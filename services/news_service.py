"""Bunker-relevant oil and shipping news for the localhost Market Report UI.

PDF generation does not import this module. Fetch failures never raise to callers.
"""

from __future__ import annotations

import json
import re
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from html import unescape
from pathlib import Path
from typing import Any
from urllib.error import URLError
from urllib.parse import quote_plus
from urllib.request import Request, urlopen

from config.paths import DATA_DIR

NEWS_WINDOW_DAYS = 7
CACHE_MAX_AGE_HOURS = 6
MAX_ITEMS = 5
UNAVAILABLE = "Market News Summary temporarily unavailable."
CACHE_PATH = DATA_DIR / "news_cache.json"
USER_AGENT = "MarketReportNews/1.0 (+localhost; bunker market briefing)"

PREMIUM_SOURCES = (
    "reuters",
    "lloyd",
    "platts",
    "s&p global",
    "sp global",
    "argus",
    "tradewinds",
    "manifold",
    "bloomberg",
    "financial times",
    "wsj",
    "wall street journal",
)
TRUSTED_SOURCES = PREMIUM_SOURCES + (
    "ft.com",
    "oilprice",
    "ship & bunker",
    "shipandbunker",
    "bunkerspot",
    "hellenic shipping",
    "splash",
    "gcaptain",
    "maritime executive",
    "lloyd's list",
    "lloyds list",
)

PRIORITY_TERMS: tuple[tuple[tuple[str, ...], int], ...] = (
    (("brent", "wti", "crude oil", "crude price"), 90),
    (("hormuz", "strait of hormuz", "red sea", "houthi", "middle east"), 85),
    (("vlcc", "product tanker", "tanker"), 75),
    (("refinery", "outage", "disruption", "gasoil", "diesel", "fuel oil"), 70),
    (("bunker", "marine fuel", "vlsfo", "hsfo", "bunkering"), 80),
    (("port congestion", "port closure", "singapore bunker", "zhoushan"), 65),
    (("freight", "war risk", "marine insurance"), 60),
    (("opec", "opec+", "inventory", "iea", "eia"), 55),
)

EXCLUDE_SOURCES = (
    "biggo",
    "future market insights",
    "inspenet",
    "pr newswire",
    "globenewswire",
    "accesswire",
)
EXCLUDE_TERMS = (
    "bitcoin",
    "crypto",
    "nft",
    "unemployment",
    "gdp growth",
    "retail sales",
    "hollywood",
    "celebrity",
    "election poll",
    "promotion",
    "team leader",
    "appoints",
    "appointed",
    "hires",
    "hiring",
    "joins as",
)

OIL_TERMS = (
    "brent",
    "wti",
    "crude",
    "opec",
    "oil price",
    "oil prices",
    "inventory",
    "refinery",
    "gasoil",
    "diesel",
    "fuel oil",
)
SHIPPING_TERMS = (
    "tanker",
    "vlcc",
    "hormuz",
    "red sea",
    "freight",
    "war risk",
    "shipping",
    "vessel",
    "port congestion",
    "suez",
)
BUNKER_TERMS = (
    "bunker",
    "marine fuel",
    "vlsfo",
    "hsfo",
    "bunkering",
    "singapore bunker",
    "fuel oil availability",
)

GOOGLE_QUERIES = (
    'Brent OR WTI OR "crude oil" (price OR inventory OR OPEC) when:7d',
    '("Strait of Hormuz" OR "Red Sea" OR Houthi) (tanker OR disruption OR freight) when:7d',
    '("bunker fuel" OR "marine fuel" OR VLSFO OR HSFO OR bunkering) when:7d',
    "(VLCC OR \"product tanker\" OR tanker) (freight OR rates OR availability) when:7d",
    "(refinery OR gasoil OR diesel OR \"fuel oil\") (disruption OR outage OR supply) when:7d",
)

DIRECT_FEEDS = (
    "https://www.manifoldtimes.com/feed/",
    "https://shipandbunker.com/news/rss",
    "https://splash247.com/feed/",
    "https://gcaptain.com/feed/",
    "https://oilprice.com/rss/main",
)

_TAG_RE = re.compile(r"<[^>]+>")
_SPACE_RE = re.compile(r"\s+")
_SENTENCE_RE = re.compile(r"(?<=[.!?])\s+")
_SOURCE_SUFFIX_RE = re.compile(r"\s+[-–—]\s+([^-–—]{2,40})$")


class NewsServiceError(RuntimeError):
    pass


def get_recent_market_news(days: int = NEWS_WINDOW_DAYS) -> list[dict[str, Any]]:
    """Fetch, filter, score and de-duplicate bunker-relevant headlines."""
    start, end = _window(days)
    raw: list[dict[str, Any]] = []
    urls = [_google_news_url(query) for query in GOOGLE_QUERIES]
    urls.extend(DIRECT_FEEDS)
    with ThreadPoolExecutor(max_workers=6) as pool:
        futures = {pool.submit(_fetch_feed, url): url for url in urls}
        for future in as_completed(futures):
            try:
                raw.extend(future.result())
            except Exception:
                continue
    filtered = [_annotate(item, start, end) for item in raw]
    filtered = [item for item in filtered if item is not None]
    return _select_items(_dedupe(filtered))


def summarize_market_news(news: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Display-ready items: headline, 1–2 sentence summary, source, date."""
    items: list[dict[str, Any]] = []
    for item in news[:MAX_ITEMS]:
        items.append(
            {
                "headline": item["headline"],
                "summary": item["summary"],
                "source": item["source"],
                "published_date": item["published_display"],
                "published_iso": item["published_iso"],
                "category": item["category"],
                "category_label": _category_label(item["category"]),
                "url": item.get("url") or "",
            }
        )
    return items


def get_weekly_market_takeaway(news: list[dict[str, Any]]) -> str:
    """2–4 sentence bunker-trader takeaway from the selected stories."""
    if not news:
        return ""
    blob = " ".join(f"{item.get('headline', '')} {item.get('summary', '')}" for item in news).lower()
    parts: list[str] = []

    if _has_any(blob, ("hormuz", "red sea", "houthi", "middle east")):
        parts.append(
            "Tanker and marine risk remains elevated around key chokepoints, "
            "which can support freight, war-risk cover and East of Suez bunker premiums."
        )
    if _has_any(blob, ("brent", "wti", "crude price", "oil settles", "oil prices")):
        if _has_any(blob, ("surge", "soar", "rally", "climb")) or (
            "higher" in blob and _has_any(blob, ("price", "brent", "wti", "crude"))
        ):
            parts.append(
                "Oil markets remain volatile, with crude prices moving higher on supply "
                "and geopolitical concerns that typically feed through into VLSFO paper."
            )
        elif _has_any(blob, ("fall", "drop", "slide", "sink", "lower", "ease", "settles lower")):
            parts.append(
                "Crude has been mixed to softer at times this week, which may ease "
                "paper-linked bunker costs if the move holds into physical stems."
            )
        else:
            parts.append(
                "Crude remains the main pricing driver for bunker traders, with Brent and WTI "
                "setting the tone for VLSFO and HSFO."
            )
    if _has_any(blob, ("tanker", "vlcc", "freight", "war risk")) and not any(
        "freight" in part.lower() or "tanker" in part.lower() for part in parts
    ):
        parts.append(
            "Tanker availability and freight remain sensitive to disruption, "
            "with possible knock-on effects for marine fuel demand and premiums."
        )
    if _has_any(blob, ("bunker", "marine fuel", "vlsfo", "hsfo", "refinery", "fuel oil")):
        parts.append(
            "Fuel-oil and bunker availability should be watched at key hubs, "
            "as refinery or supply changes can move physical premiums quickly."
        )
    if not parts:
        lead = news[0].get("summary") or news[0].get("headline") or ""
        parts.append(str(lead).strip())
        parts.append(
            "Bunker traders should watch whether these developments persist into "
            "the coming week's physical stems and East/West differentials."
        )
    return " ".join(parts[:4]).strip()


def get_news_payload(*, days: int = NEWS_WINDOW_DAYS, force: bool = False) -> dict[str, Any]:
    """Cached payload for the web UI. Never raises."""
    start, end = _window(days)
    empty = _empty_payload(start, end)
    if not force:
        cached = _read_cache()
        if cached and not _is_stale(cached):
            return cached
    try:
        from services.perf import timed

        with timed("news fetch"):
            raw = get_recent_market_news(days=days)
        items = summarize_market_news(raw)
        takeaway = get_weekly_market_takeaway(items)
        if not items:
            payload = {**empty, "error": UNAVAILABLE, "fetched_at": _now().isoformat()}
        else:
            payload = {
                "window": _window_meta(start, end),
                "takeaway": takeaway,
                "items": items,
                "fetched_at": _now().isoformat(),
                "stale": False,
                "error": None,
            }
        _write_cache(payload)
        return payload
    except Exception:
        cached = _read_cache()
        if cached and cached.get("items"):
            cached = dict(cached)
            cached["stale"] = True
            return cached
        return {**empty, "error": UNAVAILABLE, "fetched_at": _now().isoformat()}


def refresh_market_news_if_stale(*, days: int = NEWS_WINDOW_DAYS) -> dict[str, Any]:
    """Used before Create Report. Failures are swallowed by get_news_payload."""
    cached = _read_cache()
    if cached and not _is_stale(cached) and cached.get("items"):
        return cached
    return get_news_payload(days=days, force=True)


def _now() -> datetime:
    return datetime.now(timezone.utc)


def _window(days: int) -> tuple[date, date]:
    end = datetime.now().date()
    start = end - timedelta(days=max(days, 1) - 1)
    return start, end


def _window_meta(start: date, end: date) -> dict[str, str]:
    return {
        "from": start.isoformat(),
        "to": end.isoformat(),
        "label": "Last 7 days",
    }


def _empty_payload(start: date, end: date) -> dict[str, Any]:
    return {
        "window": _window_meta(start, end),
        "takeaway": "",
        "items": [],
        "fetched_at": None,
        "stale": False,
        "error": None,
    }


def _is_stale(payload: dict[str, Any]) -> bool:
    if payload.get("error") and not payload.get("items"):
        return True
    raw = payload.get("fetched_at")
    if not raw:
        return True
    try:
        fetched = datetime.fromisoformat(str(raw))
        if fetched.tzinfo is None:
            fetched = fetched.replace(tzinfo=timezone.utc)
    except ValueError:
        return True
    return _now() - fetched > timedelta(hours=CACHE_MAX_AGE_HOURS)


def _read_cache() -> dict[str, Any] | None:
    try:
        if not CACHE_PATH.is_file():
            return None
        data = json.loads(CACHE_PATH.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return None
        return data
    except Exception:
        return None


def _write_cache(payload: dict[str, Any]) -> None:
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        CACHE_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def _google_news_url(query: str) -> str:
    return (
        "https://news.google.com/rss/search?q="
        + quote_plus(query)
        + "&hl=en-US&gl=US&ceid=US:en"
    )


def _fetch_feed(url: str) -> list[dict[str, Any]]:
    request = Request(url, headers={"User-Agent": USER_AGENT, "Accept": "application/rss+xml, application/xml, text/xml"})
    try:
        with urlopen(request, timeout=10) as response:
            body = response.read()
    except (URLError, TimeoutError, OSError) as exc:
        raise NewsServiceError(f"Could not read feed: {url}") from exc
    return _parse_feed(body, url)


def _parse_feed(body: bytes, url: str) -> list[dict[str, Any]]:
    try:
        root = ET.fromstring(body)
    except ET.ParseError:
        return []
    items: list[dict[str, Any]] = []
    for node in root.iter():
        if _local(node.tag) not in {"item", "entry"}:
            continue
        title = _child_text(node, "title")
        if not title:
            continue
        link = _child_text(node, "link") or _child_attr(node, "link", "href")
        source = _child_text(node, "source") or _source_from_title(title) or _source_from_url(url)
        headline, source = _split_headline(title, source)
        description = _child_text(node, "description") or _child_text(node, "summary")
        published = _parse_date(
            _child_text(node, "pubDate")
            or _child_text(node, "published")
            or _child_text(node, "updated")
            or _child_text(node, "date")
        )
        items.append(
            {
                "headline": headline,
                "source": source,
                "url": link,
                "description": description,
                "published": published,
            }
        )
    return items


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1].lower()


def _child_text(node: ET.Element, name: str) -> str:
    want = name.lower()
    for child in list(node):
        if _local(child.tag) == want and (child.text or "").strip():
            return unescape((child.text or "").strip())
    return ""


def _child_attr(node: ET.Element, name: str, attr: str) -> str:
    want = name.lower()
    for child in list(node):
        if _local(child.tag) != want:
            continue
        if child.attrib.get(attr):
            return child.attrib[attr].strip()
        for key, value in child.attrib.items():
            if key.endswith(attr) and value:
                return value.strip()
    return ""


def _source_from_title(title: str) -> str:
    match = _SOURCE_SUFFIX_RE.search(title.strip())
    return match.group(1).strip() if match else ""


def _source_from_url(url: str) -> str:
    host = re.sub(r"^https?://(www\.)?", "", url).split("/")[0]
    mapping = {
        "manifoldtimes.com": "Manifold Times",
        "shipandbunker.com": "Ship & Bunker",
        "splash247.com": "Splash",
        "gcaptain.com": "gCaptain",
        "oilprice.com": "OilPrice",
        "news.google.com": "",
    }
    return mapping.get(host, host)


def _split_headline(title: str, source: str) -> tuple[str, str]:
    headline = _SPACE_RE.sub(" ", title).strip()
    match = _SOURCE_SUFFIX_RE.search(headline)
    if match:
        suffix = match.group(1).strip()
        headline = headline[: match.start()].strip()
        if not source:
            source = suffix
    return headline, source or "Unknown"


def _parse_date(value: str) -> datetime | None:
    text = (value or "").strip()
    if not text:
        return None
    try:
        parsed = parsedate_to_datetime(text)
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)
    except (TypeError, ValueError, OverflowError):
        pass
    cleaned = text.replace("Z", "+00:00")
    for candidate in (cleaned, cleaned[:19], text[:10]):
        for fmt in ("%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
            try:
                parsed = datetime.strptime(candidate, fmt)
                if parsed.tzinfo is None:
                    parsed = parsed.replace(tzinfo=timezone.utc)
                return parsed.astimezone(timezone.utc)
            except ValueError:
                continue
    return None


def _annotate(item: dict[str, Any], start: date, end: date) -> dict[str, Any] | None:
    published = item.get("published")
    if published is None:
        return None
    pub_day = published.date()
    if pub_day < start or pub_day > end:
        return None
    headline = item["headline"]
    description = _clean_html(item.get("description") or "")
    blob = f"{headline} {description} {item.get('source') or ''}".lower()
    source_l = (item.get("source") or "").lower()
    if any(name in source_l for name in EXCLUDE_SOURCES):
        return None
    if any(term in blob for term in EXCLUDE_TERMS):
        market_ok = _has_any(blob, OIL_TERMS + SHIPPING_TERMS + BUNKER_TERMS)
        personnel = _has_any(blob, ("promotion", "team leader", "appoints", "appointed", "hires", "hiring", "joins as"))
        if personnel or not market_ok:
            return None
    if not _has_any(blob, OIL_TERMS + SHIPPING_TERMS + BUNKER_TERMS):
        return None
    category = _category(blob)
    score = _score(blob, item.get("source") or "")
    if score < 40:
        return None
    summary = _summary_text(headline, description, category)
    return {
        "headline": headline,
        "summary": summary,
        "source": _pretty_source(item.get("source") or "Unknown"),
        "url": item.get("url") or "",
        "published": published,
        "published_iso": published.date().isoformat(),
        "published_display": _display_date(published.date()),
        "category": category,
        "score": score,
    }


def _clean_html(value: str) -> str:
    text = unescape(_TAG_RE.sub(" ", value or ""))
    text = text.replace("View Full Coverage on Google News", "")
    return _SPACE_RE.sub(" ", text).strip()


def _summary_text(headline: str, description: str, category: str) -> str:
    desc = description or ""
    if headline:
        desc = re.sub(re.escape(headline), " ", desc, flags=re.IGNORECASE)
    desc = _SPACE_RE.sub(" ", desc).strip(" |-–—")
    sentences = [part.strip() for part in _SENTENCE_RE.split(desc) if part.strip()]
    usable = [
        sentence
        for sentence in sentences
        if len(sentence) > 50
        and "google news" not in sentence.lower()
        and not sentence.lower().startswith("http")
        and sentence.lower() not in headline.lower()
    ]
    if usable:
        text = " ".join(usable[:2])
        if len(text) > 360:
            text = text[:357].rsplit(" ", 1)[0] + "."
        return text
    return _fallback_summary(headline, category)


def _fallback_summary(headline: str, category: str) -> str:
    lead = headline.rstrip(".")
    if category == "shipping":
        return (
            f"{lead}. This may affect tanker movements, freight and war-risk costs "
            "that bunker traders watch across key shipping lanes."
        )
    if category == "bunker":
        return (
            f"{lead}. Availability and premiums at major bunker hubs could react "
            "if the disruption persists."
        )
    return (
        f"{lead}. Crude and product-market moves of this type often transmit into "
        "VLSFO paper and physical bunker stems."
    )


def _category(blob: str) -> str:
    if _has_any(blob, ("bunker", "marine fuel", "vlsfo", "hsfo", "bunkering")):
        return "bunker"
    if _has_any(blob, ("settles", "oil price", "brent", "wti")) and _has_any(blob, ("fed", "policy", "inventory", "opec")):
        return "oil"
    if _has_any(blob, ("hormuz", "red sea", "houthi", "tanker", "vlcc", "supertanker", "freight", "war risk")):
        return "shipping"
    bunker_hits = sum(1 for term in BUNKER_TERMS if term in blob)
    shipping_hits = sum(1 for term in SHIPPING_TERMS if term in blob)
    oil_hits = sum(1 for term in OIL_TERMS if term in blob)
    ranked = [("oil", oil_hits), ("shipping", shipping_hits), ("bunker", bunker_hits)]
    ranked.sort(key=lambda item: item[1], reverse=True)
    return ranked[0][0] if ranked[0][1] else "oil"


def _category_label(category: str) -> str:
    return {
        "oil": "Oil / Energy",
        "shipping": "Shipping",
        "bunker": "Bunker / Supply",
    }.get(category, "Market")


def _score(blob: str, source: str) -> int:
    score = 0
    for terms, weight in PRIORITY_TERMS:
        if _has_any(blob, terms):
            score += weight
    source_l = source.lower()
    if any(name in source_l for name in PREMIUM_SOURCES):
        score += 45
    elif any(name in source_l for name in TRUSTED_SOURCES):
        score += 15
    else:
        score -= 20
    if _has_any(blob, ("bunker", "vlsfo", "marine fuel", "hormuz", "red sea")):
        score += 15
    return score


def _pretty_source(source: str) -> str:
    raw = source.strip()
    if "|" in raw:
        raw = raw.split("|")[-1].strip()
    key = raw.lower()
    aliases = {
        "s&p global": "S&P Global / Platts",
        "sp global": "S&P Global / Platts",
        "platts": "S&P Global / Platts",
        "lloyds list": "Lloyd's List",
        "lloyd's list": "Lloyd's List",
        "shipandbunker": "Ship & Bunker",
        "ship & bunker": "Ship & Bunker",
        "oilprice.com": "OilPrice",
        "oilprice": "OilPrice",
        "crude oil prices today": "OilPrice",
        "reuters.com": "Reuters",
        "bloomberg.com": "Bloomberg",
    }
    for needle, label in aliases.items():
        if needle in key:
            return label
    return raw


def _display_date(value: date) -> str:
    return f"{value.day} {value.strftime('%b %Y')}"


def _has_any(text: str, terms: tuple[str, ...]) -> bool:
    return any(term in text for term in terms)


def _normalize_title(title: str) -> str:
    return re.sub(r"[^a-z0-9\s]", "", title.lower())


def _dedupe(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    items = sorted(items, key=lambda item: item["score"], reverse=True)
    kept: list[dict[str, Any]] = []
    for item in items:
        tokens = set(_normalize_title(item["headline"]).split())
        duplicate = False
        for existing in kept:
            other = set(_normalize_title(existing["headline"]).split())
            overlap = len(tokens & other) / max(len(tokens | other), 1) if tokens and other else 0
            if overlap >= 0.42 or _same_event(item["headline"], existing["headline"]):
                if item["score"] > existing["score"]:
                    kept[kept.index(existing)] = item
                duplicate = True
                break
        if not duplicate:
            kept.append(item)
    return kept


def _same_event(left: str, right: str) -> bool:
    a = _normalize_title(left)
    b = _normalize_title(right)
    for marker in ("hormuz", "red sea", "houthi"):
        if marker in a and marker in b:
            actions = ("hit", "attack", "mine", "strike", "tanker", "supertanker")
            if any(word in a for word in actions) and any(word in b for word in actions):
                return True
    return False


def _source_rank(item: dict[str, Any]) -> int:
    source = (item.get("source") or "").lower()
    if any(name in source for name in PREMIUM_SOURCES):
        return 0
    if any(name in source for name in TRUSTED_SOURCES):
        return 1
    return 2


def _select_items(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    ranked = sorted(items, key=lambda row: (_source_rank(row), -row["score"]))
    by_cat: dict[str, list[dict[str, Any]]] = {"oil": [], "shipping": [], "bunker": []}
    for item in ranked:
        by_cat.setdefault(item["category"], []).append(item)
    chosen: list[dict[str, Any]] = []
    seen: set[str] = set()
    for category in ("oil", "shipping", "bunker"):
        for item in by_cat.get(category, []):
            key = _normalize_title(item["headline"])
            if key in seen:
                continue
            chosen.append(item)
            seen.add(key)
            break
    for item in ranked:
        if len(chosen) >= MAX_ITEMS:
            break
        key = _normalize_title(item["headline"])
        if key in seen:
            continue
        chosen.append(item)
        seen.add(key)
    order = {"oil": 0, "shipping": 1, "bunker": 2}
    chosen.sort(key=lambda item: (order.get(item["category"], 9), _source_rank(item), -item["score"]))
    return chosen[:MAX_ITEMS]
