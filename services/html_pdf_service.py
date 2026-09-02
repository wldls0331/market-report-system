"""Render the Market Report Preview HTML to PDF with Playwright/Chromium."""

from __future__ import annotations

import mimetypes
import subprocess
import sys
from pathlib import Path
from urllib.parse import unquote, urlparse

from config.paths import PROJECT_ROOT, session_output_dir
from config.runtime import is_hosted
from services.working_day_service import format_pdf_filename


class HtmlPdfError(RuntimeError):
    pass


def _chromium_args() -> list[str]:
    if not is_hosted():
        return []
    return [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
        "--font-render-hinting=none",
    ]


def _install_chromium() -> None:
    subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        check=True,
        timeout=300,
    )


def _launch_chromium(playwright):
    args = _chromium_args()
    try:
        return playwright.chromium.launch(headless=True, args=args)
    except Exception:
        _install_chromium()
        return playwright.chromium.launch(headless=True, args=args)


def preview_pdf_path(report_date) -> Path:
    session_output_dir().mkdir(parents=True, exist_ok=True)
    return session_output_dir() / format_pdf_filename(report_date)


def _fulfill_local_assets(route, *, html: str, static_root: Path) -> None:
    parsed = urlparse(route.request.url)
    path = unquote(parsed.path)
    if path.rstrip("/") == "/print" or path == "/":
        route.fulfill(status=200, content_type="text/html; charset=utf-8", body=html.encode("utf-8"))
        return
    if path.startswith("/static/"):
        relative = path[len("/static/") :]
        file_path = (static_root / relative).resolve()
        try:
            file_path.relative_to(static_root.resolve())
        except ValueError:
            route.fulfill(status=403, body=b"forbidden")
            return
        if not file_path.is_file():
            route.fulfill(status=404, body=b"not found")
            return
        mime = mimetypes.guess_type(str(file_path))[0] or "application/octet-stream"
        route.fulfill(status=200, content_type=mime, body=file_path.read_bytes())
        return
    route.continue_()


def render_preview_pdf(
    *,
    output_path: Path,
    print_url: str,
    html: str | None = None,
    static_root: Path | None = None,
) -> Path:
    """Open the live Preview print page and save it as a 2-page PDF."""
    try:
        from playwright.sync_api import TimeoutError as PlaywrightTimeout
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise HtmlPdfError(
            "Playwright is not installed. Run: pip install playwright && playwright install chromium"
        ) from exc

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = output_path.with_name(output_path.stem + ".tmp.pdf")
    if tmp_path.exists():
        try:
            tmp_path.unlink()
        except OSError as exc:
            raise HtmlPdfError(f"Could not replace temporary PDF ({tmp_path.name}): {exc}") from exc

    static_root = Path(static_root or (PROJECT_ROOT / "web" / "static"))

    try:
        with sync_playwright() as playwright:
            try:
                browser = _launch_chromium(playwright)
            except Exception as exc:
                raise HtmlPdfError(
                    "Chromium is not available for HTML-to-PDF. "
                    "Run: python -m playwright install chromium. "
                    f"Original error: {exc}"
                ) from exc
            try:
                page = browser.new_page(
                    viewport={"width": 794, "height": 1123},
                    device_scale_factor=2,
                )
                page.emulate_media(media="print")
                if html is not None:
                    page.route(
                        "**/*",
                        lambda route: _fulfill_local_assets(
                            route, html=html, static_root=static_root
                        ),
                    )
                response = page.goto(print_url, wait_until="domcontentloaded", timeout=60_000)
                if html is None and response is not None and not response.ok:
                    snippet = ""
                    try:
                        snippet = page.inner_text("body")[:1500]
                    except Exception:
                        snippet = page.content()[:1500]
                    raise HtmlPdfError(
                        f"Print page HTTP {response.status} at {print_url}.\n{snippet}"
                    )
                try:
                    page.wait_for_selector("body[data-preview-ready='1']", timeout=60_000)
                    page.wait_for_function(
                        """() => {
                          const canvases = [...document.querySelectorAll('canvas')];
                          return canvases.length >= 2 && canvases.every((c) => c.width > 100 && c.height > 100);
                        }""",
                        timeout=30_000,
                    )
                    page.evaluate(
                        """() => {
                          if (typeof charts === 'undefined') return;
                          Object.values(charts).forEach((chart) => {
                            if (chart && typeof chart.resize === 'function') chart.resize();
                          });
                        }"""
                    )
                    page.wait_for_timeout(500)
                except PlaywrightTimeout as exc:
                    snippet = ""
                    try:
                        snippet = page.inner_text("body")[:1500]
                    except Exception:
                        snippet = str(exc)
                    raise HtmlPdfError(
                        "Preview HTML did not finish rendering in Chromium "
                        f"(charts/CSS). URL: {print_url}\n{snippet}"
                    ) from exc
                page.pdf(
                    path=str(tmp_path),
                    format="A4",
                    landscape=False,
                    print_background=True,
                    prefer_css_page_size=True,
                    margin={"top": "0mm", "right": "0mm", "bottom": "0mm", "left": "0mm"},
                )
            finally:
                browser.close()
    except HtmlPdfError:
        raise
    except Exception as exc:
        raise HtmlPdfError(f"HTML/Chromium PDF rendering failed: {exc}") from exc

    if not tmp_path.exists() or tmp_path.stat().st_size < 100:
        raise HtmlPdfError("Chromium printed an empty PDF file.")
    try:
        tmp_path.replace(output_path)
    except OSError as exc:
        raise HtmlPdfError(
            f"Could not write PDF to {output_path} (file may be locked): {exc}"
        ) from exc
    return output_path
