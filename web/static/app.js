const PALETTE = ["#1b365d", "#6b7280", "#c4a35a", "#334155", "#94a3b8"];
const charts = {};

if (typeof Chart !== "undefined") {
  Chart.defaults.animation = false;
  Chart.defaults.devicePixelRatio = Math.max(window.devicePixelRatio || 1, 2);
}

function $(id) {
  return document.getElementById(id);
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function setText(id, value) {
  const el = $(id);
  if (!el) return;
  el.classList.remove("is-skeleton");
  el.textContent = value;
}

function showMessage(text, kind) {
  const el = $("message");
  if (!el) return;
  if (!text) {
    el.hidden = true;
    el.textContent = "";
    return;
  }
  el.hidden = false;
  el.className = "banner" + (kind ? " " + kind : "");
  el.textContent = text;
}

function setBusy(id, busy, label) {
  const el = $(id);
  if (!el) return;
  if (busy) {
    if (!el.dataset.label) el.dataset.label = el.textContent;
    el.disabled = true;
    if (label) el.textContent = label;
  } else {
    el.disabled = false;
    el.textContent = el.dataset.label || el.textContent;
    delete el.dataset.label;
  }
}

function setOverlay(id, hidden) {
  const el = $(id);
  if (el) el.hidden = !!hidden;
}

function setPill(id, label, value, ok) {
  const el = $(id);
  if (!el) return;
  el.textContent = `${label}: ${value || "—"}`;
  el.classList.toggle("ok", !!ok);
  el.classList.toggle("bad", ok === false);
}

async function api(url, options) {
  const response = await fetch(url, options);
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data.error || `Request failed (${response.status})`);
  }
  return data;
}

function chartOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: false,
    devicePixelRatio: Math.max(window.devicePixelRatio || 1, 2),
    plugins: {
      legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } },
    },
    scales: {
      y: { ticks: { color: "#6b7280", font: { size: 11 } }, grid: { color: "#eef2f7" } },
      x: { ticks: { color: "#6b7280", font: { size: 11 } }, grid: { display: false } },
    },
  };
}

function renderLineChart(canvasId, payload) {
  const canvas = $(canvasId);
  if (!canvas || typeof Chart === "undefined") return;
  const existing = charts[canvasId];
  if (existing) existing.destroy();
  delete charts[canvasId];
  const series = payload?.series || [];
  charts[canvasId] = new Chart(canvas, {
    type: "line",
    data: {
      labels: payload?.labels || [],
      datasets: series.map((item, index) => ({
        label: item.name,
        data: item.data,
        borderColor: PALETTE[index % PALETTE.length],
        backgroundColor: "transparent",
        spanGaps: true,
        tension: 0.15,
        pointRadius: 3,
        borderWidth: 2,
      })),
    },
    options: chartOptions(),
  });
}

function applyStatus(status) {
  if (!status) return;
  if ($("sheets-status")) {
    setPill("sheets-status", "Google Sheets", status.sheets, status.sheets_ok);
  }
  if ($("gmail-status")) {
    setPill("gmail-status", "Gmail", status.gmail, status.gmail_ok);
  }
  const sheetBtn = $("btn-sheet");
  if (!sheetBtn) return;
  if (status.sheet_url) {
    sheetBtn.href = status.sheet_url;
    sheetBtn.classList.remove("is-disabled");
  } else {
    sheetBtn.removeAttribute("href");
    sheetBtn.classList.add("is-disabled");
  }
  if ($("btn-auth-sheets")) {
    $("btn-auth-sheets").hidden = status.sheets_ok || status.sheets !== "Authorization Required";
  }
  if ($("btn-auth-gmail")) {
    $("btn-auth-gmail").hidden = status.gmail_ok || status.gmail !== "Authorization Required";
  }
}

function applyCharts(chartsPayload) {
  if (chartsPayload?.korea_premium) {
    renderLineChart("chart-korea", chartsPayload.korea_premium);
    setOverlay("korea-chart-loading", true);
  }
  if (chartsPayload?.worldwide_vlsfo) {
    renderLineChart("chart-world", chartsPayload.worldwide_vlsfo);
    setOverlay("world-chart-loading", true);
  }
}

function applyPreview(data) {
  if (data.status) applyStatus(data.status);
  const meta = data.meta || {};
  setText("meta-date", meta.report_date || "—");
  setText("meta-month", meta.pricing_month || "—");
  setText("meta-this", meta.this_week_friday || "—");
  setText("meta-prev", meta.previous_week_friday || "—");
  setText("meta-two", meta.two_weeks_ago_friday || "—");
  setText("report-title", "WEEKLY BUNKERING REPORT");
  setText("report-date-display", meta.report_date_display || meta.report_date || "—");
  const updated = meta.updated_display || meta.data_reference_date || "—";
  document.querySelectorAll(".chart-updated").forEach((el) => {
    el.textContent = updated;
  });
  const comments = data.comments || {};
  setText("comment-korea-world", comments.korea_worldwide || "TBN");
  setText("comment-singapore", comments.singapore || "TBN");
  setText("comment-china", comments.china || "TBN");
  setText("comment-japan", comments.japan || "TBN");
  const premiumRoot = $("premium-list");
  if (premiumRoot) {
    const rows = data.supplier_premiums || [];
    premiumRoot.innerHTML = rows.map((row) => `
      <div class="premium-item">
        <strong>${escapeHtml(row.label || "")}</strong>
        <span>VLSFO: ${escapeHtml(row.vlsfo || "TBN")}</span>
        <span>HSFO: ${escapeHtml(row.hsfo || "TBN")}</span>
      </div>
    `).join("");
  }
  applyCharts(data.charts);
  if (data.email && $("email-to")) {
    $("email-to").value = data.email.to || "";
    $("email-cc").value = data.email.cc || "";
    $("email-subject").value = data.email.subject || "";
    $("email-body").value = data.email.body || "";
    setText(
      "email-attachment",
      data.email.attachment ? `Attachment: ${data.email.attachment}` : "Attachment: none"
    );
  }
  const createBtn = $("btn-create");
  if (createBtn && !createBtn.dataset.label) createBtn.disabled = meta.date_ok === false;
  if (document.body.classList.contains("print-document")) return;
  if (meta.date_warning) showMessage(meta.date_warning, "error");
  else if (data.error) showMessage(data.error, "error");
  else if (data.warning) showMessage(data.warning);
}

async function loadStatus() {
  try {
    applyStatus(await api("/api/status"));
  } catch (err) {
    setPill("sheets-status", "Google Sheets", "Not Connected", false);
    setPill("gmail-status", "Gmail", "Not Connected", false);
  }
}

async function loadReportData() {
  const data = await api("/api/report-data");
  applyPreview(data);
  return data;
}

async function loadPreview() {
  const data = await api("/api/preview");
  applyPreview(data);
  return data;
}

async function refreshData() {
  setBusy("btn-refresh", true, "Refreshing...");
  try {
    const data = await api("/api/refresh", { method: "POST" });
    applyPreview(data);
    if (data.error) {
      showMessage(`Refresh failed: ${data.error}`, "error");
      return;
    }
    showMessage("Updated just now", "ok");
  } catch (err) {
    showMessage(`Refresh failed: ${err.message || err}`, "error");
    throw err;
  } finally {
    setBusy("btn-refresh", false);
  }
}

async function createReport() {
  setBusy("btn-create", true, "Creating report...");
  try {
    showMessage("Updating the dated sheet and Excel file…");
    const data = await api("/api/create-report", { method: "POST" });
    applyPreview(data);
    loadNews().catch(() => {});
    if (!data.error) {
      showMessage("Market Report generated. Excel is ready. Download PDF from the current Preview.", "ok");
    }
  } finally {
    setBusy("btn-create", false);
  }
}

function applyNews(data) {
  const windowEl = $("news-window");
  if (windowEl) windowEl.textContent = data.window?.label || "Last 7 days";
  const takeawayBox = $("news-takeaway");
  const takeawayText = $("news-takeaway-text");
  const list = $("news-list");
  const status = $("news-status");
  if (!list) return;

  const items = data.items || [];
  if (data.error && !items.length) {
    if (takeawayBox) takeawayBox.hidden = true;
    list.innerHTML = "";
    if (status) status.textContent = data.error;
    return;
  }

  if (takeawayBox && takeawayText) {
    if (data.takeaway) {
      takeawayBox.hidden = false;
      takeawayText.textContent = data.takeaway;
    } else {
      takeawayBox.hidden = true;
    }
  }

  list.innerHTML = items.map((item) => `
    <article class="news-item">
      <span class="news-kicker">${escapeHtml(item.category_label || "")}</span>
      <h4>${escapeHtml(item.headline || "")}</h4>
      <p>${escapeHtml(item.summary || "")}</p>
      <p class="news-meta">${escapeHtml(item.source || "Unknown")} | ${escapeHtml(item.published_date || "")}</p>
    </article>
  `).join("");

  if (status) {
    status.textContent = items.length
      ? (data.stale ? "Showing the last saved briefing. Refresh News for the latest 7 days." : "")
      : "No bunker-relevant headlines in the last 7 days.";
  }
}

async function loadNews() {
  const status = $("news-status");
  if (status) {
    status.innerHTML = '<span class="inline-loading"><span class="spinner spinner-sm"></span> Loading recent market news...</span>';
  }
  const data = await api("/api/news");
  applyNews(data);
}

async function refreshNews() {
  setBusy("btn-news", true, "Refreshing...");
  const status = $("news-status");
  if (status) status.textContent = "Refreshing last 7 days of market news…";
  try {
    const data = await api("/api/news/refresh", { method: "POST" });
    applyNews(data);
    if (data.error && !(data.items || []).length) {
      showMessage(data.error, "error");
      return;
    }
    showMessage("Market News Summary updated.", "ok");
  } finally {
    setBusy("btn-news", false);
  }
}

async function downloadKind(kind) {
  const btnId = kind === "pdf" ? "btn-pdf" : "btn-excel";
  setBusy(btnId, true, kind === "pdf" ? "Generating PDF..." : "Preparing Excel...");
  try {
    showMessage(kind === "pdf" ? "Rendering Preview to PDF…" : "Preparing Excel…");
    const response = await fetch(`/api/download/${kind}`);
    const disposition = response.headers.get("Content-Disposition") || "";
    const match = /filename\*?=(?:UTF-8''|")?([^\";]+)/i.exec(disposition);
    const fallback = kind === "pdf" ? "Weekly Report_Bunkering.pdf" : "Market Report.xlsx";
    const filename = match ? decodeURIComponent(match[1].replace(/"/g, "")) : fallback;
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      throw new Error(data.error || `Download failed (${response.status})`);
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    showMessage(`${filename} downloaded.`, "ok");
  } finally {
    setBusy(btnId, false);
  }
}

async function sendMail() {
  setBusy("btn-send", true, "Sending...");
  try {
    showMessage("Sending email…");
    const data = await api("/api/send-email", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        to: $("email-to").value,
        cc: $("email-cc").value,
        subject: $("email-subject").value,
        body: $("email-body").value,
      }),
    });
    showMessage(data.message || "Email sent successfully.", "ok");
  } finally {
    setBusy("btn-send", false);
  }
}

function bind(id, handler) {
  const el = $(id);
  if (!el) return;
  el.addEventListener("click", async (event) => {
    event.preventDefault();
    if (el.disabled) return;
    try {
      await handler();
    } catch (err) {
      showMessage(err.message || String(err), "error");
    }
  });
}

bind("btn-refresh", refreshData);
bind("btn-news", refreshNews);
bind("btn-create", createReport);
bind("btn-pdf", () => downloadKind("pdf"));
bind("btn-excel", () => downloadKind("excel"));
bind("btn-send", sendMail);
bind("btn-auth-sheets", async () => {
  showMessage("Authorize Google Sheets in the browser window…");
  const data = await api("/api/authorize/sheets", { method: "POST" });
  applyStatus(data);
  await loadReportData();
});
bind("btn-auth-gmail", async () => {
  showMessage("Authorize Gmail in the browser window…");
  const data = await api("/api/authorize/gmail", { method: "POST" });
  applyStatus(data);
});

if (!document.body.classList.contains("print-document")) {
  const statusTask = loadStatus();
  const reportTask = loadReportData()
    .catch((err) => {
      showMessage(err.message || String(err), "error");
    })
    .finally(() => {
      setOverlay("page-progress", true);
    });
  Promise.allSettled([statusTask, reportTask]).then(() => {
    loadNews().catch(() => {
      applyNews({ error: "Market News Summary temporarily unavailable.", items: [] });
    });
  });
}
