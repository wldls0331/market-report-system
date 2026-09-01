const PALETTE = ["#1b365d", "#6b7280", "#c4a35a", "#334155", "#94a3b8"];
const charts = {};

function $(id) {
  return document.getElementById(id);
}

function showMessage(text, kind) {
  const el = $("message");
  if (!text) {
    el.hidden = true;
    el.textContent = "";
    return;
  }
  el.hidden = false;
  el.className = "banner" + (kind ? " " + kind : "");
  el.textContent = text;
}

function setPill(id, label, value, ok) {
  const el = $(id);
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

function renderLineChart(canvasId, payload) {
  const canvas = $(canvasId);
  if (!canvas || typeof Chart === "undefined") return;
  const existing = charts[canvasId];
  if (existing) existing.destroy();
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
      })),
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: "bottom" } },
      scales: {
        y: { ticks: { color: "#6b7280" }, grid: { color: "#eef2f7" } },
        x: { ticks: { color: "#6b7280" }, grid: { display: false } },
      },
    },
  });
}

function renderBarChart(canvasId, payload) {
  const canvas = $(canvasId);
  if (!canvas || typeof Chart === "undefined") return;
  const existing = charts[canvasId];
  if (existing) existing.destroy();
  const series = payload?.series || [];
  charts[canvasId] = new Chart(canvas, {
    type: "bar",
    data: {
      labels: payload?.labels || [],
      datasets: series.map((item, index) => ({
        label: item.name,
        data: item.data,
        backgroundColor: PALETTE[index % PALETTE.length],
      })),
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: "bottom" } },
      scales: {
        y: { ticks: { color: "#6b7280" }, grid: { color: "#eef2f7" } },
        x: { ticks: { color: "#6b7280" }, grid: { display: false } },
      },
    },
  });
}

function applyStatus(status) {
  if (!status) return;
  setPill("sheets-status", "Google Sheets", status.sheets, status.sheets_ok);
  setPill("gmail-status", "Gmail", status.gmail, status.gmail_ok);
  const sheetBtn = $("btn-sheet");
  if (status.sheet_url) {
    sheetBtn.href = status.sheet_url;
    sheetBtn.classList.remove("is-disabled");
  } else {
    sheetBtn.removeAttribute("href");
    sheetBtn.classList.add("is-disabled");
  }
  $("btn-auth-sheets").hidden = status.sheets_ok || status.sheets !== "Authorization Required";
  $("btn-auth-gmail").hidden = status.gmail_ok || status.gmail !== "Authorization Required";
}

function applyPreview(data) {
  applyStatus(data.status);
  const meta = data.meta || {};
  $("meta-date").textContent = meta.report_date || "—";
  $("meta-month").textContent = meta.pricing_month || "—";
  $("meta-this").textContent = meta.this_week_friday || "—";
  $("meta-prev").textContent = meta.previous_week_friday || "—";
  $("meta-two").textContent = meta.two_weeks_ago_friday || "—";
  $("report-title").textContent = meta.report_title || "WEEKLY BUNKERING REPORT";
  const comments = data.comments || {};
  $("comment-korea").textContent = comments.korea || "TBN";
  $("comment-korea-world").textContent = comments.korea_worldwide || comments.korea || "TBN";
  $("comment-singapore").textContent = comments.singapore || "TBN";
  $("comment-china").textContent = comments.china || "TBN";
  $("comment-japan").textContent = comments.japan || "TBN";
  const strategy = (comments.strategy || []).filter(Boolean);
  $("strategy-block").textContent = strategy.length ? strategy.join("\n") : "";
  renderLineChart("chart-korea", data.charts?.korea_premium);
  renderLineChart("chart-world", data.charts?.worldwide_vlsfo);
  renderBarChart("chart-spread", data.charts?.spread);
  if (data.email) {
    $("email-to").value = data.email.to || "";
    $("email-cc").value = data.email.cc || "";
    $("email-subject").value = data.email.subject || "";
    $("email-body").value = data.email.body || "";
    $("email-attachment").textContent = data.email.attachment
      ? `Attachment: ${data.email.attachment}`
      : "Attachment: none";
  }
  const files = data.files || {};
  $("btn-pdf").classList.toggle("is-disabled", !files.has_pdf);
  $("btn-excel").classList.toggle("is-disabled", !files.has_excel);
  $("btn-create").disabled = meta.date_ok === false;
  if (meta.date_warning) showMessage(meta.date_warning, "error");
  else if (data.error) showMessage(data.error, "error");
  else if (data.warning) showMessage(data.warning);
  else showMessage("");
}

async function loadPreview() {
  showMessage("Loading report preview…");
  const data = await api("/api/preview");
  applyPreview(data);
  if (!data.error && !data.meta?.date_warning && !data.warning) showMessage("");
}

async function refreshData() {
  showMessage("Refreshing INPUT and dated report sheet…");
  const data = await api("/api/refresh", { method: "POST" });
  applyPreview(data);
  if (!data.error) {
    showMessage(
      data.synced
        ? (data.is_update ? "Dated report sheet updated. Preview refreshed." : "Dated report sheet created. Preview refreshed.")
        : "Preview refreshed.",
      "ok"
    );
  }
}

async function createReport() {
  showMessage("Updating the dated sheet and generating PDF…");
  const data = await api("/api/create-report", { method: "POST" });
  applyPreview(data);
  if (!data.error) showMessage("Market Report generated. PDF is ready to download.", "ok");
}

async function sendMail() {
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
}

function bind(id, handler) {
  $(id).addEventListener("click", async (event) => {
    event.preventDefault();
    try {
      await handler();
    } catch (err) {
      showMessage(err.message || String(err), "error");
    }
  });
}

bind("btn-refresh", refreshData);
bind("btn-create", createReport);
bind("btn-send", sendMail);
bind("btn-auth-sheets", async () => {
  showMessage("Authorize Google Sheets in the browser window…");
  const data = await api("/api/authorize/sheets", { method: "POST" });
  applyStatus(data);
  await loadPreview();
});
bind("btn-auth-gmail", async () => {
  showMessage("Authorize Gmail in the browser window…");
  const data = await api("/api/authorize/gmail", { method: "POST" });
  applyStatus(data);
});

loadPreview().catch((err) => showMessage(err.message || String(err), "error"));
