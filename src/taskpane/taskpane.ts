/**
 * PivotOps — Task Pane Entry Point
 * Wires up all UI interactions with underlying services.
 */

import "./taskpane.css";
import { signIn, signOut, getAccount, savePat, saveOAuthConfig, hasAuthConfig, getAuthMode } from "../services/auth";
import {
  syncToExcel,
  fetchAnalyticsWorkItems,
  getIterations,
  saveAdoConfig,
  getAdoConfig,
  hasAdoConfig,
  validateUpdates,
  bulkUpdateWorkItems,
  type AnalyticsWorkItem,
  type WorkItemUpdate,
  type SyncOptions,
} from "../services/ado-client";
import { initPyodide, runChart, runAnalysis, loadPredictivePackages, onStatusChange } from "../services/python-engine";
import {
  CHART_STATE_DISTRIBUTION,
  CHART_VELOCITY,
  CHART_SCOPE_DRIFT,
  CHART_CYCLE_TIME,
  CHART_CUMULATIVE_FLOW,
  CHART_BURNDOWN,
  CHART_MEMBER_VELOCITY,
  CHART_ESTIMATION_ACCURACY,
  ANALYSIS_SPRINT_PREDICTION,
  ANALYSIS_COMPLETION_FORECAST,
  ANALYSIS_SUMMARY,
} from "../python/charts";

// ── Globals ────────────────────────────────────────────────────────────────

let cachedWorkItems: AnalyticsWorkItem[] = [];
let currentChart: string | null = null;

// ── DOM Helpers ────────────────────────────────────────────────────────────

function $(id: string): HTMLElement {
  return document.getElementById(id)!;
}

function show(el: HTMLElement | string, display = "block"): void {
  const elem = typeof el === "string" ? $(el) : el;
  elem.style.display = display;
}

function hide(el: HTMLElement | string): void {
  const elem = typeof el === "string" ? $(el) : el;
  elem.style.display = "none";
}

function setResult(id: string, message: string, type: "success" | "error" | "info"): void {
  const el = $(id);
  el.textContent = message;
  el.className = `result-box ${type}`;
  show(el);
}

function setStatus(text: string, type: "idle" | "loading" | "running" | "error" = "idle"): void {
  $("status-text").textContent = text;
  $("status-bar").className = `status-bar status-${type}`;
}

// ── Initialization ─────────────────────────────────────────────────────────

Office.onReady(async () => {
  // Wire up navigation
  document.querySelectorAll<HTMLButtonElement>(".tab").forEach((tab) => {
    tab.addEventListener("click", () => switchTab(tab.dataset.tab!));
  });

  // Wire up buttons
  $("btn-connect").addEventListener("click", handleConnect);
  $("btn-settings").addEventListener("click", () => show("settings-modal", "flex"));
  $("btn-close-settings").addEventListener("click", () => hide("settings-modal"));
  $("btn-save-settings").addEventListener("click", handleSaveSettings);
  $("btn-sync").addEventListener("click", handleSync);

  // Layout toggle buttons
  document.querySelectorAll<HTMLButtonElement>(".toggle-btn[data-layout]").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll<HTMLButtonElement>(".toggle-btn[data-layout]").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
    });
  });

  $("btn-predict-velocity").addEventListener("click", handlePredictVelocity);
  $("btn-forecast").addEventListener("click", handleForecast);
  $("btn-dry-run").addEventListener("click", handleDryRun);
  $("btn-push-updates").addEventListener("click", handlePushUpdates);
  $("btn-export-chart").addEventListener("click", handleExportChart);

  // Wire up chart buttons
  document.querySelectorAll<HTMLButtonElement>(".chart-option").forEach((btn) => {
    btn.addEventListener("click", () => handleChartSelect(btn.dataset.chart!));
  });

  // Python engine status
  onStatusChange((status, detail) => {
    const emoji = status === "ready" ? "✅" : status === "loading" ? "⏳" : status === "running" ? "⚡" : status === "error" ? "❌" : "💤";
    $("python-status").textContent = `🐍 ${emoji} ${detail || status}`;
  });

  // Load settings into form
  await loadSettingsForm();

  // Check if already configured
  await updateUI();

  // Navigate to the tab requested by the ribbon button (via ?tab= query param)
  const params = new URLSearchParams(window.location.search);
  const requestedTab = params.get("tab");
  if (requestedTab) {
    switchTab(requestedTab);
  }

  // Start loading Python engine in the background
  initPyodide().catch(() => {/* Non-fatal — will retry on use */});
});

// ── UI State ───────────────────────────────────────────────────────────────

function switchTab(tabName: string): void {
  document.querySelectorAll<HTMLButtonElement>(".tab").forEach((t) => t.classList.remove("active"));
  document.querySelectorAll<HTMLElement>(".panel").forEach((p) => p.classList.remove("active"));

  document.querySelector<HTMLButtonElement>(`.tab[data-tab="${tabName}"]`)?.classList.add("active");
  $(`panel-${tabName}`).classList.add("active");
}

async function updateUI(): Promise<void> {
  const account = await getAccount();
  const hasConfig = await hasAdoConfig();

  if (account && hasConfig) {
    hide("dashboard-not-connected");
    show("dashboard-connected");
    // Only load dashboard if we already have cached items (i.e. user has synced before)
    if (cachedWorkItems.length > 0) {
      await loadDashboard();
    } else {
      setStatus("Connected — click Sync to load data", "idle");
    }
  } else {
    show("dashboard-not-connected");
    hide("dashboard-connected");
  }
}

// ── Settings ───────────────────────────────────────────────────────────────

async function loadSettingsForm(): Promise<void> {
  const adoConfig = await getAdoConfig();
  if (adoConfig) {
    (document.getElementById("setting-org") as HTMLInputElement).value = adoConfig.organization;
    (document.getElementById("setting-project") as HTMLInputElement).value = adoConfig.project;
  }
}

async function handleSaveSettings(): Promise<void> {
  const pat = (document.getElementById("setting-pat") as HTMLInputElement).value.trim();
  const clientId = (document.getElementById("setting-client-id") as HTMLInputElement).value.trim();
  const tenantId = (document.getElementById("setting-tenant-id") as HTMLInputElement).value.trim();

  // Strip full URL if user pasted https://dev.azure.com/myorg
  let org = (document.getElementById("setting-org") as HTMLInputElement).value.trim();
  org = org.replace(/^https?:\/\/dev\.azure\.com\//i, "").replace(/\/$/, "");

  const project = (document.getElementById("setting-project") as HTMLInputElement).value.trim();

  // OAuth takes priority if configured, otherwise use PAT
  if (clientId) {
    await saveOAuthConfig(clientId, tenantId || undefined);
  } else if (pat) {
    await savePat(pat);
  }

  if (org && project) {
    await saveAdoConfig({ organization: org, project });
  }

  setResult("settings-status", "Settings saved successfully.", "success");
  // Auto-close modal after a short delay so the user sees the confirmation
  setTimeout(() => hide("settings-modal"), 800);
  await updateUI();
}

// ── Connect / Sign In ──────────────────────────────────────────────────────

async function handleConnect(): Promise<void> {
  if (!(await hasAuthConfig())) {
    show("settings-modal", "flex");
    return;
  }

  try {
    setStatus("Connecting...", "loading");
    const account = await signIn();
    setStatus(`Connected as ${account.displayName}`, "idle");
    await updateUI();
  } catch (error) {
    setStatus("Connection failed", "error");
    setResult("sync-result", `Error: ${error instanceof Error ? error.message : String(error)}`, "error");
  }
}

// ── Sync ───────────────────────────────────────────────────────────────────

async function handleSync(): Promise<void> {
  try {
    setStatus("Syncing work items...", "loading");
    const wiql = (document.getElementById("wiql-query") as HTMLTextAreaElement).value.trim() || undefined;
    const layoutBtn = document.querySelector(".toggle-btn.active[data-layout]") as HTMLElement | null;
    const layout = (layoutBtn?.dataset.layout as "flat" | "hierarchical") || "flat";
    const opts: SyncOptions = { layout, wiql };
    const count = await syncToExcel(opts);
    setStatus("Sync complete", "idle");
    setResult("sync-result", `✅ Synced ${count} work items to "PivotOps Data" sheet.`, "success");

    // Refresh analytics data and update dashboard
    cachedWorkItems = await fetchAnalyticsWorkItems();
    await loadDashboard();
  } catch (error) {
    setStatus("Sync failed", "error");
    const msg = error instanceof Error ? error.message : String(error);
    const lower = msg.toLowerCase();
    const hint = lower.includes("203") || (lower.includes("401") && lower.includes("scope"))
      ? " — Your PAT is missing required scopes. Regenerate it with: Work Items (Read & Write) + Analytics (Read)."
      : lower.includes("401") || lower.includes("unauthorized")
      ? " — PAT is invalid or expired. Go to Azure DevOps → User Settings → Personal Access Tokens and generate a new one with Work Items (Read & Write) and Analytics (Read) scopes."
      : lower.includes("403") || lower.includes("forbidden") || lower.includes("not authorized") || lower.includes("scope")
      ? " — PAT does not have the required permission. Required scopes: Work Items (Read & Write), Analytics (Read)."
      : lower.includes("404") || lower.includes("not found")
      ? " — Organization or Project name not found. Check spelling and case (e.g. 'myorg', not 'https://dev.azure.com/myorg')."
      : lower.includes("failed to fetch") || lower.includes("networkerror")
      ? " — Network error. Check your internet connection and that dev.azure.com is reachable."
      : "";
    setResult("sync-result", `❌ ${msg}${hint}`, "error");
  }
}

// ── Dashboard ──────────────────────────────────────────────────────────────

async function loadDashboard(): Promise<void> {
  try {
    setStatus("Loading dashboard...", "loading");

    if (cachedWorkItems.length === 0) {
      cachedWorkItems = await fetchAnalyticsWorkItems();
    }

    // Run summary analysis
    const summary = await runAnalysis<{
      total_items: number;
      by_state: Record<string, number>;
      by_type: Record<string, number>;
      total_points?: number;
      avg_points?: number;
      cycle_time?: { median: number; p85: number; mean: number };
    }>(ANALYSIS_SUMMARY, { items: cachedWorkItems });

    // Update stat cards
    $("stat-total").textContent = String(summary.total_items);
    $("stat-active").textContent = String(summary.by_state["Active"] || 0);
    $("stat-closed").textContent = String(summary.by_state["Closed"] || 0);
    $("stat-velocity").textContent = summary.avg_points ? `${summary.avg_points} pts` : "—";

    // Generate state distribution chart
    const chartBase64 = await runChart(CHART_STATE_DISTRIBUTION, { items: cachedWorkItems });
    $("dashboard-chart").innerHTML = `<img src="data:image/png;base64,${chartBase64}" alt="State Distribution"/>`;

    setStatus("Dashboard loaded", "idle");
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    const lower = msg.toLowerCase();
    const hint = lower.includes("401") || lower.includes("unauthorized") || lower.includes("scope")
      ? " — PAT missing Analytics (Read) scope. Regenerate your PAT with Analytics (Read) + Work Items (Read & Write)."
      : lower.includes("403") || lower.includes("forbidden")
      ? " — PAT does not have Analytics permission. Required scope: Analytics (Read)."
      : lower.includes("404") || lower.includes("not found")
      ? " — Analytics not enabled on this org, or org/project name is wrong."
      : lower.includes("failed to fetch")
      ? " — Network error reaching analytics.dev.azure.com."
      : "";
    setStatus("Dashboard error", "error");
    $(`dashboard-summary`).innerHTML = `<div class="result-box error" style="margin-top:12px;">❌ ${msg}${hint}</div>`;
    show("dashboard-summary");
  }
}

// ── Charts ─────────────────────────────────────────────────────────────────

const CHART_MAP: Record<string, { script: string; feature: string }> = {
  state_distribution: { script: CHART_STATE_DISTRIBUTION, feature: "state_distribution" },
  velocity: { script: CHART_VELOCITY, feature: "velocity_chart" },
  scope_drift: { script: CHART_SCOPE_DRIFT, feature: "scope_drift" },
  cycle_time: { script: CHART_CYCLE_TIME, feature: "cycle_time_histogram" },
  cumulative_flow: { script: CHART_CUMULATIVE_FLOW, feature: "cumulative_flow" },
  burndown: { script: CHART_BURNDOWN, feature: "burndown" },
  member_velocity: { script: CHART_MEMBER_VELOCITY, feature: "member_velocity" },
  estimation_accuracy: { script: CHART_ESTIMATION_ACCURACY, feature: "estimation_accuracy" },
};

async function handleChartSelect(chartId: string): Promise<void> {
  const chart = CHART_MAP[chartId];
  if (!chart) return;

  try {
    setStatus(`Generating ${chartId} chart...`, "running");
    $("chart-output").innerHTML = '<div class="spinner"></div>';

    if (cachedWorkItems.length === 0) {
      cachedWorkItems = await fetchAnalyticsWorkItems();
    }

    // Scope drift chart needs iteration date ranges
    const chartData: Record<string, unknown> = { items: cachedWorkItems };
    if (chartId === "scope_drift") {
      const iters = await getIterations();
      chartData.iterations = iters.map((it) => ({
        path: it.path,
        name: it.name,
        startDate: it.attributes.startDate,
        finishDate: it.attributes.finishDate,
      }));
    }

    const base64 = await runChart(chart.script, chartData);
    $("chart-output").innerHTML = `<img src="data:image/png;base64,${base64}" alt="${chartId}"/>`;
    currentChart = base64;
    show("btn-export-chart");
    setStatus("Chart generated", "idle");
  } catch (error) {
    setStatus("Chart error", "error");
    $("chart-output").innerHTML = `<div style="color:var(--accent-red); padding:20px;">Error: ${error instanceof Error ? error.message : String(error)}</div>`;
  }
}

function handleExportChart(): void {
  if (!currentChart) return;
  const link = document.createElement("a");
  link.href = `data:image/png;base64,${currentChart}`;
  link.download = "pivotops-chart.png";
  link.click();
}

// ── Predictive Analytics ───────────────────────────────────────────────────

async function handlePredictVelocity(): Promise<void> {
  try {
    setStatus("Loading predictive model...", "loading");
    await loadPredictivePackages();

    setStatus("Running prediction...", "running");
    if (cachedWorkItems.length === 0) {
      cachedWorkItems = await fetchAnalyticsWorkItems();
    }

    const result = await runAnalysis<{
      predictions?: Array<{ sprint: string; predicted_velocity: number; confidence_low: number; confidence_high: number }>;
      trend?: number;
      trend_direction?: string;
      average_velocity?: number;
      error?: string;
    }>(ANALYSIS_SPRINT_PREDICTION, { items: cachedWorkItems });

    if (result.error) {
      setResult("prediction-data", result.error, "error");
      show("prediction-result");
      return;
    }

    let html = "";
    if (result.predictions) {
      for (const p of result.predictions) {
        html += `
          <div class="prediction-row">
            <span class="prediction-label">${p.sprint}</span>
            <span class="prediction-value">${p.predicted_velocity} pts <span style="font-size:10px;color:var(--text-muted)">(${p.confidence_low}–${p.confidence_high})</span></span>
          </div>`;
      }
    }

    const trendClass = `trend-${result.trend_direction || "stable"}`;
    html += `
      <div class="prediction-row">
        <span class="prediction-label">Trend</span>
        <span class="prediction-value ${trendClass}">${result.trend_direction} (${result.trend! > 0 ? "+" : ""}${result.trend}/sprint)</span>
      </div>
      <div class="prediction-row">
        <span class="prediction-label">Average Velocity</span>
        <span class="prediction-value">${result.average_velocity} pts</span>
      </div>`;

    $("prediction-data").innerHTML = html;
    show("prediction-result");
    setStatus("Prediction complete", "idle");
  } catch (error) {
    setStatus("Prediction failed", "error");
    console.error(error);
  }
}

async function handleForecast(): Promise<void> {
  try {
    const backlogInput = (document.getElementById("backlog-points") as HTMLInputElement).value;
    const backlogPoints = parseInt(backlogInput, 10);
    if (!backlogPoints || backlogPoints <= 0) {
      setResult("forecast-result", "Please enter valid backlog points.", "error");
      show("forecast-result");
      return;
    }

    setStatus("Running Monte Carlo forecast...", "running");
    await loadPredictivePackages();

    if (cachedWorkItems.length === 0) {
      cachedWorkItems = await fetchAnalyticsWorkItems();
    }

    const result = await runAnalysis<{
      p50_sprints?: number;
      p85_sprints?: number;
      p95_sprints?: number;
      average_velocity?: number;
      error?: string;
    }>(ANALYSIS_COMPLETION_FORECAST, {
      items: cachedWorkItems,
      backlog_points: backlogPoints,
    });

    if (result.error) {
      setResult("forecast-result", result.error, "error");
      show("forecast-result");
      return;
    }

    $("forecast-result").innerHTML = `
      <h4>📅 Completion Forecast (${backlogPoints} pts)</h4>
      <div class="prediction-row">
        <span class="prediction-label">50% confidence</span>
        <span class="prediction-value">${result.p50_sprints} sprints</span>
      </div>
      <div class="prediction-row">
        <span class="prediction-label">85% confidence</span>
        <span class="prediction-value">${result.p85_sprints} sprints</span>
      </div>
      <div class="prediction-row">
        <span class="prediction-label">95% confidence</span>
        <span class="prediction-value">${result.p95_sprints} sprints</span>
      </div>
      <div class="prediction-row">
        <span class="prediction-label">Avg velocity used</span>
        <span class="prediction-value">${result.average_velocity} pts/sprint</span>
      </div>
    `;
    $("forecast-result").className = "prediction-card";
    show("forecast-result");
    setStatus("Forecast complete", "idle");
  } catch (error) {
    setStatus("Forecast failed", "error");
    console.error(error);
  }
}

// ── Write-back ───────────────────────────────────────────────────────────────

async function handleDryRun(): Promise<void> {
  try {
    setStatus("Validating changes...", "loading");

    // Read changes from the PivotOps Data sheet
    const updates = await readUpdatesFromSheet();

    if (updates.length === 0) {
      setResult("dry-run-result", "No changes detected in the sheet.", "info");
      return;
    }

    const validationResults = await validateUpdates(updates);
    const hasWarnings = validationResults.some((r) => r.warnings.length > 0);

    if (hasWarnings) {
      const msgs = validationResults
        .filter((r) => r.warnings.length > 0)
        .map((r) => `#${r.id}: ${r.warnings.join(", ")}`)
        .join("\n");
      setResult("dry-run-result", `⚠ Validation warnings:\n${msgs}`, "error");
    } else {
      setResult("dry-run-result", `✅ ${updates.length} updates validated. Ready to push.`, "success");
      ($("btn-push-updates") as HTMLButtonElement).disabled = false;
    }

    setStatus("Validation complete", "idle");
  } catch (error) {
    setStatus("Validation failed", "error");
    setResult("dry-run-result", `❌ ${error instanceof Error ? error.message : String(error)}`, "error");
  }
}

async function handlePushUpdates(): Promise<void> {
  try {
    setStatus("Pushing updates...", "loading");
    const updates = await readUpdatesFromSheet();

    const result = await bulkUpdateWorkItems(updates);

    let msg = `✅ ${result.succeeded.length} items updated.`;
    if (result.failed.length > 0) {
      msg += `\n❌ ${result.failed.length} failed: ${result.failed.map((f) => `#${f.id}: ${f.error}`).join(", ")}`;
    }

    setResult("update-result", msg, result.failed.length > 0 ? "error" : "success");
    ($("btn-push-updates") as HTMLButtonElement).disabled = true;
    setStatus("Update complete", "idle");
  } catch (error) {
    setStatus("Update failed", "error");
    setResult("update-result", `❌ ${error instanceof Error ? error.message : String(error)}`, "error");
  }
}

async function readUpdatesFromSheet(): Promise<Array<{ id: number; changes: WorkItemUpdate[] }>> {
  const updates: Array<{ id: number; changes: WorkItemUpdate[] }> = [];

  await Excel.run(async (ctx) => {
    let sheet: Excel.Worksheet;
    try {
      sheet = ctx.workbook.worksheets.getItem("PivotOps Data");
    } catch {
      return;
    }

    const usedRange = sheet.getUsedRange();
    usedRange.load("values");
    await ctx.sync();

    const values = usedRange.values;
    if (!values || values.length < 2) return;

    const headers = values[0] as string[];
    const idIdx = headers.indexOf("ID");
    const stateIdx = headers.indexOf("State");
    const titleIdx = headers.indexOf("Title");
    const assignedIdx = headers.indexOf("Assigned To");
    const descriptionIdx = headers.indexOf("Description");
    const priorityIdx = headers.indexOf("Priority");
    const storyPointsIdx = headers.indexOf("Story Points");
    const iterationIdx = headers.indexOf("Iteration");
    const areaPathIdx = headers.indexOf("Area Path");
    const tagsIdx = headers.indexOf("Tags");

    if (idIdx < 0) return;

    // Read original data from a hidden property (stored during sync)
    // For now, we'll detect changes by comparing with ADO data
    for (let row = 1; row < values.length; row++) {
      const id = Number(values[row][idIdx]);
      if (!id) continue;

      const changes: WorkItemUpdate[] = [];

      // Map of header index -> ADO field path for editable fields
      const editableFields: Array<[number, string]> = [
        [stateIdx, "/fields/System.State"],
        [titleIdx, "/fields/System.Title"],
        [assignedIdx, "/fields/System.AssignedTo"],
        [descriptionIdx, "/fields/System.Description"],
        [priorityIdx, "/fields/Microsoft.VSTS.Common.Priority"],
        [storyPointsIdx, "/fields/Microsoft.VSTS.Scheduling.StoryPoints"],
        [iterationIdx, "/fields/System.IterationPath"],
        [areaPathIdx, "/fields/System.AreaPath"],
        [tagsIdx, "/fields/System.Tags"],
      ];

      for (const [idx, fieldPath] of editableFields) {
        if (idx >= 0 && values[row][idx] != null && values[row][idx] !== "") {
          changes.push({ op: "replace", path: fieldPath, value: values[row][idx] });
        }
      }

      if (changes.length > 0) {
        updates.push({ id, changes });
      }
    }
  });

  return updates;
}
