# PivotOps — Azure DevOps Analytics for Excel

A modern, multi-platform Excel Add-in that connects to Azure DevOps, delivers beautiful Python-powered analytics, and supports two-way sync — all without a server.

## Architecture

```
┌─────────────────────────────────────────────────────┐
│                    Excel (Host)                      │
│  ┌───────────────────────────────────────────────┐  │
│  │              Office Web Add-in                 │  │
│  │  ┌──────────┐  ┌──────────┐  ┌─────────────┐ │  │
│  │  │ MSAL.js  │  │ Office.js│  │  Pyodide     │ │  │
│  │  │ (Auth)   │  │ (Excel)  │  │  (Python/    │ │  │
│  │  │          │  │          │  │   WASM)      │ │  │
│  │  └────┬─────┘  └──────────┘  └──────┬──────┘ │  │
│  │       │                              │        │  │
│  │       │  Azure AD Token              │        │  │
│  │       ▼                              ▼        │  │
│  │  ┌──────────┐              ┌──────────────┐   │  │
│  │  │ ADO REST │              │ pandas       │   │  │
│  │  │ API      │───data──────▶│ matplotlib   │   │  │
│  │  │ (Direct) │              │ scikit-learn │   │  │
│  │  └──────────┘              └──────────────┘   │  │
│  └───────────────────────────────────────────────┘  │
│                   No Server Required                 │
└─────────────────────────────────────────────────────┘
```

## Features

PivotOps is a free add-in. All features are included:

- **One-way sync**: Pull work items from ADO into Excel with color-coded rows, dropdowns, and frozen header
- **Two-way sync**: Push changes back to ADO with dry-run validation
- **State Distribution** chart (pie + bar)
- **Sprint Velocity** chart with trend line
- **Scope Drift** chart — plan vs. actual per sprint with drift % overlay
- **Cycle Time Histogram** with percentile lines (P50/P85/P95)
- **Cumulative Flow Diagram**
- **Sprint Burndown** chart
- **Member Velocity** chart
- **Estimation Accuracy** chart
- **Predictive Analytics**: Sprint velocity forecast using scikit-learn
- **Completion Forecast**: Monte Carlo simulation for backlog estimation
- **Bulk Updater** with dry-run validation
- **Custom WIQL queries**
- **PAT authentication** (zero setup)
- **OAuth/MSAL authentication** (optional, auto-refresh tokens)
- Export charts as PNG

## Getting Started

### Prerequisites
- Node.js 18+
- An Azure DevOps organization with Analytics enabled
- A **Personal Access Token** (PAT) from Azure DevOps

### 1. Install Dependencies

```bash
npm install
```

### 2. Get your PAT

1. Go to Azure DevOps → **User Settings** → **Personal Access Tokens**
2. Click **New Token**
3. Set scopes: **Work Items: Read & Write**
4. Copy the token — you'll paste it into PivotOps Settings

> **Optional:** You can use OAuth instead of PAT. Register an Azure AD app with redirect URI `https://localhost:3000/taskpane.html` and scope `Azure DevOps → user_impersonation`, then enter the Client ID in Settings.

### 3. Start Development Server

```bash
npm run dev
```

### 4. Sideload in Excel

- **Desktop**: Use `npm run sideload` (requires office-addin-debugging)
- **Web**: Upload `manifest.xml` via Insert > Office Add-ins > Upload My Add-in

### 5. Configure in the Add-in

1. Click **⚙ Settings** in the task pane
2. Paste your **Personal Access Token**
3. Enter your **Organization** and **Project**
4. Click **Save Settings**, then **Connect**

## Project Structure

```
src/
├── taskpane/
│   ├── taskpane.html          # Main UI
│   ├── taskpane.css           # Dark theme styles
│   └── taskpane.ts            # UI logic & event wiring
├── commands/
│   ├── commands.html          # Ribbon command page
│   └── commands.ts            # Ribbon button handlers
├── services/
│   ├── auth.ts                # PAT / MSAL.js OAuth auth
│   ├── ado-client.ts          # ADO REST API client (direct browser calls)
│   ├── python-engine.ts       # Pyodide wrapper (Python in WASM)
│   └── storage.ts             # OfficeRuntime.storage abstraction
├── python/
│   └── charts.ts              # Python scripts for charts & analytics
index.html                     # Entry point (redirects to taskpane.html)
privacy.html                   # GDPR-compliant privacy policy
support.html                   # Support & FAQ page
manifest.xml                   # Office Add-in manifest
webpack.config.js              # Build config
azure-pipelines.yml            # CI/CD pipeline → Azure Static Web Apps
```

## Deployment

The add-in is entirely static (HTML/JS/CSS). It is hosted on **Azure Static Web Apps** and deployed automatically via the Azure DevOps pipeline.

```bash
# Production build (minified, no obfuscation)
npm run build
```

The pipeline (`azure-pipelines.yml`) replaces `localhost:3000` with the production URL, builds, and deploys to Azure Static Web Apps. Trigger it manually from the master branch in Azure DevOps.

## Tech Stack

| Component | Technology |
|-----------|-----------|
| Platform | Office Web Add-in (Office.js) |
| Language | TypeScript |
| Auth | PAT (default) / MSAL.js OAuth (optional) |
| Python | Pyodide (WebAssembly) |
| Charts | matplotlib + seaborn (via Pyodide) |
| ML | scikit-learn (via Pyodide) |
| Build | Webpack + Terser (minification) |
| Hosting | Azure Static Web Apps |
| CI/CD | Azure DevOps Pipelines |
| Distribution | Free — sideload via manifest |
