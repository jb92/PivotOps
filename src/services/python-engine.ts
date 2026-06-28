/**
 * PivotOps — Pyodide Python Engine
 * Runs Python (pandas, matplotlib, scikit-learn) entirely in the browser via WebAssembly.
 * No server required. Pyodide is loaded from CDN and cached by the browser.
 */

// Pyodide types (loaded dynamically)
interface PyodideInterface {
  runPythonAsync(code: string): Promise<unknown>;
  loadPackage(packages: string | string[]): Promise<void>;
  globals: {
    get(name: string): unknown;
    set(name: string, value: unknown): void;
  };
  FS: {
    writeFile(path: string, data: string | Uint8Array): void;
    readFile(path: string, opts: { encoding: string }): string;
  };
}

declare function loadPyodide(options?: {
  indexURL?: string;
  stdout?: (text: string) => void;
  stderr?: (text: string) => void;
}): Promise<PyodideInterface>;

const PYODIDE_CDN = "https://cdn.jsdelivr.net/pyodide/v0.27.0/full/";

let pyodideInstance: PyodideInterface | null = null;
let loadingPromise: Promise<PyodideInterface> | null = null;

export type EngineStatus = "idle" | "loading" | "ready" | "running" | "error";
type StatusListener = (status: EngineStatus, detail?: string) => void;

const listeners: StatusListener[] = [];
let currentStatus: EngineStatus = "idle";

function setStatus(status: EngineStatus, detail?: string): void {
  currentStatus = status;
  listeners.forEach((fn) => fn(status, detail));
}

export function onStatusChange(listener: StatusListener): () => void {
  listeners.push(listener);
  listener(currentStatus);
  return () => {
    const idx = listeners.indexOf(listener);
    if (idx >= 0) listeners.splice(idx, 1);
  };
}

export function getStatus(): EngineStatus {
  return currentStatus;
}

/**
 * Initialize Pyodide. Safe to call multiple times — will return cached instance.
 */
export async function initPyodide(): Promise<PyodideInterface> {
  if (pyodideInstance) return pyodideInstance;
  if (loadingPromise) return loadingPromise;

  loadingPromise = (async () => {
    setStatus("loading", "Loading Python runtime...");

    // Load Pyodide script if not already loaded
    if (typeof (globalThis as any).loadPyodide === "undefined") {
      await new Promise<void>((resolve, reject) => {
        const script = document.createElement("script");
        script.src = `${PYODIDE_CDN}pyodide.js`;
        script.onload = () => resolve();
        script.onerror = () => reject(new Error("Failed to load Pyodide from CDN"));
        document.head.appendChild(script);
      });
    }

    setStatus("loading", "Initializing Python...");
    const pyodide = await loadPyodide({ indexURL: PYODIDE_CDN });

    setStatus("loading", "Installing analytics packages...");
    await pyodide.loadPackage(["pandas", "matplotlib", "micropip"]);

    // Install additional packages via micropip
    await pyodide.runPythonAsync(`
import micropip
await micropip.install(['seaborn'])
    `);

    setStatus("ready");
    pyodideInstance = pyodide;
    return pyodide;
  })();

  try {
    return await loadingPromise;
  } catch (error) {
    setStatus("error", error instanceof Error ? error.message : String(error));
    loadingPromise = null;
    throw error;
  }
}

/**
 * Install scikit-learn (for predictive analytics).
 * Called on-demand to avoid loading it unless needed.
 */
export async function loadPredictivePackages(): Promise<void> {
  const pyodide = await initPyodide();
  setStatus("loading", "Installing predictive analytics...");
  await pyodide.loadPackage(["scikit-learn"]);
  setStatus("ready");
}

/**
 * Run a Python script with data injected as a JSON string.
 * Returns the result as a JSON-serializable object.
 */
export async function runPython<T = unknown>(
  script: string,
  data?: Record<string, unknown>
): Promise<T> {
  const pyodide = await initPyodide();
  setStatus("running");

  try {
    // Inject data into Python namespace
    if (data) {
      const jsonStr = JSON.stringify(data);
      pyodide.globals.set("__pivotops_input__", jsonStr);
      await pyodide.runPythonAsync(`
import json
__data__ = json.loads(__pivotops_input__)
      `);
    }

    const result = await pyodide.runPythonAsync(script);
    setStatus("ready");
    return result as T;
  } catch (error) {
    setStatus("error", error instanceof Error ? error.message : String(error));
    throw error;
  }
}

/**
 * Run a Python chart script and return the chart as a base64-encoded PNG.
 */
export async function runChart(
  script: string,
  data?: Record<string, unknown>
): Promise<string> {
  const pyodide = await initPyodide();
  setStatus("running");

  try {
    if (data) {
      const jsonStr = JSON.stringify(data);
      pyodide.globals.set("__pivotops_input__", jsonStr);
    }

    // Wrap the chart script to capture output as base64 PNG
    const wrappedScript = `
import json
import io
import base64
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

if '__pivotops_input__' in dir():
    __data__ = json.loads(__pivotops_input__)

# Clear any previous figures
plt.close('all')

${script}

# Capture the current figure as base64 PNG
__buf__ = io.BytesIO()
plt.savefig(__buf__, format='png', dpi=150, bbox_inches='tight', facecolor='#1e1e2e', edgecolor='none')
__buf__.seek(0)
__result__ = base64.b64encode(__buf__.read()).decode('utf-8')
plt.close('all')
__result__
    `;

    const result = await pyodide.runPythonAsync(wrappedScript);
    setStatus("ready");
    return result as string;
  } catch (error) {
    setStatus("error", error instanceof Error ? error.message : String(error));
    throw error;
  }
}

/**
 * Run a Python script that returns JSON data (for analytics results).
 */
export async function runAnalysis<T = unknown>(
  script: string,
  data?: Record<string, unknown>
): Promise<T> {
  const pyodide = await initPyodide();
  setStatus("running");

  try {
    if (data) {
      const jsonStr = JSON.stringify(data);
      pyodide.globals.set("__pivotops_input__", jsonStr);
    }

    const wrappedScript = `
import json

if '__pivotops_input__' in dir():
    __data__ = json.loads(__pivotops_input__)

${script}

# The script must set __result__ as a JSON-serializable value
import math as __math__
def __nan_to_none__(obj):
    if isinstance(obj, float) and __math__.isnan(obj):
        return None
    if isinstance(obj, dict):
        return {k: __nan_to_none__(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [__nan_to_none__(v) for v in obj]
    return obj
json.dumps(__nan_to_none__(__result__))
    `;

    const jsonResult = await pyodide.runPythonAsync(wrappedScript);
    setStatus("ready");
    return JSON.parse(jsonResult as string) as T;
  } catch (error) {
    setStatus("error", error instanceof Error ? error.message : String(error));
    throw error;
  }
}
