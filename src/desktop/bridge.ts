/**
 * PivotOps — WebView2 host bridge.
 *
 * Request/response transport between the web app running inside the WinUI shell
 * and the native host. Messages are exchanged as JSON strings over
 * `window.chrome.webview`.
 */

interface HostResponse {
  id: number;
  ok: boolean;
  result?: unknown;
  error?: string;
}

interface WebViewHost {
  postMessage(message: string): void;
  addEventListener(type: "message", listener: (event: { data: unknown }) => void): void;
}

const pending = new Map<number, { resolve: (value: never) => void; reject: (error: Error) => void }>();
let nextId = 1;
let listening = false;

function host(): WebViewHost | undefined {
  return (window as unknown as { chrome?: { webview?: WebViewHost } }).chrome?.webview;
}

export function isDesktopHost(): boolean {
  return !!host();
}

function listen(wv: WebViewHost): void {
  if (listening) return;
  listening = true;
  wv.addEventListener("message", (event) => {
    let message: HostResponse;
    try {
      message = typeof event.data === "string" ? JSON.parse(event.data) : (event.data as HostResponse);
    } catch {
      return;
    }
    const entry = pending.get(message.id);
    if (!entry) return;
    pending.delete(message.id);
    if (message.ok) entry.resolve(message.result as never);
    else entry.reject(new Error(message.error || "PivotOps host returned an error."));
  });
}

export function invoke<T>(kind: string, payload?: unknown): Promise<T> {
  const wv = host();
  if (!wv) {
    return Promise.reject(
      new Error("PivotOps desktop host is unavailable. Run this build inside the PivotOps app."),
    );
  }
  listen(wv);
  const id = nextId++;
  return new Promise<T>((resolve, reject) => {
    pending.set(id, { resolve: resolve as (value: never) => void, reject });
    wv.postMessage(JSON.stringify({ id, kind, payload: payload ?? null }));
  });
}
