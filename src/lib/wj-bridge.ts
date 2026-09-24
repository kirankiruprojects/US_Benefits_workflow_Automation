/**
 * Bridge to the desktop (pywebview) Python backend.
 *
 * IMPORTANT: the API method names and argument shapes below are exactly the
 * ones the original Workforce Junction pages used, so the existing Python
 * automation keeps working unchanged. Only the UI around them was rebuilt.
 */
import { useCallback, useEffect, useRef, useState } from "react";

export type LogTag = "info" | "ok" | "pass" | "fail" | "err" | "warn";

export interface LogLine {
  id: number;
  message: string;
  tag: LogTag;
  time: string;
}

export type StatusKind = "idle" | "pass" | "fail";

/* eslint-disable @typescript-eslint/no-explicit-any */
type AnyWindow = Window & { pywebview?: { api?: Record<string, any> } };

function topWindow(): AnyWindow | undefined {
  if (typeof window === "undefined") return undefined;
  try {
    return (window.top as AnyWindow) ?? (window as AnyWindow);
  } catch {
    return window as AnyWindow;
  }
}

export function getApi(): Record<string, any> | undefined {
  // Reading properties off window.top throws a SecurityError when the app is
  // embedded in a cross-origin frame (e.g. Kiran preview), so guard it
  // and fall back to this frame's own window.
  try {
    const api = topWindow()?.pywebview?.api;
    if (api) return api;
  } catch {
    /* cross-origin parent — ignore */
  }
  try {
    return (window as AnyWindow)?.pywebview?.api;
  } catch {
    return undefined;
  }
}

export function hasApi(): boolean {
  return Boolean(getApi());
}

/** True once the desktop Python backend is reachable (checked for ~2s). */
export function useDesktopApi() {
  const [ready, setReady] = useState(false);
  const [checked, setChecked] = useState(false);

  useEffect(() => {
    const check = () => {
      if (hasApi()) {
        setReady(true);
        setChecked(true);
        return true;
      }
      return false;
    };

    if (check()) return;

    const onPyWebviewReady = () => {
      check();
    };

    window.addEventListener("pywebviewready", onPyWebviewReady);

    let elapsed = 0;
    const tick = window.setInterval(() => {
      elapsed += 200;
      if (check() || elapsed >= 5000) {
        setChecked(true);
        window.clearInterval(tick);
      }
    }, 200);

    return () => {
      window.removeEventListener("pywebviewready", onPyWebviewReady);
      window.clearInterval(tick);
    };
  }, []);

  return { ready, checked };
}

const NORMALIZE_TAG: Record<string, LogTag> = {
  info: "info",
  ok: "ok",
  pass: "pass",
  fail: "fail",
  err: "fail",
  error: "fail",
  warn: "warn",
};

/**
 * Registers the window.wjLog / wjStatus / wjProgress callbacks that the Python
 * side calls through window.evaluate_js, plus any extra "finished" callbacks.
 */
export function useRunConsole(finishedCallbacks: string[] = []) {
  const [lines, setLines] = useState<LogLine[]>([]);
  const [status, setStatusState] = useState("Ready");
  const [statusKind, setStatusKind] = useState<StatusKind>("idle");
  const [progress, setProgress] = useState(0);
  const [running, setRunning] = useState(false);
  const counter = useRef(0);
  const keys = finishedCallbacks.join("|");

  const appendLog = useCallback((message: string, tag?: string) => {
    counter.current += 1;
    const msg = String(message);
    let resolvedTag: LogTag = "info";
    const explicit = tag ? String(tag).toLowerCase().trim() : "";

    if (explicit === "pass" || explicit === "ok") {
      resolvedTag = "pass";
    } else if (explicit === "fail" || explicit === "err" || explicit === "error") {
      resolvedTag = "fail";
    } else if (explicit === "warn" || explicit === "warning") {
      resolvedTag = "warn";
    } else {
      const lower = msg.toLowerCase();
      if (
        msg.startsWith("FAIL") ||
        msg.startsWith("✗") ||
        msg.startsWith("[✗]") ||
        lower.startsWith("error:") ||
        lower.includes("error in ") ||
        lower.includes("failed:") ||
        lower.includes("exception:") ||
        lower.includes("timed out")
      ) {
        resolvedTag = "fail";
      } else if (
        msg.startsWith("PASS") ||
        msg.startsWith("✓") ||
        msg.startsWith("[✓]") ||
        lower.includes("authenticated successfully") ||
        lower.includes("password setup completed") ||
        lower.includes("all records passed")
      ) {
        resolvedTag = "pass";
      } else if (msg.startsWith("WARN") || msg.startsWith("⚠") || lower.startsWith("warning:")) {
        resolvedTag = "warn";
      }
    }

    setLines((prev) => [
      ...prev,
      {
        id: counter.current,
        message: msg,
        tag: resolvedTag,
        time: new Date().toLocaleTimeString([], {
          hour: "2-digit",
          minute: "2-digit",
          second: "2-digit",
        }),
      },
    ]);
  }, []);

  const setStatus = useCallback((message: string, kind?: string) => {
    setStatusState(String(message));
    setStatusKind(kind === "fail" || kind === "err" ? "fail" : kind === "pass" || kind === "ok" ? "pass" : "idle");
  }, []);

  const clear = useCallback(() => {
    setLines([]);
    setProgress(0);
    setStatus("Log cleared.");
  }, [setStatus]);

  useEffect(() => {
    const targets: any[] = [];
    const own = typeof window === "undefined" ? undefined : (window as any);
    if (own) targets.push(own);
    try {
      const top = topWindow() as any;
      if (top && top !== own) {
        // touching a cross-origin parent throws — probe before using it
        void top.document;
        targets.push(top);
      }
    } catch {
      /* cross-origin parent — only register on this frame */
    }

    const names = keys ? keys.split("|") : [];
    targets.forEach((w) => {
      try {
        w.wjLog = appendLog;
        w.wjStatus = setStatus;
        w.wjProgress = (pct: number) => setProgress(Number(pct) || 0);
        names.forEach((name) => {
          w[name] = () => setRunning(false);
        });
      } catch {
        /* ignore */
      }
    });

    return () => {
      targets.forEach((w) => {
        try {
          delete w.wjLog;
          delete w.wjStatus;
          delete w.wjProgress;
          names.forEach((name) => delete w[name]);
        } catch {
          /* ignore */
        }
      });
    };
  }, [appendLog, setStatus, keys]);

  return {
    lines,
    status,
    statusKind,
    progress,
    running,
    setRunning,
    setProgress,
    appendLog,
    setStatus,
    clear,
  };
}

export async function pickExcelFile(): Promise<string | null> {
  const api = getApi();
  if (!api?.["pick_excel_file"]) return null;
  return (await api["pick_excel_file"]()) ?? null;
}

export async function pickDriverFile(): Promise<string | null> {
  const api = getApi();
  if (!api?.["pick_driver_file"]) return null;
  return (await api["pick_driver_file"]()) ?? null;
}

/** Calls a backend method by name; resolves false when unavailable. */
export async function callApi(method: string, ...args: unknown[]): Promise<boolean> {
  const api = getApi();
  if (!api?.[method]) return false;
  await api[method](...args);
  return true;
}

export async function launchScreenshotTool(): Promise<boolean> {
  const api = getApi();
  if (!api?.["launch_screenshot_tool"]) return false;
  await api["launch_screenshot_tool"]();
  return true;
}

export function fileName(path: string): string {
  return path.split(/[\\/]/).pop() ?? path;
}
