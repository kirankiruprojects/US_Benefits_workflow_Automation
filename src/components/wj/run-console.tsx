import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  AlertCircle,
  AlertTriangle,
  CheckCircle2,
  Copy,
  Check,
  FolderOpen,
  Info,
  Terminal,
  XCircle,
  Search,
} from "lucide-react";
import { LogLine, LogTag, StatusKind } from "@/lib/wj-bridge";
import { cn } from "@/lib/utils";
import { WjButton } from "./primitives";

export function DesktopNotice({ visible }: { visible: boolean }) {
  if (!visible) return null;
  return (
    <div className="flex items-start gap-3 rounded-xl border border-warning/40 bg-warning/10 p-4 text-warning-foreground">
      <AlertCircle className="size-5 shrink-0 text-warning mt-0.5" />
      <div className="text-xs leading-relaxed">
        <p className="font-bold text-foreground">Standalone Web Mode Active</p>
        <p className="mt-0.5 text-muted-foreground">
          You are viewing the web interface without the Workforce Junction desktop launcher. Automated file operations, Selenium sessions and Excel writers require launching via <code className="font-mono font-bold text-primary">run-desktop-app.bat</code>.
        </p>
      </div>
    </div>
  );
}

export function FilePickerRow({
  value,
  onBrowse,
  disabled = false,
  placeholder = "Select or enter file path…",
}: {
  value: string;
  onBrowse: () => void;
  disabled?: boolean;
  placeholder?: string;
}) {
  return (
    <div className="flex items-center gap-2">
      <input
        type="text"
        readOnly
        value={value}
        placeholder={placeholder}
        className="flex h-10 w-full rounded-lg border border-input bg-muted/40 px-3.5 py-2 text-xs font-mono text-foreground placeholder:text-muted-foreground focus-visible:outline-none disabled:cursor-not-allowed disabled:opacity-50"
      />
      <WjButton
        type="button"
        variant="secondary"
        onClick={onBrowse}
        disabled={disabled}
        className="shrink-0 text-xs font-bold"
      >
        <FolderOpen className="size-3.5 text-primary" />
        Browse…
      </WjButton>
    </div>
  );
}

export function StatusText({
  status,
  kind = "idle",
}: {
  status: string;
  kind?: StatusKind;
}) {
  const icon = {
    pass: <CheckCircle2 className="size-4 text-emerald-400 shrink-0" />,
    fail: <XCircle className="size-4 text-rose-400 shrink-0" />,
    idle: <Info className="size-4 text-sky-400 shrink-0" />,
  }[kind];

  const toneStyle = {
    pass: "border-emerald-500/40 bg-emerald-500/10 text-emerald-300 shadow-sm",
    fail: "border-rose-500/40 bg-rose-500/10 text-rose-300 shadow-sm",
    idle: "border-border bg-secondary/80 text-foreground",
  }[kind];

  return (
    <div className={cn("inline-flex items-center gap-2 rounded-lg border px-3 py-1.5 text-xs font-semibold max-w-full", toneStyle)}>
      {icon}
      <span className="truncate text-xs font-bold">{status}</span>
    </div>
  );
}

export function WjProgress({ value }: { value: number }) {
  const clamped = Math.min(100, Math.max(0, value));
  return (
    <div className="space-y-1.5">
      <div className="flex justify-between text-[0.68rem] font-bold text-muted-foreground uppercase tracking-wider">
        <span>Workflow Progress</span>
        <span className="font-mono text-foreground font-bold">{Math.round(clamped)}%</span>
      </div>
      <div className="h-2.5 w-full overflow-hidden rounded-full bg-secondary/80 border border-border/50">
        <div
          className="h-full gradient-teal transition-all duration-300 ease-out shadow-sm"
          style={{ width: `${clamped}%` }}
        />
      </div>
    </div>
  );
}

type FilterType = "all" | "pass" | "fail" | "warn";

export function LogConsole({
  lines,
  className,
  emptyText = "No log output recorded yet. Output and real-time execution logs will appear here.",
}: {
  lines: LogLine[];
  className?: string;
  emptyText?: string;
}) {
  const bottomRef = useRef<HTMLDivElement>(null);
  const [filter, setFilter] = useState<FilterType>("all");
  const [search, setSearch] = useState("");
  const [copied, setCopied] = useState(false);
  const [autoScroll, setAutoScroll] = useState(true);

  const stats = useMemo(() => {
    let pass = 0;
    let fail = 0;
    let warn = 0;
    lines.forEach((l) => {
      if (l.tag === "pass" || l.tag === "ok") pass++;
      else if (l.tag === "fail" || l.tag === "err") fail++;
      else if (l.tag === "warn") warn++;
    });
    return { pass, fail, warn, total: lines.length };
  }, [lines]);

  const filteredLines = useMemo(() => {
    return lines.filter((l) => {
      if (filter === "pass" && l.tag !== "pass" && l.tag !== "ok") return false;
      if (filter === "fail" && l.tag !== "fail" && l.tag !== "err") return false;
      if (filter === "warn" && l.tag !== "warn") return false;
      if (search.trim() && !l.message.toLowerCase().includes(search.toLowerCase())) return false;
      return true;
    });
  }, [lines, filter, search]);

  useEffect(() => {
    if (autoScroll) {
      bottomRef.current?.scrollIntoView({ behavior: "smooth" });
    }
  }, [filteredLines, autoScroll]);

  const copyLogs = () => {
    const text = lines.map((l) => `[${l.time}] [${l.tag.toUpperCase()}] ${l.message}`).join("\n");
    navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className={cn("relative flex flex-col rounded-xl border border-sidebar-border bg-sidebar font-mono text-sidebar-foreground shadow-xl overflow-hidden", className)}>
      {/* Console Top Header Bar */}
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-sidebar-border/70 bg-black/30 px-4 py-2.5">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 text-xs font-bold text-sidebar-foreground">
            <Terminal className="size-4 text-primary" />
            <span>Execution Console</span>
          </div>

          {/* Real-time Summary Counters */}
          <div className="flex items-center gap-1.5 text-[0.7rem] font-bold">
            <button
              type="button"
              onClick={() => setFilter("all")}
              className={cn(
                "rounded-md px-2 py-0.5 transition-colors cursor-pointer",
                filter === "all"
                  ? "bg-white/15 text-sidebar-foreground ring-1 ring-white/30"
                  : "bg-white/5 text-sidebar-muted hover:bg-white/10"
              )}
            >
              All: {stats.total}
            </button>
            <button
              type="button"
              onClick={() => setFilter("pass")}
              className={cn(
                "inline-flex items-center gap-1 rounded-md px-2 py-0.5 transition-colors cursor-pointer",
                filter === "pass"
                  ? "bg-emerald-500/30 text-emerald-300 ring-1 ring-emerald-500/50"
                  : "bg-emerald-500/10 text-emerald-400/80 hover:bg-emerald-500/20"
              )}
            >
              <CheckCircle2 className="size-3" />
              Passed: {stats.pass}
            </button>
            <button
              type="button"
              onClick={() => setFilter("fail")}
              className={cn(
                "inline-flex items-center gap-1 rounded-md px-2 py-0.5 transition-colors cursor-pointer",
                filter === "fail"
                  ? "bg-rose-500/30 text-rose-300 ring-1 ring-rose-500/50"
                  : "bg-rose-500/10 text-rose-400/80 hover:bg-rose-500/20"
              )}
            >
              <XCircle className="size-3" />
              Failed: {stats.fail}
            </button>
            {stats.warn > 0 && (
              <button
                type="button"
                onClick={() => setFilter("warn")}
                className={cn(
                  "inline-flex items-center gap-1 rounded-md px-2 py-0.5 transition-colors cursor-pointer",
                  filter === "warn"
                    ? "bg-amber-500/30 text-amber-300 ring-1 ring-amber-500/50"
                    : "bg-amber-500/10 text-amber-400/80 hover:bg-amber-500/20"
                )}
              >
                <AlertTriangle className="size-3" />
                Warnings: {stats.warn}
              </button>
            )}
          </div>
        </div>

        {/* Action Controls */}
        <div className="flex items-center gap-2">
          {lines.length > 5 && (
            <div className="relative">
              <Search className="size-3 text-sidebar-muted absolute left-2 top-2" />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Filter logs…"
                className="h-7 w-28 sm:w-36 rounded-md bg-black/40 pl-6 pr-2 text-[0.7rem] text-sidebar-foreground placeholder:text-sidebar-muted/60 focus:outline-none focus:ring-1 focus:ring-primary"
              />
            </div>
          )}

          <button
            type="button"
            onClick={copyLogs}
            disabled={lines.length === 0}
            title="Copy logs to clipboard"
            className="inline-flex items-center gap-1 rounded-md bg-white/5 px-2.5 py-1 text-[0.7rem] font-bold text-sidebar-muted transition-colors hover:bg-white/10 hover:text-sidebar-foreground disabled:opacity-40 cursor-pointer"
          >
            {copied ? <Check className="size-3 text-emerald-400" /> : <Copy className="size-3" />}
            {copied ? "Copied!" : "Copy"}
          </button>
        </div>
      </div>

      {/* Console Scrollable Body */}
      <div className="log-scroll h-72 overflow-y-auto p-3.5 text-xs space-y-1 bg-black/40">
        {filteredLines.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-10 text-sidebar-muted/60 text-xs">
            <Terminal className="size-7 mb-2 opacity-40" />
            <p className="italic">{search || filter !== "all" ? "No matching log entries found for active filter." : emptyText}</p>
          </div>
        ) : (
          filteredLines.map((line) => {
            const isPass = line.tag === "pass" || line.tag === "ok" || line.message.startsWith("PASS");
            const isFail = line.tag === "fail" || line.tag === "err" || line.message.startsWith("FAIL") || line.message.startsWith("Error");
            const isWarn = line.tag === "warn" || line.message.startsWith("WARN");
            const isHeader = line.message.startsWith("===") || line.message.startsWith("RESULTS SUMMARY");

            return (
              <div
                key={line.id}
                className={cn(
                  "group flex items-start gap-2.5 px-2 py-1 rounded transition-colors leading-relaxed",
                  isFail
                    ? "bg-rose-500/10 border-l-2 border-rose-500 text-rose-300 hover:bg-rose-500/15"
                    : isPass
                    ? "bg-emerald-500/10 border-l-2 border-emerald-500 text-emerald-300 hover:bg-emerald-500/15"
                    : isWarn
                    ? "bg-amber-500/10 border-l-2 border-amber-500 text-amber-300 hover:bg-amber-500/15"
                    : isHeader
                    ? "text-sidebar-primary font-bold py-1.5"
                    : "text-sidebar-foreground/90 hover:bg-white/[0.03]"
                )}
              >
                {/* Timestamp */}
                <span className="text-[0.68rem] text-sidebar-muted/60 shrink-0 select-none pt-0.5">
                  [{line.time}]
                </span>

                {/* Status Badge */}
                {isPass && (
                  <span className="inline-flex shrink-0 items-center gap-0.5 rounded bg-emerald-500/20 px-1.5 py-0.5 text-[0.65rem] font-extrabold text-emerald-300 uppercase tracking-wider">
                    <Check className="size-3" /> PASS
                  </span>
                )}
                {isFail && (
                  <span className="inline-flex shrink-0 items-center gap-0.5 rounded bg-rose-500/20 px-1.5 py-0.5 text-[0.65rem] font-extrabold text-rose-300 uppercase tracking-wider">
                    <XCircle className="size-3" /> FAIL
                  </span>
                )}
                {isWarn && (
                  <span className="inline-flex shrink-0 items-center gap-0.5 rounded bg-amber-500/20 px-1.5 py-0.5 text-[0.65rem] font-extrabold text-amber-300 uppercase tracking-wider">
                    <AlertTriangle className="size-3" /> WARN
                  </span>
                )}

                {/* Message Body */}
                <span className="break-all whitespace-pre-wrap flex-1 text-[0.8rem] font-mono">
                  {line.message}
                </span>
              </div>
            );
          })
        )}
        <div ref={bottomRef} />
      </div>
    </div>
  );
}
