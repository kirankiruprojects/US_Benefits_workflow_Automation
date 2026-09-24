import { createFileRoute } from "@tanstack/react-router";
import {
  Camera,
  CirclePlay,
  Eraser,
  FileCheck2,
  FileText,
  Keyboard,
  Layers,
  Monitor,
  Sparkles,
  Undo2,
} from "lucide-react";
import { useState } from "react";
import { toast } from "sonner";
import { PageHeader } from "@/components/wj/shell";
import {
  ProcedurePanel,
  ProcedureStep,
  WjButton,
  WjCard,
  WjCardBody,
  WjCardHeader,
} from "@/components/wj/primitives";
import { DesktopNotice, LogConsole, StatusText, WjProgress } from "@/components/wj/run-console";
import { launchScreenshotTool, useDesktopApi, useRunConsole } from "@/lib/wj-bridge";

export const Route = createFileRoute("/screenshot")({
  head: () => ({
    meta: [
      { title: "Screenshot — Workforce Junction" },
      {
        name: "description",
        content:
          "Workflow capture and screenshot documentation tool for QA testers — capture screens, log issues and notes, and auto-build Word review reports.",
      },
      { property: "og:title", content: "Screenshot — Workforce Junction" },
      { property: "og:description", content: "Workflow capture and screenshot documentation tool." },
    ],
  }),
  component: ScreenshotPage,
});

function ScreenshotPage() {
  const { ready, checked } = useDesktopApi();
  const runConsole = useRunConsole();
  const [launching, setLaunching] = useState(false);

  async function handleLaunch() {
    setLaunching(true);
    runConsole.setStatus("Launching Screenshot Tool...", "idle");
    runConsole.appendLog("Launching Workflow Capture (Screenshot) Tool...", "info");

    try {
      const ok = await launchScreenshotTool();
      if (ok) {
        toast.success("Workflow Capture (Screenshot) tool launched!");
        runConsole.setStatus("Screenshot Tool is running", "pass");
        runConsole.appendLog(
          "Screenshot tool is active. Press K or click Capture to take screenshots.",
          "pass"
        );
      } else {
        toast.error("Desktop backend not available.");
        runConsole.setStatus("Failed to launch Screenshot Tool", "fail");
        runConsole.appendLog("Desktop backend API is not connected.", "fail");
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      toast.error(`Error launching Screenshot tool: ${msg}`);
      runConsole.setStatus("Launch error", "fail");
      runConsole.appendLog(`Error: ${msg}`, "fail");
    } finally {
      setLaunching(false);
    }
  }

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Workflow Capture & QA Tools"
        icon={<Camera className="size-3.5" />}
        title="Screenshot"
        subtitle="Always-on-top workflow capture toolbar for QA testers — screenshot active monitors, document issues/notes, and auto-generate formatted Word test reports."
      />

      <DesktopNotice visible={checked && !ready} />

      {/* Workflow Capture Launch Card */}
      <WjCard>
        <WjCardHeader
          title="Workflow Capture Tool"
          hint="Always-on-top QA capture toolbar & Word report builder"
          actions={
            <span className="inline-flex items-center gap-1.5 rounded-full border border-border/80 bg-secondary/80 px-2.5 py-1 text-xs font-medium text-foreground">
              <Sparkles className="size-3 text-primary" />
              Word Review Report (.docx)
            </span>
          }
        />
        <WjCardBody>
          <div className="flex flex-col gap-5 sm:flex-row sm:items-center sm:justify-between">
            <div className="space-y-1.5">
              <h3 className="text-base font-bold text-foreground">Launch Desktop Toolbar</h3>
              <p className="text-xs text-muted-foreground leading-relaxed max-w-xl">
                Opens the lightweight floating capture toolbar. Works seamlessly across multiple monitors
                with automatic cursor detection, real-time issue indexing, and one-click undo support.
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-3 shrink-0">
              <WjButton
                size="lg"
                onClick={handleLaunch}
                disabled={launching}
                className="gap-2.5 shadow-md"
              >
                <Camera className="size-4" />
                {launching ? "Launching Toolbar..." : "Launch Screenshot Tool"}
              </WjButton>
            </div>
          </div>

          {/* Feature Badges Grid */}
          <div className="mt-5 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 pt-4 border-t border-border/60">
            <div className="flex items-start gap-2.5 rounded-lg border border-border/60 bg-secondary/30 p-2.5">
              <Keyboard className="size-4 text-primary shrink-0 mt-0.5" />
              <div>
                <div className="text-xs font-bold text-foreground">Global Hotkey</div>
                <div className="text-[11px] text-muted-foreground mt-0.5">
                  Press <kbd className="rounded bg-muted px-1.5 py-0.5 font-mono text-[10px] font-bold border border-border">K</kbd> anytime
                </div>
              </div>
            </div>

            <div className="flex items-start gap-2.5 rounded-lg border border-border/60 bg-secondary/30 p-2.5">
              <Monitor className="size-4 text-primary shrink-0 mt-0.5" />
              <div>
                <div className="text-xs font-bold text-foreground">Multi-Monitor Aware</div>
                <div className="text-[11px] text-muted-foreground mt-0.5">
                  Screenshots monitor under mouse cursor
                </div>
              </div>
            </div>

            <div className="flex items-start gap-2.5 rounded-lg border border-border/60 bg-secondary/30 p-2.5">
              <Undo2 className="size-4 text-amber-500 shrink-0 mt-0.5" />
              <div>
                <div className="text-xs font-bold text-foreground">One-Click Undo</div>
                <div className="text-[11px] text-muted-foreground mt-0.5">
                  Instantly remove mistakenly captured steps
                </div>
              </div>
            </div>

            <div className="flex items-start gap-2.5 rounded-lg border border-border/60 bg-secondary/30 p-2.5">
              <FileCheck2 className="size-4 text-emerald-500 shrink-0 mt-0.5" />
              <div>
                <div className="text-xs font-bold text-foreground">Auto Repagination</div>
                <div className="text-[11px] text-muted-foreground mt-0.5">
                  Word COM updates page fields on finish
                </div>
              </div>
            </div>
          </div>
        </WjCardBody>
      </WjCard>

      {/* Collapsible Procedure Guide */}
      <ProcedurePanel
        title="Workflow Capture Procedure & Features"
        badge="4 Steps"
        icon={<Layers className="size-4" />}
      >
        <ProcedureStep
          n="1"
          text={
            <>
              <strong className="text-foreground">Start Session:</strong> Click{" "}
              <span className="font-semibold text-primary">Launch Screenshot Tool</span> to open the floating toolbar.
              Click <strong className="text-foreground">Start</strong> on the toolbar to fill in the cover-page form
              (Client Name, Team Member, Date, Purpose of Review) and select where to save the{" "}
              <code className="text-[11px] bg-muted px-1 py-0.5 rounded border border-border">.docx</code> report.
            </>
          }
        />
        <ProcedureStep
          n="2"
          text={
            <>
              <strong className="text-foreground">Capture Screens:</strong> Walk through your application workflow.
              Press <kbd className="rounded bg-muted px-1.5 py-0.5 font-mono text-[10px] font-bold border border-border">K</kbd> or click{" "}
              <strong className="text-foreground">Capture</strong> whenever you hit a screen worth documenting. The tool
              detects which monitor your mouse is on and captures only that monitor.
            </>
          }
        />
        <ProcedureStep
          n="3"
          text={
            <>
              <strong className="text-foreground">Document Issues & Notes:</strong> Add optional{" "}
              <span className="text-rose-500 font-semibold">Issue</span> text (highlighted in red with live page numbers
              and logged in the front Issues Summary) or <span className="text-emerald-500 font-semibold">Note</span> text
              (dark green context) above any screenshot.
            </>
          }
        />
        <ProcedureStep
          n="4"
          text={
            <>
              <strong className="text-foreground">Undo or Finalize:</strong> Use <strong className="text-foreground">Undo</strong> to remove
              the last capture if taken by mistake. When testing is complete, click <strong className="text-foreground">End</strong> to
              enter pending/resolved counts. The tool writes the complete front page, updates Word page fields via COM automation, and saves the final report.
            </>
          }
        />
      </ProcedurePanel>

      {/* Live Automation Log Console */}
      <WjCard>
        <WjCardHeader
          title="Live Activity & Execution Log"
          hint="Real-time log console for active scripts and tools"
          actions={
            <span className="inline-flex items-center gap-2 rounded-full border border-border bg-secondary px-3 py-1 text-xs font-bold text-foreground">
              <CirclePlay className="size-3.5 text-primary" />
              Activity Monitor
            </span>
          }
        />
        <WjCardBody>
          <div className="mt-1">
            <WjProgress value={runConsole.progress} />
          </div>

          <div className="mt-3 flex flex-wrap items-center gap-3">
            <StatusText
              status={runConsole.status === "Ready" ? "Ready — launch Screenshot tool or start an automation script." : runConsole.status}
              kind={runConsole.statusKind}
            />
            <WjButton variant="ghost" size="sm" className="ml-auto" onClick={runConsole.clear}>
              <Eraser className="size-4" />
              Clear
            </WjButton>
          </div>

          <div className="mt-4">
            <LogConsole
              lines={runConsole.lines}
              className="h-[22rem]"
              emptyText="No activity yet — launch the Screenshot tool or run an automation script to stream execution logs here."
            />
          </div>
        </WjCardBody>
      </WjCard>
    </div>
  );
}
