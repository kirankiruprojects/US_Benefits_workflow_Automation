import { createFileRoute } from "@tanstack/react-router";
import { Eraser, Play, ShieldCheck, Loader2, Sparkles } from "lucide-react";
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
  WjField,
  WjInput,
  WjLabel,
  WjTextarea,
} from "@/components/wj/primitives";
import {
  DesktopNotice,
  FilePickerRow,
  LogConsole,
  StatusText,
  WjProgress,
} from "@/components/wj/run-console";
import { callApi, fileName, pickExcelFile, useDesktopApi, useRunConsole } from "@/lib/wj-bridge";

export const Route = createFileRoute("/auto-login-tester")({
  head: () => ({
    meta: [
      { title: "Auto-Login Tester — Workforce Junction" },
      {
        name: "description",
        content:
          "Log into the benefits portal with every row of a test-records Excel file and assert the Open Enrollment popup text.",
      },
      { property: "og:title", content: "Auto-Login Tester — Workforce Junction" },
      {
        property: "og:description",
        content: "Portal login automation driven from an Excel test-records file.",
      },
    ],
  }),
  component: AutoLoginTester,
});

const DEFAULT_ASSERTS = `Open Enrollment is Currently in Session!
The Open Enrollment period will end on May 31, 2026.
All changes will be effective 07/01/2026.`;

function AutoLoginTester() {
  const { ready, checked } = useDesktopApi();
  const [excelPath, setExcelPath] = useState("");
  const [portalUrl, setPortalUrl] = useState("https://www.benefitsjunction.com/");
  const [asserts, setAsserts] = useState(DEFAULT_ASSERTS);
  const console = useRunConsole(["wjTestFinished"]);

  async function browse() {
    const path = await pickExcelFile();
    if (path) {
      setExcelPath(path);
      console.setStatus("Loaded: " + fileName(path));
    }
  }

  async function run() {
    if (console.running) return;
    if (!excelPath) { toast.error("Please select the Excel file first."); return; }
    const lines = asserts
      .split("\n")
      .map((l) => l.trim())
      .filter(Boolean);
    if (lines.length === 0) { toast.error("Please enter expected popup text."); return; }

    console.setRunning(true);
    const ok = await callApi("run_auto_login_test", excelPath, portalUrl.trim(), lines.join("\n"));
    if (!ok) {
      console.setRunning(false);
      toast.error("Desktop backend not available.");
    }
  }

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Portal Automation"
        icon={<ShieldCheck className="size-3.5" />}
        title="Open Enrollment Auto-Login Tester"
        subtitle="Logs into the benefits portal with each row of an Excel test-records file and asserts the Open Enrollment popup text — drives a real Chrome browser on this machine."
      />

      <DesktopNotice visible={checked && !ready} />

      {/* Collapsible Procedure Accordion */}
      <ProcedurePanel
        title="Login Test Procedure & Assertion Details"
        badge="Portal Test"
        icon={<Sparkles className="size-4" />}
      >
        <ProcedureStep
          n="1"
          text="Reads the Excel file extracting 'Username' and 'Password' column values for each test employee."
        />
        <ProcedureStep
          n="2"
          text="Launches a dedicated browser session and navigates directly to the specified BenefitsJunction portal URL."
        />
        <ProcedureStep
          n="3"
          text="Enters credentials, submits login, and asserts the presence of popup message strings (PopUp_LblMsg)."
        />
        <ProcedureStep
          n="4"
          text="Logs PASS/FAIL results per employee in real-time, displaying a summary table at completion."
          highlight
        />
      </ProcedurePanel>

      {/* Main Workspace Card */}
      <WjCard>
        <WjCardHeader title="Configuration & Inputs" hint="Excel requires 'Username' & 'Password' columns" />
        <WjCardBody>
          <div className="grid gap-x-6 gap-y-4 md:grid-cols-2">
            <WjField>
              <WjLabel>Portal URL</WjLabel>
              <WjInput value={portalUrl} onChange={(e) => setPortalUrl(e.target.value)} />
            </WjField>

            <WjField>
              <WjLabel hint="(Username & Password columns)">Test Records Excel File</WjLabel>
              <FilePickerRow value={excelPath} onBrowse={browse} disabled={!ready} />
            </WjField>
          </div>

          <WjField>
            <WjLabel hint="(each line checked as a substring in popup text)">Expected Popup Text</WjLabel>
            <WjTextarea value={asserts} onChange={(e) => setAsserts(e.target.value)} />
          </WjField>

          <div className="mt-6 flex flex-wrap items-center gap-3 pt-2">
            <WjButton onClick={run} disabled={console.running || !ready}>
              {console.running ? <Loader2 className="size-4 animate-spin" /> : <Play className="size-4" />}
              {console.running ? "Running…" : "Run Tests"}
            </WjButton>
            <WjButton variant="ghost" onClick={console.clear}>
              <Eraser className="size-4" />
              Clear Log
            </WjButton>
            <StatusText status={console.status} kind={console.statusKind} />
          </div>

          <div className="my-4">
            <WjProgress value={console.progress} />
          </div>

          <LogConsole lines={console.lines} />
        </WjCardBody>
      </WjCard>
    </div>
  );
}
