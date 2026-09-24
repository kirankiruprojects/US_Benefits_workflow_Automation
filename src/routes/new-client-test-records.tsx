import { createFileRoute } from "@tanstack/react-router";
import { Eraser, FileCheck2, Loader2, Play, Sparkles } from "lucide-react";
import { useState } from "react";
import { toast } from "sonner";
import { cn } from "@/lib/utils";
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

export const Route = createFileRoute("/new-client-test-records")({
  head: () => ({
    meta: [
      { title: "Test Records — Workforce Junction" },
      {
        name: "description",
        content:
          "Run password setup, new-client auto login and Open Enrollment login validation for new-client test records in a single workflow.",
      },
      { property: "og:title", content: "Test Records — Workforce Junction" },
      {
        property: "og:description",
        content: "Combined password setup and login validation workflow for new-client test records.",
      },
    ],
  }),
  component: TestRecords,
});

type TabId = "pw" | "login" | "tester";

const TABS: { id: TabId; label: string; desc: string }[] = [
  { id: "pw", label: "New Password Creation", desc: "Temp DOB password → permanent password & security questions" },
  { id: "login", label: "New Client Auto Login", desc: "Validates direct login for newly configured employee records" },
  { id: "tester", label: "OE Auto Login", desc: "Validates active Open Enrollment session popup notifications" },
];

const DEFAULT_ASSERTS = `Open Enrollment is Currently in Session!
The Open Enrollment period will end on May 31, 2026.
All changes will be effective 07/01/2026.`;

function TestRecords() {
  const { ready, checked } = useDesktopApi();
  const [tab, setTab] = useState<TabId>("pw");

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Portal Automation"
        icon={<FileCheck2 className="size-3.5" />}
        title="Test Records"
        subtitle="Runs both the password setup and login validation steps for new-client test records in a single desktop workflow."
      />

      <DesktopNotice visible={checked && !ready} />

      {/* Collapsible Procedure Accordion */}
      <ProcedurePanel
        title="Test Records Workflow & Multi-Stage Details"
        badge="3 Modules"
        icon={<Sparkles className="size-4" />}
      >
        <ProcedureStep
          n="1"
          text={
            <span>
              <strong className="text-foreground">New Password Creation:</strong> Uses UserID + Date of Birth as temporary credentials, sets new permanent passwords, and completes security questions.
            </span>
          }
          highlight={tab === "pw"}
        />
        <ProcedureStep
          n="2"
          text={
            <span>
              <strong className="text-foreground">New Client Auto Login:</strong> Tests authentication directly against the BenefitsJunction client portal to confirm account readiness.
            </span>
          }
          highlight={tab === "login"}
        />
        <ProcedureStep
          n="3"
          text={
            <span>
              <strong className="text-foreground">OE Auto Login:</strong> Verifies that Open Enrollment popup alerts and enrollment window dates render correctly on employee dashboard.
            </span>
          }
          highlight={tab === "tester"}
        />
      </ProcedurePanel>

      <WjCard>
        <WjCardHeader title="Workflow Execution" hint="Select a step to run" />
        <WjCardBody>
          <div className="mb-6 flex flex-wrap gap-2.5 p-1 rounded-xl bg-secondary/40 border border-border/50">
            {TABS.map((t) => (
              <button
                key={t.id}
                onClick={() => setTab(t.id)}
                className={cn(
                  "rounded-lg px-4 py-2.5 text-xs font-bold transition-all duration-200 cursor-pointer",
                  tab === t.id
                    ? "gradient-teal text-primary-foreground shadow-[var(--shadow-card)] scale-[1.02]"
                    : "text-muted-foreground hover:text-foreground hover:bg-card/60",
                )}
              >
                {t.label}
              </button>
            ))}
          </div>

          <div className="animate-slide-down">
            {tab === "pw" ? <PasswordPanel ready={ready} /> : null}
            {tab === "login" ? <LoginPanel ready={ready} /> : null}
            {tab === "tester" ? <TesterPanel ready={ready} /> : null}
          </div>
        </WjCardBody>
      </WjCard>
    </div>
  );
}

function RunRow({
  running,
  runLabel,
  onRun,
  onClear,
  status,
  statusKind,
  progress,
  disabled,
}: {
  running: boolean;
  runLabel: string;
  onRun: () => void;
  onClear: () => void;
  status: string;
  statusKind: "idle" | "pass" | "fail";
  progress: number;
  disabled: boolean;
}) {
  return (
    <>
      <div className="mt-5 flex flex-wrap items-center gap-3">
        <WjButton onClick={onRun} disabled={running || disabled}>
          {running ? <Loader2 className="size-4 animate-spin" /> : <Play className="size-4" />}
          {running ? "Running…" : runLabel}
        </WjButton>
        <WjButton variant="ghost" onClick={onClear}>
          <Eraser className="size-4" />
          Clear Log
        </WjButton>
        <StatusText status={status} kind={statusKind} />
      </div>
      <div className="my-4">
        <WjProgress value={progress} />
      </div>
    </>
  );
}

function PasswordPanel({ ready }: { ready: boolean }) {
  const console = useRunConsole(["wjPwSetupFinished"]);
  const [excelPath, setExcelPath] = useState("");
  const [portalUrl, setPortalUrl] = useState("https://www.benefitsjunction.com/");
  const [password, setPassword] = useState("");

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
    if (!password.trim()) { toast.error("Please enter a new password."); return; }
    console.setRunning(true);
    const ok = await callApi("run_new_client_password_setup", excelPath, portalUrl.trim(), password.trim());
    if (!ok) {
      console.setRunning(false);
      toast.error("Desktop backend not available.");
    }
  }

  return (
    <div>
      <WjField>
        <WjLabel>Portal URL</WjLabel>
        <WjInput value={portalUrl} onChange={(e) => setPortalUrl(e.target.value)} />
      </WjField>
      <WjField>
        <WjLabel hint="(UserID & Date Of Birth columns)">Employee Excel File</WjLabel>
        <FilePickerRow value={excelPath} onBrowse={browse} disabled={!ready} />
      </WjField>
      <WjField>
        <WjLabel>New Password (used for every record)</WjLabel>
        <WjInput
          type="password"
          placeholder="Enter the new password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
        />
      </WjField>
      <RunRow
        running={console.running}
        runLabel="Run Password Setup"
        onRun={run}
        onClear={console.clear}
        status={console.status}
        statusKind={console.statusKind}
        progress={console.progress}
        disabled={!ready}
      />
      <LogConsole lines={console.lines} />
    </div>
  );
}

function LoginPanel({ ready }: { ready: boolean }) {
  const console = useRunConsole(["wjTestFinished"]);
  const [excelPath, setExcelPath] = useState("");
  const [portalUrl, setPortalUrl] = useState("https://www.benefitsjunction.com/");

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
    console.setRunning(true);
    const ok = await callApi("run_new_client_auto_login", excelPath, portalUrl.trim());
    if (!ok) {
      console.setRunning(false);
      toast.error("Desktop backend not available.");
    }
  }

  return (
    <div>
      <WjField>
        <WjLabel>Portal URL</WjLabel>
        <WjInput value={portalUrl} onChange={(e) => setPortalUrl(e.target.value)} />
      </WjField>
      <WjField>
        <WjLabel hint="(Username & Password columns)">Test Records Excel File</WjLabel>
        <FilePickerRow value={excelPath} onBrowse={browse} disabled={!ready} />
      </WjField>
      <RunRow
        running={console.running}
        runLabel="Run Auto Login Check"
        onRun={run}
        onClear={console.clear}
        status={console.status}
        statusKind={console.statusKind}
        progress={console.progress}
        disabled={!ready}
      />
      <LogConsole lines={console.lines} />
    </div>
  );
}

function TesterPanel({ ready }: { ready: boolean }) {
  const console = useRunConsole(["wjTestFinished"]);
  const [excelPath, setExcelPath] = useState("");
  const [portalUrl, setPortalUrl] = useState("https://www.benefitsjunction.com/");
  const [asserts, setAsserts] = useState(DEFAULT_ASSERTS);

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
    <div>
      <WjField>
        <WjLabel>Portal URL</WjLabel>
        <WjInput value={portalUrl} onChange={(e) => setPortalUrl(e.target.value)} />
      </WjField>
      <WjField>
        <WjLabel hint="(Username & Password columns)">Test Records Excel File</WjLabel>
        <FilePickerRow value={excelPath} onBrowse={browse} disabled={!ready} />
      </WjField>
      <WjField>
        <WjLabel hint="(each line checked as a substring)">Expected Popup Text</WjLabel>
        <WjTextarea value={asserts} onChange={(e) => setAsserts(e.target.value)} />
      </WjField>
      <RunRow
        running={console.running}
        runLabel="Run Tests"
        onRun={run}
        onClear={console.clear}
        status={console.status}
        statusKind={console.statusKind}
        progress={console.progress}
        disabled={!ready}
      />
      <LogConsole lines={console.lines} />
    </div>
  );
}
