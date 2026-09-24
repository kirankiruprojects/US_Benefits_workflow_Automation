import { createFileRoute } from "@tanstack/react-router";
import { Eraser, KeyRound, Loader2, Play, Sparkles } from "lucide-react";
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
} from "@/components/wj/primitives";
import {
  DesktopNotice,
  FilePickerRow,
  LogConsole,
  StatusText,
  WjProgress,
} from "@/components/wj/run-console";
import { callApi, fileName, pickExcelFile, useDesktopApi, useRunConsole } from "@/lib/wj-bridge";

export const Route = createFileRoute("/new-client-password-setup")({
  head: () => ({
    meta: [
      { title: "New Client Password Setup — Workforce Junction" },
      {
        name: "description",
        content:
          "Sign in with the DOB-based temporary password, set a new password and answer the security questions for every employee record.",
      },
      { property: "og:title", content: "New Client Password Setup — Workforce Junction" },
      {
        property: "og:description",
        content: "Bulk password setup automation for new client employee records.",
      },
    ],
  }),
  component: PasswordSetup,
});

function PasswordSetup() {
  const { ready, checked } = useDesktopApi();
  const [excelPath, setExcelPath] = useState("");
  const [portalUrl, setPortalUrl] = useState("https://stg.benefitsjunction.com/");
  const [password, setPassword] = useState("");
  const console = useRunConsole(["wjPwSetupFinished"]);

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
    <div className="space-y-6">
      <PageHeader
        eyebrow="Portal Automation"
        icon={<KeyRound className="size-3.5" />}
        title="New Client Password Setup"
        subtitle="Loads each employee record from an Excel file, signs in with the DOB-based temporary password, sets a new password, answers the security questions, and confirms the flow in a real Chrome browser."
      />

      <DesktopNotice visible={checked && !ready} />

      {/* Collapsible Procedure Accordion */}
      <ProcedurePanel
        title="Password Setup Workflow & Steps"
        badge="4 Stages"
        icon={<Sparkles className="size-4" />}
      >
        <ProcedureStep
          n="1"
          text="Extracts UserID and Date Of Birth (DOB) columns from the employee Excel sheet."
        />
        <ProcedureStep
          n="2"
          text="Logs in with the temporary password formatted from employee Date of Birth (MMDDYYYY)."
        />
        <ProcedureStep
          n="3"
          text="Fills in the new permanent password for both password fields and submits password change."
        />
        <ProcedureStep
          n="4"
          text="Answers all required security questions automatically and confirms successful account activation."
          highlight
        />
      </ProcedurePanel>

      {/* Main Workspace Card */}
      <WjCard>
        <WjCardHeader title="Configuration & Inputs" hint="Excel requires 'UserID' & 'Date Of Birth' columns" />
        <WjCardBody>
          <div className="grid gap-x-6 gap-y-4 md:grid-cols-2">
            <WjField>
              <WjLabel>Portal URL</WjLabel>
              <WjInput value={portalUrl} onChange={(e) => setPortalUrl(e.target.value)} />
            </WjField>

            <WjField>
              <WjLabel hint="(UserID & Date Of Birth columns)">Employee Excel File</WjLabel>
              <FilePickerRow value={excelPath} onBrowse={browse} disabled={!ready} />
            </WjField>
          </div>

          <WjField>
            <WjLabel>New Password (applied to all records)</WjLabel>
            <WjInput
              type="password"
              placeholder="Enter the new password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
            />
          </WjField>

          <div className="mt-6 flex flex-wrap items-center gap-3 pt-2">
            <WjButton onClick={run} disabled={console.running || !ready}>
              {console.running ? <Loader2 className="size-4 animate-spin" /> : <Play className="size-4" />}
              {console.running ? "Running…" : "Run Password Setup"}
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
