import { createFileRoute } from "@tanstack/react-router";
import { ClipboardCheck, Eraser, Loader2, Play, Sparkles } from "lucide-react";
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
  WjLabel,
} from "@/components/wj/primitives";
import { DesktopNotice, FilePickerRow, LogConsole, StatusText } from "@/components/wj/run-console";
import { callApi, fileName, pickExcelFile, useDesktopApi, useRunConsole } from "@/lib/wj-bridge";

export const Route = createFileRoute("/enrollment-report-audit")({
  head: () => ({
    meta: [
      { title: "Enrollment Report Audit — Workforce Junction" },
      {
        name: "description",
        content:
          "Audit active enrollment plan columns across Pre and Post enrollment reports and export a highlighted workbook.",
      },
      { property: "og:title", content: "Enrollment Report Audit — Workforce Junction" },
      {
        property: "og:description",
        content: "Compare Pre and Post enrollment reports for active employees.",
      },
    ],
  }),
  component: EnrollmentReportAudit,
});

function EnrollmentReportAudit() {
  const { ready, checked } = useDesktopApi();
  const runConsole = useRunConsole(["wjEnrollmentFinished"]);
  const [preFile, setPreFile] = useState("");
  const [postFile, setPostFile] = useState("");

  async function browse(setter: (v: string) => void, label: string) {
    const path = await pickExcelFile();
    if (path) {
      setter(path);
      runConsole.setStatus(`${label} loaded: ` + fileName(path));
    }
  }

  async function run() {
    if (runConsole.running) return;
    if (!preFile || !postFile) { toast.error("Please select both the Pre and Post files."); return; }
    runConsole.setRunning(true);
    const ok = await callApi("run_enrollment_audit", preFile, postFile);
    if (!ok) {
      runConsole.setRunning(false);
      toast.error("Desktop backend not available.");
    }
  }

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Excel Automation"
        icon={<ClipboardCheck className="size-3.5" />}
        title="Enrollment Report Audit"
        subtitle="Filters active employees, matches them on SSN and compares every plan column between the Pre and Post enrollment reports. The audited workbook is saved to your Downloads folder."
      />

      <DesktopNotice visible={checked && !ready} />

      {/* Collapsible Procedure Accordion */}
      <ProcedurePanel
        title="Audit Rules & Check Details"
        badge="Audit Logic"
        icon={<Sparkles className="size-4" />}
      >
        <ProcedureStep
          n="1"
          text={
            <span>
              Only rows with an <strong className="text-foreground font-bold">Active</strong> status on both Pre and Post reports are evaluated.
            </span>
          }
        />
        <ProcedureStep
          n="2"
          text="Employees are matched and aligned on a normalised 9-digit Employee SSN."
        />
        <ProcedureStep
          n="3"
          text={
            <span>
              Plan columns ending in <strong className="text-foreground font-bold">Plan Name</strong> are audited, excluding core medical, dental, and vision.
            </span>
          }
        />
        <ProcedureStep
          n="4"
          text="Employer-paid and volume columns are picked automatically, skipping SP/CH variants."
        />
        <ProcedureStep
          n="5"
          text={
            <span>
              All detected discrepancies are highlighted in <strong className="text-warning font-bold">yellow</strong> in the exported audit workbook in Downloads.
            </span>
          }
          highlight
        />
      </ProcedurePanel>

      {/* Main Workspace Card */}
      <WjCard>
        <WjCardHeader title="Input Files & Run" hint="Header row 5 · .xlsx files" />
        <WjCardBody>
          <div className="grid gap-x-6 gap-y-4 md:grid-cols-2">
            <WjField>
              <WjLabel>Pre Enrollment Report</WjLabel>
              <FilePickerRow value={preFile} onBrowse={() => browse(setPreFile, "Pre file")} disabled={!ready} />
            </WjField>

            <WjField>
              <WjLabel>Post Enrollment Report</WjLabel>
              <FilePickerRow value={postFile} onBrowse={() => browse(setPostFile, "Post file")} disabled={!ready} />
            </WjField>
          </div>

          <div className="mt-6 flex flex-wrap items-center gap-3 pt-2">
            <WjButton onClick={run} disabled={runConsole.running || !ready}>
              {runConsole.running ? <Loader2 className="size-4 animate-spin" /> : <Play className="size-4" />}
              {runConsole.running ? "Auditing…" : "Run Audit"}
            </WjButton>
            <WjButton variant="ghost" onClick={runConsole.clear}>
              <Eraser className="size-4" />
              Clear Log
            </WjButton>
            <StatusText status={runConsole.status} kind={runConsole.statusKind} />
          </div>

          <div className="mt-5">
            <LogConsole lines={runConsole.lines} emptyText="Select both reports and run the audit to see live output." />
          </div>
        </WjCardBody>
      </WjCard>
    </div>
  );
}
