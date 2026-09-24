import { createFileRoute } from "@tanstack/react-router";
import { Eraser, Loader2, Play, Sparkles, Wallet } from "lucide-react";
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

export const Route = createFileRoute("/salary-file-automation")({
  head: () => ({
    meta: [
      { title: "Salary File Automation — Workforce Junction" },
      {
        name: "description",
        content:
          "Compare a salary file against the BenJ work file, filter Active + FALSE rows and fill the upload template automatically.",
      },
      { property: "og:title", content: "Salary File Automation — Workforce Junction" },
      {
        property: "og:description",
        content: "Salary comparison and upload template filling in one run.",
      },
    ],
  }),
  component: SalaryFileAutomation,
});

function SalaryFileAutomation() {
  const { ready, checked } = useDesktopApi();
  const console = useRunConsole(["wjSalaryFinished"]);
  const [salaryFile, setSalaryFile] = useState("");
  const [workFile, setWorkFile] = useState("");
  const [templateFile, setTemplateFile] = useState("");

  async function browse(setter: (v: string) => void, label: string) {
    const path = await pickExcelFile();
    if (path) {
      setter(path);
      console.setStatus(`${label} loaded: ` + fileName(path));
    }
  }

  async function run() {
    if (console.running) return;
    if (!salaryFile || !workFile || !templateFile) { toast.error("Please select all three files."); return; }
    console.setRunning(true);
    const ok = await callApi("run_salary_automation", salaryFile, workFile, templateFile);
    if (!ok) {
      console.setRunning(false);
      toast.error("Desktop backend not available.");
    }
  }

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Excel Automation"
        icon={<Wallet className="size-3.5" />}
        title="Salary File Automation"
        subtitle="Salary comparison → filter (Active + FALSE) → fill upload template. Both comparison and upload template filling run seamlessly in a single step."
      />

      <DesktopNotice visible={checked && !ready} />

      {/* Collapsible Procedure Accordion */}
      <ProcedurePanel
        title="Salary Automation Process & Output Details"
        badge="4 Steps"
        icon={<Sparkles className="size-4" />}
      >
        <ProcedureStep
          n="1"
          text="The client Salary file and the BenJ work file are matched on Employee ID (with SSN fallback)."
        />
        <ProcedureStep
          n="2"
          text="Salary values are cleaned, normalised, and compared across both files to generate a comparison workbook."
        />
        <ProcedureStep
          n="3"
          text={
            <span>
              Rows that have an <strong className="text-foreground">Active</strong> status with a <strong className="text-destructive">FALSE</strong> comparison match are automatically filtered for upload.
            </span>
          }
          highlight
        />
        <ProcedureStep
          n="4"
          text="The upload template is populated and saved with yellow highlighted cell differences ready for direct portal import."
        />
      </ProcedurePanel>

      {/* Main Workspace Card */}
      <WjCard>
        <WjCardHeader title="Input Files & Run" hint="Requires Salary, Work, and Template files" />
        <WjCardBody>
          <div className="mb-4">
            <p className="mb-3 text-[0.68rem] font-extrabold uppercase tracking-[0.14em] text-primary">
              Step 1 — Salary Comparison Files
            </p>
            <div className="grid gap-x-6 gap-y-4 md:grid-cols-2">
              <WjField>
                <WjLabel>Salary File</WjLabel>
                <FilePickerRow value={salaryFile} onBrowse={() => browse(setSalaryFile, "Salary file")} disabled={!ready} />
              </WjField>

              <WjField>
                <WjLabel>Work File from BenJ</WjLabel>
                <FilePickerRow value={workFile} onBrowse={() => browse(setWorkFile, "Work file")} disabled={!ready} />
              </WjField>
            </div>
          </div>

          <div className="pt-2 border-t border-border">
            <p className="mb-3 mt-3 text-[0.68rem] font-extrabold uppercase tracking-[0.14em] text-primary">
              Step 2 — Upload Template Destination
            </p>
            <WjField>
              <WjLabel>Upload Template File</WjLabel>
              <FilePickerRow
                value={templateFile}
                onBrowse={() => browse(setTemplateFile, "Template")}
                disabled={!ready}
              />
            </WjField>
          </div>

          <div className="mt-6 flex flex-wrap items-center gap-3 pt-2">
            <WjButton onClick={run} disabled={console.running || !ready}>
              {console.running ? <Loader2 className="size-4 animate-spin" /> : <Play className="size-4" />}
              {console.running ? "Processing…" : "Generate & Fill Template"}
            </WjButton>
            <WjButton variant="ghost" onClick={console.clear}>
              <Eraser className="size-4" />
              Clear Log
            </WjButton>
            <StatusText status={console.status} kind={console.statusKind} />
          </div>

          <div className="mt-5">
            <LogConsole lines={console.lines} emptyText="Select all three files and run to stream progress here." />
          </div>
        </WjCardBody>
      </WjCard>
    </div>
  );
}
