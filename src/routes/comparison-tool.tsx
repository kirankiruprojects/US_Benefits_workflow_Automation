import { createFileRoute } from "@tanstack/react-router";
import { Eraser, GitCompareArrows, Loader2, Play, Sparkles } from "lucide-react";
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

export const Route = createFileRoute("/comparison-tool")({
  head: () => ({
    meta: [
      { title: "Pre vs Post Comparison — Workforce Junction" },
      {
        name: "description",
        content:
          "Compare Pre and Post enrollment Excel files both ways and save highlighted comparison workbooks to Downloads.",
      },
      { property: "og:title", content: "Pre vs Post Comparison — Workforce Junction" },
      {
        property: "og:description",
        content: "Two-way Pre/Post Excel comparison with highlighted compare columns.",
      },
    ],
  }),
  component: ComparisonTool,
});

function ComparisonTool() {
  const { ready, checked } = useDesktopApi();
  const console = useRunConsole(["wjComparisonFinished"]);
  const [preFile, setPreFile] = useState("");
  const [postFile, setPostFile] = useState("");

  async function browse(setter: (v: string) => void, label: string) {
    const path = await pickExcelFile();
    if (path) {
      setter(path);
      console.setStatus(`${label} loaded: ` + fileName(path));
    }
  }

  async function run() {
    if (console.running) return;
    if (!preFile || !postFile) { toast.error("Please select both Pre and Post files."); return; }
    console.setRunning(true);
    const ok = await callApi("run_comparison_tool", preFile, postFile);
    if (!ok) {
      console.setRunning(false);
      toast.error("Desktop backend not available.");
    }
  }

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Excel Automation"
        icon={<GitCompareArrows className="size-3.5" />}
        title="Plan Enrollment Finalization: Pre vs Post Comparison"
        subtitle="Compares Plan Name, Coverage, Effective Date and Enrollment Status for matching SSNs in both directions, then writes two highlighted workbooks (PRE → POST and POST → PRE) into your Downloads folder."
      />

      <DesktopNotice visible={checked && !ready} />

      {/* Collapsible Procedure Accordion */}
      <ProcedurePanel
        title="Comparison Procedure & Output Details"
        badge="4 Steps"
        icon={<Sparkles className="size-4" />}
      >
        <ProcedureStep
          n="1"
          text="Active employee records are matched bi-directionally based on normalised Employee SSN."
        />
        <ProcedureStep
          n="2"
          text="Key enrollment columns (Plan Name, Coverage, Effective Date, Enrollment Status) receive Post/Pre comparisons."
        />
        <ProcedureStep
          n="3"
          text={
            <span>
              All generated <strong className="text-foreground">Compare</strong> headers and mismatched cells are highlighted in <strong className="text-warning">yellow</strong> for fast visual auditing.
            </span>
          }
          highlight
        />
        <ProcedureStep
          n="4"
          text={
            <span>
              Two formatted workbooks land directly in your Downloads folder: <code className="rounded bg-secondary px-1.5 py-0.5 text-xs font-mono text-primary">[Name]_PRE_TO_POST.xlsx</code> and <code className="rounded bg-secondary px-1.5 py-0.5 text-xs font-mono text-primary">[Name]_POST_TO_PRE.xlsx</code>.
            </span>
          }
        />
      </ProcedurePanel>

      {/* Main Workspace Card */}
      <WjCard>
        <WjCardHeader title="Input Files & Run" hint=".xlsx files supported" />
        <WjCardBody>
          <div className="grid gap-x-6 gap-y-4 md:grid-cols-2">
            <WjField>
              <WjLabel>Pre Excel File</WjLabel>
              <FilePickerRow value={preFile} onBrowse={() => browse(setPreFile, "Pre")} disabled={!ready} />
            </WjField>

            <WjField>
              <WjLabel>Post Excel File</WjLabel>
              <FilePickerRow value={postFile} onBrowse={() => browse(setPostFile, "Post")} disabled={!ready} />
            </WjField>
          </div>

          <div className="mt-6 flex flex-wrap items-center gap-3 pt-2">
            <WjButton onClick={run} disabled={console.running || !ready}>
              {console.running ? <Loader2 className="size-4 animate-spin" /> : <Play className="size-4" />}
              {console.running ? "Generating…" : "Generate Output"}
            </WjButton>
            <WjButton variant="ghost" onClick={console.clear}>
              <Eraser className="size-4" />
              Clear Log
            </WjButton>
            <StatusText status={console.status} kind={console.statusKind} />
          </div>

          <div className="mt-5">
            <LogConsole lines={console.lines} emptyText="Select both files and generate to see output paths and live logs here." />
          </div>
        </WjCardBody>
      </WjCard>
    </div>
  );
}
