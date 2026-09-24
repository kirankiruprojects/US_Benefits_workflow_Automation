import { createFileRoute } from "@tanstack/react-router";
import { Bot, Eraser, Loader2, Play, Sparkles } from "lucide-react";
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
import { DesktopNotice, FilePickerRow, LogConsole, StatusText } from "@/components/wj/run-console";
import { callApi, pickDriverFile, useDesktopApi, useRunConsole } from "@/lib/wj-bridge";

export const Route = createFileRoute("/custom-report-generator")({
  head: () => ({
    meta: [
      { title: "Custom Report Generator — Workforce Junction" },
      {
        name: "description",
        content:
          "Drive the BenefitsJunction admin portal end-to-end to build and download a Demographics report from configurable fields.",
      },
      { property: "og:title", content: "Custom Report Generator — Workforce Junction" },
      {
        property: "og:description",
        content: "Automated Demographics report building and download for the admin portal.",
      },
    ],
  }),
  component: CustomReportGenerator,
});

function CustomReportGenerator() {
  const { ready, checked } = useDesktopApi();
  const console = useRunConsole(["wjReportFinished"]);

  const [driverPath, setDriverPath] = useState(
    "C:\\Users\\kiran.bs\\Downloads\\edgedriver_win64\\msedgedriver.exe",
  );
  const [portalUrl, setPortalUrl] = useState("https://admin.benefitsjunction.com/");
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [clientName, setClientName] = useState("W.T.Rich");
  const [reportName, setReportName] = useState("WFJ Metal Pros LLC Demographics Report");
  const [reportDate, setReportDate] = useState("10/1/2025");

  async function browseDriver() {
    const path = await pickDriverFile();
    if (path) setDriverPath(path);
  }

  async function run() {
    if (console.running) return;
    const params = {
      driver_path: driverPath.trim(),
      portal_url: portalUrl.trim(),
      username,
      password,
      client_name: clientName.trim(),
      report_name: reportName.trim(),
      report_date: reportDate.trim(),
    };
    if (!params.driver_path || !params.username || !params.password) {
      toast.error("Please fill in the driver path, username, and password.");
      return;
    }
    console.setRunning(true);
    const ok = await callApi("run_custom_report", params);
    if (!ok) {
      console.setRunning(false);
      toast.error("Desktop backend not available.");
    }
  }

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Portal Automation"
        icon={<Bot className="size-3.5" />}
        title="Custom Report Generator"
        subtitle="Drives the BenefitsJunction admin portal end-to-end to build and download a Demographics report with complete field reordering and generation."
      />

      <DesktopNotice visible={checked && !ready} />

      {/* Collapsible Procedure Accordion */}
      <ProcedurePanel
        title="Report Generation Procedure & Workflow"
        badge="End-to-end"
        icon={<Sparkles className="size-4" />}
      >
        <ProcedureStep
          n="1"
          text="Launches Microsoft Edge WebDriver and navigates to the Admin Portal login page."
        />
        <ProcedureStep
          n="2"
          text="Authenticates with credentials, selects target Client Name from dropdown, and loads Custom Reports."
        />
        <ProcedureStep
          n="3"
          text="Navigates portal iframes, selects Demographics, customizes fields, and submits report generation with date parameters."
        />
        <ProcedureStep
          n="4"
          text="Monitors 'Previously Generated Reports' queue until status changes to Completed, then automatically downloads the Excel file."
          highlight
        />
      </ProcedurePanel>

      {/* Main Workspace Card */}
      <WjCard>
        <WjCardHeader title="Admin Portal & Report Parameters" hint="Configurable fields" />
        <WjCardBody>
          <div className="grid gap-x-6 gap-y-4 md:grid-cols-2">
            <WjField>
              <WjLabel>Edge WebDriver Path</WjLabel>
              <FilePickerRow value={driverPath} onBrowse={browseDriver} disabled={!ready} />
            </WjField>

            <WjField>
              <WjLabel>Portal URL</WjLabel>
              <WjInput value={portalUrl} onChange={(e) => setPortalUrl(e.target.value)} />
            </WjField>

            <WjField>
              <WjLabel>Username</WjLabel>
              <WjInput value={username} onChange={(e) => setUsername(e.target.value)} />
            </WjField>

            <WjField>
              <WjLabel>Password</WjLabel>
              <WjInput type="password" value={password} onChange={(e) => setPassword(e.target.value)} />
            </WjField>

            <WjField>
              <WjLabel hint="(must match dropdown exactly)">Client Name</WjLabel>
              <WjInput value={clientName} onChange={(e) => setClientName(e.target.value)} />
            </WjField>

            <WjField>
              <WjLabel>Report Name</WjLabel>
              <WjInput value={reportName} onChange={(e) => setReportName(e.target.value)} />
            </WjField>

            <WjField className="md:col-span-2">
              <WjLabel hint="(M/D/YYYY format)">Report Effective Date</WjLabel>
              <WjInput value={reportDate} onChange={(e) => setReportDate(e.target.value)} />
            </WjField>
          </div>

          <div className="mt-6 flex flex-wrap items-center gap-3 pt-2">
            <WjButton onClick={run} disabled={console.running || !ready}>
              {console.running ? <Loader2 className="size-4 animate-spin" /> : <Play className="size-4" />}
              {console.running ? "Running…" : "Run Report Generation"}
            </WjButton>
            <WjButton variant="ghost" onClick={console.clear}>
              <Eraser className="size-4" />
              Clear Log
            </WjButton>
            <StatusText status={console.status} kind={console.statusKind} />
          </div>

          <div className="mt-5">
            <LogConsole lines={console.lines} />
          </div>
        </WjCardBody>
      </WjCard>
    </div>
  );
}
