import { createFileRoute } from "@tanstack/react-router";
import { ClipboardList, ExternalLink } from "lucide-react";
import { PageHeader } from "@/components/wj/shell";
import { WjButton } from "@/components/wj/primitives";

const FORM_URL = "https://client-clarifications-form.web.app";

export const Route = createFileRoute("/clarification-form")({
  head: () => ({
    meta: [
      { title: "Clarification Form — Workforce Junction" },
      {
        name: "description",
        content:
          "Client clarification form for Workforce Junction — submit questions and clarifications directly from the desktop application.",
      },
      { property: "og:title", content: "Clarification Form — Workforce Junction" },
      {
        property: "og:description",
        content: "Submit client clarifications directly from the Workforce Junction desktop app.",
      },
    ],
  }),
  component: ClarificationForm,
});

function ClarificationForm() {
  return (
    <>
      <PageHeader
        eyebrow="Forms"
        icon={<ClipboardList className="size-3.5" />}
        title="Clarification Form"
        subtitle="Submit client clarifications and questions directly from the application. The form opens securely in-app."
        actions={
          <WjButton variant="ghost" onClick={() => window.open(FORM_URL, "_blank")}>
            <ExternalLink className="size-4" />
            Open in new tab
          </WjButton>
        }
      />

      <div className="surface-card overflow-hidden">
        <iframe
          src={FORM_URL}
          title="Client Clarification Form"
          className="h-[calc(100vh-15rem)] min-h-[640px] w-full border-0 bg-white"
          allow="forms"
        />
      </div>
    </>
  );
}
