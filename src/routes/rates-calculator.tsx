import { createFileRoute } from "@tanstack/react-router";
import { Calculator, ExternalLink } from "lucide-react";
import { PageHeader } from "@/components/wj/shell";
import { WjButton } from "@/components/wj/primitives";

export const Route = createFileRoute("/rates-calculator")({
  head: () => ({
    meta: [
      { title: "Rates Calculator — Workforce Junction" },
      {
        name: "description",
        content:
          "Employee insurance rate calculator with age-banded rates, tobacco loads and Excel export — runs entirely in the app.",
      },
      { property: "og:title", content: "Rates Calculator — Workforce Junction" },
      {
        property: "og:description",
        content: "Calculate employee insurance rates and export the results.",
      },
    ],
  }),
  component: RatesCalculator,
});

function RatesCalculator() {
  return (
    <>
      <PageHeader
        eyebrow="Calculators"
        icon={<Calculator className="size-3.5" />}
        title="Rates Calculator"
        subtitle="The full Employee Insurance Rate Calculator, embedded unchanged. All calculations run locally in your browser."
        actions={
          <WjButton variant="ghost" onClick={() => window.open("/tools/rates-calculator.html", "_blank")}>
            <ExternalLink className="size-4" />
            Open in new tab
          </WjButton>
        }
      />

      <div className="surface-card overflow-hidden">
        <iframe
          src="/tools/rates-calculator.html"
          title="Employee Insurance Rate Calculator"
          className="h-[calc(100vh-15rem)] min-h-[640px] w-full border-0 bg-white"
        />
      </div>
    </>
  );
}
