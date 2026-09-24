import { createFileRoute, Link } from "@tanstack/react-router";
import { ArrowRight, CheckCircle2, LayoutGrid, MonitorCheck, Sparkles, Zap } from "lucide-react";
import { NAV_ITEMS, PageHeader } from "@/components/wj/shell";
import { useDesktopApi } from "@/lib/wj-bridge";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Automation Dashboard — Workforce Junction" },
      {
        name: "description",
        content:
          "Launch every Workforce Junction automation — portal logins, report downloads, Excel comparisons and the rates calculator — from one dashboard.",
      },
      { property: "og:title", content: "Automation Dashboard — Workforce Junction" },
      {
        property: "og:description",
        content: "All portal, Excel and rate automations in a single application.",
      },
    ],
  }),
  component: Dashboard,
});

function Dashboard() {
  const { ready, checked } = useDesktopApi();
  const tools = NAV_ITEMS.filter((i) => i.to !== "/");
  const groups = [...new Set(tools.map((t) => t.group))];

  return (
    <div className="space-y-8 animate-slide-down">
      <PageHeader
        eyebrow="Overview"
        icon={<LayoutGrid className="size-3.5" />}
        title="Automation Suite"
        subtitle="All Workforce Junction automation workflows — salary file automation, Excel comparisons, enrollment audits, portal testing (including password setup and OE auto-login inside Test Records), custom report generation, and rates calculator — in one integrated console."
      />

      <div className="grid gap-4 sm:grid-cols-3">
        <StatCard
          label="Automations Available"
          value={String(tools.length)}
          subtitle="Ready to execute"
          icon={<Zap className="size-4 text-primary" />}
        />
        <StatCard
          label="Desktop-Backed Workflows"
          value={String(tools.filter((t) => t.needsDesktop).length)}
          subtitle="Selenium & Excel powered"
          icon={<Sparkles className="size-4 text-primary" />}
        />
        <StatCard
          label="Backend Status"
          value={ready ? "Connected" : checked ? "Browser Mode" : "Detecting…"}
          subtitle={ready ? "Python pywebview bridge active" : "Standalone web preview"}
          icon={<CheckCircle2 className="size-4" />}
          tone={ready ? "ok" : checked ? "warn" : "idle"}
        />
      </div>

      {groups.map((group) => (
        <section key={group} className="space-y-3">
          <div className="flex items-center gap-2">
            <h2 className="text-xs font-extrabold uppercase tracking-[0.16em] text-muted-foreground">
              {group}
            </h2>
            <span className="h-px flex-1 bg-border/60" />
          </div>
          <div className="grid gap-4 md:grid-cols-2">
            {tools
              .filter((t) => t.group === group)
              .map((tool) => (
                <Link
                  key={tool.to}
                  to={tool.to}
                  className="surface-card group relative flex items-start gap-4 p-5 transition-all duration-200 hover:-translate-y-1 hover:border-primary/50 hover:shadow-[var(--shadow-lift)] active:translate-y-0"
                >
                  <span className="grid size-11 shrink-0 place-items-center rounded-xl bg-accent text-accent-foreground transition-all duration-200 group-hover:gradient-teal group-hover:text-primary-foreground group-hover:shadow-[var(--shadow-card)]">
                    {tool.icon}
                  </span>
                  <span className="min-w-0 flex-1">
                    <span className="flex items-center gap-2">
                      <span className="text-sm font-extrabold text-foreground group-hover:text-primary transition-colors duration-150">
                        {tool.label}
                      </span>
                      {tool.needsDesktop ? (
                        <span className="rounded-full bg-secondary/80 px-2 py-0.5 text-[0.6rem] font-bold uppercase tracking-wider text-muted-foreground border border-border/40">
                          Desktop
                        </span>
                      ) : null}
                    </span>
                    <span className="mt-1.5 block text-xs leading-relaxed text-muted-foreground line-clamp-2">
                      {tool.desc}
                    </span>
                  </span>
                  <ArrowRight className="mt-1 size-4 shrink-0 text-muted-foreground transition-all duration-200 group-hover:translate-x-1 group-hover:text-primary" />
                </Link>
              ))}
          </div>
        </section>
      ))}
    </div>
  );
}

function StatCard({
  label,
  value,
  subtitle,
  icon,
  tone = "idle",
}: {
  label: string;
  value: string;
  subtitle?: string;
  icon: React.ReactNode;
  tone?: "idle" | "ok" | "warn";
}) {
  return (
    <div className="surface-card p-5 relative overflow-hidden transition-all duration-200 hover:shadow-[var(--shadow-lift)]">
      <div className="mb-2 flex items-center justify-between">
        <span className="text-[0.68rem] font-bold uppercase tracking-[0.14em] text-muted-foreground">
          {label}
        </span>
        <span className="grid size-8 place-items-center rounded-lg bg-secondary/80">
          {icon}
        </span>
      </div>
      <p
        className={
          tone === "ok"
            ? "text-2xl font-extrabold text-success"
            : tone === "warn"
              ? "text-2xl font-extrabold text-warning"
              : "text-2xl font-extrabold text-foreground"
        }
      >
        {value}
      </p>
      {subtitle ? (
        <p className="mt-1 text-[0.72rem] font-medium text-muted-foreground">
          {subtitle}
        </p>
      ) : null}
    </div>
  );
}
