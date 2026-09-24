import React from "react";
import { Link, useRouterState } from "@tanstack/react-router";
import {
  Bot,
  Calculator,
  Camera,
  CheckCircle2,
  ClipboardCheck,
  ClipboardList,
  FileCheck2,
  GitCompareArrows,
  LayoutGrid,
  Sparkles,
  Wallet,
} from "lucide-react";
import { cn } from "@/lib/utils";
import { useDesktopApi } from "@/lib/wj-bridge";

export interface NavItem {
  to: string;
  label: string;
  group: string;
  desc: string;
  icon: React.ReactNode;
  needsDesktop?: boolean;
}

export const NAV_ITEMS: NavItem[] = [
  {
    to: "/",
    label: "Overview",
    group: "Overview",
    desc: "Main automation console and system health dashboard.",
    icon: <LayoutGrid className="size-4" />,
    needsDesktop: false,
  },
  {
    to: "/salary-file-automation",
    label: "Salary File Automation",
    group: "Excel Automations",
    desc: "Compare salary file vs BenJ work file, filter Active+FALSE and fill upload template.",
    icon: <Wallet className="size-4" />,
    needsDesktop: true,
  },
  {
    to: "/comparison-tool",
    label: "Pre vs Post Comparison",
    group: "Excel Automations",
    desc: "Two-way Excel comparison of plan name, coverage, effective date and status.",
    icon: <GitCompareArrows className="size-4" />,
    needsDesktop: true,
  },
  {
    to: "/enrollment-report-audit",
    label: "Enrollment Report Audit",
    group: "Excel Automations",
    desc: "Audit active enrollment plan columns across Pre and Post reports.",
    icon: <ClipboardCheck className="size-4" />,
    needsDesktop: true,
  },
  {
    to: "/new-client-test-records",
    label: "Test Records Workflow",
    group: "Portal Automations",
    desc: "All-in-one password setup, new client login, and OE popup testing.",
    icon: <FileCheck2 className="size-4" />,
    needsDesktop: true,
  },
  {
    to: "/custom-report-generator",
    label: "Custom Report Generator",
    group: "Portal Automations",
    desc: "Automated Demographics report generation and download via Selenium.",
    icon: <Bot className="size-4" />,
    needsDesktop: true,
  },
  {
    to: "/rates-calculator",
    label: "Rates Calculator",
    group: "Calculators & Tools",
    desc: "Interactive insurance rate calculator with age banding & Excel export.",
    icon: <Calculator className="size-4" />,
    needsDesktop: false,
  },
  {
    to: "/clarification-form",
    label: "Clarification Form",
    group: "Calculators & Tools",
    desc: "Submit client questions and clarifications directly from the application.",
    icon: <ClipboardList className="size-4" />,
    needsDesktop: false,
  },
  {
    to: "/screenshot",
    label: "Screenshot",
    group: "Monitoring",
    desc: "Live stream of status, progress bar, and execution logs for active scripts.",
    icon: <Camera className="size-4" />,
    needsDesktop: false,
  },
];

export function PageHeader({
  eyebrow,
  icon,
  title,
  subtitle,
  actions,
}: {
  eyebrow?: string;
  icon?: React.ReactNode;
  title: string;
  subtitle?: string;
  actions?: React.ReactNode;
}) {
  return (
    <div className="flex flex-col gap-4 border-b border-border/70 pb-6 sm:flex-row sm:items-center sm:justify-between">
      <div>
        {eyebrow && (
          <div className="mb-1.5 flex items-center gap-2 text-[0.68rem] font-extrabold uppercase tracking-[0.16em] text-primary">
            {icon && <span className="size-3.5">{icon}</span>}
            <span>{eyebrow}</span>
          </div>
        )}
        <h1 className="text-2xl font-extrabold tracking-tight text-foreground sm:text-3xl">
          {title}
        </h1>
        {subtitle && (
          <p className="mt-1.5 max-w-3xl text-xs text-muted-foreground leading-relaxed">
            {subtitle}
          </p>
        )}
      </div>
      {actions && <div className="flex shrink-0 items-center gap-2.5">{actions}</div>}
    </div>
  );
}

export function AppShell({ children }: { children: React.ReactNode }) {
  const { ready, checked } = useDesktopApi();
  const routerState = useRouterState();
  const pathname = routerState.location.pathname;

  const groups = [...new Set(NAV_ITEMS.map((item) => item.group))];

  return (
    <div className="flex min-h-screen bg-background text-foreground">
      {/* Sidebar */}
      <aside className="fixed inset-y-0 left-0 z-30 flex w-72 flex-col border-r border-sidebar-border bg-sidebar text-sidebar-foreground">
        {/* Brand Header */}
        <div className="flex h-16 shrink-0 items-center gap-3 border-b border-sidebar-border px-5">
          <div className="grid size-9 place-items-center rounded-xl gradient-teal text-primary-foreground shadow-md">
            <Sparkles className="size-5" />
          </div>
          <div className="min-w-0">
            <p className="text-sm font-extrabold tracking-tight text-sidebar-foreground truncate">
              Workforce Junction
            </p>
            <p className="text-[0.68rem] font-semibold text-sidebar-muted">
              Automation Suite v2.0
            </p>
          </div>
        </div>

        {/* Navigation List */}
        <div className="log-scroll flex-1 overflow-y-auto px-3 py-4 space-y-5">
          {groups.map((group) => {
            const items = NAV_ITEMS.filter((i) => i.group === group);
            return (
              <div key={group} className="space-y-1">
                <p className="px-3 pb-1 text-[0.65rem] font-extrabold uppercase tracking-[0.16em] text-sidebar-muted/70">
                  {group}
                </p>
                {items.map((item) => {
                  const isActive =
                    item.to === "/"
                      ? pathname === "/"
                      : pathname.startsWith(item.to);

                  return (
                    <Link
                      key={item.to}
                      to={item.to}
                      className={cn(
                        "group flex items-center gap-3 rounded-lg px-3 py-2 text-xs font-semibold transition-all duration-150",
                        isActive
                          ? "bg-sidebar-accent text-sidebar-accent-foreground font-bold shadow-sm"
                          : "text-sidebar-muted hover:bg-white/[0.04] hover:text-sidebar-foreground"
                      )}
                    >
                      <span
                        className={cn(
                          "transition-colors",
                          isActive
                            ? "text-sidebar-primary"
                            : "text-sidebar-muted group-hover:text-sidebar-foreground"
                        )}
                      >
                        {item.icon}
                      </span>
                      <span className="truncate">{item.label}</span>
                    </Link>
                  );
                })}
              </div>
            );
          })}
        </div>

        {/* Backend Status Footer */}
        <div className="border-t border-sidebar-border p-3.5 bg-black/10">
          <div className="flex items-center justify-between rounded-lg bg-sidebar-accent/50 p-2.5">
            <div className="flex items-center gap-2">
              <span
                className={cn(
                  "size-2 rounded-full",
                  ready
                    ? "bg-success animate-pulse"
                    : checked
                    ? "bg-warning"
                    : "bg-muted-foreground"
                )}
              />
              <span className="text-[0.7rem] font-bold text-sidebar-foreground">
                {ready ? "Desktop Bridge Active" : checked ? "Web Preview Mode" : "Connecting..."}
              </span>
            </div>
            {ready && (
              <CheckCircle2 className="size-3.5 text-success" />
            )}
          </div>
        </div>
      </aside>

      {/* Main Content Area */}
      <main className="ml-72 flex-1 min-h-screen">
        <div className="mx-auto max-w-6xl p-8">{children}</div>
      </main>
    </div>
  );
}
