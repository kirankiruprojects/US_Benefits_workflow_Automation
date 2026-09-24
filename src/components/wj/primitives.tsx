import React, { useState } from "react";
import { cn } from "@/lib/utils";
import { ChevronDown, ChevronUp } from "lucide-react";

export interface WjButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: "default" | "ghost" | "secondary" | "destructive" | "outline";
  size?: "default" | "sm" | "lg" | "icon";
}

export const WjButton = React.forwardRef<HTMLButtonElement, WjButtonProps>(
  ({ className, variant = "default", size = "default", children, ...props }, ref) => {
    const baseStyles =
      "inline-flex items-center justify-center gap-2 rounded-lg font-bold transition-all duration-200 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring disabled:pointer-events-none disabled:opacity-50 active:scale-[0.98] cursor-pointer select-none";

    const variantStyles = {
      default:
        "gradient-teal text-primary-foreground shadow-[var(--shadow-card)] hover:brightness-110 hover:shadow-[var(--shadow-lift)]",
      ghost:
        "bg-transparent text-foreground hover:bg-secondary hover:text-foreground",
      secondary:
        "bg-secondary text-secondary-foreground hover:bg-secondary/80 border border-border/50",
      destructive:
        "bg-destructive text-destructive-foreground hover:bg-destructive/90 shadow-sm",
      outline:
        "border border-border bg-card text-foreground hover:bg-secondary/80",
    }[variant];

    const sizeStyles = {
      default: "h-10 px-4 py-2 text-sm",
      sm: "h-8 px-3 text-xs rounded-md",
      lg: "h-12 px-6 text-base",
      icon: "size-10",
    }[size];

    return (
      <button
        ref={ref}
        className={cn(baseStyles, variantStyles, sizeStyles, className)}
        {...props}
      >
        {children}
      </button>
    );
  }
);
WjButton.displayName = "WjButton";

export function WjCard({ children, className }: { children: React.ReactNode; className?: string }) {
  return (
    <div className={cn("surface-card overflow-hidden", className)}>
      {children}
    </div>
  );
}

export function WjCardHeader({
  title,
  hint,
  actions,
  className,
}: {
  title: string;
  hint?: string;
  actions?: React.ReactNode;
  className?: string;
}) {
  return (
    <div className={cn("flex flex-wrap items-center justify-between gap-3 border-b border-border/60 bg-secondary/30 px-6 py-4", className)}>
      <div>
        <h3 className="text-base font-extrabold text-foreground tracking-tight">{title}</h3>
        {hint && <p className="mt-0.5 text-xs text-muted-foreground font-medium">{hint}</p>}
      </div>
      {actions && <div className="flex items-center gap-2">{actions}</div>}
    </div>
  );
}

export function WjCardBody({ children, className }: { children: React.ReactNode; className?: string }) {
  return <div className={cn("p-6", className)}>{children}</div>;
}

export function WjField({ children, className }: { children: React.ReactNode; className?: string }) {
  return <div className={cn("space-y-1.5", className)}>{children}</div>;
}

export function WjLabel({
  children,
  hint,
  className,
}: {
  children: React.ReactNode;
  hint?: string;
  className?: string;
}) {
  return (
    <label className={cn("block text-xs font-bold text-foreground/90", className)}>
      {children}
      {hint && <span className="ml-1.5 font-normal text-muted-foreground">{hint}</span>}
    </label>
  );
}

export const WjInput = React.forwardRef<HTMLInputElement, React.InputHTMLAttributes<HTMLInputElement>>(
  ({ className, type = "text", ...props }, ref) => {
    return (
      <input
        type={type}
        ref={ref}
        className={cn(
          "flex h-10 w-full rounded-lg border border-input bg-card px-3.5 py-2 text-sm text-foreground placeholder:text-muted-foreground shadow-sm transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring disabled:cursor-not-allowed disabled:opacity-50",
          className
        )}
        {...props}
      />
    );
  }
);
WjInput.displayName = "WjInput";

export const WjTextarea = React.forwardRef<HTMLTextAreaElement, React.TextareaHTMLAttributes<HTMLTextAreaElement>>(
  ({ className, rows = 4, ...props }, ref) => {
    return (
      <textarea
        ref={ref}
        rows={rows}
        className={cn(
          "flex w-full rounded-lg border border-input bg-card px-3.5 py-2.5 text-sm font-mono text-foreground placeholder:text-muted-foreground shadow-sm transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring disabled:cursor-not-allowed disabled:opacity-50 resize-y",
          className
        )}
        {...props}
      />
    );
  }
);
WjTextarea.displayName = "WjTextarea";

export function ProcedurePanel({
  title,
  badge,
  icon,
  children,
}: {
  title: string;
  badge?: string;
  icon?: React.ReactNode;
  children: React.ReactNode;
}) {
  const [open, setOpen] = useState(true);

  return (
    <div className="surface-card overflow-hidden border-primary/20 bg-primary/[0.02]">
      <button
        type="button"
        onClick={() => setOpen(!open)}
        className="flex w-full items-center justify-between gap-3 px-5 py-3.5 text-left transition-colors hover:bg-secondary/40 cursor-pointer"
      >
        <div className="flex items-center gap-2.5">
          {icon && <span className="text-primary">{icon}</span>}
          <span className="text-xs font-bold uppercase tracking-wider text-foreground">{title}</span>
          {badge && (
            <span className="rounded-full bg-primary/10 px-2 py-0.5 text-[0.65rem] font-bold text-primary">
              {badge}
            </span>
          )}
        </div>
        {open ? <ChevronUp className="size-4 text-muted-foreground" /> : <ChevronDown className="size-4 text-muted-foreground" />}
      </button>
      {open && (
        <div className="border-t border-border/50 px-5 py-4 space-y-2.5 bg-card/60">
          {children}
        </div>
      )}
    </div>
  );
}

export function ProcedureStep({
  n,
  text,
  highlight = false,
}: {
  n: string;
  text: React.ReactNode;
  highlight?: boolean;
}) {
  return (
    <div className={cn("flex items-start gap-3 text-xs leading-relaxed", highlight ? "font-semibold text-foreground" : "text-muted-foreground")}>
      <span className={cn(
        "grid size-5 shrink-0 place-items-center rounded-full text-[0.65rem] font-extrabold",
        highlight ? "bg-primary text-primary-foreground" : "bg-secondary text-secondary-foreground border border-border/60"
      )}>
        {n}
      </span>
      <div className="pt-0.5">{text}</div>
    </div>
  );
}
