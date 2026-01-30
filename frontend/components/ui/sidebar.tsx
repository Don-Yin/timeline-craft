'use client';
import * as React from "react";
import { cn } from "@/lib/utils";

type SidebarContextValue = {
  open: boolean;
  setOpen: (v: boolean | ((v: boolean) => boolean)) => void;
  toggleSidebar: () => void;
};

const SidebarContext = React.createContext<SidebarContextValue | null>(null);

export function SidebarProvider({
  children,
  defaultOpen = true,
  open: controlledOpen,
  onOpenChange,
}: {
  children: React.ReactNode;
  defaultOpen?: boolean;
  open?: boolean;
  onOpenChange?: (v: boolean) => void;
}) {
  const [uncontrolledOpen, _setOpen] = React.useState<boolean>(controlledOpen ?? defaultOpen);
  const isControlled = controlledOpen !== undefined;
  const open = isControlled ? controlledOpen! : uncontrolledOpen;

  const setOpen = React.useCallback(
    (value: boolean | ((v: boolean) => boolean)) => {
      const next = typeof value === "function" ? (value as (v: boolean) => boolean)(open) : value;
      if (onOpenChange) onOpenChange(next);
      if (!isControlled) _setOpen(next);
    },
    [isControlled, onOpenChange, open]
  );

  const toggleSidebar = React.useCallback(() => setOpen((v) => !v), [setOpen]);

  return (
    <SidebarContext.Provider value={React.useMemo(() => ({ open, setOpen, toggleSidebar }), [open, setOpen, toggleSidebar])}>
      {children}
    </SidebarContext.Provider>
  );
}

export function useSidebar() {
  const ctx = React.useContext(SidebarContext);
  if (!ctx) throw new Error("useSidebar must be used within a SidebarProvider.");
  return ctx;
}

export function Sidebar({ children, className, collapsible = "icon" }: { children: React.ReactNode; className?: string; collapsible?: "none" | "icon" }) {
  const { open } = useSidebar();
  return (
    <aside
      data-collapsible={collapsible}
      data-open={open ? "true" : "false"}
      className={cn(
        "group/sidebar relative border-r border-sidebar-border bg-sidebar text-sidebar-foreground",
        "flex flex-col transition-[width] duration-200 ease-in-out overflow-hidden",
        open ? "w-64" : "w-14",
        className
      )}
    >
      {children}
    </aside>
  );
}

export function SidebarHeader({ children, className }: { children?: React.ReactNode; className?: string }) {
  const { open } = useSidebar();
  return (
    <div className={cn("sticky top-0 z-10 border-b border-sidebar-border bg-sidebar transition-all duration-200", open ? "px-4 py-3" : "px-2 py-3", className)}>
      {children}
    </div>
  );
}

export function SidebarFooter({ children, className }: { children?: React.ReactNode; className?: string }) {
  const { open } = useSidebar();
  return (
    <div className={cn("mt-auto sticky bottom-0 z-10 border-t border-sidebar-border bg-sidebar transition-all duration-200", open ? "px-4 py-3" : "px-2 py-3 flex justify-center", className)}>
      {children}
    </div>
  );
}

export function SidebarContent({ children, className }: { children?: React.ReactNode; className?: string }) {
  const { open } = useSidebar();
  return <div className={cn("flex-1 overflow-y-auto py-3 transition-all duration-200", open ? "px-3" : "px-1", className)}>{children}</div>;
}

export function SidebarGroup({ children, className }: { children?: React.ReactNode; className?: string }) {
  return <div className={cn("mb-4", className)}>{children}</div>;
}

export function SidebarGroupLabel({ children, className }: { children?: React.ReactNode; className?: string }) {
  return <div className={cn("px-2 pb-2 text-xs font-semibold uppercase text-zinc-500", className)}>{children}</div>;
}

export function SidebarGroupContent({ children, className }: { children?: React.ReactNode; className?: string }) {
  return <div className={cn("space-y-1", className)}>{children}</div>;
}

export function SidebarMenu({ children, className }: { children?: React.ReactNode; className?: string }) {
  return <ul className={cn("flex flex-col gap-1", className)}>{children}</ul>;
}

export function SidebarMenuItem({ children, className }: { children?: React.ReactNode; className?: string }) {
  return <li className={cn(className)}>{children}</li>;
}

export function SidebarMenuButton({ children, asChild = false, className, isActive = false }: { children: React.ReactNode; asChild?: boolean; className?: string; isActive?: boolean }) {
  const { open } = useSidebar();
  const classes = cn(
    "flex w-full items-center rounded-md text-sm transition-all duration-200",
    "hover:bg-sidebar-accent hover:text-sidebar-accent-foreground",
    isActive && "bg-sidebar-accent text-sidebar-accent-foreground",
    open ? "gap-3 px-3 py-2 justify-start" : "justify-center px-0 py-2"
  );

  if (asChild) {
    const child = React.Children.only(children) as React.ReactElement<any>;
    return React.cloneElement(child, { className: cn(classes, (child.props as any)?.className, className) } as any);
  }

  return (
    <button className={cn(classes, className)} type="button" data-active={isActive ? "true" : "false"}>
      {children}
    </button>
  );
}

export function SidebarTrigger({ className, children }: { className?: string; children?: React.ReactNode }) {
  const { toggleSidebar } = useSidebar();
  return (
    <button
      onClick={toggleSidebar}
      className={cn(
        "inline-flex h-8 items-center rounded-md border border-sidebar-border bg-sidebar px-3 text-sm text-sidebar-foreground",
        "hover:bg-sidebar-accent hover:text-sidebar-accent-foreground transition-colors",
        className
      )}
      title="Toggle sidebar"
      type="button"
    >
      {children ?? "Toggle Sidebar"}
    </button>
  );
}

export function SidebarRail({ className }: { className?: string }) {
  const { toggleSidebar } = useSidebar();
  return (
    <div className={cn("pointer-events-none absolute inset-y-0 right-0 z-20 flex w-3 items-center justify-center", className)}>
      <button
        onClick={toggleSidebar}
        className="pointer-events-auto h-8 w-1.5 rounded-full bg-sidebar-border hover:bg-zinc-500 transition-colors"
        aria-label="Toggle sidebar"
        type="button"
      />
    </div>
  );
}
