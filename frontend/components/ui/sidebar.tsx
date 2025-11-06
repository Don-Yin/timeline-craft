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
  const [uncontrolledOpen, _setOpen] = React.useState<boolean>(
    controlledOpen ?? defaultOpen
  );
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

  const contextValue = React.useMemo(
    () => ({ open, setOpen, toggleSidebar }),
    [open, setOpen, toggleSidebar]
  );

  return (
    <SidebarContext.Provider value={contextValue}>
      {children}
    </SidebarContext.Provider>
  );
}

export function useSidebar() {
  const ctx = React.useContext(SidebarContext);
  if (!ctx) {
    throw new Error("useSidebar must be used within a SidebarProvider.");
  }
  return ctx;
}

export function Sidebar({
  children,
  className,
  collapsible = "icon",
}: {
  children: React.ReactNode;
  className?: string;
  collapsible?: "none" | "icon";
}) {
  const { open } = useSidebar();
  return (
    <aside
      data-collapsible={collapsible}
      data-open={open ? "true" : "false"}
      className={cn(
        "group/sidebar relative border-r border-sidebar-border bg-sidebar text-sidebar-foreground",
        "flex flex-col transition-[width] duration-200 ease-in-out",
        open ? "w-64" : "w-16",
        className
      )}
    >
      {children}
    </aside>
  );
}

export function SidebarHeader({
  children,
  className,
}: {
  children?: React.ReactNode;
  className?: string;
}) {
  return (
    <div
      className={cn(
        "sticky top-0 z-10 border-b border-sidebar-border bg-sidebar px-3 py-3",
        className
      )}
    >
      {children}
    </div>
  );
}

export function SidebarFooter({
  children,
  className,
}: {
  children?: React.ReactNode;
  className?: string;
}) {
  return (
    <div
      className={cn(
        "mt-auto sticky bottom-0 z-10 border-t border-sidebar-border bg-sidebar px-3 py-3",
        className
      )}
    >
      {children}
    </div>
  );
}

export function SidebarContent({
  children,
  className,
}: {
  children?: React.ReactNode;
  className?: string;
}) {
  return (
    <div className={cn("flex-1 overflow-y-auto px-2 py-3", className)}>
      {children}
    </div>
  );
}

export function SidebarGroup({
  children,
  className,
}: {
  children?: React.ReactNode;
  className?: string;
}) {
  return <div className={cn("mb-4", className)}>{children}</div>;
}

export function SidebarGroupLabel({
  children,
  className,
}: {
  children?: React.ReactNode;
  className?: string;
}) {
  return (
    <div className={cn("px-2 pb-2 text-xs font-semibold uppercase text-zinc-500", className)}>
      {children}
    </div>
  );
}

export function SidebarGroupContent({
  children,
  className,
}: {
  children?: React.ReactNode;
  className?: string;
}) {
  return <div className={cn("space-y-1", className)}>{children}</div>;
}

export function SidebarMenu({
  children,
  className,
}: {
  children?: React.ReactNode;
  className?: string;
}) {
  return <ul className={cn("flex flex-col gap-1", className)}>{children}</ul>;
}

export function SidebarMenuItem({
  children,
  className,
}: {
  children?: React.ReactNode;
  className?: string;
}) {
  return <li className={cn(className)}>{children}</li>;
}

export function SidebarMenuButton({
  children,
  asChild = false,
  className,
  isActive = false,
}: {
  children: React.ReactNode;
  asChild?: boolean;
  className?: string;
  isActive?: boolean;
}) {
  const classes = cn(
    "flex w-full items-center gap-3 rounded-md px-2 py-2 text-sm",
    "hover:bg-sidebar-accent hover:text-sidebar-accent-foreground",
    "transition-colors",
    isActive && "bg-sidebar-accent text-sidebar-accent-foreground"
  );
  if (asChild) {
    // Clone the only child and merge classes like shadcn's Slot pattern
    const child = React.Children.only(children) as React.ReactElement<any>;
    return React.cloneElement(child, {
      className: cn(classes, (child.props as any)?.className, className),
    } as any);
  }
  return (
    <button
      className={cn(classes, className)}
      type="button"
      data-active={isActive ? "true" : "false"}
    >
      {children}
    </button>
  );
}

export function SidebarTrigger({
  className,
  children,
}: {
  className?: string;
  children?: React.ReactNode;
}) {
  const { toggleSidebar } = useSidebar();
  return (
    <button
      onClick={toggleSidebar}
      className={cn(
        "inline-flex h-8 items-center rounded-md border border-sidebar-border bg-sidebar px-3 text-sm text-sidebar-foreground",
        "hover:bg-sidebar-accent hover:text-sidebar-accent-foreground",
        "transition-colors",
        className
      )}
      title="Toggle sidebar"
      type="button"
    >
      {children ?? "Toggle Sidebar"}
    </button>
  );
}

export function SidebarRail({
  className,
}: {
  className?: string;
}) {
  const { toggleSidebar } = useSidebar();
  return (
    <>
      <div
        className={cn(
          "pointer-events-none absolute inset-y-0 right-0 z-20 flex w-3 items-center justify-center",
          className
        )}
      >
        <button
          onClick={toggleSidebar}
          className={cn(
            "pointer-events-auto h-8 w-2 rounded-md",
            "bg-sidebar-border hover:bg-sidebar-accent transition-colors"
          )}
          aria-label="Toggle sidebar"
          type="button"
        />
      </div>
    </>
  );
}


