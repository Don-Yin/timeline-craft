'use client';
import { usePathname } from "next/navigation";
import { Home, Settings, FileText, Layers } from "lucide-react";
import {
  Sidebar,
  SidebarContent,
  SidebarFooter,
  SidebarGroup,
  SidebarGroupContent,
  SidebarHeader,
  SidebarRail,
  SidebarMenu,
  SidebarMenuButton,
  SidebarMenuItem,
  useSidebar,
} from "@/components/ui/sidebar";
import Link from "next/link";

export function AppSidebar() {
  const pathname = usePathname();
  const { open } = useSidebar();

  const items = [
    { title: "landing", url: "/landing", icon: Home },
    { title: "manage", url: "/manage", icon: FileText },
    { title: "settings", url: "/settings", icon: Settings },
  ];

  return (
    <Sidebar collapsible="icon">
      <SidebarHeader>
        <div className={`flex items-center overflow-hidden transition-all duration-200 ${open ? "gap-2 justify-start" : "justify-center"}`}>
          <Layers className="h-5 w-5 shrink-0 text-emerald-500" />
          <span className={`font-semibold whitespace-nowrap transition-all duration-200 ${open ? "opacity-100 w-auto" : "opacity-0 w-0"}`}>
            TimelineCraft
          </span>
        </div>
      </SidebarHeader>

      <SidebarContent>
        <SidebarGroup>
          <SidebarGroupContent>
            <SidebarMenu>
              {items.map((item) => (
                <SidebarMenuItem key={item.title}>
                  <SidebarMenuButton asChild isActive={pathname === item.url || pathname.startsWith(item.url + "/")}>
                    <Link href={item.url}>
                      <item.icon className="h-4 w-4 shrink-0" />
                      <span className={`transition-all duration-200 whitespace-nowrap ${open ? "opacity-100" : "opacity-0 w-0 overflow-hidden"}`}>
                        {item.title}
                      </span>
                    </Link>
                  </SidebarMenuButton>
                </SidebarMenuItem>
              ))}
            </SidebarMenu>
          </SidebarGroupContent>
        </SidebarGroup>
      </SidebarContent>

      <SidebarFooter>
        <span className={`text-xs text-zinc-500 transition-all duration-200 ${open ? "" : "text-center"}`}>
          {open ? "v0.1" : "v0"}
        </span>
      </SidebarFooter>

      <SidebarRail />
    </Sidebar>
  );
}
