import { useState, useEffect } from 'react';
import { useAuth } from '@/contexts/AuthContext';
import { LoginDialog } from '@/components/LoginDialog';
import { Button } from '@/components/ui/button';
import { Avatar, AvatarFallback } from '@/components/ui/avatar';
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu';
import {
  ClipboardList,
  LayoutDashboard,
  FileText,
  CalendarCheck,
  FolderCog,
  Users,
  UserCog,
  LogIn,
  LogOut,
  Moon,
  Sun,
  HardHat,
  PanelLeftClose,
  PanelLeftOpen,
} from 'lucide-react';
import { cn } from '@/lib/utils';

export type TabName =
  | 'summaryReport'
  | 'dashboard'
  | 'logEntry'
  | 'logStatus'
  | 'projectSetup'
  | 'inspectorManagement'
  | 'userManagement';

interface AppLayoutProps {
  activeTab: TabName;
  onTabChange: (tab: TabName) => void;
  children: React.ReactNode;
}

const allTabs: { name: TabName; label: string; icon: React.ElementType; roles?: string[] }[] = [
  { name: 'summaryReport', label: '總表', icon: FileText },
  { name: 'dashboard', label: '儀表板', icon: LayoutDashboard, roles: ['all'] },
  { name: 'logEntry', label: '填寫日誌', icon: ClipboardList, roles: ['all'] },
  { name: 'logStatus', label: '填表狀態', icon: CalendarCheck, roles: ['管理員', '主管'] },
  { name: 'projectSetup', label: '工程管理', icon: FolderCog, roles: ['all'] },
  { name: 'inspectorManagement', label: '檢驗員管理', icon: UserCog, roles: ['管理員', '主管'] },
  { name: 'userManagement', label: '使用者管理', icon: Users, roles: ['管理員'] },
];

export function AppLayout({ activeTab, onTabChange, children }: AppLayoutProps) {
  const { user, isGuest, logout } = useAuth();
  const [loginOpen, setLoginOpen] = useState(false);
  const [collapsed, setCollapsed] = useState(false);
  const [lightMode, setLightMode] = useState(false);

  useEffect(() => {
    const saved = localStorage.getItem('theme');
    if (saved === 'light') {
      setLightMode(true);
      document.documentElement.classList.add('light');
    }
  }, []);

  const toggleTheme = () => {
    const next = !lightMode;
    setLightMode(next);
    if (next) {
      document.documentElement.classList.add('light');
      localStorage.setItem('theme', 'light');
    } else {
      document.documentElement.classList.remove('light');
      localStorage.setItem('theme', 'dark');
    }
  };

  const visibleTabs = allTabs.filter((tab) => {
    if (isGuest) return tab.name === 'summaryReport';
    if (!tab.roles) return true;
    if (tab.roles.includes('all')) return true;
    return user?.role && tab.roles.includes(user.role);
  });

  return (
    <div className="flex h-screen overflow-hidden bg-background">
      {/* ===== Desktop Sidebar ===== */}
      <aside
        className={cn(
          'hidden md:flex flex-col border-r border-sidebar-border transition-all duration-300 ease-in-out',
          'bg-sidebar-background/80 backdrop-blur-xl',
          collapsed ? 'w-[68px]' : 'w-64'
        )}
      >
        {/* Logo */}
        <div className="flex items-center gap-3 px-4 py-4 border-b border-sidebar-border min-h-[64px]">
          <div className="flex items-center justify-center w-9 h-9 rounded-xl gradient-primary shrink-0">
            <HardHat className="h-5 w-5 text-white" />
          </div>
          {!collapsed && (
            <div className="overflow-hidden animate-fade-up">
              <h1 className="text-sm font-bold tracking-tight text-sidebar-foreground">每日工程日誌</h1>
              <p className="text-[10px] text-sidebar-foreground/50">Engineering Log System</p>
            </div>
          )}
        </div>

        {/* Collapse toggle */}
        <button
          onClick={() => setCollapsed(!collapsed)}
          className="mx-3 mt-3 mb-1 flex items-center justify-center h-8 rounded-lg text-sidebar-foreground/50 hover:text-sidebar-foreground hover:bg-sidebar-accent transition-colors"
        >
          {collapsed ? <PanelLeftOpen className="h-4 w-4" /> : <PanelLeftClose className="h-4 w-4" />}
        </button>

        {/* Navigation */}
        <nav className="flex-1 px-3 py-2 space-y-1 overflow-y-auto">
          {visibleTabs.map((tab) => {
            const Icon = tab.icon;
            const isActive = activeTab === tab.name;
            return (
              <div key={tab.name} className="relative group">
                <button
                  onClick={() => onTabChange(tab.name)}
                  className={cn(
                    'w-full flex items-center gap-3 rounded-xl text-sm font-medium transition-all duration-200 relative',
                    collapsed ? 'justify-center px-2 py-2.5' : 'px-3 py-2.5',
                    isActive
                      ? 'gradient-primary text-white shadow-lg shadow-primary/20 active-indicator'
                      : 'text-sidebar-foreground/60 hover:bg-sidebar-accent hover:text-sidebar-foreground'
                  )}
                >
                  <Icon className="h-[18px] w-[18px] shrink-0" />
                  {!collapsed && <span>{tab.label}</span>}
                </button>
                {collapsed && (
                  <span className="pointer-events-none absolute left-full top-1/2 -translate-y-1/2 ml-2 px-2.5 py-1 rounded-lg bg-popover text-popover-foreground text-xs font-medium shadow-lg border border-border/50 whitespace-nowrap opacity-0 scale-95 group-hover:opacity-100 group-hover:scale-100 transition-all duration-150 z-50">
                    {tab.label}
                  </span>
                )}
              </div>
            );
          })}
        </nav>

        {/* Theme Toggle */}
        <div className="px-3 py-2 relative group">
          <button
            onClick={toggleTheme}
            className={cn(
              'w-full flex items-center gap-3 rounded-xl text-sm transition-colors hover:bg-sidebar-accent text-sidebar-foreground/60 hover:text-sidebar-foreground',
              collapsed ? 'justify-center px-2 py-2.5' : 'px-3 py-2.5'
            )}
          >
            {lightMode ? <Moon className="h-[18px] w-[18px]" /> : <Sun className="h-[18px] w-[18px]" />}
            {!collapsed && <span>{lightMode ? '深色模式' : '淺色模式'}</span>}
          </button>
          {collapsed && (
            <span className="pointer-events-none absolute left-full top-1/2 -translate-y-1/2 ml-2 px-2.5 py-1 rounded-lg bg-popover text-popover-foreground text-xs font-medium shadow-lg border border-border/50 whitespace-nowrap opacity-0 scale-95 group-hover:opacity-100 group-hover:scale-100 transition-all duration-150 z-50">
              {lightMode ? '深色模式' : '淺色模式'}
            </span>
          )}
        </div>

        {/* User section */}
        <div className="px-3 py-3 border-t border-sidebar-border">
          {isGuest ? (
            <Button
              variant="outline"
              className={cn(
                'w-full border-sidebar-border text-sidebar-foreground hover:bg-sidebar-accent',
                collapsed && 'px-2'
              )}
              onClick={() => setLoginOpen(true)}
            >
              <LogIn className="h-4 w-4 shrink-0" />
              {!collapsed && <span className="ml-2">登入</span>}
            </Button>
          ) : (
            <DropdownMenu>
              <DropdownMenuTrigger asChild>
                <button className={cn(
                  'w-full flex items-center gap-3 rounded-xl hover:bg-sidebar-accent transition-colors',
                  collapsed ? 'justify-center px-2 py-2' : 'px-3 py-2'
                )}>
                  <Avatar className="h-8 w-8 shrink-0">
                    <AvatarFallback className="gradient-primary text-white text-xs font-bold">
                      {user?.name?.charAt(0) || 'U'}
                    </AvatarFallback>
                  </Avatar>
                  {!collapsed && (
                    <div className="flex-1 text-left overflow-hidden">
                      <p className="text-sm font-medium text-sidebar-foreground truncate">{user?.name}</p>
                      <p className="text-[10px] text-sidebar-foreground/50">{user?.role}</p>
                    </div>
                  )}
                </button>
              </DropdownMenuTrigger>
              <DropdownMenuContent align="end" className="w-48">
                <DropdownMenuItem onClick={logout} className="text-destructive">
                  <LogOut className="mr-2 h-4 w-4" />
                  登出
                </DropdownMenuItem>
              </DropdownMenuContent>
            </DropdownMenu>
          )}
        </div>
      </aside>

      {/* ===== Main Area ===== */}
      <div className="flex flex-col flex-1 overflow-hidden">
        {/* Mobile header */}
        <header className="md:hidden flex items-center justify-between px-4 py-3 bg-card/80 backdrop-blur-xl border-b border-border">
          <div className="flex items-center gap-2.5">
            <div className="flex items-center justify-center w-8 h-8 rounded-lg gradient-primary">
              <HardHat className="h-4 w-4 text-white" />
            </div>
            <span className="font-bold text-sm">工程日誌</span>
          </div>
          <div className="flex items-center gap-1.5">
            <Button variant="ghost" size="icon" onClick={toggleTheme} className="h-8 w-8">
              {lightMode ? <Moon className="h-4 w-4" /> : <Sun className="h-4 w-4" />}
            </Button>
            {isGuest ? (
              <Button variant="outline" size="sm" onClick={() => setLoginOpen(true)} className="h-8 text-xs">
                <LogIn className="mr-1 h-3 w-3" /> 登入
              </Button>
            ) : (
              <Button variant="ghost" size="icon" onClick={logout} className="h-8 w-8">
                <LogOut className="h-4 w-4" />
              </Button>
            )}
          </div>
        </header>

        {/* Main content */}
        <main className="flex-1 overflow-y-auto p-4 md:p-6 lg:p-8">
          <div className="animate-fade-up">
            {children}
          </div>
        </main>

        {/* ===== Mobile Bottom Navigation ===== */}
        <nav className="md:hidden flex items-center justify-around bg-card/80 backdrop-blur-xl border-t border-border px-1 py-1.5 safe-area-bottom">
          {visibleTabs.slice(0, 5).map((tab) => {
            const Icon = tab.icon;
            const isActive = activeTab === tab.name;
            return (
              <button
                key={tab.name}
                onClick={() => onTabChange(tab.name)}
                className={cn(
                  'flex flex-col items-center gap-0.5 px-2 py-1.5 rounded-xl transition-all min-w-0 flex-1',
                  isActive
                    ? 'text-primary'
                    : 'text-muted-foreground'
                )}
              >
                <div className={cn(
                  'p-1.5 rounded-xl transition-all',
                  isActive && 'gradient-primary shadow-lg shadow-primary/25'
                )}>
                  <Icon className={cn('h-4 w-4', isActive && 'text-white')} />
                </div>
                <span className="text-[10px] font-medium truncate max-w-full">{tab.label}</span>
              </button>
            );
          })}
        </nav>
      </div>

      <LoginDialog open={loginOpen} onOpenChange={setLoginOpen} />
    </div>
  );
}
