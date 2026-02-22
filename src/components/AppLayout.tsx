import { useState } from 'react';
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

const allTabs: { name: TabName; label: string; icon: React.ReactNode; roles?: string[] }[] = [
  { name: 'summaryReport', label: '總表', icon: <FileText className="h-5 w-5" /> },
  { name: 'dashboard', label: '儀表板', icon: <LayoutDashboard className="h-5 w-5" />, roles: ['all'] },
  { name: 'logEntry', label: '填寫日誌', icon: <ClipboardList className="h-5 w-5" />, roles: ['all'] },
  { name: 'logStatus', label: '填表狀態', icon: <CalendarCheck className="h-5 w-5" />, roles: ['管理員', '主管'] },
  { name: 'projectSetup', label: '工程管理', icon: <FolderCog className="h-5 w-5" />, roles: ['all'] },
  { name: 'inspectorManagement', label: '檢驗員管理', icon: <UserCog className="h-5 w-5" />, roles: ['管理員', '主管'] },
  { name: 'userManagement', label: '使用者管理', icon: <Users className="h-5 w-5" />, roles: ['管理員'] },
];

export function AppLayout({ activeTab, onTabChange, children }: AppLayoutProps) {
  const { user, isGuest, logout } = useAuth();
  const [loginOpen, setLoginOpen] = useState(false);
  const [darkMode, setDarkMode] = useState(false);

  const toggleDarkMode = () => {
    setDarkMode(!darkMode);
    document.documentElement.classList.toggle('dark');
  };

  const visibleTabs = allTabs.filter((tab) => {
    if (isGuest) return tab.name === 'summaryReport';
    if (!tab.roles) return true;
    if (tab.roles.includes('all')) return true;
    return user?.role && tab.roles.includes(user.role);
  });

  return (
    <div className="flex h-screen overflow-hidden">
      {/* Sidebar */}
      <aside className="hidden md:flex w-64 flex-col bg-sidebar text-sidebar-foreground border-r border-sidebar-border">
        {/* Logo */}
        <div className="flex items-center gap-3 px-5 py-5 border-b border-sidebar-border">
          <div className="flex items-center justify-center w-9 h-9 rounded-lg bg-sidebar-primary">
            <HardHat className="h-5 w-5 text-sidebar-primary-foreground" />
          </div>
          <div>
            <h1 className="text-sm font-bold tracking-tight">每日工程日誌</h1>
            <p className="text-xs text-sidebar-foreground/60">Engineering Log System</p>
          </div>
        </div>

        {/* Navigation */}
        <nav className="flex-1 px-3 py-4 space-y-1 overflow-y-auto">
          {visibleTabs.map((tab) => (
            <button
              key={tab.name}
              onClick={() => onTabChange(tab.name)}
              className={cn(
                'w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-all',
                activeTab === tab.name
                  ? 'bg-sidebar-primary text-sidebar-primary-foreground shadow-md'
                  : 'text-sidebar-foreground/70 hover:bg-sidebar-accent hover:text-sidebar-accent-foreground'
              )}
            >
              {tab.icon}
              {tab.label}
            </button>
          ))}
        </nav>

        {/* User section */}
        <div className="px-3 py-4 border-t border-sidebar-border">
          {isGuest ? (
            <Button
              variant="outline"
              className="w-full border-sidebar-border text-sidebar-foreground hover:bg-sidebar-accent"
              onClick={() => setLoginOpen(true)}
            >
              <LogIn className="mr-2 h-4 w-4" />
              登入
            </Button>
          ) : (
            <DropdownMenu>
              <DropdownMenuTrigger asChild>
                <button className="w-full flex items-center gap-3 px-3 py-2 rounded-lg hover:bg-sidebar-accent transition-colors">
                  <Avatar className="h-8 w-8">
                    <AvatarFallback className="bg-sidebar-primary text-sidebar-primary-foreground text-xs">
                      {user?.name?.charAt(0) || 'U'}
                    </AvatarFallback>
                  </Avatar>
                  <div className="flex-1 text-left">
                    <p className="text-sm font-medium">{user?.name}</p>
                    <p className="text-xs text-sidebar-foreground/60">{user?.role}</p>
                  </div>
                </button>
              </DropdownMenuTrigger>
              <DropdownMenuContent align="end" className="w-48">
                <DropdownMenuItem onClick={toggleDarkMode}>
                  {darkMode ? <Sun className="mr-2 h-4 w-4" /> : <Moon className="mr-2 h-4 w-4" />}
                  {darkMode ? '淺色模式' : '深色模式'}
                </DropdownMenuItem>
                <DropdownMenuSeparator />
                <DropdownMenuItem onClick={logout} className="text-destructive">
                  <LogOut className="mr-2 h-4 w-4" />
                  登出
                </DropdownMenuItem>
              </DropdownMenuContent>
            </DropdownMenu>
          )}
        </div>
      </aside>

      {/* Mobile header */}
      <div className="flex flex-col flex-1 overflow-hidden">
        <header className="md:hidden flex items-center justify-between px-4 py-3 bg-card border-b border-border">
          <div className="flex items-center gap-2">
            <HardHat className="h-5 w-5 text-primary" />
            <span className="font-bold text-sm">工程日誌</span>
          </div>
          <div className="flex items-center gap-2">
            <Button variant="ghost" size="icon" onClick={toggleDarkMode}>
              {darkMode ? <Sun className="h-4 w-4" /> : <Moon className="h-4 w-4" />}
            </Button>
            {isGuest ? (
              <Button variant="outline" size="sm" onClick={() => setLoginOpen(true)}>
                <LogIn className="mr-1 h-3 w-3" /> 登入
              </Button>
            ) : (
              <Button variant="ghost" size="sm" onClick={logout}>
                <LogOut className="h-4 w-4" />
              </Button>
            )}
          </div>
        </header>

        {/* Mobile tabs */}
        <div className="md:hidden overflow-x-auto border-b border-border bg-card">
          <div className="flex">
            {visibleTabs.map((tab) => (
              <button
                key={tab.name}
                onClick={() => onTabChange(tab.name)}
                className={cn(
                  'flex items-center gap-1.5 px-4 py-2.5 text-xs font-medium whitespace-nowrap border-b-2 transition-colors',
                  activeTab === tab.name
                    ? 'border-primary text-primary'
                    : 'border-transparent text-muted-foreground hover:text-foreground'
                )}
              >
                {tab.icon}
                {tab.label}
              </button>
            ))}
          </div>
        </div>

        {/* Main content */}
        <main className="flex-1 overflow-y-auto p-4 md:p-6 lg:p-8">
          {children}
        </main>
      </div>

      <LoginDialog open={loginOpen} onOpenChange={setLoginOpen} />
    </div>
  );
}
