import { useState, useRef, useEffect, useCallback } from "react";
import { AuthProvider } from "@/contexts/AuthContext";
import { AppLayout, type TabName } from "@/components/AppLayout";
import { SummaryReportPage } from "@/pages/SummaryReportPage";
import { DashboardPage } from "@/pages/DashboardPage";
import { LogEntryPage } from "@/pages/LogEntryPage";
import { LogStatusPage } from "@/pages/LogStatusPage";
import { ProjectSetupPage } from "@/pages/ProjectSetupPage";
import { InspectorManagementPage } from "@/pages/InspectorManagementPage";
import { UserManagementPage } from "@/pages/UserManagementPage";
import { Toaster } from "@/components/ui/toaster";

const tabOrder: TabName[] = [
  "summaryReport", "dashboard", "logEntry", "logStatus",
  "projectSetup", "inspectorManagement", "userManagement",
];

function PageTransition({ activeTab, children }: { activeTab: TabName; children: React.ReactNode }) {
  const [displayTab, setDisplayTab] = useState(activeTab);
  const [animClass, setAnimClass] = useState("translate-x-0 opacity-100");
  const prevTabRef = useRef(activeTab);

  useEffect(() => {
    if (activeTab === prevTabRef.current) return;

    const prevIdx = tabOrder.indexOf(prevTabRef.current);
    const nextIdx = tabOrder.indexOf(activeTab);
    const goRight = nextIdx > prevIdx;

    // Exit animation
    setAnimClass(goRight
      ? "-translate-x-8 opacity-0"
      : "translate-x-8 opacity-0"
    );

    const t = setTimeout(() => {
      setDisplayTab(activeTab);
      // Start off-screen on the other side
      setAnimClass(goRight
        ? "translate-x-8 opacity-0"
        : "-translate-x-8 opacity-0"
      );
      // Enter animation (next frame)
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          setAnimClass("translate-x-0 opacity-100");
        });
      });
    }, 150);

    prevTabRef.current = activeTab;
    return () => clearTimeout(t);
  }, [activeTab]);

  return (
    <div className={`transition-all duration-300 ease-out ${animClass}`} key={displayTab}>
      {children}
    </div>
  );
}

function AppContent() {
  const [activeTab, setActiveTab] = useState<TabName>("summaryReport");

  const renderPage = useCallback(() => {
    switch (activeTab) {
      case "summaryReport": return <SummaryReportPage />;
      case "dashboard": return <DashboardPage />;
      case "logEntry": return <LogEntryPage />;
      case "logStatus": return <LogStatusPage />;
      case "projectSetup": return <ProjectSetupPage />;
      case "inspectorManagement": return <InspectorManagementPage />;
      case "userManagement": return <UserManagementPage />;
      default: return <SummaryReportPage />;
    }
  }, [activeTab]);

  return (
    <AppLayout activeTab={activeTab} onTabChange={setActiveTab}>
      <PageTransition activeTab={activeTab}>
        {renderPage()}
      </PageTransition>
    </AppLayout>
  );
}

const App = () => (
  <AuthProvider>
    <AppContent />
    <Toaster />
  </AuthProvider>
);

export default App;
