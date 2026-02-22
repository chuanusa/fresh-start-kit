import { useState } from "react";
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

function AppContent() {
  const [activeTab, setActiveTab] = useState<TabName>("summaryReport");

  const renderPage = () => {
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
  };

  return (
    <AppLayout activeTab={activeTab} onTabChange={setActiveTab}>
      {renderPage()}
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
