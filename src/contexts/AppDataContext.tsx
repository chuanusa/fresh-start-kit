import React, { createContext, useContext, useState, useCallback } from 'react';
import { loadInitialData, type ProjectData, type InspectorData } from '@/lib/gas-api';
import { useToast } from '@/hooks/use-toast';

interface AppDataContextType {
  projects: ProjectData[];
  inspectors: InspectorData[];
  disasterTypes: string[];
  isLoaded: boolean;
  isLoading: boolean;
  refresh: () => Promise<void>;
}

const AppDataContext = createContext<AppDataContextType | null>(null);

export function AppDataProvider({ children }: { children: React.ReactNode }) {
  const [projects, setProjects] = useState<ProjectData[]>([]);
  const [inspectors, setInspectors] = useState<InspectorData[]>([]);
  const [disasterTypes, setDisasterTypes] = useState<string[]>([]);
  const [isLoaded, setIsLoaded] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const { toast } = useToast();

  const refresh = useCallback(async () => {
    setIsLoading(true);
    try {
      const data = await loadInitialData();
      setProjects(data.projects || []);
      setInspectors(data.inspectors || []);
      setDisasterTypes(data.disasterTypes || []);
      setIsLoaded(true);
    } catch {
      toast({ title: '載入初始資料失敗', variant: 'destructive' });
    } finally {
      setIsLoading(false);
    }
  }, [toast]);

  return (
    <AppDataContext.Provider value={{ projects, inspectors, disasterTypes, isLoaded, isLoading, refresh }}>
      {children}
    </AppDataContext.Provider>
  );
}

export function useAppData() {
  const ctx = useContext(AppDataContext);
  if (!ctx) throw new Error('useAppData must be used within AppDataProvider');
  return ctx;
}
