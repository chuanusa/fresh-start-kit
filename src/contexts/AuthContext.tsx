import React, { createContext, useContext, useState, useCallback } from 'react';
import { authenticateUser, logoutUser, type UserSession } from '@/lib/gas-api';
import { useToast } from '@/hooks/use-toast';

interface AuthContextType {
  user: UserSession | null;
  isGuest: boolean;
  isLoading: boolean;
  login: (identifier: string, password: string) => Promise<boolean>;
  logout: () => Promise<void>;
}

const AuthContext = createContext<AuthContextType | null>(null);

export function AuthProvider({ children }: { children: React.ReactNode }) {
  const [user, setUser] = useState<UserSession | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const { toast } = useToast();

  const isGuest = user === null;

  const login = useCallback(async (identifier: string, password: string) => {
    setIsLoading(true);
    try {
      const result = await authenticateUser(identifier, password);
      if (result.success && result.user) {
        setUser(result.user);
        toast({ title: '登入成功', description: result.message });
        return true;
      } else {
        toast({ title: '登入失敗', description: result.message, variant: 'destructive' });
        return false;
      }
    } catch (error: unknown) {
      toast({ title: '系統錯誤', description: (error as Error).message, variant: 'destructive' });
      return false;
    } finally {
      setIsLoading(false);
    }
  }, [toast]);

  const logout = useCallback(async () => {
    setIsLoading(true);
    try {
      await logoutUser();
      setUser(null);
      toast({ title: '已登出' });
    } catch {
      toast({ title: '登出失敗', variant: 'destructive' });
    } finally {
      setIsLoading(false);
    }
  }, [toast]);

  return (
    <AuthContext.Provider value={{ user, isGuest, isLoading, login, logout }}>
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error('useAuth must be used within AuthProvider');
  return ctx;
}
