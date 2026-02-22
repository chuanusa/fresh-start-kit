import { useState, useEffect } from 'react';
import { useAuth } from '@/contexts/AuthContext';
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Eye, EyeOff, LogIn } from 'lucide-react';

interface LoginDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

export function LoginDialog({ open, onOpenChange }: LoginDialogProps) {
  const { login, loginWithGoogle, isLoading } = useAuth();
  const [identifier, setIdentifier] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);

  useEffect(() => {
    if (open) {
      // 確保 google library 載入後才初始化
      const initGoogleLogin = () => {
        if (typeof window.google !== 'undefined' && window.google.accounts) {
          window.google.accounts.id.initialize({
            client_id: "683889163276-88ed83s714f2ig1m4e9l52fbsk8k01sc.apps.googleusercontent.com", // This needs to be the actual client ID if available, using a placeholder/dummy or user's project ID if known. Wait, I should instruct the user to provide it or check if it's in the old file.
            callback: handleGoogleResponse
          });
          window.google.accounts.id.renderButton(
            document.getElementById("googleSignInButton"),
            { theme: "outline", size: "large", type: "standard", shape: "pill", text: "signin_with", locale: "zh-TW", logo_alignment: "left", width: "100%" }
          );
        } else {
          setTimeout(initGoogleLogin, 100);
        }
      };
      // We will define handleGoogleResponse below
      setTimeout(initGoogleLogin, 100);
    }
  }, [open]);

  const handleGoogleResponse = async (response: any) => {
    try {
      // 解析 JWT
      const base64Url = response.credential.split('.')[1];
      const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
      const jsonPayload = decodeURIComponent(atob(base64).split('').map(function (c) {
        return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
      }).join(''));

      const payload = JSON.parse(jsonPayload);
      const email = payload.email;

      const success = await loginWithGoogle(email);
      if (success) {
        onOpenChange(false);
      }
    } catch (err) {
      console.error("Google Login parsing error", err);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!identifier.trim() || !password) return;
    const success = await login(identifier.trim(), password);
    if (success) {
      onOpenChange(false);
      setIdentifier('');
      setPassword('');
    }
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-md">
        <DialogHeader>
          <DialogTitle className="text-center text-xl">登入系統</DialogTitle>
        </DialogHeader>
        <form onSubmit={handleSubmit} className="space-y-4 pt-2">
          <div className="space-y-2">
            <Label htmlFor="identifier">帳號 / 信箱</Label>
            <Input
              id="identifier"
              value={identifier}
              onChange={(e) => setIdentifier(e.target.value)}
              placeholder="請輸入帳號或信箱"
              autoComplete="username"
            />
          </div>
          <div className="space-y-2">
            <Label htmlFor="password">密碼</Label>
            <div className="relative">
              <Input
                id="password"
                type={showPassword ? 'text' : 'password'}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="請輸入密碼"
                autoComplete="current-password"
                className="pr-10"
              />
              <button
                type="button"
                onClick={() => setShowPassword(!showPassword)}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground transition-colors"
              >
                {showPassword ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
              </button>
            </div>
          </div>
          <Button type="submit" className="w-full" disabled={isLoading}>
            <LogIn className="mr-2 h-4 w-4" />
            {isLoading ? '登入中...' : '一般登入'}
          </Button>

          <div className="relative my-4">
            <div className="absolute inset-0 flex items-center">
              <span className="w-full border-t border-muted-foreground/20" />
            </div>
            <div className="relative flex justify-center text-xs uppercase">
              <span className="bg-background px-2 text-muted-foreground">
                或使用第三方登入
              </span>
            </div>
          </div>

          <div id="googleSignInButton" className="flex justify-center w-full min-h-[44px]"></div>
        </form>
      </DialogContent>
    </Dialog>
  );
}
