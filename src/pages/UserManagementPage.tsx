import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Users } from 'lucide-react';

export function UserManagementPage() {
  return (
    <div className="space-y-6 max-w-5xl">
      <div>
        <h2 className="text-2xl font-bold tracking-tight">使用者管理</h2>
        <p className="text-muted-foreground text-sm">管理系統使用者帳號</p>
      </div>

      <Card className="glass-card">
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <Users className="h-5 w-5 text-primary" />
            使用者列表
          </CardTitle>
        </CardHeader>
        <CardContent>
          <p className="text-muted-foreground text-center py-12">
            使用者管理 — 待後端 API 串接完成後啟用
          </p>
        </CardContent>
      </Card>
    </div>
  );
}
