import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Users } from 'lucide-react';

export function UserManagementPage() {
  return (
    <div className="space-y-8 max-w-5xl">
      <div>
        <h2 className="text-2xl font-bold tracking-tight gradient-text">使用者管理</h2>
        <p className="text-muted-foreground text-sm mt-1">管理系統使用者帳號</p>
      </div>

      <Card className="glass-card hover-3d">
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <div className="p-2 rounded-xl gradient-primary">
              <Users className="h-4 w-4 text-white" />
            </div>
            使用者列表
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="text-center py-16">
            <Users className="h-12 w-12 mx-auto mb-3 text-muted-foreground/20" />
            <p className="text-muted-foreground">
              使用者管理 — 待後端 API 串接完成後啟用
            </p>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
