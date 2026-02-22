import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { UserCog } from 'lucide-react';

export function InspectorManagementPage() {
  return (
    <div className="space-y-6 max-w-5xl">
      <div>
        <h2 className="text-2xl font-bold tracking-tight">檢驗員管理</h2>
        <p className="text-muted-foreground text-sm">管理檢驗員資料與狀態</p>
      </div>

      <Card className="glass-card">
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <UserCog className="h-5 w-5 text-primary" />
            檢驗員列表
          </CardTitle>
        </CardHeader>
        <CardContent>
          <p className="text-muted-foreground text-center py-12">
            檢驗員管理 — 待後端 API 串接完成後啟用
          </p>
        </CardContent>
      </Card>
    </div>
  );
}
