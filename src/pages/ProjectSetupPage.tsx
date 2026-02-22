import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { FolderCog } from 'lucide-react';

export function ProjectSetupPage() {
  return (
    <div className="space-y-6 max-w-5xl">
      <div>
        <h2 className="text-2xl font-bold tracking-tight">工程管理</h2>
        <p className="text-muted-foreground text-sm">管理所有工程項目</p>
      </div>

      <Card className="glass-card">
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <FolderCog className="h-5 w-5 text-primary" />
            工程列表
          </CardTitle>
        </CardHeader>
        <CardContent>
          <p className="text-muted-foreground text-center py-12">
            工程管理 — 待後端 API 串接完成後啟用
          </p>
        </CardContent>
      </Card>
    </div>
  );
}
