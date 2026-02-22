import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { FolderCog } from 'lucide-react';

export function ProjectSetupPage() {
  return (
    <div className="space-y-8 max-w-5xl">
      <div>
        <h2 className="text-2xl font-bold tracking-tight gradient-text">工程管理</h2>
        <p className="text-muted-foreground text-sm mt-1">管理所有工程項目</p>
      </div>

      <Card className="glass-card hover-3d">
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <div className="p-2 rounded-xl gradient-primary">
              <FolderCog className="h-4 w-4 text-white" />
            </div>
            工程列表
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="text-center py-16">
            <FolderCog className="h-12 w-12 mx-auto mb-3 text-muted-foreground/20" />
            <p className="text-muted-foreground">
              工程管理 — 待後端 API 串接完成後啟用
            </p>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
