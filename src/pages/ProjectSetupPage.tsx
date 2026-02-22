import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { FolderCog, Zap } from 'lucide-react';

export function ProjectSetupPage() {
  return (
    <div className="space-y-10">
      <div className="space-y-2">
        <div className="flex items-center gap-2">
          <Zap className="h-4 w-4 text-accent" />
          <span className="text-xs font-semibold uppercase tracking-widest text-accent">Projects</span>
        </div>
        <h2 className="text-3xl md:text-4xl font-bold tracking-tight gradient-text">工程管理</h2>
        <p className="text-muted-foreground text-sm">管理所有工程項目</p>
      </div>

      <Card className="glass-card hover-3d">
        <CardHeader>
          <CardTitle className="flex items-center gap-3 text-base">
            <div className="p-2.5 rounded-xl gradient-primary shadow-md shadow-primary/15">
              <FolderCog className="h-4 w-4 text-primary-foreground" />
            </div>
            <span className="font-semibold">工程列表</span>
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="text-center py-24">
            <FolderCog className="h-16 w-16 mx-auto mb-4 text-muted-foreground/10" />
            <p className="text-lg font-medium text-muted-foreground/60">工程管理</p>
            <p className="text-sm text-muted-foreground/40 mt-1">待後端 API 串接完成後啟用</p>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
