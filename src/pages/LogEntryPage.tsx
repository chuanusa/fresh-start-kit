import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { ClipboardList } from 'lucide-react';

export function LogEntryPage() {
  return (
    <div className="space-y-8 max-w-4xl">
      <div>
        <h2 className="text-2xl font-bold tracking-tight gradient-text">填寫日誌</h2>
        <p className="text-muted-foreground text-sm mt-1">填寫每日工程施工日誌</p>
      </div>

      <Card className="glass-card hover-3d">
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <div className="p-2 rounded-xl gradient-primary">
              <ClipboardList className="h-4 w-4 text-white" />
            </div>
            日誌表單
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="text-center py-16">
            <ClipboardList className="h-12 w-12 mx-auto mb-3 text-muted-foreground/20" />
            <p className="text-muted-foreground">
              日誌填寫表單 — 待後端 API 串接完成後啟用
            </p>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
