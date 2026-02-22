import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { ClipboardList } from 'lucide-react';

export function LogEntryPage() {
  return (
    <div className="space-y-6 max-w-4xl">
      <div>
        <h2 className="text-2xl font-bold tracking-tight">填寫日誌</h2>
        <p className="text-muted-foreground text-sm">填寫每日工程施工日誌</p>
      </div>

      <Card className="glass-card">
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <ClipboardList className="h-5 w-5 text-primary" />
            日誌表單
          </CardTitle>
        </CardHeader>
        <CardContent>
          <p className="text-muted-foreground text-center py-12">
            日誌填寫表單 — 待後端 API 串接完成後啟用
          </p>
        </CardContent>
      </Card>
    </div>
  );
}
