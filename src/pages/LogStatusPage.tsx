import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { CalendarCheck } from 'lucide-react';

export function LogStatusPage() {
  return (
    <div className="space-y-6 max-w-5xl">
      <div>
        <h2 className="text-2xl font-bold tracking-tight">填表狀態</h2>
        <p className="text-muted-foreground text-sm">檢視各工程填表進度</p>
      </div>

      <Card className="glass-card">
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <CalendarCheck className="h-5 w-5 text-primary" />
            月份填表狀態
          </CardTitle>
        </CardHeader>
        <CardContent>
          <p className="text-muted-foreground text-center py-12">
            填表狀態日曆 — 待後端 API 串接完成後啟用
          </p>
        </CardContent>
      </Card>
    </div>
  );
}
