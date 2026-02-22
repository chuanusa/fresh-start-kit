import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { CalendarCheck } from 'lucide-react';

export function LogStatusPage() {
  return (
    <div className="space-y-8 max-w-5xl">
      <div>
        <h2 className="text-2xl font-bold tracking-tight gradient-text">填表狀態</h2>
        <p className="text-muted-foreground text-sm mt-1">檢視各工程填表進度</p>
      </div>

      <Card className="glass-card hover-3d">
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <div className="p-2 rounded-xl gradient-primary">
              <CalendarCheck className="h-4 w-4 text-white" />
            </div>
            月份填表狀態
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="text-center py-16">
            <CalendarCheck className="h-12 w-12 mx-auto mb-3 text-muted-foreground/20" />
            <p className="text-muted-foreground">
              填表狀態日曆 — 待後端 API 串接完成後啟用
            </p>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
