import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { CalendarCheck, Zap } from 'lucide-react';

export function LogStatusPage() {
  return (
    <div className="space-y-10">
      <div className="space-y-2">
        <div className="flex items-center gap-2">
          <Zap className="h-4 w-4 text-accent" />
          <span className="text-xs font-semibold uppercase tracking-widest text-accent">Status</span>
        </div>
        <h2 className="text-3xl md:text-4xl font-bold tracking-tight gradient-text">填表狀態</h2>
        <p className="text-muted-foreground text-sm">檢視各工程填表進度</p>
      </div>

      <Card className="glass-card hover-3d">
        <CardHeader>
          <CardTitle className="flex items-center gap-3 text-base">
            <div className="p-2.5 rounded-xl gradient-primary shadow-md shadow-primary/15">
              <CalendarCheck className="h-4 w-4 text-primary-foreground" />
            </div>
            <span className="font-semibold">月份填表狀態</span>
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="text-center py-24">
            <CalendarCheck className="h-16 w-16 mx-auto mb-4 text-muted-foreground/10" />
            <p className="text-lg font-medium text-muted-foreground/60">
              填表狀態日曆
            </p>
            <p className="text-sm text-muted-foreground/40 mt-1">待後端 API 串接完成後啟用</p>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
