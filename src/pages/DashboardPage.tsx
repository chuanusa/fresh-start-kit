import { useState, useEffect } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Badge } from '@/components/ui/badge';
import { getMonthlyDashboardData, type DashboardData } from '@/lib/gas-api';
import { useToast } from '@/hooks/use-toast';
import { BarChart3, AlertTriangle, Loader2 } from 'lucide-react';

export function DashboardPage() {
  const now = new Date();
  const [year, setYear] = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth() + 1);
  const [data, setData] = useState<DashboardData | null>(null);
  const [loading, setLoading] = useState(false);
  const { toast } = useToast();

  const loadData = async () => {
    setLoading(true);
    try {
      const result = await getMonthlyDashboardData(year, month);
      setData(result);
    } catch (e) {
      toast({ title: '載入儀表板失敗', description: (e as Error).message, variant: 'destructive' });
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, [year, month]);

  const years = Array.from({ length: 3 }, (_, i) => now.getFullYear() - i);
  const months = Array.from({ length: 12 }, (_, i) => i + 1);

  const maxWorkDays = data?.workDays?.length ? Math.max(...data.workDays.map(d => d.days)) : 1;
  const maxHazards = data?.hazards?.length ? Math.max(...data.hazards.map(h => h.count)) : 1;

  return (
    <div className="space-y-6 max-w-5xl">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-bold tracking-tight">每月儀表板</h2>
          <p className="text-muted-foreground text-sm">統計分析與圖表</p>
        </div>
        <div className="flex gap-2">
          <Select value={String(year)} onValueChange={(v: string) => setYear(Number(v))}>
            <SelectTrigger className="w-24"><SelectValue /></SelectTrigger>
            <SelectContent>
              {years.map(y => <SelectItem key={y} value={String(y)}>{y}年</SelectItem>)}
            </SelectContent>
          </Select>
          <Select value={String(month)} onValueChange={(v: string) => setMonth(Number(v))}>
            <SelectTrigger className="w-20"><SelectValue /></SelectTrigger>
            <SelectContent>
              {months.map(m => <SelectItem key={m} value={String(m)}>{m}月</SelectItem>)}
            </SelectContent>
          </Select>
        </div>
      </div>

      {loading ? (
        <div className="flex items-center justify-center py-20">
          <Loader2 className="h-8 w-8 animate-spin text-primary" />
        </div>
      ) : data ? (
        <div className="grid gap-6 md:grid-cols-2">
          {/* Work Days Chart */}
          <Card className="glass-card">
            <CardHeader className="pb-3">
              <CardTitle className="flex items-center gap-2 text-base">
                <BarChart3 className="h-5 w-5 text-primary" />
                各工程出工日數 (Top 20)
              </CardTitle>
            </CardHeader>
            <CardContent>
              {data.workDays?.length ? (
                <div className="space-y-2">
                  {data.workDays.slice(0, 20).map((item, i) => (
                    <div key={i} className="flex items-center gap-3">
                      <span className="text-xs text-muted-foreground w-28 truncate text-right" title={item.name}>{item.name}</span>
                      <div className="flex-1 h-6 bg-muted rounded-md overflow-hidden">
                        <div
                          className="h-full bg-primary rounded-md transition-all duration-500"
                          style={{ width: `${(item.days / maxWorkDays) * 100}%` }}
                        />
                      </div>
                      <Badge variant="secondary" className="text-xs min-w-[2rem] justify-center">{item.days}</Badge>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-center text-muted-foreground py-8">本月尚無出工資料</p>
              )}
            </CardContent>
          </Card>

          {/* Hazards Chart */}
          <Card className="glass-card">
            <CardHeader className="pb-3">
              <CardTitle className="flex items-center gap-2 text-base">
                <AlertTriangle className="h-5 w-5 text-destructive" />
                整體危害統計
              </CardTitle>
            </CardHeader>
            <CardContent>
              {data.hazards?.length ? (
                <div className="space-y-2">
                  {data.hazards.map((item, i) => (
                    <div key={i} className="flex items-center gap-3">
                      <span className="text-xs text-muted-foreground w-28 truncate text-right" title={item.type}>{item.type}</span>
                      <div className="flex-1 h-6 bg-muted rounded-md overflow-hidden">
                        <div
                          className="h-full bg-destructive rounded-md transition-all duration-500"
                          style={{ width: `${(item.count / maxHazards) * 100}%` }}
                        />
                      </div>
                      <Badge variant="secondary" className="text-xs min-w-[2rem] justify-center">{item.count}</Badge>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-center text-muted-foreground py-8">本月尚無危害記錄</p>
              )}
            </CardContent>
          </Card>
        </div>
      ) : (
        <p className="text-center text-muted-foreground py-20">無資料</p>
      )}
    </div>
  );
}
