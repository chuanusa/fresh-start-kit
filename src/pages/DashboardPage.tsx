import { useState, useEffect } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Badge } from '@/components/ui/badge';
import { getMonthlyDashboardData, type DashboardData } from '@/lib/gas-api';
import { useToast } from '@/hooks/use-toast';
import { BarChart3, AlertTriangle, Loader2, TrendingUp, Calendar, Shield } from 'lucide-react';

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

  const totalWorkDays = data?.workDays?.reduce((sum, d) => sum + d.days, 0) ?? 0;
  const totalHazards = data?.hazards?.reduce((sum, h) => sum + h.count, 0) ?? 0;
  const projectCount = data?.workDays?.length ?? 0;

  return (
    <div className="space-y-8 max-w-6xl">
      {/* Header */}
      <div className="flex items-center justify-between flex-wrap gap-4">
        <div>
          <h2 className="text-2xl font-bold tracking-tight gradient-text">每月儀表板</h2>
          <p className="text-muted-foreground text-sm mt-1">統計分析與視覺化圖表</p>
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
        <>
          {/* Stat Cards */}
          <div className="grid gap-4 grid-cols-1 sm:grid-cols-3">
            <Card className="glass-card hover-3d">
              <CardContent className="p-5 flex items-center gap-4">
                <div className="p-3 rounded-xl gradient-primary">
                  <Calendar className="h-5 w-5 text-white" />
                </div>
                <div>
                  <p className="text-xs text-muted-foreground font-medium">總出工日</p>
                  <p className="text-2xl font-bold">{totalWorkDays}</p>
                </div>
              </CardContent>
            </Card>
            <Card className="glass-card hover-3d">
              <CardContent className="p-5 flex items-center gap-4">
                <div className="p-3 rounded-xl bg-accent/20">
                  <TrendingUp className="h-5 w-5 text-accent" />
                </div>
                <div>
                  <p className="text-xs text-muted-foreground font-medium">進行中工程</p>
                  <p className="text-2xl font-bold">{projectCount}</p>
                </div>
              </CardContent>
            </Card>
            <Card className="glass-card hover-3d">
              <CardContent className="p-5 flex items-center gap-4">
                <div className="p-3 rounded-xl bg-destructive/20">
                  <Shield className="h-5 w-5 text-destructive" />
                </div>
                <div>
                  <p className="text-xs text-muted-foreground font-medium">危害紀錄</p>
                  <p className="text-2xl font-bold">{totalHazards}</p>
                </div>
              </CardContent>
            </Card>
          </div>

          {/* Charts */}
          <div className="grid gap-6 md:grid-cols-2">
            <Card className="glass-card hover-3d">
              <CardHeader className="pb-3">
                <CardTitle className="flex items-center gap-2 text-base">
                  <BarChart3 className="h-5 w-5 text-primary" />
                  各工程出工日數 (Top 20)
                </CardTitle>
              </CardHeader>
              <CardContent>
                {data.workDays?.length ? (
                  <div className="space-y-2.5">
                    {data.workDays.slice(0, 20).map((item, i) => (
                      <div key={i} className="flex items-center gap-3 group">
                        <span className="text-xs text-muted-foreground w-28 truncate text-right" title={item.name}>{item.name}</span>
                        <div className="flex-1 h-7 bg-muted/50 rounded-lg overflow-hidden">
                          <div
                            className="h-full gradient-primary rounded-lg transition-all duration-700 ease-out group-hover:brightness-110"
                            style={{ width: `${(item.days / maxWorkDays) * 100}%`, animationDelay: `${i * 50}ms` }}
                          />
                        </div>
                        <Badge variant="secondary" className="text-xs min-w-[2.5rem] justify-center font-mono">{item.days}</Badge>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="text-center text-muted-foreground py-8">本月尚無出工資料</p>
                )}
              </CardContent>
            </Card>

            <Card className="glass-card hover-3d">
              <CardHeader className="pb-3">
                <CardTitle className="flex items-center gap-2 text-base">
                  <AlertTriangle className="h-5 w-5 text-destructive" />
                  整體危害統計
                </CardTitle>
              </CardHeader>
              <CardContent>
                {data.hazards?.length ? (
                  <div className="space-y-2.5">
                    {data.hazards.map((item, i) => (
                      <div key={i} className="flex items-center gap-3 group">
                        <span className="text-xs text-muted-foreground w-28 truncate text-right" title={item.type}>{item.type}</span>
                        <div className="flex-1 h-7 bg-muted/50 rounded-lg overflow-hidden">
                          <div
                            className="h-full bg-gradient-to-r from-destructive to-warning rounded-lg transition-all duration-700 ease-out group-hover:brightness-110"
                            style={{ width: `${(item.count / maxHazards) * 100}%` }}
                          />
                        </div>
                        <Badge variant="secondary" className="text-xs min-w-[2.5rem] justify-center font-mono">{item.count}</Badge>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="text-center text-muted-foreground py-8">本月尚無危害記錄</p>
                )}
              </CardContent>
            </Card>
          </div>
        </>
      ) : (
        <div className="text-center py-20 text-muted-foreground">
          <BarChart3 className="h-12 w-12 mx-auto mb-3 opacity-20" />
          <p>無資料</p>
        </div>
      )}
    </div>
  );
}
