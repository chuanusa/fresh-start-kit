import { useState, useEffect } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Badge } from '@/components/ui/badge';
import { getMonthlyDashboardData, type DashboardData } from '@/lib/gas-api';
import { useToast } from '@/hooks/use-toast';
import { BarChart3, AlertTriangle, Loader2, TrendingUp, Calendar, Shield, Zap, Activity } from 'lucide-react';

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
    <div className="space-y-10">
      {/* Header */}
      <div className="flex items-end justify-between flex-wrap gap-4">
        <div className="space-y-2">
          <div className="flex items-center gap-2">
            <Activity className="h-5 w-5 text-accent" />
            <span className="text-xs font-semibold uppercase tracking-widest text-accent">Analytics</span>
          </div>
          <h2 className="text-3xl md:text-4xl font-bold tracking-tight gradient-text">每月儀表板</h2>
          <p className="text-muted-foreground text-sm">統計分析與視覺化圖表</p>
        </div>
        <div className="flex gap-2">
          <Select value={String(year)} onValueChange={(v: string) => setYear(Number(v))}>
            <SelectTrigger className="w-24 rounded-xl bg-card/60 backdrop-blur-sm border-border/50"><SelectValue /></SelectTrigger>
            <SelectContent>
              {years.map(y => <SelectItem key={y} value={String(y)}>{y}年</SelectItem>)}
            </SelectContent>
          </Select>
          <Select value={String(month)} onValueChange={(v: string) => setMonth(Number(v))}>
            <SelectTrigger className="w-20 rounded-xl bg-card/60 backdrop-blur-sm border-border/50"><SelectValue /></SelectTrigger>
            <SelectContent>
              {months.map(m => <SelectItem key={m} value={String(m)}>{m}月</SelectItem>)}
            </SelectContent>
          </Select>
        </div>
      </div>

      {loading ? (
        <div className="flex items-center justify-center py-32">
          <div className="relative">
            <Loader2 className="h-10 w-10 animate-spin text-primary" />
            <div className="absolute inset-0 h-10 w-10 animate-ping rounded-full bg-primary/10" />
          </div>
        </div>
      ) : data ? (
        <div className="space-y-8 stagger-children">
          {/* Stat Cards */}
          <div className="grid gap-5 grid-cols-1 sm:grid-cols-3">
            {/* Total Work Days */}
            <Card className="glass-card hover-3d gradient-border overflow-hidden group">
              <CardContent className="p-6 flex items-center gap-5">
                <div className="relative p-3.5 rounded-2xl gradient-primary shadow-lg shadow-primary/25 group-hover:shadow-primary/40 transition-shadow">
                  <Calendar className="h-6 w-6 text-primary-foreground" />
                  <div className="absolute -inset-1 rounded-2xl bg-primary/20 blur-lg opacity-0 group-hover:opacity-100 transition-opacity" />
                </div>
                <div>
                  <p className="text-xs text-muted-foreground font-semibold uppercase tracking-wider">總出工日</p>
                  <p className="text-3xl font-bold stat-value mt-0.5">{totalWorkDays}</p>
                </div>
              </CardContent>
            </Card>

            {/* Active Projects */}
            <Card className="glass-card hover-3d gradient-border overflow-hidden group">
              <CardContent className="p-6 flex items-center gap-5">
                <div className="relative p-3.5 rounded-2xl bg-accent/15 group-hover:bg-accent/25 transition-colors">
                  <TrendingUp className="h-6 w-6 text-accent" />
                  <div className="absolute -inset-1 rounded-2xl bg-accent/10 blur-lg opacity-0 group-hover:opacity-100 transition-opacity" />
                </div>
                <div>
                  <p className="text-xs text-muted-foreground font-semibold uppercase tracking-wider">進行中工程</p>
                  <p className="text-3xl font-bold stat-value mt-0.5">{projectCount}</p>
                </div>
              </CardContent>
            </Card>

            {/* Hazard Records */}
            <Card className="glass-card hover-3d gradient-border overflow-hidden group">
              <CardContent className="p-6 flex items-center gap-5">
                <div className="relative p-3.5 rounded-2xl bg-destructive/15 group-hover:bg-destructive/25 transition-colors">
                  <Shield className="h-6 w-6 text-destructive" />
                  <div className="absolute -inset-1 rounded-2xl bg-destructive/10 blur-lg opacity-0 group-hover:opacity-100 transition-opacity" />
                </div>
                <div>
                  <p className="text-xs text-muted-foreground font-semibold uppercase tracking-wider">危害紀錄</p>
                  <p className="text-3xl font-bold stat-value mt-0.5">{totalHazards}</p>
                </div>
              </CardContent>
            </Card>
          </div>

          {/* Charts */}
          <div className="grid gap-6 lg:grid-cols-2">
            {/* Work Days Chart */}
            <Card className="glass-card hover-3d overflow-hidden">
              <CardHeader className="pb-4">
                <CardTitle className="flex items-center gap-3 text-base">
                  <div className="p-2 rounded-xl gradient-primary shadow-md shadow-primary/15">
                    <BarChart3 className="h-4 w-4 text-primary-foreground" />
                  </div>
                  <div>
                    <span className="font-semibold">各工程出工日數</span>
                    <span className="text-xs text-muted-foreground ml-2">Top 20</span>
                  </div>
                </CardTitle>
              </CardHeader>
              <CardContent>
                {data.workDays?.length ? (
                  <div className="space-y-3">
                    {data.workDays.slice(0, 20).map((item, i) => (
                      <div key={i} className="flex items-center gap-3 group/bar">
                        <span className="text-xs text-muted-foreground w-24 truncate text-right font-medium" title={item.name}>{item.name}</span>
                        <div className="flex-1 h-8 bg-muted/30 rounded-xl overflow-hidden">
                          <div
                            className="h-full rounded-xl transition-all duration-700 ease-out group-hover/bar:brightness-110 relative overflow-hidden"
                            style={{
                              width: `${Math.max((item.days / maxWorkDays) * 100, 4)}%`,
                              background: `linear-gradient(135deg, hsl(245 58% 51%), hsl(270 60% 58%))`,
                              animationDelay: `${i * 60}ms`,
                            }}
                          >
                            <div className="absolute inset-0 animate-shimmer" />
                          </div>
                        </div>
                        <span className="text-xs font-bold stat-value min-w-[2rem] text-right text-foreground/80">{item.days}</span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="text-center py-12 text-muted-foreground">
                    <BarChart3 className="h-10 w-10 mx-auto mb-3 opacity-15" />
                    <p className="text-sm">本月尚無出工資料</p>
                  </div>
                )}
              </CardContent>
            </Card>

            {/* Hazards Chart */}
            <Card className="glass-card hover-3d overflow-hidden">
              <CardHeader className="pb-4">
                <CardTitle className="flex items-center gap-3 text-base">
                  <div className="p-2 rounded-xl bg-destructive/15">
                    <AlertTriangle className="h-4 w-4 text-destructive" />
                  </div>
                  <span className="font-semibold">整體危害統計</span>
                </CardTitle>
              </CardHeader>
              <CardContent>
                {data.hazards?.length ? (
                  <div className="space-y-3">
                    {data.hazards.map((item, i) => (
                      <div key={i} className="flex items-center gap-3 group/bar">
                        <span className="text-xs text-muted-foreground w-24 truncate text-right font-medium" title={item.type}>{item.type}</span>
                        <div className="flex-1 h-8 bg-muted/30 rounded-xl overflow-hidden">
                          <div
                            className="h-full rounded-xl transition-all duration-700 ease-out group-hover/bar:brightness-110"
                            style={{
                              width: `${Math.max((item.count / maxHazards) * 100, 4)}%`,
                              background: `linear-gradient(135deg, hsl(0 72% 51%), hsl(38 92% 50%))`,
                            }}
                          />
                        </div>
                        <span className="text-xs font-bold stat-value min-w-[2rem] text-right text-foreground/80">{item.count}</span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="text-center py-12 text-muted-foreground">
                    <AlertTriangle className="h-10 w-10 mx-auto mb-3 opacity-15" />
                    <p className="text-sm">本月尚無危害記錄</p>
                  </div>
                )}
              </CardContent>
            </Card>
          </div>
        </div>
      ) : (
        <div className="text-center py-32 text-muted-foreground">
          <div className="relative inline-block">
            <BarChart3 className="h-16 w-16 mx-auto mb-4 opacity-10" />
            <Zap className="h-6 w-6 absolute -top-1 -right-1 text-accent/30" />
          </div>
          <p className="text-lg font-medium">無資料</p>
          <p className="text-sm text-muted-foreground/60 mt-1">請選擇年月查看統計</p>
        </div>
      )}
    </div>
  );
}
