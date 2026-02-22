import { useState, useEffect } from 'react';
import { Card, CardContent } from '@/components/ui/card';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Badge } from '@/components/ui/badge';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { getSummaryData, getGuestSummaryData, type SummaryData } from '@/lib/gas-api';
import { useAuth } from '@/contexts/AuthContext';
import { useToast } from '@/hooks/use-toast';
import { FileText, CalendarDays, Loader2, Sparkles } from 'lucide-react';

export function SummaryReportPage() {
  const { isGuest } = useAuth();
  const [data, setData] = useState<SummaryData[]>([]);
  const [loading, setLoading] = useState(false);
  const [viewMode, setViewMode] = useState<'tomorrow' | 'today'>('tomorrow');
  const { toast } = useToast();

  const loadData = async () => {
    setLoading(true);
    try {
      if (isGuest) {
        const result = await getGuestSummaryData(viewMode);
        setData(result || []);
      } else {
        const now = new Date();
        const result = await getSummaryData(now.getFullYear(), now.getMonth() + 1);
        setData(result || []);
      }
    } catch (e) {
      toast({ title: '載入總表失敗', description: (e as Error).message, variant: 'destructive' });
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, [viewMode, isGuest]);

  const statusColor = (status: string) => {
    switch (status) {
      case '已填': return 'default';
      case '未填': return 'destructive';
      default: return 'secondary';
    }
  };

  return (
    <div className="space-y-10">
      {/* Header */}
      <div className="flex items-end justify-between flex-wrap gap-4">
        <div className="space-y-2">
          <div className="flex items-center gap-2">
            <Sparkles className="h-4 w-4 text-accent" />
            <span className="text-xs font-semibold uppercase tracking-widest text-accent">
              {isGuest ? 'Guest View' : 'Overview'}
            </span>
          </div>
          <h2 className="text-3xl md:text-4xl font-bold tracking-tight gradient-text">工程總表</h2>
          <p className="text-muted-foreground text-sm">
            {isGuest ? '訪客檢視模式' : '每日施工總表'}
          </p>
        </div>
        {isGuest && (
          <Tabs value={viewMode} onValueChange={(v: string) => setViewMode(v as 'tomorrow' | 'today')}>
            <TabsList className="bg-card/60 backdrop-blur-sm border border-border/50 rounded-xl">
              <TabsTrigger value="tomorrow" className="rounded-lg">
                <CalendarDays className="mr-1.5 h-3.5 w-3.5" /> 明日預定
              </TabsTrigger>
              <TabsTrigger value="today" className="rounded-lg">
                <FileText className="mr-1.5 h-3.5 w-3.5" /> 今日施工
              </TabsTrigger>
            </TabsList>
          </Tabs>
        )}
      </div>

      {/* Table Card */}
      <Card className="glass-card overflow-hidden">
        <CardContent className="p-0">
          {loading ? (
            <div className="flex items-center justify-center py-32">
              <div className="relative">
                <Loader2 className="h-10 w-10 animate-spin text-primary" />
                <div className="absolute inset-0 h-10 w-10 animate-ping rounded-full bg-primary/10" />
              </div>
            </div>
          ) : data.length > 0 ? (
            <div className="overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow className="border-border/20 bg-muted/20">
                    <TableHead className="font-semibold text-foreground/70 text-xs uppercase tracking-wider">日期</TableHead>
                    <TableHead className="font-semibold text-foreground/70 text-xs uppercase tracking-wider">工程名稱</TableHead>
                    <TableHead className="font-semibold text-foreground/70 text-xs uppercase tracking-wider">檢驗員</TableHead>
                    <TableHead className="font-semibold text-foreground/70 text-xs uppercase tracking-wider">狀態</TableHead>
                    <TableHead className="font-semibold text-foreground/70 text-xs uppercase tracking-wider">施工內容</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {data.map((row, i) => (
                    <TableRow
                      key={i}
                      className="border-border/10 hover:bg-primary/[0.03] transition-colors duration-200"
                      style={{ animationDelay: `${i * 30}ms` }}
                    >
                      <TableCell className="font-mono text-sm text-muted-foreground">{row.date}</TableCell>
                      <TableCell className="font-semibold">{row.projectName}</TableCell>
                      <TableCell className="text-foreground/80">{row.inspector}</TableCell>
                      <TableCell>
                        <Badge variant={statusColor(row.status)} className="rounded-lg font-semibold text-[11px]">{row.status}</Badge>
                      </TableCell>
                      <TableCell className="text-muted-foreground text-sm max-w-xs truncate">
                        {row.workContent || '—'}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          ) : (
            <div className="text-center py-32 text-muted-foreground">
              <FileText className="h-16 w-16 mx-auto mb-4 opacity-10" />
              <p className="text-lg font-medium">尚無資料</p>
              <p className="text-sm text-muted-foreground/60 mt-1">目前沒有施工紀錄</p>
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  );
}
