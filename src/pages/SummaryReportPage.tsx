import { useState, useEffect } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { getSummaryData, getGuestSummaryData, type SummaryData } from '@/lib/gas-api';
import { useAuth } from '@/contexts/AuthContext';
import { useToast } from '@/hooks/use-toast';
import { FileText, CalendarDays, Loader2 } from 'lucide-react';

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
    <div className="space-y-6 max-w-6xl">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-bold tracking-tight">工程總表</h2>
          <p className="text-muted-foreground text-sm">
            {isGuest ? '訪客檢視模式' : '每日施工總表'}
          </p>
        </div>
        {isGuest && (
          <Tabs value={viewMode} onValueChange={(v: string) => setViewMode(v as 'tomorrow' | 'today')}>
            <TabsList>
              <TabsTrigger value="tomorrow">
                <CalendarDays className="mr-1 h-3 w-3" /> 明日預定
              </TabsTrigger>
              <TabsTrigger value="today">
                <FileText className="mr-1 h-3 w-3" /> 今日施工
              </TabsTrigger>
            </TabsList>
          </Tabs>
        )}
      </div>

      <Card className="glass-card">
        <CardContent className="p-0">
          {loading ? (
            <div className="flex items-center justify-center py-20">
              <Loader2 className="h-8 w-8 animate-spin text-primary" />
            </div>
          ) : data.length > 0 ? (
            <div className="overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow className="bg-muted/50">
                    <TableHead className="font-semibold">日期</TableHead>
                    <TableHead className="font-semibold">工程名稱</TableHead>
                    <TableHead className="font-semibold">檢驗員</TableHead>
                    <TableHead className="font-semibold">狀態</TableHead>
                    <TableHead className="font-semibold">施工內容</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {data.map((row, i) => (
                    <TableRow key={i} className="hover:bg-muted/30">
                      <TableCell className="font-mono text-sm">{row.date}</TableCell>
                      <TableCell className="font-medium">{row.projectName}</TableCell>
                      <TableCell>{row.inspector}</TableCell>
                      <TableCell>
                        <Badge variant={statusColor(row.status)}>{row.status}</Badge>
                      </TableCell>
                      <TableCell className="text-muted-foreground text-sm max-w-xs truncate">
                        {row.workContent || '-'}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          ) : (
            <div className="text-center py-20 text-muted-foreground">
              <FileText className="h-12 w-12 mx-auto mb-3 opacity-30" />
              <p>尚無資料</p>
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  );
}
