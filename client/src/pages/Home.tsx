/**
 * 日规划转换工具 - 主页面
 * 
 * Design Philosophy: Notion-style Neo-Minimalism
 * - 功能至上，界面元素服务于功能
 * - 大量留白，让内容自然呼吸
 * - 微妙层次，通过极细微的阴影和边框建立层次
 * - 即时反馈，每个操作都有清晰的视觉反馈
 */

import { useState, useCallback, useMemo } from 'react';
import { Button } from '@/components/ui/button';
import { Card } from '@/components/ui/card';
import { Textarea } from '@/components/ui/textarea';
import { toast } from 'sonner';
import { 
  FileSpreadsheet, 
  Download, 
  AlertCircle, 
  CheckCircle2, 
  Info,
  Trash2,
  FileText,
  ArrowRight,
  Sparkles,
  Calendar
} from 'lucide-react';
import { parseMarkdown, downloadExcel, exampleMarkdown, type ParseResult } from '@/lib/converter';

export default function Home() {
  const [markdownInput, setMarkdownInput] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);

  // 实时解析 Markdown
  const parseResult = useMemo<ParseResult | null>(() => {
    if (!markdownInput.trim()) return null;
    return parseMarkdown(markdownInput);
  }, [markdownInput]);

  // 处理下载
  const handleDownload = useCallback(async () => {
    if (!parseResult || parseResult.errors.length > 0) {
      toast.error('请先修复错误后再下载');
      return;
    }

    setIsProcessing(true);
    try {
      const filename = downloadExcel(parseResult);
      toast.success(`Excel 文件已生成`, {
        description: filename,
      });
    } catch (error) {
      toast.error('生成 Excel 文件失败', {
        description: error instanceof Error ? error.message : '未知错误',
      });
    } finally {
      setIsProcessing(false);
    }
  }, [parseResult]);

  // 加载示例
  const handleLoadExample = useCallback(() => {
    setMarkdownInput(exampleMarkdown);
    toast.success('已加载示例内容');
  }, []);

  // 清空内容
  const handleClear = useCallback(() => {
    setMarkdownInput('');
    toast.info('已清空内容');
  }, []);

  // 判断是否可以下载
  const canDownload = parseResult && parseResult.errors.length === 0 && parseResult.dataRows.length > 0;

  return (
    <div className="min-h-screen bg-background">
      {/* Header */}
      <header className="border-b border-border/50 bg-white/80 backdrop-blur-sm sticky top-0 z-10">
        <div className="container py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-lg bg-primary/10 flex items-center justify-center">
                <Calendar className="w-5 h-5 text-primary" />
              </div>
              <div>
                <h1 className="text-lg font-semibold text-foreground">夜将晨-日规划转换工具</h1>
                <p className="text-xs text-muted-foreground">Markdown → Excel</p>
              </div>
            </div>
            <div className="flex items-center gap-2">
              <Button
                variant="ghost"
                size="sm"
                onClick={handleLoadExample}
                className="text-muted-foreground hover:text-foreground"
              >
                <Sparkles className="w-4 h-4 mr-1.5" />
                加载示例
              </Button>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="container py-8">
        <div className="max-w-4xl mx-auto space-y-6 animate-fade-in">
          
          {/* Input Section */}
          <Card className="p-6 border border-border/60 shadow-sm">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2">
                <FileText className="w-4 h-4 text-muted-foreground" />
                <h2 className="font-medium text-foreground">输入 Markdown 内容</h2>
              </div>
              {markdownInput && (
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={handleClear}
                  className="text-muted-foreground hover:text-destructive h-8"
                >
                  <Trash2 className="w-3.5 h-3.5 mr-1" />
                  清空
                </Button>
              )}
            </div>
            
            <Textarea
              value={markdownInput}
              onChange={(e) => setMarkdownInput(e.target.value)}
              placeholder={`粘贴您的日规划 Markdown 内容...

格式示例：
**学生名字日规划（1.13 - 1.18）执行表**
**本周核心目标：** ...

| 日期 | 星期 | 语文 | 数学 | 英语 | 物理 | 化学 | 生物 |
| ---- | ---- | ---- | ---- | ---- | ---- | ---- | ---- |
| 1.13 | 周一 | ... | ... | ... | ... | ... | ... |`}
              className="min-h-[280px] resize-y bg-muted/30 border-border/60 focus:border-primary/50 focus:ring-1 focus:ring-primary/20"
            />
            
            <p className="mt-3 text-xs text-muted-foreground">
              提示：直接从编辑器复制 Markdown 格式的日规划内容粘贴到此处
            </p>
          </Card>

          {/* Parse Result Section */}
          {parseResult && (
            <Card className="p-6 border border-border/60 shadow-sm animate-fade-in">
              <div className="flex items-center gap-2 mb-4">
                <Info className="w-4 h-4 text-muted-foreground" />
                <h2 className="font-medium text-foreground">解析结果</h2>
              </div>

              {/* Errors */}
              {parseResult.errors.length > 0 && (
                <div className="mb-4 p-4 rounded-lg bg-destructive/5 border border-destructive/20">
                  <div className="flex items-start gap-2">
                    <AlertCircle className="w-4 h-4 text-destructive mt-0.5 flex-shrink-0" />
                    <div>
                      <p className="font-medium text-destructive text-sm">解析错误</p>
                      <ul className="mt-1 text-sm text-destructive/80 space-y-0.5">
                        {parseResult.errors.map((error, i) => (
                          <li key={i}>• {error}</li>
                        ))}
                      </ul>
                    </div>
                  </div>
                </div>
              )}

              {/* Warnings */}
              {parseResult.warnings.length > 0 && (
                <div className="mb-4 p-4 rounded-lg bg-amber-50 border border-amber-200/60">
                  <div className="flex items-start gap-2">
                    <AlertCircle className="w-4 h-4 text-amber-600 mt-0.5 flex-shrink-0" />
                    <div>
                      <p className="font-medium text-amber-700 text-sm">警告</p>
                      <ul className="mt-1 text-sm text-amber-600 space-y-0.5">
                        {parseResult.warnings.map((warning, i) => (
                          <li key={i}>• {warning}</li>
                        ))}
                      </ul>
                    </div>
                  </div>
                </div>
              )}

              {/* Success Info */}
              {parseResult.errors.length === 0 && (
                <div className="space-y-4">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div className="p-4 rounded-lg bg-muted/40 border border-border/40">
                      <p className="text-xs text-muted-foreground mb-1">学生姓名</p>
                      <p className="font-medium text-foreground">{parseResult.studentName || '未识别'}</p>
                    </div>
                    <div className="p-4 rounded-lg bg-muted/40 border border-border/40">
                      <p className="text-xs text-muted-foreground mb-1">日期范围</p>
                      <p className="font-medium text-foreground">{parseResult.dateRange || '未识别'}</p>
                    </div>
                  </div>

                  {parseResult.coreTarget && (
                    <div className="p-4 rounded-lg bg-muted/40 border border-border/40">
                      <p className="text-xs text-muted-foreground mb-1">本周核心目标</p>
                      <p className="text-sm text-foreground leading-relaxed line-clamp-3">
                        {parseResult.coreTarget}
                      </p>
                    </div>
                  )}

                  <div className="p-4 rounded-lg bg-primary/5 border border-primary/20">
                    <div className="flex items-center gap-2">
                      <CheckCircle2 className="w-4 h-4 text-primary" />
                      <p className="text-sm text-primary font-medium">
                        成功解析 {parseResult.dataRows.length} 天的日规划数据（{parseResult.columns.length - 2} 个学科）
                      </p>
                    </div>
                  </div>
                </div>
              )}
            </Card>
          )}

          {/* Download Section */}
          <div className="flex justify-center pt-2">
            <Button
              size="lg"
              onClick={handleDownload}
              disabled={!canDownload || isProcessing}
              className="px-8 h-12 text-base font-medium shadow-sm hover:shadow-md transition-all duration-200 disabled:opacity-50"
            >
              {isProcessing ? (
                <>
                  <div className="w-4 h-4 mr-2 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                  生成中...
                </>
              ) : (
                <>
                  <Download className="w-4 h-4 mr-2" />
                  下载 Excel 文件
                  <ArrowRight className="w-4 h-4 ml-2" />
                </>
              )}
            </Button>
          </div>

          {/* Help Section */}
          <Card className="p-6 border border-border/40 bg-muted/20">
            <h3 className="font-medium text-foreground mb-3">使用说明</h3>
            <div className="space-y-3 text-sm text-muted-foreground">
              <div className="flex items-start gap-3">
                <span className="w-6 h-6 rounded-full bg-primary/10 text-primary text-xs font-medium flex items-center justify-center flex-shrink-0">1</span>
                <p>将 Markdown 格式的日规划内容粘贴到上方输入框</p>
              </div>
              <div className="flex items-start gap-3">
                <span className="w-6 h-6 rounded-full bg-primary/10 text-primary text-xs font-medium flex items-center justify-center flex-shrink-0">2</span>
                <p>系统会自动解析学生姓名、日期、核心目标和每日任务表格</p>
              </div>
              <div className="flex items-start gap-3">
                <span className="w-6 h-6 rounded-full bg-primary/10 text-primary text-xs font-medium flex items-center justify-center flex-shrink-0">3</span>
                <p>确认解析结果无误后，点击"下载 Excel 文件"按钮</p>
              </div>
            </div>
            
            <div className="mt-4 pt-4 border-t border-border/40">
              <p className="text-xs text-muted-foreground">
                <strong>输入格式要求：</strong>第一行为标题（包含学生姓名和日期范围），第二行为本周核心目标，后面是 Markdown 表格（包含日期、星期和各学科列）
              </p>
            </div>
          </Card>
        </div>
      </main>

      {/* Footer */}
      <footer className="border-t border-border/40 mt-auto">
        <div className="container py-4">
          <p className="text-center text-xs text-muted-foreground">
            夜将晨-日规划 Markdown 转 Excel 工具 v1.0
          </p>
        </div>
      </footer>
    </div>
  );
}
