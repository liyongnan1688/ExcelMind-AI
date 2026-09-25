export interface SnapshotItem {
  id: string;
  timestamp: string;
  timeDisplay: string;
  fileName: string;
  originalPath: string;
  promptSummary: string;
  vbaPreview: string;
}

export interface WorkbookInfo {
  name: string;
  fullName: string;
  isSaved: boolean;
  sheets: string[];
  snapshots: SnapshotItem[];
}

export interface ScriptItem {
  name: string;
  fileName: string;
  filePath: string;
  createdAt: string;
  description: string;
  code: string;
}

export interface VbaExecutionData {
  summary: string;
  error?: string;
  elapsedMs: number;
  snapshot?: SnapshotItem;
  vbaCode: string;
}

export interface BridgeResponse<T = any> {
  ok: boolean;
  action: string;
  message?: string;
  data?: T;
  error?: string;
}

type MessageHandler = (res: BridgeResponse) => void;

class NativeBridgeClient {
  private handlers = new Map<string, MessageHandler>();
  private workbookChangeListeners: ((info: WorkbookInfo) => void)[] = [];

  constructor() {
    if (typeof window !== 'undefined' && (window as any).chrome?.webview) {
      (window as any).chrome.webview.addEventListener('message', (event: any) => {
        try {
          const raw = typeof event.data === 'string' ? JSON.parse(event.data) : event.data;
          if (raw.action === 'get_workbook_info' && raw.ok && raw.data) {
            this.workbookChangeListeners.forEach((fn) => fn(raw.data));
          }
          if (raw.action && this.handlers.has(raw.action)) {
            const handler = this.handlers.get(raw.action);
            handler?.(raw);
          }
        } catch (e) {
          console.error('[NativeBridgeClient] Parse message error:', e);
        }
      });
    }
  }

  public isNative(): boolean {
    return typeof window !== 'undefined' && Boolean((window as any).chrome?.webview);
  }

  public onWorkbookChange(fn: (info: WorkbookInfo) => void) {
    this.workbookChangeListeners.push(fn);
  }

  public async send<T = any>(action: string, payload: Record<string, any> = {}): Promise<BridgeResponse<T>> {
    if (!this.isNative()) {
      return this.mockResponse<T>(action, payload);
    }

    return new Promise((resolve) => {
      const msg = JSON.stringify({ action, ...payload });
      this.handlers.set(action, (res) => {
        resolve(res);
      });
      (window as any).chrome.webview.postMessage(msg);
    });
  }

  // 纯浏览器单测 Mock 模式
  private async mockResponse<T>(action: string, payload: any): Promise<BridgeResponse<T>> {
    await new Promise((r) => setTimeout(r, 200));

    if (action === 'get_workbook_info') {
      return {
        ok: true,
        action,
        data: {
          name: '2026年9月部门预算与绩效表.xlsx',
          fullName: 'C:\\Users\\35651\\Documents\\2026年9月部门预算与绩效表.xlsx',
          isSaved: true,
          sheets: ['汇总看板', '部门明细', '人员考评'],
          snapshots: [
            {
              id: 'snap_mock_1',
              timestamp: new Date().toISOString(),
              timeDisplay: '10:45:10',
              fileName: 'snap_1.xlsx',
              originalPath: 'C:\\test.xlsx',
              promptSummary: '自动标红超出预算的金额',
              vbaPreview: 'Sub HighlightOverBudget()...',
            },
          ],
        } as any,
      };
    }

    if (action === 'execute_vba') {
      return {
        ok: true,
        action,
        message: '执行成功：已按指令完成当前工作簿操作 (Mock 模式)',
        data: {
          summary: '执行成功：已按指令在“汇总看板”生成求和公式并标注高亮',
          elapsedMs: 382,
          vbaCode: payload.code,
          snapshot: {
            id: 'snap_' + Date.now(),
            timestamp: new Date().toISOString(),
            timeDisplay: new Date().toTimeString().slice(0, 8),
            fileName: 'snap_new.xlsx',
            originalPath: 'C:\\test.xlsx',
            promptSummary: payload.prompt || '操作快照',
            vbaPreview: payload.code.slice(0, 60),
          },
        } as any,
      };
    }

    if (action === 'list_scripts') {
      return {
        ok: true,
        action,
        data: [
          {
            name: '数据区域自动边框与隔行变色',
            fileName: '数据区域自动边框与隔行变色.bas',
            filePath: 'C:\\AppData\\LeeExcel\\Scripts\\数据区域自动边框与隔行变色.bas',
            createdAt: '2026-09-25 10:15',
            description: '自动寻找当前使用区域并添加细边框与浅绿隔行底色',
            code: 'Sub FormatTable()\n    Dim rng As Range\n    Set rng = ActiveSheet.UsedRange\n    rng.Borders.LineStyle = xlContinuous\nEnd Sub',
          },
        ] as any,
      };
    }

    return {
      ok: true,
      action,
      message: 'Mock 操作成功',
      data: {} as any,
    };
  }
}

export const bridge = new NativeBridgeClient();
