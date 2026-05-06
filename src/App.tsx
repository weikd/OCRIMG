/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useMemo, useEffect, ChangeEvent } from 'react';
import { 
  Upload, 
  Table as TableIcon, 
  FileText, 
  RotateCcw, 
  Download, 
  Settings2, 
  Check, 
  AlertCircle,
  FileSpreadsheet,
  FileDown as FileWord,
  Loader2,
  Trash2,
  Type,
  Copy,
  ExternalLink
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, Table as DocxTable, TableCell, TableRow, WidthType } from 'docx';
import { saveAs } from 'file-saver';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// --- Utils ---
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Types ---
type RecognitionMode = 'offline' | 'qwen';
type ContentType = 'table' | 'text';

interface TableData {
  headers: string[];
  rows: string[][];
}

// --- App Component ---
export default function App() {
  const [image, setImage] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [mode, setMode] = useState<RecognitionMode>('offline');
  const [qwenApiKey, setQwenApiKey] = useState(() => localStorage.getItem('QWEN_API_KEY') || '');
  const [resultType, setResultType] = useState<ContentType | null>(null);
  const [tableData, setTableData] = useState<TableData | null>(null);
  const [textContent, setTextContent] = useState<string>('');
  const [error, setError] = useState<string | null>(null);
  const [showSettings, setShowSettings] = useState(false);

  useEffect(() => {
    localStorage.setItem('QWEN_API_KEY', qwenApiKey);
  }, [qwenApiKey]);

  // --- Helpers ---
  const extractJson = (text: string) => {
    try {
      const match = text.match(/\{[\s\S]*\}/);
      if (match) return JSON.parse(match[0]);
      return JSON.parse(text);
    } catch (e) {
      console.error("JSON Parse Error:", e);
      return null;
    }
  };

  // --- Handlers ---
  const handleImageUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (event) => {
        setImage(event.target?.result as string);
        setError(null);
        setTableData(null);
        setTextContent('');
        setResultType(null);
      };
      reader.readAsDataURL(file);
    }
  };

  const processImage = async () => {
    if (!image) return;
    setIsProcessing(true);
    setError(null);

    const base64 = image.split(',')[1];
    
    try {
      if (mode === 'offline') {
        await processWithTesseract();
      } else {
        await processWithQwen(base64);
      }
    } catch (err: any) {
      console.error(err);
      setError(err.message || '识别失败，请检查配置或图片内容');
    } finally {
      setIsProcessing(false);
    }
  };

  const processWithTesseract = async () => {
    const { createWorker } = await import('tesseract.js');
    const worker = await createWorker('chi_sim+eng');
    const { data: { text } } = await worker.recognize(image!);
    setResultType('text');
    setTextContent(text);
    await worker.terminate();
  };

  const processWithQwen = async (base64: string) => {
    if (!qwenApiKey) throw new Error('使用千问模式需要配置 API Key，请点击右上角设置图标填写');

    const response = await fetch('https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${qwenApiKey}`
      },
      body: JSON.stringify({
        model: "qwen-vl-max",
        messages: [
          {
            role: "user",
            content: [
              { type: "text", text: "你是一个专业的 OCR 处理专家。请识别图片中的所有内容。如果包含表格，请严格输出此 JSON 格式：{\"type\": \"table\", \"headers\": [\"列名1\", \"列名2\"], \"rows\": [[\"值1\", \"值2\"]]}。如果是文档，返回：{\"type\": \"text\", \"content\": \"文本全文\"}。严禁输出 JSON 之外的代码块标记或解释文字。" },
              { type: "image_url", image_url: { url: `data:image/jpeg;base64,${base64}` } }
            ]
          }
        ],
        max_tokens: 2000
      })
    });

    if (!response.ok) {
      const err = await response.json();
      throw new Error(err.message || '千问 API 调用失败');
    }

    const json = await response.json();
    const content = json.choices[0].message.content;
    const result = extractJson(content);
    
    if (!result) {
      setResultType('text');
      setTextContent(content);
      return;
    }

    if (result.type === 'table' || result.headers) {
      setResultType('table');
      setTableData({ 
        headers: result.headers || [], 
        rows: result.rows || [] 
      });
    } else {
      setResultType('text');
      setTextContent(result.content || result.text || "");
    }
  };

  const transposeTable = () => {
    if (!tableData) return;
    const { headers, rows } = tableData;
    const allData = [headers, ...rows];
    const transposed = allData[0].map((_, colIndex) => allData.map(row => row[colIndex]));
    setTableData({
      headers: transposed[0],
      rows: transposed.slice(1)
    });
  };

  const handleCellEdit = (rowIndex: number, colIndex: number, value: string) => {
    if (!tableData) return;
    const newRows = [...tableData.rows];
    if (rowIndex === -1) {
      const newHeaders = [...tableData.headers];
      newHeaders[colIndex] = value;
      setTableData({ ...tableData, headers: newHeaders });
    } else {
      newRows[rowIndex][colIndex] = value;
      setTableData({ ...tableData, rows: newRows });
    }
  };

  const exportToExcel = () => {
    if (!tableData) return;
    const ws = XLSX.utils.aoa_to_sheet([tableData.headers, ...tableData.rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, "识别结果.xlsx");
  };

  const exportToWord = async () => {
    let doc: Document;
    if (resultType === 'table' && tableData) {
      doc = new Document({
        sections: [{
          children: [
            new DocxTable({
              rows: [
                new TableRow({
                  children: tableData.headers.map(h => new TableCell({ children: [new Paragraph(h)] })),
                }),
                ...tableData.rows.map(row => new TableRow({
                  children: row.map(cell => new TableCell({ children: [new Paragraph(cell)] })),
                }))
              ],
              width: { size: 100, type: WidthType.PERCENTAGE }
            })
          ]
        }]
      });
    } else {
      doc = new Document({
        sections: [{
          children: textContent.split('\n').map(line => new Paragraph(line))
        }]
      });
    }

    const blob = await Packer.toBlob(doc);
    saveAs(blob, "识别结果.docx");
  };

  const copyToClipboard = () => {
    navigator.clipboard.writeText(textContent);
    // Simple alert-less feedback could be added here, but for brevity we'll stick to basic
  };

  return (
    <div className="flex flex-col h-screen bg-bg-base text-text-main font-sans overflow-hidden">
      {/* Header Section */}
      <header className="h-16 flex items-center justify-between px-6 bg-white border-b border-border-base shadow-sm shrink-0">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-primary rounded-lg flex items-center justify-center text-white font-bold text-xl shadow-lg shadow-blue-100">
            Q
          </div>
          <div>
            <h1 className="text-lg font-bold leading-tight tracking-tight font-display">Q-Flow OCR Engine</h1>
            <p className="text-[10px] text-text-muted uppercase tracking-widest font-bold">万能图片识别助手</p>
          </div>
        </div>

        <div className="flex items-center gap-4">
          <nav className="flex bg-panel-bg p-1 rounded-md border border-border-base">
            <button 
              onClick={() => setMode('offline')}
              className={cn(
                "px-4 py-1.5 text-xs font-semibold rounded transition-all",
                mode === 'offline' ? "bg-white shadow-sm text-primary" : "text-text-muted hover:text-text-main"
              )}
            >
              本地离线
            </button>
            <button 
              onClick={() => setMode('qwen')}
              className={cn(
                "px-4 py-1.5 text-xs font-semibold rounded transition-all",
                mode === 'qwen' ? "bg-white shadow-sm text-primary" : "text-text-muted hover:text-text-main"
              )}
            >
              千问 API
            </button>
          </nav>
          
          <div className="h-8 w-[1px] bg-border-base mx-2"></div>
          
          <button 
            onClick={() => setShowSettings(!showSettings)}
            className={cn(
              "px-3 py-2 rounded-lg transition-all border flex items-center gap-2",
              showSettings || (!qwenApiKey && mode === 'qwen') 
                ? "bg-primary/10 border-primary/20 text-primary" 
                : "text-text-muted hover:bg-panel-bg border-transparent"
            )}
          >
            <Settings2 className={cn("w-5 h-5", !qwenApiKey && mode === 'qwen' && "animate-pulse")} />
            {!qwenApiKey && mode === 'qwen' && (
              <span className="text-[10px] font-bold animate-pulse text-primary tracking-tight">配置 API KEY</span>
            )}
          </button>
        </div>
      </header>

      {/* Main Workspace */}
      <main className="flex-1 flex overflow-hidden">
        {/* Sidebar Controls */}
        <aside className="w-72 bg-white border-r border-border-base p-6 flex flex-col gap-8 shrink-0 overflow-y-auto">
          {/* API Setup Banner (Prominent when needed) */}
          <AnimatePresence initial={false}>
            {(showSettings || (!qwenApiKey && mode === 'qwen')) && (
              <motion.div 
                initial={{ height: 0, opacity: 0 }}
                animate={{ height: 'auto', opacity: 1 }}
                exit={{ height: 0, opacity: 0 }}
                className="overflow-hidden border-b border-border-base pb-6"
              >
                <div className="p-4 bg-primary/5 rounded-xl border border-primary/10 space-y-4">
                  <div className="flex items-center justify-between">
                    <h3 className="text-[11px] font-bold text-primary uppercase tracking-wider flex items-center gap-2">
                       <Settings2 className="w-3 h-3" /> 千问 API 配置
                    </h3>
                  </div>
                  <div className="space-y-2">
                    <div className="flex items-center justify-between">
                      <label className="text-[10px] font-bold text-text-muted uppercase">Qwen API Key</label>
                      <a 
                        href="https://bailian.console.aliyun.com/cn-beijing/?tab=model#/api-key" 
                        target="_blank" 
                        rel="noreferrer"
                        className="text-[10px] text-primary hover:underline flex items-center gap-1 font-bold"
                      >
                        获取 Key <ExternalLink className="w-2.5 h-2.5" />
                      </a>
                    </div>
                    <input 
                      type="password" 
                      value={qwenApiKey}
                      onChange={(e) => setQwenApiKey(e.target.value)}
                      placeholder="sk-xxxxxxxx"
                      className="w-full bg-white border border-border-base rounded-md px-3 py-2 text-xs focus:ring-2 focus:ring-primary/20 focus:border-primary outline-none transition-all placeholder:opacity-50"
                    />
                    <p className="text-[9px] text-text-muted leading-relaxed italic">
                      Key 仅保存在本地。您可以点击右上角图标收起此面板。
                    </p>
                  </div>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
          <div>
            <label className="text-[11px] font-bold text-text-muted uppercase tracking-wider mb-3 block">上传待识别图片</label>
            <div 
              className={cn(
                "border-2 border-dashed rounded-xl transition-all duration-300 relative",
                image 
                  ? "border-transparent" 
                  : "border-border-base hover:border-primary hover:bg-blue-50/50 p-8 text-center cursor-pointer bg-slate-50/50"
              )}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                const file = e.dataTransfer.files[0];
                if (file) {
                  const reader = new FileReader();
                  reader.onload = (ev) => setImage(ev.target?.result as string);
                  reader.readAsDataURL(file);
                }
              }}
            >
              {image ? (
                <div className="relative group rounded-lg overflow-hidden border border-border-base shadow-sm">
                  <img src={image} alt="Preview" className="w-full h-auto object-contain max-h-[300px]" />
                  <div className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center gap-2">
                    <button 
                      onClick={() => setImage(null)}
                      className="p-2 bg-white text-red-600 rounded-full shadow-lg hover:scale-110 transition-transform"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              ) : (
                <label className="cursor-pointer flex flex-col items-center">
                  <input type="file" className="hidden" accept="image/*" onChange={handleImageUpload} />
                  <Upload className="w-8 h-8 text-slate-300 mb-3" />
                  <p className="text-xs text-text-main font-semibold">Drop JPG, PNG, WEBP</p>
                  <p className="text-[10px] text-text-muted mt-1">or click to browse</p>
                </label>
              )}
            </div>
            
            {image && !isProcessing && !resultType && (
              <button 
                onClick={processImage}
                className="w-full mt-4 py-3 bg-primary text-white rounded-lg text-sm font-bold hover:bg-primary-hover shadow-lg shadow-blue-200 transition-all active:scale-[0.98]"
              >
                开始智能识别
              </button>
            )}
            
            {isProcessing && (
              <div className="mt-4 p-3 bg-blue-50 border border-blue-100 rounded-lg flex items-center justify-center gap-2">
                <Loader2 className="w-4 h-4 text-primary animate-spin" />
                <span className="text-xs font-bold text-primary">正在解析中...</span>
              </div>
            )}
            
            {error && (
              <div className="mt-4 p-3 bg-red-50 border border-red-100 rounded-lg flex items-start gap-2">
                <AlertCircle className="w-4 h-4 text-red-500 shrink-0 mt-0.5" />
                <p className="text-xs text-red-700 font-medium leading-relaxed">{error}</p>
              </div>
            )}
          </div>

          {/* Settings Section (Conditional) */}
          <AnimatePresence>
            {showSettings && (
              <motion.div 
                initial={{ height: 0, opacity: 0 }}
                animate={{ height: 'auto', opacity: 1 }}
                exit={{ height: 0, opacity: 0 }}
                className="overflow-hidden border-t border-border-base pt-6"
              >
                <div className="space-y-4">
                  <h3 className="text-[11px] font-bold text-text-muted uppercase tracking-wider flex items-center gap-2">
                    <Settings2 className="w-3 h-3" /> API 配置 (用于 GitHub 导出后)
                  </h3>
                  <div className="grid grid-cols-1 gap-4">
                    <div className="space-y-2">
                      <div className="flex items-center justify-between">
                        <label className="text-[10px] font-bold text-text-muted uppercase">Qwen API Key</label>
                        <a 
                          href="https://bailian.console.aliyun.com/cn-beijing/?tab=model#/api-key" 
                          target="_blank" 
                          rel="noreferrer"
                          className="text-[10px] text-primary hover:underline flex items-center gap-1 font-bold"
                        >
                          获取 Key <ExternalLink className="w-2.5 h-2.5" />
                        </a>
                      </div>
                      <input 
                        type="password" 
                        value={qwenApiKey}
                        onChange={(e) => setQwenApiKey(e.target.value)}
                        placeholder="sk-..."
                        className="w-full bg-white border border-border-base rounded-md px-3 py-2 text-xs focus:ring-2 focus:ring-primary/20 focus:border-primary outline-none transition-all"
                      />
                    </div>
                  </div>
                </div>
              </motion.div>
            )}
          </AnimatePresence>

          {/* Stats/Status */}
          <div className="mt-auto space-y-4">
            <div className="p-4 bg-slate-900 rounded-xl text-white shadow-xl">
              <div className="flex items-center justify-between mb-4">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">System Info</span>
                <div className="flex items-center gap-1.5">
                  <div className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse"></div>
                  <span className="text-[10px] text-emerald-400 font-bold">READY</span>
                </div>
              </div>
              <div className="grid grid-cols-1 gap-3">
                <div className="flex justify-between items-center text-xs">
                  <span className="text-slate-400">OCR Engine</span>
                  <span className="font-bold">
                    {mode === 'offline' ? 'Tesseract (本地)' : 'Qwen VL Max'}
                  </span>
                </div>
                <div className="flex justify-between items-center text-xs">
                  <span className="text-slate-400">Result Type</span>
                  <span className="font-bold text-blue-400">{resultType ? (resultType === 'table' ? '表格' : '文本') : '--'}</span>
                </div>
              </div>
            </div>
            
            <p className="text-[10px] text-center text-text-muted font-medium">支持: JPG, PNG, WEBP</p>
          </div>
        </aside>

        {/* Content Area */}
        <section className="flex-1 flex flex-col bg-bg-base p-6 overflow-hidden">
          {/* Editor Toolbar */}
          <div className="bg-white rounded-t-xl border border-border-base px-5 py-3 flex items-center justify-between shadow-sm shrink-0">
            <div className="flex items-center gap-2">
              <div className="flex gap-1 pr-5 border-r border-border-base">
                 {resultType === 'table' && (
                    <button 
                      onClick={transposeTable}
                      className="p-1.5 rounded hover:bg-panel-bg text-primary bg-blue-50/50 flex items-center gap-2 transition-all"
                    >
                      <RotateCcw className="w-4 h-4" />
                      <span className="text-[11px] font-bold">转置表格</span>
                    </button>
                 )}
                 {resultType === 'text' && (
                    <button 
                      onClick={copyToClipboard}
                      className="p-1.5 rounded hover:bg-panel-bg text-text-muted flex items-center gap-2 transition-all"
                    >
                      <Copy className="w-4 h-4" />
                      <span className="text-[11px] font-bold">复制文本</span>
                    </button>
                 )}
              </div>
              
              <div className="flex gap-4 px-4">
                 <div className="flex gap-1">
                    <button className="w-7 h-7 flex items-center justify-center rounded hover:bg-panel-bg text-sm font-bold">B</button>
                    <button className="w-7 h-7 flex items-center justify-center rounded hover:bg-panel-bg text-sm italic font-serif">I</button>
                 </div>
              </div>
              
              {resultType && (
                <span className="text-xs text-text-muted font-medium italic animate-pulse">
                  数据已解析 - 可直接在下方编辑...
                </span>
              )}
            </div>

            <div className="flex items-center gap-2">
              {resultType && (
                <>
                  <button 
                    onClick={exportToExcel}
                    className={cn(
                      "flex items-center gap-2 px-3 py-1.5 bg-white border border-slate-300 rounded-md text-[11px] font-bold hover:bg-bg-base transition-all",
                      resultType !== 'table' && "hidden"
                    )}
                  >
                    <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
                    Excel 导出
                  </button>
                  <button 
                    onClick={exportToWord}
                    className="flex items-center gap-2 px-3 py-1.5 bg-primary text-white rounded-md text-[11px] font-bold hover:bg-primary-hover shadow-lg shadow-blue-100 transition-all"
                  >
                    <FileWord className="w-3.5 h-3.5" />
                    Word 导出
                  </button>
                </>
              )}
              {resultType && (
                <div className="ml-2 text-[10px] font-bold text-primary px-2 py-1 bg-blue-50 rounded-full border border-blue-200 uppercase">
                  {resultType === 'table' ? 'Table View' : 'Text Editor'}
                </div>
              )}
            </div>
          </div>

          {/* Result Editor */}
          <div className="flex-1 bg-white border-x border-b border-border-base overflow-hidden rounded-b-xl flex flex-col shadow-sm">
            {!resultType && !isProcessing && (
              <div className="flex-1 flex flex-col items-center justify-center opacity-40 select-none px-12 text-center bg-slate-50/30">
                <div className="w-20 h-20 bg-white border border-border-base rounded-2xl flex items-center justify-center mb-6 shadow-sm">
                  <FileText className="w-10 h-10 text-slate-300" />
                </div>
                <h3 className="text-lg font-bold mb-2">等待数据输入</h3>
                <p className="text-sm max-w-xs">上传图片并选择识别模式，解析后的可编辑结果将呈现在这里。</p>
              </div>
            )}

            {isProcessing && (
              <div className="flex-1 flex flex-col items-center justify-center bg-white">
                <div className="w-48 h-1.5 bg-slate-100 rounded-full overflow-hidden mb-4">
                  <motion.div 
                    initial={{ width: 0 }}
                    animate={{ width: "100%" }}
                    transition={{ duration: 2, repeat: Infinity }}
                    className="h-full bg-primary"
                  />
                </div>
                <p className="text-xs font-bold text-text-muted uppercase tracking-widest">正在运行视觉解析引擎...</p>
              </div>
            )}

            <div className="flex-1 overflow-auto bg-white custom-scrollbar">
              {resultType === 'table' && tableData && (
                <table className="w-full border-collapse text-sm">
                  <thead className="sticky top-0 z-10">
                    <tr className="bg-bg-base border-b border-border-base">
                      <th className="w-12 p-3 text-center text-text-muted text-[10px] font-bold border-r border-border-base">#</th>
                      {tableData.headers.map((h, i) => (
                        <th key={i} className="p-3 text-left font-bold text-text-main border-r border-border-base last:border-0 relative min-w-[120px]">
                          <input 
                            value={h}
                            onChange={(e) => handleCellEdit(-1, i, e.target.value)}
                            className="w-full bg-transparent outline-none focus:text-primary transition-colors text-ellipsis"
                          />
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {tableData.rows.map((row, rowIndex) => (
                      <tr key={rowIndex} className="border-b border-slate-50 hover:bg-slate-50 transition-colors group">
                        <td className="p-3 text-center text-text-muted text-xs border-r border-slate-100 bg-slate-50/30">{rowIndex + 1}</td>
                        {row.map((cell, colIndex) => (
                          <td key={colIndex} className="p-0 border-r border-slate-100 last:border-0">
                            <textarea
                              rows={1}
                              value={cell}
                              onChange={(e) => {
                                handleCellEdit(rowIndex, colIndex, e.target.value);
                                e.target.style.height = 'auto';
                                e.target.style.height = e.target.scrollHeight + 'px';
                              }}
                              className="w-full p-3 bg-transparent outline-none focus:bg-blue-50/50 focus:ring-1 focus:ring-inset focus:ring-primary/20 resize-none overflow-hidden block transition-all"
                              style={{ minHeight: '100%' }}
                            />
                          </td>
                        ))}
                      </tr>
                    ))}
                    {/* Add Row Placeholder */}
                    <tr className="bg-slate-50/30">
                      <td colSpan={tableData.headers.length + 1} className="p-4 text-center">
                         <p className="text-[10px] text-text-muted font-bold uppercase tracking-widest">以上为完整识别结果</p>
                      </td>
                    </tr>
                  </tbody>
                </table>
              )}

              {resultType === 'text' && (
                <div className="p-10 h-full max-w-4xl mx-auto">
                  <textarea 
                    value={textContent}
                    onChange={(e) => setTextContent(e.target.value)}
                    className="w-full h-full bg-white outline-none font-sans text-base leading-loose resize-none custom-scrollbar"
                    placeholder="提取的文本将在这里解析..."
                    style={{ minHeight: '100%' }}
                  />
                </div>
              )}
            </div>
          </div>
        </section>
      </main>

      {/* Footer Bar */}
      <footer className="h-8 bg-slate-900 text-slate-400 px-6 flex items-center justify-between shrink-0">
        <div className="flex items-center gap-6">
          <div className="flex items-center gap-1.5">
            <div className="w-1.5 h-1.5 rounded-full bg-emerald-500"></div>
            <span className="text-[10px] uppercase font-extrabold tracking-tight">System Online</span>
          </div>
          <span className="text-[10px] font-medium opacity-80">支持格式: JPG, PNG, WEBP, BMP, TIFF</span>
        </div>
        <div className="text-[10px] font-mono opacity-50 uppercase font-bold">
          Q-Flow OCR v2.4.0-Stable | AI Engine: {mode === 'offline' ? 'Tesseract (Local)' : 'Qwen VL Max'}
        </div>
      </footer>
    </div>
  );
}
