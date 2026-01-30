import React, { useState, useMemo } from 'react';
import { 
  ClipboardCheck, BarChart3, Download, Users, Calendar, Plus, Trash2, PieChart as PieIcon, FileText,
  Building2, BadgeCheck, AlertCircle, CheckCircle2
} from 'lucide-react';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip, Legend } from 'recharts';
import { AppState, ComplianceStatus, FinalConclusion, ChecklistItem } from './types';
import { generateG3022Docx } from './G3022Generator';

interface Props {
  data: AppState['g3022'];
  onChange: (newData: AppState['g3022']) => void;
  onReset: () => void;
}

const COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899'];

const G3022Report: React.FC<Props> = ({ data, onChange, onReset }) => {
  const [activeTab, setActiveTab] = useState(0);
  const [isExporting, setIsExporting] = useState(false);

  const totalEmissions = useMemo(() => {
    const { cat1, cat2, cat3, cat4, cat5, cat6 } = data.emissions;
    return (cat1 || 0) + (cat2 || 0) + (cat3 || 0) + (cat4 || 0) + (cat5 || 0) + (cat6 || 0);
  }, [data.emissions]);

  const chartData = useMemo(() => [
    { name: 'Cat 1', value: data.emissions.cat1 },
    { name: 'Cat 2', value: data.emissions.cat2 },
    { name: 'Cat 3', value: data.emissions.cat3 },
    { name: 'Cat 4', value: data.emissions.cat4 },
    { name: 'Cat 5', value: data.emissions.cat5 },
    { name: 'Cat 6', value: data.emissions.cat6 },
  ].filter(i => i.value > 0), [data.emissions]);

  const updateBasic = (field: string, value: any) => {
    onChange({ ...data, basicInfo: { ...data.basicInfo, [field]: value } });
  };

  const toggleScope = (type: 'reasonable' | 'limited', scope: string) => {
    const key = type === 'reasonable' ? 'reasonableScopes' : 'limitedScopes';
    const current = data.basicInfo[key] || [];
    const next = current.includes(scope) 
      ? current.filter(s => s !== scope) 
      : [...current, scope];
    updateBasic(key, next);
  };

  const updateEmission = (field: string, value: any) => {
    if (['cat1', 'cat2', 'cat3', 'cat4', 'cat5', 'cat6'].includes(field)) {
       const num = parseFloat(value) || 0;
       onChange({ ...data, emissions: { ...data.emissions, [field]: num } });
    } else {
       onChange({ ...data, emissions: { ...data.emissions, [field]: value } });
    }
  };

  const updateChecklist = (idx: number, field: keyof ChecklistItem, value: any) => {
    const newList = [...data.checklist];
    newList[idx] = { ...newList[idx], [field]: value };
    onChange({ ...data, checklist: newList });
  };

  const setAllCompliant = () => {
    const newList = data.checklist.map(item => ({ ...item, status: ComplianceStatus.COMPLIANT }));
    onChange({ ...data, checklist: newList });
  };

  const updateConclusion = (field: string, value: any) => {
    onChange({ ...data, conclusion: { ...data.conclusion, [field]: value } });
  };

  const addInterview = () => {
    const newItem = { id: Date.now().toString(), topic: '', record: '', result: '' };
    onChange({ ...data, conclusion: { ...data.conclusion, interviews: [...data.conclusion.interviews, newItem] } });
  };

  const updateInterview = (index: number, field: string, value: string) => {
      const newInterviews = [...data.conclusion.interviews];
      // @ts-ignore
      newInterviews[index] = { ...newInterviews[index], [field]: value };
      onChange({ ...data, conclusion: { ...data.conclusion, interviews: newInterviews } });
  };
  
  const removeInterview = (index: number) => {
      const newInterviews = data.conclusion.interviews.filter((_, i) => i !== index);
      onChange({ ...data, conclusion: { ...data.conclusion, interviews: newInterviews } });
  };

  const addPending = () => {
    const newItem = { id: Date.now().toString(), content: '', response: '' };
    onChange({ ...data, conclusion: { ...data.conclusion, pendingItems: [...data.conclusion.pendingItems, newItem] } });
  };

  const updatePending = (index: number, field: string, value: string) => {
      const newPending = [...data.conclusion.pendingItems];
      // @ts-ignore
      newPending[index] = { ...newPending[index], [field]: value };
      onChange({ ...data, conclusion: { ...data.conclusion, pendingItems: newPending } });
  };

  const removePending = (index: number) => {
      const newPending = data.conclusion.pendingItems.filter((_, i) => i !== index);
      onChange({ ...data, conclusion: { ...data.conclusion, pendingItems: newPending } });
  };

  const handleExport = async () => {
    try {
      setIsExporting(true);
      const blob = await generateG3022Docx(data);
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `G-3022_Report_${data.basicInfo.caseNumber || 'Draft'}.docx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      setTimeout(() => URL.revokeObjectURL(url), 100);
    } catch (error) {
      console.error("Export failed:", error);
      alert("匯出失敗，請檢查資料是否完整");
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="space-y-6">
      {/* Action Bar */}
      <div className="flex justify-between items-center bg-white p-4 rounded-2xl shadow-sm border border-slate-100">
        <div className="flex gap-2 bg-slate-100 p-1 rounded-xl">
          {['基本資訊', '排放數據', '查驗清單', '總結與簽署'].map((label, idx) => (
            <button
              key={idx}
              onClick={() => setActiveTab(idx)}
              className={`px-6 py-2 rounded-lg text-sm font-bold transition-all ${
                activeTab === idx ? 'bg-white text-blue-700 shadow-sm' : 'text-slate-500 hover:text-slate-700'
              }`}
            >
              {label}
            </button>
          ))}
        </div>
        <div className="flex gap-3">
            <button onClick={onReset} className="flex items-center gap-2 text-slate-400 hover:text-red-500 px-4 py-2 rounded-xl font-bold transition-colors">
                <Trash2 size={16} /> 重置
            </button>
            <button 
              onClick={handleExport} 
              disabled={isExporting}
              className="flex items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-xl font-bold shadow-md shadow-blue-200 transition-all active:scale-95 disabled:opacity-50"
            >
                {isExporting ? '輸出中...' : '下載 DOCX'} <Download size={18} />
            </button>
        </div>
      </div>

      {/* Content */}
      {activeTab === 0 && (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 animate-in fade-in slide-in-from-bottom-2 duration-300">
           {/* Client Card */}
           <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-6">
               <div className="flex items-center gap-3 border-b border-slate-50 pb-4">
                  <div className="bg-indigo-50 p-2 rounded-xl text-indigo-600"><Users size={20}/></div>
                  <h3 className="font-black text-slate-800 text-lg">客戶與文件資訊</h3>
               </div>
               <div className="grid grid-cols-2 gap-4">
                   <div className="col-span-2"><Field label="委託單位名稱" value={data.basicInfo.clientName} onChange={(v: string) => updateBasic('clientName', v)} /></div>
                   <div className="col-span-2"><Field label="委託單位地址" value={data.basicInfo.clientAddress} onChange={(v: string) => updateBasic('clientAddress', v)} /></div>
                   <div className="col-span-2"><Field label="預期使用者" value={data.basicInfo.intendedUser} onChange={(v: string) => updateBasic('intendedUser', v)} /></div>
               </div>
               
               <div className="pt-4 border-t border-slate-50 space-y-4">
                 <div className="text-xs font-black text-slate-400 uppercase tracking-widest">依據文件</div>
                 <Field label="溫室氣體報告 (名稱/版次/日期)" value={data.basicInfo.reportName} onChange={(v: string) => updateBasic('reportName', v)} placeholder="例如：2023年度GHG報告書 V1.0" />
                 <Field label="盤查清冊 (名稱/版次/日期)" value={data.basicInfo.inventoryName} onChange={(v: string) => updateBasic('inventoryName', v)} placeholder="例如：2023盤查清冊 V1.0" />
                 <Field label="管理程序 (名稱/版次/日期)" value={data.basicInfo.procedureName} onChange={(v: string) => updateBasic('procedureName', v)} placeholder="例如：GHG管理程序 V2.0" />
               </div>
           </div>

           {/* Verification Params Card */}
           <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-6">
               <div className="flex items-center gap-3 border-b border-slate-50 pb-4">
                  <div className="bg-blue-50 p-2 rounded-xl text-blue-600"><ClipboardCheck size={20}/></div>
                  <h3 className="font-black text-slate-800 text-lg">查驗參數設定</h3>
               </div>
               
               <div className="grid grid-cols-2 gap-4">
                   <Field label="案件編號" value={data.basicInfo.caseNumber} onChange={(v: string) => updateBasic('caseNumber', v)} />
                   <Field label="查驗年度 (西元)" value={data.basicInfo.verificationYear} onChange={(v: string) => updateBasic('verificationYear', v)} />
                   <DateField label="書面審查日期" value={data.basicInfo.reviewDate} onChange={(v: string) => updateBasic('reviewDate', v)} />
                   <DateField label="赴廠訪談日期" value={data.basicInfo.visitDate} onChange={(v: string) => updateBasic('visitDate', v)} />
               </div>

               <div className="grid grid-cols-2 gap-4">
                   <Field label="基準年 (西元)" value={data.basicInfo.baseYear} onChange={(v: string) => updateBasic('baseYear', v)} />
                   <Field label="基準年排放量 (tCO2e)" type="number" value={data.basicInfo.baseYearEmissions} onChange={(v: string) => updateBasic('baseYearEmissions', v)} />
               </div>

               <div className="space-y-3 pt-2">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">保證等級與範疇</label>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                            <div className="text-xs font-bold text-slate-700 mb-2 flex items-center gap-1"><BadgeCheck size={14} className="text-emerald-500"/> 合理保證 (Reasonable)</div>
                            <div className="grid grid-cols-3 gap-2">
                                {['cat1', 'cat2', 'cat3', 'cat4', 'cat5', 'cat6'].map((cat, idx) => (
                                    <label key={`reasonable-${cat}`} className={`flex items-center justify-center gap-1 py-1.5 rounded-lg border cursor-pointer transition-all ${
                                        data.basicInfo.reasonableScopes?.includes(cat) 
                                        ? 'bg-emerald-500 text-white border-emerald-600 font-bold shadow-md shadow-emerald-200' 
                                        : 'bg-white border-slate-200 text-slate-400 hover:border-emerald-300'
                                    }`}>
                                        <input type="checkbox" className="hidden" checked={data.basicInfo.reasonableScopes?.includes(cat) || false} onChange={() => toggleScope('reasonable', cat)}/>
                                        <span className="text-xs">Cat {idx + 1}</span>
                                    </label>
                                ))}
                            </div>
                        </div>
                        <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                            <div className="text-xs font-bold text-slate-700 mb-2 flex items-center gap-1"><BadgeCheck size={14} className="text-blue-500"/> 有限保證 (Limited)</div>
                            <div className="grid grid-cols-3 gap-2">
                                {['cat1', 'cat2', 'cat3', 'cat4', 'cat5', 'cat6'].map((cat, idx) => (
                                    <label key={`limited-${cat}`} className={`flex items-center justify-center gap-1 py-1.5 rounded-lg border cursor-pointer transition-all ${
                                        data.basicInfo.limitedScopes?.includes(cat) 
                                        ? 'bg-blue-500 text-white border-blue-600 font-bold shadow-md shadow-blue-200' 
                                        : 'bg-white border-slate-200 text-slate-400 hover:border-blue-300'
                                    }`}>
                                        <input type="checkbox" className="hidden" checked={data.basicInfo.limitedScopes?.includes(cat) || false} onChange={() => toggleScope('limited', cat)}/>
                                        <span className="text-xs">Cat {idx + 1}</span>
                                    </label>
                                ))}
                            </div>
                        </div>
                    </div>
               </div>
               <Field label="實質性門檻說明" value={data.basicInfo.materiality} onChange={(v: string) => updateBasic('materiality', v)} />
           </div>
        </div>
      )}

      {activeTab === 1 && (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 animate-in fade-in slide-in-from-bottom-2 duration-300">
           <div className="md:col-span-1 space-y-6">
              <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm">
                  <h3 className="font-black text-slate-800 mb-6 flex items-center gap-2 text-lg"><Building2 size={20} className="text-slate-500"/> 類別排放量 (tCO2e)</h3>
                  <div className="space-y-4">
                    {['cat1', 'cat2', 'cat3', 'cat4', 'cat5', 'cat6'].map((cat, i) => (
                        <div key={cat} className="flex items-center gap-3">
                            <div className={`w-8 h-8 rounded-lg flex items-center justify-center text-xs font-bold text-white shrink-0`} style={{ backgroundColor: COLORS[i] }}>
                                C{i+1}
                            </div>
                            <Field 
                                type="number" 
                                value={data.emissions[cat as keyof typeof data.emissions]} 
                                onChange={(v: string) => updateEmission(cat, v)} 
                                placeholder="0.00"
                            />
                        </div>
                    ))}
                  </div>
                  <div className="mt-6 pt-6 border-t border-slate-100">
                      <div className="flex justify-between items-center bg-slate-50 p-4 rounded-xl border border-slate-200">
                          <span className="font-bold text-slate-500 text-xs uppercase tracking-widest">Total Emissions</span>
                          <span className="font-black text-2xl text-slate-800">{totalEmissions.toFixed(2)}</span>
                      </div>
                  </div>
              </div>
              
              <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-4">
                  <h3 className="font-black text-slate-800 flex items-center gap-2"><BarChart3 size={20} className="text-slate-500"/> 不確定性評估</h3>
                  <div className="grid grid-cols-2 gap-4">
                    <Field label="上限 (%)" value={data.emissions.uncertaintyUpper} onChange={(v: string) => updateEmission('uncertaintyUpper', v)} />
                    <Field label="下限 (%)" value={data.emissions.uncertaintyLower} onChange={(v: string) => updateEmission('uncertaintyLower', v)} />
                  </div>
              </div>
           </div>

           <div className="md:col-span-2 bg-white rounded-3xl border border-slate-100 shadow-sm p-6 flex flex-col">
              <h3 className="font-black text-slate-800 flex items-center gap-2 mb-4"><PieIcon size={20} className="text-slate-500"/> 排放占比分析</h3>
              <div className="flex-1 min-h-[400px] flex items-center justify-center bg-slate-50/50 rounded-2xl border border-slate-100">
                  {totalEmissions > 0 ? (
                    <ResponsiveContainer width="100%" height={400}>
                      <PieChart>
                        <Pie
                          data={chartData}
                          cx="50%"
                          cy="50%"
                          innerRadius={100}
                          outerRadius={140}
                          paddingAngle={5}
                          dataKey="value"
                          cornerRadius={6}
                        >
                          {chartData.map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} strokeWidth={0} />
                          ))}
                        </Pie>
                        <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
                        <Legend iconType="circle" />
                      </PieChart>
                    </ResponsiveContainer>
                  ) : (
                    <div className="text-slate-300 font-bold flex flex-col items-center gap-3">
                       <BarChart3 size={64} strokeWidth={1.5} />
                       <span>請輸入排放數據以產生圖表</span>
                    </div>
                  )}
              </div>
           </div>
        </div>
      )}

      {activeTab === 2 && (
         <div className="animate-in fade-in slide-in-from-bottom-2 duration-300">
            <div className="bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
                <div className="p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                   <h3 className="font-black text-slate-800 text-lg">查驗清單確認 (ISO 14064-1:2018)</h3>
                   <button onClick={setAllCompliant} className="text-xs font-bold text-emerald-600 hover:text-emerald-700 bg-emerald-50 hover:bg-emerald-100 px-4 py-2 rounded-xl transition-colors border border-emerald-100">
                      全部設為符合
                   </button>
                </div>
                <div className="overflow-x-auto">
                    <table className="w-full text-left text-sm">
                    <thead className="bg-slate-50 border-b border-slate-200 text-slate-500">
                        <tr>
                        <th className="p-4 w-24 text-center font-bold">編號</th>
                        <th className="p-4 font-bold">查驗項目</th>
                        <th className="p-4 w-1/3 font-bold">相關文件/證據</th>
                        <th className="p-4 w-40 text-center font-bold">狀態</th>
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                        {data.checklist.map((item, idx) => {
                        const isHeader = !item.id.includes('.');
                        return (
                        <tr key={item.id} className={`transition-colors ${isHeader ? 'bg-slate-50/80' : 'hover:bg-blue-50/30'}`}>
                            <td className={`p-4 text-center ${isHeader ? 'font-black text-slate-800' : 'font-bold text-slate-400'}`}>{item.id}</td>
                            <td 
                                className={`p-4 ${isHeader ? 'font-black text-slate-800' : 'text-slate-700 font-medium'}`}
                                colSpan={isHeader ? 3 : 1}
                            >
                                {item.name}
                            </td>
                            {!isHeader && (
                                <>
                                    <td className="p-4">
                                        <input
                                            className="w-full bg-slate-100 hover:bg-white focus:bg-white rounded-lg px-3 py-2 text-xs border border-transparent focus:border-blue-500 outline-none transition-all"
                                            placeholder="輸入文件編號..."
                                            value={item.docRef}
                                            onChange={(e) => updateChecklist(idx, 'docRef', e.target.value)}
                                        />
                                    </td>
                                    <td className="p-4 text-center">
                                        <select 
                                            className={`w-full p-2 rounded-lg text-xs font-bold appearance-none cursor-pointer outline-none border transition-all text-center ${
                                                item.status === ComplianceStatus.COMPLIANT ? 'bg-emerald-50 text-emerald-600 border-emerald-200' : 
                                                item.status === ComplianceStatus.CLARIFY ? 'bg-amber-50 text-amber-600 border-amber-200' : 'bg-slate-100 text-slate-500 border-slate-200'
                                            }`}
                                            value={item.status}
                                            onChange={(e) => updateChecklist(idx, 'status', e.target.value)}
                                        >
                                            <option value={ComplianceStatus.COMPLIANT}>符合</option>
                                            <option value={ComplianceStatus.CLARIFY}>待釐清</option>
                                            <option value={ComplianceStatus.NA}>不適用</option>
                                        </select>
                                    </td>
                                </>
                            )}
                        </tr>
                        )})}
                    </tbody>
                    </table>
                </div>
            </div>
         </div>
      )}

      {activeTab === 3 && (
        <div className="space-y-6 animate-in fade-in slide-in-from-bottom-2 duration-300">
            {/* Interviews */}
            <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-4">
               <div className="flex justify-between items-center border-b border-slate-50 pb-4">
                  <h3 className="font-black text-slate-800 text-lg">訪談紀錄</h3>
                  <button onClick={addInterview} className="bg-slate-50 hover:bg-slate-100 text-slate-600 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1 border border-slate-200 transition-colors"><Plus size={14}/> 新增</button>
               </div>
               {data.conclusion.interviews.length === 0 && <div className="text-center py-8 text-slate-400 font-medium bg-slate-50/50 rounded-2xl border border-dashed border-slate-200">尚無訪談紀錄</div>}
               {data.conclusion.interviews.map((iv, i) => (
                  <div key={iv.id} className="grid grid-cols-1 md:grid-cols-12 gap-4 p-5 bg-slate-50/50 rounded-2xl border border-slate-100 relative group hover:shadow-md transition-all">
                     <div className="md:col-span-3"><Field label="訪談對象/主題" value={iv.topic} onChange={(v: string) => updateInterview(i, 'topic', v)} /></div>
                     <div className="md:col-span-6"><Field label="訪談內容概要" value={iv.record} onChange={(v: string) => updateInterview(i, 'record', v)} /></div>
                     <div className="md:col-span-3"><Field label="訪談結果" value={iv.result} onChange={(v: string) => updateInterview(i, 'result', v)} /></div>
                     <button onClick={() => removeInterview(i)} className="absolute top-2 right-2 p-1.5 text-slate-300 hover:text-red-500 hover:bg-red-50 rounded-lg transition-all opacity-0 group-hover:opacity-100"><Trash2 size={16}/></button>
                  </div>
               ))}
            </div>

            {/* Pending Items */}
            <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-4">
               <div className="flex justify-between items-center border-b border-slate-50 pb-4">
                  <h3 className="font-black text-amber-600 text-lg flex items-center gap-2"><AlertCircle size={20}/> 待釐清/待補正事項</h3>
                  <button onClick={addPending} className="bg-amber-50 hover:bg-amber-100 text-amber-600 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1 border border-amber-100 transition-colors"><Plus size={14}/> 新增</button>
               </div>
               {data.conclusion.pendingItems.length === 0 && <div className="text-center py-8 text-slate-400 font-medium bg-slate-50/50 rounded-2xl border border-dashed border-slate-200">尚無待釐清事項</div>}
               {data.conclusion.pendingItems.map((pd, i) => (
                  <div key={pd.id} className="grid grid-cols-1 md:grid-cols-2 gap-6 p-5 bg-amber-50/30 rounded-2xl border border-amber-100 relative group hover:shadow-md transition-all">
                     <Field label="發現事項內容" value={pd.content} onChange={(v: string) => updatePending(i, 'content', v)} />
                     <Field label="回覆與處理" value={pd.response} onChange={(v: string) => updatePending(i, 'response', v)} />
                     <button onClick={() => removePending(i)} className="absolute top-2 right-2 p-1.5 text-amber-300 hover:text-red-500 hover:bg-red-50 rounded-lg transition-all opacity-0 group-hover:opacity-100"><Trash2 size={16}/></button>
                  </div>
               ))}
            </div>

            {/* Conclusion & Signatures */}
            <div className="bg-white border border-slate-200 p-8 rounded-3xl shadow-sm space-y-8">
                <div className="flex items-center gap-3 border-b border-slate-100 pb-4">
                   <div className="bg-blue-600 p-2 rounded-xl text-white"><FileText size={20}/></div>
                   <h3 className="font-black text-xl text-slate-800">查驗結論與簽署</h3>
                </div>
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                   <div className="space-y-4">
                        <label className="text-xs font-black text-slate-400 uppercase tracking-widest block">利益衝突迴避</label>
                        <div className="flex gap-4">
                            <label className={`flex items-center gap-2 px-4 py-3 rounded-xl border cursor-pointer transition-all ${data.conclusion.conflictOfInterest === 'No' ? 'bg-emerald-50 text-emerald-600 border-emerald-200' : 'bg-slate-50 text-slate-400 border-slate-200'}`}>
                                <input type="radio" className="hidden" checked={data.conclusion.conflictOfInterest === 'No'} onChange={() => updateConclusion('conflictOfInterest', 'No')} />
                                <CheckCircle2 size={18}/> <span className="font-bold text-sm">無利益衝突</span>
                            </label>
                            <label className={`flex items-center gap-2 px-4 py-3 rounded-xl border cursor-pointer transition-all ${data.conclusion.conflictOfInterest === 'Yes' ? 'bg-red-50 text-red-600 border-red-200' : 'bg-slate-50 text-slate-400 border-slate-200'}`}>
                                <input type="radio" className="hidden" checked={data.conclusion.conflictOfInterest === 'Yes'} onChange={() => updateConclusion('conflictOfInterest', 'Yes')} />
                                <AlertCircle size={18}/> <span className="font-bold text-sm">有利益衝突</span>
                            </label>
                        </div>
                        {data.conclusion.conflictOfInterest === 'Yes' && (
                            <Field label="衝突說明" value={data.conclusion.conflictDetail} onChange={(v: string) => updateConclusion('conflictDetail', v)} />
                        )}
                   </div>
                   
                   <div className="space-y-4">
                       <label className="text-xs font-black text-slate-400 uppercase tracking-widest block">查驗總結</label>
                       <select 
                          className="w-full p-4 bg-slate-50 border border-slate-200 rounded-2xl outline-none font-bold text-sm text-slate-800 focus:border-blue-500 transition-colors appearance-none"
                          value={data.conclusion.summary}
                          onChange={(e) => updateConclusion('summary', e.target.value)}
                       >
                          {Object.values(FinalConclusion).map(s => <option key={s} value={s}>{s}</option>)}
                       </select>
                   </div>
                </div>

                <div className="space-y-2">
                    <label className="text-xs font-black text-slate-400 uppercase tracking-widest block">其它補充說明</label>
                    <textarea 
                        className="w-full p-4 bg-slate-50 border border-slate-200 rounded-2xl focus:border-blue-500 outline-none font-bold text-sm text-slate-700 h-24 resize-none transition-colors"
                        placeholder="無"
                        value={data.conclusion.otherNote}
                        onChange={(e) => updateConclusion('otherNote', e.target.value)}
                    />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-6 pt-6 border-t border-slate-100">
                    <Field label="主導查驗員" value={data.conclusion.leadVerifierName} onChange={(v: string) => updateConclusion('leadVerifierName', v)} />
                    <Field label="查驗員" value={data.conclusion.verifierName} onChange={(v: string) => updateConclusion('verifierName', v)} />
                    <Field label="客戶代表" value={data.conclusion.clientRepName} onChange={(v: string) => updateConclusion('clientRepName', v)} />
                </div>

                <label className="flex items-start gap-3 cursor-pointer p-4 rounded-xl bg-slate-50/50 hover:bg-slate-50 transition-colors border border-slate-200">
                    <input 
                        type="checkbox" 
                        className="mt-1 w-5 h-5 accent-blue-500 rounded"
                        checked={data.conclusion.memoCorrection}
                        onChange={(e) => updateConclusion('memoCorrection', e.target.checked)}
                    />
                    <span className="text-xs font-bold text-slate-500 leading-relaxed">
                        備註：書面審查及/或訪談結果，尚有上述事項待釐清/補正，請於第一階段(S-1)前說明或提送補正資料。
                    </span>
                </label>
            </div>
        </div>
      )}
    </div>
  );
};

const Field = ({ label, placeholder, value, onChange, type = "text" }: any) => {
  return (
    <div className="space-y-1.5 w-full group">
      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1 group-focus-within:text-blue-500 transition-colors">{label}</label>
      {type === 'text' || type === 'number' ? (
        <input 
            type={type}
            className="w-full p-3.5 bg-slate-50 border border-slate-200 rounded-2xl focus:bg-white focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 outline-none transition-all font-bold text-sm text-slate-700 placeholder:text-slate-300" 
            placeholder={placeholder} 
            value={value} 
            onChange={(e) => onChange(e.target.value)} 
        />
      ) : (
        <textarea 
            rows={1}
            className="w-full p-3.5 bg-slate-50 border border-slate-200 rounded-2xl focus:bg-white focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 outline-none transition-all font-bold text-sm text-slate-700 placeholder:text-slate-300 min-h-[48px]" 
            placeholder={placeholder} 
            value={value} 
            onChange={(e) => onChange(e.target.value)} 
        />
      )}
    </div>
  );
};

const DateField = ({ label, value, onChange }: any) => (
  <div className="space-y-1.5 w-full group">
    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1 group-focus-within:text-blue-500 transition-colors">{label}</label>
    <div className="relative">
      <Calendar className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" size={18} />
      <input 
        type="date" 
        className="w-full pl-12 pr-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl focus:bg-white focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 outline-none text-sm font-bold text-slate-700 transition-all uppercase" 
        value={value} 
        onChange={(e) => onChange(e.target.value)} 
      />
    </div>
  </div>
);

export default G3022Report;