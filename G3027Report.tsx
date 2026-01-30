import React, { useState, useEffect } from 'react';
import { 
    Download, Plus, Trash2, Calendar, FileText, CheckSquare, 
    BarChart3, AlertTriangle, ChevronDown, ChevronUp, Lock, ArrowRightCircle, CheckCircle2
} from 'lucide-react';
import { AppState, FindingItem } from './types';
import { generateG3027Docx } from './G3027Generator';

interface Props {
  data: AppState['g3027'];
  onChange: (newData: AppState['g3027']) => void;
  onReset: () => void;
}

// ════════════════════════════════════════════════════════════════════════════
// SUB-COMPONENTS (Defined OUTSIDE to prevent re-mount/focus loss)
// ════════════════════════════════════════════════════════════════════════════

const Field = ({ label, placeholder, value, onChange, type = "text", disabled = false }: any) => (
  <div className="space-y-1.5 w-full group">
    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1 group-focus-within:text-blue-500 transition-colors">{label}</label>
    <input 
      type={type} 
      className={`w-full p-3.5 border rounded-2xl outline-none transition-all font-bold text-sm ${disabled ? 'bg-slate-100 text-slate-500 border-slate-200' : 'bg-slate-50 border-slate-200 text-slate-700 focus:bg-white focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500'}`}
      placeholder={placeholder} 
      value={value} 
      onChange={(e) => onChange(e.target.value)} 
      disabled={disabled}
    />
  </div>
);

const TextArea = ({ label, placeholder, value, onChange, disabled = false }: any) => (
  <div className="space-y-1.5 w-full group">
    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1 group-focus-within:text-blue-500 transition-colors">{label}</label>
    <textarea
      rows={2}
      className={`w-full p-3.5 border rounded-2xl outline-none transition-all font-bold text-sm resize-none ${disabled ? 'bg-slate-100 text-slate-500 border-slate-200' : 'bg-slate-50 border-slate-200 text-slate-700 focus:bg-white focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500'}`}
      placeholder={placeholder} 
      value={value} 
      onChange={(e) => onChange(e.target.value)} 
      disabled={disabled}
    />
  </div>
);

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

// Extracted FindingsList Component
const FindingsList = ({ 
    stage, 
    readOnly, 
    findings, 
    expandedIds, 
    onToggleExpand, 
    onToggleAll, 
    onAdd, 
    onRemove, 
    onUpdate 
}: {
    stage: 'S1' | 'S2';
    readOnly: boolean;
    findings: FindingItem[];
    expandedIds: Set<string>;
    onToggleExpand: (id: string) => void;
    onToggleAll: (stage: 'S1' | 'S2') => void;
    onAdd: (stage: 'S1' | 'S2') => void;
    onRemove: (id: string) => void;
    onUpdate: (id: string, field: string, value: any) => void;
}) => {
    const items = findings.filter(f => f.stage === stage);
    
    return (
      <div className="space-y-6">
          <div className="flex justify-between items-center px-2">
             <div className="flex items-center gap-3">
                 <h3 className="font-black text-slate-800 text-lg flex items-center gap-2">
                     <AlertTriangle size={20} className={stage === 'S1' ? "text-blue-500" : "text-emerald-500"}/> 
                     {stage === 'S1' ? '第一階段 (S1) 發現事項' : '第二階段 (S2) 發現事項'}
                 </h3>
                 <span className="bg-slate-100 text-slate-500 text-xs px-2 py-1 rounded-md font-bold">{items.length} 筆</span>
             </div>
             <div className="flex gap-2">
                 <button onClick={() => onToggleAll(stage)} className="bg-white border border-slate-200 text-slate-500 hover:text-blue-600 px-3 py-2 rounded-xl text-xs font-bold transition-all shadow-sm">
                     {items.length > 0 && items.every(f => expandedIds.has(f.id)) ? '全部收合' : '全部展開'}
                 </button>
                 {!readOnly && (
                     <button onClick={() => onAdd(stage)} className="bg-blue-600 text-white px-4 py-2 rounded-xl text-xs font-bold hover:bg-blue-700 flex items-center gap-1 shadow-md shadow-blue-200 transition-all active:scale-95">
                         <Plus size={14}/> 新增 {stage} 事項
                     </button>
                 )}
             </div>
          </div>

          {items.length === 0 && (
              <div className="p-12 text-center text-slate-400 font-medium bg-slate-50 rounded-3xl border border-dashed border-slate-200">
                  {readOnly 
                      ? `尚無 ${stage} 階段的發現事項。` 
                      : stage === 'S2' 
                          ? "尚無 S2 事項。若 S1 有「保留」項目，將自動帶入。" 
                          : "尚無發現事項，請點擊上方按鈕新增。"}
              </div>
          )}

          <div className="space-y-3">
          {items.map((item, idx) => {
              const isExpanded = expandedIds.has(item.id);
              return (
              <div key={item.id} className={`bg-white border transition-all duration-300 rounded-3xl overflow-hidden ${isExpanded ? 'border-blue-200 shadow-md ring-1 ring-blue-50' : 'border-slate-200 shadow-sm hover:border-blue-200'}`}>
                  {/* Header Row */}
                  <div 
                      onClick={() => onToggleExpand(item.id)}
                      className={`flex items-center gap-4 p-4 cursor-pointer transition-colors ${readOnly ? 'bg-slate-50 hover:bg-slate-100' : 'bg-white hover:bg-slate-50/50'}`}
                  >
                      <div className={`w-8 h-8 rounded-lg flex items-center justify-center font-black text-sm shrink-0 border ${isExpanded ? 'bg-blue-600 text-white border-blue-600' : 'bg-slate-100 text-slate-400 border-slate-200'}`}>
                          {idx + 1}
                      </div>
                      
                      <div className="w-24 shrink-0">
                          {item.type ? (
                              <span className={`inline-block px-2 py-1 rounded-md text-xs font-bold border ${
                                  item.type === 'CAR' ? 'bg-red-50 text-red-600 border-red-100' :
                                  item.type === 'OBS' ? 'bg-amber-50 text-amber-600 border-amber-100' :
                                  'bg-blue-50 text-blue-600 border-blue-100'
                              }`}>
                                  {item.type}
                              </span>
                          ) : (
                              <span className="text-slate-300 text-xs font-bold italic">未分類</span>
                          )}
                      </div>

                      <div className="flex-1 min-w-0">
                          <p className="text-sm font-bold text-slate-700 truncate">
                              {item.description || <span className="text-slate-300 font-normal italic">請輸入不符合事項描述...</span>}
                          </p>
                      </div>

                      {readOnly && <Lock size={14} className="text-slate-300"/>}

                      <div className="text-slate-400 shrink-0">
                          {isExpanded ? <ChevronUp size={18} /> : <ChevronDown size={18} />}
                      </div>
                  </div>

                  {/* Expandable Content */}
                  {isExpanded && (
                      <div className="p-6 pt-2 border-t border-slate-100 space-y-6 bg-slate-50/30">
                          {!readOnly && (
                              <div className="flex justify-end">
                                  <button onClick={(e) => { e.stopPropagation(); onRemove(item.id); }} className="flex items-center gap-1 text-xs font-bold text-red-400 hover:text-red-600 px-3 py-1.5 hover:bg-red-50 rounded-lg transition-all">
                                      <Trash2 size={14}/> 刪除此項目
                                  </button>
                              </div>
                          )}

                          <div className="grid grid-cols-1 md:grid-cols-12 gap-6">
                              <div className="md:col-span-4 space-y-6">
                                  <div className="space-y-1.5 group">
                                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1 flex items-center gap-1">分類</label>
                                      <select 
                                          className={`w-full p-3.5 border rounded-2xl font-bold text-sm outline-none transition-all appearance-none ${readOnly ? 'bg-slate-100 text-slate-500' : 'bg-white text-slate-700 focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500'}`}
                                          value={item.type}
                                          onChange={(e) => onUpdate(item.id, 'type', e.target.value)}
                                          disabled={readOnly}
                                      >
                                          <option value="">請選擇</option>
                                          <option value="CAR">CAR (矯正措施)</option>
                                          <option value="CR">CR (澄清要求)</option>
                                          <option value="OBS">OBS (觀察事項)</option>
                                          <option value="FAR">FAR (後續行動)</option>
                                      </select>
                                  </div>
                                  <Field label="填報查驗人員" value={item.reporter} onChange={(v: string) => onUpdate(item.id, 'reporter', v)} disabled={readOnly} />
                              </div>

                              <div className="md:col-span-8 space-y-4">
                                  <TextArea label="不符合事項描述" value={item.description} onChange={(v: string) => onUpdate(item.id, 'description', v)} disabled={readOnly} />
                                  <TextArea label="矯正措施/澄清說明 (受查方)" value={item.correctiveAction} onChange={(v: string) => onUpdate(item.id, 'correctiveAction', v)} disabled={readOnly} />
                                  <TextArea label="審查意見 (查驗方)" value={item.reviewOpinion} onChange={(v: string) => onUpdate(item.id, 'reviewOpinion', v)} disabled={readOnly} />
                              </div>
                          </div>

                          <div className="grid grid-cols-1 md:grid-cols-3 gap-6 bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                              <Field label="查驗員" value={item.reviewer} onChange={(v: string) => onUpdate(item.id, 'reviewer', v)} disabled={readOnly} />
                              <div className="space-y-1.5 group">
                                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1 group-focus-within:text-blue-500">審查結果</label>
                                  <select 
                                      className={`w-full p-3.5 border rounded-2xl font-bold text-sm outline-none transition-all appearance-none ${readOnly ? 'bg-slate-100 text-slate-500' : 'bg-white text-slate-700 focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500'}`}
                                      value={item.result}
                                      onChange={(e) => onUpdate(item.id, 'result', e.target.value)}
                                      disabled={readOnly}
                                  >
                                      <option value="">請選擇</option>
                                      <option value="Close">結案</option>
                                      <option value="Keep">保留</option>
                                  </select>
                              </div>
                              <div className="space-y-1.5 group">
                                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1 group-focus-within:text-blue-500">審查地點</label>
                                  <select 
                                      className={`w-full p-3.5 border rounded-2xl font-bold text-sm outline-none transition-all appearance-none ${readOnly ? 'bg-slate-100 text-slate-500' : 'bg-white text-slate-700 focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500'}`}
                                      value={item.location}
                                      onChange={(e) => onUpdate(item.id, 'location', e.target.value)}
                                      disabled={readOnly}
                                  >
                                      <option value="">請選擇</option>
                                      <option value="OnSite">現場</option>
                                      <option value="OffSite">非現場</option>
                                  </select>
                              </div>
                          </div>
                      </div>
                  )}
              </div>
          )})}
          </div>
      </div>
    );
};

// ════════════════════════════════════════════════════════════════════════════
// MAIN COMPONENT
// ════════════════════════════════════════════════════════════════════════════

const G3027Report: React.FC<Props> = ({ data, onChange, onReset }) => {
  const [activeTab, setActiveTab] = useState(0);
  const [isExporting, setIsExporting] = useState(false);
  const [expandedIds, setExpandedIds] = useState<Set<string>>(new Set());

  const currentStage = data.basicInfo.stage as 'S1' | 'S2';

  // --- Auto-Sync Logic: S1 Keep -> S2 ---
  // Improved: Only runs when stage or tab changes to prevent editing loops
  useEffect(() => {
    if (currentStage === 'S2') {
        const s1Keeps = data.findings.filter(f => f.stage === 'S1' && f.result === 'Keep');
        const s2Items = data.findings.filter(f => f.stage === 'S2');
        
        let hasChanges = false;
        const newFindings = [...data.findings];

        s1Keeps.forEach(s1Item => {
            // Simple check by description to avoid duplication
            const exists = s2Items.some(s2 => s2.description === s1Item.description);
            if (!exists) {
                const newItem: FindingItem = {
                    ...s1Item,
                    id: Date.now().toString() + Math.random().toString().slice(2, 5),
                    stage: 'S2',
                    result: '', 
                    reviewOpinion: '', 
                    reviewer: '',      
                };
                newFindings.push(newItem);
                hasChanges = true;
            }
        });

        if (hasChanges) {
            onChange({ ...data, findings: newFindings });
        }
    }
  }, [currentStage, activeTab]); 

  // --- Auto-Calculate Statistics ---
  useEffect(() => {
    const newStats = {
        s1: { nonConformity: '0', observation: '0', suggestion: '0' },
        s2: { nonConformity: '0', observation: '0', suggestion: '0' }
    };

    data.findings.forEach(f => {
        const target = f.stage === 'S1' ? newStats.s1 : newStats.s2;
        if (f.type === 'CAR') target.nonConformity = (parseInt(target.nonConformity) + 1).toString();
        if (f.type === 'OBS' || f.type === 'CR') target.observation = (parseInt(target.observation) + 1).toString(); 
        if (f.type === 'FAR') target.suggestion = (parseInt(target.suggestion) + 1).toString();
    });

    if (JSON.stringify(newStats) !== JSON.stringify(data.stats)) {
        onChange({ ...data, stats: newStats });
    }
  }, [data.findings]);

  const toggleExpand = (id: string) => {
      const newSet = new Set(expandedIds);
      if (newSet.has(id)) newSet.delete(id);
      else newSet.add(id);
      setExpandedIds(newSet);
  };

  const toggleAll = (stageFilter: 'S1' | 'S2') => {
      const stageIds = data.findings.filter(f => f.stage === stageFilter).map(f => f.id);
      const allExpanded = stageIds.every(id => expandedIds.has(id));
      
      const newSet = new Set(expandedIds);
      if (allExpanded) {
          stageIds.forEach(id => newSet.delete(id));
      } else {
          stageIds.forEach(id => newSet.add(id));
      }
      setExpandedIds(newSet);
  };

  const updateBasic = (field: string, value: any) => {
    onChange({ ...data, basicInfo: { ...data.basicInfo, [field]: value } });
  };

  const updateConclusion = (field: string, value: any) => {
      onChange({ ...data, conclusion: { ...data.conclusion, [field]: value } });
  };

  const addFinding = (targetStage: 'S1' | 'S2') => {
    const newItem: FindingItem = { 
        id: Date.now().toString(), 
        stage: targetStage, 
        type: '', 
        description: '', 
        reporter: '', 
        correctiveAction: '', 
        reviewOpinion: '', 
        reviewer: '', 
        result: '', 
        location: '' 
    };
    onChange({ ...data, findings: [...data.findings, newItem] });
    setExpandedIds(prev => new Set(prev).add(newItem.id));
  };

  const updateFinding = (id: string, field: string, value: any) => {
      const newFindings = data.findings.map(f => f.id === id ? { ...f, [field]: value } : f);
      onChange({ ...data, findings: newFindings });
  };

  const removeFinding = (id: string) => {
      if (window.confirm('確定要刪除此發現事項嗎？')) {
        const newFindings = data.findings.filter(f => f.id !== id);
        onChange({ ...data, findings: newFindings });
      }
  };

  const handleExport = async () => {
    try {
      setIsExporting(true);
      const blob = await generateG3027Docx(data);
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `G-3027_Report_${data.basicInfo.caseNumber || 'Draft'}_${currentStage}.docx`;
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
          {['基本資訊', '第一階段 (S1)', '第二階段 (S2)', '統計與結論'].map((label, idx) => (
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

      {activeTab === 0 && (
        <div className="space-y-8 animate-in fade-in slide-in-from-bottom-2 duration-300">
           <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
             <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-4">
                <div className="flex items-center gap-3 border-b border-slate-50 pb-4">
                   <div className="bg-blue-50 p-2 rounded-xl text-blue-600"><FileText size={20}/></div>
                   <h3 className="font-black text-slate-800 text-lg">案件資訊</h3>
                </div>
                <Field label="案件編號" value={data.basicInfo.caseNumber} onChange={(v: string) => updateBasic('caseNumber', v)} />
                <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 ring-4 ring-slate-100">
                   <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-2 flex justify-between">
                       <span>目前查驗階段</span>
                       <span className="text-blue-600">影響資料鎖定與報告輸出</span>
                   </label>
                   <div className="flex gap-2">
                      {['S1', 'S2'].map(s => (
                         <button 
                            key={s} 
                            onClick={() => updateBasic('stage', s)} 
                            className={`flex-1 py-3 rounded-xl font-black text-base transition-all flex items-center justify-center gap-2 ${
                                data.basicInfo.stage === s 
                                ? 'bg-slate-800 text-white shadow-lg shadow-slate-300 scale-100' 
                                : 'bg-white border border-slate-200 text-slate-400 hover:bg-slate-50 scale-95'
                            }`}
                         >
                            {s === 'S1' ? '第一階段 (S1)' : '第二階段 (S2)'}
                            {data.basicInfo.stage === s && <CheckCircle2 size={18} />}
                         </button>
                      ))}
                   </div>
                   {data.basicInfo.stage === 'S2' && (
                       <div className="mt-3 text-[10px] text-amber-600 font-bold bg-amber-50 p-2 rounded-lg text-center flex flex-col gap-1">
                           <span className="flex items-center justify-center gap-1"><Lock size={10}/> S1 發現事項已鎖定</span>
                           <span className="flex items-center justify-center gap-1"><ArrowRightCircle size={10}/> S1「保留」項目已自動帶入 S2 列表</span>
                       </div>
                   )}
                </div>
                <Field label="查驗年度 (民國)" value={data.basicInfo.verificationYear} onChange={(v: string) => updateBasic('verificationYear', v)} />
             </div>
             <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-4">
                <div className="flex items-center gap-3 border-b border-slate-50 pb-4">
                   <div className="bg-indigo-50 p-2 rounded-xl text-indigo-600"><CheckSquare size={20}/></div>
                   <h3 className="font-black text-slate-800 text-lg">人員與日期</h3>
                </div>
                <Field label="主導查驗員" value={data.basicInfo.leadVerifier} onChange={(v: string) => updateBasic('leadVerifier', v)} />
                <Field label="受查驗方代表" value={data.basicInfo.auditeeRep} onChange={(v: string) => updateBasic('auditeeRep', v)} />
                <DateField label="查驗日期" value={data.basicInfo.date} onChange={(v: string) => updateBasic('date', v)} />
             </div>
           </div>
        </div>
      )}

      {activeTab === 1 && (
        <div className="animate-in fade-in slide-in-from-bottom-2 duration-300">
            {/* S1 List: ReadOnly if stage is S2 */}
            <FindingsList 
                stage="S1" 
                readOnly={currentStage === 'S2'} 
                findings={data.findings}
                expandedIds={expandedIds}
                onToggleExpand={toggleExpand}
                onToggleAll={toggleAll}
                onAdd={addFinding}
                onRemove={removeFinding}
                onUpdate={updateFinding}
            />
        </div>
      )}

      {activeTab === 2 && (
        <div className="animate-in fade-in slide-in-from-bottom-2 duration-300">
            {/* S2 List: ReadOnly if stage is S1 (Future stage shouldn't be edited yet) */}
            <FindingsList 
                stage="S2" 
                readOnly={currentStage === 'S1'} 
                findings={data.findings}
                expandedIds={expandedIds}
                onToggleExpand={toggleExpand}
                onToggleAll={toggleAll}
                onAdd={addFinding}
                onRemove={removeFinding}
                onUpdate={updateFinding}
            />
        </div>
      )}

      {activeTab === 3 && (
          <div className="space-y-8 animate-in fade-in slide-in-from-bottom-2 duration-300">
             {/* Stats Section */}
             <div className="bg-white p-8 rounded-3xl border border-slate-100 shadow-sm">
                 <h3 className="font-black text-slate-800 mb-6 flex items-center gap-2 text-lg"><BarChart3 size={20} className="text-slate-500"/> 數量統計 (自動計算)</h3>
                 <div className="grid grid-cols-3 gap-6 mb-4 text-center">
                    <div className="text-xs font-black text-slate-400 uppercase tracking-widest">類型</div>
                    <div className="text-xs font-black text-slate-400 uppercase tracking-widest">第一階段 (S-1)</div>
                    <div className="text-xs font-black text-slate-400 uppercase tracking-widest">第二階段 (S-2)</div>
                 </div>
                 {['不符合事項 (CAR)', '觀察事項 (OBS/CR)', '建議事項 (FAR)'].map((type, i) => {
                     const keyMap = ['nonConformity', 'observation', 'suggestion'];
                     const key = keyMap[i];
                     return (
                        <div key={key} className="grid grid-cols-3 gap-6 mb-4 items-center">
                            <div className="text-sm font-bold text-slate-700 text-center">{type}</div>
                            <input 
                                className="w-full p-3 text-center rounded-xl border border-slate-200 font-bold bg-slate-50 focus:bg-white text-slate-800 focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 outline-none transition-all"
                                // @ts-ignore
                                value={data.stats.s1[key]}
                                readOnly
                            />
                            <input 
                                className="w-full p-3 text-center rounded-xl border border-slate-200 font-bold bg-slate-50 focus:bg-white text-slate-800 focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 outline-none transition-all"
                                // @ts-ignore
                                value={data.stats.s2[key]}
                                readOnly
                            />
                        </div>
                     );
                 })}
             </div>

             {/* Conclusion Form */}
             <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                 <div className="bg-white p-6 rounded-3xl border border-blue-100 shadow-sm space-y-4 ring-1 ring-blue-50">
                     <h3 className="font-black text-blue-600 text-lg">第一階段 (S-1) 結論</h3>
                     <select 
                        className="w-full p-4 bg-blue-50/50 border border-blue-200 rounded-2xl font-bold text-sm text-blue-900 focus:bg-white outline-none focus:ring-4 focus:ring-blue-500/20 transition-all"
                        value={data.conclusion.s1Result}
                        onChange={(e) => updateConclusion('s1Result', e.target.value)}
                     >
                        <option value="">請選擇結果</option>
                        <option value="None">未發現相關問題</option>
                        <option value="NoEffect">問題不影響第二階段</option>
                        <option value="AdjustDays">需調節查證人天/日期</option>
                        <option value="Undecided">無法決定</option>
                     </select>
                     <TextArea label="說明" value={data.conclusion.s1Note} onChange={(v: string) => updateConclusion('s1Note', v)} />
                 </div>

                 <div className="bg-white p-6 rounded-3xl border border-emerald-100 shadow-sm space-y-4 ring-1 ring-emerald-50">
                     <h3 className="font-black text-emerald-600 text-lg">第二階段 (S-2) 結論</h3>
                     <select 
                        className="w-full p-4 bg-emerald-50/50 border border-emerald-200 rounded-2xl font-bold text-sm text-emerald-900 focus:bg-white outline-none focus:ring-4 focus:ring-emerald-500/20 transition-all"
                        value={data.conclusion.s2Result}
                        onChange={(e) => updateConclusion('s2Result', e.target.value)}
                     >
                        <option value="">請選擇結果</option>
                        <option value="Corrected">已完成改正</option>
                        <option value="Agree">否 (待組織回覆)</option>
                        <option value="NoFindings">無相關發現</option>
                     </select>
                 </div>
             </div>
            
             <div className="space-y-6">
                <div className="p-6 bg-white border border-slate-100 rounded-3xl shadow-sm space-y-4">
                    <label className="text-xs font-black text-slate-400 uppercase tracking-widest">查證協議變更</label>
                    <div className="flex gap-4">
                        <label className={`flex items-center gap-2 px-4 py-3 rounded-xl border cursor-pointer transition-all ${data.conclusion.protocolChange === 'No' ? 'bg-slate-800 text-white border-slate-800' : 'bg-white border-slate-200 text-slate-500'}`}>
                            <input type="radio" className="hidden" checked={data.conclusion.protocolChange === 'No'} onChange={() => updateConclusion('protocolChange', 'No')} />
                            <span className="font-bold text-sm">無變更</span>
                        </label>
                        <label className={`flex items-center gap-2 px-4 py-3 rounded-xl border cursor-pointer transition-all ${data.conclusion.protocolChange === 'Yes' ? 'bg-blue-600 text-white border-blue-600' : 'bg-white border-slate-200 text-slate-500'}`}>
                            <input type="radio" className="hidden" checked={data.conclusion.protocolChange === 'Yes'} onChange={() => updateConclusion('protocolChange', 'Yes')} />
                            <span className="font-bold text-sm">協議變更</span>
                        </label>
                    </div>
                    {data.conclusion.protocolChange === 'Yes' && (
                        <Field label="說明" value={data.conclusion.protocolChangeNote} onChange={(v: string) => updateConclusion('protocolChangeNote', v)} />
                    )}
                </div>

                <div className="p-6 bg-white border border-slate-100 rounded-3xl shadow-sm space-y-6">
                     <TextArea label="保留意見" value={data.conclusion.reservedOpinion} onChange={(v: string) => updateConclusion('reservedOpinion', v)} />
                     <TextArea label="其他說明" value={data.conclusion.otherNote} onChange={(v: string) => updateConclusion('otherNote', v)} />
                </div>

                <div className="bg-white border border-slate-200 p-8 rounded-3xl shadow-sm grid grid-cols-1 md:grid-cols-2 gap-8">
                    <div className="space-y-4">
                        <h4 className="font-black text-lg text-slate-800">受查組織簽署</h4>
                        <DateField label="簽署日期" value={data.conclusion.auditeeDate} onChange={(v: string) => updateConclusion('auditeeDate', v)} />
                    </div>
                    <div className="space-y-4">
                        <h4 className="font-black text-lg text-slate-800">查驗機構簽署</h4>
                        <DateField label="簽署日期" value={data.conclusion.verifierDate} onChange={(v: string) => updateConclusion('verifierDate', v)} />
                    </div>
                </div>
             </div>
          </div>
      )}
    </div>
  );
};

export default G3027Report;