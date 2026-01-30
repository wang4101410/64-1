import React, { useState } from 'react';
import { Download, Plus, Trash2, Calendar, FileText, CheckCircle2, ListChecks } from 'lucide-react';
import { AppState, ComplianceStatus } from './types';
import { generateG3026Docx } from './G3026Generator';

interface Props {
  data: AppState['g3026'];
  g3022Data: AppState['g3022'];
  onChange: (newData: AppState['g3026']) => void;
  onReset: () => void;
}

const G3026Report: React.FC<Props> = ({ data, g3022Data, onChange, onReset }) => {
  const [activeTab, setActiveTab] = useState(0);
  const [isExporting, setIsExporting] = useState(false);

  const updateBasic = (field: string, value: any) => {
    onChange({ ...data, basicInfo: { ...data.basicInfo, [field]: value } });
  };

  const updateChecklist = (id: string, field: string, value: any) => {
    const newList = data.checklist.map(item => 
      item.id === id ? { ...item, [field]: value } : item
    );
    onChange({ ...data, checklist: newList });
  };

  const addSampling = () => {
    const newItem = { id: Date.now().toString(), area: '', value: '', source: '', type: '', ratio: '', remarks: '' };
    onChange({ ...data, samplingResults: [...data.samplingResults, newItem] });
  };

  const addFactor = () => {
    const newItem = { id: Date.now().toString(), item: '', source: '', description: '', remarks: '' };
    onChange({ ...data, emissionFactors: [...data.emissionFactors, newItem] });
  };

  const handleExport = async () => {
    try {
      setIsExporting(true);
      const blob = await generateG3026Docx(data, g3022Data);
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `G-3026_Report_${data.basicInfo.caseNumber || 'Draft'}.docx`;
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

  // Group definitions matching G3026Generator
  const checklistGroups = [
    { idPrefix: '1', title: "組織邊界" },
    { idPrefix: '2', title: "報告邊界" },
    { idPrefix: '3', title: "量化方法" },
    { idPrefix: '4', title: "基準年" },
    { idPrefix: '5', title: "數據品質" },
  ];

  return (
    <div className="space-y-6">
      {/* Action Bar */}
      <div className="flex justify-between items-center bg-white p-4 rounded-2xl shadow-sm border border-slate-100">
        <div className="flex gap-2 bg-slate-100 p-1 rounded-xl">
          {['基本資訊', '觀察檢核表', '取樣結果 (附一)', '係數確認 (附二)'].map((label, idx) => (
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
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 animate-in fade-in slide-in-from-bottom-2 duration-300">
           <div className="space-y-6">
             <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-4">
                <div className="flex items-center gap-3 border-b border-slate-50 pb-4">
                   <div className="bg-blue-50 p-2 rounded-xl text-blue-600"><FileText size={20}/></div>
                   <h3 className="font-black text-slate-800 text-lg">查驗基本資訊</h3>
                </div>
                <Field label="案件編號" value={data.basicInfo.caseNumber} onChange={(v: string) => updateBasic('caseNumber', v)} />
                <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                   <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-2">查驗階段</label>
                   <div className="flex gap-2">
                      {['S1', 'S2'].map(s => (
                         <button key={s} onClick={() => updateBasic('stage', s)} className={`flex-1 py-2.5 rounded-xl font-bold text-sm transition-all ${data.basicInfo.stage === s ? 'bg-slate-800 text-white shadow-lg shadow-slate-200' : 'bg-white border border-slate-200 text-slate-500 hover:bg-slate-100'}`}>{s}</button>
                      ))}
                   </div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <Field label="查驗年度 (民國)" value={data.basicInfo.year} onChange={(v: string) => updateBasic('year', v)} />
                    <DateField label="查驗日期" value={data.basicInfo.checkDate} onChange={(v: string) => updateBasic('checkDate', v)} />
                </div>
             </div>
           </div>
           
           <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-4">
              <div className="flex items-center gap-3 border-b border-slate-50 pb-4">
                   <div className="bg-indigo-50 p-2 rounded-xl text-indigo-600"><ListChecks size={20}/></div>
                   <h3 className="font-black text-slate-800 text-lg">依據文件資訊</h3>
              </div>
              <Field label="溫室氣體報告 (編號/版次/日期)" value={data.basicInfo.reportInfo} onChange={(v: string) => updateBasic('reportInfo', v)} />
              <Field label="盤查清冊 (編號/版次/日期)" value={data.basicInfo.inventoryInfo} onChange={(v: string) => updateBasic('inventoryInfo', v)} />
              <Field label="資訊管理程序 (編號/版次/日期)" value={data.basicInfo.powerFactorInfo} onChange={(v: string) => updateBasic('powerFactorInfo', v)} />
              <Field label="其他文件" value={data.basicInfo.otherInfo} onChange={(v: string) => updateBasic('otherInfo', v)} />
           </div>
        </div>
      )}

      {activeTab === 1 && (
        <div className="animate-in fade-in slide-in-from-bottom-2 duration-300 space-y-6">
           <div className="bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
             <div className="overflow-x-auto">
                <table className="w-full text-left text-sm">
                <thead className="bg-slate-50 border-b border-slate-200 text-slate-500">
                    <tr>
                        <th className="p-4 w-20 text-center font-bold">分類</th>
                        <th className="p-4 w-1/4 font-bold">查證內容</th>
                        <th className="p-4 w-1/4 font-bold">查驗文件</th>
                        <th className="p-4 font-bold">現場觀察說明</th>
                        <th className="p-4 w-32 text-center font-bold">狀態</th>
                    </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                    {checklistGroups.map((group) => {
                    const items = data.checklist.filter(i => i.id.startsWith(group.idPrefix));
                    if (items.length === 0) return null;

                    return items.map((item, index) => (
                        <tr key={item.id} className="hover:bg-slate-50/50 transition-colors">
                        {index === 0 && (
                            <td 
                            rowSpan={items.length} 
                            className="p-4 bg-slate-50/30 font-black text-slate-400 text-center border-r border-slate-100 align-middle text-xs writing-vertical"
                            >
                            <span className="block mb-2 text-lg opacity-20">{group.idPrefix}</span>
                            {group.title}
                            </td>
                        )}
                        <td className="p-4 align-top">
                            <div className="flex gap-3">
                            <span className="font-black text-slate-300 shrink-0 text-xs mt-1">{item.id}</span>
                            <span className="text-slate-700 font-bold leading-relaxed">{item.name}</span>
                            </div>
                        </td>
                        <td className="p-4 align-top">
                            <textarea 
                                className="w-full bg-slate-50 border border-slate-200 hover:bg-white focus:bg-white rounded-xl p-3 text-xs focus:ring-2 focus:ring-blue-100 focus:border-blue-500 min-h-[80px] resize-none outline-none transition-all font-medium" 
                                placeholder="輸入文件編號..."
                                value={item.docRef || ''} 
                                onChange={(e) => updateChecklist(item.id, 'docRef', e.target.value)} 
                            />
                        </td>
                        <td className="p-4 align-top">
                            <textarea 
                                className="w-full bg-slate-50 border border-slate-200 hover:bg-white focus:bg-white rounded-xl p-3 text-xs focus:ring-2 focus:ring-blue-100 focus:border-blue-500 min-h-[80px] resize-none outline-none transition-all font-medium" 
                                value={item.fieldObs || ''} 
                                onChange={(e) => updateChecklist(item.id, 'fieldObs', e.target.value)} 
                                placeholder="輸入現場觀察說明..." 
                            />
                        </td>
                        <td className="p-4 text-center align-top">
                             <select 
                                className={`w-full p-2.5 rounded-lg text-xs font-bold appearance-none cursor-pointer outline-none border transition-all text-center ${
                                    item.status === ComplianceStatus.COMPLIANT ? 'bg-emerald-50 text-emerald-600 border-emerald-200' : 
                                    item.status === ComplianceStatus.NON_COMPLIANT ? 'bg-red-50 text-red-600 border-red-200' :
                                    item.status === ComplianceStatus.NA ? 'bg-slate-100 text-slate-500 border-slate-200' : 'bg-slate-100 text-slate-500 border-slate-200'
                                }`}
                                value={item.status}
                                onChange={(e) => updateChecklist(item.id, 'status', e.target.value)}
                            >
                                <option value={ComplianceStatus.COMPLIANT}>符合</option>
                                <option value={ComplianceStatus.NON_COMPLIANT}>不符合</option>
                                <option value={ComplianceStatus.NA}>不適用</option>
                            </select>
                        </td>
                        </tr>
                    ));
                    })}
                </tbody>
                </table>
             </div>
           </div>

           <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-3">
              <label className="text-xs font-black text-slate-400 uppercase tracking-widest flex items-center gap-2"><FileText size={16}/> 其他觀察事項說明</label>
              <textarea 
                className="w-full p-4 bg-slate-50 border border-slate-200 rounded-2xl h-32 focus:bg-white focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 outline-none font-bold text-slate-700 text-sm transition-all" 
                value={data.otherObservation} 
                onChange={(e) => onChange({...data, otherObservation: e.target.value})} 
                placeholder="若有其他未列於上表的觀察事項，請在此補充..."
              />
           </div>
        </div>
      )}

      {activeTab === 2 && (
         <div className="space-y-6 animate-in fade-in slide-in-from-bottom-2 duration-300">
            <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-4">
                <div className="flex justify-between items-center border-b border-slate-50 pb-4">
                    <h3 className="font-black text-slate-800 text-lg">附件一：查驗取樣結果</h3>
                    <button onClick={addSampling} className="bg-emerald-50 text-emerald-600 border border-emerald-100 px-3 py-1.5 rounded-lg text-xs font-bold hover:bg-emerald-100 flex items-center gap-1 transition-colors"><Plus size={14}/> 新增</button>
                </div>
                {data.samplingResults.length === 0 && <div className="text-center py-10 text-slate-400 font-medium bg-slate-50 rounded-2xl border border-dashed border-slate-200">尚無取樣資料</div>}
                {data.samplingResults.map((r, i) => (
                    <div key={r.id} className="grid grid-cols-1 md:grid-cols-12 gap-4 p-5 bg-slate-50 rounded-2xl relative group border border-slate-200 hover:shadow-md transition-all">
                        <div className="md:col-span-2"><Field label="區域/排放源" value={r.area} onChange={(v: string) => { const n = [...data.samplingResults]; n[i].area = v; onChange({...data, samplingResults: n}); }} /></div>
                        <div className="md:col-span-2"><Field label="數值" value={r.value} onChange={(v: string) => { const n = [...data.samplingResults]; n[i].value = v; onChange({...data, samplingResults: n}); }} /></div>
                        <div className="md:col-span-2"><Field label="來源" value={r.source} onChange={(v: string) => { const n = [...data.samplingResults]; n[i].source = v; onChange({...data, samplingResults: n}); }} /></div>
                        <div className="md:col-span-2"><Field label="類型" value={r.type} onChange={(v: string) => { const n = [...data.samplingResults]; n[i].type = v; onChange({...data, samplingResults: n}); }} /></div>
                        <div className="md:col-span-2"><Field label="排放源佔總排放量比" value={r.ratio} onChange={(v: string) => { const n = [...data.samplingResults]; n[i].ratio = v; onChange({...data, samplingResults: n}); }} /></div>
                        <div className="md:col-span-2"><Field label="備註" value={r.remarks} onChange={(v: string) => { const n = [...data.samplingResults]; n[i].remarks = v; onChange({...data, samplingResults: n}); }} /></div>
                        <button onClick={() => onChange({...data, samplingResults: data.samplingResults.filter(sr => sr.id !== r.id)})} className="absolute top-2 right-2 p-1.5 text-slate-300 hover:text-red-500 hover:bg-red-50 rounded-lg transition-all opacity-0 group-hover:opacity-100"><Trash2 size={16}/></button>
                    </div>
                ))}
            </div>
         </div>
      )}

      {activeTab === 3 && (
         <div className="space-y-6 animate-in fade-in slide-in-from-bottom-2 duration-300">
            <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm space-y-4">
                <div className="flex justify-between items-center border-b border-slate-50 pb-4">
                    <h3 className="font-black text-slate-800 text-lg">附件二：排放係數確認</h3>
                    <button onClick={addFactor} className="bg-amber-50 text-amber-600 border border-amber-100 px-3 py-1.5 rounded-lg text-xs font-bold hover:bg-amber-100 flex items-center gap-1 transition-colors"><Plus size={14}/> 新增</button>
                </div>
                {data.emissionFactors.length === 0 && <div className="text-center py-10 text-slate-400 font-medium bg-slate-50 rounded-2xl border border-dashed border-slate-200">尚無係數資料</div>}
                {data.emissionFactors.map((f, i) => (
                <div key={f.id} className="grid grid-cols-1 md:grid-cols-4 gap-4 p-5 bg-amber-50/30 rounded-2xl relative group border border-amber-100 hover:shadow-md transition-all">
                    <Field label="係數項目" value={f.item} onChange={(v: string) => { const n = [...data.emissionFactors]; n[i].item = v; onChange({...data, emissionFactors: n}); }} />
                    <Field label="來源" value={f.source} onChange={(v: string) => { const n = [...data.emissionFactors]; n[i].source = v; onChange({...data, emissionFactors: n}); }} />
                    <div className="md:col-span-2"><Field label="說明" value={f.description} onChange={(v: string) => { const n = [...data.emissionFactors]; n[i].description = v; onChange({...data, emissionFactors: n}); }} /></div>
                    <button onClick={() => onChange({...data, emissionFactors: data.emissionFactors.filter(ef => ef.id !== f.id)})} className="absolute top-2 right-2 p-1.5 text-slate-300 hover:text-red-500 hover:bg-red-50 rounded-lg transition-all opacity-0 group-hover:opacity-100"><Trash2 size={16}/></button>
                </div>
                ))}
            </div>
            
            <div className="bg-white border border-slate-200 p-6 rounded-3xl shadow-sm">
               <div className="flex items-center gap-3 mb-6">
                  <div className="bg-blue-600 p-2 rounded-xl text-white"><CheckCircle2 size={20}/></div>
                  <h3 className="font-black text-lg text-slate-800">簽署確認</h3>
               </div>
               <Field label="主導查驗員簽名" value={data.leadVerifierName} onChange={(v: string) => onChange({...data, leadVerifierName: v})} />
            </div>
         </div>
      )}
    </div>
  );
};

const Field = ({ label, placeholder, value, onChange, type = "text" }: any) => (
  <div className="space-y-1.5 w-full group">
    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1 group-focus-within:text-blue-500 transition-colors">{label}</label>
    <input 
      type={type} 
      className="w-full p-3.5 bg-slate-50 border border-slate-200 rounded-2xl focus:bg-white focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 outline-none transition-all font-bold text-sm text-slate-700 placeholder:text-slate-300" 
      placeholder={placeholder} 
      value={value} 
      onChange={(e) => onChange(e.target.value)} 
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

export default G3026Report;