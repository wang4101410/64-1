import React, { useState, useEffect, useCallback } from 'react';
import { 
  BookOpen, Layers, LayoutDashboard, FileText, 
  AlertCircle, ClipboardList, CheckCircle2 
} from 'lucide-react';
import { 
  AppState, FinalConclusion, 
} from './types';
import { G3022_DEFAULT_CHECKLIST, G3026_DEFAULT_CHECKLIST } from './constants';
import G3022Report from './G3022Report';
import G3026Report from './G3026Report';
import G3027Report from './G3027Report';
import { apiService } from './src/api.ts';

const App: React.FC = () => {
  const [activeReport, setActiveReport] = useState<'G-3022' | 'G-3026' | 'G-3027'>('G-3022');
  const [userId] = useState(() => localStorage.getItem('userId') || 'default-user');
  
  const [state, setState] = useState<AppState>({
    reportType: 'G-3022',
    g3022: {
      basicInfo: {
        clientName: '',
        clientAddress: '',
        reviewDate: new Date().toISOString().split('T')[0],
        visitDate: new Date().toISOString().split('T')[0],
        caseNumber: '113-T-0001',
        reasonableScopes: ['cat1', 'cat2'], 
        limitedScopes: ['cat3', 'cat4', 'cat5', 'cat6'],
        materiality: '5%',
        baseYear: '2022',
        baseYearEmissions: '0',
        verificationYear: '2023',
        intendedUser: '預期使用者',
        reportName: '',
        inventoryName: '',
        procedureName: '',
      },
      emissions: {
        cat1: 0, cat2: 0, cat3: 0, cat4: 0, cat5: 0, cat6: 0,
        uncertaintyUpper: '',
        uncertaintyLower: '',
      },
      checklist: G3022_DEFAULT_CHECKLIST,
      conclusion: {
        conflictOfInterest: 'No',
        conflictDetail: '',
        summary: FinalConclusion.PASS,
        otherNote: '',
        memoCorrection: false,
        interviews: [],
        pendingItems: [],
        verifierName: '',
        leadVerifierName: '',
        clientRepName: '',
      },
    },
    g3026: {
      basicInfo: {
        caseNumber: '113-T-0001',
        stage: 'S1',
        year: '113',
        checkDate: new Date().toISOString().split('T')[0],
        reportInfo: '',
        inventoryInfo: '',
        powerFactorInfo: '',
        otherInfo: '',
        clientName: '',
        clientAddress: '',
      },
      checklist: G3026_DEFAULT_CHECKLIST,
      samplingResults: [],
      emissionFactors: [],
      otherObservation: '',
      leadVerifierName: '',
    },
    g3027: {
      basicInfo: {
        caseNumber: '113-T-0001',
        stage: 'S1',
        verificationYear: '113',
        leadVerifier: '',
        auditeeRep: '',
        date: new Date().toISOString().split('T')[0],
      },
      findings: [],
      stats: {
        s1: { nonConformity: '0', observation: '0', suggestion: '0' },
        s2: { nonConformity: '0', observation: '0', suggestion: '0' },
      },
      conclusion: {
        s1Result: '',
        s1Note: '',
        s2Result: '',
        protocolChange: 'No',
        protocolChangeNote: '',
        reservedOpinion: '',
        otherNote: '',
        auditeeDate: '',
        verifierDate: '',
      },
    },
  });

  // 自動存檔功能
  const saveData = useCallback(async (data: AppState) => {
    try {
      const response = await apiService.saveUserData(userId, data);
      if (response.success) {
        console.log('數據已自動保存');
      } else {
        console.error('自動保存失敗:', response.error);
      }
    } catch (error) {
      console.error('自動保存錯誤:', error);
    }
  }, [userId]);

  // 監聽state變化並自動保存
  useEffect(() => {
    const timeoutId = setTimeout(() => {
      saveData(state);
    }, 2000); // 2秒後保存

    return () => clearTimeout(timeoutId);
  }, [state, saveData]);

  // 載入數據
  useEffect(() => {
    const loadData = async () => {
      try {
        const response = await apiService.loadUserData(userId);
        if (response.success && response.data) {
          setState(response.data);
        }
      } catch (error) {
        console.error('載入數據錯誤:', error);
      }
    };
    loadData();
  }, [userId]);

  // --- Data Synchronization Core ---
  const syncData = (source: 'G-3022' | 'G-3026' | 'G-3027', newState: AppState) => {
    let finalState = { ...newState };
    
    // Extract key shared data points from the SOURCE
    let shared = {
      caseNumber: '',
      clientName: '',
      clientAddress: '',
      leadVerifier: '',
      clientRep: '',
      visitDate: '', // YYYY-MM-DD
      docReport: '',
      docInventory: '',
      docProcedure: '',
    };

    if (source === 'G-3022') {
      const d = newState.g3022;
      shared = {
        caseNumber: d.basicInfo.caseNumber,
        clientName: d.basicInfo.clientName,
        clientAddress: d.basicInfo.clientAddress,
        leadVerifier: d.conclusion.leadVerifierName,
        clientRep: d.conclusion.clientRepName,
        visitDate: d.basicInfo.visitDate,
        docReport: d.basicInfo.reportName,
        docInventory: d.basicInfo.inventoryName,
        docProcedure: d.basicInfo.procedureName,
      };
    } else if (source === 'G-3026') {
      const d = newState.g3026;
      shared = {
        caseNumber: d.basicInfo.caseNumber,
        clientName: d.basicInfo.clientName || '',
        clientAddress: d.basicInfo.clientAddress || '',
        leadVerifier: d.leadVerifierName,
        clientRep: finalState.g3022.conclusion.clientRepName, 
        visitDate: d.basicInfo.checkDate,
        docReport: d.basicInfo.reportInfo,
        docInventory: d.basicInfo.inventoryInfo,
        docProcedure: d.basicInfo.powerFactorInfo,
      };
    } else if (source === 'G-3027') {
      const d = newState.g3027;
      shared = {
        caseNumber: d.basicInfo.caseNumber,
        clientName: finalState.g3022.basicInfo.clientName, 
        clientAddress: finalState.g3022.basicInfo.clientAddress, 
        leadVerifier: d.basicInfo.leadVerifier,
        clientRep: d.basicInfo.auditeeRep,
        visitDate: d.basicInfo.date,
        docReport: finalState.g3022.basicInfo.reportName, 
        docInventory: finalState.g3022.basicInfo.inventoryName, 
        docProcedure: finalState.g3022.basicInfo.procedureName, 
      };
    }

    // Apply shared data to G-3022
    if (source !== 'G-3022') {
      finalState.g3022.basicInfo.caseNumber = shared.caseNumber;
      finalState.g3022.basicInfo.clientName = shared.clientName;
      finalState.g3022.basicInfo.clientAddress = shared.clientAddress;
      finalState.g3022.conclusion.leadVerifierName = shared.leadVerifier;
      finalState.g3022.conclusion.clientRepName = shared.clientRep;
      finalState.g3022.basicInfo.visitDate = shared.visitDate;
      if (source === 'G-3026') {
        finalState.g3022.basicInfo.reportName = shared.docReport;
        finalState.g3022.basicInfo.inventoryName = shared.docInventory;
        finalState.g3022.basicInfo.procedureName = shared.docProcedure;
      }
    }

    // Apply shared data to G-3026
    if (source !== 'G-3026') {
      finalState.g3026.basicInfo.caseNumber = shared.caseNumber;
      finalState.g3026.basicInfo.clientName = shared.clientName;
      finalState.g3026.basicInfo.clientAddress = shared.clientAddress;
      finalState.g3026.leadVerifierName = shared.leadVerifier;
      finalState.g3026.basicInfo.checkDate = shared.visitDate;
      finalState.g3026.basicInfo.reportInfo = shared.docReport;
      finalState.g3026.basicInfo.inventoryInfo = shared.docInventory;
      finalState.g3026.basicInfo.powerFactorInfo = shared.docProcedure;
    }

    // Apply shared data to G-3027
    if (source !== 'G-3027') {
      finalState.g3027.basicInfo.caseNumber = shared.caseNumber;
      finalState.g3027.basicInfo.leadVerifier = shared.leadVerifier;
      finalState.g3027.basicInfo.auditeeRep = shared.clientRep;
      finalState.g3027.basicInfo.date = shared.visitDate;
    }

    setState(finalState);
  };

  const updateG3022 = (newData: AppState['g3022']) => {
    syncData('G-3022', { ...state, g3022: newData });
  };

  const updateG3026 = (newData: AppState['g3026']) => {
    syncData('G-3026', { ...state, g3026: newData });
  };

  const updateG3027 = (newData: AppState['g3027']) => {
    syncData('G-3027', { ...state, g3027: newData });
  };

  // Reset handlers with Deep Updates
  const resetG3022 = () => {
    if (window.confirm("確定要重置 G-3022 嗎？")) {
      const freshChecklist = JSON.parse(JSON.stringify(G3022_DEFAULT_CHECKLIST));
      setState(prev => ({
        ...prev,
        g3022: {
           ...prev.g3022,
           checklist: freshChecklist,
           emissions: { cat1: 0, cat2: 0, cat3: 0, cat4: 0, cat5: 0, cat6: 0, uncertaintyUpper: '', uncertaintyLower: '' },
           conclusion: { 
             ...prev.g3022.conclusion, 
             conflictOfInterest: 'No',
             conflictDetail: '',
             summary: FinalConclusion.PASS,
             otherNote: '',
             memoCorrection: false,
             interviews: [], 
             pendingItems: [],
           }
        }
      }));
    }
  };

  const resetG3026 = () => {
    if (window.confirm("確定要重置 G-3026 嗎？")) {
       const freshChecklist = JSON.parse(JSON.stringify(G3026_DEFAULT_CHECKLIST));
       setState(prev => ({
         ...prev,
         g3026: {
           ...prev.g3026,
           checklist: freshChecklist,
           samplingResults: [],
           emissionFactors: [],
           otherObservation: ''
         }
       }));
    }
  };

  const resetG3027 = () => {
      if (window.confirm("確定要重置 G-3027 嗎？")) {
        setState(prev => ({
          ...prev,
          g3027: {
            ...prev.g3027,
            findings: [],
            stats: { s1: { nonConformity: '0', observation: '0', suggestion: '0' }, s2: { nonConformity: '0', observation: '0', suggestion: '0' } },
            conclusion: { ...prev.g3027.conclusion, s1Result: '', s1Note: '', s2Result: '', protocolChange: 'No', protocolChangeNote: '', reservedOpinion: '', otherNote: '' }
          }
        }));
      }
  };

  return (
    <div className="flex min-h-screen bg-[#F8FAFC] font-sans text-slate-800">
      {/* Light Sidebar */}
      <aside className="w-72 bg-white border-r border-slate-200 flex-shrink-0 sticky top-0 h-screen flex flex-col shadow-sm z-20">
        <div className="p-8 pb-4">
           <div className="flex items-center gap-3 mb-2">
             <div className="bg-blue-600 p-2 rounded-xl shadow-lg shadow-blue-600/20">
                <Layers size={24} className="text-white" />
             </div>
             <div>
                <h1 className="text-xl font-black tracking-tight text-slate-800">ISO 14064-1</h1>
                <div className="text-[10px] font-bold text-blue-600 uppercase tracking-widest">Generator Pro</div>
             </div>
           </div>
        </div>

        <nav className="flex-1 px-4 py-4 space-y-2 overflow-y-auto custom-scrollbar">
          <div className="text-xs font-black text-slate-400 uppercase tracking-widest px-4 mb-2">Reports</div>
          
          <NavItem 
            active={activeReport === 'G-3022'} 
            onClick={() => setActiveReport('G-3022')}
            icon={<LayoutDashboard size={20} />}
            title="G-3022 總結報告"
            subtitle="書面審查與訪談結論"
          />
          
          <NavItem 
            active={activeReport === 'G-3026'} 
            onClick={() => setActiveReport('G-3026')}
            icon={<ClipboardList size={20} />}
            title="G-3026 觀察報告"
            subtitle="現場查驗與取樣紀錄"
          />
          
          <NavItem 
            active={activeReport === 'G-3027'} 
            onClick={() => setActiveReport('G-3027')}
            icon={<AlertCircle size={20} />}
            title="G-3027 差異分析"
            subtitle="不符合與矯正措施"
          />
        </nav>

        <div className="p-6 border-t border-slate-100 bg-slate-50/50">
           <div className="flex items-center gap-3 text-xs font-medium text-slate-500">
              <BookOpen size={16}/>
              <span>Standard: 2018 Edition</span>
           </div>
           <div className="mt-2 text-[10px] text-slate-400">
             © 2024 旭威認證 Co., Ltd.
           </div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 min-w-0">
         {/* Top Header */}
         <header className="bg-white/80 backdrop-blur-md sticky top-0 z-10 border-b border-slate-200 px-8 py-5 flex justify-between items-center shadow-sm">
            <div>
              <h2 className="text-2xl font-black text-slate-800 tracking-tight">
                {activeReport === 'G-3022' && 'G-3022 查驗總結報告'}
                {activeReport === 'G-3026' && 'G-3026 查驗觀察報告'}
                {activeReport === 'G-3027' && 'G-3027 不符合事項摘要'}
              </h2>
              <div className="flex items-center gap-2 mt-1">
                 <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-md bg-slate-100 text-slate-600 text-xs font-bold border border-slate-200">
                    <FileText size={12}/> 案件編號: {state.g3022.basicInfo.caseNumber || '尚未輸入'}
                 </span>
                 <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-md bg-blue-50 text-blue-600 text-xs font-bold border border-blue-100">
                    <CheckCircle2 size={12}/> 狀態: 編輯中
                 </span>
              </div>
            </div>
            
            <div className="flex items-center gap-4">
              {/* Global Actions could go here */}
            </div>
         </header>

         {/* Content Area */}
         <div className="p-8 max-w-[1600px] mx-auto animate-in fade-in slide-in-from-bottom-4 duration-500">
            {activeReport === 'G-3022' && (
              <G3022Report data={state.g3022} onChange={updateG3022} onReset={resetG3022} />
            )}

            {activeReport === 'G-3026' && (
              <G3026Report 
                data={state.g3026} 
                g3022Data={state.g3022}
                onChange={updateG3026} 
                onReset={resetG3026} 
              />
            )}

            {activeReport === 'G-3027' && (
                <G3027Report
                    data={state.g3027}
                    onChange={updateG3027}
                    onReset={resetG3027}
                />
            )}
         </div>
      </main>
    </div>
  );
};

const NavItem = ({ active, onClick, icon, title, subtitle }: any) => (
  <button 
    onClick={onClick}
    className={`w-full text-left px-4 py-3 mx-2 rounded-xl transition-all duration-200 flex items-center gap-4 group relative overflow-hidden ${
      active 
      ? 'bg-blue-50 text-blue-700 ring-1 ring-blue-100 shadow-sm' 
      : 'text-slate-500 hover:bg-slate-50 hover:text-slate-700'
    }`}
  >
    <div className={`p-2 rounded-lg transition-colors ${active ? 'bg-white text-blue-600 shadow-sm' : 'bg-slate-100 text-slate-400 group-hover:bg-white group-hover:text-slate-500'}`}>
       {icon}
    </div>
    <div>
       <div className={`font-black text-sm ${active ? 'text-blue-800' : 'text-slate-600 group-hover:text-slate-800'}`}>{title}</div>
       <div className={`text-[10px] font-bold mt-0.5 ${active ? 'text-blue-400' : 'text-slate-400 group-hover:text-slate-500'}`}>{subtitle}</div>
    </div>
    {active && <div className="absolute right-0 top-1/2 -translate-y-1/2 w-1 h-8 bg-blue-600 rounded-l-full"></div>}
  </button>
);

export default App;