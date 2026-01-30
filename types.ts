
export enum ComplianceStatus {
  COMPLIANT = '符合',
  NON_COMPLIANT = '不符合',
  CLARIFY = '待釐清',
  NA = '不適用'
}

export enum FinalConclusion {
  PASS = '通過 (Pass)',
  REDUCED = '減少人天 (Reduced Days)',
  PENDING = '待釐清/補正 (Pending)'
}

export interface ChecklistItem {
  id: string;
  name: string;
  docRef: string;
  fieldObs?: string; // 用於 G-3026 的現場觀察說明
  status: ComplianceStatus;
}

export interface InterviewRecord {
  id: string;
  topic: string;
  record: string;
  result: string;
}

export interface PendingItem {
  id: string;
  content: string;
  response: string;
}

export interface SamplingResult {
  id: string;
  area: string;
  value: string;
  source: string;
  type: string;
  ratio: string;
  remarks: string;
}

export interface EmissionFactor {
  id: string;
  item: string;
  source: string;
  description: string;
  remarks: string;
}

// --- G-3027 Interfaces ---

export interface FindingItem {
  id: string;
  stage: 'S1' | 'S2'; // 階段 S1/S2
  type: 'CAR' | 'CR' | 'FAR' | 'OBS' | ''; // 發現事項分類: 增加 OBS
  description: string; // 不符合事項描述
  reporter: string; // 填報查驗人員
  correctiveAction: string; // 矯正措施/澄清說明 (Auditee)
  reviewOpinion: string; // 審查意見 (Verifier)
  reviewer: string; // 查驗員
  result: 'Close' | 'Keep' | ''; // 結案/保留
  location: 'OnSite' | 'OffSite' | ''; // 現場/非現場
}

export interface G3027Stats {
  s1: { nonConformity: string; observation: string; suggestion: string };
  s2: { nonConformity: string; observation: string; suggestion: string };
}

export interface G3027Conclusion {
  s1Result: 'None' | 'NoEffect' | 'AdjustDays' | 'Undecided' | '';
  s1Note: string; // 說明
  s2Result: 'Corrected' | 'Agree' | 'NoFindings' | '';
  protocolChange: 'No' | 'Yes';
  protocolChangeNote: string;
  reservedOpinion: string;
  otherNote: string;
  auditeeDate: string; // 受查組織簽署日期
  verifierDate: string; // 查驗機構簽署日期
}

export interface AppState {
  reportType: 'G-3022' | 'G-3026' | 'G-3027';
  g3022: {
    basicInfo: {
      clientName: string;
      clientAddress: string;
      reviewDate: string;
      visitDate: string;
      caseNumber: string;
      // Refactored Assurance Level Structure
      reasonableScopes: string[]; // e.g., ['cat1', 'cat2']
      limitedScopes: string[];    // e.g., ['cat3', 'cat4', 'cat5', 'cat6']
      materiality: string;
      baseYear: string;
      baseYearEmissions: string;
      verificationYear: string;
      intendedUser: string;
      reportName: string;
      inventoryName: string;
      procedureName: string;
    };
    emissions: {
      cat1: number;
      cat2: number;
      cat3: number;
      cat4: number;
      cat5: number;
      cat6: number;
      uncertaintyUpper: string;
      uncertaintyLower: string;
    };
    checklist: ChecklistItem[];
    conclusion: {
      conflictOfInterest: 'Yes' | 'No';
      conflictDetail: string;
      summary: FinalConclusion;
      otherNote: string;
      memoCorrection: boolean; // S-1 備註勾選
      interviews: InterviewRecord[];
      pendingItems: PendingItem[];
      verifierName: string;
      leadVerifierName: string;
      clientRepName: string;
    };
  };
  g3026: {
    basicInfo: {
      caseNumber: string;
      stage: 'S1' | 'S2';
      year: string;
      checkDate: string;
      reportInfo: string;
      inventoryInfo: string;
      powerFactorInfo: string;
      otherInfo: string;
      clientName?: string;
      clientAddress?: string;
    };
    checklist: ChecklistItem[];
    samplingResults: SamplingResult[];
    emissionFactors: EmissionFactor[];
    otherObservation: string;
    leadVerifierName: string;
  };
  g3027: {
    basicInfo: {
      caseNumber: string;
      stage: 'S1' | 'S2'; // To toggle checklist header title S1/S2 check
      verificationYear: string;
      leadVerifier: string;
      auditeeRep: string;
      date: string;
    };
    findings: FindingItem[];
    stats: G3027Stats;
    conclusion: G3027Conclusion;
  };
}