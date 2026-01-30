import { AppState, ComplianceStatus, FinalConclusion } from './types';
import { 
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
  WidthType, BorderStyle, Header, Footer, 
  VerticalAlign, ShadingType, HeightRule, TableLayoutType, SectionType, PageOrientation, PageNumber, AlignmentType, ISectionOptions
} from 'docx';

// ════════════════════════════════════════════════════════════════════════════
// DIMENSIONS & CONFIG
// ════════════════════════════════════════════════════════════════════════════
// Portrait A4 Printable Width: 21cm - 1.5cm - 1.5cm = 18cm ≈ 10205 DXA
const PORTRAIT_WIDTH = 10205;

// Margins: Top/Bottom 2cm (1134), Left/Right 1.5cm (851)
// Header/Footer: 0.8cm (454)
const MARGINS_PORTRAIT = { 
    top: 1134, 
    bottom: 1134, 
    left: 851, 
    right: 851,
    header: 454,
    footer: 454
};

const FONTS_OPTS = {
  ascii: "Times New Roman",
  hAnsi: "Times New Roman",
  eastAsia: "標楷體",
  cs: "Times New Roman",
  hint: "eastAsia" as const
};

// Unified Border Style (Consistent Thickness)
const BORDER_STD = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const BORDER_NONE = { style: BorderStyle.NONE, size: 0, color: "auto" };

// Color #C5E0B3 for all colored backgrounds
const SHADING_COLOR = "C5E0B3"; 

const LineRuleCompat = {
  AUTO: "auto",
  EXACT: "exact",
  AT_LEAST: "atLeast",
} as const;

// ════════════════════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════════════════════
const txt = (text: string, opts: { bold?: boolean; size?: number; underline?: boolean; break?: number } = {}) => {
  const runs = [];
  if (opts.break) {
      for (let i = 0; i < opts.break; i++) runs.push(new TextRun({ text: "\n" }));
  }
  runs.push(new TextRun({
    text: text,
    font: FONTS_OPTS,
    bold: opts.bold,
    size: (opts.size || 12) * 2, // helper converts pt to half-pt
    underline: opts.underline ? { type: 'single' } : undefined,
  }));
  return runs;
};

// Helper for CO2e with subscript 2
const txtCO2e = (prefix: string, suffix: string = "") => {
    return [
        new TextRun({ text: prefix, font: FONTS_OPTS, size: 24 }),
        new TextRun({ text: "CO", font: FONTS_OPTS, size: 24 }),
        new TextRun({ text: "2", font: FONTS_OPTS, size: 24, subScript: true }),
        new TextRun({ text: "e", font: FONTS_OPTS, size: 24 }),
        new TextRun({ text: suffix, font: FONTS_OPTS, size: 24 }),
    ];
};

const checkbox = (checked: boolean, sizePt: number = 12) => {
  return new TextRun({
    text: checked ? "■" : "□",
    font: FONTS_OPTS,
    size: sizePt * 2
  });
};

const stdP = (children: any[], align: string = "left", spacing?: { before?: number, after?: number, line?: number, lineRule?: typeof LineRuleCompat[keyof typeof LineRuleCompat] }, indent?: { left?: number, hanging?: number, firstLine?: number }) => {
  return new Paragraph({
    children: children.flat(),
    alignment: align as any,
    spacing: spacing || { before: 20, after: 20, line: 240 },
    indent: indent
  });
};

const formatDate = (dateStr: string) => {
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return { y: "    ", m: "  ", d: "  " };
  return {
    y: (d.getFullYear() - 1911).toString(), 
    m: (d.getMonth() + 1).toString().padStart(2, '0'),
    d: d.getDate().toString().padStart(2, '0')
  };
};

const getScopesText = (scopes: string[]) => {
  if (!scopes || scopes.length === 0) return "";
  const map: Record<string, string> = {
    'cat1': '類別1', 'cat2': '類別2', 'cat3': '類別3',
    'cat4': '類別4', 'cat5': '類別5', 'cat6': '類別6'
  };
  return scopes.map(s => map[s] || s).join('、');
};

const formatNumber = (num: string | number) => {
    if (!num) return "0";
    return Number(num).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 4 });
};

// ════════════════════════════════════════════════════════════════════════════
// GENERATOR
// ════════════════════════════════════════════════════════════════════════════
export const generateG3022Docx = async (data: AppState['g3022']): Promise<Blob> => {
  const { basicInfo, emissions, checklist, conclusion } = data;
  const reviewDate = formatDate(basicInfo.reviewDate);
  const visitDate = formatDate(basicInfo.visitDate);

  // Calculate Totals
  const total = 
    (emissions.cat1 || 0) + (emissions.cat2 || 0) + (emissions.cat3 || 0) + 
    (emissions.cat4 || 0) + (emissions.cat5 || 0) + (emissions.cat6 || 0);

  const getPercent = (val: number) => total === 0 ? "0.00" : ((val / total) * 100).toFixed(2);

  // --- Headers & Footers ---
  const mainHeader = new Header({
    children: [
        stdP([txt("旭威認證股份有限公司 查驗機構", { size: 18, bold: true })], "center"),
        stdP([txt("查驗書面審查-赴廠訪談總結報告(2018 年版)", { size: 18, bold: true })], "center"),
        stdP([txt(`案件編號：${basicInfo.caseNumber}`, { bold: true })]),
    ]
  });

  const appendixHeader = new Header({
    children: [
        stdP([txt("旭威認證股份有限公司 查驗機構", { size: 18, bold: true })], "center"),
        stdP([txt("書面審查/現場訪談報告", { size: 18, bold: true })], "center"),
        stdP([txt(`案件編號：${basicInfo.caseNumber}`, { bold: true })]),
    ]
  });

  const commonFooter = new Footer({
    children: [
        stdP([txt("正本：旭威認證股份有限公司查驗機構存檔", { size: 10 })]),
        new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: { top: BORDER_NONE, bottom: BORDER_NONE, left: BORDER_NONE, right: BORDER_NONE, insideHorizontal: BORDER_NONE, insideVertical: BORDER_NONE },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({ width: { size: 33, type: WidthType.PERCENTAGE }, children: [stdP([txt("1141017 版", { size: 10 })], "left")] }),
                        new TableCell({ width: { size: 33, type: WidthType.PERCENTAGE }, children: [
                             stdP([
                                new TextRun({ children: [PageNumber.CURRENT], size: 20, font: FONTS_OPTS }), 
                                txt(" / ", { size: 10 }),
                                new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 20, font: FONTS_OPTS }),
                             ], "center")
                        ] }),
                        new TableCell({ width: { size: 33, type: WidthType.PERCENTAGE }, children: [stdP([txt("G-3022", { size: 10 })], "right")] }),
                    ]
                })
            ]
        })
    ]
  });

  // --- Page 1 Content ---
  const criteriaPara = stdP([txt("查驗準則：ISO 14064-1(2018 年版)/CNS 14064-1(2021 年版)")], "left", { before: 200, after: 200, line: 240 });

  const clientInfoTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: { top: BORDER_NONE, bottom: BORDER_NONE, left: BORDER_NONE, right: BORDER_NONE, insideHorizontal: BORDER_NONE, insideVertical: BORDER_NONE },
      rows: [
          new TableRow({ children: [ new TableCell({ children: [stdP([txt(`委託單位名稱：${basicInfo.clientName}`)])] }) ] }),
          new TableRow({ children: [ new TableCell({ children: [stdP([txt(`委託單位地址：${basicInfo.clientAddress}`)])] }) ] }),
          new TableRow({ children: [ new TableCell({ children: [stdP([
              checkbox(true), txt(`書面審查日期： ${reviewDate.y} 年 ${reviewDate.m} 月 ${reviewDate.d} 日      `),
              checkbox(true), txt(`赴廠訪談日期： ${visitDate.y} 年 ${visitDate.m} 月 ${visitDate.d} 日`)
          ])] }) ] }),
          new TableRow({ children: [ new TableCell({ children: [stdP([txt(" ")])] }) ] }), // Spacer
      ]
  });

  const section1Title = stdP([txt("一、基本資料：", { bold: true, size: 14 })]);

  const BASIC_LABEL_WIDTH = 2750;
  const BASIC_CONTENT_WIDTH = PORTRAIT_WIDTH - BASIC_LABEL_WIDTH;
  const ROW_HEIGHT = { value: 600, rule: HeightRule.ATLEAST };

  const basicInfoTable = new Table({
      layout: TableLayoutType.FIXED,
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: { top: BORDER_STD, bottom: BORDER_STD, left: BORDER_STD, right: BORDER_STD, insideHorizontal: BORDER_STD, insideVertical: BORDER_STD },
      rows: [
          new TableRow({
              height: ROW_HEIGHT,
              children: [
                  new TableCell({ width: { size: BASIC_LABEL_WIDTH, type: WidthType.DXA }, shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, children: [stdP([txt("1.保證等級")])] }),
                  new TableCell({
                      width: { size: BASIC_CONTENT_WIDTH, type: WidthType.DXA },
                      verticalAlign: VerticalAlign.CENTER,
                      children: [
                          stdP([checkbox(basicInfo.reasonableScopes.length > 0), txt("合理等級："), txt(getScopesText(basicInfo.reasonableScopes))]),
                          stdP([checkbox(basicInfo.limitedScopes.length > 0), txt("有限等級："), txt(getScopesText(basicInfo.limitedScopes))]),
                      ]
                  })
              ]
          }),
          new TableRow({
              height: ROW_HEIGHT,
              children: [
                  new TableCell({ shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, children: [stdP([txt("2.實質性門檻")])] }),
                  new TableCell({ verticalAlign: VerticalAlign.CENTER, children: [stdP([txt(`依雙方協議訂為 ${basicInfo.materiality}`)])] })
              ]
          }),
          new TableRow({
              height: ROW_HEIGHT,
              children: [
                  new TableCell({ shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, children: [stdP([txt("3.基準年及基準年溫室氣體排放資訊")])] }),
                  new TableCell({ verticalAlign: VerticalAlign.CENTER, children: [
                      stdP([txt(`基準年設定為： ${basicInfo.baseYear} 年`)]),
                      stdP(txtCO2e(`總排放量： ${formatNumber(basicInfo.baseYearEmissions)} 公噸 `)),
                  ] })
              ]
          }),
          new TableRow({
              height: ROW_HEIGHT,
              children: [
                  new TableCell({ rowSpan: 8, shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, children: [stdP([txt("4.申請查驗年度及查驗年度溫室氣體排放資訊")])] }),
                  new TableCell({ verticalAlign: VerticalAlign.CENTER, children: [
                      stdP([txt(`查驗年度： ${basicInfo.verificationYear} 年`)]),
                      stdP(txtCO2e(`總排放量： ${formatNumber(total)} 公噸 `)),
                  ] })
              ]
          }),
          new TableRow({ height: ROW_HEIGHT, children: [new TableCell({ verticalAlign: VerticalAlign.CENTER, children: [stdP(txtCO2e(`類別一：排放量 ${formatNumber(emissions.cat1)} 公噸 `, `，佔總排放比例： ${getPercent(emissions.cat1)} %`))] })] }),
          new TableRow({ height: ROW_HEIGHT, children: [new TableCell({ verticalAlign: VerticalAlign.CENTER, children: [stdP(txtCO2e(`類別二：排放量 ${formatNumber(emissions.cat2)} 公噸 `, `，佔總排放比例： ${getPercent(emissions.cat2)} %`))] })] }),
          new TableRow({ height: ROW_HEIGHT, children: [new TableCell({ verticalAlign: VerticalAlign.CENTER, children: [stdP(txtCO2e(`類別三：排放量 ${formatNumber(emissions.cat3)} 公噸 `, `，佔總排放比例： ${getPercent(emissions.cat3)} %`))] })] }),
          new TableRow({ height: ROW_HEIGHT, children: [new TableCell({ verticalAlign: VerticalAlign.CENTER, children: [stdP(txtCO2e(`類別四：排放量 ${formatNumber(emissions.cat4)} 公噸 `, `，佔總排放比例： ${getPercent(emissions.cat4)} %`))] })] }),
          new TableRow({ height: ROW_HEIGHT, children: [new TableCell({ verticalAlign: VerticalAlign.CENTER, children: [stdP(txtCO2e(`類別五：排放量 ${formatNumber(emissions.cat5)} 公噸 `, `，佔總排放比例： ${getPercent(emissions.cat5)} %`))] })] }),
          new TableRow({ height: ROW_HEIGHT, children: [new TableCell({ verticalAlign: VerticalAlign.CENTER, children: [stdP(txtCO2e(`類別六：排放量 ${formatNumber(emissions.cat6)} 公噸 `, `，佔總排放比例： ${getPercent(emissions.cat6)} %`))] })] }),
          new TableRow({ height: ROW_HEIGHT, children: [new TableCell({ verticalAlign: VerticalAlign.CENTER, children: [stdP([txt(`盤查清冊之不確定性上、下限：上限 ${emissions.uncertaintyUpper} %，下限 ${emissions.uncertaintyLower} %`)])] })] }),
          new TableRow({
              height: ROW_HEIGHT,
              children: [
                  new TableCell({ shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, children: [stdP([txt("5.溫室氣體報告")]),stdP([txt("預期使用者")])] }),
                  new TableCell({ verticalAlign: VerticalAlign.CENTER, children: [stdP([txt(basicInfo.intendedUser)])] })
              ]
          }),
      ]
  });

  // --- Page 2 Content ---
  const fixedSpacing = { line: 520, lineRule: LineRuleCompat.EXACT, before: 0, after: 0 };

  const page2Intro = [
      stdP([txt("二、相關文件與標準條文對照審查", { bold: true, size: 14 })]),
      stdP([txt(`溫室氣體報告名稱/版次/日期： ${basicInfo.reportName}`)], "left", fixedSpacing),
      stdP([txt(`溫室氣體盤查清冊名稱/版次/日期： ${basicInfo.inventoryName}`)], "left", fixedSpacing),
      stdP([txt(`溫室氣體資訊管理程序名稱/版次/日期： ${basicInfo.procedureName}`)], "left", fixedSpacing),
      stdP([txt(" ")]), 
  ];

  const CL_WIDTH_CLAUSE = 3500;
  const CL_WIDTH_DOC = 4000;
  const CL_WIDTH_RESULT = 2165;
  const CL_WIDTH_ID = PORTRAIT_WIDTH - CL_WIDTH_CLAUSE - CL_WIDTH_DOC - CL_WIDTH_RESULT;

  const checklistHeaderRow = new TableRow({
      tableHeader: true,
      children: [
          new TableCell({ width: { size: CL_WIDTH_ID, type: WidthType.DXA }, shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, children: [stdP([txt("項目\n編號")], "center")] }),
          new TableCell({ width: { size: CL_WIDTH_CLAUSE, type: WidthType.DXA }, shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, children: [stdP([txt("ISO 14064-1: 2018 條文項目")], "center")] }),
          new TableCell({ width: { size: CL_WIDTH_DOC, type: WidthType.DXA }, shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, children: [stdP([txt("相關文件編號\n及對照章/節/頁")], "center")] }),
          new TableCell({ width: { size: CL_WIDTH_RESULT, type: WidthType.DXA }, shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, children: [
              stdP([txt("審 查 結 果")], "center"),
              stdP([txt("(符合項目以○註記，\n待釐清項目以 X 註記，\n不適用以―註記)", { size: 9 })], "center")
          ] }),
      ]
  });

  const checklistRows = checklist.map(item => {
      let mark = "";
      if (item.status === ComplianceStatus.COMPLIANT) mark = "○";
      else if (item.status === ComplianceStatus.CLARIFY) mark = "X";
      else mark = "―";

      const isMain = !item.id.includes('.'); 
      
      if (isMain) {
          return new TableRow({
              children: [
                  new TableCell({ 
                      width: { size: CL_WIDTH_ID, type: WidthType.DXA },
                      shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR },
                      children: [stdP([txt(item.id, { bold: true })], "center")] 
                  }),
                  new TableCell({ 
                      columnSpan: 3,
                      shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR },
                      children: [stdP([txt(item.name, { bold: true })])] 
                  })
              ]
          });
      }

      return new TableRow({
          children: [
              new TableCell({ children: [stdP([txt(item.id, { bold: isMain })], "center")] }),
              new TableCell({ children: [stdP([txt(item.name, { bold: isMain })])] }),
              new TableCell({ children: [stdP([txt(item.docRef)])] }),
              new TableCell({ verticalAlign: VerticalAlign.CENTER, children: [stdP([txt(mark)], "center")] }),
          ]
      });
  });

  const checklistTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      layout: TableLayoutType.FIXED,
      borders: { top: BORDER_STD, bottom: BORDER_STD, left: BORDER_STD, right: BORDER_STD, insideHorizontal: BORDER_STD, insideVertical: BORDER_STD },
      rows: [checklistHeaderRow, ...checklistRows]
  });

  // --- Page 3 Conclusion Content ---
  
  const conclusionTitle = stdP([txt("三、綜合結論", { bold: true, size: 14 })]);

  const conclusionTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: { top: BORDER_STD, bottom: BORDER_STD, left: BORDER_STD, right: BORDER_STD, insideHorizontal: BORDER_STD, insideVertical: BORDER_STD },
      rows: [
          new TableRow({
              children: [
                  new TableCell({ 
                      shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR },
                      children: [stdP([txt("(一)書面審查及/或訪談結果，對於查驗過程中是否可能與本機構之組織或個人存在潛在利益衝突", { size: 12 })])] 
                  })
              ]
          }),
          new TableRow({
              children: [
                  new TableCell({ 
                      children: [
                          stdP([
                              checkbox(conclusion.conflictOfInterest === 'No'), txt("否；"),
                              checkbox(conclusion.conflictOfInterest === 'Yes'), txt("是，說明如後："),
                              txt(conclusion.conflictOfInterest === 'Yes' ? conclusion.conflictDetail : "________________________")
                          ])
                      ] 
                  })
              ]
          }),
          new TableRow({
              children: [
                  new TableCell({ 
                      shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR },
                      children: [stdP([txt("(二)經書面審查及/或訪談結果，綜合結論如下：", { size: 12 })])] 
                  })
              ]
          }),
          new TableRow({
              height: { value: 2000, rule: HeightRule.ATLEAST }, 
              children: [
                  new TableCell({ 
                      children: [
                          stdP([
                              checkbox(conclusion.summary === FinalConclusion.PASS), 
                              txt("書面審查及/或訪談結果通過，可據以辦理後續第 1 階段查驗。")
                          ]),
                          stdP([
                              checkbox(conclusion.summary === FinalConclusion.REDUCED),
                              txt("組織層級查驗符合下述情況，最低現場查驗人天數得少於 4 人天(含)但不得少於 1 人天(含)。")
                          ]),
                          stdP([
                              txt(" 邊界及溫室氣體排放型態單純，如工廠或大樓為單一組織控制，邊界內未有涉及區域或樓層租借的狀況(即電力不須採用分配方式計算)、90 %以上溫室氣體排放量來自能源間接、無複雜之製程排放源…等情形。", { size: 10 })
                          ], "left", { before: 0, after: 100, line: 240 }, { left: 400 }),
                          stdP([
                              checkbox(conclusion.summary === FinalConclusion.PENDING),
                              txt("書面審查及/或訪談結果，尚有部分事項待釐清/補正。(參見待釐清/補正事項摘要表)")
                          ]),
                      ] 
                  })
              ]
          }),
           new TableRow({
                height: { value: 1500, rule: HeightRule.ATLEAST },
                children: [
                    new TableCell({
                        children: [stdP([txt(`其它：${conclusion.otherNote}`)])]
                    })
                ]
           })
      ]
  });

  // --- Page 4 Interview Content (CONDITIONAL) ---
  const sections: ISectionOptions[] = [
      // Section 1: Main Report
      {
        headers: { default: mainHeader },
        footers: { default: commonFooter },
        properties: {
            page: { margin: MARGINS_PORTRAIT }
        },
        children: [
            criteriaPara,
            clientInfoTable,
            section1Title,
            basicInfoTable,
            new Paragraph({ children: [], pageBreakBefore: true }),
            ...page2Intro,
            checklistTable,
            new Paragraph({ children: [] }),
            conclusionTitle,
            conclusionTable
        ]
      }
  ];

  // Section 2: Interview Report - Only add if interviews exist
  if (conclusion.interviews.length > 0) {
      const interviewHeader = [
          stdP([txt("查驗準則：ISO 14064-1(2018 年版)/CNS 14064-1(2021 年版)")]),
          stdP([txt(`現場訪談日期：     ${visitDate.y}    年     ${visitDate.m}    月     ${visitDate.d}    日`)]),
      ];

      const interviewTableRows = [
          new TableRow({
              tableHeader: true,
              children: [
                  new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, children: [stdP([txt("系統編號")], "center")] }),
                  new TableCell({ width: { size: 2400, type: WidthType.DXA }, shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, children: [stdP([txt("現場訪談事項")], "center")] }),
                  new TableCell({ width: { size: 5505, type: WidthType.DXA }, shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, children: [stdP([txt("查 核 紀 錄")], "center")] }),
                  new TableCell({ width: { size: 1500, type: WidthType.DXA }, shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, children: [stdP([txt("結 果")], "center")] }),
              ]
          }),
          ...conclusion.interviews.map((iv, i) => new TableRow({
              height: { value: 1800, rule: HeightRule.ATLEAST },
              children: [
                  new TableCell({ children: [stdP([txt(`${i + 1}`)], "center")] }),
                  new TableCell({ children: [stdP([txt(iv.topic)])] }),
                  new TableCell({ children: [stdP([txt(iv.record)])] }),
                  new TableCell({ children: [stdP([txt(iv.result)])] }),
              ]
          }))
      ];

      const interviewTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          borders: { top: BORDER_STD, bottom: BORDER_STD, left: BORDER_STD, right: BORDER_STD, insideHorizontal: BORDER_STD, insideVertical: BORDER_STD },
          rows: interviewTableRows
      });

      const interviewSigTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: { top: BORDER_NONE, left: BORDER_STD, right: BORDER_STD, bottom: BORDER_STD, insideVertical: BORDER_STD, insideHorizontal: BORDER_NONE },
          rows: [
            new TableRow({
                height: { value: 1200, rule: HeightRule.ATLEAST },
                children: [
                    new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, children: [stdP([txt("查驗員")], "center"),stdP([txt("簽 名")], "center")] }),
                    new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, children: [] }), 
                    new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, children: [stdP([txt("主導查驗員")], "center"),stdP([txt("簽 名")], "center")] }),
                    new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, children: [] }) 
                ]
            })
          ]
      });

      const interviewNote = stdP([txt("註：本頁為附件，由查驗人員視實際需要調整頁數及行距。", { size: 10 })]);

      sections.push({
        headers: { default: appendixHeader },
        footers: { default: commonFooter },
        properties: {
            type: SectionType.NEXT_PAGE,
            page: { margin: MARGINS_PORTRAIT }
        },
        children: [
            ...interviewHeader,
            interviewTable,
            interviewSigTable,
            interviewNote
        ]
      });
  }

  // --- Page 5 Pending Items Content ---
  const pendingHeaderText = [
    stdP([txt("待釐清/補正事項摘要表", { size: 14, bold: true })], "center"),
    stdP([txt("查驗準則：ISO 14064-1(2018 年版)/CNS 14064-1(2021 年版)")]),
    stdP([txt(`書面審查日期：     ${reviewDate.y}    年     ${reviewDate.m}    月     ${reviewDate.d}    日`)]),
  ];

  // Pending Table - Render only if exists or minimal 1 row if preferred, but user said "only if additions"
  // If list is empty, we still render the table structure but empty, or just one empty row. 
  // User also said "Pending items (implied dynamic)". Let's map exactly what's there. 
  // If empty, maybe show one empty row just to show the table exists? Or just header?
  // Let's stick to mapping data exactly.
  
  const pendingRows = conclusion.pendingItems.length > 0 ? conclusion.pendingItems.map((pi, i) => new TableRow({
        height: { value: 1800, rule: HeightRule.ATLEAST },
        children: [
            new TableCell({ children: [stdP([txt(`${i + 1}`)], "center")] }),
            new TableCell({ children: [stdP([txt(pi.content)])] }),
            new TableCell({ children: [stdP([txt(pi.response)])] }),
        ]
    })) : [
        new TableRow({
            height: { value: 1800, rule: HeightRule.ATLEAST },
            children: [
                new TableCell({ children: [stdP([txt(" ")], "center")] }),
                new TableCell({ children: [stdP([txt(" ")])] }),
                new TableCell({ children: [stdP([txt(" ")])] }),
            ]
        })
    ];

  const pendingTableRows = [
    new TableRow({
        tableHeader: true,
        children: [
            new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, children: [stdP([txt("項次")], "center")] }),
            new TableCell({ width: { size: 4602, type: WidthType.DXA }, shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, children: [stdP([txt("待釐清/補正事項事項內容")], "center")] }),
            new TableCell({ width: { size: 4603, type: WidthType.DXA }, shading: { fill: SHADING_COLOR, type: ShadingType.CLEAR }, children: [stdP([txt("組織回覆")], "center")] }),
        ]
    }),
    ...pendingRows
  ];

  const pendingTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      layout: TableLayoutType.FIXED,
      borders: { top: BORDER_STD, bottom: BORDER_STD, left: BORDER_STD, right: BORDER_STD, insideHorizontal: BORDER_STD, insideVertical: BORDER_STD },
      rows: pendingTableRows
  });

  const pendingSigTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: { top: BORDER_NONE, left: BORDER_STD, right: BORDER_STD, bottom: BORDER_STD, insideVertical: BORDER_STD, insideHorizontal: BORDER_STD },
      rows: [
        new TableRow({
            height: { value: 1500, rule: HeightRule.ATLEAST },
            children: [
                new TableCell({ width: { size: 33, type: WidthType.PERCENTAGE }, children: [
                    stdP([txt("填報之查驗人員簽名:")]),
                    new Paragraph({ children: [], spacing: { before: 200 } }),
                    stdP([txt(conclusion.verifierName, { size: 16 })], "center")
                ] }),
                new TableCell({ width: { size: 33, type: WidthType.PERCENTAGE }, children: [
                    stdP([txt("主導查驗員簽名:")]),
                    new Paragraph({ children: [], spacing: { before: 200 } }),
                    stdP([txt(conclusion.leadVerifierName, { size: 16 })], "center")
                ] }),
                new TableCell({ width: { size: 34, type: WidthType.PERCENTAGE }, children: [
                    stdP([txt("委託單位代表簽名:")]),
                    new Paragraph({ children: [], spacing: { before: 200 } }),
                    stdP([txt(conclusion.clientRepName, { size: 16 })], "center")
                ] })
            ]
        }),
        new TableRow({
            children: [
                new TableCell({ columnSpan: 3, borders: { top: BORDER_STD, bottom: BORDER_STD, left: BORDER_STD, right: BORDER_STD }, children: [
                    stdP([txt("備註:", { bold: true })]),
                    stdP([
                        checkbox(conclusion.memoCorrection), 
                        txt("書面審查及/或訪談結果，尚有上述事項待釐清/補正，請於第一階段(S-1)前說明或提送補正資料。", { bold: true })
                    ])
                ] })
            ]
        })
      ]
  });

  // Add Pending Items Section
  sections.push({
    headers: { default: appendixHeader },
    footers: { default: commonFooter },
    properties: {
        type: SectionType.NEXT_PAGE,
        page: { margin: MARGINS_PORTRAIT }
    },
    children: [
        ...pendingHeaderText,
        pendingTable,
        pendingSigTable
    ]
  });

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: FONTS_OPTS, size: 24 }, // 12pt
          paragraph: { spacing: { before: 40, after: 40 } }
        }
      }
    },
    sections: sections
  });

  return await Packer.toBlob(doc);
};