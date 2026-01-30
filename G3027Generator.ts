
import { AppState } from './types';
import { 
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
  WidthType, BorderStyle, Header, Footer, 
  VerticalAlign, ShadingType, HeightRule, TableLayoutType, SectionType, PageOrientation, PageNumber
} from 'docx';

// ════════════════════════════════════════════════════════════════════════════
// UNITS & DIMENSIONS
// ════════════════════════════════════════════════════════════════════════════
const LANDSCAPE_WIDTH = 14700; 
const PORTRAIT_WIDTH = 10000;

const MARGINS_LANDSCAPE = { 
    top: 1247, 
    bottom: 720, 
    left: 720, 
    right: 720,
    header: 284,
    footer: 624
};

const MARGINS_PORTRAIT = { 
    top: 720, 
    bottom: 720, 
    left: 720, 
    right: 720,
    header: 284,
    footer: 624
};

// ════════════════════════════════════════════════════════════════════════════
// STYLES
// ════════════════════════════════════════════════════════════════════════════
const FONTS_OPTS = {
  ascii: "Times New Roman",
  hAnsi: "Times New Roman",
  eastAsia: "標楷體",
  cs: "Times New Roman",
  hint: "eastAsia" as const
};

const BORDER_THICK = { style: BorderStyle.SINGLE, size: 18, color: "000000" }; 
const BORDER_THIN = { style: BorderStyle.SINGLE, size: 4, color: "000000" };   
const BORDER_NONE = { style: BorderStyle.NONE, size: 0, color: "auto" };
const BORDER_INVISIBLE = { style: BorderStyle.NIL, size: 0, color: "auto" };

const SHADING_GRAY = "D9D9D9"; 

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
    size: (opts.size || 12) * 2, 
    underline: opts.underline ? { type: 'single' } : undefined,
  }));
  return runs;
};

const checkbox = (checked: boolean, sizePt: number = 12) => {
  return new TextRun({
    text: checked ? "☑" : "□",
    font: FONTS_OPTS,
    size: sizePt * 2
  });
};

const stdP = (children: any[], align: string = "left", spacing?: { before?: number, after?: number, line?: number }) => {
  return new Paragraph({
    children: children.flat(),
    alignment: align as any,
    spacing: spacing || { before: 20, after: 20, line: 240 }
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

// ════════════════════════════════════════════════════════════════════════════
// GENERATOR
// ════════════════════════════════════════════════════════════════════════════
export const generateG3027Docx = async (data: AppState['g3027']): Promise<Blob> => {
  const { basicInfo, findings, stats, conclusion } = data;
  const date = formatDate(basicInfo.date);
  const auditeeDate = formatDate(conclusion.auditeeDate);
  const verifierDate = formatDate(conclusion.verifierDate);
  
  // Logic: Filter findings based on current stage
  const currentStage = basicInfo.stage;
  const filteredFindings = findings.filter(f => f.stage === currentStage);
  
  // Requirement 9: If S2 and no items, we might want to skip the page.
  // But typically a report still needs the section header or an "Empty" table.
  // The request says "Report output can delete this page". 
  // So if S2 and empty -> no findings section.
  const shouldRenderFindingsPage = !(currentStage === 'S2' && filteredFindings.length === 0);

  const headerContent = [
      stdP([txt("旭威認證股份有限公司 查驗機構", { size: 18, bold: true })], "center"),
      stdP([
          txt("不符合事項/觀察事項摘要表 ", { size: 16, bold: true }),
          checkbox(basicInfo.stage === 'S1', 16), txt("第一階段(S-1) ", { size: 16 }),
          checkbox(basicInfo.stage === 'S2', 16), txt("第二階段(S-2)", { size: 16 })
      ], "center"),
  ];

  const commonHeader = new Header({ children: headerContent });

  const footerContent = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
          top: BORDER_NONE,
          bottom: BORDER_NONE,
          left: BORDER_NONE,
          right: BORDER_NONE,
          insideHorizontal: BORDER_NONE,
          insideVertical: BORDER_NONE,
      },
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
                  new TableCell({ width: { size: 33, type: WidthType.PERCENTAGE }, children: [stdP([txt("G-3027", { size: 10 })], "right")] }),
              ]
          })
      ]
  });

  const commonFooter = new Footer({ children: [footerContent] });

  // ==========================================================================
  // PAGE 1: FINDINGS (LANDSCAPE)
  // ==========================================================================
  
  const basicInfoTable = new Table({
      width: { size: LANDSCAPE_WIDTH, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      rows: [
          new TableRow({
              children: [
                  new TableCell({
                      width: { size: 4500, type: WidthType.DXA },
                      borders: { top: BORDER_INVISIBLE, left: BORDER_INVISIBLE, right: BORDER_INVISIBLE, bottom: BORDER_INVISIBLE },
                      children: [stdP([txt(`案件編號：${basicInfo.caseNumber}`)])]
                  }),
                  new TableCell({
                      width: { size: 6200, type: WidthType.DXA }, 
                      borders: { top: BORDER_INVISIBLE, left: BORDER_INVISIBLE, right: BORDER_INVISIBLE, bottom: BORDER_INVISIBLE },
                      children: [stdP([txt(" ")])]
                  }),
                  new TableCell({
                      width: { size: 4000, type: WidthType.DXA },
                      borders: { top: BORDER_INVISIBLE, left: BORDER_INVISIBLE, right: BORDER_INVISIBLE, bottom: BORDER_INVISIBLE },
                      children: [stdP([txt(`查驗年度： ${basicInfo.verificationYear} 年`)])] 
                  })
              ]
          }),
          new TableRow({
              children: [
                  new TableCell({
                      borders: { top: BORDER_INVISIBLE, left: BORDER_INVISIBLE, right: BORDER_INVISIBLE, bottom: BORDER_INVISIBLE },
                      children: [stdP([txt(`主導查驗員：${basicInfo.leadVerifier}`)])]
                  }),
                  new TableCell({
                      borders: { top: BORDER_INVISIBLE, left: BORDER_INVISIBLE, right: BORDER_INVISIBLE, bottom: BORDER_INVISIBLE },
                      children: [stdP([txt(`受查驗方代表：${basicInfo.auditeeRep}`)])]
                  }),
                  new TableCell({
                      borders: { top: BORDER_INVISIBLE, left: BORDER_INVISIBLE, right: BORDER_INVISIBLE, bottom: BORDER_INVISIBLE },
                      children: [stdP([txt(`查驗日期： ${date.y} 年 ${date.m} 月 ${date.d} 日`)])]
                  })
              ]
          })
      ]
  });

  const findingsHeader = [
      new TableRow({
          children: [
              new TableCell({ rowSpan: 2, width: { size: 600, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, children: [stdP([txt("編號")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THICK, bottom: BORDER_THICK, left: BORDER_THICK, right: BORDER_THIN } }),
              new TableCell({ columnSpan: 3, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, children: [stdP([txt("查驗機構查驗發現")], "center")], borders: { top: BORDER_THICK, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, children: [stdP([txt("受查驗方回覆")], "center")], borders: { top: BORDER_THICK, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ columnSpan: 4, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, children: [stdP([txt("查驗機構審查")], "center")], borders: { top: BORDER_THICK, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THICK } }),
          ]
      }),
      new TableRow({
          children: [
              new TableCell({ width: { size: 1400, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, children: [stdP([txt("發現事項之分類")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THICK, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ width: { size: 3000, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, children: [stdP([txt("不符合事項描述")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THICK, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ width: { size: 1000, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, children: [stdP([txt("填報查\n驗人員")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THICK, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ width: { size: 3000, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, children: [stdP([txt("矯正措施/澄清說明")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THICK, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ width: { size: 3000, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, children: [stdP([txt("審查意見")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THICK, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ width: { size: 1200, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, children: [stdP([txt("查驗員")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THICK, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ width: { size: 1100, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, children: [stdP([txt("審查結果")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THICK, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ width: { size: 1100, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, children: [stdP([txt("審查地點")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THICK, left: BORDER_THIN, right: BORDER_THICK } }),
          ]
      })
  ];

  const findingRows = [];
  const minRows = 3; 
  // Use filteredFindings here
  const totalRows = Math.max(filteredFindings.length, minRows);

  for (let i = 0; i < totalRows; i++) {
      const item = filteredFindings[i] || { id: '', type: '', description: '', reporter: '', correctiveAction: '', reviewOpinion: '', reviewer: '', result: '', location: '' };
      
      const resultCell = stdP([
          checkbox(item.result === 'Close', 12), txt(" 結案"),
          new TextRun({ text: "\n" }),
          checkbox(item.result === 'Keep', 12), txt(" 保留")
      ]);

      const locCell = stdP([
          checkbox(item.location === 'OnSite', 12), txt(" 現場"),
          new TextRun({ text: "\n" }),
          checkbox(item.location === 'OffSite', 12), txt(" 非現場")
      ]);

      findingRows.push(new TableRow({
          height: { value: 900, rule: HeightRule.ATLEAST },
          children: [
              new TableCell({ children: [stdP([txt(i < filteredFindings.length ? `${i + 1}` : "")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THICK, right: BORDER_THIN } }),
              new TableCell({ children: [stdP([txt(item.type)], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ children: [stdP([txt(item.description)])], borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ children: [stdP([txt(item.reporter)], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ children: [stdP([txt(item.correctiveAction)])], borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ children: [stdP([txt(item.reviewOpinion)])], borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ children: [stdP([txt(item.reviewer)], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ children: [resultCell], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
              new TableCell({ children: [locCell], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THICK } }),
          ]
      }));
  }

  const footerRows = [
      new TableRow({
          children: [
              new TableCell({
                  columnSpan: 9,
                  shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR },
                  borders: { top: BORDER_THICK, bottom: BORDER_NONE, left: BORDER_THICK, right: BORDER_THICK },
                  children: [stdP([txt("發現事項之分類：", { size: 12 })]),
                             stdP([txt("矯正措施要求(Corrective Action Request，CAR)、澄清要求(Clarification Request，CR)與後續行動要求(Forward Action Request，FAR)", { size: 12 })])]
              })
          ]
      }),
      new TableRow({
          children: [
              new TableCell({
                  columnSpan: 9,
                  borders: { top: BORDER_NONE, bottom: BORDER_NONE, left: BORDER_THICK, right: BORDER_THICK },
                  children: [
                      stdP([txt("備註:", { size: 12 })]),
                      stdP([txt("(S-1 適用)請於第 2 階段(S-2)查驗前提送不符合事項及觀察事項因應處理方案。", { size: 12 })]),
                      stdP([txt("(S-2 適用)請提送修正後經核章之溫室氣體報告書及不符合事項及觀察事項因應處理方案)依序表列彙整， 並於 10 日內以電子郵件逕寄旭威認證股份有限公司查驗機構窗口，未送達前不予複審。", { size: 12 })]),
                      stdP([txt("➢ 矯正措施要求：若不符合相關規定、構成實質差異、個別或累積之錯誤、遺漏及誤導構成實質性之部分，查驗人員應對受查驗者提出此要求，請受查驗者進行矯正。", { size: 10 })]),
                      stdP([txt("➢ 澄清要求：若資訊不夠充分或不明確，無法確定是否符合相關規定時，查驗人員應對受查驗者提出此要求，請受查驗者提出說明以澄清。", { size: 10 })]),
                      stdP([txt("➢ 後續行動要求：針對下個查驗期間溫室氣體數據蒐集及報告特別注意或調整的部分，提出此要求。(*後續行動要求無須填寫矯正措施說明。)", { size: 10 })]),
                      stdP([txt("正本：旭威認證股份有限公司 查證機構存檔；影本：廠商存參", { size: 10 })])
                  ]
              })
          ]
      }),
      new TableRow({
          children: [
              new TableCell({
                  columnSpan: 9,
                  borders: { top: BORDER_NONE, bottom: BORDER_THICK, left: BORDER_THICK, right: BORDER_THICK },
                  children: [
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: { top: BORDER_NONE, bottom: BORDER_NONE, left: BORDER_NONE, right: BORDER_NONE, insideVertical: BORDER_NONE, insideHorizontal: BORDER_NONE },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({ children: [stdP([txt(`受查驗方代表：${basicInfo.auditeeRep}`, { size: 12 })])] }),
                                    new TableCell({ children: [stdP([txt(`回覆日期：${auditeeDate.y}年${auditeeDate.m}月${auditeeDate.d}日`, { size: 12 })])] }),
                                    new TableCell({ children: [stdP([txt(`主導查驗員：${basicInfo.leadVerifier}`, { size: 12 })])] }),
                                    new TableCell({ children: [stdP([txt(`審查日期：${date.y}年${date.m}月${date.d}日`, { size: 12 })])] }),
                                ]
                            })
                        ]
                    })
                  ]
              })
          ]
      })
  ];

  // ==========================================================================
  // PAGE 2: SUMMARY & CONCLUSIONS (PORTRAIT)
  // ==========================================================================

  const p2Header = [
      stdP([txt(`案件編號：${basicInfo.caseNumber}`, { size: 12 })]),
      stdP([txt("本階段現場查證結果(以下內容參照不符合事項/觀察事項摘要表)如下：", { size: 12 })]),
  ];

  const statsRows = [
      new TableRow({
          children: [
              new TableCell({ 
                  width: { size: 20, type: WidthType.PERCENTAGE }, 
                  children: [stdP([txt("不符合事項數量")], "center")], 
                  verticalAlign: VerticalAlign.CENTER, 
                  borders: { top: BORDER_THICK, left: BORDER_THICK, right: BORDER_THIN, bottom: BORDER_THIN } 
              }),
              new TableCell({ 
                  width: { size: 20, type: WidthType.PERCENTAGE }, 
                  children: [stdP([txt("第一階段(S-1)")], "center")], 
                  verticalAlign: VerticalAlign.CENTER, 
                  borders: { top: BORDER_THICK, left: BORDER_THIN, right: BORDER_THIN, bottom: BORDER_THIN } 
              }),
              new TableCell({ 
                  width: { size: 20, type: WidthType.PERCENTAGE }, 
                  children: [stdP([txt(stats.s1.nonConformity)], "center")], 
                  verticalAlign: VerticalAlign.CENTER, 
                  borders: { top: BORDER_THICK, left: BORDER_THIN, right: BORDER_THIN, bottom: BORDER_THIN } 
              }),
              new TableCell({ 
                  width: { size: 20, type: WidthType.PERCENTAGE }, 
                  children: [stdP([txt("第二階段(S-2)")], "center")], 
                  verticalAlign: VerticalAlign.CENTER, 
                  borders: { top: BORDER_THICK, left: BORDER_THIN, right: BORDER_THIN, bottom: BORDER_THIN } 
              }),
              new TableCell({ 
                  width: { size: 20, type: WidthType.PERCENTAGE }, 
                  children: [stdP([txt(stats.s2.nonConformity)], "center")], 
                  verticalAlign: VerticalAlign.CENTER, 
                  borders: { top: BORDER_THICK, left: BORDER_THIN, right: BORDER_THICK, bottom: BORDER_THIN } 
              }),
          ]
      }),
      new TableRow({
        children: [
            new TableCell({ children: [stdP([txt("觀察事項數量")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, left: BORDER_THICK, right: BORDER_THIN, bottom: BORDER_THIN } }),
            new TableCell({ children: [stdP([txt("第一階段(S-1)")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN, bottom: BORDER_THIN } }),
            new TableCell({ children: [stdP([txt(stats.s1.observation)], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN, bottom: BORDER_THIN } }),
            new TableCell({ children: [stdP([txt("第二階段(S-2)")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN, bottom: BORDER_THIN } }),
            new TableCell({ children: [stdP([txt(stats.s2.observation)], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, left: BORDER_THIN, right: BORDER_THICK, bottom: BORDER_THIN } }),
        ]
    }),
    new TableRow({
        children: [
            new TableCell({ children: [stdP([txt("建議事項數量")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, left: BORDER_THICK, right: BORDER_THIN, bottom: BORDER_THICK } }),
            new TableCell({ children: [stdP([txt("第一階段(S-1)")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN, bottom: BORDER_THICK } }),
            new TableCell({ children: [stdP([txt(stats.s1.suggestion)], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN, bottom: BORDER_THICK } }),
            new TableCell({ children: [stdP([txt("第二階段(S-2)")], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN, bottom: BORDER_THICK } }),
            new TableCell({ children: [stdP([txt(stats.s2.suggestion)], "center")], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, left: BORDER_THIN, right: BORDER_THICK, bottom: BORDER_THICK } }),
        ]
    })
  ];

  const conclusionRows = [
      new TableRow({
          children: [
              new TableCell({
                  columnSpan: 3,
                  width: { size: 60, type: WidthType.PERCENTAGE },
                  borders: { top: BORDER_NONE, left: BORDER_THICK, right: BORDER_THIN, bottom: BORDER_THIN },
                  margins: { top: 100, bottom: 100, left: 100, right: 100 },
                  children: [
                      stdP([txt("第一階段適用:")]),
                      stdP([checkbox(conclusion.s1Result === 'None'), txt("未發現相關問題，按原訂計畫執行第二階段查證。")]),
                      stdP([checkbox(conclusion.s1Result === 'NoEffect'), txt("本階段所發現之問題不影響第二階段查證。")]),
                      stdP([checkbox(conclusion.s1Result === 'AdjustDays'), txt("現場查證人天或第二階段查證日期需調節。")]),
                      stdP([txt(` 說明：${conclusion.s1Note || ''}`)]),
                      stdP([checkbox(conclusion.s1Result === 'Undecided'), txt("目前的情況無法決定。")]),
                  ]
              }),
              new TableCell({
                columnSpan: 2,
                width: { size: 40, type: WidthType.PERCENTAGE },
                borders: { top: BORDER_NONE, left: BORDER_THIN, right: BORDER_THICK, bottom: BORDER_THIN },
                margins: { top: 100, bottom: 100, left: 100, right: 100 },
                children: [
                    stdP([txt("第二階段適用:")]),
                    stdP([txt("不符合事項與觀察事項，是否已於第二階段前完成改正")]),
                    stdP([checkbox(conclusion.s2Result === 'Corrected'), txt("是 "), checkbox(conclusion.s2Result === 'Agree'), txt("否(待組織回覆矯正措施審查)")]),
                    stdP([checkbox(conclusion.s2Result === 'NoFindings'), txt("無相關發現")]),
                ]
            })
          ]
      }),
      new TableRow({
          children: [
              new TableCell({ width: { size: 20, type: WidthType.PERCENTAGE }, children: [stdP([txt("查證協議資訊變更")])], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THICK, right: BORDER_THIN } }),
              new TableCell({ 
                  columnSpan: 4,
                  width: { size: 80, type: WidthType.PERCENTAGE },
                  children: [
                      stdP([checkbox(conclusion.protocolChange === 'No'), txt("無變更")]),
                      stdP([checkbox(conclusion.protocolChange === 'Yes'), txt("查證協議變更")]),
                      stdP([txt(` 請說明：${conclusion.protocolChangeNote}`)])
                  ], 
                  borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THICK } 
               })
          ]
      }),
      new TableRow({
        children: [
            new TableCell({ width: { size: 20, type: WidthType.PERCENTAGE }, children: [stdP([txt("保留意見")])], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THICK, right: BORDER_THIN } }),
            new TableCell({ columnSpan: 4, width: { size: 80, type: WidthType.PERCENTAGE }, children: [stdP([txt(conclusion.reservedOpinion)])], borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THICK } })
        ]
      }),
      new TableRow({
        children: [
            new TableCell({ width: { size: 20, type: WidthType.PERCENTAGE }, children: [stdP([txt("其他說明")])], verticalAlign: VerticalAlign.CENTER, borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THICK, right: BORDER_THIN } }),
            new TableCell({ columnSpan: 4, width: { size: 80, type: WidthType.PERCENTAGE }, children: [stdP([txt(conclusion.otherNote)])], borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THICK } })
        ]
      }),
  ];

  const auditeeRows = [
      new TableRow({
        children: [
            new TableCell({
                columnSpan: 5,
                borders: { top: BORDER_THICK, bottom: BORDER_THICK, left: BORDER_THICK, right: BORDER_THICK },
                margins: { top: 100, bottom: 100, left: 100, right: 100 },
                children: [
                    stdP([txt("受查組織", { bold: true })]),
                    stdP([txt("組織代表瞭解並接受查證結果以及不符合報告的內容。組織代表亦能陳述對此次查證不滿意之處。")]),
                    stdP([txt("組織代表簽名處：", { bold: true })]),
                    new Paragraph({ children: [], spacing: { before: 400, after: 400 } }), 
                    stdP([txt(`日期：${auditeeDate.y}年${auditeeDate.m}月${auditeeDate.d}日`)], "right")
                ]
            })
        ]
      })
  ];

  const verifierRows = [
      new TableRow({
        children: [
            new TableCell({
                columnSpan: 5,
                borders: { top: BORDER_THICK, bottom: BORDER_THICK, left: BORDER_THICK, right: BORDER_THICK },
                margins: { top: 100, bottom: 100, left: 100, right: 100 },
                children: [
                    stdP([txt("查證機構", { bold: true })]),
                    stdP([txt("考慮到文件呈現方式、查證的場址，以及對問題的回應，主導查證員的簽名並不表示查證小組的查證人員或查證機構須負責意外事件或在查證程序發生後其客戶所造成的錯誤。")]),
                    stdP([txt("此階段查證機構人員簽名處：", { bold: true })]),
                    stdP([txt("(查證小組/技術專家/觀察員/見證員)", { bold: true })]),
                    new Paragraph({ children: [], spacing: { before: 400, after: 400 } }), 
                    stdP([txt(`日期：${verifierDate.y}年${verifierDate.m}月${verifierDate.d}日`)], "right")
                ]
            })
        ]
      })
  ];

  const summaryTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      layout: TableLayoutType.FIXED, 
      rows: [
          ...statsRows,
          ...conclusionRows,
          ...auditeeRows,
          ...verifierRows
      ]
  });

  const sectionArray: any[] = [];

  // Conditional: Only add Findings Page if it's not (S2 and Empty)
  if (shouldRenderFindingsPage) {
      sectionArray.push({
        headers: { default: commonHeader },
        footers: { default: commonFooter },
        properties: {
          page: { 
            size: { orientation: PageOrientation.LANDSCAPE },
            margin: MARGINS_LANDSCAPE 
          }
        },
        children: [
            basicInfoTable,
            new Paragraph({ children: [] }), 
            new Table({
                width: { size: LANDSCAPE_WIDTH, type: WidthType.DXA },
                layout: TableLayoutType.FIXED,
                rows: [...findingsHeader, ...findingRows, ...footerRows]
            })
        ]
      });
  }

  // Always add Conclusion Page
  sectionArray.push({
        headers: { default: commonHeader }, 
        footers: { default: commonFooter },
        properties: {
          type: SectionType.NEXT_PAGE,
          page: { 
            size: { orientation: PageOrientation.PORTRAIT },
            margin: MARGINS_PORTRAIT 
          }
        },
        children: [
            ...p2Header,
            new Paragraph({ children: [] }), 
            summaryTable
        ]
  });

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: FONTS_OPTS, size: 24 }, 
          paragraph: { spacing: { before: 40, after: 40 } }
        }
      }
    },
    sections: sectionArray
  });

  return await Packer.toBlob(doc);
};