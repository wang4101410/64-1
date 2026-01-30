import { AppState, ComplianceStatus } from './types';
import { 
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
  WidthType, BorderStyle, Header, Footer, 
  VerticalAlign, ShadingType, HeightRule, TableLayoutType, SectionType, TextDirection, PageNumber, AlignmentType, PageOrientation
} from 'docx';

// ════════════════════════════════════════════════════════════════════════════
// STABILITY LAYER
// ════════════════════════════════════════════════════════════════════════════
const LineRuleCompat = {
  AT_LEAST: "atLeast",
  EXACT: "exact",
  AUTO: "auto",
} as const;

// ════════════════════════════════════════════════════════════════════════════
// LAYOUT PHYSICS & UNITS
// ════════════════════════════════════════════════════════════════════════════
// Total Page Width: 11906 DXA (21cm)
// Margins: Top/Bottom 2cm (1134), Left/Right 1.5cm (851)
// Table Width: 10093 DXA
const TABLE_WIDTH_DXA = 10093;

// Landscape A4: 29.7cm width (~16838 DXA)
// Printable width = 16838 - 851 - 851 = 15136
const LANDSCAPE_TABLE_WIDTH_DXA = 15136;

const MARGIN_TOP = 1134; // 2cm
const MARGIN_BOTTOM = 1134; // 2cm
const MARGIN_LEFT = 851; // 1.5cm
const MARGIN_RIGHT = 851; // 1.5cm
const MARGIN_HEADER = 425; // 0.75cm
const MARGIN_FOOTER = 425;

// ════════════════════════════════════════════════════════════════════════════
// TYPOGRAPHY & FONTS
// ════════════════════════════════════════════════════════════════════════════
const FONTS_OPTS = {
  ascii: "Times New Roman",
  hAnsi: "Times New Roman",
  eastAsia: "標楷體",
  cs: "Times New Roman",
  hint: "eastAsia" as const
};

const BASE_SIZE = 24; // 12pt

// ════════════════════════════════════════════════════════════════════════════
// PARAGRAPH SPACING
// ════════════════════════════════════════════════════════════════════════════
const STRICT_SPACING = {
  before: 0,
  after: 0,
  line: 360, // 18pt
  lineRule: LineRuleCompat.AT_LEAST
};

// ════════════════════════════════════════════════════════════════════════════
// BORDER SYSTEM
// ════════════════════════════════════════════════════════════════════════════
const BORDER_THICK_VAL = 18; // 1.5pt
const BORDER_THIN_VAL = 12; // 1pt

const BORDER_THICK = { style: BorderStyle.SINGLE, size: BORDER_THICK_VAL, color: "000000" };
const BORDER_THIN = { style: BorderStyle.SINGLE, size: BORDER_THIN_VAL, color: "000000" };
const BORDER_NONE = { style: BorderStyle.NONE, size: 0, color: "auto" };

// ════════════════════════════════════════════════════════════════════════════
// CELL MARGINS
// ════════════════════════════════════════════════════════════════════════════
const CELL_MARGINS = {
  top: 50,
  bottom: 50,
  left: 50,
  right: 50,
};

const SHADING_GRAY = "BFBFBF";
const SHADING_LIGHT_GRAY = "D9D9D9";

// ════════════════════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════════════════════

const checkbox = (checked: boolean, sizePt: number = 12) => {
  return new TextRun({
    text: checked ? "☑" : "□",
    font: FONTS_OPTS,
    size: sizePt * 2
  });
};

const txt = (text: string, opts: { bold?: boolean; size?: number; underline?: boolean; subScript?: boolean; color?: string; position?: number } = {}) => {
  return new TextRun({
    text: text,
    font: FONTS_OPTS,
    bold: opts.bold,
    size: (opts.size || 12) * 2,
    underline: opts.underline ? { type: 'single' } : undefined,
    subScript: opts.subScript,
    color: opts.color || "000000",
    position: opts.position ? `${opts.position / 2}pt` : undefined
  });
};

// Helper to parse text and subscript "2" in "CO2e"
const parseChecklistName = (text: string, sizePt: number = 10, bold: boolean = false) => {
    const parts = text.split("CO2e");
    if (parts.length === 1) return [txt(text, { size: sizePt, bold })];
    
    const runs: TextRun[] = [];
    parts.forEach((part, i) => {
        if (part) runs.push(txt(part, { size: sizePt, bold }));
        if (i < parts.length - 1) {
            runs.push(new TextRun({ text: "CO", font: FONTS_OPTS, size: sizePt * 2, bold })); 
            runs.push(new TextRun({ text: "2", font: FONTS_OPTS, size: sizePt * 2, subScript: true, bold }));
            runs.push(new TextRun({ text: "e", font: FONTS_OPTS, size: sizePt * 2, bold }));
        }
    });
    return runs;
};

const stdP = (children: any[], align: string = "left", spacing = STRICT_SPACING) => {
  return new Paragraph({
    children: children,
    alignment: align as any,
    spacing: spacing
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

export const generateG3026Docx = async (data: AppState['g3026'], g3022Data?: AppState['g3022']): Promise<Blob> => {
  const { basicInfo, checklist, samplingResults, emissionFactors, leadVerifierName, otherObservation } = data;
  const checkDate = formatDate(basicInfo.checkDate);

  // --- Header & Footer ---
  const header = new Header({
    children: [
        stdP([txt("旭威認證股份有限公司 查驗機構", { size: 18, bold: true })], "center"),
        stdP([
            checkbox(basicInfo.stage === 'S1', 16), txt("S1", { size: 16 }), 
            txt("  "),
            checkbox(basicInfo.stage === 'S2', 16), txt("S2", { size: 16 }), 
            txt(" 查驗觀察報告", { size: 16 })
        ], "center"),
        stdP([txt(`案件編號：${basicInfo.caseNumber}`, { underline: true })], "left")
    ]
  });

  const footer = new Footer({
    children: [
        new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: { top: BORDER_NONE, bottom: BORDER_NONE, left: BORDER_NONE, right: BORDER_NONE, insideHorizontal: BORDER_NONE, insideVertical: BORDER_NONE },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({ children: [stdP([txt("1141017 版", { size: 10 })])], width: { size: 33, type: WidthType.PERCENTAGE } }),
                        new TableCell({ children: [
                             stdP([
                                new TextRun({ children: [PageNumber.CURRENT], size: 20, font: FONTS_OPTS }), 
                                txt(" / ", { size: 10 }),
                                new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 20, font: FONTS_OPTS }),
                             ], "center")
                        ], width: { size: 33, type: WidthType.PERCENTAGE } }),
                        new TableCell({ children: [stdP([txt("G-3026", { size: 10 })], "right")], width: { size: 33, type: WidthType.PERCENTAGE } }),
                    ]
                })
            ]
        })
    ]
  });

  // --- G3022 Summary Table Rows (Pages 1-2) ---
  const g3022Rows: TableRow[] = [];
  if (g3022Data) {
      g3022Data.checklist.forEach((item, index) => {
        // G3022 summary table status mark: O / X / ―
        let statusMark = '―';
        if (item.status === ComplianceStatus.COMPLIANT) statusMark = 'O';
        else if (item.status === ComplianceStatus.NON_COMPLIANT || item.status === ComplianceStatus.CLARIFY) statusMark = 'X';
        else if (item.status === ComplianceStatus.NA) statusMark = '―';

        const isHeader = !item.id.includes('.'); 
        const isLast = index === g3022Data.checklist.length - 1;
        const bottomBorder = isLast ? BORDER_THICK : BORDER_THIN;
        
        if (isHeader) {
            g3022Rows.push(new TableRow({
                children: [
                    new TableCell({
                        columnSpan: 3,
                        width: { size: TABLE_WIDTH_DXA, type: WidthType.DXA },
                        borders: { top: BORDER_THIN, bottom: bottomBorder, left: BORDER_THICK, right: BORDER_THICK },
                        shading: { fill: SHADING_LIGHT_GRAY, type: ShadingType.CLEAR },
                        margins: CELL_MARGINS,
                        children: [stdP([txt(item.id + " " + item.name, { bold: true })])]
                    })
                ]
            }));
        } else {
            g3022Rows.push(new TableRow({
                children: [
                    new TableCell({
                        width: { size: 5046, type: WidthType.DXA }, // 50%
                        borders: { top: BORDER_THIN, bottom: bottomBorder, left: BORDER_THICK, right: BORDER_THIN },
                        margins: CELL_MARGINS,
                        children: [stdP(parseChecklistName(item.id + " " + item.name, 12, false))]
                    }),
                    new TableCell({
                        width: { size: 2523, type: WidthType.DXA }, // 25%
                        borders: { top: BORDER_THIN, bottom: bottomBorder, left: BORDER_THIN, right: BORDER_THIN },
                        margins: CELL_MARGINS,
                        children: [stdP([txt(statusMark)], "center")]
                    }),
                    new TableCell({
                        width: { size: 2524, type: WidthType.DXA }, // 25%
                        borders: { top: BORDER_THIN, bottom: bottomBorder, left: BORDER_THIN, right: BORDER_THICK },
                        margins: CELL_MARGINS,
                        children: [] 
                    })
                ]
            }));
        }
      });
  }
  
  // --- G3026 Checklist Table Rows (Pages 3-6) ---
  const g3026Rows: TableRow[] = [];
  const groups = [
    { idPrefix: '1', titleChars: ["1.", "組", "織", "邊", "界"] },
    { idPrefix: '2', titleChars: ["2.", "報", "告", "邊", "界"] },
    { idPrefix: '3', titleChars: ["3.", "量", "化", "方", "法"] },
    { idPrefix: '4', titleChars: ["4.", "基", "準", "年", "排", "放", "量"] },
    { idPrefix: '5', titleChars: ["5.", "數", "據", "品", "質", "管", "理"] },
  ];

  groups.forEach((group) => {
    const items = checklist.filter(item => item.id.startsWith(group.idPrefix));
    if (items.length === 0) return;

    items.forEach((item, index) => {
        const isFirst = index === 0;
        
        // G3026 status mapping: O / X / ―
        let statusMark = '―';
        if (item.status === ComplianceStatus.COMPLIANT) statusMark = 'O';
        else if (item.status === ComplianceStatus.NON_COMPLIANT) statusMark = 'X';
        else if (item.status === ComplianceStatus.NA) statusMark = '―';
        else if (item.status === ComplianceStatus.CLARIFY) statusMark = 'X'; // Treat CLARIFY as X if present

        const rowCells: TableCell[] = [];

        // 1. Group Title (Vertical Merged)
        if (isFirst) {
            rowCells.push(new TableCell({
                rowSpan: items.length,
                width: { size: 500, type: WidthType.DXA },
                shading: { fill: SHADING_LIGHT_GRAY, type: ShadingType.CLEAR },
                borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THICK, right: BORDER_THIN },
                verticalAlign: VerticalAlign.CENTER,
                children: group.titleChars.map(char => stdP([txt(char)], "center"))
            }));
        }

        // 2. Item Content (Font size 10)
        rowCells.push(new TableCell({
            width: { size: 3000, type: WidthType.DXA },
            borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN },
            margins: CELL_MARGINS,
            children: [stdP(parseChecklistName(item.id + " " + item.name, 10, false))]
        }));

        // 3. Doc Ref
        rowCells.push(new TableCell({
            width: { size: 2500, type: WidthType.DXA },
            borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN },
            margins: CELL_MARGINS,
            children: [stdP([txt(item.docRef || "")])]
        }));

        // 4. Field Obs
        rowCells.push(new TableCell({
            width: { size: 2800, type: WidthType.DXA },
            borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN },
            margins: CELL_MARGINS,
            children: [stdP([txt(item.fieldObs || "")])]
        }));

        // 5. Status
        rowCells.push(new TableCell({
            width: { size: 1200, type: WidthType.DXA },
            borders: { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THICK },
            margins: CELL_MARGINS,
            children: [stdP([txt(statusMark)], "center")]
        }));

        g3026Rows.push(new TableRow({ children: rowCells }));
    });
  });

  // --- Document ---
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: FONTS_OPTS, size: BASE_SIZE },
          paragraph: { spacing: STRICT_SPACING }
        }
      }
    },
    sections: [
      // === Section 1: G3022 Summary (Pages 1-2) ===
      {
        headers: { default: header },
        footers: { default: footer },
        properties: {
            page: { margin: { top: MARGIN_TOP, bottom: MARGIN_BOTTOM, left: MARGIN_LEFT, right: MARGIN_RIGHT, header: MARGIN_HEADER, footer: MARGIN_FOOTER } }
        },
        children: [
            // Basic Info & Doc Checklist Table
            new Table({
                width: { size: TABLE_WIDTH_DXA, type: WidthType.DXA },
                borders: { top: BORDER_THICK, bottom: BORDER_THICK, left: BORDER_THICK, right: BORDER_THICK, insideHorizontal: BORDER_NONE, insideVertical: BORDER_NONE },
                rows: [
                    // Row 1: Date (Size 14) with Raised Position
                    new TableRow({
                        height: { value: 600, rule: HeightRule.ATLEAST },
                        children: [
                            new TableCell({
                                margins: { top: 100, bottom: 50, left: 100, right: 100 },
                                children: [
                                    stdP([
                                        txt("查驗年度：中華民國 ", { size: 14, position: 28 }),
                                        txt(` ${basicInfo.year} `, { size: 14, underline: true, position: 28 }),
                                        txt(" 年    查驗日期：中華民國 ", { size: 14, position: 28 }),
                                        txt(` ${checkDate.y} `, { size: 14, underline: true, position: 28 }), txt(" 年 ", { size: 14, position: 28 }),
                                        txt(` ${checkDate.m} `, { size: 14, underline: true, position: 28 }), txt(" 月 ", { size: 14, position: 28 }),
                                        txt(` ${checkDate.d} `, { size: 14, underline: true, position: 28 }), txt(" 日", { size: 14, position: 28 })
                                    ])
                                ]
                            })
                        ]
                    }),
                    // Row 2: Disclaimer (Size 12)
                    new TableRow({
                        children: [
                            new TableCell({
                                margins: { top: 50, bottom: 50, left: 100, right: 100 },
                                children: [stdP([txt("本次查驗活動報告係依據下列標準與文件據以查核，並依廠商現況，以抽樣原則執行：", { size: 12 })])]
                            })
                        ]
                    }),
                    // Row 3: Criteria (Size 12)
                    new TableRow({
                        children: [
                            new TableCell({
                                margins: { top: 50, bottom: 50, left: 100, right: 100 },
                                children: [stdP([txt("查驗依據 : ISO 14064-1(2018 年版)/CNS 14064-1(2021 年版)", { size: 12 })])]
                            })
                        ]
                    }),
                    // Row 4: Checkbox 1 (Size 12)
                    new TableRow({
                        children: [
                            new TableCell({
                                margins: { top: 50, bottom: 50, left: 100, right: 100 },
                                children: [stdP([checkbox(!!basicInfo.reportInfo), txt(" 溫室氣體報告（編號／版次／發行日期）：", { size: 12 }), txt(basicInfo.reportInfo || "", { size: 12 })])]
                            })
                        ]
                    }),
                    // Row 5: Checkbox 2 (Size 12)
                    new TableRow({
                        children: [
                            new TableCell({
                                margins: { top: 50, bottom: 50, left: 100, right: 100 },
                                children: [stdP([checkbox(!!basicInfo.inventoryInfo), txt(" 盤查清冊（編號／版次／發行日期）: ", { size: 12 }), txt(basicInfo.inventoryInfo || "", { size: 12 })])]
                            })
                        ]
                    }),
                    // Row 6: Checkbox 3 (Size 12)
                    new TableRow({
                        children: [
                            new TableCell({
                                margins: { top: 50, bottom: 50, left: 100, right: 100 },
                                children: [stdP([checkbox(!!basicInfo.powerFactorInfo), txt(" 溫室氣體資訊管理程序（版次或公布日期）：", { size: 12 }), txt(basicInfo.powerFactorInfo || "", { size: 12 })])]
                            })
                        ]
                    }),
                    // Row 7: Checkbox 4 (Size 12)
                    new TableRow({
                        children: [
                            new TableCell({
                                margins: { top: 50, bottom: 100, left: 100, right: 100 },
                                children: [stdP([checkbox(!!basicInfo.otherInfo), txt(" 其他：", { size: 12 }), txt(basicInfo.otherInfo || "", { size: 12 })])]
                            })
                        ]
                    }),
                ]
            }),

            new Paragraph({ children: [txt("查驗項目一覽表", { size: 16, bold: true })], alignment: "center", spacing: { before: 240, after: 240 } }),

            new Table({
                width: { size: TABLE_WIDTH_DXA, type: WidthType.DXA },
                layout: TableLayoutType.FIXED,
                borders: { top: BORDER_THICK, bottom: BORDER_THICK, left: BORDER_THICK, right: BORDER_THICK, insideHorizontal: BORDER_THIN, insideVertical: BORDER_THIN },
                rows: [
                    // Header Row
                    new TableRow({
                        tableHeader: true,
                        children: [
                            new TableCell({
                                width: { size: 5046, type: WidthType.DXA },
                                shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR },
                                borders: { top: BORDER_THICK, left: BORDER_THICK, bottom: BORDER_THICK, right: BORDER_THIN },
                                verticalAlign: VerticalAlign.CENTER,
                                children: [stdP([txt("ISO 14064-1:2018 規定項目")], "center")]
                            }),
                            new TableCell({
                                width: { size: 2523, type: WidthType.DXA },
                                shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR },
                                borders: { top: BORDER_THICK, left: BORDER_THIN, bottom: BORDER_THICK, right: BORDER_THIN },
                                verticalAlign: VerticalAlign.CENTER,
                                children: [stdP([txt("符合項目以 O 註記\n不符合項目以 X 註記\n不適用以―註記", { size: 10 })], "center")]
                            }),
                            new TableCell({
                                width: { size: 2524, type: WidthType.DXA },
                                shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR },
                                borders: { top: BORDER_THICK, left: BORDER_THIN, bottom: BORDER_THICK, right: BORDER_THICK },
                                verticalAlign: VerticalAlign.CENTER,
                                children: [stdP([txt("備註")], "center")]
                            }),
                        ]
                    }),
                    ...g3022Rows
                ]
            })
        ]
      },

      // === Section 2: G3026 Checklist (Pages 3-6) ===
      {
        headers: { default: header },
        footers: { default: footer },
        properties: {
          type: SectionType.NEXT_PAGE,
          page: { margin: { top: MARGIN_TOP, bottom: MARGIN_BOTTOM, left: MARGIN_LEFT, right: MARGIN_RIGHT, header: MARGIN_HEADER, footer: MARGIN_FOOTER } }
        },
        children: [
            new Paragraph({ children: [txt("查驗檢核表", { size: 16 })], alignment: "center", spacing: { before: 240, after: 240 } }),

            new Table({
                width: { size: TABLE_WIDTH_DXA, type: WidthType.DXA },
                layout: TableLayoutType.FIXED,
                borders: { top: BORDER_THICK, bottom: BORDER_THICK, left: BORDER_THICK, right: BORDER_THICK },
                rows: [
                    // Header Row 1
                    new TableRow({
                        tableHeader: true,
                        children: [
                            new TableCell({
                                rowSpan: 2,
                                width: { size: 500, type: WidthType.DXA },
                                shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR },
                                borders: { top: BORDER_THICK, left: BORDER_THICK, bottom: BORDER_THIN, right: BORDER_THIN },
                                verticalAlign: VerticalAlign.CENTER,
                                children: [stdP([txt("查驗\n重點")], "center")]
                            }),
                            new TableCell({
                                width: { size: 3000, type: WidthType.DXA },
                                shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR },
                                borders: { top: BORDER_THICK, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN },
                                verticalAlign: VerticalAlign.CENTER,
                                children: [stdP([txt("查 證 內 容")], "center")]
                            }),
                            new TableCell({
                                columnSpan: 3,
                                width: { size: 6593, type: WidthType.DXA },
                                shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR },
                                borders: { top: BORDER_THICK, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THICK },
                                verticalAlign: VerticalAlign.CENTER,
                                children: [
                                    stdP([txt("查 驗 情 形")], "center"),
                                    stdP([txt("(符合項目以O註記，不符合項目以X註記，不適用以―註記)", { size: 10 })], "center"),
                                ]
                            }),
                        ]
                    }),
                    // Header Row 2
                    new TableRow({
                        tableHeader: true,
                        children: [
                            new TableCell({
                                width: { size: 3000, type: WidthType.DXA },
                                shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR },
                                borders: { top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THICK, right: BORDER_THIN },
                                verticalAlign: VerticalAlign.CENTER,
                                children: [stdP([txt("項目")], "center")]
                            }),
                            new TableCell({
                                width: { size: 2500, type: WidthType.DXA },
                                shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR },
                                borders: { top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THICK, right: BORDER_THIN },
                                verticalAlign: VerticalAlign.CENTER,
                                children: [stdP([txt("查驗文件")], "center")]
                            }),
                            new TableCell({
                                width: { size: 2800, type: WidthType.DXA },
                                shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR },
                                borders: { top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THICK, right: BORDER_THIN },
                                verticalAlign: VerticalAlign.CENTER,
                                children: [stdP([txt("現場觀察說明")], "center")]
                            }),
                            new TableCell({
                                width: { size: 1200, type: WidthType.DXA },
                                shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR },
                                borders: { top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THICK, right: BORDER_THICK },
                                verticalAlign: VerticalAlign.CENTER,
                                children: [stdP([txt("符合\n與否")], "center")]
                            }),
                        ]
                    }),
                    ...g3026Rows,
                    // Footer Row 1
                    new TableRow({
                        children: [
                            new TableCell({
                                columnSpan: 5,
                                shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR },
                                borders: { top: BORDER_THICK, left: BORDER_THICK, right: BORDER_THICK, bottom: BORDER_THIN },
                                children: [stdP([txt("其他觀察事項說明:")])]
                            })
                        ]
                    }),
                    // Footer Row 2
                    new TableRow({
                        height: { value: 1800, rule: HeightRule.EXACT },
                        children: [
                            new TableCell({
                                columnSpan: 5,
                                borders: { top: BORDER_THIN, left: BORDER_THICK, right: BORDER_THICK, bottom: BORDER_THIN },
                                children: [stdP([txt(otherObservation || "")])]
                            })
                        ]
                    }),
                    // Footer Row 3 - Signature (Height Increased to 1800 ~ 3.2cm)
                    new TableRow({
                        height: { value: 1800, rule: HeightRule.EXACT },
                        children: [
                            new TableCell({
                                columnSpan: 5,
                                borders: { top: BORDER_THIN, left: BORDER_THICK, right: BORDER_THICK, bottom: BORDER_THICK },
                                verticalAlign: VerticalAlign.CENTER,
                                children: [stdP([txt("主導查驗員簽名："), txt(` ${leadVerifierName || ""} `)])]
                            })
                        ]
                    })
                ]
            })
        ]
      },

      // === Section 3: Appendix 1 (Landscape) ===
      {
        headers: { default: header },
        footers: { default: footer },
        properties: {
          type: SectionType.NEXT_PAGE,
          page: { 
            size: { orientation: PageOrientation.LANDSCAPE },
            margin: { top: MARGIN_TOP, bottom: MARGIN_BOTTOM, left: MARGIN_LEFT, right: MARGIN_RIGHT, header: MARGIN_HEADER, footer: MARGIN_FOOTER } 
          }
        },
        children: [
            new Paragraph({ children: [txt("附件一 查驗取樣結果", { size: 16 })], alignment: "left", spacing: { before: 240, after: 240 } }),
            
            new Table({
                width: { size: LANDSCAPE_TABLE_WIDTH_DXA, type: WidthType.DXA },
                rows: [
                    new TableRow({
                        tableHeader: true,
                        children: [
                            new TableCell({ width: { size: 700, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, borders: {top: BORDER_THICK, left: BORDER_THICK, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt("NO.")], "center")] }),
                            new TableCell({ width: { size: 3000, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, borders: {top: BORDER_THICK, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt("區域/\n排放源")], "center")] }),
                            new TableCell({ width: { size: 2500, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, borders: {top: BORDER_THICK, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt("抽樣活動數據\n數值/排放量\n(含單位)")], "center")] }),
                            new TableCell({ width: { size: 3000, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, borders: {top: BORDER_THICK, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt("活動數據來源")], "center")] }),
                            new TableCell({ width: { size: 2500, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, borders: {top: BORDER_THICK, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt("活動數據類型\n抽樣比例\n(抽樣數/母數)")], "center")] }),
                            new TableCell({ width: { size: 1500, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, borders: {top: BORDER_THICK, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THICK}, children: [stdP([txt("排放源佔總\n排放量比")], "center")] }),
                            new TableCell({ width: { size: 1936, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, borders: {top: BORDER_THICK, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THICK}, children: [stdP([txt("備註")], "center")] }),
                        ]
                    }),
                    ...samplingResults.map((item, idx) => new TableRow({
                        children: [
                            new TableCell({ borders: {top: BORDER_THIN, left: BORDER_THICK, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt(`${idx + 1}`)], "center")] }),
                            new TableCell({ borders: {top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt(item.area)])] }),
                            new TableCell({ borders: {top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt(item.value)])] }),
                            new TableCell({ borders: {top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt(item.source)])] }),
                            new TableCell({ borders: {top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt(`${item.type}\n${item.ratio || ''}`)])] }), // Handle ratio if merged in logic or cell? Ah, we added a dedicated column for Ratio above
                             new TableCell({ borders: {top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt(item.ratio || "")])] }), // New Cell for Ratio
                            new TableCell({ borders: {top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THICK}, children: [stdP([txt(item.remarks)])] }),
                        ]
                    })),
                    // Closure line
                    new TableRow({ height: { value: 0, rule: HeightRule.AUTO }, children: [new TableCell({children:[], columnSpan:7, borders:{top:BORDER_THICK, left:BORDER_NONE, right:BORDER_NONE, bottom:BORDER_NONE}})]})
                ]
            })
        ]
      },

      // === Section 4: Appendix 2 (Landscape) ===
      {
        headers: { default: header },
        footers: { default: footer },
        properties: {
          type: SectionType.NEXT_PAGE,
          page: { 
            size: { orientation: PageOrientation.LANDSCAPE },
            margin: { top: MARGIN_TOP, bottom: MARGIN_BOTTOM, left: MARGIN_LEFT, right: MARGIN_RIGHT, header: MARGIN_HEADER, footer: MARGIN_FOOTER } 
          }
        },
        children: [
            new Paragraph({ children: [txt("附件二 排放係數確認", { size: 16 })], alignment: "left", spacing: { before: 240, after: 240 } }),
            
            new Table({
                width: { size: LANDSCAPE_TABLE_WIDTH_DXA, type: WidthType.DXA },
                rows: [
                    new TableRow({
                        tableHeader: true,
                        children: [
                            new TableCell({ width: { size: 700, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, borders: {top: BORDER_THICK, left: BORDER_THICK, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt("NO.")], "center")] }),
                            new TableCell({ width: { size: 3500, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, borders: {top: BORDER_THICK, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt("排放係數項目")], "center")] }),
                            new TableCell({ width: { size: 3500, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, borders: {top: BORDER_THICK, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt("排放係數來源")], "center")] }),
                            new TableCell({ width: { size: 5000, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, borders: {top: BORDER_THICK, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt("排放係數說明")], "center")] }),
                            new TableCell({ width: { size: 2436, type: WidthType.DXA }, shading: { fill: SHADING_GRAY, type: ShadingType.CLEAR }, borders: {top: BORDER_THICK, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THICK}, children: [stdP([txt("備註")], "center")] }),
                        ]
                    }),
                    ...emissionFactors.map((item, idx) => new TableRow({
                        children: [
                            new TableCell({ borders: {top: BORDER_THIN, left: BORDER_THICK, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt(`${idx + 1}`)], "center")] }),
                            new TableCell({ borders: {top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt(item.item)])] }),
                            new TableCell({ borders: {top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt(item.source)])] }),
                            new TableCell({ borders: {top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN}, children: [stdP([txt(item.description)])] }),
                            new TableCell({ borders: {top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THICK}, children: [stdP([txt(item.remarks)])] }),
                        ]
                    })),
                    // Closure line
                    new TableRow({ height: { value: 0, rule: HeightRule.AUTO }, children: [new TableCell({children:[], columnSpan:5, borders:{top:BORDER_THICK, left:BORDER_NONE, right:BORDER_NONE, bottom:BORDER_NONE}})]})
                ]
            })
        ]
      }
    ]
  });

  return await Packer.toBlob(doc);
};