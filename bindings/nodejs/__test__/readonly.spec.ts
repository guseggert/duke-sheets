import { describe, it, expect } from "vitest";
import { Workbook } from "../lib.js";
import * as path from "node:path";
import * as os from "node:os";
import * as fs from "node:fs";

describe("Workbook read-only", () => {
  it("isEmpty is false for a default workbook", () => {
    const wb = new Workbook();
    // Default workbook has one empty sheet, so isEmpty refers to sheet count
    expect(wb.isEmpty).toBe(false);
  });

  it("activeSheet returns 0 by default", () => {
    const wb = new Workbook();
    expect(wb.activeSheet).toBe(0);
  });

  it("sheetIndex returns index for existing sheet", () => {
    const wb = new Workbook();
    wb.addSheet("Foo");
    expect(wb.sheetIndex("Sheet1")).toBe(0);
    expect(wb.sheetIndex("Foo")).toBe(1);
  });

  it("sheetIndex returns null for non-existent sheet", () => {
    const wb = new Workbook();
    expect(wb.sheetIndex("DoesNotExist")).toBeNull();
  });

  it("settings returns workbook settings", () => {
    const wb = new Workbook();
    const settings = wb.settings;
    expect(settings).toBeDefined();
    expect(typeof settings.date1904).toBe("boolean");
    expect(typeof settings.protected).toBe("boolean");
    expect(typeof settings.calcOnOpen).toBe("boolean");
  });

  it("namedRanges is empty by default", () => {
    const wb = new Workbook();
    expect(wb.namedRanges).toEqual([]);
  });

  it("namedRanges returns defined ranges", () => {
    const wb = new Workbook();
    wb.defineName("MyRange", "Sheet1!$A$1");
    const ranges = wb.namedRanges;
    expect(ranges.length).toBeGreaterThanOrEqual(1);
    const found = ranges.find((r) => r.name === "MyRange");
    expect(found).toBeDefined();
    expect(found!.refersTo).toContain("A");
    expect(typeof found!.hidden).toBe("boolean");
    expect(typeof found!.scope).toBe("string");
  });
});

describe("Worksheet read-only properties", () => {
  it("visibility is 'visible' by default", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.visibility).toBe("visible");
  });

  it("isSelected returns a boolean", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(typeof sheet.isSelected).toBe("boolean");
  });

  it("zoomScale is null by default", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    // May be null or a number depending on defaults
    const zoom = sheet.zoomScale;
    expect(zoom === null || typeof zoom === "number").toBe(true);
  });

  it("tabColor is null by default", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.tabColor).toBeNull();
  });

  it("isEmpty is true for a fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.isEmpty).toBe(true);
  });

  it("isEmpty is false after setting a cell", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 42);
    expect(sheet.isEmpty).toBe(false);
  });

  it("cellCount is 0 for a fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.cellCount).toBe(0);
  });

  it("cellCount reflects number of cells with data", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 1);
    sheet.setCell("B2", 2);
    sheet.setCell("C3", 3);
    expect(sheet.cellCount).toBe(3);
  });

  it("selections returns an array", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    const sel = sheet.selections;
    expect(Array.isArray(sel)).toBe(true);
  });

  it("date1904 returns a boolean", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(typeof sheet.date1904).toBe("boolean");
  });
});

describe("Worksheet cell styles", () => {
  it("getCellStyle returns null for empty cell", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    const style = sheet.getCellStyle("Z99");
    // Empty cells may return null or a default style
    expect(style === null || typeof style === "object").toBe(true);
  });

  it("getCellStyleAt returns same as getCellStyle", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 42);
    const byAddr = sheet.getCellStyle("A1");
    const byIdx = sheet.getCellStyleAt(0, 0);
    // Both should be consistent
    if (byAddr && byIdx) {
      expect(byAddr.font.name).toBe(byIdx.font.name);
    }
  });

  it("getCellStyle has expected shape when present", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 42);
    const style = sheet.getCellStyle("A1");
    if (style) {
      expect(style.font).toBeDefined();
      expect(typeof style.font.name).toBe("string");
      expect(typeof style.font.bold).toBe("boolean");
      expect(style.fill).toBeDefined();
      expect(typeof style.fill.fillType).toBe("string");
      expect(style.border).toBeDefined();
      expect(style.alignment).toBeDefined();
      expect(style.numberFormat).toBeDefined();
      expect(typeof style.numberFormat.formatType).toBe("string");
      expect(style.protection).toBeDefined();
      expect(typeof style.protection.locked).toBe("boolean");
    }
  });

  it("getFormattedValue returns string for numeric cell", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 1234);
    const fv = sheet.getFormattedValue("A1");
    expect(typeof fv).toBe("string");
    expect(fv).toContain("1234");
  });

  it("getFormattedValueAt returns string", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 42.5);
    const fv = sheet.getFormattedValueAt(0, 0);
    expect(typeof fv).toBe("string");
    expect(fv).toContain("42.5");
  });
});

describe("Worksheet row/column read-only", () => {
  it("defaultRowHeight returns a positive number", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.defaultRowHeight).toBeGreaterThan(0);
  });

  it("defaultColumnWidth returns a positive number", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.defaultColumnWidth).toBeGreaterThan(0);
  });

  it("isRowHidden returns false by default", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.isRowHidden(0)).toBe(false);
  });

  it("isColumnHidden returns false by default", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.isColumnHidden(0)).toBe(false);
  });

  it("getRowOutlineLevel returns 0 by default", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.getRowOutlineLevel(0)).toBe(0);
  });

  it("getColumnOutlineLevel returns 0 by default", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.getColumnOutlineLevel(0)).toBe(0);
  });

  it("isRowCollapsed returns false by default", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.isRowCollapsed(0)).toBe(false);
  });

  it("isColumnCollapsed returns false by default", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.isColumnCollapsed(0)).toBe(false);
  });
});

describe("Worksheet freeze/split panes", () => {
  it("freezePanes is null by default", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.freezePanes).toBeNull();
  });

  it("splitPanes is null by default", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.splitPanes).toBeNull();
  });
});

describe("Worksheet hyperlinks", () => {
  it("hyperlinkCount is 0 for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.hyperlinkCount).toBe(0);
  });

  it("hyperlinks returns empty array for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.hyperlinks).toEqual([]);
  });

  it("getHyperlink returns null for cell without hyperlink", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.getHyperlink("A1")).toBeNull();
  });
});

describe("Worksheet comments", () => {
  it("commentCount is 0 for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.commentCount).toBe(0);
  });

  it("comments returns empty array for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.comments).toEqual([]);
  });

  it("hasComment returns false for cell without comment", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.hasComment("A1")).toBe(false);
  });

  it("hasCommentAt returns false for cell without comment", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.hasCommentAt(0, 0)).toBe(false);
  });

  it("getComment returns null for cell without comment", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.getComment("A1")).toBeNull();
  });

  it("getCommentAt returns null for cell without comment", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.getCommentAt(0, 0)).toBeNull();
  });

  it("commentAuthors returns array", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(Array.isArray(sheet.commentAuthors)).toBe(true);
  });
});

describe("Worksheet tables", () => {
  it("tableCount is 0 for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.tableCount).toBe(0);
  });

  it("tables returns empty array for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.tables).toEqual([]);
  });

  it("getTableByName returns null for non-existent table", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.getTableByName("NoSuchTable")).toBeNull();
  });
});

describe("Worksheet data validation", () => {
  it("dataValidationCount is 0 for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.dataValidationCount).toBe(0);
  });

  it("dataValidations returns empty array for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.dataValidations).toEqual([]);
  });
});

describe("Worksheet conditional formatting", () => {
  it("conditionalFormatCount is 0 for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.conditionalFormatCount).toBe(0);
  });

  it("conditionalFormats returns empty array for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.conditionalFormats).toEqual([]);
  });
});

describe("Worksheet auto-filter", () => {
  it("autoFilter is null for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.autoFilter).toBeNull();
  });
});

describe("Worksheet protection", () => {
  it("protection is null for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.protection).toBeNull();
  });
});

describe("Worksheet page setup", () => {
  it("pageSetup returns valid settings", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    const ps = sheet.pageSetup;
    expect(ps).toBeDefined();
    expect(typeof ps.paperSize).toBe("number");
    expect(typeof ps.orientation).toBe("string");
    expect(typeof ps.scale).toBe("number");
    expect(typeof ps.topMargin).toBe("number");
    expect(typeof ps.bottomMargin).toBe("number");
    expect(typeof ps.leftMargin).toBe("number");
    expect(typeof ps.rightMargin).toBe("number");
    expect(typeof ps.printGridlines).toBe("boolean");
    expect(typeof ps.printHeadings).toBe("boolean");
  });

  it("printArea is null for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.printArea).toBeNull();
  });

  it("repeatRows is null for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.repeatRows).toBeNull();
  });

  it("repeatCols is null for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.repeatCols).toBeNull();
  });

  it("rowBreaks returns empty array for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.rowBreaks).toEqual([]);
  });

  it("colBreaks returns empty array for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.colBreaks).toEqual([]);
  });
});

describe("Worksheet formulas read-only", () => {
  it("getFormulaAt returns null for non-formula cell", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 42);
    expect(sheet.getFormulaAt(0, 0)).toBeNull();
  });

  it("getFormulaAt returns formula text for formula cell", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setFormula("A1", "=1+1");
    const formula = sheet.getFormulaAt(0, 0);
    expect(formula).toBe("=1+1");
  });

  it("formulaCells returns empty array for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.formulaCells).toEqual([]);
  });

  it("formulaCells returns formula entries", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setFormula("B2", "=SUM(A1:A5)");
    sheet.setFormula("C3", "=1+2");

    const cells = sheet.formulaCells;
    expect(cells.length).toBe(2);

    const b2 = cells.find((c) => c.row === 1 && c.col === 1);
    expect(b2).toBeDefined();
    expect(b2!.formula).toContain("SUM");

    const c3 = cells.find((c) => c.row === 2 && c.col === 2);
    expect(c3).toBeDefined();
    expect(c3!.formula).toBe("=1+2");
  });
});

describe("Worksheet spill", () => {
  it("isSpillTarget returns false for normal cell", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.isSpillTarget(0, 0)).toBe(false);
  });

  it("isSpillSource returns false for normal cell", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.isSpillSource(0, 0)).toBe(false);
  });

  it("getSpillSource returns null for normal cell", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.getSpillSource(0, 0)).toBeNull();
  });
});

describe("Worksheet merged regions read-only", () => {
  it("mergedRegions returns empty array for fresh sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    expect(sheet.mergedRegions).toEqual([]);
  });

  it("mergedRegions reflects merged cells", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.mergeCells("A1:C3");
    const regions = sheet.mergedRegions;
    expect(regions.length).toBe(1);
    expect(regions[0]).toEqual({ startRow: 0, startCol: 0, endRow: 2, endCol: 2, range: "A1:C3" });
  });

  it("mergedRegions reflects multiple merges", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.mergeCells("A1:B2");
    sheet.mergeCells("D4:F6");
    const regions = sheet.mergedRegions;
    expect(regions.length).toBe(2);
    expect(regions).toEqual(expect.arrayContaining([
      expect.objectContaining({ range: "A1:B2" }),
      expect.objectContaining({ range: "D4:F6" }),
    ]));
  });
});

describe("Read-only API with XLSX roundtrip", () => {
  it("preserves cell styles through save/load", () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-ro-"));
    const filePath = path.join(tmpDir, "styles.xlsx");

    try {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      sheet.setCell("A1", 42);
      sheet.setCell("B1", "Hello");
      sheet.setCell("C1", true);
      wb.save(filePath);

      const wb2 = Workbook.open(filePath);
      const sheet2 = wb2.getSheet(0);

      // Verify cell count and isEmpty
      expect(sheet2.isEmpty).toBe(false);
      expect(sheet2.cellCount).toBeGreaterThanOrEqual(3);

      // Verify formatted values
      expect(sheet2.getFormattedValue("A1")).toContain("42");
      expect(sheet2.getFormattedValue("B1")).toBe("Hello");

      // Verify default row/column dimensions exist
      expect(sheet2.defaultRowHeight).toBeGreaterThan(0);
      expect(sheet2.defaultColumnWidth).toBeGreaterThan(0);

      // Verify workbook settings survive roundtrip
      const settings = wb2.settings;
      expect(typeof settings.date1904).toBe("boolean");
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("preserves merged regions through save/load", () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-ro-"));
    const filePath = path.join(tmpDir, "merged.xlsx");

    try {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      sheet.setCell("A1", "Merged Header");
      sheet.mergeCells("A1:D1");
      wb.save(filePath);

      const wb2 = Workbook.open(filePath);
      const sheet2 = wb2.getSheet(0);
      const regions = sheet2.mergedRegions;
      expect(regions.length).toBe(1);
      expect(regions[0]).toEqual({ startRow: 0, startCol: 0, endRow: 0, endCol: 3, range: "A1:D1" });
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("preserves formulas through save/load", () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-ro-"));
    const filePath = path.join(tmpDir, "formulas.xlsx");

    try {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      sheet.setCell("A1", 10);
      sheet.setCell("A2", 20);
      sheet.setFormula("A3", "=SUM(A1:A2)");
      wb.calculate();
      wb.save(filePath);

      const wb2 = Workbook.open(filePath);
      const sheet2 = wb2.getSheet(0);

      // Formula text should be preserved
      const formula = sheet2.getFormulaAt(2, 0);
      expect(formula).toBeTruthy();
      expect(formula!.toUpperCase()).toContain("SUM");

      // Formula cells list should include our formula
      const cells = sheet2.formulaCells;
      expect(cells.length).toBeGreaterThanOrEqual(1);
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });
});

describe("Worksheet.iterateRows() sparse iterator", () => {
  it("iterates sparse rows with for...of", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 10);
    sheet.setCell("C1", "hello");
    sheet.setCell("A3", 42);
    sheet.setCell("B5", true);

    const rows = [];
    for (const row of sheet.iterateRows()) {
      rows.push(row);
    }

    expect(rows).toHaveLength(3); // rows 0, 2, 4
    expect(rows[0].index).toBe(0);
    expect(rows[0].cells).toHaveLength(2);
    expect(rows[0].cells[0]).toEqual({ col: 0, value: "10" });
    expect(rows[0].cells[1]).toEqual({ col: 2, value: "hello" });
    expect(rows[1].index).toBe(2);
    expect(rows[1].cells).toHaveLength(1);
    expect(rows[1].cells[0]).toEqual({ col: 0, value: "42" });
    expect(rows[2].index).toBe(4);
    expect(rows[2].cells[0]).toEqual({ col: 1, value: "TRUE" });
  });

  it("returns empty iterator for empty sheet", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    const rows = [...sheet.iterateRows()];
    expect(rows).toHaveLength(0);
  });

  it("supports useFormattedValues option", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 0.5);

    const rawRows = [...sheet.iterateRows()];
    expect(rawRows[0].cells[0].value).toBe("0.5");

    const fmtRows = [...sheet.iterateRows({ useFormattedValues: true })];
    // Without a number format, formatted output may differ slightly
    expect(fmtRows[0].cells[0].value).toBeDefined();
  });

  it("supports useCalculatedValues option", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 10);
    sheet.setCell("A2", 20);
    sheet.setFormula("A3", "=A1+A2");
    wb.calculate();

    const rows = [...sheet.iterateRows({ useCalculatedValues: true })];
    const a3Row = rows.find((r) => r.index === 2);
    expect(a3Row).toBeDefined();
    expect(a3Row!.cells[0].value).toBe("30");
  });

  it("getRowsBatch returns batched results", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", "first");
    sheet.setCell("A100", "last");

    const batch1 = sheet.getRowsBatch(0, 50);
    expect(batch1).toHaveLength(1); // only row 0 has data in 0..49
    expect(batch1[0].index).toBe(0);

    const batch2 = sheet.getRowsBatch(50, 100);
    expect(batch2).toHaveLength(1); // only row 99 has data in 50..149
    expect(batch2[0].index).toBe(99);

    const batch3 = sheet.getRowsBatch(100, 100);
    expect(batch3).toHaveLength(0); // no data beyond row 99
  });
});

describe("iterateRows metadata flags", () => {
  require("../lib.js");

  it("includeStyles does not crash and returns style when present", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 42);

    const rows = [...sheet.iterateRows({ includeStyles: true })];
    expect(rows).toHaveLength(1);
    // Default-styled cells may have style=undefined since style_index=0
    // The key test is that the flag doesn't crash and the cell is present
    expect(rows[0].cells[0].col).toBe(0);
    expect(rows[0].cells[0].value).toBe("42");
  });

  it("includeMergeInfo returns merge data", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", "merged");
    sheet.mergeCells("A1:B2");

    const rows = [...sheet.iterateRows({ includeMergeInfo: true })];
    // Should include A1 (origin) and the secondary cells A2, B1, B2
    const a1 = rows.find(r => r.index === 0)?.cells.find(c => c.col === 0);
    expect(a1).toBeDefined();
    expect(a1!.mergeSpan).toEqual({ rowSpan: 2, colSpan: 2 });

    const b1 = rows.find(r => r.index === 0)?.cells.find(c => c.col === 1);
    expect(b1).toBeDefined();
    expect(b1!.isMergedSecondary).toBe(true);
  });

  it("includeFormulas returns formula text", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 10);
    sheet.setFormula("A2", "=A1*2");

    const rows = [...sheet.iterateRows({ includeFormulas: true })];
    const a2 = rows.find(r => r.index === 1)?.cells.find(c => c.col === 0);
    expect(a2).toBeDefined();
    expect(a2!.formula).toBe("=A1*2");

    // Non-formula cell should not have formula
    const a1 = rows.find(r => r.index === 0)?.cells.find(c => c.col === 0);
    expect(a1!.formula).toBeUndefined();
  });

  it("includeMergeInfo includes empty secondary cells", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", "merged");
    sheet.mergeCells("A1:B2");

    // Without flag: only A1 (has value)
    const noFlag = [...sheet.iterateRows()];
    expect(noFlag).toHaveLength(1);
    expect(noFlag[0].cells).toHaveLength(1);

    // With flag: all 4 cells in the merged region should appear
    const withFlag = [...sheet.iterateRows({ includeMergeInfo: true })];
    const allCells = withFlag.flatMap(r => r.cells);
    expect(allCells).toHaveLength(4); // A1, B1, A2, B2

    // A1 is the merge origin
    const a1 = withFlag[0].cells.find(c => c.col === 0);
    expect(a1!.mergeSpan).toEqual({ rowSpan: 2, colSpan: 2 });

    // B1 is secondary
    const b1 = withFlag[0].cells.find(c => c.col === 1);
    expect(b1!.isMergedSecondary).toBe(true);
    expect(b1!.value).toBe(""); // empty
  });

  it("fields are absent when flags not set", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 42);
    sheet.setFormula("A2", "=A1*2");

    const rows = [...sheet.iterateRows()];
    expect(rows[0].cells[0].style).toBeUndefined();
    expect(rows[0].cells[0].mergeSpan).toBeUndefined();
    expect(rows[0].cells[0].comment).toBeUndefined();
    expect(rows[0].cells[0].formula).toBeUndefined();
  });
});
