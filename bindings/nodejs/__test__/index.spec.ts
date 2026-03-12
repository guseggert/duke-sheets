import { describe, it, expect } from "vitest";
import { Workbook, type CellValue } from "../index.js";
import * as path from "node:path";
import * as os from "node:os";
import * as fs from "node:fs";

// Workbook Tests

describe("Workbook", () => {
  describe("creation", () => {
    it("creates a new workbook with one sheet", () => {
      const wb = new Workbook();
      expect(wb.sheetCount).toBe(1);
    });

    it("has default sheet named Sheet1", () => {
      const wb = new Workbook();
      expect(wb.sheetNames).toEqual(["Sheet1"]);
    });
  });

  describe("sheet management", () => {
    it("adds a sheet", () => {
      const wb = new Workbook();
      const idx = wb.addSheet("NewSheet");

      expect(idx).toBe(1);
      expect(wb.sheetCount).toBe(2);
      expect(wb.sheetNames).toContain("NewSheet");
    });

    it("adds multiple sheets", () => {
      const wb = new Workbook();
      wb.addSheet("Sheet2");
      wb.addSheet("Sheet3");

      expect(wb.sheetCount).toBe(3);
      expect(wb.sheetNames).toEqual(["Sheet1", "Sheet2", "Sheet3"]);
    });

    it("removes a sheet", () => {
      const wb = new Workbook();
      wb.addSheet("ToRemove");
      expect(wb.sheetCount).toBe(2);

      wb.removeSheet(1);
      expect(wb.sheetCount).toBe(1);
      expect(wb.sheetNames).not.toContain("ToRemove");
    });

    it("gets sheet by index", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      expect(sheet.name).toBe("Sheet1");
    });

    it("gets sheet by name", () => {
      const wb = new Workbook();
      wb.addSheet("MySheet");

      const sheet = wb.getSheet("MySheet");
      expect(sheet.name).toBe("MySheet");
    });

    it("throws on invalid sheet index", () => {
      const wb = new Workbook();
      expect(() => wb.getSheet(999)).toThrow();
    });

    it("throws on invalid sheet name", () => {
      const wb = new Workbook();
      expect(() => wb.getSheet("NonExistent")).toThrow();
    });
  });

  describe("file operations", () => {
    it("saves and opens XLSX", () => {
      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-"));
      const filePath = path.join(tmpDir, "test.xlsx");

      try {
        const wb = new Workbook();
        const sheet = wb.getSheet(0);
        sheet.setCell("A1", 123);
        sheet.setCell("B1", "Hello");
        wb.save(filePath);

        expect(fs.existsSync(filePath)).toBe(true);
        expect(fs.statSync(filePath).size).toBeGreaterThan(0);

        const wb2 = Workbook.open(filePath);
        const sheet2 = wb2.getSheet(0);
        expect(sheet2.getCell("A1").asNumber()).toBe(123);
        expect(sheet2.getCell("B1").asText()).toBe("Hello");
      } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
      }
    });

    it("saves and opens CSV", () => {
      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-"));
      const filePath = path.join(tmpDir, "test.csv");

      try {
        const wb = new Workbook();
        const sheet = wb.getSheet(0);
        sheet.setCell("A1", 1);
        sheet.setCell("B1", 2);
        sheet.setCell("A2", 3);
        sheet.setCell("B2", 4);
        wb.save(filePath);

        expect(fs.existsSync(filePath)).toBe(true);

        const content = fs.readFileSync(filePath, "utf-8");
        expect(content).toContain("1");
        expect(content).toContain("2");
      } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
      }
    });

    it("loads from XLSX bytes", () => {
      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-"));
      const filePath = path.join(tmpDir, "test.xlsx");

      try {
        // Create a file first
        const wb = new Workbook();
        const sheet = wb.getSheet(0);
        sheet.setCell("A1", 42);
        wb.save(filePath);

        // Load from bytes
        const bytes = fs.readFileSync(filePath);
        const wb2 = Workbook.fromBytes(bytes);
        const sheet2 = wb2.getSheet(0);
        expect(sheet2.getCell("A1").asNumber()).toBe(42);
      } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
      }
    });

    it("loads from CSV string", () => {
      const wb = Workbook.fromCsvString("a,b,c\n1,2,3");
      const sheet = wb.getSheet(0);
      // CSV values are typically parsed as strings or numbers
      expect(wb.sheetCount).toBe(1);
    });

    it("saves to CSV string", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      sheet.setCell("A1", 1);
      sheet.setCell("B1", 2);
      sheet.setCell("A2", 3);
      sheet.setCell("B2", 4);

      const csv = wb.saveCsvString();
      expect(csv).toContain("1");
      expect(csv).toContain("2");
    });
  });

  describe("named ranges", () => {
    it("defines and gets a named range constant", () => {
      const wb = new Workbook();
      wb.defineName("TaxRate", "0.05");

      const result = wb.getNamedRange("TaxRate");
      expect(result).toBe("0.05");
    });

    it("defines a named range with cell reference", () => {
      const wb = new Workbook();
      wb.defineName("Price", "Sheet1!$A$1");

      const result = wb.getNamedRange("Price");
      expect(result).toContain("A");
      expect(result).toContain("1");
    });

    it("returns null for undefined name", () => {
      const wb = new Workbook();
      const result = wb.getNamedRange("NotDefined");
      expect(result).toBeNull();
    });
  });
});

// Worksheet Tests

describe("Worksheet", () => {
  describe("cell values", () => {
    it("sets and gets a number", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);

      sheet.setCell("A1", 42);
      const value = sheet.getCell("A1");

      expect(value.isNumber).toBe(true);
      expect(value.asNumber()).toBe(42);
    });

    it("sets and gets text", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);

      sheet.setCell("A1", "Hello");
      const value = sheet.getCell("A1");

      expect(value.isText).toBe(true);
      expect(value.asText()).toBe("Hello");
    });

    it("sets and gets a boolean", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);

      sheet.setCell("A1", true);
      const value = sheet.getCell("A1");

      expect(value.isBoolean).toBe(true);
      expect(value.asBoolean()).toBe(true);
    });

    it("clears cell with null", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);

      sheet.setCell("A1", 42);
      sheet.setCell("A1", null);
      const value = sheet.getCell("A1");

      expect(value.isEmpty).toBe(true);
    });

    it("returns empty for unset cells", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);

      const value = sheet.getCell("Z99");
      expect(value.isEmpty).toBe(true);
    });
  });

  describe("used range", () => {
    it("returns null for empty worksheet", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      expect(sheet.usedRange).toBeNull();
    });

    it("returns range with data", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);

      sheet.setCell("B2", 1);
      sheet.setCell("D4", 2);

      const range = sheet.usedRange;
      expect(range).not.toBeNull();
      expect(range!.minRow).toBeDefined();
      expect(range!.maxRow).toBeDefined();
      expect(range!.minCol).toBeDefined();
      expect(range!.maxCol).toBeDefined();
    });
  });

  describe("row/column dimensions", () => {
    it("sets row height", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);

      sheet.setRowHeight(0, 30.0);
      expect(sheet.getRowHeight(0)).toBe(30.0);
    });

    it("sets column width", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);

      sheet.setColumnWidth(0, 15.0);
      expect(sheet.getColumnWidth(0)).toBe(15.0);
    });
  });

  describe("merge cells", () => {
    it("merges cells", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);

      sheet.setCell("A1", "Merged");
      sheet.mergeCells("A1:C3");
      // No error means success
    });

    it("unmerges cells", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);

      sheet.mergeCells("A1:C3");
      const result = sheet.unmergeCells("A1:C3");
      expect(result).toBe(true);
    });
  });
});

// Formula Tests

describe("Formulas", () => {
  it("sets a formula", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setFormula("A1", "=1+1");

    expect(sheet.getFormulaAt(0, 0)).toBe("=1+1");
  });

  it("calculates cell references", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setCell("A1", 10);
    sheet.setCell("A2", 20);
    sheet.setFormula("A3", "=A1+A2");

    wb.calculate();

    const value = sheet.getCalculatedValue("A3");
    expect(value.asNumber()).toBe(30);
  });

  it("calculates SUM", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setCell("A1", 1);
    sheet.setCell("A2", 2);
    sheet.setCell("A3", 3);
    sheet.setFormula("A4", "=SUM(A1:A3)");

    wb.calculate();

    const value = sheet.getCalculatedValue("A4");
    expect(value.asNumber()).toBe(6);
  });

  it("calculates nested formulas", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setCell("A1", 5);
    sheet.setFormula("A2", "=A1*2"); // 10
    sheet.setFormula("A3", "=A2+A1"); // 15

    wb.calculate();

    expect(sheet.getCalculatedValue("A2").asNumber()).toBe(10);
    expect(sheet.getCalculatedValue("A3").asNumber()).toBe(15);
  });
});

// Calculation Tests

describe("Calculation", () => {
  it("returns stats", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setFormula("A1", "=1+1");
    sheet.setFormula("A2", "=2+2");

    const stats = wb.calculate();

    expect(stats.formulaCount).toBe(2);
    expect(stats.cellsCalculated).toBeGreaterThanOrEqual(2);
    expect(stats.errors).toBe(0);
  });

  it("calculates with options object", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setFormula("A1", "=1+1");

    const stats = wb.calculateWithOptions({ iterative: false, maxIterations: 100, maxChange: 0.001 });
    expect(stats.formulaCount).toBe(1);
    expect(sheet.getCalculatedValue("A1").asNumber()).toBe(2);
  });

  it("calculates with empty options (all defaults)", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setFormula("A1", "=1+1");

    const stats = wb.calculateWithOptions({});
    expect(stats.formulaCount).toBe(1);
    expect(sheet.getCalculatedValue("A1").asNumber()).toBe(2);
  });

  it("calculates with forceFullCalculation option", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setFormula("A1", "=1+1");

    const stats = wb.calculateWithOptions({ forceFullCalculation: true });
    expect(stats.formulaCount).toBe(1);
    expect(sheet.getCalculatedValue("A1").asNumber()).toBe(2);
  });

  it("calculates with calculateVolatile option", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setFormula("A1", "=1+1");

    const stats = wb.calculateWithOptions({ calculateVolatile: false });
    expect(stats.formulaCount).toBe(1);
  });

  it("calculates specific sheets only", () => {
    const wb = new Workbook();
    const sheet0 = wb.getSheet(0);
    sheet0.setFormula("A1", "=1+1");
    wb.addSheet("Sheet2");
    const sheet1 = wb.getSheet(1);
    sheet1.setFormula("A1", "=2+2");

    const stats = wb.calculateWithOptions({ sheets: [0] });
    expect(stats.cellsCalculated).toBeGreaterThanOrEqual(1);
    expect(sheet0.getCalculatedValue("A1").asNumber()).toBe(2);
  });

  it("calculates with maxThreads option", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setFormula("A1", "=1+1");

    const stats = wb.calculateWithOptions({ maxThreads: 1 });
    expect(stats.formulaCount).toBe(1);
    expect(sheet.getCalculatedValue("A1").asNumber()).toBe(2);
  });

  it("worksheet formulaCount getter", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    expect(sheet.formulaCount).toBe(0);
    sheet.setFormula("A1", "=1+1");
    sheet.setFormula("B1", "=2+2");
    sheet.setCell("C1", 42);
    expect(sheet.formulaCount).toBe(2);
  });
});

// CellValue Tests

describe("CellValue", () => {
  it("toJs returns number", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setCell("A1", 42.5);
    const value = sheet.getCell("A1");

    expect(value.toJs()).toBe(42.5);
  });

  it("toJs returns string", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setCell("A1", "Hello");
    const value = sheet.getCell("A1");

    expect(value.toJs()).toBe("Hello");
  });

  it("toJs returns boolean", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setCell("A1", true);
    const value = sheet.getCell("A1");

    expect(value.toJs()).toBe(true);
  });

  it("toJs returns null for empty", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    const value = sheet.getCell("Z99");
    expect(value.toJs()).toBeNull();
  });

  it("toString gives string representation", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setCell("A1", 42);
    const value = sheet.getCell("A1");

    expect(value.toString()).toBe("42");
  });
});

// CSV Roundtrip Tests

describe("CSV", () => {
  it("roundtrips through CSV string", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setCell("A1", 1);
    sheet.setCell("B1", 2);
    sheet.setCell("A2", 3);
    sheet.setCell("B2", 4);

    const csv = wb.saveCsvString();
    expect(csv).toContain("1");
    expect(csv).toContain("2");

    const wb2 = Workbook.fromCsvString(csv);
    expect(wb2.sheetCount).toBe(1);
  });
});
