import { describe, it, expect } from "vitest";
import { Workbook } from "../index.js";
import * as path from "node:path";
import * as os from "node:os";
import * as fs from "node:fs";

// Plain (unencrypted) XLS and XLSB save → open round-trips. The binding
// dispatches the writer on the file extension; these formats carry
// binary formula token streams (BIFF8 / BIFF12), so formula text and
// named-range survival exercises the compilers end to end.

function tempPath(ext: string): string {
  return path.join(
    os.tmpdir(),
    `duke-fmt-${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2)}${ext}`,
  );
}

function buildSample(): Workbook {
  const wb = new Workbook();
  wb.defineName("MyRange", "Sheet1!$A$1:$A$3");
  const sheet = wb.getSheet(0);
  if (!sheet) throw new Error("expected default sheet");
  sheet.setCell("A1", 1.0);
  sheet.setCell("A2", 2.0);
  sheet.setCell("A3", 3.0);
  sheet.setCell("B1", "label");
  sheet.setCell("B2", true);
  sheet.setFormula("C1", "=SUM(A1:A3)");
  sheet.setFormula("C2", "=IF(A1>0,A2,A3)");
  sheet.setFormula("C3", "=SUM(MyRange)");
  wb.calculate();
  return wb;
}

function assertSample(opened: Workbook) {
  const sheet = opened.getSheet(0);
  if (!sheet) throw new Error("opened workbook has no sheet 0");
  expect(sheet.getCell("A1").asNumber()).toBe(1);
  expect(sheet.getCell("B1").asText()).toBe("label");
  expect(sheet.getCell("B2").asBoolean()).toBe(true);
  expect(sheet.getFormulaAt(0, 2)).toBe("=SUM(A1:A3)");
  expect(sheet.getFormulaAt(1, 2)).toBe("=IF(A1>0,A2,A3)");
  expect(sheet.getFormulaAt(2, 2)).toBe("=SUM(MyRange)");
  expect(sheet.getCell("C1").asNumber()).toBe(6);
}

function roundTrip(ext: string) {
  const file = tempPath(ext);
  try {
    buildSample().save(file);
    const opened = Workbook.open(file);
    assertSample(opened);
  } finally {
    try {
      fs.unlinkSync(file);
    } catch {}
  }
}

function openWithMismatchedExtension(savedExt: string, openedExt: string) {
  const source = tempPath(savedExt);
  const mismatched = tempPath(openedExt);
  try {
    buildSample().save(source);
    fs.copyFileSync(source, mismatched);
    assertSample(Workbook.open(mismatched));
  } finally {
    for (const file of [source, mismatched]) {
      try {
        fs.unlinkSync(file);
      } catch {}
    }
  }
}

describe("plain format save/open round-trips", () => {
  it("xls round-trips values, formulas, and named ranges", () => {
    roundTrip(".xls");
  });

  it("xlsb round-trips values, formulas, and named ranges", () => {
    roundTrip(".xlsb");
  });

  it("xlsx round-trips values, formulas, and named ranges", () => {
    roundTrip(".xlsx");
  });

  it("opens xlsb content even with an xlsx extension", () => {
    openWithMismatchedExtension(".xlsb", ".xlsx");
  });

  it("opens xlsx content even with an xlsb extension", () => {
    openWithMismatchedExtension(".xlsx", ".xlsb");
  });

  it("opens xls content even with an xlsx extension", () => {
    openWithMismatchedExtension(".xls", ".xlsx");
  });
});
