import { describe, it, expect } from "vitest";
import { Workbook } from "../index.js";

describe("error handling", () => {
  describe("invalid inputs produce errors, not crashes", () => {
    it("rejects out-of-range sheet index", () => {
      const wb = new Workbook();
      expect(() => wb.getSheet(999)).toThrow();
    });

    it("rejects nonexistent sheet name", () => {
      const wb = new Workbook();
      expect(() => wb.getSheet("NoSuchSheet")).toThrow();
    });

    it("rejects invalid cell address", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      expect(() => sheet.getCell("!!!")).toThrow();
      expect(() => sheet.setCell("!!!", 42)).toThrow();
      expect(() => sheet.getCalculatedValue("!!!")).toThrow();
    });

    it("rejects invalid range string for merge", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      expect(() => sheet.mergeCells("not-a-range")).toThrow();
      expect(() => sheet.unmergeCells("not-a-range")).toThrow();
    });

    it("rejects removing sheet at invalid index", () => {
      const wb = new Workbook();
      expect(() => wb.removeSheet(999)).toThrow();
    });

    it("rejects opening nonexistent file", () => {
      expect(() => Workbook.open("/nonexistent/path/file.xlsx")).toThrow();
    });

    it("rejects invalid bytes", () => {
      const garbage = Buffer.from("this is not an xlsx file");
      expect(() => Workbook.fromBytes(garbage)).toThrow();
    });

    it("rejects saving to invalid path", () => {
      const wb = new Workbook();
      expect(() => wb.save("/nonexistent/dir/file.xlsx")).toThrow();
    });
  });

  describe("process survives all error cases", () => {
    it("can continue using workbook after errors", () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);

      // Trigger several errors
      try { sheet.getCell("!!!"); } catch {}
      try { wb.getSheet(999); } catch {}
      try { sheet.mergeCells("bad"); } catch {}

      // Workbook should still be fully functional
      sheet.setCell("A1", 42);
      expect(sheet.getCell("A1").asNumber()).toBe(42);

      sheet.setFormula("A2", "=A1*2");
      wb.calculate();
      expect(sheet.getCalculatedValue("A2").asNumber()).toBe(84);
    });
  });
});
