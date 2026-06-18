import { describe, expect, it } from "vitest";
import { Workbook } from "../index.js";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

function expectSameRgb(hex: string | undefined, rgb: string) {
  const normalized = hex?.length === 8 && hex.startsWith("FF") ? hex.slice(2) : hex;
  expect(normalized).toBe(rgb);
}

describe("Worksheet style setters", () => {
  it("sets font and fill styles on a cell", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setCellStyle("A1", {
      font: {
        name: "Aptos Display",
        size: 14,
        bold: true,
        color: { colorType: "rgb", hex: "FFFFFF" },
      },
      fill: {
        fillType: "solid",
        color: { hex: "1F4E79" },
      },
    });

    const style = sheet.getCellStyle("A1");
    expect(style).not.toBeNull();
    expect(style!.font.name).toBe("Aptos Display");
    expect(style!.font.size).toBe(14);
    expect(style!.font.bold).toBe(true);
    expect(style!.font.color.hex).toBe("FFFFFF");
    expect(style!.fill.fillType).toBe("solid");
    expect(style!.fill.color?.hex).toBe("1F4E79");
  });

  it("copies a full style returned by getCellStyle", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setCellStyle("A1", {
      font: { bold: true, italic: true },
      fill: { fillType: "solid", color: { hex: "D9EAF7" } },
    });
    sheet.setCellStyle("B1", {
      fill: { fillType: "solid", color: { hex: "00FF00" } },
    });

    const source = sheet.getCellStyle("A1");
    expect(source).not.toBeNull();
    sheet.setCellStyle("B1", source!);

    const copied = sheet.getCellStyle("B1");
    expect(copied!.font.bold).toBe(true);
    expect(copied!.font.italic).toBe(true);
    expect(copied!.fill.color?.hex).toBe("D9EAF7");
  });

  it("applies styles to every cell in a range", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    sheet.setRangeStyle("C1:D2", { font: { italic: true } });

    for (const addr of ["C1", "D1", "C2", "D2"]) {
      expect(sheet.getCellStyle(addr)!.font.italic).toBe(true);
    }
  });

  it("preserves style-only cells through XLSX roundtrip", () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-styles-"));
    const filePath = path.join(tmpDir, "styles.xlsx");

    try {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      sheet.setCellStyle("B2", {
        fill: { fillType: "solid", color: { hex: "FFF2CC" } },
        font: { color: { hex: "9C5700" } },
      });
      wb.save(filePath);

      const wb2 = Workbook.open(filePath);
      const style = wb2.getSheet(0).getCellStyle("B2");
      expect(style).not.toBeNull();
      expectSameRgb(style!.fill.color?.hex, "FFF2CC");
      expectSameRgb(style!.font.color.hex, "9C5700");
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("throws useful errors for invalid style values", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);

    expect(() =>
      sheet.setCellStyle("A1", { fill: { fillType: "notAType" } } as any),
    ).toThrow(/unknown fillType/);
  });
});
