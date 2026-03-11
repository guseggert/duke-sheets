import { describe, it, expect } from "vitest";
import { Workbook } from "../index.js";
import * as path from "node:path";
import * as os from "node:os";
import * as fs from "node:fs";

// The async open/fromBytes are free functions, import them
const binding = require("../index.js");
const openAsync: (path: string) => Promise<any> = binding.openAsync;
const fromBytesAsync: (data: Buffer) => Promise<any> =
  binding.fromBytesAsync;

describe("Async open", () => {
  it("openAsync loads a saved file", async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-async-"));
    const filePath = path.join(tmpDir, "test.xlsx");

    try {
      // Create and save a workbook synchronously
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      sheet.setCell("A1", 42);
      sheet.setCell("B1", "hello");
      wb.save(filePath);

      // Open it asynchronously
      const wb2 = await openAsync(filePath);
      expect(wb2).toBeDefined();

      const sheet2 = wb2.getSheet(0);
      expect(sheet2.getCell("A1").asNumber()).toBe(42);
      expect(sheet2.getCell("B1").asText()).toBe("hello");
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("openAsync rejects for non-existent file", async () => {
    await expect(openAsync("/no/such/file.xlsx")).rejects.toThrow();
  });
});

describe("Async fromBytes", () => {
  it("fromBytesAsync loads from buffer", async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-async-"));
    const filePath = path.join(tmpDir, "test.xlsx");

    try {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      sheet.setCell("A1", 99);
      wb.save(filePath);

      const buf = fs.readFileSync(filePath);
      const wb2 = await fromBytesAsync(buf);
      expect(wb2).toBeDefined();

      const sheet2 = wb2.getSheet(0);
      expect(sheet2.getCell("A1").asNumber()).toBe(99);
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("fromBytesAsync rejects for invalid bytes", async () => {
    await expect(
      fromBytesAsync(Buffer.from("not xlsx")),
    ).rejects.toThrow();
  });
});

describe("Async save", () => {
  it("saveAsync writes file to disk", async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-async-"));
    const filePath = path.join(tmpDir, "async-save.xlsx");

    try {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      sheet.setCell("A1", 123);

      await wb.saveAsync(filePath);

      expect(fs.existsSync(filePath)).toBe(true);

      // Verify the saved file is valid
      const wb2 = Workbook.open(filePath);
      const sheet2 = wb2.getSheet(0);
      expect(sheet2.getCell("A1").asNumber()).toBe(123);
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("saveAsync rejects for invalid path", async () => {
    const wb = new Workbook();
    await expect(
      wb.saveAsync("/no/such/dir/file.xlsx"),
    ).rejects.toThrow();
  });
});

describe("Async calculate", () => {
  it("calculateAsync returns stats", async () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 10);
    sheet.setCell("A2", 20);
    sheet.setFormula("A3", "=A1+A2");

    const stats = await wb.calculateAsync();
    expect(stats).toBeDefined();
    expect(stats.formulaCount).toBeGreaterThanOrEqual(1);
    expect(stats.cellsCalculated).toBeGreaterThanOrEqual(1);

    // Verify the formula was actually calculated
    expect(sheet.getCalculatedValue("A3").asNumber()).toBe(30);
  });

  it("calculateWithOptionsAsync works with iterative", async () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.setCell("A1", 1);
    sheet.setFormula("A2", "=A1*2");

    const stats = await wb.calculateWithOptionsAsync(false, 100, 0.001);
    expect(stats).toBeDefined();
    expect(stats.formulaCount).toBeGreaterThanOrEqual(1);

    expect(sheet.getCalculatedValue("A2").asNumber()).toBe(2);
  });
});

describe("Async roundtrip", () => {
  it("full async workflow: open -> calculate -> save", async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-async-"));
    const srcPath = path.join(tmpDir, "source.xlsx");
    const dstPath = path.join(tmpDir, "calculated.xlsx");

    try {
      // Create source file
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      sheet.setCell("A1", 5);
      sheet.setCell("A2", 10);
      sheet.setFormula("A3", "=SUM(A1:A2)");
      wb.save(srcPath);

      // Full async workflow
      const wb2 = await openAsync(srcPath);
      await wb2.calculateAsync();
      await wb2.saveAsync(dstPath);

      // Verify
      const wb3 = Workbook.open(dstPath);
      const sheet3 = wb3.getSheet(0);
      expect(sheet3.getCalculatedValue("A3").asNumber()).toBe(15);
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });
});
