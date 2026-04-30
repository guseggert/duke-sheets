import { describe, it, expect } from "vitest";
import { Workbook } from "../index.js";
import * as path from "node:path";
import * as os from "node:os";
import * as fs from "node:fs";

const PASSWORD = "duke-test-pw";

function buildSample(): Workbook {
  const wb = new Workbook();
  const sheet = wb.getSheet(0);
  if (!sheet) throw new Error("expected default sheet");
  sheet.setCell("A1", "hello");
  sheet.setCell("B1", 42.0);
  sheet.setCell("A2", 3.14);
  sheet.setCell("B2", true);
  return wb;
}

function tempPath(ext: string): string {
  return path.join(
    os.tmpdir(),
    `duke-pw-${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2)}${ext}`,
  );
}

function roundTrip(extension: string, profile?: string, opts?: { keyBits?: number; spinCount?: number }) {
  const file = tempPath(extension);
  try {
    const wb = buildSample();
    wb.saveWithPassword(file, PASSWORD, profile, opts?.keyBits, opts?.spinCount);
    const opened = Workbook.openWithPassword(file, PASSWORD);
    const sheet = opened.getSheet(0);
    if (!sheet) throw new Error("opened workbook has no sheet 0");
    expect(sheet.getCell("A1").asText()).toBe("hello");
    expect(sheet.getCell("B1").asNumber()).toBe(42);
  } finally {
    try { fs.unlinkSync(file); } catch {}
  }
}

describe("password-protected save/open", () => {
  describe("xlsx", () => {
    it("default profile round-trips", () => {
      roundTrip(".xlsx");
    });

    it("agile profile round-trips", () => {
      roundTrip(".xlsx", "agile", { keyBits: 256 });
    });

    it("standard profile round-trips", () => {
      roundTrip(".xlsx", "standard");
    });
  });

  describe("xls", () => {
    it("default profile round-trips", () => {
      roundTrip(".xls");
    });

    it("rc4-cryptoapi 128 round-trips", () => {
      roundTrip(".xls", "rc4-cryptoapi", { keyBits: 128 });
    });

    it("rc4-cryptoapi 40 round-trips", () => {
      roundTrip(".xls", "rc4-cryptoapi", { keyBits: 40 });
    });

    it("rc4-legacy round-trips", () => {
      roundTrip(".xls", "rc4-legacy");
    });

    it("xor round-trips via own reader", () => {
      roundTrip(".xls", "xor");
    });
  });

  describe("error paths", () => {
    it("wrong password rejects", () => {
      const file = tempPath(".xlsx");
      try {
        const wb = buildSample();
        wb.saveWithPassword(file, PASSWORD);
        expect(() => Workbook.openWithPassword(file, "wrong-password")).toThrow();
      } finally {
        try { fs.unlinkSync(file); } catch {}
      }
    });

    it("unknown profile rejects", () => {
      const file = tempPath(".xlsx");
      try {
        const wb = buildSample();
        expect(() => wb.saveWithPassword(file, PASSWORD, "not-a-thing")).toThrow();
      } finally {
        try { fs.unlinkSync(file); } catch {}
      }
    });
  });
});
