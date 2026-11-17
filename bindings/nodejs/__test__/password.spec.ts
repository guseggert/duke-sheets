import { describe, it, expect } from "vitest";
import { Workbook, decryptAsync } from "../index.js";
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

  describe("decryptAsync", () => {
    const PROFILES = [
      { extension: ".xlsx", profile: "agile" },
      { extension: ".xlsx", profile: "standard" },
      { extension: ".xls", profile: "rc4-cryptoapi" },
      { extension: ".xls", profile: "rc4-legacy" },
      { extension: ".xls", profile: "xor" },
    ];

    function encryptedBytes(extension: string, profile: string, password: string): Buffer {
      const file = tempPath(extension);
      try {
        buildSample().saveWithPassword(file, password, profile, undefined, 1000);
        return fs.readFileSync(file);
      } finally {
        try { fs.unlinkSync(file); } catch {}
      }
    }

    function expectSample(decrypted: Buffer | null) {
      if (!decrypted) throw new Error("expected decrypted bytes");
      // fromBytes refuses encrypted workbooks, so this proves the bytes are plain.
      const sheet = Workbook.fromBytes(decrypted).getSheet(0);
      if (!sheet) throw new Error("decrypted workbook has no sheet 0");
      expect(sheet.getCell("A1").asText()).toBe("hello");
      expect(sheet.getCell("B1").asNumber()).toBe(42);
    }

    it.each(PROFILES)("decrypts $profile with its password", async ({ extension, profile }) => {
      const encrypted = encryptedBytes(extension, profile, PASSWORD);
      expectSample(await decryptAsync(encrypted, PASSWORD));
    });

    it.each(PROFILES)("rejects a wrong password for $profile", async ({ extension, profile }) => {
      const encrypted = encryptedBytes(extension, profile, PASSWORD);
      await expect(decryptAsync(encrypted, "wrong-password")).rejects.toThrow("Incorrect password");
    });

    it.each(PROFILES)(
      "decrypts $profile encrypted with Excel's default password when no password is given",
      async ({ extension, profile }) => {
        const encrypted = encryptedBytes(extension, profile, "VelvetSweatshop");
        expectSample(await decryptAsync(encrypted));
      },
    );

    it("rejects an encrypted file when no password is given", async () => {
      const encrypted = encryptedBytes(".xlsx", "agile", PASSWORD);
      await expect(decryptAsync(encrypted)).rejects.toThrow("no password was supplied");
    });

    it.each([".xlsx", ".xls"])("resolves to null for an unencrypted %s", async (extension) => {
      const file = tempPath(extension);
      try {
        buildSample().save(file);
        expect(await decryptAsync(fs.readFileSync(file), PASSWORD)).toBeNull();
      } finally {
        try { fs.unlinkSync(file); } catch {}
      }
    });
  });
});
