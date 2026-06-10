import { describe, it, expect, beforeAll } from "vitest";
import { Workbook } from "../index.js";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import * as path from "node:path";
import * as os from "node:os";
import * as fs from "node:fs";

// Populated-feature reads: comments, autofilters, data validations, and
// embedded images are read-only in the binding, so fixtures carrying
// them are generated at test time by the Rust fixture generator
// (binary fixtures are never committed).

const REPO_ROOT = path.resolve(__dirname, "..", "..", "..");
let fixtureDir: string;

beforeAll(async () => {
  fixtureDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-fixtures-"));
  // Async on purpose: a long synchronous child process starves the
  // vitest worker's IPC heartbeat and fails the run with
  // `Timeout calling "onTaskUpdate"` even after all tests pass.
  await promisify(execFile)(
    "cargo",
    [
      "run",
      "-p",
      "duke-sheets",
      "--features",
      "full",
      "--example",
      "gen_binding_fixtures",
      "--",
      fixtureDir,
    ],
    { cwd: REPO_ROOT },
  );
}, 600_000);

function fixture(ext: string): string {
  return path.join(fixtureDir, `sample.${ext}`);
}

for (const ext of ["xlsx", "xls", "xlsb"]) {
  describe(`populated feature reads (.${ext})`, () => {
    it("reads the comment", () => {
      const sheet = Workbook.open(fixture(ext)).getSheet(0)!;
      const comment = sheet.getComment("A1");
      expect(comment).not.toBeNull();
      expect(comment!.author).toBe("Tester");
      expect(comment!.text).toBe("fixture comment");
      expect(sheet.commentCount).toBe(1);
    });

    it("reads the autofilter", () => {
      const sheet = Workbook.open(fixture(ext)).getSheet(0)!;
      const af = sheet.autoFilter;
      expect(af).not.toBeNull();
      expect(af!.range).toBe("A1:A4");
      expect(af!.filterColumns).toHaveLength(1);
      const col = af!.filterColumns[0];
      expect(col.colId).toBe(0);
      expect(col.filterType).toBe("values");
      expect(col.values).toEqual(["1", "3"]);
    });

    it("reads the data validation", () => {
      const sheet = Workbook.open(fixture(ext)).getSheet(0)!;
      const dvs = sheet.dataValidations;
      expect(dvs).toHaveLength(1);
      expect(dvs[0].validationType).toBe("list");
      expect(dvs[0].listSource).toBe("Red,Green,Blue");
      expect(dvs[0].ranges).toContain("C1:C5");
    });

    it("reads values and the named-range formula", () => {
      const sheet = Workbook.open(fixture(ext)).getSheet(0)!;
      expect(sheet.getCell("A1").asText()).toBe("Score");
      expect(sheet.getFormulaAt(0, 1)).toBe("=SUM(MyRange)");
    });

    if (ext !== "xlsb") {
      it("reads the embedded image", () => {
        const sheet = Workbook.open(fixture(ext)).getSheet(0)!;
        expect(sheet.imageCount).toBe(1);
        const img = sheet.images[0];
        expect(img.name).toBe("FixturePic");
        expect(img.format.toLowerCase()).toBe("png");
        expect(img.data.length).toBeGreaterThan(0);
        // PNG magic survives the container round-trip.
        expect(img.data[0]).toBe(0x89);
        expect(img.data[1]).toBe(0x50);
      });
    }
  });
}
