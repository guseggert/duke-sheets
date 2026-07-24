import { describe, expect, it } from "vitest";

import { Workbook } from "../lib.js";

describe("Worksheet.toCsvString", () => {
  it("exports any worksheet independently", () => {
    const workbook = new Workbook();
    workbook.addSheet("Second");
    const first = workbook.getSheet(0);
    const second = workbook.getSheet(1);

    first.setCell("A1", "first");
    second.setCell("A1", "second");
    second.setCell("B1", "值");

    expect(first.toCsvString()).toBe("first\r\n");
    expect(second.toCsvString()).toBe("second,值\r\n");
  });

  it("matches the first-sheet workbook compatibility method", () => {
    const workbook = new Workbook();
    const sheet = workbook.getSheet(0);
    sheet.setCell("A1", "plain");
    sheet.setCell("B1", "say \"hi\", 世界");
    sheet.setCell("A2", "line 1\nline 2");
    sheet.setCell("B2", "😀");

    expect(sheet.toCsvString()).toBe(
      "plain,\"say \"\"hi\"\", 世界\"\r\n\"line 1\nline 2\",😀\r\n",
    );
    expect(sheet.toCsvString()).toBe(workbook.saveCsvString());
  });

  it("synchronizes linked form-control state into the exported worksheet", () => {
    const workbook = new Workbook();
    const sheet = workbook.getSheet(0);
    sheet.setCell("A1", "marker");
    sheet.addDrawing({
      anchor: {
        type: "twoCell",
        from: { col: 1, row: 0 },
        to: { col: 2, row: 1 },
        editAs: "twoCell",
      },
      kind: "formControl",
      formControl: {
        kind: {
          kind: "checkbox",
          caption: { runs: [{ text: "Check" }] },
          state: "checked",
          cellLink: "$D$2",
        },
      },
    });

    expect(sheet.toCsvString()).toBe("marker,,,\r\n,,,TRUE\r\n");
    expect(sheet.toCsvString()).toBe(workbook.saveCsvString());
  });
});
