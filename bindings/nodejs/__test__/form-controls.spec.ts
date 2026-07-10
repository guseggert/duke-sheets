import { describe, expect, it } from "vitest";
import { Workbook, type JsFormControlInput } from "../index.js";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

const anchor = {
  fromCol: 1,
  fromRow: 1,
  fromColOffset: 0,
  fromRowOffset: 0,
  toCol: 3,
  toRow: 2,
  toColOffset: 0,
  toRowOffset: 0,
  editAs: "twoCell",
};

const controls: JsFormControlInput[] = [
  { anchor, kind: { kind: "button", caption: "Run" } },
  {
    anchor,
    kind: {
      kind: "checkbox",
      caption: "Check",
      state: "checked",
      no3D: true,
    },
  },
  {
    anchor,
    kind: {
      kind: "optionButton",
      caption: "Option",
      state: "unchecked",
      no3D: false,
    },
  },
  { anchor, kind: { kind: "label", caption: "Label" } },
  { anchor, kind: { kind: "groupBox", caption: "Group", no3D: false } },
  {
    anchor,
    kind: {
      kind: "listBox",
      inputRange: "$A$1:$A$3",
      selection: "multi",
      // Zero-based: first and third items.
      selected: [0, 2],
      no3D: false,
    },
  },
  {
    anchor,
    kind: {
      kind: "dropdown",
      inputRange: "$A$1:$A$3",
      selected: 2,
      lines: 8,
      no3D: false,
    },
  },
  {
    anchor,
    kind: {
      kind: "scrollbar",
      value: 5,
      min: 0,
      max: 10,
      increment: 1,
      page: 2,
      horizontal: false,
    },
  },
  {
    anchor,
    kind: { kind: "spinner", value: 2, min: 0, max: 10, increment: 1 },
  },
];

describe("form controls", () => {
  it("adds, lists, sets, and removes controls", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    for (const control of controls) sheet.addFormControl(control);

    expect(sheet.formControlCount).toBe(9);
    expect(sheet.formControls.map((c) => c.kind.kind)).toEqual([
      "button",
      "checkbox",
      "optionButton",
      "label",
      "groupBox",
      "listBox",
      "dropdown",
      "scrollbar",
      "spinner",
    ]);

    sheet.setFormControl(0, { anchor, kind: { kind: "label", caption: "Replaced" } });
    expect(sheet.formControls[0].kind).toEqual({ kind: "label", caption: "Replaced" });
    sheet.removeFormControl(0);
    expect(sheet.formControlCount).toBe(8);
    expect(sheet.formControls[0].kind.kind).toBe("checkbox");
    expect(() => sheet.removeFormControl(99)).toThrow(/out of bounds/);
  });

  for (const ext of ["xlsx", "xlsb", "xls"] as const) {
    it(`round-trips controls through .${ext}`, () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      for (const control of controls) sheet.addFormControl(control);
      const dir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-controls-"));
      const file = path.join(dir, `controls.${ext}`);
      wb.save(file);
      const reopened = Workbook.open(file).getSheet(0);
      expect(reopened.formControlCount).toBe(9);
      // Writers recompute radio grouping; the sheet group's lone
      // radio heads its group after the trip.
      expect(reopened.formControls.map((c) => c.kind)).toEqual([
        { kind: "button", caption: "Run" },
        {
          kind: "checkbox",
          caption: "Check",
          state: "checked",
          no3D: true,
        },
        {
          kind: "optionButton",
          caption: "Option",
          state: "unchecked",
          firstInGroup: true,
          no3D: false,
        },
        { kind: "label", caption: "Label" },
        { kind: "groupBox", caption: "Group", no3D: false },
        {
          kind: "listBox",
          inputRange: "$A$1:$A$3",
          selection: "multi",
          selected: [0, 2],
          no3D: false,
        },
        {
          kind: "dropdown",
          inputRange: "$A$1:$A$3",
          selected: 2,
          lines: 8,
          no3D: false,
        },
        {
          kind: "scrollbar",
          value: 5,
          min: 0,
          max: 10,
          increment: 1,
          page: 2,
          horizontal: false,
        },
        { kind: "spinner", value: 2, min: 0, max: 10, increment: 1 },
      ]);
      fs.rmSync(dir, { recursive: true, force: true });
    });
  }

  it("rejects invalid variant data", () => {
    const sheet = new Workbook().getSheet(0);
    expect(() =>
      sheet.addFormControl({
        anchor,
        kind: {
          kind: "optionButton",
          caption: "Bad",
          state: "mixed",
          no3D: false,
        },
      }),
    ).toThrow(/mixed state/);
    expect(() =>
      sheet.addFormControl({
        anchor,
        kind: {
          kind: "listBox",
          selection: "multi",
          selected: [2, 1],
          no3D: false,
        },
      }),
    ).toThrow(/sorted and unique/);
  });
});
