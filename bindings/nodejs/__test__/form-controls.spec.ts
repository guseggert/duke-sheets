import { describe, expect, it } from "vitest";
import {
  JsCheckState,
  JsListSelection,
  Workbook,
  type JsFormControlInput,
} from "../index.js";
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
      state: JsCheckState.Checked,
      no3D: true,
    },
  },
  {
    anchor,
    kind: {
      kind: "optionButton",
      caption: "Option",
      state: JsCheckState.Unchecked,
      firstInGroup: false,
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
      selection: JsListSelection.Multi,
      selected: [1, 3],
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
      expect(reopened.formControls[1].kind.kind).toBe("checkbox");
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
          state: JsCheckState.Mixed,
          firstInGroup: false,
          no3D: false,
        },
      }),
    ).toThrow(/mixed state/);
  });
});
