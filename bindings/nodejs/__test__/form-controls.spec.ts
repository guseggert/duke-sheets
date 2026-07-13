import { describe, expect, it } from "vitest";
import { Workbook, type FormControlKindInput, type TopLevelDrawingInput } from "../index.js";
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
} as const;

const caption = (value: string) => ({ runs: [{ text: value }] });
const drawing = (kind: FormControlKindInput): TopLevelDrawingInput => ({
  anchor,
  kind: "formControl",
  formControl: { kind },
});

const controls: TopLevelDrawingInput[] = [
  drawing({ kind: "button", caption: caption("Run") }),
  drawing({
    kind: "checkbox",
    caption: caption("Check"),
    state: "checked",
    cellLink: "$D$2",
    no3D: true,
  }),
  drawing({
    kind: "optionButton",
    caption: caption("Option"),
    state: "unchecked",
    no3D: false,
  }),
  drawing({ kind: "label", caption: caption("Label") }),
  drawing({ kind: "groupBox", caption: caption("Group"), no3D: false }),
  drawing({
    kind: "listBox",
    inputRange: "$A$1:$A$3",
    selection: "multi",
    selected: [0, 2],
    no3D: false,
  }),
  drawing({
    kind: "dropdown",
    inputRange: "$A$1:$A$3",
    selected: 2,
    lines: 8,
    no3D: false,
  }),
  drawing({
    kind: "scrollbar",
    value: 5,
    min: 0,
    max: 10,
    increment: 1,
    page: 2,
    horizontal: false,
  }),
  drawing({ kind: "spinner", value: 2, min: 0, max: 10, increment: 1 }),
];

describe("form-control drawings", () => {
  for (const ext of ["xlsx", "xlsb", "xls"] as const) {
    it(`round-trips controls through .${ext}`, () => {
      const wb = new Workbook();
      const sheet = wb.getSheet(0);
      for (const control of controls) sheet.addDrawing(control);
      const dir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-controls-"));
      const file = path.join(dir, `controls.${ext}`);
      wb.save(file);

      const reopened = Workbook.open(file).getSheet(0);
      expect(reopened.formControlCount).toBe(9);
      expect(reopened.formControls.map((control) => control.formControl.kind.kind)).toEqual([
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
      const checkbox = reopened.formControls[1].formControl.kind;
      if (checkbox.kind !== "checkbox") throw new Error("expected checkbox");
      expect(checkbox.caption.runs).toEqual([{ text: "Check" }]);
      expect(checkbox.state).toBe("checked");
      expect(reopened.getCell("D2").asBoolean()).toBe(true);
      fs.rmSync(dir, { recursive: true, force: true });
    });
  }

  it("rejects invalid control payloads", () => {
    const sheet = new Workbook().getSheet(0);
    expect(() =>
      sheet.addDrawing(
        drawing({
          kind: "optionButton",
          caption: caption("Bad"),
          state: "mixed" as "checked",
        }),
      ),
    ).toThrow(/mixed state/);
    expect(() =>
      sheet.addDrawing(
        drawing({
          kind: "listBox",
          selection: "multi",
          selected: [2, 1],
        }),
      ),
    ).toThrow(/sorted and unique/);
  });
});
