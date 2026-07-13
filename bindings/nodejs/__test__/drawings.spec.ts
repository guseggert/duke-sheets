import { describe, expect, it } from "vitest";
import {
  Workbook,
  type DrawingAnchor,
  type DrawingText,
  type TopLevelDrawingInput,
} from "../index.js";

function anchor(fromCol: number, fromRow: number, toCol: number, toRow: number): DrawingAnchor {
  return {
    fromCol,
    fromRow,
    fromColOffset: 0,
    fromRowOffset: 0,
    toCol,
    toRow,
    toColOffset: 0,
    toRowOffset: 0,
    editAs: "twoCell",
  };
}

function transform(xEmu = 0, yEmu = 0) {
  return { xEmu, yEmu, cxEmu: 100_000, cyEmu: 50_000 };
}

function text(value: string): DrawingText {
  return { runs: [{ text: value }] };
}

function shape(name: string, drawingAnchor: DrawingAnchor): TopLevelDrawingInput {
  return { name, anchor: drawingAnchor, kind: "shape", shape: { geometry: "rect" } };
}

describe("unified drawings", () => {
  it("preserves z-order, recursive paths, metadata, and lazy image bytes", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    const png = Buffer.from([0x89, 0x50, 0x4e, 0x47]);

    sheet.addDrawing({
      name: "Picture",
      anchor: anchor(0, 0, 2, 2),
      kind: "image",
      image: { format: "png", widthEmu: 200_000, heightEmu: 100_000, data: png },
    });
    sheet.addDrawing({
      name: "Run button",
      hidden: true,
      locked: false,
      printable: false,
      altText: "Runs the report",
      title: "Report action",
      anchor: anchor(2, 0, 4, 1),
      kind: "formControl",
      formControl: {
        kind: {
          kind: "button",
          caption: {
            runs: [{ text: "Run " }, { text: "now", font: { bold: true, size: 14 } }],
            horizontalAlignment: "center",
          },
        },
        macroName: "RunReport",
      },
    });
    sheet.addDrawing(shape("Front shape", anchor(4, 0, 6, 2)));
    sheet.addDrawing({
      name: "Group",
      anchor: anchor(0, 3, 5, 8),
      kind: "group",
      group: {
        children: [
          {
            transform: transform(),
            kind: "formControl",
            formControl: { kind: { kind: "label", caption: text("Nested") } },
          },
          {
            transform: transform(100_000, 100_000),
            kind: "group",
            group: {
              children: [
                {
                  transform: transform(10_000, 20_000),
                  kind: "shape",
                  shape: { geometry: "ellipse" },
                },
              ],
            },
          },
        ],
      },
    });

    expect(sheet.drawings.map((drawing) => drawing.kind)).toEqual([
      "image",
      "formControl",
      "shape",
      "group",
    ]);
    expect(sheet.drawings.map((drawing) => drawing.drawingPath)).toEqual([[0], [1], [2], [3]]);
    const group = sheet.drawings[3];
    expect(group.kind).toBe("group");
    if (group.kind !== "group") throw new Error("expected group");
    expect(group.group.children[0].drawingPath).toEqual([3, 0]);
    const nested = group.group.children[1];
    if (nested.kind !== "group") throw new Error("expected nested group");
    expect(nested.group.children[0].drawingPath).toEqual([3, 1, 0]);
    expect(sheet.formControls.map((control) => control.drawingPath)).toEqual([[1], [3, 0]]);

    const image = sheet.images[0];
    expect("data" in image.image).toBe(false);
    expect(sheet.drawingImageData(image.drawingPath)).toEqual(png);
    expect(sheet.drawingSvgData(image.drawingPath)).toBeNull();

    const button = sheet.formControls[0];
    expect(button).toMatchObject({
      name: "Run button",
      hidden: true,
      locked: false,
      printable: false,
      altText: "Runs the report",
      title: "Report action",
    });
    expect(button.formControl.macroName).toBe("RunReport");
    const caption = button.formControl.kind;
    if (caption.kind !== "button") throw new Error("expected button");
    expect(caption.caption.runs.map((run) => run.text).join("")).toBe("Run now");
    expect(caption.caption.runs[1].font).toMatchObject({ bold: true, size: 14 });
  });

  it("sets, removes, inserts, and moves generic drawings", () => {
    const sheet = new Workbook().getSheet(0);
    sheet.addDrawing(shape("one", anchor(0, 0, 1, 1)));
    sheet.addDrawing(shape("three", anchor(2, 0, 3, 1)));
    sheet.insertDrawing(1, shape("two", anchor(1, 0, 2, 1)));
    sheet.moveDrawing(2, 0);
    expect(sheet.drawings.map((drawing) => drawing.name)).toEqual(["three", "one", "two"]);

    sheet.setDrawing([1], shape("replaced", anchor(0, 1, 1, 2)));
    sheet.addDrawing({
      anchor: anchor(3, 0, 5, 2),
      kind: "group",
      group: {
        children: [{ transform: transform(), kind: "shape", shape: { geometry: "rect" } }],
      },
    });
    sheet.setDrawing([3, 0], {
      name: "nested label",
      transform: transform(5, 5),
      kind: "formControl",
      formControl: { kind: { kind: "label", caption: text("replacement") } },
    });
    expect(sheet.formControls[0].drawingPath).toEqual([3, 0]);
    sheet.removeDrawing([3, 0]);
    expect(sheet.formControlCount).toBe(0);
    sheet.removeDrawing([2]);
    expect(sheet.drawings.map((drawing) => drawing.name)).toEqual(["three", "replaced", undefined]);
  });

  it("applies radio semantics and updates the linked cell", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.addDrawing({
      anchor: anchor(0, 0, 4, 6),
      kind: "formControl",
      formControl: { kind: { kind: "groupBox", caption: text("Choose") } },
    });
    for (const [name, state, row] of [
      ["one", "checked", 1],
      ["two", "unchecked", 3],
    ] as const) {
      sheet.addDrawing({
        name,
        anchor: anchor(1, row, 2, row + 1),
        kind: "formControl",
        formControl: {
          kind: { kind: "optionButton", caption: text(name), state, cellLink: "$D$2" },
        },
      });
    }

    expect(wb.syncFormControls()).toBe(1);
    expect(sheet.getCell("D2").asNumber()).toBe(1);
    const result = sheet.setFormControlCheckState([2], "checked");
    expect(result).toEqual({ controlsChanged: 2, linkedCellsChanged: 1 });
    expect(sheet.getCell("D2").asNumber()).toBe(2);
    const radios = sheet.formControls.slice(1).map((control) => control.formControl.kind);
    expect(radios.map((radio) => "state" in radio ? radio.state : undefined)).toEqual([
      "unchecked",
      "checked",
    ]);
  });
});
