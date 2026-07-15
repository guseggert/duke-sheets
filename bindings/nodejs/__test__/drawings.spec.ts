import { describe, expect, it } from "vitest";
import * as path from "node:path";
import * as os from "node:os";
import * as fs from "node:fs";
import {
  Workbook,
  type DrawingAnchor,
  type DrawingText,
  type TopLevelDrawingInput,
} from "../index.js";

function anchor(fromCol: number, fromRow: number, toCol: number, toRow: number): DrawingAnchor {
  return {
    type: "twoCell",
    from: { col: fromCol, row: fromRow },
    to: { col: toCol, row: toRow },
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
  it("resolves absoluteRectEmu for top-level drawings and group children", () => {
    const wb = new Workbook();
    const sheet = wb.getSheet(0);
    sheet.addDrawing({
      anchor: anchor(0, 0, 1, 1),
      kind: "formControl",
      formControl: { kind: { kind: "label", caption: text("top") } },
    });
    sheet.addDrawing({
      anchor: anchor(0, 0, 4, 4),
      kind: "group",
      group: {
        groupTransform: { childXEmu: 0, childYEmu: 0, childCxEmu: 1000, childCyEmu: 1000 },
        children: [
          {
            transform: { xEmu: 250, yEmu: 500, cxEmu: 500, cyEmu: 250 },
            kind: "formControl",
            formControl: { kind: { kind: "label", caption: text("nested") } },
          },
        ],
      },
    });

    const drawings = sheet.drawings;
    const top = drawings[0].absoluteRectEmu;
    expect(top.xEmu).toBe(0);
    expect(top.yEmu).toBe(0);
    expect(top.widthEmu).toBeGreaterThan(0);
    expect(top.heightEmu).toBeGreaterThan(0);

    const groupRect = drawings[1].absoluteRectEmu;
    const controls = sheet.formControls;
    expect(controls).toHaveLength(2);
    const child = controls[1].absoluteRectEmu;
    expect(child.xEmu).toBe(groupRect.xEmu + groupRect.widthEmu * 0.25);
    expect(child.yEmu).toBe(groupRect.yEmu + groupRect.heightEmu * 0.5);
    expect(child.widthEmu).toBe(groupRect.widthEmu * 0.5);
    expect(child.heightEmu).toBe(groupRect.heightEmu * 0.25);

    const drawn = drawings[1];
    if (drawn.kind !== "group") throw new Error("expected group");
    expect(drawn.group.children[0].absoluteRectEmu).toEqual(child);
  });

  it("exposes the theme palette and resolves colors against it", () => {
    const wb = new Workbook();
    expect(wb.themePalette).toHaveLength(12);
    expect(wb.themePalette[4]).toBe("4F81BD");
    expect(wb.resolveColor({ colorType: "theme", index: 4, tint: 0 })).toBe("4F81BD");
    expect(wb.resolveColor({ colorType: "theme", index: 4, tint: 50 })).toBe("A7C0DE");
    expect(wb.resolveColor({ colorType: "rgb", r: 1, g: 2, b: 3 })).toBe("010203");
    expect(wb.resolveColor({ colorType: "auto" })).toBeNull();
  });

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

  it("round-trips a oneCell anchor through save, setDrawing, and save again", () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-anchors-"));
    const filePath = path.join(tmpDir, "one-cell.xlsx");
    try {
      const wb = new Workbook();
      wb.getSheet(0).addDrawing({
        name: "Pinned",
        anchor: {
          type: "oneCell",
          from: { col: 1, row: 2, colOffsetEmu: 9_525, rowOffsetEmu: 19_050 },
          widthEmu: 300_000,
          heightEmu: 200_000,
        },
        kind: "shape",
        shape: { geometry: "rect" },
      });
      wb.save(filePath);

      const first = Workbook.fromBytes(fs.readFileSync(filePath));
      const drawing = first.getSheet(0).drawings[0];
      if (drawing.kind !== "shape") throw new Error("expected shape");
      const expected = {
        type: "oneCell",
        from: { col: 1, row: 2, colOffsetEmu: 9_525, rowOffsetEmu: 19_050 },
        widthEmu: 300_000,
        heightEmu: 200_000,
      };
      expect(drawing.anchor).toEqual(expected);

      // Identity rewrite must not rewrite the anchor variant or extent.
      first.getSheet(0).setDrawing([0], drawing);
      expect(first.getSheet(0).drawings[0].anchor).toEqual(expected);
      first.save(filePath);

      const second = Workbook.fromBytes(fs.readFileSync(filePath));
      expect(second.getSheet(0).drawings[0].anchor).toEqual(expected);
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("round-trips absolute anchors and keeps twoCell markers tagged", () => {
    const sheet = new Workbook().getSheet(0);
    sheet.addDrawing({
      anchor: { type: "absolute", xEmu: 100, yEmu: 200, widthEmu: 300_000, heightEmu: 400_000 },
      kind: "shape",
      shape: { geometry: "rect" },
    });
    sheet.addDrawing(shape("plain", anchor(1, 1, 3, 4)));

    expect(sheet.drawings[0].anchor).toEqual({
      type: "absolute",
      xEmu: 100,
      yEmu: 200,
      widthEmu: 300_000,
      heightEmu: 400_000,
    });
    expect(sheet.drawings[1].anchor).toEqual({
      type: "twoCell",
      from: { col: 1, row: 1, colOffsetEmu: 0, rowOffsetEmu: 0 },
      to: { col: 3, row: 4, colOffsetEmu: 0, rowOffsetEmu: 0 },
      editAs: "twoCell",
    });
  });

  it("rejects comments as group children in setDrawing", () => {
    const sheet = new Workbook().getSheet(0);
    sheet.addDrawing({
      anchor: anchor(0, 0, 2, 2),
      kind: "group",
      group: {
        children: [{ transform: transform(), kind: "shape", shape: { geometry: "rect" } }],
      },
    });
    expect(() =>
      sheet.setDrawing([0, 0], {
        transform: transform(),
        kind: "comment",
        comment: { row: 0, col: 0, author: "a", text: "nested" },
      }),
    ).toThrow(/comments cannot be group children/);
  });

  it("preserves unknown-control passthrough data through read, setDrawing, and save", () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-drawings-"));
    const filePath = path.join(tmpDir, "unknown.xlsx");
    try {
      const wb = new Workbook();
      wb.getSheet(0).addDrawing({
        name: "Legacy editor",
        anchor: anchor(1, 2, 3, 4),
        kind: "formControl",
        formControl: {
          kind: {
            kind: "unknown",
            objectType: "EditBox",
            caption: text("Unsupported editor"),
            rawProperties: [
              ["customFlag", "kept"],
              ["val", "17"],
              ["fmlaLink", "$A$1"],
            ],
          },
          rawClientData: [Buffer.from("<x:Val>17</x:Val>")],
        },
      });
      wb.save(filePath);

      const readUnknown = (workbook: Workbook) => {
        const control = workbook.getSheet(0).formControls[0];
        const kind = control.formControl.kind;
        if (kind.kind !== "unknown") throw new Error("expected unknown control");
        return { control, kind };
      };

      const first = Workbook.fromBytes(fs.readFileSync(filePath));
      const { control, kind } = readUnknown(first);
      expect(kind.objectType).toBe("EditBox");
      expect(kind.rawProperties).toContainEqual(["customFlag", "kept"]);
      const propertyCount = kind.rawProperties.length;
      expect(propertyCount).toBeGreaterThanOrEqual(3);
      expect(control.formControl.rawClientData.length).toBeGreaterThanOrEqual(1);

      // Identity rewrite: the read snapshot keeps its passthrough data.
      // The narrowed `kind` rebuilds the payload because output states
      // (e.g. optionButton "mixed") are wider than accepted inputs.
      first
        .getSheet(0)
        .setDrawing([0], { ...control, formControl: { ...control.formControl, kind } });
      first.save(filePath);

      const second = readUnknown(Workbook.fromBytes(fs.readFileSync(filePath)));
      expect(second.kind.objectType).toBe("EditBox");
      expect(second.kind.rawProperties.length).toBe(propertyCount);
      expect(second.kind.rawProperties).toContainEqual(["customFlag", "kept"]);
      expect(
        second.control.formControl.rawClientData.some((fragment) =>
          Buffer.from(fragment).toString("utf8").includes("<x:Val>17</x:Val>"),
        ),
      ).toBe(true);
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("preserves rawClientData on modeled control kinds through save and setDrawing", () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "duke-drawings-"));
    const filePath = path.join(tmpDir, "modeled-raws.xlsx");
    try {
      const wb = new Workbook();
      wb.getSheet(0).addDrawing({
        anchor: anchor(1, 1, 3, 3),
        kind: "formControl",
        formControl: {
          kind: { kind: "checkbox", caption: text("Audit"), state: "checked" },
          rawClientData: [Buffer.from("<x:Disabled/>"), Buffer.from("<x:Accel>65</x:Accel>")],
        },
      });
      wb.save(filePath);

      const readRaws = (workbook: Workbook) =>
        workbook
          .getSheet(0)
          .formControls[0].formControl.rawClientData.map((fragment) =>
            Buffer.from(fragment).toString("utf8"),
          );

      const first = Workbook.fromBytes(fs.readFileSync(filePath));
      expect(readRaws(first)).toEqual(["<x:Disabled/>", "<x:Accel>65</x:Accel>"]);

      const control = first.getSheet(0).formControls[0];
      const kind = control.formControl.kind;
      if (kind.kind !== "checkbox") throw new Error("expected checkbox");
      first
        .getSheet(0)
        .setDrawing([0], { ...control, formControl: { ...control.formControl, kind } });
      first.save(filePath);
      const second = Workbook.fromBytes(fs.readFileSync(filePath));
      expect(readRaws(second)).toEqual(["<x:Disabled/>", "<x:Accel>65</x:Accel>"]);
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("rejects a setDrawing comment replacement that duplicates another cell's comment", () => {
    const sheet = new Workbook().getSheet(0);
    sheet.addDrawing({
      anchor: anchor(2, 0, 4, 2),
      kind: "comment",
      comment: { row: 0, col: 0, author: "a", text: "first" },
    });
    sheet.addDrawing({
      anchor: anchor(2, 4, 4, 6),
      kind: "comment",
      comment: { row: 4, col: 0, author: "a", text: "second" },
    });

    expect(() =>
      sheet.setDrawing([1], {
        anchor: anchor(2, 4, 4, 6),
        kind: "comment",
        comment: { row: 0, col: 0, author: "a", text: "duplicate" },
      }),
    ).toThrow(/already has a comment/);

    // Replacing a comment in place (same cell) stays allowed.
    sheet.setDrawing([0], {
      anchor: anchor(2, 0, 4, 2),
      kind: "comment",
      comment: { row: 0, col: 0, author: "a", text: "updated" },
    });
    expect(sheet.drawings.filter((drawing) => drawing.kind === "comment")).toHaveLength(2);
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
