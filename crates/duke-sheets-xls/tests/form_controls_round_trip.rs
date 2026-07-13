//! Round-trip tests for BIFF8 form controls.
//!
//! Exercises the in-process loop: build a workbook with form
//! controls, write to BIFF8 bytes, read back, assert kind-specific
//! properties, captions, cell links, and anchors survive.

use std::io::Cursor;

use duke_sheets_chart::{CellMarker, DrawingAnchor};
use duke_sheets_core::{
    CheckState, ControlText, DrawingObject, Drawn, FormControl, FormControlKind, ListSelection,
    RichTextRun, RunFont, Workbook, Worksheet,
};
use duke_sheets_xls::{XlsReader, XlsWriter};

fn write_then_read(wb: &Workbook) -> Workbook {
    let bytes = XlsWriter::write_to_bytes(wb).expect("serialize");
    XlsReader::read(Cursor::new(&bytes)).expect("read back")
}

fn anchor(from_col: u16, from_row: u32, to_col: u16, to_row: u32) -> DrawingAnchor {
    DrawingAnchor::TwoCell {
        from: CellMarker {
            col: from_col,
            col_offset_emu: 0,
            row: from_row,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: to_col,
            col_offset_emu: 0,
            row: to_row,
            row_offset_emu: 0,
        },
        edit_as: None,
    }
}

fn control_at(kind: FormControlKind, anchor: DrawingAnchor) -> DrawingObject {
    DrawingObject::form_control(FormControl::new(kind)).with_anchor(anchor)
}

fn controls_of(ws: &Worksheet) -> Vec<Drawn<'_, FormControl>> {
    ws.form_controls().collect()
}

fn single_control_round_trip(kind: FormControlKind) -> FormControlKind {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").expect("A1");
    ws.add_drawing(control_at(kind, anchor(1, 1, 3, 3)));

    let parsed = write_then_read(&wb);
    let controls = controls_of(parsed.worksheet(0).unwrap());
    assert_eq!(controls.len(), 1, "control must survive round-trip");
    controls[0].payload.kind.clone()
}

#[test]
fn button_round_trips() {
    let kind = single_control_round_trip(FormControlKind::Button {
        caption: "Run Report".into(),
    });
    assert_eq!(
        kind,
        FormControlKind::Button {
            caption: "Run Report".into(),
        }
    );
}

#[test]
fn checkbox_round_trips() {
    let kind = single_control_round_trip(FormControlKind::Checkbox {
        caption: "Enable audit".into(),
        state: CheckState::Checked,
        cell_link: Some("$D$2".to_string()),
        no_3d: false,
    });
    assert_eq!(
        kind,
        FormControlKind::Checkbox {
            caption: "Enable audit".into(),
            state: CheckState::Checked,
            cell_link: Some("$D$2".to_string()),
            no_3d: false,
        }
    );
}

#[test]
fn checkbox_mixed_state_round_trips() {
    let kind = single_control_round_trip(FormControlKind::Checkbox {
        caption: "Tri".into(),
        state: CheckState::Mixed,
        cell_link: None,
        no_3d: true,
    });
    assert_eq!(
        kind,
        FormControlKind::Checkbox {
            caption: "Tri".into(),
            state: CheckState::Mixed,
            cell_link: None,
            no_3d: true,
        }
    );
}

#[test]
fn option_buttons_round_trip_as_group() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for (i, (caption, state)) in [
        ("Opt A", CheckState::Checked),
        ("Opt B", CheckState::Unchecked),
        ("Opt C", CheckState::Unchecked),
    ]
    .into_iter()
    .enumerate()
    {
        ws.add_drawing(control_at(
            FormControlKind::OptionButton {
                caption: caption.into(),
                state,
                cell_link: Some("$E$1".to_string()),
                first_in_group: false,
                no_3d: false,
            },
            anchor(1, 1 + i as u32, 3, 2 + i as u32),
        ));
    }

    let parsed = write_then_read(&wb);
    let controls = controls_of(parsed.worksheet(0).unwrap());
    assert_eq!(controls.len(), 3);
    for (i, c) in controls.iter().enumerate() {
        match &c.payload.kind {
            FormControlKind::OptionButton {
                caption,
                state,
                cell_link,
                first_in_group,
                ..
            } => {
                assert_eq!(caption.plain_text(), format!("Opt {}", ["A", "B", "C"][i]));
                assert_eq!(
                    *state,
                    if i == 0 {
                        CheckState::Checked
                    } else {
                        CheckState::Unchecked
                    }
                );
                assert_eq!(cell_link.as_deref(), Some("$E$1"));
                // The writer recomputes grouping: first radio carries
                // the flag.
                assert_eq!(*first_in_group, i == 0, "radio {i} fFirstBtn");
            }
            other => panic!("expected OptionButton, got {other:?}"),
        }
    }
}

#[test]
fn label_round_trips() {
    let kind = single_control_round_trip(FormControlKind::Label {
        caption: "Status".into(),
    });
    assert_eq!(
        kind,
        FormControlKind::Label {
            caption: "Status".into(),
        }
    );
}

#[test]
fn group_box_round_trips() {
    let kind = single_control_round_trip(FormControlKind::GroupBox {
        caption: "Choices".into(),
        no_3d: true,
    });
    assert_eq!(
        kind,
        FormControlKind::GroupBox {
            caption: "Choices".into(),
            no_3d: true,
        }
    );
}

#[test]
fn list_box_round_trips() {
    let kind = single_control_round_trip(FormControlKind::ListBox {
        input_range: Some("$H$1:$H$4".to_string()),
        cell_link: Some("$D$5".to_string()),
        selection: ListSelection::Single,
        selected: vec![3],
        no_3d: true,
    });
    assert_eq!(
        kind,
        FormControlKind::ListBox {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: Some("$D$5".to_string()),
            selection: ListSelection::Single,
            selected: vec![3],
            no_3d: true,
        }
    );
}

#[test]
fn list_box_multi_select_round_trips() {
    let kind = single_control_round_trip(FormControlKind::ListBox {
        input_range: Some("$H$1:$H$5".to_string()),
        cell_link: None,
        selection: ListSelection::Multi,
        selected: vec![0, 2, 4],
        no_3d: false,
    });
    assert_eq!(
        kind,
        FormControlKind::ListBox {
            input_range: Some("$H$1:$H$5".to_string()),
            cell_link: None,
            selection: ListSelection::Multi,
            selected: vec![0, 2, 4],
            no_3d: false,
        }
    );
}

#[test]
fn dropdown_round_trips() {
    let kind = single_control_round_trip(FormControlKind::Dropdown {
        input_range: Some("$H$1:$H$4".to_string()),
        cell_link: Some("$D$4".to_string()),
        selected: Some(2),
        lines: 6,
        no_3d: true,
    });
    assert_eq!(
        kind,
        FormControlKind::Dropdown {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: Some("$D$4".to_string()),
            selected: Some(2),
            lines: 6,
            no_3d: true,
        }
    );
}

#[test]
fn scrollbar_round_trips() {
    let kind = single_control_round_trip(FormControlKind::Scrollbar {
        value: 40,
        min: 5,
        max: 95,
        increment: 2,
        page: 10,
        horizontal: true,
        cell_link: Some("$D$6".to_string()),
    });
    assert_eq!(
        kind,
        FormControlKind::Scrollbar {
            value: 40,
            min: 5,
            max: 95,
            increment: 2,
            page: 10,
            horizontal: true,
            cell_link: Some("$D$6".to_string()),
        }
    );
}

#[test]
fn spinner_round_trips() {
    let kind = single_control_round_trip(FormControlKind::Spinner {
        value: 12,
        min: 0,
        max: 30,
        increment: 3,
        cell_link: Some("$D$7".to_string()),
    });
    assert_eq!(
        kind,
        FormControlKind::Spinner {
            value: 12,
            min: 0,
            max: 30,
            increment: 3,
            cell_link: Some("$D$7".to_string()),
        }
    );
}

#[test]
fn cross_sheet_cell_link_round_trips() {
    let mut wb = Workbook::new();
    wb.add_worksheet_with_name("Data").expect("second sheet");
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_drawing(control_at(
        FormControlKind::Checkbox {
            caption: "linked".into(),
            state: CheckState::Checked,
            cell_link: Some("Data!$B$2".to_string()),
            no_3d: false,
        },
        anchor(0, 0, 2, 2),
    ));

    let parsed = write_then_read(&wb);
    let controls = controls_of(parsed.worksheet(0).unwrap());
    assert_eq!(controls.len(), 1);
    assert_eq!(controls[0].payload.cell_link(), Some("Data!$B$2"));
}

#[test]
fn unicode_caption_round_trips() {
    let kind = single_control_round_trip(FormControlKind::Checkbox {
        caption: " 自行拖车Shipper arrange ✓".into(),
        state: CheckState::Unchecked,
        cell_link: None,
        no_3d: false,
    });
    match kind {
        FormControlKind::Checkbox { caption, .. } => {
            assert_eq!(caption.plain_text(), " 自行拖车Shipper arrange ✓");
        }
        other => panic!("expected Checkbox, got {other:?}"),
    }
}

#[test]
fn control_anchor_round_trips() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_drawing(control_at(
        FormControlKind::Button {
            caption: "here".into(),
        },
        anchor(2, 3, 5, 8),
    ));

    let parsed = write_then_read(&wb);
    let controls = controls_of(parsed.worksheet(0).unwrap());
    match &controls[0].object.anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            assert_eq!(from.col, 2);
            assert_eq!(from.row, 3);
            assert_eq!(to.col, 5);
            assert_eq!(to.row, 8);
        }
        other => panic!("expected TwoCell anchor, got {other:?}"),
    }
}

#[test]
fn locked_and_printable_flags_round_trip() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let mut object = control_at(
        FormControlKind::Checkbox {
            caption: "flags".into(),
            state: CheckState::Unchecked,
            cell_link: None,
            no_3d: false,
        },
        anchor(0, 0, 2, 2),
    );
    object.meta.locked = false;
    object.meta.printable = false;
    ws.add_drawing(object);

    let parsed = write_then_read(&wb);
    let controls = controls_of(parsed.worksheet(0).unwrap());
    assert!(!controls[0].object.meta.locked);
    assert!(!controls[0].object.meta.printable);
}

#[test]
fn controls_coexist_with_comments_and_pictures() {
    // Ordering test: pictures, comments, and controls share the
    // per-sheet drawing; OBJ↔shape pairing must stay positional.
    const TEST_PNG_1X1: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F,
        0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0B, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x60,
        0x00, 0x02, 0x00, 0x00, 0x05, 0x00, 0x01, 0x7A, 0x5E, 0xAB, 0x3F, 0x00, 0x00, 0x00, 0x00,
        0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).expect("A1");
    ws.add_drawing(
        DrawingObject::image(duke_sheets_chart::EmbeddedImage {
            format: duke_sheets_chart::ImageFormat::Png,
            media_path: String::new(),
            svg_media_path: None,
            width_emu: 1_000_000,
            height_emu: 1_000_000,
            rotation: None,
            flip_h: false,
            flip_v: false,
            data: TEST_PNG_1X1.to_vec(),
            svg_data: None,
        })
        .with_anchor(anchor(6, 1, 8, 4))
        .with_name("Pic"),
    );
    ws.set_comment_at(0, 0, duke_sheets_core::CellComment::new("Author", "note"))
        .expect("set comment");
    ws.add_drawing(control_at(
        FormControlKind::Checkbox {
            caption: "check".into(),
            state: CheckState::Checked,
            cell_link: None,
            no_3d: false,
        },
        anchor(1, 1, 3, 3),
    ));
    ws.add_drawing(control_at(
        FormControlKind::Button {
            caption: "go".into(),
        },
        anchor(1, 5, 3, 7),
    ));

    let parsed = write_then_read(&wb);
    let ws2 = parsed.worksheet(0).unwrap();
    assert_eq!(ws2.image_count(), 1, "picture survives");
    assert_eq!(ws2.comment_count(), 1, "comment survives");
    let controls = controls_of(ws2);
    assert_eq!(controls.len(), 2, "controls survive");
    assert_eq!(controls[0].payload.caption_text().as_deref(), Some("check"));
    assert_eq!(controls[1].payload.caption_text().as_deref(), Some("go"));
    match &controls[0].payload.kind {
        FormControlKind::Checkbox { state, .. } => assert_eq!(*state, CheckState::Checked),
        other => panic!("expected Checkbox, got {other:?}"),
    }
}

#[test]
fn controls_on_multiple_sheets_round_trip() {
    let mut wb = Workbook::new();
    wb.add_worksheet_with_name("Second").expect("sheet 2");
    wb.worksheet_mut(0).unwrap().add_drawing(control_at(
            FormControlKind::Checkbox {
            caption: "one".into(),
                state: CheckState::Checked,
                cell_link: None,
                no_3d: false,
            },
            anchor(0, 0, 2, 2),
        ));
    wb.worksheet_mut(1).unwrap().add_drawing(control_at(
            FormControlKind::Spinner {
                value: 5,
                min: 0,
                max: 10,
                increment: 1,
                cell_link: None,
            },
            anchor(0, 0, 1, 3),
        ));

    let parsed = write_then_read(&wb);
    assert_eq!(parsed.worksheet(0).unwrap().form_control_count(), 1);
    assert_eq!(parsed.worksheet(1).unwrap().form_control_count(), 1);
    assert_eq!(
        controls_of(parsed.worksheet(0).unwrap())[0]
            .payload
            .caption_text()
            .as_deref(),
        Some("one")
    );
    match &controls_of(parsed.worksheet(1).unwrap())[0].payload.kind {
        FormControlKind::Spinner { value, max, .. } => {
            assert_eq!(*value, 5);
            assert_eq!(*max, 10);
        }
        other => panic!("expected Spinner, got {other:?}"),
    }
}

/// Build the multi-group radio workbook used by the grouping tests:
/// two group boxes with two radios each (interleaved insertion
/// order) plus a loose radio outside both boxes.
///
/// Controls in insertion order: Box A, Box B, A1, B1, A2, B2, Loose
/// (obj ids 1-7).
fn radio_groups_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let radio = |caption: &str, state: CheckState| FormControlKind::OptionButton {
        caption: caption.into(),
        state,
        cell_link: None,
        first_in_group: false,
        no_3d: false,
    };
    let group_box = |caption: &str| FormControlKind::GroupBox {
        caption: caption.into(),
        no_3d: false,
    };
    // Box A spans cols 0-2, Box B cols 4-6; radios sit inside them,
    // the loose radio at col 8 is outside both.
    ws.add_drawing(control_at(group_box("Box A"), anchor(0, 0, 2, 6)));
    ws.add_drawing(control_at(group_box("Box B"), anchor(4, 0, 6, 6)));
    ws.add_drawing(control_at(
        radio("A1", CheckState::Checked),
        anchor(1, 1, 2, 2),
    ));
    ws.add_drawing(control_at(
        radio("B1", CheckState::Unchecked),
        anchor(5, 1, 6, 2),
    ));
    ws.add_drawing(control_at(
        radio("A2", CheckState::Unchecked),
        anchor(1, 3, 2, 4),
    ));
    ws.add_drawing(control_at(
        radio("B2", CheckState::Checked),
        anchor(5, 3, 6, 4),
    ));
    ws.add_drawing(control_at(
        radio("Loose", CheckState::Unchecked),
        anchor(8, 1, 9, 2),
    ));
    wb
}

#[test]
fn radio_groups_follow_enclosing_group_boxes() {
    // Each group box forms its own radio group (fFirstBtn on its
    // first radio); radios outside every box form the sheet group.
    let parsed = write_then_read(&radio_groups_workbook());
    let controls = controls_of(parsed.worksheet(0).unwrap());
    assert_eq!(controls.len(), 7);

    let firsts: Vec<(String, bool)> = controls
        .iter()
        .filter_map(|c| match &c.payload.kind {
            FormControlKind::OptionButton {
                caption,
                first_in_group,
                ..
            } => Some((caption.plain_text(), *first_in_group)),
            _ => None,
        })
        .collect();
    assert_eq!(
        firsts,
        vec![
            ("A1".to_string(), true),
            ("B1".to_string(), true),
            ("A2".to_string(), false),
            ("B2".to_string(), false),
            ("Loose".to_string(), true),
        ],
        "each group's first radio carries fFirstBtn"
    );
}

#[test]
fn radio_chains_link_within_groups_only() {
    // Byte-level: the FtRboData idRadNext chains must be circular
    // per group, matching how Excel persists grouping. The model
    // only exposes fFirstBtn, so walk the written OBJ records.
    use duke_sheets_xls::biff::{self, obj};
    use duke_sheets_xls::cfb::CompoundFile;

    let bytes = XlsWriter::write_to_bytes(&radio_groups_workbook()).expect("serialize");
    let cfb = CompoundFile::open(Cursor::new(&bytes)).expect("cfb");
    let stream = cfb.read_stream("/Workbook").expect("workbook stream");
    let records = biff::read_all_records(&mut Cursor::new(stream)).expect("records");

    let mut chains = std::collections::HashMap::new();
    for rec in &records {
        if rec.record_type != biff::records::OBJ {
            continue;
        }
        let parsed = obj::parse_obj(&rec.data).expect("parse obj");
        if let Some((next, first)) = parsed.radio {
            chains.insert(parsed.id, (next, first));
        }
    }

    // Insertion order gives obj ids: BoxA=1, BoxB=2, A1=3, B1=4,
    // A2=5, B2=6, Loose=7.
    assert_eq!(chains[&3], (5, true), "A1 chains to A2 and heads box A");
    assert_eq!(chains[&5], (3, false), "A2 wraps back to A1");
    assert_eq!(chains[&4], (6, true), "B1 chains to B2 and heads box B");
    assert_eq!(chains[&6], (4, false), "B2 wraps back to B1");
    assert_eq!(chains[&7], (7, true), "loose radio is its own group");
}

#[test]
fn empty_caption_round_trips() {
    // Empty captions take the cchText=0 TXO path (header only, no
    // CONTINUE records).
    let kind = single_control_round_trip(FormControlKind::Checkbox {
        caption: String::new().into(),
        state: CheckState::Checked,
        cell_link: None,
        no_3d: false,
    });
    assert_eq!(
        kind,
        FormControlKind::Checkbox {
            caption: String::new().into(),
            state: CheckState::Checked,
            cell_link: None,
            no_3d: false,
        }
    );
}

#[test]
fn oversized_caption_round_trips_via_multiple_continue_records() {
    // A narrow-encoded text CONTINUE holds at most 8,223 chars (8,224
    // byte body minus the encoding grbit); 8,400 chars forces the
    // text to split across two CONTINUE records.
    let long: String = "xy".repeat(4_200);
    let kind = single_control_round_trip(FormControlKind::Label {
        caption: long.clone().into(),
    });
    match kind {
        FormControlKind::Label { caption } => {
            assert_eq!(caption.plain_text(), long);
        }
        other => panic!("expected Label, got {other:?}"),
    }
}

#[test]
fn mixed_caption_plain_run_keeps_font_none() {
    // The plain run of a [plain, styled] caption is written with TXO
    // ifnt=0 (workbook default font) and must read back as
    // font: None, not an explicit copy of the default font.
    let caption = ControlText {
        runs: vec![
            RichTextRun::plain("plain "),
            RichTextRun::with_font(
                "styled",
                RunFont {
                    bold: Some(true),
                    ..RunFont::default()
                },
            ),
        ],
        horizontal_alignment: None,
        vertical_alignment: None,
    };
    let kind = single_control_round_trip(FormControlKind::Label { caption });
    let FormControlKind::Label { caption } = kind else {
        panic!("expected Label, got {kind:?}");
    };
    assert_eq!(caption.plain_text(), "plain styled");
    assert_eq!(caption.runs.len(), 2, "runs: {:?}", caption.runs);
    assert_eq!(
        caption.runs[0].font, None,
        "plain run must stay font-less; got {:?}",
        caption.runs[0].font
    );
    let styled = caption.runs[1].font.as_ref().expect("styled run font");
    assert_eq!(styled.bold, Some(true));
}

#[test]
fn oversized_macro_name_returns_a_clean_error() {
    // Lbl cch is a u8, so a macro name over 255 UTF-16 units cannot
    // be emitted. Silently skipping its Lbl would shift every later
    // macro's PtgName index onto the wrong name; the write must fail
    // instead.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_drawing(
        DrawingObject::form_control(
            FormControl::new(FormControlKind::Button {
                caption: "First".into(),
            })
            .with_macro_name("A".repeat(300)),
        )
        .with_anchor(anchor(1, 1, 3, 3)),
    );
    ws.add_drawing(
        DrawingObject::form_control(
            FormControl::new(FormControlKind::Button {
                caption: "Second".into(),
            })
            .with_macro_name("RunSecond"),
        )
        .with_anchor(anchor(1, 5, 3, 7)),
    );
    let err = XlsWriter::write_to_bytes(&wb).expect_err("oversized macro name must fail");
    assert!(
        err.to_string().contains("procedure name"),
        "unexpected error: {err}"
    );
}

// features: Form-control macro assignment
#[test]
fn multiple_macro_names_round_trip_pointing_at_their_own_lbls() {
    // Sorted Lbl order (Alpha before Zulu) differs from insertion
    // order, so each control's PtgName index must land on its own
    // Lbl, not its positional neighbour.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_drawing(
        DrawingObject::form_control(
            FormControl::new(FormControlKind::Button {
                caption: "Z".into(),
            })
            .with_macro_name("Zulu"),
        )
        .with_anchor(anchor(1, 1, 3, 3)),
    );
    ws.add_drawing(
        DrawingObject::form_control(
            FormControl::new(FormControlKind::Button {
                caption: "A".into(),
            })
            .with_macro_name("Alpha"),
        )
        .with_anchor(anchor(1, 5, 3, 7)),
    );

    let parsed = write_then_read(&wb);
    let controls = controls_of(parsed.worksheet(0).unwrap());
    assert_eq!(controls.len(), 2);
    assert_eq!(controls[0].payload.macro_name.as_deref(), Some("Zulu"));
    assert_eq!(controls[1].payload.macro_name.as_deref(), Some("Alpha"));
}

#[test]
fn caption_over_u16_character_limit_returns_a_clean_error() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0).unwrap().add_drawing(control_at(
            FormControlKind::Label {
                // Each supplementary scalar is two UTF-16 units.
            caption: "😀".repeat(32_768).into(),
            },
            anchor(0, 0, 2, 2),
        ));
    let err = XlsWriter::write_to_bytes(&wb).expect_err("caption exceeds cchText");
    assert!(err.to_string().contains("maximum is 65535"));
}

#[test]
fn list_box_without_input_range_round_trips() {
    // No input range = ObjFmla with cbFmla=0 inside FtLbsData.
    let kind = single_control_round_trip(FormControlKind::ListBox {
        input_range: None,
        cell_link: Some("$D$5".to_string()),
        selection: ListSelection::Single,
        selected: vec![],
        no_3d: false,
    });
    assert_eq!(
        kind,
        FormControlKind::ListBox {
            input_range: None,
            cell_link: Some("$D$5".to_string()),
            selection: ListSelection::Single,
            selected: vec![],
            no_3d: false,
        }
    );
}

#[test]
fn cross_sheet_input_range_round_trips() {
    // Cross-sheet input ranges compile to PtgArea3d.
    let mut wb = Workbook::new();
    wb.add_worksheet_with_name("Data").expect("second sheet");
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_drawing(control_at(
        FormControlKind::Dropdown {
            input_range: Some("Data!$A$1:$A$5".to_string()),
            cell_link: None,
            selected: Some(4),
            lines: 8,
            no_3d: false,
        },
        anchor(0, 0, 2, 1),
    ));

    let parsed = write_then_read(&wb);
    let controls = controls_of(parsed.worksheet(0).unwrap());
    assert_eq!(controls.len(), 1);
    match &controls[0].payload.kind {
        FormControlKind::Dropdown {
            input_range,
            selected,
            ..
        } => {
            assert_eq!(input_range.as_deref(), Some("Data!$A$1:$A$5"));
            assert_eq!(*selected, Some(4));
        }
        other => panic!("expected Dropdown, got {other:?}"),
    }
}

#[test]
fn uncompilable_cell_link_returns_a_clean_error() {
    // A cell link that isn't a single reference cannot be encoded in
    // an ObjectParsedFormula; the writer drops it rather than emit an
    // out-of-spec record.
    let mut wb = Workbook::new();
    wb.worksheet_mut(0).unwrap().add_drawing(control_at(
            FormControlKind::Checkbox {
            caption: "bad link".into(),
                state: CheckState::Checked,
                cell_link: Some("SUM(A1:A3)".to_string()),
                no_3d: false,
            },
            anchor(0, 0, 2, 2),
        ));
    let err = XlsWriter::write_to_bytes(&wb).expect_err("non-reference link is invalid");
    assert!(err.to_string().contains("must be one BIFF8"));
}

#[test]
fn scrollbar_values_above_i16_max_clamp() {
    // FtSbs fields are signed 16-bit; model values above i16::MAX
    // clamp instead of wrapping negative.
    let kind = single_control_round_trip(FormControlKind::Scrollbar {
        value: 40_000,
        min: 0,
        max: 65_535,
        increment: 1,
        page: 10,
        horizontal: false,
        cell_link: None,
    });
    match kind {
        FormControlKind::Scrollbar { value, max, .. } => {
            assert_eq!(value, 32_767, "value clamps to i16::MAX");
            assert_eq!(max, 32_767, "max clamps to i16::MAX");
        }
        other => panic!("expected Scrollbar, got {other:?}"),
    }
}

#[test]
fn large_multi_select_list_uses_continue_records() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.add_drawing(control_at(
        FormControlKind::ListBox {
            input_range: Some("$H$1:$H$10000".to_string()),
            cell_link: None,
            selection: ListSelection::Multi,
            selected: vec![0, 4_999, 9_999],
            no_3d: false,
        },
        anchor(0, 0, 2, 10),
    ));

    let parsed = write_then_read(&wb);
    match &controls_of(parsed.worksheet(0).unwrap())[0].payload.kind {
        FormControlKind::ListBox { selected, .. } => {
            assert_eq!(selected, &vec![0, 4_999, 9_999]);
        }
        other => panic!("expected ListBox, got {other:?}"),
    }
}

#[test]
fn full_column_list_range_returns_a_clean_error() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0).unwrap().add_drawing(control_at(
            FormControlKind::ListBox {
                input_range: Some("$H$1:$H$65536".to_string()),
                cell_link: None,
                selection: ListSelection::Single,
                selected: Vec::new(),
                no_3d: false,
            },
            anchor(0, 0, 2, 10),
        ));

    let err = XlsWriter::write_to_bytes(&wb).expect_err("cLines exceeds the BIFF8 limit");
    assert!(err.to_string().contains("32767"));
}

#[test]
fn list_selection_outside_input_range_returns_a_clean_error() {
    // Zero-based index 4 is the first out-of-range value for 4 items.
    let mut wb = Workbook::new();
    wb.worksheet_mut(0).unwrap().add_drawing(control_at(
            FormControlKind::ListBox {
                input_range: Some("$H$1:$H$4".to_string()),
                cell_link: None,
                selection: ListSelection::Single,
                selected: vec![4],
                no_3d: false,
            },
            anchor(0, 0, 2, 4),
        ));

    let err = XlsWriter::write_to_bytes(&wb).expect_err("selection exceeds cLines");
    assert!(err.to_string().contains("selection index 4"));
}

#[test]
fn control_anchor_outside_biff8_grid_returns_a_clean_error() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0).unwrap().add_drawing(control_at(
            FormControlKind::Button {
            caption: "outside".into(),
            },
            anchor(256, 0, 257, 1),
        ));
    let err = XlsWriter::write_to_bytes(&wb).expect_err("column 256 is outside XLS");
    assert!(err.to_string().contains("BIFF8 sheet grid"));
}

#[test]
fn reversed_control_anchor_returns_a_clean_error() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0).unwrap().add_drawing(control_at(
            FormControlKind::Button {
            caption: "reversed".into(),
            },
            anchor(4, 4, 2, 2),
        ));
    let err = XlsWriter::write_to_bytes(&wb).expect_err("reversed anchor is invalid");
    assert!(err.to_string().contains("endpoints are reversed"));
}

#[test]
fn out_of_grid_control_formula_returns_a_clean_error() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0).unwrap().add_drawing(control_at(
            FormControlKind::Checkbox {
            caption: "link".into(),
                state: CheckState::Checked,
                cell_link: Some("$IW$1".to_string()),
                no_3d: false,
            },
            anchor(0, 0, 2, 2),
        ));
    let err = XlsWriter::write_to_bytes(&wb).expect_err("column IW is outside XLS");
    assert!(err.to_string().contains("must be one BIFF8"));
}

#[test]
fn invalid_scrollbar_tuple_returns_a_clean_error() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0).unwrap().add_drawing(control_at(
            FormControlKind::Scrollbar {
                value: 50,
                min: 100,
                max: 10,
                increment: 1,
                page: 10,
                horizontal: false,
                cell_link: None,
            },
            anchor(0, 0, 1, 5),
        ));
    let err = XlsWriter::write_to_bytes(&wb).expect_err("invalid scrollbar bounds");
    assert!(err.to_string().contains("min <= value <= max"));
}

#[test]
fn mixed_option_button_returns_a_clean_error() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0).unwrap().add_drawing(control_at(
            FormControlKind::OptionButton {
            caption: "mixed".into(),
                state: CheckState::Mixed,
                cell_link: None,
                first_in_group: true,
                no_3d: false,
            },
            anchor(0, 0, 2, 2),
        ));
    let err = XlsWriter::write_to_bytes(&wb).expect_err("Mixed is checkbox-only");
    assert!(
        err.to_string()
            .contains("option buttons cannot use the mixed state"),
        "unexpected error: {err}"
    );
}

/// LibreOffice envelope check: write an XLS carrying one of every
/// control kind, open via URP, read the anchor cell's value back.
///
/// Smoke test — if the OBJ subrecords or the Escher tree are
/// malformed, LO refuses to open the file and `open_workbook`
/// errors. Deeper control inspection via UNO is fragile across LO
/// versions, so this settles for confirming the file loads.
#[test]
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
fn lo_can_open_xls_with_form_controls_we_emit() {
    duke_sheets_test_harness::lo::ensure_lo();

    const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).expect("A1");
    let kinds: Vec<FormControlKind> = vec![
        FormControlKind::Button {
            caption: "Run".into(),
        },
        FormControlKind::Checkbox {
            caption: "Check".into(),
            state: CheckState::Checked,
            cell_link: Some("$D$2".to_string()),
            no_3d: false,
        },
        FormControlKind::OptionButton {
            caption: "Opt".into(),
            state: CheckState::Checked,
            cell_link: None,
            first_in_group: true,
            no_3d: false,
        },
        FormControlKind::Label {
            caption: "Info".into(),
        },
        FormControlKind::GroupBox {
            caption: "Frame".into(),
            no_3d: true,
        },
        FormControlKind::ListBox {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: None,
            selection: ListSelection::Single,
            selected: vec![2],
            no_3d: true,
        },
        FormControlKind::Dropdown {
            input_range: Some("$H$1:$H$4".to_string()),
            cell_link: None,
            selected: Some(1),
            lines: 8,
            no_3d: true,
        },
        FormControlKind::Scrollbar {
            value: 40,
            min: 0,
            max: 100,
            increment: 1,
            page: 10,
            horizontal: false,
            cell_link: None,
        },
        FormControlKind::Spinner {
            value: 3,
            min: 0,
            max: 10,
            increment: 1,
            cell_link: None,
        },
    ];
    for (i, kind) in kinds.into_iter().enumerate() {
        let row = 1 + 2 * i as u32;
        ws.add_drawing(control_at(kind, anchor(1, row, 3, row + 1)));
    }
    assert_eq!(wb.sync_form_control_links(), 1);

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    std::fs::create_dir_all(SHARED_DIR).expect("shared dir");
    let pid = std::process::id();
    let path = format!("{SHARED_DIR}/duke_form_controls_{pid}.xls");
    std::fs::write(&path, &bytes).expect("write");

    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let outcome: Result<f64, String> = rt.block_on(async {
        let mut bridge =
            duke_sheets_libreoffice::bridge::LibreOfficeBridge::connect("127.0.0.1", 2002)
                .await
                .map_err(|e| format!("connect: {e}"))?;
        let mut wb_in = bridge
            .open_workbook(&path)
            .await
            .map_err(|e| format!("open: {e}"))?;
        let a1 = wb_in
            .get_cell_value("A1")
            .await
            .map_err(|e| format!("A1: {e}"))?;
        Ok(a1)
    });
    let _ = std::fs::remove_file(&path);
    let a1 = outcome.expect("LO must open our XLS with form controls without error");
    assert!(
        (a1 - 42.0).abs() < 1e-9,
        "A1 must round-trip; got {a1} (expected 42)"
    );
}

#[test]
fn empty_workbook_emits_no_drawing_records() {
    let mut wb = Workbook::new();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", "value")
        .expect("A1");

    let bytes = XlsWriter::write_to_bytes(&wb).expect("serialize");
    let parsed = XlsReader::read(Cursor::new(&bytes)).expect("read");
    assert_eq!(parsed.worksheet(0).unwrap().form_control_count(), 0);
}
