use std::io::{Cursor, Read, Write};

use duke_sheets::{
    column_width_to_emu, radio_groups, row_height_to_emu, CellMarker, CheckState,
    DrawingAnchor, DrawingMetrics, DrawingObject, FormControl, FormControlKind, Workbook,
    Worksheet, XlsReader, XlsWriter, XlsbReader, XlsbWriter, XlsxReader, XlsxWriter,
};

fn unknown_edit_box() -> FormControlKind {
    FormControlKind::Unknown {
        object_type: "EditBox".to_string(),
        legacy_object_type: None,
        caption: "Unsupported editor".into(),
        raw_properties: vec![
            ("customFlag".to_string(), "kept".to_string()),
            ("val".to_string(), "17".to_string()),
            ("fmlaLink".to_string(), "$A$1".to_string()),
        ],
        raw_client_data: vec![
            b"<x:CustomState mode=\"alpha\"/>".to_vec(),
            b"<x:Val>17</x:Val>".to_vec(),
            b"<x:FmlaLink>$A$1</x:FmlaLink>".to_vec(),
        ],
        raw_obj: None,
    }
}

fn unknown_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    let control = FormControl::new(unknown_edit_box()).with_macro_name("RunUnknown");
    let mut object = DrawingObject::form_control(control).with_anchor(DrawingAnchor::TwoCell {
        from: CellMarker {
            col: 1,
            col_offset_emu: 95_250,
            row: 2,
            row_offset_emu: 47_625,
        },
        to: CellMarker {
            col: 3,
            col_offset_emu: 190_500,
            row: 4,
            row_offset_emu: 95_250,
        },
        edit_as: None,
    });
    object.meta.name = Some("Legacy editor".to_string());
    workbook.worksheet_mut(0).unwrap().add_drawing(object);
    workbook
}

fn assert_unknown_edit_box(workbook: &Workbook, expect_ctrl_props: bool) {
    let drawn = workbook
        .worksheet(0)
        .unwrap()
        .form_controls()
        .next()
        .expect("unknown control survives");
    assert_eq!(drawn.object.meta.name.as_deref(), Some("Legacy editor"));
    assert_eq!(drawn.payload.macro_name.as_deref(), Some("RunUnknown"));
    assert_eq!(drawn.payload.caption_text().as_deref(), Some("Unsupported editor"));
    assert_eq!(
        drawn.object.anchor,
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 1,
                col_offset_emu: 95_250,
                row: 2,
                row_offset_emu: 47_625,
            },
            to: CellMarker {
                col: 3,
                col_offset_emu: 190_500,
                row: 4,
                row_offset_emu: 95_250,
            },
            edit_as: None,
        }
    );
    let FormControlKind::Unknown {
        object_type,
        legacy_object_type,
        raw_properties,
        raw_client_data,
        raw_obj,
        ..
    } = &drawn.payload.kind
    else {
        panic!("expected unknown control, got {:?}", drawn.payload.kind);
    };
    assert_eq!(object_type, "EditBox");
    assert_eq!(*legacy_object_type, None);
    assert_eq!(raw_obj, &None);
    if expect_ctrl_props {
        assert!(raw_properties.contains(&("customFlag".to_string(), "kept".to_string())));
        assert!(raw_properties.contains(&("val".to_string(), "17".to_string())));
        assert!(raw_properties.contains(&("fmlaLink".to_string(), "$A$1".to_string())));
    }
    assert!(raw_client_data.iter().any(|xml| {
        let xml = String::from_utf8_lossy(xml);
        xml.contains("CustomState") && xml.contains("alpha")
    }));
    assert!(raw_client_data
        .iter()
        .any(|xml| String::from_utf8_lossy(xml).contains("<x:Val>17</x:Val>")));
    assert!(raw_client_data
        .iter()
        .any(|xml| String::from_utf8_lossy(xml).contains("<x:FmlaLink>$A$1</x:FmlaLink>")));
}

#[test]
fn unknown_edit_box_round_trips_xlsx_and_xlsb() {
    let mut workbook = unknown_workbook();
    assert_eq!(workbook.sync_form_control_links(), 0);

    let mut xlsx = Cursor::new(Vec::new());
    XlsxWriter::write(&workbook, &mut xlsx).expect("write xlsx");
    let xlsx = XlsxReader::read(Cursor::new(xlsx.into_inner())).expect("read xlsx");
    assert_unknown_edit_box(&xlsx, true);

    let mut xlsb = Cursor::new(Vec::new());
    XlsbWriter::write(&workbook, &mut xlsb).expect("write xlsb");
    let xlsb = XlsbReader::read(Cursor::new(xlsb.into_inner())).expect("read xlsb");
    assert_unknown_edit_box(&xlsb, false);
}

fn raw_edit_box_obj() -> Vec<u8> {
    let mut body = Vec::new();
    body.extend_from_slice(&0x0015u16.to_le_bytes());
    body.extend_from_slice(&0x0012u16.to_le_bytes());
    body.extend_from_slice(&0x000Du16.to_le_bytes());
    body.extend_from_slice(&42u16.to_le_bytes());
    body.extend_from_slice(&0x0011u16.to_le_bytes());
    body.extend_from_slice(&[0; 12]);
    body.extend_from_slice(&0x0010u16.to_le_bytes());
    body.extend_from_slice(&8u16.to_le_bytes());
    body.extend_from_slice(&[1, 2, 3, 4, 5, 6, 7, 8]);
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes());
    body
}

fn assert_xls_unknown(workbook: &Workbook, original_raw: &[u8]) {
    let drawn = workbook
        .worksheet(0)
        .unwrap()
        .form_controls()
        .next()
        .expect("XLS unknown control survives");
    assert_eq!(drawn.object.meta.name.as_deref(), Some("Edit field"));
    assert!(drawn.object.meta.hidden);
    assert!(!drawn.object.meta.locked);
    assert!(!drawn.object.meta.printable);
    assert_eq!(drawn.payload.caption_text().as_deref(), Some("Raw edit box"));
    let FormControlKind::Unknown {
        object_type,
        legacy_object_type,
        raw_obj: Some(raw_obj),
        ..
    } = &drawn.payload.kind
    else {
        panic!("expected raw-backed XLS unknown control");
    };
    assert_eq!(object_type, "EditBox");
    assert_eq!(*legacy_object_type, Some(0x000D));
    assert_eq!(&raw_obj[10..], &original_raw[10..]);
}

#[test]
fn xls_unknown_edit_box_obj_body_survives_rewrite() {
    let raw_obj = raw_edit_box_obj();
    let control = FormControl::new(FormControlKind::Unknown {
        object_type: "EditBox".to_string(),
        legacy_object_type: Some(0x000D),
        caption: "Raw edit box".into(),
        raw_properties: Vec::new(),
        raw_client_data: Vec::new(),
        raw_obj: Some(raw_obj.clone()),
    });
    let mut object = DrawingObject::form_control(control).with_anchor(DrawingAnchor::TwoCell {
        from: CellMarker {
            col: 0,
            col_offset_emu: 0,
            row: 0,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: 2,
            col_offset_emu: 0,
            row: 2,
            row_offset_emu: 0,
        },
        edit_as: None,
    });
    object.meta.name = Some("Edit field".to_string());
    object.meta.hidden = true;
    object.meta.locked = false;
    object.meta.printable = false;

    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_drawing(object);
    let first_bytes = XlsWriter::write_to_bytes(&workbook).expect("write first xls");
    let first = XlsReader::read(Cursor::new(first_bytes)).expect("read first xls");
    assert_xls_unknown(&first, &raw_obj);

    let second_bytes = XlsWriter::write_to_bytes(&first).expect("rewrite xls");
    let second = XlsReader::read(Cursor::new(second_bytes)).expect("read rewritten xls");
    assert_xls_unknown(&second, &raw_obj);

    let error = XlsWriter::write_to_bytes(&unknown_workbook()).unwrap_err();
    assert!(
        error.to_string().contains("raw OBJ body"),
        "new XLS unknown controls must fail clearly: {error}"
    );
}

fn patch_vml_object_type(bytes: Vec<u8>, replacement: &str) -> Vec<u8> {
    let mut input = zip::ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..input.len() {
        let mut file = input.by_index(index).expect("zip entry");
        let name = file.name().to_string();
        let is_dir = file.is_dir();
        let mut data = Vec::new();
        file.read_to_end(&mut data).expect("read zip entry");
        drop(file);

        if is_dir {
            output
                .add_directory(name, zip::write::SimpleFileOptions::default())
                .expect("copy directory");
            continue;
        }
        if name.contains("vmlDrawing") {
            let xml = String::from_utf8(data).expect("VML is UTF-8");
            data = xml
                .replace("ObjectType=\"EditBox\"", &format!("ObjectType=\"{replacement}\""))
                .into_bytes();
        }
        output
            .start_file(name, zip::write::SimpleFileOptions::default())
            .expect("copy file");
        output.write_all(&data).expect("write zip entry");
    }
    output.finish().expect("finish zip").into_inner()
}

#[test]
fn pict_vml_is_not_exposed_as_unknown_form_control() {
    let mut bytes = Cursor::new(Vec::new());
    XlsbWriter::write(&unknown_workbook(), &mut bytes).expect("write xlsb");
    let patched = patch_vml_object_type(bytes.into_inner(), "Pict");
    let reopened = XlsbReader::read(Cursor::new(patched)).expect("read patched xlsb");
    assert_eq!(reopened.worksheet(0).unwrap().form_control_count(), 0);

    let pict = FormControl::new(FormControlKind::Unknown {
        object_type: "Pict".to_string(),
        legacy_object_type: None,
        caption: "ActiveX placeholder".into(),
        raw_properties: Vec::new(),
        raw_client_data: Vec::new(),
        raw_obj: None,
    });
    assert!(pict.validate().is_err());
    let mut pict_workbook = Workbook::new();
    pict_workbook
        .worksheet_mut(0)
        .unwrap()
        .add_form_control(pict, DrawingAnchor::default());
    let mut output = Cursor::new(Vec::new());
    let error = XlsbWriter::write(&pict_workbook, &mut output).unwrap_err();
    assert!(error.to_string().contains("ActiveX/OLE"));

    let blank = FormControl::new(FormControlKind::Unknown {
        object_type: "  ".to_string(),
        legacy_object_type: None,
        caption: "blank type".into(),
        raw_properties: Vec::new(),
        raw_client_data: Vec::new(),
        raw_obj: None,
    });
    assert!(blank.validate().is_err());
}

fn metric_anchor_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_column_width(0, 20.0);
    sheet.set_row_height(0, 30.0);
    sheet.add_form_control(
        FormControl::new(FormControlKind::Checkbox {
            caption: "one-cell".into(),
            state: CheckState::Unchecked,
            cell_link: None,
            no_3d: false,
        }),
        DrawingAnchor::OneCell {
            from: CellMarker::default(),
            width_emu: 609_600,
            height_emu: 190_500,
        },
    );
    sheet.add_form_control(
        FormControl::new(FormControlKind::Label {
            caption: "absolute".into(),
        }),
        DrawingAnchor::Absolute {
            x_emu: 609_600,
            y_emu: 190_500,
            width_emu: 95_250,
            height_emu: 95_250,
        },
    );
    workbook
}

fn assert_metric_anchors(workbook: &Workbook) {
    let controls: Vec<_> = workbook.worksheet(0).unwrap().form_controls().collect();
    assert_eq!(controls.len(), 2);
    match &controls[0].object.anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            assert_eq!((from.col, from.col_offset_emu), (0, 0));
            assert_eq!((from.row, from.row_offset_emu), (0, 0));
            assert_eq!((to.col, to.col_offset_emu), (0, 609_600));
            assert_eq!((to.row, to.row_offset_emu), (0, 190_500));
        }
        other => panic!("expected flattened one-cell control anchor, got {other:?}"),
    }
    match &controls[1].object.anchor {
        DrawingAnchor::TwoCell { from, .. } => {
            assert_eq!((from.col, from.col_offset_emu), (0, 609_600));
            assert_eq!((from.row, from.row_offset_emu), (0, 190_500));
        }
        other => panic!("expected flattened absolute control anchor, got {other:?}"),
    }
}

#[test]
fn custom_dimensions_drive_xlsx_and_xlsb_control_anchor_flattening() {
    let workbook = metric_anchor_workbook();

    let mut xlsx = Cursor::new(Vec::new());
    XlsxWriter::write(&workbook, &mut xlsx).expect("write xlsx");
    let xlsx = XlsxReader::read(Cursor::new(xlsx.into_inner())).expect("read xlsx");
    assert_metric_anchors(&xlsx);

    let mut xlsb = Cursor::new(Vec::new());
    XlsbWriter::write(&workbook, &mut xlsb).expect("write xlsb");
    let xlsb = XlsbReader::read(Cursor::new(xlsb.into_inner())).expect("read xlsb");
    assert_metric_anchors(&xlsb);
}

#[test]
fn drawing_metric_helpers_match_excel_compatible_units() {
    assert_eq!(column_width_to_emu(8.43), 609_600);
    assert_eq!(row_height_to_emu(15.0), 190_500);
    assert_eq!(column_width_to_emu(20.0), 1_381_125);
    assert_eq!(row_height_to_emu(30.0), 381_000);

    let mut sheet = Worksheet::new("Metrics");
    sheet.set_column_width(0, 20.0);
    sheet.set_row_height(0, 30.0);
    assert_eq!(sheet.column_width_emu(0), 1_381_125);
    assert_eq!(sheet.row_height_emu(0), 381_000);
}

fn radio(caption: &str, anchor: DrawingAnchor) -> DrawingObject {
    DrawingObject::form_control(FormControl::new(FormControlKind::OptionButton {
        caption: caption.into(),
        state: CheckState::Unchecked,
        cell_link: None,
        first_in_group: false,
        no_3d: false,
    }))
    .with_anchor(anchor)
}

#[test]
fn radio_grouping_uses_custom_sheet_dimensions() {
    let mut sheet = Worksheet::new("Groups");
    sheet.set_column_width(0, 20.0);
    sheet.set_row_height(0, 30.0);
    sheet.add_drawing(
        DrawingObject::form_control(FormControl::new(FormControlKind::GroupBox {
            caption: "Box".into(),
            no_3d: false,
        }))
        .with_anchor(DrawingAnchor::Absolute {
            x_emu: 0,
            y_emu: 0,
            width_emu: 1_000_000,
            height_emu: 1_000_000,
        }),
    );
    sheet.add_drawing(radio(
        "outside with custom width",
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 1,
                col_offset_emu: 0,
                row: 0,
                row_offset_emu: 100_000,
            },
            to: CellMarker {
                col: 1,
                col_offset_emu: 100_000,
                row: 0,
                row_offset_emu: 200_000,
            },
            edit_as: None,
        },
    ));
    sheet.add_drawing(radio(
        "inside",
        DrawingAnchor::Absolute {
            x_emu: 100_000,
            y_emu: 100_000,
            width_emu: 100_000,
            height_emu: 100_000,
        },
    ));

    let placed = sheet.placed_form_controls();
    assert_eq!(placed[1].rect_emu.0, 1_381_125);
    assert_eq!(radio_groups(&placed), vec![vec![1], vec![2]]);
}
