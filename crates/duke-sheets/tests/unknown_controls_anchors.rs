use std::io::{Cursor, Read, Write};

use duke_sheets::{
    column_width_to_emu, radio_groups, row_height_to_emu, CellMarker, CheckState,
    DrawingAnchor, DrawingMetrics, DrawingObject, FormControl, FormControlKind, Workbook,
    Worksheet, XlsReader, XlsWriter, XlsbReader, XlsbWriter, XlsxReader, XlsxWriter,
};

fn unknown_edit_box() -> FormControl {
    let mut control = FormControl::new(FormControlKind::Unknown {
        object_type: "EditBox".to_string(),
        legacy_object_type: None,
        caption: "Unsupported editor".into(),
    });
    control.raw_properties = vec![
        ("customFlag".to_string(), "kept".to_string()),
        ("val".to_string(), "17".to_string()),
        ("fmlaLink".to_string(), "$A$1".to_string()),
    ];
    control
}

fn unknown_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    let mut control = unknown_edit_box().with_macro_name("RunUnknown");
    control.raw_client_data = vec![
        b"<x:CustomState mode=\"alpha\"/>".to_vec(),
        b"<x:Val>17</x:Val>".to_vec(),
        b"<x:FmlaLink>$A$1</x:FmlaLink>".to_vec(),
    ];
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
    workbook.worksheet_mut(0).unwrap().add_drawing(object).unwrap();
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
    let raw_client_data = &drawn.payload.raw_client_data;
    let raw_properties = &drawn.payload.raw_properties;
    let FormControlKind::Unknown {
        object_type,
        legacy_object_type,
        ..
    } = &drawn.payload.kind
    else {
        panic!("expected unknown control, got {:?}", drawn.payload.kind);
    };
    assert_eq!(object_type, "EditBox");
    assert_eq!(*legacy_object_type, None);
    assert_eq!(drawn.payload.raw_obj, None);
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

// features: Unknown legacy Forms controls
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
        ..
    } = &drawn.payload.kind
    else {
        panic!("expected raw-backed XLS unknown control");
    };
    assert_eq!(object_type, "EditBox");
    assert_eq!(*legacy_object_type, Some(0x000D));
    let raw_obj = drawn.payload.raw_obj.as_deref().expect("raw OBJ body captured");
    assert_eq!(&raw_obj[10..], &original_raw[10..]);
}

// features: Unknown legacy Forms controls
#[test]
fn xls_unknown_edit_box_obj_body_survives_rewrite() {
    let raw_obj = raw_edit_box_obj();
    let mut control = FormControl::new(FormControlKind::Unknown {
        object_type: "EditBox".to_string(),
        legacy_object_type: Some(0x000D),
        caption: "Raw edit box".into(),
    });
    control.raw_obj = Some(raw_obj.clone());
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
    workbook.worksheet_mut(0).unwrap().add_drawing(object).unwrap();
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

/// Count ftMacro (0x0004) subrecords in an OBJ record body by
/// walking the ft/cb framing after the 22-byte ftCmo.
fn count_macro_subrecords(body: &[u8]) -> usize {
    let mut count = 0usize;
    let mut pos = 22usize;
    while pos + 4 <= body.len() {
        let ft = u16::from_le_bytes([body[pos], body[pos + 1]]);
        let cb = u16::from_le_bytes([body[pos + 2], body[pos + 3]]) as usize;
        if ft == 0x0004 {
            count += 1;
        }
        pos += 4 + cb;
    }
    count
}

fn raw_edit_box_obj_with_macro() -> Vec<u8> {
    let mut body = Vec::new();
    // ftCmo ot=EditBox id=42 grbit=0x0011.
    body.extend_from_slice(&0x0015u16.to_le_bytes());
    body.extend_from_slice(&0x0012u16.to_le_bytes());
    body.extend_from_slice(&0x000Du16.to_le_bytes());
    body.extend_from_slice(&42u16.to_le_bytes());
    body.extend_from_slice(&0x0011u16.to_le_bytes());
    body.extend_from_slice(&[0; 12]);
    // Embedded ftMacro: ObjFmla whose PtgName points at the SOURCE
    // workbook's Lbl #7, a stale index in any rewritten file.
    body.extend_from_slice(&0x0004u16.to_le_bytes());
    body.extend_from_slice(&12u16.to_le_bytes()); // cbFmla
    body.extend_from_slice(&5u16.to_le_bytes()); // cce
    body.extend_from_slice(&[0; 4]); // unused
    body.extend_from_slice(&[0x23, 0x07, 0x00, 0x00, 0x00]); // PtgName 7
    body.push(0); // pad to even
    // FtEdoData payload (opaque here).
    body.extend_from_slice(&0x0010u16.to_le_bytes());
    body.extend_from_slice(&8u16.to_le_bytes());
    body.extend_from_slice(&[1, 2, 3, 4, 5, 6, 7, 8]);
    // ftEnd.
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes());
    body
}

// features: Unknown legacy Forms controls
// features: Unknown legacy Forms controls
#[test]
fn xls_unknown_control_embedded_macro_is_replaced_not_replayed() {
    let mut control = FormControl::new(FormControlKind::Unknown {
        object_type: "EditBox".to_string(),
        legacy_object_type: Some(0x000D),
        caption: "Macro edit box".into(),
    })
    .with_macro_name("RunEdit");
    control.raw_obj = Some(raw_edit_box_obj_with_macro());
    let object = DrawingObject::form_control(control).with_anchor(DrawingAnchor::TwoCell {
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
    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_drawing(object).unwrap();

    let bytes = XlsWriter::write_to_bytes(&workbook).expect("write xls");
    let reread = XlsReader::read(Cursor::new(bytes)).expect("read xls");
    let drawn = reread
        .worksheet(0)
        .unwrap()
        .form_controls()
        .next()
        .expect("unknown control survives");
    assert_eq!(
        drawn.payload.macro_name.as_deref(),
        Some("RunEdit"),
        "macro must reference this file's Lbl, not the source workbook's stale index"
    );
    assert!(matches!(
        drawn.payload.kind,
        FormControlKind::Unknown { .. }
    ));
    let raw_obj = drawn.payload.raw_obj.as_deref().expect("raw OBJ body captured");
    assert_eq!(
        count_macro_subrecords(raw_obj),
        1,
        "exactly one ftMacro subrecord in the emitted OBJ"
    );
    let edo_payload: &[u8] = &[1, 2, 3, 4, 5, 6, 7, 8];
    assert!(
        raw_obj
            .windows(edo_payload.len())
            .any(|window| window == edo_payload),
        "FtEdoData payload replays untouched"
    );
}

/// Old-Excel autofilter dropdown button: `x:UIObj` + `PrintObject=False`,
/// no `x:LCT` (ECMA-376 Part 4 §14.4.2.62). Modern Excel emits no legacy
/// VML for these; the shape survives only in files from older producers.
const AUX_UI_DROPDOWN_VML: &str = r##" <v:shape id="_x0000_s2001" type="#_x0000_t201" style="position:absolute">
  <x:ClientData ObjectType="Drop">
   <x:SizeWithCells/>
   <x:Anchor>0, 0, 0, 0, 1, 16, 1, 0</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:AutoLine>False</x:AutoLine>
   <x:DropStyle>Combo</x:DropStyle>
   <x:DropLines>8</x:DropLines>
   <x:Sel>0</x:Sel>
   <x:NoThreeD/>
   <x:PrintObject>False</x:PrintObject>
   <x:UIObj/>
  </x:ClientData>
 </v:shape>
"##;

/// Insert an extra `<v:shape>` before `</xml>` in every vmlDrawing part.
fn splice_vml_shape(bytes: Vec<u8>, extra_shape: &str) -> Vec<u8> {
    let mut input = zip::ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let mut spliced = false;
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
            assert!(xml.contains("</xml>"), "VML part must close with </xml>");
            data = xml
                .replace("</xml>", &format!("{extra_shape}</xml>"))
                .into_bytes();
            spliced = true;
        }
        output
            .start_file(name, zip::write::SimpleFileOptions::default())
            .expect("copy file");
        output.write_all(&data).expect("write zip entry");
    }
    assert!(spliced, "no vmlDrawing part found to splice into");
    output.finish().expect("finish zip").into_inner()
}

fn checkbox_and_comment_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_form_control(
        FormControl::new(FormControlKind::Checkbox {
            caption: "Real control".into(),
            state: CheckState::Checked,
            cell_link: None,
            no_3d: true,
        }),
        DrawingAnchor::TwoCell {
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
        },
    ).unwrap();
    sheet
        .set_comment_at(
            4,
            4,
            duke_sheets::CellComment::new("Reviewer", "still here"),
        )
        .unwrap();
    workbook
}

fn assert_aux_ui_shape_skipped(workbook: &Workbook) {
    let sheet = workbook.worksheet(0).unwrap();
    assert_eq!(
        sheet.form_control_count(),
        1,
        "UIObj auxiliary shape must not surface as a control"
    );
    let control = sheet.form_controls().next().unwrap();
    assert!(matches!(
        control.payload.kind,
        FormControlKind::Checkbox { .. }
    ));
    let comment = sheet.comment_at(4, 4).expect("comment survives");
    assert_eq!(comment.text, "still here");
}

// features: Form control: dropdown (combo box)
#[test]
fn excel_uiobj_aux_shape_is_not_a_control_xlsx() {
    let workbook = checkbox_and_comment_workbook();
    let mut bytes = Cursor::new(Vec::new());
    XlsxWriter::write(&workbook, &mut bytes).expect("write xlsx");
    let spliced = splice_vml_shape(bytes.into_inner(), AUX_UI_DROPDOWN_VML);
    let reopened = XlsxReader::read(Cursor::new(spliced)).expect("read spliced xlsx");
    assert_aux_ui_shape_skipped(&reopened);
}

// features: Form control: dropdown (combo box)
#[test]
fn excel_uiobj_aux_shape_is_not_a_control_xlsb() {
    let workbook = checkbox_and_comment_workbook();
    let mut bytes = Cursor::new(Vec::new());
    XlsbWriter::write(&workbook, &mut bytes).expect("write xlsb");
    let spliced = splice_vml_shape(bytes.into_inner(), AUX_UI_DROPDOWN_VML);
    let reopened = XlsbReader::read(Cursor::new(spliced)).expect("read spliced xlsb");
    assert_aux_ui_shape_skipped(&reopened);
}

/// Insert extra children into the ClientData of the shape whose
/// `ObjectType` matches, right before its `</x:ClientData>`.
fn splice_into_client_data(bytes: Vec<u8>, object_type: &str, insert: &str) -> Vec<u8> {
    let mut input = zip::ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let mut spliced = false;
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
            let marker = format!("ObjectType=\"{object_type}\"");
            data = if let Some(type_at) = xml.find(&marker) {
                let close_at = xml[type_at..]
                    .find("</x:ClientData>")
                    .map(|offset| type_at + offset)
                    .expect("ClientData close");
                let mut patched = String::with_capacity(xml.len() + insert.len());
                patched.push_str(&xml[..close_at]);
                patched.push_str(insert);
                patched.push_str(&xml[close_at..]);
                spliced = true;
                patched.into_bytes()
            } else {
                xml.into_bytes()
            };
        }
        output
            .start_file(name, zip::write::SimpleFileOptions::default())
            .expect("copy file");
        output.write_all(&data).expect("write zip entry");
    }
    assert!(spliced, "no vmlDrawing part found to splice into");
    output.finish().expect("finish zip").into_inner()
}

/// Read the concatenated vmlDrawing parts of an OOXML zip.
fn vml_parts(bytes: &[u8]) -> String {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes.to_vec())).expect("open zip");
    let mut out = String::new();
    for index in 0..archive.len() {
        let mut file = archive.by_index(index).expect("zip entry");
        if file.name().contains("vmlDrawing") {
            let mut s = String::new();
            file.read_to_string(&mut s).expect("read vml");
            out.push_str(&s);
        }
    }
    out
}

/// Insert extra attributes into every ctrlProps `formControlPr`
/// element of an XLSX zip.
fn splice_ctrl_prop_attrs(bytes: Vec<u8>, insert: &str) -> Vec<u8> {
    let mut input = zip::ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let mut spliced = false;
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
        if name.contains("ctrlProps") {
            let xml = String::from_utf8(data).expect("ctrlProps is UTF-8");
            data = xml
                .replace("<formControlPr ", &format!("<formControlPr {insert} "))
                .into_bytes();
            spliced = true;
        }
        output
            .start_file(name, zip::write::SimpleFileOptions::default())
            .expect("copy file");
        output.write_all(&data).expect("write zip entry");
    }
    assert!(spliced, "no ctrlProps part found to splice into");
    output.finish().expect("finish zip").into_inner()
}

// features: Form control unmodeled ctrlProps passthrough
#[test]
fn unmodeled_ctrl_props_attributes_round_trip_on_modeled_kinds_xlsx() {
    let mut workbook = Workbook::new();
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_form_control(
            FormControl::new(FormControlKind::Scrollbar {
                value: 50,
                min: 5,
                max: 100,
                increment: 2,
                page: 10,
                horizontal: false,
                cell_link: None,
            }),
            DrawingAnchor::default(),
        )
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    XlsxWriter::write(&workbook, &mut bytes).expect("write xlsx");
    let spliced = splice_ctrl_prop_attrs(bytes.into_inner(), "customFlag=\"kept\"");

    let reopened = XlsxReader::read(Cursor::new(spliced)).expect("read spliced xlsx");
    {
        let sheet = reopened.worksheet(0).unwrap();
        let drawn = sheet.form_controls().next().expect("control survives");
        let FormControlKind::Scrollbar { min, .. } = &drawn.payload.kind else {
            panic!("expected the scrollbar to stay modeled");
        };
        assert_eq!(*min, 5);
        assert_eq!(
            drawn.payload.raw_properties,
            vec![("customFlag".to_string(), "kept".to_string())],
            "unmodeled attribute captured without duplicating modeled ones"
        );
    }

    // Editing the model must win over stale raw copies: min drops to
    // its (omitted-at-write) default while the raw attribute stays.
    let mut edited = reopened;
    {
        let sheet = edited.worksheet_mut(0).unwrap();
        let control = sheet.form_control_at_path_mut(&[0]).expect("control");
        let FormControlKind::Scrollbar { min, .. } = &mut control.kind else {
            panic!("expected scrollbar");
        };
        *min = 0;
    }
    let mut rewritten = Cursor::new(Vec::new());
    XlsxWriter::write(&edited, &mut rewritten).expect("rewrite xlsx");
    let reread = XlsxReader::read(Cursor::new(rewritten.into_inner())).expect("reread xlsx");
    let sheet = reread.worksheet(0).unwrap();
    let drawn = sheet.form_controls().next().expect("control survives");
    let FormControlKind::Scrollbar { min, .. } = &drawn.payload.kind else {
        panic!("expected scrollbar");
    };
    assert_eq!(*min, 0, "model edit must not be shadowed by a stale raw attr");
    assert_eq!(
        drawn.payload.raw_properties,
        vec![("customFlag".to_string(), "kept".to_string())]
    );
}

fn control_raws(workbook: &Workbook) -> Vec<String> {
    workbook
        .worksheet(0)
        .unwrap()
        .form_controls()
        .next()
        .expect("control present")
        .payload
        .raw_client_data
        .iter()
        .map(|raw| String::from_utf8_lossy(raw).into_owned())
        .collect()
}

// features: Form control unmodeled ClientData passthrough
#[test]
fn unmodeled_client_data_round_trips_on_checkbox_xlsx() {
    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_form_control(
        FormControl::new(FormControlKind::Checkbox {
            caption: "Audit".into(),
            state: CheckState::Checked,
            cell_link: None,
            no_3d: true,
        }),
        DrawingAnchor::default(),
    ).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    XlsxWriter::write(&workbook, &mut bytes).expect("write xlsx");
    let spliced = splice_into_client_data(
        bytes.into_inner(),
        "Checkbox",
        "   <x:Disabled/>\n   <x:Accel>65</x:Accel>\n  ",
    );

    let reopened = XlsxReader::read(Cursor::new(spliced)).expect("read spliced xlsx");
    assert_eq!(
        control_raws(&reopened),
        vec!["<x:Disabled/>".to_string(), "<x:Accel>65</x:Accel>".to_string()],
        "unmodeled ClientData children captured on a modeled kind"
    );

    let mut rewritten = Cursor::new(Vec::new());
    XlsxWriter::write(&reopened, &mut rewritten).expect("rewrite xlsx");
    let rewritten = rewritten.into_inner();
    let vml = vml_parts(&rewritten);
    assert_eq!(vml.matches("<x:Disabled/>").count(), 1);
    assert_eq!(vml.matches("<x:Accel>65</x:Accel>").count(), 1);
    assert_eq!(vml.matches("<x:Checked>").count(), 1, "no double representation");

    let reread = XlsxReader::read(Cursor::new(rewritten)).expect("reread xlsx");
    assert_eq!(control_raws(&reread), control_raws(&reopened));
}

// features: Form control unmodeled ClientData passthrough
#[test]
fn unmodeled_client_data_round_trips_on_button_xlsb() {
    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_form_control(
        FormControl::new(FormControlKind::Button {
            caption: "OK".into(),
        }),
        DrawingAnchor::default(),
    ).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    XlsbWriter::write(&workbook, &mut bytes).expect("write xlsb");
    let spliced = splice_into_client_data(
        bytes.into_inner(),
        "Button",
        "   <x:Default/>\n   <x:Cancel/>\n  ",
    );

    let reopened = XlsbReader::read(Cursor::new(spliced)).expect("read spliced xlsb");
    assert_eq!(
        control_raws(&reopened),
        vec!["<x:Default/>".to_string(), "<x:Cancel/>".to_string()],
        "dialog button semantics captured on a modeled kind"
    );

    let mut rewritten = Cursor::new(Vec::new());
    XlsbWriter::write(&reopened, &mut rewritten).expect("rewrite xlsb");
    let rewritten = rewritten.into_inner();
    let vml = vml_parts(&rewritten);
    assert_eq!(vml.matches("<x:Default/>").count(), 1);
    assert_eq!(vml.matches("<x:Cancel/>").count(), 1);

    let reread = XlsbReader::read(Cursor::new(rewritten)).expect("reread xlsb");
    assert_eq!(control_raws(&reread), control_raws(&reopened));
}

// features: Form control: dropdown (combo box)
#[test]
fn uiobj_marker_wins_over_a_ctrlprops_twin_xlsx() {
    // A UIObj-marked VML shape with a worksheet <control> entry and
    // ctrlProps part is self-contradictory; the marker wins, matching
    // the XLS reader's unconditional fUIObj skip.
    let mut workbook = Workbook::new();
    workbook.worksheet_mut(0).unwrap().add_form_control(
        FormControl::new(FormControlKind::Checkbox {
            caption: "Audit".into(),
            state: CheckState::Checked,
            cell_link: None,
            no_3d: true,
        }),
        DrawingAnchor::default(),
    ).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    XlsxWriter::write(&workbook, &mut bytes).expect("write xlsx");
    let spliced = splice_into_client_data(bytes.into_inner(), "Checkbox", "   <x:UIObj/>\n  ");
    let reopened = XlsxReader::read(Cursor::new(spliced)).expect("read spliced xlsx");
    assert_eq!(
        reopened.worksheet(0).unwrap().form_control_count(),
        0,
        "UIObj-marked shape must not surface even with a ctrlProps twin"
    );
}

// features: Form control unmodeled ClientData passthrough
#[test]
fn malformed_raw_client_data_is_rejected_at_write() {
    let mut workbook = Workbook::new();
    let mut control = FormControl::new(FormControlKind::Checkbox {
        caption: "Audit".into(),
        state: CheckState::Unchecked,
        cell_link: None,
        no_3d: false,
    });
    control.raw_client_data = vec![b"<x:Oops>".to_vec()];
    workbook
        .worksheet_mut(0)
        .unwrap()
        .add_form_control(control, DrawingAnchor::default()).unwrap();

    let xlsx_err = XlsxWriter::write(&workbook, &mut Cursor::new(Vec::new()))
        .expect_err("unbalanced raw ClientData must not produce a corrupt XLSX part");
    assert!(
        xlsx_err.to_string().contains("ClientData"),
        "error names the raw ClientData fragment: {xlsx_err}"
    );
    let xlsb_err = XlsbWriter::write(&workbook, &mut Cursor::new(Vec::new()))
        .expect_err("unbalanced raw ClientData must not produce a corrupt XLSB part");
    assert!(
        xlsb_err.to_string().contains("ClientData"),
        "error names the raw ClientData fragment: {xlsb_err}"
    );

    // Fragment shapes that are balanced but would defeat the
    // duplicate-name guard or corrupt the part must also be rejected.
    let hostile: [&[u8]; 18] = [
        // Multi-root: second root smuggles a modeled element past the
        // guard (it inspects only the first element's name).
        b"<x:Disabled/><x:Checked>0</x:Checked>",
        // Comment prefix: the guard would read the name as `!--`.
        b"<!--c--><x:Checked>0</x:Checked>",
        // XML declaration mid-part is malformed.
        b"<?xml version=\"1.0\"?><x:A/>",
        // Unquoted attribute value is not well-formed XML.
        b"<x:A foo=bar>t</x:A>",
        // Undefined entity reference makes the part ill-formed.
        b"<x:A>&bogus;</x:A>",
        // Surrogate character reference is not a valid XML character.
        b"<x:A>&#xD800;</x:A>",
        // Nested XML declaration is malformed anywhere but the start.
        b"<x:A><?xml version=\"1.0\"?></x:A>",
        // DOCTYPE inside content is malformed.
        b"<x:A><!DOCTYPE d></x:A>",
        // Empty local name is not a valid element name.
        b"<x:/>",
        // Undefined entity in an attribute value.
        b"<x:A v=\"&bogus;\"/>",
        // Literal ]]> in character data (XML 1.0 section 2.4).
        b"<x:A>]]></x:A>",
        // Double hyphen inside a comment (section 2.5).
        b"<x:A><!-- -- --></x:A>",
        // Comment ending in ---> is the same violation.
        b"<x:A><!--x---></x:A>",
        // Character reference to a control char invalid in XML.
        b"<x:A>&#11;</x:A>",
        // Raw control character invalid in XML content.
        b"<x:A>\x0b</x:A>",
        // Unicode noncharacter in content.
        "<x:A>\u{FFFE}</x:A>".as_bytes(),
        // Unescaped < in an attribute value (section 3.1).
        b"<x:A b=\"<\"/>",
        // Form feed is not XML whitespace outside the element.
        b"<x:A/>\x0c",
    ];
    for fragment in hostile {
        let mut workbook = Workbook::new();
        let mut control = FormControl::new(FormControlKind::Checkbox {
            caption: "Audit".into(),
            state: CheckState::Unchecked,
            cell_link: None,
            no_3d: false,
        });
        control.raw_client_data = vec![fragment.to_vec()];
        workbook
            .worksheet_mut(0)
            .unwrap()
            .add_form_control(control, DrawingAnchor::default()).unwrap();
        let error = XlsxWriter::write(&workbook, &mut Cursor::new(Vec::new())).expect_err(
            &format!(
                "hostile fragment must be rejected: {}",
                String::from_utf8_lossy(fragment)
            ),
        );
        assert!(
            error.to_string().contains("ClientData"),
            "error names the raw ClientData fragment: {error}"
        );
    }
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

// features: Unknown legacy Forms controls
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
    });
    assert!(pict.validate().is_err());
    let mut pict_workbook = Workbook::new();
    // Unchecked: the control fails validation on purpose; the writer
    // must still reject it for files read permissively.
    pict_workbook
        .worksheet_mut(0)
        .unwrap()
        .drawings_mut()
        .push(DrawingObject::form_control(pict).with_anchor(DrawingAnchor::default()));
    let mut output = Cursor::new(Vec::new());
    let error = XlsbWriter::write(&pict_workbook, &mut output).unwrap_err();
    assert!(error.to_string().contains("ActiveX/OLE"));

    let blank = FormControl::new(FormControlKind::Unknown {
        object_type: "  ".to_string(),
        legacy_object_type: None,
        caption: "blank type".into(),
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
    ).unwrap();
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
    ).unwrap();
    workbook
}

fn assert_metric_anchors(workbook: &Workbook, tolerance_emu: i64) {
    let close = |actual: i64, expected: i64| {
        assert!(
            (actual - expected).abs() <= tolerance_emu,
            "anchor offset {actual} differs from {expected} by more than {tolerance_emu} EMU"
        );
    };
    let controls: Vec<_> = workbook.worksheet(0).unwrap().form_controls().collect();
    assert_eq!(controls.len(), 2);
    match &controls[0].object.anchor {
        DrawingAnchor::TwoCell { from, to, .. } => {
            assert_eq!((from.col, from.col_offset_emu), (0, 0));
            assert_eq!((from.row, from.row_offset_emu), (0, 0));
            assert_eq!(to.col, 0);
            close(to.col_offset_emu, 609_600);
            assert_eq!(to.row, 0);
            close(to.row_offset_emu, 190_500);
        }
        other => panic!("expected flattened one-cell control anchor, got {other:?}"),
    }
    match &controls[1].object.anchor {
        DrawingAnchor::TwoCell { from, .. } => {
            assert_eq!(from.col, 0);
            close(from.col_offset_emu, 609_600);
            assert_eq!(from.row, 0);
            close(from.row_offset_emu, 190_500);
        }
        other => panic!("expected flattened absolute control anchor, got {other:?}"),
    }
}

// features: Form-control positioning with custom cell metrics
#[test]
fn custom_dimensions_drive_all_format_control_anchor_flattening() {
    let workbook = metric_anchor_workbook();

    let mut xlsx = Cursor::new(Vec::new());
    XlsxWriter::write(&workbook, &mut xlsx).expect("write xlsx");
    let xlsx = XlsxReader::read(Cursor::new(xlsx.into_inner())).expect("read xlsx");
    assert_metric_anchors(&xlsx, 0);

    let mut xlsb = Cursor::new(Vec::new());
    XlsbWriter::write(&workbook, &mut xlsb).expect("write xlsb");
    let xlsb = XlsbReader::read(Cursor::new(xlsb.into_inner())).expect("read xlsb");
    assert_metric_anchors(&xlsb, 0);

    let xls = XlsWriter::write_to_bytes(&workbook).expect("write xls");
    let xls = XlsReader::read(Cursor::new(xls)).expect("read xls");
    // BIFF8 stores offsets in 1/1024-column and 1/256-row units.
    assert_metric_anchors(&xls, 1_500);
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

/// Hidden rows and columns render at zero extent in Excel, so drawing
/// metrics must give them zero width/height and exclude them from
/// positions when flattening anchors.
#[test]
fn hidden_rows_and_columns_have_zero_drawing_extent() {
    let mut sheet = Worksheet::new("Hidden");
    sheet.set_column_width(1, 20.0);
    sheet.set_column_hidden(1, true);
    sheet.set_column_hidden(2, true);
    sheet.set_column_hidden(2, false);
    sheet.set_row_hidden(0, true);

    assert_eq!(sheet.column_width_emu(1), 0, "hidden column has no width");
    assert_eq!(sheet.row_height_emu(0), 0, "hidden row has no height");
    assert_eq!(
        sheet.column_width_emu(2),
        column_width_to_emu(sheet.default_column_width()),
        "re-shown column recovers its width"
    );

    // Column 3 starts after two visible default columns: the hidden
    // custom column contributes nothing.
    let default_col = i128::from(column_width_to_emu(sheet.default_column_width()));
    assert_eq!(sheet.column_position_emu(3), 2 * default_col);
    // Row 2 starts after one visible default row.
    let default_row = i128::from(row_height_to_emu(sheet.default_row_height()));
    assert_eq!(sheet.row_position_emu(2), default_row);
}

/// Positions strictly inside the first visible row/column after a
/// hidden run must resolve to that visible cell, not collapse to the
/// zero-extent run head (which would truncate extents and emit
/// reversed anchors).
#[test]
fn anchors_flatten_into_the_first_visible_cell_after_a_hidden_run() {
    let mut sheet = Worksheet::new("Hidden");
    for row in 5..=8 {
        sheet.set_row_hidden(row, true);
    }
    for col in 2..=3 {
        sheet.set_column_hidden(col, true);
    }
    // Rows 5-8 are hidden, so rows 5..=9 share a start position.
    assert_eq!(sheet.row_position_emu(9), sheet.row_position_emu(5));

    let anchor = DrawingAnchor::OneCell {
        from: CellMarker {
            col: 4,
            col_offset_emu: 40_000,
            row: 9,
            row_offset_emu: 40_000,
        },
        width_emu: 100_000,
        height_emu: 100_000,
    };
    let DrawingAnchor::TwoCell { from, to, .. } = anchor.to_two_cell_with_metrics(&sheet) else {
        panic!("to_two_cell always returns TwoCell");
    };
    assert_eq!((from.col, from.col_offset_emu), (4, 40_000));
    assert_eq!((from.row, from.row_offset_emu), (9, 40_000));
    assert_eq!(
        (to.row, to.row_offset_emu),
        (9, 140_000),
        "extent ends inside the visible row, not at the hidden run head"
    );
    assert_eq!((to.col, to.col_offset_emu), (4, 140_000));

    // Exactly at the shared boundary, the run head still wins so the
    // resolution stays stable for degenerate metrics.
    let boundary = DrawingAnchor::Absolute {
        x_emu: sheet.column_position_emu(2) as i64,
        y_emu: sheet.row_position_emu(5) as i64,
        width_emu: 1_000,
        height_emu: 1_000,
    };
    let DrawingAnchor::TwoCell { from, .. } = boundary.to_two_cell_with_metrics(&sheet) else {
        panic!("to_two_cell always returns TwoCell");
    };
    assert_eq!((from.col, from.col_offset_emu), (2, 0));
    assert_eq!((from.row, from.row_offset_emu), (5, 0));
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
    ).unwrap();
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
    )).unwrap();
    sheet.add_drawing(radio(
        "inside",
        DrawingAnchor::Absolute {
            x_emu: 100_000,
            y_emu: 100_000,
            width_emu: 100_000,
            height_emu: 100_000,
        },
    )).unwrap();

    let placed = sheet.placed_form_controls();
    assert_eq!(placed[1].rect_emu.x_emu, 1_381_125);
    assert_eq!(radio_groups(&placed), vec![vec![1], vec![2]]);
}
