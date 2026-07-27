//! LibreOffice envelope checks for the XLSX writer.
//!
//! Smoke tests: write a workbook, have LibreOffice open it via URP,
//! and read a cell value back. If a part is malformed LO refuses the
//! load and `open_workbook` errors. Feature survival is asserted by
//! the Excel COM parity layer instead.

use duke_sheets_chart::{CellMarker, DrawingAnchor};
use duke_sheets_core::{CheckState, FormControl, FormControlKind, ListSelection, Workbook};
use duke_sheets_xlsx::XlsxWriter;

use crate::{lo_bridge, runtime, temp_fixture_path};

#[test]
fn lo_can_open_xlsx_with_form_controls_we_emit() {
    let anchor = |fc: u16, fr: u32, tc: u16, tr: u32| DrawingAnchor::TwoCell {
        from: CellMarker {
            col: fc,
            col_offset_emu: 0,
            row: fr,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: tc,
            col_offset_emu: 0,
            row: tr,
            row_offset_emu: 0,
        },
        edit_as: None,
    };

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
        let mut control = FormControl::new(kind);
        if matches!(control.kind, FormControlKind::Checkbox { .. }) {
            // Envelope check for the raw ClientData passthrough: the
            // replayed children must not make LO reject the VML part.
            control.raw_client_data =
                vec![b"<x:Disabled/>".to_vec(), b"<x:Accel>65</x:Accel>".to_vec()];
        }
        ws.add_form_control(control, anchor(1, row, 3, row + 1)).unwrap();
    }
    assert_eq!(wb.sync_form_control_links(), 1);

    let path = temp_fixture_path();
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, std::io::Cursor::new(&mut buf)).expect("serialize");
    std::fs::write(&path, &buf).expect("write fixture");

    let outcome: Result<f64, String> = runtime().block_on(async {
        let bridge = lo_bridge().await.expect("bridge");
        let mut bridge = bridge.lock().await;
        let mut wb_in = bridge
            .open_workbook(path.to_str().unwrap())
            .await
            .map_err(|e| format!("open: {e}"))?;
        wb_in
            .get_cell_value("A1")
            .await
            .map_err(|e| format!("A1: {e}"))
    });
    let _ = std::fs::remove_file(&path);
    let a1 = outcome.expect("LO must open our XLSX with form controls without error");
    assert!(
        (a1 - 42.0).abs() < 1e-9,
        "A1 must round-trip; got {a1} (expected 42)"
    );
}

/// A chartEx built through the model carries generated chart style and
/// chart colour style parts. LO does not render chartEx, but it must
/// still load the package: a malformed part makes its loader refuse the
/// file. Feature survival is asserted by the Excel COM parity layer.
#[test]
fn lo_can_open_xlsx_with_chartex_we_emit() {
    use duke_sheets_core::DrawingObject;

    let chart_ex = duke_sheets_chart::parse::parse_chart_ex_xml(
        &br#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>Sheet1!$A$1:$A$3</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">a</cx:pt><cx:pt idx="1">b</cx:pt><cx:pt idx="2">c</cx:pt></cx:lvl></cx:strDim><cx:numDim type="val"><cx:f>Sheet1!$B$1:$B$3</cx:f><cx:lvl ptCount="3" formatCode="General"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt><cx:pt idx="2">3</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall" uniqueId="{1D8F9C4E-1C1B-4A5F-9C6B-2E7A0F3B5D11}"><cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Series1</cx:v></cx:txData></cx:tx><cx:dataId val="0"/><cx:layoutPr><cx:subtotals><cx:idx val="0"/></cx:subtotals></cx:layoutPr></cx:series></cx:plotAreaRegion><cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>"#[..],
    )
    .expect("parse chartEx");

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", 42.0).expect("A1");
    ws.add_drawing(
        DrawingObject::chart_ex(chart_ex).with_anchor(DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 3,
                col_offset_emu: 0,
                row: 0,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 10,
                col_offset_emu: 0,
                row: 15,
                row_offset_emu: 0,
            },
            edit_as: None,
        }),
    )
    .unwrap();

    let path = temp_fixture_path();
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, std::io::Cursor::new(&mut buf)).expect("serialize");
    std::fs::write(&path, &buf).expect("write fixture");

    let outcome: Result<f64, String> = runtime().block_on(async {
        let bridge = lo_bridge().await.expect("bridge");
        let mut bridge = bridge.lock().await;
        let mut wb_in = bridge
            .open_workbook(path.to_str().unwrap())
            .await
            .map_err(|e| format!("open: {e}"))?;
        wb_in
            .get_cell_value("A1")
            .await
            .map_err(|e| format!("A1: {e}"))
    });
    let _ = std::fs::remove_file(&path);
    let a1 = outcome.expect("LO must open our XLSX with a chartEx without error");
    assert!((a1 - 42.0).abs() < 1e-9, "A1 must round-trip; got {a1}");
}
