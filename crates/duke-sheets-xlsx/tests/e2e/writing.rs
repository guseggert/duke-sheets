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
#[ignore = "requires LibreOffice URP on 127.0.0.1:2002"]
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
            caption: "Run".to_string(),
        },
        FormControlKind::Checkbox {
            caption: "Check".to_string(),
            state: CheckState::Checked,
            cell_link: Some("$D$2".to_string()),
            no_3d: false,
        },
        FormControlKind::OptionButton {
            caption: "Opt".to_string(),
            state: CheckState::Checked,
            cell_link: None,
            first_in_group: true,
            no_3d: false,
        },
        FormControlKind::Label {
            caption: "Info".to_string(),
        },
        FormControlKind::GroupBox {
            caption: "Frame".to_string(),
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
        ws.add_form_control(FormControl::with_anchor(kind, anchor(1, row, 3, row + 1)));
    }

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
