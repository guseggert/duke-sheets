#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use duke_sheets_core::style::{
        BorderEdge, BorderLineStyle, Color, FillStyle, NumberFormat, Style,
    };
    use duke_sheets_core::{CellValue, Workbook};

    use crate::reader::XlsbReader;
    use crate::writer::XlsbWriter;

    fn round_trip(wb: &Workbook) -> Workbook {
        let mut buf = Vec::new();
        XlsbWriter::write(wb, Cursor::new(&mut buf)).unwrap();
        XlsbReader::read(Cursor::new(&buf)).unwrap()
    }

    #[test]
    fn empty_workbook() {
        let wb = Workbook::new();
        let wb2 = round_trip(&wb);
        assert_eq!(wb2.sheet_count(), 1);
    }

    #[test]
    fn string_values() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "hello").unwrap();
        ws.set_cell_value_at(0, 1, "world").unwrap();
        ws.set_cell_value_at(1, 0, "hello").unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("hello"));
        assert_eq!(ws2.get_value_at(0, 1), CellValue::string("world"));
        assert_eq!(ws2.get_value_at(1, 0), CellValue::string("hello"));
    }

    #[test]
    fn numeric_values() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 42.0).unwrap();
        ws.set_cell_value_at(0, 1, -3.14).unwrap();
        ws.set_cell_value_at(0, 2, 0.0).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Number(42.0));
        assert_eq!(ws2.get_value_at(0, 1), CellValue::Number(-3.14));
        assert_eq!(ws2.get_value_at(0, 2), CellValue::Number(0.0));
    }

    #[test]
    fn boolean_values() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, true).unwrap();
        ws.set_cell_value_at(0, 1, false).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Boolean(true));
        assert_eq!(ws2.get_value_at(0, 1), CellValue::Boolean(false));
    }

    #[test]
    fn error_values() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, CellValue::Error(duke_sheets_core::CellError::Div0))
            .unwrap();
        ws.set_cell_value_at(0, 1, CellValue::Error(duke_sheets_core::CellError::Na))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(
            ws2.get_value_at(0, 0),
            CellValue::Error(duke_sheets_core::CellError::Div0)
        );
        assert_eq!(
            ws2.get_value_at(0, 1),
            CellValue::Error(duke_sheets_core::CellError::Na)
        );
    }

    #[test]
    fn multiple_sheets() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0).unwrap().set_name("First");
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "sheet1")
            .unwrap();
        wb.add_worksheet().unwrap();
        wb.worksheet_mut(1).unwrap().set_name("Second");
        wb.worksheet_mut(1)
            .unwrap()
            .set_cell_value_at(0, 0, 99.0)
            .unwrap();

        let wb2 = round_trip(&wb);
        assert_eq!(wb2.sheet_count(), 2);
        assert_eq!(wb2.worksheet(0).unwrap().name(), "First");
        assert_eq!(wb2.worksheet(1).unwrap().name(), "Second");
        assert_eq!(
            wb2.worksheet(0).unwrap().get_value_at(0, 0),
            CellValue::string("sheet1")
        );
        assert_eq!(
            wb2.worksheet(1).unwrap().get_value_at(0, 0),
            CellValue::Number(99.0)
        );
    }

    #[test]
    fn mixed_cell_types() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "text").unwrap();
        ws.set_cell_value_at(0, 1, 123.456).unwrap();
        ws.set_cell_value_at(0, 2, true).unwrap();
        ws.set_cell_value_at(1, 0, CellValue::Error(duke_sheets_core::CellError::Value))
            .unwrap();
        ws.set_cell_value_at(2, 5, "sparse").unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("text"));
        assert_eq!(ws2.get_value_at(0, 1), CellValue::Number(123.456));
        assert_eq!(ws2.get_value_at(0, 2), CellValue::Boolean(true));
        assert_eq!(
            ws2.get_value_at(1, 0),
            CellValue::Error(duke_sheets_core::CellError::Value)
        );
        assert_eq!(ws2.get_value_at(2, 5), CellValue::string("sparse"));
        assert_eq!(ws2.get_value_at(1, 1), CellValue::Empty);
    }

    #[test]
    fn unicode_strings() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "café").unwrap();
        ws.set_cell_value_at(0, 1, "日本語").unwrap();
        ws.set_cell_value_at(0, 2, "😀").unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("café"));
        assert_eq!(ws2.get_value_at(0, 1), CellValue::string("日本語"));
        assert_eq!(ws2.get_value_at(0, 2), CellValue::string("😀"));
    }

    #[test]
    fn write_file_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "file test").unwrap();
        ws.set_cell_value_at(0, 1, 42.0).unwrap();

        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.xlsb");
        XlsbWriter::write_file(&wb, &path).unwrap();

        let wb2 = XlsbReader::read_file(&path).unwrap();
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("file test"));
        assert_eq!(ws2.get_value_at(0, 1), CellValue::Number(42.0));
    }

    #[test]
    fn style_bold_font_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "bold").unwrap();
        let style = Style::new()
            .bold(true)
            .font_size(14.0)
            .font_color(Color::RED);
        ws.set_cell_style_at(0, 0, &style).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("bold"));
        let s = ws2.cell_style_at(0, 0).unwrap();
        assert!(s.font.bold);
        assert_eq!(s.font.size, 14.0);
        assert_eq!(s.font.color, Color::Rgb { r: 255, g: 0, b: 0 });
    }

    #[test]
    fn style_number_format_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 3.14159).unwrap();
        let style = Style::new().number_format("0.00");
        ws.set_cell_style_at(0, 0, &style).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Number(3.14159));
        let s = ws2.cell_style_at(0, 0).unwrap();
        assert_eq!(s.number_format, NumberFormat::Custom("0.00".to_string()));
    }

    #[test]
    fn style_fill_color_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "colored").unwrap();
        let style = Style::new().fill_color(Color::rgb(0, 128, 255));
        ws.set_cell_style_at(0, 0, &style).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let s = ws2.cell_style_at(0, 0).unwrap();
        assert_eq!(
            s.fill,
            FillStyle::Solid {
                color: Color::Rgb {
                    r: 0,
                    g: 128,
                    b: 255
                }
            }
        );
    }

    #[test]
    fn style_border_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "bordered").unwrap();
        let mut style = Style::new();
        style.border = style
            .border
            .clone()
            .with_bottom(BorderLineStyle::Thin, Color::BLACK)
            .with_top(BorderLineStyle::Medium, Color::RED);
        ws.set_cell_style_at(0, 0, &style).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let s = ws2.cell_style_at(0, 0).unwrap();
        assert_eq!(
            s.border.bottom,
            Some(BorderEdge::new(BorderLineStyle::Thin, Color::BLACK))
        );
        assert_eq!(
            s.border.top,
            Some(BorderEdge::new(BorderLineStyle::Medium, Color::RED))
        );
    }

    #[test]
    fn formula_cached_number_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 10.0).unwrap();
        ws.set_cell_value_at(0, 1, 20.0).unwrap();
        ws.set_formula_with_cached_value_at(0, 2, "=A1+B1", CellValue::Number(30.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 2), CellValue::Number(30.0));
    }

    #[test]
    fn formula_cached_string_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(
            0,
            0,
            "=CONCATENATE(\"hello\",\" \",\"world\")",
            CellValue::string("hello world"),
        )
        .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("hello world"));
    }

    #[test]
    fn formula_cached_bool_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=TRUE", CellValue::Boolean(true))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Boolean(true));
    }

    #[test]
    fn formula_cached_error_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(
            0,
            0,
            "=1/0",
            CellValue::Error(duke_sheets_core::CellError::Div0),
        )
        .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(
            ws2.get_value_at(0, 0),
            CellValue::Error(duke_sheets_core::CellError::Div0)
        );
    }

    #[test]
    fn multiple_styles_across_cells() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();

        ws.set_cell_value_at(0, 0, "bold").unwrap();
        ws.set_cell_style_at(0, 0, &Style::new().bold(true))
            .unwrap();

        ws.set_cell_value_at(0, 1, "italic").unwrap();
        ws.set_cell_style_at(0, 1, &Style::new().italic(true))
            .unwrap();

        ws.set_cell_value_at(0, 2, "red bg").unwrap();
        ws.set_cell_style_at(0, 2, &Style::new().fill_color(Color::RED))
            .unwrap();

        ws.set_cell_value_at(1, 0, "default").unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();

        let s0 = ws2.cell_style_at(0, 0).unwrap();
        assert!(s0.font.bold);
        assert!(!s0.font.italic);

        let s1 = ws2.cell_style_at(0, 1).unwrap();
        assert!(!s1.font.bold);
        assert!(s1.font.italic);

        let s2 = ws2.cell_style_at(0, 2).unwrap();
        assert_eq!(
            s2.fill,
            FillStyle::Solid {
                color: Color::Rgb { r: 255, g: 0, b: 0 }
            }
        );

        let s3 = ws2.cell_style_at(1, 0);
        assert!(s3.is_none() || *s3.unwrap() == Style::default());
    }

    #[test]
    fn style_deduplication_across_sheets() {
        let mut wb = Workbook::new();
        let bold = Style::new().bold(true);

        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "sheet1").unwrap();
        ws.set_cell_style_at(0, 0, &bold).unwrap();

        wb.add_worksheet().unwrap();
        let ws2 = wb.worksheet_mut(1).unwrap();
        ws2.set_cell_value_at(0, 0, "sheet2").unwrap();
        ws2.set_cell_style_at(0, 0, &bold).unwrap();

        let wb2 = round_trip(&wb);
        let s1 = wb2.worksheet(0).unwrap().cell_style_at(0, 0).unwrap();
        let s2 = wb2.worksheet(1).unwrap().cell_style_at(0, 0).unwrap();
        assert!(s1.font.bold);
        assert!(s2.font.bold);
    }

    #[test]
    fn formula_with_style_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=1+2", CellValue::Number(3.0))
            .unwrap();
        ws.set_cell_style_at(0, 0, &Style::new().bold(true))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Number(3.0));
        let s = ws2.cell_style_at(0, 0).unwrap();
        assert!(s.font.bold);
    }

    #[test]
    fn builtin_number_format_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 0.5).unwrap();
        let mut style = Style::new();
        style.number_format = NumberFormat::BuiltIn(10);
        ws.set_cell_style_at(0, 0, &style).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let s = ws2.cell_style_at(0, 0).unwrap();
        assert_eq!(s.number_format, NumberFormat::BuiltIn(10));
    }

    #[test]
    fn merged_cells_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "merged").unwrap();
        ws.merge_cells(&duke_sheets_core::CellRange::parse("A1:C3").unwrap())
            .unwrap();
        ws.merge_cells(&duke_sheets_core::CellRange::parse("E5:F6").unwrap())
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let regions = ws2.merged_regions();
        assert_eq!(regions.len(), 2);
        assert_eq!(regions[0].to_string(), "A1:C3");
        assert_eq!(regions[1].to_string(), "E5:F6");
    }

    #[test]
    fn row_height_and_hidden_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "row0").unwrap();
        ws.set_row_height(0, 30.0);
        ws.set_cell_value_at(1, 0, "row1").unwrap();
        ws.set_row_hidden(1, true);
        ws.set_cell_value_at(2, 0, "row2").unwrap();
        ws.set_row_height(2, 45.5);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert!((ws2.row_height(0) - 30.0).abs() < 0.1);
        assert!(ws2.is_row_hidden(1));
        assert!((ws2.row_height(2) - 45.5).abs() < 0.1);
    }

    #[test]
    fn column_width_and_hidden_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "col0").unwrap();
        ws.set_column_width(0, 20.0);
        ws.set_column_hidden(1, true);
        ws.set_column_width(2, 5.5);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert!((ws2.column_width(0) - 20.0).abs() < 0.1);
        assert!(ws2.is_column_hidden(1));
        assert!((ws2.column_width(2) - 5.5).abs() < 0.1);
    }

    #[test]
    fn freeze_pane_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "frozen").unwrap();
        ws.set_freeze_panes(2, 1);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let fp = ws2.freeze_panes().expect("freeze panes should exist");
        assert_eq!(fp.row, 2);
        assert_eq!(fp.col, 1);
    }

    #[test]
    fn page_margins_roundtrip() {
        use duke_sheets_core::worksheet::PageSetup;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "margins").unwrap();
        let mut ps = PageSetup::default();
        ps.left_margin = 1.5;
        ps.right_margin = 1.5;
        ps.top_margin = 2.0;
        ps.bottom_margin = 2.0;
        ps.header_margin = 0.5;
        ps.footer_margin = 0.5;
        ws.set_page_setup(ps);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let ps2 = ws2.page_setup();
        assert!((ps2.left_margin - 1.5).abs() < 0.001);
        assert!((ps2.right_margin - 1.5).abs() < 0.001);
        assert!((ps2.top_margin - 2.0).abs() < 0.001);
        assert!((ps2.bottom_margin - 2.0).abs() < 0.001);
        assert!((ps2.header_margin - 0.5).abs() < 0.001);
        assert!((ps2.footer_margin - 0.5).abs() < 0.001);
    }

    #[test]
    fn page_setup_landscape_roundtrip() {
        use duke_sheets_core::worksheet::{PageOrientation, PageSetup};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "landscape").unwrap();
        let mut ps = PageSetup::default();
        ps.orientation = PageOrientation::Landscape;
        ps.paper_size = 9;
        ps.scale = 75;
        ws.set_page_setup(ps);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let ps2 = ws2.page_setup();
        assert_eq!(ps2.orientation, PageOrientation::Landscape);
        assert_eq!(ps2.paper_size, 9);
        assert_eq!(ps2.scale, 75);
    }

    #[test]
    fn autofilter_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "Name").unwrap();
        ws.set_cell_value_at(0, 1, "Value").unwrap();
        ws.set_cell_value_at(1, 0, "A").unwrap();
        ws.set_cell_value_at(1, 1, 1.0).unwrap();
        let range = duke_sheets_core::CellRange::parse("A1:B2").unwrap();
        ws.set_auto_filter(Some(duke_sheets_core::AutoFilter::new(range)));

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let af = ws2.auto_filter().expect("autofilter should exist");
        assert_eq!(af.range.to_string(), "A1:B2");
    }

    #[test]
    fn header_footer_roundtrip() {
        use duke_sheets_core::worksheet::PageSetup;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "hf").unwrap();
        let mut ps = PageSetup::default();
        ps.odd_header = Some("&CPage &P".to_string());
        ps.odd_footer = Some("&LLeft&RRight".to_string());
        ws.set_page_setup(ps);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let ps2 = ws2.page_setup();
        assert_eq!(ps2.odd_header.as_deref(), Some("&CPage &P"));
        assert_eq!(ps2.odd_footer.as_deref(), Some("&LLeft&RRight"));
    }

    #[test]
    fn comments_roundtrip() {
        use duke_sheets_core::comment::CellComment;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "commented").unwrap();
        ws.set_comment_at(0, 0, CellComment::new("Author1", "First comment"));
        ws.set_cell_value_at(1, 1, "also commented").unwrap();
        ws.set_comment_at(1, 1, CellComment::new("Author2", "Second comment"));

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let c1 = ws2.comment_at(0, 0).expect("comment at A1");
        assert_eq!(c1.author, "Author1");
        assert_eq!(c1.text, "First comment");
        let c2 = ws2.comment_at(1, 1).expect("comment at B2");
        assert_eq!(c2.author, "Author2");
        assert_eq!(c2.text, "Second comment");
    }

    #[test]
    fn external_hyperlink_roundtrip() {
        use duke_sheets_core::Hyperlink;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "click me").unwrap();
        ws.set_hyperlink(
            "A1",
            Hyperlink {
                target: "https://example.com".to_string(),
                display: Some("Example".to_string()),
                tooltip: Some("Go to example".to_string()),
                location: None,
            },
        )
        .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let link = ws2.hyperlink_at(0, 0).expect("hyperlink at A1");
        assert_eq!(link.target, "https://example.com");
        assert_eq!(link.display.as_deref(), Some("Example"));
        assert_eq!(link.tooltip.as_deref(), Some("Go to example"));
    }

    #[test]
    fn internal_hyperlink_roundtrip() {
        use duke_sheets_core::Hyperlink;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "go to B5").unwrap();
        ws.set_hyperlink(
            "A1",
            Hyperlink {
                target: "#Sheet1!B5".to_string(),
                display: None,
                tooltip: None,
                location: None,
            },
        )
        .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let link = ws2.hyperlink_at(0, 0).expect("hyperlink at A1");
        assert_eq!(link.target, "#Sheet1!B5");
    }

    #[test]
    fn hidden_row_no_data_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "visible").unwrap();
        ws.set_row_hidden(5, true);
        ws.set_row_height(5, 25.0);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert!(ws2.is_row_hidden(5));
        assert!((ws2.row_height(5) - 25.0).abs() < 0.1);
    }

    #[test]
    fn data_validation_roundtrip() {
        use duke_sheets_core::validation::{DataValidation, ValidationOperator, ValidationType};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "validated").unwrap();
        let dv = DataValidation {
            validation_type: ValidationType::Whole {
                operator: ValidationOperator::GreaterThan,
                value1: "0".to_string(),
                value2: None,
            },
            ranges: vec![duke_sheets_core::CellRange::parse("A1:A10").unwrap()],
            allow_blank: true,
            show_dropdown: true,
            show_input_message: true,
            input_title: Some("Input".to_string()),
            input_message: Some("Enter a number".to_string()),
            show_error_alert: true,
            error_style: duke_sheets_core::validation::ValidationErrorStyle::Stop,
            error_title: Some("Error".to_string()),
            error_message: Some("Must be positive".to_string()),
        };
        ws.add_data_validation(dv);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let dvs = ws2.data_validations();
        assert_eq!(dvs.len(), 1);
        assert_eq!(dvs[0].ranges.len(), 1);
        assert_eq!(dvs[0].ranges[0].to_string(), "A1:A10");
        assert_eq!(dvs[0].error_title.as_deref(), Some("Error"));
        assert_eq!(dvs[0].error_message.as_deref(), Some("Must be positive"));
        assert_eq!(dvs[0].input_title.as_deref(), Some("Input"));
        assert_eq!(dvs[0].input_message.as_deref(), Some("Enter a number"));
        assert!(dvs[0].allow_blank);
        assert!(dvs[0].show_input_message);
        assert!(dvs[0].show_error_alert);
    }

    #[test]
    fn conditional_format_roundtrip() {
        use duke_sheets_core::conditional_format::{CfRuleType, ConditionalFormatRule};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 100.0).unwrap();
        let rule = ConditionalFormatRule::new(CfRuleType::DuplicateValues)
            .with_range(duke_sheets_core::CellRange::parse("A1:A10").unwrap())
            .with_priority(1);
        ws.add_conditional_format(rule);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let cfs = ws2.conditional_formats();
        assert_eq!(cfs.len(), 1);
        assert!(matches!(cfs[0].rule_type, CfRuleType::DuplicateValues));
        assert_eq!(cfs[0].ranges.len(), 1);
        assert_eq!(cfs[0].ranges[0].to_string(), "A1:A10");
    }

    #[test]
    fn combined_features_roundtrip() {
        use duke_sheets_core::comment::CellComment;
        use duke_sheets_core::worksheet::{PageOrientation, PageSetup};
        use duke_sheets_core::{AutoFilter, Hyperlink};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();

        ws.set_cell_value_at(0, 0, "Header").unwrap();
        ws.set_cell_value_at(0, 1, "Value").unwrap();
        ws.set_cell_value_at(1, 0, "A").unwrap();
        ws.set_cell_value_at(1, 1, 42.0).unwrap();

        ws.merge_cells(&duke_sheets_core::CellRange::parse("D1:E2").unwrap())
            .unwrap();
        ws.set_freeze_panes(1, 0);
        ws.set_row_height(0, 20.0);
        ws.set_column_width(0, 15.0);
        ws.set_auto_filter(Some(AutoFilter::new(
            duke_sheets_core::CellRange::parse("A1:B2").unwrap(),
        )));
        ws.set_comment_at(0, 0, CellComment::new("Test", "A note"));
        ws.set_hyperlink(
            "B1",
            Hyperlink {
                target: "https://example.com".to_string(),
                display: None,
                tooltip: None,
                location: None,
            },
        )
        .unwrap();

        let mut ps = PageSetup::default();
        ps.orientation = PageOrientation::Landscape;
        ps.odd_header = Some("&CTest Header".to_string());
        ws.set_page_setup(ps);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();

        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("Header"));
        assert_eq!(ws2.get_value_at(1, 1), CellValue::Number(42.0));
        assert_eq!(ws2.merged_regions().len(), 1);
        assert_eq!(ws2.merged_regions()[0].to_string(), "D1:E2");
        let fp = ws2.freeze_panes().unwrap();
        assert_eq!(fp.row, 1);
        assert_eq!(fp.col, 0);
        assert!((ws2.row_height(0) - 20.0).abs() < 0.1);
        assert!((ws2.column_width(0) - 15.0).abs() < 0.1);
        assert!(ws2.auto_filter().is_some());
        assert!(ws2.comment_at(0, 0).is_some());
        assert!(ws2.hyperlink_at(0, 1).is_some());
        assert_eq!(ws2.page_setup().orientation, PageOrientation::Landscape);
        assert_eq!(
            ws2.page_setup().odd_header.as_deref(),
            Some("&CTest Header")
        );
    }

    #[test]
    fn formula_text_roundtrip_simple_arithmetic() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=1+2", CellValue::Number(3.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Number(3.0));
        let fd = ws2.formula_data_at(0, 0).expect("formula should exist");
        assert_eq!(fd.text, "=1+2");
    }

    #[test]
    fn formula_text_roundtrip_sum() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 10.0).unwrap();
        ws.set_cell_value_at(1, 0, 20.0).unwrap();
        ws.set_formula_with_cached_value_at(2, 0, "=SUM(A1:A2)", CellValue::Number(30.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(2, 0), CellValue::Number(30.0));
        let fd = ws2.formula_data_at(2, 0).expect("formula should exist");
        assert_eq!(fd.text, "=SUM(A1:A2)");
    }

    #[test]
    fn formula_text_roundtrip_string_result() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=\"hello\"", CellValue::string("hello"))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("hello"));
        let fd = ws2.formula_data_at(0, 0).expect("formula should exist");
        assert_eq!(fd.text, "=\"hello\"");
    }

    #[test]
    fn formula_text_roundtrip_boolean_result() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=TRUE", CellValue::Boolean(true))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let fd = ws2.formula_data_at(0, 0).expect("formula should exist");
        assert_eq!(fd.text, "=TRUE");
    }

    #[test]
    fn formula_text_roundtrip_cross_sheet() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0).unwrap().set_name("Sheet1");
        wb.add_worksheet().unwrap();
        wb.worksheet_mut(1).unwrap().set_name("Sheet2");
        wb.worksheet_mut(1)
            .unwrap()
            .set_cell_value_at(0, 0, 42.0)
            .unwrap();
        wb.worksheet_mut(0)
            .unwrap()
            .set_formula_with_cached_value_at(0, 0, "=Sheet2!A1", CellValue::Number(42.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.get_value_at(0, 0), CellValue::Number(42.0));
        let fd = ws2.formula_data_at(0, 0).expect("formula should exist");
        assert_eq!(fd.text, "=Sheet2!A1");
    }

    #[test]
    fn formula_text_roundtrip_if_function() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 5.0).unwrap();
        ws.set_formula_with_cached_value_at(0, 1, "=IF(A1>0,TRUE,FALSE)", CellValue::Boolean(true))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let fd = ws2.formula_data_at(0, 1).expect("formula should exist");
        assert_eq!(fd.text, "=IF(A1>0,TRUE,FALSE)");
    }

    #[test]
    fn formula_text_roundtrip_cell_ref() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 10.0).unwrap();
        ws.set_formula_with_cached_value_at(0, 1, "=A1", CellValue::Number(10.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let fd = ws2.formula_data_at(0, 1).expect("formula should exist");
        assert_eq!(fd.text, "=A1");
    }

    #[test]
    fn formula_text_roundtrip_concat() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=\"a\"&\"b\"", CellValue::string("ab"))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let fd = ws2.formula_data_at(0, 0).expect("formula should exist");
        assert_eq!(fd.text, "=\"a\"&\"b\"");
    }

    #[test]
    fn formula_array_constant_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_formula_with_cached_value_at(0, 0, "=SUM({1,2,3})", CellValue::Number(6.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let fd = ws2.formula_data_at(0, 0).expect("formula should exist");
        assert_eq!(fd.text, "=SUM({1,2,3})");
    }

    #[test]
    fn formula_uplus_paren_roundtrip() {
        // Redundant parens only survive if the compiler emits PtgParen;
        // the leading plus only survives via PtgUplus.
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 2.0).unwrap();
        ws.set_formula_with_cached_value_at(0, 1, "=+A1", CellValue::Number(2.0))
            .unwrap();
        ws.set_formula_with_cached_value_at(0, 2, "=(A1+1)", CellValue::Number(3.0))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let uplus = ws2.formula_data_at(0, 1).expect("formula should exist");
        assert_eq!(uplus.text, "=+A1");
        let paren = ws2.formula_data_at(0, 2).expect("formula should exist");
        assert_eq!(paren.text, "=(A1+1)");
    }

    fn read_zip_entry(data: &[u8], name: &str) -> String {
        let cursor = Cursor::new(data);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut file = archive.by_name(name).unwrap();
        let mut buf = String::new();
        std::io::Read::read_to_string(&mut file, &mut buf).unwrap();
        buf
    }

    fn write_xlsb_bytes(wb: &Workbook) -> Vec<u8> {
        let mut buf = Vec::new();
        XlsbWriter::write(wb, Cursor::new(&mut buf)).unwrap();
        buf
    }

    #[test]
    fn writer_emits_theme_xml() {
        let wb = Workbook::new();
        let bytes = write_xlsb_bytes(&wb);
        let theme = read_zip_entry(&bytes, "xl/theme/theme1.xml");
        assert!(theme.contains("<a:theme"));
        assert!(theme.contains("<a:clrScheme"));
    }

    #[test]
    fn writer_emits_theme_content_type_and_relationship() {
        let wb = Workbook::new();
        let bytes = write_xlsb_bytes(&wb);

        let ct = read_zip_entry(&bytes, "[Content_Types].xml");
        assert!(ct.contains("/xl/theme/theme1.xml"));

        let rels = read_zip_entry(&bytes, "xl/_rels/workbook.bin.rels");
        assert!(rels.contains("Target=\"theme/theme1.xml\""));
        assert!(rels.contains("/theme"));
    }

    #[test]
    fn round_trip_preserves_rgb_colors() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "test").unwrap();
        let mut style = Style::new().font_color(Color::Rgb {
            r: 79,
            g: 129,
            b: 189,
        });
        style.fill = FillStyle::Solid {
            color: Color::Rgb {
                r: 192,
                g: 80,
                b: 77,
            },
        };
        ws.set_cell_style_at(0, 0, &style).unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let s2 = ws2.cell_style_at(0, 0).unwrap();
        assert_eq!(
            s2.font.color,
            Color::Rgb {
                r: 79,
                g: 129,
                b: 189
            }
        );
        match s2.fill {
            FillStyle::Solid { color } => {
                assert_eq!(
                    color,
                    Color::Rgb {
                        r: 192,
                        g: 80,
                        b: 77
                    }
                );
            }
            ref other => panic!("expected Solid fill, got {:?}", other),
        }
    }

    #[test]
    fn rich_text_round_trip() {
        use duke_sheets_core::rich_text::{RichTextRun, RunFont};
        use duke_sheets_core::style::Color;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();

        let runs = vec![
            RichTextRun::plain("Hello "),
            RichTextRun::with_font(
                "World",
                RunFont {
                    bold: Some(true),
                    size: Some(14.0),
                    color: Some(Color::Rgb { r: 255, g: 0, b: 0 }),
                    name: Some("Arial".to_string()),
                    ..Default::default()
                },
            ),
        ];
        ws.set_cell_value_at(0, 0, CellValue::rich_text(runs))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let val = ws2.get_value_at(0, 0);

        match val {
            CellValue::RichText(ref rt_runs) => {
                assert_eq!(rt_runs.len(), 2);
                assert_eq!(rt_runs[0].text, "Hello ");
                assert_eq!(rt_runs[1].text, "World");
                let font = rt_runs[1].font.as_ref().expect("should have font");
                assert_eq!(font.bold, Some(true));
                assert_eq!(font.size, Some(14.0));
                assert_eq!(font.name, Some("Arial".to_string()));
                assert!(matches!(
                    font.color,
                    Some(Color::Rgb { r: 255, g: 0, b: 0 })
                ));
            }
            other => panic!("Expected RichText, got {:?}", other),
        }
    }

    #[test]
    fn rich_text_with_plain_text_coexist() {
        use duke_sheets_core::rich_text::{RichTextRun, RunFont};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();

        ws.set_cell_value_at(0, 0, "plain text").unwrap();

        let runs = vec![
            RichTextRun::with_font(
                "Bold",
                RunFont {
                    bold: Some(true),
                    ..Default::default()
                },
            ),
            RichTextRun::plain(" normal"),
        ];
        ws.set_cell_value_at(0, 1, CellValue::rich_text(runs))
            .unwrap();

        ws.set_cell_value_at(0, 2, "another plain").unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();

        assert_eq!(ws2.get_value_at(0, 0), CellValue::string("plain text"));
        assert!(ws2.get_value_at(0, 1).is_rich_text());
        assert_eq!(ws2.get_value_at(0, 2), CellValue::string("another plain"));

        if let CellValue::RichText(runs) = ws2.get_value_at(0, 1) {
            assert_eq!(runs.len(), 2);
            assert_eq!(runs[0].text, "Bold");
            assert_eq!(runs[0].font.as_ref().unwrap().bold, Some(true));
            assert_eq!(runs[1].text, " normal");
        }
    }

    #[test]
    fn drawing_round_trip() {
        use crate::drawing_bundle::DrawingBundle;

        let drawing_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp><xdr:txBody><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Hello</a:t></a:r></a:p></xdr:txBody></xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let mut bundle = DrawingBundle::new();
        bundle.push("xl/drawings/drawing1.xml".into(), drawing_xml.to_vec());
        let encoded = bundle.encode();

        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .raw_drawing_objects
            .push(encoded);

        let mut buf = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

        let mut archive = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
        let entry_names: Vec<String> = (0..archive.len())
            .map(|i| archive.by_index(i).unwrap().name().to_string())
            .collect();
        assert!(
            entry_names.iter().any(|n| n == "xl/drawings/drawing1.xml"),
            "drawing XML missing from ZIP: {:?}",
            entry_names
        );
        drop(archive);

        let ct = {
            let mut a = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
            let mut f = a.by_name("[Content_Types].xml").unwrap();
            let mut s = String::new();
            std::io::Read::read_to_string(&mut f, &mut s).unwrap();
            s
        };
        assert!(
            ct.contains("drawing+xml"),
            "content types missing drawing override: {}",
            ct
        );

        let rels = {
            let mut a = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
            let mut f = a.by_name("xl/worksheets/_rels/sheet1.bin.rels").unwrap();
            let mut s = String::new();
            std::io::Read::read_to_string(&mut f, &mut s).unwrap();
            s
        };
        assert!(
            rels.contains("/drawing"),
            "sheet rels missing drawing relationship: {}",
            rels
        );

        let wb2 = XlsbReader::read(Cursor::new(&buf)).unwrap();
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.raw_drawing_objects.len(), 1);

        let bundle2 = DrawingBundle::decode(&ws2.raw_drawing_objects[0]).unwrap();
        let (_, round_tripped) = bundle2
            .entries
            .iter()
            .find(|(p, _)| p.contains("/drawings/drawing"))
            .unwrap();
        assert_eq!(round_tripped, &drawing_xml.to_vec());
    }

    #[test]
    fn drawing_with_chart_round_trip() {
        use crate::drawing_bundle::DrawingBundle;

        let drawing_xml = b"<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"><xdr:twoCellAnchor/></xdr:wsDr>";
        let chart_xml = b"<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart/></c:chartSpace>";
        let drawing_rels = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>"#;

        let mut bundle = DrawingBundle::new();
        bundle.push("xl/drawings/drawing1.xml".into(), drawing_xml.to_vec());
        bundle.push(
            "xl/drawings/_rels/drawing1.xml.rels".into(),
            drawing_rels.to_vec(),
        );
        bundle.push("xl/charts/chart1.xml".into(), chart_xml.to_vec());

        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .raw_drawing_objects
            .push(bundle.encode());

        let mut buf = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

        let mut archive = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
        let names: Vec<String> = (0..archive.len())
            .map(|i| archive.by_index(i).unwrap().name().to_string())
            .collect();
        assert!(names.iter().any(|n| n == "xl/charts/chart1.xml"));
        assert!(names.iter().any(|n| n == "xl/drawings/drawing1.xml"));
        drop(archive);

        let ct = {
            let mut a = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
            let mut f = a.by_name("[Content_Types].xml").unwrap();
            let mut s = String::new();
            std::io::Read::read_to_string(&mut f, &mut s).unwrap();
            s
        };
        assert!(ct.contains("drawingml.chart+xml"));

        let wb2 = XlsbReader::read(Cursor::new(&buf)).unwrap();
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.raw_drawing_objects.len(), 1);
        let bundle2 = DrawingBundle::decode(&ws2.raw_drawing_objects[0]).unwrap();
        assert!(bundle2
            .entries
            .iter()
            .any(|(p, _)| p == "xl/charts/chart1.xml"));
        let (_, chart_bytes) = bundle2
            .entries
            .iter()
            .find(|(p, _)| p == "xl/charts/chart1.xml")
            .unwrap();
        assert_eq!(chart_bytes, &chart_xml.to_vec());
    }

    #[test]
    fn no_drawing_produces_no_drawing_parts() {
        let wb = Workbook::new();
        let mut buf = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

        let mut archive = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
        let names: Vec<String> = (0..archive.len())
            .map(|i| archive.by_index(i).unwrap().name().to_string())
            .collect();
        assert!(
            !names.iter().any(|n| n.contains("drawing")),
            "unexpected drawing entry: {:?}",
            names
        );
    }

    #[test]
    fn anchor_xml_fallback_round_trip() {
        let anchor = b"<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp/><xdr:clientData/></xdr:twoCellAnchor>";

        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .raw_drawing_objects
            .push(anchor.to_vec());

        let mut buf = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut buf)).unwrap();

        let drawing_xml = {
            let mut a = zip::ZipArchive::new(Cursor::new(&buf)).unwrap();
            let mut f = a.by_name("xl/drawings/drawing1.xml").unwrap();
            let mut s = String::new();
            std::io::Read::read_to_string(&mut f, &mut s).unwrap();
            s
        };
        assert!(drawing_xml.contains("<xdr:wsDr"));
        assert!(drawing_xml.contains("<xdr:twoCellAnchor>"));
    }

    #[test]
    fn rich_text_deduplication() {
        use duke_sheets_core::rich_text::{RichTextRun, RunFont};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();

        let make_runs = || {
            vec![RichTextRun::with_font(
                "Italic",
                RunFont {
                    italic: Some(true),
                    ..Default::default()
                },
            )]
        };

        ws.set_cell_value_at(0, 0, CellValue::rich_text(make_runs()))
            .unwrap();
        ws.set_cell_value_at(0, 1, CellValue::rich_text(make_runs()))
            .unwrap();

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();

        assert!(ws2.get_value_at(0, 0).is_rich_text());
        assert!(ws2.get_value_at(0, 1).is_rich_text());

        if let CellValue::RichText(runs) = ws2.get_value_at(0, 0) {
            assert_eq!(runs[0].text, "Italic");
            assert_eq!(runs[0].font.as_ref().unwrap().italic, Some(true));
        }
    }

    #[test]
    fn diagnostic_style_mapping_multi_sheet() {
        use crate::writer::styles::build_style_table;
        use duke_sheets_core::style::{Color, Style};

        let mut wb = Workbook::new();

        // Sheet 0: several styled cells
        let ws0 = wb.worksheet_mut(0).unwrap();
        ws0.set_name("Styled");
        for i in 0..10 {
            ws0.set_cell_value_at(i, 0, format!("row{}", i)).unwrap();
            let style = Style::new()
                .bold(i % 2 == 0)
                .font_size(10.0 + i as f64)
                .font_color(Color::rgb(i as u8 * 25, 0, 0));
            ws0.set_cell_style_at(i, 0, &style).unwrap();
        }

        // Sheets 1-4: plain data
        for s in 1..5 {
            wb.add_worksheet().unwrap();
            let ws = wb.worksheet_mut(s).unwrap();
            ws.set_name(&format!("Sheet{}", s));
            ws.set_cell_value_at(0, 0, format!("sheet{}", s)).unwrap();
        }

        let (_table, mapping, _) = build_style_table(&wb, &[]);
        let xf_count = mapping.xf_count();
        let max_mapped = mapping.max_mapped_xf();

        eprintln!("XF count: {}", xf_count);
        eprintln!("Max mapped XF: {}", max_mapped);
        for i in 0..5 {
            if let Some(map) = mapping.sheet_map(i) {
                let max_val = map.values().copied().max().unwrap_or(0);
                eprintln!(
                    "Sheet {} map: {} entries, max XF = {}",
                    i,
                    map.len(),
                    max_val
                );
            }
        }

        assert!(
            max_mapped < xf_count,
            "OUT OF RANGE: max mapped XF {} >= xf_count {}",
            max_mapped,
            xf_count
        );

        let wb2 = round_trip(&wb);
        assert_eq!(wb2.sheet_count(), 5);
        for i in 0..10 {
            let s = wb2.worksheet(0).unwrap().cell_style_at(i, 0);
            assert!(s.is_some(), "Missing style on sheet0 row {}", i);
        }
    }

    fn count_xfs_in_binary(data: &[u8]) -> u32 {
        use crate::biff12::records;
        let mut iter = crate::biff12::RecordIter::new(std::io::Cursor::new(data));
        let mut buf = Vec::new();
        let mut in_cell_xfs = false;
        let mut count = 0u32;
        loop {
            match iter.next_record(&mut buf) {
                Ok((typ, _len)) => {
                    if typ == records::BRT_BEGIN_CELL_XFS {
                        in_cell_xfs = true;
                        if buf.len() >= 4 {
                            count = u32::from_le_bytes([buf[0], buf[1], buf[2], buf[3]]);
                        }
                    }
                    if typ == records::BRT_END_CELL_XFS {
                        break;
                    }
                    if in_cell_xfs && typ == records::BRT_XF {}
                }
                Err(_) => break,
            }
        }
        count
    }

    fn max_style_ref_in_binary(data: &[u8]) -> u32 {
        use crate::biff12::records;
        let cell_records: &[u16] = &[
            records::BRT_CELL_BLANK,
            records::BRT_CELL_RK,
            records::BRT_CELL_ERROR,
            records::BRT_CELL_BOOL,
            records::BRT_CELL_REAL,
            records::BRT_CELL_ST,
            records::BRT_CELL_ISST,
            records::BRT_FMLA_STRING,
            records::BRT_FMLA_NUM,
            records::BRT_FMLA_BOOL,
            records::BRT_FMLA_ERROR,
        ];
        let mut iter = crate::biff12::RecordIter::new(std::io::Cursor::new(data));
        let mut buf = Vec::new();
        let mut max_ref = 0u32;
        loop {
            match iter.next_record(&mut buf) {
                Ok((typ, len)) => {
                    if cell_records.contains(&typ) && len >= 8 && buf.len() >= 7 {
                        let sr = u32::from_le_bytes([buf[4], buf[5], buf[6], 0]);
                        if sr > max_ref {
                            max_ref = sr;
                        }
                    }
                }
                Err(_) => break,
            }
        }
        max_ref
    }

    #[test]
    fn active_sheet_roundtrip() {
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("Sheet2").unwrap();
        wb.add_worksheet_with_name("Sheet3").unwrap();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "s1")
            .unwrap();
        wb.worksheet_mut(1)
            .unwrap()
            .set_cell_value_at(0, 0, "s2")
            .unwrap();
        wb.worksheet_mut(2)
            .unwrap()
            .set_cell_value_at(0, 0, "s3")
            .unwrap();
        wb.set_active_sheet(2).unwrap();

        let wb2 = round_trip(&wb);
        assert_eq!(wb2.active_sheet(), 2);
    }

    #[test]
    fn active_sheet_zero_roundtrip() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "data")
            .unwrap();

        let wb2 = round_trip(&wb);
        assert_eq!(wb2.active_sheet(), 0);
    }

    #[test]
    fn sheet_visibility_roundtrip() {
        use duke_sheets_core::worksheet::SheetVisibility;

        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("Hidden").unwrap();
        wb.add_worksheet_with_name("VeryHidden").unwrap();

        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "visible")
            .unwrap();
        wb.worksheet_mut(1)
            .unwrap()
            .set_visibility(SheetVisibility::Hidden);
        wb.worksheet_mut(1)
            .unwrap()
            .set_cell_value_at(0, 0, "hidden")
            .unwrap();
        wb.worksheet_mut(2)
            .unwrap()
            .set_visibility(SheetVisibility::VeryHidden);
        wb.worksheet_mut(2)
            .unwrap()
            .set_cell_value_at(0, 0, "very hidden")
            .unwrap();

        let wb2 = round_trip(&wb);
        assert_eq!(wb2.sheet_count(), 3);
        assert_eq!(
            wb2.worksheet(0).unwrap().visibility(),
            SheetVisibility::Visible
        );
        assert_eq!(
            wb2.worksheet(1).unwrap().visibility(),
            SheetVisibility::Hidden
        );
        assert_eq!(
            wb2.worksheet(2).unwrap().visibility(),
            SheetVisibility::VeryHidden
        );
        assert_eq!(
            wb2.worksheet(0).unwrap().get_value_at(0, 0),
            CellValue::string("visible")
        );
        assert_eq!(
            wb2.worksheet(1).unwrap().get_value_at(0, 0),
            CellValue::string("hidden")
        );
        assert_eq!(
            wb2.worksheet(2).unwrap().get_value_at(0, 0),
            CellValue::string("very hidden")
        );
    }

    #[test]
    fn table_roundtrip() {
        use duke_sheets_core::table::{Table, TableColumn, TableStyleInfo};
        use duke_sheets_core::CellRange;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "Name").unwrap();
        ws.set_cell_value_at(0, 1, "Value").unwrap();
        ws.set_cell_value_at(0, 2, "Score").unwrap();
        ws.set_cell_value_at(1, 0, "Alice").unwrap();
        ws.set_cell_value_at(1, 1, 100.0).unwrap();
        ws.set_cell_value_at(1, 2, 95.0).unwrap();
        ws.set_cell_value_at(2, 0, "Bob").unwrap();
        ws.set_cell_value_at(2, 1, 200.0).unwrap();
        ws.set_cell_value_at(2, 2, 88.0).unwrap();

        let table = Table {
            id: 1,
            name: "SalesTable".to_string(),
            display_name: "SalesTable".to_string(),
            reference: CellRange::parse("A1:C3").unwrap(),
            columns: vec![
                TableColumn {
                    id: 1,
                    name: "Name".to_string(),
                    totals_row_function: None,
                    totals_row_formula: None,
                    totals_row_label: None,
                    calculated_column_formula: None,
                },
                TableColumn {
                    id: 2,
                    name: "Value".to_string(),
                    totals_row_function: None,
                    totals_row_formula: None,
                    totals_row_label: None,
                    calculated_column_formula: None,
                },
                TableColumn {
                    id: 3,
                    name: "Score".to_string(),
                    totals_row_function: None,
                    totals_row_formula: None,
                    totals_row_label: None,
                    calculated_column_formula: None,
                },
            ],
            style_info: Some(TableStyleInfo {
                name: Some("TableStyleMedium2".to_string()),
                show_first_column: false,
                show_last_column: false,
                show_row_stripes: true,
                show_column_stripes: false,
            }),
            header_row_count: 1,
            totals_row_count: 0,
            totals_row_shown: true,
        };
        ws.add_table(table);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let tables = ws2.tables();
        assert_eq!(tables.len(), 1, "should have 1 table");
        let t = &tables[0];
        assert_eq!(t.name, "SalesTable");
        assert_eq!(t.display_name, "SalesTable");
        assert_eq!(t.columns.len(), 3);
        assert_eq!(t.columns[0].name, "Name");
        assert_eq!(t.columns[1].name, "Value");
        assert_eq!(t.columns[2].name, "Score");
        assert!(t.style_info.is_some());
        let si = t.style_info.as_ref().unwrap();
        assert_eq!(si.name.as_deref(), Some("TableStyleMedium2"));
        assert!(si.show_row_stripes);
        assert!(!si.show_column_stripes);
    }

    #[test]
    fn named_range_roundtrip() {
        use duke_sheets_core::named_range::{NameScope, NamedRange};
        let mut wb = Workbook::new();
        wb.add_worksheet_with_name("Data").unwrap();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "header")
            .unwrap();
        wb.worksheet_mut(1)
            .unwrap()
            .set_cell_value_at(0, 0, 100.0)
            .unwrap();

        wb.named_ranges_mut()
            .define_or_update(NamedRange::workbook_scope("SalesData", "Data!$A$1:$D$10"));
        wb.named_ranges_mut()
            .define_or_update(NamedRange::sheet_scope("LocalRate", "Sheet1!$B$1", 0));

        let wb2 = round_trip(&wb);
        let nr = wb2.named_ranges();
        assert!(nr.get("SalesData", 0).is_some(), "SalesData should exist");
        let sd = nr.get("SalesData", 0).unwrap();
        assert_eq!(sd.scope, NameScope::Workbook);
        assert!(
            sd.refers_to.contains("Data!"),
            "SalesData refers_to should reference Data sheet"
        );

        assert!(
            nr.get("LocalRate", 0).is_some(),
            "LocalRate should exist for sheet 0"
        );
        let lr = nr.get("LocalRate", 0).unwrap();
        assert!(matches!(lr.scope, NameScope::Sheet(0)));
    }

    #[test]
    fn tab_color_rgb_roundtrip() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_tab_color(Some(Color::rgb(255, 0, 0)));

        let wb2 = round_trip(&wb);
        assert_eq!(
            wb2.worksheet(0).unwrap().tab_color(),
            Some(Color::rgb(255, 0, 0)),
        );
    }

    #[test]
    fn tab_color_none_roundtrip() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0).unwrap().set_tab_color(None);

        let wb2 = round_trip(&wb);
        assert_eq!(wb2.worksheet(0).unwrap().tab_color(), None);
    }

    #[test]
    fn cf_color_scale_roundtrip() {
        use duke_sheets_core::conditional_format::{
            CfColorValue, CfRuleType, CfValueType, ConditionalFormatRule,
        };
        use duke_sheets_core::CellRange;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 10.0).unwrap();
        ws.set_cell_value_at(1, 0, 50.0).unwrap();
        ws.set_cell_value_at(2, 0, 90.0).unwrap();

        let rule = ConditionalFormatRule {
            rule_type: CfRuleType::ColorScale {
                colors: vec![
                    CfColorValue::new(CfValueType::Min, None, Color::rgb(255, 0, 0)),
                    CfColorValue::new(CfValueType::Max, None, Color::rgb(0, 255, 0)),
                ],
            },
            ranges: vec![CellRange::parse("A1:A3").unwrap()],
            priority: 1,
            stop_if_true: false,
            format: None,
            dxf_id: None,
        };
        ws.add_conditional_format(rule);

        let mut buf = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut buf)).unwrap();
        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let rules = ws2.conditional_formats();
        assert_eq!(rules.len(), 1, "should have 1 CF rule");
        match &rules[0].rule_type {
            CfRuleType::ColorScale { colors } => {
                assert_eq!(colors.len(), 2);
                assert_eq!(colors[0].value_type, CfValueType::Min);
                assert_eq!(colors[1].value_type, CfValueType::Max);
            }
            other => panic!("expected ColorScale, got {:?}", other),
        }
    }

    #[test]
    fn cf_data_bar_roundtrip() {
        use duke_sheets_core::conditional_format::{
            CfRuleType, CfValue, CfValueType, ConditionalFormatRule,
        };
        use duke_sheets_core::CellRange;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 25.0).unwrap();
        ws.set_cell_value_at(1, 0, 75.0).unwrap();

        let rule = ConditionalFormatRule {
            rule_type: CfRuleType::DataBar {
                min_value: CfValue::min(),
                max_value: CfValue::max(),
                color: Color::rgb(99, 142, 198),
                show_value: true,
                gradient: true,
                border_color: None,
                negative_color: None,
            },
            ranges: vec![CellRange::parse("A1:A2").unwrap()],
            priority: 1,
            stop_if_true: false,
            format: None,
            dxf_id: None,
        };
        ws.add_conditional_format(rule);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let rules = ws2.conditional_formats();
        assert_eq!(rules.len(), 1);
        match &rules[0].rule_type {
            CfRuleType::DataBar {
                min_value,
                max_value,
                show_value,
                ..
            } => {
                assert_eq!(min_value.value_type, CfValueType::Min);
                assert_eq!(max_value.value_type, CfValueType::Max);
                assert!(*show_value);
            }
            other => panic!("expected DataBar, got {:?}", other),
        }
    }

    #[test]
    fn cf_icon_set_roundtrip() {
        use duke_sheets_core::conditional_format::{
            CfRuleType, CfValue, CfValueType, ConditionalFormatRule, IconSetStyle,
        };
        use duke_sheets_core::CellRange;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 10.0).unwrap();
        ws.set_cell_value_at(1, 0, 50.0).unwrap();
        ws.set_cell_value_at(2, 0, 90.0).unwrap();

        let rule = ConditionalFormatRule {
            rule_type: CfRuleType::IconSet {
                icon_style: IconSetStyle::Arrows3,
                values: vec![
                    CfValue::new(CfValueType::Percent, Some("0".to_string())),
                    CfValue::new(CfValueType::Percent, Some("33".to_string())),
                    CfValue::new(CfValueType::Percent, Some("67".to_string())),
                ],
                reverse: false,
                show_value: true,
            },
            ranges: vec![CellRange::parse("A1:A3").unwrap()],
            priority: 1,
            stop_if_true: false,
            format: None,
            dxf_id: None,
        };
        ws.add_conditional_format(rule);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let rules = ws2.conditional_formats();
        assert_eq!(rules.len(), 1);
        match &rules[0].rule_type {
            CfRuleType::IconSet {
                icon_style,
                values,
                reverse,
                show_value,
            } => {
                assert_eq!(*icon_style, IconSetStyle::Arrows3);
                assert_eq!(values.len(), 3);
                assert!(!*reverse);
                assert!(*show_value);
            }
            other => panic!("expected IconSet, got {:?}", other),
        }
    }

    #[test]
    fn cf_with_dxf_style_roundtrip() {
        use duke_sheets_core::conditional_format::{CfOperator, CfRuleType, ConditionalFormatRule};
        use duke_sheets_core::CellRange;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 100.0).unwrap();

        let format = Style::new().fill_color(Color::rgb(0, 255, 0)).bold(true);
        let rule = ConditionalFormatRule {
            rule_type: CfRuleType::CellIs {
                operator: CfOperator::GreaterThan,
                formula1: "50".to_string(),
                formula2: None,
            },
            ranges: vec![CellRange::parse("A1").unwrap()],
            priority: 1,
            stop_if_true: false,
            format: Some(format),
            dxf_id: None,
        };
        ws.add_conditional_format(rule);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let rules = ws2.conditional_formats();
        assert_eq!(rules.len(), 1);
        match &rules[0].rule_type {
            CfRuleType::CellIs {
                operator, formula1, ..
            } => {
                assert_eq!(*operator, CfOperator::GreaterThan);
                assert_eq!(formula1, "50");
            }
            other => panic!("expected CellIs, got {:?}", other),
        }
    }

    #[test]
    fn cf_three_color_scale_roundtrip() {
        use duke_sheets_core::conditional_format::{
            CfColorValue, CfRuleType, CfValueType, ConditionalFormatRule,
        };
        use duke_sheets_core::CellRange;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, 10.0).unwrap();
        ws.set_cell_value_at(1, 0, 50.0).unwrap();
        ws.set_cell_value_at(2, 0, 90.0).unwrap();

        let rule = ConditionalFormatRule {
            rule_type: CfRuleType::ColorScale {
                colors: vec![
                    CfColorValue::new(CfValueType::Min, None, Color::rgb(255, 0, 0)),
                    CfColorValue::new(
                        CfValueType::Percentile,
                        Some("50".to_string()),
                        Color::rgb(255, 255, 0),
                    ),
                    CfColorValue::new(CfValueType::Max, None, Color::rgb(0, 255, 0)),
                ],
            },
            ranges: vec![CellRange::parse("A1:A3").unwrap()],
            priority: 1,
            stop_if_true: false,
            format: None,
            dxf_id: None,
        };
        ws.add_conditional_format(rule);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let rules = ws2.conditional_formats();
        assert_eq!(rules.len(), 1);
        match &rules[0].rule_type {
            CfRuleType::ColorScale { colors } => {
                assert_eq!(colors.len(), 3);
                assert_eq!(colors[0].value_type, CfValueType::Min);
                assert_eq!(colors[1].value_type, CfValueType::Percentile);
                assert_eq!(colors[2].value_type, CfValueType::Max);
                assert_eq!(colors[0].color, Color::rgb(255, 0, 0));
                assert_eq!(colors[1].color, Color::rgb(255, 255, 0));
                assert_eq!(colors[2].color, Color::rgb(0, 255, 0));
            }
            other => panic!("expected ColorScale, got {:?}", other),
        }
    }

    #[test]
    fn zoom_scale_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "zoom").unwrap();
        ws.set_zoom_scale(Some(150));

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.zoom_scale(), Some(150));
    }

    #[test]
    fn row_outline_level_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "outline").unwrap();
        ws.set_row_outline_level(1, 1);
        ws.set_row_outline_level(2, 1);
        ws.set_row_outline_level(3, 1);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.row_outline_level(1), 1);
        assert_eq!(ws2.row_outline_level(2), 1);
        assert_eq!(ws2.row_outline_level(3), 1);
        assert_eq!(ws2.row_outline_level(0), 0);
    }

    #[test]
    fn column_outline_level_roundtrip() {
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "outline").unwrap();
        ws.set_column_outline_level(1, 1);
        ws.set_column_outline_level(2, 1);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        assert_eq!(ws2.column_outline_level(1), 1);
        assert_eq!(ws2.column_outline_level(2), 1);
        assert_eq!(ws2.column_outline_level(0), 0);
    }

    #[test]
    fn active_cell_selection_roundtrip() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "data")
            .unwrap();
        wb.worksheet_mut(0).unwrap().set_selection_active_cell(5, 3);

        let wb2 = round_trip(&wb);
        assert_eq!(
            wb2.worksheet(0).unwrap().selection_active_cell(),
            Some((5, 3))
        );
    }

    #[test]
    fn cf_time_period_roundtrip() {
        use duke_sheets_core::conditional_format::{CfRuleType, ConditionalFormatRule, TimePeriod};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "date").unwrap();
        let rule = ConditionalFormatRule::new(CfRuleType::TimePeriod {
            period: TimePeriod::ThisMonth,
        })
        .with_range(duke_sheets_core::CellRange::parse("A1:A10").unwrap())
        .with_priority(1);
        ws.add_conditional_format(rule);

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let cfs = ws2.conditional_formats();
        assert_eq!(cfs.len(), 1);
        match &cfs[0].rule_type {
            CfRuleType::TimePeriod { period } => {
                assert_eq!(*period, TimePeriod::ThisMonth);
            }
            other => panic!("expected TimePeriod, got {:?}", other),
        }
        assert_eq!(cfs[0].ranges[0].to_string(), "A1:A10");
    }

    #[test]
    fn tab_selected_roundtrip() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "s1")
            .unwrap();
        wb.add_worksheet().unwrap();
        wb.worksheet_mut(1)
            .unwrap()
            .set_cell_value_at(0, 0, "s2")
            .unwrap();
        wb.worksheet_mut(1).unwrap().set_selected(true);

        let wb2 = round_trip(&wb);
        assert!(wb2.worksheet(1).unwrap().is_selected());
    }

    #[test]
    fn split_panes_roundtrip() {
        use duke_sheets_core::worksheet::SplitPanes;

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        ws.set_cell_value_at(0, 0, "split").unwrap();
        ws.set_split_panes(Some(SplitPanes {
            x_split: 2000.0,
            y_split: 3000.0,
            top_left: Some((5, 3)),
            active_pane: Some("bottomRight".to_string()),
        }));

        let wb2 = round_trip(&wb);
        let ws2 = wb2.worksheet(0).unwrap();
        let sp = ws2.split_panes().expect("split panes should exist");
        assert!((sp.x_split - 2000.0).abs() < 0.1);
        assert!((sp.y_split - 3000.0).abs() < 0.1);
        assert_eq!(sp.top_left, Some((5, 3)));
        assert_eq!(sp.active_pane.as_deref(), Some("bottomRight"));
        assert!(ws2.freeze_panes().is_none());
    }

    #[test]
    fn sheet_protection_roundtrip() {
        use duke_sheets_core::worksheet::SheetProtection;
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "data")
            .unwrap();
        wb.worksheet_mut(0)
            .unwrap()
            .set_protection(Some(SheetProtection {
                protected: true,
                password_hash: None,
                select_locked_cells: true,
                select_unlocked_cells: true,
                format_cells: false,
                format_columns: false,
                format_rows: false,
                insert_columns: false,
                insert_rows: false,
                insert_hyperlinks: false,
                delete_columns: false,
                delete_rows: false,
                sort: false,
                auto_filter: false,
                pivot_tables: false,
            }));
        let wb2 = round_trip(&wb);
        let prot = wb2.worksheet(0).unwrap().protection().expect("protection");
        assert!(prot.protected);
        assert!(prot.select_locked_cells);
        assert!(!prot.format_cells);
    }

    #[test]
    fn page_breaks_roundtrip() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "data")
            .unwrap();
        wb.worksheet_mut(0).unwrap().add_row_break(5);
        let wb2 = round_trip(&wb);
        let breaks = wb2.worksheet(0).unwrap().row_breaks();
        assert_eq!(breaks.len(), 1);
        assert_eq!(breaks[0].id, 5);
    }

    #[test]
    fn print_area_roundtrip() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_cell_value_at(0, 0, "data")
            .unwrap();
        wb.worksheet_mut(0)
            .unwrap()
            .set_print_area(duke_sheets_core::CellRange::parse("A1:D10").unwrap());
        let wb2 = round_trip(&wb);
        assert!(wb2.worksheet(0).unwrap().print_area().is_some());
    }

    #[test]
    fn dynamic_filter_in_process_roundtrip() {
        use duke_sheets_core::auto_filter::{
            AutoFilter, ColumnFilter, DynamicFilter, DynamicFilterType, FilterColumn,
        };
        use duke_sheets_core::{CellAddress, CellRange};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        let mut af = AutoFilter::new(CellRange::new(
            CellAddress::parse("A1").unwrap(),
            CellAddress::parse("A5").unwrap(),
        ));
        af.filter_columns.push(FilterColumn::new(
            0,
            ColumnFilter::Dynamic(DynamicFilter {
                filter_type: DynamicFilterType::Today,
                val: Some(45000.5),
                max_val: Some(45001.0),
            }),
        ));
        ws.set_auto_filter(Some(af));

        let wb2 = round_trip(&wb);
        let af = wb2.worksheet(0).unwrap().auto_filter().unwrap().clone();
        let col0 = af.filter_columns.iter().find(|fc| fc.col_id == 0).unwrap();
        match &col0.filter {
            ColumnFilter::Dynamic(d) => {
                assert_eq!(d.filter_type, DynamicFilterType::Today);
                assert_eq!(d.val, Some(45000.5));
                assert_eq!(d.max_val, Some(45001.0));
            }
            other => panic!("expected Dynamic filter, got {other:?}"),
        }
    }

    #[test]
    fn color_filter_in_process_roundtrip() {
        use duke_sheets_core::auto_filter::{AutoFilter, ColorFilter, ColumnFilter, FilterColumn};
        use duke_sheets_core::{CellAddress, CellRange};

        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        let mut af = AutoFilter::new(CellRange::new(
            CellAddress::parse("A1").unwrap(),
            CellAddress::parse("A3").unwrap(),
        ));
        af.filter_columns.push(FilterColumn::new(
            0,
            ColumnFilter::Color(ColorFilter {
                dxf_id: Some(2),
                cell_color: false,
            }),
        ));
        ws.set_auto_filter(Some(af));

        let wb2 = round_trip(&wb);
        let af = wb2.worksheet(0).unwrap().auto_filter().unwrap().clone();
        let col0 = af.filter_columns.iter().find(|fc| fc.col_id == 0).unwrap();
        match &col0.filter {
            ColumnFilter::Color(c) => {
                assert_eq!(c.dxf_id, Some(2));
                assert!(!c.cell_color, "cell_color flag drifted");
            }
            other => panic!("expected Color filter, got {other:?}"),
        }
    }

    /// Raw (record_type, payload) pairs of `xl/worksheets/sheet1.bin`.
    fn sheet1_records(wb: &Workbook) -> Vec<(u16, Vec<u8>)> {
        let bytes = write_xlsb_bytes(wb);
        let bin = read_zip_entry_bytes(&bytes, "xl/worksheets/sheet1.bin");
        let mut iter = crate::biff12::RecordIter::new(Cursor::new(bin));
        let mut out = Vec::new();
        let mut buf = Vec::new();
        while let Ok((typ, len)) = iter.next_record(&mut buf) {
            out.push((typ, buf[..len].to_vec()));
        }
        out
    }

    fn read_zip_entry_bytes(data: &[u8], name: &str) -> Vec<u8> {
        let cursor = Cursor::new(data);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut file = archive.by_name(name).unwrap();
        let mut buf = Vec::new();
        std::io::Read::read_to_end(&mut file, &mut buf).unwrap();
        buf
    }

    #[test]
    fn color_filter_emits_brt_color_filter_record_id() {
        use duke_sheets_core::auto_filter::{AutoFilter, ColorFilter, ColumnFilter, FilterColumn};
        use duke_sheets_core::{CellAddress, CellRange};

        // [MS-XLSB] record enumeration: 168 (0x00A8) = BrtColorFilter
        // (§2.4.339), 169 (0x00A9) = BrtIconFilter. A struct-level
        // round-trip cannot catch a swapped ID (both payloads are two
        // u32s), so pin the raw record id and payload here.
        let mut wb = Workbook::new();
        let ws = wb.worksheet_mut(0).unwrap();
        let mut af = AutoFilter::new(CellRange::new(
            CellAddress::parse("A1").unwrap(),
            CellAddress::parse("A3").unwrap(),
        ));
        af.filter_columns.push(FilterColumn::new(
            0,
            ColumnFilter::Color(ColorFilter {
                dxf_id: Some(2),
                cell_color: false,
            }),
        ));
        ws.set_auto_filter(Some(af));

        let recs = sheet1_records(&wb);
        let color: Vec<_> = recs.iter().filter(|(t, _)| *t == 0x00A8).collect();
        assert_eq!(
            color.len(),
            1,
            "expected exactly one BrtColorFilter (0x00A8) record"
        );
        let mut expected = 2u32.to_le_bytes().to_vec();
        expected.extend_from_slice(&0u32.to_le_bytes());
        assert_eq!(color[0].1, expected, "dxfid + fCellColor payload");
        assert!(
            !recs.iter().any(|(t, _)| *t == 0x00A9),
            "0x00A9 is BrtIconFilter; a color filter must not emit it"
        );
    }

    #[test]
    fn named_range_comment_roundtrip() {
        use duke_sheets_core::named_range::{NameScope, NamedRange};

        let mut wb = Workbook::new();
        let nr =
            NamedRange::new("MyTax", "0.07", NameScope::Workbook).with_comment("Sales tax rate");
        wb.named_ranges_mut().define_or_update(nr);

        let wb2 = round_trip(&wb);
        let got = wb2
            .named_ranges()
            .iter()
            .find(|n| n.name == "MyTax")
            .expect("named range lost");
        assert_eq!(
            got.comment.as_deref(),
            Some("Sales tax rate"),
            "named range comment lost: {:?}",
            got.comment
        );
    }
}
