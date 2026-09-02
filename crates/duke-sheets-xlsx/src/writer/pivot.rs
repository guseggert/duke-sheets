use std::io::{Seek, Write};

use quick_xml::events::{BytesEnd, BytesStart, Event};

use crate::styles::XlsxStyleTable;
use duke_sheets_core::{
    CellAddress, CellRange, PivotAggregate, PivotCalculatedItem, PivotDateGroupUnit,
    PivotDatePeriod, PivotField, PivotFieldRef, PivotFilter, PivotFilterOperator, PivotGrouping,
    PivotLayoutKind, PivotShowAs, PivotSort, PivotSourceRange, PivotSubtotal, PivotTable,
    PivotValue, PivotValuesAxis, Workbook, WorkbookConnection,
};
use duke_sheets_pivot::plan::{
    pivot_measure_matches_target, FormatPivotCache, FormatPivotCacheField, FormatPivotGrouping,
    FormatPivotPlan, FormatPivotSource, FormatPivotTable,
};

use super::{write_xml_part, XlsxError, XlsxResult, XmlWriter, NS_DOC_RELS, NS_SPREADSHEET};

const NS_SPREADSHEET_X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const EXT_URI_X14_DATA_FIELD: &str = "{2946ED86-A175-432a-8AC1-64E0C546D7DE}";
pub(super) const RT_EXTERNAL_LINK_PATH: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";

type CacheField = FormatPivotCacheField;

fn xlsx_cache_fields(cache: &FormatPivotCache) -> Vec<CacheField> {
    let mut fields = cache.fields.clone();
    for field in &mut fields {
        if let Some(grouping) = &field.grouping {
            if matches!(grouping.definition, PivotGrouping::Manual { .. }) {
                field.shared_items = grouping.source_items.clone();
                field.item_ids = grouping.source_item_ids.clone();
            }
        }
    }
    for field in &cache.fields {
        let Some(grouping) = &field.grouping else {
            continue;
        };
        if !matches!(grouping.definition, PivotGrouping::Manual { .. }) {
            continue;
        }
        let mut suffix = 2;
        let name = loop {
            let candidate = format!("{}{suffix}", field.name);
            if fields
                .iter()
                .all(|existing| !existing.name.eq_ignore_ascii_case(&candidate))
            {
                break candidate;
            }
            suffix += 1;
        };
        fields.push(CacheField {
            name,
            formula: None,
            database_field: false,
            shared_items: Vec::new(),
            item_ids: Vec::new(),
            grouping: Some(grouping.clone()),
        });
    }
    fields
}

fn xlsx_table_fields(cache: &FormatPivotCache, pivot: &PivotTable) -> Vec<CacheField> {
    let mut fields = xlsx_cache_fields(cache);
    let referenced = pivot
        .rows
        .iter()
        .chain(&pivot.columns)
        .chain(&pivot.page_fields)
        .map(|field| field.field.name.as_str())
        .chain(
            pivot
                .measures
                .iter()
                .map(|measure| measure.field.name.as_str()),
        )
        .collect::<Vec<_>>();
    for (alias, target) in &cache.field_aliases {
        if !referenced
            .iter()
            .any(|name| name.eq_ignore_ascii_case(alias))
        {
            continue;
        }
        if let Some(field) = fields
            .iter_mut()
            .find(|field| field.name.eq_ignore_ascii_case(target))
        {
            field.name.clone_from(alias);
        }
    }
    fields
}
pub(super) fn workbook_cache_rid(cache_num: usize) -> String {
    format!("rIdPivotCache{}", cache_num)
}

pub(super) fn avoid_preserved_part_collisions(plan: &mut FormatPivotPlan, workbook: &Workbook) {
    let claimed = workbook
        .workbook_extension_parts()
        .iter()
        .map(|part| part.path.trim_start_matches('/').to_ascii_lowercase())
        .collect::<std::collections::HashSet<_>>();
    let mut cache_numbers = std::collections::HashMap::new();
    let mut next_cache = 1;
    for cache in &mut plan.caches {
        while [
            format!("xl/pivotcache/pivotcachedefinition{next_cache}.xml"),
            format!("xl/pivotcache/pivotcacherecords{next_cache}.xml"),
            format!("xl/pivotcache/_rels/pivotcachedefinition{next_cache}.xml.rels"),
        ]
        .iter()
        .any(|path| claimed.contains(path))
        {
            next_cache += 1;
        }
        cache_numbers.insert(cache.cache_num, next_cache);
        cache.cache_num = next_cache;
        next_cache += 1;
    }
    let mut next_table = 1;
    for table in &mut plan.tables {
        table.cache_num = cache_numbers[&table.cache_num];
        while [
            format!("xl/pivottables/pivottable{next_table}.xml"),
            format!("xl/pivottables/_rels/pivottable{next_table}.xml.rels"),
        ]
        .iter()
        .any(|path| claimed.contains(path))
        {
            next_table += 1;
        }
        table.table_num = next_table;
        next_table += 1;
    }
}

fn filter_field_ref(filter: &PivotFilter) -> Option<&PivotFieldRef> {
    match filter {
        PivotFilter::FieldItems { field, .. }
        | PivotFilter::Label { field, .. }
        | PivotFilter::LabelBetween { field, .. }
        | PivotFilter::Date { field, .. }
        | PivotFilter::DateBetween { field, .. }
        | PivotFilter::DatePeriod { field, .. }
        | PivotFilter::Value { field, .. }
        | PivotFilter::ValueBetween { field, .. }
        | PivotFilter::TopN { field, .. } => Some(field),
        PivotFilter::Unsupported { .. } => None,
    }
}

fn formula_for_cache_attr(formula: &str) -> String {
    formula.trim().trim_start_matches('=').to_string()
}

pub(super) fn write_pivot_table_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    workbook: &Workbook,
    part: &FormatPivotTable,
    cache_part: &FormatPivotCache,
    style_table: &XlsxStyleTable,
) -> XlsxResult<()> {
    let sheet = workbook
        .worksheet(part.sheet_index)
        .ok_or_else(|| XlsxError::InvalidFormat("pivot table sheet not found".into()))?;
    let pivot = sheet
        .pivot_tables()
        .get(part.pivot_index)
        .ok_or_else(|| XlsxError::InvalidFormat("pivot table not found".into()))?;
    let fields = xlsx_table_fields(cache_part, pivot);

    let path = format!("xl/pivotTables/pivotTable{}.xml", part.table_num);
    write_xml_part(zip, &path, |w| {
        let cache_id = part.cache_num.to_string();
        let row_grand = bool_attr(pivot.layout.show_row_grand_totals);
        let col_grand = bool_attr(pivot.layout.show_column_grand_totals);
        let preserve_formatting = bool_attr(pivot.refresh_policy.preserve_formatting);
        let show_headers = bool_attr(pivot.layout.show_field_headers);
        let show_drill = bool_attr(pivot.layout.show_expand_collapse);
        let print_drill = bool_attr(pivot.layout.print_drill_indicators);
        let item_print_titles = bool_attr(pivot.layout.item_print_titles);
        let field_print_titles = bool_attr(pivot.layout.field_print_titles);
        let compact = bool_attr(matches!(pivot.layout.kind, PivotLayoutKind::Compact));
        let outline = bool_attr(matches!(pivot.layout.kind, PivotLayoutKind::Outline));
        let page_wrap = pivot.layout.page_wrap.to_string();
        let page_over_then_down = bool_attr(pivot.layout.page_over_then_down);
        let merge_item = bool_attr(pivot.layout.merge_item_labels);
        let data_caption = if pivot.layout.data_caption.trim().is_empty() {
            "Values"
        } else {
            pivot.layout.data_caption.as_str()
        };
        let show_error = bool_attr(pivot.layout.show_error);
        let show_missing = bool_attr(pivot.layout.show_missing);

        let mut tag = BytesStart::new("pivotTableDefinition");
        tag.push_attribute(("xmlns", NS_SPREADSHEET));
        tag.push_attribute(("name", pivot.name.as_str()));
        tag.push_attribute(("cacheId", cache_id.as_str()));
        tag.push_attribute(("dataCaption", data_caption));
        push_bool_attr_if(
            &mut tag,
            "dataOnRows",
            matches!(pivot.layout.values_axis, PivotValuesAxis::Rows),
            false,
        );
        let data_position = pivot
            .layout
            .values_axis_position
            .map(|position| position.to_string());
        if let Some(data_position) = data_position.as_deref() {
            tag.push_attribute(("dataPosition", data_position));
        }
        if let Some(caption) = &pivot.layout.grand_total_caption {
            tag.push_attribute(("grandTotalCaption", caption.as_str()));
        }
        if let Some(caption) = &pivot.layout.error_caption {
            tag.push_attribute(("errorCaption", caption.as_str()));
        }
        if pivot.layout.show_error {
            tag.push_attribute(("showError", show_error));
        }
        if let Some(caption) = &pivot.layout.missing_caption {
            tag.push_attribute(("missingCaption", caption.as_str()));
        }
        if !pivot.layout.show_missing {
            tag.push_attribute(("showMissing", show_missing));
        }
        tag.push_attribute(("updatedVersion", "8"));
        tag.push_attribute(("minRefreshableVersion", "3"));
        push_bool_attr_if(
            &mut tag,
            "asteriskTotals",
            pivot.layout.asterisk_totals,
            false,
        );
        push_bool_attr_if(&mut tag, "showItems", pivot.layout.show_items, true);
        push_bool_attr_if(&mut tag, "editData", pivot.layout.edit_data, false);
        push_bool_attr_if(
            &mut tag,
            "disableFieldList",
            pivot.layout.disable_field_list,
            false,
        );
        push_bool_attr_if(
            &mut tag,
            "showCalcMbrs",
            pivot.layout.show_calculated_members,
            true,
        );
        push_bool_attr_if(&mut tag, "visualTotals", pivot.layout.visual_totals, true);
        push_bool_attr_if(
            &mut tag,
            "showMultipleLabel",
            pivot.layout.show_multiple_label,
            true,
        );
        push_bool_attr_if(
            &mut tag,
            "showDataDropDown",
            pivot.layout.show_data_drop_down,
            true,
        );
        tag.push_attribute(("rowGrandTotals", row_grand));
        tag.push_attribute(("colGrandTotals", col_grand));
        tag.push_attribute(("preserveFormatting", preserve_formatting));
        tag.push_attribute(("showHeaders", show_headers));
        tag.push_attribute(("showDrill", show_drill));
        tag.push_attribute(("printDrill", print_drill));
        push_bool_attr_if(
            &mut tag,
            "showMemberPropertyTips",
            pivot.layout.show_member_property_tips,
            true,
        );
        push_bool_attr_if(&mut tag, "showDataTips", pivot.layout.show_data_tips, true);
        push_bool_attr_if(&mut tag, "enableWizard", pivot.layout.enable_wizard, true);
        push_bool_attr_if(&mut tag, "enableDrill", pivot.layout.enable_drill, true);
        push_bool_attr_if(
            &mut tag,
            "enableFieldProperties",
            pivot.layout.enable_field_properties,
            true,
        );
        tag.push_attribute(("itemPrintTitles", item_print_titles));
        tag.push_attribute(("fieldPrintTitles", field_print_titles));
        tag.push_attribute(("pageWrap", page_wrap.as_str()));
        tag.push_attribute(("pageOverThenDown", page_over_then_down));
        push_bool_attr_if(
            &mut tag,
            "subtotalHiddenItems",
            pivot.layout.subtotal_hidden_items,
            false,
        );
        tag.push_attribute(("mergeItem", merge_item));
        push_bool_attr_if(
            &mut tag,
            "showDropZones",
            pivot.layout.show_drop_zones,
            true,
        );
        let indent = pivot.layout.indent.to_string();
        if pivot.layout.indent != 1 {
            tag.push_attribute(("indent", indent.as_str()));
        }
        push_bool_attr_if(
            &mut tag,
            "showEmptyRow",
            pivot.layout.show_empty_rows,
            false,
        );
        push_bool_attr_if(
            &mut tag,
            "showEmptyCol",
            pivot.layout.show_empty_columns,
            false,
        );
        tag.push_attribute(("compact", compact));
        tag.push_attribute(("outline", outline));
        w.write_event(Event::Start(tag))?;

        write_location(w, pivot, &fields, part)?;
        write_pivot_fields(w, pivot, &fields)?;
        write_axis_fields(
            w,
            "rowFields",
            &pivot.rows,
            &fields,
            values_field_on_axis(pivot, PivotValuesAxis::Rows),
            pivot.layout.values_axis_position,
        )?;
        write_axis_items(
            w,
            "rowItems",
            &pivot.rows,
            &fields,
            part.axis_tuples.rows.as_deref(),
            pivot.layout.show_row_grand_totals,
            false,
        )?;
        write_axis_fields(
            w,
            "colFields",
            &pivot.columns,
            &fields,
            values_field_on_axis(pivot, PivotValuesAxis::Columns),
            pivot.layout.values_axis_position,
        )?;
        write_axis_items(
            w,
            "colItems",
            &pivot.columns,
            &fields,
            part.axis_tuples.columns.as_deref(),
            pivot.layout.show_column_grand_totals,
            true,
        )?;
        write_page_fields(w, pivot, &fields)?;
        write_data_fields(w, pivot, &fields, style_table)?;
        write_pivot_style(w, pivot)?;
        write_pivot_filters(w, pivot, &fields)?;
        write_pivot_extensions(w, pivot)?;

        w.write_event(Event::End(BytesEnd::new("pivotTableDefinition")))?;
        Ok(())
    })
}

fn write_location(
    w: &mut XmlWriter,
    pivot: &PivotTable,
    fields: &[CacheField],
    table: &FormatPivotTable,
) -> XlsxResult<()> {
    let range = match pivot.rendered_range {
        Some(range) => range,
        None => estimated_pivot_output_range(pivot, fields, table)?,
    };
    let ref_str = range.to_a1_string();
    let first_data_col = expanded_axis_field_count(&pivot.rows, fields)
        .max(1)
        .to_string();

    let mut location = BytesStart::new("location");
    location.push_attribute(("ref", ref_str.as_str()));
    location.push_attribute(("firstHeaderRow", "1"));
    location.push_attribute(("firstDataRow", "1"));
    location.push_attribute(("firstDataCol", first_data_col.as_str()));
    w.write_event(Event::Empty(location))?;
    Ok(())
}

fn estimated_pivot_output_range(
    pivot: &PivotTable,
    fields: &[CacheField],
    table: &FormatPivotTable,
) -> XlsxResult<CellRange> {
    let row_label_width = expanded_axis_field_count(&pivot.rows, fields).max(1);
    let measure_width = pivot.measures.len().max(1);
    let column_tuple_count = table
        .axis_tuples
        .columns
        .as_ref()
        .map_or(0, Vec::len)
        .max(1);
    let value_width = if pivot.columns.is_empty() {
        measure_width
    } else {
        column_tuple_count * measure_width
    };
    let width = row_label_width + value_width;

    let row_tuple_count = table.axis_tuples.rows.as_ref().map_or(0, Vec::len);
    let data_rows = if pivot.rows.is_empty() {
        1
    } else {
        row_tuple_count + usize::from(pivot.layout.show_row_grand_totals)
    }
    .max(1);
    let header_rows = if pivot.columns.is_empty() {
        1
    } else {
        expanded_axis_field_count(&pivot.columns, fields).max(1) + 1
    };
    let height = header_rows + data_rows;

    Ok(range_from_size(pivot.target, width, height))
}

fn range_from_size(start: CellAddress, width: usize, height: usize) -> CellRange {
    const MAX_ROWS: usize = 1_048_576;
    const MAX_COLS: usize = 16_384;

    let max_width = MAX_COLS.saturating_sub(start.col as usize).max(1);
    let max_height = MAX_ROWS.saturating_sub(start.row as usize).max(1);
    let width = width.max(1).min(max_width);
    let height = height.max(1).min(max_height);
    let end_row = start.row + (height - 1) as u32;
    let end_col = start.col + (width - 1) as u16;
    CellRange::new(start, CellAddress::new(end_row, end_col))
}

fn write_pivot_fields(
    w: &mut XmlWriter,
    pivot: &PivotTable,
    fields: &[CacheField],
) -> XlsxResult<()> {
    let count = fields.len().to_string();
    let mut pivot_fields = BytesStart::new("pivotFields");
    pivot_fields.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(pivot_fields))?;

    for (index, field) in fields.iter().enumerate() {
        let mut pivot_field = BytesStart::new("pivotField");
        if pivot_field_is_data_field(pivot, field) {
            pivot_field.push_attribute(("dataField", "1"));
        }
        if let Some(axis) = field_axis(pivot, fields, index) {
            pivot_field.push_attribute(("axis", axis));
        }
        let sort = field_sort(pivot, fields, index);
        if sort != "manual" {
            pivot_field.push_attribute(("sortType", sort));
        }
        if field_is_filtered(pivot, &field.name) {
            pivot_field.push_attribute(("multipleItemSelectionAllowed", "1"));
        }
        let sort_measure_index = field_sort_measure_index(pivot, fields, index)?;
        let _item_page_count_attr = if let Some(axis_field) = pivot_axis_field(pivot, fields, index)
        {
            if let Some(caption) = &axis_field.caption {
                pivot_field.push_attribute(("name", caption.as_str()));
            }
            pivot_field.push_attribute(("showAll", bool_attr(axis_field.show_empty_items)));
            if let Some(caption) = &axis_field.subtotal_caption {
                pivot_field.push_attribute(("subtotalCaption", caption.as_str()));
            }
            push_pivot_field_option_attrs(&mut pivot_field, axis_field);
            let item_page_count_attr =
                (axis_field.item_page_count != 10).then(|| axis_field.item_page_count.to_string());
            if let Some(item_page_count) = &item_page_count_attr {
                pivot_field.push_attribute(("itemPageCount", item_page_count.as_str()));
            }
            push_subtotal_attrs(&mut pivot_field, axis_field);
            item_page_count_attr
        } else {
            None
        };

        let hidden_items = hidden_item_indexes(pivot, fields, index)?;
        let collapsed_items = collapsed_item_indexes(pivot, fields, index)?;
        let include_default = should_write_pivot_field_items(pivot, fields, index);
        if hidden_items.is_empty()
            && collapsed_items.is_empty()
            && !include_default
            && sort_measure_index.is_none()
        {
            w.write_event(Event::Empty(pivot_field))?;
        } else {
            w.write_event(Event::Start(pivot_field))?;
            if include_default || !hidden_items.is_empty() || !collapsed_items.is_empty() {
                write_pivot_field_items(
                    w,
                    field,
                    &hidden_items,
                    &collapsed_items,
                    include_default,
                )?;
            }
            if let Some(measure_index) = sort_measure_index {
                write_auto_sort_scope(w, measure_index)?;
            }
            w.write_event(Event::End(BytesEnd::new("pivotField")))?;
        }
    }

    w.write_event(Event::End(BytesEnd::new("pivotFields")))?;
    Ok(())
}

fn pivot_field_is_data_field(pivot: &PivotTable, field: &CacheField) -> bool {
    pivot
        .measures
        .iter()
        .any(|measure| measure.field.name.eq_ignore_ascii_case(&field.name))
}

fn should_write_pivot_field_items(
    pivot: &PivotTable,
    fields: &[CacheField],
    field_index: usize,
) -> bool {
    pivot_axis_field(pivot, fields, field_index).is_some()
        || fields
            .get(field_index)
            .is_some_and(|field| field.grouping.is_some())
}

fn write_pivot_field_items(
    w: &mut XmlWriter,
    field: &CacheField,
    hidden_items: &[u32],
    collapsed_items: &[u32],
    include_default: bool,
) -> XlsxResult<()> {
    let item_count = pivot_field_item_count(field);
    let count = (item_count + usize::from(include_default)).to_string();
    let mut items = BytesStart::new("items");
    items.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(items))?;
    for item_index in 0..item_count {
        let x = item_index.to_string();
        let mut item = BytesStart::new("item");
        item.push_attribute(("x", x.as_str()));
        if hidden_items.contains(&(item_index as u32)) {
            item.push_attribute(("h", "1"));
        }
        if collapsed_items.contains(&(item_index as u32)) {
            item.push_attribute(("sd", "0"));
        }
        w.write_event(Event::Empty(item))?;
    }
    if include_default {
        let mut item = BytesStart::new("item");
        item.push_attribute(("t", "default"));
        w.write_event(Event::Empty(item))?;
    }
    w.write_event(Event::End(BytesEnd::new("items")))?;
    Ok(())
}

fn pivot_field_item_count(field: &CacheField) -> usize {
    if !field.database_field {
        if let Some(grouping) = &field.grouping {
            if matches!(grouping.definition, PivotGrouping::Manual { .. }) {
                return grouping.levels[0].group_items.len();
            }
        }
    }
    field.shared_items.len()
}

fn field_axis(
    pivot: &PivotTable,
    fields: &[CacheField],
    field_index: usize,
) -> Option<&'static str> {
    let field_name = axis_semantic_field_name(fields, field_index)?;
    if pivot
        .rows
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(&field_name))
    {
        Some("axisRow")
    } else if pivot
        .columns
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(&field_name))
    {
        Some("axisCol")
    } else if pivot
        .page_fields
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(&field_name))
    {
        Some("axisPage")
    } else {
        None
    }
}

fn field_sort(pivot: &PivotTable, fields: &[CacheField], field_index: usize) -> &'static str {
    let Some(field_name) = axis_semantic_field_name(fields, field_index) else {
        return "manual";
    };
    let sort = pivot
        .rows
        .iter()
        .chain(pivot.columns.iter())
        .chain(pivot.page_fields.iter())
        .find(|field| field.field.name.eq_ignore_ascii_case(&field_name))
        .map(|field| field.sort)
        .unwrap_or(PivotSort::None);

    match sort {
        PivotSort::None => "manual",
        PivotSort::Ascending => "ascending",
        PivotSort::Descending => "descending",
    }
}

fn field_sort_measure_index(
    pivot: &PivotTable,
    fields: &[CacheField],
    field_index: usize,
) -> XlsxResult<Option<usize>> {
    let Some(field_name) = axis_semantic_field_name(fields, field_index) else {
        return Ok(None);
    };
    let Some(axis_field) = pivot
        .rows
        .iter()
        .chain(pivot.columns.iter())
        .chain(pivot.page_fields.iter())
        .find(|field| field.field.name.eq_ignore_ascii_case(&field_name))
    else {
        return Ok(None);
    };

    if matches!(axis_field.sort, PivotSort::None) {
        return Ok(None);
    }

    let Some(sort_measure) = axis_field.sort_by_measure.as_ref() else {
        return Ok(None);
    };
    pivot
        .measures
        .iter()
        .position(|measure| pivot_measure_matches_target(measure, sort_measure))
        .map(Some)
        .ok_or_else(|| {
            XlsxError::InvalidFormat(format!(
                "pivot table {} sorts field {} by an unknown measure",
                pivot.name, axis_field.field.name
            ))
        })
}

fn write_auto_sort_scope(w: &mut XmlWriter, measure_index: usize) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("autoSortScope")))?;

    let mut pivot_area = BytesStart::new("pivotArea");
    pivot_area.push_attribute(("dataOnly", "0"));
    pivot_area.push_attribute(("outline", "0"));
    pivot_area.push_attribute(("fieldPosition", "0"));
    w.write_event(Event::Start(pivot_area))?;

    let mut references = BytesStart::new("references");
    references.push_attribute(("count", "1"));
    w.write_event(Event::Start(references))?;

    let mut reference = BytesStart::new("reference");
    reference.push_attribute(("field", "4294967294"));
    reference.push_attribute(("count", "1"));
    reference.push_attribute(("selected", "0"));
    w.write_event(Event::Start(reference))?;

    let v = measure_index.to_string();
    let mut x = BytesStart::new("x");
    x.push_attribute(("v", v.as_str()));
    w.write_event(Event::Empty(x))?;

    w.write_event(Event::End(BytesEnd::new("reference")))?;
    w.write_event(Event::End(BytesEnd::new("references")))?;
    w.write_event(Event::End(BytesEnd::new("pivotArea")))?;
    w.write_event(Event::End(BytesEnd::new("autoSortScope")))?;
    Ok(())
}

fn field_is_filtered(pivot: &PivotTable, field_name: &str) -> bool {
    pivot.filters.iter().any(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems { field, .. }
                if field.name.eq_ignore_ascii_case(field_name)
        )
    })
}

fn pivot_axis_field<'a>(
    pivot: &'a PivotTable,
    fields: &'a [CacheField],
    field_index: usize,
) -> Option<&'a duke_sheets_core::PivotField> {
    let field_name = axis_semantic_field_name(fields, field_index)?;
    pivot
        .rows
        .iter()
        .chain(pivot.columns.iter())
        .chain(pivot.page_fields.iter())
        .find(|field| field.field.name.eq_ignore_ascii_case(&field_name))
}

fn axis_semantic_field_name(fields: &[CacheField], field_index: usize) -> Option<String> {
    let field = fields.get(field_index)?;
    if let Some(grouping) = &field.grouping {
        if matches!(grouping.definition, PivotGrouping::Manual { .. }) && !field.database_field {
            return fields
                .get(grouping.base_field_index)
                .map(|base| base.name.clone());
        }
    }
    if field.grouping.is_some() {
        return Some(field.name.clone());
    }
    if let Some((base, _)) = grouping_level_for_field(fields, field_index) {
        return fields
            .get(base.base_field_index)
            .map(|field| field.name.clone());
    }
    if has_grouped_children(fields, field_index) {
        return None;
    }
    Some(field.name.clone())
}

fn has_grouped_children(fields: &[CacheField], field_index: usize) -> bool {
    fields
        .get(field_index)
        .and_then(|field| field.grouping.as_ref())
        .is_some_and(|grouping| {
            grouping
                .levels
                .iter()
                .any(|level| level.field_index != field_index)
        })
}

fn push_pivot_field_option_attrs(pivot_field: &mut BytesStart<'_>, field: &PivotField) {
    if !field.show_drop_downs {
        pivot_field.push_attribute(("showDropDowns", "0"));
    }
    if !field.subtotal_top {
        pivot_field.push_attribute(("subtotalTop", "0"));
    }
    if field.insert_blank_row {
        pivot_field.push_attribute(("insertBlankRow", "1"));
    }
    if field.insert_page_break {
        pivot_field.push_attribute(("insertPageBreak", "1"));
    }
    if field.include_new_items_in_filter {
        pivot_field.push_attribute(("includeNewItemsInFilter", "1"));
    }
}

fn push_subtotal_attrs(pivot_field: &mut BytesStart<'_>, field: &duke_sheets_core::PivotField) {
    if !field.subtotals.is_empty() {
        if field.subtotals.iter().any(|subtotal| {
            matches!(subtotal, PivotSubtotal::None) || subtotal.is_custom_function()
        }) {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
        }
        for subtotal in field
            .subtotals
            .iter()
            .copied()
            .filter(|subtotal| subtotal.is_custom_function())
        {
            push_subtotal_function_attr(pivot_field, subtotal);
        }
        return;
    }

    let subtotal = field.subtotal;
    match subtotal {
        PivotSubtotal::Automatic => {}
        PivotSubtotal::None => pivot_field.push_attribute(("defaultSubtotal", "0")),
        PivotSubtotal::Sum => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            push_subtotal_function_attr(pivot_field, subtotal);
        }
        PivotSubtotal::Count => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            push_subtotal_function_attr(pivot_field, subtotal);
        }
        PivotSubtotal::CountNumbers => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            push_subtotal_function_attr(pivot_field, subtotal);
        }
        PivotSubtotal::Average => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            push_subtotal_function_attr(pivot_field, subtotal);
        }
        PivotSubtotal::Min => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            push_subtotal_function_attr(pivot_field, subtotal);
        }
        PivotSubtotal::Max => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            push_subtotal_function_attr(pivot_field, subtotal);
        }
        PivotSubtotal::Product => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            push_subtotal_function_attr(pivot_field, subtotal);
        }
        PivotSubtotal::StdDev => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            push_subtotal_function_attr(pivot_field, subtotal);
        }
        PivotSubtotal::StdDevP => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            push_subtotal_function_attr(pivot_field, subtotal);
        }
        PivotSubtotal::Var => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            push_subtotal_function_attr(pivot_field, subtotal);
        }
        PivotSubtotal::VarP => {
            pivot_field.push_attribute(("defaultSubtotal", "0"));
            push_subtotal_function_attr(pivot_field, subtotal);
        }
    }
}

fn push_subtotal_function_attr(pivot_field: &mut BytesStart<'_>, subtotal: PivotSubtotal) {
    match subtotal {
        PivotSubtotal::Automatic | PivotSubtotal::None => {}
        PivotSubtotal::Sum => pivot_field.push_attribute(("sumSubtotal", "1")),
        PivotSubtotal::Count => pivot_field.push_attribute(("countASubtotal", "1")),
        PivotSubtotal::CountNumbers => pivot_field.push_attribute(("countSubtotal", "1")),
        PivotSubtotal::Average => pivot_field.push_attribute(("avgSubtotal", "1")),
        PivotSubtotal::Min => pivot_field.push_attribute(("minSubtotal", "1")),
        PivotSubtotal::Max => pivot_field.push_attribute(("maxSubtotal", "1")),
        PivotSubtotal::Product => pivot_field.push_attribute(("productSubtotal", "1")),
        PivotSubtotal::StdDev => pivot_field.push_attribute(("stdDevSubtotal", "1")),
        PivotSubtotal::StdDevP => pivot_field.push_attribute(("stdDevPSubtotal", "1")),
        PivotSubtotal::Var => pivot_field.push_attribute(("varSubtotal", "1")),
        PivotSubtotal::VarP => pivot_field.push_attribute(("varPSubtotal", "1")),
    }
}

fn hidden_item_indexes(
    pivot: &PivotTable,
    fields: &[CacheField],
    field_index: usize,
) -> XlsxResult<Vec<u32>> {
    let field = &fields[field_index];
    let Some(PivotFilter::FieldItems { allowed_items, .. }) = pivot.filters.iter().find(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems { field: filter_field, .. }
                if filter_field.name.eq_ignore_ascii_case(&field.name)
        )
    }) else {
        return Ok(Vec::new());
    };

    let allowed = allowed_items
        .iter()
        .filter_map(|item| {
            field
                .shared_items
                .iter()
                .position(|candidate| candidate == item)
                .map(|index| index as u32)
        })
        .collect::<std::collections::HashSet<_>>();

    Ok((0..field.shared_items.len() as u32)
        .filter(|index| !allowed.contains(index))
        .collect())
}

fn collapsed_item_indexes(
    pivot: &PivotTable,
    fields: &[CacheField],
    field_index: usize,
) -> XlsxResult<Vec<u32>> {
    let Some(axis_field) = pivot_axis_field(pivot, fields, field_index) else {
        return Ok(Vec::new());
    };
    let field = &fields[field_index];
    axis_field
        .collapsed_items
        .iter()
        .map(|item| {
            pivot_field_item_index(field, item).ok_or_else(|| {
                XlsxError::InvalidFormat(format!(
                    "pivot field {} collapsed item was not found in the pivot cache: {}",
                    field.name, item
                ))
            })
        })
        .collect()
}

fn pivot_field_item_index(field: &CacheField, item: &PivotValue) -> Option<u32> {
    if !field.database_field {
        if let Some(grouping) = &field.grouping {
            if matches!(grouping.definition, PivotGrouping::Manual { .. }) {
                return grouping.levels[0]
                    .group_items
                    .iter()
                    .position(|value| value == item)
                    .map(|index| index as u32);
            }
        }
    }
    field
        .shared_items
        .iter()
        .position(|value| value == item)
        .map(|index| index as u32)
}

fn write_axis_fields(
    w: &mut XmlWriter,
    tag_name: &str,
    axis_fields: &[duke_sheets_core::PivotField],
    fields: &[CacheField],
    include_values_field: bool,
    values_position: Option<u32>,
) -> XlsxResult<()> {
    if axis_fields.is_empty() && !include_values_field {
        return Ok(());
    }

    let mut indexes = expanded_axis_field_indexes(axis_fields, fields)?
        .into_iter()
        .map(|index| index as i32)
        .collect::<Vec<_>>();
    if include_values_field {
        let position = values_position
            .map(|position| position as usize)
            .unwrap_or(indexes.len())
            .min(indexes.len());
        indexes.insert(position, -2);
    }
    let count = indexes.len().to_string();
    let mut tag = BytesStart::new(tag_name);
    tag.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(tag))?;
    for index in indexes {
        let x = index.to_string();
        let mut el = BytesStart::new("field");
        el.push_attribute(("x", x.as_str()));
        w.write_event(Event::Empty(el))?;
    }
    w.write_event(Event::End(BytesEnd::new(tag_name)))?;
    Ok(())
}

fn values_field_on_axis(pivot: &PivotTable, axis: PivotValuesAxis) -> bool {
    pivot.measures.len() > 1
        && pivot.layout.values_axis == axis
        && (matches!(axis, PivotValuesAxis::Rows) || pivot.layout.values_axis_position.is_some())
}

fn write_axis_items(
    w: &mut XmlWriter,
    tag_name: &str,
    axis_fields: &[duke_sheets_core::PivotField],
    fields: &[CacheField],
    tuples: Option<&[Vec<u32>]>,
    include_grand_total: bool,
    write_empty_item_when_no_fields: bool,
) -> XlsxResult<()> {
    if axis_fields.is_empty() {
        if write_empty_item_when_no_fields {
            let mut tag = BytesStart::new(tag_name);
            tag.push_attribute(("count", "1"));
            w.write_event(Event::Start(tag))?;
            w.write_event(Event::Empty(BytesStart::new("i")))?;
            w.write_event(Event::End(BytesEnd::new(tag_name)))?;
        }
        return Ok(());
    }

    let field_indexes = expanded_axis_field_indexes(axis_fields, fields)?;
    let tuples = tuples.unwrap_or(&[]);

    let count = (tuples.len() + usize::from(include_grand_total)).to_string();
    let mut tag = BytesStart::new(tag_name);
    tag.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(tag))?;

    for tuple in tuples {
        write_axis_item(w, None, &tuple)?;
    }
    if include_grand_total {
        let grand = vec![0; field_indexes.len()];
        write_axis_item(w, Some("grand"), &grand)?;
    }

    w.write_event(Event::End(BytesEnd::new(tag_name)))?;
    Ok(())
}

fn write_axis_item(w: &mut XmlWriter, item_type: Option<&str>, indexes: &[u32]) -> XlsxResult<()> {
    let mut item = BytesStart::new("i");
    if let Some(item_type) = item_type {
        item.push_attribute(("t", item_type));
    }
    if indexes.is_empty() {
        w.write_event(Event::Empty(item))?;
        return Ok(());
    }

    w.write_event(Event::Start(item))?;
    for index in indexes {
        let mut x = BytesStart::new("x");
        let value = index.to_string();
        if *index != 0 {
            x.push_attribute(("v", value.as_str()));
        }
        w.write_event(Event::Empty(x))?;
    }
    w.write_event(Event::End(BytesEnd::new("i")))?;
    Ok(())
}

fn expanded_axis_field_count(
    axis_fields: &[duke_sheets_core::PivotField],
    fields: &[CacheField],
) -> usize {
    axis_fields
        .iter()
        .map(|field| {
            grouped_cache_field_indexes(fields, &field.field.name)
                .map(|indexes| indexes.len())
                .unwrap_or(1)
        })
        .sum()
}

fn expanded_axis_field_indexes(
    axis_fields: &[duke_sheets_core::PivotField],
    fields: &[CacheField],
) -> XlsxResult<Vec<usize>> {
    let mut indexes = Vec::new();
    for field in axis_fields {
        if let Some(grouped_indexes) = grouped_cache_field_indexes(fields, &field.field.name) {
            indexes.extend(grouped_indexes);
            continue;
        }

        let index = field_index(fields, &field.field.name).ok_or_else(|| {
            XlsxError::InvalidFormat(format!("pivot field not found: {}", field.field.name))
        })?;
        indexes.push(index);
    }
    Ok(indexes)
}

fn grouped_cache_field_indexes(fields: &[CacheField], field_name: &str) -> Option<Vec<usize>> {
    let base = field_index(fields, field_name)?;
    let grouping = fields.get(base)?.grouping.as_ref()?;
    match &grouping.definition {
        PivotGrouping::Manual { .. } => fields
            .iter()
            .enumerate()
            .find(|(_, field)| {
                !field.database_field
                    && field.grouping.as_ref().is_some_and(|candidate| {
                        matches!(candidate.definition, PivotGrouping::Manual { .. })
                            && candidate.base_field_index == base
                    })
            })
            .map(|(grouped, _)| vec![grouped, base]),
        PivotGrouping::Date { units, .. } if units.len() > 1 => Some(
            grouping
                .levels
                .iter()
                .map(|level| level.field_index)
                .collect(),
        ),
        _ => None,
    }
}

fn write_page_fields(
    w: &mut XmlWriter,
    pivot: &PivotTable,
    fields: &[CacheField],
) -> XlsxResult<()> {
    if pivot.page_fields.is_empty() {
        return Ok(());
    }

    let count = pivot.page_fields.len().to_string();
    let mut page_fields = BytesStart::new("pageFields");
    page_fields.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(page_fields))?;
    for field in &pivot.page_fields {
        let index = field_index(fields, &field.field.name).ok_or_else(|| {
            XlsxError::InvalidFormat(format!("pivot field not found: {}", field.field.name))
        })?;
        let fld = index.to_string();
        let mut el = BytesStart::new("pageField");
        el.push_attribute(("fld", fld.as_str()));
        if let Some(item) = selected_page_item_index(pivot, &field.field.name, &fields[index]) {
            let item = item.to_string();
            el.push_attribute(("item", item.as_str()));
        }
        w.write_event(Event::Empty(el))?;
    }
    w.write_event(Event::End(BytesEnd::new("pageFields")))?;
    Ok(())
}

fn selected_page_item_index(
    pivot: &PivotTable,
    field_name: &str,
    field: &CacheField,
) -> Option<u32> {
    let PivotFilter::FieldItems { allowed_items, .. } = pivot.filters.iter().find(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems { field, .. }
                if field.name.eq_ignore_ascii_case(field_name)
        )
    })?
    else {
        return None;
    };

    let [item] = allowed_items.as_slice() else {
        return None;
    };
    field
        .shared_items
        .iter()
        .position(|candidate| candidate == item)
        .map(|index| index as u32)
}

fn write_data_fields(
    w: &mut XmlWriter,
    pivot: &PivotTable,
    fields: &[CacheField],
    style_table: &XlsxStyleTable,
) -> XlsxResult<()> {
    let count = pivot.measures.len().to_string();
    let mut data_fields = BytesStart::new("dataFields");
    data_fields.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(data_fields))?;
    for measure in &pivot.measures {
        let index = field_index(fields, &measure.field.name).ok_or_else(|| {
            XlsxError::InvalidFormat(format!("pivot field not found: {}", measure.field.name))
        })?;
        let fld = index.to_string();
        let name = measure.caption();
        let mut data_field = BytesStart::new("dataField");
        data_field.push_attribute(("name", name.as_str()));
        data_field.push_attribute(("fld", fld.as_str()));
        data_field.push_attribute(("subtotal", aggregate_name(measure.aggregate)));
        let num_fmt_id = if let Some(number_format) = &measure.number_format {
            Some(
                style_table
                    .custom_num_fmt_id(number_format)
                    .ok_or_else(|| {
                        XlsxError::InvalidFormat(format!(
                            "pivot measure number format was not registered: {number_format}"
                        ))
                    })?
                    .to_string(),
            )
        } else {
            None
        };
        if let Some(num_fmt_id) = num_fmt_id.as_deref() {
            data_field.push_attribute(("numFmtId", num_fmt_id));
        }
        if let Some(show_data_as) = show_data_as_name(&measure.show_as) {
            data_field.push_attribute(("showDataAs", show_data_as));
        }
        let base_field = show_as_base_field_index(&measure.show_as, fields)?;
        if let Some(base_field) = base_field {
            let base_field = base_field.to_string();
            data_field.push_attribute(("baseField", base_field.as_str()));
        }
        let base_item = show_as_base_item_index(&measure.show_as, fields)?;
        if let Some(base_item) = base_item {
            let base_item = base_item.to_string();
            data_field.push_attribute(("baseItem", base_item.as_str()));
        }
        if let Some(x14_show_as) = x14_show_as_name(&measure.show_as) {
            let source_field = x14_show_as_source_field_index(&measure.show_as, fields)?;
            w.write_event(Event::Start(data_field))?;
            write_data_field_ext_with_source(w, x14_show_as, source_field)?;
            w.write_event(Event::End(BytesEnd::new("dataField")))?;
        } else {
            w.write_event(Event::Empty(data_field))?;
        }
    }
    w.write_event(Event::End(BytesEnd::new("dataFields")))?;
    Ok(())
}

fn write_pivot_filters(
    w: &mut XmlWriter,
    pivot: &PivotTable,
    fields: &[CacheField],
) -> XlsxResult<()> {
    let filters = pivot
        .filters
        .iter()
        .filter(|filter| !matches!(filter, PivotFilter::FieldItems { .. }))
        .collect::<Vec<_>>();
    if filters.is_empty() {
        return Ok(());
    }

    let count = filters.len().to_string();
    let mut filters_el = BytesStart::new("filters");
    filters_el.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(filters_el))?;
    for (index, filter) in filters.into_iter().enumerate() {
        write_pivot_filter(w, pivot, fields, filter, index)?;
    }
    w.write_event(Event::End(BytesEnd::new("filters")))?;
    Ok(())
}

fn write_pivot_filter(
    w: &mut XmlWriter,
    pivot: &PivotTable,
    fields: &[CacheField],
    filter: &PivotFilter,
    index: usize,
) -> XlsxResult<()> {
    let field_name = filter_field_ref(filter)
        .map(|field| field.name.as_str())
        .ok_or_else(|| XlsxError::InvalidFormat("unsupported pivot filter".into()))?;
    let field_index = field_index(fields, field_name).ok_or_else(|| {
        XlsxError::InvalidFormat(format!("pivot filter field not found: {field_name}"))
    })?;
    let fld = field_index.to_string();
    let eval_order = index.to_string();
    let id = (index + 1).to_string();

    let mut filter_el = BytesStart::new("filter");
    filter_el.push_attribute(("fld", fld.as_str()));
    filter_el.push_attribute(("evalOrder", eval_order.as_str()));
    filter_el.push_attribute(("id", id.as_str()));

    match filter {
        PivotFilter::Label {
            operator, value, ..
        } => {
            filter_el.push_attribute(("type", label_filter_type_name(*operator)));
            filter_el.push_attribute(("stringValue1", value.as_str()));
            w.write_event(Event::Start(filter_el))?;
            let custom_value = label_custom_filter_value(*operator, value);
            write_pivot_custom_filter(w, custom_filter_operator(*operator), &custom_value)?;
            w.write_event(Event::End(BytesEnd::new("filter")))?;
        }
        PivotFilter::LabelBetween {
            start,
            end,
            not_between,
            ..
        } => {
            filter_el.push_attribute((
                "type",
                if *not_between {
                    "captionNotBetween"
                } else {
                    "captionBetween"
                },
            ));
            filter_el.push_attribute(("stringValue1", start.as_str()));
            filter_el.push_attribute(("stringValue2", end.as_str()));
            w.write_event(Event::Start(filter_el))?;
            if *not_between {
                write_pivot_custom_filters(
                    w,
                    false,
                    &[("lessThan", start.clone()), ("greaterThan", end.clone())],
                )?;
            } else {
                write_pivot_custom_filters(
                    w,
                    true,
                    &[
                        ("greaterThanOrEqual", start.clone()),
                        ("lessThanOrEqual", end.clone()),
                    ],
                )?;
            }
            w.write_event(Event::End(BytesEnd::new("filter")))?;
        }
        PivotFilter::Value {
            measure,
            operator,
            value,
            ..
        } => {
            let filter_type = value_filter_type_name(*operator).ok_or_else(|| {
                XlsxError::InvalidFormat("unsupported pivot value filter operator".into())
            })?;
            let measure_index = measure_index_for_filter(pivot, measure)?;
            let i_measure_fld = measure_index.to_string();
            let value = value.to_string();
            filter_el.push_attribute(("type", filter_type));
            filter_el.push_attribute(("iMeasureFld", i_measure_fld.as_str()));
            filter_el.push_attribute(("stringValue1", value.as_str()));
            w.write_event(Event::Start(filter_el))?;
            write_pivot_custom_filter(w, custom_filter_operator(*operator), &value)?;
            w.write_event(Event::End(BytesEnd::new("filter")))?;
        }
        PivotFilter::ValueBetween {
            measure,
            start,
            end,
            not_between,
            ..
        } => {
            let measure_index = measure_index_for_filter(pivot, measure)?;
            let i_measure_fld = measure_index.to_string();
            let start = start.to_string();
            let end = end.to_string();
            filter_el.push_attribute((
                "type",
                if *not_between {
                    "valueNotBetween"
                } else {
                    "valueBetween"
                },
            ));
            filter_el.push_attribute(("iMeasureFld", i_measure_fld.as_str()));
            filter_el.push_attribute(("stringValue1", start.as_str()));
            filter_el.push_attribute(("stringValue2", end.as_str()));
            w.write_event(Event::Start(filter_el))?;
            if *not_between {
                write_pivot_custom_filters(
                    w,
                    false,
                    &[("lessThan", start.clone()), ("greaterThan", end.clone())],
                )?;
            } else {
                write_pivot_custom_filters(
                    w,
                    true,
                    &[
                        ("greaterThanOrEqual", start.clone()),
                        ("lessThanOrEqual", end.clone()),
                    ],
                )?;
            }
            w.write_event(Event::End(BytesEnd::new("filter")))?;
        }
        PivotFilter::Date {
            operator, value, ..
        } => {
            let filter_type = date_filter_type_name(*operator).ok_or_else(|| {
                XlsxError::InvalidFormat("unsupported pivot date filter operator".into())
            })?;
            let value = value.to_string();
            filter_el.push_attribute(("type", filter_type));
            filter_el.push_attribute(("stringValue1", value.as_str()));
            w.write_event(Event::Start(filter_el))?;
            write_pivot_custom_filter(w, custom_filter_operator(*operator), &value)?;
            w.write_event(Event::End(BytesEnd::new("filter")))?;
        }
        PivotFilter::DateBetween {
            start,
            end,
            not_between,
            ..
        } => {
            let start = start.to_string();
            let end = end.to_string();
            filter_el.push_attribute((
                "type",
                if *not_between {
                    "dateNotBetween"
                } else {
                    "dateBetween"
                },
            ));
            filter_el.push_attribute(("stringValue1", start.as_str()));
            filter_el.push_attribute(("stringValue2", end.as_str()));
            w.write_event(Event::Start(filter_el))?;
            if *not_between {
                write_pivot_custom_filters(
                    w,
                    false,
                    &[("lessThan", start.clone()), ("greaterThan", end.clone())],
                )?;
            } else {
                write_pivot_custom_filters(
                    w,
                    true,
                    &[
                        ("greaterThanOrEqual", start.clone()),
                        ("lessThanOrEqual", end.clone()),
                    ],
                )?;
            }
            w.write_event(Event::End(BytesEnd::new("filter")))?;
        }
        PivotFilter::DatePeriod { period, .. } => {
            let filter_type = date_period_filter_type_name(*period).ok_or_else(|| {
                XlsxError::InvalidFormat("unsupported pivot date period filter".into())
            })?;
            filter_el.push_attribute(("type", filter_type));
            w.write_event(Event::Empty(filter_el))?;
        }
        PivotFilter::TopN {
            measure,
            n,
            top,
            percent,
            ..
        } => {
            let measure_index = measure_index_for_filter(pivot, measure)?;
            let i_measure_fld = measure_index.to_string();
            filter_el.push_attribute(("type", top_n_filter_type_name(*top, *percent)));
            filter_el.push_attribute(("iMeasureFld", i_measure_fld.as_str()));
            w.write_event(Event::Start(filter_el))?;
            write_pivot_top_filter(w, *n, *top, *percent)?;
            w.write_event(Event::End(BytesEnd::new("filter")))?;
        }
        PivotFilter::FieldItems { .. } => {}
        PivotFilter::Unsupported { kind, .. } => {
            return Err(XlsxError::InvalidFormat(format!(
                "unsupported pivot filter: {kind}"
            )));
        }
    }

    Ok(())
}

fn write_pivot_custom_filter(
    w: &mut XmlWriter,
    operator: &'static str,
    value: &str,
) -> XlsxResult<()> {
    write_pivot_custom_filters(w, true, &[(operator, value.to_string())])
}

fn write_pivot_custom_filters(
    w: &mut XmlWriter,
    and: bool,
    filters: &[(&'static str, String)],
) -> XlsxResult<()> {
    write_pivot_auto_filter_start(w)?;

    let mut custom_filters = BytesStart::new("customFilters");
    if filters.len() > 1 {
        custom_filters.push_attribute(("and", bool_attr(and)));
    }
    w.write_event(Event::Start(custom_filters))?;
    for (operator, value) in filters {
        let mut custom_filter = BytesStart::new("customFilter");
        custom_filter.push_attribute(("operator", *operator));
        custom_filter.push_attribute(("val", value.as_str()));
        w.write_event(Event::Empty(custom_filter))?;
    }
    w.write_event(Event::End(BytesEnd::new("customFilters")))?;

    write_pivot_auto_filter_end(w)?;
    Ok(())
}

fn write_pivot_top_filter(w: &mut XmlWriter, n: u32, top: bool, percent: bool) -> XlsxResult<()> {
    write_pivot_auto_filter_start(w)?;

    let n = n.to_string();
    let mut top10 = BytesStart::new("top10");
    top10.push_attribute(("top", bool_attr(top)));
    top10.push_attribute(("percent", bool_attr(percent)));
    top10.push_attribute(("val", n.as_str()));
    w.write_event(Event::Empty(top10))?;

    write_pivot_auto_filter_end(w)?;
    Ok(())
}

fn write_pivot_auto_filter_start(w: &mut XmlWriter) -> XlsxResult<()> {
    let mut auto_filter = BytesStart::new("autoFilter");
    auto_filter.push_attribute(("ref", "A1"));
    w.write_event(Event::Start(auto_filter))?;

    let mut filter_column = BytesStart::new("filterColumn");
    filter_column.push_attribute(("colId", "0"));
    w.write_event(Event::Start(filter_column))?;
    Ok(())
}

fn write_pivot_auto_filter_end(w: &mut XmlWriter) -> XlsxResult<()> {
    w.write_event(Event::End(BytesEnd::new("filterColumn")))?;
    w.write_event(Event::End(BytesEnd::new("autoFilter")))?;
    Ok(())
}

fn measure_index_for_filter(
    pivot: &PivotTable,
    filter_measure: &duke_sheets_core::PivotMeasure,
) -> XlsxResult<usize> {
    pivot
        .measures
        .iter()
        .position(|measure| {
            measure
                .field
                .name
                .eq_ignore_ascii_case(&filter_measure.field.name)
                && measure.aggregate == filter_measure.aggregate
                && match filter_measure.name.as_ref() {
                    Some(name) => measure
                        .name
                        .as_ref()
                        .map(|candidate| candidate.eq_ignore_ascii_case(name))
                        .unwrap_or_else(|| measure.caption().eq_ignore_ascii_case(name)),
                    None => true,
                }
        })
        .ok_or_else(|| {
            XlsxError::InvalidFormat(format!(
                "pivot filter measure not found: {}",
                filter_measure.caption()
            ))
        })
}

fn label_filter_type_name(operator: PivotFilterOperator) -> &'static str {
    match operator {
        PivotFilterOperator::Equals => "captionEqual",
        PivotFilterOperator::NotEquals => "captionNotEqual",
        PivotFilterOperator::LessThan => "captionLessThan",
        PivotFilterOperator::LessThanOrEqual => "captionLessThanOrEqual",
        PivotFilterOperator::GreaterThan => "captionGreaterThan",
        PivotFilterOperator::GreaterThanOrEqual => "captionGreaterThanOrEqual",
        PivotFilterOperator::BeginsWith => "captionBeginsWith",
        PivotFilterOperator::DoesNotBeginWith => "captionNotBeginsWith",
        PivotFilterOperator::EndsWith => "captionEndsWith",
        PivotFilterOperator::DoesNotEndWith => "captionNotEndsWith",
        PivotFilterOperator::Contains => "captionContains",
        PivotFilterOperator::DoesNotContain => "captionNotContains",
    }
}

fn value_filter_type_name(operator: PivotFilterOperator) -> Option<&'static str> {
    Some(match operator {
        PivotFilterOperator::Equals => "valueEqual",
        PivotFilterOperator::NotEquals => "valueNotEqual",
        PivotFilterOperator::LessThan => "valueLessThan",
        PivotFilterOperator::LessThanOrEqual => "valueLessThanOrEqual",
        PivotFilterOperator::GreaterThan => "valueGreaterThan",
        PivotFilterOperator::GreaterThanOrEqual => "valueGreaterThanOrEqual",
        PivotFilterOperator::BeginsWith
        | PivotFilterOperator::DoesNotBeginWith
        | PivotFilterOperator::EndsWith
        | PivotFilterOperator::DoesNotEndWith
        | PivotFilterOperator::Contains
        | PivotFilterOperator::DoesNotContain => return None,
    })
}

fn date_filter_type_name(operator: PivotFilterOperator) -> Option<&'static str> {
    Some(match operator {
        PivotFilterOperator::Equals => "dateEqual",
        PivotFilterOperator::NotEquals => "dateNotEqual",
        PivotFilterOperator::LessThan => "dateOlderThan",
        PivotFilterOperator::LessThanOrEqual => "dateOlderThanOrEqual",
        PivotFilterOperator::GreaterThan => "dateNewerThan",
        PivotFilterOperator::GreaterThanOrEqual => "dateNewerThanOrEqual",
        PivotFilterOperator::BeginsWith
        | PivotFilterOperator::DoesNotBeginWith
        | PivotFilterOperator::EndsWith
        | PivotFilterOperator::DoesNotEndWith
        | PivotFilterOperator::Contains
        | PivotFilterOperator::DoesNotContain => return None,
    })
}

fn date_period_filter_type_name(period: PivotDatePeriod) -> Option<&'static str> {
    Some(match period {
        PivotDatePeriod::Tomorrow => "tomorrow",
        PivotDatePeriod::Today => "today",
        PivotDatePeriod::Yesterday => "yesterday",
        PivotDatePeriod::NextWeek => "nextWeek",
        PivotDatePeriod::ThisWeek => "thisWeek",
        PivotDatePeriod::LastWeek => "lastWeek",
        PivotDatePeriod::NextMonth => "nextMonth",
        PivotDatePeriod::ThisMonth => "thisMonth",
        PivotDatePeriod::LastMonth => "lastMonth",
        PivotDatePeriod::NextQuarter => "nextQuarter",
        PivotDatePeriod::ThisQuarter => "thisQuarter",
        PivotDatePeriod::LastQuarter => "lastQuarter",
        PivotDatePeriod::NextYear => "nextYear",
        PivotDatePeriod::ThisYear => "thisYear",
        PivotDatePeriod::LastYear => "lastYear",
        PivotDatePeriod::YearToDate => "yearToDate",
        PivotDatePeriod::Quarter(1) => "Q1",
        PivotDatePeriod::Quarter(2) => "Q2",
        PivotDatePeriod::Quarter(3) => "Q3",
        PivotDatePeriod::Quarter(4) => "Q4",
        PivotDatePeriod::Month(1) => "M1",
        PivotDatePeriod::Month(2) => "M2",
        PivotDatePeriod::Month(3) => "M3",
        PivotDatePeriod::Month(4) => "M4",
        PivotDatePeriod::Month(5) => "M5",
        PivotDatePeriod::Month(6) => "M6",
        PivotDatePeriod::Month(7) => "M7",
        PivotDatePeriod::Month(8) => "M8",
        PivotDatePeriod::Month(9) => "M9",
        PivotDatePeriod::Month(10) => "M10",
        PivotDatePeriod::Month(11) => "M11",
        PivotDatePeriod::Month(12) => "M12",
        PivotDatePeriod::Month(_) | PivotDatePeriod::Quarter(_) => return None,
    })
}

fn top_n_filter_type_name(_top: bool, percent: bool) -> &'static str {
    match percent {
        true => "percent",
        false => "count",
    }
}

fn custom_filter_operator(operator: PivotFilterOperator) -> &'static str {
    match operator {
        PivotFilterOperator::Equals
        | PivotFilterOperator::BeginsWith
        | PivotFilterOperator::EndsWith
        | PivotFilterOperator::Contains => "equal",
        PivotFilterOperator::NotEquals
        | PivotFilterOperator::DoesNotBeginWith
        | PivotFilterOperator::DoesNotEndWith
        | PivotFilterOperator::DoesNotContain => "notEqual",
        PivotFilterOperator::LessThan => "lessThan",
        PivotFilterOperator::LessThanOrEqual => "lessThanOrEqual",
        PivotFilterOperator::GreaterThan => "greaterThan",
        PivotFilterOperator::GreaterThanOrEqual => "greaterThanOrEqual",
    }
}

fn label_custom_filter_value(operator: PivotFilterOperator, value: &str) -> String {
    match operator {
        PivotFilterOperator::BeginsWith | PivotFilterOperator::DoesNotBeginWith => {
            format!("{value}*")
        }
        PivotFilterOperator::EndsWith | PivotFilterOperator::DoesNotEndWith => {
            format!("*{value}")
        }
        PivotFilterOperator::Contains | PivotFilterOperator::DoesNotContain => {
            format!("*{value}*")
        }
        _ => value.to_string(),
    }
}

fn write_data_field_ext_with_source(
    w: &mut XmlWriter,
    pivot_show_as: &str,
    source_field: Option<usize>,
) -> XlsxResult<()> {
    w.write_event(Event::Start(BytesStart::new("extLst")))?;

    let mut ext = BytesStart::new("ext");
    ext.push_attribute(("uri", EXT_URI_X14_DATA_FIELD));
    w.write_event(Event::Start(ext))?;

    let mut data_field = BytesStart::new("x14:dataField");
    data_field.push_attribute(("xmlns:x14", NS_SPREADSHEET_X14));
    data_field.push_attribute(("pivotShowAs", pivot_show_as));
    let source_field = source_field.map(|field| field.to_string());
    if let Some(source_field) = source_field.as_deref() {
        data_field.push_attribute(("sourceField", source_field));
    }
    w.write_event(Event::Empty(data_field))?;

    w.write_event(Event::End(BytesEnd::new("ext")))?;
    w.write_event(Event::End(BytesEnd::new("extLst")))?;
    Ok(())
}

fn write_pivot_extensions(w: &mut XmlWriter, pivot: &PivotTable) -> XlsxResult<()> {
    if pivot
        .extensions
        .iter()
        .all(|extension| extension.payload.is_empty())
    {
        return Ok(());
    }

    w.write_event(Event::Start(BytesStart::new("extLst")))?;
    for extension in &pivot.extensions {
        if !extension.payload.is_empty() {
            w.get_mut().write_all(&extension.payload)?;
        }
    }
    w.write_event(Event::End(BytesEnd::new("extLst")))?;
    Ok(())
}

fn show_data_as_name(show_as: &PivotShowAs) -> Option<&'static str> {
    match show_as {
        PivotShowAs::Normal => None,
        PivotShowAs::PercentOfGrandTotal => Some("percentOfTotal"),
        PivotShowAs::PercentOfRowTotal => Some("percentOfRow"),
        PivotShowAs::PercentOfColumnTotal => Some("percentOfCol"),
        PivotShowAs::PercentOfParentRowTotal
        | PivotShowAs::PercentOfParentColumnTotal
        | PivotShowAs::PercentOfParentTotal { .. } => None,
        PivotShowAs::Index => Some("index"),
        PivotShowAs::RunningTotal { .. } => Some("runTotal"),
        PivotShowAs::DifferenceFrom { .. } => Some("difference"),
        PivotShowAs::PercentDifferenceFrom { .. } => Some("percentDiff"),
        PivotShowAs::RankAscending { .. } | PivotShowAs::RankDescending { .. } => None,
    }
}

fn x14_show_as_name(show_as: &PivotShowAs) -> Option<&'static str> {
    match show_as {
        PivotShowAs::PercentOfParentTotal { .. } => Some("percentOfParent"),
        PivotShowAs::PercentOfParentRowTotal => Some("percentOfParentRow"),
        PivotShowAs::PercentOfParentColumnTotal => Some("percentOfParentCol"),
        PivotShowAs::RankAscending { .. } => Some("rankAscending"),
        PivotShowAs::RankDescending { .. } => Some("rankDescending"),
        _ => None,
    }
}

fn x14_show_as_source_field_index(
    show_as: &PivotShowAs,
    fields: &[CacheField],
) -> XlsxResult<Option<usize>> {
    let PivotShowAs::PercentOfParentTotal { base_field } = show_as else {
        return Ok(None);
    };
    let base_field = &base_field.name;
    field_index(fields, base_field).map(Some).ok_or_else(|| {
        XlsxError::InvalidFormat(format!("pivot base field not found: {base_field}"))
    })
}

fn show_as_base_field_index(
    show_as: &PivotShowAs,
    fields: &[CacheField],
) -> XlsxResult<Option<usize>> {
    let base_field = match show_as {
        PivotShowAs::RunningTotal { base_field }
        | PivotShowAs::RankAscending { base_field }
        | PivotShowAs::RankDescending { base_field }
        | PivotShowAs::DifferenceFrom { base_field, .. }
        | PivotShowAs::PercentDifferenceFrom { base_field, .. } => &base_field.name,
        _ => return Ok(None),
    };
    field_index(fields, base_field).map(Some).ok_or_else(|| {
        XlsxError::InvalidFormat(format!("pivot base field not found: {base_field}"))
    })
}

fn show_as_base_item_index(
    show_as: &PivotShowAs,
    fields: &[CacheField],
) -> XlsxResult<Option<u32>> {
    let (base_field, base_item) = match show_as {
        PivotShowAs::DifferenceFrom {
            base_field,
            base_item,
        }
        | PivotShowAs::PercentDifferenceFrom {
            base_field,
            base_item,
        } => (&base_field.name, base_item),
        _ => return Ok(None),
    };
    let field_index = field_index(fields, base_field).ok_or_else(|| {
        XlsxError::InvalidFormat(format!("pivot base field not found: {base_field}"))
    })?;
    fields[field_index]
        .shared_items
        .iter()
        .position(|item| item == base_item)
        .map(|index| index as u32)
        .map(Some)
        .ok_or_else(|| {
            XlsxError::InvalidFormat(format!(
                "pivot base item not found in field {base_field}: {base_item}"
            ))
        })
}

fn write_pivot_style(w: &mut XmlWriter, pivot: &PivotTable) -> XlsxResult<()> {
    let mut style = BytesStart::new("pivotTableStyleInfo");
    if let Some(name) = &pivot.style.name {
        style.push_attribute(("name", name.as_str()));
    }
    style.push_attribute(("showRowHeaders", bool_attr(pivot.style.show_row_headers)));
    style.push_attribute(("showColHeaders", bool_attr(pivot.style.show_column_headers)));
    style.push_attribute(("showRowStripes", bool_attr(pivot.style.show_row_stripes)));
    style.push_attribute(("showColStripes", bool_attr(pivot.style.show_column_stripes)));
    style.push_attribute(("showLastColumn", bool_attr(pivot.style.show_last_column)));
    w.write_event(Event::Empty(style))?;
    Ok(())
}

pub(super) fn write_pivot_cache_definition_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    workbook: &Workbook,
    plan: &FormatPivotPlan,
    part: &FormatPivotCache,
) -> XlsxResult<()> {
    let fields = xlsx_cache_fields(part);
    let path = format!("xl/pivotCache/pivotCacheDefinition{}.xml", part.cache_num);
    write_xml_part(zip, &path, |w| {
        let record_count = part.row_count.to_string();
        let refresh_on_load = bool_attr(part.refresh_on_load);
        let background_query = bool_attr(part.background_query);
        let save_data = bool_attr(part.save_data);
        let mut tag = BytesStart::new("pivotCacheDefinition");
        tag.push_attribute(("xmlns", NS_SPREADSHEET));
        tag.push_attribute(("xmlns:r", NS_DOC_RELS));
        tag.push_attribute(("r:id", "rId1"));
        tag.push_attribute(("recordCount", record_count.as_str()));
        tag.push_attribute(("saveData", save_data));
        tag.push_attribute(("refreshOnLoad", refresh_on_load));
        tag.push_attribute(("backgroundQuery", background_query));
        let missing_items_limit = part.missing_items_limit.map(|limit| limit.to_string());
        if let Some(missing_items_limit) = missing_items_limit.as_deref() {
            tag.push_attribute(("missingItemsLimit", missing_items_limit));
        }
        tag.push_attribute(("createdVersion", "8"));
        tag.push_attribute(("refreshedVersion", "8"));
        tag.push_attribute(("minRefreshableVersion", "3"));
        w.write_event(Event::Start(tag))?;

        write_cache_source(w, workbook, part)?;

        let count = fields.len().to_string();
        let mut cache_fields = BytesStart::new("cacheFields");
        cache_fields.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(cache_fields))?;
        for (index, field) in fields.iter().enumerate() {
            write_cache_field(
                w,
                &fields,
                index,
                field,
                index < part.fields.len()
                    && cache_field_shared_items_are_metadata_only(workbook, plan, part, index),
            )?;
        }
        w.write_event(Event::End(BytesEnd::new("cacheFields")))?;

        write_calculated_items(w, part)?;

        w.write_event(Event::End(BytesEnd::new("pivotCacheDefinition")))?;
        Ok(())
    })
}

fn cache_source_tag(source_type: &'static str) -> BytesStart<'static> {
    let mut cache_source = BytesStart::new("cacheSource");
    cache_source.push_attribute(("type", source_type));
    cache_source
}

fn write_cache_source(
    w: &mut XmlWriter,
    workbook: &Workbook,
    part: &FormatPivotCache,
) -> XlsxResult<()> {
    match &part.source {
        FormatPivotSource::Worksheet { .. } => {
            w.write_event(Event::Start(cache_source_tag("worksheet")))?;
            write_worksheet_source(w, workbook, part)?;
            w.write_event(Event::End(BytesEnd::new("cacheSource")))?;
        }
        FormatPivotSource::External {
            connection_name, ..
        } => {
            let mut cache_source = cache_source_tag("external");
            if let Some(connection_id) = connection_id_attr(workbook, connection_name)? {
                cache_source.push_attribute(("connectionId", connection_id.as_str()));
            }
            w.write_event(Event::Empty(cache_source))?;
        }
        FormatPivotSource::Consolidation { ranges } => {
            w.write_event(Event::Start(cache_source_tag("consolidation")))?;
            write_consolidation_source(w, ranges)?;
            w.write_event(Event::End(BytesEnd::new("cacheSource")))?;
        }
        FormatPivotSource::Scenario { name } => {
            if name.is_empty() {
                w.write_event(Event::Empty(cache_source_tag("scenario")))?;
            } else {
                w.write_event(Event::Start(cache_source_tag("scenario")))?;
                let mut worksheet_source = BytesStart::new("worksheetSource");
                worksheet_source.push_attribute(("name", name.as_str()));
                w.write_event(Event::Empty(worksheet_source))?;
                w.write_event(Event::End(BytesEnd::new("cacheSource")))?;
            }
        }
        FormatPivotSource::Olap {
            connection_name, ..
        } => {
            let mut cache_source = cache_source_tag("olap");
            if let Some(connection_id) = connection_id_attr(workbook, connection_name)? {
                cache_source.push_attribute(("connectionId", connection_id.as_str()));
            }
            w.write_event(Event::Empty(cache_source))?;
        }
    }
    Ok(())
}

fn connection_id_attr(workbook: &Workbook, connection_name: &str) -> XlsxResult<Option<String>> {
    if connection_name.is_empty() {
        return Ok(None);
    }
    if let Some(connection) = find_workbook_connection(workbook, connection_name) {
        return Ok(Some(connection.id.to_string()));
    }
    let connection_id = connection_name.parse::<u32>().map_err(|_| {
        XlsxError::InvalidFormat(format!(
            "XLSX pivot external source connectionId must be numeric or match a workbook data connection: {connection_name}"
        ))
    })?;
    Ok(Some(connection_id.to_string()))
}

fn find_workbook_connection<'a>(
    workbook: &'a Workbook,
    connection_name: &str,
) -> Option<&'a WorkbookConnection> {
    connection_name
        .parse::<u32>()
        .ok()
        .and_then(|id| workbook.data_connection_by_id(id))
        .or_else(|| workbook.data_connection_by_name(connection_name))
}

fn write_consolidation_source(w: &mut XmlWriter, ranges: &[PivotSourceRange]) -> XlsxResult<()> {
    if ranges.is_empty() {
        return Err(XlsxError::InvalidFormat(
            "XLSX consolidation pivot sources require at least one range".into(),
        ));
    }
    let pages = consolidation_pages(ranges)?;

    let mut consolidation = BytesStart::new("consolidation");
    consolidation.push_attribute(("autoPage", "0"));
    w.write_event(Event::Start(consolidation))?;

    if !pages.is_empty() {
        let count = pages.len().to_string();
        let mut pages_el = BytesStart::new("pages");
        pages_el.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(pages_el))?;
        for page in &pages {
            let count = page.len().to_string();
            let mut page_el = BytesStart::new("page");
            page_el.push_attribute(("count", count.as_str()));
            w.write_event(Event::Start(page_el))?;
            for item in page {
                let mut page_item = BytesStart::new("pageItem");
                page_item.push_attribute(("name", item.as_str()));
                w.write_event(Event::Empty(page_item))?;
            }
            w.write_event(Event::End(BytesEnd::new("page")))?;
        }
        w.write_event(Event::End(BytesEnd::new("pages")))?;
    }

    let count = ranges.len().to_string();
    let mut range_sets = BytesStart::new("rangeSets");
    range_sets.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(range_sets))?;
    let mut external_index = 0usize;
    for range in ranges {
        if range.external_relationship_id.is_some() && range.external_relationship_target.is_none()
        {
            return Err(XlsxError::InvalidFormat(
                "XLSX external consolidation references require a relationship target".into(),
            ));
        }
        if range.range.is_none() && range.name.is_none() {
            return Err(XlsxError::InvalidFormat(
                "XLSX consolidation rangeSet requires a range or name".into(),
            ));
        }
        let ref_str = range.range.map(|range| range.to_a1_string());
        let mut range_set = BytesStart::new("rangeSet");
        if let Some(ref_str) = ref_str.as_deref() {
            range_set.push_attribute(("ref", ref_str));
        }
        if let Some(sheet) = &range.sheet {
            range_set.push_attribute(("sheet", sheet.as_str()));
        }
        if let Some(name) = &range.name {
            range_set.push_attribute(("name", name.as_str()));
        }
        let external_relationship_id = if range.external_relationship_target.is_some() {
            external_index += 1;
            Some(consolidation_external_relationship_id(
                range,
                external_index,
            ))
        } else {
            None
        };
        if let Some(external_relationship_id) = external_relationship_id.as_deref() {
            range_set.push_attribute(("r:id", external_relationship_id));
        }
        for (index, item) in range.page_items.iter().enumerate() {
            let page = pages.get(index).ok_or_else(|| {
                XlsxError::InvalidFormat(
                    "XLSX consolidation page item has no matching page field".into(),
                )
            })?;
            let Some(item_index) = page.iter().position(|candidate| candidate == item) else {
                return Err(XlsxError::InvalidFormat(format!(
                    "XLSX consolidation page item is not declared: {item}"
                )));
            };
            let attr_name = match index {
                0 => "i1",
                1 => "i2",
                2 => "i3",
                3 => "i4",
                _ => unreachable!("consolidation_pages rejects more than four page fields"),
            };
            let item_index = item_index.to_string();
            range_set.push_attribute((attr_name, item_index.as_str()));
        }
        w.write_event(Event::Empty(range_set))?;
    }
    w.write_event(Event::End(BytesEnd::new("rangeSets")))?;

    w.write_event(Event::End(BytesEnd::new("consolidation")))?;
    Ok(())
}

fn consolidation_pages(ranges: &[PivotSourceRange]) -> XlsxResult<Vec<Vec<String>>> {
    let page_count = ranges
        .iter()
        .map(|range| range.page_items.len())
        .max()
        .unwrap_or(0);
    if page_count > 4 {
        return Err(XlsxError::InvalidFormat(
            "XLSX consolidation pivot sources support at most four page fields".into(),
        ));
    }

    let mut pages = vec![Vec::<String>::new(); page_count];
    for range in ranges {
        for (index, item) in range.page_items.iter().enumerate() {
            if item.trim().is_empty() {
                return Err(XlsxError::InvalidFormat(
                    "XLSX consolidation page item names cannot be blank".into(),
                ));
            }
            if !pages[index].iter().any(|candidate| candidate == item) {
                pages[index].push(item.clone());
            }
        }
    }
    Ok(pages)
}

fn consolidation_external_relationship_id(
    range: &PivotSourceRange,
    external_index: usize,
) -> String {
    range
        .external_relationship_id
        .clone()
        .unwrap_or_else(|| format!("rIdExternal{external_index}"))
}

fn write_worksheet_source(
    w: &mut XmlWriter,
    _workbook: &Workbook,
    part: &FormatPivotCache,
) -> XlsxResult<()> {
    let mut source = BytesStart::new("worksheetSource");
    match &part.source {
        FormatPivotSource::Worksheet {
            sheet_name,
            range,
            table_name,
            ..
        } => {
            if let Some(name) = table_name {
                source.push_attribute(("name", name.as_str()));
                w.write_event(Event::Empty(source))?;
                return Ok(());
            }
            let ref_str = range.to_a1_string();
            source.push_attribute(("ref", ref_str.as_str()));
            source.push_attribute(("sheet", sheet_name.as_str()));
        }
        _ => {}
    }
    w.write_event(Event::Empty(source))?;
    Ok(())
}

fn write_calculated_items(w: &mut XmlWriter, part: &FormatPivotCache) -> XlsxResult<()> {
    if part.calculated_items.is_empty() {
        return Ok(());
    }

    let count = part.calculated_items.len().to_string();
    let mut calculated_items = BytesStart::new("calculatedItems");
    calculated_items.push_attribute(("count", count.as_str()));
    w.write_event(Event::Start(calculated_items))?;

    for item in &part.calculated_items {
        write_calculated_item(w, &part.fields, item)?;
    }

    w.write_event(Event::End(BytesEnd::new("calculatedItems")))?;
    Ok(())
}

fn write_calculated_item(
    w: &mut XmlWriter,
    fields: &[CacheField],
    item: &PivotCalculatedItem,
) -> XlsxResult<()> {
    let field_index = field_index(fields, &item.field.name).ok_or_else(|| {
        XlsxError::InvalidFormat(format!(
            "pivot calculated item references unknown source field: {}",
            item.field.name
        ))
    })?;
    let item_index = fields[field_index]
        .shared_items
        .iter()
        .position(|candidate| candidate == &item.item)
        .map(|index| index as u32)
        .ok_or_else(|| {
            XlsxError::InvalidFormat(format!(
                "pivot calculated item {} was not registered in field {}",
                item.item, item.field.name
            ))
        })?;

    let field_index = field_index.to_string();
    let item_index = item_index.to_string();
    let formula = formula_for_cache_attr(&item.formula);

    let mut calculated_item = BytesStart::new("calculatedItem");
    calculated_item.push_attribute(("field", field_index.as_str()));
    calculated_item.push_attribute(("formula", formula.as_str()));
    w.write_event(Event::Start(calculated_item))?;

    let mut pivot_area = BytesStart::new("pivotArea");
    pivot_area.push_attribute(("field", field_index.as_str()));
    pivot_area.push_attribute(("cacheIndex", "1"));
    w.write_event(Event::Start(pivot_area))?;

    let mut references = BytesStart::new("references");
    references.push_attribute(("count", "1"));
    w.write_event(Event::Start(references))?;

    let mut reference = BytesStart::new("reference");
    reference.push_attribute(("field", field_index.as_str()));
    reference.push_attribute(("count", "1"));
    w.write_event(Event::Start(reference))?;

    let mut x = BytesStart::new("x");
    x.push_attribute(("v", item_index.as_str()));
    w.write_event(Event::Empty(x))?;

    w.write_event(Event::End(BytesEnd::new("reference")))?;
    w.write_event(Event::End(BytesEnd::new("references")))?;
    w.write_event(Event::End(BytesEnd::new("pivotArea")))?;
    w.write_event(Event::End(BytesEnd::new("calculatedItem")))?;
    Ok(())
}

fn write_cache_field(
    w: &mut XmlWriter,
    fields: &[CacheField],
    field_index: usize,
    field: &CacheField,
    metadata_only: bool,
) -> XlsxResult<()> {
    let mut cache_field = BytesStart::new("cacheField");
    cache_field.push_attribute(("name", field.name.as_str()));
    if field
        .grouping
        .as_ref()
        .is_some_and(|grouping| matches!(grouping.definition, PivotGrouping::Manual { .. }))
    {
        cache_field.push_attribute(("numFmtId", "0"));
    }
    if let Some(formula) = &field.formula {
        let formula = formula_for_cache_attr(formula);
        cache_field.push_attribute(("formula", formula.as_str()));
    }
    if !field.database_field {
        cache_field.push_attribute(("databaseField", "0"));
    }
    w.write_event(Event::Start(cache_field))?;

    if !(field
        .grouping
        .as_ref()
        .is_some_and(|grouping| matches!(grouping.definition, PivotGrouping::Manual { .. }))
        && !field.database_field)
    {
        let mut shared_items = BytesStart::new("sharedItems");
        let count = field.shared_items.len().to_string();
        if !metadata_only {
            shared_items.push_attribute(("count", count.as_str()));
        }
        shared_items.push_attribute(("containsBlank", bool_attr(field_contains_blank(field))));
        shared_items.push_attribute(("containsString", bool_attr(field_contains_string(field))));
        shared_items.push_attribute(("containsNumber", bool_attr(field_contains_number(field))));
        shared_items.push_attribute(("containsMixedTypes", bool_attr(field_contains_mixed(field))));
        let (contains_integer, min_value, max_value) = numeric_shared_item_stats(field);
        if let Some(contains_integer) = contains_integer {
            shared_items.push_attribute(("containsInteger", bool_attr(contains_integer)));
        }
        if let Some(min_value) = min_value.as_deref() {
            shared_items.push_attribute(("minValue", min_value));
        }
        if let Some(max_value) = max_value.as_deref() {
            shared_items.push_attribute(("maxValue", max_value));
        }
        w.write_event(Event::Start(shared_items))?;
        if !metadata_only {
            for value in &field.shared_items {
                write_pivot_value(w, value)?;
            }
        }
        w.write_event(Event::End(BytesEnd::new("sharedItems")))?;
    }

    if let Some(grouping) = &field.grouping {
        match &grouping.definition {
            PivotGrouping::Manual { .. } if field.database_field => {
                let parent = fields
                    .iter()
                    .enumerate()
                    .find(|(_, candidate)| {
                        !candidate.database_field
                            && candidate
                                .grouping
                                .as_ref()
                                .is_some_and(|candidate_grouping| {
                                    matches!(
                                        candidate_grouping.definition,
                                        PivotGrouping::Manual { .. }
                                    ) && candidate_grouping.base_field_index == field_index
                                })
                    })
                    .map(|(index, _)| index)
                    .ok_or_else(|| {
                        XlsxError::InvalidFormat("manual pivot grouping field is missing".into())
                    })?;
                let mut field_group = BytesStart::new("fieldGroup");
                let parent = parent.to_string();
                field_group.push_attribute(("par", parent.as_str()));
                w.write_event(Event::Empty(field_group))?;
            }
            PivotGrouping::Manual { .. } => write_field_group(w, grouping, None)?,
            PivotGrouping::Date { units, .. } if units.len() > 1 => {}
            _ => write_field_group(w, grouping, None)?,
        }
    } else if let Some((grouping, level)) = grouping_level_for_field(fields, field_index) {
        write_field_group(w, grouping, Some(level))?;
    }

    w.write_event(Event::End(BytesEnd::new("cacheField")))?;
    Ok(())
}

fn cache_field_shared_items_are_metadata_only(
    workbook: &Workbook,
    plan: &FormatPivotPlan,
    cache: &FormatPivotCache,
    field_index: usize,
) -> bool {
    let field = &cache.fields[field_index];
    if field.grouping.is_some()
        || field.shared_items.is_empty()
        || !field
            .shared_items
            .iter()
            .all(|value| matches!(value, PivotValue::Number(_)))
    {
        return false;
    }
    let mut used_as_measure = false;
    for table in plan
        .tables
        .iter()
        .filter(|table| table.cache_num == cache.cache_num)
    {
        let Some(pivot) = workbook
            .worksheet(table.sheet_index)
            .and_then(|sheet| sheet.pivot_tables().get(table.pivot_index))
        else {
            return false;
        };
        let explicit = pivot
            .rows
            .iter()
            .chain(&pivot.columns)
            .chain(&pivot.page_fields)
            .map(|field| field.field.name.as_str())
            .chain(
                pivot
                    .filters
                    .iter()
                    .filter_map(filter_field_ref)
                    .map(|field| field.name.as_str()),
            )
            .chain(pivot.groupings.iter().map(|grouping| match grouping {
                PivotGrouping::Number { field, .. }
                | PivotGrouping::Date { field, .. }
                | PivotGrouping::Manual { field, .. } => field.name.as_str(),
            }))
            .any(|name| cache.field_index(name) == Some(field_index));
        if explicit {
            return false;
        }
        used_as_measure |= pivot
            .measures
            .iter()
            .any(|measure| cache.field_index(&measure.field.name) == Some(field_index));
    }
    used_as_measure
}

fn numeric_shared_item_stats(field: &CacheField) -> (Option<bool>, Option<String>, Option<String>) {
    let mut numbers = field.shared_items.iter().filter_map(|value| match value {
        PivotValue::Number(value) if value.is_finite() => Some(*value),
        _ => None,
    });
    let Some(first) = numbers.next() else {
        return (None, None, None);
    };

    let mut min = first;
    let mut max = first;
    let mut contains_integer = first.fract() == 0.0;
    for value in numbers {
        min = min.min(value);
        max = max.max(value);
        contains_integer &= value.fract() == 0.0;
    }

    (
        Some(contains_integer),
        Some(min.to_string()),
        Some(max.to_string()),
    )
}

fn grouping_level_for_field(
    fields: &[CacheField],
    field_index: usize,
) -> Option<(
    &FormatPivotGrouping,
    &duke_sheets_pivot::plan::FormatPivotGroupLevel,
)> {
    fields
        .iter()
        .filter_map(|field| field.grouping.as_ref())
        .find_map(|grouping| {
            grouping
                .levels
                .iter()
                .find(|level| {
                    level.field_index == field_index && field_index != grouping.base_field_index
                })
                .map(|level| (grouping, level))
        })
}

fn write_field_group(
    w: &mut XmlWriter,
    grouping: &FormatPivotGrouping,
    level: Option<&duke_sheets_pivot::plan::FormatPivotGroupLevel>,
) -> XlsxResult<()> {
    let mut field_group = BytesStart::new("fieldGroup");
    let base = (level.is_some() || matches!(grouping.definition, PivotGrouping::Manual { .. }))
        .then(|| grouping.base_field_index.to_string());
    let parent = level.and_then(|level| level.parent_field_index.map(|parent| parent.to_string()));
    if let Some(base) = base.as_deref() {
        field_group.push_attribute(("base", base));
    }
    if let Some(parent) = parent.as_deref() {
        field_group.push_attribute(("par", parent));
    }

    w.write_event(Event::Start(field_group))?;

    if matches!(grouping.definition, PivotGrouping::Manual { .. }) {
        let level = &grouping.levels[0];
        let item_indexes = &level.source_item_group_ids;
        let group_items = &level.group_items;
        let count = item_indexes.len().to_string();
        let mut discrete_pr = BytesStart::new("discretePr");
        discrete_pr.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(discrete_pr))?;
        for index in item_indexes {
            let value = index.to_string();
            let mut x = BytesStart::new("x");
            x.push_attribute(("v", value.as_str()));
            w.write_event(Event::Empty(x))?;
        }
        w.write_event(Event::End(BytesEnd::new("discretePr")))?;

        let count = group_items.len().to_string();
        let mut group_items_el = BytesStart::new("groupItems");
        group_items_el.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(group_items_el))?;
        for value in group_items {
            write_pivot_value(w, value)?;
        }
        w.write_event(Event::End(BytesEnd::new("groupItems")))?;
        w.write_event(Event::End(BytesEnd::new("fieldGroup")))?;
        return Ok(());
    }

    let mut range_pr = BytesStart::new("rangePr");
    match (&grouping.definition, level) {
        (
            PivotGrouping::Number {
                start,
                end,
                interval,
                ..
            },
            _,
        ) => {
            let auto_start = bool_attr(start.is_none());
            let auto_end = bool_attr(end.is_none());
            let start_num = start.map(|value| value.to_string());
            let end_num = end.map(|value| value.to_string());
            let group_interval = interval.to_string();

            range_pr.push_attribute(("autoStart", auto_start));
            range_pr.push_attribute(("autoEnd", auto_end));
            range_pr.push_attribute(("groupBy", "range"));
            if let Some(start_num) = &start_num {
                range_pr.push_attribute(("startNum", start_num.as_str()));
            }
            if let Some(end_num) = &end_num {
                range_pr.push_attribute(("endNum", end_num.as_str()));
            }
            range_pr.push_attribute(("groupInterval", group_interval.as_str()));
        }
        (PivotGrouping::Date { units, .. }, None) => {
            range_pr.push_attribute(("groupBy", date_group_by_name(units[0])));
        }
        (PivotGrouping::Date { .. }, Some(level)) => {
            range_pr.push_attribute((
                "groupBy",
                date_group_by_name(level.date_unit.ok_or_else(|| {
                    XlsxError::InvalidFormat("date grouping level has no unit".into())
                })?),
            ));
        }
        (PivotGrouping::Manual { .. }, _) => unreachable!("manual groups return before rangePr"),
    }
    w.write_event(Event::Empty(range_pr))?;

    w.write_event(Event::End(BytesEnd::new("fieldGroup")))?;
    Ok(())
}

fn date_group_by_name(unit: PivotDateGroupUnit) -> &'static str {
    match unit {
        PivotDateGroupUnit::Seconds => "seconds",
        PivotDateGroupUnit::Minutes => "minutes",
        PivotDateGroupUnit::Hours => "hours",
        PivotDateGroupUnit::Days => "days",
        PivotDateGroupUnit::Months => "months",
        PivotDateGroupUnit::Quarters => "quarters",
        PivotDateGroupUnit::Years => "years",
    }
}

pub(super) fn write_pivot_cache_records_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    workbook: &Workbook,
    plan: &FormatPivotPlan,
    part: &FormatPivotCache,
) -> XlsxResult<()> {
    let fields = xlsx_cache_fields(part);
    let path = format!("xl/pivotCache/pivotCacheRecords{}.xml", part.cache_num);
    write_xml_part(zip, &path, |w| {
        let count = part.row_count.to_string();
        let mut records = BytesStart::new("pivotCacheRecords");
        records.push_attribute(("xmlns", NS_SPREADSHEET));
        records.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(records))?;

        for row in 0..part.row_count {
            w.write_event(Event::Start(BytesStart::new("r")))?;
            for (index, field) in fields.iter().take(part.fields.len()).enumerate() {
                write_cache_record_value(
                    w,
                    field,
                    field.item_ids.get(row).copied(),
                    cache_field_shared_items_are_metadata_only(workbook, plan, part, index),
                )?;
            }
            w.write_event(Event::End(BytesEnd::new("r")))?;
        }

        w.write_event(Event::End(BytesEnd::new("pivotCacheRecords")))?;
        Ok(())
    })
}

fn write_cache_record_value(
    w: &mut XmlWriter,
    field: &CacheField,
    value_index: Option<u32>,
    metadata_only: bool,
) -> XlsxResult<()> {
    let Some(index) = value_index else {
        w.write_event(Event::Empty(BytesStart::new("m")))?;
        return Ok(());
    };

    let Some(pivot_value) = field.shared_items.get(index as usize) else {
        w.write_event(Event::Empty(BytesStart::new("m")))?;
        return Ok(());
    };
    if metadata_only {
        return write_pivot_value(w, pivot_value);
    }

    let value = index.to_string();
    let mut x = BytesStart::new("x");
    x.push_attribute(("v", value.as_str()));
    w.write_event(Event::Empty(x))?;
    Ok(())
}

pub(super) fn consolidation_external_relationships(
    source: &FormatPivotSource,
) -> XlsxResult<Vec<(String, String)>> {
    let FormatPivotSource::Consolidation { ranges } = source else {
        return Ok(Vec::new());
    };

    let mut relationships = Vec::new();
    for range in ranges {
        let Some(target) = &range.external_relationship_target else {
            continue;
        };
        if target.trim().is_empty() {
            return Err(XlsxError::InvalidFormat(
                "XLSX external consolidation relationship target cannot be blank".into(),
            ));
        }
        let id = consolidation_external_relationship_id(range, relationships.len() + 1);
        relationships.push((id, target.clone()));
    }
    Ok(relationships)
}

fn write_pivot_value(w: &mut XmlWriter, value: &PivotValue) -> XlsxResult<()> {
    match value {
        PivotValue::Blank => {
            w.write_event(Event::Empty(BytesStart::new("m")))?;
        }
        PivotValue::Boolean(value) => {
            let mut tag = BytesStart::new("b");
            tag.push_attribute(("v", if *value { "1" } else { "0" }));
            w.write_event(Event::Empty(tag))?;
        }
        PivotValue::Number(value) => {
            let number = value.to_string();
            let mut tag = BytesStart::new("n");
            tag.push_attribute(("v", number.as_str()));
            w.write_event(Event::Empty(tag))?;
        }
        PivotValue::String(value) => {
            let mut tag = BytesStart::new("s");
            tag.push_attribute(("v", value.as_str()));
            w.write_event(Event::Empty(tag))?;
        }
        PivotValue::Error(value) => {
            let mut tag = BytesStart::new("e");
            tag.push_attribute(("v", value.as_str()));
            w.write_event(Event::Empty(tag))?;
        }
    }
    Ok(())
}

fn field_contains_blank(field: &CacheField) -> bool {
    field
        .shared_items
        .iter()
        .any(|value| matches!(value, PivotValue::Blank))
}

fn field_contains_string(field: &CacheField) -> bool {
    field
        .shared_items
        .iter()
        .any(|value| matches!(value, PivotValue::String(_)))
}

fn field_contains_number(field: &CacheField) -> bool {
    field
        .shared_items
        .iter()
        .any(|value| matches!(value, PivotValue::Number(_)))
}

fn field_contains_mixed(field: &CacheField) -> bool {
    let mut kinds = std::collections::HashSet::new();
    for value in &field.shared_items {
        kinds.insert(std::mem::discriminant(value));
    }
    kinds.len() > 1
}

fn bool_attr(value: bool) -> &'static str {
    if value {
        "1"
    } else {
        "0"
    }
}

fn push_bool_attr_if(tag: &mut BytesStart<'_>, name: &'static str, value: bool, default: bool) {
    if value != default {
        tag.push_attribute((name, bool_attr(value)));
    }
}

fn field_index(fields: &[CacheField], name: &str) -> Option<usize> {
    fields
        .iter()
        .position(|field| field.name.eq_ignore_ascii_case(name))
}

fn aggregate_name(aggregate: PivotAggregate) -> &'static str {
    match aggregate {
        PivotAggregate::Average => "average",
        PivotAggregate::Count => "count",
        PivotAggregate::CountNumbers => "countNums",
        PivotAggregate::Max => "max",
        PivotAggregate::Min => "min",
        PivotAggregate::Product => "product",
        PivotAggregate::StdDev => "stdDev",
        PivotAggregate::StdDevP => "stdDevp",
        PivotAggregate::Sum => "sum",
        PivotAggregate::Var => "var",
        PivotAggregate::VarP => "varp",
    }
}
