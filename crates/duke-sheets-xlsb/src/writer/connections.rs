use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::biff12::{encode_wide_str, records, RecordWriter};
use crate::error::{XlsbError, XlsbResult};
use duke_sheets_core::{
    Workbook, WorkbookConnection, WorkbookConnectionCredentials, WorkbookConnectionKind,
    WorkbookConnectionParameter, WorkbookConnectionParameterType, WorkbookConnectionParameterValue,
};

pub(crate) const CT_CONNECTIONS: &str = "application/vnd.ms-excel.connections";
pub(crate) const RT_CONNECTIONS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";

pub(crate) fn write_connections<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    workbook: &Workbook,
) -> XlsbResult<()> {
    if workbook.data_connections().is_empty() {
        return Ok(());
    }

    zip.start_file("xl/connections.bin", *options)?;
    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    rw.write_record(records::BRT_BEGIN_EXT_CONNECTIONS, &[])?;
    for connection in workbook.data_connections() {
        write_connection(&mut rw, connection)?;
    }
    rw.write_record(records::BRT_END_EXT_CONNECTIONS, &[])?;

    drop(rw);
    zip.write_all(&buf)?;
    Ok(())
}

fn write_connection<W: Write>(
    rw: &mut RecordWriter<W>,
    connection: &WorkbookConnection,
) -> XlsbResult<()> {
    rw.write_record(
        records::BRT_BEGIN_EXT_CONNECTION,
        &ext_connection_payload(connection)?,
    )?;

    match &connection.kind {
        WorkbookConnectionKind::Database {
            connection: connection_string,
            command,
            command_type,
        } => {
            write_db_props(
                rw,
                connection_string,
                command.as_deref(),
                command_type.unwrap_or(2),
            )?;
        }
        WorkbookConnectionKind::Olap {
            connection: connection_string,
            command,
            command_type,
            local,
            local_connection,
            local_refresh,
            send_locale,
            row_drill_count,
        } => {
            if connection_string.is_some() || command.is_some() || command_type.is_some() {
                write_db_props(
                    rw,
                    connection_string.as_deref().unwrap_or(""),
                    command.as_deref(),
                    command_type.unwrap_or(1),
                )?;
            }
            write_olap_props(
                rw,
                *local,
                local_connection.as_deref(),
                *local_refresh,
                *send_locale,
                *row_drill_count,
            )?;
        }
        WorkbookConnectionKind::Web {
            url,
            xml,
            source_data,
            html_tables,
            html_format,
            post,
            edit_page,
        } => {
            write_web_props(
                rw,
                url.as_deref(),
                *xml,
                *source_data,
                *html_tables,
                html_format.as_deref(),
                post.as_deref(),
                edit_page.as_deref(),
            )?;
        }
        WorkbookConnectionKind::Text {
            source_file,
            delimiter,
            first_row,
            delimited,
            decimal,
            thousands,
        } => {
            write_text_props(
                rw,
                source_file.as_deref().unwrap_or(""),
                delimiter.as_deref(),
                *first_row,
                *delimited,
                decimal.as_deref(),
                thousands.as_deref(),
            )?;
        }
    }

    if !connection.parameters.is_empty() {
        rw.write_record(
            records::BRT_BEGIN_EC_PARAMS,
            &checked_u32(connection.parameters.len(), "connection parameter count")?.to_le_bytes(),
        )?;
        for parameter in &connection.parameters {
            rw.write_record(records::BRT_BEGIN_EC_PARAM, &parameter_payload(parameter)?)?;
            rw.write_record(records::BRT_END_EC_PARAM, &[])?;
        }
        rw.write_record(records::BRT_END_EC_PARAMS, &[])?;
    }

    rw.write_record(records::BRT_END_EXT_CONNECTION, &[])?;
    Ok(())
}

fn ext_connection_payload(connection: &WorkbookConnection) -> XlsbResult<Vec<u8>> {
    let mut flags1 = 0u16;
    if connection.keep_alive {
        flags1 |= 0x0001;
    }
    if connection.new_connection {
        flags1 |= 0x0002;
    }
    if connection.deleted {
        flags1 |= 0x0004;
    }
    if connection.only_use_connection_file {
        flags1 |= 0x0008;
    }
    if connection.background {
        flags1 |= 0x0010;
    }
    if connection.refresh_on_load {
        flags1 |= 0x0020;
    }
    if connection.save_data {
        flags1 |= 0x0040;
    }

    // MS-XLSB 2.4.80 BrtBeginExtConnection: bit K is reserved but MUST be 1.
    let mut load_flags = 0x0008u16;
    if connection.source_file.is_some() {
        load_flags |= 0x0001;
    }
    if connection.odc_file.is_some() {
        load_flags |= 0x0002;
    }
    if connection.description.is_some() {
        load_flags |= 0x0004;
    }
    if connection.single_sign_on_id.is_some() {
        load_flags |= 0x0010;
    }

    let mut payload = Vec::new();
    payload.push(connection.refreshed_version);
    payload.push(connection.min_refreshable_version);
    // MS-XLSB 2.4.80: pc=0x02 means the password is not saved.
    payload.push(if connection.save_password { 0x01 } else { 0x02 });
    payload.push(0);
    payload.extend_from_slice(
        &checked_u16(connection.interval, "connection refresh interval")?.to_le_bytes(),
    );
    payload.extend_from_slice(&flags1.to_le_bytes());
    payload.extend_from_slice(&load_flags.to_le_bytes());
    payload.extend_from_slice(&connection_type(connection).to_le_bytes());
    payload.extend_from_slice(&connection.reconnection_method.to_le_bytes());
    payload.extend_from_slice(&connection.id.to_le_bytes());
    payload.push(credentials_byte(connection.credentials));
    if let Some(source_file) = &connection.source_file {
        payload.extend_from_slice(&encode_wide_str(source_file));
    }
    if let Some(odc_file) = &connection.odc_file {
        payload.extend_from_slice(&encode_wide_str(odc_file));
    }
    if let Some(description) = &connection.description {
        payload.extend_from_slice(&encode_wide_str(description));
    }
    payload.extend_from_slice(&encode_wide_str(&connection.name));
    if let Some(single_sign_on_id) = &connection.single_sign_on_id {
        payload.extend_from_slice(&encode_wide_str(single_sign_on_id));
    }
    Ok(payload)
}

fn write_db_props<W: Write>(
    rw: &mut RecordWriter<W>,
    connection: &str,
    command: Option<&str>,
    command_type: u32,
) -> XlsbResult<()> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&command_type.to_le_bytes());
    payload.push(if command.is_some() { 0x02 } else { 0x00 });
    payload.extend_from_slice(&encode_wide_str(connection));
    if let Some(command) = command {
        payload.extend_from_slice(&encode_wide_str(command));
    }
    rw.write_record(records::BRT_BEGIN_EC_DB_PROPS, &payload)?;
    rw.write_record(records::BRT_END_EC_DB_PROPS, &[])?;
    Ok(())
}

fn write_olap_props<W: Write>(
    rw: &mut RecordWriter<W>,
    local: bool,
    local_connection: Option<&str>,
    local_refresh: bool,
    send_locale: bool,
    row_drill_count: Option<u32>,
) -> XlsbResult<()> {
    let mut flags = 0u8;
    if local {
        flags |= 0x01;
    }
    if !local_refresh {
        flags |= 0x02;
    }
    if send_locale {
        flags |= 0x40;
    }
    let mut payload = Vec::new();
    payload.push(flags);
    payload.extend_from_slice(&row_drill_count.unwrap_or(0).to_le_bytes());
    payload.push(if local_connection.is_some() {
        0x01
    } else {
        0x00
    });
    if let Some(local_connection) = local_connection {
        payload.extend_from_slice(&encode_wide_str(local_connection));
    }
    rw.write_record(records::BRT_BEGIN_EC_OLAP_PROPS, &payload)?;
    rw.write_record(records::BRT_END_EC_OLAP_PROPS, &[])?;
    Ok(())
}

fn write_web_props<W: Write>(
    rw: &mut RecordWriter<W>,
    url: Option<&str>,
    xml: bool,
    source_data: bool,
    html_tables: bool,
    html_format: Option<&str>,
    post: Option<&str>,
    edit_page: Option<&str>,
) -> XlsbResult<()> {
    let mut flags = 0u16;
    if xml {
        flags |= 0x0001;
    }
    if source_data {
        flags |= 0x0002;
    }
    if html_tables {
        flags |= 0x0100;
    }
    let mut load_flags = 0u8;
    if post.is_some() {
        load_flags |= 0x01;
    }
    if edit_page.is_some() {
        load_flags |= 0x02;
    }
    if url.is_some() {
        load_flags |= 0x04;
    }

    let mut payload = Vec::new();
    payload.push(html_format_code(html_format));
    payload.extend_from_slice(&flags.to_le_bytes());
    payload.push(0);
    payload.push(load_flags);
    if let Some(url) = url {
        payload.extend_from_slice(&encode_wide_str(url));
    }
    if let Some(post) = post {
        payload.extend_from_slice(&encode_wide_str(post));
    }
    if let Some(edit_page) = edit_page {
        payload.extend_from_slice(&encode_wide_str(edit_page));
    }
    rw.write_record(records::BRT_BEGIN_EC_WEB_PROPS, &payload)?;
    rw.write_record(records::BRT_END_EC_WEB_PROPS, &[])?;
    Ok(())
}

fn write_text_props<W: Write>(
    rw: &mut RecordWriter<W>,
    source_file: &str,
    delimiter: Option<&str>,
    first_row: u32,
    delimited: bool,
    decimal: Option<&str>,
    thousands: Option<&str>,
) -> XlsbResult<()> {
    const ICPID_WINDOWS_ANSI: u32 = 0x0000_0001;
    const F_DELIMITED: u32 = 1 << 12;
    const F_TAB: u32 = 1 << 13;
    const F_SPACE: u32 = 1 << 14;
    const F_COMMA: u32 = 1 << 15;
    const F_SEMICOLON: u32 = 1 << 16;
    const RESERVED1_MUST_BE_1: u32 = 1 << 20;
    const F_CUSTOM_DELIMITER: u32 = 1 << 22;

    let mut header = ICPID_WINDOWS_ANSI | RESERVED1_MUST_BE_1;
    let mut custom_delimiter = 0u16;
    if delimited {
        header |= F_DELIMITED;
        match delimiter {
            Some("\t") => header |= F_TAB,
            Some(" ") => header |= F_SPACE,
            Some(",") => header |= F_COMMA,
            Some(";") => header |= F_SEMICOLON,
            Some(value) => {
                if let Some(code) = single_char_code(value) {
                    header |= F_CUSTOM_DELIMITER;
                    custom_delimiter = code;
                }
            }
            None => {}
        }
    }

    let mut payload = Vec::new();
    payload.extend_from_slice(&header.to_le_bytes());
    payload.extend_from_slice(&custom_delimiter.to_le_bytes());
    payload.extend_from_slice(&first_row.to_le_bytes());
    payload.push(latin1_byte(decimal));
    payload.push(latin1_byte(thousands));
    payload.extend_from_slice(&encode_wide_str(source_file));
    rw.write_record(records::BRT_BEGIN_EC_TXT_WIZ, &payload)?;
    rw.write_record(records::BRT_BEGIN_EC_TW_FLD_INFO_LST, &1u32.to_le_bytes())?;
    rw.write_record(records::BRT_BEGIN_EC_TW_FLD_INFO, &[0; 8])?;
    rw.write_record(records::BRT_END_EC_TW_FLD_INFO_LST, &[])?;
    rw.write_record(records::BRT_END_EC_TXT_WIZ, &[])?;
    Ok(())
}

fn parameter_payload(parameter: &WorkbookConnectionParameter) -> XlsbResult<Vec<u8>> {
    let mut payload = Vec::new();
    let mut header = match parameter.parameter_type {
        WorkbookConnectionParameterType::Prompt => 0u16,
        WorkbookConnectionParameterType::Value => 1u16,
        WorkbookConnectionParameterType::Cell => 2u16,
    };
    if parameter.refresh_on_change {
        header |= 0x0008;
    }
    payload.extend_from_slice(&header.to_le_bytes());
    payload.extend_from_slice(
        &checked_u16_signed(parameter.sql_type, "connection parameter SQL type")?.to_le_bytes(),
    );

    match parameter.parameter_type {
        WorkbookConnectionParameterType::Prompt => {
            payload.extend_from_slice(&u32::from(parameter.prompt.is_some()).to_le_bytes());
            payload.extend_from_slice(&encode_wide_str(parameter.name.as_deref().unwrap_or("")));
            if let Some(prompt) = &parameter.prompt {
                payload.extend_from_slice(&encode_wide_str(prompt));
            }
        }
        WorkbookConnectionParameterType::Value => {
            let data_type = parameter_value_type(&parameter.value)?;
            payload.extend_from_slice(&data_type.to_le_bytes());
            payload.extend_from_slice(&encode_wide_str(parameter.name.as_deref().unwrap_or("")));
            write_parameter_value(&mut payload, &parameter.value)?;
        }
        WorkbookConnectionParameterType::Cell => {
            return Err(XlsbError::InvalidFormat(
                "XLSB connection cell-bound parameter authoring is not supported yet".into(),
            ));
        }
    }
    Ok(payload)
}

fn write_parameter_value(
    payload: &mut Vec<u8>,
    value: &WorkbookConnectionParameterValue,
) -> XlsbResult<()> {
    match value {
        WorkbookConnectionParameterValue::None => {}
        WorkbookConnectionParameterValue::Double(value) => {
            payload.extend_from_slice(&value.to_le_bytes());
        }
        WorkbookConnectionParameterValue::String(value) => {
            payload.extend_from_slice(&encode_wide_str(value));
        }
        WorkbookConnectionParameterValue::Boolean(value) => {
            payload.push(u8::from(*value));
        }
        WorkbookConnectionParameterValue::Integer(value) => {
            payload.extend_from_slice(&(*value as f64).to_le_bytes());
        }
        WorkbookConnectionParameterValue::Cell(_) => {
            return Err(XlsbError::InvalidFormat(
                "XLSB connection cell references require parameterType=cell".into(),
            ));
        }
    }
    Ok(())
}

fn parameter_value_type(value: &WorkbookConnectionParameterValue) -> XlsbResult<u32> {
    Ok(match value {
        WorkbookConnectionParameterValue::None => 0,
        WorkbookConnectionParameterValue::Double(_) => 1,
        WorkbookConnectionParameterValue::String(_) => 2,
        WorkbookConnectionParameterValue::Boolean(_) => 4,
        WorkbookConnectionParameterValue::Integer(_) => 0x0800,
        WorkbookConnectionParameterValue::Cell(_) => {
            return Err(XlsbError::InvalidFormat(
                "XLSB connection cell references require parameterType=cell".into(),
            ));
        }
    })
}

fn connection_type(connection: &WorkbookConnection) -> u32 {
    connection
        .connection_type
        .unwrap_or(match &connection.kind {
            WorkbookConnectionKind::Database { .. } | WorkbookConnectionKind::Olap { .. } => 5,
            WorkbookConnectionKind::Web { .. } => 4,
            WorkbookConnectionKind::Text { .. } => 6,
        })
}

fn credentials_byte(credentials: Option<WorkbookConnectionCredentials>) -> u8 {
    match credentials {
        Some(WorkbookConnectionCredentials::Integrated) => 0,
        Some(WorkbookConnectionCredentials::None) => 1,
        Some(WorkbookConnectionCredentials::Stored) => 2,
        Some(WorkbookConnectionCredentials::Prompt) => 3,
        None => 1,
    }
}

fn html_format_code(html_format: Option<&str>) -> u8 {
    match html_format {
        Some("none") => 0,
        Some("rtf") => 1,
        Some("all") => 2,
        _ => 2,
    }
}

fn latin1_byte(value: Option<&str>) -> u8 {
    value
        .and_then(|text| text.chars().next())
        .and_then(|ch| (ch as u32 <= u8::MAX as u32).then_some(ch as u8))
        .unwrap_or(0)
}

fn single_char_code(value: &str) -> Option<u16> {
    let mut chars = value.chars();
    let ch = chars.next()?;
    if chars.next().is_some() {
        return None;
    }
    (ch as u32 <= u16::MAX as u32).then_some(ch as u16)
}

fn checked_u16(value: u32, label: &str) -> XlsbResult<u16> {
    u16::try_from(value)
        .map_err(|_| XlsbError::InvalidFormat(format!("{label} exceeds BIFF12 u16 range")))
}

fn checked_u16_signed(value: i32, label: &str) -> XlsbResult<u16> {
    if !(0..=u16::MAX as i32).contains(&value) {
        return Err(XlsbError::InvalidFormat(format!(
            "{label} exceeds BIFF12 u16 range"
        )));
    }
    Ok(value as u16)
}

fn checked_u32(value: usize, label: &str) -> XlsbResult<u32> {
    u32::try_from(value)
        .map_err(|_| XlsbError::InvalidFormat(format!("{label} exceeds BIFF12 u32 range")))
}
