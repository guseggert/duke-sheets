use std::io::{Read, Seek};

use duke_sheets_core::{
    WorkbookConnection, WorkbookConnectionCredentials, WorkbookConnectionKind,
    WorkbookConnectionParameter, WorkbookConnectionParameterType, WorkbookConnectionParameterValue,
};
use duke_sheets_formula::decompile::{decompiler, FormulaContext};

use crate::biff12::{parser, records, token_parser, RecordIter};
use crate::error::{XlsbError, XlsbResult};

pub(crate) fn read_connections<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: Option<&str>,
    formula_ctx: &FormulaContext,
) -> XlsbResult<Vec<WorkbookConnection>> {
    let Some(path) = path else {
        return Ok(Vec::new());
    };
    let file = match archive.by_name(path) {
        Ok(file) => file,
        Err(_) => return Ok(Vec::new()),
    };

    let mut iter = RecordIter::new(file);
    let mut buf = Vec::with_capacity(1024);
    let mut connections = Vec::new();
    let mut current: Option<ParsedConnection> = None;

    loop {
        let record = iter.next_record(&mut buf);
        let (record_type, len) = match record {
            Ok(record) => record,
            Err(XlsbError::Parse(message)) if message.contains("unexpected end") => break,
            Err(err) => return Err(err),
        };
        let payload = &buf[..len];
        match record_type {
            records::BRT_BEGIN_EXT_CONNECTION => {
                if let Some(connection) = current.take().and_then(ParsedConnection::build) {
                    connections.push(connection);
                }
                current = parse_ext_connection(payload)?;
            }
            records::BRT_BEGIN_EC_DB_PROPS => {
                if let Some(connection) = &mut current {
                    connection.apply_db_props(payload)?;
                }
            }
            records::BRT_BEGIN_EC_OLAP_PROPS => {
                if let Some(connection) = &mut current {
                    connection.apply_olap_props(payload)?;
                }
            }
            records::BRT_BEGIN_EC_WEB_PROPS => {
                if let Some(connection) = &mut current {
                    connection.kind = parse_web_props(payload)?;
                }
            }
            records::BRT_BEGIN_EC_TXT_WIZ => {
                if let Some(connection) = &mut current {
                    connection.kind = parse_text_props(payload)?;
                }
            }
            records::BRT_BEGIN_EC_PARAM => {
                if let Some(connection) = &mut current {
                    if let Some(parameter) = parse_parameter(payload, formula_ctx)? {
                        connection.parameters.push(parameter);
                    }
                }
            }
            _ => {}
        }
    }

    if let Some(connection) = current.and_then(ParsedConnection::build) {
        connections.push(connection);
    }

    Ok(connections)
}

#[derive(Debug)]
struct ParsedConnection {
    id: u32,
    name: String,
    source_file: Option<String>,
    odc_file: Option<String>,
    description: Option<String>,
    connection_type: Option<u32>,
    refreshed_version: u8,
    min_refreshable_version: u8,
    keep_alive: bool,
    interval: u32,
    reconnection_method: u32,
    refresh_on_load: bool,
    background: bool,
    save_data: bool,
    save_password: bool,
    new_connection: bool,
    deleted: bool,
    only_use_connection_file: bool,
    credentials: Option<WorkbookConnectionCredentials>,
    single_sign_on_id: Option<String>,
    kind: Option<WorkbookConnectionKind>,
    db_props: Option<ConnectionDbProps>,
    parameters: Vec<WorkbookConnectionParameter>,
}

impl ParsedConnection {
    fn apply_db_props(&mut self, payload: &[u8]) -> XlsbResult<()> {
        let Some(db_props) = parse_db_props(payload)? else {
            return Ok(());
        };
        self.kind = Some(match self.kind.take() {
            Some(WorkbookConnectionKind::Olap {
                local,
                local_connection,
                local_refresh,
                send_locale,
                row_drill_count,
                ..
            }) => db_props.olap_kind(
                local,
                local_connection,
                local_refresh,
                send_locale,
                row_drill_count,
            ),
            _ => db_props.database_kind(),
        });
        self.db_props = Some(db_props);
        Ok(())
    }

    fn apply_olap_props(&mut self, payload: &[u8]) -> XlsbResult<()> {
        self.kind = parse_olap_props(payload, self.db_props.as_ref())?;
        Ok(())
    }

    fn build(mut self) -> Option<WorkbookConnection> {
        if self.id == 0 || self.name.trim().is_empty() {
            return None;
        }
        let kind = self
            .kind
            .take()
            .unwrap_or_else(|| fallback_kind(self.connection_type, self.source_file.clone()));
        Some(WorkbookConnection {
            id: self.id,
            name: self.name,
            source_file: self.source_file,
            odc_file: self.odc_file,
            description: self.description,
            connection_type: self.connection_type,
            kind,
            refreshed_version: self.refreshed_version,
            min_refreshable_version: self.min_refreshable_version,
            keep_alive: self.keep_alive,
            interval: self.interval,
            reconnection_method: self.reconnection_method,
            refresh_on_load: self.refresh_on_load,
            background: self.background,
            save_data: self.save_data,
            save_password: self.save_password,
            new_connection: self.new_connection,
            deleted: self.deleted,
            only_use_connection_file: self.only_use_connection_file,
            credentials: self.credentials,
            single_sign_on_id: self.single_sign_on_id,
            parameters: self.parameters,
        })
    }
}

fn parse_ext_connection(payload: &[u8]) -> XlsbResult<Option<ParsedConnection>> {
    if payload.len() < 23 {
        return Ok(None);
    }

    let flags1 = parser::read_u16(payload, 6);
    let flags2 = parser::read_u16(payload, 8);
    let idbtype = parser::read_u32(payload, 10);
    let mut offset = 23usize;

    let source_file = read_optional_wide(payload, &mut offset, flags2 & 0x0001 != 0)?;
    let odc_file = read_optional_wide(payload, &mut offset, flags2 & 0x0002 != 0)?;
    let description = read_optional_wide(payload, &mut offset, flags2 & 0x0004 != 0)?;
    let name = read_wide(payload, &mut offset)?;
    let single_sign_on_id = read_optional_wide(payload, &mut offset, flags2 & 0x0010 != 0)?;

    Ok(Some(ParsedConnection {
        id: parser::read_u32(payload, 18),
        name,
        source_file,
        odc_file,
        description,
        connection_type: Some(idbtype),
        refreshed_version: payload[0],
        min_refreshable_version: payload[1],
        keep_alive: flags1 & 0x0001 != 0,
        interval: parser::read_u16(payload, 4) as u32,
        reconnection_method: parser::read_u32(payload, 14),
        refresh_on_load: flags1 & 0x0020 != 0,
        background: flags1 & 0x0010 != 0,
        save_data: flags1 & 0x0040 != 0,
        save_password: payload[2] == 0x01,
        new_connection: flags1 & 0x0002 != 0,
        deleted: flags1 & 0x0004 != 0,
        only_use_connection_file: flags1 & 0x0008 != 0,
        credentials: credentials_from_byte(payload[22]),
        single_sign_on_id,
        kind: None,
        db_props: None,
        parameters: Vec::new(),
    }))
}

#[derive(Debug, Clone)]
struct ConnectionDbProps {
    connection: String,
    command: Option<String>,
    command_type: Option<u32>,
}

impl ConnectionDbProps {
    fn database_kind(&self) -> WorkbookConnectionKind {
        WorkbookConnectionKind::Database {
            connection: self.connection.clone(),
            command: self.command.clone(),
            command_type: self.command_type,
        }
    }

    fn olap_kind(
        &self,
        local: bool,
        local_connection: Option<String>,
        local_refresh: bool,
        send_locale: bool,
        row_drill_count: Option<u32>,
    ) -> WorkbookConnectionKind {
        WorkbookConnectionKind::Olap {
            connection: Some(self.connection.clone()),
            command: self.command.clone(),
            command_type: self.command_type,
            local,
            local_connection,
            local_refresh,
            send_locale,
            row_drill_count,
        }
    }
}

fn parse_db_props(payload: &[u8]) -> XlsbResult<Option<ConnectionDbProps>> {
    if payload.len() < 5 {
        return Ok(None);
    }
    let command_type = parser::read_u32(payload, 0);
    let flags = payload[4];
    let mut offset = 5usize;
    let connection = read_wide(payload, &mut offset)?;
    let command = read_optional_wide(payload, &mut offset, flags & 0x02 != 0)?;
    let _server_command = read_optional_wide(payload, &mut offset, flags & 0x01 != 0)?;
    Ok(Some(ConnectionDbProps {
        connection,
        command,
        command_type: Some(command_type),
    }))
}

fn parse_olap_props(
    payload: &[u8],
    db_props: Option<&ConnectionDbProps>,
) -> XlsbResult<Option<WorkbookConnectionKind>> {
    if payload.len() < 6 {
        return Ok(None);
    }
    let flags = payload[0];
    let row_drill_count = parser::read_u32(payload, 1);
    let mut offset = 6usize;
    let local_connection = read_optional_wide(payload, &mut offset, payload[5] & 0x01 != 0)?;
    let local = flags & 0x01 != 0;
    let local_refresh = flags & 0x02 == 0;
    let send_locale = flags & 0x40 != 0;
    let row_drill_count = (row_drill_count > 0).then_some(row_drill_count);
    Ok(Some(if let Some(db_props) = db_props {
        db_props.olap_kind(
            local,
            local_connection,
            local_refresh,
            send_locale,
            row_drill_count,
        )
    } else {
        WorkbookConnectionKind::Olap {
            connection: None,
            command: None,
            command_type: None,
            local,
            local_connection,
            local_refresh,
            send_locale,
            row_drill_count,
        }
    }))
}

fn parse_web_props(payload: &[u8]) -> XlsbResult<Option<WorkbookConnectionKind>> {
    if payload.len() < 5 {
        return Ok(None);
    }
    let html_format = match payload[0] {
        0 => Some("none".to_string()),
        1 => Some("rtf".to_string()),
        2 => Some("all".to_string()),
        _ => None,
    };
    let flags = parser::read_u16(payload, 1);
    let load_flags = payload[4];
    let mut offset = 5usize;
    let url = read_optional_wide(payload, &mut offset, load_flags & 0x04 != 0)?;
    let post = read_optional_wide(payload, &mut offset, load_flags & 0x01 != 0)?;
    let edit_page = read_optional_wide(payload, &mut offset, load_flags & 0x02 != 0)?;
    Ok(Some(WorkbookConnectionKind::Web {
        url,
        xml: flags & 0x0001 != 0,
        source_data: flags & 0x0002 != 0,
        html_tables: flags & 0x0100 != 0,
        html_format,
        post,
        edit_page,
    }))
}

fn parse_text_props(payload: &[u8]) -> XlsbResult<Option<WorkbookConnectionKind>> {
    if payload.len() < 12 {
        return Ok(None);
    }
    let header = parser::read_u32(payload, 0);
    let custom = parser::read_u16(payload, 4);
    let first_row = parser::read_u32(payload, 6);
    let decimal = char_from_latin1(payload[10]);
    let thousands = char_from_latin1(payload[11]);
    let mut offset = 12usize;
    let source_file = Some(read_wide(payload, &mut offset)?);
    Ok(Some(WorkbookConnectionKind::Text {
        source_file,
        delimiter: text_delimiter(header, custom),
        first_row,
        delimited: header & (1 << 12) != 0,
        decimal,
        thousands,
    }))
}

fn parse_parameter(
    payload: &[u8],
    formula_ctx: &FormulaContext,
) -> XlsbResult<Option<WorkbookConnectionParameter>> {
    if payload.len() < 4 {
        return Ok(None);
    }
    let header = parser::read_u16(payload, 0);
    let parameter_kind = header & 0x0007;
    let refresh_on_change = header & 0x0008 != 0;
    let sql_type = parser::read_u16(payload, 2) as i32;
    let mut offset = 4usize;
    let mut data_type = None;
    let mut load_prompt = false;
    match parameter_kind {
        0 => {
            if payload.len() < offset + 4 {
                return Ok(None);
            }
            load_prompt = parser::read_u32(payload, offset) != 0;
            offset += 4;
        }
        1 | 2 => {
            if payload.len() < offset + 4 {
                return Ok(None);
            }
            data_type = Some(parser::read_u32(payload, offset));
            offset += 4;
        }
        _ => return Ok(None),
    }

    let name = non_empty(read_wide(payload, &mut offset)?);
    let prompt = read_optional_wide(payload, &mut offset, parameter_kind == 0 && load_prompt)?;
    let value = match parameter_kind {
        0 => WorkbookConnectionParameterValue::None,
        1 => parse_parameter_literal(payload, &mut offset, data_type.unwrap_or(0))?,
        2 => parse_parameter_formula(payload, offset, formula_ctx),
        _ => WorkbookConnectionParameterValue::None,
    };
    let parameter_type = match parameter_kind {
        1 => WorkbookConnectionParameterType::Value,
        2 => WorkbookConnectionParameterType::Cell,
        _ => WorkbookConnectionParameterType::Prompt,
    };

    Ok(Some(WorkbookConnectionParameter {
        name,
        sql_type,
        parameter_type,
        refresh_on_change,
        prompt,
        value,
    }))
}

fn parse_parameter_literal(
    payload: &[u8],
    offset: &mut usize,
    data_type: u32,
) -> XlsbResult<WorkbookConnectionParameterValue> {
    Ok(match data_type {
        1 => {
            if payload.len() < *offset + 8 {
                return Ok(WorkbookConnectionParameterValue::None);
            }
            let value = parser::read_f64(payload, *offset);
            *offset += 8;
            WorkbookConnectionParameterValue::Double(value)
        }
        2 => WorkbookConnectionParameterValue::String(read_wide(payload, offset)?),
        4 => {
            if payload.len() <= *offset {
                return Ok(WorkbookConnectionParameterValue::None);
            }
            let value = payload[*offset] != 0;
            *offset += 1;
            WorkbookConnectionParameterValue::Boolean(value)
        }
        0x800 => {
            if payload.len() < *offset + 8 {
                return Ok(WorkbookConnectionParameterValue::None);
            }
            let value = parser::read_f64(payload, *offset) as i32;
            *offset += 8;
            WorkbookConnectionParameterValue::Integer(value)
        }
        0x8000 => {
            if payload.len() < *offset + 8 {
                return Ok(WorkbookConnectionParameterValue::None);
            }
            let value = parser::read_f64(payload, *offset) as i32;
            *offset += 8;
            WorkbookConnectionParameterValue::Integer(value)
        }
        _ => WorkbookConnectionParameterValue::None,
    })
}

fn parse_parameter_formula(
    payload: &[u8],
    offset: usize,
    formula_ctx: &FormulaContext,
) -> WorkbookConnectionParameterValue {
    if offset + 4 > payload.len() {
        return WorkbookConnectionParameterValue::None;
    }
    let cce = parser::read_u32(payload, offset) as usize;
    let tokens_start = offset + 4;
    let tokens_end = tokens_start + cce;
    if cce == 0 || tokens_end > payload.len() {
        return WorkbookConnectionParameterValue::None;
    }
    let cb_offset = tokens_end;
    let extra = if cb_offset + 4 <= payload.len() {
        let cb = parser::read_u32(payload, cb_offset) as usize;
        let extra_start = cb_offset + 4;
        let extra_end = extra_start + cb;
        if cb > 0 && extra_end <= payload.len() {
            &payload[extra_start..extra_end]
        } else {
            &[] as &[u8]
        }
    } else {
        &[] as &[u8]
    };
    let tokens = token_parser::parse_tokens_with_extra(&payload[tokens_start..tokens_end], extra);
    if tokens.is_empty() {
        return WorkbookConnectionParameterValue::None;
    }
    let formula = decompiler::decompile(&tokens, formula_ctx);
    if formula.is_empty() {
        WorkbookConnectionParameterValue::None
    } else {
        WorkbookConnectionParameterValue::Cell(formula)
    }
}

fn read_optional_wide(
    payload: &[u8],
    offset: &mut usize,
    present: bool,
) -> XlsbResult<Option<String>> {
    if present {
        read_wide(payload, offset).map(Some)
    } else {
        Ok(None)
    }
}

fn read_wide(payload: &[u8], offset: &mut usize) -> XlsbResult<String> {
    let (value, consumed) = parser::wide_str(payload, *offset)?;
    *offset += consumed;
    Ok(value)
}

fn non_empty(value: String) -> Option<String> {
    (!value.is_empty()).then_some(value)
}

fn char_from_latin1(value: u8) -> Option<String> {
    (value != 0).then(|| char::from(value).to_string())
}

fn text_delimiter(header: u32, custom: u16) -> Option<String> {
    if header & (1 << 12) == 0 {
        return None;
    }
    if header & (1 << 22) != 0 {
        return char::from_u32(custom as u32).map(|ch| ch.to_string());
    }
    if header & (1 << 15) != 0 {
        Some(",".to_string())
    } else if header & (1 << 16) != 0 {
        Some(";".to_string())
    } else if header & (1 << 13) != 0 {
        Some("\t".to_string())
    } else if header & (1 << 14) != 0 {
        Some(" ".to_string())
    } else {
        None
    }
}

fn credentials_from_byte(value: u8) -> Option<WorkbookConnectionCredentials> {
    match value {
        0 => Some(WorkbookConnectionCredentials::Integrated),
        1 => Some(WorkbookConnectionCredentials::None),
        2 => Some(WorkbookConnectionCredentials::Stored),
        3 => Some(WorkbookConnectionCredentials::Prompt),
        _ => None,
    }
}

fn fallback_kind(
    connection_type: Option<u32>,
    source_file: Option<String>,
) -> WorkbookConnectionKind {
    match connection_type {
        Some(4) => WorkbookConnectionKind::Web {
            url: source_file,
            xml: false,
            source_data: false,
            html_tables: false,
            html_format: None,
            post: None,
            edit_page: None,
        },
        Some(6 | 103) => WorkbookConnectionKind::Text {
            source_file,
            delimiter: None,
            first_row: 1,
            delimited: true,
            decimal: None,
            thousands: None,
        },
        _ => WorkbookConnectionKind::Database {
            connection: String::new(),
            command: None,
            command_type: None,
        },
    }
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use super::*;
    use crate::biff12::{encode_wide_str, RecordWriter};

    // features: Data connections (basic)
    #[test]
    fn reads_xlsb_database_connection_records() {
        let connections = read_test_connections(database_connections_bin());
        assert_eq!(connections.len(), 1);
        let connection = &connections[0];
        assert_eq!(connection.id, 7);
        assert_eq!(connection.name, "SalesConnection");
        assert_eq!(connection.source_file.as_deref(), Some("sales.csv"));
        assert_eq!(connection.odc_file.as_deref(), Some("sales.odc"));
        assert_eq!(connection.description.as_deref(), Some("Sales Reporting"));
        assert_eq!(connection.connection_type, Some(5));
        assert_eq!(connection.refreshed_version, 3);
        assert_eq!(connection.min_refreshable_version, 0);
        assert!(connection.keep_alive);
        assert_eq!(connection.interval, 15);
        assert_eq!(connection.reconnection_method, 2);
        assert!(connection.refresh_on_load);
        assert!(connection.background);
        assert!(connection.save_data);
        assert!(connection.save_password);
        assert_eq!(
            connection.credentials,
            Some(WorkbookConnectionCredentials::Stored)
        );
        assert_eq!(connection.single_sign_on_id.as_deref(), Some("SalesSso"));
        match &connection.kind {
            WorkbookConnectionKind::Database {
                connection,
                command,
                command_type,
            } => {
                assert_eq!(connection, "Provider=MSDASQL;DSN=Sales;");
                assert_eq!(command.as_deref(), Some("select * from Sales"));
                assert_eq!(*command_type, Some(2));
            }
            other => panic!("unexpected connection kind: {other:?}"),
        }
    }

    #[test]
    fn reads_xlsb_web_connection_records_and_parameters() {
        let connections = read_test_connections(web_connections_bin());
        assert_eq!(connections.len(), 1);
        let connection = &connections[0];
        assert_eq!(connection.id, 8);
        assert_eq!(connection.name, "WebSales");
        assert_eq!(connection.connection_type, Some(4));
        match &connection.kind {
            WorkbookConnectionKind::Web {
                url,
                xml,
                source_data,
                html_tables,
                html_format,
                post,
                edit_page,
            } => {
                assert_eq!(url.as_deref(), Some("http://127.0.0.1/duke-sheets/sales"));
                assert!(*xml);
                assert!(*source_data);
                assert!(*html_tables);
                assert_eq!(html_format.as_deref(), Some("all"));
                assert_eq!(post.as_deref(), Some("region=west"));
                assert_eq!(
                    edit_page.as_deref(),
                    Some("http://127.0.0.1/duke-sheets/edit")
                );
            }
            other => panic!("unexpected connection kind: {other:?}"),
        }
        assert_eq!(connection.parameters.len(), 2);
        assert_eq!(connection.parameters[0].name.as_deref(), Some("Region"));
        assert_eq!(
            connection.parameters[0].parameter_type,
            WorkbookConnectionParameterType::Prompt
        );
        assert_eq!(
            connection.parameters[0].prompt.as_deref(),
            Some("Choose region")
        );
        assert_eq!(connection.parameters[1].name.as_deref(), Some("Limit"));
        assert_eq!(
            connection.parameters[1].parameter_type,
            WorkbookConnectionParameterType::Value
        );
        assert_eq!(
            connection.parameters[1].value,
            WorkbookConnectionParameterValue::Integer(25)
        );
    }

    #[test]
    fn reads_xlsb_text_connection_records() {
        let connections = read_test_connections(text_connections_bin());
        assert_eq!(connections.len(), 1);
        let connection = &connections[0];
        assert_eq!(connection.id, 9);
        assert_eq!(connection.name, "CsvSales");
        assert_eq!(connection.connection_type, Some(6));
        match &connection.kind {
            WorkbookConnectionKind::Text {
                source_file,
                delimiter,
                first_row,
                delimited,
                decimal,
                thousands,
            } => {
                assert_eq!(source_file.as_deref(), Some("/data/sales.csv"));
                assert_eq!(delimiter.as_deref(), Some(","));
                assert_eq!(*first_row, 2);
                assert!(*delimited);
                assert_eq!(decimal.as_deref(), Some("."));
                assert_eq!(thousands.as_deref(), Some(","));
            }
            other => panic!("unexpected connection kind: {other:?}"),
        }
    }

    #[test]
    fn reads_xlsb_olap_connection_records() {
        let connections = read_test_connections(olap_connections_bin());
        assert_eq!(connections.len(), 1);
        let connection = &connections[0];
        assert_eq!(connection.id, 10);
        assert_eq!(connection.name, "CubeSales");
        assert_eq!(connection.connection_type, Some(5));
        match &connection.kind {
            WorkbookConnectionKind::Olap {
                connection,
                command,
                command_type,
                local,
                local_connection,
                local_refresh,
                send_locale,
                row_drill_count,
            } => {
                assert_eq!(
                    connection.as_deref(),
                    Some("Provider=MSOLAP;Data Source=olapserver;")
                );
                assert_eq!(command.as_deref(), Some("SalesCube"));
                assert_eq!(*command_type, Some(1));
                assert!(*local);
                assert_eq!(local_connection.as_deref(), Some("CubeFile=cube.cub"));
                assert!(*local_refresh);
                assert!(*send_locale);
                assert_eq!(*row_drill_count, Some(1000));
            }
            other => panic!("unexpected connection kind: {other:?}"),
        }
    }

    fn read_test_connections(connections_bin: Vec<u8>) -> Vec<WorkbookConnection> {
        let mut archive_bytes = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut archive_bytes));
            let opts = zip::write::SimpleFileOptions::default();
            zip.start_file("xl/connections.bin", opts).unwrap();
            zip.write_all(&connections_bin).unwrap();
            zip.finish().unwrap();
        }

        let mut archive = zip::ZipArchive::new(Cursor::new(archive_bytes)).unwrap();
        read_connections(
            &mut archive,
            Some("xl/connections.bin"),
            &FormulaContext::new(vec!["Sheet1".to_string()]),
        )
        .unwrap()
    }

    fn database_connections_bin() -> Vec<u8> {
        let mut out = Vec::new();
        let mut rw = RecordWriter::new(&mut out);
        rw.write_record(records::BRT_BEGIN_EXT_CONNECTIONS, &0u32.to_le_bytes())
            .unwrap();
        rw.write_record(records::BRT_BEGIN_EXT_CONNECTION, &ext_connection_payload())
            .unwrap();
        rw.write_record(records::BRT_BEGIN_EC_DB_PROPS, &db_props_payload())
            .unwrap();
        out
    }

    fn web_connections_bin() -> Vec<u8> {
        let mut out = Vec::new();
        let mut rw = RecordWriter::new(&mut out);
        rw.write_record(records::BRT_BEGIN_EXT_CONNECTIONS, &0u32.to_le_bytes())
            .unwrap();
        rw.write_record(
            records::BRT_BEGIN_EXT_CONNECTION,
            &web_ext_connection_payload(),
        )
        .unwrap();
        rw.write_record(records::BRT_BEGIN_EC_WEB_PROPS, &web_props_payload())
            .unwrap();
        rw.write_record(records::BRT_BEGIN_EC_PARAMS, &2u32.to_le_bytes())
            .unwrap();
        rw.write_record(records::BRT_BEGIN_EC_PARAM, &prompt_param_payload())
            .unwrap();
        rw.write_record(records::BRT_END_EC_PARAM, &[]).unwrap();
        rw.write_record(records::BRT_BEGIN_EC_PARAM, &integer_param_payload())
            .unwrap();
        rw.write_record(records::BRT_END_EC_PARAM, &[]).unwrap();
        rw.write_record(records::BRT_END_EC_PARAMS, &[]).unwrap();
        rw.write_record(records::BRT_END_EC_WEB_PROPS, &[]).unwrap();
        out
    }

    fn text_connections_bin() -> Vec<u8> {
        let mut out = Vec::new();
        let mut rw = RecordWriter::new(&mut out);
        rw.write_record(records::BRT_BEGIN_EXT_CONNECTIONS, &0u32.to_le_bytes())
            .unwrap();
        rw.write_record(
            records::BRT_BEGIN_EXT_CONNECTION,
            &text_ext_connection_payload(),
        )
        .unwrap();
        rw.write_record(records::BRT_BEGIN_EC_TXT_WIZ, &text_props_payload())
            .unwrap();
        rw.write_record(records::BRT_END_EC_TXT_WIZ, &[]).unwrap();
        out
    }

    fn olap_connections_bin() -> Vec<u8> {
        let mut out = Vec::new();
        let mut rw = RecordWriter::new(&mut out);
        rw.write_record(records::BRT_BEGIN_EXT_CONNECTIONS, &0u32.to_le_bytes())
            .unwrap();
        rw.write_record(
            records::BRT_BEGIN_EXT_CONNECTION,
            &olap_ext_connection_payload(),
        )
        .unwrap();
        rw.write_record(records::BRT_BEGIN_EC_DB_PROPS, &olap_db_props_payload())
            .unwrap();
        rw.write_record(records::BRT_BEGIN_EC_OLAP_PROPS, &olap_props_payload())
            .unwrap();
        out
    }

    fn ext_connection_payload() -> Vec<u8> {
        ext_connection_payload_for(5, 7, "SalesConnection", 1, 0x001F)
    }

    fn web_ext_connection_payload() -> Vec<u8> {
        ext_connection_payload_for(4, 8, "WebSales", 2, 0x0008)
    }

    fn text_ext_connection_payload() -> Vec<u8> {
        ext_connection_payload_for(6, 9, "CsvSales", 2, 0x0008)
    }

    fn olap_ext_connection_payload() -> Vec<u8> {
        ext_connection_payload_for(5, 10, "CubeSales", 2, 0x0008)
    }

    fn ext_connection_payload_for(
        source_type: u32,
        id: u32,
        name: &str,
        password_saved: u8,
        load_flags: u16,
    ) -> Vec<u8> {
        let mut payload = vec![3, 0, password_saved, 0];
        payload.extend_from_slice(&15u16.to_le_bytes());
        payload.extend_from_slice(&(0x0001u16 | 0x0010 | 0x0020 | 0x0040).to_le_bytes());
        payload.extend_from_slice(&load_flags.to_le_bytes());
        payload.extend_from_slice(&source_type.to_le_bytes());
        payload.extend_from_slice(&2u32.to_le_bytes());
        payload.extend_from_slice(&id.to_le_bytes());
        payload.push(2);
        if load_flags & 0x0001 != 0 {
            payload.extend_from_slice(&encode_wide_str("sales.csv"));
        }
        if load_flags & 0x0002 != 0 {
            payload.extend_from_slice(&encode_wide_str("sales.odc"));
        }
        if load_flags & 0x0004 != 0 {
            payload.extend_from_slice(&encode_wide_str("Sales Reporting"));
        }
        payload.extend_from_slice(&encode_wide_str(name));
        if load_flags & 0x0010 != 0 {
            payload.extend_from_slice(&encode_wide_str("SalesSso"));
        }
        payload
    }

    fn db_props_payload() -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&2u32.to_le_bytes());
        payload.push(0x02);
        payload.extend_from_slice(&encode_wide_str("Provider=MSDASQL;DSN=Sales;"));
        payload.extend_from_slice(&encode_wide_str("select * from Sales"));
        payload
    }

    fn olap_db_props_payload() -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&1u32.to_le_bytes());
        payload.push(0x02);
        payload.extend_from_slice(&encode_wide_str("Provider=MSOLAP;Data Source=olapserver;"));
        payload.extend_from_slice(&encode_wide_str("SalesCube"));
        payload
    }

    fn olap_props_payload() -> Vec<u8> {
        let mut payload = Vec::new();
        payload.push(0x41);
        payload.extend_from_slice(&1000u32.to_le_bytes());
        payload.push(0x01);
        payload.extend_from_slice(&encode_wide_str("CubeFile=cube.cub"));
        payload
    }

    fn web_props_payload() -> Vec<u8> {
        let mut payload = vec![2];
        payload.extend_from_slice(&(0x0001u16 | 0x0002 | 0x0100).to_le_bytes());
        payload.push(0);
        payload.push(0x07);
        payload.extend_from_slice(&encode_wide_str("http://127.0.0.1/duke-sheets/sales"));
        payload.extend_from_slice(&encode_wide_str("region=west"));
        payload.extend_from_slice(&encode_wide_str("http://127.0.0.1/duke-sheets/edit"));
        payload
    }

    fn text_props_payload() -> Vec<u8> {
        let mut payload = Vec::new();
        let header = 0x0000_0001u32 | (1 << 12) | (1 << 15) | (1 << 20);
        payload.extend_from_slice(&header.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&2u32.to_le_bytes());
        payload.push(b'.');
        payload.push(b',');
        payload.extend_from_slice(&encode_wide_str("/data/sales.csv"));
        payload
    }

    fn prompt_param_payload() -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&12u16.to_le_bytes());
        payload.extend_from_slice(&1u32.to_le_bytes());
        payload.extend_from_slice(&encode_wide_str("Region"));
        payload.extend_from_slice(&encode_wide_str("Choose region"));
        payload
    }

    fn integer_param_payload() -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&4u16.to_le_bytes());
        payload.extend_from_slice(&0x800u32.to_le_bytes());
        payload.extend_from_slice(&encode_wide_str("Limit"));
        payload.extend_from_slice(&25.0f64.to_le_bytes());
        payload
    }
}
