import csv
import datetime as dt
import json
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import pymysql


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = get_base_dir()
CONFIG_PATH = BASE_DIR / "db_backup_config.json"
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"


class DatabaseConnectionError(RuntimeError):
    pass


class Logger:
    def __init__(self, log_path: Path):
        self.log_path = log_path
        self.log_path.parent.mkdir(parents=True, exist_ok=True)

    def _write(self, message: str, stream):
        timestamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"{timestamp} {message}"
        print(line, file=stream)
        with self.log_path.open("a", encoding="utf-8") as f:
            f.write(line + "\n")

    def info(self, message: str):
        self._write(message, sys.stdout)

    def error(self, message: str):
        self._write(message, sys.stderr)


def pause_if_frozen():
    if getattr(sys, "frozen", False):
        try:
            input("Press Enter to exit...")
        except EOFError:
            pass


def load_config(config_path: Path) -> dict:
    with config_path.open("r", encoding="utf-8") as f:
        return json.load(f)


def today_suffix() -> str:
    return dt.datetime.now().strftime("%Y%m%d")


def datetime_suffix() -> str:
    return dt.datetime.now().strftime("%Y%m%d_%H%M%S")


def quote_mysql_identifier(identifier: str) -> str:
    return "`" + identifier.replace("`", "``") + "`"


def safe_name(value: str) -> str:
    return "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in value)


def safe_unlink(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except FileNotFoundError:
        pass


def get_bool_config(project: dict, key: str, default: bool) -> bool:
    value = project.get(key, default)
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"true", "1", "yes", "y"}:
            return True
        if normalized in {"false", "0", "no", "n"}:
            return False
    return bool(value)


def get_int_config(project: dict, key: str, default: int) -> int:
    value = project.get(key, default)
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def validate_project(project: dict) -> None:
    required = ["project_name", "host", "user", "password", "database", "table"]
    missing = [key for key in required if not project.get(key)]
    if missing:
        raise ValueError(f"Missing project fields: {', '.join(missing)}")


def connect_mysql(project: dict):
    try:
        return pymysql.connect(
            host=project["host"],
            port=int(project.get("port", 3306)),
            user=project["user"],
            password=project["password"],
            database=project["database"],
            charset=project.get("charset", "utf8mb4"),
            autocommit=False,
            cursorclass=pymysql.cursors.Cursor,
        )
    except pymysql.MySQLError as exc:
        raise DatabaseConnectionError(str(exc)) from exc


def fetch_columns(connection, table_name: str):
    sql = (
        "SELECT COLUMN_NAME, DATA_TYPE "
        "FROM INFORMATION_SCHEMA.COLUMNS "
        "WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = %s "
        "ORDER BY ORDINAL_POSITION"
    )
    with connection.cursor() as cursor:
        cursor.execute(sql, (table_name,))
        rows = cursor.fetchall()
        return [row[0] for row in rows], {row[0]: row[1] for row in rows}


def column_letter(index: int) -> str:
    letters = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def make_sheet_xml(rows):
    lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    lines.append(
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    )
    lines.append("<sheetData>")
    for row_idx, row in enumerate(rows, start=1):
        lines.append(f'<row r="{row_idx}">')
        for col_idx, value in enumerate(row, start=1):
            cell_ref = f"{column_letter(col_idx)}{row_idx}"
            cell_value = escape("" if value is None else str(value))
            lines.append(
                f'<c r="{cell_ref}" t="inlineStr"><is><t>{cell_value}</t></is></c>'
            )
        lines.append("</row>")
    lines.append("</sheetData>")
    lines.append("</worksheet>")
    return "".join(lines)


def create_xlsx_from_sheet_xml(sheet_xml_path: Path, xlsx_path: Path, sheet_name: str) -> None:
    sheet_xml = sheet_xml_path.read_text(encoding="utf-8")
    with zipfile.ZipFile(xlsx_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>""",
        )
        zf.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>""",
        )
        zf.writestr(
            "docProps/core.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Python</dc:creator>
  <cp:lastModifiedBy>Python</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{dt.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")}</dcterms:created>
</cp:coreProperties>""",
        )
        zf.writestr(
            "docProps/app.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Python</Application>
</Properties>""",
        )
        zf.writestr(
            "xl/workbook.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="{escape(sheet_name[:31])}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>""",
        )
        zf.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""",
        )
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def create_zip(zip_path: Path, file_path: Path) -> None:
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(file_path, arcname=file_path.name)


def resolve_output_base_name(output_dir: Path, prefix: str) -> str:
    base_name = f"{prefix}_{datetime_suffix()}"
    if not any((output_dir / f"{base_name}{ext}").exists() for ext in (".xlsx", ".zip")):
        return base_name

    index = 1
    while True:
        candidate = f"{base_name}_{index}"
        if not any((output_dir / f"{candidate}{ext}").exists() for ext in (".xlsx", ".zip")):
            return candidate
        index += 1


def begin_sheet_xml(sheet_xml_path: Path) -> None:
    with sheet_xml_path.open("w", encoding="utf-8", newline="") as f:
        f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        f.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        f.write("<sheetData>")


def append_sheet_rows(sheet_xml_path: Path, rows, start_row_index: int) -> int:
    row_index = start_row_index
    with sheet_xml_path.open("a", encoding="utf-8", newline="") as f:
        for row in rows:
            f.write(f'<row r="{row_index}">')
            for col_idx, value in enumerate(row, start=1):
                cell_ref = f"{column_letter(col_idx)}{row_index}"
                cell_value = escape("" if value is None else str(value))
                f.write(f'<c r="{cell_ref}" t="inlineStr"><is><t>{cell_value}</t></is></c>')
            f.write("</row>")
            row_index += 1
    return row_index


def finish_sheet_xml(sheet_xml_path: Path) -> None:
    with sheet_xml_path.open("a", encoding="utf-8", newline="") as f:
        f.write("</sheetData>")
        f.write("</worksheet>")


def export_table_to_sheet_xml(connection, project: dict, table_name: str, columns, sheet_xml_path: Path, logger: Logger) -> int:
    batch_size = get_int_config(project, "batch_size", 10000)
    batch_size = max(batch_size, 1)
    batch_key = project.get("batch_key")
    exported_rows = 0
    row_index = 1

    begin_sheet_xml(sheet_xml_path)
    row_index = append_sheet_rows(sheet_xml_path, [columns], row_index)

    with connection.cursor() as cursor:
        if batch_key:
            logger.info(f"[{project['project_name']}] Paging mode: keyset, batch_key={batch_key}, batch_size={batch_size}")
            last_value = None
            while True:
                if last_value is None:
                    sql = (
                        f"SELECT * FROM {quote_mysql_identifier(table_name)} "
                        f"ORDER BY {quote_mysql_identifier(batch_key)} ASC LIMIT %s"
                    )
                    cursor.execute(sql, (batch_size,))
                else:
                    sql = (
                        f"SELECT * FROM {quote_mysql_identifier(table_name)} "
                        f"WHERE {quote_mysql_identifier(batch_key)} > %s "
                        f"ORDER BY {quote_mysql_identifier(batch_key)} ASC LIMIT %s"
                    )
                    cursor.execute(sql, (last_value, batch_size))

                rows = cursor.fetchall()
                if not rows:
                    break

                row_index = append_sheet_rows(sheet_xml_path, rows, row_index)
                exported_rows += len(rows)
                key_index = columns.index(batch_key)
                last_value = rows[-1][key_index]

                logger.info(f"[{project['project_name']}] Exported rows: {exported_rows}")
                if len(rows) < batch_size:
                    break
        else:
            logger.info(f"[{project['project_name']}] Paging mode: offset fallback, batch_size={batch_size}")
            logger.info(f"[{project['project_name']}] WARNING: batch_key is not configured; large tables may export more slowly")
            offset = 0
            while True:
                sql = (
                    f"SELECT * FROM {quote_mysql_identifier(table_name)} "
                    f"LIMIT %s OFFSET %s"
                )
                cursor.execute(sql, (batch_size, offset))
                rows = cursor.fetchall()
                if not rows:
                    break

                row_index = append_sheet_rows(sheet_xml_path, rows, row_index)
                exported_rows += len(rows)
                offset += len(rows)

                logger.info(f"[{project['project_name']}] Exported rows: {exported_rows}")
                if len(rows) < batch_size:
                    break

    finish_sheet_xml(sheet_xml_path)
    return exported_rows


def process_project(project: dict, logger: Logger) -> None:
    validate_project(project)

    project_name = project["project_name"]
    table_name = project["table"]
    keep_excel = get_bool_config(project, "keep_excel", True)

    connection = connect_mysql(project)
    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_base_name = resolve_output_base_name(OUTPUT_DIR, f"{safe_name(project_name)}_{safe_name(table_name)}")
        xlsx_path = OUTPUT_DIR / f"{output_base_name}.xlsx"
        zip_path = OUTPUT_DIR / f"{output_base_name}.zip"
        columns, _ = fetch_columns(connection, table_name)

        temp_dir = Path(tempfile.mkdtemp(prefix="coupang_export_", dir=str(OUTPUT_DIR)))
        sheet_xml_path = temp_dir / "sheet1.xml"
        try:
            exported_rows = export_table_to_sheet_xml(connection, project, table_name, columns, sheet_xml_path, logger)
            create_xlsx_from_sheet_xml(sheet_xml_path, xlsx_path, sheet_name=table_name)
        finally:
            safe_unlink(sheet_xml_path)
            try:
                if temp_dir.exists():
                    temp_dir.rmdir()
            except OSError:
                pass

        create_zip(zip_path, xlsx_path)
    finally:
        connection.close()

    if not keep_excel:
        safe_unlink(xlsx_path)
        logger.info(f"[{project_name}] Excel removed because keep_excel is false")

    logger.info(f"[{project_name}] SUCCESS")
    logger.info(f"[{project_name}] keep_excel: {keep_excel}")
    logger.info(f"[{project_name}] Source table: {table_name}")
    logger.info(f"[{project_name}] Output base name: {output_base_name}")
    logger.info(f"[{project_name}] Exported rows: {exported_rows}")
    logger.info(f"[{project_name}] Excel: {xlsx_path}")
    logger.info(f"[{project_name}] ZIP: {zip_path}")
    logger.info(f"[{project_name}] Database connection closed")


def main():
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(f"Config file not found: {CONFIG_PATH}")

    config = load_config(CONFIG_PATH)
    projects = config.get("projects")
    if not isinstance(projects, list) or not projects:
        raise ValueError("Config must contain a non-empty 'projects' list")

    log_path = LOG_DIR / f"db_backup_{today_suffix()}.log"
    logger = Logger(log_path)
    logger.info(f"Base directory: {BASE_DIR}")
    logger.info(f"Config path: {CONFIG_PATH}")
    logger.info(f"Output directory: {OUTPUT_DIR}")
    logger.info(f"Log directory: {LOG_DIR}")
    logger.info(f"Log file: {log_path}")
    logger.info(f"Project count: {len(projects)}")

    failed_projects = []
    for project in projects:
        project_name = project.get("project_name", "UNKNOWN")
        try:
            process_project(project, logger)
        except DatabaseConnectionError as exc:
            failed_projects.append(project_name)
            logger.error(f"[{project_name}] DATABASE CONNECTION FAILED")
            logger.error(f"[{project_name}] Detail: {exc}")
            logger.error(f"[{project_name}] This project has been skipped and requires manual handling")
        except Exception as exc:
            failed_projects.append(project_name)
            logger.error(f"[{project_name}] FAILED: {exc}")

    logger.info("-" * 60)
    logger.info(f"Total projects: {len(projects)}")
    logger.info(f"Success count: {len(projects) - len(failed_projects)}")
    logger.info(f"Failed count: {len(failed_projects)}")
    if failed_projects:
        logger.error("Failed projects: " + ", ".join(failed_projects))
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        fallback_log_dir = BASE_DIR / "logs"
        fallback_log_dir.mkdir(parents=True, exist_ok=True)
        fallback_log_path = fallback_log_dir / f"db_backup_fatal_{today_suffix()}.log"
        timestamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        message = f"{timestamp} FATAL: {exc}"
        print(message, file=sys.stderr)
        with fallback_log_path.open("a", encoding="utf-8") as f:
            f.write(message + "\n")
        pause_if_frozen()
        raise
