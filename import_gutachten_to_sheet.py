import re
import sys
from os import path, getenv
from datetime import datetime
from zoneinfo import ZoneInfo
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Konfiguration
SERVICE_ACCOUNT_FILE = getenv('GOOGLE_APPLICATION_CREDENTIALS', 'ux-dashboard-465511-29cd7fce4011.json')
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

SPREADSHEET_ID = getenv('SHEET_ID', '10mfm9SVVDiWcxnfK2QuUCj3msaVFBQIQx34NnPlUEo4')
TAB_NAME = getenv('SHEET_TAB_NAME', 'Dashboard')
STATISTIK_TAB = getenv('SHEET_STATISTIK_TAB', 'Statistik')
LEGACY_IMPORT_LOG_TAB = getenv('SHEET_IMPORT_LOG_TAB', 'ImportLog')
LEGACY_TAGES_STAT_TAB = getenv('SHEET_TAGES_STAT_TAB', 'TagesStat')
WOCHEN_STAT_TAB = getenv('SHEET_WOCHEN_STAT_TAB', 'WochenStat')
WOCHEN_STAT_ARCHIV_TAB = getenv('SHEET_WOCHEN_STAT_ARCHIV_TAB', 'WochenStat_Archiv')
WOCHEN_STAT_HEADERS = ['Jahr', 'KW', 'Wochenstart', 'Nummer', 'Anzahl', 'Aktualisiert']
WOCHEN_STAT_RANGE = 'A2:F'
WOCHEN_STAT_ARCHIV_FROZEN = getenv('WOCHEN_STAT_ARCHIV_FROZEN', '1').strip().lower() not in ('0', 'false', 'no')
WOCHEN_STAT_MANUAL = getenv('WOCHEN_STAT_MANUAL', '1').strip().lower() not in ('0', 'false', 'no')
WOCHEN_STAT_AUSWERTUNG_TAB = getenv('SHEET_WOCHEN_STAT_AUSWERTUNG_TAB', 'WochenStat_Auswertung')
FOLDER_ID = getenv('DRIVE_FOLDER_ID', '1FVnM3Y_ktIvXMUPuAQTpJ-sMB5yI1gYf')
RB_FOLDER_ID = getenv('DRIVE_RB_FOLDER_ID', '1Lpzu-pK94B2asLbbUgzAaOxZ3oj_DrnR')
INCLUDE_ALL_DRIVES = getenv('DRIVE_INCLUDE_ALL_DRIVES', '1').strip().lower() not in ('0', 'false', 'no')
LOCAL_TIMEZONE = getenv('LOCAL_TIMEZONE', 'Europe/Berlin')


def local_now():
    """Lokale Zeit (Standard: Europe/Berlin) — unabhängig von Container-UTC."""
    return datetime.now(ZoneInfo(LOCAL_TIMEZONE))


_now = local_now()
current_year = str(_now.year % 100).zfill(2)
previous_year = str((_now.year - 1) % 100).zfill(2)
AKTE_REGEX = re.compile(r'(\d{3,5})\s*[/_:\-]\s*(\d{2})')
COMPACT_AKTE_REGEX = re.compile(r'\b(\d{8})\b')


def parse_allowed_years(raw_value):
    years = set()
    for token in str(raw_value or '').split(','):
        clean = re.sub(r'[^0-9]', '', token)
        if len(clean) == 2:
            years.add(clean)
        elif len(clean) == 4:
            years.add(clean[-2:])
    return years


ALLOWED_YEARS = parse_allowed_years(getenv('AKTE_ALLOWED_YEARS', f'{previous_year},{current_year}')) or {current_year}


def parse_csv_tokens(raw_value, default_value):
    value = raw_value if raw_value is not None else default_value
    return [token.strip().lower() for token in str(value).split(',') if token.strip()]


ACCEPTED_FOLDER_TERMS = parse_csv_tokens(
    getenv('AKTE_ACCEPTED_FOLDER_TERMS'),
    'gutachten,beratungsleistungen,kva,kostenvoranschlag'
)
IGNORED_PREFIXES = parse_csv_tokens(
    getenv('AKTE_IGNORED_PREFIXES'),
    'rb'
)


def normalize_number(value):
    return re.sub(r'[^0-9]', '', str(value or ''))


def has_accepted_folder_term(name):
    lower_name = name.lower()
    return any(term in lower_name for term in ACCEPTED_FOLDER_TERMS)


def has_ignored_prefix(name):
    return any(re.match(rf'^{re.escape(prefix)}[\s_:\-]?', name, re.IGNORECASE) for prefix in IGNORED_PREFIXES)


def extract_number_and_year(name):
    match = AKTE_REGEX.search(name)
    if match:
        return normalize_number(match.group(1)), match.group(2)

    compact_match = COMPACT_AKTE_REGEX.search(name)
    if compact_match:
        compact_value = compact_match.group(1)
        return normalize_number(compact_value), compact_value[-2:]

    return '', ''


def list_drive_files(drive_service, folder_id=None, folders_only=False):
    folder_id = folder_id or FOLDER_ID
    query_parts = [f"'{folder_id}' in parents", 'trashed = false']
    if folders_only:
        query_parts.append("mimeType = 'application/vnd.google-apps.folder'")

    files = []
    page_token = None

    while True:
        request = drive_service.files().list(
            q=' and '.join(query_parts),
            fields='nextPageToken, files(name, mimeType)',
            pageSize=200,
            pageToken=page_token,
            includeItemsFromAllDrives=INCLUDE_ALL_DRIVES,
            supportsAllDrives=INCLUDE_ALL_DRIVES
        )
        response = request.execute()

        files.extend(response.get('files', []))
        page_token = response.get('nextPageToken')
        if not page_token:
            break

    return files


def count_rb_folders(drive_service):
    if not RB_FOLDER_ID:
        return 0

    try:
        items = list_drive_files(drive_service, RB_FOLDER_ID, folders_only=True)
    except HttpError as exc:
        print(f'⚠️ RB-Ordner konnte nicht gelesen werden: {exc}', file=sys.stderr)
        return 0

    count = 0
    for item in items:
        name = item.get('name', '').strip()
        if not name or name.lower() == 'organisation':
            continue
        count += 1
    return count


def import_numbers_for_date(log_rows, date_str):
    seen = set()
    for row in log_rows:
        if len(row) < 3 or str(row[0]).strip() != date_str:
            continue
        nummer = normalize_number(row[2])
        if nummer:
            seen.add(nummer)
    return seen


SHORTCODE_TO_BEARBEITER = {
    'RO': 'Robar Kassem',
    'HB': 'Hussein Selman',
    'MZ': 'Mohamed Zahreddine',
    'HJ': 'Hussein Jaber',
}


def extract_folder_shortcode(name):
    match = re.search(r'\(([A-Za-z]{1,4})\)\s*$', str(name or '').strip())
    return match.group(1).upper() if match else ''


def bearbeiter_from_folder_name(name):
    shortcode = extract_folder_shortcode(name)
    return SHORTCODE_TO_BEARBEITER.get(shortcode, ''), shortcode


def find_new_entries(dateien, filtered_rows, skip_numbers=None):
    vorhandene = {
        normalize_number(row[0])
        for row in filtered_rows
        if row and normalize_number(row[0])
    }
    if skip_numbers:
        vorhandene |= {
            normalize_number(nummer)
            for nummer in skip_numbers
            if normalize_number(nummer)
        }
    neue_eintraege = []
    gesehen = set(vorhandene)

    for file in dateien:
        name = file['name'].strip()

        if name.lower() == 'organisation':
            print(f'⏭️ Übersprungen (Organisation): {name}')
            continue
        if has_ignored_prefix(name):
            print(f'⏭️ Übersprungen (Prefix ignoriert): {name}')
            continue
        if not has_accepted_folder_term(name):
            print(f'⏭️ Übersprungen (kein erlaubter Ordnertyp): {name}')
            continue

        nummer, jahr = extract_number_and_year(name)
        if not nummer or not jahr:
            print(f'⏭️ Kein gültiges Format: {name}')
            continue

        if jahr not in ALLOWED_YEARS:
            print(f'⏭️ Übersprungen (Jahr {jahr} nicht erlaubt): {name}')
            continue

        if not nummer or nummer in gesehen:
            continue

        gesehen.add(nummer)
        bearbeiter, shortcode = bearbeiter_from_folder_name(name)
        neue_eintraege.append({
            'nummer': nummer,
            'folder_name': name,
            'bearbeiter': bearbeiter,
            'shortcode': shortcode,
        })

    return neue_eintraege


def is_valid_drive_entry(name):
    name = name.strip()
    if name.lower() == 'organisation':
        return False
    if has_ignored_prefix(name):
        return False
    if not has_accepted_folder_term(name):
        return False

    nummer, jahr = extract_number_and_year(name)
    if not nummer or not jahr:
        return False
    if jahr not in ALLOWED_YEARS:
        return False
    return True


def ensure_tab(sheets_service, title, headers):
    meta = sheets_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    existing = {sheet['properties']['title'] for sheet in meta.get('sheets', [])}
    if title in existing:
        return

    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={'requests': [{'addSheet': {'properties': {'title': title}}}]}
    ).execute()
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f'{title}!A1',
        valueInputOption='RAW',
        body={'values': [headers]}
    ).execute()
    print(f'ℹ️ Tab "{title}" angelegt.')


def tab_exists(sheets_service, title):
    meta = sheets_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    existing = {sheet['properties']['title'] for sheet in meta.get('sheets', [])}
    return title in existing


def read_sheet_values(sheets_service, tab, cell_range):
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f'{tab}!{cell_range}'
        ).execute()
    except HttpError:
        return []
    return result.get('values', []) or []


def ensure_statistik_tab(sheets_service):
    ensure_tab(sheets_service, STATISTIK_TAB, ['Datum', 'Uhrzeit', 'Aktennummer'])
    sheets_service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={'valueInputOption': 'RAW', 'data': [
            {'range': f'{STATISTIK_TAB}!A1:C1', 'values': [['Datum', 'Uhrzeit', 'Aktennummer']]},
            {'range': f'{STATISTIK_TAB}!E1', 'values': [['Letzter_Lauf']]},
            {'range': f'{STATISTIK_TAB}!H1:J1', 'values': [['Datum', 'Sync', 'RB_Offene']]},
        ]}
    ).execute()


def migrate_statistik_data(sheets_service):
    import_rows = read_sheet_values(sheets_service, STATISTIK_TAB, 'A2:C')
    daily_rows = read_sheet_values(sheets_service, STATISTIK_TAB, 'H2:J')
    run_time = read_sheet_values(sheets_service, STATISTIK_TAB, 'F1')

    updates = []

    if not import_rows and tab_exists(sheets_service, LEGACY_IMPORT_LOG_TAB):
        legacy_import = read_sheet_values(sheets_service, LEGACY_IMPORT_LOG_TAB, 'A2:C')
        if legacy_import:
            updates.append({'range': f'{STATISTIK_TAB}!A2:C', 'values': legacy_import})
            print(f'ℹ️ {len(legacy_import)} ImportLog-Zeilen nach Statistik übernommen.')

        legacy_run = read_sheet_values(sheets_service, LEGACY_IMPORT_LOG_TAB, 'D1:E1')
        if legacy_run and legacy_run[0] and len(legacy_run[0]) >= 2 and not run_time:
            updates.append({'range': f'{STATISTIK_TAB}!F1', 'values': [[legacy_run[0][1]]]})

    if not daily_rows and tab_exists(sheets_service, LEGACY_TAGES_STAT_TAB):
        legacy_daily = read_sheet_values(sheets_service, LEGACY_TAGES_STAT_TAB, 'A2:F')
        migrated = [normalize_tages_stat_row(row) for row in legacy_daily]
        migrated = [row for row in migrated if row and re.match(r'^\d{4}-\d{2}-\d{2}$', str(row[0]).strip())]
        if migrated:
            updates.append({'range': f'{STATISTIK_TAB}!H2:J', 'values': migrated})
            print(f'ℹ️ {len(migrated)} TagesStat-Zeilen nach Statistik übernommen.')

    if updates:
        sheets_service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={'valueInputOption': 'RAW', 'data': updates}
        ).execute()


def append_import_log(sheets_service, nummern):
    if not nummern:
        return

    ensure_statistik_tab(sheets_service)
    migrate_statistik_data(sheets_service)
    now = local_now()
    today = now.strftime('%Y-%m-%d')
    logged_today = import_numbers_for_date(
        read_sheet_values(sheets_service, STATISTIK_TAB, 'A2:C'),
        today,
    )
    to_log = [
        nummer for nummer in nummern
        if normalize_number(nummer) not in logged_today
    ]
    if not to_log:
        return

    rows = [[today, now.strftime('%H:%M:%S'), nummer] for nummer in to_log]
    # Kein INSERT_ROWS: würde ganze Tabellenzeilen einfügen und H:J mit nach unten schieben.
    sheets_service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f'{STATISTIK_TAB}!A:C',
        valueInputOption='RAW',
        body={'values': rows}
    ).execute()


def record_import_run(sheets_service):
    ensure_statistik_tab(sheets_service)
    migrate_statistik_data(sheets_service)
    now = local_now()
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f'{STATISTIK_TAB}!F1',
        valueInputOption='RAW',
        body={'values': [[now.strftime('%Y-%m-%d %H:%M:%S')]]}
    ).execute()


def read_import_log_rows(sheets_service):
    ensure_statistik_tab(sheets_service)
    migrate_statistik_data(sheets_service)
    return read_sheet_values(sheets_service, STATISTIK_TAB, 'A2:C')


def count_imports_for_date(log_rows, date_str):
    return len(import_numbers_for_date(log_rows, date_str))


def normalize_tages_stat_row(row):
    if not row:
        return row
    # Altes Format: Datum, Neu, Offen_Tagesstart, Drive_Ordner, Sync, RB_Offene
    if len(row) >= 6:
        return [row[0], row[4], row[5]]
    # Vorheriges Format: Datum, Neu, Offen_Tagesstart, Sync, RB_Offene
    if len(row) >= 5:
        return [row[0], row[3], row[4]]
    if len(row) == 3:
        return row
    if len(row) == 4:
        return [row[0], row[2], row[3] if len(row) > 3 else '']
    return row


def update_tages_stat(sheets_service, sync_ok, rb_count=0):
    ensure_statistik_tab(sheets_service)
    migrate_statistik_data(sheets_service)

    today = local_now().strftime('%Y-%m-%d')
    sync_label = 'OK' if sync_ok else 'Offen'

    rows = [
        normalize_tages_stat_row(row)
        for row in read_sheet_values(sheets_service, STATISTIK_TAB, 'H2:J')
    ]
    rows = [row for row in rows if row and re.match(r'^\d{4}-\d{2}-\d{2}$', str(row[0]).strip())]

    today_idx = next((i for i, row in enumerate(rows) if row and row[0] == today), None)
    if today_idx is None:
        rows.append([today, sync_label, str(rb_count)])
    else:
        row = list(rows[today_idx])
        while len(row) < 3:
            row.append('')
        row[1] = sync_label
        row[2] = str(rb_count)
        rows[today_idx] = row

    sheets_service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f'{STATISTIK_TAB}!H2:J'
    ).execute()
    if rows:
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f'{STATISTIK_TAB}!H2',
            valueInputOption='RAW',
            body={'values': rows}
        ).execute()


def iso_year_week(dt=None):
    dt = dt or local_now()
    iso = dt.isocalendar()
    return iso[0], iso[1]


def get_max_akten_nummer(rows, extra_nummern=None):
    max_num = 0
    for row in rows:
        nummer = normalize_number(row[0] if row else '')
        if nummer.isdigit():
            max_num = max(max_num, int(nummer))
    for nummer in extra_nummern or []:
        clean = normalize_number(nummer)
        if clean.isdigit():
            max_num = max(max_num, int(clean))
    return max_num


def current_calendar_year():
    return str(local_now().year)


def wochen_row_year(row):
    if not row:
        return ''
    return str(row[0]).strip()


def split_wochen_rows(rows):
    year = current_calendar_year()
    current = []
    archive = []
    for raw in rows:
        row = normalize_wochen_stat_row(raw)
        if not row or len(row) < 2:
            continue
        if wochen_row_year(row) == year:
            current.append(row)
        else:
            archive.append(row)
    current.sort(key=lambda r: int(str(r[1]).strip() or 0))
    archive.sort(key=lambda r: (int(wochen_row_year(r) or 0), int(str(r[1]).strip() or 0)))
    return current, archive


def is_wochen_section_row(row):
    if not row:
        return False
    year = wochen_row_year(row)
    if not re.fullmatch(r'\d{4}', year):
        return False
    if len(row) < 2:
        return True
    kw = str(row[1]).strip()
    if kw:
        return False
    return not any(str(row[i]).strip() for i in range(2, min(6, len(row))))


def is_wochen_summary_row(row):
    if not row:
        return False
    for idx in (1, 3):
        if idx < len(row) and str(row[idx]).strip().lower() in ('jahressumme', 'summe', 'gesamt'):
            return True
    return False


def year_anzahl_sum(rows):
    return sum(int(normalize_number(row[4]) or 0) for row in rows)


def append_year_summary_row(rows):
    if not rows:
        return rows, []
    total = year_anzahl_sum(rows)
    return rows + [['', '', '', 'Jahressumme', str(total), '']], [len(rows) + 2]


def read_wochenstat_tab(sheets_service, tab):
    data_rows, _ = read_wochenstat_tab_with_trailing(sheets_service, tab)
    return data_rows


def read_wochenstat_tab_with_trailing(sheets_service, tab):
    if not tab_exists(sheets_service, tab):
        return [], []
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f'{tab}!{WOCHEN_STAT_RANGE}',
    ).execute()
    data_rows = []
    trailing_rows = []
    for row in (result.get('values', []) or []):
        normalized = normalize_wochen_stat_row(row)
        if not normalized:
            continue
        if is_wochen_section_row(normalized) or is_wochen_summary_row(normalized):
            trailing_rows.append(normalized)
            continue
        if len(normalized) < 2 or not str(normalized[1]).strip():
            trailing_rows.append(normalized)
            continue
        data_rows.append(normalized)
    return data_rows, trailing_rows


def read_all_wochenstat(sheets_service):
    rows = read_wochenstat_tab(sheets_service, WOCHEN_STAT_TAB)
    rows.extend(read_wochenstat_tab(sheets_service, WOCHEN_STAT_ARCHIV_TAB))
    return rows


def get_sheet_id(sheets_service, title):
    meta = sheets_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for sheet in meta.get('sheets', []):
        if sheet.get('properties', {}).get('title') == title:
            return sheet['properties']['sheetId']
    return None


def decorate_archive_rows(rows):
    output = []
    section_sheet_rows = []
    summary_sheet_rows = []
    by_year = {}
    for row in rows:
        by_year.setdefault(wochen_row_year(row), []).append(row)

    for year in sorted(by_year, key=lambda value: int(value or 0)):
        year_rows = by_year[year]
        section_sheet_rows.append(len(output) + 2)
        output.append([year, '', '', '', '', ''])
        output.extend(year_rows)
        total = year_anzahl_sum(year_rows)
        summary_sheet_rows.append(len(output) + 2)
        output.append(['', '', '', 'Jahressumme', str(total), ''])

    return output, section_sheet_rows, summary_sheet_rows


def _color(red, green, blue):
    return {'red': red, 'green': green, 'blue': blue}


def format_wochenstat_archiv(sheets_service, total_rows, section_sheet_rows, summary_sheet_rows=None):
    sheet_id = get_sheet_id(sheets_service, WOCHEN_STAT_ARCHIV_TAB)
    if sheet_id is None or total_rows < 2:
        return

    requests = [
        {
            'updateSheetProperties': {
                'properties': {
                    'sheetId': sheet_id,
                    'gridProperties': {'frozenRowCount': 1},
                },
                'fields': 'gridProperties.frozenRowCount',
            }
        },
        {
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': 6,
                },
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': _color(0.12, 0.16, 0.22),
                        'horizontalAlignment': 'CENTER',
                        'textFormat': {
                            'bold': True,
                            'fontSize': 10,
                            'foregroundColor': _color(0.95, 0.97, 1.0),
                        },
                    }
                },
                'fields': 'userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)',
            }
        },
        {
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 1,
                    'endRowIndex': total_rows,
                    'startColumnIndex': 0,
                    'endColumnIndex': 6,
                },
                'cell': {
                    'userEnteredFormat': {
                        'verticalAlignment': 'MIDDLE',
                        'wrapStrategy': 'CLIP',
                    }
                },
                'fields': 'userEnteredFormat(verticalAlignment,wrapStrategy)',
            }
        },
        {
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 1,
                    'endRowIndex': total_rows,
                    'startColumnIndex': 1,
                    'endColumnIndex': 2,
                },
                'cell': {
                    'userEnteredFormat': {'horizontalAlignment': 'CENTER'}
                },
                'fields': 'userEnteredFormat.horizontalAlignment',
            }
        },
        {
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 1,
                    'endRowIndex': total_rows,
                    'startColumnIndex': 2,
                    'endColumnIndex': 5,
                },
                'cell': {
                    'userEnteredFormat': {'horizontalAlignment': 'RIGHT'}
                },
                'fields': 'userEnteredFormat.horizontalAlignment',
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 0,
                    'endIndex': 1,
                },
                'properties': {'pixelSize': 64},
                'fields': 'pixelSize',
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 1,
                    'endIndex': 2,
                },
                'properties': {'pixelSize': 52},
                'fields': 'pixelSize',
            }
        },
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 2,
                    'endIndex': 6,
                },
                'properties': {'pixelSize': 108},
                'fields': 'pixelSize',
            }
        },
        {
            'setBasicFilter': {
                'filter': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': 0,
                        'endRowIndex': total_rows,
                        'startColumnIndex': 0,
                        'endColumnIndex': 6,
                    }
                }
            }
        },
        {
            'unmergeCells': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 1,
                    'endRowIndex': max(total_rows, 2),
                    'startColumnIndex': 0,
                    'endColumnIndex': 6,
                }
            }
        },
    ]

    for sheet_row in section_sheet_rows:
        row_index = sheet_row - 1
        requests.append({
            'mergeCells': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': row_index,
                    'endRowIndex': row_index + 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': 6,
                },
                'mergeType': 'MERGE_ALL',
            }
        })
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': row_index,
                    'endRowIndex': row_index + 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': 6,
                },
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': _color(0.86, 0.91, 0.98),
                        'horizontalAlignment': 'LEFT',
                        'padding': {'left': 12},
                        'textFormat': {
                            'bold': True,
                            'fontSize': 12,
                            'foregroundColor': _color(0.12, 0.23, 0.45),
                        },
                    }
                },
                'fields': 'userEnteredFormat(backgroundColor,horizontalAlignment,padding,textFormat)',
            }
        })

    for sheet_row in summary_sheet_rows or []:
        row_index = sheet_row - 1
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': row_index,
                    'endRowIndex': row_index + 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': 6,
                },
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': _color(0.91, 0.96, 0.91),
                        'textFormat': {
                            'bold': True,
                            'fontSize': 11,
                            'foregroundColor': _color(0.1, 0.35, 0.16),
                        },
                    }
                },
                'fields': 'userEnteredFormat(backgroundColor,textFormat)',
            }
        })
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': row_index,
                    'endRowIndex': row_index + 1,
                    'startColumnIndex': 3,
                    'endColumnIndex': 4,
                },
                'cell': {
                    'userEnteredFormat': {'horizontalAlignment': 'RIGHT'}
                },
                'fields': 'userEnteredFormat.horizontalAlignment',
            }
        })

    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={'requests': requests},
    ).execute()


def write_wochenstat_tab(
    sheets_service,
    tab,
    rows,
    *,
    format_archiv=False,
    section_sheet_rows=None,
    summary_sheet_rows=None,
):
    ensure_tab(sheets_service, tab, WOCHEN_STAT_HEADERS)
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f'{tab}!A1',
        valueInputOption='RAW',
        body={'values': [WOCHEN_STAT_HEADERS]},
    ).execute()
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f'{tab}!{WOCHEN_STAT_RANGE}',
    ).execute()
    if rows:
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f'{tab}!A2',
            valueInputOption='RAW',
            body={'values': rows},
        ).execute()
    if format_archiv:
        format_wochenstat_archiv(
            sheets_service,
            1 + len(rows),
            section_sheet_rows or [],
            summary_sheet_rows or [],
        )


def write_wochenstat_preserve(sheets_service, data_rows, trailing_rows=None):
    output = list(data_rows)
    if trailing_rows:
        output.extend(trailing_rows)
    write_wochenstat_tab(sheets_service, WOCHEN_STAT_TAB, output)


def merge_wochenstat_current_rows(sheets_service, incoming_rows):
    """Bestehende KW-Zeilen behalten, nur fehlende ergänzen."""
    year = current_calendar_year()
    existing_data, trailing = read_wochenstat_tab_with_trailing(sheets_service, WOCHEN_STAT_TAB)
    merged = {
        (wochen_row_year(row), str(row[1]).strip()): row
        for row in existing_data
        if wochen_row_year(row) == year and str(row[1]).strip()
    }

    for raw in incoming_rows:
        row = normalize_wochen_stat_row(raw)
        if not row or wochen_row_year(row) != year or not str(row[1]).strip():
            continue
        key = (wochen_row_year(row), str(row[1]).strip())
        if WOCHEN_STAT_MANUAL and key in merged:
            continue
        if key not in merged or not WOCHEN_STAT_MANUAL:
            merged[key] = row

    data_rows = sorted(
        merged.values(),
        key=lambda row: int(str(row[1]).strip() or 0),
    )
    return data_rows, trailing


def write_wochenstat_current(sheets_service, rows):
    """Nur WochenStat (aktuelles Jahr) – bestehende Zeilen bleiben unverändert."""
    year = current_calendar_year()
    incoming = [
        normalize_wochen_stat_row(row)
        for row in rows
        if wochen_row_year(normalize_wochen_stat_row(row)) == year
    ]
    if WOCHEN_STAT_MANUAL:
        data_rows, trailing = merge_wochenstat_current_rows(sheets_service, incoming)
        write_wochenstat_preserve(sheets_service, data_rows, trailing)
        return

    data_rows, _ = append_year_summary_row(incoming)
    write_wochenstat_tab(sheets_service, WOCHEN_STAT_TAB, data_rows)


def write_wochenstat_tabs(sheets_service, rows, *, rewrite_archiv=False):
    current, archive = split_wochen_rows(rows)
    write_wochenstat_current(sheets_service, current)

    if WOCHEN_STAT_ARCHIV_FROZEN and not rewrite_archiv:
        return

    decorated, section_rows, summary_rows = decorate_archive_rows(archive)
    write_wochenstat_tab(
        sheets_service,
        WOCHEN_STAT_ARCHIV_TAB,
        decorated,
        format_archiv=True,
        section_sheet_rows=section_rows,
        summary_sheet_rows=summary_rows,
    )


def normalize_wochen_stat_row(row):
    if not row or len(row) < 3:
        return row
    # Sehr alt: Jahr, KW, Nummer, Aktualisiert
    if len(row) >= 4 and re.match(r'^\d{4}-\d{2}-\d{2}$', str(row[3]).strip()):
        nummer = str(row[2]).strip()
        return [row[0], row[1], nummer, nummer, '1', row[3]]
    # Ohne Anzahl: Jahr, KW, Wochenstart, Nummer, Aktualisiert
    if len(row) == 5 and re.match(r'^\d{4}-\d{2}-\d{2}$', str(row[4]).strip()):
        start = int(normalize_number(row[2]) or 0)
        end = int(normalize_number(row[3]) or 0)
        anzahl = str(max(1, end - start + 1)) if start and end >= start else '1'
        return [row[0], row[1], row[2], row[3], anzahl, row[4]]
    normalized = list(row)
    while len(normalized) < 6:
        normalized.append('')
    return normalized[:6]


def update_wochen_stat(sheets_service, max_nummer):
    if max_nummer <= 0:
        return

    year, kw = iso_year_week()
    today = local_now().strftime('%Y-%m-%d')
    calendar_year = current_calendar_year()

    data_rows, trailing = read_wochenstat_tab_with_trailing(sheets_service, WOCHEN_STAT_TAB)
    pruned = [row for row in data_rows if wochen_row_year(row) == calendar_year]

    existing_kw = next(
        (
            row for row in pruned
            if str(row[0]).strip() == str(year) and str(row[1]).strip() == str(kw)
        ),
        None,
    )

    if existing_kw is not None:
        # Aktuelle KW immer aus Drive/Dashboard aktualisieren; WOCHEN_STAT_MANUAL
        # schützt nur ältere KW-Zeilen (merge_wochenstat_current_rows).

        row = list(existing_kw)
        while len(row) < 6:
            row.append('')
        if not normalize_number(row[2]):
            row[2] = str(max_nummer)
        start = int(normalize_number(row[2]) or max_nummer)
        existing = int(normalize_number(row[3]) or 0)
        end = max(existing, max_nummer)
        row[3] = str(end)
        row[4] = str(max(1, end - start + 1))
        row[5] = today
        pruned = [
            row if str(r[0]).strip() == str(year) and str(r[1]).strip() == str(kw) else r
            for r in pruned
        ]
        pruned.sort(key=lambda row: int(str(row[1]).strip() or 0))
        write_wochenstat_preserve(sheets_service, pruned, trailing)
        return

    pruned.append([str(year), str(kw), str(max_nummer), str(max_nummer), '1', today])
    pruned.sort(key=lambda row: int(str(row[1]).strip() or 0))
    write_wochenstat_preserve(sheets_service, pruned, trailing)


def main():
    if not path.exists(SERVICE_ACCOUNT_FILE):
        print(f'❌ Service-Account-Datei nicht gefunden: {SERVICE_ACCOUNT_FILE}', file=sys.stderr)
        return 1

    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    drive_service = build('drive', 'v3', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)

    # 1. Alle Daten aus dem Sheet lesen (A-C, ab Zeile 2)
    sheet = sheets_service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=f'{TAB_NAME}!A2:C').execute()
    rows = result.get('values', [])

    if not rows:
        rows = []

    # 2. Alle Zeilen mit Status 'versendet' aussortieren
    filtered_rows = []
    for row in rows:
        status = row[2].strip().lower() if len(row) > 2 else ''
        if not status.startswith('versendet'):
            filtered_rows.append(row)

    # 3. Alte Daten löschen
    sheet.values().clear(spreadsheetId=SPREADSHEET_ID, range=f'{TAB_NAME}!A2:C').execute()

    # 4. Gefilterte Daten zurückschreiben
    if filtered_rows:
        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f'{TAB_NAME}!A2',
            valueInputOption='RAW',
            body={'values': filtered_rows}
        ).execute()

    print(f"✅ {len(rows) - len(filtered_rows)} versendete Zeilen gelöscht.")

    # 5. Neue Einträge aus Google Drive abrufen
    print(f"ℹ️ Drive-Folder-ID: {FOLDER_ID}")
    print(f"ℹ️ Erlaubte Jahre: {', '.join(sorted(ALLOWED_YEARS))}")
    print(f"ℹ️ Erlaubte Ordnerarten: {', '.join(ACCEPTED_FOLDER_TERMS)}")
    try:
        dateien = list_drive_files(drive_service)
    except HttpError as exc:
        print(f'❌ Drive-Abfrage fehlgeschlagen: {exc}', file=sys.stderr)
        print('❌ Prüfe, ob der Service-Account Zugriff auf den Ordner hat.', file=sys.stderr)
        return 1
    print(f"ℹ️ Dateien im Drive-Ordner gefunden: {len(dateien)}")
    log_rows = read_import_log_rows(sheets_service)
    today = local_now().strftime('%Y-%m-%d')
    logged_today = import_numbers_for_date(log_rows, today)
    neue_eintraege = find_new_entries(dateien, filtered_rows, skip_numbers=logged_today)
    neue_nummern = [entry['nummer'] for entry in neue_eintraege]

    # 6. Neue Einträge gezielt in Spalte A (+ Bearbeiter aus Drive-Kürzel in B) schreiben
    startzeile = len(filtered_rows) + 2
    if neue_eintraege:
        values = [
            [entry['nummer'], entry['bearbeiter']] if entry['bearbeiter'] else [entry['nummer']]
            for entry in neue_eintraege
        ]
        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f'{TAB_NAME}!A{startzeile}',
            valueInputOption='RAW',
            body={'values': values}
        ).execute()
        append_import_log(sheets_service, neue_nummern)
        for entry in neue_eintraege:
            if entry['bearbeiter']:
                print(
                    f"👤 {entry['nummer']} ← {entry['bearbeiter']} "
                    f"(Kürzel {entry['shortcode']})"
                )

    if neue_nummern:
        print(f"✅ {len(neue_nummern)} neue Einträge eingetragen.")
    else:
        print("✅ Keine neuen Einträge eingetragen.")

    log_rows = read_import_log_rows(sheets_service)
    imports_today = count_imports_for_date(log_rows, today)
    sheet_numbers = {normalize_number(row[0]) for row in filtered_rows if row}
    sheet_numbers.update(normalize_number(nummer) for nummer in neue_nummern)
    drive_numbers = set()
    for file in dateien:
        if not is_valid_drive_entry(file['name']):
            continue
        nummer, _ = extract_number_and_year(file['name'])
        if nummer:
            drive_numbers.add(nummer)
    sync_ok = drive_numbers.issubset(sheet_numbers)
    rb_count = count_rb_folders(drive_service)
    max_nummer = get_max_akten_nummer(filtered_rows, neue_nummern)
    year, kw = iso_year_week()

    try:
        update_tages_stat(
            sheets_service,
            sync_ok=sync_ok,
            rb_count=rb_count
        )
        update_wochen_stat(sheets_service, max_nummer)
        print(
            f"📊 Heute importiert: {imports_today} | Offen: {len(filtered_rows) + len(neue_nummern)} | "
            f"RB offen: {rb_count} | Sync: {'OK' if sync_ok else 'Offen'} | "
            f"KW {kw}/{year}: {max_nummer}"
        )
    except HttpError as exc:
        print(f'⚠️ Statistik konnte nicht aktualisiert werden: {exc}', file=sys.stderr)

    try:
        record_import_run(sheets_service)
    except HttpError as exc:
        print(f'⚠️ Skript-Laufzeit konnte nicht gespeichert werden: {exc}', file=sys.stderr)

    return 0

if __name__ == '__main__':
    raise SystemExit(main())
