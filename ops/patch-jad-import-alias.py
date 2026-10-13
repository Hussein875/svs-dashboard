"""Einzeiler für den Import-Cron (svs-dashboard-backend).

In import_gutachten_to_sheet.py bei SHEET_ASSIGNEE_ALIASES ergänzen:

    'jad': 'Jad',

Sonst leert sync_sheet_assignees den Bearbeiter „Jad“ wieder.
"""

from pathlib import Path

path = Path('import_gutachten_to_sheet.py')
text = path.read_text(encoding='utf-8')
needle = "    'robar kassam': 'Robar',\n"
insert = "    'robar kassam': 'Robar',\n    'jad': 'Jad',\n"
if "'jad': 'Jad'" in text or '"jad": "Jad"' in text:
    print('bereits vorhanden')
elif needle in text:
    path.write_text(text.replace(needle, insert, 1), encoding='utf-8')
    print('ok: jad Alias ergänzt')
else:
    raise SystemExit('Insert-Stelle nicht gefunden – bitte manuell ergänzen')
