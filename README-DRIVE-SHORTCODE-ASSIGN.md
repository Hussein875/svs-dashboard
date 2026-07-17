# Drive-Kürzel → Bearbeiter (Auto-Assign)

Wenn ein neuer Gutachten-Ordner in Drive ankommt und der UltraExpert-Bot die Akte öffnet, wird das Kürzel in Klammern am Ordnernamen ausgewertet.

| Kürzel | Bearbeiter | Dashboard |
|--------|------------|-----------|
| `RO` | Robar Kassem | Spalte Robar |
| `HB` | Hussein Selman | Badge B |
| `MZ` | Mohamed Zahreddine | Badge M |
| `HJ` | Hussein Jaber | Badge HJ |
| sonst (`OS`, `HA`, `HK`, …) | keine Zuweisung | — |

## Komponenten

- `code/UX-Watcher/assign-server.js` — API `POST /api/assign-from-folder`
- `code/UX-Watcher/assign-sachbearbeiter.js` — UX „Weitere Sachbearbeiter“ setzen
- `code/UX-Watcher/google.js` — Sheet-Bearbeiter schreiben (legt Zeile ggf. neu an)
- `ultraexpert-login-bot/server.mjs` — nach Akte-Erstellung Assign-API aufrufen
- `import_gutachten_to_sheet.py` — schreibt Bearbeiter aus Kürzel auch beim Sheet-Import

## Env

Beide Services brauchen denselben Wert:

```env
FOLDER_ASSIGN_SECRET=...
```

Deploy-Hinweise: siehe `deploy/assign-service/`.
