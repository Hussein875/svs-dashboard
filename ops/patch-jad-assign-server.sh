#!/usr/bin/env bash
# Auf dem VPS ausführen (wo der Container "assign-service" läuft).
# Schaltet Jad für Drag-and-Drop + UX-Zuweisung frei und patched google.js.
set -euo pipefail

CONTAINER="${ASSIGN_CONTAINER:-assign-service}"

if ! docker ps --format '{{.Names}}' | grep -qx "$CONTAINER"; then
  echo "Container '$CONTAINER' läuft nicht." >&2
  docker ps --format '{{.Names}}' >&2 || true
  exit 1
fi

echo "▶ Aktuelle ASSIGN_COLUMNS:"
docker exec "$CONTAINER" node -e "
const fs=require('fs');
const s=fs.readFileSync('/app/assign-server.js','utf8');
const m=s.match(/ASSIGN_COLUMNS\\s*=\\s*new Set\\(\\[[^\\]]+\\]\\)/);
console.log(m ? m[0] : '(nicht gefunden)');
"

docker exec -i "$CONTAINER" node <<'NODE'
const fs = require('fs');

function patchAssignServer(src) {
  let s = src;
  s = s.replace(
    /ASSIGN_COLUMNS\s*=\s*new Set\(\[[^\]]*\]\)/,
    "ASSIGN_COLUMNS = new Set(['Ramazan', 'Robar', 'Jad'])",
  );
  s = s.replace(
    /column muss[^'`"]+/,
    'column muss Ramazan, Robar oder Jad sein',
  );
  return s;
}

function ensureWorkerMaps(src) {
  let s = src;
  if (!/^\s*jad:\s*'Jad'/m.test(s) && !/jad:\s*'Jad'/.test(s)) {
    if (/osama:\s*'Osama Sleiman',/.test(s)) {
      s = s.replace(
        /osama:\s*'Osama Sleiman',/,
        "osama: 'Osama Sleiman',\n  jad: 'Jad',",
      );
    } else if (/robar:\s*'Robar Kassem',/.test(s)) {
      s = s.replace(
        /robar:\s*'Robar Kassem',/,
        "robar: 'Robar Kassem',\n  jad: 'Jad',",
      );
    }
  }
  if (!/jad:\s*\[\s*'Jad'/.test(s)) {
    if (/osama:\s*\[[^\]]+\],/.test(s)) {
      s = s.replace(
        /osama:\s*\[[^\]]+\],/,
        "osama: ['Osama Sleiman', 'Herr Osama Sleiman'],\n  jad: ['Jad'],",
      );
    }
  }
  return s;
}

function ensureGoogleShortName(src) {
  let s = src;
  if (/Jad:\s*'Jad'/.test(s)) return s;
  if (/Robar:\s*'Robar Kassem',/.test(s)) {
    return s.replace(
      /Robar:\s*'Robar Kassem',/,
      "Robar: 'Robar Kassem',\n  Jad: 'Jad',",
    );
  }
  if (/SHORT_NAME_TO_SHEET\s*=\s*\{/.test(s)) {
    return s.replace(
      /SHORT_NAME_TO_SHEET\s*=\s*\{/,
      "SHORT_NAME_TO_SHEET = {\n  Jad: 'Jad',",
    );
  }
  return s;
}

for (const [file, fn] of [
  ['/app/assign-server.js', patchAssignServer],
  ['/app/assign-sachbearbeiter.js', ensureWorkerMaps],
  ['/app/google.js', ensureGoogleShortName],
]) {
  const before = fs.readFileSync(file, 'utf8');
  const after = fn(before);
  if (after !== before) {
    fs.writeFileSync(file, after);
    console.log('patched', file);
  } else {
    console.log('unchanged', file);
  }
}
NODE

echo "▶ Neustart $CONTAINER …"
docker restart "$CONTAINER" >/dev/null
sleep 2
docker exec "$CONTAINER" node -e "
const fs=require('fs');
const s=fs.readFileSync('/app/assign-server.js','utf8');
const m=s.match(/ASSIGN_COLUMNS\\s*=\\s*new Set\\(\\[[^\\]]+\\]\\)/);
console.log('Danach:', m ? m[0] : '(nicht gefunden)');
"
echo "✅ Fertig. Im Dashboard: Hard-Reload, Admin entsperren, Akte auf Jad ziehen."
