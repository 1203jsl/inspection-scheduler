#!/usr/bin/env bash
# Builds dist/검수일정관리_엑셀.zip (workbook + VBA files + guide)
set -euo pipefail
cd "$(dirname "$0")"
rm -rf dist && mkdir -p dist/pkg
cat vba/part1_core.bas vba/part2_logic.bas vba/part3_events.bas vba/part4_views.bas vba/part5_mail.bas vba/part6_import.bas vba/part7_setup.bas > vba/modInspection.bas
python3 test/name_check.py >/dev/null
python3 build_workbook.py "dist/pkg/검수일정관리.xlsx"
# VBA import needs CRLF line endings
python3 - <<'PY'
import os
for src, dst in (('vba/modInspection.bas', 'dist/pkg/modInspection.bas'), ('vba/clsAppEvents.cls', 'dist/pkg/clsAppEvents.cls')):
    s = open(src, encoding='ascii').read().replace('\r\n', '\n').replace('\n', '\r\n')
    open(dst, 'w', encoding='ascii', newline='').write(s)
s = open('help.txt', encoding='utf-8').read().replace('\n', '\r\n')
open('dist/pkg/설치·사용 안내.txt', 'w', encoding='utf-8-sig', newline='').write(s)
import zipfile
with zipfile.ZipFile('dist/검수일정관리_엑셀.zip', 'w', zipfile.ZIP_DEFLATED) as z:
    for f in sorted(os.listdir('dist/pkg')):
        z.write(os.path.join('dist/pkg', f), f)
print('zip:', sorted(os.listdir('dist/pkg')))
PY
