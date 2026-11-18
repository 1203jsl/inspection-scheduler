"""Static check: local variables/parameters must not shadow module procedures or constants (VBA is case-insensitive)."""
import re, sys, os
here = os.path.dirname(os.path.abspath(__file__))
src = open(os.path.join(here, '..', 'vba', 'modInspection.bas')).read().replace('\r\n', '\n')
# join continuation lines
src = re.sub(r' _\n\s*', ' ', src)
procs = re.findall(r'^(?:Public |Private )?(?:Sub|Function) (\w+)\((.*?)\)(?: As \w+)?\s*$(.*?)^End (?:Sub|Function)', src, re.M | re.S)
names = {p[0].lower() for p in procs}
consts = {m.lower() for m in re.findall(r'^Public Const (\w+)', src, re.M)}
globals_ = {m.lower() for m in re.findall(r'^(?:Public|Private) (\w+) As ', src, re.M)}
known = names | consts
bad = 0
for nm, params, body in procs:
    locs = set()
    for p in params.split(','):
        m = re.search(r'(\w+)(?:\(\))?\s+As\b', p) or re.search(r'(?:ByVal|ByRef|Optional|ParamArray)\s+(\w+)', p) or re.match(r'\s*(\w+)\s*$', p)
        if m: locs.add(m.group(1))
    for d in re.findall(r'^\s*Dim (.*)$', body, re.M):
        for part in re.split(r',(?![^(]*\))', d):
            m = re.match(r'\s*(\w+)', part)
            if m: locs.add(m.group(1))
    for l in locs:
        if l.lower() in known and l.lower() != nm.lower():
            print(f'{nm}: local "{l}" shadows module name')
            bad += 1
        if l.lower() in globals_:
            print(f'{nm}: local "{l}" shadows module variable')
            bad += 1
print('shadowing problems:', bad)
sys.exit(1 if bad else 0)
