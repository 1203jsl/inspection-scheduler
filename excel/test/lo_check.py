"""Syntax-check each procedure of the VBA module with LibreOffice Basic (VBA mode),
then run pure-function tests from test/tests.bas.
usage: python3 lo_check.py"""
import uno, subprocess, time, os, re, sys, tempfile
from com.sun.star.beans import PropertyValue
here = os.path.dirname(os.path.abspath(__file__))
src = open(os.path.join(here, '..', 'vba', 'modInspection.bas')).read().replace('\r\n', '\n')
src = '\n'.join(l for l in src.split('\n') if not l.startswith('Attribute VB_'))
# split into declarations + procedures
lines = src.split('\n')
procs, decl, cur, name = [], [], None, None
start_re = re.compile(r'^(Public |Private )?(Sub|Function) (\w+)')
for ln in lines:
    m = start_re.match(ln)
    if m and cur is None:
        cur, name = [ln], m.group(3)
        if ln.rstrip().endswith('End Sub') or ln.rstrip().endswith('End Function'):
            procs.append((name, cur)); cur = None
        continue
    if cur is not None:
        cur.append(ln)
        if re.match(r'^End (Sub|Function)\s*$', ln):
            procs.append((name, cur)); cur = None
    else:
        decl.append(ln)
decl_txt = '\n'.join(l for l in decl if not l.startswith('Private gEvents'))  # class type unknown to LO
prof = tempfile.mkdtemp(prefix='loprof')
os.makedirs(prof + '/user')
open(prof + '/user/registrymodifications.xcu', 'w').write('<?xml version="1.0"?><oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema"></oor:items>')
state = {}
def start():
    port = 3100 + (os.getpid() + len(state.get('runs', []))) % 500
    state.setdefault('runs', []).append(port)
    state['proc'] = subprocess.Popen(['soffice', f'-env:UserInstallation=file://{prof}', '--headless', '--norestore',
                         f'--accept=socket,host=localhost,port={port};urp;'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    ctx = None
    for i in range(80):
        try:
            local = uno.getComponentContext()
            res = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
            ctx = res.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext"); break
        except Exception:
            time.sleep(0.5)
    gl = ctx.ServiceManager.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer", ctx)
    if not gl.hasByName("Chk"): gl.createLibrary("Chk")
    gl.loadLibrary("Chk")
    state['lib'] = gl.getByName("Chk")
    for m in state['lib'].getElementNames(): state['lib'].removeByName(m)
    state['sp'] = ctx.ServiceManager.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", ctx).createScriptProvider("")
start()
n = [0]
def compiles(body):
    n[0] += 1
    mod = f'M{n[0]}'
    lib, sp = state['lib'], state['sp']
    try:
        lib.insertByName(mod, 'Option VBASupport 1\n' + body + '\nFunction ZzPing() As String\n ZzPing = "ok"\nEnd Function\n')
        r = sp.getScript(f"vnd.sun.star.script:Chk.{mod}.ZzPing?language=Basic&location=application").invoke((), (), ())
        ok = r[0] == 'ok'
        lib.removeByName(mod)
        return ok
    except Exception as e:
        subprocess.run(['pkill', '-f', prof])
        time.sleep(1)
        start()
        return 'CRASH'
bad = []
if not compiles(decl_txt):
    bad.append('(declarations)')
for nm, body in ([] if '--tests-only' in sys.argv else procs):
    r = compiles(decl_txt + '\n' + '\n'.join(body))
    tries = 0
    while r == 'CRASH' and tries < 2:
        tries += 1
        r = compiles(decl_txt + '\n' + '\n'.join(body))
    if r is not True:
        bad.append(nm + (' (crash)' if r == 'CRASH' else ''))
        if r is False:
            struct = re.compile(r'^\s*(If\b.*Then\s*$|ElseIf\b|Else\s*$|End\b|For\b|Next\b|Do\b|Loop\b|Select\b|Case\b|With\b|While\b|Wend\b|\w+:\s*$)')
            simple = [i for i, l in enumerate(body) if i > 0 and not struct.match(l) and not l.rstrip().endswith('_') and not (i > 0 and body[i-1].rstrip().endswith('_'))]
            cand = simple
            while len(cand) > 1:
                half = cand[:len(cand)//2]
                trial = [('' if i in half else l) for i, l in enumerate(body)]
                if compiles(decl_txt + '\n' + '\n'.join(trial)) is True:
                    cand = half
                else:
                    cand = cand[len(cand)//2:]
            if cand:
                trial = [('' if i == cand[0] else l) for i, l in enumerate(body)]
                if compiles(decl_txt + '\n' + '\n'.join(trial)) is True:
                    print('   ', nm, 'line:', body[cand[0]].strip())
                else:
                    print('   ', nm, 'error in a structural line or multi-line statement')
print(f'procedures: {len(procs)}, failed: {len(bad)}')
for b in bad: print('  FAIL', b)
# pure-function tests
tests = os.path.join(here, 'tests.bas')
if os.path.exists(tests) and '--no-tests' not in sys.argv:
    lib, sp = state['lib'], state['sp']
    stubbed = {'T', 'SetV'}
    body_all = '\n'.join('\n'.join(b) for nm, b in procs if nm not in stubbed)
    if os.environ.get('DBG'):
        fn = os.environ['DBG']
        kw = 'Function' if ('Function ' + fn + '(') in body_all else 'Sub'
        i0 = body_all.index(kw + ' ' + fn + '(')
        j = body_all.index('\n', i0)
        while body_all[:j].rstrip().endswith('_'):
            j = body_all.index('\n', j + 1)
        body_all = body_all[:j] + '\n    On Error GoTo dbgeh' + body_all[j:]
        k = body_all.index('\nEnd ' + kw, i0)
        body_all = body_all[:k] + '\n    Exit ' + kw + '\ndbgeh:\n    gFails = gFails & " DBG ' + fn + ' " & Err.Description & " line " & Erl' + body_all[k:]
        open('/tmp/claude-0/-home-user/63ac5177-41fc-5ffc-89f3-5843df79f64b/scratchpad/tt.bas','w').write(body_all)
    tl = open(tests).read().split('\n')
    tdecl = '\n'.join(l for l in tl if l.startswith('Private g'))
    trest = '\n'.join(l for l in tl if not l.startswith('Private g'))
    lib.insertByName('TT', 'Option VBASupport 1\n' + decl_txt + '\n' + tdecl + '\n' + body_all + '\n' + trest)
    r = sp.getScript("vnd.sun.star.script:Chk.TT.RunTests?language=Basic&location=application").invoke((), (), ())
    print('TESTS:', r[0])
subprocess.run(['pkill', '-f', prof])
