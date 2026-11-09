import { useEffect, useRef, useState } from 'react';
import type { AppErrorInfo, BackupEntry, ConflictRules, CurrentUser, DataFiles, MailKind, MailSettings, MailTemplate, RuleNum, SmtpAccountInfo } from '../shared/types';
import { MAIL_KIND_LABEL } from '../shared/types';
import { DEFAULT_MAIL_SETTINGS, DEFAULT_TEMPLATES, PLACEHOLDERS, parseAddresses } from '../shared/mail';
import { DEFAULT_CONFLICT_RULES } from '../shared/conflicts';
import { fmtYMD } from '../shared/dates';
import { api, fmtDateTime, type LockStatus } from './api';
import { Alert, Dialog, ErrorAlert, Icon } from './components';

/**
 * 설정: 구역 탭 + 화면 맨 아래 [설정 저장] 하나 (디자인 ScreenSettings).
 * 알림 탭은 7단계에서 추가합니다.
 */
type Tab = 'conflict' | 'mailTpl' | 'mailSend' | 'excel' | 'folder' | 'user';
const TABS: { id: Tab; label: string }[] = [
  { id: 'conflict', label: '충돌 조건' },
  { id: 'mailTpl', label: '메일 템플릿' },
  { id: 'mailSend', label: '메일 발송 설정' },
  { id: 'excel', label: '엑셀 양식' },
  { id: 'folder', label: '데이터 폴더' },
  { id: 'user', label: '사용자' },
];
/** 저장 후 다시 그려도 보던 탭 유지 */
let lastTab: Tab = 'conflict';

const NUM_RULES: { key: 'maxPerDay' | 'maxPerInspector' | 'maxQtyPerDay' | 'maxPerVendor'; label: string; unit: string }[] = [
  { key: 'maxPerDay', label: '하루 최대 검수 건수', unit: '건' },
  { key: 'maxPerInspector', label: '검수자 1인당 하루 최대 건수', unit: '건' },
  { key: 'maxQtyPerDay', label: '하루 최대 검수 수량 합계', unit: '개' },
  { key: 'maxPerVendor', label: '같은 업체 하루 최대 건수', unit: '건' },
];

function Switch({ label, checked, onChange }: { label: string; checked: boolean; onChange: (v: boolean) => void }) {
  return (
    <>
      <input type="checkbox" className="switch" role="switch" aria-label={label} checked={checked} onChange={(e) => onChange(e.target.checked)} />
      <span className="nm">{label}</span>
      <span className="muted" style={{ width: 48 }}>{checked ? '켜짐' : '꺼짐'}</span>
    </>
  );
}

export function SettingsPage({ data, user, dataDir, onChangeFolder, onSaved, onSaveError, onLockChange }: {
  data: DataFiles;
  user: CurrentUser;
  dataDir: string;
  onChangeFolder: () => void;
  onSaved: (message?: string) => void;
  onSaveError: (e: AppErrorInfo) => AppErrorInfo | null;
  onLockChange: (l: LockStatus) => void;
}) {
  const s = data.settings;
  const [tab, setTabState] = useState<Tab>(lastTab);
  const setTab = (t: Tab) => { lastTab = t; setTabState(t); };
  const [userMap, setUserMap] = useState<[string, string][]>(Object.entries(s.userMap ?? {}));
  const [lockMin, setLockMin] = useState(String(s.lockStaleMinutes));
  const [keep, setKeep] = useState(String(s.backupKeep));
  const [items, setItems] = useState<string[]>(s.inspectionItems ?? []);
  const [newItem, setNewItem] = useState('');
  const [rules, setRules] = useState<ConflictRules>(structuredClone(s.conflictRules ?? DEFAULT_CONFLICT_RULES));
  const [numText, setNumText] = useState<Record<string, string>>(
    Object.fromEntries(NUM_RULES.map((r) => [r.key, String((s.conflictRules ?? DEFAULT_CONFLICT_RULES)[r.key].value)])),
  );
  const [templates, setTemplates] = useState<Record<MailKind, MailTemplate>>(structuredClone(s.mailTemplates ?? DEFAULT_TEMPLATES));
  const [tplKind, setTplKind] = useState<MailKind>('registered');
  const [mail, setMail] = useState<MailSettings>(structuredClone(s.mail ?? DEFAULT_MAIL_SETTINGS));
  const [portText, setPortText] = useState(String((s.mail ?? DEFAULT_MAIL_SETTINGS).smtp.port));
  const [acct, setAcct] = useState<SmtpAccountInfo | null>(null);
  const [acctText, setAcctText] = useState('');
  const [pw, setPw] = useState('');
  const [clearPw, setClearPw] = useState(false);
  const [testing, setTesting] = useState(false);
  const subjRef = useRef<HTMLInputElement>(null);
  const bodyRef = useRef<HTMLTextAreaElement>(null);
  const lastField = useRef<'subject' | 'body'>('body');
  const [hDate, setHDate] = useState('');
  const [hName, setHName] = useState('');
  const [dirty, setDirty] = useState(false);
  const [error, setError] = useState<AppErrorInfo | null>(null);
  const [notice, setNotice] = useState('');
  const [backups, setBackups] = useState<BackupEntry[]>([]);
  const [restoreTarget, setRestoreTarget] = useState<BackupEntry | null>(null);

  useEffect(() => {
    api.listBackups().then((r) => r.ok && setBackups(r.value));
  }, [data]);
  useEffect(() => {
    api.mailAccount().then((r) => { if (r.ok) { setAcct(r.value); setAcctText(r.value.account); } });
  }, []);

  const setTpl = (k: MailKind, t: MailTemplate) => { touch(); setTemplates({ ...templates, [k]: t }); };
  /** 치환값 칩: 마지막으로 누른 칸(제목/본문)의 커서 위치에 넣음 */
  const insertPh = (ph: string) => {
    const el = lastField.current === 'subject' ? subjRef.current : bodyRef.current;
    const t = templates[tplKind];
    const key = lastField.current;
    const cur = t[key];
    const a = el?.selectionStart ?? cur.length, b = el?.selectionEnd ?? cur.length;
    setTpl(tplKind, { ...t, [key]: cur.slice(0, a) + ph + cur.slice(b) });
    requestAnimationFrame(() => { el?.focus(); el?.setSelectionRange(a + ph.length, a + ph.length); });
  };
  const setM = (m: Partial<MailSettings>) => { touch(); setMail({ ...mail, ...m }); };
  const acctChanged = !!acct && (acctText.trim() !== acct.account || pw !== '' || clearPw);

  /** 처음 고칠 때 "편집 중" 표시 */
  const touch = () => {
    setNotice('');
    if (!dirty) {
      setDirty(true);
      api.lockAcquire().then((r) => r.ok && onLockChange(r.value));
    }
  };
  const setRule = <K extends keyof ConflictRules>(k: K, v: ConflictRules[K]) => { touch(); setRules({ ...rules, [k]: v }); };

  const lockN = Number(lockMin);
  const keepN = Number(keep);
  const numErr = (k: string) => {
    const n = Number(numText[k]);
    return rules[k as keyof ConflictRules] && (rules[k as keyof ConflictRules] as RuleNum).enabled && !(Number.isInteger(n) && n >= 1)
      ? '1 이상의 정수로 입력하세요.' : '';
  };
  const invalid = {
    lock: !Number.isInteger(lockN) || lockN < 5 || lockN > 1440 ? '5~1440 사이의 정수(분)로 입력하세요.' : '',
    keep: !Number.isInteger(keepN) || keepN < 1 || keepN > 500 ? '1~500 사이의 정수로 입력하세요.' : '',
    map: userMap.some(([a, n]) => !a.trim() || !n.trim()) ? '계정과 표시 이름을 모두 입력하세요. 필요 없는 줄은 삭제하세요.' : '',
    items: items.length === 0 ? '검수 품목을 하나 이상 등록하세요. 빈 양식의 드롭다운에 쓰입니다.' : '',
    rules: NUM_RULES.some((r) => numErr(r.key)) ? '충돌 조건의 숫자를 확인하세요.' : '',
    tpl: (Object.keys(templates) as MailKind[]).filter((k) => !templates[k].subject.trim()).map((k) => MAIL_KIND_LABEL[k]).join(', ')
      ? `제목이 비어 있는 템플릿이 있습니다: ${(Object.keys(templates) as MailKind[]).filter((k) => !templates[k].subject.trim()).map((k) => MAIL_KIND_LABEL[k]).join(', ')}` : '',
    port: !(Number.isInteger(Number(portText)) && Number(portText) >= 1 && Number(portText) <= 65535) ? '1~65535 사이의 숫자로 입력하세요. (보통 SSL/TLS=465, STARTTLS=587)' : '',
    webmail: mail.webmailUrl.trim() && !/^https?:\/\/\S+$/.test(mail.webmailUrl.trim()) ? 'http:// 또는 https:// 로 시작하는 주소로 입력하세요.' : '',
    cc: parseAddresses(mail.defaultCc).bad.length ? `이메일 형식이 올바르지 않습니다: ${parseAddresses(mail.defaultCc).bad.join(', ')}` : '',
    acct: acctText.trim() && parseAddresses(acctText).bad.length ? '메일 계정은 이메일 주소로 입력하세요. 예: hong@company.co.kr' : '',
  };
  const tabHasError: Record<Tab, boolean> = {
    conflict: !!invalid.rules, mailTpl: !!invalid.tpl, mailSend: !!(invalid.port || invalid.webmail || invalid.cc || invalid.acct),
    excel: !!invalid.items, folder: !!(invalid.lock || invalid.keep), user: !!invalid.map,
  };
  const hasInvalid = Object.values(invalid).some(Boolean);

  const trimmed = newItem.trim();
  const dupItem = trimmed && items.includes(trimmed) ? '이미 있는 품목입니다.' : '';
  const addItem = () => {
    if (!trimmed || dupItem) return;
    touch();
    setItems([...items, trimmed]);
    setNewItem('');
  };

  const dupHoliday = hDate && rules.holidays.dates.some((h) => h.date === hDate) ? '이미 등록된 날짜입니다.' : '';
  const addHoliday = () => {
    if (!hDate || dupHoliday) return;
    setRule('holidays', { ...rules.holidays, dates: [...rules.holidays.dates, { date: hDate, name: hName.trim() }].sort((a, b) => a.date.localeCompare(b.date)) });
    setHDate('');
    setHName('');
  };

  const save = async () => {
    setError(null);
    const map: Record<string, string> = {};
    for (const [a, n] of userMap) map[a.trim()] = n.trim();
    const nextRules: ConflictRules = { ...rules };
    for (const r of NUM_RULES) {
      const n = Number(numText[r.key]);
      nextRules[r.key] = { ...rules[r.key], value: Number.isInteger(n) && n >= 1 ? n : rules[r.key].value };
    }
    const nextMail: MailSettings = {
      ...mail, webmailUrl: mail.webmailUrl.trim(), defaultCc: mail.defaultCc.trim(),
      smtp: { ...mail.smtp, host: mail.smtp.host.trim(), port: Number(portText) },
    };
    const r = await api.save({
      settings: {
        ...s, userMap: map, lockStaleMinutes: lockN, backupKeep: keepN, inspectionItems: items, conflictRules: nextRules,
        mailTemplates: templates, mail: nextMail,
      },
    });
    if (r.ok) {
      // 메일 계정·비밀번호는 이 PC에만 (공유 설정 파일에는 넣지 않음)
      let acctMsg = '';
      if (acctChanged) {
        const a = await api.mailSaveAccount({ account: acctText.trim(), password: clearPw ? '' : pw || undefined });
        if (a.ok) { setAcct(a.value); setPw(''); setClearPw(false); }
        else {
          acctMsg = ` 단, ${a.error.message}`;
          const info = await api.mailAccount();
          if (info.ok) setAcct(info.value);
        }
      }
      await api.lockRelease();
      setDirty(false);
      onSaved('설정을 저장했습니다.' + acctMsg);
    } else {
      setError(onSaveError(r.error));
    }
  };

  const cancel = async () => {
    await api.lockRelease();
    onSaved(); // 다시 불러와 원래 값으로
  };

  const doRestore = async () => {
    if (!restoreTarget) return;
    const r = await api.restoreBackup(restoreTarget.id);
    setRestoreTarget(null);
    if (r.ok) onSaved(`${fmtDateTime(restoreTarget.createdAt)} 시점으로 복원했습니다. 복원 전 상태도 백업해 두었습니다.`);
    else setError(r.error);
  };

  const downloadTemplate = async () => {
    const r = await api.saveTemplate();
    if (!r.ok) setError(r.error);
    else if (r.value) setNotice(`빈 양식을 저장했습니다: ${r.value}`);
  };

  return (
    <>
      {error && <ErrorAlert error={error} />}
      {notice && <Alert kind="ok">{notice}</Alert>}

      <div className="tabs" role="tablist">
        {TABS.map((t) => (
          <button key={t.id} className="tab" role="tab" type="button" aria-selected={tab === t.id} onClick={() => setTab(t.id)}>
            {t.label}{tabHasError[t.id] && <Icon name="error" style={{ marginLeft: 4, color: 'var(--danger-fg)' }} />}
          </button>
        ))}
      </div>

      {tab === 'conflict' && (
        <div className="card" role="tabpanel">
          <div className="card-h">충돌 조건</div>
          <div className="help" style={{ marginBottom: 4 }}>
            조건을 넘는 날짜는 달력·목록·가져오기에서 충돌 경고로 표시되고, 확정·변경할 때 경고합니다. 취소된 건은 세지 않습니다.
          </div>
          {NUM_RULES.slice(0, 2).map(numRule)}
          <div className="rule">
            <Switch label="같은 시간대·같은 장소 중복 금지" checked={rules.sameSlot.enabled} onChange={(v) => setRule('sameSlot', { ...rules.sameSlot, enabled: v })} />
            <span className="help" style={{ width: 260 }}>시작 시각이 {rules.sameSlot.minutes}분 안쪽이면 같은 시간대</span>
          </div>
          {NUM_RULES.slice(2).map(numRule)}
          <div className="rule" style={{ alignItems: 'flex-start' }}>
            <Switch label="검수 불가일(휴일) 지정" checked={rules.holidays.enabled} onChange={(v) => setRule('holidays', { ...rules.holidays, enabled: v })} />
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8, maxWidth: 560 }}>
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                {rules.holidays.dates.length === 0 && <span className="help">등록된 날짜가 없습니다.</span>}
                {rules.holidays.dates.map((h) => (
                  <span key={h.date} className="tagchip">{fmtYMD(h.date)}{h.name ? ` ${h.name}` : ''}
                    <button type="button" aria-label={`${h.date} 삭제`} onClick={() => setRule('holidays', { ...rules.holidays, dates: rules.holidays.dates.filter((x) => x.date !== h.date) })}>×</button>
                  </span>
                ))}
              </div>
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                <input type="date" className={`input${dupHoliday ? ' invalid' : ''}`} style={{ width: 160 }} aria-label="추가할 날짜" value={hDate} onChange={(e) => setHDate(e.target.value)} />
                <input className="input" style={{ width: 160 }} aria-label="이름 (선택)" placeholder="이름 (예: 성탄절)" value={hName}
                  onChange={(e) => setHName(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && addHoliday()} />
                <button className="btn btn-sm" type="button" style={{ height: 32 }} onClick={addHoliday} disabled={!hDate || !!dupHoliday}>추가</button>
              </div>
              {dupHoliday && <span className="help help-err"><Icon name="error" />{dupHoliday}</span>}
            </div>
          </div>
          <div className="rule">
            <Switch label="충돌이 있어도 사유를 입력하면 강제 진행 허용" checked={rules.allowForce} onChange={(v) => setRule('allowForce', v)} />
            <span className="help" style={{ width: 260 }}>{rules.allowForce ? '사유는 변경 이력에 남습니다' : '충돌이 있으면 확정·변경할 수 없습니다'}</span>
          </div>
        </div>
      )}

      {tab === 'mailTpl' && (
        <div className="card" role="tabpanel" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          <div className="card-h" style={{ margin: 0 }}>메일 템플릿 <span className="help" style={{ fontWeight: 400 }}>업체 담당자에게 가는 메일. 보내기 전 미리보기에서 다시 고칠 수 있습니다.</span></div>
          <div className="tabs" role="tablist" aria-label="메일 종류">
            {(Object.keys(MAIL_KIND_LABEL) as MailKind[]).map((k) => (
              <button key={k} className="tab" role="tab" type="button" aria-selected={tplKind === k} onClick={() => setTplKind(k)}>
                {MAIL_KIND_LABEL[k]}{!templates[k].subject.trim() && <Icon name="error" style={{ marginLeft: 4, color: 'var(--danger-fg)' }} />}
              </button>
            ))}
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) 280px', gap: 16 }}>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              <div className="field">
                <label htmlFor="ts">제목</label>
                <input id="ts" ref={subjRef} spellCheck={false} className={`input${!templates[tplKind].subject.trim() ? ' invalid' : ''}`} value={templates[tplKind].subject}
                  onFocus={() => (lastField.current = 'subject')} onChange={(e) => setTpl(tplKind, { ...templates[tplKind], subject: e.target.value })} />
              </div>
              <div className="field">
                <label htmlFor="tb">본문</label>
                <textarea id="tb" ref={bodyRef} spellCheck={false} className="input" style={{ height: 260, lineHeight: '22px' }} value={templates[tplKind].body}
                  onFocus={() => (lastField.current = 'body')} onChange={(e) => setTpl(tplKind, { ...templates[tplKind], body: e.target.value })} />
              </div>
              <div className="btn-row">
                <button className="btn btn-sm" type="button" onClick={() => setTpl(tplKind, { ...DEFAULT_TEMPLATES[tplKind] })}>기본 문구로 되돌리기</button>
              </div>
            </div>
            <div className="field">
              <span className="lbl">사용 가능한 치환값</span>
              <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                {PLACEHOLDERS.map((ph) => (
                  <button key={ph} type="button" className="tag" style={{ cursor: 'pointer', background: 'var(--surface)' }} onClick={() => insertPh(ph)}>{ph}</button>
                ))}
              </div>
              <span className="help">누르면 마지막으로 누른 칸(제목/본문)의 커서 위치에 들어갑니다. {'{결과}·{불합격사유}'}는 합격/불합격 메일에 씁니다.</span>
            </div>
          </div>
        </div>
      )}

      {tab === 'mailSend' && (
        <div className="card" role="tabpanel" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          <div className="card-h" style={{ margin: 0 }}>메일 발송 설정</div>
          <fieldset className="field" style={{ border: 0, margin: 0, padding: 0 }}>
            <legend className="lbl" style={{ padding: '0 0 6px' }}>발송 방식</legend>
            <label className="radio"><input type="radio" name="mm" checked={mail.method === 'smtp'} onChange={() => setM({ method: 'smtp' })} />자동 SMTP (미리보기 후 [발송])</label>
            <label className="radio"><input type="radio" name="mm" checked={mail.method === 'manual'} onChange={() => setM({ method: 'manual' })} />수동 복사 (웹메일에 붙여넣기)</label>
            <span className="help">어느 방식이든 업체 메일은 미리보기를 거쳐야 보냅니다. 모든 PC에 같이 적용됩니다.</span>
          </fieldset>
          <div className="form-row" style={{ maxWidth: 760 }}>
            <div className="field">
              <label htmlFor="sh">SMTP 서버</label>
              <input id="sh" className="input" disabled={mail.method !== 'smtp'} placeholder="사내 IT에 확인 (예: smtp.회사도메인)" value={mail.smtp.host}
                onChange={(e) => setM({ smtp: { ...mail.smtp, host: e.target.value } })} />
            </div>
            <div className="field">
              <label htmlFor="sp">포트 · 보안 방식</label>
              <div style={{ display: 'flex', gap: 8 }}>
                <input id="sp" className={`input${invalid.port ? ' invalid' : ''}`} style={{ width: 96 }} inputMode="numeric" disabled={mail.method !== 'smtp'} value={portText}
                  onChange={(e) => { touch(); setPortText(e.target.value); }} />
                <select className="sel" aria-label="보안 방식" disabled={mail.method !== 'smtp'} value={mail.smtp.security}
                  onChange={(e) => setM({ smtp: { ...mail.smtp, security: e.target.value as MailSettings['smtp']['security'] } })}>
                  <option value="starttls">STARTTLS (보통 587)</option>
                  <option value="ssl">SSL/TLS (보통 465)</option>
                  <option value="none">암호화 없음 (사내 전용 서버만)</option>
                </select>
              </div>
              {invalid.port && <span className="help help-err"><Icon name="error" />{invalid.port}</span>}
            </div>
            <div className="field">
              <label htmlFor="sa">메일 계정 (이 PC)</label>
              <input id="sa" className={`input${invalid.acct ? ' invalid' : ''}`} disabled={mail.method !== 'smtp'} placeholder="예: hong@company.co.kr" value={acctText}
                onChange={(e) => { touch(); setAcctText(e.target.value); }} />
              {invalid.acct ? <span className="help help-err"><Icon name="error" />{invalid.acct}</span> : <span className="help">보내는 사람 주소로도 쓰입니다. PC마다 따로 저장합니다.</span>}
            </div>
            <div className="field">
              <label htmlFor="sw">비밀번호 (이 PC)</label>
              <input id="sw" className="input" type="password" autoComplete="off" disabled={mail.method !== 'smtp' || clearPw}
                placeholder={acct?.hasPassword ? '저장됨 · 바꿀 때만 입력' : '입력하세요'} value={pw}
                onChange={(e) => { touch(); setPw(e.target.value); }} />
              <span className="help">
                Windows 계정으로 암호화해 이 PC에만 저장하고, 공유 폴더·기록에는 남기지 않습니다.
                {acct && !acct.encryptionAvailable && ' 이 PC에서는 암호화를 쓸 수 없어 저장할 수 없습니다.'}
              </span>
              {acct?.hasPassword && (
                <label className="switch-row"><input type="checkbox" className="chk" checked={clearPw} onChange={(e) => { touch(); setClearPw(e.target.checked); setPw(''); }} />저장된 비밀번호 지우기</label>
              )}
            </div>
          </div>
          <div className="btn-row">
            <button className="btn" type="button" disabled={mail.method !== 'smtp' || dirty || testing} onClick={async () => {
              setTesting(true); setError(null); setNotice('');
              const r = await api.mailTest();
              setTesting(false);
              if (r.ok) setNotice(`테스트 메일을 ${r.value.to}(으)로 보냈습니다. 받은편지함을 확인하세요.`);
              else setError(r.error);
            }}>{testing ? '보내는 중…' : '테스트 메일 보내기'}</button>
            <span className="help">{dirty ? '바꾼 설정을 먼저 저장한 뒤 시험하세요.' : '설정한 계정으로 본인에게 시험 메일을 보냅니다.'}</span>
          </div>
          <div className="form-row" style={{ maxWidth: 760 }}>
            <div className="field">
              <label htmlFor="wu">웹메일 주소 (수동 발송용)</label>
              <input id="wu" className={`input${invalid.webmail ? ' invalid' : ''}`} placeholder="https://" value={mail.webmailUrl} onChange={(e) => setM({ webmailUrl: e.target.value })} />
              {invalid.webmail && <span className="help help-err"><Icon name="error" />{invalid.webmail}</span>}
            </div>
            <div className="field">
              <label htmlFor="dc">업체 메일 기본 참조</label>
              <input id="dc" className={`input${invalid.cc ? ' invalid' : ''}`} placeholder="예: quality@company.co.kr (여러 명은 쉼표)" value={mail.defaultCc} onChange={(e) => setM({ defaultCc: e.target.value })} />
              {invalid.cc && <span className="help help-err"><Icon name="error" />{invalid.cc}</span>}
            </div>
          </div>
        </div>
      )}

      {tab === 'excel' && (
        <div className="card" role="tabpanel" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          <div className="card-h" style={{ margin: 0 }}>엑셀 양식</div>
          <div className="field">
            <span className="lbl">검수 품목 선택 목록</span>
            <span className="help">빈 양식의 드롭다운 목록, 가져오기 때 '알 수 없는 품목' 점검에 쓰입니다.</span>
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
              {items.map((it) => (
                <span key={it} className="tagchip">{it}
                  <button type="button" aria-label={`${it} 삭제`} onClick={() => { touch(); setItems(items.filter((x) => x !== it)); }}>×</button>
                </span>
              ))}
              <input className={`input${dupItem ? ' invalid' : ''}`} style={{ width: 160 }} aria-label="추가할 검수 품목" placeholder="새 품목" value={newItem}
                onChange={(e) => setNewItem(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && addItem()} />
              <button className="btn btn-sm" type="button" style={{ height: 32 }} onClick={addItem} disabled={!trimmed || !!dupItem}>추가</button>
            </div>
            {(dupItem || invalid.items) && <span className="help help-err"><Icon name="error" />{dupItem || invalid.items}</span>}
          </div>
          <div className="btn-row">
            <button className="btn" type="button" onClick={downloadTemplate} disabled={dirty}><Icon name="download" />빈 양식 내려받기</button>
            <span className="help">{dirty ? '바꾼 설정을 먼저 저장해야 새 목록으로 양식이 만들어집니다.' : '목록을 바꾼 뒤에는 새 양식을 만들어 업체에 보내세요.'}</span>
          </div>
        </div>
      )}

      {tab === 'folder' && (
        <>
          <div className="card" role="tabpanel" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            <div className="card-h" style={{ margin: 0 }}>데이터 폴더 <span className="help" style={{ fontWeight: 400 }}>폴더 위치는 이 PC에만 저장</span></div>
            <div className="field">
              <label htmlFor="dd">클라우드독 동기화 폴더 경로</label>
              <div style={{ display: 'flex', gap: 8 }}>
                <input id="dd" className="input" value={dataDir} readOnly style={{ flex: 1 }} />
                <button className="btn" type="button" onClick={onChangeFolder}>폴더 선택</button>
              </div>
            </div>
            <div className="form-row" style={{ maxWidth: 760 }}>
              <div className="field">
                <label htmlFor="bk">백업 보관 개수</label>
                <input id="bk" className={`input${invalid.keep ? ' invalid' : ''}`} value={keep} inputMode="numeric" style={{ width: 96 }}
                  onChange={(e) => { touch(); setKeep(e.target.value); }} />
                {invalid.keep ? <span className="help help-err"><Icon name="error" />{invalid.keep}</span>
                  : <span className="help">저장할 때마다 이전 상태를 백업/ 에 남기고, 개수를 넘으면 오래된 것부터 지웁니다.</span>}
              </div>
              <div className="field">
                <label htmlFor="lk">'편집 중' 표시 유효 시간 (분)</label>
                <input id="lk" className={`input${invalid.lock ? ' invalid' : ''}`} value={lockMin} inputMode="numeric" style={{ width: 96 }}
                  onChange={(e) => { touch(); setLockMin(e.target.value); }} />
                {invalid.lock ? <span className="help help-err"><Icon name="error" />{invalid.lock}</span>
                  : <span className="help">이 시간 넘게 갱신이 없는 '편집 중' 표시는 무시합니다.</span>}
              </div>
            </div>
          </div>
          <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            <div className="card-h" style={{ margin: 0 }}>백업에서 복원 <span className="help" style={{ fontWeight: 400 }}>최근 {backups.length}개</span></div>
            {backups.length === 0 ? (
              <span className="help">아직 백업이 없습니다. 처음 저장할 때부터 생깁니다.</span>
            ) : (
              <div className="tbl-wrap" style={{ maxHeight: 320, overflow: 'auto' }}>
                <table className="tbl">
                  <thead><tr><th>백업 시각</th><th>저장한 사람</th><th>폴더 이름</th><th /></tr></thead>
                  <tbody>
                    {backups.map((b) => (
                      <tr key={b.id}>
                        <td>{fmtDateTime(b.createdAt)}</td>
                        <td>{b.createdBy}</td>
                        <td className="muted">{b.id}</td>
                        <td><button className="btn btn-sm" type="button" onClick={() => setRestoreTarget(b)}>이 시점으로 복원</button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </>
      )}

      {tab === 'user' && (
        <div className="card" role="tabpanel" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          <div className="card-h" style={{ margin: 0 }}>사용자 <span className="help" style={{ fontWeight: 400 }}>Windows 계정 → 변경 이력과 협의 기록의 작성자로 보일 이름 (모든 PC 공통)</span></div>
          <div className="tbl-wrap">
            <table className="tbl">
              <thead><tr><th style={{ width: '40%' }}>Windows 계정</th><th>표시 이름</th><th style={{ width: 80 }} /></tr></thead>
              <tbody>
                {userMap.map(([a, n], i) => (
                  <tr key={i} className={a === user.account ? 'sel' : undefined}>
                    <td style={{ maxWidth: 'none' }}>
                      <input className="input" aria-label="Windows 계정" value={a}
                        onChange={(e) => { touch(); setUserMap(userMap.map((r, j) => (j === i ? [e.target.value, r[1]] : r))); }} />
                    </td>
                    <td style={{ maxWidth: 'none' }}>
                      <input className="input" aria-label="표시 이름" value={n}
                        onChange={(e) => { touch(); setUserMap(userMap.map((r, j) => (j === i ? [r[0], e.target.value] : r))); }} />
                    </td>
                    <td><button className="btn btn-sm btn-danger" type="button" onClick={() => { touch(); setUserMap(userMap.filter((_, j) => j !== i)); }}>삭제</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {invalid.map && <span className="help help-err"><Icon name="error" />{invalid.map}</span>}
          <div className="btn-row">
            <button className="btn" type="button" onClick={() => { touch(); setUserMap([...userMap, ['', '']]); }}>줄 추가</button>
            <span className="help">이 PC 계정: {user.account} (선택된 줄)</span>
          </div>
        </div>
      )}

      <div className="toolbar" style={{ justifyContent: 'flex-end' }}>
        {hasInvalid && dirty && <span className="help help-err" style={{ marginRight: 'auto' }}><Icon name="error" />빨간 표시가 있는 탭의 입력값을 고쳐야 저장할 수 있습니다.</span>}
        <button className="btn" type="button" disabled={!dirty} onClick={cancel}>되돌리기</button>
        <button className="btn btn-primary" type="button" disabled={!dirty || hasInvalid} onClick={save}>설정 저장</button>
      </div>

      {restoreTarget && (
        <Dialog
          title="백업에서 복원"
          footer={<>
            <button className="btn" type="button" onClick={() => setRestoreTarget(null)}>닫기</button>
            <button className="btn btn-primary" type="button" onClick={doRestore}>복원</button>
          </>}
        >
          <Alert kind="warn">
            모든 데이터(검수 건·업체·검수자·설정·메일 기록)를 <b>{fmtDateTime(restoreTarget.createdAt)}</b> 시점으로 되돌립니다.
            그 뒤에 저장된 내용은 사라집니다.
          </Alert>
          <span className="help">복원 직전 상태도 백업해 두므로, 잘못 복원했으면 다시 되돌릴 수 있습니다. 다른 PC는 새로 불러와야 바뀐 내용이 보입니다.</span>
        </Dialog>
      )}
    </>
  );

  /** 숫자 규칙 한 줄 (컴포넌트로 만들면 입력할 때마다 칸이 새로 생겨 포커스를 잃으므로 함수로 그림) */
  function numRule({ key: k, label, unit }: (typeof NUM_RULES)[number]) {
    const r = rules[k];
    const err = numErr(k);
    return (
      <div className="rule" key={k}>
        <Switch label={label} checked={r.enabled} onChange={(v) => setRule(k, { ...r, enabled: v })} />
        <input className={`input${err ? ' invalid' : ''}`} style={{ width: 80 }} inputMode="numeric" aria-label={`${label} 값`}
          disabled={!r.enabled} value={numText[k]}
          onChange={(e) => { touch(); setNumText({ ...numText, [k]: e.target.value }); }} />
        <span className="muted" style={{ width: 20 }}>{unit}</span>
        {err && <span className="help help-err"><Icon name="error" />{err}</span>}
      </div>
    );
  }
}
