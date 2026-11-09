import { useState } from 'react';
import type { AppErrorInfo, Inspection, Inspector, RequestFields, Schedule, Vendor } from '../shared/types';
import { normalizeTime } from '../shared/dates';
import { conflictsFor, toSlot } from '../shared/conflicts';
import { checkForce, checkSchedule, WorkflowError, type ForceInfo } from '../shared/workflow';
import { Alert, Dialog, ErrorAlert, Icon } from './components';
import { useApp, useEditLock } from './ctx';

function FieldErr({ msg }: { msg?: string }) {
  return msg ? <span className="help help-err"><Icon name="error" />{msg}</span> : null;
}

/** 공통: 팝업 안에서 저장하고 결과를 보여줌 */
function useSubmit(onClose: () => void) {
  const [error, setError] = useState<AppErrorInfo | null>(null);
  const [busy, setBusy] = useState(false);
  const run = async (fn: () => Promise<AppErrorInfo | 'handled' | null>) => {
    setBusy(true);
    setError(null);
    try {
      const r = await fn();
      if (r === null || r === 'handled') onClose();
      else setError(r);
    } catch (e) {
      if (e instanceof WorkflowError) setError({ code: 'UNKNOWN', message: e.message });
      else throw e;
    } finally {
      setBusy(false);
    }
  };
  return { error, busy, run };
}

export function locationOptions(items: Inspection[]): string[] {
  return [...new Set(items.flatMap((i) => [i.schedule.location, i.location]).filter(Boolean))].sort();
}

// ───────── 확정 / 변경 ─────────

export function ScheduleDialog({ mode, insp, onClose, onSubmit }: {
  mode: 'confirm' | 'change';
  insp: Inspection;
  onClose: () => void;
  onSubmit: (s: Schedule, reason: string, force?: ForceInfo) => Promise<AppErrorInfo | 'handled' | null>;
}) {
  useEditLock(true);
  const { data, inspectorName } = useApp();
  const inspectors = data.inspectors.items;
  const [s, setS] = useState<Schedule>({ ...insp.schedule });
  const [reason, setReason] = useState('');
  const [tried, setTried] = useState(false);
  const { error, busy, run } = useSubmit(onClose);
  const errs = checkSchedule(s, inspectors);
  const show = (k: keyof Schedule) => (tried ? errs[k] : undefined);
  const active = inspectors.filter((p) => p.active || p.id === s.inspectorId);
  const reconfirm = mode === 'confirm' && insp.status === 'changed';
  const title = mode === 'change' ? '일정 변경' : reconfirm ? '일정 재확정' : '일정 확정';

  // 충돌 조건 (입력할 때마다 다시 검사)
  const rules = data.settings.conflictRules;
  const [forceOn, setForceOn] = useState(false);
  const [forceReason, setForceReason] = useState('');
  const fixedTime = s.time ? normalizeTime(s.time) ?? s.time : s.time;
  const conflicts = rules
    ? conflictsFor(toSlot({ ...insp, schedule: { ...s, time: fixedTime } }), data.inspections.items.map(toSlot), rules, (id) => inspectorName(id)).map((c) => c.message)
    : [];
  const blocked = conflicts.length > 0 && !rules.allowForce;
  const needReason = conflicts.length > 0 && forceOn && tried && !forceReason.trim();

  const submit = () => {
    setTried(true);
    const fixed = { ...s, time: fixedTime };
    if (Object.keys(checkSchedule(fixed, inspectors)).length) { setS(fixed); return; }
    if (conflicts.length && (blocked || !forceOn || !forceReason.trim())) return;
    run(async () => onSubmit(fixed, reason, checkForce(insp.id, conflicts, rules.allowForce, forceReason)));
  };
  const verb = mode === 'change' ? '변경' : reconfirm ? '재확정' : '확정';

  return (
    <Dialog
      title={title}
      width={560}
      footer={<>
        <button className="btn" type="button" onClick={onClose}>닫기</button>
        <button className="btn btn-primary" type="button" onClick={submit} disabled={busy || blocked || (conflicts.length > 0 && !forceOn)}>
          {conflicts.length && forceOn ? `강제 진행하여 ${verb}` : verb}
        </button>
      </>}
    >
      <p className="help" style={{ margin: 0 }}>{insp.id} · {insp.vendorName} · {insp.itemName} · 수량 {insp.quantity ?? '-'}</p>
      {mode === 'change' && <Alert kind="info">변경하면 상태가 '변경됨'이 되고, 다시 [확정]해야 합니다.</Alert>}
      <div className="form-row">
        <div className="field">
          <label htmlFor="sd">검수일 <span className="req">*</span></label>
          <input id="sd" type="date" className={`input${show('date') ? ' invalid' : ''}`} value={s.date ?? ''} onChange={(e) => setS({ ...s, date: e.target.value || null })} />
          <FieldErr msg={show('date')} />
        </div>
        <div className="field">
          <label htmlFor="st">검수 시간 <span className="req">*</span></label>
          <input id="st" className={`input${show('time') ? ' invalid' : ''}`} inputMode="numeric" maxLength={8} placeholder="예: 13:30"
            value={s.time ?? ''} onChange={(e) => setS({ ...s, time: e.target.value || null })}
            onBlur={() => { const n = s.time ? normalizeTime(s.time) : null; if (n) setS({ ...s, time: n }); }} />
          {!show('time') && <span className="help">24시간 형식 (예: 09:30, 13:30)</span>}
          <FieldErr msg={show('time')} />
        </div>
        <div className="field">
          <label htmlFor="sl">검수 장소 <span className="req">*</span></label>
          <input id="sl" className={`input${show('location') ? ' invalid' : ''}`} list="loc-list" value={s.location} onChange={(e) => setS({ ...s, location: e.target.value })} />
          <datalist id="loc-list">{locationOptions(data.inspections.items).map((l) => <option key={l} value={l} />)}</datalist>
          <FieldErr msg={show('location')} />
        </div>
        <div className="field">
          <label htmlFor="si">검수자 <span className="req">*</span></label>
          <select id="si" className={`sel${show('inspectorId') ? ' invalid' : ''}`} value={s.inspectorId ?? ''} onChange={(e) => setS({ ...s, inspectorId: e.target.value || null })}>
            <option value="">선택하세요</option>
            {active.map((p) => <option key={p.id} value={p.id}>{p.name}{p.dept ? ` (${p.dept})` : ''}{p.active ? '' : ' · 사용 안 함'}</option>)}
          </select>
          <FieldErr msg={show('inspectorId')} />
          {inspectors.length === 0 && <span className="help">검수자 관리 화면에서 먼저 검수자를 추가하세요.</span>}
        </div>
      </div>
      {conflicts.length > 0 && (
        <div className={`alert ${blocked ? 'alert-danger' : 'alert-warn'}`} role="status" style={{ flexDirection: 'column', gap: 4 }}>
          <span style={{ display: 'flex', gap: 8, alignItems: 'center' }}><Icon name="warn" /><b>충돌 조건 {conflicts.length}건에 걸립니다</b></span>
          {conflicts.map((c) => <span key={c}>· {c}</span>)}
          <span>{blocked
            ? '설정에서 강제 진행이 꺼져 있어 진행할 수 없습니다. 다른 날짜나 검수자를 고르세요.'
            : '다른 날짜나 검수자를 고르거나, 꼭 필요하면 사유를 적고 강제로 진행하세요.'}</span>
        </div>
      )}
      {conflicts.length > 0 && !blocked && (
        <>
          <label className="switch-row" style={{ fontSize: 14 }}>
            <input type="checkbox" className="chk" checked={forceOn} onChange={(e) => setForceOn(e.target.checked)} />
            충돌을 확인했으며 강제로 진행합니다
          </label>
          {forceOn && (
            <div className="field">
              <label htmlFor="fr">강제 진행 사유 <span className="req">*</span></label>
              <textarea id="fr" className={`input${needReason ? ' invalid' : ''}`} style={{ minHeight: 64 }} value={forceReason}
                onChange={(e) => setForceReason(e.target.value)} placeholder="예: 업체 장비 반출 일정상 이 날 외 불가. 팀장 승인" />
              {needReason ? <FieldErr msg="강제 진행 사유를 입력하세요." /> : <span className="help">걸린 충돌 조건과 함께 변경 이력에 기록됩니다.</span>}
            </div>
          )}
        </>
      )}
      {mode === 'change' && (
        <div className="field">
          <label htmlFor="sr">변경 사유</label>
          <textarea id="sr" className="input" style={{ minHeight: 64 }} value={reason} onChange={(e) => setReason(e.target.value)} placeholder="예: 업체 입고 지연으로 일정 연기" />
          <span className="help">변경 이력에 남습니다.</span>
        </div>
      )}
      {error && <ErrorAlert error={error} />}
    </Dialog>
  );
}

// ───────── 취소 ─────────

export function CancelDialog({ label, onClose, onSubmit }: {
  label: string;
  onClose: () => void;
  onSubmit: (reason: string) => Promise<AppErrorInfo | 'handled' | null>;
}) {
  useEditLock(true);
  const [reason, setReason] = useState('');
  const [tried, setTried] = useState(false);
  const { error, busy, run } = useSubmit(onClose);
  const empty = !reason.trim();
  return (
    <Dialog
      title="검수 취소"
      footer={<>
        <button className="btn" type="button" onClick={onClose}>닫기</button>
        <button className="btn btn-danger" type="button" disabled={busy} onClick={() => { setTried(true); if (!empty) run(() => onSubmit(reason)); }}>취소 처리</button>
      </>}
    >
      <Alert kind="warn">{label}을(를) 취소합니다. 취소한 건은 다시 확정할 수 없습니다.</Alert>
      <div className="field">
        <label htmlFor="cr">취소 사유 <span className="req">*</span></label>
        <textarea id="cr" className={`input${tried && empty ? ' invalid' : ''}`} style={{ minHeight: 72 }} value={reason} onChange={(e) => setReason(e.target.value)} placeholder="예: 업체 요청으로 납품 취소" />
        {tried && empty ? <FieldErr msg="취소 사유를 입력하세요." /> : <span className="help">변경 이력에 남습니다.</span>}
      </div>
      {error && <ErrorAlert error={error} />}
    </Dialog>
  );
}

// ───────── 기본정보 수정 ─────────

const INFO_FIELDS: { key: keyof RequestFields; label: string; type?: string }[] = [
  { key: 'vendorName', label: '업체명' },
  { key: 'category', label: '검수 품목' },
  { key: 'itemName', label: '품명' },
  { key: 'spec', label: '품번/규격' },
  { key: 'quantity', label: '수량', type: 'number' },
  { key: 'requestedDate', label: '검수 희망일', type: 'date' },
  { key: 'requestedTime', label: '희망 시간', type: 'hhmm' },
  { key: 'location', label: '검수 장소(신청)' },
  { key: 'contactName', label: '업체 담당자' },
  { key: 'contactPhone', label: '연락처' },
  { key: 'contactEmail', label: '이메일' },
  { key: 'note', label: '비고' },
];

export function EditInfoDialog({ insp, onClose, onSubmit }: {
  insp: Inspection;
  onClose: () => void;
  onSubmit: (patch: Partial<RequestFields>, reason: string) => Promise<AppErrorInfo | 'handled' | null>;
}) {
  useEditLock(true);
  const { data } = useApp();
  const [v, setV] = useState<Record<string, string>>(
    Object.fromEntries(INFO_FIELDS.map((f) => [f.key, insp[f.key] === null || insp[f.key] === undefined ? '' : String(insp[f.key])])),
  );
  const [reason, setReason] = useState('');
  const { error, busy, run } = useSubmit(onClose);
  const qty = v.quantity.trim();
  const errs: Record<string, string> = {};
  if (qty && !(Number.isInteger(Number(qty)) && Number(qty) >= 1)) errs.quantity = '1 이상의 정수로 입력하세요.';
  if (v.contactEmail.trim() && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(v.contactEmail.trim())) errs.contactEmail = '이메일 형식이 올바르지 않습니다. 예: name@company.co.kr';
  if (!v.vendorName.trim()) errs.vendorName = '업체명을 입력하세요.';
  if (!v.itemName.trim()) errs.itemName = '품명을 입력하세요.';
  if (v.requestedTime.trim() && !normalizeTime(v.requestedTime)) errs.requestedTime = '24시간 형식으로 입력하세요. 예: 13:30';

  const submit = () => {
    if (Object.keys(errs).length) return;
    const patch: Partial<RequestFields> = {};
    for (const f of INFO_FIELDS) {
      const raw = v[f.key].trim();
      (patch as any)[f.key] = f.key === 'quantity' ? (raw ? Number(raw) : null)
        : f.type === 'hhmm' ? (raw ? normalizeTime(raw) : null)
        : f.type === 'date' ? (raw || null) : raw;
    }
    run(() => onSubmit(patch, reason));
  };

  return (
    <Dialog
      title="기본정보 수정"
      width={720}
      footer={<>
        <button className="btn" type="button" onClick={onClose}>닫기</button>
        <button className="btn btn-primary" type="button" disabled={busy || Object.keys(errs).length > 0} onClick={submit}>저장</button>
      </>}
    >
      <span className="help">업체가 신청한 내용입니다. 실제 검수 일정은 [확정]/[변경]에서 정합니다. 처음 신청한 값(원본)은 그대로 남습니다.</span>
      <div className="form-row" style={{ gridTemplateColumns: 'repeat(3, minmax(0, 1fr))' }}>
        {INFO_FIELDS.map((f) => (
          <div className="field" key={f.key}>
            <label htmlFor={`f-${f.key}`}>{f.label}</label>
            {f.key === 'category' ? (
              <select id="f-category" className="sel" value={v.category} onChange={(e) => setV({ ...v, category: e.target.value })}>
                {[...new Set([...(data.settings.inspectionItems ?? []), v.category].filter(Boolean))].map((c) => <option key={c}>{c}</option>)}
                {!v.category && <option value="">선택하세요</option>}
              </select>
            ) : (
              <input id={`f-${f.key}`} className={`input${errs[f.key] ? ' invalid' : ''}`} type={f.type === 'date' ? 'date' : 'text'}
                placeholder={f.type === 'hhmm' ? '예: 13:30' : undefined}
                inputMode={f.type === 'number' ? 'numeric' : undefined}
                value={v[f.key]} onChange={(e) => setV({ ...v, [f.key]: e.target.value })} />
            )}
            <FieldErr msg={errs[f.key]} />
          </div>
        ))}
      </div>
      <div className="field">
        <label htmlFor="er">수정 사유</label>
        <input id="er" className="input" value={reason} onChange={(e) => setReason(e.target.value)} placeholder="예: 업체 전화로 수량 정정" />
      </div>
      {error && <ErrorAlert error={error} />}
    </Dialog>
  );
}

// ───────── 업체 ─────────

export function VendorDialog({ vendor, locked, onClose, onSubmit }: {
  vendor: Vendor | null;
  /** 검수 건이 있어 업체명을 바꿀 수 없음 */
  locked: boolean;
  onClose: () => void;
  onSubmit: (v: Vendor) => Promise<AppErrorInfo | 'handled' | null>;
}) {
  useEditLock(true);
  const { data } = useApp();
  const [v, setV] = useState<Vendor>(vendor ?? { name: '', contactName: '', contactPhone: '', contactEmail: '', updatedAt: '' });
  const [tried, setTried] = useState(false);
  const { error, busy, run } = useSubmit(onClose);
  const errs: Record<string, string> = {};
  if (!v.name.trim()) errs.name = '업체명을 입력하세요.';
  else if (data.vendors.items.some((x) => x.name === v.name.trim() && x !== vendor)) errs.name = '이미 등록된 업체명입니다.';
  if (v.contactEmail.trim() && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(v.contactEmail.trim())) errs.contactEmail = '이메일 형식이 올바르지 않습니다. 메일 받는 사람 기본값으로 쓰이니 확인하세요.';
  const show = (k: string) => (tried || k === 'contactEmail' ? errs[k] : undefined);

  return (
    <Dialog
      title={vendor ? '업체 정보 수정' : '업체 추가'}
      footer={<>
        <button className="btn" type="button" onClick={onClose}>닫기</button>
        <button className="btn btn-primary" type="button" disabled={busy} onClick={() => {
          setTried(true);
          if (Object.keys(errs).length) return;
          run(() => onSubmit({ ...v, name: v.name.trim(), contactEmail: v.contactEmail.trim() }));
        }}>저장</button>
      </>}
    >
      <div className="field">
        <label htmlFor="vn">업체명 <span className="req">*</span></label>
        <input id="vn" className={`input${show('name') ? ' invalid' : ''}`} value={v.name} disabled={locked} onChange={(e) => setV({ ...v, name: e.target.value })} />
        {locked ? <span className="help">검수 건이 있는 업체는 이름을 바꿀 수 없습니다.</span> : <FieldErr msg={show('name')} />}
      </div>
      <div className="form-row">
        <div className="field"><label htmlFor="vc">담당자</label><input id="vc" className="input" value={v.contactName} onChange={(e) => setV({ ...v, contactName: e.target.value })} /></div>
        <div className="field"><label htmlFor="vp">연락처</label><input id="vp" className="input" value={v.contactPhone} onChange={(e) => setV({ ...v, contactPhone: e.target.value })} /></div>
      </div>
      <div className="field">
        <label htmlFor="ve">이메일</label>
        <input id="ve" className={`input${show('contactEmail') ? ' invalid' : ''}`} value={v.contactEmail} onChange={(e) => setV({ ...v, contactEmail: e.target.value })} />
        <FieldErr msg={show('contactEmail')} />
      </div>
      {error && <ErrorAlert error={error} />}
    </Dialog>
  );
}

// ───────── 검수자 ─────────

export function InspectorDialog({ inspector, onClose, onSubmit }: {
  inspector: Inspector | null;
  onClose: () => void;
  onSubmit: (p: Inspector) => Promise<AppErrorInfo | 'handled' | null>;
}) {
  useEditLock(true);
  const { data } = useApp();
  const [p, setP] = useState<Inspector>(inspector ?? { id: '', name: '', dept: '', email: '', active: true, createdAt: '' });
  const [tried, setTried] = useState(false);
  const { error, busy, run } = useSubmit(onClose);
  const errs: Record<string, string> = {};
  if (!p.name.trim()) errs.name = '이름을 입력하세요.';
  else if (data.inspectors.items.some((x) => x.name === p.name.trim() && x.id !== p.id)) errs.name = '같은 이름의 검수자가 있습니다. 구분할 수 있게 적어 주세요 (예: 김민준B).';
  if (p.email.trim() && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(p.email.trim())) errs.email = '이메일 형식이 올바르지 않습니다. 알림 메일을 받을 주소입니다.';
  const show = (k: string) => (tried || k === 'email' ? errs[k] : undefined);

  return (
    <Dialog
      title={inspector ? '검수자 수정' : '검수자 추가'}
      footer={<>
        <button className="btn" type="button" onClick={onClose}>닫기</button>
        <button className="btn btn-primary" type="button" disabled={busy} onClick={() => {
          setTried(true);
          if (Object.keys(errs).length) return;
          run(() => onSubmit({ ...p, name: p.name.trim(), dept: p.dept.trim(), email: p.email.trim() }));
        }}>저장</button>
      </>}
    >
      <div className="form-row">
        <div className="field">
          <label htmlFor="pn">이름 <span className="req">*</span></label>
          <input id="pn" className={`input${show('name') ? ' invalid' : ''}`} value={p.name} onChange={(e) => setP({ ...p, name: e.target.value })} />
          <FieldErr msg={show('name')} />
        </div>
        <div className="field"><label htmlFor="pd">부서</label><input id="pd" className="input" value={p.dept} onChange={(e) => setP({ ...p, dept: e.target.value })} /></div>
      </div>
      <div className="field">
        <label htmlFor="pe">이메일</label>
        <input id="pe" className={`input${show('email') ? ' invalid' : ''}`} value={p.email} onChange={(e) => setP({ ...p, email: e.target.value })} />
        <FieldErr msg={show('email')} />
      </div>
      <label className="switch-row">
        <input type="checkbox" className="switch" role="switch" checked={p.active} onChange={(e) => setP({ ...p, active: e.target.checked })} />
        <span>{p.active ? '사용' : '사용 안 함'}</span>
      </label>
      <span className="help">'사용 안 함'인 검수자는 배정 목록에서 빠집니다. 이미 배정된 건은 그대로 남습니다.</span>
      {error && <ErrorAlert error={error} />}
    </Dialog>
  );
}
