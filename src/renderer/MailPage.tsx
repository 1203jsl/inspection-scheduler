import { useState } from 'react';
import type { AppErrorInfo, Inspection, MailDraft, MailKind } from '../shared/types';
import { MAIL_KIND_LABEL } from '../shared/types';
import { fmtMD } from '../shared/dates';
import { fillTemplate, kindForStatus, mailCovers, mailGroup, mailVars, parseAddresses, pendingMail, recipientsFor } from '../shared/mail';
import { api } from './api';
import { Alert, ErrorAlert, Icon } from './components';
import { useApp, useEditLock, type Route } from './ctx';
import { StatusBadge } from './ListPage';

const KINDS = Object.keys(MAIL_KIND_LABEL) as MailKind[];

export function MailPage({ id, kind: initialKind, from }: { id: string; kind?: MailKind; from?: Route }) {
  const { data, user, go, inspectorName, reload } = useApp();
  const all = data.inspections.items;
  const log = data.mail_log.items;
  const insp = all.find((i) => i.id === id);
  const settings = data.settings;
  const vendor = data.vendors.items.find((v) => v.name === insp?.vendorName);
  const fallback = vendor ? { email: vendor.contactEmail, name: vendor.contactName } : undefined;

  /** 묶음 후보와 기본 선택: 아직 이 메일을 안 받은 건 (없으면 이 건만) */
  const groupOf = (k: MailKind) => {
    if (!insp) return { group: [] as Inspection[], sel: new Set<string>() };
    const group = mailGroup(insp, all, k);
    const unsent = group.filter((g) => !log.some((m) => mailCovers(m, g.id) && m.kind === k && m.result !== 'failed'));
    const sel = new Set((unsent.length ? unsent : [insp]).map((g) => g.id));
    sel.add(insp.id);
    return { group, sel };
  };
  const fill = (k: MailKind, items: Inspection[]) => {
    const vars = mailVars(items, inspectorName, user.displayName, k);
    const t = settings.mailTemplates[k];
    const s = fillTemplate(t.subject, vars), b = fillTemplate(t.body, vars);
    return { subject: s.text, body: b.text, unknown: [...new Set([...s.unknown, ...b.unknown])], to: recipientsFor(items, fallback) };
  };

  const startKind: MailKind = initialKind ?? (insp && (pendingMail(insp, log) ?? kindForStatus(insp))) ?? 'registered';
  const [init] = useState(() => {
    const g = groupOf(startKind);
    return { ...g, ...fill(startKind, g.group.filter((x) => g.sel.has(x.id))) };
  });
  const [kind, setKind] = useState<MailKind>(startKind);
  const [group, setGroup] = useState<Inspection[]>(init.group);
  const [sel, setSel] = useState<Set<string>>(init.sel);
  const [to, setTo] = useState(init.to);
  const [cc, setCc] = useState(settings.mail.defaultCc);
  const [subject, setSubject] = useState(init.subject);
  const [body, setBody] = useState(init.body);
  const [unknown, setUnknown] = useState(init.unknown);
  const [extra, setExtra] = useState<{ path: string; on: boolean }[]>([]);
  const [attachOn, setAttachOn] = useState<Set<string>>(new Set());
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<AppErrorInfo | null>(null);
  const [done, setDone] = useState('');
  const [copied, setCopied] = useState('');
  useEditLock(!done);

  if (!insp) {
    return <Alert kind="warn" action={<button className="btn btn-sm" type="button" onClick={() => go({ page: 'list' })}>목록으로</button>}>신청번호 {id}를 찾을 수 없습니다.</Alert>;
  }

  const items = group.filter((g) => sel.has(g.id));
  const apply = (k: MailKind, list: Inspection[]) => {
    const f = fill(k, list);
    setSubject(f.subject);
    setBody(f.body);
    setUnknown(f.unknown);
    setTo(f.to);
  };
  const changeKind = (k: MailKind) => {
    const g = groupOf(k);
    setKind(k);
    setGroup(g.group);
    setSel(g.sel);
    apply(k, g.group.filter((x) => g.sel.has(x.id)));
  };
  const toggleItem = (gid: string) => {
    const s = new Set(sel);
    if (s.has(gid)) s.delete(gid); else s.add(gid);
    if (s.size === 0) return; // 최소 한 건
    setSel(s);
    apply(kind, group.filter((g) => s.has(g.id)));
  };

  const manual = settings.mail.method === 'manual';
  const candidates = [...new Set(items.flatMap((i) => [...(i.attachments ?? []).map((a) => a.file), ...i.sources.map((s) => s.file)]))];
  const attachList = [...candidates.map((p) => ({ path: p, local: false })), ...extra.map((e) => ({ path: e.path, local: true }))];
  const draft = (): MailDraft => ({
    inspectionId: items[0].id,
    inspectionIds: items.map((i) => i.id),
    kind, to, cc, subject, body,
    attachments: attachList.filter((a) => attachOn.has(a.path)).map((a) => a.path),
  });
  const toCheck = parseAddresses(to), ccCheck = parseAddresses(cc);
  const addrErr = !to.trim() ? '받는 사람 이메일을 입력하세요. 업체 관리에서 업체 이메일을 등록할 수도 있습니다.'
    : toCheck.bad.length ? `이메일 형식이 올바르지 않습니다: ${toCheck.bad.join(', ')}` : '';
  const ccErr = ccCheck.bad.length ? `이메일 형식이 올바르지 않습니다: ${ccCheck.bad.join(', ')}` : '';
  const canSend = !addrErr && !ccErr && subject.trim() && !busy && !done;
  const label = items.length > 1 ? `${items.length}건` : items[0].id;

  const run = async (fn: () => ReturnType<typeof api.mailSend>, okMsg: string) => {
    setBusy(true);
    setError(null);
    const r = await fn();
    setBusy(false);
    if (r.ok) setDone(okMsg);
    else setError(r.error);
    await reload(); // 성공·실패 모두 이력에 보이게
  };
  const copy = async (what: string, text: string) => {
    const r = await api.copyText(text);
    if (r.ok) setCopied(what);
  };
  const back = () => go(from ?? { page: 'detail', id: insp.id });

  return (
    <>
      <Alert kind="info"><b>안내</b> · 업체에 보내는 메일은 반드시 미리보기를 확인한 뒤 발송합니다. 본문은 직접 수정할 수 있습니다.</Alert>
      {done && <Alert kind="ok" action={<button className="btn btn-sm" type="button" onClick={back}>돌아가기</button>}>{done}</Alert>}
      {error && <ErrorAlert error={error} />}

      <div className="split" style={{ gridTemplateColumns: 'minmax(0, 1fr) 400px' }}>
        <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          <div className="card-h" style={{ margin: 0 }}>
            {MAIL_KIND_LABEL[kind]} 메일 · {insp.vendorName} {items.length > 1 ? `${items.length}건` : insp.id} {items.length === 1 && <StatusBadge i={insp} />}
          </div>
          <div className="field">
            <label htmlFor="mk">메일 종류</label>
            <select id="mk" className="sel" style={{ maxWidth: 260 }} value={kind} disabled={!!done} onChange={(e) => changeKind(e.target.value as MailKind)}>
              {KINDS.map((k) => <option key={k} value={k}>{MAIL_KIND_LABEL[k]}{k === kindForStatus(insp) ? ' (현재 상태)' : ''}</option>)}
            </select>
            <span className="help">종류나 함께 보낼 건을 바꾸면 받는 사람·제목·본문이 다시 채워집니다.</span>
          </div>

          {group.length > 1 && (
            <fieldset className="field" style={{ border: '1px solid var(--border)', borderRadius: 'var(--radius-md)', padding: 12, margin: 0 }}>
              <legend className="lbl" style={{ padding: '0 4px' }}>함께 보낼 건 · 같은 업체에 같은 날 확정한 {group.length}건</legend>
              {group.map((g) => {
                const sent = log.some((m) => mailCovers(m, g.id) && m.kind === kind && m.result !== 'failed');
                return (
                  <label key={g.id} className="radio" style={{ fontSize: 13 }}>
                    <input type="checkbox" className="chk" checked={sel.has(g.id)} disabled={!!done || (sel.size === 1 && sel.has(g.id))} onChange={() => toggleItem(g.id)} />
                    <b>{g.id}</b> {g.itemName} · {fmtMD(g.schedule.date)} {g.schedule.time ?? ''}
                    {sent && <span className="cell-flag" style={{ color: 'var(--ok-fg)' }}><Icon name="ok" />이미 보냄</span>}
                  </label>
                );
              })}
            </fieldset>
          )}

          <div className="field">
            <label htmlFor="mto">받는 사람 <span className="req">*</span></label>
            <input id="mto" className={`input${addrErr ? ' invalid' : ''}`} value={to} onChange={(e) => setTo(e.target.value)} disabled={!!done} />
            {addrErr ? <span className="help help-err"><Icon name="error" />{addrErr}</span> : <span className="help">여러 명은 쉼표(,)로 구분합니다.</span>}
          </div>
          <div className="field">
            <label htmlFor="mcc">참조</label>
            <input id="mcc" className={`input${ccErr ? ' invalid' : ''}`} value={cc} onChange={(e) => setCc(e.target.value)} disabled={!!done} />
            {ccErr && <span className="help help-err"><Icon name="error" />{ccErr}</span>}
          </div>
          <div className="field">
            <label htmlFor="msub">제목 <span className="req">*</span></label>
            <input id="msub" className="input" spellCheck={false} value={subject} onChange={(e) => setSubject(e.target.value)} disabled={!!done} />
          </div>
          <div className="field">
            <label htmlFor="mbody">본문 (템플릿 값이 채워진 상태)</label>
            <textarea id="mbody" className="input" spellCheck={false} style={{ height: 250, lineHeight: '22px' }} value={body} onChange={(e) => setBody(e.target.value)} disabled={!!done} />
          </div>
          {unknown.length > 0 && (
            <Alert kind="warn">채워지지 않은 치환값이 있습니다: {unknown.join(', ')} — 본문에서 직접 고치거나, 설정 &gt; 메일 템플릿의 이름을 확인하세요.</Alert>
          )}
          <div className="field">
            <span className="lbl">첨부파일</span>
            {attachList.length === 0 && <span className="help">첨부 후보가 없습니다.</span>}
            {attachList.map((a) => (
              <label key={a.path} className="radio">
                <input type="checkbox" className="chk" checked={attachOn.has(a.path)} disabled={!!done}
                  onChange={() => { const s = new Set(attachOn); s.has(a.path) ? s.delete(a.path) : s.add(a.path); setAttachOn(s); }} />
                <Icon name="file" />{a.path.split(/[\\/]/).pop()}
                {a.local && <span className="muted">(이 PC의 파일)</span>}
              </label>
            ))}
            <button className="btn btn-sm" type="button" style={{ alignSelf: 'flex-start' }} disabled={!!done} onClick={async () => {
              const r = await api.mailPickAttachments();
              if (!r.ok) return;
              const fresh = r.value.filter((p) => !attachList.some((a) => a.path === p));
              setExtra([...extra, ...fresh.map((p) => ({ path: p, on: true }))]);
              setAttachOn(new Set([...attachOn, ...fresh]));
            }}>파일 추가</button>
            {manual && attachOn.size > 0 && <span className="help">수동 발송에서는 첨부파일을 웹메일에 직접 올려야 합니다.</span>}
          </div>
          <div className="btn-row" style={{ justifyContent: 'flex-end' }}>
            <button className="btn" type="button" onClick={back}>{done ? '닫기' : '취소'}</button>
            {!manual && (
              <button className="btn btn-primary" type="button" disabled={!canSend}
                onClick={() => run(() => api.mailSend(draft()), `${label}: ${MAIL_KIND_LABEL[kind]} 메일을 보냈습니다. 메일 이력에 기록했습니다.`)}>
                <Icon name="mail" />{busy ? '보내는 중…' : '발송'}
              </button>
            )}
          </div>
        </div>

        {manual ? (
          <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 12, alignSelf: 'start' }}>
            <div className="card-h" style={{ margin: 0 }}>수동 발송</div>
            <div className="help">발송 방식이 '수동'입니다. 아래 순서로 웹메일에서 보낸 뒤 기록하세요.</div>
            <div className="field">
              <span className="lbl">1. 내용 복사</span>
              <div className="btn-row">
                <button className="btn btn-sm" type="button" onClick={() => copy('받는 사람', [to, cc && `참조: ${cc}`].filter(Boolean).join('\n'))}>받는 사람 복사</button>
                <button className="btn btn-sm" type="button" onClick={() => copy('제목', subject)}>제목 복사</button>
                <button className="btn btn-sm" type="button" onClick={() => copy('본문', body)}>본문 복사</button>
              </div>
              {copied && <span className="help" role="status">{copied}을(를) 복사했습니다. 웹메일에 붙여넣으세요.</span>}
            </div>
            <div className="field">
              <span className="lbl">2. 웹메일에 붙여넣어 발송</span>
              <button className="btn" type="button" style={{ alignSelf: 'flex-start' }} disabled={!settings.mail.webmailUrl}
                onClick={async () => { const r = await api.mailOpenWebmail(settings.mail.webmailUrl); if (!r.ok) setError(r.error); }}>웹메일 열기</button>
              {!settings.mail.webmailUrl && <span className="help">설정 &gt; 메일 발송 설정에 웹메일 주소를 넣으면 바로 열 수 있습니다.</span>}
            </div>
            <div className="field">
              <span className="lbl">3. 발송을 마쳤다면</span>
              <button className="btn btn-primary" type="button" style={{ alignSelf: 'flex-start' }} disabled={!canSend}
                onClick={() => run(() => api.mailRecordManual(draft()), `${label}: ${MAIL_KIND_LABEL[kind]} 메일을 수동 발송으로 기록했습니다.`)}>
                <Icon name="ok" />발송 완료로 기록
              </button>
            </div>
            <Alert kind="warn">[발송 완료로 기록]은 실제로 메일을 보낸 뒤에만 누르세요. 이력에 '발송'으로 남습니다.</Alert>
          </div>
        ) : (
          <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 8, alignSelf: 'start' }}>
            <div className="card-h" style={{ margin: 0 }}>자동 발송 (SMTP)</div>
            <span className="help">서버: {settings.mail.smtp.host || '(입력 안 됨)'}:{settings.mail.smtp.port}</span>
            <span className="help">[발송]을 누르면 이 PC에 저장한 내 메일 계정으로 보냅니다. 성공·실패는 메일 이력에 남습니다.</span>
            <button className="btn btn-sm" type="button" style={{ alignSelf: 'flex-start' }} onClick={() => go({ page: 'settings' })}>메일 발송 설정 열기</button>
          </div>
        )}
      </div>
    </>
  );
}
