import { useMemo, useState } from 'react';
import type { AppErrorInfo, Inspection, MailDraft, MailKind } from '../shared/types';
import { MAIL_KIND_LABEL } from '../shared/types';
import { fillTemplate, kindForStatus, mailVars, parseAddresses, pendingMail } from '../shared/mail';
import { api } from './api';
import { Alert, ErrorAlert, Icon } from './components';
import { useApp, useEditLock, type Route } from './ctx';
import { StatusBadge } from './ListPage';

const KINDS = Object.keys(MAIL_KIND_LABEL) as MailKind[];

/** 받는 사람 기본값: 엑셀에 적힌 업체 담당자, 없으면 업체 관리의 이메일 */
function defaultTo(i: Inspection, vendorEmail: string, vendorName: string): string {
  const email = i.contactEmail || vendorEmail;
  if (!email) return '';
  const name = i.contactEmail ? i.contactName : vendorName;
  return name ? `${name} <${email}>` : email;
}

export function MailPage({ id, kind: initialKind, from }: { id: string; kind?: MailKind; from?: Route }) {
  const { data, user, go, inspectorName, reload } = useApp();
  const insp = data.inspections.items.find((i) => i.id === id);
  const settings = data.settings;
  const vendor = data.vendors.items.find((v) => v.name === insp?.vendorName);

  const startKind: MailKind = initialKind ?? (insp && (pendingMail(insp, data.mail_log.items) ?? kindForStatus(insp))) ?? 'registered';
  const [kind, setKind] = useState<MailKind>(startKind);
  const build = (k: MailKind) => {
    if (!insp) return { subject: '', body: '', unknown: [] as string[] };
    const vars = mailVars(insp, inspectorName, user.displayName);
    const t = settings.mailTemplates[k];
    const s = fillTemplate(t.subject, vars), b = fillTemplate(t.body, vars);
    return { subject: s.text, body: b.text, unknown: [...new Set([...s.unknown, ...b.unknown])] };
  };
  const first = useMemo(() => build(startKind), []); // eslint-disable-line react-hooks/exhaustive-deps
  const [to, setTo] = useState(insp ? defaultTo(insp, vendor?.contactEmail ?? '', vendor?.contactName ?? '') : '');
  const [cc, setCc] = useState(settings.mail.defaultCc);
  const [subject, setSubject] = useState(first.subject);
  const [body, setBody] = useState(first.body);
  const [unknown, setUnknown] = useState(first.unknown);
  const candidates = (insp?.sources ?? []).map((s) => s.file);
  const [attach, setAttach] = useState<{ path: string; on: boolean }[]>(candidates.map((p) => ({ path: p, on: false })));
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<AppErrorInfo | null>(null);
  const [done, setDone] = useState('');
  const [copied, setCopied] = useState('');
  useEditLock(!done);

  if (!insp) {
    return <Alert kind="warn" action={<button className="btn btn-sm" type="button" onClick={() => go({ page: 'list' })}>목록으로</button>}>신청번호 {id}를 찾을 수 없습니다.</Alert>;
  }

  const manual = settings.mail.method === 'manual';
  const changeKind = (k: MailKind) => {
    setKind(k);
    const b = build(k);
    setSubject(b.subject);
    setBody(b.body);
    setUnknown(b.unknown);
  };
  const draft = (): MailDraft => ({
    inspectionId: insp.id, kind, to, cc, subject, body, attachments: attach.filter((a) => a.on).map((a) => a.path),
  });
  const toCheck = parseAddresses(to), ccCheck = parseAddresses(cc);
  const addrErr = !to.trim() ? '받는 사람 이메일을 입력하세요. 업체 관리에서 업체 이메일을 등록할 수도 있습니다.'
    : toCheck.bad.length ? `이메일 형식이 올바르지 않습니다: ${toCheck.bad.join(', ')}` : '';
  const ccErr = ccCheck.bad.length ? `이메일 형식이 올바르지 않습니다: ${ccCheck.bad.join(', ')}` : '';
  const canSend = !addrErr && !ccErr && subject.trim() && !busy && !done;

  const run = async (fn: () => ReturnType<typeof api.mailSend>, okMsg: string) => {
    setBusy(true);
    setError(null);
    const r = await fn();
    setBusy(false);
    if (r.ok) {
      setDone(okMsg);
      await reload();
    } else {
      setError(r.error);
      await reload(); // 실패 기록도 이력에 보이게
    }
  };
  const copy = async (label: string, text: string) => {
    const r = await api.copyText(text);
    if (r.ok) setCopied(label);
  };
  const back = () => go(from ?? { page: 'detail', id: insp.id });

  return (
    <>
      <Alert kind="info"><b>안내</b> · 업체에 보내는 메일은 반드시 미리보기를 확인한 뒤 발송합니다. 본문은 직접 수정할 수 있습니다.</Alert>
      {done && (
        <Alert kind="ok" action={<button className="btn btn-sm" type="button" onClick={back}>상세 화면으로</button>}>{done}</Alert>
      )}
      {error && <ErrorAlert error={error} />}

      <div className="split" style={{ gridTemplateColumns: 'minmax(0, 1fr) 400px' }}>
        <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          <div className="card-h" style={{ margin: 0 }}>
            {MAIL_KIND_LABEL[kind]} 메일 · {insp.id} {insp.vendorName} <StatusBadge i={insp} />
          </div>
          <div className="field">
            <label htmlFor="mk">메일 종류</label>
            <select id="mk" className="sel" style={{ maxWidth: 260 }} value={kind} disabled={!!done} onChange={(e) => changeKind(e.target.value as MailKind)}>
              {KINDS.map((k) => <option key={k} value={k}>{MAIL_KIND_LABEL[k]}{k === kindForStatus(insp) ? ' (현재 상태)' : ''}</option>)}
            </select>
            <span className="help">종류를 바꾸면 제목·본문이 그 템플릿으로 다시 채워집니다.</span>
          </div>
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
            {attach.length === 0 && <span className="help">첨부 후보가 없습니다.</span>}
            {attach.map((a, k) => (
              <label key={a.path} className="radio">
                <input type="checkbox" className="chk" checked={a.on} disabled={!!done}
                  onChange={(e) => setAttach(attach.map((x, j) => (j === k ? { ...x, on: e.target.checked } : x)))} />
                <Icon name="file" />{a.path.split(/[\\/]/).pop()}
                {!candidates.includes(a.path) && <span className="muted">(이 PC의 파일)</span>}
              </label>
            ))}
            <button className="btn btn-sm" type="button" style={{ alignSelf: 'flex-start' }} disabled={!!done} onClick={async () => {
              const r = await api.mailPickAttachments();
              if (r.ok) setAttach([...attach, ...r.value.filter((p) => !attach.some((a) => a.path === p)).map((p) => ({ path: p, on: true }))]);
            }}>파일 추가</button>
            {manual && attach.some((a) => a.on) && <span className="help">수동 발송에서는 첨부파일을 웹메일에 직접 올려야 합니다.</span>}
          </div>
          <div className="btn-row" style={{ justifyContent: 'flex-end' }}>
            <button className="btn" type="button" onClick={back}>{done ? '닫기' : '취소'}</button>
            {!manual && (
              <button className="btn btn-primary" type="button" disabled={!canSend}
                onClick={() => run(() => api.mailSend(draft()), `${insp.id}: ${MAIL_KIND_LABEL[kind]} 메일을 보냈습니다. 메일 이력에 기록했습니다.`)}>
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
                onClick={() => run(() => api.mailRecordManual(draft()), `${insp.id}: ${MAIL_KIND_LABEL[kind]} 메일을 수동 발송으로 기록했습니다.`)}>
                <Icon name="ok" />발송 완료로 기록
              </button>
            </div>
            <Alert kind="warn">[발송 완료로 기록]은 실제로 메일을 보낸 뒤에만 누르세요. 이력에 '발송'으로 남습니다.</Alert>
          </div>
        ) : (
          <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 8, alignSelf: 'start' }}>
            <div className="card-h" style={{ margin: 0 }}>자동 발송 (SMTP)</div>
            <span className="help">서버: {settings.mail.smtp.host || '(입력 안 됨)'}:{settings.mail.smtp.port}</span>
            <span className="help">[발송]을 누르면 이 PC에 저장한 메일 계정으로 보냅니다. 성공·실패는 메일 이력에 남습니다.</span>
            <button className="btn btn-sm" type="button" style={{ alignSelf: 'flex-start' }} onClick={() => go({ page: 'settings' })}>메일 발송 설정 열기</button>
          </div>
        )}
      </div>
    </>
  );
}
