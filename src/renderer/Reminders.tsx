import { useState } from 'react';
import type { Inspection, MailDraft } from '../shared/types';
import { fmtMD, parseYmd } from '../shared/dates';
import { api, type ReminderRun } from './api';
import { Alert, Dialog, Icon } from './components';

const dayDiff = (from: string, to: string) => Math.round((parseYmd(to).getTime() - parseYmd(from).getTime()) / 86_400_000);

/** 프로그램 시작·날짜가 바뀔 때 뜨는 다가오는 검수 팝업 (디자인 Dialog '프로그램 시작 알림') */
export function ReminderPopup({ items, today, who, onClose, onOpen }: {
  items: Inspection[];
  today: string;
  who: string;
  onClose: () => void;
  onOpen: (i: Inspection | null) => void;
}) {
  const allTomorrow = items.every((i) => dayDiff(today, i.schedule.date!) === 1);
  const allToday = items.every((i) => i.schedule.date === today);
  const title = allToday ? `오늘 검수 예정 ${items.length}건` : allTomorrow ? `내일 검수 예정 ${items.length}건` : `다가오는 검수 ${items.length}건`;
  return (
    <Dialog
      title={title}
      icon={<Icon name="calendar" lg />}
      width={520}
      footer={<>
        <button className="btn" type="button" onClick={onClose} autoFocus>닫기</button>
        <button className="btn btn-primary" type="button" onClick={() => onOpen(items.length === 1 ? items[0] : null)}>상세 보기</button>
      </>}
    >
      <span className="help">{who} 님에게 배정된 검수입니다.</span>
      {items.map((i) => {
        const d = dayDiff(today, i.schedule.date!);
        return (
          <div key={i.id} className="list-i">
            <span className={`dday${d <= 0 ? '' : d === 1 ? ' dday-soon' : ' dday-far'}`}>{d <= 0 ? i.schedule.time ?? 'D-Day' : `D-${d}`}</span>
            <span>
              <b>{i.vendorName}</b> · {i.itemName} · {i.schedule.location}
              <span className="muted"> · {fmtMD(i.schedule.date)} {i.schedule.time ?? ''}</span>
              {i.status === 'changed' && <span className="badge st-changed" style={{ marginLeft: 6 }}>변경됨</span>}
            </span>
          </div>
        );
      })}
    </Dialog>
  );
}

/** 알림 메일 미리보기 (설정에서 '미리보기'를 켰을 때) */
export function ReminderPreview({ drafts, onDone }: { drafts: MailDraft[]; onDone: (r: ReminderRun | null, error?: string) => void }) {
  const [list, setList] = useState(drafts);
  const [k, setK] = useState(0);
  const [busy, setBusy] = useState(false);
  const d = list[k];
  return (
    <Dialog
      title={`검수자 알림 메일 ${list.length}통`}
      icon={<Icon name="mail" lg />}
      width={640}
      footer={<>
        <button className="btn" type="button" onClick={() => onDone(null)}>오늘은 보내지 않기</button>
        <button className="btn btn-primary" type="button" disabled={busy} onClick={async () => {
          setBusy(true);
          const r = await api.reminderSend(list);
          setBusy(false);
          onDone(r.ok ? r.value : null, r.ok ? undefined : r.error.message);
        }}>{busy ? '보내는 중…' : `${list.length}통 모두 보내기`}</button>
      </>}
    >
      <Alert kind="info">오늘 처음 프로그램을 연 PC라서 검수자 알림 메일을 보냅니다. 다른 PC에서는 다시 보내지 않습니다.</Alert>
      <div className="tabs" role="tablist" aria-label="받는 사람">
        {list.map((m, n) => (
          <button key={m.to} className="tab" role="tab" type="button" aria-selected={n === k} onClick={() => setK(n)}>{m.to}</button>
        ))}
      </div>
      <div className="field">
        <label htmlFor="rs">제목</label>
        <input id="rs" className="input" value={d.subject} onChange={(e) => setList(list.map((x, n) => (n === k ? { ...x, subject: e.target.value } : x)))} />
      </div>
      <div className="field">
        <label htmlFor="rb">본문</label>
        <textarea id="rb" className="input" style={{ height: 200 }} value={d.body} onChange={(e) => setList(list.map((x, n) => (n === k ? { ...x, body: e.target.value } : x)))} />
      </div>
    </Dialog>
  );
}
