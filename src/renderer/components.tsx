import type { ReactNode } from 'react';
import type { AppErrorInfo } from '../shared/types';

export type IconName =
  | 'home' | 'calendar' | 'list' | 'upload' | 'download' | 'building' | 'user' | 'gear'
  | 'mail' | 'search' | 'warn' | 'error' | 'info' | 'ok' | 'chev' | 'file' | 'clock';

export function Icon({ name, lg, style }: { name: IconName; lg?: boolean; style?: React.CSSProperties }) {
  return <span className={`ico i-${name}${lg ? ' ico-lg' : ''}`} aria-hidden="true" style={style} />;
}

type AlertKind = 'warn' | 'danger' | 'info' | 'ok';
const ALERT_ICON: Record<AlertKind, IconName> = { warn: 'warn', danger: 'error', info: 'info', ok: 'ok' };

export function Alert({ kind, icon, children, action }: { kind: AlertKind; icon?: IconName; children: ReactNode; action?: ReactNode }) {
  return (
    <div className={`alert alert-${kind}`} role={kind === 'danger' ? 'alert' : 'status'}>
      <Icon name={icon ?? ALERT_ICON[kind]} />
      <span>{children}</span>
      {action && <span className="act">{action}</span>}
    </div>
  );
}

export function ErrorAlert({ error, onRetry }: { error: AppErrorInfo; onRetry?: () => void }) {
  return (
    <Alert kind="danger" action={onRetry && <button className="btn btn-sm" type="button" onClick={onRetry}>다시 시도</button>}>
      {error.message}
    </Alert>
  );
}

export function Dialog({ title, icon, children, footer, role = 'dialog', width }: {
  title: string; icon?: ReactNode; children: ReactNode; footer: ReactNode; role?: 'dialog' | 'alertdialog'; width?: number;
}) {
  return (
    <div className="scrim" style={{ position: 'fixed', inset: 0, minHeight: 0, zIndex: 50 }}>
      <div className="dlg" role={role} aria-modal="true" aria-labelledby="dlg-title" style={width ? { width } : undefined}>
        <div className="dlg-h" id="dlg-title">{icon}{title}</div>
        <div className="dlg-b">{children}</div>
        <div className="dlg-f">{footer}</div>
      </div>
    </div>
  );
}

/** 저장 충돌(덮어쓰기 방지) 안내 */
export function ConflictDialog({ error, onReload }: { error: AppErrorInfo; onReload: () => void }) {
  return (
    <Dialog
      role="alertdialog"
      width={540}
      title="저장하지 못했습니다"
      icon={<Icon name="error" lg style={{ color: 'var(--danger-fg)' }} />}
      footer={<button className="btn btn-primary" type="button" onClick={onReload} autoFocus>새로 불러오기</button>}
    >
      <Alert kind="danger">{error.message}</Alert>
      <p style={{ margin: 0 }}>
        내가 불러온 뒤 다른 사람이 데이터를 바꿨기 때문에, 덮어쓰지 않도록 저장을 멈췄습니다. 파일은 바뀌지 않았습니다.
        새로 불러온 뒤 변경 내용을 다시 입력하세요.
      </p>
    </Dialog>
  );
}
