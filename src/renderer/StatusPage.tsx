import type { DataFileKey, DataFiles } from '../shared/types';
import { DATA_FILE_KEYS } from '../shared/types';
import { fmtDateTime } from './api';
import { Alert } from './components';

const LABEL: Record<DataFileKey, string> = {
  inspections: '검수 건',
  vendors: '업체',
  inspectors: '검수자',
  settings: '설정',
  mail_log: '메일 발송 기록',
};

/** 아직 만들지 않은 화면 자리: 개발 단계 안내 + 데이터 파일 상태 */
export function StatusPage({ data, step, label }: { data: DataFiles; step: string; label: string }) {
  return (
    <>
      <Alert kind="info">'{label}' 화면은 {step}에서 만듭니다. 지금은 데이터 파일 상태만 보여줍니다.</Alert>
      <div className="tbl-wrap">
        <table className="tbl">
          <thead>
            <tr><th>파일</th><th>내용</th><th className="num">항목 수</th><th className="num">저장 번호</th><th>마지막 저장</th></tr>
          </thead>
          <tbody>
            {DATA_FILE_KEYS.map((k) => {
              const f = data[k] as DataFiles[DataFileKey];
              const n = 'items' in f ? `${f.items.length}` : '-';
              return (
                <tr key={k}>
                  <td>data/{k}.json</td>
                  <td>{LABEL[k]}</td>
                  <td className="num">{n}</td>
                  <td className="num">{f.revision}</td>
                  <td>{f.updatedByName || f.updatedBy} · {fmtDateTime(f.updatedAt)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </>
  );
}
