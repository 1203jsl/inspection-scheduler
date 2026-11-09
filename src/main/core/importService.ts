import { promises as fs } from 'node:fs';
import path from 'node:path';
import type { DataFiles, ImportPreview, ImportSummary, PreviewFile, PreviewRow } from '../../shared/types';
import { DataStore, DIRS } from './dataStore';
import { AppError, toErrorInfo } from './errors';
import { readWorkbook } from './excel/read';
import { applyImport, buildPreviewRows, finalizePreview, originalFileName, type ValidateContext } from './importer';
import { importConflicts } from './importConflicts';

const pad = (n: number) => String(n).padStart(2, '0');
export const ymd = (d: Date) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

/**
 * 엑셀 가져오기 순서:
 *  preview(파일들) → 사용자가 행 선택 → commit(선택한 행)
 * commit 은 화면이 보낸 값이 아니라 preview 때 읽어 둔 값을 씁니다.
 * preview 이후 다른 사람이 데이터를 바꿨으면 저장 단계에서 CONFLICT 로 멈춥니다.
 */
export class ImportService {
  private last: {
    preview: ImportPreview;
    buffers: Buffer[];
    /** 미리보기 때 읽은 데이터 (이 위에 반영) */
    inspFile: DataFiles['inspections'];
    vendFile: DataFiles['vendors'];
  } | null = null;

  constructor(
    private store: DataStore,
    private user: () => { account: string; displayName: string },
    private now: () => Date = () => new Date(),
  ) {}

  async preview(paths: string[]): Promise<ImportPreview> {
    const settings = await this.store.load('settings');
    const inspections = await this.store.load('inspections');
    const vendors = await this.store.load('vendors');
    const inspectors = await this.store.load('inspectors');
    const ctx: ValidateContext = {
      settings,
      existing: inspections.items,
      today: ymd(this.now()),
      checkConflicts: (rows) => importConflicts(rows, inspections.items, settings, inspectors.items),
    };

    const files: PreviewFile[] = [];
    const buffers: Buffer[] = [];
    let rows: PreviewRow[] = [];
    for (const [i, p] of paths.entries()) {
      const name = path.basename(p);
      let buf: Buffer;
      try {
        buf = await fs.readFile(p);
      } catch {
        files.push({ path: p, name, rowCount: 0, error: { code: 'UNKNOWN', message: `파일을 열 수 없습니다 (${name}). 엑셀에서 열려 있으면 닫고 다시 넣으세요.` } });
        buffers.push(Buffer.alloc(0));
        continue;
      }
      buffers.push(buf);
      if (!/\.xlsx$/i.test(name)) {
        files.push({ path: p, name, rowCount: 0, error: { code: 'UNKNOWN', message: `${name}: .xlsx 파일만 가져올 수 있습니다. 엑셀에서 "Excel 통합 문서(.xlsx)"로 저장한 뒤 넣으세요.` } });
        continue;
      }
      try {
        const res = await readWorkbook(buf, settings);
        files.push({ path: p, name, templateVersion: res.templateVersion, rowCount: res.rows.length, warnings: res.warnings });
        rows = rows.concat(buildPreviewRows(res.rows, i, name, res.templateVersion, ctx));
      } catch (e) {
        const info = toErrorInfo(e);
        files.push({ path: p, name, rowCount: 0, error: { ...info, message: `${name}: ${info.message}` } });
      }
    }
    const preview = { files, rows: finalizePreview(rows, ctx) };
    this.last = { preview, buffers, inspFile: inspections, vendFile: vendors };
    return preview;
  }

  async commit(includeKeys: string[]): Promise<ImportSummary> {
    if (!this.last) throw new AppError('UNKNOWN', '미리보기를 먼저 해야 합니다. 파일을 다시 넣어 주세요.');
    const { preview, buffers, inspFile, vendFile } = this.last;
    const include = new Set(includeKeys);
    const now = this.now();

    // 1) 원본 보관 (포함한 행이 있는 파일만)
    const usedFiles = new Set(preview.rows.filter((r) => include.has(r.key)).map((r) => r.fileIndex));
    const month = `${now.getFullYear()}-${pad(now.getMonth() + 1)}`;
    const destDir = path.join(this.store.root, DIRS.original, month);
    const sourcePaths: string[] = [];
    const copied: string[] = [];
    try {
      for (const [i, f] of preview.files.entries()) {
        if (!usedFiles.has(i)) {
          sourcePaths.push('');
          continue;
        }
        const vendors = preview.rows.filter((r) => r.fileIndex === i).map((r) => r.parsed.vendorName);
        let name = originalFileName(vendors, now, f.name);
        await fs.mkdir(destDir, { recursive: true });
        for (let n = 2; await fs.access(path.join(destDir, name)).then(() => true, () => false); n++) {
          name = name.replace(/(\(\d+\))?\.xlsx$/i, `(${n}).xlsx`);
        }
        await fs.writeFile(path.join(destDir, name), buffers[i], { flag: 'wx' });
        copied.push(path.join(destDir, name));
        sourcePaths.push(`${DIRS.original}/${month}/${name}`);
      }
    } catch (e) {
      for (const c of copied) await fs.rm(c, { force: true });
      throw new AppError('UNKNOWN', `원본 파일을 원본/ 폴더에 복사하지 못했습니다. 동기화 폴더 연결을 확인한 뒤 다시 가져오세요. (${(e as Error).message})`);
    }

    // 2) 반영 후 저장 — 미리보기 때 읽은 데이터 기준 (그 뒤 바뀌었으면 save 가 CONFLICT)
    try {
      const out = applyImport({
        rows: preview.rows, includeKeys: include, inspections: inspFile.items, vendors: vendFile.items,
        sourcePaths, user: this.user(), now,
      });
      await this.store.save({
        inspections: { ...inspFile, items: out.inspections },
        vendors: { ...vendFile, items: out.vendors },
      });
      this.last = null;
      return {
        created: out.created.map((c) => ({ id: c.id, vendorName: c.vendorName, itemName: c.itemName, contactEmail: c.contactEmail })),
        updated: out.updated.map((u) => u.id),
        excluded: out.excluded,
        originals: sourcePaths.filter(Boolean),
      };
    } catch (e) {
      for (const c of copied) await fs.rm(c, { force: true });
      throw e;
    }
  }
}
