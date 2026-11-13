/**
 * 설치된 프로그램 실행 시험 (GitHub Actions 의 Windows PC 에서 사용)
 *   node scripts/smoke-windows.cjs "<설치된 InspectionScheduler.exe 경로>"
 *
 * 순서: 실행 → 첫 실행 화면에서 새 데이터 폴더 만들기 → 시험용 엑셀 가져오기 → 목록·대시보드 확인
 * 각 단계 화면을 smoke-output/ 에 저장합니다. 화면 오류가 하나라도 나면 실패로 끝납니다.
 */
const { _electron } = require('playwright');
const fs = require('fs');
const os = require('os');
const path = require('path');

const exe = process.argv[2];
const OUT = path.resolve('smoke-output');
const SAMPLE = path.resolve('samples', '1_정상_신규3건.xlsx');

(async () => {
  if (!exe || !fs.existsSync(exe)) throw new Error(`실행 파일이 없습니다: ${exe}`);
  fs.mkdirSync(OUT, { recursive: true });
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'is-ud-'));
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'is-data-'));
  const errors = [];

  const args = [`--user-data-dir=${userData}`, ...(process.platform === 'linux' ? ['--no-sandbox'] : [])];
  const app = await _electron.launch({ executablePath: exe, args, timeout: 60_000 });
  const w = await app.firstWindow();
  w.on('pageerror', (e) => errors.push(e.message));
  const shot = (n) => w.screenshot({ path: path.join(OUT, `${n}.png`) });

  await w.waitForSelector('text=데이터 폴더를 선택하세요', { timeout: 60_000 });
  await shot('1-first-run');
  await app.evaluate(({ dialog }, d) => { dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [d] }); }, dataDir);
  await w.click('text=새 데이터 폴더 만들기');
  await w.click('button:has-text("폴더 선택")');
  await w.fill('#nm', '시험사용자');
  await w.click('button:has-text("폴더 만들고 시작")');
  await w.waitForSelector('text=상태별 건수');
  await shot('2-dashboard');

  await w.click('a.nav-i:has-text("엑셀 가져오기")');
  await app.evaluate(({ dialog }, p) => { dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [p] }); }, SAMPLE);
  await w.click('button:has-text("파일 선택")');
  await w.waitForSelector('text=건 가져오기');
  await shot('3-import-preview');
  await w.click('button:has-text("건 가져오기")');
  await w.waitForSelector('text=가져오기 완료');

  await w.click('a.nav-i:has-text("목록")');
  await w.waitForSelector('text=IN-');
  await shot('4-list');

  const insp = JSON.parse(fs.readFileSync(path.join(dataDir, 'data', 'inspections.json'), 'utf8'));
  const originals = fs.readdirSync(path.join(dataDir, '원본'));
  await app.close();

  console.log(`가져온 건수: ${insp.items.length}, 원본 보관 폴더: ${originals.join(', ')}`);
  if (insp.items.length !== 3) throw new Error(`가져온 건수가 3이 아닙니다: ${insp.items.length}`);
  if (errors.length) throw new Error(`화면 오류: ${errors.join(' / ')}`);
  console.log('실행 시험 통과');
})().catch((e) => {
  console.error('실행 시험 실패:', e);
  process.exit(1);
});
