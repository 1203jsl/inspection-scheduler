import { build } from 'esbuild';
await build({
  entryPoints: { main: 'src/main/main.ts', preload: 'src/main/preload.ts' },
  outdir: 'dist/main',
  bundle: true,
  platform: 'node',
  format: 'cjs',
  target: 'node20',
  external: ['electron', 'exceljs'],
  sourcemap: true,
  logLevel: 'info',
});
