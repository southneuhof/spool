import { build } from 'esbuild'
import { execFileSync } from 'node:child_process'
await build({entryPoints: ['src/index.ts', 'src/vue/index.ts'], outdir: 'dist', bundle: true, format: 'esm', platform: 'browser', target: 'es2022', packages: 'external'})
execFileSync('tsc', ['--ignoreConfig', '--declaration', '--emitDeclarationOnly', '--outDir', 'dist', '--rootDir', 'src', '--types', 'node', ...['src/index.ts', 'src/vue/index.ts'], '--target', 'ES2022', '--module', 'ESNext', '--moduleResolution', 'Bundler', '--strict', '--noUncheckedIndexedAccess', '--exactOptionalPropertyTypes', '--skipLibCheck', '--lib', 'ES2022,DOM,DOM.Iterable'], {stdio:'inherit'})
