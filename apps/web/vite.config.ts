import { defineConfig, type Plugin } from 'vite';
import react from '@vitejs/plugin-react';
import { cpSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ---------------------------------------------------------------------------
// COOP/COEP — обязательное условие работы SharedArrayBuffer, который использует
// многопоточный WASM-модуль LibreOffice (Emscripten + pthreads).
//
//   Cross-Origin-Opener-Policy:   same-origin  — изоляция окна
//   Cross-Origin-Embedder-Policy: require-corp — разрешаем только «корпоративные» ресурсы
//   Cross-Origin-Resource-Policy: same-origin  — помечаем собственные ресурсы
//
// Заголовки настраиваются и для dev-сервера, и для `vite preview`.
// В production их должен отдавать ваш веб-сервер (nginx, Caddy и т. п.) — см. README.
// ---------------------------------------------------------------------------
const CROSS_ORIGIN_ISOLATION_HEADERS: Record<string, string> = {
  'Cross-Origin-Opener-Policy': 'same-origin',
  'Cross-Origin-Embedder-Policy': 'require-corp',
  'Cross-Origin-Resource-Policy': 'same-origin',
};

// ---------------------------------------------------------------------------
// Поиск каталога @matbee/libreoffice-converter.
// Библиотека объявлена зависимостью пакета office-wasm, а при строгой раскладке
// pnpm (node-linker=isolated) она лежит в node_modules самого пакета, а не
// приложения. Поэтому резолвим её из контекста пакета office-wasm, а не apps/web.
// ---------------------------------------------------------------------------
const officeWasmRequire = createRequire(path.join(__dirname, '../../packages/office-wasm/index.ts'));
const LIB_PACKAGE_DIR = path.dirname(
  officeWasmRequire.resolve('@matbee/libreoffice-converter/package.json'),
);
const PUBLIC_DIR = path.join(__dirname, 'public');
const MARKER_FILE = path.join(PUBLIC_DIR, '.office-wasm-version');

/**
 * Копирует статические ресурсы LibreOffice WASM в public/:
 *   public/wasm/* — ~200 МБ WASM-бинарников и данных движка (URL из createWasmPaths('/wasm/'))
 *   public/dist/* — browser.worker.global.js и сопутствующие модули библиотеки
 *
 * Копирование выполняется один раз на версию библиотеки (метка в .office-wasm-version).
 * При сборке Vite сам переносит содержимое public/ в выходной каталог,
 * при разработке — раздаёт файлы напрямую.
 */
function copyOfficeAssets(): void {
  const { version } = JSON.parse(
    readFileSync(path.join(LIB_PACKAGE_DIR, 'package.json'), 'utf8'),
  ) as { version: string };

  const alreadyCopied =
    existsSync(MARKER_FILE) &&
    readFileSync(MARKER_FILE, 'utf8').trim() === version &&
    existsSync(path.join(PUBLIC_DIR, 'wasm')) &&
    existsSync(path.join(PUBLIC_DIR, 'dist'));

  if (alreadyCopied) return;

  mkdirSync(PUBLIC_DIR, { recursive: true });
  cpSync(path.join(LIB_PACKAGE_DIR, 'wasm'), path.join(PUBLIC_DIR, 'wasm'), { recursive: true });
  cpSync(path.join(LIB_PACKAGE_DIR, 'dist'), path.join(PUBLIC_DIR, 'dist'), { recursive: true });
  writeFileSync(MARKER_FILE, version);

  console.log(`[office-wasm] ресурсы LibreOffice WASM v${version} скопированы в public/`);
}

function officeWasmAssetsPlugin(): Plugin {
  return {
    name: 'office-wasm:copy-lib-assets',
    // buildStart вызывается и при старте dev-сервера, и в начале vite build
    buildStart() {
      copyOfficeAssets();
    },
  };
}

export default defineConfig({
  plugins: [react(), officeWasmAssetsPlugin()],

  server: {
    headers: CROSS_ORIGIN_ISOLATION_HEADERS,
  },

  preview: {
    headers: CROSS_ORIGIN_ISOLATION_HEADERS,
  },

  worker: {
    // Воркер конвертера (из пакета office-converter) — ES-модуль с import'ами
    // (подключает пакет office-wasm), поэтому воркеры собираются в формате ES-модулей
    format: 'es',
  },

  optimizeDeps: {
    // office-wasm и office-converter подключены как workspace-пакеты (симлинки),
    // Vite обрабатывает их как исходный код; пребандлинг им не нужен
    // и может помешать (в office-converter воркер подключается через ?worker)
    exclude: ['office-wasm', 'office-converter'],
  },
});
