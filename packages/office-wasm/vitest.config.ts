import { configDefaults, defineConfig } from 'vitest/config';

// e2e-тесты поднимают движок LibreOffice WASM (~240 МБ, десятки секунд
// на первую инициализацию) — в обычный прогон и CI они не входят,
// запускаются явно: pnpm --filter office-wasm test:e2e
const runE2e = process.env.E2E === '1';

export default defineConfig({
  test: {
    exclude: runE2e ? configDefaults.exclude : [...configDefaults.exclude, '**/*.e2e.test.ts']
  }
});
