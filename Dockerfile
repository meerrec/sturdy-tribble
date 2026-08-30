# ---------------------------------------------------------------------------
# Многоэтапная сборка приложения Office → PDF.
#
# Этап 1 (build): установка зависимостей pnpm-монорепо и production-сборка
# Vite-приложения apps/web. WASM-ресурсы LibreOffice (~240 МБ) копируются
# Vite-плагином (copyOfficeAssets в apps/web/vite.config.ts) из node_modules
# в выходной каталог автоматически.
#
# Этап 2 (nginx): раздача статики с обязательными заголовками Cross-Origin
# Isolation (COOP/COEP/CORP) — без них браузер не отдаст SharedArrayBuffer
# и многопоточный WASM-движок LibreOffice не запустится (см. README).
# ---------------------------------------------------------------------------

# --- Этап 1: сборка ---------------------------------------------------------
FROM node:22-alpine AS build

# Версия pnpm зафиксирована в корневом package.json (packageManager)
RUN npm install --global pnpm@11.21.0

WORKDIR /app

# Сначала копируем только манифесты — тогда слой с зависимостями
# переиспользуется из кэша Docker при изменении исходников
COPY package.json pnpm-lock.yaml pnpm-workspace.yaml ./
COPY packages/office-converter/package.json packages/office-converter/package.json
COPY packages/office-wasm/package.json packages/office-wasm/package.json
COPY apps/web/package.json apps/web/package.json

# --frozen-lockfile: ставим ровно то, что зафиксировано в pnpm-lock.yaml.
# postinstall esbuild разрешён настройкой allowBuilds в pnpm-workspace.yaml.
RUN pnpm install --frozen-lockfile

# Исходники монорепо (node_modules, dist, .git отсекаются .dockerignore)
COPY . .

# Корневой скрипт build = pnpm --filter web build (tsc -b && vite build)
RUN pnpm build

# --- Этап 2: веб-сервер -----------------------------------------------------
FROM nginx:alpine

# Кастомная конфигурация: заголовки COOP/COEP/CORP, SPA-fallback,
# кэширование WASM-ресурсов, эндпоинт /healthz
COPY docker/nginx.conf /etc/nginx/conf.d/default.conf

# Готовая статика: index.html, JS/CSS-бандлы и WASM-ресурсы (~240 МБ)
COPY --from=build /app/apps/web/dist /usr/share/nginx/html

EXPOSE 80

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD wget -q --spider http://127.0.0.1/healthz || exit 1
