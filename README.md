# Office → PDF

Конвертация документов **DOCX** и **XLSX** в **PDF** полностью в браузере, на стороне клиента.
Движок — **LibreOffice, скомпилированный в WebAssembly** (пакет [`@matbee/libreoffice-converter`](https://www.npmjs.com/package/@matbee/libreoffice-converter)).

> Примечание: npm-пакета с именем `wasm-office` не существует. Реальным аналогом описанного
> порта LibreOffice в WASM является `@matbee/libreoffice-converter` — он и используется
> в пакете-обёртке `office-wasm` с требуемым API (`initOffice()`, `convertDocumentToPdf()`).

Благодаря полноценному LibreOffice конвертация сохраняет **стили, изображения, колонтитулы,
таблицы, формулы, диаграммы** — качество как у настольного LibreOffice, без сервера и
без отправки файлов куда-либо.

## Возможности

- Конвертация выполняется в **Web Worker** — интерфейс не блокируется
- WASM-модуль инициализируется **один раз** и кэшируется на время сессии
- Прогресс инициализации и конвертации, состояния загрузки и ошибок
- Предпросмотр результата в `<iframe>` через **Blob URL**, скачивание PDF
- Drag & drop и выбор файла диалогом
- Тёмная/светлая тема (по системной настройке)

## Структура проекта

```text
├── pnpm-workspace.yaml          # workspace: packages/* и apps/*
├── package.json                 # корневой package.json (скрипты dev/build/preview)
├── packages/
│   ├── office-wasm/             # обёртка над LibreOffice WASM
│   │   ├── index.ts             #   initOffice(), convertDocumentToPdf(), disposeOffice()
│   │   └── package.json
│   └── office-converter/        # конвертер: RPC-клиент + Web Worker
│       ├── client.ts            #   ConverterClient: worker, init()/convert()/dispose()
│       ├── worker/converter.worker.ts # Web Worker с протоколом postMessage
│       ├── protocol.ts          #   типы сообщений (WorkerRequest/WorkerResponse)
│       ├── detect-file-type.ts  #   определение типа файла по расширению
│       ├── index.ts             #   публичный API пакета
│       └── package.json
└── apps/
    └── web/                     # React 18 + Vite + Jotai + TypeScript
        ├── vite.config.ts      # COOP/COEP, копирование WASM-ресурсов, worker: es
        ├── index.html
        └── src/
            ├── main.tsx         # точка входа
            ├── App.tsx          # компоновка: загрузка, статусы, предпросмотр
            ├── atoms.ts         # глобальное состояние (Jotai)
            ├── styles.css
            ├── hooks/useConverter.ts   # адаптер: клиент пакета → Jotai-атомы
            └── components/      # FileUploader, ConvertButton, PDFViewer
```

## Требования

- **Node.js ≥ 18** и **pnpm ≥ 9**
- Браузер с поддержкой `SharedArrayBuffer`: **Chrome / Edge 92+**, **Firefox 79+**
  (Safari не поддерживает вложенные worker'ы, которые использует движок)
- ~240 МБ дискового пространства для WASM-ресурсов (скачиваются один раз, далее — из кэша браузера)

## Установка и запуск

```bash
# 1. Установка зависимостей (первый раз скачает ~240 МБ пакета LibreOffice WASM)
pnpm install

# 2. Режим разработки — http://localhost:5173
pnpm dev

# 3. Production-сборка + локальный предпросмотр
pnpm build
pnpm preview
```

При старте Vite-плагин из `apps/web/vite.config.ts` копирует WASM-ресурсы и
worker-скрипт библиотеки из `node_modules` в `apps/web/public/` (каталоги
`wasm/` и `dist/`, см. `.gitignore`). Если версия библиотеки обновилась —
ресурсы перекопируются автоматически.

## Как это работает

```text
┌─────────────── Главный поток ───────────────┐        ┌────────── Web Worker ──────────┐
│  React + Jotai (UI, статусы, прогресс)      │        │  converter.worker.ts         │
│  useConverter → client (office-converter)   │        │    └─ office-wasm             │
│        │ ─────── postMessage ─────────────► │        │        initOffice() (1 раз,   │
│        ▲                                    │        │        кэшируется)            │
│  Blob URL ──── iframe (предпросмотр PDF)    │        │        convertDocumentToPdf() │
└───────┬─────────────────────────────────────┘        │              │                 │
        │   PDF-байты (Transferable)                   │        ┌─────▼──────────┐     │
        └──────────────────────────────────────────────│        │ browser.worker │     │
                                                       │        │ (вложенный     │     │
                                                       │        │  worker библ.) │     │
                                                       │        │   LibreOffice  │     │
                                                       │        │   WASM + SAB   │     │
                                                       │        └────────────────┘     │
                                                       └───────────────────────────────┘
```

1. При открытии страницы `useConverter` через клиент пакета `office-converter`
   поднимает Web Worker и сразу запрашивает инициализацию движка
   (`initOffice()` кэширует промис в пакете `office-wasm`).
2. Пользователь выбирает DOCX/XLSX — содержимое читается в `ArrayBuffer`
   и кладётся в Jotai-атом.
3. По кнопке «Конвертировать» буфер уходит в worker сообщением `postMessage`;
   конвертация выполняется внутри worker'а (LibreOffice WASM), главный поток свободен.
4. Готовые PDF-байты возвращаются с transfer list, на главном потоке из них
   создаётся `Blob` и `URL.createObjectURL(...)`, который показывается в `<iframe>`.

## Cross-Origin Isolation (COOP/COEP)

WASM-модуль LibreOffice многопоточный (Emscripten + pthreads), поэтому странице
нужен `SharedArrayBuffer`, а ему — заголовки кросс-доменной изоляции:

```http
Cross-Origin-Opener-Policy: same-origin
Cross-Origin-Embedder-Policy: require-corp
Cross-Origin-Resource-Policy: same-origin
```

Для dev-сервера и `vite preview` они уже настроены в `vite.config.ts`.
В production их должен отдавать веб-сервер, например:

```nginx
# nginx
add_header Cross-Origin-Opener-Policy "same-origin";
add_header Cross-Origin-Embedder-Policy "require-corp";
add_header Cross-Origin-Resource-Policy "same-origin";
```

```caddy
# Caddy
header Cross-Origin-Opener-Policy "same-origin"
header Cross-Origin-Embedder-Policy "require-corp"
header Cross-Origin-Resource-Policy "same-origin"
```

Проверить изоляцию можно в DevTools: `console.log(crossOriginIsolated)` → `true`,
или на вкладке Application → Frames → Cross-Origin Isolated.

## Ограничения

- Поддерживаются только **DOCX** и **XLSX** (расширяется в `packages/office-wasm/index.ts`
  и `packages/office-converter/detect-file-type.ts` — движок умеет и другие форматы).
- Первая конвертация после открытия страницы дольше последующих (прогрев движка).
- Пути к WASM-ресурсам (`/wasm/`, `/dist/`) абсолютные — при деплое под поддоменом
  в не-корень потребуется база Vite (`base`).
