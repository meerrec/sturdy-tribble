/**
 * office-wasm — обёртка над LibreOffice, скомпилированным в WebAssembly.
 *
 * Модуль предназначен для запуска ИЗНУТРИ Web Worker: инициализация движка и
 * конвертация занимают секунды и сотни мегабайт памяти, поэтому на главном
 * потоке его подключать не стоит.
 *
 * Базовый пакет: @matbee/libreoffice-converter (LibreOffice → WASM).
 * Примечание: npm-пакета с именем «wasm-office» не существует — реальный
 * аналог описанного порта LibreOffice именно этот пакет.
 */
import { WorkerBrowserConverter, createWasmPaths } from '@matbee/libreoffice-converter/browser';

/** Типы исходных документов, поддерживаемые приложением. */
export type SupportedFileType = 'docx' | 'xlsx';

/** Событие прогресса инициализации или конвертации. */
export interface OfficeProgressInfo {
  /** Процент выполнения, 0..100. */
  percent: number;
  /** Человекочитаемое описание текущего этапа. */
  message: string;
}

/** Опции инициализации движка. */
export interface InitOfficeOptions {
  /**
   * Колбэк прогресса загрузки WASM-модуля.
   * Вызывается многократно: сначала при инициализации движка,
   * затем при каждой конвертации.
   */
  onProgress?: (info: OfficeProgressInfo) => void;
}

// ---------------------------------------------------------------------------
// Настройки размещения WASM-ресурсов.
// Файлы из node_modules/@matbee/libreoffice-converter/{wasm,dist} копируются
// в public/ приложения плагином из vite.config.ts и раздаются по этим путям.
// Пути абсолютные, чтобы они одинаково работали и из главного потока, и из worker'а.
// ---------------------------------------------------------------------------
const WASM_BASE_URL = '/wasm/';
// Worker-скрипт библиотеки, в котором выполняется конвертация (см. dist/ пакета)
const BROWSER_WORKER_URL = '/dist/browser.worker.global.js';

// Единственный экземпляр конвертера на всё время жизни worker'а.
// Worker живёт в течение всей сессии приложения, поэтому «кэш инициализации»
// совпадает с модульным скоупом: повторные вызовы initOffice() ничего не делают.
let converterInstance: WorkerBrowserConverter | null = null;
let initPromise: Promise<WorkerBrowserConverter> | null = null;

/** Отображение поддерживаемых типов файлов на расширение для LibreOffice. */
const FILE_TYPE_TO_EXTENSION: Record<SupportedFileType, string> = {
  docx: 'docx',
  xlsx: 'xlsx'
};

/**
 * Все поддерживаемые типы исходных документов — единый источник правды:
 * от него строятся detectFileType в office-converter и accept у <input type="file">.
 * Добавление нового формата = одна строка в FILE_TYPE_TO_EXTENSION.
 */
export const SUPPORTED_FILE_TYPES = Object.keys(FILE_TYPE_TO_EXTENSION) as SupportedFileType[];

/** MIME-типы форматов (для атрибута accept и файловых диалогов). */
export const FILE_TYPE_TO_MIME: Record<SupportedFileType, string> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

/** Проверяет, поддерживается ли расширение файла (без точки). */
export function isSupportedFileType(extension: string): extension is SupportedFileType {
  return extension in FILE_TYPE_TO_EXTENSION;
}

/**
 * Инициализирует движок LibreOffice WASM.
 *
 * Выполняется один раз: результат кэшируется в initPromise, поэтому функцию
 * можно безопасно вызывать сколько угодно раз — повторные вызовы просто
 * дождутся уже идущей инициализации. При ошибке кэш сбрасывается, чтобы
 * следующий вызов мог попробовать заново (например, после обрыва сети).
 *
 * @param options — необязательные опции (см. {@link InitOfficeOptions}).
 * @returns Резолвится экземпляром конвертера, когда движок готов.
 */
export async function initOffice(options: InitOfficeOptions = {}): Promise<WorkerBrowserConverter> {
  if (initPromise) {
    return initPromise;
  }

  initPromise = (async () => {
    const { onProgress } = options;

    // WorkerBrowserConverter сам запускает отдельный worker (browser.worker.global.js),
    // внутри которого работает LibreOffice WASM. Наш worker в этом случае —
    // «диспетчер»: главный поток общается только с ним по postMessage,
    // а вся тяжёлая работа остаётся за пределами главного потока.
    const converter = new WorkerBrowserConverter({
      ...createWasmPaths(WASM_BASE_URL), // пути к WASM-бинарникам и данным
      browserWorkerJs: BROWSER_WORKER_URL,
      onProgress: (info) => {
        onProgress?.({ percent: info.percent, message: info.message });
      }
    });

    // Загружает WASM-модуль (~240 МБ при первом запуске, далее — из кэша браузера)
    // и готовит офисный движок к работе.
    await converter.initialize();

    converterInstance = converter;
    return converter;
  })().catch((error) => {
    // Сбрасываем кэш, чтобы неудачная инициализация не «залипла» навсегда
    initPromise = null;
    converterInstance = null;
    throw error;
  });

  return initPromise;
}

/**
 * Конвертирует документ DOCX или XLSX в PDF.
 *
 * Если движок ещё не инициализирован, инициализация выполнится автоматически.
 * Конвертация выполняется в worker'е библиотеки (внутри нашего worker'а),
 * поэтому главный поток не блокируется.
 *
 * @param fileBuffer — бинарное содержимое файла.
 * @param fileType — тип исходного документа. Определяет,
 *        какой фильтр импорта использует LibreOffice (по расширению имени файла).
 * @returns Байты готового PDF-документа.
 * @throws {TypeError} — если `fileType` не входит в {@link SupportedFileType}.
 */
export async function convertDocumentToPdf(
  fileBuffer: ArrayBuffer | Uint8Array,
  fileType: SupportedFileType
): Promise<Uint8Array> {
  const extension = FILE_TYPE_TO_EXTENSION[fileType];
  if (!extension) {
    throw new TypeError(
      `Неподдерживаемый тип файла: "${fileType}". Поддерживаются только: ${Object.keys(FILE_TYPE_TO_EXTENSION).join(', ')}`
    );
  }

  const converter = await initOffice();

  // Имя файла важно: LibreOffice выбирает фильтр импорта по расширению.
  // fileType задаётся приложением явно, а не выводится из имени пользовательского файла.
  const sourceFilename = `document.${extension}`;

  const result = await converter.convert(fileBuffer, { outputFormat: 'pdf' }, sourceFilename);

  return result.data;
}

/**
 * Завершает работу движка и освобождает ресурсы.
 * Полезно вызывать при выгрузке приложения; после этого initOffice() снова
 * выполнит полную инициализацию.
 */
export async function disposeOffice(): Promise<void> {
  const converter = converterInstance;
  initPromise = null;
  converterInstance = null;
  if (converter) {
    await converter.destroy();
  }
}
