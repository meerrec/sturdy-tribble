/**
 * Web Worker, выполняющий конвертацию документов.
 *
 * Сюда приходят сообщения с главного потока, вся работа с LibreOffice WASM
 * (включая его загрузку на ~240 МБ) выполняется здесь и НИКОГДА не блокирует UI.
 *
 * Протокол сообщений (главный поток → worker):
 *   { type: 'init' }                                   — инициализировать движок
 *   { type: 'convert', requestId, payload: { buffer, fileType } } — конвертировать файл
 *   { type: 'dispose' }                                — освободить ресурсы и завершиться
 *
 * Протокол сообщений (worker → главный поток):
 *   { type: 'progress', phase: 'init' | 'convert', requestId, percent, message }
 *   { type: 'init-done' } | { type: 'init-error', message }
 *   { type: 'convert-done', requestId, pdfBuffer } | { type: 'convert-error', requestId, message }
 *
 * Точные типы сообщений описаны в ../protocol.ts.
 */
import { initOffice, convertDocumentToPdf, disposeOffice } from 'office-wasm';
import type { WorkerRequest, WorkerResponse } from '../protocol';

// В tsconfig приложения подключён lib "DOM", в котором self имеет тип Window
// и не описывает контракт worker'а. Заводить отдельный tsconfig ради одного
// файла избыточно, поэтому объявляем нужный минимум DedicatedWorkerGlobalScope
// структурно и приводим тип.
interface WorkerScope {
  onmessage: ((event: MessageEvent<WorkerRequest>) => void) | null;
  postMessage(message: WorkerResponse, transfer?: Transferable[]): void;
  close(): void;
}

const ctx = self as unknown as WorkerScope;

let initDone = false; // движок успешно инициализирован
let activeRequestId: number | null = null; // id конвертации, выполняющейся в данный момент

/** Ошибку сериализуем в строку — Error-объекты не проходят через postMessage. */
function postError(
  type: 'init-error' | 'convert-error',
  requestId: number | null,
  error: unknown,
): void {
  console.error(`[converter.worker] ${type}:`, error);
  const message = error instanceof Error ? error.message : String(error);

  // У инициализации нет requestId; у конвертации он есть всегда.
  if (type === 'init-error') {
    ctx.postMessage({ type, message });
  } else if (requestId !== null) {
    ctx.postMessage({ type, requestId, message });
  }
}

ctx.onmessage = async (event) => {
  const { data } = event;

  switch (data.type) {
    case 'init': {
      try {
        // initOffice кэширует инициализацию внутри пакета office-wasm:
        // повторные вызовы ничего не скачивают и не запускают заново.
        await initOffice({
          onProgress: (info) => {
            // Колбэк прогресса у библиотеки один на весь жизненный цикл конвертера,
            // поэтому фазу определяем сами: до окончания init — загрузка движка,
            // после — прогресс текущей конвертации.
            ctx.postMessage({
              type: 'progress',
              phase: initDone ? 'convert' : 'init',
              requestId: initDone ? activeRequestId : null,
              percent: info.percent,
              message: info.message,
            });
          },
        });
        initDone = true;
        ctx.postMessage({ type: 'init-done' });
      } catch (error) {
        // Кэш initOffice при ошибке сбрасывается внутри пакета,
        // поэтому повторный 'init' выполнит инициализацию заново.
        postError('init-error', null, error);
      }
      break;
    }

    case 'convert': {
      const { requestId, payload } = data;
      activeRequestId = requestId;
      try {
        const pdfBytes = await convertDocumentToPdf(payload.buffer, payload.fileType);

        // slice() создаёт собственную копию: результат библиотеки может быть
        // представлением (view) на более крупном ArrayBuffer, а передавать
        // «лишние» байты через transfer нельзя без потери смещения.
        const copy = pdfBytes.slice();

        // Transfer List: байты PDF уезжают на главный поток без копирования —
        // PDF бывают большими, а лишнее клонирование ни к чему.
        ctx.postMessage({ type: 'convert-done', requestId, pdfBuffer: copy }, [copy.buffer]);
      } catch (error) {
        postError('convert-error', requestId, error);
      } finally {
        activeRequestId = null;
      }
      break;
    }

    case 'dispose': {
      try {
        await disposeOffice();
      } catch (error) {
        console.error('[converter.worker] ошибка при завершении:', error);
      }
      ctx.close(); // завершаем поток worker'а
      break;
    }
  }
};
