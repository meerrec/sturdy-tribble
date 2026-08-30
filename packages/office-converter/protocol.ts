import type { SupportedFileType } from 'office-wasm';

// ---------------------------------------------------------------------------
// Типы протокола обмена сообщениями между главным потоком и converter.worker.
// Модуль содержит только типы: в сборку ничего не попадает (import type),
// а worker и главный поток разделяют один источник правды о формате сообщений.
// ---------------------------------------------------------------------------

/**
 * Запрос главного потока к worker'у.
 *   { type: 'init' } — инициализировать движок
 *   { type: 'convert', requestId, payload } — конвертировать файл
 *   { type: 'dispose' } — освободить ресурсы и завершиться
 */
export type WorkerRequest =
  | { type: 'init' }
  | {
      type: 'convert';
      requestId: number;
      payload: { buffer: ArrayBuffer; fileType: SupportedFileType };
    }
  | { type: 'dispose' };

/**
 * Ответ worker'а главному потоку.
 * Прогресс приходит и от инициализации движка, и от конвертации;
 * фаза определяется полем phase, requestId заполнен только для конвертации.
 */
export type WorkerResponse =
  | {
      type: 'progress';
      phase: 'init' | 'convert';
      requestId: number | null;
      percent: number;
      message: string;
    }
  | { type: 'init-done' }
  | { type: 'init-error'; message: string }
  | { type: 'convert-done'; requestId: number; pdfBuffer: Uint8Array<ArrayBuffer> }
  | { type: 'convert-error'; requestId: number; message: string };
