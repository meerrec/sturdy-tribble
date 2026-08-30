import { useCallback, useEffect } from 'react';
import { useAtomValue, useSetAtom } from 'jotai';
import type { SupportedFileType } from 'office-wasm';
import type { WorkerRequest, WorkerResponse } from '../worker/protocol';
import {
  officeStateAtom,
  officeInitProgressAtom,
  officeInitErrorAtom,
  fileAtom,
  conversionStateAtom,
  conversionProgressAtom,
  conversionErrorAtom,
  pdfUrlAtom,
  canConvertAtom,
} from '../atoms';

// ---------------------------------------------------------------------------
// Жизненный цикл worker'а — модульный синглтон.
//
// Worker намеренно НЕ завершается при размонтировании компонента: в dev-режиме
// StrictMode монтирует эффекты дважды, а повторная инициализация WASM стоит
// секунды и ~240 МБ повторной загрузки. Worker живёт до выгрузки страницы
// (браузер сам завершит его) или до HMR-обновления этого модуля.
// ---------------------------------------------------------------------------

/**
 * Незавершённый RPC-вызов конвертации: promise, ожидающий ответа worker'а.
 * pdfBuffer приходит через transfer list и всегда лежит на собственном
 * ArrayBuffer (worker передаёт копию) — поэтому тип строже, чем Uint8Array.
 */
interface PendingRequest {
  resolve: (pdfBuffer: Uint8Array<ArrayBuffer>) => void;
  reject: (error: Error) => void;
}

let worker: Worker | null = null;
let requestSeq = 0;
let convertInFlight = false; // синхронный guard: атомы обновляются асинхронно
const pendingRequests = new Map<number, PendingRequest>(); // requestId -> { resolve, reject }
let activeMessageHandler: ((message: WorkerResponse) => void) | null = null;

function getWorker(): Worker {
  if (worker) return worker;

  worker = new Worker(new URL('../worker/converter.worker.ts', import.meta.url), {
    // worker собирается Vite как ES-модуль и подключает пакет office-wasm
    type: 'module',
  });

  worker.onmessage = (event: MessageEvent<WorkerResponse>) => {
    activeMessageHandler?.(event.data);
  };

  worker.onerror = (event) => {
    console.error('[useConverter] ошибка worker:', event);
    activeMessageHandler?.({
      type: 'init-error',
      message: event.message || 'Не удалось загрузить Web Worker',
    });
  };

  // Сразу запрашиваем инициализацию движка — пока пользователь выбирает файл,
  // ~240 МБ WASM-данных успеют скачаться из сети (или кэша браузера).
  worker.postMessage({ type: 'init' } satisfies WorkerRequest);

  return worker;
}

// При HMR-обновлении этого модуля старый worker корректно завершаем
if (import.meta.hot) {
  import.meta.hot.dispose(() => {
    worker?.terminate();
    worker = null;
    pendingRequests.clear();
  });
}

/** Определяет тип файла по расширению имени. Возвращает null для неподдерживаемых. */
export function detectFileType(
  file: File | null | undefined,
): SupportedFileType | null {
  const name = (file?.name ?? '').toLowerCase();
  const extension = name.split('.').pop();
  if (extension === 'docx' || extension === 'xlsx') return extension;
  return null;
}

/**
 * Главный хук приложения: владеет worker'ом и всеми переходами состояния.
 * Возвращает значения атомов и действия для компонентов.
 */
export function useConverter() {
  // --- Чтение состояния ------------------------------------------------------
  const officeState = useAtomValue(officeStateAtom);
  const officeInitProgress = useAtomValue(officeInitProgressAtom);
  const officeInitError = useAtomValue(officeInitErrorAtom);
  const file = useAtomValue(fileAtom);
  const conversionState = useAtomValue(conversionStateAtom);
  const conversionProgress = useAtomValue(conversionProgressAtom);
  const conversionError = useAtomValue(conversionErrorAtom);
  const pdfUrl = useAtomValue(pdfUrlAtom);
  const canConvert = useAtomValue(canConvertAtom);

  // --- Запись состояния ------------------------------------------------------
  const setOfficeState = useSetAtom(officeStateAtom);
  const setOfficeInitProgress = useSetAtom(officeInitProgressAtom);
  const setOfficeInitError = useSetAtom(officeInitErrorAtom);
  const setFile = useSetAtom(fileAtom);
  const setConversionState = useSetAtom(conversionStateAtom);
  const setConversionProgress = useSetAtom(conversionProgressAtom);
  const setConversionError = useSetAtom(conversionErrorAtom);
  const setPdfUrl = useSetAtom(pdfUrlAtom);

  // --- Разбор сообщений от worker'а -----------------------------------------
  const handleMessage = useCallback(
    (data: WorkerResponse) => {
      switch (data.type) {
        case 'progress':
          // Прогресс приходит и от инициализации, и от конвертации;
          // worker помечает сообщение фазой
          if (data.phase === 'init') {
            setOfficeInitProgress({ percent: data.percent, message: data.message });
          } else {
            setConversionProgress({ percent: data.percent, message: data.message });
          }
          break;

        case 'init-done':
          setOfficeState('ready');
          setOfficeInitProgress(null);
          setOfficeInitError(null);
          break;

        case 'init-error':
          setOfficeState('error');
          setOfficeInitError(data.message);
          break;

        case 'convert-done': {
          const entry = pendingRequests.get(data.requestId);
          pendingRequests.delete(data.requestId);
          entry?.resolve(data.pdfBuffer);
          break;
        }

        case 'convert-error': {
          const entry = pendingRequests.get(data.requestId);
          pendingRequests.delete(data.requestId);
          entry?.reject(new Error(data.message || 'Неизвестная ошибка конвертации'));
          break;
        }
      }
    },
    [
      setOfficeState,
      setOfficeInitProgress,
      setOfficeInitError,
      setConversionProgress,
    ],
  );

  // --- Монтирование: регистрируем обработчик и поднимаем worker ---------------
  useEffect(() => {
    activeMessageHandler = handleMessage;
    getWorker();
    return () => {
      if (activeMessageHandler === handleMessage) activeMessageHandler = null;
    };
  }, [handleMessage]);

  // --- Действия ---------------------------------------------------------------

  /** Читает выбранный файл в ArrayBuffer и кладёт его в состояние. */
  const selectFile = useCallback(
    async (file: File) => {
      const type = detectFileType(file);
      if (!type) {
        throw new Error('Неподдерживаемый формат. Выберите файл DOCX или XLSX.');
      }

      const buffer = await file.arrayBuffer();

      // Выбрали новый файл — сбрасываем результат предыдущей конвертации
      if (pdfUrl) URL.revokeObjectURL(pdfUrl);
      setPdfUrl(null);
      setConversionState('idle');
      setConversionError(null);
      setConversionProgress(null);

      setFile({ name: file.name, size: file.size, type, buffer });
    },
    [pdfUrl, setFile, setPdfUrl, setConversionState, setConversionError, setConversionProgress],
  );

  /** Запускает конвертацию текущего файла в PDF. */
  const convert = useCallback(async () => {
    if (!file || conversionState === 'converting' || convertInFlight) return;

    convertInFlight = true;
    setConversionState('converting');
    setConversionProgress(null);
    setConversionError(null);

    const requestId = ++requestSeq;
    try {
      // RPC-вызов worker'а: обещание резолвится входящим convert-done
      const pdfBuffer = await new Promise<Uint8Array<ArrayBuffer>>((resolve, reject) => {
        pendingRequests.set(requestId, { resolve, reject });
        // Буфер передаём копией (structured clone), а не через transfer list:
        // тогда он остаётся в atom, и документ можно конвертировать повторно.
        // Для файлов в десятки мегабайт копирование на десктопе незаметно.
        getWorker().postMessage({
          type: 'convert',
          requestId,
          payload: { buffer: file.buffer, fileType: file.type },
        } satisfies WorkerRequest);
      });

      // Старый Blob URL больше не нужен — освобождаем память
      if (pdfUrl) URL.revokeObjectURL(pdfUrl);
      const blob = new Blob([pdfBuffer], { type: 'application/pdf' });
      setPdfUrl(URL.createObjectURL(blob));
      setConversionState('done');
    } catch (error) {
      setConversionError(error instanceof Error ? error.message : String(error));
      setConversionState('error');
    } finally {
      convertInFlight = false;
    }
  }, [file, conversionState, pdfUrl, setConversionState, setConversionProgress, setConversionError, setPdfUrl]);

  /** Повторная попытка инициализации движка после ошибки. */
  const retryInit = useCallback(() => {
    setOfficeState('initializing');
    setOfficeInitError(null);
    setOfficeInitProgress(null);
    getWorker().postMessage({ type: 'init' } satisfies WorkerRequest);
  }, [setOfficeState, setOfficeInitError, setOfficeInitProgress]);

  return {
    // Состояние движка
    officeState,
    officeInitProgress,
    officeInitError,
    // Выбранный файл и конвертация
    file,
    conversionState,
    conversionProgress,
    conversionError,
    pdfUrl,
    canConvert,
    // Действия
    selectFile,
    convert,
    retryInit,
  };
}
