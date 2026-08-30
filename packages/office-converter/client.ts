import type { SupportedFileType } from 'office-wasm';
import type { WorkerRequest, WorkerResponse } from './protocol';
import ConverterWorker from './worker/converter.worker.ts?worker';

// ---------------------------------------------------------------------------
// RPC-клиент конвертера: владеет Web Worker'ом и превращает протокол
// postMessage (см. protocol.ts) в удобные вызовы. Пакет намеренно не знает
// ни о каком фреймворке — всё состояние приложения (React/Jotai) живёт
// в потребителе, а сюда приходят только сообщения через subscribe().
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

/** Подписчик на сообщения worker'а: прогресс, статусы инициализации и т. д. */
export type ConverterListener = (message: WorkerResponse) => void;

/** Опции клиента; createWorker позволяет подменить фабрику worker'а (например, в тестах). */
export interface ConverterClientOptions {
  createWorker?: () => Worker;
}

/**
 * Клиент конвертера — модульный синглтон (см. {@link getConverterClient}).
 * Одна страница — один worker: повторное создание двигателя стоит
 * сотни мегабайт повторной загрузки WASM-данных.
 */
export class ConverterClient {
  private worker: Worker | null = null;
  private readonly listeners = new Set<ConverterListener>();
  private requestSeq = 0;
  private convertInFlight = false; // синхронный guard: состояния потребителя обновляются асинхронно
  private readonly pendingRequests = new Map<number, PendingRequest>(); // requestId -> { resolve, reject }

  constructor(private readonly options: ConverterClientOptions = {}) {}

  /**
   * Создаёт worker при первом обращении и сразу запрашивает инициализацию
   * движка — пока пользователь выбирает файл, ~240 МБ WASM-данных успеют
   * скачаться из сети (или кэша браузера).
   */
  private ensureWorker(): Worker {
    if (this.worker) return this.worker;

    // worker собирается Vite как ES-модуль (см. worker.format в vite.config)
    // и подключает пакет office-wasm
    const worker = this.options.createWorker ? this.options.createWorker() : new ConverterWorker();

    worker.onmessage = (event: MessageEvent<WorkerResponse>) => {
      this.handleMessage(event.data);
    };

    worker.onerror = (event) => {
      console.error('[office-converter] ошибка worker:', event);
      // Worker аварийно завершился: незавершённые вызовы отклоняем,
      // чтобы promise'ы и guard не зависли навсегда. Следующий init()/convert()
      // создаст нового worker'а.
      this.worker?.terminate();
      this.worker = null;
      for (const { reject } of this.pendingRequests.values()) {
        reject(new Error('Web Worker аварийно завершился'));
      }
      this.pendingRequests.clear();
      this.convertInFlight = false;
      this.notify({
        type: 'init-error',
        message: event.message || 'Не удалось загрузить Web Worker'
      });
    };

    // Инициализацию запускаем сразу при создании worker'а, чтобы она шла
    // параллельно остальным действиям приложения
    worker.postMessage({ type: 'init' } satisfies WorkerRequest);

    this.worker = worker;
    return worker;
  }

  /** Подписывает на все сообщения worker'а. Возвращает функцию отписки. */
  subscribe(listener: ConverterListener): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  /**
   * Запускает инициализацию движка в worker'е (worker при необходимости
   * создаётся). Безопасно вызывать повторно: внутри worker'а initOffice
   * кэширует инициализацию, а при ошибке кэш сбрасывается — поэтому этот же
   * метод подходит и для повторной попытки после сбоя.
   */
  init(): void {
    this.ensureWorker().postMessage({ type: 'init' } satisfies WorkerRequest);
  }

  /**
   * Конвертирует документ в PDF внутри worker'а.
   *
   * По умолчанию буфер передаётся копией (structured clone) и остаётся
   * у вызывающего кода. С опцией `transfer` буфер переходит во владение
   * worker'а без копирования (в вызывающем коде он становится detached) —
   * так дешевле передавать большие файлы; для повторной конвертации
   * потребуется прочитать файл заново.
   *
   * @returns Байты готового PDF-документа.
   * @throws {Error} — если конвертация уже выполняется или worker вернул ошибку.
   */
  convert(
    buffer: ArrayBuffer,
    fileType: SupportedFileType,
    options?: { transfer?: boolean }
  ): Promise<Uint8Array<ArrayBuffer>> {
    if (this.convertInFlight) {
      // Практически недостижимо, если вызывающий код тоже следит за своим
      // состоянием; защита на случай гонок между независимыми вызовами.
      return Promise.reject(new Error('Конвертация уже выполняется'));
    }

    this.convertInFlight = true;
    const requestId = ++this.requestSeq;

    // RPC-вызов: обещание резолвится входящим convert-done
    const promise = new Promise<Uint8Array<ArrayBuffer>>((resolve, reject) => {
      this.pendingRequests.set(requestId, { resolve, reject });
      this.ensureWorker().postMessage(
        {
          type: 'convert',
          requestId,
          payload: { buffer, fileType }
        } satisfies WorkerRequest,
        options?.transfer ? [buffer] : []
      );
    });

    return promise.finally(() => {
      this.convertInFlight = false;
    });
  }

  /**
   * Завершает worker и освобождает ресурсы. Worker'у даётся возможность
   * корректно освободить движок (disposeOffice) и закрыться; если он
   * не успеет за 1 с — завершается принудительно. Незавершённые конвертации
   * отклоняются ошибкой. Клиент остаётся пригодным: следующий вызов
   * init()/convert() создаст worker заново.
   */
  dispose(): void {
    const worker = this.worker;
    this.worker = null;
    if (worker) {
      worker.postMessage({ type: 'dispose' } satisfies WorkerRequest);
      // Страховка на случай зависания: закрытый worker завершается сам,
      // для него terminate — безопасный no-op
      setTimeout(() => worker.terminate(), 1000);
    }
    for (const { reject } of this.pendingRequests.values()) {
      reject(new Error('Конвертер уничтожен до завершения конвертации'));
    }
    this.pendingRequests.clear();
    this.convertInFlight = false;
  }

  /** Разбирает входящее сообщение: резолвит RPC-вызовы и уведомляет подписчиков. */
  private handleMessage(data: WorkerResponse): void {
    if (data.type === 'convert-done' || data.type === 'convert-error') {
      const entry = this.pendingRequests.get(data.requestId);
      this.pendingRequests.delete(data.requestId);
      if (data.type === 'convert-done') {
        entry?.resolve(data.pdfBuffer);
      } else {
        entry?.reject(new Error(data.message || 'Неизвестная ошибка конвертации'));
      }
    }
    this.notify(data);
  }

  private notify(data: WorkerResponse): void {
    this.listeners.forEach((listener) => listener(data));
  }
}

// Единственный экземпляр клиента на всё время жизни приложения.
let client: ConverterClient | null = null;

/** Возвращает синглтон клиента конвертера, создавая его при первом вызове. */
export function getConverterClient(): ConverterClient {
  if (!client) client = new ConverterClient();
  return client;
}

// При HMR-обновлении этого модуля старый worker корректно завершаем
if (import.meta.hot) {
  import.meta.hot.dispose(() => {
    client?.dispose();
  });
}
