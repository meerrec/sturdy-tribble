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

  /**
   * Создаёт worker при первом обращении и сразу запрашивает инициализацию
   * движка — пока пользователь выбирает файл, ~240 МБ WASM-данных успеют
   * скачаться из сети (или кэша браузера).
   */
  private ensureWorker(): Worker {
    if (this.worker) return this.worker;

    // worker собирается Vite как ES-модуль (см. worker.format в vite.config)
    // и подключает пакет office-wasm
    const worker = new ConverterWorker();

    worker.onmessage = (event: MessageEvent<WorkerResponse>) => {
      this.handleMessage(event.data);
    };

    worker.onerror = (event) => {
      console.error('[office-converter] ошибка worker:', event);
      this.notify({
        type: 'init-error',
        message: event.message || 'Не удалось загрузить Web Worker',
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
   * Буфер передаётся копией (structured clone), а не через transfer list:
   * тогда он остаётся у вызывающего кода, и документ можно конвертировать
   * повторно. Для файлов в десятки мегабайт копирование на десктопе незаметно.
   *
   * @returns Байты готового PDF-документа.
   * @throws {Error} — если конвертация уже выполняется или worker вернул ошибку.
   */
  convert(buffer: ArrayBuffer, fileType: SupportedFileType): Promise<Uint8Array<ArrayBuffer>> {
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
      this.ensureWorker().postMessage({
        type: 'convert',
        requestId,
        payload: { buffer, fileType },
      } satisfies WorkerRequest);
    });

    return promise.finally(() => {
      this.convertInFlight = false;
    });
  }

  /**
   * Завершает worker и освобождает ресурсы. Незавершённые конвертации
   * отклоняются ошибкой. Клиент остаётся пригодным: следующий вызов
   * init()/convert() создаст worker заново.
   */
  dispose(): void {
    if (this.worker) {
      this.worker.terminate();
      this.worker = null;
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
