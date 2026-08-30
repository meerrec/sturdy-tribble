import { describe, expect, it, vi } from 'vitest';
import { ConverterClient } from './client';
import type { WorkerRequest } from './protocol';

// В тестах реальный worker-модуль (Vite-специфичный импорт ?worker) не загружаем:
// рабочие экземпляры подставляются через createWorker.
vi.mock('./worker/converter.worker.ts?worker', () => ({ default: class {} }));

/** Минимальная заглушка Web Worker'а для тестов RPC-клиента. */
class FakeWorker {
  posted: { message: unknown; transfer?: Transferable[] }[] = [];
  terminated = false;
  onmessage: ((event: MessageEvent) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;

  postMessage(message: unknown, transfer?: Transferable[]): void {
    this.posted.push({ message, transfer });
  }

  terminate(): void {
    this.terminated = true;
  }

  /** Имитирует входящее сообщение от worker'а. */
  emit(data: unknown): void {
    this.onmessage?.({ data } as MessageEvent);
  }

  /** Имитирует аварийное завершение worker'а. */
  fail(message = 'boom'): void {
    this.onerror?.({ message } as ErrorEvent);
  }
}

/** Создаёт клиент, worker'ом которого управляет тест. */
function createTestClient(worker: FakeWorker): ConverterClient {
  return new ConverterClient({ createWorker: () => worker as unknown as Worker });
}

/** Находит среди отправленных сообщений сообщение заданного типа. */
function findPosted(worker: FakeWorker, type: WorkerRequest['type']) {
  return worker.posted.find((m) => (m.message as WorkerRequest).type === type);
}

describe('ConverterClient', () => {
  it('init создаёт worker лениво и шлёт init при каждом вызове', () => {
    const fake = new FakeWorker();
    const client = createTestClient(fake);

    client.init();
    client.init();

    // Первое сообщение — автозапуск инициализации при создании worker'а,
    // ещё два — по одному на каждый вызов init()
    expect(fake.posted).toHaveLength(3);
    expect(fake.posted.every((m) => (m.message as WorkerRequest).type === 'init')).toBe(true);
  });

  it('convert передаёт буфер по transfer list и резолвится pdfBuffer', async () => {
    const fake = new FakeWorker();
    const client = createTestClient(fake);
    const buffer = new ArrayBuffer(8);
    const pdf = new Uint8Array([1, 2, 3]);

    const promise = client.convert(buffer, 'docx', { transfer: true });
    const convertMsg = findPosted(fake, 'convert');
    expect(convertMsg?.transfer).toEqual([buffer]);

    fake.emit({ type: 'convert-done', requestId: 1, pdfBuffer: pdf });
    await expect(promise).resolves.toBe(pdf);
  });

  it('без transfer буфер передаётся копией и остаётся у вызывающего кода', () => {
    const fake = new FakeWorker();
    const client = createTestClient(fake);
    const buffer = new ArrayBuffer(8);

    void client.convert(buffer, 'xlsx');
    const convertMsg = findPosted(fake, 'convert');
    expect(convertMsg?.transfer).toEqual([]);
    expect(buffer.byteLength).toBe(8);
  });

  it('convert отклоняется при convert-error от worker', async () => {
    const fake = new FakeWorker();
    const client = createTestClient(fake);

    const promise = client.convert(new ArrayBuffer(4), 'xlsx');
    fake.emit({ type: 'convert-error', requestId: 1, message: 'движок сломался' });
    await expect(promise).rejects.toThrow('движок сломался');
  });

  it('отклоняет вторую конвертацию, пока первая выполняется', async () => {
    const fake = new FakeWorker();
    const client = createTestClient(fake);

    void client.convert(new ArrayBuffer(4), 'docx');
    await expect(client.convert(new ArrayBuffer(4), 'docx')).rejects.toThrow('уже выполняется');
  });

  it('аварийное завершение worker отклоняет вызовы и позволяет пересоздать worker', async () => {
    const workers = [new FakeWorker(), new FakeWorker()];
    let created = 0;
    const client = new ConverterClient({
      createWorker: () => workers[created++] as unknown as Worker
    });
    const listener = vi.fn();
    client.subscribe(listener);

    const promise = client.convert(new ArrayBuffer(4), 'docx');
    workers[0].fail('boom');

    await expect(promise).rejects.toThrow('аварийно');
    expect(listener).toHaveBeenCalledWith({ type: 'init-error', message: 'boom' });

    // Следующий вызов создаёт нового worker'а, и конвертация на нём проходит
    client.init();
    expect(findPosted(workers[1], 'init')).toBeDefined();

    const next = client.convert(new ArrayBuffer(4), 'docx');
    workers[1].emit({ type: 'convert-done', requestId: 2, pdfBuffer: new Uint8Array([1]) });
    await expect(next).resolves.toEqual(new Uint8Array([1]));
  });

  it('dispose отклоняет незавершённые вызовы, шлёт dispose и завершает worker по таймауту', async () => {
    vi.useFakeTimers();
    try {
      const fake = new FakeWorker();
      const client = createTestClient(fake);
      const promise = client.convert(new ArrayBuffer(4), 'docx');

      client.dispose();
      await expect(promise).rejects.toThrow('уничтожен');
      expect(findPosted(fake, 'dispose')).toBeDefined();
      expect(fake.terminated).toBe(false);

      vi.advanceTimersByTime(1000);
      expect(fake.terminated).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });

  it('subscribe рассылает сообщения и отписывает', () => {
    const fake = new FakeWorker();
    const client = createTestClient(fake);
    const listener = vi.fn();
    const unsubscribe = client.subscribe(listener);

    client.init(); // создаёт worker и подключает обработчик onmessage
    fake.emit({ type: 'init-done' });
    expect(listener).toHaveBeenCalledWith({ type: 'init-done' });

    unsubscribe();
    fake.emit({ type: 'init-done' });
    expect(listener).toHaveBeenCalledTimes(1);
  });
});
