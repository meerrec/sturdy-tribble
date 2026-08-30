/**
 * office-converter — конвертер документов Office в PDF поверх пакета office-wasm.
 *
 * Пакет упаковывает весь «конвейер» конвертации в переиспользуемую единицу:
 *   - client.ts — RPC-клиент на главном потоке: владеет Web Worker'ом,
 *     превращает протокол postMessage в вызовы init()/convert()/dispose()
 *     и рассылает события подписчикам через subscribe();
 *   - worker/converter.worker.ts — Web Worker, выполняющий инициализацию
 *     движка и конвертацию вне главного потока;
 *   - protocol.ts — типы сообщений протокола (один источник правды
 *     для клиента и worker'а);
 *   - detect-file-type.ts — определение типа файла по расширению.
 *
 * Пакет не зависит от фреймворков: всё UI-состояние (например, атомы Jotai
 * приложения) живёт у потребителя, который подписывается на события клиента.
 */
export { ConverterClient, getConverterClient } from './client';
export type { ConverterListener } from './client';
export { detectFileType } from './detect-file-type';
export type { WorkerRequest, WorkerResponse } from './protocol';
