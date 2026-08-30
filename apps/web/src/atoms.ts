import { atom } from 'jotai';
import type { SupportedFileType } from 'office-wasm';

// ---------------------------------------------------------------------------
// Глобальное состояние приложения (Jotai-атомы).
// Все переходы состояния выполняются в хуке useConverter, компоненты только читают.
// ---------------------------------------------------------------------------

/** Статус офисного движка: 'initializing' | 'ready' | 'error'. */
export type OfficeState = 'initializing' | 'ready' | 'error';

/** Прогресс загрузки/инициализации WASM-модуля или конвертации. */
export interface ProgressInfo {
  /** Процент выполнения, 0..100. */
  percent: number;
  /** Человекочитаемое описание текущего этапа. */
  message: string;
}

/**
 * Выбранный пользователем файл.
 * Буфер лежит в атоме, чтобы его можно было передать в worker повторно
 * (например, при повторной конвертации после ошибки).
 */
export interface UploadedFile {
  name: string;
  size: number;
  type: SupportedFileType;
  buffer: ArrayBuffer;
}

/** Статус конвертации: 'idle' | 'converting' | 'done' | 'error'. */
export type ConversionState = 'idle' | 'converting' | 'done' | 'error';

/** Статус офисного движка. */
export const officeStateAtom = atom<OfficeState>('initializing');

/** Прогресс загрузки/инициализации WASM-модуля (или null). */
export const officeInitProgressAtom = atom<ProgressInfo | null>(null);

/** Текст ошибки инициализации движка (или null). */
export const officeInitErrorAtom = atom<string | null>(null);

/** Выбранный пользователем файл (или null). */
export const fileAtom = atom<UploadedFile | null>(null);

/** Статус конвертации. */
export const conversionStateAtom = atom<ConversionState>('idle');

/** Прогресс текущей конвертации (или null). */
export const conversionProgressAtom = atom<ProgressInfo | null>(null);

/** Текст ошибки конвертации (или null). */
export const conversionErrorAtom = atom<string | null>(null);

/** Blob URL готового PDF для <iframe> (или null). */
export const pdfUrlAtom = atom<string | null>(null);

/** Производный атом: можно ли запустить конвертацию прямо сейчас. */
export const canConvertAtom = atom(
  (get) =>
    Boolean(get(fileAtom)) &&
    get(officeStateAtom) === 'ready' &&
    get(conversionStateAtom) !== 'converting',
);
