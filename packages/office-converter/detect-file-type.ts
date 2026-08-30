import { isSupportedFileType } from 'office-wasm';
import type { SupportedFileType } from 'office-wasm';

/**
 * Определяет тип файла по имени. Возвращает null для неподдерживаемых.
 * Список форматов живёт в пакете office-wasm (FILE_TYPE_TO_EXTENSION) —
 * здесь ничего не нужно менять при добавлении нового формата.
 */
export function detectFileType(name: string): SupportedFileType | null {
  const extension = name.toLowerCase().split('.').pop() ?? '';
  return isSupportedFileType(extension) ? extension : null;
}
