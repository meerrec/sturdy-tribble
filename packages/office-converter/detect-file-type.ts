import type { SupportedFileType } from 'office-wasm';

/** Определяет тип файла по расширению имени. Возвращает null для неподдерживаемых. */
export function detectFileType(
  file: File | null | undefined,
): SupportedFileType | null {
  const name = (file?.name ?? '').toLowerCase();
  const extension = name.split('.').pop();
  if (extension === 'docx' || extension === 'xlsx') return extension;
  return null;
}
