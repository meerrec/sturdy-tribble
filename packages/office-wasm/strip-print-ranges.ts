/**
 * Удаление областей печати из XLSX перед конвертацией в PDF.
 *
 * LibreOffice Calc экспортирует лист в PDF строго в пределах области печати
 * (definedName `_xlnm.Print_Area`, которую Excel сохраняет в xl/workbook.xml).
 * Если область задана не на весь лист — а Excel делает это молча, например,
 * при «Set Print Area» или «Fit to page» — всё содержимое за её пределами
 * из PDF пропадает. Поэтому перед конвертацией области печати вырезаются:
 * без них Calc печатает весь used range листа с обычной разбивкой на страницы.
 *
 * `_xlnm.Print_Titles` (повторяющиеся на каждой странице строки/столбцы)
 * НЕ удаляются: они не обрезают данные, а только дублируют заголовки —
 * поведение сохраняем как задумано автором файла.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

/** Имя definedName, в котором Excel хранит область печати листа. */
const PRINT_AREA_DEFINED_NAME = '_xlnm.Print_Area';

/**
 * Элемент `<definedName name="_xlnm.Print_Area">…</definedName>`.
 * Префикс пространства имён (например, `<x:definedName>`), порядок и кавычки
 * атрибутов не важны — совпадение идёт по атрибуту `name`.
 */
const PRINT_AREA_RE = new RegExp(
  `<(?:[\\w-]+:)?definedName\\b[^>]*\\bname\\s*=\\s*["']${PRINT_AREA_DEFINED_NAME}["'][^>]*>` +
    '[\\s\\S]*?' +
    '<\\/(?:[\\w-]+:)?definedName\\s*>',
  'g'
);

/** Пустая обёртка `<definedNames></definedNames>`, оставшаяся после вырезания. */
const EMPTY_DEFINED_NAMES_RE =
  /<(?:[\w-]+:)?definedNames\b[^>]*>\s*<\/(?:[\w-]+:)?definedNames\s*>/g;

/**
 * Возвращает копию XLSX-файла без областей печати.
 *
 * @param buffer — содержимое .xlsx (zip-архив).
 * @returns Новый XLSX-буфер, либо `null`, если областей печати в файле нет
 *          (или файл — не zip-архив): тогда конвертировать нужно исходник.
 */
export function stripPrintRangesFromXlsx(buffer: Uint8Array): Uint8Array | null {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(buffer);
  } catch {
    // Не zip-архив (повреждённый файл или не xlsx) — оставляем как есть:
    // LibreOffice сам сообщит об ошибке при конвертации.
    return null;
  }

  // Имена внутри zip чувствительны к регистру, но Excel иногда пишет
  // «xl/Workbook.xml» — ищем без учёта регистра.
  const workbookPath = Object.keys(files).find((path) => path.toLowerCase() === 'xl/workbook.xml');
  if (!workbookPath) {
    return null;
  }

  const originalXml = strFromU8(files[workbookPath]);
  const strippedXml = originalXml.replace(PRINT_AREA_RE, '').replace(EMPTY_DEFINED_NAMES_RE, '');

  // Быстрый путь: областей печати не было — исходный буфер не трогаем.
  if (strippedXml === originalXml) {
    return null;
  }

  files[workbookPath] = strToU8(strippedXml);
  return zipSync(files);
}
