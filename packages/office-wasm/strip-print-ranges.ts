/**
 * Препроцессинг XLSX перед конвертацией в PDF.
 *
 * 1. Удаление областей печати.
 *    При обычной разбивке на страницы LibreOffice Calc экспортирует лист
 *    строго в пределах области печати (definedName `_xlnm.Print_Area`,
 *    которую Excel сохраняет в xl/workbook.xml и выставляет молча, например,
 *    при «Set Print Area» или «Fit to page») — всё содержимое за её
 *    пределами из PDF пропадает. Конвертация XLSX идёт с SinglePageSheets
 *    (см. index.ts), при которой движок экспортирует весь used range листа
 *    и область печати игнорирует, так что вырезка здесь — страховка
 *    на случай конвертации без этой опции: семантику «в PDF попадает весь
 *    лист» не должен менять даже такой файл.
 *
 *    `_xlnm.Print_Titles` (повторяющиеся на каждой странице строки/столбцы)
 *    НЕ удаляются: они не обрезают данные, а только дублируют заголовки —
 *    поведение сохраняем как задумано автором файла.
 *
 * 2. Удаление скрытых листов.
 *    При SinglePageSheets движок экспортирует ВСЕ листы, включая скрытые
 *    (`state="hidden"` / `state="veryHidden"`). При обычной печати и экспорте
 *    PDF скрытые листы не печатаются, поэтому вырезаем их здесь, из
 *    workbook.xml — единственного места, где перечисляются листы книги.
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
 * Элемент `<sheet … state="hidden" … />` (или парный вариант) — скрытый лист.
 * Порядок и кавычки атрибутов не важны — совпадение идёт по атрибуту `state`.
 * `veryHidden` — скрытый лист, который нельзя показать из интерфейса Excel;
 * для экспорта он скрыт точно так же.
 */
const HIDDEN_SHEET_RE = new RegExp(
  `<(?:[\\w-]+:)?sheet\\b[^>]*\\bstate\\s*=\\s*["'](?:hidden|veryHidden)["'][^>]*\\/?>` +
    '|' +
    `<(?:[\\w-]+:)?sheet\\b[^>]*\\bstate\\s*=\\s*["'](?:hidden|veryHidden)["'][^>]*>[\\s\\S]*?<\\/(?:[\\w-]+:)?sheet\\s*>`,
  'g'
);

/** Открывающий тег `<sheet …>` — для проверки, что в книге остался хотя бы один лист. */
const ANY_SHEET_RE = /<(?:[\w-]+:)?sheet\b[^>]*>/;

/** Преобразование XML-содержимого workbook.xml. */
type WorkbookTransformer = (xml: string) => string;

/**
 * Применяет преобразование к xl/workbook.xml внутри xlsx-архива.
 *
 * @param buffer — содержимое .xlsx (zip-архив).
 * @param transformer — функция, возвращающая НОВЫЙ xml (или исходный, если менять нечего).
 * @returns Новый XLSX-буфер, либо `null`, если содержимое не изменилось
 *          (или файл — не zip-архив): тогда конвертировать нужно исходник.
 */
function transformWorkbookXml(
  buffer: Uint8Array,
  transformer: WorkbookTransformer
): Uint8Array | null {
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
  const transformedXml = transformer(originalXml);

  // Быстрый путь: менять было нечего — исходный буфер не трогаем.
  if (transformedXml === originalXml) {
    return null;
  }

  files[workbookPath] = strToU8(transformedXml);
  return zipSync(files);
}

/**
 * Возвращает копию XLSX-файла без областей печати.
 *
 * @param buffer — содержимое .xlsx (zip-архив).
 * @returns Новый XLSX-буфер, либо `null`, если областей печати в файле нет
 *          (или файл — не zip-архив): тогда конвертировать нужно исходник.
 */
export function stripPrintRangesFromXlsx(buffer: Uint8Array): Uint8Array | null {
  return transformWorkbookXml(buffer, (xml) =>
    xml.replace(PRINT_AREA_RE, '').replace(EMPTY_DEFINED_NAMES_RE, '')
  );
}

/**
 * Возвращает копию XLSX-файла без скрытых листов.
 *
 * Если после вырезания в книге не остаётся ни одного листа (все листы скрыты),
 * файл не меняется: отдавать книгу вообще без листов хуже, чем книгу
 * со скрытыми листами — пусть движок сам решит, что с ней делать.
 *
 * @param buffer — содержимое .xlsx (zip-архив).
 * @returns Новый XLSX-буфер, либо `null`, если скрытых листов в файле нет.
 */
export function stripHiddenSheetsFromXlsx(buffer: Uint8Array): Uint8Array | null {
  return transformWorkbookXml(buffer, (xml) => {
    let hiddenCount = 0;
    const withoutHidden = xml.replace(HIDDEN_SHEET_RE, () => {
      hiddenCount += 1;
      return '';
    });
    if (hiddenCount === 0 || !ANY_SHEET_RE.test(withoutHidden)) {
      return xml;
    }
    return withoutHidden;
  });
}

/**
 * Полная подготовка XLSX к конвертации в PDF: один проход по архиву.
 *
 * Объединяет {@link stripPrintRangesFromXlsx} и {@link stripHiddenSheetsFromXlsx}:
 * области печати вырезаются как страховка, а скрытые листы иначе попадали бы
 * в PDF при SinglePageSheets (см. index.ts).
 *
 * @param buffer — содержимое .xlsx (zip-архив).
 * @returns Новый XLSX-буфер, либо `null`, если менять нечего:
 *          тогда конвертировать нужно исходник.
 */
export function prepareXlsxForPdf(buffer: Uint8Array): Uint8Array | null {
  return transformWorkbookXml(buffer, (xml) => {
    const withoutPrintAreas = xml.replace(PRINT_AREA_RE, '').replace(EMPTY_DEFINED_NAMES_RE, '');

    let hiddenCount = 0;
    const withoutHidden = withoutPrintAreas.replace(HIDDEN_SHEET_RE, () => {
      hiddenCount += 1;
      return '';
    });
    if (hiddenCount > 0 && ANY_SHEET_RE.test(withoutHidden)) {
      return withoutHidden;
    }
    return withoutPrintAreas;
  });
}
