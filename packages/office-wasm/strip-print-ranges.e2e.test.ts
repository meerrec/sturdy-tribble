/**
 * e2e-тесты: реальный движок LibreOffice WASM принимает zip от fflate и после
 * препроцессинга (prepareXlsxForPdf) экспортирует в PDF весь лист — без
 * обрезания областью печати и без скрытых листов.
 *
 * Конвертация XLSX идёт с SinglePageSheets: каждый лист — одна страница,
 * размер которой движок подгоняет под содержимое (весь used range).
 * Область печати при этом движком игнорируется, поэтому «весь лист попал
 * в PDF» проверяется размером страницы: он должен соответствовать всем
 * колонкам и строкам листа, а не области печати.
 *
 * Исключены из обычного прогона (см. vitest.config.ts): поднимают движок WASM
 * (~240 МБ, десятки секунд на первую инициализацию). Запуск:
 *   pnpm --filter office-wasm test:e2e
 */
import { convertDocument } from '@matbee/libreoffice-converter';
import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { XLSX_PDF_FILTER_OPTIONS } from './index';
import { prepareXlsxForPdf } from './strip-print-ranges';

const wasmOptions = {
  wasmPath: new URL('./node_modules/@matbee/libreoffice-converter/wasm', import.meta.url).pathname
};

/** Те же опции конвертации, что использует convertDocumentToPdf для XLSX. */
const pdfOptions = { outputFormat: 'pdf' as const, filterOptions: XLSX_PDF_FILTER_OPTIONS };

/**
 * Собирает xlsx c 26 колонками и заданным числом строк.
 * На 1500 строках лист при SinglePageSheets превысил бы максимальный размер
 * PDF-страницы (14 400 pt), поэтому используется 300 строк: ~4500 pt высоты.
 */
function buildBigXlsx(rows: number): Uint8Array {
  const cells = (r: number) =>
    'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
      .split('')
      .map((col, c) => `<c r="${col}${r}" t="inlineStr"><is><t>R${r}C${c + 1}</t></is></c>`)
      .join('');
  const rowsXml = Array.from(
    { length: rows },
    (_, i) => `<row r="${i + 1}">${cells(i + 1)}</row>`
  ).join('');
  const files = {
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      </Types>`
    ),
    '_rels/.rels': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
      </Relationships>`
    ),
    'xl/workbook.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
        <definedNames>
          <definedName name="_xlnm.Print_Area" localSheetId="0">'Sheet1'!$A$1:$E$5</definedName>
        </definedNames>
      </workbook>`
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      </Relationships>`
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rowsXml}</sheetData></worksheet>`
    )
  };
  return zipSync(files);
}

/** Минимальная книга из двух листов, один из которых скрыт. */
function buildXlsxWithHiddenSheet(): Uint8Array {
  const worksheet = (cell: string) =>
    strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${cell}</t></is></c></row></sheetData>
      </worksheet>`
    );
  const files = {
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      </Types>`
    ),
    '_rels/.rels': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
      </Relationships>`
    ),
    'xl/workbook.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
          <sheet name="Видимый" sheetId="1" r:id="rId1"/>
          <sheet name="Скрытый" sheetId="2" state="hidden" r:id="rId2"/>
        </sheets>
      </workbook>`
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
      </Relationships>`
    ),
    'xl/worksheets/sheet1.xml': worksheet('Видимый лист'),
    'xl/worksheets/sheet2.xml': worksheet('Скрытый лист')
  };
  return zipSync(files);
}

describe('e2e: конвертация xlsx с SinglePageSheets', () => {
  it('область печати не обрезает лист: в PDF попадает весь used range', async () => {
    const source = buildBigXlsx(300);

    // Движок при SinglePageSheets экспортирует весь used range листа,
    // игнорируя область печати A1:E5: страница размером со ВСЕ 26 колонок
    // и все 300 строк (~1250×3840 pt), а не с 5×5 ячеек (~240×65 pt).
    const truncatedPdf = await convertDocument(source, pdfOptions, wasmOptions);
    expect(countPdfPages(truncatedPdf.data)).toBe(1);
    const truncatedBox = mediaBox(truncatedPdf.data);
    expect(truncatedBox).not.toBeNull();
    expect(truncatedBox![0]).toBeGreaterThan(1000); // 26 колонок поместились
    expect(truncatedBox![1]).toBeGreaterThan(3000); // все 300 строк поместились

    // Вырезка областей печати (страховка на случай конвертации без
    // SinglePageSheets) не ломает конвертацию и не меняет полноту листа.
    const prepared = prepareXlsxForPdf(source)!;
    const pdf = await convertDocument(prepared, pdfOptions, wasmOptions);
    expect(countPdfPages(pdf.data)).toBe(1);
    expect(mediaBox(pdf.data)).toEqual(truncatedBox);
  }, 120_000);

  it('скрытые листы вырезаются и не попадают в PDF', async () => {
    const source = buildXlsxWithHiddenSheet();

    // Контроль: движок при SinglePageSheets экспортирует и скрытый лист —
    // без препроцессинга страниц было бы две.
    const withHidden = await convertDocument(source, pdfOptions, wasmOptions);
    expect(countPdfPages(withHidden.data)).toBe(2);

    // После препроцессинга скрытый лист удалён из книги — одна страница.
    const prepared = prepareXlsxForPdf(source)!;
    const pdf = await convertDocument(prepared, pdfOptions, wasmOptions);
    expect(countPdfPages(pdf.data)).toBe(1);
  }, 120_000);
});

/** Грубый подсчёт страниц PDF по маркерам /Type /Page. */
function countPdfPages(pdf: Uint8Array): number {
  const text = new TextDecoder('latin1').decode(pdf);
  return (text.match(/\/Type\s*\/Page\b/g) ?? []).length;
}

/** Ширина и высота первой страницы PDF из MediaBox, в пунктах. */
function mediaBox(pdf: Uint8Array): [number, number] | null {
  const text = new TextDecoder('latin1').decode(pdf);
  const match = text.match(/\/MediaBox\s*\[\s*[\d.-]+\s+[\d.-]+\s+([\d.]+)\s+([\d.]+)\s*\]/);
  return match ? [Number(match[1]), Number(match[2])] : null;
}
