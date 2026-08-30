/**
 * e2e-тест: реальный движок LibreOffice WASM принимает zip от fflate и после
 * удаления области печати экспортирует весь лист, а не только область печати.
 *
 * Исключён из обычного прогона (см. vitest.config.ts): поднимает движок WASM
 * (~240 МБ, десятки секунд на первую инициализацию). Запуск:
 *   pnpm --filter office-wasm test:e2e
 */
import { convertDocument } from '@matbee/libreoffice-converter';
import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { stripPrintRangesFromXlsx } from './strip-print-ranges';

const wasmOptions = {
  wasmPath: new URL('./node_modules/@matbee/libreoffice-converter/wasm', import.meta.url).pathname
};

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

describe('e2e: конвертация xlsx без областей печати', () => {
  it('весь лист (1500 строк) попадает в PDF после stripPrintRangesFromXlsx', async () => {
    const source = buildBigXlsx(1500);

    // Контроль: с областью печати A1:E5 PDF получается одностраничным.
    const truncatedPdf = await convertDocument(source, { outputFormat: 'pdf' }, wasmOptions);
    expect(countPdfPages(truncatedPdf.data)).toBe(1);

    // Основная проверка: после удаления области печати страниц должно быть много.
    const stripped = stripPrintRangesFromXlsx(source)!;
    const pdf = await convertDocument(stripped, { outputFormat: 'pdf' }, wasmOptions);
    const pageCount = countPdfPages(pdf.data);
    expect(pageCount).toBeGreaterThan(50);
  }, 120_000);
});

/** Грубый подсчёт страниц PDF по маркерам /Type /Page. */
function countPdfPages(pdf: Uint8Array): number {
  const text = new TextDecoder('latin1').decode(pdf);
  return (text.match(/\/Type\s*\/Page\b/g) ?? []).length;
}
