import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

// Мок движка: настоящий WASM в юнит-тестах не поднимаем — проверяем только,
// КАК convertDocumentToPdf готовит данные и опции перед вызовом convert.
const { convertMock } = vi.hoisted(() => ({
  convertMock: vi.fn().mockResolvedValue({ data: new Uint8Array([1, 2, 3]) })
}));

vi.mock('@matbee/libreoffice-converter/browser', () => ({
  WorkerBrowserConverter: vi.fn().mockImplementation(() => ({
    initialize: vi.fn().mockResolvedValue(undefined),
    convert: convertMock,
    destroy: vi.fn().mockResolvedValue(undefined)
  })),
  createWasmPaths: vi.fn(() => ({}))
}));

import {
  convertDocumentToPdf,
  disposeOffice,
  XLSX_PDF_FILTER_OPTIONS,
  type SupportedFileType
} from './index';

/** Минимальный xlsx со скрытым листом — для проверки препроцессинга. */
function buildXlsxWithHiddenSheet(): Uint8Array {
  const files = {
    '[Content_Types].xml': strToU8('<Types/>'),
    '_rels/.rels': strToU8('<Relationships/>'),
    'xl/workbook.xml': strToU8(`
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Видимый" sheetId="1" r:id="rId1"/>
    <sheet name="Скрытый" sheetId="2" state="hidden" r:id="rId2"/>
  </sheets>
</workbook>`),
    'xl/worksheets/sheet1.xml': strToU8('<worksheet/>')
  };
  return zipSync(files);
}

beforeEach(() => {
  convertMock.mockClear();
});

afterEach(async () => {
  // Сбрасываем кэш initOffice между тестами: каждый тест создаёт
  // «свежий» конвертер и convertMock видит ровно один вызов.
  await disposeOffice();
});

describe('convertDocumentToPdf', () => {
  it('для xlsx передаёт SinglePageSheets и вырезает скрытые листы', async () => {
    await convertDocumentToPdf(buildXlsxWithHiddenSheet(), 'xlsx');

    expect(convertMock).toHaveBeenCalledTimes(1);
    const [input, options, filename] = convertMock.mock.calls[0];

    expect(options).toEqual({ outputFormat: 'pdf', filterOptions: XLSX_PDF_FILTER_OPTIONS });
    expect(filename).toBe('document.xlsx');

    const workbookXml = strFromU8(unzipSync(input as Uint8Array)['xl/workbook.xml']);
    expect(workbookXml).toContain('Видимый');
    expect(workbookXml).not.toContain('Скрытый');
  });

  it('для docx опции SinglePageSheets не передаёт и файл не трогает', async () => {
    const docx = strToU8('не настоящий docx — для теста важен только вызов');

    await convertDocumentToPdf(docx, 'docx');

    expect(convertMock).toHaveBeenCalledTimes(1);
    const [input, options, filename] = convertMock.mock.calls[0];

    expect(options).toEqual({ outputFormat: 'pdf' });
    expect(filename).toBe('document.docx');
    expect(input).toBe(docx); // тот же буфер — препроцессинг не применялся
  });

  it('бросает TypeError для неподдерживаемого типа', async () => {
    await expect(
      convertDocumentToPdf(new Uint8Array(), 'pptx' as SupportedFileType)
    ).rejects.toThrow(TypeError);
    expect(convertMock).not.toHaveBeenCalled();
  });
});
