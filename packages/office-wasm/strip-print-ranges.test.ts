import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import {
  prepareXlsxForPdf,
  stripHiddenSheetsFromXlsx,
  stripPrintRangesFromXlsx
} from './strip-print-ranges';

/**
 * Собирает минимальный xlsx (zip) для тестов.
 * Лист с данными — заглушка: проверяем только workbook.xml, где живут области печати.
 */
function buildXlsx(workbookXml: string): Uint8Array {
  const files = {
    '[Content_Types].xml': strToU8('<Types/>'),
    '_rels/.rels': strToU8('<Relationships/>'),
    'xl/workbook.xml': strToU8(workbookXml),
    'xl/worksheets/sheet1.xml': strToU8('<worksheet/>')
  };
  return zipSync(files);
}

const workbookWithDefinedNames = `
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">'Sheet1'!$A$1:$E$10</definedName>
    <definedName name="_xlnm.Print_Titles" localSheetId="0">'Sheet1'!$1:$2</definedName>
    <definedName name="MyRange">Sheet1!$A$1:$A$5</definedName>
  </definedNames>
</workbook>`;

describe('stripPrintRangesFromXlsx', () => {
  it('удаляет область печати, сохраняя остальные definedName', () => {
    const result = stripPrintRangesFromXlsx(buildXlsx(workbookWithDefinedNames));

    expect(result).not.toBeNull();
    const xml = strFromU8(unzipSync(result!)['xl/workbook.xml']);
    expect(xml).not.toContain('_xlnm.Print_Area');
    expect(xml).not.toContain("'Sheet1'!$A$1:$E$10");
    expect(xml).toContain('_xlnm.Print_Titles');
    expect(xml).toContain('MyRange');
  });

  it('не трогает файл без областей печати', () => {
    const workbookWithoutPrintAreas = `
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <definedNames>
    <definedName name="MyRange">Sheet1!$A$1:$A$5</definedName>
  </definedNames>
</workbook>`;

    expect(stripPrintRangesFromXlsx(buildXlsx(workbookWithoutPrintAreas))).toBeNull();
  });

  it('убирает пустую обёртку definedNames', () => {
    const workbookOnlyPrintArea = `
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">'Sheet1'!$A$1:$E$10</definedName>
  </definedNames>
</workbook>`;

    const xml = strFromU8(
      unzipSync(stripPrintRangesFromXlsx(buildXlsx(workbookOnlyPrintArea))!)['xl/workbook.xml']
    );
    expect(xml).not.toContain('definedNames');
  });

  it('работает с префиксами пространства имён и одинарными кавычками', () => {
    const namespacedWorkbook = `
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:definedNames>
    <x:definedName localSheetId="0" name='_xlnm.Print_Area'>'Sheet1'!$A$1:$E$10</x:definedName>
  </x:definedNames>
</x:workbook>`;

    const result = stripPrintRangesFromXlsx(buildXlsx(namespacedWorkbook));
    expect(result).not.toBeNull();
    expect(strFromU8(unzipSync(result!)['xl/workbook.xml'])).not.toContain('Print_Area');
  });

  it('удаляет все области печати, если листов несколько', () => {
    const multiSheetWorkbook = `
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">'Sheet1'!$A$1:$E$10</definedName>
    <definedName name="_xlnm.Print_Area" localSheetId="1">'Sheet2'!$B$2:$D$20</definedName>
  </definedNames>
</workbook>`;

    const xml = strFromU8(
      unzipSync(stripPrintRangesFromXlsx(buildXlsx(multiSheetWorkbook))!)['xl/workbook.xml']
    );
    expect(xml).not.toContain('_xlnm.Print_Area');
  });

  it('находит workbook.xml без учёта регистра имени', () => {
    const workbookXml = `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">'Sheet1'!$A$1:$E$10</definedName>
  </definedNames>
</workbook>`;
    const files = {
      '[Content_Types].xml': strToU8('<Types/>'),
      '_rels/.rels': strToU8('<Relationships/>'),
      'xl/Workbook.xml': strToU8(workbookXml),
      'xl/worksheets/sheet1.xml': strToU8('<worksheet/>')
    };
    const result = stripPrintRangesFromXlsx(zipSync(files));
    expect(result).not.toBeNull();
    expect(strFromU8(unzipSync(result!)['xl/Workbook.xml'])).not.toContain('Print_Area');
  });

  it('возвращает null для не-zip данных', () => {
    expect(stripPrintRangesFromXlsx(strToU8('это не zip-архив'))).toBeNull();
  });

  it('результат остаётся валидным zip со всеми исходными файлами', () => {
    const result = stripPrintRangesFromXlsx(buildXlsx(workbookWithDefinedNames))!;
    const entries = unzipSync(result);

    expect(Object.keys(entries).sort()).toEqual(
      ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml'].sort()
    );
    expect(strFromU8(entries['xl/worksheets/sheet1.xml'])).toBe('<worksheet/>');
  });
});

describe('stripHiddenSheetsFromXlsx', () => {
  it('удаляет скрытые листы, сохраняя видимые', () => {
    const workbookWithHiddenSheets = `
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Видимый" sheetId="1" r:id="rId1"/>
    <sheet name="Скрытый" sheetId="2" state="hidden" r:id="rId2"/>
    <sheet name="ОченьСкрытый" sheetId="3" state="veryHidden" r:id="rId3"/>
  </sheets>
</workbook>`;

    const xml = strFromU8(
      unzipSync(stripHiddenSheetsFromXlsx(buildXlsx(workbookWithHiddenSheets))!)['xl/workbook.xml']
    );

    expect(xml).toContain('Видимый');
    expect(xml).not.toContain('Скрытый');
    expect(xml).not.toContain('veryHidden');
    expect(xml).not.toContain('r:id="rId2"');
  });

  it('не трогает файл без скрытых листов', () => {
    const workbookWithoutHiddenSheets = `
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Лист1" sheetId="1" r:id="rId1"/>
    <sheet name="Лист2" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`;

    expect(stripHiddenSheetsFromXlsx(buildXlsx(workbookWithoutHiddenSheets))).toBeNull();
  });

  it('работает с префиксами пространства имён, любым порядком атрибутов и одинарными кавычками', () => {
    const namespacedWorkbook = `
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:sheets>
    <x:sheet r:id="rId1" state='hidden' name="Скрытый" sheetId="2"/>
    <x:sheet name="Видимый" sheetId="1" r:id="rId2"/>
  </x:sheets>
</x:workbook>`;

    const xml = strFromU8(
      unzipSync(stripHiddenSheetsFromXlsx(buildXlsx(namespacedWorkbook))!)['xl/workbook.xml']
    );

    expect(xml).toContain('Видимый');
    expect(xml).not.toContain('Скрытый');
  });

  it('не удаляет последний лист, даже если он скрыт', () => {
    const workbookWithOnlyHiddenSheet = `
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="ЕдинственныйСкрытый" sheetId="1" state="hidden" r:id="rId1"/>
  </sheets>
</workbook>`;

    expect(stripHiddenSheetsFromXlsx(buildXlsx(workbookWithOnlyHiddenSheet))).toBeNull();
  });
});

describe('prepareXlsxForPdf', () => {
  it('за один проход удаляет области печати и скрытые листы', () => {
    const workbookWithBoth = `
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Видимый" sheetId="1" r:id="rId1"/>
    <sheet name="Скрытый" sheetId="2" state="hidden" r:id="rId2"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">'Видимый'!$A$1:$E$10</definedName>
  </definedNames>
</workbook>`;

    const xml = strFromU8(
      unzipSync(prepareXlsxForPdf(buildXlsx(workbookWithBoth))!)['xl/workbook.xml']
    );

    expect(xml).toContain('Видимый');
    expect(xml).not.toContain('Скрытый');
    expect(xml).not.toContain('Print_Area');
    expect(xml).not.toContain('definedNames');
  });

  it('возвращает null, если менять нечего', () => {
    const plainWorkbook = `
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Лист1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

    expect(prepareXlsxForPdf(buildXlsx(plainWorkbook))).toBeNull();
  });

  it('возвращает null для не-zip данных', () => {
    expect(prepareXlsxForPdf(strToU8('это не zip-архив'))).toBeNull();
  });
});
