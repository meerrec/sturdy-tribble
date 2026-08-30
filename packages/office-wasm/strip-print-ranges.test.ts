import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { stripPrintRangesFromXlsx } from './strip-print-ranges';

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
