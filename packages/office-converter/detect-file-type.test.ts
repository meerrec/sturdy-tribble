import { describe, expect, it } from 'vitest';
import { detectFileType } from './detect-file-type';

describe('detectFileType', () => {
  it('определяет docx и xlsx без учёта регистра', () => {
    expect(detectFileType('отчёт.DOCX')).toBe('docx');
    expect(detectFileType('таблица.xlsx')).toBe('xlsx');
  });

  it('берёт последнее расширение для имён с несколькими точками', () => {
    expect(detectFileType('archive.tar.docx')).toBe('docx');
  });

  it('возвращает null для неподдерживаемых расширений', () => {
    expect(detectFileType('документ.pdf')).toBeNull();
    expect(detectFileType('архив.zip')).toBeNull();
    expect(detectFileType('картинка.png')).toBeNull();
  });

  it('возвращает null для имени без расширения', () => {
    expect(detectFileType('README')).toBeNull();
    expect(detectFileType('file.')).toBeNull();
  });
});
