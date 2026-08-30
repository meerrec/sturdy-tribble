import { describe, expect, it } from 'vitest';
import { formatBytes } from './format';

describe('formatBytes', () => {
  it('форматирует байты, кило- и мегабайты', () => {
    expect(formatBytes(512)).toBe('512 Б');
    expect(formatBytes(2048)).toBe('2.0 КБ');
    expect(formatBytes(10 * 1024 * 1024)).toBe('10 МБ');
  });

  it('не показывает десятые доли у круглых крупных значений', () => {
    expect(formatBytes(1.5 * 1024 * 1024)).toBe('1.5 МБ');
  });

  it('возвращает пустую строку для нечисловых значений', () => {
    expect(formatBytes(Number.NaN)).toBe('');
    expect(formatBytes(Number.POSITIVE_INFINITY)).toBe('');
  });
});
