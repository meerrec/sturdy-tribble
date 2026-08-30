interface ConvertButtonProps {
  disabled: boolean;
  converting: boolean;
  onConvert: () => void;
}

/**
 * Кнопка запуска конвертации.
 * disabled приходит из атома canConvertAtom: нужны выбранный файл,
 * готовый движок и отсутствие уже идущей конвертации.
 */
export default function ConvertButton({ disabled, converting, onConvert }: ConvertButtonProps) {
  return (
    <button
      type="button"
      className="btn"
      disabled={disabled}
      onClick={onConvert}
    >
      {converting ? (
        <>
          <span className="spinner" aria-hidden="true" />
          Конвертация…
        </>
      ) : (
        'Конвертировать в PDF'
      )}
    </button>
  );
}
