import { useRef, useState } from 'react';
import type { ChangeEvent, DragEvent } from 'react';
import { FILE_TYPE_TO_MIME, SUPPORTED_FILE_TYPES } from 'office-wasm';
import type { UploadedFile } from '../atoms';
import { formatBytes } from '../lib/format';

/**
 * Форматы, которые принимает <input type="file">: расширения + MIME-типы.
 * Строится от общего списка форматов пакета office-wasm — при добавлении
 * нового формата здесь ничего менять не нужно.
 */
const ACCEPT = SUPPORTED_FILE_TYPES.flatMap((type) => [`.${type}`, FILE_TYPE_TO_MIME[type]]).join(
  ','
);

interface FileUploaderProps {
  /** Уже выбранный файл (для показа карточки). */
  file: UploadedFile | null;
  /** Блокировка зоны во время конвертации. */
  disabled: boolean;
  /** Кладёт выбранный файл в состояние; бросает Error для неподдерживаемых форматов. */
  onSelectFile: (file: File) => void;
}

/**
 * Зона загрузки файла: drag & drop + выбор через системный диалог.
 * Доступность: скрытый, но фокусируемый <input> внутри <label> — клик по зоне
 * открывает диалог, а с клавиатуры он активируется по Enter/Space.
 */
export default function FileUploader({ file, disabled, onSelectFile }: FileUploaderProps) {
  const [dragging, setDragging] = useState(false);
  const [localError, setLocalError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  // dragenter/dragleave срабатывают и на дочерних элементах зоны,
  // поэтому считаем глубину вложенности вместо простого флага
  const dragDepth = useRef(0);

  const pickFile = (pickedFile: File | null | undefined) => {
    if (!pickedFile || disabled) return;
    setLocalError(null);
    try {
      // onSelectFile (useConverter.selectFile) сам проверяет формат
      // и бросает Error с понятным сообщением для неподдерживаемых файлов
      onSelectFile(pickedFile);
    } catch (error) {
      setLocalError(error instanceof Error ? error.message : String(error));
    }
  };

  const handleDrop = (event: DragEvent<HTMLLabelElement>) => {
    event.preventDefault();
    dragDepth.current = 0;
    setDragging(false);
    const files = event.dataTransfer.files;
    if (files.length > 1) {
      setLocalError('Перетащите только один файл.');
      return;
    }
    pickFile(files[0]);
  };

  return (
    <div>
      <label
        className={`dropzone${dragging ? ' dragging' : ''}`}
        aria-disabled={disabled}
        onDragEnter={(event) => {
          if (disabled) return;
          event.preventDefault();
          dragDepth.current += 1;
          setDragging(true);
        }}
        onDragOver={(event) => {
          // Без preventDefault браузер откроет файл вместо drop-обработчика
          event.preventDefault();
        }}
        onDragLeave={() => {
          dragDepth.current -= 1;
          if (dragDepth.current <= 0) {
            dragDepth.current = 0;
            setDragging(false);
          }
        }}
        onDrop={handleDrop}
      >
        <input
          ref={inputRef}
          className="sr-only"
          type="file"
          accept={ACCEPT}
          disabled={disabled}
          onChange={(event: ChangeEvent<HTMLInputElement>) => {
            pickFile(event.target.files?.[0]);
            // Сбрасываем значение, чтобы повторный выбор того же файла срабатывал
            event.target.value = '';
          }}
        />
        <span className="dropzone-icon" aria-hidden="true">
          📄
        </span>
        <span className="dropzone-title">
          {file ? 'Выбрать другой файл' : 'Перетащите файл сюда'}
        </span>
        <span className="dropzone-hint">
          или нажмите, чтобы выбрать — поддерживаются{' '}
          {SUPPORTED_FILE_TYPES.map((type) => type.toUpperCase()).join(' и ')}
        </span>
      </label>

      {file && (
        <div className="file-chip">
          <span className="file-type-badge">{file.type.toUpperCase()}</span>
          <div className="file-meta-text">
            <div className="file-name">{file.name}</div>
            <div className="file-size">{formatBytes(file.size)}</div>
          </div>
        </div>
      )}

      {localError && (
        <p className="dropzone-error" role="alert">
          {localError}
        </p>
      )}
    </div>
  );
}
