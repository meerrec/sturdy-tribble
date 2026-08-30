import FileUploader from './components/FileUploader';
import ConvertButton from './components/ConvertButton';
import PDFViewer from './components/PDFViewer';
import { useConverter } from './hooks/useConverter';
import type { OfficeState, ProgressInfo } from './atoms';

/** Человекочитаемый размер файла (для карточки выбранного файла). */
function formatBytes(bytes: number): string {
  const units = ['Б', 'КБ', 'МБ', 'ГБ'];
  let value = bytes;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit += 1;
  }
  return `${value.toFixed(value >= 10 || unit === 0 ? 0 : 1)} ${units[unit]}`;
}

interface OfficeStatusProps {
  state: OfficeState;
  progress: ProgressInfo | null;
  onRetry: () => void;
}

/** Пилюля статуса офисного движка в шапке. */
function OfficeStatus({ state, progress, onRetry }: OfficeStatusProps) {
  if (state === 'ready') {
    return <div className="pill pill-ready">Движок готов</div>;
  }
  if (state === 'error') {
    return (
      <div className="pill pill-error">
        Ошибка инициализации движка
        <button type="button" className="link-btn" onClick={onRetry}>
          повторить
        </button>
      </div>
    );
  }
  const percent = progress ? ` ${progress.percent}%` : '';
  return (
    <div className="pill pill-busy">
      <span className="spinner" aria-hidden="true" />
      Загрузка LibreOffice…{percent}
    </div>
  );
}

interface StatusBlockProps {
  progress: ProgressInfo | null;
  label: string;
  note?: string;
}

/**
 * Блок хода выполнения: прогресс-бар с процентами, если библиотека их
 * сообщает, иначе — анимированный «неопределённый» прогресс.
 */
function StatusBlock({ progress, label, note }: StatusBlockProps) {
  const percent =
    progress != null && Number.isFinite(progress.percent)
      ? progress.percent
      : null;
  return (
    <div className="status">
      <div className="status-label">
        {label}
        {percent != null ? ` ${percent}%` : ''}
      </div>
      <div className={`progress${percent != null ? '' : ' indeterminate'}`}>
        <div
          className="progress-fill"
          style={percent != null ? { width: `${percent}%` } : undefined}
        />
      </div>
      {progress?.message && <div className="status-message">{progress.message}</div>}
      {note && <div className="status-note">{note}</div>}
    </div>
  );
}

export default function App() {
  const {
    officeState,
    officeInitProgress,
    officeInitError,
    file,
    conversionState,
    conversionProgress,
    conversionError,
    pdfUrl,
    canConvert,
    selectFile,
    convert,
    retryInit,
  } = useConverter();

  const isConverting = conversionState === 'converting';

  return (
    <div className="page">
      <header className="app-header">
        <h1>Office → PDF</h1>
        <p className="app-subtitle">
          Конвертация DOCX и XLSX в PDF прямо в браузере — файлы не покидают ваш
          компьютер
        </p>
        <OfficeStatus
          state={officeState}
          progress={officeInitProgress}
          onRetry={retryInit}
        />
      </header>

      <main className="content">
        <section className="card" aria-label="Загрузка файла">
          <FileUploader
            file={file}
            disabled={isConverting}
            onSelectFile={selectFile}
          />
        </section>

        {file && (
          <section className="card actions-card">
            <div className="file-meta">
              <span className="file-type-badge">{file.type.toUpperCase()}</span>
              <div className="file-meta-text">
                <div className="file-name">{file.name}</div>
                <div className="file-size">{formatBytes(file.size)}</div>
              </div>
            </div>
            <ConvertButton
              disabled={!canConvert}
              converting={isConverting}
              onConvert={convert}
            />
          </section>
        )}

        {/* aria-live: скринридеры озвучивают смену статуса */}
        <div aria-live="polite">
          {officeState === 'initializing' && (
            <section className="card">
              <StatusBlock
                progress={officeInitProgress}
                label="Инициализация LibreOffice…"
                note="Первый запуск скачивает ≈240 МБ данных движка; при следующих запусках используется кэш браузера."
              />
            </section>
          )}

          {isConverting && (
            <section className="card">
              <StatusBlock
                progress={conversionProgress}
                label="Конвертация документа…"
              />
            </section>
          )}

          {officeState === 'error' && (
            <section className="card error-banner" role="alert">
              <strong>Не удалось инициализировать LibreOffice WASM.</strong>
              <div>{officeInitError || 'Неизвестная ошибка'}</div>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={retryInit}
              >
                Повторить инициализацию
              </button>
            </section>
          )}

          {conversionState === 'error' && (
            <section className="card error-banner" role="alert">
              <strong>Конвертация не удалась.</strong>
              <div>{conversionError || 'Неизвестная ошибка'}</div>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={convert}
                disabled={!canConvert}
              >
                Попробовать снова
              </button>
            </section>
          )}
        </div>

        {pdfUrl && (
          <section className="card viewer-card">
            <PDFViewer pdfUrl={pdfUrl} sourceFileName={file?.name} />
          </section>
        )}
      </main>

      <footer className="app-footer">
        <p>
          Движок — LibreOffice в WebAssembly (<code>@matbee/libreoffice-converter</code>).
          Для работы требуется браузер с поддержкой SharedArrayBuffer и
          заголовками COOP/COEP (Chrome, Edge, Firefox).
        </p>
      </footer>
    </div>
  );
}
