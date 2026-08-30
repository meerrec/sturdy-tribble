import { useState } from 'react';

interface PDFViewerProps {
  /** Blob URL готового PDF. */
  pdfUrl: string;
  /** Имя исходного файла — из него получается имя для скачивания. */
  sourceFileName?: string;
}

/**
 * Отображение результата: <iframe> с Blob URL готового PDF.
 * PDF рендерится встроенным просмотрщиком браузера (Chromium/Firefox).
 *
 * Ключевые моменты:
 *  - компонент рендерится с key={pdfUrl} (см. App): при новом URL он
 *    пересоздаётся целиком, и просмотрщик не показывает предыдущий
 *    документ из своего кэша, а состояние загрузки сбрасывается само;
 *  - Blob URL живёт в атоме pdfUrlAtom и отзывается (revokeObjectURL)
 *    при замене файла или результата — см. useConverter.
 */
export default function PDFViewer({ pdfUrl, sourceFileName }: PDFViewerProps) {
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState(false);

  if (!pdfUrl) {
    return <div className="viewer-placeholder">Результат конвертации появится здесь</div>;
  }

  // Имя для скачивания: «отчёт.xlsx» → «отчёт.pdf»
  const downloadName = sourceFileName
    ? `${sourceFileName.replace(/\.[^.]+$/, '')}.pdf`
    : 'document.pdf';

  return (
    <div className="viewer">
      <div className="viewer-toolbar">
        <span className="viewer-title">Предпросмотр PDF</span>
        {/* Скачивание прямо из Blob URL — файл не покидает браузер */}
        <a className="btn btn-secondary" href={pdfUrl} download={downloadName}>
          Скачать PDF
        </a>
      </div>
      <div className="viewer-frame-wrap">
        {loading && !loadError && (
          <div className="viewer-loading">
            <span className="spinner" aria-hidden="true" />
            Загрузка PDF…
          </div>
        )}
        {loadError && (
          <div className="viewer-error" role="alert">
            Не удалось отобразить предпросмотр PDF.
            <a className="btn btn-secondary" href={pdfUrl} download={downloadName}>
              Скачать PDF
            </a>
          </div>
        )}
        {!loadError && (
          <iframe
            title="Предпросмотр сконвертированного PDF-документа"
            src={pdfUrl}
            onLoad={() => {
              setLoading(false);
              setLoadError(false);
            }}
            onError={() => {
              setLoading(false);
              setLoadError(true);
            }}
          />
        )}
      </div>
    </div>
  );
}
