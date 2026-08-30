import { useCallback, useEffect, useRef } from 'react';
import { useAtomValue, useSetAtom } from 'jotai';
import { detectFileType, getConverterClient } from 'office-converter';
import type { WorkerResponse } from 'office-converter';
import {
  officeStateAtom,
  officeInitProgressAtom,
  officeInitErrorAtom,
  fileAtom,
  conversionStateAtom,
  conversionProgressAtom,
  conversionErrorAtom,
  pdfUrlAtom,
  canConvertAtom
} from '../atoms';

// ---------------------------------------------------------------------------
// Адаптер над пакетом office-converter: связывает фреймворк-независимый
// RPC-клиент (Web Worker + протокол postMessage) с Jotai-атомами приложения.
//
// Жизненным циклом worker'а владеет клиент пакета: worker переживает
// размонтирование компонента (StrictMode монтирует эффекты дважды, а повторная
// инициализация WASM стоит секунды и ~240 МБ) и живёт до выгрузки страницы
// или HMR-обновления пакета office-converter.
// ---------------------------------------------------------------------------

const client = getConverterClient();

/**
 * Главный хук приложения: подписывается на события клиента конвертера
 * и переводит их в состояние атомов. Возвращает значения атомов и действия.
 */
export function useConverter() {
  // --- Чтение состояния ------------------------------------------------------
  const officeState = useAtomValue(officeStateAtom);
  const officeInitProgress = useAtomValue(officeInitProgressAtom);
  const officeInitError = useAtomValue(officeInitErrorAtom);
  const file = useAtomValue(fileAtom);
  const conversionState = useAtomValue(conversionStateAtom);
  const conversionProgress = useAtomValue(conversionProgressAtom);
  const conversionError = useAtomValue(conversionErrorAtom);
  const pdfUrl = useAtomValue(pdfUrlAtom);
  const canConvert = useAtomValue(canConvertAtom);

  // --- Запись состояния ------------------------------------------------------
  const setOfficeState = useSetAtom(officeStateAtom);
  const setOfficeInitProgress = useSetAtom(officeInitProgressAtom);
  const setOfficeInitError = useSetAtom(officeInitErrorAtom);
  const setFile = useSetAtom(fileAtom);
  const setConversionState = useSetAtom(conversionStateAtom);
  const setConversionProgress = useSetAtom(conversionProgressAtom);
  const setConversionError = useSetAtom(conversionErrorAtom);
  const setPdfUrl = useSetAtom(pdfUrlAtom);

  // Синхронный guard: атомы обновляются асинхронно, поэтому для защиты
  // от двойного клика держим флаг в ref, а не в состоянии.
  const convertInFlightRef = useRef(false);
  // Последнее сообщение прогресса: библиотека может слать сотни одинаковых
  // событий на одну конвертацию — дубли не пропускаем в атомы (и в React).
  const lastProgressRef = useRef<(WorkerResponse & { type: 'progress' }) | null>(null);

  // --- Разбор сообщений клиента ----------------------------------------------
  const handleMessage = useCallback(
    (data: WorkerResponse) => {
      switch (data.type) {
        case 'progress': {
          const prev = lastProgressRef.current;
          if (
            prev &&
            prev.phase === data.phase &&
            prev.percent === data.percent &&
            prev.message === data.message
          ) {
            break;
          }
          lastProgressRef.current = data;
          // Прогресс приходит и от инициализации, и от конвертации;
          // worker помечает сообщение фазой
          if (data.phase === 'init') {
            setOfficeInitProgress({ percent: data.percent, message: data.message });
          } else {
            setConversionProgress({ percent: data.percent, message: data.message });
          }
          break;
        }

        case 'init-done':
          setOfficeState('ready');
          setOfficeInitProgress(null);
          setOfficeInitError(null);
          break;

        case 'init-error':
          setOfficeState('error');
          setOfficeInitError(data.message);
          break;

        // convert-done / convert-error резолвятся внутри клиента пакета:
        // они только завершают promise из client.convert(), и здесь
        // состояние уже обновляет сам convert()
        default:
          break;
      }
    },
    [setOfficeState, setOfficeInitProgress, setOfficeInitError, setConversionProgress]
  );

  // --- Монтирование: подписываемся и запускаем инициализацию движка ----------
  useEffect(() => {
    client.init();
    return client.subscribe(handleMessage);
  }, [handleMessage]);

  // --- Действия ---------------------------------------------------------------

  /**
   * Проверяет формат и кладёт выбранный файл в состояние.
   * Содержимое не читается: буфер понадобится только при конвертации.
   */
  const selectFile = useCallback(
    (file: File) => {
      const type = detectFileType(file.name);
      if (!type) {
        throw new Error('Неподдерживаемый формат. Выберите файл DOCX или XLSX.');
      }

      // Выбрали новый файл — сбрасываем результат предыдущей конвертации
      if (pdfUrl) URL.revokeObjectURL(pdfUrl);
      setPdfUrl(null);
      setConversionState('idle');
      setConversionError(null);
      setConversionProgress(null);
      lastProgressRef.current = null;

      setFile({ name: file.name, size: file.size, type, file });
    },
    [pdfUrl, setFile, setPdfUrl, setConversionState, setConversionError, setConversionProgress]
  );

  /** Запускает конвертацию текущего файла в PDF. */
  const convert = useCallback(async () => {
    if (!file || conversionState === 'converting' || convertInFlightRef.current) return;

    convertInFlightRef.current = true;
    setConversionState('converting');
    setConversionProgress(null);
    setConversionError(null);
    lastProgressRef.current = null;

    try {
      // Буфер читаем на каждый запуск: client передаёт его в worker
      // по transfer list (без копирования) и забирает владение,
      // поэтому держать буфер между конвертациями нельзя
      const buffer = await file.file.arrayBuffer();
      const pdfBuffer = await client.convert(buffer, file.type, { transfer: true });

      // Старый Blob URL больше не нужен — освобождаем память
      if (pdfUrl) URL.revokeObjectURL(pdfUrl);
      const blob = new Blob([pdfBuffer], { type: 'application/pdf' });
      setPdfUrl(URL.createObjectURL(blob));
      setConversionState('done');
    } catch (error) {
      setConversionError(error instanceof Error ? error.message : String(error));
      setConversionState('error');
    } finally {
      convertInFlightRef.current = false;
    }
  }, [
    file,
    conversionState,
    pdfUrl,
    setConversionState,
    setConversionProgress,
    setConversionError,
    setPdfUrl
  ]);

  /** Повторная попытка инициализации движка после ошибки. */
  const retryInit = useCallback(() => {
    setOfficeState('initializing');
    setOfficeInitError(null);
    setOfficeInitProgress(null);
    lastProgressRef.current = null;
    client.init();
  }, [setOfficeState, setOfficeInitError, setOfficeInitProgress]);

  return {
    // Состояние движка
    officeState,
    officeInitProgress,
    officeInitError,
    // Выбранный файл и конвертация
    file,
    conversionState,
    conversionProgress,
    conversionError,
    pdfUrl,
    canConvert,
    // Действия
    selectFile,
    convert,
    retryInit
  };
}
