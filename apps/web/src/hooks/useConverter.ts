import { useCallback, useEffect } from 'react';
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
  canConvertAtom,
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

let convertInFlight = false; // синхронный guard: атомы обновляются асинхронно

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

  // --- Разбор сообщений клиента ----------------------------------------------
  const handleMessage = useCallback(
    (data: WorkerResponse) => {
      switch (data.type) {
        case 'progress':
          // Прогресс приходит и от инициализации, и от конвертации;
          // worker помечает сообщение фазой
          if (data.phase === 'init') {
            setOfficeInitProgress({ percent: data.percent, message: data.message });
          } else {
            setConversionProgress({ percent: data.percent, message: data.message });
          }
          break;

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
    [
      setOfficeState,
      setOfficeInitProgress,
      setOfficeInitError,
      setConversionProgress,
    ],
  );

  // --- Монтирование: подписываемся и запускаем инициализацию движка ----------
  useEffect(() => {
    client.init();
    return client.subscribe(handleMessage);
  }, [handleMessage]);

  // --- Действия ---------------------------------------------------------------

  /** Читает выбранный файл в ArrayBuffer и кладёт его в состояние. */
  const selectFile = useCallback(
    async (file: File) => {
      const type = detectFileType(file);
      if (!type) {
        throw new Error('Неподдерживаемый формат. Выберите файл DOCX или XLSX.');
      }

      const buffer = await file.arrayBuffer();

      // Выбрали новый файл — сбрасываем результат предыдущей конвертации
      if (pdfUrl) URL.revokeObjectURL(pdfUrl);
      setPdfUrl(null);
      setConversionState('idle');
      setConversionError(null);
      setConversionProgress(null);

      setFile({ name: file.name, size: file.size, type, buffer });
    },
    [pdfUrl, setFile, setPdfUrl, setConversionState, setConversionError, setConversionProgress],
  );

  /** Запускает конвертацию текущего файла в PDF. */
  const convert = useCallback(async () => {
    if (!file || conversionState === 'converting' || convertInFlight) return;

    convertInFlight = true;
    setConversionState('converting');
    setConversionProgress(null);
    setConversionError(null);

    try {
      const pdfBuffer = await client.convert(file.buffer, file.type);

      // Старый Blob URL больше не нужен — освобождаем память
      if (pdfUrl) URL.revokeObjectURL(pdfUrl);
      const blob = new Blob([pdfBuffer], { type: 'application/pdf' });
      setPdfUrl(URL.createObjectURL(blob));
      setConversionState('done');
    } catch (error) {
      setConversionError(error instanceof Error ? error.message : String(error));
      setConversionState('error');
    } finally {
      convertInFlight = false;
    }
  }, [file, conversionState, pdfUrl, setConversionState, setConversionProgress, setConversionError, setPdfUrl]);

  /** Повторная попытка инициализации движка после ошибки. */
  const retryInit = useCallback(() => {
    setOfficeState('initializing');
    setOfficeInitError(null);
    setOfficeInitProgress(null);
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
    retryInit,
  };
}
