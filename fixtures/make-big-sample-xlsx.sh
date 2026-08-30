#!/bin/bash

# Генератор XLSX-файла ~10 МБ с использованием виртуального окружения
# Использование: ./generate_xlsx.sh [имя_файла]

set -e

# 1. Проверка наличия python3
if ! command -v python3 &> /dev/null; then
    echo "Ошибка: python3 не установлен."
    exit 1
fi

# 2. Создаём виртуальное окружение, если его нет
VENV_DIR=".venv_xlsx"
if [ ! -d "$VENV_DIR" ]; then
    echo "Создаём виртуальное окружение в $VENV_DIR..."
    python3 -m venv "$VENV_DIR"
fi

# 3. Устанавливаем openpyxl в виртуальное окружение (если ещё не установлена)
if ! "$VENV_DIR/bin/python" -c "import openpyxl" 2>/dev/null; then
    echo "Устанавливаем openpyxl в виртуальное окружение..."
    "$VENV_DIR/bin/pip" install openpyxl
fi

# 4. Создаём временный Python-скрипт для генерации
PYSCRIPT=$(mktemp)
cat > "$PYSCRIPT" << 'EOF'
import openpyxl
from openpyxl import Workbook
import sys
import os

def generate_xlsx(filename, num_rows):
    wb = Workbook()
    ws = wb.active
    headers = [f'Col_{i}' for i in range(1, 11)]
    ws.append(headers)
    for i in range(num_rows):
        row = [f'Data_{i}_{j}' for j in range(10)]
        ws.append(row)
        if (i + 1) % 50000 == 0:
            print(f'Сгенерировано {i + 1} строк...')
    wb.save(filename)
    size = os.path.getsize(filename)
    print(f'Создано {num_rows} строк, размер: {size} байт ({size / 1024 / 1024:.2f} МБ)')
    return size

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Использование: python3 script.py <имя_файла> <количество_строк>")
        sys.exit(1)
    filename = sys.argv[1]
    num_rows = int(sys.argv[2])
    generate_xlsx(filename, num_rows)
EOF

# 5. Параметры
OUTPUT="${1:-large_file.xlsx}"
TARGET_SIZE=$((10 * 1024 * 1024))   # 10 МБ
ROWS=100000
INCREMENT=50000

echo "Генерация XLSX-файла '$OUTPUT' с целевым размером ~10 МБ..."

# 6. Цикл подбора количества строк
while true; do
    echo "Пробуем с $ROWS строками..."
    "$VENV_DIR/bin/python" "$PYSCRIPT" "$OUTPUT" "$ROWS"
    SIZE=$(wc -c < "$OUTPUT")
    echo "Текущий размер: $SIZE байт"
    if [ "$SIZE" -ge "$TARGET_SIZE" ]; then
        echo "Целевой размер достигнут."
        break
    else
        echo "Размер слишком мал, увеличиваем количество строк на $INCREMENT..."
        ROWS=$((ROWS + INCREMENT))
    fi
done

# 7. Очистка временного скрипта
rm "$PYSCRIPT"

echo "Готово. Итоговый файл: $OUTPUT ($(du -h "$OUTPUT" | cut -f1))"
echo "Виртуальное окружение оставлено в $VENV_DIR (можно удалить вручную: rm -rf $VENV_DIR)"