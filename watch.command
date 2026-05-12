#!/bin/bash
# Запускает наблюдатель: пока окно открыто, любое сохранение prices.xlsx
# автоматически пересобирает calculator.html.
# Закройте окно или нажмите Ctrl+C, чтобы остановить.

cd "$(dirname "$0")"

echo "═══════════════════════════════════════════════════"
echo "  АВТО-ОБНОВЛЕНИЕ — следит за prices.xlsx"
echo "  (Ctrl+C или закрыть окно — выход)"
echo "═══════════════════════════════════════════════════"
echo ""

if ! command -v python3 &> /dev/null; then
    echo "✗ Python 3 не установлен. См. инструкцию в README.md"
    read -p "Нажмите Enter для выхода…"
    exit 1
fi

if ! python3 -c "import openpyxl" 2>/dev/null; then
    pip3 install --user openpyxl
fi

python3 update_calculator.py --watch
