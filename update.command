#!/bin/bash
# Двойной клик на этом файле обновляет calculator.html из prices.xlsx
# При первом запуске разрешите выполнение: ПКМ на файле → «Открыть»

cd "$(dirname "$0")"

echo "═══════════════════════════════════════════════════"
echo "  Обновление калькулятора стоимости строительства"
echo "═══════════════════════════════════════════════════"
echo ""

# Проверка Python 3
if ! command -v python3 &> /dev/null; then
    echo "✗ Python 3 не установлен."
    echo ""
    echo "Установите Homebrew (если ещё нет):"
    echo '  /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"'
    echo "затем:"
    echo "  brew install python"
    echo ""
    read -p "Нажмите Enter для выхода…"
    exit 1
fi

# Проверка openpyxl
if ! python3 -c "import openpyxl" 2>/dev/null; then
    echo "⚙  Устанавливаю недостающую библиотеку openpyxl…"
    pip3 install --user openpyxl
    echo ""
fi

# Запуск обновления
python3 update_calculator.py
RESULT=$?

echo ""
if [ $RESULT -eq 0 ]; then
    echo "✓ Готово. Файл calculator.html обновлён."
    echo ""
    read -p "Открыть calculator.html в браузере? (y/n): " open_it
    if [[ "$open_it" =~ ^[Yy] ]]; then
        open calculator.html
    fi
else
    echo "✗ Обновление не удалось. Проверьте сообщения выше."
fi

echo ""
read -p "Нажмите Enter для выхода…"
