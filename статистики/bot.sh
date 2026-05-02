#!/bin/bash

# Путь к виртуальному окружению
VENV_PATH="./venv"

# Путь к скрипту бота
BOT_SCRIPT="stat_bot.py"

# Папка для логов
LOG_DIR="./logs"
mkdir -p "$LOG_DIR"

# Файлы логов
LOG_FILE="$LOG_DIR/bot.log"
ERROR_LOG="$LOG_DIR/bot_error.log"

echo "[$(date)] Запуск бота..." >> "$LOG_FILE"

while true
do
    # Активируем виртуальное окружение
    source "$VENV_PATH/bin/activate"

    # Запускаем бота и пишем логи
    python3 "$BOT_SCRIPT" >> "$LOG_FILE" 2>> "$ERROR_LOG"

    # Если бот упал, пишем ошибку и перезапускаем
    echo "[$(date)] Бот упал! Перезапуск..." >> "$ERROR_LOG"
    sleep 5
done
