#!/usr/bin/env bash

# Автообновление marketplace-репозитория плагина mindbox-editor-plugin.
# Запускается при каждом старте сессии Claude Code.
# Для пользователя полностью бесшумный — логи пишутся в файл.

# Подавить весь вывод в stdout/stderr хука
exec >/dev/null 2>&1

PLUGIN_NAME="mindbox-editor-plugin"
LOG_DIR="$HOME/.claude/plugins/logs"
LOG_FILE="$LOG_DIR/$PLUGIN_NAME-update.log"
MARKETPLACE_DIR="$HOME/.claude/plugins/marketplaces/$PLUGIN_NAME"
MAX_LOG_LINES=50

log() {
  local level="$1"
  local message="$2"
  local timestamp
  timestamp=$(date '+%Y-%m-%d %H:%M:%S')

  mkdir -p "$LOG_DIR"
  echo "[$timestamp] $level: $message" >> "$LOG_FILE"

  # Ротация: оставить последние MAX_LOG_LINES строк
  if [ -f "$LOG_FILE" ]; then
    local line_count
    line_count=$(wc -l < "$LOG_FILE")
    if [ "$line_count" -gt "$MAX_LOG_LINES" ]; then
      tail -n "$MAX_LOG_LINES" "$LOG_FILE" > "$LOG_FILE.tmp" && mv "$LOG_FILE.tmp" "$LOG_FILE"
    fi
  fi
}

# Проверить что marketplace-директория существует и является git-репозиторием
if [ ! -d "$MARKETPLACE_DIR/.git" ]; then
  log "SKIP" "Marketplace directory not found or not a git repo: $MARKETPLACE_DIR"
  exit 0
fi

cd "$MARKETPLACE_DIR" || { log "ERROR" "Cannot cd to $MARKETPLACE_DIR"; exit 0; }

# Определить дефолтную ветку
DEFAULT_BRANCH=$(git symbolic-ref refs/remotes/origin/HEAD 2>/dev/null | sed 's@^refs/remotes/origin/@@')
if [ -z "$DEFAULT_BRANCH" ]; then
  DEFAULT_BRANCH="main"
fi

# Получить обновления только для целевой ветки
if ! git fetch --quiet origin "$DEFAULT_BRANCH" 2>/dev/null; then
  log "ERROR" "git fetch failed (network issue?)"
  exit 0
fi

# Сравнить локальный и удалённый SHA
LOCAL_SHA=$(git rev-parse --short=7 HEAD 2>/dev/null)
REMOTE_SHA=$(git rev-parse --short=7 "origin/$DEFAULT_BRANCH" 2>/dev/null)

if [ -z "$LOCAL_SHA" ] || [ -z "$REMOTE_SHA" ]; then
  log "ERROR" "Cannot resolve HEAD or origin/$DEFAULT_BRANCH"
  exit 0
fi

if [ "$LOCAL_SHA" = "$REMOTE_SHA" ]; then
  log "OK" "Already up-to-date ($LOCAL_SHA)"
  exit 0
fi

# Применить обновления
if git pull --ff-only origin "$DEFAULT_BRANCH" 2>/dev/null; then
  NEW_SHA=$(git rev-parse --short=7 HEAD 2>/dev/null)
  log "UPDATED" "Updated from $LOCAL_SHA to $NEW_SHA"
else
  log "ERROR" "git pull --ff-only failed (not a fast-forward?). Will retry next session."
fi

exit 0
