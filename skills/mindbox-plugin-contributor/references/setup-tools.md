# Установка инструментов

Пошаговая инструкция для установки всех необходимых инструментов. Выполнять по порядку.

## 1. Homebrew

Проверка: `brew --version`

Если не установлен:
```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

## 2. GitHub CLI (gh)

Проверка: `gh --version`

Если не установлен (git установится автоматически как зависимость):
```bash
brew install gh
```

## 3. Авторизация в GitHub

```bash
gh auth login --web --git-protocol https
```

Команда откроет браузер для авторизации и автоматически настроит протокол HTTPS.

Проверить: `gh auth status`

## 4. Настройка git

```bash
gh auth setup-git
git config --global user.name "Имя Фамилия"
git config --global user.email "email@example.com"
```

Спросить у пользователя его имя и email для настройки.