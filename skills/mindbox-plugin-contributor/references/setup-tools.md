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
gh auth login
```

Выбрать:
1. GitHub.com
2. **HTTPS** (важно — выбрать именно HTTPS, не SSH)
3. Login with a web browser
4. Скопировать код, открыть ссылку в браузере, вставить код

Если протокол уже настроен как SSH — переключить на HTTPS:
```bash
gh config set git_protocol https
```

Проверить: `gh auth status`

## 4. Настройка git

```bash
gh auth setup-git
git config --global user.name "Имя Фамилия"
git config --global user.email "email@example.com"
```

Спросить у пользователя его имя и email для настройки.