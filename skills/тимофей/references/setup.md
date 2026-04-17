# Как подключить базу знаний Mindbox к Claude Code

База знаний Mindbox — это MCP-сервер, который позволяет Claude искать по статьям поддержки Mindbox. Без него скилл «Тимофей» работает только по CRM-урокам. С ним — ещё и по базе знаний платформы.

## Подключение

1. Открой терминал.
2. Выполни команду:

```bash
claude mcp add --transport http mindbox-knowledge https://ai-support-mcp.mindbox.ru/mcp
```

3. Готово.

## Проверка

Спроси Claude что-нибудь про продукт, например: «Как настроить триггерную рассылку в Mindbox?»

Если в ответе Claude использует инструменты `search_knowledge` или `get_article` — подключено.

## Если не работает

- Убедись, что Claude Code обновлён до последней версии: `claude update`
- Проверь список подключённых MCP-серверов: `claude mcp list`
- Если `mindbox-knowledge` нет в списке — повтори шаг 2
- Если есть, но не работает — удали и подключи заново:

```bash
claude mcp remove mindbox-knowledge
claude mcp add --transport http mindbox-knowledge https://ai-support-mcp.mindbox.ru/mcp
```
