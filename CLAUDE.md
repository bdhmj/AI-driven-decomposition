# AI-driven Project Estimation Workspace

Рабочее пространство для оценки и декомпозиции IT-проектов. Работает через Claude Code (VS Code extension).

## Workflow

PM кидает описание проекта → Claude проходит три этапа → на выходе ТЗ (.docx) + декомпозиция (.xlsx с GANTT).

### Этап 1: Анализ запроса
Прочитай описание клиента из `input/`. Определи, достаточно ли информации для ТЗ.
- Если достаточно → переходи к Этапу 2
- Если нет → сформулируй уточняющие вопросы (3-6 штук, на русском)

Используй скилл: `skills/analyze-request/SKILL.md`

### Этап 2: Генерация ТЗ
Сгенерируй трёхчастное техническое задание:
1. **MVP — Базовый скоуп** — что точно делаем, с явными границами
2. **Открытые вопросы** — что может повлиять на оценку
3. **Чеклист расширений** — таблица всего, что можно добавить поверх MVP

Используй скилл: `skills/generate-spec/SKILL.md`
Шаблон вывода: `skills/generate-spec/references/spec-template.md`

Сохрани результат в `output/spec.md`.
Для конвертации в .docx: `python scripts/build_docx.py output/spec.md output/Техническое_задание.docx`

### Этап 3: Декомпозиция и оценка
Декомпозируй ТЗ на задачи по модулям с оценками в днях. Каждая задача помечена фазой: `mvp` или `post-mvp`.

Используй скилл: `skills/decompose-tasks/SKILL.md`
Референс: `skills/decompose-tasks/references/estimation-guide.md`

Сохрани JSON в `output/decomposition.json`.
Для генерации .xlsx: `python scripts/build_xlsx.py output/decomposition.json output/Оценка_проекта.xlsx`

## Структура проекта

```
input/              — описания проектов от клиентов
output/             — результаты: spec.md, decomposition.json, .docx, .xlsx
skills/             — скиллы (промпты + референсы)
  analyze-request/  — анализ полноты запроса
  generate-spec/    — генерация трёхчастного ТЗ
  decompose-tasks/  — декомпозиция на задачи
scripts/            — утилиты для генерации документов
  build_docx.py     — markdown → .docx
  build_xlsx.py     — decomposition.json → .xlsx с GANTT
```

## Архивация проекта

После завершения оценки заархивируй результаты:
```bash
./scripts/archive.sh "Название проекта"
```
Это перенесёт всё из `input/` и `output/` в `projects/YYYY-MM-DD_Название/`. Папки input/output останутся чистыми для следующего проекта.

## Правила
- Язык ТЗ, вопросов, задач: всегда русский (независимо от языка входа)
- Формулировки однозначные — никаких "и т.д.", "при необходимости"
- Каждое утверждение проверяемое ("API < 500ms", не "быстрый API")
- Чеклист расширений — исчерпывающий, без ограничений по количеству

## Зависимости
```
pip install python-docx openpyxl
```
