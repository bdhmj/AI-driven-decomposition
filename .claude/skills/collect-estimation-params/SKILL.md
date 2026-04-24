---
name: collect-estimation-params
description: Collects estimation parameters from PM after decomposition — confirms specialist roster, then project coefficients (K), specialist rates, margin. Use when user says "собери параметры оценки", "настрой коэффициенты", "введи ставки", "добавь маржу", or when moving to estimation step 3.5 (between decomposition and xlsx generation).
metadata:
  author: metalamp
  version: 1.1.0
---

# Collect Estimation Parameters

## Role

You are a project manager's assistant that collects estimation parameters needed to calculate final project cost. Work strictly in Russian.

## Instructions

### Step 1: Read inputs

1. Read `output/decomposition.json` — extract unique list of specialists from `modules[].tasks[].specialist`.
2. Read `.claude/skills/collect-estimation-params/references/defaults.json` — get defaults (`mandatory_specialists`, `coefficients`, `rates`, `margin_pct`).
3. Build initial specialist roster:
   - Start with unique specialists from decomposition
   - Union with `mandatory_specialists` from defaults (currently: `PM`, `DevOps`, `QA`) — add any missing ones
   - Preserve insertion order: first decomposition order, then mandatory ones not yet present

### Step 2: Confirm specialist roster (block 0)

Show PM the proposed roster:

```
👥 СОСТАВ КОМАНДЫ ПРОЕКТА

Программа определила необходимых специалистов из декомпозиции и добавила обязательных (PM, DevOps, QA).

1. Backend
2. Frontend
3. UX/UI дизайнер
4. Аналитик
5. QA                ← обязательный
6. DevOps            ← обязательный
7. PM                ← обязательный

Отправь «ок» если состав устраивает, или команды редактирования:
  -N          — убрать специалиста под номером N (включая обязательных)
  +Название   — добавить специалиста (например: +Mobile Developer)
  можно несколько команд подряд через пробел, например: -3 +Data Engineer
```

Mark mandatory specialists (`PM`, `DevOps`, `QA`) with a note «← обязательный» so PM видит, какие добавлены автоматически.

Loop until user sends `ок` / `ok` / `далее`:
- Parse `-N` → remove Nth specialist from roster
- Parse `+Name` → append specialist to roster (if not already present); if rate unknown, add to rates with default 20 and flag «нет в дефолтах — поставлю ставку 20, уточни на блоке ставок»
- Multiple commands on one line allowed, parse sequentially
- After each change, show updated numbered list and re-prompt

**Rules for this step:**
- Mandatory specialists can be removed, but warn once: «Убираю DevOps/QA/PM — убедись, что это осознанное решение (обычно они нужны на любом проекте).»
- If user removes a specialist who has tasks in `decomposition.json` — warn: «У "Frontend" есть N задач в декомпозиции. Если убрать — его стоимость не посчитается, но задачи останутся. Продолжить? (да/нет)»
- Roster cannot be empty — reject: «Список не может быть пустым».

Save final roster in memory for next steps.

### Step 3: Collect coefficients (block 1)

Show PM the coefficients block with defaults and computed K:

```
⚙️ КОЭФФИЦИЕНТЫ ПРОЕКТА

1. Проверка и отладка задач (%): 20
2. Код ревью (часов/день): 1
3. Коммуникации (часов/неделю): 3
4. Тестировщик (% от общего): 0
5. Буфер на риски (%): 20
6. DevOps доп. (%): 0
7. Менеджер (% от макс. специалиста): 30

📐 Итого K = 1.60

Отправь номер и новое значение через пробел (например: 5 25), или «ок» если всё верно.
```

**K formula** (qa_pct и pm_pct НЕ входят в K):
```
K = 1 + code_review_hours/8 + communication_hours/40 + debug_pct/100 + risk_buffer_pct/100 + devops_pct/100
```

Loop until user sends `ок` / `ok` / `далее`:
- Parse `N VALUE` → update coefficient N (1-based index over the 7 params above)
- Recompute K, show updated block

### Step 4: Collect rates (block 2)

Show PM the rates block for the confirmed roster from Step 2 (in the same order):

```
💰 СТАВКИ ($/час, внутренние — на руки)

1. Backend: 25
2. Frontend: 20
3. UX/UI дизайнер: 20
4. Аналитик: 20
5. QA: 15
6. DevOps: 25
7. PM: 25

Отправь номер и новое значение (например: 1 30), или «ок» если всё верно.
```

Loop until `ок`. Parse `N VALUE` → update rate N.

### Step 5: Collect margin (block 3)

```
📊 МАРЖА

Наценка на внутренние ставки (%): 100

Клиентская ставка = внутренняя × (1 + маржа/100).
Например: Backend 25$/час + 100% маржа → клиенту 50$/час.

Отправь новое значение маржи или «ок» если оставляем 100.
```

Loop until `ок`. Parse single number → update `margin_pct`.

### Step 6: Gantt chart (block 4)

Ask whether to include the Gantt visualization sheet in the xlsx:

```
🗓 GANTT-CHART

Нужен ли лист с графиком Ганта в xlsx? Отдельный лист-визуализация сроков по задачам.
На цифры оценки (стоимость, сроки, недели) это не влияет — Gantt только рисует
уже посчитанные данные.

y / да / ок — включить (по умолчанию)
n / нет     — пропустить (xlsx на 5 листов без GANTT Chart)
```

Parse `y / да / ok / ок / yes` → `generate_gantt: true`. Parse `n / нет / no / skip` → `generate_gantt: false`. On unrecognised input — re-ask.

Default if skipped: `true`.

### Step 7: Save output

Save to `output/estimation_params.json` — include `specialists` field (final roster in order) and `generate_gantt` flag:

```json
{
  "specialists": ["Backend", "Frontend", "UX/UI дизайнер", "Аналитик", "QA", "DevOps", "PM"],
  "coefficients": {
    "debug_pct": 20,
    "code_review_hours": 1,
    "communication_hours": 3,
    "qa_pct": 0,
    "risk_buffer_pct": 20,
    "devops_pct": 0,
    "pm_pct": 30
  },
  "K": 1.60,
  "rates": {
    "Backend": 25,
    "Frontend": 20,
    "UX/UI дизайнер": 20,
    "Аналитик": 20,
    "QA": 15,
    "DevOps": 25,
    "PM": 25
  },
  "margin_pct": 100,
  "generate_gantt": true,
  "currency": "$"
}
```

Confirm to user: «✅ Параметры сохранены. Генерирую финальную таблицу...» and proceed to run `build_xlsx.py --params`.

## Rules

- Always in Russian
- Show K computed to 2 decimal places
- If user's input is not a recognised command or `ок` — politely re-ask
- Negative values — reject with short message («Значение должно быть ≥ 0»)
- `PM`, `DevOps`, `QA` are mandatory by default — always pre-added to roster, but user can remove them explicitly at Step 2

## Error handling

- `output/decomposition.json` not found → «Сначала запусти этап декомпозиции (skill: decompose-tasks)»
- `defaults.json` not found → use hardcoded fallback: mandatory `["PM","DevOps","QA"]`, coefficients and rates as listed above, margin 100
- Invalid JSON in decomposition → «Файл output/decomposition.json повреждён. Перезапусти декомпозицию»
