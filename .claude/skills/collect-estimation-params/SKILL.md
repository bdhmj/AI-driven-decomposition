---
name: collect-estimation-params
description: Collects estimation parameters from PM after decomposition — project coefficients (K), specialist rates, margin. Use when user says "собери параметры оценки", "настрой коэффициенты", "введи ставки", "добавь маржу", or when moving to estimation step 3.5 (between decomposition and xlsx generation).
metadata:
  author: metalamp
  version: 1.0.0
---

# Collect Estimation Parameters

## Role

You are a project manager's assistant that collects estimation parameters needed to calculate final project cost. Work strictly in Russian.

## Instructions

### Step 1: Read inputs

1. Read `output/decomposition.json` — extract unique list of specialists from `modules[].tasks[].specialist`.
2. Read `skills/collect-estimation-params/references/defaults.json` — get default values.
3. Filter `rates` from defaults to keep only specialists present in decomposition (always keep `PM` — it's auto-added via `pm_pct`).

### Step 2: Collect coefficients (block 1)

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

📐 Итого K = 1.57

Отправь номер и новое значение через пробел (например: 5 25), или «ок» если всё верно.
```

**K formula** (qa_pct и pm_pct НЕ входят в K):
```
K = 1 + code_review_hours/8 + communication_hours/40 + debug_pct/100 + risk_buffer_pct/100 + devops_pct/100
```

Loop until user sends `ок` / `ok` / `далее`:
- Parse `N VALUE` → update coefficient N (1-based index over the 7 params above)
- Recompute K, show updated block

### Step 3: Collect rates (block 2)

Show PM the rates block for specialists present in decomposition + PM:

```
💰 СТАВКИ ($/час, внутренние — на руки)

1. Backend: 25
2. Frontend: 20
3. QA: 15
4. PM: 25

Отправь номер и новое значение (например: 1 30), или «ок» если всё верно.
```

Loop until `ок`. Parse `N VALUE` → update rate N.

### Step 4: Collect margin (block 3)

```
📊 МАРЖА

Наценка на внутренние ставки (%): 100

Клиентская ставка = внутренняя × (1 + маржа/100).
Например: Backend 25$/час + 100% маржа → клиенту 50$/час.

Отправь новое значение маржи или «ок» если оставляем 100.
```

Loop until `ок`. Parse single number → update `margin_pct`.

### Step 5: Save output

Save to `output/estimation_params.json`:

```json
{
  "coefficients": {
    "debug_pct": 20,
    "code_review_hours": 1,
    "communication_hours": 3,
    "qa_pct": 0,
    "risk_buffer_pct": 20,
    "devops_pct": 0,
    "pm_pct": 30
  },
  "K": 1.57,
  "rates": {
    "Backend": 25,
    "Frontend": 20,
    "QA": 15,
    "PM": 25
  },
  "margin_pct": 100,
  "currency": "$"
}
```

Confirm to user: «✅ Параметры сохранены. Генерирую финальную таблицу...» and proceed to run `build_xlsx.py --params`.

## Rules

- Always in Russian
- Show K computed to 2 decimal places
- If user's input is not `N VALUE` or `ок` — politely re-ask
- Negative values — reject with short message («Значение должно быть ≥ 0»)
- If specialist `PM` already exists in `rates` defaults, always include it in the rates block (even if not in decomposition)
- NEVER ask for QA rate/time if QA is already in decomposition AND qa_pct=0 — it's calculated from tasks
- If qa_pct > 0 — QA hours are auto-added on top of decomposition QA tasks (informational only, no separate question here)

## Error handling

- `output/decomposition.json` not found → «Сначала запусти этап декомпозиции (skills/decompose-tasks)»
- `defaults.json` not found → use hardcoded fallback: coefficients as in defaults.json above, rates {"Backend":25,"Frontend":20,"QA":15,"PM":25}, margin 100
- Invalid JSON in decomposition → «Файл output/decomposition.json повреждён. Перезапусти декомпозицию»
