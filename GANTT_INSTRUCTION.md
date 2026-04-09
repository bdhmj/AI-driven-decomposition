# Инструкция: построение Excel Gantt-чарта из CSV (roadmap)

Цель документа — позволить новому экземпляру Claude в другом проекте полностью воспроизвести ту же логику и визуальный стиль Excel-файла, что был отработан ранее. Читай целиком, ничего не пропуская: каждое «очевидное» решение — результат итераций и правок пользователя.

---

## 1. КОНТЕКСТ ЗАДАЧИ

**Что строим:** дневной Gantt-chart в одном Excel-листе по roadmap'у IT-проекта (P2P-платформа). Каждая строка — задача, каждый столбец в data-области = 1 календарный день. Задачи сгруппированы по «фазам» (команды/направления: DevOps, Smart contracts, Backend #3, Backend #4, Frontend, QA, Analyst).

**Зачем:** у пользователя есть CSV roadmap с задачами, где даты проставлены как если бы работали без выходных. Нужно:
1. Отобразить задачи на таймлайне.
2. Явно показать выходные (суббота/воскресенье) как серые узкие колонки с разрывом бара.
3. **Пересчитать длительности так, чтобы исходные дни стали РАБОЧИМИ днями** (то есть реальная дата окончания каждой задачи сдвигается из-за выходных).
4. Внутри фазы задачи, идущие друг за другом, должны каскадироваться: сдвиг одной задачи тянет следующие.
5. Стилизовать под «красивый» корпоративный вид (каждая фаза — свой цвет заголовка, фона строк и полосок бара).

**Конечный пользователь:** менеджер проекта, который смотрит файл в Excel, ожидает «гугл-таблица-подобный» вид со всплывающими цветами фаз и чётким разделением по неделям.

**Что это НЕ:** не интерактивный Gantt с формулами, не JS-рендер, не PDF. Это статический Excel, собранный openpyxl'ом, в котором бары нарисованы через `PatternFill` отдельных ячеек (а не через chart/shapes).

---

## 2. ВХОДНЫЕ ДАННЫЕ

### 2.1. Формат CSV
Пользователь присылает CSV (выгрузка из GitHub Projects). Кодировка UTF-8. Первая строка = `GANTT Chart`, вторая — заголовок:

```
"Task ID","Title","Description","Status","Assignee","Start Date","End Date","Effort","Priority","Blocking","Blocked by","Phase"
```

**Важные поля (всё остальное игнорируем):**
- `Title` — название задачи. **Может содержать переносы строк и повторения** (например, строка разбита на многострочный блок с повторами названия — это GitHub-эффект, из него берём первое вменяемое название).
- `Start Date` — ISO `2025-06-02T00:00:00.000Z`. Берём только дату.
- `End Date` — такой же формат.
- `Phase` — одна из 7 фаз (см. ниже).

Остальные поля (`Task ID`, `Description`, `Status`, `Assignee`, `Effort`, `Priority`, `Blocking`, `Blocked by`) — **опциональные, обычно пустые, игнорируем**.

### 2.2. Список фаз (строго эти 7 строк, regex-match)
```
Infrastructure (DevOps)
Smart contracts
Backend #3
Backend #4
Frontend
QA
Analyst
```

### 2.3. Как парсить
CSV не парсь питоновским `csv` модулем «в лоб» — из-за многострочных Title-блоков с кавычками и переносами это хрупко. **Правильный подход:** либо вручную прочитать CSV и нормализовать, либо (проще и надёжнее) — прочитать файл только ради того, чтобы увидеть глазами порядок и даты задач, а затем **захардкодить данные в Python-скрипт** в виде списка кортежей:

```python
csv_data = [
    # (phase, title, orig_start, orig_end)
    ("Infrastructure (DevOps)", "Создание и настройка репозиториев", "2025-06-02", "2025-06-03"),
    ...
]
```

Это решение было принято после нескольких итераций: пользователь много раз присылает обновлённые CSV с переставленными задачами, и проще обновлять список кортежей вручную по данным из свежего CSV, чем поддерживать парсер.

### 2.4. Порядок задач в `csv_data`
**КРИТИЧЕСКИ ВАЖНО.** Задачи внутри фазы должны идти в `csv_data` в том же порядке, в котором они идут в CSV. Каскадная логика (см. §5) опирается на порядок. Если в исходном CSV задача A стоит перед задачей B — в `csv_data` должно быть так же. Если пользователь пересылает обновлённый CSV с переставленными задачами — обновляй порядок в `csv_data`.

### 2.5. Особенности дат в CSV
- Длительность задачи в CSV = `(end - start).days + 1` (включая оба конца).
- Даты в CSV не учитывают выходные (как будто суббота = рабочий день).
- Разные фазы могут начинаться с разных дат (например, User Module в Backend #4 стартует 03.06, а не 02.06).
- Бывают параллельные задачи внутри фазы (две задачи с одинаковыми датами или пересекающимися). Пример: `Telegram (уведомление) 18-19.07` и `Telegram (OTC) 19.07` — должны заканчиваться одновременно.
- Бывают длинные фоновые задачи, перекрывающие всю фазу (Smart contracts → «Аудит безопасности» 13.06–30.07).

---

## 3. СТРУКТУРА ИТОГОВОГО XLSX

**Один лист**, `title = "GANTT Chart"`. Другие листы не создаём.

### 3.1. Раскладка колонок

| Колонка | Буква | Ширина | Содержимое |
|---|---|---|---|
| 1 | A | 22 | Фаза |
| 2 | B | 14 | Роль |
| 3 | C | 42 | Задача |
| 4 | D | 11 | Старт (DD.MM.YY) |
| 5 | E | 6  | Дней (int, рабочие дни) |
| 6 | F | 11 | Конец (DD.MM.YY) |
| 7..N | G..? | 3.8 (будни) / 2.5 (выходные) | По одной колонке на каждый календарный день |

Дата-область начинается с колонки **G** (индекс 7). В коде это константа `DATA_COL_START = 7`.

### 3.2. Раскладка строк

- **Строка 1** — месяцы: `Июнь 2025`, `Июль 2025`, … Объединённые ячейки по диапазону колонок текущего месяца в data-области.
- **Строка 2** — заголовки колонок A–F (`Фаза`, `Роль`, `Задача`, `Старт`, `Дней`, `Конец`) + числа дней (1, 2, … 31) в data-области.
- **Строка 3** — пустые A–F (только заливка), дни недели (`Пн`, `Вт`, `Ср`, `Чт`, `Пт`, `Сб`, `Вс`) в data-области.
- **Строка 4 и далее** — данные. Каждая фаза начинается с **строки-заголовка фазы** (merge A:F, название фазы на цветном фоне), затем строки задач.

### 3.3. Freeze panes
`ws.freeze_panes = 'G4'` — закрепляем первые 3 строки и 6 колонок (A–F).

### 3.4. Высоты строк
- Строка 1 (месяцы): 20
- Строка 2 (числа/заголовки): 18
- Строка 3 (дни недели): 14
- Строка-заголовок фазы: 22
- Строка задачи: 20

### 3.5. Нет диаграмм, нет условного форматирования, нет формул, нет автофильтров
Всё «рисуется» руками через `PatternFill`.

---

## 4. БИБЛИОТЕКА

**openpyxl** (любая относительно свежая версия, 3.x подходит). Причина выбора: нужна полная работа с merge'ами, `PatternFill`, `Border`, установкой размеров колонок и freeze panes. xlsxwriter не подходит, потому что нужно поддерживать итеративное редактирование файла (хотя конкретно в текущем решении мы файл всегда создаём с нуля).

Импорты:
```python
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
```

---

## 5. ЛОГИКА ПЕРЕСЧЁТА ДАТ (ядро задачи — не упростить!)

### 5.1. Базовая идея
Исходные `start`/`end` в CSV трактуем так, будто каждый день был рабочим. Значит `duration_workdays = (orig_end - orig_start).days + 1`. Новое окно задачи = `duration_workdays` **рабочих** дней, начиная с корректного рабочего старта, пропуская сб/вс.

### 5.2. Вспомогательные функции
```python
def next_workday(dt):
    """Если dt попадает на сб/вс — сдвинуть на ближайший пн."""
    while dt.weekday() >= 5:
        dt += timedelta(days=1)
    return dt

def add_workdays(start_dt, num_workdays):
    """Начиная со start_dt (уже рабочий день), вернуть дату,
    на которую приходится num_workdays-й рабочий день включительно."""
    cur = next_workday(start_dt)
    counted = 1
    while counted < num_workdays:
        cur += timedelta(days=1)
        if cur.weekday() < 5:
            counted += 1
    return cur
```

### 5.3. Каскадная логика внутри фазы (v6/v7c — проверенная версия!)
Это место правилось МНОЖЕСТВО раз, финальная рабочая версия ниже. **НЕ меняй условие `gap <= 1` на `gap == 1` — это ломает параллельные задачи.**

```python
tasks = []
phase_prev = {}  # phase -> (prev_orig_end, prev_new_end)

for phase, title, orig_start_str, orig_end_str in csv_data:
    os_dt = datetime.strptime(orig_start_str, "%Y-%m-%d")
    oe_dt = datetime.strptime(orig_end_str, "%Y-%m-%d")
    duration = (oe_dt - os_dt).days + 1

    if phase in phase_prev:
        prev_orig_end, prev_new_end = phase_prev[phase]
        gap = (os_dt - prev_orig_end).days
        if gap <= 1:
            # Последовательная или параллельная задача: цепляем к new_end предыдущей
            new_start = next_workday(prev_new_end + timedelta(days=gap))
        else:
            # Задача с разрывом: применяем накопленный сдвиг
            accumulated = (prev_new_end - prev_orig_end).days
            new_start = next_workday(os_dt + timedelta(days=accumulated))
    else:
        new_start = next_workday(os_dt)

    new_end = add_workdays(new_start, duration)
    tasks.append((phase, title, new_start, new_end, duration))
    phase_prev[phase] = (oe_dt, new_end)
```

### 5.4. Что означают ветки
- `gap <= 1` включает `gap == 1` (следующий день — нормальное последовательное продолжение), `gap == 0` (параллельная задача того же дня), `gap < 0` (параллельная задача, начавшаяся раньше конца предыдущей — например, overlap). Во всех этих случаях привязываем к `prev_new_end`, смещая на `gap` дней.
- `gap > 1` = между задачами есть пауза (в CSV между ними несколько «свободных» дней). Тогда не цепляем к prev_new_end, а берём `orig_start` и двигаем на накопленный сдвиг (`prev_new_end - prev_orig_end`).
- `phase_prev[phase] = (oe_dt, new_end)` — ВСЕГДА обновляется, даже если параллельная задача заканчивается раньше. Это намеренно: по v6-логике цепочка считается от ПОСЛЕДНЕЙ обработанной задачи, а не от самой поздней. Пользователь явно подтвердил такое поведение.

### 5.5. Подводный камень с параллельными Telegram/Backend #4
В одной из версий я пытался «исправить» поведение для параллельных задач, сохраняя максимальный `new_end`. Это сломало каскад для аналитики, идущей после параллельных задач. **Правильно — оставить тупое перезаписывание `phase_prev[phase]`, как в v6.**

---

## 6. ПОСТРОЕНИЕ ДИАПАЗОНА ДАТ ДЛЯ ЗАГОЛОВКА

```python
project_start = min(t[2] for t in tasks)
while project_start.weekday() != 0:   # понедельник
    project_start -= timedelta(days=1)
project_end = max(t[3] for t in tasks)
while project_end.weekday() != 4:     # пятница
    project_end += timedelta(days=1)
```

То есть визуальная сетка ВСЕГДА начинается с понедельника и заканчивается пятницей той недели, где заканчивается последняя задача (чтобы неделя была целой). Затем собираем `all_days = [project_start, project_start+1day, …, project_end]`.

---

## 7. СТИЛИЗАЦИЯ — ТОЧНЫЕ ЗНАЧЕНИЯ

### 7.1. Цвета фаз (hex без #)

```python
phase_colors = {
    "Infrastructure (DevOps)": {"header": "1F4E79", "fill": "D6E4F0", "bar": "5B9BD5"},
    "Smart contracts":         {"header": "7B2D26", "fill": "F2DCDB", "bar": "C0504D"},
    "Backend #3":              {"header": "4F6228", "fill": "EBF1DE", "bar": "9BBB59"},
    "Backend #4":              {"header": "31859C", "fill": "DAEEF3", "bar": "4BACC6"},
    "Frontend":                {"header": "E36C09", "fill": "FDE9D9", "bar": "F79646"},
    "QA":                      {"header": "60497A", "fill": "E4DFEC", "bar": "8064A2"},
    "Analyst":                 {"header": "4A452A", "fill": "F2F2E6", "bar": "948A54"},
}
```

Смысл:
- `header` — тёмный цвет для merge-строки заголовка фазы (и белый текст на нём).
- `fill` — светлый фон для всех ячеек строки задачи (в т.ч. «пустые» ячейки data-области вне диапазона задачи).
- `bar` — насыщенный цвет для ячеек, попадающих внутрь диапазона задачи (Пн–Пт).

### 7.2. Цвет выходных в строках задач
Берём `fill` фазы и ДЕЛАЕМ ЕГО НА 25 ЕДИНИЦ ТЕМНЕЕ по каждому каналу, чтобы сб/вс внутри task-row выделялись:
```python
rb, gb, bb = int(colors["fill"][:2],16), int(colors["fill"][2:4],16), int(colors["fill"][4:6],16)
wknd_row_color = f"{max(0,rb-25):02X}{max(0,gb-25):02X}{max(0,bb-25):02X}"
```

### 7.3. Общие цвета
- `header_fill` (строки 1 и 2, ячейки A–F и область месяцев/чисел): `2F5496`
- `weekend_header_fill` (числа/дни недели в сб/вс в шапке): `C0C0C0`
- `date_label_fill` (фон числа/дня недели в будни + пустая строка 3 в A–F): `D6DCE4`

### 7.4. Шрифты (Calibri везде)
- Заголовки колонок и месяцев: `Calibri 10 bold white (FFFFFF)`, для ячейки месяца — `size=11 bold`.
- Числа дней в строке 2: `Calibri 7 color 44546A`.
- Дни недели (Пн–Пт): `Calibri 6 color 808080`.
- Дни недели (Сб–Вс): `Calibri 6 bold color 999999`.
- Задачи в таблице: `Calibri 9`.
- Заголовок фазы (merge): `Calibri 10 bold white`.

### 7.5. Границы
Все ячейки — тонкая граница `BFBFBF`. Плюс специальная граница для колонок-воскресений (последний день недели) — правая сторона `medium 808080`, создаёт визуальный разделитель недель.

```python
thin_border = Border(
    left=Side(style='thin', color='BFBFBF'),
    right=Side(style='thin', color='BFBFBF'),
    top=Side(style='thin', color='BFBFBF'),
    bottom=Side(style='thin', color='BFBFBF'),
)
week_sep_border = Border(
    left=Side(style='thin', color='BFBFBF'),
    right=Side(style='medium', color='808080'),
    top=Side(style='thin', color='BFBFBF'),
    bottom=Side(style='thin', color='BFBFBF'),
)
```

Правило: если `day.weekday() == 6` (воскресенье) — ставим `week_sep_border`, иначе `thin_border`. **Это правило применяется В ТОМ ЧИСЛЕ в строках заголовка фазы и в шапке, чтобы линия недели была сквозной.**

### 7.6. Выравнивание
- Ячейки A–F в строках-задачах: `left_wrap = Alignment(vertical='center', wrap_text=True)` для колонок 1,2,3. `center_align` для 4,5,6.
- Все ячейки шапки, дата-области, заголовок фазы — `center_align` (кроме самого merge-заголовка фазы — там `horizontal='left', vertical='center'`).

### 7.7. Ширины колонок
```python
ws.column_dimensions['A'].width = 22
ws.column_dimensions['B'].width = 14
ws.column_dimensions['C'].width = 42
ws.column_dimensions['D'].width = 11
ws.column_dimensions['E'].width = 6
ws.column_dimensions['F'].width = 11
for i in range(num_days):
    col_letter = get_column_letter(DATA_COL_START + i)
    ws.column_dimensions[col_letter].width = 2.5 if all_days[i].weekday() >= 5 else 3.8
```

Выходные уже́е будней — это визуальный приём, делающий неделю чуть ритмичнее.

---

## 8. ПОШАГОВЫЙ АЛГОРИТМ ПОСТРОЕНИЯ ЛИСТА

1. `wb = openpyxl.Workbook(); ws = wb.active; ws.title = "GANTT Chart"`.
2. Задай `csv_data` (список кортежей).
3. Прогони через каскадную логику (§5), получи `tasks`.
4. Посчитай `project_start` (сдвинутый на понедельник), `project_end` (сдвинутый на пятницу), `all_days`, `num_days`.
5. Установи ширины колонок A–F и G…G+num_days-1 (§7.7).
6. **Строка 1 (месяцы):**
   - Высота 20.
   - A–F: пустые, заливка `header_fill`, тонкая граница.
   - Data-область: найди промежутки месяцев (`month_spans`), merge-ни их, напиши `"{Месяц} {Год}"` (русские названия), примени `header_fill`, центр, белый текст bold 11. Всем ячейкам внутри merge-диапазона проставь заливку и границу.
7. **Строка 2 (числа дней / заголовки):**
   - Высота 18.
   - A–F: `["Фаза","Роль","Задача","Старт","Дней","Конец"]`, `header_font` (10 bold white), `header_fill`, центр, тонкая граница.
   - Data-область: число `day.day`, шрифт 7pt `44546A`, заливка `D6DCE4` (будни) / `C0C0C0` (выходные), центр, для воскресений — `week_sep_border`.
8. **Строка 3 (дни недели):**
   - Высота 14.
   - A–F: пустые, заливка `date_label_fill`, тонкая граница.
   - Data-область: `["Пн","Вт","Ср","Чт","Пт","Сб","Вс"][day.weekday()]`, шрифт 6pt (будни 808080 обычный, выходные 999999 bold), соответствующая заливка, `week_sep_border` для воскресений.
9. **Строки задач (с 4-й строки):**
   - Перед каждой новой фазой — вставь **строку-заголовок фазы**: merge A:F, название фазы, белый bold, фон `header`-цвет. **Плюс** для A–F и для ВСЕХ колонок data-области той же строки — заливка `header`-цветом и корректные границы (включая week_sep_border для воскресений). Высота 22.
   - Затем строки задач:
     - A: phase, B: role (из `phase_to_role`), C: title, D: `start.strftime("%d.%m.%y")`, E: duration (int), F: `end.strftime("%d.%m.%y")`.
     - Шрифт `Calibri 9`, заливка `fill`-цвет фазы, граница тонкая, выравнивание: 1,2,3 — left_wrap, 4,5,6 — center.
     - Data-область: для каждого дня `all_days`:
       - Граница week_sep_border если вс, иначе thin_border.
       - Если `start_dt <= day <= end_dt` И `day.weekday() < 5` → заливка `bar` (цветной бар).
       - Иначе если `day.weekday() >= 5` → заливка «темнее fill на 25» (`wknd_row`).
       - Иначе → заливка `fill` (нейтральный фон строки).
     - Высота 20.
10. `ws.freeze_panes = 'G4'`.
11. `wb.save(output_path)`.

---

## 9. МАППИНГ phase → role (для колонки B)

```python
phase_to_role = {
    "Infrastructure (DevOps)": "DevOps",
    "Smart contracts": "Smart Contract",
    "Backend #3": "Backend #3",
    "Backend #4": "Backend #4",
    "Frontend": "Frontend",
    "QA": "QA",
    "Analyst": "Analyst",
}
```

---

## 10. НАЗВАНИЯ МЕСЯЦЕВ ПО-РУССКИ
```python
month_names_ru = {
    1:"Январь",2:"Февраль",3:"Март",4:"Апрель",5:"Май",6:"Июнь",
    7:"Июль",8:"Август",9:"Сентябрь",10:"Октябрь",11:"Ноябрь",12:"Декабрь"
}
```

---

## 11. ПОДВОДНЫЕ КАМНИ (все реальные ошибки из истории итераций)

1. **Не заменять длительность задач просто гэпами выходных.** Первая попытка: показал сб/вс как серые колонки поверх существующего бара — пользователь возмутился, что сроки не увеличились. Нужно именно пересчитать `end_date` методом «исходная длительность = рабочие дни».
2. **Не ломать каскад** внутри фазы. Если просто пересчитать каждую задачу независимо — backend'ы как заканчивались в конце июля, так и будут. Нужно тянуть `phase_prev` по фазе.
3. **`gap <= 1`, а не `gap == 1`.** Замена на `==` была попыткой чинить параллельные Frontend-задачи и сломала параллельный Telegram. После отката всё работает.
4. **Параллельные задачи должны заканчиваться одновременно** (Telegram-бот и Telegram-канал, 18-19.07 и 19.07). При `gap == 0` обе привязываются к prev_new_end через `next_workday(prev_new_end + timedelta(days=0))`.
5. **phase_prev перезаписывается всегда**, даже если новая задача короче предыдущей. Попытка сохранять max(new_end) сломала логику — откатил.
6. **PermissionError при сохранении**: если xlsx открыт в Excel — openpyxl не сможет перезаписать. Решение: сохранять в файл с инкрементом версии (`v11`, `v11b`, `v12`, ...). Не пытаться закрыть Excel программно.
7. **UnicodeEncodeError в print**: консоль Windows cp1251 не переваривает `→`. Использовать `->` или любые ASCII-символы в `print`.
8. **Порядок задач в `csv_data` = порядок в CSV.** Если пользователь переставил в CSV — переставить и в коде. v9 провалилась потому, что у Infrastructure первой стояла Nginx (17.06), а не репозитории (02.06) — визуально казалось что DevOps «не с первого дня начинается».
9. **Строка-заголовок фазы должна заливаться и в дата-области**, не только в A–F. Иначе визуально фаза «обрывается» на колонке F.
10. **Границы воскресений должны быть сквозными** — и в шапке, и в строках-заголовках фаз, и в строках задач. Иначе week separator ломается по вертикали.
11. **Длинные фоновые задачи** (Аудит безопасности ~1.5 месяца) тоже проходят через каскад и корректно растягиваются.
12. **Windows-путь сохранения** с raw-строкой: `r"C:\Users\...\file.xlsx"`.
13. **Не используй built-in chart openpyxl.** Бары рисуются руками через PatternFill — это намеренное решение.
14. **Shell**: скрипт запускается `python create_roadmap.py` из директории скрипта. Не забывай про `cd` в bash-сессии Windows.

---

## 12. ЭТАЛОННЫЕ ФАЙЛЫ

В папке `C:\Users\gregory\Desktop\` лежат инкрементальные версии xlsx:
- `P2P_Roadmap_Gantt_v1.xlsx` … `P2P_Roadmap_Gantt_v11b.xlsx`.
- **Финальный рабочий = `P2P_Roadmap_Gantt_v11b.xlsx`**. Все более ранние — промежуточные итерации, их открывать только для сравнения.
- Исходный скрипт: `C:\Users\gregory\Desktop\AI-driven decomposition\create_roadmap.py`.
- Входные CSV: `D:\Download\p2p roadmap.csv`, `p2p roadmap (1).csv` … `(5).csv`. Самая свежая = с максимальным номером в скобках.

Ожидаемые диапазоны дат последней версии:
```
Range: 02.06.25 - 05.09.25 (96 cal days)
Backend #4:              07.08.25
Backend #3:              15.08.25
Infrastructure (DevOps): 20.08.25
Smart contracts:         21.08.25
Frontend:                28.08.25
QA:                      05.09.25
Analyst:                 05.09.25
```

---

## 13. МИНИМАЛЬНО-ПОЛНЫЙ ШАБЛОН КОДА

```python
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "GANTT Chart"

csv_data = [
    # (phase, title, orig_start, orig_end) — заполнить по текущему CSV,
    # порядок внутри фазы должен совпадать с порядком в CSV!
]

# --- workday helpers ---
def next_workday(dt):
    while dt.weekday() >= 5:
        dt += timedelta(days=1)
    return dt

def add_workdays(start_dt, num_workdays):
    cur = next_workday(start_dt)
    counted = 1
    while counted < num_workdays:
        cur += timedelta(days=1)
        if cur.weekday() < 5:
            counted += 1
    return cur

# --- cascade (v6 logic — не менять!) ---
tasks = []
phase_prev = {}
for phase, title, os_str, oe_str in csv_data:
    os_dt = datetime.strptime(os_str, "%Y-%m-%d")
    oe_dt = datetime.strptime(oe_str, "%Y-%m-%d")
    duration = (oe_dt - os_dt).days + 1
    if phase in phase_prev:
        prev_orig_end, prev_new_end = phase_prev[phase]
        gap = (os_dt - prev_orig_end).days
        if gap <= 1:
            new_start = next_workday(prev_new_end + timedelta(days=gap))
        else:
            accumulated = (prev_new_end - prev_orig_end).days
            new_start = next_workday(os_dt + timedelta(days=accumulated))
    else:
        new_start = next_workday(os_dt)
    new_end = add_workdays(new_start, duration)
    tasks.append((phase, title, new_start, new_end, duration))
    phase_prev[phase] = (oe_dt, new_end)

# --- см. §7–§10 для стилизации и §8 для пошагового заполнения листа ---
# (полный код см. в create_roadmap.py)

ws.freeze_panes = 'G4'
wb.save(r"C:\Users\gregory\Desktop\P2P_Roadmap_Gantt_vXX.xlsx")
```

**Полный референс — файл `create_roadmap.py` (≈290 строк).** Скопируй его целиком в новый проект и обновляй только:
1. список `csv_data` (под свежий CSV);
2. путь `output_path` (инкрементная версия).

Всё остальное (цвета, размеры, каскад, рендер) — не трогать без явного запроса пользователя.

---

## 14. КОНТРОЛЬНЫЙ ЧЕК-ЛИСТ ПЕРЕД ОТДАЧЕЙ ФАЙЛА

- [ ] Все 7 фаз присутствуют и в правильном порядке DevOps → Smart contracts → Backend #3 → Backend #4 → Frontend → QA → Analyst.
- [ ] Первая задача каждой фазы (кроме User Module в Backend #4) стартует **02.06.2025**.
- [ ] Колонка G = понедельник, последняя колонка = пятница.
- [ ] Выходные в шапке — серые `C0C0C0`, узкие (width 2.5).
- [ ] Каждое воскресенье имеет правую границу `medium 808080` во всех строках.
- [ ] Заголовки фаз — merge A:F + заливка header-цветом через всю ширину листа.
- [ ] Бары в будни = `bar`-цвет, в выходные = «темнее fill на 25».
- [ ] freeze_panes = G4.
- [ ] Длительности (колонка E) = рабочим дням, end_date каждой задачи >= original end_date.
- [ ] Файл открывается в Excel без warning'ов.
