"""Claude API service for project estimation bot."""

import json
import anthropic


def create_client(api_key: str) -> anthropic.Anthropic:
    return anthropic.Anthropic(api_key=api_key)


def _call_claude(client: anthropic.Anthropic, system: str, user_msg: str, max_tokens: int = 8192) -> str:
    with client.messages.stream(
        model="claude-opus-4-20250514",
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": user_msg}],
    ) as stream:
        return stream.get_final_text()


def _parse_json(text: str) -> dict:
    """Strip markdown fences and parse JSON."""
    text = text.strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1] if "\n" in text else text[3:]
        if text.endswith("```"):
            text = text[:-3]
        text = text.strip()
    return json.loads(text)


def summarize_website(client: anthropic.Anthropic, pages_content: str) -> str:
    """Analyze crawled website pages and produce a project description."""
    system = """Ты — опытный бизнес-аналитик в IT-студии. Тебе предоставлено содержимое всех страниц сайта (парсинг).

Твоя задача — подробно проанализировать ВСЕ разделы и функционал сайта и составить детальное описание проекта,
как если бы клиент прислал запрос на разработку аналогичного сайта/приложения.

Описание должно включать:
1. Общее описание проекта (что это за продукт, для кого)
2. Детальный перечень всех разделов и страниц сайта
3. Функционал каждого раздела (что пользователь может делать)
4. Интеграции (платёжные системы, API, внешние сервисы — если видно)
5. Особенности UI/UX (если заметны)
6. Технические особенности (SPA, мобильная адаптация, анимации и т.д.)

Пиши подробно, ВСЕГДА на русском языке, даже если сайт англоязычный. Это описание будет использоваться для составления ТЗ и оценки проекта."""

    return _call_claude(client, system, f"Содержимое страниц сайта:\n\n{pages_content}", max_tokens=16384)


def analyze_request(client: anthropic.Anthropic, request_text: str) -> dict:
    """Analyze if client request has enough info."""
    system = """Ты — опытный проектный менеджер в IT-студии. Твоя задача — определить, достаточно ли информации в клиентском запросе для составления технического задания и оценки проекта.

Если информации достаточно для формирования ТЗ (есть понимание что нужно сделать, какой продукт, основной функционал), верни JSON:
{"sufficient": true}

Если информации недостаточно, сформулируй от 3 до 6 конкретных вопросов, ответы на которые помогут составить ТЗ, декомпозировать задачи и оценить проект. Верни JSON:
{"sufficient": false, "questions": ["вопрос1", "вопрос2", ...]}

ВАЖНО:
- Отвечай ТОЛЬКО валидным JSON без markdown-разметки, без ```json блоков, просто чистый JSON.
- Клиентский запрос может быть на любом языке (русский, английский и др.), но вопросы ВСЕГДА формулируй на русском языке."""

    result = _call_claude(client, system, f"Клиентский запрос:\n\n{request_text}")
    return _parse_json(result)


def generate_spec(client: anthropic.Anthropic, request_text: str, answers: str = None, feedback: str = None, previous_spec: str = None) -> str:
    """Generate technical specification from client request."""
    system = """Ты — опытный проектный менеджер и системный аналитик в IT-студии. Твоя задача — составить чёткое и структурированное техническое задание (ТЗ) на основе клиентского запроса.

ТЗ должно включать:
1. Название проекта
2. Общее описание проекта и его цели
3. Целевая аудитория
4. Функциональные требования (разбитые по модулям/разделам)
5. Нефункциональные требования (производительность, безопасность, масштабируемость)
6. Технологический стек (предложи оптимальный стек, если клиент не указал)
7. Основные экраны/страницы (если применимо)
8. Интеграции (если применимо)
9. Ограничения и допущения

Клиентский запрос может быть на любом языке (русский, английский и др.), но ТЗ ВСЕГДА пиши на русском языке.
Будь конкретен, избегай воды. Формат — структурированный текст с нумерацией."""

    user_msg = f"Клиентский запрос:\n\n{request_text}"
    if answers:
        user_msg += f"\n\nДополнительные ответы на уточняющие вопросы:\n\n{answers}"
    if feedback and previous_spec:
        user_msg += f"\n\nПредыдущая версия ТЗ:\n\n{previous_spec}\n\nЗамечания администратора к ТЗ:\n\n{feedback}"

    return _call_claude(client, system, user_msg)


def decompose_tasks(client: anthropic.Anthropic, spec: str) -> dict:
    """Decompose spec into task modules with min/max day estimates."""
    system = """Ты — опытный проектный менеджер в IT-студии. Твоя задача — декомпозировать техническое задание на конкретные задачи, сгруппированные по модулям, назначить специалистов и дать оценку в ДНЯХ (минимум и максимум).

ПРАВИЛА ОЦЕНКИ:
1. Оценка производится в днях и полуднях (0.5), НЕ в часах
2. Если задача занимает больше 5 дней — разбей её на подзадачи
3. Задачи по фронтенду должны включать работы по интеграции с бэкендом
4. Каждая задача должна быть конкретной и измеримой
5. PM и QA НЕ включаются в список задач — их часы рассчитываются автоматически коэффициентами

СПЕЦИАЛИСТЫ (используй только тех, кто реально нужен):
DevOps, Smart contract, Backend, Frontend, QA, UX/UI дизайнер, Аналитик, Mobile Developer, Data Engineer

СТРУКТУРА МОДУЛЕЙ (группируй задачи логически):
- DevOps: Инфраструктура
- Смарт-контракты (если применимо)
- Backend: Инициализация
- Backend: [Название модуля] (User Module, Trading Module, Payment Module и т.д.)
- Backend: Миграции и документация
- Интеграции
- Frontend: Старт проекта
- Frontend: Страницы
- Frontend: Админ-панель (если нужна)
- Frontend: Адаптив
- QA: Тестирование
- Другие работы

ПРИМЕР ФОРМАТА ЗАДАЧИ:
{"task": "Настройка CI/CD pipelines", "specialist": "DevOps", "comment": "Сборка, линтинг, тесты, деплой для всех репозиториев", "min_days": 1.0, "max_days": 2.0}

Верни ТОЛЬКО валидный JSON (без markdown, без ```json) в формате:
{
  "modules": [
    {
      "name": "DevOps: Инфраструктура",
      "tasks": [
        {"task": "Название задачи", "specialist": "DevOps", "comment": "Детали", "min_days": 0.5, "max_days": 1.0}
      ]
    }
  ]
}

ВАЖНО:
- Названия задач, модулей и комментарии ВСЕГДА на русском языке
- Будь реалистичен в оценках — ориентируйся на опыт реальных проектов
- НЕ указывай PM и QA в задачах — они рассчитываются отдельно
- Если проект предполагает несколько Backend-разработчиков разного уровня (Senior для сложных задач, Middle для типовых), используй "Backend" для обоих — уровень определяется ставкой

РЕАЛЬНЫЙ ПРИМЕР — P2P Platform with OTC (Ethereum Escrow, ~2200ч, 8 специалистов):
P2P-обмен крипто/фиат: KYC, стакан (Buy/Sell/Pit), чат с чек-листом, Escrow-контракт на сделку, диспуты, рейтинг, рефералы, админ-панель, Telegram-бот. Стек: Next.js, NestJS, PostgreSQL, Redis, Ethers.js.

Декомпозиция:
{"modules": [
  {"name": "DevOps: Инфраструктура", "tasks": [
    {"task": "Создание и настройка репозиториев", "specialist": "DevOps", "comment": "Backend, Frontend, Contracts — структура, README, права доступа", "min_days": 1.0, "max_days": 1.5},
    {"task": "Настройка CI/CD pipelines", "specialist": "DevOps", "comment": "Сборка, линтинг, тесты, деплой для всех репозиториев", "min_days": 2.0, "max_days": 3.0},
    {"task": "Настройка PostgreSQL", "specialist": "DevOps", "comment": "Установка, конфигурация, резервное копирование", "min_days": 2.0, "max_days": 3.0},
    {"task": "Настройка Redis", "specialist": "DevOps", "comment": "Кэш, очереди, сессии", "min_days": 2.0, "max_days": 3.0},
    {"task": "Настройка HashiCorp Vault", "specialist": "DevOps", "comment": "Хранение секретов, ключей шифрования", "min_days": 2.0, "max_days": 3.0},
    {"task": "Настройка Nginx и доменов", "specialist": "DevOps", "comment": "Reverse proxy, SSL, rate limiting, DNS", "min_days": 3.0, "max_days": 5.0},
    {"task": "Настройка мониторинга и алертинга", "specialist": "DevOps", "comment": "Prometheus, Grafana, PagerDuty/Slack интеграция", "min_days": 3.0, "max_days": 5.0},
    {"task": "Настройка централизованного логирования", "specialist": "DevOps", "comment": "ELK/Loki стек", "min_days": 2.5, "max_days": 4.0}
  ]},
  {"name": "Smart contract: Escrow", "tasks": [
    {"task": "Разработка EscrowFactory", "specialist": "Smart contract", "comment": "Создание прокси-контрактов, события, access control", "min_days": 1.5, "max_days": 2.5},
    {"task": "Разработка Escrow Implementation", "specialist": "Smart contract", "comment": "Логика депозита, релиза, refund, referral link, resolveDispute, интеграция USDT", "min_days": 3.0, "max_days": 5.0},
    {"task": "Написание unit и integration тестов", "specialist": "Smart contract", "comment": "Покрытие всех функций, happy path, edge cases, gas optimization", "min_days": 2.5, "max_days": 4.0},
    {"task": "Деплой в testnet (Sepolia)", "specialist": "Smart contract", "comment": "Factory + Implementation, верификация в Etherscan", "min_days": 0.5, "max_days": 1.0},
    {"task": "Деплой в mainnet", "specialist": "Smart contract", "comment": "Factory + Implementation, настройка мультисиг/Fireblocks", "min_days": 2.5, "max_days": 4.0},
    {"task": "Аудит безопасности", "specialist": "Smart contract", "comment": "Внешний аудит, исправление найденных проблем", "min_days": 1.0, "max_days": 2.0}
  ]},
  {"name": "Backend: Инициализация", "tasks": [
    {"task": "Настройка инфраструктуры", "specialist": "Backend", "comment": "TypeORM (PostgreSQL), Redis, Bull Queue, валидация", "min_days": 3.0, "max_days": 5.0},
    {"task": "Разработка Shared Layer", "specialist": "Backend", "comment": "Database Service, Cache Service, Queue Service, Auth Service", "min_days": 3.0, "max_days": 5.0},
    {"task": "Разработка Blockchain Service", "specialist": "Backend", "comment": "Ethers.js интеграция, RPC, индексатор событий", "min_days": 3.0, "max_days": 5.0},
    {"task": "Разработка Encryption Service", "specialist": "Backend", "comment": "AES-256 для шифрования реквизитов", "min_days": 2.0, "max_days": 3.0}
  ]},
  {"name": "Backend: User Module", "tasks": [
    {"task": "Создание сущностей", "specialist": "Backend", "comment": "User, Wallet, KycStatus, UserLevel — таблицы, связи, индексы", "min_days": 2.5, "max_days": 4.0},
    {"task": "Разработка API", "specialist": "Backend", "comment": "Регистрация, авторизация (Web3 + JWT), профиль, привязка кошелька", "min_days": 4.0, "max_days": 5.0},
    {"task": "Реализация уровней доступа", "specialist": "Backend", "comment": "Проверка прав, лимиты", "min_days": 3.0, "max_days": 5.0}
  ]},
  {"name": "Backend: Trading Module", "tasks": [
    {"task": "Создание сущностей", "specialist": "Backend", "comment": "Order, Deal, EscrowContract — таблицы, связи", "min_days": 2.0, "max_days": 3.0},
    {"task": "Разработка Order API", "specialist": "Backend", "comment": "CRUD ордеров, стакан, фильтрация, WebSocket для real-time", "min_days": 5.0, "max_days": 5.0},
    {"task": "Разработка Deal API", "specialist": "Backend", "comment": "Инициация, деплой эскроу, обработка депозита, подтверждения", "min_days": 5.0, "max_days": 5.0},
    {"task": "Реализация жизненного цикла сделки", "specialist": "Backend", "comment": "Таймауты, авто-отмена, релиз, комиссия", "min_days": 5.0, "max_days": 5.0},
    {"task": "Реализация режима Яма", "specialist": "Backend", "comment": "Переговорные сделки, фиксация курса", "min_days": 5.0, "max_days": 5.0}
  ]},
  {"name": "Backend: Payment Module", "tasks": [
    {"task": "Создание сущности PaymentDetails", "specialist": "Backend", "comment": "Таблица, шифрование", "min_days": 1.0, "max_days": 1.5},
    {"task": "Разработка API", "specialist": "Backend", "comment": "CRUD реквизитов, передача в сделке после депозита", "min_days": 1.0, "max_days": 2.0}
  ]},
  {"name": "Backend: Dispute Module", "tasks": [
    {"task": "Создание сущностей", "specialist": "Backend", "comment": "Dispute — таблицы, статусы", "min_days": 1.0, "max_days": 2.0},
    {"task": "Разработка API", "specialist": "Backend", "comment": "Создание диспута, добавление доказательств, продление блокировки", "min_days": 3.0, "max_days": 5.0},
    {"task": "Разработка Admin API", "specialist": "Backend", "comment": "Список диспутов, просмотр, принятие решения, вызов resolveDispute", "min_days": 3.0, "max_days": 5.0}
  ]},
  {"name": "Backend: Rating Module", "tasks": [
    {"task": "Создание сущностей", "specialist": "Backend", "comment": "Rating, RatingHistory", "min_days": 1.0, "max_days": 2.0},
    {"task": "Разработка Rating Service", "specialist": "Backend", "comment": "Расчет рейтинга (сделки, отмены, диспуты, объем)", "min_days": 2.0, "max_days": 3.0}
  ]},
  {"name": "Backend: Referral Module", "tasks": [
    {"task": "API генерации реферальной ссылки", "specialist": "Backend", "comment": "Уникальный код, URL", "min_days": 1.0, "max_days": 1.5},
    {"task": "Привязка при регистрации", "specialist": "Backend", "comment": "Парсинг ?ref=, сохранение referrer_id/wallet", "min_days": 0.5, "max_days": 1.0},
    {"task": "Интеграция со сделками", "specialist": "Backend", "comment": "Получение referrer_wallet при создании сделки, передача в контракт", "min_days": 2.0, "max_days": 3.0},
    {"task": "Создание сущности ReferralReward", "specialist": "Backend", "comment": "Хранение истории выплат", "min_days": 1.0, "max_days": 1.5},
    {"task": "API реферальной статистики", "specialist": "Backend", "comment": "Кол-во рефералов, сделок, сумма выплат", "min_days": 2.0, "max_days": 3.0}
  ]},
  {"name": "Backend: Миграции и документация", "tasks": [
    {"task": "Создание миграций", "specialist": "Backend", "comment": "Все таблицы: users, wallets, orders, deals, disputes, rating, referral_rewards", "min_days": 3.0, "max_days": 5.0},
    {"task": "Настройка Swagger/OpenAPI", "specialist": "Backend", "comment": "Автогенерация документации для всех endpoints", "min_days": 2.0, "max_days": 3.0}
  ]},
  {"name": "Интеграции", "tasks": [
    {"task": "KYC интеграция", "specialist": "Backend", "comment": "SumSub/Onfido — вебхуки, SDK, обработка статусов", "min_days": 3.0, "max_days": 5.0},
    {"task": "Чат интеграция", "specialist": "Backend", "comment": "Stream/PubNub — создание каналов, участники, история", "min_days": 4.0, "max_days": 5.0},
    {"task": "Аналитика", "specialist": "Backend", "comment": "Подключение Redash или аналога", "min_days": 1.0, "max_days": 1.5},
    {"task": "Telegram-бот уведомлений", "specialist": "Backend", "comment": "Уведомления о статусах сделок, авторизация", "min_days": 2.0, "max_days": 3.0},
    {"task": "Telegram-канал OTC", "specialist": "Backend", "comment": "Трансляция заявок из Ямы в канал", "min_days": 1.0, "max_days": 1.5},
    {"task": "Unit-тестирование Backend", "specialist": "Backend", "comment": "Jest, покрытие сервисов > 70%", "min_days": 3.0, "max_days": 5.0}
  ]},
  {"name": "Frontend: Старт проекта", "tasks": [
    {"task": "Инициализация проекта", "specialist": "Frontend", "comment": "Next.js, Tailwind, Wagmi, структура, базовая верстка", "min_days": 5.0, "max_days": 5.0}
  ]},
  {"name": "Frontend: Страницы", "tasks": [
    {"task": "Страница Авторизации", "specialist": "Frontend", "comment": "Подключение кошелька, подпись сообщения, JWT", "min_days": 2.0, "max_days": 3.0},
    {"task": "Страница Профиля", "specialist": "Frontend", "comment": "Информация, KYC статус, кошельки, реквизиты, рейтинг", "min_days": 3.0, "max_days": 5.0},
    {"task": "Страница Стакана (Order Book)", "specialist": "Frontend", "comment": "Три колонки (Sell/Pit/Buy), фильтры, WebSocket, карточки ордеров", "min_days": 5.0, "max_days": 5.0},
    {"task": "Страница Создания Ордера", "specialist": "Frontend", "comment": "Тип, направление, валюта, сумма, цена, банки, управление объявлениями", "min_days": 3.0, "max_days": 5.0},
    {"task": "Страница Сделки", "specialist": "Frontend", "comment": "Информационная панель, шаги, депозит, подтверждения, чат, диспут", "min_days": 5.0, "max_days": 5.0},
    {"task": "Страница Яма (Pit)", "specialist": "Frontend", "comment": "Чат, предложение курса, принятие, фиксация", "min_days": 1.0, "max_days": 2.0},
    {"task": "Страница Истории Сделок", "specialist": "Frontend", "comment": "Список, фильтры, детальная информация", "min_days": 1.0, "max_days": 1.5},
    {"task": "Раздел Рефералов в ЛК", "specialist": "Frontend", "comment": "Ссылка, таблица сделок, статистика, интеграция API", "min_days": 4.0, "max_days": 5.0}
  ]},
  {"name": "Frontend: Админ-панель", "tasks": [
    {"task": "Layout и Dashboard", "specialist": "Frontend", "comment": "Сайдбар, навигация, статистика, метрики, управление пользователями, диспуты", "min_days": 5.0, "max_days": 5.0}
  ]},
  {"name": "Frontend: Интеграции и адаптив", "tasks": [
    {"task": "Интеграция WebSocket", "specialist": "Frontend", "comment": "Стакан, уведомления, обработка ошибок", "min_days": 2.0, "max_days": 3.0},
    {"task": "KYC интеграция (фронт)", "specialist": "Frontend", "comment": "Изучение документации, виджет", "min_days": 1.0, "max_days": 2.0},
    {"task": "Чат интеграция (фронт)", "specialist": "Frontend", "comment": "Изучение документации, компоненты", "min_days": 1.0, "max_days": 2.0},
    {"task": "Адаптив под мобильные устройства", "specialist": "Frontend", "comment": "Доработка всех интерфейсов", "min_days": 4.5, "max_days": 5.0}
  ]},
  {"name": "Дизайн", "tasks": [
    {"task": "Сопровождение разработки, доработки UX/UI", "specialist": "Аналитик", "comment": "Прототипирование, доработки по ходу разработки", "min_days": 5.0, "max_days": 5.0}
  ]}
]}"""

    result = _call_claude(client, system, f"Техническое задание:\n\n{spec}", max_tokens=16384)
    return _parse_json(result)
