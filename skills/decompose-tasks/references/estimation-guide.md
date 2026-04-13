# Estimation Reference Guide

## Calibration example: P2P Platform with OTC

**Project**: P2P crypto/fiat exchange with Ethereum Escrow (~2200 hours, 8 specialists)
**Features**: KYC, order book (Buy/Sell/Pit), chat with checklist, Escrow smart contract, disputes, ratings, referrals, admin panel, Telegram bot
**Stack**: Next.js, NestJS, PostgreSQL, Redis, Ethers.js

### Reference decomposition

```json
{"modules": [
  {"name": "DevOps: Инфраструктура", "tasks": [
    {"task": "Создание и настройка репозиториев", "specialist": "DevOps", "comment": "Backend, Frontend, Contracts — структура, README, права доступа", "min_days": 1.0, "max_days": 1.5, "phase": "mvp"},
    {"task": "Настройка CI/CD pipelines", "specialist": "DevOps", "comment": "Сборка, линтинг, тесты, деплой для всех репозиториев", "min_days": 2.0, "max_days": 3.0, "phase": "mvp"},
    {"task": "Настройка PostgreSQL", "specialist": "DevOps", "comment": "Установка, конфигурация, резервное копирование", "min_days": 2.0, "max_days": 3.0, "phase": "mvp"},
    {"task": "Настройка Redis", "specialist": "DevOps", "comment": "Кэш, очереди, сессии", "min_days": 2.0, "max_days": 3.0, "phase": "mvp"},
    {"task": "Настройка HashiCorp Vault", "specialist": "DevOps", "comment": "Хранение секретов, ключей шифрования", "min_days": 2.0, "max_days": 3.0, "phase": "mvp"},
    {"task": "Настройка Nginx и доменов", "specialist": "DevOps", "comment": "Reverse proxy, SSL, rate limiting, DNS", "min_days": 3.0, "max_days": 5.0, "phase": "mvp"},
    {"task": "Настройка мониторинга и алертинга", "specialist": "DevOps", "comment": "Prometheus, Grafana, PagerDuty/Slack интеграция", "min_days": 3.0, "max_days": 5.0, "phase": "post-mvp"},
    {"task": "Настройка централизованного логирования", "specialist": "DevOps", "comment": "ELK/Loki стек", "min_days": 2.5, "max_days": 4.0, "phase": "post-mvp"}
  ]},
  {"name": "Smart contract: Escrow", "tasks": [
    {"task": "Разработка EscrowFactory", "specialist": "Smart contract", "comment": "Создание прокси-контрактов, события, access control", "min_days": 1.5, "max_days": 2.5, "phase": "mvp"},
    {"task": "Разработка Escrow Implementation", "specialist": "Smart contract", "comment": "Логика депозита, релиза, refund, referral link, resolveDispute, интеграция USDT", "min_days": 3.0, "max_days": 5.0, "phase": "mvp"},
    {"task": "Написание unit и integration тестов", "specialist": "Smart contract", "comment": "Покрытие всех функций, happy path, edge cases, gas optimization", "min_days": 2.5, "max_days": 4.0, "phase": "mvp"},
    {"task": "Деплой в testnet (Sepolia)", "specialist": "Smart contract", "comment": "Factory + Implementation, верификация в Etherscan", "min_days": 0.5, "max_days": 1.0, "phase": "mvp"},
    {"task": "Деплой в mainnet", "specialist": "Smart contract", "comment": "Factory + Implementation, настройка мультисиг/Fireblocks", "min_days": 2.5, "max_days": 4.0, "phase": "mvp"},
    {"task": "Аудит безопасности", "specialist": "Smart contract", "comment": "Внешний аудит, исправление найденных проблем", "min_days": 1.0, "max_days": 2.0, "phase": "mvp"}
  ]},
  {"name": "Backend: Инициализация", "tasks": [
    {"task": "Настройка инфраструктуры", "specialist": "Backend", "comment": "TypeORM (PostgreSQL), Redis, Bull Queue, валидация", "min_days": 3.0, "max_days": 5.0, "phase": "mvp"},
    {"task": "Разработка Shared Layer", "specialist": "Backend", "comment": "Database Service, Cache Service, Queue Service, Auth Service", "min_days": 3.0, "max_days": 5.0, "phase": "mvp"},
    {"task": "Разработка Blockchain Service", "specialist": "Backend", "comment": "Ethers.js интеграция, RPC, индексатор событий", "min_days": 3.0, "max_days": 5.0, "phase": "mvp"},
    {"task": "Разработка Encryption Service", "specialist": "Backend", "comment": "AES-256 для шифрования реквизитов", "min_days": 2.0, "max_days": 3.0, "phase": "mvp"}
  ]}
]}
```

## Estimation principles

1. **Infrastructure tasks** (DevOps, project setup) are often underestimated — account for configuration, testing, and documentation
2. **Integration tasks** have high variance — simple API = 1-2 days, complex third-party with webhooks = 3-5 days
3. **Frontend tasks** include API integration time, not just UI
4. **Smart contract tasks** include testing time (critical for security)
5. **"Настройка" tasks** include both setup AND verification/testing
6. **Admin panels** are typically 30-50% of the main frontend effort
7. **Mobile adaptive** is 15-25% of total frontend work if designed mobile-first, 30-40% if desktop-first
