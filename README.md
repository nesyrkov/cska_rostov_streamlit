# ЦСКА × Ростов Дон — Streamlit + DuckDB

Интерактивный проект для Streamlit Community Cloud по истории встреч ЦСКА и Ростов Дон.

## Данные

Файлы лежат в папке `data/`:

- `Matches.xlsx` — матчи, основной ключ `ID_match`
- `Players.xlsx` — игроки и официальные лица
- `Result.xlsx` — индивидуальная и агрегированная статистика игроков и голкиперов

Основная связь:

```text
matches.ID_match -> player_stats.ID_match
matches.ID_match -> goalkeeper_stats.ID_match
matches.ID_match -> aggregate_players.ID_match
matches.ID_match -> aggregate_goalkeepers.ID_match
```

## Вкладки приложения

- 🎮 Матчи
- 🧍 Игроки
- 📊 Статистика игроков
- 🧤 Статистика голкиперов
- 🦆 DuckDB

## Локальный запуск

```bash
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
# .venv\Scripts\activate   # Windows
pip install -r requirements.txt
streamlit run app.py
```

## Деплой на Streamlit Community Cloud

1. Создайте GitHub-репозиторий.
2. Загрузите файлы проекта в репозиторий.
3. В Streamlit Community Cloud выберите репозиторий.
4. Main file path: `app.py`.
5. Deploy.

## Примечание по данным

В `Result.xlsx` есть лист `AggregateStatisicsGoalkeepers` с опечаткой в слове `Statistics`. Приложение поддерживает это название автоматически.
