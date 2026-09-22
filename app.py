from __future__ import annotations

from pathlib import Path
import re

import duckdb
import pandas as pd
import plotly.express as px
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
MATCHES_FILE = DATA_DIR / "Matches.xlsx"
PLAYERS_FILE = DATA_DIR / "Players.xlsx"
RESULT_FILE = DATA_DIR / "Result.xlsx"

CSKA = "ЦСКА"
ROSTOV = "Ростов Дон"


st.set_page_config(
    page_title="ЦСКА vs Ростов | Streamlit + DuckDB",
    page_icon="⚽",
    layout="wide",
)


CSS = """
<style>
:root {
    --cska-red: #D50032;
    --cska-blue: #0033A0;
    --rostov-yellow: #FFD200;
    --rostov-black: #111111;
    --card-bg: #171B24;
}

.block-container {
    padding-top: 1.5rem;
    padding-bottom: 2rem;
}

.main-title {
    border-radius: 24px;
    padding: 28px 24px;
    margin-bottom: 18px;
    color: #fff;
    background:
        radial-gradient(circle at 8% 15%, rgba(255,210,0,.75), transparent 22%),
        radial-gradient(circle at 92% 25%, rgba(213,0,50,.70), transparent 24%),
        linear-gradient(110deg, #0033A0 0%, #D50032 40%, #FFD200 72%, #111111 100%);
    box-shadow: 0 20px 50px rgba(0,0,0,.40);
}

.main-title h1 {
    font-size: 44px;
    line-height: 1.05;
    margin: 0 0 8px 0;
    font-weight: 950;
    letter-spacing: -0.03em;
}

.main-title p {
    font-size: 16px;
    margin: 0;
    opacity: .92;
}

.score-card {
    border-radius: 22px;
    padding: 18px;
    background: var(--card-bg);
    border: 1px solid rgba(255,255,255,.08);
    box-shadow: 0 12px 28px rgba(0,0,0,.28);
}

.cska-card {
    border-left: 8px solid var(--cska-red);
    background: linear-gradient(135deg, rgba(213,0,50,.92), rgba(0,51,160,.92));
}

.rostov-card {
    border-left: 8px solid var(--rostov-yellow);
    background: linear-gradient(135deg, rgba(255,210,0,.95), rgba(17,17,17,.92));
    color: white;
}

.neutral-card {
    border-left: 8px solid rgba(255,255,255,.35);
}

.card-title {
    font-size: 13px;
    opacity: .78;
    text-transform: uppercase;
    letter-spacing: .06em;
    margin-bottom: 4px;
}

.card-value {
    font-size: 32px;
    font-weight: 900;
    line-height: 1.1;
}

.small-note {
    opacity: .72;
    font-size: 13px;
}

[data-testid="stMetricValue"] {
    font-weight: 900;
}
</style>
"""

st.markdown(CSS, unsafe_allow_html=True)


@st.cache_data(show_spinner="Загружаю Excel-файлы...")
def read_excel_data() -> dict[str, pd.DataFrame]:
    """Read all project files from /data. Designed for Streamlit Community Cloud."""
    missing = [str(p.relative_to(BASE_DIR)) for p in [MATCHES_FILE, PLAYERS_FILE, RESULT_FILE] if not p.exists()]
    if missing:
        raise FileNotFoundError("Не найдены файлы: " + ", ".join(missing))

    matches = pd.read_excel(MATCHES_FILE, sheet_name=0, engine="openpyxl")
    players = pd.read_excel(PLAYERS_FILE, sheet_name="Лист1", engine="openpyxl")
    officials = pd.read_excel(PLAYERS_FILE, sheet_name="Лист2", engine="openpyxl")

    result_xls = pd.ExcelFile(RESULT_FILE, engine="openpyxl")
    player_stats = pd.read_excel(result_xls, sheet_name="A")
    goalkeeper_stats = pd.read_excel(result_xls, sheet_name="B")
    aggregate_players = pd.read_excel(result_xls, sheet_name="AggregateStatisticsPlayers")

    # In the source workbook the goalkeeper aggregate sheet is intentionally supported
    # with both spellings: AggregateStatisicsGoalkeepers / AggregateStatisticsGoalkeepers.
    gk_sheet = next(
        s for s in result_xls.sheet_names
        if s.lower().replace("statistics", "statisics") == "aggregatestatisicsgoalkeepers".lower()
    )
    aggregate_goalkeepers = pd.read_excel(result_xls, sheet_name=gk_sheet)

    tables = {
        "matches": clean_matches(matches),
        "players": clean_base_table(players),
        "officials": clean_base_table(officials),
        "player_stats": clean_base_table(player_stats),
        "goalkeeper_stats": clean_base_table(goalkeeper_stats),
        "aggregate_players": clean_base_table(aggregate_players),
        "aggregate_goalkeepers": clean_base_table(aggregate_goalkeepers),
    }

    tables["players"]["player_key"] = tables["players"]["Players"].map(normalize_name)
    tables["players"]["No_num"] = pd.to_numeric(tables["players"].get("No"), errors="coerce")
    tables["players_lookup"] = make_players_lookup(tables["players"])

    # Individual player sheet uses a blank row as a separator between home and away teams.
    tables["player_stats"] = add_team_from_blank_separator(
        tables["player_stats"], tables["matches"], name_col="Players"
    )
    tables["player_stats"] = tables["player_stats"][tables["player_stats"]["Players"].notna()].copy()
    tables["player_stats"]["player_key"] = tables["player_stats"]["Players"].map(normalize_name)

    # Goalkeeper sheet has no blank separator, so the home/away split is inferred
    # from source order and roster matches by name/number.
    tables["goalkeeper_stats"] = add_team_to_goalkeepers(
        tables["goalkeeper_stats"], tables["matches"], tables["players"]
    )
    tables["goalkeeper_stats"]["player_key"] = tables["goalkeeper_stats"]["goalkeeper"].map(normalize_name)

    tables["aggregate_players"] = add_team_from_match_order(
        tables["aggregate_players"], tables["matches"]
    )
    tables["aggregate_goalkeepers"] = add_team_from_match_order(
        tables["aggregate_goalkeepers"], tables["matches"]
    )

    tables["player_stats_enriched"] = enrich_player_stats(tables)
    tables["goalkeeper_stats_enriched"] = enrich_goalkeeper_stats(tables)

    return tables


def clean_base_table(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if "ID_match" in df.columns:
        df["ID_match"] = pd.to_numeric(df["ID_match"], errors="coerce").astype("Int64")
    return df


def clean_matches(df: pd.DataFrame) -> pd.DataFrame:
    df = clean_base_table(df)
    if " Score" in df.columns:
        df = df.rename(columns={" Score": "Score"})
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df["Date_text"] = df["Date"].dt.strftime("%d.%m.%Y %H:%M")
    if "Score" in df.columns:
        goals = df["Score"].astype(str).str.extract(r"(?P<Home_Goals>\d+)\s*:\s*(?P<Away_Goals>\d+)")
        df["Home_Goals"] = pd.to_numeric(goals["Home_Goals"], errors="coerce")
        df["Away_Goals"] = pd.to_numeric(goals["Away_Goals"], errors="coerce")
    df["Winner"] = df.apply(match_winner, axis=1)
    return df


def normalize_name(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).lower().replace("ё", "е")
    text = re.sub(r"[^а-яa-z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def match_winner(row: pd.Series) -> str:
    result = str(row.get("Result", ""))
    if CSKA.lower() in result.lower():
        return CSKA
    if "ростов" in result.lower():
        return ROSTOV
    home_goals = row.get("Home_Goals")
    away_goals = row.get("Away_Goals")
    if pd.notna(home_goals) and pd.notna(away_goals):
        if home_goals > away_goals:
            return str(row.get("Home_Team", ""))
        if away_goals > home_goals:
            return str(row.get("Away_Team", ""))
    return "Ничья"


def make_players_lookup(players: pd.DataFrame) -> pd.DataFrame:
    """One roster row per player/team for safer joins after team is already known."""
    cols = [
        c for c in ["player_key", "Team", "Position", "Citizenship", "Height", "Age", "No", "No_num"]
        if c in players.columns
    ]
    lookup = players[cols].dropna(subset=["player_key", "Team"]).copy()
    return lookup.drop_duplicates(subset=["player_key", "Team"], keep="first")


def attach_match_context_and_team(df: pd.DataFrame, matches: pd.DataFrame) -> pd.DataFrame:
    if "ID_match" not in df.columns:
        return df
    out = df.copy()
    out = out.merge(
        matches[["ID_match", "Date", "Date_text", "Season", "Tournament", "Stage", "Home_Team", "Away_Team", "Score", "Winner"]],
        on="ID_match",
        how="left",
    )
    if "side_number" in out.columns:
        side = pd.to_numeric(out["side_number"], errors="coerce")
    else:
        side = pd.Series(pd.NA, index=out.index)

    out["Team"] = "Не определено"

    mask_home = side.eq(1)
    mask_away = side.eq(2)

    if "Home_Team" in out.columns:
        out.loc[mask_home, "Team"] = out.loc[mask_home, "Home_Team"].fillna("Не определено")

    if "Away_Team" in out.columns:
        out.loc[mask_away, "Team"] = out.loc[mask_away, "Away_Team"].fillna("Не определено")

    return out


def add_team_from_blank_separator(df: pd.DataFrame, matches: pd.DataFrame, name_col: str) -> pd.DataFrame:
    """Assign team by row blocks: before blank separator = home, after separator = away."""
    if "ID_match" not in df.columns or name_col not in df.columns:
        return df
    out = df.copy()
    out["side_number"] = pd.NA

    for _, group in out.groupby("ID_match", sort=False):
        side = 1
        for idx, row in group.iterrows():
            if pd.isna(row.get(name_col)):
                out.at[idx, "side_number"] = pd.NA
                side = 2
            else:
                out.at[idx, "side_number"] = side

    out["side_number"] = pd.to_numeric(out["side_number"], errors="coerce").astype("Int64")
    return attach_match_context_and_team(out, matches)


def roster_candidates(players: pd.DataFrame, name: str, number) -> set[str]:
    key = normalize_name(name)
    no = pd.to_numeric(number, errors="coerce")
    if not key or pd.isna(no):
        return set()
    mask = (players["player_key"] == key) & (players["No_num"] == float(no))
    return set(players.loc[mask, "Team"].dropna().astype(str).tolist())


def add_team_to_goalkeepers(df: pd.DataFrame, matches: pd.DataFrame, players: pd.DataFrame) -> pd.DataFrame:
    """Infer home/away goalkeeper blocks for each match using roster name+number candidates."""
    if "ID_match" not in df.columns:
        return df
    out = df.copy()
    out["side_number"] = pd.NA
    match_lookup = matches.set_index("ID_match")

    for match_id, group in out.groupby("ID_match", sort=False):
        if pd.isna(match_id) or match_id not in match_lookup.index:
            continue
        home = str(match_lookup.loc[match_id, "Home_Team"])
        away = str(match_lookup.loc[match_id, "Away_Team"])
        indices = list(group.index)
        n = len(indices)
        if n == 0:
            continue

        candidates = [
            roster_candidates(players, out.at[idx, "goalkeeper"], out.at[idx, "no"])
            for idx in indices
        ]

        if n == 1:
            best_split = 1
        else:
            best_score = -1
            best_split = max(1, n // 2)
            for split in range(1, n):
                score = 0.0
                for pos, teams in enumerate(candidates):
                    assigned = home if pos < split else away
                    if assigned in teams:
                        score += 1.0
                    elif not teams:
                        score += 0.05
                if score > best_score:
                    best_score = score
                    best_split = split

        for pos, idx in enumerate(indices):
            out.at[idx, "side_number"] = 1 if pos < best_split else 2

    out["side_number"] = pd.to_numeric(out["side_number"], errors="coerce").astype("Int64")
    return attach_match_context_and_team(out, matches)


def add_team_from_match_order(df: pd.DataFrame, matches: pd.DataFrame) -> pd.DataFrame:
    """Aggregate sheets have one row per team per ID_match. Row 1 = home team, row 2 = away team."""
    if "ID_match" not in df.columns:
        return df
    out = df.copy()
    out["side_number"] = out.groupby("ID_match").cumcount() + 1
    out = out.merge(
        matches[["ID_match", "Date", "Season", "Tournament", "Stage", "Home_Team", "Away_Team", "Score", "Winner"]],
        on="ID_match",
        how="left",
    )
    out["Team"] = out.apply(
        lambda r: r["Home_Team"] if r["side_number"] == 1 else r["Away_Team"], axis=1
    )
    return out


def split_goal_attempt(value) -> tuple[float | None, float | None]:
    if pd.isna(value):
        return None, None
    text = str(value).strip()
    match = re.match(r"^(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)$", text)
    if not match:
        return None, None
    return float(match.group(1)), float(match.group(2))


def add_goal_columns(df: pd.DataFrame, source_col: str = "Goals", prefix: str = "goals") -> pd.DataFrame:
    out = df.copy()
    if source_col in out.columns:
        pairs = out[source_col].map(split_goal_attempt)
        out[f"{prefix}_made"] = pairs.map(lambda x: x[0])
        out[f"{prefix}_attempts"] = pairs.map(lambda x: x[1])
        out[f"{prefix}_pct_calc"] = out.apply(
            lambda r: round(r[f"{prefix}_made"] / r[f"{prefix}_attempts"] * 100, 1)
            if pd.notna(r[f"{prefix}_made"]) and pd.notna(r[f"{prefix}_attempts"]) and r[f"{prefix}_attempts"]
            else None,
            axis=1,
        )
    return out


def enrich_player_stats(tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    con = duckdb.connect(database=":memory:")
    con.register("player_stats", tables["player_stats"])
    con.register("players_lookup", tables["players_lookup"])
    df = con.execute(
        """
        SELECT
            ps.*,
            p.Position,
            p.Citizenship,
            p.Height,
            p.Age
        FROM player_stats ps
        LEFT JOIN players_lookup p
            ON ps.player_key = p.player_key AND ps.Team = p.Team
        """
    ).df()
    df = add_goal_columns(df, source_col="Goals", prefix="goals")
    return df


def enrich_goalkeeper_stats(tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    con = duckdb.connect(database=":memory:")
    con.register("goalkeeper_stats", tables["goalkeeper_stats"])
    con.register("players_lookup", tables["players_lookup"])
    df = con.execute(
        """
        SELECT
            gs.*,
            p.Position,
            p.Citizenship,
            p.Height,
            p.Age
        FROM goalkeeper_stats gs
        LEFT JOIN players_lookup p
            ON gs.player_key = p.player_key AND gs.Team = p.Team
        """
    ).df()
    df = add_goal_columns(df, source_col="total", prefix="saves")
    return df


def team_card_class(team: str) -> str:
    text = str(team).lower()
    if "цска" in text:
        return "cska-card"
    if "ростов" in text:
        return "rostov-card"
    return "neutral-card"


def render_card(title: str, value: str, subtitle: str = "", card_class: str = "neutral-card") -> None:
    st.markdown(
        f"""
        <div class="score-card {card_class}">
            <div class="card-title">{title}</div>
            <div class="card-value">{value}</div>
            <div class="small-note">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def filter_by_sidebar(matches: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Фильтры")
    filtered = matches.copy()

    if "Season" in filtered.columns:
        seasons = sorted(filtered["Season"].dropna().astype(str).unique())
        selected = st.sidebar.multiselect("Сезон", seasons, default=seasons)
        filtered = filtered[filtered["Season"].astype(str).isin(selected)]

    if "Tournament" in filtered.columns:
        tournaments = sorted(filtered["Tournament"].dropna().astype(str).unique())
        selected = st.sidebar.multiselect("Турнир", tournaments, default=tournaments)
        filtered = filtered[filtered["Tournament"].astype(str).isin(selected)]

    winner_options = ["Все", CSKA, ROSTOV]
    winner = st.sidebar.radio("Победитель", winner_options, horizontal=False)
    if winner != "Все":
        filtered = filtered[filtered["Winner"] == winner]

    return filtered


def display_matches_tab(matches: pd.DataFrame, aggregate_players: pd.DataFrame) -> None:
    st.subheader("Матчи")

    if matches.empty:
        st.info("По выбранным фильтрам матчей не найдено.")
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Матчей", len(matches))
    c2.metric("Побед ЦСКА", int((matches["Winner"] == CSKA).sum()))
    c3.metric("Побед Ростова", int((matches["Winner"] == ROSTOV).sum()))
    c4.metric("Голов всего", int(matches["Number_of_Goals"].fillna(0).sum()))

    st.dataframe(
        matches[[
            "ID_match", "Date_text", "Season", "Tournament", "Stage", "Home_Team",
            "Away_Team", "First_Half_Score", "Score", "Number_of_Goals", "Result", "Arena"
        ]],
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("### Детализация выбранного матча")
    match_options = matches.sort_values("Date", ascending=False).copy()
    match_options["label"] = match_options.apply(
        lambda r: f"#{int(r['ID_match'])} | {r.get('Date_text', '')} | {r['Home_Team']} {r['Score']} {r['Away_Team']}",
        axis=1,
    )
    selected_label = st.selectbox("Выберите матч", match_options["label"])
    selected_id = int(match_options.loc[match_options["label"] == selected_label, "ID_match"].iloc[0])
    selected_match = match_options[match_options["ID_match"] == selected_id].iloc[0]

    left, middle, right = st.columns([1, .75, 1])
    with left:
        render_card(
            "Хозяева",
            str(selected_match["Home_Team"]),
            f"Голы: {int(selected_match['Home_Goals']) if pd.notna(selected_match['Home_Goals']) else '-'}",
            team_card_class(selected_match["Home_Team"]),
        )
    with middle:
        render_card(
            "Счёт",
            str(selected_match["Score"]),
            f"1-й тайм: {selected_match.get('First_Half_Score', '-')}",
            "neutral-card",
        )
    with right:
        render_card(
            "Гости",
            str(selected_match["Away_Team"]),
            f"Голы: {int(selected_match['Away_Goals']) if pd.notna(selected_match['Away_Goals']) else '-'}",
            team_card_class(selected_match["Away_Team"]),
        )

    match_team_stats = aggregate_players[aggregate_players["ID_match"] == selected_id].copy()
    if not match_team_stats.empty:
        st.markdown("### Командная статистика по матчу")
        visible_cols = [
            c for c in ["Team", "Goals", "Percentage", "Assist", "Turnover", "Total_turnovers", "Steal/Interception", "2-minutes"]
            if c in match_team_stats.columns
        ]
        st.dataframe(match_team_stats[visible_cols], use_container_width=True, hide_index=True)
    else:
        st.info("Для выбранного матча нет агрегированной статистики игроков.")


def display_players_tab(players: pd.DataFrame, officials: pd.DataFrame) -> None:
    st.subheader("Игроки и штаб")

    col1, col2, col3 = st.columns(3)
    with col1:
        team = st.selectbox("Команда", ["Все"] + sorted(players["Team"].dropna().unique().tolist()))
    with col2:
        position = st.selectbox("Амплуа", ["Все"] + sorted(players["Position"].dropna().unique().tolist()))
    with col3:
        search = st.text_input("Поиск игрока")

    filtered = players.copy()
    if team != "Все":
        filtered = filtered[filtered["Team"] == team]
    if position != "Все":
        filtered = filtered[filtered["Position"] == position]
    if search:
        filtered = filtered[filtered["Players"].astype(str).str.contains(search, case=False, na=False)]

    c1, c2, c3 = st.columns(3)
    c1.metric("Игроков", len(filtered))
    c2.metric("Вратарей", int((filtered["Position"].astype(str).str.lower() == "вратарь").sum()))
    if "Height" in filtered.columns:
        c3.metric("Средний рост", round(filtered["Height"].dropna().mean(), 1) if filtered["Height"].notna().any() else "-")

    st.dataframe(filtered.drop(columns=["player_key"], errors="ignore"), use_container_width=True, hide_index=True)

    with st.expander("Официальные лица"):
        st.dataframe(officials, use_container_width=True, hide_index=True)


def stats_filters(df: pd.DataFrame, key_prefix: str) -> pd.DataFrame:
    filtered = df.copy()
    c1, c2, c3 = st.columns(3)
    with c1:
        team = st.selectbox("Команда", ["Все"] + sorted(filtered["Team"].dropna().astype(str).unique()), key=f"{key_prefix}_team")
    with c2:
        seasons = ["Все"] + sorted(filtered["Season"].dropna().astype(str).unique())
        season = st.selectbox("Сезон", seasons, key=f"{key_prefix}_season")
    with c3:
        match_ids = ["Все"] + [str(int(x)) for x in sorted(filtered["ID_match"].dropna().unique())]
        match_id = st.selectbox("ID матча", match_ids, key=f"{key_prefix}_match")

    if team != "Все":
        filtered = filtered[filtered["Team"].astype(str) == team]
    if season != "Все":
        filtered = filtered[filtered["Season"].astype(str) == season]
    if match_id != "Все":
        filtered = filtered[filtered["ID_match"].astype("Int64") == int(match_id)]

    return filtered


def display_player_stats_tab(player_stats: pd.DataFrame) -> None:
    st.subheader("Статистика игроков")
    filtered = stats_filters(player_stats, "players")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Записей", len(filtered))
    c2.metric("Игроков", filtered["Players"].nunique())
    c3.metric("Голов", int(filtered["goals_made"].fillna(0).sum()))
    c4.metric("Бросков", int(filtered["goals_attempts"].fillna(0).sum()))

    chart_data = (
        filtered.dropna(subset=["goals_made"])
        .groupby(["Players", "Team"], as_index=False)["goals_made"]
        .sum()
        .sort_values("goals_made", ascending=False)
        .head(15)
    )
    if not chart_data.empty:
        fig = px.bar(
            chart_data,
            x="Players",
            y="goals_made",
            color="Team",
            title="Топ-15 игроков по голам в выбранной выборке",
            labels={"goals_made": "Голы", "Players": "Игрок"},
            color_discrete_map={CSKA: "#D50032", ROSTOV: "#FFD200", "Не определено": "#777777"},
        )
        fig.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(fig, use_container_width=True)

    visible_cols = [
        "ID_match", "Date_text", "Season", "Tournament", "Home_Team", "Away_Team", "Score",
        "Team", "Player_number/Jersey_number", "Players", "Position", "Goals", "Percentage",
        "Assist", "Turnover", "Total_turnovers", "Steal/Interception", "2-minutes", "Playing_time"
    ]
    visible_cols = [c for c in visible_cols if c in filtered.columns]
    st.dataframe(filtered[visible_cols], use_container_width=True, hide_index=True)


def display_goalkeeper_stats_tab(goalkeeper_stats: pd.DataFrame, aggregate_goalkeepers: pd.DataFrame) -> None:
    st.subheader("Статистика голкиперов")
    filtered = stats_filters(goalkeeper_stats, "gk")

    c1, c2, c3 = st.columns(3)
    c1.metric("Записей", len(filtered))
    c2.metric("Голкиперов", filtered["goalkeeper"].nunique())
    if "save_pct" in filtered.columns:
        c3.metric("Средний % сейвов", round(filtered["save_pct"].dropna().mean(), 1) if filtered["save_pct"].notna().any() else "-")

    chart_data = (
        filtered.dropna(subset=["save_pct"])
        .groupby(["goalkeeper", "Team"], as_index=False)["save_pct"]
        .mean()
        .sort_values("save_pct", ascending=False)
        .head(15)
    )
    if not chart_data.empty:
        fig = px.bar(
            chart_data,
            x="goalkeeper",
            y="save_pct",
            color="Team",
            title="Топ голкиперов по среднему проценту сейвов",
            labels={"save_pct": "% сейвов", "goalkeeper": "Голкипер"},
            color_discrete_map={CSKA: "#D50032", ROSTOV: "#FFD200", "Не определено": "#777777"},
        )
        fig.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(fig, use_container_width=True)

    visible_cols = [
        "ID_match", "Date_text", "Season", "Tournament", "Home_Team", "Away_Team", "Score",
        "Team", "no", "goalkeeper", "total", "save_pct", "six_m", "six_m_save_pct",
        "wing", "wing_save_pct", "close_range", "close_range_save_pct", "long_range", "long_range_save_pct",
        "seven_m", "seven_m_save_pct"
    ]
    visible_cols = [c for c in visible_cols if c in filtered.columns]
    st.dataframe(filtered[visible_cols], use_container_width=True, hide_index=True)

    with st.expander("Командная агрегированная статистика голкиперов"):
        st.dataframe(aggregate_goalkeepers, use_container_width=True, hide_index=True)


def display_duckdb_tab(tables: dict[str, pd.DataFrame]) -> None:
    st.subheader("DuckDB-связи")
    st.markdown(
        """
        Основная связь проекта: `matches.ID_match` → статистические таблицы `ID_match`.
        Ниже можно выполнить безопасный SELECT-запрос к зарегистрированным таблицам.
        """
    )
    st.code(
        """
Доступные таблицы:
- matches
- players
- officials
- player_stats
- goalkeeper_stats
- aggregate_players
- aggregate_goalkeepers
- player_stats_enriched
- goalkeeper_stats_enriched
        """.strip()
    )

    default_sql = """
SELECT
    m.ID_match,
    m.Date_text,
    m.Home_Team,
    m.Away_Team,
    m.Score,
    m.Winner,
    COUNT(ps.Players) AS player_stat_rows
FROM matches m
LEFT JOIN player_stats ps ON m.ID_match = ps.ID_match
GROUP BY 1,2,3,4,5,6
ORDER BY m.ID_match;
""".strip()

    sql = st.text_area("SQL", default_sql, height=180)
    if not sql.strip().lower().startswith("select"):
        st.error("Для безопасности в демо разрешены только SELECT-запросы.")
        return

    con = duckdb.connect(database=":memory:")
    for name, table in tables.items():
        con.register(name, table)
    try:
        result = con.execute(sql).df()
        st.dataframe(result, use_container_width=True, hide_index=True)
    except Exception as exc:
        st.error(f"Ошибка SQL: {exc}")


def main() -> None:
    st.markdown(
        """
        <div class="main-title">
            <h1>ЦСКА × Ростов Дон</h1>
            <p>История встреч, игроки, статистика полевых игроков и голкиперов · Streamlit + DuckDB</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    try:
        tables = read_excel_data()
    except Exception as exc:
        st.error("Не удалось загрузить данные из папки data/.")
        st.exception(exc)
        st.stop()

    filtered_matches = filter_by_sidebar(tables["matches"])
    allowed_ids = set(filtered_matches["ID_match"].dropna().astype(int).tolist())

    player_stats = tables["player_stats_enriched"]
    goalkeeper_stats = tables["goalkeeper_stats_enriched"]
    aggregate_players = tables["aggregate_players"]
    aggregate_goalkeepers = tables["aggregate_goalkeepers"]

    if allowed_ids:
        player_stats = player_stats[player_stats["ID_match"].astype("Int64").isin(allowed_ids)]
        goalkeeper_stats = goalkeeper_stats[goalkeeper_stats["ID_match"].astype("Int64").isin(allowed_ids)]
        aggregate_players = aggregate_players[aggregate_players["ID_match"].astype("Int64").isin(allowed_ids)]
        aggregate_goalkeepers = aggregate_goalkeepers[aggregate_goalkeepers["ID_match"].astype("Int64").isin(allowed_ids)]

    tab_matches, tab_players, tab_player_stats, tab_gk_stats, tab_sql = st.tabs([
        "🎮 Матчи",
        "🧍 Игроки",
        "📊 Статистика игроков",
        "🧤 Статистика голкиперов",
        "🦆 DuckDB",
    ])

    with tab_matches:
        display_matches_tab(filtered_matches, aggregate_players)
    with tab_players:
        display_players_tab(tables["players"], tables["officials"])
    with tab_player_stats:
        display_player_stats_tab(player_stats)
    with tab_gk_stats:
        display_goalkeeper_stats_tab(goalkeeper_stats, aggregate_goalkeepers)
    with tab_sql:
        display_duckdb_tab(tables)


if __name__ == "__main__":
    main()
