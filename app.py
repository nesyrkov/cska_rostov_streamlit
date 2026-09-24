from __future__ import annotations

from pathlib import Path
import re

import duckdb
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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

    tables["players"] = normalize_team_column(tables["players"])
    tables["officials"] = normalize_team_column(tables["officials"])
    tables["players"]["player_key"] = tables["players"]["Players"].map(normalize_name)
    tables["players"]["No_num"] = pd.to_numeric(tables["players"].get("No"), errors="coerce")
    tables["players_lookup"] = make_players_lookup(tables["players"])

    # The updated Result.xlsx already contains Team in sheets A and B.
    # Use Team from the workbook as the primary source for player affiliation
    # at the moment of each match. The old side_number logic is kept only
    # as a fallback if Team is absent or empty.
    tables["player_stats"] = prepare_individual_stats(
        tables["player_stats"], tables["matches"], name_col="Players"
    )
    tables["player_stats"] = tables["player_stats"][tables["player_stats"]["Players"].notna()].copy()
    tables["player_stats"]["player_key"] = tables["player_stats"]["Players"].map(normalize_name)

    tables["goalkeeper_stats"] = prepare_individual_stats(
        tables["goalkeeper_stats"], tables["matches"], name_col="goalkeeper"
    )
    tables["goalkeeper_stats"] = tables["goalkeeper_stats"][tables["goalkeeper_stats"]["goalkeeper"].notna()].copy()
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
    for col in ["Home_Team", "Away_Team"]:
        if col in df.columns:
            df[col] = df[col].map(normalize_team_name)
    if " Score" in df.columns:
        df = df.rename(columns={" Score": "Score"})
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df["Date_text"] = df["Date"].dt.strftime("%d.%m.%Y %H:%M")
    if "Score" in df.columns:
        goals = df["Score"].astype(str).str.extract(r"(?P<Home_Goals>\d+)\s*:\s*(?P<Away_Goals>\d+)")
        df["Home_Goals"] = pd.to_numeric(goals["Home_Goals"], errors="coerce")
        df["Away_Goals"] = pd.to_numeric(goals["Away_Goals"], errors="coerce")
    for col in ["Number_of_Goals", "Capacity", "Attendance_Rate"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    if "Arena" in df.columns:
        df["Arena"] = df["Arena"].astype(str).str.strip()
        df["Arena_City"] = df["Arena"].map(arena_city)
        df["Arena_short"] = df["Arena"].map(short_arena_name)
    if "Referees" in df.columns:
        df["Referees"] = df["Referees"].map(normalize_referee_pair)
    df["Winner"] = df.apply(match_winner, axis=1)
    return df


def normalize_name(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).lower().replace("ё", "е")
    text = re.sub(r"[^а-яa-z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_team_name(value):
    """Normalize team labels coming from Matches, Players and Result workbooks."""
    if pd.isna(value):
        return pd.NA
    text = str(value).strip()
    lower = text.lower().replace("ё", "е")
    if "цска" in lower:
        return CSKA
    if "ростов" in lower:
        return ROSTOV
    return text if text else pd.NA


def normalize_team_column(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "Team" in out.columns:
        out["Team"] = out["Team"].map(normalize_team_name)
    return out


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
    """Attach match columns and preserve Team from Result.xlsx when it exists.

    In the updated Result.xlsx the individual player and goalkeeper sheets already
    contain Team. That value is the source of truth for which club the player
    represented in a specific match. side_number is only a fallback for old files.
    """
    if "ID_match" not in df.columns:
        return normalize_team_column(df)

    out = normalize_team_column(df)
    match_cols = [
        "ID_match", "Date", "Date_text", "Season", "Tournament", "Stage",
        "Home_Team", "Away_Team", "Score", "Winner",
    ]
    match_cols = [c for c in match_cols if c in matches.columns]
    out = out.merge(matches[match_cols], on="ID_match", how="left")

    if "Team" not in out.columns:
        out["Team"] = pd.NA

    if "side_number" in out.columns:
        side = pd.to_numeric(out["side_number"], errors="coerce")
        missing_team = out["Team"].isna() | (out["Team"].astype(str).str.strip() == "")

        home_mask = missing_team & side.eq(1)
        away_mask = missing_team & side.eq(2)

        if "Home_Team" in out.columns:
            out.loc[home_mask, "Team"] = out.loc[home_mask, "Home_Team"].map(normalize_team_name)
        if "Away_Team" in out.columns:
            out.loc[away_mask, "Team"] = out.loc[away_mask, "Away_Team"].map(normalize_team_name)

    out["Team"] = out["Team"].map(normalize_team_name).fillna("Не определено")
    return out


def prepare_individual_stats(df: pd.DataFrame, matches: pd.DataFrame, name_col: str) -> pd.DataFrame:
    """Prepare sheets A/B from Result.xlsx. Prefer explicit Team, fallback to old row order."""
    out = clean_base_table(df)

    if "Team" in out.columns and out["Team"].notna().any():
        return attach_match_context_and_team(out, matches)

    # Backward compatibility for older files without Team in sheet A.
    if name_col in out.columns:
        return add_team_from_blank_separator(out, matches, name_col=name_col)

    return attach_match_context_and_team(out, matches)


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
    text = str(value).strip().replace(",", ".")
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


def normalize_referee_pair(value) -> str:
    if pd.isna(value):
        return "Не указано"
    text = re.sub(r"\s+", " ", str(value)).strip()
    text = re.sub(r"\s*,\s*", ", ", text)
    return text if text else "Не указано"


def arena_city(value) -> str:
    text = str(value).lower()
    if "ростов" in text:
        return "Ростов-на-Дону"
    if "москва" in text:
        return "Москва"
    return "Нейтральная арена"


def short_arena_name(value) -> str:
    text = str(value).strip()
    if "," in text:
        return text.split(",", 1)[1].strip()
    return text


def default_home_arenas(matches: pd.DataFrame) -> list[str]:
    if "Arena" not in matches.columns:
        return []
    source = matches.copy()
    if "Arena_City" not in source.columns:
        source["Arena_City"] = source["Arena"].map(arena_city)
    home = source[source["Arena_City"].isin(["Москва", "Ростов-на-Дону"])]
    if home.empty:
        home = source
    counts = home["Arena"].dropna().astype(str).value_counts()
    return counts.head(4).index.tolist()


def build_referee_rating(matches: pd.DataFrame) -> pd.DataFrame:
    if "Referees" not in matches.columns:
        return pd.DataFrame()

    source = matches.dropna(subset=["Referees"]).copy()
    if source.empty:
        return pd.DataFrame()

    source["referee_score"] = source["Winner"].map({ROSTOV: 1, CSKA: -1}).fillna(0).astype(int)
    rating = (
        source.groupby("Referees", as_index=False)
        .agg(
            Rating=("referee_score", "sum"),
            Matches=("ID_match", "count"),
            Rostov_wins=("Winner", lambda s: int((s == ROSTOV).sum())),
            CSKA_wins=("Winner", lambda s: int((s == CSKA).sum())),
            Draws=("Winner", lambda s: int((s == "Ничья").sum())),
        )
        .sort_values(["Rating", "Rostov_wins", "CSKA_wins", "Matches"], ascending=[False, False, True, False])
        .reset_index(drop=True)
    )
    rating["Rating_per_match"] = (rating["Rating"] / rating["Matches"]).round(2)
    rating["Direction"] = rating["Rating"].map(
        lambda x: "Ростов +" if x > 0 else ("ЦСКА -" if x < 0 else "Баланс")
    )
    return rating


def display_attendance_chart(matches: pd.DataFrame) -> None:
    st.markdown("### Посещаемость по аренам")

    required = {"Arena", "Attendance_Rate"}
    if not required.issubset(matches.columns):
        st.info("В Matches.xlsx нужны колонки Arena и Attendance_Rate для графика посещаемости.")
        return

    source = matches.copy()
    source["Attendance_Rate"] = pd.to_numeric(source["Attendance_Rate"], errors="coerce")
    source = source.dropna(subset=["Arena", "Attendance_Rate"])

    if source.empty:
        st.info("Нет данных посещаемости по выбранным фильтрам.")
        return

    source["Arena_City"] = source.get("Arena_City", source["Arena"].map(arena_city))
    source["Arena_short"] = source.get("Arena_short", source["Arena"].map(short_arena_name))

    arenas = sorted(source["Arena"].dropna().astype(str).unique())
    default_arenas = [arena for arena in default_home_arenas(source) if arena in arenas]
    if not default_arenas:
        default_arenas = arenas[:4]

    selected_arenas = st.multiselect(
        "Арены для гистограммы",
        arenas,
        default=default_arenas,
        key="matches_attendance_arenas",
    )
    if selected_arenas:
        source = source[source["Arena"].isin(selected_arenas)]

    grouped = (
        source.groupby(["Arena", "Arena_short", "Arena_City"], as_index=False)
        .agg(
            avg_attendance=("Attendance_Rate", "mean"),
            matches=("ID_match", "count"),
            avg_capacity=("Capacity", "mean") if "Capacity" in source.columns else ("Attendance_Rate", "size"),
        )
        .sort_values("avg_attendance", ascending=False)
    )

    grouped["avg_attendance"] = grouped["avg_attendance"].round(0).astype(int)
    if "avg_capacity" in grouped.columns:
        grouped["avg_capacity"] = grouped["avg_capacity"].round(0).astype(int)

    colors = grouped["Arena_City"].map(
        {"Ростов-на-Дону": "#FFD200", "Москва": "#D50032", "Нейтральная арена": "#777777"}
    ).fillna("#777777")

    fig = go.Figure(
        data=[
            go.Bar(
                x=grouped["Arena_short"],
                y=grouped["avg_attendance"],
                marker_color=colors,
                text=grouped["avg_attendance"].map(lambda x: f"{x:,}".replace(",", " ")),
                textposition="outside",
                hovertemplate=(
                    "<b>%{customdata[0]}</b><br>"
                    "Средняя посещаемость: %{y:.0f}<br>"
                    "Матчей: %{customdata[1]}<br>"
                    "Средняя вместимость: %{customdata[2]:.0f}<extra></extra>"
                ),
                customdata=grouped[["Arena", "matches", "avg_capacity"]],
            )
        ]
    )
    fig.update_layout(
        title="Средняя посещаемость Attendance_Rate по 4 основным аренам",
        xaxis_title="Арена",
        yaxis_title="Средняя посещаемость, чел.",
        bargap=0.28,
        height=460,
        margin=dict(t=70, b=110),
    )
    fig.update_xaxes(tickangle=-18)
    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(
        grouped.rename(
            columns={
                "Arena": "Арена",
                "Arena_City": "Город",
                "avg_attendance": "Средняя посещаемость",
                "matches": "Матчей",
                "avg_capacity": "Средняя вместимость",
            }
        )[["Арена", "Город", "Средняя посещаемость", "Матчей", "Средняя вместимость"]],
        use_container_width=True,
        hide_index=True,
    )


def display_referee_rating(matches: pd.DataFrame) -> None:
    st.markdown("### Рейтинг пар судей")
    st.caption("Формула рейтинга: победа Ростова = +1, победа ЦСКА = -1, ничья = 0. Сортировка: от наиболее ростовского значения к наиболее цэсковскому.")

    rating = build_referee_rating(matches)
    if rating.empty:
        st.info("Нет данных по парам судей в выбранной выборке.")
        return

    min_matches = st.slider(
        "Минимум матчей для рейтинга",
        min_value=1,
        max_value=max(1, int(rating["Matches"].max())),
        value=1,
        key="matches_referee_min_matches",
    )
    rating = rating[rating["Matches"] >= min_matches].copy()
    if rating.empty:
        st.info("После фильтра по количеству матчей пар судей не осталось.")
        return

    colors = rating["Rating"].map(lambda x: "#FFD200" if x > 0 else ("#D50032" if x < 0 else "#777777"))
    chart = rating.iloc[::-1].copy()
    chart_colors = colors.iloc[::-1]

    fig = go.Figure(
        data=[
            go.Bar(
                x=chart["Rating"],
                y=chart["Referees"],
                orientation="h",
                marker_color=chart_colors,
                text=chart["Rating"],
                textposition="outside",
                hovertemplate=(
                    "<b>%{y}</b><br>"
                    "Рейтинг: %{x}<br>"
                    "Матчей: %{customdata[0]}<br>"
                    "Победы Ростова: %{customdata[1]}<br>"
                    "Победы ЦСКА: %{customdata[2]}<extra></extra>"
                ),
                customdata=chart[["Matches", "Rostov_wins", "CSKA_wins"]],
            )
        ]
    )
    fig.add_vline(x=0, line_width=1, line_dash="dash", line_color="#AAAAAA")
    fig.update_layout(
        title="Пары судей: от удобных для Ростова к удобным для ЦСКА",
        xaxis_title="Рейтинг: Ростов +1 / ЦСКА -1",
        yaxis_title="Пара судей",
        height=max(420, 52 * len(chart)),
        margin=dict(l=260, r=40, t=70, b=60),
    )
    st.plotly_chart(fig, use_container_width=True)

    visible = rating.rename(
        columns={
            "Referees": "Пара судей",
            "Rating": "Рейтинг",
            "Rating_per_match": "Рейтинг за матч",
            "Matches": "Матчей",
            "Rostov_wins": "Побед Ростова",
            "CSKA_wins": "Побед ЦСКА",
            "Draws": "Ничьих",
            "Direction": "Направление",
        }
    )
    st.dataframe(
        visible[["Пара судей", "Рейтинг", "Рейтинг за матч", "Матчей", "Побед Ростова", "Побед ЦСКА", "Ничьих", "Направление"]],
        use_container_width=True,
        hide_index=True,
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

    st.markdown("### Фильтры вкладки")
    match_view = matches.copy()

    f1, f2, f3 = st.columns(3)
    with f1:
        arena_options = sorted(match_view["Arena"].dropna().astype(str).unique()) if "Arena" in match_view.columns else []
        selected_arenas = st.multiselect(
            "Арены",
            arena_options,
            default=arena_options,
            key="matches_tab_arenas",
        )
        if selected_arenas:
            match_view = match_view[match_view["Arena"].astype(str).isin(selected_arenas)]

    with f2:
        referee_options = sorted(match_view["Referees"].dropna().astype(str).unique()) if "Referees" in match_view.columns else []
        selected_referees = st.multiselect(
            "Пары судей",
            referee_options,
            default=referee_options,
            key="matches_tab_referees",
        )
        if selected_referees:
            match_view = match_view[match_view["Referees"].astype(str).isin(selected_referees)]

    with f3:
        winner_options = ["Все"] + [x for x in [ROSTOV, CSKA, "Ничья"] if x in set(match_view["Winner"].dropna().astype(str))]
        selected_winner = st.selectbox("Победитель", winner_options, key="matches_tab_winner")
        if selected_winner != "Все":
            match_view = match_view[match_view["Winner"].astype(str) == selected_winner]

    if "Date" in match_view.columns and match_view["Date"].notna().any():
        min_date = match_view["Date"].min().date()
        max_date = match_view["Date"].max().date()
        date_range = st.date_input(
            "Период матчей",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            key="matches_tab_date_range",
        )
        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_date, end_date = date_range
            match_view = match_view[
                (match_view["Date"].dt.date >= start_date) &
                (match_view["Date"].dt.date <= end_date)
            ]

    if match_view.empty:
        st.info("После фильтров вкладки матчей не осталось.")
        return

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Матчей", len(match_view))
    c2.metric("Побед ЦСКА", int((match_view["Winner"] == CSKA).sum()))
    c3.metric("Побед Ростова", int((match_view["Winner"] == ROSTOV).sum()))
    c4.metric("Голов всего", int(match_view["Number_of_Goals"].fillna(0).sum()))
    avg_attendance = "-"
    if "Attendance_Rate" in match_view.columns and match_view["Attendance_Rate"].notna().any():
        avg_attendance = int(round(match_view["Attendance_Rate"].dropna().mean(), 0))
    c5.metric("Средняя посещаемость", avg_attendance)

    st.dataframe(
        match_view[[
            "ID_match", "Date_text", "Season", "Tournament", "Stage", "Home_Team",
            "Away_Team", "First_Half_Score", "Score", "Number_of_Goals", "Result",
            "Referees", "Attendance_Rate", "Capacity", "Arena"
        ]],
        use_container_width=True,
        hide_index=True,
    )

    display_attendance_chart(match_view)
    display_referee_rating(match_view)

    st.markdown("### Детализация выбранного матча")
    match_options = match_view.sort_values("Date", ascending=False).copy()
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

    if "Referees" in selected_match.index or "Attendance_Rate" in selected_match.index:
        d1, d2, d3 = st.columns(3)
        d1.metric("Пара судей", str(selected_match.get("Referees", "-")))
        d2.metric("Посещаемость", int(selected_match["Attendance_Rate"]) if pd.notna(selected_match.get("Attendance_Rate")) else "-")
        d3.metric("Арена", str(selected_match.get("Arena_short", selected_match.get("Arena", "-"))))

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


def is_slash_stat_column(series: pd.Series) -> bool:
    sample = series.dropna().astype(str).str.strip().head(200)
    if sample.empty:
        return False
    pattern = r"^\d+(?:[\.,]\d+)?\s*/\s*\d+(?:[\.,]\d+)?$"
    return bool(sample.str.match(pattern).any())


def format_stat_number(value) -> str:
    if pd.isna(value):
        return ""
    value = float(value)
    if value.is_integer():
        return str(int(value))
    return str(round(value, 1))


def format_slash_sum(made: float, attempts: float) -> str:
    if not made and not attempts:
        return ""
    return f"{format_stat_number(made)}/{format_stat_number(attempts)}"


def build_player_summary_table(player_stats: pd.DataFrame) -> pd.DataFrame:
    """Aggregate field-player statistics by player and team.

    Columns with values like 3/5 are aggregated as sum(left)/sum(right),
    for example 3/5 + 2/4 = 5/9. Goalkeepers are excluded by Position.
    """
    if player_stats.empty or "Players" not in player_stats.columns:
        return pd.DataFrame()

    df = player_stats.copy()
    df = df[df["Players"].notna()].copy()

    if "Position" in df.columns:
        df = df[~df["Position"].astype(str).str.lower().str.contains("вратар", na=False)].copy()

    if df.empty:
        return pd.DataFrame()

    if "Team" not in df.columns:
        df["Team"] = "Не определено"
    df["Team"] = df["Team"].fillna("Не определено")

    excluded_cols = {
        "ID_match", "Date", "Date_text", "Season", "Tournament", "Stage", "Home_Team", "Away_Team",
        "Score", "Winner", "Players", "Team", "Position", "Citizenship", "Height", "Age", "player_key",
        "Player_number/Jersey_number", "No", "No_num", "Percentage", "Playing_time",
        "goals_made", "goals_attempts", "goals_pct_calc",
    }

    slash_cols = [
        col for col in df.columns
        if col not in excluded_cols and is_slash_stat_column(df[col])
    ]

    numeric_cols = []
    for col in df.select_dtypes(include="number").columns:
        if col in excluded_cols:
            continue
        if col.endswith("_made") or col.endswith("_attempts") or col.endswith("_pct_calc"):
            continue
        numeric_cols.append(col)

    rows: list[dict[str, object]] = []
    group_cols = ["Players", "Team"]

    for (player, team), group in df.groupby(group_cols, dropna=False, sort=False):
        row: dict[str, object] = {
            "Игрок": player,
            "Команда": team if pd.notna(team) else "Не определено",
            "Матчи": int(group["ID_match"].nunique()) if "ID_match" in group.columns else int(len(group)),
        }

        if "Position" in group.columns:
            position_values = group["Position"].dropna().astype(str)
            row["Амплуа"] = position_values.iloc[0] if not position_values.empty else ""

        goals_made_for_sort = 0.0
        goals_attempts_for_sort = 0.0

        for col in slash_cols:
            pairs = group[col].map(split_goal_attempt)
            made = sum(x[0] for x in pairs if x[0] is not None)
            attempts = sum(x[1] for x in pairs if x[1] is not None)
            row[col] = format_slash_sum(made, attempts)

            if col == "Goals":
                goals_made_for_sort = made
                goals_attempts_for_sort = attempts
                row["% голов"] = round(made / attempts * 100, 1) if attempts else ""

        for col in numeric_cols:
            values = pd.to_numeric(group[col], errors="coerce")
            total = values.sum(min_count=1)
            row[col] = format_stat_number(total) if pd.notna(total) else ""

        row["__goals_made"] = goals_made_for_sort
        row["__goals_attempts"] = goals_attempts_for_sort
        rows.append(row)

    summary = pd.DataFrame(rows)
    if summary.empty:
        return summary

    if "Амплуа" in summary.columns:
        base_cols = ["Игрок", "Команда", "Амплуа", "Матчи"]
    else:
        base_cols = ["Игрок", "Команда", "Матчи"]

    stat_cols = [
        c for c in summary.columns
        if c not in base_cols and not c.startswith("__")
    ]
    ordered_cols = base_cols + stat_cols + ["__goals_made", "__goals_attempts"]
    summary = summary[ordered_cols]
    summary = summary.sort_values(["__goals_made", "__goals_attempts", "Матчи"], ascending=[False, False, False])
    return summary.reset_index(drop=True)


def display_players_summary(player_stats: pd.DataFrame) -> None:
    st.markdown("### Сводная статистика полевых игроков")
    st.caption(
        "Вратари исключены. Показатели формата `голы/попытки` суммируются отдельно до и после слэша: "
        "например, `3/5 + 2/4 = 5/9`."
    )

    summary = build_player_summary_table(player_stats)
    if summary.empty:
        st.info("Нет данных по полевым игрокам для выбранной выборки матчей.")
        return

    c1, c2 = st.columns([2, 1])
    with c1:
        player_options = ["Все"] + sorted(summary["Игрок"].dropna().astype(str).unique().tolist())
        selected_player = st.selectbox("Игрок", player_options, key="players_summary_player")
    with c2:
        team_options = ["Все"] + sorted(summary["Команда"].dropna().astype(str).unique().tolist())
        selected_team = st.selectbox("Команда в статистике", team_options, key="players_summary_team")

    view = summary.copy()
    if selected_player != "Все":
        view = view[view["Игрок"].astype(str) == selected_player]
    if selected_team != "Все":
        view = view[view["Команда"].astype(str) == selected_team]

    m1, m2, m3 = st.columns(3)
    m1.metric("Игроков в таблице", view["Игрок"].nunique())
    m2.metric("Командо-записей", len(view))
    if "__goals_made" in view.columns:
        m3.metric("Голов", int(view["__goals_made"].fillna(0).sum()))

    display_cols = [c for c in view.columns if not c.startswith("__")]
    st.dataframe(view[display_cols], use_container_width=True, hide_index=True)


def first_non_empty_text(series: pd.Series) -> str:
    values = series.dropna().astype(str).str.strip()
    values = values[values != ""]
    return values.iloc[0] if not values.empty else ""


def build_goalkeeper_summary_table(goalkeeper_stats: pd.DataFrame) -> pd.DataFrame:
    """Aggregate goalkeeper statistics by goalkeeper and team.

    Columns with values like 9/27 are aggregated as sum(left)/sum(right),
    for example 9/27 + 4/10 = 13/37. Field players are not included because
    the source table is the goalkeeper statistics sheet.
    """
    if goalkeeper_stats.empty or "goalkeeper" not in goalkeeper_stats.columns:
        return pd.DataFrame()

    df = goalkeeper_stats.copy()
    df = df[df["goalkeeper"].notna()].copy()
    df = df[~df["goalkeeper"].astype(str).str.strip().str.lower().isin(["итого", "итого:", "total"])]

    if df.empty:
        return pd.DataFrame()

    if "Team" not in df.columns:
        df["Team"] = "Не определено"
    df["Team"] = df["Team"].fillna("Не определено")

    excluded_cols = {
        "ID_match", "Date", "Date_text", "Season", "Tournament", "Stage", "Home_Team", "Away_Team",
        "Score", "Winner", "goalkeeper", "Team", "player_key", "No_num", "no",
    }
    pct_cols = {c for c in df.columns if str(c).endswith("_pct") or str(c).endswith("_save_pct") or c == "save_pct"}

    slash_cols = [
        col for col in df.columns
        if col not in excluded_cols and col not in pct_cols and is_slash_stat_column(df[col])
    ]

    preferred_slash_order = [
        "total", "six_m", "wing", "close_range", "long_range", "seven_m", "counterattack", "one_on_one", "quick_start"
    ]
    slash_cols = [c for c in preferred_slash_order if c in slash_cols] + [c for c in slash_cols if c not in preferred_slash_order]

    pct_by_slash = {
        "total": "save_pct",
        "six_m": "six_m_save_pct",
        "wing": "wing_save_pct",
        "close_range": "close_range_save_pct",
        "long_range": "long_range_save_pct",
        "seven_m": "seven_m_save_pct",
    }

    rows: list[dict[str, object]] = []
    group_cols = ["goalkeeper", "Team"]

    for (goalkeeper, team), group in df.groupby(group_cols, dropna=False, sort=False):
        row: dict[str, object] = {
            "Вратарь": goalkeeper,
            "Команда": team if pd.notna(team) else "Не определено",
            "Матчи": int(group["ID_match"].nunique()) if "ID_match" in group.columns else int(len(group)),
        }

        if "no" in group.columns:
            row["№"] = first_non_empty_text(group["no"])

        total_saves_for_sort = 0.0
        total_shots_for_sort = 0.0

        for col in slash_cols:
            pairs = group[col].map(split_goal_attempt)
            made = sum(x[0] for x in pairs if x[0] is not None)
            attempts = sum(x[1] for x in pairs if x[1] is not None)
            row[col] = format_slash_sum(made, attempts)

            pct_col = pct_by_slash.get(col)
            if pct_col:
                row[pct_col] = round(made / attempts * 100, 1) if attempts else ""

            if col == "total":
                total_saves_for_sort = made
                total_shots_for_sort = attempts

        row["__saves"] = total_saves_for_sort
        row["__shots"] = total_shots_for_sort
        rows.append(row)

    summary = pd.DataFrame(rows)
    if summary.empty:
        return summary

    base_cols = ["Вратарь", "Команда"]
    if "№" in summary.columns:
        base_cols.append("№")
    base_cols.append("Матчи")

    ordered_stat_cols: list[str] = []
    for col in slash_cols:
        if col in summary.columns:
            ordered_stat_cols.append(col)
        pct_col = pct_by_slash.get(col)
        if pct_col in summary.columns:
            ordered_stat_cols.append(pct_col)

    other_cols = [c for c in summary.columns if c not in base_cols + ordered_stat_cols and not c.startswith("__")]
    summary = summary[base_cols + ordered_stat_cols + other_cols + ["__saves", "__shots"]]
    summary = summary.sort_values(["__saves", "__shots", "Матчи"], ascending=[False, False, False])
    return summary.reset_index(drop=True)


def display_goalkeepers_summary(goalkeeper_stats: pd.DataFrame) -> None:
    st.markdown("### Сводная статистика вратарей")
    st.caption(
        "Полевые игроки исключены. Показатели формата `сейвы/броски` суммируются отдельно до и после слэша: "
        "например, `9/27 + 4/10 = 13/37`. Проценты пересчитываются от итоговой дроби."
    )

    summary = build_goalkeeper_summary_table(goalkeeper_stats)
    if summary.empty:
        st.info("Нет данных по вратарям для выбранной выборки матчей.")
        return

    c1, c2 = st.columns([2, 1])
    with c1:
        goalkeeper_options = ["Все"] + sorted(summary["Вратарь"].dropna().astype(str).unique().tolist())
        selected_goalkeeper = st.selectbox("Вратарь", goalkeeper_options, key="goalkeepers_summary_player")
    with c2:
        team_options = ["Все"] + sorted(summary["Команда"].dropna().astype(str).unique().tolist())
        selected_team = st.selectbox("Команда в статистике", team_options, key="goalkeepers_summary_team")

    view = summary.copy()
    if selected_goalkeeper != "Все":
        view = view[view["Вратарь"].astype(str) == selected_goalkeeper]
    if selected_team != "Все":
        view = view[view["Команда"].astype(str) == selected_team]

    m1, m2, m3 = st.columns(3)
    m1.metric("Вратарей в таблице", view["Вратарь"].nunique())
    m2.metric("Командо-записей", len(view))
    if "__saves" in view.columns and "__shots" in view.columns:
        saves = int(view["__saves"].fillna(0).sum())
        shots = int(view["__shots"].fillna(0).sum())
        m3.metric("Сейвы/броски", f"{saves}/{shots}" if shots else "-")

    display_cols = [c for c in view.columns if not c.startswith("__")]
    st.dataframe(view[display_cols], use_container_width=True, hide_index=True)


def display_players_tab(players: pd.DataFrame, officials: pd.DataFrame, player_stats: pd.DataFrame, goalkeeper_stats: pd.DataFrame) -> None:
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

    display_players_summary(player_stats)
    display_goalkeepers_summary(goalkeeper_stats)

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



def player_stats_chart_source(filtered: pd.DataFrame) -> pd.DataFrame:
    """Return field players only for charts on the player statistics tab."""
    df = filtered.copy()
    if "Position" in df.columns:
        df = df[~df["Position"].astype(str).str.lower().str.contains("вратар", na=False)].copy()
    return df


def numeric_stat(df: pd.DataFrame, col: str) -> pd.Series:
    """Safe numeric series for optional stat columns."""
    if col not in df.columns:
        return pd.Series(0, index=df.index, dtype="float64")
    return pd.to_numeric(df[col], errors="coerce").fillna(0)


def build_goals_assists_top(filtered: pd.DataFrame) -> pd.DataFrame:
    """Top field players by goals + assists for the selected filters."""
    df = player_stats_chart_source(filtered)
    if df.empty or "Players" not in df.columns:
        return pd.DataFrame()

    df = df.copy()
    if "Team" not in df.columns:
        df["Team"] = "Не определено"
    if "goals_made" not in df.columns:
        if "Goals" in df.columns:
            df["goals_made"] = df["Goals"].map(lambda x: split_goal_attempt(x)[0] or 0)
        else:
            df["goals_made"] = 0
    df["Assist"] = numeric_stat(df, "Assist")

    grouped = (
        df.groupby(["Players", "Team"], dropna=False, as_index=False)
        .agg(
            Голы=("goals_made", "sum"),
            Пасы=("Assist", "sum"),
            Матчи=("ID_match", "nunique") if "ID_match" in df.columns else ("Players", "size"),
        )
    )
    grouped["Гол+пас"] = grouped["Голы"].fillna(0) + grouped["Пасы"].fillna(0)
    grouped["Игрок"] = grouped["Players"].astype(str) + " · " + grouped["Team"].astype(str)
    return grouped.sort_values(["Гол+пас", "Голы", "Пасы", "Матчи"], ascending=False).head(15)


def build_defenders_top(filtered: pd.DataFrame) -> pd.DataFrame:
    """Top field players by blocks + interceptions for the selected filters."""
    df = player_stats_chart_source(filtered)
    if df.empty or "Players" not in df.columns:
        return pd.DataFrame()

    df = df.copy()
    if "Team" not in df.columns:
        df["Team"] = "Не определено"

    block_col = "Close-range_shot_(<9m).1"
    if block_col not in df.columns:
        block_candidates = [c for c in df.columns if str(c).endswith(".1") and "Close-range" in str(c)]
        block_col = block_candidates[0] if block_candidates else ""

    df["__blocks"] = numeric_stat(df, block_col) if block_col else 0
    df["__interceptions"] = numeric_stat(df, "Steal/Interception")

    grouped = (
        df.groupby(["Players", "Team"], dropna=False, as_index=False)
        .agg(
            Блоки=("__blocks", "sum"),
            Перехваты=("__interceptions", "sum"),
            Матчи=("ID_match", "nunique") if "ID_match" in df.columns else ("Players", "size"),
        )
    )
    grouped["Блок+перехват"] = grouped["Блоки"].fillna(0) + grouped["Перехваты"].fillna(0)
    grouped["Игрок"] = grouped["Players"].astype(str) + " · " + grouped["Team"].astype(str)
    return grouped.sort_values(["Блок+перехват", "Блоки", "Перехваты", "Матчи"], ascending=False).head(10)


def build_suspensions_top(filtered: pd.DataFrame) -> pd.DataFrame:
    """Top field players by average two-minute suspension time per match."""
    df = player_stats_chart_source(filtered)
    if df.empty or "Players" not in df.columns or "2-minutes" not in df.columns:
        return pd.DataFrame()

    df = df.copy()
    if "Team" not in df.columns:
        df["Team"] = "Не определено"

    df["__suspensions"] = numeric_stat(df, "2-minutes")
    df = df[df["__suspensions"] > 0].copy()
    if df.empty:
        return pd.DataFrame()

    grouped = (
        df.groupby(["Players", "Team"], dropna=False, as_index=False)
        .agg(
            Удаления=("__suspensions", "sum"),
            Матчи=("ID_match", "nunique") if "ID_match" in df.columns else ("Players", "size"),
        )
    )
    grouped["Минуты удалений"] = grouped["Удаления"].fillna(0) * 2
    grouped["Среднее время удалений за матч"] = grouped.apply(
        lambda r: round(r["Минуты удалений"] / r["Матчи"], 2) if r["Матчи"] else 0,
        axis=1,
    )
    grouped["Игрок"] = grouped["Players"].astype(str) + " · " + grouped["Team"].astype(str)
    return grouped.sort_values(
        ["Среднее время удалений за матч", "Минуты удалений", "Удаления", "Матчи"],
        ascending=False,
    ).head(10)


def team_palette_for_chart(team_value: object) -> tuple[str, str, str]:
    """Return first-stack color, second-stack color and outline color by team."""
    team = str(team_value).lower()

    if "ростов" in team or "rostov" in team:
        # Ростов: желтый + черный. Черный сегмент получает желтую обводку,
        # чтобы он был читаемым на темной теме Streamlit.
        return "#FFD200", "#111111", "#FFD200"

    if "цска" in team or "cska" in team:
        # ЦСКА: красный + синий.
        return "#D50032", "#0033A0", "#FFFFFF"

    return "#8A8A8A", "#4A4A4A", "#FFFFFF"


def metric_colors_by_team(data: pd.DataFrame, metric_index: int) -> tuple[list[str], list[str]]:
    """Build per-player bar colors. First metric uses team primary, second uses team secondary."""
    fill_colors: list[str] = []
    line_colors: list[str] = []

    team_values = data["Team"] if "Team" in data.columns else pd.Series("Не определено", index=data.index)
    for team in team_values:
        primary, secondary, outline = team_palette_for_chart(team)
        fill_colors.append(primary if metric_index == 0 else secondary)
        line_colors.append(outline)

    return fill_colors, line_colors


def render_stacked_player_chart(data: pd.DataFrame, x_col: str, y_cols: list[str], title: str, y_title: str) -> None:
    if data.empty:
        st.info("Нет данных для построения графика по выбранным фильтрам.")
        return

    fig = go.Figure()

    for metric_index, col in enumerate(y_cols):
        fill_colors, line_colors = metric_colors_by_team(data, metric_index)
        fig.add_trace(
            go.Bar(
                x=data[x_col],
                y=data[col],
                name=col,
                marker=dict(
                    color=fill_colors,
                    line=dict(color=line_colors, width=1.2),
                ),
                customdata=data[["Team"]] if "Team" in data.columns else None,
                hovertemplate=(
                    "%{x}<br>"
                    + ("Команда: %{customdata[0]}<br>" if "Team" in data.columns else "")
                    + f"{col}: %{{y}}<extra></extra>"
                ),
            )
        )

    fig.update_layout(
        title=title,
        barmode="stack",
        xaxis_title="Игрок · команда",
        yaxis_title=y_title,
        legend_title="Показатель",
        xaxis_tickangle=-35,
        margin=dict(l=10, r=10, t=70, b=135),
        annotations=[
            dict(
                text="Ростов: желтый/черный · ЦСКА: красный/синий",
                xref="paper",
                yref="paper",
                x=0,
                y=1.12,
                showarrow=False,
                align="left",
                font=dict(size=12),
            )
        ],
    )
    st.plotly_chart(fig, use_container_width=True)


def display_player_stats_tab(player_stats: pd.DataFrame) -> None:
    st.subheader("Статистика игроков")
    filtered = stats_filters(player_stats, "players")
    chart_source = player_stats_chart_source(filtered)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Записей", len(filtered))
    c2.metric("Полевых игроков", chart_source["Players"].nunique() if "Players" in chart_source.columns else 0)
    c3.metric("Голы", int(chart_source["goals_made"].fillna(0).sum()) if "goals_made" in chart_source.columns else 0)
    c4.metric("Пасы", int(numeric_stat(chart_source, "Assist").sum()) if not chart_source.empty else 0)

    st.markdown("### Топ-15 по гол+пас")
    attack_top = build_goals_assists_top(filtered)
    render_stacked_player_chart(
        attack_top,
        x_col="Игрок",
        y_cols=["Голы", "Пасы"],
        title="Топ-15 полевых игроков по сумме гол+пас",
        y_title="Голы + пасы",
    )
    if not attack_top.empty:
        st.dataframe(
            attack_top[["Players", "Team", "Матчи", "Голы", "Пасы", "Гол+пас"]]
            .rename(columns={"Players": "Игрок", "Team": "Команда"}),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("### Топ-10 защитников по блок+перехваты")
    defense_top = build_defenders_top(filtered)
    render_stacked_player_chart(
        defense_top,
        x_col="Игрок",
        y_cols=["Блоки", "Перехваты"],
        title="Топ-10 полевых игроков по сумме блок+перехват",
        y_title="Блоки + перехваты",
    )
    if not defense_top.empty:
        st.dataframe(
            defense_top[["Players", "Team", "Матчи", "Блоки", "Перехваты", "Блок+перехват"]]
            .rename(columns={"Players": "Игрок", "Team": "Команда"}),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("### Топ-10 по среднему времени удалений")
    suspensions_top = build_suspensions_top(filtered)
    if suspensions_top.empty:
        st.info("Нет данных по удалениям для выбранных фильтров.")
    else:
        render_stacked_player_chart(
            suspensions_top,
            x_col="Игрок",
            y_cols=["Среднее время удалений за матч"],
            title="Топ-10 полевых игроков по среднему времени удалений за матч",
            y_title="Минуты удалений за матч",
        )
        st.dataframe(
            suspensions_top[["Players", "Team", "Матчи", "Удаления", "Минуты удалений", "Среднее время удалений за матч"]]
            .rename(columns={"Players": "Игрок", "Team": "Команда"}),
            use_container_width=True,
            hide_index=True,
        )

    visible_cols = [
        "ID_match", "Date_text", "Season", "Tournament", "Home_Team", "Away_Team", "Score",
        "Team", "Player_number/Jersey_number", "Players", "Position", "Goals", "Assist",
        "Steal/Interception", "Close-range_shot_(<9m).1", "Turnover", "Total_turnovers",
        "2-minutes", "Playing_time"
    ]
    visible_cols = [c for c in visible_cols if c in filtered.columns]
    with st.expander("Индивидуальные строки статистики игроков"):
        st.dataframe(filtered[visible_cols], use_container_width=True, hide_index=True)


def build_goalkeeper_save_rating(goalkeeper_stats: pd.DataFrame) -> pd.DataFrame:
    """Aggregate individual goalkeeper rows as saves/shots and recalculate save %."""
    if goalkeeper_stats.empty or "goalkeeper" not in goalkeeper_stats.columns or "total" not in goalkeeper_stats.columns:
        return pd.DataFrame()

    df = goalkeeper_stats.copy()
    df = df[df["goalkeeper"].notna()].copy()
    df = df[~df["goalkeeper"].astype(str).str.strip().str.lower().isin(["итого", "итого:", "total"])]
    if df.empty:
        return pd.DataFrame()

    if "Team" not in df.columns:
        df["Team"] = "Не определено"
    df["Team"] = df["Team"].fillna("Не определено")

    pairs = df["total"].map(split_goal_attempt)
    df["__saves"] = pairs.map(lambda x: x[0] if x[0] is not None else 0)
    df["__shots"] = pairs.map(lambda x: x[1] if x[1] is not None else 0)

    grouped = (
        df.groupby(["goalkeeper", "Team"], dropna=False, as_index=False)
        .agg(
            Матчи=("ID_match", "nunique") if "ID_match" in df.columns else ("goalkeeper", "size"),
            Отбитые=("__saves", "sum"),
            Броски=("__shots", "sum"),
        )
    )
    grouped["Сейвы/броски"] = grouped.apply(
        lambda r: f"{format_stat_number(r['Отбитые'])}/{format_stat_number(r['Броски'])}" if r["Броски"] else "",
        axis=1,
    )
    grouped["% сейвов"] = grouped.apply(
        lambda r: round(r["Отбитые"] / r["Броски"] * 100, 2) if r["Броски"] else None,
        axis=1,
    )
    grouped = grouped.rename(columns={"goalkeeper": "Вратарь", "Team": "Команда"})
    return grouped.sort_values(["% сейвов", "Броски", "Матчи"], ascending=[False, False, False]).reset_index(drop=True)


def display_goalkeeper_stats_tab(goalkeeper_stats: pd.DataFrame, aggregate_goalkeepers: pd.DataFrame) -> None:
    st.subheader("Статистика вратарей")
    filtered = stats_filters(goalkeeper_stats, "gk")

    rating_all = build_goalkeeper_save_rating(filtered)
    if rating_all.empty:
        st.info("Нет данных по индивидуальной статистике вратарей для выбранных фильтров.")
        return

    max_matches = int(rating_all["Матчи"].max()) if rating_all["Матчи"].notna().any() else 1
    min_matches = st.slider(
        "Минимальное количество матчей",
        min_value=1,
        max_value=max(1, max_matches),
        value=1,
        step=1,
    )
    rating = rating_all[rating_all["Матчи"] >= min_matches].copy()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Записей", len(filtered))
    c2.metric("Вратарей", rating["Вратарь"].nunique())
    c3.metric("Отбитые/броски", f"{int(rating['Отбитые'].sum())}/{int(rating['Броски'].sum())}" if not rating.empty else "-")
    total_shots = rating["Броски"].sum() if not rating.empty else 0
    total_saves = rating["Отбитые"].sum() if not rating.empty else 0
    c4.metric("Общий % сейвов", f"{round(total_saves / total_shots * 100, 2)}%" if total_shots else "-")

    if rating.empty:
        st.info("После фильтра по минимальному количеству матчей данных не осталось.")
    else:
        chart_data = rating.head(15)
        fig = px.bar(
            chart_data,
            x="Вратарь",
            y="% сейвов",
            color="Команда",
            text="% сейвов",
            title="Рейтинг вратарей по проценту отбитых бросков",
            labels={"% сейвов": "% сейвов", "Вратарь": "Вратарь"},
            color_discrete_map={CSKA: "#D50032", ROSTOV: "#FFD200", "Не определено": "#777777"},
        )
        fig.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
        fig.update_layout(xaxis_tickangle=-35, yaxis_ticksuffix="%", margin=dict(l=10, r=10, t=70, b=120))
        st.plotly_chart(fig, use_container_width=True)

        table_cols = ["Вратарь", "Команда", "Матчи", "Отбитые", "Броски", "Сейвы/броски", "% сейвов"]
        st.dataframe(rating[table_cols], use_container_width=True, hide_index=True)

    visible_cols = [
        "ID_match", "Date_text", "Season", "Tournament", "Home_Team", "Away_Team", "Score",
        "Team", "no", "goalkeeper", "total", "save_pct", "six_m", "six_m_save_pct",
        "wing", "wing_save_pct", "close_range", "close_range_save_pct", "long_range", "long_range_save_pct",
        "seven_m", "seven_m_save_pct"
    ]
    visible_cols = [c for c in visible_cols if c in filtered.columns]
    with st.expander("Индивидуальные строки статистики вратарей"):
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
        display_players_tab(tables["players"], tables["officials"], player_stats, goalkeeper_stats)
    with tab_player_stats:
        display_player_stats_tab(player_stats)
    with tab_gk_stats:
        display_goalkeeper_stats_tab(goalkeeper_stats, aggregate_goalkeepers)
    with tab_sql:
        display_duckdb_tab(tables)


if __name__ == "__main__":
    main()
