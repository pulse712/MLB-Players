"""
MLB Series Dog Scatter Analysis
================================
Interactive scatter plot: Dog Win% vs Fav Win%
Color: Blue = WIN, Orange = LOSS
Plots today's qualifying matchups as star markers.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
import os
from datetime import date, timedelta
import requests

st.set_page_config(
    page_title="Scatter Analysis — MLB Betting",
    page_icon="📊",
    layout="wide"
)

MARTINGALE_FILE = 'Martingale_Series_Analysis.xlsx'

# ── LOAD HISTORICAL DATA ──────────────────────────────────────────
@st.cache_data(ttl=86400, show_spinner="Loading historical series data...")
def load_series_data(file_bytes):
    import io as _io
    df = pd.read_excel(_io.BytesIO(file_bytes), sheet_name='Raw Series Data', skiprows=1)
    df.columns = ['Year','Date','Road_Team','Home_Team','Dog_Location',
                  'Dog_WinPct','Fav_WinPct','Avg_Line','Results','Won_1_Game']
    df = df.dropna(subset=['Dog_WinPct','Fav_WinPct'])
    df['Dog_WinPct'] = pd.to_numeric(df['Dog_WinPct'], errors='coerce')
    df['Fav_WinPct'] = pd.to_numeric(df['Fav_WinPct'], errors='coerce')
    df['Avg_Line']   = pd.to_numeric(df['Avg_Line'],   errors='coerce')
    df['Won_1_Game'] = df['Won_1_Game'].astype(str).str.strip().str.upper()
    df['Year']       = pd.to_numeric(df['Year'], errors='coerce').astype('Int64')
    df['Date']       = pd.to_datetime(df['Date'], errors='coerce')
    df = df.dropna(subset=['Dog_WinPct','Fav_WinPct','Won_1_Game'])
    return df


# ── FETCH TODAY'S GAMES WIN% FROM MLB API ─────────────────────────
@st.cache_data(ttl=1800, show_spinner="Fetching today's standings...")
def get_team_winpcts(report_date_str):
    """Fetch current season win% for all teams from MLB Stats API."""
    report_date = date.fromisoformat(report_date_str)
    season = report_date.year
    url = f'https://statsapi.mlb.com/api/v1/standings?leagueId=103,104&season={season}&standingsTypes=regularSeason'
    try:
        r = requests.get(url, timeout=10)
        data = r.json()
    except Exception:
        return {}

    winpcts = {}
    for record in data.get('records', []):
        for tr in record.get('teamRecords', []):
            name = tr['team']['name']
            wins = tr.get('wins', 0)
            losses = tr.get('losses', 0)
            total = wins + losses
            winpcts[name] = round(wins / total, 4) if total > 0 else 0.500
    return winpcts


@st.cache_data(ttl=1800, show_spinner="Fetching today's schedule...")
def get_todays_games(report_date_str):
    """Fetch today's schedule from MLB Stats API."""
    url = f'https://statsapi.mlb.com/api/v1/schedule?sportId=1&date={report_date_str}'
    try:
        r = requests.get(url, timeout=10)
        data = r.json()
    except Exception:
        return []

    games = []
    if not data.get('dates'):
        return []
    for g in data['dates'][0]['games']:
        away = g['teams']['away']['team']['name']
        home = g['teams']['home']['team']['name']
        games.append({'away': away, 'home': home})
    return games


def determine_dog_fav(game, winpcts):
    """
    Given a game dict with 'away'/'home' keys and a winpcts lookup,
    determine which team is the dog and which is the fav by win%.
    Returns dict with dog/fav info, or None if can't determine.
    """
    away, home = game['away'], game['home']
    away_wp = winpcts.get(away)
    home_wp = winpcts.get(home)
    if away_wp is None or home_wp is None:
        return None
    if away_wp < home_wp:
        dog, fav = away, home
        dog_wp, fav_wp = away_wp, home_wp
        dog_loc = 'ROAD'
    elif home_wp < away_wp:
        dog, fav = home, away
        dog_wp, fav_wp = home_wp, away_wp
        dog_loc = 'HOME'
    else:
        # Equal win% — still show it
        dog, fav = away, home
        dog_wp, fav_wp = away_wp, home_wp
        dog_loc = 'ROAD'
    return {
        'dog': dog, 'fav': fav,
        'dog_wp': dog_wp, 'fav_wp': fav_wp,
        'dog_loc': dog_loc,
        'matchup': f"{away} @ {home}"
    }


def above_line(dog_wp, fav_wp, x1, y1, x2, y2):
    """Check if point (dog_wp, fav_wp) is above the dividing line."""
    if x2 == x1:
        return dog_wp > x1
    slope = (y2 - y1) / (x2 - x1)
    y_line = y1 + slope * (dog_wp - x1)
    return fav_wp > y_line


# ── PAGE HEADER ───────────────────────────────────────────────────
st.title("📊 Series Dog Scatter Analysis")
st.markdown("Dog Win% vs Favorite Win% — Historical series results with today's games overlaid.")
st.markdown("---")

# ── LOAD DATA ─────────────────────────────────────────────────────
# Try local file first (dev), then ask for upload (Streamlit Cloud)
df = None

if os.path.exists(MARTINGALE_FILE):
    try:
        with open(MARTINGALE_FILE, 'rb') as f:
            df = load_series_data(f.read())
    except Exception as e:
        st.warning(f"Could not load local file: {e}")

if df is None:
    st.info("📤 Upload `Martingale_Series_Analysis.xlsx` to load historical data.")
    uploaded_file = st.file_uploader(
        "Upload Martingale_Series_Analysis.xlsx",
        type=['xlsx'],
        key='martingale_upload'
    )
    if uploaded_file is not None:
        try:
            df = load_series_data(uploaded_file.read())
            st.success(f"✓ Loaded {len(df):,} series records")
        except Exception as e:
            st.error(f"Could not read file: {e}")
            st.stop()
    else:
        st.stop()


# ── SIDEBAR CONTROLS ──────────────────────────────────────────────
st.sidebar.header("⚙️ Controls")

report_date = st.sidebar.date_input("Report Date", value=date.today())
report_date_str = report_date.isoformat()

st.sidebar.markdown("---")
st.sidebar.subheader("📅 Year Filter")
available_years = sorted(df['Year'].dropna().unique().tolist())
selected_years = st.sidebar.multiselect(
    "Show years:", available_years, default=available_years
)

st.sidebar.markdown("---")
st.sidebar.subheader("📍 Dog Location")
dog_locations = st.sidebar.multiselect(
    "Dog location:", ['ROAD', 'HOME'], default=['ROAD', 'HOME']
)

st.sidebar.markdown("---")
st.sidebar.subheader("📏 Dividing Line")
st.sidebar.caption("Define the dividing line. Points ABOVE the line are flagged as high-confidence.")
lx1 = st.sidebar.number_input("Line start X (Dog Win%)", value=0.450, min_value=0.0, max_value=1.0, step=0.01, format="%.3f")
ly1 = st.sidebar.number_input("Line start Y (Fav Win%)", value=0.580, min_value=0.0, max_value=1.0, step=0.01, format="%.3f")
lx2 = st.sidebar.number_input("Line end X (Dog Win%)",   value=0.750, min_value=0.0, max_value=1.0, step=0.01, format="%.3f")
ly2 = st.sidebar.number_input("Line end Y (Fav Win%)",   value=0.620, min_value=0.0, max_value=1.0, step=0.01, format="%.3f")

st.sidebar.markdown("---")
st.sidebar.subheader("🎨 Display")
show_today = st.sidebar.checkbox("Show today's games", value=True)
show_line  = st.sidebar.checkbox("Show dividing line", value=True)
point_size = st.sidebar.slider("Point size", 6, 20, 10)

# ── FILTER DATA ───────────────────────────────────────────────────
filtered = df[
    (df['Year'].isin(selected_years)) &
    (df['Dog_Location'].isin(dog_locations))
].copy()

if filtered.empty:
    st.warning("No data matches the current filters.")
    st.stop()

# Classify above/below line
filtered['above_line'] = filtered.apply(
    lambda r: above_line(r['Dog_WinPct'], r['Fav_WinPct'], lx1, ly1, lx2, ly2), axis=1
)

wins = filtered[filtered['Won_1_Game'] == 'YES']
losses = filtered[filtered['Won_1_Game'] == 'NO']

# ── BUILD SCATTER PLOT ────────────────────────────────────────────
fig = go.Figure()

# Losses (orange)
fig.add_trace(go.Scatter(
    x=losses['Dog_WinPct'],
    y=losses['Fav_WinPct'],
    mode='markers',
    name='Loss (Dog lost series)',
    marker=dict(
        color='rgba(255, 127, 14, 0.65)',
        size=point_size,
        line=dict(color='rgba(200,90,0,0.8)', width=0.5),
        symbol='circle'
    ),
    customdata=np.stack([
        losses['Year'].astype(str),
        losses['Road_Team'],
        losses['Home_Team'],
        losses['Dog_Location'],
        losses['Avg_Line'].round(1).astype(str),
        losses['Results']
    ], axis=-1),
    hovertemplate=(
        "<b>LOSS</b><br>"
        "Year: %{customdata[0]}<br>"
        "Matchup: %{customdata[1]} @ %{customdata[2]}<br>"
        "Dog Location: %{customdata[3]}<br>"
        "Dog Win%%: %{x:.3f}<br>"
        "Fav Win%%: %{y:.3f}<br>"
        "Avg Line: +%{customdata[4]}<br>"
        "Results: %{customdata[5]}<extra></extra>"
    )
))

# Wins (blue)
fig.add_trace(go.Scatter(
    x=wins['Dog_WinPct'],
    y=wins['Fav_WinPct'],
    mode='markers',
    name='Win (Dog won ≥1 game)',
    marker=dict(
        color='rgba(31, 119, 180, 0.65)',
        size=point_size,
        line=dict(color='rgba(10,80,150,0.8)', width=0.5),
        symbol='circle'
    ),
    customdata=np.stack([
        wins['Year'].astype(str),
        wins['Road_Team'],
        wins['Home_Team'],
        wins['Dog_Location'],
        wins['Avg_Line'].round(1).astype(str),
        wins['Results']
    ], axis=-1),
    hovertemplate=(
        "<b>WIN</b><br>"
        "Year: %{customdata[0]}<br>"
        "Matchup: %{customdata[1]} @ %{customdata[2]}<br>"
        "Dog Location: %{customdata[3]}<br>"
        "Dog Win%%: %{x:.3f}<br>"
        "Fav Win%%: %{y:.3f}<br>"
        "Avg Line: +%{customdata[4]}<br>"
        "Results: %{customdata[5]}<extra></extra>"
    )
))

# ── DIVIDING LINE ─────────────────────────────────────────────────
if show_line:
    fig.add_trace(go.Scatter(
        x=[lx1, lx2],
        y=[ly1, ly2],
        mode='lines+markers',
        name=f'Dividing line ({lx1:.3f},{ly1:.3f})→({lx2:.3f},{ly2:.3f})',
        line=dict(color='red', width=2.5, dash='dash'),
        marker=dict(color='red', size=8, symbol='x'),
        hoverinfo='skip'
    ))

# ── TODAY'S GAMES ─────────────────────────────────────────────────
if show_today:
    winpcts  = get_team_winpcts(report_date_str)
    today_games = get_todays_games(report_date_str)

    today_points = []
    for g in today_games:
        info = determine_dog_fav(g, winpcts)
        if info:
            is_above = above_line(info['dog_wp'], info['fav_wp'], lx1, ly1, lx2, ly2)
            info['above_line'] = is_above
            today_points.append(info)

    if today_points:
        tx = [p['dog_wp'] for p in today_points]
        ty = [p['fav_wp'] for p in today_points]
        tlabels = [
            f"{'⭐ HIGH CONFIDENCE' if p['above_line'] else '📍 Below line'}<br>"
            f"Matchup: {p['matchup']}<br>"
            f"Dog: {p['dog']} ({p['dog_wp']:.3f})<br>"
            f"Fav: {p['fav']} ({p['fav_wp']:.3f})<br>"
            f"Dog location: {p['dog_loc']}"
            for p in today_points
        ]
        tcolors = ['gold' if p['above_line'] else 'lime' for p in today_points]
        tsymbols = ['star' if p['above_line'] else 'diamond' for p in today_points]

        fig.add_trace(go.Scatter(
            x=tx, y=ty,
            mode='markers+text',
            name="Today's games",
            marker=dict(
                color=tcolors,
                size=18,
                symbol=tsymbols,
                line=dict(color='black', width=1.5)
            ),
            text=[p['dog'].split()[-1] for p in today_points],
            textposition='top center',
            textfont=dict(size=10, color='black'),
            customdata=tlabels,
            hovertemplate="%{customdata}<extra></extra>"
        ))

# ── LAYOUT ────────────────────────────────────────────────────────
fig.update_layout(
    title=dict(
        text=f"Dog Win% vs Fav Win%  |  {len(filtered):,} series  |  {selected_years[0] if len(selected_years)==1 else f'{min(selected_years)}–{max(selected_years)}'}",
        font=dict(size=16)
    ),
    xaxis=dict(
        title="Dog Win% (season record of underdog)",
        tickformat=".0%",
        range=[0.20, 0.75],
        showgrid=True, gridcolor='rgba(200,200,200,0.4)',
        zeroline=False
    ),
    yaxis=dict(
        title="Fav Win% (season record of favorite)",
        tickformat=".0%",
        range=[0.30, 0.75],
        showgrid=True, gridcolor='rgba(200,200,200,0.4)',
        zeroline=False
    ),
    legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
    hovermode='closest',
    height=620,
    margin=dict(l=60, r=30, t=80, b=60),
    plot_bgcolor='white',
    paper_bgcolor='white'
)

st.plotly_chart(fig, use_container_width=True)

# ── RECORD ABOVE / BELOW LINE ─────────────────────────────────────
st.markdown("---")
st.subheader("📊 Win/Loss Record by Zone")

above = filtered[filtered['above_line']]
below = filtered[~filtered['above_line']]

def zone_stats(subset, label):
    total = len(subset)
    w = (subset['Won_1_Game'] == 'YES').sum()
    l = (subset['Won_1_Game'] == 'NO').sum()
    pct = w / total * 100 if total > 0 else 0
    return label, w, l, total, pct

c1, c2, c3 = st.columns(3)

for col, (label, w, l, total, pct), color in zip(
    [c1, c2, c3],
    [
        zone_stats(above, "⭐ Above Line"),
        zone_stats(below, "📉 Below Line"),
        zone_stats(filtered, "📊 All Filtered"),
    ],
    ['success', 'error', 'info']
):
    getattr(col, color)(f"**{label}**  \n{w}W – {l}L  ({pct:.1f}% win rate)  \nn={total}")

# ── YEAR-BY-YEAR BREAKDOWN (above line only) ──────────────────────
if not above.empty:
    st.markdown("---")
    st.subheader("📅 Year-by-Year  (Above Line only)")
    yr_rows = []
    for yr in sorted(above['Year'].dropna().unique()):
        ydf = above[above['Year'] == yr]
        w = (ydf['Won_1_Game'] == 'YES').sum()
        l = (ydf['Won_1_Game'] == 'NO').sum()
        t = w + l
        yr_rows.append({'Year': int(yr), 'W': int(w), 'L': int(l),
                        'Total': int(t), 'Win%': f"{w/t*100:.1f}%" if t > 0 else "—"})
    st.dataframe(pd.DataFrame(yr_rows), hide_index=True, use_container_width=False)

# ── TODAY'S HIGH-CONFIDENCE GAMES ─────────────────────────────────
if show_today and 'today_points' in dir() and today_points:
    st.markdown("---")
    st.subheader(f"⭐ Today's Games  ({report_date.strftime('%A, %B %d, %Y')})")

    today_rows = []
    for p in today_points:
        today_rows.append({
            'Matchup':    p['matchup'],
            'Dog':        p['dog'],
            'Dog Win%':   f"{p['dog_wp']:.3f}",
            'Fav':        p['fav'],
            'Fav Win%':   f"{p['fav_wp']:.3f}",
            'Dog Loc':    p['dog_loc'],
            'Zone':       '⭐ ABOVE LINE' if p['above_line'] else 'Below line',
        })

    today_df = pd.DataFrame(today_rows)

    def highlight_zone(row):
        if row['Zone'] == '⭐ ABOVE LINE':
            return ['background-color: #FFF9C4; font-weight: bold'] * len(row)
        return [''] * len(row)

    st.dataframe(
        today_df.style.apply(highlight_zone, axis=1),
        hide_index=True,
        use_container_width=True
    )

    above_today = [p for p in today_points if p['above_line']]
    if above_today:
        st.success(f"✅ {len(above_today)} game(s) fall above the line today — high-confidence zone")
    else:
        st.info("No games fall above the dividing line today")

elif show_today:
    st.markdown("---")
    st.info(f"No games scheduled for {report_date.strftime('%B %d, %Y')} or unable to fetch schedule.")

# ── INTERPRETATION GUIDE ──────────────────────────────────────────
with st.expander("ℹ️ How to read this chart", expanded=False):
    st.markdown("""
    **Axes:**
    - **X-axis (Dog Win%)** — The underdog's season win percentage at time of series
    - **Y-axis (Fav Win%)** — The favorite's season win percentage at time of series

    **Points:**
    - 🔵 **Blue** — The underdog won at least one game in the series
    - 🟠 **Orange** — The underdog lost all games (favorite swept)
    - ⭐ **Gold star** — Today's game, falls ABOVE the dividing line (high-confidence zone)
    - 💚 **Green diamond** — Today's game, falls below the line

    **Dividing Line:**
    - Adjust the line coordinates in the sidebar
    - The record above vs below updates automatically
    - **Above line = both teams are strong** — dog isn't outclassed, favorite is genuinely good

    **Key insight:**
    - Upper-right cluster = competitive matchups where the dog has a real chance
    - Lower-left = weak dog vs strong fav — avoid betting the dog
    - The line helps separate these zones — adjust it to find the optimal split

    **Sample size warning:**
    - ~11 plays/year above the line — build conviction over 2–3 seasons before increasing stake size
    """)
