"""
MLB Daily Betting Report - Streamlit Web App
Client opens URL in browser, enters odds, clicks button, downloads Excel report.
No Python installation needed on client side.
"""

import streamlit as st
import pandas as pd
import requests
import xlsxwriter
import io
import os
import re
from datetime import date, datetime, timedelta
from openpyxl import load_workbook

# ── Page config ───────────────────────────────────────────────────
st.set_page_config(
    page_title="MLB Daily Betting Report",
    page_icon="⚾",
    layout="wide"
)

# ── Import core logic from daily_report.py ────────────────────────
from daily_report import (
    load_historical_data, fetch_recent_results, compute_all_states,
    fetch_todays_schedule, get_team_state, build_game_row,
    check_scenarios, API_TO_CANONICAL, DIVISIONS, SCENARIO_DEFS,
    NEEDS_OPP_STREAK, NEEDS_OPP_ROAD_WP, title_case, fmt_line
)

# ── Cache heavy data loading ──────────────────────────────────────
@st.cache_data(ttl=3600, show_spinner="Loading historical data...")
def load_enriched_data(report_date_str):
    """Load and enrich all historical data. Cached for 1 hour."""
    df = load_historical_data()
    df = fetch_recent_results(df, date.fromisoformat(report_date_str))
    enriched = compute_all_states(df)
    return enriched

@st.cache_data(ttl=1800, show_spinner="Fetching today's schedule...")
def get_schedule(report_date_str):
    return fetch_todays_schedule(date.fromisoformat(report_date_str))

# ── Build Excel report in memory ──────────────────────────────────
def build_report_bytes(games, triggers, report_date, odds):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})

    NAVY='#1B2A4A'; STEEL='#2E5F8A'; LIGHT_STEEL='#D0E4F5'
    FADE_RED='#C00000'; FADE_BG='#FFE7E7'
    WATCH_AMB='#7D5A00'; WATCH_BG='#FFF3CC'
    AWAY_BG='#F0F5FB'; HOME_BG='#FFFFFF'
    GAME_HDR_BG='#1B2A4A'; SCEN_BG='#E8F0FA'; GRAY_TEXT='#666666'

    f_banner   = wb.add_format({'bold':True,'font_size':16,'font_name':'Calibri','font_color':'white','bg_color':NAVY,'align':'center','valign':'vcenter'})
    f_subtitle = wb.add_format({'font_size':11,'font_name':'Calibri','italic':True,'font_color':LIGHT_STEEL,'bg_color':NAVY,'align':'center','valign':'vcenter'})
    f_col_hdr  = wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':'white','bg_color':STEEL,'align':'center','valign':'vcenter','border':1,'border_color':'#1A4A72'})
    f_game_hdr = wb.add_format({'bold':True,'font_size':11,'font_name':'Calibri','font_color':'white','bg_color':GAME_HDR_BG,'align':'left','valign':'vcenter','left':2,'left_color':STEEL,'top':1,'top_color':'#0A1828','bottom':1,'bottom_color':'#0A1828'})
    f_away_lbl = wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':STEEL,'bg_color':AWAY_BG,'align':'left','valign':'vcenter','left':2,'left_color':STEEL,'border':1,'border_color':'#B8C9DC','indent':1})
    f_home_lbl = wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':'#333333','bg_color':HOME_BG,'align':'left','valign':'vcenter','left':2,'left_color':STEEL,'border':1,'border_color':'#B8C9DC','indent':1})
    f_away_c   = wb.add_format({'font_size':10,'font_name':'Calibri','bg_color':AWAY_BG,'align':'center','valign':'vcenter','border':1,'border_color':'#B8C9DC'})
    f_home_c   = wb.add_format({'font_size':10,'font_name':'Calibri','bg_color':HOME_BG,'align':'center','valign':'vcenter','border':1,'border_color':'#B8C9DC'})
    f_fade_a   = wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':FADE_RED,'bg_color':FADE_BG,'align':'center','valign':'vcenter','border':1,'border_color':'#E8AAAA'})
    f_fade_h   = wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':FADE_RED,'bg_color':FADE_BG,'align':'center','valign':'vcenter','border':1,'border_color':'#E8AAAA'})
    f_watch    = wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':WATCH_AMB,'bg_color':WATCH_BG,'align':'center','valign':'vcenter','border':1,'border_color':'#D4B800'})
    f_no_play_a= wb.add_format({'italic':True,'font_size':9,'font_name':'Calibri','font_color':GRAY_TEXT,'bg_color':AWAY_BG,'align':'center','valign':'vcenter','border':1,'border_color':'#B8C9DC'})
    f_no_play_h= wb.add_format({'italic':True,'font_size':9,'font_name':'Calibri','font_color':GRAY_TEXT,'bg_color':HOME_BG,'align':'center','valign':'vcenter','border':1,'border_color':'#B8C9DC'})
    f_scen_a   = wb.add_format({'font_size':9,'font_name':'Calibri','italic':True,'font_color':STEEL,'bg_color':SCEN_BG,'align':'left','valign':'vcenter','border':1,'border_color':'#B8C9DC','text_wrap':True,'indent':1})
    f_scen_h   = wb.add_format({'font_size':9,'font_name':'Calibri','italic':True,'font_color':'#555555','bg_color':HOME_BG,'align':'left','valign':'vcenter','border':1,'border_color':'#B8C9DC','text_wrap':True,'indent':1})
    f_no_scen_a= wb.add_format({'italic':True,'font_size':9,'font_color':GRAY_TEXT,'bg_color':SCEN_BG,'align':'left','valign':'vcenter','border':1,'border_color':'#B8C9DC','indent':1})
    f_no_scen_h= wb.add_format({'italic':True,'font_size':9,'font_color':GRAY_TEXT,'bg_color':HOME_BG,'align':'left','valign':'vcenter','border':1,'border_color':'#B8C9DC','indent':1})
    f_no_games = wb.add_format({'italic':True,'font_size':11,'font_color':GRAY_TEXT,'align':'center','valign':'vcenter'})
    f_sum_lbl  = wb.add_format({'bold':True,'font_size':11,'font_name':'Calibri','font_color':NAVY,'bg_color':'#EDF3FB','border':1,'border_color':'#B8C9DC','align':'left','valign':'vcenter','indent':1})
    f_sum_val  = wb.add_format({'bold':True,'font_size':14,'font_name':'Calibri','font_color':NAVY,'bg_color':'#EDF3FB','border':1,'border_color':'#B8C9DC','align':'center','valign':'vcenter'})
    f_sum_fl   = wb.add_format({'bold':True,'font_size':11,'font_name':'Calibri','font_color':FADE_RED,'bg_color':FADE_BG,'border':1,'border_color':'#E8AAAA','align':'left','valign':'vcenter','indent':1})
    f_sum_fv   = wb.add_format({'bold':True,'font_size':14,'font_name':'Calibri','font_color':FADE_RED,'bg_color':FADE_BG,'border':1,'border_color':'#E8AAAA','align':'center','valign':'vcenter'})
    f_sum_wl   = wb.add_format({'bold':True,'font_size':11,'font_name':'Calibri','font_color':WATCH_AMB,'bg_color':WATCH_BG,'border':1,'border_color':'#D4B800','align':'left','valign':'vcenter','indent':1})
    f_sum_wv   = wb.add_format({'bold':True,'font_size':14,'font_name':'Calibri','font_color':WATCH_AMB,'bg_color':WATCH_BG,'border':1,'border_color':'#D4B800','align':'center','valign':'vcenter'})

    COL_WIDTHS  = [5, 8, 26, 10, 22, 48]
    COL_HEADERS = ['#', 'H/A', 'Team', 'Odds', 'Play', 'Scenario(s)']
    NCOLS = 6

    def write_tab(ws, tab_label, verdict_filter, tab_color):
        ws.set_tab_color(tab_color)
        for ci, w in enumerate(COL_WIDTHS):
            ws.set_column(ci, ci, w)
        ws.set_row(0, 30); ws.set_row(1, 18)
        ws.merge_range(0,0,0,NCOLS-1, f'⚾  MLB DAILY BETTING REPORT  —  {tab_label}', f_banner)
        ws.merge_range(1,0,1,NCOLS-1, f'{report_date.strftime("%A, %B %d, %Y")}  |  Generated {datetime.now().strftime("%I:%M %p")}', f_subtitle)
        ws.set_row(2, 20)
        for ci, h in enumerate(COL_HEADERS):
            ws.write(2, ci, h, f_col_hdr)
        ws.freeze_panes(3, 0)

        row = 3; written = 0
        for g_idx, g in enumerate(games, 1):
            away, home = g['away_team'], g['home_team']
            away_tc, home_tc = title_case(away), title_case(home)
            at = [t for t in triggers if t['team']==away and t['opponent']==home and t['verdict']==verdict_filter]
            ht = [t for t in triggers if t['team']==home and t['opponent']==away and t['verdict']==verdict_filter]
            if not at and not ht: continue
            written += 1

            ws.set_row(row, 18)
            ws.merge_range(row,0,row,NCOLS-1, f'  Game {written}  ▸  {away_tc}  @  {home_tc}', f_game_hdr)
            row += 1

            ws.set_row(row, 20)
            a_scen = '  |  '.join(f"#{t['scenario_id']} {t['scenario']}" for t in at) if at else '—'
            a_pf   = f_fade_a if (at and at[0]['verdict']=='CLEAR FADE') else (f_watch if at else f_no_play_a)
            ws.write(row,0,g_idx,f_away_c); ws.write(row,1,'AWAY',f_away_c)
            ws.write(row,2,away_tc,f_away_lbl)
            ws.write(row,3,fmt_line(odds.get(away)),f_away_c)
            ws.write(row,4,at[0]['play'] if at else '—',a_pf)
            ws.write(row,5,a_scen,f_scen_a if at else f_no_scen_a)
            row += 1

            ws.set_row(row, 20)
            h_scen = '  |  '.join(f"#{t['scenario_id']} {t['scenario']}" for t in ht) if ht else '—'
            h_pf   = f_fade_h if (ht and ht[0]['verdict']=='CLEAR FADE') else (f_watch if ht else f_no_play_h)
            ws.write(row,0,g_idx,f_home_c); ws.write(row,1,'HOME',f_home_c)
            ws.write(row,2,home_tc,f_home_lbl)
            ws.write(row,3,fmt_line(odds.get(home)),f_home_c)
            ws.write(row,4,ht[0]['play'] if ht else '—',h_pf)
            ws.write(row,5,h_scen,f_scen_h if ht else f_no_scen_h)
            row += 2

        if written == 0:
            ws.merge_range(row,0,row,NCOLS-1,f'No {verdict_filter.title()} scenarios triggered today.',f_no_games)
        return written

    ws1 = wb.add_worksheet('🔴 Clear Fade')
    ws2 = wb.add_worksheet('🟡 Inconsistent')
    n1 = write_tab(ws1,'CLEAR FADE','CLEAR FADE','#FF0000')
    n2 = write_tab(ws2,'INCONSISTENT','INCONSISTENT','#FFD700')

    # Summary tab
    ws3 = wb.add_worksheet('📊 Summary'); ws3.set_tab_color('#1B2A4A')
    ws3.set_column(0,0,35); ws3.set_column(1,1,18); ws3.set_column(2,2,35); ws3.set_column(3,3,18)
    ws3.set_row(0,30); ws3.set_row(1,18)
    ws3.merge_range(0,0,0,3,'⚾  MLB DAILY BETTING REPORT  —  SUMMARY',f_banner)
    ws3.merge_range(1,0,1,3,f'{report_date.strftime("%A, %B %d, %Y")}  |  Generated {datetime.now().strftime("%I:%M %p")}',f_subtitle)
    n_ft = sum(1 for t in triggers if t['verdict']=='CLEAR FADE')
    n_it = sum(1 for t in triggers if t['verdict']=='INCONSISTENT')
    ws3.set_row(3,28); ws3.set_row(4,28)
    ws3.write(3,0,'Total Games Today',f_sum_lbl); ws3.write(3,1,len(games),f_sum_val)
    ws3.write(3,2,'Games With Triggers',f_sum_lbl); ws3.write(3,3,n1+n2,f_sum_val)
    ws3.write(4,0,'🔴  Clear Fade Triggers',f_sum_fl); ws3.write(4,1,n_ft,f_sum_fv)
    ws3.write(4,2,'🟡  Inconsistent Triggers',f_sum_wl); ws3.write(4,3,n_it,f_sum_wv)
    ws3.merge_range(6,0,6,3,'TOP PLAYS TODAY',f_col_hdr)
    ws3.set_row(7,20)
    for ci,h in enumerate(['Team','Odds','Play','Scenario']):
        ws3.write(7,ci,h,f_col_hdr)
    sr = 8
    for t in [x for x in triggers if x['verdict']=='CLEAR FADE']:
        ws3.write(sr,0,title_case(t['team']),f_sum_fl); ws3.write(sr,1,fmt_line(t['line']),f_sum_fv)
        ws3.write(sr,2,t['play'],f_sum_fl); ws3.write(sr,3,f"#{t['scenario_id']} {t['scenario']}",f_sum_fl); sr+=1
    for t in [x for x in triggers if x['verdict']=='INCONSISTENT']:
        ws3.write(sr,0,title_case(t['team']),f_sum_wl); ws3.write(sr,1,fmt_line(t['line']),f_sum_wv)
        ws3.write(sr,2,t['play'],f_sum_wl); ws3.write(sr,3,f"#{t['scenario_id']} {t['scenario']}",f_sum_wl); sr+=1
    ws3.freeze_panes(8,0)

    wb.close(); output.seek(0)
    return output.getvalue(), n1, n2


# ── MAIN UI ───────────────────────────────────────────────────────

st.title("⚾ MLB Daily Betting Report")
st.markdown("---")

# Date selector
col1, col2 = st.columns([2,4])
with col1:
    report_date = st.date_input("Report Date", value=date.today())

report_date_str = report_date.isoformat()

# Load data
with st.spinner("Loading team data and schedule..."):
    try:
        enriched = load_enriched_data(report_date_str)
        games    = get_schedule(report_date_str)
    except Exception as e:
        st.error(f"Error loading data: {e}")
        st.stop()

if not games:
    st.warning("No games scheduled for this date.")
    st.stop()

st.success(f"✓ {len(games)} games scheduled for {report_date.strftime('%A, %B %d, %Y')}")
st.markdown("---")

# ── ODDS INPUT TABLE ──────────────────────────────────────────────
st.subheader("📋 Enter Today's Moneylines")
st.caption("Enter the moneyline for each team (e.g. +130 or -150). Leave blank if unavailable.")

# Build odds input form
odds_data = []
for g in games:
    odds_data.append({
        'Away Team': g['away_team'],
        'Home Team': g['home_team'],
        'Away Line': '',
        'Home Line': '',
    })

odds_df = pd.DataFrame(odds_data)
edited = st.data_editor(
    odds_df,
    column_config={
        'Away Team': st.column_config.TextColumn('Away Team', disabled=True, width='large'),
        'Home Team': st.column_config.TextColumn('Home Team', disabled=True, width='large'),
        'Away Line': st.column_config.NumberColumn('Away Line', help='e.g. 130 or -150', width='medium'),
        'Home Line': st.column_config.NumberColumn('Home Line', help='e.g. -130 or 150', width='medium'),
    },
    hide_index=True,
    use_container_width=True,
)

st.markdown("---")

# ── GENERATE REPORT ───────────────────────────────────────────────
if st.button("⚾ Generate Daily Report", type="primary", use_container_width=True):
    # Parse odds from editor
    odds = {}
    for _, row in edited.iterrows():
        if row['Away Line'] not in [None,'','nan']:
            try: odds[str(row['Away Team']).upper()] = int(row['Away Line'])
            except: pass
        if row['Home Line'] not in [None,'','nan']:
            try: odds[str(row['Home Team']).upper()] = int(row['Home Line'])
            except: pass

    if not odds:
        st.warning("Please enter at least some moneylines before generating the report.")
        st.stop()

    with st.spinner("Running scenario analysis..."):
        # Build opponent streak and road win% lookups
        opp_streaks = {}
        opp_road_wpct = {}
        for team in API_TO_CANONICAL.values():
            tdf = enriched[(enriched['team']==team) & (enriched['date'].dt.date < report_date)]
            if not tdf.empty:
                last = tdf.iloc[-1]
                opp_streaks[team] = last['streak_before'] + (1 if last['result']=='W' else -1)
                road = tdf[tdf['home_away']=='away']
                if not road.empty:
                    rw = (road['result']=='W').sum()
                    rl = (road['result']=='L').sum()
                    opp_road_wpct[team] = rw/(rw+rl) if (rw+rl)>0 else None

        # Check scenarios for each game
        all_triggers = []
        for g in games:
            away, home = g['away_team'], g['home_team']
            away_state = get_team_state(enriched, away, report_date)
            home_state = get_team_state(enriched, home, report_date)
            away_row = build_game_row(away_state, 'away', home, odds.get(away))
            home_row = build_game_row(home_state, 'home', away, odds.get(home))
            triggers = check_scenarios([away_row, home_row], opp_streaks, opp_road_wpct)
            all_triggers.extend(triggers)

        # Build Excel
        excel_bytes, n_fade, n_inc = build_report_bytes(games, all_triggers, report_date, odds)

    # Show summary
    st.markdown("---")
    st.subheader("📊 Results")
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Games", len(games))
    m2.metric("🔴 Clear Fade Triggers", sum(1 for t in all_triggers if t['verdict']=='CLEAR FADE'))
    m3.metric("🟡 Inconsistent Triggers", sum(1 for t in all_triggers if t['verdict']=='INCONSISTENT'))

    # Show triggers in UI
    if all_triggers:
        st.markdown("### Clear Fade Plays")
        fade_list = [t for t in all_triggers if t['verdict']=='CLEAR FADE']
        if fade_list:
            fade_df = pd.DataFrame([{
                'Team': t['team'],
                'vs': t['opponent'],
                'H/A': t['home_away'].upper(),
                'Line': t['line'],
                'Play': t['play'],
                'Scenario': f"#{t['scenario_id']} {t['scenario']}",
            } for t in fade_list])
            st.dataframe(fade_df, use_container_width=True, hide_index=True)
        else:
            st.info("No Clear Fade scenarios triggered today.")

        st.markdown("### Inconsistent Plays")
        inc_list = [t for t in all_triggers if t['verdict']=='INCONSISTENT']
        if inc_list:
            inc_df = pd.DataFrame([{
                'Team': t['team'],
                'vs': t['opponent'],
                'H/A': t['home_away'].upper(),
                'Line': t['line'],
                'Play': t['play'],
                'Scenario': f"#{t['scenario_id']} {t['scenario']}",
            } for t in inc_list])
            st.dataframe(inc_df, use_container_width=True, hide_index=True)
        else:
            st.info("No Inconsistent scenarios triggered today.")

    # Download button
    st.markdown("---")
    fname = f'MLB_Daily_Report_{report_date.strftime("%Y-%m-%d")}.xlsx'
    st.download_button(
        label="📥 Download Excel Report",
        data=excel_bytes,
        file_name=fname,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
        type="primary",
    )
