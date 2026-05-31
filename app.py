"""
PCO TechOps - Shift Capacity Planning Web Application v5
========================================================
Enhanced Streamlit app with:
- Effective_Workload_% (effort-based utilization)
- Expected_Volume_Trend (WoW growth from incoming volumes)
- 65% utilization target for Expected_Util_%
- Expected SR Overhead Factor (coordination hours)
- MMI-aware: MMI SDs do not consume shift capacity
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.graph_objects as go
import io
import re
import warnings
warnings.filterwarnings('ignore')

# =============================================================================
# PAGE CONFIG
# =============================================================================
st.set_page_config(
    page_title="PCO TechOps - Shift Capacity Planning",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =============================================================================
# CONSTANTS
# =============================================================================

SHIFT_HOURS = {'S1': 8, 'S2': 9, 'S3': 7}
SHIFT_TIMING = {'S1': '6AM-2PM IST', 'S2': '2PM-11PM IST', 'S3': '11PM-6AM IST'}

# Flat 65% target for all tracks
TARGET_UTIL = 0.65

# Non-MMI Effort/Duration ratio (37.5%)
# MMI SDs get slots even when capacity is exceeded - they don't consume shift capacity
EFFORT_RATIO_NON_MMI = 0.375

# SR coordination overhead: 15 minutes per internal SR
SR_COORDINATION_MINUTES = 15

# Trend cap
TREND_CAP = 1.15


# =============================================================================
# DATA LOADING
# =============================================================================

def load_excel_data(uploaded_file):
    """Load partner data from SPC Excel export, auto-detecting format."""
    xls = pd.ExcelFile(uploaded_file)
    tracks_found = {}
    
    for sheet in xls.sheet_names:
        # Try multiple header positions
        for header_row in [0, 4]:
            try:
                raw = pd.read_excel(xls, sheet_name=sheet, header=header_row)
                raw = raw.dropna(axis=1, how='all')
                
                # Remove unnamed leading column
                if raw.columns[0] != 'Team ID':
                    if 'Team ID' in raw.columns:
                        pass
                    elif str(raw.columns[0]).startswith('Unnamed') or pd.isna(raw.columns[0]):
                        raw = raw.iloc[:, 1:]
                    else:
                        continue
                
                if 'Team ID' not in raw.columns or 'Shift' not in raw.columns:
                    continue
                
                raw = raw[raw['Shift'].isin(['S1', 'S2', 'S3'])]
                if len(raw) < 20:
                    continue
                
                # Group by Team ID
                for tid, tdf in raw.groupby('Team ID'):
                    if len(tdf) < 20:
                        continue
                    
                    team_name = tdf['Team Name'].iloc[0] if 'Team Name' in tdf.columns else ''
                    track = detect_track(team_name)
                    partner = detect_partner(team_name)
                    
                    key = f"{partner}|{track}|{tid}"
                    if key not in tracks_found:
                        tracks_found[key] = {
                            'data': tdf,
                            'team_id': tid,
                            'team_name': team_name,
                            'track': track,
                            'partner': partner,
                            'sheet': sheet,
                            'rows': len(tdf),
                        }
                    else:
                        tracks_found[key]['data'] = pd.concat(
                            [tracks_found[key]['data'], tdf], ignore_index=True)
                        tracks_found[key]['rows'] = len(tracks_found[key]['data'])
                
                break  # Found valid data at this header position
            except Exception:
                continue
    
    return tracks_found


def detect_track(team_name):
    """Detect track from team name."""
    name = str(team_name).upper()
    parts = name.split(':')
    if len(parts) > 1:
        after_colon = parts[1]
        if 'SM' in after_colon:
            return 'SM'
        elif 'DB' in after_colon:
            return 'DB'
    return 'Basis'


def detect_partner(team_name):
    """Extract partner name from team name."""
    name = str(team_name)
    for prefix in ['PCO SR Delivery: Basis - ', 'PCO SR Delivery: SM L2 - ',
                   'PCO SR Delivery: SM - ', 'PCO SR Delivery: DB - ',
                   'PCO SR Delivery: ', 'ECS SR Delivery: Basis - ',
                   'ECS SR Delivery: SM L2 - ', 'ECS SR Delivery: SM - ',
                   'ECS SR Delivery: DB - ', 'ECS SR Delivery: ']:
        if name.startswith(prefix):
            return name[len(prefix):]
    if ' - ' in name:
        return name.split(' - ')[-1]
    return name


def prepare_data(df):
    """Add computed columns for analysis."""
    df = df.copy()
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df = df.dropna(subset=['Date'])
    
    df['Ʃ Total Demand Hour(s)'] = pd.to_numeric(df['Ʃ Total Demand Hour(s)'], errors='coerce')
    df['Ʃ Total Capacity Hour(s)'] = pd.to_numeric(df['Ʃ Total Capacity Hour(s)'], errors='coerce')
    df = df.dropna(subset=['Ʃ Total Demand Hour(s)', 'Ʃ Total Capacity Hour(s)'])
    
    df['DayOfWeekNum'] = df['Date'].dt.dayofweek
    df['DayOfWeek'] = df['Date'].dt.day_name()
    df['WeekOfMonth'] = ((df['Date'].dt.day - 1) // 7) + 1
    df['YearMonth'] = df['Date'].dt.to_period('M')
    df['WeekOfYear'] = df['Date'].dt.isocalendar().week.astype(int)
    df['Shift_Hours'] = df['Shift'].map(SHIFT_HOURS)
    df['Executors'] = df['Ʃ Total Capacity Hour(s)'] / df['Shift_Hours']
    df['Utilization_Pct'] = (df['Ʃ Total Demand Hour(s)'] / df['Ʃ Total Capacity Hour(s)']) * 100
    
    return df


# =============================================================================
# VOLUME TREND & SR OVERHEAD COMPUTATION
# =============================================================================

def compute_volume_trend(volume_file, partner_name):
    """
    Compute Expected_Volume_Trend from Partners_Incoming_volumes.xlsx.
    Returns weekly growth rate (%) based on linear regression on CW02-CW15.
    """
    if volume_file is None:
        return 0.0
    
    try:
        vol_xls = pd.ExcelFile(volume_file)
    except Exception:
        return 0.0
    
    target_sheet = find_partner_sheet(vol_xls.sheet_names, partner_name)
    if target_sheet is None:
        return 0.0
    
    data = pd.read_excel(vol_xls, sheet_name=target_sheet)
    cw_cols = [c for c in data.columns if c.startswith('CW')]
    csr_row = data[data['Request Type'] == 'Customer Service Request']
    
    if len(csr_row) == 0 or len(cw_cols) < 3:
        return 0.0
    
    # Exclude CW01 (holiday anomaly)
    steady_cols = [c for c in cw_cols if c != 'CW01']
    vals = csr_row[steady_cols].values.flatten().astype(float)
    vals = vals[~np.isnan(vals)]
    
    if len(vals) < 3:
        return 0.0
    
    x = np.arange(len(vals))
    slope, intercept = np.polyfit(x, vals, 1)
    mean_val = vals.mean()
    
    weekly_trend_pct = (slope / mean_val) * 100 if mean_val > 0 else 0
    return round(weekly_trend_pct, 2)


def compute_sr_overhead(volume_file, partner_name):
    """
    Compute Expected SR Overhead Factor in hours per shift.
    """
    if volume_file is None:
        return 0.0
    
    try:
        vol_xls = pd.ExcelFile(volume_file)
    except Exception:
        return 0.0
    
    target_sheet = find_partner_sheet(vol_xls.sheet_names, partner_name)
    if target_sheet is None:
        return 0.0
    
    data = pd.read_excel(vol_xls, sheet_name=target_sheet)
    cw_cols = [c for c in data.columns if c.startswith('CW')]
    sr_row = data[data['Request Type'] == 'Service Request']
    
    if len(sr_row) == 0:
        return 0.0
    
    steady_cols = [c for c in cw_cols if c != 'CW01']
    sr_avg_week = sr_row[steady_cols].mean(axis=1).values[0]
    
    # Hours per shift = SR_count × 15 min / 60 / 21 shifts per week
    overhead_hrs = sr_avg_week * SR_COORDINATION_MINUTES / 60 / 21
    return round(overhead_hrs, 1)


def compute_sr_overhead_weekly(volume_file, partner_name):
    """Compute SR overhead in hours per WEEK for display."""
    if volume_file is None:
        return 0.0
    
    try:
        vol_xls = pd.ExcelFile(volume_file)
    except Exception:
        return 0.0
    
    target_sheet = find_partner_sheet(vol_xls.sheet_names, partner_name)
    if target_sheet is None:
        return 0.0
    
    data = pd.read_excel(vol_xls, sheet_name=target_sheet)
    cw_cols = [c for c in data.columns if c.startswith('CW')]
    sr_row = data[data['Request Type'] == 'Service Request']
    
    if len(sr_row) == 0:
        return 0.0
    
    steady_cols = [c for c in cw_cols if c != 'CW01']
    sr_avg_week = sr_row[steady_cols].mean(axis=1).values[0]
    
    # Total overhead hours per week
    overhead_hrs_wk = sr_avg_week * SR_COORDINATION_MINUTES / 60
    return round(overhead_hrs_wk, 1)


def find_partner_sheet(sheet_names, partner_name):
    """Find matching sheet name for a partner."""
    partner_lower = str(partner_name).lower().strip()
    
    # Direct match
    for sheet in sheet_names:
        if sheet.lower() == partner_lower:
            return sheet
    
    # Partial match
    for sheet in sheet_names:
        if partner_lower in sheet.lower() or sheet.lower() in partner_lower:
            return sheet
    
    # Keyword match
    partner_keywords = {
        'hcl': 'HCL', 'ntt': 'NTT', 'tcs': 'TCS', 'accenture': 'Accenture',
        'kyndryl': 'Kyndryl', 'tech mahindra': 'Tech Mahindra',
        'infosys': 'Infosys', 'gio': 'GIO', 'lenovo': 'Lenovo',
        'tata': 'TCS', 'global internal': 'GIO',
    }
    for kw, sname in partner_keywords.items():
        if kw in partner_lower:
            if sname in sheet_names:
                return sname
    
    return None


# =============================================================================
# DEMAND FORECASTING
# =============================================================================

class DemandForecaster:
    """Forecasts demand based on day-of-week + shift + week-of-month patterns."""
    
    def __init__(self, data):
        self.data = data
        self.profiles = {}
        self.wom_factors = {}
        self.trend_factor = 1.0
        self._fit()
    
    def _fit(self):
        for shift in ['S1', 'S2', 'S3']:
            for dow in range(7):
                subset = self.data[
                    (self.data['Shift'] == shift) & (self.data['DayOfWeekNum'] == dow)
                ]
                if len(subset) > 0:
                    self.profiles[(dow, shift)] = {
                        'mean': subset['Ʃ Total Demand Hour(s)'].mean(),
                        'std': subset['Ʃ Total Demand Hour(s)'].std(),
                    }
        
        overall_mean = self.data['Ʃ Total Demand Hour(s)'].mean()
        for wom in range(1, 6):
            subset = self.data[self.data['WeekOfMonth'] == wom]
            if len(subset) > 0 and overall_mean > 0:
                self.wom_factors[wom] = subset['Ʃ Total Demand Hour(s)'].mean() / overall_mean
            else:
                self.wom_factors[wom] = 1.0
        
        months = sorted(self.data['YearMonth'].unique())
        if len(months) >= 4:
            recent = self.data[self.data['YearMonth'].isin(months[-3:])]
            earlier = self.data[self.data['YearMonth'].isin(months[:3])]
            if len(earlier) > 0 and earlier['Ʃ Total Demand Hour(s)'].mean() > 0:
                self.trend_factor = (
                    recent['Ʃ Total Demand Hour(s)'].mean() /
                    earlier['Ʃ Total Demand Hour(s)'].mean()
                )
            self.trend_factor = min(max(self.trend_factor, 0.8), 1.3)
    
    def forecast(self, date, shift):
        dow = date.weekday()
        wom = min(((date.day - 1) // 7) + 1, 5)
        
        profile = self.profiles.get((dow, shift))
        if profile is None:
            return {'mean': 50, 'std': 15}
        
        wom_adj = self.wom_factors.get(wom, 1.0)
        trend_adj = min(self.trend_factor, TREND_CAP)
        
        return {
            'mean': profile['mean'] * wom_adj * trend_adj,
            'std': profile['std'],
        }


# =============================================================================
# CAPACITY PLANNING
# =============================================================================

def compute_capacity_plan(forecaster, start_date, track='Basis', days=90,
                          volume_trend_pct=0.0, sr_overhead_hrs=0.0):
    """Generate capacity plan with all enhanced columns."""
    target_util = TARGET_UTIL
    plan = []
    
    for d in range(days):
        date = start_date + timedelta(days=d)
        for shift in ['S1', 'S2', 'S3']:
            fc = forecaster.forecast(date, shift)
            sh = SHIFT_HOURS[shift]
            
            required_cap = fc['mean'] / target_util
            recommended_exec = int(np.ceil(required_cap / sh))
            recommended_exec = max(recommended_exec, 1)
            
            total_cap = recommended_exec * sh
            util_mean = (fc['mean'] / total_cap) * 100 if total_cap > 0 else 0
            buffer = total_cap - fc['mean']
            
            # Effective Workload: actual effort = util × effort_ratio
            effective_workload = util_mean * EFFORT_RATIO_NON_MMI
            
            plan.append({
                'Date': date,
                'Day': date.strftime('%A'),
                'DayOfWeekNum': date.weekday(),
                'Shift': shift,
                'Shift_Timing': SHIFT_TIMING[shift],
                'Recommended_Executors': recommended_exec,
                'Total_Capacity_Hrs': total_cap,
                'Forecast_Hours': round(fc['mean'], 1),
                'Expected_Util_%': round(util_mean, 1),
                'Effective_Workload_%': round(effective_workload, 1),
                'Buffer_Hours': round(buffer, 1),
                'Expected_Volume_Trend': volume_trend_pct,
                'Expected_SR_Overhead_Hrs': sr_overhead_hrs,
            })
    
    return pd.DataFrame(plan)


# =============================================================================
# HTML DASHBOARD
# =============================================================================

def generate_html_dashboard(data, plan_df, forecaster, team_id, team_name,
                            partner_name, track, volume_trend, sr_overhead_hrs,
                            sr_overhead_weekly):
    """Generate interactive HTML dashboard."""
    
    target_util = TARGET_UTIL * 100
    dow_names = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    
    colors = {
        'primary': '#0070F2', 'secondary': '#4CB1FF', 'success': '#36A41D',
        'warning': '#E76500', 'danger': '#BB0000',
        'S1': '#0070F2', 'S2': '#E76500', 'S3': '#36A41D',
    }
    
    # Chart 1: Demand by Day
    fig_dow = go.Figure()
    for shift in ['S1', 'S2', 'S3']:
        s_data = data[data['Shift'] == shift]
        means = s_data.groupby('DayOfWeekNum')['Ʃ Total Demand Hour(s)'].mean()
        fig_dow.add_trace(go.Bar(name=shift, x=dow_names,
            y=[means.get(i, 0) for i in range(7)], marker_color=colors[shift]))
    
    avg_cap = data.groupby('DayOfWeekNum')['Ʃ Total Capacity Hour(s)'].mean()
    threshold_line = [avg_cap.get(i, 0) * 0.70 for i in range(7)]
    fig_dow.add_trace(go.Scatter(name='70% Threshold', x=dow_names, y=threshold_line,
        mode='lines+markers', line=dict(color=colors['danger'], dash='dash', width=2)))
    fig_dow.update_layout(title=dict(text='Historical Demand by Day & Shift', x=0.5),
        barmode='group', yaxis_title='Demand Hours', height=380,
        margin=dict(t=60, b=40, l=60, r=20), legend=dict(orientation='h', y=-0.15))
    
    # Chart 2: Resource Gap
    gap_data = []
    for dow in range(7):
        for shift in ['S1', 'S2', 'S3']:
            current = data[(data['DayOfWeekNum'] == dow) & (data['Shift'] == shift)]['Executors'].mean()
            recommended = plan_df[(plan_df['DayOfWeekNum'] == dow) & (plan_df['Shift'] == shift)]['Recommended_Executors'].mean()
            gap_data.append({'Day': dow_names[dow][:3], 'Shift': shift, 'Gap': recommended - current})
    
    gap_df = pd.DataFrame(gap_data)
    fig_gap = go.Figure()
    for shift in ['S1', 'S2', 'S3']:
        s_gap = gap_df[gap_df['Shift'] == shift]
        fig_gap.add_trace(go.Bar(name=shift, x=s_gap['Day'], y=s_gap['Gap'],
            marker_color=colors[shift],
            text=[f'+{v:.0f}' if v > 0 else f'{v:.0f}' for v in s_gap['Gap']], textposition='outside'))
    fig_gap.update_layout(title=dict(text='Resource GAP: Additional Executors Needed', x=0.5),
        barmode='group', yaxis_title='Additional Executors', height=380,
        margin=dict(t=60, b=40, l=60, r=20), legend=dict(orientation='h', y=-0.15))
    fig_gap.add_hline(y=0, line_color='black', line_width=1)
    
    # KPIs
    total_shifts = len(data)
    pct_exceeded = len(data[data['Total Average Utilization Status'] == 'Exceeded']) / total_shifts * 100
    pct_threshold = len(data[data['Total Average Utilization Status'] == 'Threshold Exceeded']) / total_shifts * 100
    pct_sufficient = len(data[data['Total Average Utilization Status'] == 'Sufficient']) / total_shifts * 100
    avg_util = data['Utilization_Pct'].mean()
    avg_demand = data['Ʃ Total Demand Hour(s)'].mean()
    avg_cap_val = data['Ʃ Total Capacity Hour(s)'].mean()
    cap_demand_ratio = avg_cap_val / avg_demand if avg_demand > 0 else 0
    
    sim_ok = len(plan_df[plan_df['Expected_Util_%'] <= target_util])
    sim_pct_ok = sim_ok / len(plan_df) * 100
    
    # Weekly template table
    weekly_template = plan_df.groupby(['DayOfWeekNum', 'Day', 'Shift']).agg({
        'Recommended_Executors': 'mean',
        'Forecast_Hours': 'mean',
        'Expected_Util_%': 'mean',
        'Effective_Workload_%': 'mean',
    }).reset_index().round(0).sort_values(['DayOfWeekNum', 'Shift'])
    
    table_rows = ""
    for _, row in weekly_template.iterrows():
        shift_time = SHIFT_TIMING[row['Shift']]
        current_avg = data[(data['DayOfWeekNum'] == row['DayOfWeekNum']) & (data['Shift'] == row['Shift'])]['Executors'].mean()
        gap = row['Recommended_Executors'] - current_avg
        gap_color = colors['danger'] if gap > 3 else (colors['warning'] if gap > 1 else colors['success'])
        table_rows += f"""
        <tr>
            <td>{row['Day']}</td>
            <td>{row['Shift']} ({shift_time})</td>
            <td>{current_avg:.1f}</td>
            <td><strong>{int(row['Recommended_Executors'])}</strong></td>
            <td style="color:{gap_color}; font-weight:bold;">{'+' if gap > 0 else ''}{gap:.1f}</td>
            <td>{int(row['Forecast_Hours'])}h</td>
            <td>{int(row['Expected_Util_%'])}%</td>
            <td>{int(row['Effective_Workload_%'])}%</td>
        </tr>"""
    
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PCO TechOps - Shift Capacity Report | {partner_name} - {track}</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: '72', 'Segoe UI', Arial, sans-serif; background: #F5F6F7; color: #000; padding: 20px; }}
        .header {{ background: linear-gradient(135deg, #0070F2, #0054B4); color: white; padding: 30px 40px; border-radius: 12px; margin-bottom: 24px; }}
        .header h1 {{ font-size: 24px; margin-bottom: 8px; }}
        .header .subtitle {{ font-size: 14px; opacity: 0.9; }}
        .header .meta {{ display: flex; gap: 30px; margin-top: 12px; font-size: 13px; opacity: 0.85; flex-wrap: wrap; }}
        .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 16px; margin-bottom: 24px; }}
        .kpi-card {{ background: white; border-radius: 10px; padding: 18px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-top: 3px solid #0070F2; }}
        .kpi-card.danger {{ border-top-color: #BB0000; }}
        .kpi-card.warning {{ border-top-color: #E76500; }}
        .kpi-card.success {{ border-top-color: #36A41D; }}
        .kpi-card .value {{ font-size: 26px; font-weight: bold; margin: 6px 0; }}
        .kpi-card .label {{ font-size: 11px; color: #666; text-transform: uppercase; }}
        .chart-grid {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin-bottom: 24px; }}
        .chart-card {{ background: white; border-radius: 10px; padding: 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
        .section-title {{ font-size: 18px; font-weight: bold; margin: 30px 0 16px 0; padding-bottom: 8px; border-bottom: 2px solid #0070F2; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
        th {{ background: #0070F2; color: white; padding: 10px 12px; text-align: center; }}
        td {{ padding: 8px 12px; text-align: center; border-bottom: 1px solid #E1E2E6; }}
        tr:nth-child(even) {{ background: #F8F9FA; }}
        tr:hover {{ background: #E1F4FF; }}
        .insight-box {{ background: #E1F4FF; border-left: 4px solid #0070F2; padding: 16px 20px; border-radius: 0 8px 8px 0; margin: 16px 0; font-size: 14px; }}
        .insight-box.alert {{ background: #FFF3E0; border-left-color: #E76500; }}
        .insight-box.info {{ background: #F0F7FF; border-left-color: #4CB1FF; }}
        .footer {{ text-align: center; margin-top: 40px; padding: 20px; font-size: 12px; color: #666; }}
        @media (max-width: 900px) {{ .chart-grid {{ grid-template-columns: 1fr; }} .kpi-grid {{ grid-template-columns: repeat(2, 1fr); }} }}
    </style>
</head>
<body>
<div class="header">
    <h1>PCO TechOps - Shift Capacity Planning Report</h1>
    <div class="subtitle">{team_name} | Track: {track} | Team ID: {team_id}</div>
    <div class="meta">
        <span>Analysis: {data['Date'].min().strftime('%b %Y')} - {data['Date'].max().strftime('%b %Y')}</span>
        <span>Shifts Analyzed: {total_shifts}</span>
        <span>Target Util: {target_util:.0f}%</span>
        <span>Report: {datetime.now().strftime('%d %b %Y')}</span>
    </div>
</div>

<div class="kpi-grid">
    <div class="kpi-card danger"><div class="label">Exceeded</div><div class="value">{pct_exceeded:.1f}%</div></div>
    <div class="kpi-card warning"><div class="label">Threshold Exceeded</div><div class="value">{pct_threshold:.1f}%</div></div>
    <div class="kpi-card success"><div class="label">Sufficient</div><div class="value">{pct_sufficient:.1f}%</div></div>
    <div class="kpi-card"><div class="label">Avg Utilization</div><div class="value">{avg_util:.0f}%</div></div>
    <div class="kpi-card"><div class="label">Cap/Demand</div><div class="value">{cap_demand_ratio:.2f}x</div></div>
    <div class="kpi-card success"><div class="label">Proj. Availability</div><div class="value">{sim_pct_ok:.0f}%</div></div>
</div>

<div class="insight-box alert">
    <strong>Key Finding:</strong> {pct_exceeded + pct_threshold:.0f}% of shifts are currently problematic.
    With recommended allocation at {target_util:.0f}% target, projected slot availability improves to {sim_pct_ok:.0f}%.
</div>

<div class="insight-box info">
    <strong>MMI Note:</strong> Service Definitions marked with Minimal Manual Intervention (MMI) do NOT consume shift capacity.
    They receive slots even when capacity is exceeded. The demand figures here reflect only non-MMI workload.
</div>

<div class="section-title">Demand Patterns & Resource Gap</div>
<div class="chart-grid">
    <div class="chart-card"><div id="chart_dow"></div></div>
    <div class="chart-card"><div id="chart_gap"></div></div>
</div>

<div class="section-title">Weekly Executor Planning Template</div>
<div class="insight-box">
    Target utilization for <strong>{track}</strong> track: <strong>{target_util:.0f}%</strong>.
    Effective Workload shows actual manual effort (37.5% of slot-blocking utilization).
    Expected SR Overhead: <strong>{sr_overhead_hrs:.1f} hrs/shift</strong> ({sr_overhead_weekly:.0f} hrs/week) for internal coordination.
</div>
<div class="chart-card" style="margin-top:16px; overflow-x:auto;">
    <table>
        <thead><tr><th>Day</th><th>Shift</th><th>Current</th><th>Recommended</th><th>Gap</th><th>Forecast</th><th>Util%</th><th>Eff. Workload%</th></tr></thead>
        <tbody>{table_rows}</tbody>
    </table>
</div>

<div class="section-title">Model Parameters</div>
<div class="chart-card" style="margin-top:8px; padding:20px;">
    <table style="max-width:600px;">
        <tr><td style="text-align:left;"><strong>Track</strong></td><td>{track}</td></tr>
        <tr><td style="text-align:left;"><strong>Target Utilization</strong></td><td>{target_util:.0f}%</td></tr>
        <tr><td style="text-align:left;"><strong>Effort Ratio (non-MMI)</strong></td><td>37.5%</td></tr>
        <tr><td style="text-align:left;"><strong>Expected Volume Trend</strong></td><td>{volume_trend:+.2f}% per week</td></tr>
        <tr><td style="text-align:left;"><strong>Expected SR Overhead</strong></td><td>{sr_overhead_hrs:.1f} hrs/shift ({sr_overhead_weekly:.0f} hrs/week)</td></tr>
        <tr><td style="text-align:left;"><strong>Demand Trend Factor</strong></td><td>{forecaster.trend_factor:.3f}</td></tr>
        <tr><td style="text-align:left;"><strong>MMI Impact</strong></td><td>Excluded (MMI SDs get slots regardless of capacity)</td></tr>
    </table>
</div>

<div class="footer">
    PCO TechOps Shift Capacity Planning Tool v5 | Target: 65% Utilization | MMI SDs excluded from capacity consumption
</div>

<script>
    var chart_dow = {fig_dow.to_json()};
    var chart_gap = {fig_gap.to_json()};
    Plotly.newPlot('chart_dow', chart_dow.data, chart_dow.layout, {{responsive: true}});
    Plotly.newPlot('chart_gap', chart_gap.data, chart_gap.layout, {{responsive: true}});
</script>
</body>
</html>"""
    
    return html_content


# =============================================================================
# EXCEL OUTPUT
# =============================================================================

def generate_excel_output(plan_df):
    """Generate Excel capacity plan with all columns."""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        plan_export = plan_df[[
            'Date', 'Day', 'Shift',
            'Recommended_Executors', 'Total_Capacity_Hrs',
            'Forecast_Hours', 'Expected_Util_%', 'Effective_Workload_%',
            'Buffer_Hours', 'Expected_Volume_Trend', 'Expected_SR_Overhead_Hrs'
        ]].copy()
        plan_export.to_excel(writer, sheet_name='90Day_Capacity_Plan', index=False)
    
    return output.getvalue()


# =============================================================================
# STREAMLIT UI
# =============================================================================

def main():
    # Custom CSS
    st.markdown("""
    <style>
        .main-header { font-size: 28px; font-weight: bold; color: #0070F2; margin-bottom: 4px; }
        .sub-header { font-size: 14px; color: #666; margin-bottom: 20px; }
        .track-badge { display: inline-block; padding: 4px 12px; border-radius: 4px; font-size: 12px; font-weight: bold; margin-right: 8px; }
        .track-basis { background: #E1F4FF; color: #0070F2; }
        .track-sm { background: #FFF3E0; color: #E76500; }
        .track-db { background: #E8F5E9; color: #36A41D; }
        .info-box { background: #F0F7FF; border-left: 3px solid #4CB1FF; padding: 12px 16px; border-radius: 0 6px 6px 0; margin: 8px 0; font-size: 13px; }
        .stDownloadButton button { background-color: #0070F2 !important; color: white !important; }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="main-header">PCO TechOps - Shift Capacity Planning</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Upload SPC export to generate demand forecast and capacity plan with effort-based workload analysis</div>', unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.image("https://www.sap.com/dam/application/shared/logos/sap-logo-svg.svg", width=100)
        st.markdown("### Configuration")
        
        st.markdown("**Target:** 65% Expected Utilization")
        
        st.markdown("---")
        st.markdown("**Key Concepts:**")
        st.markdown("""
        - **Expected_Util_%**: Slot-blocking utilization in SAP4ME
        - **Effective_Workload_%**: Actual manual effort (37.5% of util)
        - **MMI SDs**: Do NOT consume capacity; get slots even when exceeded
        - **SR Overhead**: Internal coordination time (not in SAP4ME)
        """)
        
        st.markdown("---")
        st.markdown("**Shift Schedule:**")
        st.markdown("""
        - S1: 6AM - 2PM IST (8h)
        - S2: 2PM - 11PM IST (9h)
        - S3: 11PM - 6AM IST (7h)
        """)
        
        st.markdown("---")
        forecast_days = st.slider("Forecast Days", 30, 120, 90, 15)
        
        st.markdown("---")
        st.markdown("**v5** | 65% Target | MMI-Aware")
    
    # File uploads
    col_upload1, col_upload2 = st.columns(2)
    
    with col_upload1:
        uploaded_spc = st.file_uploader(
            "Upload SPC Excel Export (Required)",
            type=['xlsx', 'xls'],
            help="Capacity Demand Overview export from SPC (6-7 months data)"
        )
    
    with col_upload2:
        uploaded_volumes = st.file_uploader(
            "Upload Partners Incoming Volumes (Optional)",
            type=['xlsx', 'xls'],
            help="Partners_Incoming_volumes.xlsx for volume trend and SR overhead calculation"
        )
    
    if uploaded_spc is not None:
        with st.spinner("Analyzing file structure..."):
            tracks_found = load_excel_data(uploaded_spc)
        
        if not tracks_found:
            st.error("Could not find valid shift data. Ensure the file has: Team ID, Date, Shift, Capacity Hours, Demand Hours.")
            return
        
        st.success(f"Detected **{len(tracks_found)} track(s)** in the uploaded file")
        
        # Track selection
        track_options = {}
        for key, info in tracks_found.items():
            label = f"{info['partner']} - {info['track']} ({info['rows']} shifts, {info['team_id']})"
            track_options[label] = key
        
        if len(track_options) == 1:
            selected_label = list(track_options.keys())[0]
            st.info(f"Processing: **{selected_label}**")
        else:
            selected_label = st.selectbox(
                "Select track to analyze:",
                options=list(track_options.keys())
            )
        
        selected_key = track_options[selected_label]
        track_info = tracks_found[selected_key]
        track = track_info['track']
        partner_name = track_info['partner']
        
        # Show target for selected track
        target = TARGET_UTIL
        track_colors = {'Basis': 'track-basis', 'SM': 'track-sm', 'DB': 'track-db'}
        st.markdown(f"""
        <div class="info-box">
            <span class="track-badge {track_colors.get(track, 'track-basis')}">{track}</span>
            Target utilization: <strong>{target*100:.0f}%</strong> | 
            Partner: <strong>{partner_name}</strong> |
            MMI SDs: excluded from capacity (get slots regardless)
        </div>
        """, unsafe_allow_html=True)
        
        # Process button
        if st.button("Generate Forecast & Capacity Plan", type="primary", use_container_width=True):
            with st.spinner("Building demand model and generating forecast..."):
                # Prepare data
                data = prepare_data(track_info['data'])
                
                if len(data) < 30:
                    st.error(f"Insufficient data: only {len(data)} valid shifts. Need at least 30.")
                    return
                
                # Compute volume trend and SR overhead
                volume_trend = compute_volume_trend(uploaded_volumes, partner_name)
                sr_overhead = compute_sr_overhead(uploaded_volumes, partner_name)
                sr_overhead_weekly = compute_sr_overhead_weekly(uploaded_volumes, partner_name)
                
                # Build forecast
                forecaster = DemandForecaster(data)
                
                # Generate plan
                start_date = data['Date'].max().normalize() + timedelta(days=1)
                plan_df = compute_capacity_plan(
                    forecaster, start_date.to_pydatetime(),
                    track=track, days=forecast_days,
                    volume_trend_pct=volume_trend,
                    sr_overhead_hrs=sr_overhead,
                )
                
                # Store in session state
                st.session_state['data'] = data
                st.session_state['plan_df'] = plan_df
                st.session_state['forecaster'] = forecaster
                st.session_state['track_info'] = track_info
                st.session_state['volume_trend'] = volume_trend
                st.session_state['sr_overhead'] = sr_overhead
                st.session_state['sr_overhead_weekly'] = sr_overhead_weekly
                st.session_state['processed'] = True
        
        # Display results
        if st.session_state.get('processed', False):
            data = st.session_state['data']
            plan_df = st.session_state['plan_df']
            forecaster = st.session_state['forecaster']
            track_info = st.session_state['track_info']
            volume_trend = st.session_state['volume_trend']
            sr_overhead = st.session_state['sr_overhead']
            sr_overhead_weekly = st.session_state['sr_overhead_weekly']
            
            partner_name = track_info['partner']
            team_id = track_info['team_id']
            team_name = track_info['team_name']
            track = track_info['track']
            target = TARGET_UTIL
            
            st.markdown("---")
            st.markdown(f"### Results: {partner_name} - {track}")
            
            # KPI Metrics Row 1
            total_shifts = len(data)
            pct_exceeded = len(data[data['Total Average Utilization Status'] == 'Exceeded']) / total_shifts * 100
            pct_threshold = len(data[data['Total Average Utilization Status'] == 'Threshold Exceeded']) / total_shifts * 100
            pct_sufficient = len(data[data['Total Average Utilization Status'] == 'Sufficient']) / total_shifts * 100
            avg_util = data['Utilization_Pct'].mean()
            avg_eff_workload = avg_util * EFFORT_RATIO_NON_MMI
            
            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("Exceeded", f"{pct_exceeded:.1f}%")
            col2.metric("Threshold Exc.", f"{pct_threshold:.1f}%")
            col3.metric("Sufficient", f"{pct_sufficient:.1f}%")
            col4.metric("Avg Util (slot)", f"{avg_util:.0f}%")
            col5.metric("Eff. Workload", f"{avg_eff_workload:.0f}%")
            
            # KPI Metrics Row 2
            col6, col7 = st.columns(2)
            col6.metric("Target Util", f"{target*100:.0f}%")
            col7.metric("Trend Factor", f"{forecaster.trend_factor:.3f}")
            
            st.markdown(f"**Period:** {data['Date'].min().strftime('%b %Y')} - {data['Date'].max().strftime('%b %Y')} | "
                       f"**Shifts analyzed:** {total_shifts}")
            
            # MMI info box
            st.info("**MMI Note:** Service Definitions with Minimal Manual Intervention do NOT consume shift capacity. "
                   "They receive scheduling slots even when capacity is exceeded. Demand figures reflect non-MMI workload only.")
            
            # Charts
            col_left, col_right = st.columns(2)
            
            with col_left:
                dow_names = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
                fig_dow = go.Figure()
                for shift in ['S1', 'S2', 'S3']:
                    s_data = data[data['Shift'] == shift]
                    means = s_data.groupby('DayOfWeekNum')['Ʃ Total Demand Hour(s)'].mean()
                    fig_dow.add_trace(go.Bar(name=shift, x=dow_names,
                        y=[means.get(i, 0) for i in range(7)]))
                avg_cap_day = data.groupby('DayOfWeekNum')['Ʃ Total Capacity Hour(s)'].mean()
                fig_dow.add_trace(go.Scatter(name='70% Threshold', x=dow_names,
                    y=[avg_cap_day.get(i, 0) * 0.70 for i in range(7)],
                    mode='lines+markers', line=dict(color='red', dash='dash', width=2)))
                fig_dow.update_layout(title='Demand by Day & Shift', barmode='group',
                    yaxis_title='Hours', height=350, margin=dict(t=40, b=20))
                st.plotly_chart(fig_dow, use_container_width=True)
            
            with col_right:
                gap_data = []
                for dow in range(7):
                    for shift in ['S1', 'S2', 'S3']:
                        current = data[(data['DayOfWeekNum'] == dow) & (data['Shift'] == shift)]['Executors'].mean()
                        recommended = plan_df[(plan_df['DayOfWeekNum'] == dow) & (plan_df['Shift'] == shift)]['Recommended_Executors'].mean()
                        gap_data.append({'Day': dow_names[dow][:3], 'Shift': shift, 'Gap': recommended - current})
                gap_df_chart = pd.DataFrame(gap_data)
                fig_gap = go.Figure()
                for shift in ['S1', 'S2', 'S3']:
                    s_gap = gap_df_chart[gap_df_chart['Shift'] == shift]
                    fig_gap.add_trace(go.Bar(name=shift, x=s_gap['Day'], y=s_gap['Gap']))
                fig_gap.add_hline(y=0, line_color='black', line_width=1)
                fig_gap.update_layout(title='Resource Gap (Recommended - Current)', barmode='group',
                    yaxis_title='Executors', height=350, margin=dict(t=40, b=20))
                st.plotly_chart(fig_gap, use_container_width=True)
            
            # Weekly Template Table
            st.markdown("#### Weekly Executor Planning Template")
            template_data = []
            dow_names = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
            for dow in range(7):
                for shift in ['S1', 'S2', 'S3']:
                    subset = plan_df[(plan_df['DayOfWeekNum'] == dow) & (plan_df['Shift'] == shift)]
                    hist = data[(data['DayOfWeekNum'] == dow) & (data['Shift'] == shift)]
                    current = hist['Executors'].mean() if len(hist) > 0 else 0
                    recommended = int(subset['Recommended_Executors'].mean())
                    gap = recommended - current
                    util = subset['Expected_Util_%'].mean()
                    eff_wl = subset['Effective_Workload_%'].mean()
                    template_data.append({
                        'Day': dow_names[dow],
                        'Shift': shift,
                        'Current': round(current, 1),
                        'Recommended': recommended,
                        'Gap': f"+{gap:.1f}" if gap > 0 else f"{gap:.1f}",
                        'Expected Util%': f"{util:.0f}%",
                        'Eff. Workload%': f"{eff_wl:.0f}%",
                    })
            
            st.dataframe(pd.DataFrame(template_data), use_container_width=True, hide_index=True)
            
            # Model Parameters Summary
            st.markdown("#### Model Parameters")
            param_col1, param_col2 = st.columns(2)
            with param_col1:
                st.markdown(f"""
                | Parameter | Value |
                |-----------|-------|
                | Track | {track} |
                | Target Utilization | {target*100:.0f}% |
                | Effort Ratio (non-MMI) | 37.5% |
                | Demand Trend Factor | {forecaster.trend_factor:.3f} |
                """)
            with param_col2:
                st.markdown(f"""
                | Parameter | Value |
                |-----------|-------|
                | Effective Workload Target | ~{target*100*EFFORT_RATIO_NON_MMI:.0f}% |
                | Expected Volume Trend | {volume_trend:+.2f}%/week |
                | Expected SR Overhead | {sr_overhead:.1f} hrs/shift |
                | MMI SDs | Excluded from capacity |
                """)
            
            # Downloads
            st.markdown("---")
            st.markdown("#### Download Outputs")
            
            dl_col1, dl_col2 = st.columns(2)
            
            with dl_col1:
                html_content = generate_html_dashboard(
                    data, plan_df, forecaster, team_id, team_name,
                    partner_name, track, volume_trend, sr_overhead, sr_overhead_weekly
                )
                safe_name = re.sub(r'[<>:"/\\|?*]', '', partner_name).strip().replace(' ', '_')
                st.download_button(
                    label="Download HTML Dashboard",
                    data=html_content.encode('utf-8'),
                    file_name=f"PCO_TechOps_Report_{safe_name}_{track}.html",
                    mime="text/html",
                    use_container_width=True
                )
            
            with dl_col2:
                excel_bytes = generate_excel_output(plan_df)
                st.download_button(
                    label="Download Excel Capacity Plan",
                    data=excel_bytes,
                    file_name=f"PCO_TechOps_Plan_{safe_name}_{track}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )


if __name__ == '__main__':
    main()
