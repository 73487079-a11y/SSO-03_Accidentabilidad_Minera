import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from io import BytesIO
from datetime import datetime
import warnings
from matplotlib.backends.backend_pdf import PdfPages

warnings.filterwarnings('ignore')

st.set_page_config(
    page_title="SSO-03: Sistema de Accidentabilidad Minera",
    page_icon="⛏️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Paleta de colores
C = {
    'primary': '#5DD62C',
    'primary_dark': '#337418',
    'secondary': '#1F1F1F',
    'text_dark': '#1F1F1F',
    'text_medium': '#333333',
    'text_light': '#555555',
    'white': '#FFFFFF',
    'black': '#000000',
    'background': '#F8F8F8',
    'card_bg': '#FFFFFF',
    'input_bg': '#2C2C2C',
    'input_text': '#FFFFFF',
    'border': '#E0E0E0',
    'success': '#5DD62C',
    'warning': '#FFA500',
    'danger': '#FF4444',
    'light_gray': '#F8F9FA',
    'medium_gray': '#E0E0E0',
    'dark_gray': '#6C757D',
    'file_uploader_bg': '#2C2C2C',
    'file_uploader_text': '#FFFFFF'
}

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Times+New+Roman&display=swap');
    
    * {{
        font-family: 'Times New Roman', Times, serif !important;
    }}
    
    .stApp {{
        background-color: {C['background']};
    }}
    
    .hero-section {{
        background: linear-gradient(135deg, {C['primary_dark']} 0%, {C['primary']} 100%);
        padding: 2.5rem 2rem;
        border-radius: 24px;
        margin-bottom: 2rem;
        text-align: center;
        position: relative;
        overflow: hidden;
        box-shadow: 0 10px 30px rgba(0,0,0,0.15);
        border-bottom: 4px solid #FFD700;
    }}
    
    .hero-section * {{
        color: {C['white']} !important;
    }}
    
    .hero-section::before {{
        content: "⛏️";
        position: absolute;
        font-size: 180px;
        opacity: 0.1;
        left: -30px;
        bottom: -50px;
        transform: rotate(-15deg);
    }}
    
    .hero-section::after {{
        content: "📊";
        position: absolute;
        font-size: 160px;
        opacity: 0.1;
        right: -30px;
        top: -40px;
        transform: rotate(15deg);
    }}
    
    .hero-title {{
        font-size: 2.8rem;
        font-weight: 800;
        margin-bottom: 0.5rem;
        position: relative;
        z-index: 1;
        letter-spacing: -0.02em;
    }}
    
    .hero-subtitle {{
        font-size: 1.1rem;
        margin-bottom: 1rem;
        position: relative;
        z-index: 1;
    }}
    
    .hero-badge {{
        display: inline-block;
        background: rgba(255,255,255,0.2);
        backdrop-filter: blur(10px);
        padding: 0.5rem 1.2rem;
        border-radius: 50px;
        font-size: 0.8rem;
        position: relative;
        z-index: 1;
    }}
    
    .stat-card {{
        background: {C['card_bg']};
        border-radius: 20px;
        padding: 1.2rem;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        transition: all 0.3s ease;
        border-top: 4px solid {C['primary']};
        cursor: pointer;
    }}
    
    .stat-card:hover {{
        transform: translateY(-5px);
        box-shadow: 0 15px 30px rgba(0,0,0,0.12);
    }}
    
    .stat-number {{
        font-size: 2rem;
        font-weight: 800;
        margin: 0;
        color: {C['text_dark']};
    }}
    
    .stat-label {{
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-top: 0.3rem;
        color: {C['dark_gray']};
    }}
    
    .metric-card {{
        background: {C['card_bg']};
        border-radius: 16px;
        padding: 1rem;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        border-left: 4px solid {C['primary']};
        transition: all 0.3s ease;
    }}
    
    .metric-card:hover {{
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(0,0,0,0.1);
    }}
    
    .metric-label {{
        font-size: 0.7rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        font-weight: 600;
        color: {C['dark_gray']};
    }}
    
    .metric-value {{
        font-size: 1.5rem;
        font-weight: 700;
        margin: 0.25rem 0;
        color: {C['text_dark']};
    }}
    
    .metric-unit {{
        font-size: 0.7rem;
        color: {C['dark_gray']};
    }}
    
    .stTabs [data-baseweb="tab-list"] {{
        gap: 0.5rem;
        background: {C['card_bg']};
        padding: 0.5rem;
        border-radius: 16px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        margin-bottom: 1.5rem;
    }}
    
    .stTabs [data-baseweb="tab"] {{
        border-radius: 12px;
        padding: 0.5rem 1.5rem;
        font-weight: 600;
        font-size: 0.85rem;
        background: transparent;
        transition: all 0.3s ease;
        color: {C['text_dark']};
    }}
    
    .stTabs [aria-selected="true"] {{
        background: linear-gradient(135deg, {C['primary_dark']} 0%, {C['primary']} 100%);
        color: {C['white']} !important;
    }}
    
    .stTabs [aria-selected="true"] * {{
        color: {C['white']} !important;
    }}
    
    .stTabs [aria-selected="false"]:hover {{
        background: {C['light_gray']};
    }}
    
    .stButton > button {{
        background: linear-gradient(135deg, {C['primary_dark']} 0%, {C['primary']} 100%);
        border: none;
        border-radius: 12px;
        padding: 0.6rem 1.5rem;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 2px 8px rgba(0,0,0,0.2);
        color: {C['white']} !important;
    }}
    
    .stButton > button * {{
        color: {C['white']} !important;
    }}
    
    .stButton > button:hover {{
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(93,214,44,0.4);
    }}
    
    .stDownloadButton > button {{
        background: {C['card_bg']} !important;
        border: 2px solid {C['primary']} !important;
        color: {C['text_dark']} !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1) !important;
    }}
    
    .stDownloadButton > button:hover {{
        background: {C['light_gray']} !important;
        border-color: {C['primary_dark']} !important;
        color: {C['text_dark']} !important;
    }}
    
    .stDownloadButton > button * {{
        color: {C['text_dark']} !important;
    }}
    
    .section-title {{
        font-size: 1.4rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        display: inline-block;
        border-bottom: 3px solid {C['primary']};
        padding-bottom: 0.3rem;
        color: {C['text_dark']};
    }}
    
    .section-subtitle {{
        font-size: 0.85rem;
        margin-bottom: 1.5rem;
        color: {C['dark_gray']};
    }}
    
    div[data-testid="stNumberInput"] label, 
    div[data-testid="stTextInput"] label,
    div[data-testid="stSelectbox"] label {{
        font-weight: 600;
        font-size: 0.8rem;
        color: {C['text_dark']} !important;
    }}
    
    div[data-testid="stNumberInput"] input,
    div[data-testid="stTextInput"] input {{
        border-radius: 12px;
        border: 1px solid {C['primary']};
        padding: 0.4rem 0.8rem;
        background-color: {C['input_bg']} !important;
        color: {C['input_text']} !important;
        font-weight: 500;
    }}
    
    div[data-testid="stNumberInput"] input:focus,
    div[data-testid="stTextInput"] input:focus {{
        border-color: {C['primary']};
        box-shadow: 0 0 5px {C['primary']};
    }}
    
    .stFileUploader {{
        background-color: {C['input_bg']} !important;
        border-radius: 12px;
        padding: 1rem;
        border: 1px solid {C['primary']};
    }}
    
    .stFileUploader label {{
        color: {C['input_text']} !important;
    }}
    
    .stFileUploader div, .stFileUploader span, .stFileUploader p {{
        color: {C['input_text']} !important;
    }}
    
    .uploadedFileName {{
        color: {C['input_text']} !important;
    }}
    
    .stExpander {{
        border-radius: 16px;
        border: none;
        background: {C['card_bg']};
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }}
    
    .stExpander > div:first-child {{
        background: {C['light_gray']};
        border-radius: 16px;
        font-weight: 600;
        color: {C['text_dark']};
    }}
    
    .stExpander * {{
        color: {C['text_dark']} !important;
    }}
    
    .dataframe {{
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid {C['medium_gray']};
    }}
    
    hr {{
        margin: 1.5rem 0;
        border: none;
        height: 2px;
        background: linear-gradient(90deg, transparent, {C['primary']}, {C['primary_dark']}, transparent);
    }}
    
    .footer {{
        text-align: center;
        padding: 1.5rem;
        font-size: 0.75rem;
        color: {C['dark_gray']};
    }}
    
    .eval-box {{
        border-radius: 12px;
        padding: 0.8rem 1rem;
        margin: 1rem 0;
        border-left: 4px solid;
    }}
    
    .eval-success {{
        background: #E8F5E9;
        border-left-color: {C['success']};
        color: {C['text_dark']};
    }}
    
    .eval-warning {{
        background: #FFF8E1;
        border-left-color: {C['warning']};
        color: {C['text_dark']};
    }}
    
    .eval-error {{
        background: #FFEBEE;
        border-left-color: {C['danger']};
        color: {C['text_dark']};
    }}
    
    .info-note {{
        background: #E3F2FD;
        padding: 0.8rem;
        border-radius: 8px;
        font-size: 0.8rem;
        margin: 0.5rem 0;
        border-left: 4px solid {C['primary']};
        color: {C['text_dark']};
    }}
    
    h1, h2, h3, h4, h5, h6, p, span, div, label, .stMarkdown {{
        color: {C['text_dark']} !important;
    }}
    
    .stMetric label, .stMetric div {{
        color: {C['text_dark']} !important;
    }}
    
    [data-testid="stMetricValue"] {{
        color: {C['text_dark']} !important;
    }}
    
    .stAlert {{
        color: {C['text_dark']} !important;
    }}
    
    .stAlert p {{
        color: {C['text_dark']} !important;
    }}
    
    table, th, td {{
        color: {C['text_dark']} !important;
    }}
    
    th {{
        background-color: {C['primary']} !important;
        color: {C['white']} !important;
    }}
    
    th * {{
        color: {C['white']} !important;
    }}
    
    .stMarkdown strong {{
        color: {C['text_dark']};
    }}
    
    .stCheckbox label {{
        color: {C['text_dark']} !important;
    }}
    
    .interpretacion-box {{
        background-color: #F0F7F0;
        border-left: 4px solid {C['primary']};
        padding: 0.8rem 1rem;
        border-radius: 8px;
        margin-top: 0.5rem;
        margin-bottom: 1rem;
        font-size: 0.85rem;
    }}
</style>
""", unsafe_allow_html=True)

# Configuración de matplotlib
plt.rcParams['font.size'] = 10
plt.rcParams['axes.labelsize'] = 11
plt.rcParams['axes.titlesize'] = 12
plt.rcParams['legend.fontsize'] = 9
plt.rcParams['font.family'] = 'serif'
plt.rcParams['font.serif'] = ['Times New Roman', 'Times', 'DejaVu Serif']
plt.rcParams['figure.facecolor'] = 'white'
plt.rcParams['axes.facecolor'] = '#F8F9FA'
plt.rcParams['axes.grid'] = True
plt.rcParams['grid.alpha'] = 0.3
plt.rcParams['grid.linestyle'] = '--'
plt.rcParams['axes.spines.top'] = False
plt.rcParams['axes.spines.right'] = False

# ==================== CONFIGURACIÓN DE NORMATIVAS ====================

NORMATIVAS = {
    "D.S. 024-2016-EM (Perú - Minería)": {
        "k_default": 1000000,
        "incluye_am_en_frecuencia": False,
        "nombre_frecuencia": "IF",
        "nombre_severidad": "IS",
        "nombre_global": "IA",
        "referencia": "D.S. 024-2016-EM - Reglamento de Seguridad y Salud Ocupacional en Minería (Perú)",
        "penalizacion_referencia": "Anexo N°33",
        "footer": "D.S. 024-2016-EM | Anexo N°33"
    },
    "OSHA (EE.UU. - General Industrial)": {
        "k_default": 200000,
        "incluye_am_en_frecuencia": True,
        "nombre_frecuencia": "TIFAR",
        "nombre_severidad": "TISSAR",
        "nombre_global": "IA",
        "referencia": "OSHA 29 CFR 1904 - Recording and Reporting Occupational Injuries and Illnesses",
        "penalizacion_referencia": "ANSI Z16.1 (referencia)",
        "footer": "OSHA 29 CFR 1904"
    },
    "ANSI Z16.1 (EE.UU. - Severidad)": {
        "k_default": 1000000,
        "incluye_am_en_frecuencia": False,
        "nombre_frecuencia": "TIFAR",
        "nombre_severidad": "TISSAR",
        "nombre_global": "IA",
        "referencia": "ANSI Z16.1 - Method of Recording and Measuring Work Injury Experience",
        "penalizacion_referencia": "ANSI Z16.1",
        "footer": "ANSI Z16.1"
    }
}

# ==================== FUNCIONES DE CÁLCULO ====================

def calc_k_estandarizacion(trabajadores, horas_semana=40, semanas=50):
    return trabajadores * horas_semana * semanas

def calc_if(atp, fat, am, k, hht, normativa):
    config = NORMATIVAS[normativa]
    if config["incluye_am_en_frecuencia"]:
        return ((am + atp + fat) * k) / hht
    else:
        return ((atp + fat) * k) / hht

def calc_dias_cargados(df, fat):
    return df + (fat * 6000)

def calc_is(dias_cargados, k, hht):
    return (dias_cargados * k) / hht

def calc_ia(frecuencia_valor, severidad_valor):
    return (frecuencia_valor * severidad_valor) / 1000

def calc_far(fat, hht):
    return (fat * 100000000) / hht

def calc_dmi(df, atp):
    return df / atp if atp > 0 else 0

def calc_variacion(valor_actual, valor_anterior):
    return ((valor_actual - valor_anterior) / valor_anterior) * 100 if valor_anterior > 0 else 0

def monte_carlo(params, n_simulations=10000):
    np.random.seed(42)
    n = n_simulations
    
    hht_samples = np.random.normal(params['hht'], params['hht'] * 0.05, n)
    hht_samples = np.maximum(hht_samples, 1)
    am_samples = np.random.poisson(max(0.1, params['am']), n)
    atp_samples = np.random.poisson(max(0.1, params['atp']), n)
    fat_samples = np.random.poisson(max(0.1, params['fat']), n)
    df_samples = np.random.normal(params['df'], params['df'] * 0.15, n)
    df_samples = np.maximum(df_samples, 0)
    k_samples = np.ones(n) * params['k_factor']
    
    if params['incluye_am_en_frecuencia']:
        freq_samples = ((am_samples + atp_samples + fat_samples) * k_samples) / hht_samples
    else:
        freq_samples = ((atp_samples + fat_samples) * k_samples) / hht_samples
    
    dias_cargados_samples = df_samples + (fat_samples * 6000)
    severity_samples = (dias_cargados_samples * k_samples) / hht_samples
    ia_samples = (freq_samples * severity_samples) / 1000
    far_samples = (fat_samples * 100000000) / hht_samples
    dmi_samples = np.where(atp_samples > 0, df_samples / atp_samples, 0)
    
    nombre_frec = params['nombre_frecuencia']
    nombre_sev = params['nombre_severidad']
    
    stats_results = {}
    for name, data in [(nombre_frec, freq_samples), (nombre_sev, severity_samples), ('IA', ia_samples), ('FAR', far_samples), ('DMI', dmi_samples)]:
        stats_results[name] = {
            'mean': np.mean(data), 'std': np.std(data),
            'p5': np.percentile(data, 5), 'p50': np.percentile(data, 50),
            'p90': np.percentile(data, 90), 'p95': np.percentile(data, 95),
            'samples': data[:500]
        }
    
    input_matrix = np.column_stack([hht_samples, am_samples, atp_samples, fat_samples, df_samples])
    correlations = {}
    for i, name in enumerate(['HHT', 'AM', 'ATP', 'FAT', 'DF']):
        correlations[name] = np.corrcoef(input_matrix[:, i], freq_samples)[0, 1]
    
    samples = {'hht': hht_samples, 'am': am_samples, 'atp': atp_samples, 'fat': fat_samples, 'df': df_samples}
    
    return stats_results, correlations, samples

def analisis_historico(df, k_factor, forecast_months=12, anomaly_threshold=2.5, mc_forecast=False, mc_samples=10000, normativa="D.S. 024-2016-EM (Perú - Minería)"):
    df = df.copy()
    df['Periodo'] = pd.to_datetime(df['Periodo'])
    df = df.sort_values('Periodo').reset_index(drop=True)
    
    required_cols = ['HHT', 'ATP', 'FAT', 'DF']
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"La columna '{col}' es obligatoria en el archivo Excel")
    
    if 'AM' not in df.columns:
        df['AM'] = 0
    
    config = NORMATIVAS[normativa]
    nombre_frec = config['nombre_frecuencia']
    nombre_sev = config['nombre_severidad']
    incluye_am = config['incluye_am_en_frecuencia']
    
    if incluye_am:
        df[nombre_frec] = ((df['AM'] + df['ATP'] + df['FAT']) * k_factor) / df['HHT']
    else:
        df[nombre_frec] = ((df['ATP'] + df['FAT']) * k_factor) / df['HHT']
    
    df['Dias_Cargados'] = df['DF'] + (df['FAT'] * 6000)
    df[nombre_sev] = (df['Dias_Cargados'] * k_factor) / df['HHT']
    df['IA'] = (df[nombre_frec] * df[nombre_sev]) / 1000
    df['FAR'] = (df['FAT'] * 100000000) / df['HHT']
    df['DMI'] = df.apply(lambda row: row['DF'] / row['ATP'] if row['ATP'] > 0 else 0, axis=1)
    
    n_meses = len(df)
    n_anios = df['Periodo'].dt.year.nunique()
    fecha_inicio = df['Periodo'].iloc[0].strftime('%Y-%m')
    fecha_fin = df['Periodo'].iloc[-1].strftime('%Y-%m')
    
    years = df['Periodo'].dt.year.unique()
    anual = []
    for year in years:
        year_data = df[df['Periodo'].dt.year == year]
        anual.append({
            'Año': year,
            f'{nombre_frec}_Promedio': year_data[nombre_frec].mean(),
            f'{nombre_frec}_Max': year_data[nombre_frec].max(),
            f'{nombre_frec}_Min': year_data[nombre_frec].min(),
            f'{nombre_frec}_Std': year_data[nombre_frec].std(),
            'Acc_Fatales': year_data['FAT'].sum(),
            'Acc_ATP': year_data['ATP'].sum(),
            'Acc_AM': year_data['AM'].sum(),
            'Dias_Cargados_Total': year_data['Dias_Cargados'].sum(),
            'FAR_Promedio': year_data['FAR'].mean()
        })
    
    x = np.arange(len(df))
    slope, intercept, r_value, p_value, std_err = stats.linregress(x, df[nombre_frec])
    r2 = r_value ** 2
    
    forecast_x = np.arange(len(df), len(df) + forecast_months)
    forecast_val = intercept + slope * forecast_x
    
    se_forecast = std_err * np.sqrt(1 + 1/len(df) + (forecast_x - np.mean(x))**2 / np.sum((x - np.mean(x))**2))
    t_val = stats.t.ppf(0.975, len(df) - 2)
    
    forecast_upper = forecast_val + t_val * se_forecast
    forecast_lower = forecast_val - t_val * se_forecast
    
    z_scores = np.abs(stats.zscore(df[nombre_frec]))
    df['Anomalia'] = z_scores > anomaly_threshold
    anomalias = df[df['Anomalia']].copy()
    
    mc_forecast_results = None
    last_date = df['Periodo'].iloc[-1]
    forecast_dates = []
    for i in range(1, forecast_months + 1):
        next_date = last_date + pd.DateOffset(months=i)
        forecast_dates.append(next_date)
    
    if mc_forecast:
        np.random.seed(42)
        mc_freq = np.zeros((mc_samples, forecast_months))
        
        hht_mean = df['HHT'].mean()
        hht_std = df['HHT'].std()
        am_mean = df['AM'].mean()
        atp_mean = df['ATP'].mean()
        fat_mean = df['FAT'].mean()
        df_mean = df['DF'].mean()
        
        for i in range(mc_samples):
            for j in range(forecast_months):
                hht_sim = np.random.normal(hht_mean, hht_std)
                hht_sim = max(hht_sim, 1)
                am_sim = np.random.poisson(max(0.1, am_mean))
                atp_sim = np.random.poisson(max(0.1, atp_mean))
                fat_sim = np.random.poisson(max(0.1, fat_mean))
                df_sim = np.random.normal(df_mean, df_mean * 0.15)
                df_sim = max(df_sim, 0)
                
                if incluye_am:
                    freq_sim = ((am_sim + atp_sim + fat_sim) * k_factor) / hht_sim
                else:
                    freq_sim = ((atp_sim + fat_sim) * k_factor) / hht_sim
                mc_freq[i, j] = freq_sim
        
        mc_forecast_results = {
            'mean': np.mean(mc_freq, axis=0),
            'p5': np.percentile(mc_freq, 5, axis=0),
            'p50': np.percentile(mc_freq, 50, axis=0),
            'p95': np.percentile(mc_freq, 95, axis=0)
        }
    
    forecast_df = pd.DataFrame({
        'Periodo': forecast_dates,
        f'{nombre_frec}_Pronostico': forecast_val,
        'Limite_Inferior_95': forecast_lower,
        'Limite_Superior_95': forecast_upper
    })
    
    if mc_forecast_results:
        forecast_df[f'{nombre_frec}_MC_P50'] = mc_forecast_results['p50']
        forecast_df[f'{nombre_frec}_MC_P5'] = mc_forecast_results['p5']
        forecast_df[f'{nombre_frec}_MC_P95'] = mc_forecast_results['p95']
    
    return {
        'df_historico': df,
        'resumen_anual': pd.DataFrame(anual),
        'modelo_ols': {
            'slope': slope, 'intercept': intercept, 'r2': r2, 'p_value': p_value, 'std_err': std_err,
            'tendencia_texto': 'AL ALZA (deterioro)' if slope > 0 else 'A LA BAJA (mejora)',
            'nombre_frec': nombre_frec
        },
        'anomalias': anomalias,
        'forecast': forecast_df,
        'mc_forecast': mc_forecast_results,
        'n_meses': n_meses,
        'n_anios': n_anios,
        'fecha_inicio': fecha_inicio,
        'fecha_fin': fecha_fin,
        'forecast_months': forecast_months,
        'normativa': normativa,
        'nombre_frec': nombre_frec,
        'nombre_sev': nombre_sev
    }

def generar_plantilla_excel(incluir_columnas_k=False):
    output = BytesIO()
    
    fechas = [f"2024-{i:02d}" for i in range(1, 13)]
    
    datos_basicos = {
        'Periodo': fechas,
        'HHT': [88000, 85600, 88400, 87200, 86400, 87600, 86000, 86800, 88000, 87200, 86400, 88000],
        'AM': [3, 2, 1, 4, 5, 3, 2, 2, 3, 4, 2, 3],
        'ATP': [2, 1, 0, 3, 2, 2, 1, 1, 2, 3, 1, 2],
        'FAT': [0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0],
        'DF': [15, 10, 0, 30, 20, 15, 8, 10, 18, 25, 12, 20]
    }
    
    if incluir_columnas_k:
        datos_basicos['Trabajadores'] = [500, 495, 502, 498, 501, 497, 503, 496, 500, 499, 502, 498]
        datos_basicos['Horas_Semana'] = [40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40]
        datos_basicos['Semanas'] = [50, 50, 50, 50, 50, 50, 50, 50, 50, 50, 50, 50]
        k_calculadas = []
        for i in range(len(fechas)):
            k = datos_basicos['Trabajadores'][i] * datos_basicos['Horas_Semana'][i] * datos_basicos['Semanas'][i]
            k_calculadas.append(k)
        datos_basicos['K_calculada'] = k_calculadas
    
    df = pd.DataFrame(datos_basicos)
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Datos_Mensuales', index=False)
        
        instrucciones = pd.DataFrame({
            'Instrucción': [
                'BIENVENIDO A LA PLANTILLA SSO-03', '',
                'COLUMNAS OBLIGATORIAS:',
                '1. Periodo: Formato AAAA-MM (ejemplo: 2024-01, 2024-02, ...)',
                '2. HHT: Horas Hombre Trabajadas (número entero, >0)',
                '3. AM: Accidentes Médicos (número entero, >=0)',
                '4. ATP: Accidentes con Tiempo Perdido (número entero, >=0)',
                '5. FAT: Accidentes Fatales (número entero, >=0)',
                '6. DF: Días Físicos Perdidos (número entero, >=0)',
                '',
                'NOTAS IMPORTANTES:',
                '- Complete un registro por cada mes',
                '- Los datos de ejemplo son solo referenciales',
                '- Mantenga el formato de fecha AAAA-MM',
                '',
                'PENALIZACIÓN POR FATALIDAD:',
                '- Cada accidente fatal (FAT) suma automáticamente 6,000 días cargados',
                '',
                'NORMATIVAS SOPORTADAS:',
                '- D.S. 024-2016-EM (Perú): K=1,000,000, NO incluye AM en IF',
                '- OSHA (EE.UU.): K=200,000, SÍ incluye AM en TIFAR',
                '- ANSI Z16.1 (EE.UU.): K=1,000,000, NO incluye AM en TIFAR'
            ]
        })
        instrucciones.to_excel(writer, sheet_name='Instrucciones', index=False, header=False)
    
    output.seek(0)
    return output

def export_excel_mensual(params, deterministic, stats, correlations, samples, normativa):
    output = BytesIO()
    wb = openpyxl.Workbook()
    
    header_fill = PatternFill(start_color='337418', end_color='337418', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    config = NORMATIVAS[normativa]
    nombre_frec = config['nombre_frecuencia']
    nombre_sev = config['nombre_severidad']
    
    ws1 = wb.active
    ws1.title = '1_Datos_Entrada'
    ws1.append(['Parámetro', 'Símbolo', 'Valor', 'Unidades', 'Descripción'])
    for cell in ws1[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
    
    entrada_data = [
        ['Normativa Aplicada', '-', normativa, '-', config['referencia']],
        ['Horas Hombre Trabajadas', 'HHT', params['hht'], 'horas', 'Tiempo total de exposición'],
        ['Número de Trabajadores', 'N', params['trabajadores'], 'trabajadores', 'Para cálculo de K'],
        ['Semanas Laborales al año', 's', params['semanas'], 'semanas', 'Para cálculo de K'],
        ['Horas por Semana', 'h', params['horas_semana'], 'horas/semana', 'Para cálculo de K'],
        ['Constante de Estandarización', 'K', params['k_factor'], 'horas', 'Según normativa seleccionada'],
        ['Accidentes Médicos', 'AM', params['am'], 'eventos', f'{"SÍ" if config["incluye_am_en_frecuencia"] else "NO"} se incluye en {nombre_frec}'],
        ['Accidentes Tiempo Perdido', 'ATP', params['atp'], 'eventos', 'Incapacitantes'],
        ['Accidentes Fatales', 'FAT', params['fat'], 'eventos', 'Penalización 6,000 días'],
        ['Días Físicos Perdidos', 'DF', params['df'], 'días', 'Descanso médico'],
        [f'{nombre_frec} Periodo Anterior', f'{nombre_frec}_ant', params['freq_prev'], 'adim', 'Para cálculo de variación'],
        ['Simulaciones Monte Carlo', 'N_sim', params['n_simulations'], '-', 'Iteraciones'],
    ]
    
    for row_data in entrada_data:
        ws1.append(row_data)
        for cell in ws1[ws1.max_row]:
            cell.border = thin_border
    
    ws2 = wb.create_sheet('2_Resultados_Deterministicos')
    ws2.append(['Variable', 'Valor', 'Unidades', 'Interpretación'])
    for cell in ws2[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
    
    freq_val = deterministic['frecuencia']
    if freq_val < 10:
        interp = "Riesgo controlado. Desempeño favorable."
    elif freq_val < 25:
        interp = "Riesgo moderado. Requiere seguimiento."
    elif freq_val < 50:
        interp = "Riesgo severo. Implementar acciones correctivas."
    else:
        interp = "RIESGO CRÍTICO. Intervención inmediata."
    
    ws2.append([f'{nombre_frec} (Índice de Frecuencia)', freq_val, 'adimensional', interp])
    ws2.append([f'{nombre_sev} (Índice de Severidad)', deterministic['severidad'], 'adimensional', 'Días cargados proyectados'])
    ws2.append(['IA (Índice de Accidentabilidad)', deterministic['ia'], 'adimensional', 'Calificación unificada del riesgo'])
    ws2.append(['FAR (Tasa Mortalidad)', deterministic['far'], 'adimensional', 'Referencia ICMM (100M horas)'])
    ws2.append(['DMI (Duración Media Incapacidad)', deterministic['dmi'], 'días/accidente', 'Promedio de días por ATP'])
    ws2.append(['Días Cargados', deterministic['dias_cargados'], 'días', f'Según {config["penalizacion_referencia"]}'])
    ws2.append([f'Variación {nombre_frec}', f"{deterministic['variacion']:+.2f}%", 'porcentaje', 'vs periodo anterior'])
    
    for cell in ws2[ws2.max_row]:
        cell.border = thin_border
    
    ws3 = wb.create_sheet('3_MonteCarlo_Estadisticas')
    ws3.append(['Variable', 'Media', 'Std Dev', 'P5', 'P50', 'P90', 'P95'])
    for cell in ws3[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
    
    for var, s in stats.items():
        ws3.append([var, s['mean'], s['std'], s['p5'], s['p50'], s['p90'], s['p95']])
        for cell in ws3[ws3.max_row]:
            cell.border = thin_border
    
    ws4 = wb.create_sheet('4_MonteCarlo_Muestra')
    ws4.append(['Muestra #', 'HHT', 'AM', 'ATP', 'FAT', 'DF', nombre_frec, nombre_sev, 'IA', 'FAR'])
    for cell in ws4[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
    
    for i in range(min(500, len(samples['hht']))):
        ws4.append([i+1, f"{samples['hht'][i]:.0f}", int(samples['am'][i]), int(samples['atp'][i]),
                   int(samples['fat'][i]), f"{samples['df'][i]:.1f}", f"{stats[nombre_frec]['samples'][i]:.4f}",
                   f"{stats[nombre_sev]['samples'][i]:.4f}", f"{stats['IA']['samples'][i]:.4f}",
                   f"{stats['FAR']['samples'][i]:.4f}"])
        for cell in ws4[ws4.max_row]:
            cell.border = thin_border
    
    ws5 = wb.create_sheet('5_Correlaciones')
    ws5.append(['Variable de Entrada', f'Correlación con {nombre_frec}', 'Sensibilidad'])
    for cell in ws5[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
    
    for var, corr in correlations.items():
        if abs(corr) > 0.5:
            sens = "ALTA"
        elif abs(corr) > 0.2:
            sens = "MEDIA"
        else:
            sens = "BAJA"
        ws5.append([var, f"{corr:.4f}", sens])
        for cell in ws5[ws5.max_row]:
            cell.border = thin_border
    
    for col in ['A', 'B', 'C', 'D', 'E']:
        ws1.column_dimensions[col].width = 30
        ws2.column_dimensions[col].width = 30
        ws3.column_dimensions[col].width = 18
        ws5.column_dimensions[col].width = 30
    
    wb.save(output)
    output.seek(0)
    return output

def export_historico_excel(resultados):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        resultados['df_historico'].to_excel(writer, sheet_name='1_Datos_Historicos', index=False)
        resultados['resumen_anual'].to_excel(writer, sheet_name='2_Resumen_Anual', index=False)
        ols_df = pd.DataFrame([{
            'Pendiente': resultados['modelo_ols']['slope'],
            'Intercept': resultados['modelo_ols']['intercept'],
            'R2': resultados['modelo_ols']['r2'],
            'p_valor': resultados['modelo_ols']['p_value'],
            'Error_Std': resultados['modelo_ols']['std_err'],
            'Tendencia': resultados['modelo_ols']['tendencia_texto']
        }])
        ols_df.to_excel(writer, sheet_name='3_Modelo_OLS', index=False)
        if len(resultados['anomalias']) > 0:
            resultados['anomalias'].to_excel(writer, sheet_name='4_Anomalias', index=False)
        resultados['forecast'].to_excel(writer, sheet_name='5_Pronostico', index=False)
        if resultados['mc_forecast']:
            mc_df = pd.DataFrame({
                'Mes': range(1, resultados['forecast_months'] + 1),
                'P50': resultados['mc_forecast']['p50'],
                'P5': resultados['mc_forecast']['p5'],
                'P95': resultados['mc_forecast']['p95']
            })
            mc_df.to_excel(writer, sheet_name='6_MonteCarlo_Forecast', index=False)
    output.seek(0)
    return output

# ==================== FUNCIONES DE INTERPRETACIÓN ====================

def interpretacion_radar(valores_normalizados, nombre_frec):
    max_valor = max(valores_normalizados)
    if max_valor <= 1:
        return f"Todos los indicadores están dentro del rango de referencia (≤1). El perfil de riesgo es FAVORABLE. El {nombre_frec} y la severidad están controlados."
    elif max_valor <= 1.5:
        return f"Algunos indicadores superan ligeramente la referencia. El valor más alto es {max_valor:.2f}x el benchmark. Se recomienda revisar el indicador con mayor desviación."
    else:
        return f"El indicador con mayor desviación alcanza {max_valor:.2f}x el valor de referencia. RIESGO CRÍTICO. Se requiere intervención inmediata."

def interpretacion_histogramas(stats, nombre_frec):
    mean_val = stats[nombre_frec]['mean']
    std_val = stats[nombre_frec]['std']
    p5_val = stats[nombre_frec]['p5']
    p95_val = stats[nombre_frec]['p95']
    rango = p95_val - p5_val
    cv = std_val / mean_val if mean_val > 0 else 0
    
    texto = f"El valor esperado (media) es {mean_val:.2f}. "
    if cv < 0.15:
        texto += f"Incertidumbre BAJA (CV={cv:.2f}), resultados confiables. "
    elif cv < 0.30:
        texto += f"Incertidumbre MODERADA (CV={cv:.2f}). "
    else:
        texto += f"Incertidumbre ALTA (CV={cv:.2f}), revisar calidad de datos. "
    
    texto += f"El 90% de los escenarios se encuentran entre {p5_val:.2f} y {p95_val:.2f} (rango {rango:.2f})."
    return texto

def interpretacion_tornado(correlations, nombre_frec):
    max_corr = max(correlations.values())
    max_var = max(correlations, key=correlations.get)
    min_corr = min(correlations.values())
    min_var = min(correlations, key=correlations.get)
    
    texto = f"Variable con MAYOR impacto en {nombre_frec}: {max_var} (correlación {max_corr:.3f}). "
    if max_corr > 0.7:
        texto += "Correlación MUY FUERTE. "
    texto += f"Variable con menor impacto: {min_var} (correlación {min_corr:.3f}). "
    
    if max_corr > 0.5:
        texto += f"Para reducir el {nombre_frec}, priorice acciones sobre {max_var}."
    return texto

def interpretacion_pastel(am, atp, fat):
    total = am + atp + fat
    if total == 0:
        return "No se registraron accidentes. Resultado EXCELENTE. Mantener las condiciones actuales."
    
    pct_atp = (atp / total) * 100
    
    if fat > 0:
        return f"Se registró al menos un accidente FATAL. Esto es CRÍTICO. Según Pirámide de Bird, por cada fatalidad debería haber ~10 ATP y ~600 incidentes."
    elif pct_atp > 30:
        return f"{pct_atp:.1f}% de accidentes son con tiempo perdido (ATP). Proporción ALTA. Refuerce controles en tareas de alto riesgo."
    else:
        return f"{100-pct_atp:.1f}% de accidentes son médicos (AM). Favorable, pero no genere complacencia."

def interpretacion_ols_trend(slope, r2, nombre_frec):
    if slope > 0:
        texto = f"Pendiente POSITIVA (β₁={slope:.4f}): {nombre_frec} está AUMENTANDO con el tiempo. "
        if r2 > 0.65:
            texto += f"R²={r2:.3f} indica CONFIANZA ALTA. Se requiere acción inmediata."
        elif r2 > 0.4:
            texto += f"R²={r2:.3f} indica confianza MODERADA."
        else:
            texto += f"R²={r2:.3f} es BAJO, hay otros factores influyentes."
    else:
        texto = f"Pendiente NEGATIVA (β₁={slope:.4f}): {nombre_frec} está DISMINUYENDO con el tiempo. "
        if r2 > 0.65:
            texto += f"R²={r2:.3f} indica mejora consistente y confiable."
        elif r2 > 0.4:
            texto += f"R²={r2:.3f} indica confianza MODERADA en la mejora."
        else:
            texto += f"R²={r2:.3f} es BAJO, hay variabilidad estacional u otros factores."
    return texto

def interpretacion_boxplot_anual(df, nombre_frec):
    df_year = df.copy()
    df_year['Año'] = df_year['Periodo'].dt.year
    años = sorted(df_year['Año'].unique())
    if len(años) < 2:
        return "Se requiere al menos 2 años de datos para analizar tendencia anual."
    
    medias_anuales = []
    for año in años:
        medias_anuales.append(df_year[df_year['Año'] == año][nombre_frec].mean())
    
    if medias_anuales[-1] > medias_anuales[0]:
        return f"La mediana del {nombre_frec} ha AUMENTADO de {medias_anuales[0]:.2f} a {medias_anuales[-1]:.2f}. Indica DETERIORO sostenido."
    else:
        return f"La mediana del {nombre_frec} ha DISMINUIDO de {medias_anuales[0]:.2f} a {medias_anuales[-1]:.2f}. Indica MEJORA sostenida."

def interpretacion_heatmap(df, nombre_frec):
    df_mes = df.copy()
    df_mes['Mes'] = df_mes['Periodo'].dt.month
    meses_promedio = df_mes.groupby('Mes')[nombre_frec].mean()
    mes_max = meses_promedio.idxmax()
    mes_min = meses_promedio.idxmin()
    nombre_meses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
    
    if meses_promedio.max() > meses_promedio.mean() * 1.3:
        return f"Se detecta ESTACIONALIDAD: mes de {nombre_meses[mes_max-1]} tiene el valor más alto ({meses_promedio[mes_max]:.2f}). Refuerce controles en ese mes."
    else:
        return f"No se detecta estacionalidad significativa. Comportamiento consistente todo el año."

def interpretacion_mc_forecast(mc_forecast, nombre_frec):
    if mc_forecast is None:
        return ""
    p50_final = mc_forecast['p50'][-1]
    p50_inicial = mc_forecast['p50'][0]
    
    if p50_final > p50_inicial * 1.1:
        return f"Pronóstico: {nombre_frec} AUMENTARÁ un {((p50_final/p50_inicial)-1)*100:.1f}% en los próximos meses. Tome acciones preventivas."
    elif p50_final < p50_inicial * 0.9:
        return f"Pronóstico: {nombre_frec} DISMINUIRÁ un {((p50_inicial/p50_final)-1)*100:.1f}% en los próximos meses."
    else:
        return f"Pronóstico: {nombre_frec} se mantendrá ESTABLE en los próximos meses."

# ==================== GRÁFICOS ====================

def generar_grafico_radar(deterministic, params, nombre_frec, nombre_sev, normativa):
    benchmarks = {'Frecuencia': 15, 'Severidad': 5000, 'IA': 75, 'FAR': 50, 'DMI': 15}
    
    freq_norm = min(deterministic['frecuencia'] / benchmarks['Frecuencia'], 2)
    sev_norm = min(deterministic['severidad'] / benchmarks['Severidad'], 2)
    ia_norm = min(deterministic['ia'] / benchmarks['IA'], 2)
    far_norm = min(deterministic['far'] / benchmarks['FAR'], 2)
    dmi_norm = min(deterministic['dmi'] / benchmarks['DMI'], 2)
    
    valores_norm = [freq_norm, sev_norm, ia_norm, far_norm, dmi_norm]
    categorias = [f'{nombre_frec}', f'{nombre_sev}', 'IA', 'FAR', 'DMI']
    
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(projection='polar'))
    fig.patch.set_facecolor('white')
    
    angulos = np.linspace(0, 2 * np.pi, len(categorias), endpoint=False).tolist()
    valores_plot = valores_norm + valores_norm[:1]
    angulos_plot = angulos + angulos[:1]
    
    ax.plot(angulos_plot, valores_plot, 'o-', linewidth=2, color='#5DD62C', markersize=8)
    ax.fill(angulos_plot, valores_plot, alpha=0.25, color='#5DD62C')
    ax.plot(angulos_plot, [1] * len(angulos_plot), 'r--', linewidth=1.5, alpha=0.7, label='Referencia')
    
    ax.set_xticks(angulos)
    ax.set_xticklabels(categorias, fontsize=9)
    ax.set_ylim(0, 2)
    ax.set_yticks([0.5, 1, 1.5, 2])
    ax.set_yticklabels(['0.5x', '1x', '1.5x', '2x'], fontsize=8)
    ax.set_title(f'Perfil de Riesgo - {normativa}', fontsize=12, fontweight='bold', pad=20)
    ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.0))
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    return fig, valores_norm

def generar_grafico_histogramas(stats, nombre_frec, nombre_sev):
    fig, axes = plt.subplots(2, 2, figsize=(12, 10))
    fig.patch.set_facecolor('white')
    vars_plot = [nombre_frec, nombre_sev, 'IA', 'FAR']
    for idx, var in enumerate(vars_plot):
        ax = axes[idx // 2, idx % 2]
        ax.set_facecolor('#F8F9FA')
        data = stats[var]['samples']
        ax.hist(data, bins=50, color='#5DD62C', edgecolor='white', alpha=0.7, density=True)
        ax.axvline(stats[var]['mean'], color='#FF4444', linestyle='--', linewidth=2, label=f'Media: {stats[var]["mean"]:.2f}')
        ax.axvline(stats[var]['p5'], color='#FFA500', linestyle=':', linewidth=2, label=f'P5: {stats[var]["p5"]:.2f}')
        ax.axvline(stats[var]['p95'], color='#FFA500', linestyle=':', linewidth=2, label=f'P95: {stats[var]["p95"]:.2f}')
        ax.set_xlabel('Valor', fontsize=10)
        ax.set_ylabel('Densidad', fontsize=10)
        ax.set_title(f'Distribución de {var}', fontsize=11, fontweight='bold')
        ax.legend(fontsize=8, loc='upper right')
        ax.grid(True, alpha=0.3, linestyle='--')
    plt.suptitle('Análisis Monte Carlo - Distribuciones de Probabilidad', fontsize=14, fontweight='bold')
    plt.tight_layout()
    return fig

def generar_grafico_tornado(correlations, nombre_frec):
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FA')
    names = list(correlations.keys())
    values = list(correlations.values())
    colors_t = ['#FF4444' if v > 0 else '#5DD62C' for v in values]
    bars = ax.barh(names, values, color=colors_t, alpha=0.7, edgecolor='white', linewidth=1)
    ax.axvline(0, color='#1F1F1F', linewidth=1.5, linestyle='-')
    ax.set_xlabel('Coeficiente de Correlación de Pearson', fontsize=11, fontweight='bold')
    ax.set_title(f'Diagrama de Tornado - Sensibilidad del {nombre_frec}', fontsize=12, fontweight='bold')
    for bar, val in zip(bars, values):
        ax.text(val + (0.02 if val >= 0 else -0.08), bar.get_y() + bar.get_height()/2, 
               f'{val:.3f}', va='center', fontsize=9, fontweight='bold')
    ax.grid(True, alpha=0.3, axis='x', linestyle='--')
    plt.tight_layout()
    return fig

def generar_grafico_pastel(params):
    fig, ax = plt.subplots(figsize=(8, 8))
    fig.patch.set_facecolor('white')
    sizes = [params['am'], params['atp'], params['fat']]
    labels = ['Accidentes Médicos\n(AM)', 'Accidentes Tiempo Perdido\n(ATP)', 'Accidentes Fatales\n(FAT)']
    colors_p = ['#5DD62C', '#FF4444', '#337418']
    explode = (0, 0.05, 0.1) if sum(sizes) > 0 else (0, 0, 0)
    if sum(sizes) > 0:
        wedges, texts, autotexts = ax.pie(sizes, explode=explode, labels=labels, colors=colors_p,
                                          autopct=lambda pct: f'{pct:.1f}%\n({int(pct/100*sum(sizes))})',
                                          startangle=90, textprops={'fontsize': 11, 'fontweight': 'bold'})
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(10)
    else:
        ax.pie([1], labels=['Sin accidentes'], colors=['#5DD62C'], autopct='100%')
    ax.set_title('Distribución de Eventos por Severidad', fontsize=12, fontweight='bold')
    plt.tight_layout()
    return fig

def generar_grafico_ols_trend(params, deterministic, freq_prev, nombre_frec):
    meses_simulados = np.arange(1, 13)
    freq_simulado = freq_prev * (1 + np.sin(np.linspace(0, 2*np.pi, 12)) * 0.3)
    freq_simulado = np.maximum(freq_simulado, 1)
    freq_simulado[11] = deterministic['frecuencia']
    
    x = np.arange(len(meses_simulados))
    slope, intercept = np.polyfit(x, freq_simulado, 1)
    trend = intercept + slope * x
    fut_meses = np.arange(13, 19)
    fut_trend = intercept + slope * fut_meses
    
    residuos = freq_simulado - trend
    se = np.sqrt(np.sum(residuos**2) / (len(x) - 2))
    r2 = 1 - np.sum(residuos**2) / np.sum((freq_simulado - np.mean(freq_simulado))**2)
    t_val = 1.96
    upper_band = trend + t_val * se
    lower_band = trend - t_val * se
    
    fig, ax = plt.subplots(figsize=(11, 6))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FA')
    ax.plot(meses_simulados, freq_simulado, 'o-', label=f'{nombre_frec} Histórico', color='#5DD62C', linewidth=2, markersize=8, markerfacecolor='white', markeredgewidth=2)
    ax.plot(meses_simulados, trend, '--', label=f'Tendencia OLS (β₁={slope:.3f})', color='#FF4444', linewidth=2)
    ax.fill_between(meses_simulados, lower_band, upper_band, alpha=0.2, color='#FFA500', label='Intervalo 95%')
    ax.plot(fut_meses, fut_trend, '--', color='#337418', linewidth=2, label='Proyección 6 meses')
    
    ax.set_xlabel('Periodo (meses)', fontsize=11, fontweight='bold')
    ax.set_ylabel(f'{nombre_frec}', fontsize=11, fontweight='bold')
    ax.set_title(f'Análisis Predictivo OLS - Tendencia del {nombre_frec}', fontsize=12, fontweight='bold')
    ax.legend(loc='upper left', fontsize=9)
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.set_xlim(0, 19)
    plt.tight_layout()
    return fig, slope, r2

def generar_grafico_boxplot_anual(df, nombre_frec):
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FA')
    
    df['Año'] = df['Periodo'].dt.year
    data_by_year = [df[df['Año'] == year][nombre_frec].values for year in sorted(df['Año'].unique())]
    
    bp = ax.boxplot(data_by_year, labels=sorted(df['Año'].unique()), patch_artist=True)
    for patch in bp['boxes']:
        patch.set_facecolor('#5DD62C')
        patch.set_alpha(0.7)
    for median in bp['medians']:
        median.set_color('#FF4444')
        median.set_linewidth(2)
    
    ax.set_xlabel('Año', fontsize=11, fontweight='bold')
    ax.set_ylabel(f'{nombre_frec}', fontsize=11, fontweight='bold')
    ax.set_title(f'Distribución Anual del {nombre_frec}', fontsize=12, fontweight='bold')
    ax.grid(True, alpha=0.3, axis='y', linestyle='--')
    plt.tight_layout()
    return fig

def generar_grafico_heatmap(df, nombre_frec):
    fig, ax = plt.subplots(figsize=(12, 8))
    fig.patch.set_facecolor('white')
    
    df['Año'] = df['Periodo'].dt.year
    df['Mes'] = df['Periodo'].dt.month
    pivot = df.pivot_table(index='Año', columns='Mes', values=nombre_frec, fill_value=0)
    
    im = ax.imshow(pivot.values, cmap='RdYlGn_r', aspect='auto', vmin=0, vmax=df[nombre_frec].quantile(0.9))
    ax.set_xticks(np.arange(len(pivot.columns)))
    ax.set_yticks(np.arange(len(pivot.index)))
    ax.set_xticklabels(['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'])
    ax.set_yticklabels(pivot.index)
    
    for i in range(len(pivot.index)):
        for j in range(len(pivot.columns)):
            ax.text(j, i, f'{pivot.values[i, j]:.1f}', ha='center', va='center', color='black', fontsize=8, fontweight='bold')
    
    plt.colorbar(im, ax=ax, label=nombre_frec)
    ax.set_xlabel('Mes', fontsize=11, fontweight='bold')
    ax.set_ylabel('Año', fontsize=11, fontweight='bold')
    ax.set_title(f'Heatmap - {nombre_frec} por Mes y Año', fontsize=12, fontweight='bold')
    plt.tight_layout()
    return fig

def generar_grafico_tendencia_estacional(df, ols):
    fig, ax = plt.subplots(figsize=(12, 6))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FA')
    nombre_frec = ols['nombre_frec']
    
    x = np.arange(len(df))
    trend_line = ols['intercept'] + ols['slope'] * x
    
    ax.plot(df['Periodo'], df[nombre_frec], 'o-', color='#5DD62C', alpha=0.6, markersize=4, label='Datos reales')
    ax.plot(df['Periodo'], trend_line, 'r-', linewidth=2.5, label=f'Tendencia (β₁={ols["slope"]:.3f})')
    
    media_movil = df[nombre_frec].rolling(window=12, center=True).mean()
    ax.plot(df['Periodo'], media_movil, '--', color='#FFA500', linewidth=2, label='Media móvil 12m')
    
    ax.set_xlabel('Periodo', fontsize=11, fontweight='bold')
    ax.set_ylabel(f'{nombre_frec}', fontsize=11, fontweight='bold')
    ax.set_title(f'Tendencia y Estacionalidad - R² = {ols["r2"]:.3f}', fontsize=12, fontweight='bold')
    ax.legend(loc='upper left', fontsize=9)
    ax.grid(True, alpha=0.3, linestyle='--')
    plt.xticks(rotation=45)
    plt.tight_layout()
    return fig

def generar_grafico_comparativo_anual(df, nombre_frec):
    fig, ax = plt.subplots(figsize=(12, 6))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FA')
    
    df['Año'] = df['Periodo'].dt.year
    for year in sorted(df['Año'].unique()):
        year_data = df[df['Año'] == year]
        ax.plot(year_data['Periodo'].dt.month, year_data[nombre_frec], 'o-', label=str(year), linewidth=2, markersize=6)
    
    ax.set_xlabel('Mes', fontsize=11, fontweight='bold')
    ax.set_ylabel(f'{nombre_frec}', fontsize=11, fontweight='bold')
    ax.set_title(f'Comparativa Anual del {nombre_frec}', fontsize=12, fontweight='bold')
    ax.set_xticks(range(1, 13))
    ax.set_xticklabels(['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'])
    ax.legend(loc='upper left', fontsize=9)
    ax.grid(True, alpha=0.3, linestyle='--')
    plt.tight_layout()
    return fig

def generar_grafico_mc_forecast(forecast, mc_forecast, nombre_frec):
    fig, ax = plt.subplots(figsize=(12, 6))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FA')
    
    ax.plot(forecast['Periodo'], forecast[f'{nombre_frec}_Pronostico'], '--', color='#FF4444', linewidth=2, label='Pronóstico OLS')
    ax.fill_between(forecast['Periodo'], forecast['Limite_Inferior_95'], forecast['Limite_Superior_95'], alpha=0.2, color='#FFA500', label='Intervalo 95% OLS')
    
    if mc_forecast:
        ax.plot(forecast['Periodo'], mc_forecast['p50'], '-', color='#5DD62C', linewidth=2.5, label='Monte Carlo P50')
        ax.fill_between(forecast['Periodo'], mc_forecast['p5'], mc_forecast['p95'], alpha=0.3, color='#5DD62C', label='MC P5-P95')
    
    ax.set_xlabel('Periodo', fontsize=11, fontweight='bold')
    ax.set_ylabel(f'{nombre_frec}', fontsize=11, fontweight='bold')
    ax.set_title('Pronóstico con Monte Carlo', fontsize=12, fontweight='bold')
    ax.legend(loc='upper left', fontsize=9)
    ax.grid(True, alpha=0.3, linestyle='--')
    plt.xticks(rotation=45)
    plt.tight_layout()
    return fig

# ==================== REPORTES PDF ====================

def generar_reporte_pdf_mensual(params, deterministic, stats_results, correlations, normativa, interpretaciones):
    buffer = BytesIO()
    
    plt.rcParams['font.family'] = 'serif'
    plt.rcParams['font.serif'] = ['Times New Roman']
    plt.rcParams['font.size'] = 12
    
    config = NORMATIVAS[normativa]
    nombre_frec = config['nombre_frecuencia']
    nombre_sev = config['nombre_severidad']
    
    with PdfPages(buffer) as pdf:
        # SECCIÓN 1: CÁLCULOS (VERTICAL)
        fig1 = plt.figure(figsize=(8.27, 11.69))
        fig1.patch.set_facecolor('white')
        ax1 = fig1.add_subplot(111)
        ax1.axis('off')
        
        fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
        freq_val = deterministic['frecuencia']
        
        if freq_val < 10:
            nivel_riesgo = "FAVORABLE"
            recomendacion = "Riesgo controlado. Mantener programas de prevención actuales."
        elif freq_val < 25:
            nivel_riesgo = "MODERADO"
            recomendacion = "Riesgo moderado. Requiere seguimiento especial."
        elif freq_val < 50:
            nivel_riesgo = "SEVERO"
            recomendacion = "Riesgo severo. Implementar plan de acción correctivo inmediato."
        else:
            nivel_riesgo = "CRITICO"
            recomendacion = "Riesgo crítico. Suspender actividades de alto riesgo inmediatamente."

        texto_calculos = f"""
SSO-03: SISTEMA DE CALCULO Y PROYECCION DE TASAS DE ACCIDENTABILIDAD MINERA
INFORME TECNICO - ANALISIS MENSUAL

================================================================================
1. DATOS GENERALES
================================================================================
  Normativa aplicada        : {normativa}
  Fecha de emision          : {fecha}
  Simulaciones Monte Carlo  : {params['n_simulations']:,}

================================================================================
2. DATOS DE ENTRADA
================================================================================

  EXPOSICION:
     • Horas Hombre Trabajadas (HHT)    : {params['hht']:,} horas
     • Numero de Trabajadores (N)       : {params['trabajadores']:,}
     • Semanas laborales por año (s)    : {params['semanas']}
     • Horas por semana (h)             : {params['horas_semana']}
     • Constante de Estandarizacion (K) : {params['k_factor']:,} horas

  ACCIDENTES:
     • AM (Accidentes Medicos)          : {params['am']} eventos {"(NO incluido en IF)" if not config['incluye_am_en_frecuencia'] else "(incluido en TIFAR)"}
     • ATP (Accidentes Tiempo Perdido)  : {params['atp']} eventos
     • FAT (Accidentes Fatales)         : {params['fat']} eventos
     • DF (Dias Fisicos Perdidos)       : {params['df']} dias
     • Dias Cargados Totales            : {deterministic['dias_cargados']:,.0f} dias (DF + FAT×6,000)

================================================================================
3. RESULTADOS DETERMINISTICOS
================================================================================

  INDICADOR                         VALOR        INTERPRETACION
  ------------------------------------------------------------------------------
  {nombre_frec:<30} {freq_val:>12.4f}   {"Riesgo controlado" if freq_val<10 else "Riesgo moderado" if freq_val<25 else "Riesgo severo" if freq_val<50 else "RIESGO CRITICO"}
  {nombre_sev:<30} {deterministic['severidad']:>12.2f}   Dias cargados proyectados por K horas
  IA (Accidentabilidad Global)      {deterministic['ia']:>12.2f}   Calificacion unificada del riesgo
  FAR (Tasa de Mortalidad)          {deterministic['far']:>12.2f}   Fatalidades por 100M horas (ICMM)
  DMI (Duracion Media Incapacidad)  {deterministic['dmi']:>12.2f}   Dias/accidente
  Variacion {nombre_frec:<14} {deterministic['variacion']:>+12.2f}%   vs periodo anterior

================================================================================
4. EVALUACION DE RIESGO
================================================================================

  NIVEL DE RIESGO: {nivel_riesgo}

  RECOMENDACION: {recomendacion}

================================================================================
5. INTERPRETACION DE RESULTADOS NUMERICOS
================================================================================

  {nombre_frec} (Valor: {freq_val:.4f}):
     • {"Valor dentro del rango favorable (<10). La frecuencia de accidentes es baja." if freq_val < 10 else "Valor supera el rango favorable. Se requiere atencion para reducir la frecuencia."}

  {nombre_sev} (Valor: {deterministic['severidad']:.2f}):
     • Representa los dias cargados proyectados por cada {params['k_factor']:,} horas trabajadas.
     • {"Incluye penalizacion de 6,000 dias por accidente fatal." if params['fat'] > 0 else "No se registraron fatalidades."}

  IA (Valor: {deterministic['ia']:.2f}):
     • {"Valor aceptable." if deterministic['ia'] < 100 else "Valor elevado, requiere atencion prioritaria."}

  FAR (Valor: {deterministic['far']:.2f}):
     • Proyeccion a 100 millones de horas.
     • {"Sin fatalidades en este periodo." if params['fat'] == 0 else f"Se proyectan {deterministic['far']:.0f} fatalidades a esa exposicion."}

  DMI (Valor: {deterministic['dmi']:.2f} dias/accidente):
     • {"Duracion media baja, los trabajadores se recuperan rapidamente." if deterministic['dmi'] < 15 else "Duracion media alta, los accidentes son graves."}

  Variacion: {deterministic['variacion']:+.2f}%
     • {"El desempeno de seguridad ha empeorado respecto al periodo anterior." if deterministic['variacion'] > 0 else "El desempeno de seguridad ha mejorado respecto al periodo anterior."}

================================================================================
FIN DE LA SECCION DE CALCULOS
================================================================================
"""
        
        ax1.text(0.05, 0.95, texto_calculos, transform=ax1.transAxes, fontsize=12, verticalalignment='top', fontfamily='serif')
        pdf.savefig(fig1, bbox_inches='tight')
        plt.close(fig1)
        
        # SECCIÓN 2: INTERPRETACIÓN DE GRÁFICOS (VERTICAL)
        fig2 = plt.figure(figsize=(8.27, 11.69))
        fig2.patch.set_facecolor('white')
        ax2 = fig2.add_subplot(111)
        ax2.axis('off')
        
        texto_interpretaciones = f"""
SSO-03: SISTEMA DE CALCULO Y PROYECCION DE TASAS DE ACCIDENTABILIDAD MINERA
INTERPRETACION DE GRAFICOS

================================================================================
1. GRAFICO RADAR - PERFIL DE RIESGO
================================================================================
{interpretaciones['radar']}

-------------------------------------------------------------------------------
2. HISTOGRAMAS MONTE CARLO - DISTRIBUCIONES DE PROBABILIDAD
================================================================================
{interpretaciones['histogramas']}

-------------------------------------------------------------------------------
3. DIAGRAMA DE TORNADO - ANALISIS DE SENSIBILIDAD
================================================================================
{interpretaciones['tornado']}

-------------------------------------------------------------------------------
4. GRAFICO DE PASTEL - DISTRIBUCION DE EVENTOS
================================================================================
{interpretaciones['pastel']}

-------------------------------------------------------------------------------
5. ANALISIS PREDICTIVO OLS - TENDENCIA TEMPORAL
================================================================================
{interpretaciones['ols']}

-------------------------------------------------------------------------------
6. PERCENTILES MONTE CARLO
================================================================================
Los percentiles permiten entender la incertidumbre de los indicadores:

  • P5 (percentil 5):  Escenario optimista (solo 5% de probabilidad de ser menor)
  • P50 (mediana):     Escenario mas probable (50% de probabilidad)
  • P95 (percentil 95): Escenario pesimista (solo 5% de probabilidad de ser mayor)

INTERPRETACION DEL RANGO P5-P95:
  • Si el rango es estrecho  → alta confianza en la estimacion
  • Si el rango es amplio    → alta incertidumbre, revisar calidad de datos

================================================================================
FIN DE LA SECCION DE INTERPRETACIONES
================================================================================
"""
        
        ax2.text(0.05, 0.95, texto_interpretaciones, transform=ax2.transAxes, fontsize=12, verticalalignment='top', fontfamily='serif')
        pdf.savefig(fig2, bbox_inches='tight')
        plt.close(fig2)
        
        # SECCIÓN 3: GRÁFICOS (HORIZONTAL)
        fig_radar, _ = generar_grafico_radar(deterministic, params, nombre_frec, nombre_sev, normativa)
        fig_radar.set_size_inches(11.69, 8.27)
        pdf.savefig(fig_radar, bbox_inches='tight')
        plt.close(fig_radar)
        
        fig_hist = generar_grafico_histogramas(stats_results, nombre_frec, nombre_sev)
        fig_hist.set_size_inches(11.69, 8.27)
        pdf.savefig(fig_hist, bbox_inches='tight')
        plt.close(fig_hist)
        
        fig_tornado = generar_grafico_tornado(correlations, nombre_frec)
        fig_tornado.set_size_inches(11.69, 8.27)
        pdf.savefig(fig_tornado, bbox_inches='tight')
        plt.close(fig_tornado)
        
        fig_pastel = generar_grafico_pastel(params)
        fig_pastel.set_size_inches(11.69, 8.27)
        pdf.savefig(fig_pastel, bbox_inches='tight')
        plt.close(fig_pastel)
        
        fig_trend, _, _ = generar_grafico_ols_trend(params, deterministic, params['freq_prev'], nombre_frec)
        fig_trend.set_size_inches(11.69, 8.27)
        pdf.savefig(fig_trend, bbox_inches='tight')
        plt.close(fig_trend)
    
    buffer.seek(0)
    return buffer

def generar_reporte_pdf_historico(resultados, interpretaciones):
    buffer = BytesIO()
    
    plt.rcParams['font.family'] = 'serif'
    plt.rcParams['font.serif'] = ['Times New Roman']
    plt.rcParams['font.size'] = 12
    
    config = NORMATIVAS[resultados['normativa']]
    nombre_frec = resultados['nombre_frec']
    nombre_sev = resultados['nombre_sev']
    df = resultados['df_historico']
    ols = resultados['modelo_ols']
    
    with PdfPages(buffer) as pdf:
        # SECCIÓN 1: CÁLCULOS (VERTICAL)
        fig1 = plt.figure(figsize=(8.27, 11.69))
        fig1.patch.set_facecolor('white')
        ax1 = fig1.add_subplot(111)
        ax1.axis('off')
        
        fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
        
        texto_calculos = f"""
SSO-03: SISTEMA DE CALCULO Y PROYECCION DE TASAS DE ACCIDENTABILIDAD MINERA
INFORME TECNICO - SERIES TEMPORALES

================================================================================
1. DATOS GENERALES
================================================================================
  Normativa aplicada        : {resultados['normativa']}
  Fecha de emision          : {fecha}
  Periodo analizado         : {resultados['fecha_inicio']} a {resultados['fecha_fin']}
  Meses analizados          : {resultados['n_meses']} meses
  Años distintos            : {resultados['n_anios']}

================================================================================
2. ESTADISTICAS GLOBALES DEL PERIODO
================================================================================

  INDICADOR                                     VALOR
  ------------------------------------------------------------------------------
  {nombre_frec:<40} {df[nombre_frec].mean():>12.4f} (promedio)
  {nombre_frec} Maximo                          {df[nombre_frec].max():>12.4f}
  {nombre_frec} Minimo                          {df[nombre_frec].min():>12.4f}
  {nombre_sev:<40} {df[nombre_sev].mean():>12.2f} (promedio)
  IA (Accidentabilidad Global) Promedio         {df['IA'].mean():>12.2f}
  FAR (Tasa Mortalidad) Promedio                {df['FAR'].mean():>12.2f}

================================================================================
3. TOTALES ACUMULADOS
================================================================================
  • Total Accidentes Fatales (FAT)    : {df['FAT'].sum()}
  • Total Accidentes ATP              : {df['ATP'].sum()}
  • Total Accidentes AM (referencia)  : {df['AM'].sum()}
  • Total Dias Cargados               : {df['Dias_Cargados'].sum():,.0f} dias

================================================================================
4. MODELO DE TENDENCIA OLS
================================================================================

  PARAMETRO                        VALOR         INTERPRETACION
  ------------------------------------------------------------------------------
  Pendiente (β₁)                   {ols['slope']:>14.6f}   {"Tendencia al alza" if ols['slope'] > 0 else "Tendencia a la baja"}
  Coeficiente R²                   {ols['r2']:>14.4f}   {"Confiabilidad ALTA" if ols['r2'] > 0.65 else "Confiabilidad MEDIA" if ols['r2'] > 0.4 else "Confiabilidad BAJA"}
  Error estandar                   {ols['std_err']:>14.4f}   Precision del modelo
  p-valor                          {ols['p_value']:>14.6f}   {"Significativo" if ols['p_value'] < 0.05 else "No significativo"}

================================================================================
5. EVALUACION DE TENDENCIA
================================================================================
  TENDENCIA DETECTADA: {ols['tendencia_texto']}
  
  {interpretacion_ols_trend(ols['slope'], ols['r2'], nombre_frec)}

================================================================================
6. MESES CON ANOMALIAS DETECTADAS
================================================================================
  {"• Se detectaron " + str(len(resultados['anomalias'])) + " meses con comportamiento atipico." if len(resultados['anomalias']) > 0 else "• No se detectaron anomalias en la serie temporal."}

================================================================================
FIN DE LA SECCION DE CALCULOS
================================================================================
"""
        
        ax1.text(0.05, 0.95, texto_calculos, transform=ax1.transAxes, fontsize=12, verticalalignment='top', fontfamily='serif')
        pdf.savefig(fig1, bbox_inches='tight')
        plt.close(fig1)
        
        # SECCIÓN 2: INTERPRETACIÓN DE GRÁFICOS (VERTICAL)
        fig2 = plt.figure(figsize=(8.27, 11.69))
        fig2.patch.set_facecolor('white')
        ax2 = fig2.add_subplot(111)
        ax2.axis('off')
        
        texto_interpretaciones = f"""
SSO-03: SISTEMA DE CALCULO Y PROYECCION DE TASAS DE ACCIDENTABILIDAD MINERA
INTERPRETACION DE GRAFICOS - SERIES TEMPORALES

================================================================================
1. BOXPLOT ANUAL - DISTRIBUCION POR AÑO
================================================================================
{interpretaciones['boxplot']}

-------------------------------------------------------------------------------
2. HEATMAP - ESTACIONALIDAD MENSUAL
================================================================================
{interpretaciones['heatmap']}

-------------------------------------------------------------------------------
3. TENDENCIA Y ESTACIONALIDAD
================================================================================
{interpretaciones['tendencia']}

-------------------------------------------------------------------------------
4. COMPARATIVA ANUAL
================================================================================
{interpretaciones['comparativa']}

-------------------------------------------------------------------------------
5. PRONOSTICO MONTE CARLO
================================================================================
{interpretaciones['forecast']}

================================================================================
FIN DE LA SECCION DE INTERPRETACIONES
================================================================================
"""
        
        ax2.text(0.05, 0.95, texto_interpretaciones, transform=ax2.transAxes, fontsize=12, verticalalignment='top', fontfamily='serif')
        pdf.savefig(fig2, bbox_inches='tight')
        plt.close(fig2)
        
        # SECCIÓN 3: GRÁFICOS (HORIZONTAL)
        fig_box = generar_grafico_boxplot_anual(df, nombre_frec)
        fig_box.set_size_inches(11.69, 8.27)
        pdf.savefig(fig_box, bbox_inches='tight')
        plt.close(fig_box)
        
        fig_heat = generar_grafico_heatmap(df, nombre_frec)
        fig_heat.set_size_inches(11.69, 8.27)
        pdf.savefig(fig_heat, bbox_inches='tight')
        plt.close(fig_heat)
        
        fig_tend = generar_grafico_tendencia_estacional(df, ols)
        fig_tend.set_size_inches(11.69, 8.27)
        pdf.savefig(fig_tend, bbox_inches='tight')
        plt.close(fig_tend)
        
        fig_comp = generar_grafico_comparativo_anual(df, nombre_frec)
        fig_comp.set_size_inches(11.69, 8.27)
        pdf.savefig(fig_comp, bbox_inches='tight')
        plt.close(fig_comp)
        
        if resultados['mc_forecast']:
            fig_mc = generar_grafico_mc_forecast(resultados['forecast'], resultados['mc_forecast'], nombre_frec)
            fig_mc.set_size_inches(11.69, 8.27)
            pdf.savefig(fig_mc, bbox_inches='tight')
            plt.close(fig_mc)
    
    buffer.seek(0)
    return buffer

# ==================== PARÁMETROS ====================

PARAM_INFO = {
    'hht': {'default': 88000},
    'trabajadores': {'default': 500},
    'semanas': {'default': 50},
    'horas_semana': {'default': 40},
    'am': {'default': 3},
    'atp': {'default': 3},
    'fat': {'default': 1},
    'df': {'default': 31},
    'freq_prev': {'default': 6.47}
}

# ==================== MODALES ====================

if 'modal' not in st.session_state:
    st.session_state.modal = None

def close_modal():
    st.session_state.modal = None

def open_modal(modal_name):
    st.session_state.modal = modal_name

# ==================== INTERFAZ PRINCIPAL ====================

st.markdown(f"""
<div class="hero-section">
    <div class="hero-title">SSO-03</div>
    <div class="hero-subtitle">Sistema de Calculo y Proyeccion de Tasas de Accidentabilidad Minera</div>
    <div class="hero-badge">⛏️ D.S. 024-2016-EM | OSHA | ANSI Z16.1</div>
</div>
""", unsafe_allow_html=True)

# Tarjetas clickeables
col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)

with col_stat1:
    if st.button("📊 IF/TIFAR\nIndice Frecuencia", key="btn_if", use_container_width=True):
        open_modal("if")

with col_stat2:
    if st.button("📊 IS/TISSAR\nIndice Severidad", key="btn_is", use_container_width=True):
        open_modal("is")

with col_stat3:
    if st.button("📊 IA\nAccidentabilidad Global", key="btn_ia", use_container_width=True):
        open_modal("ia")

with col_stat4:
    if st.button("📊 FAR\nTasa Mortalidad", key="btn_far", use_container_width=True):
        open_modal("far")

# Modales mejorados
if st.session_state.modal == "if":
    with st.expander("📊 IF/TIFAR - Índice de Frecuencia", expanded=True):
        st.markdown("""
        ### 📐 Fórmula General
        
        $$
        \\text{IF} = \\frac{\\text{ATP} + \\text{FAT} + (\\text{AM} \\times w)}{\\text{HHT}} \\times K
        $$
        
        ### 📋 Según Normativa
        
        | Normativa | Fórmula | K | Incluye AM |
        |-----------|---------|---|------------|
        | **D.S. 024-2016-EM (Perú)** | IF = (ATP + FAT) × K ÷ HHT | 1,000,000 | ❌ No |
        | **OSHA (EE.UU.)** | TIFAR = (AM + ATP + FAT) × K ÷ HHT | 200,000 | ✅ Sí |
        | **ANSI Z16.1 (EE.UU.)** | TIFAR = (ATP + FAT) × K ÷ HHT | 1,000,000 | ❌ No |
        
        ### 🔍 Componentes
        
        - **ATP** = Accidentes con Tiempo Perdido (eventos)
        - **FAT** = Accidentes Fatales (eventos)
        - **AM** = Accidentes Médicos (eventos)
        - **w** = Ponderador (1 = incluye AM, 0 = excluye AM)
        - **HHT** = Horas Hombre Trabajadas (horas)
        - **K** = Factor de estandarización (1,000,000 o 200,000)
        
        ### 💡 Interpretación
        
        - **IF < 10** → Riesgo controlado ✅
        - **IF 10 - 25** → Riesgo moderado ⚠️
        - **IF 25 - 50** → Riesgo severo 🔴
        - **IF > 50** → Riesgo crítico 🔴🔴
        """)
        if st.button("Cerrar", key="close_if"):
            close_modal()

if st.session_state.modal == "is":
    with st.expander("📊 IS/TISSAR - Índice de Severidad", expanded=True):
        st.markdown("""
        ### 📐 Fórmula General
        
        $$
        \\text{IS} = \\frac{\\text{Días Cargados}}{\\text{HHT}} \\times K
        $$
        
        ### 📋 Detalle de Días Cargados
        
        $$
        \\text{Días Cargados} = \\text{DF} + (6000 \\times \\text{FAT})
        $$
        
        ### 🔍 Componentes
        
        - **DF** = Días Físicos Perdidos (días)
        - **FAT** = Accidentes Fatales (eventos)
        - **6,000** = Penalización estándar por fatalidad (días)
        - **HHT** = Horas Hombre Trabajadas (horas)
        - **K** = Factor de estandarización
        
        ### ⚖️ Base Normativa
        
        - **Anexo N°33 del D.S. 024-2016-EM** → 6,000 días por fatalidad
        - **ANSI Z16.1** → 6,000 días por fatalidad
        - **OSHA** → 6,000 días por fatalidad (estándar internacional)
        
        ### 💡 Interpretación
        
        - **IS < 50** → Severidad baja ✅
        - **IS 50 - 200** → Severidad moderada ⚠️
        - **IS 200 - 500** → Severidad alta 🔴
        - **IS > 500** → Severidad crítica 🔴🔴
        """)
        if st.button("Cerrar", key="close_is"):
            close_modal()

if st.session_state.modal == "ia":
    with st.expander("📊 IA - Índice de Accidentabilidad Global", expanded=True):
        st.markdown("""
        ### 📐 Fórmula General
        
        $$
        \\text{IA} = \\frac{\\text{IF} \\times \\text{IS}}{1000}
        $$
        
        ### 🔍 Componentes
        
        - **IF** = Índice de Frecuencia (accidentes / 1,000,000 horas)
        - **IS** = Índice de Severidad (días perdidos / 1,000,000 horas)
        - **÷ 1,000** = Factor de estandarización para escala manejable
        
        ### 📊 Matriz de Riesgo (IA)
        
        | Severidad \\ Frecuencia | Baja (IF<10) | Media (10-25) | Alta (25-50) | Crítica (>50) |
        |------------------------|--------------|---------------|--------------|---------------|
        | Baja (IS<50) | **Controlado** | Moderado | Significativo | Alto |
        | Media (50-200) | Moderado | **Significativo** | Alto | Muy Alto |
        | Alta (200-500) | Significativo | Alto | **Muy Alto** | Extremo |
        | Crítica (>500) | Alto | Muy Alto | Extremo | **Crítico** |
        
        ### 💡 Interpretación
        
        - **IA < 0.5** → Riesgo bajo ✅
        - **IA 0.5 - 2.0** → Riesgo moderado ⚠️
        - **IA 2.0 - 5.0** → Riesgo alto 🔴
        - **IA > 5.0** → Riesgo crítico 🔴🔴
        """)
        if st.button("Cerrar", key="close_ia"):
            close_modal()

if st.session_state.modal == "far":
    with st.expander("📊 FAR - Tasa de Accidentes Fatales", expanded=True):
        st.markdown("""
        ### 📐 Fórmula General
        
        $$
        \\text{FAR} = \\frac{\\text{FAT}}{\\text{HHT}} \\times 100,000,000
        $$
        
        ### 🔍 Componentes
        
        - **FAT** = Accidentes Fatales (eventos con pérdida de vida)
        - **HHT** = Horas Hombre Trabajadas (horas totales de exposición)
        - **100,000,000** = Base de exposición estándar ICMM (100 millones de horas)
        
        ### 🌍 Referencia ICMM
        
        > El FAR indica el **número de fatalidades que ocurrirían** si la exposición fuera de 100 millones de horas.
        
        ### 📊 Escala de Evaluación (ICMM)
        
        - **FAR = 0** → Sin fatalidades en el período evaluado ✅
        - **FAR entre 1 y 50** → Riesgo moderado (rango estándar industrial) ⚠️
        - **FAR > 50** → Alerta crítica, requiere intervención inmediata 🔴
        
        ### ⚠️ Nota Importante
        
        > Este indicador **NO es exigido por el D.S. 024-2016-EM (Perú)**.  
        > Es una **referencia internacional del ICMM** (International Council on Mining and Metals) utilizada para comparar tasas de mortalidad entre diferentes operaciones mineras a nivel global.
        """)
        if st.button("Cerrar", key="close_far"):
            close_modal()

st.markdown("---")

tab1, tab2, tab3 = st.tabs(["📊 Analisis Mensual", "📈 Series Temporales", "📋 Plantilla Excel"])

# ==================== TAB 1: ANÁLISIS MENSUAL ====================
with tab1:
    st.markdown("### 📝 Datos de Entrada")
    
    col_norm1, col_norm2 = st.columns([2, 1])
    with col_norm1:
        normativa_seleccionada = st.radio(
            "Seleccione la normativa:",
            options=list(NORMATIVAS.keys()),
            index=0,
            horizontal=True
        )
    
    config_seleccionada = NORMATIVAS[normativa_seleccionada]
    nombre_frec = config_seleccionada['nombre_frecuencia']
    nombre_sev = config_seleccionada['nombre_severidad']
    
    with col_norm2:
        st.info(f"K = {config_seleccionada['k_default']:,} horas | Incluye AM: {'SI' if config_seleccionada['incluye_am_en_frecuencia'] else 'NO'}")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**👥 Datos de Exposicion**")
        hht = st.number_input("HHT", min_value=1, value=88000, step=1000, key="hht")
        trabajadores = st.number_input("N° Trabajadores", min_value=1, value=500, step=10, key="trab")
        semanas = st.number_input("Semanas/año", min_value=40, max_value=52, value=50, key="sem")
        horas_semana = st.number_input("Horas/semana", min_value=20, max_value=60, value=40, key="hs")
        
        usar_k_auto = st.checkbox("Calcular K automaticamente", value=True)
        if usar_k_auto:
            k_factor = calc_k_estandarizacion(trabajadores, horas_semana, semanas)
            st.info(f"K = {k_factor:,} horas")
        else:
            k_factor = st.number_input("K manual", min_value=100000, value=config_seleccionada['k_default'], step=100000)
    
    with col2:
        st.markdown("**📋 Accidentes**")
        atp = st.number_input("ATP", min_value=0, value=3, key="atp")
        fat = st.number_input("FAT", min_value=0, value=1, key="fat")
        df_dias = st.number_input("DF", min_value=0, value=31, key="df")
        am = st.number_input("AM (referencia)", min_value=0, value=3, key="am")
    
    with col3:
        st.markdown("**📊 Parametros**")
        freq_prev = st.number_input(f"{nombre_frec} anterior", min_value=0.0, value=6.47, step=0.1)
        n_simulations = st.number_input("Simulaciones MC", min_value=1000, value=10000, step=5000)
        mostrar_reporte = st.checkbox("Generar reporte PDF", value=True)
    
    if st.button("🔍 Calcular", type="primary", use_container_width=True):
        params = {
            'hht': hht, 'am': am, 'atp': atp, 'fat': fat, 'df': df_dias,
            'k_factor': k_factor, 'freq_prev': freq_prev, 'n_simulations': n_simulations,
            'trabajadores': trabajadores, 'semanas': semanas, 'horas_semana': horas_semana,
            'incluye_am_en_frecuencia': config_seleccionada['incluye_am_en_frecuencia'],
            'nombre_frecuencia': nombre_frec, 'nombre_severidad': nombre_sev
        }
        
        freq_val = calc_if(atp, fat, am, k_factor, hht, normativa_seleccionada)
        dias_cargados = calc_dias_cargados(df_dias, fat)
        sev_val = calc_is(dias_cargados, k_factor, hht)
        ia_val = calc_ia(freq_val, sev_val)
        far_val = calc_far(fat, hht)
        dmi_val = calc_dmi(df_dias, atp)
        variacion = calc_variacion(freq_val, freq_prev)
        
        deterministic = {
            'frecuencia': freq_val, 'dias_cargados': dias_cargados,
            'severidad': sev_val, 'ia': ia_val, 'far': far_val, 'dmi': dmi_val, 'variacion': variacion
        }
        
        with st.spinner("Simulando Monte Carlo..."):
            stats_results, correlations, samples = monte_carlo(params, n_simulations)
        
        st.markdown("---")
        st.markdown(f"### 📊 Resultados ({normativa_seleccionada})")
        
        r1, r2, r3, r4 = st.columns(4)
        r1.metric(nombre_frec, f"{freq_val:.4f}")
        r2.metric(nombre_sev, f"{sev_val:.2f}")
        r3.metric("IA", f"{ia_val:.2f}")
        r4.metric("FAR", f"{far_val:.2f}")
        
        r5, r6, r7, r8 = st.columns(4)
        r5.metric("DMI", f"{dmi_val:.2f} dias/acc")
        r6.metric("Dias Cargados", f"{dias_cargados:,.0f}")
        r7.metric("Variacion", f"{variacion:+.2f}%", delta_color="normal" if variacion <= 0 else "inverse")
        r8.metric("AM (ref)", f"{am}")
        
        st.markdown("---")
        st.markdown("### 📈 Evaluacion de Riesgo")
        if freq_val < 10:
            st.success("✅ RIESGO CONTROLADO - Desempeno favorable")
        elif freq_val < 25:
            st.warning("⚠️ RIESGO MODERADO - Requiere seguimiento")
        elif freq_val < 50:
            st.warning("🔴 RIESGO SEVERO - Acciones correctivas")
        else:
            st.error("🔴🔴 RIESGO CRITICO - Intervencion inmediata")
        
        st.markdown("#### 🎯 Perfil de Riesgo (Radar Chart)")
        fig_radar, valores_norm = generar_grafico_radar(deterministic, params, nombre_frec, nombre_sev, normativa_seleccionada)
        st.pyplot(fig_radar)
        plt.close(fig_radar)
        st.markdown(f'<div class="interpretacion-box">📖 Interpretacion: {interpretacion_radar(valores_norm, nombre_frec)}</div>', unsafe_allow_html=True)
        
        st.markdown("#### 📊 Distribuciones Monte Carlo")
        fig_hist = generar_grafico_histogramas(stats_results, nombre_frec, nombre_sev)
        st.pyplot(fig_hist)
        plt.close(fig_hist)
        st.markdown(f'<div class="interpretacion-box">📖 Interpretacion: {interpretacion_histogramas(stats_results, nombre_frec)}</div>', unsafe_allow_html=True)
        
        g1, g2 = st.columns(2)
        with g1:
            st.markdown("#### 🔥 Diagrama de Tornado")
            fig_tornado = generar_grafico_tornado(correlations, nombre_frec)
            st.pyplot(fig_tornado)
            plt.close(fig_tornado)
            st.markdown(f'<div class="interpretacion-box">📖 Interpretacion: {interpretacion_tornado(correlations, nombre_frec)}</div>', unsafe_allow_html=True)
        
        with g2:
            st.markdown("#### 🥧 Distribucion de Eventos")
            fig_pastel = generar_grafico_pastel(params)
            st.pyplot(fig_pastel)
            plt.close(fig_pastel)
            st.markdown(f'<div class="interpretacion-box">📖 Interpretacion: {interpretacion_pastel(am, atp, fat)}</div>', unsafe_allow_html=True)
        
        st.markdown("#### 📈 Analisis Predictivo OLS")
        fig_trend, slope, r2 = generar_grafico_ols_trend(params, deterministic, freq_prev, nombre_frec)
        st.pyplot(fig_trend)
        plt.close(fig_trend)
        st.markdown(f'<div class="interpretacion-box">📖 Interpretacion: {interpretacion_ols_trend(slope, r2, nombre_frec)}</div>', unsafe_allow_html=True)
        
        st.markdown("### 🎲 Monte Carlo - Percentiles")
        mc_stats = {}
        for var in [nombre_frec, nombre_sev, 'IA', 'FAR']:
            mc_stats[var] = {
                'Media': f"{stats_results[var]['mean']:.4f}",
                'P5': f"{stats_results[var]['p5']:.4f}",
                'P50': f"{stats_results[var]['p50']:.4f}",
                'P95': f"{stats_results[var]['p95']:.4f}"
            }
        st.dataframe(pd.DataFrame(mc_stats).T, use_container_width=True)
        
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            excel_data = export_excel_mensual(params, deterministic, stats_results, correlations, samples, normativa_seleccionada)
            st.download_button("📥 Exportar Excel", excel_data, f"SSO03_{nombre_frec}.xlsx", use_container_width=True)
        with col_d2:
            if mostrar_reporte:
                interpretaciones = {
                    'radar': interpretacion_radar(valores_norm, nombre_frec),
                    'histogramas': interpretacion_histogramas(stats_results, nombre_frec),
                    'tornado': interpretacion_tornado(correlations, nombre_frec),
                    'pastel': interpretacion_pastel(am, atp, fat),
                    'ols': interpretacion_ols_trend(slope, r2, nombre_frec)
                }
                pdf_data = generar_reporte_pdf_mensual(params, deterministic, stats_results, correlations, normativa_seleccionada, interpretaciones)
                st.download_button("📄 Exportar PDF", pdf_data, f"SSO03_Reporte_{nombre_frec}.pdf", use_container_width=True)

# ==================== TAB 2: ANÁLISIS HISTÓRICO ====================
with tab2:
    st.markdown("### 📂 Carga de Datos Historicos")
    st.info("Columnas obligatorias: Periodo (AAAA-MM), HHT, AM, ATP, FAT, DF")
    
    usar_k_por_periodo = st.checkbox("Calcular K por periodo", value=False)
    if not usar_k_por_periodo:
        k_factor_hist = st.number_input("K fijo", min_value=100000, value=config_seleccionada['k_default'], step=100000)
    
    uploaded_file = st.file_uploader("Seleccionar Excel", type=['xlsx', 'xls'])
    
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        with st.expander("📋 Vista previa de los datos cargados"):
            st.dataframe(df.head(10), use_container_width=True)
            st.caption(f"Columnas encontradas: {list(df.columns)}")
        
        if 'AM' not in df.columns:
            df['AM'] = 0
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            forecast_months = st.number_input("Meses proyeccion", min_value=1, max_value=36, value=12)
            anomaly_threshold = st.number_input("Umbral anomalias", min_value=1.5, max_value=4.0, value=2.5, step=0.1)
        with col_f2:
            mc_check = st.checkbox("Incluir Monte Carlo", value=True)
            mc_samples = 10000 if mc_check else 0
        
        if st.button("🔍 Analizar", type="primary", use_container_width=True):
            with st.spinner("Procesando..."):
                if usar_k_por_periodo and all(col in df.columns for col in ['Trabajadores', 'Semanas', 'Horas_Semana']):
                    k_factors = []
                    for i, row in df.iterrows():
                        k = calc_k_estandarizacion(row['Trabajadores'], row['Horas_Semana'], row['Semanas'])
                        k_factors.append(k)
                    k_factor_hist = np.mean(k_factors)
                    st.info(f"K promedio: {k_factor_hist:,.0f} horas")
                elif usar_k_por_periodo:
                    st.error("No se encontraron columnas")
                    k_factor_hist = config_seleccionada['k_default']
                else:
                    if 'k_factor_hist' not in locals():
                        k_factor_hist = config_seleccionada['k_default']
                
                resultados = analisis_historico(df, k_factor_hist, forecast_months, anomaly_threshold, mc_check, mc_samples, normativa_seleccionada)
            
            st.success(f"✅ Analisis completado: {resultados['n_meses']} meses")
            
            ols = resultados['modelo_ols']
            st.markdown("### 📈 Modelo OLS")
            o1, o2, o3 = st.columns(3)
            o1.metric("Pendiente", f"{ols['slope']:.4f}")
            o2.metric("R²", f"{ols['r2']:.4f}")
            o3.metric("Tendencia", ols['tendencia_texto'])
            
            st.markdown("#### 📦 Distribucion Anual")
            fig_box = generar_grafico_boxplot_anual(resultados['df_historico'], resultados['nombre_frec'])
            st.pyplot(fig_box)
            plt.close(fig_box)
            st.markdown(f'<div class="interpretacion-box">📖 Interpretacion: {interpretacion_boxplot_anual(resultados["df_historico"], resultados["nombre_frec"])}</div>', unsafe_allow_html=True)
            
            st.markdown("#### 🗺️ Heatmap Mensual")
            fig_heat = generar_grafico_heatmap(resultados['df_historico'], resultados['nombre_frec'])
            st.pyplot(fig_heat)
            plt.close(fig_heat)
            st.markdown(f'<div class="interpretacion-box">📖 Interpretacion: {interpretacion_heatmap(resultados["df_historico"], resultados["nombre_frec"])}</div>', unsafe_allow_html=True)
            
            st.markdown("#### 📈 Tendencia y Estacionalidad")
            fig_tend = generar_grafico_tendencia_estacional(resultados['df_historico'], ols)
            st.pyplot(fig_tend)
            plt.close(fig_tend)
            st.markdown(f'<div class="interpretacion-box">📖 Interpretacion: {interpretacion_ols_trend(ols["slope"], ols["r2"], resultados["nombre_frec"])}</div>', unsafe_allow_html=True)
            
            st.markdown("#### 📅 Comparativa por Año")
            fig_comp = generar_grafico_comparativo_anual(resultados['df_historico'], resultados['nombre_frec'])
            st.pyplot(fig_comp)
            plt.close(fig_comp)
            
            if resultados['mc_forecast']:
                st.markdown("#### 🔮 Pronostico Monte Carlo")
                fig_mc = generar_grafico_mc_forecast(resultados['forecast'], resultados['mc_forecast'], resultados['nombre_frec'])
                st.pyplot(fig_mc)
                plt.close(fig_mc)
                st.markdown(f'<div class="interpretacion-box">📖 Interpretacion: {interpretacion_mc_forecast(resultados["mc_forecast"], resultados["nombre_frec"])}</div>', unsafe_allow_html=True)
            
            st.markdown("### 📅 Resumen Anual")
            st.dataframe(resultados['resumen_anual'].round(4), use_container_width=True)
            
            if len(resultados['anomalias']) > 0:
                st.markdown("### 🚨 Anomalias Detectadas")
                st.dataframe(resultados['anomalias'][['Periodo', resultados['nombre_frec'], resultados['nombre_sev'], 'IA']], use_container_width=True)
            
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                excel_data = export_historico_excel(resultados)
                st.download_button("📥 Exportar Excel", excel_data, "SSO03_Historico.xlsx", use_container_width=True)
            with col_e2:
                interpretaciones_hist = {
                    'boxplot': interpretacion_boxplot_anual(resultados['df_historico'], resultados['nombre_frec']),
                    'heatmap': interpretacion_heatmap(resultados['df_historico'], resultados['nombre_frec']),
                    'tendencia': interpretacion_ols_trend(ols['slope'], ols['r2'], resultados['nombre_frec']),
                    'comparativa': "El grafico comparativo anual permite visualizar la evolucion mes a mes. Si un año completo esta sistematicamente por encima de los anteriores, la tendencia es negativa.",
                    'forecast': interpretacion_mc_forecast(resultados['mc_forecast'], resultados['nombre_frec']) if resultados['mc_forecast'] else "No se genero pronostico"
                }
                pdf_data = generar_reporte_pdf_historico(resultados, interpretaciones_hist)
                st.download_button("📄 Exportar PDF", pdf_data, "SSO03_Historico.pdf", use_container_width=True)

# ==================== TAB 3: PLANTILLA EXCEL ====================
with tab3:
    st.markdown("### 📋 Plantilla Excel")
    st.markdown("""
    **Columnas requeridas:**
    - **Periodo**: Fecha en formato AAAA-MM (ej: 2024-01)
    - **HHT**: Horas Hombre Trabajadas
    - **AM**: Accidentes Medicos
    - **ATP**: Accidentes con Tiempo Perdido
    - **FAT**: Accidentes Fatales
    - **DF**: Dias Fisicos Perdidos
    
    **Notas importantes:**
    - Complete un registro por cada mes
    - Mantenga el formato de fecha AAAA-MM
    - Cada accidente fatal suma 6,000 dias cargados
    """)
    
    incluir_k = st.checkbox("Incluir columnas para calculo de K (Trabajadores, Horas_Semana, Semanas, K_calculada)", value=False)
    
    col_preview1, col_preview2 = st.columns(2)
    
    with col_preview1:
        if st.button("📋 Vista Previa de Plantilla", use_container_width=True):
            if incluir_k:
                datos_preview = {
                    'Periodo': ['2024-01', '2024-02', '2024-03'],
                    'HHT': [88000, 85600, 88400],
                    'AM': [3, 2, 1],
                    'ATP': [2, 1, 0],
                    'FAT': [0, 0, 0],
                    'DF': [15, 10, 0],
                    'Trabajadores': [500, 495, 502],
                    'Horas_Semana': [40, 40, 40],
                    'Semanas': [50, 50, 50],
                    'K_calculada': [1000000, 990000, 1004000]
                }
            else:
                datos_preview = {
                    'Periodo': ['2024-01', '2024-02', '2024-03'],
                    'HHT': [88000, 85600, 88400],
                    'AM': [3, 2, 1],
                    'ATP': [2, 1, 0],
                    'FAT': [0, 0, 0],
                    'DF': [15, 10, 0]
                }
            df_preview = pd.DataFrame(datos_preview)
            st.dataframe(df_preview, use_container_width=True)
    
    with col_preview2:
        if st.button("📋 Generar Plantilla", use_container_width=True, type="primary"):
            plantilla = generar_plantilla_excel(incluir_k)
            st.download_button("📥 Descargar Plantilla", plantilla, "Plantilla_SSO03.xlsx", use_container_width=True)

st.markdown("---")
st.markdown(f"""
<div class="footer">
    <p>⛏️ SSO-03 - Normativas: D.S. 024-2016-EM (Peru) | OSHA | ANSI Z16.1</p>
</div>
""", unsafe_allow_html=True)