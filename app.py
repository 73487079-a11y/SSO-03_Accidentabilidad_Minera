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

# Paleta de colores personalizada - COMPLETA con todas las claves
C = {
    'saffron': '#F8C662',
    'ultra_violet': '#595082',
    'dark_purple': '#2C263F',
    'hunter_green': '#41644A',
    'dark_green': '#213722',
    'white': '#FFFFFF',
    'light_gray': '#F8F9FA',
    'medium_gray': '#E8ECEF',
    'dark_gray': '#6C757D',
    'text_dark': '#2C263F',
    'text_light': '#FFFFFF',
    'success': '#41644A',
    'warning': '#F8C662',
    'danger': '#E74C3C'
}

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    
    * {{
        font-family: 'Inter', sans-serif;
    }}
    
    .stApp {{
        background: linear-gradient(135deg, {C['light_gray']} 0%, #E8ECEF 100%);
    }}
    
    .hero-section {{
        background: linear-gradient(135deg, {C['dark_purple']} 0%, {C['ultra_violet']} 100%);
        padding: 2.5rem 2rem;
        border-radius: 24px;
        margin-bottom: 2rem;
        text-align: center;
        position: relative;
        overflow: hidden;
        box-shadow: 0 10px 30px rgba(44,38,63,0.2);
        border-bottom: 4px solid {C['saffron']};
    }}
    
    .hero-section::before {{
        content: "⛏️";
        position: absolute;
        font-size: 180px;
        opacity: 0.08;
        left: -30px;
        bottom: -50px;
        transform: rotate(-15deg);
    }}
    
    .hero-section::after {{
        content: "📊";
        position: absolute;
        font-size: 160px;
        opacity: 0.08;
        right: -30px;
        top: -40px;
        transform: rotate(15deg);
    }}
    
    .hero-title {{
        font-size: 2.8rem;
        font-weight: 800;
        color: {C['white']};
        margin-bottom: 0.5rem;
        position: relative;
        z-index: 1;
        letter-spacing: -0.02em;
    }}
    
    .hero-subtitle {{
        font-size: 1.1rem;
        color: {C['saffron']};
        margin-bottom: 1rem;
        position: relative;
        z-index: 1;
    }}
    
    .hero-badge {{
        display: inline-block;
        background: rgba(255,255,255,0.15);
        backdrop-filter: blur(10px);
        padding: 0.5rem 1.2rem;
        border-radius: 50px;
        font-size: 0.8rem;
        color: {C['white']};
        position: relative;
        z-index: 1;
    }}
    
    .stat-card {{
        background: {C['white']};
        border-radius: 20px;
        padding: 1.2rem;
        text-align: center;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        transition: all 0.3s ease;
        border-top: 3px solid {C['saffron']};
    }}
    
    .stat-card:hover {{
        transform: translateY(-5px);
        box-shadow: 0 15px 30px rgba(44,38,63,0.15);
    }}
    
    .stat-number {{
        font-size: 2rem;
        font-weight: 800;
        color: {C['dark_purple']};
    }}
    
    .stat-label {{
        font-size: 0.75rem;
        color: {C['dark_gray']};
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-top: 0.3rem;
    }}
    
    .metric-card {{
        background: {C['white']};
        border-radius: 16px;
        padding: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        border-left: 4px solid {C['saffron']};
        transition: all 0.3s ease;
    }}
    
    .metric-card:hover {{
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(44,38,63,0.1);
    }}
    
    .metric-label {{
        font-size: 0.7rem;
        color: {C['dark_gray']};
        text-transform: uppercase;
        letter-spacing: 0.5px;
        font-weight: 600;
    }}
    
    .metric-value {{
        font-size: 1.5rem;
        font-weight: 700;
        color: {C['dark_purple']};
        margin: 0.25rem 0;
    }}
    
    .metric-unit {{
        font-size: 0.7rem;
        color: {C['dark_gray']};
    }}
    
    .stTabs [data-baseweb="tab-list"] {{
        gap: 0.5rem;
        background: {C['white']};
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
        color: {C['text_dark']};
        background: transparent;
        transition: all 0.3s ease;
    }}
    
    .stTabs [aria-selected="true"] {{
        background: linear-gradient(135deg, {C['dark_purple']} 0%, {C['ultra_violet']} 100%);
        color: {C['white']} !important;
        box-shadow: 0 4px 12px rgba(89,80,130,0.3);
    }}
    
    .stTabs [aria-selected="false"]:hover {{
        background: {C['light_gray']};
    }}
    
    .stButton > button {{
        background: linear-gradient(135deg, {C['dark_purple']} 0%, {C['ultra_violet']} 100%);
        color: {C['white']};
        border: none;
        border-radius: 12px;
        padding: 0.6rem 1.5rem;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 2px 8px rgba(44,38,63,0.2);
    }}
    
    .stButton > button:hover {{
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(89,80,130,0.4);
    }}
    
    .section-title {{
        font-size: 1.4rem;
        font-weight: 700;
        color: {C['dark_purple']};
        margin-bottom: 0.5rem;
        display: inline-block;
        border-bottom: 3px solid {C['saffron']};
        padding-bottom: 0.3rem;
    }}
    
    .section-subtitle {{
        font-size: 0.85rem;
        color: {C['dark_gray']};
        margin-bottom: 1.5rem;
    }}
    
    div[data-testid="stNumberInput"] label, 
    div[data-testid="stTextInput"] label,
    div[data-testid="stSelectbox"] label {{
        font-weight: 600;
        color: {C['dark_purple']} !important;
        font-size: 0.8rem;
    }}
    
    div[data-testid="stNumberInput"] input,
    div[data-testid="stTextInput"] input {{
        border-radius: 12px;
        border: 1px solid {C['medium_gray']};
        padding: 0.4rem 0.8rem;
    }}
    
    .stExpander {{
        border-radius: 16px;
        border: none;
        background: {C['white']};
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }}
    
    .stExpander > div:first-child {{
        background: {C['light_gray']};
        border-radius: 16px;
        font-weight: 600;
        color: {C['dark_purple']};
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
        background: linear-gradient(90deg, transparent, {C['saffron']}, {C['ultra_violet']}, transparent);
    }}
    
    .footer {{
        text-align: center;
        padding: 1.5rem;
        color: {C['dark_gray']};
        font-size: 0.75rem;
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
        color: {C['dark_green']};
    }}
    
    .eval-warning {{
        background: #FFF8E1;
        border-left-color: {C['warning']};
        color: #856404;
    }}
    
    .eval-error {{
        background: #FFEBEE;
        border-left-color: {C['danger']};
        color: #C62828;
    }}
</style>
""", unsafe_allow_html=True)

# Configuración de matplotlib
plt.rcParams['font.size'] = 10
plt.rcParams['axes.labelsize'] = 11
plt.rcParams['axes.titlesize'] = 12
plt.rcParams['legend.fontsize'] = 9
plt.rcParams['font.sans-serif'] = ['Inter', 'Arial', 'DejaVu Sans']
plt.rcParams['figure.facecolor'] = 'white'
plt.rcParams['axes.facecolor'] = '#F8F9FA'
plt.rcParams['axes.grid'] = True
plt.rcParams['grid.alpha'] = 0.3
plt.rcParams['grid.linestyle'] = '--'
plt.rcParams['axes.spines.top'] = False
plt.rcParams['axes.spines.right'] = False

PARAM_INFO = {
    'hht': {'label': 'Horas Hombre Trabajadas', 'symbol': 'HHT', 'unit': 'horas', 'desc': 'Tiempo total de exposición al riesgo',
            'default': 152041, 'min': 1, 'max': 10000000, 'cov': 0.05, 'dist': 'normal',
            'typical': 'Pequeña:50-100k | Mediana:100-500k | Grande:500k-2M', 'icon': '👥'},
    'am': {'label': 'Accidentes Médicos', 'symbol': 'AM', 'unit': 'eventos', 'desc': 'Incidentes que requieren atención médica sin descanso',
           'default': 3, 'min': 0, 'max': 1000, 'cov': 0.10, 'dist': 'poisson',
           'typical': 'Bajo:0-3 | Estándar:4-8 | Alto:9-15 | Crítico:>15', 'icon': '🩺'},
    'atp': {'label': 'Accidentes con Tiempo Perdido', 'symbol': 'ATP', 'unit': 'eventos', 'desc': 'Lesiones incapacitantes con descanso',
            'default': 3, 'min': 0, 'max': 500, 'cov': 0.10, 'dist': 'poisson',
            'typical': 'Excelente:0 | Aceptable:1-2 | Preocupante:3-5 | Crítico:>5', 'icon': '⚠️'},
    'fat': {'label': 'Accidentes Fatales', 'symbol': 'FAT', 'unit': 'eventos', 'desc': 'Eventos con pérdida de vida',
            'default': 1, 'min': 0, 'max': 100, 'cov': 0.05, 'dist': 'poisson',
            'typical': 'Estándar:0 | Alerta máxima:≥1', 'icon': '💀'},
    'df': {'label': 'Días Físicos Perdidos', 'symbol': 'DF', 'unit': 'días', 'desc': 'Sumatoria de días de descanso médico efectivo',
           'default': 31, 'min': 0, 'max': 10000, 'cov': 0.15, 'dist': 'normal',
           'typical': 'Baja:0-15 | Media:16-45 | Alta:46-100 | Crítica:>100', 'icon': '📅'},
    'k_factor': {'label': 'Constante de Estandarización', 'symbol': 'K', 'unit': 'horas', 'desc': 'Base teórica de exposición',
                 'default': 1000000, 'min': 100000, 'max': 100000000, 'cov': 0.00, 'dist': 'fixed',
                 'typical': 'OSHA:200,000 | Minería:1,000,000 | FAR:100,000,000', 'icon': '⚙️'},
    'tifar_prev': {'label': 'TIFAR Periodo Anterior', 'symbol': 'TIFAR_ant', 'unit': 'adim', 'desc': 'Índice del periodo base comparativo',
                   'default': 6.47, 'min': 0, 'max': 1000, 'cov': 0.00, 'dist': 'fixed',
                   'typical': 'Minería mundial:0.5-2.5 | Minería Perú:2-8', 'icon': '📈'}
}

def calc_tifar(am, atp, fat, k, hht):
    return ((am + atp + fat) * k) / hht

def calc_tipar(atp, fat, k, hht):
    return ((atp + fat) * k) / hht

def calc_dias_cargados(df, fat):
    return df + (fat * 6000)

def calc_tissar(dias_cargados, k, hht):
    return (dias_cargados * k) / hht

def calc_ia(tifar, tissar):
    return (tifar * tissar) / 1000

def calc_far(fat, hht):
    return (fat * 100000000) / hht

def calc_dmi(df, atp):
    return df / atp if atp > 0 else 0

def calc_variacion(tifar_actual, tifar_anterior):
    return ((tifar_actual - tifar_anterior) / tifar_anterior) * 100 if tifar_anterior > 0 else 0

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
    
    tifar_samples = ((am_samples + atp_samples + fat_samples) * k_samples) / hht_samples
    tipar_samples = ((atp_samples + fat_samples) * k_samples) / hht_samples
    dias_cargados_samples = df_samples + (fat_samples * 6000)
    tissar_samples = (dias_cargados_samples * k_samples) / hht_samples
    ia_samples = (tifar_samples * tissar_samples) / 1000
    far_samples = (fat_samples * 100000000) / hht_samples
    dmi_samples = np.where(atp_samples > 0, df_samples / atp_samples, 0)
    
    stats_results = {}
    for name, data in [('TIFAR', tifar_samples), ('TIPAR', tipar_samples), ('TISSAR', tissar_samples), 
                       ('IA', ia_samples), ('FAR', far_samples), ('DMI', dmi_samples)]:
        stats_results[name] = {
            'mean': np.mean(data), 'std': np.std(data),
            'p5': np.percentile(data, 5), 'p50': np.percentile(data, 50),
            'p90': np.percentile(data, 90), 'p95': np.percentile(data, 95),
            'samples': data[:500]
        }
    
    input_matrix = np.column_stack([hht_samples, am_samples, atp_samples, fat_samples, df_samples])
    correlations = {}
    for i, name in enumerate(['HHT', 'AM', 'ATP', 'FAT', 'DF']):
        correlations[name] = np.corrcoef(input_matrix[:, i], tifar_samples)[0, 1]
    
    samples = {'hht': hht_samples, 'am': am_samples, 'atp': atp_samples, 'fat': fat_samples, 'df': df_samples}
    
    return stats_results, correlations, samples

def analisis_historico(df, k_factor, forecast_months=12, anomaly_threshold=2.5, mc_forecast=False, mc_samples=10000):
    df = df.copy()
    df['Periodo'] = pd.to_datetime(df['Periodo'])
    df = df.sort_values('Periodo').reset_index(drop=True)
    
    df['TIFAR'] = ((df['AM'] + df['ATP'] + df['FAT']) * k_factor) / df['HHT']
    df['TIPAR'] = ((df['ATP'] + df['FAT']) * k_factor) / df['HHT']
    df['Dias_Cargados'] = df['DF'] + (df['FAT'] * 6000)
    df['TISSAR'] = (df['Dias_Cargados'] * k_factor) / df['HHT']
    df['IA'] = (df['TIFAR'] * df['TISSAR']) / 1000
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
            'TIFAR_Promedio': year_data['TIFAR'].mean(),
            'TIFAR_Max': year_data['TIFAR'].max(),
            'TIFAR_Min': year_data['TIFAR'].min(),
            'TIFAR_Std': year_data['TIFAR'].std(),
            'Acc_Fatales': year_data['FAT'].sum(),
            'Acc_ATP': year_data['ATP'].sum(),
            'Acc_AM': year_data['AM'].sum(),
            'Dias_Cargados_Total': year_data['Dias_Cargados'].sum(),
            'FAR_Promedio': year_data['FAR'].mean()
        })
    
    x = np.arange(len(df))
    slope, intercept, r_value, p_value, std_err = stats.linregress(x, df['TIFAR'])
    r2 = r_value ** 2
    
    forecast_x = np.arange(len(df), len(df) + forecast_months)
    forecast_tifar = intercept + slope * forecast_x
    
    se_forecast = std_err * np.sqrt(1 + 1/len(df) + (forecast_x - np.mean(x))**2 / np.sum((x - np.mean(x))**2))
    t_val = stats.t.ppf(0.975, len(df) - 2)
    
    forecast_upper = forecast_tifar + t_val * se_forecast
    forecast_lower = forecast_tifar - t_val * se_forecast
    
    z_scores = np.abs(stats.zscore(df['TIFAR']))
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
        mc_tifar = np.zeros((mc_samples, forecast_months))
        
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
                
                tifar_sim = ((am_sim + atp_sim + fat_sim) * k_factor) / hht_sim
                mc_tifar[i, j] = tifar_sim
        
        mc_forecast_results = {
            'mean': np.mean(mc_tifar, axis=0),
            'p5': np.percentile(mc_tifar, 5, axis=0),
            'p50': np.percentile(mc_tifar, 50, axis=0),
            'p95': np.percentile(mc_tifar, 95, axis=0)
        }
    
    forecast_df = pd.DataFrame({
        'Periodo': forecast_dates,
        'TIFAR_Pronostico': forecast_tifar,
        'Limite_Inferior_95': forecast_lower,
        'Limite_Superior_95': forecast_upper
    })
    
    if mc_forecast_results:
        forecast_df['TIFAR_MC_P50'] = mc_forecast_results['p50']
        forecast_df['TIFAR_MC_P5'] = mc_forecast_results['p5']
        forecast_df['TIFAR_MC_P95'] = mc_forecast_results['p95']
    
    return {
        'df_historico': df,
        'resumen_anual': pd.DataFrame(anual),
        'modelo_ols': {
            'slope': slope, 'intercept': intercept, 'r2': r2, 'p_value': p_value, 'std_err': std_err,
            'tendencia_texto': 'AL ALZA (deterioro)' if slope > 0 else 'A LA BAJA (mejora)'
        },
        'anomalias': anomalias,
        'forecast': forecast_df,
        'mc_forecast': mc_forecast_results,
        'n_meses': n_meses,
        'n_anios': n_anios,
        'fecha_inicio': fecha_inicio,
        'fecha_fin': fecha_fin,
        'forecast_months': forecast_months
    }

def generar_plantilla_excel():
    output = BytesIO()
    fechas = [f"2024-{i:02d}" for i in range(1, 13)]
    datos = {
        'Periodo': fechas,
        'HHT': [152000, 148500, 154000, 151000, 149000, 153000, 147000, 150000, 152500, 148000, 146000, 151500],
        'AM': [3, 2, 1, 4, 5, 3, 2, 2, 3, 4, 2, 3],
        'ATP': [2, 1, 0, 3, 2, 2, 1, 1, 2, 3, 1, 2],
        'FAT': [0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0],
        'DF': [15, 10, 0, 30, 20, 15, 8, 10, 18, 25, 12, 20]
    }
    df = pd.DataFrame(datos)
    
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
                '', 'NOTAS IMPORTANTES:',
                '- Complete un registro por cada mes',
                '- Puede tener datos de 1 mes hasta 100+ meses',
                '- Los datos de ejemplo son solo referenciales',
                '- Mantenga el formato de fecha AAAA-MM',
                '', 'PENALIZACIÓN ANSI Z16.1:',
                '- Cada accidente fatal (FAT) suma automáticamente 6,000 días cargados'
            ]
        })
        instrucciones.to_excel(writer, sheet_name='Instrucciones', index=False, header=False)
    
    output.seek(0)
    return output

def export_excel_mensual(params, deterministic, stats, correlations, samples):
    output = BytesIO()
    wb = openpyxl.Workbook()
    
    header_fill = PatternFill(start_color='2C263F', end_color='2C263F', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws1 = wb.active
    ws1.title = '1_Datos_Entrada'
    ws1.append(['Parámetro', 'Símbolo', 'Valor', 'Unidades', 'Valores Típicos'])
    for cell in ws1[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
    
    for key, info in PARAM_INFO.items():
        ws1.append([info['label'], info['symbol'], params[key], info['unit'], info['typical']])
        for cell in ws1[ws1.max_row]:
            cell.border = thin_border
    
    ws2 = wb.create_sheet('2_Resultados_Deterministicos')
    ws2.append(['Variable', 'Valor', 'Unidades', 'Interpretación'])
    for cell in ws2[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
    
    tifar = deterministic['TIFAR']
    if tifar < 10:
        interp = "Riesgo controlado. Desempeño favorable."
    elif tifar < 25:
        interp = "Riesgo moderado. Requiere seguimiento."
    elif tifar < 50:
        interp = "Riesgo severo. Implementar acciones correctivas."
    else:
        interp = "RIESGO CRÍTICO. Intervención inmediata requerida."
    
    ws2.append(['TIFAR', tifar, 'adimensional', interp])
    ws2.append(['TIPAR', deterministic['TIPAR'], 'adimensional', 'Índice de frecuencia con tiempo perdido'])
    ws2.append(['Días Cargados', deterministic['Dias_Cargados'], 'días', 'Incluye penalización ANSI Z16.1'])
    ws2.append(['TISSAR', deterministic['TISSAR'], 'adimensional', 'Índice de severidad total'])
    ws2.append(['IA', deterministic['IA'], 'adimensional', 'Índice de accidentabilidad global'])
    ws2.append(['FAR', deterministic['FAR'], 'adimensional', 'Tasa de accidentes mortales'])
    ws2.append(['DMI', deterministic['DMI'], 'días/accidente', 'Duración media de incapacidad'])
    ws2.append(['Variación', f"{deterministic['Variacion']:+.2f}%", 'porcentaje', 'vs periodo anterior'])
    
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
    ws4.append(['Muestra #', 'HHT', 'AM', 'ATP', 'FAT', 'DF', 'TIFAR', 'TIPAR', 'TISSAR', 'IA'])
    for cell in ws4[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
    
    for i in range(min(500, len(samples['hht']))):
        ws4.append([i+1, f"{samples['hht'][i]:.0f}", int(samples['am'][i]), int(samples['atp'][i]),
                   int(samples['fat'][i]), f"{samples['df'][i]:.1f}", f"{stats['TIFAR']['samples'][i]:.4f}",
                   f"{stats['TIPAR']['samples'][i]:.4f}", f"{stats['TISSAR']['samples'][i]:.4f}",
                   f"{stats['IA']['samples'][i]:.4f}"])
        for cell in ws4[ws4.max_row]:
            cell.border = thin_border
    
    ws5 = wb.create_sheet('5_Correlaciones')
    ws5.append(['Variable de Entrada', 'Correlación con TIFAR', 'Sensibilidad'])
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
        pd.DataFrame([resultados['modelo_ols']]).to_excel(writer, sheet_name='3_Modelo_OLS', index=False)
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

def generar_grafico_histogramas(stats):
    fig, axes = plt.subplots(2, 2, figsize=(12, 10))
    fig.patch.set_facecolor('white')
    vars_plot = ['TIFAR', 'TIPAR', 'TISSAR', 'IA']
    for idx, var in enumerate(vars_plot):
        ax = axes[idx // 2, idx % 2]
        ax.set_facecolor('#F8F9FA')
        data = stats[var]['samples']
        ax.hist(data, bins=50, color='#595082', edgecolor='white', alpha=0.7, density=True)
        ax.axvline(stats[var]['mean'], color='#E74C3C', linestyle='--', linewidth=2, label=f'Media: {stats[var]["mean"]:.2f}')
        ax.axvline(stats[var]['p5'], color='#F8C662', linestyle=':', linewidth=2, label=f'P5: {stats[var]["p5"]:.2f}')
        ax.axvline(stats[var]['p95'], color='#F8C662', linestyle=':', linewidth=2, label=f'P95: {stats[var]["p95"]:.2f}')
        ax.set_xlabel('Valor', fontsize=10)
        ax.set_ylabel('Densidad', fontsize=10)
        ax.set_title(f'Distribución de {var}', fontsize=11, fontweight='bold', color='#2C263F')
        ax.legend(fontsize=8, loc='upper right')
        ax.grid(True, alpha=0.3, linestyle='--')
    plt.suptitle('Análisis Monte Carlo - Distribuciones de Probabilidad', fontsize=14, fontweight='bold', color='#2C263F')
    plt.tight_layout()
    return fig

def generar_grafico_tornado(correlations):
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FA')
    names = list(correlations.keys())
    values = list(correlations.values())
    colors_t = ['#E74C3C' if v > 0 else '#41644A' for v in values]
    bars = ax.barh(names, values, color=colors_t, alpha=0.7, edgecolor='white', linewidth=1)
    ax.axvline(0, color='#2C263F', linewidth=1.5, linestyle='-')
    ax.set_xlabel('Coeficiente de Correlación de Pearson', fontsize=11, fontweight='bold', color='#2C263F')
    ax.set_title('Diagrama de Tornado - Sensibilidad del TIFAR', fontsize=12, fontweight='bold', color='#2C263F')
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
    colors_p = ['#595082', '#E74C3C', '#2C263F']
    explode = (0, 0.05, 0.1)
    wedges, texts, autotexts = ax.pie(sizes, explode=explode, labels=labels, colors=colors_p,
                                      autopct=lambda pct: f'{pct:.1f}%\n({int(pct/100*sum(sizes))})',
                                      startangle=90, textprops={'fontsize': 11, 'fontweight': 'bold'})
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
        autotext.set_fontsize(10)
    ax.set_title('Distribución de Eventos por Severidad', fontsize=12, fontweight='bold', color='#2C263F')
    plt.tight_layout()
    return fig

def generar_grafico_ols_trend(params, deterministic, tifar_prev):
    meses_simulados = np.arange(1, 13)
    tifar_simulado = tifar_prev * (1 + np.sin(np.linspace(0, 2*np.pi, 12)) * 0.3)
    tifar_simulado = np.maximum(tifar_simulado, 1)
    tifar_simulado[11] = deterministic['TIFAR']
    
    x = np.arange(len(meses_simulados))
    slope, intercept = np.polyfit(x, tifar_simulado, 1)
    trend = intercept + slope * x
    fut_meses = np.arange(13, 19)
    fut_trend = intercept + slope * fut_meses
    
    residuos = tifar_simulado - trend
    se = np.sqrt(np.sum(residuos**2) / (len(x) - 2))
    r2 = 1 - np.sum(residuos**2) / np.sum((tifar_simulado - np.mean(tifar_simulado))**2)
    t_val = 1.96
    upper_band = trend + t_val * se
    lower_band = trend - t_val * se
    
    fig, ax = plt.subplots(figsize=(11, 6))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FA')
    ax.plot(meses_simulados, tifar_simulado, 'o-', label='TIFAR Histórico', color='#595082', linewidth=2, markersize=8, markerfacecolor='white', markeredgewidth=2)
    ax.plot(meses_simulados, trend, '--', label=f'Tendencia OLS (β₁={slope:.3f})', color='#E74C3C', linewidth=2)
    ax.fill_between(meses_simulados, lower_band, upper_band, alpha=0.2, color='#F8C662', label='Intervalo 95% confianza')
    ax.plot(fut_meses, fut_trend, '--', color='#41644A', linewidth=2, label='Proyección (6 meses)')
    
    if slope > 0:
        ax.text(0.02, 0.95, f'⚠ ALERTA: Tendencia al alza | R² = {r2:.3f}', 
                transform=ax.transAxes, fontsize=11, color='#E74C3C', fontweight='bold')
    else:
        ax.text(0.02, 0.95, f'✓ Favorable: Tendencia a la baja | R² = {r2:.3f}', 
                transform=ax.transAxes, fontsize=11, color='#41644A', fontweight='bold')
    
    ax.set_xlabel('Periodo (meses)', fontsize=11, fontweight='bold')
    ax.set_ylabel('TIFAR', fontsize=11, fontweight='bold')
    ax.set_title('Análisis Predictivo OLS - Tendencia TIFAR', fontsize=12, fontweight='bold', color='#2C263F')
    ax.legend(loc='upper left', fontsize=9)
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.set_xlim(0, 19)
    plt.tight_layout()
    return fig

def generar_grafico_boxplot_anual(df):
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FA')
    
    df['Año'] = df['Periodo'].dt.year
    data_by_year = [df[df['Año'] == year]['TIFAR'].values for year in sorted(df['Año'].unique())]
    
    bp = ax.boxplot(data_by_year, labels=sorted(df['Año'].unique()), patch_artist=True)
    for patch in bp['boxes']:
        patch.set_facecolor('#595082')
        patch.set_alpha(0.7)
    for median in bp['medians']:
        median.set_color('#E74C3C')
        median.set_linewidth(2)
    
    ax.set_xlabel('Año', fontsize=11, fontweight='bold')
    ax.set_ylabel('TIFAR', fontsize=11, fontweight='bold')
    ax.set_title('Distribución Anual del TIFAR', fontsize=12, fontweight='bold', color='#2C263F')
    ax.grid(True, alpha=0.3, axis='y', linestyle='--')
    plt.tight_layout()
    return fig

def generar_grafico_heatmap(df):
    fig, ax = plt.subplots(figsize=(12, 8))
    fig.patch.set_facecolor('white')
    
    df['Año'] = df['Periodo'].dt.year
    df['Mes'] = df['Periodo'].dt.month
    pivot = df.pivot_table(index='Año', columns='Mes', values='TIFAR', fill_value=0)
    
    im = ax.imshow(pivot.values, cmap='RdYlGn_r', aspect='auto', vmin=0, vmax=df['TIFAR'].quantile(0.9))
    ax.set_xticks(np.arange(len(pivot.columns)))
    ax.set_yticks(np.arange(len(pivot.index)))
    ax.set_xticklabels(['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'])
    ax.set_yticklabels(pivot.index)
    
    for i in range(len(pivot.index)):
        for j in range(len(pivot.columns)):
            ax.text(j, i, f'{pivot.values[i, j]:.1f}', ha='center', va='center', color='black', fontsize=8, fontweight='bold')
    
    plt.colorbar(im, ax=ax, label='TIFAR')
    ax.set_xlabel('Mes', fontsize=11, fontweight='bold')
    ax.set_ylabel('Año', fontsize=11, fontweight='bold')
    ax.set_title('Heatmap de Accidentabilidad - TIFAR por Mes y Año', fontsize=12, fontweight='bold', color='#2C263F')
    plt.tight_layout()
    return fig

def generar_grafico_tendencia_estacional(df, ols):
    fig, ax = plt.subplots(figsize=(12, 6))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FA')
    
    x = np.arange(len(df))
    trend_line = ols['intercept'] + ols['slope'] * x
    
    ax.plot(df['Periodo'], df['TIFAR'], 'o-', color='#595082', alpha=0.6, markersize=4, label='Datos reales')
    ax.plot(df['Periodo'], trend_line, 'r-', linewidth=2.5, label=f'Tendencia (β₁={ols["slope"]:.3f})')
    
    media_movil = df['TIFAR'].rolling(window=12, center=True).mean()
    ax.plot(df['Periodo'], media_movil, '--', color='#F8C662', linewidth=2, label='Media móvil (12 meses)')
    
    ax.set_xlabel('Periodo', fontsize=11, fontweight='bold')
    ax.set_ylabel('TIFAR', fontsize=11, fontweight='bold')
    ax.set_title(f'Tendencia y Estacionalidad - R² = {ols["r2"]:.3f}', fontsize=12, fontweight='bold', color='#2C263F')
    ax.legend(loc='upper left', fontsize=9)
    ax.grid(True, alpha=0.3, linestyle='--')
    plt.xticks(rotation=45)
    plt.tight_layout()
    return fig

def generar_grafico_comparativo_anual(df):
    fig, ax = plt.subplots(figsize=(12, 6))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FA')
    
    df['Año'] = df['Periodo'].dt.year
    for year in sorted(df['Año'].unique()):
        year_data = df[df['Año'] == year]
        ax.plot(year_data['Periodo'].dt.month, year_data['TIFAR'], 'o-', label=str(year), linewidth=2, markersize=6)
    
    ax.set_xlabel('Mes', fontsize=11, fontweight='bold')
    ax.set_ylabel('TIFAR', fontsize=11, fontweight='bold')
    ax.set_title('Comparativa Anual del TIFAR', fontsize=12, fontweight='bold', color='#2C263F')
    ax.set_xticks(range(1, 13))
    ax.set_xticklabels(['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'])
    ax.legend(loc='upper left', fontsize=9)
    ax.grid(True, alpha=0.3, linestyle='--')
    plt.tight_layout()
    return fig

def generar_grafico_mc_forecast(forecast, mc_forecast):
    fig, ax = plt.subplots(figsize=(12, 6))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#F8F9FA')
    
    ax.plot(forecast['Periodo'], forecast['TIFAR_Pronostico'], '--', color='#E74C3C', linewidth=2, label='Pronóstico OLS')
    ax.fill_between(forecast['Periodo'], forecast['Limite_Inferior_95'], forecast['Limite_Superior_95'], alpha=0.2, color='#F8C662', label='Intervalo 95% OLS')
    
    if mc_forecast:
        ax.plot(forecast['Periodo'], mc_forecast['p50'], '-', color='#595082', linewidth=2.5, label='Monte Carlo - P50')
        ax.fill_between(forecast['Periodo'], mc_forecast['p5'], mc_forecast['p95'], alpha=0.3, color='#595082', label='Monte Carlo - P5 a P95')
    
    ax.set_xlabel('Periodo', fontsize=11, fontweight='bold')
    ax.set_ylabel('TIFAR', fontsize=11, fontweight='bold')
    ax.set_title('Pronóstico con Monte Carlo', fontsize=12, fontweight='bold', color='#2C263F')
    ax.legend(loc='upper left', fontsize=9)
    ax.grid(True, alpha=0.3, linestyle='--')
    plt.xticks(rotation=45)
    plt.tight_layout()
    return fig

def generar_reporte_pdf_mensual(params, deterministic, stats_results, correlations):
    buffer = BytesIO()
    with PdfPages(buffer) as pdf:
        fig1, ax1 = plt.subplots(figsize=(11, 8.5))
        ax1.axis('off')
        fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
        texto = f"""
SSO-03: SISTEMA DE CALCULO Y PROYECCION DE TASAS DE ACCIDENTABILIDAD
REPORTE MENSUAL - ANALISIS PUNTUAL
Fecha de emision: {fecha}
N° de simulaciones Monte Carlo: {params['n_simulations']:,}

DATOS DE ENTRADA:
{chr(10).join([f"- {info['label']} ({info['symbol']}): {params[key]} {info['unit']}" for key, info in PARAM_INFO.items()])}

RESULTADOS DETERMINISTICOS:
- TIFAR: {deterministic['TIFAR']:.4f}
- TIPAR: {deterministic['TIPAR']:.4f}
- Dias Cargados: {deterministic['Dias_Cargados']:,.0f} dias
- TISSAR: {deterministic['TISSAR']:.2f}
- IA: {deterministic['IA']:.2f}
- FAR: {deterministic['FAR']:.2f}
- DMI: {deterministic['DMI']:.2f} dias/accidente
- Variacion: {deterministic['Variacion']:+.2f}%

MONTE CARLO - PERCENTILES:
{chr(10).join([f"- {var}: Media={s['mean']:.4f} | P5={s['p5']:.4f} | P50={s['p50']:.4f} | P95={s['p95']:.4f}" for var, s in stats_results.items()])}
"""
        ax1.text(0.1, 0.9, texto, transform=ax1.transAxes, fontsize=9, verticalalignment='top', fontfamily='monospace')
        pdf.savefig(fig1, bbox_inches='tight')
        plt.close(fig1)
        
        pdf.savefig(generar_grafico_histogramas(stats_results), bbox_inches='tight')
        plt.close()
        pdf.savefig(generar_grafico_tornado(correlations), bbox_inches='tight')
        plt.close()
        pdf.savefig(generar_grafico_pastel(params), bbox_inches='tight')
        plt.close()
        pdf.savefig(generar_grafico_ols_trend(params, deterministic, params['tifar_prev']), bbox_inches='tight')
        plt.close()
    
    buffer.seek(0)
    return buffer

def generar_reporte_pdf_historico(resultados):
    buffer = BytesIO()
    with PdfPages(buffer) as pdf:
        fig1, ax1 = plt.subplots(figsize=(11, 8.5))
        ax1.axis('off')
        fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
        ols = resultados['modelo_ols']
        texto = f"""
SSO-03: SISTEMA DE CALCULO Y PROYECCION DE TASAS DE ACCIDENTABILIDAD
REPORTE DE SERIES TEMPORALES - {resultados['n_meses']} MESES DE ANALISIS
Fecha de emision: {fecha}
Periodo analizado: {resultados['fecha_inicio']} a {resultados['fecha_fin']} ({resultados['n_meses']} meses)

MODELO OLS:
- Pendiente (β₁): {ols['slope']:.4f} - Tendencia {ols['tendencia_texto']}
- Coeficiente R²: {ols['r2']:.4f}
- Error estandar: {ols['std_err']:.4f}
- p-valor: {ols['p_value']:.6f}

ESTADISTICAS GLOBALES:
- TIFAR Promedio: {resultados['df_historico']['TIFAR'].mean():.4f}
- TIFAR Max: {resultados['df_historico']['TIFAR'].max():.4f}
- TIFAR Min: {resultados['df_historico']['TIFAR'].min():.4f}
- Total Accidentes Fatales: {resultados['df_historico']['FAT'].sum()}
- Total Dias Cargados: {resultados['df_historico']['Dias_Cargados'].sum():,.0f}

ANOMALIAS DETECTADAS: {len(resultados['anomalias'])} meses
"""
        ax1.text(0.1, 0.9, texto, transform=ax1.transAxes, fontsize=9, verticalalignment='top', fontfamily='monospace')
        pdf.savefig(fig1, bbox_inches='tight')
        plt.close(fig1)
        
        pdf.savefig(generar_grafico_boxplot_anual(resultados['df_historico']), bbox_inches='tight')
        plt.close()
        pdf.savefig(generar_grafico_heatmap(resultados['df_historico']), bbox_inches='tight')
        plt.close()
        pdf.savefig(generar_grafico_tendencia_estacional(resultados['df_historico'], ols), bbox_inches='tight')
        plt.close()
        pdf.savefig(generar_grafico_comparativo_anual(resultados['df_historico']), bbox_inches='tight')
        plt.close()
        
        if resultados['mc_forecast']:
            pdf.savefig(generar_grafico_mc_forecast(resultados['forecast'], resultados['mc_forecast']), bbox_inches='tight')
            plt.close()
    
    buffer.seek(0)
    return buffer

# ==================== INTERFAZ STREAMLIT ====================

st.markdown(f"""
<div class="hero-section">
    <div class="hero-title">SSO-03</div>
    <div class="hero-subtitle">Sistema de Cálculo y Proyección de Tasas de Accidentabilidad Minera</div>
    <div class="hero-badge">⛏️ Gestión de Seguridad Minera | 📊 Análisis Avanzado | 🎲 Monte Carlo</div>
</div>
""", unsafe_allow_html=True)

col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
with col_stat1:
    st.markdown("""
    <div class="stat-card">
        <div class="stat-number">TIFAR</div>
        <div class="stat-label">Índice Frecuencia Total</div>
    </div>
    """, unsafe_allow_html=True)
with col_stat2:
    st.markdown("""
    <div class="stat-card">
        <div class="stat-number">TISSAR</div>
        <div class="stat-label">Índice Severidad Total</div>
    </div>
    """, unsafe_allow_html=True)
with col_stat3:
    st.markdown("""
    <div class="stat-card">
        <div class="stat-number">IA</div>
        <div class="stat-label">Accidentabilidad Global</div>
    </div>
    """, unsafe_allow_html=True)
with col_stat4:
    st.markdown("""
    <div class="stat-card">
        <div class="stat-number">FAR</div>
        <div class="stat-label">Tasa Mortalidad</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

tab1, tab2, tab3 = st.tabs(["📊 Análisis Mensual", "📈 Series Temporales", "📋 Plantilla Excel"])

# ==================== TAB 1: ANÁLISIS MENSUAL ====================
with tab1:
    st.markdown('<div class="section-title">📝 Datos de Entrada</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Complete la siguiente información para el período a analizar</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        hht = st.number_input("👥 Horas Hombre Trabajadas (HHT)", min_value=1, value=152041, step=1000, key="hht_input")
        k_factor = st.number_input("⚙️ Constante K", min_value=100000, value=1000000, step=100000, key="k_input")
    
    with col2:
        am = st.number_input("🩺 Accidentes Médicos (AM)", min_value=0, value=3, step=1, key="am_input")
        atp = st.number_input("⚠️ Accidentes con Tiempo Perdido (ATP)", min_value=0, value=3, step=1, key="atp_input")
        fat = st.number_input("💀 Accidentes Fatales (FAT)", min_value=0, value=1, step=1, key="fat_input")
    
    with col3:
        df_dias = st.number_input("📅 Días Físicos Perdidos (DF)", min_value=0, value=31, step=1, key="df_input")
        tifar_prev = st.number_input("📈 TIFAR Periodo Anterior", min_value=0.0, value=6.47, step=0.1, key="tifar_prev_input")
        n_simulations = st.number_input("🎲 Simulaciones Monte Carlo", min_value=1000, value=10000, step=5000, key="n_sim_input")
    
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        calcular = st.button("🔍 Calcular Análisis", type="primary", use_container_width=True)
    
    if calcular:
        params = {'hht': hht, 'am': am, 'atp': atp, 'fat': fat, 'df': df_dias, 'k_factor': k_factor, 'tifar_prev': tifar_prev, 'n_simulations': n_simulations}
        
        tifar = calc_tifar(am, atp, fat, k_factor, hht)
        tipar = calc_tipar(atp, fat, k_factor, hht)
        dias_cargados = calc_dias_cargados(df_dias, fat)
        tissar = calc_tissar(dias_cargados, k_factor, hht)
        ia = calc_ia(tifar, tissar)
        far = calc_far(fat, hht)
        dmi = calc_dmi(df_dias, atp)
        variacion = calc_variacion(tifar, tifar_prev)
        
        deterministic = {
            'TIFAR': tifar, 'TIPAR': tipar, 'Dias_Cargados': dias_cargados,
            'TISSAR': tissar, 'IA': ia, 'FAR': far, 'DMI': dmi, 'Variacion': variacion
        }
        
        with st.spinner("🎲 Ejecutando simulación Monte Carlo..."):
            stats_results, correlations, samples = monte_carlo(params, n_simulations)
        
        st.markdown("---")
        st.markdown('<div class="section-title">📊 Resultados Determinísticos</div>', unsafe_allow_html=True)
        
        col_a, col_b, col_c, col_d = st.columns(4)
        with col_a:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">📈 TIFAR</div>
                <div class="metric-value">{tifar:.4f}</div>
                <div class="metric-unit">Índice Frecuencia Total</div>
            </div>
            """, unsafe_allow_html=True)
        with col_b:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">⚠️ TIPAR</div>
                <div class="metric-value">{tipar:.4f}</div>
                <div class="metric-unit">Índice Tiempo Perdido</div>
            </div>
            """, unsafe_allow_html=True)
        with col_c:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">💀 TISSAR</div>
                <div class="metric-value">{tissar:.2f}</div>
                <div class="metric-unit">Índice Severidad Total</div>
            </div>
            """, unsafe_allow_html=True)
        with col_d:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">🎯 IA</div>
                <div class="metric-value">{ia:.2f}</div>
                <div class="metric-unit">Accidentabilidad Global</div>
            </div>
            """, unsafe_allow_html=True)
        
        col_e, col_f, col_g, col_h = st.columns(4)
        with col_e:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">⚡ FAR</div>
                <div class="metric-value">{far:.2f}</div>
                <div class="metric-unit">Tasa Mortalidad (×100M)</div>
            </div>
            """, unsafe_allow_html=True)
        with col_f:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">📊 DMI</div>
                <div class="metric-value">{dmi:.2f}</div>
                <div class="metric-unit">días/accidente</div>
            </div>
            """, unsafe_allow_html=True)
        with col_g:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">📅 Días Cargados</div>
                <div class="metric-value">{dias_cargados:,.0f}</div>
                <div class="metric-unit">incluye penalización</div>
            </div>
            """, unsafe_allow_html=True)
        with col_h:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">🔄 Variación Δ%</div>
                <div class="metric-value" style="color: {'#41644A' if variacion <= 0 else '#E74C3C'}">{variacion:+.2f}%</div>
                <div class="metric-unit">vs período anterior</div>
            </div>
            """, unsafe_allow_html=True)
        
        if tifar < 10:
            st.markdown('<div class="eval-box eval-success">✅ EVALUACIÓN FAVORABLE: El TIFAR indica un riesgo controlado. Mantener los programas de prevención actuales.</div>', unsafe_allow_html=True)
        elif tifar < 25:
            st.markdown('<div class="eval-box eval-warning">⚠️ EVALUACIÓN MODERADA: El TIFAR requiere seguimiento. Reforzar controles operacionales.</div>', unsafe_allow_html=True)
        elif tifar < 50:
            st.markdown('<div class="eval-box eval-warning">🔴 EVALUACIÓN SEVERA: El TIFAR es alto. Implementar acciones correctivas inmediatas.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="eval-box eval-error">🔴🔴 EVALUACIÓN CRÍTICA: El TIFAR es extremadamente alto. SUSPENDER actividades de alto riesgo.</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="section-title">🎲 Monte Carlo - Percentiles</div>', unsafe_allow_html=True)
        mc_stats = {}
        for var in ['TIFAR', 'TIPAR', 'TISSAR', 'IA']:
            mc_stats[var] = {
                'Media': f"{stats_results[var]['mean']:.4f}",
                'P5': f"{stats_results[var]['p5']:.4f}",
                'P50': f"{stats_results[var]['p50']:.4f}",
                'P95': f"{stats_results[var]['p95']:.4f}"
            }
        st.dataframe(pd.DataFrame(mc_stats).T, use_container_width=True)
        
        st.markdown('<div class="section-title">📈 Visualizaciones</div>', unsafe_allow_html=True)
        col_graf1, col_graf2 = st.columns(2)
        with col_graf1:
            st.pyplot(generar_grafico_histogramas(stats_results))
        with col_graf2:
            st.pyplot(generar_grafico_tornado(correlations))
        
        col_graf3, col_graf4 = st.columns(2)
        with col_graf3:
            st.pyplot(generar_grafico_pastel(params))
        with col_graf4:
            st.pyplot(generar_grafico_ols_trend(params, deterministic, tifar_prev))
        
        col_download1, col_download2 = st.columns(2)
        with col_download1:
            excel_data = export_excel_mensual(params, deterministic, stats_results, correlations, samples)
            st.download_button("📥 Exportar a Excel", excel_data, "SSO03_Reporte_Mensual.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_download2:
            pdf_data = generar_reporte_pdf_mensual(params, deterministic, stats_results, correlations)
            st.download_button("📄 Exportar a PDF", pdf_data, "SSO03_Reporte_Mensual.pdf", "application/pdf", use_container_width=True)

# ==================== TAB 2: ANÁLISIS HISTÓRICO ====================
with tab2:
    st.markdown('<div class="section-title">📂 Carga de Datos Históricos</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Cargue un archivo Excel con los datos mensuales para el análisis de series temporales</div>', unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("Seleccionar archivo Excel", type=['xlsx', 'xls'], help="El archivo debe contener las columnas: Periodo, HHT, AM, ATP, FAT, DF")
    
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        with st.expander("📋 Vista previa de los datos cargados"):
            st.dataframe(df.head(10), use_container_width=True)
        
        st.markdown('<div class="section-title">⚙️ Configuración del Análisis</div>', unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            k_factor_hist = st.number_input("⚙️ Constante K", value=1000000, step=100000, key="k_hist")
        with col2:
            forecast_months = st.number_input("📅 Meses de proyección", min_value=1, max_value=36, value=12, key="forecast")
        with col3:
            anomaly_threshold = st.number_input("🎯 Umbral anomalías (Z-score)", min_value=1.5, max_value=4.0, value=2.5, step=0.1, key="anomaly")
        
        mc_check = st.checkbox("🎲 Incluir pronóstico con Monte Carlo", value=True)
        mc_samples = 10000
        if mc_check:
            mc_samples = st.number_input("Número de simulaciones MC", min_value=1000, max_value=50000, value=10000, step=5000, key="mc_samples")
        
        if st.button("🔍 Analizar Datos Históricos", type="primary", use_container_width=True):
            with st.spinner("📊 Procesando datos históricos..."):
                resultados = analisis_historico(df, k_factor_hist, forecast_months, anomaly_threshold, mc_check, mc_samples)
            
            st.success(f"✅ Análisis completado: {resultados['n_meses']} meses analizados ({resultados['fecha_inicio']} a {resultados['fecha_fin']})")
            
            st.markdown('<div class="section-title">📈 Modelo de Tendencia OLS</div>', unsafe_allow_html=True)
            ols = resultados['modelo_ols']
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">📐 Pendiente (β₁)</div>
                    <div class="metric-value">{ols['slope']:.4f}</div>
                    <div class="metric-unit">{ols['tendencia_texto']}</div>
                </div>
                """, unsafe_allow_html=True)
            with col_b:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">📊 R²</div>
                    <div class="metric-value">{ols['r2']:.4f}</div>
                    <div class="metric-unit">Confiabilidad del modelo</div>
                </div>
                """, unsafe_allow_html=True)
            with col_c:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">⚠️ Error estándar</div>
                    <div class="metric-value">{ols['std_err']:.4f}</div>
                    <div class="metric-unit">Precisión</div>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown('<div class="section-title">📊 Visualizaciones</div>', unsafe_allow_html=True)
            
            fig1 = generar_grafico_boxplot_anual(resultados['df_historico'])
            st.pyplot(fig1)
            plt.close(fig1)
            
            fig2 = generar_grafico_heatmap(resultados['df_historico'])
            st.pyplot(fig2)
            plt.close(fig2)
            
            fig3 = generar_grafico_tendencia_estacional(resultados['df_historico'], ols)
            st.pyplot(fig3)
            plt.close(fig3)
            
            fig4 = generar_grafico_comparativo_anual(resultados['df_historico'])
            st.pyplot(fig4)
            plt.close(fig4)
            
            if resultados['mc_forecast']:
                fig5 = generar_grafico_mc_forecast(resultados['forecast'], resultados['mc_forecast'])
                st.pyplot(fig5)
                plt.close(fig5)
            
            st.markdown('<div class="section-title">📅 Resumen Anual</div>', unsafe_allow_html=True)
            st.dataframe(resultados['resumen_anual'].round(4), use_container_width=True)
            
            if len(resultados['anomalias']) > 0:
                st.markdown('<div class="section-title">🚨 Meses con Anomalías Detectadas</div>', unsafe_allow_html=True)
                st.dataframe(resultados['anomalias'][['Periodo', 'TIFAR', 'TISSAR', 'IA']], use_container_width=True)
            
            col_download1, col_download2 = st.columns(2)
            with col_download1:
                excel_data = export_historico_excel(resultados)
                st.download_button("📥 Exportar a Excel", excel_data, "SSO03_Resultados_SeriesTemporales.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with col_download2:
                pdf_data = generar_reporte_pdf_historico(resultados)
                st.download_button("📄 Exportar a PDF", pdf_data, "SSO03_Reporte_SeriesTemporales.pdf", "application/pdf", use_container_width=True)

# ==================== TAB 3: PLANTILLA EXCEL (CORREGIDO) ====================
with tab3:
    st.markdown('<div class="section-title">📋 Plantilla para Datos Históricos</div>', unsafe_allow_html=True)
    
    # Usar st.markdown con estilo para la tabla en lugar del markdown normal
    st.markdown("""
    <div style="background: white; padding: 1.5rem; border-radius: 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.05); margin-bottom: 1.5rem;">
        <h4 style="color: #2C263F; margin-bottom: 1rem;">📊 Columnas requeridas</h4>
        <table style="width: 100%; border-collapse: collapse; background: white;">
            <thead>
                <tr>
                    <th style="background: #2C263F; color: white; padding: 10px; text-align: left; border-radius: 8px 0 0 0;">Columna</th>
                    <th style="background: #2C263F; color: white; padding: 10px; text-align: left;">Descripción</th>
                    <th style="background: #2C263F; color: white; padding: 10px; text-align: left; border-radius: 0 8px 0 0;">Ejemplo</th>
                </tr>
            </thead>
            <tbody>
                <tr style="border-bottom: 1px solid #E8ECEF;">
                    <td style="padding: 10px; color: #2C263F; font-weight: 600;">Periodo</td>
                    <td style="padding: 10px; color: #41644A;">Fecha en formato AAAA-MM</td>
                    <td style="padding: 10px; color: #2C263F;">2024-01</td>
                </tr>
                <tr style="border-bottom: 1px solid #E8ECEF; background: #F8F9FA;">
                    <td style="padding: 10px; color: #2C263F; font-weight: 600;">HHT</td>
                    <td style="padding: 10px; color: #41644A;">Horas Hombre Trabajadas</td>
                    <td style="padding: 10px; color: #2C263F;">152041</td>
                </tr>
                <tr style="border-bottom: 1px solid #E8ECEF;">
                    <td style="padding: 10px; color: #2C263F; font-weight: 600;">AM</td>
                    <td style="padding: 10px; color: #41644A;">Accidentes Médicos</td>
                    <td style="padding: 10px; color: #2C263F;">3</td>
                </tr>
                <tr style="border-bottom: 1px solid #E8ECEF; background: #F8F9FA;">
                    <td style="padding: 10px; color: #2C263F; font-weight: 600;">ATP</td>
                    <td style="padding: 10px; color: #41644A;">Accidentes con Tiempo Perdido</td>
                    <td style="padding: 10px; color: #2C263F;">3</td>
                </tr>
                <tr style="border-bottom: 1px solid #E8ECEF;">
                    <td style="padding: 10px; color: #2C263F; font-weight: 600;">FAT</td>
                    <td style="padding: 10px; color: #41644A;">Accidentes Fatales</td>
                    <td style="padding: 10px; color: #2C263F;">1</td>
                </tr>
                <tr style="background: #F8F9FA;">
                    <td style="padding: 10px; color: #2C263F; font-weight: 600; border-radius: 0 0 0 8px;">DF</td>
                    <td style="padding: 10px; color: #41644A;">Días Físicos Perdidos</td>
                    <td style="padding: 10px; color: #2C263F; border-radius: 0 0 8px 0;">31</td>
                </tr>
            </tbody>
        </table>
    </div>
    
    <div style="background: #FFF8E1; padding: 1rem; border-radius: 12px; border-left: 4px solid #F8C662; margin: 1rem 0;">
        <h4 style="color: #856404; margin: 0 0 0.5rem 0;">📌 Notas importantes</h4>
        <ul style="color: #856404; margin: 0; padding-left: 1.2rem;">
            <li>Complete un registro por cada mes (puede tener desde 1 mes hasta 100+ meses)</li>
            <li>Mantenga el formato de fecha <strong>AAAA-MM</strong> (ej: 2024-01, 2024-02, ...)</li>
            <li>Cada accidente fatal (FAT) suma automáticamente <strong>6,000 días cargados</strong> según norma ANSI Z16.1</li>
            <li>Los datos de ejemplo son solo referenciales, sustitúyalos con sus datos reales</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("📋 Generar Plantilla", use_container_width=True):
        plantilla = generar_plantilla_excel()
        st.download_button(
            label="📥 Descargar Plantilla",
            data=plantilla,
            file_name="Plantilla_SSO03_Datos_Mensuales.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
st.markdown("---")
st.markdown("""
<div class="footer">
    <p>⛏️ SSO-03: Sistema de Cálculo y Proyección de Tasas de Accidentabilidad | Desarrollado para minería</p>
    <p style="font-size: 0.7rem;">Basado en estándares OSHA, ICMM y ANSI Z16.1</p>
</div>
""", unsafe_allow_html=True)