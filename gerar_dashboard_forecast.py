import pandas as pd
import os
from datetime import datetime
import json
import glob
import re
import shutil
import sys
import time
import warnings
import webbrowser
import html
import traceback
import requests
from dotenv import load_dotenv

# --- VERIFICAÇÃO DE AMBIENTE VIRTUAL ---
if sys.prefix == sys.base_prefix:
    print("\n" + "="*80)
    print("❌ ERRO CRÍTICO: O script não está rodando no ambiente virtual (venv).")
    print("   A execução foi interrompida para garantir a integridade das dependências.")
    print("\n   COMO CORRIGIR:")
    print("   1. Abra um terminal (PowerShell) na pasta do projeto.")
    print("   2. Ative o ambiente virtual: > .\\venv\\Scripts\\Activate.ps1")
    print("   3. Rode novamente: > python gerar_dashboard_forecast.py")
    print("="*80 + "\n")
    sys.exit(1)

load_dotenv(override=True)
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

# --- TRANSLATIONS (PT / EN / ES) ---
TRANSLATIONS = {
    "pt": {
        "realized": "Realizado", "var_cycle": "Var. Ciclo", "seg": "🚜 Segmento", "total": "Total:",
        "proj_vol": "Projeção de Volumes", "export": "Exportar para Excel",
        "version": "Versão", "updated": "Atualizado",
        "kpi_vol_title": "Volume Total Projetado", "kpi_var_title": "Variação de Volume",
        "select_prod": "Selecione a Família de Produto:", "month": "Mês", "current_forecast": "Atual Forecast",
        "positive": "Positivo", "critical": "Crítico", "warning": "Atenção", "var_yoy": "Var. YoY",
        "positivo": "Positivo", "negativo": "Negativo", "neutro": "Neutro", "incerto": "Incerto",
        "insights_agribusiness_title": "Insights de Agribusiness (IA)",
        "jan": "Janeiro", "feb": "Fevereiro", "mar": "Março", "apr": "Abril", "may": "Maio", "jun": "Junho", "jul": "Julho", "aug": "Agosto", "sep": "Setembro", "oct": "Outubro", "nov": "Novembro", "dec": "Dezembro", "current": "Atual",
        "scenario_general": "Cenário Geral do Mercado", "relevant_news": "Notícias Relevantes",
        "brasil": "Brasil", "argentina": "Argentina", "mexico": "México", "osa": "OSA",
        "ta": "Tratores", "co": "Colheitadeiras", "pa": "Plantadeiras", "pu": "Pulverizadores",
        "metodologia": "Metodologia",
        "met_title": "Metodologia do Forecast", "met_badge": "Modelo preditivo",
        "met_detalhamento_title": "Detalhamento por Região, Produto e Segmento",
        "met_col_seg": "Segmento", "met_col_vars": "Variáveis do modelo (features do XGBoost)", "met_col_n": "Nº",
        "met_legend_title": "Natureza das variáveis:", "met_temporal": "Temporal", "met_agro": "Agronômica / Setorial", "met_macro": "Macroeconômica & Preços",
    },
    "en": {
        "realized": "Realized", "var_cycle": "Cycle Var.", "seg": "🚜 Segment", "total": "Total:",
        "proj_vol": "Volume Projection", "export": "Export to Excel",
        "version": "Version", "updated": "Updated",
        "kpi_vol_title": "Total Projected Volume", "kpi_var_title": "Volume Variation",
        "select_prod": "Select Product Family:", "month": "Month", "current_forecast": "Current Forecast",
        "positive": "Positive", "critical": "Critical", "warning": "Warning", "var_yoy": "YoY Var.",
        "positivo": "Positive", "negativo": "Negative", "neutro": "Neutral", "incerto": "Uncertain",
        "insights_agribusiness_title": "Agribusiness Insights (AI)",
        "jan": "January", "feb": "February", "mar": "March", "apr": "April", "may": "May", "jun": "June", "jul": "July", "aug": "August", "sep": "September", "oct": "October", "nov": "November", "dec": "December", "current": "Current",
        "scenario_general": "General Market Scenario", "relevant_news": "Relevant News",
        "brasil": "Brazil", "argentina": "Argentina", "mexico": "Mexico", "osa": "OSA",
        "ta": "Tractors", "co": "Combines", "pa": "Planters", "pu": "Sprayers",
        "metodologia": "Methodology",
        "met_title": "Forecast Methodology", "met_badge": "Predictive model",
        "met_detalhamento_title": "Breakdown by Region, Product and Segment",
        "met_col_seg": "Segment", "met_col_vars": "Model variables (XGBoost features)", "met_col_n": "No.",
        "met_legend_title": "Nature of variables:", "met_temporal": "Temporal", "met_agro": "Agronomic / Sector", "met_macro": "Macroeconomic & Prices",
    },
    "es": {
        "realized": "Realizado", "var_cycle": "Var. Ciclo", "seg": "🚜 Segmento", "total": "Total:",
        "proj_vol": "Proyección de Volúmenes", "export": "Exportar a Excel",
        "version": "Versión", "updated": "Actualizado",
        "kpi_vol_title": "Volumen Total Proyectado", "kpi_var_title": "Variación de Volumen",
        "select_prod": "Seleccione Familia de Producto:", "month": "Mes", "current_forecast": "Forecast Actual",
        "positive": "Positivo", "critical": "Crítico", "warning": "Atención", "var_yoy": "Var. YoY",
        "positivo": "Positivo", "negativo": "Negativo", "neutro": "Neutro", "incerto": "Incierto",
        "insights_agribusiness_title": "Insights de Agronegocios (IA)",
        "jan": "Enero", "feb": "Febrero", "mar": "Marzo", "apr": "Abril", "may": "Mayo", "jun": "Junio", "jul": "Julio", "aug": "Agosto", "sep": "Septiembre", "oct": "Octubre", "nov": "Noviembre", "dec": "Diciembre", "current": "Actual",
        "scenario_general": "Escenario General del Mercado", "relevant_news": "Noticias Relevantes",
        "brasil": "Brasil", "argentina": "Argentina", "mexico": "México", "osa": "OSA",
        "ta": "Tractores", "co": "Cosechadoras", "pa": "Sembradoras", "pu": "Pulverizadores",
        "metodologia": "Metodología",
        "met_title": "Metodología del Forecast", "met_badge": "Modelo predictivo",
        "met_detalhamento_title": "Detalle por Región, Producto y Segmento",
        "met_col_seg": "Segmento", "met_col_vars": "Variables del modelo (features de XGBoost)", "met_col_n": "Nº",
        "met_legend_title": "Naturaleza de las variables:", "met_temporal": "Temporal", "met_agro": "Agronómica / Sectorial", "met_macro": "Macroeconómica y Precios",
    }
}

def i18n(key, lang='pt'):
    return TRANSLATIONS.get(lang, TRANSLATIONS['pt']).get(key, key)

def ml(pt, en, es, tag=None, cls=""):
    """Texto multilíngue no padrão do template (.i18n-text lang-XX)."""
    inner = (f'<span class="i18n-text lang-pt">{html.escape(pt)}</span>'
             f'<span class="i18n-text lang-en">{html.escape(en)}</span>'
             f'<span class="i18n-text lang-es">{html.escape(es)}</span>')
    if tag:
        c = f' class="{cls}"' if cls else ""
        return f'<{tag}{c}>{inner}</{tag}>'
    return inner

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAMINHO_HTML = os.path.join(BASE_DIR, 'index.html')

COLUMN_MAPPING_CONFIG = {
    'col_m': ['COUNTRY', 'PAIS', 'PAÍS', 'MERCADO', 'MARKET', 'GEOGRAFIA', 'REGION'],
    'col_p': ['PRODUCT', 'PRODUTO', 'FAMILY', 'FAMILIA', 'EQUIPMENT'],
    'col_s': ['SEGMENT', 'SEGMENTO', 'POTENCIA', 'HP', 'CLASS', 'CLASSE', 'GROUP'],
    'col_y': ['YEAR', 'ANO', 'DATA', 'DATE', 'MÊS', 'MES', 'PERIOD', 'TIME'],
    'col_v': ['IND', 'VOL', 'QTY', 'QUANT', 'FORECAST', 'BASE', 'VALUE', 'VALOR', 'UNIDADES']
}

def validar_e_mapear_colunas(df):
    mapped_cols = {key: None for key in COLUMN_MAPPING_CONFIG}
    for col in df.columns:
        col_up = str(col).upper()
        for key, keywords in COLUMN_MAPPING_CONFIG.items():
            if not mapped_cols[key] and any(kw in col_up for kw in keywords):
                mapped_cols[key] = col
                break
    if not mapped_cols['col_v']:
        for col in df.select_dtypes(include=['number']).columns:
            if col != mapped_cols.get('col_y'):
                mapped_cols['col_v'] = col
                break
    if not mapped_cols['col_v']:
        mapped_cols['col_v'] = df.columns[-1]
    return mapped_cols['col_m'], mapped_cols['col_p'], mapped_cols['col_s'], mapped_cols['col_y'], mapped_cols['col_v']

def ordenar_segmentos(segmentos):
    def segment_key(seg):
        s = str(seg).upper()
        if 'VIII' in s: return 8000
        if 'VII' in s: return 7000
        if 'VI' in s: return 6000
        if 'IV' in s: return 4000
        if 'V' in s: return 5000
        match = re.search(r'\d+', s)
        if match: return int(match.group())
        if s == 'ALL': return -1
        return 9999
    return sorted(segmentos, key=segment_key)

def obter_insights_agribusiness():
    early_signals_dir = os.path.join(BASE_DIR, '..', 'New Early signals')
    cache_pattern = os.path.join(early_signals_dir, 'cache_dados_ia_*.json')
    cache_files = sorted(glob.glob(cache_pattern), reverse=True)
    insights_path = None
    if cache_files:
        insights_path = cache_files[0]
        print(f"\nℹ️  Buscando insights do cache mais recente: {os.path.basename(insights_path)}")
    else:
        fallback_path = os.path.join(early_signals_dir, 'dados_paises.json')
        if os.path.exists(fallback_path):
            insights_path = fallback_path
            print(f"\nℹ️  Usando 'dados_paises.json' do 'New Early signals'.")
        else:
            print(f"⚠️ AVISO: Nenhum arquivo de insights encontrado no projeto 'New Early signals'.")
            return {}
    try:
        with open(insights_path, 'r', encoding='utf-8') as f:
            dados_externos = json.load(f)
        print("✅ Insights de agribusiness carregados com sucesso.")
        insights_por_pais = {}
        country_code_map = {'BR': 'BRASIL', 'AR': 'ARGENTINA', 'MX': 'MEXICO', 'CL': 'CHILE', 'PY': 'PARAGUAI', 'UY': 'URUGUAI', 'PE': 'PERU', 'BO': 'BOLIVIA', 'CO': 'COLOMBIA', 'EC': 'EQUADOR'}
        main_countries = ['BRASIL', 'ARGENTINA', 'MEXICO']
        osa_countries = [v for v in country_code_map.values() if v not in main_countries]
        osa_fatores, osa_noticias = [], []
        for pais_code, dados_pais in dados_externos.items():
            pais_full_name = country_code_map.get(pais_code.upper(), pais_code.upper())
            if pais_full_name in main_countries:
                insights_por_pais[pais_full_name] = dados_pais
            elif pais_full_name in osa_countries:
                if dados_pais.get('fatores_economicos'):
                    for fator in dados_pais['fatores_economicos']:
                        fator['fonte_pais'] = pais_full_name
                        osa_fatores.append(fator)
                if dados_pais.get('noticias'):
                    for noticia in dados_pais['noticias']:
                        noticia['fonte_pais'] = pais_full_name
                        osa_noticias.append(noticia)
        if osa_fatores or osa_noticias:
            insights_por_pais['OSA'] = {'fatores_economicos': osa_fatores, 'noticias': sorted(osa_noticias, key=lambda x: x.get('tendencia_noticia', 'neutro') != 'negativo')}
        return insights_por_pais
    except (json.JSONDecodeError, Exception) as e:
        print(f"❌ ERRO ao ler insights: {e}")
        return {}

def gerar_html_insights_agribusiness(insights_mercado, mes_nome_ano, mes_key):
    if not insights_mercado:
        return ""
    fatores = insights_mercado.get('fatores_economicos', [])
    noticias = insights_mercado.get('noticias', [])
    if not fatores and not noticias:
        return ""
    html_resumo = ""
    if not fatores:
        html_resumo = "<p>—</p>"
    else:
        resumo_parts = []
        sentiment_map = {'positivo': 'positive', 'baixa': 'positive', 'expansiva': 'positive', 'negativo': 'negative', 'alta': 'negative', 'restritiva': 'negative'}
        sentiment_text_map = {'positivo': 'positivo', 'baixa': 'positivo', 'expansiva': 'positivo', 'negativo': 'negativo', 'alta': 'negativo', 'restritiva': 'negativo', 'incerto': 'incerto', 'estavel': 'neutro'}
        for fator in fatores:
            descricao_obj = fator.get('descricao', {})
            desc_pt = descricao_obj.get('pt', '') if isinstance(descricao_obj, dict) else str(descricao_obj)
            desc_en = descricao_obj.get('en', desc_pt) if isinstance(descricao_obj, dict) else str(descricao_obj)
            desc_es = descricao_obj.get('es', desc_pt) if isinstance(descricao_obj, dict) else str(descricao_obj)
            tendencia = fator.get('tendencia', 'incerto').lower()
            i18n_key = sentiment_text_map.get(tendencia, 'incerto')
            sentiment_class = sentiment_map.get(tendencia, 'neutral')
            badge_html = f'<span class="summary-sentiment-badge {sentiment_class}" data-i18n="{i18n_key}">{i18n(i18n_key)}</span>'
            desc_span = (f'<span class="i18n-text lang-pt">{html.escape(desc_pt)}</span>'
                         f'<span class="i18n-text lang-en">{html.escape(desc_en)}</span>'
                         f'<span class="i18n-text lang-es">{html.escape(desc_es)}</span>')
            resumo_parts.append(f'<div class="summary-item {sentiment_class}"><span class="summary-item-icon">{fator.get("icone", "📊")}</span><div class="summary-item-text"><strong>{fator.get("titulo", "")}:</strong> {desc_span}</div>{badge_html}</div>')
        html_resumo = "".join(resumo_parts)
    html_noticias = ""
    if not noticias:
        html_noticias = "<p>—</p>"
    else:
        for noticia in noticias:
            titulo_obj = noticia.get('titulo_noticia', {})
            titulo_pt = titulo_obj.get('pt', 'Sem título') if isinstance(titulo_obj, dict) else str(titulo_obj)
            titulo_en = titulo_obj.get('en', titulo_pt) if isinstance(titulo_obj, dict) else str(titulo_obj)
            titulo_es = titulo_obj.get('es', titulo_pt) if isinstance(titulo_obj, dict) else str(titulo_obj)
            resumo_obj = noticia.get('corpo_noticia', {})
            resumo_pt = resumo_obj.get('pt', 'Sem resumo.') if isinstance(resumo_obj, dict) else str(resumo_obj)
            resumo_en = resumo_obj.get('en', resumo_pt) if isinstance(resumo_obj, dict) else str(resumo_obj)
            resumo_es = resumo_obj.get('es', resumo_pt) if isinstance(resumo_obj, dict) else str(resumo_obj)
            sentimento = noticia.get('tendencia_noticia', 'neutro').lower()
            card_class = {'positivo': 'positive-card', 'negativo': 'negative-card'}.get(sentimento, 'neutral-card')
            cor_sentimento = {'positivo': 'var(--pos-color)', 'negativo': 'var(--neg-color)', 'neutro': 'var(--neutral-color)'}.get(sentimento, 'var(--neutral-color)')
            sentimento_texto = i18n(sentimento)
            fonte_pais_html = f'<span class="news-source-country">{noticia["fonte_pais"]}</span>' if 'fonte_pais' in noticia else ""
            titulo_ml = (f'<span class="i18n-text lang-pt">{html.escape(titulo_pt)}</span>'
                         f'<span class="i18n-text lang-en">{html.escape(titulo_en)}</span>'
                         f'<span class="i18n-text lang-es">{html.escape(titulo_es)}</span>')
            resumo_ml = (f'<span class="i18n-text lang-pt">{html.escape(resumo_pt)}</span>'
                         f'<span class="i18n-text lang-en">{html.escape(resumo_en)}</span>'
                         f'<span class="i18n-text lang-es">{html.escape(resumo_es)}</span>')
            html_noticias += f"""
            <div class="news-card {card_class}">
                <div class="news-card-header">
                    <span class="news-sentiment-dot" style="background-color: {cor_sentimento};"></span>
                    <span class="news-sentiment-label" style="color: {cor_sentimento};" data-i18n="{sentimento}">{sentimento_texto}</span>
                    <h5 class="news-title">{titulo_ml}</h5>
                </div>
                <p class="news-summary">{resumo_ml} {fonte_pais_html}</p>
            </div>"""
    return f"""
    <div class="insights-section">
        <hr class="agco-divider">
        <h3 class="section-title" data-i18n="insights_agribusiness_title">{i18n("insights_agribusiness_title")}
            <span class="month-badge" data-i18n="{mes_key}">{mes_nome_ano}</span></h3>
        <div class="insights-container">
            <div class="insight-summary-card"><h4 data-i18n="scenario_general">{i18n("scenario_general")}</h4>
                <div class="summary-content">{html_resumo}</div></div>
            <div class="insight-news-container"><h4 data-i18n="relevant_news">{i18n("relevant_news")}</h4>
                <div class="news-grid">{html_noticias}</div></div>
        </div>
    </div>"""

def limpar_historico_antigo(hist_dir, manter=10):
    try:
        padrao = os.path.join(hist_dir, "Forecast_*+*.xlsx")
        arquivos = sorted(glob.glob(padrao), key=os.path.getmtime, reverse=True)
        if len(arquivos) > manter:
            for arquivo_para_deletar in arquivos[manter:]:
                os.remove(arquivo_para_deletar)
    except Exception:
        pass

def carregar_bases():
    hist_dir = os.path.join(BASE_DIR, 'historico_forecast')
    os.makedirs(hist_dir, exist_ok=True)
    arquivos_validos = []
    for f in glob.glob(os.path.join(BASE_DIR, "*.xlsx")):
        nome = os.path.basename(f).lower()
        if not nome.startswith("~$") and ("forecast" in nome or "cenario" in nome):
            arquivos_validos.append(f)
    if not arquivos_validos: return {}
    caminho_base_atual = max(arquivos_validos, key=os.path.getmtime)
    match = re.search(r'(\d+)\s*\+\s*(\d+)', os.path.basename(caminho_base_atual))
    if match:
        versao_atual = f"{match.group(1)}+{match.group(2)}"
    else:
        now = datetime.now()
        actuals = now.month - 1 if now.month > 1 else 12
        versao_atual = f"{actuals}+{12 - actuals}"
    caminho_hist = os.path.join(hist_dir, f'Forecast_{versao_atual}.xlsx')
    try:
        shutil.copy2(caminho_base_atual, caminho_hist)
    except PermissionError:
        print("\n❌ FECHE o arquivo no Excel e rode o script novamente.\n")
        return {}
    padrao = os.path.join(hist_dir, "Forecast_*+*.xlsx")
    arquivos = sorted(glob.glob(padrao), key=os.path.getmtime)
    bases_temp = {}
    for arq in arquivos:
        m = re.search(r'Forecast_(\d+\+\d+)\.xlsx', os.path.basename(arq))
        versao = m.group(1) if m else os.path.basename(arq)
        bases_temp[versao] = arq
    todas = list(bases_temp.keys())
    ultimas = []
    if todas:
        recente = todas[-1]
        if recente == '7+5':
            print("\nℹ️  REGRA ESPECIAL: Detectada versão '7+5'. Comparando com '6+6'.")
            sys.stdout.flush()
            if '6+6' in todas:
                ultimas = ['6+6', '7+5']
            else:
                print("   ⚠️ AVISO: Versão '6+6' não encontrada. Usando a versão anterior disponível.")
                ultimas = todas[-2:] if len(todas) > 1 else todas[-1:]
        else:
            ultimas = todas[-2:] if len(todas) > 1 else todas[-1:]
    bases = {v: pd.read_excel(bases_temp[v]) for v in ultimas}
    limpar_historico_antigo(hist_dir)
    return bases

def renderizar_template_final(context):
    template_path = os.path.join(BASE_DIR, 'template.html')
    if not os.path.exists(template_path):
        print("❌ ERRO: 'template.html' não encontrado na pasta.")
        return
    with open(template_path, 'r', encoding='utf-8') as f:
        html_completo = f.read()
    for key, value in context.items():
        if key == 'injected_css': continue
        html_completo = html_completo.replace(f'{{{key}}}', str(value))
    if context.get('injected_css'):
        html_completo = html_completo.replace('</head>', f'<style>{context["injected_css"]}</style>\n</head>')
    with open(CAMINHO_HTML, 'w', encoding='utf-8') as f:
        f.write(html_completo)

def gerar_conteudo_produto(produto, df_segmentos, p_idx, mercado, mercado_limpo, anos_unicos, versoes, current_year, mes_key, cols):
    col_m, col_p, col_s, col_y, col_v = cols
    versao_rec = versoes[-1]
    versao_ant = versoes[0] if len(versoes) > 1 else None
    produto_limpo = str(produto).replace(" ", "_").upper()
    prod_id = f"{mercado_limpo}_{produto_limpo}"
    p_active_class = "active" if p_idx == 0 else ""
    p_display_style = "block" if p_idx == 0 else "none"
    produto_key_i18n = str(produto).lower().replace(' ', '_')
    html_prod_tab = f'<button class="prod-tab-btn {p_active_class}" onclick="openProduct(event, \'{prod_id}\', \'{mercado_limpo}\')" data-i18n="{produto_key_i18n}">{i18n(produto_key_i18n)}</button>'
    segmentos_unicos = ordenar_segmentos(df_segmentos[col_s].unique())
    thb = 'border-bottom: 1px solid #e2e8f0; color: #1e293b; font-size: 13px; font-weight: 600;'
    ths = 'font-size: 11px; color: #64748b; font-weight: 500;'
    thsb = 'font-size: 11px; color: #0f172a; font-weight: 600;'
    if versao_ant:
        h1 = f'<th rowspan="2" style="width: 15%; {thb}" data-i18n="seg">{i18n("seg")}</th>'
        h2 = ''
        for i, ano in enumerate(anos_unicos):
            is_past = int(ano) < current_year if ano.isdigit() else False
            add_yoy = (not is_past and i > 0)
            if is_past:
                h1 += f'<th style="text-align:center; {thb}">{ano}</th>'
                h2 += f'<th style="text-align:center; {ths} border-bottom:1px solid #e2e8f0;" data-i18n="realized">{i18n("realized")}</th>'
            elif add_yoy:
                ano_ant = anos_unicos[i-1]
                h1 += f'<th colspan="2" style="text-align:center; {thb}">{ano}</th>'
                h2 += f'<th style="text-align:center; {thsb} width:12%; border-bottom:1px solid #e2e8f0;">{versao_rec}</th>'
                h2 += f'<th style="text-align:center; {ths} width:12%; border-bottom:1px solid #e2e8f0;"><span data-i18n="var_yoy">{i18n("var_yoy")}</span> ({ano[-2:]} vs {ano_ant[-2:]})</th>'
            else:
                h1 += f'<th colspan="1" style="text-align:center; {thb}">{ano}</th>'
                h2 += f'<th style="text-align:center; {thsb} width:15%; border-bottom:1px solid #e2e8f0;">{versao_rec}</th>'
        html_thead = f'<thead><tr>{h1}</tr><tr>{h2}</tr></thead>'
    else:
        h1 = f'<th style="width:20%; color:#1e293b;" data-i18n="seg">{i18n("seg")}</th>' + ''.join([f'<th style="text-align:center; font-weight:600; color:#1e293b;">{a}</th>' for a in anos_unicos])
        for i in range(1, len(anos_unicos)):
            h1 += f'<th style="text-align:center; background:#f8fafc; font-weight:600; color:#1e293b;"><span data-i18n="var_yoy">{i18n("var_yoy")}</span> ({anos_unicos[i-1]} ➔ {anos_unicos[i]})</th>'
        html_thead = f'<thead><tr>{h1}</tr></thead>'
    linhas = ""
    totais_ano = {ano: {v: 0 for v in versoes} for ano in anos_unicos}
    for segmento in segmentos_unicos:
        dseg = df_segmentos[df_segmentos[col_s] == segmento]
        tds = ""
        prev = 0
        for i, ano in enumerate(anos_unicos):
            vrec = dseg[(dseg[col_y].astype(str) == str(ano)) & (dseg['Versao'] == versao_rec)][col_v].sum()
            totais_ano[ano][versao_rec] += vrec
            if versao_ant:
                vant = dseg[(dseg[col_y].astype(str) == str(ano)) & (dseg['Versao'] == versao_ant)][col_v].sum()
                totais_ano[ano][versao_ant] += vant
                is_past = int(ano) < current_year if ano.isdigit() else False
                add_yoy = (not is_past and i > 0)
                if is_past:
                    tds += f'<td class="num" style="font-weight:500; color:#334155; background:#f8fafc; border-left:1px dashed #e2e8f0;">{vrec:,.0f}</td>'
                else:
                    tds += f'<td class="num" style="font-weight:500; color:#0f172a; background:#f8fafc; border-left:1px dashed #e2e8f0;">{vrec:,.0f}</td>'
                    if add_yoy:
                        pct = ((vrec / prev) - 1) * 100 if prev > 0 else None
                        if pct is not None:
                            cls = "positive" if pct >= 0 else "negative"
                            s = "+" if pct > 0 else ""
                            ys = f'<span class="{cls}" style="font-weight:500;">{s}{pct:.1f}%</span>'
                        else:
                            ys = '<span class="positive" style="font-weight:500;">+100.0%</span>' if vrec > 0 else "-"
                        tds += f'<td class="num">{ys}</td>'
            else:
                tds += f'<td class="num" style="font-weight:500; color:#0f172a;">{vrec:,.0f}</td>'
            prev = vrec
        linhas += f'<tr><td style="font-weight:600; color:#334155;">{segmento}</td>{tds}</tr>'
    tds_tot = ""
    prev_t = 0
    for i, ano in enumerate(anos_unicos):
        trec = totais_ano[ano][versao_rec]
        if versao_ant:
            is_past = int(ano) < current_year if ano.isdigit() else False
            add_yoy = (not is_past and i > 0)
            tds_tot += f'<td class="num" style="font-weight:700; font-size:1.05em; color:#0f172a; background:#f8fafc; border-left:1px dashed #e2e8f0;">{trec:,.0f}</td>'
            if not is_past and add_yoy:
                pct = ((trec / prev_t) - 1) * 100 if prev_t > 0 else None
                if pct is not None:
                    cls = "positive" if pct >= 0 else "negative"
                    s = "+" if pct > 0 else ""
                    ys = f'<span class="{cls}">{s}{pct:.1f}%</span>'
                else:
                    ys = '<span class="positive">+100.0%</span>' if trec > 0 else "-"
                tds_tot += f'<td class="num" style="font-weight:700; font-size:1em;">{ys}</td>'
        prev_t = trec
    body = f"""
        <div class="header-action-container" style="display:flex; justify-content:space-between; align-items:center;">
            <h3 class="section-title" data-i18n="proj_vol">{i18n("proj_vol")}</h3>
            <button class="btn-excel" onclick="exportProdExcel('{prod_id}', '{mercado}_{produto}')">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                <span data-i18n="export">{i18n("export")}</span>
            </button>
        </div>
        <div class="table-container">
            <table id="table_{prod_id}">{html_thead}
                <tbody>{linhas}</tbody>
                <tfoot><tr style="background:#f1f5f9; border-top:1px solid #cbd5e1;"><td style="text-align:left; font-weight:700; color:#0f172a; text-transform:uppercase;" data-i18n="total">{i18n("total")}</td>{tds_tot}</tr></tfoot>
            </table>
        </div>"""
    html_prod_content = f'<div id="{prod_id}" class="prod-tab-content {p_active_class}" style="display:{p_display_style};">{body}</div>'
    return html_prod_tab, html_prod_content

def gerar_conteudo_pais(mercado, df_mercado, is_active, anos_unicos, versoes, current_year, mes_key, cols, insights_agribusiness, mes_nome_ano):
    col_m, col_p, col_s, col_y, col_v = cols
    mercado_limpo = str(mercado).replace(" ", "_").upper()
    active_class = "active" if is_active else ""
    display_style = "block" if is_active else "none"
    mkt_i18n = str(mercado).lower().replace(' ', '_')
    html_country_tab = f'<button class="tab-btn country-tab-btn {active_class}" onclick="openCountry(event, \'{mercado_limpo}\')" data-i18n="{mkt_i18n}">{i18n(mkt_i18n)}</button>'
    ordem = ['TA', 'CO', 'PA', 'PU']
    produtos = sorted(df_mercado[col_p].unique(), key=lambda p: next(((i, p) for i, pr in enumerate(ordem) if str(p).upper().startswith(pr)), (len(ordem), p)))
    ptabs, pconts = "", ""
    for p_idx, produto in enumerate(produtos):
        dprod = df_mercado[df_mercado[col_p] == produto]
        t, c = gerar_conteudo_produto(produto, dprod, p_idx, mercado, mercado_limpo, anos_unicos, versoes, current_year, mes_key, cols)
        ptabs += t; pconts += c
    mapa = {'BRASIL': 'BRASIL', 'BRA': 'BRASIL', 'ARGENTINA': 'ARGENTINA', 'ARG': 'ARGENTINA', 'MEXICO': 'MEXICO', 'MEX': 'MEXICO', 'OSA': 'OSA'}
    ins = insights_agribusiness.get(mapa.get(mercado.upper(), 'OSA'), {})
    html_ins = gerar_html_insights_agribusiness(ins, mes_nome_ano, mes_key)
    html_country_content = f"""
    <div id="{mercado_limpo}" class="country-tab-content {active_class}" style="display:{display_style};">
        <div class="prod-tabs-container"><h2 class="country-title" data-i18n="select_prod">{i18n("select_prod")}</h2>
        <div class="prod-tabs-nav">{ptabs}</div></div>
        {pconts}{html_ins}
    </div>"""
    return html_country_tab, html_country_content

# ===== ABA METODOLOGIA (XGBoost) — usa sistema de idioma nativo do template =====
METODO_FEATURES = {
    "TA": {"icon": "🚜", "data": {
        "BRASIL": {
            "0-49 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Feijão","Leite","Mandioca","Suínos","Tomate","Uva","Taxa Juros (%)"],
            "50-79 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Café","Laranja","Leite","Milho","Suínos","Tomate","Maçã","Taxa Juros (%)","Inflação (%)"],
            "80-119 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Arroz VP","Bovinos VP","Leite VP","Milho VP","Soja VP","Trigo VP","Taxa Juros (%)","Inflação (%)","Taxa Câmbio (USD)"],
            "120-169 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Milho VP","Soja VP","Wheat Value of Production","Taxa Juros (%)","Taxa Câmbio (USD)","Preço Soja (USD/Ton)"],
            "170-239 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Milho VP","Soja VP","Taxa Juros (%)","Preço Soja (USD/Ton)"],
            "240-339 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Milho VP","Soja VP","Corn Production","Corn Wholesale Price"],
            "+340 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Milho VP","Soja VP","Taxa Juros (%)","Inflação (%)","Taxa Câmbio (USD)","Preço Soja (USD/Ton)"],
        },
        "ARGENTINA": {
            "0-49 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Apple PR","Tomato Production","Potato PR","Taxa Juros (%)","Inflação (%)"],
            "50-79 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Apple PR","Beef and Veal Meat Domestic","Tomato Production","Olive Producer Price","Peach & Nectarine PR","Taxa Juros (%)","Inflação (%)"],
            "80-119 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Beef and Veal Meat Domestic","Wheat Value of Production","Corn Port Price"],
            "120-169 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Wheat Value of Production","Corn Value of Production","Soybean farm prices","Taxa Juros (%)"],
            "170-239 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Wheat Value of Production","Corn Value of Production","Soybean oil production"],
            "240-339 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Wheat Value of Production","Corn Value of Production","Soybean PR"],
            "+340 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Wheat Value of Production","Corn Value of Production","Soybean oil production"],
        },
        "OSA": {
            "0-49 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Apple PR","Coffee Area Harvested","Coffee Producer Price","Coffee Production","Olive Producer Price","Orange Wholesale Price","Peach & Nectarine PR"],
            "50-79 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Apple PR","Coffee Area Harvested","Coffee Producer Price","Coffee Production","Olive Producer Price","Orange Wholesale Price","Peach & Nectarine PR"],
            "80-119 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Beef and Veal Meat Domestic","Soybean PR"],
            "120-169 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Soybean oil production","Soybean PR"],
            "170-239 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Soybean oil production","Soybean PR"],
            "240-339 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Soybean oil production","Soybean PR"],
            "+340 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Soybean oil production","Soybean PR"],
        },
        "MEXICO": {
            "0-49 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Avocado Area Harvested","Corn Area Harvested"],
            "50-79 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Avocado Area Harvested","Corn Area Harvested"],
            "80-119 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Sugar Cane Area Harvested"],
            "120-169 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Corn Production"],
            "170-239 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Sugar Cane Area Harvested"],
            "240-339 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Sugar Cane Area Harvested"],
            "+340 HP": ["Sazonalidade (ano)","Sazonalidade (mês)","Sugar Cane Area Harvested"],
        },
    }},
    "CO": {"icon": "🌾", "data": {
        "BRASIL": {
            "Classe 4": ["Sazonalidade (ano)","Sazonalidade (mês)","Arroz VP","Milho VP","Soja VP","Trigo VP","Corn Wholesale Price","Taxa Juros (%)","Soybean port prices"],
            "Classe 5": ["Sazonalidade (ano)","Sazonalidade (mês)","Arroz VP","Milho VP","Soja VP","Trigo VP","Corn Wholesale Price","Taxa Juros (%)","Soybean port prices"],
            "Classe 6": ["Sazonalidade (ano)","Sazonalidade (mês)","Milho VP","Soja VP","Taxa Juros (%)"],
            "Classe 7": ["Sazonalidade (ano)","Sazonalidade (mês)","Algodão VP","Milho VP","Soybean VP","Taxa Juros (%)"],
            "Classe 8+": ["Sazonalidade (ano)","Sazonalidade (mês)","Algodão VP","Milho VP","Soybean VP","Taxa Juros (%)"],
        },
        "ARGENTINA": {
            "Classe 4 e 5": ["Sazonalidade (ano)","Sazonalidade (mês)","Wheat Value of Production","Corn Area Harvested","Corn Port Price","Corn Production","Corn Value of Production","Corn Wholesale Price","Taxa Juros (%)","Índice Crédito Agrícola","Soybean oil production"],
            "Classe 6": ["Sazonalidade (ano)","Sazonalidade (mês)","Wheat Value of Production","Corn Port Price","Corn Value of Production","Soybean oil production"],
            "Classe 7": ["Sazonalidade (ano)","Sazonalidade (mês)","Wheat Value of Production","Corn Value of Production","Soybean oil production"],
            "Classe 8+": ["Sazonalidade (ano)","Sazonalidade (mês)","Wheat Value of Production","Corn Value of Production","Soybean oil production"],
        },
        "OSA": {
            "Classe 4": ["Sazonalidade (ano)","Sazonalidade (mês)","Rice VP"],
            "Classe 5": ["Sazonalidade (ano)","Sazonalidade (mês)","Rice VP","Soybean PR"],
            "Classe 6": ["Sazonalidade (ano)","Sazonalidade (mês)","Rice VP","Soybean oil production"],
            "Classe 7": ["Sazonalidade (ano)","Sazonalidade (mês)","Soybeans production"],
            "Classe 8+": ["Sazonalidade (ano)","Sazonalidade (mês)","Soybeans production"],
        },
    }},
    "PA": {"icon": "🌱", "data": {
        "BRASIL": {
            "< 20 linhas": ["Sazonalidade (ano)","Sazonalidade (mês)","Milho VP","Soja VP","Trigo VP","Taxa Juros (%)","Preço Soja (USD/Ton)"],
            "> 20 linhas": ["Sazonalidade (ano)","Sazonalidade (mês)","Milho VP","Soja VP","Taxa Juros (%)","Soybean port prices"],
        },
        "OSA": {"Todos os segmentos": ["Sazonalidade (ano)","Sazonalidade (mês)","Soybeans PR"]},
    }},
    "PU": {"icon": "💧", "data": {
        "BRASIL": {"Todos os segmentos": ["Sazonalidade (ano)","Sazonalidade (mês)","Cana-de-Açúcar VP","Milho VP","Soja VP","Taxa Juros (%)","Taxa Câmbio (USD)"]},
        "ARGENTINA": {"Todos os segmentos": ["Sazonalidade (ano)","Sazonalidade (mês)","Wheat Value of Production","Corn Value of Production","Soybean PR","Taxa Juros (%)"]},
        "OSA": {"Todos os segmentos": ["Sazonalidade (ano)","Sazonalidade (mês)","Soybean PR"]},
    }},
}
METODO_PRODUCT_ORDER = ["TA", "CO", "PA", "PU"]
METODO_MARKET_ORDER = ["BRASIL", "ARGENTINA", "OSA", "MEXICO"]
METODO_FLAG = {"BRASIL": "🇧🇷", "ARGENTINA": "🇦🇷", "OSA": "🌎", "MEXICO": "🇲🇽"}
METODO_PROD_I18N = {"TA": "ta", "CO": "co", "PA": "pa", "PU": "pu"}

def _metodo_classify(var):
    v = var.lower()
    if "sazonalidade" in v: return "temporal"
    if any(k in v for k in ["juros","inflação","inflacao","câmbio","cambio","crédito","credito","índice","indice"]): return "macro"
    if any(k in v for k in ["price","preço","preco","usd","wholesale","port price","farm price","port prices"]) or v.endswith(" pr"): return "macro"
    return "agro"

def _metodo_pill(var):
    cat = {"temporal": "vp-temporal", "agro": "vp-agro", "macro": "vp-macro"}[_metodo_classify(var)]
    return f'<span class="vpill {cat}">{html.escape(var)}</span>'

def _metodo_market_block(pcode, market):
    data = METODO_FEATURES[pcode]["data"].get(market)
    mkt_i18n = market.lower()
    prod_i18n = METODO_PROD_I18N[pcode]
    head = (f'<div class="market-head"><span class="market-flag">{METODO_FLAG[market]}</span>'
            f'<h4 data-i18n="{mkt_i18n}">{i18n(mkt_i18n)}</h4>'
            f'<span class="market-tag" data-i18n="{prod_i18n}">{i18n(prod_i18n)}</span></div>')
    if not data:
        na = ml("Não aplicável — este mercado não comercializa esta família.",
                "Not applicable — this market does not sell this family.",
                "No aplicable — este mercado no comercializa esta familia.")
        return f'<div class="market-block">{head}<div class="detal-card"><div class="na-state">➖ {na}</div></div></div>'
    rows = ""
    for seg, varlist in data.items():
        pills = "".join(_metodo_pill(v) for v in varlist)
        rows += f'<tr><td class="seg-cell">{html.escape(seg)}</td><td class="var-cell">{pills}</td><td class="count-cell"><span class="count-badge">{len(varlist)}</span></td></tr>'
    return (f'<div class="market-block">{head}<div class="detal-card"><table class="detal-table">'
            f'<thead><tr><th style="width:16%" data-i18n="met_col_seg">{i18n("met_col_seg")}</th>'
            f'<th data-i18n="met_col_vars">{i18n("met_col_vars")}</th>'
            f'<th style="width:8%;text-align:center" data-i18n="met_col_n">{i18n("met_col_n")}</th></tr></thead>'
            f'<tbody>{rows}</tbody></table></div></div>')

def _metodo_product_panel(pcode, active):
    blocks = "".join(_metodo_market_block(pcode, m) for m in METODO_MARKET_ORDER)
    return f'<div id="detal-{pcode}" class="detal-panel" style="display:{"block" if active else "none"}">{blocks}</div>'

def _metodo_detalhamento():
    legend = (f'<div class="detal-legend"><span class="legend-title" data-i18n="met_legend_title">{i18n("met_legend_title")}</span>'
              f'<span class="vpill vp-temporal" data-i18n="met_temporal">{i18n("met_temporal")}</span>'
              f'<span class="vpill vp-agro" data-i18n="met_agro">{i18n("met_agro")}</span>'
              f'<span class="vpill vp-macro" data-i18n="met_macro">{i18n("met_macro")}</span></div>')
    intro = '<div class="method-card"><p>' + ml(
        "Abaixo, o conjunto de variáveis (features) que alimenta o modelo XGBoost em cada segmento, por mercado. São os elementos de maior importância preditiva por família de produto. Todos os modelos incluem a sazonalidade (ano e mês) como base temporal, somada a drivers agronômicos/setoriais e macroeconômicos específicos de cada segmento.",
        "Below is the set of variables (features) that feeds the XGBoost model in each segment, by market. These are the elements of greatest predictive importance per product family. All models include seasonality (year and month) as the temporal base, plus agronomic/sector and macroeconomic drivers specific to each segment.",
        "A continuación, el conjunto de variables (features) que alimenta el modelo XGBoost en cada segmento, por mercado. Son los elementos de mayor importancia predictiva por familia de producto. Todos los modelos incluyen la estacionalidad (año y mes) como base temporal, sumada a drivers agronómicos/sectoriales y macroeconómicos específicos de cada segmento.") + '</p></div>'
    prod_tabs = "".join(
        f'<button class="detal-prod-btn {"active" if i==0 else ""}" onclick="openProd(event, \'{c}\')">{METODO_FEATURES[c]["icon"]} <span data-i18n="{METODO_PROD_I18N[c]}">{i18n(METODO_PROD_I18N[c])}</span></button>'
        for i, c in enumerate(METODO_PRODUCT_ORDER))
    panels = "".join(_metodo_product_panel(c, i == 0) for i, c in enumerate(METODO_PRODUCT_ORDER))
    complete = ml(
        "Detalhamento completo: as quatro famílias (Tratores, Colheitadeiras, Plantadeiras e Pulverizadores) estão mapeadas por mercado e segmento. Onde um mercado não comercializa determinada família, o item aparece como \"Não aplicável\".",
        "Complete breakdown: the four families (Tractors, Combines, Planters and Sprayers) are mapped by market and segment. Where a market does not sell a given family, the item appears as \"Not applicable\".",
        "Detalle completo: las cuatro familias (Tractores, Cosechadoras, Sembradoras y Pulverizadores) están mapeadas por mercado y segmento. Donde un mercado no comercializa una familia, el ítem aparece como \"No aplicable\".")
    return (f'<h3 class="section-title" style="margin-top:1.6rem" data-i18n="met_detalhamento_title">{i18n("met_detalhamento_title")}</h3>'
            f'<hr class="agco-divider">{intro}{legend}<div class="detal-prod-nav">{prod_tabs}</div>{panels}'
            f'<div class="next" style="margin-top:1.2rem">✅ {complete}</div>')

def gerar_conteudo_metodologia():
    tab = f'<button class="tab-btn country-tab-btn" onclick="openCountry(event, \'METODOLOGIA\')" data-i18n="metodologia">{i18n("metodologia")}</button>'
    open_prod_script = ("<script>function openProd(evt, code){"
        "document.querySelectorAll('.detal-panel').forEach(function(el){el.style.display='none';});"
        "document.querySelectorAll('.detal-prod-btn').forEach(function(b){b.classList.remove('active');});"
        "var t=document.getElementById('detal-'+code); if(t){t.style.display='block';}"
        "if(evt&&evt.currentTarget){evt.currentTarget.classList.add('active');}}</script>")
    lead = ml(
        "Todas as projeções deste dashboard — para todos os mercados (Brasil, Argentina, México e OSA) e todas as famílias de produto (Tratores, Colheitadeiras, Plantadeiras e Pulverizadores) — são geradas por um mesmo motor estatístico: o algoritmo XGBoost (eXtreme Gradient Boosting).",
        "All projections in this dashboard — for all markets (Brazil, Argentina, Mexico and OSA) and all product families (Tractors, Combines, Planters and Sprayers) — are generated by a single statistical engine: the XGBoost (eXtreme Gradient Boosting) algorithm.",
        "Todas las proyecciones de este dashboard — para todos los mercados (Brasil, Argentina, México y OSA) y todas las familias de producto (Tractores, Cosechadoras, Sembradoras y Pulverizadores) — son generadas por un mismo motor estadístico: el algoritmo XGBoost (eXtreme Gradient Boosting).")
    lead2 = ml(
        "Adotar um único método para toda a base garante consistência metodológica e permite comparar mercados e segmentos sob o mesmo critério, sem vieses de modelos diferentes.",
        "Using a single method across the whole base ensures methodological consistency and allows comparing markets and segments under the same criterion, without biases from different models.",
        "Adoptar un único método para toda la base garantiza consistencia metodológica y permite comparar mercados y segmentos bajo el mismo criterio, sin sesgos de modelos diferentes.")
    whatT = ml("O que é o XGBoost?", "What is XGBoost?", "¿Qué es XGBoost?", tag="h4")
    whatP = ml(
        "O XGBoost é um algoritmo de machine learning baseado em árvores de decisão combinadas pela técnica de gradient boosting. Em vez de treinar um único modelo, ele constrói centenas de pequenas árvores em sequência, onde cada nova árvore aprende a corrigir os erros deixados pelas anteriores. A previsão final é a soma das contribuições de todas as árvores.",
        "XGBoost is a machine learning algorithm based on decision trees combined through the gradient boosting technique. Instead of training a single model, it builds hundreds of small trees in sequence, where each new tree learns to correct the errors left by the previous ones. The final prediction is the sum of the contributions of all trees.",
        "XGBoost es un algoritmo de machine learning basado en árboles de decisión combinados por la técnica de gradient boosting. En lugar de entrenar un único modelo, construye cientos de pequeños árboles en secuencia, donde cada nuevo árbol aprende a corregir los errores dejados por los anteriores. La previsión final es la suma de las contribuciones de todos los árboles.")
    golf = ml(
        "🏌️ Analogia do golfe: a primeira árvore dá a \"tacada inicial\" em direção ao alvo (a demanda real). Ela não acerta de primeira. A segunda árvore corrige a distância que faltou; a terceira dá um ajuste fino — e assim por diante, até chegar bem perto do valor correto.",
        "🏌️ Golf analogy: the first tree takes the \"tee shot\" toward the target (real demand). It doesn't hit it right away. The second tree corrects the remaining distance; the third makes a fine adjustment — and so on, until it gets very close to the correct value.",
        "🏌️ Analogía del golf: el primer árbol da el \"golpe inicial\" hacia el objetivo (la demanda real). No acierta de primera. El segundo árbol corrige la distancia que faltó; el tercero hace un ajuste fino — y así sucesivamente, hasta acercarse mucho al valor correcto.")
    stepsT = ml("Como o modelo aprende — em 4 passos", "How the model learns — in 4 steps", "Cómo aprende el modelo — en 4 pasos", tag="h4")
    s1t = ml("Ponto de partida","Starting point","Punto de partida",tag="h5")
    s1p = ml("O modelo começa com uma estimativa-base (ex.: a média histórica de vendas do segmento).","The model starts with a base estimate (e.g., the segment's historical sales average).","El modelo comienza con una estimación base (ej.: la media histórica de ventas del segmento).")
    s2t = ml("Mede o erro","Measures the error","Mide el error",tag="h5")
    s2p = ml("Compara a previsão com o realizado e calcula o resíduo — o quanto errou para mais ou para menos.","Compares the forecast with the actual and calculates the residual — how much it missed, over or under.","Compara la previsión con lo realizado y calcula el residuo — cuánto se desvió por encima o por debajo.")
    s3t = ml("Corrige","Corrects","Corrige",tag="h5")
    s3p = ml("Uma nova árvore é treinada focando exatamente nos pontos mais difíceis de prever.","A new tree is trained focusing exactly on the hardest points to predict.","Se entrena un nuevo árbol enfocándose exactamente en los puntos más difíciles de prever.")
    s4t = ml("Repete e soma","Repeat and sum","Repite y suma",tag="h5")
    s4p = ml("O ciclo se repete por centenas de rodadas; a projeção final é a soma de todas as árvores.","The cycle repeats for hundreds of rounds; the final projection is the sum of all trees.","El ciclo se repite por cientos de rondas; la proyección final es la suma de todos los árboles.")
    whyT = ml("Por que escolhemos o XGBoost","Why we chose XGBoost","Por qué elegimos XGBoost",tag="h4")
    f1t = ml("Alta precisão em dados tabulares","High accuracy on tabular data","Alta precisión en datos tabulares",tag="b")
    f1p = ml("É referência de mercado para dados estruturados como os nossos (volumes por ano, região, produto e segmento).","It is a market reference for structured data like ours (volumes by year, region, product and segment).","Es referencia de mercado para datos estructurados como los nuestros (volúmenes por año, región, producto y segmento).")
    f2t = ml("Captura relações não-lineares","Captures non-linear relationships","Captura relaciones no lineales",tag="b")
    f2p = ml("Entende interações complexas entre variáveis (ex.: preço da soja + câmbio + crédito agindo juntos).","Understands complex interactions between variables (e.g., soy price + FX + credit acting together).","Entiende interacciones complejas entre variables (ej.: precio de la soja + cambio + crédito actuando juntos).")
    f3t = ml("Regularização anti-overfitting","Anti-overfitting regularization","Regularización anti-overfitting",tag="b")
    f3p = ml("Penalidades L1/L2 evitam que o modelo \"decore\" o passado e o mantêm robusto para prever o futuro.","L1/L2 penalties prevent the model from \"memorizing\" the past and keep it robust to predict the future.","Las penalidades L1/L2 evitan que el modelo \"memorice\" el pasado y lo mantienen robusto para prever el futuro.")
    f4t = ml("Importância de variáveis","Variable importance","Importancia de variables",tag="b")
    f4p = ml("Mostra quais fatores mais pesam em cada previsão — a base do detalhamento por região/produto/segmento a seguir.","Shows which factors weigh most in each prediction — the basis of the region/product/segment breakdown below.","Muestra qué factores pesan más en cada previsión — la base del detalle por región/producto/segmento a continuación.")
    f5t = ml("Rápido e escalável","Fast and scalable","Rápido y escalable",tag="b")
    f5p = ml("Treina em segundos e roda igual para os 4 mercados e 4 produtos, com processamento paralelo.","Trains in seconds and runs the same for the 4 markets and 4 products, with parallel processing.","Entrena en segundos y funciona igual para los 4 mercados y 4 productos, con procesamiento paralelo.")
    f6t = ml("Lida com dados faltantes","Handles missing data","Maneja datos faltantes",tag="b")
    f6p = ml("Trata automaticamente lacunas e outliers, comuns em séries históricas de mercado.","Automatically handles gaps and outliers, common in market historical series.","Trata automáticamente lagunas y outliers, comunes en series históricas de mercado.")
    content = (
        '<div id="METODOLOGIA" class="country-tab-content" style="display:none;">'
        f'<h3 class="section-title"><span data-i18n="met_title">{i18n("met_title")}</span> <span class="method-pill" data-i18n="met_badge">{i18n("met_badge")}</span></h3>'
        '<hr class="agco-divider">'
        f'<div class="method-card"><p class="lead">{lead}</p><p>{lead2}</p></div>'
        f'<div class="method-card">{whatT}<p>{whatP}</p><div class="callout">{golf}</div></div>'
        f'<div class="method-card">{stepsT}<div class="steps">'
        f'<div class="step"><div class="n">1</div>{s1t}<p>{s1p}</p></div>'
        f'<div class="step"><div class="n">2</div>{s2t}<p>{s2p}</p></div>'
        f'<div class="step"><div class="n">3</div>{s3t}<p>{s3p}</p></div>'
        f'<div class="step"><div class="n">4</div>{s4t}<p>{s4p}</p></div></div></div>'
        f'<div class="method-card">{whyT}<div class="feature-grid">'
        f'<div class="feature"><span class="ic">🎯</span><div>{f1t}<span>{f1p}</span></div></div>'
        f'<div class="feature"><span class="ic">🔗</span><div>{f2t}<span>{f2p}</span></div></div>'
        f'<div class="feature"><span class="ic">🛡️</span><div>{f3t}<span>{f3p}</span></div></div>'
        f'<div class="feature"><span class="ic">📊</span><div>{f4t}<span>{f4p}</span></div></div>'
        f'<div class="feature"><span class="ic">⚡</span><div>{f5t}<span>{f5p}</span></div></div>'
        f'<div class="feature"><span class="ic">🕳️</span><div>{f6t}<span>{f6p}</span></div></div></div></div>'
        + _metodo_detalhamento() + open_prod_script + '</div>')
    return tab, content

CSS_METODOLOGIA = r"""
    /* ============================================================
       IDENTIDADE VISUAL "EARLY SIGNALS" — AGCO
       Paleta, tipografia e layout alinhados ao dashboard
       Early Signals - LATAM Executive Intelligence
       ============================================================ */
    :root {
        --agco-red: #cc0000;
        --agco-red-dark: #a30000;
        --agco-red-soft: #fce4d6;
        --agco-black: #111111;
        --agco-dark-gray: #333333;
        --agco-light-gray: #f4f4f4;
        --text-main: #222222;
        --text-muted: #666666;
        --white: #ffffff;
        --primary-bg: #e9ecef;
        --card-bg: #ffffff;
        --surface-2: #fafafa;
        --border-color: #e0e0e0;
        --border-strong: #dddddd;

        /* Faróis Early Signals */
        --farol-positive-bg: #e2f0d9;
        --farol-positive-text: #385723;
        --farol-positive-dot: #70ad47;
        --farol-warning-bg: #fff2cc;
        --farol-warning-text: #7f6000;
        --farol-warning-dot: #ffc000;
        --farol-critical-bg: #fce4d6;
        --farol-critical-text: #c65911;
        --farol-critical-dot: #c00000;

        --pos-color: #385723;
        --neg-color: #c00000;
        --neutral-color: #7f6000;
        --color-hp: #385723; --bg-hp: #e2f0d9;
        --color-n:  #7f6000; --bg-n:  #fff2cc;
        --color-hn: #c65911; --bg-hn: #fce4d6;

        --shadow-soft: 0 4px 6px rgba(0,0,0,0.02);
        --shadow-hover: 0 8px 16px rgba(0,0,0,0.07);
        --radius: 4px;
    }
    *, *::before, *::after { box-sizing: border-box; }

    body {
        font-family: 'Arial', sans-serif !important;
        background-color: var(--primary-bg) !important;
        color: var(--text-main) !important;
        margin: 0; padding: 20px;
        line-height: 1.5;
    }
    h1, h2, h3, h4, h5 { font-family: 'Arial', sans-serif; color: var(--agco-black); }
    ::selection { background: rgba(204,0,0,.16); }

    /* ===== CONTAINER / HEADER (padrão Early Signals) ===== */
    .dashboard-header {
        max-width: 1350px; margin: 0 auto 25px auto;
        background-color: var(--agco-black) !important;
        color: var(--white);
        padding: 30px 40px !important;
        border-radius: var(--radius) !important;
        border-bottom: 6px solid var(--agco-red);
        box-shadow: 0 10px 25px rgba(0,0,0,0.15) !important;
        display: flex; justify-content: space-between; align-items: center; gap: 20px;
    }
    .header-title h1 {
        color: var(--white) !important; margin: 0;
        font-size: 28px; font-weight: 900;
        text-transform: uppercase; letter-spacing: 1.5px;
    }
    .header-title p { color: #b0b0b0 !important; font-size: 12px; text-transform: uppercase; letter-spacing: .5px; }

    .kpi-container { display: flex; gap: 15px; }
    .kpi-card {
        background: #1d1d1d !important; border: 1px solid #2f2f2f;
        border-left: 4px solid var(--agco-red) !important;
        border-radius: var(--radius); padding: 14px 22px; min-width: 170px;
    }
    .kpi-card.positive { border-left-color: var(--farol-positive-dot) !important; }
    .kpi-card.negative { border-left-color: var(--agco-red) !important; }
    .kpi-title { color: #a0a0a0 !important; font-size: 11px; text-transform: uppercase; letter-spacing: .5px; font-weight: bold; }
    .kpi-value { color: var(--white) !important; font-size: 24px; font-weight: 900; }

    /* ===== CONTROLES DE IDIOMA ===== */
    .lang-controls { max-width: 1350px; margin: 0 auto 12px auto; text-align: right; }
    .lang-btn {
        background-color: #2a2a2a; border: none; color: #a0a0a0;
        padding: 6px 12px; margin-left: 4px; font-size: 12px; font-weight: bold;
        border-radius: 2px; cursor: pointer; transition: all .2s ease-in-out;
    }
    .lang-btn:hover { background-color: #444; color: var(--white); }
    .lang-btn.active { background-color: var(--agco-red); color: var(--white); }

    /* ===== ABAS DE PAÍS (tab-btn Early Signals) ===== */
    .country-tabs-nav {
        max-width: 1350px; margin: 0 auto 25px auto;
        display: flex; flex-wrap: wrap; gap: 5px;
        border-bottom: 3px solid var(--agco-black); padding-bottom: 5px;
    }
    .country-tab-btn {
        background-color: var(--agco-light-gray); color: var(--agco-dark-gray);
        border: none; padding: 12px 20px; font-size: 13px; font-weight: bold;
        cursor: pointer; text-transform: uppercase; letter-spacing: .5px;
        border-radius: 4px 4px 0 0; transition: all .2s;
    }
    .country-tab-btn:hover { background-color: #e0e0e0; color: var(--agco-black); }
    .country-tab-btn.active, .country-tab-btn.active:hover {
        background-color: var(--agco-black) !important; color: var(--white) !important;
        border-bottom: 3px solid var(--agco-red) !important;
    }
    .country-tab-content { max-width: 1350px; margin: 0 auto; }

    /* ===== ABAS DE PRODUTO ===== */
    .prod-tabs-container {
        display: flex; align-items: center; gap: 16px; flex-wrap: wrap;
        border-bottom: 2px solid var(--border-strong); margin-bottom: 20px;
    }
    .country-title {
        font-size: 13px; color: var(--agco-dark-gray); margin: 0 1rem 0 0;
        font-weight: 800; text-transform: uppercase; letter-spacing: 1px; border: none;
    }
    .prod-tabs-nav { display: flex; gap: 5px; border: none; padding-bottom: 0; }
    .prod-tab-btn {
        background: transparent; border: none; padding: 12px 18px;
        color: var(--text-muted); font-size: 13px; font-weight: bold;
        text-transform: uppercase; letter-spacing: .5px; cursor: pointer;
        border-bottom: 3px solid transparent; margin-bottom: -2px; border-radius: 0;
    }
    .prod-tab-btn:hover { color: var(--agco-black); background: var(--agco-light-gray); }
    .prod-tab-btn.active {
        color: var(--agco-black) !important; background: transparent !important;
        border-bottom-color: var(--agco-red) !important;
    }

    /* ===== BOTÃO EXCEL ===== */
    .btn-excel { background: var(--agco-black); border-radius: 2px; text-transform: uppercase; font-size: 11px; letter-spacing: .5px; }
    .btn-excel:hover { background: var(--agco-red); }

    /* ===== TÍTULOS DE SEÇÃO ===== */
    .section-title {
        font-size: 18px; color: var(--agco-dark-gray); font-weight: 700;
        text-transform: uppercase; letter-spacing: 1px;
        display: flex; align-items: center; position: relative;
        padding-left: .8rem; margin: 10px 0;
    }
    .section-title::before {
        content: ''; position: absolute; left: 0; top: 50%; transform: translateY(-50%);
        width: 5px; height: 1.1em; background: var(--agco-red);
    }
    .agco-divider { border: 0; height: 3px; background: var(--agco-black); margin: 8px 0 18px; }
    .header-action-container { display: flex; justify-content: space-between; align-items: center; }
    .month-badge {
        font-size: 11px; font-weight: bold; color: #fff; background: var(--agco-red);
        padding: 6px 14px; border-radius: 2px; margin-left: 10px;
        text-transform: uppercase; letter-spacing: 1px;
    }

    /* ===== TABELAS ===== */
    .table-container {
        background: var(--white); border: 1px solid var(--border-color);
        border-top: 5px solid var(--agco-red); border-radius: var(--radius);
        box-shadow: var(--shadow-soft); padding: 20px; overflow-x: auto;
    }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { padding: 10px 14px; text-align: left; vertical-align: middle; border-bottom: 1px solid var(--border-color); }
    thead th {
        background-color: var(--agco-light-gray); color: var(--agco-black);
        text-transform: uppercase; letter-spacing: .5px; font-size: 11px; font-weight: bold;
    }
    tbody tr:nth-child(even) { background: #fbfbfb; }
    tbody tr:hover { background: var(--agco-light-gray); }
    td.num { text-align: center; font-variant-numeric: tabular-nums; }
    tfoot td { border-top: 3px solid var(--agco-black); }
    .positive { color: var(--farol-positive-text); }
    .negative { color: var(--farol-critical-dot); }

    /* ===== INSIGHTS / FARÓIS ===== */
    .insights-section { margin-top: 30px; }
    .insights-container { display: grid; grid-template-columns: 3fr 7fr; gap: 25px; align-items: start; margin-top: 10px; }
    .insight-summary-card, .insight-news-container {
        background: var(--white); border: 1px solid var(--border-color);
        border-radius: var(--radius); padding: 18px 20px; box-shadow: var(--shadow-soft);
    }
    .insight-summary-card h4, .insight-news-container h4 {
        margin: 0 0 14px; font-size: 11px; color: var(--agco-black);
        text-transform: uppercase; letter-spacing: .5px; font-weight: bold;
        border-bottom: 1px solid var(--border-color); padding-bottom: 6px;
    }
    .summary-item { display: flex; align-items: center; gap: 12px; padding: 10px 8px; border-radius: var(--radius); }
    .summary-item:hover { background: var(--surface-2); }
    .summary-item.positive { border-left: 5px solid var(--farol-positive-dot); }
    .summary-item.negative { border-left: 5px solid var(--farol-critical-dot); }
    .summary-item.neutral  { border-left: 5px solid var(--farol-warning-dot); }
    .summary-item-icon { font-size: 20px; }
    .summary-item-text { flex-grow: 1; font-size: 13px; color: var(--agco-dark-gray); }
    .summary-item-text strong { color: var(--agco-black); text-transform: uppercase; font-size: 11px; letter-spacing: .5px; }
    .summary-sentiment-badge {
        font-size: 11px; font-weight: bold; text-transform: uppercase; letter-spacing: .5px;
        padding: 4px 10px; border-radius: 4px; margin-left: auto; white-space: nowrap;
    }
    .summary-sentiment-badge.positive { color: var(--farol-positive-text); background: var(--farol-positive-bg); }
    .summary-sentiment-badge.negative { color: var(--farol-critical-text); background: var(--farol-critical-bg); }
    .summary-sentiment-badge.neutral  { color: var(--farol-warning-text);  background: var(--farol-warning-bg); }

    .news-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; max-height: 560px; overflow: auto; }
    @media (max-width: 1024px) { .news-grid { grid-template-columns: 1fr; } }
    .news-card {
        border: 1px solid var(--border-color); border-radius: var(--radius);
        padding: 0; background: var(--white); box-shadow: var(--shadow-soft);
        transition: transform .25s ease, box-shadow .25s ease; overflow: hidden;
    }
    .news-card:hover { transform: translateY(-4px); box-shadow: var(--shadow-hover); }
    .news-card.positive-card { border-left: 5px solid var(--farol-positive-dot); }
    .news-card.negative-card { border-left: 5px solid var(--farol-critical-dot); }
    .news-card.neutral-card  { border-left: 5px solid var(--farol-warning-dot); }
    .news-card-header {
        display: flex; align-items: center; gap: 10px; flex-wrap: wrap;
        background-color: var(--agco-light-gray); padding: 12px 16px; margin-bottom: 0;
    }
    .news-sentiment-dot { width: 8px; height: 8px; border-radius: 50%; }
    .news-sentiment-label { font-size: 11px; font-weight: bold; text-transform: uppercase; letter-spacing: .5px; }
    .news-title { margin: 0; font-size: 13px; font-weight: bold; color: var(--agco-black); text-transform: uppercase; line-height: 1.4; width: 100%; }
    .news-summary { margin: 0; padding: 16px; font-size: 13px; color: var(--agco-dark-gray); line-height: 1.5; text-align: justify; }
    .news-source-country {
        font-size: 10px; font-weight: bold; color: #fff; background: var(--agco-black);
        padding: 3px 8px; border-radius: 2px; margin-left: 6px; letter-spacing: .5px;
    }

    /* ===== ABA METODOLOGIA ===== */
    .method-card {
        background: var(--card-bg); border: 1px solid var(--border-color);
        border-top: 5px solid var(--agco-red); border-radius: var(--radius);
        box-shadow: var(--shadow-soft); padding: 20px 24px; margin-bottom: 18px;
    }
    .method-card p { margin: .55rem 0; line-height: 1.55; color: var(--agco-dark-gray); text-align: justify; }
    .method-card .lead { font-size: 15px; color: var(--text-main); }
    .method-card h4 { font-size: 14px; text-transform: uppercase; letter-spacing: .5px; margin: 0 0 10px; border-bottom: 1px solid var(--border-color); padding-bottom: 6px; }
    .method-pill {
        display: inline-block; font-size: 11px; font-weight: bold; color: #fff;
        background: var(--agco-red); padding: 5px 12px; border-radius: 2px;
        text-transform: uppercase; letter-spacing: 1px; margin-left: 12px;
    }
    .model-name { font-weight: 900; color: var(--agco-red); }

    .steps { display: grid; grid-template-columns: repeat(4,1fr); gap: 15px; margin-top: .4rem; }
    .step { background: var(--white); border: 2px solid var(--agco-red); border-radius: 6px; padding: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.04); transition: transform .25s ease, box-shadow .25s ease; }
    .step:hover { transform: translateY(-4px); box-shadow: var(--shadow-hover); }
    .step .n {
        width: 34px; height: 34px; border-radius: 8px; background: var(--agco-black);
        color: #fff; font-weight: 900; font-size: 14px;
        display: flex; align-items: center; justify-content: center; margin-bottom: 12px;
    }
    .step h5 { margin: 0 0 .3rem; font-size: 12px; text-transform: uppercase; letter-spacing: .5px; }
    .step p { margin: 0; font-size: 12px; color: var(--text-muted); line-height: 1.5; }

    .feature-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px,1fr)); gap: 15px; margin-top: .4rem; }
    .feature {
        display: flex; gap: 12px; align-items: flex-start; background: var(--white);
        border: 2px solid var(--agco-red); border-radius: 6px; padding: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04); transition: transform .25s ease, box-shadow .25s ease;
    }
    .feature:hover { transform: translateY(-4px); box-shadow: var(--shadow-hover); }
    .feature .ic {
        font-size: 20px; width: 44px; height: 44px; flex-shrink: 0; border-radius: 8px;
        display: flex; align-items: center; justify-content: center;
        background: var(--farol-critical-bg); color: var(--farol-critical-text);
    }
    .feature b { color: var(--agco-black); font-size: 12px; text-transform: uppercase; letter-spacing: .5px; }
    .feature span { display: block; font-size: 12px; color: var(--text-muted); margin-top: .25rem; }

    .callout {
        background: var(--agco-light-gray); border-left: 5px solid var(--agco-red);
        padding: 14px 18px; font-size: 13px; color: var(--agco-dark-gray); margin-top: .8rem;
    }
    .callout strong { color: var(--agco-red); }
    .next {
        border: 1px dashed var(--border-strong); border-radius: var(--radius);
        padding: 14px 18px; background: var(--white); color: var(--agco-dark-gray); font-size: 13px;
    }

    .detal-legend { display: flex; align-items: center; flex-wrap: wrap; gap: 8px; margin: .2rem 0 1rem; font-size: 12px; }
    .detal-legend .legend-title { font-weight: bold; color: var(--agco-dark-gray); text-transform: uppercase; letter-spacing: .5px; font-size: 11px; margin-right: 6px; }
    .detal-prod-nav { display: flex; flex-wrap: wrap; gap: 5px; border-bottom: 3px solid var(--agco-black); padding-bottom: 5px; margin-bottom: 20px; }
    .detal-prod-btn {
        background-color: var(--agco-light-gray); color: var(--agco-dark-gray); border: none;
        padding: 12px 20px; font-size: 13px; font-weight: bold; text-transform: uppercase;
        letter-spacing: .5px; cursor: pointer; border-radius: 4px 4px 0 0; transition: all .2s;
    }
    .detal-prod-btn:hover { background-color: #e0e0e0; color: var(--agco-black); }
    .detal-prod-btn.active { background-color: var(--agco-black); color: var(--white); border-bottom: 3px solid var(--agco-red); }

    .market-block { margin-bottom: 25px; }
    .market-head { display: flex; align-items: center; gap: 10px; margin: .2rem 0 .6rem; }
    .market-head h4 { margin: 0; font-size: 16px; text-transform: uppercase; letter-spacing: 1px; font-weight: 800; }
    .market-flag { font-size: 22px; }
    .market-tag {
        font-size: 10px; font-weight: bold; text-transform: uppercase; letter-spacing: .5px;
        color: var(--white); background: var(--agco-red); padding: 4px 10px; border-radius: 2px;
    }
    .detal-card { background: #fff; border: 1px solid var(--border-color); border-top: 5px solid var(--agco-red); border-radius: var(--radius); overflow: hidden; }
    .detal-table { width: 100%; border-collapse: collapse; }
    .detal-table td, .detal-table th { padding: 12px 14px; border-bottom: 1px solid var(--border-color); vertical-align: top; text-align: left; }
    .detal-table th { background: var(--agco-light-gray); font-size: 11px; text-transform: uppercase; letter-spacing: .5px; color: var(--agco-black); font-weight: bold; }
    .na-state { padding: 18px 20px; color: var(--text-muted); font-size: 13px; }
    .seg-cell { font-weight: bold; color: var(--agco-black); font-size: 12px; text-transform: uppercase; white-space: nowrap; }
    .var-cell { line-height: 2.1; }
    .count-cell { text-align: center; }
    .count-badge {
        display: inline-block; min-width: 28px; font-weight: bold; font-size: 12px;
        color: var(--agco-black); background: var(--agco-light-gray);
        border: 1px solid var(--border-color); border-radius: 4px; padding: 3px 9px;
    }
    .vpill {
        display: inline-block; font-size: 11px; font-weight: bold; padding: 4px 10px;
        border-radius: 4px; margin: 0 6px 5px 0; white-space: nowrap;
        text-transform: uppercase; letter-spacing: .3px; border: 1px solid transparent;
    }
    .vp-temporal { color: var(--farol-warning-text); background: var(--farol-warning-bg); border-color: #f0dca0; }
    .vp-agro     { color: var(--farol-positive-text); background: var(--farol-positive-bg); border-color: #c6e0b4; }
    .vp-macro    { color: var(--farol-critical-text); background: var(--farol-critical-bg); border-color: #f2c2a8; }

    @media (max-width: 992px) {
        .steps { grid-template-columns: 1fr 1fr; }
        .insights-container { grid-template-columns: 1fr; }
        .dashboard-header { flex-direction: column; align-items: flex-start; }
    }
    @media print { body { background: #fff; } .lang-controls, .btn-excel { display: none !important; } }
"""

def gerar_dashboard():
    bases = carregar_bases()
    if not bases:
        print("❌ Nenhuma base de forecast (.xlsx) encontrada na pasta.")
        return
    now = datetime.now()
    current_year = now.year
    versoes = list(bases.keys())
    versao_rec = versoes[-1]
    versao_ant = versoes[0] if len(versoes) > 1 else None
    cols = validar_e_mapear_colunas(bases[versao_rec])
    col_m, col_p, col_s, col_y, col_v = cols
    if not all([col_m, col_p, col_s, col_y, col_v]):
        print("\n❌ ERRO: Não foi possível mapear todas as colunas obrigatórias no Excel.")
        return
    dados = []
    for versao, df in bases.items():
        temp = df.copy(); temp['Versao'] = versao; dados.append(temp)
    df_full = pd.concat(dados, ignore_index=True).sort_values(by=[col_m, col_p, col_s, col_y])
    mercados = sorted(list(df_full[col_m].dropna().unique()))
    if not mercados:
        print("\n❌ ERRO CRÍTICO: Nenhum dado de país/mercado foi encontrado.")
        sys.exit(1)
    anos_unicos = sorted(list({str(int(y)) if isinstance(y, (int, float)) and pd.notna(y) else str(y) for y in df_full[col_y].unique()}))
    mes_keys = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"]
    mes_key = mes_keys[now.month - 1]
    mes_nome_ano = f"{i18n(mes_key,'pt')} de {now.year}"
    insights_agribusiness = obter_insights_agribusiness()
    total_atual = df_full[df_full['Versao'] == versao_rec][col_v].sum()
    total_atual_str = f"{total_atual:,.0f}".replace(",", ".")
    if versao_ant:
        total_ant = df_full[df_full['Versao'] == versao_ant][col_v].sum()
        var_pct = ((total_atual / total_ant) - 1) * 100 if total_ant > 0 else 0
        sinal = "+" if var_pct >= 0 else ""
        cor_kpi = "positive" if var_pct >= 0 else "negative"
        texto_var = f"{sinal}{var_pct:.1f}% vs {versao_ant}"
    else:
        print("\nℹ️  AVISO: Apenas uma versão de forecast foi encontrada.")
        texto_var = "-"; cor_kpi = "info"
    cor_texto = {"positive": "#28a745", "negative": "#dc3545", "info": "#6c757d"}.get(cor_kpi, "#6c757d")

    injected_css = CSS_METODOLOGIA

    for br_tag in ['BRASIL', 'BRA']:
        if br_tag in mercados:
            mercados.insert(0, mercados.pop(mercados.index(br_tag))); break

    html_country_tabs, html_country_contents = "", ""
    for idx, mercado in enumerate(mercados):
        df_mercado = df_full[df_full[col_m] == mercado]
        print(f"\n--- Processando País: {mercado} ---"); sys.stdout.flush()
        t, c = gerar_conteudo_pais(mercado, df_mercado, idx == 0, anos_unicos, versoes, current_year, mes_key, cols, insights_agribusiness, mes_nome_ano)
        html_country_tabs += t; html_country_contents += c

    # ABA METODOLOGIA (Matriz de Fatores foi removida)
    mt, mc = gerar_conteudo_metodologia()
    html_country_tabs += mt; html_country_contents += mc

    # Script i18n: mostra .i18n-text no idioma inicial e re-sincroniza após clique nos botões PT/EN/ES do template
    i18n_init = """
    <script>
    (function(){
      function showLang(lang){
        var all=document.querySelectorAll('.i18n-text');
        for(var i=0;i<all.length;i++){all[i].style.display='none';}
        var sel=document.querySelectorAll('.i18n-text.lang-'+lang);
        for(var j=0;j<sel.length;j++){sel[j].style.display='inline';}
      }
      function setActive(lang){
        var b=document.querySelectorAll('.lang-btn');
        for(var i=0;i<b.length;i++){var oc=b[i].getAttribute('onclick')||'';
          if(oc.indexOf("'"+lang+"'")>-1){b[i].classList.add('active');}else{b[i].classList.remove('active');}}
      }
      function boot(){
        if(typeof window.setLanguage==='function' && !window.__wrapLang){
          var _o=window.setLanguage;
          window.setLanguage=function(l){try{_o.call(this,l);}catch(e){} showLang(l); setActive(l); try{window.currentLang=l;}catch(e){}};
          window.__wrapLang=true;
        }
        var lang='pt'; try{lang=window.currentLang||'pt';}catch(e){}
        showLang(lang); setActive(lang);
      }
      if(document.readyState==='loading'){document.addEventListener('DOMContentLoaded',boot);}else{boot();}
    })();
    </script>"""
    html_country_contents += i18n_init

    context = {
        'injected_css': injected_css,
        'translations_json': json.dumps(TRANSLATIONS, ensure_ascii=False),  # <-- CORREÇÃO CRÍTICA (faltava)
        'all_charts_json': json.dumps({}),
        'forecast_version': str(versao_rec),
        'update_date': now.strftime("%d/%m/%Y"),
        'total_atual_str': str(total_atual_str),
        'cor_kpi': str(cor_kpi),
        'cor_texto': str(cor_texto),  # <-- CORREÇÃO (faltava)
        'texto_var': str(texto_var),
        'html_country_tabs': html_country_tabs,
        'html_country_contents': html_country_contents,
    }
    renderizar_template_final(context)
    print(f"\n✅ Sucesso! Dashboard gerado em: {CAMINHO_HTML}")

if __name__ == "__main__":
    gerar_dashboard()
