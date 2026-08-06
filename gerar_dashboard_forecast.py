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
    print("   2. Ative o ambiente virtual com o comando:")
    print("      > .\\venv\\Scripts\\Activate.ps1")
    print("   3. Após a ativação, o seu prompt de comando deve começar com '(venv)'.")
    print("   4. Rode o script novamente:")
    print("      > python gerar_dashboard_forecast.py")
    print("="*80 + "\n")
    sys.exit(1)

# --- CARREGAR VARIÁVEIS DE AMBIENTE ---
load_dotenv(override=True)
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

# --- TRANSLATIONS (PT / EN / ES) ---
TRANSLATIONS = {
    "pt": {
        "realized": "Realizado", "var_cycle": "Var. Ciclo", "seg": "🚜 Segmento", "total": "Total:",
        "proj_vol": "Projeção de Volumes", "export": "Exportar para Excel",
        "version": "Versão", "updated": "Atualizado",
        "select_prod": "Selecione a Família de Produto:", "month": "Mês", "current_forecast": "Atual Forecast",
        "positive": "Positivo", "critical": "Crítico", "warning": "Atenção",
        "var_yoy": "Var. YoY",
        "positivo": "Positivo", "negativo": "Negativo", "neutro": "Neutro", "incerto": "Incerto",
        "insights_agribusiness_title": "Insights de Agribusiness (IA)",
        "jan": "Janeiro", "feb": "Fevereiro", "mar": "Março", "apr": "Abril", "may": "Maio", "jun": "Junho", "jul": "Julho", "aug": "Agosto", "sep": "Setembro", "oct": "Outubro", "nov": "Novembro", "dec": "Dezembro", "current": "Atual",
        "scenario_general": "Cenário Geral do Mercado", "relevant_news": "Notícias Relevantes",
        "brasil": "Brasil", "argentina": "Argentina", "mexico": "México", "osa": "OSA",
        "ta": "Tratores", "co": "Colheitadeiras", "pa": "Plantadeiras", "pu": "Pulverizadores",
        "metodologia": "Metodologia",
    },
    "en": {
        "realized": "Realized", "var_cycle": "Cycle Var.", "seg": "🚜 Segment", "total": "Total:",
        "proj_vol": "Volume Projection", "export": "Export to Excel",
        "version": "Version", "updated": "Updated",
        "select_prod": "Select Product Family:", "month": "Month", "current_forecast": "Current Forecast",
        "positive": "Positive", "critical": "Critical", "warning": "Warning",
        "var_yoy": "YoY Var.",
        "positivo": "Positive", "negativo": "Negative", "neutro": "Neutral", "incerto": "Uncertain",
        "insights_agribusiness_title": "Agribusiness Insights (AI)",
        "jan": "January", "feb": "February", "mar": "March", "apr": "April", "may": "May", "jun": "June", "jul": "July", "aug": "August", "sep": "September", "oct": "October", "nov": "November", "dec": "December", "current": "Current",
        "scenario_general": "General Market Scenario", "relevant_news": "Relevant News",
        "brasil": "Brazil", "argentina": "Argentina", "mexico": "Mexico", "osa": "OSA",
        "ta": "Tractors", "co": "Combines", "pa": "Planters", "pu": "Sprayers",
        "metodologia": "Methodology",
    },
    "es": {
        "realized": "Realizado", "var_cycle": "Var. Ciclo", "seg": "🚜 Segmento", "total": "Total:",
        "proj_vol": "Proyección de Volúmenes", "export": "Exportar a Excel",
        "version": "Versión", "updated": "Actualizado",
        "select_prod": "Seleccione Familia de Producto:", "month": "Mes", "current_forecast": "Forecast Actual",
        "positive": "Positivo", "critical": "Crítico", "warning": "Atención",
        "var_yoy": "Var. YoY",
        "positivo": "Positivo", "negativo": "Negativo", "neutro": "Neutro", "incerto": "Incierto",
        "insights_agribusiness_title": "Insights de Agronegocios (IA)",
        "jan": "Enero", "feb": "Febrero", "mar": "Marzo", "apr": "Abril", "may": "Mayo", "jun": "Junio", "jul": "Julio", "aug": "Agosto", "sep": "Septiembre", "oct": "Octubre", "nov": "Noviembre", "dec": "Diciembre", "current": "Actual",
        "scenario_general": "Escenario General del Mercado", "relevant_news": "Noticias Relevantes",
        "brasil": "Brasil", "argentina": "Argentina", "mexico": "México", "osa": "OSA",
        "ta": "Tractores", "co": "Cosechadoras", "pa": "Sembradoras", "pu": "Pulverizadores",
        "metodologia": "Metodología",
    }
}

def i18n(key, lang='pt'):
    return TRANSLATIONS.get(lang, TRANSLATIONS['pt']).get(key, key)

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
        traceback.print_exc()
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
        html_resumo = "<p>Nenhum resumo geral disponível para este período.</p>"
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
            desc_span = f'<span data-lang-pt="{html.escape(desc_pt)}" data-lang-en="{html.escape(desc_en)}" data-lang-es="{html.escape(desc_es)}">{desc_pt}</span>'
            resumo_parts.append(f'<div class="summary-item {sentiment_class}"><span class="summary-item-icon">{fator.get("icone", "📊")}</span><div class="summary-item-text"><strong>{fator.get("titulo", "")}:</strong> {desc_span}</div>{badge_html}</div>')
        html_resumo = "".join(resumo_parts)
    html_noticias = ""
    if not noticias:
        html_noticias = "<p>Nenhuma notícia relevante encontrada para este período.</p>"
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
            cor_sentimento = {'positivo': 'var(--color-hp)', 'negativo': 'var(--color-hn)', 'neutro': 'var(--color-n)'}.get(sentimento, 'var(--color-n)')
            sentimento_texto = i18n(sentimento)
            fonte_pais_html = f'<span class="news-source-country">{noticia["fonte_pais"]}</span>' if 'fonte_pais' in noticia else ""
            html_noticias += f"""
            <div class="news-card {card_class}">
                <div class="news-card-header">
                    <span class="news-sentiment-dot" style="background-color: {cor_sentimento};"></span>
                    <span class="news-sentiment-label" style="color: {cor_sentimento};" data-i18n="{sentimento}">{sentimento_texto}</span>
                    <h5 class="news-title" data-lang-pt="{html.escape(titulo_pt)}" data-lang-en="{html.escape(titulo_en)}" data-lang-es="{html.escape(titulo_es)}">{titulo_pt}</h5>
                </div>
                <p class="news-summary"><span data-lang-pt="{html.escape(resumo_pt)}" data-lang-en="{html.escape(resumo_en)}" data-lang-es="{html.escape(resumo_es)}">{resumo_pt}</span>{fonte_pais_html}</p>
            </div>
            """
    return f"""
    <div class="insights-section">
        <hr class="agco-divider">
        <h3 class="section-title" data-i18n="insights_agribusiness_title">
            {i18n("insights_agribusiness_title")}
            <span class="month-badge" data-i18n="{mes_key}">{mes_nome_ano}</span>
        </h3>
        <div class="insights-container">
            <div class="insight-summary-card">
                <h4 data-i18n="scenario_general">{i18n("scenario_general")}</h4>
                <div class="summary-content">{html_resumo}</div>
            </div>
            <div class="insight-news-container">
                <h4 data-i18n="relevant_news">{i18n("relevant_news")}</h4>
                <div class="news-grid">{html_noticias}</div>
            </div>
        </div>
    </div>
    """

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
        forecasts = 12 - actuals
        versao_atual = f"{actuals}+{forecasts}"
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
    # --- SELEÇÃO DE VERSÕES PARA COMPARAÇÃO ---
    todas_versoes_ordenadas = list(bases_temp.keys())
    ultimas_versoes = []
    if todas_versoes_ordenadas:
        versao_mais_recente = todas_versoes_ordenadas[-1]
        # REGRA ESPECIAL: Se a versão atual for '7+5', força a comparação com '6+6'
        if versao_mais_recente == '7+5':
            print("\nℹ️  REGRA ESPECIAL: Detectada versão '7+5'. Comparando com '6+6'.")
            sys.stdout.flush()
            if '6+6' in todas_versoes_ordenadas:
                ultimas_versoes = ['6+6', '7+5']
            else:
                print("   ⚠️ AVISO: Versão '6+6' não encontrada. Usando a versão anterior disponível.")
                sys.stdout.flush()
                ultimas_versoes = todas_versoes_ordenadas[-2:] if len(todas_versoes_ordenadas) > 1 else todas_versoes_ordenadas[-1:]
        else:
            ultimas_versoes = todas_versoes_ordenadas[-2:] if len(todas_versoes_ordenadas) > 1 else todas_versoes_ordenadas[-1:]
    bases = {v: pd.read_excel(bases_temp[v]) for v in ultimas_versoes}
    limpar_historico_antigo(hist_dir)
    return bases

def renderizar_template_final(context):
    template_path = os.path.join(BASE_DIR, 'template.html')
    if not os.path.exists(template_path): return
    with open(template_path, 'r', encoding='utf-8') as f:
        html_completo = f.read()
    for key, value in context.items():
        if key == 'injected_css': continue
        html_completo = html_completo.replace(f'{{{key}}}', str(value))
    if 'injected_css' in context:
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
    produto_traduzido = i18n(produto_key_i18n)
    html_prod_tab = f'<button class="prod-tab-btn {p_active_class}" onclick="openProduct(event, \'{prod_id}\', \'{mercado_limpo}\')" data-i18n="{produto_key_i18n}">{produto_traduzido}</button>'
    segmentos_unicos = ordenar_segmentos(df_segmentos[col_s].unique())
    th_style_bottom = 'border-bottom: 1px solid #e2e8f0; color: #1e293b; font-size: 13px; font-weight: 600;'
    th_style_sub = 'font-size: 11px; color: #64748b; font-weight: 500;'
    th_style_sub_bold = 'font-size: 11px; color: #0f172a; font-weight: 600;'
    if versao_ant:
        header_row1 = f'<th rowspan="2" style="width: 15%; border-bottom: 1px solid #e2e8f0; color: #1e293b;" data-i18n="seg">{i18n("seg")}</th>'
        header_row2 = ''
        for i, ano in enumerate(anos_unicos):
            is_past = int(ano) < current_year if ano.isdigit() else False
            add_yoy_column = False
            if not is_past and i > 0:
                add_yoy_column = True
            if is_past:
                header_row1 += f'<th style="text-align: center; {th_style_bottom}">{ano}</th>'
                header_row2 += f'<th style="text-align: center; {th_style_sub} border-bottom: 1px solid #e2e8f0;" data-i18n="realized">{i18n("realized")}</th>'
            elif add_yoy_column:
                ano_anterior = anos_unicos[i-1]
                header_row1 += f'<th colspan="2" style="text-align: center; {th_style_bottom}">{ano}</th>'
                header_row2 += f'<th style="text-align: center; {th_style_sub_bold} width: 12%; border-bottom: 1px solid #e2e8f0;">{versao_rec}</th>'
                header_row2 += f'<th style="text-align: center; {th_style_sub} width: 12%; border-bottom: 1px solid #e2e8f0;" data-i18n="var_yoy">Var. YoY ({ano[-2:]} vs {ano_anterior[-2:]})</th>'
            else:
                header_row1 += f'<th colspan="1" style="text-align: center; {th_style_bottom}">{ano}</th>'
                header_row2 += f'<th style="text-align: center; {th_style_sub_bold} width: 15%; border-bottom: 1px solid #e2e8f0;">{versao_rec}</th>'
        html_thead = f'<thead><tr>{header_row1}</tr><tr>{header_row2}</tr></thead>'
    else:
        header_row1 = f'<th style="width: 20%; color: #1e293b;" data-i18n="seg">{i18n("seg")}</th>' + ''.join([f'<th style="text-align: center; font-weight: 600; color: #1e293b;">{ano}</th>' for ano in anos_unicos])
        for i in range(1, len(anos_unicos)):
            header_row1 += f'<th style="text-align: center; background-color: #f8fafc; font-weight: 600; color: #1e293b;">Var. % ({anos_unicos[i-1]} ➔ {anos_unicos[i]})</th>'
        html_thead = f'<thead><tr>{header_row1}</tr></thead>'
    linhas_tabela_dados = ""
    totais_ano = {ano: {v: 0 for v in versoes} for ano in anos_unicos}
    for segmento in segmentos_unicos:
        df_seg_data = df_segmentos[df_segmentos[col_s] == segmento]
        tds_anos, vals = "", []
        v_rec_val_ano_anterior = 0
        for i, ano in enumerate(anos_unicos):
            v_rec_val = df_seg_data[(df_seg_data[col_y].astype(str) == str(ano)) & (df_seg_data['Versao'] == versao_rec)][col_v].sum()
            totais_ano[ano][versao_rec] += v_rec_val
            if versao_ant:
                v_ant_val = df_seg_data[(df_seg_data[col_y].astype(str) == str(ano)) & (df_seg_data['Versao'] == versao_ant)][col_v].sum()
                totais_ano[ano][versao_ant] += v_ant_val
                is_past = int(ano) < current_year if ano.isdigit() else False
                add_yoy_column = False
                if not is_past and i > 0:
                    add_yoy_column = True
                if is_past:
                    tds_anos += f'<td class="num" style="font-weight: 500; color: #334155; background-color: #f8fafc; border-left: 1px dashed #e2e8f0;">{v_rec_val:,.0f}</td>'
                else:
                    tds_anos += f'<td class="num" style="font-weight: 500; color: #0f172a; background-color: #f8fafc; border-left: 1px dashed #e2e8f0;">{v_rec_val:,.0f}</td>'
                    if add_yoy_column:
                        var_yoy_pct = ((v_rec_val / v_rec_val_ano_anterior) - 1) * 100 if v_rec_val_ano_anterior > 0 else None
                        if var_yoy_pct is not None:
                            var_yoy_class = "positive" if var_yoy_pct >= 0 else "negative"
                            sinal_yoy = "+" if var_yoy_pct > 0 else ""
                            var_yoy_str = f'<span class="{var_yoy_class}" style="font-weight: 500;">{sinal_yoy}{var_yoy_pct:.1f}%</span>'
                        else:
                            var_yoy_str = f'<span class="positive" style="font-weight: 500;">+100.0%</span>' if v_rec_val > 0 else "-"
                        tds_anos += f'<td class="num">{var_yoy_str}</td>'
            else:
                tds_anos += f'<td class="num" style="font-weight: 500; color: #0f172a;">{v_rec_val:,.0f}</td>'
                vals.append(v_rec_val)
            v_rec_val_ano_anterior = v_rec_val
        linhas_tabela_dados += f'<tr><td style="font-weight: 600; color: #334155;">{segmento}</td>{tds_anos}</tr>'
    tds_totais = ""
    total_rec_ano_anterior = 0
    for i, ano in enumerate(anos_unicos):
        t_rec = totais_ano[ano][versao_rec]
        if versao_ant:
            is_past = int(ano) < current_year if ano.isdigit() else False
            add_yoy_column = False
            if not is_past and i > 0:
                add_yoy_column = True
            if is_past:
                tds_totais += f'<td class="num" style="font-weight: 700; font-size: 1.05em; color: #0f172a; background-color: #f8fafc; border-left: 1px dashed #e2e8f0;">{t_rec:,.0f}</td>'
            else:
                tds_totais += f'<td class="num" style="font-weight: 700; font-size: 1.05em; color: #0f172a; background-color: #f8fafc; border-left: 1px dashed #e2e8f0;">{t_rec:,.0f}</td>'
                if add_yoy_column:
                    var_yoy_pct = ((t_rec / total_rec_ano_anterior) - 1) * 100 if total_rec_ano_anterior > 0 else None
                    if var_yoy_pct is not None:
                        var_yoy_class = "positive" if var_yoy_pct >= 0 else "negative"
                        sinal_yoy = "+" if var_yoy_pct > 0 else ""
                        var_yoy_str = f'<span class="{var_yoy_class}">{sinal_yoy}{var_yoy_pct:.1f}%</span>'
                    else:
                        var_yoy_str = f'<span class="positive">+100.0%</span>' if t_rec > 0 else "-"
                    tds_totais += f'<td class="num" style="font-weight: 700; font-size: 1em;">{var_yoy_str}</td>'
        total_rec_ano_anterior = t_rec
    body_content = f"""
        <div class="header-action-container">
            <h3 class="section-title" data-i18n="proj_vol">{i18n("proj_vol")}</h3>
            <button class="btn-excel elegant-btn" onclick="exportProdExcel('{prod_id}', '{mercado}_{produto}')">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                <span data-i18n="export">{i18n("export")}</span>
            </button>
        </div>
        <div class="table-container elegant-card">
            <table id="table_{prod_id}">
                {html_thead}
                <tbody>{linhas_tabela_dados}</tbody>
                <tfoot><tr style="background-color: #f1f5f9; border-top: 1px solid #cbd5e1;"><td style="text-align: left; font-weight: 700; color: #0f172a; text-transform: uppercase;" data-i18n="total">{i18n("total")}</td>{tds_totais}</tr></tfoot>
            </table>
        </div>
    """
    html_prod_content = f"""
    <div id="{prod_id}" class="prod-tab-content {p_active_class}" style="display: {p_display_style};">
        {body_content}
    </div>
    """
    return html_prod_tab, html_prod_content

def gerar_conteudo_pais(mercado, df_mercado, is_active, anos_unicos, versoes, current_year, mes_key, cols, insights_agribusiness, mes_nome_ano):
    col_m, col_p, col_s, col_y, col_v = cols
    mercado_limpo = str(mercado).replace(" ", "_").upper()
    active_class = "active" if is_active else ""
    display_style = "block" if is_active else "none"
    mercado_key_i18n = str(mercado).lower().replace(' ', '_')
    mercado_traduzido = i18n(mercado_key_i18n)
    html_country_tab = f'<button class="tab-btn country-tab-btn {active_class}" onclick="openCountry(event, \'{mercado_limpo}\')" data-i18n="{mercado_key_i18n}">{mercado_traduzido}</button>'
    ordem_produtos = ['TA', 'CO', 'PA', 'PU']
    produtos = sorted(df_mercado[col_p].unique(), key=lambda p: next(((i, p) for i, prefix in enumerate(ordem_produtos) if str(p).upper().startswith(prefix)), (len(ordem_produtos), p)))
    html_prod_tabs, html_prod_contents = "", ""
    for p_idx, produto in enumerate(produtos):
        df_produto = df_mercado[df_mercado[col_p] == produto]
        tab, content = gerar_conteudo_produto(produto, df_produto, p_idx, mercado, mercado_limpo, anos_unicos, versoes, current_year, mes_key, cols)
        html_prod_tabs += tab
        html_prod_contents += content
    mapa_mercado = {'BRASIL': 'BRASIL', 'BRA': 'BRASIL', 'ARGENTINA': 'ARGENTINA', 'ARG': 'ARGENTINA', 'MEXICO': 'MEXICO', 'MEX': 'MEXICO', 'OSA': 'OSA'}
    mercado_key = mapa_mercado.get(mercado.upper(), 'OSA')
    insights_para_mercado = insights_agribusiness.get(mercado_key, {})
    html_insights = gerar_html_insights_agribusiness(insights_para_mercado, mes_nome_ano, mes_key)
    html_country_content = f"""
    <div id="{mercado_limpo}" class="country-tab-content {active_class}" style="display: {display_style};">
        <div class="prod-tabs-container">
            <h2 class="country-title" data-i18n="select_prod">{i18n("select_prod")}</h2>
            <div class="prod-tabs-nav">{html_prod_tabs}</div>
        </div>
        {html_prod_contents}
        {html_insights}
    </div>"""
    return html_country_tab, html_country_content

# ==========================================================================
# ABA METODOLOGIA (XGBoost) - documenta o método e as features por regiao/produto/segmento
# ==========================================================================
METODO_FEATURES = {
    "TA": {"label": "Tratores", "icon": "🚜", "data": {
        "BRASIL": {
            "0-49 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Feijão", "Leite", "Mandioca", "Suínos", "Tomate", "Uva", "Taxa Juros (%)"],
            "50-79 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Café", "Laranja", "Leite", "Milho", "Suínos", "Tomate", "Maçã", "Taxa Juros (%)", "Inflação (%)"],
            "80-119 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Arroz VP", "Bovinos VP", "Leite VP", "Milho VP", "Soja VP", "Trigo VP", "Taxa Juros (%)", "Inflação (%)", "Taxa Câmbio (USD)"],
            "120-169 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Milho VP", "Soja VP", "Wheat Value of Production", "Taxa Juros (%)", "Taxa Câmbio (USD)", "Preço Soja (USD/Ton)"],
            "170-239 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Milho VP", "Soja VP", "Taxa Juros (%)", "Preço Soja (USD/Ton)"],
            "240-339 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Milho VP", "Soja VP", "Corn Production", "Corn Wholesale Price"],
            "+340 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Milho VP", "Soja VP", "Taxa Juros (%)", "Inflação (%)", "Taxa Câmbio (USD)", "Preço Soja (USD/Ton)"],
        },
        "ARGENTINA": {
            "0-49 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Apple PR", "Tomato Production", "Potato PR", "Taxa Juros (%)", "Inflação (%)"],
            "50-79 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Apple PR", "Beef and Veal Meat Domestic", "Tomato Production", "Olive Producer Price", "Peach & Nectarine PR", "Taxa Juros (%)", "Inflação (%)"],
            "80-119 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Beef and Veal Meat Domestic", "Wheat Value of Production", "Corn Port Price"],
            "120-169 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Wheat Value of Production", "Corn Value of Production", "Soybean farm prices", "Taxa Juros (%)"],
            "170-239 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Wheat Value of Production", "Corn Value of Production", "Soybean oil production"],
            "240-339 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Wheat Value of Production", "Corn Value of Production", "Soybean PR"],
            "+340 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Wheat Value of Production", "Corn Value of Production", "Soybean oil production"],
        },
        "OSA": {
            "0-49 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Apple PR", "Coffee Area Harvested", "Coffee Producer Price", "Coffee Production", "Olive Producer Price", "Orange Wholesale Price", "Peach & Nectarine PR"],
            "50-79 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Apple PR", "Coffee Area Harvested", "Coffee Producer Price", "Coffee Production", "Olive Producer Price", "Orange Wholesale Price", "Peach & Nectarine PR"],
            "80-119 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Beef and Veal Meat Domestic", "Soybean PR"],
            "120-169 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Soybean oil production", "Soybean PR"],
            "170-239 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Soybean oil production", "Soybean PR"],
            "240-339 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Soybean oil production", "Soybean PR"],
            "+340 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Soybean oil production", "Soybean PR"],
        },
        "MEXICO": {
            "0-49 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Avocado Area Harvested", "Corn Area Harvested"],
            "50-79 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Avocado Area Harvested", "Corn Area Harvested"],
            "80-119 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Sugar Cane Area Harvested"],
            "120-169 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Corn Production"],
            "170-239 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Sugar Cane Area Harvested"],
            "240-339 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Sugar Cane Area Harvested"],
            "+340 HP": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Sugar Cane Area Harvested"],
        },
    }},
    "CO": {"label": "Colheitadeiras", "icon": "🌾", "data": {
        "BRASIL": {
            "Classe 4": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Arroz VP", "Milho VP", "Soja VP", "Trigo VP", "Corn Wholesale Price", "Taxa Juros (%)", "Soybean port prices"],
            "Classe 5": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Arroz VP", "Milho VP", "Soja VP", "Trigo VP", "Corn Wholesale Price", "Taxa Juros (%)", "Soybean port prices"],
            "Classe 6": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Milho VP", "Soja VP", "Taxa Juros (%)"],
            "Classe 7": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Algodão VP", "Milho VP", "Soybean VP", "Taxa Juros (%)"],
            "Classe 8+": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Algodão VP", "Milho VP", "Soybean VP", "Taxa Juros (%)"],
        },
        "ARGENTINA": {
            "Classe 4 e 5": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Wheat Value of Production", "Corn Area Harvested", "Corn Port Price", "Corn Production", "Corn Value of Production", "Corn Wholesale Price", "Taxa Juros (%)", "Índice Crédito Agrícola", "Soybean oil production"],
            "Classe 6": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Wheat Value of Production", "Corn Port Price", "Corn Value of Production", "Soybean oil production"],
            "Classe 7": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Wheat Value of Production", "Corn Value of Production", "Soybean oil production"],
            "Classe 8+": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Wheat Value of Production", "Corn Value of Production", "Soybean oil production"],
        },
        "OSA": {
            "Classe 4": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Rice VP"],
            "Classe 5": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Rice VP", "Soybean PR"],
            "Classe 6": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Rice VP", "Soybean oil production"],
            "Classe 7": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Soybeans production"],
            "Classe 8+": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Soybeans production"],
        },
    }},
    "PA": {"label": "Plantadeiras", "icon": "🌱", "data": {
        "BRASIL": {
            "< 20 linhas": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Milho VP", "Soja VP", "Trigo VP", "Taxa Juros (%)", "Preço Soja (USD/Ton)"],
            "> 20 linhas": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Milho VP", "Soja VP", "Taxa Juros (%)", "Soybean port prices"],
        },
        "OSA": {"Todos os segmentos": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Soybeans PR"]},
    }},
    "PU": {"label": "Pulverizadores", "icon": "💧", "data": {
        "BRASIL": {"Todos os segmentos": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Cana-de-Açúcar VP", "Milho VP", "Soja VP", "Taxa Juros (%)", "Taxa Câmbio (USD)"]},
        "ARGENTINA": {"Todos os segmentos": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Wheat Value of Production", "Corn Value of Production", "Soybean PR", "Taxa Juros (%)"]},
        "OSA": {"Todos os segmentos": ["Sazonalidade (ano)", "Sazonalidade (mês)", "Soybean PR"]},
    }},
}
METODO_PRODUCT_ORDER = ["TA", "CO", "PA", "PU"]
METODO_MARKET_ORDER = ["BRASIL", "ARGENTINA", "OSA", "MEXICO"]
METODO_MARKET_META = {
    "BRASIL": {"flag": "🇧🇷", "label": "Brasil"},
    "ARGENTINA": {"flag": "🇦🇷", "label": "Argentina"},
    "OSA": {"flag": "🌎", "label": "OSA (Outros da América do Sul)"},
    "MEXICO": {"flag": "🇲🇽", "label": "México"},
}

def _metodo_classify(var):
    v = var.lower()
    if "sazonalidade" in v:
        return "temporal"
    macro_kw = ["juros", "inflação", "inflacao", "câmbio", "cambio", "crédito", "credito", "índice", "indice"]
    price_kw = ["price", "preço", "preco", "usd", "wholesale", "port price", "farm price", "port prices"]
    if any(k in v for k in macro_kw):
        return "macro"
    if any(k in v for k in price_kw) or v.endswith(" pr"):
        return "macro"
    return "agro"

def _metodo_pill(var):
    cat = {"temporal": "vp-temporal", "agro": "vp-agro", "macro": "vp-macro"}[_metodo_classify(var)]
    return f'<span class="vpill {cat}">{html.escape(var)}</span>'

def _metodo_market_block(pcode, market):
    p = METODO_FEATURES[pcode]
    meta = METODO_MARKET_META[market]
    data = p["data"].get(market)
    if not data:
        return (f'<div class="market-block"><div class="market-head"><span class="market-flag">{meta["flag"]}</span>'
                f'<h4>{meta["label"]}</h4><span class="market-tag">{p["label"]}</span></div>'
                f'<div class="elegant-card detal-card"><div class="na-state">➖ <strong>Não aplicável.</strong> '
                f'Este mercado ({meta["label"]}) não comercializa a família <strong>{p["label"]}</strong> — '
                f'portanto, não há forecast para este produto.</div></div></div>')
    rows = ""
    for seg, varlist in data.items():
        pills = "".join(_metodo_pill(v) for v in varlist)
        rows += (f'<tr><td class="seg-cell">{html.escape(seg)}</td><td class="var-cell">{pills}</td>'
                 f'<td class="count-cell"><span class="count-badge">{len(varlist)}</span></td></tr>')
    return (f'<div class="market-block"><div class="market-head"><span class="market-flag">{meta["flag"]}</span>'
            f'<h4>{meta["label"]}</h4><span class="market-tag">{p["label"]}</span></div>'
            f'<div class="elegant-card detal-card"><table class="detal-table">'
            f'<thead><tr><th style="width:16%">Segmento</th><th>Variáveis do modelo (features do XGBoost)</th>'
            f'<th style="width:8%;text-align:center">Nº</th></tr></thead><tbody>{rows}</tbody></table></div></div>')

def _metodo_product_panel(pcode, active):
    blocks = "".join(_metodo_market_block(pcode, m) for m in METODO_MARKET_ORDER)
    disp = "block" if active else "none"
    return f'<div id="detal-{pcode}" class="detal-panel" style="display:{disp}">{blocks}</div>'

def _metodo_detalhamento():
    legend = ('<div class="detal-legend"><span class="legend-title">Natureza das variáveis:</span>'
              '<span class="vpill vp-temporal">Temporal</span>'
              '<span class="vpill vp-agro">Agronômica / Setorial</span>'
              '<span class="vpill vp-macro">Macroeconômica &amp; Preços</span></div>')
    intro = ('<div class="method-card"><p>Abaixo, o <strong>conjunto de variáveis (features)</strong> que alimenta o modelo '
             '<span class="model-name">XGBoost</span> em cada <strong>segmento</strong>, por mercado. São os elementos de maior '
             'importância preditiva por família de produto. Todos os modelos incluem a <strong>sazonalidade</strong> (ano e mês) '
             'como base temporal, somada a drivers agronômicos/setoriais e macroeconômicos específicos de cada segmento.</p></div>')
    prod_tabs = "".join(
        f'<button class="detal-prod-btn {"active" if i == 0 else ""}" onclick="openProd(event, \'{c}\')">{METODO_FEATURES[c]["icon"]} {METODO_FEATURES[c]["label"]}</button>'
        for i, c in enumerate(METODO_PRODUCT_ORDER))
    panels = "".join(_metodo_product_panel(c, i == 0) for i, c in enumerate(METODO_PRODUCT_ORDER))
    return (f'<h3 class="section-title" style="margin-top:1.6rem">Detalhamento por Região, Produto e Segmento</h3>'
            f'<hr class="agco-divider">{intro}{legend}<div class="detal-prod-nav">{prod_tabs}</div>{panels}'
            f'<div class="next" style="margin-top:1.2rem">✅ <strong>Detalhamento completo:</strong> as quatro famílias '
            f'(Tratores, Colheitadeiras, Plantadeiras e Pulverizadores) estão mapeadas por mercado e segmento. Onde um mercado '
            f'não comercializa determinada família, o item aparece como <em>"Não aplicável"</em>.</div>')

def gerar_conteudo_metodologia():
    tab = f'<button class="tab-btn country-tab-btn" onclick="openCountry(event, \'METODOLOGIA\')" data-i18n="metodologia">{i18n("metodologia")}</button>'
    open_prod_script = (
        "<script>"
        "function openProd(evt, code){"
        "document.querySelectorAll('.detal-panel').forEach(function(el){el.style.display='none';});"
        "document.querySelectorAll('.detal-prod-btn').forEach(function(b){b.classList.remove('active');});"
        "var t=document.getElementById('detal-'+code); if(t){t.style.display='block';}"
        "if(evt&&evt.currentTarget){evt.currentTarget.classList.add('active');}"
        "}</script>"
    )
    content = (
        '<div id="METODOLOGIA" class="country-tab-content" style="display: none;">'
        '<h3 class="section-title">Metodologia do Forecast <span class="method-pill">Modelo preditivo</span></h3>'
        '<hr class="agco-divider">'
        '<div class="method-card"><p class="lead">Todas as projeções deste dashboard — para <strong>todos os mercados</strong> '
        '(Brasil, Argentina, México e OSA) e <strong>todas as famílias de produto</strong> (Tratores, Colheitadeiras, '
        'Plantadeiras e Pulverizadores) — são geradas por um mesmo motor estatístico: o algoritmo '
        '<span class="model-name">XGBoost</span> (<em>eXtreme Gradient Boosting</em>).</p>'
        '<p>Adotar um único método para toda a base garante <strong>consistência metodológica</strong> e permite comparar '
        'mercados e segmentos sob o mesmo critério, sem vieses de modelos diferentes.</p></div>'
        '<div class="method-card"><h4 style="margin-top:0">O que é o XGBoost?</h4>'
        '<p>O XGBoost é um algoritmo de <strong>machine learning</strong> baseado em <strong>árvores de decisão</strong> '
        'combinadas pela técnica de <em>gradient boosting</em>. Em vez de treinar um único modelo, ele constrói '
        '<strong>centenas de pequenas árvores em sequência</strong>, onde cada nova árvore aprende a <strong>corrigir os erros</strong> '
        'deixados pelas anteriores. A previsão final é a soma das contribuições de todas as árvores.</p>'
        '<div class="callout">🏌️ <strong>Analogia do golfe:</strong> a primeira árvore dá a "tacada inicial" em direção ao alvo '
        '(a demanda real). Ela não acerta de primeira. A segunda árvore corrige a distância que faltou; a terceira dá um ajuste '
        'fino — e assim por diante, até chegar bem perto do valor correto.</div></div>'
        '<div class="method-card"><h4 style="margin-top:0">Como o modelo aprende — em 4 passos</h4><div class="steps">'
        '<div class="step"><div class="n">1</div><h5>Ponto de partida</h5><p>O modelo começa com uma estimativa-base (ex.: a média histórica de vendas do segmento).</p></div>'
        '<div class="step"><div class="n">2</div><h5>Mede o erro</h5><p>Compara a previsão com o realizado e calcula o resíduo — o quanto errou para mais ou para menos.</p></div>'
        '<div class="step"><div class="n">3</div><h5>Corrige</h5><p>Uma nova árvore é treinada focando exatamente nos pontos mais difíceis de prever.</p></div>'
        '<div class="step"><div class="n">4</div><h5>Repete e soma</h5><p>O ciclo se repete por centenas de rodadas; a projeção final é a soma de todas as árvores.</p></div>'
        '</div></div>'
        '<div class="method-card"><h4 style="margin-top:0">Por que escolhemos o XGBoost</h4><div class="feature-grid">'
        '<div class="feature"><span class="ic">🎯</span><div><b>Alta precisão em dados tabulares</b><span>É referência de mercado para dados estruturados como os nossos (volumes por ano, região, produto e segmento).</span></div></div>'
        '<div class="feature"><span class="ic">🔗</span><div><b>Captura relações não-lineares</b><span>Entende interações complexas entre variáveis (ex.: preço da soja + câmbio + crédito agindo juntos).</span></div></div>'
        '<div class="feature"><span class="ic">🛡️</span><div><b>Regularização anti-overfitting</b><span>Penalidades L1/L2 evitam que o modelo "decore" o passado e o mantêm robusto para prever o futuro.</span></div></div>'
        '<div class="feature"><span class="ic">📊</span><div><b>Importância de variáveis</b><span>Mostra quais fatores mais pesam em cada previsão — a base do detalhamento por região/produto/segmento a seguir.</span></div></div>'
        '<div class="feature"><span class="ic">⚡</span><div><b>Rápido e escalável</b><span>Treina em segundos e roda igual para os 4 mercados e 4 produtos, com processamento paralelo.</span></div></div>'
        '<div class="feature"><span class="ic">🕳️</span><div><b>Lida com dados faltantes</b><span>Trata automaticamente lacunas e outliers, comuns em séries históricas de mercado.</span></div></div>'
        '</div></div>'
        + _metodo_detalhamento()
        + open_prod_script
        + '</div>'
    )
    return tab, content

CSS_METODOLOGIA = r"""
    .method-card { background: var(--card-bg); border: 1px solid var(--border-color); border-top: 3px solid var(--agco-red); border-radius: var(--radius); box-shadow: var(--shadow-soft); padding: 1.5rem 1.7rem; margin-bottom: 1.1rem; }
    .method-card p { margin: 0.55rem 0; } .method-card .lead { font-size: 1rem; }
    .method-pill { display: inline-block; font-size: 0.72rem; font-weight: 700; color: #fff; background: linear-gradient(180deg, var(--agco-red), var(--agco-red-dark)); padding: 0.22rem 0.7rem; border-radius: 999px; text-transform: uppercase; letter-spacing: 0.04em; box-shadow: 0 2px 8px rgba(200,16,46,.30); margin-left: 0.7rem; }
    .model-name { font-weight: 800; color: var(--agco-red-dark); }
    .steps { display: grid; grid-template-columns: repeat(4,1fr); gap: 0.9rem; margin-top: 0.4rem; }
    .step { background: var(--surface-2, #F7F8FA); border: 1px solid var(--border-color); border-radius: 12px; padding: 1rem 1.05rem; }
    .step .n { width: 30px; height: 30px; border-radius: 50%; background: linear-gradient(180deg, var(--agco-red), var(--agco-red-dark)); color: #fff; font-family: 'Poppins'; font-weight: 700; font-size: 0.9rem; display: flex; align-items: center; justify-content: center; box-shadow: 0 3px 8px rgba(200,16,46,.28); margin-bottom: 0.6rem; }
    .step h5 { margin: 0 0 0.3rem; font-size: 0.9rem; } .step p { margin: 0; font-size: 0.82rem; color: var(--text-muted); line-height: 1.5; }
    .feature-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0.7rem; margin-top: 0.4rem; }
    .feature { display: flex; gap: 0.7rem; align-items: flex-start; background: #fff; border: 1px solid var(--border-color); border-left: 3px solid var(--agco-red); border-radius: 10px; padding: 0.75rem 0.9rem; }
    .feature .ic { font-size: 1.15rem; } .feature b { color: var(--text-heading); font-size: 0.9rem; } .feature span { display: block; font-size: 0.82rem; color: var(--text-muted); margin-top: 0.15rem; }
    .callout { background: #FEF6F7; border-left: 3px solid var(--agco-red); border-radius: 8px; padding: 0.85rem 1rem; font-size: 0.86rem; color: #521b25; margin-top: 0.6rem; } .callout strong { color: var(--agco-red-dark); }
    .next { border: 1px dashed var(--border-strong, #D9DDE3); border-radius: 12px; padding: 1rem 1.2rem; background: #fff; color: var(--text-muted); font-size: 0.88rem; }
    .detal-legend { display: flex; align-items: center; flex-wrap: wrap; gap: 0.5rem; margin: 0.2rem 0 1rem; font-size: 0.8rem; }
    .detal-legend .legend-title { font-weight: 700; color: var(--text-muted); text-transform: uppercase; letter-spacing: .05em; font-size: 0.72rem; margin-right: 0.3rem; }
    .detal-prod-nav { display: inline-flex; gap: 0.4rem; padding: 0.3rem; background: var(--surface-2, #F7F8FA); border: 1px solid var(--border-color); border-radius: 10px; margin-bottom: 1rem; flex-wrap: wrap; }
    .detal-prod-btn { padding: 0.5rem 1.1rem; border-radius: 7px; border: none; background: transparent; color: var(--text-muted); font-family: 'Poppins', sans-serif; font-weight: 600; font-size: 0.85rem; cursor: pointer; transition: all 0.2s; }
    .detal-prod-btn:hover { color: var(--agco-black); background: #fff; }
    .detal-prod-btn.active { background: linear-gradient(180deg, var(--agco-red), var(--agco-red-dark)); color: #fff; box-shadow: 0 3px 10px rgba(200,16,46,.30); }
    .market-block { margin-bottom: 1.4rem; }
    .market-head { display: flex; align-items: center; gap: 0.6rem; margin: 0.2rem 0 0.5rem; }
    .market-head h4 { margin: 0; font-size: 1.02rem; } .market-flag { font-size: 1.25rem; }
    .market-tag { font-size: 0.66rem; font-weight: 700; text-transform: uppercase; letter-spacing: .05em; color: var(--agco-red-dark); background: var(--agco-red-soft, #FCE1E4); padding: 0.16rem 0.6rem; border-radius: 999px; }
    .detal-card { margin-bottom: 0; } .detal-table td { vertical-align: top; }
    .na-state { padding: 1.1rem 1.3rem; color: var(--text-muted); font-size: 0.9rem; }
    .seg-cell { font-family: 'JetBrains Mono', monospace; font-weight: 600; color: var(--text-heading); font-size: 0.82rem; white-space: nowrap; }
    .var-cell { line-height: 2.1; } .count-cell { text-align: center; }
    .count-badge { display: inline-block; min-width: 26px; font-family: 'JetBrains Mono', monospace; font-weight: 700; font-size: 0.78rem; color: var(--text-heading); background: var(--surface-2, #F7F8FA); border: 1px solid var(--border-color); border-radius: 999px; padding: 2px 8px; }
    .vpill { display: inline-block; font-size: 0.74rem; font-weight: 600; padding: 0.18rem 0.62rem; border-radius: 999px; margin: 0 0.32rem 0.3rem 0; white-space: nowrap; border: 1px solid transparent; }
    .vp-temporal { color: #475569; background: #EEF0F2; border-color: #E2E6EA; }
    .vp-agro { color: var(--agco-red-dark); background: var(--agco-red-soft, #FCE1E4); border-color: #F6C9CF; }
    .vp-macro { color: #E9EDF2; background: linear-gradient(180deg, #2b2b2b, #1a1a1a); border-color: #111; }
    @media (max-width: 992px) { .steps { grid-template-columns: 1fr 1fr; } .feature-grid { grid-template-columns: 1fr; } }
"""

def gerar_dashboard():
    bases = carregar_bases()
    if not bases: return
    now = datetime.now()
    current_year = now.year
    versoes = list(bases.keys())
    versao_rec = versoes[-1]
    versao_ant = versoes[0] if len(versoes) > 1 else None
    forecast_version = versao_rec
    cols = validar_e_mapear_colunas(bases[versao_rec])
    col_m, col_p, col_s, col_y, col_v = cols
    if not all([col_m, col_p, col_s, col_y, col_v]):
        print("\n❌ ERRO: Não foi possível mapear todas as colunas obrigatórias no Excel.")
        return
    dados = []
    for versao, df in bases.items():
        temp = df.copy()
        temp['Versao'] = versao
        dados.append(temp)
    df_full = pd.concat(dados, ignore_index=True)
    df_full = df_full.sort_values(by=[col_m, col_p, col_s, col_y])
    mercados = sorted(list(df_full[col_m].dropna().unique()))
    if not mercados:
        print("\n❌ ERRO CRÍTICO: Nenhum dado de país/mercado foi encontrado.")
        sys.exit(1)
    anos_formatados = {str(int(y)) if isinstance(y, (int, float)) and pd.notna(y) else str(y) for y in df_full[col_y].unique()}
    anos_unicos = sorted(list(anos_formatados))
    mes_keys = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
    mes_key = mes_keys[now.month - 1]
    mes_nome_ia = i18n(mes_key, 'pt')
    mes_nome_ano = f"{mes_nome_ia} de {now.year}"
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
        texto_var = "-"
        cor_kpi = "info"

    # ==========================================================================
    # LAYOUT REFINADO AGCO (topo vermelho, gradientes, cantos suaves, hover suave)
    # ==========================================================================
    injected_css = """
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Poppins:wght@500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');

    .kpi-container { display: none !important; }

    :root {
        --agco-red: #C8102E; --agco-red-hover: #A30B22; --agco-red-dark: #8B0A1D;
        --agco-red-soft: #FCE1E4;
        --agco-black: #111111; --agco-dark: #1F1F1F; --agco-gray-dark: #2B2B2B;
        --agco-gray: #F1F3F5; --agco-silver: #D8D8D8; --white: #FFFFFF;
        --primary-bg: #EEF0F3; --card-bg: #FFFFFF; --surface-2: #F7F8FA;
        --text-heading: #0F172A; --text-body: #384152; --text-muted: #6B7280;
        --border-color: #E7EAEE; --border-strong: #D9DDE3;
        --shadow-soft: 0 1px 2px rgba(16,24,40,.04), 0 6px 16px -6px rgba(16,24,40,.10);
        --shadow-hover: 0 14px 30px -10px rgba(200,16,46,.22), 0 6px 14px -8px rgba(16,24,40,.12);
        --radius: 14px; --radius-sm: 9px;
        --color-hp: #0B7A3B; --bg-hp: #D6F5E3;
        --color-mp: #22A05B; --bg-mp: #E9FBF0;
        --color-n: #6B7280;  --bg-n: #EEF0F2;
        --color-mn: #E5484D; --bg-mn: #FDECEC;
        --color-hn: #C8102E; --bg-hn: #FCE1E4;
    }

    * { box-sizing: border-box; }
    body {
        background:
          radial-gradient(1200px 460px at 100% -12%, rgba(200,16,46,.06), transparent 62%),
          radial-gradient(900px 360px at -10% 0%, rgba(17,17,17,.04), transparent 60%),
          var(--primary-bg);
        font-family: 'Inter', sans-serif; color: var(--text-body);
        -webkit-font-smoothing: antialiased; line-height: 1.55;
        border-top: 4px solid var(--agco-red);
    }
    h1, h2, h3, h4, h5 { font-family: 'Poppins', sans-serif; letter-spacing: -.01em; color: var(--text-heading); }
    ::selection { background: rgba(200,16,46,.16); }

    button { outline: none !important; -webkit-tap-highlight-color: transparent !important; -webkit-appearance: none !important; appearance: none !important; }
    button:focus, button:active, button:focus-visible, button:-webkit-any-link { outline: none !important; box-shadow: none !important; -webkit-tap-highlight-color: transparent !important; }
    *:focus { outline-color: var(--agco-red) !important; }

    .num { font-family: 'JetBrains Mono', monospace; font-size: 0.9em !important; font-variant-numeric: tabular-nums; }
    .positive { color: var(--color-hp) !important; }
    .negative { color: var(--color-hn) !important; }
    .info     { color: var(--text-muted) !important; }

    .section-title { color: var(--text-heading); font-size: 1.18rem; font-weight: 700; margin-bottom: 0.6rem; padding-left: 0.8rem; display: flex; align-items: center; position: relative; }
    .section-title::before { content: ''; position: absolute; left: 0; top: 50%; transform: translateY(-50%); width: 4px; height: 1.15em; border-radius: 3px; background: linear-gradient(180deg, var(--agco-red), var(--agco-red-dark)); }
    .month-badge { font-family: 'Inter', sans-serif; font-size: 0.72rem; font-weight: 700; color: #FFFFFF; background: linear-gradient(180deg, var(--agco-red), var(--agco-red-dark)); padding: 0.22rem 0.65rem; border-radius: 999px; margin-left: 0.65rem; text-transform: uppercase; letter-spacing: 0.04em; box-shadow: 0 2px 8px rgba(200,16,46,.30); }
    .agco-divider { border: 0; height: 2px; background: linear-gradient(90deg, var(--agco-red) 0%, var(--agco-red) 22%, transparent 100%); margin-bottom: 0.8rem; }
    .insights-section { margin-top: 2rem; }
    .insights-container { display: grid; grid-template-columns: 3fr 7fr; gap: 1.2rem; align-items: start; }
    .insight-summary-card { background: var(--card-bg); border: 1px solid var(--border-color); border-radius: var(--radius); padding: 1.15rem 1.3rem; box-shadow: var(--shadow-soft); }
    .insight-summary-card h4, .insight-news-container h4 { font-size: 0.78rem; color: var(--text-muted); margin-bottom: 0.75rem; text-transform: uppercase; letter-spacing: 0.06em; font-weight: 700; }
    .insight-summary-card .summary-content { font-size: 0.85rem; line-height: 1.6; }
    .summary-item { display: flex; align-items: center; gap: 0.75rem; padding: 0.58rem 0.65rem; border-radius: var(--radius-sm); transition: background .2s ease; }
    .summary-item:hover { background: var(--surface-2); }
    .summary-item-icon { font-size: 1.15rem; line-height: 1; }
    .summary-item-text { flex-grow: 1; }
    .summary-item:not(:last-child) { margin-bottom: 0.15rem; }
    .summary-item.positive { border-left: 3px solid var(--color-hp); }
    .summary-item.negative { border-left: 3px solid var(--color-hn); }
    .summary-item.neutral  { border-left: 3px solid var(--color-n); }
    .summary-sentiment-badge { font-size: 0.66rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.04em; padding: 0.16rem 0.55rem; border-radius: 999px; white-space: nowrap; }
    .summary-sentiment-badge.positive { color: var(--color-hp); background: var(--bg-hp); }
    .summary-sentiment-badge.negative { color: var(--color-hn); background: var(--bg-hn); }
    .summary-sentiment-badge.neutral  { color: var(--color-n);  background: var(--bg-n); }
    .news-grid { display: grid; grid-template-columns: 1fr; gap: 0.7rem; max-height: 410px; overflow-y: auto; padding-right: 8px; }
    .news-grid::-webkit-scrollbar { width: 6px; }
    .news-grid::-webkit-scrollbar-track { background: #eef0f2; border-radius: 3px; }
    .news-grid::-webkit-scrollbar-thumb { background: #cfd3d9; border-radius: 3px; }
    .news-grid::-webkit-scrollbar-thumb:hover { background: var(--agco-red); }
    .news-card { background: var(--card-bg); border: 1px solid var(--border-color); border-radius: 12px; padding: 0.9rem 1.05rem; box-shadow: var(--shadow-soft); transition: transform .18s ease, box-shadow .18s ease, border-color .18s ease; }
    .news-card:hover { transform: translateY(-2px); box-shadow: var(--shadow-hover); border-color: #dfe3e8; }
    .news-card.positive-card { border-left: 4px solid var(--color-hp); }
    .news-card.negative-card { border-left: 4px solid var(--color-hn); }
    .news-card.neutral-card  { border-left: 4px solid var(--color-n); }
    .news-card-header { display: flex; align-items: center; gap: 0.5rem; margin-bottom: 0.3rem; }
    .news-sentiment-label { font-size: 0.68rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; }
    .news-card-header .news-title { flex-grow: 1; }
    .news-sentiment-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
    .news-title { margin: 0; font-size: 0.9rem; font-weight: 600; color: var(--text-heading); }
    .news-summary { margin: 0; font-size: 0.8rem; color: var(--text-muted); line-height: 1.5; }
    .news-source-country { font-size: 0.66rem; font-weight: 700; color: #fff; background: var(--agco-gray-dark); padding: 2px 7px; border-radius: 999px; margin-left: 8px; vertical-align: middle; }

    .language-selector-container { position: absolute; top: 15px; right: 20px; z-index: 1000; }
    .lang-toggle { display: inline-flex; padding: 0.25rem; background: var(--card-bg); border: 1px solid var(--border-color); border-radius: 999px; box-shadow: var(--shadow-soft); gap: 0.15rem; }
    .lang-btn { border: none; background: transparent; color: var(--text-muted); font-family: 'Poppins', sans-serif; font-weight: 700; font-size: 0.78rem; letter-spacing: .03em; padding: 0.4rem 0.85rem; border-radius: 999px; cursor: pointer; transition: all 0.2s ease; }
    .lang-btn:hover { color: var(--agco-black); background: var(--surface-2); }
    .lang-btn.active { background: linear-gradient(180deg, var(--agco-red), var(--agco-red-dark)); color: #FFFFFF; box-shadow: 0 3px 10px rgba(200,16,46,.32); }

    .country-tabs-nav { display: inline-flex; gap: 0.3rem; padding: 0.35rem; background: linear-gradient(180deg, #1f1f1f, #0b0b0b); border-radius: 12px; margin-bottom: 0.9rem; box-shadow: 0 6px 16px -8px rgba(0,0,0,.45); flex-wrap: wrap; }
    .country-tab-btn { padding: 0.55rem 1.4rem; border-radius: 8px; border: none; background: transparent; color: #C4C4C4; font-family: 'Poppins', sans-serif; font-weight: 600; font-size: 0.88rem; cursor: pointer; transition: all 0.25s ease; text-transform: uppercase; letter-spacing: 0.05em; }
    .country-tab-btn:hover { color: #FFFFFF; background: rgba(255,255,255,.07); }
    .country-tab-btn.active, .country-tab-btn.active:focus, .country-tab-btn.active:active { background: linear-gradient(180deg, var(--agco-red), var(--agco-red-dark)) !important; color: #FFFFFF !important; box-shadow: 0 4px 12px rgba(200,16,46,0.45) !important; }

    .prod-tabs-container { display: flex; align-items: center; border-bottom: 2px solid var(--border-color); margin-bottom: 0.75rem; }
    .country-title { font-size: 0.78rem; color: var(--text-muted); margin: 0 1.5rem 0 0; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; }
    .prod-tabs-nav { display: flex; gap: 1.7rem; }
    .prod-tab-btn { padding: 0.65rem 0; border: none; background: transparent; color: var(--text-muted); font-family: 'Poppins', sans-serif; font-weight: 600; font-size: 0.92rem; cursor: pointer; border-bottom: 3px solid transparent; margin-bottom: -2px; transition: color 0.2s ease, border-color 0.2s ease; text-transform: uppercase; }
    .prod-tab-btn:hover { color: var(--agco-black); }
    .prod-tab-btn.active, .prod-tab-btn.active:hover, .prod-tab-btn.active:focus, .prod-tab-btn.active:active { color: var(--agco-black) !important; border-bottom-color: var(--agco-red) !important; background: transparent !important; outline: none !important; box-shadow: none !important; }

    .header-action-container { display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 0.55rem; }
    .elegant-btn { display: inline-flex; align-items: center; gap: 0.45rem; background-color: var(--card-bg); color: var(--text-heading); padding: 0.5rem 0.95rem; border-radius: 8px; border: 1px solid var(--border-color); font-family: 'Inter', sans-serif; font-weight: 600; font-size: 0.82rem; cursor: pointer; transition: all 0.2s; box-shadow: var(--shadow-soft); }
    .elegant-btn:hover { background-color: var(--agco-red); color: #FFFFFF; transform: translateY(-1px); box-shadow: var(--shadow-hover); border-color: var(--agco-red); }
    .elegant-btn svg { color: var(--text-body); transition: color 0.2s; }
    .elegant-btn:hover svg { color: #FFFFFF; }

    .elegant-card { background: var(--card-bg); border-radius: var(--radius); box-shadow: var(--shadow-soft); border: 1px solid var(--border-color); border-top: 3px solid var(--agco-red); overflow: hidden; margin-bottom: 0.8rem; }
    table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 0.84rem; }
    th, td { padding: 0.74rem 1.1rem; text-align: left; vertical-align: middle; }
    thead th { background: linear-gradient(180deg, #F7F8FA, #EDF0F3); color: var(--text-heading); font-family: 'Inter', sans-serif; text-transform: uppercase; letter-spacing: 0.05em; font-size: 0.7rem; position: sticky; top: 0; z-index: 2; border-bottom: 1px solid var(--border-strong); }
    tbody tr { transition: background .15s ease; }
    tbody tr:nth-child(even) { background: #FBFCFD; }
    tbody tr:hover { background: var(--agco-red-soft); }
    tbody td { border-bottom: 1px solid #F0F2F4; }
    tfoot tr { position: sticky; bottom: 0; }

    .country-tab-content { padding-top: 0.25rem; animation: fadeIn .3s ease; }
    .prod-tab-content    { padding-top: 0.35rem; animation: fadeIn .25s ease; }
    .prod-tab-content > *:first-child { margin-top: 0.2rem; }
    @keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
    hr { margin: 0.7rem 0; }

    @media (max-width: 992px) {
        .insights-container { grid-template-columns: 1fr; }
        .country-tab-btn { padding: 0.5rem 1rem; font-size: 0.8rem; }
        th, td { padding: 0.6rem 0.7rem; }
    }
    @media print {
        body { background: #fff; }
        .country-tabs-nav, .language-selector-container, .elegant-btn { display: none !important; }
        .elegant-card, .news-card, .insight-summary-card, .method-card { box-shadow: none; }
    }
    """
    # >>> Anexa o CSS da aba Metodologia
    injected_css += CSS_METODOLOGIA

    for br_tag in ['BRASIL', 'BRA']:
        if br_tag in mercados:
            mercados.insert(0, mercados.pop(mercados.index(br_tag)))
            break

    html_country_tabs, html_country_contents = "", ""
    for idx, mercado in enumerate(mercados):
        df_mercado = df_full[df_full[col_m] == mercado]
        print(f"\n--- Processando País: {mercado} ---")
        sys.stdout.flush()
        tab, content = gerar_conteudo_pais(mercado, df_mercado, idx == 0, anos_unicos, versoes, current_year, mes_key, cols, insights_agribusiness, mes_nome_ano)
        html_country_tabs += tab
        html_country_contents += content

    # >>> ABA METODOLOGIA (a Matriz de Fatores foi removida)
    metodo_tab, metodo_content = gerar_conteudo_metodologia()
    html_country_tabs += metodo_tab
    html_country_contents += metodo_content

    all_translations_js = json.dumps(TRANSLATIONS, ensure_ascii=False)

    context = {
        'injected_css': injected_css,
        'all_charts_json': json.dumps({}),
        'forecast_version': str(forecast_version),
        'update_date': now.strftime("%d/%m/%Y"),
        'total_atual_str': str(total_atual_str),
        'versao_rec': str(versao_rec),
        'cor_kpi': str(cor_kpi),
        'texto_var': str(texto_var),
        'html_country_tabs': html_country_tabs,
        'html_country_contents': html_country_contents,
    }

    # --- SELETOR DE IDIOMA: botão PT | EN no topo (ES disponível via função) ---
    language_selector_html = f"""
    <div class="language-selector-container">
        <div class="lang-toggle" role="group" aria-label="Idioma">
            <button type="button" class="lang-btn" id="lang-btn-pt" onclick="translatePage('pt')">PT</button>
            <button type="button" class="lang-btn" id="lang-btn-en" onclick="translatePage('en')">EN</button>
        </div>
    </div>
    <script>
        const ALL_TRANSLATIONS = {all_translations_js};
        let currentLang = localStorage.getItem('dashboardLang') || 'pt';
        function updateLanguageSelector() {{
            ['pt','en','es'].forEach(function(l){{
                var b = document.getElementById('lang-btn-' + l);
                if (b) b.classList.toggle('active', l === currentLang);
            }});
        }}
        function translatePage(lang) {{
            currentLang = lang;
            localStorage.setItem('dashboardLang', lang);
            document.querySelectorAll('[data-i18n]').forEach(function(element) {{
                const key = element.getAttribute('data-i18n');
                if (ALL_TRANSLATIONS[lang] && ALL_TRANSLATIONS[lang][key]) {{ element.textContent = ALL_TRANSLATIONS[lang][key]; }}
                else {{ element.textContent = ALL_TRANSLATIONS['pt'][key] || key; }}
            }});
            document.querySelectorAll('[data-lang-pt]').forEach(function(element) {{
                const tr = element.getAttribute('data-lang-' + lang);
                const fb = element.getAttribute('data-lang-pt');
                element.textContent = (tr !== null && tr !== undefined) ? tr : (fb || '');
            }});
            document.dispatchEvent(new CustomEvent('languageChanged', {{ detail: {{ lang: lang }} }}));
            updateLanguageSelector();
        }}
        document.addEventListener('DOMContentLoaded', function() {{ translatePage(currentLang); }});
    </script>
    """
    context['language_selector'] = language_selector_html

    renderizar_template_final(context)
    print(f"\nSucesso! Dashboard gerado em: {CAMINHO_HTML}")

if __name__ == "__main__":
    gerar_dashboard()
