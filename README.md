# 📈 AGCO LATAM Forecast Intelligence

## 🎯 Visão de Negócio

**O Problema:** O processo de forecast de mercado para máquinas agrícolas é complexo e dinâmico. A cada ciclo, os números mudam, e entender *o que* mudou e, mais importante, *por que* mudou, é um desafio. Analistas gastam horas cruzando planilhas, comparando versões e tentando conectar as variações do forecast com as condições de mercado (preços de commodities, juros, câmbio, clima), um processo manual, lento e sujeito a inconsistências.

**A Solução:** Este projeto automatiza a criação de um **Dashboard Executivo de Inteligência de Forecast**. Ele não apenas compara diferentes versões do forecast e destaca as variações, mas também integra uma camada de inteligência artificial (usando Google Gemini) para fornecer análises de mercado em tempo real, explicando o "porquê" por trás dos números. O sistema conecta as projeções de volume com os drivers de mercado que realmente impactam cada segmento de produto.

### ✨ Valor para o Negócio

-   **Agilidade na Análise:** Reduz de horas para minutos o tempo necessário para analisar as mudanças entre os ciclos de forecast.
-   **Inteligência Integrada:** Conecta os "números" (forecast) com a "narrativa" (drivers de mercado), fornecendo uma visão completa e contextualizada.
-   **Decisões Baseadas em Dados:** Permite que a liderança e as equipes de produto identifiquem rapidamente riscos e oportunidades, entendendo quais segmentos são mais sensíveis a quais fatores de mercado.
-   **Comunicação Eficiente:** Padroniza a análise do forecast em um formato visual e interativo, facilitando a comunicação e o alinhamento entre as equipes de finanças, produto e vendas.

---

## 🛠️ Descrição Técnica

O sistema é um script Python que orquestra a leitura de dados de planilhas Excel, o enriquecimento com análises de IA e a geração de um dashboard HTML interativo.

### 🔄 Fluxo de Execução

1.  **Carregamento de Versões:** O script localiza o arquivo de forecast mais recente (ex: `Forecast_8+4.xlsx`) na pasta raiz e o compara com a versão imediatamente anterior, que está arquivada na pasta `historico_forecast/`.
2.  **Análise de Mercado com IA:** O script verifica se existe uma análise de mercado em cache para o mês atual (`insights_ia_{mes}.json`).
    -   **Se não houver cache:** Ele se conecta à API do Google Gemini com um prompt detalhado, solicitando uma análise sobre a tendência ("up", "down", "neutral") e o motivo para dezenas de drivers de mercado (Soja, Milho, Juros, Câmbio, etc.) para cada região (Brasil, Argentina, México, OSA).
    -   O resultado da IA é salvo em cache para otimizar custos e velocidade em execuções futuras no mesmo mês.
3.  **Geração do Dashboard:** O script processa os dados de forecast e os combina com as análises da IA.
    -   Ele calcula as variações de volume entre os ciclos e entre os anos.
    -   Utiliza uma **Matriz de Impacto** (lógica interna) para cruzar quais drivers de mercado afetam quais segmentos de produto (ex: Tratores de Baixa Potência no Brasil são altamente impactados por Juros e Agricultura Familiar).
    -   Injeta todas essas informações (tabelas, gráficos, KPIs, análises da IA) em um arquivo `template.html`.
4.  **Saída Final:** O resultado é um único arquivo, `dashboard_forecast.html`, que contém o dashboard completo e interativo, pronto para ser aberto em qualquer navegador.

### 📂 Componentes Principais

-   `gerar_dashboard_forecast.py`: O script Python principal que contém toda a lógica.
-   `template.html`: Template HTML que serve de base para o dashboard, com placeholders para os dados.
-   `*.xlsx`: Arquivos de entrada com os dados do forecast. O nome deve conter a versão (ex: `Forecast_8+4.xlsx`).
-   `historico_forecast/`: Pasta criada automaticamente para armazenar o histórico de arquivos de forecast, permitindo a comparação entre ciclos.
-   `insights_ia_{mes}.json`: Arquivo de cache com as análises geradas pela IA para o mês corrente.
-   `dashboard_forecast.html`: O arquivo final gerado, contendo o dashboard.

---

## 🚀 Como Usar (Ambiente Virtual Recomendado)

1.  **Preparar o Ambiente Virtual:**
    -   Abra um terminal na pasta do projeto.
    -   Crie o ambiente virtual (só precisa fazer isso uma vez):
        ```bash
        python -m venv venv
        ```
    -   Ative o ambiente:
        -   **Windows (PowerShell):** `.\venv\Scripts\Activate.ps1`
        -   **Windows (CMD):** `.\venv\Scripts\activate.bat`
        -   **macOS/Linux:** `source venv/bin/activate`
    -   *Seu terminal deve agora mostrar `(venv)` no início da linha.*

2.  **Instalar Dependências:**
    -   Com o ambiente ativado, instale as bibliotecas necessárias:
        ```bash
        pip install -r requirements.txt
        ```

3.  **Configurar a Chave de API:**
    -   Renomeie o arquivo `.env.example` para `.env`.
    -   Abra o novo arquivo `.env` e substitua `"sua-chave-aqui"` pela sua chave de API real do Google Gemini.
        ```
        GEMINI_API_KEY="sua-chave-real-aqui"
        ```
    -   O arquivo `.gitignore` já está configurado para impedir que este arquivo seja enviado para o repositório, mantendo sua chave segura.

4.  **Preparar o Arquivo de Entrada:**
    -   Coloque o arquivo Excel com o forecast mais recente na mesma pasta do script. O nome do arquivo deve indicar a versão, por exemplo: `Meu_Forecast_LATAM_8+4.xlsx`.

5.  **Executar:**
    -   Certifique-se de que o ambiente virtual ainda está ativado.
        ```bash
        python gerar_dashboard_forecast.py
        ```

6.  **Visualizar o Resultado:**
    -   Abra o arquivo `dashboard_forecast.html` que foi gerado na pasta.

7.  **Desativar o Ambiente (Opcional):**
    -   Quando terminar, você pode desativar o ambiente com o comando:
        ```powershell
        deactivate
        ```

---

## ✨ Funcionalidades

-   **Comparação de Ciclos:** Visualização clara das mudanças de volume entre a versão atual e a anterior do forecast.
-   **Análise Multi-dimensional:** Filtros por País (abas) e Família de Produto (sub-abas).
-   **KPIs Executivos:** Cards no topo com o volume total e a variação percentual consolidada.
-   **Inteligência de Mercado (IA):** Seção com o status e a análise dos principais drivers de mercado, atualizada mensalmente pela IA.
-   **Matriz de Impacto:** Tabela visual que mostra a sensibilidade de cada segmento de produto aos drivers de mercado.
-   **Gráficos Dinâmicos:** Gráficos de barra que comparam os volumes por segmento ou por ano.
-   **Exportação para Excel:** Botão para exportar os dados da tabela visualizada para um arquivo `.xlsx`.
-   **Suporte a Múltiplos Idiomas:** A interface pode ser traduzida (a base está preparada com tags `data-i18n`).