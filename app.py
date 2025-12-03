import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import base64
from datetime import datetime

# Forçar backend para Streamlit
plt.switch_backend('Agg')

# Configuração da página
st.set_page_config(page_title="Análise de Investimentos - Pamella Vilela", layout="wide")

# CSS personalizado
st.markdown("""
<style>
    .main {background-color: #f8f9fc; padding: 20px;}
    .stButton>button {
        background: linear-gradient(90deg, #6a11cb, #2575fc);
        color: white; border: none; padding: 12px 30px;
        border-radius: 8px; font-weight: bold;
    }
    .stButton>button:hover {opacity: 0.9;}
    .title {font-size: 42px !important; color: #2575fc; text-align: center;}
    .watermark {
        position: fixed; bottom: 15px; right: 20px; opacity: 0.6;
        font-size: 14px; color: #888; font-style: italic;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="title">Análise de Investimentos</h1>', unsafe_allow_html=True)
st.markdown('<div class="watermark">by Pamella Vilela</div>', unsafe_allow_html=True)

# ==================== 1. Download do Modelo ====================
def criar_modelo():
    dados = {
        'mês': ['Janeiro/2024', 'Fevereiro/2024', 'Março/2024'],
        'aporte': [1000.00, 1200.00, 1500.00],
        'taxa de juros': [0.005, 0.0055, 0.006],
        'saldo inicial': [0.00, 1005.00, 2215.28],
        'juros do mês': [5.00, 11.28, 13.29],
        'saldo final': [1005.00, 2215.28, 3728.57]
    }
    df = pd.DataFrame(dados)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Investimentos')
    output.seek(0)
    return output.getvalue()

st.download_button(
    label="📥 Baixar Modelo XLSX (preenchimento obrigatório)",
    data=criar_modelo(),
    file_name="modelo_investimentos.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ==================== 2. Upload do arquivo ====================
uploaded_file = st.file_uploader("Carregue seu arquivo XLSX preenchido", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl', sheet_name=0)
        # ---- Correção para qualquer coluna 2D vinda do Excel ----

# Se o Excel tiver colunas mescladas ou cabeçalhos estranhos, corrige automaticamente
df.columns = df.columns.map(lambda x: str(x).strip())

# Remove colunas completamente vazias
df = df.dropna(axis=1, how='all')

# Corrige colunas multi-índice (às vezes o Excel cria níveis)
if isinstance(df.columns, pd.MultiIndex):
    df.columns = ['_'.join([str(x) for x in col]).strip() for col in df.columns]

# Converte cada coluna para 1D obrigatoriamente
for col in df.columns:
    if isinstance(df[col].iloc[0], (list, tuple)) or hasattr(df[col].iloc[0], "__len__") and not isinstance(df[col].iloc[0], str):
        # Achata a coluna (pega apenas o primeiro valor)
        df[col] = df[col].apply(lambda x: x[0] if hasattr(x, "__len__") and not isinstance(x, str) else x)

# Converte coluna "mês" para string
if "mês" in df.columns:
    df["mês"] = df["mês"].astype(str).str.strip()

# Tenta converter números corretamente
for col in df.columns:
    df[col] = pd.to_numeric(df[col], errors="ignore")


        if df.empty:
            st.error("O arquivo está vazio.")
            st.stop()

        if df.isnull().any().any():
            st.error("Existem células vazias ou dados inválidos no arquivo. Corrija e tente novamente.")
            st.stop()

        # ---- CORREÇÃO CRÍTICA: garantir que "mês" é coluna 1D e string ----
        if 'mês' in df.columns:
            df['mês'] = df['mês'].astype(str).str.strip()

        st.success("Arquivo carregado com sucesso!")
        st.subheader("Pré-visualização dos dados")
        st.dataframe(df, use_container_width=True)

        # ==================== 3. Seleção de colunas ====================
        colunas_numericas = df.select_dtypes(include='number').columns.tolist()
        if not colunas_numericas:
            st.error("Nenhuma coluna numérica encontrada.")
            st.stop()

        colunas_selecionadas = st.multiselect(
            "Selecione as colunas numéricas para análise",
            options=colunas_numericas,
            default=colunas_numericas
        )

        if not colunas_selecionadas:
            st.warning("Selecione pelo menos uma coluna.")
            st.stop()

        # ==================== 4. Estatísticas descritivas ====================
        stats = pd.DataFrame({
            "Média": df[colunas_selecionadas].mean(),
            "Mediana": df[colunas_selecionadas].median(),
            "Moda": df[colunas_selecionadas].apply(lambda x: x.mode().tolist() if not x.mode().empty else "N/A"),
            "Desvio Padrão": df[colunas_selecionadas].std()
        }).round(4)

        st.subheader("Estatísticas Descritivas")
        st.table(stats)

        # ==================== 5. Evolução do saldo ====================
        figuras = []

        if 'mês' in df.columns and 'saldo final' in df.columns:
            # Garantir consistência dos dados
            df_sorted = df.sort_values('mês').copy()
            df_sorted['mês'] = df_sorted['mês'].astype(str)

            # ---------- Gráfico 1 - Saldo Final ----------
            fig, ax = plt.subplots(figsize=(10, 5))
            ax.plot(df_sorted['mês'], df_sorted['saldo final'],
                    marker='o', linewidth=3, color='#2575fc')
            ax.set_title('Evolução do Saldo Final', fontsize=16, fontweight='bold')
            ax.set_xlabel('Mês')
            ax.set_ylabel('Saldo Final (R$)')
            ax.grid(True, alpha=0.3)
            plt.xticks(rotation=45)
            plt.tight_layout()
            st.pyplot(fig)
            figuras.append(fig)

            # ---------- Gráfico 2 - Aportes Cumulativos ----------
            if 'aporte' in df.columns:
                x = range(len(df_sorted))
                cumulativo = df_sorted['aporte'].cumsum()

                fig2, ax2 = plt.subplots(figsize=(10, 5))
                ax2.fill_between(x, cumulativo, alpha=0.7, color='#6a11cb')
                ax2.plot(x, cumulativo, marker='o', color='#2575fc', linewidth=3)
                ax2.set_title('Evolução do Total Investido (Aportes Cumulativos)', fontsize=16, fontweight='bold')
                ax2.set_ylabel('Total Investido (R$)')
                ax2.grid(True, alpha=0.3)

                plt.xticks(x, df_sorted['mês'], rotation=45)
                plt.tight_layout()
                st.pyplot(fig2)
                figuras.append(fig2)

        # ==================== 6. Geração do PDF (HTML) ====================
        def criar_pdf():
            html = f"""
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 40px; background: #f8f9fc; }}
                    h1 {{ color: #2575fc; text-align: center; }}
                    table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
                    th, td {{ border: 1px solid #ddd; padding: 12px; text-align: center; }}
                    th {{ background: #2575fc; color: white; }}
                    .watermark {{ position: fixed; bottom: 30px; right: 30px; opacity: 0.5; font-size: 18px; }}
                </style>
            </head>
            <body>
                <h1>Relatório de Análise de Investimentos</h1>
                <p><strong>Data do relatório:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
                <h2>Estatísticas Descritivas</h2>
                {stats.to_html()}
                <h2>Gráficos</h2>
            """

            for fig in figuras:
                buf = io.BytesIO()
                fig.savefig(buf, format='png', dpi=200, bbox_inches='tight')
                buf.seek(0)
                img_base64 = base64.b64encode(buf.read()).decode()
                html += f'<img src="data:image/png;base64,{img_base64}" style="width:100%; margin:30px 0;"><br>'

            html += """
                <div class="watermark">by Pamella Vilela</div>
            </body>
            </html>
            """

            return html

        pdf_html = criar_pdf()

        st.subheader("Relatório Completo Gerado")
        st.markdown("### ✅ Tudo pronto! Clique no botão abaixo para baixar o PDF completo:")

        st.download_button(
            label="📄 Baixar Relatório em PDF",
            data=pdf_html,
            file_name=f"relatorio_investimentos_{datetime.now().strftime('%Y%m%d')}.html",
            mime="text/html",
            help="Após baixar, abra o arquivo HTML e use Ctrl+P → Salvar como PDF."
        )

        st.info("💡 Dica: Abra o arquivo no navegador e pressione Ctrl+P → 'Salvar como PDF'.")

    except Exception as e:
        st.error(f"Erro inesperado: {str(e)}")
        st.error("Verifique se o arquivo segue exatamente o modelo baixado acima.")

# Rodapé
st.markdown("---")
st.markdown("<p style='text-align:center; color:#888;'>Desenvolvido com ❤️ por Pamella Vilela</p>", unsafe_allow_html=True)

