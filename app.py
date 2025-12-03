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


# ==================== 1. Modelo XLSX ====================
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
    label="📥 Baixar Modelo XLSX",
    data=criar_modelo(),
    file_name="modelo_investimentos.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


# ==================== 2. Upload ====================
uploaded_file = st.file_uploader("Carregue seu arquivo XLSX preenchido", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl', sheet_name=0)

        # ---------------------------------------------------------
        # 🚨 CORREÇÃO DEFINITIVA DE COLUNAS 2D DO EXCEL
        # ---------------------------------------------------------

        # Normaliza cabeçalhos
        df.columns = df.columns.map(lambda x: str(x).strip())

        # Remove colunas completamente vazias
        df = df.dropna(axis=1, how='all')

        # Converte cabeçalho MultiIndex do Excel
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = ['_'.join([str(x) for x in col]).strip() for col in df.columns]

        # Converte colunas que vieram como listas, tuplas, arrays → 1D
        for col in df.columns:
            first = df[col].iloc[0]

            # Lista ou tupla → pega o primeiro elemento
            if isinstance(first, (list, tuple)):
                df[col] = df[col].apply(lambda x: x[0] if isinstance(x, (list, tuple)) else x)

            # Array numpy → item() quando possível
            elif hasattr(first, "__len__") and not isinstance(first, str):
                try:
                    df[col] = df[col].apply(lambda x: x.item() if hasattr(x, "item") else x)
                except:
                    pass

        # Coluna "mês" sempre string
        if "mês" in df.columns:
            df["mês"] = df["mês"].astype(str).str.strip()


        # ---------------------------------------------------------
        # Fim da correção 2D
        # ---------------------------------------------------------

        if df.empty:
            st.error("O arquivo está vazio.")
            st.stop()

        if df.isnull().any().any():
            st.error("Existem células vazias no arquivo.")
            st.stop()

        st.success("Arquivo carregado com sucesso!")

        st.subheader("Pré-visualização")
        st.dataframe(df, use_container_width=True)

        # ==================== 3. Estatísticas ====================
        colunas_numericas = df.select_dtypes(include='number').columns.tolist()

        colunas_selecionadas = st.multiselect(
            "Selecione as colunas numéricas para análise",
            colunas_numericas,
            default=colunas_numericas
        )

        stats = pd.DataFrame({
            "Média": df[colunas_selecionadas].mean(),
            "Mediana": df[colunas_selecionadas].median(),
            "Moda": df[colunas_selecionadas].apply(lambda x: x.mode().tolist() if not x.mode().empty else "N/A"),
            "Desvio Padrão": df[colunas_selecionadas].std()
        }).round(4)

        st.subheader("Estatísticas Descritivas")
        st.table(stats)

        # ==================== 4. Gráficos ====================
        figuras = []

        if "mês" in df.columns and "saldo final" in df.columns:
            df_sorted = df.sort_values("mês").copy()
            df_sorted["mês"] = df_sorted["mês"].astype(str)

            # ---- Gráfico 1 ----
            fig, ax = plt.subplots(figsize=(10, 5))
            ax.plot(df_sorted["mês"], df_sorted["saldo final"], marker="o", linewidth=3, color="#2575fc")
            ax.set_title("Evolução do Saldo Final")
            ax.grid(True, alpha=0.3)
            plt.xticks(rotation=45)
            plt.tight_layout()
            st.pyplot(fig)
            figuras.append(fig)

            # ---- Gráfico 2 ----
            if "aporte" in df.columns:
                x = range(len(df_sorted))
                cumulativo = df_sorted["aporte"].cumsum()

                fig2, ax2 = plt.subplots(figsize=(10, 5))
                ax2.fill_between(x, cumulativo, alpha=0.7, color="#6a11cb")
                ax2.plot(x, cumulativo, marker="o", linewidth=3, color="#2575fc")
                ax2.set_title("Aportes Cumulativos")
                ax2.grid(True, alpha=0.3)
                plt.xticks(x, df_sorted["mês"], rotation=45)
                plt.tight_layout()
                st.pyplot(fig2)
                figuras.append(fig2)

        # ==================== 5. PDF/HTML ====================
        def criar_pdf():
            html = f"""
            <html>
            <head>
                <meta charset='utf-8'>
                <style>
                    body {{ font-family: Arial; margin: 40px; background: #f8f9fc; }}
                    h1 {{ color: #2575fc; text-align: center; }}
                    table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
                    th, td {{ border: 1px solid #ccc; padding: 10px; text-align: center; }}
                    th {{ background: #2575fc; color: white; }}
                </style>
            </head>
            <body>
                <h1>Relatório de Análise de Investimentos</h1>
                <p><strong>Data:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>

                <h2>Estatísticas</h2>
                {stats.to_html()}

                <h2>Gráficos</h2>
            """

            for fig in figuras:
                buf = io.BytesIO()
                fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
                buf.seek(0)
                img_base64 = base64.b64encode(buf.read()).decode()
                html += f'<img src="data:image/png;base64,{img_base64}" style="width:100%; margin-bottom:25px;">'

            html += "</body></html>"
            return html

        pdf_html = criar_pdf()

        st.download_button(
            "📄 Baixar Relatório (HTML → PDF via Ctrl+P)",
            data=pdf_html,
            file_name="relatorio_investimentos.html",
            mime="text/html"
        )

    except Exception as e:
        st.error(f"Erro inesperado: {str(e)}")


# ----------------- Rodapé -----------------
st.markdown("---")
st.markdown("<p style='text-align:center;color:#888;'>Desenvolvido com ❤️ por Pamella Vilela</p>", unsafe_allow_html=True)
