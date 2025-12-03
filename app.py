import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import base64
from datetime import datetime

plt.switch_backend('Agg')

st.set_page_config(page_title="Análise de Investimentos - Pamella Vilela", layout="wide")


# ==================== CSS ====================
st.markdown("""
<style>
    .main {background-color: #f8f9fc; padding: 20px;}
    .stButton>button {
        background: linear-gradient(90deg, #6a11cb, #2575fc);
        color: white; border: none; padding: 12px 30px;
        border-radius: 8px; font-weight: bold;
    }
    .title {font-size: 42px !important; color: #2575fc; text-align: center;}
    .watermark {
        position: fixed; bottom: 15px; right: 20px; opacity: 0.6;
        font-size: 14px; color: #888;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="title">Análise de Investimentos</h1>', unsafe_allow_html=True)
st.markdown('<div class="watermark">by Pamella Vilela</div>', unsafe_allow_html=True)


# ==================== 1. NOVO MODELO XLSX (LIMPO E CORRETO) ====================
def criar_modelo():
    df = pd.DataFrame({
        "mês": ["Janeiro/2024", "Fevereiro/2024", "Março/2024"],
        "aporte": [1000, 1200, 1500],
        "taxa de juros": [0.005, 0.0055, 0.006],
        "saldo inicial": [0.0, 1005.0, 2215.28],
        "juros do mês": [5.0, 11.28, 13.29],
        "saldo final": [1005.0, 2215.28, 3728.57]
    })

    # Geração do XLSX sem risco de multi-index ou merges
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Investimentos")

    output.seek(0)
    return output.getvalue()


st.download_button(
    "📥 Baixar Modelo XLSX (corrigido)",
    data=criar_modelo(),
    file_name="modelo_investimentos.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


# ==================== 2. UPLOAD DO ARQUIVO ====================
uploaded_file = st.file_uploader("Carregue seu arquivo XLSX preenchido", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

        # LIMPEZA GARANTIDA — elimina qualquer estrutura 2D
        df.columns = df.columns.map(str)

        # Força todas as colunas a serem 1D
        for col in df.columns:
            df[col] = df[col].apply(
                lambda x: x[0] if isinstance(x, (list, tuple)) else x
            )

        if "mês" in df.columns:
            df["mês"] = df["mês"].astype(str)

        st.success("Arquivo carregado com sucesso!")
        st.dataframe(df, use_container_width=True)


        # ==================== 3. ESTATÍSTICAS ====================
        num_cols = df.select_dtypes(include="number").columns.tolist()

        stats = pd.DataFrame({
            "Média": df[num_cols].mean(),
            "Mediana": df[num_cols].median(),
            "Moda": df[num_cols].apply(lambda x: x.mode()[0] if not x.mode().empty else "N/A"),
            "Desvio Padrão": df[num_cols].std()
        })

        st.subheader("Estatísticas Descritivas")
        st.table(stats)


        # ==================== 4. GRÁFICOS ====================
        figs = []

        if "mês" in df.columns and "saldo final" in df.columns:
            df_sorted = df.sort_values("mês")

            fig, ax = plt.subplots(figsize=(9, 4))
            ax.plot(df_sorted["mês"], df_sorted["saldo final"], marker="o", linewidth=3)
            ax.grid(True)
            plt.xticks(rotation=45)
            st.pyplot(fig)
            figs.append(fig)

        if "aporte" in df.columns:
            df_sorted = df.sort_values("mês")
            cumul = df_sorted["aporte"].cumsum()
            x = range(len(cumul))

            fig2, ax2 = plt.subplots(figsize=(9, 4))
            ax2.fill_between(x, cumul, alpha=.6)
            ax2.plot(x, cumul, marker="o")
            ax2.grid(True)
            plt.xticks(x, df_sorted["mês"], rotation=45)
            st.pyplot(fig2)
            figs.append(fig2)


        # ==================== 5. GERAR PDF (HTML) ====================
        def criar_pdf():
            html = f"""
            <html><body>
            <h1>Relatório de Análise</h1>
            <p><b>Gerado em:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
            <h2>Estatísticas</h2>
            {stats.to_html()}
            """

            for f in figs:
                buf = io.BytesIO()
                f.savefig(buf, format="png", dpi=180, bbox_inches="tight")
                buf.seek(0)
                img = base64.b64encode(buf.read()).decode()
                html += f'<img src="data:image/png;base64,{img}" style="width:100%; margin:20px 0;">'

            html += "</body></html>"
            return html

        html_relatorio = criar_pdf()

        st.download_button(
            "📄 Baixar Relatório em HTML",
            data=html_relatorio,
            file_name="relatorio_investimentos.html",
            mime="text/html"
        )


    except Exception as e:
        st.error(f"Erro inesperado: {e}")


# ==================== Rodapé ====================
st.markdown("---")
st.markdown("<p style='text-align:center;'>Feito com ❤️ por Pamella Vilela</p>", unsafe_allow_html=True)
