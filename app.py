import streamlit as st
import pandas as pd
from io import BytesIO
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# Função para criar o modelo XLSX (blueprint)
def create_blueprint():
    data = {
        'mês': ['Janeiro', 'Fevereiro', 'Março'],
        'aporte': [1000.0, 1500.0, 2000.0],
        'taxa de juros': [0.005, 0.006, 0.0055],
        'saldo inicial': [0.0, 1005.0, 2513.03],
        'juros do mês': [5.0, 8.03, 14.12],
        'saldo final': [1005.0, 2513.03, 4527.15]
    }
    df = pd.DataFrame(data)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
    buffer.seek(0)
    return buffer.getvalue()

# Função para gerar PDF do relatório
def generate_pdf(df, selected_cols, stats_df):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    y = height - 50

    c.drawString(50, y, "Relatório de Análise Financeira")
    y -= 30

    # Adicionar estatísticas
    c.drawString(50, y, "Estatísticas Descritivas:")
    y -= 20
    for col in selected_cols:
        c.drawString(50, y, f"Coluna: {col}")
        y -= 20
        stats = stats_df.loc[col]
        c.drawString(70, y, f"Média: {stats['média']:.2f}")
        y -= 20
        c.drawString(70, y, f"Mediana: {stats['mediana']:.2f}")
        y -= 20
        moda = stats['moda']
        c.drawString(70, y, f"Moda: {moda:.2f}" if pd.notnull(moda) else "Moda: Nenhuma")
        y -= 20
        c.drawString(70, y, f"Desvio Padrão: {stats['desvio padrão']:.2f}")
        y -= 30

    # Gerar e adicionar gráficos como imagens
    if 'saldo final' in df.columns and 'mês' in df.columns:
        # Evolução do Saldo
        fig1, ax1 = plt.subplots(figsize=(6, 4))
        ax1.plot(df['mês'], df['saldo final'], marker='o')
        ax1.set_title('Evolução do Saldo')
        ax1.set_xlabel('Mês')
        ax1.set_ylabel('Saldo Final')
        plt.xticks(rotation=45)
        fig1.tight_layout()
        img_buffer1 = BytesIO()
        fig1.savefig(img_buffer1, format='png')
        img_buffer1.seek(0)
        img1 = ImageReader(img_buffer1)
        c.drawImage(img1, 50, y - 250, width=400, height=200)
        y -= 270

        # Saldo Mensal
        monthly_balance = df['saldo final'] - df['saldo inicial']
        fig2, ax2 = plt.subplots(figsize=(6, 4))
        ax2.plot(df['mês'], monthly_balance, marker='o')
        ax2.set_title('Saldo Mensal')
        ax2.set_xlabel('Mês')
        ax2.set_ylabel('Saldo Mensal')
        plt.xticks(rotation=45)
        fig2.tight_layout()
        img_buffer2 = BytesIO()
        fig2.savefig(img_buffer2, format='png')
        img_buffer2.seek(0)
        img2 = ImageReader(img_buffer2)
        c.drawImage(img2, 50, y - 250, width=400, height=200)
        y -= 270

        # Saldo de Investimento (cumulativo de aportes)
        investment_balance = df['aporte'].cumsum()
        fig3, ax3 = plt.subplots(figsize=(6, 4))
        ax3.plot(df['mês'], investment_balance, marker='o')
        ax3.set_title('Saldo de Investimento')
        ax3.set_xlabel('Mês')
        ax3.set_ylabel('Investimento Cumulativo')
        plt.xticks(rotation=45)
        fig3.tight_layout()
        img_buffer3 = BytesIO()
        fig3.savefig(img_buffer3, format='png')
        img_buffer3.seek(0)
        img3 = ImageReader(img_buffer3)
        c.drawImage(img3, 50, y - 250, width=400, height=200)

    c.save()
    buffer.seek(0)
    return buffer.getvalue()

# Interface principal
st.title("Análise Financeira de XLSX")

# Botão para baixar modelo
st.download_button(
    label="Baixar modelo XLSX",
    data=create_blueprint(),
    file_name="modelo.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Upload do arquivo
uploaded_file = st.file_uploader("Carregue o arquivo XLSX", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Ler o arquivo (apenas a primeira sheet)
        df = pd.read_excel(uploaded_file, engine='openpyxl', sheet_name=0)

        # Verificar colunas obrigatórias
        expected_cols = ['mês', 'aporte', 'taxa de juros', 'saldo inicial', 'juros do mês', 'saldo final']
        missing_cols = [col for col in expected_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Colunas faltando: {', '.join(missing_cols)}")

        # Verificar se colunas (exceto 'mês') são numéricas
        for col in expected_cols[1:]:
            if not pd.api.types.is_numeric_dtype(df[col]):
                raise ValueError(f"A coluna '{col}' deve ser numérica.")

        # Verificar dados faltando ou inválidos
        if df.isnull().any().any():
            raise ValueError("Existem dados faltando ou inválidos no arquivo.")

        # Exibir dados
        st.subheader("Dados Carregados")
        st.dataframe(df)

        # Seleção de colunas numéricas
        numeric_cols = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col])]
        selected_cols = st.multiselect("Selecione colunas para análise", options=numeric_cols)

        if selected_cols:
            # Calcular estatísticas
            stats = {}
            for col in selected_cols:
                series = df[col]
                mode = series.mode()
                stats[col] = {
                    'média': series.mean(),
                    'mediana': series.median(),
                    'moda': mode[0] if not mode.empty else None,
                    'desvio padrão': series.std()
                }
            stats_df = pd.DataFrame(stats).T

            # Exibir relatório
            st.subheader("Estatísticas Descritivas")
            st.dataframe(stats_df)

            # Evolução do saldo (fixa em colunas específicas, se presentes)
            if all(col in df.columns for col in ['mês', 'saldo final', 'saldo inicial', 'aporte']):
                st.subheader("Evolução do Saldo")
                st.line_chart(df.set_index('mês')['saldo final'])

                monthly_balance = df['saldo final'] - df['saldo inicial']
                st.subheader("Saldo Mensal")
                st.line_chart(pd.DataFrame({'mês': df['mês'], 'saldo mensal': monthly_balance}).set_index('mês'))

                investment_balance = df['aporte'].cumsum()
                st.subheader("Saldo de Investimento")
                st.line_chart(pd.DataFrame({'mês': df['mês'], 'saldo de investimento': investment_balance}).set_index('mês'))

            # Botão para baixar PDF (no final da página)
            pdf_data = generate_pdf(df, selected_cols, stats_df)
            st.download_button(
                label="Baixar Relatório em PDF",
                data=pdf_data,
                file_name="relatorio_analise.pdf",
                mime="application/pdf"
            )

    except ValueError as ve:
        st.error(f"Erro: {ve}")
    except Exception as e:
        st.error(f"Erro inesperado: {str(e)}")
        st.error(f"Erro inesperado: {str(e)}. Por favor, verifique o arquivo e tente novamente.")

