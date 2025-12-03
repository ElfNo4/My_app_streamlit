import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use('Agg') # Backend não-interativo, essencial
import matplotlib.pyplot as plt
import io
from datetime import datetime
from fpdf import FPDF # Para PDF natiimport streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
from io import BytesIO
import base64
import os
import tempfile

# Função para criar o blueprint XLSX
def create_blueprint():
    data = {
        'Mês': ['Janeiro', 'Fevereiro'],
        'Aporte': [1000, 1000],
        'Taxa de Juros': [0.01, 0.01],
        'Saldo Inicial': [0, 1000],
        'Juros do Mês': [10, 10],
        'Saldo Final': [1010, 2020]
    }
    df = pd.DataFrame(data)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer

# Função para gerar PDF
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Relatório de Análise', 0, 1, 'C')

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()} - by Pamella Vilela', 0, 0, 'C')

def generate_pdf(report_text, chart_images):
    pdf = PDF()
    pdf.add_page()
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(0, 10, report_text)
    for img_data in chart_images:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
            tmp.write(img_data)
            tmp_path = tmp.name
        pdf.add_page()
        pdf.image(tmp_path, x=10, y=20, w=190)
        os.unlink(tmp_path)
    buffer = BytesIO()
    pdf.output(buffer)
    buffer.seek(0)
    return buffer

# Configuração da página
st.set_page_config(page_title="Analisador de Investimentos", layout="wide")

# CSS para estilização (moderno, escuro)
st.markdown("""
    <style>
    .stApp {
        background-color: #1e1e1e;
        color: #ffffff;
    }
    .stButton > button {
        background-color: #4CAF50;
        color: white;
        border: none;
        padding: 10px 24px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 4px;
    }
    .stButton > button:hover {
        background-color: #45a049;
    }
    h1, h2, h3 {
        color: #4CAF50;
    }
    .stMarkdown {
        color: #ffffff;
    }
    footer {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# Watermark no rodapé
st.markdown('<div style="position: fixed; bottom: 10px; right: 10px; color: #888888;">by Pamella Vilela</div>', unsafe_allow_html=True)

# Título
st.title("Analisador de Investimentos")

# Botão para baixar blueprint
blueprint_buffer = create_blueprint()
st.download_button(
    label="Baixar Modelo XLSX",
    data=blueprint_buffer,
    file_name="modelo_investimentos.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Upload do arquivo
uploaded_file = st.file_uploader("Carregue o arquivo XLSX", type="xlsx")

if uploaded_file:
    try:
        # Ler o XLSX (uma folha apenas)
        df = pd.read_excel(uploaded_file, engine='openpyxl')

        # Verificar dados faltantes ou inválidos
        if df.isnull().values.any():
            raise ValueError("Dados faltantes detectados no arquivo.")
        numeric_cols = df.select_dtypes(include=['number']).columns
        if len(numeric_cols) < len(df.columns) - 1:  # Assumindo pelo menos uma coluna não-numérica (Mês)
            raise ValueError("Dados inválidos: colunas numéricas esperadas não são numéricas.")

        # Seleção de colunas
        columns = st.multiselect("Selecione as colunas para análise", options=df.columns.tolist())

        if columns:
            selected_df = df[columns]

            # Cálculos estatísticos (apenas colunas numéricas)
            stats = {}
            for col in columns:
                if pd.api.types.is_numeric_dtype(selected_df[col]):
                    stats[col] = {
                        'Média': selected_df[col].mean(),
                        'Mediana': selected_df[col].median(),
                        'Moda': selected_df[col].mode()[0] if not selected_df[col].mode().empty else 'N/A',
                        'Desvio Padrão': selected_df[col].std()
                    }

            # Evolução do saldo (assumindo colunas específicas presentes)
            if 'Mês' in df.columns and 'Saldo Final' in df.columns and 'Aporte' in df.columns and 'Juros do Mês' in df.columns:
                df['Variação Mensal'] = df['Saldo Final'].diff().fillna(0)
                total_aporte = df['Aporte'].sum()
                total_juros = df['Juros do Mês'].sum()
                saldo_investimento = total_aporte + total_juros

                # Gráficos otimizados
                fig_balance = px.line(df, x='Mês', y='Saldo Final', title='Evolução do Saldo Final')
                fig_contrib = px.bar(df, x='Mês', y=['Aporte', 'Juros do Mês'], title='Aportes e Juros Mensais')

                # Exibir relatório
                st.header("Relatório de Análise")
                st.subheader("Estatísticas Descritivas")
                st.table(pd.DataFrame(stats).T)

                st.subheader("Evolução do Saldo")
                st.write(f"Saldo do Investimento Total: {saldo_investimento}")
                st.table(df[['Mês', 'Saldo Final', 'Variação Mensal']])

                st.subheader("Gráficos")
                st.plotly_chart(fig_balance)
                st.plotly_chart(fig_contrib)

                # Preparar texto do relatório para PDF
                report_text = "Estatísticas Descritivas:\n"
                for col, vals in stats.items():
                    report_text += f"{col}: Média={vals['Média']}, Mediana={vals['Mediana']}, Moda={vals['Moda']}, Desvio Padrão={vals['Desvio Padrão']}\n"
                report_text += f"\nEvolução do Saldo:\nSaldo do Investimento Total: {saldo_investimento}\n"
                report_text += df[['Mês', 'Saldo Final', 'Variação Mensal']].to_string()

                # Salvar gráficos como imagens para PDF
                chart_images = []
                balance_img = fig_balance.to_image(format="png")
                contrib_img = fig_contrib.to_image(format="png")
                chart_images.extend([balance_img, contrib_img])

                # Botão para baixar PDF
                pdf_buffer = generate_pdf(report_text, chart_images)
                st.download_button(
                    label="Baixar Relatório em PDF",
                    data=pdf_buffer,
                    file_name="relatorio_analise.pdf",
                    mime="application/pdf"
                )
            else:
                st.error("Colunas necessárias para evolução do saldo não encontradas (Mês, Saldo Final, Aporte, Juros do Mês).")
    except ValueError as ve:
        st.error(f"Erro nos dados: {ve}")
    except Exception as e:
        st.error(f"Erro inesperado: {e}")
