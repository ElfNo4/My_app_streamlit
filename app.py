import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch
import io
import os

# Configuração da página para estilo moderno
st.set_page_config(page_title="Análise de Investimentos", layout="wide")
st.markdown("""
    <style>
    .main {background-color: #f0f4f8;}
    .stButton>button {background-color: #4CAF50; color: white; border: none; padding: 10px 24px; text-align: center; text-decoration: none; display: inline-block; font-size: 16px; margin: 4px 2px; cursor: pointer; border-radius: 4px;}
    .stButton>button:hover {background-color: #45a049;}
    .stMarkdown {font-family: 'Arial', sans-serif;}
    .watermark {position: fixed; bottom: 10px; right: 10px; color: #888; font-size: 12px;}
    </style>
    """, unsafe_allow_html=True)

# Watermark
st.markdown('<div class="watermark">by Pamella Vilela</div>', unsafe_allow_html=True)

st.title("Análise Estatística de Dados Financeiros")

# Botão para download do blueprint XLSX
def generate_blueprint():
    columns = ['mês', 'aporte', 'taxa de juros', 'saldo inicial', 'juros do mês', 'saldo final']
    sample_data = {
        'mês': ['Janeiro', 'Fevereiro', 'Março'],
        'aporte': [1000.0, 1500.0, 2000.0],
        'taxa de juros': [0.005, 0.006, 0.0055],
        'saldo inicial': [0.0, 1005.0, 2520.5],
        'juros do mês': [5.0, 15.5, 20.0],
        'saldo final': [1005.0, 2520.5, 4540.5]
    }
    df_sample = pd.DataFrame(sample_data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_sample.to_excel(writer, index=False)
    output.seek(0)
    return output

st.download_button(
    label="Baixar Modelo XLSX",
    data=generate_blueprint(),
    file_name="modelo.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Upload do arquivo XLSX
uploaded_file = st.file_uploader("Carregar arquivo XLSX", type="xlsx")

if uploaded_file is not None:
    try:
        # Ler o arquivo (primeira sheet, uma por vez)
        df = pd.read_excel(uploaded_file, engine='openpyxl', sheet_name=0)
        
        # Verificar dados faltando
        if df.isnull().any().any():
            raise ValueError("Erro: Dados faltando ou inválidos no arquivo.")
        
        # Exibir preview dos dados
        st.subheader("Pré-visualização dos Dados")
        st.dataframe(df.head())
        
        # Seleção de colunas pelo usuário
        selected_columns = st.multiselect("Selecionar colunas para análise", df.columns)
        
        if selected_columns:
            # Verificar se colunas selecionadas são numéricas
            df_selected = df[selected_columns].apply(pd.to_numeric, errors='coerce')
            if df_selected.isnull().any().any():
                raise ValueError("Erro: Colunas selecionadas contêm dados não numéricos ou inválidos.")
            
            # Calcular estatísticas descritivas
            stats = {}
            for col in selected_columns:
                series = df_selected[col]
                stats[col] = {
                    'Média': series.mean(),
                    'Mediana': series.median(),
                    'Moda': series.mode().tolist() if not series.mode().empty else 'N/A',
                    'Desvio Padrão': series.std()
                }
            stats_df = pd.DataFrame(stats).T
            
            # Exibir estatísticas
            st.subheader("Estatísticas Descritivas")
            st.table(stats_df)
            
            # Calcular evolução do saldo (se colunas relevantes existirem)
            if 'mês' in df.columns and 'saldo final' in df.columns:
                df['mês'] = pd.Categorical(df['mês'])  # Para ordenação categórica se necessário
                df_sorted = df.sort_values('mês')
                
                # Saldo mensal (saldo final ao longo do tempo)
                fig1, ax1 = plt.subplots()
                ax1.plot(df_sorted['mês'], df_sorted['saldo final'], marker='o')
                ax1.set_title('Evolução do Saldo Mensal')
                ax1.set_xlabel('Mês')
                ax1.set_ylabel('Saldo Final')
                plt.xticks(rotation=45)
                st.pyplot(fig1)
                
                # Saldo de investimento (cumulativo de aportes, se 'aporte' existir)
                if 'aporte' in df.columns:
                    df_sorted['saldo investimento'] = df_sorted['aporte'].cumsum()
                    fig2, ax2 = plt.subplots()
                    ax2.plot(df_sorted['mês'], df_sorted['saldo investimento'], marker='o', color='green')
                    ax2.set_title('Evolução do Saldo de Investimento (Cumulativo)')
                    ax2.set_xlabel('Mês')
                    ax2.set_ylabel('Saldo Investimento')
                    plt.xticks(rotation=45)
                    st.pyplot(fig2)
                
                # Mudança mensal (saldo final - saldo inicial, se 'saldo inicial' existir)
                if 'saldo inicial' in df.columns:
                    df_sorted['mudança mensal'] = df_sorted['saldo final'] - df_sorted['saldo inicial']
                    fig3, ax3 = plt.subplots()
                    ax3.bar(df_sorted['mês'], df_sorted['mudança mensal'], color='blue')
                    ax3.set_title('Mudança Mensal no Saldo')
                    ax3.set_xlabel('Mês')
                    ax3.set_ylabel('Mudança')
                    plt.xticks(rotation=45)
                    st.pyplot(fig3)
            
            # Gerar PDF do relatório
            def generate_pdf():
                buffer = io.BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=letter)
                styles = getSampleStyleSheet()
                elements = []
                
                # Título
                elements.append(Paragraph("Relatório de Análise", styles['Title']))
                elements.append(Spacer(1, 0.2 * inch))
                
                # Tabela de estatísticas
                elements.append(Paragraph("Estatísticas Descritivas", styles['Heading2']))
                data = [stats_df.columns.tolist()] + stats_df.values.tolist()
                t = Table(data)
                t.setStyle([('BACKGROUND', (0,0), (-1,0), colors.grey), ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke)])
                elements.append(t)
                elements.append(Spacer(1, 0.2 * inch))
                
                # Gráficos (salvar como imagens e adicionar)
                if 'mês' in df.columns and 'saldo final' in df.columns:
                    img_buffer1 = io.BytesIO()
                    fig1.savefig(img_buffer1, format='png')
                    img_buffer1.seek(0)
                    elements.append(Paragraph("Evolução do Saldo Mensal", styles['Heading2']))
                    elements.append(Image(img_buffer1, width=5*inch, height=3*inch))
                    elements.append(Spacer(1, 0.2 * inch))
                    
                    if 'aporte' in df.columns:
                        img_buffer2 = io.BytesIO()
                        fig2.savefig(img_buffer2, format='png')
                        img_buffer2.seek(0)
                        elements.append(Paragraph("Evolução do Saldo de Investimento", styles['Heading2']))
                        elements.append(Image(img_buffer2, width=5*inch, height=3*inch))
                        elements.append(Spacer(1, 0.2 * inch))
                    
                    if 'saldo inicial' in df.columns:
                        img_buffer3 = io.BytesIO()
                        fig3.savefig(img_buffer3, format='png')
                        img_buffer3.seek(0)
                        elements.append(Paragraph("Mudança Mensal no Saldo", styles['Heading2']))
                        elements.append(Image(img_buffer3, width=5*inch, height=3*inch))
                
                # Watermark no PDF
                elements.append(Spacer(1, 1 * inch))
                elements.append(Paragraph("by Pamella Vilela", styles['Normal']))
                
                doc.build(elements)
                buffer.seek(0)
                return buffer
            
            # Botão para download do PDF
            st.download_button(
                label="Baixar Relatório em PDF",
                data=generate_pdf(),
                file_name="relatorio_analise.pdf",
                mime="application/pdf"
            )
    
    except ValueError as ve:
        st.error(str(ve))
    except Exception as e:
        st.error(f"Erro inesperado: {str(e)}. Por favor, verifique o arquivo e tente novamente.")