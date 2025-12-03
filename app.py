import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Backend não-interativo, essencial
import matplotlib.pyplot as plt
import io
from datetime import datetime
from fpdf import FPDF  # Para PDF nativo e estável

# Configuração da página
st.set_page_config(page_title="Análise de Investimentos - Pamella Vilela", layout="wide")

# CSS para estilo moderno
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
@st.cache_data
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
        # Leitura robusta: especifica colunas como strings para evitar inferência errada
        df = pd.read_excel(uploaded_file, engine='openpyxl', sheet_name=0, dtype=str)
        
        # Converter colunas numéricas explicitamente após leitura
        colunas_numericas_possiveis = ['aporte', 'taxa de juros', 'saldo inicial', 'juros do mês', 'saldo final']
        for col in colunas_numericas_possiveis:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Verificar se está vazio
        if df.empty:
            raise ValueError("O arquivo está vazio.")

        # Verificar dados inválidos/NaN após conversão
        if df[colunas_numericas_possiveis].isnull().any().any():
            raise ValueError("Existem células vazias ou dados inválidos (não numéricos) nas colunas numéricas. Corrija e tente novamente.")

        st.success("✅ Arquivo carregado com sucesso!")
        st.subheader("Pré-visualização dos dados")
        st.dataframe(df, use_container_width=True)

        # ==================== 3. Seleção de colunas ====================
        colunas_numericas = df.select_dtypes(include='number').columns.tolist()
        if not colunas_numericas:
            raise ValueError("Nenhuma coluna numérica encontrada no arquivo.")

        colunas_selecionadas = st.multiselect(
            "Selecione as colunas numéricas para análise",
            options=colunas_numericas,
            default=colunas_numericas  # Seleciona todas por default
        )

        if not colunas_selecionadas:
            st.warning("⚠️ Selecione pelo menos uma coluna para prosseguir.")
            st.stop()

        # ==================== 4. Estatísticas descritivas (CORRIGIDO) ====================
        # Inicializar dicionários para stats escalares (1D sempre)
        medias = df[colunas_selecionadas].mean()
        medianas = df[colunas_selecionadas].median()
        desvios = df[colunas_selecionadas].std()
        
        # Moda corrigida: extrair apenas o primeiro modo como string/float
        modas = {}
        for col in colunas_selecionadas:
            modo_series = df[col].mode()
            if not modo_series.empty:
                modas[col] = modo_series.iloc[0]  # Primeiro modo apenas
            else:
                modas[col] = "N/A"
        
        # Construir DataFrame linha por linha para evitar ndarray 2D
        stats_data = []
        for col in colunas_selecionadas:
            stats_data.append({
                'Coluna': col,
                'Média': round(medias[col], 4),
                'Mediana': round(medianas[col], 4),
                'Moda': modas[col],
                'Desvio Padrão': round(desvios[col], 4)
            })
        stats = pd.DataFrame(stats_data).set_index('Coluna')

        st.subheader("📊 Estatísticas Descritivas")
        st.table(stats.style.format({'Média': '{:.4f}', 'Mediana': '{:.4f}', 'Desvio Padrão': '{:.4f}'}))

        # ==================== 5. Evolução do saldo ====================
        figuras = []
        tem_mes = 'mês' in df.columns
        tem_saldo_final = 'saldo final' in df.columns

        if tem_mes and tem_saldo_final:
            # Ordenar por mês (tratar como categórico se string)
            if df['mês'].dtype == 'object':
                df['mês_ordinal'] = pd.Categorical(df['mês'], categories=sorted(df['mês'].unique()), ordered=True)
                df_sorted = df.sort_values('mês_ordinal')
            else:
                df_sorted = df.sort_values('mês')
            
            del df_sorted['mês_ordinal'] if 'mês_ordinal' in df_sorted.columns else None

            # Gráfico 1 - Saldo Final
            fig1, ax1 = plt.subplots(figsize=(10, 5))
            ax1.plot(df_sorted['mês'], df_sorted['saldo final'], marker='o', linewidth=3, color='#2575fc')
            ax1.set_title('Evolução do Saldo Final', fontsize=16, fontweight='bold')
            ax1.set_xlabel('Mês')
            ax1.set_ylabel('Saldo Final (R$)')
            ax1.grid(True, alpha=0.3)
            plt.xticks(rotation=45)
            plt.tight_layout()
            figuras.append(fig1)
            st.pyplot(fig1)

            # Gráfico 2 - Aportes cumulativos (se existir)
            if 'aporte' in df.columns:
                fig2, ax2 = plt.subplots(figsize=(10, 5))
                cumulativo = df_sorted['aporte'].cumsum()
                ax2.fill_between(df_sorted['mês'], cumulativo, alpha=0.7, color='#6a11cb')
                ax2.plot(df_sorted['mês'], cumulativo, marker='o', color='#2575fc', linewidth=3)
                ax2.set_title('Evolução do Total Investido (Aportes Cumulativos)', fontsize=16, fontweight='bold')
                ax2.set_ylabel('Total Investido (R$)')
                ax2.grid(True, alpha=0.3)
                plt.xticks(rotation=45)
                plt.tight_layout()
                figuras.append(fig2)
                st.pyplot(fig2)

        # ==================== 6. Geração do PDF (agora com FPDF nativo) ====================
        @st.cache_data
        def gerar_pdf(_stats, _figuras, _df):
            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.add_page()
            pdf.add_font('Arial', '', 'Arial.ttf', uni=True)  # Suporte a acentos
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, 'Relatório de Análise de Investimentos', ln=1, align='C')
            pdf.set_font('Arial', '', 12)
            pdf.cell(0, 10, f'Data do relatório: {datetime.now().strftime("%d/%m/%Y %H:%M")}', ln=1, align='C')
            
            # Estatísticas
            pdf.ln(10)
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Estatísticas Descritivas', ln=1)
            pdf.set_font('Arial', '', 10)
            # Cabeçalho da tabela
            pdf.cell(40, 8, 'Coluna', 1)
            pdf.cell(35, 8, 'Média', 1)
            pdf.cell(35, 8, 'Mediana', 1)
            pdf.cell(35, 8, 'Moda', 1)
            pdf.cell(45, 8, 'Desv. Padrão', 1)
            pdf.ln()
            # Linhas
            for idx, row in _stats.iterrows():
                pdf.cell(40, 8, str(idx), 1)
                pdf.cell(35, 8, f'{row["Média"]:.4f}', 1)
                pdf.cell(35, 8, f'{row["Mediana"]:.4f}', 1)
                pdf.cell(35, 8, str(row["Moda"]), 1)
                pdf.cell(45, 8, f'{row["Desvio Padrão"]:.4f}', 1)
                pdf.ln()
            
            # Gráficos
            if _figuras:
                pdf.ln(10)
                pdf.set_font('Arial', 'B', 14)
                pdf.cell(0, 10, 'Gráficos da Análise', ln=1)
                for i, fig in enumerate(_figuras):
                    img_buffer = io.BytesIO()
                    fig.savefig(img_buffer, format='PNG', dpi=150, bbox_inches='tight')
                    img_buffer.seek(0)
                    pdf.image(img_buffer, x=10, y=pdf.get_y(), w=190)
                    pdf.ln(10)
            
            # Watermark
            pdf.set_font('Arial', 'I', 12)
            pdf.set_xy(150, 270)
            pdf.cell(0, 10, 'by Pamella Vilela', 0, align='R')
            
            output_pdf = io.BytesIO()
            output_pdf.write(pdf.output(dest='S').encode('latin1'))  # Encode para bytes
            output_pdf.seek(0)
            return output_pdf.getvalue()

        # Botão de download do PDF
        pdf_data = gerar_pdf(stats, figuras, df)
        st.subheader("📄 Relatório Completo Gerado")
        st.success("✅ Análise realizada! Baixe o PDF com estatísticas e gráficos.")
        st.download_button(
            label="Baixar Relatório em PDF",
            data=pdf_data,
            file_name=f"relatorio_investimentos_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf"
        )

    except ValueError as ve:
        st.error(f"❌ Erro de validação: {str(ve)}")
        st.info("💡 Dica: Use exatamente o modelo baixado e preencha apenas números nas colunas indicadas.")
    except Exception as e:
        st.error(f"❌ Erro inesperado: {str(e)}")
        st.error("Verifique o console do terminal para mais detalhes e tente novamente.")

# Rodapé
st.markdown("---")
st.markdown("<p style='text-align:center; color:#888;'>Desenvolvido com ❤️ por Pamella Vilela</p>", unsafe_allow_html=True)
