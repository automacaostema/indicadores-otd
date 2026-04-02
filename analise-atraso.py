import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime
import os

# Configuração da página Streamlit
st.set_page_config(page_title="Indicadores de Produção", layout="wide")

# Classe personalizada para controlar Cabeçalho e Rodapé do PDF
class PDF(FPDF):
    def __init__(self, nome_arquivo=""):
        super().__init__()
        self.nome_arquivo = nome_arquivos

    def header(self):
        if os.path.exists("logo.png"):
            self.image("logo.png", 10, 8, 33)
            self.ln(15)
        else:
            self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font("helvetica", "I", 8)
        texto_rodape = f"Adriano Toracelli - Analista da Qualidade - Pag {self.page_no()} de {{nb}}"
        self.cell(0, 10, texto_rodape.encode('latin-1', 'replace').decode('latin-1'), align='R')

st.title("📊 Analisador de OTD e Lead Time")

uploaded_file = st.file_uploader("Arraste o arquivo Excel (.xlsx) aqui", type="xlsx")

if uploaded_file:
    nome_do_arquivo = uploaded_file.name
    
    # 1. Leitura e Limpeza
    df = pd.read_excel(uploaded_file, engine='calamine', header=None)
    
    col_names = ['Emp.', 'Servico', 'Lancamento', 'Titulo Cliente', 'Extra_E', 
                 'N Pedido', 'Previsao de Entrega', 'Aprovado', 'Conclusao', 'Status']
    df.columns = col_names[:len(df.columns)]
    df = df.drop(columns=['Extra_E'], errors='ignore')
    
    # Remover duplicatas e converter datas
    df = df.drop_duplicates(subset='Servico').copy()
    colunas_data = ['Lancamento', 'Previsao de Entrega', 'Aprovado', 'Conclusao']
    for col in colunas_data:
        df[col] = pd.to_datetime(df[col], errors='coerce')

    # 2. Cálculos
    df['Lead Time'] = (df['Conclusao'] - df['Lancamento']).dt.days
    df['Dias Atraso'] = (df['Conclusao'] - df['Previsao de Entrega']).dt.days
    
    # Filtrar e Ordenar por maior atraso
    df_atrasados = df[df['Dias Atraso'] > 0].copy()
    df_atrasados = df_atrasados.sort_values(by='Dias Atraso', ascending=False)
    
    total_pedidos = len(df.dropna(subset=['Conclusao']))
    qtd_atraso = len(df_atrasados)
    no_prazo = total_pedidos - qtd_atraso
    otd = (no_prazo / total_pedidos) * 100 if total_pedidos > 0 else 0
    lt_medio = df['Lead Time'].mean()

    # 3. Dashboard na tela
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Serviços Únicos", len(df))
    c2.metric("Qtd. em Atraso", qtd_atraso)
    c3.metric("OTD %", f"{otd:.2f}%")
    c4.metric("Lead Time Médio", f"{lt_medio:.1f} dias")

    st.subheader(f"⚠️ Itens em Atraso - Arquivo: {nome_do_arquivo}")
    
    if not df_atrasados.empty:
        df_show = df_atrasados.copy()
        for col in colunas_data:
            df_show[col] = df_show[col].dt.strftime('%d/%m/%Y')
        st.dataframe(df_show[['Servico', 'Titulo Cliente', 'Previsao de Entrega', 'Conclusao', 'Dias Atraso']], use_container_width=True)
    else:
        st.success("Tudo em dia!")

    # 4. Função para Gerar PDF
    def gerar_pdf():
        pdf = PDF(nome_arquivo=nome_do_arquivo)
        pdf.alias_nb_pages()
        pdf.add_page()
        
        # Título
        pdf.set_font("helvetica", "B", 16)
        pdf.cell(0, 10, "Relatorio de OTD e Lead Time", ln=True, align='C')
        
        # Cabeçalho de Informações
        pdf.set_font("helvetica", "", 11)
        pdf.ln(5)
        pdf.cell(0, 7, f"Data do Relatorio: {datetime.now().strftime('%d/%m/%Y')}", ln=True)
        pdf.set_font("helvetica", "I", 10)
        pdf.cell(0, 7, f"Arquivo Origem: {nome_do_arquivo}", ln=True)
        pdf.ln(2)
        
        pdf.set_font("helvetica", "", 11)
        pdf.cell(0, 7, f"Total de Servicos Processados: {len(df)}", ln=True)
        pdf.cell(0, 7, f"Quantidade de Servicos em Atraso: {qtd_atraso}", ln=True)
        pdf.cell(0, 7, f"OTD Geral: {otd:.2f}%", ln=True)
        pdf.cell(0, 7, f"Lead Time Medio: {lt_medio:.2f} dias", ln=True)
        
        pdf.ln(10)
        pdf.set_font("helvetica", "B", 12)
        pdf.cell(0, 10, "Lista de Itens em Atraso (Prioridade por Dias):", ln=True)
        
        # Listagem de Atrasos
        pdf.set_font("helvetica", "", 8)
        for i, row in df_atrasados.iterrows():
            prev = row['Previsao de Entrega'].strftime('%d/%m/%Y')
            conc = row['Conclusao'].strftime('%d/%m/%Y')
            serv = str(row['Servico'])
            dias = str(row['Dias Atraso'])
            tit = str(row['Titulo Cliente'])[:25]
            
            txt = f"Servicos: {serv} | Prev: {prev} | Conclusao: {conc} | Atraso: {dias} dias | Item: {tit}"
            pdf.cell(0, 7, txt.encode('latin-1', 'replace').decode('latin-1'), ln=True, border='B')
        
        return pdf.output()

    # 5. Botão de Download
    st.download_button(
        label="📥 Exportar Relatório em PDF",
        data=bytes(gerar_pdf()),
        file_name=f"relatorio_otd_leadtime_{datetime.now().strftime('%d_%m')}.pdf",
        mime="application/pdf"
    )

    with st.expander("Ver base completa extraída"):
        st.dataframe(df)