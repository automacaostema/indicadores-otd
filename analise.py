import pandas as pd

file_path = 'RU1-loja-1-marco.xlsx'

# Configurações de exibição
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.float_format', '{:.0f}'.format)

# 1. Ler o arquivo
df = pd.read_excel(file_path, engine='calamine', header=None)

# 2. Definir colunas
colunas_nomes = ['Emp.', 'Servico', 'Lancamento', 'Titulo Cliente', 'Extra_E', 
                 'N Pedido', 'Previsao de Entrega', 'Aprovado', 'Conclusao', 'Status']
df.columns = colunas_nomes[:len(df.columns)]
df = df.drop(columns=['Extra_E'], errors='ignore')

# 3. Remover duplicatas de Serviço (Mantém a primeira ocorrência)
df = df.drop_duplicates(subset='Servico').copy()

# 4. Converter colunas para Data
colunas_data = ['Lancamento', 'Previsao de Entrega', 'Aprovado', 'Conclusao']
for col in colunas_data:
    df[col] = pd.to_datetime(df[col], errors='coerce')

# 5. Cálculos de Indicadores
# Lead Time (Dias entre Lançamento e Conclusão)
df['Lead Time'] = (df['Conclusao'] - df['Lancamento']).dt.days

# Dias de Atraso (Positivo significa atraso, Negativo ou Zero significa no prazo)
df['Dias Atraso'] = (df['Conclusao'] - df['Previsao de Entrega']).dt.days

# 6. Identificar Atrasos (Atrasado se Conclusão > Previsão)
df_atrasados = df[df['Dias Atraso'] > 0].copy()

# 7. Cálculo do OTD (%)
total_pedidos = len(df.dropna(subset=['Conclusao']))
no_prazo = total_pedidos - len(df_atrasados)
otd = (no_prazo / total_pedidos) * 100 if total_pedidos > 0 else 0

# 8. Print dos Indicadores
print(f"--- INDICADORES ---")
print(f"Total de Serviços Únicos: {len(df)}")
print(f"OTD (On-Time Delivery): {otd:.2f}%")
print(f"Lead Time Médio: {df['Lead Time'].mean():.2f} dias")
print(f"\n--- ITENS EM ATRASO ---")

if df_atrasados.empty:
    print("Nenhum item em atraso.")
else:
    # Formatação para o print
    df_print = df_atrasados.copy()
    for col in colunas_data:
        df_print[col] = df_print[col].dt.strftime('%d/%m/%Y')
    
    print(df_print[['Servico', 'Titulo Cliente', 'Previsao de Entrega', 'Conclusao', 'Dias Atraso']])