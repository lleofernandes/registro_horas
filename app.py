import pdfplumber
import pandas as pd

pdf_path_hedgepoint = 'Project Builder - Relatório de Registros - HedgePoint.pdf'
pdf_path_imaps = 'Project Builder - Relatório de Registros - iMaps.pdf'

# Função para processar um PDF e retornar um DataFrame
def process_pdf(pdf_path, projeto):
    # Lista para armazenar todos os dados
    all_rows = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:            
                all_rows.extend(table[1:]) # Pula a primeira linha (cabeçalho) e adiciona o resto

    # Criando o DataFrame diretamente da lista de linhas
    df = pd.DataFrame(all_rows, columns=['A', 'B', 'C'])

    # Lista para armazenar os dados processados
    processed_data = []
    current_date = None
    current_name = None

    # Processando linha por linha
    for idx in range(len(df)):
        # Convertendo valores None para string vazia
        col_a = str(df.iloc[idx, 0]) if df.iloc[idx, 0] is not None else ""
        col_b = str(df.iloc[idx, 1]) if df.iloc[idx, 1] is not None else ""
        col_c = str(df.iloc[idx, 2]) if df.iloc[idx, 2] is not None else ""
        
        # Verifica se é uma data (formato DD/MM/YYYY)
        if (col_a.strip() and '/' in col_a and len(col_a.strip().split('/')) == 3):
            current_date = col_a
        # Se a linha tem todas as colunas preenchidas
        elif (col_a.strip() and col_b.strip() and col_c.strip()):
            # Se a coluna A contém 'Total do dia', usa o nome atual
            if 'Total do dia' in col_a:
                name_to_use = current_name
            else:
                name_to_use = col_a
                current_name = col_a
            
            # Adiciona os dados processados à lista
            processed_data.append({
                'Data': current_date,
                'Nome': name_to_use,
                'Horas': col_c,
                'Projeto': projeto
            })

    # Criando um novo DataFrame com os dados processados
    return pd.DataFrame(processed_data)

# Processando cada PDF
df_hedgepoint = process_pdf(pdf_path_hedgepoint, 'HedgePoint')
df_imaps = process_pdf(pdf_path_imaps, 'iMaps')

# Mesclando os DataFrames
final_df = pd.concat([df_hedgepoint, df_imaps], ignore_index=True)

# Reordenando as colunas
final_df = final_df[['Data', 'Horas', 'Projeto', 'Nome']]

print(final_df.head())

# Salvando o resultado em um arquivo Excel
final_df.to_excel('registros_horas.xlsx', index=False)
