import pandas as pd

# Carregar os dados da planilha enviada
file_path = 'valores_frete.xlsx'
df = pd.read_excel(file_path)

# Verificar as colunas do DataFrame
print("Colunas do DataFrame:", df.columns)

# Separar os dados por estado
estados = df['Estado'].unique()

# Criar um writer para salvar os dados em um novo arquivo Excel
output_path = 'valores_frete_separado_por_estado.xlsx'
with pd.ExcelWriter(output_path) as writer:
    for estado in estados:
        df_estado = df[df['Estado'] == estado]
        df_estado.to_excel(writer, sheet_name=estado, index=False)

print(f"Dados separados por estado foram salvos em {output_path}.")
