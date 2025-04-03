import os
import pandas as pd

def consolidate_excel_sheets():
    try:
        # Obter o diretório do script atual
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Definir caminhos completos
        input_file = os.path.join(script_dir, 'SSA2.xlsx')
        output_file = os.path.join(script_dir, 'SSA2_Consolidado.xlsx')
        
        # Verificar se o arquivo de entrada existe
        if not os.path.exists(input_file):
            print(f"Erro: Arquivo {input_file} não encontrado!")
            return
        
        # Ler todas as abas do arquivo Excel
        xlsx = pd.ExcelFile(input_file)
        
        # Lista para armazenar todos os DataFrames
        all_dataframes = []
        
        # Percorrer todas as abas do arquivo
        for sheet_name in xlsx.sheet_names:
            # Ler a aba atual
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            
            # Adicionar coluna com o nome da aba original
            df['Aba_Original'] = sheet_name
            
            # Adicionar o DataFrame à lista
            all_dataframes.append(df)
        
        # Concatenar todos os DataFrames
        consolidated_df = pd.concat(all_dataframes, ignore_index=True)
        
        # Salvar o DataFrame consolidado em um novo arquivo Excel
        consolidated_df.to_excel(output_file, index=False)
        
        print(f"Caminho do script: {script_dir}")
        print(f"Arquivo de entrada: {input_file}")
        print(f"Arquivo de saída: {output_file}")
        print(f"Planilha consolidada salva com sucesso!")
        print(f"Total de linhas consolidadas: {len(consolidated_df)}")
        print(f"Abas processadas: {xlsx.sheet_names}")
    
    except Exception as e:
        print(f"Erro detalhado ao processar o arquivo: {e}")
        import traceback
        traceback.print_exc()

# Executar a consolidação
consolidate_excel_sheets()

# Pausa para ver a saída
input("Pressione Enter para sair...")