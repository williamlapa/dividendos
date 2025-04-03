import os
import pandas as pd
import tabula
import PyPDF2
from google.colab import files

def extract_pdf_to_excel(pdf_path, output_excel_path):
    try:
        # Método 1: Extração de tabelas com Tabula
        print("Tentando extrair tabelas do PDF...")
        
        # Tente extrair tabelas de todas as páginas
        dfs = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
        
        # Verifica se encontrou tabelas
        if dfs and len(dfs) > 0:
            # Consolidar tabelas encontradas
            consolidated_df = pd.concat(dfs, ignore_index=True)
            
            # Salvar tabelas em Excel
            with pd.ExcelWriter(output_excel_path) as writer:
                # Salvar tabelas consolidadas
                consolidated_df.to_excel(writer, sheet_name='Tabelas', index=False)
                
                # Salvar cada tabela original em uma aba separada
                for i, df in enumerate(dfs, 1):
                    df.to_excel(writer, sheet_name=f'Tabela_{i}', index=False)
            
            print(f"Tabelas extraídas com sucesso! Total de tabelas: {len(dfs)}")
        
        # Método 2: Extração de texto com PyPDF2
        if not dfs or len(dfs) == 0:
            print("Nenhuma tabela encontrada. Extraindo texto...")
            
            # Abrir PDF
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                
                # Extrair texto de todas as páginas
                text_pages = []
                for page in reader.pages:
                    text_pages.append(page.extract_text())
                
                # Converter texto para DataFrame
                text_df = pd.DataFrame(text_pages, columns=['Conteudo_Pagina'])
                
                # Salvar texto em Excel
                text_df.to_excel(output_excel_path, index=False, sheet_name='Texto_PDF')
                
                print("Texto extraído e salvo em Excel")
        
        # Download do arquivo
        files.download(output_excel_path)
        
        return True
    
    except Exception as e:
        print(f"Erro na extração: {e}")
        import traceback
        traceback.print_exc()
        return False

def process_pdf():
    # Upload do PDF
    print("Faça upload do PDF")
    uploaded = files.upload()
    
    if not uploaded:
        print("Nenhum arquivo PDF foi carregado!")
        return
    
    # Obter nome do arquivo
    pdf_path = list(uploaded.keys())[0]
    output_excel_path = 'PDF_Convertido.xlsx'
    
    # Executar conversão
    result = extract_pdf_to_excel(pdf_path, output_excel_path)
    
    if result:
        print("Conversão concluída com sucesso!")
    else:
        print("Falha na conversão do PDF")

# Executar processamento
process_pdf()