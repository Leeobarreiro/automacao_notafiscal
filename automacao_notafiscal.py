import xmltodict
import pandas as pd
import os
from datetime import datetime

def extrair_dados_nota(arquivo_xml):
    with open(arquivo_xml, 'rb') as arquivo:
        dicionario = xmltodict.parse(arquivo)
        
    itens = []
    try:
        produtos = dicionario['nfeProc']['NFe']['infNFe']['det']
        data_emissao = dicionario['nfeProc']['NFe']['infNFe']['ide']['dhEmi'][:10]  # Pega apenas a data (YYYY-MM-DD)
        if isinstance(produtos, dict):  # Se for um único produto, transforma em lista
            produtos = [produtos]
        
        for item in produtos:
            produto = item['prod']['xProd']
            quantidade = float(item['prod']['qCom'])
            valor_unitario = float(item['prod']['vUnCom'])
            valor_total = float(item['prod']['vProd'])
            
            itens.append({
                'Produto': produto,
                'Data Emissão': datetime.strptime(data_emissao, '%Y-%m-%d').strftime('%Y-%m'),  # Converte para formato YYYY-MM
                'Quantidade Consumida': quantidade,
                'Valor Unitário': valor_unitario,
                'Valor Total': valor_total,
            })
    except Exception as e:
        print(f"Erro ao processar arquivo {arquivo_xml}: {e}")
    
    return itens

def main():
    local_atual = os.getcwd()
    caminho_pasta = os.path.join(local_atual, 'Notas_Fiscais')
    lista_arquivos = os.listdir(caminho_pasta)
    
    todos_itens = []
    
    for arquivo in lista_arquivos:
        if arquivo.endswith('.xml'):
            caminho_completo = os.path.join(caminho_pasta, arquivo)
            itens = extrair_dados_nota(caminho_completo)
            todos_itens.extend(itens)
    
    df = pd.DataFrame(todos_itens)
    
    if not df.empty:
        # Agrupa por Produto e Data Emissão
        df_consumo_mensal = df.groupby(['Produto', 'Data Emissão']).agg({
            'Quantidade Consumida': 'sum',
            'Valor Unitário': 'mean',
            'Valor Total': 'sum'
        }).reset_index()
        
        # Preparar a planilha final
        with pd.ExcelWriter('Resumo_Consumo_Material.xlsx') as writer:
            df.groupby('Produto').agg({
                'Quantidade Consumida': 'sum',
                'Valor Unitário': 'mean',
                'Valor Total': 'sum'
            }).reset_index().sort_values(by='Valor Total', ascending=False).to_excel(writer, sheet_name='Resumo', index=False)
            
            df.groupby('Produto').agg({
                'Quantidade Consumida': 'sum',
                'Valor Total': 'sum'
            }).reset_index().nlargest(5, 'Valor Total').to_excel(writer, sheet_name='Top 5 Valor', index=False)
            
            df.groupby('Produto').agg({
                'Quantidade Consumida': 'sum',
                'Valor Total': 'sum'
            }).reset_index().nlargest(5, 'Quantidade Consumida').to_excel(writer, sheet_name='Top 5 Quantidade', index=False)
            
            # Nova aba "Consumo Mensal"
            df_consumo_mensal.to_excel(writer, sheet_name='Consumo Mensal', index=False)
        
        print("Planilha 'Resumo_Consumo_Material.xlsx' criada com sucesso.")
    else:
        print("Nenhum dado encontrado para processar.")

if __name__ == "__main__":
    main()
