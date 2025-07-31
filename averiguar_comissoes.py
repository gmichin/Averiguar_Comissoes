import pandas as pd
from datetime import datetime
import os
import numpy as np

def encontrar_oferta_mais_proxima(df_ofertas, codproduto, data_venda):
    """Encontra a oferta mais recente anterior à data da venda"""
    ofertas_produto = df_ofertas[df_ofertas['COD'] == codproduto]
    ofertas_anteriores = ofertas_produto[ofertas_produto['Data'] <= data_venda]
    
    if len(ofertas_anteriores) > 0:
        return ofertas_anteriores.loc[ofertas_anteriores['Data'].idxmax()]
    return None

def criar_regras_comissao_kg():
    """Cria as regras para classificação de comissão por kg"""
    regras = {
        # Regra geral para LOURENCINI (todos os vendedores)
        'TODOS': {
            'grupo': ['LOURENCINI'] 
        },
        
        # Regras específicas por vendedor
        'FELIPE RAMALHO GOMES': {
            'grupo_codigos': {'BERGAMINI': [700]} 
        },
        'LUIZ FERNANDO VOLTERO BARBOSA': {
            'grupo_codigos': {
                'REDE PLUS': [812],
                'CHAMA': [812]
            }
        },
        'VALDENIR VOLTERO - PRETO': {
            'grupo_codigos': {'RICOY': [812, 937]},
            'razao_codigos': {'LATICINIO SOBERANO LTDA VILA ALPINA': [1707, 1708, 1709]}
        },
        'VERA LUCIA MUNIZ': {
            'grupo_codigos': {'MOTA NOVO': [812]},
            'razao_codigos': {'SUPERMERCADO FEDERZONI LTDA': [812]}
        },
        'PAMELA FERREIRA VIEIRA': {
            'grupo_codigos': {
                'REDE PLUS': [812],
                'VIOLETA': [812],
                'HANJO': [812]
            },
            'razao_codigos': {
                'RODOSNACK G E G LANCHONETE E RESTAURANTE': [812],
                'SUPERMERCADO CATANDUVA LTDA': [812],
                'MERCADINHO SUBLIME MARTINS LTDA': [812],
                'JMW FOODS DISTRIBUIDORA DE ALIMENTOS LTDA': [812],
                'JMW FOODS DISTRIBUIDORA DE ALIMENTOS LTD': [812]
            }
        },
        'ROSE VOLTERO': {
            'razao_codigos': {
                'SUPERMERCADO REMIX EMPORIO JAVRI LTDA': [812],
                'HORTIFRUTI CHACARA FLORA LTDA': [812],
                'JC MIXMERC LTDA': [812]
            }
        }
    }
    return regras

def pertence_comissao_kg(row, regras):
    """Verifica se o registro pertence à comissão por kg"""
    vendedor = str(row['VENDEDOR']).strip().upper()
    grupo = str(row['GRUPO']).strip().upper()
    razao = str(row['RAZAO']).strip().upper()
    codproduto = row['CODPRODUTO']
    descricao = str(row['DESCRICAO']).strip().upper()
    
    # Primeiro verifica a regra geral (LOURENCINI)
    if 'TODOS' in regras:
        if 'grupo' in regras['TODOS']:
            if grupo in regras['TODOS']['grupo']:
                return True
    
    # Depois verifica as regras específicas por vendedor
    if vendedor in regras:
        regras_vendedor = regras[vendedor]
        
        # Verifica por grupo com códigos específicos
        if 'grupo_codigos' in regras_vendedor:
            for grp, codigos in regras_vendedor['grupo_codigos'].items():
                if grupo == grp and (codproduto in codigos or 'TODOS' in codigos):
                    return True
        
        # Verifica por razão social com códigos específicos
        if 'razao_codigos' in regras_vendedor:
            for rz, codigos in regras_vendedor['razao_codigos'].items():
                if razao == rz:
                    if 'PURURUCA 1KG' in codigos:  # Caso especial do produto
                        return 'PURURUCA 1KG' in descricao
                    return codproduto in codigos
    return False

def processar_planilhas():
    caminho_origem = r"C:\Users\win11\OneDrive\Documentos\Margens de fechamento\Margem_250730 - wapp.xlsx"
    caminho_downloads = os.path.join(os.path.expanduser('~'), 'Downloads', 'Averiguar_Comissoes (MARGEM).xlsx')
    
    try:
        print("Iniciando processamento...")
        
        # 1. Ler a aba 'Base (3,5%)'
        print("Lendo aba Base (3,5%)...")
        df_base = pd.read_excel(caminho_origem, sheet_name='Base (3,5%)', header=8)
        print(f"Total de registros na base: {len(df_base)}")
        
        # Selecionar colunas
        colunas_base = ['CF', 'RAZAO', 'GRUPO', 'NF-E', 'DATA', 'VENDEDOR', 'CODPRODUTO',
                       'GRUPO PRODUTO', 'DESCRICAO', 'P. Com', 'Preço Venda ']
        df_base = df_base[colunas_base].rename(columns={'Preço Venda ': 'Preço_Venda'})
        
        # Converter e formatar dados
        df_base['DATA'] = pd.to_datetime(df_base['DATA']).dt.date
        df_base['CODPRODUTO'] = pd.to_numeric(df_base['CODPRODUTO'], errors='coerce').fillna(0).astype('int64')
        df_base['P. Com'] = (pd.to_numeric(df_base['P. Com'], errors='coerce') * 100).round().astype('Int64')
        df_base['Preço_Venda'] = pd.to_numeric(df_base['Preço_Venda'], errors='coerce')
        
        # 2. Ler a aba 'OFERTAS_VOG'
        print("Lendo aba OFERTAS_VOG...")
        df_ofertas = pd.read_excel(caminho_origem, sheet_name='OFERTAS_VOG')
        df_ofertas = df_ofertas[['COD', 'ITENS', '3%', 'Data']].dropna(subset=['COD', 'Data'])
        
        # Converter e formatar dados
        df_ofertas['Data'] = pd.to_datetime(df_ofertas['Data']).dt.date
        df_ofertas['COD'] = pd.to_numeric(df_ofertas['COD'], errors='coerce').fillna(0).astype('int64')
        df_ofertas['3%'] = pd.to_numeric(df_ofertas['3%'], errors='coerce')
        print(f"Total de ofertas cadastradas: {len(df_ofertas)}")
        
        # 3. Classificar registros
        print("\nClassificando registros por tipo de comissão...")
        
        # Criar regras de comissão por kg
        regras_comissao_kg = criar_regras_comissao_kg()
        
        # Aplicar regras para identificar comissão por kg
        df_base['Comissao_Kg'] = df_base.apply(lambda row: pertence_comissao_kg(row, regras_comissao_kg), axis=1)
        
        # Valores que devem ser verificados nas ofertas (1, -1, 3, -3)
        valores_verificar = [-3, -1, 1, 3]  
        
        # Valores que vão para "2%" (2, -2)
        valores_2porcento = [2, -2]         
        
        # Valores que vão para "Sem Oferta" (0)
        valores_sem_oferta = [0]            
        
        # Primeiro separar comissão por kg
        df_comissao_kg = df_base[df_base['Comissao_Kg'] == True].copy()
        df_restante = df_base[df_base['Comissao_Kg'] == False].copy()
        
        # Classificar o restante
        mask_verificar = df_restante['P. Com'].isin(valores_verificar)
        mask_2porcento = df_restante['P. Com'].isin(valores_2porcento)
        mask_sem_oferta = df_restante['P. Com'].isin(valores_sem_oferta)
        mask_outros = ~(mask_verificar | mask_2porcento | mask_sem_oferta)
        
        df_para_verificar = df_restante[mask_verificar].copy()
        df_2porcento = df_restante[mask_2porcento].copy()
        df_sem_oferta = df_restante[mask_sem_oferta].copy()
        df_outros = df_restante[mask_outros].copy()
        
        print(f"Registros para Comissão por Kg: {len(df_comissao_kg)}")
        print(f"Registros para verificar nas ofertas: {len(df_para_verificar)}")
        print(f"Registros para 2%: {len(df_2porcento)}")
        print(f"Registros para Sem Oferta: {len(df_sem_oferta)}")
        print(f"Outros registros: {len(df_outros)}\n")
        
        # 4. Processar verificações
        resultados = []
        
        if len(df_para_verificar) > 0:
            for idx, row in df_para_verificar.iterrows():
                cod = row['CODPRODUTO']
                data = row['DATA']
                p_com = row['P. Com']
                preco_venda = row['Preço_Venda']
                
                # Verificar correspondência exata primeiro
                oferta = df_ofertas[(df_ofertas['COD'] == cod) & (df_ofertas['Data'] == data)]
                
                # Se não encontrar, buscar a oferta mais recente anterior
                if len(oferta) == 0:
                    oferta_proxima = encontrar_oferta_mais_proxima(df_ofertas, cod, data)
                    if oferta_proxima is not None:
                        oferta = pd.DataFrame([oferta_proxima])
                        tipo = 'Data Proxima'
                    else:
                        tipo = 'Sem Oferta'
                else:
                    tipo = 'Exata'
                
                # Processar oferta encontrada
                if len(oferta) > 0:
                    oferta = oferta.iloc[0]
                    valor_oferta = oferta['3%']
                    
                    is_devolucao = str(row['CF']).startswith('DEV')
                    
                    if preco_venda >= valor_oferta:
                        comissao_correta = 3
                    else:
                        comissao_correta = 1
                    
                    if is_devolucao:
                        comissao_correta *= -1
                    
                    status = 'Correto' if comissao_correta == p_com else 'Incorreto'
                    
                    resultados.append({
                        **row.to_dict(),
                        'Preço_Oferta': valor_oferta,
                        'Data_Oferta': oferta['Data'],
                        'Comissão_Correta': comissao_correta,
                        'Status': status,
                        'Tipo': tipo
                    })
                else:
                    df_sem_oferta = pd.concat([df_sem_oferta, pd.DataFrame([row])])
        
        # 5. Criar DataFrames de resultados
        df_resultados = pd.DataFrame(resultados) if resultados else pd.DataFrame()
        df_corretos = df_resultados[df_resultados['Status'] == 'Correto'] if not df_resultados.empty else pd.DataFrame()
        df_incorretos = df_resultados[df_resultados['Status'] == 'Incorreto'] if not df_resultados.empty else pd.DataFrame()
        
        # 6. Preparar saída
        colunas_saida = ['CF', 'RAZAO', 'GRUPO', 'NF-E', 'DATA', 'VENDEDOR', 
                        'CODPRODUTO', 'GRUPO PRODUTO', 'DESCRICAO', 
                        'P. Com', 'Preço_Venda', 'Preço_Oferta']
        
        colunas_corretas = colunas_saida + ['Data_Oferta', 'Tipo']
        colunas_incorretas = colunas_saida + ['Comissão_Correta', 'Status', 'Data_Oferta', 'Tipo']
        
        # 7. Salvar resultados
        print(f"\nSalvando resultados em: {caminho_downloads}")
        with pd.ExcelWriter(caminho_downloads, engine='openpyxl') as writer:
            # Comissões Corretas
            if not df_corretos.empty:
                df_corretos[colunas_corretas].to_excel(
                    writer, sheet_name='Corretos', index=False)
            
            # Comissões Incorretas
            if not df_incorretos.empty:
                df_incorretos[colunas_incorretas].to_excel(
                    writer, sheet_name='Incorretos', index=False)
            
            # Registros sem oferta
            if not df_sem_oferta.empty:
                df_sem_oferta[colunas_saida[:-1]].to_excel(
                    writer, sheet_name='Sem Oferta', index=False)
            
            # Comissão por Kg
            if not df_comissao_kg.empty:
                df_comissao_kg[colunas_saida[:-1]].to_excel(
                    writer, sheet_name='Comissão por Kg', index=False)
            
            # 2%
            if not df_2porcento.empty:
                df_2porcento[colunas_saida[:-1]].to_excel(
                    writer, sheet_name='2%', index=False)
            
            # Outros registros (se houver)
            if not df_outros.empty:
                df_outros[colunas_saida[:-1]].to_excel(
                    writer, sheet_name='Outros', index=False)
        
        # 8. Resumo final
        print("\n=== RESUMO FINAL ===")
        print(f"Total de registros na base: {len(df_base)}")
        print(f"Registros com comissão por kg: {len(df_comissao_kg)}")
        print(f"Registros em 2%: {len(df_2porcento)}")
        print(f"Registros em Sem Oferta: {len(df_sem_oferta)}")
        print(f"Registros verificados nas ofertas: {len(df_para_verificar)}")
        
        if not df_resultados.empty:
            print(f"  → Corretos: {len(df_corretos)}")
            print(f"  → Incorretos: {len(df_incorretos)}")
            print(f"Tipo de correspondência:")
            print(f"  → Exatas: {len(df_resultados[df_resultados['Tipo'] == 'Exata'])}")
            print(f"  → Por data próxima: {len(df_resultados[df_resultados['Tipo'] == 'Data Proxima'])}")
        
        print(f"Outros registros: {len(df_outros)}")
        
        # Verificação adicional para LOURENCINI
        lourencini_total = len(df_base[df_base['GRUPO'].str.upper() == 'LOURENCINI'])
        lourencini_kg = len(df_comissao_kg[df_comissao_kg['GRUPO'].str.upper() == 'LOURENCINI'])
        print(f"\nVerificação LOURENCINI:")
        print(f"Total de registros LOURENCINI: {lourencini_total}")
        print(f"Registros LOURENCINI em Comissão por Kg: {lourencini_kg}")
        if lourencini_total > 0 and lourencini_kg < lourencini_total:
            print("AVISO: Nem todos os registros LOURENCINI foram para Comissão por Kg!")
        
        print("\nProcessamento concluído com sucesso!")
        
    except Exception as e:
        print(f"\nErro durante o processamento: {str(e)}")
        raise

if __name__ == "__main__":
    processar_planilhas()