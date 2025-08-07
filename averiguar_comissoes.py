import pandas as pd
from datetime import datetime
import os
import numpy as np

def encontrar_oferta_mais_proxima(df_ofertas, codproduto, data_venda):
    try:
        # Conversão robusta dos dados de entrada
        cod = int(float(codproduto))
        data = pd.to_datetime(data_venda).date()
        
        # Filtrar ofertas para o código do produto
        ofertas_cod = df_ofertas[df_ofertas['COD'] == cod].copy()
        
        if ofertas_cod.empty:
            return None
            
        # Converter datas e ordenar
        ofertas_cod['Data'] = pd.to_datetime(ofertas_cod['Data']).dt.date
        ofertas_cod = ofertas_cod.sort_values('Data', ascending=False)
        
        # 1. Buscar oferta na data exata
        oferta_exata = ofertas_cod[ofertas_cod['Data'] == data]
        if not oferta_exata.empty:
            return oferta_exata.iloc[0]
        
        # 2. Buscar a oferta mais recente anterior
        ofertas_anteriores = ofertas_cod[ofertas_cod['Data'] < data]
        if not ofertas_anteriores.empty:
            return ofertas_anteriores.iloc[0]  # Já está ordenado
        
        return None
        
    except Exception as e:
        print(f"Erro ao buscar oferta para {codproduto}: {str(e)}")
        return None


def criar_regras_comissao_kg():
    regras = {
        'TODOS': {
            'grupo': ['LOURENCINI'] 
        },
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
                'JMW FOODS DISTRIBUIDORA DE ALIMENTOS LTD': [812],
                'MERCADINHO SUBLIME CUMBICA LTDA': [812]
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

def criar_regras_comissao_fixa():
    regras = {
        'geral': {
            '0%': {
                'grupos': [
                    'AKKI ATACADISTA', 'ANDORINHA', 'BERGAMINI', 'DA PRACA', 'DOVALE',
                    'MERCADAO ATACADISTA', 'REIMBERG', 'SEMAR', 'TRIMAIS', 'VOVO ZUZU'
                ],
                'razoes': [
                    'COMERCIO DE CARNES E ROTISSERIE DUTRA LT',
                    'SUPERMERCADO HIGAS ITAQUERA LTDA', 'DISTRIBUIDORA E COMERCIO UAI SP LTDA',
                    "GARFETO'S FORNECIMENTO DE REFEICOES LTDA", "LATICINIO SOBERANO LTDA VILA ALPINA",
                    "SAO LORENZO ALIMENTOS LTDA"
                ]
            },
            '3%': {
                'grupos': ['CALVO', 'CENCOSUD', 'CHAMA', 'ESTRELA AZUL', 'TENDA']
            },
            '1%': {
                'grupos': ['ROLDAO'],
                'razoes': ['SHOPPING FARTURA VALINHOS COMERCIO LTDA']
            }
        },
        'razoes_especificas': { 
            'ACOUGUE E ROTISSERIE E GRIL BEEF LTDA': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            },
            'COMERCIO DAN DOG DE ALIMENTOS E LANCHES': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            },
            'LATICIO SOBERANO LTDA.': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            },
            'LINO ASTROLINO JUNIOR': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            },
            'M.F RODRIGUES JUNIOR ACOUGUE': {
                '2%': ['MIUDOS BOVINOS', 'CORTES BOVINOS']  # SEM CORTES SUINOS
            },
            'MARIA DE LOURDES ALBUQUERQUE MINIMERCADO': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            },
            'MERCADO SHOPAN E PADARIA LTDA': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            },
            'NASCIMENTO E SILVA COMERCIO DE PRODUTOS': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            },
            'PADARIA E CONFEITARIA ENCANTO DOS PAES L': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            },
            'PADARIA MAO NA MASSA LTDA': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            },
            'PAES E DOCES COMENDADOR LTDA': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            },
            'PAES E DOCES LEKA LTDA': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            },
            'PAES E DOCES MICHELLI LTDA': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            },
            'PRODUTORA DE CHARQUE SERTAO LTDA': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            },
            'REI DO BIFE CASA DE CARNES ACOUGUE LTDA': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            },
            'STYLLUS GRILL COMERCIO DE CARNES LTDA': {
                '2%': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
            }
        },
        'grupos_especificos': {
            'STYLLUS': {
                '0%': {
                    'grupos_produto': ['TORRESMO', 'SALAME UAI', 'EMPANADOS']
                }
            },
            'ROSSI': {
                '3%': [1288, 1289, 1287, 937, 1698, 1701, 1587, 1700, 1586, 1699],
                '1%': [1265, 1266, 812, 1115, 798],
                '2%': {
                    'grupos_produto': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'SALGADOS SUINOS A GRANEL'],
                    'codigos': [700]
                },
                '0%': {
                    'grupos_produto': ['EMBUTIDOS', 'SALAME UAI']
                }
            },
            'MUSSA': {
                '2%': {
                    'grupos_produto': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
                },
            },
            'FRIGOSHOW': {
                '2%': {
                    'grupos_produto': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
                },
            },
            'CLAYTON': {
                '2%': {
                    'grupos_produto': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
                },
            },
            'ROD E RAF': {
                '2%': {
                    'grupos_produto': ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']
                },
            },
            'REDE PLUS': {
                '3%': {
                    'grupos_produto': ['CORTES SUINOS CONGELADOS', 'CORTES BOVINOS'],
                    'codigos': [812]
                }
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

def aplicar_regras_comissao_fixa(row, regras):
    """Aplica as regras de comissão fixa de forma dinâmica"""
    grupo = str(row['GRUPO']).strip().upper()
    razao = str(row['RAZAO']).strip().upper()
    codproduto = row['CODPRODUTO']
    grupo_produto = str(row['GRUPO PRODUTO']).strip().upper()
    is_devolucao = str(row['CF']).startswith('DEV')
    

    # --- NOVO BLOCO PARA RAZÕES ESPECÍFICAS ---
    if 'razoes_especificas' in regras:
        if razao in regras['razoes_especificas']:
            grupos_permitidos = regras['razoes_especificas'][razao]['2%']
            if grupo_produto in grupos_permitidos:
                return _ajustar_para_devolucao(2, is_devolucao)
            
    # 1. Verificar regras específicas
    if grupo == 'ROSSI':
        if grupo_produto in ['MIUDOS BOVINOS', 'CORTES SUINOS CONGELADOS', 'SALGADOS SUINOS A GRANEL']:
            return _ajustar_para_devolucao(2, is_devolucao)
        
        if codproduto == 700:
            return _ajustar_para_devolucao(2, is_devolucao)
        
    if grupo == 'REDE PLUS':
        if grupo_produto in ['CORTES SUINOS CONGELADOS', 'CORTES BOVINOS']:
            return _ajustar_para_devolucao(3, is_devolucao)
        
        if codproduto == 812:
            return _ajustar_para_devolucao(3, is_devolucao)
        
    # 2. Verifica outras regras específicas por grupo
    if grupo in regras['grupos_especificos']:
        regras_grupo = regras['grupos_especificos'][grupo]
        
        for porcentagem, condicoes in regras_grupo.items():
            # Se for lista simples de códigos
            if isinstance(condicoes, list):
                if codproduto in condicoes:
                    return _ajustar_para_devolucao(int(porcentagem.replace('%', '')), is_devolucao)
            
            # Se for dicionário com condições complexas
            elif isinstance(condicoes, dict):
                match = True
                
                # Verifica grupos de produto
                if 'grupos_produto' in condicoes:
                    if grupo_produto not in condicoes['grupos_produto']:
                        match = False
                
                # Verifica códigos específicos
                if 'codigos' in condicoes and match:
                    if codproduto not in condicoes['codigos']:
                        match = False
                
                if match:
                    return _ajustar_para_devolucao(int(porcentagem.replace('%', '')), is_devolucao)
    
    # 3. Verifica regras gerais
    for porcentagem, condicoes in regras['geral'].items():
        porcentagem_num = int(porcentagem.replace('%', ''))
        
        # Verifica por grupo
        if 'grupos' in condicoes:
            if grupo in condicoes['grupos']:
                return _ajustar_para_devolucao(porcentagem_num, is_devolucao)
        
        # Verifica por razão social
        if 'razoes' in condicoes:
            if razao in condicoes['razoes']:
                return _ajustar_para_devolucao(porcentagem_num, is_devolucao)
    
    return None

def _ajustar_para_devolucao(valor, is_devolucao):
    """Ajusta o valor da comissão para devoluções"""
    return valor if not is_devolucao else -valor

def processar_planilhas():
    caminho_origem = r"C:\Users\win11\Downloads\Margem_250806 - wapp.xlsx"
    caminho_downloads = os.path.join(os.path.expanduser('~'), 'Downloads', 'Averiguar_Comissoes (MARGEM).xlsx')
    
    try:
        print("=== INÍCIO DO PROCESSAMENTO ===")
        
        # 1. Ler os dados
        df_base = pd.read_excel(caminho_origem, sheet_name='Base (3,5%)', header=8)
        print(f"TOTAL DE REGISTROS NA BASE: {len(df_base)}")
        
        colunas_base = ['CF', 'RAZAO', 'GRUPO', 'NF-E', 'DATA', 'VENDEDOR', 'CODPRODUTO',
                       'GRUPO PRODUTO', 'DESCRICAO', 'P. Com', 'Preço Venda ']
        df_base = df_base[colunas_base].rename(columns={'Preço Venda ': 'Preço_Venda'})
        
        # Converter e formatar dados
        df_base['DATA'] = pd.to_datetime(df_base['DATA']).dt.date
        df_base['CODPRODUTO'] = pd.to_numeric(df_base['CODPRODUTO'], errors='coerce').fillna(0).astype('int64')
        df_base['P. Com'] = (pd.to_numeric(df_base['P. Com'], errors='coerce') * 100).round().astype('Int64')
        df_base['Preço_Venda'] = pd.to_numeric(df_base['Preço_Venda'], errors='coerce')
        
        df_ofertas = pd.read_excel(caminho_origem, sheet_name='OFERTAS_VOG')
        df_ofertas = df_ofertas[['COD', 'ITENS', '3%', 'Data']].dropna(subset=['COD', 'Data'])
        
        df_ofertas['Data'] = pd.to_datetime(df_ofertas['Data']).dt.date
        df_ofertas['COD'] = pd.to_numeric(df_ofertas['COD'], errors='coerce').fillna(0).astype('int64')
        df_ofertas['3%'] = pd.to_numeric(df_ofertas['3%'], errors='coerce')
        print(f"- Total de ofertas cadastradas: {len(df_ofertas)}")
        
         # 2. Aplicar regras de comissão por kg (primeiro filtro)
        regras_comissao_kg = criar_regras_comissao_kg()
        df_base['Comissao_Kg'] = df_base.apply(
            lambda row: pertence_comissao_kg(row, regras_comissao_kg), axis=1)
    
        df_comissao_kg = df_base[df_base['Comissao_Kg'] == True].copy()
        df_sem_kg = df_base[df_base['Comissao_Kg'] == False].copy()
    
        print(f"- Itens para comissão por kg: {len(df_comissao_kg)}")
    
        # 3. Aplicar TODAS as regras fixas (incluindo as de 2%)
        regras_comissao_fixa = criar_regras_comissao_fixa()
        df_sem_kg['Comissao_Esperada'] = df_sem_kg.apply(
            lambda row: aplicar_regras_comissao_fixa(row, regras_comissao_fixa), axis=1)
        
        # Separar os que tem regra aplicada dos que não tem
        mask_regras = df_sem_kg['Comissao_Esperada'].notna()
        df_regras = df_sem_kg[mask_regras].copy()
        df_sem_regra = df_sem_kg[~mask_regras].copy()
        
        # Verificar se a comissão aplicada está correta
        df_regras['Status'] = df_regras.apply(
            lambda row: 'Correto' if row['P. Com'] == row['Comissao_Esperada'] else 'Incorreto', axis=1)
        
        df_regras_corretas = df_regras[df_regras['Status'] == 'Correto']
        df_regras_incorretas = df_regras[df_regras['Status'] == 'Incorreto']
        
        print(f"- Registros com regras fixas aplicadas: {len(df_regras)}")
        print(f"  → Corretos: {len(df_regras_corretas)}")
        print(f"  → Incorretos: {len(df_regras_incorretas)}")
        
        # 4. Verificação das ofertas (apenas para os que não se enquadram nas regras anteriores)
        resultados_ofertas = []
        registros_sem_oferta = []
        logs_erros = []
        
        # Otimização: Criar dicionário de ofertas por código para acesso rápido
        ofertas_por_codigo = df_ofertas.groupby('COD')
        
        for idx, row in df_sem_regra.iterrows():
            try:
                cod = int(float(row['CODPRODUTO']))
                data = pd.to_datetime(row['DATA']).date()
                preco = float(row['Preço_Venda'])
                is_devolucao = str(row['CF']).startswith('DEV')
        
                # Verificar se existe oferta para este código
                if cod not in ofertas_por_codigo.groups:
                    registros_sem_oferta.append(row.to_dict())
                    continue
                
                # Buscar oferta específica
                ofertas_cod = ofertas_por_codigo.get_group(cod)
                oferta = encontrar_oferta_mais_proxima(ofertas_cod, cod, data)
        
                if oferta is not None:
                    preco_oferta = float(oferta['3%'])
                    
                    # Nova lógica de classificação:
                    if preco >= preco_oferta:
                        comissao = 3
                    else:
                        comissao = 1
        
                    if is_devolucao:
                        comissao *= -1
        
                    resultados_ofertas.append({
                        **row.to_dict(),
                        'Preço_Oferta': preco_oferta,
                        'Data_Oferta': oferta['Data'],
                        'Comissão_Correta': comissao,
                        'Status': 'Correto' if row['P. Com'] == comissao else 'Incorreto',
                        'Tipo': 'Exata' if oferta['Data'] == data else 'Data Proxima',
                        'Diferença_Preço': f"{(preco - preco_oferta)/preco_oferta:.2%}"
                    })
                else:
                    registros_sem_oferta.append(row.to_dict())
        
            except Exception as e:
                logs_erros.append({
                    'CODPRODUTO': row['CODPRODUTO'],
                    'DATA': row['DATA'],
                    'GRUPO': row['GRUPO'],
                    'RAZAO': row['RAZAO'],
                    'Mensagem': f'Erro ao processar: {str(e)}'
                })
    
        # Criar DataFrames de resultados
        df_resultados_ofertas = pd.DataFrame(resultados_ofertas) if resultados_ofertas else pd.DataFrame()
        df_sem_oferta_final = pd.DataFrame(registros_sem_oferta) if registros_sem_oferta else pd.DataFrame()
        df_logs_erros = pd.DataFrame(logs_erros) if logs_erros else pd.DataFrame()

        print(f"- Itens com oferta encontrada: {len(df_resultados_ofertas)}")
        print(f"- Itens sem oferta encontrada: {len(df_sem_oferta_final)}")
        print(f"- Erros durante o processamento: {len(df_logs_erros)}")

       # 5. Exportar para Excel
        print(f"\n8. SALVANDO RESULTADOS EM: {caminho_downloads}")
        with pd.ExcelWriter(caminho_downloads, engine='openpyxl') as writer:
            # 1. Comissão por Kg - Remover apenas Comissao_Kg e renomear aba (substituindo / por -)
            if not df_comissao_kg.empty:
                df_comissao_kg.drop(columns=['Comissao_Kg'], errors='ignore').to_excel(
                    writer, sheet_name='Comissão-kg', index=False)  # Alterado para hífen

            # 2. Regras Corretas - Reordenar e renomear colunas
            if not df_regras_corretas.empty:
                df_regras_corretas = df_regras_corretas.drop(
                    columns=['Comissao_Kg', 'Status'], errors='ignore')
                
                # Reordenar colunas e renomear
                colunas_ordenadas = ['RAZAO', 'GRUPO', 'NF-E', 'DATA', 'VENDEDOR', 'CODPRODUTO',
                                    'GRUPO PRODUTO', 'DESCRICAO', 'Preço_Venda', 'P. Com', 'Comissao_Esperada']
                df_regras_corretas = df_regras_corretas[colunas_ordenadas].rename(
                    columns={'Comissao_Esperada': 'O Com'})
                
                # Adicionar espaçamento
                df_regras_corretas = df_regras_corretas.style.set_properties(**{'text-align': 'left'})
                df_regras_corretas.to_excel(writer, sheet_name='O Regras', index=False)

            # 3. Regras Incorretas - Reordenar e renomear colunas
            if not df_regras_incorretas.empty:
                df_regras_incorretas = df_regras_incorretas.drop(
                    columns=['Comissao_Kg', 'Status'], errors='ignore')
                
                # Reordenar colunas e renomear
                colunas_ordenadas = ['RAZAO', 'GRUPO', 'NF-E', 'DATA', 'VENDEDOR', 'CODPRODUTO',
                                    'GRUPO PRODUTO', 'DESCRICAO', 'Preço_Venda', 'P. Com', 'Comissao_Esperada']
                df_regras_incorretas = df_regras_incorretas[colunas_ordenadas].rename(
                    columns={'Comissao_Esperada': 'O Com'})
                
                # Adicionar espaçamento
                df_regras_incorretas = df_regras_incorretas.style.set_properties(**{'text-align': 'left'})
                df_regras_incorretas.to_excel(writer, sheet_name='X Regras', index=False)

            # 4. Ofertas Corretas - Reordenar, renomear e remover colunas
            if not df_resultados_ofertas.empty:
                df_ofertas_corretas = df_resultados_ofertas[df_resultados_ofertas['Status'] == 'Correto']
                if not df_ofertas_corretas.empty:
                    df_ofertas_corretas = df_ofertas_corretas.drop(
                        columns=['Comissao_Kg', 'Comissao_Esperada', 'Diferença_Preço', 'Status', 'Tipo'], 
                        errors='ignore')
                    
                    # Reordenar colunas e renomear
                    colunas_ordenadas = ['CF', 'RAZAO', 'GRUPO', 'NF-E', 'VENDEDOR', 'CODPRODUTO',
                                       'GRUPO PRODUTO', 'DESCRICAO', 'Preço_Venda', 'DATA',
                                       'P. Com', 'Preço_Oferta', 'Data_Oferta', 'Comissão_Correta']
                    df_ofertas_corretas = df_ofertas_corretas[colunas_ordenadas].rename(
                        columns={'Comissão_Correta': 'O Com'})
                    
                    # Adicionar espaçamento
                    df_ofertas_corretas = df_ofertas_corretas.style.set_properties(**{'text-align': 'left'})
                    df_ofertas_corretas.to_excel(writer, sheet_name='O Ofertas', index=False)

            # 5. Ofertas Incorretas - Reordenar, renomear e remover colunas
            if not df_resultados_ofertas.empty:
                df_ofertas_incorretas = df_resultados_ofertas[df_resultados_ofertas['Status'] == 'Incorreto']
                if not df_ofertas_incorretas.empty:
                    df_ofertas_incorretas = df_ofertas_incorretas.drop(
                        columns=['Comissao_Kg', 'Comissao_Esperada', 'Diferença_Preço', 'Status', 'Tipo'], 
                        errors='ignore')
                    
                    # Reordenar colunas e renomear
                    colunas_ordenadas = ['CF', 'RAZAO', 'GRUPO', 'NF-E', 'VENDEDOR', 'CODPRODUTO',
                                       'GRUPO PRODUTO', 'DESCRICAO', 'Preço_Venda', 'DATA',
                                       'P. Com', 'Preço_Oferta', 'Data_Oferta', 'Comissão_Correta']
                    df_ofertas_incorretas = df_ofertas_incorretas[colunas_ordenadas].rename(
                        columns={'Comissão_Correta': 'O Com'})
                    
                    # Adicionar espaçamento
                    df_ofertas_incorretas = df_ofertas_incorretas.style.set_properties(**{'text-align': 'left'})
                    df_ofertas_incorretas.to_excel(writer, sheet_name='X Ofertas', index=False)

            # 6. Sem Oferta - Sem alterações além do drop já existente
            if not df_sem_oferta_final.empty:
                df_sem_oferta_final = df_sem_oferta_final.drop(
                    columns=['Comissao_Kg', 'Comissao_Esperada'], 
                    errors='ignore')
                
                # Adicionar espaçamento
                df_sem_oferta_final = df_sem_oferta_final.style.set_properties(**{'text-align': 'left'})
                df_sem_oferta_final.to_excel(writer, sheet_name='Sem Oferta', index=False)

            # 7. Logs de erros - Sem alterações
            if not df_logs_erros.empty:
                df_logs_erros = df_logs_erros.style.set_properties(**{'text-align': 'left'})
                df_logs_erros.to_excel(writer, sheet_name='Logs Erros', index=False)

        print("\n=== PROCESSAMENTO CONCLUÍDO COM SUCESSO ===")
        
    except Exception as e:
        print(f"\nERRO CRÍTICO DURANTE O PROCESSAMENTO: {str(e)}")
        raise

if __name__ == "__main__":
    processar_planilhas()