import pandas as pd
from datetime import datetime
import os
import numpy as np

def encontrar_oferta_mais_proxima(df_ofertas, codproduto, data_venda, coluna_comissao='3%'):
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

def encontrar_oferta_cb_mais_proxima(df_ofertas_cb, codproduto, data_venda):
    try:
        # Conversão robusta dos dados de entrada
        cod = int(float(codproduto))
        data = pd.to_datetime(data_venda).date()
        
        # Filtrar ofertas para o código do produto
        ofertas_cod = df_ofertas_cb[df_ofertas_cb['CD_PROD'] == cod].copy()
        
        if ofertas_cod.empty:
            return None
            
        # Converter datas e ordenar
        ofertas_cod['DT_REF'] = pd.to_datetime(ofertas_cod['DT_REF']).dt.date
        ofertas_cod = ofertas_cod.sort_values('DT_REF', ascending=False)
        
        # 1. Buscar oferta na data exata
        oferta_exata = ofertas_cod[ofertas_cod['DT_REF'] == data]
        if not oferta_exata.empty:
            return oferta_exata.iloc[0]
        
        # 2. Buscar a oferta mais recente anterior
        ofertas_anteriores = ofertas_cod[ofertas_cod['DT_REF'] < data]
        if not ofertas_anteriores.empty:
            return ofertas_anteriores.iloc[0]  # Já está ordenado
        
        return None
        
    except Exception as e:
        print(f"Erro ao buscar oferta CB para {codproduto}: {str(e)}")
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
                'MERCADINHO VILA NOVA BONSUCESSO LTDA': [812],
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
                'JC MIXMERC LTDA': [812],
                'SUPERMERCADO EMPORIO MIX LTDA': [812],
                'SUPERMERCADO DOM PETROPOLIS LTDA': [812]
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
                    'MERCADAO ATACADISTA', 'REIMBERG', 'SEMAR', 'TRIMAIS', 'VOVO ZUZU',
                    'BENGALA', 'OURINHOS'
                ],
                'razoes': [
                    'COMERCIO DE CARNES E ROTISSERIE DUTRA LT',
                    'DISTRIBUIDORA E COMERCIO UAI SP LTDA',
                    "GARFETO'S FORNECIMENTO DE REFEICOES LTDA", "LATICINIO SOBERANO LTDA VILA ALPINA",
                    "SAO LORENZO ALIMENTOS LTDA",
                    "QUE DELICIA MENDES COMERCIO DE ALIMENTOS",
                    "MARIANA OLIVEIRA MAZZEI"
                ]
            },
            '3%': {
                'grupos': ['CALVO', 'CHAMA', 'ESTRELA AZUL', 'TENDA', 'HIGAS'],
            },
            '1%': {
                'grupos': ['ROLDAO'],
                'razoes': ['SHOPPING FARTURA VALINHOS COMERCIO LTDA']
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
                '1%': [1265, 1266, 812, 1115, 798, 1211],
                '0%': {
                    'grupos_produto': ['EMBUTIDOS', 'EMBUTIDOS NOBRE', 'EMBUTIDOS SADIA', 
                                       'EMBUTIDOS PERDIGAO', 'EMBUTIDOS AURORA', 'EMBUTIDOS SEARA', 
                                       'SALAME UAI'],
                    'codigos': [1139]
                },
                '2%': {
                    'grupos_produto': ['MIUDOS BOVINOS', 'SUINOS', 'SALGADOS SUINOS A GRANEL'],
                    'codigos': [700]
                }
            },
            'REDE PLUS': {
                '3%': {
                    'grupos_produto': ['TEMPERADOS'],
                    'codigos': [812]
                }
            },
            'CENCOSUD': {
                '1%': {
                    'grupos_produto': ['SALAME UAI']
                },
                '3%': {
                    'todos_exceto': ['SALGADOS SUINOS EMBALADOS']
                }
            },
            'ROLDAO': {
                '0%': {
                    'grupos_produto': [
                        'CONGELADOS', 'CORTES BOVINOS', 'CORTES DE FRANGO', 'EMBUTIDOS', 
                        'EMBUTIDOS AURORA', 'EMBUTIDOS NOBRE', 'EMBUTIDOS PERDIGÃO', 
                        'EMBUTIDOS SADIA', 'EMBUTIDOS SEARA', 'EMPANADOS', 
                        'KITS FEIJOADDA', 'MIUDOS BOVINOS', 'SUINOS', 'TEMPERADOS'
                    ]
                },
                '1%': {
                    'todos_exceto': [
                        'CONGELADOS', 'CORTES BOVINOS', 'CORTES DE FRANGO', 'EMBUTIDOS', 
                        'EMBUTIDOS AURORA', 'EMBUTIDOS NOBRE', 'EMBUTIDOS PERDIGÃO', 
                        'EMBUTIDOS SADIA', 'EMBUTIDOS SEARA', 'EMPANADOS', 
                        'KITS FEIJOADDA', 'MIUDOS BOVINOS', 'SUINOS', 'TEMPERADOS'
                    ]
                }
            }
        },
        'razoes_especificas': {
            'PAES E DOCES LEKA LTDA': {
                '3%_codigos': [1893, 1886]
            },
            'PAES E DOCES MICHELLI LTDA': {
                '3%_codigos': [1893, 1886]
            },
            'WANDERLEY GOMES MORENO': {
                '3%_codigos': [1893, 1886]
            }
        }
    }
    return regras

def pertence_comissao_kg(row, regras):
    """Verifica se o registro pertence à comissão por kg"""
    # ... (mantido igual ao código original)
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
    # ... (mantido igual ao código original)
    vendedor = str(row['VENDEDOR']).strip().upper()
    grupo = str(row['GRUPO']).strip().upper()
    razao = str(row['RAZAO']).strip().upper()
    codproduto = row['CODPRODUTO']
    grupo_produto = str(row['GRUPO PRODUTO']).strip().upper()
    nfe = str(row['NF-E']).strip()
    is_devolucao = str(row['CF']).startswith('DEV')
    
    # --- NOVA REGRA: POR NF-E E PRODUTO ---
    # NF-E 107523 o produto 869 seria 1%
    if nfe == '107523' and codproduto == 869:
        return _ajustar_para_devolucao(1, is_devolucao)
    
    # --- NOVA REGRA: POR PRODUTO ---
    # Todos os produtos de código 1807 vai ser 1%
    if codproduto == 1807:
        return _ajustar_para_devolucao(1, is_devolucao)
    
    # --- VERIFICAÇÃO POR VENDEDOR ESPECÍFICO ---
    if 'vendedores_especificos' in regras:
        if vendedor in regras['vendedores_especificos']:
            regras_vendedor = regras['vendedores_especificos'][vendedor]
            
            for porcentagem, condicoes in regras_vendedor.items():
                if 'grupos_produto' in condicoes:
                    if grupo_produto in condicoes['grupos_produto']:
                        return _ajustar_para_devolucao(int(porcentagem.replace('%', '')), is_devolucao)
    
    # --- NOVA REGRA PARA CALVO - deve ser processada por ofertas ---
    if grupo == 'CALVO':
        # Se for MIUDOS BOVINOS, CORTES DE FRANGO ou SUINOS, processa por ofertas
        if grupo_produto in ['MIUDOS BOVINOS', 'CORTES DE FRANGO', 'SUINOS']:
            return None  # Retorna None para que seja processado por ofertas
        
        # Todo o resto do CALVO é 3%
        return _ajustar_para_devolucao(3, is_devolucao)
    
    if 'CENCOSUD' in grupo:
        if 'SALAME UAI' in grupo_produto:
            return _ajustar_para_devolucao(1, is_devolucao)
        return _ajustar_para_devolucao(3, is_devolucao)
    
    # --- BLOCO PARA RAZÕES ESPECÍFICAS ---
    if 'razoes_especificas' in regras:
        if razao in regras['razoes_especificas']:
            regras_razao = regras['razoes_especificas'][razao]
            
            for chave, valores in regras_razao.items():
                if chave.endswith('_codigos') and codproduto in valores:
                    porcentagem = int(chave.split('_')[0].replace('%', ''))
                    return _ajustar_para_devolucao(porcentagem, is_devolucao)
            
            for chave, valores in regras_razao.items():
                if not chave.endswith('_codigos') and grupo_produto in valores:
                    porcentagem = int(chave.replace('%', ''))
                    return _ajustar_para_devolucao(porcentagem, is_devolucao)
    
    # --- REGRAS ESPECÍFICAS POR GRUPO (COM ORDEM DE PRIORIDADE) ---
    if grupo == 'ROSSI':
        # PRIMEIRO verifica a regra de 0% (mais específica)
        if codproduto == 1139:
            return _ajustar_para_devolucao(0, is_devolucao)
        
        if grupo_produto in ['EMBUTIDOS', 'EMBUTIDOS NOBRE', 'EMBUTIDOS SADIA', 
                           'EMBUTIDOS PERDIGAO', 'EMBUTIDOS AURORA', 'EMBUTIDOS SEARA', 
                           'SALAME UAI']:
            return _ajustar_para_devolucao(0, is_devolucao)
        
        # DEPOIS verifica a regra de 2%
        if grupo_produto in ['MIUDOS BOVINOS', 'SUINOS', 'SALGADOS SUINOS A GRANEL']:
            return _ajustar_para_devolucao(2, is_devolucao)
        
        if codproduto == 700:
            return _ajustar_para_devolucao(2, is_devolucao)
        
        # FINALMENTE verifica as listas de códigos
        if codproduto in [1288, 1289, 1287, 937, 1698, 1701, 1587, 1700, 1586, 1699]:
            return _ajustar_para_devolucao(3, is_devolucao)
        
        if codproduto in [1265, 1266, 812, 1115, 798]:
            return _ajustar_para_devolucao(1, is_devolucao)
    
    if grupo == 'REDE PLUS':
        if grupo_produto in ['TEMPERADOS']:
            return _ajustar_para_devolucao(3, is_devolucao)
        
        if codproduto == 812:
            return _ajustar_para_devolucao(3, is_devolucao)
    
    # 2. Verifica outras regras específicas por grupo (mantendo a estrutura original)
    if grupo in regras['grupos_especificos']:
        regras_grupo = regras['grupos_especificos'][grupo]
        
        # Verificar primeiro as regras mais específicas (0%, depois 2%, etc.)
        for porcentagem in ['0%', '2%', '1%', '3%']:  # Ordem de prioridade
            if porcentagem in regras_grupo:
                condicoes = regras_grupo[porcentagem]
                
                if isinstance(condicoes, list):
                    if codproduto in condicoes:
                        return _ajustar_para_devolucao(int(porcentagem.replace('%', '')), is_devolucao)
                
                elif isinstance(condicoes, dict):
                    match = True
                    
                    if 'grupos_produto' in condicoes:
                        if grupo_produto not in condicoes['grupos_produto']:
                            match = False
                    
                    if 'codigos' in condicoes and match:
                        if codproduto not in condicoes['codigos']:
                            match = False
                    
                    if match:
                        return _ajustar_para_devolucao(int(porcentagem.replace('%', '')), is_devolucao)
    
    # 3. Verifica regras gerais
    for porcentagem, condicoes in regras['geral'].items():
        porcentagem_num = int(porcentagem.replace('%', ''))
        
        if 'grupos' in condicoes:
            if grupo in condicoes['grupos']:
                return _ajustar_para_devolucao(porcentagem_num, is_devolucao)
        
        if 'razoes' in condicoes:
            if razao in condicoes['razoes']:
                return _ajustar_para_devolucao(porcentagem_num, is_devolucao)
    
    return None

def _ajustar_para_devolucao(valor, is_devolucao):
    return valor if not is_devolucao else -valor

def processar_planilhas():
    caminho_origem = r"C:\Users\win11\Downloads\Margem_250925 - wapp.xlsx"
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
        
        # Ler as duas abas de ofertas
        df_ofertas_vog = pd.read_excel(caminho_origem, sheet_name='OFF_VOG')
        df_ofertas_vog = df_ofertas_vog[['COD', 'ITENS', '3%', 'Data']].dropna(subset=['COD', 'Data'])
        df_ofertas_vog['Data'] = pd.to_datetime(df_ofertas_vog['Data']).dt.date
        df_ofertas_vog['COD'] = pd.to_numeric(df_ofertas_vog['COD'], errors='coerce').fillna(0).astype('int64')
        df_ofertas_vog['3%'] = pd.to_numeric(df_ofertas_vog['3%'], errors='coerce')
        print(f"- Total de OFERTAS_VOG cadastradas: {len(df_ofertas_vog)}")
        
        df_ofertas_cb = pd.read_excel(caminho_origem, sheet_name='OFF_VOG_CB')
        df_ofertas_cb = df_ofertas_cb[['CD_PROD', 'GP_PROD', 'DS_PROD', '2%', '1%', 'DT_REF', 'QTDE_VENDAS', 'PK_OFF_CB']].dropna(subset=['CD_PROD', 'DT_REF'])
        df_ofertas_cb['DT_REF'] = pd.to_datetime(df_ofertas_cb['DT_REF']).dt.date
        df_ofertas_cb['CD_PROD'] = pd.to_numeric(df_ofertas_cb['CD_PROD'], errors='coerce').fillna(0).astype('int64')
        df_ofertas_cb['2%'] = pd.to_numeric(df_ofertas_cb['2%'], errors='coerce')
        print(f"- Total de OFERTAS_CB cadastradas: {len(df_ofertas_cb)}")
        
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
        
        # 4. Verificação das ofertas - PRIMEIRO CB, DEPOIS VOG
        resultados_ofertas_cb = []
        resultados_ofertas_vog = []
        registros_sem_oferta = []
        logs_erros = []
        
        # Otimização: Criar dicionários de ofertas por código para acesso rápido
        ofertas_cb_por_codigo = df_ofertas_cb.groupby('CD_PROD')
        ofertas_vog_por_codigo = df_ofertas_vog.groupby('COD')
        
        for idx, row in df_sem_regra.iterrows():
            try:
                cod = int(float(row['CODPRODUTO']))
                data = pd.to_datetime(row['DATA']).date()
                preco = float(row['Preço_Venda'])
                is_devolucao = str(row['CF']).startswith('DEV')
                grupo = str(row['GRUPO']).strip().upper()
                grupo_produto = str(row['GRUPO PRODUTO']).strip().upper()
        
                # PRIMEIRO: Verificar se existe oferta CB para este código
                if cod in ofertas_cb_por_codigo.groups:
                    # Buscar oferta CB específica
                    ofertas_cod_cb = ofertas_cb_por_codigo.get_group(cod)
                    oferta_cb = encontrar_oferta_cb_mais_proxima(ofertas_cod_cb, cod, data)
        
                    if oferta_cb is not None:
                        preco_oferta_cb = float(oferta_cb['2%'])
                        
                        # Aplicar desconto de 5% para grupos especiais
                        grupos_especiais = ['STYLLUS', 'ROD E RAF']
                        if grupo == 'CALVO' and grupo_produto in ['MIUDOS BOVINOS', 'CORTES DE FRANGO', 'SUINOS']:
                            grupos_especiais.append('CALVO')
                        
                        if grupo in grupos_especiais:
                            preco_comparacao = preco * 0.95  # Preço - 5%
                        else:
                            preco_comparacao = preco  # Mantém o preço normal
                        
                        # Lógica de classificação para CB: 2% se >=, 1% se <
                        if preco_comparacao >= preco_oferta_cb:
                            comissao = 2
                        else:
                            comissao = 1
            
                        if is_devolucao:
                            comissao *= -1
            
                        # Adicionar o preço -5% apenas para os grupos especiais
                        preco_menos_5 = preco * 0.95 if grupo in grupos_especiais else None
                        
                        resultados_ofertas_cb.append({
                            **row.to_dict(),
                            'Preço_Oferta': preco_oferta_cb,
                            'Preço - 5%': preco_menos_5,
                            'Data_Oferta': oferta_cb['DT_REF'],
                            'Comissão_Correta': comissao,
                            'Status': 'Correto' if row['P. Com'] == comissao else 'Incorreto',
                            'Tipo': 'CB',
                            'Tipo_Oferta': 'Exata' if oferta_cb['DT_REF'] == data else 'Data Proxima',
                            'Diferença_Preço': f"{(preco_comparacao - preco_oferta_cb)/preco_oferta_cb:.2%}" if preco_oferta_cb != 0 else 'Div/Zero'
                        })
                        continue  # Pula para o próximo registro, já encontrou em CB
        
                # SEGUNDO: Se não encontrou em CB, verificar em VOG
                if cod in ofertas_vog_por_codigo.groups:
                    # Buscar oferta VOG específica
                    ofertas_cod_vog = ofertas_vog_por_codigo.get_group(cod)
                    oferta_vog = encontrar_oferta_mais_proxima(ofertas_cod_vog, cod, data)
        
                    if oferta_vog is not None:
                        preco_oferta_vog = float(oferta_vog['3%'])
                        
                        # Aplicar desconto de 5% para grupos especiais
                        grupos_especiais = ['STYLLUS', 'ROD E RAF']
                        if grupo == 'CALVO' and grupo_produto in ['MIUDOS BOVINOS', 'CORTES DE FRANGO', 'SUINOS']:
                            grupos_especiais.append('CALVO')
                        
                        if grupo in grupos_especiais:
                            preco_comparacao = preco * 0.95  # Preço - 5%
                        else:
                            preco_comparacao = preco  # Mantém o preço normal
                        
                        # Lógica de classificação para VOG: 3% se >=, 1% se <
                        if preco_comparacao >= preco_oferta_vog:
                            comissao = 3
                        else:
                            comissao = 1
            
                        if is_devolucao:
                            comissao *= -1
            
                        # Adicionar o preço -5% apenas para os grupos especiais
                        preco_menos_5 = preco * 0.95 if grupo in grupos_especiais else None
                        
                        resultados_ofertas_vog.append({
                            **row.to_dict(),
                            'Preço_Oferta': preco_oferta_vog,
                            'Preço - 5%': preco_menos_5,
                            'Data_Oferta': oferta_vog['Data'],
                            'Comissão_Correta': comissao,
                            'Status': 'Correto' if row['P. Com'] == comissao else 'Incorreto',
                            'Tipo': 'VOG',
                            'Tipo_Oferta': 'Exata' if oferta_vog['Data'] == data else 'Data Proxima',
                            'Diferença_Preço': f"{(preco_comparacao - preco_oferta_vog)/preco_oferta_vog:.2%}" if preco_oferta_vog != 0 else 'Div/Zero'
                        })
                    else:
                        registros_sem_oferta.append(row.to_dict())
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
    
        # Combinar resultados de ofertas CB e VOG
        resultados_ofertas = resultados_ofertas_cb + resultados_ofertas_vog
        
        # Criar DataFrames de resultados
        df_resultados_ofertas = pd.DataFrame(resultados_ofertas) if resultados_ofertas else pd.DataFrame()
        df_sem_oferta_final = pd.DataFrame(registros_sem_oferta) if registros_sem_oferta else pd.DataFrame()
        df_logs_erros = pd.DataFrame(logs_erros) if logs_erros else pd.DataFrame()

        print(f"- Itens com oferta CB encontrada: {len(resultados_ofertas_cb)}")
        print(f"- Itens com oferta VOG encontrada: {len(resultados_ofertas_vog)}")
        print(f"- Itens sem oferta encontrada: {len(df_sem_oferta_final)}")
        print(f"- Erros durante o processamento: {len(df_logs_erros)}")

        # 5. Exportar para Excel
        print(f"\n8. SALVANDO RESULTADOS EM: {caminho_downloads}")
        with pd.ExcelWriter(caminho_downloads, engine='openpyxl') as writer:
            # 1. Comissão por Kg
            if not df_comissao_kg.empty:
                df_comissao_kg.drop(columns=['Comissao_Kg'], errors='ignore').to_excel(
                    writer, sheet_name='Comissão-kg', index=False)

            # 2. Regras Corretas
            if not df_regras_corretas.empty:
                df_regras_corretas = df_regras_corretas.drop(
                    columns=['Comissao_Kg', 'Status'], errors='ignore')
                
                colunas_ordenadas = ['RAZAO', 'GRUPO', 'NF-E', 'DATA', 'VENDEDOR', 'CODPRODUTO',
                                    'GRUPO PRODUTO', 'DESCRICAO', 'Preço_Venda', 'P. Com', 'Comissao_Esperada']
                df_regras_corretas = df_regras_corretas[colunas_ordenadas].rename(
                    columns={'Comissao_Esperada': 'O Com'})
                
                df_regras_corretas = df_regras_corretas.style.set_properties(**{'text-align': 'left'})
                df_regras_corretas.to_excel(writer, sheet_name='O Regras', index=False)

            # 3. Regras Incorretas
            if not df_regras_incorretas.empty:
                df_regras_incorretas = df_regras_incorretas.drop(
                    columns=['Comissao_Kg', 'Status'], errors='ignore')
                
                colunas_ordenadas = ['RAZAO', 'GRUPO', 'NF-E', 'DATA', 'VENDEDOR', 'CODPRODUTO',
                                    'GRUPO PRODUTO', 'DESCRICAO', 'Preço_Venda', 'P. Com', 'Comissao_Esperada']
                df_regras_incorretas = df_regras_incorretas[colunas_ordenadas].rename(
                    columns={'Comissao_Esperada': 'O Com'})
                
                df_regras_incorretas = df_regras_incorretas.style.set_properties(**{'text-align': 'left'})
                df_regras_incorretas.to_excel(writer, sheet_name='X Regras', index=False)

            # 4. Ofertas Corretas
            if not df_resultados_ofertas.empty:
                df_ofertas_corretas = df_resultados_ofertas[df_resultados_ofertas['Status'] == 'Correto']
                if not df_ofertas_corretas.empty:
                    df_ofertas_corretas = df_ofertas_corretas.drop(
                        columns=['Comissao_Kg', 'Comissao_Esperada', 'Diferença_Preço', 'Status', 'Tipo_Oferta'], 
                        errors='ignore')
                    
                    colunas_ordenadas = ['CF', 'RAZAO', 'GRUPO', 'NF-E', 'VENDEDOR', 'CODPRODUTO',
                                       'GRUPO PRODUTO', 'DESCRICAO', 'Preço_Venda', 'Preço - 5%', 'DATA',
                                       'P. Com', 'Preço_Oferta', 'Data_Oferta', 'Comissão_Correta', 'Tipo']
                    df_ofertas_corretas = df_ofertas_corretas[colunas_ordenadas].rename(
                        columns={'Comissão_Correta': 'O Com'})
                    
                    df_ofertas_corretas = df_ofertas_corretas.style.set_properties(**{'text-align': 'left'})
                    df_ofertas_corretas.to_excel(writer, sheet_name='O Ofertas', index=False)

            # 5. Ofertas Incorretas
            if not df_resultados_ofertas.empty:
                df_ofertas_incorretas = df_resultados_ofertas[df_resultados_ofertas['Status'] == 'Incorreto']
                if not df_ofertas_incorretas.empty:
                    df_ofertas_incorretas = df_ofertas_incorretas.drop(
                        columns=['Comissao_Kg', 'Comissao_Esperada', 'Diferença_Preço', 'Status', 'Tipo_Oferta'], 
                        errors='ignore')
                    
                    colunas_ordenadas = ['CF', 'RAZAO', 'GRUPO', 'NF-E', 'VENDEDOR', 'CODPRODUTO',
                                       'GRUPO PRODUTO', 'DESCRICAO', 'Preço_Venda', 'Preço - 5%', 'DATA',
                                       'P. Com', 'Preço_Oferta', 'Data_Oferta', 'Comissão_Correta', 'Tipo']
                    df_ofertas_incorretas = df_ofertas_incorretas[colunas_ordenadas].rename(
                        columns={'Comissão_Correta': 'O Com'})
                    
                    df_ofertas_incorretas = df_ofertas_incorretas.style.set_properties(**{'text-align': 'left'})
                    df_ofertas_incorretas.to_excel(writer, sheet_name='X Ofertas', index=False)

            # 6. Sem Oferta
            if not df_sem_oferta_final.empty:
                df_sem_oferta_final = df_sem_oferta_final.drop(
                    columns=['Comissao_Kg', 'Comissao_Esperada'], 
                    errors='ignore')
                
                df_sem_oferta_final = df_sem_oferta_final.style.set_properties(**{'text-align': 'left'})
                df_sem_oferta_final.to_excel(writer, sheet_name='Sem Oferta', index=False)

            # 7. Logs de erros
            if not df_logs_erros.empty:
                df_logs_erros = df_logs_erros.style.set_properties(**{'text-align': 'left'})
                df_logs_erros.to_excel(writer, sheet_name='Logs Erros', index=False)

        print("\n=== PROCESSAMENTO CONCLUÍDO COM SUCESSO ===")
        
    except Exception as e:
        print(f"\nERRO CRÍTICO DURANTE O PROCESSAMENTO: {str(e)}")
        raise

if __name__ == "__main__":
    processar_planilhas()