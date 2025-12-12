import pandas as pd
from datetime import datetime
import os
import numpy as np
from openpyxl.styles import numbers

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
            'grupo': ['REDE LOURENCINI'] 
        },
        'FELIPE RAMALHO GOMES': {
            'grupo_codigos': {
                'VAREJO BERGAMINI': [700],
                'REDE PEDREIRA': [700]
                } 
        },
        'LUIZ FERNANDO VOLTERO BARBOSA': {
            'grupo_codigos': {
                'REDE PLUS': [812],
                'REDE CHAMA': [812],
                'REDE PARANA': [812]
            }
        },
        'VALDENIR VOLTERO - PRETO': {
            'grupo_codigos': {'REDE RICOY': [812, 937, 1624]},
        },
        'VERA LUCIA MUNIZ': {
            'grupo_codigos': {'VAREJO MOTA NOVO': [812]},
            'razao_codigos': {'SUPERMERCADO FEDERZONI LTDA': [812]}
        },
        'PAMELA FERREIRA VIEIRA': {
            'grupo_codigos': {
                'REDE PLUS': [812],
                'REDE VIOLETA': [812],
                'REDE HANJO': [812]
            },
            'razao_codigos': {
                'MERCADINHO VILA NOVA BONSUCESSO LTDA': [812],
                'RODOSNACK G E G LANCHONETE E RESTAURANTE': [812],
                'SUPERMERCADO CATANDUVA LTDA': [812],
                'MERCADINHO SUBLIME MARTINS LTDA': [812],
                'JMW FOODS DISTRIBUIDORA DE ALIMENTOS LTDA': [812],
                'JMW FOODS DISTRIBUIDORA DE ALIMENTOS LTD': [812],
                'MERCADINHO SUBLIME CUMBICA LTDA': [812],
                'SUPER E DIST D ALIM E HORTF BRASIL LTDA': [812]
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
    return {
        'geral': {
            0.00: { 
                'grupos': [
                    'REDE AKKI', 'VAREJO ANDORINHA', 'VAREJO BERGAMINI', 'REDE DA PRACA', 'REDE DOVALE',
                    'REDE MERCADAO', 'REDE REIMBERG', 'REDE SEMAR', 'REDE TRIMAIS', 'REDE VOVO ZUZU',
                    'REDE BENGALA', 'VAREJO OURINHOS', 'REDE RICOY'
                ],
                'razoes': [
                    'COMERCIO DE CARNES E ROTISSERIE DUTRA LT',
                    'COMERCIO DE CARNES E ROTISSERIE DUTRA LTDA',
                    'DISTRIBUIDORA E COMERCIO UAI SP LTDA',
                    "GARFETO'S FORNECIMENTO DE REFEICOES LTDA", 
                    "LATICINIO SOBERANO LTDA VILA ALPINA",
                    "SAO LORENZO ALIMENTOS LTDA",
                    "QUE DELICIA MENDES COMERCIO DE ALIMENTOS",
                    "MARIANA OLIVEIRA MAZZEI",
                    "LS SANTOS COMERCIO DE ALIMENTOS LTDA",
                    "MERCADINHO LESSA LTDA",
                    "JSV SUPERMERCADOS EIRELI- LOJA 3"
                ]
            },
            0.03: {
                'grupos': ['VAREJO CALVO', 'REDE CHAMA', 'REDE ESTRELA AZUL', 'REDE TENDA', 'REDE HIGAS']
            },
            0.01: { 
                'razoes': ['SHOPPING FARTURA VALINHOS COMERCIO LTDA']
            }
        },
        'grupos_especificos': {
            'REDE STYLLUS': {
                0.00: {
                    'grupos_produto': ['TORRESMO', 'SALAME UAI', 'EMPANADOS']
                }
            },
            'REDE ROSSI': {
                0.03: {
                    'codigos': [937, 1698, 1701, 1587, 1700, 1586, 1699, 943, 1735, 1624, 1134]
                },
                0.01: [1265, 1266, 812, 1115, 798, 1211],
                0.00: {
                    'grupos_produto': [
                        'EMBUTIDOS', 'EMBUTIDOS NOBRE', 'EMBUTIDOS SADIA', 
                        'EMBUTIDOS PERDIGAO', 'EMBUTIDOS AURORA', 'EMBUTIDOS SEARA', 
                        'SALAME UAI'
                    ],
                    'codigos': [1139]
                },
                0.02: {
                    'grupos_produto': ['MIUDOS BOVINOS', 'SUINOS', 'SALGADOS SUINOS A GRANEL'],
                    'codigos': [700]
                }
            },
            'REDE PLUS': {
                0.03: {
                    'grupos_produto': ['TEMPERADOS'],
                    'codigos': [812]
                }
            },
            'REDE CENCOSUD': {
                0.01: {
                    'grupos_produto': ['SALAME UAI']
                },
            },
            'REDE AYUMI': {
                0.03: {
                    'grupos_produto': ['SALAME UAI']
                },
            },
            'REDE ROLDAO': {
                0.02: {
                    'grupos_produto': [
                        'CONGELADOS', 'CORTES BOVINOS', 'CORTES DE FRANGO', 'EMBUTIDOS', 
                        'EMBUTIDOS AURORA', 'EMBUTIDOS NOBRE', 'EMBUTIDOS PERDIGÃO', 
                        'EMBUTIDOS SADIA', 'EMBUTIDOS SEARA', 'EMPANADOS', 
                        'KITS FEIJOADA', 'MIUDOS BOVINOS', 'SUINOS', 'TEMPERADOS'
                    ]
                },
                0.00: {
                    'todos_exceto': [
                        'CONGELADOS', 'CORTES BOVINOS', 'CORTES DE FRANGO', 'EMBUTIDOS', 
                        'EMBUTIDOS AURORA', 'EMBUTIDOS NOBRE', 'EMBUTIDOS PERDIGÃO', 
                        'EMBUTIDOS SADIA', 'EMBUTIDOS SEARA', 'EMPANADOS', 
                        'KITS FEIJOADA', 'MIUDOS BOVINOS', 'SUINOS', 'TEMPERADOS'
                    ]
                }
            }
        },
        'razoes_especificas': {
            'PAES E DOCES LEKA LTDA': {
                0.03: [1893, 1886]
            },
            'PAES E DOCES MICHELLI LTDA': {
                0.03: [1893, 1886]
            },
            'WANDERLEY GOMES MORENO': {
                0.03: [1893, 1886]
            }
        }
    }

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
    vendedor = str(row['VENDEDOR']).strip().upper()
    grupo = str(row['GRUPO']).strip().upper()
    razao = str(row['RAZAO']).strip().upper()
    codproduto = row['CODPRODUTO']
    grupo_produto = str(row['GRUPO PRODUTO']).strip().upper()
    nfe = str(row['NF-E']).strip()
    is_devolucao = str(row['CF']).startswith('DEV')

    if nfe == '111880' and codproduto == 1950:
        return _ajustar_para_devolucao(0.01, is_devolucao)  # Alterado para 1% como decimal

    # --- NOVA REGRA: POR PRODUTO ---
    if codproduto == 1807 or codproduto == 947 or codproduto == 1914 or codproduto == 2000:
        return _ajustar_para_devolucao(0.01, is_devolucao)  # Alterado para 1% como decimal
    
    if vendedor == "PROPRIO":
        return _ajustar_para_devolucao(0.00, is_devolucao)  # Alterado para 0% como decimal
    
    # --- NOVA REGRA: TODOS OS ITENS DA REDE RICOY = 0% ---
    if grupo == 'REDE RICOY':
        return _ajustar_para_devolucao(0.00, is_devolucao)  # Alterado para 0% como decimal
    
    # --- REGRA ESPECÍFICA PARA REDE ROLDAO ---
    if grupo == 'REDE ROLDAO':
        grupos_2_percent = [
            'CONGELADOS', 'CORTES BOVINOS', 'CORTES DE FRANGO', 'EMBUTIDOS', 
            'EMBUTIDOS AURORA', 'EMBUTIDOS NOBRE', 'EMBUTIDOS PERDIGÃO', 
            'EMBUTIDOS SADIA', 'EMBUTIDOS SEARA', 'EMPANADOS', 
            'KITS FEIJOADA', 'MIUDOS BOVINOS', 'SUINOS', 'TEMPERADOS'
        ]
        
        if grupo_produto in grupos_2_percent:
            return _ajustar_para_devolucao(0.02, is_devolucao)  # Alterado para 2% como decimal
        else:
            return _ajustar_para_devolucao(0.00, is_devolucao)  # Alterado para 0% como decimal

    # --- NOVA REGRA PARA CALVO - deve ser processada por ofertas ---
    if grupo == 'VAREJO CALVO':
        # Se for MIUDOS BOVINOS, CORTES DE FRANGO ou SUINOS, processa por ofertas
        if grupo_produto in ['MIUDOS BOVINOS', 'CORTES DE FRANGO', 'SUINOS']:
            return None  # Retorna None para que seja processado por ofertas
        
        # Todo o resto do CALVO é 3%
        return _ajustar_para_devolucao(0.03, is_devolucao)  # Alterado para 3% como decimal
    
    if 'REDE CENCOSUD' in grupo:
        if 'SALAME UAI' in grupo_produto:
            return _ajustar_para_devolucao(0.01, is_devolucao)  # Alterado para 1% como decimal
        return _ajustar_para_devolucao(0.03, is_devolucao)  # Alterado para 3% como decimal
    
    # --- REGRAS ESPECÍFICAS POR GRUPO (COM ORDEM DE PRIORIDADE) ---
    if grupo == 'REDE ROSSI':
        # PRIMEIRO verifica a regra de TORRESMO (3%) - APENAS POR CÓDIGOS
        if codproduto in [937, 1698, 1701, 1587, 1700, 1586, 1699, 943, 1735, 1624, 1134]:
            return _ajustar_para_devolucao(0.03, is_devolucao)  # Alterado para 3% como decimal
        
        # SEGUNDO verifica a regra de 0% 
        if codproduto == 1139:
            return _ajustar_para_devolucao(0.00, is_devolucao)  # Alterado para 0% como decimal
        
        if grupo_produto in ['EMBUTIDOS', 'EMBUTIDOS NOBRE', 'EMBUTIDOS SADIA', 
                           'EMBUTIDOS PERDIGAO', 'EMBUTIDOS AURORA', 'EMBUTIDOS SEARA', 
                           'SALAME UAI']:
            return _ajustar_para_devolucao(0.00, is_devolucao)  # Alterado para 0% como decimal
        
        # TERCEIRO verifica a regra de 2%
        if grupo_produto in ["MIUDOS BOVINOS", "SUINOS", "SALGADOS SUINOS A GRANEL", 
                             "SALGADOS SUINOS EMBALADOS", "CORTES DE FRANGO"]:
            return _ajustar_para_devolucao(0.02, is_devolucao)  # Alterado para 2% como decimal
        
        if codproduto == 700:
            return _ajustar_para_devolucao(0.02, is_devolucao)  # Alterado para 2% como decimal
        
        # QUARTO verifica a regra de 1%
        if codproduto in [1265, 1266, 812, 1115, 798, 1211]:
            return _ajustar_para_devolucao(0.01, is_devolucao)  # Alterado para 1% como decimal
    
    if grupo == 'REDE PLUS':
        if grupo_produto in ['TEMPERADOS']:
            return _ajustar_para_devolucao(0.03, is_devolucao)  # Alterado para 3% como decimal
        
        if codproduto == 812:
            return _ajustar_para_devolucao(0.03, is_devolucao)  # Alterado para 3% como decimal
    
    # 2. Verifica outras regras específicas por grupo
    if grupo in regras['grupos_especificos']:
        regras_grupo = regras['grupos_especificos'][grupo]
        
        # Verificar primeiro as regras mais específicas (0%, depois 2%, etc.)
        for porcentagem in [0.00, 0.02, 0.01, 0.03]:  # Manter como floats
            if porcentagem in regras_grupo:
                condicoes = regras_grupo[porcentagem]
                
                if isinstance(condicoes, list):
                    if codproduto in condicoes:
                        return _ajustar_para_devolucao(porcentagem, is_devolucao)
                
                elif isinstance(condicoes, dict):
                    match = True
                    
                    if 'grupos_produto' in condicoes:
                        if grupo_produto not in condicoes['grupos_produto']:
                            match = False
                    
                    if 'codigos' in condicoes and match:
                        if codproduto not in condicoes['codigos']:
                            match = False
                    
                    if match:
                        return _ajustar_para_devolucao(porcentagem, is_devolucao)
    
    # 3. Verifica regras gerais
    for porcentagem, condicoes in regras['geral'].items():
        
        if 'grupos' in condicoes:
            if grupo in condicoes['grupos']:
                return _ajustar_para_devolucao(porcentagem, is_devolucao)
        
        if 'razoes' in condicoes:
            if razao in condicoes['razoes']:
                return _ajustar_para_devolucao(porcentagem, is_devolucao)
    
    return None

def _ajustar_para_devolucao(valor, is_devolucao):
    return valor if not is_devolucao else -valor

def _converter_para_decimal_percentual(valor):
    """
    Converte um valor para decimal de porcentagem.
    Pode lidar com strings com vírgula ou ponto decimal.
    """
    try:
        if pd.isna(valor):
            return None
        
        # Se for string, substituir vírgula por ponto
        if isinstance(valor, str):
            valor = valor.replace(',', '.').replace('%', '').strip()
        
        # Converter para float
        valor_float = float(valor)
        
        # Se o valor for maior que 1, provavelmente está em formato decimal
        # (ex: 1.11 significa 1.11% = 0.0111)
        if valor_float > 1:
            return valor_float / 100.0
        else:
            return valor_float
            
    except Exception as e:
        print(f"Erro ao converter valor '{valor}': {str(e)}")
        return None

def _comparar_comissoes(comissao_atual, comissao_esperada, decimal_places=4):
    """
    Compara duas comissões arredondando para N casas decimais
    Retorna True se forem iguais após o arredondamento
    """
    try:
        # Converter ambos os valores para decimal percentual
        atual = _converter_para_decimal_percentual(comissao_atual)
        esperada = _converter_para_decimal_percentual(comissao_esperada)
        
        if atual is None or esperada is None:
            return False
        
        # Arredondar para o número especificado de casas decimais
        atual_rounded = round(atual, decimal_places)
        esperada_rounded = round(esperada, decimal_places)
        
        # Debug: mostrar comparações se necessário
        # print(f"DEBUG: atual={atual}, esperada={esperada}, arred_atual={atual_rounded}, arred_esperada={esperada_rounded}, iguais={atual_rounded == esperada_rounded}")
        
        return atual_rounded == esperada_rounded
    except (ValueError, TypeError) as e:
        print(f"Erro na comparação: {str(e)}, atual={comissao_atual}, esperada={comissao_esperada}")
        return False

def processar_planilhas():
    caminho_origem = r"C:\Users\win11\OneDrive\Documentos\Margens de fechamento\MRG_251130 - Fechamento.xlsx"
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
        
        # REMOVER CONVERSÃO PARA INTEIRO - manter como float
        # IMPORTANTE: Converter vírgula para ponto decimal
        df_base['P. Com'] = df_base['P. Com'].apply(
            lambda x: float(str(x).replace(',', '.')) if isinstance(x, str) else float(x)
        )
        df_base['Preço_Venda'] = df_base['Preço_Venda'].apply(
            lambda x: float(str(x).replace(',', '.')) if isinstance(x, str) else float(x)
        )
        
        # Ler as duas abas de ofertas
        # AGORA LER A NOVA COLUNA "2%" NA ABA OFF_VOG
        df_ofertas_vog = pd.read_excel(caminho_origem, sheet_name='OFF_VOG')
        # Verifica se a coluna '2%' existe, se não existir, cria uma coluna vazia
        colunas_ofertas_vog = ['COD', 'ITENS', '3%', 'Data']
        if '2%' in df_ofertas_vog.columns:
            colunas_ofertas_vog.insert(2, '2%')  # Insere '2%' entre 'ITENS' e '3%'
        df_ofertas_vog = df_ofertas_vog[colunas_ofertas_vog].dropna(subset=['COD', 'Data'])
        df_ofertas_vog['Data'] = pd.to_datetime(df_ofertas_vog['Data']).dt.date
        df_ofertas_vog['COD'] = pd.to_numeric(df_ofertas_vog['COD'], errors='coerce').fillna(0).astype('int64')
        
        # Converter vírgula para ponto decimal nas ofertas
        df_ofertas_vog['3%'] = df_ofertas_vog['3%'].apply(
            lambda x: float(str(x).replace(',', '.')) if isinstance(x, str) else float(x)
        )
        # Se existir a coluna '2%', converte para numérico
        if '2%' in df_ofertas_vog.columns:
            df_ofertas_vog['2%'] = df_ofertas_vog['2%'].apply(
                lambda x: float(str(x).replace(',', '.')) if isinstance(x, str) else float(x)
            )
        print(f"- Total de OFERTAS_VOG cadastradas: {len(df_ofertas_vog)}")
        
        # AGORA LER A NOVA COLUNA "3%" NA ABA OFF_VOG_CB (entre 'DS_PROD' e '2%')
        df_ofertas_cb = pd.read_excel(caminho_origem, sheet_name='OFF_VOG_CB')
        # Verificar se a coluna '3%' existe
        colunas_ofertas_cb = ['CD_PROD', 'GP_PROD', 'DS_PROD', '2%', '1%', 'DT_REF', 'QTDE_VENDAS', 'PK_OFF_CB']
        if '3%' in df_ofertas_cb.columns:
            # Inserir '3%' entre 'DS_PROD' e '2%'
            colunas_ofertas_cb = ['CD_PROD', 'GP_PROD', 'DS_PROD', '3%', '2%', '1%', 'DT_REF', 'QTDE_VENDAS', 'PK_OFF_CB']
        
        df_ofertas_cb = df_ofertas_cb[colunas_ofertas_cb].dropna(subset=['CD_PROD', 'DT_REF'])
        df_ofertas_cb['DT_REF'] = pd.to_datetime(df_ofertas_cb['DT_REF']).dt.date
        df_ofertas_cb['CD_PROD'] = pd.to_numeric(df_ofertas_cb['CD_PROD'], errors='coerce').fillna(0).astype('int64')
        
        # Converter vírgula para ponto decimal nas ofertas CB
        df_ofertas_cb['2%'] = df_ofertas_cb['2%'].apply(
            lambda x: float(str(x).replace(',', '.')) if isinstance(x, str) else float(x)
        )
        df_ofertas_cb['1%'] = df_ofertas_cb['1%'].apply(
            lambda x: float(str(x).replace(',', '.')) if isinstance(x, str) else float(x)
        )
        
        # Se existir a coluna '3%', converter para numérico
        if '3%' in df_ofertas_cb.columns:
            df_ofertas_cb['3%'] = df_ofertas_cb['3%'].apply(
                lambda x: float(str(x).replace(',', '.')) if isinstance(x, str) else float(x)
            )
        
        print(f"- Total de OFERTAS_CB cadastradas: {len(df_ofertas_cb)}")
        if '3%' in df_ofertas_cb.columns:
            print(f"- Coluna '3%' encontrada em OFF_VOG_CB: {df_ofertas_cb['3%'].notna().sum()} registros com valores")
        
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
        
        # Verificar se a comissão aplicada está correta usando arredondamento de 2 casas
        df_regras['Status'] = df_regras.apply(
            lambda row: 'Correto' if _comparar_comissoes(row['P. Com'], row['Comissao_Esperada'], 4) else 'Incorreto', 
            axis=1)
        
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
                        # VERIFICAR SE TEM COLUNA '3%' NA OFERTA CB
                        tem_3_percent_cb = '3%' in oferta_cb and pd.notna(oferta_cb['3%']) and oferta_cb['3%'] > 0
                        tem_2_percent_cb = pd.notna(oferta_cb['2%']) and oferta_cb['2%'] > 0
                        
                        if tem_3_percent_cb:
                            # NOVA LÓGICA COM 3 FAIXAS (3%, 2%, 1%) para CB
                            preco_oferta_3_cb = float(oferta_cb['3%'])
                            preco_oferta_2_cb = float(oferta_cb['2%'])
                            
                            # Aplicar desconto de 5% para grupos especiais
                            grupos_especiais = ['REDE STYLLUS', 'REDE ROD E RAF']
                            if grupo == 'VAREJO CALVO' and grupo_produto in ['MIUDOS BOVINOS', 'CORTES DE FRANGO', 'SUINOS']:
                                grupos_especiais.append('VAREJO CALVO')
                            
                            if grupo in grupos_especiais:
                                preco_comparacao = preco * 0.95  # Preço - 5%
                            else:
                                preco_comparacao = preco  # Mantém o preço normal
                            
                            # Lógica de classificação para CB com 3 faixas
                            if preco_comparacao >= preco_oferta_3_cb:
                                comissao = 0.03  # Alterado para 3% como decimal
                            elif preco_comparacao >= preco_oferta_2_cb:
                                comissao = 0.02  # Alterado para 2% como decimal
                            else:
                                comissao = 0.01  # Alterado para 1% como decimal
                                
                            preco_oferta_2 = preco_oferta_2_cb
                            preco_oferta_3 = preco_oferta_3_cb
                            tem_2_percent = True
                            
                        else:
                            # LÓGICA ANTIGA (apenas 2%, 1%) para CB
                            preco_oferta_2_cb = float(oferta_cb['2%'])
                            
                            # Aplicar desconto de 5% para grupos especiais
                            grupos_especiais = ['REDE STYLLUS', 'REDE ROD E RAF']
                            if grupo == 'VAREJO CALVO' and grupo_produto in ['MIUDOS BOVINOS', 'CORTES DE FRANGO', 'SUINOS']:
                                grupos_especiais.append('VAREJO CALVO')
                            
                            if grupo in grupos_especiais:
                                preco_comparacao = preco * 0.95  # Preço - 5%
                            else:
                                preco_comparacao = preco  # Mantém o preço normal
                            
                            # Lógica de classificação para CB: 2% se >=, 1% se <
                            if preco_comparacao >= preco_oferta_2_cb:
                                comissao = 0.02  # Alterado para 2% como decimal
                            else:
                                comissao = 0.01  # Alterado para 1% como decimal
                                
                            preco_oferta_2 = preco_oferta_2_cb
                            preco_oferta_3 = None
                            tem_2_percent = False
            
                        if is_devolucao:
                            comissao *= -1  # Mantém como decimal negativo
            
                        # Adicionar o preço -5% apenas para os grupos especiais
                        preco_menos_5 = preco * 0.95 if grupo in grupos_especiais else None
                        
                        # Preparar dados para o resultado CB
                        resultado_cb = {
                            **row.to_dict(),
                            'Preço - 5%': preco_menos_5,
                            'Data_Oferta': oferta_cb['DT_REF'],
                            'Comissão_Correta': comissao,
                            'Status': 'Correto' if _comparar_comissoes(row['P. Com'], comissao, 4) else 'Incorreto',
                            'Tipo': 'CB',
                            'Tipo_Oferta': 'Exata' if oferta_cb['DT_REF'] == data else 'Data Proxima'
                        }
                        
                        # Adicionar preços de oferta conforme disponibilidade
                        if tem_3_percent_cb:
                            resultado_cb['Preço_Oferta_3%'] = preco_oferta_3
                            resultado_cb['Preço_Oferta_2%'] = preco_oferta_2
                            # Calcular diferença com base no preço de referência correto
                            if comissao == 0.03:
                                preco_referencia = preco_oferta_3
                            elif comissao == 0.02:
                                preco_referencia = preco_oferta_2
                            else:
                                preco_referencia = None
                                
                            if preco_referencia:
                                resultado_cb['Diferença_Preço'] = f"{(preco_comparacao - preco_referencia)/preco_referencia:.2%}"
                            else:
                                resultado_cb['Diferença_Preço'] = 'N/A'
                        else:
                            resultado_cb['Preço_Oferta'] = preco_oferta_2
                            resultado_cb['Diferença_Preço'] = f"{(preco_comparacao - preco_oferta_2)/preco_oferta_2:.2%}" if preco_oferta_2 != 0 else 'Div/Zero'
                        
                        resultados_ofertas_cb.append(resultado_cb)
                        continue  # Pula para o próximo registro, já encontrou em CB
        
                # SEGUNDO: Se não encontrou em CB, verificar em VOG
                if cod in ofertas_vog_por_codigo.groups:
                    # Buscar oferta VOG específica
                    ofertas_cod_vog = ofertas_vog_por_codigo.get_group(cod)
                    oferta_vog = encontrar_oferta_mais_proxima(ofertas_cod_vog, cod, data)
        
                    if oferta_vog is not None:
                        preco_oferta_3 = float(oferta_vog['3%'])
                        
                        # Verificar se tem coluna '2%' e se tem valor válido
                        tem_2_percent = '2%' in oferta_vog and pd.notna(oferta_vog['2%']) and oferta_vog['2%'] > 0
                        preco_oferta_2 = float(oferta_vog['2%']) if tem_2_percent else None
                        
                        # Aplicar desconto de 5% para grupos especiais
                        grupos_especiais = ['REDE STYLLUS', 'REDE ROD E RAF']
                        if grupo == 'VAREJO CALVO' and grupo_produto in ['MIUDOS BOVINOS', 'CORTES DE FRANGO', 'SUINOS']:
                            grupos_especiais.append('VAREJO CALVO')
                        
                        if grupo in grupos_especiais:
                            preco_comparacao = preco * 0.95  # Preço - 5%
                        else:
                            preco_comparacao = preco  # Mantém o preço normal
                        
                        # NOVA LÓGICA DE CLASSIFICAÇÃO PARA VOG COM 3 FAIXAS
                        if tem_2_percent:
                            # Com três faixas (3%, 2%, 1%)
                            if preco_comparacao >= preco_oferta_3:
                                comissao = 0.03  # Alterado para 3% como decimal
                            elif preco_comparacao >= preco_oferta_2:
                                comissao = 0.02  # Alterado para 2% como decimal
                            else:
                                comissao = 0.01  # Alterado para 1% como decimal
                        else:
                            # Com apenas duas faixas (3%, 1%) - lógica anterior
                            if preco_comparacao >= preco_oferta_3:
                                comissao = 0.03  # Alterado para 3% como decimal
                            else:
                                comissao = 0.01  # Alterado para 1% como decimal
            
                        if is_devolucao:
                            comissao *= -1  # Mantém como decimal negativo
            
                        # Adicionar o preço -5% apenas para os grupos especiais
                        preco_menos_5 = preco * 0.95 if grupo in grupos_especiais else None
                        
                        # Preparar dados para o resultado
                        resultado = {
                            **row.to_dict(),
                            'Preço_Oferta_3%': preco_oferta_3,
                            'Preço - 5%': preco_menos_5,
                            'Data_Oferta': oferta_vog['Data'],
                            'Comissão_Correta': comissao,
                            'Status': 'Correto' if _comparar_comissoes(row['P. Com'], comissao, 4) else 'Incorreto',
                            'Tipo': 'VOG',
                            'Tipo_Oferta': 'Exata' if oferta_vog['Data'] == data else 'Data Proxima'
                        }
                        
                        # Adicionar preço da oferta de 2% se existir
                        if tem_2_percent:
                            resultado['Preço_Oferta_2%'] = preco_oferta_2
                            # Calcular diferença com base no preço de referência correto
                            if comissao == 0.03:
                                preco_referencia = preco_oferta_3
                            elif comissao == 0.02:
                                preco_referencia = preco_oferta_2
                            else:
                                preco_referencia = None
                                
                            if preco_referencia:
                                resultado['Diferença_Preço'] = f"{(preco_comparacao - preco_referencia)/preco_referencia:.2%}"
                            else:
                                resultado['Diferença_Preço'] = 'N/A'
                        else:
                            # Para ofertas sem 2%, usar a lógica anterior
                            resultado['Diferença_Preço'] = f"{(preco_comparacao - preco_oferta_3)/preco_oferta_3:.2%}" if preco_oferta_3 != 0 else 'Div/Zero'
                        
                        resultados_ofertas_vog.append(resultado)
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
        
        # Exibir estatísticas das novas colunas
        if '3%' in df_ofertas_cb.columns:
            cb_com_3_percent = df_ofertas_cb['3%'].notna().sum()
            cb_com_2_percent = df_ofertas_cb['2%'].notna().sum()
            print(f"\nEstatísticas OFF_VOG_CB:")
            print(f"  - Registros com coluna '3%': {cb_com_3_percent}")
            print(f"  - Registros com coluna '2%': {cb_com_2_percent}")
            print(f"  - Ofetas CB com 3 faixas (3%,2%,1%): {cb_com_3_percent}")
            print(f"  - Ofetas CB com 2 faixas (2%,1%): {cb_com_2_percent - cb_com_3_percent}")

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
                
                colunas_ordenadas = ['CF', 'RAZAO', 'GRUPO', 'NF-E', 'DATA', 'VENDEDOR', 'CODPRODUTO',
                                    'GRUPO PRODUTO', 'DESCRICAO', 'Preço_Venda', 'P. Com', 'Comissao_Esperada']
                df_regras_corretas = df_regras_corretas[colunas_ordenadas].rename(
                    columns={'Comissao_Esperada': 'O Com'})
                
                df_regras_corretas.to_excel(writer, sheet_name='O Regras', index=False)

            # 3. Regras Incorretas
            if not df_regras_incorretas.empty:
                df_regras_incorretas = df_regras_incorretas.drop(
                    columns=['Comissao_Kg', 'Status'], errors='ignore')
                
                colunas_ordenadas = ['CF', 'RAZAO', 'GRUPO', 'NF-E', 'DATA', 'VENDEDOR', 'CODPRODUTO',
                                    'GRUPO PRODUTO', 'DESCRICAO', 'Preço_Venda', 'P. Com', 'Comissao_Esperada']
                df_regras_incorretas = df_regras_incorretas[colunas_ordenadas].rename(
                    columns={'Comissao_Esperada': 'O Com'})
                
                df_regras_incorretas.to_excel(writer, sheet_name='X Regras', index=False)

            # 4. Ofertas Corretas
            if not df_resultados_ofertas.empty:
                df_ofertas_corretas = df_resultados_ofertas[df_resultados_ofertas['Status'] == 'Correto']
                if not df_ofertas_corretas.empty:
                    df_ofertas_corretas = df_ofertas_corretas.drop(
                        columns=['Comissao_Kg', 'Comissao_Esperada', 'Diferença_Preço', 'Status', 'Tipo_Oferta'], 
                        errors='ignore')
                    
                    # Ajustar colunas baseado no que existe
                    colunas_base = ['CF', 'RAZAO', 'GRUPO', 'NF-E', 'VENDEDOR', 'CODPRODUTO',
                                   'GRUPO PRODUTO', 'DESCRICAO', 'Preço_Venda', 'Preço - 5%', 'DATA',
                                   'P. Com', 'Data_Oferta', 'Comissão_Correta', 'Tipo']
                    
                    # Adicionar colunas específicas conforme disponíveis
                    colunas_finais = colunas_base.copy()
                    
                    # Verificar quais colunas de preço de oferta existem
                    if 'Preço_Oferta_3%' in df_ofertas_corretas.columns:
                        colunas_finais.insert(12, 'Preço_Oferta_3%')
                    if 'Preço_Oferta_2%' in df_ofertas_corretas.columns:
                        colunas_finais.insert(13 if 'Preço_Oferta_3%' in colunas_finais else 12, 'Preço_Oferta_2%')
                    if 'Preço_Oferta' in df_ofertas_corretas.columns:
                        colunas_finais.insert(12, 'Preço_Oferta')
                    
                    df_ofertas_corretas = df_ofertas_corretas[colunas_finais].rename(
                        columns={'Comissão_Correta': 'O Com'})
                    
                    df_ofertas_corretas.to_excel(writer, sheet_name='O Ofertas', index=False)

            # 5. Ofertas Incorretas
            if not df_resultados_ofertas.empty:
                df_ofertas_incorretas = df_resultados_ofertas[df_resultados_ofertas['Status'] == 'Incorreto']
                if not df_ofertas_incorretas.empty:
                    df_ofertas_incorretas = df_ofertas_incorretas.drop(
                        columns=['Comissao_Kg', 'Comissao_Esperada', 'Diferença_Preço', 'Status', 'Tipo_Oferta'], 
                        errors='ignore')
                    
                    # Ajustar colunas baseado no que existe
                    colunas_base = ['CF', 'RAZAO', 'GRUPO', 'NF-E', 'VENDEDOR', 'CODPRODUTO',
                                   'GRUPO PRODUTO', 'DESCRICAO', 'Preço_Venda', 'Preço - 5%', 'DATA',
                                   'P. Com', 'Data_Oferta', 'Comissão_Correta', 'Tipo']
                    
                    # Adicionar colunas específicas conforme disponíveis
                    colunas_finais = colunas_base.copy()
                    
                    # Verificar quais colunas de preço de oferta existem
                    if 'Preço_Oferta_3%' in df_ofertas_incorretas.columns:
                        colunas_finais.insert(12, 'Preço_Oferta_3%')
                    if 'Preço_Oferta_2%' in df_ofertas_incorretas.columns:
                        colunas_finais.insert(13 if 'Preço_Oferta_3%' in colunas_finais else 12, 'Preço_Oferta_2%')
                    if 'Preço_Oferta' in df_ofertas_incorretas.columns:
                        colunas_finais.insert(12, 'Preço_Oferta')
                    
                    df_ofertas_incorretas = df_ofertas_incorretas[colunas_finais].rename(
                        columns={'Comissão_Correta': 'O Com'})
                    
                    df_ofertas_incorretas.to_excel(writer, sheet_name='X Ofertas', index=False)

            # 6. Sem Oferta
            if not df_sem_oferta_final.empty:
                df_sem_oferta_final = df_sem_oferta_final.drop(
                    columns=['Comissao_Kg', 'Comissao_Esperada'], 
                    errors='ignore')
                
                df_sem_oferta_final.to_excel(writer, sheet_name='Sem Oferta', index=False)

            # 7. Logs de erros
            if not df_logs_erros.empty:
                df_logs_erros.to_excel(writer, sheet_name='Logs Erros', index=False)
                
            # Aplicar formatação de porcentagem com 2 decimais às colunas de comissão
            workbook = writer.book
            
            # Para cada planilha, aplicar formatação às colunas de comissão
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                
                # Encontrar índices das colunas que contêm "Com" (comissão)
                for col_idx, col_name in enumerate(next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True))):
                    col_str = str(col_name)
                    if any(keyword in col_str for keyword in ['P. Com', 'O Com', 'Comissão_Correta', 'Comissao_Esperada']):
                        # Aplicar formatação de porcentagem a partir da linha 2
                        for row in worksheet.iter_rows(min_row=2, min_col=col_idx+1, max_col=col_idx+1):
                            for cell in row:
                                # Formatar como porcentagem com 2 decimais
                                cell.number_format = '0.00%'

        print("\n=== PROCESSAMENTO CONCLUÍDO COM SUCESSO ===")
        
    except Exception as e:
        print(f"\nERRO CRÍTICO DURANTE O PROCESSAMENTO: {str(e)}")
        raise

if __name__ == "__main__":
    processar_planilhas()