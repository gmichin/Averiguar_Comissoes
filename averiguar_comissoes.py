import pandas as pd
from datetime import datetime
import os
import numpy as np
from openpyxl.styles import numbers

def encontrar_oferta_mais_proxima(df_ofertas, codproduto, data_venda):
    """
    Encontra a oferta mais próxima no tempo (anterior ou posterior)
    Prioridade: data exata > mais recente anterior > mais próxima futura
    """
    try:
        cod = int(float(codproduto))
        data = pd.to_datetime(data_venda).date()
        
        # Filtrar ofertas para o código do produto
        ofertas_cod = df_ofertas[df_ofertas['COD'] == cod].copy()
        
        if ofertas_cod.empty:
            return None
            
        # Converter datas
        ofertas_cod['DT_REF_OFF'] = pd.to_datetime(ofertas_cod['DT_REF_OFF']).dt.date
        
        # DEBUG: Mostrar todas as ofertas disponíveis (opcional)
        # print(f"DEBUG - Produto {cod}, Data venda: {data}")
        # for idx, row in ofertas_cod.iterrows():
        #     print(f"  Data: {row['DT_REF_OFF']}, 3%: {row.get('3%', 'N/A')}, 2%: {row.get('2%', 'N/A')}")
        
        # 1. Buscar oferta na data exata
        oferta_exata = ofertas_cod[ofertas_cod['DT_REF_OFF'] == data]
        if not oferta_exata.empty:
            return oferta_exata.iloc[0]
        
        # 2. Buscar a oferta mais recente ANTERIOR
        ofertas_anteriores = ofertas_cod[ofertas_cod['DT_REF_OFF'] < data]
        if not ofertas_anteriores.empty:
            # Ordenar por data DESCENDENTE (mais recente primeiro)
            ofertas_anteriores = ofertas_anteriores.sort_values('DT_REF_OFF', ascending=False)
            return ofertas_anteriores.iloc[0]
        
        # 3. Se não houver anteriores, buscar a PRIMEIRA oferta POSTERIOR
        ofertas_posteriores = ofertas_cod[ofertas_cod['DT_REF_OFF'] > data]
        if not ofertas_posteriores.empty:
            # Ordenar por data ASCENDENTE (mais próxima no futuro)
            ofertas_posteriores = ofertas_posteriores.sort_values('DT_REF_OFF', ascending=True)
            return ofertas_posteriores.iloc[0]
        
        return None
        
    except Exception as e:
        print(f"Erro ao buscar oferta para {codproduto}: {str(e)}")
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
            'razao_codigos': {'SUPERMERCADO FEDERZONI LTDA': [812],
                              'HX3698 HOUSE CENTER LTDA': [812]}
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
                    'REDE REIMBERG', 'REDE SEMAR', 'REDE TRIMAIS', 'REDE VOVO ZUZU',
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
            'REDE MERCADAO': {
                0.005: {
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

    if nfe == '122022' and codproduto == 1451:
        return _ajustar_para_devolucao(0.03, is_devolucao) 
    if nfe == '121905' and codproduto == 1477:
        return _ajustar_para_devolucao(0.01, is_devolucao) 
    if nfe == '122520' and codproduto == 2045:
        return _ajustar_para_devolucao(0.01, is_devolucao) 
    if nfe == '122643' and codproduto == 1443:
        return _ajustar_para_devolucao(0.01, is_devolucao) 
    if nfe == '122886' and codproduto == 2010:
        return _ajustar_para_devolucao(0.01, is_devolucao) 
    if nfe == '122221' and codproduto == 2026:
        return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122225' and codproduto == 2026:
        return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122265' and codproduto == 2026:
        return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122401' and codproduto == 2026:
        return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122559' and codproduto == 2026:
        return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '112636' and codproduto == 2050:
        return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122637' and codproduto == 2050:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122657' and codproduto == 2050:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122658' and codproduto == 2050:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122727' and codproduto == 2050:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122792' and codproduto == 2050:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122958' and codproduto == 2050:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122991' and codproduto == 2050:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '123007' and codproduto == 2050:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '123047' and codproduto == 2050:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '123061' and codproduto == 2050:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '123273' and codproduto == 2050:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '123282' and codproduto == 2050:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122636' and codproduto == 2050:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122705' and codproduto == 1505:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122706' and codproduto == 1505:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122720' and codproduto == 1505:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122722' and codproduto == 1505:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122987' and codproduto == 1505:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '123004' and codproduto == 1505:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '123052' and codproduto == 1505:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '123062' and codproduto == 1505:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '123064' and codproduto == 1505:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '123282' and codproduto == 1505:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122638' and codproduto == 2052:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '122685' and codproduto == 2052:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '123061' and codproduto == 1513:
            return _ajustar_para_devolucao(0.005, is_devolucao) 
    if nfe == '124082' and codproduto == 2045:
            return _ajustar_para_devolucao(0.01, is_devolucao) 
    if nfe == '124033' and codproduto == 1547:
            return _ajustar_para_devolucao(0.01, is_devolucao) 
    if nfe == '124156' and codproduto == 2045:
            return _ajustar_para_devolucao(0.01, is_devolucao) 


    if codproduto == 1807 or codproduto == 947 or codproduto == 1914 or codproduto == 2000:
        return _ajustar_para_devolucao(0.01, is_devolucao)
    
    if vendedor == "PROPRIO":
        return _ajustar_para_devolucao(0.00, is_devolucao)
    
    if grupo == 'REDE RICOY':
        return _ajustar_para_devolucao(0.00, is_devolucao)
    
    if grupo == 'REDE ROLDAO':
        grupos_2_percent = [
            'CONGELADOS', 'CORTES BOVINOS', 'CORTES DE FRANGO', 'EMBUTIDOS', 
            'EMBUTIDOS AURORA', 'EMBUTIDOS NOBRE', 'EMBUTIDOS PERDIGÃO', 
            'EMBUTIDOS SADIA', 'EMBUTIDOS SEARA', 'EMPANADOS', 
            'KITS FEIJOADA', 'MIUDOS BOVINOS', 'SUINOS', 'TEMPERADOS'
        ]
        
        if grupo_produto in grupos_2_percent:
            return _ajustar_para_devolucao(0.02, is_devolucao)
        else:
            return _ajustar_para_devolucao(0.00, is_devolucao)
        
    if grupo == 'REDE MERCADAO':
        grupos_05_percent = [
            'CONGELADOS', 'CORTES BOVINOS', 'CORTES DE FRANGO', 'EMBUTIDOS', 
            'EMBUTIDOS AURORA', 'EMBUTIDOS NOBRE', 'EMBUTIDOS PERDIGÃO', 
            'EMBUTIDOS SADIA', 'EMBUTIDOS SEARA', 'EMPANADOS', 
            'KITS FEIJOADA', 'MIUDOS BOVINOS', 'SUINOS', 'TEMPERADOS'
        ]
        
        if grupo_produto in grupos_05_percent:
            return _ajustar_para_devolucao(0.005, is_devolucao)
        else:
            return _ajustar_para_devolucao(0.00, is_devolucao)

    if grupo == 'VAREJO CALVO':
        if grupo_produto in ['MIUDOS BOVINOS', 'CORTES DE FRANGO', 'SUINOS']:
            return None  # Processar por ofertas
        
        return _ajustar_para_devolucao(0.03, is_devolucao)
    
    if 'REDE CENCOSUD' in grupo:
        if 'SALAME UAI' in grupo_produto:
            return _ajustar_para_devolucao(0.01, is_devolucao)
        return _ajustar_para_devolucao(0.03, is_devolucao)
    
    if grupo == 'REDE ROSSI':
        if codproduto in [937, 1698, 1701, 1587, 1700, 1586, 1699, 943, 1735, 1624, 1134]:
            return _ajustar_para_devolucao(0.03, is_devolucao)
        
        if codproduto == 1139:
            return _ajustar_para_devolucao(0.00, is_devolucao)
        
        if grupo_produto in ['EMBUTIDOS', 'EMBUTIDOS NOBRE', 'EMBUTIDOS SADIA', 
                           'EMBUTIDOS PERDIGAO', 'EMBUTIDOS AURORA', 'EMBUTIDOS SEARA', 
                           'SALAME UAI']:
            return _ajustar_para_devolucao(0.00, is_devolucao)
        
        if grupo_produto in ["MIUDOS BOVINOS", "SUINOS", "SALGADOS SUINOS A GRANEL", 
                             "SALGADOS SUINOS EMBALADOS", "CORTES DE FRANGO"]:
            return _ajustar_para_devolucao(0.02, is_devolucao)
        
        if codproduto == 700:
            return _ajustar_para_devolucao(0.02, is_devolucao)
        
        if codproduto in [1265, 1266, 812, 1115, 798, 1211]:
            return _ajustar_para_devolucao(0.01, is_devolucao)
    
    if grupo == 'REDE PLUS':
        if grupo_produto in ['TEMPERADOS']:
            return _ajustar_para_devolucao(0.03, is_devolucao)
        
        if codproduto == 812:
            return _ajustar_para_devolucao(0.03, is_devolucao)
    
    if grupo in regras['grupos_especificos']:
        regras_grupo = regras['grupos_especificos'][grupo]
        
        for porcentagem in [0.00, 0.02, 0.01, 0.03]:
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
    try:
        if pd.isna(valor):
            return None
        
        if isinstance(valor, str):
            valor = valor.replace(',', '.').replace('%', '').strip()
        
        valor_float = float(valor)
        
        if valor_float > 1:
            return valor_float / 100.0
        else:
            return valor_float
            
    except Exception as e:
        print(f"Erro ao converter valor '{valor}': {str(e)}")
        return None

def _comparar_comissoes(comissao_atual, comissao_esperada, decimal_places=4):
    try:
        atual = _converter_para_decimal_percentual(comissao_atual)
        esperada = _converter_para_decimal_percentual(comissao_esperada)
        
        if atual is None or esperada is None:
            return False
        
        atual_rounded = round(atual, decimal_places)
        esperada_rounded = round(esperada, decimal_places)
        
        return atual_rounded == esperada_rounded
    except (ValueError, TypeError) as e:
        print(f"Erro na comparação: {str(e)}, atual={comissao_atual}, esperada={comissao_esperada}")
        return False

def _converter_valor_oferta(valor):
    try:
        if pd.isna(valor):
            return np.nan
        
        if isinstance(valor, str):
            valor_str = valor.strip().upper()
            
            if valor_str in ['-', '#N/D', 'N/A', 'NAN', 'NULL', '', 'N/D']:
                return np.nan
            
            valor_str = valor_str.replace(',', '.')
            
            return float(valor_str)
        
        return float(valor)
        
    except Exception as e:
        print(f"Erro ao converter valor de oferta '{valor}': {str(e)}")
        return np.nan

def classificar_comissao_por_oferta(preco, preco_oferta_3, preco_oferta_2, preco_oferta_1, grupo, grupo_produto, is_devolucao):
    """
    Classifica a comissão baseada nos preços de oferta
    Lógica:
    - Se preço >= oferta_3%: 3%
    - Se preço >= oferta_2%: 2% 
    - Se preço >= oferta_1%: 1%
    - Caso contrário: 1% (fallback)
    
    Aplica -5% para grupos especiais antes da comparação
    """
    
    # Aplicar desconto de 5% para grupos especiais
    grupos_especiais = ['REDE STYLLUS', 'REDE ROD E RAF']
    if grupo == 'VAREJO CALVO' and grupo_produto in ['MIUDOS BOVINOS', 'CORTES DE FRANGO', 'SUINOS']:
        grupos_especiais.append('VAREJO CALVO')
    
    if grupo in grupos_especiais:
        # Aplicar desconto de 5% e ARREDONDAR para 2 casas decimais
        preco_comparacao = round(preco * 0.95, 2)
    else:
        preco_comparacao = preco
    
    # Verificar se os preços de oferta são válidos
    preco_oferta_3_valido = preco_oferta_3 is not None and not np.isnan(preco_oferta_3) and preco_oferta_3 > 0
    preco_oferta_2_valido = preco_oferta_2 is not None and not np.isnan(preco_oferta_2) and preco_oferta_2 > 0
    preco_oferta_1_valido = preco_oferta_1 is not None and not np.isnan(preco_oferta_1) and preco_oferta_1 > 0
    
    # Lógica de classificação
    if preco_oferta_3_valido and preco_comparacao >= preco_oferta_3:
        comissao = 0.03
    elif preco_oferta_2_valido and preco_comparacao >= preco_oferta_2:
        comissao = 0.02
    elif preco_oferta_1_valido:
        comissao = 0.01
    else:
        comissao = 0.01  # Fallback padrão
    
    # Ajustar para devolução
    if is_devolucao:
        comissao *= -1
    
    return comissao

def processar_planilhas():
    caminho_origem = r"C:\Users\win11\Downloads\260222_MRG - wapp.xlsx"
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
        
        # Converter preços
        df_base['P. Com'] = df_base['P. Com'].apply(
            lambda x: _converter_valor_oferta(x) if pd.notna(x) else np.nan
        )
        df_base['Preço_Venda'] = df_base['Preço_Venda'].apply(
            lambda x: _converter_valor_oferta(x) if pd.notna(x) else np.nan
        )
        
        # 2. Ler apenas a aba OFF_VOG
        print("\n--- Lendo aba OFF_VOG ---")
        df_ofertas_vog = pd.read_excel(caminho_origem, sheet_name='OFF_VOG')
        
        # Verificar quais colunas existem
        colunas_vog_disponiveis = df_ofertas_vog.columns.tolist()
        print(f"Colunas disponíveis em OFF_VOG: {colunas_vog_disponiveis}")
        
        # Construir lista de colunas baseada no que existe
        colunas_ofertas_vog = []
        if 'COD' in colunas_vog_disponiveis:
            colunas_ofertas_vog.append('COD')
        if 'ITENS' in colunas_vog_disponiveis:
            colunas_ofertas_vog.append('ITENS')
        if '3%' in colunas_vog_disponiveis:
            colunas_ofertas_vog.append('3%')
        if '2%' in colunas_vog_disponiveis:
            colunas_ofertas_vog.append('2%')
        if '1%' in colunas_vog_disponiveis:
            colunas_ofertas_vog.append('1%')
        
        # Verificar se existe DT_REF_OFF ou Data
        if 'DT_REF_OFF' in colunas_vog_disponiveis:
            colunas_ofertas_vog.append('DT_REF_OFF')
        elif 'Data' in colunas_vog_disponiveis:
            colunas_ofertas_vog.append('Data')
            df_ofertas_vog = df_ofertas_vog.rename(columns={'Data': 'DT_REF_OFF'})
        else:
            raise ValueError("Não encontrada coluna de data em OFF_VOG")
        
        # Se existirem outras colunas úteis
        for col in ['Coluna1', 'PK_OFF']:
            if col in colunas_vog_disponiveis:
                colunas_ofertas_vog.append(col)
        
        # Filtrar apenas as colunas necessárias
        df_ofertas_vog = df_ofertas_vog[colunas_ofertas_vog].dropna(subset=['COD', 'DT_REF_OFF'])
        
        # Converter tipos
        df_ofertas_vog['DT_REF_OFF'] = pd.to_datetime(df_ofertas_vog['DT_REF_OFF']).dt.date
        df_ofertas_vog['COD'] = pd.to_numeric(df_ofertas_vog['COD'], errors='coerce').fillna(0).astype('int64')
        
        # Converter preços das ofertas
        if '3%' in colunas_ofertas_vog:
            df_ofertas_vog['3%'] = df_ofertas_vog['3%'].apply(
                lambda x: _converter_valor_oferta(x) if pd.notna(x) else np.nan
            )
        if '2%' in colunas_ofertas_vog:
            df_ofertas_vog['2%'] = df_ofertas_vog['2%'].apply(
                lambda x: _converter_valor_oferta(x) if pd.notna(x) else np.nan
            )
        if '1%' in colunas_ofertas_vog:
            df_ofertas_vog['1%'] = df_ofertas_vog['1%'].apply(
                lambda x: _converter_valor_oferta(x) if pd.notna(x) else np.nan
            )
        
        print(f"Total de ofertas cadastradas: {len(df_ofertas_vog)}")
        print(f"Colunas em OFF_VOG: {colunas_ofertas_vog}")
        
        # 3. Aplicar regras de comissão por kg
        regras_comissao_kg = criar_regras_comissao_kg()
        df_base['Comissao_Kg'] = df_base.apply(
            lambda row: pertence_comissao_kg(row, regras_comissao_kg), axis=1)
        
        df_comissao_kg = df_base[df_base['Comissao_Kg'] == True].copy()
        df_sem_kg = df_base[df_base['Comissao_Kg'] == False].copy()
        
        print(f"- Itens para comissão por kg: {len(df_comissao_kg)}")
        
        # 4. Aplicar regras fixas
        regras_comissao_fixa = criar_regras_comissao_fixa()
        df_sem_kg['Comissao_Esperada'] = df_sem_kg.apply(
            lambda row: aplicar_regras_comissao_fixa(row, regras_comissao_fixa), axis=1)
        
        # Separar os que tem regra aplicada
        mask_regras = df_sem_kg['Comissao_Esperada'].notna()
        df_regras = df_sem_kg[mask_regras].copy()
        df_sem_regra = df_sem_kg[~mask_regras].copy()
        
        # Verificar se a comissão aplicada está correta
        df_regras['Status'] = df_regras.apply(
            lambda row: 'Correto' if _comparar_comissoes(row['P. Com'], row['Comissao_Esperada'], 4) else 'Incorreto', 
            axis=1)
        
        df_regras_corretas = df_regras[df_regras['Status'] == 'Correto']
        df_regras_incorretas = df_regras[df_regras['Status'] == 'Incorreto']
        
        print(f"- Registros com regras fixas aplicadas: {len(df_regras)}")
        print(f"  → Corretos: {len(df_regras_corretas)}")
        print(f"  → Incorretos: {len(df_regras_incorretas)}")
        
        # 5. Verificação das ofertas VOG
        resultados_ofertas = []
        registros_sem_oferta = []
        logs_erros = []
        
        # Otimização: Criar dicionário de ofertas por código para acesso rápido
        ofertas_por_codigo = df_ofertas_vog.groupby('COD')
        codigos_com_oferta = set(ofertas_por_codigo.groups.keys())
        
        print(f"\n--- Processando ofertas VOG ---")
        print(f"Códigos com oferta disponível: {len(codigos_com_oferta)}")
        
        for idx, row in df_sem_regra.iterrows():
            try:
                cod = int(float(row['CODPRODUTO']))
                data = pd.to_datetime(row['DATA']).date()
                preco = float(row['Preço_Venda'])
                is_devolucao = str(row['CF']).startswith('DEV')
                grupo = str(row['GRUPO']).strip().upper()
                grupo_produto = str(row['GRUPO PRODUTO']).strip().upper()
                
                # Verificar se existe oferta para este código
                if cod in codigos_com_oferta:
                    # Buscar oferta específica
                    ofertas_cod = ofertas_por_codigo.get_group(cod)
                    oferta = encontrar_oferta_mais_proxima(ofertas_cod, cod, data)
                    
                    if oferta is not None:
                        # Extrair preços da oferta
                        preco_oferta_3 = oferta.get('3%')
                        preco_oferta_2 = oferta.get('2%')
                        preco_oferta_1 = oferta.get('1%')
                        
                        # Converter para float se necessário
                        preco_oferta_3 = _converter_valor_oferta(preco_oferta_3) if preco_oferta_3 is not None else None
                        preco_oferta_2 = _converter_valor_oferta(preco_oferta_2) if preco_oferta_2 is not None else None
                        preco_oferta_1 = _converter_valor_oferta(preco_oferta_1) if preco_oferta_1 is not None else None
                        
                        # Classificar usando a função
                        comissao = classificar_comissao_por_oferta(
                            preco, 
                            preco_oferta_3, 
                            preco_oferta_2, 
                            preco_oferta_1,
                            grupo,
                            grupo_produto,
                            is_devolucao
                        )
                        
                        # Determinar tipo de oferta
                        if oferta['DT_REF_OFF'] == data:
                            tipo_oferta = 'Exata'
                        elif oferta['DT_REF_OFF'] < data:
                            tipo_oferta = 'Data Proxima Ant'
                        else:
                            tipo_oferta = 'Data Proxima Pos'
                        
                        # Preparar dados para o resultado
                        resultado = {
                            **row.to_dict(),
                            'Preço - 5%': preco * 0.95 if grupo in ['REDE STYLLUS', 'REDE ROD E RAF'] else None,
                            'Data_Oferta': oferta['DT_REF_OFF'],
                            'Comissão_Correta': comissao,
                            'Status': 'Correto' if _comparar_comissoes(row['P. Com'], comissao, 4) else 'Incorreto',
                            'Tipo': 'VOG',
                            'Tipo_Oferta': tipo_oferta
                        }
                        
                        # Adicionar preços de oferta conforme disponibilidade
                        if preco_oferta_3 is not None and not np.isnan(preco_oferta_3):
                            resultado['Preço_Oferta_3%'] = preco_oferta_3
                        if preco_oferta_2 is not None and not np.isnan(preco_oferta_2):
                            resultado['Preço_Oferta_2%'] = preco_oferta_2
                        if preco_oferta_1 is not None and not np.isnan(preco_oferta_1):
                            resultado['Preço_Oferta_1%'] = preco_oferta_1
                            
                        resultados_ofertas.append(resultado)
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
        
        # Criar DataFrames de resultados
        df_resultados_ofertas = pd.DataFrame(resultados_ofertas) if resultados_ofertas else pd.DataFrame()
        df_sem_oferta_final = pd.DataFrame(registros_sem_oferta) if registros_sem_oferta else pd.DataFrame()
        df_logs_erros = pd.DataFrame(logs_erros) if logs_erros else pd.DataFrame()

        print(f"- Itens com oferta VOG encontrada: {len(df_resultados_ofertas)}")
        print(f"- Itens sem oferta encontrada: {len(df_sem_oferta_final)}")
        print(f"- Erros durante o processamento: {len(df_logs_erros)}")
        
        # Separar ofertas corretas e incorretas
        if not df_resultados_ofertas.empty:
            df_ofertas_corretas = df_resultados_ofertas[df_resultados_ofertas['Status'] == 'Correto']
            df_ofertas_incorretas = df_resultados_ofertas[df_resultados_ofertas['Status'] == 'Incorreto']
            print(f"  → Ofertas Corretas: {len(df_ofertas_corretas)}")
            print(f"  → Ofertas Incorretas: {len(df_ofertas_incorretas)}")
        
        # 6. Exportar para Excel
        print(f"\nSALVANDO RESULTADOS EM: {caminho_downloads}")
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
                        columns=['Comissao_Kg', 'Comissao_Esperada', 'Status', 'Tipo_Oferta'], 
                        errors='ignore')
                    
                    # Ajustar colunas
                    colunas_base = ['CF', 'RAZAO', 'GRUPO', 'NF-E', 'VENDEDOR', 'CODPRODUTO',
                                   'GRUPO PRODUTO', 'DESCRICAO', 'Preço_Venda', 'Preço - 5%', 'DATA',
                                   'P. Com', 'Data_Oferta', 'Comissão_Correta', 'Tipo']
                    
                    colunas_finais = colunas_base.copy()
                    
                    # Adicionar colunas de preço conforme disponíveis
                    if 'Preço_Oferta_3%' in df_ofertas_corretas.columns:
                        colunas_finais.insert(12, 'Preço_Oferta_3%')
                    if 'Preço_Oferta_2%' in df_ofertas_corretas.columns:
                        colunas_finais.insert(13 if 'Preço_Oferta_3%' in colunas_finais else 12, 'Preço_Oferta_2%')
                    if 'Preço_Oferta_1%' in df_ofertas_corretas.columns:
                        colunas_finais.insert(14 if 'Preço_Oferta_2%' in colunas_finais else 12, 'Preço_Oferta_1%')
                    
                    df_ofertas_corretas = df_ofertas_corretas[colunas_finais].rename(
                        columns={'Comissão_Correta': 'O Com'})
                    
                    df_ofertas_corretas.to_excel(writer, sheet_name='O Ofertas', index=False)

            # 5. Ofertas Incorretas
            if not df_resultados_ofertas.empty:
                df_ofertas_incorretas = df_resultados_ofertas[df_resultados_ofertas['Status'] == 'Incorreto']
                if not df_ofertas_incorretas.empty:
                    df_ofertas_incorretas = df_ofertas_incorretas.drop(
                        columns=['Comissao_Kg', 'Comissao_Esperada', 'Status', 'Tipo_Oferta'], 
                        errors='ignore')
                    
                    # Ajustar colunas
                    colunas_base = ['CF', 'RAZAO', 'GRUPO', 'NF-E', 'VENDEDOR', 'CODPRODUTO',
                                   'GRUPO PRODUTO', 'DESCRICAO', 'Preço_Venda', 'Preço - 5%', 'DATA',
                                   'P. Com', 'Data_Oferta', 'Comissão_Correta', 'Tipo']
                    
                    colunas_finais = colunas_base.copy()
                    
                    # Adicionar colunas de preço conforme disponíveis
                    if 'Preço_Oferta_3%' in df_ofertas_incorretas.columns:
                        colunas_finais.insert(12, 'Preço_Oferta_3%')
                    if 'Preço_Oferta_2%' in df_ofertas_incorretas.columns:
                        colunas_finais.insert(13 if 'Preço_Oferta_3%' in colunas_finais else 12, 'Preço_Oferta_2%')
                    if 'Preço_Oferta_1%' in df_ofertas_incorretas.columns:
                        colunas_finais.insert(14 if 'Preço_Oferta_2%' in colunas_finais else 12, 'Preço_Oferta_1%')
                    
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
            
            # Aplicar formatação de porcentagem
            workbook = writer.book
            
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                
                for col_idx, col_name in enumerate(next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True))):
                    col_str = str(col_name)
                    if any(keyword in col_str for keyword in ['P. Com', 'O Com', 'Comissão_Correta', 'Comissao_Esperada']):
                        for row in worksheet.iter_rows(min_row=2, min_col=col_idx+1, max_col=col_idx+1):
                            for cell in row:
                                cell.number_format = '0.00%'

        print("\n=== PROCESSAMENTO CONCLUÍDO COM SUCESSO ===")
        print(f"Arquivo salvo em: {caminho_downloads}")
        
    except Exception as e:
        print(f"\nERRO CRÍTICO DURANTE O PROCESSAMENTO: {str(e)}")
        import traceback
        traceback.print_exc()
        raise

if __name__ == "__main__":
    processar_planilhas()