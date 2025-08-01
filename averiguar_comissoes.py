import pandas as pd
from datetime import datetime
import os
import numpy as np

def encontrar_oferta_mais_proxima(df_ofertas, codproduto, data_venda):
    """
    Versão otimizada para encontrar ofertas de forma confiável
    - Primeiro verifica oferta na data exata
    - Se não encontrar, busca a oferta mais recente anterior à data de venda
    - Retorna None se não encontrar oferta válida
    """
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
    """Cria as regras para classificação de comissão por kg"""
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

def criar_regras_comissao_fixa():
    """Cria as regras para comissões fixas por grupo/razão social"""
    regras = {
        '0%': {
            'grupos': [
                'AKKI ATACADISTA', 'ANDORINHA', 'BERGAMINI', 'DA PRACA', 'DOVALE',
                'MERCADAO ATACADISTA', 'REIMBERG', 'SEMAR', 'TRIMAIS', 'VOVO ZUZU'
            ],
            'razoes': [
                'COMERCIO DE CARNES E ROTISSERIE DUTRA LT',
                'SUPERMERCADO HIGAS ITAQUERA LTDA'
            ]
        },
        '3%': {
            'grupos': ['CALVO', 'CENCOSUD', 'CHAMA', 'ESTRELA AZUL', 'TENDA']
        },
        '1%': {
            'grupos': ['ROLDAO']
        },
        'STYLLUS': {
            '0%': {
                'grupos_produto': ['TORRESMO', 'SALAME UAI', 'EMPANADOS']
            },
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
    """Aplica as regras de comissão fixa e retorna a comissão esperada ou None"""
    grupo = str(row['GRUPO']).strip().upper()
    razao = str(row['RAZAO']).strip().upper()
    codproduto = row['CODPRODUTO']
    grupo_produto = str(row['GRUPO PRODUTO']).strip().upper()
    is_devolucao = str(row['CF']).startswith('DEV')
    
    # Verifica primeiro as regras específicas do STYLLUS
    if grupo == 'STYLLUS' and 'STYLLUS' in regras:
        if '0%' in regras['STYLLUS']:
            if 'grupos_produto' in regras['STYLLUS']['0%']:
                if grupo_produto in regras['STYLLUS']['0%']['grupos_produto']:
                    return 0 if not is_devolucao else -0
    
    # Verifica regras gerais de 0%
    if '0%' in regras:
        if grupo in regras['0%'].get('grupos', []):
            return 0 if not is_devolucao else -0
        if razao in regras['0%'].get('razoes', []):
            return 0 if not is_devolucao else -0
    
    # Verifica regras de 3%
    if '3%' in regras:
        if grupo in regras['3%'].get('grupos', []):
            return 3 if not is_devolucao else -3
    
    # Verifica regras de 1%
    if '1%' in regras:
        if grupo in regras['1%'].get('grupos', []):
            return 1 if not is_devolucao else -1
    
    # Verifica regras específicas para ROSSI
    if grupo == 'ROSSI':
        if codproduto in regras['ROSSI'].get('3%', []):
            return 3 if not is_devolucao else -3
        if codproduto in regras['ROSSI'].get('1%', []):
            return 1 if not is_devolucao else -1
        if codproduto in regras['ROSSI'].get('2%', {}).get('codigos', []):
            return 2 if not is_devolucao else -2
        if grupo_produto in regras['ROSSI'].get('2%', {}).get('grupos_produto', []):
            return 2 if not is_devolucao else -2
        if grupo_produto in regras['ROSSI'].get('0%', {}).get('grupos_produto', []):
            return 0 if not is_devolucao else -0
    
    return None

def processar_planilhas():
    caminho_origem = r"C:\Users\win11\OneDrive\Documentos\Margens de fechamento\Margem_250731 - wapp.xlsx"
    caminho_downloads = os.path.join(os.path.expanduser('~'), 'Downloads', 'Averiguar_Comissoes (MARGEM).xlsx')
    
    try:
        print("=== INÍCIO DO PROCESSAMENTO ===")
        print("1. LENDO DADOS DE ENTRADA...")
        
        # 1. Ler os dados
        print("- Lendo aba Base (3,5%)...")
        df_base = pd.read_excel(caminho_origem, sheet_name='Base (3,5%)', header=8)
        print(f"  Total de registros na base: {len(df_base)}")
        
        colunas_base = ['CF', 'RAZAO', 'GRUPO', 'NF-E', 'DATA', 'VENDEDOR', 'CODPRODUTO',
                       'GRUPO PRODUTO', 'DESCRICAO', 'P. Com', 'Preço Venda ']
        df_base = df_base[colunas_base].rename(columns={'Preço Venda ': 'Preço_Venda'})
        
        # Converter e formatar dados
        df_base['DATA'] = pd.to_datetime(df_base['DATA']).dt.date
        df_base['CODPRODUTO'] = pd.to_numeric(df_base['CODPRODUTO'], errors='coerce').fillna(0).astype('int64')
        df_base['P. Com'] = (pd.to_numeric(df_base['P. Com'], errors='coerce') * 100).round().astype('Int64')
        df_base['Preço_Venda'] = pd.to_numeric(df_base['Preço_Venda'], errors='coerce')
        
        print("- Lendo aba OFERTAS_VOG...")
        df_ofertas = pd.read_excel(caminho_origem, sheet_name='OFERTAS_VOG')
        df_ofertas = df_ofertas[['COD', 'ITENS', '3%', 'Data']].dropna(subset=['COD', 'Data'])
        
        df_ofertas['Data'] = pd.to_datetime(df_ofertas['Data']).dt.date
        df_ofertas['COD'] = pd.to_numeric(df_ofertas['COD'], errors='coerce').fillna(0).astype('int64')
        df_ofertas['3%'] = pd.to_numeric(df_ofertas['3%'], errors='coerce')
        print(f"  Total de ofertas cadastradas: {len(df_ofertas)}")
        print(f"  Produtos distintos com oferta: {df_ofertas['COD'].nunique()}")
        print(f"  Datas distintas com oferta: {df_ofertas['Data'].nunique()}")
        
        # 2. Aplicar regras de comissão por kg
        print("\n2. APLICANDO REGRAS DE COMISSÃO POR KG...")
        regras_comissao_kg = criar_regras_comissao_kg()
        df_base['Comissao_Kg'] = df_base.apply(
            lambda row: pertence_comissao_kg(row, regras_comissao_kg), axis=1)

        df_comissao_kg = df_base[df_base['Comissao_Kg'] == True].copy()
        df_sem_kg = df_base[df_base['Comissao_Kg'] == False].copy()

        print(f"- Itens para comissão por kg: {len(df_comissao_kg)}")
        print(f"- Itens para próxima etapa: {len(df_sem_kg)}")

        # 3. Aplicar regras fixas
        print("\n3. APLICANDO REGRAS FIXAS DE COMISSÃO...")
        regras_comissao_fixa = criar_regras_comissao_fixa()
        df_sem_kg['Comissao_Esperada'] = df_sem_kg.apply(
            lambda row: aplicar_regras_comissao_fixa(row, regras_comissao_fixa), axis=1)
        
        mask_regras = df_sem_kg['Comissao_Esperada'].notna()
        df_regras = df_sem_kg[mask_regras].copy()
        df_sem_regra = df_sem_kg[~mask_regras].copy()
        
        df_regras['Status'] = df_regras.apply(
            lambda row: 'Correto' if row['P. Com'] == row['Comissao_Esperada'] else 'Incorreto', axis=1)
        
        df_regras_corretas = df_regras[df_regras['Status'] == 'Correto']
        df_regras_incorretas = df_regras[df_regras['Status'] == 'Incorreto']
        
        print(f"- Registros com regras fixas aplicadas: {len(df_regras)}")
        print(f"  → Corretos: {len(df_regras_corretas)}")
        print(f"  → Incorretos: {len(df_regras_incorretas)}")
        
        # 4. Verificação detalhada das ofertas
        print("\n4. VERIFICANDO OFERTAS PARA ITENS SEM REGRA...")
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
                    diferenca_percentual = abs(preco - preco_oferta) / preco_oferta

                    # Tolerância de 5% no preço para considerar como oferta
                    if diferenca_percentual <= 0.0001:
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
                        'Diferença_Preço': f"{diferenca_percentual:.2%}"
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

        print("\n5. RESUMO DOS RESULTADOS:")
        print(f"- Itens com comissão por kg: {len(df_comissao_kg)}")
        print(f"- Itens com regras fixas corretas: {len(df_regras_corretas)}")
        print(f"- Itens com regras fixas incorretas: {len(df_regras_incorretas)}")
        print(f"- Itens com ofertas encontradas: {len(df_resultados_ofertas)}")
        print(f"- Itens sem oferta encontrada: {len(df_sem_oferta_final)}")
        print(f"- Erros durante o processamento: {len(df_logs_erros)}")

        # Verificação crítica dos resultados
        if not df_resultados_ofertas.empty:
            print("\n6. DETALHES DAS OFERTAS ENCONTRADAS:")
            print(f"- Ofertas na data exata: {len(df_resultados_ofertas[df_resultados_ofertas['Tipo'] == 'Exata'])}")
            print(f"- Ofertas em data próxima: {len(df_resultados_ofertas[df_resultados_ofertas['Tipo'] == 'Data Proxima'])}")
            print(f"- Ofertas corretas: {len(df_resultados_ofertas[df_resultados_ofertas['Status'] == 'Correto'])}")
            print(f"- Ofertas incorretas: {len(df_resultados_ofertas[df_resultados_ofertas['Status'] == 'Incorreto'])}")

        # Verificar registros problemáticos
        if not df_sem_oferta_final.empty:
            print("\n7. VERIFICAÇÃO DE REGISTROS SEM OFERTA:")
            # Encontrar códigos que estão em "sem oferta" mas têm ofertas no sistema
            codigos_sem_oferta = set(df_sem_oferta_final['CODPRODUTO'].unique())
            codigos_com_oferta = set(df_ofertas['COD'].unique())
            codigos_problematicos = codigos_sem_oferta.intersection(codigos_com_oferta)
            
            if codigos_problematicos:
                print(f"- AVISO: {len(codigos_problematicos)} códigos foram para 'sem oferta' mas têm ofertas no sistema")
                print("  Exemplos de códigos problemáticos:", list(codigos_problematicos)[:5])
                
                # Mostrar detalhes para análise
                df_problematicos = df_sem_oferta_final[df_sem_oferta_final['CODPRODUTO'].isin(codigos_problematicos)]
                print("\n  Exemplos de registros problemáticos:")
                for _, row in df_problematicos.head(3).iterrows():
                    cod = row['CODPRODUTO']
                    data = row['DATA']
                    print(f"\n  Código: {cod}, Data: {data}")
                    print("  Ofertas disponíveis para este código:")
                    print(df_ofertas[df_ofertas['COD'] == cod][['COD', 'Data', '3%']].sort_values('Data'))

        # Lista de colunas a serem removidas
        colunas_remover = ['Comissao_Kg', 'Comissao_Esperada', 'Diferença_Preço']
        
        # 8. Exportar para Excel
        print(f"\n8. SALVANDO RESULTADOS EM: {caminho_downloads}")
        with pd.ExcelWriter(caminho_downloads, engine='openpyxl') as writer:
            # Comissão por Kg - Remover colunas indesejadas
            if not df_comissao_kg.empty:
                df_comissao_kg_final = df_comissao_kg.drop(columns=[col for col in colunas_remover if col in df_comissao_kg.columns])
                df_comissao_kg_final.to_excel(writer, sheet_name='Comissão por Kg', index=False)
        
            # Regras Corretas - Remover colunas indesejadas
            if not df_regras_corretas.empty:
                df_regras_corretas_final = df_regras_corretas.drop(columns=[col for col in colunas_remover if col in df_regras_corretas.columns])
                df_regras_corretas_final.to_excel(writer, sheet_name='O Regras', index=False)
        
            # Regras Incorretas - Remover colunas indesejadas
            if not df_regras_incorretas.empty:
                df_regras_incorretas_final = df_regras_incorretas.drop(columns=[col for col in colunas_remover if col in df_regras_incorretas.columns])
                df_regras_incorretas_final.to_excel(writer, sheet_name='X Regras', index=False)
        
            # Ofertas Corretas - Remover colunas indesejadas
            if not df_resultados_ofertas.empty:
                df_ofertas_corretas = df_resultados_ofertas[df_resultados_ofertas['Status'] == 'Correto']
                if not df_ofertas_corretas.empty:
                    df_ofertas_corretas_final = df_ofertas_corretas.drop(columns=[col for col in colunas_remover if col in df_ofertas_corretas.columns])
                    df_ofertas_corretas_final.to_excel(writer, sheet_name='O Ofertas', index=False)
        
            # Ofertas Incorretas - Remover colunas indesejadas
            if not df_resultados_ofertas.empty:
                df_ofertas_incorretas = df_resultados_ofertas[df_resultados_ofertas['Status'] == 'Incorreto']
                if not df_ofertas_incorretas.empty:
                    df_ofertas_incorretas_final = df_ofertas_incorretas.drop(columns=[col for col in colunas_remover if col in df_ofertas_incorretas.columns])
                    df_ofertas_incorretas_final.to_excel(writer, sheet_name='X Ofertas', index=False)
        
            # Sem Oferta - Remover colunas indesejadas
            if not df_sem_oferta_final.empty:
                df_sem_oferta_final_final = df_sem_oferta_final.drop(columns=[col for col in colunas_remover if col in df_sem_oferta_final.columns])
                df_sem_oferta_final_final.to_excel(writer, sheet_name='Sem Oferta', index=False)
        
            # Logs de erros - Não precisa remover colunas (não tem essas colunas)
            if not df_logs_erros.empty:
                df_logs_erros.to_excel(writer, sheet_name='Logs Erros', index=False)

        print("\n=== PROCESSAMENTO CONCLUÍDO COM SUCESSO ===")
        
    except Exception as e:
        print(f"\nERRO CRÍTICO DURANTE O PROCESSAMENTO: {str(e)}")
        raise

if __name__ == "__main__":
    processar_planilhas()