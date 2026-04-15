import pandas as pd
import os

# --- Constantes ---
OUTPUT_FILENAME = 'geral.xlsx'
CSV_ENCODING = 'cp1252'
# --- Fim das Constantes ---

def consolidar_csvs_para_excel():
    """
    Consolida todos os arquivos CSV do diretório atual em um único arquivo Excel.
    Cada arquivo CSV é salvo em uma aba separada no arquivo Excel.
    """
    # Define os caminhos baseados no local onde o script está salvo
    diretorio_script = os.path.dirname(os.path.abspath(__file__))
    diretorio_csvs = os.path.join(diretorio_script, 'csvs')
    caminho_saida = os.path.join(diretorio_script, OUTPUT_FILENAME)

    # Verifica se a pasta 'csvs' existe
    if not os.path.exists(diretorio_csvs):
        print(f"Erro: A pasta '{diretorio_csvs}' não foi encontrada.")
        return

    # Lista todos os arquivos .csv dentro da pasta 'csvs'
    arquivos_csv = [f for f in os.listdir(diretorio_csvs) if f.lower().endswith('.csv')]

    if not arquivos_csv:
        print(f"Nenhum arquivo .csv encontrado em: {diretorio_csvs}")
        return

    print(f"Encontrados {len(arquivos_csv)} arquivos. Iniciando consolidação...")

    processou_algum = False
    dados_para_analise = []
    detalhes_para_analise = []
    # Utiliza o ExcelWriter para gerenciar a criação de múltiplas abas
    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        for arquivo in arquivos_csv:
            caminho_csv = os.path.join(diretorio_csvs, arquivo)
            
            # Nome da aba: nome do arquivo sem extensão (limitado a 31 caracteres)
            nome_aba = os.path.splitext(arquivo)[0][:31]
            
            try:
                # Lê o CSV pulando as 16 linhas iniciais de cabeçalho decorativo
                # A linha 17 (índice 16) será tratada como o cabeçalho das colunas
                df = pd.read_csv(
                    caminho_csv, 
                    encoding=CSV_ENCODING, 
                    sep=';', 
                    skiprows=16,
                    on_bad_lines='warn'
                )
                
                # AJUSTES REFINADOS:

                # 1º Ajuste Refinado: remove linhas onde a coluna H (8ª coluna, índice 7) é '-'
                if len(df.columns) >= 8:
                    df = df[df.iloc[:, 7] != '-']
                
                # 2º Ajuste Refinado: trata duplicatas na coluna K (11ª coluna, índice 10)
                if len(df.columns) >= 11:
                    # Mantém a linha se for '-' OU se não for uma duplicata (mantendo a primeira ocorrência)
                    mask_manter = (df.iloc[:, 10] == '-') | (~df.iloc[:, 10].duplicated(keep='first'))
                    df = df[mask_manter]
                
                # 3º Ajuste Refinado: trata duplicatas na coluna L (12ª coluna, índice 11)
                if len(df.columns) >= 12:
                    # Mantém a linha se for '-' OU se não for uma duplicata (mantendo a primeira ocorrência)
                    mask_manter_l = (df.iloc[:, 11] == '-') | (~df.iloc[:, 11].duplicated(keep='first'))
                    df = df[mask_manter_l]
                
                # Coleta dados para o resumo (Aba ANALISYS)
                partes = nome_aba.split('-')
                if len(partes) >= 2:
                    equipe = f"EQUIPE {partes[0]}"
                    micro = f"MICRO {partes[1]}"
                    dados_para_analise.append({
                        'EQUIPE': equipe,
                        'MICRO': micro,
                        'CONTAGEM': len(df)
                    })
                    
                    # B) Coleta colunas A e B para o resumo detalhado por logradouro
                    if len(df.columns) >= 2:
                        temp_df = df.iloc[:, [0, 1]].copy()
                        temp_df.columns = ['TIPO', 'LOGRADOURO']
                        temp_df['CHAVE_EQUIPE_MICRO'] = f"{equipe} - {micro}"
                        # Organiza as colunas na ordem: Equipe-Micro, Tipo, Nome
                        detalhes_para_analise.append(temp_df[['CHAVE_EQUIPE_MICRO', 'TIPO', 'LOGRADOURO']])
                
                df.to_excel(writer, sheet_name=nome_aba, index=False)
                

                #CONFIGURAÇÕES ADICIONAIS/CUSTOMIZAÇÕES:

                # I) Congelar a linha 1 (cabeçalho) e obter referência da aba
                ws = writer.sheets[nome_aba]
                ws.freeze_panes = 'A2'

                # II) Ocultar colunas 'F' e 'G' para proteger dados sensíveis
                ws.column_dimensions['F'].hidden = True
                ws.column_dimensions['G'].hidden = True

                # III) Ativar o filtro automático nas colunas 'A' a 'O'
                ws.auto_filter.ref = f"A1:O{ws.max_row}"
                
                print(f"Processado: {arquivo} -> Aba: {nome_aba}")
                processou_algum = True
            except Exception as e:
                print(f"Erro ao processar {arquivo}: {e}")
                continue

        # A) Geração da aba ANALISYS com o resumo consolidado
        if dados_para_analise:
            df_resumo = pd.DataFrame(dados_para_analise)
            # Ordena os dados para garantir a sequência correta de Equipes e Micros
            df_resumo = df_resumo.sort_values(['EQUIPE', 'MICRO'])
            
            linha_atual = 0
            # 1º Ajuste Final: Formatação da primeira tabela em blocos por EQUIPE
            for equipe_nome, grupo in df_resumo.groupby('EQUIPE'):
                # Escreve o cabeçalho do bloco (EQUIPE | MICRO)
                pd.DataFrame([['EQUIPE', 'MICRO', '']]).to_excel(
                    writer, sheet_name='ANALISYS', startrow=linha_atual, index=False, header=False
                )
                linha_atual += 1
                
                # Escreve os dados das Micros (EQUIPE | MICRO | CONTAGEM)
                grupo[['EQUIPE', 'MICRO', 'CONTAGEM']].to_excel(
                    writer, sheet_name='ANALISYS', startrow=linha_atual, index=False, header=False
                )
                linha_atual += len(grupo)
                
                # Escreve a linha de TOTAL da equipe específica
                total_equipe = grupo['CONTAGEM'].sum()
                pd.DataFrame([['TOTAL', equipe_nome, total_equipe]]).to_excel(
                    writer, sheet_name='ANALISYS', startrow=linha_atual, index=False, header=False
                )
                
                # Incrementa para deixar uma linha em branco entre blocos de equipes
                linha_atual += 2
            
            # B) Segunda tabela: contagem por Equipe-Micro + Logradouro (Colunas A e B)
            if detalhes_para_analise:
                df_b = pd.concat(detalhes_para_analise)
                # Agrupa pelas 3 colunas e conta as ocorrências (size)
                resumo_b = df_b.groupby(['CHAVE_EQUIPE_MICRO', 'TIPO', 'LOGRADOURO']).size().reset_index()
                
                # 2º Ajuste Final: Mantém a estrutura e pula duas linhas após o fim da primeira tabela
                resumo_b.to_excel(writer, sheet_name='ANALISYS', startrow=linha_atual + 1, index=False, header=False)

            print("Resumo 'ANALISYS' gerado com sucesso.")

    if processou_algum:
        print(f"\nProcesso concluído! Arquivo criado em: {caminho_saida}")
    else:
        print("\nNenhum arquivo foi processado com sucesso.")

if __name__ == "__main__":
    consolidar_csvs_para_excel()