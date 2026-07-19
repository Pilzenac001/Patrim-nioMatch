import os
from excel import carregar_planilha, detectar_colunas
from matcher import carregar_configuracao, processar_matching
from relatorio import gerar_relatorio_excel
from normalizador import carregar_sinonimos

def testar_fluxo():
    print("Iniciando teste de fluxo automático...")
    
    # 1. Carregar Configurações e Sinônimos
    config = carregar_configuracao("config.json")
    sinonimos = carregar_sinonimos("sinonimos.json")
    print(f"Configurações carregadas: {config}")
    print(f"Sinônimos carregados: {len(sinonimos)} itens.")
    
    # 2. Carregar Planilhas
    gemat_path = "entrada/GEMAT.xlsx"
    sam_path = "entrada/SAM.xlsx"
    
    print(f"Lendo planilha GEMAT de: {gemat_path}")
    df_gemat = carregar_planilha(gemat_path)
    colunas_gemat = detectar_colunas(df_gemat)
    print(f"Colunas GEMAT mapeadas: {colunas_gemat}")
    
    print(f"Lendo planilha SAM de: {sam_path}")
    df_sam = carregar_planilha(sam_path)
    colunas_sam = detectar_colunas(df_sam)
    print(f"Colunas SAM mapeadas: {colunas_sam}")
    
    # 3. Processar Matching
    def callback_progresso(msg):
        print(f"  [Progresso] {msg}")
        
    print("Executando o matcher...")
    resultados = processar_matching(
        df_gemat, colunas_gemat,
        df_sam, colunas_sam,
        config, sinonimos,
        callback_progresso
    )
    
    # 4. Gravar Relatório
    saida_path = "saida/Relatorio_PatrimonioMatch.xlsx"
    print(f"Gravando relatório final em: {saida_path}")
    gerar_relatorio_excel(resultados, saida_path)
    
    # Verificar se o arquivo foi criado
    if os.path.exists(saida_path):
        print(f"SUCESSO! O relatório foi criado e possui {os.path.getsize(saida_path)} bytes.")
    else:
        print("ERRO: O arquivo de saída não foi encontrado!")
        
if __name__ == "__main__":
    testar_fluxo()
