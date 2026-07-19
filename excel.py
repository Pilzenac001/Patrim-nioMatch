import pandas as pd
import os
from normalizador import normalizar_texto

# Mapeamento de termos comuns para identificar cabeçalhos das colunas (Exato vs Parcial)
MAPEAMENTO_COLUNAS = {
    "patrimonio": {
        "exato": ["patrimonio", "codigo", "registro", "nº", "no", "id", "inventario", "plaqueta", "patr"],
        "contem": ["patrimonio", "codigo", "registro", "inventario", "plaqueta", "cod. bem", "cod bem"]
    },
    "descricao": {
        "exato": ["descricao", "especificacao", "nome", "bem", "desc", "material", "historico", "detalhe"],
        "contem": ["descricao", "especificacao", "nome do bem", "nome do item", "nome do material", "desc. bem", "desc bem", "especif", "descricaomaterial"]
    },
    "marca": {
        "exato": ["marca", "fabricante"],
        "contem": ["marca", "fabricante", "fabr"]
    },
    "modelo": {
        "exato": ["modelo"],
        "contem": ["modelo"]
    },
    "local": {
        "exato": ["local", "setor", "sala", "destino", "lotacao", "departamento", "divisao", "predio"],
        "contem": ["localizacao", "lotacao", "setor", "departamento", "lugar", "unidade", "dep."]
    },
    "observacao": {
        "exato": ["observacao", "observacoes", "obs", "nota", "detalhes"],
        "contem": ["observac", "obs", "nota"]
    }
}

def detectar_colunas(df):
    """
    Detecta automaticamente as colunas da planilha e retorna um dicionário
    mapeando as chaves internas para os nomes de colunas reais presentes no DataFrame.
    Prioriza correspondências exatas em todas as chaves antes de tentar correspondências parciais.
    """
    mapeamento_final = {}
    colunas_df = list(df.columns)
    colunas_df_norm = [normalizar_texto(str(col)) for col in colunas_df]
    
    # 1ª Passada: Buscar correspondência EXATA para todas as chaves
    for chave, config_busca in MAPEAMENTO_COLUNAS.items():
        mapeado = False
        for alias in config_busca["exato"]:
            alias_norm = normalizar_texto(alias)
            for i, col_norm in enumerate(colunas_df_norm):
                if col_norm == alias_norm or f" {alias_norm} " in f" {col_norm} ":
                    if colunas_df[i] not in mapeamento_final.values():
                        mapeamento_final[chave] = colunas_df[i]
                        mapeado = True
                        break
            if mapeado:
                break
                
    # 2ª Passada: Para chaves que não foram mapeadas, busca por correspondência parcial (contém)
    for chave, config_busca in MAPEAMENTO_COLUNAS.items():
        if chave in mapeamento_final:
            continue
            
        mapeado = False
        for alias in config_busca["contem"]:
            alias_norm = normalizar_texto(alias)
            for i, col_norm in enumerate(colunas_df_norm):
                if alias_norm in col_norm:
                    if colunas_df[i] not in mapeamento_final.values():
                        mapeamento_final[chave] = colunas_df[i]
                        mapeado = True
                        break
            if mapeado:
                break
                
    # Caso crítico: se a descrição ainda não foi mapeada, procura qualquer coluna contendo "desc" ou "nome" que sobrou
    if "descricao" not in mapeamento_final:
        for i, col_norm in enumerate(colunas_df_norm):
            if "desc" in col_norm or "nome" in col_norm or "bem" in col_norm:
                if colunas_df[i] not in mapeamento_final.values():
                    mapeamento_final["descricao"] = colunas_df[i]
                    break
                    
    return mapeamento_final

def carregar_planilha(caminho_arquivo):
    """
    Carrega uma planilha Excel (.xlsx ou .xls) em um DataFrame Pandas.
    Se a planilha tiver linhas vazias no topo antes dos cabeçalhos,
    o algoritmo procura a primeira linha contendo dados estruturados.
    """
    if not os.path.exists(caminho_arquivo):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")
        
    try:
        # Carregamos inicialmente para analisar
        df = pd.read_excel(caminho_arquivo)
    except Exception as e:
        raise ValueError(f"Erro ao ler o arquivo {caminho_arquivo}: {str(e)}")
        
    # Se o DataFrame for vazio, retorna
    if df.empty:
        return df
        
    # Heurística para pular cabeçalhos vazios/sujeira no topo da planilha
    # Se a primeira linha tiver muitas colunas sem nome (ex: "Unnamed: 0") e a próxima linha
    # tiver menos "Unnamed", podemos ajustar os cabeçalhos.
    unnamed_cols = [c for c in df.columns if "Unnamed:" in str(c)]
    if len(unnamed_cols) > len(df.columns) / 2 and len(df) > 1:
        # Tenta ler as primeiras 10 linhas para achar a linha de cabeçalho real
        for skip in range(1, min(10, len(df))):
            df_temp = pd.read_excel(caminho_arquivo, skiprows=skip)
            temp_unnamed = [c for c in df_temp.columns if "Unnamed:" in str(c)]
            if len(temp_unnamed) < len(unnamed_cols) and not df_temp.empty:
                df = df_temp
                break
                
    # Converte todas as colunas que puderem ser strings/números limpando strings nulas
    df = df.fillna("")
    return df
