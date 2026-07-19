import os
import json
import time
import re
import pandas as pd
from rapidfuzz import fuzz
from normalizador import normalizar_texto, extrair_caracteristicas

def carregar_configuracao(caminho_config):
    config_padrao = {
        "confianca_confirmada": 65,
        "confianca_provavel": 35,
        "confianca_minima": 10,
        "usar_modelo": True,
        "usar_marca": True,
        "usar_medidas": True,
        "pesos": {
            "tipo_bem": 40,
            "marca": 15,
            "modelo": 20,
            "medidas": 10,
            "material": 5,
            "similaridade_textual": 10
        }
    }
    if not os.path.exists(caminho_config):
        return config_padrao
    try:
        with open(caminho_config, 'r', encoding='utf-8') as f:
            user_config = json.load(f)
            for k, v in config_padrao.items():
                if k not in user_config:
                    user_config[k] = v
                elif k == "pesos" and isinstance(user_config[k], dict):
                    for pk, pv in v.items():
                        if pk not in user_config[k]:
                            user_config[k][pk] = pv
            return user_config
    except Exception:
        return config_padrao

def calcular_confianca(gemat_caract, sam_caract, gemat_desc, sam_desc, config):
    pesos = config.get("pesos", {})
    soma_pesos = 0
    soma_pontos = 0
    
    def possui_valor(val):
        return val and str(val).strip() and str(val).upper() not in ["NULO", "NAN", "IGNORADO", "", "N/A"]

    sim_txt = fuzz.token_set_ratio(gemat_desc, sam_desc)

    # 1. Tipo de Bem
    if possui_valor(gemat_caract.get("tipo")) and possui_valor(sam_caract.get("tipo")):
        w_tipo = pesos.get("tipo_bem", 40)
        soma_pesos += w_tipo
        if gemat_caract["tipo"] == sam_caract["tipo"]:
            soma_pontos += w_tipo * 100
        elif gemat_caract["tipo"] in sam_caract["tipo"] or sam_caract["tipo"] in gemat_caract["tipo"]:
            soma_pontos += w_tipo * 90
        else:
            soma_pontos += w_tipo * max(sim_txt, 50)
    else:
        w_tipo = pesos.get("tipo_bem", 40)
        soma_pesos += w_tipo
        soma_pontos += w_tipo * sim_txt
        
    # 2. Marca e Modelo
    if config.get("usar_marca", True):
        g, s = gemat_caract.get("marca"), sam_caract.get("marca")
        if possui_valor(g) and possui_valor(s):
            w = pesos.get("marca", 15)
            soma_pesos += w
            soma_pontos += w * (100 if g == s else fuzz.token_sort_ratio(g, s))

    if config.get("usar_modelo", True):
        g, s = gemat_caract.get("modelo"), sam_caract.get("modelo")
        if possui_valor(g) and possui_valor(s):
            w = pesos.get("modelo", 20)
            soma_pesos += w
            soma_pontos += w * fuzz.token_sort_ratio(g, s)
            
    # 3. Similaridade Textual
    w_textual = pesos.get("similaridade_textual", 10)
    soma_pesos += w_textual
    soma_pontos += w_textual * sim_txt
    
    score_final = soma_pontos / soma_pesos if soma_pesos > 0 else 0
    
    if sim_txt >= 75 and score_final < 40:
        score_final = max(score_final, sim_txt * 0.75)
        
    return score_final

def processar_matching(df_gemat, colunas_gemat, df_sam, colunas_sam, config, sinonimos, callback_progresso=None):
    tempo_inicio = time.time()
    
    def preparar(df, cols):
        lista = []
        p_col, d_col = cols.get('patrimonio', 'patrimonio'), cols.get('descricao', 'descricao')
        m_col, mod_col = cols.get('marca', 'marca'), cols.get('modelo', 'modelo')
        l_col, o_col = cols.get('local', 'local'), cols.get('observacao', 'observacao')
        
        for idx, row in df.iterrows():
            desc_orig = str(row.get(d_col, "")) if pd.notna(row.get(d_col)) else ""
            desc_limpa = re.sub(r'^\d+\s*[-–—]?\s*', '', desc_orig).strip()
            desc_norm = normalizar_texto(desc_limpa, sinonimos)
            
            lista.append({
                "index": idx, 
                "patrimonio": str(row.get(p_col, "")).split('.')[0].strip(),
                "descricao_orig": desc_orig, 
                "marca_orig": str(row.get(m_col, "")),
                "modelo_orig": str(row.get(mod_col, "")), 
                "local_orig": str(row.get(l_col, "")),
                "obs_orig": str(row.get(o_col, "")), 
                "desc_norm": desc_norm,
                "caract": extrair_caracteristicas(desc_norm)
            })
        return lista

    gemat_processado = preparar(df_gemat, colunas_gemat)
    sam_processado = preparar(df_sam, colunas_sam)

    sam_disponiveis = sam_processado.copy()
    matches_resolvidos = {}

    # ETAPA 1: Matching com Barreira de Tipo (Alta precisão)
    for g_item in gemat_processado:
        melhor_score = -1.0
        melhor_sam_idx = -1
        tipo_gemat = g_item["caract"].get("tipo")
        
        for i, s_item in enumerate(sam_disponiveis):
            tipo_sam = s_item["caract"].get("tipo")
            if tipo_gemat and tipo_sam and tipo_gemat != tipo_sam:
                continue
                
            score = calcular_confianca(g_item["caract"], s_item["caract"], g_item["desc_norm"], s_item["desc_norm"], config)
            if score >= 40.0 and score > melhor_score:
                melhor_score = score
                melhor_sam_idx = i
                
        if melhor_sam_idx != -1:
            sam_escolhido = sam_disponiveis.pop(melhor_sam_idx)
            matches_resolvidos[g_item["index"]] = (sam_escolhido, melhor_score)

    # ETAPA 2: Varredura de Recuperação (Rede de proteção para alta similaridade)
    for g_item in gemat_processado:
        if g_item["index"] in matches_resolvidos or not sam_disponiveis:
            continue
            
        melhor_score = -1.0
        melhor_sam_idx = -1
        
        for i, s_item in enumerate(sam_disponiveis):
            sim_txt = fuzz.token_set_ratio(g_item["desc_norm"], s_item["desc_norm"])
            if sim_txt > 85:
                score = calcular_confianca(g_item["caract"], s_item["caract"], g_item["desc_norm"], s_item["desc_norm"], config)
                if score > melhor_score:
                    melhor_score = score
                    melhor_sam_idx = i
        
        if melhor_sam_idx != -1:
            sam_escolhido = sam_disponiveis.pop(melhor_sam_idx)
            matches_resolvidos[g_item["index"]] = (sam_escolhido, melhor_score)

    # Montagem do Resultado
    confirmadas, provaveis, manual, apenas_gemat = [], [], [], []
    c_conf, c_prov = config.get("confianca_confirmada", 65), config.get("confianca_provavel", 35)
    
    for g_item in gemat_processado:
        g_idx = g_item["index"]
        linha = {
            "gemat_patrimonio": g_item["patrimonio"], "gemat_descricao": g_item["descricao_orig"],
            "gemat_marca": g_item["marca_orig"], "gemat_modelo": g_item["modelo_orig"],
            "gemat_local": g_item["local_orig"], "gemat_observacao": g_item["obs_orig"],
            "sam_patrimonio": "", "sam_descricao": "", "sam_marca": "", "sam_modelo": "",
            "sam_local": "", "sam_observacao": "", "confianca": 0.0, "motivo": ""
        }
        
        if g_idx in matches_resolvidos:
            sam_item, score = matches_resolvidos[g_idx]
            linha.update({
                "sam_patrimonio": sam_item["patrimonio"], "sam_descricao": sam_item["descricao_orig"],
                "sam_marca": sam_item["marca_orig"], "sam_modelo": sam_item["modelo_orig"],
                "sam_local": sam_item["local_orig"], "sam_observacao": sam_item["obs_orig"],
                "confianca": round(score, 1)
            })
            if score >= c_conf:
                linha["motivo"] = f"Alta Confiança ({score:.1f}%)."
                confirmadas.append(linha)
            elif score >= c_prov:
                linha["motivo"] = f"Correspondência Provável ({score:.1f}%)."
                provaveis.append(linha)
            else:
                linha["motivo"] = f"Revisão manual ({score:.1f}%)."
                manual.append(linha)
        else:
            linha["motivo"] = "Nenhum correspondente encontrado."
            apenas_gemat.append(linha)
            
    return {
        "confirmadas": confirmadas, "provaveis": provaveis, "manual": manual, "apenas_gemat": apenas_gemat,
        "estatisticas": {"total_gemat": len(df_gemat), "confirmadas_qtd": len(confirmadas), "provaveis_qtd": len(provaveis), "manual_qtd": len(manual)}
    }