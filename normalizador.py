import re
import json
import os
from unidecode import unidecode

def carregar_sinonimos(caminho_sinonimos):
    """
    Carrega o dicionário de sinônimos de um arquivo JSON.
    """
    if not os.path.exists(caminho_sinonimos):
        return {}
    try:
        with open(caminho_sinonimos, 'r', encoding='utf-8') as f:
            dados = json.load(f)
            return {str(k).upper().strip(): str(v).upper().strip() for k, v in dados.items()}
    except Exception:
        return {}

def normalizar_texto(texto, sinonimos=None):
    """
    Remove acentos, pontuação, reduz espaços, passa para MAIÚSCULO
    e substitui sinônimos conhecidos.
    """
    if not isinstance(texto, str) or not texto.strip():
        return ""
    
    texto = unidecode(texto)
    texto = texto.upper()
    texto = re.sub(r'[^A-Z0-9\s\-/X]', ' ', texto)
    texto = re.sub(r'(?<!\d)0+(\d+)', r'\1', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    
    if sinonimos:
        sinonimos_maiusculos = {str(k).upper(): str(v).upper() for k, v in sinonimos.items()}
        sinonimos_ordenados = sorted(sinonimos_maiusculos.items(), key=lambda x: len(x[0]), reverse=True)
        
        for original, substituto in sinonimos_ordenados:
            pattern = r'\b' + re.escape(original) + r'\b'
            texto = re.sub(pattern, substituto, texto)
            
        texto = re.sub(r'\s+', ' ', texto).strip()
        
    return texto

def extrair_caracteristicas(texto_normalizado):
    """
    Extrai características chave em MAIÚSCULAS através de varredura ampla.
    """
    caract = {
        "tipo": "",
        "marca": "",
        "modelo": "",
        "medidas": "",
        "material": ""
    }
    
    if not texto_normalizado:
        return caract
        
    # 1. MATERIAL
    materiais_comuns = ["ACO", "MADEIRA", "PLASTICO", "VIDRO", "MDF", "METAL", "FORMICA", "TECIDO", "COURO", "FERRO"]
    for mat in materiais_comuns:
        if re.search(r'\b' + mat + r'\b', texto_normalizado):
            caract["material"] = mat
            break
            
    # 2. MEDIDAS
    medida_pat = re.search(r'(\d+[\s]*[xX][\s]*\d+)|(\d+[\s]*(?:CM|MM|M|PORTAS|GAVETAS)\b)', texto_normalizado)
    if medida_pat:
        caract["medidas"] = medida_pat.group(0).replace(" ", "")
        
    # 3. MARCA
    marcas_comuns = [
        "DELL", "HP", "LENOVO", "SAMSUNG", "LG", "POSITIVO", "EPSON", "ITAUTEC", "FLEXFORM", 
        "CAVALETTI", "ALBERFLEX", "CONSUL", "BRASTEMP", "ELECTROLUX", "PHILCO", "AOC", "BENQ", 
        "SONY", "INTELBRAS", "PANDIN", "RICOH", "BROTHER", "LOGITECH", "CANON", "ACER", "ASUS",
        "MARELLI", "GIROFLEX", "COBRA", "CISCO", "DESKTOP"
    ]
    for marca in marcas_comuns:
        if re.search(r'\b' + marca + r'\b', texto_normalizado):
            caract["marca"] = marca
            break
            
    # 4. TIPO DO BEM (CORREÇÃO CRÍTICA: Busca mapeada por palavras-chave em qualquer posição do texto)
    dicionario_tipos = {
        "AR CONDICIONADO": ["AR CONDICIONADO", "CONDICIONADOR AR", "SPLIT"],
        "LAPTOP": ["LAPTOP", "NOTEBOOK", "COMPUTADOR PORTATIL", "MICROCOMPUTADOR PORTATIL"],
        "MICROCOMPUTADOR": ["MICROCOMPUTADOR", "COMPUTADOR", "DESKTOP", "CPU", "GABINETE"],
        "MONITOR": ["MONITOR", "TELA", "DISP"],
        "IMPRESSORA": ["IMPRESSORA", "MULTIFUNCIONAL", "RICOH", "BROTHER", "EPSON"],
        "NOBREAK": ["NO BREAK", "NOBREAK", "ESTABILIZADOR"],
        "PROJETOR": ["PROJETOR", "DATA SHOW", "DATASHOW"],
        "QUADRO": ["QUADRO BRANCO", "QUADRO VERDE", "QUADRO"],
        "BEBEDOURO": ["BEBEDOURO COLUNA", "BEBEDOURO MESA", "BEBEDOURO", "PURIFICADOR"],
        "CADEIRA": ["CADEIRA UNIVERSITARIA", "CADEIRA GIRATORIA", "CADEIRA FIXA", "CADEIRA", "POLTRONA"],
        "MESA": ["MESA REUNIAO", "MESA TRABALHO", "MESA ESCRITORIO", "MESA", "ESCRIVANINHA"],
        "ARMARIO": ["ARMARIO ACO", "ARMARIO MADEIRA", "ARMARIO ALTO", "ARMARIO BAIXO", "ARMARIO", "ARQUIVO"],
        "BALCAO": ["BALCAO", "BALCOES"],
        "GAVETEIRO": ["GAVETEIRO", "VOLANTE"]
    }
    
    tipo_encontrado = False
    for tipo_padrao, variantes in dicionario_tipos.items():
        for variante in variantes:
            if re.search(r'\b' + re.escape(variante) + r'\b', texto_normalizado):
                caract["tipo"] = tipo_padrao
                tipo_encontrado = True
                break
        if tipo_encontrado:
            break
            
    # Fallback se não bater com o dicionário: usa a lógica de primeira palavra válida
    if not caract["tipo"]:
        stop_words = ["DE", "PARA", "COM", "EM", "UM", "UMA", "O", "A", "OS", "AS", "DO", "DA", "DOS", "DAS", "C/", "SEM", "SOB"]
        palavras = [p for p in texto_normalizado.split() if p not in stop_words]
        if palavras:
            caract["tipo"] = palavras[0]
                
    # 5. MODELO
    if caract["marca"]:
        marca_idx = texto_normalizado.find(caract["marca"])
        restante = texto_normalizado[marca_idx + len(caract["marca"]):].strip()
        palavras_restante = restante.split()
        if palavras_restante:
            if len(palavras_restante) > 1 and len(palavras_restante[0]) > 2:
                caract["modelo"] = f"{palavras_restante[0]} {palavras_restante[1]}"
            else:
                caract["modelo"] = palavras_restante[0]
    else:
        match_modelo = re.search(r'\b([A-Z]+\d+|\d+[A-Z]+|[A-Z]+-\d+|\d+-[A-Z]+)\b', texto_normalizado)
        if match_modelo:
            caract["modelo"] = match_modelo.group(0)
            
    return caract