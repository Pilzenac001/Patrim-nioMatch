import os
import xlsxwriter
from estatisticas import calcular_percentuais

def gerar_relatorio_excel(resultados, caminho_saida):
    """
    Grava os resultados da conciliação patrimonial em uma planilha Excel
    completamente formatada e estilizada usando xlsxwriter.
    """
    # Garantir que o diretório de destino exista
    diretorio = os.path.dirname(caminho_saida)
    if diretorio and not os.path.exists(diretorio):
        os.makedirs(diretorio, exist_ok=True)
        
    # Verificar se o arquivo já existe e está aberto/bloqueado
    if os.path.exists(caminho_saida):
        try:
            with open(caminho_saida, 'r+'):
                pass
        except IOError:
            raise PermissionError(
                f"O arquivo de relatório '{os.path.basename(caminho_saida)}' está aberto no Excel ou em outro programa. "
                "Por favor, feche-o antes de iniciar o processamento."
            )
            
    workbook = xlsxwriter.Workbook(caminho_saida)
    
    # --- Paleta de Cores & Fontes Modernas (Glassmorphism inspired) ---
    FONT_FAMILY = "Segoe UI"
    
    # Cores Principais
    COLOR_PRIMARY = "#1A365D"    # Azul Escuro Premium
    COLOR_SECONDARY = "#2B6CB0"  # Azul Destaque
    COLOR_WHITE = "#FFFFFF"
    COLOR_LIGHT_BG = "#F7FAFC"   # Fundo suave
    COLOR_BORDER = "#E2E8F0"     # Borda elegante
    
    # Cores de Status
    COLOR_CONFIRMED_BG = "#E6FFFA"  # Verde Menta Suave
    COLOR_CONFIRMED_TEXT = "#004D40"
    COLOR_PROBABLE_BG = "#FEFCBF"   # Amarelo Suave
    COLOR_PROBABLE_TEXT = "#744210"
    COLOR_MANUAL_BG = "#FFF5F5"     # Coral Claro Suave
    COLOR_MANUAL_TEXT = "#742A2A"
    
    # Formatos de Texto
    fmt_header = workbook.add_format({
        'bold': True,
        'align': 'left',
        'valign': 'vcenter',
        'fg_color': COLOR_PRIMARY,
        'font_color': COLOR_WHITE,
        'font_name': FONT_FAMILY,
        'font_size': 10,
        'border': 1,
        'border_color': COLOR_BORDER
    })
    
    fmt_cell = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'font_name': FONT_FAMILY,
        'font_size': 9,
        'border': 1,
        'border_color': COLOR_BORDER
    })
    
    fmt_cell_center = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'font_name': FONT_FAMILY,
        'font_size': 9,
        'border': 1,
        'border_color': COLOR_BORDER
    })
    
    fmt_cell_pct = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'font_name': FONT_FAMILY,
        'font_size': 9,
        'border': 1,
        'border_color': COLOR_BORDER,
        'num_format': '0.0"%"'
    })
    
    # Formatos por aba de status
    fmt_cell_confirmed = workbook.add_format({
        'align': 'left', 'valign': 'vcenter', 'font_name': FONT_FAMILY, 'font_size': 9,
        'border': 1, 'border_color': COLOR_BORDER, 'fg_color': COLOR_CONFIRMED_BG, 'font_color': COLOR_CONFIRMED_TEXT
    })
    fmt_cell_probable = workbook.add_format({
        'align': 'left', 'valign': 'vcenter', 'font_name': FONT_FAMILY, 'font_size': 9,
        'border': 1, 'border_color': COLOR_BORDER, 'fg_color': COLOR_PROBABLE_BG, 'font_color': COLOR_PROBABLE_TEXT
    })
    fmt_cell_manual = workbook.add_format({
        'align': 'left', 'valign': 'vcenter', 'font_name': FONT_FAMILY, 'font_size': 9,
        'border': 1, 'border_color': COLOR_BORDER, 'fg_color': COLOR_MANUAL_BG, 'font_color': COLOR_MANUAL_TEXT
    })
    
    # Cabeçalho das Abas de Matching
    headers_match = [
        "Patrimônio GEMAT", "Descrição GEMAT", "Marca GEMAT", "Modelo GEMAT", "Local GEMAT", "Observação GEMAT",
        "Patrimônio SAM", "Descrição SAM", "Marca SAM", "Modelo SAM", "Local SAM", "Observação SAM",
        "Confiança", "Motivo / Status"
    ]
    
    def preencher_aba_match(worksheet, dados, custom_fmt=None):
        # Escreve cabeçalhos
        for col_idx, text in enumerate(headers_match):
            worksheet.write(0, col_idx, text, fmt_header)
        worksheet.set_row(0, 26)
        
        if not dados:
            worksheet.write(1, 0, "Nenhuma correspondência registrada nesta categoria.", fmt_cell)
            worksheet.set_row(1, 20)
            return
            
        for row_idx, item in enumerate(dados, start=1):
            worksheet.set_row(row_idx, 20)
            fmt_linha = custom_fmt if custom_fmt else fmt_cell
            
            worksheet.write(row_idx, 0, item.get("gemat_patrimonio", ""), fmt_cell_center)
            worksheet.write(row_idx, 1, item.get("gemat_descricao", ""), fmt_linha)
            worksheet.write(row_idx, 2, item.get("gemat_marca", ""), fmt_linha)
            worksheet.write(row_idx, 3, item.get("gemat_modelo", ""), fmt_linha)
            worksheet.write(row_idx, 4, item.get("gemat_local", ""), fmt_linha)
            worksheet.write(row_idx, 5, item.get("gemat_observacao", ""), fmt_linha)
            
            worksheet.write(row_idx, 6, item.get("sam_patrimonio", ""), fmt_cell_center)
            worksheet.write(row_idx, 7, item.get("sam_descricao", ""), fmt_linha)
            worksheet.write(row_idx, 8, item.get("sam_marca", ""), fmt_linha)
            worksheet.write(row_idx, 9, item.get("sam_modelo", ""), fmt_linha)
            worksheet.write(row_idx, 10, item.get("sam_local", ""), fmt_linha)
            worksheet.write(row_idx, 11, item.get("sam_observacao", ""), fmt_linha)
            
            worksheet.write(row_idx, 12, item.get("confianca", 0.0), fmt_cell_pct)
            worksheet.write(row_idx, 13, item.get("motivo", ""), fmt_linha)
            
        # Largura das colunas padrão ajustadas
        larguras = [18, 30, 15, 15, 20, 25, 18, 30, 15, 15, 20, 25, 12, 35]
        for idx, w in enumerate(larguras):
            worksheet.set_column(idx, idx, w)

    # 1. Correspondências Confirmadas
    ws_conf = workbook.add_worksheet("Conf. Confirmadas")
    preencher_aba_match(ws_conf, resultados.get("confirmadas", []), fmt_cell_confirmed)
    
    # 2. Correspondências Prováveis
    ws_prov = workbook.add_worksheet("Conf. Provaveis")
    preencher_aba_match(ws_prov, resultados.get("provaveis", []), fmt_cell_probable)
    
    # 3. Conferência Manual
    ws_man = workbook.add_worksheet("Conferencia Manual")
    preencher_aba_match(ws_man, resultados.get("manual", []), fmt_cell_manual)
    
    # 4. Apenas GEMAT
    ws_apenas_gemat = workbook.add_worksheet("Apenas GEMAT")
    headers_gemat = ["Patrimônio GEMAT", "Descrição GEMAT", "Marca GEMAT", "Modelo GEMAT", "Local GEMAT", "Observação GEMAT", "Status"]
    for col_idx, text in enumerate(headers_gemat):
        ws_apenas_gemat.write(0, col_idx, text, fmt_header)
    ws_apenas_gemat.set_row(0, 26)
    
    dados_apenas_gemat = resultados.get("apenas_gemat", [])
    if not dados_apenas_gemat:
        ws_apenas_gemat.write(1, 0, "Nenhum item restou exclusivo no GEMAT.", fmt_cell)
        ws_apenas_gemat.set_row(1, 20)
    else:
        for row_idx, item in enumerate(dados_apenas_gemat, start=1):
            ws_apenas_gemat.set_row(row_idx, 20)
            ws_apenas_gemat.write(row_idx, 0, item.get("gemat_patrimonio", ""), fmt_cell_center)
            ws_apenas_gemat.write(row_idx, 1, item.get("gemat_descricao", ""), fmt_cell)
            ws_apenas_gemat.write(row_idx, 2, item.get("gemat_marca", ""), fmt_cell)
            ws_apenas_gemat.write(row_idx, 3, item.get("gemat_modelo", ""), fmt_cell)
            ws_apenas_gemat.write(row_idx, 4, item.get("gemat_local", ""), fmt_cell)
            ws_apenas_gemat.write(row_idx, 5, item.get("gemat_observacao", ""), fmt_cell)
            ws_apenas_gemat.write(row_idx, 6, item.get("motivo", ""), fmt_cell)
    
    larguras_gemat = [18, 35, 15, 15, 20, 25, 35]
    for idx, w in enumerate(larguras_gemat):
        ws_apenas_gemat.set_column(idx, idx, w)

    # 5. Apenas SAM
    ws_apenas_sam = workbook.add_worksheet("Apenas SAM")
    headers_sam = ["Patrimônio SAM", "Descrição SAM", "Marca SAM", "Modelo SAM", "Local SAM", "Observação SAM"]
    for col_idx, text in enumerate(headers_sam):
        ws_apenas_sam.write(0, col_idx, text, fmt_header)
    ws_apenas_sam.set_row(0, 26)
    
    dados_apenas_sam = resultados.get("apenas_sam", [])
    if not dados_apenas_sam:
        ws_apenas_sam.write(1, 0, "Nenhum item restou exclusivo no SAM.", fmt_cell)
        ws_apenas_sam.set_row(1, 20)
    else:
        for row_idx, item in enumerate(dados_apenas_sam, start=1):
            ws_apenas_sam.set_row(row_idx, 20)
            ws_apenas_sam.write(row_idx, 0, item.get("sam_patrimonio", ""), fmt_cell_center)
            ws_apenas_sam.write(row_idx, 1, item.get("sam_descricao", ""), fmt_cell)
            ws_apenas_sam.write(row_idx, 2, item.get("sam_marca", ""), fmt_cell)
            ws_apenas_sam.write(row_idx, 3, item.get("sam_modelo", ""), fmt_cell)
            ws_apenas_sam.write(row_idx, 4, item.get("sam_local", ""), fmt_cell)
            ws_apenas_sam.write(row_idx, 5, item.get("sam_observacao", ""), fmt_cell)
            
    larguras_sam = [18, 35, 15, 15, 20, 25]
    for idx, w in enumerate(larguras_sam):
        ws_apenas_sam.set_column(idx, idx, w)

    # 6. Estatísticas (Dashboard simples)
    ws_est = workbook.add_worksheet("Estatisticas")
    estats_raw = resultados.get("estatisticas", {})
    estats = calcular_percentuais(estats_raw)
    
    # Título do Dashboard
    fmt_title = workbook.add_format({
        'bold': True, 'font_name': FONT_FAMILY, 'font_size': 14, 'font_color': COLOR_PRIMARY
    })
    ws_est.write("A1", "Painel de Resultados - Saneamento Patrimonial", fmt_title)
    
    fmt_card_title = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_name': FONT_FAMILY,
        'font_size': 10, 'fg_color': COLOR_SECONDARY, 'font_color': COLOR_WHITE, 'border': 1, 'border_color': COLOR_BORDER
    })
    
    fmt_card_value = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_name': FONT_FAMILY,
        'font_size': 14, 'border': 1, 'border_color': COLOR_BORDER, 'fg_color': COLOR_LIGHT_BG
    })
    
    fmt_card_value_pct = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_name': FONT_FAMILY,
        'font_size': 14, 'border': 1, 'border_color': COLOR_BORDER, 'fg_color': COLOR_LIGHT_BG,
        'num_format': '0.0"%"'
    })
    
    # Criar cards de dados em formato tabela
    indicadores = [
        ("Total GEMAT", estats.get("total_gemat", 0), "A4", "A5"),
        ("Total SAM", estats.get("total_sam", 0), "B4", "B5"),
        ("Conciliados (Confirmados + Prováveis)", estats.get("confirmadas_qtd", 0) + estats.get("provaveis_qtd", 0), "C4", "C5"),
        ("Percentual Sucesso", estats.get("percentual_sucesso", 0.0), "D4", "D5"),
        ("Total de Pendências (Manual + Apenas GEMAT)", estats.get("total_pendencias", 0), "E4", "E5")
    ]
    
    for title, val, cell_t, cell_v in indicadores:
        ws_est.write(cell_t, title, fmt_card_title)
        if "Percentual" in title:
            ws_est.write(cell_v, val, fmt_card_value_pct)
        else:
            ws_est.write(cell_v, val, fmt_card_value)
            
    # Ajuste de altura das linhas do dashboard
    ws_est.set_row(3, 25)
    ws_est.set_row(4, 30)
    
    # Tabela Detalhada abaixo dos cards
    ws_est.write("A8", "Resumo de Categorização", workbook.add_format({'bold': True, 'font_name': FONT_FAMILY, 'font_size': 11, 'font_color': COLOR_PRIMARY}))
    
    headers_table = ["Categoria / Aba", "Quantidade", "Proporção (% GEMAT)"]
    for col_idx, text in enumerate(headers_table):
        ws_est.write(8, col_idx, text, fmt_header)
        
    categorias_est = [
        ("Conf. Confirmadas (Alta Confiança)", estats.get("confirmadas_qtd", 0)),
        ("Conf. Prováveis (Média Confiança)", estats.get("provaveis_qtd", 0)),
        ("Conferência Manual (Baixa Confiança ou Conflito)", estats.get("manual_qtd", 0)),
        ("Apenas GEMAT (Sem correspondente)", estats.get("apenas_gemat_qtd", 0))
    ]
    
    total_g = estats.get("total_gemat", 1) # evitar divisão por zero
    for row_idx, (cat_name, qtd) in enumerate(categorias_est, start=9):
        ws_est.set_row(row_idx, 20)
        ws_est.write(row_idx, 0, cat_name, fmt_cell)
        ws_est.write(row_idx, 1, qtd, fmt_cell_center)
        ws_est.write(row_idx, 2, (qtd / total_g) * 100, fmt_cell_pct)
        
    # Adicionar totalizadores na tabela
    total_row = 9 + len(categorias_est)
    ws_est.set_row(total_row, 20)
    fmt_total = workbook.add_format({'bold': True, 'font_name': FONT_FAMILY, 'font_size': 9, 'border': 1, 'border_color': COLOR_BORDER})
    fmt_total_center = workbook.add_format({'bold': True, 'align': 'center', 'font_name': FONT_FAMILY, 'font_size': 9, 'border': 1, 'border_color': COLOR_BORDER})
    fmt_total_pct = workbook.add_format({'bold': True, 'align': 'center', 'font_name': FONT_FAMILY, 'font_size': 9, 'border': 1, 'border_color': COLOR_BORDER, 'num_format': '0.0"%"'})
    
    ws_est.write(total_row, 0, "Total Geral GEMAT", fmt_total)
    ws_est.write(total_row, 1, estats.get("total_gemat", 0), fmt_total_center)
    ws_est.write(total_row, 2, 100.0, fmt_total_pct)
    
    ws_est.set_column(0, 0, 42)
    ws_est.set_column(1, 1, 15)
    ws_est.set_column(2, 2, 20)
    ws_est.set_column(3, 3, 20)
    ws_est.set_column(4, 4, 35)

    # 7. Log de Processamento
    ws_log = workbook.add_worksheet("Log")
    ws_log.write(0, 0, "Registro de Eventos / Execução", fmt_header)
    ws_log.set_row(0, 26)
    
    logs = resultados.get("logs", [])
    for row_idx, log_line in enumerate(logs, start=1):
        ws_log.set_row(row_idx, 18)
        ws_log.write(row_idx, 0, log_line, fmt_cell)
        
    ws_log.set_column(0, 0, 100)
    
    workbook.close()
