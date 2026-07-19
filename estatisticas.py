def calcular_percentuais(estats):
    """
    Calcula o percentual de sucesso do matching e o total de pendências
    para subsidiar a aba Estatísticas.
    """
    total_gemat = estats.get("total_gemat", 0)
    confirmadas = estats.get("confirmadas_qtd", 0)
    provaveis = estats.get("provaveis_qtd", 0)
    manual = estats.get("manual_qtd", 0)
    apenas_gemat = estats.get("apenas_gemat_qtd", 0)
    
    # O sucesso de saneamento é a quantidade de correspondências confirmadas e prováveis
    total_resolvidos = confirmadas + provaveis
    percentual_sucesso = (total_resolvidos / total_gemat * 100) if total_gemat > 0 else 0.0
    
    # Pendências: registros que necessitam de intervenção ou revisão manual
    total_pendencias = manual + apenas_gemat
    
    estats["percentual_sucesso"] = percentual_sucesso
    estats["total_pendencias"] = total_pendencias
    return estats
