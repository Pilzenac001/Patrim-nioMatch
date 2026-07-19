import pandas as pd
import os

# Dados fictícios do GEMAT
dados_gemat = {
    "Plaqueta/Patrimônio": [1001, 1002, 1003, 1004, 1005, 1006, 1007],
    "Descrição do Bem": [
        "ARMÁRIO DE AÇO C/02 PORTAS PANDIN",
        "NOTEBOOK DELL LATITUDE 3420",
        "CPU ITAUTEC INFOWAY",
        "CARTEIRA UNIVERSITARIA EM MADEIRA E METAL",
        "AR CONDICIONADO CONSUL 12000 BTU",
        "CADEIRA GIRATORIA COM BRACOS COURO",
        "DATA SHOW EPSON POWERLITE"
    ],
    "Setor/Local": ["Secretaria", "TI - Suporte", "Recursos Humanos", "Sala de Aula 2", "Diretoria", "Financeiro", "Auditório"],
    "Observações Gerais": ["Estado regular", "Com carregador", "Funcionando", "Sem avarias", "Instalado na parede", "Assento rasgado", "Sem controle remoto"]
}

# Dados fictícios do SAM
dados_sam = {
    "Registro/Código": ["SAM-098", "SAM-102", "SAM-205", "SAM-401", "SAM-512", "SAM-703", "SAM-800"],
    "Especificação do Material": [
        "armario aco 2 portas",
        "laptop dell latitude 3420",
        "microcomputador itautec infoway",
        "cadeira universitaria madeira",
        "condicionador de ar consul 12000btu",
        "projetor epson powerlite",
        "gaveteiro mdf 3 gavetas"  # Sem correspondente no GEMAT
    ],
    "Lotação Atual": ["Secretaria Geral", "Departamento TI", "RH", "Bloco A", "Gabinete Direção", "Auditório Central", "Almoxarifado"],
    "Obs": ["Ativo", "Ativo", "Obsoleto", "Ativo", "Revisado", "Lâmpada nova", "Sem uso"]
}

os.makedirs("entrada", exist_ok=True)

df_gemat = pd.DataFrame(dados_gemat)
df_sam = pd.DataFrame(dados_sam)

df_gemat.to_excel("entrada/GEMAT.xlsx", index=False)
df_sam.to_excel("entrada/SAM.xlsx", index=False)

print("Arquivos de teste GEMAT.xlsx e SAM.xlsx criados com sucesso na pasta 'entrada'!")
