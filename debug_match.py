import pandas as pd
import matcher
import normalizador

df_g = pd.read_excel('entrada/GEMAT.xlsx')
df_s = pd.read_excel('entrada/SAM.xlsx')
config = matcher.carregar_configuracao('config.json')
sinonimos = normalizador.carregar_sinonimos('sinonimos.json')

col_g = {'patrimonio':'Plaqueta/Patrimônio', 'descricao':'Descrição do Bem', 'local':'Setor/Local', 'observacao':'Observações Gerais'}
col_s = {'patrimonio':'Registro/Código', 'descricao':'Especificação do Material', 'local':'Lotação Atual', 'observacao':'Obs'}

g_row = df_g.iloc[0]
desc_g = g_row['Descrição do Bem']
desc_g_norm = normalizador.normalizar_texto(desc_g, sinonimos)
caract_g = normalizador.extrair_caracteristicas(desc_g_norm)

print("GEMAT Item 1001:")
print("  Original:", desc_g)
print("  Normalizado:", desc_g_norm)
print("  Caract:", caract_g)

print("\nSAM Items:")
for i, s_row in df_s.iterrows():
    desc_s = s_row['Especificação do Material']
    desc_s_norm = normalizador.normalizar_texto(desc_s, sinonimos)
    caract_s = normalizador.extrair_caracteristicas(desc_s_norm)
    score = matcher.calcular_confianca(caract_g, caract_s, desc_g_norm, desc_s_norm, config)
    print(f"  SAM {s_row['Registro/Código']}:")
    print(f"    Original: {desc_s}")
    print(f"    Normalizado: {desc_s_norm}")
    print(f"    Caract: {caract_s}")
    print(f"    Score: {score}")
