import pandas as pd
import json
from datetime import datetime

caminho_arquivo = "Dados/Ferias SECAP.xlsx"

# 1. Carregar ajustando o cabeçalho para a linha correta
df = pd.read_excel(caminho_arquivo, header=1)

pessoas = []
unidade_atual = ""
pessoa_atual = None
periodo_atual = ""

# Função auxiliar para converter datas de forma segura sem estourar NaTType error
def formatar_data_segura(val):
    if pd.isna(val) or val is None:
        return None
    dt = pd.to_datetime(val, errors='coerce', dayfirst=True)
    if pd.notna(dt):
        return dt.strftime('%Y-%m-%d')
    return None

# 2. Iterar sobre as linhas da planilha replicando a lógica do Código.gs
for idx, row in df.iterrows():
    uni = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
    nome = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
    cargo = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
    periodo = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
    
    # Preenchimento de células mescladas (Unidade)
    if uni and uni.lower() != "nan":
        unidade_atual = uni
        
    # Quando encontra um novo Servidor
    if nome and nome.lower() != "nan":
        pessoa_atual = {
            "unidade": unidade_atual,
            "nome": nome,
            "cargo": cargo if cargo.lower() != "nan" else "",
            "ferias": []
        }
        pessoas.append(pessoa_atual)
        periodo_atual = ""
        
    if not pessoa_atual:
        continue
        
    if periodo and periodo.lower() != "nan":
        periodo_atual = periodo
        
    # Tratar datas de início e fim
    ini_val = row.iloc[4]
    fim_val = row.iloc[5]
    
    ini_str = formatar_data_segura(ini_val)
    fim_str = formatar_data_segura(fim_val)
    
    texto_combinado = f"{str(ini_val)} {str(fim_val)}".lower()
    
    if ini_str and fim_str:
        pessoa_atual["ferias"].append({
            "inicio": ini_str,
            "fim": fim_str,
            "periodo": periodo_atual,
            "definir": False
        })
    elif "definir" in texto_combinado:
        pessoa_atual["ferias"].append({
            "inicio": None,
            "fim": None,
            "periodo": periodo_atual,
            "definir": True
        })

# 3. Estrutura idêntica ao retorno do getData() no Apps Script
dados_finais = {
    "titulo": "Painel de Férias SECAP",
    "atualizadoEm": datetime.now().strftime('%d/%m/%Y %H:%M'),
    "pessoas": pessoas
}

# 4. Exportar para dados_dashboard.json
with open("dados_dashboard.json", "w", encoding="utf-8") as f:
    json.dump(dados_finais, f, ensure_ascii=False, indent=2)

print("✅ 'dados_dashboard.json' gerado com sucesso!")