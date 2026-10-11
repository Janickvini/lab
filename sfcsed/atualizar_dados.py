import openpyxl
import json
import os
import sys
from datetime import datetime

base_dir = os.path.dirname(os.path.abspath(__file__))

# 1. Localizar arquivo Excel
excel_files = [f for f in os.listdir(base_dir) if f.endswith('.xlsx') and not f.startswith('~$')]
if not excel_files:
    print("Erro: Nenhum arquivo .xlsx encontrado na pasta sfcsed.")
    sys.exit(1)

# Prioriza "SF CSED.xlsx" se existir
if "SF CSED.xlsx" in excel_files:
    target_excel = "SF CSED.xlsx"
else:
    target_excel = sorted(excel_files)[-1]

wb_path = os.path.join(base_dir, target_excel)
print(f"Lendo dados de: {wb_path}...")

try:
    wb = openpyxl.load_workbook(wb_path, data_only=True)
except Exception as e:
    print(f"Erro ao abrir arquivo Excel: {e}")
    sys.exit(1)

# Validação das abas obrigatórias
required_sheets = ["SF (Scopus)", "Scopus 3 (Limpo)"]
for sheet_name in required_sheets:
    if sheet_name not in wb.sheetnames:
        print(f"Erro: Aba obrigatória '{sheet_name}' não encontrada na planilha. Abas disponíveis: {wb.sheetnames}")
        sys.exit(1)

def clean_value(val):
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        return val
    s = str(val).strip()
    return s

# 1. Processar Artigos Scopus (Scopus 3 (Limpo))
ws_scopus = wb["Scopus 3 (Limpo)"]
scopus_rows = list(ws_scopus.iter_rows(values_only=True))
sc_headers = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(scopus_rows[0])]

articles_dict = {}
articles_list = []

for row in scopus_rows[1:]:
    if not any(row):
        continue
    item = {}
    for idx, val in enumerate(row):
        if idx < len(sc_headers):
            item[sc_headers[idx]] = clean_value(val)
    
    art_id = item.get("ID")
    if art_id != "" and art_id is not None:
        try:
            art_id_int = int(art_id)
        except:
            art_id_int = art_id
        
        doi_link = item.get("DOILINK", "")
        doi_val = item.get("DOI", "")
        if not doi_link and doi_val:
            doi_link = f"https://doi.org/{doi_val}" if not str(doi_val).startswith("http") else str(doi_val)
            
        clean_art = {
            "id": art_id_int,
            "authors": item.get("Authors", "Autores não informados"),
            "title": item.get("Title", "Título não informado"),
            "year": item.get("Year", ""),
            "sourceTitle": item.get("Source title", ""),
            "doi": doi_val,
            "doiLink": doi_link,
            "abstract": item.get("Abstract", ""),
            "authorKeywords": item.get("Author Keywords", ""),
            "indexKeywords": item.get("Index Keywords", ""),
            "documentType": item.get("Document Type", ""),
            "associatedSFs": []
        }
        articles_dict[art_id_int] = clean_art
        articles_list.append(clean_art)

# 2. Processar Sementes de Futuro (SF (Scopus))
ws_sf = wb["SF (Scopus)"]
sf_rows = list(ws_sf.iter_rows(values_only=True))
sf_headers = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(sf_rows[0])]

sementes_list = []

for row in sf_rows[1:]:
    if not any(row):
        continue
    item = {}
    for idx, val in enumerate(row):
        if idx < len(sf_headers):
            item[sf_headers[idx]] = clean_value(val)
    
    sf_id = item.get("ID da SF", "")
    if not sf_id:
        continue

    # Extrair IDs numéricos de artigos associados
    raw_art_ids = str(item.get("Artigos Scopus Associados (IDs)", ""))
    parsed_art_ids = []
    for p in raw_art_ids.replace(";", ",").split(","):
        p = p.strip()
        if p.isdigit():
            art_num = int(p)
            parsed_art_ids.append(art_num)
            # Vincular a SF de volta ao artigo
            if art_num in articles_dict and sf_id not in articles_dict[art_num]["associatedSFs"]:
                articles_dict[art_num]["associatedSFs"].append(sf_id)

    # Processar PESTEL (lista limpa de dimensões)
    raw_pestel = str(item.get("Dimensões PESTEL", ""))
    pestel_list = [dim.strip() for dim in raw_pestel.replace(";", ",").split(",") if dim.strip()]

    # Tratar campo revisada
    raw_revisada = item.get("Revisada", 0)
    is_revisada = (raw_revisada == 1 or str(raw_revisada).strip().lower() in ["1", "sim", "true", "s"])

    sf_obj = {
        "id": sf_id,
        "topico": item.get("Tópico Mais Amplo Padronizado", "Geral"),
        "titulo": item.get("Título da Semente de Futuro", "Sem título"),
        "descricao": item.get("Descrição Global do Fenômeno (com Citações Numéricas)", ""),
        "classificacao": item.get("Classificação Prospectiva", "Não classificada"),
        "pestel": pestel_list,
        "eixo": item.get("Eixo", "Não categorizado"),
        "qtdArtigos": len(parsed_art_ids) if parsed_art_ids else (item.get("Qtd Artigos") or 0),
        "artigosIds": parsed_art_ids,
        "referenciasResumidas": item.get("Referências Bibliográficas Resumidas (Autores, Ano, Título)", ""),
        "implicacoes": item.get("Conexão / Implicações Estratégicas para o Ensino Superior Privado até 2040", ""),
        "revisada": is_revisada
    }
    sementes_list.append(sf_obj)

# Ordenar SFs numericamente (ex: SF01, SF02... SF62)
def sf_sort_key(sf):
    code = sf["id"]
    digits = "".join([c for c in code if c.isdigit()])
    return int(digits) if digits else 9999

sementes_list.sort(key=sf_sort_key)

# Consolidar dados globais e estatísticas
data_payload = {
    "metadata": {
        "sourceFile": target_excel,
        "totalSF": len(sementes_list),
        "totalArticles": len(articles_list),
        "totalArticlesCited": len([a for a in articles_list if len(a["associatedSFs"]) > 0]),
        "totalRevisadas": len([sf for sf in sementes_list if sf["revisada"]]),
        "generatedAt": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    },
    "sementes": sementes_list,
    "artigos": articles_list
}

# Salvar como data.js
data_js_path = os.path.join(base_dir, "data.js")
with open(data_js_path, "w", encoding="utf-8") as f:
    f.write("// Dados exportados automaticamente de " + target_excel + "\n")
    f.write("const sfData = ")
    json.dump(data_payload, f, ensure_ascii=False, indent=2)
    f.write(";\n")

print(f"\n--- Sincronização Concluída com Sucesso ---")
print(f"Arquivo lido: {target_excel}")
print(f"Sementes de Futuro extraídas: {len(sementes_list)} (Revisadas: {data_payload['metadata']['totalRevisadas']})")
print(f"Artigos Scopus mapeados: {len(articles_list)} (Artigos citados: {data_payload['metadata']['totalArticlesCited']})")
print(f"Arquivo gerado: {data_js_path}")
