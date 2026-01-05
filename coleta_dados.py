#%% === Importar Bibliotecas ===
import pdfplumber
import pandas as pd
import re
import os

#%% === Definir funções e caminho do arquivo ===
pasta_pdfs = "CTEs Todos Novembro"  # nome da pasta onde estão seus PDFs
arquivos_pdfs = [f for f in os.listdir(pasta_pdfs) if f.lower().endswith(".pdf")]

# Limpa e padroniza texto
def normalize_line(s):
    return re.sub(r'\s+', ' ', s).strip()

# Localiza informações no texto
def find_line_index_containing(lines, patterns):
    for i, ln in enumerate(lines):
        for p in patterns:
            if p.lower() in ln.lower():
                return i
    return -1

# Captura texto
def get_following_text(lines, start_idx, stop_tokens=None, max_lines=3):
    stop_tokens = stop_tokens or []
    collected = []
    for j in range(start_idx + 1, min(start_idx + 1 + max_lines, len(lines))):
        ln = lines[j].strip()
        if not ln:
            continue
        low = ln.lower()
        if any(tok.lower() in low for tok in stop_tokens):
            break
        collected.append(ln)
    return " ".join(collected).strip()

#%% === Loop para ler vários PDFs ===
todos_dados = []
Contador = 0

for arquivo_pdf in arquivos_pdfs:
    caminho_pdf = os.path.join(pasta_pdfs, arquivo_pdf)
    Contador += 1
    print(f"🔍 {Contador} - Processando: {arquivo_pdf}")
    # === Leitura do PDF ===
    #with pdfplumber.open(caminho_pdf) as pdf:
    #    texto_pages = []
    #    for p in pdf.pages:
    #        t = p.extract_text()
    #        if t:
    #            texto_pages.append(t)
                
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto_pages = []
            for p in pdf.pages:
                t = p.extract_text()
                if t:
                    texto_pages.append(t)
    except Exception as e:
        print(f"❌ Erro ao abrir {arquivo_pdf}: {e}")
        # Salva o erro mas continua o loop
        dados = {
            "ARQUIVO": arquivo_pdf,
            "DTA. HRA. DE EMISSÃO": "",
            "INÍCIO DA PRESTAÇÃO": "",
            "TÉRMINO DA PRESTAÇÃO": "",
            "QTD": "",
            "VALOR TOTAL DA MERCADORIA": "",
            "FRETE": "",
            "NÚMERO DOCUMENTO": "",
            "MOTORISTA": "",
            "PLACAS": "",
            "EMISSOR": "",
            "ERRO": "PDF CORROMPIDO OU INVÁLIDO"
        }
        todos_dados.append(dados)
        continue

    texto_raw = "\n".join(texto_pages)

    lines = [ln for ln in [l.strip() for l in texto_raw.splitlines()] if ln]
    texto_norm = re.sub(r'\s+', ' ', texto_raw)

    # === Campos desejados (mantém igual ao seu código) ===
    dados = {
        "ARQUIVO": arquivo_pdf,
        "DTA. HRA. DE EMISSÃO": "",
        "INÍCIO DA PRESTAÇÃO": "",
        "TÉRMINO DA PRESTAÇÃO": "",
        "QTD": "",
        "VALOR TOTAL DA MERCADORIA": "",
        "FRETE": "",
        "NÚMERO DOCUMENTO": "",
        "MOTORISTA": "",
        "PLACAS": "",
        "EMISSOR": ""
    }
    
    # ---------- 1) DTA. HRA. DE EMISSÃO ----------
    m = re.search(r'(\d{2}/\d{2}/\d{2}\s*\d{2}:\d{2}:\d{2})', texto_norm)
    if m:
        dados["DTA. HRA. DE EMISSÃO"] = m.group(1)
    
    # ---------- 2) INÍCIO / TÉRMINO DA PRESTAÇÃO ----------
    idx_head = find_line_index_containing(lines, ["INÍCIO DA PRESTAÇÃO", "INICIO DA PRESTACAO"])
    if idx_head != -1:
        next_line = ""
        for j in range(idx_head + 1, min(idx_head + 4, len(lines))):
            if lines[j].strip():
                next_line = lines[j].strip()
                break
        if next_line:
            parts = re.split(r'\s{2,}', next_line)
            if len(parts) >= 2:
                dados["INÍCIO DA PRESTAÇÃO"] = normalize_line(parts[0])
                dados["TÉRMINO DA PRESTAÇÃO"] = normalize_line(parts[1])
            else:
                matches = list(re.finditer(r'\b[A-Z]{2}\s*-\s*', next_line))
                if len(matches) >= 2:
                    second_start = matches[1].start()
                    dados["INÍCIO DA PRESTAÇÃO"] = normalize_line(next_line[:second_start])
                    dados["TÉRMINO DA PRESTAÇÃO"] = normalize_line(next_line[second_start:])
                else:
                    tokens = next_line.split()
                    half = len(tokens) // 2
                    dados["INÍCIO DA PRESTAÇÃO"] = normalize_line(" ".join(tokens[:half]))
                    dados["TÉRMINO DA PRESTAÇÃO"] = normalize_line(" ".join(tokens[half:]))
    
    # ---------- 3) QTD ----------
    m = re.search(r'\bQTD[:\s]*([\d\.,]+)', texto_norm, re.IGNORECASE)
    if not m:
        m = re.search(r'\bKG[:\s]*([\d\.,]+)', texto_norm, re.IGNORECASE)
    if m:
        dados["QTD"] = m.group(1)
    
    # ---------- 4) VALOR TOTAL DA MERCADORIA ----------
    m = re.search(r'VALOR\s*TOTAL\s*DA\s*MERCADORIA[:\s]*([\d\.,]+)', texto_norm, re.IGNORECASE)
    if not m:
        idx_label = find_line_index_containing(lines, ["VALOR TOTAL DA MERCADORIA", "TOTAL MERCADORIA"])
        if idx_label != -1 and idx_label + 1 < len(lines):
            cand = re.search(r'([\d\.,]+)', lines[idx_label + 1])
            if cand:
                dados["VALOR TOTAL DA MERCADORIA"] = cand.group(1)
    else:
        dados["VALOR TOTAL DA MERCADORIA"] = m.group(1)
    
    # ---------- 5) FRETE ----------
    m = re.search(r'\bFRETE[:\s]*([\d\.,]+)', texto_norm, re.IGNORECASE)
    if m:
        dados["FRETE"] = m.group(1)
    
    # ---------- 6) NÚMERO DOCUMENTO ----------
    idx_doc = find_line_index_containing(lines, ["NÚMERO DOCUMENTO", "NUMERO DOCUMENTO"])
    if idx_doc != -1 and idx_doc + 1 < len(lines):
        linha_doc = lines[idx_doc + 1].strip()
        m = re.search(r'\b(\d{4,})\b', linha_doc)
        if m:
            dados["NÚMERO DOCUMENTO"] = m.group(1)
    
    # ---------- 7) MOTORISTA / PLACAS / EMISSOR ----------
    m_obs = re.search(r'OBSERVAÇÕES GERAIS(.*)', texto_raw, re.IGNORECASE | re.DOTALL)
    if m_obs:
        bloco = normalize_line(m_obs.group(1))
    
        # === MOTORISTA ===
        # Tenta capturar casos normais (com "MOTORISTA")
        m_mot = re.search(
            r'MOTORISTA[\s:/\-]+([A-ZÀ-ÿ\s\.]+?)(?=\s*[-/]*\s*(PLACA|EMISSOR|NF|$))',
            bloco,
            re.IGNORECASE
        )
    
        # Se não encontrar, tenta capturar nome antes de "PLACA" ou "EMISSOR"
        if not m_mot:
            m_mot = re.search(
                r'NF\s*\d+\s*[/\-:\s]*([A-ZÀ-ÿ\s\.]+?)(?=\s*(PLACA|EMISSOR|NF|$))',
                bloco,
                re.IGNORECASE
            )
    
        if m_mot:
            dados["MOTORISTA"] = m_mot.group(1).strip()
    
        # === PLACAS ===
        # Captura placas com ou sem a palavra "PLACA", aceitando ":" e variantes
        placas_encontradas = re.findall(
            r'(?:PLACA[S]?(?:\s*(?:CAVALO|CARRETA)?)[\s:]*|[/\s])([A-Z0-9/]{5,})',
            bloco,
            re.IGNORECASE
        )
    
        if placas_encontradas:
            # Filtra falsos positivos (como NF, ICMS, etc.)
            placas_limpas = []
            for p in placas_encontradas:
                p = p.strip().upper()
                # Normaliza placas separadas por "/"
                subplacas = [sp.strip() for sp in p.split('/') if sp.strip()]
                for sp in subplacas:
                    # Mantém apenas padrões válidos de placas (AAA0A00, ABC1234, etc.)
                    if re.match(r'^[A-Z]{3,4}\d[A-Z0-9]{2,3}$', sp):
                        placas_limpas.append(sp)
            if placas_limpas:
                dados["PLACAS"] = "/".join(placas_limpas)

        # === EMISSOR ===
        m_emissor = re.search(
            r'EMISSOR[\s:/\-]+([A-ZÀ-ÿ\s\.]+?)(?=\s+(?:INFORMA|USO|DATA|$))',
            bloco,
            re.IGNORECASE
        )
        if m_emissor:
            dados["EMISSOR"] = m_emissor.group(1).strip()
                
    # === Salva o resultado de cada PDF ===
    todos_dados.append(dados)

#%% === Exporta tudo para o Excel ===
df = pd.DataFrame(todos_dados)
df.to_excel("dados_dacte_todos22.xlsx", index=False)
print("\n✅ Extração concluída. Dados salvos em 'dados_dacte_todos.xlsx'")
