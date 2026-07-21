import pandas as pd
import re
from decimal import Decimal, ROUND_HALF_UP

print("Iniciando validação...\n")

# =========================
# NORMALIZAR TEXTO
# =========================
def norm(txt):
    if pd.isna(txt):
        return ""
    return str(txt).strip().lower() #transforma em texto,remove espaços no começo e no fim, deixa tudo minúsculo

# =========================
# TRATAR NÚMERO (CORRIGE VÍRGULA)
# =========================
def tratar_numero(valor):
    if pd.isna(valor):
        return None

    # Se já é numérico (int/float nativo do Excel via pandas), retorna direto.
    if isinstance(valor, (int, float)):
        return float(valor)

    valor = str(valor).strip()
    if valor in ["", "-", "nan"]:
        return None

    # Só troca separador BR (1.234,56 -> 1234.56) se houver vírgula.
    if "," in valor:
        valor = valor.replace(".", "").replace(",", ".")

    try:
        return float(valor)
    except:
        return None

# =========================
# DETECTAR DISCIPLINA (apenas informativo)
# =========================
def detectar_disciplina(nome_aba):
    nome_aba = nome_aba.lower()
    if nome_aba.startswith("d_1"):
        return "língua portuguesa"
    elif nome_aba.startswith("d_2"):
        return "matemática"
    elif nome_aba.startswith("d_29"):
        return "ciências da natureza"
    elif nome_aba.startswith("d_30"):
        return "ciências humanas"
    else:
        raise ValueError("Não foi possível identificar a disciplina pela aba!")

# =========================
# EXTRAIR FAIXAS
# =========================
def extrair_faixas(texto):
    partes = str(texto).split(";")
    regras = []

    for p in partes:
        p = p.strip()
        nome = p.split("(")[0].strip().lower()
        nums = list(map(float, re.findall(r"\d+\.?\d*", p)))

        if "até" in p.lower() and len(nums) == 1:
            min_v = 0
            max_v = nums[0]
        elif "ou mais" in p.lower() or "acima" in p.lower():
            min_v = nums[0]
            max_v = float("inf")
        else:
            if len(nums) >= 2:
                min_v = nums[0]
                max_v = nums[1]
            else:
                continue

        regras.append({
            "min": min_v,
            "max": max_v,
            "classificacao": norm(nome)
        })

    return regras

# =========================
# BUSCAR REGRAS
# Critério: medida_referencia + etapa + disciplina
# Fallback: medida_referencia + "todas" + "todas"
# =========================
def buscar_regras(medida_ref, etapa, disciplina, df_faixas):
    medida_norm     = norm(medida_ref)
    etapa_norm      = norm(etapa)
    disciplina_norm = norm(disciplina)

    filtro = df_faixas[
        (df_faixas["medida_ref_norm"] == medida_norm) &
        (df_faixas["etapa_norm"] == etapa_norm) &
        (df_faixas["disciplina_norm"] == disciplina_norm)
    ]

    if filtro.empty:
        filtro = df_faixas[
            (df_faixas["medida_ref_norm"] == medida_norm) &
            (df_faixas["etapa_norm"] == "todas") &
            (df_faixas["disciplina_norm"] == "todas")
        ]

    if filtro.empty:
        raise ValueError(
            f"Sem faixa para medida_referencia: '{medida_ref}' | etapa: '{etapa}' | disciplina: '{disciplina}'"
        )

    texto_faixa = filtro.iloc[0]["faixas"]
    return extrair_faixas(texto_faixa)

# =========================
# CLASSIFICAR
# =========================
def classificar(valor, regras):
    valor = int(Decimal(str(valor)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))  # >=.5 sobe, <.5 desce
    for r in regras:
        if r["min"] <= valor <= r["max"]:
            return r["classificacao"]
    return "fora_da_faixa"

# =========================
# INPUT
# =========================
arquivo = input("Digite o caminho do Excel: ").strip().strip('"') #salva o caminho do arquivo, remove espaços no começo e final, remove aspas duplas 

xls = pd.ExcelFile(arquivo) # usa o pandas para abrir o arquivo, xls vira um objeto que contém: nome das abas e estrutura do Excel

print("\nAbas disponíveis:")
for aba in xls.sheet_names: #lista com nomes das abas
    print("-", aba)

#Salva o nome das abas que contem a informação das faixas de corte e dos dados a serem validados
aba_faixas = input("\nDigite a aba de faixas (ex: G_1): ").strip() 
aba_dados  = input("Digite a aba de dados (ex: D_1): ").strip()

# =========================
# LEITURA
# =========================

#utiliza da função do canvas para ler as abas, converte tudo em uma tabela (dataframe)
df_faixas = pd.read_excel(arquivo, sheet_name=aba_faixas)
df        = pd.read_excel(arquivo, sheet_name=aba_dados)

# =========================
# NORMALIZAR(PADRONIZAR) COLUNAS
# =========================

#pega os nomes das colunas, remove espaços, transforma em minúsculo
df.columns        = df.columns.str.strip().str.lower()
df_faixas.columns = df_faixas.columns.str.strip().str.lower()

print("\nColunas encontradas (aba de dados):")
for c in df.columns:
    print("-", c)

print("\nColunas encontradas (aba de faixas):")
for c in df_faixas.columns:
    print("-", c)

# =========================
# DETECTAR COLUNAS
# =========================
def detectar_coluna(df, possiveis, nome_logico):
    for col in possiveis:
        if col in df.columns:
            print(f"{nome_logico}: '{col}'")
            return col
    raise ValueError(f"Coluna '{nome_logico}' não encontrada!\nDisponíveis: {df.columns.tolist()}")

def detectar_coluna_opcional(df, possiveis, nome_logico):
    for col in possiveis:
        if col in df.columns:
            print(f"{nome_logico} (opcional): '{col}'")
            return col
    print(f"{nome_logico} (opcional): não encontrada")
    return None

# --- Aba de faixas: coluna de medida de referência (obrigatória) ---
if "medida de referência" in df_faixas.columns:
    col_faixas_medida = "medida de referência"
elif "medida de referencia" in df_faixas.columns:
    col_faixas_medida = "medida de referencia"
else:
    raise ValueError(f"Coluna 'medida de referência' não encontrada na aba de faixas!\nDisponíveis: {df_faixas.columns.tolist()}")

if "etapa" not in df_faixas.columns:
    raise ValueError("Coluna 'etapa' não encontrada na aba de faixas!")
if "disciplina" not in df_faixas.columns:
    raise ValueError("Coluna 'disciplina' não encontrada na aba de faixas!")
if "faixas" not in df_faixas.columns:
    raise ValueError("Coluna 'faixas' não encontrada na aba de faixas!")

df_faixas["medida_ref_norm"] = df_faixas[col_faixas_medida].apply(norm)  #pega cada linha da coluna e roda a função norm nela
df_faixas["etapa_norm"]      = df_faixas["etapa"].apply(norm)
df_faixas["disciplina_norm"] = df_faixas["disciplina"].apply(norm)
"""

# --- Aba de dados: colunas do fluxo TRI (Proficiência) ---
col_prof  = detectar_coluna(df, ["proficiência - professor", "proficiência"], "proficiência (TRI)") #procura na tabela a coluna referente a medida
col_nivel = detectar_coluna(df, ["níveis de aprendizagem", "padrão de desempenho", "padrões de desempenho"], "classificação (TRI)")

# se não existir a coluna etapa e disciplina, dá erro e para tudo
if "etapa" not in df.columns:
    raise ValueError("Coluna 'etapa' não encontrada!")
if "disciplina" not in df.columns:
    raise ValueError("Coluna 'disciplina' não encontrada!")

# --- Aba de dados: colunas do fluxo TCT (Acertos no Teste %) --- Ele tenta descobrir se existem essas duas colunas no seu Excel
col_acertos_pct = detectar_coluna_opcional(df, ["acertos no teste %"], "acertos no teste % (TCT)")
col_categoria   = detectar_coluna_opcional(df, ["categoria de desempenho"], "categoria de desempenho (TCT)")

#só ativa o TCT se as DUAS colunas existirem
valida_tct = col_acertos_pct is not None and col_categoria is not None
if valida_tct:
    print("\n✅ Validação TCT (Acertos no Teste % x Categoria de Desempenho) ATIVADA.")
else:
    print("\nℹ️  Validação TCT não será feita (colunas não encontradas nesta aba).")


"""
# =========================
# DETECTAR COLUNAS TRI
# =========================
col_prof  = detectar_coluna_opcional(
    df,
    ["proficiência - professor", "proficiência"],
    "proficiência (TRI)"
)

col_nivel = detectar_coluna_opcional(
    df,
    ["níveis de aprendizagem", "padrão de desempenho", "padrões de desempenho"],
    "classificação (TRI)"
) #Ele tenta descobrir se existem essas duas colunas no seu Excel

# verifica se TRI pode ser usado
valida_tri = col_prof is not None and col_nivel is not None

if valida_tri:
    print("\n✅ Validação TRI (Proficiência x Classificação) ATIVADA.")
else:
    print("\nℹ️ Validação TRI não será feita (colunas não encontradas).")


# =========================
# COLUNAS OBRIGATÓRIAS GERAIS
# =========================
if "etapa" not in df.columns:
    raise ValueError("Coluna 'etapa' não encontrada!")

if "disciplina" not in df.columns:
    raise ValueError("Coluna 'disciplina' não encontrada!")


# =========================
# DETECTAR COLUNAS TCT (JÁ ERA OPCIONAL)
# =========================
col_acertos_pct = detectar_coluna_opcional(
    df,
    ["acertos no teste %"],
    "acertos no teste % (TCT)"
)

col_categoria = detectar_coluna_opcional(
    df,
    ["categoria de desempenho"],
    "categoria de desempenho (TCT)"
)

# verifica se TCT pode ser usado
valida_tct = col_acertos_pct is not None and col_categoria is not None

if valida_tct:
    print("\n✅ Validação TCT (Acertos no Teste % x Categoria de Desempenho) ATIVADA.")
else:
    print("\nℹ️ Validação TCT não será feita (colunas não encontradas).")


# =========================
# REGRA FINAL (PELO MENOS UM TEM QUE EXISTIR)
# =========================
if not valida_tri and not valida_tct:
    raise ValueError(
        "Nenhum dado válido encontrado! É necessário ter TRI ou TCT na planilha."
    )

# =========================
# DISCIPLINA (apenas informativo)
# =========================
disciplina_aba = detectar_disciplina(aba_dados)
print(f"\nDisciplina detectada pela aba: {disciplina_aba}")

# =========================
# CACHE DE REGRAS — chave (medida_ref, etapa, disciplina)
# =========================
cache_regras = {}

MEDIDA_PROFICIENCIA  = "proficiência"
MEDIDA_ACERTOS_TESTE = "porcentagem de acerto no teste"

def carregar_regra(medida_val, etapa_val, disciplina_val):
    chave = (norm(medida_val), norm(etapa_val), norm(disciplina_val))
    if chave not in cache_regras:
        try:
            cache_regras[chave] = buscar_regras(medida_val, etapa_val, disciplina_val, df_faixas)
        except Exception as e:
            print(f"⚠️  Sem regra para medida='{medida_val}' | etapa='{etapa_val}' | disciplina='{disciplina_val}': {e}")
    return chave

# Pré-carrega regras TRI (Proficiência) para todas as combinações etapa/disciplina existentes
combinacoes = df[["etapa", "disciplina"]].dropna().drop_duplicates()
for _, row in combinacoes.iterrows():
    carregar_regra(MEDIDA_PROFICIENCIA, row["etapa"], row["disciplina"])

# Pré-carrega a regra genérica TCT (Acertos no Teste, Todas/Todas), se aplicável
if valida_tct:
    carregar_regra(MEDIDA_ACERTOS_TESTE, "todas", "todas")

# =========================
# MOSTRAR REGRAS INTERPRETADAS
# =========================
print("\nREGRAS INTERPRETADAS:\n")

for (medida_k, etapa_k, disc_k), regras in cache_regras.items():
    print(f"\nMedida: {medida_k} | Etapa: {etapa_k} | Disciplina: {disc_k}")
    for r in regras:
        max_v = "∞" if r["max"] == float("inf") else r["max"]
        print(f"  - {r['classificacao']} → {r['min']} até {max_v}")

# =========================
# VALIDAÇÃO — DUAS VALIDAÇÕES INDEPENDENTES POR LINHA
# =========================
df["calc_tri"]       = ""
df["validacao_tri"]  = ""
df["calc_tct"]       = ""
df["validacao_tct"]  = ""

for i, row in df.iterrows():
    etapa_norm      = norm(row["etapa"])
    disciplina_norm = norm(row["disciplina"])
    linha_excel     = i + 2  # +2: índice começa em 0 e Excel tem cabeçalho na linha 1

    # ---------- VALIDAÇÃO TRI (Proficiência x Níveis de Aprendizagem) ----------
    try:
        valor_tri = tratar_numero(row[col_prof])

        if valor_tri is None:
            df.at[i, "validacao_tri"] = "valor inválido"
        else:
            planilha_tri = norm(row[col_nivel])
            chave_tri    = (MEDIDA_PROFICIENCIA, etapa_norm, disciplina_norm)
            regras_tri   = cache_regras.get(chave_tri)

            if not regras_tri:
                df.at[i, "validacao_tri"] = "sem regra"
            else:
                calc_tri = classificar(valor_tri, regras_tri)
                df.at[i, "calc_tri"] = calc_tri

                if calc_tri == planilha_tri:
                    df.at[i, "validacao_tri"] = "ok"
                else:
                    df.at[i, "validacao_tri"] = f"calc: {calc_tri} | planilha: {planilha_tri}"
                    print(f"\n❌ ERRO TRI — Linha Excel {linha_excel} (índice {i})")
                    print(f"  Etapa:       {row['etapa']}")
                    print(f"  Disciplina:  {row['disciplina']}")
                    print(f"  Profic.:     {valor_tri}")
                    print(f"  Planilha:    {planilha_tri}")
                    print(f"  Calculado:   {calc_tri}")
                    print(f"  Regras:      {regras_tri}")

    except Exception as e:
        df.at[i, "validacao_tri"] = f"erro: {str(e)}"

    # ---------- VALIDAÇÃO TCT (Acertos no Teste % x Categoria de Desempenho) ----------
    if valida_tct:
        try:
            valor_tct = tratar_numero(row[col_acertos_pct])

            if valor_tct is None:
                df.at[i, "validacao_tct"] = "valor inválido"
            else:
                planilha_tct = norm(row[col_categoria])
                chave_tct    = (MEDIDA_ACERTOS_TESTE, "todas", "todas")
                regras_tct   = cache_regras.get(chave_tct)

                if not regras_tct:
                    df.at[i, "validacao_tct"] = "sem regra"
                else:
                    calc_tct = classificar(valor_tct, regras_tct)
                    df.at[i, "calc_tct"] = calc_tct

                    if calc_tct == planilha_tct:
                        df.at[i, "validacao_tct"] = "ok"
                    else:
                        df.at[i, "validacao_tct"] = f"calc: {calc_tct} | planilha: {planilha_tct}"
                        print(f"\n❌ ERRO TCT — Linha Excel {linha_excel} (índice {i})")
                        print(f"  Etapa:       {row['etapa']}")
                        print(f"  Disciplina:  {row['disciplina']}")
                        print(f"  Acertos %:   {valor_tct}")
                        print(f"  Planilha:    {planilha_tct}")
                        print(f"  Calculado:   {calc_tct}")
                        print(f"  Regras:      {regras_tct}")

        except Exception as e:
            df.at[i, "validacao_tct"] = f"erro: {str(e)}"

# =========================
# RESUMO FINAL
# =========================
total = len(df)

corretas_tri  = (df["validacao_tri"] == "ok").sum()
erradas_tri   = df["validacao_tri"].str.contains("calc:", na=False).sum()
invalidas_tri = df["validacao_tri"].str.contains("inválido|erro|sem regra", na=False).sum()

print("\n==============================")
print("📊 RESUMO FINAL — TRI (Proficiência)")
print("==============================")
print(f"Total de linhas : {total}")
print(f"✔  Corretas     : {corretas_tri}")
print(f"❌  Erradas      : {erradas_tri}")
print(f"⚠️  Inválidas    : {invalidas_tri}")

if valida_tct:
    corretas_tct  = (df["validacao_tct"] == "ok").sum()
    erradas_tct   = df["validacao_tct"].str.contains("calc:", na=False).sum()
    invalidas_tct = df["validacao_tct"].str.contains("inválido|erro|sem regra", na=False).sum()

    print("\n==============================")
    print("📊 RESUMO FINAL — TCT (Acertos no Teste %)")
    print("==============================")
    print(f"Total de linhas : {total}")
    print(f"✔  Corretas     : {corretas_tct}")
    print(f"❌  Erradas      : {erradas_tct}")
    print(f"⚠️  Inválidas    : {invalidas_tct}")

# =========================
# EXPORTAR
# =========================
saida = "resultado_validado.xlsx"
df.to_excel(saida, index=False)
print(f"\nFinalizado! Arquivo gerado: {saida}")