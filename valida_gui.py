# -*- coding: utf-8 -*-
"""
ValidaCorte— Interface Gráfica
Valida Proficiência (TRI) e Acertos no Teste % (TCT) contra as faixas
de corte definidas na aba de faixas (ex: G_1).
"""

# =============================================================================
# IMPORTAÇÕES
# os         → manipular caminhos de arquivo (pasta, nome do arquivo, etc.)
# re         → expressões regulares (extrair números das faixas de texto)
# threading  → rodar a validação em segundo plano sem travar a interface
# traceback  → capturar o erro completo quando algo dá errado
# Decimal    → arredondamento preciso: 30.5 → 31, nunca 30 (evita bug do Python)
# pandas     → ler e manipular planilhas Excel
# tkinter    → biblioteca padrão do Python para criar interfaces gráficas
# =============================================================================
import os
import re
import threading
import traceback
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd

import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# =============================================================================
# BLOCO 1 — FUNÇÕES UTILITÁRIAS
# =============================================================================

def norm(txt):
    """
    Normaliza um texto: remove espaços das bordas e deixa tudo minúsculo.
    Usado para comparar etapa, disciplina, classificação sem se preocupar
    com maiúsculas/minúsculas ou espaços extras.
    Ex: "  Língua Portuguesa  " → "língua portuguesa"
    """
    if pd.isna(txt):
        return ""
    return str(txt).strip().lower()


def tratar_numero(valor):
    """
    Converte qualquer valor em float, lidando com dois cenários:
    1. Valor já veio como número nativo do pandas (ex: 49.5) → retorna direto.
       Sem essa verificação, str(49.5) → "49.5" → remove o ponto → "495" → ERRO.
    2. Valor veio como string no formato BR (ex: "49,50" ou "1.234,56")
       → troca vírgula por ponto decimal só se houver vírgula.
    Retorna None se o valor for vazio, traço ou inválido.
    """
    if pd.isna(valor):
        return None

    # Se já é número nativo do Excel via pandas, retorna direto (sem tratar string)
    if isinstance(valor, (int, float)):
        return float(valor)

    valor = str(valor).strip()
    if valor in ["", "-", "nan"]:
        return None

    # Só substitui separadores se houver vírgula (formato BR: 1.234,56 → 1234.56)
    if "," in valor:
        valor = valor.replace(".", "").replace(",", ".")

    try:
        return float(valor)
    except Exception:
        return None


def detectar_disciplina(nome_aba):
    """
    Descobre a disciplina pelo nome da aba de dados.
    Usado apenas para exibir no log — não afeta a validação.
    Ex: "D_1" → "língua portuguesa", "D_2" → "matemática"
    """
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
        return "não identificada"


# =============================================================================
# BLOCO 2 — LEITURA E INTERPRETAÇÃO DAS FAIXAS DE CORTE
# =============================================================================

def extrair_faixas(texto):
    """
    Converte o texto de faixas (como está na planilha G_1) em uma lista de regras.
    Exemplo de entrada:
      "Defasagem (Até 30); Aprendizado Intermediário (31 até 70); Aprendizado Adequado (71 até 100)"
    Resultado:
      [
        {"min": 0,  "max": 30.0,  "classificacao": "defasagem"},
        {"min": 31, "max": 70.0,  "classificacao": "aprendizado intermediário"},
        {"min": 71, "max": 100.0, "classificacao": "aprendizado adequado"},
      ]
    Reconhece três formatos:
      - "Até X"         → min=0, max=X
      - "X ou mais"     → min=X, max=infinito
      - "X até Y"       → min=X, max=Y
    """
    partes = str(texto).split(";")
    regras = []

    for p in partes:
        p = p.strip()
        # Pega o nome antes do parêntese: "Defasagem (Até 30)" → "defasagem"
        nome = p.split("(")[0].strip().lower()
        # Extrai todos os números do texto (ex: "31 até 70" → [31.0, 70.0])
        nums = list(map(float, re.findall(r"\d+\.?\d*", p)))

        if "até" in p.lower() and len(nums) == 1:
            # Formato "Até X" → começa em 0
            min_v = 0
            max_v = nums[0]
        elif "ou mais" in p.lower() or "acima" in p.lower():
            # Formato "X ou mais" → vai até infinito
            min_v = nums[0]
            max_v = float("inf")
        else:
            # Formato "X até Y"
            if len(nums) >= 2:
                min_v = nums[0]
                max_v = nums[1]
            else:
                continue  # ignora linha malformada

        regras.append({
            "min": min_v,
            "max": max_v,
            "classificacao": norm(nome)
        })

    return regras


def buscar_regras(medida_ref, etapa, disciplina, df_faixas):
    """
    Busca na aba de faixas (G_1) a regra correspondente à combinação:
    medida_ref + etapa + disciplina.

    Se não encontrar a combinação exata, tenta o fallback genérico:
    mesma medida de referência + etapa="todas" + disciplina="todas".

    Ex: "Proficiência" + "ENSINO MÉDIO - 2ª SÉRIE" + "Matemática"
    Se não existir essa linha no G_1, tenta "Proficiência" + "Todas" + "Todas".
    """
    medida_norm = norm(medida_ref)
    etapa_norm = norm(etapa)
    disciplina_norm = norm(disciplina)

    # Busca exata
    filtro = df_faixas[
        (df_faixas["medida_ref_norm"] == medida_norm) &
        (df_faixas["etapa_norm"] == etapa_norm) &
        (df_faixas["disciplina_norm"] == disciplina_norm)
    ]

    # Fallback: mesma medida, mas etapa/disciplina genéricas
    if filtro.empty:
        filtro = df_faixas[
            (df_faixas["medida_ref_norm"] == medida_norm) &
            (df_faixas["etapa_norm"] == "todas") &
            (df_faixas["disciplina_norm"] == "todas")
        ]

    if filtro.empty:
        raise ValueError(
            f"Sem faixa para medida_referencia: '{medida_ref}' | "
            f"etapa: '{etapa}' | disciplina: '{disciplina}'"
        )

    texto_faixa = filtro.iloc[0]["faixas"]
    return extrair_faixas(texto_faixa)


def classificar(valor, regras):
    """
    Arredonda o valor (>=0.5 sobe, <0.5 desce) e descobre em qual faixa ele se encaixa.
    Usa ROUND_HALF_UP do módulo Decimal porque o round() nativo do Python
    usa arredondamento bancário (30.5 → 30, não 31), o que daria classificações erradas.
    Retorna "fora_da_faixa" se o valor não se encaixar em nenhuma regra.
    """
    valor = int(Decimal(str(valor)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    for r in regras:
        if r["min"] <= valor <= r["max"]:
            return r["classificacao"]
    return "fora_da_faixa"


def detectar_coluna_opcional(df, possiveis, nome_logico, log):
    """
    Tenta encontrar uma coluna no dataframe a partir de uma lista de nomes possíveis.
    Diferente da versão obrigatória, não lança erro se não encontrar — apenas avisa no log.
    Retorna o nome real da coluna ou None.
    """
    for col in possiveis:
        if col in df.columns:
            log(f"{nome_logico}: '{col}'")
            return col
    log(f"{nome_logico}: não encontrada")
    return None


# Constantes com os nomes normalizados das medidas de referência usadas no G_1
MEDIDA_PROFICIENCIA = "proficiência"
MEDIDA_ACERTOS_TESTE = "porcentagem de acerto no teste"


# =============================================================================
# BLOCO 3 — FUNÇÃO PRINCIPAL DE VALIDAÇÃO
# Recebe callbacks de log e progresso para atualizar a interface em tempo real
# =============================================================================

def rodar_validacao(arquivo, aba_faixas, aba_dados, log, progress_callback):
    """
    Executa toda a lógica de validação.
    - log(str)                   → escreve uma linha no log da interface
    - progress_callback(0..100)  → atualiza a barra de progresso
    Retorna o caminho do arquivo resultado_validado.xlsx gerado.
    """

    log("Iniciando validação...\n")

    # Lê as duas abas do Excel em DataFrames
    df_faixas = pd.read_excel(arquivo, sheet_name=aba_faixas)
    df = pd.read_excel(arquivo, sheet_name=aba_dados)

    # Padroniza os nomes das colunas: remove espaços e deixa minúsculo
    df.columns = df.columns.str.strip().str.lower()
    df_faixas.columns = df_faixas.columns.str.strip().str.lower()

    log("Colunas encontradas (aba de dados):")
    for c in df.columns:
        log(f"- {c}")

    log("\nColunas encontradas (aba de faixas):")
    for c in df_faixas.columns:
        log(f"- {c}")

    # --- Valida que a aba de faixas tem a estrutura esperada ---
    # Aceita com ou sem acento na palavra "referência"
    if "medida de referência" in df_faixas.columns:
        col_faixas_medida = "medida de referência"
    elif "medida de referencia" in df_faixas.columns:
        col_faixas_medida = "medida de referencia"
    else:
        raise ValueError(
            f"Coluna 'medida de referência' não encontrada na aba de faixas!\n"
            f"Disponíveis: {df_faixas.columns.tolist()}"
        )

    if "etapa" not in df_faixas.columns:
        raise ValueError("Coluna 'etapa' não encontrada na aba de faixas!")
    if "disciplina" not in df_faixas.columns:
        raise ValueError("Coluna 'disciplina' não encontrada na aba de faixas!")
    if "faixas" not in df_faixas.columns:
        raise ValueError("Coluna 'faixas' não encontrada na aba de faixas!")

    # Pré-normaliza colunas do G_1 para não repetir esse trabalho dentro dos loops
    df_faixas["medida_ref_norm"] = df_faixas[col_faixas_medida].apply(norm)
    df_faixas["etapa_norm"] = df_faixas["etapa"].apply(norm)
    df_faixas["disciplina_norm"] = df_faixas["disciplina"].apply(norm)

    # --- Detecta colunas do fluxo TRI (ambas precisam existir para ativar) ---
    col_prof = detectar_coluna_opcional(
        df, ["proficiência - professor", "proficiência"], "proficiência (TRI)", log
    )
    col_nivel = detectar_coluna_opcional(
        df,
        ["níveis de aprendizagem", "padrão de desempenho", "padrões de desempenho"],
        "classificação (TRI)",
        log,
    )
    valida_tri = col_prof is not None and col_nivel is not None
    log("\n✅ Validação TRI ATIVADA." if valida_tri else "\nℹ️ Validação TRI não será feita.")

    # Colunas obrigatórias para qualquer validação
    if "etapa" not in df.columns:
        raise ValueError("Coluna 'etapa' não encontrada!")
    if "disciplina" not in df.columns:
        raise ValueError("Coluna 'disciplina' não encontrada!")

    # --- Detecta colunas do fluxo TCT (ambas precisam existir para ativar) ---
    col_acertos_pct = detectar_coluna_opcional(
        df, ["acertos no teste %"], "acertos no teste % (TCT)", log
    )
    col_categoria = detectar_coluna_opcional(
        df, ["categoria de desempenho"], "categoria de desempenho (TCT)", log
    )
    valida_tct = col_acertos_pct is not None and col_categoria is not None
    log("✅ Validação TCT ATIVADA." if valida_tct else "ℹ️ Validação TCT não será feita.")

    # Pelo menos um dos fluxos tem que estar disponível
    if not valida_tri and not valida_tct:
        raise ValueError("Nenhum dado válido encontrado! É necessário ter TRI ou TCT na planilha.")

    disciplina_aba = detectar_disciplina(aba_dados)
    log(f"\nDisciplina detectada pela aba: {disciplina_aba}")

    # --- Cache de regras: evita buscar a mesma combinação repetidas vezes ---
    # Chave: (medida_ref_normalizada, etapa_normalizada, disciplina_normalizada)
    # Valor: lista de regras [{"min": X, "max": Y, "classificacao": "..."}]
    cache_regras = {}

    def carregar_regra(medida_val, etapa_val, disciplina_val):
        chave = (norm(medida_val), norm(etapa_val), norm(disciplina_val))
        if chave not in cache_regras:
            try:
                cache_regras[chave] = buscar_regras(
                    medida_val, etapa_val, disciplina_val, df_faixas
                )
            except Exception as e:
                log(f"⚠️  Sem regra: medida='{medida_val}' | etapa='{etapa_val}' | disciplina='{disciplina_val}': {e}")
        return chave

    # Pré-carrega regras TRI para todas as combinações etapa/disciplina únicas dos dados
    if valida_tri:
        combinacoes = df[["etapa", "disciplina"]].dropna().drop_duplicates()
        for _, row in combinacoes.iterrows():
            carregar_regra(MEDIDA_PROFICIENCIA, row["etapa"], row["disciplina"])

    # Pré-carrega regra TCT genérica (Todas/Todas)
    if valida_tct:
        carregar_regra(MEDIDA_ACERTOS_TESTE, "todas", "todas")

    # Mostra no log todas as regras que foram carregadas
    log("\nREGRAS INTERPRETADAS:\n")
    for (medida_k, etapa_k, disc_k), regras in cache_regras.items():
        log(f"Medida: {medida_k} | Etapa: {etapa_k} | Disciplina: {disc_k}")
        for r in regras:
            max_v = "∞" if r["max"] == float("inf") else r["max"]
            log(f"  - {r['classificacao']} → {r['min']} até {max_v}")

    # --- Cria as colunas de resultado no dataframe ---
    df["calc_tri"] = ""       # classificação calculada (TRI)
    df["validacao_tri"] = ""  # "ok" ou "calc: X | planilha: Y" (TRI)
    df["calc_tct"] = ""       # classificação calculada (TCT)
    df["validacao_tct"] = ""  # "ok" ou "calc: X | planilha: Y" (TCT)

    total = len(df)

    for i, row in df.iterrows():
        etapa_norm = norm(row["etapa"])
        disciplina_norm = norm(row["disciplina"])
        linha_excel = i + 2  # +2: índice pandas começa em 0 e linha 1 é o cabeçalho

        # ======= VALIDAÇÃO TRI (Proficiência x Níveis de Aprendizagem) =======
        if valida_tri:
            try:
                valor_tri = tratar_numero(row[col_prof])

                if valor_tri is None:
                    # Célula vazia ou inválida — registra mas não trava
                    df.at[i, "validacao_tri"] = "valor inválido"
                else:
                    planilha_tri = norm(row[col_nivel])
                    chave_tri = (MEDIDA_PROFICIENCIA, etapa_norm, disciplina_norm)
                    regras_tri = cache_regras.get(chave_tri)

                    if not regras_tri:
                        df.at[i, "validacao_tri"] = "sem regra"
                    else:
                        calc_tri = classificar(valor_tri, regras_tri)
                        df.at[i, "calc_tri"] = calc_tri

                        if calc_tri == planilha_tri:
                            df.at[i, "validacao_tri"] = "ok"
                        else:
                            # Discrepância: mostra no log e grava no Excel
                            df.at[i, "validacao_tri"] = f"calc: {calc_tri} | planilha: {planilha_tri}"
                            log(f"\n❌ ERRO TRI — Linha Excel {linha_excel}")
                            log(f"  Etapa:      {row['etapa']}")
                            log(f"  Disciplina: {row['disciplina']}")
                            log(f"  Profic.:    {valor_tri}")
                            log(f"  Planilha:   {planilha_tri}")
                            log(f"  Calculado:  {calc_tri}")

            except Exception as e:
                df.at[i, "validacao_tri"] = f"erro: {str(e)}"

        # ======= VALIDAÇÃO TCT (Acertos no Teste % x Categoria de Desempenho) =======
        if valida_tct:
            try:
                valor_tct = tratar_numero(row[col_acertos_pct])

                if valor_tct is None:
                    df.at[i, "validacao_tct"] = "valor inválido"
                else:
                    planilha_tct = norm(row[col_categoria])
                    # TCT sempre usa a regra genérica "todas/todas"
                    chave_tct = (MEDIDA_ACERTOS_TESTE, "todas", "todas")
                    regras_tct = cache_regras.get(chave_tct)

                    if not regras_tct:
                        df.at[i, "validacao_tct"] = "sem regra"
                    else:
                        calc_tct = classificar(valor_tct, regras_tct)
                        df.at[i, "calc_tct"] = calc_tct

                        if calc_tct == planilha_tct:
                            df.at[i, "validacao_tct"] = "ok"
                        else:
                            df.at[i, "validacao_tct"] = f"calc: {calc_tct} | planilha: {planilha_tct}"
                            log(f"\n❌ ERRO TCT — Linha Excel {linha_excel}")
                            log(f"  Etapa:      {row['etapa']}")
                            log(f"  Disciplina: {row['disciplina']}")
                            log(f"  Acertos %:  {valor_tct}")
                            log(f"  Planilha:   {planilha_tct}")
                            log(f"  Calculado:  {calc_tct}")

            except Exception as e:
                df.at[i, "validacao_tct"] = f"erro: {str(e)}"

        # Atualiza a barra de progresso a cada linha processada
        if progress_callback and total > 0:
            progress_callback(int((i + 1) / total * 100))

    # --- Resumo final ---
    log("\n==============================")
    log("📊 RESUMO FINAL")
    log("==============================")
    log(f"Total de linhas : {total}")

    if valida_tri:
        corretas_tri = (df["validacao_tri"] == "ok").sum()
        erradas_tri = df["validacao_tri"].str.contains("calc:", na=False).sum()
        invalidas_tri = df["validacao_tri"].str.contains("inválido|erro|sem regra", na=False).sum()
        log(f"\n--- TRI (Proficiência) ---")
        log(f"✔  Corretas  : {corretas_tri}")
        log(f"❌  Erradas   : {erradas_tri}")
        log(f"⚠️  Inválidas : {invalidas_tri}")

    if valida_tct:
        corretas_tct = (df["validacao_tct"] == "ok").sum()
        erradas_tct = df["validacao_tct"].str.contains("calc:", na=False).sum()
        invalidas_tct = df["validacao_tct"].str.contains("inválido|erro|sem regra", na=False).sum()
        log(f"\n--- TCT (Acertos no Teste %) ---")
        log(f"✔  Corretas  : {corretas_tct}")
        log(f"❌  Erradas   : {erradas_tct}")
        log(f"⚠️  Inválidas : {invalidas_tct}")

    log(f"\nValidação concluída!")
    return None


# =============================================================================
# BLOCO 4 — INTERFACE GRÁFICA (Tkinter)
# Tkinter é a biblioteca padrão do Python para interfaces — não precisa instalar.
# Toda a interface é montada dentro da classe ValidadorApp.
# =============================================================================

class ValidadorApp:
    def __init__(self, root):
        """Inicializa a janela principal e cria as variáveis de estado."""
        self.root = root
        self.root.title("ValidaCorte")
        self.root.geometry("780x640")   # tamanho inicial da janela
        self.root.minsize(680, 520)     # tamanho mínimo (não deixa ficar pequeno demais)

        # Variáveis ligadas aos widgets (atualizam automaticamente quando o widget muda)
        self.arquivo_path = tk.StringVar()      # caminho do Excel selecionado
        self.aba_faixas_var = tk.StringVar()    # aba escolhida no dropdown de faixas
        self.aba_dados_var = tk.StringVar()     # aba escolhida no dropdown de dados
        self.abas_disponiveis = []              # lista de abas do Excel carregado

        self._montar_interface()  # constrói todos os widgets visuais

    def _montar_interface(self):
        """Cria e posiciona todos os widgets da janela."""
        pad = {"padx": 10, "pady": 6}  # espaçamento padrão usado em vários frames

        # ----- LINHA 1: Seleção de arquivo -----
        frame_arquivo = ttk.Frame(self.root)
        frame_arquivo.pack(fill="x", **pad)

        ttk.Label(frame_arquivo, text="Arquivo Excel:").pack(side="left")

        # Campo de texto (readonly) mostrando o caminho do arquivo selecionado
        self.entry_arquivo = ttk.Entry(
            frame_arquivo, textvariable=self.arquivo_path, state="readonly"
        )
        self.entry_arquivo.pack(side="left", fill="x", expand=True, padx=8)

        # Botão que abre o explorador de arquivos do Windows
        ttk.Button(
            frame_arquivo, text="Selecionar arquivo...", command=self.selecionar_arquivo
        ).pack(side="left")

        # ----- LINHA 2: Seleção de abas (dropdowns) -----
        frame_abas = ttk.Frame(self.root)
        frame_abas.pack(fill="x", **pad)

        ttk.Label(frame_abas, text="Aba de faixas:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        # Combobox começa desabilitado — só ativa depois de carregar o arquivo
        self.combo_faixas = ttk.Combobox(
            frame_abas, textvariable=self.aba_faixas_var, state="disabled", width=20
        )
        self.combo_faixas.grid(row=0, column=1, padx=(0, 20))

        ttk.Label(frame_abas, text="Aba de dados:").grid(row=0, column=2, sticky="w", padx=(0, 6))
        self.combo_dados = ttk.Combobox(
            frame_abas, textvariable=self.aba_dados_var, state="disabled", width=20
        )
        self.combo_dados.grid(row=0, column=3)

        # ----- LINHA 3: Botões de ação -----
        frame_botao = ttk.Frame(self.root)
        frame_botao.pack(fill="x", **pad)

        # Botão principal — desabilitado até um arquivo ser carregado
        self.btn_validar = ttk.Button(
            frame_botao, text="▶ Validar", command=self.iniciar_validacao, state="disabled"
        )
        self.btn_validar.pack(side="left")

        # Aparece só depois que a validação termina
        self.btn_abrir_pasta = ttk.Button(
            frame_botao, text="Abrir pasta do resultado",
            command=self.abrir_pasta, state="disabled"
        )
        self.btn_abrir_pasta.pack(side="left", padx=10)

        # ----- LINHA 4: Barra de progresso + status -----
        frame_progresso = ttk.Frame(self.root)
        frame_progresso.pack(fill="x", **pad)

        # Barra de progresso (0 a 100) — atualizada a cada linha processada
        self.progress = ttk.Progressbar(
            frame_progresso, orient="horizontal", mode="determinate", maximum=100
        )
        self.progress.pack(fill="x")

        # Texto de status abaixo da barra
        self.label_status = ttk.Label(frame_progresso, text="Aguardando arquivo...")
        self.label_status.pack(anchor="w", pady=(4, 0))

        # ----- ÁREA PRINCIPAL: Caixa de log -----
        frame_log = ttk.Frame(self.root)
        frame_log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        ttk.Label(frame_log, text="Log de validação:").pack(anchor="w")

        # Caixa de texto estilo terminal (fundo escuro) — apenas leitura pelo usuário
        self.text_log = tk.Text(
            frame_log,
            wrap="word",
            state="disabled",          # usuário não pode editar
            bg="#1e1e1e",              # fundo escuro (estilo terminal)
            fg="#d4d4d4",              # texto cinza claro
            insertbackground="#d4d4d4",
            font=("Consolas", 9)       # fonte monospace para alinhar melhor
        )
        scrollbar = ttk.Scrollbar(frame_log, command=self.text_log.yview)
        self.text_log.configure(yscrollcommand=scrollbar.set)

        self.text_log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.saida_path = None  # guarda o caminho do arquivo gerado (para abrir pasta)

    # =========================================================================
    # MÉTODOS DE ATUALIZAÇÃO DA INTERFACE
    # Todos usam self.root.after(0, ...) para garantir que a atualização
    # acontece na thread principal do Tkinter (obrigatório — widgets não podem
    # ser atualizados de threads secundárias diretamente).
    # =========================================================================

    def log(self, msg):
        """Adiciona uma linha ao log. Chamado pela thread de validação via callback."""
        def _append():
            self.text_log.configure(state="normal")
            self.text_log.insert("end", str(msg) + "\n")
            self.text_log.see("end")   # rola automáticamente para o final
            self.text_log.configure(state="disabled")
        self.root.after(0, _append)

    def set_progress(self, valor):
        """Atualiza a barra de progresso (0-100). Chamado pela thread de validação."""
        self.root.after(0, lambda: self.progress.configure(value=valor))

    def set_status(self, texto):
        """Atualiza o texto de status abaixo da barra de progresso."""
        self.root.after(0, lambda: self.label_status.configure(text=texto))

    # =========================================================================
    # MÉTODOS DE AÇÃO (ligados aos botões)
    # =========================================================================

    def selecionar_arquivo(self):
        """
        Abre o explorador de arquivos do Windows, carrega o Excel escolhido
        e preenche os dropdowns com as abas disponíveis.
        Tenta pré-selecionar automaticamente a primeira aba G_* e D_*.
        """
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if not caminho:
            return  # usuário cancelou a seleção

        self.arquivo_path.set(caminho)

        # Limpa o log e reseta a barra de progresso ao trocar de arquivo
        self.text_log.configure(state="normal")
        self.text_log.delete("1.0", "end")
        self.text_log.configure(state="disabled")
        self.progress.configure(value=0)
        self.btn_abrir_pasta.configure(state="disabled")
        self.saida_path = None

        try:
            # Lê só os metadados do Excel (nomes das abas, sem carregar os dados ainda)
            xls = pd.ExcelFile(caminho)
            self.abas_disponiveis = xls.sheet_names

            # Preenche os dropdowns com as abas encontradas e libera a seleção
            self.combo_faixas.configure(values=self.abas_disponiveis, state="readonly")
            self.combo_dados.configure(values=self.abas_disponiveis, state="readonly")

            # Tenta pré-selecionar: primeira aba G_* para faixas, primeira D_* para dados
            self.aba_faixas_var.set("")
            self.aba_dados_var.set("")
            for nome in self.abas_disponiveis:
                if nome.upper().startswith("G_") and not self.aba_faixas_var.get():
                    self.aba_faixas_var.set(nome)
                if nome.upper().startswith("D_") and not self.aba_dados_var.get():
                    self.aba_dados_var.set(nome)

            self.btn_validar.configure(state="normal")
            self.set_status(
                f"Arquivo carregado: {os.path.basename(caminho)} "
                f"— {len(self.abas_disponiveis)} abas encontradas"
            )
            self.log(f"Arquivo selecionado: {caminho}")
            self.log(f"Abas disponíveis: {', '.join(self.abas_disponiveis)}")

        except Exception as e:
            messagebox.showerror("Erro ao abrir arquivo", str(e))
            self.set_status("Erro ao carregar o arquivo.")

    def abrir_pasta(self):
        """Abre o Explorer do Windows na pasta onde o resultado foi salvo."""
        if self.saida_path and os.path.exists(self.saida_path):
            os.startfile(os.path.dirname(self.saida_path))
        else:
            messagebox.showinfo("Aviso", "Ainda não há resultado gerado.")

    def iniciar_validacao(self):
        """
        Chamado ao clicar em "▶ Validar".
        Valida os inputs, desabilita o botão durante o processo,
        e inicia a validação em uma thread separada para não travar a interface.
        """
        arquivo = self.arquivo_path.get()
        aba_faixas = self.aba_faixas_var.get()
        aba_dados = self.aba_dados_var.get()

        if not arquivo:
            messagebox.showwarning("Atenção", "Selecione um arquivo Excel primeiro.")
            return
        if not aba_faixas or not aba_dados:
            messagebox.showwarning("Atenção", "Selecione a aba de faixas e a aba de dados.")
            return

        # Desabilita botões e reseta progresso enquanto roda
        self.btn_validar.configure(state="disabled")
        self.btn_abrir_pasta.configure(state="disabled")
        self.progress.configure(value=0)
        self.set_status("Validando, aguarde...")

        # Roda a validação em thread separada — sem isso, a janela travaria
        thread = threading.Thread(
            target=self._executar_validacao_thread,
            args=(arquivo, aba_faixas, aba_dados),
            daemon=True  # a thread morre automaticamente se a janela for fechada
        )
        thread.start()

    def _executar_validacao_thread(self, arquivo, aba_faixas, aba_dados):
        """
        Executa rodar_validacao() em background.
        Usa self.log e self.set_progress como callbacks para atualizar a interface.
        Ao terminar, reabilita os botões e mostra o resultado.
        """
        try:
            rodar_validacao(
                arquivo, aba_faixas, aba_dados,
                log=self.log,
                progress_callback=self.set_progress
            )
            self.set_status("Validação concluída!")
            self.root.after(
                0, lambda: messagebox.showinfo("Concluído", "Validação finalizada!\nVeja o resumo no log.")
            )

        except Exception as e:
            # Captura o erro completo (com traceback) e exibe no log
            self.log(f"\n🛑 ERRO FATAL: {e}")
            self.log(traceback.format_exc())
            self.set_status("Erro durante a validação. Veja o log.")
            self.root.after(0, lambda: messagebox.showerror("Erro", str(e)))

        finally:
            # Sempre reabilita o botão, independente de ter dado erro ou não
            self.root.after(0, lambda: self.btn_validar.configure(state="normal"))


# =============================================================================
# BLOCO 5 — PONTO DE ENTRADA
# =============================================================================

def main():
    root = tk.Tk()
    # Tenta usar o tema visual "vista" do Windows (mais moderno que o padrão)
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass  # se falhar, continua com o tema padrão sem erro

    app = ValidadorApp(root)
    root.mainloop()  # inicia o loop de eventos da interface (fica aberto até fechar a janela)


if __name__ == "__main__":
    # Só executa se rodar diretamente (não se importado como módulo)
    main()
