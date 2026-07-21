# ValidaCorte

Ferramenta com interface gráfica para validar automaticamente os resultados de proficiência e acertos no teste das planilhas de resultados, comparando os valores com as faixas de corte definidas pela equipe.

---

## O que o programa faz

Para cada linha da planilha de dados, o programa:

1. Lê o valor de **Proficiência** e compara com a classificação na coluna **Níveis de Aprendizagem** (validação TRI)
2. Lê o valor de **Acertos no Teste %** e compara com a coluna **Categoria de Desempenho** (validação TCT)
3. Aplica as faixas de corte definidas na aba de faixas (ex: G_1) para cada combinação de etapa, disciplina e medida de referência
4. Mostra no log todas as linhas onde a classificação calculada diverge da que está na planilha
5. Exibe um resumo com total de linhas corretas, erradas e inválidas — separado por TRI e TCT

---

## Estrutura esperada da planilha Excel

### Aba de faixas (ex: G_1)

| Medida de Referência | Etapa | Disciplina | Faixas |
|---|---|---|---|
| Porcentagem de acerto no teste | Todas | Todas | Muito baixo (Até 25); Baixo (26 até 50); Médio (51 até 75); Alto (76 até 100) |
| Proficiência | ENSINO MÉDIO - 2ª SÉRIE | Língua Portuguesa | Defasagem (Até 30); Aprendizado Intermediário (31 até 70); Aprendizado Adequado (71 até 100) |
| ... | ... | ... | ... |

**Colunas obrigatórias:** `Medida de Referência`, `Etapa`, `Disciplina`, `Faixas`

As faixas são separadas por `;` e os limites precisam estar entre parênteses. Formatos reconhecidos:
- `Até X` → de 0 até X
- `X até Y` → de X até Y
- `X ou mais` / `Acima de X` → de X até infinito

### Aba de dados (ex: D_1)

Precisa ter obrigatoriamente as colunas `Etapa` e `Disciplina`.

Para a **validação TRI**, precisa ter:
- `Proficiência` ou `Proficiência - Professor`
- `Níveis de Aprendizagem` ou `Padrão de Desempenho` ou `Padrões de Desempenho`

Para a **validação TCT**, precisa ter:
- `Acertos no Teste %`
- `Categoria de Desempenho`

Se apenas uma das validações (TRI ou TCT) tiver as colunas necessárias, o programa valida só essa. Se ambas existirem, valida as duas de forma independente para cada linha.

---

## Como usar

### Rodando o script diretamente (Python instalado)

**Requisitos:**
```
pip install pandas openpyxl
```

**Execução:**
```
python valida_gui.py
```

### Rodando o executável (.exe)

Basta dar duplo clique no `ValidadorSAEPE.exe`. Não precisa de Python instalado.

Para gerar o `.exe` a partir do script:
```
pip install pyinstaller
python -m PyInstaller --onefile --windowed --name "ValidadorSAEPE" valida_gui.py
```
O executável gerado fica na pasta `dist/`.

---

## Passo a passo de uso

1. Abra o programa
2. Clique em **Selecionar arquivo...** e escolha o `.xlsx`
3. O programa detecta automaticamente a aba de faixas (primeira que começa com `G_`) e a aba de dados (primeira que começa com `D_`). Você pode trocar pelos dropdowns se necessário
4. Clique em **▶ Validar**
5. Acompanhe o progresso na barra e no log em tempo real
6. Ao terminar, o resumo aparece no log

---

## Interpretando o log

### Regras interpretadas
Logo no início, o programa mostra as faixas que encontrou e como as interpretou:
```
Medida: proficiência | Etapa: ensino médio - 2ª série | Disciplina: língua portuguesa
  - defasagem → 0 até 30.0
  - aprendizado intermediário → 31.0 até 70.0
  - aprendizado adequado → 71.0 até 100.0
```
Se uma regra aparecer errada aqui, o problema está no texto da célula na aba de faixas.

### Erros de validação
```
❌ ERRO TRI — Linha Excel 47
  Etapa:      ENSINO MÉDIO - 2ª SÉRIE
  Disciplina: Língua Portuguesa
  Profic.:    185.3
  Planilha:   básico
  Calculado:  adequado
```
Indica que o valor `185.3` foi calculado como `adequado`, mas a planilha diz `básico`.

### Resumo final
```
📊 RESUMO FINAL
Total de linhas : 15936

--- TRI (Proficiência) ---
✔  Corretas  : 15800
❌  Erradas   : 12
⚠️  Inválidas : 124

--- TCT (Acertos no Teste %) ---
✔  Corretas  : 15900
❌  Erradas   : 0
⚠️  Inválidas : 36
```

| Status | Significado |
|---|---|
| ✔ Corretas | Classificação calculada bate com a planilha |
| ❌ Erradas | Classificação calculada é diferente da planilha |
| ⚠️ Inválidas | Valor numérico ausente, inválido, ou sem faixa de corte cadastrada para aquela etapa/disciplina |

---

## Regras de arredondamento

O programa usa arredondamento padrão: valores com `.5` ou mais sobem, abaixo de `.5` descem.

Exemplos:
- `30.5` → `31` → **Aprendizado Intermediário** (não Defasagem)
- `30.4` → `30` → **Defasagem**
- `230.96` → `231` → cai na faixa seguinte

O Python nativo (`round()`) usa arredondamento bancário e produziria resultados incorretos nesses casos. O programa usa `Decimal` com `ROUND_HALF_UP` para evitar esse problema.

---

## Detecção de disciplina por aba

O programa identifica a disciplina pelo nome da aba de dados (usado apenas no log, não afeta a validação):

| Aba começa com | Disciplina |
|---|---|
| D_1 | Língua Portuguesa |
| D_2 | Matemática |
| D_29 | Ciências da Natureza |
| D_30 | Ciências Humanas |

---

## Dependências

| Biblioteca | Uso |
|---|---|
| `pandas` | Leitura e manipulação das planilhas Excel |
| `openpyxl` | Motor de leitura de arquivos `.xlsx` usado pelo pandas |
| `tkinter` | Interface gráfica (já vem com Python, não precisa instalar) |
| `decimal` | Arredondamento preciso (já vem com Python) |