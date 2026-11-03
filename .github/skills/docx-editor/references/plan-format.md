# Formato do plano de edição

Referência de `scripts/edit_docx.py`. O plano é um JSON em UTF-8. Chaves desconhecidas são recusadas, o que pega
erros de digitação antes de qualquer alteração.

## Índice

- [Estrutura geral](#estrutura-geral)
- [Seletores de parágrafo](#seletores-de-parágrafo)
- [Conteúdo: text e runs](#conteúdo-text-e-runs)
- [Operações](#operações)
- [Exemplo completo](#exemplo-completo)
- [Relatório de saída](#relatório-de-saída)

## Estrutura geral

```json
{
  "track_changes": "auto",
  "author": "Nome do revisor",
  "date": "2026-11-03T12:00:00Z",
  "operations": [ { "op": "replace_text", "find": "2023", "replace": "2024" } ]
}
```

| Chave | Padrão | Significado |
|---|---|---|
| `operations` | obrigatória | Lista executada em ordem; qualquer erro cancela tudo |
| `track_changes` | `"auto"` | `"auto"` segue `w:trackRevisions` do documento; `true` ou `false` forçam |
| `author` | `"docx-editor"` | Autor das revisões |
| `date` | agora (UTC) | Data das revisões |

## Seletores de parágrafo

Usados em `target`, `anchor` e `like`. **Sempre se referem ao documento original**, mesmo depois de inserções ou
remoções feitas por operações anteriores do mesmo plano.

| Seletor | Exemplo | Observação |
|---|---|---|
| Índice | `{"p": 12, "expect": "Cláusula"}` | Índice do `inspect`. `expect` é opcional e recomendado: falha se o texto não estiver ali |
| Texto contido | `{"contains": "prazo de vigência"}` | Precisa ser único. Com mais de um, use `occurrence` (base 0) |
| Texto exato | `{"equals": "Objeto"}` | O parágrafo inteiro |
| Estilo | `{"style": "Heading 2", "contains": "Anexo"}` | Aceita styleId ou nome; combina com `contains` e `equals` |
| Indicador | `{"bookmark": "valor_total"}` | Parágrafo onde o indicador começa |
| `w14:paraId` | `{"para_id": "1A2B3C4D"}` | Visto no `inspect --json` |
| Célula | `{"table": 0, "row": 1, "cell": 2, "cell_p": 0}` | Índices base 0; `cell_p` é o parágrafo dentro da célula |

Para outra parte do pacote, acrescente `"part"`: `{"part": "word/footer1.xml", "p": 0}`. O índice vale dentro dessa
parte, conforme `inspect --part`.

Em `like`, a forma `{"style": "Heading 2"}` procura o exemplar **mais próximo antes da âncora**, copiando também a
formatação direta que o autor costuma aplicar a esse estilo.

## Conteúdo: text e runs

**`text`** é uma string com a formatação herdada do contexto. `\t` vira tabulação e `\n` vira quebra de linha
(`w:br`, não parágrafo novo).

**`runs`** é uma lista para trechos com formatação diferente:

```json
"runs": [
  { "text": "Nota:", "like": { "p": 4, "text": "Importante" } },
  { "text": " o prazo é de 30 dias." },
  { "text": "Lei 8.666", "rstyle": "Strong" }
]
```

| Chave | Efeito |
|---|---|
| `text` | Obrigatória |
| `like` | Copia a formatação de um run existente que contenha `text` por inteiro, no parágrafo do seletor. **Preferida** |
| `rstyle` | Aplica estilo de caractere existente (id ou nome); estilo inexistente é recusado |
| `bold`, `italic`, `underline` | Formatação direta (`true`/`false`). Use só se o documento já faz assim; o `compare` aponta formatação sem precedente |

## Operações

### replace_text

Troca um trecho, mesmo que esteja dividido em vários runs.

| Chave | Significado |
|---|---|
| `find` | Texto exato, sensível a maiúsculas. Não atravessa parágrafos nem objetos (imagem, nota, campo, quebra de página) |
| `replace` ou `runs` | O novo conteúdo; `"replace": ""` apaga o trecho |
| `target` | Restringe a um parágrafo |
| `parts` | Sem `target`: `"main"` (padrão), `"all"` ou lista de partes |
| `count` | Ocorrências esperadas: `1` (padrão), um número ou `"all"`. Divergência é erro |
| `allow_field_result` | `true` permite editar texto calculado (sumário, `REF`). O Word pode sobrescrever ao atualizar campos |

O prefixo e o sufixo comuns são preservados por limite de palavra, logo `"vence em 2023"` → `"vence em 2030"` altera só
`2023`. Em caixas de texto, a edição também é aplicada à cópia em `mc:Fallback`.

Para acrescentar texto dentro de um parágrafo, troque um trecho por ele mesmo mais o acréscimo:
`"find": "30 dias", "replace": "30 dias corridos"`.

### rewrite_paragraph

Dá ao parágrafo um novo texto: `target` + (`text` ou `runs`).

- Com `text` e semelhança de 50% ou mais, aplica diff por palavra: o que não mudou mantém formatação, notas, links e
  comentários.
- Com pouca semelhança, troca o texto todo usando a **formatação dominante** do parágrafo; objetos permanecem.
- Com `runs`, troca o texto todo. Parágrafo com resultado de campo é recusado.

`pPr`, incluindo quebra de seção, nunca é alterado.

### insert_paragraph

| Chave | Significado |
|---|---|
| `anchor` | Parágrafo de referência |
| `position` | `"after"` (padrão) ou `"before"` |
| `text` ou `runs` | Conteúdo de um parágrafo |
| `paragraphs` | Lista de strings ou de `{ "text" \| "runs", "like" }`, inseridos em ordem |
| `like` | Modelo. Padrão: a própria âncora |

Copia do modelo as propriedades de parágrafo (incluindo numeração) e a formatação dominante do texto. **Não** copia
`sectPr`, `paraId`, indicadores, comentários nem objetos. Inserções repetidas "after" na mesma âncora saem na ordem do plano.

Se a âncora termina uma seção, `"after"` coloca o parágrafo na seção seguinte (há aviso); use `"before"` para ficar na mesma.

### delete_paragraph

`target` + `force` opcional (obrigatório se houver imagem ou objeto).

- Recusa: parágrafo com quebra de seção, parte de campo que atravessa parágrafos, separador entre duas tabelas ou
  parágrafo exigido depois da última tabela.
- Único bloco de célula, cabeçalho ou rodapé: é esvaziado em vez de removido, pois o formato exige um parágrafo.
- Indicadores e comentários vão para o parágrafo vizinho (há aviso). Notas e imagens usadas só por ele saem do pacote.
- Com revisões, o conteúdo fica marcado como excluído e nada sai do pacote.

### insert_table_row

| Chave | Significado |
|---|---|
| `table`, `row` | Tabela (índice do `--tables`) e linha de referência |
| `cells` | Um item por célula da linha modelo: `null`, texto, `{"runs": [...]}` ou `{"paragraphs": [...]}` |
| `position` | `"after"` (padrão) ou `"before"` |
| `like_row` | Linha modelo. Padrão: `row`. Ao inserir depois do cabeçalho, aponte para uma linha de dados |
| `part` | Parte, se não for o documento principal |

Clona `trPr`, `tcPr` e o primeiro parágrafo de cada célula. Se a nova linha cair dentro de uma mesclagem vertical, a
célula continua a mesclagem e exige `null`; caso contrário, a mesclagem do modelo é descartada. O marcador de cabeçalho
repetido (`w:tblHeader`) não é clonado.

### delete_table_row

`table`, `row`, `part` e `force` opcionais. Se a linha inicia uma mesclagem vertical, a linha seguinte herda o início e o
conteúdo. A única linha da tabela é recusada: use `delete_table`.

### insert_table

| Chave | Significado |
|---|---|
| `anchor`, `position` | Onde inserir |
| `like_table` | Índice da tabela modelo (`like_part` para outra parte) |
| `rows` | Lista de linhas; cada linha tem um item por célula do modelo |
| `header` | A primeira linha usa o cabeçalho do modelo. Padrão: `true` se o modelo tem mais de uma linha |

A grade é a do modelo: o número de colunas precisa coincidir. Quando necessário, um parágrafo vazio é acrescentado para
evitar fusão com tabela vizinha ou atender à exigência de parágrafo após a tabela.

### delete_table

`table`, `part` e `force` opcionais. Indicadores e comentários vão para o parágrafo vizinho.

### insert_image

| Chave | Significado |
|---|---|
| `anchor`, `position` | Onde inserir, sempre em parágrafo próprio |
| `image` | Caminho de PNG, JPEG ou GIF |
| `width_cm`, `height_cm` | Tamanho; com um só, a proporção é mantida. Sem nenhum, usa 96 dpi |
| `alt` | Texto alternativo. Omitir gera aviso |
| `like` | Modelo do parágrafo. Prefira uma figura já existente |

A imagem é limitada à largura útil da seção. Entrelinha exata direta no modelo é removida, pois cortaria a figura.

## Exemplo completo

```json
{
  "author": "Copilot",
  "operations": [
    { "op": "replace_text", "find": "31 de dezembro de 2025", "replace": "31 de dezembro de 2026", "count": "all" },
    { "op": "rewrite_paragraph",
      "target": { "p": 18, "expect": "reajuste" },
      "text": "O valor será reajustado anualmente pelo IPCA, a partir da data de assinatura." },
    { "op": "insert_paragraph",
      "anchor": { "p": 24, "expect": "fornecer os materiais" },
      "text": "manter sigilo sobre as informações recebidas;" },
    { "op": "insert_table_row",
      "table": 1, "row": 3,
      "cells": [ "Suporte técnico", "12 meses", "R$ 4.800,00" ] },
    { "op": "delete_paragraph", "target": { "contains": "Esta cláusula será revista" } },
    { "op": "replace_text", "find": "v1.0", "replace": "v2.0", "parts": [ "word/footer1.xml" ] }
  ]
}
```

## Relatório de saída

`edit_docx.py` imprime um JSON (também gravado com `--report arquivo.json`):

```json
{
  "ok": true,
  "applied": [ { "op": "replace_text", "find": "...", "occurrences": 2 } ],
  "warnings": [ "operação #5 (delete_paragraph): indicadores ['x'] foram preservados no parágrafo vizinho" ],
  "new_findings": [],
  "track_changes": false,
  "output": "C:/caminho/doc_editado.docx"
}
```

Código de saída 0 indica sucesso. Código 2 indica plano recusado ou saída descartada: `error` traz o motivo e nenhum
arquivo é gravado.
