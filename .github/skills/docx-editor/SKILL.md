---
name: docx-editor
description: >-
  Edita documentos Word (.docx) existentes com fidelidade total ao padrão visual e estrutural do
  original: insere, substitui, remove ou reescreve conteúdo preservando estilos, numeração,
  cabeçalhos e rodapés, seções, tabelas, imagens, campos, comentários, revisões e propriedades, sem
  gerar avisos de reparo no Word, LibreOffice ou Google Docs. Use sempre que o pedido envolver alterar
  um .docx já existente, mesmo sem citar a palavra "docx": contrato, relatório, proposta, ata, modelo
  ou template em Word; atualizar cláusula, trocar datas ou valores, acrescentar seção ou linha de
  tabela, corrigir texto, editar com controle de alterações. Também para "edit this Word document",
  "update the contract", "tracked changes". Não use para criar documento do zero sem modelo, nem para
  apenas ler ou resumir um arquivo.
---

# docx-editor

Edita um `.docx` existente de modo que o resultado pareça escrito pelo mesmo autor, no mesmo template,
e que tudo o que não foi pedido continue idêntico.

A abordagem é cirúrgica: os scripts alteram só os elementos necessários do pacote OOXML e devolvem as
demais partes byte a byte. Evite reconstruir o documento (python-docx salvando tudo, conversão via
Markdown ou HTML, "Salvar como" pelo LibreOffice): esses caminhos descartam silenciosamente estilos,
temas e metadados, e é exatamente isso que o usuário pediu para não acontecer.

Nas instruções abaixo, `$SKILL` é a pasta deste arquivo (normalmente `.github/skills/docx-editor`).

## Preparação

Os scripts exigem Python 3.10+ e `lxml`. Verifique com `python -c "import lxml"`. Se faltar, instale em
um ambiente virtual fora do projeto do usuário e use o Python dele nos comandos:

```
python -m venv <pasta-temporaria>/venv
<pasta-temporaria>/venv/<bin-ou-Scripts>/python -m pip install -r $SKILL/requirements.txt
```

O original nunca é sobrescrito. Salvo pedido em contrário, grave `<nome>_editado.docx` ao lado dele.

## Fluxo de trabalho

### 1. Entenda o documento antes de escrever qualquer coisa

```
python $SKILL/scripts/inspect_docx.py doc.docx --profile
python $SKILL/scripts/inspect_docx.py doc.docx
```

`--profile` traz o padrão a imitar: estilos em uso e sua formatação, listas, tabelas, seções, controle de
alterações e sinais de redação. A listagem mostra cada parágrafo assim:

```
[12] Titulo2 num=3:0 T0R1C2 {campo,nota} "Texto do parágrafo"
 │    │       │      │       │            └ texto visível, pronto para colar no plano
 │    │       │      │       └ objetos presentes no parágrafo
 │    │       │      └ posição em tabela (tabela, linha, célula)
 │    │       └ numeração (numId:nível)
 │    └ styleId (costuma ser localizado: Ttulo1, Titulo2...)
 └ índice usado nos alvos do plano
```

Outras visões úteis: `--part word/header1.xml`, `--grep "termo"`, `--range 10:40 --runs` (formatação run
a run), `--tables` e `--xml N`. Para ler cabeçalhos, rodapés e notas, liste as partes indicadas em
`partes_de_texto` no perfil.

Ainda nesta etapa, procure as **consequências indiretas** do pedido: outra menção ao valor alterado, a
mesma data no cabeçalho, o sumário citando o título que vai mudar, uma referência cruzada ao item que vai
sair. Avise o usuário sobre as que estiverem fora do pedido, em vez de alterá-las por conta própria.

### 2. Monte o plano de edição

O plano é um JSON aplicado de uma vez. Formato completo e exemplos: [references/plan-format.md](references/plan-format.md).

| Pedido | Operação | Como preserva o padrão |
|---|---|---|
| Trocar um trecho | `replace_text` | Atravessa runs divididos e herda a formatação do trecho trocado |
| Reescrever um parágrafo | `rewrite_paragraph` | Diff por palavra: o que não mudou mantém negrito, notas e links |
| Novo parágrafo ou item de lista | `insert_paragraph` | Clona as propriedades do modelo, sem levar IDs nem quebra de seção |
| Remover parágrafo | `delete_paragraph` | Guarda indicadores e comentários no vizinho; protege seções e campos |
| Linha de tabela | `insert_table_row`, `delete_table_row` | Clona a linha modelo e mantém mesclagens coerentes |
| Tabela inteira | `insert_table`, `delete_table` | Clona estilo, grade e formatação de uma tabela existente |
| Imagem | `insert_image` | Registra mídia, relacionamento e tipo; limita à largura da seção |

Regras para o plano:

- **Prefira a menor operação que resolve.** `replace_text` em uma palavra é mais seguro do que reescrever
  o parágrafo, e reescrever é mais seguro do que apagar e inserir de novo.
- **Escolha como modelo (`like`) um parágrafo do mesmo papel**: novo item de lista a partir de um item, novo
  título a partir de um título. Quase toda a fidelidade visual vem dessa escolha.
- **Use `p` com `expect`** para que um índice errado falhe em vez de editar o parágrafo errado. Os índices
  sempre se referem ao documento original, então não os recalcule depois de inserções.
- **Não simule formatação.** Numeração digitada à mão, parágrafo vazio como espaçamento, negrito no lugar de
  estilo de título ou tabulações para alinhar quebram a estrutura mesmo quando parecem certos.
- **Formatação direta (`bold`, `italic`) só quando o documento já faz isso no mesmo contexto.** Melhor ainda:
  copie de um trecho real com `like` ou use um estilo de caractere existente com `rstyle`.

### 3. Escreva como o autor

O texto novo deve passar despercebido. Antes de redigir, leia os parágrafos vizinhos e use os sinais de
`redacao` no perfil:

- **Idioma, tratamento e tempo verbal** dos parágrafos ao redor ("a Contratada deverá", não "você deve").
- **Tamanho das frases e dos parágrafos**: a mediana do documento é a referência.
- **Títulos**: mesma caixa e pontuação final dos demais (`caixa_dos_titulos`).
- **Itens de lista**: mesma terminação (`fim_de_item_de_lista`, por exemplo `;` e `.` no último).
- **Tipografia**: aspas curvas ou retas, travessão ou hífen, espaço depois do ponto.
- **Números e datas**: mesmo formato decimal e de data (`numeros`, `formatos_de_data`).
- **Termos definidos**: se o documento diz "Contratante", não escreva "cliente". Procure o termo com `--grep`.

Não acrescente conteúdo que não foi pedido nem "melhore" trechos vizinhos.

### 4. Controle de alterações

Por padrão (`track_changes: "auto"`) as edições viram revisões quando o documento já está com o controle de
alterações ativo. Se o usuário pedir revisões, use `true` e informe `author`. Revisões pendentes de terceiros
nunca são aceitas nem rejeitadas pelos scripts: se um trecho alvo estiver dentro de uma, a operação é recusada
e a decisão é do usuário.

### 5. Aplique, valide e compare

```
python $SKILL/scripts/edit_docx.py doc.docx plano.json --dry-run
python $SKILL/scripts/edit_docx.py doc.docx plano.json -o doc_editado.docx
python $SKILL/scripts/validate_docx.py doc_editado.docx --baseline doc.docx
python $SKILL/scripts/compare_docx.py doc.docx doc_editado.docx
python $SKILL/scripts/render_docx.py doc_editado.docx --outdir <pasta-temporaria> --compare doc.docx
```

- `edit_docx.py` é tudo-ou-nada e descarta a saída se ela introduzir erro estrutural. **Leia `warnings`**:
  são decisões que você precisa conferir ou relatar.
- `validate_docx.py --baseline` só reprova problemas **novos**. Os herdados do original não são seus, mas vale
  mencioná-los se parecerem graves.
- `compare_docx.py` é a prova de escopo. Confira cada linha de modificados, inseridos e removidos contra o pedido.
  **VIOLAÇÃO** é alteração colateral: corrija o plano em vez de ignorar. **ATENÇÃO** pede conferência.
- `--allow` libera outras partes. Por padrão só `word/document.xml` pode mudar, então libere o que o plano
  realmente exige:

| O plano mexe em | Acrescente |
|---|---|
| Cabeçalho ou rodapé | `--allow word/document.xml --allow word/header1.xml` (a parte editada) |
| Parágrafo com nota de rodapé removido | `--allow word/document.xml --allow word/footnotes.xml` |
| Imagem nova ou removida | `--allow word/document.xml --allow "word/media/*" --allow "word/_rels/*" --allow "[Content_Types].xml"` |

- `render_docx.py` abre o arquivo de verdade no LibreOffice e compara o número de páginas. Se ele sair com
  código 3 (LibreOffice ausente), diga que esse teste não foi feito.

Se algo reprovar, corrija o plano e gere de novo a partir do **original**, nunca em cima da saída anterior.

### 6. Relate o resultado

Termine com um resumo curto e fiel:

- o que mudou, parágrafo a parágrafo, e o caminho do arquivo gerado;
- avisos relevantes e consequências indiretas encontradas (por exemplo, outras menções ao valor antigo);
- **o que o usuário precisa fazer**: se títulos ou conteúdo mudaram e há sumário ou referências cruzadas, atualizar
  os campos no Word (`Ctrl+A`, `F9`);
- **o que foi e o que não foi verificado.** Estes scripts não abrem o Word nem o Google Docs. A validação estrutural
  reduz muito o risco, mas não substitui abrir o arquivo. Nunca afirme que "abre sem erros no Word" se isso não foi
  testado; diga que passou nas validações e recomende abrir o arquivo para conferir.

## Proteções e o que fazer quando disparam

As recusas existem para evitar danos silenciosos. Não as contorne por reflexo: leia a mensagem.

| Mensagem | Por que existe | O que fazer |
|---|---|---|
| Resultado de campo | O Word recalcula sumário, `REF` e `PAGE`, e a edição seria perdida | Edite a origem (o título, por exemplo). Use `allow_field_result: true` só para manter o texto em cache coerente |
| Atravessa um objeto | A troca apagaria nota, imagem ou campo no meio do trecho | Edite cada lado do objeto em operações separadas |
| Quebra de seção | Remover o parágrafo funde seções e muda cabeçalho, margens e orientação | Esvazie o texto com `rewrite_paragraph` ou confirme a fusão com o usuário e use a rota manual |
| Campo atravessa parágrafos | Remover um pedaço corrompe o campo | Peça ao usuário para atualizar o sumário no Word ou use a rota manual para remover o campo inteiro |
| Imagem ou objeto, use `force` | Evita remover figura junto com o texto | Use `force: true` só se o pedido incluir a figura |
| Revisão pendente | Não dá para sobrepor revisão de outra pessoa com segurança | Peça para aceitar ou rejeitar antes, ou use a rota manual |
| `count` não bate | A troca atingiria mais ou menos do que você supôs | Confira as ocorrências; use `target` ou `count` exato |
| Texto não encontrado, com sugestão | Aspas curvas, espaço não separável ou travessão diferentes | Copie o texto exato da sugestão ou do `inspect` |
| Parágrafo vinculado a dados (aviso) | O Word pode restaurar o valor vinculado ao abrir | Avise o usuário: pode ser preciso alterar a propriedade do documento |

Remover um parágrafo também remove o texto das notas que só ele referenciava e as imagens que só ele usava, para que
o conteúdo apagado não fique escondido no arquivo. Comentários e indicadores, por outro lado, são preservados no
parágrafo vizinho: são trabalho de terceiros ou alvos de referências. Se o usuário pediu para apagar um comentário,
isso é um pedido à parte, feito pela rota manual.

## O que evitar

- **`w:updateFields` em `settings.xml`**: atualiza o sumário, mas faz o Word exibir um aviso ao abrir, o que viola o
  requisito de abrir sem avisos. Oriente o usuário a atualizar os campos.
- **Alterar `docProps/core.xml`** (autor, datas, título) sem pedido explícito.
- **Pretty-print em XML de documento**: pode criar espaços visíveis e diffs enormes.
- **Copiar `w:sectPr`, `w14:paraId` ou `wp:docPr` ao clonar à mão**: cria seção fantasma ou IDs duplicados.

## Rota manual (exceção)

Para o que as operações não cobrem, como fundir seções, remover um sumário inteiro, excluir comentário ou alterar
um estilo a pedido, edite o XML diretamente:

```
python $SKILL/scripts/unpack_docx.py doc.docx <pasta>
# edite só os trechos necessários
python $SKILL/scripts/pack_docx.py <pasta> doc_editado.docx --original doc.docx
```

Leia antes [references/ooxml-pitfalls.md](references/ooxml-pitfalls.md), que lista a ordem obrigatória de elementos e
os erros que disparam o aviso de reparo. Depois valide e compare como na etapa 5.

## Limites conhecidos

- Não abre `.doc` legado, arquivo protegido por senha nem Strict Open XML: os scripts recusam com mensagem clara.
  Peça uma cópia `.docx` comum, sem tentar remover a proteção.
- `find` não atravessa parágrafos: para trechos longos, combine `rewrite_paragraph` e `delete_paragraph`.
- Fórmulas (OMML), gráficos e objetos embutidos são preservados, mas não editados.
- Quebra de layout (linhas viúvas, tabela mudando de página) só aparece na renderização; o número de páginas é um
  indício, não uma prova.
