# Armadilhas de OOXML para edição manual

Leia antes de usar a rota `unpack_docx.py` / `pack_docx.py`. As operações de `edit_docx.py` já tratam tudo isto; a
lista só importa quando você mesmo mexe no XML.

## Índice

- [Regras que evitam o aviso de reparo](#regras-que-evitam-o-aviso-de-reparo)
- [Ordem obrigatória dos filhos](#ordem-obrigatória-dos-filhos)
- [Texto e runs](#texto-e-runs)
- [Revisões](#revisões)
- [Campos](#campos)
- [Tabelas](#tabelas)
- [Seções](#seções)
- [Numeração](#numeração)
- [Imagens e relacionamentos](#imagens-e-relacionamentos)
- [Comentários e notas](#comentários-e-notas)
- [Compatibilidade entre aplicativos](#compatibilidade-entre-aplicativos)
- [Tarefas que só a rota manual cobre](#tarefas-que-só-a-rota-manual-cobre)

## Regras que evitam o aviso de reparo

1. Respeite a ordem dos filhos em `pPr`, `tblPr`, `tcPr` e `sectPr`.
2. Toda `w:tc` termina em `w:p`, mesmo que vazio. Tabela aninhada não pode ser o último bloco da célula.
3. Todo `w:id`, `r:id`, `r:embed`, `w:numId` e `w:pStyle` usado precisa existir em sua parte de destino.
4. Toda parte precisa de tipo em `[Content_Types].xml`, e todo `Override` precisa apontar para uma parte existente.
5. Não remova declarações `xmlns` da raiz. Todo prefixo em `mc:Ignorable` precisa estar declarado.
6. `w14:paraId` precisa ser único e menor que `0x80000000`.
7. `w:fldChar` precisa estar balanceado: `begin`, `separate` (opcional) e `end`.

## Ordem obrigatória dos filhos

**`w:pPr`**: `pStyle, keepNext, keepLines, pageBreakBefore, framePr, widowControl, numPr, suppressLineNumbers, pBdr, shd,
tabs, suppressAutoHyphens, kinsoku, wordWrap, overflowPunct, topLinePunct, autoSpaceDE, autoSpaceDN, bidi, adjustRightInd,
snapToGrid, spacing, ind, contextualSpacing, mirrorIndents, suppressOverlap, jc, textDirection, textAlignment,
textboxTightWrap, outlineLvl, divId, cnfStyle, rPr, sectPr, pPrChange`

**`w:rPr`**: `ins, del, moveFrom, moveTo` (só na marca de parágrafo), depois `rStyle, rFonts, b, bCs, i, iCs, caps, smallCaps,
strike, dstrike, outline, shadow, emboss, imprint, noProof, snapToGrid, vanish, webHidden, color, spacing, w, kern, position,
sz, szCs, highlight, u, effect, bdr, shd, fitText, vertAlign, rtl, cs, em, lang, eastAsianLayout, specVanish, oMath, rPrChange`

**`w:tblPr`**: `tblStyle, tblpPr, tblOverlap, bidiVisual, tblStyleRowBandSize, tblStyleColBandSize, tblW, jc, tblCellSpacing,
tblInd, tblBorders, shd, tblLayout, tblCellMar, tblLook, tblCaption, tblDescription, tblPrChange`

**`w:tcPr`**: `cnfStyle, tcW, gridSpan, hMerge, vMerge, tcBorders, shd, noWrap, tcMar, textDirection, tcFitText, vAlign,
hideMark, headers, cellIns, cellDel, cellMerge, tcPrChange`

**`w:sectPr`**: `headerReference*, footerReference*, footnotePr, endnotePr, type, pgSz, pgMar, paperSrc, pgBorders, lnNumType,
pgNumType, cols, formProt, vAlign, noEndnote, titlePg, textDirection, bidi, rtlGutter, docGrid, printerSettings, sectPrChange`

**`w:trPr`**: elementos comuns em qualquer ordem, seguidos de `ins, del, trPrChange`.

## Texto e runs

- Use `xml:space="preserve"` em `w:t` com espaço nas pontas, ou o espaço some.
- O Word divide o texto em runs arbitrariamente (`O contr` + `ato`). Nunca presuma que uma frase está em um único `w:t`.
- `w:rPr` é sempre o primeiro filho de `w:r`; `w:pPr` é sempre o primeiro filho de `w:p`.
- Tabulação é `<w:tab/>` e quebra é `<w:br/>`: caracteres `\t` e `\n` dentro de `w:t` não funcionam como esperado.
- Escape `&`, `<` e `>`. Caracteres de controle são proibidos em XML 1.0.

## Revisões

- Dentro de `w:del`, use `w:delText` e `w:delInstrText` no lugar de `w:t` e `w:instrText`.
- Toda revisão precisa de `w:id` e `w:author`; use `w:date` em UTC. Mantenha os `w:id` únicos.
- Parágrafo inserido: `<w:ins/>` em `w:pPr/w:rPr` **e** os runs em `w:ins`. Excluído: o mesmo com `w:del`.
- Linha inserida ou excluída: `<w:ins/>` ou `<w:del/>` em `w:trPr`, além do conteúdo marcado.
- Excluir texto de uma inserção de terceiro: `<w:ins ...><w:del ...><w:r>...`. Não aninhe `w:ins` em `w:ins`.

## Campos

- Um campo complexo pode atravessar parágrafos: o sumário começa em um parágrafo e termina vários depois. Remova do
  `begin` até o `end` por inteiro.
- O texto entre `separate` e `end` é resultado em cache: mudar só ele não muda a origem.
- **Não use `w:updateFields`**: o Word exibe um aviso ao abrir. Peça ao usuário para pressionar `F9`.

## Tabelas

- `w:tblGrid` é obrigatório, e a soma de `gridSpan`, `gridBefore` e `gridAfter` de cada linha deve bater com ele.
- Mesclagem vertical: `<w:vMerge w:val="restart"/>` inicia e `<w:vMerge/>` continua. A célula de continuação ainda
  precisa de `w:p`. Ao remover a linha inicial, transfira o `restart` e o conteúdo para a seguinte.
- Duas tabelas sem parágrafo entre elas são fundidas pelo Word.
- A tabela no fim do corpo precisa de um parágrafo depois dela.
- Ao clonar linha, remova `w14:paraId` e `w14:textId` e não leve `w:tblHeader` sem intenção.

## Seções

- O `w:sectPr` direto em `w:body` (último filho) define a **última** seção. Cada `w:sectPr` em `w:pPr` encerra a seção
  naquele parágrafo.
- Remover um parágrafo com `sectPr` faz o conteúdo anterior adotar a configuração da seção **seguinte**: margens,
  orientação, cabeçalhos e rodapés mudam. Confirme com o usuário.
- Ao clonar parágrafo, nunca copie `sectPr`, ou uma seção nova surge.

## Numeração

- O item usa `w:numPr` direto ou herdado do estilo. Para continuar a lista, repita o mesmo `numId`.
- Para reiniciar, crie um `w:num` novo com `w:lvlOverride/w:startOverride`. Trocar o `abstractNum` altera todas as listas
  que o usam.
- Em `numbering.xml`, todos os `w:abstractNum` vêm **antes** de todos os `w:num`.
- `numId="0"` significa "sem numeração".
- Não digite "1." no texto: quebra a renumeração automática.

## Imagens e relacionamentos

Uma imagem nova exige quatro coisas coerentes:

1. o arquivo em `word/media/`;
2. `<Default Extension="png" ContentType="image/png"/>` em `[Content_Types].xml`;
3. um `Relationship` em `word/_rels/document.xml.rels` (ou no `.rels` do cabeçalho ou rodapé, que têm os seus);
4. `r:embed` apontando para ele, com `wp:docPr id` único em todo o documento.

Tamanhos usam EMU: 1 cm = 360000. Mantenha `wp:extent` e `a:ext` iguais.

## Comentários e notas

- Comentário: `commentRangeStart`, `commentRangeEnd`, um run com `commentReference` e `w:comment` em `comments.xml`.
  Ao remover, retire também as entradas de mesmo `paraId` em `commentsExtended.xml`, `commentsIds.xml` e
  `commentsExtensible.xml`, quando existirem.
- `footnotes.xml` tem separadores com `id` `-1` e `0` (ou `0` e `1`), que são referenciados por `settings.xml`. Não os remova.

## Compatibilidade entre aplicativos

- **`mc:AlternateContent`**: o Word usa `mc:Choice`; outros aplicativos podem usar `mc:Fallback`. Edite as duas cópias.
- **LibreOffice**: dependente de renderização real; use `render_docx.py`.
- **Google Docs**: não há como testar localmente. Importa bem estilos, listas, tabelas e imagens PNG/JPEG. Formas, caixas de
  texto, `EMF`/`WMF` e seções complexas podem mudar de aparência, inclusive no **original**. Compare com o original antes
  de atribuir a diferença à edição.

## Tarefas que só a rota manual cobre

| Tarefa | Cuidado principal |
|---|---|
| Fundir duas seções | Decidir qual configuração prevalece e conferir cabeçalhos |
| Remover o sumário inteiro | Do `begin` ao `end`; costuma estar em `w:sdt` |
| Excluir um comentário | Âncoras, `comments.xml` e partes estendidas |
| Aceitar ou rejeitar revisões a pedido | Tratar `w:ins`, `w:del`, `w:rPrChange`, `w:pPrChange` e marcas de parágrafo |
| Alterar um estilo a pedido | Afeta todos os parágrafos que o usam; rode o `compare` com `--allow word/styles.xml` |
| Editar dentro de revisão pendente | Dividir o `w:ins` em dois e colocar a nova revisão entre eles |
