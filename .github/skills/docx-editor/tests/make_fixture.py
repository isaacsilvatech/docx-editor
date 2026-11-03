"""Gera um .docx sintético com os elementos que costumam quebrar em edições.

O arquivo é montado à mão, em XML, para controlar exatamente cada caso difícil:
texto dividido em vários runs, numeração, campos, sumário, notas, comentários,
indicadores, seções, mesclagens, caixa de texto com mc:Fallback, revisões e
namespaces em mc:Ignorable. Nenhum binário é versionado.
"""
from __future__ import annotations

import struct
import zipfile
import zlib

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = (
    f'xmlns:w="{W}" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'mc:Ignorable="w14 w15"'
)
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'


def png_bytes(width=4, height=2):
    def chunk(tag, data):
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)
    raw = b"".join(b"\x00" + b"\x20\x60\xa0" * width for _ in range(height))
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw))
            + chunk(b"IEND", b""))


def inline_image(rid, docpr, name):
    return (
        f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="914400" cy="457200"/><wp:docPr id="{docpr}" name="{name}"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{docpr}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    )


TEXTBOX = (
    '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" '
    'simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
    '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
    '<wp:extent cx="914400" cy="457200"/><wp:wrapNone/><wp:docPr id="3" name="Caixa"/>'
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
    '<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>'
    '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Texto na caixa</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
    '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>'
    '<mc:Fallback><w:pict><v:shape id="s1" style="width:72pt;height:36pt"><v:textbox><w:txbxContent>'
    '<w:p><w:r><w:t>Texto na caixa</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>'
    '</mc:AlternateContent></w:r></w:p>'
)


def field(instr, result):
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r><w:t>{result}</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


BODY = (
    # 0-2: sumário cujo campo atravessa parágrafos
    '<w:p><w:pPr><w:pStyle w:val="Sumario1"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>'
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Introdução</w:t></w:r></w:p>'
    '<w:p><w:pPr><w:pStyle w:val="Sumario1"/></w:pPr><w:r><w:t>Resultados</w:t></w:r></w:p>'
    '<w:p><w:pPr><w:pStyle w:val="Sumario1"/></w:pPr><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    # 3: título com indicador
    '<w:p w14:paraId="1A2B3C01" w14:textId="77777777" w:rsidR="00AA0001"><w:pPr><w:pStyle w:val="Titulo1"/></w:pPr>'
    '<w:bookmarkStart w:id="1" w:name="_Toc100"/><w:r><w:t>Introdução</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>'
    # 4: corpo com texto dividido entre runs de formatos diferentes
    '<w:p w14:paraId="1A2B3C02" w14:textId="77777777"><w:pPr><w:jc w:val="both"/></w:pPr>'
    '<w:r w:rsidR="00AA0002"><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">O contr</w:t></w:r>'
    '<w:proofErr w:type="spellStart"/>'
    '<w:r w:rsidR="00AA0003"><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">ato vigente </w:t></w:r>'
    '<w:proofErr w:type="spellEnd"/>'
    '<w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>vence em 2023</w:t></w:r>'
    '<w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve"> e pode ser renovado por acordo entre as partes.</w:t></w:r></w:p>'
    # 5: nota de rodapé e comentário
    '<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:commentRangeStart w:id="0"/>'
    '<w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="22"/></w:rPr><w:t>O prazo segue a norma</w:t></w:r>'
    '<w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="RefComentario"/></w:rPr><w:commentReference w:id="0"/></w:r>'
    '<w:r><w:rPr><w:rStyle w:val="RefNota"/></w:rPr><w:footnoteReference w:id="1"/></w:r>'
    '<w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve"> aplicável ao caso.</w:t></w:r></w:p>'
    # 6-8: lista numerada
    '<w:p><w:pPr><w:pStyle w:val="ItemLista"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Primeiro item;</w:t></w:r></w:p>'
    '<w:p><w:pPr><w:pStyle w:val="ItemLista"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Segundo item;</w:t></w:r></w:p>'
    '<w:p><w:pPr><w:pStyle w:val="ItemLista"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Subitem.</w:t></w:r></w:p>'
    # 9: hiperlink
    '<w:p><w:r><w:t xml:space="preserve">Veja </w:t></w:r><w:hyperlink r:id="rId20"><w:r><w:rPr><w:rStyle w:val="Link"/></w:rPr><w:t>o portal oficial</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> para detalhes.</w:t></w:r></w:p>'
    # 10: imagem + texto no mesmo run
    f'<w:p><w:r><w:t xml:space="preserve">Figura: </w:t></w:r>{inline_image("rId21", 1, "Figura 1")}<w:r><w:t xml:space="preserve"> legenda da figura.</w:t></w:r></w:p>'
    # 11: quebra de seção carregada por parágrafo
    '<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId22"/><w:footerReference w:type="default" r:id="rId23"/>'
    '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1701" w:bottom="1417" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:pPr>'
    '<w:r><w:t>Fim da primeira seção.</w:t></w:r></w:p>'
    # 12: título; depois tabela (Tabela 0) com cabeçalho e mesclagem vertical
    '<w:p><w:pPr><w:pStyle w:val="Titulo1"/></w:pPr><w:r><w:t>Resultados</w:t></w:r></w:p>'
    '<w:tbl><w:tblPr><w:tblStyle w:val="TabelaPadrao"/><w:tblW w:w="5000" w:type="pct"/><w:tblLook w:val="04A0"/></w:tblPr>'
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
    '<w:tr><w:trPr><w:tblHeader/></w:trPr>'
    '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Grupo</w:t></w:r></w:p></w:tc>'
    '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Item</w:t></w:r></w:p></w:tc>'
    '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Valor</w:t></w:r></w:p></w:tc></w:tr>'
    '<w:tr w14:paraId="1A2B3C10"><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Grupo A</w:t></w:r></w:p></w:tc>'
    '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:bookmarkStart w:id="2" w:name="valor_alfa"/><w:r><w:t>Alfa</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p></w:tc>'
    '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:color w:val="1F4E79"/></w:rPr><w:t>10,5</w:t></w:r></w:p></w:tc></w:tr>'
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>'
    '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Beta</w:t></w:r></w:p></w:tc>'
    '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:color w:val="1F4E79"/></w:rPr><w:t>20,0</w:t></w:r></w:p></w:tc></w:tr>'
    '</w:tbl>'
    # parágrafo com referência cruzada (REF) a um indicador
    f'<w:p><w:r><w:t xml:space="preserve">Conforme o valor </w:t></w:r>{field("REF valor_alfa", "Alfa")}<w:r><w:t>.</w:t></w:r></w:p>'
    # controle de conteúdo vinculado a dados
    '<w:p><w:sdt><w:sdtPr><w:alias w:val="Título"/><w:dataBinding w:prefixMappings="" w:xpath="/x" w:storeItemID="{00000000-0000-0000-0000-000000000000}"/></w:sdtPr>'
    '<w:sdtContent><w:r><w:t>Valor vinculado</w:t></w:r></w:sdtContent></w:sdt></w:p>'
    # caixa de texto com mc:Fallback
    + TEXTBOX +
    # revisões pendentes de outro autor
    '<w:p><w:r><w:t xml:space="preserve">Texto com </w:t></w:r>'
    '<w:ins w:id="10" w:author="Maria" w:date="2024-01-01T00:00:00Z"><w:r><w:t>inserção</w:t></w:r></w:ins>'
    '<w:del w:id="11" w:author="Maria" w:date="2024-01-01T00:00:00Z"><w:r><w:delText>removido</w:delText></w:r></w:del>'
    '<w:r><w:t xml:space="preserve"> pendente.</w:t></w:r></w:p>'
    # parágrafo simples removível, separador
    '<w:p><w:r><w:t>Parágrafo descartável.</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>Último parágrafo.</w:t></w:r></w:p>'
    '<w:sectPr><w:headerReference w:type="default" r:id="rId22"/><w:footerReference w:type="default" r:id="rId23"/>'
    '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1417" w:right="1701" w:bottom="1417" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
)

STYLES = (
    DECL + f'<w:styles {NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="22"/><w:lang w:val="pt-BR"/></w:rPr></w:rPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="character" w:default="1" w:styleId="FonteParagrafoPadrao"><w:name w:val="Default Paragraph Font"/></w:style>'
    '<w:style w:type="table" w:default="1" w:styleId="TabelaNormal"><w:name w:val="Normal Table"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Titulo1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="Titulo2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="Sumario1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="ItemLista"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style>'
    '<w:style w:type="character" w:styleId="Link"><w:name w:val="Hyperlink"/><w:rPr><w:u w:val="single"/></w:rPr></w:style>'
    '<w:style w:type="character" w:styleId="Forte"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>'
    '<w:style w:type="character" w:styleId="RefNota"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
    '<w:style w:type="character" w:styleId="RefComentario"><w:name w:val="annotation reference"/></w:style>'
    '<w:style w:type="table" w:styleId="TabelaPadrao"><w:name w:val="Table Grid"/><w:basedOn w:val="TabelaNormal"/></w:style>'
    '</w:styles>'
)
NUMBERING = (
    DECL + f'<w:numbering {NS}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl>'
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/></w:lvl>'
    '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
)
FOOTNOTES = (
    DECL + f'<w:footnotes {NS}>'
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
    '<w:footnote w:id="1"><w:p><w:r><w:rPr><w:rStyle w:val="RefNota"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Nota sobre a norma.</w:t></w:r></w:p></w:footnote>'
    '</w:footnotes>'
)
COMMENTS = (
    DECL + f'<w:comments {NS}><w:comment w:id="0" w:author="Revisor" w:date="2024-01-01T00:00:00Z" w:initials="R">'
    '<w:p w14:paraId="1A2B3C20"><w:r><w:rPr><w:rStyle w:val="RefComentario"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>Conferir a norma.</w:t></w:r></w:p></w:comment></w:comments>'
)
HEADER = DECL + f'<w:hdr {NS}><w:p><w:r><w:t>Empresa Exemplo</w:t></w:r></w:p></w:hdr>'
FOOTER = DECL + (
    f'<w:ftr {NS}><w:p><w:r><w:t xml:space="preserve">Página </w:t></w:r>' + field("PAGE", "1") + '</w:p></w:ftr>'
)
SETTINGS_PLAIN = DECL + f'<w:settings {NS}><w:zoom w:percent="100"/><w:rsids><w:rsidRoot w:val="00AA0001"/></w:rsids></w:settings>'
SETTINGS_TRACKED = DECL + f'<w:settings {NS}><w:zoom w:percent="100"/><w:trackRevisions/></w:settings>'
CORE = DECL + (
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Documento de teste</dc:title>'
    '<dc:creator>Autora Original</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created></cp:coreProperties>'
)
CONTENT_TYPES = DECL + (
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
    '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    '</Types>'
)
ROOT_RELS = DECL + (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '</Relationships>'
)
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DOC_RELS = DECL + (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{R}/styles" Target="styles.xml"/>'
    f'<Relationship Id="rId2" Type="{R}/numbering" Target="numbering.xml"/>'
    f'<Relationship Id="rId3" Type="{R}/settings" Target="settings.xml"/>'
    f'<Relationship Id="rId4" Type="{R}/footnotes" Target="footnotes.xml"/>'
    f'<Relationship Id="rId5" Type="{R}/comments" Target="comments.xml"/>'
    f'<Relationship Id="rId20" Type="{R}/hyperlink" Target="https://example.com/" TargetMode="External"/>'
    f'<Relationship Id="rId21" Type="{R}/image" Target="media/image1.png"/>'
    f'<Relationship Id="rId22" Type="{R}/header" Target="header1.xml"/>'
    f'<Relationship Id="rId23" Type="{R}/footer" Target="footer1.xml"/>'
    '</Relationships>'
)


def build(path, tracked=False):
    document = DECL + f'<w:document {NS}><w:body>{BODY}</w:body></w:document>'
    entries = [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("word/document.xml", document),
        ("word/_rels/document.xml.rels", DOC_RELS),
        ("word/styles.xml", STYLES),
        ("word/numbering.xml", NUMBERING),
        ("word/settings.xml", SETTINGS_TRACKED if tracked else SETTINGS_PLAIN),
        ("word/footnotes.xml", FOOTNOTES),
        ("word/comments.xml", COMMENTS),
        ("word/header1.xml", HEADER),
        ("word/footer1.xml", FOOTER),
        ("word/media/image1.png", png_bytes()),
        ("docProps/core.xml", CORE),
    ]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in entries:
            zf.writestr(name, data.encode("utf-8") if isinstance(data, str) else data)
    return path


if __name__ == "__main__":
    import sys
    build(sys.argv[1], tracked=len(sys.argv) > 2 and sys.argv[2] == "tracked")
