"""Biblioteca compartilhada da skill docx-editor.

Trabalha direto no pacote OOXML: só as partes alteradas são reserializadas, e
tudo o mais volta byte a byte, na ordem original do ZIP. É isso que garante que
"o que não foi pedido para mudar" continue idêntico.
"""
from __future__ import annotations

import fnmatch
import os
import posixpath
import re
import sys
import tempfile
import zipfile
from collections import Counter
from dataclasses import dataclass, field

try:
    from lxml import etree
except ImportError:  # pragma: no cover
    sys.exit(
        "ERRO: dependência 'lxml' ausente. Instale com: "
        "python -m pip install -r <pasta-da-skill>/requirements.txt"
    )

# ---------------------------------------------------------------------------
# Namespaces
# ---------------------------------------------------------------------------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_STRICT_NS = "http://purl.oclc.org/ooxml/wordprocessingml/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
XML_NS = "http://www.w3.org/XML/1998/namespace"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

RT_OFFICE_DOC = R_NS + "/officeDocument"
RT_IMAGE = R_NS + "/image"

XML_SPACE = "{%s}space" % XML_NS
MC_FALLBACK = "{%s}Fallback" % MC_NS
MC_CHOICE = "{%s}Choice" % MC_NS
MC_ALT = "{%s}AlternateContent" % MC_NS
MC_IGNORABLE = "{%s}Ignorable" % MC_NS
W14_PARA_ID = "{%s}paraId" % W14_NS
W14_TEXT_ID = "{%s}textId" % W14_NS


def w(tag: str) -> str:
    """Nome qualificado de um elemento/atributo no namespace w:."""
    return "{%s}%s" % (W_NS, tag)


def wval(el, default=None):
    """Valor de w:val de um elemento (ou default)."""
    return default if el is None else el.get(w("val"), default)


def local(el) -> str:
    tag = el.tag
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def is_w(el) -> bool:
    return isinstance(el.tag, str) and el.tag.startswith("{%s}" % W_NS)


def is_on(el) -> bool:
    """Interpreta um elemento booleano OOXML (<w:b/>, <w:b w:val="0"/>...)."""
    if el is None:
        return False
    return el.get(w("val"), "1").lower() not in ("0", "false", "off")


# ---------------------------------------------------------------------------
# Ordem de filhos exigida pelo esquema (ECMA-376, transitional).
# Inserir fora de ordem é a causa nº 1 de "o Word encontrou conteúdo ilegível".
# ---------------------------------------------------------------------------
PPR_ORDER = (
    "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr "
    "suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap "
    "overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd "
    "snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc "
    "textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle "
    "rPr sectPr pPrChange"
).split()
RPR_ORDER = (
    "ins del moveFrom moveTo rStyle rFonts b bCs i iCs caps smallCaps strike "
    "dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden "
    "color spacing w kern position sz szCs highlight u effect bdr shd fitText "
    "vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange"
).split()
TBLPR_ORDER = (
    "tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize "
    "tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd tblLayout "
    "tblCellMar tblLook tblCaption tblDescription tblPrChange"
).split()
TCPR_ORDER = (
    "cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar "
    "textDirection tcFitText vAlign hideMark headers cellIns cellDel cellMerge "
    "tcPrChange"
).split()
SECTPR_ORDER = (
    "headerReference footerReference footnotePr endnotePr type pgSz pgMar "
    "paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote "
    "titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange"
).split()
TRPR_TAIL_ORDER = ["ins", "del", "trPrChange"]


def ordered_insert(parent, child, order):
    """Insere `child` em `parent` respeitando a ordem do esquema."""
    name = local(child)
    rank = order.index(name) if name in order else len(order)
    for idx, sibling in enumerate(parent):
        sib = local(sibling)
        if is_w(sibling) and sib in order and order.index(sib) > rank:
            parent.insert(idx, child)
            return child
    parent.append(child)
    return child


# ---------------------------------------------------------------------------
# Erros
# ---------------------------------------------------------------------------
class DocxError(Exception):
    """Erro esperado (entrada inválida, operação insegura) com mensagem útil."""


def setup_stdio():
    """Evita UnicodeEncodeError em consoles Windows (cp1252)."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
        except (AttributeError, ValueError):
            pass


# ---------------------------------------------------------------------------
# Pacote
# ---------------------------------------------------------------------------
_PARSER_OPTS = dict(
    remove_blank_text=False,
    remove_comments=False,
    remove_pis=False,
    strip_cdata=False,
    resolve_entities=False,  # bloqueia XXE
    no_network=True,
    huge_tree=True,
)
_DECL_RE = re.compile(rb"^\s*<\?xml[^>]*\?>[ \t\r\n]*")
_DEFAULT_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
MAX_UNCOMPRESSED = 4 * 1024 ** 3

STORY_ROOTS = {w(x) for x in ("document", "hdr", "ftr", "footnotes", "endnotes", "comments")}


def parse_xml(data: bytes):
    return etree.fromstring(data, etree.XMLParser(**_PARSER_OPTS))


def serialize_xml(root, original: bytes | None = None) -> bytes:
    """Serializa preservando a declaração XML original quando ela é UTF-8."""
    decl = _DEFAULT_DECL
    if original is not None:
        stripped = original[3:] if original.startswith(b"\xef\xbb\xbf") else original
        match = _DECL_RE.match(stripped)
        if match and b"utf-8" in match.group(0).lower():
            decl = match.group(0)
    return decl + etree.tostring(root.getroottree(), encoding="UTF-8", xml_declaration=False)


class Package:
    """Um .docx em memória. Só partes marcadas como sujas são reserializadas."""

    def __init__(self, path):
        self.path = str(path)
        _check_signature(self.path)
        try:
            with zipfile.ZipFile(self.path) as zf:
                self.infos = zf.infolist()
                names = [i.filename for i in self.infos]
                dup = [n for n, c in Counter(names).items() if c > 1]
                if dup:
                    raise DocxError(f"ZIP com entradas duplicadas: {dup}")
                if sum(i.file_size for i in self.infos) > MAX_UNCOMPRESSED:
                    raise DocxError("Pacote descompactado excede 4 GiB; recusado por segurança.")
                self.data = {i.filename: zf.read(i.filename) for i in self.infos}
        except zipfile.BadZipFile as exc:
            raise DocxError(f"{self.path} não é um ZIP/.docx válido: {exc}") from exc
        self._trees = {}
        self.dirty = set()
        self.removed = set()
        if "[Content_Types].xml" not in self.data:
            raise DocxError(f"{self.path} não é OOXML: falta [Content_Types].xml")
        self.main_part = self._find_main_part()
        self._check_not_strict()

    def _find_main_part(self) -> str:
        if "_rels/.rels" in self.data:
            for rel in parse_xml(self.data["_rels/.rels"]):
                if rel.get("Type") == RT_OFFICE_DOC:
                    target = rel.get("Target", "").lstrip("/")
                    if target in self.data:
                        return target
        if "word/document.xml" in self.data:
            return "word/document.xml"
        raise DocxError(f"{self.path}: documento principal não encontrado")

    def _check_not_strict(self):
        root = self.xml(self.main_part)
        if root.tag.startswith("{%s}" % W_STRICT_NS):
            raise DocxError(
                "Documento em Strict Open XML (namespace purl.oclc.org). Não suportado: "
                "peça ao usuário para salvar como 'Documento do Word (*.docx)' transitional."
            )
        if root.tag != w("document"):
            raise DocxError(f"Parte principal inesperada: {root.tag}")

    # ---- partes -----------------------------------------------------------
    def names(self):
        return [i.filename for i in self.infos if i.filename not in self.removed]

    def has(self, name) -> bool:
        return name in self.data and name not in self.removed

    def xml(self, name):
        if name not in self._trees:
            if not self.has(name):
                raise DocxError(f"Parte inexistente no pacote: {name}")
            try:
                self._trees[name] = parse_xml(self.data[name])
            except etree.XMLSyntaxError as exc:
                raise DocxError(f"XML malformado em {name}: {exc}") from exc
        return self._trees[name]

    def touch(self, name):
        """Marca a parte para ser reserializada ao salvar."""
        self.xml(name)
        self.dirty.add(name)

    def add_part(self, name, data: bytes):
        if self.has(name):
            raise DocxError(f"Parte já existe: {name}")
        self.removed.discard(name)
        self.data[name] = data
        if not any(i.filename == name for i in self.infos):
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            self.infos.append(info)

    def remove_part(self, name):
        self.removed.add(name)
        self.dirty.discard(name)
        self._trees.pop(name, None)

    def story_parts(self):
        """Partes que contêm texto: principal, cabeçalhos, rodapés, notas, comentários."""
        out = [self.main_part]
        for name in self.names():
            if name == self.main_part or not re.match(r"^word/[^/]+\.xml$", name):
                continue
            try:
                if self.xml(name).tag in STORY_ROOTS:
                    out.append(name)
            except DocxError:
                continue
        return out

    def rendered(self, name) -> bytes:
        if name in self.dirty:
            return serialize_xml(self._trees[name], self.data.get(name))
        return self.data[name]

    def save(self, out_path):
        """Grava de forma atômica, mantendo ordem e metadados das entradas ZIP."""
        out_path = os.path.abspath(str(out_path))
        fd, tmp = tempfile.mkstemp(suffix=".docx", dir=os.path.dirname(out_path))
        os.close(fd)
        try:
            self.write_to(tmp)
            os.replace(tmp, out_path)
        except BaseException:
            if os.path.exists(tmp):
                os.unlink(tmp)
            raise

    def write_to(self, path):
        with zipfile.ZipFile(path, "w", allowZip64=True) as zout:
            for info in self.infos:
                if info.filename in self.removed:
                    continue
                zout.writestr(clone_zipinfo(info), self.rendered(info.filename))

    # ---- relacionamentos ----------------------------------------------------
    def rels_name(self, part) -> str:
        folder, base = posixpath.split(part)
        return posixpath.join(folder, "_rels", base + ".rels")

    def rels(self, part):
        """Lista de elementos <Relationship> de uma parte (vazia se não houver .rels)."""
        name = self.rels_name(part)
        return list(self.xml(name)) if self.has(name) else []

    def rel_target(self, part, rel) -> str:
        return resolve_target(part, rel.get("Target", ""))

    def add_relationship(self, part, rel_type, target) -> str:
        name = self.rels_name(part)
        if not self.has(name):
            root = etree.Element("{%s}Relationships" % RELS_NS, nsmap={None: RELS_NS})
            self.add_part(name, serialize_xml(root))
        root = self.xml(name)
        used = {r.get("Id") for r in root}
        n = 1
        while f"rId{n}" in used:
            n += 1
        rid = f"rId{n}"
        etree.SubElement(
            root,
            "{%s}Relationship" % RELS_NS,
            {"Id": rid, "Type": rel_type, "Target": target},
        )
        self.touch(name)
        return rid

    def ensure_default_content_type(self, ext, content_type):
        root = self.xml("[Content_Types].xml")
        for el in root:
            if local(el) == "Default" and el.get("Extension", "").lower() == ext.lower():
                return
        default = etree.Element("{%s}Default" % CT_NS, {"Extension": ext, "ContentType": content_type})
        root.insert(0, default)
        self.touch("[Content_Types].xml")

    def remove_override(self, part):
        root = self.xml("[Content_Types].xml")
        for el in list(root):
            if local(el) == "Override" and el.get("PartName") == "/" + part:
                root.remove(el)
                self.touch("[Content_Types].xml")

    # ---- settings -----------------------------------------------------------
    def settings(self):
        return self.xml("word/settings.xml") if self.has("word/settings.xml") else None

    def tracks_revisions(self) -> bool:
        settings = self.settings()
        return settings is not None and is_on(settings.find(w("trackRevisions")))


def clone_zipinfo(info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    new = zipfile.ZipInfo(info.filename, date_time=info.date_time)
    new.compress_type = info.compress_type
    new.external_attr = info.external_attr
    new.create_system = info.create_system
    new.comment = info.comment
    return new


def _check_signature(path):
    if not os.path.isfile(path):
        raise DocxError(f"Arquivo não encontrado: {path}")
    with open(path, "rb") as fh:
        head = fh.read(8)
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        raise DocxError(
            f"{path} é um contêiner OLE: .doc legado ou .docx protegido por senha. "
            "Peça ao usuário uma cópia .docx sem senha; não tente contornar a proteção."
        )


def resolve_target(part, target) -> str:
    target = target.split("#", 1)[0]
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def glob_match(name, patterns) -> bool:
    # A comparação exata vem primeiro: "[Content_Types].xml" seria lido como classe de caracteres.
    return any(name == p or fnmatch.fnmatchcase(name, p) for p in patterns)


# ---------------------------------------------------------------------------
# Parágrafos e tabelas
# ---------------------------------------------------------------------------
def in_fallback(el) -> bool:
    return any(a.tag == MC_FALLBACK for a in el.iterancestors())


def iter_paragraphs(root):
    """Parágrafos em ordem de documento, sem as duplicatas de mc:Fallback."""
    for p in root.iter(w("p")):
        if not in_fallback(p):
            yield p


def iter_tables(root):
    for tbl in root.iter(w("tbl")):
        if not in_fallback(tbl):
            yield tbl


def field_depths(root):
    """Profundidade de campos complexos no início de cada parágrafo.

    Um campo pode atravessar parágrafos (o sumário é o caso clássico), logo o
    estado precisa ser calculado para a parte inteira.
    """
    depths = {}
    depth = 0
    for el in root.iter(w("p"), w("fldChar")):
        if in_fallback(el):
            continue
        if el.tag == w("p"):
            depths[el] = depth
        else:
            kind = el.get(w("fldCharType"))
            if kind == "begin":
                depth += 1
            elif kind == "end":
                depth = max(0, depth - 1)
    return depths


_DESCEND = {w(x) for x in ("hyperlink", "ins", "smartTag", "customXml", "moveTo", "dir", "bdo")}
_OPAQUE_INLINE = {MC_ALT, "{%s}oMath" % M_NS, "{%s}oMathPara" % M_NS}
_SIMPLE_CHARS = {
    w("tab"): "\t",
    w("cr"): "\n",
    w("noBreakHyphen"): "‑",
    w("softHyphen"): "­",
}
BARRIER = "￼"


@dataclass(eq=False)
class Atom:
    """Unidade de conteúdo do parágrafo.

    `text is None` indica uma barreira: objeto que não é texto (imagem, nota,
    marcador de campo, quebra de página). Uma busca nunca atravessa barreiras,
    de modo que substituir texto não apaga esses objetos por acidente.
    """

    run: object
    child: object
    text: str | None
    start: int = 0
    in_field: bool = False

    @property
    def length(self) -> int:
        return len(self.text) if self.text else 0


class ParagraphText:
    """Mapa do texto visível de um parágrafo para os elementos que o contêm."""

    def __init__(self, p, start_depth: int = 0):
        self.p = p
        self.atoms: list[Atom] = []
        self._depth = start_depth
        self._walk(p, False)
        pos = 0
        for atom in self.atoms:
            atom.start = pos
            pos += atom.length
        self.text = "".join(a.text for a in self.atoms if a.text)

    def _walk(self, parent, simple_field):
        for child in parent:
            tag = child.tag
            if tag == w("r"):
                self._run(child, simple_field)
            elif tag in _DESCEND:
                self._walk(child, simple_field)
            elif tag == w("sdt"):
                content = child.find(w("sdtContent"))
                if content is not None:
                    self._walk(content, simple_field)
            elif tag == w("fldSimple"):
                self._walk(child, True)
            elif tag in _OPAQUE_INLINE:
                self.atoms.append(Atom(None, child, None))

    def _run(self, run, simple_field):
        for child in run:
            tag = child.tag
            if tag == w("rPr"):
                continue
            field_now = simple_field or self._depth > 0
            if tag == w("t"):
                self.atoms.append(Atom(run, child, child.text or "", in_field=field_now))
            elif tag in _SIMPLE_CHARS:
                self.atoms.append(Atom(run, child, _SIMPLE_CHARS[tag], in_field=field_now))
            elif tag == w("br") and child.get(w("type"), "textWrapping") == "textWrapping":
                self.atoms.append(Atom(run, child, "\n", in_field=field_now))
            elif tag == w("lastRenderedPageBreak"):
                self.atoms.append(Atom(run, child, "", in_field=field_now))
            else:
                if tag == w("fldChar"):
                    kind = child.get(w("fldCharType"))
                    if kind == "begin":
                        self._depth += 1
                    elif kind == "end":
                        self._depth = max(0, self._depth - 1)
                self.atoms.append(Atom(run, child, None, in_field=field_now))

    @property
    def searchable(self):
        """(texto com BARRIER nas barreiras, mapa posição-busca -> posição-visível)."""
        chars, mapping = [], []
        for atom in self.atoms:
            if atom.text is None:
                chars.append(BARRIER)
                mapping.append(atom.start)
            else:
                for offset, char in enumerate(atom.text):
                    chars.append(char)
                    mapping.append(atom.start + offset)
        mapping.append(len(self.text))
        return "".join(chars), mapping

    def text_atoms_in(self, start, end):
        return [a for a in self.atoms if a.length and a.start < end and a.start + a.length > start]

    def has_barrier(self, tags) -> bool:
        return any(a.text is None and a.child.tag in tags for a in self.atoms)


def visible_text(p, depth=0) -> str:
    return ParagraphText(p, depth).text


def paragraph_style(p):
    ppr = p.find(w("pPr"))
    return wval(ppr.find(w("pStyle"))) if ppr is not None else None


def paragraph_numbering(p):
    ppr = p.find(w("pPr"))
    if ppr is None:
        return None
    num = ppr.find(w("numPr"))
    if num is None:
        return None
    return wval(num.find(w("numId"))), wval(num.find(w("ilvl")), "0")


def has_section_break(p) -> bool:
    ppr = p.find(w("pPr"))
    return ppr is not None and ppr.find(w("sectPr")) is not None


# ---------------------------------------------------------------------------
# Assinaturas de formatação (independentes de rsid e de serialização)
# ---------------------------------------------------------------------------
_NOISE_ATTRS = re.compile(r"\}(rsid\w*|paraId|textId)$")
_NOISE_TAGS = {w("lastRenderedPageBreak"), w("proofErr")}


def canon(el, skip_tags=()) -> str:
    """Forma canônica de um elemento, ignorando rsids e paraIds."""
    if el is None or not isinstance(el.tag, str):
        return ""
    attrs = sorted((k, v) for k, v in el.attrib.items() if not _NOISE_ATTRS.search(k))
    kids = [canon(c, skip_tags) for c in el if isinstance(c.tag, str) and c.tag not in skip_tags and c.tag not in _NOISE_TAGS]
    raw = el.text or ""
    text = raw if local(el) in ("t", "instrText", "delText") else raw.strip()
    return f"<{local(el)} {attrs}>{text}{''.join(kids)}</>"


_MARK_REVISION = {w("ins"), w("del"), w("moveFrom"), w("moveTo")}


def ppr_signature(p) -> str:
    """Assinatura do parágrafo sem marcas de revisão e sem quebra de seção."""
    ppr = p.find(w("pPr"))
    sig = canon(ppr, skip_tags=_MARK_REVISION | {w("sectPr"), w("pPrChange"), w("rPrChange")})
    # Sem a marca de revisão, um rPr/pPr vazio equivale a nenhum.
    sig = sig.replace("<rPr []></>", "")
    return "" if sig == "<pPr []></>" else sig


def has_deleted_mark(p) -> bool:
    """A marca de parágrafo foi excluída com revisão (w:pPr/w:rPr/w:del)."""
    ppr = p.find(w("pPr"))
    mark = ppr.find(w("rPr")) if ppr is not None else None
    return mark is not None and mark.find(w("del")) is not None


def rpr_signature(run) -> str:
    if run is None:
        return ""
    return canon(run.find(w("rPr")), skip_tags={w("rPrChange")})


def char_formats(pt: ParagraphText):
    """Assinatura de rPr de cada caractere visível."""
    out = []
    for atom in pt.atoms:
        if atom.length:
            sig = rpr_signature(atom.run)
            out.extend([sig] * atom.length)
    return out


def human_rpr(rpr) -> str:
    """Resumo legível de um w:rPr (para o inspect)."""
    if rpr is None:
        return ""
    bits = []
    for child in rpr:
        name = local(child)
        if name in ("lang", "noProof", "rPrChange", "ins", "del", "bCs", "iCs", "szCs"):
            continue
        if name == "rFonts":
            font = child.get(w("ascii")) or child.get(w("asciiTheme")) or child.get(w("hAnsi"))
            if font:
                bits.append(f"font={font}")
        elif name == "sz":
            try:
                bits.append(f"{int(wval(child)) / 2:g}pt")
            except (TypeError, ValueError):
                bits.append("sz=?")
        elif name in ("b", "i", "caps", "smallCaps", "strike", "vanish"):
            bits.append(name if is_on(child) else f"!{name}")
        elif wval(child) is not None:
            bits.append(f"{name}={wval(child)}")
        else:
            bits.append(name)
    return " ".join(bits)


# ---------------------------------------------------------------------------
# Estilos
# ---------------------------------------------------------------------------
class Styles:
    def __init__(self, pkg: Package):
        self.by_id = {}
        self.defaults = {}
        self.root = None
        if pkg.has("word/styles.xml"):
            self.root = pkg.xml("word/styles.xml")
            for style in self.root.findall(w("style")):
                sid = style.get(w("styleId"))
                if sid:
                    self.by_id[sid] = style
                if style.get(w("default")) in ("1", "true"):
                    self.defaults[style.get(w("type"))] = sid

    def name_of(self, sid) -> str | None:
        style = self.by_id.get(sid)
        return wval(style.find(w("name"))) if style is not None else None

    def type_of(self, sid):
        style = self.by_id.get(sid)
        return style.get(w("type")) if style is not None else None

    def resolve(self, wanted, kind="paragraph") -> str | None:
        """Aceita styleId ou nome de exibição, sem diferenciar maiúsculas.

        IDs costumam ser localizados (Ttulo1 no Word em português), então o
        agente pode tanto usar o ID visto no inspect quanto o nome.
        """
        if wanted in self.by_id and self.type_of(wanted) == kind:
            return wanted
        target = str(wanted).casefold()
        for sid, style in self.by_id.items():
            if style.get(w("type")) != kind:
                continue
            name = wval(style.find(w("name")), "")
            if sid.casefold() == target or name.casefold() == target:
                return sid
        return None

    def is_heading(self, sid) -> bool:
        seen = set()
        while sid and sid in self.by_id and sid not in seen:
            seen.add(sid)
            style = self.by_id[sid]
            name = (wval(style.find(w("name")), "") or "").casefold()
            if re.match(r"(heading|t[ií]tulo)\s*\d", name):
                return True
            ppr = style.find(w("pPr"))
            if ppr is not None and ppr.find(w("outlineLvl")) is not None:
                return True
            sid = wval(style.find(w("basedOn")))
        return False

    def numbering_of(self, sid):
        """(numId, ilvl) herdados pelo estilo, seguindo basedOn; ou None."""
        seen = set()
        while sid and sid in self.by_id and sid not in seen:
            seen.add(sid)
            style = self.by_id[sid]
            ppr = style.find(w("pPr"))
            num = ppr.find(w("numPr")) if ppr is not None else None
            if num is not None and num.find(w("numId")) is not None:
                return wval(num.find(w("numId"))), wval(num.find(w("ilvl")), "0")
            sid = wval(style.find(w("basedOn")))
        return None

    def effective_run_props(self, sid) -> dict:
        """Fonte/tamanho/cor/negrito efetivos, seguindo basedOn e docDefaults."""
        chain = []
        seen = set()
        while sid and sid in self.by_id and sid not in seen:
            seen.add(sid)
            chain.append(self.by_id[sid])
            sid = wval(self.by_id[sid].find(w("basedOn")))
        sources = []
        if self.root is not None:
            dd = self.root.find(w("docDefaults"))
            if dd is not None:
                sources.append(dd.find("%s/%s" % (w("rPrDefault"), w("rPr"))))
        sources += [s.find(w("rPr")) for s in reversed(chain)]
        props: dict = {}
        for rpr in sources:
            if rpr is None:
                continue
            fonts = rpr.find(w("rFonts"))
            if fonts is not None:
                font = fonts.get(w("ascii")) or fonts.get(w("asciiTheme"))
                if font:
                    props["font"] = font
            size = rpr.find(w("sz"))
            if size is not None and wval(size):
                props["size_pt"] = int(wval(size)) / 2
            color = rpr.find(w("color"))
            if color is not None:
                props["color"] = wval(color)
            for tag, key in (("b", "bold"), ("i", "italic"), ("caps", "caps")):
                el = rpr.find(w(tag))
                if el is not None:
                    props[key] = is_on(el)
        return props


# ---------------------------------------------------------------------------
# Validação estrutural
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class Finding:
    severity: str  # error | warning | info
    code: str
    part: str
    detail: str
    where: str = field(default="", compare=False)

    @property
    def key(self):
        return (self.code, self.part, self.detail)

    def __str__(self):
        loc = f" [{self.where}]" if self.where else ""
        return f"{self.severity.upper():7} {self.code:24} {self.part}: {self.detail}{loc}"


_BLOCK_PARENTS = {
    w(x)
    for x in (
        "body", "tc", "txbxContent", "sdtContent", "comment", "footnote", "endnote",
        "hdr", "ftr", "customXml", "ins", "del", "moveFrom", "moveTo", "docPartBody",
    )
}
_RUN_PARENTS = {
    w(x)
    for x in (
        "p", "hyperlink", "ins", "del", "moveFrom", "moveTo", "smartTag", "sdtContent",
        "customXml", "fldSimple", "dir", "bdo", "rt", "rubyBase",
    )
}
_DELETED_CONTAINERS = {w("del"), w("moveFrom")}
_ORDERS = [("pPr", PPR_ORDER, "error"), ("tblPr", TBLPR_ORDER, "error"),
           ("tcPr", TCPR_ORDER, "error"), ("sectPr", SECTPR_ORDER, "error"),
           ("rPr", RPR_ORDER, "warning")]
_EXOTIC_IMAGES = {"emf", "wmf", "svg", "tif", "tiff"}


class _Validator:
    def __init__(self, path):
        self.path = str(path)
        self.findings: list[Finding] = []
        self.pkg = None

    def add(self, severity, code, part, detail, where=""):
        self.findings.append(Finding(severity, code, part, detail, where))

    def run(self):
        try:
            self.pkg = Package(self.path)
        except DocxError as exc:
            self.add("error", "PACKAGE_UNREADABLE", "-", str(exc))
            return self.findings
        self.check_zip()
        self.check_xml_parts()
        if any(f.code == "XML_MALFORMED" for f in self.findings):
            return self.findings
        self.check_content_types()
        self.check_relationships()
        self.check_ignorable()
        self.styles = Styles(self.pkg)
        self.num_ids = self.load_numbering()
        self.notes = {
            "footnoteReference": self.note_ids("word/footnotes.xml", "footnote"),
            "endnoteReference": self.note_ids("word/endnotes.xml", "endnote"),
            "commentReference": self.note_ids("word/comments.xml", "comment"),
        }
        self.bookmark_names = Counter()
        self.docpr_ids = Counter()
        self.para_ids = Counter()
        for part in self.pkg.story_parts():
            self.check_story(part)
        self.check_global_uniques()
        self.check_settings()
        return self.findings

    # ---- pacote -------------------------------------------------------------
    def check_zip(self):
        names = [i.filename for i in self.pkg.infos]
        if names[0] != "[Content_Types].xml":
            self.add("warning", "CT_NOT_FIRST", "[Content_Types].xml",
                     "não é a primeira entrada do ZIP")
        for name in names:
            if name.endswith("/"):
                self.add("warning", "ZIP_DIR_ENTRY", name, "entrada de diretório no pacote")
            if "\\" in name or name.startswith("/") or ".." in name.split("/"):
                self.add("error", "ZIP_BAD_NAME", name, "nome de entrada inválido")

    def check_xml_parts(self):
        for name in self.pkg.names():
            if name.endswith((".xml", ".rels")):
                try:
                    self.pkg.xml(name)
                except DocxError as exc:
                    self.add("error", "XML_MALFORMED", name, str(exc))

    def check_content_types(self):
        root = self.pkg.xml("[Content_Types].xml")
        defaults, overrides = set(), set()
        for el in root:
            if local(el) == "Default":
                defaults.add(el.get("Extension", "").lower())
            elif local(el) == "Override":
                overrides.add(el.get("PartName", ""))
                part = el.get("PartName", "").lstrip("/")
                if not self.pkg.has(part):
                    self.add("error", "CT_OVERRIDE_DANGLING", "[Content_Types].xml",
                             f"Override aponta para parte inexistente: {part}")
        for name in self.pkg.names():
            if name == "[Content_Types].xml" or name.endswith("/"):
                continue
            ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
            if "/" + name not in overrides and ext not in defaults:
                self.add("error", "CT_MISSING", name, "parte sem tipo de conteúdo declarado")

    def check_relationships(self):
        for name in self.pkg.names():
            if not name.endswith(".rels"):
                continue
            ids = Counter()
            base = name.replace("_rels/", "", 1)[:-5] if name != "_rels/.rels" else ""
            for rel in self.pkg.xml(name):
                ids[rel.get("Id")] += 1
                if rel.get("TargetMode") == "External":
                    continue
                target = resolve_target(base, rel.get("Target", ""))
                if not self.pkg.has(target):
                    self.add("error", "REL_DANGLING", name,
                             f"{rel.get('Id')} aponta para parte inexistente: {target}")
                ext = target.rsplit(".", 1)[-1].lower()
                if rel.get("Type") == RT_IMAGE and ext in _EXOTIC_IMAGES:
                    self.add("info", "IMAGE_FORMAT_COMPAT", name,
                             f"imagem .{ext} pode não renderizar no Google Docs/LibreOffice: {target}")
            for rid, count in ids.items():
                if count > 1:
                    self.add("error", "REL_DUPLICATE_ID", name, f"Id repetido: {rid}")

    def check_ignorable(self):
        for name in self.pkg.names():
            if not re.match(r"^word/[^/]+\.xml$", name):
                continue
            root = self.pkg.xml(name)
            for prefix in (root.get(MC_IGNORABLE) or "").split():
                if prefix not in root.nsmap:
                    self.add("error", "MC_IGNORABLE_UNDECLARED", name,
                             f"prefixo '{prefix}' em mc:Ignorable sem xmlns declarado")
            for choice in root.iter(MC_CHOICE):
                for prefix in (choice.get("Requires") or "").split():
                    if prefix not in choice.nsmap:
                        self.add("error", "MC_REQUIRES_UNDECLARED", name,
                                 f"prefixo '{prefix}' em mc:Choice/@Requires sem xmlns")

    # ---- referências --------------------------------------------------------
    def load_numbering(self):
        if not self.pkg.has("word/numbering.xml"):
            return set()
        root = self.pkg.xml("word/numbering.xml")
        abstracts = {a.get(w("abstractNumId")) for a in root.findall(w("abstractNum"))}
        ids = set()
        for num in root.findall(w("num")):
            ids.add(num.get(w("numId")))
            ref = wval(num.find(w("abstractNumId")))
            if ref not in abstracts:
                self.add("error", "NUM_ABSTRACT_MISSING", "word/numbering.xml",
                         f"numId {num.get(w('numId'))} referencia abstractNum {ref} inexistente")
        return ids

    def note_ids(self, part, tag):
        if not self.pkg.has(part):
            return set()
        return {el.get(w("id")) for el in self.pkg.xml(part).findall(w(tag))}

    # ---- conteúdo -----------------------------------------------------------
    def check_story(self, part):
        root = self.pkg.xml(part)
        rel_ids = {r.get("Id") for r in self.pkg.rels(part)}
        field_depth = 0
        bm_starts, bm_ends = Counter(), Counter()
        cm_starts, cm_ends = Counter(), Counter()
        for el in root.iter():
            if not isinstance(el.tag, str) or in_fallback(el):
                continue
            for attr, value in el.attrib.items():
                if attr.startswith("{%s}" % R_NS) and value and value not in rel_ids:
                    self.add("error", "RID_UNRESOLVED", part,
                             f"{local(el)} referencia relacionamento inexistente: {value}")
            if el.tag == "{%s}docPr" % WP_NS:
                self.docpr_ids[el.get("id")] += 1
            if W14_PARA_ID in el.attrib:
                self.check_para_id(part, el)
            if not is_w(el):
                continue
            name = local(el)
            parent = el.getparent()
            parent_w = parent is not None and is_w(parent)
            if name == "fldChar":
                kind = el.get(w("fldCharType"))
                if kind == "begin":
                    field_depth += 1
                elif kind == "end":
                    field_depth -= 1
                    if field_depth < 0:
                        self.add("error", "FIELD_UNBALANCED", part, "fldChar end sem begin")
                        field_depth = 0
                elif kind == "separate" and field_depth == 0:
                    self.add("error", "FIELD_UNBALANCED", part, "fldChar separate fora de campo")
            elif name == "p":
                if parent_w and parent.tag not in _BLOCK_PARENTS:
                    self.add("error", "BAD_NESTING", part, f"w:p dentro de w:{local(parent)}")
                self.check_first_child(part, el, "pPr")
            elif name == "r":
                if parent_w and parent.tag not in _RUN_PARENTS:
                    self.add("error", "BAD_NESTING", part, f"w:r dentro de w:{local(parent)}")
                self.check_first_child(part, el, "rPr")
            elif name in ("t", "delText", "instrText", "delInstrText"):
                self.check_text(part, el, name)
            elif name == "tbl":
                self.check_table(part, el)
            elif name == "tc":
                self.check_cell(part, el)
            elif name == "tr" and el.find(w("tc")) is None and el.find(w("sdt")) is None \
                    and el.find(w("customXml")) is None:
                self.add("error", "ROW_NO_CELLS", part, "w:tr sem células")
            elif name == "body":
                self.check_body(part, el)
            elif name in ("hdr", "ftr", "footnote", "endnote", "comment"):
                if el.find(w("p")) is None and el.find(w("tbl")) is None and el.find(w("sdt")) is None:
                    self.add("error", "STORY_EMPTY", part, f"w:{name} sem conteúdo de bloco")
            elif name == "pStyle":
                self.check_style(part, wval(el), "paragraph")
            elif name == "rStyle":
                self.check_style(part, wval(el), "character")
            elif name == "tblStyle":
                self.check_style(part, wval(el), "table")
            elif name == "numId" and parent_w and local(parent) == "numPr":
                if wval(el) != "0" and wval(el) not in self.num_ids:
                    self.add("error", "NUM_ID_MISSING", part, f"numId {wval(el)} não existe em numbering.xml")
            elif name in self.notes:
                if el.get(w("id")) not in self.notes[name]:
                    self.add("error", "NOTE_REF_MISSING", part, f"{name} id={el.get(w('id'))} sem entrada correspondente")
            elif name == "bookmarkStart":
                bm_starts[el.get(w("id"))] += 1
                self.bookmark_names[el.get(w("name"))] += 1
            elif name == "bookmarkEnd":
                bm_ends[el.get(w("id"))] += 1
            elif name == "commentRangeStart":
                cm_starts[el.get(w("id"))] += 1
            elif name == "commentRangeEnd":
                cm_ends[el.get(w("id"))] += 1
            if name in ("ins", "del", "moveFrom", "moveTo") and el.get(w("author")) is None:
                self.add("error", "REVISION_NO_AUTHOR", part, f"w:{name} sem w:author")
            if name in ("ins", "del", "moveFrom", "moveTo") and el.get(w("id")) is None:
                self.add("error", "REVISION_NO_ID", part, f"w:{name} sem w:id")
            for container, order, severity in _ORDERS:
                if name == container:
                    self.check_order(part, el, order, severity)
        if field_depth > 0:
            self.add("error", "FIELD_UNBALANCED", part, f"{field_depth} campo(s) sem fldChar end")
        self.pair_check(part, "BOOKMARK_UNPAIRED", "bookmark", bm_starts, bm_ends)
        self.pair_check(part, "COMMENT_RANGE_UNPAIRED", "commentRange", cm_starts, cm_ends)
        for bid, count in bm_starts.items():
            if count > 1:
                self.add("warning", "BOOKMARK_DUP_ID", part, f"id de indicador repetido: {bid}")

    def pair_check(self, part, code, label, starts, ends):
        for key in sorted(set(starts) | set(ends), key=str):
            if starts[key] != ends[key]:
                self.add("warning", code, part,
                         f"{label} id={key}: {starts[key]} início(s), {ends[key]} fim(ns)")

    def check_first_child(self, part, el, tag):
        props = el.findall(w(tag))
        if len(props) > 1:
            self.add("error", "DUPLICATE_PROPS", part, f"w:{local(el)} com mais de um w:{tag}")
        elif props and any(isinstance(c.tag, str) for c in el[:el.index(props[0])]):
            self.add("error", "PROPS_NOT_FIRST", part, f"w:{tag} não é o primeiro filho de w:{local(el)}")

    def check_order(self, part, el, order, severity):
        last = -1
        for child in el:
            if not isinstance(child.tag, str) or not is_w(child) or local(child) not in order:
                continue
            idx = order.index(local(child))
            if idx < last:
                self.add(severity, "CHILD_ORDER", part,
                         f"w:{local(child)} fora de ordem em w:{local(el)}")
                return
            last = idx

    def check_text(self, part, el, name):
        parent = el.getparent()
        if parent.tag != w("r"):
            self.add("error", "BAD_NESTING", part, f"w:{name} dentro de w:{local(parent)}")
            return
        deleted = any(a.tag in _DELETED_CONTAINERS for a in parent.iterancestors())
        if name in ("t", "instrText") and deleted:
            self.add("error", "TEXT_IN_DELETION", part, f"w:{name} dentro de revisão de exclusão (use w:delText)")
        if name in ("delText", "delInstrText") and not deleted:
            self.add("error", "DELTEXT_OUTSIDE_DELETION", part, f"w:{name} fora de w:del")
        text = el.text or ""
        if name in ("t", "delText") and text != text.strip() and el.get(XML_SPACE) != "preserve":
            self.add("warning", "SPACE_NOT_PRESERVED", part,
                     "w:t com espaço nas pontas sem xml:space='preserve'", repr(text[:40]))

    def check_table(self, part, tbl):
        if tbl.find(w("tblPr")) is None:
            self.add("error", "TABLE_NO_TBLPR", part, "w:tbl sem w:tblPr")
        grid = tbl.find(w("tblGrid"))
        if grid is None:
            self.add("error", "TABLE_NO_GRID", part, "w:tbl sem w:tblGrid")
        rows = tbl.findall(w("tr"))
        if not rows and tbl.find(w("sdt")) is None and tbl.find(w("customXml")) is None:
            self.add("error", "TABLE_NO_ROWS", part, "w:tbl sem linhas")
        if grid is not None:
            cols = len(grid.findall(w("gridCol")))
            for row in rows:
                if row.find(w("sdt")) is not None or row.find(w("customXml")) is not None:
                    continue
                total = row_grid_width(row)
                if total != cols:
                    self.add("warning", "GRID_MISMATCH", part, f"linha ocupa {total} colunas; tblGrid tem {cols}")
        next_el = tbl.getnext()
        if next_el is not None and next_el.tag == w("tbl"):
            self.add("warning", "ADJACENT_TABLES", part, "tabelas sem parágrafo entre elas serão mescladas pelo Word")
        for row, above in zip(rows[1:], rows[:-1]):
            above_map = row_cells_by_grid(above)
            for start, cell in row_cells_by_grid(row).items():
                if cell_vmerge(cell) == "continue":
                    prev = above_map.get(start)
                    if prev is None or cell_vmerge(prev) is None:
                        self.add("warning", "VMERGE_ORPHAN", part, "vMerge de continuação sem célula inicial acima")

    def check_cell(self, part, tc):
        blocks = [c for c in tc if isinstance(c.tag, str) and c.tag != w("tcPr") and is_w(c)
                  and c.tag not in MARKER_TAGS]
        if not blocks:
            self.add("error", "CELL_NO_PARAGRAPH", part, "w:tc sem parágrafo")
        elif blocks[-1].tag == w("tbl"):
            self.add("error", "CELL_ENDS_WITH_TABLE", part, "w:tc precisa terminar em w:p")

    def check_body(self, part, body):
        kids = [c for c in body if isinstance(c.tag, str)]
        sects = [c for c in kids if c.tag == w("sectPr")]
        if len(sects) > 1 or (sects and kids[-1] is not sects[0]):
            self.add("error", "BODY_SECTPR", part, "w:sectPr deve ser único e o último filho de w:body")
        if not sects:
            self.add("warning", "BODY_NO_SECTPR", part, "w:body sem w:sectPr final")
        content = [c for c in kids if c.tag != w("sectPr") and c.tag not in MARKER_TAGS]
        if content and content[-1].tag == w("tbl"):
            self.add("warning", "BODY_ENDS_WITH_TABLE", part, "tabela no fim do corpo sem parágrafo depois dela")

    def check_style(self, part, sid, kind):
        if sid not in self.styles.by_id:
            self.add("error", "STYLE_MISSING", part, f"estilo '{sid}' não existe em styles.xml")
        elif self.styles.type_of(sid) != kind:
            self.add("warning", "STYLE_WRONG_TYPE", part, f"estilo '{sid}' não é do tipo {kind}")

    def check_para_id(self, part, el):
        value = el.get(W14_PARA_ID)
        self.para_ids[value] += 1
        try:
            if int(value, 16) >= 0x80000000:
                self.add("error", "PARA_ID_RANGE", part, f"w14:paraId fora do intervalo: {value}")
        except ValueError:
            self.add("error", "PARA_ID_RANGE", part, f"w14:paraId não hexadecimal: {value}")

    def check_global_uniques(self):
        for name, count in self.bookmark_names.items():
            if count > 1:
                self.add("warning", "BOOKMARK_DUP_NAME", "*", f"nome de indicador repetido: {name}")
        for value, count in self.docpr_ids.items():
            if count > 1:
                self.add("warning", "DOCPR_DUP_ID", "*", f"wp:docPr id repetido: {value}")
        for value, count in self.para_ids.items():
            if count > 1:
                self.add("warning", "PARA_ID_DUP", "*", f"w14:paraId repetido: {value}")

    def check_settings(self):
        settings = self.pkg.settings()
        if settings is not None and is_on(settings.find(w("updateFields"))):
            self.add("warning", "UPDATE_FIELDS_PROMPT", "word/settings.xml",
                     "w:updateFields faz o Word exibir um aviso ao abrir o documento")


MARKER_TAGS = {
    w(x)
    for x in (
        "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd",
        "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
        "permStart", "permEnd",
    )
}


def row_grid_width(row) -> int:
    trpr = row.find(w("trPr"))
    total = 0
    if trpr is not None:
        total += int(wval(trpr.find(w("gridBefore")), "0") or 0)
        total += int(wval(trpr.find(w("gridAfter")), "0") or 0)
    for tc in row.findall(w("tc")):
        total += cell_span(tc)
    return total


def cell_span(tc) -> int:
    tcpr = tc.find(w("tcPr"))
    if tcpr is None:
        return 1
    return int(wval(tcpr.find(w("gridSpan")), "1") or 1)


def cell_vmerge(tc):
    """None, 'restart' ou 'continue'."""
    tcpr = tc.find(w("tcPr"))
    if tcpr is None:
        return None
    vm = tcpr.find(w("vMerge"))
    if vm is None:
        return None
    return "restart" if wval(vm) == "restart" else "continue"


def row_cells_by_grid(row) -> dict:
    """{coluna inicial na grade: célula}."""
    trpr = row.find(w("trPr"))
    pos = int(wval(trpr.find(w("gridBefore")), "0") or 0) if trpr is not None else 0
    out = {}
    for tc in row.findall(w("tc")):
        out[pos] = tc
        pos += cell_span(tc)
    return out


def validate(path) -> list[Finding]:
    return _Validator(path).run()


def regressions(findings, baseline) -> list[Finding]:
    """Achados que existem em `findings` mas não existiam (ou eram menos) no baseline."""
    remaining = Counter(f.key for f in baseline)
    out = []
    for finding in findings:
        if remaining[finding.key] > 0:
            remaining[finding.key] -= 1
        else:
            out.append(finding)
    return out
