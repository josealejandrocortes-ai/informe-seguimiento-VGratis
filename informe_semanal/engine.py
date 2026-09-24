"""
Motor para diligenciar el Informe Seguimiento Semanal (Word) a partir del Excel de soporte.

Qué hace:
  1. Lee el Excel (.xlsm/.xlsx) con los valores guardados (las tablas dinámicas deben estar
     actualizadas y el archivo guardado antes de subirlo).
  2. Reemplaza en el Word los marcadores {...} conservando el formato de cada texto.
  3. Reemplaza los marcadores <instrucción> con un análisis redactado (Claude) apoyado en el Excel.
  4. Actualiza los gráficos del Word (gráficos nativos vinculados a Excel) con los datos nuevos,
     sin tocar su formato, tamaño ni posición.

Sintaxis de marcadores (ver README.md):
  {'Hoja'!B2}                 valor de la celda, tal como lo muestra Excel (formato de número de la celda)
  {'Hoja'!B2|M}               modificadores: M, MM, mill, pct, pct0, pct2, pp, abs, dir, up, cap, txt, num0/1/2
  {@fecha_corte|larga}        fecha de corte (larga, dm, ab, mes, iso) / {@fecha_anterior|...}
  {@mov.obligaciones.dep_art} mayor variación semanal por dependencia y rubro (calculada)
  {@buscar('Hoja'!Q18; 'Otra'!B:C)}   búsqueda tipo BUSCARV
  <instrucción para el análisis>      texto generado con IA
"""
from __future__ import annotations

import copy
import datetime as dt
import io
import re
import zipfile
from decimal import ROUND_HALF_UP, Decimal
from typing import Callable, Optional

import openpyxl
from lxml import etree
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import to_excel

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
W = "{%s}" % NS["w"]
C = "{%s}" % NS["c"]
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]
MESES_AB = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


# ----------------------------------------------------------------------------------------
# Formato numérico estilo Excel (es-CO: punto para miles, coma para decimales)
# ----------------------------------------------------------------------------------------
def _round(v: float, dec: int) -> Decimal:
    q = Decimal(1).scaleb(-dec)
    return Decimal(repr(float(v))).quantize(q, rounding=ROUND_HALF_UP)


def es_num(v: float, dec: int = 0, miles: bool = True, min_int: int = 1) -> str:
    d = _round(abs(v), dec)
    ent, _, frac = f"{d:.{dec}f}".partition(".")
    ent = ent.zfill(min_int)
    if miles and len(ent) > 3:
        ent = re.sub(r"(?<=\d)(?=(\d{3})+$)", ".", ent)
    out = ent + ("," + frac if dec else "")
    return ("-" if v < 0 and d != 0 else "") + out


def _clean_literals(s: str) -> str:
    s = re.sub(r"\[[^\]]*\]", "", s)          # [Red], [$-409]
    s = re.sub(r"_.", "", s)                   # _-  (espacio de relleno)
    s = re.sub(r"\*.", "", s)                  # *   (relleno)
    s = s.replace("\\", "").replace('"', "")
    return s


def excel_display(v, fmt: Optional[str]) -> str:
    """Devuelve el valor tal como lo mostraría Excel con el formato de número de la celda."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "VERDADERO" if v else "FALSO"
    if isinstance(v, (dt.datetime, dt.date)):
        f = (fmt or "").lower()
        if "mmmm" in f:
            return f"{MESES[v.month - 1]} {v.year}" if "yy" in f else MESES[v.month - 1]
        if "mmm" in f:
            return f"{MESES_AB[v.month - 1].lower()}-{str(v.year)[2:]}"
        return v.strftime("%d/%m/%Y")
    if isinstance(v, str):
        return v
    fmt = fmt or "General"
    if fmt == "General" or re.fullmatch(r"[a-zA-Z ]+", fmt):
        if float(v).is_integer():
            return es_num(v, 0, miles=False)
        return es_num(v, 10, miles=False).rstrip("0").rstrip(",")
    sec = fmt.split(";")[0]
    clean = _clean_literals(sec)
    m = re.search(r"[#0?,\.]*[#0?][#0?,\.]*", clean)
    if not m:
        return es_num(v, 0)
    tok = m.group(0)
    prefix, suffix = clean[: m.start()].strip(), clean[m.end():].strip()
    trail = len(tok) - len(tok.rstrip(","))
    core = tok.rstrip(",")
    intpart, _, fracpart = core.partition(".")
    dec = len([ch for ch in fracpart if ch in "0#?"])
    miles = "," in intpart
    min_int = intpart.count("0")
    val = float(v)
    if "%" in clean:
        val *= 100
    val /= 1000 ** trail
    txt = es_num(val, dec, miles=miles, min_int=max(min_int, 1))
    if txt.startswith("-"):
        return "-" + prefix + txt[1:] + suffix
    return prefix + txt + suffix


def _fmt_mod(v, mod: str, disp: str) -> tuple:
    """Aplica un modificador. Devuelve (nuevo_valor_numérico_o_texto, es_texto_final)."""
    m = mod.strip().lower()
    num = isinstance(v, (int, float)) and not isinstance(v, bool)
    if m == "txt":
        return str(v if v is not None else ""), True
    if m == "up":
        return disp.upper(), True
    if m == "cap":
        return disp[:1].upper() + disp[1:].lower(), True
    if not num:
        return v, False
    if m == "abs":
        return abs(v), False
    if m == "m":                     # pesos -> $1.234M
        return f"${es_num(v / 1e6, 0)}M", True
    if m == "mm":                    # ya en millones -> $1.234M
        return f"${es_num(v, 0)}M", True
    if m == "mill":                  # pesos -> 1.234
        return es_num(v / 1e6, 0), True
    if m in ("pct", "pct1"):
        return es_num(v * 100, 1) + "%", True
    if m == "pct0":
        return es_num(v * 100, 0) + "%", True
    if m == "pct2":
        return es_num(v * 100, 2) + "%", True
    if m == "pp":                    # fracción -> puntos porcentuales (0,2)
        return es_num(v * 100, 1), True
    if m == "dir":                   # aumento / disminución
        return ("aumento" if v > 0 else "disminución" if v < 0 else "variación nula"), True
    if m in ("num0", "num1", "num2"):
        return es_num(v, int(m[-1])), True
    return v, False


# ----------------------------------------------------------------------------------------
# Libro de Excel
# ----------------------------------------------------------------------------------------
_REF_RE = re.compile(
    r"^\s*(?:'(?P<q>[^']+)'|(?P<u>[^'!\s][^'!]*?))!\s*(?P<a>\$?[A-Za-z]{1,3}\$?\d*)(?::(?P<b>\$?[A-Za-z]{1,3}\$?\d*))?\s*$"
)


def _norm_sheet(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip().casefold()


def _split_cell(a: str, default_row: int) -> tuple:
    m = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d*)", a.strip())
    col = column_index_from_string(m.group(1).upper())
    row = int(m.group(2)) if m.group(2) else default_row
    return col, row


class Libro:
    """Envoltura de openpyxl con búsqueda tolerante de hojas y valores formateados."""

    def __init__(self, data: bytes | str):
        src = io.BytesIO(data) if isinstance(data, (bytes, bytearray)) else data
        self.wb = openpyxl.load_workbook(src, data_only=True)
        self._map = {_norm_sheet(n): n for n in self.wb.sheetnames}

    @property
    def sheetnames(self):
        return self.wb.sheetnames

    def hoja(self, nombre: str):
        k = _norm_sheet(nombre)
        if k not in self._map:
            raise KeyError(f"No existe la hoja '{nombre}' en el Excel")
        return self.wb[self._map[k]]

    def parse_ref(self, ref: str):
        m = _REF_RE.match(ref)
        if not m:
            raise ValueError(f"Referencia no válida: {ref}")
        ws = self.hoja(m.group("q") or m.group("u"))
        c1, r1 = _split_cell(m.group("a"), 1)
        if m.group("b"):
            c2, r2 = _split_cell(m.group("b"), ws.max_row)
        else:
            c2, r2 = c1, r1
        return ws, c1, r1, c2, r2

    def celdas(self, ref: str):
        ws, c1, r1, c2, r2 = self.parse_ref(ref)
        return [[ws.cell(r, c) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]

    def valor(self, ref: str):
        return self.celdas(ref)[0][0].value

    def mostrar(self, ref: str, mods: str = "") -> str:
        cell = self.celdas(ref)[0][0]
        v = cell.value
        disp = excel_display(v, cell.number_format)
        if not mods:
            return disp
        cur = v
        for mod in mods.split(","):
            cur, final = _fmt_mod(cur, mod, disp)
            if final:
                return str(cur)
        return excel_display(cur, cell.number_format) if not isinstance(cur, str) else cur


# ----------------------------------------------------------------------------------------
# Datos calculados (@...)
# ----------------------------------------------------------------------------------------
def adivinar_fecha_corte(lib: Libro) -> dt.date:
    """Toma 'Sem 18 Sep' de la hoja de ejecución y el año de la fecha de extracción."""
    year, extr = dt.date.today().year, None
    try:
        v = lib.valor("'Ejec Presupuestal Vigente'!A1")
        if isinstance(v, (dt.datetime, dt.date)):
            extr = v.date() if isinstance(v, dt.datetime) else v
            year = extr.year
    except Exception:
        pass
    try:
        txt = str(lib.valor("'Ejecución presupuestal'!A2"))
        m = re.search(r"(\d{1,2})\s+([A-Za-zñÑ]{3})", txt)
        if m:
            mes = [x.lower() for x in MESES_AB].index(m.group(2).lower()[:3]) + 1
            return dt.date(year, mes, int(m.group(1)))
    except Exception:
        pass
    return (extr - dt.timedelta(days=1)) if extr else dt.date.today()


def fmt_fecha(d: dt.date, f: str) -> str:
    f = (f or "larga").lower()
    if f == "dm":
        return f"{d.day} de {MESES[d.month - 1]}"
    if f == "ab":
        return f"{d.day} {MESES_AB[d.month - 1]}"
    if f == "mes":
        return MESES[d.month - 1]
    if f == "iso":
        return d.isoformat()
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def _filas_ejecucion(lib: Libro, hoja: str):
    ws = lib.hoja(hoja)
    hdr = {}
    for r in range(1, 6):
        vals = [str(c.value or "").strip().upper().replace("\n", " ") for c in ws[r]]
        if "COMPROMISO" in vals and "OBLIGACION" in vals:
            hdr = {v: i for i, v in enumerate(vals) if v}
            first = r + 1
            break
    else:
        raise ValueError(f"No se encontró el encabezado de ejecución en '{hoja}'")
    ix = lambda k: hdr[k]
    out = []
    for row in ws.iter_rows(min_row=first, values_only=True):
        comp = row[ix("COMPROMISO")]
        if row[0] is None or not isinstance(comp, (int, float)):
            continue
        out.append({
            "despacho": str(row[ix("DESPACHO")]).strip(),
            "rubro": str(row[ix("RUBRO NORMALIZADO")]).strip(),
            "desc": str(row[ix("DESCRIPCION")] or "").strip(),
            "compromisos": comp or 0,
            "obligaciones": row[ix("OBLIGACION")] or 0,
            "pagos": row[ix("PAGOS")] or 0,
        })
    return out


def _etiqueta_dep(lib: Libro, despacho: str) -> str:
    try:
        for fila in lib.celdas("'Ejecución presupuestal'!B8:B11"):
            v = fila[0].value
            if v and _norm_sheet(str(v)) == _norm_sheet(despacho):
                return str(v)
    except Exception:
        pass
    return despacho.capitalize()


def calcular_movimientos(lib: Libro) -> dict:
    """Mayor variación semanal por momento (compromisos/obligaciones/pagos): dependencia y rubro."""
    actual = _filas_ejecucion(lib, "Ejec Presupuestal Vigente")
    previa = _filas_ejecucion(lib, "Ejec Presupuestal Sem Ant")

    def agg(rows):
        d: dict = {}
        for r in rows:
            k = (r["despacho"], r["rubro"])
            e = d.setdefault(k, {"compromisos": 0, "obligaciones": 0, "pagos": 0, "desc": r["desc"]})
            for m in ("compromisos", "obligaciones", "pagos"):
                e[m] += r[m]
        return d

    a, b = agg(actual), agg(previa)
    res = {}
    for mom in ("compromisos", "obligaciones", "pagos"):
        por_dep: dict = {}
        base_dep: dict = {}
        for k in set(a) | set(b):
            dep, rub = k
            va, vb = a.get(k, {}).get(mom, 0), b.get(k, {}).get(mom, 0)
            por_dep.setdefault(dep, []).append((va - vb, rub, (a.get(k) or b.get(k))["desc"]))
            base_dep[dep] = base_dep.get(dep, 0) + vb
        # Igual que el Excel: la dependencia con mayor variación PORCENTUAL en el momento
        dep = max(por_dep, key=lambda d: (sum(x[0] for x in por_dep[d]) / base_dep[d]) if base_dep[d] else 0)
        total = sum(x[0] for x in por_dep[dep])
        dr, rub, desc = max(por_dep[dep], key=lambda x: x[0])
        nombre = _etiqueta_dep(lib, dep)
        art = "La" if nombre.lower().startswith("secretar") else "El"
        res[mom] = {
            "dep": nombre, "dep_art": f"{art} {nombre}", "valor": f"${es_num(total / 1e6, 0)}M",
            "pct": es_num(total / base_dep[dep] * 100, 1) + "%" if base_dep[dep] else "n.a.",
            "rubro": rub, "rubro_valor": f"${es_num(dr / 1e6, 0)}M", "rubro_desc": desc,
            "tipo": "funcionamiento" if rub.upper().startswith("A") else "inversión",
        }
    return res


# ----------------------------------------------------------------------------------------
# Marcadores
# ----------------------------------------------------------------------------------------
TOKEN_RE = re.compile(r"<[^<>]{12,}>|\{[^{}]+\}")


class Contexto:
    def __init__(self, lib: Libro, fecha_corte: dt.date, glosario: Optional[dict] = None):
        self.lib, self.fecha, self.glosario = lib, fecha_corte, glosario or {}
        self._mov = None

    def mov(self):
        if self._mov is None:
            self._mov = calcular_movimientos(self.lib)
        return self._mov

    def glosa(self, s: str) -> str:
        for a, b in self.glosario.items():
            s = s.replace(a, b)
        return s

    def resolver(self, cuerpo: str) -> str:
        """Resuelve el contenido de {…}. Lanza excepción si no se puede."""
        cuerpo = cuerpo.strip()
        if cuerpo.startswith("@"):
            return self._resolver_especial(cuerpo[1:])
        ref, _, mods = cuerpo.partition("|")
        return self.glosa(self.lib.mostrar(ref.strip(), mods))

    def _resolver_especial(self, k: str) -> str:
        k, _, mods = k.partition("|")
        k = k.strip()
        if k in ("fecha_corte", "fecha_anterior"):
            d = self.fecha - (dt.timedelta(days=7) if k == "fecha_anterior" else dt.timedelta(0))
            return fmt_fecha(d, mods)
        if k == "anio":
            return str(self.fecha.year)
        if k.startswith("mov."):
            _, mom, campo = k.split(".")
            return self.glosa(str(self.mov()[mom][campo]))
        m = re.fullmatch(r"buscar\((.+?);\s*(.+?)(?:;\s*(\d+))?\)", k)
        if m:
            clave = self.resolver(m.group(1))
            ws, c1, r1, c2, r2 = self.lib.parse_ref(m.group(2))
            col = int(m.group(3) or 2) - 1
            for r in range(r1, r2 + 1):
                if str(ws.cell(r, c1).value).strip() == clave.strip():
                    c = ws.cell(r, c1 + col)
                    return self.glosa(excel_display(c.value, c.number_format))
            raise KeyError(f"'{clave}' no encontrado")
        raise KeyError(f"Marcador especial desconocido: @{k}")

    def resolver_instruccion(self, texto: str) -> str:
        """Sustituye {…} dentro de una instrucción <…> por sus valores."""
        def rep(m):
            try:
                return self.resolver(m.group(1))
            except Exception:
                return m.group(0)
        return re.sub(r"\{([^{}]+)\}", rep, texto)


# ----------------------------------------------------------------------------------------
# Word: utilidades de párrafos y runs (conservan el formato)
# ----------------------------------------------------------------------------------------
def _runs_propios(p):
    return [r for r in p.iter(W + "r") if _padre_p(r) is p]


def _padre_p(el):
    a = el.getparent()
    while a is not None and a.tag != W + "p":
        a = a.getparent()
    return a


def _texto_run(r) -> str:
    return "".join(t.text or "" for t in r.findall(W + "t"))


def _set_texto_run(r, txt: str):
    ts = r.findall(W + "t")
    if not ts:
        t = etree.SubElement(r, W + "t")
        ts = [t]
    ts[0].text = txt
    ts[0].set(XML_SPACE, "preserve")
    for extra in ts[1:]:
        r.remove(extra)


def texto_parrafo(p) -> str:
    return "".join(_texto_run(r) for r in _runs_propios(p))


def reemplazar_en_parrafo(p, inicio: int, fin: int, nuevo: str):
    """Reemplaza el rango [inicio, fin) del texto del párrafo conservando el formato del primer run."""
    runs = _runs_propios(p)
    textos = [_texto_run(r) for r in runs]
    acum, primero = 0, None
    for i, t in enumerate(textos):
        a, b = acum, acum + len(t)
        acum = b
        if b <= inicio:
            continue
        if a >= fin:
            break
        pre = t[: max(0, inicio - a)]
        post = t[fin - a:] if fin < b else ""
        if primero is None:
            primero = i
            _set_texto_run(runs[i], pre + nuevo + post)
        else:
            _set_texto_run(runs[i], post)


def _es_solo_marcador(p, tok_ini, tok_fin) -> bool:
    return texto_parrafo(p).strip() == texto_parrafo(p)[tok_ini:tok_fin].strip()


def _partes_xml(zin: zipfile.ZipFile):
    return [n for n in zin.namelist()
            if re.fullmatch(r"word/(document|header\d*|footer\d*)\.xml", n)]


# ----------------------------------------------------------------------------------------
# Escaneo y construcción
# ----------------------------------------------------------------------------------------
def _norm_ws(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def escanear(plantilla: bytes, ctx: Contexto) -> dict:
    """Lista marcadores de valor (con su resultado) y de análisis (con la instrucción resuelta)."""
    valores, analisis, vistos = [], [], set()
    with zipfile.ZipFile(io.BytesIO(plantilla)) as z:
        for parte in _partes_xml(z):
            root = etree.fromstring(z.read(parte))
            for p in root.iter(W + "p"):
                txt = texto_parrafo(p)
                for m in TOKEN_RE.finditer(txt):
                    tok = m.group(0)
                    if tok.startswith("<"):
                        instr = ctx.resolver_instruccion(_norm_ws(tok[1:-1]))
                        if instr not in vistos:
                            vistos.add(instr)
                            analisis.append({"instruccion": instr, "parte": parte})
                    else:
                        try:
                            v, err = ctx.resolver(tok[1:-1]), None
                        except Exception as e:  # noqa: BLE001
                            v, err = None, f"{type(e).__name__}: {e}"
                        valores.append({"marcador": tok, "valor": v, "error": err, "parte": parte})
    return {"valores": valores, "analisis": analisis}


def construir(plantilla: bytes, ctx: Contexto, analisis: Optional[dict] = None,
              refrescar_graficos: bool = True) -> tuple[bytes, dict]:
    """Devuelve (docx_bytes, informe) con conteos y advertencias."""
    analisis = analisis or {}
    info = {"reemplazados": 0, "sin_resolver": [], "analisis_pendientes": [], "graficos": []}
    zin = zipfile.ZipFile(io.BytesIO(plantilla))
    salida = io.BytesIO()
    with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in _partes_xml(zin):
                data = _rellenar_parte(data, ctx, analisis, info)
            elif refrescar_graficos and re.fullmatch(r"word/charts/chart\d+\.xml", item.filename):
                data = _refrescar_grafico(data, ctx.lib, item.filename, info)
            zout.writestr(item, data)
    return salida.getvalue(), info


def _rellenar_parte(xml: bytes, ctx: Contexto, analisis: dict, info: dict) -> bytes:
    root = etree.fromstring(xml)
    for p in list(root.iter(W + "p")):
        txt = texto_parrafo(p)
        toks = list(TOKEN_RE.finditer(txt))
        if not toks:
            continue
        extra_parrafos = []
        for m in reversed(toks):
            tok = m.group(0)
            if tok.startswith("<"):
                instr = ctx.resolver_instruccion(_norm_ws(tok[1:-1]))
                res = analisis.get(instr)
                if not res:
                    info["analisis_pendientes"].append(instr)
                    continue
                trozos = [t.strip() for t in re.split(r"\n\s*\n", res.strip()) if t.strip()]
                if len(trozos) > 1 and _es_solo_marcador(p, m.start(), m.end()):
                    extra_parrafos = trozos[1:]
                    res = trozos[0]
                else:
                    res = " ".join(trozos)
                reemplazar_en_parrafo(p, m.start(), m.end(), res)
                info["reemplazados"] += 1
            else:
                try:
                    val = ctx.resolver(tok[1:-1])
                except Exception as e:  # noqa: BLE001
                    info["sin_resolver"].append(f"{tok} → {type(e).__name__}: {e}")
                    continue
                reemplazar_en_parrafo(p, m.start(), m.end(), val)
                info["reemplazados"] += 1
        for t in reversed(extra_parrafos):
            nuevo = copy.deepcopy(p)
            runs = _runs_propios(nuevo)
            if runs:
                _set_texto_run(runs[0], t)
                for r in runs[1:]:
                    _set_texto_run(r, "")
            p.addnext(nuevo)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ----------------------------------------------------------------------------------------
# Gráficos: se reescriben las cachés (c:numCache / c:strCache) desde el Excel
# ----------------------------------------------------------------------------------------
def _valores_ref(lib: Libro, f: str):
    ws, c1, r1, c2, r2 = lib.parse_ref(f)
    out = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            out.append((cell.value, cell.number_format))
    return out


def _refrescar_grafico(xml: bytes, lib: Libro, nombre: str, info: dict) -> bytes:
    root = etree.fromstring(xml)
    n_ref, errores = 0, []
    for ref in list(root.iter(C + "numRef")) + list(root.iter(C + "strRef")):
        f_el = ref.find(C + "f")
        if f_el is None or not f_el.text:
            continue
        try:
            vals = _valores_ref(lib, f_el.text)
        except Exception as e:  # noqa: BLE001
            errores.append(f"{f_el.text}: {e}")
            continue
        es_num_ref = ref.tag == C + "numRef"
        cache = ref.find(C + ("numCache" if es_num_ref else "strCache"))
        if cache is None:
            cache = etree.SubElement(ref, C + ("numCache" if es_num_ref else "strCache"))
        fc = cache.find(C + "formatCode")
        for ch in list(cache):
            if ch is not fc:
                cache.remove(ch)
        if es_num_ref and fc is None:
            fc = etree.Element(C + "formatCode")
            fc.text = vals[0][1] if vals else "General"
            cache.insert(0, fc)
        pc = etree.SubElement(cache, C + "ptCount")
        pc.set("val", str(len(vals)))
        for i, (v, fmt) in enumerate(vals):
            if v is None:
                continue
            if es_num_ref and isinstance(v, str):
                v = v if _es_numero(v) else 0      # Excel grafica el texto/errores como 0
            if es_num_ref and isinstance(v, (dt.datetime, dt.date)):
                v = to_excel(v)
            txt = str(v) if es_num_ref else excel_display(v, fmt)
            if es_num_ref and isinstance(v, float):
                txt = repr(v)
            pt = etree.SubElement(cache, C + "pt")
            pt.set("idx", str(i))
            ve = etree.SubElement(pt, C + "v")
            ve.text = txt
        n_ref += 1
    info["graficos"].append({"parte": nombre, "referencias": n_ref, "errores": errores})
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _es_numero(s: str) -> bool:
    try:
        float(s)
        return True
    except ValueError:
        return False


# ----------------------------------------------------------------------------------------
# Análisis con IA (Claude)
# ----------------------------------------------------------------------------------------
SISTEMA = (
    "Eres analista financiero y presupuestal de una entidad pública colombiana. Redactas apartados de un "
    "informe semanal de seguimiento a la cadena presupuestal (compromisos, obligaciones, pagos, rezago, "
    "legalizaciones). Reglas: (1) usa EXCLUSIVAMENTE las cifras del contexto; no inventes ni recalcules "
    "cifras que ya vengan dadas; (2) cifras en millones de pesos con el formato $1.234M y porcentajes con "
    "coma decimal (0,6%); (3) redacción formal, en español, en tercera persona, sin viñetas ni títulos salvo "
    "que la instrucción lo pida; (4) responde SOLO con el texto que reemplaza al marcador, sin preámbulos "
    "ni comillas; (5) si el contexto no alcanza para afirmar algo, dilo en una frase corta en lugar de suponer."
)


def _celda_ia(c) -> str:
    """Valor mostrado + valor crudo (los formatos en millones ocultan la cifra real)."""
    disp = excel_display(c.value, c.number_format)
    if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
        return f"{disp} [{round(c.value, 4)}]"
    return disp


def contexto_para_ia(ctx: Contexto, instruccion: str, max_celdas: int = 400) -> str:
    """Arma el contexto: rangos citados en la instrucción + movimientos calculados."""
    partes = []
    refs = re.findall(r"(?:'[^']+'|[A-Za-zÁÉÍÓÚáéíóúñÑ_]+)![$A-Za-z0-9:]+", instruccion)
    for ref in dict.fromkeys(refs):
        try:
            celdas = ctx.lib.celdas(ref)
        except Exception:
            continue
        filas, n = [], 0
        for fila in celdas:
            if n > max_celdas:
                filas.append("… (recortado)")
                break
            filas.append(" | ".join(_celda_ia(c) for c in fila))
            n += len(fila)
        partes.append(f"### {ref}\n" + "\n".join(filas))
    try:
        mov = ctx.mov()
        lineas = [f"- {k}: " + "; ".join(f"{a}={b}" for a, b in v.items()) for k, v in mov.items()]
        partes.append("### Mayor variación semanal por momento (calculada de 'Ejec Presupuestal Vigente' vs "
                      "'Ejec Presupuestal Sem Ant')\n" + "\n".join(lineas))
    except Exception:
        pass
    partes.append(f"### Fechas\nCorte: {fmt_fecha(ctx.fecha, 'larga')}; semana anterior: "
                  f"{fmt_fecha(ctx.fecha - dt.timedelta(days=7), 'larga')}")
    return "\n\n".join(partes)


PROVEEDORES = {
    "Google Gemini (plan gratuito)": {
        "tipo": "openai", "url": "https://generativelanguage.googleapis.com/v1beta/openai",
        "modelo": "gemini-2.5-flash", "env": "GEMINI_API_KEY", "gratis": True, "requiere_key": True,
        "nota": "Clave gratis en aistudio.google.com. En el plan gratuito Google puede usar los datos enviados "
                "para mejorar sus productos."},
    "Groq (plan gratuito)": {
        "tipo": "openai", "url": "https://api.groq.com/openai/v1",
        "modelo": "openai/gpt-oss-120b", "env": "GROQ_API_KEY", "gratis": True, "requiere_key": True,
        "nota": "Clave gratis en console.groq.com/keys. Modelos abiertos (gpt-oss, Llama); límite diario de tokens."},
    "OpenRouter (modelos :free)": {
        "tipo": "openai", "url": "https://openrouter.ai/api/v1",
        "modelo": "", "env": "OPENROUTER_API_KEY", "gratis": True, "requiere_key": True,
        "nota": "Clave en openrouter.ai. Escribe el nombre de un modelo que termine en ':free' (cambian con frecuencia)."},
    "Ollama (local, gratis y privado)": {
        "tipo": "openai", "url": "http://localhost:11434/v1",
        "modelo": "llama3.1", "env": "", "gratis": True, "requiere_key": False,
        "nota": "Corre en tu computador (no funciona desde Streamlit Cloud). Los datos no salen de tu equipo."},
    "Otro compatible con OpenAI": {
        "tipo": "openai", "url": "", "modelo": "", "env": "LLM_API_KEY", "gratis": False, "requiere_key": True,
        "nota": "Cualquier servicio con API estilo OpenAI (/chat/completions): indica URL base, modelo y clave."},
    "Anthropic Claude (de pago)": {
        "tipo": "anthropic", "url": "", "modelo": "claude-sonnet-5", "env": "ANTHROPIC_API_KEY",
        "gratis": False, "requiere_key": True, "nota": "Requiere saldo en console.anthropic.com."},
}
PROVEEDOR_DEFECTO = "Google Gemini (plan gratuito)"


def _mensaje_usuario(ctx: Contexto, instruccion: str) -> str:
    return (f"CONTEXTO (extraído del Excel de soporte):\n{contexto_para_ia(ctx, instruccion)}\n\n"
            f"INSTRUCCIÓN DEL APARTADO:\n{instruccion}")


def analizar_con_claude(ctx: Contexto, instruccion: str, api_key: str,
                        modelo: str = "claude-sonnet-5") -> str:
    import anthropic

    cliente = anthropic.Anthropic(api_key=api_key)
    r = cliente.messages.create(
        model=modelo, max_tokens=900, system=SISTEMA,
        messages=[{"role": "user", "content": _mensaje_usuario(ctx, instruccion)}],
    )
    return "".join(b.text for b in r.content if getattr(b, "type", "") == "text").strip()


def analizar_openai_compatible(ctx: Contexto, instruccion: str, url_base: str, api_key: str,
                               modelo: str, timeout: int = 90) -> str:
    """Llama a cualquier API estilo OpenAI (Gemini, Groq, OpenRouter, Ollama…)."""
    import requests

    if not url_base or not modelo:
        raise ValueError("Falta la URL base o el nombre del modelo")
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"
    cuerpo = {"model": modelo, "temperature": 0.2, "max_tokens": 900,
              "messages": [{"role": "system", "content": SISTEMA},
                           {"role": "user", "content": _mensaje_usuario(ctx, instruccion)}]}
    r = requests.post(url_base.rstrip("/") + "/chat/completions", headers=headers, json=cuerpo, timeout=timeout)
    if r.status_code != 200:
        raise RuntimeError(f"Error {r.status_code} del proveedor: {r.text[:300]}")
    txt = (r.json()["choices"][0]["message"].get("content") or "").strip()
    txt = re.sub(r"<think>.*?</think>", "", txt, flags=re.S).strip()   # algunos modelos devuelven su razonamiento
    if not txt:
        raise RuntimeError("El modelo devolvió una respuesta vacía")
    return txt


def analizar(ctx: Contexto, instruccion: str, proveedor: str, api_key: str = "",
             modelo: str = "", url_base: str = "") -> str:
    """Punto único de entrada: elige el proveedor según PROVEEDORES."""
    cfg = PROVEEDORES[proveedor]
    modelo = modelo or cfg["modelo"]
    if cfg["tipo"] == "anthropic":
        return analizar_con_claude(ctx, instruccion, api_key, modelo)
    return analizar_openai_compatible(ctx, instruccion, url_base or cfg["url"], api_key, modelo)


def generar_analisis(ctx: Contexto, instrucciones: list[str],
                     analizador: Optional[Callable[[str], str]] = None) -> dict:
    return {i: analizador(i) for i in instrucciones} if analizador else {}
