#!/usr/bin/env python3
"""
Excel Automatizado - Resistencia de Materiales
Genera un archivo .xlsx con fórmulas automatizadas para:
- Propiedades de sección (6 tipos)
- Carga admisible con corte crítico
- Tabla multi-ejercicios
- Columna circular (ACI / NSR-84)
- Columnas en I
"""

import zipfile
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element, SubElement
import os
import math

# ============================================================
# UTILIDADES PARA GENERAR XLSX (formato OpenXML)
# ============================================================

NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace('', NAMESPACE)
ET.register_namespace('r', REL_NS)


def col_letter(col_idx):
    """Convert 0-based column index to Excel column letter."""
    result = ""
    while col_idx >= 0:
        result = chr(col_idx % 26 + 65) + result
        col_idx = col_idx // 26 - 1
    return result


def cell_ref(row, col):
    """Convert 0-based row,col to Excel cell reference like A1."""
    return f"{col_letter(col)}{row + 1}"



class XlsxWriter:
    """Minimal XLSX writer that supports formulas, strings, and numbers."""

    def __init__(self):
        self.sheets = []  # list of (name, cells_dict)
        self.shared_strings = []
        self.string_index = {}

    def add_sheet(self, name):
        """Add a new sheet and return its cells dictionary."""
        cells = {}
        self.sheets.append((name, cells))
        return cells

    def _get_string_index(self, s):
        if s not in self.string_index:
            self.string_index[s] = len(self.shared_strings)
            self.shared_strings.append(s)
        return self.string_index[s]

    def set_cell(self, cells, row, col, value, bold=False):
        """Set a cell value. value can be str, int, float, or formula (starts with =)."""
        cells[(row, col)] = value

    def save(self, filename):
        """Write the .xlsx file."""
        # First pass: generate all sheet XMLs (this populates shared_strings)
        sheet_xmls = []
        for i, (name, cells) in enumerate(self.sheets):
            sheet_xmls.append(self._sheet_xml(cells))

        with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr('[Content_Types].xml', self._content_types())
            zf.writestr('_rels/.rels', self._root_rels())
            zf.writestr('xl/workbook.xml', self._workbook())
            zf.writestr('xl/_rels/workbook.xml.rels', self._workbook_rels())
            zf.writestr('xl/styles.xml', self._styles())
            # Write sheets (already generated)
            for i, xml_content in enumerate(sheet_xmls):
                zf.writestr(f'xl/worksheets/sheet{i+1}.xml', xml_content)
            # Write shared strings AFTER sheets (so all strings are collected)
            zf.writestr('xl/sharedStrings.xml', self._shared_strings_xml())


    def _content_types(self):
        root = Element('Types')
        root.set('xmlns', CONTENT_NS)
        SubElement(root, 'Default', Extension='rels',
                   ContentType='application/vnd.openxmlformats-package.relationships+xml')
        SubElement(root, 'Default', Extension='xml', ContentType='application/xml')
        SubElement(root, 'Override', PartName='/xl/workbook.xml',
                   ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml')
        SubElement(root, 'Override', PartName='/xl/sharedStrings.xml',
                   ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml')
        SubElement(root, 'Override', PartName='/xl/styles.xml',
                   ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml')
        for i in range(len(self.sheets)):
            SubElement(root, 'Override', PartName=f'/xl/worksheets/sheet{i+1}.xml',
                       ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
        return ET.tostring(root, encoding='unicode', xml_declaration=True)

    def _root_rels(self):
        root = Element('Relationships')
        root.set('xmlns', RELS_NS)
        SubElement(root, 'Relationship', Id='rId1',
                   Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
                   Target='xl/workbook.xml')
        return ET.tostring(root, encoding='unicode', xml_declaration=True)


    def _workbook(self):
        root = Element('workbook')
        root.set('xmlns', NAMESPACE)
        root.set('xmlns:r', REL_NS)
        sheets_el = SubElement(root, 'sheets')
        for i, (name, _) in enumerate(self.sheets):
            s = SubElement(sheets_el, 'sheet')
            s.set('name', name)
            s.set('sheetId', str(i+1))
            s.set('r:id', f'rId{i+1}')
        return ET.tostring(root, encoding='unicode', xml_declaration=True)

    def _workbook_rels(self):
        root = Element('Relationships')
        root.set('xmlns', RELS_NS)
        for i in range(len(self.sheets)):
            SubElement(root, 'Relationship', Id=f'rId{i+1}',
                       Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                       Target=f'worksheets/sheet{i+1}.xml')
        SubElement(root, 'Relationship', Id=f'rId{len(self.sheets)+1}',
                   Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
                   Target='sharedStrings.xml')
        SubElement(root, 'Relationship', Id=f'rId{len(self.sheets)+2}',
                   Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
                   Target='styles.xml')
        return ET.tostring(root, encoding='unicode', xml_declaration=True)


    def _styles(self):
        root = Element('styleSheet')
        root.set('xmlns', NAMESPACE)
        # Fonts
        fonts = SubElement(root, 'fonts', count='2')
        font_normal = SubElement(fonts, 'font')
        SubElement(font_normal, 'sz', val='11')
        SubElement(font_normal, 'name', val='Calibri')
        font_bold = SubElement(fonts, 'font')
        SubElement(font_bold, 'b')
        SubElement(font_bold, 'sz', val='11')
        SubElement(font_bold, 'name', val='Calibri')
        # Fills
        fills = SubElement(root, 'fills', count='2')
        fill1 = SubElement(fills, 'fill')
        SubElement(fill1, 'patternFill', patternType='none')
        fill2 = SubElement(fills, 'fill')
        SubElement(fill2, 'patternFill', patternType='gray125')
        # Borders
        borders = SubElement(root, 'borders', count='1')
        border = SubElement(borders, 'border')
        SubElement(border, 'left')
        SubElement(border, 'right')
        SubElement(border, 'top')
        SubElement(border, 'bottom')
        SubElement(border, 'diagonal')
        # Cell style xfs
        csxfs = SubElement(root, 'cellStyleXfs', count='1')
        SubElement(csxfs, 'xf', numFmtId='0', fontId='0', fillId='0', borderId='0')
        # Cell xfs
        cxfs = SubElement(root, 'cellXfs', count='2')
        SubElement(cxfs, 'xf', numFmtId='0', fontId='0', fillId='0', borderId='0', xfId='0')
        xf_bold = SubElement(cxfs, 'xf', numFmtId='0', fontId='1', fillId='0', borderId='0', xfId='0')
        xf_bold.set('applyFont', '1')
        return ET.tostring(root, encoding='unicode', xml_declaration=True)


    def _shared_strings_xml(self):
        root = Element('sst')
        root.set('xmlns', NAMESPACE)
        root.set('count', str(len(self.shared_strings)))
        root.set('uniqueCount', str(len(self.shared_strings)))
        for s in self.shared_strings:
            si = SubElement(root, 'si')
            t = SubElement(si, 't')
            t.text = s
        return ET.tostring(root, encoding='unicode', xml_declaration=True)

    def _sheet_xml(self, cells):
        root = Element('worksheet')
        root.set('xmlns', NAMESPACE)
        sd = SubElement(root, 'sheetData')

        if not cells:
            return ET.tostring(root, encoding='unicode', xml_declaration=True)

        # Group cells by row
        rows_dict = {}
        for (r, c), val in cells.items():
            rows_dict.setdefault(r, []).append((c, val))

        for row_idx in sorted(rows_dict.keys()):
            row_el = SubElement(sd, 'row', r=str(row_idx + 1))
            for col_idx, val in sorted(rows_dict[row_idx]):
                ref = cell_ref(row_idx, col_idx)
                c_el = SubElement(row_el, 'c', r=ref)
                if isinstance(val, str) and val.startswith('='):
                    # Formula - no type attribute, Excel will evaluate
                    f_el = SubElement(c_el, 'f')
                    f_el.text = val[1:]  # remove leading =
                elif isinstance(val, (int, float)):
                    v_el = SubElement(c_el, 'v')
                    v_el.text = str(val)
                else:
                    # String - use shared strings
                    idx = self._get_string_index(str(val))
                    c_el.set('t', 's')
                    v_el = SubElement(c_el, 'v')
                    v_el.text = str(idx)

        return ET.tostring(root, encoding='unicode', xml_declaration=True)



# ============================================================
# HOJA 1: PROPIEDADES DE SECCIÓN
# ============================================================





def llenar_prop_seccion_completa(wb):
    """Llena la hoja Prop. Seccion con las 6 secciones."""
    cells = wb.add_sheet("Prop. Seccion")

    # === ENCABEZADO ===
    cells[(0, 0)] = "PROPIEDADES DE SECCIÓN - CÁLCULO AUTOMÁTICO"
    cells[(1, 0)] = "Universidad Nacional de Loja - Resistencia de Materiales"
    cells[(2, 0)] = "Ingrese los datos y los resultados se calculan automáticamente"
    cells[(3, 0)] = "π ="
    cells[(3, 1)] = 3.141592653589793

    # Referencia a PI en B4
    PI = "$B$4"

    # =========================================================
    # 1. SECCIÓN CIRCULAR SÓLIDA
    # =========================================================
    r = 5
    cells[(r, 0)] = "═══ 1. SECCIÓN CIRCULAR SÓLIDA ═══"
    cells[(r+1, 0)] = "DATOS:"
    cells[(r+2, 0)] = "r (radio)"
    cells[(r+2, 1)] = 0.05
    cells[(r+2, 2)] = "m"
    # r está en B8

    cells[(r+4, 0)] = "RESULTADOS (automáticos):"
    cells[(r+5, 0)] = "d (diámetro)"
    cells[(r+5, 1)] = "=2*B8"
    cells[(r+5, 2)] = "m"

    cells[(r+6, 0)] = "A (área) = π·r²"
    cells[(r+6, 1)] = f"={PI}*B8^2"
    cells[(r+6, 2)] = "m²"

    cells[(r+7, 0)] = "I (inercia) = π·r⁴/4"
    cells[(r+7, 1)] = f"={PI}*B8^4/4"
    cells[(r+7, 2)] = "m⁴"

    cells[(r+8, 0)] = "S (mód. sección) = π·r³/4"
    cells[(r+8, 1)] = f"={PI}*B8^3/4"
    cells[(r+8, 2)] = "m³"

    cells[(r+9, 0)] = "r_g (radio de giro) = r/2"
    cells[(r+9, 1)] = "=B8/2"
    cells[(r+9, 2)] = "m"


    # =========================================================
    # 2. SECCIÓN CIRCULAR HUECA
    # =========================================================
    r = 16
    cells[(r, 0)] = "═══ 2. SECCIÓN CIRCULAR HUECA ═══"
    cells[(r+1, 0)] = "DATOS:"
    cells[(r+2, 0)] = "R_ext (radio exterior)"
    cells[(r+2, 1)] = 0.075
    cells[(r+2, 2)] = "m"
    cells[(r+3, 0)] = "R_int (radio interior)"
    cells[(r+3, 1)] = 0.065
    cells[(r+3, 2)] = "m"
    # R_ext en B19, R_int en B20

    cells[(r+5, 0)] = "RESULTADOS (automáticos):"
    cells[(r+6, 0)] = "D_ext (diámetro ext.)"
    cells[(r+6, 1)] = "=2*B19"
    cells[(r+6, 2)] = "m"

    cells[(r+7, 0)] = "D_int (diámetro int.)"
    cells[(r+7, 1)] = "=2*B20"
    cells[(r+7, 2)] = "m"

    cells[(r+8, 0)] = "A (área) = π·(R²-r²)"
    cells[(r+8, 1)] = f"={PI}*(B19^2-B20^2)"
    cells[(r+8, 2)] = "m²"

    cells[(r+9, 0)] = "I (inercia) = π·(R⁴-r⁴)/4"
    cells[(r+9, 1)] = f"={PI}*(B19^4-B20^4)/4"
    cells[(r+9, 2)] = "m⁴"

    cells[(r+10, 0)] = "S (mód. sección) = I/R_ext"
    cells[(r+10, 1)] = f"=({PI}*(B19^4-B20^4)/4)/B19"
    cells[(r+10, 2)] = "m³"

    cells[(r+11, 0)] = "r_g (radio de giro) = √(I/A)"
    cells[(r+11, 1)] = f"=SQRT(({PI}*(B19^4-B20^4)/4)/({PI}*(B19^2-B20^2)))"
    cells[(r+11, 2)] = "m"


    # =========================================================
    # 3. SECCIÓN CUADRADA SÓLIDA
    # =========================================================
    r = 29
    cells[(r, 0)] = "═══ 3. SECCIÓN CUADRADA SÓLIDA ═══"
    cells[(r+1, 0)] = "DATOS:"
    cells[(r+2, 0)] = "a (lado)"
    cells[(r+2, 1)] = 0.10
    cells[(r+2, 2)] = "m"
    # a en B32

    cells[(r+4, 0)] = "RESULTADOS (automáticos):"
    cells[(r+5, 0)] = "A (área) = a²"
    cells[(r+5, 1)] = "=B32^2"
    cells[(r+5, 2)] = "m²"

    cells[(r+6, 0)] = "I (inercia) = a⁴/12"
    cells[(r+6, 1)] = "=B32^4/12"
    cells[(r+6, 2)] = "m⁴"

    cells[(r+7, 0)] = "S (mód. sección) = a³/6"
    cells[(r+7, 1)] = "=B32^3/6"
    cells[(r+7, 2)] = "m³"

    cells[(r+8, 0)] = "r_g (radio de giro) = a/√12"
    cells[(r+8, 1)] = "=B32/SQRT(12)"
    cells[(r+8, 2)] = "m"

    # =========================================================
    # 4. SECCIÓN CUADRADA HUECA
    # =========================================================
    r = 39
    cells[(r, 0)] = "═══ 4. SECCIÓN CUADRADA HUECA ═══"
    cells[(r+1, 0)] = "DATOS:"
    cells[(r+2, 0)] = "a_ext (lado exterior)"
    cells[(r+2, 1)] = 0.12
    cells[(r+2, 2)] = "m"
    cells[(r+3, 0)] = "a_int (lado interior)"
    cells[(r+3, 1)] = 0.10
    cells[(r+3, 2)] = "m"
    # a_ext en B42, a_int en B43

    cells[(r+5, 0)] = "RESULTADOS (automáticos):"
    cells[(r+6, 0)] = "A (área) = a_ext² - a_int²"
    cells[(r+6, 1)] = "=B42^2-B43^2"
    cells[(r+6, 2)] = "m²"

    cells[(r+7, 0)] = "I (inercia) = (a_ext⁴-a_int⁴)/12"
    cells[(r+7, 1)] = "=(B42^4-B43^4)/12"
    cells[(r+7, 2)] = "m⁴"

    cells[(r+8, 0)] = "S (mód. sección) = I/(a_ext/2)"
    cells[(r+8, 1)] = "=((B42^4-B43^4)/12)/(B42/2)"
    cells[(r+8, 2)] = "m³"

    cells[(r+9, 0)] = "r_g (radio de giro) = √(I/A)"
    cells[(r+9, 1)] = "=SQRT(((B42^4-B43^4)/12)/(B42^2-B43^2))"
    cells[(r+9, 2)] = "m"


    # =========================================================
    # 5. SECCIÓN RECTANGULAR SÓLIDA
    # =========================================================
    r = 50
    cells[(r, 0)] = "═══ 5. SECCIÓN RECTANGULAR SÓLIDA ═══"
    cells[(r+1, 0)] = "DATOS:"
    cells[(r+2, 0)] = "b (base)"
    cells[(r+2, 1)] = 0.08
    cells[(r+2, 2)] = "m"
    cells[(r+3, 0)] = "h (altura)"
    cells[(r+3, 1)] = 0.15
    cells[(r+3, 2)] = "m"
    # b en B53, h en B54

    cells[(r+5, 0)] = "RESULTADOS (automáticos):"
    cells[(r+6, 0)] = "A (área) = b·h"
    cells[(r+6, 1)] = "=B53*B54"
    cells[(r+6, 2)] = "m²"

    cells[(r+7, 0)] = "Ix (inercia eje x) = b·h³/12"
    cells[(r+7, 1)] = "=B53*B54^3/12"
    cells[(r+7, 2)] = "m⁴"

    cells[(r+8, 0)] = "Iy (inercia eje y) = h·b³/12"
    cells[(r+8, 1)] = "=B54*B53^3/12"
    cells[(r+8, 2)] = "m⁴"

    cells[(r+9, 0)] = "Sx (mód. sección x) = b·h²/6"
    cells[(r+9, 1)] = "=B53*B54^2/6"
    cells[(r+9, 2)] = "m³"

    cells[(r+10, 0)] = "Sy (mód. sección y) = h·b²/6"
    cells[(r+10, 1)] = "=B54*B53^2/6"
    cells[(r+10, 2)] = "m³"

    cells[(r+11, 0)] = "rx (radio giro x) = h/√12"
    cells[(r+11, 1)] = "=B54/SQRT(12)"
    cells[(r+11, 2)] = "m"

    cells[(r+12, 0)] = "ry (radio giro y) = b/√12"
    cells[(r+12, 1)] = "=B53/SQRT(12)"
    cells[(r+12, 2)] = "m"


    # =========================================================
    # 6. SECCIÓN RECTANGULAR HUECA
    # =========================================================
    r = 64
    cells[(r, 0)] = "═══ 6. SECCIÓN RECTANGULAR HUECA ═══"
    cells[(r+1, 0)] = "DATOS:"
    cells[(r+2, 0)] = "b_ext (base exterior)"
    cells[(r+2, 1)] = 0.12
    cells[(r+2, 2)] = "m"
    cells[(r+3, 0)] = "h_ext (altura exterior)"
    cells[(r+3, 1)] = 0.20
    cells[(r+3, 2)] = "m"
    cells[(r+4, 0)] = "b_int (base interior)"
    cells[(r+4, 1)] = 0.10
    cells[(r+4, 2)] = "m"
    cells[(r+5, 0)] = "h_int (altura interior)"
    cells[(r+5, 1)] = 0.18
    cells[(r+5, 2)] = "m"
    # b_ext=B67, h_ext=B68, b_int=B69, h_int=B70

    cells[(r+7, 0)] = "RESULTADOS (automáticos):"
    cells[(r+8, 0)] = "A (área) = b_ext·h_ext - b_int·h_int"
    cells[(r+8, 1)] = "=B67*B68-B69*B70"
    cells[(r+8, 2)] = "m²"

    cells[(r+9, 0)] = "Ix (inercia x) = (b_ext·h_ext³ - b_int·h_int³)/12"
    cells[(r+9, 1)] = "=(B67*B68^3-B69*B70^3)/12"
    cells[(r+9, 2)] = "m⁴"

    cells[(r+10, 0)] = "Iy (inercia y) = (h_ext·b_ext³ - h_int·b_int³)/12"
    cells[(r+10, 1)] = "=(B68*B67^3-B70*B69^3)/12"
    cells[(r+10, 2)] = "m⁴"

    cells[(r+11, 0)] = "Sx = Ix / (h_ext/2)"
    cells[(r+11, 1)] = "=((B67*B68^3-B69*B70^3)/12)/(B68/2)"
    cells[(r+11, 2)] = "m³"

    cells[(r+12, 0)] = "Sy = Iy / (b_ext/2)"
    cells[(r+12, 1)] = "=((B68*B67^3-B70*B69^3)/12)/(B67/2)"
    cells[(r+12, 2)] = "m³"

    cells[(r+13, 0)] = "rx = √(Ix/A)"
    cells[(r+13, 1)] = "=SQRT(((B67*B68^3-B69*B70^3)/12)/(B67*B68-B69*B70))"
    cells[(r+13, 2)] = "m"

    cells[(r+14, 0)] = "ry = √(Iy/A)"
    cells[(r+14, 1)] = "=SQRT(((B68*B67^3-B70*B69^3)/12)/(B67*B68-B69*B70))"
    cells[(r+14, 2)] = "m"

    return cells



# ============================================================
# HOJA 2: CARGA ADMISIBLE CON CORTE CRÍTICO
# ============================================================

def crear_hoja_carga_admisible(wb):
    """Crea la hoja de carga admisible con pandeo de Euler."""
    cells = wb.add_sheet("Carga admisible_cc")

    # Encabezados
    cells[(0, 0)] = "Universidad Nacional de Loja"
    cells[(1, 0)] = "Nombre: Israel Alvarado"
    cells[(2, 0)] = "Asignatura: Resistencia de materiales"
    cells[(3, 0)] = "CARGAS ADMISIBLES CON CORTE CRÍTICO - CÁLCULO AUTOMATIZADO"

    # === 1. DATOS DE ENTRADA ===
    cells[(5, 0)] = "1. DATOS DE ENTRADA"
    cells[(6, 0)] = "E — Módulo de elasticidad"
    cells[(6, 1)] = 200
    cells[(6, 2)] = "GPa"

    cells[(7, 0)] = "Fy — Esfuerzo de fluencia (A992)"
    cells[(7, 1)] = 345
    cells[(7, 2)] = "MPa"

    cells[(8, 0)] = "A — Área de la sección"
    cells[(8, 1)] = 0.0074
    cells[(8, 2)] = "m²"

    cells[(9, 0)] = "Ix — Momento de inercia, eje fuerte x"
    cells[(9, 1)] = 0.0000873
    cells[(9, 2)] = "m⁴"

    cells[(10, 0)] = "Iy — Momento de inercia, eje débil y"
    cells[(10, 1)] = 0.0000188
    cells[(10, 2)] = "m⁴"

    cells[(11, 0)] = "L — Longitud total de la columna"
    cells[(11, 1)] = 12.0
    cells[(11, 2)] = "m"

    cells[(12, 0)] = "Kx — Factor de longitud efectiva, eje x"
    cells[(12, 1)] = 1.0
    cells[(12, 2)] = "-"

    cells[(13, 0)] = "Lx,no-arriostrada — long. libre eje x"
    cells[(13, 1)] = 12.0
    cells[(13, 2)] = "m"

    cells[(14, 0)] = "Ky — Factor de longitud efectiva, eje y"
    cells[(14, 1)] = 1.0
    cells[(14, 2)] = "-"

    cells[(15, 0)] = "Ly,no-arriostrada — long. libre eje y"
    cells[(15, 1)] = 6.0
    cells[(15, 2)] = "m"

    cells[(16, 0)] = "F.S. — Factor de seguridad contra pandeo"
    cells[(16, 1)] = 2.0
    cells[(16, 2)] = "-"


    # === 2. DESARROLLO DEL CÁLCULO ===
    # Referencias: E=B7(GPa), Fy=B8(MPa), A=B9, Ix=B10, Iy=B11
    # L=B12, Kx=B13, Lx=B14, Ky=B15, Ly=B16, FS=B17

    cells[(18, 0)] = "2. DESARROLLO DEL CÁLCULO"
    cells[(19, 0)] = "E en Pa"
    cells[(19, 1)] = "=B7*1E9"
    cells[(19, 2)] = "Pa"

    cells[(20, 0)] = "Fy en Pa"
    cells[(20, 1)] = "=B8*1E6"
    cells[(20, 2)] = "Pa"

    cells[(21, 0)] = "(K·L)x — long. efectiva eje x"
    cells[(21, 1)] = "=B13*B14"
    cells[(21, 2)] = "m"

    cells[(22, 0)] = "(K·L)y — long. efectiva eje y"
    cells[(22, 1)] = "=B15*B16"
    cells[(22, 2)] = "m"

    cells[(23, 0)] = "Pcr,x = π²·E·Ix / (K·L)x²"
    cells[(23, 1)] = "=PI()^2*B20*B10/B22^2"
    cells[(23, 2)] = "N"

    cells[(24, 0)] = "Pcr,y = π²·E·Iy / (K·L)y²"
    cells[(24, 1)] = "=PI()^2*B20*B11/B23^2"
    cells[(24, 2)] = "N"

    cells[(25, 0)] = "Eje que gobierna el pandeo"
    cells[(25, 1)] = '=IF(B24<B25,"Eje x (fuerte)","Eje y (débil)")'

    cells[(26, 0)] = "Pcr (pandeo) = mín(Pcr,x; Pcr,y)"
    cells[(26, 1)] = "=MIN(B24,B25)"
    cells[(26, 2)] = "N"

    cells[(27, 0)] = "Pcr (pandeo) en kN"
    cells[(27, 1)] = "=B27/1000"
    cells[(27, 2)] = "kN"


    # === 3. VERIFICACIÓN POR FLUENCIA ===
    cells[(29, 0)] = "3. VERIFICACIÓN POR FLUENCIA"
    cells[(30, 0)] = "σ pandeo = Pcr/A"
    cells[(30, 1)] = "=B27/B9"
    cells[(30, 2)] = "Pa"

    cells[(31, 0)] = "σ en MPa"
    cells[(31, 1)] = "=B31/1E6"
    cells[(31, 2)] = "MPa"

    cells[(32, 0)] = "Pcr,fluencia = Fy · A"
    cells[(32, 1)] = "=B21*B9"
    cells[(32, 2)] = "N"

    cells[(33, 0)] = "¿σ < Fy? (si NO, rige fluencia)"
    cells[(33, 1)] = '=IF(B32<B8,"OK — rige pandeo (Euler)","NO — rige fluencia")'

    cells[(34, 0)] = "Pcr final (gobernante)"
    cells[(34, 1)] = '=IF(B32<B8,B27,B33)'
    cells[(34, 2)] = "N"

    # === 4. CARGA ADMISIBLE ===
    cells[(36, 0)] = "4. CARGA ADMISIBLE"
    cells[(37, 0)] = "P_admisible = Pcr,final / F.S."
    cells[(37, 1)] = "=B35/B17"
    cells[(37, 2)] = "N"

    cells[(38, 0)] = "P_admisible en kN"
    cells[(38, 1)] = "=B38/1000"
    cells[(38, 2)] = "kN"

    cells[(40, 0)] = "══════════════════════════════"
    cells[(41, 0)] = "RESULTADO: FUERZA ADMISIBLE P"
    cells[(41, 1)] = "=B39"
    cells[(41, 2)] = "kN"

    return cells



# ============================================================
# HOJA 3: TABLA MULTI-EJERCICIOS
# ============================================================

def crear_hoja_tabla_multi(wb):
    """Tabla para resolver múltiples ejercicios de pandeo."""
    cells = wb.add_sheet("Tabla Multi-Ejercicios")

    # Encabezados de la tabla
    headers = [
        "N.º", "Descripción / sección", "E (GPa)", "Fy (MPa)",
        "A (m²)", "Ix (m⁴)", "Iy (m⁴)", "Kx", "Lx no-arr. (m)",
        "Ky", "Ly no-arr. (m)", "(KL)x (m)", "(KL)y (m)",
        "Pcr,x (N)", "Pcr,y (N)", "Eje que gobierna",
        "Pcr pandeo (kN)", "σ pandeo (MPa)", "Pcr fluencia (kN)",
        "Modo de falla", "F.S.", "P admisible (kN)"
    ]

    cells[(0, 0)] = "TABLA MULTI-EJERCICIOS — Carga Admisible por Pandeo (Automatizada)"
    cells[(1, 0)] = "Complete las columnas A-K (datos) y las columnas L-V se calculan automáticamente"

    row_h = 3
    for col, h in enumerate(headers):
        cells[(row_h, col)] = h

    # Fila 1 con datos del ejemplo
    r = 4  # primera fila de datos (row 5 en Excel)
    cells[(r, 0)] = 1
    cells[(r, 1)] = "FP 13.3 — Columna A992, arriostrada eje débil a media altura"
    cells[(r, 2)] = 200
    cells[(r, 3)] = 345
    cells[(r, 4)] = 0.0074
    cells[(r, 5)] = 0.0000873
    cells[(r, 6)] = 0.0000188
    cells[(r, 7)] = 1
    cells[(r, 8)] = 12
    cells[(r, 9)] = 1
    cells[(r, 10)] = 6
    cells[(r, 20)] = 2.0


    # Fórmulas automáticas para las 10 filas
    for i in range(10):
        row = 4 + i
        rn = row + 1  # Excel row number (1-based)

        # Columnas de entrada: A=col0..K=col10, FS=col20
        # Col L (11): (KL)x = Kx * Lx = H*I
        cells[(row, 11)] = f"=IF(H{rn}=\"\",\"\",H{rn}*I{rn})"
        # Col M (12): (KL)y = Ky * Ly = J*K
        cells[(row, 12)] = f"=IF(J{rn}=\"\",\"\",J{rn}*K{rn})"
        # Col N (13): Pcr,x = π²·E·Ix/(KL)x²
        cells[(row, 13)] = f"=IF(L{rn}=\"\",\"\",PI()^2*(C{rn}*1E9)*F{rn}/L{rn}^2)"
        # Col O (14): Pcr,y = π²·E·Iy/(KL)y²
        cells[(row, 14)] = f"=IF(M{rn}=\"\",\"\",PI()^2*(C{rn}*1E9)*G{rn}/M{rn}^2)"
        # Col P (15): Eje que gobierna
        cells[(row, 15)] = f'=IF(N{rn}="","",IF(N{rn}<O{rn},"Eje x","Eje y"))'
        # Col Q (16): Pcr pandeo (kN)
        cells[(row, 16)] = f"=IF(N{rn}=\"\",\"\",MIN(N{rn},O{rn})/1000)"
        # Col R (17): σ pandeo (MPa) = Pcr/A / 1E6
        cells[(row, 17)] = f"=IF(Q{rn}=\"\",\"\",Q{rn}*1000/E{rn}/1E6)"
        # Col S (18): Pcr fluencia (kN) = Fy*A/1000
        cells[(row, 18)] = f"=IF(E{rn}=\"\",\"\",(D{rn}*1E6)*E{rn}/1000)"
        # Col T (19): Modo de falla
        cells[(row, 19)] = f'=IF(Q{rn}="","",IF(Q{rn}<S{rn},"Pandeo (Euler)","Fluencia"))'
        # Col V (21): P admisible (kN) = min(Pcr_pandeo, Pcr_fluencia) / FS
        cells[(row, 21)] = f"=IF(Q{rn}=\"\",\"\",MIN(Q{rn},S{rn})/U{rn})"

    # Filas 2-10: solo número y FS por defecto
    for i in range(1, 10):
        row = 4 + i
        cells[(row, 0)] = i + 1
        cells[(row, 20)] = 2.0

    return cells



# ============================================================
# HOJA 4: COLUMNA CIRCULAR (ACI / NSR-84)
# ============================================================

def crear_hoja_columna_circular(wb):
    """Columna circular con normas ACI y NSR-84."""
    cells = wb.add_sheet("Columna Circular")

    cells[(0, 0)] = "COLUMNA CIRCULAR — Normas ACI y NSR-84 (AUTOMATIZADO)"
    cells[(1, 0)] = "Calcular la carga axial permisible para la columna circular hueca"
    cells[(2, 0)] = "La columna es un tubo de acero empotrado en la base."

    # === DATOS ===
    cells[(4, 0)] = "═══ DATOS DE ENTRADA ═══"
    cells[(5, 0)] = "E ="
    cells[(5, 1)] = 200
    cells[(5, 2)] = "GPa"
    cells[(5, 3)] = "=B6*1E9"
    cells[(5, 4)] = "Pa"

    cells[(6, 0)] = "Fy ="
    cells[(6, 1)] = 250
    cells[(6, 2)] = "MPa"
    cells[(6, 3)] = "=B7*1E6"
    cells[(6, 4)] = "Pa"

    cells[(7, 0)] = "k (factor long. efectiva) ="
    cells[(7, 1)] = 2
    cells[(7, 2)] = "(empotrado-libre)"

    cells[(9, 0)] = "LONGITUDES:"
    cells[(10, 0)] = "L1 ="
    cells[(10, 1)] = 2
    cells[(10, 2)] = "m"

    cells[(11, 0)] = "L2 ="
    cells[(11, 1)] = 2.5
    cells[(11, 2)] = "m"

    cells[(12, 0)] = "L3 ="
    cells[(12, 1)] = 3.5
    cells[(12, 2)] = "m"

    cells[(14, 0)] = "DIMENSIONES:"
    cells[(15, 0)] = "D.int ="
    cells[(15, 1)] = 13
    cells[(15, 2)] = "cm"
    cells[(15, 3)] = "=B16/100"
    cells[(15, 4)] = "m"

    cells[(16, 0)] = "D.ext ="
    cells[(16, 1)] = 15
    cells[(16, 2)] = "cm"
    cells[(16, 3)] = "=B17/100"
    cells[(16, 4)] = "m"


    # === PROPIEDADES CALCULADAS ===
    cells[(18, 0)] = "═══ PROPIEDADES CALCULADAS ═══"
    # I = pi*(D_ext^4 - D_int^4)/64  (en m)
    cells[(19, 0)] = "I (inercia) ="
    cells[(19, 1)] = "=PI()*(D17^4-D16^4)/64"
    cells[(19, 2)] = "m⁴"

    # A = pi*(D_ext^2 - D_int^2)/4
    cells[(20, 0)] = "A (área) ="
    cells[(20, 1)] = "=PI()*(D17^2-D16^2)/4"
    cells[(20, 2)] = "m²"

    # r = sqrt(I/A)
    cells[(21, 0)] = "r (radio de giro) ="
    cells[(21, 1)] = "=SQRT(B20/B21)"
    cells[(21, 2)] = "m"

    # Cc = sqrt(2*pi^2*E/Fy)  (en Pa)
    cells[(22, 0)] = "Cc (relación de esbeltez crítica) ="
    cells[(22, 1)] = "=SQRT(2*PI()^2*D6/D7)"

    # === NORMA ACI — EULER ===
    cells[(24, 0)] = "═══ NORMA ACI (Euler) ═══"

    # Para cada longitud L1, L2, L3
    col_offset = [0, 4, 8]
    l_refs = ["B11", "B12", "B13"]  # L1, L2, L3

    for idx, (co, lr) in enumerate(zip(col_offset, l_refs)):
        r_base = 25
        letter_a = chr(65 + co)       # A, E, I
        letter_b = chr(65 + co + 1)   # B, F, J

        cells[(r_base, co)] = f"L{idx+1} ="
        cells[(r_base, co+1)] = f"={lr}"
        cells[(r_base, co+2)] = "m"

        cells[(r_base+1, co)] = "Le = k·L ="
        cells[(r_base+1, co+1)] = f"=B8*{lr}"
        cells[(r_base+1, co+2)] = "m"

        cells[(r_base+2, co)] = "kL/r ="
        cells[(r_base+2, co+1)] = f"=(B8*{lr})/B22"

        cells[(r_base+3, co)] = "Pcr (Euler) ="
        cells[(r_base+3, co+1)] = f"=PI()^2*D6*B20/(B8*{lr})^2"
        cells[(r_base+3, co+2)] = "N"

        cells[(r_base+4, co)] = "Pcr (kN) ="
        cells[(r_base+4, co+1)] = f"={chr(65+co+1)}{r_base+4}/1000"
        cells[(r_base+4, co+2)] = "kN"

        cells[(r_base+5, co)] = "σ cr ="
        cells[(r_base+5, co+1)] = f"={chr(65+co+1)}{r_base+4}/B21"
        cells[(r_base+5, co+2)] = "Pa"

        cells[(r_base+6, co)] = "σ cr (MPa) ="
        cells[(r_base+6, co+1)] = f"={chr(65+co+1)}{r_base+6}/1E6"
        cells[(r_base+6, co+2)] = "MPa"


    # === NORMA NSR-84 ===
    cells[(34, 0)] = "═══ NORMA NSR-84 ═══"

    for idx, (co, lr) in enumerate(zip(col_offset, l_refs)):
        r_base = 35

        cells[(r_base, co)] = f"L{idx+1} ="
        cells[(r_base, co+1)] = f"={lr}"
        cells[(r_base, co+2)] = "m"

        cells[(r_base+1, co)] = "Esbeltez kL/r ="
        cells[(r_base+1, co+1)] = f"=(B8*{lr})/B22"

        cells[(r_base+2, co)] = "Cc ="
        cells[(r_base+2, co+1)] = "=B23"

        cells[(r_base+3, co)] = "Comparación:"
        # Si kL/r < Cc: columna corta, si no: columna larga
        esb_ref = f"{chr(65+co+1)}{r_base+2}"
        cc_ref = f"{chr(65+co+1)}{r_base+3}"

        cells[(r_base+4, co)] = "Tipo columna:"
        cells[(r_base+4, co+1)] = f'=IF({esb_ref}<{cc_ref},"Corta (kL/r < Cc)","Larga (kL/r > Cc)")'

        # Fórmula NSR-84:
        # Si corta: σa = Fy * [1 - (kL/r)²/(2·Cc²)] / [5/3 + 3(kL/r)/(8Cc) - (kL/r)³/(8Cc³)]
        # Si larga: σa = 12·π²·E / (23·(kL/r)²)
        num_corta = f"(D7*(1-{esb_ref}^2/(2*{cc_ref}^2)))"
        den_corta = f"(5/3+3*{esb_ref}/(8*{cc_ref})-{esb_ref}^3/(8*{cc_ref}^3))"
        formula_larga = f"(12*PI()^2*D6/(23*{esb_ref}^2))"

        cells[(r_base+6, co)] = "σa (esfuerzo admisible) ="
        cells[(r_base+6, co+1)] = f"=IF({esb_ref}<{cc_ref},{num_corta}/{den_corta},{formula_larga})"
        cells[(r_base+6, co+2)] = "Pa"

        cells[(r_base+7, co)] = "σa (MPa) ="
        cells[(r_base+7, co+1)] = f"={chr(65+co+1)}{r_base+7}/1E6"
        cells[(r_base+7, co+2)] = "MPa"

        cells[(r_base+8, co)] = "P admisible ="
        cells[(r_base+8, co+1)] = f"={chr(65+co+1)}{r_base+7}*B21"
        cells[(r_base+8, co+2)] = "N"

        cells[(r_base+9, co)] = "P admisible (kN) ="
        cells[(r_base+9, co+1)] = f"={chr(65+co+1)}{r_base+9}/1000"
        cells[(r_base+9, co+2)] = "kN"


    # === TABLA COMPARATIVA ===
    cells[(47, 0)] = "═══ TABLA COMPARATIVA ACI vs NSR-84 ═══"
    cells[(48, 0)] = "Norma"
    cells[(48, 1)] = "L1"
    cells[(48, 2)] = "L2"
    cells[(48, 3)] = "L3"
    cells[(48, 4)] = "Unidad"

    cells[(49, 0)] = "ACI — Pcr"
    cells[(49, 1)] = "=B30"
    cells[(49, 2)] = "=F30"
    cells[(49, 3)] = "=J30"
    cells[(49, 4)] = "kN"

    cells[(50, 0)] = "ACI — σcr"
    cells[(50, 1)] = "=B32"
    cells[(50, 2)] = "=F32"
    cells[(50, 3)] = "=J32"
    cells[(50, 4)] = "MPa"

    cells[(51, 0)] = "NSR-84 — P adm"
    cells[(51, 1)] = "=B45"
    cells[(51, 2)] = "=F45"
    cells[(51, 3)] = "=J45"
    cells[(51, 4)] = "kN"

    cells[(52, 0)] = "NSR-84 — σa"
    cells[(52, 1)] = "=B43"
    cells[(52, 2)] = "=F43"
    cells[(52, 3)] = "=J43"
    cells[(52, 4)] = "MPa"

    return cells



# ============================================================
# HOJA 5: COLUMNAS EN I
# ============================================================

def crear_hoja_columnas_I(wb):
    """Columnas con perfil I — múltiples condiciones de apoyo."""
    cells = wb.add_sheet("Columnas en I")

    cells[(0, 0)] = "COLUMNAS EN I — ANÁLISIS CON DIFERENTES CONDICIONES DE APOYO"

    # === DATOS ===
    cells[(2, 0)] = "═══ DATOS DEL MATERIAL ═══"
    cells[(3, 0)] = "E ="
    cells[(3, 1)] = 30000000
    cells[(3, 2)] = "Lb/pulg² (PSI)"

    cells[(4, 0)] = "Fy ="
    cells[(4, 1)] = 20000
    cells[(4, 2)] = "Lb/pulg² (PSI)"

    cells[(6, 0)] = "═══ FACTORES K ═══"
    cells[(7, 0)] = "k1 (articulado-articulado) ="
    cells[(7, 1)] = 1
    cells[(8, 0)] = "k2 (empotrado-articulado) ="
    cells[(8, 1)] = 0.7
    cells[(9, 0)] = "k3 (empotrado-empotrado) ="
    cells[(9, 1)] = 0.5
    cells[(10, 0)] = "k4 (empotrado-libre) ="
    cells[(10, 1)] = 2

    cells[(12, 0)] = "═══ LONGITUD ═══"
    cells[(13, 0)] = "L ="
    cells[(13, 1)] = 2.5
    cells[(13, 2)] = "m"
    cells[(13, 3)] = "=B14*39.27"
    cells[(13, 4)] = "pulg"

    cells[(15, 0)] = "═══ PROPIEDADES DEL PERFIL (W10x60) ═══"
    cells[(16, 0)] = "Peso"
    cells[(16, 1)] = 60
    cells[(16, 2)] = "Lb/pie"

    cells[(17, 0)] = "A (área)"
    cells[(17, 1)] = 17.6
    cells[(17, 2)] = "pulg²"

    cells[(18, 0)] = "d (altura)"
    cells[(18, 1)] = 10.22
    cells[(18, 2)] = "pulg"

    cells[(19, 0)] = "tw (espesor alma)"
    cells[(19, 1)] = 0.42
    cells[(19, 2)] = "pulg"

    cells[(20, 0)] = "bf (ancho patín)"
    cells[(20, 1)] = 10.08
    cells[(20, 2)] = "pulg"

    cells[(21, 0)] = "tf (espesor patín)"
    cells[(21, 1)] = 0.68
    cells[(21, 2)] = "pulg"


    cells[(23, 0)] = "Ix ="
    cells[(23, 1)] = 341
    cells[(23, 2)] = "pulg⁴"

    cells[(24, 0)] = "Sx ="
    cells[(24, 1)] = 66.7
    cells[(24, 2)] = "pulg³"

    cells[(25, 0)] = "rx ="
    cells[(25, 1)] = 4.39
    cells[(25, 2)] = "pulg"

    cells[(26, 0)] = "Iy ="
    cells[(26, 1)] = 116
    cells[(26, 2)] = "pulg⁴"

    cells[(27, 0)] = "Sy ="
    cells[(27, 1)] = 23
    cells[(27, 2)] = "pulg³"

    cells[(28, 0)] = "ry ="
    cells[(28, 1)] = 2.57
    cells[(28, 2)] = "pulg"

    # === CÁLCULOS PARA CADA K ===
    cells[(30, 0)] = "═══ RESULTADOS POR CONDICIÓN DE APOYO ═══"

    # Headers
    cells[(31, 0)] = "Condición"
    cells[(31, 1)] = "k"
    cells[(31, 2)] = "Le (pulg)"
    cells[(31, 3)] = "kL/ry"
    cells[(31, 4)] = "Pcr,y (Lb)"
    cells[(31, 5)] = "σcr (PSI)"
    cells[(31, 6)] = "Cc"
    cells[(31, 7)] = "Tipo"
    cells[(31, 8)] = "σa NSR (PSI)"
    cells[(31, 9)] = "P adm (Lb)"

    # k refs: B8, B9, B10, B11
    k_names = ["Art-Art", "Emp-Art", "Emp-Emp", "Emp-Libre"]
    k_refs = ["B8", "B9", "B10", "B11"]

    for i, (kn, kr) in enumerate(zip(k_names, k_refs)):
        row = 32 + i
        rn = row + 1

        cells[(row, 0)] = kn
        cells[(row, 1)] = f"={kr}"
        # Le = k * L(pulg)
        cells[(row, 2)] = f"={kr}*D14"
        # kL/ry
        cells[(row, 3)] = f"=C{rn}/B29"
        # Pcr,y = π²·E·Iy / Le²
        cells[(row, 4)] = f"=PI()^2*B4*B27/C{rn}^2"
        # σcr = Pcr/A
        cells[(row, 5)] = f"=E{rn}/B18"
        # Cc
        cells[(row, 6)] = "=SQRT(2*PI()^2*B4/B5)"
        # Tipo
        cells[(row, 7)] = f'=IF(D{rn}<G{rn},"Corta","Larga")'
        # σa NSR
        # Corta: Fy[1-(kL/r)²/(2Cc²)] / [5/3 + 3(kL/r)/(8Cc) - (kL/r)³/(8Cc³)]
        # Larga: 12π²E/(23(kL/r)²)
        num = f"(B5*(1-D{rn}^2/(2*G{rn}^2)))"
        den = f"(5/3+3*D{rn}/(8*G{rn})-D{rn}^3/(8*G{rn}^3))"
        larga = f"(12*PI()^2*B4/(23*D{rn}^2))"
        cells[(row, 8)] = f"=IF(D{rn}<G{rn},{num}/{den},{larga})"
        # P adm = σa * A
        cells[(row, 9)] = f"=I{rn}*B18"

    return cells



# ============================================================
# MAIN — GENERAR EL ARCHIVO EXCEL
# ============================================================

def main():
    wb = XlsxWriter()

    # Hoja 1: Propiedades de Sección
    llenar_prop_seccion_completa(wb)

    # Hoja 2: Carga Admisible
    crear_hoja_carga_admisible(wb)

    # Hoja 3: Tabla Multi-Ejercicios
    crear_hoja_tabla_multi(wb)

    # Hoja 4: Columna Circular
    crear_hoja_columna_circular(wb)

    # Hoja 5: Columnas en I
    crear_hoja_columnas_I(wb)

    # Guardar
    output_file = "Excel_Automatizado_RM.xlsx"
    wb.save(output_file)
    print(f"✅ Archivo generado exitosamente: {output_file}")
    print(f"   Tamaño: {os.path.getsize(output_file)} bytes")
    print(f"   Hojas: {[name for name, _ in wb.sheets]}")
    print()
    print("📋 Contenido del archivo:")
    print("   1. Prop. Seccion — Propiedades de 6 tipos de sección (circular, circular hueca,")
    print("      cuadrada, cuadrada hueca, rectangular, rectangular hueca)")
    print("   2. Carga admisible_cc — Cálculo de pandeo de Euler con verificación de fluencia")
    print("   3. Tabla Multi-Ejercicios — Tabla para resolver hasta 10 ejercicios a la vez")
    print("   4. Columna Circular — Normas ACI y NSR-84 para 3 longitudes")
    print("   5. Columnas en I — Perfil W10x60 con 4 condiciones de apoyo")


if __name__ == "__main__":
    main()
