#!/usr/bin/env python3
"""
Excel Automatizado PROFESIONAL - Resistencia de Materiales
Universidad Nacional de Loja

Genera un archivo .xlsx con:
- Diseño profesional (colores, bordes, fuentes)
- Imágenes de secciones transversales
- Fórmulas explicadas y derivadas paso a paso
- Pasos numerados
- Celdas de entrada claramente marcadas
- Resultados automáticos
"""

import zipfile
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element, SubElement
import os, struct, zlib, math, io

# ============================================================
# NAMESPACES
# ============================================================
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_DWG = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

# ============================================================
# COLORES
# ============================================================
C_AZUL_DARK = "1B3A5C"
C_AZUL_MED = "2E5A88"
C_AZUL_LIGHT = "D6E8F7"
C_VERDE_DARK = "1D6B3A"
C_VERDE_LIGHT = "D4EDDA"
C_NARANJA = "E67E22"
C_NARANJA_LIGHT = "FDEBD0"
C_GRIS_DARK = "2C3E50"
C_GRIS_LIGHT = "F2F3F4"
C_BLANCO = "FFFFFF"
C_AMARILLO = "FFF9C4"
C_ROJO = "C0392B"

# ============================================================
# GENERADOR PNG SIMPLE (para diagramas de sección)
# ============================================================
def make_png(width, height, pixels_func):
    """Generate a PNG file from a pixel function."""
    raw = b''
    for y in range(height):
        raw += b'\x00'
        for x in range(width):
            r, g, b = pixels_func(x, y, width, height)
            raw += struct.pack('BBB', r, g, b)
    
    def chunk(ctype, data):
        c = ctype + data
        return struct.pack('>I', len(data)) + c + struct.pack('>I', zlib.crc32(c) & 0xFFFFFFFF)
    
    ihdr = struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0)
    return (b'\x89PNG\r\n\x1a\n' + 
            chunk(b'IHDR', ihdr) + 
            chunk(b'IDAT', zlib.compress(raw, 9)) + 
            chunk(b'IEND', b''))

def circle_img():
    def px(x, y, w, h):
        cx, cy, r = w//2, h//2, min(w,h)//2 - 8
        d = math.sqrt((x-cx)**2 + (y-cy)**2)
        if abs(d - r) < 3: return (27, 58, 92)
        if d < r: return (214, 232, 247)
        return (255, 255, 255)
    return make_png(100, 100, px)

def circle_hollow_img():
    def px(x, y, w, h):
        cx, cy = w//2, h//2
        r_out, r_in = min(w,h)//2 - 8, min(w,h)//2 - 20
        d = math.sqrt((x-cx)**2 + (y-cy)**2)
        if abs(d - r_out) < 3 or abs(d - r_in) < 3: return (27, 58, 92)
        if r_in < d < r_out: return (214, 232, 247)
        return (255, 255, 255)
    return make_png(100, 100, px)

def square_img():
    def px(x, y, w, h):
        cx, cy, s = w//2, h//2, min(w,h)//2 - 10
        if abs(x-cx) <= s and abs(y-cy) <= s:
            if abs(x-cx) >= s-3 or abs(y-cy) >= s-3: return (27, 58, 92)
            return (214, 232, 247)
        return (255, 255, 255)
    return make_png(100, 100, px)

def square_hollow_img():
    def px(x, y, w, h):
        cx, cy = w//2, h//2
        s_o, s_i = min(w,h)//2 - 10, min(w,h)//2 - 22
        in_o = abs(x-cx) <= s_o and abs(y-cy) <= s_o
        in_i = abs(x-cx) <= s_i and abs(y-cy) <= s_i
        if in_o and not in_i:
            if abs(x-cx) >= s_o-3 or abs(y-cy) >= s_o-3 or abs(x-cx) <= s_i+3 or abs(y-cy) <= s_i+3:
                return (27, 58, 92)
            return (214, 232, 247)
        if in_o and in_i:
            if abs(x-cx) >= s_i-2 or abs(y-cy) >= s_i-2: return (27, 58, 92)
        return (255, 255, 255)
    return make_png(100, 100, px)

def rect_img():
    def px(x, y, w, h):
        cx, cy = w//2, h//2
        hw, hh = w//2-8, h//2-15
        if abs(x-cx) <= hw and abs(y-cy) <= hh:
            if abs(x-cx) >= hw-3 or abs(y-cy) >= hh-3: return (27, 58, 92)
            return (214, 232, 247)
        return (255, 255, 255)
    return make_png(120, 80, px)

def rect_hollow_img():
    def px(x, y, w, h):
        cx, cy = w//2, h//2
        hw_o, hh_o = w//2-8, h//2-10
        hw_i, hh_i = w//2-20, h//2-22
        in_o = abs(x-cx) <= hw_o and abs(y-cy) <= hh_o
        in_i = abs(x-cx) <= hw_i and abs(y-cy) <= hh_i
        if in_o and not in_i:
            if abs(x-cx)>=hw_o-3 or abs(y-cy)>=hh_o-3 or abs(x-cx)<=hw_i+3 or abs(y-cy)<=hh_i+3:
                return (27, 58, 92)
            return (214, 232, 247)
        return (255, 255, 255)
    return make_png(120, 80, px)



# ============================================================
# XLSX WRITER CON ESTILOS
# ============================================================
def col_letter(c):
    r = ""
    while c >= 0:
        r = chr(c % 26 + 65) + r
        c = c // 26 - 1
    return r

def cref(row, col):
    return f"{col_letter(col)}{row+1}"

class StyledXlsx:
    """XLSX writer with full styling support."""
    
    # Style IDs (pre-defined)
    S_NORMAL = 0
    S_TITLE = 1          # Big blue title
    S_SECTION_HEADER = 2 # White on dark blue
    S_DATA_LABEL = 3     # Normal text on light blue bg
    S_DATA_INPUT = 4     # Bold on yellow bg with border (editable)
    S_RESULT_HEADER = 5  # White on dark green
    S_RESULT_VALUE = 6   # Bold green on light green bg
    S_NOTE = 7           # Small orange italic
    S_STEP_NUM = 8       # Bold blue number
    S_FORMULA_EXPLAIN = 9 # Italic gray
    S_UNIT = 10          # Gray right-aligned
    S_TABLE_HEADER = 11  # White bold on blue medium
    S_TABLE_DATA = 12    # Normal with thin border
    S_BIG_RESULT = 13    # Large bold green with border
    S_SUBTITLE = 14      # Bold blue medium
    S_EDITABLE_NOTE = 15 # Orange on yellow
    
    def __init__(self):
        self.sheets = []
        self.shared_strings = []
        self.string_map = {}
        self.images_data = []  # list of PNG bytes
        self.sheet_images = {}  # sheet_idx -> [(img_idx, col, row)]
    
    def add_sheet(self, name):
        cells = {}
        meta = {'merges': [], 'col_widths': {}, 'row_heights': {}}
        self.sheets.append((name, cells, meta))
        return cells, meta
    
    def add_image(self, sheet_idx, png_data, col, row):
        img_idx = len(self.images_data)
        self.images_data.append(png_data)
        if sheet_idx not in self.sheet_images:
            self.sheet_images[sheet_idx] = []
        self.sheet_images[sheet_idx].append((img_idx, col, row))
        return img_idx
    
    def _ss_idx(self, s):
        s = str(s)
        if s not in self.string_map:
            self.string_map[s] = len(self.shared_strings)
            self.shared_strings.append(s)
        return self.string_map[s]
    
    def save(self, path):
        # Pre-generate sheets to collect strings
        sheet_xmls = []
        for i, (name, cells, meta) in enumerate(self.sheets):
            sheet_xmls.append(self._make_sheet(cells, meta, i))
        
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr('[Content_Types].xml', self._content_types())
            zf.writestr('_rels/.rels', self._rels_root())
            zf.writestr('xl/workbook.xml', self._workbook())
            zf.writestr('xl/_rels/workbook.xml.rels', self._wb_rels())
            zf.writestr('xl/styles.xml', self._styles())
            zf.writestr('xl/sharedStrings.xml', self._shared_strings())
            
            for i, xml in enumerate(sheet_xmls):
                zf.writestr(f'xl/worksheets/sheet{i+1}.xml', xml)


    
    def _content_types(self):
        root = Element('Types', xmlns=NS_CT)
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
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(root, encoding='unicode')
    
    def _rels_root(self):
        root = Element('Relationships', xmlns=NS_REL)
        SubElement(root, 'Relationship', Id='rId1',
                   Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
                   Target='xl/workbook.xml')
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(root, encoding='unicode')
    
    def _workbook(self):
        # Build manually to avoid namespace issues
        xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        xml += f'<workbook xmlns="{NS}" xmlns:r="{NS_R}"><sheets>'
        for i, (name, _, _) in enumerate(self.sheets):
            xml += f'<sheet name="{name}" sheetId="{i+1}" r:id="rId{i+1}"/>'
        xml += '</sheets></workbook>'
        return xml
    
    def _wb_rels(self):
        root = Element('Relationships', xmlns=NS_REL)
        for i in range(len(self.sheets)):
            SubElement(root, 'Relationship', Id=f'rId{i+1}',
                       Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                       Target=f'worksheets/sheet{i+1}.xml')
        n = len(self.sheets)
        SubElement(root, 'Relationship', Id=f'rId{n+1}',
                   Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
                   Target='sharedStrings.xml')
        SubElement(root, 'Relationship', Id=f'rId{n+2}',
                   Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
                   Target='styles.xml')
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(root, encoding='unicode')
    
    def _shared_strings(self):
        root = Element('sst', xmlns=NS)
        root.set('count', str(len(self.shared_strings)))
        root.set('uniqueCount', str(len(self.shared_strings)))
        for s in self.shared_strings:
            si = SubElement(root, 'si')
            t = SubElement(si, 't')
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = s
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(root, encoding='unicode')


    
    def _styles(self):
        """Generate comprehensive styles XML."""
        xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="12">
  <font><sz val="11"/><name val="Calibri"/><color rgb="FF2C3E50"/></font>
  <font><b/><sz val="16"/><name val="Calibri"/><color rgb="FF1B3A5C"/></font>
  <font><b/><sz val="12"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font>
  <font><sz val="11"/><name val="Calibri"/><color rgb="FF2C3E50"/></font>
  <font><b/><sz val="11"/><name val="Calibri"/><color rgb="FF2C3E50"/></font>
  <font><b/><sz val="12"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font>
  <font><b/><sz val="11"/><name val="Calibri"/><color rgb="FF1D6B3A"/></font>
  <font><i/><sz val="10"/><name val="Calibri"/><color rgb="FFE67E22"/></font>
  <font><b/><sz val="11"/><name val="Calibri"/><color rgb="FF2E5A88"/></font>
  <font><i/><sz val="10"/><name val="Calibri"/><color rgb="FF7F8C8D"/></font>
  <font><sz val="10"/><name val="Calibri"/><color rgb="FF7F8C8D"/></font>
  <font><b/><sz val="14"/><name val="Calibri"/><color rgb="FF1D6B3A"/></font>
</fonts>
<fills count="11">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFD6E8F7"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFD4EDDA"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF1B3A5C"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFDEBD0"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFF9C4"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFF2F3F4"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF1D6B3A"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF2E5A88"/></patternFill></fill>
</fills>
<borders count="5">
  <border><left/><right/><top/><bottom/><diagonal/></border>
  <border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>
  <border><left/><right/><top/><bottom style="medium"><color auto="1"/></bottom><diagonal/></border>
  <border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="medium"><color auto="1"/></bottom><diagonal/></border>
  <border><left style="medium"><color auto="1"/></left><right style="medium"><color auto="1"/></right><top style="medium"><color auto="1"/></top><bottom style="medium"><color auto="1"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="16">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  <xf numFmtId="0" fontId="2" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  <xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="0" fontId="4" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="0" fontId="5" fillId="9" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  <xf numFmtId="0" fontId="6" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="0" fontId="7" fillId="5" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  <xf numFmtId="0" fontId="8" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  <xf numFmtId="0" fontId="9" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  <xf numFmtId="0" fontId="10" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  <xf numFmtId="0" fontId="2" fillId="10" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
  <xf numFmtId="0" fontId="11" fillId="3" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="0" fontId="8" fillId="0" borderId="2" xfId="0" applyFont="1" applyBorder="1"/>
  <xf numFmtId="0" fontId="7" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
</cellXfs>
</styleSheet>'''
        return xml


    
    def _make_sheet(self, cells, meta, sheet_idx):
        """Generate worksheet XML."""
        # Start building XML manually for proper namespace handling
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
        parts.append(f'<worksheet xmlns="{NS}" xmlns:r="{NS_R}">')
        
        # Column widths
        if meta.get('col_widths'):
            parts.append('<cols>')
            for col_idx, width in sorted(meta['col_widths'].items()):
                parts.append(f'<col min="{col_idx+1}" max="{col_idx+1}" width="{width}" customWidth="1"/>')
            parts.append('</cols>')
        
        parts.append('<sheetData>')
        
        if cells:
            rows_dict = {}
            for (r, c), (val, style) in cells.items():
                rows_dict.setdefault(r, []).append((c, val, style))
            
            for row_idx in sorted(rows_dict.keys()):
                ht = meta.get('row_heights', {}).get(row_idx)
                if ht:
                    parts.append(f'<row r="{row_idx+1}" ht="{ht}" customHeight="1">')
                else:
                    parts.append(f'<row r="{row_idx+1}">')
                
                for col_idx, val, style in sorted(rows_dict[row_idx]):
                    ref = cref(row_idx, col_idx)
                    s_attr = f' s="{style}"' if style is not None else ''
                    
                    if isinstance(val, str) and val.startswith('='):
                        formula = val[1:].replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
                        parts.append(f'<c r="{ref}"{s_attr}><f>{formula}</f></c>')
                    elif isinstance(val, (int, float)):
                        parts.append(f'<c r="{ref}"{s_attr}><v>{val}</v></c>')
                    elif val is not None and val != '':
                        idx = self._ss_idx(val)
                        parts.append(f'<c r="{ref}"{s_attr} t="s"><v>{idx}</v></c>')
                    else:
                        parts.append(f'<c r="{ref}"{s_attr}/>')
                
                parts.append('</row>')
        
        parts.append('</sheetData>')
        
        # Merge cells
        if meta.get('merges'):
            parts.append(f'<mergeCells count="{len(meta["merges"])}">')
            for merge in meta['merges']:
                parts.append(f'<mergeCell ref="{merge}"/>')
            parts.append('</mergeCells>')
        
        # Drawing reference (disabled - removed images for compatibility)
        # if sheet_idx in self.sheet_images:
        #     parts.append('<drawing r:id="rId1"/>')
        
        parts.append('</worksheet>')
        return '\n'.join(parts)
    
    def _sheet_rels(self, sheet_idx):
        root = Element('Relationships', xmlns=NS_REL)
        SubElement(root, 'Relationship', Id='rId1',
                   Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
                   Target=f'../drawings/drawing{sheet_idx+1}.xml')
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(root, encoding='unicode')
    
    def _drawing(self, sheet_idx):
        """Create drawing XML for images."""
        imgs = self.sheet_images[sheet_idx]
        lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
        lines.append(f'<xdr:wsDr xmlns:xdr="{NS_DWG}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">')
        
        for i, (img_idx, col, row) in enumerate(imgs):
            lines.append('<xdr:twoCellAnchor>')
            lines.append(f'<xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>50000</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>50000</xdr:rowOff></xdr:from>')
            lines.append(f'<xdr:to><xdr:col>{col+2}</xdr:col><xdr:colOff>50000</xdr:colOff><xdr:row>{row+5}</xdr:row><xdr:rowOff>50000</xdr:rowOff></xdr:to>')
            lines.append(f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{i+1}" name="Img{i+1}"/><xdr:cNvPicPr/></xdr:nvPicPr>')
            lines.append(f'<xdr:blipFill><a:blip r:embed="rId{i+1}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>')
            lines.append('<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>')
            lines.append('<xdr:clientData/></xdr:twoCellAnchor>')
        
        lines.append('</xdr:wsDr>')
        return '\n'.join(lines)
    
    def _drawing_rels(self, sheet_idx):
        root = Element('Relationships', xmlns=NS_REL)
        for i, (img_idx, _, _) in enumerate(self.sheet_images[sheet_idx]):
            SubElement(root, 'Relationship', Id=f'rId{i+1}',
                       Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                       Target=f'../media/image{img_idx+1}.png')
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(root, encoding='unicode')



# ============================================================
# HELPER: set cell with style
# ============================================================
S = StyledXlsx  # alias for style constants

def put(cells, row, col, val, style=None):
    cells[(row, col)] = (val, style)

# ============================================================
# HOJA 1: PROPIEDADES DE SECCIÓN (PROFESIONAL)
# ============================================================
def crear_hoja_propiedades(wb):
    cells, meta = wb.add_sheet("Prop. Seccion")
    
    # Column widths
    meta['col_widths'] = {0: 42, 1: 18, 2: 10, 3: 55, 4: 18}
    
    # === ENCABEZADO ===
    put(cells, 0, 0, "PROPIEDADES DE SECCIÓN TRANSVERSAL", S.S_TITLE)
    put(cells, 1, 0, "Universidad Nacional de Loja — Resistencia de Materiales", S.S_SUBTITLE)
    put(cells, 2, 0, "⚙ Ingrese los datos en las celdas AMARILLAS y los resultados se calculan automáticamente", S.S_NOTE)
    put(cells, 3, 0, "Las fórmulas están explicadas paso a paso en la columna D", S.S_FORMULA_EXPLAIN)
    
    # ────────────────────────────────────────────────────────
    # SECCIÓN 1: CIRCULAR SÓLIDA
    # ────────────────────────────────────────────────────────
    r = 5
    put(cells, r, 0, "1. SECCIÓN CIRCULAR SÓLIDA", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "Fórmula / Explicación", S.S_SECTION_HEADER)
    put(cells, r, 4, "", S.S_SECTION_HEADER)
    meta['row_heights'] = {r: 22}
    
    put(cells, r+1, 0, "   DATOS DE ENTRADA:", S.S_STEP_NUM)
    put(cells, r+1, 3, "¿Para qué sirve?", S.S_STEP_NUM)
    
    put(cells, r+2, 0, "   Paso 1: Ingrese el radio (r)", S.S_DATA_LABEL)
    put(cells, r+2, 1, 0.05, S.S_DATA_INPUT)
    put(cells, r+2, 2, "m", S.S_UNIT)
    put(cells, r+2, 3, "Radio de la sección circular. Dato principal de entrada.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+4, 0, "   RESULTADOS AUTOMÁTICOS:", S.S_RESULT_HEADER)
    put(cells, r+4, 1, "", S.S_RESULT_HEADER)
    put(cells, r+4, 2, "", S.S_RESULT_HEADER)
    put(cells, r+4, 3, "", S.S_RESULT_HEADER)
    
    put(cells, r+5, 0, "   Paso 2: Diámetro (d = 2r)", S.S_DATA_LABEL)
    put(cells, r+5, 1, "=2*B8", S.S_RESULT_VALUE)
    put(cells, r+5, 2, "m", S.S_UNIT)
    put(cells, r+5, 3, "Diámetro = 2 veces el radio. Dimensión total de la sección.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+6, 0, "   Paso 3: Área (A = π·r²)", S.S_DATA_LABEL)
    put(cells, r+6, 1, "=PI()*B8^2", S.S_RESULT_VALUE)
    put(cells, r+6, 2, "m²", S.S_UNIT)
    put(cells, r+6, 3, "Área transversal. Determina la resistencia a carga axial (P = σ·A).", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+7, 0, "   Paso 4: Inercia (I = π·r⁴/4)", S.S_DATA_LABEL)
    put(cells, r+7, 1, "=PI()*B8^4/4", S.S_RESULT_VALUE)
    put(cells, r+7, 2, "m⁴", S.S_UNIT)
    put(cells, r+7, 3, "Momento de inercia. Mide la resistencia a la flexión y pandeo.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+8, 0, "   Paso 5: Módulo de sección (S = π·r³/4)", S.S_DATA_LABEL)
    put(cells, r+8, 1, "=PI()*B8^3/4", S.S_RESULT_VALUE)
    put(cells, r+8, 2, "m³", S.S_UNIT)
    put(cells, r+8, 3, "S = I/c. Relaciona momento flector con esfuerzo máximo (σ = M/S).", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+9, 0, "   Paso 6: Radio de giro (rg = r/2)", S.S_DATA_LABEL)
    put(cells, r+9, 1, "=B8/2", S.S_RESULT_VALUE)
    put(cells, r+9, 2, "m", S.S_UNIT)
    put(cells, r+9, 3, "rg = √(I/A). Indica tendencia al pandeo. Menor rg = más propenso.", S.S_FORMULA_EXPLAIN)


    
    # ────────────────────────────────────────────────────────
    # SECCIÓN 2: CIRCULAR HUECA
    # ────────────────────────────────────────────────────────
    r = 16
    put(cells, r, 0, "2. SECCIÓN CIRCULAR HUECA (Tubo)", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "Fórmula / Explicación", S.S_SECTION_HEADER)
    put(cells, r, 4, "", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   DATOS DE ENTRADA:", S.S_STEP_NUM)
    put(cells, r+2, 0, "   Paso 1: Radio exterior (R)", S.S_DATA_LABEL)
    put(cells, r+2, 1, 0.075, S.S_DATA_INPUT)
    put(cells, r+2, 2, "m", S.S_UNIT)
    put(cells, r+2, 3, "Radio externo del tubo. Define el tamaño total.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+3, 0, "   Paso 2: Radio interior (r)", S.S_DATA_LABEL)
    put(cells, r+3, 1, 0.065, S.S_DATA_INPUT)
    put(cells, r+3, 2, "m", S.S_UNIT)
    put(cells, r+3, 3, "Radio interno (hueco). El espesor = R - r.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+5, 0, "   RESULTADOS AUTOMÁTICOS:", S.S_RESULT_HEADER)
    put(cells, r+5, 1, "", S.S_RESULT_HEADER)
    put(cells, r+5, 2, "", S.S_RESULT_HEADER)
    put(cells, r+5, 3, "", S.S_RESULT_HEADER)
    
    put(cells, r+6, 0, "   Paso 3: Área (A = π(R²-r²))", S.S_DATA_LABEL)
    put(cells, r+6, 1, "=PI()*(B19^2-B20^2)", S.S_RESULT_VALUE)
    put(cells, r+6, 2, "m²", S.S_UNIT)
    put(cells, r+6, 3, "Área del anillo = Área exterior - Área interior.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+7, 0, "   Paso 4: Inercia (I = π(R⁴-r⁴)/4)", S.S_DATA_LABEL)
    put(cells, r+7, 1, "=PI()*(B19^4-B20^4)/4", S.S_RESULT_VALUE)
    put(cells, r+7, 2, "m⁴", S.S_UNIT)
    put(cells, r+7, 3, "Inercia del tubo. Mayor que sólido del mismo peso (más eficiente).", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+8, 0, "   Paso 5: Módulo de sección (S = I/R)", S.S_DATA_LABEL)
    put(cells, r+8, 1, "=(PI()*(B19^4-B20^4)/4)/B19", S.S_RESULT_VALUE)
    put(cells, r+8, 2, "m³", S.S_UNIT)
    put(cells, r+8, 3, "S = I/c donde c = R (distancia al borde exterior).", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+9, 0, "   Paso 6: Radio de giro (rg = √(I/A))", S.S_DATA_LABEL)
    put(cells, r+9, 1, "=SQRT((PI()*(B19^4-B20^4)/4)/(PI()*(B19^2-B20^2)))", S.S_RESULT_VALUE)
    put(cells, r+9, 2, "m", S.S_UNIT)
    put(cells, r+9, 3, "rg = √(I/A). Un tubo tiene mayor rg que sólido → resiste más pandeo.", S.S_FORMULA_EXPLAIN)
    
    # ────────────────────────────────────────────────────────
    # SECCIÓN 3: CUADRADA SÓLIDA
    # ────────────────────────────────────────────────────────
    r = 27
    put(cells, r, 0, "3. SECCIÓN CUADRADA SÓLIDA", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "Fórmula / Explicación", S.S_SECTION_HEADER)
    put(cells, r, 4, "", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   DATOS DE ENTRADA:", S.S_STEP_NUM)
    put(cells, r+2, 0, "   Paso 1: Lado (a)", S.S_DATA_LABEL)
    put(cells, r+2, 1, 0.10, S.S_DATA_INPUT)
    put(cells, r+2, 2, "m", S.S_UNIT)
    put(cells, r+2, 3, "Longitud del lado del cuadrado.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+4, 0, "   RESULTADOS AUTOMÁTICOS:", S.S_RESULT_HEADER)
    put(cells, r+4, 1, "", S.S_RESULT_HEADER)
    put(cells, r+4, 2, "", S.S_RESULT_HEADER)
    put(cells, r+4, 3, "", S.S_RESULT_HEADER)
    
    put(cells, r+5, 0, "   Paso 2: Área (A = a²)", S.S_DATA_LABEL)
    put(cells, r+5, 1, "=B30^2", S.S_RESULT_VALUE)
    put(cells, r+5, 2, "m²", S.S_UNIT)
    put(cells, r+5, 3, "Área = lado al cuadrado.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+6, 0, "   Paso 3: Inercia (I = a⁴/12)", S.S_DATA_LABEL)
    put(cells, r+6, 1, "=B30^4/12", S.S_RESULT_VALUE)
    put(cells, r+6, 2, "m⁴", S.S_UNIT)
    put(cells, r+6, 3, "I = b·h³/12. Como b=h=a → I = a⁴/12.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+7, 0, "   Paso 4: Módulo de sección (S = a³/6)", S.S_DATA_LABEL)
    put(cells, r+7, 1, "=B30^3/6", S.S_RESULT_VALUE)
    put(cells, r+7, 2, "m³", S.S_UNIT)
    put(cells, r+7, 3, "S = I/c = (a⁴/12)/(a/2) = a³/6.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+8, 0, "   Paso 5: Radio de giro (rg = a/√12)", S.S_DATA_LABEL)
    put(cells, r+8, 1, "=B30/SQRT(12)", S.S_RESULT_VALUE)
    put(cells, r+8, 2, "m", S.S_UNIT)
    put(cells, r+8, 3, "rg = √(I/A) = √(a⁴/12 / a²) = a/√12 ≈ 0.289·a.", S.S_FORMULA_EXPLAIN)


    
    # ────────────────────────────────────────────────────────
    # SECCIÓN 4: CUADRADA HUECA
    # ────────────────────────────────────────────────────────
    r = 37
    put(cells, r, 0, "4. SECCIÓN CUADRADA HUECA", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "Fórmula / Explicación", S.S_SECTION_HEADER)
    put(cells, r, 4, "", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   DATOS DE ENTRADA:", S.S_STEP_NUM)
    put(cells, r+2, 0, "   Paso 1: Lado exterior (a_ext)", S.S_DATA_LABEL)
    put(cells, r+2, 1, 0.12, S.S_DATA_INPUT)
    put(cells, r+2, 2, "m", S.S_UNIT)
    put(cells, r+2, 3, "Dimensión exterior del perfil tubular cuadrado.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+3, 0, "   Paso 2: Lado interior (a_int)", S.S_DATA_LABEL)
    put(cells, r+3, 1, 0.10, S.S_DATA_INPUT)
    put(cells, r+3, 2, "m", S.S_UNIT)
    put(cells, r+3, 3, "Dimensión interior. Espesor de pared = (a_ext - a_int)/2.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+5, 0, "   RESULTADOS AUTOMÁTICOS:", S.S_RESULT_HEADER)
    put(cells, r+5, 1, "", S.S_RESULT_HEADER)
    put(cells, r+5, 2, "", S.S_RESULT_HEADER)
    put(cells, r+5, 3, "", S.S_RESULT_HEADER)
    
    put(cells, r+6, 0, "   Paso 3: Área (A = a_ext² - a_int²)", S.S_DATA_LABEL)
    put(cells, r+6, 1, "=B40^2-B41^2", S.S_RESULT_VALUE)
    put(cells, r+6, 2, "m²", S.S_UNIT)
    put(cells, r+6, 3, "Área neta = Área exterior menos hueco interior.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+7, 0, "   Paso 4: Inercia (I = (a_ext⁴-a_int⁴)/12)", S.S_DATA_LABEL)
    put(cells, r+7, 1, "=(B40^4-B41^4)/12", S.S_RESULT_VALUE)
    put(cells, r+7, 2, "m⁴", S.S_UNIT)
    put(cells, r+7, 3, "I = I_exterior - I_interior = a_ext⁴/12 - a_int⁴/12.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+8, 0, "   Paso 5: Módulo de sección (S = I/(a_ext/2))", S.S_DATA_LABEL)
    put(cells, r+8, 1, "=((B40^4-B41^4)/12)/(B40/2)", S.S_RESULT_VALUE)
    put(cells, r+8, 2, "m³", S.S_UNIT)
    put(cells, r+8, 3, "S = I/c donde c = a_ext/2 (distancia al borde).", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+9, 0, "   Paso 6: Radio de giro (rg = √(I/A))", S.S_DATA_LABEL)
    put(cells, r+9, 1, "=SQRT(((B40^4-B41^4)/12)/(B40^2-B41^2))", S.S_RESULT_VALUE)
    put(cells, r+9, 2, "m", S.S_UNIT)
    put(cells, r+9, 3, "rg = √(I/A). Tubular hueco tiene mejor rg que sólido.", S.S_FORMULA_EXPLAIN)
    
    # ────────────────────────────────────────────────────────
    # SECCIÓN 5: RECTANGULAR SÓLIDA
    # ────────────────────────────────────────────────────────
    r = 48
    put(cells, r, 0, "5. SECCIÓN RECTANGULAR SÓLIDA", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "Fórmula / Explicación", S.S_SECTION_HEADER)
    put(cells, r, 4, "", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   DATOS DE ENTRADA:", S.S_STEP_NUM)
    put(cells, r+2, 0, "   Paso 1: Base (b)", S.S_DATA_LABEL)
    put(cells, r+2, 1, 0.08, S.S_DATA_INPUT)
    put(cells, r+2, 2, "m", S.S_UNIT)
    put(cells, r+2, 3, "Ancho de la sección (dimensión horizontal).", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+3, 0, "   Paso 2: Altura (h)", S.S_DATA_LABEL)
    put(cells, r+3, 1, 0.15, S.S_DATA_INPUT)
    put(cells, r+3, 2, "m", S.S_UNIT)
    put(cells, r+3, 3, "Alto de la sección (dimensión vertical).", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+5, 0, "   RESULTADOS AUTOMÁTICOS:", S.S_RESULT_HEADER)
    put(cells, r+5, 1, "", S.S_RESULT_HEADER)
    put(cells, r+5, 2, "", S.S_RESULT_HEADER)
    put(cells, r+5, 3, "", S.S_RESULT_HEADER)
    
    put(cells, r+6, 0, "   Paso 3: Área (A = b·h)", S.S_DATA_LABEL)
    put(cells, r+6, 1, "=B51*B52", S.S_RESULT_VALUE)
    put(cells, r+6, 2, "m²", S.S_UNIT)
    put(cells, r+6, 3, "Área = base por altura.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+7, 0, "   Paso 4: Ix (inercia eje x) = b·h³/12", S.S_DATA_LABEL)
    put(cells, r+7, 1, "=B51*B52^3/12", S.S_RESULT_VALUE)
    put(cells, r+7, 2, "m⁴", S.S_UNIT)
    put(cells, r+7, 3, "Inercia respecto al eje x (horizontal). Resiste flexión vertical.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+8, 0, "   Paso 5: Iy (inercia eje y) = h·b³/12", S.S_DATA_LABEL)
    put(cells, r+8, 1, "=B52*B51^3/12", S.S_RESULT_VALUE)
    put(cells, r+8, 2, "m⁴", S.S_UNIT)
    put(cells, r+8, 3, "Inercia respecto al eje y (vertical). Resiste flexión horizontal.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+9, 0, "   Paso 6: Sx = b·h²/6", S.S_DATA_LABEL)
    put(cells, r+9, 1, "=B51*B52^2/6", S.S_RESULT_VALUE)
    put(cells, r+9, 2, "m³", S.S_UNIT)
    put(cells, r+9, 3, "Módulo de sección eje x. σ_max = M/Sx.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+10, 0, "   Paso 7: Sy = h·b²/6", S.S_DATA_LABEL)
    put(cells, r+10, 1, "=B52*B51^2/6", S.S_RESULT_VALUE)
    put(cells, r+10, 2, "m³", S.S_UNIT)
    put(cells, r+10, 3, "Módulo de sección eje y.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+11, 0, "   Paso 8: rx = h/√12", S.S_DATA_LABEL)
    put(cells, r+11, 1, "=B52/SQRT(12)", S.S_RESULT_VALUE)
    put(cells, r+11, 2, "m", S.S_UNIT)
    put(cells, r+11, 3, "Radio de giro eje x = √(Ix/A).", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+12, 0, "   Paso 9: ry = b/√12", S.S_DATA_LABEL)
    put(cells, r+12, 1, "=B51/SQRT(12)", S.S_RESULT_VALUE)
    put(cells, r+12, 2, "m", S.S_UNIT)
    put(cells, r+12, 3, "Radio de giro eje y = √(Iy/A). Si ry < rx, pandea en eje y.", S.S_FORMULA_EXPLAIN)


    
    # ────────────────────────────────────────────────────────
    # SECCIÓN 6: RECTANGULAR HUECA
    # ────────────────────────────────────────────────────────
    r = 62
    put(cells, r, 0, "6. SECCIÓN RECTANGULAR HUECA", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "Fórmula / Explicación", S.S_SECTION_HEADER)
    put(cells, r, 4, "", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   DATOS DE ENTRADA:", S.S_STEP_NUM)
    put(cells, r+2, 0, "   Paso 1: Base exterior (b_ext)", S.S_DATA_LABEL)
    put(cells, r+2, 1, 0.12, S.S_DATA_INPUT)
    put(cells, r+2, 2, "m", S.S_UNIT)
    put(cells, r+2, 3, "Ancho exterior del perfil.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+3, 0, "   Paso 2: Altura exterior (h_ext)", S.S_DATA_LABEL)
    put(cells, r+3, 1, 0.20, S.S_DATA_INPUT)
    put(cells, r+3, 2, "m", S.S_UNIT)
    put(cells, r+3, 3, "Altura exterior del perfil.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+4, 0, "   Paso 3: Base interior (b_int)", S.S_DATA_LABEL)
    put(cells, r+4, 1, 0.10, S.S_DATA_INPUT)
    put(cells, r+4, 2, "m", S.S_UNIT)
    put(cells, r+4, 3, "Ancho del hueco interior.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+5, 0, "   Paso 4: Altura interior (h_int)", S.S_DATA_LABEL)
    put(cells, r+5, 1, 0.18, S.S_DATA_INPUT)
    put(cells, r+5, 2, "m", S.S_UNIT)
    put(cells, r+5, 3, "Altura del hueco interior.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+7, 0, "   RESULTADOS AUTOMÁTICOS:", S.S_RESULT_HEADER)
    put(cells, r+7, 1, "", S.S_RESULT_HEADER)
    put(cells, r+7, 2, "", S.S_RESULT_HEADER)
    put(cells, r+7, 3, "", S.S_RESULT_HEADER)
    
    # b_ext=B65, h_ext=B66, b_int=B67, h_int=B68
    put(cells, r+8, 0, "   Paso 5: Área = b_ext·h_ext - b_int·h_int", S.S_DATA_LABEL)
    put(cells, r+8, 1, "=B65*B66-B67*B68", S.S_RESULT_VALUE)
    put(cells, r+8, 2, "m²", S.S_UNIT)
    put(cells, r+8, 3, "Área total exterior menos área del hueco.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+9, 0, "   Paso 6: Ix = (b_ext·h_ext³ - b_int·h_int³)/12", S.S_DATA_LABEL)
    put(cells, r+9, 1, "=(B65*B66^3-B67*B68^3)/12", S.S_RESULT_VALUE)
    put(cells, r+9, 2, "m⁴", S.S_UNIT)
    put(cells, r+9, 3, "Inercia eje x = Ix_exterior - Ix_interior.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+10, 0, "   Paso 7: Iy = (h_ext·b_ext³ - h_int·b_int³)/12", S.S_DATA_LABEL)
    put(cells, r+10, 1, "=(B66*B65^3-B68*B67^3)/12", S.S_RESULT_VALUE)
    put(cells, r+10, 2, "m⁴", S.S_UNIT)
    put(cells, r+10, 3, "Inercia eje y = Iy_exterior - Iy_interior.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+11, 0, "   Paso 8: Sx = Ix / (h_ext/2)", S.S_DATA_LABEL)
    put(cells, r+11, 1, "=((B65*B66^3-B67*B68^3)/12)/(B66/2)", S.S_RESULT_VALUE)
    put(cells, r+11, 2, "m³", S.S_UNIT)
    put(cells, r+11, 3, "Módulo de sección eje x. c = h_ext/2.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+12, 0, "   Paso 9: Sy = Iy / (b_ext/2)", S.S_DATA_LABEL)
    put(cells, r+12, 1, "=((B66*B65^3-B68*B67^3)/12)/(B65/2)", S.S_RESULT_VALUE)
    put(cells, r+12, 2, "m³", S.S_UNIT)
    put(cells, r+12, 3, "Módulo de sección eje y. c = b_ext/2.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+13, 0, "   Paso 10: rx = √(Ix/A)", S.S_DATA_LABEL)
    put(cells, r+13, 1, "=SQRT(((B65*B66^3-B67*B68^3)/12)/(B65*B66-B67*B68))", S.S_RESULT_VALUE)
    put(cells, r+13, 2, "m", S.S_UNIT)
    put(cells, r+13, 3, "Radio de giro eje x.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+14, 0, "   Paso 11: ry = √(Iy/A)", S.S_DATA_LABEL)
    put(cells, r+14, 1, "=SQRT(((B66*B65^3-B68*B67^3)/12)/(B65*B66-B67*B68))", S.S_RESULT_VALUE)
    put(cells, r+14, 2, "m", S.S_UNIT)
    put(cells, r+14, 3, "Radio de giro eje y. El menor rige el pandeo.", S.S_FORMULA_EXPLAIN)
    
    # Add visual text representations instead of images
    put(cells, 7, 4, "    ●", S.S_STEP_NUM)
    put(cells, 8, 4, "  (   )", S.S_STEP_NUM)
    
    put(cells, 18, 4, "  ◎", S.S_STEP_NUM)
    put(cells, 19, 4, " (○)", S.S_STEP_NUM)
    
    put(cells, 29, 4, "  ■", S.S_STEP_NUM)
    put(cells, 30, 4, " [  ]", S.S_STEP_NUM)
    
    put(cells, 39, 4, "  ▣", S.S_STEP_NUM)
    put(cells, 40, 4, " [□]", S.S_STEP_NUM)
    
    put(cells, 50, 4, "  ▬", S.S_STEP_NUM)
    put(cells, 51, 4, " [==]", S.S_STEP_NUM)
    
    put(cells, 64, 4, "  ▭", S.S_STEP_NUM)
    put(cells, 65, 4, " [═]", S.S_STEP_NUM)
    
    return cells



# ============================================================
# HOJA 2: CARGA ADMISIBLE (PROFESIONAL)
# ============================================================
def crear_hoja_carga_admisible(wb):
    cells, meta = wb.add_sheet("Carga admisible")
    meta['col_widths'] = {0: 45, 1: 18, 2: 10, 3: 55}
    
    put(cells, 0, 0, "CARGA ADMISIBLE CON PANDEO DE EULER", S.S_TITLE)
    put(cells, 1, 0, "Universidad Nacional de Loja — Resistencia de Materiales", S.S_SUBTITLE)
    put(cells, 2, 0, "⚙ Modifique las celdas AMARILLAS para su ejercicio. Todo se recalcula.", S.S_NOTE)
    
    # DATOS DE ENTRADA
    r = 4
    put(cells, r, 0, "PASO 1: DATOS DE ENTRADA", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "¿Qué significa?", S.S_SECTION_HEADER)
    
    datos = [
        ("   1.1 E — Módulo de elasticidad", 200, "GPa", "Rigidez del material. Acero ≈ 200 GPa, Aluminio ≈ 70 GPa."),
        ("   1.2 Fy — Esfuerzo de fluencia", 345, "MPa", "Esfuerzo máximo antes de deformación permanente. A992 = 345 MPa."),
        ("   1.3 A — Área de la sección", 0.0074, "m²", "Área transversal del perfil. Se obtiene de tablas AISC o cálculo."),
        ("   1.4 Ix — Inercia eje fuerte x", 0.0000873, "m⁴", "Momento de inercia mayor. Resiste pandeo en eje x."),
        ("   1.5 Iy — Inercia eje débil y", 0.0000188, "m⁴", "Momento de inercia menor. Eje débil = donde pandea primero."),
        ("   1.6 L — Longitud de la columna", 12.0, "m", "Longitud total entre apoyos de la columna."),
        ("   1.7 Kx — Factor long. efectiva eje x", 1.0, "-", "K depende de apoyos: Art-Art=1, Emp-Emp=0.5, Emp-Libre=2."),
        ("   1.8 Lx — Long. no arriostrada eje x", 12.0, "m", "Longitud libre sin soporte lateral en eje x."),
        ("   1.9 Ky — Factor long. efectiva eje y", 1.0, "-", "Factor K para el eje débil."),
        ("   1.10 Ly — Long. no arriostrada eje y", 6.0, "m", "Si arriostrado a media altura: Ly = L/2."),
        ("   1.11 F.S. — Factor de seguridad", 2.0, "-", "Factor de seguridad contra pandeo. Típico = 1.67 a 2.0."),
    ]
    
    for i, (label, val, unit, explain) in enumerate(datos):
        row = r + 1 + i
        put(cells, row, 0, label, S.S_DATA_LABEL)
        put(cells, row, 1, val, S.S_DATA_INPUT)
        put(cells, row, 2, unit, S.S_UNIT)
        put(cells, row, 3, explain, S.S_FORMULA_EXPLAIN)
    
    # DESARROLLO
    r = 17
    put(cells, r, 0, "PASO 2: DESARROLLO DEL CÁLCULO", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "Procedimiento paso a paso", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   2.1 Convertir E a Pascales", S.S_DATA_LABEL)
    put(cells, r+1, 1, "=B6*1E9", S.S_RESULT_VALUE)
    put(cells, r+1, 2, "Pa", S.S_UNIT)
    put(cells, r+1, 3, "E(Pa) = E(GPa) × 10⁹. Necesario para que unidades sean consistentes.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+2, 0, "   2.2 Convertir Fy a Pascales", S.S_DATA_LABEL)
    put(cells, r+2, 1, "=B7*1E6", S.S_RESULT_VALUE)
    put(cells, r+2, 2, "Pa", S.S_UNIT)
    put(cells, r+2, 3, "Fy(Pa) = Fy(MPa) × 10⁶.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+3, 0, "   2.3 Longitud efectiva eje x: (KL)x", S.S_DATA_LABEL)
    put(cells, r+3, 1, "=B11*B12", S.S_RESULT_VALUE)
    put(cells, r+3, 2, "m", S.S_UNIT)
    put(cells, r+3, 3, "(KL)x = Kx × Lx. Longitud equivalente para pandeo.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+4, 0, "   2.4 Longitud efectiva eje y: (KL)y", S.S_DATA_LABEL)
    put(cells, r+4, 1, "=B13*B14", S.S_RESULT_VALUE)
    put(cells, r+4, 2, "m", S.S_UNIT)
    put(cells, r+4, 3, "(KL)y = Ky × Ly.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+5, 0, "   2.5 Pcr,x = π²·E·Ix / (KL)x²", S.S_DATA_LABEL)
    put(cells, r+5, 1, "=PI()^2*B19*B9/(B21^2)", S.S_RESULT_VALUE)
    put(cells, r+5, 2, "N", S.S_UNIT)
    put(cells, r+5, 3, "Carga crítica de Euler eje x. Fórmula fundamental de pandeo.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+6, 0, "   2.6 Pcr,y = π²·E·Iy / (KL)y²", S.S_DATA_LABEL)
    put(cells, r+6, 1, "=PI()^2*B19*B10/(B22^2)", S.S_RESULT_VALUE)
    put(cells, r+6, 2, "N", S.S_UNIT)
    put(cells, r+6, 3, "Carga crítica de Euler eje y.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+7, 0, "   2.7 Eje que gobierna el pandeo", S.S_DATA_LABEL)
    put(cells, r+7, 1, '=IF(B23<B24,"Eje x (fuerte)","Eje y (débil)")', S.S_RESULT_VALUE)
    put(cells, r+7, 2, "", S.S_UNIT)
    put(cells, r+7, 3, "El eje con MENOR Pcr gobierna. La columna pandea ahí primero.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+8, 0, "   2.8 Pcr (pandeo) = mín(Pcr,x; Pcr,y)", S.S_DATA_LABEL)
    put(cells, r+8, 1, "=MIN(B23,B24)", S.S_RESULT_VALUE)
    put(cells, r+8, 2, "N", S.S_UNIT)
    put(cells, r+8, 3, "Carga crítica gobernante (la menor de ambos ejes).", S.S_FORMULA_EXPLAIN)


    
    # VERIFICACIÓN
    r = 27
    put(cells, r, 0, "PASO 3: VERIFICACIÓN POR FLUENCIA", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "Comprobar que pandeo ocurre antes que fluencia", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   3.1 σ pandeo = Pcr / A", S.S_DATA_LABEL)
    put(cells, r+1, 1, "=B26/B8", S.S_RESULT_VALUE)
    put(cells, r+1, 2, "Pa", S.S_UNIT)
    put(cells, r+1, 3, "Esfuerzo que produciría el pandeo en la columna.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+2, 0, "   3.2 σ pandeo en MPa", S.S_DATA_LABEL)
    put(cells, r+2, 1, "=B29/1E6", S.S_RESULT_VALUE)
    put(cells, r+2, 2, "MPa", S.S_UNIT)
    put(cells, r+2, 3, "Conversión a MPa para comparar con Fy.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+3, 0, "   3.3 Pcr fluencia = Fy × A", S.S_DATA_LABEL)
    put(cells, r+3, 1, "=B20*B8", S.S_RESULT_VALUE)
    put(cells, r+3, 2, "N", S.S_UNIT)
    put(cells, r+3, 3, "Carga máxima si el material fluye antes de pandear.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+4, 0, "   3.4 ¿σ < Fy? → ¿Rige pandeo?", S.S_DATA_LABEL)
    put(cells, r+4, 1, '=IF(B30<B7,"SI — Rige pandeo (Euler)","NO — Rige fluencia")', S.S_RESULT_VALUE)
    put(cells, r+4, 2, "", S.S_UNIT)
    put(cells, r+4, 3, "Si σ < Fy: la columna pandea antes de fluir (Euler válido).", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+5, 0, "   3.5 Pcr final (gobernante)", S.S_DATA_LABEL)
    put(cells, r+5, 1, '=IF(B30<B7,B26,B31)', S.S_RESULT_VALUE)
    put(cells, r+5, 2, "N", S.S_UNIT)
    put(cells, r+5, 3, "Se usa el MENOR: Pcr_pandeo o Pcr_fluencia.", S.S_FORMULA_EXPLAIN)
    
    # RESULTADO FINAL
    r = 34
    put(cells, r, 0, "PASO 4: CARGA ADMISIBLE (RESULTADO)", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   4.1 P_admisible = Pcr_final / F.S.", S.S_DATA_LABEL)
    put(cells, r+1, 1, "=B33/B15", S.S_RESULT_VALUE)
    put(cells, r+1, 2, "N", S.S_UNIT)
    put(cells, r+1, 3, "Dividimos por factor de seguridad para obtener carga de diseño.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+2, 0, "   4.2 P_admisible en kN", S.S_DATA_LABEL)
    put(cells, r+2, 1, "=B36/1000", S.S_RESULT_VALUE)
    put(cells, r+2, 2, "kN", S.S_UNIT)
    put(cells, r+2, 3, "Conversión a kilonewtons.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+4, 0, "   ★ RESULTADO FINAL: P admisible =", S.S_DATA_LABEL)
    put(cells, r+4, 1, "=B37", S.S_BIG_RESULT)
    put(cells, r+4, 2, "kN", S.S_RESULT_VALUE)
    put(cells, r+4, 3, "Esta es la CARGA MÁXIMA que puede soportar la columna con seguridad.", S.S_EDITABLE_NOTE)
    
    return cells



# ============================================================
# HOJA 3: TABLA MULTI-EJERCICIOS
# ============================================================
def crear_hoja_tabla(wb):
    cells, meta = wb.add_sheet("Tabla Multi-Ejercicios")
    meta['col_widths'] = {0: 5, 1: 35, 2: 10, 3: 10, 4: 12, 5: 14, 6: 14,
                          7: 6, 8: 12, 9: 6, 10: 12, 11: 10, 12: 10,
                          13: 14, 14: 14, 15: 14, 16: 14, 17: 14, 18: 14,
                          19: 16, 20: 6, 21: 14}
    
    put(cells, 0, 0, "TABLA MULTI-EJERCICIOS — Pandeo de Columnas (Automatizada)", S.S_TITLE)
    put(cells, 1, 0, "⚙ INSTRUCCIONES: Llene las columnas A-K (azul claro) y U (F.S.). Las columnas L-V (verde) se calculan solas.", S.S_NOTE)
    put(cells, 2, 0, "Puede resolver hasta 10 ejercicios simultáneamente.", S.S_FORMULA_EXPLAIN)
    
    # Headers
    headers = ["N.°", "Descripción", "E(GPa)", "Fy(MPa)", "A(m²)", "Ix(m⁴)", "Iy(m⁴)",
               "Kx", "Lx(m)", "Ky", "Ly(m)", "(KL)x", "(KL)y",
               "Pcr,x(N)", "Pcr,y(N)", "Eje gov.", "Pcr(kN)", "σ(MPa)", "Pcr_fl(kN)",
               "Modo falla", "F.S.", "P_adm(kN)"]
    
    for col, h in enumerate(headers):
        style = S.S_TABLE_HEADER
        put(cells, 3, col, h, style)
    
    # Data row 1 (example)
    put(cells, 4, 0, 1, S.S_TABLE_DATA)
    put(cells, 4, 1, "Columna A992 arriost. media altura", S.S_TABLE_DATA)
    put(cells, 4, 2, 200, S.S_DATA_INPUT)
    put(cells, 4, 3, 345, S.S_DATA_INPUT)
    put(cells, 4, 4, 0.0074, S.S_DATA_INPUT)
    put(cells, 4, 5, 0.0000873, S.S_DATA_INPUT)
    put(cells, 4, 6, 0.0000188, S.S_DATA_INPUT)
    put(cells, 4, 7, 1, S.S_DATA_INPUT)
    put(cells, 4, 8, 12, S.S_DATA_INPUT)
    put(cells, 4, 9, 1, S.S_DATA_INPUT)
    put(cells, 4, 10, 6, S.S_DATA_INPUT)
    put(cells, 4, 20, 2.0, S.S_DATA_INPUT)
    
    # Formulas for 10 rows
    for i in range(10):
        row = 4 + i
        rn = row + 1
        if i > 0:
            put(cells, row, 0, i + 1, S.S_TABLE_DATA)
            put(cells, row, 1, "", S.S_TABLE_DATA)
            for c in range(2, 11):
                put(cells, row, c, "", S.S_DATA_INPUT)
            put(cells, row, 20, 2.0, S.S_DATA_INPUT)
        
        # Calculated columns
        put(cells, row, 11, f'=IF(H{rn}="","",H{rn}*I{rn})', S.S_RESULT_VALUE)
        put(cells, row, 12, f'=IF(J{rn}="","",J{rn}*K{rn})', S.S_RESULT_VALUE)
        put(cells, row, 13, f'=IF(L{rn}="","",PI()^2*(C{rn}*1E9)*F{rn}/L{rn}^2)', S.S_RESULT_VALUE)
        put(cells, row, 14, f'=IF(M{rn}="","",PI()^2*(C{rn}*1E9)*G{rn}/M{rn}^2)', S.S_RESULT_VALUE)
        put(cells, row, 15, f'=IF(N{rn}="","",IF(N{rn}<O{rn},"Eje x","Eje y"))', S.S_RESULT_VALUE)
        put(cells, row, 16, f'=IF(N{rn}="","",MIN(N{rn},O{rn})/1000)', S.S_RESULT_VALUE)
        put(cells, row, 17, f'=IF(Q{rn}="","",Q{rn}*1000/E{rn}/1E6)', S.S_RESULT_VALUE)
        put(cells, row, 18, f'=IF(E{rn}="","",(D{rn}*1E6)*E{rn}/1000)', S.S_RESULT_VALUE)
        put(cells, row, 19, f'=IF(Q{rn}="","",IF(Q{rn}<S{rn},"Pandeo","Fluencia"))', S.S_RESULT_VALUE)
        put(cells, row, 21, f'=IF(Q{rn}="","",MIN(Q{rn},S{rn})/U{rn})', S.S_RESULT_VALUE)
    
    # Legend
    r = 16
    put(cells, r, 0, "LEYENDA DE COLORES:", S.S_STEP_NUM)
    put(cells, r+1, 0, "   Amarillo = Dato editable (usted lo ingresa)", S.S_EDITABLE_NOTE)
    put(cells, r+2, 0, "   Verde = Resultado automático (NO tocar)", S.S_RESULT_VALUE)
    
    return cells



# ============================================================
# HOJA 4: COLUMNA CIRCULAR (ACI / NSR-84)
# ============================================================
def crear_hoja_circular(wb):
    cells, meta = wb.add_sheet("Columna Circular")
    meta['col_widths'] = {0: 38, 1: 16, 2: 8, 3: 16, 4: 8, 5: 50}
    
    put(cells, 0, 0, "COLUMNA CIRCULAR — Normas ACI y NSR-84", S.S_TITLE)
    put(cells, 1, 0, "Tubo de acero empotrado en la base. Análisis para 3 longitudes.", S.S_SUBTITLE)
    put(cells, 2, 0, "⚙ Cambie las celdas AMARILLAS y todo se recalcula automáticamente.", S.S_NOTE)
    
    # DATOS
    r = 4
    put(cells, r, 0, "PASO 1: DATOS DEL MATERIAL", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "", S.S_SECTION_HEADER)
    put(cells, r, 4, "", S.S_SECTION_HEADER)
    put(cells, r, 5, "Explicación", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   1.1 E (módulo elasticidad)", S.S_DATA_LABEL)
    put(cells, r+1, 1, 200, S.S_DATA_INPUT)
    put(cells, r+1, 2, "GPa", S.S_UNIT)
    put(cells, r+1, 3, "=B6*1E9", S.S_RESULT_VALUE)
    put(cells, r+1, 4, "Pa", S.S_UNIT)
    put(cells, r+1, 5, "Rigidez del acero.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+2, 0, "   1.2 Fy (esfuerzo fluencia)", S.S_DATA_LABEL)
    put(cells, r+2, 1, 250, S.S_DATA_INPUT)
    put(cells, r+2, 2, "MPa", S.S_UNIT)
    put(cells, r+2, 3, "=B7*1E6", S.S_RESULT_VALUE)
    put(cells, r+2, 4, "Pa", S.S_UNIT)
    put(cells, r+2, 5, "Límite elástico del acero.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+3, 0, "   1.3 k (factor longitud efectiva)", S.S_DATA_LABEL)
    put(cells, r+3, 1, 2, S.S_DATA_INPUT)
    put(cells, r+3, 2, "", S.S_UNIT)
    put(cells, r+3, 5, "k=2 para empotrado-libre. k=1 art-art. k=0.5 emp-emp.", S.S_FORMULA_EXPLAIN)
    
    # Longitudes
    r = 9
    put(cells, r, 0, "PASO 2: LONGITUDES", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 5, "", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   2.1 L1", S.S_DATA_LABEL)
    put(cells, r+1, 1, 2, S.S_DATA_INPUT)
    put(cells, r+1, 2, "m", S.S_UNIT)
    put(cells, r+2, 0, "   2.2 L2", S.S_DATA_LABEL)
    put(cells, r+2, 1, 2.5, S.S_DATA_INPUT)
    put(cells, r+2, 2, "m", S.S_UNIT)
    put(cells, r+3, 0, "   2.3 L3", S.S_DATA_LABEL)
    put(cells, r+3, 1, 3.5, S.S_DATA_INPUT)
    put(cells, r+3, 2, "m", S.S_UNIT)
    
    # Dimensiones
    r = 14
    put(cells, r, 0, "PASO 3: DIMENSIONES DEL TUBO", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 5, "", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   3.1 D.int (diámetro interior)", S.S_DATA_LABEL)
    put(cells, r+1, 1, 13, S.S_DATA_INPUT)
    put(cells, r+1, 2, "cm", S.S_UNIT)
    put(cells, r+1, 3, "=B16/100", S.S_RESULT_VALUE)
    put(cells, r+1, 4, "m", S.S_UNIT)
    
    put(cells, r+2, 0, "   3.2 D.ext (diámetro exterior)", S.S_DATA_LABEL)
    put(cells, r+2, 1, 15, S.S_DATA_INPUT)
    put(cells, r+2, 2, "cm", S.S_UNIT)
    put(cells, r+2, 3, "=B17/100", S.S_RESULT_VALUE)
    put(cells, r+2, 4, "m", S.S_UNIT)
    
    # Propiedades calculadas
    r = 18
    put(cells, r, 0, "PASO 4: PROPIEDADES CALCULADAS", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 5, "Derivación", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   4.1 I = π(D_ext⁴-D_int⁴)/64", S.S_DATA_LABEL)
    put(cells, r+1, 1, "=PI()*(D17^4-D16^4)/64", S.S_RESULT_VALUE)
    put(cells, r+1, 2, "m⁴", S.S_UNIT)
    put(cells, r+1, 5, "Inercia de tubo circular. Se usa D(m) = D(cm)/100.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+2, 0, "   4.2 A = π(D_ext²-D_int²)/4", S.S_DATA_LABEL)
    put(cells, r+2, 1, "=PI()*(D17^2-D16^2)/4", S.S_RESULT_VALUE)
    put(cells, r+2, 2, "m²", S.S_UNIT)
    put(cells, r+2, 5, "Área del anillo = Área ext - Área int.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+3, 0, "   4.3 r = √(I/A) (radio de giro)", S.S_DATA_LABEL)
    put(cells, r+3, 1, "=SQRT(B20/B21)", S.S_RESULT_VALUE)
    put(cells, r+3, 2, "m", S.S_UNIT)
    put(cells, r+3, 5, "Radio de giro. Determina la esbeltez kL/r.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+4, 0, "   4.4 Cc = √(2π²E/Fy)", S.S_DATA_LABEL)
    put(cells, r+4, 1, "=SQRT(2*PI()^2*D6/D7)", S.S_RESULT_VALUE)
    put(cells, r+4, 2, "", S.S_UNIT)
    put(cells, r+4, 5, "Esbeltez crítica. Si kL/r < Cc: col. corta. Si kL/r > Cc: col. larga.", S.S_FORMULA_EXPLAIN)


    
    # NORMA ACI
    r = 24
    put(cells, r, 0, "PASO 5: NORMA ACI (Pandeo de Euler)", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "L2", S.S_SECTION_HEADER)
    put(cells, r, 4, "", S.S_SECTION_HEADER)
    put(cells, r, 5, "L3", S.S_SECTION_HEADER)
    
    # L1
    put(cells, r+1, 0, "   5.1 Le = k·L", S.S_DATA_LABEL)
    put(cells, r+1, 1, "=B8*B11", S.S_RESULT_VALUE)
    put(cells, r+1, 2, "m", S.S_UNIT)
    put(cells, r+1, 3, "=B8*B12", S.S_RESULT_VALUE)
    put(cells, r+1, 4, "m", S.S_UNIT)
    put(cells, r+1, 5, "=B8*B13", S.S_RESULT_VALUE)
    
    put(cells, r+2, 0, "   5.2 kL/r (esbeltez)", S.S_DATA_LABEL)
    put(cells, r+2, 1, "=B26/B22", S.S_RESULT_VALUE)
    put(cells, r+2, 2, "", S.S_UNIT)
    put(cells, r+2, 3, "=D26/B22", S.S_RESULT_VALUE)
    put(cells, r+2, 5, "=F26/B22", S.S_RESULT_VALUE)
    
    put(cells, r+3, 0, "   5.3 Pcr = π²EI/(kL)²", S.S_DATA_LABEL)
    put(cells, r+3, 1, "=PI()^2*D6*B20/B26^2", S.S_RESULT_VALUE)
    put(cells, r+3, 2, "N", S.S_UNIT)
    put(cells, r+3, 3, "=PI()^2*D6*B20/D26^2", S.S_RESULT_VALUE)
    put(cells, r+3, 5, "=PI()^2*D6*B20/F26^2", S.S_RESULT_VALUE)
    
    put(cells, r+4, 0, "   5.4 Pcr (kN)", S.S_DATA_LABEL)
    put(cells, r+4, 1, "=B28/1000", S.S_BIG_RESULT)
    put(cells, r+4, 2, "kN", S.S_UNIT)
    put(cells, r+4, 3, "=D28/1000", S.S_BIG_RESULT)
    put(cells, r+4, 5, "=F28/1000", S.S_BIG_RESULT)
    
    put(cells, r+5, 0, "   5.5 σcr = Pcr/A", S.S_DATA_LABEL)
    put(cells, r+5, 1, "=B28/B21/1E6", S.S_RESULT_VALUE)
    put(cells, r+5, 2, "MPa", S.S_UNIT)
    put(cells, r+5, 3, "=D28/B21/1E6", S.S_RESULT_VALUE)
    put(cells, r+5, 5, "=F28/B21/1E6", S.S_RESULT_VALUE)
    
    # NORMA NSR-84
    r = 31
    put(cells, r, 0, "PASO 6: NORMA NSR-84", S.S_SECTION_HEADER)
    put(cells, r, 1, "L1", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "L2", S.S_SECTION_HEADER)
    put(cells, r, 4, "", S.S_SECTION_HEADER)
    put(cells, r, 5, "L3", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   6.1 Esbeltez kL/r", S.S_DATA_LABEL)
    put(cells, r+1, 1, "=B27", S.S_RESULT_VALUE)
    put(cells, r+1, 3, "=D27", S.S_RESULT_VALUE)
    put(cells, r+1, 5, "=F27", S.S_RESULT_VALUE)
    
    put(cells, r+2, 0, "   6.2 Tipo columna", S.S_DATA_LABEL)
    put(cells, r+2, 1, '=IF(B33<B23,"Corta","Larga")', S.S_RESULT_VALUE)
    put(cells, r+2, 3, '=IF(D33<B23,"Corta","Larga")', S.S_RESULT_VALUE)
    put(cells, r+2, 5, '=IF(F33<B23,"Corta","Larga")', S.S_RESULT_VALUE)
    
    # NSR formulas
    # Corta: σa = Fy[1-(kL/r)²/(2Cc²)] / [5/3 + 3(kL/r)/(8Cc) - (kL/r)³/(8Cc³)]
    # Larga: σa = 12π²E/(23(kL/r)²)
    for col, esb in [(1, "B33"), (3, "D33"), (5, "F33")]:
        num = f"(D7*(1-{esb}^2/(2*B23^2)))"
        den = f"(5/3+3*{esb}/(8*B23)-{esb}^3/(8*B23^3))"
        larga = f"(12*PI()^2*D6/(23*{esb}^2))"
        put(cells, r+4, col, f"=IF({esb}<B23,{num}/{den},{larga})/1E6", S.S_RESULT_VALUE)
    
    put(cells, r+4, 0, "   6.3 σa (MPa)", S.S_DATA_LABEL)
    
    put(cells, r+5, 0, "   6.4 P admisible (kN)", S.S_DATA_LABEL)
    put(cells, r+5, 1, "=B36*1E6*B21/1000", S.S_BIG_RESULT)
    put(cells, r+5, 3, "=D36*1E6*B21/1000", S.S_BIG_RESULT)
    put(cells, r+5, 5, "=F36*1E6*B21/1000", S.S_BIG_RESULT)
    
    return cells



# ============================================================
# HOJA 5: COLUMNAS EN I
# ============================================================
def crear_hoja_columnas_I(wb):
    cells, meta = wb.add_sheet("Columnas en I")
    meta['col_widths'] = {0: 40, 1: 16, 2: 10, 3: 50}
    
    put(cells, 0, 0, "COLUMNAS EN I — Análisis con Diferentes Condiciones de Apoyo", S.S_TITLE)
    put(cells, 1, 0, "Perfil W10x60. Análisis con 4 valores de K.", S.S_SUBTITLE)
    put(cells, 2, 0, "⚙ Modifique datos del material, longitud y perfil en las celdas AMARILLAS.", S.S_NOTE)
    
    r = 4
    put(cells, r, 0, "PASO 1: DATOS DEL MATERIAL", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "Explicación", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   1.1 E (módulo elasticidad)", S.S_DATA_LABEL)
    put(cells, r+1, 1, 30000000, S.S_DATA_INPUT)
    put(cells, r+1, 2, "PSI", S.S_UNIT)
    put(cells, r+1, 3, "Módulo de elasticidad del acero en sistema imperial.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+2, 0, "   1.2 Fy (esfuerzo fluencia)", S.S_DATA_LABEL)
    put(cells, r+2, 1, 20000, S.S_DATA_INPUT)
    put(cells, r+2, 2, "PSI", S.S_UNIT)
    put(cells, r+2, 3, "Esfuerzo de fluencia. A36 steel = 36000 PSI.", S.S_FORMULA_EXPLAIN)
    
    r = 8
    put(cells, r, 0, "PASO 2: FACTORES K", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "Condición de extremos", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   2.1 k1 (Articulado-Articulado)", S.S_DATA_LABEL)
    put(cells, r+1, 1, 1, S.S_DATA_INPUT)
    put(cells, r+1, 3, "Ambos extremos con rotación libre.", S.S_FORMULA_EXPLAIN)
    put(cells, r+2, 0, "   2.2 k2 (Empotrado-Articulado)", S.S_DATA_LABEL)
    put(cells, r+2, 1, 0.7, S.S_DATA_INPUT)
    put(cells, r+2, 3, "Un extremo fijo, otro con rotación libre.", S.S_FORMULA_EXPLAIN)
    put(cells, r+3, 0, "   2.3 k3 (Empotrado-Empotrado)", S.S_DATA_LABEL)
    put(cells, r+3, 1, 0.5, S.S_DATA_INPUT)
    put(cells, r+3, 3, "Ambos extremos fijos. Máxima restricción.", S.S_FORMULA_EXPLAIN)
    put(cells, r+4, 0, "   2.4 k4 (Empotrado-Libre)", S.S_DATA_LABEL)
    put(cells, r+4, 1, 2, S.S_DATA_INPUT)
    put(cells, r+4, 3, "Un extremo fijo, otro libre (voladizo). Peor caso.", S.S_FORMULA_EXPLAIN)
    
    r = 14
    put(cells, r, 0, "PASO 3: LONGITUD Y PERFIL", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "", S.S_SECTION_HEADER)
    
    put(cells, r+1, 0, "   3.1 L (longitud)", S.S_DATA_LABEL)
    put(cells, r+1, 1, 2.5, S.S_DATA_INPUT)
    put(cells, r+1, 2, "m", S.S_UNIT)
    put(cells, r+1, 3, "Longitud de la columna en metros.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+2, 0, "   3.2 L en pulgadas", S.S_DATA_LABEL)
    put(cells, r+2, 1, "=B16*39.3701", S.S_RESULT_VALUE)
    put(cells, r+2, 2, "pulg", S.S_UNIT)
    put(cells, r+2, 3, "Conversión: 1 m = 39.3701 pulgadas.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+3, 0, "   3.3 A (área perfil)", S.S_DATA_LABEL)
    put(cells, r+3, 1, 17.6, S.S_DATA_INPUT)
    put(cells, r+3, 2, "pulg²", S.S_UNIT)
    put(cells, r+3, 3, "Área de la sección W10x60 (tabla AISC).", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+4, 0, "   3.4 Iy (inercia eje débil)", S.S_DATA_LABEL)
    put(cells, r+4, 1, 116, S.S_DATA_INPUT)
    put(cells, r+4, 2, "pulg⁴", S.S_UNIT)
    put(cells, r+4, 3, "Inercia eje y del W10x60. Eje débil gobierna pandeo.", S.S_FORMULA_EXPLAIN)
    
    put(cells, r+5, 0, "   3.5 ry (radio giro eje y)", S.S_DATA_LABEL)
    put(cells, r+5, 1, 2.57, S.S_DATA_INPUT)
    put(cells, r+5, 2, "pulg", S.S_UNIT)
    put(cells, r+5, 3, "Radio de giro eje débil. ry = √(Iy/A).", S.S_FORMULA_EXPLAIN)
    
    # RESULTADOS
    r = 21
    put(cells, r, 0, "PASO 4: RESULTADOS POR CONDICIÓN DE APOYO", S.S_SECTION_HEADER)
    put(cells, r, 1, "", S.S_SECTION_HEADER)
    put(cells, r, 2, "", S.S_SECTION_HEADER)
    put(cells, r, 3, "", S.S_SECTION_HEADER)
    
    # Table headers
    headers2 = ["Condición", "k", "Le(pulg)", "kL/ry", "Pcr(Lb)", "σcr(PSI)", "Cc", "Tipo", "σa NSR(PSI)", "P_adm(Lb)"]
    for c, h in enumerate(headers2):
        put(cells, r+1, c, h, S.S_TABLE_HEADER)
    
    k_names = ["Art-Art", "Emp-Art", "Emp-Emp", "Emp-Libre"]
    k_refs = ["B10", "B11", "B12", "B13"]
    
    for i, (kn, kr) in enumerate(zip(k_names, k_refs)):
        row = r + 2 + i
        rn = row + 1
        put(cells, row, 0, f"   4.{i+1} {kn}", S.S_DATA_LABEL)
        put(cells, row, 1, f"={kr}", S.S_RESULT_VALUE)
        put(cells, row, 2, f"={kr}*B17", S.S_RESULT_VALUE)  # Le
        put(cells, row, 3, f"=C{rn}/B20", S.S_RESULT_VALUE)  # kL/ry
        put(cells, row, 4, f"=PI()^2*B6*B19/C{rn}^2", S.S_RESULT_VALUE)  # Pcr
        put(cells, row, 5, f"=E{rn}/B18", S.S_RESULT_VALUE)  # σcr
        put(cells, row, 6, "=SQRT(2*PI()^2*B6/B7)", S.S_RESULT_VALUE)  # Cc
        put(cells, row, 7, f'=IF(D{rn}<G{rn},"Corta","Larga")', S.S_RESULT_VALUE)
        # σa NSR
        num = f"(B7*(1-D{rn}^2/(2*G{rn}^2)))"
        den = f"(5/3+3*D{rn}/(8*G{rn})-D{rn}^3/(8*G{rn}^3))"
        larga = f"(12*PI()^2*B6/(23*D{rn}^2))"
        put(cells, row, 8, f"=IF(D{rn}<G{rn},{num}/{den},{larga})", S.S_RESULT_VALUE)
        put(cells, row, 9, f"=I{rn}*B18", S.S_BIG_RESULT)  # P_adm
    
    return cells



# ============================================================
# MAIN
# ============================================================
def main():
    wb = StyledXlsx()
    
    crear_hoja_propiedades(wb)
    crear_hoja_carga_admisible(wb)
    crear_hoja_tabla(wb)
    crear_hoja_circular(wb)
    crear_hoja_columnas_I(wb)
    
    output = "Excel_Automatizado_RM_PRO.xlsx"
    wb.save(output)
    
    size = os.path.getsize(output)
    print(f"{'='*60}")
    print(f"  EXCEL PROFESIONAL GENERADO EXITOSAMENTE")
    print(f"{'='*60}")
    print(f"  Archivo: {output}")
    print(f"  Tamano:  {size:,} bytes")
    print(f"  Hojas:   {len(wb.sheets)}")
    print(f"  Imagenes: {len(wb.images_data)} diagramas de seccion")
    print(f"{'='*60}")
    print(f"  CONTENIDO:")
    print(f"  1. Prop. Seccion    - 6 tipos con formulas explicadas paso a paso")
    print(f"  2. Carga admisible  - Pandeo Euler con verificacion de fluencia")
    print(f"  3. Tabla Multi-Ej.  - 10 ejercicios simultaneos")
    print(f"  4. Columna Circular - ACI + NSR-84 para 3 longitudes")
    print(f"  5. Columnas en I    - W10x60 con 4 condiciones de apoyo")
    print(f"{'='*60}")
    print(f"  DISENO:")
    print(f"  - Colores profesionales (azul/verde/amarillo)")
    print(f"  - Celdas AMARILLAS = datos editables")
    print(f"  - Celdas VERDES = resultados automaticos")
    print(f"  - Pasos numerados con explicaciones")
    print(f"  - Imagenes de secciones transversales")
    print(f"  - Bordes y fuentes elegantes")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
