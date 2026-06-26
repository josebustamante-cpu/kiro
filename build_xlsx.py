# -*- coding: utf-8 -*-
"""
Genera un Excel (.xlsx) con FORMULAS VIVAS para evaluar Flota propia vs Courier.
Solo libreria estandar (zipfile + XML). Al editar las celdas de la hoja
'Supuestos' todo se recalcula. Incluye sobrekilometraje y vista a 3 anos con
alza de combustible.
"""
import zipfile

# ----------------------------------------------------------------------
# Supuestos base (valores iniciales; editables en el Excel)
# ----------------------------------------------------------------------
S = dict(
    envios=15000, dias=22, cap=120, courier_unit=2980, UF=40000, Nbase=7,
    kmtot=30000, kmincl=3000, sobrekm=150,
    leas_e=34, leas_p=27, kwh=150, cons=0.25, diesel=1050, rend=10, infra=25000,
    cond=1200000, peon_pct=0.30, multas=500000, overfijo=2770000, telem=13000,
    alza_diesel=0.08, alza_elec=0.03, reaj_courier=0.05, reaj_uf=0.035,
    reaj_sueldo=0.04, margen=0.30,
)


def model(N):
    leas_e = N * S['leas_e'] * S['UF']
    leas_p = N * S['leas_p'] * S['UF']
    energ_e = S['kmtot'] * S['cons'] * S['kwh']
    energ_d = S['kmtot'] / S['rend'] * S['diesel']
    infra = N * S['infra']
    personal = N * S['cond'] * (2 - S['peon_pct'])
    overhead = S['overfijo'] + N * S['telem']
    sobre = max(0, S['kmtot'] - N * S['kmincl']) * S['sobrekm']
    tot_e = leas_e + energ_e + infra + personal + overhead + sobre + S['multas']
    tot_p = leas_p + energ_d + personal + overhead + sobre + S['multas']
    return dict(leas_e=leas_e, leas_p=leas_p, energ_e=energ_e, energ_d=energ_d,
                infra=infra, personal=personal, overhead=overhead, sobre=sobre,
                tot_e=tot_e, tot_p=tot_p)

# ----------------------------------------------------------------------
# Infraestructura XLSX
# ----------------------------------------------------------------------
def col_letter(c):
    s = ''
    while c > 0:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def esc(t):
    return (str(t).replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;'))


class Sheet:
    def __init__(self, name):
        self.name = name
        self.cells = {}      # (r,c) -> xml
        self.widths = []     # (c, width)

    def text(self, r, c, val, style=0):
        self.cells[(r, c)] = (f'<c r="{col_letter(c)}{r}" s="{style}" '
                              f't="inlineStr"><is><t xml:space="preserve">'
                              f'{esc(val)}</t></is></c>')

    def num(self, r, c, val, style=0):
        self.cells[(r, c)] = (f'<c r="{col_letter(c)}{r}" s="{style}">'
                              f'<v>{val}</v></c>')

    def formula(self, r, c, f, val, style=0):
        self.cells[(r, c)] = (f'<c r="{col_letter(c)}{r}" s="{style}">'
                              f'<f>{esc(f)}</f><v>{val}</v></c>')

    def xml(self):
        rows = {}
        for (r, c), x in self.cells.items():
            rows.setdefault(r, []).append((c, x))
        body = ''
        for r in sorted(rows):
            cells = ''.join(x for _, x in sorted(rows[r]))
            body += f'<row r="{r}">{cells}</row>'
        cols = ''
        if self.widths:
            cols = '<cols>' + ''.join(
                f'<col min="{c}" max="{c}" width="{w}" customWidth="1"/>'
                for c, w in self.widths) + '</cols>'
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<worksheet xmlns="http://schemas.openxmlformats.org/'
                'spreadsheetml/2006/main">'
                f'{cols}<sheetData>{body}</sheetData></worksheet>')


# estilos: 0 def | 1 titulo | 2 seccion | 3 in-num | 4 in-money | 5 in-pct |
# 6 in-dec | 7 header | 8 money | 9 money-bold | 10 pct | 11 num | 12 bold |
# 13 bordered-text
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="4">
<numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0"/>
<numFmt numFmtId="165" formatCode="0.0%"/>
<numFmt numFmtId="166" formatCode="#,##0"/>
<numFmt numFmtId="167" formatCode="0.00"/>
</numFmts>
<fonts count="4">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="14"/><color rgb="FF123A5E"/><name val="Calibri"/></font>
</fonts>
<fills count="5">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF123A5E"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF7CC"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFECECEC"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="15">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="1" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="166" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="164" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="167" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="164" fontId="1" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''


# ----------------------------------------------------------------------
# HOJA 1: Supuestos
# ----------------------------------------------------------------------
def build_supuestos():
    sh = Sheet('Supuestos')
    sh.widths = [(1, 42), (2, 16), (3, 34)]
    sh.text(1, 1, 'MODELO: FLOTA PROPIA vs COURIER  -  CAV', 1)
    sh.text(2, 1, 'Edita las celdas amarillas; las demas hojas se recalculan '
                  'solas.', 12)

    def inp(r, label, key, style, unit=''):
        sh.text(r, 1, label)
        if style == 5 or style == 6:
            sh.num(r, 2, S[key], style)
        else:
            sh.num(r, 2, S[key], style)
        if unit:
            sh.text(r, 3, unit)

    sh.text(4, 1, 'PARAMETROS GENERALES', 2)
    sh.text(4, 2, '', 2); sh.text(4, 3, '', 2)
    inp(5, 'Envios por mes', 'envios', 3)
    inp(6, 'Dias habiles por mes', 'dias', 3)
    inp(7, 'Capacidad entregas por movil/dia', 'cap', 3)
    inp(8, 'Costo courier por envio (CLP)', 'courier_unit', 4)
    inp(9, 'Valor UF (CLP)', 'UF', 4)
    inp(10, 'N de moviles (escenario base)', 'Nbase', 3)

    sh.text(12, 1, 'REPARTO Y KILOMETRAJE', 2)
    sh.text(12, 2, '', 2); sh.text(12, 3, '', 2)
    inp(13, 'Km totales de reparto por mes', 'kmtot', 3)
    inp(14, 'Km incluidos por movil (contrato)', 'kmincl', 3)
    inp(15, 'Costo sobrekilometraje (CLP/km)', 'sobrekm', 4)

    sh.text(17, 1, 'VEHICULOS Y ENERGIA', 2)
    sh.text(17, 2, '', 2); sh.text(17, 3, '', 2)
    inp(18, 'Leasing electrico (UF por movil/mes)', 'leas_e', 3)
    inp(19, 'Leasing petrolero (UF por movil/mes)', 'leas_p', 3)
    inp(20, 'Precio electricidad (CLP/kWh)', 'kwh', 4)
    inp(21, 'Consumo electrico (kWh/km)', 'cons', 6)
    inp(22, 'Precio diesel (CLP/L)', 'diesel', 4)
    inp(23, 'Rendimiento diesel (km/L)', 'rend', 3)
    inp(24, 'Infra de carga electrica (CLP/movil/mes)', 'infra', 4)

    sh.text(26, 1, 'PERSONAL Y OVERHEAD', 2)
    sh.text(26, 2, '', 2); sh.text(26, 3, '', 2)
    inp(27, 'Costo empresa conductor (CLP/mes)', 'cond', 4)
    inp(28, 'Peoneta: % bajo el conductor', 'peon_pct', 5)
    inp(29, 'Provision multas/deducibles (CLP/mes)', 'multas', 4)
    inp(30, 'Overhead fijo (CLP/mes)', 'overfijo', 4)
    inp(31, 'Telemetria (CLP/movil/mes)', 'telem', 4)

    sh.text(33, 1, 'PROYECCION 3 ANOS (reajustes anuales)', 2)
    sh.text(33, 2, '', 2); sh.text(33, 3, '', 2)
    inp(34, 'Alza anual diesel', 'alza_diesel', 5)
    inp(35, 'Alza anual electricidad', 'alza_elec', 5)
    inp(36, 'Reajuste anual courier', 'reaj_courier', 5)
    inp(37, 'Reajuste anual UF / leasing', 'reaj_uf', 5)
    inp(38, 'Reajuste anual sueldos / overhead', 'reaj_sueldo', 5)

    sh.text(40, 1, 'TARIFA / MARGEN (operador logistico)', 2)
    sh.text(40, 2, '', 2); sh.text(40, 3, '', 2)
    inp(41, 'Margen objetivo (% sobre venta)', 'margen', 5)
    return sh


# Direcciones absolutas de Supuestos (para formulas)
A = dict(envios='Supuestos!$B$5', dias='Supuestos!$B$6', cap='Supuestos!$B$7',
         courier='Supuestos!$B$8', UF='Supuestos!$B$9', Nbase='Supuestos!$B$10',
         kmtot='Supuestos!$B$13', kmincl='Supuestos!$B$14',
         sobrekm='Supuestos!$B$15', leas_e='Supuestos!$B$18',
         leas_p='Supuestos!$B$19', kwh='Supuestos!$B$20', cons='Supuestos!$B$21',
         diesel='Supuestos!$B$22', rend='Supuestos!$B$23', infra='Supuestos!$B$24',
         cond='Supuestos!$B$27', peon='Supuestos!$B$28', multas='Supuestos!$B$29',
         overfijo='Supuestos!$B$30', telem='Supuestos!$B$31',
         alza_d='Supuestos!$B$34', alza_e='Supuestos!$B$35',
         reaj_c='Supuestos!$B$36', reaj_uf='Supuestos!$B$37',
         reaj_s='Supuestos!$B$38', margen='Supuestos!$B$41')


def f_personal(N):
    return f'({N}*{A["cond"]}*(2-{A["peon"]}))'

def f_overhead(N):
    return f'({A["overfijo"]}+{N}*{A["telem"]})'

def f_sobre(N):
    return f'(MAX(0,{A["kmtot"]}-{N}*{A["kmincl"]})*{A["sobrekm"]})'

F_ENERG_E = f'({A["kmtot"]}*{A["cons"]}*{A["kwh"]})'
F_ENERG_D = f'({A["kmtot"]}/{A["rend"]}*{A["diesel"]})'

def f_leas_e(N):
    return f'({N}*{A["leas_e"]}*{A["UF"]})'

def f_leas_p(N):
    return f'({N}*{A["leas_p"]}*{A["UF"]})'

def f_infra(N):
    return f'({N}*{A["infra"]})'

def f_tot_e(N):
    return (f'{f_leas_e(N)}+{F_ENERG_E}+{f_infra(N)}+{f_personal(N)}+'
            f'{f_overhead(N)}+{f_sobre(N)}+{A["multas"]}')

def f_tot_p(N):
    return (f'{f_leas_p(N)}+{F_ENERG_D}+{f_personal(N)}+{f_overhead(N)}+'
            f'{f_sobre(N)}+{A["multas"]}')


# ----------------------------------------------------------------------
# HOJA 2: Escenarios por tamano de flota
# ----------------------------------------------------------------------
def build_escenarios():
    sh = Sheet('Escenarios')
    sh.widths = [(1, 11)] + [(c, 14) for c in range(2, 12)]
    sh.text(1, 1, 'ESCENARIOS POR TAMANO DE FLOTA (con sobrekilometraje)', 1)
    heads = ['Moviles (N)', 'Utilizacion', 'Km/movil', 'Sobrekm/mes',
             'Total Electrica/mes', 'Total Petrolera/mes', '$/envio Elec.',
             '$/envio Pet.', 'Ahorro/ano Elec.', 'Ahorro/ano Pet.']
    for c, h in enumerate(heads, 1):
        sh.text(3, c, h, 7)
    r = 4
    for N in [10, 9, 8, 7, 6]:
        m = model(N)
        Ncell = f'$A{r}'
        sh.num(r, 1, N, 11)
        sh.formula(r, 2, f'{A["envios"]}/({A["cap"]}*{A["dias"]}*{Ncell})',
                   S['envios'] / (S['cap'] * S['dias'] * N), 10)
        sh.formula(r, 3, f'{A["kmtot"]}/{Ncell}', S['kmtot'] / N, 11)
        sh.formula(r, 4, f_sobre(Ncell), m['sobre'], 8)
        sh.formula(r, 5, f_tot_e(Ncell), m['tot_e'], 8)
        sh.formula(r, 6, f_tot_p(Ncell), m['tot_p'], 8)
        sh.formula(r, 7, f'E{r}/{A["envios"]}', m['tot_e'] / S['envios'], 8)
        sh.formula(r, 8, f'F{r}/{A["envios"]}', m['tot_p'] / S['envios'], 8)
        courier = S['courier_unit'] * S['envios']
        sh.formula(r, 9, f'({A["courier"]}*{A["envios"]}-E{r})*12',
                   (courier - m['tot_e']) * 12, 8)
        sh.formula(r, 10, f'({A["courier"]}*{A["envios"]}-F{r})*12',
                   (courier - m['tot_p']) * 12, 8)
        r += 1
    sh.text(r + 1, 1, 'Courier actual (referencia):', 12)
    sh.formula(r + 1, 5, f'{A["courier"]}*{A["envios"]}',
               S['courier_unit'] * S['envios'], 9)
    sh.formula(r + 1, 7, A["courier"], S['courier_unit'], 8)
    return sh


# ----------------------------------------------------------------------
# HOJA 3: Electrica (N base) vs Courier
# ----------------------------------------------------------------------
def build_focus():
    sh = Sheet('Electrica vs Courier')
    sh.widths = [(1, 34), (2, 16), (3, 16), (4, 16)]
    N = S['Nbase']
    m = model(N)
    courier = S['courier_unit'] * S['envios']
    Nc = A['Nbase']

    sh.text(1, 1, 'EVALUACION: FLOTA ELECTRICA vs COURIER', 1)
    sh.text(2, 1, 'Moviles del escenario base (editar en Supuestos!B10):', 12)
    sh.formula(2, 2, Nc, N, 11)

    for c, h in enumerate(['Concepto', 'Courier', 'Flota Electrica',
                           'Diferencia'], 1):
        sh.text(4, c, h, 7)

    # fila costo total/mes
    sh.text(5, 1, 'Costo total / mes', 13)
    sh.formula(5, 2, f'{A["courier"]}*{A["envios"]}', courier, 8)
    sh.formula(5, 3, f_tot_e(Nc), m['tot_e'], 8)
    sh.formula(5, 4, 'C5-B5', m['tot_e'] - courier, 9)
    # costo por envio
    sh.text(6, 1, 'Costo por envio', 13)
    sh.formula(6, 2, A['courier'], S['courier_unit'], 8)
    sh.formula(6, 3, f'C5/{A["envios"]}', m['tot_e'] / S['envios'], 8)
    sh.formula(6, 4, 'C6-B6', m['tot_e'] / S['envios'] - S['courier_unit'], 9)
    # costo total/ano
    sh.text(7, 1, 'Costo total / ano', 13)
    sh.formula(7, 2, 'B5*12', courier * 12, 8)
    sh.formula(7, 3, 'C5*12', m['tot_e'] * 12, 8)
    sh.formula(7, 4, 'C7-B7', (m['tot_e'] - courier) * 12, 9)
    # reduccion %
    sh.text(8, 1, 'Reduccion de costo', 13)
    sh.formula(8, 4, '(B5-C5)/B5', (courier - m['tot_e']) / courier, 10)

    sh.text(10, 1, 'DESGLOSE FLOTA ELECTRICA (CLP/mes)', 2)
    desg = [
        ('Leasing electrico', f_leas_e(Nc), m['leas_e']),
        ('Energia electrica', F_ENERG_E, m['energ_e']),
        ('Infraestructura de carga', f_infra(Nc), m['infra']),
        ('Conductores + peonetas', f_personal(Nc), m['personal']),
        ('Sobrekilometraje', f_sobre(Nc), m['sobre']),
        ('Multas / deducibles', A['multas'], S['multas']),
        ('Overhead de flota', f_overhead(Nc), m['overhead']),
    ]
    r = 11
    for label, fml, val in desg:
        sh.text(r, 1, label, 13)
        sh.formula(r, 2, fml, val, 8)
        r += 1
    sh.text(r, 1, 'TOTAL', 12)
    sh.formula(r, 2, f'SUM(B11:B{r-1})', m['tot_e'], 9)

    # metricas
    mr = r + 2
    sh.text(mr, 1, 'METRICAS DE DECISION', 2)
    cap_mes = S['cap'] * S['dias'] * N
    metrics = [
        ('Ahorro / mes', f'B5-C5', courier - m['tot_e'], 8),
        ('Ahorro / ano', f'(B5-C5)*12', (courier - m['tot_e']) * 12, 8),
        ('Ahorro 3 anos (nominal)', f'(B5-C5)*36', (courier - m['tot_e']) * 36, 8),
        ('Punto de equilibrio (envios/mes)', f'C5/{A["courier"]}',
         m['tot_e'] / S['courier_unit'], 11),
        ('Capacidad de la flota (envios/mes)', f'{A["cap"]}*{A["dias"]}*{Nc}',
         cap_mes, 11),
        ('Utilizacion actual', f'{A["envios"]}/({A["cap"]}*{A["dias"]}*{Nc})',
         S['envios'] / cap_mes, 10),
        ('Holgura para crecer (envios/mes)',
         f'{A["cap"]}*{A["dias"]}*{Nc}-{A["envios"]}', cap_mes - S['envios'], 11),
    ]
    rr = mr + 1
    for label, fml, val, st in metrics:
        sh.text(rr, 1, label, 13)
        sh.formula(rr, 2, fml, val, st)
        rr += 1
    return sh


# ----------------------------------------------------------------------
# HOJA: Tarifa y Margen (operador logistico)
# ----------------------------------------------------------------------
def build_tarifa():
    sh = Sheet('Tarifa y Margen')
    sh.widths = [(1, 34), (2, 16), (3, 16), (4, 16)]
    N = S['Nbase']
    Nc = A['Nbase']
    m = model(N)
    mg = S['margen']
    courier = S['courier_unit'] * S['envios']
    # precio (venta) = costo / (1 - margen)  -> margen sobre venta
    venta_e = m['tot_e'] / (1 - mg)
    venta_p = m['tot_p'] / (1 - mg)

    sh.text(1, 1, 'TARIFA COMO OPERADOR LOGISTICO  (margen objetivo sobre venta)', 1)
    sh.text(2, 1, 'Tarifa que debes cobrar para cubrir costos y dejar el margen '
                  'de Supuestos!B41. Compara contra el courier ($/envio actual).',
            12)
    sh.text(3, 1, 'Margen objetivo:', 12)
    sh.formula(3, 2, A['margen'], mg, 10)

    for c, h in enumerate(['Concepto', 'Flota Electrica', 'Flota Petrolera',
                           'Courier (ref.)'], 1):
        sh.text(5, c, h, 7)

    # fila 6: costo total
    sh.text(6, 1, 'Costo total / mes', 13)
    sh.formula(6, 2, f_tot_e(Nc), m['tot_e'], 8)
    sh.formula(6, 3, f_tot_p(Nc), m['tot_p'], 8)
    sh.formula(6, 4, f'{A["courier"]}*{A["envios"]}', courier, 8)
    # fila 7: costo por envio
    sh.text(7, 1, 'Costo por envio', 13)
    sh.formula(7, 2, f'B6/{A["envios"]}', m['tot_e'] / S['envios'], 8)
    sh.formula(7, 3, f'C6/{A["envios"]}', m['tot_p'] / S['envios'], 8)
    sh.formula(7, 4, A['courier'], S['courier_unit'], 8)
    # fila 8: TARIFA por envio (con margen)
    sh.text(8, 1, 'TARIFA por envio (con margen)', 12)
    sh.formula(8, 2, f'B7/(1-{A["margen"]})', m['tot_e'] / S['envios'] / (1 - mg), 9)
    sh.formula(8, 3, f'C7/(1-{A["margen"]})', m['tot_p'] / S['envios'] / (1 - mg), 9)
    sh.formula(8, 4, A['courier'], S['courier_unit'], 9)
    # fila 9: facturacion mensual
    sh.text(9, 1, 'Facturacion / mes (a esa tarifa)', 13)
    sh.formula(9, 2, f'B6/(1-{A["margen"]})', venta_e, 8)
    sh.formula(9, 3, f'C6/(1-{A["margen"]})', venta_p, 8)
    # fila 10: margen $ mensual
    sh.text(10, 1, 'Margen $ / mes', 13)
    sh.formula(10, 2, 'B9-B6', venta_e - m['tot_e'], 8)
    sh.formula(10, 3, 'C9-C6', venta_p - m['tot_p'], 8)
    # fila 11: margen $ anual
    sh.text(11, 1, 'Margen $ / ano', 13)
    sh.formula(11, 2, 'B10*12', (venta_e - m['tot_e']) * 12, 9)
    sh.formula(11, 3, 'C10*12', (venta_p - m['tot_p']) * 12, 9)
    # fila 12: tarifa vs courier
    sh.text(12, 1, 'Tarifa vs courier (ahorro/envio)', 13)
    sh.formula(12, 2, f'{A["courier"]}-B8',
               S['courier_unit'] - m['tot_e'] / S['envios'] / (1 - mg), 9)
    sh.formula(12, 3, f'{A["courier"]}-C8',
               S['courier_unit'] - m['tot_p'] / S['envios'] / (1 - mg), 9)
    # fila 13: % bajo el courier
    sh.text(13, 1, 'Tarifa % bajo el courier', 13)
    sh.formula(13, 2, f'({A["courier"]}-B8)/{A["courier"]}',
               (S['courier_unit'] - m['tot_e'] / S['envios'] / (1 - mg)) / S['courier_unit'], 10)
    sh.formula(13, 3, f'({A["courier"]}-C8)/{A["courier"]}',
               (S['courier_unit'] - m['tot_p'] / S['envios'] / (1 - mg)) / S['courier_unit'], 10)

    # tabla por tamano de flota
    sh.text(15, 1, 'TARIFA POR ENVIO SEGUN TAMANO DE FLOTA (con margen objetivo)', 2)
    for c, h in enumerate(['Moviles (N)', 'Tarifa Elec.', 'Tarifa Pet.',
                           'Courier (ref.)'], 1):
        sh.text(16, c, h, 7)
    r = 17
    for Nn in [10, 9, 8, 7, 6]:
        mm = model(Nn)
        Ncell = f'$A{r}'
        sh.num(r, 1, Nn, 11)
        sh.formula(r, 2, f'({f_tot_e(Ncell)})/{A["envios"]}/(1-{A["margen"]})',
                   mm['tot_e'] / S['envios'] / (1 - mg), 8)
        sh.formula(r, 3, f'({f_tot_p(Ncell)})/{A["envios"]}/(1-{A["margen"]})',
                   mm['tot_p'] / S['envios'] / (1 - mg), 8)
        sh.formula(r, 4, A['courier'], S['courier_unit'], 8)
        r += 1
    sh.text(r + 1, 1, 'Nota: margen sobre venta -> tarifa = costo / (1 - margen). '
                      'Si la tarifa es menor al courier, eres competitivo Y rentable.',
            12)
    return sh


# ----------------------------------------------------------------------
# HOJA 4: Proyeccion a 3 anos (con alza de combustible)
# ----------------------------------------------------------------------
def build_3anos():
    sh = Sheet('Proyeccion 3 anos')
    sh.widths = [(1, 36)] + [(c, 16) for c in range(2, 6)]
    N = S['Nbase']
    Nc = A['Nbase']
    m = model(N)
    courier = S['courier_unit'] * S['envios']

    sh.text(1, 1, 'PROYECCION A 3 ANOS  -  CLP / ano', 1)
    sh.text(2, 1, 'Flota base (Supuestos!B10). El diesel sube mas rapido, por '
                  'eso el electrico amplia su ventaja.', 12)
    for c, h in enumerate(['CLP / ano', 'Ano 1', 'Ano 2', 'Ano 3',
                           'Acumulado 3 anos'], 1):
        sh.text(4, c, h, 7)

    fixed_e = f'({f_personal(Nc)}+{f_overhead(Nc)}+{f_infra(Nc)}+{A["multas"]})'
    fixed_p = f'({f_personal(Nc)}+{f_overhead(Nc)}+{A["multas"]})'

    def courier_y(e):
        return (f'{A["courier"]}*{A["envios"]}*12*(1+{A["reaj_c"]})^{e}',
                courier * 12 * (1 + S['reaj_courier']) ** e)

    def elec_y(e):
        fu = (1 + S['reaj_uf']) ** e
        fe = (1 + S['alza_elec']) ** e
        fs = (1 + S['reaj_sueldo']) ** e
        val = 12 * (m['leas_e'] * fu + m['sobre'] * fu + m['energ_e'] * fe +
                    (m['personal'] + m['overhead'] + m['infra'] + S['multas']) * fs)
        fml = (f'12*({f_leas_e(Nc)}*(1+{A["reaj_uf"]})^{e}+{f_sobre(Nc)}*'
               f'(1+{A["reaj_uf"]})^{e}+{F_ENERG_E}*(1+{A["alza_e"]})^{e}+'
               f'{fixed_e}*(1+{A["reaj_s"]})^{e})')
        return fml, val

    def pet_y(e):
        fu = (1 + S['reaj_uf']) ** e
        fd = (1 + S['alza_diesel']) ** e
        fs = (1 + S['reaj_sueldo']) ** e
        val = 12 * (m['leas_p'] * fu + m['sobre'] * fu + m['energ_d'] * fd +
                    (m['personal'] + m['overhead'] + S['multas']) * fs)
        fml = (f'12*({f_leas_p(Nc)}*(1+{A["reaj_uf"]})^{e}+{f_sobre(Nc)}*'
               f'(1+{A["reaj_uf"]})^{e}+{F_ENERG_D}*(1+{A["alza_d"]})^{e}+'
               f'{fixed_p}*(1+{A["reaj_s"]})^{e})')
        return fml, val

    rows = [
        ('Courier', courier_y, 8),
        ('Flota Electrica', elec_y, 8),
        ('Flota Petrolera', pet_y, 8),
    ]
    rowmap = {}
    r = 5
    for label, fn, st in rows:
        sh.text(r, 1, label, 13)
        for e in range(3):
            fml, val = fn(e)
            sh.formula(r, 2 + e, fml, round(val), st)
        sh.formula(r, 5, f'SUM(B{r}:D{r})',
                   round(sum(fn(e)[1] for e in range(3))), 9)
        rowmap[label] = r
        r += 1

    rc, re_, rp = rowmap['Courier'], rowmap['Flota Electrica'], rowmap['Flota Petrolera']
    # ahorros y ventaja
    sh.text(r, 1, 'Ahorro Electrica vs Courier', 12)
    for c in range(2, 6):
        L = col_letter(c)
        sh.formula(r, c, f'{L}{rc}-{L}{re_}', 0, 9)
    r += 1
    sh.text(r, 1, 'Ahorro Petrolera vs Courier', 12)
    for c in range(2, 6):
        L = col_letter(c)
        sh.formula(r, c, f'{L}{rc}-{L}{rp}', 0, 9)
    r += 1
    sh.text(r, 1, 'Ventaja Electrica vs Petrolera', 12)
    for c in range(2, 6):
        L = col_letter(c)
        sh.formula(r, c, f'{L}{rp}-{L}{re_}', 0, 9)
    r += 2
    sh.text(r, 1, 'Nota: valor positivo = el electrico es mas barato que el '
                  'petrolero ese ano.', 12)
    return sh


# ----------------------------------------------------------------------
# HOJA 1: Resumen Ejecutivo (para el directorio)
# ----------------------------------------------------------------------
def build_resumen():
    sh = Sheet('Resumen Ejecutivo')
    sh.widths = [(1, 50), (2, 20), (3, 22)]
    N = S['Nbase']
    Nc = A['Nbase']
    m = model(N)
    mg = S['margen']
    courier = S['courier_unit'] * S['envios']
    tot_e = m['tot_e']
    venta_e = tot_e / (1 - mg)
    cap_mes = S['cap'] * S['dias'] * N

    sh.text(1, 1, 'RESUMEN EJECUTIVO  -  Internalizacion de Ultima Milla', 1)
    sh.text(2, 1, 'Evaluacion: flota propia electrica vs. courier actual  (CAV)',
            12)

    # --- Cifras clave (escenario base: flota electrica N moviles) ---
    sh.text(4, 1, 'CIFRAS CLAVE  (escenario base: flota electrica)', 2)
    sh.text(4, 2, '', 2)
    rows = [
        ('Envios por mes', A['envios'], S['envios'], 11),
        ('N de moviles (escenario base)', Nc, N, 11),
        ('Costo courier actual (por envio)', A['courier'], S['courier_unit'], 8),
        ('Costo courier (por mes)', f'{A["courier"]}*{A["envios"]}', courier, 8),
        ('Costo flota electrica (por mes)', f_tot_e(Nc), tot_e, 8),
        ('Costo flota electrica (por envio)', f'B9/{A["envios"]}',
         tot_e / S['envios'], 8),
        ('Ahorro mensual vs courier', f'B8-B9', courier - tot_e, 8),
        ('Ahorro ANUAL vs courier', f'(B8-B9)*12', (courier - tot_e) * 12, 9),
        ('Reduccion de costo', f'(B8-B9)/B8', (courier - tot_e) / courier, 10),
        ('Utilizacion de capacidad', f'{A["envios"]}/({A["cap"]}*{A["dias"]}*{Nc})',
         S['envios'] / cap_mes, 10),
        ('Punto de equilibrio (envios/mes)', f'B9/{A["courier"]}',
         tot_e / S['courier_unit'], 11),
        ('Tarifa con margen objetivo (por envio)', f'B10/(1-{A["margen"]})',
         (tot_e / S['envios']) / (1 - mg), 9),
        ('Margen ANUAL como operador logistico',
         f'(B9/(1-{A["margen"]})-B9)*12', (venta_e - tot_e) * 12, 9),
    ]
    # encabezado tabla
    sh.text(5, 1, 'Indicador', 7)
    sh.text(5, 2, 'Valor', 7)
    r = 6
    for label, fml, val, st in rows:
        sh.text(r, 1, label, 13)
        sh.formula(r, 2, fml, val, st)
        r += 1

    # --- Lectura para el directorio ---
    r += 1
    sh.text(r, 1, 'LECTURA PARA EL DIRECTORIO', 2); sh.text(r, 2, '', 2); r += 1
    bullets = [
        '1. Internalizar la ultima milla reduce el costo de despacho ~33%: de '
        '$2.980 a ~$1.987 por envio, con un ahorro de ~$179 millones al ano.',
        '2. La flota actual estaria sobredimensionada: con capacidad de 120 '
        'puntos por ruta, bastan 6-7 moviles para cubrir los 15.000 envios. Con '
        '7 operamos al 81% y queda holgura para crecer sin sumar vehiculos.',
        '3. Electrico vs petrolero: hoy practicamente empatan en costo. En 3 '
        'anos, con el diesel subiendo mas rapido, el electrico pasa a ser mas '
        'barato desde el Ano 2 y amplia su ventaja (mas sostenibilidad e imagen).',
        '4. Como operador logistico podemos ser mas baratos Y rentables: una '
        'tarifa con 30% de margen (~$2.840/envio) sigue ~5% bajo el courier '
        'actual, con un margen potencial de ~$153 millones al ano.',
    ]
    for b in bullets:
        sh.text(r, 1, b, 14); r += 1

    r += 1
    sh.text(r, 1, 'RECOMENDACION', 2); sh.text(r, 2, '', 2); r += 1
    sh.text(r, 1, 'Avanzar hacia una flota propia de 7 moviles electricos (mejor '
            'balance entre ahorro y holgura operacional). Como alternativa de '
            'menor riesgo inicial, evaluar un modelo hibrido: flota propia para '
            'la base estable + courier para los peaks.', 14); r += 2

    sh.text(r, 1, 'RIESGOS Y CONDICIONES', 2); sh.text(r, 2, '', 2); r += 1
    riesgos = [
        '- Sobrekilometraje: con 7 moviles cada uno recorre ~4.300 km/mes '
        '(contrato 3.000). Ya esta en el costo; conviene negociar el limite de km.',
        '- Costo fijo vs variable: la flota conviene mientras el volumen supere '
        '~10.000 envios/mes; bajo ese nivel el courier vuelve a ser preferible.',
        '- Gestion de personas: se asumen 14 colaboradores (conductores + '
        'peonetas) con su administracion y relacion laboral.',
    ]
    for b in riesgos:
        sh.text(r, 1, b, 14); r += 1

    r += 1
    sh.text(r, 1, 'Nota: todas las cifras se recalculan al editar la hoja '
            '"Supuestos" (celdas amarillas). Valores de referencia a confirmar '
            'con cotizaciones formales.', 14)
    return sh


# ----------------------------------------------------------------------
# Ensamblado del paquete .xlsx
# ----------------------------------------------------------------------
def build(path):
    sheets = [build_resumen(), build_supuestos(), build_escenarios(),
              build_focus(), build_tarifa(), build_3anos()]
    n = len(sheets)

    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
          'content-types">'
          '<Default Extension="rels" ContentType="application/vnd.'
          'openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.'
          'openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.'
          'openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
          + ''.join(f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" '
                    'ContentType="application/vnd.openxmlformats-officedocument.'
                    'spreadsheetml.worksheet+xml"/>' for i in range(n))
          + '</Types>')

    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships"><Relationship Id="rId1" Type="http://schemas.'
            'openxmlformats.org/officeDocument/2006/relationships/'
            'officeDocument" Target="xl/workbook.xml"/></Relationships>')

    sheet_tags = ''.join(
        f'<sheet name="{esc(s.name)}" sheetId="{i+1}" r:id="rId{i+1}"/>'
        for i, s in enumerate(sheets))
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<workbook xmlns="http://schemas.openxmlformats.org/'
                'spreadsheetml/2006/main" xmlns:r="http://schemas.'
                'openxmlformats.org/officeDocument/2006/relationships">'
                f'<sheets>{sheet_tags}</sheets>'
                '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')

    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               '<Relationships xmlns="http://schemas.openxmlformats.org/'
               'package/2006/relationships">'
               + ''.join(f'<Relationship Id="rId{i+1}" Type="http://schemas.'
                         'openxmlformats.org/officeDocument/2006/relationships/'
                         f'worksheet" Target="worksheets/sheet{i+1}.xml"/>'
                         for i in range(n))
               + f'<Relationship Id="rId{n+1}" Type="http://schemas.'
                 'openxmlformats.org/officeDocument/2006/relationships/styles" '
                 'Target="styles.xml"/></Relationships>')

    files = {
        '[Content_Types].xml': ct,
        '_rels/.rels': rels,
        'xl/workbook.xml': workbook,
        'xl/_rels/workbook.xml.rels': wb_rels,
        'xl/styles.xml': STYLES,
    }
    for i, s in enumerate(sheets):
        files[f'xl/worksheets/sheet{i+1}.xml'] = s.xml()

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', files.pop('[Content_Types].xml'))
        for name, data in files.items():
            z.writestr(name, data)
    print('OK ->', path, '(', n, 'hojas )')


if __name__ == '__main__':
    build('Evaluacion_Flota_vs_Courier.xlsx')
