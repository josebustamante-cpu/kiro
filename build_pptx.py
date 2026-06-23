# -*- coding: utf-8 -*-
"""
Generador de PPTX (Office Open XML) usando solo la libreria estandar.
Crea: Programa de Transformacion Omnicanal y Ultima Milla (6 laminas).
Diseno corporativo: navy + acento naranja, tarjetas, pilares, timeline, grafico de barras.
"""
import zipfile, os

# ---------- Geometria ----------
EMU = 914400  # por pulgada
SW = 12192000  # 13.333 in (16:9)
SH = 6858000   # 7.5 in

def IN(v):  # pulgadas -> EMU
    return int(v * EMU)

# ---------- Paleta (referencia CAV) ----------
CHARCOAL = "3A3A3A"   # gris oscuro de cabecera
DARKBOX  = "262626"   # caja negra (callouts)
NAVY     = "123A5E"   # azul marino de iconos
BLUE     = "123A5E"   # = navy
TEAL     = "1C5A8C"   # azul medio (variacion)
GREEN    = "1C5A8C"   # mapeado a azul medio (sin verde, paleta de marca)
ORANGE   = "E8711E"   # naranja de acento
RED      = "D71920"   # rojo del logo CAV
LIGHT    = "F4F4F4"   # fondo claro
CARD     = "FFFFFF"
GREY     = "6B6B6B"
DARK     = "333333"
LINEG    = "DCDCDC"

FONT = "Calibri"
FONTH = "Calibri"

# ---------- Helpers XML ----------
def esc(t):
    return (t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
             .replace('"', '&quot;'))

def run(text, sz=1800, b=False, i=False, color=DARK, font=FONT, spc=None):
    spcx = f' spc="{spc}"' if spc is not None else ''
    return (f'<a:r><a:rPr lang="es-ES" sz="{sz}" b="{1 if b else 0}" '
            f'i="{1 if i else 0}"{spcx} dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{font}"/><a:cs typeface="{font}"/></a:rPr>'
            f'<a:t>{esc(text)}</a:t></a:r>')

def para(runs_xml, algn='l', level=0, bullet=None, bullet_color=ORANGE,
         before=0, after=400, line=None):
    if bullet is None:
        bu = '<a:buNone/>'
        marL, indent = 0, 0
    else:
        bu = (f'<a:buClr><a:srgbClr val="{bullet_color}"/></a:buClr>'
              f'<a:buFont typeface="Arial"/><a:buChar char="{bullet}"/>')
        marL = 228600 * (level + 1)
        indent = -228600
    sp = ''
    if line:
        sp += f'<a:lnSpc><a:spcPct val="{line}"/></a:lnSpc>'
    if before:
        sp += f'<a:spcBef><a:spcPts val="{before}"/></a:spcBef>'
    if after:
        sp += f'<a:spcAft><a:spcPts val="{after}"/></a:spcAft>'
    pPr = (f'<a:pPr marL="{marL}" indent="{indent}" lvl="{level}" '
           f'algn="{algn}">{sp}{bu}</a:pPr>')
    return f'<a:p>{pPr}{runs_xml}</a:p>'

def empty_para(sz=600):
    return f'<a:p><a:pPr><a:buNone/></a:pPr><a:endParaRPr lang="es-ES" sz="{sz}"/></a:p>'

class IdGen:
    def __init__(self):
        self.n = 1
    def __call__(self):
        self.n += 1
        return self.n

def textbox(idg, x, y, cx, cy, paras, anchor='t', wrap='square', name='tb'):
    body = ''.join(paras)
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{idg()}" name="{name}"/>'
            f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/>'
            f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="{wrap}" anchor="{anchor}" lIns="73152" '
            f'tIns="36576" rIns="73152" bIns="36576"><a:normAutofit/></a:bodyPr>'
            f'<a:lstStyle/>{body}</p:txBody></p:sp>')

def shape(idg, x, y, cx, cy, fill=None, line=None, line_w=12700,
          prst='rect', paras=None, anchor='t', shadow=False, name='sp',
          grad=None, flipv=False, fliph=False):
    if grad:
        c1, c2 = grad
        fillxml = (f'<a:gradFill rotWithShape="1"><a:gsLst>'
                   f'<a:gs pos="0"><a:srgbClr val="{c1}"/></a:gs>'
                   f'<a:gs pos="100000"><a:srgbClr val="{c2}"/></a:gs></a:gsLst>'
                   f'<a:lin ang="5400000" scaled="0"/></a:gradFill>')
    elif fill:
        fillxml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
    else:
        fillxml = '<a:noFill/>'
    lnxml = (f'<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{line}"/>'
             f'</a:solidFill></a:ln>') if line else '<a:ln><a:noFill/></a:ln>'
    effect = ('<a:effectLst><a:outerShdw blurRad="50000" dist="25000" '
              'dir="5400000" rotWithShape="0"><a:srgbClr val="000000">'
              '<a:alpha val="20000"/></a:srgbClr></a:outerShdw></a:effectLst>'
              ) if shadow else ''
    if paras:
        body = (f'<p:txBody><a:bodyPr wrap="square" anchor="{anchor}" '
                f'lIns="73152" tIns="36576" rIns="73152" bIns="36576">'
                f'<a:normAutofit/></a:bodyPr><a:lstStyle/>{"".join(paras)}'
                f'</p:txBody>')
    else:
        body = ('<p:txBody><a:bodyPr/><a:lstStyle/><a:p>'
                '<a:endParaRPr lang="es-ES"/></a:p></p:txBody>')
    geom = f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>'
    flip = ''
    if flipv:
        flip += ' flipV="1"'
    if fliph:
        flip += ' flipH="1"'
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{idg()}" name="{name}"/>'
            f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm{flip}><a:off x="{x}" y="{y}"/>'
            f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>{geom}{fillxml}{lnxml}'
            f'{effect}</p:spPr>{body}</p:sp>')

def slide_xml(shapes, bg=CARD):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:bg><p:bgPr>'
        f'<a:solidFill><a:srgbClr val="{bg}"/></a:solidFill>'
        '<a:effectLst/></p:bgPr></p:bg><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + ''.join(shapes) +
        '</p:spTree></p:cSld>'
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')

# ---------- Componentes de diseno ----------
def cav_logo(idg, x, y, d):
    """Logo CAV: circulo rojo con borde blanco punteado y texto CAV."""
    s = []
    s.append(shape(idg, x, y, d, d, fill='FFFFFF', prst='ellipse', name='logoring'))
    inset = int(d * 0.085)
    s.append(shape(idg, x + inset, y + inset, d - 2*inset, d - 2*inset, fill=RED,
                   prst='ellipse',
                   paras=[para(run('CAV', sz=1500, b=True, color='FFFFFF',
                                   spc=120), algn='ctr', after=0)], anchor='ctr',
                   name='logo'))
    return s

def header(idg, title, kicker, idx):
    """Banda superior carbon con titulo blanco y logo CAV (estilo CAV)."""
    s = []
    bandh = IN(1.32)
    # banda carbon
    s.append(shape(idg, 0, 0, SW, bandh, fill=CHARCOAL, name='band'))
    # logo CAV a la derecha
    s += cav_logo(idg, IN(11.55), IN(0.18), IN(0.95))
    # kicker
    s.append(textbox(idg, IN(0.6), IN(0.2), IN(10.5), IN(0.4),
        [para(run(kicker.upper(), sz=1050, b=True, color=ORANGE, spc=200),
              after=0)]))
    # titulo
    s.append(textbox(idg, IN(0.58), IN(0.5), IN(10.6), IN(0.8),
        [para(run(title, sz=2600, b=True, color='FFFFFF'), after=0)]))
    # numero de lamina en circulo naranja
    s.append(shape(idg, IN(0.62), IN(1.5), IN(0.46), IN(0.46), fill=ORANGE,
                   prst='ellipse',
                   paras=[para(run(f'{idx}', sz=1300, b=True, color='FFFFFF'),
                               algn='ctr', after=0)], anchor='ctr'))
    # linea divisoria naranja bajo cabecera
    s.append(shape(idg, 0, bandh, SW, IN(0.05), fill=ORANGE, name='rule'))
    return s

def footer(idg):
    return [textbox(idg, IN(0.6), IN(7.05), IN(11), IN(0.35),
        [para(run('CAV  |  Programa de Transformacion Omnicanal y Ultima Milla',
                  sz=900, color=GREY), after=0)])]

def footer(idg):
    return [textbox(idg, IN(0.6), IN(7.0), IN(11), IN(0.4),
        [para(run('CAV  |  Programa de Transformacion Omnicanal y Ultima Milla',
                  sz=900, color=GREY), after=0)])]

# ---------- LAMINA 1: Vision General (portada) ----------
def slide1():
    idg = IdGen()
    s = []
    # fondo carbon con degradado
    s.append(shape(idg, 0, 0, SW, SH, grad=('454545', '2B2B2B'), name='bg'))
    # banda lateral derecha azul marino
    s.append(shape(idg, IN(8.7), 0, IN(4.7), SH, fill=NAVY, prst='rect',
                   name='band'))
    s.append(shape(idg, IN(8.4), 0, IN(0.16), SH, fill=ORANGE, name='line'))
    # logo CAV
    s += cav_logo(idg, IN(11.7), IN(0.55), IN(1.05))
    # etiqueta superior
    s.append(textbox(idg, IN(0.85), IN(0.95), IN(8), IN(0.5),
        [para(run('LAMINA 1  ·  VISION GENERAL', sz=1200, b=True,
                  color=ORANGE, spc=240), after=0)]))
    # titulo principal
    s.append(textbox(idg, IN(0.8), IN(1.75), IN(8.1), IN(2.6),
        [para(run('Programa de Transformacion', sz=4000, b=True,
                  color='FFFFFF'), after=100),
         para(run('Omnicanal y Ultima Milla', sz=4000, b=True,
                  color=ORANGE), after=0)]))
    # subtitulo objetivo
    s.append(textbox(idg, IN(0.85), IN(4.05), IN(7.4), IN(0.6),
        [para(run('OBJETIVO', sz=1300, b=True, color=ORANGE, spc=200), after=0)]))
    s.append(textbox(idg, IN(0.85), IN(4.45), IN(7.55), IN(0.7),
        [para(run('Mejorar la experiencia de compra y entrega de nuestros '
                  'socios mediante:', sz=1500, color='D7E2EC'), after=0)]))
    # bullets de objetivo
    objetivos = [
        'Direcciones de clientes confiables.',
        'Automatizacion de incidencias de distribucion.',
        'Entregas mas rapidas.',
        'Mayor disponibilidad de inventario.',
        'Integracion omnicanal de tiendas y centros de distribucion.',
    ]
    pp = []
    for o in objetivos:
        pp.append(para(run(o, sz=1350, color='FFFFFF'), bullet='▶',
                       bullet_color=ORANGE, after=320, level=0))
    s.append(textbox(idg, IN(0.85), IN(5.15), IN(7.6), IN(2.0), pp))
    # iconos / chips en banda derecha
    chips = ['DATOS', 'LOGISTICA', 'IA EN RUTA', 'OMNICANAL']
    y = IN(2.0)
    for c in chips:
        s.append(shape(idg, IN(9.0), y, IN(3.9), IN(0.78), fill='0C2742',
                       line=ORANGE, line_w=9525, prst='roundRect',
                       paras=[para(run(c, sz=1400, b=True, color='FFFFFF',
                                       spc=160), algn='ctr', after=0)],
                       anchor='ctr'))
        y += IN(1.0)
    return slide_xml(s, bg=NAVY)

# ---------- Iconos (formas nativas PPT, estilo insignia CAV) ----------
LW = 19050   # grosor de linea de glifos (~1.5pt)
WH = 'FFFFFF'

def icon_badge(idg, cx, cy, d, color):
    return [shape(idg, cx - d // 2, cy - d // 2, d, d, fill=color,
                  prst='ellipse', name='badge', shadow=True)]

def icon_datos(idg, cx, cy):
    w, h = IN(0.46), IN(0.6)
    return [shape(idg, cx - w // 2, cy - h // 2, w, h, prst='can',
                  line=WH, line_w=LW, name='db')]

def icon_distribucion(idg, cx, cy):
    s = []
    s.append(shape(idg, cx - IN(0.44), cy - IN(0.22), IN(0.52), IN(0.36),
                   prst='roundRect', line=WH, line_w=LW, name='cargo'))
    s.append(shape(idg, cx + IN(0.1), cy - IN(0.04), IN(0.3), IN(0.2),
                   prst='round1Rect', line=WH, line_w=LW, name='cab'))
    for dx in (-IN(0.26), IN(0.18)):
        s.append(shape(idg, cx + dx, cy + IN(0.16), IN(0.17), IN(0.17),
                       prst='ellipse', line=WH, line_w=LW, name='wheel'))
    return s

def icon_cliente(idg, cx, cy):
    s = []
    s.append(shape(idg, cx - IN(0.16), cy - IN(0.36), IN(0.3), IN(0.3),
                   prst='ellipse', line=WH, line_w=LW, name='head'))
    s.append(shape(idg, cx - IN(0.3), cy + IN(0.0), IN(0.58), IN(0.34),
                   prst='trapezoid', line=WH, line_w=LW, name='body'))
    s.append(shape(idg, cx + IN(0.12), cy - IN(0.42), IN(0.24), IN(0.22),
                   prst='heart', fill=WH, name='heart'))
    return s

def icon_retail(idg, cx, cy):
    s = []
    s.append(shape(idg, cx - IN(0.36), cy - IN(0.34), IN(0.72), IN(0.18),
                   prst='trapezoid', line=WH, line_w=LW, name='awning',
                   flipv=True))
    s.append(shape(idg, cx - IN(0.28), cy - IN(0.16), IN(0.56), IN(0.5),
                   prst='rect', line=WH, line_w=LW, name='store'))
    s.append(shape(idg, cx - IN(0.09), cy + IN(0.05), IN(0.18), IN(0.29),
                   prst='rect', line=WH, line_w=LW, name='door'))
    return s

ICON_FN = {
    'Datos': icon_datos,
    'Distribucion': icon_distribucion,
    'Experiencia Cliente': icon_cliente,
    'Retail': icon_retail,
}

# ---------- LAMINA 2: Situacion Actual (4 tarjetas verticales) ----------
def slide2():
    idg = IdGen()
    s = header(idg, 'Situacion Actual', 'El desafio de CAV', 2)
    # parrafo introductorio (texto mejorado)
    s.append(shape(idg, IN(0.62), IN(1.55), IN(0.1), IN(1.0), fill=ORANGE,
                   name='introtab'))
    s.append(textbox(idg, IN(0.85), IN(1.5), IN(11.9), IN(1.1),
        [para(run('El Club de Amantes del Vino (CAV) enfrenta desafios '
                  'criticos en su operacion logistica y comercial: direcciones '
                  'de clientes inexactas que derivan en entregas fallidas, '
                  'tiempos de despacho elevados en la Region Metropolitana y '
                  'canales de venta que operan de forma aislada. Este programa '
                  'aborda estos retos de manera integral, bajo una unica '
                  'estrategia omnicanal.', sz=1400, color=DARK), after=0,
              line=112000)]))
    # 4 tarjetas verticales (diseno dinamico, insignias naranjas)
    cards = [
        ('Datos', [
            'Mas de 35.000 clientes con direcciones de calidad dispar.',
            'Multiples canales de captura sin estandar unico.']),
        ('Distribucion', [
            'Entregas fallidas por errores de direccion.',
            'Reprocesos y carga retornada al CD.']),
        ('Experiencia Cliente', [
            'Retrasos en las entregas.',
            'Falta de entregas same day en la RM.']),
        ('Retail', [
            'Venta limitada al stock de cada tienda.',
            'Sin visibilidad del inventario de la red.']),
    ]
    x0 = IN(0.62)
    cw = IN(2.875)
    gx = IN(0.205)
    y0 = IN(2.7)
    ch = IN(4.05)
    HALO = 'FBE2CB'  # naranja claro (halo)
    for k, (title, items) in enumerate(cards):
        cx = x0 + k * (cw + gx)
        ccx = cx + cw // 2
        # tarjeta
        s.append(shape(idg, cx, y0, cw, ch, fill=CARD, line=LINEG, line_w=9525,
                       prst='roundRect', shadow=True, name='card'))
        # marco superior navy (esquinas sup. redondeadas)
        s.append(shape(idg, cx, y0, cw, IN(0.16), fill=NAVY,
                       prst='round2SameRect', name='top'))
        # marco inferior naranja (esquinas inf. redondeadas)
        s.append(shape(idg, cx, y0 + ch - IN(0.16), cw, IN(0.16), fill=ORANGE,
                       prst='round2SameRect', flipv=True, name='bot'))
        # halo + insignia naranja + icono blanco
        bcy = y0 + IN(1.0)
        s.append(shape(idg, ccx - IN(0.8), bcy - IN(0.8), IN(1.6), IN(1.6),
                       fill=HALO, prst='ellipse', name='halo'))
        s += icon_badge(idg, ccx, bcy, IN(1.3), ORANGE)
        s += ICON_FN[title](idg, ccx, bcy)
        # titulo
        s.append(textbox(idg, cx + IN(0.1), y0 + IN(1.78), cw - IN(0.2), IN(0.7),
            [para(run(title, sz=1500, b=True, color=NAVY), algn='ctr',
                  after=0, line=96000)]))
        # subrayado naranja
        s.append(shape(idg, ccx - IN(0.32), y0 + IN(2.42), IN(0.64), IN(0.06),
                       fill=ORANGE, prst='roundRect', name='uline'))
        # bullets
        pp = [para(run(it, sz=1200, color=DARK), bullet='▸', bullet_color=ORANGE,
                   after=260, line=104000) for it in items]
        s.append(textbox(idg, cx + IN(0.24), y0 + IN(2.62), cw - IN(0.44),
                         ch - IN(2.9), pp))
    return slide_xml(s)

# ---------- LAMINA 3: Los 4 Frentes (pilares) ----------
def slide3():
    idg = IdGen()
    s = header(idg, 'Los 4 Frentes del Programa', 'Cuatro pilares de trabajo', 3)
    pilares = [
        ('01', 'Normalizacion Historica de Direcciones', BLUE, [
            'Corregir la base actual de clientes.',
            'Meta: normalizar 35.000 registros.',
            'Duracion: 1 mes.']),
        ('02', 'Captura Correcta desde el Origen', TEAL, [
            'Formularios estandarizados en Web, App, Carrito, ERP, QR ejecutivos, Eventos y Tiendas.',
            'Resultado: evitar que el problema vuelva a generarse.']),
        ('03', 'IA para Gestion de Incidencias en Ruta', ORANGE, [
            'Agente IA SimpliRoute: contacto automatico, correccion de direcciones, reagendamiento y actualizacion en tiempo real.',
            'Objetivo: reducir entregas fallidas y retornos al CD; mejorar experiencia.']),
        ('04', 'Omnicanalidad Retail', NAVY, [
            'Tiendas venden inventario de toda la red.',
            'Modelo: Cliente > Tienda > Cualquier Bodega > Despacho.',
            'Beneficios: mayor disponibilidad, menos ventas perdidas.']),
    ]
    x0 = IN(0.62)
    cw = IN(2.93)
    gx = IN(0.13)
    y0 = IN(1.95)
    ch = IN(4.65)
    for k, (num, title, color, items) in enumerate(pilares):
        cx = x0 + k * (cw + gx)
        # cuerpo de la tarjeta
        s.append(shape(idg, cx, y0, cw, ch, fill=CARD, line=LINEG, line_w=9525,
                       prst='roundRect', shadow=True, name='pillar'))
        # cabecera de color
        s.append(shape(idg, cx, y0, cw, IN(1.15), fill=color, prst='round2SameRect',
                       name='phead'))
        # numero
        s.append(textbox(idg, cx, y0 + IN(0.12), cw, IN(0.55),
            [para(run('PILAR ' + num, sz=1150, b=True, color='FFFFFF',
                      spc=180), algn='ctr', after=0)]))
        s.append(textbox(idg, cx + IN(0.12), y0 + IN(0.5), cw - IN(0.24),
                         IN(0.65),
            [para(run(title, sz=1180, b=True, color='FFFFFF'), algn='ctr',
                  after=0, line=96000)]))
        pp = []
        for it in items:
            pp.append(para(run(it, sz=1180, color=DARK), bullet='•',
                           bullet_color=color, after=280, line=100000))
        s.append(textbox(idg, cx + IN(0.16), y0 + IN(1.30), cw - IN(0.32),
                         ch - IN(1.45), pp))
    s += footer(idg)
    return slide_xml(s)

# ---------- LAMINA 4: Caso de Negocio IA SimpliRoute ----------
def slide4():
    idg = IdGen()
    s = header(idg, 'Caso de Negocio  ·  IA SimpliRoute',
               'Agente IA de gestion de entregas', 4)
    # tres KPIs
    kpis = [
        ('Inversion', 'USD 1.500', '/ mes', NAVY),
        ('Beneficio estimado', 'USD 3.500', '/ mes en ahorro', NAVY),
        ('Beneficio neto', 'USD 2.000', '/ mes  ·  ROI ~133%', ORANGE),
    ]
    x0 = IN(0.62)
    cw = IN(3.83)
    gx = IN(0.2)
    y0 = IN(1.9)
    ch = IN(1.85)
    for k, (label, val, sub, color) in enumerate(kpis):
        cx = x0 + k * (cw + gx)
        s.append(shape(idg, cx, y0, cw, ch, fill=CARD, line=LINEG, line_w=9525,
                       prst='roundRect', shadow=True))
        s.append(shape(idg, cx, y0, IN(0.12), ch, fill=color))
        s.append(textbox(idg, cx + IN(0.3), y0 + IN(0.18), cw - IN(0.5), IN(0.45),
            [para(run(label.upper(), sz=1100, b=True, color=GREY, spc=120),
                  after=0)]))
        s.append(textbox(idg, cx + IN(0.28), y0 + IN(0.58), cw - IN(0.5), IN(0.8),
            [para(run(val, sz=3200, b=True, color=color), after=0)]))
        s.append(textbox(idg, cx + IN(0.32), y0 + IN(1.32), cw - IN(0.5), IN(0.4),
            [para(run(sub, sz=1150, color=DARK), after=0)]))
    # ---- Grafico de barras (dibujado con rectangulos) ----
    gx0, gy0 = IN(0.62), IN(4.05)
    gW, gH = IN(6.0), IN(2.45)
    s.append(shape(idg, gx0, gy0, gW, gH, fill=LIGHT, line=LINEG, line_w=9525,
                   prst='roundRect'))
    s.append(textbox(idg, gx0 + IN(0.25), gy0 + IN(0.12), gW - IN(0.5), IN(0.4),
        [para(run('Comparativo mensual (USD)', sz=1250, b=True, color=NAVY),
              after=0)]))
    bars = [('Costo', 1500, '7A7A7A'), ('Ahorro', 3500, NAVY),
            ('Beneficio Neto', 2000, ORANGE)]
    maxv = 3500
    base_y = gy0 + gH - IN(0.55)
    plot_h = IN(1.35)
    bx = gx0 + IN(0.55)
    bw = IN(1.15)
    bgap = IN(0.62)
    # eje
    s.append(shape(idg, gx0 + IN(0.35), base_y, gW - IN(0.7), IN(0.018),
                   fill=LINEG))
    for label, val, color in bars:
        h = int(plot_h * (val / maxv))
        by = base_y - h
        s.append(shape(idg, bx, by, bw, h, fill=color, prst='roundRect'))
        # valor
        s.append(textbox(idg, bx - IN(0.1), by - IN(0.42), bw + IN(0.2), IN(0.4),
            [para(run(f'{val:,}'.replace(',', '.'), sz=1200, b=True,
                      color=DARK), algn='ctr', after=0)]))
        # etiqueta
        s.append(textbox(idg, bx - IN(0.15), base_y + IN(0.06), bw + IN(0.3),
                         IN(0.4),
            [para(run(label, sz=1050, color=GREY), algn='ctr', after=0)]))
        bx += bw + bgap
    # ---- Beneficios adicionales ----
    px, py = IN(6.95), IN(4.05)
    pw, ph = IN(5.78), IN(2.45)
    s.append(shape(idg, px, py, pw, ph, fill=NAVY, prst='roundRect', shadow=True))
    s.append(textbox(idg, px + IN(0.32), py + IN(0.18), pw - IN(0.5), IN(0.45),
        [para(run('BENEFICIOS ADICIONALES', sz=1200, b=True, color=ORANGE,
                  spc=160), after=0)]))
    extra = ['Menor carga operativa.', 'Menos devoluciones.',
             'Mejor NPS.', 'Mayor satisfaccion de clientes.']
    pp = [para(run(e, sz=1300, color='FFFFFF'), bullet='✔',
               bullet_color=ORANGE, after=300) for e in extra]
    s.append(textbox(idg, px + IN(0.3), py + IN(0.66), pw - IN(0.55),
                     ph - IN(0.8), pp))
    s += footer(idg)
    return slide_xml(s)

# ---------- LAMINA 5: Roadmap (timeline) ----------
def slide5():
    idg = IdGen()
    s = header(idg, 'Roadmap', 'Hoja de ruta del programa', 7)
    steps = [
        ('May - Jun', 'Evaluacion IA SimpliRoute', True, BLUE),
        ('Junio', 'Correccion masiva base de clientes', True, BLUE),
        ('Jun - Jul', 'Desarrollo formularios normalizados', True, TEAL),
        ('Julio', 'Implementacion Agente IA', True, TEAL),
        ('Ago - Sep', 'Expansion Same Day RM', False, ORANGE),
        ('Sep - Oct', 'Proyecto Omnicanalidad Retail', False, ORANGE),
    ]
    # linea horizontal central
    line_y = IN(3.75)
    s.append(shape(idg, IN(0.9), line_y, IN(11.5), IN(0.05), fill=LINEG))
    n = len(steps)
    x0 = IN(1.05)
    span = IN(11.2)
    step_w = span // n
    for i, (period, desc, done, color) in enumerate(steps):
        cx = x0 + i * step_w + step_w // 2
        # nodo
        node = IN(0.5)
        s.append(shape(idg, cx - node // 2, line_y - node // 2 + IN(0.02), node,
                       node, fill=color, prst='ellipse',
                       paras=[para(run('✓' if done else '~', sz=1500, b=True,
                                       color='FFFFFF'), algn='ctr', after=0)],
                       anchor='ctr'))
        cardw = step_w - IN(0.2)
        cardx = cx - cardw // 2
        if i % 2 == 0:
            # arriba
            cardy = line_y - IN(1.85)
            s.append(shape(idg, cardx, cardy, cardw, IN(1.5), fill=CARD,
                           line=LINEG, line_w=9525, prst='roundRect',
                           shadow=True))
            s.append(textbox(idg, cardx, cardy + IN(0.14), cardw, IN(0.4),
                [para(run(period.upper(), sz=1050, b=True, color=color),
                      algn='ctr', after=0)]))
            s.append(textbox(idg, cardx + IN(0.06), cardy + IN(0.52), cardw - IN(0.12),
                             IN(0.9),
                [para(run(desc, sz=1150, color=DARK), algn='ctr', after=0,
                      line=98000)]))
            # conector
            s.append(shape(idg, cx - IN(0.01), cardy + IN(1.5), IN(0.02),
                           line_y - (cardy + IN(1.5)), fill=color))
        else:
            cardy = line_y + IN(0.35)
            s.append(shape(idg, cx - IN(0.01), line_y, IN(0.02),
                           cardy - line_y, fill=color))
            s.append(shape(idg, cardx, cardy, cardw, IN(1.5), fill=CARD,
                           line=LINEG, line_w=9525, prst='roundRect',
                           shadow=True))
            s.append(textbox(idg, cardx, cardy + IN(0.14), cardw, IN(0.4),
                [para(run(period.upper(), sz=1050, b=True, color=color),
                      algn='ctr', after=0)]))
            s.append(textbox(idg, cardx + IN(0.06), cardy + IN(0.52), cardw - IN(0.12),
                             IN(0.9),
                [para(run(desc, sz=1150, color=DARK), algn='ctr', after=0,
                      line=98000)]))
    # leyenda
    s.append(shape(idg, IN(0.62), IN(6.45), IN(0.28), IN(0.28), fill=NAVY,
                   prst='ellipse'))
    s.append(textbox(idg, IN(0.95), IN(6.4), IN(3), IN(0.4),
        [para(run('Completado', sz=1050, color=GREY), after=0)]))
    s.append(shape(idg, IN(2.7), IN(6.45), IN(0.28), IN(0.28), fill=ORANGE,
                   prst='ellipse'))
    s.append(textbox(idg, IN(3.03), IN(6.4), IN(3), IN(0.4),
        [para(run('En curso / planificado', sz=1050, color=GREY), after=0)]))
    return slide_xml(s)

# ---------- LAMINA 6: Impacto Esperado ----------
def slide6():
    idg = IdGen()
    s = header(idg, 'Impacto Esperado', 'Resultados en cuatro dimensiones', 8)
    blocks = [
        ('Operacion', NAVY, ['Menos entregas fallidas.', 'Menos reprocesos.',
                             'Menos retornos al CD.']),
        ('Cliente', ORANGE, ['Direcciones correctas.', 'Entregas mas rapidas.',
                           'Mejor experiencia de compra.']),
        ('Comercial', NAVY, ['Mas ventas.', 'Mayor disponibilidad de inventario.',
                               'Omnicanalidad.']),
        ('Financiero', ORANGE, ['Ahorro operativo.', 'Mayor productividad logistica.',
                               'Incremento potencial de ingresos.']),
    ]
    x0 = IN(0.62)
    cw = IN(2.93)
    gx = IN(0.13)
    y0 = IN(1.9)
    ch = IN(2.75)
    for k, (title, color, items) in enumerate(blocks):
        cx = x0 + k * (cw + gx)
        s.append(shape(idg, cx, y0, cw, ch, fill=CARD, line=LINEG, line_w=9525,
                       prst='roundRect', shadow=True))
        s.append(shape(idg, cx, y0, cw, IN(0.7), fill=color,
                       prst='round2SameRect',
                       paras=[para(run(title, sz=1500, b=True, color='FFFFFF'),
                                   algn='ctr', after=0)], anchor='ctr'))
        pp = [para(run(it, sz=1200, color=DARK), bullet='•', bullet_color=color,
                   after=300, line=100000) for it in items]
        s.append(textbox(idg, cx + IN(0.18), y0 + IN(0.85), cw - IN(0.36),
                         ch - IN(1.0), pp))
    # mensaje final (caja negra estilo callout CAV)
    my = IN(4.95)
    s.append(shape(idg, IN(0.62), my, IN(12.1), IN(1.7), fill=DARKBOX,
                   prst='roundRect', shadow=True))
    s.append(shape(idg, IN(0.62), my, IN(0.14), IN(1.7), fill=ORANGE))
    s.append(textbox(idg, IN(1.0), my + IN(0.16), IN(11.4), IN(0.4),
        [para(run('MENSAJE FINAL PARA GERENCIA', sz=1150, b=True, color=ORANGE,
                  spc=180), after=0)]))
    s.append(textbox(idg, IN(1.0), my + IN(0.55), IN(11.45), IN(1.05),
        [para(run('Estamos construyendo las capacidades para que CAV opere con '
                  'una vision omnicanal real, donde los datos de clientes, la '
                  'logistica inteligente y la disponibilidad de inventario '
                  'trabajen integradas para mejorar la experiencia de compra, '
                  'reducir costos operacionales y habilitar el crecimiento '
                  'futuro del negocio.', sz=1300, color='E6EDF3', i=True),
              after=0, line=108000)]))
    return slide_xml(s)

# ---------- Modelo economico (parametros compartidos) ----------
FALLIDOS_MES = 1051
PCT_ATAC = 0.75
PCT_REDES = 0.80
COSTO_UNIT = 3467
USD_CLP = 950
COSTO_IA_MES = 1500 * USD_CLP            # 1.425.000
COSTO_IA_ANIO = COSTO_IA_MES * 12        # 17.100.000
HHEE_VALOR = 7533
HHEE_NORM = 40
NORMALIZACION = HHEE_NORM * HHEE_VALOR    # 301.320 (inversion unica)


def clp(n, signed=False):
    """Formatea un numero como pesos chilenos: $1.234.567 / +$.. / -$.."""
    s_ = f"{abs(int(round(n))):,.0f}".replace(",", ".")
    if signed:
        return ("+$" if n >= 0 else "-$") + s_
    return ("-$" if n < 0 else "$") + s_


def draw_table(idg, x, y, col_defs, rows, header_h=0.45, row_h=0.5,
               header_fill=NAVY, header_sz=950, cell_sz=1000):
    """Dibuja una tabla. col_defs: [(titulo, ancho_in, alineacion)].
    rows: [(celdas, fill)] donde celdas=[(texto, color, bold)] y fill=hex|None."""
    sh = []
    cx = x
    for (htext, w, al) in col_defs:
        cw = IN(w)
        sh.append(shape(idg, cx, y, cw, IN(header_h), fill=header_fill,
                        line='FFFFFF', line_w=6350,
                        paras=[para(run(htext, sz=header_sz, b=True,
                                        color='FFFFFF'), algn=al, after=0,
                                    line=90000)], anchor='ctr', name='th'))
        cx += cw
    ry = y + IN(header_h)
    for ri, (cells, rfill) in enumerate(rows):
        fill = rfill if rfill else (CARD if ri % 2 == 0 else LIGHT)
        cx = x
        for ci, (_h, w, al) in enumerate(col_defs):
            cw = IN(w)
            txt, col, bold = cells[ci]
            sh.append(shape(idg, cx, ry, cw, IN(row_h), fill=fill,
                            line=LINEG, line_w=6350,
                            paras=[para(run(txt, sz=cell_sz, b=bold, color=col),
                                        algn=al, after=0, line=90000)],
                            anchor='ctr', name='td'))
            cx += cw
        ry += IN(row_h)
    return sh, ry


def supuestos_strip(idg, sy):
    s = []
    s.append(shape(idg, IN(0.62), sy, IN(12.1), IN(0.78), fill=DARKBOX,
                   prst='roundRect', name='supuestos'))
    s.append(shape(idg, IN(0.62), sy, IN(0.12), IN(0.78), fill=ORANGE))
    s.append(textbox(idg, IN(0.95), sy + IN(0.04), IN(11.6), IN(0.62),
        [para(run('SUPUESTOS   ', sz=1000, b=True, color=ORANGE, spc=120) +
              run('1.051 despachos fallidos/mes (promedio 13 meses)  ·  75% '
                  'atacable (788/mes)  ·  80% de los retornos genera redespacho '
                  ' ·  costo del redespacho $3.467  ·  TC $950/USD  ·  Agente '
                  'IA USD 1.500/mes  ·  Normalizacion 40 HHEE x $7.533',
                  sz=1000, color='E6EDF3'),
              algn='l', after=0)], anchor='ctr'))
    return s


# ---------- LAMINA: Analisis Economico (1) Dimension del problema ----------
def slide_econo_problema():
    idg = IdGen()
    s = header(idg, 'Analisis Economico (1)  ·  Dimension del Problema',
               'Costo actual de los redespachos', 5)

    m = FALLIDOS_MES
    u = COSTO_UNIT
    # franja de supuestos
    s += supuestos_strip(idg, IN(1.5))

    # ---- Cuadro 1 (tabla, lado izquierdo) ----
    s.append(textbox(idg, IN(0.62), IN(2.42), IN(7.2), IN(0.32),
        [para(run('Cuadro 1  ·  Dimension del problema (costo actual)',
                  sz=1250, b=True, color=NAVY), after=0)]))

    NV = 'E4ECF3'   # tinte navy claro
    OR = 'FCE9D6'   # tinte naranja claro
    cols = [('Concepto', 3.55, 'l'), ('Mensual', 1.75, 'ctr'),
            ('Anual', 1.85, 'ctr')]

    def r(c, mes, anio, fill=None, bold=False, ccol=DARK, vcol=DARK):
        return ([(c, ccol, bold), (mes, vcol, bold), (anio, vcol, bold)], fill)

    rows = [
        r('Despachos fallidos', '1.051', '12.612'),
        r('Costo unitario del redespacho', '$3.467', '$3.467'),
        r('Redespachos (80% de los fallidos)', str(round(m*PCT_REDES)),
          str(round(m*PCT_REDES*12))),
        r('Costo total de redespachos', clp(m*PCT_REDES*u),
          clp(m*PCT_REDES*u*12), fill=NV, bold=True, ccol=NAVY, vcol=NAVY),
        r('Fallidos atacables (75%)', str(round(m*PCT_ATAC)),
          str(round(m*PCT_ATAC*12))),
        r('Redespachos atacables (75% x 80%)', str(round(m*PCT_ATAC*PCT_REDES)),
          str(round(m*PCT_ATAC*PCT_REDES*12))),
        r('Costo de redespachos atacables', clp(m*PCT_ATAC*PCT_REDES*u),
          clp(m*PCT_ATAC*PCT_REDES*u*12), fill=OR, bold=True, ccol=ORANGE,
          vcol=ORANGE),
    ]
    tbl, _ = draw_table(idg, IN(0.62), IN(2.78), cols, rows,
                        header_h=0.45, row_h=0.5)
    s += tbl

    # ---- Tarjetas de resumen (lado derecho) ----
    cards = [
        ('COSTO TOTAL DE REDESPACHOS', clp(m*PCT_REDES*u*12),
         'CLP / ano  (100% de los fallidos)', NAVY),
        ('COSTO ATACABLE POR EL PROYECTO', clp(m*PCT_ATAC*PCT_REDES*u*12),
         'CLP / ano  (75% x 80%)', ORANGE),
    ]
    cx = IN(8.05)
    cwd = IN(4.67)
    cy = IN(2.78)
    chh = IN(1.55)
    for label, val, sub, color in cards:
        s.append(shape(idg, cx, cy, cwd, chh, fill=CARD, line=LINEG,
                       line_w=9525, prst='roundRect', shadow=True))
        s.append(shape(idg, cx, cy, IN(0.13), chh, fill=color))
        s.append(textbox(idg, cx + IN(0.32), cy + IN(0.18), cwd - IN(0.5),
                         IN(0.4),
            [para(run(label, sz=1000, b=True, color=GREY, spc=80), after=0,
                  line=96000)]))
        s.append(textbox(idg, cx + IN(0.3), cy + IN(0.6), cwd - IN(0.5), IN(0.6),
            [para(run(val, sz=2800, b=True, color=color), after=0)]))
        s.append(textbox(idg, cx + IN(0.34), cy + IN(1.16), cwd - IN(0.5),
                         IN(0.35),
            [para(run(sub, sz=1000, color=DARK), after=0)]))
        cy += chh + IN(0.15)

    # nota al pie
    s.append(textbox(idg, IN(0.62), IN(6.82), IN(12.1), IN(0.4),
        [para(run('* 3 casuisticas atacables: nadie en casa  ·  direccion '
                  'erronea / no encontrada  ·  cliente ya no vive en la '
                  'direccion.', sz=900, i=True, color=GREY), after=0)]))
    return slide_xml(s)


# ---------- LAMINA: Analisis Economico (2) Ahorro y Beneficio ----------
def slide_econo_beneficio():
    idg = IdGen()
    s = header(idg, 'Analisis Economico (2)  ·  Ahorro y Beneficio',
               'Retorno del Agente IA', 6)

    m = FALLIDOS_MES
    u = COSTO_UNIT
    base_m = m * PCT_ATAC * PCT_REDES * u      # ahorro maximo mensual
    OR = 'FCE9D6'
    escen = [('Conservador (50%)', 0.50), ('Moderado (70%)', 0.70),
             ('Optimista (85%)', 0.85)]

    # ---- Cuadro 2: Ahorro potencial ----
    s.append(textbox(idg, IN(0.62), IN(1.5), IN(12.1), IN(0.32),
        [para(run('Cuadro 2  ·  Ahorro potencial segun efectividad del Agente '
                  'IA', sz=1250, b=True, color=NAVY), after=0)]))
    cols2 = [('Escenario', 3.25, 'l'), ('Casos resueltos/mes', 2.6, 'ctr'),
             ('Ahorro mensual', 3.05, 'ctr'), ('Ahorro anual', 3.05, 'ctr')]
    rows2 = []
    for i, (name, p) in enumerate(escen):
        fill = OR if i == 1 else None
        bold = (i == 1)
        rows2.append((
            [(name, NAVY if i == 1 else DARK, True),
             (str(round(m*PCT_ATAC*p)), DARK, bold),
             (clp(base_m*p), DARK, bold),
             (clp(base_m*p*12), DARK, True)], fill))
    t2y = IN(1.85)
    tbl2, _ = draw_table(idg, IN(0.62), t2y, cols2, rows2,
                         header_h=0.45, row_h=0.52)
    s += tbl2
    # borde naranja fila Moderado
    s.append(shape(idg, IN(0.62), t2y + IN(0.45) + IN(0.52), IN(11.95), IN(0.52),
                   fill=None, line=ORANGE, line_w=22225, prst='rect'))

    # ---- Cuadro 3: Beneficio neto y ROI ----
    s.append(textbox(idg, IN(0.62), IN(3.98), IN(12.1), IN(0.32),
        [para(run('Cuadro 3  ·  Beneficio neto y ROI (descontando el Agente IA: '
                  '$17.100.000/ano)', sz=1250, b=True, color=NAVY), after=0)]))
    cols3 = [('Escenario', 2.55, 'l'), ('Ahorro anual', 2.55, 'ctr'),
             ('Costo Agente IA', 2.5, 'ctr'),
             ('Beneficio neto anual', 2.85, 'ctr'), ('ROI', 1.5, 'ctr')]
    rows3 = []
    for i, (name, p) in enumerate(escen):
        ah_a = round(base_m * p * 12)
        neto = ah_a - COSTO_IA_ANIO
        roi = neto / COSTO_IA_ANIO
        ncol = NAVY if neto >= 0 else RED
        fill = OR if i == 1 else None
        rows3.append((
            [(name, NAVY if i == 1 else DARK, True),
             (clp(ah_a), DARK, False),
             (clp(COSTO_IA_ANIO), GREY, False),
             (clp(neto, signed=True), ncol, True),
             (f'{roi*100:+.0f}%', ncol, True)], fill))
    t3y = IN(4.32)
    tbl3, _ = draw_table(idg, IN(0.62), t3y, cols3, rows3,
                         header_h=0.45, row_h=0.52)
    s += tbl3
    s.append(shape(idg, IN(0.62), t3y + IN(0.45) + IN(0.52), IN(11.95), IN(0.52),
                   fill=None, line=ORANGE, line_w=22225, prst='rect'))

    # ---- Conclusion ----
    cy = IN(6.42)
    s.append(shape(idg, IN(0.62), cy, IN(12.1), IN(0.6), fill=DARKBOX,
                   prst='roundRect', shadow=True))
    s.append(shape(idg, IN(0.62), cy, IN(0.14), IN(0.6), fill=ORANGE))
    s.append(textbox(idg, IN(0.98), cy + IN(0.05), IN(11.6), IN(0.55),
        [para(run('CONCLUSION   ', sz=1000, b=True, color=ORANGE, spc=120) +
              run('Punto de equilibrio ~65% de efectividad. Inversion unica de '
                  'normalizacion: $301.320 (40 HHEE x $7.533). Cifras '
                  'conservadoras: no incluyen mejor CX, ventas recuperadas ni '
                  'menor inventario inmovilizado.', sz=1000, color='E6EDF3'),
              algn='l', after=0)], anchor='ctr'))
    return slide_xml(s)

# ==================================================================
# Boilerplate OOXML (theme, master, layout, presentation, content types)
# ==================================================================
CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
{slide_overrides}
</Types>'''

RELS_ROOT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>'''

def presentation_xml(n):
    ids = ''.join(f'<p:sldId id="{256+i}" r:id="rId{2+i}"/>' for i in range(n))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        f'<p:sldIdLst>{ids}</p:sldIdLst>'
        f'<p:sldSz cx="{SW}" cy="{SH}" type="screen16x9"/>'
        '<p:notesSz cx="6858000" cy="9144000"/></p:presentation>')

def presentation_rels(n):
    rels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>']
    for i in range(n):
        rels.append(f'<Relationship Id="rId{2+i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>')
    rels.append(f'<Relationship Id="rId{2+n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>')
    rels.append(f'<Relationship Id="rId{3+n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + ''.join(rels) + '</Relationships>')

PRES_PROPS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>')

def theme_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="CAV">'
    '<a:themeElements><a:clrScheme name="CAV">'
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
    f'<a:dk2><a:srgbClr val="{NAVY}"/></a:dk2>'
    f'<a:lt2><a:srgbClr val="{LIGHT}"/></a:lt2>'
    f'<a:accent1><a:srgbClr val="{BLUE}"/></a:accent1>'
    f'<a:accent2><a:srgbClr val="{ORANGE}"/></a:accent2>'
    f'<a:accent3><a:srgbClr val="{TEAL}"/></a:accent3>'
    f'<a:accent4><a:srgbClr val="{GREEN}"/></a:accent4>'
    f'<a:accent5><a:srgbClr val="{GREY}"/></a:accent5>'
    f'<a:accent6><a:srgbClr val="{DARK}"/></a:accent6>'
    '<a:hlink><a:srgbClr val="1F6FB2"/></a:hlink>'
    '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
    f'<a:fontScheme name="CAV"><a:majorFont><a:latin typeface="{FONTH}"/>'
    '<a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
    f'<a:minorFont><a:latin typeface="{FONT}"/><a:ea typeface=""/>'
    '<a:cs typeface=""/></a:minorFont></a:fontScheme>'
    '<a:fmtScheme name="CAV">'
    '<a:fillStyleLst>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
    '</a:fillStyleLst>'
    '<a:lnStyleLst>'
    '<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
    '<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
    '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
    '</a:lnStyleLst>'
    '<a:effectStyleLst>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>'
    '</a:effectStyleLst>'
    '<a:bgFillStyleLst>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
    '</a:bgFillStyleLst>'
    '</a:fmtScheme></a:themeElements></a:theme>')

SLIDE_MASTER = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>'
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
    'accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" '
    'accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
    '<p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="2800">'
    '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
    '<a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>'
    '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800"><a:solidFill>'
    '<a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/>'
    '</a:defRPr></a:lvl1pPr></p:bodyStyle><p:otherStyle/></p:txStyles></p:sldMaster>')

SLIDE_MASTER_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
    '</Relationships>')

SLIDE_LAYOUT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'type="blank" preserve="1"><p:cSld name="Blank">'
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>'
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>')

SLIDE_LAYOUT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
    '</Relationships>')

def slide_rels():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        '</Relationships>')

# ==================================================================
def build(path):
    slides = [slide1(), slide2(), slide3(), slide4(), slide_econo_problema(),
              slide_econo_beneficio(), slide5(), slide6()]
    n = len(slides)
    overrides = '\n'.join(
        f'<Override PartName="/ppt/slides/slide{i+1}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(n))
    files = {
        '[Content_Types].xml': CONTENT_TYPES.format(slide_overrides=overrides),
        '_rels/.rels': RELS_ROOT,
        'ppt/presentation.xml': presentation_xml(n),
        'ppt/_rels/presentation.xml.rels': presentation_rels(n),
        'ppt/presProps.xml': PRES_PROPS,
        'ppt/theme/theme1.xml': theme_xml(),
        'ppt/slideMasters/slideMaster1.xml': SLIDE_MASTER,
        'ppt/slideMasters/_rels/slideMaster1.xml.rels': SLIDE_MASTER_RELS,
        'ppt/slideLayouts/slideLayout1.xml': SLIDE_LAYOUT,
        'ppt/slideLayouts/_rels/slideLayout1.xml.rels': SLIDE_LAYOUT_RELS,
    }
    for i, sx in enumerate(slides):
        files[f'ppt/slides/slide{i+1}.xml'] = sx
        files[f'ppt/slides/_rels/slide{i+1}.xml.rels'] = slide_rels()

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        # content types primero
        z.writestr('[Content_Types].xml', files.pop('[Content_Types].xml'))
        for name, data in files.items():
            z.writestr(name, data)
    print('OK ->', path, '(', n, 'laminas )')

if __name__ == '__main__':
    build('Transformacion_Omnicanal_Ultima_Milla.pptx')
