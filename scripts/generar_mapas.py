#!/usr/bin/env python3
"""
generar_mapas.py
Lee el Excel de Knox (Tabletas_Estructura.xlsx) y genera 35 HTML por entidad.
Ejecutado por GitHub Actions cuando se sube un nuevo Excel.
"""
import pandas as pd
import json, os, re
from datetime import datetime, timezone, timedelta
from pathlib import Path

# ── Configuracion ──────────────────────────────────────────────────
OUTPUT_DIR  = Path('docs')
DATA_DIR    = Path('data')
ZONA_MEXICO = timezone(timedelta(hours=-6))
FECHA_HOY   = datetime.now(ZONA_MEXICO).strftime('%d/%m/%Y %H:%M hora Centro')

ESTADOS = {
    '01':('Aguascalientes','Aguascalientes'),
    '02':('Baja_California','Baja California'),
    '03':('Baja_California_Sur','Baja California Sur'),
    '04':('Campeche','Campeche'),
    '05':('Coahuila','Coahuila'),
    '06':('Colima','Colima'),
    '07':('Chiapas','Chiapas'),
    '08':('Chihuahua','Chihuahua'),
    '09':('CDMX','Ciudad de Mexico'),
    '10':('Durango','Durango'),
    '11':('Guanajuato','Guanajuato'),
    '12':('Guerrero','Guerrero'),
    '13':('Hidalgo','Hidalgo'),
    '14':('Jalisco','Jalisco'),
    '15':('Estado_de_Mexico','Estado de Mexico'),
    '16':('Michoacan','Michoacan'),
    '17':('Morelos','Morelos'),
    '18':('Nayarit','Nayarit'),
    '19':('Nuevo_Leon','Nuevo Leon'),
    '20':('Oaxaca','Oaxaca'),
    '21':('Puebla','Puebla'),
    '22':('Queretaro','Queretaro'),
    '23':('Quintana_Roo','Quintana Roo'),
    '24':('San_Luis_Potosi','San Luis Potosi'),
    '25':('Sinaloa','Sinaloa'),
    '26':('Sonora','Sonora'),
    '27':('Tabasco','Tabasco'),
    '28':('Tamaulipas','Tamaulipas'),
    '29':('Tlaxcala','Tlaxcala'),
    '30':('Veracruz','Veracruz'),
    '31':('Yucatan','Yucatan'),
    '32':('Zacatecas','Zacatecas'),
    'COTS_REFUERZO':('COTS_Refuerzo','COTS Refuerzo'),
    'DISTRITALES':('Distritales','Distritales'),
    'FINANZAS':('Finanzas','Finanzas'),
}

ETIQUETA_MAP = {}
_base = {
    '01':['01_Aguascalientes','1_Aguascalientes'],
    '02':['02_Baja_California'],
    '03':['03_Baja_California_Sur'],
    '04':['04_Campeche','Estruct_Campeche'],
    '05':['05_Coahuila','Estruct_Coahuila','PROCESO_COAHUILA_ESTRUCT'],
    '06':['06_Colima'],
    '07':['07_Chiapas','Estruct_Chiapas'],
    '08':['08_Chihuahua','Estruct_Chihuahua'],
    '09':['09_CDMX','09_CDMX Offline','Estruct_CDMX','Estruct_CDMX_Distrito 12'],
    '10':['10_Durango','Estruct_Durango'],
    '11':['11_Guanajuato','Estruct_GTO','11_Estructura_GUANAJUATO'],
    '12':['12_Guerrero','Estruct_GRO'],
    '13':['13_Hidalgo','13_HGO Offline','Estruct_HGO'],
    '14':['14_Jalisco','14_JLO Offline','Estruct_Jalisco'],
    '15':['15_Edomex','_EDOMEX','Estruct_EDOMEX'],
    '16':['16_Michoacan','Estruct_Michoacan'],
    '17':['17_Morelos','Estruct_Morelos'],
    '18':['18_Nayarit','Estruct_Nayarit'],
    '19':['19_Nuevo_Leon','Estruct_NL'],
    '20':['20_Oaxaca','Oaxacaa','Estruct_Oax','Oaxaca_ROBADOS'],
    '21':['21_Puebla','ESTRUCTURA_PUEBLA','Estruct_Puebla'],
    '22':['22_Queretaro','Estruct_Queretaro'],
    '23':['23_Quintana Roo','Estruct_Quintana Roo'],
    '24':['24_San Luis Potosi','Estruct_SLP'],
    '25':['25_Sinaloa','Estruct_Sinaloa'],
    '26':['26_Sonora','Estruct_Sonora'],
    '27':['27_Tabasco','Estruct_Tabasco'],
    '28':['28_Tamaulipas','Estruct_Tamaulipas'],
    '29':['29_Tlaxcala'],
    '30':['30_Veracruz','Estruct_Veracruz'],
    '31':['31_Yucatan','Estruct_YUC'],
    '32':['32_Zacatecas','Estruct_ZACAT'],
}
SUFIJOS = ['_ROBADOS','_EXTRAVIADOS','_ROBADO','_EXTRAVIADO',' Offline','_Offline']
for num, ets in _base.items():
    for e in ets:
        ETIQUETA_MAP[e] = num
        for suf in SUFIJOS:
            ETIQUETA_MAP[e + suf] = num
ETIQUETA_MAP.update({
    'Estruct_CDMX_ROBADOS':'09','Estruct_EDOMEX_ROBADOS':'15',
    'Estruct_Durango_ROBADOS':'10','Estruct_Sinaloa_ROBADOS':'25',
    'Estruct_Sonora_ROBADOS':'26','Estruct_Coahuila_ROBADOS':'05',
    'Estruct_Campeche_ROBADOS':'04','Estruct_Puebla_ROBADOS':'21',
    'ESTRUCTURA_PUEBLA_ROBADOS':'21','Estruct_Queretaro_EXTRAVIADOS':'22',
    'Estruct_SLP_ROBADOS':'24','11_Guanajuato_ROBADOS':'11',
    '_EDOMEX_ROBADOS':'15',
})

# ── Parsear GPS ────────────────────────────────────────────────────
def parse_gps(val):
    if pd.isna(val) or str(val).strip() == '':
        return None, None
    m = re.match(r'([-\d.]+),\s*([-\d.]+)\s*\(([^)]+)\)', str(val).strip())
    if m:
        return float(m.group(1)), float(m.group(2))
    return None, None

# ── Clasificacion ──────────────────────────────────────────────────
def clasificar(grupo_raw, etiqueta_raw):
    g = str(grupo_raw).strip() if not pd.isna(grupo_raw) else ''
    e = str(etiqueta_raw).strip() if not pd.isna(etiqueta_raw) else ''
    tokens = [t.strip() for t in g.split(',') if t.strip()]

    if e.startswith('Fin_') or e == 'Finanzas':
        return 'FINANZAS', False

    estados_grupo = []
    for tok in tokens:
        m = re.match(r'^(\d{2})_(?:Estructura|Proceso)_', tok, re.I)
        if m and m.group(1).zfill(2) in ESTADOS:
            estados_grupo.append(m.group(1).zfill(2))

    tiene_cots   = any(re.match(r'^05_COTS-REFUERZO$', t, re.I) for t in tokens)
    tiene_brigad = any(re.match(r'^05_Estructura_BRIGADISTAS$', t, re.I) for t in tokens)
    tiene_dist   = any(t.upper() == 'DISTRITALES' for t in tokens)
    tiene_fin    = any(t.upper() == 'FINANZAS' for t in tokens)

    if estados_grupo:
        return estados_grupo[0], tiene_cots
    if tiene_brigad:
        return '05', False
    if tiene_cots:
        return 'COTS_REFUERZO', False
    if tiene_dist:
        return 'DISTRITALES', False
    if tiene_fin:
        return 'FINANZAS', False
    if e in ETIQUETA_MAP:
        return ETIQUETA_MAP[e], False
    if 'REF_PROCESO_RS' in e.upper() or 'PROCESO_COAHUILA' in e.upper():
        return '05', False

    return None, False

def get_incidencia(etiqueta, grupo):
    t = ((str(etiqueta) if not pd.isna(etiqueta) else '') + ' ' +
         (str(grupo) if not pd.isna(grupo) else '')).upper()
    if 'ROBADO' in t:     return 'ROBADO'
    if 'EXTRAVIADO' in t: return 'EXTRAVIADO'
    return ''

def get_estatus(visto_str, estado_mdm, inc):
    if inc in ('ROBADO', 'EXTRAVIADO'):
        return 'robado'
    mdm = str(estado_mdm).strip().lower()
    if 'licencia ha caducado' in mdm or mdm == 'suministro':
        return 'caducada'
    if 'no inscrito' in mdm:
        return 'offline'
    if not visto_str or visto_str == 'nan':
        return 'offline'
    m = re.match(r'^(\d+)([dhm])', visto_str)
    if m:
        n, u = int(m.group(1)), m.group(2)
        dias = n if u == 'd' else n/24 if u == 'h' else n/1440
        if dias <= 3:  return 'activo'
        if dias <= 60: return 'offline'
        return 'caducada'
    return 'activo'

# ── Leer Excel ────────────────────────────────────────────────────
def leer_excel():
    # Buscar el Excel en la carpeta data/
    excels = list(DATA_DIR.glob('*.xlsx')) + list(DATA_DIR.glob('*.xls'))
    if not excels:
        raise FileNotFoundError('No se encontro ningun archivo Excel en la carpeta data/')

    xlsx = sorted(excels)[-1]  # el mas reciente
    print('Leyendo Excel: ' + str(xlsx))
    df = pd.read_excel(xlsx)
    print('Total filas: ' + str(len(df)))
    print('Columnas: ' + str(df.columns.tolist()))
    return df

# ── Clasificar filas ───────────────────────────────────────────────
def procesar_df(df):
    # Detectar columnas por posicion (A=0, B=1, ... I=8)
    col_estado   = df.columns[0]   # A - Estado (estatus MDM)
    col_visto    = df.columns[1]   # B - Visto por ultima vez
    col_imei     = df.columns[2]   # C - IMEI/MEID
    col_serie    = df.columns[3]   # D - Numero de serie
    col_etiqueta = df.columns[4]   # E - Etiqueta del dispositivo
    col_tel      = df.columns[5]   # F - Numero de telefono
    col_grupo    = df.columns[6]   # G - Grupo asignado
    col_gps      = df.columns[7]   # H - Ultima ubicacion
    col_apps     = df.columns[8] if len(df.columns) > 8 else None  # I - Apps no instaladas

    print('Columnas detectadas:')
    print('  Estado MDM:  ' + str(col_estado))
    print('  Visto:       ' + str(col_visto))
    print('  IMEI:        ' + str(col_imei))
    print('  Grupo:       ' + str(col_grupo))
    print('  GPS:         ' + str(col_gps))

    grupos = {}
    excluidos = 0
    dupes = set()
    imei_count = df[col_imei].value_counts()
    imei_dupes = set(imei_count[imei_count > 1].index)

    for _, row in df.iterrows():
        etiq    = row[col_etiqueta]
        grupo   = row[col_grupo]
        inc     = get_incidencia(etiq, grupo)
        visto   = str(row[col_visto]).strip() if pd.notna(row[col_visto]) else ''
        estatus = get_estatus(visto, row[col_estado], inc)

        # IMEI limpio
        imei_raw = str(row[col_imei]).strip() if pd.notna(row[col_imei]) else ''
        try:
            imei = str(int(float(imei_raw))) if imei_raw and imei_raw != 'nan' else ''
        except Exception:
            imei = imei_raw

        is_dupe = (row[col_imei] in imei_dupes) if imei_dupes else False

        # GPS
        lat, lng = parse_gps(row[col_gps])

        # V52 pendiente (col I)
        v52_ni = False
        if col_apps is not None:
            apps_str = str(row[col_apps]).strip() if pd.notna(row[col_apps]) else ''
            v52_ni = 'Sumate_V52-PROD-V1' in apps_str

        rec = {
            'imei':    imei,
            'serie':   str(row[col_serie]).strip() if pd.notna(row[col_serie]) else '',
            'etiqueta':str(etiq).strip() if not pd.isna(etiq) else '',
            'tel':     str(row[col_tel]).replace("'","").lstrip('+').strip() if pd.notna(row[col_tel]) else '',
            'grupo':   str(grupo).strip() if not pd.isna(grupo) else '',
            'visto':   visto,
            'lat':     round(lat, 6) if lat is not None else None,
            'lng':     round(lng, 6) if lng is not None else None,
            'inc':     inc,
            'estatus': estatus,
            'v52_ni':  v52_ni,
            'dupe':    is_dupe,
        }

        bucket, tambien_cots = clasificar(grupo, etiq)
        if bucket is None:
            excluidos += 1
            continue

        grupos.setdefault(bucket, []).append(rec)
        if tambien_cots:
            grupos.setdefault('COTS_REFUERZO', []).append(rec)

    print('Excluidas (sin estado): ' + str(excluidos))
    return grupos

# ── Logo ───────────────────────────────────────────────────────────
def get_logo():
    logo_path = Path('scripts/logo.b64')
    if logo_path.exists():
        return logo_path.read_text().strip()
    return 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVQI12NgAAIABQAABjE+ibYAAAAASUVORK5CYII='

# ── HTML por estado ────────────────────────────────────────────────
def generar_html(bucket, nombre, devices, logo, fecha):
    lats  = [d['lat'] for d in devices if d['lat'] is not None]
    lngs  = [d['lng'] for d in devices if d['lng'] is not None]
    clat  = round(sum(lats)/len(lats), 4) if lats else 23.6
    clng  = round(sum(lngs)/len(lngs), 4) if lngs else -102.5
    czoom = 8 if lats else 5
    total   = len(devices)
    con_gps = len(lats)
    v52c    = sum(1 for d in devices if d.get('v52_ni'))
    rob     = sum(1 for d in devices if d['inc'] in ('ROBADO','EXTRAVIADO'))
    cad     = sum(1 for d in devices if d['estatus'] == 'caducada')
    dupes   = sum(1 for d in devices if d.get('dupe'))

    dj = json.dumps(devices, ensure_ascii=False, separators=(',', ':'))

    dupe_banner = ''
    if dupes > 0:
        dupe_banner = '<div style="background:#fff3cd;border-bottom:1px solid #ffc107;padding:5px 20px;font-size:11px;color:#856404;font-weight:600;">&#9888; ' + str(dupes) + ' tableta(s) con IMEI duplicado en este estado</div>'

    css = """:root{--vino:#6D1130;--vl:#f5e8ec;--bl:#fff;--gs:#f7f4f5;--gb:#e0d4d8;--gt:#5a4a50;
--ve:#1a7a45;--ro:#b83232;--am:#c97a00;--gr:#555;--na:#d45500;--v52:#d46800;--v52l:#fff3e0;}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
html,body{height:100%;font-family:'Barlow',sans-serif;background:var(--gs);color:#222;overflow:hidden;}
#hd{background:var(--bl);border-bottom:4px solid var(--vino);padding:0 20px;display:flex;align-items:center;gap:14px;height:66px;z-index:1000;box-shadow:0 2px 12px rgba(109,17,48,.10);}
#hd img{height:46px;width:auto;object-fit:contain;flex-shrink:0;border-radius:4px;}
#hd .tt h1{font-family:'Barlow Condensed',sans-serif;font-size:19px;font-weight:700;color:var(--vino);letter-spacing:.5px;text-transform:uppercase;}
#hd .tt p{font-size:11px;color:var(--gt);margin-top:2px;}
#hd .ac{margin-left:auto;text-align:right;font-size:11px;color:var(--gt);line-height:1.6;}
#hd .ac strong{color:var(--vino);font-size:12px;display:block;}
#st{background:var(--vino);display:flex;padding:0 20px;}
.sc{flex:1;padding:8px 10px;border-right:1px solid rgba(255,255,255,.15);display:flex;align-items:center;gap:8px;}
.sc:last-child{border-right:none;}
.sn{font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:700;color:#fff;line-height:1;}
.sn.v52a{color:#ffd080;}
.si{display:flex;flex-direction:column;}
.sl{font-size:9px;color:rgba(255,255,255,.7);text-transform:uppercase;letter-spacing:.5px;}
.ss{font-size:9px;color:rgba(255,255,255,.85);margin-top:1px;}
#tb{background:var(--bl);border-bottom:1px solid var(--gb);padding:7px 20px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;}
#tb input[type=text],#tb select{border:1px solid var(--gb);border-radius:6px;padding:5px 9px;font-size:12px;font-family:'Barlow',sans-serif;color:#333;background:var(--gs);outline:none;}
#tb input[type=text]{width:195px;}
#tb input:focus,#tb select:focus{border-color:var(--vino);}
.bv52{background:transparent;border:1.5px solid var(--v52);color:var(--v52);padding:5px 11px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;font-family:'Barlow',sans-serif;display:flex;align-items:center;gap:5px;}
.bv52.on{background:var(--v52);color:#fff;}
.sep{width:1px;height:24px;background:var(--gb);flex-shrink:0;}
.ie{margin-left:auto;font-size:11px;color:var(--gt);background:var(--vl);padding:4px 12px;border-radius:6px;border:1px solid var(--gb);font-weight:600;white-space:nowrap;}
#main{display:flex;height:calc(100vh - 66px - 44px - 43px);min-height:400px;}
#map{flex:1;min-width:0;position:relative;}
.leaflet-popup-content-wrapper{border-radius:8px!important;border:2px solid var(--vino)!important;padding:0!important;}
.leaflet-popup-content{margin:0!important;}
.leaflet-popup-tip{background:var(--vino)!important;}
.ley{position:absolute;bottom:14px;left:14px;z-index:900;background:#fff;border:1.5px solid var(--vino);border-radius:8px;padding:9px 13px;font-size:11px;}
.ley h4{font-family:'Barlow Condensed',sans-serif;color:var(--vino);font-size:12px;font-weight:700;text-transform:uppercase;margin-bottom:6px;}
.lr{display:flex;align-items:center;gap:7px;margin-bottom:3px;font-size:11px;}
.ld{width:10px;height:10px;border-radius:50%;flex-shrink:0;}
.lsep{height:1px;background:var(--gb);margin:5px 0;}
#pn{width:320px;flex-shrink:0;background:var(--bl);border-left:1px solid var(--gb);display:flex;flex-direction:column;overflow:hidden;}
#ph{padding:10px 14px;border-bottom:2px solid var(--vino);background:var(--vl);}
#ph h3{font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:700;color:var(--vino);text-transform:uppercase;letter-spacing:.5px;}
#ph p{font-size:10px;color:var(--gt);margin-top:2px;}
#lista{flex:1;overflow-y:auto;padding:7px;}
.card{border:1px solid var(--gb);border-radius:8px;padding:10px 11px;margin-bottom:6px;cursor:pointer;transition:border-color .15s,background .15s;background:var(--bl);border-left:4px solid var(--gb);}
.card:hover{border-color:var(--vino);background:var(--vl);}
.card.activo{border-left-color:var(--ve);}
.card.offline{border-left-color:var(--ro);}
.card.caducada{border-left-color:var(--gr);}
.card.robado{border-left-color:var(--na);}
.card.sg{opacity:.75;}
.card.v52p{box-shadow:0 0 0 1.5px var(--v52) inset;}
.card.dupew{outline:2px solid #c97a00;}
.card.sel{outline:2px solid var(--vino);}
.cd{display:inline-flex;align-items:center;font-size:11px;font-weight:700;padding:2px 9px;border-radius:20px;margin-bottom:5px;}
.cd.ok{background:#e6f4ec;color:var(--ve);}
.cd.warn{background:#fff4e0;color:var(--am);}
.cd.danger{background:#faeaea;color:var(--ro);}
.cd.venc{background:#eee;color:var(--gr);}
.cd.inc{background:#fde8d9;color:var(--na);}
.ct{display:flex;align-items:center;justify-content:space-between;margin-bottom:3px;gap:6px;}
.ce{font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:700;color:var(--vino);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:155px;}
.bgs{display:flex;gap:3px;flex-wrap:wrap;justify-content:flex-end;}
.bg{font-size:9px;font-weight:700;padding:2px 6px;border-radius:20px;text-transform:uppercase;white-space:nowrap;}
.bg.activo{background:#e6f4ec;color:var(--ve);}
.bg.offline{background:#faeaea;color:var(--ro);}
.bg.caducada{background:#eee;color:var(--gr);}
.bg.robado{background:#fde8d9;color:var(--na);}
.bg.v52{background:var(--v52l);color:var(--v52);}
.bg.ng{background:#eee;color:#888;}
.bg.dup{background:#fff3cd;color:#856404;}
.cdat{display:grid;grid-template-columns:58px 1fr;gap:2px 8px;font-size:10px;margin-top:5px;}
.cl{color:#bbb;font-size:9px;text-transform:uppercase;letter-spacing:.3px;padding-top:1px;}
.cv{font-family:'Source Code Pro',monospace;font-size:10px;color:#333;word-break:break-all;}
.cv.nl{font-family:'Barlow',sans-serif;}
.cv.v52t{font-family:'Barlow',sans-serif;color:var(--v52);font-weight:600;font-size:9px;}
#ft{background:var(--bl);border-top:3px solid var(--vino);padding:6px 20px;display:flex;align-items:center;justify-content:space-between;font-size:10px;color:var(--gt);}
#ft .marca{color:var(--vino);font-weight:700;font-size:11px;}
#ntf{position:fixed;bottom:20px;left:50%;transform:translateX(-50%);background:var(--vino);color:#fff;padding:9px 20px;border-radius:8px;font-size:12px;font-weight:600;z-index:9998;opacity:0;transition:opacity .3s;pointer-events:none;}
#ntf.on{opacity:1;}
@media(max-width:768px){#pn{display:none;}}"""

    js = """var D=""" + dj + """;
var mp,ly,rf={},sv=false,td=[];
var CO={activo:'#1a7a45',offline:'#b83232',caducada:'#555',robado:'#d45500'};
var LB={activo:'ACTIVO',offline:'OFFLINE',caducada:'LIC. CADUCADA',robado:'ROBADO/EXTRAVIADO'};
function mi(e,v52,dp){
  var c=CO[e]||'#6D1130';
  var rg=v52?'<circle cx="14" cy="4.5" r="4" fill="none" stroke="#d46800" stroke-width="2.5"/>':'';
  var drg=dp?'<circle cx="24" cy="5" r="4" fill="#c97a00"/>':'';
  var s='<svg xmlns="http://www.w3.org/2000/svg" width="28" height="36" viewBox="0 0 28 36">'
    +'<path d="M14 0C6.27 0 0 6.27 0 14c0 9.25 14 22 14 22S28 23.25 28 14C28 6.27 21.73 0 14 0z" fill="'+c+'"/>'
    +'<rect x="7" y="7" width="14" height="17" rx="2" fill="white"/>'
    +'<rect x="10" y="5.5" width="8" height="2" rx="1" fill="white"/>'
    +'<rect x="9" y="10" width="10" height="1.5" rx=".5" fill="'+c+'"/>'
    +'<rect x="9" y="13" width="10" height="1.5" rx=".5" fill="'+c+'"/>'
    +'<rect x="9" y="16" width="6" height="1.5" rx=".5" fill="'+c+'"/>'+rg+drg+'</svg>';
  return L.divIcon({html:s,iconSize:[28,36],iconAnchor:[14,36],popupAnchor:[0,-38],className:''});
}
function mpop(d){
  var c=CO[d.estatus]||'#6D1130';
  var ub=d.lat!==null?d.lat.toFixed(5)+', '+d.lng.toFixed(5):'Sin GPS';
  var v52r=d.v52_ni?'<tr><td class="pp">V52</td><td style="color:#d46800;font-weight:700;">&#9888; Pendiente</td></tr>':'';
  var ir=d.inc?'<tr><td class="pp">Incidencia</td><td style="color:#d45500;font-weight:700;">'+d.inc+'</td></tr>':'';
  var gr=d.grupo?'<tr><td class="pp">Grupo</td><td style="font-size:9px;color:#555;">'+d.grupo+'</td></tr>':'';
  var dr=d.dupe?'<tr><td class="pp">Aviso</td><td style="color:#856404;font-weight:700;">IMEI duplicado</td></tr>':'';
  return '<div style="font-family:Barlow,sans-serif;min-width:240px;border-radius:6px;overflow:hidden;">'
    +'<div style="background:'+c+';color:#fff;padding:6px 12px;font-size:10px;font-weight:700;">&bull; '+(LB[d.estatus]||d.estatus.toUpperCase())+' &mdash; MDM</div>'
    +'<div style="padding:10px 12px;">'
    +'<div style="font-family:Barlow Condensed,sans-serif;font-size:15px;font-weight:700;color:#6D1130;margin-bottom:6px;">'+(d.etiqueta||d.serie)+'</div>'
    +'<style>.pp{color:#bbb;font-size:9px;text-transform:uppercase;padding:2px 0;width:65px;}</style>'
    +'<table style="font-size:10px;width:100%;border-collapse:collapse;">'
    +'<tr><td class="pp">IMEI</td><td style="font-family:monospace;color:#333;">'+d.imei+'</td></tr>'
    +'<tr><td class="pp">N&deg; Serie</td><td style="font-family:monospace;color:#555;font-size:9px;">'+d.serie+'</td></tr>'
    +'<tr><td class="pp">Tel&eacute;fono</td><td style="color:#333;">'+(d.tel||'&mdash;')+'</td></tr>'
    +'<tr><td class="pp">Visto</td><td style="color:#333;">'+(d.visto||'&mdash;')+'</td></tr>'
    +'<tr><td class="pp">Ubicaci&oacute;n</td><td style="font-family:monospace;color:#555;font-size:9px;">'+ub+'</td></tr>'
    +gr+v52r+ir+dr+'</table></div></div>';
}
function dc(d){
  if(d.estatus==='caducada') return 'venc';
  if(d.estatus==='robado')   return 'inc';
  var m=(d.visto||'').match(/^(\\d+)([dhm])/);
  if(!m) return 'ok';
  var n=parseInt(m[1]),u=m[2],di=u==='d'?n:u==='h'?n/24:n/1440;
  return di===0?'ok':di<=7?'warn':'danger';
}
function mcard(d,i){
  var tg=d.lat!==null;
  var ub=tg?d.lat.toFixed(5)+', '+d.lng.toFixed(5):'Sin GPS';
  var dt='';
  if(d.estatus==='caducada') dt='Licencia caducada'+(d.visto?' &middot; '+d.visto:'');
  else if(d.estatus==='robado') dt=(d.inc||'Incidencia')+(d.visto?' &middot; '+d.visto:'');
  else dt=d.visto||'Sin datos';
  var v52b=d.v52_ni?'<span class="bg v52">V52 &#9888;</span>':'';
  var gpb=!tg?'<span class="bg ng">Sin GPS</span>':'';
  var dpb=d.dupe?'<span class="bg dup">IMEI Dup.</span>':'';
  var v52r=d.v52_ni?'<span class="cl">Sumate V52</span><span class="cv v52t">&#9888; Pendiente</span>':'';
  var ir=d.inc?'<span class="cl">Incidencia</span><span class="cv nl" style="color:#d45500;font-weight:700;">'+d.inc+'</span>':'';
  var grr=d.grupo?'<span class="cl">Grupo</span><span class="cv nl" style="font-size:9px;color:#777;">'+(d.grupo.length>42?d.grupo.substring(0,40)+'&hellip;':d.grupo)+'</span>':'';
  var elbl=d.estatus==='caducada'?'Caducada':d.estatus==='robado'?(d.inc||'Robado'):d.estatus;
  return '<div class="card '+d.estatus+(!tg?' sg':'')+(d.v52_ni?' v52p':'')+(d.dupe?' dupew':'')+'" data-i="'+i+'" onclick="ir('+i+')">'
    +'<div class="ct"><span class="ce" title="'+(d.etiqueta||d.serie)+'">'+(d.etiqueta||d.serie)+'</span>'
    +'<div class="bgs">'+v52b+dpb+'<span class="bg '+d.estatus+'">'+elbl+'</span>'+gpb+'</div></div>'
    +'<div class="cd '+dc(d)+'">'+dt+'</div>'
    +'<div class="cdat"><span class="cl">IMEI</span><span class="cv">'+d.imei+'</span>'
    +'<span class="cl">Serie</span><span class="cv">'+d.serie+'</span>'
    +'<span class="cl">GPS</span><span class="cv nl">'+ub+'</span>'
    +grr+v52r+ir+'</div></div>';
}
function rend(lista){
  ly.clearLayers();rf={};
  var el=document.getElementById('lista');el.innerHTML='';
  for(var i=0;i<lista.length;i++){
    var d=lista[i];
    if(d.lat!==null){
      var mk=L.marker([d.lat,d.lng],{icon:mi(d.estatus,d.v52_ni,d.dupe)});
      mk.bindPopup(mpop(d),{maxWidth:290});mk.addTo(ly);rf[i]={mk:mk,d:d};
    }
    el.insertAdjacentHTML('beforeend',mcard(d,i));
  }
  var t=lista.length,g=0,v52n=0,inc=0,cad=0;
  for(var j=0;j<lista.length;j++){
    if(lista[j].lat!==null) g++;
    if(lista[j].v52_ni) v52n++;
    if(lista[j].inc) inc++;
    if(lista[j].estatus==='caducada') cad++;
  }
  document.getElementById('s0').textContent=t;
  document.getElementById('s1').textContent=g;
  document.getElementById('s2').textContent=t-g;
  document.getElementById('s3').textContent=v52n;
  document.getElementById('s4').textContent=inc;
  document.getElementById('s5').textContent=cad;
  document.getElementById('pc').textContent=t+' tabletas \u00b7 clic para centrar';
  document.getElementById('fc').textContent=t+' de '+td.length+' tabletas';
}
function ir(i){
  document.querySelectorAll('.card').forEach(function(c){c.classList.remove('sel');});
  var card=document.querySelector('.card[data-i="'+i+'"]');
  if(card){card.classList.add('sel');card.scrollIntoView({behavior:'smooth',block:'nearest'});}
  var r=rf[i];if(!r){ntf('Sin GPS registrado');return;}
  mp.setView([r.d.lat,r.d.lng],15);r.mk.openPopup();
}
function tv(){
  sv=!sv;
  document.getElementById('bv').classList.toggle('on',sv);
  fil();
}
function fil(){
  var q=document.getElementById('bq').value.toLowerCase();
  var fe=document.getElementById('fe').value;
  var fg=document.getElementById('fg').value;
  var l=[];
  for(var i=0;i<td.length;i++){
    var d=td[i];
    var mq=!q||[d.imei,d.serie,d.etiqueta,d.tel,d.grupo].some(function(v){return (v||'').toLowerCase().indexOf(q)>=0;});
    var me=!fe||d.estatus===fe;
    var mg=!fg||(fg==='con'&&d.lat!==null)||(fg==='sin'&&d.lat===null);
    var mv=!sv||d.v52_ni;
    if(mq&&me&&mg&&mv) l.push(d);
  }
  rend(l);
}
function ntf(msg){var el=document.getElementById('ntf');el.textContent=msg;el.classList.add('on');setTimeout(function(){el.classList.remove('on');},3000);}
window.addEventListener('load',function(){
  mp=L.map('map',{zoomControl:true}).setView([""" + str(clat) + """,""" + str(clng) + """],""" + str(czoom) + """);
  L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png',
    {attribution:'\u00a9 OpenStreetMap \u00a9 CARTO',maxZoom:19,subdomains:'abcd'}).addTo(mp);
  ly=L.layerGroup().addTo(mp);
  fetch('https://raw.githubusercontent.com/angelnmara/geojson/master/mexicoHigh.json')
    .then(function(r){return r.json();})
    .then(function(data){L.geoJSON(data,{
      style:{color:'#6D1130',weight:1.5,fillColor:'#f9f0f3',fillOpacity:.18},
      onEachFeature:function(f,l){if(f.properties&&f.properties.name)l.bindTooltip(f.properties.name,{permanent:false,direction:'center'});}
    }).addTo(mp);}).catch(function(){});
  td=D;rend(td);
  var cg=[];for(var i=0;i<td.length;i++){if(td[i].lat!==null)cg.push([td[i].lat,td[i].lng]);}
  if(cg.length>0)mp.fitBounds(cg,{padding:[40,40],maxZoom:11});
});"""

    return """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>MORENA - Estructura - """ + nombre + """</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css"/>
<link href="https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700;800&family=Source+Code+Pro:wght@400;500&display=swap" rel="stylesheet"/>
<script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js"></script>
<style>""" + css + """</style>
</head>
<body>
<div id="hd">
  <img src=\"""" + logo + """\" alt="MORENA"/>
  <div class="tt"><h1>MORENA - Estructura - """ + nombre + """</h1><p>Monitoreo de tabletas - """ + fecha + """</p></div>
  <div class="ac"><strong>""" + fecha + """</strong>Actualizacion automatica</div>
</div>
""" + dupe_banner + """
<div id="st">
  <div class="sc"><div class="sn" id="s0">0</div><div class="si"><span class="sl">Total</span><span class="ss">Tabletas</span></div></div>
  <div class="sc"><div class="sn" id="s1">0</div><div class="si"><span class="sl">Con GPS</span><span class="ss">Ubicadas</span></div></div>
  <div class="sc"><div class="sn" id="s2">0</div><div class="si"><span class="sl">Sin GPS</span><span class="ss">Solo lista</span></div></div>
  <div class="sc"><div class="sn v52a" id="s3">0</div><div class="si"><span class="sl">V52 Pendiente</span><span class="ss">Sin instalar</span></div></div>
  <div class="sc"><div class="sn" id="s4">0</div><div class="si"><span class="sl">Incidencias</span><span class="ss">Robado/Extraviado</span></div></div>
  <div class="sc"><div class="sn" id="s5">0</div><div class="si"><span class="sl">Lic. Caducada</span><span class="ss">Sin renovar</span></div></div>
</div>
<div id="tb">
  <input type="text" id="bq" placeholder="Buscar IMEI, serie, etiqueta..." oninput="fil()"/>
  <select id="fe" onchange="fil()">
    <option value="">Todos los estatus</option>
    <option value="activo">Activo</option>
    <option value="offline">Offline</option>
    <option value="caducada">Lic. Caducada</option>
    <option value="robado">Robado / Extraviado</option>
  </select>
  <select id="fg" onchange="fil()">
    <option value="">Con y sin GPS</option>
    <option value="con">Solo con GPS</option>
    <option value="sin">Solo sin GPS</option>
  </select>
  <div class="sep"></div>
  <button class="bv52" id="bv" onclick="tv()">
    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>
    Solo V52 Pendiente
  </button>
  <span class="ie">""" + nombre + """ - """ + str(total) + """ tab - """ + str(con_gps) + """ GPS - """ + str(v52c) + """ V52 pend.</span>
</div>
<div id="main">
  <div id="map">
    <div class="ley">
      <h4>Leyenda</h4>
      <div class="lr"><div class="ld" style="background:#1a7a45"></div><span>Activo (0-3 dias)</span></div>
      <div class="lr"><div class="ld" style="background:#b83232"></div><span>Offline (4+ dias)</span></div>
      <div class="lr"><div class="ld" style="background:#555"></div><span>Lic. Caducada</span></div>
      <div class="lr"><div class="ld" style="background:#d45500"></div><span>Robado / Extraviado</span></div>
      <div class="lsep"></div>
      <div class="lr"><div style="width:10px;height:10px;border-radius:50%;border:2.5px solid #d46800;flex-shrink:0;"></div><span style="color:#d46800;font-weight:600;">V52 Pendiente</span></div>
    </div>
  </div>
  <div id="pn">
    <div id="ph"><h3>""" + nombre + """</h3><p id="pc">""" + str(total) + """ tabletas</p></div>
    <div id="lista"></div>
  </div>
</div>
<div id="ft">
  <span>Monitoreo MDM - Tenant Estructura - """ + nombre + """</span>
  <span class="marca">MORENA -</span>
  <span id="fc">""" + str(total) + """ tabletas</span>
</div>
<div id="ntf"></div>
<script>
""" + js + """
</script>
</body>
</html>"""


# ── Index ──────────────────────────────────────────────────────────
def generar_index(grupos, fecha):
    filas = ''
    orden = [str(n).zfill(2) for n in range(1, 33)] + ['COTS_REFUERZO', 'DISTRITALES', 'FINANZAS']
    total_gral = 0
    for b in orden:
        if b not in grupos:
            continue
        key, nombre = ESTADOS[b]
        total   = len(grupos[b])
        con_gps = sum(1 for d in grupos[b] if d['lat'] is not None)
        v52c    = sum(1 for d in grupos[b] if d.get('v52_ni'))
        rob     = sum(1 for d in grupos[b] if d['inc'] in ('ROBADO', 'EXTRAVIADO'))
        if b not in ('COTS_REFUERZO', 'DISTRITALES', 'FINANZAS'):
            total_gral += total
        filas += ('<tr><td>' + b + '</td>'
                  '<td><a href="' + b + '_' + key + '.html">' + nombre + '</a></td>'
                  '<td>' + str(total) + '</td>'
                  '<td>' + str(con_gps) + '</td>'
                  '<td>' + str(v52c) + '</td>'
                  '<td>' + str(rob) + '</td></tr>\n')

    return """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"/>
<title>MORENA - Tenant Estructura - Indice de Mapas</title>
<link href="https://fonts.googleapis.com/css2?family=Barlow:wght@400;600;700&family=Barlow+Condensed:wght@700&display=swap" rel="stylesheet"/>
<style>
body{font-family:'Barlow',sans-serif;background:#f7f4f5;color:#222;margin:0;padding:20px;}
h1{font-family:'Barlow Condensed',sans-serif;color:#6D1130;font-size:28px;margin-bottom:4px;}
p{color:#5a4a50;font-size:13px;margin-bottom:20px;}
table{width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.08);}
th{background:#6D1130;color:#fff;padding:10px 14px;text-align:left;font-size:12px;text-transform:uppercase;letter-spacing:.5px;}
td{padding:9px 14px;border-bottom:1px solid #e0d4d8;font-size:13px;}
tr:last-child td{border-bottom:none;}
a{color:#6D1130;font-weight:600;text-decoration:none;}
a:hover{text-decoration:underline;}
</style>
</head>
<body>
<h1>MORENA - Tenant Estructura - Indice de Mapas</h1>
<p>Ultima actualizacion: """ + fecha + """</p>
<table>
<tr><th>#</th><th>Entidad</th><th>Total</th><th>Con GPS</th><th>V52 Pend.</th><th>Incidencias</th></tr>
""" + filas + """</table>
</body>
</html>"""


# ── Main ───────────────────────────────────────────────────────────
if __name__ == '__main__':
    OUTPUT_DIR.mkdir(exist_ok=True)
    logo = get_logo()

    df     = leer_excel()
    grupos = procesar_df(df)

    print('Generando HTML...')
    orden = [str(n).zfill(2) for n in range(1, 33)] + ['COTS_REFUERZO', 'DISTRITALES', 'FINANZAS']
    for bucket in orden:
        if bucket not in grupos:
            continue
        key, nombre = ESTADOS[bucket]
        devs     = grupos[bucket]
        html     = generar_html(bucket, nombre, devs, logo, FECHA_HOY)
        out_file = OUTPUT_DIR / (bucket + '_' + key + '.html')
        out_file.write_text(html, encoding='utf-8')
        t = len(devs)
        g = sum(1 for d in devs if d['lat'] is not None)
        print('  OK ' + bucket + ' ' + nombre + ': ' + str(t) + ' tab, ' + str(g) + ' GPS')

    index_html = generar_index(grupos, FECHA_HOY)
    (OUTPUT_DIR / 'index.html').write_text(index_html, encoding='utf-8')
    print('Completado: ' + str(len(grupos)) + ' archivos generados en docs/')
