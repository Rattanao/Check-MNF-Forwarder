#!/usr/bin/env python
"""ตรวจ EDI: MANIFEST × ENTER (ไฟล์ ENTER ชุดรวมหลาย forwarder)

ไฟล์ ENTER เป็น PDF ชุดรวม: ARRIVAL NOTICE/สำเนา B/L ของ Heung-A (+ ATTACHED RIDER),
ใบ MARK&NO./CNTR/B/L NO., ตาราง AMENDMENT (Cargoport / ASIA ATLANTIC / INITIAL / Handle ฯลฯ),
จดหมายแก้ไขแบบข้อความอิสระ, หน้าภาพสแกน และฟอร์มของ forwarder ที่ skill มีตัวอ่านอยู่แล้ว
(MAC-NELS, THE M CORP/CMS, DSV)

สคริปต์นี้แยกหน้า ENTER ตามชนิดฟอร์ม แล้วอ่านแต่ละชนิดด้วยตัวอ่านของตัวเอง — ฟอร์มที่ skill
manifest-enter-edi-check อ่านได้อยู่แล้วจะส่งเฉพาะหน้านั้นให้ build_edi.py ของ skill อ่าน (เรียกใช้
อย่างเดียว ไม่แก้ไขไฟล์ของ skill) จากนั้นใช้ parse_manifest()/build_report() ของ skill สร้าง EDI.xlsx
รูปแบบเดิม แล้วเติมเลขหน้า ENTER ของแต่ละ B/L และหมายเหตุหน้าภาพสแกนลงคอลัมน์หมายเหตุ

ใช้:  python bundle_edi.py                      (หาไฟล์ใน input/ ให้อัตโนมัติ, ผลอยู่ใน output/EDI.xlsx)
      python bundle_edi.py --manifest input/MANIFEST.xls --enter input/ENTER.pdf --shed 0141
"""
import argparse
import glob
import importlib.util
import os
import re
import sys
import tempfile
from collections import Counter

import openpyxl
import pymupdf
from openpyxl.styles import PatternFill

# build_edi.py: ใช้สำเนาที่วางข้างไฟล์นี้ก่อน (ชุดที่แจกเพื่อน) ไม่มีจึงใช้ของ skill manifest-enter-edi-check
_HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_SCRIPT = next((p for p in (os.path.join(_HERE, 'build_edi.py'),
                                 os.path.join(os.path.expanduser('~'), '.claude', 'skills', 'manifest-enter-edi-check',
                                              'scripts', 'build_edi.py')) if os.path.exists(p)), '')


def _load_shared():
    if not SKILL_SCRIPT:
        sys.exit('!! ไม่พบ build_edi.py (วางไว้โฟลเดอร์เดียวกับ bundle_edi.py)')
    spec = importlib.util.spec_from_file_location('build_edi_shared', SKILL_SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


be = _load_shared()

# เลขตู้ ISO 6346: ตัวที่ 4 เป็น U/J/Z เสมอ — กันเลข SHIPMENT ID ('SSGN0000001') และส่วนท้ายของเลข B/L
CONT2_RE = re.compile(r'(?<![A-Z])[A-Z]{3}[UJZ]\d{7}')
BL_TXT_RE = re.compile(r'HASL[A-Z]\d{8,}[A-Z]?')
BL_IN_TXT_RE = re.compile(r'HASL[A-Z]\d{8,11}[A-Z]?(?![A-Za-z0-9])')   # ในข้อความที่มีช่องว่าง
STATUS_KW = re.compile(r'STATUS|ลากตู้|ตู้ลาก|ลากตต|เปิดตู้|ขนส่ง|\bCFS\b', re.I)
PKG_WORD_RE = re.compile(r'^(?:WOODEN\s+)?(PALLETS?|PLTS?|CARTONS?|CTNS?|CASES?|BOX(?:ES)?|BOX\(ES\)|PACKAGES?|PKGS?|IBCS?|PAILS?|'
                         r'SETS?|ROLLS?|REELS?|DRUMS?|BAGS?|BALES?|CRATES?|UNITS?|PIECES?|BUNDLES?|SKIDS?)\b', re.I)
REEFER_RE = re.compile(r'(?:MINUS|-)\s*(\d{1,2}(?:\.\d)?)\s*(?:DEG|DEGREES|°|\'\s*C)', re.I)


# ═══════════════ ตัวช่วยอ่านหน้า PDF ═══════════════
def page_rot(pg):
    """ทิศข้อความส่วนใหญ่ของหน้า — (0,-1)/(0,1) = เนื้อหาถูกวาดตะแคง 90° ทั้งที่ page.rotation เป็น 0
       (พบในฟอร์ม UNION WORLD ของ ENTER-SWSP) ต้องหมุนพิกัดกลับก่อนอ่านคอลัมน์ แบบ build_edi._page_lines"""
    dirs = Counter(tuple(round(v) for v in l['dir'])
                   for b in pg.get_text('dict')['blocks'] for l in b.get('lines', []))
    return dirs.most_common(1)[0][0] if dirs else (1, 0)


def _unrotate(pg, rot, x0, y0, x1, y1):
    """คืน bbox (x0, y0, x1, y1) ในแนวอ่านปกติ"""
    W, H = pg.rect.width, pg.rect.height
    if rot == (0, -1):      # ข้อความวิ่งจากล่างขึ้นบน → หมุนตามเข็ม 90°
        return H - y1, x0, H - y0, x1
    if rot == (0, 1):       # ข้อความวิ่งจากบนลงล่าง → หมุนทวนเข็ม 90°
        return y0, W - x1, y1, W - x0
    return x0, y0, x1, y1


def page_lines(pg):
    rot = page_rot(pg)
    out = []
    for b in pg.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            # '。' (U+3002) ใช้แทนช่องว่างใน PDF บางไฟล์ ('SHIPPER'S。LOAD') — ถือเป็นช่องว่าง
            t = ''.join(s['text'] for s in l['spans']).replace('。', ' ').strip()
            # ชั้นข้อความ OCR ของเครื่องสแกนอ่าน ' เป็น I ('SHIPPERIS LOAD') — ป้ายสำคัญของ ARRIVAL NOTICE
            t = re.sub(r"\bSHIPPER\s*I\s*S\s+LOAD\b", "SHIPPER'S LOAD", t)
            if t:
                x0, y0, x1, _ = _unrotate(pg, rot, *l['bbox'])
                out.append({'x': x0, 'y': y0, 'x1': x1, 't': t})
    return out


def page_annots(pg):
    rot = page_rot(pg)
    out = []
    for a in pg.annots() or []:
        if a.type[1] != 'FreeText':
            continue
        txt = a.info.get('content', '') or ''
        parts = [re.sub(r'\s+', ' ', p).strip() for p in re.split(r'[\r\n]+', txt)]
        parts = [p for p in parts if p]
        if parts:
            r = a.rect if rot not in ((0, -1), (0, 1)) else pymupdf.Rect(_unrotate(pg, rot, *a.rect))
            out.append({'rect': r, 'lines': parts, 't': ' '.join(parts)})
    return out


def nospace(s):
    return re.sub(r'\s+', '', s).upper()


_FORM_LABELS = {'SHIPPER', 'CONSIGNEE', 'NOTIFYPARTY'}


def in_annot(l, annots):
    """บรรทัดนี้เป็นข้อความที่ FreeText annotation วาดไว้บนหน้า (อยู่ในกรอบของมัน)
       ป้ายของฟอร์มเอง (Shipper / Consignee / Notify Party) ไม่ใช่ข้อความที่ลูกค้าพิมพ์ทับ แม้กรอบสีแดงจะชิดป้าย"""
    if nospace(l['t']) in _FORM_LABELS:
        return False
    # บรรทัดของฟอร์มที่กรอบ annotation บังบางส่วน แต่ไม่ใช่ข้อความของ annotation (RICH FREIGHT 'CNEE : …'
    # อยู่ใต้กรอบ CTN/STATUS ที่ลูกค้าแปะ, SWDN 2026-11-27) — ไม่นับเป็นข้อความ annotation
    lt = nospace(l['t'])
    return any(a['rect'].x0 - 3 <= l['x'] <= a['rect'].x1 and a['rect'].y0 - 3 <= l['y'] <= a['rect'].y1 + 1
               and (l['x'] >= a['rect'].x0 - 1 or len(lt) < 6 or lt[:6] in nospace(a['t']))
               for a in annots)


def conts_in(*txts):
    seen = []
    for t in txts:
        for c in CONT2_RE.findall(str(t or '').upper()):
            if c not in seen:
                seen.append(c)
    return seen


def split_pkg(s):
    """'38 BUNDLES' / '600BAGS' / '1 PALLET(S)' / '222 PLTS' -> (จำนวน, ชนิด)"""
    m = re.match(r'^\s*(\d[\d,]*)\s*([A-Za-z].*)?$', s or '')
    if not m:
        return None, ''
    return int(m.group(1).replace(',', '')), (m.group(2) or '').strip()


# ฟอนต์ไทยบางไฟล์ (UNION WORLD/AGN ที่วาดตะแคง) ถอดสระบน/ล่าง+วรรณยุกต์ออกมาเป็นอักษร ASCII คนละตัวในแต่ละไฟล์
# ('เร<งเป?ดตBCเขCาโกดHงด<วน', 'เร>งเปAดตDEเขEาโกดJงด>วน', 'ลากต?@') — จับกับวลี STATUS ที่รู้จัก
# โดยให้สระ/วรรณยุกต์แต่ละตัวแทนด้วยอักษรที่ไม่ใช่ไทยตัวใดก็ได้ 1 ตัว แล้วคืนวลีที่ถูกต้อง
_TH_MARKS = 'ัิีึืุู่้๊๋็์'
_TH_PHRASES = ['เร่งเปิดตู้เข้าโกดังด่วน', 'เปิดตู้เข้าโกดังด่วน', 'เปิดตู้เข้าโกดัง', 'เปิดตู้',
               'ขนส่งหน้าตู้แรงงานเรือ', 'ขนส่งแรงงานเรือ', 'ขนส่งหน้าตู้', 'ลากตู้ด่วน', 'ลากตู้']
_TH_GARBLE = [(re.compile(''.join(r'[!-~]' if ch in _TH_MARKS else re.escape(ch) for ch in ph)), ph)
              for ph in _TH_PHRASES]


# ฟอนต์ที่ถอดไทยออกมาเป็นอักษรซีริลลิกทั้งคำ (UNICORN: '***јѥдшѬҖ***' = ลากตู้) — จดเป็นคำ ๆ
_TH_LITERAL = {'јѥдшѬҖ': 'ลากตู้'}


def fix_thai_garble(s):
    for k, v in _TH_LITERAL.items():
        s = (s or '').replace(k, v)
    if not s or not re.search(r'[฀-๿][!-~]|[!-~][฀-๿]', s):
        return s
    for rx, ph in _TH_GARBLE:
        s = rx.sub(ph, s)
    return s


def clean_status(parts):
    out = []
    for p in parts:
        p = fix_thai_garble(p.replace('ลากตต', 'ลากตู้').replace('ตู้ลาก', 'ลากตู้'))   # 'ตู้ลากค่ะ' = ลากตู้
        p = re.sub(r'^\s*(?:STATUS|REMARKS?)\s*\.?\s*[:=]?\s*', '', p, flags=re.I).strip(' *"“”:')   # 'Status.: CY-CY'
        p = re.sub(r'^[=>\-\s]+', '', p)                                    # 'STATUS ===> ลากตู้ ด่วนมากค่ะ'
        if p and p not in out:
            out.append(p)
    return ' '.join(out)


def is_status_line(t):
    return bool(STATUS_KW.search(t))


def find_reefer(*txts):
    for t in txts:                      # 'MINUS 18 DEGREES' / 'TEMP : - 18'C' ก่อน — find_temp อ่านแบบมีช่องว่างไม่ออก
        m = REEFER_RE.search(t or '')
        if m:
            return f'-{m.group(1)}°C'
    for t in txts:                      # 'TEMP: 6 DEGREE C' (find_temp อ่าน DEGREE C ไม่ออก คืนทั้งก้อน)
        m = re.search(r'TEMP\w*\s*[:=]?\s*(\+?\s*\d{1,2}(?:\.\d)?)\s*DEG(?:REES?)?\.?\s*C\b', t or '', re.I)
        if m:
            return m.group(1).replace(' ', '').lstrip('+') + '°C'
    return be.find_temp(*txts)


def company_name(lines):
    """ชื่อบริษัทจากบล็อกชื่อ+ที่อยู่: เอาบรรทัดจนเจอคำลงท้ายบริษัท (ไม่เกิน 3 บรรทัด)"""
    lines = [l for l in lines if l.strip()]
    if not lines:
        return ''
    out = []
    for l in lines[:3]:
        out.append(l.strip())
        if re.search(r'\b(LTD|LIMITED|CORP(?:ORATION)?|INC|PLC|LLC|CO\.?,?\s*LTD)\b\.?\s*\*?$', l.strip(), re.I) \
                or re.search(r'\b(LTD|LIMITED)\b', l, re.I):
            break
    else:
        out = [lines[0].strip()]
    s = ' '.join(out)
    s = re.sub(r'\s+', ' ', s.replace(' D.', 'D.')).strip(' *')
    return re.sub(r'LT D\.', 'LTD.', s)


def blank(bl, page):
    return {'bl': bl, 'cons': '', 'cont': '', 'pkgs': None, 'pkgtype': '', 'gw': None, 'meas': None,
            'marks': '', 'desc': '', 'status_raw': '', 'transit_raw': '', 'dg': '', 'reefer': '',
            'low_confidence': False, 'pages': [page]}


def finish(rec, *extra_txt):
    blob = ' '.join([rec['marks'], rec['desc']] + [t for t in extra_txt if t])
    rec['dg'] = rec['dg'] or be.find_dg(blob)
    rec['reefer'] = rec['reefer'] or find_reefer(blob)
    if not rec['transit_raw'] and be.TRANSIT_RE.search(blob):
        rec['transit_raw'] = blob
    return rec


# ═══════════════ จำแนกชนิดหน้า ═══════════════
# ลายเซ็นหน้าของฟอร์ม forwarder ที่ build_edi.py มีตัวอ่านเฉพาะ (ตามตัวตรวจฟอร์มของเว็บแต่ละเจ้า)
# test(ns, lines): ns = ข้อความทั้งหน้าแบบไม่มีช่องว่าง ตัวพิมพ์ใหญ่ ; ลำดับสำคัญ (ตัวที่เจาะจงกว่ามาก่อน)
_has_line = lambda lines, rx: any(re.match(rx, l['t'], re.I) for l in lines)
FORWARDER_PAGES = [
    ('SHARED_SINOTRANS', lambda ns, ls: 'MB/LNO' in ns and 'ISSUEDD/OASFOLLOWING' in ns and 'MARKS&NOS' in ns),   # layout L
    ('SHARED_LXPANTOS', lambda ns, ls: 'LXPANTOS' in ns and 'AMENDMENTOFMANIFESTOF' in ns),                         # layout J
    ('SHARED_FREIGHTLINKS', lambda ns, ls: 'IMPORTCARGOMANIFEST' in ns and 'CUSTOMSB/LNO' in ns),                   # layout M
    ('SHARED_SHIPCO', lambda ns, ls: 'CUSTOMSIMPORTMANIFEST' in ns and 'ACTUALCONSIGNEE' in ns),                     # layout E
    ('SHARED_PLANET', lambda ns, ls: 'B/LCHANGENO' in ns and 'HB/LNO' in ns and _has_line(ls, r'^CNTR\s*NO\.?\s*:')),   # layout K
    ('SHARED_UNIONWORLD', lambda ns, ls: 'B/LCHANGENO' in ns and 'HB/LNO' in ns and _has_line(ls, r'^STATUS\s*=')),    # layout I
    ('SHARED_AGN', lambda ns, ls: 'B/LCHANGENO' in ns and 'HB/LNO' not in ns
                                  and 'UNICORNGLOBALLOGISTICS' not in ns),   # UNICORN = TABLE + โปรไฟล์ (SWDN 2026-11-27)                                     # layout D (AGN / Handle Inter)
    ('SHARED_LAYOUTB', lambda ns, ls: 'MARKS&NOS' in ns and 'DESCRIPTIONSOFGOODS' in ns and 'GROSSWT.&MEAS' in ns),  # layout B (Able / Quality)
]
# หน้าต่อเนื่อง/ใบปะหน้าของฟอร์มเดียวกัน (ลายเซ็นพิมพ์แค่หน้าแรก เช่นหน้า 2… ของ Able/Quality, ใบปะหน้า Freight Links)
# — หน้าที่ปกติจะเป็น LETTER แต่อยู่ติดกับหน้าของเจ้านั้นและผ่านเงื่อนไขนี้ ให้รวมกลุ่มเดียวกัน
FORWARDER_MEMBER = {
    'SHARED_LAYOUTB': lambda ns: 'CONSIGNEE' in ns and 'B/LNO' in ns and ('CONT.NO' in ns or 'CONTNO' in ns),
    'SHARED_FREIGHTLINKS': lambda ns: 'FREIGHTLINKS' in ns or 'IMPORTCARGOMANIFEST' in ns,
    'SHARED_SHIPCO': lambda ns: 'SHIPCO' in ns,
    'SHARED_SINOTRANS': lambda ns: 'SINOTRANS' in ns,
    'SHARED_LXPANTOS': lambda ns: 'LXPANTOS' in ns,
}


# ฟอร์มที่พิมพ์ลายเซ็น (หัวตาราง) เฉพาะหน้าแรกของเอกสาร — หน้าที่มีลายเซ็นเองเริ่มเอกสารใหม่เสมอ แม้หน้าก่อนเป็นฟอร์มเดียวกัน
# (ชุดรวมที่มี ENTER ของ Able ต่อด้วย Quality: ห้ามอ่านรวมเป็นเอกสารเดียว ไม่งั้น B/L แรกของ Quality ได้ข้อความท้ายหน้าของ Able)
DOC_HEADER_KINDS = {'SHARED_LAYOUTB'}


def join_forwarder_pages(kinds, L):
    """LETTER ที่อยู่ติดกลุ่มหน้าของ forwarder และเป็นส่วนหนึ่งของฟอร์มนั้น → ชนิดเดียวกับกลุ่ม (ทำซ้ำจนนิ่ง)"""
    changed = True
    while changed:
        changed = False
        for i, k in enumerate(kinds):
            if k != 'LETTER':
                continue
            ns = nospace(' '.join(l['t'] for l in L[i]))
            for j in (i - 1, i + 1):
                nk = kinds[j] if 0 <= j < len(kinds) else ''
                if nk in FORWARDER_MEMBER and FORWARDER_MEMBER[nk](ns):
                    kinds[i] = nk
                    changed = True
                    break
    return kinds


def page_kind(pg, lines):
    ns = nospace(' '.join(l['t'] for l in lines))
    if not ns:
        # หน้าเปล่า (กระดาษคั่นในชุดสแกน) — ไม่ใช่ "หน้าที่อ่านไม่ได้"
        if not list(pg.annots() or []):
            pix = pg.get_pixmap(dpi=20, colorspace=pymupdf.csGRAY)
            if sum(1 for b in pix.samples if b < 200) <= 3:
                return 'BLANK'
        return 'SCAN'
    if 'MAC-NELS' in ns and 'ATTACHEDSHEETS' in ns:
        return 'SHARED_MACNELS'
    if 'INBOUNDCARGOMANIFEST' in ns:
        return 'SHARED_CMS'
    if 'AMENDMENTOFMANIFEST' in ns and 'B/LCHANGED' in ns:
        return 'SHARED_DSV'
    # ฟอร์มของ forwarder ที่มีตัวอ่านเฉพาะอยู่แล้วใน build_edi.py — เจอในชุดรวมเมื่อไรส่งให้ตัวอ่านนั้นอ่านอัตโนมัติ
    for kind, test in FORWARDER_PAGES:
        if test(ns, lines):
            return kind
    if 'ATTACHEDRIDER' in ns and "SHIPPER'SLOAD" not in ns:
        return 'RIDER'
    if 'MARK&NO.' in ns and 'CNTR:' in ns and 'B/LNO.:' in ns:
        return 'MARKNO'
    if "SHIPPER'SLOAD" in ns:
        return 'BLFORM'
    up = [l['t'].upper() for l in lines]
    if any(u in ('CONSIGNEE',) for u in up) and any(u.startswith('QUANTITY') for u in up) \
            and any(u.startswith('DESCRIPTION') for u in up) \
            and any(u.startswith('GROSS WEIGHT') or u.startswith('WEIGHT/VOLUME') or u == 'G.W.' for u in up):
        # LEO GLOBAL: WEIGHT/VOLUME ; SHIP MASTER: 'G.W.'
        return 'TABLE'
    if len(lines) <= 3 and not BL_TXT_RE.search(ns):
        return 'SCAN'       # ภาพสแกนที่มีแค่ FreeText แปะทับ (เช่นเลขตู้)
    return 'LETTER'


# ═══════════════ ARRIVAL NOTICE / สำเนา B/L ของ Heung-A ═══════════════
def _lab(lines, key):
    for l in lines:
        if nospace(l['t']).startswith(key):
            return l
    return None


def red_boxes(pg):
    """กรอบสี่เหลี่ยมสีแดงที่ลูกค้าวาดบน PDF (Square annotation เส้นแดง) — พิกัดแบบเดียวกับ page_lines"""
    rot = page_rot(pg)
    out = []
    for a in pg.annots() or []:
        st = (a.colors or {}).get('stroke') or []
        if a.type[1] == 'Square' and len(st) >= 3 and st[0] > 0.7 and st[1] < 0.4 and st[2] < 0.4:
            out.append(a.rect if rot not in ((0, -1), (0, 1)) else pymupdf.Rect(_unrotate(pg, rot, *a.rect)))
    # กรอบแดงที่ถูก "พิมพ์รวม" เป็นเส้นในหน้า (ไม่ใช่ annotation) — SWDN 2026-11-27 หน้า 21: เส้นแดงสี่เหลี่ยมไม่มีสีพื้น
    try:
        drs = pg.get_drawings()
    except Exception:
        drs = []
    for dr in drs:
        c = dr.get('color') or ()
        r = dr.get('rect')
        if len(c) >= 3 and c[0] > 0.7 and c[1] < 0.4 and c[2] < 0.4 and not dr.get('fill') and r is not None \
                and r.width > 25 and r.height > 10 \
                and {it[0] for it in dr.get('items') or []} <= {'re', 'l', 'qu'} \
                and not any(abs(r.x0 - o.x0) < 2 and abs(r.y0 - o.y0) < 2 for o in out):
            out.append(r if rot not in ((0, -1), (0, 1)) else pymupdf.Rect(_unrotate(pg, rot, *r)))
    return out


def parse_blform(lines, annots, page, gap_cut=False, boxes=None):
    """gap_cut (OCR ของสแกน): ตัดพื้นที่สินค้าที่ช่องว่างแนวตั้ง > 70pt — ใต้นั้นคือโน้ตลายมือ/ตรายางของลูกค้า
       boxes: กรอบสีแดงของลูกค้า — MARKS/DESC ตรวจเฉพาะข้อความในกรอบ"""
    bl_line = next((l for l in lines if BL_TXT_RE.fullmatch(nospace(l['t'])) and l['x'] > 330 and l['y'] < 130), None)
    if not bl_line:
        return {}
    rec = blank(nospace(bl_line['t']), page)
    rec['src'] = 'BLFORM'
    body = [l for l in lines if not in_annot(l, annots)]

    sl = _lab(body, "SHIPPER'SLOAD")
    if sl is None:
        # จำนวนหีบห่อพิมพ์ติดป้ายเป็นบรรทัดเดียว ('8 TANK SHIPPER'S LOAD, COUNT & WEIGHT' — SAWASDEE MIMOSA 2611S)
        jl = next((l for l in body if "SHIPPER'S LOAD" in l['t'].upper()), None)
        if jl is None:
            return {}
        i = jl['t'].upper().find("SHIPPER'S LOAD")
        xs = jl['x'] + (jl['x1'] - jl['x']) * i / len(jl['t'])
        body = [o for o in body if o is not jl] + [dict(jl, x1=xs - 4, t=jl['t'][:i].strip()), dict(jl, x=xs, t=jl['t'][i:])]
        sl = body[-1]
    y_sl, x_sl = sl['y'], sl['x']
    tots = [l['y'] for l in (_lab(body, '(*)TOTALNO'), _lab(body, 'SAY:')) if l]
    # สแกน/OCR: '(*) Totar ilo. ...' / '(i) Total 16.' — บรรทัดรวมที่ขึ้นต้น '(x) TOTA'
    tots += [l['y'] for l in body if (re.match(r'^\(.\)TOTA', nospace(l['t'])) or re.search(r'TOTAL(NO|OF)', nospace(l['t'])))
             and l['y'] > y_sl]
    y_t = min(tots) if tots else y_sl + 230
    # น้ำหนัก = ตัวเลขติดหน่วย KGS ('25,040.000 KGS') — ไม่ใช่ '720 PKGS' ในช่องจำนวน (SAWASDEE MIMOSA 2611S หน้า 175)
    tol = 6 if gap_cut else 4            # ภาพสแกน (OCR) แถวเอียงได้ 4–5pt (SWDN 2026-11-27 หน้า 24)
    kgs = next((l for l in body if abs(l['y'] - y_sl) < tol and re.search(r'KGS', l['t']) and not re.search(r'PKGS?\b', l['t'])), None)
    cbm = next((l for l in body if abs(l['y'] - y_sl) < tol and re.search(r'CBM', l['t'])), None)
    if gap_cut and not cbm and kgs:      # OCR อ่านคำ CBM ไม่ออก: ตัวเลขทศนิยมขวาของน้ำหนักในแถวเดียวกัน = MEAS.
        cbm = next((l for l in body if abs(l['y'] - y_sl) < tol and l['x'] > kgs['x1']
                    and re.fullmatch(r'[\d,]+\.\d+', l['t'].strip()) and be.num(l['t']) < 1000), None)
    pkl = next((l for l in body if abs(l['y'] - y_sl) < 4 and l['x'] < x_sl - 2 and re.match(r'^\d', l['t'])), None)
    x_gw = (kgs['x'] if kgs else 420) - 6
    rec['gw'] = be.num(kgs['t']) if kgs else None
    rec['meas'] = be.num(cbm['t']) if cbm else None
    if cbm is not None and cbm is kgs:   # 'KGS 89.0000 CBM' / '41,000.000 KGS 41.000 CBM' ชิ้นเดียวกัน → ตัวเลขหลัง KGS
        mm = re.search(r'KGS\W*([\d,]+\.?\d*)', cbm['t'])
        rec['meas'] = be.num(mm.group(1)) if mm else None
    if pkl:
        rec['pkgs'], rec['pkgtype'] = split_pkg(pkl['t'])
        if not rec['pkgtype']:            # '7' กับ 'PALLETS' เป็นคนละชิ้นข้อความในแถวเดียวกัน
            pw = next((l for l in body if abs(l['y'] - y_sl) < 4 and pkl['x1'] - 2 <= l['x'] < x_sl - 2
                       and PKG_WORD_RE.match(l['t'].strip())), None)
            rec['pkgtype'] = pw['t'].strip() if pw else ''

    # ผู้รับ / ผู้รับแจ้ง: ใช้ป้าย Consignee / Notify Party / Pre Carriage ถ้ามี ไม่งั้นแบ่งบล็อกจากช่องว่างแนวตั้ง
    lc, ln, lp = _lab(body, 'CONSIGNEE'), _lab(body, 'NOTIFYPARTY'), _lab(body, 'PRECARRIAGE')
    left = sorted((l for l in body if l['x'] < 200 and l['y'] < y_sl - 30 and not BL_TXT_RE.search(nospace(l['t']))),
                  key=lambda l: l['y'])
    if lc and ln:
        y_end = lp['y'] if lp else ln['y'] + 100
        cons_lines = [l['t'] for l in left if lc['y'] + 3 < l['y'] < ln['y'] - 1]
        noti_lines = [l['t'] for l in left if ln['y'] + 3 < l['y'] < y_end - 1]
    else:
        blocks, cur = [], []
        for l in left:
            if cur and l['y'] - cur[-1]['y'] > 16:
                blocks.append(cur)
                cur = []
            cur.append(l)
        if cur:
            blocks.append(cur)
        cons_lines = [l['t'] for l in blocks[1]] if len(blocks) > 1 else []
        noti_lines = [l['t'] for l in blocks[2]] if len(blocks) > 2 else []
    cons = company_name(cons_lines)
    notify = company_name(noti_lines)
    if re.match(r'^TO\s+(?:THE\s+)?ORDER', cons, re.I) and notify and not re.match(r'^SAME\s+AS', notify, re.I):
        rec['cons_note'] = f'ENTER Consignee = {cons} → ใช้ชื่อ Notify'
        cons = notify
    rec['cons'] = cons
    if notify and not re.match(r'^SAME\s+AS', notify, re.I):
        rec['notify'] = notify

    # พื้นที่สินค้า: ใต้แถว SHIPPER'S LOAD ถึงบรรทัด (*) Total No.
    cargo = [l for l in body if y_sl + 3 < l['y'] < y_t - 1]
    if gap_cut and cargo:
        ys = sorted(l['y'] for l in cargo)
        cut = next((b for a, b in zip(ys, ys[1:]) if b - a > 70), None)
        if cut is not None:
            cargo = [l for l in cargo if l['y'] < cut]
    cont_line = lambda t: bool(CONT2_RE.search(t.upper())) and bool(
        re.match(r"^(?:R?\d{2}[A-Z]{0,3}'?)?\s*[A-Z]{3}[UJZ]\d{7}", nospace(t)))
    noise = lambda t: bool(re.match(r"^\d{2}'?[A-Z]*X\d+$", nospace(t))) or nospace(t).startswith('PARTICULARSFURNISHED') \
        or nospace(t) in ("SAIDTOCONTAIN", "SHIPPER'SLOAD,COUNT&WEIGHT")
    marks, desc, typed = [], [], []
    box_of = lambda l: next((b for b in boxes or [] if b.x0 - 2 <= l['x'] and l['x1'] <= b.x1 + 2
                             and b.y0 - 2 <= l['y'] + 4 <= b.y1), None)
    # ใช้เฉพาะกรอบที่ครอบ "บรรทัดแรก" ของ MARKS/DESC — กรอบอื่นในช่องสินค้าคือโน้ตของลูกค้า ('PAY DO BKK')
    box_m = box_d = None
    in_box = lambda l: False
    bm, bd = [], []                         # บรรทัดที่อยู่ในกรอบสีแดงของลูกค้า
    for l in sorted(cargo, key=lambda l: (l['y'], l['x'])):
        t = l['t']
        if l is pkl or noise(t):
            continue
        if is_status_line(t) and t != 'CY/CY':
            typed.append(t)
            continue
        if l['x'] < x_sl - 10:
            if not cont_line(t):
                if not marks:
                    box_m = box_of(l)
                marks.append(t)
                if box_m is not None and box_of(l) is box_m:
                    bm.append(t)
        elif l['x'] < x_gw:
            if not desc:
                box_d = box_of(l)
            desc.append(t)
            if box_d is not None and box_of(l) is box_d:
                bd.append(t)
    rec['marks'] = ' '.join(marks)
    rec['desc'] = ' '.join(desc)
    # กฎผู้ใช้ (2026-11-25): ลูกค้าตีกรอบสีแดงที่ MARKS/DESC = ตรวจเฉพาะข้อความในกรอบ (SAWASDEE MIMOSA 2611S หน้า 138)
    if bd:
        rec['desc_boxed'], rec['desc_boxed_full'] = ' '.join(bd), rec['desc']
    if bm:
        rec['marks_boxed'], rec['marks_boxed_full'] = ' '.join(bm), rec['marks']
    # กฎผู้ใช้ (2026-11-27, SWDN 1100797): กรอบแดงครอบช่องตัวเลข (จำนวน / G.W. / MEAS.) = ลูกค้าแจ้งเฉพาะช่องในกรอบ —
    # ช่องตัวเลขที่ไม่อยู่ในกรอบ "ไม่ได้ใส่ข้อมูลนั้น" → ไม่ตรวจ (ตีกรอบแค่ Gross weight → MEAS. ไม่ตรวจ)
    _num = {'pkgs': pkl, 'gw': kgs, 'meas': cbm}
    _in = {k for k, l in _num.items() if l is not None and box_of(l)}
    if _in:
        rec['skip_fields'] = [k for k in _num if k not in _in and rec.get(k) is not None]
    if rec['pkgs'] and not rec['pkgtype']:   # ชนิดหีบห่ออยู่ในรายละเอียด: '920 PACKAGES = 920 BASKETS'
        mp = re.search(rf"\b{rec['pkgs']}\s+([A-Z()]+)", rec['desc'], re.I)
        if mp and PKG_WORD_RE.match(mp.group(1)):
            rec['pkgtype'] = mp.group(1).upper()
    rec['cont'] = ';'.join(conts_in(*[l['t'] for l in cargo if l['x'] < x_sl - 10 and cont_line(l['t'])],
                                    *[a['t'] for a in annots]))

    # STATUS ที่ลูกค้าพิมพ์เพิ่ม (ทั้งข้อความแบนราบบนหน้า และ FreeText annotation)
    for l in body:
        if l['y'] >= y_t - 1 or l['y'] <= y_sl + 3:
            if is_status_line(l['t']) and l['t'] != 'CY/CY':
                typed.append(l['t'])
    labelled = _apply_labelled_annots(rec, annots, lc, ln)
    for a in annots:
        if a['t'] in labelled:
            continue
        # FreeText สั้น ๆ ที่ลงท้ายด้วย CY / LCL / FCL ('CY', 'ตามหน้า BL  CY') = STATUS ที่ลูกค้าแจ้ง (SWDN 2026-11-27)
        typed += [p for p in a['lines'] if is_status_line(p)
                  or (len(p.strip()) <= 30 and re.search(r'(?:^|\s)(?:CY(?:/CY)?|FCL|LCL)\s*$', p, re.I))]
        rec['dg'] = rec['dg'] or be.find_dg(a['t'])
        rec['reefer'] = rec['reefer'] or find_reefer(a['t'])
        if be.TRANSIT_RE.search(a['t']):
            rec['transit_raw'] = a['t']
    rec['status_raw'] = rec['status_raw'] or clean_status(typed)
    return {rec['bl']: finish(rec)}


def _apply_labelled_annots(rec, annots, lc, ln):
    """FreeText ที่ลูกค้าพิมพ์ทับ: ช่อง Consignee (แก้ชื่อผู้รับ) หรือข้อความมีป้าย 'QUANTITY PACKAGE:' ฯลฯ"""
    used = set()
    LBL = [('pkgs', r'(?:QUANTITY\s*PACKAGE|NO\.?\s*OF\s*PKGS?|PACKAGES?)'), ('desc', r'(?:CARGO\s*DESCRIPTION|DESCRIPTION)'),
           ('marks', r'MARKS?\s*&\s*NOS?\.?'), ('gw', r'GROSS\s*WEIGHT'), ('meas', r'(?:MEAS(?:UREMENT)?|CBM)'),
           ('status', r'STATUS')]
    for a in annots:
        hits = {}
        free = []
        for p in a['lines']:
            for k, pat in LBL:
                m = re.match(rf'^{pat}\s*[:=]\s*(.*)$', p, re.I)
                if m:
                    hits[k] = m.group(1).strip()
                    break
            else:
                free.append(p)
        in_cons_box = lc is not None and ln is not None and lc['y'] - 10 <= a['rect'].y0 < ln['y'] - 2
        if hits or in_cons_box:
            used.add(a['t'])
            # ชื่อผู้รับ = บรรทัดแรกที่ไม่มีป้าย เฉพาะเมื่อพิมพ์ทับช่อง Consignee หรือเป็นใบแก้ไขหลายช่อง (≥2 ป้าย)
            if free and (in_cons_box or len(hits) >= 2) and not CONT2_RE.search(free[0].upper()):
                rec['cons'] = free[0]
            if 'pkgs' in hits:
                rec['pkgs'], rec['pkgtype'] = split_pkg(hits['pkgs'])
            for k in ('desc', 'marks'):
                if k in hits:
                    rec[k] = hits[k]
            if 'gw' in hits:
                rec['gw'] = be.num(hits['gw'])
            if 'meas' in hits:
                rec['meas'] = be.num(hits['meas'])
            if 'status' in hits:
                rec['status_raw'] = clean_status([hits['status']])
    return used


def parse_rider(lines, page):
    """ATTACHED RIDER ของ B/L: คืน (bl, marks, typed status lines, desc)
       desc = รายละเอียดใต้ '<< Description Information >>' (แทน 'AS PER ATTACHED DESCRIPTIONS' ของ B/L)"""
    bl = next((nospace(l['t']) for l in lines if BL_TXT_RE.fullmatch(nospace(l['t']))), '') \
        or next((m.group(0) for l in lines for m in [re.search(r'HASL[A-Z0-9]{8,14}', nospace(l['t']))]
                 if m and nospace(l['t']).startswith('B/LNO')), '')       # 'B/L NO :HASLS2…' ในบรรทัดเดียว
    # 'Mark Information' — สแกนอ่านเป็น 'Mank Informati-on'
    mi = next((l for l in lines if re.search(r'MA.KINFORMATI', nospace(l['t']))), None)
    di = next((l for l in lines if 'DESCRIPTIONINFORMATION' in nospace(l['t'])), None)
    marks, typed, desc = [], [], []
    stop = None
    for l in sorted(lines, key=lambda l: (l['y'], l['x'])):
        if is_status_line(l['t']):
            typed.append(l['t'])
            continue
        if mi and l['y'] > mi['y'] + 3 and l['x'] < 250:
            marks.append(l['t'])
        if di and l['y'] > di['y'] + 3 and l['x'] > di['x'] - 20 and stop is None:
            if re.match(r'^(\*|CONTACT|E-?MAIL|TEL|FAX|ALSO\s+NOTIFY|TAX\s+CODE)', l['t'], re.I):
                stop = l['y']
            else:
                desc.append(l['t'])
    return bl, ' '.join(marks), typed, ' '.join(desc)


# ═══════════════ ใบ MARK&NO. / NO. OF PKGS / DESCRIPTION OF GOODS (2 B/L ต่อหน้า) ═══════════════
def parse_markno(lines, page):
    out = {}
    lbls = sorted((l for l in lines if nospace(l['t']) == 'B/LNO.:'), key=lambda l: l['y'])
    hdr = next((l for l in lines if nospace(l['t']) == 'MARK&NO.'), None)
    prev_y = (hdr['y'] + 5) if hdr else 0
    for lb in lbls:
        val = next((l for l in lines if abs(l['y'] - lb['y']) < 3 and l['x'] > lb['x'] + 20
                    and BL_TXT_RE.fullmatch(nospace(l['t']))), None)
        if not val:
            continue
        rec = blank(nospace(val['t']), page)
        blk = [l for l in lines if prev_y < l['y'] <= lb['y'] + 3]
        prev_y = lb['y'] + 3
        cl = next((l for l in blk if nospace(l['t']) == 'CNTR:'), None)
        y_cntr = cl['y'] if cl else lb['y'] - 20
        skip_y = [y_cntr, lb['y']]
        marks, desc, mid = [], [], []
        for l in sorted(blk, key=lambda l: (l['y'], l['x'])):
            t = l['t']
            if t.startswith('G.W.'):
                rec['gw'] = be.num(t)
            elif t.startswith('MEAS.'):
                rec['meas'] = be.num(t)
            elif any(abs(l['y'] - y) < 3 for y in skip_y) and l['x'] < 200:
                if abs(l['y'] - y_cntr) < 3:
                    rec['cont'] = ';'.join(conts_in(t)) or rec['cont']
            elif l['x'] < 200:
                marks.append(t)
            elif l['x'] < 340:
                mid.append(l)
            elif l['y'] < y_cntr:
                desc.append(t)
        cnt_l = next((l for l in mid if re.fullmatch(r'[\d,]+', l['t'])), None)
        if cnt_l:
            rec['pkgs'] = int(cnt_l['t'].replace(',', ''))
            pt = next((l for l in mid if l is not cnt_l and l['y'] < cnt_l['y'] + 30 and not re.fullmatch(r'[\d,]+', l['t'])), None)
            rec['pkgtype'] = pt['t'] if pt else ''
            rec['cons'] = ' '.join(l['t'] for l in mid if l is not cnt_l and l is not pt and l['y'] < y_cntr)
        rec['marks'] = ' '.join(marks)
        rec['desc'] = ' '.join(desc)
        out[rec['bl']] = finish(rec)
    return out


# ═══════════════ ตาราง AMENDMENT (Cargoport / ASIA ATLANTIC / INITIAL / Handle ฯลฯ) ═══════════════
def _status_value(lines, l):
    m = re.match(r'^\s*STATUS\s*[:=]?\s*(.*)$', l['t'], re.I)
    v = (m.group(1) if m else l['t']).strip(' :')
    if not v:
        v = ' '.join(o['t'] for o in sorted(lines, key=lambda o: o['x'])
                     if o is not l and abs(o['y'] - l['y']) < 3 and o['x'] > l['x'] and o['t'].strip(' :'))
    return clean_status([v])


def parse_table_page(lines, annots, page):
    """คืน (rows:dict[bl->rec], page_status, page_conts)"""
    hdr = next(l for l in lines if l['t'].upper() == 'CONSIGNEE')
    hy = hdr['y']
    H = [l for l in lines if abs(l['y'] - hy) < 6]
    col = {}
    for l in H:
        u = l['t'].upper()
        if u.startswith('MARKS') or u.startswith('MARK &'):        # 'MARK & NUMBER' (BUGATTI)
            col['MARKS'] = l['x']
        elif u.startswith('QUANTITY'):
            col['QTY'] = l['x']
        elif u.startswith('DESCRIPTION'):
            col['DESC'] = l['x']
        elif u == 'CONSIGNEE':
            col['CONS'] = l['x']
        elif u.startswith('GROSS WEIGHT') or u.startswith('WEIGHT/VOLUME') or u == 'G.W.':
            col['GW'] = l['x']
        elif u.startswith('HB/L'):
            col['HBL'] = l['x']
    bounds = sorted((x - 45, k) for k, x in col.items())

    def col_of(x):
        name = 'BL'
        for left, k in bounds:
            if x >= left:
                name = k
        return name

    tot = [l['y'] for l in lines if l['y'] > hy + 5 and (l['t'].upper() == 'TOTAL' or l['t'].upper().startswith('TOTAL ')
                                                       or l['t'].upper().startswith('CREATED BY'))]   # SHIP MASTER ท้ายตาราง
    # ('Page 1 of 1' ที่หัวกระดาษเหนือตาราง ไม่ใช่ท้ายตาราง — BUGATTI)
    pg_end = [l['y'] for l in lines if (re.match(r'^Page \d+ of \d+', l['t']) or l['t'].startswith('[FM-')) and l['y'] > hy]
    end_y = min(tot + pg_end + [1e9])

    anchors = []
    used = set()
    for l in sorted(lines, key=lambda l: l['y']):
        if l['y'] > hy + 5 and l['y'] < end_y and col_of(l['x']) == 'BL' and re.match(r'^HASL[A-Z0-9]+$', l['t']):
            bl = l['t']
            # ส่วนท้ายเลข B/L ที่ตัดขึ้นบรรทัดใหม่: '…104' + '9' หรือ 'HASLC0000000' + '2029A' (B/L ยังไม่ครบ 16 ตัว)
            suf = next((o for o in lines if o is not l and abs(o['x'] - l['x']) < 5 and 3 < o['y'] - l['y'] < 16
                        and (re.fullmatch(r'[A-Z0-9]{1,2}', o['t'])
                             or (len(bl) < 16 and re.fullmatch(r'\d[0-9]{0,5}[A-Z]?', o['t'])))), None)
            if suf:
                bl += suf['t']
                used.add(id(suf))
            used.add(id(l))
            anchors.append((l['y'], bl))
    rows = {}
    for i, (ay, bl) in enumerate(anchors):
        ny = anchors[i + 1][0] if i + 1 < len(anchors) else end_y
        rl = [l for l in lines if ay - 4 <= l['y'] < ny - 4 and id(l) not in used]
        rec = blank(bl, page)
        parts = {k: [] for k in ('BL', 'MARKS', 'QTY', 'DESC', 'CONS', 'GW', 'HBL')}
        st = []
        for l in sorted(rl, key=lambda l: (l['y'], l['x'])):
            c = col_of(l['x'])
            if re.match(r'^\s*STATUS\b', l['t'], re.I):
                st.append(_status_value(rl, l))
                continue
            # SHIP MASTER พิมพ์คำสั่งไว้ท้ายช่อง DESCRIPTION: '***ลากตู้***' / '1X20'DC' / 'CY/CY'
            if c in ('DESC', 'QTY') and (is_status_line(l['t']) or re.fullmatch(r"\d+\s*X\s*\d{2}'?\s*[A-Z]{0,4}|(?:CY|CFS)\s*/\s*(?:CY|CFS)",
                                                                       l['t'].strip(), re.I)):
                if is_status_line(l['t']) or '/' in l['t']:
                    st.append(clean_status([l['t'].strip('* ')]))
                continue
            parts[c].append(l['t'])
        # รายละเอียดสินค้าชิดซ้ายเกินเส้นแบ่งเข้าไปในคอลัมน์ QUANTITY (LEO GLOBAL: DESCRIPTION หัว x405 แต่ข้อมูล x357)
        # → ในคอลัมน์ QUANTITY เก็บเฉพาะตัวเลข/ชนิดหีบห่อ ที่เหลือย้ายไป DESC ตามลำดับ
        if parts['QTY']:
            keep = [t for t in parts['QTY'] if re.fullmatch(r'\d[\d,]*', t.strip()) or PKG_WORD_RE.fullmatch(t.strip())
                    or re.fullmatch(r'(?:WOODEN\s+)?[A-Z]+\(S\)', t.strip())
                    or (split_pkg(t)[0] and PKG_WORD_RE.match(split_pkg(t)[1] or '') and len(t.split()) <= 3)]
            if keep and len(keep) < len(parts['QTY']):
                parts['DESC'] = [t for t in parts['QTY'] if t not in keep] + parts['DESC']
                parts['QTY'] = keep
        rec['marks'] = ' '.join(parts['MARKS'])
        rec['desc'] = ' '.join(parts['DESC'])
        rec['cons'] = ' '.join(parts['CONS'])
        rec['pkgs'], rec['pkgtype'] = split_pkg(' '.join(parts['QTY']))
        gwt = ' '.join(parts['GW'])
        mg = re.search(r'G\.?\s*W\.?\s*([\d,]+\.?\d*)', gwt) or re.search(r'([\d,]+\.\d+)\s*KGS', gwt)
        mm = re.search(r'([\d,]+\.?\d*)\s*(?:M3|CBM)', gwt)
        mm_v = None
        if not mm:   # '25.0000' / 'M3' แยกบรรทัด
            nums = [x for x in re.findall(r'[\d,]+\.\d+', gwt)]
            mm_v = be.num(nums[1]) if len(nums) > 1 else None
        rec['gw'] = be.num(mg.group(1)) if mg else None
        rec['meas'] = be.num(mm.group(1)) if mm else mm_v
        # UNION WORLD: เลขตู้อยู่ในคอลัมน์ HB/L ใต้เลข HB/L (คั่นด้วย , และอาจขึ้นบรรทัดกลางเลข)
        toks = [t for t in re.split(r'[,\s]+', ' '.join(parts['HBL'])) if t]
        hbl_conts = []
        for i, t in enumerate(toks):
            if CONT2_RE.fullmatch(t):
                hbl_conts.append(t)
            elif i + 1 < len(toks) and re.fullmatch(r'[A-Z]{3}[UJZ]?\d{0,6}', t) and CONT2_RE.fullmatch(t + toks[i + 1]):
                hbl_conts.append(t + toks[i + 1])
        rec['cont'] = ';'.join(conts_in(*parts['BL'], *parts['MARKS'], *parts['DESC'], ' '.join(hbl_conts)))
        rec['status_raw'] = ' '.join(st)
        rows[bl] = rec
    page_status = [_status_value(lines, l) for l in lines
                   if re.match(r'^\s*STATUS\b', l['t'], re.I) and not any(ay - 4 <= l['y'] < end_y for ay, _ in anchors[:1])]
    page_conts = conts_in(*[l['t'] for l in lines if l['y'] < hy], *[a['t'] for a in annots])
    for a in annots:
        if is_status_line(a['t']):
            page_status.append(clean_status([a['t']]))
    return rows, [s for s in page_status if s], page_conts


# ═══════════════ จดหมายแก้ไขแบบข้อความอิสระ ═══════════════
_LETTER_STOP = re.compile(r'^(CONSIGNEE|CNEE|CONTAINER|CONT\.?\s*NO|\(?TOTAL|-{5,}|WE |YOUR|THANK|STATUS|REMARK|'
                          r'NOTIFY|PLACE OF|VESSEL|จ่าย|AS THE SHIPMENT TERM)', re.I)   # '(TOTAL : …' WISMA, 'AS THE SHIPMENT TERM' Masterwork
_L_CONS = r'(?:CONSIGNEE\s*NAME|CONSIGNEE|CNEE)'
_L_CONT = r'(?:CONTAINER\s*NO\.?|CONT(?:R)?\.?\s*(?:NO\.?)?)'
_L_PKG = r'(?:PACKAGES?|NO\.\s*OF\s*PCS|QUANTITY\s*PACKAGE|NO\.?\s*OF\s*P[’\']?\s*KGS?|QUANTITY(?=\s*[:\d]|\s*$))'   # 'NO OF P’KGS: 80 DRUM(S)', 'Quantity : 1105 CARTONS'
_L_GW = r'(?:WEIGHT\s*&\s*CBM|GROSS\s*(?:WEIGHT|WT\.?)|G\.\s*W\.?)'
_L_MEAS = r'(?:GROSS\s*CBM|CBM\.?|MEASUREMENT|MEAS\.?)'
_L_MARKS = r'(?:MARKS?\s*(?:AND|&)\s*(?:NUMBERS?|NOS?\.?|NO\.?)|SHIPPING\s*MARKS?)'
_L_DESC = r'(?:CARGO\s*DESCRIPTION|DESCRIPTION(?:\s*OF\s*GOODS)?)'


def _label_value(lines, pat):
    for l in sorted(lines, key=lambda l: l['y']):
        m = re.match(rf'^{pat}\s*[:=]?\s*(.*)$', l['t'], re.I)
        if not m:
            continue
        v = m.group(1).strip(' :=')
        if not v:
            # ค่าทางขวาแถวเดียวกัน — ต้องอยู่ใกล้ป้าย ('CONSIGNEE:' x55 กับ 'QUANTITY 38 CASES' x303 คนละช่อง — ชื่อผู้รับอยู่บรรทัดล่าง)
            right = sorted((o for o in lines if o is not l and abs(o['y'] - l['y']) < 6 and o['x'] > l['x']
                            and o['x'] - l.get('x1', l['x']) < 100), key=lambda o: o['x'])
            v = ' '.join(o['t'] for o in right).strip(' :=')
        if not v:
            near = sorted((o for o in lines if o['x'] > l['x'] + 40 and o['t'].startswith(':')
                           and abs(o['y'] - l['y']) < 22), key=lambda o: abs(o['y'] - l['y']))
            if near:
                v = near[0]['t'].strip(' :=')
        if not v:
            # ค่าอยู่บรรทัดล่างชิดซ้ายเดียวกับป้าย ('DESCRIPTION OF GOODS :' / 'FRUCTOSE SYRUP' — ABC CHEMICAL)
            below = sorted((o for o in lines if 0 < o['y'] - l['y'] < 30 and abs(o['x'] - l['x']) < 6),
                           key=lambda o: o['y'])
            if not below and re.match(rf'^{_L_CONS}', l['t'], re.I):
                # 'CONSIGNEE NAME :' พิมพ์กึ่งกลาง ชื่ออยู่บรรทัดล่างแต่คนละ x (GENIUS INTERFREIGHT)
                below = sorted((o for o in lines if 0 < o['y'] - l['y'] < 25 and o['x'] < l.get('x1', l['x'] + 100)
                                and o.get('x1', o['x'] + 100) > l['x']), key=lambda o: o['y'])
            if below and not re.match(r"^[A-Z .&'/’()-]+\s*[:=]", below[0]['t']) and not _LETTER_STOP.match(below[0]['t']) \
                    and not re.search(r'GROSS\s*WEIGHT|MEASU|QUANTITY|NO\.?\s*OF\s*P', below[0]['t'], re.I):   # ไม่ใช่หัวตาราง
                v = below[0]['t'].strip(' :=')
        if v:
            return v, l
    return '', None


# ═══════════════ โปรไฟล์ฟอร์มของแต่ละเจ้า (ผู้ใช้ชี้ตำแหน่งให้ แล้วจดไว้ที่นี่ 1 รายการต่อ 1 ฟอร์ม) ═══════════════
# sig  = ข้อความที่มีเฉพาะฟอร์มของเจ้านั้น (regex บนข้อความทั้งหน้า)
# hdr  = ป้ายหัวคอลัมน์ตามที่พิมพ์บนฟอร์ม → ช่องในรายงาน (MARKS/PKG/DESC/GW/MEAS)
# stop = บรรทัดแรกใต้ตารางที่ไม่ใช่ข้อมูลสินค้าแล้ว
# แต่ละบรรทัดใต้หัวตารางลงคอลัมน์ตาม "จุดกึ่งกลาง" ของบรรทัด เทียบกับเส้นแบ่งกึ่งกลางระหว่างหัวคอลัมน์ที่ติดกัน
# (หัวคอลัมน์มักพิมพ์กึ่งกลางช่อง ส่วนข้อมูลชิดซ้าย — ใช้ x ซ้ายตรง ๆ จะตกผิดช่อง)
FORM_PROFILES = [
    # ── SAWASDEE DENEB 2612S (2026-11-27) ──
    # CK LINE 'AMENDMENT' — ตารางแถวเดียว B/L / D/O / CONSIGNEE / MARKS / QUANTITY / DESCRIPTION / G.W / Measurment
    # ค่าชิดซ้ายแต่หัวอยู่กลางคอลัมน์; แถวยอดรวมอยู่ห่างลงไป ~200pt → 'gap' หยุดอ่าน
    {'name': 'CK LINE', 'sig': r'CK\s+LINE\s+\(THAILAND\)',
     'hdr': {'B/L NO.': 'BL', 'D/O NO.': 'REMARK', 'CONSIGNEE': 'CONS', 'MARKS & NUMBER': 'MARKS', 'QUANTITY': 'PKG',
             'DESCRIPTION': 'DESC', 'G.W': 'GW', 'MEASURMENT': 'MEAS', 'MEASUREMENT': 'MEAS'},
     'gap': 60, 'stop': r'^(TOTAL|REMARK|THANK)'},
    # EXTRA MARITIME — 'Gross Weigh Volume' หัวเดียวคลุม 2 ค่า: '12,600.00' (ไม่มีหน่วย = G.W.) / '30.000 CBM'
    {'name': 'EXTRA MARITIME', 'sig': r'E\s*X\s*T\s*R\s*A\s+M\s*A\s*R\s*I\s*T\s*I\s*M\s*E',
     'hdr': {'MARKS & NUMBER': 'MARKS', 'NBR OF PKGS': 'PKG', 'DESCRIPTION OF GOODS': 'DESC',
             'GROSS WEIGH VOLUME': 'GW', 'GROSS WEIGHT VOLUME': 'GW'},
     'skip': r'^KGS\.?$', 'stop': r'^(CNEE|STATUS|B/L NO)'},
    # RICH FREIGHT 'ATTACHED LIST' (Page 2/2) — B/L อยู่ในบรรทัด 'DO NO. : …', CNEE/CTN/STATUS ใต้ตาราง
    {'name': 'RICH FREIGHT', 'sig': r'RICH\s+FREIGHT\s+FORWARDER',
     'hdr': {'MARKS': 'MARKS', 'NO.OF PACKAGES': 'PKG', 'DESCRIPTION OF GOODS': 'DESC', 'GROSS WEIGHT': 'GW',
             'MEASUERMENT': 'MEAS', 'MEASUREMENT': 'MEAS'},
     'rx': {'CONS': r'^CNEE\s*:\s*(.+)$'},
     'stop': r'^(CNEE|DO NO|CTN|STATUS|\*\*\*)'},
    # สหไทย เทอร์มินอล (Sahathai) 'pre-lading manifest' — หัว Consignee Name / No of package / Description / G.W. /
    # Mark & No.; 'Meas.' พิมพ์แยกใต้ G.W. ค่าอยู่บรรทัดล่าง ('68.000CBM' = MEAS. ตามหน่วย)
    {'name': 'SAHATHAI TERMINAL', 'sig': r'Sahathai\s+Terminal',
     'hdr': {'CONSIGNEE NAME': 'CONS', 'NO OF PACKAGE': 'PKG', 'DESCRIPTION OF GOODS': 'DESC', 'G.W.': 'GW',
             'MARK & NO.': 'MARKS'},
     'skip': r'^MEAS\.?$', 'stop': r'^(REMARK|THANK|YOURS)'},
    {'name': 'H & FRIENDS GTL', 'sig': r'H\s*&\s*FRIENDS\s+GTL',          # 2026-11-03 DONGJIN CONFIDENT 0153S
     'hdr': {'MARKS': 'MARKS', 'QUANTITY': 'PKG', 'DESCRIPTION': 'DESC', 'WEIGHT': 'GW', 'CBM': 'MEAS'},
     'stop': r'^(จ่ายเงิน|Consignee Name|Place Of Delivery|Status|Notify Party|WE HEREBY)'},
    # จำนวน+รายละเอียดอยู่คอลัมน์เดียว ('29 PALLETS' ชิดซ้าย รายละเอียดอยู่ขวา) → QTYDESC แยกเองทีละบรรทัด
    {'name': 'BILLION LOGISTICS', 'sig': r'BILLION\s+LOGISTICS',            # 2026-11-03 DONGJIN CONFIDENT 0153S
     'hdr': {'MARKS & NOS': 'MARKS', 'QUANTITY & DESCRIPTION OF GOOD': 'QTYDESC',
             'QUANTITY & DESCRIPTION OF GOODS': 'QTYDESC', 'GROSS WEIGHT': 'GW', 'MEAS': 'MEAS'},
     'stop': r'^(\*|CONTAINER NO|TOTAL\s*:|AS THE ORIGINAL)'},
    # LOGISALL (LAT) 'AMEND MANIFEST' — ตารางมีคอลัมน์ B/L และ CONSIGNEE; หัวบางคอลัมน์พิมพ์ 2 ชั้น
    # (PACKAGES/QUANTITY, CONTENTS/DESCRIPTION, WEIGHT/KGS). DESC อาจเป็น 'AS PER ATTACHED DESCRIPTIONS'
    # แล้วรายละเอียดจริงอยู่หน้า ATTACHED RIDER ถัดไป
    {'name': 'LOGISALL (LAT)', 'sig': r'KAE\s+LAT',                         # 2026-11-03 DONGJIN CONFIDENT 0153S
     'hdr': {'B/L': 'BL', 'MARKS & NOS.': 'MARKS', 'PACKAGES': 'PKG', 'QUANTITY': 'PKG', 'CONTENTS': 'DESC',
             'DESCRIPTION': 'DESC', 'CONSIGNEE': 'CONS', 'WEIGHT': 'GW', 'KGS': 'GW', 'CBM': 'MEAS',
             'REMARK': 'REMARK'},
     'stop': r'^(SAY\s*:|STATUS|TEMP\s*:|YOURS)'},
    # YOUR SHIPPING (โลโก้เป็นรูป) — หัวตารางและข้อมูลพิมพ์เป็นข้อความบรรทัดเดียวคั่นด้วยช่องว่างยาว
    # ('467 PACKAGES      PERSONAL EFFECT (HOUSEHOLD GOODS)      4,765.00 KGS') → mode 'spaced':
    # แยกด้วยช่องว่าง ≥2 ตัว แล้วจับคู่จากขวา (คอลัมน์ MARKS มักอยู่บรรทัดแยก)
    {'name': 'YOUR SHIPPING', 'sig': r'WE SHOULD VERY MUCH APPRECIATE.*PRINCIPALS MANIFEST',  # 2026-11-03 0153S
     'mode': 'spaced',
     'hdr': {'MARKS AND NOS.': 'MARKS', 'QUANTITY': 'PKG', 'DESCRIPTION OF GOODS': 'DESC',
             'GROSS WEIGHT': 'GW', 'MEASUREMENT': 'MEAS', 'CBM': 'MEAS'},
     'status': r'STATUS\s*:?\s*(\S.*)$',
     'stop': r'^(CONSIGNEE|NOTIFY|YOURS|THANK)'},
    # rx = ค่าที่พิมพ์แบบ 'ป้าย : ค่า' นอกตาราง (อ่านจากแถวที่รวมบรรทัดระดับเดียวกันแล้ว) — ใช้ทับค่าจากตาราง
    {'name': 'ACE INTER FREIGHT', 'sig': r'ACE\s+INTER\s+FREIGHT',          # 2026-11-03 DONGJIN CONFIDENT 0153S
     'hdr': {'MARK AND NOS.': 'MARKS', 'NO.OF PKGS. DESCRIPTION OF GOODS': 'QTYDESC'},
     'min_hdr': 2,
     'rx': {'GW': r'GROSS WEIGHT\s*=\s*([\d,]+\.?\d*)', 'MEAS': r'CBM\.?\s*=\s*([\d,]+\.?\d*)',
            'CONS': r'^CNEE\s*:\s*(.+)$', 'STATUS': r'\*\*\s*STATUS\s+(.+?)\s*\*\*'},
     'stop': r'^(GROSS WEIGHT|SHIPPER|CNEE|TOTAL QUANTITY)'},
    {'name': 'UNICORN GLOBAL LOGISTICS', 'sig': r'UNICORN\s+GLOBAL\s+LOGISTICS',   # 2026-11-03 0153S (หน้า TABLE)
     'hdr': {'B/L CHANGE NO.': 'BL', 'MARK & NUMBERS': 'MARKS', 'QUANTITY': 'PKG', 'DESCRIPTION': 'DESC',
             'CONSIGNEE': 'CONS', 'GROSS WEIGHT': 'GW'},
     'rx': {'STATUS': r'^STATUS\s*:\s*(.+)$'},
     'stop': r'^(TOTAL|\d[\d,]*\s+PACKAGES?\s*/|THANKS)'},
    # Gebruder Weiss — ชื่อบริษัทเป็นรูปโลโก้ จึงจับจากประโยคในจดหมาย; GWT/MEA อยู่ท้ายกรอบ ห่างจากตาราง
    {'name': 'GEBRUDER WEISS', 'sig': r'Should be amended as follow.*We further guarantee',   # 2026-11-03 0153S
     'hdr': {'MARKS': 'MARKS', "P'KGS": 'PKG', 'DESCRIPTION': 'DESC'},
     'rx': {'GW': r'GWT\s*:\s*([\d,]+\.?\d*)', 'MEAS': r'MEA\s*:\s*([\d,]+\.?\d*)',
            'CONS': r'^1\.\s*CONSIGNEE\s+(.+)$', 'STATUS': r'\*\*\s*STATUS\s*:\s*(.+?)\s*\*\*'},
     'stop': r'^(GWT|MEA|CTNR)'},
    # MARKS มีแค่บรรทัดใต้หัว 'MARKS AND NOS.' — บรรทัด 'CNEE : …' ด้านล่างไม่ใช่ marks
    {'name': 'VISION LOGISTICS', 'sig': r'VISION\s+LOGISTICS',              # 2026-11-03 DONGJIN CONFIDENT 0153S
     'hdr': {'MARKS AND NOS.': 'MARKS', "NO.OF P'KGS": 'PKG', 'DESCRIPTION OF GOODS': 'DESC'},
     'rx': {'GW': r'GROSS\s*WT\.?\s*=\s*([\d,]+\.?\d*)', 'MEAS': r'CBM\.?\s*=\s*([\d,]+\.?\d*)',
            'CONS': r'^CNEE\s*:\s*(.+)$', 'STATUS': r'\*\*\s*(.+?)\s*\*\*'},
     'stop': r'^(GROSS\s*WT|CBM\.?\s*=|CNEE|WE HEREBY)'},
    # CNEE/STATUS เป็นป้ายบนหัวจดหมาย; G.W./MEAS เป็นแถว 'G.W.: … KGS.  MEAS: … CBMS.' ใต้ตาราง
    {'name': 'T.V.L. GLOBAL LOGISTICS', 'sig': r'T\.V\.L\.\s*GLOBAL\s+LOGISTICS',   # 2026-11-03 DONGJIN 0153S
     'hdr': {'MARKS & NOS': 'MARKS', 'PACKAGES': 'PKG', 'DESCRIPTION': 'DESC'},
     'rx': {'GW': r'G\.W\.\s*:\s*([\d,]+\.?\d*)', 'MEAS': r'MEAS\s*:\s*([\d,]+\.?\d*)',
            'CONS': r'^CNEE\s*:\s*(.+)$', 'STATUS': r'^STATUS\s*:\s*(.+?)(?:\s+SHED\s*:.*)?$'},
     'stop': r'^(G\.W\.|MEAS|WE HE)'},
    # Star Concord 'Sea Freight Summary Manifest' (แนวนอน) — ตาราง Bill Number / Wgt-Vol-Pkg / Consignee / Origin /
    # Dest / Mark / Goods Description ; ใต้ตาราง 'STATUS:CY/CY' และ 'Cargo movement = 7' แถวเดียวกัน
    {'name': 'STAR CONCORD', 'sig': r'Sea\s+Freight\s+Summary\s+Manifest',     # 2026-11-24 SAWASDEE VEGA 2614S
     'hdr': {'BILL NUMBER': 'BL', 'WGT / VOL / PKG': 'PKG', 'CONSIGNEE': 'CONS', 'ORIGIN': 'REMARK', 'DEST': 'REMARK',
             'MARK': 'MARKS', 'GOODS DESCRIPTION': 'DESC'},
     'rx': {'STATUS': r'STATUS\s*:\s*([A-Z]+(?:\s*/\s*[A-Z]+)?)', 'MOVEMENT': r'Cargo\s*movement\s*[=:]?\s*([0-9A-Z])\b'},
     'stop': r'^(Total Shipments|NOTIFY PARTY|Container No)'},
    # NEW HANYUAN — ตาราง MARKS / NO. OF PKGS. / DESCRIPTION ; ใต้ตาราง 'CNEE : …', 'G.W. : … KGS / … CBM', 'STATUS : …'
    {'name': 'NEW HANYUAN LOGISTICS', 'sig': r'NEW\s+HANYUAN\s+LOGISTICS',   # 2026-11-24 SAWASDEE VEGA 2614S
     'hdr': {'MARKS': 'MARKS', 'NO. OF PKGS.': 'PKG', 'DESCRIPTION': 'DESC'},
     'rx': {'CONS': r'^CNEE\s*:\s*(.+)$', 'GW': r'([\d,]+\.\d+)\s*KGS', 'MEAS': r'([\d,]+\.\d+)\s*CBM',
            'STATUS': r'^STATUS\s*:\s*(.+)$'},
     'stop': r'^(CNEE|G\.W\.|STATUS|WE HEREBY)'},
    # NIPPON EXPRESS — ตาราง SHIPPING MARKS / QUANTITY & DESCRIPTIONS / CONSIGNEE ; คอลัมน์กลางพิมพ์ป้าย 'PKG:' 'W/T:'
    # 'M/M:' กับตัวเลข/หน่วยเป็นชิ้นแยก (อ่านด้วย rx แล้วข้ามชิ้นเหล่านั้น) ; DESC 'DESCRIPTION AS PER ATTACHED' =
    # หน้า 'ATTACHED SHEET (DESCRIPTION)' ที่ตามมา ; B/L ย่อยลูกค้าพิมพ์ตัวอักษรสีแดงต่อท้าย M-B/L ('A')
    {'name': 'NIPPON EXPRESS', 'sig': r'NIPPON\s+EXPRESS\s+LOGISTICS',       # 2026-11-24 SAWASDEE VEGA 2614S
     'hdr': {'SHIPPING MARKS': 'MARKS', 'QUANTITY & DESCRIPTIONS': 'DESC', 'CONSIGNEE': 'CONS'},
     'skip': r'^((PKG|W/T|M/M)\s*:?|[\d,.]+|KGS?|M3|CBM|PACKAGES?|PALLETS?|CARTONS?|CASES?|BOX(ES)?)$',
     'rx': {'PKG': r'PKG\s*:\s*(\d[\d,]*\s+[A-Z]+)', 'GW': r'W/T\s*:\s*([\d,]+\.?\d*)',
            'MEAS': r'M/M\s*:\s*([\d,]+\.?\d*)', 'STATUS': r'^SERVICE TYPE\s*:\s*(.+)$'},
     'stop': r'^(TOTAL|WE HEREBY)'},
    # SINTOP — หัวตาราง 'MARKS : NO. OF PKGS. DESCRIPTION OF GOODS G.W.' เป็นข้อความบรรทัดเดียว → แบ่งด้วยขอบซ้าย x
    {'name': 'SINTOP INTERNATIONAL', 'sig': r'SINTOP\s+INTERNATIONAL',        # 2026-11-24 SAWASDEE VEGA 2614S
     'hdr': {},
     'bands': {'hdr': r'^MARKS\s*:', 'cols': [(110, 'MARKS'), (185, 'PKG'), (400, 'DESC'), (1e9, 'GWM')]},
     'rx': {'CONS': r'^CONSIGNEE\s*:\s*(.+)$', 'STATUS': r'^STATUS\s*:\s*(.+)$',
            'GW': r'([\d,]+\.\d+)\s*KGS', 'MEAS': r'([\d,]+\.\d+)\s*CBM'},
     'stop': r'^(CONSIGNEE|BEST REGARDS|_{5,})'},
    # Maxtop — 'AMENDMENT OF MANIFEST' หัว MBL NO / HBL NO ; ตาราง MARKS / NO OF P'KGS. / DESCRIPTION OF GOODS ;
    # ใต้ตาราง 'CONSIGNEE NAME:', 'STATUS : CY/CY', 'CONTAINER NO :', 'GW : … KGS', 'CBM : … CBM'
    {'name': 'MAXTOP LOGISTICS', 'sig': r'Maxtop\s+Logistics',               # 2026-11-24 DONGJIN CONFIDENT 0154S
     'hdr': {'MARKS': 'MARKS', "NO OF P’ KGS.": 'PKG', "NO OF P' KGS.": 'PKG', 'DESCRIPTION OF GOODS': 'DESC'},
     'rx': {'CONS': r'^CONSIGNEE\s*NAME\s*:\s*(.+)$', 'STATUS': r'^STATUS\s*:\s*(.+)$',
            'GW': r'^GW\s*:\s*([\d,]+\.?\d*)', 'MEAS': r'^CBM\s*:\s*([\d,]+\.?\d*)'},
     'stop': r'^(CONSIGNEE|STATUS|CONTAINER|GW\b|CBM\b|\*)'},
    # FUKUYAMA GRAND / INTERCOSPED (โปรแกรมเดียวกัน) — 'WE SHALL BE OBLIGED IF YOU KINDLY AMEND THE FOLLOWING MANIFEST';
    # ตาราง MARKS / NO.OF PACKAGES / DESCRIPTION OF GOODS / GROSS WEIGHT / MEASUERMENT ; 'CNEE :', 'STATUS * CY *'
    {'name': 'FUKUYAMA / INTERCOSPED', 'sig': r'KINDLY AMEND THE FOLLOWING MANIFEST AS PER FOLLOWING',  # 2026-11-24 SWMM 2611S
     'hdr': {'MARKS': 'MARKS', 'NO.OF PACKAGES': 'PKG', 'DESCRIPTION OF GOODS': 'DESC', 'GROSS WEIGHT': 'GW',
             'MEASUERMENT': 'MEAS', 'MEASUREMENT': 'MEAS'},
     'rx': {'CONS': r'^CNEE\s*:?\s*(.+)$', 'STATUS': r'^STATUS\s*:?\s*\**\s*(.+?)\s*\**$'},
     'stop': r'^(CNEE|CTN|STATUS|WE HEREBY)'},
    # GENIUS INTERFREIGHT — 'CONSIGNEE NAME :' กึ่งกลาง ชื่ออยู่บรรทัดล่าง ; ตาราง MARKS / NO.OF PACKAGE /
    # DESCRIPTION OF GOODS (2 คอลัมน์ย่อย x266 และ x407) / WEIGHT ('3,783.00' 'KGS' / '27.600' 'CBM' เป็นชิ้นแยก)
    {'name': 'GENIUS INTERFREIGHT', 'sig': r'GENIUS\s+INTERFREIGHT',          # 2026-11-25 SWMM 2611S
     'hdr': {},
     'bands': {'hdr': r'^MARKS$', 'cols': [(160, 'MARKS'), (255, 'PKG'), (430, 'DESC'), (1e9, 'GWM')]},
     'skip': r'^(\*+|\*+\s*\S*\s*ช.*|GENIUS INTERFREIGHT.*)$',
     'rx': {'GW': r'([\d,]+\.\d+)\s*KGS', 'MEAS': r'([\d,]+\.\d+)\s*CBM'},
     'stop': r'^(แจ้ง|CTNR|Thanks)'},
    {'name': 'HIPPO LOGISTICS', 'sig': r'HIPPO\s+LOGISTICS',          # 2026-11-03 DONGJIN CONFIDENT 0153S
     'hdr': {'MARKS': 'MARKS', 'NO.OF PKGS.': 'PKG', 'DESCRIPTION OF GOODS': 'DESC', 'GROSS WEIGHT': 'GW'},
     'rx': {'STATUS': r'^STATUS\s*:\s*(.+)$', 'CONS': r'^CONSIGNEE\s*:\s*(.+)$'},
     'stop': r'^(YOURS|THANK)'},
]


def _rows(lines, tol=4):
    """รวมบรรทัดที่อยู่ระดับเดียวกัน (|Δy| ≤ tol) เป็นแถวข้อความเดียว เรียงซ้าย→ขวา"""
    out = []
    for l in sorted(lines, key=lambda l: (l['y'], l['x'])):
        if out and abs(l['y'] - out[-1][0]) <= tol:
            out[-1][1].append(l)
        else:
            out.append((l['y'], [l]))
    return [' '.join(re.sub(r'\s+', ' ', o['t']).strip() for o in sorted(ls, key=lambda o: o['x'])) for _, ls in out]


def _apply_rx(pf, rec, body):
    rows = _rows(body)
    for k, pat in pf.get('rx', {}).items():
        v = next((m.group(1).strip() for r in rows for m in [re.search(pat, r, re.I)] if m), None)
        if not v:
            continue
        if k in ('GW', 'MEAS'):
            rec['gw' if k == 'GW' else 'meas'] = be.num(v)
        elif k == 'CONS':
            rec['cons'] = v
        elif k == 'STATUS':
            rec['status_raw'] = clean_status([v])
        elif k == 'PKG':
            rec['pkgs'], rec['pkgtype'] = split_pkg(v)
        elif k == 'MOVEMENT':                              # 'Cargo movement = 7' (Star Concord)
            rec['movement_raw'] = v.upper()


def _apply_spaced(pf, rec, body):
    labels = {k.upper(): v for k, v in pf['hdr'].items()}
    split = lambda t: [s.strip() for s in re.split(r'\s{2,}', t) if s.strip()]
    hl = next((l for l in body if sum(s.upper() in labels for s in split(l['t'])) >= 3), None)
    if not hl:
        return False
    cols = [labels[s.upper()] for s in split(hl['t']) if s.upper() in labels]
    parts = {}
    for l in sorted((o for o in body if o['y'] > hl['y'] + 3), key=lambda o: (round(o['y']), o['x'])):
        if re.match(pf['stop'], l['t'].strip(), re.I):
            break
        toks = split(l['t'])
        if len(toks) == 1 and l['x'] < hl['x'] + 30 and not parts.get('PKG'):
            parts.setdefault('MARKS', []).append(toks[0])     # MARKS บรรทัดแยก ชิดซ้าย
            continue
        for k, t in zip(cols[-len(toks):], toks):              # จับคู่จากขวา
            parts.setdefault(k, []).append(t)
    if parts.get('MARKS'):
        rec['marks'] = ' '.join(parts['MARKS'])
    if parts.get('DESC'):
        rec['desc'] = ' '.join(parts['DESC'])
    if parts.get('PKG'):
        rec['pkgs'], rec['pkgtype'] = split_pkg(parts['PKG'][0])
    for k, f in (('GW', 'gw'), ('MEAS', 'meas')):
        n = next((m.group(0) for t in parts.get(k, []) for m in [re.match(r'^[\d,]+(?:\.\d+)?', t)] if m), None)
        if n:
            rec[f] = be.num(n)
    return True


def _assign_parts(rec, parts):
    """ค่าที่แยกตามคอลัมน์ของโปรไฟล์ → ช่องของ rec"""
    if parts.get('CONS'):
        cl = []
        for t in parts['CONS']:                      # ชื่อตัดคำท้ายบรรทัด ('…SYSTEM' + 'S') ต่อกลับไม่เว้นวรรค
            if cl and len(t) <= 2 and t.isalpha():
                cl[-1] += t
            elif cl and re.fullmatch(r'\((?:HEAD\s*OFFICE|BRANCH[^)]*|สำนักงานใหญ่)\)', t, re.I):
                cl[-1] += t                          # '…CO., LTD.' + '(HEAD OFFICE)' (Sahathai, SWDN 2026-11-27)
            else:
                cl.append(t)
        rec['cons'] = company_name(cl)
    if parts.get('MARKS'):
        rec['marks'] = ' '.join(parts['MARKS'])
    if parts.get('DESC'):
        rec['desc'] = ' '.join(parts['DESC'])
    if parts.get('PKG'):
        rec['pkgs'], rec['pkgtype'] = split_pkg(' '.join(parts['PKG']))
    for k, f in (('GW', 'gw'), ('MEAS', 'meas')):
        n = next((m.group(0) for t in parts.get(k, []) for m in [re.match(r'^[\d,]+(?:\.\d+)?', t.strip())] if m), None)
        if n:
            rec[f] = be.num(n)


def _profile_for(lines):
    txt = ' '.join(l['t'] for l in lines)
    return next((pf for pf in FORM_PROFILES if re.search(pf['sig'], txt, re.I)), None)


def apply_profile(rec, lines, annots):
    """อ่านตารางสินค้าตามโปรไฟล์ของเจ้านั้น แล้วเขียนทับค่าที่ตัวอ่านจดหมายทั่วไปอ่านได้"""
    pf = _profile_for(lines)
    if not pf:
        return rec
    body = [l for l in lines if not in_annot(l, annots)]
    if pf.get('status'):
        st = next((m.group(1) for l in body for m in [re.search(pf['status'], l['t'].strip(), re.I)] if m), None)
        if st:
            rec['status_raw'] = clean_status([st])
    if pf.get('mode') == 'spaced':
        if _apply_spaced(pf, rec, body):
            rec['combined_text'] = False
            rec['profile'] = pf['name']
        return rec
    if pf.get('bands'):
        # หัวตารางพิมพ์เป็นข้อความบรรทัดเดียว (SINTOP) — แบ่งคอลัมน์ด้วยขอบซ้าย x ที่กำหนดในโปรไฟล์
        bd = pf['bands']
        hl0 = next((l for l in sorted(body, key=lambda l: l['y']) if re.match(bd['hdr'], l['t'].strip(), re.I)), None)
        if hl0:
            parts = {}
            for l in sorted((o for o in body if o['y'] > hl0['y'] + 3), key=lambda o: (round(o['y']), o['x'])):
                if re.match(pf['stop'], l['t'].strip(), re.I):
                    break
                if pf.get('skip') and re.match(pf['skip'], l['t'].strip(), re.I):
                    continue
                k = next(k for lim, k in bd['cols'] if l['x'] < lim)
                parts.setdefault(k, []).append(l['t'].strip())
            _assign_parts(rec, parts)
            rec['combined_text'] = False
            rec['profile'] = pf['name']
        _apply_rx(pf, rec, body)
        return rec
    # ช่องว่างซ้อนไม่นับ: PyMuPDF บนเว็บ (1.26) อ่าน 'Gross Weigh  Volume' มี 2 ช่องว่าง ส่วนในเครื่อง (1.28) มี 1 (2026-11-27)
    _sp = lambda s: re.sub(r'\s+', ' ', s).strip().upper()
    labels = {_sp(k): v for k, v in pf['hdr'].items()}
    hdr, hl = {}, []
    for l in sorted(body, key=lambda l: l['y']):
        k = labels.get(_sp(l['t']))
        if k and (not hl or l['y'] - hl[0]['y'] < 30):      # หัวตาราง 1–2 ชั้นแรกเท่านั้น
            hl.append(l)
            hdr.setdefault(k, l)
    if len(hdr) < pf.get('min_hdr', 3):
        _apply_rx(pf, rec, body)
        return rec
    hy = max(h['y'] for h in hl)
    cols = sorted(((h['x'] + h['x1']) / 2, k) for k, h in hdr.items())
    cuts = [((a[0] + b[0]) / 2, a[1]) for a, b in zip(cols, cols[1:])] + [(1e9, cols[-1][1])]
    parts = {k: [] for _, k in cols}
    last_y = hy
    for l in sorted((o for o in body if o['y'] > hy + 3), key=lambda o: (round(o['y']), o['x'])):
        if re.match(pf['stop'], l['t'], re.I):
            break
        if pf.get('gap') and l['y'] - last_y > pf['gap']:   # แถวยอดรวมที่อยู่ห่างลงไป (CK LINE)
            break
        last_y = l['y']
        c = (l['x'] + l['x1']) / 2
        k = next(k for lim, k in cuts if c < lim)
        t = l['t'].strip()
        if pf.get('skip') and re.match(pf['skip'], t, re.I):   # ป้าย/ชิ้นตัวเลขที่ rx อ่านแทน (NIPPON 'PKG:' '12' 'KG')
            continue
        if re.fullmatch(r'[\d,]+\.?\d*\s*(?:CBMS?|M3)\.?', t, re.I):     # ค่าที่มีหน่วยกำกับ ไม่ว่าอยู่คอลัมน์ไหน
            k = 'MEAS'
        elif re.fullmatch(r'[\d,]+\.?\d*\s*KGS?\.?', t, re.I):
            k = 'GW'
        elif k == 'QTYDESC':
            pk = split_pkg(t)
            k = 'PKG' if pk[0] and PKG_WORD_RE.match(pk[1] or '') and not parts.get('PKG') else 'DESC'
        elif k == 'BL':                                   # 'HASLK00000000000 ABC PHARMA' = B/L + marks บรรทัดแรก
            t = BL_IN_TXT_RE.sub('', t).strip()
            k = 'MARKS'
            if not t or CONT2_RE.search(t) or re.match(r'^(CNTR\b|H?B\s*/?\s*L\s*:?$)', t, re.I):   # 'CNTR :' / 'HLHU…/40'HQ' / 'HBL :'
                continue
        if k == 'MARKS' and not parts.get('PKG'):        # 'ABC FOOD … LIMITED 1,234 CARTONS' = marks + จำนวน
            mp = re.match(r'^(.*\S)\s+(\d[\d,]*\s+\S+)$', t)
            if mp and PKG_WORD_RE.match(split_pkg(mp.group(2))[1] or ''):
                parts.setdefault('MARKS', []).append(mp.group(1))
                parts.setdefault('PKG', []).append(mp.group(2))
                continue
        parts.setdefault(k, []).append(t)
    _assign_parts(rec, parts)
    _apply_rx(pf, rec, body)
    rec['combined_text'] = False
    rec['profile'] = pf['name']
    return rec


# ═══════════════ KUEHNE + NAGEL 'CORRECTION MANIFEST' (ตัวอักษรความกว้างเท่ากัน คอลัมน์คั่นด้วยช่องว่างยาว) ═══════════════
# หน้า 1: 'MBL NO. : HASL…', ป้าย CONSIGNEE + ชื่อบรรทัดล่าง, 'MOVEMENT : CY/CY', แถว '4 40' HC  AS PER ATTACHED  33223.217  264.396'
# หน้า ATTACHMENT: ทีละตู้ — 'SKHU0000000  1 40' HC  CONTAINER SAID TO CONTAIN:  8214.350  66.514' /
#   'SEAL NO. …  1730 CARTON(S)' / 'ABC  ABC ARTICLES' / 'HS CODE: …' ; 'TOTAL: 13252 CARTONS' ; แถวรวม '4 40' HC … น้ำหนัก ปริมาตร'
# (SAWASDEE MIMOSA 2611S, HASLS00000000000 — ผู้ใช้ชี้ตำแหน่ง 2026-11-25)
def _is_kn(lines):
    ns = nospace(' '.join(l['t'] for l in lines))
    return 'CORRECTIONMANIFEST' in ns and ('KNACCOUNTINGNO' in ns or 'KUEHNE+NAGEL' in ns)


def parse_kn_page(lines, annots, page, last, put):
    ns = nospace(' '.join(l['t'] for l in lines))
    rows = [re.split(r'\s{2,}', l['t'].strip()) for l in sorted(lines, key=lambda l: (round(l['y']), l['x']))]
    bl = None
    for l in lines:
        if re.match(r'^MBL\s*NO\.?$', l['t'].strip(), re.I):
            v = next((o['t'] for o in lines if abs(o['y'] - l['y']) < 3 and o['x'] > l['x'] + 20), '')
            mb = BL_IN_TXT_RE.search(v.replace(' ', ''))
            bl = mb.group(0) if mb else bl
    if bl:                                         # ใบหลัก
        rec = blank(bl, page)
        cl = next((l for l in lines if l['t'].strip().upper() == 'CONSIGNEE'), None)
        if cl:
            nxt = min((o for o in lines if 0 < o['y'] - cl['y'] < 30 and abs(o['x'] - cl['x']) < 10),
                      key=lambda o: o['y'], default=None)
            rec['cons'] = nxt['t'].strip() if nxt else ''
        mv = next((o['t'] for l in lines if re.match(r'^MOVEMENT$', l['t'].strip(), re.I)
                   for o in lines if abs(o['y'] - l['y']) < 3 and o['x'] > l['x'] + 20), '')
        rec['status_raw'] = clean_status([mv.strip(' :')]) if mv.strip(' :') else ''
        tot = next((r for r in rows if len(r) >= 3 and re.match(r"^\d+\s+\d{2}'", r[0])
                    and re.fullmatch(r'[\d,.]+', r[-1]) and re.fullmatch(r'[\d,.]+', r[-2])), None)
        if tot:
            rec['gw'], rec['meas'] = be.num(tot[-2]), be.num(tot[-1])
        rec['combined_text'] = False
        put({bl: finish(rec)})
        return rec
    if last is not None and 'ATTACHMENT' in ns:     # หน้าแนบของใบหลักก่อนหน้า
        conts, marks, desc = [], [], []
        for r in rows:
            for c in r:
                c = c.strip()
                if CONT2_RE.fullmatch(c):
                    conts.append(c)
        in_tbl = False
        for r in rows:
            if r[0].upper().startswith('MARKS & NOS'):
                in_tbl = True
                continue
            if not in_tbl or r[0].upper().startswith('TOTAL'):
                if r[0].upper().startswith('TOTAL'):
                    in_tbl = False
                continue
            first = r[0].strip()
            if len(r) >= 2 and not CONT2_RE.fullmatch(first) and not re.match(r'^(SEAL\s*NO|HS\s*CODE)', first, re.I):
                marks.append(first)
            for c in r[1:]:
                c = c.strip()
                if re.fullmatch(r"[\d,.]+|\d+\s+\d{2}'\s*\w+|CONTAINER SAID TO CONTAIN:?|\d[\d,]*\s+\w+\(S\)", c, re.I):
                    continue
                desc.append(c)
        mt = next((re.search(r'TOTAL\s*:\s*([\d,]+\s+\w+)', l['t'], re.I) for l in lines if re.match(r'^TOTAL\s*:', l['t'], re.I)), None)
        tot = next((r for r in rows if len(r) >= 3 and re.match(r"^\d+\s+\d{2}'", r[0])
                    and re.fullmatch(r'[\d,.]+', r[-1]) and re.fullmatch(r'[\d,.]+', r[-2])), None)
        last['cont'] = ';'.join(dict.fromkeys(conts_in(*[c for c in conts]))) or last['cont']
        last['marks'] = ' '.join(dict.fromkeys(marks))
        last['desc'] = ' '.join(dict.fromkeys(desc))
        if mt:
            last['pkgs'], last['pkgtype'] = split_pkg(mt.group(1))
        if tot:
            last['gw'], last['meas'] = be.num(tot[-2]), be.num(tot[-1])
        last['pages'] = sorted(set(last['pages'] + [page]))
        finish(last)
        return None
    return last


def _letter_blocks(lines):
    """จดหมายที่มีหลาย B/L ในหน้าเดียว (แต่ละ B/L มีป้ายผู้รับของตัวเอง) → แบ่งเป็นช่วงตาม y ของบรรทัดแรกของแต่ละ B/L
       (SAWASDEE MIMOSA 2611S หน้า 20: 'ENTER B/L …5764' … '____' … 'ENTER B/L …5764A'); หน้าปกติคืน [lines]"""
    first = {}
    for l in sorted(lines, key=lambda l: l['y']):
        for b in BL_IN_TXT_RE.findall(l['t'].upper()):
            first.setdefault(b, l['y'])
    n_cons = sum(1 for l in lines if re.match(rf'^{_L_CONS}\s*[:=]', l['t'], re.I))
    if len(first) < 2 or n_cons < 2:
        return [lines]
    ys = sorted(first.values())
    cuts = [y for y in ys[1:] if y - ys[0] > 60]
    if not cuts:
        return [lines]
    edges = [-1e9] + [c - 2 for c in cuts] + [1e9]
    return [[l for l in lines if a <= l['y'] < b] for a, b in zip(edges, edges[1:])]


def parse_letter(lines, annots, page, allow_no_bl=False):
    body = [l for l in lines if not in_annot(l, annots)]
    txt = ' '.join(l['t'] for l in body)
    bls = BL_IN_TXT_RE.findall(txt)
    if not bls:
        if not (allow_no_bl and re.search(rf'\b{_L_CONS}\s*[:=]', txt, re.I)):
            return {}
        bls = ['']
    bl0 = bls[0]
    # B/L ย่อย: ลูกค้าพิมพ์ตัวอักษรเดียวสีแดง (FreeText) ต่อท้ายเลข B/L บนบรรทัดเดียวกัน ('HASLS…966' + 'A')
    bl_line = next((l for l in body if bl0 in nospace(l['t'])), None)
    if bl_line:
        suf = next((a['t'] for a in annots if re.fullmatch(r'[A-Z]', a['t'])
                    and abs(a['rect'].y0 - bl_line['y']) < 8 and a['rect'].x0 > bl_line['x']), None)
        if suf:
            bl0 += suf
    rec = blank(bl0, page)
    if re.search(r'ATTACHED\s+SHEET\s*\(\s*DESCRIPTION\s*\)', txt, re.I):
        # หน้ารายละเอียดแนบ (NIPPON) — เอาเฉพาะบรรทัดใต้หัวเอกสาร (ป้าย M-B/L / HOUSE-B/L / M/M) เป็น DESC
        hy = max((l['y'] for l in body if re.match(r'^(HOUSE-B/L NO|M/M|M-B/L NO|W/T|PKG)\s*:?', l['t'], re.I)), default=0)
        rec['desc'] = ' '.join(l['t'] for l in sorted(body, key=lambda l: (round(l['y']), l['x'])) if l['y'] > hy + 10)
        rec['attached_desc'] = True
        return {rec['bl']: rec}

    cons, _ = _label_value(body, _L_CONS)
    rec['cons'] = cons
    rec['as_per_attached'] = bool(re.search(r'AS\s*PER\s*ATTACHED', cons, re.I))
    rec['cont'] = ';'.join(conts_in(txt, *[a['t'] for a in annots]))

    # ตารางที่มีหัว MARKS ... DESCRIPTION บนแถวเดียวกัน
    # หัวตารางจริงไม่มี ':' (ต่างจากป้าย 'Marks and Number : ABC' ที่มีค่าต่อท้าย)
    hm = next((l for l in body if re.match(_L_MARKS, l['t'], re.I) and ':' not in l['t'] and
               any(re.search(r'DESCRIP', o['t'], re.I) and ':' not in o['t'] and abs(o['y'] - l['y']) < 8
                   for o in body)), None)
    if hm:
        hrow = [o for o in body if abs(o['y'] - hm['y']) < 10]
        hx = {}
        for o in hrow:
            u = o['t'].upper()
            if re.match(_L_MARKS, u):
                hx.setdefault('MARKS', o['x'])
            if re.search(r'NUMBER$|NO\.?\s*OF\s*P|PKGS|PACKAGES', u) and not re.match(_L_MARKS, u):
                hx.setdefault('PKG', o['x'])
            if 'DESCRIP' in u and not re.match(_L_MARKS, u):
                hx.setdefault('DESC', o['x'])
            if 'WEIGHT' in u and not re.match(_L_MARKS, u):
                hx.setdefault('GW', o['x'])
        single = len(hrow) == 1          # หัวตารางพิมพ์เป็นข้อความบรรทัดเดียว (ไม่รู้ตำแหน่งคอลัมน์)
        rows, last_y = [], max(o['y'] for o in hrow)
        for l in sorted((o for o in body if o['y'] > last_y + 3), key=lambda o: (o['y'], o['x'])):
            if rows and l['y'] - max(r['y'] for r in rows) > 45:
                break
            # หัวตารางบรรทัดที่ 2 ('WEIGHT(KGS.)  MEAS.(CBM)' / 'GOODS  (KGS)  (M3)') และเศษจุด '.'
            if not re.search(r'[^\W_]', l['t']) or (l['y'] - last_y < 25 and all(re.fullmatch(   # (บรรทัดภาษาไทยไม่ใช่เศษ)
                    r'WEIGHT\S*|MEAS\S*|GOODS|PACKAGES?|\((?:KGS?|M3|CBM)\)', w, re.I) for w in l['t'].split())):
                continue
            # บรรทัดขึ้นต้น ':' = ค่าของป้ายทางซ้าย (': ABC TRADING …' ของ 'CNEE' — y น้อยกว่าป้ายเล็กน้อยจึงมาก่อน)
            if _LETTER_STOP.match(l['t']) or is_status_line(l['t']) or (rows and l['t'].startswith(':')):
                if rows:
                    break
                continue
            rows.append(l)
        marks, desc, gwm, nums_only = [], [], [], []
        pkg_l = next((r for r in rows if PKG_WORD_RE.match(split_pkg(r['t'])[1] or '') or
                      (re.fullmatch(r'\d[\d,]*', r['t']) and 'PKG' in hx and abs(r['x'] - hx['PKG']) < 40)), None)
        if pkg_l is None:
            # ตัวเลขล้วนที่มีชนิดหีบห่ออยู่บรรทัดล่าง ('5,251' / 'CARTONS', '19' / 'Pallets') — หัวคอลัมน์จำนวน
            # หาตำแหน่งไม่ได้ (พิมพ์รวมบรรทัดเดียวกับ MARKS หรือแยก 2 บรรทัด 'NO.OF' / 'PACKAGES') — SWMM 2611S
            pkg_l = next((r for r in rows if re.fullmatch(r'\d[\d,]*', r['t'].strip()) and any(
                PKG_WORD_RE.fullmatch(o['t'].strip()) and 0 < o['y'] - r['y'] < 25 and abs(o['x'] - r['x']) < 40
                for o in rows)), None)
        if pkg_l:
            rec['pkgs'], rec['pkgtype'] = split_pkg(pkg_l['t'])
            if not rec['pkgtype']:
                pw = next((r for r in rows if PKG_WORD_RE.fullmatch(r['t'].strip()) and abs(r['x'] - pkg_l['x']) < 40), None)
                rec['pkgtype'] = pw['t'].upper() if pw else ''
        pkg_x = pkg_l['x'] if pkg_l else hx.get('PKG', 200)
        desc_x = hx.get('DESC')
        gw_x = hx.get('GW')
        for r in rows:
            t = r['t']
            if r is pkg_l or (PKG_WORD_RE.fullmatch(t.strip()) and abs(r['x'] - pkg_x) < 40):
                continue
            if re.fullmatch(r'[\d,]+\.\d+', t):
                nums_only.append(r)
            elif re.fullmatch(r'(KG[SM]?|CBM|M3)\.?', t.strip(), re.I):
                continue                              # หน่วยแยกบรรทัด — จับคู่กับตัวเลขด้านล่าง
            elif re.search(r'\d\s*(?:KG[SM]?|CBMS?|M3)\b\.?$', t, re.I) or (gw_x and r['x'] >= gw_x - 30) \
                    or re.match(r'^(?:GROSS\s*(?:WT|WEIGHT)|G\.?\s*W|CBM|MEAS)\S*\s*[:=]', t, re.I):   # 'CBM. = 40.000'
                # (เฉพาะค่าที่ลงท้ายด้วยหน่วย — 'ANTIOXIDANT … IN 900KG (NET)' เป็นรายละเอียด ไม่ใช่น้ำหนัก)
                gwm.append(t)
            elif r['x'] < pkg_x - 5:
                marks.append(t)
            elif single or desc_x is None or r['x'] >= (pkg_x + desc_x) / 2:
                desc.append(t)
        rec['marks'] = ' '.join(marks)
        rec['desc'] = ' '.join(desc)
        # ตัวเลขล้วน ('10,774.40') กับหน่วยที่พิมพ์บรรทัดล่าง ('KGS') — จับคู่ตามคอลัมน์ (x ใกล้กัน)
        for r in nums_only:
            # หน่วยอยู่บรรทัดล่าง ('10,774.40' / 'KGS') หรือชิ้นถัดไปทางขวาในแถวเดียวกัน ('9,390.000' 'KGS.' — EXTRA MARITIME)
            unit = next((u for u in rows if re.fullmatch(r'(KG[SM]?|CBM|M3)\.?', u['t'].strip(), re.I)
                         and ((0 < u['y'] - r['y'] < 30 and abs(u['x'] - r['x']) < 25)
                              or (abs(u['y'] - r['y']) < 3 and 0 <= u['x'] - r['x1'] < 40))), None)
            gwm.append(f"{r['t']} {unit['t'] if unit else ''}".strip())
        g = ' '.join(gwm)
        mg = re.search(r'([\d,]+\.?\d*)\s*KG[SM]?', g, re.I)          # 'KGM.' (WISMA) เท่ากับ KGS
        mm = re.search(r'([\d,]+\.?\d*)\s*(?:CBM|M3)', g, re.I) or re.search(r'CBM\.?\s*=?\s*([\d,]+\.?\d*)', g, re.I)
        nums = re.findall(r'\d[\d,]*\.\d+', g)
        rec['gw'] = be.num(mg.group(1)) if mg else (be.num(nums[0]) if nums else None)
        rec['meas'] = be.num(mm.group(1)) if mm else (be.num(nums[1]) if len(nums) > 1 else None)
        # น้ำหนัก/ปริมาตรที่พิมพ์แบบมีป้ายใต้ตาราง ('GROSS WT. = 9,550.000 KGS.' / 'CBM. = …') ชนะค่าที่เดาจากตาราง
        for k, pat in (('gw', r'GROSS\s*(?:WEIGHT|WT\.?)\s*[:=]'), ('meas', r'CBM\.?\s*[:=]')):
            lv = next((m.group(1) for l in body for m in [re.search(pat + r'\s*([\d,]+\.?\d*)', l['t'], re.I)] if m), None)
            if lv:
                rec[k] = be.num(lv)
        if rec['gw'] is None:
            v, _ = _label_value(body, _L_GW)
            rec['gw'] = be.num(v) if v else None
        if rec['meas'] is None:
            v, _ = _label_value(body, _L_MEAS)
            rec['meas'] = be.num(v) if v else None
        if rec['pkgs'] is None:                       # 'Quantity : 1105 CARTONS' นอกตาราง (Masterwork)
            v, _ = _label_value(body, _L_PKG)
            pk = split_pkg(v) if v else (None, '')
            if pk[0]:
                rec['pkgs'], rec['pkgtype'] = pk
    else:
        v, _ = _label_value(body, _L_PKG)
        if v:
            rec['pkgs'], rec['pkgtype'] = split_pkg(v)
        v, _ = _label_value(body, _L_GW)
        if v:
            rec['gw'] = be.num(v)
            mm = re.search(r'([\d,]+\.?\d*)\s*(?:M3|CBM)', v, re.I)
            if mm:
                rec['meas'] = be.num(mm.group(1))
        if rec['meas'] is None:
            v, _ = _label_value(body, _L_MEAS)
            rec['meas'] = be.num(v) if v else None
        for key, pat in (('marks', _L_MARKS), ('desc', _L_DESC)):
            v, l = _label_value(body, pat)
            if l is not None:
                vx = l['x'] + 60
                more = []
                for o in sorted((o for o in body if l['y'] + 5 < o['y'] < l['y'] + 60 and o['x'] >= vx),
                                key=lambda o: o['y']):
                    if _LETTER_STOP.match(o['t']) or re.match(r'^[A-Z .&\'/-]+\s*:', o['t']) or o['t'].startswith(':') \
                            or any(p['x'] < vx and abs(p['y'] - o['y']) < 4 for p in body):
                        break                 # ค่าของป้ายถัดไป (มีป้ายอยู่ทางซ้ายแถวเดียวกัน)
                    more.append(o['t'])
                rec[key] = ' '.join([v] + more).strip()
    typed = [l['t'] for l in body if is_status_line(l['t']) and not re.match(_L_DESC, l['t'], re.I)]
    rec['status_raw'] = clean_status(typed + [p for a in annots for p in a['lines'] if is_status_line(p)])
    rec['combined_text'] = True           # ข้อความอิสระ — MARKS/DESC ไม่ตรงให้เป็นสีส้ม (ให้คนยืนยัน)
    for a in annots:
        rec['dg'] = rec['dg'] or be.find_dg(a['t'])
        rec['reefer'] = rec['reefer'] or find_reefer(a['t'])
    return {rec['bl']: finish(rec, *[a['t'] for a in annots])}


# ═══════════════ ฟอร์มที่ skill อ่านได้อยู่แล้ว: ส่งเฉพาะหน้านั้นให้ build_edi.py ═══════════════
def parse_shared(doc, pages, bl_re, man_bls=()):
    sub = pymupdf.open()
    for p in pages:
        sub.insert_pdf(doc, from_page=p - 1, to_page=p - 1, annots=True)
    fd, tmp = tempfile.mkstemp(suffix='.pdf')
    os.close(fd)
    try:
        sub.save(tmp)
        sub.close()
        ent, declared = be.parse_enter(tmp, bl_re)
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass
    page_bls = {}
    page_ns = {p: nospace(doc[p - 1].get_text()) for p in pages}
    for p in pages:
        for bl in BL_IN_TXT_RE.findall(doc[p - 1].get_text()):
            page_bls.setdefault(bl, []).append(p)
    man_set = set(man_bls)
    out = {}
    for bl, rec in ent.items():
        # เลข B/L บางฟอร์มพิมพ์แยก 2 บรรทัด ('…2381' + '3A') หรือรูปแบบอื่น จึงหาในข้อความหน้าไม่เจอตรง ๆ —
        # รับเมื่อเป็น B/L ของ MF หรือเลขฐาน (ตัดท้าย ≤3 ตัว) อยู่ในหน้า ; อย่างอื่นถือเป็นขยะจากการอ่าน
        hit = page_bls.get(bl) or [p for p in pages if any(bl[:len(bl) - k] in page_ns[p] for k in range(0, 4))]
        if not hit and bl not in man_set:
            continue
        rec['pages'] = sorted(set(hit or pages))
        out[bl] = rec
    return out, declared


# ═══════════════ อ่าน ENTER ทั้งชุด ═══════════════
def parse_enter_bundle(path, bl_re, man_bls=()):
    doc = pymupdf.open(path)
    kinds, L, A = [], [], []
    garbled = set()                       # หน้าที่ชั้นข้อความเป็น OCR ของเครื่องสแกน (เลข B/L มี '■')
    for i in range(doc.page_count):
        pg = doc[i]
        ls, an = page_lines(pg), page_annots(pg)
        if _fix_garbled_bl(ls, man_bls):
            garbled.add(i + 1)
            ls = _fix_an_labels(ls)
        L.append(ls)
        A.append(an)
        k = page_kind(pg, ls)
        if i + 1 not in garbled and _needs_ocr(k, ls, man_bls):
            k = 'SCAN'                    # ชั้นข้อความของสแกนอ่านไม่ได้ → อ่านภาพด้วย OCR แทน (DONGJIN CONFIDENT 0154S หน้า 10)
        kinds.append(k)
    direct = list(kinds)                 # ชนิดที่หน้านั้นแสดงเอง (ก่อนรวมหน้าต่อเนื่อง)
    join_forwarder_pages(kinds, L)

    ENT, LETTERS, scans = {}, {}, []
    info = {'kinds': kinds, 'scans': scans, 'groups': []}

    def put(recs):
        for bl, rec in recs.items():
            if bl in ENT:
                ENT[bl]['pages'] = sorted(set(ENT[bl]['pages'] + rec['pages']))
                for k, v in rec.items():
                    if k != 'pages' and v not in (None, '') and ENT[bl].get(k) in (None, ''):
                        ENT[bl][k] = v
            else:
                ENT[bl] = rec

    # กลุ่มหน้าที่ติดกันชนิดเดียวกัน
    groups, i = [], 0
    while i < len(kinds):
        j = i
        while j + 1 < len(kinds) and kinds[j + 1] == kinds[i] and kinds[i].startswith('SHARED') \
                and not (kinds[i] in DOC_HEADER_KINDS and direct[j + 1] == kinds[i]):
            j += 1
        groups.append((kinds[i], list(range(i + 1, j + 2))))
        i = j + 1
    info['groups'] = groups

    table_doc = None                      # ตาราง AMENDMENT หลายหน้า: STATUS/เลขตู้ระบุครั้งเดียวทั้งฉบับ
    table_docs = []
    pending_rider = None                  # (หน้า, rec) ใบแก้ไขที่รอรายละเอียดจากหน้า RIDER ถัดไป
    ocr_todo = []                         # หน้าที่ไม่มีตัวอักษร → อ่านด้วย OCR ตอนท้าย
    no_bl = []                            # จดหมายที่ไม่มีเลข B/L (ช่อง B/L NO. ว่าง)
    kn_last = None                        # KUEHNE + NAGEL: ใบหลักล่าสุด รอหน้า ATTACHMENT
    for kind, pages in groups:
        p = pages[0]
        ls, an = L[p - 1], A[p - 1]
        if kind.startswith('SHARED'):
            recs, _ = parse_shared(doc, pages, bl_re, man_bls)
            if kind == 'SHARED_DSV':
                _fix_dsv(recs, doc)
            if kind in dict(FORWARDER_PAGES):
                for r in recs.values():
                    # จำฟอร์มได้จากลายเซ็นแล้ว — ไม่ใช่ "ความมั่นใจต่ำ" แม้หน้านั้นมี B/L น้อยกว่า 3 ตัว (ไม่ให้แดงกลายเป็นส้ม)
                    r['low_confidence'] = False
                    # หน้าที่มี B/L เดียว: เลขตู้ทุกตัวที่พิมพ์บนหน้าเป็นของ B/L นี้ (Handle Inter พิมพ์ตู้ต่อบรรทัดใต้ 'Cont No :')
                    one = [q for q in r['pages'] if sum(1 for o in recs.values() if q in o['pages']) == 1]
                    extra = conts_in(*[doc[q - 1].get_text() for q in one])
                    if one and len(extra) > len(conts_in(r.get('cont') or '')):
                        r['cont'] = ';'.join(extra)
            # FreeText STATUS ที่หัวใบแรก (เหนือ B/L แรก) = คำสั่งของทั้งชุด ไม่ใช่แค่ B/L แรก
            top = min((l['y'] for l in ls if BL_TXT_RE.search(nospace(l['t']))), default=0)
            head_st = clean_status([q for a in an if a['rect'].y0 < top for q in a['lines'] if is_status_line(q)])
            if head_st:
                for r in recs.values():
                    if not r.get('status_raw'):
                        r['status_raw'] = head_st
                        r['pages'] = sorted(set(r['pages'] + [p]))      # หน้าที่มีคำสั่ง STATUS
            if kind == 'SHARED_CMS':
                # บล็อกคร่อมหน้า: สินค้าอยู่ท้ายหน้าก่อน แต่ D/O No. (เลข B/L) อยู่หัวหน้าถัดไป — ใส่หน้าก่อนด้วย
                for r in recs.values():
                    q = min(r['pages'])
                    ql = L[q - 1]
                    do_y = next((l['y'] for l in ql if r['bl'] in nospace(l['t'])), None)
                    if q > pages[0] and do_y is not None and not any('KGS' in l['t'] and l['y'] < do_y for l in ql):
                        r['pages'] = sorted(set(r['pages'] + [q - 1]))
            put(recs)
            continue
        if kind != 'TABLE':
            table_doc = None
        if kind == 'BLFORM':
            recs = parse_blform(ls, an, p, boxes=red_boxes(doc[p - 1]))
            if p in garbled:
                for r in recs.values():
                    r['ocr'] = True               # ค่าจากชั้นข้อความ OCR ของสแกน — ขึ้นสีส้ม ตรวจด้วยตา
                    r['low_confidence'] = True
            put(recs)
        elif kind == 'RIDER':
            # ใบแก้ไขตามโปรไฟล์ที่ DESC = 'AS PER ATTACHED DESCRIPTIONS' → รายละเอียดจริงคือหน้า RIDER ที่ตามมาทันที
            if pending_rider and pending_rider[0] == p - 1:
                prec = pending_rider[1]
                top = max((l['y'] for l in ls if re.search(r'Vessel/Voy|B/L No\.?:|ATTACHED RIDER', l['t'], re.I)), default=0)
                rl = [l['t'] for l in sorted(ls, key=lambda l: (l['y'], l['x'])) if l['y'] > top + 3
                      and not in_annot(l, an) and not re.match(r'^(Page\s*:|\*\s*INVOICE)', l['t'], re.I)]
                if rl:
                    prec['desc_attached_text'] = prec['desc']
                    prec['desc'] = ' '.join(rl)
                    prec['pages'] = sorted(set(prec['pages'] + [p]))
                pending_rider = None
                continue
            bl, marks, typed, rdesc = parse_rider(ls, p)
            if bl:
                ENT.setdefault(bl, blank(bl, p))
                r = ENT[bl]
                r['pages'] = sorted(set(r['pages'] + [p]))
                if marks:
                    r['rider_marks'] = marks
                if rdesc:
                    r['rider_desc'] = rdesc
                if typed and not r['status_raw']:
                    r['status_raw'] = clean_status(typed)
        elif kind == 'MARKNO':
            put(parse_markno(ls, p))
        elif kind == 'TABLE':
            is_head = any('WE WOULD LIKE TO REQUEST' in l['t'].upper() or l['t'].upper() == 'AMENDMENT' for l in ls)
            if is_head or table_doc is None:
                table_doc = {'rows': {}, 'status': [], 'conts': [], 'st_pages': [], 'ct_pages': []}
                table_docs.append(table_doc)
            rows, st, cs = parse_table_page(ls, an, p)
            if len(rows) == 1 and _profile_for(ls):        # ฟอร์มที่จดโปรไฟล์ไว้ (เช่น UNICORN) — 1 B/L ต่อหน้า
                apply_profile(next(iter(rows.values())), ls, an)
            table_doc['rows'].update(rows)
            table_doc['status'] += st
            table_doc['conts'] += [c for c in cs if c not in table_doc['conts']]
            if st:
                table_doc['st_pages'].append(p)
            if cs:
                table_doc['ct_pages'].append(p)
        elif kind == 'LETTER' and _is_kn(ls):
            kn_last = parse_kn_page(ls, an, p, kn_last, put)
        elif kind == 'LETTER':
            recs = {}
            for blk in _letter_blocks(ls):        # จดหมายเดียวหลาย B/L (ENTER B/L … / ____ / ENTER B/L …A)
                recs.update(parse_letter(blk, an, p))
            if not recs:
                nr = parse_letter(ls, an, p, allow_no_bl=True)
                if nr:
                    no_bl.append(next(iter(nr.values())))   # จดหมายที่ลูกค้าไม่ได้ใส่เลข B/L — จับคู่กับ MF ตอนท้าย
                else:
                    # หน้า 2 ของจดหมาย ('Page 2 of 2', ไม่มีเลข B/L) — STATUS/เลขตู้ในหน้านั้นเป็นของ B/L หน้าก่อน
                    mpg = re.search(r'Page\s+(\d+)\s+of\s+\d+', ' '.join(l['t'] for l in ls), re.I)
                    prev = next((r for r in list(LETTERS.values()) + list(ENT.values()) if p - 1 in r['pages']), None)
                    if mpg and int(mpg.group(1)) > 1 and prev is not None:
                        prev['pages'] = sorted(set(prev['pages'] + [p]))
                        st = clean_status([l['t'] for l in ls if is_status_line(l['t'])]
                                          + [q for a in an for q in a['lines'] if is_status_line(q)])
                        if st and not prev.get('status_raw'):
                            prev['status_raw'] = st
                        cs = conts_in(*[l['t'] for l in ls], *[a['t'] for a in an])
                        if cs and not prev.get('cont'):
                            prev['cont'] = ';'.join(cs)
                        continue
            for bl, rec in recs.items():
                if rec.get('attached_desc'):
                    # หน้ารายละเอียดแนบของ B/L ที่อ่านไปแล้ว: แทน 'AS PER ATTACHED' — ไม่เขียนทับใบหลักทั้งใบ
                    lr = LETTERS.get(bl)
                    if lr is not None:
                        if not lr.get('desc') or re.search(r'AS\s*PER\s*ATTACHED', lr['desc'], re.I):
                            lr['desc'] = rec['desc']
                        lr['pages'] = sorted(set(lr['pages'] + rec['pages']))
                        continue
                    rec.pop('attached_desc', None)
                apply_profile(rec, ls, an)
                if bl in LETTERS:              # B/L เดียวกันหลายหน้า: เติมช่องที่ยังว่าง ไม่ทับของเดิม
                    lr = LETTERS[bl]
                    lr['pages'] = sorted(set(lr['pages'] + rec['pages']))
                    for k, v in rec.items():
                        if k != 'pages' and v not in (None, '') and lr.get(k) in (None, ''):
                            lr[k] = v
                    continue
                LETTERS[bl] = rec
                if rec.get('profile') and re.match(r'AS\s*PER\s*ATTACHED', rec.get('desc') or '', re.I):
                    pending_rider = (p, rec)
            if not recs and sum(not in_annot(l, an) for l in ls) < 5:
                ocr_todo.append(p)        # ข้อความถูกแปลงเป็นเส้น: เหลือแต่ FreeText ที่ลูกค้าพิมพ์ทับ
        elif kind == 'SCAN':
            scans.append({'page': p, 'conts': conts_in(*[a['t'] for a in an])})
            ocr_todo.append(p)

    for td in table_docs:
        for bl, rec in td['rows'].items():
            # STATUS/เลขตู้ที่ระบุครั้งเดียวทั้งฉบับ — ใส่หน้าที่มาด้วย เพื่อย้อนตรวจได้
            if not rec['status_raw'] and td['status']:
                rec['status_raw'] = ' '.join(dict.fromkeys(td['status']))
                rec['pages'] = sorted(set(rec['pages'] + td['st_pages']))
            if not rec['cont'] and td['conts']:
                rec['cont'] = ';'.join(td['conts'])
                rec['pages'] = sorted(set(rec['pages'] + td['ct_pages']))
            put({bl: finish(rec)})

    # ATTACHED RIDER: แทนคำ 'AS PER ATTACHED RIDER' ใน MARKS ด้วยเครื่องหมายจริงจาก RIDER
    for rec in ENT.values():
        rm = rec.pop('rider_marks', '')
        if rm:
            if re.search(r'AS\s*PER\s*ATTACHED\s*RIDER', rec['marks'], re.I):
                rec['marks'] = re.sub(r'AS\s*PER\s*ATTACHED\s*RIDER', rm, rec['marks'], flags=re.I).strip()
            elif not rec['marks']:
                rec['marks'] = rm
        rd = rec.pop('rider_desc', '')
        if rd and (not rec['desc'] or re.search(r'AS\s*PER\s*ATTACHED', rec['desc'], re.I)):
            rec['desc'] = rd

    # จดหมาย: B/L ที่มีฟอร์มโครงสร้างอยู่แล้ว ใช้จดหมายเติมเฉพาะช่องว่าง (STATUS ฯลฯ);
    # จดหมาย 'CONSIGNEE AS PER ATTACHED SHEET' ส่ง STATUS ให้ B/L ย่อยทุกตัวของ B/L นั้นด้วย
    for bl, lr in LETTERS.items():
        sibs = [b for b in ENT if b == bl or (lr.get('as_per_attached') and re.fullmatch(re.escape(bl) + r'[A-Z]?', b))]
        if sibs:
            for b in sibs:
                e = ENT[b]
                e['pages'] = sorted(set(e['pages'] + lr['pages']))
                if not e['status_raw'] and lr['status_raw']:
                    e['status_raw'] = lr['status_raw']
                if b == bl:
                    for k in ('cont', 'dg', 'reefer', 'transit_raw'):
                        if not e.get(k) and lr.get(k):
                            e[k] = lr[k]
        else:
            lr.pop('as_per_attached', None)
            ENT[bl] = lr
    for rec in ENT.values():
        rec.pop('as_per_attached', None)

    # OCR: เฉพาะ B/L ที่ยังไม่มีข้อความ ENTER จริง — ไม่เขียนทับค่าที่อ่านจากข้อความ
    info['ocr'] = []
    if ocr_todo and man_bls:
        OCR = {}
        for p in ocr_todo:
            ol = ocr_page(doc[p - 1])
            pool = {b: (man_bls[b] if isinstance(man_bls, dict) else {}) for b in man_bls}
            bl = _ocr_bl(ol, pool)                     # เทียบกับ MF ทุก B/L แล้วค่อยดูว่ามี ENTER จากข้อความแล้วหรือยัง
            done = bl if bl in ENT else None
            if bl in ENT:
                bl = None
            if not bl:
                # สำเนาภาพของ B/L ที่อ่านจากข้อความได้แล้ว (ARRIVAL NOTICE ที่ลูกค้าเขียนกำกับแล้วสแกนแนบ)
                # → ใส่เลขหน้าให้ B/L นั้น ไม่นับเป็นหน้าที่อ่านไม่ได้ (ค่าใช้จากหน้าข้อความ)
                # หน้าต่อของเอกสารหลายหน้า ('2/3' 'MARKS (ต่อ)') — เป็นของ B/L ที่อยู่หน้าก่อน
                top = next((l['t'] for l in sorted(ol, key=lambda l: l['y'])), '')
                mpg = re.fullmatch(r'\s*(\d+)\s*/\s*(\d+)\s*', top)
                if not done and mpg and int(mpg.group(1)) > 1:
                    done = next((b for b, r in ENT.items() if p - 1 in r['pages']), None)
                if done:
                    ENT[done]['pages'] = sorted(set(ENT[done]['pages'] + [p]))
                    st = clean_status([a['t'] for a in A[p - 1] if is_status_line(a['t'])])
                    if st and not ENT[done].get('status_raw'):
                        ENT[done]['status_raw'] = st
                    info.setdefault('ocr_copy', []).append((p, done))
                    scans[:] = [s for s in scans if s['page'] != p]
                continue
            r = parse_ocr_page(ol, A[p - 1], p)
            if any(re.search(r'ATTACHED\s*RIDER', l['t'], re.I) for l in ol)                     and not any("SHIPPER'S LOAD" in l['t'].upper() for l in ol):
                # ATTACHED RIDER ที่เป็นภาพสแกน (SWDN 2026-11-27 หน้า 18) — รายละเอียดแทน 'AS PER ATTACHED DESCRIPTIONS'
                _rb, rmk, _rt, rdesc = parse_rider([dict(l, t=_ocr_num(l['t'])) for l in ol], p)
                if rdesc:
                    r['rider_desc'] = rdesc
                if rmk:
                    r['rider_marks'] = rmk
            if any("SHIPPER'S LOAD" in l['t'].upper() for l in ol):
                # ARRIVAL NOTICE สแกน: ใช้ตัวอ่าน BLFORM ตัวเดียวกับหน้าที่มีข้อความ (OCR มีพิกัดครบ)
                # — ใช้เมื่ออ่านจำนวนหีบห่อและน้ำหนักได้ (ภาพเอียงจะอ่านคอลัมน์ไม่ได้ → ใช้ตัวอ่าน OCR เดิม)
                rb = parse_blform(_ocr_blform_lines(ol, bl), A[p - 1], p, gap_cut=True).get(bl)
                if rb and rb.get('gw'):         # อ่านคอลัมน์น้ำหนักได้ = จับแนวฟอร์มได้ (จำนวนหีบห่ออาจว่างถ้าเลขเลอะ)
                    for k, v in r.items():      # ช่องที่ BLFORM ว่าง (เช่น STATUS CY/CY ที่พิมพ์บนฟอร์ม) เติมจากตัวอ่านเดิม
                        if k != 'pages' and v not in (None, '') and rb.get(k) in (None, ''):
                            rb[k] = v
                    r = rb
            info['ocr'].append((p, bl))
            scans[:] = [s for s in scans if s['page'] != p]     # OCR อ่านได้แล้ว — ไม่ใช่ "หน้าที่อ่านไม่ได้"
            if bl not in OCR:
                r['bl'] = bl
                OCR[bl] = r
                continue
            o = OCR[bl]                       # หน้าที่ 2 ของ B/L เดียวกัน (เช่น ATTACHED RIDER)
            o['pages'] = sorted(set(o['pages'] + [p]))
            o['cont'] = ';'.join(dict.fromkeys([c for c in (o['cont'] + ';' + r['cont']).split(';') if c]))
            for k, v in r.items():
                if k not in ('pages', 'bl') and v not in (None, '') and o.get(k) in (None, ''):
                    o[k] = v
        for bl, r in OCR.items():
            r['ocr'] = True
            r['low_confidence'] = True
            rd = r.pop('rider_desc', '')
            if rd and (not r.get('desc') or re.search(r'AS\s*PER\s*ATTACHED', r['desc'], re.I)):
                r['desc'] = rd
            rm = r.pop('rider_marks', '')     # หน้า ATTACHED RIDER ที่เป็นภาพสแกนด้วย
            if rm and re.search(r'AS\s*PER\s*ATTACHED\s*RIDER', r['marks'] or '', re.I):
                r['marks'] = re.sub(r'AS\s*PER\s*ATTACHED\s*RIDER', rm, r['marks'], flags=re.I).strip()
            elif rm and not r['marks']:
                r['marks'] = rm
            ENT[bl] = finish(r)

    # จดหมายที่ช่อง B/L NO. ว่าง: จับคู่กับ B/L ใน MF ที่ยังไม่มี ENTER — G.W. ตรง และ (จำนวนหีบห่อหรือชื่อผู้รับ) ตรง
    # ต้องได้ตัวเดียวเท่านั้น; แถวเป็นสีส้ม พร้อมหมายเหตุให้ตรวจด้วยตา (SAWASDEE MIMOSA 2611S หน้า 30)
    if no_bl and isinstance(man_bls, dict):
        def _n(v):
            try:
                return be.num(re.sub(r'[^\d.,]', '', str(v or '')) or None)
            except Exception:
                return None
        for r in no_bl:
            if r.get('gw') is None:
                continue
            cw = set(re.findall(r'[A-Z]{3,}', (r.get('cons') or '').upper())) - {'CO', 'LTD', 'COMPANY', 'LIMITED', 'THE'}
            cand = []
            for b, m in man_bls.items():
                if b in ENT or _n(m.get('gw')) is None or abs(_n(m.get('gw')) - r['gw']) > 0.5:
                    continue
                mw = set(re.findall(r'[A-Z]{3,}', str(m.get('cons') or '').upper()))
                mp = _n((str(m.get('pkg_hdr') or '').split() or [''])[0])       # '18 PX (PALLET(s))'
                if (r.get('pkgs') and mp == r['pkgs']) or len(cw & mw) >= 2:
                    cand.append(b)
            if len(cand) == 1:
                r['bl'] = cand[0]
                r['low_confidence'] = True
                r['cons_note'] = (r.get('cons_note', '') + ' ; ' if r.get('cons_note') else '') + \
                    'ENTER ไม่มีเลข B/L — จับคู่กับ MF จาก G.W./จำนวน/ผู้รับ ตรวจด้วยตา'
                ENT[cand[0]] = r
                info.setdefault('no_bl', []).append((r['pages'][0], cand[0]))
    return ENT, info


# ═══════════════ OCR สำรอง สำหรับหน้าที่ไม่มีตัวอักษร (ภาพสแกน / ข้อความถูกแปลงเป็นเส้น) ═══════════════
# ใช้ Windows.Media.Ocr ผ่าน win_ocr.ps1 (มีใน Windows 10/11 ไม่ต้องติดตั้ง). ค่าที่ได้ "ไม่น่าเชื่อถือเต็มที่"
# (ฟอร์ม Heung-A ใช้เลข 0 มีขีด → OCR อ่านเป็น Ø/ø/e) — B/L ที่ได้จาก OCR จึงขึ้นสีส้ม ตรวจด้วยตาเสมอ
OCR_PS = os.path.join(_HERE, 'win_ocr.ps1')
_OCR_DIGIT = str.maketrans({'Ø': '0', 'ø': '0', 'e': '0', 'o': '0', 'O': '0', 'D': '0', 'Q': '0',
                            'g': '9', 'l': '1', 'I': '1', '|': '1', 'B': '8'})
_OCR_PKG_RE = re.compile(r'\b(\d[\d,]*)\s+(CASES?|CARTONS?|CTNS?|PALLETS?|PLTS?|PKGS?|PACKAGES?|BAGS?|DRUMS?|'
                         r'ROLLS?|PIECES?|PCS|SETS?|BOX(?:ES)?|BUNDLES?|UNITS?)\b', re.I)


# หน้าเว็บ (Pyodide ในเบราว์เซอร์): ไม่มี OCR ของ Windows — รอบแรกจด "หน้าที่ต้อง OCR" ไว้ใน OCR_WANTED,
# หน้าเว็บ OCR ด้วย Tesseract.js แล้วใส่ผลใน BROWSER_OCR {เลขหน้า 0-based: [{x,y,x1,y1,t} พิกเซลที่ 300 dpi]} แล้วรันรอบสอง
BROWSER_OCR = None
OCR_WANTED = set()


def ocr_page(pg, dpi=300):
    """คืนบรรทัด OCR แบบเดียวกับ page_lines (พิกัด pt) — ไม่มี OCR (ไม่ใช่ Windows) คืน []"""
    if BROWSER_OCR is not None:
        raw = BROWSER_OCR.get(pg.number)
        if raw is None:
            OCR_WANTED.add(pg.number)
            return []
        k = 72 / dpi
        return [{'x': l['x'] * k, 'y': l['y'] * k, 'x1': l['x1'] * k, 't': l['t']} for l in raw if l.get('t')]
    if os.name != 'nt' or not os.path.exists(OCR_PS):
        return []
    import json
    import subprocess
    import hashlib
    # แคชผล OCR ตามเนื้อภาพ (ชุดใหญ่ 300 หน้า/สแกน 100 หน้า รันซ้ำไม่ต้อง OCR ใหม่) — อยู่ในโฟลเดอร์ temp ของเครื่อง
    pix = pg.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY)
    cdir = os.path.join(tempfile.gettempdir(), 'bundle_enter_ocr_cache')
    cfile = os.path.join(cdir, hashlib.md5(pix.samples).hexdigest() + '.json')
    raw = None
    if os.path.exists(cfile):
        try:
            with open(cfile, encoding='utf-8') as f:
                raw = json.load(f)
        except Exception:
            raw = None
    if raw is None:
        fd, png = tempfile.mkstemp(suffix='.png')
        os.close(fd)
        try:
            pix.save(png)
            r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', OCR_PS, png],
                               capture_output=True, timeout=120)
            raw = json.loads(r.stdout.decode('utf-8-sig') or '[]')
            try:
                os.makedirs(cdir, exist_ok=True)
                with open(cfile, 'w', encoding='utf-8') as f:
                    json.dump(raw, f, ensure_ascii=False)
            except OSError:
                pass
        except Exception:
            return []
        finally:
            try:
                os.remove(png)
            except OSError:
                pass
    k = 72 / dpi
    return [{'x': l['x'] * k, 'y': l['y'] * k, 'x1': l['x1'] * k, 't': l['t']} for l in raw if l.get('t')]


def _ocr_num(t):
    """แก้ตัวที่ OCR อ่านเลข 0 มีขีดผิด เฉพาะคำที่เป็นตัวเลขเป็นหลัก ('2eø.øeee' → '200.0000')"""
    def fix(w):
        if re.search(r'\d', w) and re.fullmatch(r'[\dØøeoODQgIl|B,./:\-]+', w):     # '2B.eeøe' = 28.0000
            return w.translate(_OCR_DIGIT)
        return w
    # ตัวเลขที่ OCR แยกด้วยช่องว่างหลังจุด/จุลภาค ('55. eoeo', '49, øee' — SWDN 2026-11-27) → ต่อกลับ
    t = re.sub(r'(\d[.,])\s+(?=[\dØøeoOD]{2,}(?![A-Za-z]))', r'\1', t)
    return ' '.join(fix(w) for w in t.split())


_OCR_LOOKALIKE = {frozenset(p) for p in ('08', '06', '68', '38', '56', '17', '09', '35')}


def _lev(a, b):
    """ระยะแก้ไข — ตัวเลขที่ OCR สับสนบ่อย (0/8, 6/8, 1/7 …) คิด 0.5 ('HASLC00000000000' ↔ HASLC00000000000)"""
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            sub = 0 if ca == cb else (0.5 if frozenset((ca, cb)) in _OCR_LOOKALIKE else 1)
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + sub))
        prev = cur
    return prev[-1]


def _ocr_bl(lines, man_bls):
    """เลข B/L จาก OCR ('HASLKØ226egee795') → B/L ใน MANIFEST ที่ใกล้ที่สุด (ต่างไม่เกิน 2 ตัว และไม่กำกวม)"""
    best = []
    for l in lines:
        # ตัวอักษรหลัง HASL อาจหาย ('HASL1022611øe793' = HASLK00000000000) หรือเลขถูกอ่านเป็น O/L ('HASLKO000LOO00L0')
        # ตัวพิมพ์เล็กก็ได้ ('HASLce1261e17675', 'HAsL01261ee4eee' — SAWASDEE MIMOSA 2611S)
        for m in re.finditer(r'HAS[Ll]([A-Za-z0-9])([0-9A-Za-zØø]{8,14})', l['t'].replace(' ', ''), re.I):
            g1 = m.group(1).upper() if m.group(1).isalpha() else m.group(1).translate(_OCR_DIGIT)
            cand = 'HASL' + g1 + m.group(2).translate(_OCR_DIGIT).upper().replace('L', '1')
            for b in man_bls:
                # OCR อาจอ่านเลข B/L ติดกับข้อความถัดไป ('HASLC0000000000THSC' = HASLC00000000000 + 'THS C…')
                d = min(_lev(cand, b), _lev(cand[:len(b)], b) + 0.5 if len(cand) > len(b) else 99)
                if d <= 2:
                    best.append((d, b))
    cset = set()
    if isinstance(man_bls, dict):
        # เลขตู้บนหน้า (OCR อ่านเลขตู้ได้ดีกว่าเลข B/L) → B/L ใน MF ที่มีตู้นั้น
        pconts = set(conts_in(_ocr_cont_fix(' '.join(_ocr_num(l['t']) for l in lines))))
        if pconts:
            cset = {b for b, m in man_bls.items() if pconts & set(re.findall(r'[A-Z]{3}[UJZ]\d{7}', str(m.get('cont') or '')))}
    if not best:
        # เลข B/L อ่านไม่ออก ('HASLce1261æe11e') แต่เลขตู้ชี้ B/L ได้ตัวเดียว (SAWASDEE MIMOSA 2611S หน้า 143)
        return cset.pop() if len(cset) == 1 else None
    dmin = min(d for d, _ in best)
    top = {b for d, b in best if d == dmin}
    if len(top) > 1 and cset and len(top & cset) == 1:
        return (top & cset).pop()
    if len(cset) == 1 and not (top & cset) and dmin > 1:      # เลข B/L ใกล้ตัวอื่นแบบหลวม ๆ แต่เลขตู้ชัด → เชื่อเลขตู้
        return cset.pop()
    if len(top) > 1 and isinstance(man_bls, dict):
        # เลขกำกวม: เลือกตัวที่ชื่อผู้รับใน MF ปรากฏบนหน้านั้น ('…1077' + 'THS CO., LTD.' = HASLC00000000000)
        page = re.sub(r'[^A-Z0-9]', '', ' '.join(l['t'] for l in lines).upper())
        key = lambda b: re.sub(r'[^A-Z0-9]', '', str(man_bls[b].get('cons') or '').upper())[:8]
        top = {b for b in top if len(key(b)) >= 4 and key(b) in page}
    return top.pop() if len(top) == 1 else None


def _ocr_cont_fix(t):
    """เลขตู้จาก OCR ที่เลข 0 มีขีดถูกอ่านเป็น e/ø ('JPHU25eø588' → 'JPHU0000000')"""
    return re.sub(r'([A-Z]{3}[UJZ])([0-9eoOØøDQIl|]{7})(?![0-9A-Za-z])',
                  lambda m: m.group(1) + m.group(2).translate(_OCR_DIGIT), t)


def _needs_ocr(kind, lines, man_bls):
    """ฟอร์มพิมพ์ของ Heung-A (ARRIVAL NOTICE / RIDER) ที่ชั้นข้อความเป็น OCR คุณภาพต่ำของเครื่องสแกน:
       ไม่มีเลข B/L ของ MF ตรงตัวเลย แต่มีเลขที่ใกล้ B/L ใน MF ตัวเดียว ('HASLKO000LOO00L0') → อ่านหน้านี้ด้วย OCR ใหม่
       (ฟอร์มพิมพ์จากระบบพิมพ์เลข B/L ผิดไม่ได้ จึงไม่ใช่กรณีลูกค้าพิมพ์เลขผิด)"""
    if kind not in ('BLFORM', 'RIDER') or not man_bls:
        return False
    ns = nospace(' '.join(l['t'] for l in lines))
    if any(b in ns for b in man_bls):
        return False
    return _ocr_bl(lines, man_bls) is not None


def _fix_garbled_bl(lines, man_bls):
    """ชั้นข้อความ OCR ของเครื่องสแกนมีตัวที่อ่านไม่ออกเป็น '■' ('HASLS0000■000000')
       → เทียบกับ B/L ใน MF ทีละตัว ('■' = ตัวใดก็ได้ 1 ตัว) ถ้าตรงตัวเดียวพอดีให้แทนเลขนั้น คืน True ถ้าแก้"""
    fixed = False
    for l in lines:
        m = re.search(r'HASL[A-Z0-9■]{8,14}', nospace(l['t']))
        if not m or '■' not in m.group(0):
            continue
        rx = re.compile(re.escape(m.group(0)).replace('■', '.') + r'$')
        hit = [b for b in man_bls if rx.match(b)]
        if len(hit) == 1:
            l['t'] = l['t'].replace(m.group(0), hit[0]) if m.group(0) in l['t'] else hit[0]
            fixed = True
    return fixed


def _ocr_blform_lines(ol, bl):
    """บรรทัด OCR ของ ARRIVAL NOTICE → รูปแบบที่ parse_blform อ่านได้: แก้เลข 0 มีขีด, เลข B/L = B/L ใน MF ที่จับคู่แล้ว,
       ป้ายช่องที่ OCR อ่านเพี้ยน ('Cons i onee', 'btlfy Party'), แยก 'PACKAGESi SHIPPER'S LOAD' ที่ติดกันเป็นสองชิ้น"""
    # ชิ้นที่อยู่แถวเดียวกันและห่างกันไม่ถึง 12pt = คำเดียวกันที่ OCR ตัดเป็นสองชิ้น ('ABC FOODS CO.' + 'LTD.')
    merged = []
    for l in sorted(ol, key=lambda l: (round(l['y']), l['x'])):
        p = merged[-1] if merged else None
        if p and abs(p['y'] - l['y']) < 2 and 0 <= l['x'] - p['x1'] < 12 and re.match(r'[A-Za-z(]', l['t']):
            # (ไม่รวมชิ้นที่ขึ้นต้นด้วยตัวเลข: '15,333.000 KGS' กับ '42.5000 CBM' เป็นคนละคอลัมน์)
            merged[-1] = dict(p, x1=l['x1'], t=p['t'] + ' ' + l['t'])
        else:
            merged.append(dict(l))
    out = []
    for l in merged:
        t = _ocr_num(l['t']).replace('Ø', '0').replace('ø', '0')     # 'R4ØH SKRU0000000'
        t = re.sub(r'(?<=[A-Z]{4})[a-z]\b', '', t)                     # 'PACKAGESi' → 'PACKAGES'
        t = _ocr_cont_fix(t)
        # เส้นแบ่งคอลัมน์ของฟอร์มที่ OCR อ่านเป็น ':' ท้าย/หน้าชิ้น ('24,016.000 KGS :', SWDN 2026-11-27)
        t = re.sub(r'^[:;|]\s+|\s+[:;|]$', '', t)
        if re.fullmatch(r'(?:COUNT\s*)?&\s*WEIGHT', t.strip(), re.I):   # ท้าย "SHIPPER'S LOAD, COUNT & WEIGHT" ที่ OCR ตัดแยก
            continue
        # จำนวนหีบห่อที่ถูกวง/ไฮไลต์: '1±350 CTNS' / '1:800 BAGS' (จุลภาคหลักพันอ่านเป็นตัวอื่น)
        t = re.sub(r'^(\d)[:;±•·.](\d{3})(?=\s+[A-Z])', r'\1\2', t)
        t = re.sub(r'^[lI|][dDoOØø](?=\s+[A-Z]{3})', '10', t)          # 'ld PALLETS' = 10 PALLETS
        if l['y'] < 130 and l['x'] > 330 and re.search(r'HASL', t.replace(' ', ''), re.I):
            t = bl
        mc = re.search(r'[A-Z]{3}[UJZ]\d{7}', t)
        if mc and mc.start() > 0 and l['x'] < 120:          # '2el SKLU0000000' / 'aeH JPHU0000000' — เศษขนาดตู้หน้าเลขตู้
            t = t[mc.start():]
        # 'G.W. KGS : 72.0000' — ปากกาไฮไลต์กลบคำ CBM → ตัวเลขหลัง KGS ในแถวเดียวกันคือ MEAS
        mk = re.match(r'^([\d,]+\.\d+\s*KGS)\s*[:;]?\s*([\d,]+\.\d+)$', t)
        if mk:
            xs = l['x'] + (l['x1'] - l['x']) * mk.start(2) / len(t)
            out.append(dict(l, x1=xs - 2, t=mk.group(1)))
            out.append(dict(l, x=xs, t=mk.group(2) + ' CBM'))
            continue
        l = dict(l, t=t)
        # แยกชิ้นที่ OCR ต่อกันข้ามคอลัมน์: 'PACKAGESi SHIPPER'S LOAD', '40RHX1 SAID TO CONTAIN'
        for key in ("SHIPPER'S LOAD", 'SAID TO CONTAIN'):
            i = l['t'].upper().find(key)
            if i > 0:
                xs = l['x'] + (l['x1'] - l['x']) * i / len(l['t'])
                left = l['t'][:i].strip()
                # ขนาด×จำนวนตู้หน้า SAID TO CONTAIN ('40 'HQXI :', '2eix1') ไม่ใช่ MARKS
                if not (key == 'SAID TO CONTAIN' and re.fullmatch(r"[\w' ]{2,9}X[\dIl1]\s*:?", left, re.I)):
                    out.append(dict(l, x1=xs - 2, t=left))
                l = dict(l, x=xs, t=l['t'][i:])
        out.append(l)
    # ขอบซ้ายคอลัมน์รายละเอียด: x ของชิ้นที่แยกได้เป็นค่าประมาณ — ใช้ค่าน้อยสุดของ SHIPPER'S LOAD / SAID TO CONTAIN
    sl = next((o for o in out if o['t'].upper().startswith("SHIPPER'S LOAD")), None)
    sd = next((o for o in out if o['t'].upper().startswith('SAID TO CONTAIN')), None)
    if sl and sd and 0 < sd['y'] - sl['y'] < 25 and sd['x'] < sl['x']:
        sl['x'] = sd['x']
    return _fix_an_labels(out)


def _fix_an_labels(lines):
    """ARRIVAL NOTICE จากสแกน: ป้ายช่องที่อ่านเพี้ยน ('Cons i onee', 'Corel!re', 'btlfy Party') → ป้ายจริง
       ให้ parse_blform แบ่งช่อง Consignee / Notify ได้; ทิ้งบรรทัดที่ไม่มีตัวอักษรละติน/ตัวเลข (เศษสแกน)
       และเศษท้ายบรรทัด ('ABC TRADING CO., LTD -')"""
    out = []
    for l in lines:
        t = l['t']
        if not re.search(r'[A-Za-z0-9]', t):
            continue
        t = re.sub(r"[\s\-‐_′″`´]+$", '', t)
        t = re.sub(r'(LTD\.?|LIMITED)\s+[7/vV✓]$', r'\1', t)    # เครื่องหมาย ✓ ที่ลูกค้าติ๊กท้ายชื่อ (OCR อ่านเป็น 7 หรือ /)
        if l['x'] < 90 and 100 < l['y'] < 140 and len(nospace(t)) <= 12 and re.match(r'^(Co|I\s*gn)', t, re.I):
            t = 'Consignee'                                       # 'Cons i onee' / 'Corel!re' / 'I gn•e' (ไม่ใช่ 'TO THE ORDER')
        elif l['x'] < 90 and re.search(r'party$', t, re.I) and len(t) < 16:
            t = 'Notify Party'
        elif l['x'] < 90 and re.match(r'^pr[eo]\s*carr', t, re.I):
            t = 'Pre Carriage by'
        out.append(dict(l, t=t))
    return out


def parse_ocr_page(lines, annots, page):
    """อ่านค่าจากบรรทัด OCR ของ ARRIVAL NOTICE / ATTACHED RIDER (ไม่ผูกกับตำแหน่งคอลัมน์ — ภาพเอียงได้)"""
    rec = blank('', page)
    norm = [dict(l, t=_ocr_cont_fix(_ocr_num(l['t']))) for l in sorted(lines, key=lambda l: (round(l['y'] / 4), l['x']))]
    txt = ' '.join(l['t'] for l in norm)
    rec['cont'] = ';'.join(conts_in(txt, *[a['t'] for a in annots]))
    m = _OCR_PKG_RE.search(txt)
    if m:
        rec['pkgs'], rec['pkgtype'] = int(m.group(1).replace(',', '')), m.group(2)
    gw = [l for l in norm if re.search(r'[\d,]+\.\d+\s*KGS?\b', l['t'], re.I) and 'NET' not in l['t'].upper()]
    if gw:
        rec['gw'] = be.num(re.search(r'([\d,]+\.\d+)\s*KGS?', gw[0]['t'], re.I).group(1))
    mm = re.search(r'([\d,]+\.\d+)\s*(?:CBM|M3)\b', txt, re.I)
    if mm:
        rec['meas'] = be.num(mm.group(1))
    st = next((l['t'].strip().upper() for l in norm
               if re.fullmatch(r'(CY/CY|CFS/CFS|CY/CFS|CFS/CY|CY|CFS|LCL|FCL)', l['t'].strip(), re.I)), '')
    st = st or clean_status([a['t'] for a in annots if is_status_line(a['t'])])
    rec['status_raw'] = st
    lab = next((l for l in norm if re.match(r'CONS\s*I\s*GNEE$', l['t'].strip(), re.I)), None)
    if lab:
        nxt = next((l for l in norm if 0 < l['y'] - lab['y'] < 16 and abs(l['x'] - lab['x']) < 15), None)
        rec['cons'] = nxt['t'] if nxt else ''
    di = next((l for l in norm if re.search(r'Description\s+Information', l['t'], re.I)), None)
    if di:                                    # ATTACHED RIDER: รายละเอียดสินค้าใต้หัว Description Information
        desc = []
        for l in norm:
            if l['y'] <= di['y'] + 3:
                continue
            if re.match(r'(NET\s*WEIGHT|TEMP|\*|CONTACT|TEL|FAX)', l['t'], re.I) or CONT2_RE.search(l['t']):
                break
            desc.append(l['t'])
        rec['desc'] = ' '.join(desc)
    mi = next((l for l in norm if re.search(r'Mark\s+Information', l['t'], re.I)), None)
    if mi:                                    # ATTACHED RIDER: เครื่องหมายใต้หัว Mark Information (แทน 'AS PER ATTACHED RIDER')
        rec['rider_marks'] = ' '.join(l['t'] for l in norm if l['y'] > mi['y'] + 3 and l['x'] < 250)
    # 'TEMP : -25'C' / OCR: ': -25 •c' (บรรทัด TEMP กับค่าอาจสลับลำดับ) — ไม่ใช้ be.find_temp (คืนทั้งก้อนข้อความ)
    mt = re.search(r"(?:TEMP\s*:?\s*)?(-\s*\d{1,2}(?:\.\d)?)\s*['’•°º]?\s*[cC]\b", txt + ' ' + ' '.join(a['t'] for a in annots))
    rec['reefer'] = f"{mt.group(1).replace(' ', '')}°C" if mt and re.search(r'TEMP', txt + ' '.join(a['t'] for a in annots), re.I) else ''
    return rec


def _fix_dsv(recs, doc):
    """DSV หน้ากว้าง (737pt) มีหลายตู้แบบ 'HLHU0000000 (40HC)' — layout G อ่านได้ตู้เดียว/ปนใน DESC"""
    for rec in recs.values():
        txt = ' '.join(doc[p - 1].get_text() for p in rec['pages'])
        cs = conts_in(txt)
        if cs:
            rec['cont'] = ';'.join(cs)
        rec['desc'] = ' '.join(w for w in re.split(r'(?<=\))\s+|\s{2,}', rec['desc'] or '')
                               if not re.fullmatch(r"[A-Z]{3}[UJZ]\d{7}\s*\(?[^)]*\)?", w.strip()))
        rec['desc'] = re.sub(r"\s*[A-Z]{3}[UJZ]\d{7}\s*\(\d{2}[A-Z]{0,2}\)", '', rec['desc']).strip()
        # เลขอ้างอิงภายในของ DSV ท้ายช่องสินค้า ไม่ใช่รายละเอียดสินค้า (SAWASDEE VEGA 2614S)
        rec['desc'] = re.sub(r"\s*\b(?:SHIPMENT\s*NO|COST\s*CENTER)\s*:\s*\S+", '', rec['desc'], flags=re.I).strip()


# ═══════════════ กฎของ ENTER ชุดรวมก่อนเทียบ ═══════════════
_ASPER_RE = re.compile(r'\*?\s*' + r'\s*'.join('ASPERATTACHED') + r'\s*(?:' + r'\s*'.join('SHEET') + '|'
                       + r'\s*'.join('RIDER') + r')?\s*\*?', re.I)   # 'AS PER ATTACHE D SHEET' (ช่องว่างกลางคำ)


def _subseq_extra(m_txt, e_txt):
    """MANIFEST เป็นส่วนหนึ่งของข้อความ B/L ใน ENTER (ตัวอักษร/ตัวเลขต่อกันตามลำดับ ไม่นับช่องว่าง)?
       คืนข้อความส่วนเกินของ ENTER หรือ None ถ้าไม่ใช่"""
    m_core = re.sub(r'[^A-Z0-9]', '', _ASPER_RE.sub(' ', str(m_txt or '').upper()))
    e_words = str(e_txt or '').split()
    if not m_core or not e_words:
        return None
    e_core = re.sub(r'[^A-Z0-9]', '', str(e_txt).upper())
    if m_core not in e_core:
        # ข้อความ MANIFEST อาจข้ามบรรทัดของ B/L (เช่นตัด 'HS CODE' ตรงกลาง) — ยอมให้เป็นลำดับย่อยทีละคำ
        it = iter(re.findall(r'[A-Z0-9]+', str(e_txt).upper()))
        if not all(any(w == x for x in it) for w in re.findall(r'[A-Z0-9]+', _ASPER_RE.sub(' ', str(m_txt).upper()))):
            return None
    extra = []
    rest = m_core
    for w in e_words:
        c = re.sub(r'[^A-Z0-9]', '', w.upper())
        if c and rest.startswith(c):
            rest = rest[len(c):]
        elif c:
            extra.append(w)
    return ' '.join(extra)


def apply_bundle_rules(MAN, ENT):
    for bl, m in MAN.items():
        # MANIFEST 'TEMP : - 18'C' (มีช่องว่างหลังเครื่องหมายลบ) — find_temp ของ skill อ่านไม่ออก
        if m.get('reefer') and '°C' not in m['reefer'] or (not m.get('reefer') and REEFER_RE.search(m.get('desc') or '')):
            m['reefer'] = find_reefer(m.get('reefer') or '', m.get('desc') or '') or m.get('reefer') or ''
        # กฎผู้ใช้ (2026-11-25): MF พิมพ์ STATUS ไว้หน้ารายละเอียด 'LCL AGENT'S LABOUR …' = STATUS LCL
        # (ช่อง STATUS ของ MF ว่าง) — ย้ายเป็น STATUS แล้วตัดออกจาก DESC
        ml = re.match(r"^\s*(LCL|CFS|CY)\s+AGENT['’]?S\s+LABOU?R\s*", m.get('desc') or '', re.I)
        if ml:
            if not (m.get('status') or '').strip():
                m['status'] = ml.group(1).upper()
            m['desc'] = m['desc'][ml.end():]
        # MF บางลำ (SWDN 2026-11-27) พิมพ์จำนวนตู้ 'X2' เป็นบรรทัดแรกของช่องรายละเอียด (ใต้ STATUS) —
        # ไม่ใช่รายละเอียดสินค้า ตัดออกเมื่อเลขเท่ากับจำนวนตู้ของ B/L นั้น
        mx = re.match(r'^\s*X\s?(\d{1,3})\s+', m.get('desc') or '')
        if mx and int(mx.group(1)) == max(1, len(m.get('cont') or [])):
            m['desc'] = m['desc'][mx.end():]
        # รหัสหีบห่อของ MF กำกวม: 'PL' ปกติ = PALLET แต่ MF เขียนชื่อเต็มกำกับ '240 PL (PAIL(s))' → ใช้ชื่อเต็ม
        mp = re.match(r'^\s*([\d,]+)\s+(\S+)\s+\(([A-Z ]+)\(s\)\)', m.get('pkg_hdr') or '', re.I)
        if mp and be.PKG_KIND.get(mp.group(2).upper(), mp.group(3).upper()) != mp.group(3).upper() \
                and mp.group(3).upper() != 'REEL':                   # REEL = ROLL อยู่แล้ว (กฎเดิม)
            m['pkg_hdr'] = f"{mp.group(1)} {mp.group(3).upper()} ({mp.group(3)}(s))"
    _an = lambda s: re.sub(r'[^A-Z0-9]', '', str(s or '').upper())
    for bl, e in ENT.items():
        # กฎผู้ใช้ (2026-11-27, SWDN 1002364): 'CL9' / 'CL 9' / 'CL.9' = 'CLASS 9' — แทนเฉพาะเมื่อ MF ของ B/L นั้น
        # เขียน CLASS เลขเดียวกัน (กัน 'CL1' ที่เป็นรหัสสินค้า) ; DG ของ ENTER อ่านใหม่ให้ได้ CLASS ด้วย
        mt = ' '.join(str((MAN.get(bl) or {}).get(k) or '') for k in ('desc', 'marks')).upper()   # ไม่ใช้ dg (MF CL9 → dg CLASS 9)
        for k in ('desc', 'desc_boxed', 'desc_boxed_full', 'marks', 'marks_boxed', 'marks_boxed_full'):
            if isinstance(e.get(k), str) and e[k]:
                e[k] = re.sub(r'\bCL\.?\s?(\d(?:\.\d)?)\b',
                              lambda q: f'CLASS {q.group(1)}' if re.search(rf'\bCLASS\s*{re.escape(q.group(1))}\b', mt)
                              else q.group(0), e[k], flags=re.I)
        _dt = ' '.join(str(e.get(k) or '') for k in ('desc', 'desc_boxed', 'desc_boxed_full'))
        if e.get('dg') and 'CLASS' not in e['dg'].upper() and re.search(r'\bCLASS\s*\d', _dt, re.I):
            e['dg'] = be.find_dg(_dt) or e['dg']
        if re.search(r'\bREELS?\b', e.get('pkgtype') or '', re.I):
            e['pkgtype'] = 'REEL (ROLL)'
        # 'FLEXIBAG' / 'FLEXITANK' = ถุงของเหลวในตู้ = TANK ของ MF (SWDN 2026-11-27 '1 FLEXIBAG, FLEXI')
        if re.search(r'FLEXI', e.get('pkgtype') or '', re.I):
            e['pkgtype'] = 'TANK'
        # STATUS พิมพ์สระผิด 'ลากตุ้' (ุ แทน ู) / ฟอนต์ถอดเป็นซีริลลิก — ซ่อมก่อนแปลงเป็น CY
        if e.get('status_raw'):
            e['status_raw'] = fix_thai_garble(e['status_raw'].replace('ตุ้', 'ตู้'))
        m = MAN.get(bl)
        # กฎผู้ใช้ (2026-11-03): MANIFEST '<ข้อความ> *AS PER ATTACHED SHEET*' — ถ้าทุกตัวอักษรก่อนคำนั้นตรงกับ
        # ต้นข้อความ ENTER (ไม่นับช่องว่าง/เครื่องหมาย) ถือว่าตรง ส่วนที่เหลือของ ENTER อยู่ใน ATTACHED SHEET
        for k in ('marks', 'desc') if m else ():
            ma = re.search(r'\*?\s*AS\s*PER\s*ATTA\s*CHED\s*SHEETS?\s*\*?', m.get(k) or '', re.I)
            if not ma or e.get(k + '_full'):
                continue
            # กฎผู้ใช้ (2026-11-25): ตรวจแค่ 30 ตัวอักษรแรกก่อน AS PER ATTACHED SHEET — ตรงกันถือว่าถูก ไม่ต้องเตือน
            pre = _an(m[k][:ma.start()])[:30]
            if pre and _an(e.get(k)).startswith(pre) and _an(e.get(k)) != _an(m[k]):
                e[k + '_full'] = e[k]
                e[k + '_attached'] = True
                e[k] = m[k]
        m = MAN.get(bl)
        if not m or e.get('src') != 'BLFORM':
            continue
        # กรอบสีแดงของลูกค้า: ตรวจเฉพาะข้อความในกรอบ — ถ้า MF มีข้อความนั้นครบถือว่าตรง
        for k in ('marks', 'desc'):
            bx = e.pop(k + '_boxed', None)
            full = e.pop(k + '_boxed_full', None)
            if not bx:
                continue
            e[k + '_full'] = full
            e[k + '_inbox'] = bx
            e[k] = m[k] if _an(bx) and _an(bx) in _an(m.get(k)) else bx
        # ช่องตัวเลขนอกกรอบแดง (ลูกค้าตีกรอบช่องอื่นในแถวเดียวกัน) = ไม่ได้แจ้ง → ใช้ค่า MF (ไม่ตรวจ) บอกค่าเดิมในหมายเหตุ
        _mfv = {'pkgs': lambda: be.pkgnum(m.get('pkg_hdr')), 'gw': lambda: be.num(m.get('gw')),
                'meas': lambda: be.num(m.get('meas'))}
        for k in e.pop('skip_fields', None) or []:
            v = _mfv[k]()
            if v is not None:
                e.setdefault('unboxed', {})[k] = e.get(k)
                e[k] = v
        # ARRIVAL NOTICE/สำเนา B/L ของ Heung-A เอง: MANIFEST ตัดข้อความ B/L บางส่วนออก (HS CODE, ที่อยู่ '*…',
        # TEL ฯลฯ) — ถ้า MANIFEST เป็นส่วนหนึ่งของ ENTER ตามลำดับ ถือว่าตรง แต่บอกส่วนเกินไว้ในหมายเหตุ
        for k in ('marks', 'desc'):
            if be.same_detail(e[k], m[k]) is False:
                extra = _subseq_extra(m[k], e[k])
                if extra is not None:
                    e[k + '_full'] = e[k]
                    e[k + '_extra'] = extra
                    e[k] = m[k]


def _one_edit(a, b):
    """a กับ b ต่างกันแค่ 1 ตัวอักษร (เพิ่ม/ขาด/ผิด 1 ตัว)"""
    if abs(len(a) - len(b)) > 1 or a == b:
        return False
    if len(a) == len(b):
        return sum(x != y for x, y in zip(a, b)) == 1
    s, l = (a, b) if len(a) < len(b) else (b, a)
    return any(l[:i] + l[i + 1:] == s for i in range(len(l)))


def pair_bl_typos(MAN, ENT):
    """B/L ใน ENTER ที่ไม่มีใน MANIFEST แต่ต่างจาก B/L ใน MANIFEST (ที่ยังไม่มี ENTER) แค่ 1 ตัวอักษร
       = ลูกค้าพิมพ์เลข B/L ผิด (เช่น 'HASLK0000000000' ขาด 0 จาก HASLK00000000000) — ย้ายไปเทียบกับ
       B/L ที่ถูก และจำเลขที่พิมพ์ผิดไว้แจ้งในรายงาน. จับคู่เฉพาะเมื่อมีตัวเลือกเดียว ไม่เดาถ้ากำกวม"""
    pairs = []
    for bad in [b for b in ENT if b not in MAN]:
        cand = [m for m in MAN if m not in ENT and _one_edit(bad, m)]
        if len(cand) == 1:
            good = cand[0]
            rec = ENT.pop(bad)
            rec['bl'] = good
            rec['bl_typo'] = bad
            ENT[good] = rec
            pairs.append((good, bad))
    return pairs


def use_shed_by_status(ENT, shed_lcl, shed_other):
    """SHED ตาม STATUS: LCL / LCL/CFS → shed_lcl, อื่น ๆ (CY) → shed_other.
       STATUS ที่ใช้ = คำสั่งใน ENTER ถ้ามี (ลูกค้าสั่งจริง) ไม่มีจึงใช้ STATUS ของ MANIFEST"""
    orig = be.compute_extra_checks

    def wrapped(m, e):
        st = be.canon_status((e or {}).get('status_raw') or '') or be.canon_status(m.get('status') or '')
        be.EXPECTED_SHED = shed_lcl if st in ('LCL', 'LCL/CFS') else shed_other
        return orig(m, e)
    be.compute_extra_checks = wrapped


def mf_without_rule_lines(MAN):
    """MF ที่ส่งออกมาแบบไม่มีคอลัมน์ '1) SHED NO. / 2) ประเทศปลายทาง / 3) CARGO MOVEMENT' เลยสักบรรทัด (mf.xls ของ
       SAWASDEE VEGA 2614S) — ช่องว่าง ≠ ผิด: SHED/ประเทศ/MOVEMENT เป็นส้ม 'ไม่มีใน MF ฉบับนี้ — ตรวจในระบบ' ไม่ใช่แดง"""
    if not MAN or any(m.get('shed_no') or m.get('dest_code_raw') or m.get('movement_raw') for m in MAN.values()):
        return False
    for m in MAN.values():
        m['no_rule_lines'] = True
    orig = be.compute_extra_checks

    def wrapped(m, e):
        out = orig(m, e)
        sh = out['shed']
        if sh['ok'] is False and not (m.get('shed_no') or '').strip():
            exp = re.search(r'ต้องเป็น (\d{4})', sh['note'] or sh['text'] or '')
            exp = exp.group(1) if exp else ''
            out['shed'] = {'ok': 'review', 'text': f'ไม่มีใน MF ฉบับนี้  (ต้องเป็น {exp})',
                           'note': f'SHED NO. — MF ฉบับนี้ไม่มีบรรทัด 1) SHED NO. ต้องตรวจในระบบว่าเป็น {exp}'}
        for k in ('dest', 'mv'):
            for f in ('text', 'note'):
                if out[k].get(f):
                    out[k][f] = out[k][f].replace('MANIFEST PDF', 'MF ฉบับนี้')
        return out
    be.compute_extra_checks = wrapped
    return True


# ═══════════════ เติมเลขหน้า ENTER / หน้าสแกน ลงรายงาน ═══════════════
def annotate_report(outpath, MAN, ENT, info):
    wb = openpyxl.load_workbook(outpath)
    ws = wb.active
    head = [c.value for c in ws[4]]
    c_bl = head.index('B/L NO.') + 1
    c_note = len(head)
    c_full = {'marks': head.index('MARKS (ENTER)') + 1, 'desc': head.index('DESC. (ENTER)') + 1}
    c_sum = 1
    f_org = PatternFill('solid', fgColor='FCE4D6')
    scan_pages = [s['page'] for s in info['scans']]
    by_cont = {}
    for s in info['scans']:
        for c in s['conts']:
            for bl, m in MAN.items():
                if c in m['cont']:
                    by_cont.setdefault(bl, []).append(s['page'])
    for r in range(5, ws.max_row + 1):
        mbl = BL_TXT_RE.search(str(ws.cell(r, c_bl).value or ''))   # แถวเตือนขึ้นต้นด้วย '⚠ '
        if not mbl:
            continue
        bl = mbl.group(0)
        note = ws.cell(r, c_note)
        prefix = ''
        if bl in ENT and ENT[bl].get('pages'):
            pg = ENT[bl]['pages']
            prefix = f"[ENTER หน้า {', '.join(map(str, pg))}] "
            if ENT[bl].get('ocr'):
                # ค่าฝั่ง ENTER อ่านจากภาพด้วย OCR — ห้ามขึ้นเขียว/แดงแบบมั่นใจ: ทั้งแถวเป็นส้ม ตรวจด้วยตา
                prefix += "[ENTER อ่านด้วย OCR จากภาพ — ตัวเลขอาจผิด ตรวจกับหน้า PDF ด้วยตา] "
                for c in range(1, c_note + 1):
                    ws.cell(r, c).fill = f_org
                s = ws.cell(r, c_sum)
                s.value = "⚠ อ่านด้วย OCR — ตรวจด้วยตา" + (f" ; {s.value}" if s.value and 'ผ่าน' not in str(s.value) else '')
            if ENT[bl].get('bl_typo'):
                prefix += f"⚠ เลข B/L ใน ENTER พิมพ์เป็น {ENT[bl]['bl_typo']} (ผิด 1 ตัว) — เทียบกับ {bl} แทน ; "
                s = ws.cell(r, c_sum)
                s.value = f"⚠ เลข B/L ใน ENTER ผิด ({ENT[bl]['bl_typo']})" + (f" ; {s.value}" if s.value and 'ผ่าน' not in str(s.value) else '')
                s.fill = PatternFill('solid', fgColor='FDD7D7')
            if ENT[bl].get('cons_note'):
                prefix += ENT[bl]['cons_note'] + ' ; '
            for k, lbl in (('marks', 'MARKS'), ('desc', 'DESC.')):
                if ENT[bl].get(k + '_full'):
                    ws.cell(r, c_full[k]).value = ENT[bl][k + '_full']
                    if ENT[bl].get(k + '_attached'):
                        prefix += f"{lbl} ตรง (30 ตัวอักษรแรกก่อน *AS PER ATTACHED SHEET* ใน MANIFEST ตรงกับต้นข้อความ ENTER) ; "
                    if ENT[bl].get(k + '_extra'):
                        prefix += f"{lbl} ตรง (MANIFEST เป็นส่วนหนึ่งของ B/L) — ENTER มีเพิ่ม: {ENT[bl][k + '_extra']} ; "
                    if ENT[bl].get(k + '_inbox'):
                        prefix += f"{lbl} ตรวจเฉพาะในกรอบแดงของลูกค้า: {ENT[bl][k + '_inbox']} ; "
            _lb = {'pkgs': 'TOTAL PKG', 'gw': 'G.W.', 'meas': 'MEAS.'}
            for k, v in (ENT[bl].get('unboxed') or {}).items():
                prefix += f"{_lb[k]} ไม่อยู่ในกรอบแดงของลูกค้า = ไม่ได้แจ้ง ไม่ตรวจ (ENTER พิมพ์ {v:,.3f}) ; "
        elif bl in MAN and bl not in ENT:
            if bl in by_cont:
                pg = by_cont[bl]
                msg = (f"ENTER น่าจะเป็นหน้า {', '.join(map(str, pg))} (ภาพสแกน อ่านอัตโนมัติไม่ได้ — เลขตู้ที่พิมพ์ทับตรงกับ "
                       f"MANIFEST) — ตรวจด้วยตา")
            elif scan_pages:
                msg = (f"ไม่พบในข้อความ ENTER — มีหน้าภาพสแกนที่อ่านอัตโนมัติไม่ได้: หน้า {', '.join(map(str, scan_pages))} "
                       f"— ตรวจด้วยตา")
            else:
                msg = ''
            if msg:
                prefix = f'[{msg}] '
                # ผู้ใช้ 2026-11-27: B/L ที่ไม่มี ENTER = เหลืองเฉพาะช่อง B/L NO. ; ส้มทั้งแถวเฉพาะเมื่อเลขตู้ชี้ไปที่หน้าสแกน
                if bl in by_cont:
                    for c in range(1, c_note + 1):
                        if c != 2:
                            ws.cell(r, c).fill = f_org
                s = ws.cell(r, c_sum)
                s.value = f"⚠ ENTER เป็นภาพสแกน — ตรวจด้วยตา" if bl in by_cont else s.value
        if prefix:
            note.value = prefix + (note.value or '')
    # ข้อความที่ขึ้นต้นด้วย '=' (ลูกค้าพิมพ์ 'STATUS ===> ลากตู้') openpyxl เขียนเป็นสูตร → Excel เปิดแล้วต้อง Repair
    # และลบค่าทิ้ง — บังคับเป็นข้อความ (รายงานนี้ไม่มีสูตรจริง)
    for wsx in wb.worksheets:
        for row in wsx.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    c.data_type = 's'
    wb.save(outpath)


# ═══════════════ main ═══════════════
def _find_one(patterns, folder):
    for pat in patterns:
        hits = glob.glob(os.path.join(folder, pat))
        if hits:
            return hits[0]
    return None


def main(argv=None):
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--manifest', default=None)
    ap.add_argument('--enter', default=None)
    ap.add_argument('--indir', default='.', help='โฟลเดอร์งาน (มี input/) — ค่าเริ่มต้น: โฟลเดอร์ปัจจุบัน')
    ap.add_argument('--outdir', default=None)
    ap.add_argument('--shed', default='', help='SHED NO. เดียวทุก B/L เช่น 0141')
    ap.add_argument('--shed-lcl', default='', help='SHED ของ B/L ที่ STATUS เป็น CFS/LCL เช่น 0141')
    ap.add_argument('--shed-other', default='', help='SHED ของ B/L ที่เหลือ (CY) เช่น 0122')
    ap.add_argument('--name', default='EDI.xlsx', help='ชื่อไฟล์รายงาน (ค่าเริ่มต้น EDI.xlsx)')
    ap.add_argument('--all-mf', action='store_true', help='(ค่าเริ่มต้นอยู่แล้ว) แสดงทุก B/L ของ MF ตามลำดับใน MF')
    ap.add_argument('--enter-only', action='store_true', help='แบบเดิม: ลงผลเฉพาะ B/L ที่มีใน ENTER')
    ap.add_argument('--check-shed', action='store_true', help='แบบเดิม: ตรวจ SHED กับ --shed / กฎ SHED ของ MF')
    a = ap.parse_args(argv)
    # ผู้ใช้ 2026-11-27: "การตรวจ m/f แยกละฟอร์ม เราจะไม่ตรวจจับ SHED และหาก BL ไหนไม่มี ENTER ให้ใส่สีเหลืองช่อง
    # B/L NO. เรียงลำดับตาม m/f" → ค่าเริ่มต้น: ไม่ตรวจ SHED, แสดงทุก B/L ของ MF ตามลำดับ MF, ไม่มี ENTER = B/L เหลือง
    be.CHECK_SHED = a.check_shed
    be.MF_ORDER = True
    a.all_mf = not a.enter_only
    _s = lambda v: (lambda m: m.group(0).zfill(4) if m else '')(re.search(r'\d{3,4}', v or ''))
    if a.check_shed:
        be.EXPECTED_SHED = _s(a.shed)
        if _s(a.shed_lcl) or _s(a.shed_other):
            use_shed_by_status(None, _s(a.shed_lcl), _s(a.shed_other))
    elif a.shed or a.shed_lcl or a.shed_other:
        print('(ไม่ตรวจ SHED แล้ว — ช่อง SHED แสดงค่าใน MF เฉย ๆ ; ใช้ --check-shed ถ้าต้องการตรวจแบบเดิม)')

    dirs = [os.path.join(a.indir, 'input'), a.indir]
    manifest = a.manifest or next((_find_one(['MANIFEST*.xls', 'MANIFEST*.xlsx', '*.xls', '*.xlsx'], d)
                                   for d in dirs if _find_one(['MANIFEST*.xls', 'MANIFEST*.xlsx', '*.xls', '*.xlsx'], d)), None)
    enter = a.enter or next((_find_one(['ENTER*.pdf', '*.pdf'], d) for d in dirs if _find_one(['ENTER*.pdf', '*.pdf'], d)), None)
    if not manifest or not enter:
        sys.exit('!! ไม่พบไฟล์ MANIFEST (.xls) หรือ ENTER (.pdf) — วางไว้ใน input/ หรือระบุ --manifest/--enter')
    outdir = a.outdir or os.path.join(a.indir, 'output')
    os.makedirs(outdir, exist_ok=True)
    outpath = os.path.join(outdir, a.name)
    if os.path.exists(outpath):
        sys.exit(f'!! มีไฟล์ {outpath} อยู่แล้ว — ไม่เขียนทับ ใช้ --name ชื่อใหม่')
    print(f'MANIFEST : {manifest}')
    print(f'ENTER    : {enter}')

    MAN, vessel, man_declared, bl_re = be.parse_manifest(manifest)
    if mf_without_rule_lines(MAN):
        print('⚠️  MF ฉบับนี้ไม่มีบรรทัด 1) SHED NO. / 2) ประเทศปลายทาง / 3) CARGO MOVEMENT — 3 ช่องนี้เป็นสีส้ม '
              '"ไม่มีใน MF ฉบับนี้" ต้องตรวจในระบบ (ใช้ MF ฉบับที่มีคอลัมน์นี้จะตรวจได้ครบ)')
    ENT, info = parse_enter_bundle(enter, bl_re, MAN)       # dict: ใช้เป็นรายการ B/L และใช้จับคู่จดหมายที่ไม่มีเลข B/L
    typos = pair_bl_typos(MAN, ENT)

    # MANIFEST พิมพ์จำนวน/ชนิดหีบห่อซ้ำในวงเล็บต้น DESCRIPTION ('(5 WOODEN CASE) ...') — เหมือน build_edi.main()
    _n = lambda s: re.sub(r'[^A-Z0-9]', '', str(s).upper())
    for bl, e in ENT.items():
        mm = re.match(r'^(\([^()]*\))\s*', (MAN.get(bl) or {}).get('desc') or '')
        if mm and e.get('pkgs') and _n(mm.group(1)) == _n(f"{e['pkgs']}{e.get('pkgtype', '')}") \
                and not (e.get('desc') or '').startswith('('):
            e['desc'] = f"{mm.group(1)} {e.get('desc') or ''}".strip()

    apply_bundle_rules(MAN, ENT)
    # กฎผู้ใช้: ลงผลเฉพาะ B/L ที่มีใน ENTER (MF ทั้งลำมี B/L อื่นไม่ต้องลงผล) — B/L ที่มีใน ENTER แต่ไม่มีใน MF
    # ยังขึ้นแถววิกฤตเหมือนเดิม ; GRAND TOTAL ของ MF ทั้งลำจึงไม่ใช้เทียบ ; --all-mf = แสดงทุก B/L ของ MF แบบเดิม
    mf_skipped = []
    if not a.all_mf:
        mf_skipped = [b for b in MAN if b not in ENT]
        MAN = {b: m for b, m in MAN.items() if b in ENT}
        man_declared = None
    stats = be.build_report(MAN, ENT, vessel, man_declared, {}, outpath)
    annotate_report(outpath, MAN, ENT, info)

    kinds = info['kinds']
    print('=' * 70)
    print('ชนิดหน้า ENTER:', ', '.join(f'{k}×{kinds.count(k)}' for k in dict.fromkeys(kinds)))
    if info.get('ocr'):
        print('หน้าภาพสแกนที่อ่านด้วย OCR (สีส้ม ตรวจด้วยตา):', ', '.join(f'{p} = {b}' for p, b in info['ocr']))
    if info.get('ocr_copy'):
        print('หน้าสแกนที่เป็นสำเนาของ B/L ที่อ่านจากข้อความแล้ว:', ', '.join(f'{p} = {b}' for p, b in info['ocr_copy']))
    if 'BLANK' in kinds:
        print('หน้าว่าง (ข้าม):', ', '.join(str(i + 1) for i, k in enumerate(kinds) if k == 'BLANK'))
    if info['scans']:
        print('หน้าภาพสแกน (อ่านอัตโนมัติไม่ได้):', ', '.join(str(s['page']) for s in info['scans']))
    print(f'บันทึกแล้ว: {outpath}')
    print(f'B/L ทั้งหมด {stats["bl"]} | จุดต้องตรวจ {stats["issue_rows"]} แถว | STATUS ไม่ตรง {stats["st_bad"]} | '
          f'DG {stats["dg"]} | REEFER {stats["reefer"]} | TRANSIT {stats["transit"]}')
    print((f'SHED NO. ไม่ตรง {stats["shed_bad"]} | ' if be.CHECK_SHED else 'SHED NO. ไม่ตรวจ | ') + f'ประเทศปลายทาง ไม่ตรง {stats["dest_bad"]} | '
          f'CARGO MOVEMENT ไม่ตรง {stats["mv_bad"]} | TAX ID ไม่ตรง NOTIFY {stats["tax_bad"]}')
    for good, bad in typos:
        print(f'⚠️  เลข B/L ใน ENTER พิมพ์ผิด: {bad} → น่าจะเป็น {good} (เทียบค่ากับ {good} แล้ว)')
    missing = [b for b in MAN if b not in ENT]
    if missing:
        print(f'⚠️  B/L ของ MF ที่ไม่มี ENTER ({len(missing)}, ช่อง B/L สีเหลือง): {", ".join(missing)}')
    if mf_skipped:
        print(f'(ไม่ลงผล {len(mf_skipped)} B/L ของ MF ที่ไม่มีใน ENTER — ใช้ --all-mf ถ้าต้องการเห็นทั้งหมด)')
        if info['scans']:
            print(f'⚠️  มีหน้า ENTER ที่อ่านไม่ได้ (หน้า {", ".join(str(s["page"]) for s in info["scans"])}) — '
                  f'B/L ในหน้านั้นอาจอยู่ในกลุ่มที่ไม่ลงผล ตรวจด้วยตา')
    extra =[b for b in ENT if b not in MAN]
    if extra:
        print(f'⚠️  B/L ที่มีใน ENTER แต่ไม่มีใน MANIFEST ({len(extra)}): {", ".join(extra)}')
    if stats['crit']:
        print('\nจุดที่ต้องตรวจสอบ:')
        for x in stats['crit']:
            print('  -', x)
    print('=' * 70)


if __name__ == '__main__':
    main()
