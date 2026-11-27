"""ทดสอบฟอร์มของทุกเจ้าในชุดรวม: รวม ENTER ของหลายเจ้าเป็น PDF เดียว (ชุดรวมจำลอง) → skill ต้องแยกหน้าและอ่านทุก B/L ได้เท่ากับตัวอ่านเฉพาะของแต่ละเจ้า"""
import importlib.util, os, re, sys, tempfile, collections
sys.stdout.reconfigure(encoding='utf-8')
sys.argv = [sys.argv[0]]
K = os.path.expanduser(r'~\.claude\skills\bundle-enter-check\scripts')
s2 = importlib.util.spec_from_file_location('bundle', os.path.join(K, 'bundle_edi.py'))
B = importlib.util.module_from_spec(s2); s2.loader.exec_module(B)
be = B.be
import pymupdf, glob
D = os.path.expanduser(r'~\Downloads')
JOBS = ['Sinotrans', 'LX Pantos', 'Planet', 'union world', 'Freight links', 'Shipco', 'Able', 'Quality', 'AGN-LAO', 'Handle inter']
F = ['cons', 'cont', 'pkgs', 'pkgtype', 'gw', 'meas', 'marks', 'desc', 'status_raw', 'dg', 'reefer', 'transit_raw']
norm = lambda v: ' '.join(str(v if v is not None else '').split())
bl_re = re.compile(r'^HASL[A-Z0-9]{5,}$')

mix = pymupdf.open()
ref_all, man_all, src = {}, [], {}
for job in JOBS:
    inp = os.path.join(D, f'check edi-{job}', 'input')
    pdfs = sorted(glob.glob(os.path.join(inp, '*.pdf')), key=lambda p: (len(p), p))
    mans = [p for p in glob.glob(os.path.join(inp, '*')) if os.path.splitext(p)[1].lower() in ('.xls', '.xlsx')] \
        or [p for p in pdfs if 'MANIFEST' in os.path.basename(p).upper()]
    ents = [p for p in pdfs if 'MANIFEST' not in os.path.basename(p).upper()]
    one = pymupdf.open()
    for p in ents:
        one.insert_pdf(pymupdf.open(p), annots=True)
    fd, tmp = tempfile.mkstemp(suffix='.pdf'); os.close(fd); one.save(tmp)
    MAN, _, _, jre = be.parse_manifest(mans[0])
    ref, _ = be.parse_enter(tmp, jre)
    for bl, r in ref.items():
        ref_all[bl] = r; src[bl] = job
    man_all += list(MAN)
    mix.insert_pdf(pymupdf.open(tmp), annots=True)
fd, mixed = tempfile.mkstemp(suffix='-mixed.pdf'); os.close(fd); mix.save(mixed)
print('ชุดรวมจำลอง:', mix.page_count, 'หน้า จาก', len(JOBS), 'เจ้า')
got, info = B.parse_enter_bundle(mixed, bl_re, man_all)
print('ชนิดหน้า:', dict(collections.Counter(info['kinds'])))
per = collections.defaultdict(lambda: [0, 0, 0])
for bl, r in ref_all.items():
    j = src[bl]; per[j][0] += 1
    g = got.get(bl)
    if g is None:
        per[j][2] += 1; print('   ไม่เจอ:', j, bl); continue
    bad = [f for f in F if norm(r.get(f)) != norm(g.get(f)) and not (f == 'status_raw' and j == 'Shipco')]
    if bad:
        per[j][2] += 1; print('   ต่าง:', j, bl, bad)
    else:
        per[j][1] += 1
for j in JOBS:
    n, ok, bad = per[j]
    print(f'  {j:14} B/L {n:3}  ตรงตัวอ่านเฉพาะ {ok:3}  ต่าง/หาย {bad}')
print('B/L เกินที่ไม่ควรมี:', [b for b in got if b not in ref_all][:10])
