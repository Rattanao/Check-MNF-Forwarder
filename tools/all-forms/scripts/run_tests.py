"""ทดสอบทุกแบบฟอร์มในที่เดียว: รัน bundle_edi.py กับทุกงานใน <tests>/*/input แล้ว (ถ้าให้ --base) เทียบทุกช่อง
   กับสคริปต์รุ่นก่อน — ทุกช่องที่ต่างต้องอธิบายได้ด้วยสิ่งที่เพิ่งแก้

   python run_tests.py                                   (tests = Downloads/8-Various forms/tests)
   python run_tests.py --base /path/old/bundle_edi.py    (คัดลอกสคริปต์เดิมไว้ก่อนแก้ แล้วเทียบ)
"""
import argparse
import os
import subprocess
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TESTS = os.path.join(os.path.expanduser('~'), 'Downloads', '8-Various forms', 'tests')


def run(script, job, outdir, name):
    cmd = [sys.executable, '-I', script, '--indir', job, '--outdir', outdir, '--name', name, '--shed', '0141', '--all-mf']
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=900)
    summary = [l for l in r.stdout.splitlines() if l.startswith(('ชนิดหน้า', 'B/L ทั้งหมด'))]
    return os.path.join(outdir, name), summary, (r.stderr.strip().splitlines()[-1] if r.returncode else '')


def cells(path):
    ws = openpyxl.load_workbook(path).active
    return {(c.row, c.column): (c.value, c.fill.fgColor.rgb) for row in ws.iter_rows() for c in row}


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser()
    ap.add_argument('--tests', default=DEFAULT_TESTS)
    ap.add_argument('--base', default='', help='สคริปต์รุ่นก่อน (bundle_edi.py) ที่จะเทียบ')
    a = ap.parse_args()
    new = os.path.join(HERE, 'bundle_edi.py')
    jobs = sorted(d for d in os.listdir(a.tests) if os.path.isdir(os.path.join(a.tests, d, 'input')))
    if not jobs:
        sys.exit(f'!! ไม่พบงานทดสอบใน {a.tests} (ต้องเป็น <งาน>/input/)')
    out = tempfile.mkdtemp(prefix='bundle-tests-')
    total_diff = 0
    for j in jobs:
        job = os.path.join(a.tests, j)
        p_new, s_new, err = run(new, job, out, f'{j}-new.xlsx')
        print(f'== {j}')
        for l in s_new:
            print('  ', l)
        if err:
            print('   !! ERROR:', err)
            total_diff += 1
            continue
        if a.base:
            p_old, _, err_old = run(a.base, job, out, f'{j}-old.xlsx')
            if err_old:
                print('   !! base ERROR:', err_old)
                continue
            x, y = cells(p_old), cells(p_new)
            diff = [k for k in sorted(set(x) | set(y)) if x.get(k) != y.get(k)]
            total_diff += len(diff)
            print(f'   ช่องที่ต่างจากรุ่นก่อน: {len(diff)}')
            for k in diff[:40]:
                print(f'     r{k[0]} c{k[1]}: {str(x.get(k, ("", ""))[0])[:70]!r} -> {str(y.get(k, ("", ""))[0])[:70]!r}')
    print(f'\nผลลัพธ์อยู่ที่ {out}')
    print('รวมช่องที่ต่าง/ข้อผิดพลาด:', total_diff)


if __name__ == '__main__':
    main()
