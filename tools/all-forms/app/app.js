// Check EDI Various Forms — รันโปรแกรม Python ตัวเดียวกับบนเครื่อง (scripts/bundle_edi.py + build_edi.py)
// ในเบราว์เซอร์ด้วย Pyodide; หน้าสแกน OCR ด้วย Tesseract.js (แทน OCR ของ Windows)
// หน้าตาเหมือนเว็บตรวจ EDI ของ forwarder เจ้าอื่น (M INTER ฯลฯ)
'use strict';
(() => {
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  let files = [];         // {name, size, kind: 'MANIFEST'|'ENTER'|null, file}
  let py = null;          // Pyodide
  let reportBytes = null;
  let table = null;       // {head, rows:[[{v,s,bad}]], vessel}

  // ---------- ไฟล์ ----------
  const kindOf = (f) => {
    const n = f.name.toLowerCase();
    if (/\.xlsx?$/.test(n)) return 'MANIFEST';
    if (/\.pdf$/.test(n)) return /^(mf|manifest)/.test(n) ? 'MANIFEST' : 'ENTER';
    return null;
  };
  function addFiles(list) {
    $('msg').textContent = '';
    for (const f of list) {
      if (files.some((x) => x.name === f.name && x.size === f.size)) continue;
      const kind = kindOf(f);
      if (kind) files = files.filter((x) => x.kind !== kind);   // MANIFEST 1 ไฟล์ + ENTER 1 ไฟล์
      files.push({ name: f.name, size: f.size, kind, file: f });
    }
    renderFiles();
  }
  function renderFiles() {
    $('files').innerHTML = files.map((f, i) =>
      `<span class="file">${esc(f.name)} <span class="k ${f.kind || 'NONE'}">${f.kind || 'ไม่รู้จัก'}</span>` +
      `<button class="x" data-i="${i}" title="เอาออก">×</button></span>`).join('');
    const has = (k) => files.some((f) => f.kind === k);
    $('run').disabled = !(has('MANIFEST') && has('ENTER'));
  }

  const setStatus = (t, pct) => {
    $('prog').classList.remove('hidden');
    $('status').textContent = t;
    if (pct !== undefined) $('bar').style.width = Math.max(0, Math.min(100, pct)) + '%';
  };

  // ไม่ตรวจ SHED แล้ว (ผู้ใช้ 2026-11-27) — ไม่มีช่องใส่ SHED

  // ---------- ตัวประมวลผล ----------
  async function boot() {
    if (py) return py;
    setStatus('กำลังโหลดตัวประมวลผล Python (ครั้งแรกใช้เวลาสักครู่)…', 5);
    py = await loadPyodide();
    setStatus('กำลังโหลดไลบรารีอ่าน PDF / Excel…', 15);
    await py.loadPackage(['pymupdf', 'xlrd', 'micropip']);
    await py.runPythonAsync("import micropip\nawait micropip.install(['openpyxl'])");
    setStatus('กำลังโหลดโปรแกรมตรวจ…', 25);
    py.FS.mkdirTree('/work/scripts');
    for (const f of ['bundle_edi.py', 'build_edi.py']) {
      const r = await fetch('scripts/' + f + '?v=' + Date.now(), { cache: 'no-cache' });
      if (!r.ok) throw new Error('โหลด ' + f + ' ไม่ได้');
      py.FS.writeFile('/work/scripts/' + f, new Uint8Array(await r.arrayBuffer()));
    }
    await py.runPythonAsync(`
import sys, io, os, json
sys.path.insert(0, '/work/scripts')
import bundle_edi as B
def _run(argv):
    buf = io.StringIO(); old = sys.stdout; sys.stdout = buf
    try:
        try:
            B.main(list(argv))
        except SystemExit as e:
            if e.code not in (None, 0): print(e)
    finally:
        sys.stdout = old
    return buf.getvalue()
def _png(path, i):
    import pymupdf
    d = pymupdf.open(path)
    return d[i].get_pixmap(dpi=300, colorspace=pymupdf.csGRAY).tobytes('png')
_FILL = {'FDD7D7': 'diff', 'FCE4D6': 'warn', 'FFF2CC': 'yel', 'C6EFCE': 'grn', '1F4E78': 'navy', '1F3864': 'navy'}
def _table(path):
    import openpyxl
    ws = openpyxl.load_workbook(path).active
    head = [ws.cell(4, c).value for c in range(1, ws.max_column + 1)]
    n = max(i for i, h in enumerate(head, 1) if h)
    rows = []
    for r in range(5, ws.max_row + 1):
        if not any(ws.cell(r, c).value not in (None, '') for c in range(1, n + 1)):
            continue
        cells = []
        for c in range(1, n + 1):
            x = ws.cell(r, c)
            rgb = (x.fill.fgColor.rgb or '')[-6:] if x.fill is not None and x.fill.fill_type == 'solid' and isinstance(x.fill.fgColor.rgb, str) else ''
            col = (x.font.color.rgb or '')[-6:] if x.font is not None and x.font.color is not None and isinstance(x.font.color.rgb, str) else ''
            cells.append({'v': '' if x.value is None else str(x.value), 's': _FILL.get(rgb.upper(), ''), 'bad': col.upper() == 'B00000'})
        rows.append(cells)
    return json.dumps({'head': head[:n], 'rows': rows, 'vessel': ws.cell(2, 1).value or ''}, ensure_ascii=False)
`);
    return py;
  }

  // Tesseract รวมข้อความข้ามคอลัมน์เป็นบรรทัดเดียว ('MSC CO., LTD. : HASLK…') — OCR ของ Windows แยกเป็นชิ้นตามช่องว่าง
  // จึงตัดบรรทัดเมื่อช่องว่างระหว่างคำกว้างกว่าความสูงตัวอักษร และตัด ':' / '|' เดี่ยว (เส้นประแบ่งช่องของฟอร์ม)
  function tessLines(data) {
    const out = [];
    const clean = (t) => t.replace(/[©®]/g, '0').replace(/\s+/g, ' ').trim();
    const emit = (ws) => {
      const t = clean(ws.map((w) => w.text).join(' '));
      if (!t) return;
      out.push({ x: Math.min(...ws.map((w) => w.bbox.x0)), y: Math.min(...ws.map((w) => w.bbox.y0)),
                 x1: Math.max(...ws.map((w) => w.bbox.x1)), y1: Math.max(...ws.map((w) => w.bbox.y1)), t });
    };
    const line = (l) => {
      const words = (l.words || []).filter((w) => (w.text || '').trim());
      if (!words.length) { const t = clean(l.text || ''); if (t) out.push({ x: l.bbox.x0, y: l.bbox.y0, x1: l.bbox.x1, y1: l.bbox.y1, t }); return; }
      const h = Math.max(8, (l.bbox.y1 - l.bbox.y0));
      let seg = [];
      for (const w of words) {
        const sep = /^[:|¦!]+$/.test(w.text.trim());
        const prev = seg[seg.length - 1];
        if (sep || (prev && w.bbox.x0 - prev.bbox.x1 > 1.1 * h)) { if (seg.length) emit(seg); seg = []; }
        if (!sep) seg.push(w);
      }
      if (seg.length) emit(seg);
    };
    if (data.blocks) for (const b of data.blocks) for (const p of (b.paragraphs || [])) for (const l of (p.lines || [])) line(l);
    else if (data.lines) data.lines.forEach(line);
    return out;
  }

  // ---------- ตรวจ ----------
  async function check() {
    $('run').disabled = true; $('msg').textContent = ''; $('result').classList.add('hidden');
    reportBytes = null; table = null;
    const mf = files.find((f) => f.kind === 'MANIFEST'), en = files.find((f) => f.kind === 'ENTER');
    try {
      await boot();
      setStatus('กำลังอ่านไฟล์…', 30);
      const mfName = 'MF' + mf.name.slice(mf.name.lastIndexOf('.')).toLowerCase();
      for (const d of ['/work/input', '/work/output']) py.FS.mkdirTree(d);
      const clearOut = () => { for (const f of py.FS.readdir('/work/output')) if (f !== '.' && f !== '..') py.FS.unlink('/work/output/' + f); };
      clearOut();
      py.FS.writeFile('/work/input/' + mfName, new Uint8Array(await mf.file.arrayBuffer()));
      py.FS.writeFile('/work/input/ENTER.pdf', new Uint8Array(await en.file.arrayBuffer()));
      const argv = ['--manifest', '/work/input/' + mfName, '--enter', '/work/input/ENTER.pdf',
                    '--outdir', '/work/output', '--name', 'EDI.xlsx'];
      const run = py.globals.get('_run'), png = py.globals.get('_png'), B = py.pyimport('bundle_edi');

      // รอบ 1: แยกชนิดหน้า อ่านทุกฟอร์ม และจดหน้าที่ต้อง OCR
      setStatus('กำลังแยกชนิดหน้าและอ่านทุกฟอร์ม…', 35);
      await new Promise((r) => setTimeout(r, 30));
      B.BROWSER_OCR = py.toPy(new Map()); B.OCR_WANTED.clear();
      let out = run(py.toPy(argv));
      const wanted = [...B.OCR_WANTED.toJs()].sort((a, b) => a - b);

      if (wanted.length) {
        const worker = await Tesseract.createWorker('eng');
        const ocr = new Map();
        for (let k = 0; k < wanted.length; k++) {
          const i = wanted[k];
          setStatus(`OCR หน้าภาพสแกน ${k + 1}/${wanted.length} (หน้า ${i + 1})…`, 40 + 50 * k / wanted.length);
          const bytes = png(py.toPy('/work/input/ENTER.pdf'), i).toJs();
          const { data } = await worker.recognize(new Blob([bytes], { type: 'image/png' }), {}, { blocks: true });
          ocr.set(i, tessLines(data));          // คีย์ = เลขหน้า 0-based (ตัวเลข)
        }
        await worker.terminate();
        setStatus('กำลังเทียบกับ MANIFEST (รอบสุดท้าย)…', 92);
        await new Promise((r) => setTimeout(r, 30));
        B.BROWSER_OCR = py.toPy(ocr);
        clearOut();
        out = run(py.toPy(argv));
      }
      $('log').textContent = out.trim();
      if (!py.FS.analyzePath('/work/output/EDI.xlsx').exists) throw new Error(out.trim().split('\n').pop() || 'ไม่ได้รายงาน');
      reportBytes = py.FS.readFile('/work/output/EDI.xlsx');
      table = JSON.parse(py.globals.get('_table')('/work/output/EDI.xlsx'));
      renderResult(out);
      setStatus('เสร็จแล้ว', 100);
    } catch (e) {
      console.error(e);
      setStatus('ตรวจไม่สำเร็จ', 0);
      $('msg').textContent = 'เกิดข้อผิดพลาด: ' + (e.message || e);
    } finally {
      renderFiles();
    }
  }

  // ---------- แสดงผล ----------
  const isTotal = (c) => /^TOTAL/.test(c[1].v);
  function renderResult(out) {
    const rows = table.rows.filter((c) => !isTotal(c));
    const total = table.rows.find(isTotal);
    const issues = rows.filter((c) => c[0].s !== 'grn').length;
    const m = (re) => (out.match(re) || [])[1];
    const vessel = (String(table.vessel).match(/Vessel\s+(.*?)\s*\|/) || [])[1] || '(ไม่ระบุ)';
    const skipped = m(/ไม่ลงผล (\d+) B\/L/);
    const unread = m(/หน้าภาพสแกน \(อ่านอัตโนมัติไม่ได้\): ([^\n]+)/);
    const ocr = m(/หน้าภาพสแกนที่อ่านด้วย OCR[^:]*: ([^\n]+)/);
    $('info').innerHTML =
      `<div class="k">VESSEL</div><div>${esc(vessel)}</div>` +
      `<div class="k">ENTER</div><div>${esc(files.find((f) => f.kind === 'ENTER').name)}</div>` +
      (ocr ? `<div class="k">OCR</div><div>หน้า ${esc(ocr.split(',').map((s) => s.split('=')[0].trim()).join(', '))}</div>` : '') +
      (unread ? `<div class="k">อ่านไม่ได้</div><div>หน้า ${esc(unread)}</div>` : '');
    const crit = rows.filter((c) => /ไม่มีใน MANIFEST/.test(c[0].v)).length;
    $('chips').innerHTML =
      `<span class="chip all">ทั้งหมด ${rows.length} B/L</span><span class="chip bad">⚠ ต้องตรวจ ${issues}</span>` +
      `<span class="chip ok">✔ ผ่าน ${rows.length - issues}</span>` +
      (crit ? `<span class="chip bad">⚠⚠ ไม่มีใน MANIFEST ${crit}</span>` : '') +
      (skipped ? `<span class="chip warn">B/L ใน MF ที่ไม่มีใน ENTER ${skipped} (ไม่ลงผล)</span>` : '') +
      (total && total[0].s === 'grn' ? '<span class="chip ok">ยอดรวมตรงกัน</span>' : '<span class="chip bad">⚠ ยอดรวมไม่ตรง</span>');
    const NUM = /^(TOTAL PKG|G\.W\.|MEAS\.)/, WIDE = /MARKS|DESC|หมายเหตุ|สรุปผล|CNEE|CNTRS/;
    $('thead').innerHTML = '<tr>' + table.head.map((h, i) => `<th class="${i < 2 ? 's' + i : ''}">${esc(h)}</th>`).join('') + '</tr>';
    table._cls = table.head.map((h, j) => [NUM.test(h) ? 'r' : '', WIDE.test(h) ? 'wide' : '', j < 2 ? 's' + j : ''].join(' '));
    drawRows();
    $('result').classList.remove('hidden');
    $('result').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function tr(cells, cls) {
    return `<tr class="${cls || ''}">` + cells.map((c, j) =>
      `<td class="${[c.s, c.bad ? 'bad' : '', table._cls[j]].join(' ').trim()}">${esc(c.v)}</td>`).join('') + '</tr>';
  }
  function drawRows() {
    const q = $('search').value.trim().toLowerCase(), only = $('onlyBad').checked;
    const rows = table.rows.filter((c) => !isTotal(c));
    const total = table.rows.find(isTotal);
    const shown = rows.filter((c) => (!only || c[0].s !== 'grn') && (!q || c.map((x) => x.v).join(' ').toLowerCase().includes(q)));
    $('tbody').innerHTML = shown.map((c) => tr(c)).join('') + (total ? tr(total, 'total') : '');
  }

  function download() {
    if (!reportBytes) return;
    const en = files.find((f) => f.kind === 'ENTER');
    const base = en ? en.name.replace(/\.pdf$/i, '') : 'ENTER';
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([reportBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    a.download = 'EDI ' + base.replace(/[^\w\-ก-๙ ]+/g, '_').slice(0, 40).trim() + '.xlsx';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  // ---------- events ----------
  const drop = $('drop');
  drop.onclick = () => $('file').click();
  $('file').onchange = (e) => { addFiles([...e.target.files]); e.target.value = ''; };
  drop.ondragover = (e) => { e.preventDefault(); drop.classList.add('over'); };
  drop.ondragleave = () => drop.classList.remove('over');
  drop.ondrop = (e) => { e.preventDefault(); drop.classList.remove('over'); addFiles([...e.dataTransfer.files]); };
  $('files').onclick = (e) => { const i = e.target.dataset.i; if (i !== undefined) { files.splice(+i, 1); renderFiles(); } };
  $('run').onclick = check;
  $('search').oninput = () => table && drawRows();
  $('onlyBad').onchange = () => table && drawRows();
  $('download').onclick = download;

  // ใช้ทดสอบอัตโนมัติ: window.CheckVarious.run(mfFile, enterFile)
  window.CheckVarious = {
    py: () => py,
    boot: () => boot(),
    async run(mf, en, opt = {}) {
      files = [{ name: mf.name, size: mf.size, kind: 'MANIFEST', file: mf }, { name: en.name, size: en.size, kind: 'ENTER', file: en }];
      renderFiles();
      await check(); return { ok: !!reportBytes, log: $('log').textContent, table };
    },
  };
})();
