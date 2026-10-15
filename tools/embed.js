// โหลดในทุกเครื่องมือ (ใส่ไว้ใน <head>) เพื่อเพิ่มของที่เว็บรวมต้องการโดยไม่แก้โค้ดต้นฉบับ:
//  1. แสดงในเว็บรวม: ซ่อนหัว/ท้ายของเครื่องมือ ใช้ธีมสว่าง
//  2. ดาวน์โหลด: ให้เลือกโฟลเดอร์ที่จะบันทึก (File System Access API) ถ้าเบราว์เซอร์ไม่รองรับใช้การดาวน์โหลดปกติ
//  3. ปุ่ม "ลบไฟล์" ข้างปุ่มประมวลผล สำหรับเครื่องมือที่ยังไม่มี
(function () {
  const root = document.documentElement;
  if (window.self !== window.top) {
    root.classList.add("embedded");
    root.setAttribute("data-theme", "light");
  }

  const style = document.createElement("style");
  style.textContent = `
    html.embedded body>header, html.embedded body>footer { display: none !important; }
    html.embedded body { background: #fff !important; }
    html.embedded main { max-width: none !important; padding: 18px 20px 28px !important; }
    html.embedded .wrap { max-width: none !important; }
    .x-btns { display: flex; gap: 10px; align-items: stretch; }
    button.x-clear { font: inherit; font-weight: 600; font-size: 15px; padding: 10px 18px; border-radius: 6px; cursor: pointer;
      background: #fff !important; color: var(--navy, #1f3f94) !important; border: 1px solid var(--navy, #1f3f94) !important; white-space: nowrap; }
    button.x-clear:hover { background: #eef2fb !important; }
    .x-toast { position: fixed; right: 16px; bottom: 16px; z-index: 9999; max-width: min(440px, calc(100vw - 32px));
      background: #1f3f94; color: #fff; padding: 10px 16px; border-radius: 8px; font: 500 14px/1.4 "Sarabun", Tahoma, sans-serif;
      box-shadow: 0 4px 14px rgba(0,0,0,.2); transition: opacity .3s; }
    .x-toast.warn { background: #b4441c; }
    .x-toast-btn { font: inherit; font-weight: 600; margin-left: 6px; padding: 2px 10px; border-radius: 5px; cursor: pointer;
      background: #fff; color: #1f3f94; border: 0; }
  `;
  document.head.appendChild(style);

  let toastTimer;
  function toast(text, warn, action) {
    let t = document.querySelector(".x-toast");
    if (!t) { t = document.createElement("div"); t.className = "x-toast"; t.setAttribute("role", "status"); document.body.appendChild(t); }
    t.textContent = text;
    if (action) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "x-toast-btn";
      b.textContent = action.label;
      b.addEventListener("click", action.run);
      t.append(" ", b);
    }
    t.classList.toggle("warn", !!warn);
    t.style.opacity = "1";
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.style.opacity = "0"; }, 5000);
  }

  // ---- ดาวน์โหลดแบบเลือกโฟลเดอร์ ----
  // เครื่องมือทุกตัวดาวน์โหลดด้วย <a download href="blob:..."> แล้วเรียก a.click() จึงดักที่จุดนี้จุดเดียว
  if ("showDirectoryPicker" in window) {
    const realClick = HTMLAnchorElement.prototype.click;
    const realRevoke = URL.revokeObjectURL.bind(URL);
    // เครื่องมือ revoke blob หลังคลิก 1-5 วินาที แต่ผู้ใช้อาจใช้เวลาเลือกโฟลเดอร์นานกว่านั้น
    URL.revokeObjectURL = url => setTimeout(() => realRevoke(url), 120000);

    // id แยกตามเครื่องมือ: เบราว์เซอร์จำโฟลเดอร์ล่าสุดของแต่ละเครื่องมือไว้เปิดครั้งถัดไป
    // เครื่องมืออยู่ที่ tools/<ชื่อ>/index.html จึงใช้ชื่อโฟลเดอร์
    const parts = location.pathname.split("/").filter(p => p && p !== "index.html");
    const pickerId = ("t-" + (parts.pop() || "index").replace(/\W/g, "")).slice(0, 32);
    let batch = null;   // ไฟล์ที่ดาวน์โหลดจากการกดครั้งเดียวกัน (เช่น DO ได้ 2 ไฟล์) ใช้โฟลเดอร์เดียวกัน ถามครั้งเดียว

    // จำโฟลเดอร์ที่เลือกไว้ใน IndexedDB: ครั้งต่อไปบันทึกลงโฟลเดอร์เดิมทันทีโดยไม่ต้องเลือกใหม่
    // ถ้าเบราว์เซอร์ถามสิทธิ์อีก ให้กด "อนุญาตทุกครั้งที่เข้าชม" แล้วจะไม่ถามอีก
    function db() {
      return new Promise((ok, bad) => {
        const r = indexedDB.open("edi-folders", 1);
        r.onupgradeneeded = () => r.result.createObjectStore("dirs");
        r.onsuccess = () => ok(r.result);
        r.onerror = () => bad(r.error);
      });
    }
    async function dbDo(mode, fn) {
      const d = await db();
      return new Promise((ok, bad) => {
        const r = fn(d.transaction("dirs", mode).objectStore("dirs"));
        r.onsuccess = () => ok(r.result);
        r.onerror = () => bad(r.error);
      });
    }
    const loadDir = () => dbDo("readonly", s => s.get(pickerId)).catch(() => null);
    const saveDir = h => dbDo("readwrite", s => s.put(h, pickerId)).catch(() => {});
    const forgetDir = () => dbDo("readwrite", s => s.delete(pickerId)).catch(() => {});

    async function getDir() {
      const saved = await loadDir();
      if (saved) {
        const opts = { mode: "readwrite" };
        try {
          if (await saved.queryPermission(opts) === "granted") return saved;
          if (await saved.requestPermission(opts) === "granted") return saved;
        } catch (e) { console.warn("saved folder unavailable:", e); }
      }
      const dir = await window.showDirectoryPicker({ id: pickerId, mode: "readwrite" });
      await saveDir(dir);
      return dir;
    }

    function pickFolder() {
      if (batch && (batch.doneAt === null || Date.now() - batch.doneAt < 10000)) return batch.dir;
      const b = { doneAt: null };
      b.dir = getDir();
      b.dir.then(() => { b.doneAt = Date.now(); }, () => { b.doneAt = Date.now(); });
      batch = b;
      return b.dir;
    }

    HTMLAnchorElement.prototype.click = function () {
      const name = this.getAttribute("download");
      if (name === null || !/^(blob|data):/.test(this.href)) return realClick.call(this);
      const a = this, fileName = name || "download";
      const data = fetch(a.href).then(r => r.blob());
      pickFolder()
        .then(async dir => {
          const fh = await dir.getFileHandle(fileName, { create: true });
          const w = await fh.createWritable();
          await w.write(await data);
          await w.close();
          toast(`บันทึก ${fileName} ลงโฟลเดอร์ ${dir.name} แล้ว`, false,
            { label: "เปลี่ยนโฟลเดอร์", run: () => { forgetDir(); toast("ครั้งหน้าจะให้เลือกโฟลเดอร์ใหม่"); } });
        })
        .catch(err => {
          // ไม่ว่าจะกดยกเลิก เลือกโฟลเดอร์ที่เบราว์เซอร์บล็อก (เช่น Downloads ชั้นนอกสุด, Desktop) ไม่อนุญาต
          // หรือหมดเวลาการกดปุ่ม ก็ยังต้องได้ไฟล์: ดาวน์โหลดแบบปกติลงโฟลเดอร์ Downloads แทน
          console.warn("folder save failed, falling back to download:", err);
          realClick.call(a);
          toast(err && err.name === "AbortError"
            ? "ไม่ได้บันทึกลงโฟลเดอร์ที่เลือก จึงดาวน์โหลดลง Downloads แทน · เลือกโฟลเดอร์ Downloads หรือ Desktop ตรง ๆ ไม่ได้ ให้เลือกโฟลเดอร์ย่อยข้างใน"
            : "บันทึกลงโฟลเดอร์ไม่ได้ จึงดาวน์โหลดลง Downloads แทน", true);
        });
    };
  }

  // ---- ปุ่มลบไฟล์ ----
  document.addEventListener("DOMContentLoaded", () => {
    if (document.getElementById("clear")) return;            // เครื่องมือมีปุ่มของตัวเองแล้ว (DO)
    const run = document.querySelector("#run, #go");
    if (!run) return;
    const b = document.createElement("button");
    b.type = "button";
    b.className = "x-clear";
    b.textContent = "ลบไฟล์";
    b.title = "ล้างไฟล์และผลลัพธ์ทั้งหมด เริ่มใหม่";
    // โหลดหน้าใหม่ด้วย URL ใหม่ เพื่อไม่ให้เบราว์เซอร์คืนค่าที่กรอกไว้
    b.addEventListener("click", () => location.replace(location.pathname + "?clear=" + Date.now()));
    if (run.parentElement.classList.contains("row")) {
      run.after(b);
    } else {
      const wrap = document.createElement("div");
      wrap.className = "x-btns";
      run.replaceWith(wrap);
      wrap.append(run, b);
    }
  });
})();
