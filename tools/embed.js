// โหลดในทุกเครื่องมือ (ใส่ไว้ใน <head>) เพื่อเพิ่มของที่เว็บรวมต้องการโดยไม่แก้โค้ดต้นฉบับ:
//  1. แสดงในเว็บรวม: ซ่อนหัว/ท้ายของเครื่องมือ ใช้ธีมสว่าง
//  2. ปุ่ม "ลบไฟล์" ข้างปุ่มประมวลผล สำหรับเครื่องมือที่ยังไม่มี
//  3. ตั้งชื่อไฟล์ดาวน์โหลดตามเครื่องมือ เช่น EDI.xlsx -> "EDI Union World.xlsx"
// ดาวน์โหลด: ใช้การดาวน์โหลดปกติของเบราว์เซอร์ ไฟล์ลงโฟลเดอร์ Downloads เสมอ ไม่ถามโฟลเดอร์
(function () {
  // ชื่อที่ใช้ต่อท้ายไฟล์ ตามโฟลเดอร์ tools/<ชื่อ>/ (เพิ่มเครื่องมือใหม่: เพิ่มที่นี่ด้วย)
  const NAMES = {
    "able": "Able", "agn": "AGN", "cargoport": "Cargoport", "cms": "CMS", "dsv": "DSV", "eastern-air": "Eastern Air",
    "gateway": "Gateway", "lx-pantos": "LX Pantos", "m-corp": "M-CORP", "mac-nels": "MAC-NELS", "penanshin": "Penanshin", "planet": "Planet",
    "quality": "Quality", "shipco": "Shipco", "union-world": "Union World"
  };
  const folder = location.pathname.split("/").filter(p => p && p !== "index.html").pop() || "";
  const toolName = NAMES[folder] || document.title.replace(/^check\s+edi\s*/i, "").trim();

  // เครื่องมือทุกตัวดาวน์โหลดด้วย <a download="EDI.xlsx" href="blob:..."> แล้วเรียก a.click() จึงเปลี่ยนชื่อที่จุดนี้
  if (toolName) {
    const realClick = HTMLAnchorElement.prototype.click;
    HTMLAnchorElement.prototype.click = function () {
      const name = this.getAttribute("download");
      if (name && !name.includes(toolName)) {
        const dot = name.lastIndexOf(".");
        const base = dot > 0 ? name.slice(0, dot) : name, ext = dot > 0 ? name.slice(dot) : "";
        this.download = `${base} ${toolName}${ext}`;
      }
      return realClick.call(this);
    };
  }

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
  `;
  document.head.appendChild(style);

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
