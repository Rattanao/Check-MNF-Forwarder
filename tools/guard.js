// เปิดหน้าเครื่องมือตรง ๆ (ไม่ผ่านเว็บรวม) ตอนที่ auth.js ตั้ง lock: true และยังไม่ได้เข้าระบบ
// -> ส่งไปหน้าเข้าระบบของเว็บรวม แล้วกลับมาที่เครื่องมือนี้หลังเข้าระบบ
// ในเว็บรวม (อยู่ใน iframe) ไม่ต้องตรวจซ้ำ เพราะเว็บรวมไม่เปิดเครื่องมือจนกว่าจะเข้าระบบ
(function () {
  if (window.self !== window.top) return;
  const root = new URL("../", document.currentScript.src);
  const html = document.documentElement;
  html.style.visibility = "hidden";                     // ไม่ให้เห็นหน้าก่อนตรวจเสร็จ
  const reveal = () => { html.style.visibility = ""; };
  const s = document.createElement("script");
  s.src = new URL("auth.js", root).href;
  s.onload = () => {
    const A = window.AUTH;
    if (!A || !A.lock || A.current()) return reveal();
    const parts = location.pathname.split("/").filter(Boolean);
    let id = parts.pop() || "";
    if (id === "index.html") id = parts.pop() || "";
    location.replace(root.href + "#" + id.replace(/\.html$/, ""));
  };
  s.onerror = reveal;
  document.head.appendChild(s);
})();
