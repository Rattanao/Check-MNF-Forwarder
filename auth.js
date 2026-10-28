// ผู้ใช้ที่เข้าเว็บได้ (User ID / Password)
//
// เปิดการล็อก: เปลี่ยน lock: false เป็น lock: true แล้ว commit + push
// เพิ่มผู้ใช้: เปิด tools/add-user.html ใส่ User ID กับ Password แล้วคัดลอกบรรทัดที่ได้มาวางใน users ด้านล่าง
// ลบผู้ใช้: ลบบรรทัดของคนนั้นออก
//
// เก็บเฉพาะค่า hash ของรหัสผ่าน ไม่เก็บรหัสผ่านจริงในไฟล์นี้
// hash = SHA-256 ของ "Check-MNF-Forwarder|<USER ID ตัวใหญ่>|<password>"
window.AUTH = {
  lock: false,
  users: [
    { id: "KAN", hash: "efd97f093cc8fa9dcb039c7ac977da916fd9f8c68169e77dbff594ff1aaaf98d" },
    { id: "AOR", hash: "5d54563652a3336db999a0d114c35cf58878a1c23bb1d90c5028fd614cdfce36" },
    { id: "YAH", hash: "c9157a01c652b875d016f718053c6210a505db4dfe621ea4a62599308d953986" },
    { id: "FAH", hash: "317019279f6cb1d8885866a1ff451b878c8e1060568712fd28e089680dcdc036" },
    { id: "BEN", hash: "efb5fd8412c3d8137e4b6820feaaa8d5c19a54074e568300950a04e1c27bef3d" },
    { id: "BEE", hash: "a92136b39785321a5167d7dc0409614dd639216d02ee1c65e251b5349aa8fbb1" },
  ],
};

// ใช้ร่วมกันในหน้าเว็บรวมและหน้าเครื่องมือ
window.AUTH.hash = async (id, password) => {
  const data = new TextEncoder().encode("Check-MNF-Forwarder|" + id.trim().toUpperCase() + "|" + password);
  const buf = await crypto.subtle.digest("SHA-256", data);
  return [...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, "0")).join("");
};
window.AUTH.check = async (id, password) => {
  const h = await window.AUTH.hash(id, password);
  const u = window.AUTH.users.find(u => u.id.toUpperCase() === id.trim().toUpperCase());
  return u && u.hash === h ? u.id : null;
};
// ผู้ใช้ที่เข้าระบบอยู่ (จำไว้จนปิดแท็บ/เบราว์เซอร์)
window.AUTH.current = () => { try { return sessionStorage.getItem("hal-user"); } catch (e) { return null; } };
