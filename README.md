# Check MNF Forwarder

รวมเครื่องมือตรวจเทียบ MANIFEST กับ ENTER ของ Forwarder ทั้ง 12 รายไว้ในเว็บเดียว (รูปแบบเดียวกับ [WWW-HAL](https://rattanao.github.io/WWW-HAL/))

เปิดใช้งาน: https://rattanao.github.io/Check-MNF-Forwarder/

| เมนู | ต้นฉบับ |
|---|---|
| Able | [check-edi-Able](https://github.com/Rattanao/check-edi-Able) |
| AGN | [check-edi-AGN](https://github.com/Rattanao/check-edi-AGN) |
| Cargoport | [Cargoport](https://github.com/Rattanao/Cargoport) |
| CMS | [CMS](https://github.com/Rattanao/CMS) |
| DSV | [check-edi-DSV](https://github.com/Rattanao/check-edi-DSV) |
| Gateway | [Gateway](https://github.com/Rattanao/Gateway) |
| M-CORP | [M-CORP](https://github.com/Rattanao/M-CORP) |
| MAC-NELS | [MAC-NELS](https://github.com/Rattanao/MAC-NELS) |
| Penanshin | [Penanshin](https://github.com/Rattanao/Penanshin) |
| Quality | [check-edi-Quality](https://github.com/Rattanao/check-edi-Quality) |
| Shipco | [check-edi-Shipco](https://github.com/Rattanao/check-edi-Shipco) |
| Union World | [Union-world](https://github.com/Rattanao/Union-world) |

แต่ละเครื่องมืออยู่ใน `tools/<ชื่อ>/` (index.html + core.js คัดลอกจากต้นฉบับ เพิ่มเฉพาะบรรทัด `<script src="../embed.js">`)
ไฟล์ทั้งหมดประมวลผลในเบราว์เซอร์ ไม่มีการอัปโหลด

## อัปเดตเครื่องมือจาก repo ต้นฉบับ

```bash
bash sync.sh
```

แล้ว commit + push · เพิ่ม Forwarder ใหม่: เพิ่มบรรทัดท้าย `sync.sh` และเพิ่มเมนูใน `index.html` (GROUPS)
