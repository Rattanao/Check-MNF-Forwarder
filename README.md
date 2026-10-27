# Check MNF Forwarder

รวมเครื่องมือตรวจเทียบ MANIFEST กับ ENTER ของ Forwarder ทั้ง 18 รายไว้ในเว็บเดียว (รูปแบบเดียวกับ [WWW-HAL](https://rattanao.github.io/WWW-HAL/))

เปิดใช้งาน: https://rattanao.github.io/Check-MNF-Forwarder/ · คู่มือ (ไทย / English): https://rattanao.github.io/Check-MNF-Forwarder/#guide

เมนูบนสุด **Check all files** (`tools/all/`) รับ MANIFEST / ENTER ของเจ้าไหนก็ได้ ลองอ่าน ENTER ด้วยตัวตรวจของทุกเจ้า เลือกเจ้าที่รู้จักฟอร์มและอ่านเลข B/L ตรงกับ MANIFEST ได้มากที่สุด แล้วแสดงผลด้วยหน้าของเจ้านั้น (เพิ่มเครื่องมือใหม่: เพิ่มใน `TOOLS` ของ `tools/all/index.html` ด้วย)

| เมนู | ต้นฉบับ |
|---|---|
| Able | [check-edi-Able](https://github.com/Rattanao/check-edi-Able) |
| AGN | [check-edi-AGN](https://github.com/Rattanao/check-edi-AGN) |
| Cargoport | [Cargoport](https://github.com/Rattanao/Cargoport) |
| CMS | [CMS](https://github.com/Rattanao/CMS) |
| DSV | [check-edi-DSV](https://github.com/Rattanao/check-edi-DSV) |
| Eastern Air | [Eastern-Air](https://github.com/Rattanao/Eastern-Air) |
| Freight Links | [Freight-links](https://github.com/Rattanao/Freight-links) |
| Gateway | [Gateway](https://github.com/Rattanao/Gateway) |
| Handle Inter | [Handle-inter](https://github.com/Rattanao/Handle-inter) |
| LX Pantos | [LX-Pantos](https://github.com/Rattanao/LX-Pantos) |
| M-CORP | [M-CORP](https://github.com/Rattanao/M-CORP) |
| MAC-NELS | [MAC-NELS](https://github.com/Rattanao/MAC-NELS) |
| Penanshin | [Penanshin](https://github.com/Rattanao/Penanshin) |
| Planet | [Planet](https://github.com/Rattanao/Planet) |
| Quality | [check-edi-Quality](https://github.com/Rattanao/check-edi-Quality) |
| Shipco | [check-edi-Shipco](https://github.com/Rattanao/check-edi-Shipco) |
| Sinotrans | [Sinotrans](https://github.com/Rattanao/Sinotrans) |
| Union World | [Union-world](https://github.com/Rattanao/Union-world) |

แต่ละเครื่องมืออยู่ใน `tools/<ชื่อ>/` (index.html + core.js คัดลอกจากต้นฉบับ เพิ่มเฉพาะบรรทัด `<script src="../embed.js">`)
ไฟล์ทั้งหมดประมวลผลในเบราว์เซอร์ ไม่มีการอัปโหลด

## อัปเดตเครื่องมือจาก repo ต้นฉบับ

```bash
bash sync.sh
```

แล้ว commit + push · เพิ่ม Forwarder ใหม่: เพิ่มบรรทัดท้าย `sync.sh` และเพิ่มเมนูใน `index.html` (GROUPS)
