// แปลหน้าเครื่องมือเป็นภาษาอังกฤษตอนแสดงผล (ไม่แก้ไฟล์ต้นฉบับ ซึ่ง sync.sh ดึงทับทุกครั้ง)
// แปลข้อความบนหน้า, placeholder/title, ข้อความแจ้งเตือน, ตารางผล และไฟล์ EDI.xlsx ที่ดาวน์โหลด
// ข้อความของลูกค้าในเอกสาร (เช่น "เร่งเปิดตู้เข้าโกดังด่วน") ไม่แปล เพราะเป็นข้อมูลจาก ENTER
// เพิ่มคำใหม่: เพิ่มแถวใน RULES (เรียงจากวลียาว/เจาะจงก่อน คำสั้นไว้ท้าย)
(function () {
  const TH = /[\u0E00-\u0E7F]/;
  const OWNER = "([A-Z][A-Za-z0-9 .&/-]*?(?: \\([A-Za-z ]+\\))?)";
  const END = "(?=\\)|\\s\\(|:|\\s·|\\s—|\\s*$)";
  const R = (re, to) => [new RegExp(re, "g"), to];

  const RULES = [
    // ---- หัว/คำอธิบายในไฟล์ EDI.xlsx (ก่อนคำสั้นด้านล่าง) ----
    R(`รายงานตรวจสอบเอกสารขาเข้าเรือ  EDI  —  เทียบ ENTER \\(เอกสารลูกค้า\\) กับ MANIFEST \\(สายเรือ\\)  \\|  จุดสำคัญที่มีผลกับใบขนสินค้าขาเข้า`,
      "Import vessel document check  EDI  —  ENTER (customer document) vs MANIFEST (carrier)  |  points that affect the import customs declaration"),
    R(`สี: ขาว=ไม่ผิด \\| เขียว=ผ่าน \\(ช่องสรุป\\) \\| แดง ⚠=ไม่ตรง/จุดเร่งด่วน \\| ส้ม=ต้องยืนยัน \\| เหลือง=ไม่มี ENTER เทียบ`,
      "Colours: white=no error | green=pass (summary cell) | red ⚠=mismatch/urgent | orange=needs confirming | yellow=no ENTER to compare"),
    R(`STATUS: CY,CY/CY,FCL,ลากตู้=CY ; LCL,ขน\\(ส่ง\\)=LCL ; CFS,LCL/CFS,เปิดตู้=LCL/CFS`, "STATUS: CY,CY/CY,FCL,ลากตู้(haul)=CY ; LCL,ขน(ส่ง)(carry)=LCL ; CFS,LCL/CFS,เปิดตู้(unstuff)=LCL/CFS"),
    R(`SHED NO\\.: ต้องเป็น (\\S+) ทุก B/L \\(ตรง = -\\) ยกเว้น LAOS=0124 \\(DISCHARGE BANGKOK\\)`, "SHED NO.: must be $1 for every B/L (match = -) except LAOS=0124 (DISCHARGE BANGKOK)"),
    R(`SHED NO\\.: USED ENGINE/LAOS=0126/0124 \\(เฉพาะ DISCHARGE BANGKOK\\), DG=2826 \\(เฉพาะ LAEM CHABANG\\)`, "SHED NO.: USED ENGINE/LAOS=0126/0124 (DISCHARGE BANGKOK only), DG=2826 (LAEM CHABANG only)"),
    R(`CARGO MOVEMENT: TRANSHIPMENT ต้องเป็น \\(3-TRANSHIPMENT\\) เสมอ ; TRANSIT ต้องเป็น \\(7-TRANSIT\\) เสมอ ยกเว้นไปลาวต้องตรงกับ ENTER เท่านั้น`,
      "CARGO MOVEMENT: TRANSHIPMENT must always be (3-TRANSHIPMENT) ; TRANSIT must always be (7-TRANSIT), except Laos, which must match the ENTER"),

    // ---- ข้อความแจ้งเตือน ----
    R(`ไฟล์นี้ไม่ใช่ ?(?:ENTER )?(.+?) \\(ไม่พบ ?(.+?)\\) — หน้านี้รองรับเฉพาะ ?(.+?)(?: \\(และ CMS ซึ่งเป็นฟอร์มเดียวกัน\\))?:`, (m, f, nf, o) =>
      `This file is not the ${form(f)} (not found: ${notFound(nf)}) — this page supports only the ${form(o.replace(/^ฟอร์ม(?:ของ)? ?/, "ฟอร์ม ")).replace(/^(?!.* form$)(.*)$/, "$1 form")}${/และ CMS/.test(m) ? " (and CMS, which uses the same form)" : ""}:`),
    R(`อ่านไฟล์ (.+?) ไม่ได้: `, "Could not read file $1: "),
    R(`อ่าน B/L จาก ENTER ไม่ได้ — เลข B/L ใน ENTER ไม่ตรงกับรูปแบบเลข B/L ใน MANIFEST`, "Could not read any B/L from the ENTER — its B/L numbers do not match the MANIFEST's B/L number pattern"),
    R(`อ่าน B/L จาก ENTER ไม่ได้ — ไฟล์นี้ไม่ใช่แบบฟอร์ม AMENDMENT ของ Cargoport ที่หน้านี้รองรับ`, "Could not read any B/L from the ENTER — this file is not the Cargoport AMENDMENT form this page supports"),
    R(`อ่าน B/L จาก MANIFEST ไม่ได้ ตรวจว่าเป็นรายงาน CARGO MANIFEST ของสายเรือ`, "Could not read any B/L from the MANIFEST — check that it is the carrier's CARGO MANIFEST report"),
    R(`เกิดข้อผิดพลาด: `, "Error: "),

    // ---- หัวหน้า / ช่องใส่ไฟล์ ----
    R(`ตรวจเทียบ MANIFEST`, "Compare the MANIFEST"),
    R(` กับ ENTER \\((.+?)\\):\\s*`, (m, f) => ` with the ENTER (${form(f)}): `),
    R(`^ไฟล์ MANIFEST \\((.+?)\\) และ ENTER (.+?) \\(\\.pdf`, (m, a, f) => `MANIFEST file (${a}) and ${form(f)} ENTER (.pdf`),
    R(` — ไฟล์ละ B/L ใส่พร้อมกันได้หลายไฟล์`, " — one B/L per file, several files at once"),
    R(` — ไฟล์เดียวหรือหลายไฟล์ก็ได้`, " — one or several files"),
    R(` — ได้หลายไฟล์`, " — several files OK"),
    R(`ลากไฟล์มาวางพร้อมกัน หรือ`, "Drop the files here together, or"),
    R(`คลิกเพื่อเลือกไฟล์`, "click to choose files"),
    R(`SHED NO\\. ที่แจ้ง \\(ถ้ามี\\):`, "SHED NO. given (optional):"),
    R(`— ทุก B/L ต้องตรงเลขนี้ ยกเว้น LAOS`, "— every B/L must match it, except LAOS"),
    R(`^เช่น 0141$`, "e.g. 0141"),
    R(`^ล้างไฟล์และผลลัพธ์ทั้งหมด เริ่มใหม่$`, "Remove all files and results and start again"),
    R(`^คู่มือการใช้งาน / ดาวน์โหลดโปรแกรม$`, "User guide / download the program"),

    // ---- ย่อหน้ากฎของแต่ละหน้า ----
    R(`หน้านี้รองรับเฉพาะ ENTER `, "This page supports only the ENTER "),
    R(`หน้านี้รองรับเฉพาะ`, "This page supports only the "),
    R(`\\(ไฟล์ CNTRS ไม่นำมาเทียบ\\)`, "(CNTRS files are not compared)"),
    R(`; ไฟล์ CNTRS ไม่นำมาเทียบ\\)`, "; CNTRS files are not compared)"),
    R(`แยกประเภทไฟล์ให้เองจากนามสกุล \\(ตั้งชื่ออะไรก็ได้\\)`, "file types are sorted by extension (any file name works)"),
    R(`แยกประเภทไฟล์ให้เอง \\(ตั้งชื่ออะไรก็ได้\\):`, "file types are sorted automatically (any file name works):"),
    R(`และ PDF ที่เป็น CARGO MANIFEST = MANIFEST, PDF อื่น = ENTER`, "and a PDF CARGO MANIFEST = MANIFEST, any other PDF = ENTER"),
    R(`MANIFEST แบบ PDF ไม่พิมพ์ช่อง ประเทศปลายทาง / CARGO MOVEMENT — B/L ที่ TRANSIT จะขึ้น`, "a PDF MANIFEST does not print DESTINATION COUNTRY / CARGO MOVEMENT — TRANSIT B/Ls show"),
    R(`^สีส้ม$`, "orange"),
    R(`ให้ตรวจในระบบเอง`, "— check them in the system"),
    R(`ENTER ไฟล์ละ B/L ของเรือลำเดียวกันใส่พร้อมกันได้ทุกไฟล์`, "one-B/L ENTER files for the same vessel can all go in at once"),
    R(`ENTER ของเรือลำเดียวกันใส่พร้อมกันได้หลายไฟล์`, "several ENTER files for the same vessel can go in at once"),
    R(`ENTER หลายไฟล์ของเรือลำเดียวกันใส่พร้อมกันได้`, "several ENTER files for the same vessel can go in at once"),
    R(`ป้าย `, "labels "),
    R(`หน้าละ 1 B/L`, "one B/L per page"),
    R(`หัวตาราง `, "table heading "),
    R(`ในคอลัมน์ MARKS`, "in the MARKS column"),
    R(`ใต้เลขตู้`, "under the container number"),
    R(`: CY/CY-CY/FCL/ลากตู้ =`, ": CY/CY-CY/FCL/ลากตู้ (haul container) ="),
    R(`, LCL/ขน =`, ", LCL/ขน (carry) ="),
    R(`, CFS/LCL-CFS/เปิดตู้ =`, ", CFS/LCL-CFS/เปิดตู้ (unstuff) ="),
    R(`เทียบเข้ม \\(ตัดแค่ช่องว่าง/วรรคตอน\\)`, "compared strictly (only spaces/punctuation ignored)"),
    R(`พบฝั่งใดก็เตือนเสมอ`, "always flagged when found on either side"),
    R(`STATUS อ่านจากบรรทัด`, "STATUS is read from the line"),
    R(`ของแต่ละหน้า`, "of each page"),
    R(`B/L ที่มีหลายตู้ ต้องตรงกับ MANIFEST ครบทุกตู้`, "a B/L with several containers must match the MANIFEST for every container"),
    R(`ข้อความสีแดงที่ลูกค้าพิมพ์ทับ PDF ใช้กับ B/L ที่อยู่ใต้ข้อความ`, "red text the customer typed over the PDF applies to the B/Ls below it"),
    R(`ข้อความสีแดงที่ลูกค้าพิมพ์ทับ PDF ใช้กับ B/L ของหน้านั้น`, "red text the customer typed over the PDF applies to the B/L on that page"),
    R(`ข้อความสีแดงที่ลูกค้าพิมพ์ทับใช้กับ B/L ของหน้านั้น`, "red text the customer typed over it applies to the B/L on that page"),
    R(`ข้อความสีแดงที่ลูกค้าพิมพ์ทับ เช่น`, "red text the customer typed over it, e.g."),
    R(`ข้อความสีแดงที่พิมพ์ใต้บรรทัด TOTAL \\(เช่น`, "red text typed under the TOTAL line (e.g."),
    R(`\\) = STATUS ของทุก B/L`, ") = the STATUS of every B/L"),
    R(`\\(ไม่นับตรา`, "(ignoring the stamp"),
    R(`ท้ายหน้า\\)`, "at the page foot)"),
    R(`\\(ถ้ามีคำว่า`, "(if it says"),
    R(`ใช้กับทุก B/L\\)`, "it applies to every B/L)"),
    R(`ใช้กับทุก B/L`, "applies to every B/L"),
    R(`ถ้าถูกต้องตามกฎจะแสดง`, "show this when correct:"),
    R(`ข้อจำกัด:`, "Limits:"),
    R(`ไฟล์ถูกประมวลผลในเบราว์เซอร์ของคุณเท่านั้น ไม่มีการอัปโหลดไปที่ใด`, "Files are processed only in your browser — nothing is uploaded anywhere"),
    R(`= วิกฤต`, "= critical"),

    // ---- ผลตรวจ ----
    R(`^สี:$`, "Colours:"),
    R(`^ไม่ผิด$`, "no error"), R(`^ผ่าน$`, "pass"), R(`^ไม่ตรง ⚠$`, "mismatch ⚠"), R(`^ต้องยืนยัน$`, "needs confirming"),
    R(`^ไม่มี ENTER เทียบ$`, "no ENTER to compare"),
    R(`แสดงเฉพาะที่ไม่ตรง`, "Show mismatches only"),
    R(`ดาวน์โหลด EDI\\.xlsx`, "Download EDI.xlsx"),
    R(`ค้นหา B/L / ผู้รับ / ตู้ / หมายเหตุ`, "Search B/L / consignee / container / note"),
    R(`^ไม่รู้จัก$`, "unknown"), R(`^เอาออก$`, "Remove"),
    R(`\\(ไม่ระบุ\\)`, "(not given)"),
    R(`ทั้งหมด (\\d+) B/L`, "All $1 B/L"),
    R(`ต้องตรวจ (\\d+)$`, "To check $1"),
    R(`✔ ผ่าน (\\d+)`, "✔ Pass $1"),
    R(`ยอดรวมตรงกัน แต่มีจุดต้องตรวจ:`, "Totals match, but there are points to check:"),
    R(`ยอดรวมตรงกัน`, "Totals match"),
    R(`ยอดรวมไม่ตรง — ตรวจด่วน`, "Totals do not match — check urgently"),
    R(`ยอดรวมไม่ตรง`, "Totals do not match"),
    R(`ยอดรวม ENTER = MANIFEST`, "ENTER total = MANIFEST total"),
    R(`ยอดที่ระบุในเอกสาร — `, "Totals stated in the documents — "),
    R(`ตรวจยอดรวมจากรายการข้างบน:`, "Totals summed from the rows above:"),
    R(`TOTAL / ยอดรวม`, "TOTAL"),
    R(`สรุปผลเร่งด่วน`, "Urgent summary"),
    R(`หมายเหตุ / จุดที่ไม่ตรงกัน`, "Notes / mismatches"),
    R(`ประเทศปลายทาง`, "DESTINATION COUNTRY"),
    R(`(\\d+) B/L ย่อย`, "$1 sub-B/Ls"),
    R(`(\\d+) ตู้: `, (m, n) => `${n} container${n === "1" ? "" : "s"}: `),
    R(`พบข้อความคล้ายเลข B/L นี้ใน ENTER \\(ความมั่นใจต่ำ\\) แต่ไม่มีใน MANIFEST — ตรวจสอบด้วยคนว่าเป็น B/L จริงหรือไม่`, "text like this B/L number found in the ENTER (low confidence) but not in the MANIFEST — check by hand whether it is a real B/L"),
    R(`มีใน ENTER แต่ไม่มีใน MANIFEST — เสี่ยงตกหล่นจากใบขน`, "in the ENTER but not in the MANIFEST — at risk of being left off the declaration"),
    R(`B/L นี้ไม่พบใน MANIFEST — ดึงข้อมูลมาด้วยความมั่นใจต่ำ ตรวจสอบด้วยคนก่อนแจ้งสายเรือ/ลูกค้า`, "this B/L is not in the MANIFEST — read with low confidence; check by hand before telling the carrier/customer"),
    R(`B/L นี้ไม่มีอยู่ใน MANIFEST เลย — ของอาจตกหล่นจากใบขนทั้งรายการ ต้องแจ้งสายเรือ/ลูกค้าด่วน`, "this B/L is not in the MANIFEST at all — the cargo may be left off the declaration entirely; tell the carrier/customer urgently"),
    R(`อยู่ใน ENTER \\(layout ไม่รู้จัก\\) แต่ไม่มีใน MANIFEST — ตรวจสอบด้วยคน`, "in the ENTER (unknown layout) but not in the MANIFEST — check by hand"),
    R(`อยู่ใน ENTER แต่ไม่มีใน MANIFEST เลย \\(วิกฤต\\)`, "in the ENTER but not in the MANIFEST at all (critical)"),
    R(`B/L ที่มีใน ENTER แต่ไม่มีใน MANIFEST`, "A B/L in the ENTER but not in the MANIFEST"),
    R(`ไม่มีเอกสาร ENTER เทียบ`, "no ENTER document to compare"),
    R(`ไม่มีเอกสาร ENTER`, "no ENTER document"),
    R(`ไม่พบ B/L นี้ในเอกสาร ENTER`, "this B/L is not in the ENTER"),
    R(`\\(ไม่ระบุใน ENTER\\)`, "(not stated in the ENTER)"),
    R(`SHED NO\\. ไม่ตรง — `, "SHED NO. mismatch — "),
    R(`SHED ที่แจ้ง`, "the given SHED"),
    R(`DG ที่ DISCHARGE LAEM CHABANG`, "DG discharged at LAEM CHABANG"),
    R(`LAOS ที่ DISCHARGE BANGKOK`, "LAOS discharged at BANGKOK"),
    R(`USED ENGINE ที่ DISCHARGE BANGKOK`, "USED ENGINE discharged at BANGKOK"),
    R(`ตาม ENTER = `, "per the ENTER = "),
    R(`\\(ต้องเทียบ ENTER\\)`, "(compare with the ENTER)"),
    R(`ไม่มีใน MANIFEST PDF`, "not in the PDF MANIFEST"),
    R(`ประเทศปลายทาง — MANIFEST PDF ไม่พิมพ์ช่องนี้ ต้องตรวจในระบบว่าเป็น`, "DESTINATION COUNTRY — the PDF MANIFEST does not print this field; check in the system that it is"),
    R(`DESTINATION COUNTRY — MANIFEST PDF ไม่พิมพ์ช่องนี้ ต้องตรวจในระบบว่าเป็น`, "DESTINATION COUNTRY — the PDF MANIFEST does not print this field; check in the system that it is"),
    R(`CARGO MOVEMENT — MANIFEST PDF ไม่พิมพ์ช่องนี้ ต้องตรวจในระบบว่าเป็น`, "CARGO MOVEMENT — the PDF MANIFEST does not print this field; check in the system that it is"),
    R(`CARGO MOVEMENT ไปลาว — ไม่พบค่า CARGO MOVEMENT ใน ENTER ของ B/L นี้ ต้องตรวจสอบด้วยคน`, "CARGO MOVEMENT to Laos — no CARGO MOVEMENT in this B/L's ENTER; check by hand"),
    R(`CARGO MOVEMENT ไม่ตรง — `, "CARGO MOVEMENT mismatch — "),
    R(`พบคำ transit/transhipment — (.+?) \\(ไม่ใช่ LOCAL, ตรวจประเภทใบขน: ถ่ายลำ/ผ่านแดน\\)`, "transit/transhipment found — $1 (not LOCAL; check the declaration type: transhipment/transit)"),
    R(`ต้องสำแดงตู้เย็น/ยืนยันอุณหภูมิในใบขน`, "declare as reefer / confirm the temperature in the declaration"),
    R(`ต้องสำแดงวัตถุอันตราย \\(CLASS/UN\\) \\+ แนบ MSDS/DGD`, "declare as dangerous goods (CLASS/UN) + attach MSDS/DGD"),
    R(`อีกฝั่งไม่ระบุ ; `, "the other side does not state it ; "),
    R(` — MANIFEST = 0\\.000 \\(ค่าหาย\\) แต่ ENTER มีค่าจริง`, " — MANIFEST = 0.000 (missing) but the ENTER has a value"),
    R(`ไม่ตรงกัน \\(ไม่นับช่องว่าง\\)`, "does not match (spaces ignored)"),
    R(`ไม่เหมือนกัน`, "differs"),
    R(`ENTER เอกสารนี้ดึงข้อมูลด้วยความมั่นใจต่ำ \\(layout ไม่ชัดเจน/พบ anchor น้อยเกินไป\\) — ตรวจสอบทุกช่องที่ไฮไลต์ส้มด้วยคน`, "this ENTER was read with low confidence (unclear layout / too few anchors) — check every orange cell by hand"),
    R(`ต้องตรวจสอบด้วยคน`, "check by hand"),
    R(`ต้องตรวจสอบ: `, "to check: "),
    R(`ต้องตรวจ (\\d+) B/L`, "to check $1 B/L"),
    R(`ต้องตรวจ`, "to check"),
    R(`ต้องเป็น`, "must be"),
    R(`ปลายทางไม่ตรง`, "destination mismatch"),
    R(`ปลายทาง`, "DESTINATION"),
    R(`PACKAGE ตรง`, "PACKAGE match"), R(`GROSS WEIGHT ตรง`, "GROSS WEIGHT match"), R(`MEASUREMENT ตรง`, "MEASUREMENT match"),
    R(` ไม่ตรง`, " mismatch"),
    R(`ไม่ตรง: `, "mismatch: "),
    R(`ไม่ตรง`, "mismatch"),
    R(`\\(ว่าง\\)`, "(blank)"),
    R(`\\(ไม่มี\\)`, "(none)"),
    R(`^ตรวจสอบ$`, "Check"),
    R(`^ลบไฟล์$`, "Clear files"),
  ];

  // สิ่งที่ตัวตรวจหาไม่พบ: "หัวตาราง 'B/L CHANGE NO.'" -> "table heading 'B/L CHANGE NO.'"
  function notFound(s) {
    return s.trim()
      .replace(/หัวตาราง ?/g, "table heading ").replace(/หัวเอกสาร ?/g, "document heading ").replace(/ป้าย ?/g, "labels ")
      .replace(/หน้า ?/g, "page ").replace(/หัว ?/g, "heading ").replace(/ ?ของฟอร์ม/g, " of the form")
      .replace(/ ?ในคอลัมน์ ?/g, " in column ").replace(/ ?ใต้เลขตู้/g, " under the container number").replace(/ และ /g, " and ")
      .replace(/^document heading$/, "document heading").replace(/\s+/g, " ").trim();
  }
  // ชื่อฟอร์ม: "ฟอร์ม AMENDMENT ของ AGN" -> "AGN AMENDMENT form"
  function form(f) {
    return f.trim()
      .replace(/^ENTER /, "")
      .replace(/^ฟอร์ม (.+?) ของ (.+)$/, "$2 $1 form")
      .replace(/^ฟอร์มของ (.+)$/, "$1 form")
      .replace(/^ฟอร์ม (.+)$/, "$1 form")
      .replace(/^(?:แบบฟอร์ม )?AMENDMENT ของ (.+)$/, "$1 AMENDMENT")
      .replace(/^'AMENDMENT' ของ (.+)$/, "$1 'AMENDMENT'")
      .replace(/^จดหมายแก้ MANIFEST ของ (.+)$/, "$1 manifest correction letter")
      .replace(/^ของ (.+)$/, "$1")
      .replace(/ หรือ /g, " or ")
      .replace(/ form form$/, " form");
  }
  // ชิ้นที่ RULES ข้างบนไม่ได้แปลทั้งประโยค: ชื่อฟอร์มที่เหลือในวงเล็บหรือหลัง "the"
  const TAIL = [
    [new RegExp(`(?:ENTER )?ฟอร์ม (.+?) ของ ${OWNER}${END}`, "g"), (m, a, b) => `${b} ${a} form`],
    [new RegExp(`(?:ENTER )?ฟอร์มของ ${OWNER}${END}`, "g"), (m, a) => `${a} form`],
    [new RegExp(`ENTER จดหมายแก้ MANIFEST ของ ${OWNER}${END}`, "g"), (m, a) => `${a} manifest correction letter`],
    [new RegExp(`ENTER 'AMENDMENT' ของ ${OWNER}${END}`, "g"), (m, a) => `${a} 'AMENDMENT' ENTER`],
    [new RegExp(`ฟอร์ม (\\[[^\\]]+\\]):`, "g"), "form $1:"],
    [/ หรือ /g, " or "], [/ และ(?= |$)/g, " and"], [/^ทุก BL$/, "ทุก BL (all B/Ls)"],
  ];

  function tr(s) {
    if (typeof s !== "string" || !TH.test(s)) return s;
    // แปลเฉพาะเนื้อความ เก็บช่องว่างหน้า/หลังไว้ (ข้อความบนหน้าเว็บมักมีช่องว่างคั่นกับแท็กข้าง ๆ)
    const [, lead, body, trail] = s.match(/^(\s*)([\s\S]*?)(\s*)$/);
    s = body;
    for (const [re, to] of RULES) { re.lastIndex = 0; s = s.replace(re, to); if (!TH.test(s)) return lead + s + trail; }
    for (const [re, to] of TAIL) s = s.replace(re, to);
    return lead + s + trail;
  }
  window.__toEnglish = tr;

  // ---- หน้าเว็บ ----
  const done = new WeakMap();
  function node(n) {
    if (n.nodeType === 3) {
      const p = n.parentElement;
      if (!p || p.closest("script,style,textarea")) return;
      if (done.get(n) === n.nodeValue) return;
      const v = tr(n.nodeValue);
      if (v !== n.nodeValue) n.nodeValue = v;
      done.set(n, n.nodeValue);
    } else if (n.nodeType === 1) {
      for (const a of ["placeholder", "title", "aria-label"]) { const v = n.getAttribute(a); if (v && TH.test(v)) n.setAttribute(a, tr(v)); }
      if (n.tagName !== "SCRIPT" && n.tagName !== "STYLE") n.childNodes.forEach(node);
    }
  }
  function start() {
    document.documentElement.lang = "en";
    node(document.body);
    new MutationObserver(ms => ms.forEach(m => {
      if (m.type === "characterData") node(m.target);
      else if (m.type === "attributes") node(m.target);
      else m.addedNodes.forEach(node);
    })).observe(document.body, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ["placeholder", "title"] });
    if (TH.test(document.title)) document.title = tr(document.title);
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start); else start();

  // ---- ไฟล์ EDI.xlsx: แปลค่าข้อความในทุกช่องก่อนเขียนไฟล์ ----
  const excel = () => {
    if (!window.ExcelJS) return;
    try {
      const cell = new ExcelJS.Workbook().addWorksheet("x").getCell(1, 1);
      let proto = Object.getPrototypeOf(cell), d;
      while (proto && !(d = Object.getOwnPropertyDescriptor(proto, "value"))) proto = Object.getPrototypeOf(proto);
      if (!d || !d.set) return;
      Object.defineProperty(proto, "value", { get: d.get, set(v) { d.set.call(this, tr(v)); }, configurable: true });
    } catch (e) { console.warn("en.js: Excel", e); }
  };
  if (document.readyState === "complete") excel(); else addEventListener("load", excel);
})();
