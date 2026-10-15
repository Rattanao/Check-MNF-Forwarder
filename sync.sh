#!/usr/bin/env bash
# ดึงเครื่องมือเวอร์ชันล่าสุดจาก repo ต้นฉบับมาไว้ใน tools/<ชื่อ>/
# ใช้: bash sync.sh   (ต้องมี git) แล้ว commit + push
# เพิ่มเครื่องมือใหม่: เพิ่มบรรทัด "<ชื่อโฟลเดอร์> <repo>" ด้านล่าง และเพิ่มเมนูใน index.html (GROUPS)
set -e
cd "$(dirname "$0")"
tmp=$(mktemp -d)
trap 'rm -rf "$tmp"' EXIT

while read -r id repo; do
  [ -z "$id" ] && continue
  git clone -q --depth 1 "https://github.com/Rattanao/$repo.git" "$tmp/$id"
  src="$tmp/$id/docs"; [ -f "$src/index.html" ] || src="$tmp/$id"
  mkdir -p "tools/$id"
  cp "$src/core.js" "tools/$id/core.js"
  # เพิ่ม embed.js ต่อจาก <title> (ซ่อนหัว/ท้าย, จำโฟลเดอร์บันทึก, ปุ่มลบไฟล์)
  sed 's#^\(<title>.*</title>\)\r\{0,1\}$#\1\n<script src="../embed.js"></script>#' "$src/index.html" > "tools/$id/index.html"
  grep -q 'src="../embed.js"' "tools/$id/index.html" || { echo "ใส่ embed.js ไม่ได้: $id" >&2; exit 1; }
  echo "$id <- $repo $(git -C "$tmp/$id" log -1 --format='%h %cs')"
done <<'EOF'
able check-edi-Able
agn check-edi-AGN
cargoport Cargoport
cma CMA
dsv check-edi-DSV
gateway Gateway
m-corp M-CORP
mac-nels MAC-NELS
penanshin Penanshin
quality check-edi-Quality
shipco check-edi-Shipco
EOF
