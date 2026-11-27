#!/usr/bin/env bash
# ดึงเครื่องมือเวอร์ชันล่าสุดจาก repo ต้นฉบับมาไว้ใน tools/<ชื่อ>/
# ใช้: bash sync.sh   (ต้องมี git) แล้ว commit + push
# เพิ่มเครื่องมือใหม่: เพิ่มบรรทัด "<ชื่อโฟลเดอร์> <repo>" ด้านล่าง, เพิ่มเมนูใน index.html (GROUPS)
#   และชื่อไฟล์ดาวน์โหลดใน tools/embed.js (NAMES)
set -e
cd "$(dirname "$0")"
tmp=$(mktemp -d)
trap 'rm -rf "$tmp"' EXIT

while read -r id repo; do
  [ -z "$id" ] && continue
  git clone -q --depth 1 "https://github.com/Rattanao/$repo.git" "$tmp/$id"
  src="$tmp/$id/docs"
  { [ -f "$src/index.html" ] && [ ! -d "$tmp/$id/app" ]; } || src="$tmp/$id"
  rm -rf "tools/$id" && mkdir -p "tools/$id"
  if [ "$src" = "$tmp/$id/docs" ]; then
    cp -r "$src"/. "tools/$id/"        # docs/ ทั้งโฟลเดอร์ (core.js, โลโก้, assets/)
  elif [ -d "$src/app" ]; then
    # repo ที่เวปอยู่ที่ root + app/ + scripts/*.py รันด้วย Python ในเบราว์เซอร์ (Various-forms)
    cp -r "$src/app" "tools/$id/app"
    mkdir -p "tools/$id/scripts" && cp "$src"/scripts/*.py "tools/$id/scripts/"
  else
    cp "$src/core.js" "tools/$id/core.js"   # repo ที่วางเวปไว้ที่ root (Cargoport)
  fi
  # เพิ่ม embed.js ต่อจาก <title> (ซ่อนหัว/ท้าย, จำโฟลเดอร์บันทึก, ปุ่มลบไฟล์)
  sed 's#^\(<title>.*</title>\)\r\{0,1\}$#\1\n<script src="../embed.js"></script>#' "$src/index.html" > "tools/$id/index.html"
  grep -q 'src="../embed.js"' "tools/$id/index.html" || { echo "ใส่ embed.js ไม่ได้: $id" >&2; exit 1; }
  echo "$id <- $repo $(git -C "$tmp/$id" log -1 --format='%h %cs')"
done <<'EOF'
all-forms Various-forms
able check-edi-Able
agn check-edi-AGN
cargoport Cargoport
cms CMS
dsv check-edi-DSV
eastern-air Eastern-Air
freight-links Freight-links
gateway Gateway
handle-inter Handle-inter
leo-global Leo-Globa
lx-pantos LX-Pantos
m-corp M-CORP
m-inter M-INTER
mac-nels MAC-NELS
orca Orca
penanshin Penanshin
planet Planet
quality check-edi-Quality
shipco check-edi-Shipco
sinotrans Sinotrans
union-world Union-world
EOF
