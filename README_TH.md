# Lighting Maintenance Dashboard (MapLibre + PostgreSQL Login)

โปรเจกต์นี้อ่านข้อมูลจาก Excel Template แล้วสร้าง Dashboard แผนที่ พร้อมระบบ Login
ที่ตรวจสอบบัญชีผู้ใช้จาก PostgreSQL ก่อนอนุญาตให้เปิดหน้า `/lighting_heatmap`

## ไฟล์

- `generate_lighting_dashboard.py` — โปรแกรมอ่าน Excel และสร้าง Dashboard
- `lighting_dashboard.html` — Dashboard ที่สร้างจากไฟล์ตัวอย่าง `test.xlsx`
- `app.py` — Flask Backend สำหรับ Login, Session และให้บริการ Dashboard
- `templates/login.html` — หน้า Login และโลโก้เสาไฟ/โคมไฟ
- `requirements.txt` — Python dependencies
- `.env.example` — ตัวอย่างการตั้งค่า PostgreSQL และบัญชีผู้ดูแลครั้งแรก
- `run_dashboard.bat` — ตัวช่วยสำหรับ Windows

## วิธีใช้งาน

วางไฟล์ Excel ไว้โฟลเดอร์เดียวกับโปรแกรม แล้วรัน:

```bash
python generate_lighting_dashboard.py --excel test.xlsx --output lighting_dashboard.html
```

หากข้อมูลอยู่คนละ Worksheet:

```bash
python generate_lighting_dashboard.py --excel test.xlsx --sheet Sheet1 --output lighting_dashboard.html
```

เปิด `lighting_dashboard.html` ด้วย Chrome หรือ Edge ได้โดยตรงสำหรับตรวจหน้าตาแบบไม่มี Login
หรือเปิดผ่าน Local Server:

```bash
python -m http.server 8000
```

แล้วเปิด `http://localhost:8000/lighting_dashboard.html`

> ต้องเชื่อมต่ออินเทอร์เน็ตเพื่อโหลด MapLibre GL JS และ Basemap OpenFreeMap

## ตั้งค่า Login และ PostgreSQL

ระบบจะสร้างตารางต่อไปนี้ในฐานข้อมูลที่กำหนดโดยอัตโนมัติ:

- `app_users` — บัญชีผู้ใช้ รหัสผ่านแบบ Hash สถานะบัญชี และเวลาที่เข้าสู่ระบบล่าสุด
- `app_login_audit` — ประวัติการเข้าสู่ระบบสำเร็จ/ไม่สำเร็จ

รหัสผ่าน PostgreSQL, Session Secret และรหัสผ่านผู้ดูแล **ห้าม commit ขึ้น GitHub**
ให้สร้างไฟล์ `.env` บน Ubuntu:

```bash
cd ~/lighting_dashboard_project
cp .env.example .env
nano .env
chmod 600 .env
```

กำหนดค่าฐานข้อมูลตาม Server จริง และสร้าง `SECRET_KEY` ด้วย:

```bash
python3 -c "import secrets; print(secrets.token_urlsafe(48))"
```

ค่าที่ต้องตรวจสอบใน `.env`:

```dotenv
DB_HOST=192.168.120.98
DB_PORT=5433
DB_NAME=bma_lighting_heatmap
DB_USER=postgres
DB_PASSWORD=ใส่รหัสผ่านจริง
SECRET_KEY=ใส่ค่าสุ่มอย่างน้อย_32_ตัวอักษร
APP_ADMIN_USERNAME=admin
APP_ADMIN_PASSWORD=ตั้งรหัสผ่านผู้ดูแลอย่างน้อย_10_ตัวอักษร
APP_ADMIN_DISPLAY_NAME=ผู้ดูแลระบบ
SESSION_COOKIE_SECURE=true
```

`APP_ADMIN_USERNAME` และ `APP_ADMIN_PASSWORD` ใช้สร้างผู้ดูแลเมื่อเริ่มระบบครั้งแรกเท่านั้น
หากมี Username นี้ในฐานข้อมูลแล้ว ระบบจะไม่เปลี่ยนรหัสผ่านเดิม

## เปิดใช้งานด้วย Docker Compose

หลังสร้าง `.env` แล้ว:

```bash
docker compose up -d --build
docker compose ps
```

Compose จะสร้าง Dashboard จาก `test.xlsx`, เริ่ม Flask ผ่าน Gunicorn, เชื่อม PostgreSQL
และเปิดบริการที่ Port `1112`

## อัปเดตจาก GitHub บน Ubuntu

หลัง commit และ push จาก VS Code ไปยัง branch `main` แล้ว ให้รันบน Ubuntu:

```bash
cd ~/lighting_dashboard_project
bash deploy.sh
```

สคริปต์จะดึงโค้ดด้วย `git pull --ff-only`, build image ใหม่, recreate container
และรอจน healthcheck เป็น `healthy` หาก server มีไฟล์ที่แก้ค้างอยู่ สคริปต์จะหยุดก่อน
เพื่อป้องกันการเขียนทับไฟล์

คำสั่งแบบ manual ที่ให้ผลเดียวกัน:

```bash
cd ~/lighting_dashboard_project
git pull --ff-only origin main
docker compose up -d --build --remove-orphans
docker compose ps
```

Dashboard จะเปิดที่:

```text
http://SERVER_IP:1112/lighting_heatmap
```

หากยังไม่ได้ Login ระบบจะส่งไปยัง:

```text
http://SERVER_IP:1112/lighting_heatmap/login
```

หากมี Reverse Proxy หลักของโดเมนชี้ Path `/lighting_heatmap` มายังพอร์ต `1112`
จะเข้า Dashboard โดยไม่ต้องระบุพอร์ตได้ที่:

```text
https://apps.infra-corp.co/lighting_heatmap
```

สำหรับ HTTPS ให้ตั้ง `SESSION_COOKIE_SECURE=true` หากทดสอบผ่าน HTTP โดยตรงจึงค่อยเปลี่ยนเป็น `false`

ตรวจสอบสถานะและ log:

```bash
docker compose ps
docker compose logs -f lighting-dashboard
```

หากต้องการเปลี่ยนพอร์ต ให้กำหนด `DASHBOARD_PORT` ก่อนสั่ง Compose เช่น:

```bash
DASHBOARD_PORT=8090 docker compose up -d
```

## ความปลอดภัยของระบบ Login

- เก็บรหัสผ่านด้วย Password Hash ของ Werkzeug ไม่เก็บรหัสผ่านแบบข้อความตรง
- ใช้ Server-side verification ผ่าน PostgreSQL
- ใช้ Session Cookie แบบ `HttpOnly` และ `SameSite=Lax`
- ป้องกัน Login และ Logout ด้วย CSRF Token
- เมื่อกรอกรหัสผ่านผิดครบ 5 ครั้ง บัญชีจะถูกพัก 15 นาที
- บันทึกประวัติการ Login โดยไม่บันทึกรหัสผ่าน
- หน้า Dashboard และข้อมูลภายในต้องผ่าน Login ก่อน

## คอลัมน์หลักที่โปรแกรมใช้

- รหัสเสาไฟฟ้า
- ชนิดโคม
- lat
- lon
- ประเภทความเสียหาย
- อาการที่ตรวจสอบ
- วิธีแก้ไข
- รายละเอียดเพิ่มเติม
- ระยะเวลาดำเนินการ
- ticket_id
- สถานะข้อร้องเรียน
- เขต

โปรแกรมรองรับชื่อคอลัมน์ใกล้เคียง เช่น `latitude`, `longitude`, `วิธีการซ่อม`, `สถานะงาน` ด้วย

## หลักการนับข้อมูล

- ตัดรายการที่ `วิธีแก้ไข` ระบุว่า `ไม่เสียหาย`, `ไม่พบความเสียหาย`, `ตรวจสอบแล้วปกติ` หรือ `ปกติ` ออกจาก Dashboard เพราะไม่ใช่เหตุโคมไฟเสียจริง
- ยึด `รหัสเสาไฟฟ้า` เป็นตัวระบุเสาหลัก (ใช้พิกัดแทนเฉพาะรหัสที่ว่างหรือเป็น Placeholder)
- Ticket ID คือรหัสข้อร้องเรียน จึงใช้แสดงจำนวนข้อร้องเรียนเท่านั้น
- ข้อร้องเรียนหลายรายการของเสาเดียวกันที่มีช่วงเวลาดำเนินการเดียวกัน จะรวมเป็นรอบการซ่อม/เหตุไฟดับเดียว
- เสาซ่อมซ้ำ หมายถึงรหัสเสาที่มีมากกว่า 1 รอบการซ่อม ไม่ใช่เสาที่มีมากกว่า 1 Ticket

## สิ่งที่ Dashboard แสดง

- ปุ่ม Dashboard สรุปข้อมูล เปิดหน้าวิเคราะห์ KPI, อัตราปิดงาน, Insight, เขต, วิธีแก้ไข, ประเภทความเสียหาย และสถานะตามตัวกรองปัจจุบัน โดยคลิกกราฟเพื่อดูรายชื่อและรายละเอียดเสาที่อยู่ในหมวดนั้นได้
- Drawer ตัวกรองรองรับการเลือก `เขต` ร่วมกับ `วิธีการแก้ไข` และใช้เงื่อนไขเดียวกันกับแผนที่ Cluster และ Dashboard สรุป
- โหมด `การซ่อม Heatmap` แสดงเฉพาะรหัสเสาที่มีจำนวนรอบซ่อมตั้งแต่ 2 ครั้งขึ้นไป
- จุดเสาไฟดับตามจำนวนรอบที่ไฟดับจริง
- วิธีซ่อมหลักของแต่ละเสา
- เสาที่มีมากกว่า 1 รอบซ่อม
- KPI: ข้อร้องเรียนไม่ซ้ำ, จำนวนเสา, ครั้งที่ไฟดับ, อัตราปิดงานซ่อม, Median เวลาซ่อม, เสาซ่อมซ้ำ
- อันดับเขตและวิธีซ่อม
- รายละเอียดเสาที่เลือกจากแผนที่ พร้อมประวัติ สาเหตุ จุดที่เสีย วิธีซ่อม และจำนวนข้อร้องเรียน
- Export CSV ตามตัวกรอง

## คอลัมน์แนะนำสำหรับระบบจริง

- `reported_at`, `inspected_at`, `repair_started_at`, `repair_completed_at`
- `post_repair_status` เช่น ติดปกติ / ยังไม่ติด / ติดไม่เต็มกำลัง
- `root_cause` เช่น Driver, สายไฟ, หลอด, ตู้ควบคุม, ไฟฟ้าต้นทาง
- `repeat_of_ticket_id`, `repeat_failure_date`
- `material_cost`, `labor_cost`, `spare_part`, `contractor`
- `sla_hours`, `warranty_status`
- `before_photo_url`, `after_photo_url`
