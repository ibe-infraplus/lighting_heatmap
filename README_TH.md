# Lighting Maintenance Dashboard (MapLibre)

โปรเจกต์นี้อ่านข้อมูลจาก Excel Template แล้วสร้าง Dashboard เป็นไฟล์ HTML เดียว โดยไม่ต้องมี Backend หรือฐานข้อมูล

## ไฟล์

- `generate_lighting_dashboard.py` — โปรแกรมอ่าน Excel และสร้าง Dashboard
- `lighting_dashboard.html` — Dashboard ที่สร้างจากไฟล์ตัวอย่าง `test.xlsx`
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

เปิด `lighting_dashboard.html` ด้วย Chrome หรือ Edge ได้โดยตรง หรือเปิดผ่าน Local Server:

```bash
python -m http.server 8000
```

แล้วเปิด `http://localhost:8000/lighting_dashboard.html`

> ต้องเชื่อมต่ออินเทอร์เน็ตเพื่อโหลด MapLibre GL JS และ Basemap OpenFreeMap

## เปิดใช้งานด้วย Docker Compose

วางไฟล์ต่อไปนี้ไว้ในโฟลเดอร์เดียวกันบน Ubuntu:

- `docker-compose.yml`
- `nginx.conf`
- `lighting_dashboard.html`

เริ่มระบบด้วยคำสั่ง:

```bash
docker compose up -d
```

Compose จะ build Dashboard จาก `test.xlsx` ด้วย `generate_lighting_dashboard.py`
ภายใน Docker image โดยอัตโนมัติ

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

หากมี Reverse Proxy หลักของโดเมนชี้ path `/lighting_heatmap` มายังพอร์ต `1112`
จะเข้า Dashboard โดยไม่ต้องระบุพอร์ตได้ที่:

```text
http://apps.infra-corp.co/lighting_heatmap
```

ตรวจสอบสถานะและ log:

```bash
docker compose ps
docker compose logs -f lighting-dashboard
```

หากต้องการเปลี่ยนพอร์ต ให้กำหนด `DASHBOARD_PORT` ก่อนสั่ง Compose เช่น:

```bash
DASHBOARD_PORT=8090 docker compose up -d
```

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

- Heatmap ความหนาแน่นเฉพาะเสาที่มีมากกว่า 1 รอบซ่อม โดยหนึ่งจุดแทนหนึ่งรหัสเสา
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
