# Turtle23 Messenger Log — auto-sync จาก SharePoint

หน้าเว็บดูข้อมูลการจองงาน Messenger + ใบงานพิมพ์รายวัน ที่ซิงก์ข้อมูลจากไฟล์
`TURTLE GROUP MESSENGER 2025_🛵.xlsx` บน SharePoint โดยอัตโนมัติผ่าน
GitHub Actions (ไม่ต้องพึ่งคอมพิวเตอร์เครื่องไหนเปิดอยู่)

ทุกครั้งที่ workflow รัน จะทำ 2 อย่าง:
1. **อัปเดตหน้าเว็บ** (`docs/index.html`) ให้เป็นข้อมูลล่าสุด
2. **สร้าง/อัปเดตไฟล์ใบงานของวันนี้** (`MESSENGER {d}-{m}-{yyyy}.xlsx`) แล้วอัปโหลด
   กลับไปวางไว้ในโฟลเดอร์ SharePoint เดียวกับไฟล์ template `MESSENGER From 1 2 (2).xlsx`
   (คนละไฟล์กับ template เดิม — ไม่แตะ/ไม่ทับไฟล์ template) รันซ้ำในวันเดียวกันจะ
   อัปเดตไฟล์ของวันนั้นให้ล่าสุดเรื่อยๆ ตามข้อมูลที่เข้ามาใหม่

## โครงสร้าง repo

```
.github/workflows/sync.yml     ← workflow ที่รันตามตารางเวลา
scripts/sync_data.py           ← ดึงไฟล์จาก SharePoint, แปลงเป็น JSON, สั่ง build เว็บ + อัปโหลดใบงาน
scripts/job_sheet.py           ← สร้างไฟล์ใบงาน Excel 4 ชีต (ส่งเอกสาร/รับฝากเอกสารกลับ x เช้า/บ่าย)
scripts/messenger_log.tpl.html ← เทมเพลตหน้าเว็บ (ห้ามลบ placeholder __BOOKING_DATA_JSON__ ฯลฯ)
scripts/requirements.txt
docs/index.html                ← หน้าเว็บที่ build เสร็จแล้ว (GitHub Pages เสิร์ฟจากตรงนี้)
docs/booking_data.json         ← ข้อมูลล่าสุดในรูปแบบ JSON (เผื่อใช้ต่อ)
```

## ตั้งค่าครั้งแรก (ทำ 2 อย่าง)

### 1) เพิ่ม GitHub Secrets

ไปที่ **Settings → Secrets and variables → Actions → New repository secret**
แล้วเพิ่มทั้ง 5 ตัวนี้:

| ชื่อ Secret | ค่า |
|---|---|
| `AZURE_TENANT_ID` | Tenant ID ของ Turtle23 ใน Azure AD |
| `AZURE_CLIENT_ID` | Application (client) ID ของ App Registration |
| `AZURE_CLIENT_SECRET` | ค่า Client Secret (ไม่ใช่ Secret ID) |
| `SHAREPOINT_FILE_URL` | ลิงก์แชร์ไฟล์ queue `TURTLE GROUP MESSENGER 2025_🛵.xlsx` |
| `SHAREPOINT_TEMPLATE_URL` | ลิงก์แชร์ไฟล์ `MESSENGER From 1 2 (2).xlsx` — ใช้แค่หาตำแหน่งโฟลเดอร์ปลายทาง (ไม่ได้อ่าน/แก้ไฟล์นี้เลย) |

> **สิทธิ์ที่ App Registration ต้องมี (API permissions, ประเภท Application) พร้อม Grant admin consent แล้ว:**
> ต้องมีทั้งอ่านและเขียน เพราะตอนนี้ต้องอัปโหลดไฟล์ใบงานกลับขึ้น SharePoint ด้วย —
> ใช้ **`Sites.ReadWrite.All`** (แนะนำ ครอบคลุมทั้งองค์กรรวมถึง OneDrive ส่วนตัว)
> หรือ `Files.ReadWrite.All` ก็ได้ — แค่ `Sites.Read.All`/`Files.Read.All` (อ่านอย่างเดียว)
> **ไม่พอ** จะอัปโหลดใบงานไม่ได้ (ถ้าต้องการแค่ซิงก์หน้าเว็บโดยไม่อัปโหลดใบงาน ดูหัวข้อ
> "ปิดการอัปโหลดใบงาน" ด้านล่าง แล้วใช้สิทธิ์อ่านอย่างเดียวพอ)
>
> เนื่องจากไฟล์ทั้งสองอยู่ใน OneDrive ส่วนตัว (`personal/sarocha_k_...`) สิทธิ์แบบ
> `Sites.ReadWrite.All` (ครอบคลุมทั้งองค์กร) จะเข้าถึง/เขียนได้แน่นอน — ถ้าใช้
> `Sites.Selected` แทน ต้องให้แอดมิน grant สิทธิ์เขียน (write) เจาะจงไปที่ไซต์ OneDrive
> ของคุณสโรชาเพิ่มเติม

### 2) เปิดใช้ GitHub Pages

ไปที่ **Settings → Pages** → ในหัวข้อ "Build and deployment" เลือก
**Source: Deploy from a branch** → Branch: `main` และโฟลเดอร์ `/docs` → Save

หลัง workflow รันครั้งแรกสำเร็จ หน้าเว็บจะอยู่ที่ (ตัวอย่าง):
`https://<your-github-username>.github.io/<repo-name>/`

## การทำงาน

- Workflow รันอัตโนมัติทุกชั่วโมง (`0 * * * *`, เวลา UTC) — แก้ความถี่ได้ที่
  `.github/workflows/sync.yml` บรรทัด `cron:`
- กด **Actions → Sync Messenger booking data → Run workflow** เพื่อสั่งซิงก์ทันทีได้ตลอด
- ถ้าข้อมูลไม่เปลี่ยน workflow จะไม่สร้าง commit ใหม่ในหน้าเว็บ (กัน repo รกจาก commit ว่างๆ) —
  แต่ไฟล์ใบงานบน SharePoint จะยัง PUT ทับซ้ำทุกครั้งที่รัน (ไม่มีผลเสีย เพราะเนื้อหาจะเหมือนเดิมถ้าไม่มีข้อมูลใหม่)
- ถ้าดึงข้อมูลได้ 0 แถว สคริปต์จะ "ปฏิเสธ" ไม่เขียนทับ docs/ หรืออัปโหลดอะไรทั้งสิ้น
  (กันหน้าเว็บ/ใบงานว่างเปล่าจาก error ชั่วคราว) — เช็ค log ใน Actions tab ได้
- ไฟล์ใบงานตั้งชื่อ `MESSENGER {d}-{m}-{yyyy}.xlsx` (เช่น `MESSENGER 8-9-2026.xlsx`)
  แยกไฟล์ใหม่ทุกวัน โดยไม่แตะไฟล์ template เดิมเลย

### ปิดการอัปโหลดใบงาน (ถ้าต้องการแค่ซิงก์หน้าเว็บ)

ตั้ง secret `SKIP_JOB_SHEET_UPLOAD` เป็น `1` (หรือไม่ใส่ `SHAREPOINT_TEMPLATE_URL`) —
ตอนนั้นใช้สิทธิ์อ่านอย่างเดียว (`Sites.Read.All`) ได้พอ ไม่ต้องขอสิทธิ์เขียน

## ทดสอบว่า Secret ถูกต้องหรือยัง

หลังใส่ secret ครบแล้ว กด **Actions → Sync Messenger booking data → Run workflow**
แล้วดู log:
- ถ้าเจอ `Failed to obtain Graph access token` → เช็ค `AZURE_TENANT_ID` / `AZURE_CLIENT_ID` / `AZURE_CLIENT_SECRET`
- ถ้าเจอ `Failed to resolve SharePoint share to a driveItem` (403/404) → เช็คว่า admin consent สิทธิ์แล้ว หรือ URL ใน `SHAREPOINT_FILE_URL` / `SHAREPOINT_TEMPLATE_URL` ไม่ตรง/หมดอายุ
- ถ้าเจอ `Failed to upload job sheet to SharePoint` (403) → สิทธิ์ที่ให้เป็นแบบอ่านอย่างเดียว ต้องเปลี่ยนเป็น `Sites.ReadWrite.All`/`Files.ReadWrite.All` แล้วขอ admin consent ใหม่

## หมายเหตุความปลอดภัย

- Client Secret มีวันหมดอายุ (Azure AD กำหนดตอนสร้าง ปกติ 6 เดือน–2 ปี) — ต้องสร้างใหม่และอัปเดต secret ก่อนหมดอายุ
- Personal Access Token ที่ใช้ตอน push โค้ดครั้งแรกสามารถ revoke ได้ทันทีหลังตั้งค่าเสร็จ เพราะหลังจากนี้ workflow ใช้ `GITHUB_TOKEN` ในตัว (auto) สำหรับ commit ไม่ต้องพึ่ง PAT อีก
