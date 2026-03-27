<<<<<<< HEAD
# Run Power Automate Desktop Flow via Dataverse API

เครื่องมือนี้ช่วย "เรียก (Run)" Desktop Flow (Power Automate Desktop) ผ่าน Dataverse Web API แล้ว "เฝ้าดู (Monitor)" จนจบ พร้อมรายงานสถานะสำเร็จ/ล้มเหลว และข้อความ Error ถ้ามี

## สิ่งที่ต้องมี (Prerequisites)
- Windows + Python 3.10 ขึ้นไป
- App registration บน Microsoft Entra (Azure AD)
  - มี `Client ID`, `Client Secret`, `Tenant ID`
  - ลงทะเบียนเป็น Application User ใน Dataverse Environment เดียวกับ Flow และมอบสิทธิ์ที่เหมาะสมให้สามารถรัน Desktop Flow ได้
- ค่าแวดล้อมของ Environment/Flow
  - `DATAVERSE_URL` เช่น `https://orge8e671e1.crm5.dynamics.com`
  - `WORKFLOW_ID` เป็น GUID ของ Flow: `1015b2f8-5575-45dd-b1ba-adca4f1f5957`

## การติดตั้ง
```bash
cd "d:\Python Project\Run Power Automate"
python -m venv .venv
. .venv/Scripts/activate
pip install -r requirements.txt
```

คัดลอกไฟล์ `.env.example` เป็น `.env` แล้วกรอกค่า:
```
TENANT_ID=<your-tenant-id>
CLIENT_ID=<your-client-id>
CLIENT_SECRET=<your-client-secret>
DATAVERSE_URL=https://orge8e671e1.crm5.dynamics.com
WORKFLOW_ID=1015b2f8-5575-45dd-b1ba-adca4f1f5957
POLL_INTERVAL_SEC=5
POLL_TIMEOUT_SEC=1200
```

## การใช้งาน (2 โหมด)

### โหมด A: Client Credentials (TENANT_ID/CLIENT_ID/CLIENT_SECRET)
รัน Desktop Flow และเฝ้าดูจนเสร็จ:
```bash
python run_desktop_flow.py
```

ถ้า Flow ต้องการ Input parameters (JSON):
```bash
python run_desktop_flow.py --inputs inputs.json
```

หากไม่ได้ใช้ไฟล์อินพุต สามารถกำหนดค่าบังคับผ่าน CLI หรือ .env:
```bash
python run_desktop_flow.py --run-mode Unattended --connection-name shared_uiflow
```

สคริปต์จะ:
1) ขอโทเค็น OAuth2 ผ่าน MSAL (client credentials) ไปยัง `DATAVERSE_URL/.default`
2) เรียก Action: `POST /api/data/v9.2/workflows(<WORKFLOW_ID>)/Microsoft.Dynamics.CRM.RunDesktopFlow`
3) สอบถามตาราง `flowruns` ล่าสุดของ Flow เดียวกัน เพื่อดึงสถานะจนกว่าจะ `Succeeded/Completed` หรือ `Failed/Cancelled` หรือครบกำหนดเวลา (`POLL_TIMEOUT_SEC`)

ผลลัพธ์สุดท้ายจะแสดง:
- Outcome: Succeeded / Failed / TimedOut
- FlowRunId, Status และ ErrorCode/ErrorMessage (หากล้มเหลว)

โค้ดนี้ใช้ตาราง `flowruns` ใน Dataverse เพื่อตรวจสอบสถานะการรัน โดยอิง `createdon` หลังเวลาที่เรา Trigger เพื่อให้ตรงกับรันที่เพิ่งเริ่มต้น

### โหมด B: Token-only (มี Access Token อยู่แล้ว)
ตั้งค่าใน `.env` เพิ่ม `ACCESS_TOKEN` (หรือส่งผ่านพารามิเตอร์ `--token` ที่สคริปต์):
```bash
python run_desktop_flow_token.py
# หรือ
python run_desktop_flow_token.py --token "<access_token>"
# ส่งอินพุต
python run_desktop_flow_token.py --token "<access_token>" --inputs inputs.json
```
หรือกำหนดค่าบังคับผ่าน CLI:
```bash
python run_desktop_flow_token.py --token "<access_token>" --run-mode Unattended --connection-name shared_uiflow
```
หมายเหตุ: Access Token ต้องออกสำหรับ Resource เดียวกับ `DATAVERSE_URL` (audience เท่ากับ URL นั้น) เพื่อเรียก `/api/data/v9.2/...` ได้สำเร็จ

## ข้อควรทราบ
- สิทธิ์: Application user ต้องมีสิทธิ์รัน Desktop Flow ใน Environment นั้น
- การ Throttling: ถ้าถูกจำกัดอัตรา (429) อาจต้องเพิ่ม `POLL_INTERVAL_SEC` หรือลองใหม่
- สถานะที่รองรับ: `Succeeded/Completed`, `Failed/Cancelled` และสถานะระหว่างทาง เช่น `Running/Queued`
- หาก Action ของ Flow ต้องการพารามิเตอร์ ให้เตรียม `inputs.json` ตามสัญญาของ Flow นั้น ๆ

## โครงสร้าง `inputs.json` (ตัวอย่าง)
อย่างต่ำ Action ต้องการ `runMode` และ `connectionName` ถ้าไม่ใส่ใน `.env` หรือ CLI ให้ระบุไว้ในไฟล์:
```json
{
  "runMode": "Unattended",
  "connectionName": "shared_uiflow",
  "inputs": {
    "SampleParameter": "value"
  }
}
```

## การแก้ไขปัญหาเบื้องต้น
- 401/403: ตรวจสอบ Tenant/Client/Secret และสิทธิ์ Application user
- 404: ตรวจสอบ `DATAVERSE_URL` และ `WORKFLOW_ID`
- 5xx: ลองใหม่ภายหลัง, ตรวจสอบ Health ของบริการ
=======
# Python_PowerAutomate
>>>>>>> 51980058e8e809ecbc6b5d583d0d5a5b00d771cb
