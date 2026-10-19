# Pub2Doc Studio & Cloud Service 📄➡️📝

**Pub2Doc** คือแพลตฟอร์มแปลงไฟล์ **Microsoft Publisher (`.pub`)** ให้กลายเป็น **Microsoft Word (`.docx`)** โดยคงความสมบูรณ์ของ **ภาษาไทย สระ-วรรณยุกต์ ตารางคำนวณ รูปภาพความละเอียดสูง และโครงสร้างฟอร์มเอกสาร** ไว้อย่างถูกต้อง 100%

สร้างขึ้นเพื่อเป็นโซลูชันรองรับการยุติการให้บริการอย่างเป็นทางการของ **Microsoft Publisher (End of Life: ตุลาคม 2026)** และแก้ปัญหาข้อผิดพลาดร้ายแรงจากการแปลงผ่าน PDF (สระลอย, วรรณยุกต์หลุด, การเคาะวรรคหลอก `ส ว ั ส ด ี`, และกล่องข้อความแตกกระจาย)

---

## 🌟 ฟีเจอร์เด่น (Key Features)

* **ความแม่นยำภาษาไทยระดับ 100% (Thai Complex Script):**  
  ถอดรหัสข้อความจากหน่วยความจำโดยตรง ไม่ผ่านกระบวนการแปลงเป็น PDF เวกเตอร์ ทำให้สระบน/ล่าง (ิ, ี, ึ, ื, ุ, ู, ั, ำ) และวรรณยุกต์ (่, ้, ๊, ๋, ์) อยู่ในตำแหน่งตรรกะที่ถูกต้อง ไม่ลอย ไม่ซ้อน และไม่กลายเป็นกล่องสี่เหลี่ยม (`□`)
* **รองรับตัวเอียง ตัวหนา และสีตัวอักษรสำหรับภาษาไทย:**  
  เขียนคำสั่ง OpenXML พิเศษ (`<w:iCs/>`, `<w:bCs/>`, `<w:color/>`) ทำให้ฟอนต์ภาษาไทยแสดงผลตัวเอียง ตัวหนา และสี (เช่น ตัวเอียงสีแดงในหมายเหตุ) ได้อย่างถูกต้องใน Microsoft Word
* **ระบบจัดฟอร์มอัจฉริยะ (Smart Form Layout):**  
  ตรวจจับการจัดวางแนวนอนอัตโนมัติ (Spatial Band Clustering) จัดข้อมูลส่วนหัวซ้าย-ขวา (ข้อมูลบริษัท vs โลโก้/เลขผู้เสียภาษี, ข้อมูลลูกค้า vs วันที่ใบแจ้งหนี้) และส่วนลายเซ็นต์ 3 คอลัมน์ลงในตารางโครงร่างไร้ขอบ (Invisible Layout Grid) เอกสารจึงเรียบร้อยและจบพอดีใน 1 หน้า A4
* **แปลงตารางเป็น Word Table แท้:**  
  แปลงตารางคำนวณและรายการสินค้าเป็นตารางจริงของ Word พร้อมจัดสัดส่วนความกว้างคอลัมน์ ชิดขวาสำหรับตัวเลข และชิดซ้ายสำหรับข้อความ
* **ดึงรูปภาพคมชัดระดับต้นฉบับ:**  
  สกัดโลโก้บริษัท ภาพตราประทับ และภาพลายเซ็นต์ตัวจริงออกมาฝังลงในเอกสารโดยตรง
* **สถาปัตยกรรม Dual-Engine (Auto-Detect):**  
  ตรวจจับอัตโนมัติว่าในระบบมี MS Publisher หรือไม่:
  * 🟢 **Native COM Engine:** เมื่อมี MS Publisher บน Windows (ให้คุณภาพสูงสุด 100%)
  * 🟡 **Standalone Fallback Engine:** เมื่อนำไปรันบนระบบที่ไม่มี Publisher (เช่น Linux VPS หรือเครื่องทั่วไป) โดยใช้ LibreOffice Headless + ระบบซ่อมสระไทย PyThaiNLP

---

## 📁 โครงสร้างไฟล์ในโครงการ (Project Structure)

```text
Pub2Doc/
├── app.py                   # โปรแกรมหน้าต่าง Desktop GUI (Tkinter + sv-ttk Modern UI)
├── pub2doc_engine.py        # เอนจินหลักสำหรับประมวลผลแปลงไฟล์ .pub -> .docx
├── form_converter.py        # สคริปต์ตัวอย่างจัดโครงร่างฟอร์มใบแจ้งหนี้ (Invoice)
├── web_app.py               # เว็บแอปพลิเคชัน & REST API (FastAPI + Tailwind Web UI)
├── Run_Pub2Doc.bat          # ไฟล์ Shortcut ดับเบิ้ลคลิกเปิดโปรแกรม Desktop GUI ได้ทันที
├── PRD.md                   # เอกสารวิเคราะห์ความต้องการเชิงลึก (Product Requirements Document)
├── Dockerfile               # Docker configuration สำหรับ Deploy บน Linux VPS
├── docker-compose.yml       # คำสั่งเปิดบริการ Cloud Service แบบบรรทัดเดียว
├── requirements.txt         # รายการไลบรารี Python สำหรับรันระบบ
├── extracted_images/        # โฟลเดอร์เก็บรูปภาพและลายเซ็นต์ที่สกัดออกมา
└── test_images/             # รูปภาพทดสอบ
```

---

## 💻 1. การติดตั้งและใช้งานแบบ Desktop App (Standalone บน Windows)

### ข้อกำหนดของระบบ
* Windows 10 หรือ Windows 11 (64-bit)
* Python 3.10 ขึ้นไป
* *(แนะนำ)* มี Microsoft Office / Publisher ติดตั้งอยู่ในเครื่อง

### ขั้นตอนการติดตั้ง
1. เปิด **PowerShell** หรือ **Command Prompt** เข้าไปยังโฟลเดอร์โครงการ:
   ```powershell
   cd d:\2026\2026_10Oct\Pub2Doc
   ```
2. ติดตั้งไลบรารีที่จำเป็น:
   ```powershell
   python -m pip install -r requirements.txt pywin32 sv-ttk
   ```

### วิธีเปิดใช้งาน
* **วิธีที่ 1 (ง่ายที่สุด):** ดับเบิ้ลคลิกที่ไฟล์ **`Run_Pub2Doc.bat`**
* **วิธีที่ 2:** รันผ่านคำสั่ง Terminal:
  ```powershell
  python app.py
  ```

### วิธีใช้งานหน้าต่างโปรแกรม
1. สังเกตไฟสถานะมุมบนซ้าย:  
   `🟢 ตรวจพบ MS Publisher (v16.0) - ใช้งานโหมดความแม่นยำสูง (Native COM Mode)`
2. กดปุ่ม **`➕ เพิ่มไฟล์ .PUB`** เพื่อเลือกไฟล์เดี่ยว หรือกด **`📁 เพิ่มทั้งโฟลเดอร์`** เพื่อเลือกโฟลเดอร์ที่มีไฟล์ Publisher จำนวนมาก (Batch Convert)
3. เลือกโฟลเดอร์ปลายทางที่ต้องการบันทึกไฟล์ Word
4. กดปุ่ม **`🚀 เริ่มแปลงไฟล์เป็น Word (.docx)`**
5. เมื่อแปลงเสร็จ โปรแกรมจะแจ้งเตือนและมีปุ่มให้กดเปิดโฟลเดอร์ผลลัพธ์ได้ทันที

---

## 🌐 2. การติดตั้งและเปิดให้บริการผ่าน Web Hosting / VPS (Cloud Service)

Pub2Doc มาพร้อมกับระบบ **Web UI แบบลากและวาง (Drag & Drop)** และ **REST API** ในตัว สามารถเปิดให้บริการแก่ผู้ใช้งานทั่วไปหรือพนักงานในองค์กรผ่านเบราว์เซอร์ได้

---

### วิธีที่ 2.1: ทดสอบรันเป็น Web Server ภายในเครื่อง / เครือข่าย LAN

1. สั่งรันเว็บแอปด้วยคำสั่ง:
   ```powershell
   python web_app.py
   ```
2. เปิดเบราว์เซอร์แล้วเข้าไปที่:
   👉 **`http://localhost:8000`**
3. เครื่องอื่นๆ ในวงแลนเดียวกันสามารถเข้าใช้งานผ่าน `http://<IP-เครื่องของคุณ>:8000`

---

### วิธีที่ 2.2: นำขึ้น Linux VPS (Ubuntu) ด้วย Docker (แนะนำสำหรับ Public Service)

วิธีนี้เหมาะสำหรับการเปิดให้บริการผ่านอินเทอร์เน็ตจริง โดยใช้ Linux VPS ราคาประหยัด ($5 - $10 / เดือน) เช่น DigitalOcean, Hetzner, AWS Lightsail หรือ Linode

#### ขั้นที่ 1: เตรียม VPS (Ubuntu 22.04 หรือ 24.04 LTS)
เชื่อมต่อเข้า VPS ผ่าน SSH:
```bash
ssh root@your-vps-ip
```

#### ขั้นที่ 2: ติดตั้ง Docker และ Docker Compose
```bash
sudo apt update && sudo apt install -y docker.io docker-compose
sudo systemctl enable docker
sudo systemctl start docker
```

#### ขั้นที่ 3: อัปโหลดโฟลเดอร์ `Pub2Doc` ขึ้นเซิร์ฟเวอร์
สามารถใช้โปรแกรม FileZilla / WinSCP อัปโหลด หรือใช้คำสั่ง rsync / git:
```bash
# ตัวอย่างการ clone หรือ copy โฟลเดอร์มาที่ /var/www/pub2doc
cd /var/www/pub2doc
```

#### ขั้นที่ 4: สั่งรัน Container
```bash
docker compose up -d --build
```
*ระบบจะดาวน์โหลด Base Image, ติดตั้ง LibreOffice Headless สำหรับถอดรหัส .pub, ติดตั้งชุดฟอนต์ภาษาไทย (TLWG, Sarabun) และเปิดบริการที่พอร์ต 8000*

#### ขั้นที่ 5: ตั้งค่าชื่อโดเมนและ HTTPS (SSL ฟรี) ด้วย Caddy
1. ติดตั้ง Caddy บน VPS:
   ```bash
   sudo apt install -y debian-keyring debian-archive-keyring apt-transport-https curl
   curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/gpg.key' | sudo gpg --dearmor -o /usr/share/keyrings/caddy-stable-archive-keyring.gpg
   curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/debian.deb.txt' | sudo tee /etc/apt/sources.list.d/caddy-stable.list
   sudo apt update && sudo apt install -y caddy
   ```
2. แก้ไขไฟล์ `/etc/caddy/Caddyfile`:
   ```caddy
   pub2doc.yourdomain.com {
       reverse_proxy localhost:8000
   }
   ```
3. รีสตาร์ท Caddy:
   ```bash
   sudo systemctl restart caddy
   ```
*Caddy จะทำการขอใบรับรองความปลอดภัย HTTPS (SSL) จาก Let's Encrypt และต่ออายุให้อัตโนมัติ สามารถเข้าใช้งานผ่าน `https://pub2doc.yourdomain.com` ได้ทันที*

---

## 🔌 3. การเชื่อมต่อผ่าน REST API (Developer Integration)

คุณสามารถนำ Pub2Doc ไปเชื่อมต่อเข้ากับระบบ ERP, เว็บไซต์ หรือ Workflow อัตโนมัติอื่นๆ ผ่าน HTTP API:

### 1. ส่งไฟล์ `.pub` เข้าไปแปลง
* **Endpoint:** `POST /api/convert`
* **Content-Type:** `multipart/form-data`
* **Body:** `file` (ไฟล์นามสกุล `.pub`)

**ตัวอย่างคำสั่งด้วย cURL:**
```bash
curl -X POST "https://pub2doc.yourdomain.com/api/convert" \
     -F "file=@Invoice_Sample.pub"
```

**ตัวอย่างคำสั่งด้วย Python:**
```python
import requests

url = "https://pub2doc.yourdomain.com/api/convert"
with open("Invoice_Sample.pub", "rb") as f:
    response = requests.post(url, files={"file": f})
    data = response.json()
    print("Download URL:", data["download_url"])
```

**ผลลัพธ์ที่ได้ (JSON Response):**
```json
{
  "status": "success",
  "filename": "Invoice_Sample.docx",
  "download_url": "/api/download/Invoice_Sample_a1b2c3d4.docx"
}
```

### 2. ดาวน์โหลดไฟล์ผลลัพธ์
* **Endpoint:** `GET /api/download/{filename}`
* ส่งไฟล์ Word `.docx` กลับมาให้เบราว์เซอร์หรือโปรแกรมดาวน์โหลดทันที

---

## 🛠️ การแก้ไขปัญหาที่พบบ่อย (Troubleshooting & FAQs)

### Q1: ขึ้นแจ้งเตือน "The file is locked because it is currently being used by another program"
* **สาเหตุ:** มีโปรแกรม Microsoft Publisher ค้างอยู่ในหน่วยความจำเบื้องหลัง (Zombie Process)
* **วิธีแก้:** เปิด PowerShell แล้วสั่งคำสั่งเพื่อปิด Process ที่ค้างอยู่:
  ```powershell
  Get-Process MSPUB -ErrorAction SilentlyContinue | Stop-Process -Force
  ```

### Q2: ฟอนต์ภาษาไทยในไฟล์ Word เป็นตัวตรง ไม่ยอมเอียง
* **สาเหตุ:** สำหรับภาษาไทย (Complex Script) การตั้งค่า `run.italic = True` ธรรมดาใน Word จะเอียงเฉพาะภาษาอังกฤษ
* **วิธีแก้:** ตัวเอนจินของ Pub2Doc ได้ใส่คำสั่ง `<w:iCs/>` กำกับไว้ให้เรียบร้อยแล้ว จึงมั่นใจได้ว่าสระและตัวอักษรภาษาไทยจะเอียงตามต้นฉบับอย่างถูกต้อง

### Q3: ต้องการแปลงไฟล์เอกสาร `.pub` ทั้งหมดของบริษัทล่วงหน้า ต้องทำอย่างไร?
* **คำแนะนำ:** แนะนำให้ใช้โปรแกรม **Pub2Doc Desktop GUI** บนเครื่องคอมพิวเตอร์ปัจจุบันที่มี MS Publisher กดปุ่ม **"📁 เพิ่มทั้งโฟลเดอร์"** แล้วกดแปลงรวดเดียวทั้งหมด คุณจะได้คลังไฟล์ `.docx` แท้ 100% เก็บไว้ใช้งานต่อเนื่องได้ตลอดไปก่อนที่ Microsoft จะปิดบริการ Publisher ในเดือนตุลาคม 2026

---

## 📄 ลิขสิทธิ์และการพัฒนา (License)
พัฒนาขึ้นเพื่อการอนุรักษ์และสืบทอดข้อมูลเอกสารขององค์กรและผู้ใช้งาน Microsoft Publisher ในประเทศไทย  
สามารถนำไปปรับแต่ง ติดตั้ง และใช้งานเชิงพาณิชย์ได้อย่างอิสระ
