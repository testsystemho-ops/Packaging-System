# Makro Packaging Count System

ระบบบันทึกการตรวจนับ Packing — Fresh Food (CP Axtra / Makro)

แยกออกจากไฟล์ `Packaging-System.html` (ไฟล์เดียว 3,100+ บรรทัด) เป็น 5 ไฟล์
เพื่อให้ดูแล แก้ไข และอัปขึ้น GitHub Pages ได้ง่ายขึ้น โดย**ไม่มีการเปลี่ยนแปลง
ฟังก์ชันการทำงานใดๆ ทั้งสิ้น** — โค้ดเดิมทุกบรรทัดถูกย้ายมาแบบ 1:1

## โครงสร้างไฟล์

```
├── index.html      โครงหน้าเว็บ (login screen + app shell) + ลิงก์ไฟล์ทั้งหมด
├── style.css       ดีไซน์ทั้งหมด (สี ฟอนต์ เลย์เอาต์ responsive)
├── data.json       ข้อมูลตั้งต้น: รายการสินค้า 283 รายการ, รายชื่อสาขา 208 สาขา,
│                   บัญชี Admin, โลโก้ Makro (base64)
├── firebase.js     ตั้งค่าการเชื่อมต่อ Firebase + โหลด data.json
└── app.js          โค้ด logic ทั้งหมด (login, บันทึกตรวจนับ, แดชบอร์ด,
                    หน้าผู้ดูแลระบบ, export Excel ฯลฯ)
```

### ทำไมแยกแบบนี้
- **data.json** แยกข้อมูลดิบ (สินค้า/สาขา) ออกจากโค้ด ทำให้แก้ไข/เพิ่มรายการ
  สินค้าใหม่ได้โดยไม่ต้องแตะโค้ด JavaScript เลย
- **firebase.js** รวม config การเชื่อมต่อฐานข้อมูลไว้จุดเดียว ถ้าต้องเปลี่ยน
  Firebase project ในอนาคต แก้ไฟล์นี้ไฟล์เดียวพอ
- **app.js** คือ logic ทั้งหมดของระบบ (เดิมเป็น 5 ไฟล์ย่อยในตัว bundle —
  รวมเป็นไฟล์เดียวเพื่อให้ง่ายต่อการอัปโหลดและ deploy)
- **style.css** แยกความสวยงามออกจากโครงสร้าง ทำให้ปรับธีม/สีได้ง่าย

## วิธีติดตั้งขึ้น GitHub Pages

1. สร้าง repository ใหม่บน GitHub (หรือใช้ repo เดิม)
2. อัปโหลดไฟล์ทั้ง 5 ไฟล์ (`index.html`, `style.css`, `data.json`,
   `firebase.js`, `app.js`) ไว้ที่ root ของ repo — **ลาก-วาง (drag & drop)
   ผ่านหน้าเว็บ GitHub ได้เลย** ไม่ต้องใช้ command line
3. ไปที่ **Settings → Pages**
4. เลือก Source = `Deploy from a branch`, Branch = `main` (หรือ `master`),
   Folder = `/ (root)` แล้วกด **Save**
5. รอประมาณ 1–2 นาที จะได้ลิงก์ เช่น
   `https://<username>.github.io/<repo-name>/`
6. เข้าลิงก์นั้น ระบบจะพร้อมใช้งานทันที (ไม่ต้องตั้งค่าอะไรเพิ่ม
   เพราะ Firebase config ฝังอยู่ใน `firebase.js` แล้ว)

> **หมายเหตุ:** ทั้ง 5 ไฟล์ต้องอยู่ในโฟลเดอร์เดียวกันเสมอ (เช่นเดียวกับ
> ตอนนี้) เพราะ `index.html` อ้างอิงไฟล์อื่นด้วย path สัมพัทธ์
> (`style.css`, `app.js`, `firebase.js`, `data.json`)

## การแก้ไขที่พบบ่อย

| ต้องการแก้ | แก้ไฟล์ |
|---|---|
| เพิ่ม/แก้ไขรายการสินค้า, ราคา, สาขา | `data.json` |
| เปลี่ยนสี ธีม ฟอนต์ ขนาดตัวอักษร | `style.css` |
| เปลี่ยน Firebase project / database URL | `firebase.js` |
| แก้ logic การคำนวณ, เพิ่มหน้าจอใหม่, แก้ฟอร์ม | `app.js` |
| เปลี่ยนข้อความหน้า login, โครงหน้าเว็บ | `index.html` |

## หมายเหตุทางเทคนิค

- ระบบยังใช้ Firebase Realtime Database (compat SDK) และ SheetJS (สำหรับ
  export Excel) เหมือนเดิมทุกประการ โหลดผ่าน CDN ใน `index.html`
- `firebase.js` จะ `fetch('data.json')` ตอนเปิดหน้าเว็บ แล้วนำข้อมูลไปใส่ใน
  ตัวแปร global เดิม (`ITEMS_DATA`, `STORES_DATA`, `ADMIN_ACCOUNT`,
  `MAKRO_LOGO_DATA_URI`) — `app.js` จะรอข้อมูลนี้โหลดเสร็จก่อนเริ่มทำงาน
  อัตโนมัติ ไม่ต้องทำอะไรเพิ่ม
- เนื่องจากใช้ `fetch()` โหลด `data.json` หากเปิดไฟล์ `index.html` ตรงๆ
  จากเครื่อง (`file://...`) browser บางตัวอาจบล็อกการโหลด — แนะนำให้ทดสอบ
  ผ่าน local server (เช่น `python3 -m http.server`) หรือเปิดผ่าน
  GitHub Pages ซึ่งจะไม่มีปัญหานี้
