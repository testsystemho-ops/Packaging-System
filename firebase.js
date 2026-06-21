/* ============================================================
   firebase.js
   - ตั้งค่าการเชื่อมต่อ Firebase Realtime Database
   - โหลดข้อมูลหลัก (รายการสินค้า / รายชื่อสาขา / บัญชี Admin / โลโก้)
     จากไฟล์ data.json แล้วประกาศเป็นตัวแปร global ตัวเดิม
     (ITEMS_DATA, STORES_DATA, ADMIN_ACCOUNT, MAKRO_LOGO_DATA_URI)
     เพื่อให้ app.js ใช้งานได้เหมือนเดิมทุกประการ โดยไม่ต้องแก้โค้ดส่วนอื่น
   ============================================================ */

// ---------- 1) Firebase Configuration ----------
// TODO: เปลี่ยนค่าด้านล่างนี้หากต้องการย้ายไปใช้ Firebase Project อื่น
const firebaseConfig = {
  apiKey: "AIzaSyDYSCqX52ubME3P3rfO7bcqg0TTfUVvWNc",
  authDomain: "packing-cost.firebaseapp.com",
  databaseURL: "https://packing-cost-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "packing-cost",
  storageBucket: "packing-cost.firebasestorage.app",
  messagingSenderId: "384560968075",
  appId: "1:384560968075:web:f66bee0746bada60a5bf8e"
};

firebase.initializeApp(firebaseConfig);
const db = firebase.database();

// ---------- 2) โหลดข้อมูลหลักจาก data.json ----------
// ตัวแปรเหล่านี้ถูกใช้งานทั่วทั้งแอป (app.js, auth.js, admin-view.js, dashboard.js, utils.js)
let ITEMS_DATA = [];
let STORES_DATA = [];
let ADMIN_ACCOUNT = {};
let MAKRO_LOGO_DATA_URI = '';

// Promise ที่ resolve เมื่อโหลด data.json เสร็จแล้ว — app.js จะรอ Promise นี้ก่อนเริ่มทำงาน
const DATA_READY = fetch('data.json')
  .then(res => {
    if (!res.ok) throw new Error('ไม่สามารถโหลด data.json ได้ (HTTP ' + res.status + ')');
    return res.json();
  })
  .then(json => {
    ITEMS_DATA = json.items || [];
    STORES_DATA = json.stores || [];
    ADMIN_ACCOUNT = json.admin || {};
    MAKRO_LOGO_DATA_URI = json.logo || '';
  })
  .catch(err => {
    console.error('โหลด data.json ไม่สำเร็จ:', err);
    alert('ไม่สามารถโหลดไฟล์ data.json ได้ กรุณาตรวจสอบว่าไฟล์อยู่ในโฟลเดอร์เดียวกับ index.html');
  });
