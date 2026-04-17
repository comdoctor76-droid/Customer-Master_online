/* Firebase 설정 - 실시간 데이터 동기화
 *
 * 사용 방법:
 * 1) https://console.firebase.google.com 에서 프로젝트 생성
 * 2) 웹 앱 등록 후 config 값을 아래 FIREBASE_CONFIG 에 붙여넣기
 * 3) Firestore Database 생성 후 보안 규칙 설정
 *
 * Firestore 콜렉션 구조:
 *   students/{empNo} = {
 *     region, center, branch, cohort, empNo, name, phone,
 *     base, target, honors, updatedAt
 *   }
 */

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import {
  getFirestore, collection, doc, setDoc, deleteDoc, onSnapshot, serverTimestamp
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

// ⚠ 실제 값으로 교체
const FIREBASE_CONFIG = {
  apiKey: "YOUR_API_KEY",
  authDomain: "YOUR_PROJECT.firebaseapp.com",
  projectId: "YOUR_PROJECT_ID",
  storageBucket: "YOUR_PROJECT.appspot.com",
  messagingSenderId: "YOUR_SENDER_ID",
  appId: "YOUR_APP_ID"
};

const IS_CONFIGURED = FIREBASE_CONFIG.apiKey !== "YOUR_API_KEY";

let db = null;
if (IS_CONFIGURED) {
  const app = initializeApp(FIREBASE_CONFIG);
  db = getFirestore(app);
}

// 앱에서 사용할 전역 데이터 API
window.DataAPI = {
  configured: IS_CONFIGURED,

  // 실시간 구독
  subscribe(callback) {
    if (!IS_CONFIGURED) {
      // 로컬 모드: localStorage 에서 로드
      const load = () => {
        const raw = localStorage.getItem("students");
        callback(raw ? JSON.parse(raw) : []);
      };
      load();
      window.addEventListener("storage", load);
      return () => window.removeEventListener("storage", load);
    }
    const ref = collection(db, "students");
    return onSnapshot(ref, (snap) => {
      const list = [];
      snap.forEach((d) => list.push(d.data()));
      callback(list);
    });
  },

  // 사번 기준 저장 (upsert)
  async save(student) {
    const record = { ...student, updatedAt: new Date().toISOString() };
    if (!IS_CONFIGURED) {
      const raw = localStorage.getItem("students");
      const list = raw ? JSON.parse(raw) : [];
      const idx = list.findIndex((s) => s.empNo === record.empNo);
      if (idx >= 0) list[idx] = { ...list[idx], ...record };
      else list.push(record);
      localStorage.setItem("students", JSON.stringify(list));
      window.dispatchEvent(new Event("storage"));
      return;
    }
    await setDoc(doc(db, "students", record.empNo), {
      ...record,
      updatedAt: serverTimestamp()
    }, { merge: true });
  },

  // 사번 기준 삭제
  async remove(empNo) {
    if (!IS_CONFIGURED) {
      const raw = localStorage.getItem("students");
      const list = raw ? JSON.parse(raw) : [];
      const next = list.filter((s) => s.empNo !== empNo);
      localStorage.setItem("students", JSON.stringify(next));
      window.dispatchEvent(new Event("storage"));
      return;
    }
    await deleteDoc(doc(db, "students", empNo));
  }
};

// 연결 상태 표시
document.addEventListener("DOMContentLoaded", () => {
  const badge = document.getElementById("connection-status");
  if (!badge) return;
  if (IS_CONFIGURED) {
    badge.textContent = "Firebase 연결";
    badge.classList.add("online");
  } else {
    badge.textContent = "로컬 모드";
    badge.classList.add("offline");
  }
});
