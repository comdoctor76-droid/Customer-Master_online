/* Firebase 설정 - 실시간 데이터 동기화
 *
 * Firestore 콜렉션 구조:
 *   students/{empNo} = {
 *     region, center, branch, cohort, empNo, name, phone,
 *     base, target, honors, updatedAt(serverTimestamp)
 *   }
 *
 * Firestore 보안 규칙 (테스트용 - 실제 운영 전 반드시 인증 추가):
 *   rules_version = '2';
 *   service cloud.firestore {
 *     match /databases/{database}/documents {
 *       match /students/{empNo} {
 *         allow read, write: if true;
 *       }
 *     }
 *   }
 */

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { getAnalytics, isSupported as isAnalyticsSupported }
  from "https://www.gstatic.com/firebasejs/10.12.0/firebase-analytics.js";
import {
  getFirestore,
  collection,
  doc,
  setDoc,
  deleteDoc,
  onSnapshot,
  serverTimestamp,
  writeBatch,
  enableIndexedDbPersistence
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

const FIREBASE_CONFIG = {
  apiKey: "AIzaSyCOUmChGShJn7VsicetIWO0yVSBAWI5srY",
  authDomain: "customer-master-online.firebaseapp.com",
  projectId: "customer-master-online",
  storageBucket: "customer-master-online.firebasestorage.app",
  messagingSenderId: "307685872563",
  appId: "1:307685872563:web:d4b59af631584c654271c6",
  measurementId: "G-S3LDN27Y59"
};

const app = initializeApp(FIREBASE_CONFIG);
const db = getFirestore(app);

// 오프라인 캐시 (네트워크 끊겨도 로컬 유지 후 재접속 시 동기화)
enableIndexedDbPersistence(db).catch((err) => {
  if (err.code === "failed-precondition") {
    console.warn("[Firebase] 여러 탭에서 열려있어 오프라인 캐시를 건너뜁니다.");
  } else if (err.code === "unimplemented") {
    console.warn("[Firebase] 브라우저가 IndexedDB 캐시를 지원하지 않습니다.");
  }
});

// Analytics (지원되는 환경에서만)
isAnalyticsSupported().then((ok) => {
  if (ok) {
    try { getAnalytics(app); } catch (e) { console.warn("[Firebase] Analytics 초기화 실패:", e); }
  }
});

// 사번 정규화: Firestore doc ID 에 사용할 수 없는 문자 방지
function normalizeEmpNo(raw) {
  return String(raw || "").trim().replace(/[\/\\\s]/g, "");
}

// 앱에서 사용할 전역 데이터 API
window.DataAPI = {
  configured: true,

  // 실시간 구독
  subscribe(callback) {
    const ref = collection(db, "students");
    return onSnapshot(
      ref,
      (snap) => {
        const list = [];
        snap.forEach((d) => list.push(d.data()));
        callback(list);
        setStatus("online", "실시간 연결");
      },
      (err) => {
        console.error("[Firebase] 구독 오류:", err);
        setStatus("offline", "연결 오류");
      }
    );
  },

  // 사번 기준 저장 (upsert) — 동일 사번이면 merge, 없으면 신규
  async save(student) {
    const empNo = normalizeEmpNo(student.empNo);
    if (!empNo) throw new Error("사번이 비어있습니다.");
    const record = {
      region: student.region || "",
      center: student.center || "",
      branch: student.branch || "",
      cohort: student.cohort || "",
      empNo,
      name: student.name || "",
      phone: student.phone || "",
      base: Number(student.base || 0),
      target: Number(student.target || 0),
      honors: Number(student.honors || 0),
      updatedAt: serverTimestamp()
    };
    if (student.tenureMonths !== undefined && student.tenureMonths !== "") {
      record.tenureMonths = Number(student.tenureMonths) || 0;
    }
    await setDoc(doc(db, "students", empNo), record, { merge: true });
  },

  // 여러건 일괄 저장 — writeBatch 로 한 번의 네트워크 호출에 묶어서 처리
  // Firestore 제한: 배치 1개당 최대 500개 작업
  async saveMany(students) {
    const errors = [];
    const valid = [];
    students.forEach((student, idx) => {
      const empNo = normalizeEmpNo(student.empNo);
      if (!empNo) { errors.push({ idx, empNo: student.empNo, message: "사번 누락" }); return; }
      const record = {
        region: student.region || "",
        center: student.center || "",
        branch: student.branch || "",
        cohort: student.cohort || "",
        empNo,
        name: student.name || "",
        phone: student.phone || "",
        base: Number(student.base || 0),
        target: Number(student.target || 0),
        honors: Number(student.honors || 0),
        updatedAt: serverTimestamp()
      };
      if (student.tenureMonths !== undefined && student.tenureMonths !== "") {
        record.tenureMonths = Number(student.tenureMonths) || 0;
      }
      valid.push({ empNo, record });
    });

    let committed = 0;
    const CHUNK = 450; // 500 한도보다 안전 마진
    for (let i = 0; i < valid.length; i += CHUNK) {
      const slice = valid.slice(i, i + CHUNK);
      const batch = writeBatch(db);
      slice.forEach(({ empNo, record }) => {
        batch.set(doc(db, "students", empNo), record, { merge: true });
      });
      try {
        await batch.commit();
        committed += slice.length;
      } catch (err) {
        slice.forEach(({ empNo }) => {
          errors.push({ empNo, message: err.message || String(err) });
        });
      }
    }
    return { committed, errors };
  },

  // 사번 기준 삭제
  async remove(empNo) {
    const id = normalizeEmpNo(empNo);
    if (!id) return;
    await deleteDoc(doc(db, "students", id));
  }
};

// 온라인/오프라인 상태 배지 제어
function setStatus(cls, text) {
  const badge = document.getElementById("connection-status");
  if (!badge) return;
  badge.classList.remove("online", "offline");
  badge.classList.add(cls);
  badge.textContent = text;
}

document.addEventListener("DOMContentLoaded", () => setStatus("online", "Firebase 연결중..."));
window.addEventListener("online", () => setStatus("online", "실시간 연결"));
window.addEventListener("offline", () => setStatus("offline", "오프라인 (캐시 사용)"));
