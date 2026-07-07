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
  initializeFirestore,
  getFirestore,
  persistentLocalCache,
  persistentMultipleTabManager,
  collection,
  collectionGroup,
  doc,
  addDoc,
  setDoc,
  getDoc,
  deleteDoc,
  onSnapshot,
  getDocs,
  query,
  orderBy,
  serverTimestamp,
  writeBatch,
  increment,
  arrayUnion,
  arrayRemove
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
// Firebase 10.x 방식: initializeFirestore 로 영구 캐시를 한 번에 설정
// (구 enableIndexedDbPersistence 는 deprecated — writeBatch 와 충돌하여 hang 유발)
let db;
try {
  db = initializeFirestore(app, {
    localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() })
  });
} catch (e) {
  // 시크릿 모드·구형 브라우저 등 persistentLocalCache 미지원 시 메모리 캐시 폴백
  console.warn("[Firebase] persistentLocalCache 초기화 실패, 기본 설정 사용:", e.message || e.code);
  db = getFirestore(app);
}

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
    let firstFire = true;
    return onSnapshot(
      ref,
      (snap) => {
        const list = [];
        // 시스템 문서(_award_plans_ 등)는 교육생 목록에서 제외
        snap.forEach((d) => { if (!d.id.startsWith("_")) list.push({ _docId: d.id, ...d.data() }); });
        const fromCache = snap.metadata.fromCache;
        const pending  = snap.metadata.hasPendingWrites;
        if (firstFire) {
          console.info(`[Firebase] 첫 데이터 도착: ${list.length}건 (cache=${fromCache}, pending=${pending})`);
          firstFire = false;
        }
        callback(list, { fromCache, pending });
        if (fromCache && list.length === 0) {
          setStatus("offline", "오프라인 (서버 미응답)");
        } else if (fromCache) {
          setStatus("offline", "캐시 표시중");
        } else {
          setStatus("online", "실시간 연결");
        }
      },
      (err) => {
        console.error("[Firebase] 구독 오류:", err);
        setStatus("offline", "연결 오류: " + (err.message || err.code || "unknown"));
        if (typeof window.__onSubscribeError === "function") window.__onSubscribeError(err);
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
    // 실적진도용 확장 필드
    if (student.current !== undefined)      record.current      = Number(student.current)      || 0;
    if (student.pgBase   !== undefined)     record.pgBase       = Number(student.pgBase)       || 0;
    if (student.pgCurrent !== undefined)    record.pgCurrent    = Number(student.pgCurrent)    || 0;
    if (student.ipumCount !== undefined)    record.ipumCount    = Number(student.ipumCount)    || 0;
    if (student.ipumAmt !== undefined)      record.ipumAmt      = Number(student.ipumAmt)      || 0;
    if (student.insAvg !== undefined)       record.insAvg       = Number(student.insAvg)       || 0;
    if (student.curAct !== undefined)       record.curAct       = Number(student.curAct)       || 0;
    // 실적진도현황 붙여넣기 인품 필드 (인품왕 표시에 사용)
    if (student.pgIpumCount !== undefined)  record.pgIpumCount  = Number(student.pgIpumCount)  || 0;
    if (student.pgIpumAmt   !== undefined)  record.pgIpumAmt    = Number(student.pgIpumAmt)    || 0;
    // 실적진도 사전데이터
    if (student.pgPreIns    !== undefined)  record.pgPreIns     = Number(student.pgPreIns)     || 0;
    if (student.pgPreConv   !== undefined)  record.pgPreConv    = Number(student.pgPreConv)    || 0;
    if (student.pgPreIncome !== undefined)  record.pgPreIncome  = Number(student.pgPreIncome)  || 0;
    if (student.pgLeader    !== undefined)  record.pgLeader     = student.pgLeader || "";
    if (student.pgMonth     !== undefined)  record.pgMonth      = student.pgMonth  || "";
    // Step 2 전용 필드
    if (student.pgBase2 !== undefined)      record.pgBase2      = Number(student.pgBase2)      || 0;
    if (student.pgCurrent2 !== undefined)   record.pgCurrent2   = Number(student.pgCurrent2)   || 0;
    if (student.pgIpumCount2 !== undefined) record.pgIpumCount2 = Number(student.pgIpumCount2) || 0;
    if (student.pgIpumAmt2 !== undefined)   record.pgIpumAmt2   = Number(student.pgIpumAmt2)   || 0;
    if (student.hiCap2 !== undefined)       record.hiCap2       = Number(student.hiCap2)       || 0;
    // 팀 배정
    if (student.team !== undefined)         record.team         = String(student.team || "");
    // 직책
    if (student.jobTitle !== undefined)     record.jobTitle     = String(student.jobTitle || "");
    await setDoc(doc(db, "students", empNo), record, { merge: true });
  },

  // 여러건 일괄 저장 — writeBatch 로 한 번의 네트워크 호출에 묶어서 처리
  // Firestore 제한: 배치 1개당 최대 500개 작업
  async saveMany(students, onProgress) {
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
      if (student.current !== undefined)      record.current      = Number(student.current)      || 0;
      if (student.pgBase   !== undefined)     record.pgBase       = Number(student.pgBase)       || 0;
      if (student.pgCurrent !== undefined)    record.pgCurrent    = Number(student.pgCurrent)    || 0;
      if (student.hiCap !== undefined)        record.hiCap        = Number(student.hiCap)        || 0;
      if (student.ipumCount !== undefined)    record.ipumCount    = Number(student.ipumCount)    || 0;
      if (student.ipumAmt !== undefined)      record.ipumAmt      = Number(student.ipumAmt)      || 0;
      if (student.insAvg !== undefined)       record.insAvg       = Number(student.insAvg)       || 0;
      if (student.curAct !== undefined)       record.curAct       = Number(student.curAct)       || 0;
      // 실적진도현황 붙여넣기 인품 필드 (인품왕 표시에 사용)
      if (student.pgIpumCount !== undefined)  record.pgIpumCount  = Number(student.pgIpumCount)  || 0;
      if (student.pgIpumAmt   !== undefined)  record.pgIpumAmt    = Number(student.pgIpumAmt)    || 0;
      // 실적진도 사전데이터
      if (student.pgPreIns    !== undefined)  record.pgPreIns     = Number(student.pgPreIns)     || 0;
      if (student.pgPreConv   !== undefined)  record.pgPreConv    = Number(student.pgPreConv)    || 0;
      if (student.pgPreIncome !== undefined)  record.pgPreIncome  = Number(student.pgPreIncome)  || 0;
      if (student.pgLeader    !== undefined)  record.pgLeader     = student.pgLeader || "";
      if (student.pgMonth     !== undefined)  record.pgMonth      = student.pgMonth  || "";
      // Step 2 전용 필드
      if (student.pgBase2 !== undefined)      record.pgBase2      = Number(student.pgBase2)      || 0;
      if (student.pgCurrent2 !== undefined)   record.pgCurrent2   = Number(student.pgCurrent2)   || 0;
      if (student.pgIpumCount2 !== undefined) record.pgIpumCount2 = Number(student.pgIpumCount2) || 0;
      if (student.pgIpumAmt2 !== undefined)   record.pgIpumAmt2   = Number(student.pgIpumAmt2)   || 0;
      if (student.hiCap2 !== undefined)       record.hiCap2       = Number(student.hiCap2)       || 0;
      if (student.team !== undefined)         record.team         = String(student.team || "");
      valid.push({ empNo, record });
    });

    let committed = 0;
    const CHUNK = 50; // 작은 청크로 진행 표시를 빠르게 갱신
    const TIMEOUT_MS = 10000; // 10초 타임아웃 — 무한 대기 방지 (짧게 해서 빠른 오류 확인)
    const commitWithTimeout = (batch) => Promise.race([
      batch.commit(),
      new Promise((_, rej) => setTimeout(() => rej(new Error("Firestore 응답 시간 초과 (10초)")), TIMEOUT_MS))
    ]);
    console.info(`[saveMany] 시작: ${valid.length}건, 청크=${CHUNK}`);
    for (let i = 0; i < valid.length; i += CHUNK) {
      const slice = valid.slice(i, i + CHUNK);
      const batch = writeBatch(db);
      slice.forEach(({ empNo, record }) => {
        batch.set(doc(db, "students", empNo), record, { merge: true });
      });
      const chunkIdx = Math.floor(i / CHUNK) + 1;
      console.info(`[saveMany] 청크 ${chunkIdx} 커밋 시도 (${i}~${i + slice.length - 1}번)`);
      try {
        await commitWithTimeout(batch);
        committed += slice.length;
        console.info(`[saveMany] 청크 ${chunkIdx} 완료 — 누적 ${committed}건`);
      } catch (err) {
        console.error(`[saveMany] 청크 ${chunkIdx} 오류:`, err.message || err);
        slice.forEach(({ empNo }) => {
          errors.push({ empNo, message: err.message || String(err) });
        });
      }
      if (onProgress) onProgress(committed, valid.length, errors);
    }
    console.info(`[saveMany] 완료: committed=${committed}, errors=${errors.length}`);
    return { committed, errors };
  },

  // 사번 기준 삭제
  async remove(empNo) {
    const id = normalizeEmpNo(empNo);
    if (!id) return;
    await deleteDoc(doc(db, "students", id));
  },

  // 교육생 + 모든 면담기록 원자적 삭제 — docId 직접 지정 (empNo 없는 문서 포함)
  async removeByDocId(docId) {
    if (!docId) return;
    const consultRef = collection(db, "students", docId, "consultations");
    const snap = await getDocs(consultRef);
    const CHUNK = 450;
    const docs = [];
    snap.forEach((d) => docs.push(d.ref));
    for (let i = 0; i < docs.length; i += CHUNK) {
      const batch = writeBatch(db);
      docs.slice(i, i + CHUNK).forEach((ref) => batch.delete(ref));
      await batch.commit();
    }
    await deleteDoc(doc(db, "students", docId));
  },

  // 교육생 + 모든 면담기록 원자적 삭제 (서브컬렉션 포함)
  async removeStudentWithConsultations(empNo) {
    const id = normalizeEmpNo(empNo);
    if (!id) return;
    const consultRef = collection(db, "students", id, "consultations");
    const snap = await getDocs(consultRef);
    // 최대 500 batch 제한 고려 (면담 + 학생 = N+1)
    const CHUNK = 450;
    const docs = [];
    snap.forEach((d) => docs.push(d.ref));
    for (let i = 0; i < docs.length; i += CHUNK) {
      const batch = writeBatch(db);
      docs.slice(i, i + CHUNK).forEach((ref) => batch.delete(ref));
      await batch.commit();
    }
    await deleteDoc(doc(db, "students", id));
  },

  // ========== 면담 기록 ==========
  // 특정 교육생의 면담 기록 실시간 구독
  subscribeConsultations(empNo, callback) {
    const id = normalizeEmpNo(empNo);
    if (!id) { callback([]); return () => {}; }
    const ref = collection(db, "students", id, "consultations");
    const q = query(ref, orderBy("date", "desc"));
    return onSnapshot(
      q,
      (snap) => {
        const list = [];
        snap.forEach((d) => list.push({ id: d.id, ...d.data() }));
        callback(list);
      },
      (err) => { console.error("[Firebase] 면담 구독 오류:", err); callback([]); }
    );
  },

  // 면담 기록 일회성 fetch (출력용 — 여러 교육생 집계 시 사용)
  async getConsultationsOnce(empNo) {
    const id = normalizeEmpNo(empNo);
    if (!id) return [];
    const ref = collection(db, "students", id, "consultations");
    const q = query(ref, orderBy("date", "desc"));
    const snap = await getDocs(q);
    const list = [];
    snap.forEach((d) => list.push({ id: d.id, ...d.data() }));
    return list;
  },

  // 면담 기록 추가 (Phase 1: 확장된 필드 지원)
  async addConsultation(empNo, entry) {
    const id = normalizeEmpNo(empNo);
    if (!id) throw new Error("사번이 비어있습니다.");
    const toNum = (v) => {
      if (v === undefined || v === null || v === "") return 0;
      const n = Number(String(v).replace(/[,\s]/g, ""));
      return Number.isFinite(n) ? n : 0;
    };
    const record = {
      date: entry.date || new Date().toISOString().slice(0, 10),
      content: entry.content || "",
      seq: String(entry.seq || "").trim(),
      ins: toNum(entry.ins),
      tgt: toNum(entry.tgt),
      pct: toNum(entry.pct),
      curAct: toNum(entry.curAct),
      plan: toNum(entry.plan),
      hap: toNum(entry.hap),
      exp: toNum(entry.exp),
      close1: toNum(entry.close1),
      close2: toNum(entry.close2),
      coach: (entry.coach || "").trim(),
      clients: Array.isArray(entry.clients) ? entry.clients.slice(0, 5).map((c) => ({
        name: (c.name || "").trim(),
        types: Array.isArray(c.types) ? c.types : [],
        consult: Array.isArray(c.consult) ? c.consult : [],
        material: Array.isArray(c.material) ? c.material : [],
        amount: Array.isArray(c.amount) ? c.amount : [],
        amountDirect: (c.amountDirect || "").toString(),
        bj: Array.isArray(c.bj) ? c.bj : [],
        memo: (c.memo || "").toString()
      })) : [],
      // Phase 3 시상 계산기 스냅샷
      calcAvg: (entry.calcAvg || "").toString(),
      calcBaseTgt: (entry.calcBaseTgt || "").toString(),
      calcTgt: (entry.calcTgt || "").toString(),
      calcComment: (entry.calcComment || "").toString(),
      createdAt: serverTimestamp()
    };
    await addDoc(collection(db, "students", id, "consultations"), record);
    // 사이드바 면담 횟수 배지용 카운터
    try {
      await setDoc(doc(db, "students", id), { consultCount: increment(1), updatedAt: serverTimestamp() }, { merge: true });
    } catch (e) { console.warn("[Firebase] consultCount +1 실패:", e); }
  },

  // 면담 기록 수정 (Phase 4)
  async updateConsultation(empNo, consultationId, entry) {
    const id = normalizeEmpNo(empNo);
    if (!id || !consultationId) throw new Error("필수 식별자 누락");
    const toNum = (v) => {
      if (v === undefined || v === null || v === "") return 0;
      const n = Number(String(v).replace(/[,\s]/g, ""));
      return Number.isFinite(n) ? n : 0;
    };
    const patch = {
      date: entry.date || new Date().toISOString().slice(0, 10),
      seq: String(entry.seq || "").trim(),
      ins: toNum(entry.ins),
      tgt: toNum(entry.tgt),
      pct: toNum(entry.pct),
      curAct: toNum(entry.curAct),
      plan: toNum(entry.plan),
      hap: toNum(entry.hap),
      exp: toNum(entry.exp),
      close1: toNum(entry.close1),
      close2: toNum(entry.close2),
      coach: (entry.coach || "").trim(),
      clients: Array.isArray(entry.clients) ? entry.clients.slice(0, 5).map((c) => ({
        name: (c.name || "").trim(),
        types: Array.isArray(c.types) ? c.types : [],
        consult: Array.isArray(c.consult) ? c.consult : [],
        material: Array.isArray(c.material) ? c.material : [],
        amount: Array.isArray(c.amount) ? c.amount : [],
        amountDirect: (c.amountDirect || "").toString(),
        bj: Array.isArray(c.bj) ? c.bj : [],
        memo: (c.memo || "").toString()
      })) : [],
      calcAvg: (entry.calcAvg || "").toString(),
      calcBaseTgt: (entry.calcBaseTgt || "").toString(),
      calcTgt: (entry.calcTgt || "").toString(),
      calcComment: (entry.calcComment || "").toString(),
      updatedAt: serverTimestamp()
    };
    await setDoc(doc(db, "students", id, "consultations", consultationId), patch, { merge: true });
  },

  // 면담 기록 삭제
  async removeConsultation(empNo, consultationId) {
    const id = normalizeEmpNo(empNo);
    if (!id || !consultationId) return;
    await deleteDoc(doc(db, "students", id, "consultations", consultationId));
    try {
      await setDoc(doc(db, "students", id), { consultCount: increment(-1), updatedAt: serverTimestamp() }, { merge: true });
    } catch (e) { console.warn("[Firebase] consultCount -1 실패:", e); }
  },

  // ========== 면담 댓글 (결제 승인 연동 예정) ==========
  // 댓글 추가 — consultations/{id}.comments 배열에 arrayUnion
  async addConsultationComment(empNo, consultationId, comment) {
    const id = normalizeEmpNo(empNo);
    if (!id || !consultationId) throw new Error("대상 면담 정보가 없습니다.");
    if (!comment || !comment.text) throw new Error("댓글 내용이 비어있습니다.");
    const safeComment = {
      id: comment.id || ("cm_" + Date.now()),
      role: String(comment.role || ""),
      author: String(comment.author || ""),
      text: String(comment.text || ""),
      createdAt: comment.createdAt || new Date().toISOString()
    };
    await setDoc(
      doc(db, "students", id, "consultations", consultationId),
      { comments: arrayUnion(safeComment) },
      { merge: true }
    );
  },

  // 댓글 삭제 — id 일치 항목 arrayRemove
  // arrayRemove 는 객체 완전 일치가 필요해서 원본을 다시 받아 삭제
  async removeConsultationComment(empNo, consultationId, commentId) {
    const id = normalizeEmpNo(empNo);
    if (!id || !consultationId || !commentId) return;
    const ref = doc(db, "students", id, "consultations", consultationId);
    const snap = await getDocs(query(collection(db, "students", id, "consultations")));
    let target = null;
    snap.forEach((d) => {
      if (d.id !== consultationId) return;
      const list = (d.data().comments || []);
      target = list.find((c) => c.id === commentId) || null;
    });
    if (!target) return;
    await setDoc(ref, { comments: arrayRemove(target) }, { merge: true });
  },

  // 기존 교육생의 consultCount 재계산 (자기치유: 구독 시 최신값으로 동기화)
  async syncConsultCount(empNo, count) {
    const id = normalizeEmpNo(empNo);
    if (!id) return;
    await setDoc(doc(db, "students", id), { consultCount: Number(count) || 0 }, { merge: true });
  },

  // 전체 학생의 면담 횟수를 한번에 수집 (사이드바 배지용 사전계산)
  // 1) 우선 collectionGroup("consultations") 으로 단일 쿼리 시도
  // 2) 실패(규칙/인덱스 미설정) 시 학생별 getDocs 로 폴백 (병렬 8개씩)
  async fetchAllConsultCounts(empNos) {
    const counts = {};
    try {
      const cg = await getDocs(collectionGroup(db, "consultations"));
      cg.forEach((d) => {
        const parent = d.ref.parent.parent; // students/{empNo}
        const id = parent && parent.id;
        if (!id) return;
        counts[id] = (counts[id] || 0) + 1;
      });
      return counts;
    } catch (err) {
      console.warn("[Firebase] collectionGroup 실패 → 폴백:", err && (err.code || err.message));
    }
    // 폴백: 학생별 병렬 조회 (8개씩 청크)
    const ids = Array.isArray(empNos) ? empNos.map(normalizeEmpNo).filter(Boolean) : [];
    const CHUNK = 8;
    for (let i = 0; i < ids.length; i += CHUNK) {
      const slice = ids.slice(i, i + CHUNK);
      await Promise.all(slice.map(async (id) => {
        try {
          const snap = await getDocs(collection(db, "students", id, "consultations"));
          counts[id] = snap.size;
        } catch (e) {
          counts[id] = 0;
        }
      }));
    }
    return counts;
  },

  // 교육생의 인보험 평균 동기화 (최근 면담값)
  async updateStudentInsAvg(empNo, insAvg) {
    const id = normalizeEmpNo(empNo);
    if (!id) return;
    const val = Number(insAvg);
    if (!Number.isFinite(val)) return;
    await setDoc(
      doc(db, "students", id),
      { insAvg: val, updatedAt: serverTimestamp() },
      { merge: true }
    );
  },

  // 오류신고 저장
  async addErrorReport({ title, content, reporterName, reporterContact, imageBase64 }) {
    const ref = collection(db, "errorReports");
    const docRef = await addDoc(ref, {
      title: title || "",
      content: content || "",
      reporterName: reporterName || "",
      reporterContact: reporterContact || "",
      imageBase64: imageBase64 || null,
      resolved: false,
      createdAt: serverTimestamp()
    });
    return docRef.id;
  },

  // 오류신고 목록 실시간 구독 (최신순)
  subscribeErrorReports(callback) {
    const ref = query(collection(db, "errorReports"), orderBy("createdAt", "desc"));
    return onSnapshot(ref, (snap) => {
      const list = [];
      snap.forEach((d) => list.push({ id: d.id, ...d.data() }));
      callback(list);
    }, (err) => {
      console.error("[Firebase] 오류신고 구독 오류:", err);
    });
  },

  // 오류신고 업데이트 (해결 여부 등)
  async updateErrorReport(id, data) {
    await setDoc(doc(db, "errorReports", id), data, { merge: true });
  },

  // 오류신고 삭제
  async deleteErrorReport(id) {
    await deleteDoc(doc(db, "errorReports", id));
  },

  // 시상안 전체 저장 (students/_award_plans_ — __이름__ 형식은 Firestore 예약어로 금지)
  async saveAwardPlans(plansObj) {
    await setDoc(
      doc(db, "students", "_award_plans_"),
      { plans: plansObj, updatedAt: serverTimestamp() }
    );
  },

  // 시상안 전체 불러오기
  async loadAwardPlans() {
    try {
      const snap = await getDoc(doc(db, "students", "_award_plans_"));
      if (snap.exists()) return snap.data().plans || {};
    } catch (e) {
      console.warn("[Firebase] 시상안 불러오기 실패:", e);
    }
    return {};
  },

  // ── 관리자 계정 관리 (students 컬렉션, _adm_ 접두사 문서) ──────
  // students 구독은 _ 시작 문서를 자동 제외하므로 교육생 목록과 충돌 없음
  subscribeAdmins(callback) {
    return onSnapshot(collection(db, "students"), (snap) => {
      const admins = snap.docs
        .filter(d => d.id.startsWith("_adm_"))
        .map(d => ({ docId: d.id, ...d.data() }));
      callback(admins);
    });
  },
  async saveAdmin(admin) {
    const empNo = String(admin.empNo).trim();
    const ref = doc(db, "students", "_adm_" + empNo);
    await setDoc(ref, {
      empNo,
      name:     admin.name     || "",
      phone:    admin.phone    || "",
      region:   admin.region   || "",
      center:   admin.center   || "",
      branch:   admin.branch   || "",
      role:     admin.role     || "기타",
      password: admin.password !== undefined ? admin.password : "0000",
      updatedAt: serverTimestamp(),
    }, { merge: true });
  },
  async removeAdmin(empNo) {
    await deleteDoc(doc(db, "students", "_adm_" + String(empNo).trim()));
  },
  async getAdminByEmpNo(empNo) {
    const snap = await getDoc(doc(db, "students", "_adm_" + String(empNo).trim()));
    return snap.exists() ? { docId: snap.id, ...snap.data() } : null;
  },
  async getStudentByEmpNo(empNo) {
    const id = normalizeEmpNo(empNo);
    if (!id) return null;
    const snap = await getDoc(doc(db, "students", id));
    return snap.exists() ? { docId: snap.id, ...snap.data() } : null;
  },

  // ── 교육생 접속 로그 (students/_slog_* 문서) ─────────────────
  async logStudentAccess(empNo, name, region, cohort, action) {
    const date = new Date().toISOString().slice(0, 10);
    const docId = "_slog_" + String(empNo).trim() + "_" + date.replace(/-/g, "");
    const ref = doc(db, "students", docId);
    const patch = {
      _type: "student_log",
      empNo: String(empNo).trim(), name: name || "",
      region: region || "", cohort: cohort || "",
      date, actions: arrayUnion(action),
      updatedAt: serverTimestamp(),
    };
    if (action === "로그인") patch.loginCount = increment(1);
    await setDoc(ref, patch, { merge: true });
  },
  async fetchStudentLogs() {
    const snap = await getDocs(collection(db, "students"));
    const logs = [];
    snap.forEach(d => {
      if (d.id.startsWith("_slog_")) logs.push({ docId: d.id, ...d.data() });
    });
    return logs.sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  },
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
