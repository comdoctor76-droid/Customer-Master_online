# 고객컨설팅 마스터과정 운영관리

현대해상 고객컨설팅 마스터과정 교육생 운영관리 프로그램 (웹 기반)

## 기능 개요

- 지역단 / 비전센터 / 지점 기반 조직 관리 (팝업 선택)
- 교육생 등록 / 수정 / 삭제 (사번 기준 업데이트)
- 단건 입력 + 엑셀 붙여넣기 방식 일괄 등록
- 지점별 교육생 현황 및 KPI 통계
- 기수별 필터, 검색, CSV 내보내기
- Firebase Firestore 기반 실시간 전국 동기화

## 데이터 필드

| 필드 | 설명 |
| --- | --- |
| region | 지역단 |
| center | 비전센터 |
| branch | 지점 |
| cohort | 마스터과정 기수 |
| empNo | 사번 (Primary Key) |
| name | 교육생 이름 |
| phone | 연락처 |
| base | 기준실적 |
| target | 목표실적 |
| honors | 아너스실적 |

## 파일 구조

```
/
├── index.html              # 메인 화면
├── css/style.css           # 스타일 (오렌지 기반 UI)
├── js/app.js               # 앱 로직
├── js/data.js              # 지역단/센터/지점 샘플 데이터
└── js/firebase-config.js   # Firebase 연동 설정
```

## 로컬 실행

별도 빌드 없이 정적 파일로 동작합니다.

```bash
# 간단 서버 실행 (Python 3)
python3 -m http.server 8000
# 브라우저에서 http://localhost:8000 접속
```

또는 VS Code Live Server 확장을 사용해도 됩니다.

## Firebase 연동 (전국 실시간 공유)

1. https://console.firebase.google.com 에서 프로젝트 생성
2. Firestore Database 생성 (프로덕션 또는 테스트 모드)
3. 웹 앱 등록 후 config 값을 복사
4. `js/firebase-config.js` 의 `FIREBASE_CONFIG` 에 붙여넣기
5. Firestore 보안 규칙 설정 예시:

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /students/{empNo} {
      allow read, write: if true;  // 초기 테스트용 - 반드시 인증 추가할 것
    }
  }
}
```

> 실제 운영 시에는 Firebase Authentication 으로 사내 계정 인증을 걸고,
> 규칙에 `request.auth != null` 조건을 추가하는 것을 권장합니다.

## GitHub Pages 배포

1. 이 저장소를 GitHub 에 푸시
2. 저장소 Settings → Pages → Source 를 `main` 브랜치 루트로 설정
3. 배포된 URL (예: `https://<계정>.github.io/Customer-Master_online/`) 을 공유

## 조직 데이터 수정

`js/data.js` 의 `ORG_DATA` 를 실제 현대해상 조직 체계로 교체하세요.
구조: `regions → centers → branches`

## 엑셀 붙여넣기 포맷

다음 10개 컬럼 순서로 복사하여 붙여넣습니다. (TAB 또는 쉼표 구분)

```
지역단 | 비전센터 | 지점 | 기수 | 사번 | 이름 | 연락처 | 기준실적 | 목표실적 | 아너스실적
```

동일 사번은 자동 업데이트, 신규 사번은 신규 등록됩니다.

## 안내

- 로고 및 CI 는 자체 제작한 표식(오렌지 H 마크)을 기본 제공합니다.
  실제 현대해상 공식 로고/CI 는 법무/브랜드팀 승인 후 `index.html` 헤더 영역에 교체 적용하세요.
- 초기 버전은 권한/감사 로그가 없는 단일 공용 콜렉션 구조입니다.
  운영 배포 전 Firebase Auth 및 접근 권한 체계를 반드시 구성하시기 바랍니다.
