// GOLJA-FIREBASE-AND-DEPLOY-20260906-G9: Firebase 프로젝트 연결(golja-app, 2026-09-06).
// 아래 값은 딘이 Firebase 콘솔에서 직접 발급받은 실제 웹앱 config입니다(Firebase
// 웹 config는 원래 클라이언트에 공개되는 값이라 그 자체로는 비밀번호가 아니지만,
// Firestore 보안 규칙이 열려 있으면 누구나 읽고 쓸 수 있으니 주의 — 이 프로젝트는
// "request.auth != null" 규칙을 쓴다(정확한 규칙 텍스트는
// reports/GOLJA_TASK_9_FIREBASE_DEPLOY_REPORT.md 참조, 딘이 콘솔에 직접 붙여넣음).
const FIREBASE_CONFIG = {
  apiKey: "AIzaSyDwvi6M-iXFqlR-4LZTarOc1alrop2L5Cs",
  authDomain: "golja-app.firebaseapp.com",
  projectId: "golja-app",
  storageBucket: "golja-app.firebasestorage.app",
  messagingSenderId: "79603514301",
  appId: "1:79603514301:web:b036b8b999862964fcf0d4"
};
function isFirebaseConfigReady(){
  return !!(FIREBASE_CONFIG.apiKey && FIREBASE_CONFIG.projectId);
}

// index.html에서 이 스크립트를 Firebase SDK 스크립트들 "뒤"에 로드하므로 이
// 시점엔 이미 firebase 전역이 존재한다. db는 골자 앱 인라인 스크립트(같은
// 페이지의 다음 <script> 블록)에서 golja_data 컬렉션 읽기/쓰기에 그대로
// 사용한다(같은 페이지 안의 여러 <script> 태그는 let/const 전역 바인딩을
// 공유하므로 index.html 쪽에서 별도 선언 없이 db를 바로 참조할 수 있다).
let db=null;
if(isFirebaseConfigReady()&&typeof firebase!=='undefined'){
  firebase.initializeApp(FIREBASE_CONFIG);
  db=firebase.firestore();
  console.log('[Firebase] 초기화 완료 (projectId=' + FIREBASE_CONFIG.projectId + ')');
}else if(typeof firebase==='undefined'){
  console.warn('[Firebase] SDK 스크립트가 로드되지 않았습니다(index.html 순서 확인 필요).');
}

// 1인 사용 앱이라 별도 로그인 화면 없이, 앱을 열면 조용히 익명 로그인만
// 자동 실행한다(sokcho-map과 동일 패턴). 실패해도 로컬 데이터로 계속
// 동작해야 하므로 에러는 콘솔 로그만 남기고 앱을 막지 않는다.
if(db&&typeof firebase.auth==='function'){
  firebase.auth().signInAnonymously().then(function(cred){
    console.log('[Firebase] 익명 로그인 성공 (uid=' + (cred&&cred.user&&cred.user.uid) + ')');
  }).catch(function(err){
    console.error('[Firebase] 익명 로그인 실패(로컬 데이터로 계속 동작):',err&&err.code,err&&err.message);
  });
}else if(db){
  console.warn('[Firebase] auth SDK가 로드되지 않아 익명 로그인을 건너뜁니다(index.html에 firebase-auth-compat.js 확인 필요).');
}
