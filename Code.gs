/**
 * 높은뜻푸른교회 교육·세미나 신청 시스템 — 백엔드 (Google Apps Script)
 *
 * 시트 구성 (이 스크립트는 스프레드시트에 "연결된(container-bound)" 스크립트로 사용):
 *   1) members      : 이름 | 전화번호 | 생년월일 | 성별 | 초장 | 직분 | 세례여부   ← 교적 추출 데이터 붙여넣기
 *      (세례여부: Y / 세례 / 입교 → 세례교인 인정. N / 빈칸 / 유아세례만 → 미인정. 교적 값 그대로 사용 가능)
 *   2) courses      : 과정ID | 과정명 | 설명 | 교육기간 | 신청시작일 | 신청마감일 | 정원 | 상태 | 추가질문 | 분반 | 신청자격 | 신청서유형
 *      (신청자격: 비움 또는 '전체' = 누구나 / '교인' = 등록교인만 / '세례교인' = 세례교인만)
 *      (신청서유형: 비움 = 기본 신청서 / '유아세례' = 자녀·배우자 정보를 함께 받는 확장 신청서)
 *   3) applications : 타임스탬프 | 과정ID | 과정명 | 분반 | 이름 | 전화번호 | 생년월일 | 성별 | 초장 | 교적매칭 | 추가질문답변 | 상태
 *                     | 신청자구분 | 자녀이름 | 자녀성별 | 자녀생년월일 | 배우자이름 | 배우자연락처 | 배우자등록 | 배우자세례
 *
 * 배포: 배포 > 새 배포 > 웹 앱
 *   - 실행 사용자: 나
 *   - 액세스 권한: 모든 사용자
 */

const SHEET_MEMBERS = 'members';
const SHEET_COURSES = 'courses';
const SHEET_APPS = 'applications';

// ─────────────────────────────────────────────
// 최초 1회 실행: 관리자 비밀번호 설정
// 아래 ' 여기에비밀번호 ' 부분을 원하는 비밀번호로 바꾼 뒤
// 에디터에서 setAdminPassword 함수를 한 번 실행하세요.
// 실행 후에는 비밀번호를 코드에서 지워도 됩니다.
// ─────────────────────────────────────────────
function setAdminPassword() {
  const password = '여기에비밀번호';
  const hash = sha256Hex_(password);
  PropertiesService.getScriptProperties().setProperty('ADMIN_HASH', hash);
  Logger.log('관리자 비밀번호가 설정되었습니다.');
}

// ─────────────────────────────────────────────
// 진입점
// ─────────────────────────────────────────────
function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: '잘못된 요청 형식입니다.' });
  }

  try {
    switch (req.action) {
      // 신청자용
      case 'getCourses':        return json_(getCourses_());
      case 'verifyMember':      return json_(verifyMember_(req));
      case 'submitApplication': return json_(submitApplication_(req));
      // 관리자용 (비밀번호 필요)
      case 'adminGetCourses':      return adminGuard_(req, () => getCourses_(true));
      case 'adminSaveCourse':      return adminGuard_(req, () => saveCourse_(req.course));
      case 'adminGetApplications': return adminGuard_(req, () => getApplications_(req.courseId));
      case 'adminCancelApplication': return adminGuard_(req, () => cancelApplication_(req.rowId));
      default:
        return json_({ ok: false, error: '알 수 없는 요청입니다.' });
    }
  } catch (err) {
    return json_({ ok: false, error: '처리 중 오류: ' + err.message });
  }
}

function doGet(e) {
  // page 파라미터가 'admin'이면 admin.html을, 그렇지 않으면 index.html을 보여줍니다.
  const page = (e.parameter && e.parameter.page === 'admin') ? 'admin' : 'index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(page === 'admin' ? '관리자 페이지' : '교육훈련 통합 신청')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ─────────────────────────────────────────────
// 신청자 기능
// ─────────────────────────────────────────────

/** 모집중인 과정 목록 (includeAll=true면 관리자용 전체 목록) */
function getCourses_(includeAll) {
  const rows = readSheet_(SHEET_COURSES);
  const today = startOfDay_(new Date());
  const counts = applicationCounts_();

  const courses = rows.map(r => {
    const capacity = Number(r['정원']) || 0;
    const applied = counts[String(r['과정ID'])] || 0;
    return {
      courseId: String(r['과정ID']),
      name: r['과정명'],
      description: r['설명'],
      period: r['교육기간'],
      applyStart: dateStr_(r['신청시작일']),
      applyEnd: dateStr_(r['신청마감일']),
      capacity: capacity,
      applied: applied,
      full: capacity > 0 && applied >= capacity,
      status: r['상태'],
      extraQuestion: r['추가질문'] || '',
      sections: String(r['분반'] || '').split(',').map(s => s.trim()).filter(s => s),
      eligibility: String(r['신청자격'] || '전체').trim() || '전체',
      formType: String(r['신청서유형'] || '').trim()
    };
  });

  if (includeAll) return { ok: true, courses: courses };

  // 신청자에게는: 상태=모집중 + 신청기간 내 과정만 노출
  const open = courses.filter(c => {
    if (c.status !== '모집중') return false;
    const s = c.applyStart ? startOfDay_(new Date(c.applyStart)) : null;
    const e = c.applyEnd ? startOfDay_(new Date(c.applyEnd)) : null;
    if (s && today < s) return false;
    if (e && today > e) return false;
    return true;
  });
  return { ok: true, courses: open };
}

/** 본인 확인: 이름 + 전화번호 뒤 4자리 */
function verifyMember_(req) {
  const name = String(req.name || '').trim();
  const last4 = String(req.last4 || '').trim();
  if (!name || last4.length !== 4) {
    return { ok: false, error: '이름과 전화번호 뒤 4자리를 입력해 주세요.' };
  }

  const cleanInput = name.replace(/\s+/g, '');
  const members = readSheet_(SHEET_MEMBERS);
  const matches = members.filter(m => {
    const phone = digits_(m['전화번호']);
    const cleanName = String(m['이름']).replace(/\s+/g, '');
    return cleanName.startsWith(cleanInput) && phone.slice(-4) === last4;
  });

  if (matches.length === 0) {
    return { ok: true, found: false }; // 교적에 없음 → 프론트에서 수기입력으로 전환
  }
  if (matches.length > 1) {
    // 동명이인 + 뒤 4자리까지 동일한 드문 경우 → 전체 번호로 재확인 요청
    return { ok: true, found: false, needFullPhone: true };
  }
  const m = matches[0];
  return {
    ok: true,
    found: true,
    member: {
      name: String(m['이름']).trim(),
      phone: digits_(m['전화번호']),
      birth: birthStr_(m['생년월일']),
      gender: String(m['성별'] || '').trim(),
      group: m['초장'] || '',
      position: m['직분'] || '',
      baptized: isBaptized_(m['세례여부'])
    }
  };
}

/** 신청 저장 */
function submitApplication_(req) {
  const courseId = String(req.courseId || '').trim();
  const section = String(req.section || '').trim();
  const name = String(req.name || '').trim();
  const phone = digits_(req.phone || '');
  const birth = String(req.birth || '').trim();
  const gender = String(req.gender || '').trim();
  const group = String(req.group || '').trim();
  const isMember = req.isMember === true;
  const extraAnswer = String(req.extraAnswer || '').trim();
  // 유아세례 확장 신청서 항목
  const applicantRole = String(req.applicantRole || '').trim(); // 아빠 / 엄마
  const childName = String(req.childName || '').trim();
  const childGender = String(req.childGender || '').trim();
  const childBirth = String(req.childBirth || '').trim();
  const spouseName = String(req.spouseName || '').trim();
  const spousePhone = digits_(req.spousePhone || '');
  const spouseRegistered = String(req.spouseRegistered || '').trim();
  const spouseBaptized = String(req.spouseBaptized || '').trim();

  if (!courseId || !name || phone.length < 9) {
    return { ok: false, error: '필수 정보(과정, 이름, 전화번호)를 확인해 주세요.' };
  }

  // 동시 신청 경합 방지
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // 과정 유효성 + 정원 확인
    const result = getCourses_(true);
    const course = result.courses.find(c => c.courseId === courseId);
    if (!course) return { ok: false, error: '존재하지 않는 과정입니다.' };
    if (course.status !== '모집중') return { ok: false, error: '현재 모집중인 과정이 아닙니다.' };
    if (course.full) return { ok: false, error: '정원이 마감되었습니다.' };
    if (course.sections.length > 0 && !course.sections.includes(section)) {
      return { ok: false, error: '분반을 선택해 주세요.' };
    }

    // 신청자격 검증 (서버 측에서 교적 재대조)
    if (course.eligibility === '교인' || course.eligibility === '세례교인') {
      const m = findMember_(name, phone, birth);
      if (!m) {
        return { ok: false, error: '이 과정은 등록 교인만 신청하실 수 있습니다. 교적 정보 확인이 필요하시면 교회 사무실로 문의해 주세요.' };
      }
      if (course.eligibility === '세례교인' && !isBaptized_(m['세례여부'])) {
        return { ok: false, error: '이 과정은 세례교인(입교 포함)만 신청하실 수 있습니다. 문의는 교회 사무실로 부탁드립니다.' };
      }
    }

    // 유아세례 확장 신청서 필수 항목 검증
    if (course.formType === '유아세례') {
      if (!applicantRole || !childName || !childGender || !childBirth) {
        return { ok: false, error: '신청자 구분과 자녀 정보(이름·성별·생년월일)를 모두 입력해 주세요.' };
      }
    }

    // 중복 신청 확인 (같은 과정 + 같은 전화번호)
    const apps = readSheet_(SHEET_APPS);
    const dup = apps.some(a =>
      String(a['과정ID']) === courseId &&
      digits_(a['전화번호']) === phone &&
      a['상태'] !== '취소'
    );
    if (dup) return { ok: false, error: '이미 신청하신 과정입니다.' };

    const sheet = ss_().getSheetByName(SHEET_APPS);
    sheet.appendRow([
      new Date(), courseId, course.name, section, name, phone, birth, gender, group,
      isMember ? 'Y' : 'N(교적외)', extraAnswer, '신청',
      applicantRole, childName, childGender, childBirth,
      spouseName, spousePhone, spouseRegistered, spouseBaptized
    ]);

    return { ok: true, message: '신청이 완료되었습니다.' };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────
// 관리자 기능
// ─────────────────────────────────────────────

function adminGuard_(req, fn) {
  const stored = PropertiesService.getScriptProperties().getProperty('ADMIN_HASH');
  if (!stored || sha256Hex_(String(req.password || '')) !== stored) {
    return json_({ ok: false, error: '관리자 인증에 실패했습니다.' });
  }
  return json_(fn());
}

/** 과정 추가/수정 (course.courseId가 기존에 있으면 수정, 없으면 추가) */
function saveCourse_(course) {
  if (!course || !course.courseId || !course.name) {
    return { ok: false, error: '과정ID와 과정명은 필수입니다.' };
  }
  const sheet = ss_().getSheetByName(SHEET_COURSES);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idCol = header.indexOf('과정ID');

  const row = [
    course.courseId, course.name, course.description || '',
    course.period || '', course.applyStart || '', course.applyEnd || '',
    Number(course.capacity) || 0, course.status || '모집중',
    course.extraQuestion || '', course.sections || '', course.eligibility || '전체',
    course.formType || ''
  ];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(course.courseId)) {
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { ok: true, message: '과정이 수정되었습니다.' };
    }
  }
  sheet.appendRow(row);
  return { ok: true, message: '과정이 등록되었습니다.' };
}

/** 특정 과정(또는 전체) 신청자 명단 */
function getApplications_(courseId) {
  const apps = readSheet_(SHEET_APPS).map((a, i) => ({
    rowId: i + 2, // 시트 행 번호 (헤더 다음부터)
    timestamp: dateTimeStr_(a['타임스탬프']),
    courseId: String(a['과정ID']),
    courseName: a['과정명'],
    section: a['분반'] || '',
    name: a['이름'],
    phone: String(a['전화번호']),
    birth: dateStr_(a['생년월일']),
    gender: a['성별'] || '',
    group: a['초장'],
    isMember: a['교적매칭'],
    extraAnswer: a['추가질문답변'],
    status: a['상태'],
    applicantRole: a['신청자구분'] || '',
    childName: a['자녀이름'] || '',
    childGender: a['자녀성별'] || '',
    childBirth: dateStr_(a['자녀생년월일']),
    spouseName: a['배우자이름'] || '',
    spousePhone: String(a['배우자연락처'] || ''),
    spouseRegistered: a['배우자등록'] || '',
    spouseBaptized: a['배우자세례'] || ''
  }));
  const filtered = courseId ? apps.filter(a => a.courseId === String(courseId)) : apps;
  return { ok: true, applications: filtered };
}

/** 신청 취소 처리 (행 삭제 대신 상태 변경) */
function cancelApplication_(rowId) {
  const sheet = ss_().getSheetByName(SHEET_APPS);
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = header.indexOf('상태') + 1;
  if (!rowId || statusCol === 0) return { ok: false, error: '취소 처리에 실패했습니다.' };
  sheet.getRange(Number(rowId), statusCol).setValue('취소');
  return { ok: true, message: '취소 처리되었습니다.' };
}

// ─────────────────────────────────────────────
// 유틸리티
// ─────────────────────────────────────────────

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

/** 시트를 [{헤더:값}] 배열로 읽기 */
function readSheet_(name) {
  const sheet = ss_().getSheetByName(name);
  if (!sheet) throw new Error('시트를 찾을 수 없습니다: ' + name);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const header = data[0].map(h => String(h).trim());
  return data.slice(1)
    .filter(r => r.some(v => v !== ''))
    .map(r => {
      const obj = {};
      header.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
}

/** 과정별 유효 신청 수 */
function applicationCounts_() {
  const apps = readSheet_(SHEET_APPS);
  const counts = {};
  apps.forEach(a => {
    if (a['상태'] === '취소') return;
    const id = String(a['과정ID']);
    counts[id] = (counts[id] || 0) + 1;
  });
  return counts;
}

/** 세례교인 판정: 세례여부 칸 값이 Y, 세례, 입교, 영세면 인정. '유아세례'만 있는 경우는 제외 */
function isBaptized_(v) {
  const s = String(v || '').trim().toUpperCase();
  if (!s || s === 'N') return false;
  if (s === 'Y') return true;
  if (s.indexOf('유아세례') !== -1 && s.indexOf('입교') === -1) return false;
  return s.indexOf('세례') !== -1 || s.indexOf('입교') !== -1 || s.indexOf('영세') !== -1;
}

/** 생년월일을 yyyy-MM-dd로 정규화 (날짜셀, 1980.05.14, 1980/05/14, 19800514, 800514 모두 처리) */
function birthStr_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  const d = digits_(v);
  if (d.length === 8) return d.slice(0, 4) + '-' + d.slice(4, 6) + '-' + d.slice(6, 8);
  if (d.length === 6) {
    // 6자리(YYMMDD)는 26 이하 → 2000년대, 그 외 → 1900년대로 추정
    const century = Number(d.slice(0, 2)) <= 26 ? '20' : '19';
    return century + d.slice(0, 2) + '-' + d.slice(2, 4) + '-' + d.slice(4, 6);
  }
  return String(v);
}

function digits_(v) { return String(v || '').replace(/\D/g, ''); }

/** 교적 재대조: 이름이 같고, (전화번호 뒤4자리 또는 생년월일)이 일치하는 교인 찾기 */
function findMember_(name, phone, birth) {
  const last4 = digits_(phone).slice(-4);
  const cleanInput = String(name || '').replace(/\s+/g, '');
  const members = readSheet_(SHEET_MEMBERS);
  return members.find(m => {
    const cleanName = String(m['이름']).replace(/\s+/g, '');
    if (!cleanName.startsWith(cleanInput)) return false;
    if (last4 && digits_(m['전화번호']).slice(-4) === last4) return true;
    if (birth && birthStr_(m['생년월일']) === birth) return true;
    return false;
  }) || null;
}

function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

function dateStr_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  return String(v);
}

function dateTimeStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  return String(v || '');
}

function sha256Hex_(str) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
