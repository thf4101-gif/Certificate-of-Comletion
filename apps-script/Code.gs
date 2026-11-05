/**
 * 교직원 연수 이수증 수합 확인 서버 (연수별 개별 시트 자동 생성 및 이수번호 등록 지원)
 * - 연수별 시트 자동 생성: 성인지교육, 4대폭력예방교육, 감염병 예방 등 각 연수마다 독립 시트 생성
 * - 각 연수 시트 열: [번호, 직위, 이름, 이수번호, 연수 형태, 제출일시, 이수증 첨부]
 * - 이수증 자동 스캔 및 이수번호 기록 연동
 * - 통합 제출기록 시트 및 대시보드 실시간 연동
 */

const STAFF_SHEETS = {
  roster: '교직원명단',
  tasks: '연수목록',
  records: '제출기록'
};

/**
 * 초기 설정 함수 (스프레드시트의 확장 프로그램 → Apps Script에서 최초 1회 실행)
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('스프레드시트의 확장 프로그램 → Apps Script에서 실행하세요.');
  const p = PropertiesService.getScriptProperties();
  p.setProperty('TRAINING_SHEET_ID', ss.getId());

  ensureSheet_(ss, STAFF_SHEETS.roster, ['번호', '직위', '이름']);
  ensureSheet_(ss, STAFF_SHEETS.tasks, ['연수명', '마감일', '등록일시']);
  ensureSheet_(ss, STAFF_SHEETS.records, ['제출일시', '번호', '이름', '연수명', '비고', '이수번호', '연수형태', '이수증링크']);

  if (!p.getProperty('TRAINING_PIN_HASH')) setPin_('1234');

  // 등록된 연수가 있으면 연수별 시트 자동 생성
  syncAllTaskSheets_(ss);

  console.log('설정 완료: 초기 비밀번호 1234. 연결 후 즉시 관리자 페이지에서 변경하세요.');
}

/**
 * 연수별 개별 시트 즉시 생성 및 동기화 (Apps Script 에디터에서 직접 실행용)
 * 웹앱 재배포 전이라도, 상단 함수 목록에서 이 함수를 선택 후 ▶ 실행을 누르면 모든 시트가 즉시 생성됩니다.
 */
function createAllTaskSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || book_();
  syncAllTaskSheets_(ss);
  console.log('✅ 모든 연수별 시트 생성 및 동기화가 성공적으로 완료되었습니다.');
}

/**
 * 기본 시트 존재 여부 확인 및 생성
 */
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setBackground('#f1f3f4').setFontWeight('bold');
  }
  return sh;
}

function book_() {
  const id = PropertiesService.getScriptProperties().getProperty('TRAINING_SHEET_ID');
  if (!id) throw new Error('먼저 setup 함수를 실행하세요.');
  return SpreadsheetApp.openById(id);
}

function rows_(ss, name, cols) {
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  return sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, cols).getDisplayValues() : [];
}

function stamp_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

function text_(v, label, max) {
  if (typeof v !== 'string' && typeof v !== 'number') throw new Error(label + ' 형식이 올바르지 않습니다.');
  const s = String(v).trim();
  if (!s || s.length > max || /[\r\n\x00]/.test(s)) throw new Error(label + '을(를) 확인하세요.');
  return s;
}

function cell_(s) {
  return /^[=+@-]/.test(String(s)) ? "'" + s : String(s || '');
}

function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function setPin_(pin) {
  const p = PropertiesService.getScriptProperties(), salt = Utilities.getUuid();
  p.setProperty('TRAINING_PIN_SALT', salt);
  p.setProperty('TRAINING_PIN_HASH', hash_(salt + ':' + pin));
  p.deleteProperty('TRAINING_PIN_FAILS');
  p.deleteProperty('TRAINING_PIN_LOCK_UNTIL');
}

function requirePin_(pin) {
  const p = PropertiesService.getScriptProperties();
  const until = Number(p.getProperty('TRAINING_PIN_LOCK_UNTIL') || 0);
  if (until > Date.now()) {
    const e = new Error('비밀번호 오류가 누적되어 잠겼습니다. 잠시 후 다시 시도하세요.');
    e.code = 'PIN';
    throw e;
  }
  if (until) {
    p.deleteProperty('TRAINING_PIN_FAILS');
    p.deleteProperty('TRAINING_PIN_LOCK_UNTIL');
  }
  const salt = p.getProperty('TRAINING_PIN_SALT'), target = p.getProperty('TRAINING_PIN_HASH');
  if (!target) throw new Error('먼저 setup 함수를 실행하세요.');
  if (hash_(salt + ':' + String(pin || '')) !== target) {
    const fails = Number(p.getProperty('TRAINING_PIN_FAILS') || 0) + 1;
    p.setProperty('TRAINING_PIN_FAILS', String(fails));
    if (fails >= 10) p.setProperty('TRAINING_PIN_LOCK_UNTIL', String(Date.now() + 600000));
    const e = new Error('관리자 비밀번호가 올바르지 않습니다.');
    e.code = 'PIN';
    throw e;
  }
  p.deleteProperty('TRAINING_PIN_FAILS');
  p.deleteProperty('TRAINING_PIN_LOCK_UNTIL');
}

/**
 * 교직원 직위와 이름 분리 (예: '돌봄전담사 김민선' -> { role: '돌봄전담사', name: '김민선' })
 */
function parseStaffInfo_(rawName) {
  const str = String(rawName || '').trim();
  const knownRoles = [
    '교장', '교감', '원장', '원감', '행정실장',
    '돌봄전담사', '돌봄전담', '돌봄교사', '동승보호자', '통학차량동승자', '통학차량',
    '조리사', '조리원', '영양교사', '영양사',
    '보건교사', '보건', '전문상담교사', '상담교사', '특수교사', '사서교사',
    '수석교사', '부장교사', '교사', '선생님',
    '교무행정사', '교무실무사', '행정실무사', '교육공무직',
    '주무관', '시설관리원', '당직전담원', '청소원', '기사'
  ];
  for (let i = 0; i < knownRoles.length; i++) {
    const r = knownRoles[i];
    if (str.startsWith(r + ' ') || str.startsWith(r + '\t')) {
      const name = str.slice(r.length).trim();
      if (name) return { role: r, name: name };
    }
  }
  const m = /^(.+?)[\s/()]+([가-힣\w]+)[)]?$/.exec(str);
  if (m) {
    for (let i = 0; i < knownRoles.length; i++) {
      if (m[2].indexOf(knownRoles[i]) >= 0) return { role: knownRoles[i], name: m[1].trim() };
    }
  }
  return { role: '교직원', name: str };
}

/**
 * 구글 시트 탭 이름으로 안전하게 변환 (최대 40자, 특수문자 제거)
 */
function sanitizeSheetName_(name) {
  let clean = String(name || '').replace(/[\\\/?:*\[\]]/g, ' ').replace(/\s+/g, ' ').trim();
  if (clean.length > 40) clean = clean.slice(0, 40).trim();
  return clean || '연수';
}

/**
 * 특정 연수에 대한 개별 시트 보장 및 교직원 명단 기본 행 세팅
 * 열 구조: [번호, 직위, 이름, 이수번호, 연수 형태, 제출일시, 이수증 첨부]
 */
function ensureTaskSheet_(ss, taskName, rosterList) {
  const sheetName = sanitizeSheetName_(taskName);
  let sh = ss.getSheetByName(sheetName);
  const headers = ['번호', '직위', '이름', '이수번호', '연수 형태', '제출일시', '이수증 첨부'];
  let isNew = false;
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    isNew = true;
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    const hRange = sh.getRange(1, 1, 1, headers.length);
    hRange.setBackground('#e8f0fe').setFontWeight('bold').setHorizontalAlignment('center');
    try {
      sh.setColumnWidth(1, 60);  // 번호
      sh.setColumnWidth(2, 110); // 직위
      sh.setColumnWidth(3, 110); // 이름
      sh.setColumnWidth(4, 280); // 이수번호
      sh.setColumnWidth(5, 180); // 연수 형태
      sh.setColumnWidth(6, 160); // 제출일시
      sh.setColumnWidth(7, 240); // 이수증 첨부
    } catch (ignore) {}
  }

  // 교직원 명단 기본 행 채우기 (시트에 헤더만 있는 경우)
  if (rosterList && rosterList.length > 0 && sh.getLastRow() <= 1) {
    const rows = rosterList.map(s => {
      const p = parseStaffInfo_(s.name);
      const role = s.role || p.role;
      const name = s.cleanName || p.name;
      return [cell_(s.no), cell_(role), cell_(name), '', '', '', ''];
    });
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    try {
      sh.getRange(2, 1, rows.length, 3).setHorizontalAlignment('center');
      sh.getRange(2, 6, rows.length, 1).setHorizontalAlignment('center');
    } catch (ignore) {}
  }
  return sh;
}

/**
 * 연수별 개별 시트에 제출 기록 쓰기/갱신
 */
function updateTaskSheetRecord_(ss, taskName, s, certNo, certType, time, fileUrl, rosterList) {
  try {
    const sh = ensureTaskSheet_(ss, taskName, rosterList);
    const lastRow = sh.getLastRow();
    const p = parseStaffInfo_(s.name);
    const role = s.role || p.role;
    const cleanName = s.cleanName || p.name;
    let targetRow = -1;

    if (lastRow > 1) {
      const values = sh.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
      for (let i = 0; i < values.length; i++) {
        const noVal = String(values[i][0]).trim();
        const nameVal = String(values[i][2] || values[i][1]).trim();
        if (noVal === String(s.no).trim() || nameVal === cleanName || nameVal === s.name) {
          targetRow = i + 2;
          break;
        }
      }
    }

    const fileCell = fileUrl ? (fileUrl.indexOf('http') === 0 ? '=HYPERLINK("' + fileUrl + '", "이수증 보기")' : fileUrl) : '';
    const certCell = certNo ? cell_(certNo) : '이수 완료';
    const typeCell = cell_(certType || '원격 연수');

    if (targetRow > 0) {
      // 해당 행의 [이수번호, 연수 형태, 제출일시, 이수증 첨부] 갱신
      sh.getRange(targetRow, 4, 1, 4).setValues([[
        certCell,
        typeCell,
        cell_(time),
        fileCell
      ]]);
    } else {
      // 행이 없으면 추가
      sh.appendRow([
        cell_(s.no),
        cell_(role),
        cell_(cleanName),
        certCell,
        typeCell,
        cell_(time),
        fileCell
      ]);
    }
  } catch (err) {
    console.error('연수 시트 업데이트 오류 (' + taskName + '): ' + err);
  }
}

/**
 * 연수별 시트에서 제출 취소 시 행 내용 초기화
 */
function clearTaskSheetRecord_(ss, taskName, s) {
  try {
    const sheetName = sanitizeSheetName_(taskName);
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() <= 1) return;
    const lastRow = sh.getLastRow();
    const values = sh.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
    const p = parseStaffInfo_(s.name);
    const cleanName = s.cleanName || p.name;
    for (let i = 0; i < values.length; i++) {
      const noVal = String(values[i][0]).trim();
      const nameVal = String(values[i][2] || values[i][1]).trim();
      if (noVal === String(s.no).trim() || nameVal === cleanName || nameVal === s.name) {
        sh.getRange(i + 2, 4, 1, 4).clearContent();
        break;
      }
    }
  } catch (err) {
    console.error('연수 시트 클리어 오류 (' + taskName + '): ' + err);
  }
}

/**
 * 모든 연수별 시트 생성 및 동기화
 */
function syncAllTaskSheets_(ss) {
  const d = data_(ss);
  const tasks = d.tasks;
  const roster = d.students;
  const records = d.records;

  tasks.forEach(t => {
    const sh = ensureTaskSheet_(ss, t.name, roster);
    const taskRecords = records.filter(r => r.task === t.name);
    taskRecords.forEach(r => {
      const st = { no: r.no, name: r.name };
      let certNo = r.certNo || '';
      let certType = r.certType || '원격 연수';
      let fileUrl = r.fileUrl || '';
      if (!certNo && r.memo) {
        const mCert = /이수번호:\s*([^/|\]\n]+)/.exec(r.memo);
        if (mCert) certNo = mCert[1].trim();
      }
      if (!certType && r.memo) {
        const mType = /연수형태:\s*([^/|\]\n]+)/.exec(r.memo);
        if (mType) certType = mType[1].trim();
      }
      if (!fileUrl && r.memo) {
        const mFile = /이수증:\s*(https?:\/\/[^\s]+)/.exec(r.memo);
        if (mFile) fileUrl = mFile[1].trim();
      }
      updateTaskSheetRecord_(ss, t.name, st, certNo, certType, r.time, fileUrl, roster);
    });
  });
}

/**
 * 전체 데이터 읽기
 */
function data_(ss) {
  const rosterRaw = rows_(ss, STAFF_SHEETS.roster, 3);
  const students = rosterRaw.filter(r => r[1] || r[2]).map(r => {
    // 3열인 경우 [번호, 직위, 이름] 또는 2열 [번호, 이름]
    if (r[2]) {
      return { no: r[0], role: r[1], cleanName: r[2], name: r[1] ? (r[1] + ' ' + r[2]) : r[2] };
    }
    const p = parseStaffInfo_(r[1]);
    return { no: r[0], role: p.role, cleanName: p.name, name: r[1] };
  });

  const tasks = rows_(ss, STAFF_SHEETS.tasks, 3).filter(r => r[0]).map(r => ({ name: r[0], due: r[1] }));
  const recordsRaw = rows_(ss, STAFF_SHEETS.records, 8);
  const records = recordsRaw.filter(r => r[1] && r[3]).map(r => ({
    time: r[0],
    no: r[1],
    name: r[2],
    task: r[3],
    memo: r[4] || '',
    certNo: r[5] || '',
    certType: r[6] || '',
    fileUrl: r[7] || ''
  }));

  return { students, tasks, records };
}

function staff_(d, no, name) {
  no = text_(no, '번호', 20);
  name = text_(name, '이름', 80);
  if (!d.students.some(s => s.no === no && (s.name === name || s.cleanName === name))) {
    throw new Error('등록된 교직원이 아닙니다.');
  }
  return { no, name };
}

function doGet() {
  return json_({ ok: true, data: { service: '교직원 연수 이수증 수합', message: '화면에서 제출 여부를 등록하세요.' } });
}

function getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function saveFileToDrive_(fileData, s, task) {
  if (!fileData || !fileData.dataUrl) return '';
  try {
    const parts = fileData.dataUrl.split(',');
    if (parts.length < 2) return '';
    const base64 = parts[1];
    const contentType = (fileData.type || 'image/jpeg');
    const ext = fileData.name && fileData.name.lastIndexOf('.') > 0
      ? fileData.name.slice(fileData.name.lastIndexOf('.'))
      : (contentType.indexOf('pdf') >= 0 ? '.pdf' : '.jpg');
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), contentType, `${s.no}번_${s.name}_${task}_이수증${ext}`);
    const folder = getOrCreateFolder_('교직원연수_이수증첨부');
    const driveFile = folder.createFile(blob);
    try {
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (ignore) {}
    return driveFile.getUrl();
  } catch (err) {
    return '파일저장오류: ' + String(err.message || err);
  }
}

/**
 * 구글 드라이브(DriveApp) 접근 권한 1회 승인 및 시트 동기화 함수
 * Apps Script 상단 함수 선택창에서 'authorizeDriveAndSetup'을 고르고 ▶ 실행을 누르면 드라이브 권한 승인 창이 뜹니다.
 */
function authorizeDriveAndSetup() {
  const folder = getOrCreateFolder_('교직원연수_이수증첨부');
  Logger.log('구글 드라이브 폴더 생성 및 권한 승인 성공: ' + folder.getName());
  createAllTaskSheets();
  return '구글 드라이브 접근 권한 승인 및 모든 연수별 시트 생성이 완료되었습니다.';
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    if (!e || !e.postData || e.postData.contents.length > 15000000) throw new Error('요청을 확인하세요. (용량 초과)');
    const q = JSON.parse(e.postData.contents);
    if (!q || typeof q !== 'object' || Array.isArray(q)) throw new Error('요청 형식이 올바르지 않습니다.');
    const allowed = ['getStudentData','submit','verifyPin','getDashboard','toggleRecord','addTask','deleteTask','saveStudents','changePin','syncTaskSheets'];
    if (allowed.indexOf(q.action) < 0) throw new Error('알 수 없는 요청입니다.');
    if (!lock.tryLock(20000)) throw new Error('요청이 많습니다. 잠시 후 다시 시도하세요.');
    locked = true;
    const ss = book_();
    if (q.action !== 'getStudentData' && q.action !== 'submit') requirePin_(q.pin);
    let out;

    if (q.action === 'verifyPin') {
      out = { sheetName: ss.getName(), sheetUrl: ss.getUrl() };
    } else if (q.action === 'changePin') {
      const pin = String(q.newPin || '');
      if (!/^\d{4,8}$/.test(pin)) throw new Error('비밀번호는 숫자 4~8자리로 정하세요.');
      setPin_(pin);
      out = true;
    } else if (q.action === 'syncTaskSheets') {
      syncAllTaskSheets_(ss);
      out = { success: true, count: data_(ss).tasks.length };
    } else if (q.action === 'getStudentData' || q.action === 'getDashboard') {
      out = data_(ss);
      if (q.action === 'getDashboard') {
        Object.assign(out, { updated: stamp_().slice(11), sheetName: ss.getName(), sheetUrl: ss.getUrl() });
      } else {
        out.records = out.records.map(r => ({ no: r.no, name: r.name, task: r.task, certNo: r.certNo, certType: r.certType }));
      }
    } else if (q.action === 'saveStudents') {
      if (!Array.isArray(q.students) || q.students.length > 1000) throw new Error('명단은 최대 1000명까지 입력하세요.');
      const seen = new Set();
      const list = q.students.map(s => {
        const no = text_(s.no, '번호', 20), name = text_(s.name, '이름', 80);
        if (!/^\d+$/.test(no)) throw new Error('번호는 숫자로 입력하세요.');
        if (seen.has(no)) throw new Error('중복된 번호가 있습니다: ' + no);
        seen.add(no);
        const p = parseStaffInfo_(name);
        return { no, role: s.role || p.role, name: p.name, fullName: name };
      });
      const sh = ss.getSheetByName(STAFF_SHEETS.roster);
      if (list.length + 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), list.length + 1 - sh.getMaxRows());
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
      if (list.length) {
        sh.getRange(2, 1, list.length, 3).setNumberFormat('@').setValues(list.map(s => [cell_(s.no), cell_(s.role), cell_(s.name)]));
      }
      // 교직원 명단 변경 시 모든 연수별 시트의 명단 행도 동기화
      syncAllTaskSheets_(ss);
      out = list.map(s => ({ no: s.no, name: s.fullName }));
    } else if (q.action === 'addTask') {
      const name = text_(q.name, '연수명', 120), due = String(q.due || '').trim();
      if (due && !/^\d{4}-\d{2}-\d{2}$/.test(due)) throw new Error('마감일 형식을 확인하세요.');
      const d = data_(ss);
      if (d.tasks.some(t => t.name === name)) throw new Error('같은 이름의 연수가 이미 있습니다.');
      ss.getSheetByName(STAFF_SHEETS.tasks).appendRow([cell_(name), cell_(due), stamp_()]);
      // 연수 추가 시 개별 시트 즉시 생성
      ensureTaskSheet_(ss, name, d.students);
      out = data_(ss).tasks;
    } else if (q.action === 'deleteTask') {
      const name = text_(q.name, '연수명', 120), sh = ss.getSheetByName(STAFF_SHEETS.tasks);
      const rows = rows_(ss, STAFF_SHEETS.tasks, 3);
      for (let i = rows.length - 1; i >= 0; i--) {
        if (rows[i][0] === name) sh.deleteRow(i + 2);
      }
      out = data_(ss).tasks;
    } else if (q.action === 'submit') {
      const d = data_(ss), s = staff_(d, q.no, q.name), sh = ss.getSheetByName(STAFF_SHEETS.records);
      if (!Array.isArray(q.tasks) || !q.tasks.length || q.tasks.length > 100) throw new Error('연수를 선택하세요.');
      const tasks = [...new Set(q.tasks.map(t => text_(t, '연수명', 120)))];
      if (tasks.some(t => !d.tasks.some(x => x.name === t))) throw new Error('현재 등록되지 않은 연수가 있습니다. 화면을 새로고침하세요.');

      const filesMap = (q.files && typeof q.files === 'object') ? q.files : {};
      const memosMap = (q.memos && typeof q.memos === 'object') ? q.memos : {};
      const certNosMap = (q.certNos && typeof q.certNos === 'object') ? q.certNos : {};
      const certTypesMap = (q.certTypes && typeof q.certTypes === 'object') ? q.certTypes : {};

      const dups = [];
      const added = [];
      const time = stamp_();

      tasks.forEach(task => {
        let certNo = certNosMap[task] ? String(certNosMap[task]).trim() : '';
        let certType = certTypesMap[task] ? String(certTypesMap[task]).trim() : '원격 연수';
        let memo = memosMap[task] ? String(memosMap[task]).slice(0, 300) : '';
        let fileUrl = '';

        if (filesMap[task]) {
          fileUrl = saveFileToDrive_(filesMap[task], s, task);
          const fileText = fileUrl ? '이수증: ' + fileUrl : (filesMap[task].name ? '첨부: ' + filesMap[task].name : '이수증 첨부');
          memo = memo ? (memo + ' / ' + fileText) : fileText;
        }

        const existingRecord = d.records.find(r => r.no === s.no && (r.name === s.name || r.name.indexOf(s.name) >= 0) && r.task === task);

        if (existingRecord) {
          // 이미 제출된 연수인 경우: 새로 첨부된 이수증 파일이나 이수번호가 있으면 기존 행 갱신!
          const hasNewData = !!(filesMap[task] || certNo || memo || (certType && certType !== existingRecord.certType));
          if (!hasNewData) {
            dups.push(task);
            return;
          }

          if (!certNo && existingRecord.certNo && existingRecord.certNo !== '이수 완료') certNo = existingRecord.certNo;
          if (!certType && existingRecord.certType) certType = existingRecord.certType;
          if (!fileUrl && existingRecord.fileUrl) fileUrl = existingRecord.fileUrl;

          const fullMemoParts = [];
          if (certNo) fullMemoParts.push('이수번호: ' + certNo);
          if (certType) fullMemoParts.push('연수형태: ' + certType);
          if (memo) fullMemoParts.push(memo);
          else if (existingRecord.memo) fullMemoParts.push(existingRecord.memo);
          const fullMemo = fullMemoParts.join(' / ');

          const fileLinkCell = fileUrl ? (fileUrl.indexOf('http') === 0 ? '=HYPERLINK("' + fileUrl + '", "이수증 보기")' : fileUrl) : '';
          const certCell = certNo ? cell_(certNo) : '이수 완료';

          // 1. 제출기록 시트에서 기존 행 찾아서 업데이트
          const lastRow = sh.getLastRow();
          let updatedRow = false;
          if (lastRow > 1) {
            const recRows = sh.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
            for (let i = 0; i < recRows.length; i++) {
              if (String(recRows[i][1]).trim() === String(s.no).trim() &&
                  (recRows[i][2] === s.name || recRows[i][2].indexOf(s.name) >= 0) &&
                  recRows[i][3] === task) {
                sh.getRange(i + 2, 1, 1, 8).setValues([[
                  time,
                  cell_(s.no),
                  cell_(s.name),
                  cell_(task),
                  cell_(fullMemo),
                  certCell,
                  cell_(certType),
                  fileLinkCell
                ]]);
                updatedRow = true;
                break;
              }
            }
          }
          if (!updatedRow) {
            sh.appendRow([
              time,
              cell_(s.no),
              cell_(s.name),
              cell_(task),
              cell_(fullMemo),
              certCell,
              cell_(certType),
              fileLinkCell
            ]);
          }

          // 2. 해당 연수 전용 시트(성인지교육, 4대폭력예방교육 등)의 이수번호, 연수형태, 이수증 링크 갱신
          updateTaskSheetRecord_(ss, task, s, certNo, certType, time, fileUrl, d.students);
          added.push(task);
        } else {
          // 신규 제출
          const fullMemoParts = [];
          if (certNo) fullMemoParts.push('이수번호: ' + certNo);
          if (certType) fullMemoParts.push('연수형태: ' + certType);
          if (memo) fullMemoParts.push(memo);
          const fullMemo = fullMemoParts.join(' / ');
          const fileLinkCell = fileUrl ? (fileUrl.indexOf('http') === 0 ? '=HYPERLINK("' + fileUrl + '", "이수증 보기")' : fileUrl) : '';
          const certCell = certNo ? cell_(certNo) : '이수 완료';

          // 1. 전체 통합 기록 시트(제출기록)에 추가
          sh.appendRow([
            time,
            cell_(s.no),
            cell_(s.name),
            cell_(task),
            cell_(fullMemo),
            certCell,
            cell_(certType),
            fileLinkCell
          ]);

          // 2. 해당 연수 전용 시트(성인지교육, 4대폭력예방교육 등)에 갱신/추가
          updateTaskSheetRecord_(ss, task, s, certNo, certType, time, fileUrl, d.students);

          added.push(task);
        }
      });
      out = { added, dups, time };
    } else if (q.action === 'toggleRecord') {
      const d = data_(ss), s = staff_(d, q.no, q.name), sh = ss.getSheetByName(STAFF_SHEETS.records);
      const task = text_(q.task, '연수명', 120);
      if (!d.tasks.some(t => t.name === task)) throw new Error('등록된 연수가 아닙니다.');
      const rows = rows_(ss, STAFF_SHEETS.records, 5);
      let removed = false;
      for (let i = rows.length - 1; i >= 0; i--) {
        if (rows[i][1] === s.no && (rows[i][2] === s.name || rows[i][2].indexOf(s.name) >= 0) && rows[i][3] === task) {
          sh.deleteRow(i + 2);
          removed = true;
        }
      }
      if (!removed) {
        const time = stamp_();
        sh.appendRow([time, cell_(s.no), cell_(s.name), cell_(task), '담당자 확인', '', '대면/직접확인', '']);
        updateTaskSheetRecord_(ss, task, s, '담당자 직접 확인', '대면/직접확인', time, '', d.students);
      } else {
        clearTaskSheetRecord_(ss, task, s);
      }
      out = { submitted: !removed };
    }

    return json_({ ok: true, data: out });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err), code: err.code || 'ERROR' });
  } finally {
    if (locked) lock.releaseLock();
  }
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
