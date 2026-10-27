/**
 * 교직원 연수 이수증 수합 확인 서버 (새로 작성한 코드)
 * 기존 프런트엔드 API 형식(students/tasks)을 유지합니다.
 * 반드시 새 스프레드시트에 연결하여 setup()을 실행하세요.
 * 파일 업로드/사용자 본인 인증 기능은 없습니다.
 */
const STAFF_SHEETS = {
  roster: '교직원명단', tasks: '연수목록', records: '제출기록'
};
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('스프레드시트의 확장 프로그램 → Apps Script에서 실행하세요.');
  const p = PropertiesService.getScriptProperties();
  p.setProperty('TRAINING_SHEET_ID', ss.getId());
  ensureSheet_(ss, STAFF_SHEETS.roster, ['번호', '이름']);
  ensureSheet_(ss, STAFF_SHEETS.tasks, ['연수명', '마감일', '등록일시']);
  ensureSheet_(ss, STAFF_SHEETS.records, ['제출일시', '번호', '이름', '연수명', '비고']);
  if (!p.getProperty('TRAINING_PIN_HASH')) setPin_('1234');
  console.log('설정 완료: 초기 비밀번호 1234. 연결 후 즉시 변경하세요. 기존 비밀번호가 있으면 유지됩니다.');
}
function ensureSheet_(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    const actual = sh.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    if (actual.join('|') !== headers.join('|')) throw new Error(name + ' 시트의 머리글이 다릅니다. 새 스프레드시트를 사용하세요.');
  }
}
function book_() {
  const id = PropertiesService.getScriptProperties().getProperty('TRAINING_SHEET_ID');
  if (!id) throw new Error('먼저 setup 함수를 실행하세요.');
  return SpreadsheetApp.openById(id);
}
function rows_(ss, name, cols) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('시트가 없습니다. setup 함수를 실행하세요.');
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
function cell_(s) { return /^[=+@-]/.test(String(s)) ? "'" + s : String(s); }
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
  if (until > Date.now()) { const e = new Error('비밀번호 오류가 누적되어 잠겼습니다. 잠시 후 다시 시도하세요.'); e.code = 'PIN'; throw e; }
  if (until) { p.deleteProperty('TRAINING_PIN_FAILS'); p.deleteProperty('TRAINING_PIN_LOCK_UNTIL'); }
  const salt = p.getProperty('TRAINING_PIN_SALT'), target = p.getProperty('TRAINING_PIN_HASH');
  if (!target) throw new Error('먼저 setup 함수를 실행하세요.');
  if (hash_(salt + ':' + String(pin || '')) !== target) {
    const fails = Number(p.getProperty('TRAINING_PIN_FAILS') || 0) + 1;
    p.setProperty('TRAINING_PIN_FAILS', String(fails));
    if (fails >= 10) p.setProperty('TRAINING_PIN_LOCK_UNTIL', String(Date.now() + 600000));
    const e = new Error('관리자 비밀번호가 올바르지 않습니다.'); e.code = 'PIN'; throw e;
  }
  p.deleteProperty('TRAINING_PIN_FAILS');
  p.deleteProperty('TRAINING_PIN_LOCK_UNTIL');
}
function data_(ss) {
  return {
    students: rows_(ss, STAFF_SHEETS.roster, 2).filter(r => r[1]).map(r => ({ no: r[0], name: r[1] })),
    tasks: rows_(ss, STAFF_SHEETS.tasks, 3).filter(r => r[0]).map(r => ({ name: r[0], due: r[1] })),
    records: rows_(ss, STAFF_SHEETS.records, 5).filter(r => r[1] && r[3]).map(r => ({ time: r[0], no: r[1], name: r[2], task: r[3], memo: r[4] }))
  };
}
function staff_(d, no, name) {
  no = text_(no, '번호', 20); name = text_(name, '이름', 80);
  if (!d.students.some(s => s.no === no && s.name === name)) throw new Error('등록된 교직원이 아닙니다.');
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
    const ext = fileData.name && fileData.name.lastIndexOf('.') > 0 ? fileData.name.slice(fileData.name.lastIndexOf('.')) : (contentType.indexOf('pdf') >= 0 ? '.pdf' : '.jpg');
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

function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    if (!e || !e.postData || e.postData.contents.length > 15000000) throw new Error('요청을 확인하세요. (용량 초과)');
    const q = JSON.parse(e.postData.contents);
    if (!q || typeof q !== 'object' || Array.isArray(q)) throw new Error('요청 형식이 올바르지 않습니다.');
    const allowed = ['getStudentData','submit','verifyPin','getDashboard','toggleRecord','addTask','deleteTask','saveStudents','changePin'];
    if (allowed.indexOf(q.action) < 0) throw new Error('알 수 없는 요청입니다.');
    if (!lock.tryLock(20000)) throw new Error('요청이 많습니다. 잠시 후 다시 시도하세요.');
    locked = true;
    const ss = book_();
    if (q.action !== 'getStudentData' && q.action !== 'submit') requirePin_(q.pin);
    let out;
    if (q.action === 'verifyPin') out = { sheetName: ss.getName(), sheetUrl: ss.getUrl() };
    else if (q.action === 'changePin') {
      const pin = String(q.newPin || '');
      if (!/^\d{4,8}$/.test(pin)) throw new Error('비밀번호는 숫자 4~8자리로 정하세요.');
      setPin_(pin); out = true;
    } else if (q.action === 'getStudentData' || q.action === 'getDashboard') {
      out = data_(ss);
      if (q.action === 'getDashboard') Object.assign(out, { updated: stamp_().slice(11), sheetName: ss.getName(), sheetUrl: ss.getUrl() });
      else out.records = out.records.map(r => ({ no: r.no, name: r.name, task: r.task }));
    } else if (q.action === 'saveStudents') {
      if (!Array.isArray(q.students) || q.students.length > 1000) throw new Error('명단은 최대 1000명까지 입력하세요.');
      const seen = new Set();
      const list = q.students.map(s => {
        const no = text_(s.no, '번호', 20), name = text_(s.name, '이름', 80);
        if (!/^\d+$/.test(no)) throw new Error('번호는 숫자로 입력하세요.');
        if (seen.has(no)) throw new Error('중복된 번호가 있습니다: ' + no);
        seen.add(no); return { no, name };
      });
      const sh = ss.getSheetByName(STAFF_SHEETS.roster);
      if (list.length + 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), list.length + 1 - sh.getMaxRows());
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
      if (list.length) sh.getRange(2, 1, list.length, 2).setNumberFormat('@').setValues(list.map(s => [cell_(s.no), cell_(s.name)]));
      out = list;
    } else if (q.action === 'addTask') {
      const name = text_(q.name, '연수명', 120), due = String(q.due || '').trim();
      if (due && !/^\d{4}-\d{2}-\d{2}$/.test(due)) throw new Error('마감일 형식을 확인하세요.');
      if (data_(ss).tasks.some(t => t.name === name)) throw new Error('같은 이름의 연수가 이미 있습니다.');
      ss.getSheetByName(STAFF_SHEETS.tasks).appendRow([cell_(name), cell_(due), stamp_()]);
      out = data_(ss).tasks;
    } else if (q.action === 'deleteTask') {
      const name = text_(q.name, '연수명', 120), sh = ss.getSheetByName(STAFF_SHEETS.tasks);
      const rows = rows_(ss, STAFF_SHEETS.tasks, 3);
      for (let i = rows.length - 1; i >= 0; i--) if (rows[i][0] === name) sh.deleteRow(i + 2);
      out = data_(ss).tasks;
    } else if (q.action === 'submit' || q.action === 'toggleRecord') {
      const d = data_(ss), s = staff_(d, q.no, q.name), sh = ss.getSheetByName(STAFF_SHEETS.records);
      if (q.action === 'submit') {
        if (!Array.isArray(q.tasks) || !q.tasks.length || q.tasks.length > 100) throw new Error('연수를 선택하세요.');
        const tasks = [...new Set(q.tasks.map(t => text_(t, '연수명', 120)))];
        if (tasks.some(t => !d.tasks.some(x => x.name === t))) throw new Error('현재 등록되지 않은 연수가 있습니다. 화면을 새로고침하세요.');
        const added = [], dups = [], time = stamp_();
        const filesMap = (q.files && typeof q.files === 'object') ? q.files : {};
        tasks.forEach(task => {
          if (d.records.some(r => r.no === s.no && r.name === s.name && r.task === task)) dups.push(task);
          else {
            let memo = '';
            if (filesMap[task]) {
              const fileUrl = saveFileToDrive_(filesMap[task], s, task);
              memo = fileUrl ? '이수증: ' + fileUrl : (filesMap[task].name ? '첨부: ' + filesMap[task].name : '이수증 첨부');
            }
            sh.appendRow([time, cell_(s.no), cell_(s.name), cell_(task), cell_(memo)]);
            added.push(task);
          }
        });
        out = { added, dups, time };
      } else {
        const task = text_(q.task, '연수명', 120);
        if (!d.tasks.some(t => t.name === task)) throw new Error('등록된 연수가 아닙니다.');
        const rows = rows_(ss, STAFF_SHEETS.records, 5); let removed = false;
        for (let i = rows.length - 1; i >= 0; i--) {
          if (rows[i][1] === s.no && rows[i][2] === s.name && rows[i][3] === task) { sh.deleteRow(i + 2); removed = true; }
        }
        if (!removed) sh.appendRow([stamp_(), cell_(s.no), cell_(s.name), cell_(task), '담당자 확인']);
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
