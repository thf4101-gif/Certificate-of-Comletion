/**
 * 구글 스프레드시트(Apps Script) API 연결 + 데모 모드
 *  - 연동 URL이 있으면 구글 시트에 기록
 *  - 없으면 이 브라우저(localStorage)에 임시 저장하는 데모 모드
 */
(function () {
  const CFG = window.APP_CONFIG || {};
  const URL_KEY = 'trainingCertificateTracker.apiUrl';
  const DEMO_KEY = 'trainingCertificateTracker.demo.v5';

  const store = {
    get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch (e) {} },
  };
  const session = {
    get(k) { try { return sessionStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { v == null ? sessionStorage.removeItem(k) : sessionStorage.setItem(k, v); } catch (e) {} },
  };

  const p2 = n => String(n).padStart(2, '0');
  const fmt = d => `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())} ${p2(d.getHours())}:${p2(d.getMinutes())}:${p2(d.getSeconds())}`;
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const due = s => { const m = /^\d{4}-(\d{2})-(\d{2})/.exec(s || ''); return m ? `${+m[1]}/${+m[2]}` : (s || ''); };

  /* ───── 구글 시트 ───── */
  async function remote(url, action, params) {
    let res;
    try {
      res = await fetch(url, { method: 'POST', body: JSON.stringify(Object.assign({ action }, params)) });
    } catch (e) {
      throw new Error('구글 스프레드시트에 연결할 수 없습니다. 인터넷 연결과 웹 앱 URL을 확인하세요.');
    }
    let j;
    try { j = await res.json(); }
    catch (e) { throw new Error('구글 시트 응답을 읽을 수 없습니다. 웹 앱 URL과 배포 설정(액세스: 모든 사용자)을 확인하세요.'); }
    if (!j.ok) { const err = new Error(j.error || '알 수 없는 오류'); err.code = j.code; throw err; }
    return j.data;
  }

  /* ───── 데모 (브라우저 저장) ───── */
  let mem = null;
  function dsave(d) { mem = d; store.set(DEMO_KEY, JSON.stringify(d)); }
  function dload() {
    const s = store.get(DEMO_KEY);
    if (s) { try { return (mem = JSON.parse(s)); } catch (e) {} }
    if (mem) return mem;
    const staffList = [
      { no: '1', name: '교장 김철수' },
      { no: '2', name: '교감 이영희' },
      { no: '3', name: '행정실장 박문수' },
      { no: '4', name: '교사 홍길동' },
      { no: '5', name: '교사 신사임당' },
      { no: '6', name: '주무관 강감찬' },
    ];
    const today = new Date();
    const day = n => { const d = new Date(today); d.setDate(d.getDate() + n); return fmt(d).slice(0, 10); };
    const d = {
      pin: '1234',
      students: staffList,
      tasks: [
        { name: '개인정보 보호 연수', due: day(-1) },
        { name: '심폐소생술 및 응급처치 교육', due: day(2) },
        { name: '활동성 결핵검진 (흉부X선 검사)', due: day(14) },
        { name: '4대 폭력 예방교육 (일반)', due: day(3) },
        { name: '4대 폭력 예방교육 (고위직 대면)', due: day(3) },
        { name: '다문화 교육 (교사 전용)', due: day(4) },
        { name: '아동학대 신고의무자 교육', due: day(5) },
        { name: '장애인식 개선 교육', due: day(7) },
        { name: '청렴 교육', due: day(10) },
      ],
      records: [],
    };
    dsave(d);
    return d;
  }
  function pinErr() { const e = new Error('비밀번호가 올바르지 않습니다. (데모 비밀번호: 1234)'); e.code = 'PIN'; return e; }
  const has = (d, no, name, t) => d.records.some(r => r.no === no && r.name === name && r.task === t);

  const demo = {
    getStudentData() { const d = dload(); return { students: d.students, tasks: d.tasks, records: d.records }; },
    submit({ no, name, tasks, files, memos }) {
      const d = dload(), time = fmt(new Date()), added = [], dups = [];
      tasks.forEach(t => {
        if (has(d, no, name, t)) dups.push(t);
        else {
          added.push(t);
          const f = files && files[t] ? files[t] : null;
          let memoText = memos && memos[t] ? memos[t] : '';
          if (f && f.name) {
            memoText = memoText ? `${memoText} / 파일: ${f.name}` : `파일첨부: ${f.name}`;
          }
          d.records.push({ time, no, name, task: t, memo: memoText, file: f });
        }
      });
      dsave(d); return { added, dups, time };
    },
    verifyPin({ pin }) { if (String(pin) !== dload().pin) throw pinErr(); return { sheetName: '', sheetUrl: '' }; },
  };
  const demoAdmin = {
    getDashboard() { const d = dload(); return { students: d.students, tasks: d.tasks, records: d.records, updated: fmt(new Date()).slice(11), sheetName: '', sheetUrl: '' }; },
    toggleRecord({ no, name, task }) {
      const d = dload(), i = d.records.findIndex(r => r.no === no && r.name === name && r.task === task);
      if (i >= 0) d.records.splice(i, 1); else d.records.push({ time: fmt(new Date()), no, name, task, memo: '담당자 확인' });
      dsave(d); return { submitted: i < 0 };
    },
    addTask({ name, due }) {
      const d = dload(); name = String(name || '').trim();
      if (!name) throw new Error('연수명을 입력하세요.');
      if (d.tasks.some(t => t.name === name)) throw new Error('같은 이름의 연수가 이미 있습니다.');
      d.tasks.push({ name, due: due || '' }); dsave(d); return d.tasks;
    },
    deleteTask({ name }) { const d = dload(); d.tasks = d.tasks.filter(t => t.name !== name); dsave(d); return d.tasks; },
    saveStudents({ students }) { const d = dload(); d.students = students.filter(s => s.name); dsave(d); return d.students; },
    changePin({ newPin }) {
      if (!/^\d{4,8}$/.test(String(newPin))) throw new Error('비밀번호는 숫자 4~8자리로 정하세요.');
      const d = dload(); d.pin = String(newPin); dsave(d); return true;
    },
  };

  /* ───── 공개 API ───── */
  window.Api = {
    get url() {
      const v = store.get(URL_KEY);
      if (v === 'demo') return '';
      return v || CFG.apiUrl || '';
    },
    get isDemo() { return !this.url; },
    setUrl(u) { store.set(URL_KEY, u ? u.trim() : 'demo'); },

    async call(action, params = {}) {
      if (this.url) return remote(this.url, action, params);
      await new Promise(r => setTimeout(r, 150));
      if (demo[action]) return demo[action](params);
      if (demoAdmin[action]) { if (String(params.pin) !== dload().pin) throw pinErr(); return demoAdmin[action](params); }
      throw new Error('알 수 없는 요청: ' + action);
    },
    /** 저장하기 전에 URL과 비밀번호가 맞는지 확인 */
    test(url, pin) { return remote(url.trim(), 'verifyPin', { pin }); },
    resetDemo() { store.set(DEMO_KEY, null); mem = null; },
    DEMO_KEY,
  };
  window.U = { esc, fmt, due, session, CFG };
})();
