/**
 * 이수증 자동 스캔 및 이수번호/연수형태 자동 인식 모듈
 * - 클라이언트(브라우저)에서 안전하게 실행 (서버 비용 0원, GitHub Pages 완벽 호환)
 * - Tesseract OCR 엔진 연동 및 한국 교육청/학교/사이버연수원 이수증 특화 번호 추출기
 */
(function() {
  let tesseractLoaded = false;
  let loadingPromise = null;

  // Tesseract 라이브러리 동적 로드
  function loadTesseract() {
    if (window.Tesseract) return Promise.resolve(window.Tesseract);
    if (loadingPromise) return loadingPromise;
    loadingPromise = new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = 'https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js';
      script.async = true;
      script.onload = () => {
        tesseractLoaded = true;
        resolve(window.Tesseract);
      };
      script.onerror = () => reject(new Error('OCR 엔진(Tesseract)을 불러오지 못했습니다.'));
      document.head.appendChild(script);
    });
    return loadingPromise;
  }

  // 이미지 전처리 (흑백 대비 강화하여 OCR 인식률 대폭 향상)
  function preprocessImage(imageSrc) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload = () => {
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        const maxDim = 1600;
        let w = img.width, h = img.height;
        if (w > maxDim || h > maxDim) {
          if (w > h) { h = Math.round((h * maxDim) / w); w = maxDim; }
          else { w = Math.round((w * maxDim) / h); h = maxDim; }
        }
        canvas.width = w;
        canvas.height = h;
        ctx.drawImage(img, 0, 0, w, h);

        const imgData = ctx.getImageData(0, 0, w, h);
        const d = imgData.data;
        // 그레이스케일 및 콘트라스트 강화
        const contrast = 1.25; // 25% 대비 증가
        const factor = (259 * (contrast + 255)) / (255 * (259 - contrast));
        for (let i = 0; i < d.length; i += 4) {
          const gray = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
          const c = factor * (gray - 128) + 128;
          const val = Math.min(255, Math.max(0, c));
          d[i] = val;
          d[i + 1] = val;
          d[i + 2] = val;
        }
        ctx.putImageData(imgData, 0, 0);
        resolve(canvas);
      };
      img.onerror = () => reject(new Error('이미지를 처리할 수 없습니다.'));
      img.src = imageSrc;
    });
  }

  /**
   * 한국 이수증 텍스트에서 이수번호 및 연수 형태 스마트 추출
   */
  function extractCertificateInfo(rawText) {
    if (!rawText || typeof rawText !== 'string') {
      return { certNo: '', certType: '원격 연수', rawText: '' };
    }

    const text = rawText.replace(/\r/g, '\n');
    const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

    let certNo = '';
    let certType = '원격 연수';

    // 1. 연수 형태 추정
    if (/내부교직원.*1시간|내부.*대면.*1시간/.test(text)) {
      certType = '내부교직원 활용 대면교육 1시간';
    } else if (/외부강사.*2시간|외부.*대면.*2시간/.test(text)) {
      certType = '외부강사 활용 대면교육 2시간';
    } else if (/내부교직원/.test(text)) {
      certType = '내부교직원 활용 대면교육 1시간';
    } else if (/외부강사/.test(text)) {
      certType = '외부강사 활용 대면교육 2시간';
    } else if (/대면교육|집합교육|대면\s*연수|출석교육/.test(text)) {
      certType = '대면교육';
    } else if (/원격|사이버|온라인|e-러닝|인터넷/.test(text)) {
      certType = '원격 연수';
    }

    // 2. 이수번호 후보군 탐색 및 우선순위 적용
    const candidates = [];

    // 패턴 A: NO. 공무원 성인지 사이버-2026-0988935 형태 (자주 쓰이는 중앙교육연수원/공무원 사이버 연수 형태)
    const mNo = /NO[.:·\s]*([A-Za-z0-9가-힣\s\-_.()/]{4,40})/gi;
    let match;
    while ((match = mNo.exec(text)) !== null) {
      const val = ('NO. ' + match[1].trim().replace(/^NO[.:·\s]*/i, '')).trim();
      if (/\d/.test(val) && val.length >= 6) {
        candidates.push({ val, priority: 100 });
      }
    }

    // 패턴 B: 수비초-7571(2026. 9. 2.) 형태 (학교 공문서 등록번호 형태)
    const mSchoolDoc = /([가-힣]{2,}초?[\s\-_.·]+(?:\d{2,6}|제?\d+호?)(?:\s*\(\s*\d{4}[.\s]+\d{1,2}[.\s]+\d{1,2}[.\s]*\))?)/g;
    while ((match = mSchoolDoc.exec(text)) !== null) {
      const val = match[1].trim();
      if (/\d/.test(val)) {
        candidates.push({ val, priority: 95 });
      }
    }

    // 패턴 C: "제 2026-1092021호" 또는 "제 2026-가-1234호" 형태
    const mJe = /제\s*([0-9A-Za-z가-힣\s\-_.()]{3,35})\s*호/g;
    while ((match = mJe.exec(text)) !== null) {
      const val = ('제 ' + match[1].trim() + '호').trim();
      if (/\d/.test(val)) {
        candidates.push({ val, priority: 90 });
      }
    }

    // 패턴 D: "이수번호 : 2026-0988935" 또는 "등록번호 : ..."
    const mLabel = /(?:이수\s*번호|등록\s*번호|발급\s*번호|증서\s*번호|연수\s*번호)[\s:：]+([A-Za-z0-9가-힣\s\-_.()/]{3,40})/gi;
    while ((match = mLabel.exec(text)) !== null) {
      const val = match[1].trim();
      if (/\d/.test(val)) {
        candidates.push({ val, priority: 95 });
      }
    }

    // 패턴 E: 연수코드 또는 사이버-2026-xxxx 형태
    const mCyber = /([A-Za-z0-9가-힣]{2,15}-\d{4}-\d{3,10})/g;
    while ((match = mCyber.exec(text)) !== null) {
      const val = match[1].trim();
      candidates.push({ val, priority: 80 });
    }

    // 패턴 F: 줄 단위에서 이수번호/등록번호가 포함된 경우
    lines.forEach(line => {
      if (/이수.*번호|등록.*번호|발급.*번호/i.test(line)) {
        const clean = line.replace(/.*(?:이수.*번호|등록.*번호|발급.*번호)[:：\s]*/i, '').trim();
        if (clean && /\d/.test(clean) && clean.length >= 4) {
          candidates.push({ val: clean, priority: 85 });
        }
      }
    });

    if (candidates.length > 0) {
      // 우선순위 정렬 후 가장 적합한 번호 선택
      candidates.sort((a, b) => b.priority - a.priority);
      // 가장 긴 유의미한 문자열 우선
      certNo = candidates[0].val.replace(/\s{2,}/g, ' ');
    }

    return {
      certNo: certNo || '',
      certType: certType || '원격 연수',
      rawText: text
    };
  }

  /**
   * 이수증 이미지(또는 DataUrl) 스캔 실행 함수
   * @param {string} dataUrl - 이미지 DataURL
   * @param {function} onProgress - (progress: number, status: string) => void
   * @returns {Promise<{ certNo: string, certType: string, rawText: string }>}
   */
  async function scanCertificateImage(dataUrl, onProgress) {
    if (typeof onProgress === 'function') onProgress(10, 'OCR 엔진 로드 중...');
    const Tesseract = await loadTesseract();

    if (typeof onProgress === 'function') onProgress(25, '이미지 화질 및 대비 보정 중...');
    const canvas = await preprocessImage(dataUrl);

    if (typeof onProgress === 'function') onProgress(40, '문자 및 이수번호 인식 중...');

    // Tesseract v5 worker 생성 및 언어 설정 (한국어 + 영어)
    const worker = await Tesseract.createWorker(['kor', 'eng'], 1, {
      logger: m => {
        if (m.status === 'recognizing text' && typeof onProgress === 'function') {
          const p = 40 + Math.round((m.progress || 0) * 55);
          onProgress(Math.min(95, p), `이수번호 분석 중... (${p}%)`);
        }
      }
    });

    try {
      const ret = await worker.recognize(canvas);
      const rawText = (ret && ret.data && ret.data.text) ? ret.data.text : '';
      await worker.terminate();

      if (typeof onProgress === 'function') onProgress(98, '이수번호 패턴 추출 중...');
      const parsed = extractCertificateInfo(rawText);
      if (typeof onProgress === 'function') onProgress(100, '완료!');
      return parsed;
    } catch (err) {
      try { await worker.terminate(); } catch (e) {}
      throw err;
    }
  }

  window.CertScanner = {
    loadTesseract,
    scanCertificateImage,
    extractCertificateInfo,
  };
})();
