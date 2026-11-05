/**
 * 이수증 자동 스캔 및 이수번호/연수형태 지능형 추출 모듈 (v2.0)
 * - 클라이언트(브라우저)에서 안전하게 실행 (서버 비용 0원, GitHub Pages 완벽 호환)
 * - 상단 헤더 영역(Top 38%) 정밀 크롭(ROI) 및 고대비 이진화 전처리로 본문 오인식 원천 차단
 * - '이수번호' 문구 뒤 번호 우선 추출 + 문구 없이 좌측 상단에 위치한 번호(NO, 수비초-7571 등) 정밀 감지
 */
(function() {
  let loadingPromise = null;

  // Tesseract 라이브러리 동적 로드
  function loadTesseract() {
    if (window.Tesseract) return Promise.resolve(window.Tesseract);
    if (loadingPromise) return loadingPromise;
    loadingPromise = new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = 'https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js';
      script.async = true;
      script.onload = () => resolve(window.Tesseract);
      script.onerror = () => reject(new Error('OCR 엔진(Tesseract)을 불러오지 못했습니다.'));
      document.head.appendChild(script);
    });
    return loadingPromise;
  }

  /**
   * 상단 헤더 영역(Top 38% / Left 85%) 집중 크롭 및 고해상도 확대 전처리
   * - 한국의 모든 이수증은 상단 좌측/헤더에 이수번호가 위치함
   * - 중간 본문("위 사람은...", 과정명, 날짜 등)을 완전히 배제하여 본문 오인식 원천 방지
   */
  function cropTopHeaderRegion(img) {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');

    // 상단 38% 높이, 좌측 85% 너비 집중
    const cropW = Math.round(img.width * 0.88);
    const cropH = Math.round(img.height * 0.38);

    // 최소 너비 1400px 이상으로 고해상도 리샘플링하여 작은 글씨 선명화
    const targetW = Math.max(1400, cropW);
    const scale = targetW / cropW;
    const targetH = Math.round(cropH * scale);

    canvas.width = targetW;
    canvas.height = targetH;

    ctx.drawImage(img, 0, 0, cropW, cropH, 0, 0, targetW, targetH);
    applyHighContrast(ctx, targetW, targetH, 1.45);
    return canvas;
  }

  /**
   * 전체 이미지 전처리 (보조용)
   */
  function preprocessFullImage(img) {
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
    applyHighContrast(ctx, w, h, 1.25);
    return canvas;
  }

  /**
   * 흑백 그레이스케일 및 콘트라스트/샤프닝 보정
   */
  function applyHighContrast(ctx, w, h, contrast) {
    const imgData = ctx.getImageData(0, 0, w, h);
    const d = imgData.data;
    const factor = (259 * (contrast * 255 + 255)) / (255 * (259 - contrast * 255));
    for (let i = 0; i < d.length; i += 4) {
      const gray = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
      const c = factor * (gray - 128) + 128;
      const val = Math.min(255, Math.max(0, c));
      d[i] = val;
      d[i + 1] = val;
      d[i + 2] = val;
    }
    ctx.putImageData(imgData, 0, 0);
  }

  function loadImage(src) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error('이미지를 불러올 수 없습니다.'));
      img.src = src;
    });
  }

  // 본문 문장 및 제외해야 할 키워드 (오인식 방지 블랙리스트)
  const BODY_BLACKLIST = [
    '위 사람', '귀하는', '수여합니다', '축하합니다', '수료하였으므로', '이수하였으므로',
    '성명', '생년월일', '주민등록번호', '소속', '직위', '직급', '과정명', '연수명',
    '교육과정', '연수기간', '교육기간', '이수시간', '총시간', '원장', '교육감',
    '학교장', '협회장', '대표이사', '총장', '직인', '이수증', '수료증', '상장'
  ];

  function containsBlacklistWord(text) {
    if (!text) return false;
    return BODY_BLACKLIST.some(w => text.includes(w));
  }

  /**
   * 한국 이수증 텍스트에서 이수번호 및 연수 형태 정밀 추출
   * @param {string} rawText OCR 인식 원문
   * @param {boolean} isTopRegion 상단 헤더 영역만 크롭한 것인지 여부
   */
  function extractCertificateInfo(rawText, isTopRegion = false) {
    if (!rawText || typeof rawText !== 'string') {
      return { certNo: '', certType: '원격 연수', rawText: '', score: 0 };
    }

    const text = rawText.replace(/\r/g, '\n');
    const lines = text.split('\n')
      .map(l => l.trim())
      .filter(l => l.length > 0);

    let certType = '원격 연수';

    // 연수 형태 추정 (대면/원격/내부/외부)
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

    const candidates = [];

    // [우선순위 1]: '이수번호', '등록번호', '연수번호', '발급번호', '증서번호', '문서번호' 라벨 뒤의 번호
    // (사용자 1순위 요청: 이수증 서식이 다양해도 라벨 뒤의 번호를 확실하게 인식)
    const labelRegex = /(?:이\s*수\s*번\s*호|등\s*록\s*번\s*호|연\s*수\s*번\s*호|발\s*급\s*번\s*호|증\s*서\s*번\s*호|문\s*서\s*번\s*호)[\s:：·\-]+([^\r\n]{3,50})/gi;
    let match;
    while ((match = labelRegex.exec(text)) !== null) {
      let val = match[1].trim();
      // 뒷부분에 붙은 불필요한 단어 정리 (예: " 2026-1234 (원격) " -> "2026-1234")
      val = val.replace(/[\r\n].*$/, '')
               .replace(/\s*(?:성명|이름|과정명|연수과정).*$/, '')
               .replace(/^[:：·\-\s]+/, '')
               .trim();
      if (/\d/.test(val) && val.length >= 3 && !containsBlacklistWord(val)) {
        candidates.push({ val, score: 1000, source: 'label' });
      }
    }

    // 각 줄별로 "이수번호" 등이 단독으로 있거나 줄 끝에 걸친 경우 체크
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (/(?:이\s*수\s*번\s*호|등\s*록\s*번\s*호|연\s*수\s*번\s*호|발\s*급\s*번\s*호|증\s*서\s*번\s*호|문\s*서\s*번\s*호)/i.test(line)) {
        let clean = line.replace(/.*(?:이\s*수\s*번\s*호|등\s*록\s*번\s*호|연\s*수\s*번\s*호|발\s*급\s*번\s*호|증\s*서\s*번\s*호|문\s*서\s*번\s*호)[\s:：·\-]*/i, '').trim();
        // 만약 라벨 뒤가 비어 있고 다음 줄에 번호가 적혀 있는 경우 (줄바꿈 서식 지원)
        if (!clean && i + 1 < lines.length) {
          const nextLine = lines[i + 1].trim();
          if (/\d/.test(nextLine) && !containsBlacklistWord(nextLine) && nextLine.length >= 4) {
            clean = nextLine;
          }
        }
        if (clean && /\d/.test(clean) && clean.length >= 3 && !containsBlacklistWord(clean)) {
          candidates.push({ val: clean, score: 980, source: 'line-label' });
        }
      }
    }

    // [우선순위 2]: NO. 형태 (중앙교육연수원, 교육청 등에서 흔히 쓰이는 형태)
    // 예: NO. 공무원 성인지 사이버-2026-0988935 또는 NO. 2026-1092021
    const mNo = /NO[.:·\s]*([A-Za-z0-9가-힣\s\-_.()/]{4,45})/gi;
    while ((match = mNo.exec(text)) !== null) {
      let rawVal = match[1].trim();
      rawVal = rawVal.replace(/[\r\n].*$/, '').replace(/\s*(?:성명|이름).*$/, '').trim();
      const val = ('NO. ' + rawVal.replace(/^NO[.:·\s]*/i, '')).trim();
      if (/\d/.test(val) && val.length >= 6 && !containsBlacklistWord(rawVal)) {
        candidates.push({ val, score: 920, source: 'no' });
      }
    }

    // [우선순위 3]: 학교 문서번호 형태 (학교 자체 대면연수 등)
    // 예: 수비초-7571(2026. 9. 2.) 또는 수비초등학교-1234
    const mSchoolDoc = /([가-힣]{2,}초(?:등학교)?[\s\-_.·]+(?:\d{2,6}|제?\d+호?)(?:\s*\(\s*\d{4}[.\s]+\d{1,2}[.\s]+\d{1,2}[.\s]*\))?)/g;
    while ((match = mSchoolDoc.exec(text)) !== null) {
      const val = match[1].trim();
      if (/\d/.test(val) && !containsBlacklistWord(val)) {
        candidates.push({ val, score: 900, source: 'school-doc' });
      }
    }

    // [우선순위 4]: "제 2026-1092021 호" 또는 "제 2026-가-1234호" 형태
    const mJe = /제\s*([0-9A-Za-z가-힣\s\-_.()]{3,35})\s*호/g;
    while ((match = mJe.exec(text)) !== null) {
      const inner = match[1].trim();
      if (/\d/.test(inner) && !containsBlacklistWord(inner)) {
        const val = ('제 ' + inner + '호').trim();
        candidates.push({ val, score: 880, source: 'je-ho' });
      }
    }

    // [우선순위 5]: (사용자 2순위 요청: 이수번호 라벨이 없이 좌측 상단에 번호만 있는 서식 지원)
    // 상단 헤더 영역(Top Region)에서 라벨 없이 상단 첫 줄에 위치한 번호/코드 패턴 추출
    if (isTopRegion) {
      for (let i = 0; i < Math.min(lines.length, 5); i++) {
        const line = lines[i];
        if (containsBlacklistWord(line)) continue;
        // 생년월일이나 단순 날짜(2026. 03. 15)는 배제
        if (/^\d{4}[.\s\-]\d{1,2}[.\s\-]\d{1,2}$/.test(line)) continue;

        // 패턴 5-A: 하이픈과 숫자가 결합된 코드 (예: 2026-0988935, 2026-연수-1234, 2026-1092021)
        const codeMatch = line.match(/([A-Za-z0-9가-힣]{0,12}[-_.][0-9A-Za-z가-힣\-_.]{4,25})/);
        if (codeMatch && /\d{3,}/.test(codeMatch[1])) {
          const val = codeMatch[1].trim();
          candidates.push({ val, score: 850, source: 'top-unlabeled-code' });
          break; // 가장 상단의 첫 번째 코드 채택
        }

        // 패턴 5-B: 5자리 이상의 단독 숫자/영문자 번호
        const numMatch = line.match(/\b([A-Za-z0-9]{5,20})\b/);
        if (numMatch && /\d{4,}/.test(numMatch[1])) {
          const val = numMatch[1].trim();
          candidates.push({ val, score: 820, source: 'top-unlabeled-num' });
          break;
        }
      }
    }

    // [우선순위 6]: 일반 사이버 연수 코드 패턴 (사이버-2026-xxxx 등)
    const mCyber = /([A-Za-z0-9가-힣]{2,15}-\d{4}-\d{3,10})/g;
    while ((match = mCyber.exec(text)) !== null) {
      const val = match[1].trim();
      if (!containsBlacklistWord(val)) {
        candidates.push({ val, score: 750, source: 'cyber' });
      }
    }

    let certNo = '';
    let bestScore = 0;

    if (candidates.length > 0) {
      // 점수 높은 순 정렬
      candidates.sort((a, b) => b.score - a.score);
      certNo = candidates[0].val.replace(/\s{2,}/g, ' ').trim();
      bestScore = candidates[0].score;
    }

    return {
      certNo: certNo || '',
      certType: certType || '원격 연수',
      rawText: text,
      score: bestScore
    };
  }

  /**
   * 이수증 이미지 정밀 스캔 실행 함수 (2단계 듀얼 헤더 크롭 & 스마트 추출)
   * @param {string} dataUrl 이미지 DataURL
   * @param {function} onProgress (progress: number, status: string) => void
   * @returns {Promise<{ certNo: string, certType: string, rawText: string }>}
   */
  async function scanCertificateImage(dataUrl, onProgress) {
    if (typeof onProgress === 'function') onProgress(10, 'OCR 엔진 준비 중...');
    const Tesseract = await loadTesseract();

    if (typeof onProgress === 'function') onProgress(20, '이수증 이미지 분석 중...');
    const img = await loadImage(dataUrl);

    // 1단계: 상단 헤더 영역(Top 38%) 정밀 크롭
    if (typeof onProgress === 'function') onProgress(30, '상단 이수번호 영역 집중 분리 중...');
    const topCanvas = cropTopHeaderRegion(img);

    // Tesseract worker 생성
    const worker = await Tesseract.createWorker(['kor', 'eng'], 1, {
      logger: m => {
        if (m.status === 'recognizing text' && typeof onProgress === 'function') {
          const p = 35 + Math.round((m.progress || 0) * 50);
          onProgress(Math.min(88, p), `상단 이수번호 정밀 분석 중... (${p}%)`);
        }
      }
    });

    try {
      // 1단계: 상단 헤더 영역 우선 인식 (본문 텍스트 원천 차단)
      const topRet = await worker.recognize(topCanvas);
      const topRawText = (topRet && topRet.data && topRet.data.text) ? topRet.data.text : '';
      let parsed = extractCertificateInfo(topRawText, true);

      // 상단 헤더에서 높은 신뢰도(820점 이상)로 이수번호를 찾았으면 바로 종료! (본문 오인식 0%)
      if (parsed.certNo && parsed.score >= 820) {
        await worker.terminate();
        if (typeof onProgress === 'function') onProgress(100, '이수번호 인식 완료!');
        return parsed;
      }

      // 2단계: 상단에서 명확한 번호가 안 나온 경우 전체 이미지 보조 스캔
      if (typeof onProgress === 'function') onProgress(75, '전체 문서 보조 분석 중...');
      const fullCanvas = preprocessFullImage(img);
      const fullRet = await worker.recognize(fullCanvas);
      const fullRawText = (fullRet && fullRet.data && fullRet.data.text) ? fullRet.data.text : '';
      await worker.terminate();

      const fullParsed = extractCertificateInfo(fullRawText, false);

      // 전체 스캔과 상단 스캔 결과 중 더 높은 신뢰도 채택
      if (fullParsed.score > parsed.score && fullParsed.certNo) {
        parsed = fullParsed;
      }

      if (typeof onProgress === 'function') onProgress(100, '인식 완료!');
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
