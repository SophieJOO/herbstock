// ========================================
// 🏥 아현재한의원 약재관리 통합 자동화 시스템
// OCR 자동화 (Vision API + Gemini) + FIFO 선입선출 + 실시간 원가계산
// Version: 8.1 (Gemini API 통합)
// ========================================

// ========================================
// 공통 유틸리티
// ========================================

/**
 * 설정값 가져오기
 */
function getConfig(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('설정');
  if (!sheet) {
    throw new Error('설정 시트가 없습니다. 먼저 설정 시트를 만들어주세요.');
  }
  
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return null;
}

/**
 * 폴더 생성 또는 가져오기
 */
function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

/**
 * 오류 로깅
 */
function logError(fileName, errorMessage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let errorSheet = ss.getSheetByName('오류로그');
  
  if (!errorSheet) {
    errorSheet = ss.insertSheet('오류로그');
    errorSheet.appendRow(['일시', '파일명', '오류내용']);
  }
  
  errorSheet.appendRow([
    new Date(),
    fileName,
    errorMessage
  ]);
}

/**
 * 재고 부족 Slack 알람
 */
function sendSlackAlert(message) {
  const webhookUrl = getConfig('slack_긴급알람_webhook');
  
  if (!webhookUrl) {
    Logger.log('⚠️ Slack 긴급알람 Webhook URL이 설정되지 않았습니다.');
    return;
  }
  
  const payload = {
    text: message,
    username: '약재재고알람',
    icon_emoji: ':pill:'
  };
  
  sendSlackMessage(webhookUrl, payload);
  Logger.log('✅ Slack 알람 발송: ' + message);
}

/**
 * 일반 Slack 알림 (기존 알림용)
 */
function sendSlackNotification(message) {
  const webhookUrl = getConfig('slack_일반알림_webhook');
  
  if (!webhookUrl) {
    Logger.log('⚠️ Slack 일반알림 Webhook URL이 설정되지 않았습니다.');
    return;
  }
  
  const payload = {
    text: message,
    username: '한의원자동화',
    icon_emoji: ':herb:'
  };
  
  sendSlackMessage(webhookUrl, payload);
}

// ========================================
// 📥 입고 자동화 - PART 1: OCR 처리
// ========================================

/**
 * 입고서 이미지를 OCR 처리하여 임시입고 시트에 입력
 * 트리거: 5분마다 자동 실행
 */
function processIncomingImagesOCR() {
  const folderId = getConfig('입고서_폴더_ID');
  if (!folderId) {
    Logger.log('❌ 입고서 폴더 ID가 설정되지 않았습니다.');
    Logger.log('설정 시트에 "입고서_폴더_ID" 항목을 추가해주세요.');
    return;
  }

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();

  let processedCount = 0;
  let errorCount = 0;
  const MAX_FILES_PER_RUN = 10;  // ✅ 한 번에 최대 10개만 처리

  while (files.hasNext() && (processedCount + errorCount) < MAX_FILES_PER_RUN) {
    const file = files.next();
    const mimeType = file.getMimeType();
    
    // 이미지 파일만 처리
    if (mimeType.includes('image')) {
      try {
        Logger.log('📸 입고서 OCR 처리 중: ' + file.getName());
        
        // Google Vision API로 OCR 실행
        const ocrText = extractTextFromImage(file);
        Logger.log('OCR 결과:\n' + ocrText);
        
        // Gemini로 구조화된 데이터 추출
        const parsedData = parseIncomingDraftWithGemini(ocrText, file.getName());
        Logger.log('파싱 결과: ' + JSON.stringify(parsedData));
        
        if (parsedData && parsedData.items && parsedData.items.length > 0) {
          // 임시입고 시트에 추가
          addToTempIncomingSheet(parsedData, file);
          
          // 처리 완료 폴더로 이동
          const processedFolder = getOrCreateFolder(folder, 'OCR완료');
          file.moveTo(processedFolder);
          
          processedCount++;
          Logger.log('✅ OCR 추출 완료: ' + file.getName());
          
          // 슬랙 알림
          sendOCRCompletedSlack(parsedData, processedCount);
        }
        
      } catch (error) {
        Logger.log('❌ OCR 오류: ' + error.message);
        errorCount++;
        
        logError(file.getName(), error.message);
        
        const errorFolder = getOrCreateFolder(folder, '오류');
        file.moveTo(errorFolder);
      }
    }
  }
  
  if (processedCount > 0 || errorCount > 0) {
    Logger.log(`📊 OCR 처리 완료: ${processedCount}건 성공, ${errorCount}건 오류`);
  }
}

/**
 * Google Vision API로 이미지에서 텍스트 추출
 */
function extractTextFromImage(file) {
  const apiKey = getConfig('VISION_API_KEY');
  if (!apiKey) {
    throw new Error('VISION_API_KEY가 설정되지 않았습니다.');
  }
  
  const imageBlob = file.getBlob();
  const base64Image = Utilities.base64Encode(imageBlob.getBytes());
  
  const url = 'https://vision.googleapis.com/v1/images:annotate?key=' + apiKey;
  const payload = {
    requests: [{
      image: { content: base64Image },
      features: [{ type: 'TEXT_DETECTION' }]
    }]
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());
  
  if (result.responses && result.responses[0].fullTextAnnotation) {
    return result.responses[0].fullTextAnnotation.text;
  }
  
  throw new Error('OCR 실패: 텍스트를 추출할 수 없습니다.');
}

/**
 * Gemini API로 입고서 데이터 파싱 (JSON 복구 로직 포함)
 */
function parseIncomingDraftWithGemini(ocrText, fileName) {
  const apiKey = getConfig('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY가 설정되지 않았습니다.');
  }

  // ✅ OCR 텍스트 전처리 (불필요한 부분 제거)
  let cleanedText = ocrText;

  Logger.log(`📊 원본 OCR 텍스트 길이: ${cleanedText.length}자`);

  // 1. 연속된 공백/줄바꿈 정리
  cleanedText = cleanedText.replace(/\s+/g, ' ').trim();

  // 2. 특수문자 제거 (한글, 숫자, 기본 구두점만 남김)
  cleanedText = cleanedText.replace(/[^\u3131-\u318E\uAC00-\uD7A3a-zA-Z0-9\s\.,:\-\/]/g, '');

  // 3. 텍스트가 너무 길면 제한 (단계적 제한)
  const MAX_LENGTH = 3000;  // 5000 → 3000으로 더 줄임

  if (cleanedText.length > MAX_LENGTH) {
    Logger.log(`⚠️ OCR 텍스트가 ${cleanedText.length}자로 너무 깁니다. ${MAX_LENGTH}자로 제한합니다.`);
    cleanedText = cleanedText.substring(0, MAX_LENGTH);
  }

  Logger.log(`📊 정리된 OCR 텍스트 길이: ${cleanedText.length}자`);

  const prompt = `한의원 약재 입고서 OCR 텍스트를 분석하여 JSON으로 변환하세요.

아래 JSON 형식으로만 응답하세요 (설명 없이 JSON만):
{
  "incomingDate": "YYYY-MM-DD",
  "supplier": "공급처명",
  "items": [
    {
      "herbName": "약재명",
      "bagSize": 600,
      "quantity": 2,
      "unitPrice": 11000,
      "totalPrice": 22000,
      "confidence": "high"
    }
  ]
}

confidence: high/medium/low 중 선택
반드시 완전한 JSON 출력, 끝에 ] } 닫기

OCR 텍스트:
${cleanedText}`;

  // ✅ 토큰 수 증가 + 더 안정적인 모델
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  
  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 8192,  // ✅ 토큰 제한 증가 (4096 → 8192)
      topP: 0.8,
      topK: 40
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // ✅ 재시도 로직 (503 에러 대응)
  const MAX_RETRIES = 3;
  let lastError = null;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      if (attempt > 1) {
        const waitTime = attempt * 2000; // 2초, 4초, 6초
        Logger.log(`⏳ ${waitTime/1000}초 대기 후 재시도 (${attempt}/${MAX_RETRIES})...`);
        Utilities.sleep(waitTime);
      }

      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      Logger.log('=== Gemini API 응답 (입고서) ===');
      Logger.log('HTTP 상태: ' + responseCode);
      Logger.log('응답 길이: ' + responseText.length + ' 문자');
      if (attempt > 1) {
        Logger.log(`✅ 재시도 ${attempt}번째 성공`);
      }

      // ✅ 503 에러는 재시도
      if (responseCode === 503) {
        Logger.log('⚠️ 503 에러: Gemini API 과부하');
        lastError = new Error('Gemini API 서버 과부하 (503)');
        continue; // 재시도
      }

      if (responseCode !== 200) {
        Logger.log('❌ 전체 응답: ' + responseText);
        throw new Error(`Gemini API 오류 (HTTP ${responseCode}): ${responseText}`);
      }

      const result = JSON.parse(responseText);

      if (result.error) {
        // 503 에러 체크
        if (result.error.code === 503) {
          Logger.log('⚠️ 503 에러: ' + result.error.message);
          lastError = new Error(`Gemini API 서버 과부하: ${result.error.message}`);
          continue; // 재시도
        }
        throw new Error(`Gemini API 오류: ${result.error.message} (코드: ${result.error.code})`);
      }

      if (!result.candidates || !result.candidates[0]) {
        throw new Error('Gemini API 응답에 candidates가 없습니다.');
      }

      const candidate = result.candidates[0];

    // finishReason 확인 - 중단되었는지 체크
    const finishReason = candidate.finishReason || 'UNKNOWN';
    Logger.log(`📌 종료 이유: ${finishReason}`);

    // MAX_TOKENS로 잘렸고 content가 없거나 너무 짧으면 재시도
    if (finishReason === 'MAX_TOKENS') {
      Logger.log('⚠️ 토큰 제한으로 응답이 잘렸습니다.');

      // content가 없거나 비어있으면 에러
      if (!candidate.content || !candidate.content.parts || !candidate.content.parts[0] || !candidate.content.parts[0].text) {
        Logger.log('❌ MAX_TOKENS이지만 응답 내용이 없습니다. OCR 텍스트가 너무 길거나 복잡합니다.');
        throw new Error('Gemini 토큰 제한 초과: OCR 텍스트가 너무 길어 처리할 수 없습니다. 이미지를 더 깔끔하게 찍어주세요.');
      }

      // 응답이 있지만 잘렸다면 복구 시도
      Logger.log('⚠️ 응답이 잘렸지만 일부 내용이 있습니다. 복구 시도...');
    }

    if (!candidate.content || !candidate.content.parts || !candidate.content.parts[0]) {
      Logger.log('❌ 응답 구조: ' + JSON.stringify(candidate));
      throw new Error('Gemini API 응답 구조가 올바르지 않습니다.');
    }
    
    let textContent = candidate.content.parts[0].text;
    Logger.log('원본 응답 (처음 500자): ' + textContent.substring(0, 500));
    Logger.log('원본 응답 (마지막 200자): ' + textContent.substring(Math.max(0, textContent.length - 200)));
    
    // JSON 추출 및 정리
    textContent = textContent.trim();
    textContent = textContent.replace(/```json\s*/gi, '');
    textContent = textContent.replace(/```\s*/g, '');
    textContent = textContent.trim();
    
    // JSON 객체 추출
    const jsonStart = textContent.indexOf('{');
    const jsonEnd = textContent.lastIndexOf('}');
    
    if (jsonStart === -1) {
      Logger.log('❌ JSON 시작 찾기 실패. 전체 텍스트: ' + textContent);
      throw new Error('응답에서 JSON 형식을 찾을 수 없습니다.');
    }
    
    let jsonText;
    
    // ✅ JSON 복구 로직 (개선)
    if (jsonEnd === -1 || jsonEnd < jsonStart) {
      Logger.log('⚠️ JSON이 불완전합니다. 자동 복구 시도...');

      jsonText = textContent.substring(jsonStart);

      // 1. 불완전한 필드 제거 (마지막 쉼표 이후)
      const lastComma = jsonText.lastIndexOf(',');
      const lastCloseBrace = jsonText.lastIndexOf('}');
      const lastCloseBracket = jsonText.lastIndexOf(']');

      // 마지막 완전한 객체까지만 사용
      if (lastCloseBrace !== -1 && lastComma > lastCloseBrace) {
        // 마지막 완전한 객체 이후 불완전한 부분 제거
        jsonText = jsonText.substring(0, lastCloseBrace + 1);
      }

      // 2. items 배열 닫기
      if (jsonText.includes('"items"') && !jsonText.includes('items":[')) {
        // items가 시작조차 안된 경우
        jsonText += ', "items": []}';
      } else if (jsonText.includes('"items"') && jsonText.lastIndexOf(']') < jsonText.lastIndexOf('[')) {
        // items 배열이 열렸지만 닫히지 않은 경우
        jsonText += '\n  ]\n}';
      } else if (!jsonText.endsWith('}')) {
        // 최종 객체가 닫히지 않은 경우
        jsonText += '\n}';
      }

      Logger.log('✅ 복구된 JSON (처음 500자): ' + jsonText.substring(0, 500));
      Logger.log('✅ 복구된 JSON (마지막 200자): ' + jsonText.substring(Math.max(0, jsonText.length - 200)));
    } else {
      jsonText = textContent.substring(jsonStart, jsonEnd + 1);
    }
    
    Logger.log('최종 JSON (파싱 시도): ' + jsonText);
    
    try {
      const parsed = JSON.parse(jsonText);
      parsed.fileName = fileName;
      
      if (!parsed.items || !Array.isArray(parsed.items) || parsed.items.length === 0) {
        throw new Error('약재 항목이 없습니다.');
      }
      
      Logger.log('✅ JSON 파싱 성공: ' + parsed.items.length + '개 항목');
      return parsed;  // ✅ 성공 - 재시도 루프 탈출

    } catch (parseError) {
      Logger.log('❌ JSON 파싱 오류: ' + parseError.message);
      Logger.log('파싱 시도한 JSON: ' + jsonText);
      throw new Error(`JSON 파싱 실패: ${parseError.message}`);
    }

    } catch (error) {
      // 503 에러는 재시도, 다른 에러는 즉시 throw
      if (error.message && error.message.includes('503')) {
        lastError = error;
        Logger.log(`⚠️ 시도 ${attempt}/${MAX_RETRIES} 실패: ${error.message}`);
        if (attempt < MAX_RETRIES) {
          continue; // 재시도
        }
      } else {
        // 503이 아닌 다른 에러는 즉시 throw
        Logger.log('❌ Gemini API 호출 오류 (재시도 불가): ' + error.message);
        throw error;
      }
    }
  }

  // 모든 재시도 실패
  Logger.log(`❌ ${MAX_RETRIES}번 재시도 모두 실패`);
  throw lastError || new Error('Gemini API 호출 실패');
}

/**
 * 임시입고 시트에 OCR 결과 추가
 */
function addToTempIncomingSheet(data, file) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tempSheet = ss.getSheetByName('임시입고');
  
  // 시트가 없으면 생성
  if (!tempSheet) {
    tempSheet = ss.insertSheet('임시입고');
    
    const headers = [
      '입고일', '공급처', '약재명', '수량(봉지)', '봉지단위(g)', 
      '총량(g)', '단가(원/봉)', '총금액', 'g당단가(원/g)', '유통기한_입력',
      '확신도', '입고서파일', '✅처리완료', '비고'
    ];
    tempSheet.appendRow(headers);
    
    // 헤더 스타일링
    const headerRange = tempSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // 열 너비 조정
    tempSheet.setColumnWidth(3, 120);  // 약재명
    tempSheet.setColumnWidth(9, 100);  // g당단가
    tempSheet.setColumnWidth(10, 200); // 유통기한 입력
    tempSheet.setColumnWidth(12, 200); // 입고서파일
    tempSheet.setColumnWidth(14, 250); // 비고
  }
  
  const fileUrl = file.getUrl();
  
  // 각 약재별로 행 추가
  data.items.forEach(item => {
    const totalAmount = item.bagSize && item.quantity ? item.bagSize * item.quantity : '';
    const unitPrice = item.totalPrice && item.quantity ? Math.round(item.totalPrice / item.quantity) : '';
    
    // g당 단가 계산
    let pricePerGram = '';
    if (item.totalPrice && totalAmount) {
      pricePerGram = Math.round((item.totalPrice / totalAmount) * 10) / 10;
    } else if (unitPrice && item.bagSize) {
      pricePerGram = Math.round((unitPrice / item.bagSize) * 10) / 10;
    }
    
    // 유통기한 입력 가이드
    let expiryDateGuide = '';
    if (item.quantity && item.quantity > 1) {
      const dates = [];
      for (let i = 1; i <= item.quantity; i++) {
        dates.push(`봉지${i}: YYYY-MM-DD`);
      }
      expiryDateGuide = dates.join(', ');
    } else {
      expiryDateGuide = 'YYYY-MM-DD';
    }
    
    const row = [
      data.incomingDate || new Date().toISOString().split('T')[0],
      data.supplier || '',
      item.herbName,
      item.quantity,
      item.bagSize || '',
      totalAmount,
      unitPrice,
      item.totalPrice || '',
      pricePerGram,
      expiryDateGuide,
      item.confidence || 'unknown',
      fileUrl,
      '',  // 처리완료 체크박스
      item.bagSize ? `✅ 자동입력 완료 (g당 ${pricePerGram}원) → 유통기한만 입력` : '⚠️ 봉지단위 입력 필요'
    ];
    
    tempSheet.appendRow(row);
    
    const lastRow = tempSheet.getLastRow();
    
    // 확신도 색상 표시
    const confidenceCell = tempSheet.getRange(lastRow, 11);
    if (item.confidence === 'high') {
      confidenceCell.setBackground('#d9ead3');
    } else if (item.confidence === 'medium') {
      confidenceCell.setBackground('#fff2cc');
    } else {
      confidenceCell.setBackground('#f4cccc');
    }
    
    // g당 단가 색상
    if (pricePerGram) {
      tempSheet.getRange(lastRow, 9).setBackground('#d9ead3');
    }
    
    // 유통기한 입력란 강조
    tempSheet.getRange(lastRow, 10).setBackground('#fff2cc');
    
    // 봉지단위 누락 시 강조
    if (!item.bagSize) {
      tempSheet.getRange(lastRow, 5).setBackground('#fff2cc');
    }
    
    // 처리완료 체크박스 생성
    const checkboxCell = tempSheet.getRange(lastRow, 13);
    checkboxCell.insertCheckboxes();
    checkboxCell.setValue(false);
    checkboxCell.setHorizontalAlignment('center');
  });
  
  Logger.log(`✅ 임시입고 시트에 ${data.items.length}개 약재 추가됨`);
}

// ========================================
// 📥 입고 자동화 - PART 2: 약재입고 이동 (FIFO 준비)
// ========================================

/**
 * 편집 트리거: 처리완료 체크 시 자동 입고
 */
function onTempIncomingEdit(e) {
  try {
    if (!e || !e.source) {
      Logger.log('❌ 이 함수는 수동 실행할 수 없습니다.');
      Browser.msgBox('안내', '스프레드시트에서 "처리완료" 체크박스를 체크하세요.', Browser.Buttons.OK);
      return;
    }
    
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    
    if (sheet.getName() !== '임시입고') return;
    
    // 13열(M열)이 처리완료 컬럼
    if (range.getColumn() === 13 && range.getValue() === true) {
      const row = range.getRow();
      if (row === 1) return;  // 헤더 제외
      
      Logger.log(`✅ 처리완료 체크: ${row}행 자동 입고 시작`);
      moveToIncomingSheet(row);
    }
  } catch (error) {
    Logger.log('편집 트리거 오류: ' + error.message);
    Browser.msgBox('오류', '처리 중 오류 발생: ' + error.message, Browser.Buttons.OK);
  }
}

/**
 * 임시입고 → 약재입고 (봉지별 분리 + 잔량 관리)
 */
/**
 * 임시입고 → 약재입고 시트로 이동 (편집 트리거 최적화)
 */
function moveToIncomingSheet(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.getSheetByName('임시입고');
  let incomingSheet = ss.getSheetByName('약재입고');
  
  // 약재입고 시트가 없으면 생성
  if (!incomingSheet) {
    incomingSheet = ss.insertSheet('약재입고');
    
    const headers = [
      '입고번호', '입고일', '약재명', '수량(g)', '유통기한', '잔량(g)',
      '단가(원/g)', '공급처', '입고가격(원)', '비고', '원본파일'
    ];
    incomingSheet.appendRow(headers);
    
    const headerRange = incomingSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#34a853');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
  }
  
  // 임시입고 시트에서 데이터 읽기
  const data = tempSheet.getRange(row, 1, 1, 14).getValues()[0];
  
  const incomingDate = data[0];
  const supplier = data[1];
  const herbName = data[2];
  const quantity = parseInt(data[3]) || 0;
  const bagSize = parseFloat(data[4]) || 0;
  const totalAmount = data[5];
  const unitPrice = data[6];
  const totalPrice = data[7];
  const pricePerGram = data[8];
  const expiryDateInput = data[9];
  const fileUrl = data[11];
  
  Logger.log(`=== 입고 처리 시작 ===`);
  Logger.log(`약재명: ${herbName}`);
  Logger.log(`수량(봉지): ${quantity}`);
  Logger.log(`봉지단위: ${bagSize}g`);
  
  // 유효성 검사
  if (!quantity || quantity <= 0) {
    throw new Error('봉지 수량이 올바르지 않습니다: ' + quantity);
  }
  
  if (!bagSize || bagSize <= 0) {
    throw new Error('봉지 단위(g)가 올바르지 않습니다: ' + bagSize);
  }
  
  const expiryDates = parseExpiryDates(expiryDateInput, quantity);
  
  if (expiryDates.length === 0) {
    throw new Error('유통기한 형식이 올바르지 않습니다: ' + expiryDateInput);
  }
  
  if (expiryDates.length !== quantity) {
    Logger.log(`⚠️ 봉지 수(${quantity})와 유통기한 수(${expiryDates.length}) 불일치 - 마지막 값으로 채움`);
  }
  
  // 한 번에 여러 행 추가
  const rowsToAdd = [];
  
  Logger.log(`\n🔄 ${quantity}개 봉지를 입고 처리합니다...`);
  
  for (let i = 0; i < quantity; i++) {
    const incomingNumber = generateIncomingNumber(incomingDate);
    const expiryDate = expiryDates[i] || expiryDates[expiryDates.length - 1];
    const amount = bagSize;
    
    rowsToAdd.push([
      incomingNumber,
      incomingDate,
      herbName,
      amount,
      expiryDate,
      amount,  // 초기 잔량 = 수량
      pricePerGram,
      supplier,
      unitPrice,
      `${i + 1}/${quantity} 봉지`,
      fileUrl
    ]);
    
    Logger.log(`📦 봉지 ${i + 1}: ${incomingNumber} | ${amount}g | ${expiryDate}`);
  }
  
  // 한 번에 모든 행 추가
  if (rowsToAdd.length > 0) {
    const lastRow = incomingSheet.getLastRow();
    incomingSheet.getRange(lastRow + 1, 1, rowsToAdd.length, rowsToAdd[0].length)
      .setValues(rowsToAdd);
  }
  
  Logger.log(`✅ 입고 완료: ${herbName} ${quantity}봉 (총 ${bagSize * quantity}g)`);
  
  // 임시입고 시트에서 해당 행 삭제
  tempSheet.deleteRow(row);
  
  // ✅ 약재마스터 재고 즉시 업데이트 (이 약재만)
  updateSingleHerbStock(herbName);
  
  Logger.log(`=== 입고 처리 종료 ===\n`);
}

/**
 * 유통기한 파싱 (개선 버전 - Date 객체, 문자열 모두 처리)
 */
function parseExpiryDates(expiryDateInput, quantity) {
  const expiryDates = [];
  
  // 빈 값 체크
  if (!expiryDateInput) {
    Logger.log('⚠️ 유통기한이 입력되지 않았습니다.');
    return expiryDates;
  }
  
  // Date 객체인 경우 (Google Sheets가 자동 변환한 경우)
  if (expiryDateInput instanceof Date) {
    Logger.log('✅ Date 객체로 입력됨: ' + expiryDateInput);
    const formattedDate = Utilities.formatDate(expiryDateInput, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    // 봉지 수만큼 같은 유통기한으로 채우기
    for (let i = 0; i < quantity; i++) {
      expiryDates.push(formattedDate);
    }
    
    Logger.log(`✅ 유통기한 ${quantity}개 생성: ${formattedDate}`);
    return expiryDates;
  }
  
  // 문자열로 변환
  let dateString = String(expiryDateInput).trim();
  
  if (dateString === '') {
    Logger.log('⚠️ 유통기한이 빈 문자열입니다.');
    return expiryDates;
  }
  
  Logger.log('입력된 유통기한 문자열: ' + dateString);
  
  // "봉지1: 2026-01-15, 봉지2: 2026-02-20" 형식 파싱
  if (dateString.includes('봉지')) {
    const parts = dateString.split(',');
    for (const part of parts) {
      // YYYY-MM-DD 또는 YYYY/MM/DD 또는 YYYY.MM.DD 형식 모두 허용
      const match = part.match(/(\d{4}[-/.]?\d{1,2}[-/.]?\d{1,2})/);
      if (match) {
        const dateStr = match[1].replace(/[/.]/g, '-'); // 구분자를 -로 통일
        const date = new Date(dateStr);
        if (!isNaN(date.getTime())) {
          const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          expiryDates.push(formattedDate);
          Logger.log(`✅ 파싱 성공: ${formattedDate}`);
        }
      }
    }
  } else {
    // 단일 날짜 (다양한 형식 허용)
    // YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD, YYYYMMDD 등
    const dateStr = dateString.replace(/[/.]/g, '-'); // 구분자를 -로 통일
    
    // YYYY-MM-DD 형식 시도
    let match = dateStr.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (match) {
      const year = match[1];
      const month = match[2].padStart(2, '0');
      const day = match[3].padStart(2, '0');
      const normalizedDate = `${year}-${month}-${day}`;
      const date = new Date(normalizedDate);
      
      if (!isNaN(date.getTime())) {
        const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        
        // 봉지 수만큼 같은 유통기한으로 채우기
        for (let i = 0; i < quantity; i++) {
          expiryDates.push(formattedDate);
        }
        
        Logger.log(`✅ 유통기한 ${quantity}개 생성: ${formattedDate}`);
        return expiryDates;
      }
    }
    
    // YYYYMMDD 형식 시도
    match = dateString.match(/(\d{8})/);
    if (match) {
      const dateStr = match[1];
      const year = dateStr.substring(0, 4);
      const month = dateStr.substring(4, 6);
      const day = dateStr.substring(6, 8);
      const date = new Date(`${year}-${month}-${day}`);
      
      if (!isNaN(date.getTime())) {
        const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        
        for (let i = 0; i < quantity; i++) {
          expiryDates.push(formattedDate);
        }
        
        Logger.log(`✅ 유통기한 ${quantity}개 생성: ${formattedDate}`);
        return expiryDates;
      }
    }
    
    Logger.log('⚠️ 날짜 형식을 인식할 수 없습니다: ' + dateString);
  }
  
  // 부족한 경우 마지막 날짜로 채우기
  if (expiryDates.length > 0 && expiryDates.length < quantity) {
    const lastDate = expiryDates[expiryDates.length - 1];
    Logger.log(`⚠️ 유통기한이 부족합니다. 마지막 날짜(${lastDate})로 채웁니다.`);
    while (expiryDates.length < quantity) {
      expiryDates.push(lastDate);
    }
  }
  
  return expiryDates;
}

/**
 * 입고번호 생성 (IN20251020-001 형식)
 */
function generateIncomingNumber(incomingDate) {
  const date = incomingDate ? new Date(incomingDate) : new Date();
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const incomingSheet = ss.getSheetByName('약재입고');
  
  if (!incomingSheet) {
    return `IN${dateStr}-001`;
  }
  
  const data = incomingSheet.getDataRange().getValues();
  let todayCount = 0;
  const prefix = `IN${dateStr}-`;
  
  for (let i = 1; i < data.length; i++) {
    const incomingNumber = data[i][0];
    if (incomingNumber && incomingNumber.toString().startsWith(prefix)) {
      todayCount++;
    }
  }
  
  const serialNumber = String(todayCount + 1).padStart(3, '0');
  return `${prefix}${serialNumber}`;
}

// ========================================
// 📤 처방 자동화 - PART 1: OCR 처리
// ========================================

/**
 * 처방전 이미지를 OCR 처리하여 처방입력/처방상세 시트에 입력
 * 트리거: 5분마다 자동 실행
 */
function processPrescriptionImages() {
  const folderId = getConfig('처방전_폴더_ID');
  if (!folderId) {
    Logger.log('❌ 처방전 폴더 ID가 설정되지 않았습니다.');
    Logger.log('설정 시트에 "처방전_폴더_ID" 항목을 추가해주세요.');
    return;
  }

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();

  let processedCount = 0;
  let errorCount = 0;
  const MAX_FILES_PER_RUN = 10;  // ✅ 한 번에 최대 10개만 처리

  while (files.hasNext() && (processedCount + errorCount) < MAX_FILES_PER_RUN) {
    const file = files.next();
    const mimeType = file.getMimeType();

    if (mimeType.includes('image')) {
      try {
        Logger.log('📋 처방전 OCR 처리 중: ' + file.getName());

        const ocrText = extractTextFromImage(file);
        Logger.log('OCR 결과:\n' + ocrText);

        const parsedData = parsePrescriptionWithGemini(ocrText);
        Logger.log('파싱 결과: ' + JSON.stringify(parsedData));

        if (parsedData && parsedData.herbs) {
          // 처방입력 시트에 추가 (처방번호 반환)
          const prescNumber = addPrescriptionToSheet(parsedData);

          // 처방상세 시트에 추가 (약재 목록)
          addPrescriptionDetailsToSheet(prescNumber, parsedData);

          const processedFolder = getOrCreateFolder(folder, '처리완료');
          file.moveTo(processedFolder);

          processedCount++;
          Logger.log('✅ 처방 입력 완료: ' + file.getName());
          Logger.log(`   - 처방번호: ${prescNumber}`);
          Logger.log(`   - 환자: ${parsedData.patientName}`);
          Logger.log(`   - 약재: ${parsedData.herbs.length}개`);

          sendPrescriptionProcessedSlack(parsedData);
        }

      } catch (error) {
        Logger.log('❌ 처방 OCR 오류: ' + error.message);
        errorCount++;

        logError(file.getName(), error.message);

        const errorFolder = getOrCreateFolder(folder, '오류');
        file.moveTo(errorFolder);
      }
    }
  }
  
  if (processedCount > 0 || errorCount > 0) {
    Logger.log(`📊 처방 OCR 처리 완료: ${processedCount}건 성공, ${errorCount}건 오류`);
  }
}

/**
 * Gemini API로 처방전 데이터 파싱 (아현재한의원 맞춤)
 */
function parsePrescriptionWithGemini(ocrText) {
  const apiKey = getConfig('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY가 설정되지 않았습니다.');
  }

  // ✅ OCR 텍스트 전처리 (입고서와 동일)
  let cleanedText = ocrText;

  Logger.log(`📊 원본 OCR 텍스트 길이: ${cleanedText.length}자`);

  // 1. 연속된 공백/줄바꿈 정리
  cleanedText = cleanedText.replace(/\s+/g, ' ').trim();

  // 2. 특수문자 제거 (한글, 숫자, 기본 구두점만 남김)
  cleanedText = cleanedText.replace(/[^\u3131-\u318E\uAC00-\uD7A3a-zA-Z0-9\s\.,:\-\/\(\)]/g, '');

  // 3. 텍스트가 너무 길면 제한
  const MAX_LENGTH = 4000;  // 처방전은 입고서보다 길 수 있음

  if (cleanedText.length > MAX_LENGTH) {
    Logger.log(`⚠️ OCR 텍스트가 ${cleanedText.length}자로 너무 깁니다. ${MAX_LENGTH}자로 제한합니다.`);
    cleanedText = cleanedText.substring(0, MAX_LENGTH);
  }

  Logger.log(`📊 정리된 OCR 텍스트 길이: ${cleanedText.length}자`);

  const prompt = `한의원 처방전 OCR 텍스트를 JSON으로 변환하세요.

아래 JSON 형식으로만 응답 (설명 없이 JSON만):
{
  "prescriptionNumber": "19357",
  "prescriptionDate": "2025-10-20",
  "prescriptionName": "사물탕가미",
  "cheops": 15,
  "patientName": "김경희",
  "chartNumber": "003379",
  "gender": "여",
  "age": 67,
  "birthDate": "1958-07-20",
  "doctorName": "주치형",
  "clinicName": "아현재한의원",
  "herbs": [
    {"name": "숙지황", "amountPerCheop": 5.6},
    {"name": "백작약", "amountPerCheop": 5.6}
  ]
}

정보 없으면 "", null 사용. 완전한 JSON 출력, 끝에 ] } 닫기

OCR 텍스트:
${cleanedText}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  
  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 8192,  // 약재가 많을 수 있으므로 8192로 증가
      topP: 0.8,
      topK: 40
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // ✅ 재시도 로직 (503 에러 대응)
  const MAX_RETRIES = 3;
  let lastError = null;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      if (attempt > 1) {
        const waitTime = attempt * 2000; // 2초, 4초, 6초
        Logger.log(`⏳ ${waitTime/1000}초 대기 후 재시도 (${attempt}/${MAX_RETRIES})...`);
        Utilities.sleep(waitTime);
      }

      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      Logger.log('=== Gemini API 응답 (처방전) ===');
      Logger.log('HTTP 상태: ' + responseCode);
      Logger.log('응답 길이: ' + responseText.length + ' 문자');
      if (attempt > 1) {
        Logger.log(`✅ 재시도 ${attempt}번째 성공`);
      }

      // ✅ 503 에러는 재시도
      if (responseCode === 503) {
        Logger.log('⚠️ 503 에러: Gemini API 과부하');
        lastError = new Error('Gemini API 서버 과부하 (503)');
        continue; // 재시도
      }

      if (responseCode !== 200) {
        Logger.log('❌ 전체 응답: ' + responseText);
        throw new Error(`Gemini API 오류 (HTTP ${responseCode}): ${responseText}`);
      }

      const result = JSON.parse(responseText);

      if (result.error) {
        // 503 에러 체크
        if (result.error.code === 503) {
          Logger.log('⚠️ 503 에러: ' + result.error.message);
          lastError = new Error(`Gemini API 서버 과부하: ${result.error.message}`);
          continue; // 재시도
        }
        throw new Error(`Gemini API 오류: ${result.error.message} (코드: ${result.error.code})`);
      }

      if (!result.candidates || !result.candidates[0]) {
        throw new Error('Gemini API 응답에 candidates가 없습니다.');
      }
    
    const candidate = result.candidates[0];

    // ✅ finishReason 확인 - MAX_TOKENS 처리 (입고서와 동일)
    const finishReason = candidate.finishReason || 'UNKNOWN';
    Logger.log(`📌 종료 이유: ${finishReason}`);

    if (finishReason === 'MAX_TOKENS') {
      Logger.log('⚠️ 토큰 제한으로 응답이 잘렸습니다.');

      if (!candidate.content || !candidate.content.parts || !candidate.content.parts[0] || !candidate.content.parts[0].text) {
        Logger.log('❌ MAX_TOKENS이지만 응답 내용이 없습니다.');
        throw new Error('Gemini 토큰 제한 초과: OCR 텍스트가 너무 복잡합니다. 이미지를 더 깔끔하게 찍어주세요.');
      }

      Logger.log('⚠️ 응답이 잘렸지만 일부 내용이 있습니다. 복구 시도...');
    }

    if (!candidate.content || !candidate.content.parts || !candidate.content.parts[0]) {
      Logger.log('❌ 응답 구조: ' + JSON.stringify(candidate));
      throw new Error('Gemini API 응답 구조가 올바르지 않습니다.');
    }

    let textContent = candidate.content.parts[0].text;
    Logger.log('추출된 텍스트 (첫 800자): ' + textContent.substring(0, 800));

    textContent = textContent.trim();
    textContent = textContent.replace(/```json\s*/gi, '');
    textContent = textContent.replace(/```\s*/g, '');
    textContent = textContent.trim();

    const jsonStart = textContent.indexOf('{');
    const jsonEnd = textContent.lastIndexOf('}');

    if (jsonStart === -1) {
      Logger.log('❌ JSON 찾기 실패. 전체 텍스트: ' + textContent);
      throw new Error('응답에서 JSON 형식을 찾을 수 없습니다.');
    }

    let jsonText;

    // ✅ JSON 복구 로직 (입고서와 유사)
    if (jsonEnd === -1 || jsonEnd < jsonStart) {
      Logger.log('⚠️ JSON이 불완전합니다. 자동 복구 시도...');

      jsonText = textContent.substring(jsonStart);

      // herbs 배열이 닫히지 않은 경우 처리
      const lastComma = jsonText.lastIndexOf(',');
      const lastCloseBrace = jsonText.lastIndexOf('}');

      if (lastCloseBrace !== -1 && lastComma > lastCloseBrace) {
        jsonText = jsonText.substring(0, lastCloseBrace + 1);
      }

      if (jsonText.includes('"herbs"') && jsonText.lastIndexOf(']') < jsonText.lastIndexOf('[')) {
        jsonText += '\n  ]\n}';
      } else if (!jsonText.endsWith('}')) {
        jsonText += '\n}';
      }

      Logger.log('✅ 복구된 JSON (처음 500자): ' + jsonText.substring(0, 500));
    } else {
      jsonText = textContent.substring(jsonStart, jsonEnd + 1);
    }

    Logger.log('추출된 JSON (길이: ' + jsonText.length + ')');
    
    try {
      const parsed = JSON.parse(jsonText);
      
      // 데이터 검증
      if (!parsed.herbs || !Array.isArray(parsed.herbs) || parsed.herbs.length === 0) {
        throw new Error('약재 항목이 없습니다.');
      }
      
      if (!parsed.patientName) {
        throw new Error('환자명이 없습니다.');
      }
      
      if (!parsed.cheops || parsed.cheops <= 0) {
        throw new Error('첩수가 올바르지 않습니다.');
      }
      
      // 기본값 설정
      parsed.prescriptionNumber = parsed.prescriptionNumber || '';
      parsed.prescriptionDate = parsed.prescriptionDate || new Date().toISOString().split('T')[0];
      parsed.prescriptionName = parsed.prescriptionName || '';
      parsed.chartNumber = parsed.chartNumber || '';
      parsed.gender = parsed.gender || '';
      parsed.age = parsed.age || null;
      parsed.birthDate = parsed.birthDate || '';
      parsed.doctorName = parsed.doctorName || '';
      parsed.clinicName = parsed.clinicName || '';
      
      // 총 용량 계산 추가
      parsed.herbs = parsed.herbs.map(herb => ({
        ...herb,
        totalAmount: herb.amountPerCheop * parsed.cheops
      }));
      
      // 약재 목록을 문자열로 변환 (처방입력 시트용)
      parsed.herbsList = parsed.herbs.map(h => h.name).join(', ');
      
      Logger.log('✅ 처방전 JSON 파싱 성공:');
      Logger.log(`  - 처방전번호: ${parsed.prescriptionNumber}`);
      Logger.log(`  - 처방일: ${parsed.prescriptionDate}`);
      Logger.log(`  - 환자: ${parsed.patientName} (${parsed.gender}, ${parsed.age}세)`);
      Logger.log(`  - 생년월일: ${parsed.birthDate}`);
      Logger.log(`  - 차트번호: ${parsed.chartNumber}`);
      Logger.log(`  - 처방: ${parsed.prescriptionName} (${parsed.cheops}첩)`);
      Logger.log(`  - 처방의: ${parsed.doctorName}`);
      Logger.log(`  - 약재: ${parsed.herbs.length}개`);

      return parsed;  // ✅ 성공 - 재시도 루프 탈출

    } catch (parseError) {
      Logger.log('❌ JSON 파싱 오류: ' + parseError.message);
      Logger.log('파싱 시도한 JSON 앞부분: ' + jsonText.substring(0, 500));
      throw new Error(`JSON 파싱 실패: ${parseError.message}`);
    }

    } catch (error) {
      // 503 에러는 재시도, 다른 에러는 즉시 throw
      if (error.message && error.message.includes('503')) {
        lastError = error;
        Logger.log(`⚠️ 시도 ${attempt}/${MAX_RETRIES} 실패: ${error.message}`);
        if (attempt < MAX_RETRIES) {
          continue; // 재시도
        }
      } else {
        // 503이 아닌 다른 에러는 즉시 throw
        Logger.log('❌ Gemini API 호출 오류 (재시도 불가): ' + error.message);
        throw error;
      }
    }
  }

  // 모든 재시도 실패
  Logger.log(`❌ ${MAX_RETRIES}번 재시도 모두 실패`);
  throw lastError || new Error('Gemini API 호출 실패');
}

/**
 * 처방전 데이터를 처방입력/처방상세 시트에 추가
 */
function addPrescriptionToSheet(parsedData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 처방입력 시트 처리
  let prescInputSheet = ss.getSheetByName('처방입력');
  
  if (!prescInputSheet) {
    prescInputSheet = ss.insertSheet('처방입력');
    
    const headers = [
      '처방일', '처방명', '차트번호', '환자명', '첩수', 
      '성별', '나이', '생년월일', '처방의', '약재목록(자동)', '처리상태'
    ];
    prescInputSheet.appendRow(headers);
    
    const headerRange = prescInputSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a73e8');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
  }
  
  // 처방입력 시트에 추가
  prescInputSheet.appendRow([
    parsedData.prescriptionDate,
    parsedData.prescriptionName,
    parsedData.chartNumber,
    parsedData.patientName,
    parsedData.cheops,
    parsedData.gender,
    parsedData.age,
    parsedData.birthDate,
    parsedData.doctorName,
    parsedData.herbsList,
    '대기중'
  ]);
  
  Logger.log(`✅ 처방입력 시트에 추가: ${parsedData.patientName} - ${parsedData.prescriptionName}`);
  
  // 2. 처방상세 시트 처리
  let prescDetailSheet = ss.getSheetByName('처방상세');
  
  if (!prescDetailSheet) {
    prescDetailSheet = ss.insertSheet('처방상세');
    
    const headers = [
      '처방전번호', '처방명', '처방일', '환자명', '차트번호', 
      '약재명', '용량(g/첩)', '첩수', '총수량(g)', '조제완료'
    ];
    prescDetailSheet.appendRow(headers);
    
    const headerRange = prescDetailSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a73e8');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
  }
  
  // 각 약재를 처방상세 시트에 추가
  parsedData.herbs.forEach(herb => {
    prescDetailSheet.appendRow([
      parsedData.prescriptionNumber,
      parsedData.prescriptionName,
      parsedData.prescriptionDate,
      parsedData.patientName,
      parsedData.chartNumber,
      herb.name,
      herb.amountPerCheop,
      parsedData.cheops,
      herb.totalAmount,
      false  // 조제완료 체크박스
    ]);
  });
  
  // 조제완료 체크박스 추가
  const lastRow = prescDetailSheet.getLastRow();
  const firstRow = lastRow - parsedData.herbs.length + 1;
  const checkboxRange = prescDetailSheet.getRange(firstRow, 10, parsedData.herbs.length, 1);
  checkboxRange.insertCheckboxes();
  checkboxRange.setHorizontalAlignment('center');
  
  Logger.log(`✅ 처방상세 시트에 ${parsedData.herbs.length}개 약재 추가`);
}

// ========================================
// 📤 처방 자동화 - PART 2: FIFO 자동 차감
// ========================================

/**
 * 처방상세 시트 편집 시 자동 조제 처리
 */
function onPrescriptionEdit_DISABLED(e) {
  try {
    if (!e || !e.source) {
      Logger.log('⚠️ 이 함수는 자동 트리거로만 실행됩니다.');
      return;
    }
    
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    
    Logger.log(`🔔 편집 감지: ${sheet.getName()}, ${range.getRow()}행, ${range.getColumn()}열`);
    
    // 처방상세 시트가 아니면 무시
    if (sheet.getName() !== '처방상세') {
      return;
    }
    
    // 10번째 컬럼(조제완료)이 아니면 무시
    if (range.getColumn() !== 10) {
      return;
    }
    
    // 체크박스가 true로 변경되었는지 확인
    if (range.getValue() !== true) {
      return;
    }
    
    const editedRow = range.getRow();
    
    // 헤더 행은 무시
    if (editedRow === 1) {
      return;
    }
    
    Logger.log(`✅ 조제 처리 시작: ${editedRow}행`);
    
    // 약재출고 처리 (함수 이름 수정!)
    try {
      processPrescriptionDispense(editedRow);  // ✅ 정확한 함수 이름
      Logger.log('✅ 조제 처리 완료');
      
    } catch (error) {
      Logger.log('❌ 조제 처리 오류: ' + error.message);
      Logger.log('상세:\n' + error.stack);
      
      // 체크 해제
      range.setValue(false);
      
      // 사용자 알림
      SpreadsheetApp.getActive().toast(
        error.message,
        '조제 처리 오류',
        10
      );
    }
    
  } catch (error) {
    Logger.log('❌ onPrescriptionEdit 전체 오류: ' + error.message);
  }
}

/**
 * 처방상세 한 행의 조제 처리 (FIFO 차감) - 원가 계산 추가
 */
function processPrescriptionDispense(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailSheet = ss.getSheetByName('처방상세');
  
  if (!detailSheet) {
    throw new Error('처방상세 시트가 없습니다.');
  }
  
  // 처방상세 시트에서 데이터 읽기
  const data = detailSheet.getRange(row, 1, 1, 10).getValues()[0];
  
  const prescriptionNumber = data[0] || '';
  const prescriptionName = data[1] || '';
  const prescriptionDate = data[2] || new Date();
  const patientName = data[3] || '';
  const chartNumber = data[4] || '';
  const herbName = data[5];
  const totalAmount = parseFloat(data[8]) || 0;
  
  Logger.log(`  조제 처리: ${herbName} ${totalAmount}g`);
  
  if (!herbName || totalAmount <= 0) {
    throw new Error('약재명 또는 수량이 올바르지 않습니다.');
  }
  
  // FIFO 할당 및 차감
  const batchInfo = allocateStockFIFO(herbName, totalAmount);
  
  // ✅ 이 약재의 원가 계산
  const herbCost = batchInfo.reduce((sum, batch) => sum + batch.금액, 0);
  Logger.log(`  ${herbName} 원가: ${herbCost}원`);
  
  // 약재출고 시트
  let dispenseSheet = ss.getSheetByName('약재출고');
  if (!dispenseSheet) {
    throw new Error('약재출고 시트가 없습니다.');
  }
  
  // FIFO상세추적 시트
  let fifoDetailSheet = ss.getSheetByName('FIFO상세추적');
  if (!fifoDetailSheet) {
    fifoDetailSheet = ss.insertSheet('FIFO상세추적');
    
    const headers = [
      '출고일', '처방전번호', '처방명', '환자명', '약재명',
      '입고번호', '입고일', '유통기한', '출고량(g)', 
      '단가(원/g)', '금액(원)', '공급처'
    ];
    fifoDetailSheet.appendRow(headers);
    
    const headerRange = fifoDetailSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#34a853');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
  }
  
  // 처방의 정보 가져오기
  let doctor = '';
  const prescriptionSheet = ss.getSheetByName('처방입력');
  if (prescriptionSheet) {
    const prescData = prescriptionSheet.getDataRange().getValues();
    for (let i = 1; i < prescData.length; i++) {
      if (prescData[i][0] === prescriptionNumber) {
        doctor = prescData[i][9] || '';
        break;
      }
    }
  }
  
  const batchSummary = batchInfo.map(b => `${b.입고번호}(${b.출고량}g)`).join(', ');
  const currentDate = new Date();
  
  // 약재출고 시트에 기록
  dispenseSheet.appendRow([
    currentDate,
    prescriptionNumber,
    herbName,
    totalAmount,
    doctor,
    patientName,
    chartNumber,
    batchSummary
  ]);
  
  // FIFO상세추적 시트에 기록
  batchInfo.forEach(batch => {
    fifoDetailSheet.appendRow([
      currentDate,
      prescriptionNumber,
      prescriptionName,
      patientName,
      herbName,
      batch.입고번호,
      batch.입고일,
      batch.유통기한,
      batch.출고량,
      batch.단가,
      batch.금액,
      batch.공급처
    ]);
  });
  
  // ✅ 출고 즉시 원가 누적 업데이트
  updatePrescriptionCostIncremental(prescriptionNumber, herbCost);
  
  // 처방상세에서 해당 행 삭제
  detailSheet.deleteRow(row);
  
  // 처방 완료 확인
  checkAndCompletePrescription(prescriptionNumber);
  
  Logger.log(`  ✅ ${herbName} ${totalAmount}g 출고 완료 (원가: ${herbCost}원)`);
}

/**
 * 처방 원가를 점진적으로 업데이트 (출고할 때마다 누적)
 */
function updatePrescriptionCostIncremental(prescriptionNumber, additionalCost) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prescSheet = ss.getSheetByName('처방입력');
  
  if (!prescSheet) {
    Logger.log('⚠️ 처방입력 시트를 찾을 수 없습니다.');
    return;
  }
  
  const data = prescSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === prescriptionNumber) { // A열: 처방전번호
      const row = i + 1;
      const currentCost = parseFloat(data[i][12]) || 0; // M열: 원가(원)
      const newCost = Math.round((currentCost + additionalCost) * 10) / 10;
      
      prescSheet.getRange(row, 13).setValue(newCost); // M열 업데이트
      
      Logger.log(`  ✅ 원가 누적: ${currentCost.toLocaleString()}원 → ${newCost.toLocaleString()}원 (+${additionalCost.toLocaleString()}원)`);
      return;
    }
  }
  
  Logger.log(`  ⚠️ 처방번호 ${prescriptionNumber}를 찾을 수 없습니다.`);
}

/**
 * 처방이 모두 완료되었는지 확인하고 완료 처리
 */
function checkAndCompletePrescription(prescriptionNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailSheet = ss.getSheetByName('처방상세');
  const prescSheet = ss.getSheetByName('처방입력');
  
  if (!detailSheet || !prescSheet) {
    Logger.log('⚠️ 시트를 찾을 수 없습니다.');
    return;
  }
  
  // 처방상세에 아직 남아있는 항목 확인
  const detailData = detailSheet.getDataRange().getValues();
  let hasRemaining = false;
  
  for (let i = 1; i < detailData.length; i++) {
    if (detailData[i][0] === prescriptionNumber) { // A열: 처방전번호
      hasRemaining = true;
      break;
    }
  }
  
  if (hasRemaining) {
    Logger.log(`  처방 ${prescriptionNumber}: 아직 미완료 항목 있음`);
    return;
  }
  
  // 모두 완료됨 - 처방입력 시트 업데이트
  Logger.log(`  ✅ 처방 ${prescriptionNumber}: 모든 약재 조제 완료!`);
  
  const prescData = prescSheet.getDataRange().getValues();
  
  for (let i = 1; i < prescData.length; i++) {
    if (prescData[i][0] === prescriptionNumber) { // A열: 처방전번호
      const row = i + 1;
      
      // 처리상태를 '완료'로 변경
      prescSheet.getRange(row, 12).setValue('완료'); // L열: 처리상태
      
      // 완료일시 기록
      prescSheet.getRange(row, 14).setValue(new Date()); // N열: 완료일시
      
      // ✅ 원가는 이미 누적되어 있음 - 최종 검증만
      const finalCost = parseFloat(prescData[i][12]) || 0;
      const calculatedCost = calculatePrescriptionCost(prescriptionNumber);
      
      if (Math.abs(finalCost - calculatedCost) > 1) {
        Logger.log(`  ⚠️ 원가 불일치 감지: 기록값 ${finalCost}원, 계산값 ${calculatedCost}원 - 재계산 적용`);
        prescSheet.getRange(row, 13).setValue(calculatedCost);
      } else {
        Logger.log(`  ✅ 원가 검증 완료: ${finalCost.toLocaleString()}원`);
      }
      
      // ✅ Slack 완료 알림
      try {
        const patientName = prescData[i][4]; // E열: 환자명
        const prescName = prescData[i][2]; // C열: 처방명
        const finalCostValue = prescSheet.getRange(row, 13).getValue();
        
        const message = `✅ *조제 완료*\n\n` +
          `• 처방전: ${prescriptionNumber}\n` +
          `• 환자: ${patientName}\n` +
          `• 처방: ${prescName}\n` +
          `• 원가: ${finalCostValue.toLocaleString()}원`;
        
        sendSlackNotification(message);
        Logger.log(`  ✅ Slack 완료 알림 발송`);
      } catch (error) {
        Logger.log(`  ⚠️ Slack 알림 실패: ${error.message}`);
      }
      
      break;
    }
  }
}

/**
 * FIFO 방식으로 재고 할당 및 차감 (트랜잭션 방식)
 */
function allocateStockFIFO(herbName, requiredAmount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const incomingSheet = ss.getSheetByName('약재입고');
  
  if (!incomingSheet) {
    Logger.log('⚠️ 약재입고 시트가 없습니다.');
    throw new Error('약재입고 시트가 없습니다.');
  }
  
  const data = incomingSheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    Logger.log(`⚠️ ${herbName}: 약재입고 시트에 데이터가 없습니다.`);
    throw new Error(`${herbName}의 입고 기록이 없습니다.`);
  }
  
  let allocated = [];
  let remaining = requiredAmount;
  
  // 유통기한 빠른 순으로 정렬
  let batches = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === herbName) {
      const rowNum = i + 1;
      const batchId = data[i][0];
      const incomingDate = data[i][1];
      const expiryDateValue = data[i][4];
      const remainingAmount = parseFloat(data[i][5]) || 0;
      const pricePerGram = parseFloat(data[i][6]) || 0;
      const supplier = data[i][7];
      
      let expiryDate;
      if (expiryDateValue && expiryDateValue instanceof Date) {
        expiryDate = expiryDateValue;
      } else if (expiryDateValue) {
        expiryDate = new Date(expiryDateValue);
      } else {
        expiryDate = new Date('2099-12-31');
      }
      
      if (remainingAmount > 0) {
        batches.push({
          rowNum: rowNum,
          batchId: batchId,
          incomingDate: incomingDate,
          expiryDate: expiryDate,
          available: remainingAmount,
          pricePerGram: pricePerGram,
          supplier: supplier
        });
      }
    }
  }
  
  if (batches.length === 0) {
    Logger.log(`⚠️ ${herbName}: 가용 재고가 없습니다.`);
    throw new Error(`${herbName}의 재고가 없습니다.`);
  }
  
  batches.sort((a, b) => a.expiryDate - b.expiryDate);
  
  Logger.log(`\n📦 ${herbName} FIFO 할당 시작`);
  Logger.log(`필요량: ${requiredAmount}g`);
  Logger.log(`가용 재고: ${batches.length}개 입고 건`);
  
  // ===== 1단계: 할당 가능 여부만 체크 (차감하지 않음!) =====
  let allocationPlan = [];
  let tempRemaining = requiredAmount;
  
  for (let batch of batches) {
    if (tempRemaining <= 0) break;
    
    const allocateAmount = Math.min(tempRemaining, batch.available);
    const allocatePrice = Math.round(allocateAmount * batch.pricePerGram * 10) / 10;
    
    allocationPlan.push({
      rowNum: batch.rowNum,
      batch: batch,
      allocateAmount: allocateAmount,
      newRemaining: Math.round((batch.available - allocateAmount) * 10) / 10,
      출고정보: {
        입고번호: batch.batchId,
        입고일: Utilities.formatDate(new Date(batch.incomingDate), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        유통기한: Utilities.formatDate(batch.expiryDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        출고량: allocateAmount,
        단가: batch.pricePerGram,
        금액: allocatePrice,
        공급처: batch.supplier
      }
    });
    
    tempRemaining -= allocateAmount;
  }
  
  // ===== 재고 부족 체크 =====
  if (tempRemaining > 0) {
    const currentStock = allocationPlan.reduce((sum, plan) => sum + plan.allocateAmount, 0);
    Logger.log(`❌ ${herbName} 재고 부족: 필요 ${requiredAmount}g, 가용 ${currentStock}g, 부족 ${tempRemaining}g`);
    
    // ❌ 여기서는 아무것도 차감하지 않음!
    throw new Error(`${herbName}의 재고가 ${tempRemaining}g 부족합니다. (필요: ${requiredAmount}g, 가용: ${currentStock}g)`);
  }
  
  // ===== 2단계: 할당 가능하면 실제로 차감 =====
  Logger.log(`✅ 재고 충분, 실제 차감 시작`);
  
  for (let plan of allocationPlan) {
    incomingSheet.getRange(plan.rowNum, 6).setValue(plan.newRemaining);
    
    Logger.log(`✅ ${plan.batch.batchId}: ${plan.allocateAmount}g 출고, 잔량 ${plan.batch.available}g → ${plan.newRemaining}g`);
    
    allocated.push(plan.출고정보);
  }
  
  Logger.log(`✅ FIFO 할당 완료: ${allocated.length}개 입고분 사용\n`);
  
  return allocated;
}

// ========================================
// 📊 재고 관리
// ========================================

/**
 * 약재마스터 시트 현재 재고 자동 업데이트
 */
function updateCurrentStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('약재마스터');
  const incomingSheet = ss.getSheetByName('약재입고');
  const dispenseSheet = ss.getSheetByName('약재출고');
  
  if (!masterSheet) {
    Logger.log('❌ 약재마스터 시트가 없습니다.');
    return;
  }
  
  const masterData = masterSheet.getDataRange().getValues();
  
  // 약재입고 데이터
  let incomingData = [];
  if (incomingSheet) {
    incomingData = incomingSheet.getDataRange().getValues();
  } else {
    Logger.log('⚠️ 약재입고 시트가 없습니다.');
  }
  
  // 약재출고 데이터
  let dispenseData = [];
  if (dispenseSheet) {
    dispenseData = dispenseSheet.getDataRange().getValues();
  } else {
    Logger.log('⚠️ 약재출고 시트가 없습니다.');
  }
  
  Logger.log('=== 재고 업데이트 시작 ===');
  
  for (let i = 1; i < masterData.length; i++) {
    const herbName = masterData[i][0];
    
    if (!herbName || herbName.trim() === '') {
      continue;
    }
    
    // 총 입고량 및 공급처 수집
    let totalIncoming = 0;
    let suppliers = new Set();
    
    for (let j = 1; j < incomingData.length; j++) {
      if (incomingData[j][2] === herbName) {  // C열: 약재명
        totalIncoming += parseFloat(incomingData[j][3]) || 0;  // D열: 수량
        
        const supplier = incomingData[j][7];  // H열: 공급처
        if (supplier && supplier.trim() !== '') {
          suppliers.add(supplier.trim());
        }
      }
    }
    
    // 총 출고량
    let totalDispensed = 0;
    for (let k = 1; k < dispenseData.length; k++) {
      if (dispenseData[k][2] === herbName) {  // C열: 약재명
        totalDispensed += parseFloat(dispenseData[k][3]) || 0;  // D열: 출고량
      }
    }
    
    // 현재 재고 = 입고 - 출고
    const currentStock = totalIncoming - totalDispensed;
    
    // C열: 현재재고 업데이트
    masterSheet.getRange(i + 1, 3).setValue(currentStock);
    
    // G열: 가장 이른 유통기한 업데이트
    const nearestExpiry = getNearestExpiryDate(herbName);
    if (nearestExpiry) {
      masterSheet.getRange(i + 1, 7).setValue(nearestExpiry);
    } else {
      masterSheet.getRange(i + 1, 7).setValue('');
    }
    
    // H열: 공급처 자동 업데이트
    if (suppliers.size > 0) {
      const supplierList = Array.from(suppliers).join(', ');
      masterSheet.getRange(i + 1, 8).setValue(supplierList);
    }
    
    Logger.log(`${herbName}: 입고 ${totalIncoming}g - 출고 ${totalDispensed}g = 재고 ${currentStock}g`);
  }
  
  Logger.log('✅ 약재마스터 현재 재고 업데이트 완료');
}

/**
 * 가장 빠른 유통기한 가져오기 (잔량이 있는 것만)
 */
function getNearestExpiryDate(herbName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const incomingSheet = ss.getSheetByName('약재입고');
  
  if (!incomingSheet) {
    return null;
  }
  
  const data = incomingSheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return null;
  }
  
  let nearestDate = null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === herbName) {  // C열: 약재명
      const expiryDateValue = data[i][4];  // E열: 유통기한
      const remainingAmount = parseFloat(data[i][5]) || 0;  // F열: 잔량
      
      // 유통기한 파싱
      let expiryDate;
      if (expiryDateValue && expiryDateValue instanceof Date) {
        expiryDate = expiryDateValue;
      } else if (expiryDateValue) {
        try {
          expiryDate = new Date(expiryDateValue);
        } catch (e) {
          continue;
        }
      } else {
        continue;
      }
      
      // 잔량이 있는 입고분만 확인
      if (remainingAmount > 0) {
        if (!nearestDate || expiryDate < nearestDate) {
          nearestDate = expiryDate;
        }
      }
    }
  }
  
  return nearestDate;
}

/**
 * 평균 일일 소비량 계산 (120일 기준)
 */
function calculateAverageDailyUsage(herbName, days = 120) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dispenseSheet = ss.getSheetByName('약재출고');
  
  if (!dispenseSheet) {
    return 0;
  }
  
  const data = dispenseSheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return 0;
  }
  
  const today = new Date();
  const startDate = new Date(today.getTime() - (days * 24 * 60 * 60 * 1000));
  
  let totalUsage = 0;
  
  for (let i = 1; i < data.length; i++) {
    const dateValue = data[i][0];  // A열: 출고일
    const name = data[i][2];  // C열: 약재명
    const amount = parseFloat(data[i][3]) || 0;  // D열: 출고량
    
    let date;
    if (dateValue instanceof Date) {
      date = dateValue;
    } else {
      date = new Date(dateValue);
    }
    
    if (name === herbName && date >= startDate && date <= today) {
      totalUsage += amount;
    }
  }
  
  const actualDays = Math.max(1, Math.floor((today - startDate) / (1000 * 60 * 60 * 24)));
  return totalUsage / actualDays;
}

/**
 * 최소재고량 AI 자동 계산 (120일 기준)
 */
function autoUpdateMinimumStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('약재마스터');
  
  if (!masterSheet) {
    Logger.log('❌ 약재마스터 시트가 없습니다.');
    return;
  }
  
  const data = masterSheet.getDataRange().getValues();
  
  Logger.log('=== 최소재고량 자동 업데이트 시작 ===');
  
  for (let i = 1; i < data.length; i++) {
    const herbName = data[i][0];
    
    if (!herbName || herbName.trim() === '') {
      continue;
    }
    
    // 평균 일일 소비량 계산
    const avgDailyUsage = calculateAverageDailyUsage(herbName, 120);
    
    // F열에 평균일일소비량 업데이트
    masterSheet.getRange(i + 1, 6).setValue(Math.round(avgDailyUsage * 10) / 10);
    
    // 안전재고 계산 (리드타임 7일 + 안전계수 1.2배)
    const safetyStock = avgDailyUsage * 7 * 1.2;
    const minimumStock = Math.round(safetyStock);
    
    // D열에 최소재고량 업데이트
    masterSheet.getRange(i + 1, 4).setValue(minimumStock);
    
    Logger.log(`${herbName}: 평균 ${Math.round(avgDailyUsage)}g/일 → 최소재고 ${minimumStock}g`);
  }
  
  Logger.log('✅ 최소재고량 자동 업데이트 완료');
}

/**
 * 유통기한 임박 약재 확인 (30일 이내)
 */
function checkExpiringHerbs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('약재마스터');
  
  if (!masterSheet) {
    return;
  }
  
  const data = masterSheet.getDataRange().getValues();
  const today = new Date();
  const threshold = new Date(today.getTime() + (30 * 24 * 60 * 60 * 1000));
  
  let expiringHerbs = [];
  
  for (let i = 1; i < data.length; i++) {
    const herbName = data[i][0];
    const currentStock = data[i][2];
    const expiryDateValue = data[i][6];  // G열: 가장이른_유통기한
    
    if (!herbName || !expiryDateValue || currentStock <= 0) {
      continue;
    }
    
    let expiryDate;
    if (expiryDateValue instanceof Date) {
      expiryDate = expiryDateValue;
    } else {
      try {
        expiryDate = new Date(expiryDateValue);
      } catch (e) {
        continue;
      }
    }
    
    if (expiryDate <= threshold) {
      const daysLeft = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));
      expiringHerbs.push({
        herbName: herbName,
        expiryDate: expiryDate,
        daysLeft: daysLeft,
        currentStock: currentStock
      });
      
      // 셀 색상 변경 (빨간색)
      masterSheet.getRange(i + 1, 7).setBackground('#f4cccc');
    }
  }
  
  if (expiringHerbs.length > 0) {
    Logger.log(`⚠️ 유통기한 임박 약재: ${expiringHerbs.length}개`);
    sendExpiringHerbsAlert(expiringHerbs);
  }
}

// ========================================
// 🔔 슬랙 알림
// ========================================

function sendOCRCompletedSlack(data, count) {
  const webhookUrl = getConfig('SLACK_WEBHOOK_URL');
  if (!webhookUrl) return;
  
  const itemsList = data.items.slice(0, 3).map(item => {
    const bagInfo = item.bagSize ? `${item.bagSize}g × ${item.quantity}봉` : `${item.quantity}봉`;
    const priceInfo = item.totalPrice && item.bagSize && item.quantity ? 
      ` (${Math.round((item.totalPrice / (item.bagSize * item.quantity)) * 10) / 10}원/g)` : '';
    return `• ${item.herbName}: ${bagInfo}${priceInfo}`;
  }).join('\n');
  
  const moreItems = data.items.length > 3 ? `\n... 외 ${data.items.length - 3}개` : '';
  
  const payload = {
    text: `📸 입고서 OCR 완료 (${count}건)`,
    blocks: [{
      "type": "section",
      "text": {
        "type": "mrkdwn",
        "text": `*📸 입고서 OCR 완료*\n\n${itemsList}${moreItems}\n\n⚠️ *임시입고 시트*에서 유통기한 입력 후 처리완료 체크!`
      }
    }]
  };
  
  sendSlackMessage(webhookUrl, payload);
}

function sendIncomingCompletedSlack(data) {
  const webhookUrl = getConfig('SLACK_WEBHOOK_URL');
  if (!webhookUrl) return;
  
  const payload = {
    text: `✅ 입고 완료: ${data.herbName}`,
    blocks: [{
      "type": "section",
      "text": {
        "type": "mrkdwn",
        "text": `*✅ 약재 입고 완료 (✨ FIFO 원가 계산 준비)*\n\n*약재명:* ${data.herbName}\n*수량:* ${data.quantity}봉 × ${data.bagSize}g = ${data.totalAmount}g\n*g당 단가:* ${data.pricePerGram}원/g\n\n📦 처방 시 실제 구매 가격으로 정확한 원가 계산됩니다!`
      }
    }]
  };
  
  sendSlackMessage(webhookUrl, payload);
}

function sendPrescriptionProcessedSlack(data) {
  const webhookUrl = getConfig('SLACK_WEBHOOK_URL');
  if (!webhookUrl) return;
  
  const herbsList = data.herbs.slice(0, 5).map(herb => {
    return `• ${herb.name}: ${herb.totalAmount}g`;
  }).join('\n');
  
  const moreHerbs = data.herbs.length > 5 ? `\n... 외 ${data.herbs.length - 5}개` : '';
  
  const payload = {
    text: `📋 처방 자동 입력 완료: ${data.patientName}`,
    blocks: [{
      "type": "section",
      "text": {
        "type": "mrkdwn",
        "text": `*📋 처방 자동 입력 완료*\n\n*환자:* ${data.patientName} (${data.chartNumber})\n*처방명:* ${data.prescriptionName}\n*첩수:* ${data.cheops}첩\n\n${herbsList}${moreHerbs}\n\n⚠️ 조제 완료 후 *처방상세 시트*에서 조제완료 체크!`
      }
    }]
  };
  
  sendSlackMessage(webhookUrl, payload);
}

function sendLowStockAlert(herbName, shortageAmount) {
  const webhookUrl = getConfig('SLACK_WEBHOOK_URL');
  if (!webhookUrl) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('약재마스터');
  
  if (!masterSheet) return;
  
  const data = masterSheet.getDataRange().getValues();
  let currentStock = 0;
  let minimumStock = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === herbName) {
      currentStock = data[i][2];
      minimumStock = data[i][3];
      break;
    }
  }
  
  const payload = {
    text: `⚠️ 재고 부족: ${herbName}`,
    blocks: [{
      "type": "section",
      "text": {
        "type": "mrkdwn",
        "text": `*⚠️ 재고 부족 경고*\n\n*약재명:* ${herbName}\n*현재 재고:* ${currentStock}g\n*최소 재고:* ${minimumStock}g\n*부족량:* ${shortageAmount}g\n\n🚨 조제 진행 불가! 긴급 발주가 필요합니다.`
      }
    }]
  };
  
  sendSlackMessage(webhookUrl, payload);
}

function sendExpiringHerbsAlert(expiringHerbs) {
  const webhookUrl = getConfig('SLACK_WEBHOOK_URL');
  if (!webhookUrl) return;
  
  const herbsList = expiringHerbs.slice(0, 5).map(herb => {
    return `• ${herb.herbName}: ${herb.daysLeft}일 남음 (${herb.currentStock}g)`;
  }).join('\n');
  
  const moreHerbs = expiringHerbs.length > 5 ? `\n... 외 ${expiringHerbs.length - 5}개` : '';
  
  const payload = {
    text: `🚨 유통기한 임박: ${expiringHerbs.length}개`,
    blocks: [{
      "type": "section",
      "text": {
        "type": "mrkdwn",
        "text": `*🚨 유통기한 임박 (30일 이내)*\n\n${herbsList}${moreHerbs}\n\n⚠️ 조속히 사용하세요!`
      }
    }]
  };
  
  sendSlackMessage(webhookUrl, payload);
}

// ========================================
// 🔧 트리거 설정
// ========================================

/**
 * 모든 트리거 한 번에 설정
 */
function setupAllTriggers() {
  // 기존 트리거 삭제
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  Logger.log('기존 트리거 삭제 완료');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 입고서 OCR (5분마다)
  ScriptApp.newTrigger('processIncomingImagesOCR')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('✅ processIncomingImagesOCR 트리거 생성');
  
  // 2. 처방전 OCR (5분마다)
  ScriptApp.newTrigger('processPrescriptionImages')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('✅ processPrescriptionImages 트리거 생성');
  
  // 3. 재고 자동 업데이트 (1시간마다)
  ScriptApp.newTrigger('updateCurrentStock')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('✅ updateCurrentStock 트리거 생성');
  
  // 4. 유통기한 확인 (매일 오전 9시)
  ScriptApp.newTrigger('checkExpiringHerbs')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
  Logger.log('✅ checkExpiringHerbs 트리거 생성');
  
  // 5. 최소재고량 자동 계산 (매주 월요일 오전 10시)
  ScriptApp.newTrigger('autoUpdateMinimumStock')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .create();
  Logger.log('✅ autoUpdateMinimumStock 트리거 생성');
  
  // 6. 임시입고 편집 트리거
  ScriptApp.newTrigger('onTempIncomingEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  Logger.log('✅ onTempIncomingEdit 트리거 생성');
  
  // 7. 처방상세 편집 트리거 ⭐ 중요!
  ScriptApp.newTrigger('onPrescriptionEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  Logger.log('✅ onPrescriptionEdit 트리거 생성');
  
  Logger.log('\n✅✅✅ 모든 트리거 설정 완료!');
  Browser.msgBox('완료', '모든 트리거가 설정되었습니다!', Browser.Buttons.OK);
}

// ========================================
// 🧪 테스트 및 유틸리티
// ========================================

/**
 * 시스템 테스트
 */
function testSystem() {
  Logger.log('=== 약재관리 자동화 시스템 테스트 ===\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log('1. 시트 확인');
  const sheets = ['설정', '약재마스터', '임시입고', '약재입고', '처방입력', '처방상세', '약재출고'];
  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    Logger.log(`${sheetName}: ${sheet ? '✅' : '❌'}`);
  });
  
  Logger.log('\n2. 설정 확인');
  const configs = ['GEMINI_API_KEY', 'VISION_API_KEY', 'SLACK_WEBHOOK_URL', '입고서_폴더_ID', '처방전_폴더_ID'];
  configs.forEach(key => {
    const value = getConfig(key);
    Logger.log(`${key}: ${value ? '✅' : '❌'}`);
  });
  
  Logger.log('\n3. 트리거 확인');
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log(`설정된 트리거 수: ${triggers.length}`);
  triggers.forEach(trigger => {
    Logger.log(`- ${trigger.getHandlerFunction()}`);
  });
  
  Logger.log('\n✨ v8.1: Vision API + Gemini API + FIFO 선입선출 통합 시스템');
  Logger.log('=== 테스트 완료 ===');
}

/**
 * 체크된 처방 수동 처리 (확인 후 처리)
 */
function processCheckedNow() {
  Logger.log('=== 체크된 처방 확인 시작 ===\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('처방상세');
  
  if (!sheet) {
    Browser.msgBox('오류', '처방상세 시트를 찾을 수 없습니다.', Browser.Buttons.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    Browser.msgBox('알림', '처방상세 시트에 데이터가 없습니다.', Browser.Buttons.OK);
    return;
  }
  
  // ===== 1단계: 체크된 항목 수집 =====
  let checkedItems = [];
  
  for (let row = 2; row <= lastRow; row++) {
    const isChecked = sheet.getRange(row, 10).getValue();
    
    if (isChecked === true) {
      const prescriptionNumber = sheet.getRange(row, 1).getValue();
      const prescriptionName = sheet.getRange(row, 2).getValue();
      const patientName = sheet.getRange(row, 4).getValue();
      const herbName = sheet.getRange(row, 6).getValue();
      const amount = sheet.getRange(row, 9).getValue();
      
      checkedItems.push({
        row: row,
        prescriptionNumber: prescriptionNumber,
        prescriptionName: prescriptionName,
        patientName: patientName,
        herbName: herbName,
        amount: amount
      });
    }
  }
  
  if (checkedItems.length === 0) {
    Browser.msgBox('알림', '체크된 항목이 없습니다.', Browser.Buttons.OK);
    return;
  }
  
  Logger.log(`체크된 항목: ${checkedItems.length}개`);
  
  // ===== 2단계: 재고 확인 =====
  let stockCheckResults = [];
  let allAvailable = true;
  
  for (let item of checkedItems) {
    try {
      // 재고만 확인 (차감하지 않음)
      const stockCheck = checkStockAvailability(item.herbName, item.amount);
      stockCheckResults.push({
        item: item,
        available: true,
        message: `✅ ${item.herbName} ${item.amount}g (재고: ${stockCheck.totalAvailable}g)`
      });
    } catch (error) {
      allAvailable = false;
      stockCheckResults.push({
        item: item,
        available: false,
        message: `❌ ${item.herbName} ${item.amount}g (${error.message})`
      });
    }
  }
  
  // ===== 3단계: 사용자 확인 =====
  const ui = SpreadsheetApp.getUi();
  let confirmMessage = `처리할 항목: ${checkedItems.length}개\n\n`;
  
  if (allAvailable) {
    confirmMessage += '✅ 모든 약재 재고 충분\n\n';
    stockCheckResults.forEach(result => {
      confirmMessage += result.message + '\n';
    });
    confirmMessage += '\n처리하시겠습니까?';
    
    const response = ui.alert(
      '조제 처리 확인',
      confirmMessage,
      ui.ButtonSet.YES_NO
    );
    
    Logger.log(`사용자 응답 (모든 재고 충분): ${response}`);
    
    if (response !== ui.Button.YES) {
      Logger.log('사용자가 처리를 취소했습니다.');
      return;
    }
    
  } else {
    // 재고 부족 항목 있음
    confirmMessage += '⚠️ 일부 약재 재고 부족\n\n';
    stockCheckResults.forEach(result => {
      confirmMessage += result.message + '\n';
    });
    confirmMessage += '\n✅ 표시된 항목만 처리하시겠습니까?';
    
    const response = ui.alert(
      '재고 부족 항목 있음',
      confirmMessage,
      ui.ButtonSet.YES_NO
    );
    
    Logger.log(`사용자 응답 (재고 부족 항목 있음): ${response}`);
    
    if (response !== ui.Button.YES) {
      Logger.log('사용자가 처리를 취소했습니다.');
      return;
    }
  }
  
  Logger.log('사용자가 처리를 확인했습니다. 처리 시작...\n');
  
  // ===== 4단계: 실제 처리 =====
  Logger.log('===== 실제 처리 시작 =====');
  Logger.log(`처리할 항목 수: ${stockCheckResults.length}`);
  
  let successCount = 0;
  let errorCount = 0;
  let errorMessages = [];
  let processedHerbs = new Set(); // ✅ 처리된 약재 목록
  
  // 뒤에서부터 처리 (행 삭제 대비)
  for (let i = stockCheckResults.length - 1; i >= 0; i--) {
    const result = stockCheckResults[i];
    
    Logger.log(`\n[${i}] 처리 시작: ${result.item.herbName} ${result.item.amount}g, 행번호: ${result.item.row}`);
    
    if (!result.available) {
      // 재고 부족 처리...
      errorCount++;
      errorMessages.push(`${result.item.herbName}: 재고 부족`);
      continue;
    }
    
    try {
      Logger.log(`  처리 시작: processPrescriptionDispense(${result.item.row})`);
      processPrescriptionDispense(result.item.row);
      successCount++;
      processedHerbs.add(result.item.herbName); // ✅ 처리된 약재 기록
      Logger.log(`  ✅ 처리 성공`);
      
    } catch (error) {
      Logger.log(`  ❌ 처리 실패: ${error.message}`);
      errorCount++;
      errorMessages.push(`${result.item.herbName}: ${error.message}`);
      
      // 체크박스 해제...
    }
  }
  
  Logger.log(`\n===== 처리 완료 =====`);
  Logger.log(`✅ 성공: ${successCount}개`);
  Logger.log(`❌ 실패: ${errorCount}개`);
  
  // ✅ 처리된 약재들의 마스터 재고 일괄 업데이트
  if (processedHerbs.size > 0) {
    Logger.log(`\n===== 약재마스터 재고 업데이트 =====`);
    processedHerbs.forEach(herbName => {
      updateSingleHerbStock(herbName);
    });
    Logger.log(`✅ ${processedHerbs.size}개 약재 재고 업데이트 완료`);
  }
  
  // ===== 5단계: 결과 알림 =====
  let resultMessage = `조제 처리 완료\n\n✅ 성공: ${successCount}개\n❌ 실패: ${errorCount}개`;
  
  if (errorMessages.length > 0) {
    resultMessage += '\n\n실패 내역:\n' + errorMessages.join('\n');
  }
  
  Browser.msgBox('처리 완료', resultMessage, Browser.Buttons.OK);
}

/**
 * 약재입고 시트에서 입고번호 없는 행 찾기
 */
function findMissingIncomingNumbers() {
  Logger.log('=== 입고번호 누락 확인 ===\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const incomingSheet = ss.getSheetByName('약재입고');
  
  if (!incomingSheet) {
    Logger.log('❌ 약재입고 시트 없음');
    return;
  }
  
  const data = incomingSheet.getDataRange().getValues();
  let problemRows = [];
  
  for (let i = 1; i < data.length; i++) {
    const incomingNumber = data[i][0];  // A열: 입고번호
    const herbName = data[i][2];        // C열: 약재명
    const remaining = data[i][5];       // F열: 잔량
    
    // 입고번호가 없는데 잔량이 있는 경우
    if (!incomingNumber && remaining > 0) {
      Logger.log(`⚠️ ${i+1}행: 입고번호 없음 - ${herbName} (잔량: ${remaining}g)`);
      problemRows.push({
        row: i + 1,
        herbName: herbName,
        remaining: remaining
      });
    }
  }
  
  if (problemRows.length === 0) {
    Logger.log('✅ 모든 입고 행에 입고번호가 있습니다.');
  } else {
    Logger.log(`\n❌ 입고번호 없는 행: ${problemRows.length}개`);
    Logger.log('\n해결 방법:');
    Logger.log('1. 약재입고 시트로 이동');
    Logger.log('2. 해당 행들의 입고번호(A열)를 채워주세요');
    Logger.log('   예: IN20251025-001, IN20251025-002 등');
  }
  
  Logger.log('\n=== 확인 완료 ===');
  
  return problemRows;
}

/**
 * 입고번호 없는 행에 자동으로 번호 부여
 */
function autoAssignIncomingNumbers() {
  Logger.log('=== 자동 입고번호 부여 시작 ===\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const incomingSheet = ss.getSheetByName('약재입고');
  
  if (!incomingSheet) {
    Logger.log('❌ 약재입고 시트 없음');
    return;
  }
  
  const data = incomingSheet.getDataRange().getValues();
  let assignedCount = 0;
  
  // 오늘 날짜로 시작하는 입고번호 중 가장 큰 번호 찾기
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  let maxSeq = 0;
  
  for (let i = 1; i < data.length; i++) {
    const incomingNumber = data[i][0];
    
    if (incomingNumber && incomingNumber.startsWith('IN' + today)) {
      const seqStr = incomingNumber.split('-')[1];
      const seq = parseInt(seqStr) || 0;
      if (seq > maxSeq) {
        maxSeq = seq;
      }
    }
  }
  
  Logger.log(`오늘 날짜(${today})의 최대 번호: ${maxSeq}`);
  
  // 입고번호 없는 행에 부여
  for (let i = 1; i < data.length; i++) {
    const incomingNumber = data[i][0];
    const herbName = data[i][2];
    const remaining = data[i][5];
    
    // 입고번호가 없고 약재명이 있는 경우
    if (!incomingNumber && herbName) {
      maxSeq++;
      const newNumber = `IN${today}-${String(maxSeq).padStart(3, '0')}`;
      
      incomingSheet.getRange(i + 1, 1).setValue(newNumber);
      assignedCount++;
      
      Logger.log(`✅ ${i+1}행: ${herbName} → ${newNumber}`);
    }
  }
  
  Logger.log(`\n=== 완료: ${assignedCount}개 행에 입고번호 부여 ===`);
  
  Browser.msgBox(
    '완료',
    `${assignedCount}개 행에 입고번호가 자동으로 부여되었습니다.`,
    Browser.Buttons.OK
  );
}

/**
 * 처방입력 시트에 원가 컬럼 추가
 */
function addCostColumnToPrescriptionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prescSheet = ss.getSheetByName('처방입력');
  
  if (!prescSheet) {
    Logger.log('❌ 처방입력 시트 없음');
    return;
  }
  
  const headers = prescSheet.getRange(1, 1, 1, prescSheet.getLastColumn()).getValues()[0];
  
  Logger.log('현재 헤더: ' + headers.join(', '));
  
  // 원가(원) 컬럼이 이미 있는지 확인
  if (headers.includes('원가(원)')) {
    Logger.log('✅ 원가(원) 컬럼이 이미 있습니다.');
    return;
  }
  
  // 처리상태 다음에 원가(원), 완료일시 컬럼 추가
  const lastCol = prescSheet.getLastColumn();
  
  prescSheet.getRange(1, lastCol + 1).setValue('원가(원)');
  prescSheet.getRange(1, lastCol + 2).setValue('완료일시');
  
  // 헤더 스타일
  const newHeaderRange = prescSheet.getRange(1, lastCol + 1, 1, 2);
  newHeaderRange.setBackground('#1a73e8');
  newHeaderRange.setFontColor('white');
  newHeaderRange.setFontWeight('bold');
  
  Logger.log('✅ 원가(원), 완료일시 컬럼 추가 완료');
  Browser.msgBox('완료', '원가(원), 완료일시 컬럼이 추가되었습니다.', Browser.Buttons.OK);
}

/**
 * 처방전번호로 원가 계산
 */
function calculatePrescriptionCost(prescriptionNumber) {
  if (!prescriptionNumber) {
    return 0;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fifoSheet = ss.getSheetByName('FIFO상세추적');
  
  if (!fifoSheet) {
    Logger.log('⚠️ FIFO상세추적 시트 없음');
    return 0;
  }
  
  const data = fifoSheet.getDataRange().getValues();
  let totalCost = 0;
  
  // 처방전번호가 일치하는 행의 금액 합산
  for (let i = 1; i < data.length; i++) {
    const prescNum = data[i][1];  // 2열: 처방전번호
    const amount = parseFloat(data[i][10]) || 0;  // 11열: 금액(원)
    
    if (prescNum === prescriptionNumber) {
      totalCost += amount;
    }
  }
  
  return Math.round(totalCost);
}

/**
 * 모든 처방의 원가 업데이트
 */
function updateAllPrescriptionCosts() {
  Logger.log('=== 전체 처방 원가 업데이트 시작 ===\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prescSheet = ss.getSheetByName('처방입력');
  
  if (!prescSheet) {
    Logger.log('❌ 처방입력 시트 없음');
    return;
  }
  
  const headers = prescSheet.getRange(1, 1, 1, prescSheet.getLastColumn()).getValues()[0];
  const costColIndex = headers.indexOf('원가(원)') + 1;
  const statusColIndex = headers.indexOf('처리상태') + 1;
  
  if (costColIndex === 0) {
    Logger.log('❌ 원가(원) 컬럼이 없습니다. addCostColumnToPrescriptionSheet()를 먼저 실행하세요.');
    Browser.msgBox('오류', '원가(원) 컬럼이 없습니다.\naddCostColumnToPrescriptionSheet() 함수를 먼저 실행하세요.', Browser.Buttons.OK);
    return;
  }
  
  const lastRow = prescSheet.getLastRow();
  
  if (lastRow <= 1) {
    Logger.log('⚠️ 데이터 없음');
    return;
  }
  
  const data = prescSheet.getRange(2, 1, lastRow - 1, prescSheet.getLastColumn()).getValues();
  let updatedCount = 0;
  
  for (let i = 0; i < data.length; i++) {
    const row = i + 2;
    const prescriptionNumber = data[i][0];  // 처방전번호 (첫 번째 컬럼)
    const status = data[i][statusColIndex - 1];  // 처리상태
    
    // 완료된 처방만 원가 계산
    if (status === '완료' || status === '조제완료') {
      const cost = calculatePrescriptionCost(prescriptionNumber);
      
      if (cost > 0) {
        prescSheet.getRange(row, costColIndex).setValue(cost);
        updatedCount++;
        Logger.log(`✅ ${row}행: ${prescriptionNumber} → ${cost}원`);
      }
    }
  }
  
  Logger.log(`\n=== 업데이트 완료: ${updatedCount}개 처방 ===`);
  Browser.msgBox('완료', `${updatedCount}개 처방의 원가가 업데이트되었습니다.`, Browser.Buttons.OK);
}

/**
 * 체크된 모든 처방을 한 번에 조제 처리
 */
function processAllCheckedPrescriptions() {
  Logger.log('=== 체크된 모든 처방 일괄 처리 ===\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prescDetailSheet = ss.getSheetByName('처방상세');
  
  if (!prescDetailSheet) {
    Logger.log('❌ 처방상세 시트 없음');
    Browser.msgBox('오류', '처방상세 시트가 없습니다.', Browser.Buttons.OK);
    return;
  }
  
  const lastRow = prescDetailSheet.getLastRow();
  
  if (lastRow <= 1) {
    Logger.log('⚠️ 처방상세 시트에 데이터가 없습니다.');
    Browser.msgBox('알림', '처방상세 시트에 데이터가 없습니다.', Browser.Buttons.OK);
    return;
  }
  
  const data = prescDetailSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  
  let processedCount = 0;
  let errorCount = 0;
  const errors = [];
  
  // 뒤에서부터 처리 (행 삭제로 인한 인덱스 변경 방지)
  for (let i = data.length - 1; i >= 0; i--) {
    const rowIndex = i + 2;  // 실제 시트 행 번호
    const row = data[i];
    const isChecked = row[9];  // 10번째 컬럼 (조제완료)
    
    if (isChecked === true) {
      Logger.log(`\n📌 ${rowIndex}행 처리 중:`);
      Logger.log(`  처방: ${row[1]}`);
      Logger.log(`  환자: ${row[3]}`);
      Logger.log(`  약재: ${row[5]} ${row[8]}g`);
      
      try {
        processPrescriptionDispense(rowIndex);
        processedCount++;
        Logger.log('  ✅ 조제 완료');
        
      } catch (error) {
        errorCount++;
        const errorMsg = `${row[5]} (${rowIndex}행): ${error.message}`;
        errors.push(errorMsg);
        Logger.log('  ❌ 오류: ' + error.message);
      }
    }
  }
  
  Logger.log(`\n=== 처리 완료 ===`);
  Logger.log(`✅ 성공: ${processedCount}개`);
  Logger.log(`❌ 실패: ${errorCount}개`);
  
  // 결과 메시지
  let resultMsg = `조제 처리가 완료되었습니다.\n\n`;
  resultMsg += `✅ 성공: ${processedCount}개\n`;
  
  if (errorCount > 0) {
    resultMsg += `❌ 실패: ${errorCount}개\n\n`;
    resultMsg += `오류 내역:\n`;
    errors.forEach(err => {
      resultMsg += `- ${err}\n`;
    });
  }
  
  Browser.msgBox('조제 처리 완료', resultMsg, Browser.Buttons.OK);
  
  if (processedCount === 0 && errorCount === 0) {
    Logger.log('\n💡 체크된 행이 없습니다.');
  }
}

/**
 * 슬랙 웹훅 URL 설정
 * 스크립트 속성에 저장하여 코드에서 URL 숨김
 */
function setupSlackWebhooks() {
  const ui = SpreadsheetApp.getUi();
  
  // 일반 알림 웹훅
  const normalResponse = ui.prompt(
    '슬랙 웹훅 설정',
    '일반 알림 채널(#약재관리-일반)의 웹훅 URL을 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (normalResponse.getSelectedButton() === ui.Button.OK) {
    const normalWebhook = normalResponse.getResponseText();
    PropertiesService.getScriptProperties().setProperty('SLACK_WEBHOOK_NORMAL', normalWebhook);
    Logger.log('✅ 일반 알림 웹훅 저장 완료');
  }
  
  // 긴급 알림 웹훅
  const urgentResponse = ui.prompt(
    '슬랙 웹훅 설정',
    '긴급 알림 채널(#약재관리-긴급)의 웹훅 URL을 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (urgentResponse.getSelectedButton() === ui.Button.OK) {
    const urgentWebhook = urgentResponse.getResponseText();
    PropertiesService.getScriptProperties().setProperty('SLACK_WEBHOOK_URGENT', urgentWebhook);
    Logger.log('✅ 긴급 알림 웹훅 저장 완료');
  }
  
  Browser.msgBox('완료', '슬랙 웹훅 설정이 완료되었습니다!', Browser.Buttons.OK);
}

/**
 * 슬랙 웹훅 URL 가져오기
 */
function getSlackWebhook(type = 'normal') {
  const props = PropertiesService.getScriptProperties();
  
  if (type === 'urgent') {
    return props.getProperty('SLACK_WEBHOOK_URGENT');
  }
  
  return props.getProperty('SLACK_WEBHOOK_NORMAL');
}

/**
 * EMR 스프레드시트 ID 설정
 */
function setupEMRLink() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'EMR 시스템 연동 설정',
    'EMR 스프레드시트 ID를 입력하세요:\n\n(EMR 스프레드시트 URL에서 /d/ 다음의 긴 문자열)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const emrId = response.getResponseText().trim();
    
    // ID 검증
    try {
      const testSS = SpreadsheetApp.openById(emrId);
      const testName = testSS.getName();
      
      // 저장
      PropertiesService.getScriptProperties().setProperty('EMR_SPREADSHEET_ID', emrId);
      
      Logger.log(`✅ EMR 시스템 연동 완료: ${testName}`);
      Browser.msgBox(
        '연동 완료', 
        `EMR 시스템 "${testName}"과(와) 연동되었습니다!`, 
        Browser.Buttons.OK
      );
      
    } catch (error) {
      Browser.msgBox(
        '오류',
        '올바른 스프레드시트 ID가 아니거나 접근 권한이 없습니다.\n\n확인 후 다시 시도하세요.',
        Browser.Buttons.OK
      );
      Logger.log('❌ EMR 연동 실패: ' + error.message);
    }
  }
}

/**
 * EMR 스프레드시트 ID 가져오기
 */
function getEMRSpreadsheetId() {
  return PropertiesService.getScriptProperties().getProperty('EMR_SPREADSHEET_ID');
}

/**
 * EMR 연동 상태 확인
 */
function checkEMRConnection() {
  const emrId = getEMRSpreadsheetId();
  
  if (!emrId) {
    Logger.log('❌ EMR 시스템이 연동되지 않았습니다.');
    return false;
  }
  
  try {
    const emrSS = SpreadsheetApp.openById(emrId);
    const name = emrSS.getName();
    Logger.log(`✅ EMR 시스템 연결됨: ${name}`);
    return true;
  } catch (error) {
    Logger.log('❌ EMR 연결 오류: ' + error.message);
    return false;
  }
}

// ============================================
// EMR 시스템 데이터 조회
// ============================================

/**
 * EMR에서 환자 기본정보 가져오기
 */
function getPatientInfoFromEMR(chartNumber) {
  const emrId = getEMRSpreadsheetId();
  
  if (!emrId) {
    Logger.log('⚠️ EMR 연동 안됨');
    return null;
  }
  
  try {
    const emrSS = SpreadsheetApp.openById(emrId);
    const patientSheet = emrSS.getSheetByName('환자정보');
    
    if (!patientSheet) {
      Logger.log('⚠️ 환자정보 시트 없음');
      return null;
    }
    
    const data = patientSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === chartNumber) {
        return {
          chartNumber: data[i][0],
          name: data[i][1],
          birthDate: data[i][2],
          gender: data[i][3],
          phone: data[i][4],
          address: data[i][5],
          firstVisit: data[i][6],
          lastVisit: data[i][7],
          totalVisits: data[i][8],
          note: data[i][9]
        };
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log(`❌ 환자 정보 조회 오류: ${error.message}`);
    return null;
  }
}

// ============================================
// EMR 시스템 데이터 동기화
// ============================================

/**
 * 처방 입력 시 EMR 환자정보 자동 업데이트
 */
function syncPatientToEMR(chartNumber, patientName, additionalInfo = {}) {
  const emrId = getEMRSpreadsheetId();
  
  if (!emrId) {
    Logger.log('⚠️ EMR 동기화 건너뜀 (연동 안됨)');
    return;
  }
  
  try {
    const emrSS = SpreadsheetApp.openById(emrId);
    const patientSheet = emrSS.getSheetByName('환자정보');
    
    if (!patientSheet) {
      Logger.log('⚠️ 환자정보 시트 없음');
      return;
    }
    
    const data = patientSheet.getDataRange().getValues();
    let patientRow = -1;
    
    // 기존 환자 찾기
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === chartNumber) {
        patientRow = i + 1;
        break;
      }
    }
    
    const today = new Date();
    
    // 신규 환자 등록
    if (patientRow === -1) {
      patientSheet.appendRow([
        chartNumber,
        patientName,
        additionalInfo.birthDate || '',
        additionalInfo.gender || '',
        additionalInfo.phone || '',
        additionalInfo.address || '',
        today,  // 초진일
        today,  // 최종방문일
        1,      // 총방문횟수
        '약재관리 시스템에서 자동 등록'
      ]);
      
      Logger.log(`✅ EMR 신규 환자 등록: ${patientName} (${chartNumber})`);
    }
    // 기존 환자 업데이트
    else {
      // 최종방문일
      patientSheet.getRange(patientRow, 8).setValue(today);
      
      // 총방문횟수 +1
      const currentVisits = patientSheet.getRange(patientRow, 9).getValue() || 0;
      patientSheet.getRange(patientRow, 9).setValue(currentVisits + 1);
      
      Logger.log(`✅ EMR 환자 정보 업데이트: ${patientName} (${chartNumber})`);
    }
    
  } catch (error) {
    Logger.log(`❌ EMR 환자 동기화 오류: ${error.message}`);
  }
}

/**
 * 처방 입력 시 EMR 진료기록 자동 생성
 */
function syncPrescriptionToEMR(prescriptionData) {
  const emrId = getEMRSpreadsheetId();
  
  if (!emrId) {
    Logger.log('⚠️ EMR 동기화 건너뜀');
    return;
  }
  
  try {
    const emrSS = SpreadsheetApp.openById(emrId);
    const recordSheet = emrSS.getSheetByName('진료기록');
    
    if (!recordSheet) {
      Logger.log('⚠️ 진료기록 시트 없음');
      return;
    }
    
    // 진료번호 생성
    const timestamp = Utilities.formatDate(
      new Date(), 
      Session.getScriptTimeZone(), 
      'yyyyMMddHHmmss'
    );
    const recordNumber = `R${timestamp}`;
    
    // 진료기록 추가
    recordSheet.appendRow([
      recordNumber,                      // 진료번호
      new Date(),                        // 진료일시
      prescriptionData.chartNumber,      // 차트번호
      prescriptionData.patientName,      // 환자명
      '',                                // 주소(CC)
      '',                                // 현병력(PI)
      '',                                // 진단
      prescriptionData.prescriptionName, // 처방명
      prescriptionData.doctor,           // 처방의
      '',                                // 녹음파일ID
      '',                                // AI차팅
      '약재관리 시스템에서 동기화됨'    // 비고
    ]);
    
    Logger.log(`✅ EMR 진료기록 동기화: ${recordNumber}`);
    
  } catch (error) {
    Logger.log(`❌ EMR 진료기록 동기화 오류: ${error.message}`);
  }
}

// ============================================
// 기존 addPrescriptionToSheet 함수 수정
// ============================================

function addPrescriptionToSheet(parsedData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prescSheet = ss.getSheetByName('처방입력');
  
  if (!prescSheet) {
    throw new Error('처방입력 시트가 없습니다.');
  }
  
  // 데이터 소스 구분 (OCR vs EMR)
  const isOCR = parsedData.herbs && Array.isArray(parsedData.herbs);
  
  let prescriptionNumber;
  
  if (isOCR) {
    // ===== OCR 데이터 처리 =====
    prescriptionNumber = parsedData.prescriptionNumber || '';
    
    prescSheet.appendRow([
      prescriptionNumber,                   // A: 처방전번호
      parsedData.prescriptionDate || '',    // B: 처방일
      parsedData.prescriptionName || '',    // C: 처방명
      parsedData.chartNumber || '',         // D: 차트번호
      parsedData.patientName || '',         // E: 환자명
      parsedData.cheops || 1,               // F: 첩수
      parsedData.gender || '',              // G: 성별
      parsedData.age || '',                 // H: 나이
      parsedData.birthDate || '',           // I: 생년월일
      parsedData.doctorName || '',          // J: 처방의
      parsedData.herbsList || '',           // K: 약재목록(자동)
      '대기',                               // L: 처리상태
      '',                                   // M: 원가(원)
      ''                                    // N: 완료일시
    ]);
    
    Logger.log(`✅ [OCR] 처방입력: ${prescriptionNumber} - ${parsedData.patientName}`);
    
    // ✅ OCR 데이터도 EMR 동기화
    if (parsedData.chartNumber && parsedData.patientName) {
      try {
        // 환자정보 동기화 (추가 정보 포함)
        syncPatientToEMR(
          parsedData.chartNumber,
          parsedData.patientName,
          {
            birthDate: parsedData.birthDate || '',
            gender: parsedData.gender || '',
            phone: '',
            address: ''
          }
        );
        
        // 진료기록 동기화
        syncPrescriptionToEMR({
          chartNumber: parsedData.chartNumber,
          patientName: parsedData.patientName,
          prescriptionName: parsedData.prescriptionName || '',
          doctor: parsedData.doctorName || ''
        });
        
        Logger.log(`✅ [OCR] EMR 동기화 완료`);
      } catch (error) {
        Logger.log(`⚠️ [OCR] EMR 동기화 실패: ${error.message}`);
      }
    }
    
  } else {
    // ===== EMR 데이터 처리 =====
    prescriptionNumber = parsedData.visitNumber || parsedData.prescriptionNumber || '';
    
    prescSheet.appendRow([
      prescriptionNumber,                   // A: 처방전번호
      parsedData.prescriptionDate || parsedData.visitDateTime || '', // B: 처방일
      parsedData.prescriptionName || '',    // C: 처방명
      parsedData.chartNumber || '',         // D: 차트번호
      parsedData.patientName || '',         // E: 환자명
      '',                                   // F: 첩수
      '',                                   // G: 성별
      '',                                   // H: 나이
      '',                                   // I: 생년월일
      parsedData.doctor || '',              // J: 처방의
      '',                                   // K: 약재목록(자동)
      '대기',                               // L: 처리상태
      '',                                   // M: 원가(원)
      ''                                    // N: 완료일시
    ]);
    
    Logger.log(`✅ [EMR] 처방입력: ${prescriptionNumber} - ${parsedData.patientName}`);
    
    // EMR 동기화
    try {
      syncPatientToEMR(
        parsedData.chartNumber,
        parsedData.patientName
      );
      
      syncPrescriptionToEMR({
        chartNumber: parsedData.chartNumber,
        patientName: parsedData.patientName,
        prescriptionName: parsedData.prescriptionName,
        doctor: parsedData.doctor
      });
      
      Logger.log(`✅ [EMR] EMR 동기화 완료`);
    } catch (error) {
      Logger.log(`⚠️ [EMR] EMR 동기화 실패: ${error.message}`);
    }
  }
  
  return prescriptionNumber;
}

function addPrescriptionDetailsToSheet(prescNumber, parsedData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('처방상세');
  
  if (!sheet) {
    throw new Error('처방상세 시트를 찾을 수 없습니다.');
  }
  
  // OCR 데이터만 처리 (EMR은 약재 정보 없음)
  if (!parsedData.herbs || !Array.isArray(parsedData.herbs)) {
    Logger.log('⚠️ 약재 정보가 없습니다. (EMR 데이터는 처방상세 추가 안함)');
    return;
  }
  
  const startRow = sheet.getLastRow() + 1;
  let addedCount = 0;
  
  parsedData.herbs.forEach((herb) => {
    sheet.appendRow([
      prescNumber,                          // A: 처방전번호
      parsedData.prescriptionName || '',    // B: 처방명
      parsedData.prescriptionDate || '',    // C: 처방일
      parsedData.patientName || '',         // D: 환자명
      parsedData.chartNumber || '',         // E: 챠트번호
      herb.name,                            // F: 약재명
      herb.amountPerCheop,                  // G: 용량
      parsedData.cheops || 1,               // H: 첩수
      herb.totalAmount,                     // I: 총수량
      ''                                    // J: 조제완료
    ]);
    addedCount++;
  });
  
  if (addedCount > 0) {
    const checkboxRange = sheet.getRange(startRow, 10, addedCount, 1);
    checkboxRange.insertCheckboxes();
    Logger.log(`✅ 처방상세 시트 추가: ${addedCount}개 약재 (체크박스 포함)`);
  }
}

// ============================================
// 메뉴 업데이트
// ============================================

/**
 * 스프레드시트 열 때 메뉴 추가 (업데이트 버전)
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🏥 약재관리')
    .addItem('💊 체크된 조제 처리', 'processCheckedNow')
    .addItem('📦 체크된 입고 처리', 'processAllCheckedIncoming')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 재고 관리')
      .addItem('🔄 약재마스터 재고 업데이트', 'updateCurrentStock')
      .addItem('⏰ 자동 업데이트 트리거 설정', 'setupStockUpdateTrigger')
      .addSeparator()
      .addItem('🔍 약재입고 시트 구조 확인', 'checkIncomingSheetStructure')
      .addItem('🔍 약재출고 시트 구조 확인', 'checkDispenseSheetStructure'))
    .addSeparator()
    .addItem('💰 전체 처방 원가 업데이트', 'updateAllPrescriptionCosts')
    .addSeparator()
    .addSubMenu(ui.createMenu('📸 드라이브 OCR')
      .addItem('📋 처방전 OCR 처리', 'processPrescriptionImages')
      .addItem('📦 입고서 OCR 처리', 'processIncomingImagesOCR')
      .addItem('🔄 전체 OCR 처리', 'processAllDriveFiles')
      .addSeparator()
      .addItem('📁 드라이브 폴더 설정', 'setupDriveFolders')
      .addItem('🔍 드라이브 폴더 확인', 'checkDriveFolders'))
    .addSeparator()
    .addItem('🔗 EMR 시스템 연동 설정', 'setupEMRLink')
    .addItem('🔍 EMR 연결 확인', 'testEMRConnection')
    .addToUi();
}

/**
 * EMR 연결 테스트
 */
function testEMRConnection() {
  const emrId = getEMRSpreadsheetId();
  
  if (!emrId) {
    Browser.msgBox(
      'EMR 연동 없음',
      'EMR 시스템이 연동되지 않았습니다.\n\n메뉴: 🏥 약재관리 > 🔗 EMR 시스템 연동 설정',
      Browser.Buttons.OK
    );
    return;
  }
  
  try {
    const emrSS = SpreadsheetApp.openById(emrId);
    const name = emrSS.getName();
    const sheets = emrSS.getSheets().map(s => s.getName()).join(', ');
    
    Browser.msgBox(
      'EMR 연결 성공',
      `EMR 시스템: ${name}\n시트: ${sheets}`,
      Browser.Buttons.OK
    );
    
  } catch (error) {
    Browser.msgBox(
      'EMR 연결 실패',
      '연결에 실패했습니다.\n\n' + error.message,
      Browser.Buttons.OK
    );
  }
}

// ============================================
// 구글 드라이브 자동 OCR 시스템
// ============================================

/**
 * 드라이브 폴더 ID 설정 (최초 1회)
 */
function setupDriveFolders() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    '드라이브 폴더 설정',
    '4개의 폴더 ID를 차례로 입력합니다:\n\n1. 처방전_대기\n2. 처방전_완료\n3. 입고서_대기\n4. 입고서_완료\n\n각 폴더를 미리 만들어두세요!',
    ui.ButtonSet.OK
  );
  
  // 1. 처방전_대기
  const prescWaitResponse = ui.prompt(
    '처방전_대기 폴더',
    '처방전_대기 폴더 ID를 입력하세요:\n(폴더 URL의 /folders/ 다음 부분)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (prescWaitResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const prescWaitId = prescWaitResponse.getResponseText().trim();
  PropertiesService.getScriptProperties().setProperty('DRIVE_PRESC_WAIT', prescWaitId);
  
  // 2. 처방전_완료
  const prescDoneResponse = ui.prompt(
    '처방전_완료 폴더',
    '처방전_완료 폴더 ID를 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (prescDoneResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const prescDoneId = prescDoneResponse.getResponseText().trim();
  PropertiesService.getScriptProperties().setProperty('DRIVE_PRESC_DONE', prescDoneId);
  
  // 3. 입고서_대기
  const incWaitResponse = ui.prompt(
    '입고서_대기 폴더',
    '입고서_대기 폴더 ID를 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (incWaitResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const incWaitId = incWaitResponse.getResponseText().trim();
  PropertiesService.getScriptProperties().setProperty('DRIVE_INC_WAIT', incWaitId);
  
  // 4. 입고서_완료
  const incDoneResponse = ui.prompt(
    '입고서_완료 폴더',
    '입고서_완료 폴더 ID를 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (incDoneResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const incDoneId = incDoneResponse.getResponseText().trim();
  PropertiesService.getScriptProperties().setProperty('DRIVE_INC_DONE', incDoneId);
  
  Browser.msgBox('완료', '드라이브 폴더 설정이 완료되었습니다!', Browser.Buttons.OK);
}

/**
 * 폴더 ID 가져오기
 */
function getDriveFolderId(type) {
  const props = PropertiesService.getScriptProperties();
  
  switch(type) {
    case 'presc_wait':
      return props.getProperty('DRIVE_PRESC_WAIT');
    case 'presc_done':
      return props.getProperty('DRIVE_PRESC_DONE');
    case 'inc_wait':
      return props.getProperty('DRIVE_INC_WAIT');
    case 'inc_done':
      return props.getProperty('DRIVE_INC_DONE');
    default:
      return null;
  }
}

/**
 * 드라이브 폴더 확인
 */
function checkDriveFolders() {
  const prescWait = getDriveFolderId('presc_wait');
  const prescDone = getDriveFolderId('presc_done');
  const incWait = getDriveFolderId('inc_wait');
  const incDone = getDriveFolderId('inc_done');
  
  let message = '드라이브 폴더 설정:\n\n';
  
  if (prescWait) {
    try {
      const folder = DriveApp.getFolderById(prescWait);
      message += `✅ 처방전_대기: ${folder.getName()}\n`;
    } catch (e) {
      message += `❌ 처방전_대기: 접근 불가\n`;
    }
  } else {
    message += `❌ 처방전_대기: 미설정\n`;
  }
  
  if (prescDone) {
    try {
      const folder = DriveApp.getFolderById(prescDone);
      message += `✅ 처방전_완료: ${folder.getName()}\n`;
    } catch (e) {
      message += `❌ 처방전_완료: 접근 불가\n`;
    }
  } else {
    message += `❌ 처방전_완료: 미설정\n`;
  }
  
  if (incWait) {
    try {
      const folder = DriveApp.getFolderById(incWait);
      message += `✅ 입고서_대기: ${folder.getName()}\n`;
    } catch (e) {
      message += `❌ 입고서_대기: 접근 불가\n`;
    }
  } else {
    message += `❌ 입고서_대기: 미설정\n`;
  }
  
  if (incDone) {
    try {
      const folder = DriveApp.getFolderById(incDone);
      message += `✅ 입고서_완료: ${folder.getName()}\n`;
    } catch (e) {
      message += `❌ 입고서_완료: 접근 불가\n`;
    }
  } else {
    message += `❌ 입고서_완료: 미설정\n`;
  }
  
  Browser.msgBox('드라이브 폴더 확인', message, Browser.Buttons.OK);
}

/**
 * 모든 대기 파일 한번에 처리
 */
function processAllDriveFiles() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '전체 OCR 처리',
    '처방전과 입고서를 모두 처리하시겠습니까?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // 처방전 처리
  processPrescriptionImages();

  // 잠시 대기
  Utilities.sleep(2000);

  // 입고서 처리
  processIncomingImagesOCR();
}

/**
 * 재고 가용성만 확인 (차감하지 않음)
 */
function checkStockAvailability(herbName, requiredAmount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const incomingSheet = ss.getSheetByName('약재입고');
  
  if (!incomingSheet) {
    throw new Error('약재입고 시트가 없습니다.');
  }
  
  const data = incomingSheet.getDataRange().getValues();
  
  let totalAvailable = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === herbName) {
      const remainingAmount = parseFloat(data[i][5]) || 0;
      totalAvailable += remainingAmount;
    }
  }
  
  if (totalAvailable < requiredAmount) {
    throw new Error(`재고 부족 (필요: ${requiredAmount}g, 가용: ${totalAvailable}g)`);
  }
  
  return {
    herbName: herbName,
    requiredAmount: requiredAmount,
    totalAvailable: totalAvailable,
    sufficient: true
  };
}

/**
 * 특정 약재 1개만 재고 업데이트 (빠른 버전)
 */
function updateSingleHerbStock(herbName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('약재마스터');
  const incomingSheet = ss.getSheetByName('약재입고');
  const dispenseSheet = ss.getSheetByName('약재출고');
  
  if (!masterSheet || !incomingSheet || !dispenseSheet) {
    Logger.log('⚠️ 필요한 시트가 없습니다.');
    return;
  }
  
  // 약재마스터에서 해당 약재 찾기
  const masterData = masterSheet.getDataRange().getValues();
  let masterRow = -1;
  
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][0] === herbName) { // A열: 약재명
      masterRow = i + 1;
      break;
    }
  }
  
  if (masterRow === -1) {
    Logger.log(`  ⚠️ 약재마스터에 ${herbName} 없음`);
    return;
  }
  
  // 총 입고량 계산
  const incomingData = incomingSheet.getDataRange().getValues();
  let totalIncoming = 0;
  
  for (let i = 1; i < incomingData.length; i++) {
    if (incomingData[i][2] === herbName) { // C열: 약재명
      totalIncoming += parseFloat(incomingData[i][3]) || 0; // D열: 입고량
    }
  }
  
  // 총 출고량 계산
  const dispenseData = dispenseSheet.getDataRange().getValues();
  let totalDispensed = 0;
  
  for (let i = 1; i < dispenseData.length; i++) {
    if (dispenseData[i][2] === herbName) { // C열: 약재명
      totalDispensed += parseFloat(dispenseData[i][3]) || 0; // D열: 출고량
    }
  }
  
  // 현재 재고 = 입고 - 출고
  const currentStock = Math.round((totalIncoming - totalDispensed) * 10) / 10;
  
  // 약재마스터 C열 업데이트
  masterSheet.getRange(masterRow, 3).setValue(currentStock);
  
  Logger.log(`  ✅ 약재마스터 업데이트: ${herbName} → ${currentStock}g`);
  
  // 유통기한도 업데이트
  try {
    const nearestExpiry = getNearestExpiryDate(herbName);
    if (nearestExpiry) {
      masterSheet.getRange(masterRow, 7).setValue(nearestExpiry);
    }
  } catch (e) {
    Logger.log(`  ⚠️ 유통기한 업데이트 실패: ${e.message}`);
  }
}

/**
 * 메뉴 강제 업데이트 (테스트용)
 */
function forceUpdateMenu() {
  onOpen();
  Browser.msgBox('완료', '메뉴가 업데이트되었습니다.', Browser.Buttons.OK);
}

function setupOnOpenTrigger() {
  // 기존 onOpen 트리거 삭제
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onOpen') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 새 onOpen 트리거 생성
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
  
  Browser.msgBox('완료', 'onOpen 트리거가 재생성되었습니다. 새로고침하세요!', Browser.Buttons.OK);
}
