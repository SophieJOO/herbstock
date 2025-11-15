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
 * 약재입고 시트 F열(잔량) 편집 트리거: 해당 약재 재고 즉시 업데이트 + 조정이력 기록
 */
function onIncomingStockEdit(e) {
  try {
    if (!e || !e.source) {
      Logger.log('❌ 이 함수는 수동 실행할 수 없습니다.');
      return;
    }

    const sheet = e.source.getActiveSheet();
    const range = e.range;

    // 약재입고 시트가 아니면 무시
    if (sheet.getName() !== '약재입고') return;

    // F열(6열, 잔량)이 아니면 무시
    if (range.getColumn() !== 6) return;

    const row = range.getRow();
    if (row === 1) return;  // 헤더 제외

    // 편집된 행의 데이터 추출
    const rowData = sheet.getRange(row, 1, 1, 11).getValues()[0];
    const incomingNumber = rowData[0];  // A열: 입고번호
    const incomingDate = rowData[1];    // B열: 입고일
    const herbName = rowData[2];        // C열: 약재명
    const incomingAmount = rowData[3];  // D열: 입고량
    const expiryDate = rowData[4];      // E열: 유통기한
    const newRemaining = parseFloat(range.getValue()) || 0;  // F열: 새 잔량
    const oldRemaining = parseFloat(e.oldValue) || 0;  // 이전 잔량

    if (!herbName || herbName.trim() === '') {
      Logger.log('⚠️ 약재명이 없습니다.');
      return;
    }

    // 값이 변경되지 않았으면 무시
    if (oldRemaining === newRemaining) {
      return;
    }

    const difference = newRemaining - oldRemaining;

    Logger.log(`🔄 잔량 수정 감지: ${herbName} (${row}행) ${oldRemaining}g → ${newRemaining}g (${difference > 0 ? '+' : ''}${difference}g)`);

    // 재고조정이력 기록
    recordStockAdjustment(incomingNumber, incomingDate, herbName, incomingAmount, expiryDate, oldRemaining, newRemaining, difference);

    // 해당 약재만 업데이트
    updateSingleHerbStock(herbName);

    Logger.log(`✅ ${herbName} 약재마스터 업데이트 완료`);

  } catch (error) {
    Logger.log(`⚠️ 약재입고 편집 트리거 오류: ${error.message}`);
  }
}

/**
 * 재고조정이력 기록
 */
function recordStockAdjustment(incomingNumber, incomingDate, herbName, incomingAmount, expiryDate, oldRemaining, newRemaining, difference) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let adjustmentSheet = ss.getSheetByName('재고조정이력');

  // 시트가 없으면 생성
  if (!adjustmentSheet) {
    adjustmentSheet = ss.insertSheet('재고조정이력');
    const headers = [
      '조정일시', '입고번호', '약재명', '입고량(g)', '유통기한',
      '조정 전 잔량(g)', '조정 후 잔량(g)', '조정량(g)', '조정 유형', '조정 사유', '담당자'
    ];
    adjustmentSheet.appendRow(headers);

    const headerRange = adjustmentSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#ff9900');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
  }

  // 조정 유형 선택 UI
  const ui = SpreadsheetApp.getUi();

  const typeResponse = ui.prompt(
    '재고 조정 유형',
    `${herbName} 잔량이 ${oldRemaining}g → ${newRemaining}g로 변경되었습니다.\n\n` +
    `조정 유형을 선택하세요:\n` +
    `1. 폐기 (유통기한 임박, 변질, 파손 등)\n` +
    `2. 타 한의원 대여\n` +
    `3. 단순 조정 (재고 실사, 오입력 수정 등)\n\n` +
    `번호를 입력하세요 (1-3):`,
    ui.ButtonSet.OK_CANCEL
  );

  let typeLabel = '';
  let reason = '';

  if (typeResponse.getSelectedButton() === ui.Button.CANCEL) {
    typeLabel = '단순 조정';
    reason = '사유 미입력 (취소됨)';
  } else {
    const typeNum = typeResponse.getResponseText().trim();

    if (typeNum === '1') {
      // 폐기
      typeLabel = '폐기';
      const response = ui.prompt(
        '폐기 사유',
        '폐기 사유를 입력하세요 (예: 유통기한 임박, 변질, 파손 등):',
        ui.ButtonSet.OK_CANCEL
      );

      if (response.getSelectedButton() === ui.Button.OK) {
        reason = response.getResponseText();
      } else {
        reason = '사유 미입력';
      }
    } else if (typeNum === '2') {
      // 타 한의원 대여
      typeLabel = '타 한의원 대여';
      const response = ui.prompt(
        '대여 정보',
        '대여처 정보를 입력하세요 (예: OO한의원, 반환예정일: 2024-01-15):',
        ui.ButtonSet.OK_CANCEL
      );

      if (response.getSelectedButton() === ui.Button.OK) {
        reason = response.getResponseText();
      } else {
        reason = '정보 미입력';
      }
    } else {
      // 3 또는 기타 = 단순 조정
      typeLabel = '단순 조정';
      const response = ui.prompt(
        '조정 사유',
        '조정 사유를 입력하세요 (예: 재고 실사, 오입력 수정 등):',
        ui.ButtonSet.OK_CANCEL
      );

      if (response.getSelectedButton() === ui.Button.OK) {
        reason = response.getResponseText() || '사유 미입력';
      } else {
        reason = '사유 미입력';
      }
    }
  }

  // 담당자 (현재 사용자)
  const user = Session.getActiveUser().getEmail();

  // 조정 일시
  const now = new Date();
  const adjustmentTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // 유통기한 포맷
  let expiryDateStr = '';
  if (expiryDate instanceof Date) {
    expiryDateStr = Utilities.formatDate(expiryDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } else if (expiryDate) {
    expiryDateStr = String(expiryDate);
  }

  // 입고일 포맷
  let incomingDateStr = '';
  if (incomingDate instanceof Date) {
    incomingDateStr = Utilities.formatDate(incomingDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } else if (incomingDate) {
    incomingDateStr = String(incomingDate);
  }

  // 데이터 추가
  const newRow = [
    adjustmentTime,
    incomingNumber,
    herbName,
    incomingAmount,
    expiryDateStr,
    oldRemaining,
    newRemaining,
    difference,
    typeLabel,
    reason,
    user
  ];

  adjustmentSheet.appendRow(newRow);

  // 마지막 행 색상 구분 (조정량이 음수면 빨강, 양수면 파랑)
  const lastRow = adjustmentSheet.getLastRow();
  const colorRange = adjustmentSheet.getRange(lastRow, 8);  // H열: 조정량

  if (difference < 0) {
    colorRange.setBackground('#f4cccc');  // 빨강 (감소)
  } else if (difference > 0) {
    colorRange.setBackground('#d9ead3');  // 초록 (증가)
  }

  Logger.log(`✅ 재고조정이력 기록: ${herbName} ${difference > 0 ? '+' : ''}${difference}g (${typeLabel})`);
}

/**
 * 통합 편집 트리거 (모든 시트 편집 감지)
 * 주의: 함수명을 onEdit으로 하면 Simple Trigger와 중복 실행되므로 다른 이름 사용
 */
function onEditHandler(e) {
  // 임시입고 시트 처리완료 체크
  onTempIncomingEdit(e);

  // 처방상세 시트 조제완료 체크
  onPrescriptionEdit(e);

  // 약재입고 시트 잔량 수정
  onIncomingStockEdit(e);
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

  // 💰 가격 변동 체크 및 알림
  try {
    checkAndNotifyPriceChange(herbName, pricePerGram, supplier);
  } catch (priceCheckError) {
    Logger.log(`⚠️ 가격 변동 체크 중 오류: ${priceCheckError.message}`);
    // 가격 체크 실패해도 입고는 계속 진행
  }

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
 * 이전 단가 조회 (약재입고 시트에서 최근 입고 단가)
 */
function getPreviousPrice(herbName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const incomingSheet = ss.getSheetByName('약재입고');

  if (!incomingSheet) {
    return null;
  }

  const data = incomingSheet.getDataRange().getValues();

  // 뒤에서부터 검색 (최근 입고)
  for (let i = data.length - 1; i >= 1; i--) {
    const rowHerbName = data[i][2];  // C열: 약재명
    const pricePerGram = parseFloat(data[i][6]);  // G열: 단가(원/g)

    if (rowHerbName === herbName && pricePerGram > 0) {
      return {
        pricePerGram: pricePerGram,
        incomingDate: data[i][1],  // B열: 입고일
        supplier: data[i][7]  // H열: 공급처
      };
    }
  }

  return null;
}

/**
 * 가격이력 시트에 변동 기록
 */
function recordPriceChange(herbName, previousPrice, newPrice, supplier, changePercent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let priceHistorySheet = ss.getSheetByName('가격이력');

  // 가격이력 시트가 없으면 생성
  if (!priceHistorySheet) {
    priceHistorySheet = ss.insertSheet('가격이력');

    const headers = [
      '변동일시', '약재명', '이전단가(원/g)', '신규단가(원/g)',
      '변동금액(원/g)', '변동률(%)', '공급처', '비고'
    ];
    priceHistorySheet.appendRow(headers);

    const headerRange = priceHistorySheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#f4b400');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
  }

  const currentDate = new Date();
  const priceChange = newPrice - previousPrice;
  const changeDirection = priceChange > 0 ? '⬆️ 상승' : '⬇️ 하락';

  priceHistorySheet.appendRow([
    currentDate,
    herbName,
    previousPrice,
    newPrice,
    priceChange,
    changePercent,
    supplier,
    changeDirection
  ]);

  // 변동률에 따라 색상 구분
  const lastRow = priceHistorySheet.getLastRow();
  const changePercentCell = priceHistorySheet.getRange(lastRow, 6);

  if (Math.abs(changePercent) >= 20) {
    changePercentCell.setBackground('#f4cccc');  // 20% 이상: 빨강
  } else if (Math.abs(changePercent) >= 10) {
    changePercentCell.setBackground('#fff2cc');  // 10% 이상: 노랑
  }

  Logger.log(`✅ 가격이력 기록: ${herbName} ${changePercent}% ${changeDirection}`);
}

/**
 * 단가 변동 체크 및 슬랙 알림
 */
function checkAndNotifyPriceChange(herbName, newPricePerGram, supplier) {
  const previousInfo = getPreviousPrice(herbName);

  if (!previousInfo) {
    Logger.log(`ℹ️ ${herbName}: 첫 입고 - 가격 비교 없음`);
    return;
  }

  const previousPrice = previousInfo.pricePerGram;
  const priceChange = newPricePerGram - previousPrice;
  const changePercent = ((priceChange / previousPrice) * 100).toFixed(1);

  Logger.log(`💰 ${herbName} 단가 비교:`);
  Logger.log(`   이전: ${previousPrice}원/g`);
  Logger.log(`   신규: ${newPricePerGram}원/g`);
  Logger.log(`   변동: ${priceChange > 0 ? '+' : ''}${priceChange}원/g (${changePercent}%)`);

  // 가격 변동이 있으면 무조건 기록
  if (priceChange !== 0) {
    // 가격이력 시트에 기록 (변동이 조금이라도 있으면 무조건 기록)
    recordPriceChange(herbName, previousPrice, newPricePerGram, supplier, parseFloat(changePercent));
    Logger.log(`✅ 가격이력 기록: ${herbName} ${changePercent}% 변동`);

    // 슬랙 알림은 5% 이상 변동 시에만 발송 (너무 많은 알림 방지)
    const ALERT_THRESHOLD = 5;

    if (Math.abs(parseFloat(changePercent)) >= ALERT_THRESHOLD) {
      Logger.log(`⚠️ ${ALERT_THRESHOLD}% 이상 변동 감지 - 슬랙 알림 발송`);

      const emoji = priceChange > 0 ? '📈' : '📉';
      const direction = priceChange > 0 ? '상승' : '하락';
      const color = priceChange > 0 ? '#ea4335' : '#34a853';

      const message = {
        text: `${emoji} *단가 변동 알림*`,
        attachments: [{
          color: color,
          fields: [
            {
              title: '약재명',
              value: herbName,
              short: true
            },
            {
              title: '공급처',
              value: supplier,
              short: true
            },
            {
              title: '이전 단가',
              value: `${previousPrice}원/g`,
              short: true
            },
            {
              title: '신규 단가',
              value: `${newPricePerGram}원/g`,
              short: true
            },
            {
              title: '변동금액',
              value: `${priceChange > 0 ? '+' : ''}${priceChange}원/g`,
              short: true
            },
            {
              title: '변동률',
              value: `${priceChange > 0 ? '+' : ''}${changePercent}% ${direction}`,
              short: true
            }
          ],
          footer: '가격이력 시트에서 전체 이력 확인 가능',
          ts: Math.floor(Date.now() / 1000)
        }]
      };

      try {
        sendSlackAlert(JSON.stringify(message));
        Logger.log(`✅ 슬랙 알림 발송 완료`);
      } catch (error) {
        Logger.log(`⚠️ 슬랙 알림 실패: ${error.message}`);
      }
    }
  } else {
    Logger.log(`ℹ️ 가격 변동 없음 - 기록 생략`);
  }
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

    // ✅ JSON 복구 로직 (강화)
    if (jsonEnd === -1 || jsonEnd < jsonStart) {
      Logger.log('⚠️ JSON이 불완전합니다. 자동 복구 시도...');

      jsonText = textContent.substring(jsonStart);

      // 1. 불완전한 마지막 항목 제거
      const lastComma = jsonText.lastIndexOf(',');
      const lastCloseBrace = jsonText.lastIndexOf('}');
      const lastCloseBracket = jsonText.lastIndexOf(']');

      // herbs 배열 내 마지막 완전한 객체까지만 사용
      if (lastCloseBrace !== -1) {
        // 마지막 }가 herbs 배열 안에 있는지 확인
        const herbsStart = jsonText.indexOf('"herbs"');
        if (herbsStart !== -1 && lastCloseBrace > herbsStart) {
          // 마지막 } 이후의 불완전한 부분 제거
          if (lastComma > lastCloseBrace) {
            jsonText = jsonText.substring(0, lastCloseBrace + 1);
          }
        }
      }

      // 2. herbs 배열 닫기
      if (jsonText.includes('"herbs"')) {
        const herbsArrayStart = jsonText.indexOf('"herbs"');
        const bracketAfterHerbs = jsonText.indexOf('[', herbsArrayStart);

        if (bracketAfterHerbs !== -1 && jsonText.lastIndexOf(']') < bracketAfterHerbs) {
          // herbs 배열이 열렸지만 닫히지 않음
          jsonText += '\n  ]\n}';
          Logger.log('✅ herbs 배열 자동 닫기');
        } else if (!jsonText.trim().endsWith('}')) {
          // herbs 배열은 닫혔지만 전체 객체가 안 닫힘
          jsonText += '\n}';
          Logger.log('✅ 전체 객체 자동 닫기');
        }
      } else if (!jsonText.trim().endsWith('}')) {
        jsonText += '\n}';
      }

      Logger.log('✅ 복구된 JSON (처음 500자): ' + jsonText.substring(0, 500));
      Logger.log('✅ 복구된 JSON (마지막 200자): ' + jsonText.substring(Math.max(0, jsonText.length - 200)));
    } else {
      jsonText = textContent.substring(jsonStart, jsonEnd + 1);
    }

    Logger.log('추출된 JSON (길이: ' + jsonText.length + ')');
    
    try {
      const parsed = JSON.parse(jsonText);
      
      // ✅ 데이터 검증 (관대하게)
      if (!parsed.herbs || !Array.isArray(parsed.herbs) || parsed.herbs.length === 0) {
        throw new Error('약재 항목이 없습니다.');
      }

      // ⚠️ 중요 필드 체크 (경고만, 에러 아님)
      if (!parsed.patientName) {
        Logger.log('⚠️ 환자명이 없습니다. "미상"으로 설정합니다.');
        parsed.patientName = '미상';
      }

      if (!parsed.cheops || parsed.cheops <= 0) {
        Logger.log('⚠️ 첩수가 없거나 올바르지 않습니다. 기본값 1로 설정합니다.');
        parsed.cheops = 1;
      }

      // 기본값 설정
      parsed.prescriptionNumber = parsed.prescriptionNumber || '';
      parsed.prescriptionDate = parsed.prescriptionDate || new Date().toISOString().split('T')[0];
      parsed.prescriptionName = parsed.prescriptionName || '처방명 미상';
      parsed.chartNumber = parsed.chartNumber || '';
      parsed.gender = parsed.gender || '';
      parsed.age = parsed.age || null;
      parsed.birthDate = parsed.birthDate || '';
      parsed.doctorName = parsed.doctorName || '';
      parsed.clinicName = parsed.clinicName || '';

      // ✅ 약재 항목 안전 처리
      parsed.herbs = parsed.herbs.filter(herb => herb.name).map(herb => ({
        name: herb.name,
        amountPerCheop: parseFloat(herb.amountPerCheop) || 0,
        totalAmount: (parseFloat(herb.amountPerCheop) || 0) * parsed.cheops
      }));

      // 빈 약재 항목 제거 후 재확인
      if (parsed.herbs.length === 0) {
        throw new Error('유효한 약재 항목이 없습니다.');
      }
      
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
/**
 * 처방상세 시트 편집 트리거 (조제완료 체크)
 */
function onPrescriptionEdit(e) {
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

  // ✅ 약재마스터 재고 즉시 업데이트 (출고 즉시 반영)
  updateSingleHerbStock(herbName);

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
  
  Logger.log('=== 재고 업데이트 시작 (✅ 잔량 기준 계산 - 버전 eb14291) ===');

  for (let i = 1; i < masterData.length; i++) {
    const herbName = masterData[i][0];
    
    if (!herbName || herbName.trim() === '') {
      continue;
    }
    
    // 현재 재고 = 약재입고 시트의 잔량(F열) 합계
    // F열은 이미 출고를 반영한 실제 남은 재고량이므로 출고량을 별도로 빼지 않음
    let currentStock = 0;
    let suppliers = new Set();

    for (let j = 1; j < incomingData.length; j++) {
      if (incomingData[j][2] === herbName) {  // C열: 약재명
        const remainingAmount = parseFloat(incomingData[j][5]) || 0;  // F열: 잔량
        currentStock += remainingAmount;

        const supplier = incomingData[j][7];  // H열: 공급처
        if (supplier && supplier.trim() !== '') {
          suppliers.add(supplier.trim());
        }
      }
    }
    
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

    Logger.log(`${herbName}: 현재 재고 ${currentStock}g (약재입고 시트 잔량 합계)`);

    // 💰 재고 부족 체크 및 알림
    try {
      const minimumStock = masterData[i][3]; // D열: 최소재고량

      if (minimumStock && minimumStock > 0 && currentStock < minimumStock) {
        const shortageAmount = minimumStock - currentStock;
        Logger.log(`🚨 재고 부족: ${herbName} (현재: ${currentStock}g, 최소: ${minimumStock}g, 부족: ${shortageAmount}g)`);
        sendLowStockAlert(herbName, shortageAmount);
      }
    } catch (e) {
      Logger.log(`⚠️ ${herbName} 재고 부족 체크 실패: ${e.message}`);
    }
  }

  Logger.log('✅ 약재마스터 현재 재고 업데이트 완료');
}

/**
 * 재고분석 시트 자동 업데이트
 * 약재명|총입고|총출고|현재고|재고회전율
 */
function updateInventoryAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('약재마스터');
  const incomingSheet = ss.getSheetByName('약재입고');
  const prescDetailSheet = ss.getSheetByName('처방상세');
  let analysisSheet = ss.getSheetByName('재고분석');

  if (!masterSheet) {
    Logger.log('❌ 약재마스터 시트가 없습니다.');
    return;
  }

  // 재고분석 시트가 없으면 생성
  if (!analysisSheet) {
    analysisSheet = ss.insertSheet('재고분석');
    const headers = ['약재명', '총입고', '총출고', '현재고', '재고회전율'];
    analysisSheet.appendRow(headers);

    const headerRange = analysisSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a73e8');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    Logger.log('✅ 재고분석 시트 생성 완료');
  }

  const masterData = masterSheet.getDataRange().getValues();

  // 약재입고 데이터
  let incomingData = [];
  if (incomingSheet) {
    incomingData = incomingSheet.getDataRange().getValues();
  } else {
    Logger.log('⚠️ 약재입고 시트가 없습니다.');
  }

  // 처방상세 데이터
  let prescDetailData = [];
  if (prescDetailSheet) {
    prescDetailData = prescDetailSheet.getDataRange().getValues();
  } else {
    Logger.log('⚠️ 처방상세 시트가 없습니다.');
  }

  Logger.log('=== 재고분석 업데이트 시작 ===');

  // 기존 데이터 초기화 (헤더 제외)
  if (analysisSheet.getLastRow() > 1) {
    analysisSheet.getRange(2, 1, analysisSheet.getLastRow() - 1, 5).clearContent();
  }

  const analysisData = [];

  // 약재마스터의 모든 약재에 대해 계산
  for (let i = 1; i < masterData.length; i++) {
    const herbName = masterData[i][0];  // A열: 약재명
    const currentStock = parseFloat(masterData[i][2]) || 0;  // C열: 현재재고

    if (!herbName || herbName.trim() === '') {
      continue;
    }

    // 총입고량 계산 (약재입고 시트의 D열: 입고량 합계)
    let totalIncoming = 0;
    for (let j = 1; j < incomingData.length; j++) {
      if (incomingData[j][2] === herbName) {  // C열: 약재명
        const incomingAmount = parseFloat(incomingData[j][3]) || 0;  // D열: 입고량
        totalIncoming += incomingAmount;
      }
    }

    // 총출고량 계산 (처방상세 시트의 I열: 총수량 합계)
    let totalDispensed = 0;
    for (let k = 1; k < prescDetailData.length; k++) {
      if (prescDetailData[k][5] === herbName) {  // F열: 약재명
        const dispensedAmount = parseFloat(prescDetailData[k][8]) || 0;  // I열: 총수량(g)
        totalDispensed += dispensedAmount;
      }
    }

    // 재고회전율 계산 (총출고 ÷ 현재고)
    let turnoverRate = '';
    if (currentStock > 0 && totalDispensed > 0) {
      turnoverRate = (totalDispensed / currentStock).toFixed(2);
    } else if (currentStock === 0 && totalDispensed > 0) {
      turnoverRate = '∞';  // 재고 없이 출고만 있는 경우
    } else {
      turnoverRate = 'N/A';  // 출고 없음
    }

    analysisData.push([
      herbName,
      Math.round(totalIncoming * 10) / 10,  // 소수점 1자리
      Math.round(totalDispensed * 10) / 10,
      Math.round(currentStock * 10) / 10,
      turnoverRate
    ]);

    Logger.log(`${herbName}: 입고 ${totalIncoming}g, 출고 ${totalDispensed}g, 재고 ${currentStock}g, 회전율 ${turnoverRate}`);
  }

  // 재고분석 시트에 데이터 입력
  if (analysisData.length > 0) {
    analysisSheet.getRange(2, 1, analysisData.length, 5).setValues(analysisData);
  }

  // 숫자 컬럼 정렬 및 포맷
  if (analysisData.length > 0) {
    // B~D열 숫자 포맷 (천단위 구분)
    analysisSheet.getRange(2, 2, analysisData.length, 3).setNumberFormat('#,##0.0');

    // E열 재고회전율 (소수점 2자리)
    analysisSheet.getRange(2, 5, analysisData.length, 1).setHorizontalAlignment('right');
  }

  Logger.log(`✅ 재고분석 업데이트 완료 (${analysisData.length}개 약재)`);
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
 * 감모율 분석 (폐기 이력 기반)
 */
function analyzeSpoilageRate(herbName, days = 365) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adjustmentSheet = ss.getSheetByName('재고조정이력');
  const incomingSheet = ss.getSheetByName('약재입고');

  // 재고조정이력 시트가 없으면 감모 없음으로 처리
  if (!adjustmentSheet) {
    return {
      totalSpoilage: 0,
      spoilageAmount: 0,
      spoilageRate: 0,
      avgSpoilagePerMonth: 0,
      totalIncoming: 0
    };
  }

  const today = new Date();
  const startDate = new Date(today.getTime() - (days * 24 * 60 * 60 * 1000));

  // 재고조정이력에서 폐기 데이터 수집
  const adjustmentData = adjustmentSheet.getDataRange().getValues();
  let totalSpoilage = 0;
  let spoilageRecords = [];

  for (let i = 1; i < adjustmentData.length; i++) {
    const adjustmentTime = adjustmentData[i][0];  // A열: 조정일시
    const adjustHerbName = adjustmentData[i][2];  // C열: 약재명
    const adjustmentAmount = parseFloat(adjustmentData[i][7]) || 0;  // H열: 조정량
    const adjustmentType = adjustmentData[i][8];  // I열: 조정 유형

    if (adjustHerbName !== herbName) continue;
    if (adjustmentType !== '폐기') continue;

    // 날짜 파싱
    let adjDate;
    if (adjustmentTime instanceof Date) {
      adjDate = adjustmentTime;
    } else {
      try {
        adjDate = new Date(adjustmentTime);
      } catch (e) {
        continue;
      }
    }

    if (adjDate >= startDate && adjDate <= today) {
      // 폐기는 음수로 기록되므로 절대값 사용
      const spoilageAmount = Math.abs(adjustmentAmount);
      totalSpoilage += spoilageAmount;
      spoilageRecords.push({
        date: adjDate,
        amount: spoilageAmount
      });
    }
  }

  // 동일 기간 총 입고량 계산
  if (!incomingSheet) {
    return {
      totalSpoilage: totalSpoilage,
      spoilageAmount: 0,
      spoilageRate: 0,
      avgSpoilagePerMonth: totalSpoilage / (days / 30),
      totalIncoming: 0
    };
  }

  const incomingData = incomingSheet.getDataRange().getValues();
  let totalIncoming = 0;
  let totalSpoilageValue = 0;

  for (let i = 1; i < incomingData.length; i++) {
    const incomingDate = incomingData[i][1];  // B열: 입고일
    const incomingHerbName = incomingData[i][2];  // C열: 약재명
    const incomingAmount = parseFloat(incomingData[i][3]) || 0;  // D열: 입고량
    const pricePerGram = parseFloat(incomingData[i][6]) || 0;  // G열: 단가

    if (incomingHerbName !== herbName) continue;

    // 날짜 파싱
    let incDate;
    if (incomingDate instanceof Date) {
      incDate = incomingDate;
    } else {
      try {
        incDate = new Date(incomingDate);
      } catch (e) {
        continue;
      }
    }

    if (incDate >= startDate && incDate <= today) {
      totalIncoming += incomingAmount;

      // 폐기 금액 계산 (평균 단가 사용)
      totalSpoilageValue += totalSpoilage * pricePerGram;
    }
  }

  // 감모율 계산 (%)
  const spoilageRate = totalIncoming > 0 ? (totalSpoilage / totalIncoming) * 100 : 0;

  // 월평균 폐기량
  const avgSpoilagePerMonth = totalSpoilage / (days / 30);

  return {
    totalSpoilage: Math.round(totalSpoilage * 10) / 10,
    spoilageAmount: Math.round(totalSpoilageValue),
    spoilageRate: Math.round(spoilageRate * 100) / 100,
    avgSpoilagePerMonth: Math.round(avgSpoilagePerMonth * 10) / 10,
    totalIncoming: Math.round(totalIncoming * 10) / 10
  };
}

/**
 * 약재 출고 히스토리 수집 (AI 분석용)
 */
function getUsageHistory(herbName, days = 120) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dispenseSheet = ss.getSheetByName('약재출고');

  if (!dispenseSheet) {
    return [];
  }

  const data = dispenseSheet.getDataRange().getValues();

  if (data.length <= 1) {
    return [];
  }

  const today = new Date();
  const startDate = new Date(today.getTime() - (days * 24 * 60 * 60 * 1000));

  const history = [];

  for (let i = 1; i < data.length; i++) {
    const dateValue = data[i][0];  // A열: 출고일
    const name = data[i][2];  // C열: 약재명
    const amount = parseFloat(data[i][3]) || 0;  // D열: 출고량

    if (name !== herbName) continue;

    let date;
    if (dateValue instanceof Date) {
      date = dateValue;
    } else {
      date = new Date(dateValue);
    }

    if (date >= startDate && date <= today) {
      history.push({
        date: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        amount: amount
      });
    }
  }

  // 날짜순 정렬
  history.sort((a, b) => new Date(a.date) - new Date(b.date));

  return history;
}

/**
 * AI 기반 최적재고량 분석 (Gemini API + 감모율 반영)
 */
function analyzeOptimalStockWithAI(herbName, usageHistory) {
  // 출고 데이터가 부족하면 기본값 사용
  if (usageHistory.length < 7) {
    Logger.log(`⚠️ ${herbName}: 데이터 부족 (${usageHistory.length}건) - 기본 계산 사용`);
    const avgUsage = calculateAverageDailyUsage(herbName, 120);
    return {
      optimalStock: Math.round(avgUsage * 7 * 1.2),
      avgDailyUsage: avgUsage,
      confidence: 'low',
      reason: '데이터 부족으로 기본 계산 사용',
      spoilageRate: 0
    };
  }

  const apiKey = getConfig('GEMINI_API_KEY');
  if (!apiKey) {
    Logger.log('⚠️ Gemini API 키가 없습니다 - 기본 계산 사용');
    const avgUsage = calculateAverageDailyUsage(herbName, 120);
    return {
      optimalStock: Math.round(avgUsage * 7 * 1.2),
      avgDailyUsage: avgUsage,
      confidence: 'low',
      reason: 'API 키 없음',
      spoilageRate: 0
    };
  }

  // 감모율 분석 (연간)
  const spoilageAnalysis = analyzeSpoilageRate(herbName, 365);

  // 주간 사용량 집계 (AI 분석 효율화)
  const weeklyData = [];
  let weekStart = null;
  let weekTotal = 0;

  usageHistory.forEach((record, idx) => {
    const recordDate = new Date(record.date);

    if (!weekStart) {
      weekStart = record.date;
      weekTotal = record.amount;
    } else {
      const daysDiff = Math.floor((recordDate - new Date(weekStart)) / (1000 * 60 * 60 * 24));

      if (daysDiff < 7) {
        weekTotal += record.amount;
      } else {
        weeklyData.push({ week: weekStart, total: Math.round(weekTotal) });
        weekStart = record.date;
        weekTotal = record.amount;
      }
    }

    // 마지막 주 추가
    if (idx === usageHistory.length - 1 && weekTotal > 0) {
      weeklyData.push({ week: weekStart, total: Math.round(weekTotal) });
    }
  });

  const prompt = `당신은 한의원 약재 재고 관리 전문가입니다.

약재명: ${herbName}
분석 기간: 최근 ${usageHistory.length}일 (${weeklyData.length}주)

주간 사용량 데이터:
${weeklyData.map((w, i) => `${i + 1}주차 (${w.week}): ${w.total}g`).join('\n')}

📊 감모율 분석 (최근 1년):
- 감모율: ${spoilageAnalysis.spoilageRate}%
- 총 폐기량: ${spoilageAnalysis.totalSpoilage}g
- 폐기 금액: ${spoilageAnalysis.spoilageAmount.toLocaleString()}원
- 월평균 폐기: ${spoilageAnalysis.avgSpoilagePerMonth}g

다음을 분석하여 JSON으로 응답하세요:
1. 평균 일일 소비량 (avgDailyUsage: 숫자)
2. 계절성 패턴 (seasonality: "높음/중간/낮음")
3. 증가/감소 트렌드 (trend: "증가/안정/감소")
4. 최근 변동성 (volatility: "높음/중간/낮음")
5. 권장 최소재고량 (optimalStock: 숫자, 단위 g)
   - 리드타임 7일 고려
   - 변동성에 따른 안전계수 (낮음: 1.2배, 중간: 1.3배, 높음: 1.5배)
   - 트렌드 반영 (증가 추세면 더 높게)
   - 🔥 감모율 반영 (매우 중요):
     * 감모율 10% 이상: 소량 주문 권장 (안전계수 1.1배로 낮춤)
     * 감모율 3-10%: 정상 운영 (기본 안전계수)
     * 감모율 3% 미만: 대량 주문 가능 (안전계수 1.5배로 높임)
6. 신뢰도 (confidence: "high/medium/low")
7. 분석 근거 (reason: 한줄 설명, 감모율 언급 필수)

응답 형식 (JSON만):
{
  "avgDailyUsage": 숫자,
  "seasonality": "높음/중간/낮음",
  "trend": "증가/안정/감소",
  "volatility": "높음/중간/낮음",
  "optimalStock": 숫자,
  "confidence": "high/medium/low",
  "reason": "분석 근거 (감모율 ${spoilageAnalysis.spoilageRate}% 반영)"
}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent?key=${apiKey}`;

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.3,
      maxOutputTokens: 1024,
      responseMimeType: "application/json"
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    let attempt = 0;
    const maxRetries = 3;

    while (attempt < maxRetries) {
      attempt++;

      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();

      if (statusCode === 503) {
        Logger.log(`⚠️ Gemini API 503 오류 (${attempt}/${maxRetries})`);
        if (attempt < maxRetries) {
          const waitTime = attempt * 5;
          Logger.log(`⏳ ${waitTime}초 대기 중...`);
          Utilities.sleep(waitTime * 1000);
          continue;
        } else {
          throw new Error('Gemini API 503 오류 (재시도 실패)');
        }
      }

      if (statusCode !== 200) {
        throw new Error(`Gemini API 오류: ${statusCode} - ${response.getContentText()}`);
      }

      const result = JSON.parse(response.getContentText());

      if (!result.candidates || result.candidates.length === 0) {
        throw new Error('Gemini API 응답 없음');
      }

      const textContent = result.candidates[0].content.parts[0].text;
      const analysis = JSON.parse(textContent);

      // 감모율 정보 추가
      analysis.spoilageRate = spoilageAnalysis.spoilageRate;
      analysis.spoilageAmount = spoilageAnalysis.spoilageAmount;

      Logger.log(`✅ ${herbName} AI 분석 완료: 평균 ${analysis.avgDailyUsage}g/일, 최적재고 ${analysis.optimalStock}g`);
      Logger.log(`   트렌드: ${analysis.trend}, 변동성: ${analysis.volatility}, 감모율: ${analysis.spoilageRate}%`);
      Logger.log(`   신뢰도: ${analysis.confidence}, 이유: ${analysis.reason}`);

      return analysis;

    }

    // 모든 재시도 실패
    throw new Error('Gemini API 재시도 모두 실패');

  } catch (error) {
    Logger.log(`❌ ${herbName} AI 분석 실패: ${error.message} - 기본 계산 사용`);
    const avgUsage = calculateAverageDailyUsage(herbName, 120);
    return {
      optimalStock: Math.round(avgUsage * 7 * 1.2),
      avgDailyUsage: avgUsage,
      confidence: 'low',
      reason: `AI 분석 실패: ${error.message}`
    };
  }
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

  Logger.log('=== AI 기반 최소재고량 자동 업데이트 시작 ===');

  for (let i = 1; i < data.length; i++) {
    const herbName = data[i][0];

    if (!herbName || herbName.trim() === '') {
      continue;
    }

    // 출고 히스토리 수집
    const usageHistory = getUsageHistory(herbName, 120);

    // AI 분석
    const analysis = analyzeOptimalStockWithAI(herbName, usageHistory);

    // F열에 평균일일소비량 업데이트
    masterSheet.getRange(i + 1, 6).setValue(Math.round(analysis.avgDailyUsage * 10) / 10);

    // D열에 최소재고량 업데이트
    masterSheet.getRange(i + 1, 4).setValue(analysis.optimalStock);

    // I열에 감모율 업데이트 (%)
    masterSheet.getRange(i + 1, 9).setValue(analysis.spoilageRate);

    // E열에 분석 결과 메모 (선택사항 - 없으면 무시)
    try {
      const memo = `${analysis.trend} / ${analysis.volatility} / ${analysis.confidence}`;
      masterSheet.getRange(i + 1, 5).setNote(analysis.reason);
    } catch (e) {
      // E열이 없거나 권한 문제면 무시
    }

    Logger.log(`${herbName}: 평균 ${Math.round(analysis.avgDailyUsage)}g/일 → 최적재고 ${analysis.optimalStock}g (감모율 ${analysis.spoilageRate}%, ${analysis.confidence})`);
  }

  Logger.log('✅ AI 기반 최소재고량 자동 업데이트 완료');
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

/**
 * Slack 메시지 전송 (공통 함수)
 */
function sendSlackMessage(webhookUrl, payload) {
  if (!webhookUrl) {
    Logger.log('⚠️ Slack Webhook URL이 설정되지 않았습니다.');
    return;
  }

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    const statusCode = response.getResponseCode();

    if (statusCode === 200) {
      Logger.log('✅ Slack 메시지 전송 성공');
    } else {
      Logger.log(`⚠️ Slack 메시지 전송 실패: ${statusCode} - ${response.getContentText()}`);
    }
  } catch (error) {
    Logger.log(`❌ Slack 메시지 전송 오류: ${error.message}`);
  }
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

  // 3-1. 재고분석 자동 업데이트 (1시간마다)
  ScriptApp.newTrigger('updateInventoryAnalysis')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('✅ updateInventoryAnalysis 트리거 생성');

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
  
  // 6. 통합 편집 트리거 (임시입고, 처방상세, 약재입고)
  ScriptApp.newTrigger('onEditHandler')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  Logger.log('✅ 통합 onEditHandler 트리거 생성 (임시입고/처방상세/약재입고)');

  Logger.log('\n✅✅✅ 모든 트리거 설정 완료!');
  Browser.msgBox('완료', '모든 트리거가 설정되었습니다!', Browser.Buttons.OK);
}

/**
 * 재고 업데이트 트리거만 설정 (개별 설정용)
 */
function setupStockUpdateTrigger() {
  // 기존 재고 업데이트 트리거 삭제
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'updateCurrentStock' ||
        trigger.getHandlerFunction() === 'updateInventoryAnalysis') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  Logger.log('기존 재고 업데이트 트리거 삭제 완료');

  // 재고 자동 업데이트 트리거 생성
  ScriptApp.newTrigger('updateCurrentStock')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('✅ updateCurrentStock 트리거 생성');

  ScriptApp.newTrigger('updateInventoryAnalysis')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('✅ updateInventoryAnalysis 트리거 생성');

  Browser.msgBox('완료', '재고 업데이트 트리거가 설정되었습니다!\n\n- 약재마스터 재고 업데이트 (1시간마다)\n- 재고분석 업데이트 (1시간마다)', Browser.Buttons.OK);
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
      .addItem('📊 재고분석 업데이트', 'updateInventoryAnalysis')
      .addItem('⏰ 자동 업데이트 트리거 설정', 'setupStockUpdateTrigger')
      .addSeparator()
      .addItem('🔍 약재입고 시트 구조 확인', 'checkIncomingSheetStructure')
      .addItem('🔍 약재출고 시트 구조 확인', 'checkDispenseSheetStructure'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 관리')
      .addItem('⚙️ 모든 트리거 설정', 'setupAllTriggers')
      .addItem('📋 시스템 테스트', 'testSystem')
      .addSeparator()
      .addItem('🔍 트리거 상태 확인', 'checkTriggerStatus')
      .addItem('🧪 조제 처리 테스트', 'testPrescriptionProcessing')
      .addItem('🔍 처방상세 시트 구조 확인', 'checkPrescriptionSheetStructure'))
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
  
  // 현재 재고 = 약재입고 시트의 잔량(F열) 합계
  // F열은 이미 출고를 반영한 실제 남은 재고량이므로 출고량을 별도로 빼지 않음
  const incomingData = incomingSheet.getDataRange().getValues();
  let currentStock = 0;

  for (let i = 1; i < incomingData.length; i++) {
    if (incomingData[i][2] === herbName) { // C열: 약재명
      const remainingAmount = parseFloat(incomingData[i][5]) || 0; // F열: 잔량
      currentStock += remainingAmount;
    }
  }

  currentStock = Math.round(currentStock * 10) / 10;

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

  // 💰 재고 부족 체크 및 알림
  try {
    const minimumStock = masterData[masterRow - 1][3]; // D열: 최소재고량

    if (minimumStock && minimumStock > 0 && currentStock < minimumStock) {
      const shortageAmount = minimumStock - currentStock;
      Logger.log(`  🚨 재고 부족: ${herbName} (현재: ${currentStock}g, 최소: ${minimumStock}g, 부족: ${shortageAmount}g)`);
      sendLowStockAlert(herbName, shortageAmount);
    }
  } catch (e) {
    Logger.log(`  ⚠️ 재고 부족 체크 실패: ${e.message}`);
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

// ========================================
// 🔍 진단 및 테스트 함수
// ========================================

/**
 * 처방상세 시트 구조 확인
 */
function checkPrescriptionSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('처방상세');

  if (!sheet) {
    Browser.msgBox('오류', '처방상세 시트가 없습니다.', Browser.Buttons.OK);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dataRowCount = sheet.getLastRow() - 1;

  let message = '처방상세 시트 구조:\n\n';
  headers.forEach((header, index) => {
    message += `${index + 1}열: ${header}\n`;
  });

  message += `\n총 ${dataRowCount}개의 조제 대기 항목`;

  if (headers[9] === '조제완료') {
    message += '\n\n✅ 조제완료 컬럼 위치: 10열 (정상)';
  } else {
    message += `\n\n⚠️ 10열이 "조제완료"가 아닙니다: "${headers[9]}"`;
  }

  Browser.msgBox('처방상세 시트 구조', message, Browser.Buttons.OK);
}

/**
 * 약재입고 시트 구조 확인
 */
function checkIncomingSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('약재입고');

  if (!sheet) {
    Browser.msgBox('오류', '약재입고 시트가 없습니다.', Browser.Buttons.OK);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dataRowCount = sheet.getLastRow() - 1;

  let message = '약재입고 시트 구조:\n\n';
  headers.forEach((header, index) => {
    message += `${index + 1}열: ${header}\n`;
  });

  message += `\n총 ${dataRowCount}개의 입고 기록`;

  Browser.msgBox('약재입고 시트 구조', message, Browser.Buttons.OK);
}

/**
 * 약재출고 시트 구조 확인
 */
function checkDispenseSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('약재출고');

  if (!sheet) {
    Browser.msgBox('안내', '약재출고 시트가 아직 생성되지 않았습니다.\n\n첫 조제 처리 시 자동으로 생성됩니다.', Browser.Buttons.OK);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dataRowCount = sheet.getLastRow() - 1;

  let message = '약재출고 시트 구조:\n\n';
  headers.forEach((header, index) => {
    message += `${index + 1}열: ${header}\n`;
  });

  message += `\n총 ${dataRowCount}개의 출고 기록`;

  Browser.msgBox('약재출고 시트 구조', message, Browser.Buttons.OK);
}

/**
 * 트리거 상태 확인
 */
function checkTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();

  let message = '📊 현재 설정된 트리거:\n\n';

  if (triggers.length === 0) {
    message += '⚠️ 설정된 트리거가 없습니다!\n\n';
    message += '메뉴: 🏥 약재관리 > 🔧 관리 > ⚙️ 모든 트리거 설정\n을 실행하세요.';
  } else {
    const triggerInfo = {};

    triggers.forEach(trigger => {
      const handlerName = trigger.getHandlerFunction();
      const eventType = trigger.getEventType();

      if (!triggerInfo[handlerName]) {
        triggerInfo[handlerName] = [];
      }

      if (eventType === ScriptApp.EventType.ON_EDIT) {
        triggerInfo[handlerName].push('편집 시 실행');
      } else if (eventType === ScriptApp.EventType.CLOCK) {
        const source = trigger.getTriggerSource();
        if (source === ScriptApp.TriggerSource.CLOCK) {
          triggerInfo[handlerName].push('시간 기반');
        }
      } else if (eventType === ScriptApp.EventType.ON_OPEN) {
        triggerInfo[handlerName].push('시트 열 때 실행');
      }
    });

    for (let func in triggerInfo) {
      message += `✅ ${func}: ${triggerInfo[func].join(', ')}\n`;
    }

    message += `\n총 ${triggers.length}개 트리거 실행 중`;

    // onEditHandler 확인
    if (triggerInfo['onEditHandler']) {
      message += '\n\n✅ onEditHandler 트리거 정상';
    } else {
      message += '\n\n⚠️ onEditHandler 트리거 없음!\n조제완료, 입고완료, 재고조정이 작동하지 않습니다.';
    }
  }

  Browser.msgBox('트리거 상태', message, Browser.Buttons.OK);
}

/**
 * 처방 조제 테스트 (수동 실행)
 */
function testPrescriptionProcessing() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('처방상세');

  if (!sheet) {
    Browser.msgBox('오류', '처방상세 시트가 없습니다.', Browser.Buttons.OK);
    return;
  }

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    Browser.msgBox('안내', '처방상세 시트에 조제할 항목이 없습니다.', Browser.Buttons.OK);
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '조제 테스트',
    `처방상세 시트의 몇 번째 행을 조제 처리하시겠습니까?\n(2~${lastRow}):`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const row = parseInt(response.getResponseText());

  if (isNaN(row) || row < 2 || row > lastRow) {
    Browser.msgBox('오류', `2~${lastRow} 사이의 숫자를 입력하세요.`, Browser.Buttons.OK);
    return;
  }

  try {
    Logger.log('=== 수동 조제 테스트 시작 ===');
    processPrescriptionDispense(row);
    Browser.msgBox('성공', `${row}행 조제 처리가 완료되었습니다!\n\n약재출고 및 FIFO상세추적 시트를 확인하세요.`, Browser.Buttons.OK);
  } catch (error) {
    Logger.log('❌ 조제 테스트 실패: ' + error.message);
    Logger.log(error.stack);
    Browser.msgBox('조제 처리 오류', error.message, Browser.Buttons.OK);
  }
}
