/**
 * 수시 지원 입력 수신기 (가벼운 버전)
 * - 학생 HTML에서 POST 요청 받음
 * - responses 시트에 저장만 함
 * - db 시트는 안 읽음 (학생 화면에서 처리하니까)
 * - 30초 타임아웃 걱정 없음
 */

const SHEET_RESPONSES = 'responses';

// ========== 웹앱 진입점 (POST 수신) ==========
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    if (!data.hakbun) {
      return ContentService.createTextOutput(JSON.stringify({
        ok: false, msg: '학번이 없습니다.'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_RESPONSES);
    
    // 시트 없으면 자동 생성
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_RESPONSES);
      sheet.getRange(1, 1, 1, 7).setValues([
        ['학번', '슬롯', '대학명', '세부전형', '모집단위', '제출시각', '수정시각']
      ]);
      sheet.getRange(1, 1, 1, 7)
        .setFontWeight('bold')
        .setBackground('#3a6647')
        .setFontColor('white');
      sheet.setFrozenRows(1);
    }
    
    const hakbun = String(data.hakbun);
    const now = new Date();
    
    // 이 학생의 기존 행 삭제 (재제출이면)
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const hakbunCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      const rowsToDelete = [];
      for (let i = hakbunCol.length - 1; i >= 0; i--) {
        if (String(hakbunCol[i][0]) === hakbun) {
          rowsToDelete.push(i + 2);
        }
      }
      rowsToDelete.forEach(r => sheet.deleteRow(r));
    }
    
    // 새 데이터 추가
    const slots = data.slots || [];
    if (slots.length > 0) {
      const rows = slots.map(s => [
        hakbun,
        s.slot,
        s.univ || '',
        s.type || '',
        s.dept || '',
        now,
        now
      ]);
      
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      ok: true,
      msg: '저장 완료',
      count: slots.length
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      msg: '오류: ' + err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== GET 요청 (테스트용) ==========
function doGet(e) {
  return ContentService.createTextOutput(
    'Susi Response Receiver is running. POST data here to save.'
  ).setMimeType(ContentService.MimeType.TEXT);
}

// ========== 테스트 (Apps Script 편집기에서 직접 실행) ==========
function 테스트_제출() {
  const testData = {
    postData: {
      contents: JSON.stringify({
        hakbun: '99999',
        slots: [
          { slot: 1, univ: '고려대', type: '학생부종합(계열적합형)', dept: '통계학과' },
          { slot: 2, univ: '성균관대', type: '학생부종합', dept: '자연과학계열' }
        ],
        timestamp: new Date().toISOString()
      })
    }
  };
  
  const result = doPost(testData);
  Logger.log(result.getContent());
  SpreadsheetApp.getUi().alert('테스트 완료. responses 시트 확인.');
}
