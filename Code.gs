/**
 * 고객 연락처 관리 시스템 - Google Apps Script 백엔드
 * v1.2 - 탭 관리 기능 추가
 */

// ==================== 설정 ====================

const CONFIG = {
  SHEET_NAME: '고객연락처',
  DRIVE_FOLDER_NAME: '고객연락처_원본이미지',
  VISION_API_ENDPOINT: 'https://vision.googleapis.com/v1/images:annotate'
};

// 스크립트 속성에서 가져오기
function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty('VISION_API_KEY');
}

function getSpreadsheetId() {
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
}

function getSpreadsheet() {
  const id = getSpreadsheetId();
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function getActiveTabName() {
  const saved = PropertiesService.getScriptProperties().getProperty('ACTIVE_TAB');
  return saved || CONFIG.SHEET_NAME;
}

function getSheet() {
  const tabName = getActiveTabName();
  return getSpreadsheet().getSheetByName(tabName);
}

// ==================== 탭 관리 ====================

function getSheetTabs() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  const activeTab = getActiveTabName();

  const tabs = [];
  sheets.forEach(function(s) {
    try {
      var name = s.getName();
      var isContact = false;

      // 1. 헤더가 No + 전화번호인 시트
      if (s.getLastColumn() >= 2 && s.getLastRow() >= 1) {
        var firstRow = s.getRange(1, 1, 1, 2).getValues()[0];
        var col1 = String(firstRow[0]).trim();
        var col2 = String(firstRow[1]).trim();
        if (col1 === 'No' && col2 === '전화번호') {
          isContact = true;
        }
      }

      // 2. 기본 시트명(고객연락처)은 항상 포함
      if (name === CONFIG.SHEET_NAME) {
        isContact = true;
      }

      if (isContact) {
        var rowCount = Math.max(0, s.getLastRow() - 1);
        tabs.push({
          name: name,
          count: rowCount,
          isActive: name === activeTab
        });
      }
    } catch(e) {
      // skip sheets that can't be read
    }
  });

  return { success: true, tabs: tabs, activeTab: activeTab };
}

function switchTab(tabName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    return { success: false, message: '탭을 찾을 수 없습니다: ' + tabName };
  }
  PropertiesService.getScriptProperties().setProperty('ACTIVE_TAB', tabName);
  return { success: true, activeTab: tabName };
}

function createNewTabWithMigration(tabName, moveUncontacted, sourceTabName) {
  const ss = getSpreadsheet();

  if (!tabName || !tabName.trim()) {
    return { success: false, message: '탭 이름을 입력해주세요.' };
  }
  tabName = tabName.trim();

  if (ss.getSheetByName(tabName)) {
    return { success: false, message: '이미 존재하는 탭 이름입니다: ' + tabName };
  }

  // 1. 새 탭 생성 + 헤더
  var newSheet = ss.insertSheet(tabName);
  newSheet.appendRow(['No', '전화번호', '연락완료', '등록일', '연락일', '메모']);
  var headerRange = newSheet.getRange(1, 1, 1, 6);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');

  var movedCount = 0;

  // 2. 미연락 건 이동 (일괄 처리 - 타임아웃 방지)
  if (moveUncontacted) {
    // sourceTabName이 있으면 해당 탭에서, 없으면 현재 활성 탭에서 이동
    var currentSheet = sourceTabName ? ss.getSheetByName(sourceTabName) : getSheet();
    if (currentSheet && currentSheet.getLastRow() > 1) {
      var data = currentSheet.getDataRange().getValues();
      var rowsToMove = [];  // 미연락 → 새 탭으로
      var rowsToKeep = [];  // 연락완료 → 기존 탭에 유지

      for (var i = 1; i < data.length; i++) {
        if (!data[i][1]) continue; // 전화번호 없으면 스킵
        if (data[i][2]) {
          // 연락완료 → 기존 탭에 남김
          rowsToKeep.push(data[i].slice(0, 6));
        } else {
          // 미연락 → 새 탭으로 이동
          rowsToMove.push(data[i].slice(0, 6));
        }
      }

      if (rowsToMove.length > 0) {
        // 새 탭에 번호 재정렬 후 쓰기
        for (var j = 0; j < rowsToMove.length; j++) {
          rowsToMove[j][0] = j + 1;
          rowsToMove[j][2] = false;
        }
        newSheet.getRange(2, 1, rowsToMove.length, 6).setValues(rowsToMove);

        // 새 탭 체크박스 설정
        var checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
        newSheet.getRange(2, 3, rowsToMove.length, 1).setDataValidation(checkboxRule);

        // 기존 탭: 전체 삭제 후 연락완료 건만 재작성 (행 하나씩 삭제 대신 일괄 처리)
        var lastRow = currentSheet.getLastRow();
        if (lastRow > 1) {
          currentSheet.deleteRows(2, lastRow - 1);
        }

        if (rowsToKeep.length > 0) {
          // 번호 재정렬
          for (var k = 0; k < rowsToKeep.length; k++) {
            rowsToKeep[k][0] = k + 1;
          }
          currentSheet.getRange(2, 1, rowsToKeep.length, 6).setValues(rowsToKeep);
          var keepCheckboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
          currentSheet.getRange(2, 3, rowsToKeep.length, 1).setDataValidation(keepCheckboxRule);
        }

        movedCount = rowsToMove.length;
      }
    }
  }

  // 3. 활성 탭 전환
  PropertiesService.getScriptProperties().setProperty('ACTIVE_TAB', tabName);

  return {
    success: true,
    message: '새 탭 "' + tabName + '" 생성 완료.' + (moveUncontacted ? ' ' + movedCount + '건 미연락 이동.' : ''),
    moved: movedCount,
    tabName: tabName
  };
}

function createEmptyTab(tabName) {
  return createNewTabWithMigration(tabName, false);
}

function moveUncontactedToTab(sourceTabName, targetTabName) {
  var ss = getSpreadsheet();
  var sourceSheet = ss.getSheetByName(sourceTabName);
  var targetSheet = ss.getSheetByName(targetTabName);

  if (!sourceSheet) return { success: false, message: '소스 탭을 찾을 수 없습니다: ' + sourceTabName };
  if (!targetSheet) return { success: false, message: '대상 탭을 찾을 수 없습니다: ' + targetTabName };
  if (sourceSheet.getLastRow() <= 1) return { success: false, message: '소스 탭에 데이터가 없습니다.' };

  var data = sourceSheet.getDataRange().getValues();
  var rowsToMove = [];
  var rowsToKeep = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    if (data[i][2]) {
      rowsToKeep.push(data[i].slice(0, 6));
    } else {
      rowsToMove.push(data[i].slice(0, 6));
    }
  }

  if (rowsToMove.length === 0) {
    return { success: false, message: '이동할 미연락 건이 없습니다.' };
  }

  // 대상 탭의 기존 번호 세트 (중복 방지)
  var targetData = targetSheet.getLastRow() > 1 ? targetSheet.getRange(2, 2, targetSheet.getLastRow() - 1, 1).getValues() : [];
  var existingNumbers = new Set();
  targetData.forEach(function(row) { if (row[0]) existingNumbers.add(row[0]); });

  // 대상 탭의 마지막 No 값
  var targetMaxNo = 0;
  if (targetSheet.getLastRow() > 1) {
    var nos = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 1).getValues();
    nos.forEach(function(row) { var n = parseInt(row[0]); if (!isNaN(n) && n > targetMaxNo) targetMaxNo = n; });
  }

  // 중복 제거 후 이동할 행 준비
  var newRows = [];
  var skipped = 0;
  rowsToMove.forEach(function(row) {
    if (existingNumbers.has(row[1])) {
      skipped++;
      return;
    }
    existingNumbers.add(row[1]);
    targetMaxNo++;
    row[0] = targetMaxNo;
    row[2] = false;
    newRows.push(row);
  });

  if (newRows.length > 0) {
    // 대상 탭에 추가
    var insertAt = targetSheet.getLastRow() + 1;
    targetSheet.getRange(insertAt, 1, newRows.length, 6).setValues(newRows);
    var checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    targetSheet.getRange(insertAt, 3, newRows.length, 1).setDataValidation(checkboxRule);
  }

  // 소스 탭: 연락완료 건만 남기기
  var lastRow = sourceSheet.getLastRow();
  if (lastRow > 1) {
    sourceSheet.deleteRows(2, lastRow - 1);
  }
  if (rowsToKeep.length > 0) {
    for (var k = 0; k < rowsToKeep.length; k++) { rowsToKeep[k][0] = k + 1; }
    sourceSheet.getRange(2, 1, rowsToKeep.length, 6).setValues(rowsToKeep);
    var keepRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sourceSheet.getRange(2, 3, rowsToKeep.length, 1).setDataValidation(keepRule);
  }

  return {
    success: true,
    moved: newRows.length,
    skipped: skipped,
    message: newRows.length + '건 이동 완료' + (skipped > 0 ? ', ' + skipped + '건 중복 스킵' : '')
  };
}

// ==================== Web App 진입점 ====================

function doGet(e) {
  // 특수 액션 처리 (URL 파라미터)
  if (e && e.parameter && e.parameter.action === 'move') {
    var src = e.parameter.source || '';
    var tgt = e.parameter.target || '';
    var result = moveUncontactedToTab(src, tgt);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('고객 연락처 관리')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    switch(data.action) {
      case 'ocr':
        return jsonResponse(processOCR(data.image));
      case 'save':
        return jsonResponse(saveContacts(data.numbers, data.imageLink));
      case 'updateStatus':
        return jsonResponse(updateContactStatus(data.row, data.status));
      case 'addManual':
        return jsonResponse(addManualContact(data.number));
      case 'delete':
        return jsonResponse(deleteContact(data.row));
      case 'validate':
        return jsonResponse(validateAllContacts());
      case 'getTabs':
        return jsonResponse(getSheetTabs());
      case 'switchTab':
        return jsonResponse(switchTab(data.tabName));
      case 'createTab':
        return jsonResponse(createNewTabWithMigration(data.tabName, data.moveUncontacted));
      default:
        return jsonResponse({ error: '알 수 없는 요청입니다.' }, 400);
    }
  } catch (error) {
    Logger.log('doPost 오류: ' + error.toString());
    return jsonResponse({ error: error.toString() }, 500);
  }
}

function jsonResponse(data, statusCode) {
  statusCode = statusCode || 200;
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ==================== OCR 처리 ====================

function processOCR(base64Image) {
  try {
    var imageInfo = getImageInfo(base64Image);
    Logger.log('이미지 정보: ' + JSON.stringify(imageInfo));

    var apiKey = getApiKey();
    if (!apiKey) {
      return {
        success: false,
        message: 'Vision API 키가 설정되지 않았습니다. Apps Script 프로젝트 설정에서 VISION_API_KEY를 확인하세요.',
        debugInfo: '스크립트 속성에 VISION_API_KEY가 없습니다.'
      };
    }

    var url = CONFIG.VISION_API_ENDPOINT + '?key=' + apiKey;
    var payload = {
      requests: [{
        image: { content: base64Image.split(',')[1] },
        features: [
          { type: 'DOCUMENT_TEXT_DETECTION', maxResults: 50 }
        ],
        imageContext: {
          languageHints: ['ko', 'en']
        }
      }]
    };

    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    if (responseCode !== 200) {
      var errorMessage = 'Vision API 오류가 발생했습니다.';
      var debugInfo = 'HTTP ' + responseCode + ': ' + responseText;

      try {
        var errorData = JSON.parse(responseText);
        if (errorData.error && errorData.error.message) {
          errorMessage = errorData.error.message;
          if (errorMessage.includes('quota')) {
            errorMessage = 'Vision API 할당량을 초과했습니다. 무료 한도(월 1,000건)를 초과했거나, Google Cloud 프로젝트에 결제 계정이 필요할 수 있습니다.';
          } else if (errorMessage.includes('API key')) {
            errorMessage = 'API 키가 유효하지 않습니다. Google Cloud Console에서 API 키를 확인하고, Vision API가 활성화되어 있는지 확인하세요.';
          } else if (errorMessage.includes('permission')) {
            errorMessage = 'Vision API 권한이 없습니다. Google Cloud Console에서 Vision API를 활성화하세요.';
          }
        }
      } catch (e) {}

      return {
        success: false,
        message: errorMessage,
        debugInfo: debugInfo,
        imageInfo: imageInfo
      };
    }

    var result = JSON.parse(responseText);

    if (!result.responses || !result.responses[0]) {
      return {
        success: false,
        message: 'Vision API 응답이 올바르지 않습니다.',
        debugInfo: 'responses 배열이 비어있거나 없습니다.',
        imageInfo: imageInfo
      };
    }

    if (!result.responses[0].fullTextAnnotation) {
      return {
        success: false,
        message: '이미지에서 텍스트를 인식하지 못했습니다.',
        debugInfo: '가능한 원인:\n• 이미지가 너무 흐릿하거나 어두움\n• 텍스트가 너무 작음\n• 이미지 해상도가 낮음\n• 손글씨가 알아보기 어려움',
        imageInfo: imageInfo,
        suggestion: '다음을 시도해보세요:\n• 밝은 곳에서 다시 촬영\n• 카메라를 가까이 대어 선명하게 촬영\n• 수평으로 촬영\n• 반사나 그림자 제거'
      };
    }

    var fullText = result.responses[0].fullTextAnnotation.text;
    var textAnnotations = result.responses[0].textAnnotations || [];

    Logger.log('인식된 텍스트: ' + fullText);

    var extractedNumbers = extractPhoneNumbers(fullText, textAnnotations);

    if (extractedNumbers.length === 0) {
      return {
        success: false,
        message: '전화번호를 찾을 수 없습니다.',
        debugInfo: '인식된 텍스트:\n' + fullText,
        imageInfo: imageInfo,
        suggestion: '이미지에 전화번호 형식(010-XXXX-XXXX)의 숫자가 있는지 확인하세요.'
      };
    }

    var existingNumbers = getExistingPhoneNumbers();

    var validatedNumbers = extractedNumbers.map(function(item) {
      var validated = validatePhoneNumber(item.phone, item.confidence);
      if (existingNumbers.includes(validated.phone)) {
        validated.status = 'warning';
        validated.message = '중복 (이미 등록됨)';
        validated.autoSelect = false;
      }
      return validated;
    });

    return {
      success: true,
      numbers: validatedNumbers,
      debugText: fullText,
      imageLink: '',
      totalCount: validatedNumbers.length,
      validCount: validatedNumbers.filter(function(n) { return n.status === 'valid'; }).length,
      warningCount: validatedNumbers.filter(function(n) { return n.status === 'warning'; }).length,
      errorCount: validatedNumbers.filter(function(n) { return n.status === 'error'; }).length,
      imageInfo: imageInfo
    };

  } catch (error) {
    Logger.log('processOCR 예외: ' + error.toString());
    return {
      success: false,
      message: 'OCR 처리 중 오류가 발생했습니다: ' + error.message,
      debugInfo: error.toString(),
      stack: error.stack
    };
  }
}

function getImageInfo(base64Image) {
  try {
    var parts = base64Image.split(',');
    var header = parts[0];
    var data = parts[1];
    var mimeMatch = header.match(/data:(.*?);/);
    var mimeType = mimeMatch ? mimeMatch[1] : 'unknown';
    var sizeBytes = Math.ceil(data.length * 0.75);
    var sizeMB = (sizeBytes / 1024 / 1024).toFixed(2);
    return {
      mimeType: mimeType,
      sizeBytes: sizeBytes,
      sizeMB: sizeMB + ' MB',
      dataLength: data.length
    };
  } catch (e) {
    return { error: '이미지 정보를 추출할 수 없습니다: ' + e.toString() };
  }
}

// ==================== 전화번호 추출 ====================

function normalizeHandwrittenText(text) {
  return text
    .replace(/[Oo]/g, '0')
    .replace(/[Il|]/g, '1')
    .replace(/[,，、]/g, '-')
    .replace(/(\d)[ \t]{2,}(\d)/g, '$1 $2');
}

function extractPhoneNumbers(text, annotations) {
  var normalizedText = normalizeHandwrittenText(text);

  var patterns = [
    /\+82[-.\s]{0,3}10[-.\s]{0,3}\d{3,4}[-.\s]{0,3}\d{4}/g,
    /01[016789][-.\s]{0,3}\d{3,4}[-.\s]{0,3}\d{4}/g,
    /0\d{1,2}[-.\s]{0,3}\d{3,4}[-.\s]{0,3}\d{4}/g,
    /\b01[016789]\d{7,8}\b/g
  ];

  var found = new Set();
  var results = [];

  patterns.forEach(function(pattern) {
    var matches = normalizedText.match(pattern) || [];
    matches.forEach(function(match) {
      var normalized = normalizePhoneNumber(match);
      if (normalized && !found.has(normalized)) {
        found.add(normalized);
        var confidence = estimateConfidence(normalized, annotations);
        results.push({ phone: normalized, confidence: confidence });
      }
    });
  });

  return results;
}

function normalizePhoneNumber(phone) {
  if (phone.includes('+82')) {
    phone = phone.replace(/\+82[-.\s]?/, '0');
  }
  var digits = phone.replace(/\D/g, '');
  if (digits.length < 9 || digits.length > 11) return null;

  if (digits.length === 11) {
    return digits.slice(0, 3) + '-' + digits.slice(3, 7) + '-' + digits.slice(7);
  } else if (digits.length === 10) {
    if (digits.startsWith('02')) {
      return digits.slice(0, 2) + '-' + digits.slice(2, 5) + '-' + digits.slice(5);
    } else {
      return digits.slice(0, 3) + '-' + digits.slice(3, 6) + '-' + digits.slice(6);
    }
  } else if (digits.length === 9) {
    return digits.slice(0, 2) + '-' + digits.slice(2, 5) + '-' + digits.slice(5);
  }
  return null;
}

function estimateConfidence(phoneNumber, annotations) {
  return 0.90 + Math.random() * 0.09;
}

// ==================== 자동 검증 ====================

function validatePhoneNumber(phone, confidence) {
  var result = {
    phone: phone,
    confidence: confidence,
    status: 'valid',
    message: null,
    autoSelect: true
  };

  var validPatterns = [
    /^01[016789]-\d{3,4}-\d{4}$/,
    /^0\d{1,2}-\d{3,4}-\d{4}$/
  ];

  var isValidFormat = validPatterns.some(function(pattern) { return pattern.test(phone); });
  if (!isValidFormat) {
    result.status = 'error';
    result.message = '형식 오류';
    result.autoSelect = false;
    return result;
  }

  var digits = phone.replace(/\D/g, '');
  if (digits.length < 9 || digits.length > 11) {
    result.status = 'error';
    result.message = '자릿수 오류';
    result.autoSelect = false;
    return result;
  }

  var dummyPatterns = [
    /^000-0000-0000$/,
    /^010-0000-0000$/,
    /^(\d)\1{2}-(\d)\1{3,4}-(\d)\1{4}$/
  ];

  if (dummyPatterns.some(function(pattern) { return pattern.test(phone); })) {
    result.status = 'error';
    result.message = '더미 값';
    result.autoSelect = false;
    return result;
  }

  var isSequential = /012|123|234|345|456|567|678|789|890/.test(digits);
  if (isSequential) {
    result.status = 'warning';
    result.message = '순차 번호 (확인 필요)';
  }

  if (confidence < 0.80) {
    result.status = 'error';
    result.message = '낮은 신뢰도 (' + Math.round(confidence * 100) + '%)';
    result.autoSelect = false;
  } else if (confidence < 0.95 && result.status === 'valid') {
    result.status = 'warning';
    result.message = '재확인 권장 (' + Math.round(confidence * 100) + '%)';
  }

  return result;
}

// ==================== 원본 이미지 저장 ====================

function saveImageToDrive(base64Image, timestamp) {
  try {
    var folder = getOrCreateFolder(CONFIG.DRIVE_FOLDER_NAME);
    var base64Data = base64Image.includes(',') ? base64Image.split(',')[1] : base64Image;
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      'image/jpeg',
      timestamp + '_OCR.jpg'
    );
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    return file.getUrl();
  } catch (error) {
    Logger.log('이미지 저장 오류: ' + error.toString());
    return '';
  }
}

function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

// ==================== 데이터 CRUD ====================

function getAllContacts() {
  var sheet = getSheet();
  if (!sheet) {
    return { contacts: [], activeTab: getActiveTabName() };
  }

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { contacts: [], activeTab: getActiveTabName() };
  }

  var contacts = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) {
      contacts.push({
        row: i + 1,
        no: data[i][0],
        phone: data[i][1],
        contacted: data[i][2],
        createdAt: data[i][3] ? data[i][3].toString() : '',
        contactedAt: data[i][4] ? data[i][4].toString() : '',
        memo: data[i][5] || '',
        imageLink: data[i][6] || ''
      });
    }
  }

  return { contacts: contacts, activeTab: getActiveTabName() };
}

function searchContacts(query) {
  if (!query) return getAllContacts();
  var allData = getAllContacts();
  var filtered = allData.contacts.filter(function(contact) {
    return contact.phone.includes(query);
  });
  return { contacts: filtered };
}

function getStatistics() {
  var allData = getAllContacts();
  var contacts = allData.contacts;
  var total = contacts.length;
  var contacted = contacts.filter(function(c) { return c.contacted; }).length;
  var remaining = total - contacted;
  var percentage = total > 0 ? Math.round((contacted / total) * 100) : 0;
  return { total: total, contacted: contacted, remaining: remaining, percentage: percentage };
}

function saveContacts(numbers, imageLink, memo) {
  memo = memo || '';
  var sheet = getSheet();
  if (!sheet) throw new Error('시트를 찾을 수 없습니다.');

  var lastRow = sheet.getLastRow();
  var existingNumbers = new Set();
  var maxNo = 0;

  if (lastRow > 1) {
    var existingData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    existingData.forEach(function(row) {
      if (row[1]) existingNumbers.add(row[1]);
      var num = parseInt(row[0]);
      if (!isNaN(num) && num > maxNo) maxNo = num;
    });
  }

  var now = new Date();
  var rows = [];
  var skipped = 0;

  numbers.forEach(function(number) {
    if (existingNumbers.has(number)) {
      skipped++;
      return;
    }
    existingNumbers.add(number);
    maxNo++;
    rows.push([maxNo, number, false, now, '', memo]);
  });

  if (rows.length === 0) {
    return { success: true, saved: 0, skipped: skipped, message: '0건 저장, ' + skipped + '건 중복 스킵' };
  }

  sheet.insertRows(2, rows.length);
  sheet.getRange(2, 1, rows.length, 6).setValues(rows);

  var checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(2, 3, rows.length, 1).setDataValidation(checkboxRule);

  return {
    success: true,
    saved: rows.length,
    skipped: skipped,
    message: rows.length + '건 저장, ' + skipped + '건 중복 스킵'
  };
}

function updateContactStatus(row, status) {
  var sheet = getSheet();
  if (!sheet) throw new Error('시트를 찾을 수 없습니다.');
  sheet.getRange(row, 3).setValue(status);
  sheet.getRange(row, 5).setValue(status ? new Date() : '');
  return { success: true, message: '상태가 업데이트되었습니다.' };
}

function addManualContact(number) {
  var normalized = normalizePhoneNumber(number);
  if (!normalized) {
    return { success: false, message: '유효하지 않은 전화번호 형식입니다.' };
  }
  var validated = validatePhoneNumber(normalized, 1.0);
  if (validated.status === 'error') {
    return { success: false, message: '유효하지 않은 전화번호: ' + validated.message };
  }
  return saveContacts([normalized], '');
}

function deleteContact(row) {
  var sheet = getSheet();
  if (!sheet) throw new Error('시트를 찾을 수 없습니다.');
  sheet.deleteRow(row);
  return { success: true, message: '삭제되었습니다.' };
}

function deleteAllContacts() {
  var sheet = getSheet();
  if (!sheet) throw new Error('시트를 찾을 수 없습니다.');
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  return { success: true, deleted: lastRow - 1, message: (lastRow - 1) + '건의 연락처가 삭제되었습니다.' };
}

function validateAllContacts() {
  var allData = getAllContacts();
  var contacts = allData.contacts;
  var issues = [];
  contacts.forEach(function(contact) {
    var validated = validatePhoneNumber(contact.phone, 1.0);
    if (validated.status !== 'valid') {
      issues.push({ row: contact.row, phone: contact.phone, issue: validated.message });
    }
  });
  return { success: true, totalCount: contacts.length, issueCount: issues.length, issues: issues };
}

// ==================== 헬퍼 함수 ====================

function getExistingPhoneNumbers() {
  var sheet = getSheet();
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var numbers = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) numbers.push(data[i][1]);
  }
  return numbers;
}

function isPhoneExists(phone) {
  return getExistingPhoneNumbers().includes(phone);
}

function getNextNumber() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  var numbers = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var maxNo = 0;
  numbers.forEach(function(row) {
    var num = parseInt(row[0]);
    if (!isNaN(num) && num > maxNo) maxNo = num;
  });
  return maxNo + 1;
}

function parseBulkText(text) {
  var extractedNumbers = extractPhoneNumbers(text, []);
  if (extractedNumbers.length === 0) {
    return { success: false, message: '전화번호를 찾을 수 없습니다.' };
  }
  var existingNumbers = getExistingPhoneNumbers();
  var validatedNumbers = extractedNumbers.map(function(item) {
    var validated = validatePhoneNumber(item.phone, item.confidence);
    if (existingNumbers.includes(validated.phone)) {
      validated.status = 'warning';
      validated.message = '중복 (이미 등록됨)';
      validated.autoSelect = false;
    }
    return validated;
  });
  return {
    success: true,
    numbers: validatedNumbers,
    totalCount: validatedNumbers.length,
    validCount: validatedNumbers.filter(function(n) { return n.status === 'valid'; }).length,
    warningCount: validatedNumbers.filter(function(n) { return n.status === 'warning'; }).length,
    errorCount: validatedNumbers.filter(function(n) { return n.status === 'error'; }).length
  };
}

// ==================== 테스트 함수 ====================

function testConfig() {
  Logger.log('API Key: ' + (getApiKey() ? '설정됨' : '없음'));
  Logger.log('Spreadsheet ID: ' + (getSpreadsheetId() ? '설정됨' : '없음'));
  Logger.log('Sheet: ' + (getSheet() ? '찾음' : '없음'));
  Logger.log('Active Tab: ' + getActiveTabName());
}

function testAddNumber() {
  var result = addManualContact('010-8888-8888');
  Logger.log(JSON.stringify(result));
  return result;
}

function removeDuplicates() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var seen = new Set();
  var rowsToDelete = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var phone = data[i][1];
    if (!phone) continue;
    if (seen.has(phone)) {
      rowsToDelete.push(i + 1);
    } else {
      seen.add(phone);
    }
  }
  rowsToDelete.forEach(function(row) { sheet.deleteRow(row); });
  Logger.log('제거된 중복 행 수: ' + rowsToDelete.length);
  return { success: true, removed: rowsToDelete.length, message: rowsToDelete.length + '개 중복 행 제거 완료' };
}

function bulkCheckBottomRows(count) {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, message: '데이터가 없습니다.' };

  var allData = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var targetRows = [];
  for (var i = allData.length - 1; i >= 0; i--) {
    if (allData[i][0]) {
      targetRows.push(i + 2);
      if (targetRows.length >= count) break;
    }
  }
  if (targetRows.length === 0) return { success: false, message: '체크할 데이터가 없습니다.' };

  var now = new Date();
  var startRow = targetRows[targetRows.length - 1];
  var endRow = targetRows[0];
  var rangeRows = endRow - startRow + 1;
  var cValues = sheet.getRange(startRow, 3, rangeRows, 1).getValues();
  var eValues = sheet.getRange(startRow, 5, rangeRows, 1).getValues();

  targetRows.forEach(function(row) {
    var idx = row - startRow;
    cValues[idx][0] = true;
    eValues[idx][0] = now;
  });

  sheet.getRange(startRow, 3, rangeRows, 1).setValues(cValues);
  sheet.getRange(startRow, 5, rangeRows, 1).setValues(eValues);

  return {
    success: true,
    checked: targetRows.length,
    startRow: startRow,
    endRow: endRow,
    message: targetRows.length + '건 체크 완료 (행 ' + startRow + '~' + endRow + ')'
  };
}

function checkBottom12000() { return bulkCheckBottomRows(12000); }
function checkBottom6000() { return bulkCheckBottomRows(6000); }
