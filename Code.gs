const SHEET_NAME = 'Waitlist';
const HEADERS = ['暱稱', 'Email', '需求', '年齡', '推薦人', '推薦人數', '時間'];

function doGet(e) {
  try {
    const action = String(e && e.parameter && e.parameter.action || '').trim();

    if (action === 'getLeaderboard') {
      return jsonOutput_(getLeaderboard_());
    }

    if (action === 'checkNickname') {
      const name = String(e && e.parameter && e.parameter.name || '').trim();
      return jsonOutput_(name ? nicknameExists_(getWaitlistSheet_(), name) : false);
    }

    return jsonOutput_({ success: false, message: '不支援的操作。' });
  } catch (error) {
    console.error('doGet failed', error);
    return jsonOutput_({ success: false, message: '服務暫時無法使用，請稍後再試。' });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(10000)) {
      return jsonOutput_({ success: false, message: '目前報名人數較多，請稍後再試。' });
    }

    const params = e && e.parameter ? e.parameter : {};
    const nickname = cleanText_(params.nickname, 24);
    const email = cleanText_(params.email, 254);
    const need = cleanText_(params.need, 1000);
    const ageText = cleanText_(params.age, 3);
    const referrer = cleanText_(params.referrer, 24);
    if (!nickname) {
      return jsonOutput_({ success: false, message: '請填寫暱稱。' });
    }

    if (!isValidEmail_(email)) {
      return jsonOutput_({ success: false, message: '請填寫有效的 Email。' });
    }

    let age = '';
    if (ageText) {
      age = Number(ageText);
      if (!Number.isInteger(age) || age < 1 || age > 120) {
        return jsonOutput_({ success: false, message: '年齡需為 1 到 120 的整數。' });
      }
    }

    const sheet = getWaitlistSheet_();
    if (nicknameExists_(sheet, nickname)) {
      return jsonOutput_({
        success: false,
        code: 'DUPLICATE_NICKNAME',
        message: '這個暱稱已經被拿走了'
      });
    }

    let referrerCredited = false;
    if (referrer && normalizeName_(referrer) !== normalizeName_(nickname)) {
      const referrerRow = findNicknameRow_(sheet, referrer);
      if (referrerRow > 1) {
        const countCell = sheet.getRange(referrerRow, 6);
        const currentCount = Number(countCell.getValue()) || 0;
        countCell.setValue(currentCount + 1);
        referrerCredited = true;
      }
    }

    const nextRow = sheet.getLastRow() + 1;
    const textValues = [[
      neutralizeFormula_(nickname),
      neutralizeFormula_(email),
      neutralizeFormula_(need),
      age,
      neutralizeFormula_(referrer),
      0,
      new Date()
    ]];
    sheet.getRange(nextRow, 1, 1, HEADERS.length).setValues(textValues);
    SpreadsheetApp.flush();

    return jsonOutput_({
      success: true,
      referrerCredited: referrerCredited
    });
  } catch (error) {
    console.error('doPost failed', error);
    return jsonOutput_({ success: false, message: '報名失敗，請稍後再試。' });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function getLeaderboard_() {
  const sheet = getWaitlistSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getDisplayValues();
  return rows
    .map(function(row, index) {
      return {
        '暱稱': String(row[0] || '').trim(),
        '推薦人數': Math.max(0, Number(row[5]) || 0),
        rowOrder: index
      };
    })
    .filter(function(entry) {
      return Boolean(entry['暱稱']);
    })
    .sort(function(a, b) {
      return b['推薦人數'] - a['推薦人數'] || a.rowOrder - b.rowOrder;
    })
    .slice(0, 10)
    .map(function(entry) {
      return {
        '暱稱': entry['暱稱'],
        '推薦人數': entry['推薦人數']
      };
    });
}

function getWaitlistSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('找不到綁定的 Google 試算表。');

  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('找不到工作表：' + SHEET_NAME);

  const actualHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
  const matches = HEADERS.every(function(header, index) {
    return String(actualHeaders[index] || '').trim() === header;
  });
  if (!matches) {
    throw new Error('第一列欄位必須依序為：' + HEADERS.join('、'));
  }

  return sheet;
}

function nicknameExists_(sheet, name) {
  return findNicknameRow_(sheet, name) > 1;
}

function findNicknameRow_(sheet, name) {
  const target = normalizeName_(name);
  const lastRow = sheet.getLastRow();
  if (!target || lastRow < 2) return -1;

  const names = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  for (let index = 0; index < names.length; index += 1) {
    if (normalizeName_(names[index][0]) === target) return index + 2;
  }
  return -1;
}

function normalizeName_(value) {
  return String(value || '').trim().toLocaleLowerCase();
}

function cleanText_(value, maxLength) {
  return String(value || '').trim().slice(0, maxLength);
}

function neutralizeFormula_(value) {
  const text = String(value || '');
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function isValidEmail_(value) {
  return value.length <= 254 && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function jsonOutput_(data) {
  // Apps Script 的 TextOutput 無法自行加入 Access-Control-Allow-Origin。
  // 公開部署後由 Google 提供跨網域回應；前端使用 URLSearchParams，避免觸發 POST 預檢。
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
