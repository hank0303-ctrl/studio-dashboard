const SHEET_NAME = 'studio_data';
const DEFAULT_APP_ID = 'studio-dashboard';
const DEFAULT_SPREADSHEET_ID = '1-DtjW2azPOspnc5XLbzGunN4sqMbkyDaErrXrEMUllQ';

function doGet(e) {
  try {
    const params = e.parameter || {};
    assertAllowed(params.token || '');
    const action = params.action || 'load';
    if (action === 'load') {
      return jsonOutput(loadAppData(params.app || DEFAULT_APP_ID));
    }
    if (action === 'gmailPurchases') {
      return jsonOutput(getGmailPurchases(params));
    }
    return jsonOutput({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    assertAllowed(body.token || '');
    if (body.action !== 'save') {
      return jsonOutput({ ok: false, error: 'Unknown action' });
    }
    return jsonOutput(saveAppData(body.app || DEFAULT_APP_ID, body.data || {}, body.meta || {}));
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message });
  }
}

function saveAppData(appId, data, meta) {
  const sheet = getDataSheet();
  const values = sheet.getDataRange().getValues();
  const now = new Date().toISOString();
  const nextMeta = Object.assign({}, meta, { cloudUpdatedAt: now });
  const row = [appId, JSON.stringify(data), JSON.stringify(nextMeta), now];

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === appId) {
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { ok: true, data, meta: nextMeta };
    }
  }

  sheet.appendRow(row);
  return { ok: true, data, meta: nextMeta };
}

function loadAppData(appId) {
  const sheet = getDataSheet();
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === appId) {
      return {
        ok: true,
        data: safeJson(values[i][1], {}),
        meta: safeJson(values[i][2], {}),
      };
    }
  }

  return { ok: true, data: {}, meta: {} };
}

function getGmailPurchases(params) {
  const query = params.query || 'newer_than:180d from:wushujoyful@gmail.com subject:每日營收';
  const max = Math.min(Number(params.max || 150), 300);
  const messages = uniqueMessages(searchGmailMessages(query, max));
  const purchases = [];

  const selectedMessages = selectBestDailyRevenueMessages(messages);
  selectedMessages.forEach(message => {
    purchases.push(...parseDailyRevenueMail(message, message.getSubject() || '', message.getPlainBody() || ''));
  });
  const combined = dedupePurchases(purchases);

  return {
    ok: true,
    query,
    forcedQuery: query,
    messageCount: messages.length,
    selectedMessageCount: selectedMessages.length,
    sheetPurchaseCount: 0,
    sheetNames: [],
    count: combined.length,
    purchases: combined,
    subjects: selectedMessages.slice(0, 12).map(message => message.getSubject()),
  };
}

function getSheetPurchases(params) {
  const purchases = [];
  const sheetNames = [];
  const ss = getSpreadsheet(params);
  ss.getSheets().forEach(sheet => {
    if (sheet.getName() === SHEET_NAME) return;
    const values = sheet.getDataRange().getDisplayValues();
    if (!values.length) return;
    const found = parsePurchaseRowsFromValues(values, sheet.getName());
    if (found.length) {
      sheetNames.push(sheet.getName());
      purchases.push(...found);
    }
  });
  return { purchases, sheetNames };
}

function parsePurchaseRowsFromValues(values, sheetName) {
  const rows = [];
  for (let headerRow = 0; headerRow < Math.min(values.length, 12); headerRow++) {
    const headers = values[headerRow].map(h => cleanValue(h));
    const col = makeColumnFinder(headers);
    const studentCol = col(['會員姓名', '學員姓名', '學生姓名', '姓名', '購買人', '會員名稱']);
    const planCol = col(['商品名稱', '方案名稱', '購買方案', '票券名稱', '課程方案']);
    if (studentCol < 0 || planCol < 0) continue;

    const dateCol = col(['訂購時間', '購買時間', '付款時間', '購買日期', '訂單日期', '時間戳記', 'Timestamp']);
    const amountCol = col(['付款內容', '付款金額', '總金額', '金額', '實收金額']);
    const periodCol = col(['使用期限', '有效期限']);
    const startCol = col(['啟用日期', '開始日期', '起始日期']);
    const endCol = col(['到期日期', '結束日期', '截止日期']);
    const payMethodCol = col(['付款方式']);
    const typeCol = col(['票券種類', '商品類型', '部門']);
    const codeCol = col(['商品編號', '訂單編號', '編號']);

    for (let r = headerRow + 1; r < values.length; r++) {
      const row = values[r];
      const student = cleanValue(row[studentCol]);
      const plan = cleanValue(row[planCol]);
      if (!student || !plan) continue;

      const period = parseDateRange(periodCol >= 0 ? row[periodCol] : '');
      const date = normalizeDate(dateCol >= 0 ? row[dateCol] : '') || formatDate(new Date());
      const startDate = (startCol >= 0 ? normalizeDate(row[startCol]) : '') || period.startDate || date;
      const endDate = (endCol >= 0 ? normalizeDate(row[endCol]) : '') || period.endDate || '';
      const amountText = amountCol >= 0 ? row[amountCol] : '';
      const typeText = typeCol >= 0 ? row[typeCol] : '';
      const code = codeCol >= 0 ? cleanValue(row[codeCol]) : '';

      rows.push({
        date,
        student,
        plan,
        amount: parseAmount(amountText),
        dept: guessDept(plan + ' ' + typeText + ' ' + sheetName),
        payMethod: guessPayMethod(payMethodCol >= 0 ? row[payMethodCol] : ''),
        serviceMonth: startDate ? startDate.slice(0, 7) : date.slice(0, 7),
        startDate,
        endDate,
        sessions: sessionsFromPlan(plan + ' ' + typeText),
        note: '試算表匯入：' + sheetName + (code ? ' / ' + code : ''),
        sourceId: 'sheet:' + sheetName + ':' + (code || r),
      });
    }
    break;
  }
  return rows;
}

function makeColumnFinder(headers) {
  return function(names) {
    for (const name of names) {
      const needle = String(name).toLowerCase();
      const idx = headers.findIndex(h => String(h).toLowerCase().includes(needle));
      if (idx >= 0) return idx;
    }
    return -1;
  };
}

function dedupePurchases(items) {
  const seen = {};
  const result = [];
  items.forEach(item => {
    const key = [
      item.date || '',
      cleanValue(item.student || ''),
      cleanValue(item.plan || ''),
      Number(item.amount || 0),
      item.startDate || '',
      item.endDate || '',
    ].join('|').toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    result.push(item);
  });
  return result;
}

function searchGmailMessages(query, max) {
  const messages = [];
  GmailApp.search(query, 0, max).forEach(thread => {
    thread.getMessages().forEach(message => messages.push(message));
  });
  return messages;
}

function uniqueMessages(messages) {
  const seen = {};
  const unique = [];
  messages.forEach(message => {
    const id = message.getId();
    if (seen[id]) return;
    seen[id] = true;
    unique.push(message);
  });
  return unique;
}

function selectBestDailyRevenueMessages(messages) {
  const dailyByDate = {};
  const others = [];

  messages.forEach(message => {
    const subject = message.getSubject() || '';
    const body = message.getPlainBody() || '';
    const text = subject + '\n' + body;
    if (!/每日營收|今日購買明細/.test(text)) {
      others.push(message);
      return;
    }
    const date = normalizeDate(firstMatch(text, [
      /(\d{4}\/\d{1,2}\/\d{1,2})/,
      /(\d{4}-\d{1,2}-\d{1,2})/,
    ])) || formatDate(message.getDate());
    const score = (/啟用日/.test(text) ? 10 : 0) + (text.match(/\$[\d,]+/g) || []).length;
    if (!dailyByDate[date] || score > dailyByDate[date].score) {
      dailyByDate[date] = { message, score };
    }
  });

  return [...others, ...Object.values(dailyByDate).map(item => item.message)];
}

function parsePurchaseMail(message) {
  const subject = message.getSubject() || '';
  const body = message.getPlainBody() || '';
  const text = normalizeMailText(subject + '\n' + body);

  if (/購買成功通知|bookfastpos|servicealarm\.mirle|商品明細|購買人資料/.test(text)) {
    const parsedBookFast = parseBookFastPurchaseMail(message, subject, body);
    if (parsedBookFast) return parsedBookFast;
  }

  if (/每日營收|今日購買明細/.test(text)) {
    return parseDailyRevenueMail(message, subject, body);
  }

  const student = firstMatch(text, [
    /(?:學員姓名|學生姓名|孩子姓名|姓名|購買人|報名人|家長姓名)\s*[:：]\s*([^\n\r]+)/i,
    /(?:學員|學生|孩子)\s*[:：]\s*([^\n\r]+)/i,
  ]);
  const plan = firstMatch(text, [
    /(?:方案名稱|購買方案|課程方案|方案|商品名稱|品項|項目|報名期數)\s*[:：]\s*([^\n\r]+)/i,
    /(?:購買|訂購)\s*[:：]\s*([^\n\r]+)/i,
  ]);
  const amountText = firstMatch(text, [
    /(?:付款金額|金額|總金額|總費用|實收金額|訂單金額|應付金額)\s*[:：]\s*([$＄]?\s*[\d,]+)/i,
    /(?:NT\$|TWD|\$|＄)\s*([\d,]+)/i,
  ]);

  if (!student || !plan) return null;

  const mailDate = message.getDate();
  const date = normalizeDate(firstMatch(text, [
    /(?:購買日期|付款日期|訂單日期|報名日期|日期)\s*[:：]\s*([0-9\/\-.年月日]+)/i,
  ])) || formatDate(mailDate);
  const startDate = normalizeDate(firstMatch(text, [
    /(?:啟用日期|開始日期|起始日期|開卡日期)\s*[:：]\s*([0-9\/\-.年月日]+)/i,
  ])) || date;
  const endDate = normalizeDate(firstMatch(text, [
    /(?:到期日期|結束日期|有效期限|截止日期)\s*[:：]\s*([0-9\/\-.年月日]+)/i,
  ])) || '';
  const sessions = Number(firstMatch(text, [
    /(?:堂數|購買堂數|剩餘堂數)\s*[:：]\s*(\d+)/i,
    /(\d+)\s*堂/i,
  ]) || 0);
  const serviceMonth = normalizeMonth(firstMatch(text, [
    /(?:服務月份|歸屬月份|課程月份)\s*[:：]\s*([0-9\/\-.年月]+)/i,
  ])) || (date ? date.slice(0, 7) : '');

  return {
    date,
    student: cleanValue(student),
    plan: cleanValue(plan),
    amount: parseAmount(amountText),
    dept: guessDept(plan + ' ' + subject),
    payMethod: guessPayMethod(text),
    serviceMonth,
    startDate,
    endDate,
    sessions,
    note: 'Gmail 匯入：' + subject.slice(0, 60),
    sourceId: message.getId(),
  };
}

function parseBookFastPurchaseMail(message, subject, body) {
  const text = normalizeMailText(body);
  const student = fieldValue(text, ['會員姓名', '購買人', '姓名', '會員名稱']);
  const productName = fieldValue(text, ['商品名稱', '方案名稱', '票券名稱']);
  const orderTime = fieldValue(text, ['訂購時間', '購買時間', '付款時間']);
  const periodText = fieldValue(text, ['使用期限', '有效期限']);
  const payContent = fieldValue(text, ['付款內容', '付款金額', '總金額', '金額']);
  const payMethodText = fieldValue(text, ['付款方式']);
  const ticketType = fieldValue(text, ['票券種類', '商品類型']);
  const productCode = fieldValue(text, ['商品編號', '訂單編號']);

  if (!student || !productName) return null;

  const period = parseDateRange(periodText);
  const date = normalizeDate(orderTime) || formatDate(message.getDate());
  const startDate = period.startDate || date;
  const endDate = period.endDate || '';
  const amount = parseAmount(payContent);

  return {
    date,
    student: cleanValue(student),
    plan: cleanValue(productName),
    amount,
    dept: guessDept(productName + ' ' + ticketType + ' ' + subject),
    payMethod: guessPayMethod(payMethodText || text),
    serviceMonth: startDate ? startDate.slice(0, 7) : date.slice(0, 7),
    startDate,
    endDate,
    sessions: sessionsFromPlan(productName + ' ' + ticketType),
    note: 'BookFast 匯入' + (productCode ? '：' + productCode : ''),
    sourceId: message.getId(),
  };
}

function parseDailyRevenueMail(message, subject, body) {
  const text = normalizeMailText(body);
  if (/今日尚無購買紀錄/.test(text)) return [];

  const date = normalizeDate(firstMatch(subject + '\n' + text, [
    /(\d{4}\/\d{1,2}\/\d{1,2})/,
    /(\d{4}-\d{1,2}-\d{1,2})/,
  ])) || formatDate(message.getDate());

  const lines = text.split('\n').map(line => line.trim()).filter(Boolean);
  const rows = [];
  let inTable = false;

  lines.forEach(line => {
    if (/今日購買明細/.test(line)) {
      inTable = true;
      return;
    }
    if (!inTable) return;
    if (/^(會員\s+方案\s+金額|[*＊]|系統自動發送)/.test(line)) return;
    if (!/\$[\d,]+/.test(line)) return;

    const match = line.match(/^(\S+)\s+(.+?)\s+\$([\d,]+)(?:\s+(\d{4}-\d{1,2}-\d{1,2})(?:\s+(\d{4}-\d{1,2}-\d{1,2}|-))?)?$/);
    if (!match) return;

    const student = cleanValue(match[1]);
    const plan = cleanValue(match[2]);
    const amount = parseAmount(match[3]);
    const startDate = normalizeDate(match[4]) || date;
    const endDate = match[5] === '-' ? '' : normalizeDate(match[5]);

    rows.push({
      date,
      student,
      plan,
      amount,
      dept: guessDept(plan + ' ' + subject),
      payMethod: 'transfer',
      serviceMonth: startDate ? startDate.slice(0, 7) : date.slice(0, 7),
      startDate,
      endDate,
      sessions: sessionsFromPlan(plan),
      note: 'Gmail 日報匯入：' + subject.slice(0, 60),
      sourceId: message.getId() + ':' + rows.length,
    });
  });

  return rows;
}

function normalizeMailText(text) {
  return String(text || '')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n');
}

function firstMatch(text, patterns) {
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match && match[1]) return cleanValue(match[1]);
  }
  return '';
}

function cleanValue(value) {
  return String(value || '')
    .replace(/<[^>]+>/g, '')
    .replace(/[　\t]/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim()
    .replace(/^[-：:]+/, '')
    .trim();
}

function fieldValue(text, labels) {
  const lines = normalizeMailText(text).split('\n').map(line => line.trim()).filter(Boolean);
  for (let i = 0; i < lines.length; i++) {
    for (const label of labels) {
      const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const sameLine = lines[i].match(new RegExp('^' + escaped + '(?:\\s+|\\s*[:：]\\s*)(.+)$'));
      if (sameLine && sameLine[1]) return cleanValue(sameLine[1]);
      if (lines[i] === label && lines[i + 1]) return cleanValue(lines[i + 1]);
    }
  }
  return '';
}

function parseAmount(value) {
  const text = String(value || '').replace(/,/g, '');
  const money = text.match(/(?:NT\$|TWD|\$|＄)?\s*(\d{1,7})(?:\s*元)?/i);
  const n = money ? money[1] : '';
  return n ? Number(n) : 0;
}

function normalizeDate(value) {
  if (!value) return '';
  const raw = String(value).replace(/[年月]/g, '/').replace(/[日.]/g, '').replace(/-/g, '/').trim();
  const match = raw.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (!match) return '';
  return [match[1], pad2(match[2]), pad2(match[3])].join('-');
}

function normalizeMonth(value) {
  if (!value) return '';
  const raw = String(value).replace(/[年月]/g, '/').replace(/[日.]/g, '').replace(/-/g, '/').trim();
  const match = raw.match(/(\d{4})\/(\d{1,2})/);
  if (!match) return '';
  return [match[1], pad2(match[2])].join('-');
}

function parseDateRange(value) {
  if (!value) return { startDate: '', endDate: '' };
  const matches = String(value).match(/\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2}/g) || [];
  return {
    startDate: normalizeDate(matches[0]),
    endDate: normalizeDate(matches[1]),
  };
}

function formatDate(date) {
  return [
    date.getFullYear(),
    pad2(date.getMonth() + 1),
    pad2(date.getDate()),
  ].join('-');
}

function pad2(n) {
  return String(n).padStart(2, '0');
}

function guessDept(text) {
  if (/兒童|幼兒|小孩|安親|課後/i.test(text)) return 'kids';
  if (/共用|全教室/i.test(text)) return 'shared';
  return 'sport';
}

function guessPayMethod(text) {
  if (/刷卡|信用卡|card/i.test(text)) return 'card';
  if (/現金/i.test(text)) return 'cash';
  if (/運動幣|動滋|voucher/i.test(text)) return 'voucher';
  if (/匯款|轉帳|銀行/i.test(text)) return 'transfer';
  return 'transfer';
}

function sessionsFromPlan(plan) {
  const match = String(plan || '').match(/(\d+)\s*堂/);
  return match ? Number(match[1]) : 0;
}

function getDataSheet() {
  const ss = getSpreadsheet({});
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['app', 'data_json', 'meta_json', 'updated_at']);
  }
  return sheet;
}

function getSpreadsheet(params) {
  const requestedId = params && (params.spreadsheetId || params.sheetId);
  const configuredId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const id = requestedId || configuredId || DEFAULT_SPREADSHEET_ID;
  if (id) return SpreadsheetApp.openById(id);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function assertAllowed(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('SYNC_TOKEN');
  if (expected && token !== expected) {
    throw new Error('同步密碼不正確');
  }
}

function safeJson(text, fallback) {
  try {
    return JSON.parse(text || '');
  } catch (err) {
    return fallback;
  }
}

function jsonOutput(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
