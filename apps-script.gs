const SHEET_NAME = 'studio_data';
const DEFAULT_APP_ID = 'studio-dashboard';

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
  const query = params.query || 'newer_than:90d {武樂 方案 購買 報名 付款 訂單 月卡 堂數}';
  const max = Math.min(Number(params.max || 80), 150);
  const threads = GmailApp.search(query, 0, max);
  const purchases = [];

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const parsed = parsePurchaseMail(message);
      if (parsed) purchases.push(parsed);
    });
  });

  return { ok: true, query, count: purchases.length, purchases };
}

function parsePurchaseMail(message) {
  const subject = message.getSubject() || '';
  const body = message.getPlainBody() || '';
  const text = normalizeMailText(subject + '\n' + body);

  const student = firstMatch(text, [
    /(?:學員姓名|學生姓名|孩子姓名|姓名|購買人|報名人|家長姓名)\s*[:：]\s*([^\n\r]+)/i,
    /(?:學員|學生|孩子)\s*[:：]\s*([^\n\r]+)/i,
  ]);
  const plan = firstMatch(text, [
    /(?:方案名稱|購買方案|課程方案|方案|商品名稱|品項|項目)\s*[:：]\s*([^\n\r]+)/i,
    /(?:購買|訂購)\s*[:：]\s*([^\n\r]+)/i,
  ]);
  const amountText = firstMatch(text, [
    /(?:付款金額|金額|總金額|實收金額|訂單金額|應付金額)\s*[:：]\s*([$＄]?\s*[\d,]+)/i,
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

function parseAmount(value) {
  const n = String(value || '').replace(/[^\d]/g, '');
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

function getDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['app', 'data_json', 'meta_json', 'updated_at']);
  }
  return sheet;
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
