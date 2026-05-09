const SHEET_NAME = 'studio_data';
const DEFAULT_APP_ID = 'studio-dashboard';

function doGet(e) {
  try {
    const params = e.parameter || {};
    assertAllowed(params.token || '');
    if ((params.action || 'load') !== 'load') {
      return jsonOutput({ ok: false, error: 'Unknown action' });
    }
    return jsonOutput(loadAppData(params.app || DEFAULT_APP_ID));
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
