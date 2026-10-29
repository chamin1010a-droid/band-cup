/**
 * 밴드 취향 월드컵 서버 (구글 시트 + Apps Script)
 *
 * members 시트: 한 줄에 한 사람. id | 닉네임 | 단계 | 판수 | 갱신 시각 | 기록(JSON)
 * config 시트: A1 "open", B1 TRUE/FALSE (비교 열림 여부)
 *
 * 처음 한 번 setup()을 실행하면 시트가 만들어지고 관리자 키가 로그에 찍힘.
 */

const MAX_JSON = 45000; // 시트 한 칸 한도(5만 자)보다 작게

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let m = ss.getSheetByName('members');
  if (!m) m = ss.insertSheet('members');
  if (m.getLastRow() === 0) m.appendRow(['id', '닉네임', '단계', '판수', '갱신 시각', '기록']);
  let c = ss.getSheetByName('config');
  if (!c) c = ss.insertSheet('config');
  c.getRange('A1:B1').setValues([['open', false]]);

  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('ADMIN_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    props.setProperty('ADMIN_KEY', key);
  }
  Logger.log('관리자 키: ' + key);
}

function doGet() {
  return json_(readState_());
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: 'busy' });
  try {
    if (req.action === 'save') return json_(saveMember_(req.id, req.data));
    if (req.action === 'open') {
      if (!req.key || req.key !== PropertiesService.getScriptProperties().getProperty('ADMIN_KEY')) {
        return json_({ ok: false, error: 'not_admin' });
      }
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config').getRange('B1').setValue(true);
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } finally {
    lock.releaseLock();
  }
}

function readState_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = ss.getSheetByName('members').getDataRange().getValues().slice(1);
  const members = {};
  rows.forEach(r => {
    if (!r[0] || !r[5]) return;
    try { members[r[0]] = JSON.parse(r[5]); } catch (err) {}
  });
  const open = ss.getSheetByName('config').getRange('B1').getValue() === true;
  return { ok: true, members: members, config: { open: open } };
}

function saveMember_(id, data) {
  if (typeof id !== 'string' || !/^[0-9a-f]{16,40}$/.test(id)) return { ok: false, error: 'bad_id' };
  if (!data || typeof data !== 'object' || data.v !== 1 || !Array.isArray(data.matches)) return { ok: false, error: 'bad_data' };
  if (data.matches.length > 300 || (data.added || []).length > 2) return { ok: false, error: 'too_big' };
  data.name = String(data.name || '').slice(0, 20);

  const text = JSON.stringify(data);
  if (text.length > MAX_JSON) return { ok: false, error: 'too_big' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('members');
  const row = [id, data.name, data.phase || '', data.matches.length, new Date(), text];
  const ids = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues().map(r => r[0]);
  const at = ids.indexOf(id);
  if (at > 0) sheet.getRange(at + 1, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
  return { ok: true };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
