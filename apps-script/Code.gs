/**
 * 도장 도안 제출 수신기 (Google Apps Script 웹앱)
 * - 학생 앱에서 POST로 보낸 PNG를 내 드라이브의 "도장도안_제출" 폴더에 저장한다.
 * - 같은 폴더의 "제출기록" 스프레드시트에 한 줄씩 기록한다.
 * - 같은 학급_번호_이름 파일이 이미 있으면 기존 파일은 "이전제출" 하위 폴더로 옮기고 새 파일을 저장한다.
 *
 * 배포: 배포 > 새 배포 > 유형 "웹 앱", 실행 주체 "나", 액세스 권한 "모든 사용자"
 */
const FOLDER_NAME = '도장도안_제출';
const OLD_FOLDER_NAME = '이전제출';
const SHEET_NAME = '제출기록';

function getFolder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function getSubFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function getSheet_(folder) {
  const files = folder.getFilesByName(SHEET_NAME);
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    ss.getActiveSheet().appendRow(['제출시각', '학급', '번호', '이름', '위 문구', '아래 문구', '파일명', '링크']);
  }
  return ss.getActiveSheet();
}

function safeName_(s) {
  return String(s || '').replace(/[\\\/:*?"<>|\r\n]/g, '').trim();
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d.image || d.image.indexOf('data:image/png;base64,') !== 0) {
      return out_({ ok: false, error: 'PNG 데이터가 아닙니다.' });
    }
    const cls = safeName_(d.cls) || '학급';
    const num = safeName_(d.num).padStart(2, '0');
    const name = safeName_(d.name) || '이름없음';
    const base = cls + '_' + num + '_' + name;
    const fname = base + '.png';

    const folder = getFolder_();
    const existing = folder.getFilesByName(fname);
    if (existing.hasNext()) {
      const old = getSubFolder_(folder, OLD_FOLDER_NAME);
      const stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmmss');
      while (existing.hasNext()) {
        const f = existing.next();
        f.setName(base + '_' + stamp + '.png');
        f.moveTo(old);
      }
    }
    const bytes = Utilities.base64Decode(d.image.split(',')[1]);
    const file = folder.createFile(Utilities.newBlob(bytes, 'image/png', fname));
    getSheet_(folder).appendRow([new Date(), cls, num, name, d.top || '', d.bottom || '', fname, file.getUrl()]);
    return out_({ ok: true, file: fname });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return out_({ ok: true, msg: '도장 도안 수신기 작동 중' });
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 편집기에서 한 번 실행: 폴더·시트를 미리 만들고 권한을 승인한다. */
function setup() {
  const folder = getFolder_();
  getSubFolder_(folder, OLD_FOLDER_NAME);
  getSheet_(folder);
  Logger.log('폴더 준비 완료: ' + folder.getUrl());
}
