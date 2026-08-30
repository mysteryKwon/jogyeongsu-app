/**
 * 아파트 조경수 실태조사 - 스마트폰 입력앱 백엔드
 * -------------------------------------------------
 * 이 스크립트는 구글 스프레드시트에 연결된 Apps Script 프로젝트에 붙여넣고
 * "웹 앱으로 배포"하여 사용합니다. (설치가이드.md 참고)
 *
 * 처리 흐름
 *  - doPost(type: 'tree'|'bed'): 새 조사 데이터를 해당 시트에 한 행 추가
 *  - doPost(type: 'tree_update'|'bed_update'): No로 기존 행을 찾아 값 갱신
 *  - doPost(type: 'tree_delete'|'bed_delete'): No로 기존 행을 찾아 삭제
 *  - doGet(?action=list&type=tree|bed&limit=20): 최근 조사 내역 조회 (수정/삭제 화면용)
 *  - 사진이 첨부된 경우 구글 드라이브 폴더에 저장 후, 공유 링크를 시트에 기록
 *  - 사진관리 시트에도 자동으로 사진 정보를 추가
 *  - 동/구역 목록은 "동구역목록" 시트에 저장되어 모든 사용자(기기)가 공유합니다
 *  - doGet(?action=zones)로 동/구역 목록 조회, 버전 정보(SCRIPT_VERSION)도 함께 반환
 *
 * Code.gs를 수정한 뒤에는 반드시
 * [배포 → 배포 관리 → 편집(연필) → 새 버전으로 배포] 를 다시 실행해야 반영됩니다.
 */

// ===== 설정 =====
// 백엔드(Apps Script) 버전 - 프론트엔드 index.html의 APP_VERSION과 비교해 설정 탭에 표시됩니다.
const SCRIPT_VERSION = '3.1.0';

// 사진을 저장할 구글 드라이브 폴더 이름 (없으면 자동 생성됨)
const PHOTO_FOLDER_NAME = '조경수조사_사진';

// 시트 이름 (엑셀 원본과 동일하게 유지)
const SHEET_TREE = '수목실태조사';
const SHEET_BED = '화단구역조사';
const SHEET_PHOTO = '사진관리';
const SHEET_ZONES = '동구역목록'; // 없으면 자동 생성됩니다

// 각 시트의 열 순서 (헤더와 동일한 순서로 값이 채워집니다)
const COLS_TREE = ['No', '점검일', '동/구역', '상세 위치', '수종', '수량', '생육상태',
  '잎·가지 상태', '토양수분', '토양상태', '주변환경', '고사위험', '관찰내용',
  '추정원인(선택)', '우선조치', '사진번호', '재점검일', '비고'];

const COLS_BED = ['No', '점검일', '동/구역', '화단/구역 위치', '토양수분', '토양다짐',
  '표토상태', '낙엽상태', '멀칭', '배수상태', '잔디·지피 생육', '관수시설',
  '수목 전반상태', '주요 문제', '우선 개선사항', '사진번호', '비고'];

const COLS_PHOTO = ['사진번호', '촬영일', '동/구역', '위치', '대상 수목/화단',
  '촬영구분', '파일명/구글드라이브 링크', '설명', '재촬영 예정일'];

// ===== 진입점 =====
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    if (action === 'zones') {
      return jsonOut({ ok: true, zones: getZoneList(ss), version: SCRIPT_VERSION });
    }
    if (action === 'list') {
      const type = e.parameter.type; // 'tree' | 'bed'
      const limit = Number(e.parameter.limit) || 20;
      return jsonOut(listSurveyRows(ss, type, limit));
    }
    if (action === 'stats') {
      const existingStatSheet = ss.getSheetByName('자동통계');
      const needsFullRecompute = !existingStatSheet
        || !existingStatSheet.getRange(ZONE_TABLE_HEADER_ROW, 1).getValue(); // 구버전 시트(구역별 집계 표 없음) 감지
      if (needsFullRecompute) {
        // 통계 시트가 아직 없거나, 예전 버전이라 "구역별 집계" 표가 없을 때만 전체 데이터를 계산합니다.
        const stats = computeStats(ss);
        writeStatsToSheet(ss, stats);
        return jsonOut({ ok: true, stats: stats, version: SCRIPT_VERSION });
      }
      // 이후에는 저장·수정·삭제 시점에 이미 갱신해둔 값을 그대로 읽기만 해서 빠르게 응답합니다.
      const stats = readStatsFromSheet(existingStatSheet);
      return jsonOut({ ok: true, stats: stats, version: SCRIPT_VERSION });
    }
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }

  return jsonOut({ ok: true, message: '조경수 조사 API 정상 작동 중', version: SCRIPT_VERSION });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const type = body.type; // 'tree'|'bed'|'tree_update'|'bed_update'|'tree_delete'|'bed_delete'|'zone_add'|'zone_remove'
    const data = body.data || {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ---- 동/구역 목록 관리 ----
    if (type === 'zone_add') {
      const zones = addZoneToSheet(ss, (data.name || '').trim());
      return jsonOut({ ok: true, zones: zones, version: SCRIPT_VERSION });
    }
    if (type === 'zone_remove') {
      const zones = removeZoneFromSheet(ss, (data.name || '').trim());
      return jsonOut({ ok: true, zones: zones, version: SCRIPT_VERSION });
    }

    // ---- 조사 데이터 생성/수정/삭제 ----
    const photos = body.photos || []; // [{name, mime, base64}]

    if (type === 'tree' || type === 'bed') {
      return jsonOut(createSurveyRow(ss, type, data, photos));
    }
    if (type === 'tree_update' || type === 'bed_update') {
      return jsonOut(updateSurveyRow(ss, type, data, photos));
    }
    if (type === 'tree_delete' || type === 'bed_delete') {
      return jsonOut(deleteSurveyRow(ss, type, data));
    }

    throw new Error('알 수 없는 요청 유형입니다: ' + type);
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 조사 데이터 CRUD =====
function sheetAndColsFor(type) {
  const isTree = type.indexOf('tree') === 0;
  return {
    isTree: isTree,
    sheetName: isTree ? SHEET_TREE : SHEET_BED,
    cols: isTree ? COLS_TREE : COLS_BED
  };
}

function getNextNo(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return 1;
  const values = sheet.getRange(2, 1, last - 1, 1).getValues()
    .map(r => Number(r[0])).filter(n => !isNaN(n));
  if (values.length === 0) return 1;
  return Math.max.apply(null, values) + 1;
}

function findRowByNo(sheet, no) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const values = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (Number(values[i][0]) === Number(no)) return i + 2; // 실제 시트 행 번호(헤더=1행)
  }
  return -1;
}

function createSurveyRow(ss, type, data, photos) {
  const { sheetName, cols } = sheetAndColsFor(type);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('시트를 찾을 수 없습니다: ' + sheetName);

  let photoNumber = '';
  if (photos.length > 0) {
    const links = photos.map(p => savePhotoToDrive(p.name, p.mime, p.base64));
    photoNumber = generatePhotoNumber(ss);
    appendPhotoRows(ss, photoNumber, data, links, type);
  }

  const nextNo = getNextNo(sheet);
  data['No'] = nextNo;
  data['사진번호'] = photoNumber || data['사진번호'] || '';

  const row = cols.map(c => data[c] !== undefined ? data[c] : '');
  sheet.appendRow(row);
  writeStatsToSheet(ss, computeStats(ss));

  return { ok: true, no: nextNo, photoNumber: photoNumber, version: SCRIPT_VERSION };
}

function updateSurveyRow(ss, type, data, photos) {
  const { isTree, sheetName, cols } = sheetAndColsFor(type);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('시트를 찾을 수 없습니다: ' + sheetName);

  const no = data['No'];
  if (no === undefined || no === '') throw new Error('수정할 항목의 No가 없습니다.');
  const rowIndex = findRowByNo(sheet, no);
  if (rowIndex === -1) throw new Error('해당 No(' + no + ') 항목을 찾을 수 없습니다. 삭제되었거나 새로고침이 필요할 수 있습니다.');

  const photoColIdx = cols.indexOf('사진번호');
  const existingRow = sheet.getRange(rowIndex, 1, 1, cols.length).getValues()[0];
  const existingPhotoNo = photoColIdx > -1 ? (existingRow[photoColIdx] || '') : '';

  let newPhotoNumber = '';
  if (photos.length > 0) {
    const links = photos.map(p => savePhotoToDrive(p.name, p.mime, p.base64));
    newPhotoNumber = generatePhotoNumber(ss);
    appendPhotoRows(ss, newPhotoNumber, data, links, isTree ? 'tree' : 'bed');
  }

  data['No'] = no;
  data['사진번호'] = newPhotoNumber
    ? (existingPhotoNo ? existingPhotoNo + ', ' + newPhotoNumber : newPhotoNumber)
    : existingPhotoNo;

  const row = cols.map(c => data[c] !== undefined ? data[c] : '');
  sheet.getRange(rowIndex, 1, 1, cols.length).setValues([row]);
  writeStatsToSheet(ss, computeStats(ss));

  return { ok: true, no: no, photoNumber: data['사진번호'], version: SCRIPT_VERSION };
}

function deleteSurveyRow(ss, type, data) {
  const { sheetName } = sheetAndColsFor(type);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('시트를 찾을 수 없습니다: ' + sheetName);

  const no = data['No'];
  if (no === undefined || no === '') throw new Error('삭제할 항목의 No가 없습니다.');
  const rowIndex = findRowByNo(sheet, no);
  if (rowIndex === -1) throw new Error('해당 No(' + no + ') 항목을 찾을 수 없습니다. 이미 삭제되었을 수 있습니다.');

  sheet.deleteRow(rowIndex);
  writeStatsToSheet(ss, computeStats(ss));

  return { ok: true, no: no, version: SCRIPT_VERSION };
}

function listSurveyRows(ss, type, limit) {
  const { sheetName, cols } = sheetAndColsFor(type === 'tree' ? 'tree' : 'bed');
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('시트를 찾을 수 없습니다: ' + sheetName);

  const last = sheet.getLastRow();
  if (last < 2) return { ok: true, items: [], version: SCRIPT_VERSION };

  // 시트 전체를 읽지 않고, 최근 항목이 있는 "뒤쪽 구간"만 읽어서 로딩 속도를 개선합니다.
  // (새 조사는 항상 맨 아래에 추가되므로, 최신 No는 대부분 뒤쪽에 있습니다)
  const windowSize = Math.max(limit * 4, 60);
  const startRow = Math.max(2, last - windowSize + 1);
  const rowCount = last - startRow + 1;

  const values = sheet.getRange(startRow, 1, rowCount, cols.length).getValues();
  const items = values
    .map(row => {
      const obj = {};
      cols.forEach((c, idx) => { obj[c] = row[idx]; });
      return obj;
    })
    .filter(o => o['No'] !== '' && o['No'] !== undefined && o['No'] !== null)
    .sort((a, b) => Number(b['No']) - Number(a['No']))
    .slice(0, limit);

  return { ok: true, items: items, version: SCRIPT_VERSION };
}

// ===== 동/구역 목록 (구글시트 저장) =====
function ensureZoneSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_ZONES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ZONES);
    sheet.appendRow(['동/구역', '등록일']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getZoneList(ss) {
  const sheet = ensureZoneSheet(ss);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, 1).getValues().map(r => r[0]).filter(v => v !== '');
  return values.sort((a, b) => String(a).localeCompare(String(b), 'ko'));
}

function addZoneToSheet(ss, name) {
  if (!name) throw new Error('구역명이 비어있습니다.');
  const sheet = ensureZoneSheet(ss);
  const existing = getZoneList(ss);
  if (existing.indexOf(name) === -1) {
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    sheet.appendRow([name, today]);
  }
  return getZoneList(ss);
}

function removeZoneFromSheet(ss, name) {
  if (!name) throw new Error('구역명이 비어있습니다.');
  const sheet = ensureZoneSheet(ss);
  const last = sheet.getLastRow();
  if (last >= 2) {
    const values = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === name) {
        sheet.deleteRow(i + 2);
        break;
      }
    }
  }
  return getZoneList(ss);
}

// ===== 사진 저장 =====
function getPhotoFolder() {
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function savePhotoToDrive(name, mime, base64) {
  const folder = getPhotoFolder();
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mime || 'image/jpeg', name || ('photo_' + Date.now() + '.jpg'));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function generatePhotoNumber(ss) {
  const sheet = ss.getSheetByName(SHEET_PHOTO);
  const count = sheet.getLastRow(); // 헤더 제외 현재까지 사진 행 수 = lastRow - 1, 새 번호는 그 다음
  const num = count; // lastRow가 헤더포함이므로 count가 곧 다음 순번
  return 'P-' + ('0000' + num).slice(-4);
}

function appendPhotoRows(ss, photoNumber, data, links, type) {
  const sheet = ss.getSheetByName(SHEET_PHOTO);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const target = type === 'tree' ? (data['수종'] || '') : (data['화단/구역 위치'] || '');
  const location = type === 'tree' ? (data['상세 위치'] || '') : (data['화단/구역 위치'] || '');

  links.forEach((url, idx) => {
    const num = links.length > 1 ? photoNumber + '-' + (idx + 1) : photoNumber;
    const row = [
      num,
      today,
      data['동/구역'] || '',
      location,
      target,
      type === 'tree' ? '수목조사' : '화단조사',
      url,
      data['관찰내용'] || data['주요 문제'] || '',
      ''
    ];
    sheet.appendRow(row);
  });
}

// ===== 자동통계 (생육상태 · 핵심지표) =====
// 원본 엑셀의 "자동통계" 시트 레이아웃과 동일하게 유지합니다:
//   A3:B3 헤더(생육상태/건수), A4:A8 데이터
//   D3:E3 헤더(핵심지표/값),   D4:D9 데이터
const STAT_GROWTH_ORDER = ['정상', '주의', '생육불량', '고사위험', '고사'];
const STAT_INDICATOR_LABELS = ['조사 건수', '고사위험+고사', '긴급관리 대상',
  '토양 매우건조', '토양다짐 의심', '화단 조사구역'];
const BED_STATE_ORDER = ['양호', '일부 불량', '다수 불량', '고사목 있음'];

function computeStats(ss) {
  const tree = ss.getSheetByName(SHEET_TREE);
  const bed = ss.getSheetByName(SHEET_BED);

  const growth = {}; STAT_GROWTH_ORDER.forEach(k => growth[k] = 0);
  let treeCount = 0, urgent = 0, veryDry = 0, compact = 0;

  // 구역별 집계용 맵. 등록된 동/구역 목록을 먼저 채워서, 조사 실적이 0건인 구역도 함께 보여줍니다.
  const registeredZones = getZoneList(ss);
  const zoneMap = {};
  function zoneEntry(z) {
    if (!zoneMap[z]) {
      zoneMap[z] = {
        zone: z, treeTotal: 0, bedTotal: 0,
        treeGrowth: { 정상: 0, 주의: 0, 생육불량: 0, 고사위험: 0, 고사: 0 },
        bedState: { '양호': 0, '일부 불량': 0, '다수 불량': 0, '고사목 있음': 0 }
      };
    }
    return zoneMap[z];
  }
  registeredZones.forEach(z => zoneEntry(z));

  if (tree && tree.getLastRow() > 1) {
    const data = tree.getRange(2, 1, tree.getLastRow() - 1, COLS_TREE.length).getValues();
    const idx = name => COLS_TREE.indexOf(name);
    const stateIdx = idx('생육상태'), riskIdx = idx('고사위험'),
          moistIdx = idx('토양수분'), soilIdx = idx('토양상태'), noIdx = idx('No'),
          zoneIdx = idx('동/구역');

    data.forEach(r => {
      if (r[noIdx] === '' || r[noIdx] === null) return; // 빈 행 제외
      treeCount++;
      const st = r[stateIdx];
      if (growth[st] !== undefined) growth[st]++;
      if (r[riskIdx] === '긴급') urgent++;              // 고사위험 컬럼: 낮음/중간/높음/긴급/고사
      if (r[moistIdx] === '매우 건조') veryDry++;         // 토양수분 컬럼 (공백 포함, 시트 목록과 동일)
      if (r[soilIdx] === '단단함/다짐') compact++;        // 토양상태 컬럼

      const z = r[zoneIdx] || '(미지정)';
      const ze = zoneEntry(z);
      ze.treeTotal++;
      if (ze.treeGrowth[st] !== undefined) ze.treeGrowth[st]++;
    });
  }

  let bedCount = 0;
  if (bed && bed.getLastRow() > 1) {
    const idx = name => COLS_BED.indexOf(name);
    const noIdx = idx('No'), zoneIdx = idx('동/구역'), stateIdx = idx('수목 전반상태');
    const data = bed.getRange(2, 1, bed.getLastRow() - 1, COLS_BED.length).getValues();
    data.forEach(r => {
      if (r[noIdx] === '' || r[noIdx] === null) return;
      bedCount++;
      const z = r[zoneIdx] || '(미지정)';
      const ze = zoneEntry(z);
      ze.bedTotal++;
      const st = r[stateIdx];
      if (ze.bedState[st] !== undefined) ze.bedState[st]++;
    });
  }

  const byZone = Object.keys(zoneMap)
    .map(z => zoneMap[z])
    .filter(ze => ze.treeTotal > 0 || ze.bedTotal > 0 || registeredZones.indexOf(ze.zone) !== -1)
    .sort((a, b) => String(a.zone).localeCompare(String(b.zone), 'ko'));

  return {
    updatedAt: new Date().toISOString(),
    growth: STAT_GROWTH_ORDER.map(k => ({ label: k, count: growth[k] })),
    indicators: [
      { label: STAT_INDICATOR_LABELS[0], value: treeCount },
      { label: STAT_INDICATOR_LABELS[1], value: growth['고사위험'] + growth['고사'] },
      { label: STAT_INDICATOR_LABELS[2], value: urgent },
      { label: STAT_INDICATOR_LABELS[3], value: veryDry },
      { label: STAT_INDICATOR_LABELS[4], value: compact },
      { label: STAT_INDICATOR_LABELS[5], value: bedCount }
    ],
    byZone: byZone
  };
}

const ZONE_TABLE_HEADER_ROW = 12;
const ZONE_TABLE_DATA_ROW = 13;
const ZONE_TABLE_MAX_ROWS = 300;
const ZONE_TABLE_COLS = ['동/구역', '수목 건수', '정상', '주의', '생육불량', '고사위험', '고사',
  '화단 건수', '양호', '일부 불량', '다수 불량', '고사목 있음'];

function writeStatsToSheet(ss, stats) {
  try {
    const stat = ensureStatsSheet(ss);
    stat.getRange(4, 1, STAT_GROWTH_ORDER.length, 1)
      .setValues(STAT_GROWTH_ORDER.map(k => [k])); // 라벨도 항상 최신 상태로 맞춰줌
    stat.getRange(4, 2, STAT_GROWTH_ORDER.length, 1)
      .setValues(stats.growth.map(g => [g.count]));
    stat.getRange(4, 4, stats.indicators.length, 1)
      .setValues(STAT_INDICATOR_LABELS.map(k => [k])); // 라벨도 항상 최신 상태로 맞춰줌
    stat.getRange(4, 5, stats.indicators.length, 1)
      .setValues(stats.indicators.map(i => [i.value]));
    stat.getRange('G1').setValue('마지막 갱신: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
    stat.getRange('I1').setValue(stats.updatedAt); // 프로그램에서 다시 읽을 때 쓰는 원본 타임스탬프(숨김용)

    // 구역별 집계 표 (동/화단 조사를 동/구역별로 구분)
    stat.getRange(ZONE_TABLE_HEADER_ROW - 1, 1).setValue('구역별 집계').setFontWeight('bold').setFontSize(13);
    stat.getRange(ZONE_TABLE_HEADER_ROW, 1, 1, ZONE_TABLE_COLS.length)
      .setValues([ZONE_TABLE_COLS]).setFontWeight('bold');
    // 구역 수가 줄어들 때를 대비해 매번 넓게 지운 뒤 다시 씀
    stat.getRange(ZONE_TABLE_DATA_ROW, 1, ZONE_TABLE_MAX_ROWS, ZONE_TABLE_COLS.length).clearContent();
    const byZone = stats.byZone || [];
    if (byZone.length) {
      const rows = byZone.map(z => [
        z.zone, z.treeTotal,
        z.treeGrowth['정상'], z.treeGrowth['주의'], z.treeGrowth['생육불량'], z.treeGrowth['고사위험'], z.treeGrowth['고사'],
        z.bedTotal,
        z.bedState['양호'], z.bedState['일부 불량'], z.bedState['다수 불량'], z.bedState['고사목 있음']
      ]);
      stat.getRange(ZONE_TABLE_DATA_ROW, 1, rows.length, ZONE_TABLE_COLS.length).setValues(rows);
    }
  } catch (err) {
    // 통계 시트 갱신 실패는 조사 저장 자체를 막지 않도록 조용히 무시
  }
}

// 이미 시트에 계산되어 있는 통계 값을 "다시 계산하지 않고" 그대로 읽기만 함 (통계 탭 로딩 속도 개선용)
function readStatsFromSheet(stat) {
  const growthCounts = stat.getRange(4, 2, STAT_GROWTH_ORDER.length, 1).getValues();
  const growth = STAT_GROWTH_ORDER.map((label, i) => ({ label: label, count: Number(growthCounts[i][0]) || 0 }));

  const indicatorValues = stat.getRange(4, 5, STAT_INDICATOR_LABELS.length, 1).getValues();
  const indicators = STAT_INDICATOR_LABELS.map((label, i) => ({ label: label, value: Number(indicatorValues[i][0]) || 0 }));

  let updatedAt = stat.getRange('I1').getValue();
  updatedAt = updatedAt ? String(updatedAt) : new Date().toISOString();

  const zoneValues = stat.getRange(ZONE_TABLE_DATA_ROW, 1, ZONE_TABLE_MAX_ROWS, ZONE_TABLE_COLS.length).getValues();
  const byZone = zoneValues
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      zone: r[0],
      treeTotal: Number(r[1]) || 0,
      treeGrowth: { 정상: Number(r[2]) || 0, 주의: Number(r[3]) || 0, 생육불량: Number(r[4]) || 0, 고사위험: Number(r[5]) || 0, 고사: Number(r[6]) || 0 },
      bedTotal: Number(r[7]) || 0,
      bedState: { '양호': Number(r[8]) || 0, '일부 불량': Number(r[9]) || 0, '다수 불량': Number(r[10]) || 0, '고사목 있음': Number(r[11]) || 0 }
    }));

  return { updatedAt: updatedAt, growth: growth, indicators: indicators, byZone: byZone };
}

function ensureStatsSheet(ss) {
  let stat = ss.getSheetByName('자동통계');
  if (stat) return stat;

  stat = ss.insertSheet('자동통계');
  stat.getRange('A1').setValue('아파트 조경수 1차 실태조사 자동통계');
  stat.getRange('A3:B3').setValues([['생육상태', '건수']]);
  stat.getRange('D3:E3').setValues([['핵심지표', '값']]);
  stat.getRange('A4:A8').setValues(STAT_GROWTH_ORDER.map(k => [k]));
  stat.getRange('D4:D9').setValues(STAT_INDICATOR_LABELS.map(k => [k]));
  stat.getRange('A1').setFontWeight('bold').setFontSize(13);
  stat.getRange('A3:B3').setFontWeight('bold');
  stat.getRange('D3:E3').setFontWeight('bold');
  stat.autoResizeColumns(1, 6);
  return stat;
}

