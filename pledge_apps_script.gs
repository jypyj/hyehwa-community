/**
 * 2026 혜화공동체 약속 - 온라인 서약 데이터 수집 Apps Script
 *
 * [설정 방법]
 * 1. Google Sheets에서 새 스프레드시트를 만듭니다.
 * 2. 확장 프로그램 > Apps Script 를 클릭합니다.
 * 3. 이 코드를 전체 복사하여 붙여넣습니다.
 * 4. 아래 SPREADSHEET_ID를 스프레드시트 URL에서 복사하여 넣습니다.
 *    (URL 예시: https://docs.google.com/spreadsheets/d/여기가_ID/edit)
 * 5. 배포 > 새 배포 > 유형: 웹 앱 선택
 *    - 실행 주체: 본인
 *    - 액세스 권한: 모든 사용자
 * 6. 배포 후 나오는 URL을 pledge.html의 SCRIPT_URL에 붙여넣습니다.
 */

const SPREADSHEET_ID = ''; // ← 스프레드시트 ID를 여기에 입력

/* ===== 초기 시트 설정 (최초 1회 실행) ===== */
function setupSheet() {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  // 시트가 없으면 생성
  let sheet = ss.getSheetByName('서약현황');
  if (!sheet) {
    sheet = ss.insertSheet('서약현황');
  }

  // 헤더 설정
  const headers = [
    '타임스탬프', '구분', '학년', '반', '번호', '이름',
    '약속 이해 동의', '실천 다짐 동의', '실천 다짐 한마디', '서약완료 이미지'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 헤더 서식
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#2C5F8A');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  // 열 너비 조정
  sheet.setColumnWidth(1, 160); // 타임스탬프
  sheet.setColumnWidth(2, 80);  // 구분
  sheet.setColumnWidth(3, 60);  // 학년
  sheet.setColumnWidth(4, 60);  // 반
  sheet.setColumnWidth(5, 60);  // 번호
  sheet.setColumnWidth(6, 100); // 이름
  sheet.setColumnWidth(7, 120); // 약속 이해
  sheet.setColumnWidth(8, 120); // 실천 다짐
  sheet.setColumnWidth(9, 300); // 다짐 한마디
  sheet.setColumnWidth(10, 200); // 서약완료 이미지

  // 행 고정
  sheet.setFrozenRows(1);

  // 통계 시트 생성
  let statSheet = ss.getSheetByName('통계');
  if (!statSheet) {
    statSheet = ss.insertSheet('통계');
  }
  setupStatSheet(statSheet);

  SpreadsheetApp.getUi().alert('시트 설정이 완료되었습니다!');
}

/* ===== 통계 시트 설정 ===== */
function setupStatSheet(sheet) {
  sheet.clear();

  // 제목
  sheet.getRange('A1').setValue('2026 혜화공동체 약속 서약 현황');
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold').setFontColor('#2C5F8A');

  // 전체 참여 현황
  sheet.getRange('A3').setValue('전체 참여 현황').setFontWeight('bold');
  sheet.getRange('A4').setValue('전체 서약 인원');
  sheet.getRange('B4').setFormula('=COUNTA(서약현황!A:A)-1');

  sheet.getRange('A6').setValue('구분별 참여 현황').setFontWeight('bold');
  sheet.getRange('A7').setValue('학생');
  sheet.getRange('B7').setFormula('=COUNTIF(서약현황!B:B,"학생")');
  sheet.getRange('A8').setValue('학부모');
  sheet.getRange('B8').setFormula('=COUNTIF(서약현황!B:B,"학부모")');
  sheet.getRange('A9').setValue('교직원');
  sheet.getRange('B9').setFormula('=COUNTIF(서약현황!B:B,"교직원")');

  // 학년별 학생 참여 현황
  sheet.getRange('A11').setValue('학년별 학생 참여 현황').setFontWeight('bold');
  sheet.getRange('A12').setValue('1학년');
  sheet.getRange('B12').setFormula('=COUNTIFS(서약현황!B:B,"학생",서약현황!C:C,"1")');
  sheet.getRange('A13').setValue('2학년');
  sheet.getRange('B13').setFormula('=COUNTIFS(서약현황!B:B,"학생",서약현황!C:C,"2")');
  sheet.getRange('A14').setValue('3학년');
  sheet.getRange('B14').setFormula('=COUNTIFS(서약현황!B:B,"학생",서약현황!C:C,"3")');

  // 반별 현황 (1학년)
  sheet.getRange('D6').setValue('1학년 반별 현황').setFontWeight('bold');
  for (let i = 1; i <= 7; i++) {
    sheet.getRange('D' + (6 + i)).setValue(i + '반');
    sheet.getRange('E' + (6 + i)).setFormula(
      '=COUNTIFS(서약현황!B:B,"학생",서약현황!C:C,"1",서약현황!D:D,"' + i + '")'
    );
  }

  // 반별 현황 (2학년)
  sheet.getRange('F6').setValue('2학년 반별 현황').setFontWeight('bold');
  for (let i = 1; i <= 7; i++) {
    sheet.getRange('F' + (6 + i)).setValue(i + '반');
    sheet.getRange('G' + (6 + i)).setFormula(
      '=COUNTIFS(서약현황!B:B,"학생",서약현황!C:C,"2",서약현황!D:D,"' + i + '")'
    );
  }

  // 반별 현황 (3학년)
  sheet.getRange('H6').setValue('3학년 반별 현황').setFontWeight('bold');
  for (let i = 1; i <= 7; i++) {
    sheet.getRange('H' + (6 + i)).setValue(i + '반');
    sheet.getRange('I' + (6 + i)).setFormula(
      '=COUNTIFS(서약현황!B:B,"학생",서약현황!C:C,"3",서약현황!D:D,"' + i + '")'
    );
  }

  // 열 너비
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 80);
}

/* ===== POST 요청 처리 ===== */
function doPost(e) {
  try {
    const data = e.parameter && e.parameter.role
      ? e.parameter
      : JSON.parse(e.postData.contents);

    if (data.action === 'submitImage') {
      return handleImageSubmit(data);
    }

    const ss = SPREADSHEET_ID
      ? SpreadsheetApp.openById(SPREADSHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('서약현황');

    if (!sheet) {
      return makeResponse(false, '시트를 찾을 수 없습니다. setupSheet()을 먼저 실행하세요.');
    }

    const row = [
      data.timestamp || new Date().toLocaleString('ko-KR'),
      data.role || '',
      data.grade || '',
      data.classNum || '',
      data.number || '',
      data.name || '',
      data.pledge1 || '',
      data.pledge2 || '',
      data.resolve || ''
    ];

    sheet.appendRow(row);

    return makeResponse(true, '서약이 기록되었습니다.');
  } catch (err) {
    return makeResponse(false, err.message);
  }
}

/* ===== 서약완료 이미지 처리 ===== */
function handleImageSubmit(data) {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  // 서약완료_이미지 시트 (없으면 생성)
  let imgSheet = ss.getSheetByName('서약완료_이미지');
  if (!imgSheet) {
    imgSheet = ss.insertSheet('서약완료_이미지');
    const headers = ['제출시간', '구분', '학년', '반', '번호', '이름'];
    imgSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    imgSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#2C5F8A')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center');
    imgSheet.setFrozenRows(1);
    imgSheet.setColumnWidth(1, 160);
    imgSheet.setColumnWidth(2, 80);
    imgSheet.setColumnWidth(3, 60);
    imgSheet.setColumnWidth(4, 60);
    imgSheet.setColumnWidth(5, 60);
    imgSheet.setColumnWidth(6, 100);
  }

  const imageBase64 = data.imageData.replace(/^data:image\/\w+;base64,/, '');
  const decoded = Utilities.base64Decode(imageBase64);

  const now = new Date();
  const timestamp = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  const nameStr = data.name || '';
  const fileName = (data.role || '') + '_' + (nameStr || '익명') + '_' + Utilities.formatDate(now, 'Asia/Seoul', 'yyyyMMdd_HHmmss') + '.jpg';
  const blob = Utilities.newBlob(decoded, 'image/jpeg', fileName);

  // 정보 행 추가
  const newRow = imgSheet.getLastRow() + 1;
  imgSheet.getRange(newRow, 1).setValue(timestamp);
  imgSheet.getRange(newRow, 2).setValue(data.role || '');
  imgSheet.getRange(newRow, 3).setValue(data.grade || '');
  imgSheet.getRange(newRow, 4).setValue(data.classNum || '');
  imgSheet.getRange(newRow, 5).setValue(data.number || '');
  imgSheet.getRange(newRow, 6).setValue(nameStr);

  // 이미지 직접 삽입 (7번째 열 위치)
  imgSheet.insertImage(blob, 7, newRow);
  imgSheet.setRowHeight(newRow, 300);

  // 서약현황 시트에도 제출 기록
  const sheet = ss.getSheetByName('서약현황');
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (let i = dataRange.length - 1; i >= 0; i--) {
        if (dataRange[i][1] === data.role && String(dataRange[i][5]) === nameStr) {
          sheet.getRange(i + 2, 10).setValue('제출완료 (' + timestamp + ')');
          break;
        }
      }
    }
  }

  return makeResponse(true, '이미지가 저장되었습니다.');
}

/* ===== GET 요청 처리 (테스트용) ===== */
function doGet(e) {
  return makeResponse(true, '혜화공동체 약속 서약 API가 정상 작동 중입니다.');
}

/* ===== 응답 생성 ===== */
function makeResponse(success, message) {
  const output = JSON.stringify({ success: success, message: message });
  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}
