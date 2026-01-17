// ============================================
// 설정
// ============================================
const CONFIG = {
  CLIENT_ID: '47931668114-m851pfnh8l7so2b7qqiki2lrrjoq6iqe.apps.googleusercontent.com',
  CLIENT_SECRET: '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',
  REFRESH_TOKEN: '1//040I3uTHU27TZCgYIARAAGAQSNwF-L9IrPfyga-dYlm5urSlfxYlnGxA_sfXh9M_ulF9N9Hy7T4WjH4cJbVch-rXdW2FtrlvAUB4'
};

// ============================================
// Access Token 발급
// ============================================
function getAccessToken() {
  const response = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: {
      client_id: CONFIG.CLIENT_ID,
      client_secret: CONFIG.CLIENT_SECRET,
      refresh_token: CONFIG.REFRESH_TOKEN,
      grant_type: 'refresh_token'
    }
  });
  return JSON.parse(response.getContentText()).access_token;
}

// ============================================
// AdMob 계정 ID 가져오기
// ============================================
function getAdMobAccountId() {
  const accessToken = getAccessToken();
  const response = UrlFetchApp.fetch('https://admob.googleapis.com/v1/accounts', {
    headers: { 'Authorization': 'Bearer ' + accessToken }
  });
  const data = JSON.parse(response.getContentText());
  if (data.account && data.account.length > 0) {
    return data.account[0].name;
  }
  throw new Error('AdMob 계정을 찾을 수 없습니다.');
}

// ============================================
// 특정 기간 리포트 가져오기
// ============================================
function fetchReportByDateRange(startDate, endDate) {
  const accessToken = getAccessToken();
  const accountId = getAdMobAccountId();
  
  const requestBody = {
    reportSpec: {
      dateRange: {
        startDate: {
          year: startDate.getFullYear(),
          month: startDate.getMonth() + 1,
          day: startDate.getDate()
        },
        endDate: {
          year: endDate.getFullYear(),
          month: endDate.getMonth() + 1,
          day: endDate.getDate()
        }
      },
      dimensions: ['DATE', 'APP', 'AD_UNIT', 'FORMAT'],
      metrics: ['IMPRESSIONS', 'CLICKS', 'ESTIMATED_EARNINGS', 'IMPRESSION_RPM'],
      sortConditions: [{ dimension: 'DATE', order: 'DESCENDING' }]
    }
  };
  
  const response = UrlFetchApp.fetch(
    `https://admob.googleapis.com/v1/${accountId}/networkReport:generate`,
    {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + accessToken,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(requestBody)
    }
  );
  
  return JSON.parse(response.getContentText());
}

// ============================================
// 날짜를 문자열로 변환 (YYYY-MM-DD)
// ============================================
function formatDateToString(value) {
  if (!value) return '';
  if (typeof value === 'string') {
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
    if (/^\d{8}$/.test(value)) {
      return `${value.slice(0,4)}-${value.slice(4,6)}-${value.slice(6,8)}`;
    }
  }
  if (value instanceof Date) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, '0');
    const day = String(value.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  return String(value);
}

// ============================================
// 리포트 행 파싱
// ============================================
function parseReportRow(row) {
  try {
    const dimensions = row.dimensionValues || {};
    const metrics = row.metricValues || {};
    
    const dateStr = dimensions.DATE?.value;
    if (!dateStr) return null;
    
    const date = `${dateStr.slice(0,4)}-${dateStr.slice(4,6)}-${dateStr.slice(6,8)}`;
    const appName = dimensions.APP?.displayLabel || dimensions.APP?.value || 'Unknown';
    const adUnit = dimensions.AD_UNIT?.displayLabel || dimensions.AD_UNIT?.value || 'Unknown';
    const adFormat = dimensions.FORMAT?.value || 'Unknown';
    
    const impressions = parseInt(metrics.IMPRESSIONS?.integerValue || '0');
    const clicks = parseInt(metrics.CLICKS?.integerValue || '0');
    const earningsMicros = parseInt(metrics.ESTIMATED_EARNINGS?.microsValue || '0');
    const revenue = earningsMicros / 1000000;
    const ecpm = metrics.IMPRESSION_RPM?.doubleValue || 0;
    
    return [date, appName, adUnit, adFormat, impressions, clicks, revenue, ecpm];
  } catch (e) {
    Logger.log('행 파싱 오류: ' + e.message);
    return null;
  }
}

// ============================================
// 리포트 데이터를 시트에 저장/업데이트
// ============================================
function saveReportToSheet(reportData, updateExisting = false) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('daily');
  if (!sheet) throw new Error('daily 시트를 찾을 수 없습니다.');
  
  const rows = [];
  
  if (Array.isArray(reportData)) {
    reportData.forEach(item => {
      if (item.row) {
        const row = parseReportRow(item.row);
        if (row) rows.push(row);
      }
    });
  } else if (reportData.rows) {
    reportData.rows.forEach(row => {
      const parsed = parseReportRow(row);
      if (parsed) rows.push(parsed);
    });
  }
  
  if (rows.length === 0) {
    Logger.log('저장할 데이터가 없습니다.');
    return 0;
  }
  
  const lastRow = sheet.getLastRow();
  
  if (updateExisting && lastRow > 1) {
    // 기존 데이터 가져오기
    const existingData = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const existingMap = new Map();
    
    existingData.forEach((row, idx) => {
      const dateStr = formatDateToString(row[0]);
      const key = `${dateStr}|${row[1]}|${row[2]}|${row[3]}`;
      existingMap.set(key, idx + 2); // 행 번호 저장 (헤더 제외)
    });
    
    let updatedCount = 0;
    let addedCount = 0;
    const newRows = [];
    
    rows.forEach(row => {
      const key = `${row[0]}|${row[1]}|${row[2]}|${row[3]}`;
      
      if (existingMap.has(key)) {
        // 기존 데이터 업데이트
        const rowNum = existingMap.get(key);
        sheet.getRange(rowNum, 1, 1, 8).setValues([row]);
        updatedCount++;
      } else {
        // 새 데이터 추가
        newRows.push(row);
        addedCount++;
      }
    });
    
    if (newRows.length > 0) {
      const insertRow = lastRow + 1;
      sheet.getRange(insertRow, 1, newRows.length, 8).setValues(newRows);
    }
    
    Logger.log(`업데이트: ${updatedCount}개, 추가: ${addedCount}개`);
    return updatedCount + addedCount;
    
  } else {
    // 중복 체크 후 새 데이터만 추가
    const existingKeys = new Set();
    
    if (lastRow > 1) {
      const existingData = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
      existingData.forEach(row => {
        const dateStr = formatDateToString(row[0]);
        const key = `${dateStr}|${row[1]}|${row[2]}|${row[3]}`;
        existingKeys.add(key);
      });
    }
    
    const newRows = rows.filter(row => {
      const key = `${row[0]}|${row[1]}|${row[2]}|${row[3]}`;
      return !existingKeys.has(key);
    });
    
    if (newRows.length > 0) {
      const insertRow = lastRow > 0 ? lastRow + 1 : 2;
      sheet.getRange(insertRow, 1, newRows.length, 8).setValues(newRows);
      Logger.log(`${newRows.length}개의 새 데이터가 추가되었습니다.`);
    }
    
    return newRows.length;
  }
}

// ============================================
// ★ 오늘 데이터 가져오기 (실시간)
// ============================================
function fetchTodayReport() {
  try {
    Logger.log('오늘 데이터 가져오기 시작...');
    const today = new Date();
    const reportData = fetchReportByDateRange(today, today);
    const count = saveReportToSheet(reportData, true); // 기존 데이터 업데이트
    Logger.log(`완료! ${count}개 데이터 처리됨`);
    return count;
  } catch (e) {
    Logger.log('오류 발생: ' + e.message);
    throw e;
  }
}

// ============================================
// ★ 최근 2일 데이터 가져오기 (어제 + 오늘)
// ============================================
function fetchRecentData() {
  try {
    Logger.log('최근 2일 데이터 가져오기 시작...');
    
    const today = new Date();
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    
    const reportData = fetchReportByDateRange(yesterday, today);
    const count = saveReportToSheet(reportData, true); // 기존 데이터 업데이트
    
    Logger.log(`완료! ${count}개 데이터 처리됨`);
    return count;
  } catch (e) {
    Logger.log('오류 발생: ' + e.message);
    throw e;
  }
}

// ============================================
// 어제 데이터 가져오기 (기존 - 자동 실행용)
// ============================================
function fetchDailyReport() {
  try {
    Logger.log('AdMob 일일 리포트 가져오기 시작...');
    
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    
    const reportData = fetchReportByDateRange(yesterday, yesterday);
    const count = saveReportToSheet(reportData, false);
    
    Logger.log(`완료! ${count}개 데이터 저장됨`);
  } catch (e) {
    Logger.log('오류 발생: ' + e.message);
    throw e;
  }
}

// ============================================
// 데이터 새로고침 (수동)
// ============================================
function refreshData() {
  fetchRecentData();
}

// ============================================
// 최근 30일 데이터 가져오기
// ============================================
function fetchLast30Days() {
  try {
    Logger.log('최근 30일 데이터 가져오기 시작...');
    
    const today = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - 30);
    
    const reportData = fetchReportByDateRange(startDate, today);
    const count = saveReportToSheet(reportData, true);
    
    Logger.log(`완료! ${count}개 데이터 처리됨`);
  } catch (e) {
    Logger.log('오류 발생: ' + e.message);
    throw e;
  }
}

// ============================================
// 시트 초기화 (헤더 다시 설정)
// ============================================
function resetSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('daily');
  sheet.clear();
  sheet.getRange(1, 1, 1, 8).setValues([
    ['date', 'app_name', 'ad_unit', 'ad_format', 'impressions', 'clicks', 'revenue', 'ecpm']
  ]);
  Logger.log('시트가 초기화되었습니다.');
}

// ============================================
// 테스트: 연결 확인
// ============================================
function testConnection() {
  try {
    Logger.log('AdMob API 연결 테스트...');
    const accountId = getAdMobAccountId();
    Logger.log('성공! 계정 ID: ' + accountId);
    return accountId;
  } catch (e) {
    Logger.log('연결 실패: ' + e.message);
    throw e;
  }
}

// ============================================
// 웹앱용 API
// ============================================
function doGet(e) {
  const action = e.parameter.action || 'getData';
  
  try {
    let result;
    switch (action) {
      case 'getData':
        result = getAllData();
        break;
      case 'refresh':
        // 최근 2일 데이터 새로 가져오기 (오늘 + 어제)
        fetchRecentData();
        result = getAllData();
        result.refreshed = true;
        result.refreshedAt = new Date().toISOString();
        break;
      case 'refreshToday':
        // 오늘 데이터만 새로 가져오기
        fetchTodayReport();
        result = getAllData();
        result.refreshed = true;
        result.refreshedAt = new Date().toISOString();
        break;
      default:
        result = getAllData();
    }
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// 전체 데이터 가져오기
// ============================================
function getAllData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('daily');
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return { data: [], summary: {} };
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  
  const formatted = data.map(row => ({
    date: formatDateToString(row[0]),
    app_name: row[1],
    ad_unit: row[2],
    ad_format: row[3],
    impressions: row[4],
    clicks: row[5],
    revenue: row[6],
    ecpm: row[7]
  }));
  
  // 날짜 내림차순 정렬
  formatted.sort((a, b) => b.date.localeCompare(a.date));
  
  return {
    data: formatted,
    summary: calculateSummary(formatted),
    lastUpdated: new Date().toISOString()
  };
}

// ============================================
// 요약 통계 계산
// ============================================
function calculateSummary(data) {
  if (!data || data.length === 0) {
    return { totalRevenue: 0, totalImpressions: 0, totalClicks: 0, avgEcpm: 0, avgCtr: 0 };
  }
  
  const totalRevenue = data.reduce((sum, row) => sum + (row.revenue || 0), 0);
  const totalImpressions = data.reduce((sum, row) => sum + (row.impressions || 0), 0);
  const totalClicks = data.reduce((sum, row) => sum + (row.clicks || 0), 0);
  
  return {
    totalRevenue,
    totalImpressions,
    totalClicks,
    avgEcpm: totalImpressions > 0 ? (totalRevenue / totalImpressions) * 1000 : 0,
    avgCtr: totalImpressions > 0 ? (totalClicks / totalImpressions) * 100 : 0
  };
}
