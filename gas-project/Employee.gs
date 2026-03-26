// ============================================================
// Employee.gs - Sync and manage employee list
// ============================================================

// Column indices in the employee source sheet (0-based)
// AC = column 29 (到職日/加保), AE = column 31 (離職日), AK = column 37 (算入考核)
// 假設 A欄=姓名, B欄=部門 (需根據實際sheet調整)
var EMP_COL = {
  NAME:        0,  // A - 姓名
  DEPT:        1,  // B - 部門/科別
  EMP_ID:      2,  // C - 員工編號
  HIRE_DATE:   28, // AC - 到職日(加保)
  RESIGN_DATE: 30, // AE - 離職日
  IN_REVIEW:   36  // AK - 算入考核
};

/**
 * Sync employees from external Google Sheet.
 * Only syncs rows where AK column = '算入考核'.
 * Filters out employees resigned before start of current month.
 */
function syncEmployees() {
  var config = getConfig();
  var sourceSheetId = config.employeeSheetId;

  if (!sourceSheetId) {
    return { success: false, message: '未設定員工資料表ID' };
  }

  try {
    var sourceSheet = SpreadsheetApp.openById(sourceSheetId).getSheets()[0];
    var data = sourceSheet.getDataRange().getValues();

    if (data.length < 2) {
      return { success: false, message: '員工資料表無資料' };
    }

    var now = new Date();
    var startOfCurrentMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    var employees = [];

    // Start from row 2 (index 1) to skip header
    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // Must be marked 算入考核
      var inReview = row[EMP_COL.IN_REVIEW];
      if (!inReview || inReview.toString().indexOf('算入考核') === -1) continue;

      var name = (row[EMP_COL.NAME] || '').toString().trim();
      if (!name) continue;

      var dept = (row[EMP_COL.DEPT] || '').toString().trim();
      var empId = row[EMP_COL.EMP_ID] ? row[EMP_COL.EMP_ID].toString().trim() : ('EMP_' + (i + 1));

      // Parse hire date
      var hireDateRaw = row[EMP_COL.HIRE_DATE];
      var hireDate = parseFlexibleDate(hireDateRaw);
      if (!hireDate) continue; // Must have hire date

      // Parse resign date
      var resignDateRaw = row[EMP_COL.RESIGN_DATE];
      var resignDate = parseFlexibleDate(resignDateRaw);

      // Exclude: resigned before start of current month
      if (resignDate && resignDate < startOfCurrentMonth) continue;

      // Calculate tenure
      var tenureStr = calcTenure(hireDate, now);

      employees.push({
        id:           empId,
        name:         name,
        dept:         dept,
        hireDate:     Utilities.formatDate(hireDate, 'Asia/Taipei', 'yyyy/MM/dd'),
        resignDate:   resignDate ? Utilities.formatDate(resignDate, 'Asia/Taipei', 'yyyy/MM/dd') : '',
        tenure:       tenureStr,
        isResigned:   resignDate ? true : false,
        rowIndex:     i + 1
      });
    }

    // Write to cache sheet
    var ss = getMainSpreadsheet();
    var cacheSheet = ss.getSheetByName('員工名單快取');
    if (!cacheSheet) {
      cacheSheet = ss.insertSheet('員工名單快取');
    }

    // Clear and rewrite
    cacheSheet.clearContents();
    cacheSheet.getRange('A1:H1').setValues([[
      '員工ID', '姓名', '部門', '到職日', '離職日', '在職狀態', '年資', '更新時間'
    ]]);

    var updateTime = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    if (employees.length > 0) {
      var rows = employees.map(function(e) {
        return [e.id, e.name, e.dept, e.hireDate, e.resignDate,
                e.isResigned ? '離職' : '在職', e.tenure, updateTime];
      });
      cacheSheet.getRange(2, 1, rows.length, 8).setValues(rows);
    }

    Logger.log('Employee sync complete: ' + employees.length + ' employees');
    return {
      success: true,
      message: '同步完成，共 ' + employees.length + ' 位同仁',
      count: employees.length,
      employees: employees
    };

  } catch (e) {
    Logger.log('syncEmployees error: ' + e + '\n' + e.stack);
    return { success: false, message: '同步失敗: ' + e.message };
  }
}

/**
 * Get cached employees. If cache is empty, trigger sync.
 * Optionally filters by review period (to handle resigned before period).
 */
function getCachedEmployees(period) {
  var ss = getMainSpreadsheet();
  var cacheSheet = ss.getSheetByName('員工名單快取');
  var data = [];

  if (cacheSheet) {
    var allData = cacheSheet.getDataRange().getValues();
    if (allData.length > 1) {
      for (var i = 1; i < allData.length; i++) {
        var row = allData[i];
        if (!row[0]) continue;
        data.push({
          id:         row[0].toString(),
          name:       row[1].toString(),
          dept:       row[2].toString(),
          hireDate:   row[3].toString(),
          resignDate: row[4].toString(),
          status:     row[5].toString(),
          tenure:     row[6].toString()
        });
      }
    }
  }

  // If no cache, sync now
  if (data.length === 0) {
    var result = syncEmployees();
    if (result.success) data = result.employees || [];
  }

  // Filter by period: if employee resigned before the period's start month, exclude
  if (period && period.startDate) {
    var periodStart = new Date(period.startDate);
    var startOfPeriodMonth = new Date(periodStart.getFullYear(), periodStart.getMonth(), 1);

    data = data.filter(function(emp) {
      if (!emp.resignDate) return true;
      var rd = parseFlexibleDate(emp.resignDate);
      if (!rd) return true;
      // Resigned before start of period month: exclude
      return rd >= startOfPeriodMonth;
    });
  }

  return data;
}

// ============================================================
// Helper functions
// ============================================================

/**
 * Parse various date formats: Date object, 'yyyy/MM/dd', 'yyyy-MM-dd',
 * Taiwanese year '115/03/02', or Excel serial number.
 */
function parseFlexibleDate(raw) {
  if (!raw) return null;
  if (raw instanceof Date) return isNaN(raw.getTime()) ? null : raw;

  var str = raw.toString().trim();
  if (!str) return null;

  // Excel serial number
  if (/^\d{5}$/.test(str)) {
    var d = new Date((parseInt(str) - 25569) * 86400 * 1000);
    return isNaN(d.getTime()) ? null : d;
  }

  // Taiwanese year: 115/03/02 or 115-03-02
  var twMatch = str.match(/^(\d{2,3})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (twMatch) {
    var year = parseInt(twMatch[1]);
    if (year < 200) year += 1911; // Convert ROC year
    var m = parseInt(twMatch[2]) - 1;
    var day = parseInt(twMatch[3]);
    var d2 = new Date(year, m, day);
    return isNaN(d2.getTime()) ? null : d2;
  }

  // ISO: yyyy-MM-dd or yyyy/MM/dd
  var isoMatch = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (isoMatch) {
    var d3 = new Date(parseInt(isoMatch[1]), parseInt(isoMatch[2]) - 1, parseInt(isoMatch[3]));
    return isNaN(d3.getTime()) ? null : d3;
  }

  // Try native parse
  var d4 = new Date(str);
  return isNaN(d4.getTime()) ? null : d4;
}

/**
 * Calculate tenure between two dates.
 * Returns string like '3年2個月'.
 */
function calcTenure(hireDate, toDate) {
  if (!hireDate || !toDate) return '';
  var to = toDate || new Date();
  var years = to.getFullYear() - hireDate.getFullYear();
  var months = to.getMonth() - hireDate.getMonth();
  if (months < 0) { years--; months += 12; }
  if (years < 0) return '未滿1個月';
  if (years === 0) return months + '個月';
  return years + '年' + months + '個月';
}
