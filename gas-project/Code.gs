// ============================================================
// Code.gs - Main entry point and routing
// ============================================================

/**
 * Web app entry point: serves HTML based on user role.
 */
function doGet(e) {
  var userEmail = Session.getActiveUser().getEmail();
  var page = (e.parameter && e.parameter.page) ? e.parameter.page : 'main';

  // Check if user is admin
  var config = getConfig();
  var isAdmin = config.adminEmails && config.adminEmails.indexOf(userEmail) !== -1;

  // Identify supervisor role
  var supervisorRole = getSupervisorRole(userEmail, config);

  if (!isAdmin && !supervisorRole) {
    return HtmlService.createHtmlOutput(
      '<h2 style="font-family:sans-serif;color:#c0392b;padding:40px">存取被拒絕 Access Denied</h2>' +
      '<p style="font-family:sans-serif;padding:0 40px">您的帳號 ' + userEmail + ' 未被授權使用此系統。</p>'
    ).setTitle('員工考核系統');
  }

  var template = HtmlService.createTemplateFromFile('index');
  template.userEmail = userEmail;
  template.isAdmin = isAdmin;
  template.supervisorRole = supervisorRole || '';
  template.page = page;

  return template.evaluate()
    .setTitle('員工考核系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Include HTML sub-files (CSS, JS partials).
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// Server-side API functions called via google.script.run
// ============================================================

/** Returns current user info */
function getCurrentUser() {
  var email = Session.getActiveUser().getEmail();
  var config = getConfig();
  var isAdmin = config.adminEmails && config.adminEmails.indexOf(email) !== -1;
  var supervisorRole = getSupervisorRole(email, config);
  return { email: email, isAdmin: isAdmin, supervisorRole: supervisorRole };
}

/** Returns full config (admin only for sensitive parts) */
function getConfigForUI() {
  var email = Session.getActiveUser().getEmail();
  var config = getConfig();
  var isAdmin = config.adminEmails && config.adminEmails.indexOf(email) !== -1;
  if (!isAdmin) {
    // Strip admin email list for non-admins
    var safeConfig = JSON.parse(JSON.stringify(config));
    delete safeConfig.adminEmails;
    return safeConfig;
  }
  return config;
}

/** Admin: save full config */
function saveConfigFromUI(configJson) {
  requireAdmin();
  saveConfig(configJson);
  return { success: true };
}

/** Admin: manually trigger employee sync */
function syncEmployeesFromUI() {
  requireAdmin();
  return syncEmployees();
}

/** Get employees for a scoring period (supervisor-facing) */
function getEmployeesForScoring(periodId) {
  var email = Session.getActiveUser().getEmail();
  var config = getConfig();
  var supervisorRole = getSupervisorRole(email, config);
  var isAdmin = config.adminEmails && config.adminEmails.indexOf(email) !== -1;

  if (!supervisorRole && !isAdmin) throw new Error('未授權');

  var period = getPeriodById(config, periodId);
  if (!period) throw new Error('找不到考核期間');

  // Get cached employees
  var employees = getCachedEmployees(period);

  // Filter employees assigned to this supervisor role
  if (!isAdmin && supervisorRole) {
    employees = filterEmployeesBySupervisor(employees, supervisorRole, config);
  }

  return employees;
}

/** Get all employees for admin view */
function getAllEmployeesForAdmin(periodId) {
  requireAdmin();
  var config = getConfig();
  var period = getPeriodById(config, periodId);
  return getCachedEmployees(period);
}

/** Get scoring items for a dept */
function getScoringItemsForDept(dept) {
  var config = getConfig();
  return getScoringItemsByDept(config, dept);
}

/** Save scores submitted by a supervisor */
function saveScores(periodId, employeeId, scores) {
  var email = Session.getActiveUser().getEmail();
  var config = getConfig();
  var supervisorRole = getSupervisorRole(email, config);
  var isAdmin = config.adminEmails && config.adminEmails.indexOf(email) !== -1;
  if (!supervisorRole && !isAdmin) throw new Error('未授權');

  var period = getPeriodById(config, periodId);
  if (!period) throw new Error('找不到考核期間');

  // Check deadline
  var now = new Date();
  var deadline = new Date(period.deadlineDate + 'T23:59:59+08:00');
  if (now > deadline) throw new Error('評分截止日期已過 (' + period.deadlineDate + ')');

  return writeScores(periodId, employeeId, supervisorRole || 'admin', email, scores);
}

/** Get existing scores for a supervisor/period */
function getExistingScores(periodId, supervisorRole) {
  var email = Session.getActiveUser().getEmail();
  var config = getConfig();
  var myRole = getSupervisorRole(email, config);
  var isAdmin = config.adminEmails && config.adminEmails.indexOf(email) !== -1;
  if (!myRole && !isAdmin) throw new Error('未授權');

  var role = supervisorRole || myRole;
  return readScores(periodId, role);
}

/** Admin: get calculated results for a period */
function getCalculatedResults(periodId) {
  requireAdmin();
  return calculateResults(periodId);
}

/** Admin: get scoring completion status */
function getScoringStatus(periodId) {
  requireAdmin();
  return buildScoringStatus(periodId);
}

/** Admin: export to Excel-friendly format */
function exportResultsData(periodId) {
  requireAdmin();
  return buildExportData(periodId);
}

/** Admin: get list of review periods */
function getReviewPeriods() {
  var config = getConfig();
  return config.reviewPeriods || [];
}

/** Get active/current review period */
function getActivePeriod() {
  var config = getConfig();
  var periods = config.reviewPeriods || [];
  var now = new Date();
  // Find period where today is within deadline window
  for (var i = 0; i < periods.length; i++) {
    var p = periods[i];
    var start = new Date(p.startDate);
    var deadline = new Date(p.deadlineDate + 'T23:59:59+08:00');
    if (now >= start && now <= deadline) return p;
  }
  // Return latest if none active
  if (periods.length > 0) return periods[periods.length - 1];
  return null;
}

// ============================================================
// Time-driven triggers (set up via admin UI)
// ============================================================

/** Send reminder notifications - called by time trigger */
function sendReminderNotifications() {
  var config = getConfig();
  var periods = config.reviewPeriods || [];
  var now = new Date();
  var todayStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd');

  periods.forEach(function(period) {
    var notifDates = period.notificationDates || [];
    var notifTime = period.notificationTime || '09:00';

    notifDates.forEach(function(nd) {
      if (nd === todayStr) {
        var currentHour = parseInt(Utilities.formatDate(now, 'Asia/Taipei', 'HH'));
        var targetHour = parseInt(notifTime.split(':')[0]);
        // Trigger is set hourly; check if within the right hour
        if (currentHour === targetHour) {
          sendUnsubmittedReminders(period, config);
        }
      }
    });
  });
}

/** Install time-driven triggers */
function installTriggers() {
  requireAdmin();
  // Remove existing
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendReminderNotifications') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Add hourly trigger
  ScriptApp.newTrigger('sendReminderNotifications')
    .timeBased()
    .everyHours(1)
    .create();
  return { success: true, message: '已設定每小時提醒觸發器' };
}

// ============================================================
// Helper
// ============================================================
function requireAdmin() {
  var email = Session.getActiveUser().getEmail();
  var config = getConfig();
  if (!config.adminEmails || config.adminEmails.indexOf(email) === -1) {
    throw new Error('需要管理員權限');
  }
}

function getPeriodById(config, periodId) {
  var periods = config.reviewPeriods || [];
  for (var i = 0; i < periods.length; i++) {
    if (periods[i].id === periodId) return periods[i];
  }
  return null;
}

function filterEmployeesBySupervisor(employees, supervisorRole, config) {
  var deptConfigs = config.deptReviewConfig || [];
  var allowedDepts = [];

  deptConfigs.forEach(function(dc) {
    var reviewers = dc.reviewers || [];
    reviewers.forEach(function(r) {
      if (r.role === supervisorRole) {
        allowedDepts.push(dc.dept);
      }
    });
  });

  return employees.filter(function(emp) {
    return allowedDepts.indexOf(emp.dept) !== -1;
  });
}

function getSupervisorRole(email, config) {
  var supervisors = config.supervisors || [];
  for (var i = 0; i < supervisors.length; i++) {
    if (supervisors[i].email === email) return supervisors[i].role;
  }
  return null;
}
