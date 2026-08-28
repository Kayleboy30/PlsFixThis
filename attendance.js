/**
 * ============================================================================
 * TIME & ATTENDANCE MODULE (Attendance.gs)
 * ============================================================================
 */

function getAttendanceSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance');
  if (!sheet) {
    sheet = ss.insertSheet('Attendance');
    sheet.appendRow(['Log ID', 'Date', 'User ID', 'Employee Name', 'Clock In', 'Clock Out', 'Total Hours', 'Regular Hours', 'Overtime Hours', 'Status', 'Remarks']);
    sheet.getRange('A1:K1').setFontWeight('bold').setBackground('#f1f5f9');
  }
  return sheet;
}

function fetchAttendanceData(token, filterDate) {
  try {
    const sheet = getAttendanceSheet_();
    const data = sheet.getDataRange().getValues();
    let currentActiveLog = null;
    let currentUser = null;

    if (token) {
      const cached = CacheService.getScriptCache().get('sess_' + token);
      if (cached) currentUser = JSON.parse(cached);
    }

    if (data.length <= 1) {
      return { success: true, logs: [], activeLog: null };
    }

    const logs = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const log = {
        rowIndex: i + 1,
        logId: String(r[0] || ''),
        date: r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone() || 'GMT+8', 'yyyy-MM-dd') : String(r[1] || ''),
        userId: String(r[2] || ''),
        employeeName: String(r[3] || ''),
        clockIn: String(r[4] || ''),
        clockOut: String(r[5] || ''),
        totalHours: Number(r[6]) || 0,
        regularHours: Number(r[7]) || 0,
        overtimeHours: Number(r[8]) || 0,
        status: String(r[9] || 'Completed'),
        remarks: String(r[10] || '')
      };

      if (filterDate && log.date !== filterDate) continue;

      if (currentUser && log.userId === currentUser.userId && (!log.clockOut || log.status === 'Active')) {
        currentActiveLog = log;
      }
      logs.push(log);
    }

    logs.reverse(); // Latest records first
    return { success: true, logs: logs, activeLog: currentActiveLog };
  } catch (err) {
    return { success: false, error: err.message, logs: [] };
  }
}

function recordAttendanceClock(token, action, remarks) {
  try {
    const user = requireRole_(token, ['ADMIN', 'ENCODER', 'VIEWER']);
    const sheet = getAttendanceSheet_();
    const data = sheet.getDataRange().getValues();
    const tz = Session.getScriptTimeZone() || 'GMT+8';
    const now = new Date();
    const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const timeStr = Utilities.formatDate(now, tz, 'hh:mm:ss a');

    let activeRowIndex = -1;
    let clockInTimeStr = '';

    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const uId = String(r[2] || '');
      const cOut = String(r[5] || '');
      const st = String(r[9] || '');
      if (uId === user.userId && (!cOut || st === 'Active')) {
        activeRowIndex = i + 1;
        clockInTimeStr = String(r[4] || '');
        break;
      }
    }

    if (action === 'CLOCK_IN') {
      if (activeRowIndex > 1) {
        throw new Error('You already have an active Clock In session. Please Clock Out first.');
      }
      const logId = 'ATT-' + Utilities.formatString('%05d', data.length);
      const newRow = [
        logId, todayStr, user.userId, user.fullName || user.username,
        timeStr, '', 0, 0, 0, 'Active', remarks || 'Shift started'
      ];
      sheet.appendRow(newRow);
      return { success: true, message: 'Clocked in successfully at ' + timeStr };
    } else if (action === 'CLOCK_OUT') {
      if (activeRowIndex <= 1) {
        throw new Error('No active Clock In session found for your account.');
      }

      let totalHrs = 0;
      try {
        const d1 = new Date(todayStr + ' ' + clockInTimeStr);
        const d2 = now;
        const diffMs = Math.max(0, d2.getTime() - d1.getTime());
        totalHrs = Number((diffMs / 3600000).toFixed(2));
      } catch (e) {
        totalHrs = 8.0;
      }

      const regHrs = Math.min(8.0, totalHrs);
      const otHrs = Math.max(0, totalHrs - 8.0);

      sheet.getRange(activeRowIndex, 6).setValue(timeStr);
      sheet.getRange(activeRowIndex, 7).setValue(totalHrs);
      sheet.getRange(activeRowIndex, 8).setValue(regHrs);
      sheet.getRange(activeRowIndex, 9).setValue(otHrs);
      sheet.getRange(activeRowIndex, 10).setValue('Completed');
      if (remarks) sheet.getRange(activeRowIndex, 11).setValue(remarks);

      return { success: true, message: 'Clocked out successfully at ' + timeStr + '. Total Hours: ' + totalHrs + 'h' };
    }

    throw new Error('Invalid clock action');
  } catch (err) {
    return { success: false, error: err.message };
  }
}
