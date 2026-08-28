/**
 * ============================================================================
 * WEB APP STABILITY FIXES
 * ============================================================================
 * Loaded after the existing application files.
 *
 * Goals:
 *  - keep the existing UI
 *  - make navigation resilient if an older client handler is missing
 *  - expose a health check for the spreadsheet/authentication layer
 *  - ensure the Users sheet/admin account can be initialized safely
 * ============================================================================
 */

function getAppSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error(
    'No spreadsheet is bound to this Apps Script project. Open the script from the spreadsheet (Extensions > Apps Script), or configure a spreadsheet ID in the project.'
  );
}

function getAppHealth() {
  try {
    const ss = getAppSpreadsheet_();
    const required = ['Users'];
    const optional = ['Incoming', 'Outgoing', 'IncDB', 'OutDB', 'Stock Balance', 'Inventory', 'Division Budget'];
    const sheets = ss.getSheets().map(s => s.getName());

    return {
      success: true,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      sheets: sheets,
      usersSheet: sheets.includes('Users'),
      requiredMissing: required.filter(n => !sheets.includes(n)),
      dataSheets: optional.filter(n => sheets.includes(n)),
      timestamp: new Date().toISOString()
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function initializeApplication() {
  try {
    const ss = getAppSpreadsheet_();
    const users = getUsersSheet_();

    // Guarantee the Users sheet has a usable header row.
    if (users.getLastColumn() < 8) {
      users.getRange(1, 1, 1, 8).setValues([[
        'Username', 'PasswordHash', 'Role', 'Name', 'Active',
        'MustChangePassword', 'LastLogin', 'Salt'
      ]]);
    }

    // Keep the built-in admin account available for first-time setup.
    createAdminAccount('admin123');

    return {
      success: true,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      message: 'Application initialized successfully.'
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Replace the web-app entry point with a version that adds a tiny recovery
 * layer to the existing HTML. The original getHtmlContent() remains intact.
 */
function doGet() {
  let html = getHtmlContent();

  const recoveryScript = `
<script>
(function () {
  'use strict';

  function stopLinkNavigation(e) {
    var href = this.getAttribute('href');
    if (href === '#') e.preventDefault();
  }

  function installNavigationRecovery() {
    document.querySelectorAll('a[href="#"]').forEach(function (a) {
      a.addEventListener('click', stopLinkNavigation, false);
    });

    // If the application already provides its own navigation function, use it.
    // Otherwise provide a safe fallback for common view/tab IDs.
    document.querySelectorAll('[data-view], [data-tab]').forEach(function (el) {
      if (el.dataset.recoveryBound) return;
      el.dataset.recoveryBound = '1';
      el.addEventListener('click', function (e) {
        var target = this.dataset.view || this.dataset.tab;
        if (!target) return;
        e.preventDefault();
        if (typeof window.switchView === 'function') return window.switchView(target);
        if (typeof window.switchTab === 'function') return window.switchTab(target);
      });
    });
  }

  // Make google.script.run failures visible instead of leaving the UI frozen.
  window.__asaServerError = function (err) {
    console.error('ASA server error:', err);
    var msg = err && err.message ? err.message : String(err || 'Unknown server error');
    var el = document.getElementById('appError') || document.getElementById('errorMessage');
    if (el) {
      el.textContent = msg;
      el.classList.remove('hidden');
    }
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', installNavigationRecovery);
  } else {
    installNavigationRecovery();
  }
})();
</script>`;

  html = html.replace(/<\\/body>/i, recoveryScript + '</body>');

  return HtmlService.createHtmlOutput(html)
    .setTitle('Logistics Inventory Dashboard & Delivery Receipts - ASA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
