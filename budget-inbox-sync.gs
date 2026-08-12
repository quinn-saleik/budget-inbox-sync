/**
 * Budget Inbox Sync — Personal Budget Apps Script
 * ETL: Gmail (banking notifications) → Google Sheets
 *
 * One-time setup: run `bootstrapAll()` from the Apps Script editor.
 * Then seed the Categories sheet with your keyword rules.
 *
 * See bottom of file for trigger setup instructions.
 */

// ============================================================================
// CONFIG
// ============================================================================

const DEFAULT_CONFIG = {
  lookback_days: '3',
  bank_senders: 'capitalone,venmo,chase',
  chase_balance_cell: 'E1',
  digest_recipient: '',
  digest_hours_back: '24'
  // Note: fidelity_<last4>_cell entries aren't defaulted here — they're
  // account-specific and get added to the Config sheet as needed (see
  // fetchFidelityBalance, which logs the exact key to add the first time
  // it sees a new account).
};

function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const config = Object.assign({}, DEFAULT_CONFIG);
  if (!sheet || sheet.getLastRow() < 2) return config;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  data.forEach(row => {
    const key = String(row[0] || '').trim();
    const val = String(row[1] || '').trim();
    if (key) config[key] = val;
  });
  return config;
}

function cfg(key) {
  return getConfig()[key] || DEFAULT_CONFIG[key] || '';
}

// ============================================================================
// CATEGORY RULES
// ============================================================================

function loadCategoryRules() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return data
    .map(row => ({
      keyword: String(row[0] || '').trim().toLowerCase(),
      category: String(row[1] || '').trim()
    }))
    .filter(r => r.keyword && r.category);
}

function applyCategoryRules(description, rules) {
  if (!description) return '';
  const lower = String(description).toLowerCase();
  for (const rule of rules) {
    if (lower.includes(rule.keyword)) return rule.category;
  }
  return '';
}

/**
 * Backfill empty Category cells in Transactions (Auto) using current rules.
 * Never overwrites a non-empty (manual) category.
 */
function recategorizeAll() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions (Auto)');
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('No Transactions (Auto) data to recategorize.');
    return;
  }
  const rules = loadCategoryRules();
  if (rules.length === 0) {
    Logger.log('No category rules found — seed the Categories sheet first.');
    return;
  }
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6);
  const values = range.getValues();
  let updated = 0;
  values.forEach(row => {
    if (!row[5]) {
      const cat = applyCategoryRules(row[3], rules);
      if (cat) { row[5] = cat; updated++; }
    }
  });
  range.setValues(values);
  Logger.log(`Recategorized ${updated} rows.`);
}

// ============================================================================
// FETCH TRANSACTION EMAILS
// ============================================================================

function fetchTransactionEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions (Auto)');
  if (!sheet) {
    Logger.log('Transactions (Auto) sheet not found.');
    return;
  }

  const lastRow = sheet.getLastRow();
  const existingTransactionIDs = lastRow < 2
    ? []
    : sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();

  const lookbackDays = parseInt(cfg('lookback_days'), 10) || 3;
  const sinceDate = new Date(Date.now() - lookbackDays * 24 * 60 * 60 * 1000);
  const dateString = Utilities.formatDate(sinceDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  const senders = cfg('bank_senders').split(',').map(s => s.trim()).filter(Boolean);
  const fromClause = senders.map(s => `from:${s}`).join(' OR ');
  const query = `(${fromClause}) after:${dateString}`;

  const threads = GmailApp.search(query);
  const rules = loadCategoryRules();

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const sender = message.getFrom().toLowerCase();
      const subject = message.getSubject() || '';
      const body = message.getPlainBody();

      let appName = null;
      if (sender.includes('capitalone')) appName = 'Capital One';
      else if (sender.includes('venmo')) appName = 'Venmo';
      else if (sender.includes('chase')) appName = 'Chase';
      else if (sender.includes('americanexpress')) appName = 'Amex';
      if (!appName) return;

      const parsed = parseTransaction(appName, subject, body);
      if (!parsed) return;

      const { amount, description } = parsed;
      const timestamp = message.getDate();
      const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');
      const transactionId = `${appName}|${formattedDate}|${amount}|${description}`;

      if (!existingTransactionIDs.includes(transactionId)) {
        const category = applyCategoryRules(description, rules);
        sheet.insertRowAfter(2);
        sheet.getRange(3, 1, 1, 6).setValues([[
          transactionId, formattedDate, amount, description, appName, category
        ]]);
        existingTransactionIDs.push(transactionId);
      }
    });
  });
}

function parseTransaction(appName, subject, body) {
  if (appName === 'Capital One') return parseCapitalOne(subject, body);
  if (appName === 'Venmo') return parseVenmo(subject, body);
  if (appName === 'Chase') return parseChase(subject, body);
  if (appName === 'Amex') return parseAmEx(subject, body);
  return null;
}

function parseCapitalOne(subject, body) {
  const lower = body.toLowerCase();

  // Skip Venmo-related withdrawals (Venmo parser handles these)
  if (lower.includes('withdrawal notice') && lower.includes('venmo')) return null;

  if (lower.includes('withdrawal notice')) {
    const amountMatch = body.match(/Amount:\s*\$([\d,.]+)/);
    const descriptionMatch = body.match(/withdrawal notice\s*([^\n]+?)(?=\shas\sinitiated)/i);
    if (amountMatch && descriptionMatch) {
      return {
        amount: -Math.abs(parseAmount(amountMatch[1])),
        description: descriptionMatch[1].trim()
      };
    }
    return null;
  }

  // Standard purchase — handles "April 25, 2026" OR "Apr. 25, 2026"
  const purchaseMatch = body.match(
    /on ([\w.]+ \d+, \d+), at (.+?), a pending authorization or purchase in the amount of \$([\d,.]+) was placed or charged/i
  );
  if (purchaseMatch) {
    return {
      amount: -Math.abs(parseAmount(purchaseMatch[3])),
      description: purchaseMatch[2].trim()
    };
  }

  if (lower.includes('rewards credit')) {
    const m = body.match(/\$([\d,.]+)/);
    if (m) return { amount: Math.abs(parseAmount(m[1])), description: 'Rewards Credit' };
  }
  if (lower.includes('refund')) {
    const m = body.match(/\$([\d,.]+)/);
    if (m) return { amount: Math.abs(parseAmount(m[1])), description: 'Refund' };
  }
  if (lower.includes('card payment')) {
    const m = body.match(/\$([\d,.]+)/);
    if (m) return { amount: Math.abs(parseAmount(m[1])), description: 'Card Payment' };
  }
  if (lower.includes('payment posted')) {
    const m = body.match(/\$([\d,.]+)/);
    if (m) return { amount: Math.abs(parseAmount(m[1])), description: 'Payment Posted' };
  }
  if (lower.includes('card transaction notice')) {
    const amountMatch = body.match(/Amount:\s*\$([\d,.]+)/);
    const descriptionMatch = body.match(/Description:\s*(.+)/);
    if (amountMatch && descriptionMatch) {
      return {
        amount: -Math.abs(parseAmount(amountMatch[1])),
        description: descriptionMatch[1].trim()
      };
    }
  }
  return null;
}

function parseVenmo(subject, body) {
  const lower = body.toLowerCase();

  // Counterparty from subject: "John Doe paid you $25" / "You paid John Doe $25"
  let counterparty = '';
  const subjPaidYou = subject.match(/^(.*?)\s+paid you/i);
  const subjYouPaid = subject.match(/You paid\s+(.+?)\s*\$/i);
  if (subjPaidYou) counterparty = subjPaidYou[1].trim();
  else if (subjYouPaid) counterparty = subjYouPaid[1].trim();

  const note = extractVenmoNote(body);

  let description = 'Venmo';
  if (counterparty) description = `Venmo: ${counterparty}`;
  if (note) description += ` — ${note}`;

  if (lower.includes('money credited') || lower.includes('paid you')) {
    const m = body.match(/\$([\d,.]+)/);
    if (m) return { amount: Math.abs(parseAmount(m[1])), description };
  }
  if (lower.includes('you paid')) {
    const m = body.match(/\$\s*([\d,.]+)/);
    if (m) return { amount: -Math.abs(parseAmount(m[1])), description };
  }
  return null;
}

function extractVenmoNote(body) {
  const m = body.match(/Note:\s*(.+?)(?:\n|$)/i);
  if (m) return m[1].trim().slice(0, 80);
  return '';
}

function parseChase(subject, body) {
  const lower = body.toLowerCase();

  // Outgoing purchase
  if (lower.includes('you made a')) {
    const amountMatch = body.match(/\$([\d,.]+)/);
    const descMatch = body.match(/Merchant\s+(.*?)\s+Amount/i);
    if (amountMatch && descMatch) {
      return {
        amount: -Math.abs(parseAmount(amountMatch[1])),
        description: descMatch[1].trim()
      };
    }
  }

  // Card payment
  if (lower.includes('payment scheduled') || lower.includes('payment posted') || lower.includes('payment received')) {
    const m = body.match(/\$([\d,.]+)/);
    if (m) return { amount: Math.abs(parseAmount(m[1])), description: 'Chase Card Payment' };
  }

  // Refund / credit
  if (lower.includes('refund') || lower.includes('credit posted')) {
    const m = body.match(/\$([\d,.]+)/);
    if (m) return { amount: Math.abs(parseAmount(m[1])), description: 'Chase Refund' };
  }

  // Deposit
  if (lower.includes('deposit') && lower.includes('posted')) {
    const m = body.match(/\$([\d,.]+)/);
    if (m) return { amount: Math.abs(parseAmount(m[1])), description: 'Chase Deposit' };
  }

  return null;
}

function parseAmEx(subject, body) {
  const lower = body.toLowerCase();

  // Large purchase notification (merchant name on its own line, followed by $amount*)
  if (lower.includes('large purchase')) {
    const m = body.match(/\n([A-Z0-9][A-Z0-9 .,&'-]*[A-Z0-9])\s*\n\$([\d,.]+)\*?/);
    if (m) {
      return {
        amount: -Math.abs(parseAmount(m[2])),
        description: m[1].trim()
      };
    }
  }

  return null;
}

function parseAmount(str) {
  return parseFloat(String(str).replace(/,/g, ''));
}

// ============================================================================
// FETCH BALANCES
// ============================================================================

function fetchChaseBalance() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Widget');
  if (!sheet) return;

  const yesterday = new Date(Date.now() - 24 * 60 * 60 * 1000);
  const dateString = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const threads = GmailApp.search(`from:no.reply.alerts@chase.com after:${dateString}`);

  for (const thread of threads) {
    for (const message of thread.getMessages()) {
      const subject = message.getSubject();
      const match = subject.match(/Your Chase Sapphire Preferred Visa balance is \$([\d,]+\.\d{2})/);
      if (match) {
        const balance = parseFloat(match[1].replace(/,/g, ''));
        sheet.getRange(cfg('chase_balance_cell')).setValue(balance);
        appendBalanceHistory('Chase Sapphire', balance);
        return;
      }
    }
  }
}

function fetchFidelityBalance() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Widget');
  if (!sheet) return;

  const yesterday = new Date(Date.now() - 24 * 60 * 60 * 1000);
  const dateString = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const threads = GmailApp.search(`from:fidelity.com subject:"Account Summary" after:${dateString}`);

  // Pick the latest email per account if multiple arrive same day
  const latestByAccount = {};

  for (const thread of threads) {
    for (const message of thread.getMessages()) {
      const body = message.getPlainBody();
      const accountMatch = body.match(/Account:\s*X+(\d{4})/);
      const balanceMatch = body.match(/Total Account Value\s*\$([\d,]+\.\d{2})/);
      if (!accountMatch || !balanceMatch) continue;

      const last4 = accountMatch[1];
      const balance = parseFloat(balanceMatch[1].replace(/,/g, ''));
      const date = message.getDate();

      if (!latestByAccount[last4] || date > latestByAccount[last4].date) {
        latestByAccount[last4] = { balance, date };
      }
    }
  }

  for (const last4 in latestByAccount) {
    const cellKey = `fidelity_${last4}_cell`;
    const cell = cfg(cellKey);
    if (!cell) {
      Logger.log(`No mapping for Fidelity *${last4}. Add "${cellKey}" to Config sheet.`);
      continue;
    }
    const balance = latestByAccount[last4].balance;
    sheet.getRange(cell).setValue(balance);
    appendBalanceHistory(`Fidelity *${last4}`, balance);
  }
}

function appendBalanceHistory(account, balance) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BalanceHistory');
  if (!sheet) {
    sheet = ss.insertSheet('BalanceHistory');
    sheet.getRange(1, 1, 1, 3).setValues([['Date', 'Account', 'Balance']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const lastRow = sheet.getLastRow();

  // Update if same-day row exists for this account
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < data.length; i++) {
      const rowDate = data[i][0] instanceof Date
        ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(data[i][0]);
      if (rowDate === today && data[i][1] === account) {
        sheet.getRange(i + 2, 3).setValue(balance);
        return;
      }
    }
  }

  sheet.appendRow([today, account, balance]);
}

// ============================================================================
// DAILY DIGEST EMAIL
// ============================================================================

function sendDailyDigest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName('Transactions (Auto)');
  const widgetSheet = ss.getSheetByName('Widget');

  const hoursBack = parseInt(cfg('digest_hours_back'), 10) || 24;
  const cutoff = new Date(Date.now() - hoursBack * 60 * 60 * 1000);

  let recent = [];
  if (transSheet && transSheet.getLastRow() >= 2) {
    const data = transSheet.getRange(2, 1, transSheet.getLastRow() - 1, 6).getValues();
    recent = data.filter(row => {
      const d = parseAutoDate(row[1]);
      return d && d >= cutoff;
    });
  }

  const config = getConfig();
  const chaseBalance = widgetSheet ? widgetSheet.getRange(cfg('chase_balance_cell')).getValue() : null;

  // Pull every configured Fidelity balance cell generically instead of
  // hardcoding a fixed number of accounts.
  const balanceLines = [];
  if (chaseBalance != null && chaseBalance !== '') {
    balanceLines.push(`Chase Sapphire: $${Number(chaseBalance).toFixed(2)}`);
  }
  if (widgetSheet) {
    Object.keys(config)
      .filter(k => k.startsWith('fidelity_') && k.endsWith('_cell'))
      .forEach((k, i) => {
        const val = widgetSheet.getRange(config[k]).getValue();
        if (val != null && val !== '') balanceLines.push(`Fidelity (${i + 1}): $${Number(val).toFixed(2)}`);
      });
  }

  let html = `<div style="font-family:sans-serif;font-size:14px;">`;
  html += `<h2 style="margin-bottom:4px;">Daily Budget Digest</h2>`;
  html += `<p style="color:#666;margin-top:0;">Last ${hoursBack}h: ${recent.length} transaction${recent.length === 1 ? '' : 's'}</p>`;

  if (recent.length > 0) {
    html += `<table cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:13px;width:100%;">`;
    html += `<tr style="background:#f5f5f5;text-align:left;">`;
    html += `<th>Date</th><th>Amount</th><th>Description</th><th>Account</th><th>Category</th></tr>`;
    recent.forEach(row => {
      const amt = typeof row[2] === 'number' ? row[2] : parseFloat(row[2]);
      const color = amt < 0 ? '#c00' : '#080';
      const sign = amt < 0 ? '-' : '+';
      html += `<tr style="border-top:1px solid #eee;">`;
      html += `<td style="white-space:nowrap;">${escapeHtml(String(row[1]))}</td>`;
      html += `<td style="color:${color};text-align:right;white-space:nowrap;">${sign}$${Math.abs(amt).toFixed(2)}</td>`;
      html += `<td>${escapeHtml(String(row[3]))}</td>`;
      html += `<td>${escapeHtml(String(row[4]))}</td>`;
      html += `<td>${escapeHtml(String(row[5] || ''))}</td>`;
      html += `</tr>`;
    });
    html += `</table>`;
  }

  html += `<h3 style="margin-top:24px;">Balances</h3><ul>`;
  balanceLines.forEach(line => { html += `<li>${line}</li>`; });
  html += `</ul></div>`;

  const recipient = cfg('digest_recipient') || Session.getActiveUser().getEmail();
  MailApp.sendEmail({
    to: recipient,
    subject: `Daily Budget Digest — ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d')}`,
    htmlBody: html
  });
}

function parseAutoDate(str) {
  if (str instanceof Date) return str;
  if (!str) return null;
  const m = String(str).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})/);
  if (!m) return null;
  return new Date(+m[3], +m[1] - 1, +m[2], +m[4], +m[5], +m[6]);
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[c]));
}

// ============================================================================
// BOOTSTRAP (run once from the Apps Script editor)
// ============================================================================

function bootstrapAll() {
  bootstrapConfigSheet();
  bootstrapCategoriesSheet();
  bootstrapBalanceHistorySheet();
  Logger.log('Bootstrap complete. Review Config and seed Categories sheet.');
}

function bootstrapConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');
  if (sheet) { Logger.log('Config sheet exists, skipping.'); return; }
  sheet = ss.insertSheet('Config');
  sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  const rows = Object.entries(DEFAULT_CONFIG).map(([k, v]) => [k, v]);
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  sheet.autoResizeColumns(1, 2);
}

function bootstrapCategoriesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Categories');
  if (sheet) { Logger.log('Categories sheet exists, skipping.'); return; }
  sheet = ss.insertSheet('Categories');
  sheet.getRange(1, 1, 1, 2).setValues([['Keyword', 'Category']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);
}

function bootstrapBalanceHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BalanceHistory');
  if (sheet) { Logger.log('BalanceHistory sheet exists, skipping.'); return; }
  sheet = ss.insertSheet('BalanceHistory');
  sheet.getRange(1, 1, 1, 3).setValues([['Date', 'Account', 'Balance']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
}

// ============================================================================
// DEBUG HELPERS
// ============================================================================

function debugFidelityEmail() {
  const threads = GmailApp.search('from:fidelity.com subject:"Account Summary"', 0, 5);
  threads.forEach((thread, i) => {
    const message = thread.getMessages()[0];
    Logger.log(`--- Email ${i + 1} ---`);
    Logger.log(`Subject: ${message.getSubject()}`);
    Logger.log(`Date: ${message.getDate()}`);
    Logger.log(`Body excerpt: ${message.getPlainBody().substring(0, 500)}`);
  });
}

// ============================================================================
// EXISTING: WEEKLY AI REPORT (preserved from original)
// ============================================================================

function sendComprehensiveWeeklyReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var transSheet = ss.getSheetByName("All Transactions");
  var budgetSheet = ss.getSheetByName("Budget");
  var trendsSheet = ss.getSheetByName("Trends");

  if (!transSheet || !budgetSheet || !trendsSheet) {
    MailApp.sendEmail(Session.getActiveUser().getEmail(), "Error: Budget Script", "One or more sheets (Transactions, Budget, Trends) not found.");
    return;
  }

  var transString  = transSheet.getDataRange().getValues().map(r => r.join(" | ")).join("\n");
  var budgetString = budgetSheet.getDataRange().getValues().map(r => r.join(" | ")).join("\n");
  var trendsString = trendsSheet.getDataRange().getValues().map(r => r.join(" | ")).join("\n");

  var prompt = `
    Act as a professional financial analyst.
    Review the following data from a user's Google Sheet:

    --- TRANSACTIONS (Past Week) ---
    ${transString}

    --- BUDGET TARGETS ---
    ${budgetString}

    --- HISTORICAL TRENDS ---
    ${trendsString}
    --- END DATA ---
    Generate a brief, high-impact weekly financial report. Be direct and encouraging — not preachy.
    Use Markdown. Keep each section short: bullets over paragraphs, numbers over narrative.

    1. **This Week at a Glance** — One-line verdict (on/over/under budget) with the total spend vs. target and the single biggest driver.
    2. **Dining Out Check-In** — Current week vs. budget vs. recent trend. One sentence of context, not a lecture.
    3. **Where It's Leaking** — List only categories over budget, with $ and % overage. Skip the rest.
    4. **Seasonal Expense Runway** — [Example: rent/insurance/tuition increases for a set number of months, then resets]. Compute the total added cost vs. baseline. Since the increase will be funded from savings, show: current monthly savings rate → adjusted rate during the higher-expense window → whether savings growth stays positive or goes flat/negative. Confirm I'm still in a good position, or flag if not. If a small trim elsewhere would keep savings healthier, name it.
    5. **Three Moves for Next Week** — Three concrete, small actions (not "spend less on dining"). Each with an estimated $ impact.
    6. **Budget Tune-Up** — 1–2 line-item adjustments grounded in the trend data — where targets are unrealistically tight or consistently loose. Propose new numbers.

    Close with one encouraging sentence. No filler, no recap.
  `;

  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    MailApp.sendEmail(Session.getActiveUser().getEmail(), "Error: Budget Script", "API Key not found in Script Properties.");
    return;
  }

  var modelsToTry = [
    "gemini-flash-latest",
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-flash-lite-latest",
    "gemini-2.5-pro"
  ];

  var aiReport = null;
  var lastError = null;

  for (var m = 0; m < modelsToTry.length && !aiReport; m++) {
    try {
      aiReport = callGeminiWithBackoff(modelsToTry[m], apiKey, prompt);
    } catch (e) {
      lastError = e;
      Logger.log("Model " + modelsToTry[m] + " failed: " + e.message);
    }
  }

  var recipient = Session.getActiveUser().getEmail();
  if (aiReport) {
    MailApp.sendEmail({
      to: recipient,
      subject: "📊 AI-Powered Holistic Financial Analysis",
      htmlBody: aiReport.replace(/\n/g, '<br>')
    });
  } else {
    MailApp.sendEmail(recipient, "Error: Budget Script",
      "All Gemini models failed after retries.\nLast error: " + (lastError ? lastError.message : "unknown"));
  }
}

function callGeminiWithBackoff(model, apiKey, prompt) {
  var apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + apiKey;
  var payload = { "contents": [{ "parts": [{ "text": prompt }] }] };
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  var maxAttempts = 5;
  var baseDelayMs = 2000;

  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    var response = UrlFetchApp.fetch(apiUrl, options);
    var code = response.getResponseCode();
    var body = response.getContentText();
    var json;
    try { json = JSON.parse(body); } catch (e) { json = null; }

    if (code >= 200 && code < 300 && json && json.candidates && json.candidates.length > 0) {
      return json.candidates[0].content.parts[0].text;
    }

    var errMsg = (json && json.error && json.error.message) ? json.error.message : ("HTTP " + code + ": " + body);
    var retryable = (code === 429 || code === 500 || code === 502 || code === 503 || code === 504);

    if (!retryable || attempt === maxAttempts) {
      throw new Error(errMsg);
    }

    var delay = baseDelayMs * Math.pow(2, attempt - 1);
    var jitter = Math.floor(Math.random() * 1000);
    Utilities.sleep(delay + jitter);
  }

  throw new Error("Unreachable");
}
