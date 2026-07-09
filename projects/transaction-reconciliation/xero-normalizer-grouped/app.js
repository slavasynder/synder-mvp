// Xero Assisted Reconciliation — Normalizer (grouped-by-account variant)
// Reads a Xero "Account Transactions" export saved with **Grouping = Account**
// (produces per-account sections + running balance) and normalizes it to the
// same 7-column output as the flat-format normalizer. Grouped input has no
// `Account` column — instead each account is a labelled section terminated by
// a `Total <account>` row. Running balance passes through the parser silently
// (available for downstream balance-reconciliation work).

// ─── DOM plumbing ─────────────────────────────────────────────────────────────
const $ = (id) => document.getElementById(id);
const state = {
    workbook: null,
    parsed: null,       // { headers, rows, clearingAccounts, arRows, apRows, golRows }
    selectedClearing: null,
    result: null,       // { rows, summary, warnings }
};

const dropZone = $('dropZone'), fileInput = $('fileInput');
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drop-active'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drop-active'));
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drop-active');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFile(e.target.files[0]);
});
$('clearFile').addEventListener('click', (e) => { e.stopPropagation(); resetUI(); });
$('normalizeBtn').addEventListener('click', runNormalize);
$('downloadBtn').addEventListener('click', downloadResult);
$('tryDemoBtn').addEventListener('click', (e) => { e.stopPropagation(); loadDemoFile(); });

async function loadDemoFile() {
    try {
        const resp = await fetch('sample.xlsx');
        if (!resp.ok) throw new Error(`Couldn't fetch sample file (HTTP ${resp.status})`);
        const blob = await resp.blob();
        const file = new File([blob], 'sample.xlsx', { type: blob.type });
        handleFile(file);
    } catch (err) {
        showError(String(err));
    }
}

function resetUI() {
    state.workbook = null; state.parsed = null; state.result = null;
    fileInput.value = '';
    $('dropPrompt').classList.remove('hidden');
    $('dropFile').classList.add('hidden');
    $('parseError').classList.add('hidden');
    $('configStep').classList.add('hidden');
    $('resultStep').classList.add('hidden');
}

function showError(msg) {
    const el = $('parseError');
    el.textContent = msg;
    el.classList.remove('hidden');
}

async function handleFile(file) {
    resetUI();
    $('dropPrompt').classList.add('hidden');
    $('dropFile').classList.remove('hidden');
    $('fileName').textContent = file.name;
    $('fileMeta').textContent = `${(file.size / 1024).toFixed(1)} KB`;

    try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array', cellDates: true });
        state.workbook = wb;
        const parsed = parseWorkbook(wb);
        state.parsed = parsed;
        showConfigStep(parsed);
    } catch (err) {
        console.error(err);
        showError(`Couldn't read this file as a Xero Account Transactions export. ${err.message ?? err}`);
    }
}

function showConfigStep(parsed) {
    $('configStep').classList.remove('hidden');
    if (parsed.clearingAccounts.length > 1) {
        const picker = $('clearingPicker');
        picker.classList.remove('hidden');
        const sel = $('clearingSelect');
        sel.innerHTML = parsed.clearingAccounts.map(a => `<option value="${escapeHtml(a)}">${escapeHtml(a)}</option>`).join('');
        state.selectedClearing = parsed.clearingAccounts[0];
        sel.addEventListener('change', () => { state.selectedClearing = sel.value; });
    } else {
        state.selectedClearing = parsed.clearingAccounts[0];
        $('clearingPicker').classList.add('hidden');
    }
    // Show the base-currency input only when we parsed a single-currency file.
    if (!parsed.isMultiCurrency) {
        $('baseCurrencyPicker').classList.remove('hidden');
    } else {
        $('baseCurrencyPicker').classList.add('hidden');
    }
}

function runNormalize() {
    const mode = document.querySelector('input[name="mode"]:checked').value;
    const baseCurrency = ($('baseCurrencyInput').value || 'USD').trim().toUpperCase();
    try {
        state.result = normalize(state.parsed, state.selectedClearing, mode, { baseCurrency });
        renderResult(state.result);
    } catch (err) {
        console.error(err);
        showError(`Normalization failed: ${err.message ?? err}`);
    }
}

function renderResult({ rows, summary, warnings }) {
    $('resultStep').classList.remove('hidden');

    const cards = [
        ['Rows normalized', rows.length],
        ['Clearing account', state.selectedClearing],
        ['Mode', summary.mode === 'source' ? 'Source (PayPal-style)' : 'Converted (clearing)'],
        ['Source-substituted rows', summary.substitutedCount],
    ];
    $('summary').innerHTML = cards.map(([k, v]) => `
        <div class="bg-gray-50 rounded-lg p-3">
            <div class="text-xs font-medium text-gray-500 uppercase tracking-wider">${escapeHtml(k)}</div>
            <div class="text-sm font-semibold text-gray-900 mt-1 truncate" title="${escapeHtml(String(v))}">${escapeHtml(String(v))}</div>
        </div>
    `).join('');

    const body = $('previewBody');
    body.innerHTML = rows.slice(0, 25).map(r => `
        <tr>
            <td class="px-3 py-2 font-mono text-xs">${escapeHtml(r['Primary ID'] ?? '')}</td>
            <td class="px-3 py-2 font-mono text-xs">${escapeHtml(r['Secondary ID'] ?? '')}</td>
            <td class="px-3 py-2 text-right ${r.Amount < 0 ? 'text-red-600' : 'text-gray-900'}">${fmtNum(r.Amount)}</td>
            <td class="px-3 py-2">${escapeHtml(r.Currency ?? '')}</td>
            <td class="px-3 py-2">${fmtDate(r.Date)}</td>
            <td class="px-3 py-2">${escapeHtml(r.Type ?? '')}</td>
            <td class="px-3 py-2 text-gray-600 max-w-xs truncate" title="${escapeHtml(r.Description ?? '')}">${escapeHtml(r.Description ?? '')}</td>
        </tr>
    `).join('');

    const warnEl = $('warnings');
    warnEl.innerHTML = warnings.length
        ? `<div class="p-3 bg-amber-50 border border-amber-200 rounded text-sm text-amber-800">
             <div class="font-semibold mb-1">Notes (${warnings.length})</div>
             <ul class="list-disc ml-5 space-y-0.5">${warnings.map(w => `<li>${escapeHtml(w)}</li>`).join('')}</ul>
           </div>`
        : '';
}

function downloadResult() {
    if (!state.result) return;
    const rows = state.result.rows;
    const header = ['Primary ID', 'Secondary ID', 'Amount', 'Currency', 'Date', 'Type', 'Description'];
    const data = [header, ...rows.map(r => [
        r['Primary ID'] ?? '',
        r['Secondary ID'] ?? '',
        r.Amount,
        r.Currency ?? '',
        r.Date instanceof Date ? r.Date : (r.Date ? new Date(r.Date) : ''),
        r.Type ?? '',
        r.Description ?? '',
    ])];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 36 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 26 }, { wch: 60 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Normalized');
    const mode = state.result.summary.mode === 'source' ? 'PayPal' : 'non-PayPal';
    const filename = `Xero_Normalized_${mode}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, filename);
}

// ─── Utilities ────────────────────────────────────────────────────────────────
function fmtNum(n) { return typeof n === 'number' ? n.toFixed(2) : ''; }
function fmtDate(d) {
    if (!d) return '';
    const dt = d instanceof Date ? d : new Date(d);
    return isNaN(dt) ? String(d) : dt.toISOString().slice(0, 10);
}
function escapeHtml(s) {
    if (s == null) return '';
    return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ─── Parsing ──────────────────────────────────────────────────────────────────
const SPECIAL_ACCOUNTS = new Set(['Accounts Receivable', 'Accounts Payable', 'Realized Currency Gains']);

function parseWorkbook(wb) {
    // Grouped-format exports don't have a "4 in 1" sheet name — take the first
    // sheet whose title mentions "Account Transactions" or the first sheet in the workbook.
    const sheetName = wb.SheetNames.find(n => /account\s+transactions|group/i.test(n))
        || wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error('No sheet found.');
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });

    // Find header row: the first row containing `Date` AND either `Net (Source)` (multi-currency)
    // or plain `Net` (single-currency).
    const headerIdx = raw.findIndex(r =>
        Array.isArray(r) && r.includes('Date') && (r.includes('Net (Source)') || r.includes('Net'))
    );
    if (headerIdx === -1) throw new Error('Header row not found. Expected a row containing Date and either Net (Source) or Net.');
    const headers = raw[headerIdx];
    const idx = colIndexer(headers);

    // Detect variant:
    //   multi-currency  → header carries Net (Source) and a Net ({currency-code}) column
    //   single-currency → header carries plain Net (no Source variant, no Currency column)
    const hasNetSource = idx('Net (Source)') !== -1;
    const baseNetCol = headers.find(h =>
        typeof h === 'string' && /^Net \([A-Z]{3}\)$/.test(h) && h !== 'Net (Source)'
    );
    const isMultiCurrency = hasNetSource || !!baseNetCol;

    // Column resolution:
    //   netSourceCol → the row's own-currency Net column
    //   baseNetCol   → the base-currency Net column (equal to netSourceCol in single-currency)
    //   currencyCol  → per-row currency column ('Currency'), or null when the file has no such column
    const netSourceCol = isMultiCurrency ? 'Net (Source)' : 'Net';
    const resolvedBaseNetCol = isMultiCurrency ? baseNetCol : 'Net';
    const currencyCol = isMultiCurrency ? 'Currency' : null;

    // Also detect the Running Balance column. Multi-currency labels it with the
    // base currency code (`Running Balance (USD)`); single-currency drops the parens.
    const runningBalanceCol = headers.find(h =>
        typeof h === 'string' && (/^Running Balance \([A-Z]{3}\)$/.test(h) || h === 'Running Balance')
    );

    // Required columns per variant. Missing → hard error.
    const COMMON_REQUIRED = ['Date', 'Source', 'Contact', 'Description', 'Reference'];
    const missingCols = COMMON_REQUIRED.filter(c => idx(c) === -1);
    if (isMultiCurrency) {
        if (idx('Net (Source)') === -1) missingCols.push('Net (Source)');
        if (!baseNetCol) missingCols.push('Net ({currency}) — the base-currency Net column, e.g. Net (USD) in a USD org, Net (EUR) in a EUR org');
        if (idx('Currency') === -1) missingCols.push('Currency');
        if (!runningBalanceCol || !/^Running Balance \([A-Z]{3}\)$/.test(runningBalanceCol)) {
            missingCols.push('Running Balance ({currency}) — Xero doesn\'t always tick this by default when Grouping = Account is set; make sure to select it in the column picker');
        }
    } else {
        if (idx('Net') === -1) missingCols.push('Net');
        if (!runningBalanceCol || runningBalanceCol !== 'Running Balance') {
            missingCols.push('Running Balance — Xero doesn\'t always tick this by default when Grouping = Account is set; make sure to select it in the column picker');
        }
    }
    if (missingCols.length > 0) {
        throw new Error(`Required column(s) missing from the Xero export: ${missingCols.join('; ')}. Re-export from Xero → Reporting → Account Transactions with Grouping = Account and the full column set selected.`);
    }

    // Walk rows below the header, tracking the current account section.
    // A section starts on a row whose first cell is a non-Date string that
    // doesn't begin with "Total ". A section ends on a "Total <account>" row.
    // Data rows have a Date in the Date column.
    const dateIdx = idx('Date');
    const buckets = {};
    let currentAccount = null;
    const dataRows = []; // (kept for downstream logic that iterates all data rows)

    for (let i = headerIdx + 1; i < raw.length; i++) {
        const r = raw[i];
        if (!Array.isArray(r) || r.every(c => c == null || c === '')) continue;

        const first = r[0];
        const isDate = first instanceof Date;
        const isTotalRow = typeof first === 'string' && /^Total\b/i.test(first.trim());
        const isSectionHeader = typeof first === 'string'
            && !isTotalRow
            && !isDate
            && r.slice(1).every(c => c == null || c === '');

        if (isSectionHeader) {
            currentAccount = first.trim();
            continue;
        }
        if (isTotalRow) {
            currentAccount = null;
            continue;
        }
        // Data row — must be under a section
        if (currentAccount == null) continue;

        // Coerce numeric columns if they came through as strings (accounting-format
        // negatives). Skip Running Balance since it's often a formula cell.
        for (const cn of [netSourceCol, resolvedBaseNetCol, 'Debit (Source)', 'Credit (Source)']) {
            const ci = idx(cn);
            if (ci >= 0 && typeof r[ci] === 'string') {
                r[ci] = parseAccountingNumber(r[ci]);
            }
        }

        // Attach the section's account to the row so downstream code can use idx('Account').
        // We append an extra column at the end for the account label and expose it via
        // a synthetic index that points to that trailing column.
        r._account = currentAccount;
        dataRows.push(r);
        (buckets[currentAccount] ||= []).push(r);
    }
    const clearingAccounts = Object.keys(buckets).filter(a => !SPECIAL_ACCOUNTS.has(a));
    if (clearingAccounts.length === 0) throw new Error('No clearing account rows found. Ensure the Xero export includes the clearing account.');

    // AR and AP are optional at the file level — a reconciliation period may
    // legitimately have no receivable or no payable activity, in which case
    // Xero omits that account's section from the export. We surface a soft
    // warning in `normalize` when the missing side is actually needed.
    const softMissing = [];
    if (!buckets['Accounts Receivable']) softMissing.push('Accounts Receivable');
    if (!buckets['Accounts Payable']) softMissing.push('Accounts Payable');

    return {
        headers, idx,
        isMultiCurrency,
        netSourceCol,
        baseNetCol: resolvedBaseNetCol,
        currencyCol,
        runningBalanceCol,
        // Kept for backwards compatibility with any external code that reads this
        baseColumn: resolvedBaseNetCol,
        clearingAccounts,
        arRows: buckets['Accounts Receivable'] || [],
        apRows: buckets['Accounts Payable'] || [],
        golRows: buckets['Realized Currency Gains'] || [],
        buckets,
        softMissing,
    };
}

function colIndexer(headers) {
    const map = new Map();
    headers.forEach((h, i) => { if (h != null) map.set(String(h).trim(), i); });
    return (name) => map.has(name) ? map.get(name) : -1;
}

function parseDate(s) {
    if (!s) return null;
    const dt = new Date(s);
    return isNaN(dt) ? null : dt;
}

// Xero exports may format negatives in accounting notation: "(28.38)" instead of "-28.38".
function parseAccountingNumber(s) {
    if (typeof s === 'number') return s;
    if (s == null || s === '') return 0;
    const trimmed = String(s).trim().replace(/,/g, '');
    const negParens = /^\((.+)\)$/.exec(trimmed);
    const cleaned = negParens ? '-' + negParens[1] : trimmed;
    const n = Number(cleaned);
    return isNaN(n) ? 0 : n;
}

// ─── Normalization algorithm ─────────────────────────────────────────────────
function normalize(parsed, clearingAccount, mode, opts = {}) {
    const warnings = [];
    const clearingRows = parsed.buckets[clearingAccount] || [];
    const { headers, idx, arRows, apRows, golRows, netSourceCol, baseNetCol, currencyCol, isMultiCurrency } = parsed;

    // For single-currency Xero orgs the file has no Currency column. The org's
    // base currency comes from Xero's account preference (via the connector in
    // production). For the MVP we accept it as an opts parameter, defaulting
    // to USD as the demo assumption.
    const singleCurrencyBase = (opts.baseCurrency || 'USD').trim();

    // Track already-assigned AR/AP and GoL rows so we don't double-count.
    const usedAr = new Set();
    const usedGol = new Set();

    const ctx = { headers, idx, arRows, apRows, golRows, netSourceCol, baseNetCol, currencyCol, isMultiCurrency, singleCurrencyBase };

    const rows = [];
    let substitutedCount = 0;
    let unresolvedFxCount = 0;

    for (const cr of clearingRows) {
        const primaryId = String(cr[idx('Reference')] ?? '').trim();
        const date = cr[idx('Date')];
        const type = cr[idx('Source')] ?? '';
        const contact = cr[idx('Contact')] ?? '';
        const description = cr[idx('Description')] ?? '';
        const netSrc = toNum(cr[idx(netSourceCol)]);
        const currency = currencyCol ? (cr[idx(currencyCol)] ?? singleCurrencyBase) : singleCurrencyBase;

        let outAmount = netSrc;
        let outCurrency = currency;

        // Bank Transfer rows pass through — no AR/AP matching, mode doesn't apply.
        const isBankTransfer = /^Bank\s*Transfer/i.test(String(type));

        if (mode === 'source' && !isBankTransfer) {
            const match = findMatchingArAp(cr, ctx, usedAr, usedGol);
            if (match) {
                // Source substitution: pull amount+currency from matched AR/AP row.
                // Sign follows the clearing side (clearing debit = +; credit = −).
                // Runs whether or not a Gain/Loss row was involved — a
                // cross-currency payment settled at the invoice's own rate
                // still needs the AR/AP row's Net (Source) + Currency, even
                // though Xero didn't book a realized gain or loss.
                const matchedNetSrc = toNum(match.row[idx(netSourceCol)]);
                const matchedCurrency = currencyCol
                    ? (match.row[idx(currencyCol)] ?? currency)
                    : singleCurrencyBase;
                const magnitude = Math.abs(matchedNetSrc);
                outAmount = Math.sign(netSrc || 0) * magnitude;
                if (outAmount === 0 && netSrc === 0) outAmount = matchedNetSrc;
                outCurrency = matchedCurrency;
                if (matchedCurrency !== currency || Math.abs(magnitude - Math.abs(netSrc)) > 0.011) {
                    substitutedCount++;
                }
            } else {
                unresolvedFxCount++;
            }
        }

        rows.push({
            'Primary ID': primaryId,
            'Secondary ID': null,
            'Amount': outAmount,
            'Currency': outCurrency,
            'Date': date,
            'Type': type,
            'Description': `${contact}/${description}`,
        });
    }

    if (mode === 'source' && parsed.golRows.length === 0) {
        warnings.push('The Realized Currency Gains account is not present in the export. In Source mode, that means no rows will be substituted — if you expect FX activity, re-export with that account ticked in the Xero account picker.');
    }
    if (mode === 'source' && parsed.softMissing && parsed.softMissing.length) {
        warnings.push(`The following accounts weren't in the export: ${parsed.softMissing.join(', ')}. This is fine if the period had no activity on those sides; otherwise re-export with them ticked in the Xero account picker so source-currency substitution can find the counterparts.`);
    }
    if (unresolvedFxCount > 0) {
        warnings.push(`${unresolvedFxCount} clearing row(s) had no matching AR/AP counterpart (by reference or by date+contact+hybrid). These rows kept their clearing amount/currency. Common causes: the counterpart row is outside the export window, or a data-quality issue.`);
    }

    return { rows, summary: { mode, substitutedCount }, warnings };
}

// Match a clearing row to (AR/AP, optional GoL) using date+contact+source-type
// enumeration and the double-entry hybrid invariant:
//   Clearing.Net(USD) + GoL.Net(Source) + AR.Net(USD) = 0
// This does NOT rely on the AR sibling-invoice row being in the export (which
// might be out of period), so it's robust to partial-period exports.
function findMatchingArAp(cr, ctx, usedAr, usedGol) {
    const { idx, arRows, apRows, golRows, baseNetCol, netSourceCol } = ctx;
    const cType = String(cr[idx('Source')] ?? '');
    const isPayable = /^Payable/i.test(cType);
    const pool = isPayable ? apRows : arRows;

    const cRef = String(cr[idx('Reference')] ?? '').trim();
    const cContact = cr[idx('Contact')];
    const cDate = cr[idx('Date')];
    const cNetBase = toNum(cr[idx(baseNetCol)]);

    // Candidate AR/AP rows on same (date, contact, source-type), unused, and
    // payment-like (skip invoice rows on the AR side).
    const arCandidates = pool.filter(r =>
        !usedAr.has(r) &&
        isPaymentLikeRow(r[idx('Source')]) &&
        sameDate(r[idx('Date')], cDate) &&
        sameContact(r[idx('Contact')], cContact) &&
        matchesType(r[idx('Source')], cType)
    );
    if (arCandidates.length === 0) return null;

    // Candidate GoL rows on same (date, contact, source-type), unused.
    const golCandidates = golRows.filter(r =>
        !usedGol.has(r) &&
        sameDate(r[idx('Date')], cDate) &&
        sameContact(r[idx('Contact')], cContact) &&
        matchesType(r[idx('Source')], cType)
    );

    // Prefer AR candidates whose Reference matches the clearing row's Reference
    // (Sales Receipt / passthrough case). Fall back to the rest.
    const refMatched = arCandidates.filter(r =>
        cRef && String(r[idx('Reference')] ?? '').trim() === cRef
    );
    const orderedGroups = [refMatched, arCandidates.filter(r => !refMatched.includes(r))];

    const targetAbs = Math.abs(cNetBase);
    let best = null; // { row, gol, diff }

    // Xero's Net convention differs by account type: Asset & Expense use
    // Debit − Credit, Liability uses Credit − Debit. That means the hybrid
    // invariant flips sign between AR-side and AP-side clearing rows:
    //   AR-side: |Clearing.Net({base})| = |AR.Net({base}) + GoL.Net(Source)|
    //   AP-side: |Clearing.Net({base})| = |AP.Net({base}) − GoL.Net(Source)|
    // In practice we compare absolute values, so the sign of the GoL term is
    // what matters. `{base}` is the org's base currency column (e.g. `Net (USD)`
    // for a USD-based org) — detected once at parse time.
    const golSign = isPayable ? -1 : +1;

    for (const group of orderedGroups) {
        for (const ar of group) {
            const arNetBase = toNum(ar[idx(baseNetCol)]);
            // Option: no GoL (same-currency or same-rate cross-currency)
            {
                const diff = Math.abs(Math.abs(arNetBase) - targetAbs);
                if (best == null || diff < best.diff) best = { row: ar, gol: null, diff };
            }
            // Option: each unused GoL candidate
            for (const gol of golCandidates) {
                const hybrid = arNetBase + golSign * toNum(gol[idx(netSourceCol)]);
                const diff = Math.abs(Math.abs(hybrid) - targetAbs);
                if (diff < best.diff) best = { row: ar, gol, diff };
            }
        }
        // If we found an exact match inside the ref-matched group, stop —
        // reference match is authoritative.
        if (best && best.diff <= 0.011 && group === refMatched && refMatched.length) break;
    }

    if (!best || best.diff > 0.011) return null;

    usedAr.add(best.row);
    if (best.gol) usedGol.add(best.gol);
    return { row: best.row, gol: best.gol, hasGoL: !!best.gol };
}

function isPaymentLikeRow(source) {
    const s = String(source || '');
    return /Payment$|Credit Note Refund$/i.test(s);
}
function sameContact(a, b) {
    return String(a ?? '').trim() === String(b ?? '').trim();
}
function matchesType(candidateSource, clearingSource) {
    // Match Receivable Payment ↔ Receivable Payment (etc.). Bank Transfers excluded upstream.
    return String(candidateSource) === String(clearingSource);
}
function sameDate(a, b) {
    if (!a || !b) return false;
    const da = a instanceof Date ? a : new Date(a);
    const db = b instanceof Date ? b : new Date(b);
    if (isNaN(da) || isNaN(db)) return false;
    return da.getFullYear() === db.getFullYear() && da.getMonth() === db.getMonth() && da.getDate() === db.getDate();
}
function keyDayContact(r, idx) {
    const d = r[idx('Date')];
    const c = r[idx('Contact')] ?? '';
    if (!d) return `~|${c}`;
    const dt = d instanceof Date ? d : new Date(d);
    return `${dt.toISOString().slice(0, 10)}|${c}`;
}
function toNum(v) {
    if (typeof v === 'number') return v;
    if (v == null || v === '') return 0;
    const n = Number(String(v).replace(/,/g, ''));
    return isNaN(n) ? 0 : n;
}
function approxEq(a, b, tol = 0.011) { return Math.abs(a - b) <= tol; }
function groupBy(arr, keyFn) {
    const m = new Map();
    for (const x of arr) {
        const k = keyFn(x);
        if (k == null) continue;
        if (!m.has(k)) m.set(k, []);
        m.get(k).push(x);
    }
    return m;
}

// ─── Test hook (Node / browser) ──────────────────────────────────────────────
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { normalize, parseWorkbook };
}
