// Xero Assisted Reconciliation — Normalizer
// Reads a Xero "Account Transactions" export (single file, 4 accounts) and produces
// a normalized reconciliation-ready sheet. See ../fdd.md for algorithm rationale.

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
}

function runNormalize() {
    const mode = document.querySelector('input[name="mode"]:checked').value;
    try {
        state.result = normalize(state.parsed, state.selectedClearing, mode);
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
    // Prefer a sheet with "4 in 1" in its name; otherwise the first sheet with an
    // 'Account Transactions' shape.
    const sheetName = wb.SheetNames.find(n => /4[\s_-]?in[\s_-]?1/i.test(n))
        || wb.SheetNames.find(n => /account\s+transactions/i.test(n))
        || wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error('No sheet found.');
    // raw: true keeps numbers as numbers and dates as Date objects (paired with cellDates
    // on XLSX.read). Formatted-string mode breaks on accounting-format negatives like "(28.38)".
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });

    // Find header row: the first row containing 'Account' and 'Date'
    const headerIdx = raw.findIndex(r => Array.isArray(r) && r.includes('Account') && r.includes('Date'));
    if (headerIdx === -1) throw new Error('Header row not found. Expected columns like Date, Account, Net, Reference.');
    const headers = raw[headerIdx];

    // Data rows: everything after header, until the 'Total' row or end
    const dataRows = [];
    for (let i = headerIdx + 1; i < raw.length; i++) {
        const r = raw[i];
        if (!Array.isArray(r) || r.every(c => c == null || c === '')) continue;
        // Total row detection: first non-empty cell equals 'Total'
        const first = r.find(c => c != null && c !== '');
        if (first === 'Total') continue;
        dataRows.push(r);
    }

    // With raw:true, cells are already native types. Fall back to coercion only when
    // sheet_to_json returns a string (e.g. formula-derived cells).
    const idx = colIndexer(headers);
    for (const r of dataRows) {
        for (const cn of ['Net (Source)', 'Net (USD)', 'Debit (Source)', 'Credit (Source)']) {
            const ci = idx(cn);
            if (ci >= 0 && typeof r[ci] === 'string') {
                r[ci] = parseAccountingNumber(r[ci]);
            }
        }
    }

    // Split by Account
    const buckets = {};
    for (const r of dataRows) {
        const a = r[idx('Account')];
        if (a == null || a === '') continue;
        (buckets[a] ||= []).push(r);
    }
    const clearingAccounts = Object.keys(buckets).filter(a => !SPECIAL_ACCOUNTS.has(a));
    if (clearingAccounts.length === 0) throw new Error('No clearing account rows found. Ensure the Xero export includes the clearing account.');
    const missing = [];
    if (!buckets['Accounts Receivable']) missing.push('Accounts Receivable');
    if (!buckets['Accounts Payable']) missing.push('Accounts Payable');
    // Realized Currency Gains is allowed to be empty; only warn if entirely missing
    // (Xero may omit the account name from the export when there are no rows, so
    // this isn't a hard failure — we just note it as a warning during normalize).

    return {
        headers, idx,
        clearingAccounts,
        arRows: buckets['Accounts Receivable'] || [],
        apRows: buckets['Accounts Payable'] || [],
        golRows: buckets['Realized Currency Gains'] || [],
        buckets,
        missingCritical: missing,
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
function normalize(parsed, clearingAccount, mode) {
    const warnings = [];
    if (parsed.missingCritical.length) {
        return { rows: [], summary: { mode, substitutedCount: 0 }, warnings: [
            `Required account missing from the file: ${parsed.missingCritical.join(', ')}. Re-export from Xero with all four accounts (clearing + Accounts Receivable + Accounts Payable + Realized Currency Gains) selected.`,
        ] };
    }

    const clearingRows = parsed.buckets[clearingAccount] || [];
    const { headers, idx, arRows, apRows, golRows } = parsed;

    // Pre-index AR/AP by Reference and by (date + contact)
    const arByRef = groupBy(arRows, r => r[idx('Reference')]);
    const apByRef = groupBy(apRows, r => r[idx('Reference')]);
    const arByDayContact = groupBy(arRows, r => keyDayContact(r, idx));
    const apByDayContact = groupBy(apRows, r => keyDayContact(r, idx));
    // Gain/Loss rows indexed by Reference (Xero writes Invoice Number here)
    const golByRef = groupBy(golRows, r => r[idx('Reference')]);

    const rows = [];
    let substitutedCount = 0;

    for (const cr of clearingRows) {
        const primaryId = String(cr[idx('Reference')] ?? '').trim();
        const date = cr[idx('Date')];
        const type = cr[idx('Source')] ?? '';
        const contact = cr[idx('Contact')] ?? '';
        const description = cr[idx('Description')] ?? '';
        const netSrc = toNum(cr[idx('Net (Source)')]);
        const netUsd = toNum(cr[idx('Net (USD)')]);
        const currency = cr[idx('Currency')] ?? '';

        let outAmount = netSrc;
        let outCurrency = currency;

        // Bank Transfer rows pass through — no AR/AP matching, mode doesn't apply.
        const isPayable = /^Payable/i.test(String(type));
        const isBankTransfer = /^Bank\s*Transfer/i.test(String(type));

        if (mode === 'source' && !isBankTransfer) {
            const match = findMatchingArAp(cr, {
                headers, idx, arRows, apRows, golRows,
                arByRef, apByRef, arByDayContact, apByDayContact,
                golByRef,
                isPayable,
            });
            if (match && match.hasGoL) {
                // Source substitution: pull amount+currency from matched AR/AP row.
                // Sign follows the clearing side (clearing debit = +; credit = −).
                const matchedNetSrc = toNum(match.row[idx('Net (Source)')]);
                const magnitude = Math.abs(matchedNetSrc);
                outAmount = Math.sign(netSrc || 0) * magnitude;
                if (outAmount === 0 && netSrc === 0) outAmount = matchedNetSrc; // extreme edge case
                outCurrency = match.row[idx('Currency')] ?? currency;
                substitutedCount++;
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

    // Post hoc warning: Gain/Loss bucket entirely absent from the export.
    if (mode === 'source' && parsed.golRows.length === 0) {
        warnings.push('The Realized Currency Gains account is not present in the export. In Source mode, that means no rows will be substituted — if you expect FX activity, re-export with that account ticked in the Xero account picker.');
    }

    return { rows, summary: { mode, substitutedCount }, warnings };
}

function findMatchingArAp(cr, ctx) {
    const { idx, isPayable } = ctx;
    const pool = isPayable ? ctx.apRows : ctx.arRows;
    const byRef = isPayable ? ctx.apByRef : ctx.arByRef;
    const byDayContact = isPayable ? ctx.apByDayContact : ctx.arByDayContact;

    const cRef = String(cr[idx('Reference')] ?? '').trim();
    const cContact = cr[idx('Contact')];
    const cDate = cr[idx('Date')];
    const cNetUsd = toNum(cr[idx('Net (USD)')]);
    const cType = String(cr[idx('Source')] ?? '');

    // 1. Reference join. Restrict to a payment/refund row (skip invoice rows).
    if (cRef) {
        const candidates = (byRef.get(cRef) || []).filter(r => isPaymentLikeRow(r[idx('Source')]));
        // Prefer a same-date row; else any candidate.
        const best = candidates.find(r => sameDate(r[idx('Date')], cDate)) || candidates[0];
        if (best) {
            return { row: best, hasGoL: hasLinkedGoL(best, ctx) };
        }
    }

    // 2. Fallback: date + contact + hybrid amount match.
    const dayKey = keyDayContact(cr, idx);
    const candidates = (byDayContact.get(dayKey) || []).filter(r => matchesType(r[idx('Source')], cType));
    if (candidates.length === 0) return null;

    // Compute hybrid_amount for each candidate:
    //   hybrid = candidate.Net(USD) + linkedGoL.Net(Source)
    // where linkedGoL is the GoL row that shares the invoice number of the
    // candidate's sibling invoice row, at the sort position of the payment
    // among that invoice's payments.
    const scored = candidates.map(cand => {
        const gol = findLinkedGoL(cand, ctx, candidates);
        const candNetUsd = toNum(cand[idx('Net (USD)')]);
        const golAmt = gol ? toNum(gol[idx('Net (Source)')]) : 0;
        const hybrid = candNetUsd + golAmt;
        return { row: cand, gol, hybrid };
    });

    const targetAbs = Math.abs(cNetUsd);
    // Best match: absolute value of hybrid == absolute value of clearing.Net(USD).
    const best = scored.find(s => approxEq(Math.abs(s.hybrid), targetAbs));
    if (best) return { row: best.row, hasGoL: !!best.gol };
    // No exact hybrid match — take the closest candidate but don't claim GoL.
    scored.sort((a, b) => Math.abs(Math.abs(a.hybrid) - targetAbs) - Math.abs(Math.abs(b.hybrid) - targetAbs));
    return scored[0] ? { row: scored[0].row, hasGoL: !!scored[0].gol } : null;
}

// Returns the linked GoL row for a payment row, if one exists.
function findLinkedGoL(paymentRow, ctx, sameDayContactCandidates) {
    const { idx, arRows, apRows, golByRef, isPayable } = ctx;
    const pool = isPayable ? apRows : arRows;
    const pRef = paymentRow[idx('Reference')];
    if (!pRef) return null;

    // Sibling invoice row (same Reference, but Source is *Invoice)
    const invRow = pool.find(r =>
        r[idx('Reference')] === pRef && isInvoiceRow(r[idx('Source')])
    );
    if (!invRow) return null;
    const invNum = invRow[idx('Invoice Number')];
    if (!invNum) return null;

    const golRows = golByRef.get(invNum) || [];
    if (golRows.length === 0) return null;

    // Position of this payment among all payments for the same invoice within
    // the day-contact candidate pool (position is preserved in the AR ordering).
    const paymentsForInvoice = pool.filter(r =>
        isPaymentLikeRow(r[idx('Source')]) &&
        r[idx('Reference')] === pRef  // shares payment reference == same sibling invoice
    );
    const pos = paymentsForInvoice.indexOf(paymentRow);
    if (pos < 0) return null;
    return golRows[pos] || null;
}

function hasLinkedGoL(paymentRow, ctx) {
    return !!findLinkedGoL(paymentRow, ctx);
}

function isPaymentLikeRow(source) {
    const s = String(source || '');
    return /Payment$|Credit Note Refund$/i.test(s);
}
function isInvoiceRow(source) {
    const s = String(source || '');
    return /Invoice$|Credit Note$/i.test(s) && !/Refund/i.test(s);
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
