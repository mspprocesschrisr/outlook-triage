/* =============================================================
   taskpane.js  —  Inbox Triage Outlook Add-in
   =============================================================
   Communicates with Exchange via EWS (makeEwsRequestAsync).
   No separate Azure AD registration required — uses the
   user's existing authenticated Outlook session.
   ============================================================= */

'use strict';

// ── Default configuration (editable in the UI) ──────────────
const DEFAULTS = {
  highSenders:  'yourboss@company.com, ceo@company.com',
  lowSenders:   'noreply@, notifications@, newsletter@, donotreply@, no-reply@, alerts@, updates@, mailer@, promo@, marketing@',
  highSubjects: 'urgent, action required, decision needed, approval, critical, important',
  lowSubjects:  'unsubscribe, newsletter, your receipt, subscription, sale, offer, free, webinar, digest, weekly update, monthly report, automated, notification',
  daysBack:     7,
  maxItems:     50,   // max fetched from EWS
};

// ── State ────────────────────────────────────────────────────
let lastScoredItems = [];   // [{id, subject, from, received, score, isLowPriority}]

// ── Office initialisation ────────────────────────────────────
Office.onReady(() => {
  // Pre-fill config fields with defaults
  document.getElementById('cfgHighSenders').value = DEFAULTS.highSenders;
  document.getElementById('cfgLowSenders').value  = DEFAULTS.lowSenders;
});

// ── Read config from UI ──────────────────────────────────────
function getConfig() {
  return {
    highSenders:  splitList(document.getElementById('cfgHighSenders').value),
    lowSenders:   splitList(document.getElementById('cfgLowSenders').value),
    highSubjects: splitList(DEFAULTS.highSubjects),
    lowSubjects:  splitList(DEFAULTS.lowSubjects),
    daysBack:     parseInt(document.getElementById('cfgDays').value, 10) || 7,
  };
}

function splitList(str) {
  return str.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
}

// ── UI helpers ───────────────────────────────────────────────
function setStatus(msg, loading = false) {
  const el = document.getElementById('status');
  el.innerHTML = loading
    ? `<span class="spinner"></span>${msg}`
    : msg;
}

function setButtonsDisabled(disabled) {
  ['btnDryRun','btnTriage','btnMarkRead'].forEach(id => {
    document.getElementById(id).disabled = disabled;
  });
}

function renderEmpty(msg) {
  document.getElementById('results').innerHTML =
    `<div class="empty-state"><div class="empty-icon">📭</div>${msg}</div>`;
}

// ── Main triage entry point ──────────────────────────────────
async function runTriage(dryRun) {
  setButtonsDisabled(true);
  setStatus('Fetching unread messages…', true);
  document.getElementById('results').innerHTML = '';

  try {
    const cfg   = getConfig();
    const items = await fetchUnreadMessages(cfg.daysBack);

    if (!items.length) {
      renderEmpty('No unread messages found in the last ' + cfg.daysBack + ' day(s).');
      setStatus('Inbox looks clear!');
      return;
    }

    setStatus(`Scoring ${items.length} message(s)…`, true);

    // Score every item
    lastScoredItems = items.map(item => ({
      ...item,
      score:         scoreItem(item, cfg),
      isLowPriority: isLowPriority(item, cfg),
    }));

    const priorityItems = lastScoredItems
      .filter(i => !i.isLowPriority)
      .sort((a, b) => b.score - a.score);

    const lowItems = lastScoredItems.filter(i => i.isLowPriority);

    // In live mode, mark low-priority items as read
    if (!dryRun && lowItems.length) {
      setStatus(`Marking ${lowItems.length} low-priority message(s) as read…`, true);
      await markItemsAsRead(lowItems.map(i => i.id));
    }

    renderResults(priorityItems, lowItems, dryRun);

    const modeTag  = dryRun ? ' <span class="tag-dryrun">dry run</span>' : '';
    const marked   = dryRun ? `Would mark ${lowItems.length}` : `Marked ${lowItems.length}`;
    setStatus(`${marked} as read · ${priorityItems.length} need replies${modeTag}`);

  } catch (err) {
    setStatus('Error: ' + (err.message || err));
    renderEmpty('Something went wrong — see status bar above.');
    console.error(err);
  } finally {
    setButtonsDisabled(false);
  }
}

// ── Mark-only action (no report) ────────────────────────────
async function markLowPriorityRead() {
  setButtonsDisabled(true);
  setStatus('Fetching unread messages…', true);

  try {
    const cfg   = getConfig();
    const items = await fetchUnreadMessages(cfg.daysBack);
    const low   = items.filter(i => isLowPriority(i, cfg));

    if (!low.length) {
      setStatus('No low-priority unread messages found.');
      return;
    }

    setStatus(`Marking ${low.length} message(s) as read…`, true);
    await markItemsAsRead(low.map(i => i.id));
    setStatus(`✓ Marked ${low.length} low-priority message(s) as read.`);

  } catch (err) {
    setStatus('Error: ' + (err.message || err));
  } finally {
    setButtonsDisabled(false);
  }
}

// ── Render results ───────────────────────────────────────────
function renderResults(priorityItems, lowItems, dryRun) {
  const container = document.getElementById('results');
  let html = '';

  // Priority reply list
  if (priorityItems.length === 0) {
    html += `<div class="empty-state"><div class="empty-icon">✅</div>No emails needing replies.</div>`;
  } else {
    html += `<div class="section-header">Suggested reply order — ${priorityItems.length} email(s)</div>`;
    priorityItems.forEach((item, i) => {
      const level     = badgeLevel(item.score);
      const timeStr   = formatAge(item.received);
      html += `
        <div class="email-card ${level}">
          <div class="card-top">
            <span class="badge ${level}">${level}</span>
            <span class="score">${item.score}pts</span>
          </div>
          <div class="card-subject" title="${esc(item.subject)}">${i+1}. ${esc(item.subject)}</div>
          <div class="card-meta">
            <span class="card-from" title="${esc(item.from)}">${esc(item.from)}</span>
            <span style="margin-left:auto;flex-shrink:0">${timeStr}</span>
          </div>
        </div>`;
    });
  }

  // Low-priority section
  if (lowItems.length) {
    html += `<div class="divider"></div>`;
    const verb = dryRun ? 'Would be marked' : 'Marked';
    html += `<div class="section-header">${verb} as read — ${lowItems.length} email(s)</div>`;
    lowItems.slice(0, 20).forEach(item => {
      html += `
        <div class="email-card" style="opacity:0.55">
          <div class="card-subject" title="${esc(item.subject)}">${esc(item.subject)}</div>
          <div class="card-meta">
            <span class="card-from">${esc(item.from)}</span>
            <span style="margin-left:auto">${formatAge(item.received)}</span>
          </div>
        </div>`;
    });
    if (lowItems.length > 20) {
      html += `<div style="padding:6px 16px;font-size:10px;color:var(--muted)">…and ${lowItems.length - 20} more</div>`;
    }
  }

  container.innerHTML = html;
}

// ── Scoring engine ───────────────────────────────────────────
function scoreItem(item, cfg) {
  if (isLowPriority(item, cfg)) return 0;

  let score = 10;   // baseline

  const sender  = item.from.toLowerCase();
  const subj    = item.subject.toLowerCase();
  const body    = (item.bodyPreview || '').toLowerCase();

  // VIP sender
  if (cfg.highSenders.some(s => sender.includes(s))) score += 50;

  // High-priority subject keywords
  if (cfg.highSubjects.some(k => subj.includes(k))) score += 30;

  // Recency
  const hoursOld = (Date.now() - new Date(item.received).getTime()) / 3_600_000;
  if      (hoursOld < 4)  score += 20;
  else if (hoursOld < 24) score += 10;
  else if (hoursOld < 48) score += 5;

  // Direct recipient (in To:, not just Cc)
  if (item.isDirect) score += 15;

  // Reply-requesting language in body preview
  const replyPhrases = ['please reply','let me know','your thoughts','waiting for','your feedback','can you'];
  if (replyPhrases.some(p => body.includes(p))) score += 10;

  // Outlook importance flag
  if (item.importance === 'High')  score += 25;
  if (item.importance === 'Low')   score -= 10;

  return Math.max(score, 1);
}

function isLowPriority(item, cfg) {
  const sender = item.from.toLowerCase();
  const subj   = item.subject.toLowerCase();

  if (cfg.lowSenders.some(s => sender.includes(s))) return true;
  if (cfg.lowSubjects.some(k => subj.includes(k)))  return true;
  if (item.importance === 'Low' && !item.isDirect)   return true;
  return false;
}

function badgeLevel(score) {
  if (score >= 80) return 'high';
  if (score >= 40) return 'med';
  return 'low';
}

// ── EWS: Fetch unread messages ───────────────────────────────
function fetchUnreadMessages(daysBack) {
  return new Promise((resolve, reject) => {

    const since = new Date(Date.now() - daysBack * 86_400_000).toISOString();

    // FindItem SOAP — retrieves up to DEFAULTS.maxItems unread messages
    const soap = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Body>
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>AllProperties</t:BaseShape>
        <t:BodyType>Text</t:BodyType>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="${DEFAULTS.maxItems}"
        Offset="0" BasePoint="Beginning"/>
      <m:Restriction>
        <t:And>
          <t:IsEqualTo>
            <t:FieldURI FieldURI="message:IsRead"/>
            <t:FieldURIOrConstant>
              <t:Constant Value="false"/>
            </t:FieldURIOrConstant>
          </t:IsEqualTo>
          <t:IsGreaterThan>
            <t:FieldURI FieldURI="item:DateTimeReceived"/>
            <t:FieldURIOrConstant>
              <t:Constant Value="${since}"/>
            </t:FieldURIOrConstant>
          </t:IsGreaterThan>
        </t:And>
      </m:Restriction>
      <m:SortOrder>
        <t:FieldOrder Order="Descending">
          <t:FieldURI FieldURI="item:DateTimeReceived"/>
        </t:FieldOrder>
      </m:SortOrder>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="inbox"/>
      </m:ParentFolderIds>
    </m:FindItem>
  </soap:Body>
</soap:Envelope>`;

    Office.context.mailbox.makeEwsRequestAsync(soap, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        return reject(new Error(result.error.message));
      }
      try {
        resolve(parseEwsFindItemResponse(result.value));
      } catch (e) {
        reject(e);
      }
    });
  });
}

// ── EWS: Mark items as read ──────────────────────────────────
function markItemsAsRead(itemIds) {
  if (!itemIds.length) return Promise.resolve();

  // Build UpdateItem for each id
  const itemChanges = itemIds.map(id => `
    <t:ItemChange>
      <t:ItemId Id="${id}"/>
      <t:Updates>
        <t:SetItemField>
          <t:FieldURI FieldURI="message:IsRead"/>
          <t:Message>
            <t:IsRead>true</t:IsRead>
          </t:Message>
        </t:SetItemField>
      </t:Updates>
    </t:ItemChange>`).join('\n');

  const soap = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Body>
    <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AutoResolve">
      <m:ItemChanges>
        ${itemChanges}
      </m:ItemChanges>
    </m:UpdateItem>
  </soap:Body>
</soap:Envelope>`;

  return new Promise((resolve, reject) => {
    Office.context.mailbox.makeEwsRequestAsync(soap, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        return reject(new Error(result.error.message));
      }
      resolve();
    });
  });
}

// ── Parse EWS FindItem XML response ─────────────────────────
function parseEwsFindItemResponse(xmlStr) {
  const parser = new DOMParser();
  const doc    = parser.parseFromString(xmlStr, 'text/xml');

  const messages = doc.getElementsByTagNameNS(
    'http://schemas.microsoft.com/exchange/services/2006/types', 'Message');

  const items = [];

  for (const msg of messages) {
    const get = (ns, tag) => {
      const el = msg.getElementsByTagNameNS(ns, tag)[0];
      return el ? el.textContent.trim() : '';
    };
    const T = 'http://schemas.microsoft.com/exchange/services/2006/types';

    // Item ID (needed for mark-as-read)
    const itemIdEl = msg.getElementsByTagNameNS(T, 'ItemId')[0];
    const id       = itemIdEl ? itemIdEl.getAttribute('Id') : '';

    // From display name + address
    const fromEl      = msg.getElementsByTagNameNS(T, 'From')[0];
    const fromName    = fromEl ? (fromEl.getElementsByTagNameNS(T, 'Name')[0] || {}).textContent || '' : '';
    const fromAddr    = fromEl ? (fromEl.getElementsByTagNameNS(T, 'EmailAddress')[0] || {}).textContent || '' : '';
    const from        = fromName && fromName !== fromAddr ? `${fromName} <${fromAddr}>` : (fromAddr || fromName);

    // Is the current user in To: (not just Cc)?
    const toRec    = msg.getElementsByTagNameNS(T, 'ToRecipients')[0];
    const myEmail  = (Office.context.mailbox.userProfile.emailAddress || '').toLowerCase();
    let   isDirect = false;
    if (toRec) {
      for (const mb of toRec.getElementsByTagNameNS(T, 'Mailbox')) {
        const addr = (mb.getElementsByTagNameNS(T, 'EmailAddress')[0] || {}).textContent || '';
        if (addr.toLowerCase() === myEmail) { isDirect = true; break; }
      }
    }

    items.push({
      id,
      subject:     get(T, 'Subject')          || '(no subject)',
      from:        from                        || '(unknown)',
      received:    get(T, 'DateTimeReceived'),
      importance:  get(T, 'Importance'),
      bodyPreview: get(T, 'Body'),
      isDirect,
    });
  }

  return items;
}

// ── Utilities ────────────────────────────────────────────────
function formatAge(isoString) {
  if (!isoString) return '';
  const diff  = Date.now() - new Date(isoString).getTime();
  const hours = diff / 3_600_000;
  if (hours < 1)   return Math.round(hours * 60) + 'm ago';
  if (hours < 24)  return Math.round(hours) + 'h ago';
  const days = Math.round(hours / 24);
  return days + 'd ago';
}

function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
