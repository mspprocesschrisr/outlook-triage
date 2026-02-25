/* =============================================================
   taskpane.js  —  Inbox Triage Outlook Add-in  (v2 — REST API)
   =============================================================
   Uses the Outlook REST API via getCallbackTokenAsync instead
   of EWS makeEwsRequestAsync. This is more reliable across
   modern Microsoft 365 tenants and Outlook for Mac.

   No Azure AD app registration is required — the token is
   issued by Office.js from the user's existing session.
   ============================================================= */

'use strict';

// ── Default configuration (editable in the UI) ──────────────
const DEFAULTS = {
  highSenders:  'yourboss@company.com, ceo@company.com',
  lowSenders:   'noreply@, notifications@, newsletter@, donotreply@, no-reply@, alerts@, updates@, mailer@, promo@, marketing@',
  highSubjects: 'urgent, action required, decision needed, approval, critical, important',
  lowSubjects:  'unsubscribe, newsletter, your receipt, subscription, sale, offer, free, webinar, digest, weekly update, monthly report, automated, notification',
  daysBack:     7,
  maxItems:     50,
};

// ── State ────────────────────────────────────────────────────
let cachedToken   = null;
let cachedRestUrl = null;

// ── Office initialisation ────────────────────────────────────
Office.onReady(() => {
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

// ── Auth: get REST callback token ────────────────────────────
function getRestToken() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, result => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        return reject(new Error(
          'Could not get auth token: ' + (result.error.message || result.error.code) +
          '\n\nMake sure you are signed in to a Microsoft 365 or Exchange account.'
        ));
      }
      cachedToken   = result.value;
      cachedRestUrl = Office.context.mailbox.restUrl;
      resolve({ token: cachedToken, restUrl: cachedRestUrl });
    });
  });
}

// ── Main triage entry point ──────────────────────────────────
async function runTriage(dryRun) {
  setButtonsDisabled(true);
  setStatus('Authenticating…', true);
  document.getElementById('results').innerHTML = '';

  try {
    const { token, restUrl } = await getRestToken();

    setStatus('Fetching unread messages…', true);
    const cfg   = getConfig();
    const items = await fetchUnreadMessages(token, restUrl, cfg.daysBack);

    if (!items.length) {
      renderEmpty('No unread messages found in the last ' + cfg.daysBack + ' day(s).');
      setStatus('Inbox looks clear!');
      return;
    }

    setStatus(`Scoring ${items.length} message(s)…`, true);

    const scored = items.map(item => ({
      ...item,
      score:         scoreItem(item, cfg),
      isLowPriority: isLowPriority(item, cfg),
    }));

    const priorityItems = scored
      .filter(i => !i.isLowPriority)
      .sort((a, b) => b.score - a.score);

    const lowItems = scored.filter(i => i.isLowPriority);

    if (!dryRun && lowItems.length) {
      setStatus(`Marking ${lowItems.length} low-priority message(s) as read…`, true);
      await markItemsAsRead(token, restUrl, lowItems.map(i => i.id));
    }

    renderResults(priorityItems, lowItems, dryRun);

    const modeTag = dryRun ? ' <span class="tag-dryrun">dry run</span>' : '';
    const verb    = dryRun ? 'Would mark' : 'Marked';
    setStatus(`${verb} ${lowItems.length} as read · ${priorityItems.length} need replies${modeTag}`);

  } catch (err) {
    const msg = err.message || String(err);
    setStatus('⚠️ ' + msg);
    renderEmpty('Something went wrong.<br><small style="color:var(--muted)">' + esc(msg) + '</small>');
    console.error(err);
  } finally {
    setButtonsDisabled(false);
  }
}

// ── Mark-only action ─────────────────────────────────────────
async function markLowPriorityRead() {
  setButtonsDisabled(true);
  setStatus('Authenticating…', true);

  try {
    const { token, restUrl } = await getRestToken();
    const cfg   = getConfig();

    setStatus('Fetching unread messages…', true);
    const items = await fetchUnreadMessages(token, restUrl, cfg.daysBack);
    const low   = items.filter(i => isLowPriority(i, cfg));

    if (!low.length) {
      setStatus('No low-priority unread messages found.');
      return;
    }

    setStatus(`Marking ${low.length} message(s) as read…`, true);
    await markItemsAsRead(token, restUrl, low.map(i => i.id));
    setStatus(`✓ Marked ${low.length} low-priority message(s) as read.`);

  } catch (err) {
    setStatus('⚠️ ' + (err.message || err));
  } finally {
    setButtonsDisabled(false);
  }
}

// ── Outlook REST API: Fetch unread messages ──────────────────
async function fetchUnreadMessages(token, restUrl, daysBack) {
  const since = new Date(Date.now() - daysBack * 86_400_000).toISOString();

  const params = new URLSearchParams({
    '$filter':  `isRead eq false and receivedDateTime ge ${since}`,
    '$orderby': 'receivedDateTime desc',
    '$top':     String(DEFAULTS.maxItems),
    '$select':  'id,subject,from,toRecipients,receivedDateTime,importance,bodyPreview,isRead',
  });

  const url = `${restUrl}/v2.0/me/mailfolders/inbox/messages?${params}`;

  const response = await fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token,
      'Accept':        'application/json',
    },
  });

  if (!response.ok) {
    const body = await response.text().catch(() => '');
    throw new Error(
      `REST API error ${response.status}: ${response.statusText}` +
      (body ? '\n' + body.slice(0, 300) : '')
    );
  }

  const data = await response.json();
  const myEmail = (Office.context.mailbox.userProfile.emailAddress || '').toLowerCase();

  return (data.value || []).map(msg => {
    const fromAddr = msg.from?.emailAddress?.address || '';
    const fromName = msg.from?.emailAddress?.name    || '';
    const from     = fromName && fromName !== fromAddr
      ? `${fromName} <${fromAddr}>`
      : (fromAddr || fromName || '(unknown)');

    const isDirect = (msg.toRecipients || []).some(
      r => (r.emailAddress?.address || '').toLowerCase() === myEmail
    );

    return {
      id:          msg.id,
      subject:     msg.subject     || '(no subject)',
      from,
      fromAddr:    fromAddr.toLowerCase(),
      received:    msg.receivedDateTime,
      importance:  msg.importance  || 'Normal',
      bodyPreview: msg.bodyPreview || '',
      isDirect,
    };
  });
}

// ── Outlook REST API: Mark messages as read ──────────────────
async function markItemsAsRead(token, restUrl, ids) {
  const requests = ids.map(id =>
    fetch(`${restUrl}/v2.0/me/messages/${id}`, {
      method:  'PATCH',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type':  'application/json',
      },
      body: JSON.stringify({ isRead: true }),
    }).then(r => {
      if (!r.ok) console.warn(`Failed to mark ${id} as read: ${r.status}`);
    })
  );
  await Promise.all(requests);
}

// ── Scoring engine ───────────────────────────────────────────
function scoreItem(item, cfg) {
  if (isLowPriority(item, cfg)) return 0;

  let score = 10;

  const sender = item.fromAddr;
  const subj   = item.subject.toLowerCase();
  const body   = item.bodyPreview.toLowerCase();

  if (cfg.highSenders.some(s => sender.includes(s)))   score += 50;
  if (cfg.highSubjects.some(k => subj.includes(k)))    score += 30;

  const hoursOld = (Date.now() - new Date(item.received).getTime()) / 3_600_000;
  if      (hoursOld < 4)  score += 20;
  else if (hoursOld < 24) score += 10;
  else if (hoursOld < 48) score += 5;

  if (item.isDirect) score += 15;

  const replyPhrases = ['please reply','let me know','your thoughts','waiting for','your feedback','can you'];
  if (replyPhrases.some(p => body.includes(p))) score += 10;

  if (item.importance === 'High') score += 25;
  if (item.importance === 'Low')  score -= 10;

  return Math.max(score, 1);
}

function isLowPriority(item, cfg) {
  const sender = item.fromAddr;
  const subj   = item.subject.toLowerCase();

  if (cfg.lowSenders.some(s => sender.includes(s)))  return true;
  if (cfg.lowSubjects.some(k => subj.includes(k)))   return true;
  if (item.importance === 'Low' && !item.isDirect)    return true;
  return false;
}

function badgeLevel(score) {
  if (score >= 80) return 'high';
  if (score >= 40) return 'med';
  return 'low';
}

// ── Render results ───────────────────────────────────────────
function renderResults(priorityItems, lowItems, dryRun) {
  const container = document.getElementById('results');
  let html = '';

  if (priorityItems.length === 0) {
    html += `<div class="empty-state"><div class="empty-icon">✅</div>No emails needing replies.</div>`;
  } else {
    html += `<div class="section-header">Suggested reply order — ${priorityItems.length} email(s)</div>`;
    priorityItems.forEach((item, i) => {
      const level = badgeLevel(item.score);
      html += `
        <div class="email-card ${level}">
          <div class="card-top">
            <span class="badge ${level}">${level}</span>
            <span class="score">${item.score}pts</span>
          </div>
          <div class="card-subject" title="${esc(item.subject)}">${i+1}. ${esc(item.subject)}</div>
          <div class="card-meta">
            <span class="card-from" title="${esc(item.from)}">${esc(item.from)}</span>
            <span style="margin-left:auto;flex-shrink:0">${formatAge(item.received)}</span>
          </div>
        </div>`;
    });
  }

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

// ── Utilities ────────────────────────────────────────────────
function formatAge(isoString) {
  if (!isoString) return '';
  const diff  = Date.now() - new Date(isoString).getTime();
  const hours = diff / 3_600_000;
  if (hours < 1)  return Math.round(hours * 60) + 'm ago';
  if (hours < 24) return Math.round(hours) + 'h ago';
  return Math.round(hours / 24) + 'd ago';
}

function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
