/**
 * The Gray House – landlord onboarding (standalone).
 * Paste this whole file into a new Apps Script project (script.google.com), then Deploy > New deployment > Web app,
 * Execute as: Me, Who has access: Anyone. Copy the /exec web app URL into the page (see GO-LIVE.md).
 *
 * What it does:
 *  - submitPack:  emails the signed landlord pack (PDF) plus the landlord's uploads to Michelle, and a copy of the PDF to the landlord.
 *  - getSlots:    free valuation times from Michelle's Google Calendar.
 *  - book:        puts the visit in Michelle's calendar and emails both sides. Michelle's email has the "Enter the agreed rent" button.
 *  - getVisit / issueLetter / signLetter: Michelle enters the rent, the landlord signs the Rental Value Confirmation, both get the PDF.
 * Nothing is stored except the calendar event (which holds the visit details) and the emails sent.
 */

/* ======== SETTINGS: change these on go-live day ======== */
const CONFIG = {
  MICHELLE_EMAIL: 'charlotte@hammer-and-pen.com',   // where onboarding packs, bookings and signed letters go. Change to Michelle's on go-live.
  MICHELLE_NAME: 'Michelle Barber',
  CALENDAR_ID: 'primary',                           // 'primary' = this Google account's own calendar, or Michelle's calendar address once shared
  SITE_URL: 'https://charlotte644.github.io/gray-house-onboarding/', // the landlord page
  TERMS_PDF: 'terms-of-business-v9.pdf',             // the full Terms of Business, next to the page; attached to the landlord's copy
  FROM_NAME: 'The Gray House',
  OFFICE_PHONE: '02920 692222',
  SLOT_TIMES: ['09:30', '11:00', '13:30', '15:00', '16:30'],
  VISIT_MINUTES: 45,
  DAYS_AHEAD: 21,
  MIN_NOTICE_HOURS: 24,
  TIMEZONE: 'Europe/London'
};
/* ======================================================= */

const API = { getSlots, submitPack, book, getVisit, issueLetter, signLetter };

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const fn = API[body.fn];
    if (!fn) throw new Error('Unknown request.');
    return out_({ ok: true, data: fn.apply(null, body.args || []) });
  } catch (err) {
    console.error(err);
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}
function doGet() { return out_({ ok: true, data: 'The Gray House onboarding service is running.' }); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- run this once from the editor to grant permissions and check the set-up ---------- */
function setUpAndTest() {
  secret_();
  const c = cal_();
  Logger.log('Calendar: ' + c.getName());
  Logger.log('Free days found: ' + Object.keys(getSlots()).length);
  MailApp.sendEmail({ to: CONFIG.MICHELLE_EMAIL, subject: 'The Gray House onboarding: set-up test', htmlBody: email_('Set-up test', '<p>The onboarding service can send email and read the calendar. You can delete this email.</p>'), name: CONFIG.FROM_NAME });
  Logger.log('Test email sent to ' + CONFIG.MICHELLE_EMAIL + '. Remaining email quota today: ' + MailApp.getRemainingDailyQuota());
}

/* ---------- helpers ---------- */
function secret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function cal_() {
  const c = CONFIG.CALENDAR_ID === 'primary' ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!c) throw new Error('The calendar could not be opened. Check CALENDAR_ID and that it is shared with this account.');
  return c;
}
function b64_(s) { return Utilities.base64EncodeWebSafe(s).replace(/=+$/, ''); }
function unb64_(s) { while (s.length % 4) s += '='; return Utilities.newBlob(Utilities.base64DecodeWebSafe(s)).getDataAsString(); }
function sig_(kind, id) { return b64_(Utilities.computeHmacSha256Signature(kind + ':' + id, secret_())).slice(0, 24); }
function token_(kind, id) { return kind + '.' + b64_(id) + '.' + sig_(kind, id); }
function readToken_(t, kind) {
  const p = String(t || '').split('.');
  if (p.length !== 3 || p[0] !== kind) throw new Error('This link is not valid.');
  const id = unb64_(p[1]);
  if (sig_(kind, id) !== p[2]) throw new Error('This link is not valid.');
  return id;
}
function event_(id) {
  const ev = cal_().getEventById(id);
  if (!ev) throw new Error('This visit could not be found. It may have been cancelled. Please call ' + CONFIG.OFFICE_PHONE + '.');
  return ev;
}
const TAGS = ['first', 'last', 'email', 'phone', 'addr1', 'town', 'pc', 'second', 's_first', 's_last', 'tier', 'tierName', 'furnish', 'signedAt', 'ref', 'note', 'rent', 'from', 'pets', 'notes', 'issuedAt', 'lsig', 'lSignedAt'];
function tags_(ev) { const o = {}; TAGS.forEach(k => o[k] = ev.getTag(k) || ''); return o; }
function setTag_(ev, k, v) { ev.setTag(k, String(v == null ? '' : v).slice(0, 900)); }
function esc_(v) { return String(v == null ? '' : v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function clean_(v, n) { return String(v == null ? '' : v).trim().slice(0, n || 200); }
function okEmail_(e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); }
function fmt_(d, f) { return Utilities.formatDate(d, CONFIG.TIMEZONE, f); }
function slotDate_(day, time) { return Utilities.parseDate(day + ' ' + time, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm'); }
function when_(ev) { return fmt_(ev.getStartTime(), "EEEE d MMMM 'at' HH:mm"); }
function throttle_(key, max) {
  const c = CacheService.getScriptCache(), k = 'n:' + key, n = Number(c.get(k) || 0);
  if (n >= max) throw new Error('Too many requests. Please try again in an hour, or call ' + CONFIG.OFFICE_PHONE + '.');
  c.put(k, String(n + 1), 3600);
}
function send_(o) {
  MailApp.sendEmail(Object.assign({ name: CONFIG.FROM_NAME, replyTo: CONFIG.MICHELLE_EMAIL }, o));
}
function email_(title, body, button) {
  return '<div style="background:#f5f3f0;padding:24px 12px;font-family:Arial,Helvetica,sans-serif;color:#2e2f36">' +
    '<div style="max-width:600px;margin:0 auto;background:#ffffff;border:1px solid #e4dfd9;border-radius:10px;overflow:hidden">' +
    '<div style="background:#4b4c57;color:#f3f1ee;padding:16px 22px;font-size:16px;letter-spacing:.04em">THE GRAY HOUSE</div>' +
    '<div style="padding:22px;font-size:15px;line-height:1.55"><h2 style="font-size:19px;margin:0 0 12px;color:#4b4c57">' + esc_(title) + '</h2>' + body +
    (button ? '<p style="margin:22px 0 6px"><a href="' + button.url + '" style="background:#4b4c57;color:#ffffff;text-decoration:none;padding:12px 20px;border-radius:999px;display:inline-block;font-weight:bold">' + esc_(button.label) + '</a></p>' : '') +
    '</div><div style="padding:14px 22px;border-top:1px solid #e4dfd9;font-size:12px;color:#6e6a69">The Gray House Letting &amp; Property Management Ltd · Birchgrove House, 102–104 Caerphilly Road, Cardiff CF14 4AG · ' + CONFIG.OFFICE_PHONE + '</div></div></div>';
}
function rows_(pairs) {
  return '<table style="border-collapse:collapse;width:100%;margin:10px 0;font-size:14px">' + pairs.filter(p => p[1]).map(p =>
    '<tr><td style="padding:6px 10px 6px 0;color:#6e6a69;vertical-align:top;width:38%">' + esc_(p[0]) + '</td><td style="padding:6px 0;vertical-align:top">' + p[1] + '</td></tr>').join('') + '</table>';
}
const PDF_CSS = 'body{font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#2e2f36;line-height:1.5}' +
  '.paper{page-break-after:always;padding:6px 4px}.paper:last-child{page-break-after:auto}' +
  '.lh{border-bottom:1px solid #e4dfd9;padding-bottom:10px;margin-bottom:14px;overflow:hidden}.lh img{height:46px;float:left}.lh .addr{float:right;text-align:right;font-size:9px;color:#6e6a69}' +
  '.sub{font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:#857a74;clear:both}h3{font-size:16px;margin:4px 0 10px;color:#4b4c57}h4{font-size:12px;margin:14px 0 6px;color:#4b4c57}' +
  'table{border-collapse:collapse;width:100%;margin:6px 0}th,td{text-align:left;padding:5px 8px 5px 0;border-bottom:1px solid #e4dfd9;vertical-align:top;font-size:10.5px}th{font-weight:normal;color:#4b4c57;width:45%}' +
  '.ref{color:#857a74;font-size:9px}.st-ok{color:#4c7559}.st-arr{color:#9a6a24}.st-mut,.muted{color:#6e6a69}' +
  '.sigs{margin-top:16px;overflow:hidden}.sigs .s{float:left;width:46%;margin-right:4%}.ink{font-family:"Brush Script MT",cursive;font-size:24px;border-bottom:1px solid #c5bdb3;min-height:34px;color:#4b4c57}.ink.blank{font-family:Arial,sans-serif;font-size:10px;color:#c5bdb3}.who{font-size:9.5px;color:#6e6a69}' +
  '.audit{clear:both;margin-top:14px;padding:8px 10px;background:#f5f3f0;font-size:9px;color:#6e6a69}p{margin:0 0 6px}';
function pdf_(html, name) {
  const doc = '<html><head><meta charset="utf-8"><style>' + PDF_CSS + '</style></head><body>' + html + '</body></html>';
  return Utilities.newBlob(doc, 'text/html', 'doc.html').getAs('application/pdf').setName(name);
}
function safeName_(s) { return String(s).replace(/[\\/:*?"<>|]+/g, ' ').trim(); }

/* ---------- 1. the signed pack ---------- */
function submitPack(p) {
  const w = p.who || {};
  const name = clean_(p.name, 120), email = clean_(w.email, 160);
  if (!name || !okEmail_(email)) throw new Error('Please check your name and email address.');
  throttle_('pack:' + email.toLowerCase(), 4); throttle_('pack:all', 40);
  if (!p.html || String(p.html).length > 3e6) throw new Error('The signed pack could not be read.');
  const addr = clean_(p.addr, 200), ref = clean_(p.ref, 40);
  const pack = pdf_(p.html, safeName_('Landlord pack - ' + name + ' - ' + (w.addr1 || addr) + '.pdf'));
  const files = (p.files || []).slice(0, 30);
  let total = 0;
  const uploads = files.map(f => { const b = Utilities.newBlob(Utilities.base64Decode(f.data), f.type || 'application/octet-stream', safeName_(f.name || 'upload')); total += b.getBytes().length; return b; });
  if (total > 20 * 1024 * 1024) throw new Error('The uploads are too large to email. Please send smaller copies.');
  let terms = null;
  try { terms = UrlFetchApp.fetch(CONFIG.SITE_URL + CONFIG.TERMS_PDF).getBlob().setName('The Gray House - Terms of Business v9.0 (Wales).pdf'); } catch (e) { console.error('Terms PDF not attached: ' + e); }
  const list = arr => arr && arr.length ? '<ul style="margin:6px 0 0 18px;padding:0">' + arr.map(x => '<li>' + esc_(x) + '</li>').join('') + '</ul>' : '';
  send_({
    to: CONFIG.MICHELLE_EMAIL, replyTo: email,
    subject: 'Onboarding complete: ' + name + ', ' + (w.addr1 || addr),
    htmlBody: email_('New landlord onboarding', '<p>' + esc_(name) + (p.name2 ? ' and ' + esc_(p.name2) : '') + ' signed the Terms of Business, onboarding checklist and property questionnaire.</p>' +
      rows_([['Landlord', esc_(name) + (p.name2 ? ' and ' + esc_(p.name2) : '')], ['Email', esc_(email)], ['Phone', esc_(w.phone)], ['Property', esc_(addr)], ['Service', esc_(p.tier)], ['Reference', esc_(ref)], ['Signed', esc_(fmt_(new Date(p.signedAt || Date.now()), 'd MMMM yyyy, HH:mm'))]]) +
      '<p><strong>Attached:</strong> the signed pack (PDF)' + (uploads.length ? ' and ' + uploads.length + ' upload' + (uploads.length > 1 ? 's' : '') + ': ' + files.map(f => esc_(f.label || f.name)).join(', ') : '') + '.</p>' +
      (p.arrange && p.arrange.length ? '<p><strong>You have been asked to arrange:</strong>' + list(p.arrange) + '</p>' : '') +
      (p.outstanding && p.outstanding.length ? '<p><strong>Still to come from the landlord:</strong>' + list(p.outstanding) + '</p>' : '') +
      '<p>They are now booking a valuation visit. You will get a separate email when they do.</p>'),
    attachments: [pack].concat(uploads)
  });
  send_({
    to: email,
    subject: 'Your signed Terms of Business with The Gray House',
    htmlBody: email_('Thank you, ' + (w.first || name), '<p>Thank you for signing your Terms of Business, onboarding checklist and property questionnaire for <strong>' + esc_(addr) + '</strong>. Your signed copy is attached' + (terms ? ', with the full wording of our Terms of Business' : '') + '.</p><p>Your reference is ' + esc_(ref) + '. If you haven\'t booked your valuation visit yet, go back to the page to choose a time.</p><p>' + esc_(CONFIG.MICHELLE_NAME) + '<br>The Gray House</p>'),
    attachments: [pack.copyBlob()].concat(terms ? [terms] : [])
  });
  return { sent: true };
}

/* ---------- 2. availability ---------- */
function getSlots() {
  const now = new Date(), earliest = new Date(now.getTime() + CONFIG.MIN_NOTICE_HOURS * 3600e3), end = new Date(now.getTime() + (CONFIG.DAYS_AHEAD + 1) * 864e5);
  const busy = cal_().getEvents(now, end).filter(ev => !ev.isAllDayEvent()).map(ev => [ev.getStartTime().getTime(), ev.getEndTime().getTime()]);
  const out = {};
  for (let i = 0; i <= CONFIG.DAYS_AHEAD; i++) {
    const d = new Date(now.getTime() + i * 864e5), day = fmt_(d, 'yyyy-MM-dd'), dow = Number(fmt_(d, 'u'));
    if (dow > 5) continue;
    const free = CONFIG.SLOT_TIMES.filter(t => {
      const s = slotDate_(day, t).getTime(), e = s + CONFIG.VISIT_MINUTES * 60000;
      return s >= earliest.getTime() && !busy.some(b => b[0] < e && b[1] > s);
    });
    if (free.length) out[day] = free;
  }
  return out;
}

/* ---------- 3. book the valuation ---------- */
function book(p) {
  const w = p.who || {}, name = clean_(p.name, 120), email = clean_(w.email, 160);
  if (!name || !okEmail_(email)) throw new Error('Please check your name and email address.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(p.day) || CONFIG.SLOT_TIMES.indexOf(p.time) < 0) throw new Error('Please choose a time again.');
  throttle_('book:' + email.toLowerCase(), 4);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const free = getSlots();
    if (!free[p.day] || free[p.day].indexOf(p.time) < 0) throw new Error('Sorry, that time has just been taken. Please choose another.');
    const start = slotDate_(p.day, p.time), end = new Date(start.getTime() + CONFIG.VISIT_MINUTES * 60000);
    const addr = clean_(p.addr, 200);
    const ev = cal_().createEvent('Valuation visit: ' + name + ', ' + clean_(w.addr1, 80), start, end, {
      location: addr,
      description: 'Landlord: ' + name + (p.name2 ? ' and ' + clean_(p.name2, 120) : '') + '\nPhone: ' + clean_(w.phone, 40) + '\nEmail: ' + email + '\nService: ' + clean_(p.tierName, 60) + '\nReference: ' + clean_(p.ref, 40) + (p.note ? '\nNote: ' + clean_(p.note, 400) : '') + '\n\nBooked through The Gray House onboarding page.'
    });
    const v = { first: w.first, last: w.last, email, phone: w.phone, addr1: w.addr1, town: w.town, pc: w.pc, second: w.second ? '1' : '', s_first: w.s_first, s_last: w.s_last, tier: p.tier, tierName: p.tierName, furnish: p.furnish, signedAt: p.signedAt, ref: p.ref, note: p.note };
    Object.keys(v).forEach(k => setTag_(ev, k, clean_(v[k], 400)));
    const when = when_(ev), link = CONFIG.SITE_URL + '?r=' + token_('r', ev.getId());
    send_({
      to: CONFIG.MICHELLE_EMAIL, replyTo: email,
      subject: 'New booking: valuation visit, ' + fmt_(start, 'EEE d MMM, HH:mm'),
      htmlBody: email_('New valuation visit', rows_([['When', esc_(when) + ' (' + CONFIG.VISIT_MINUTES + ' mins)'], ['Where', esc_(addr)], ['Landlord', esc_(name) + ' · ' + esc_(w.phone) + ' · ' + esc_(email)], ['Service', esc_(p.tierName)], ['Note', esc_(p.note)]]) +
        '<p>It is in your calendar. After the visit, use the button below to enter the agreed rent. The Rental Value Confirmation is then created with your signature and sent to ' + esc_(w.first || name) + ' to sign.</p>', { url: link, label: 'Enter the agreed rent' })
    });
    send_({
      to: email,
      subject: 'Your valuation visit is booked: ' + fmt_(start, 'EEE d MMM, HH:mm'),
      htmlBody: email_('Your visit is booked', rows_([['When', esc_(when)], ['Where', esc_(addr)], ['With', esc_(CONFIG.MICHELLE_NAME)], ['Length', 'About ' + CONFIG.VISIT_MINUTES + ' minutes']]) +
        '<p>Michelle will look round the property with you and agree the monthly rent we\'ll market it at. Afterwards you\'ll get your Rental Value Confirmation to sign by email.</p><p>Need to change the time? Reply to this email or call ' + CONFIG.OFFICE_PHONE + '.</p>')
    });
    return { when };
  } finally { lock.releaseLock(); }
}

/* ---------- 4. Michelle enters the rent; the landlord signs ---------- */
function getVisit(t) {
  const kind = String(t || '').charAt(0);
  if (kind !== 'r' && kind !== 's') throw new Error('This link is not valid.');
  const ev = event_(readToken_(t, kind)), v = tags_(ev);
  if (!v.email) throw new Error('This link is not valid.');
  v.day = fmt_(ev.getStartTime(), 'yyyy-MM-dd'); v.time = fmt_(ev.getStartTime(), 'HH:mm');
  return v;
}
function issueLetter(t, L) {
  const ev = event_(readToken_(t, 'r')), v = tags_(ev);
  if (v.lSignedAt) throw new Error(v.first + ' has already signed the Rental Value Confirmation.');
  const rent = clean_(L.rent, 20);
  if (!/^[\d,]+(\.\d{1,2})?$/.test(rent)) throw new Error('Please enter the agreed monthly rent.');
  setTag_(ev, 'rent', rent); setTag_(ev, 'from', clean_(L.from, 20)); setTag_(ev, 'pets', clean_(L.pets, 40)); setTag_(ev, 'notes', clean_(L.notes, 800)); setTag_(ev, 'issuedAt', new Date().toISOString());
  ev.setTitle('Rental value sent: ' + v.first + ' ' + v.last + ', ' + v.addr1);
  const link = CONFIG.SITE_URL + '?s=' + token_('s', ev.getId());
  send_({
    to: v.email,
    subject: 'Your Rental Value Confirmation is ready to sign',
    htmlBody: email_('Your Rental Value Confirmation', '<p>Dear ' + esc_(v.first) + ',</p><p>Thank you for meeting ' + esc_(CONFIG.MICHELLE_NAME) + ' at ' + esc_(v.addr1) + '. Your Rental Value Confirmation, at <strong>£' + esc_(rent) + ' per month</strong>, is ready to read and sign. It is supplementary to your Terms of Business.</p>', { url: link, label: 'Read and sign' })
  });
  send_({
    to: CONFIG.MICHELLE_EMAIL,
    subject: 'Sent to sign: Rental Value Confirmation, ' + v.addr1,
    htmlBody: email_('Sent to ' + v.first + ' to sign', '<p>The Rental Value Confirmation for ' + esc_(v.addr1) + ' at <strong>£' + esc_(rent) + ' pcm</strong> has been emailed to ' + esc_(v.first + ' ' + v.last) + '. You\'ll both receive the signed PDF when they sign.</p>')
  });
  return { sent: true };
}
function signLetter(t, sig, html) {
  const ev = event_(readToken_(t, 's')), v = tags_(ev);
  if (!v.rent) throw new Error('This confirmation is not ready yet.');
  if (v.lSignedAt) return { already: true };
  sig = clean_(sig, 120);
  if (!sig) throw new Error('Please add your signature.');
  if (!html || String(html).length > 2e6) throw new Error('The confirmation could not be read.');
  const now = new Date().toISOString();
  setTag_(ev, 'lsig', sig); setTag_(ev, 'lSignedAt', now);
  ev.setTitle('Rental value agreed: ' + v.first + ' ' + v.last + ', ' + v.addr1 + ' (£' + v.rent + ' pcm)');
  const pdf = pdf_(html, safeName_('Rental Value Confirmation - ' + v.first + ' ' + v.last + ' - ' + v.addr1 + '.pdf'));
  send_({
    to: v.email,
    subject: 'Signed: your Rental Value Confirmation',
    htmlBody: email_('All signed', '<p>Dear ' + esc_(v.first) + ',</p><p>Thank you. Your signed Rental Value Confirmation for ' + esc_(v.addr1) + ' at <strong>£' + esc_(v.rent) + ' per month</strong> is attached. We\'ll be in touch about the next steps for marketing your property.</p><p>' + esc_(CONFIG.MICHELLE_NAME) + '<br>The Gray House</p>'),
    attachments: [pdf]
  });
  send_({
    to: CONFIG.MICHELLE_EMAIL, replyTo: v.email,
    subject: 'Signed: Rental Value Confirmation, ' + v.addr1 + ' (£' + v.rent + ' pcm)',
    htmlBody: email_('Rental value agreed', '<p>' + esc_(v.first + ' ' + v.last) + ' has signed the Rental Value Confirmation for <strong>' + esc_(v.addr1) + '</strong> at <strong>£' + esc_(v.rent) + ' pcm</strong>. The signed PDF is attached, and the same copy has gone to ' + esc_(v.email) + '.</p>'),
    attachments: [pdf.copyBlob()]
  });
  return { signed: true };
}
