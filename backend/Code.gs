/**
 * يوم ريادة الأعمال – كلية البريمي الجامعية
 * موقع التسجيل + حفظ البيانات في Google Sheets + إشعار بالإيميل
 */

// ====== الإعدادات ======
const NOTIFY_EMAIL = true;      // إرسال إيميل لك مع كل تسجيل جديد (اجعلها false لإيقافه)
const EXTRA_EMAILS = '';        // إيميلات إضافية تستلم الإشعار، مفصولة بفاصلة (اختياري)
const TZ = 'Asia/Muscat';
const SHEET_ID = 'YOUR_SHEET_ID'; // ملف التسجيلات
function book_() { return SpreadsheetApp.openById(SHEET_ID); }

// أرقام جامعية تحصل على طاولات VIP (بترتيب التسجيل: VIP-01, VIP-02 ...)
const VIP_IDS = {
  // 'الرقم الجامعي': 'الاسم',
};

const SHEETS = {
  attendee: {
    name: 'الحضور',
    headers: ['رقم التسجيل', 'تاريخ التسجيل', 'الاسم الكامل', 'رقم الهاتف', 'البريد الإلكتروني', 'الصفة', 'الحالة', 'ملاحظات'],
    widths: [110, 140, 200, 140, 210, 250, 120, 240],
    statusCol: 7, phoneCol: 4,
  },
  project: {
    name: 'المشاريع',
    headers: ['رقم المشروع', 'تاريخ التسجيل', 'اسم صاحب المشروع', 'رقم الهاتف', 'الرقم الجامعي', 'اسم المشروع', 'نوع المشاركة', 'نبذة عن المشروع', 'إنستغرام', 'يحتاج طاولة', 'الحالة', 'ملاحظات'],
    widths: [110, 140, 190, 140, 120, 180, 140, 340, 150, 90, 120, 240],
    statusCol: 11, phoneCol: 4, nameCol: 6, idCol: 5,
  },
};
const STATUSES = ['جديد', 'تم التواصل', 'مؤكد', 'مرفوض'];
const ROLES = { buc_student: 'طالب/طالبة في كلية البريمي الجامعية', other_student: 'طالب/طالبة في كلية أو جامعة أخرى', visitor: 'زائر' };
const PTYPES = { student: 'مشروع طلابي', private: 'مشروع تجاري خاص', startup: 'شركة ناشئة', other: 'مشروع آخر' };

// ====== صفحة الموقع ======
const PAGE_FILE_ID = 'YOUR_DRIVE_FILE_ID';
const SITE_FILE_ID = 'YOUR_DRIVE_FILE_ID'; // نسخة الموقع المنشورة على GitHub Pages
const PRINT_KEY = 'YOUR_PRINT_KEY'; // صفحة طباعة البطاقات: رابط الموقع + ?print=المفتاح // تصميم الموقع (ملف HTML في Drive)
function doGet(e) {
  if (e && e.parameter && e.parameter.raw === 'site') {
    return ContentService.createTextOutput(DriveApp.getFileById(SITE_FILE_ID).getBlob().getDataAsString('UTF-8')).setMimeType(ContentService.MimeType.TEXT);
  }
  const cache = CacheService.getScriptCache();
  let html = cache.get('page');
  if (!html) {
    html = DriveApp.getFileById(PAGE_FILE_ID).getBlob().getDataAsString('UTF-8');
    try { cache.put('page', html, 21600); } catch (e) {}
  }
  if (e && e.parameter && e.parameter.print === PRINT_KEY) {
    html = html.replace('<body>', () => '<body><script>window.PRINT_DATA=' + JSON.stringify(cards_()).replace(/</g, '\\u003c') + ';</script>');
  }
  return HtmlService.createHtmlOutput(html)
    .setTitle('التسجيل | يوم ريادة الأعمال – كلية البريمي الجامعية')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ====== واجهة برمجية للموقع الخارجي (GitHub Pages) ======
function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  const out = register(body);
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ====== استقبال التسجيل ======
function register(body) {
  body = body || {};
  if (body.website) return { ok: true, ref: makeRef_(body.type) };                    // مصيدة الروبوتات
  if (Number(body.elapsed_ms) > 0 && Number(body.elapsed_ms) < 2000) return { ok: false, error: 'تم الإرسال بسرعة كبيرة، يرجى المحاولة مرة أخرى' };

  const v = validate_(body);
  if (v.errors) return { ok: false, error: 'يرجى مراجعة الحقول المطلوبة', fields: v.errors };
  const d = v.data;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'الضغط عالٍ حاليًا، يرجى المحاولة بعد لحظات' };
  try {
    const cfg = SHEETS[d.type];
    const sh = ensureSheet_(cfg);

    // منع التسجيل المكرر
    const last = sh.getLastRow();
    if (last > 1) {
      const phones = sh.getRange(2, cfg.phoneCol, last - 1, 1).getValues().map((r) => String(r[0]).replace(/^'/, ''));
      if (d.type === 'attendee' && phones.indexOf(d.phone) !== -1) {
        return { ok: false, duplicate: true, error: 'هذا الرقم مسجّل مسبقًا للحضور. لا داعي للتسجيل مرة أخرى.' };
      }
      if (d.type === 'project') {
        const sids = sh.getRange(2, cfg.idCol, last - 1, 1).getValues().map((r) => String(r[0]).replace(/^'/, '').trim());
        if (sids.indexOf(d.student_id) !== -1) {
          return { ok: false, duplicate: true, error: 'هذا الرقم الجامعي مسجّل مسبقًا. لا داعي للتسجيل مرة أخرى.' };
        }
        const names = sh.getRange(2, cfg.nameCol, last - 1, 1).getValues();
        for (let i = 0; i < phones.length; i++) {
          if (phones[i] === d.phone && String(names[i][0]).trim().toLowerCase() === d.project_name.toLowerCase()) {
            return { ok: false, duplicate: true, error: 'تم تسجيل هذا المشروع بنفس الرقم مسبقًا. سيتواصل معك فريق التنظيم قريبًا.' };
          }
        }
      }
    }

    const ref = d.type === 'project' ? (VIP_IDS[d.student_id] ? nextVip_() : nextNumber_()) : makeRef_(d.type);
    const now = new Date();
    const row = d.type === 'attendee'
      ? [ref, now, safe_(d.full_name), "'" + d.phone, safe_(d.email || ''), ROLES[d.attendee_role], 'جديد', '']
      : [ref, now, safe_(d.full_name), "'" + d.phone, d.student_id ? "'" + d.student_id : '', safe_(d.project_name), PTYPES[d.project_type], safe_(d.description), d.instagram || '', d.wants_booth ? 'نعم' : 'لا', 'جديد', ''];
    sh.appendRow(row);
    SpreadsheetApp.flush();
    lock.releaseLock();

    if (NOTIFY_EMAIL) notify_(d, ref, now);
    return { ok: true, ref: ref };
  } catch (e) {
    console.error(e);
    return { ok: false, error: 'حدث خطأ غير متوقع، يرجى المحاولة لاحقًا' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ====== التحقق من البيانات ======
function validate_(b) {
  const errors = {};
  const type = b.type === 'project' ? 'project' : b.type === 'attendee' ? 'attendee' : null;
  if (!type) return { errors: { type: 'نوع التسجيل غير صالح' } };
  const full_name = clean_(b.full_name, 100);
  if (full_name.length < 3) errors.full_name = 'يرجى كتابة الاسم الكامل';
  const phone = normalizePhone_(b.phone);
  if (!phone) errors.phone = 'يرجى إدخال رقم هاتف صحيح (مثال: 9XXXXXXX)';
  const out = { type: type, full_name: full_name, phone: phone };

  if (type === 'attendee') {
    const email = clean_(b.email, 254).toLowerCase();
    if (email && !/^[^\s@]{1,64}@[^\s@]{1,255}\.[^\s@]{2,}$/.test(email)) errors.email = 'البريد الإلكتروني غير صحيح';
    if (!ROLES[b.attendee_role]) errors.attendee_role = 'يرجى اختيار الصفة';
    out.email = email; out.attendee_role = b.attendee_role;
  } else {
    const project_name = clean_(b.project_name, 120);
    if (project_name.length < 2) errors.project_name = 'يرجى كتابة اسم المشروع';
    if (!PTYPES[b.project_type]) errors.project_type = 'يرجى اختيار نوع المشاركة';
    const description = String(b.description || '').replace(/\r\n?/g, '\n').trim().slice(0, 600);
    if (description.length < 10) errors.description = 'يرجى كتابة نبذة قصيرة (10 أحرف على الأقل)';
    let ig = clean_(b.instagram, 120).replace(/^https?:\/\/(www\.)?instagram\.com\//i, '').replace(/[/?#].*$/, '').replace(/^@/, '');
    if (ig && !/^[A-Za-z0-9._]{1,30}$/.test(ig)) errors.instagram = 'اسم الحساب غير صحيح';
    const wb = b.wants_booth === 'yes' ? true : b.wants_booth === 'no' ? false : null;
    if (wb === null) errors.wants_booth = 'يرجى اختيار نعم أو لا';
    out.project_name = project_name; out.project_type = b.project_type; out.description = description;
    out.instagram = ig ? '@' + ig : ''; out.wants_booth = wb;
    const sid = toLatin_(clean_(b.student_id, 30)).replace(/\s/g, '');
    if (!sid) errors.student_id = 'يرجى كتابة الرقم الجامعي';
    else if (!/^[A-Za-z0-9\-]{3,20}$/.test(sid)) errors.student_id = 'الرقم الجامعي غير صحيح';
    out.student_id = sid;
  }
  return Object.keys(errors).length ? { errors: errors } : { data: out };
}
function clean_(v, max) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max); }
function safe_(s) { s = String(s || ''); return /^[=+\-@]/.test(s) ? "'" + s : s; } // منع تنفيذ الصيغ
function toLatin_(s) {
  return String(s || '').replace(/[٠-٩]/g, (d) => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[۰-۹]/g, (d) => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d));
}
function normalizePhone_(raw) {
  let s = toLatin_(raw).replace(/[\s\-().]/g, '');
  if (s.indexOf('00') === 0) s = '+' + s.slice(2);
  if (/^\+968\d{8}$/.test(s)) s = s.slice(4);
  else if (/^968[79]\d{7}$/.test(s)) s = s.slice(3);
  if (/^[792]\d{7}$/.test(s)) return '+968' + s;
  if (/^\+[1-9]\d{7,14}$/.test(s) && s.indexOf('+968') !== 0) return s;
  return null;
}
/** بيانات بطاقات الطاولات لكل المشاريع غير المرفوضة */
function cards_() {
  const sh = book_().getSheetByName(SHEETS.project.name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 11).getValues()
    .filter((r) => r[0] && r[10] !== 'مرفوض')
    .map((r) => ({ num: String(r[0]), owner: String(r[2]), project: String(r[5]) }));
}

/** تشغيل يدوي: يمسح نسخة الصفحة المخزنة حتى يظهر آخر تحديث للتصميم فورًا */
function resetPage() { CacheService.getScriptCache().remove('page'); }

/** رقم تسلسلي لكل مشروع: P-001, P-002 ... (يُستدعى داخل القفل) */
function nextNumber_() {
  const p = PropertiesService.getScriptProperties();
  const n = Number(p.getProperty('project_seq') || 0) + 1;
  p.setProperty('project_seq', String(n));
  return 'P-' + ('00' + n).slice(-3);
}

/** رقم VIP لأرقام الفريق الجامعية: VIP-01, VIP-02 ... */
function nextVip_() {
  const p = PropertiesService.getScriptProperties();
  const n = Number(p.getProperty('vip_seq') || 0) + 1;
  p.setProperty('vip_seq', String(n));
  return 'VIP-' + ('0' + n).slice(-2);
}

/** تشغيل يدوي: يحذف التسجيلات التجريبية (اسمها يبدأ بـ «تجربة من Claude») ويعيد الترقيم */
function removeTestRows() {
  const sh = book_().getSheetByName(SHEETS.project.name);
  for (let r = sh.getLastRow(); r >= 2; r--) {
    if (String(sh.getRange(r, 3).getValue()).indexOf('تجربة من Claude') === 0) sh.deleteRow(r);
  }
  const last = sh.getLastRow();
  let max = 0;
  if (last >= 2) sh.getRange(2, 1, last - 1, 1).getValues().forEach((v) => { const m = String(v[0]).match(/^P-(\d+)$/); if (m) max = Math.max(max, Number(m[1])); });
  PropertiesService.getScriptProperties().setProperty('project_seq', String(max));
  let vmax = 0;
  if (last >= 2) sh.getRange(2, 1, last - 1, 1).getValues().forEach((v) => { const t = String(v[0]); if (t.indexOf('VIP-') === 0) vmax = Math.max(vmax, Number(t.slice(4)) || 0); });
  PropertiesService.getScriptProperties().setProperty('vip_seq', String(vmax));
}

function makeRef_(type) {
  const a = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 6; i++) s += a.charAt(Math.floor(Math.random() * a.length));
  return (type === 'project' ? 'P-' : 'A-') + s;
}

// ====== تجهيز الجدول ======
function ensureSheet_(cfg) {
  const ss = book_();
  return ss.getSheetByName(cfg.name) || (setup(), ss.getSheetByName(cfg.name));
}

/** شغّل هذه الدالة مرة واحدة من محرر Apps Script لتجهيز الجدول ومنح الصلاحيات */
function setup() {
  const ss = book_();
  ss.setSpreadsheetTimeZone(TZ);
  ss.setSpreadsheetLocale('ar');

  Object.keys(SHEETS).forEach((k, idx) => {
    const cfg = SHEETS[k];
    let sh = ss.getSheetByName(cfg.name);
    if (sh && sh.getLastRow() <= 1 && sh.getRange(1, 1, 1, sh.getMaxColumns()).getValues()[0].filter(String).length !== cfg.headers.length) { ss.deleteSheet(sh); sh = null; }
    if (!sh) sh = ss.insertSheet(cfg.name, idx);
    sh.setRightToLeft(true);
    const n = cfg.headers.length;
    sh.getRange(1, 1, 1, n).setValues([cfg.headers])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#26495d')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
    sh.setRowHeight(1, 34);
    sh.setFrozenRows(1);
    cfg.widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
    const max = sh.getMaxRows();
    sh.getRange(2, 2, max - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange(2, cfg.phoneCol, max - 1, 1).setNumberFormat('@').setHorizontalAlignment('left');
    sh.getRange(2, 1, max - 1, n).setVerticalAlignment('top').setWrap(true);

    // قائمة الحالة + ألوانها
    const st = sh.getRange(2, cfg.statusCol, max - 1, 1);
    st.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());
    const colors = { 'جديد': ['#e6eff5', '#2f5f7c'], 'تم التواصل': ['#f6efe0', '#7d6334'], 'مؤكد': ['#e3f1e8', '#2f6b4c'], 'مرفوض': ['#f8e4e2', '#a33a32'] };
    sh.setConditionalFormatRules(Object.keys(colors).map((s) =>
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s).setBackground(colors[s][0]).setFontColor(colors[s][1]).setBold(true).setRanges([st]).build()));
    if (!sh.getFilter()) sh.getRange(1, 1, Math.max(sh.getLastRow(), 2), n).createFilter();
    sh.setTabColor(k === 'project' ? '#c2a772' : '#6095b3');
  });

  // ورقة الملخص
  let sum = ss.getSheetByName('الملخص');
  if (!sum) sum = ss.insertSheet('الملخص', 0);
  sum.setRightToLeft(true).setTabColor('#26495d');
  sum.getRange('A1:B1').setValues([['يوم ريادة الأعمال – ملخص التسجيلات', '']]).merge()
    .setFontSize(15).setFontWeight('bold').setFontColor('#ffffff').setBackground('#26495d');
  sum.getRange('A3:B9').setValues([
    ['المسجلون للحضور', "=COUNTA('الحضور'!A2:A)"],
    ['المشاريع المسجلة', "=COUNTA('المشاريع'!A2:A)"],
    ['طلبات طاولة', "=COUNTIF('المشاريع'!J2:J,\"نعم\")"],
    ['حضور بانتظار التواصل (جديد)', "=COUNTIF('الحضور'!G2:G,\"جديد\")"],
    ['مشاريع بانتظار التواصل (جديد)', "=COUNTIF('المشاريع'!K2:K,\"جديد\")"],
    ['مشاريع مؤكدة', "=COUNTIF('المشاريع'!K2:K,\"مؤكد\")"],
    ['إجمالي التسجيلات', '=B3+B4'],
  ]);
  sum.getRange('A3:A9').setFontWeight('bold').setBackground('#f6f4f0');
  sum.getRange('B3:B9').setFontSize(16).setFontWeight('bold').setFontColor('#26495d').setHorizontalAlignment('center');
  sum.setColumnWidth(1, 280); sum.setColumnWidth(2, 140);
  sum.setRowHeights(3, 7, 30);

  // حذف الورقة الافتراضية الفارغة
  ss.getSheets().forEach((s) => {
    const nm = s.getName();
    if (['الحضور', 'المشاريع', 'الملخص'].indexOf(nm) === -1 && s.getLastRow() === 0 && ss.getSheets().length > 3) ss.deleteSheet(s);
  });
}

// ====== الإشعار بالإيميل ======
function notify_(d, ref, now) {
  try {
    const to = [Session.getEffectiveUser().getEmail()].concat(EXTRA_EMAILS.split(',').map((s) => s.trim()).filter(String)).join(',');
    const isP = d.type === 'project';
    const rows = isP
      ? [['صاحب المشروع', d.full_name], ['الهاتف', d.phone], ['الرقم الجامعي', d.student_id || '—'], ['المشروع', d.project_name], ['النوع', PTYPES[d.project_type]],
         ['النبذة', d.description], ['إنستغرام', d.instagram || '—'], ['يحتاج طاولة', d.wants_booth ? 'نعم' : 'لا']]
      : [['الاسم', d.full_name], ['الهاتف', d.phone], ['البريد', d.email || '—'], ['الصفة', ROLES[d.attendee_role]]];
    const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    const wa = 'https://wa.me/' + d.phone.replace(/\D/g, '');
    const html = '<div dir="rtl" style="font-family:Tahoma,Arial;font-size:15px;color:#2e3133">'
      + '<h2 style="color:#26495d;margin:0 0 12px">' + (isP ? 'مشروع جديد مسجّل' : 'تسجيل حضور جديد') + ' – ' + ref + '</h2>'
      + '<table cellpadding="8" style="border-collapse:collapse">'
      + rows.map((r) => '<tr><td style="color:#6d7275;border-bottom:1px solid #eee">' + r[0] + '</td><td style="border-bottom:1px solid #eee"><b>' + esc(r[1]).replace(/\n/g, '<br>') + '</b></td></tr>').join('')
      + '</table><p><a href="' + wa + '" style="background:#2f7d57;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none">مراسلة واتساب</a> '
      + '&nbsp; <a href="' + book_().getUrl() + '">فتح جدول التسجيلات</a></p>'
      + '<p style="color:#999;font-size:12px">' + Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm') + '</p></div>';
    MailApp.sendEmail({ to: to, subject: (isP ? (ref.indexOf('VIP') === 0 ? '⭐ ' : '🟡 ') + 'مشروع جديد ' + ref + ': ' + d.project_name : '🔵 تسجيل حضور: ' + d.full_name) + ' | يوم ريادة الأعمال', htmlBody: html, name: 'يوم ريادة الأعمال' });
  } catch (e) { console.error('notify failed', e); }
}
