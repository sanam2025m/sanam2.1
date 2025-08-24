/** Sanam Attendance – Google Apps Script (STRICT A..L ORDER)
 * Sheet tab: "البيانات"
 * Timezone : Asia/Riyadh
 * Images   : uploaded to FOLDER_ID
 */
const SHEET_NAME = 'البيانات';
const FOLDER_ID  = '1Qb89RWoNedpJkNYn3AyryOD2J1jh8q02';
const TZ         = 'Asia/Riyadh';

/** العناوين النهائية (A..L) */
const HEADERS = [
  'الطابع الزمني',        // A
  'الاسم الرباعي',        // B
  'رقم الهوية',           // C
  'رقم الجوال',           // D
  'الموقع',               // E
  'الوردية',              // F
  'الحالة',               // G
  'نوع الدوام',           // H
  'نوع العملية',          // I
  'فرق الدقائق',          // J
  'سبب الانصراف المبكر',  // K
  'رابط الصورة'           // L
];

function doGet(){ return json_({ ok:true, message:'Sanam WebApp up', headers: HEADERS }); }
function doOptions(){ return json_({ ok:true }); }

function doPost(e){
  const lock = LockService.getScriptLock();
  try {
    lock.tryLock(30000);

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    ensureHeadersTop_(sheet); // يثبت الصف 1 بنفس الترتيب تمامًا

    // --- قراءة البيانات ---
    let data = {}, photoBlob = null;
    if (e && e.postData && e.postData.type === 'application/json'){
      data = safelyParseJson_(e.postData.contents || '{}');
      if (data.imageBase64){
        const conv = base64ToBlob_(data.imageBase64);
        conv.blob.setName(makeFilename_(data, conv.ext));
        photoBlob = conv.blob;
      }
    }
    if ((!photoBlob || !data || Object.keys(data).length === 0) && e && e.parameter){
      const p = e.parameter;
      data = {
        action:     (p.action     || '').trim(),
        fullName:   (p.fullName   || '').trim(),
        nationalId: (p.nationalId || '').trim(),
        phone:      (p.phone      || '').trim(),
        location:   (p.location   || '').trim(),
        shift:      (p.shift      || '').trim(),
        workType:   (p.workType   || '').trim(),
        status:     (p.status     || '').trim(),
        minutes:     p.minutes    || '',
        reason:     (p.reason     || '').trim()
      };
      if (e.files && e.files.photo){
        photoBlob = e.files.photo;
        const ct = photoBlob.getContentType();
        const ext = (ct && ct.indexOf('/')>-1) ? ct.split('/')[1] : 'jpg';
        photoBlob.setName(makeFilename_(data, ext));
      }
      if (!photoBlob && p.imageBase64){
        const conv = base64ToBlob_(p.imageBase64);
        conv.blob.setName(makeFilename_(data, conv.ext));
        photoBlob = conv.blob;
      }
    }

    // --- رفع الصورة (اختياري) ---
    let imageUrl = '';
    if (photoBlob){
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const file = folder.createFile(photoBlob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(_){}
      imageUrl = `https://drive.google.com/file/d/${file.getId()}/view?usp=drivesdk`;
    }

    // --- كتابة صف واحد بترتيب A..L مهما كان شكل الشيت ---
    const ts = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    const row = [
      ts,                     // A: الطابع الزمني
      data.fullName,          // B: الاسم الرباعي
      data.nationalId,        // C: رقم الهوية
      data.phone,             // D: رقم الجوال
      data.location,          // E: الموقع
      data.shift,             // F: الوردية
      data.status,            // G: الحالة
      data.workType,          // H: نوع الدوام
      data.action,            // I: نوع العملية
      data.minutes,           // J: فرق الدقائق
      data.reason,            // K: سبب الانصراف المبكر
      imageUrl                // L: رابط الصورة
    ];

    // نستخدم appendRow لتعبئة من العمود A مباشرة
    sheet.appendRow(row);

    return json_({ ok:true, ts, imageUrl });
  } catch (err) {
    return json_({ ok:false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_){}
  }
}

/* ========= Helpers ========= */
function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin','*')
    .setHeader('Access-Control-Allow-Methods','GET,POST,OPTIONS')
    .setHeader('Access-Control-Allow-Headers','Content-Type');
}
function safelyParseJson_(txt){ try{ return JSON.parse(txt); } catch(_){ return {}; } }
function base64ToBlob_(dataUrlOrRaw){
  let mime = 'image/jpeg', b64 = dataUrlOrRaw || '';
  const m = String(b64).match(/^data:(.+?);base64,(.+)$/);
  if (m){ mime = m[1]; b64 = m[2]; }
  const bytes = Utilities.base64Decode(b64);
  const ext = (mime.indexOf('/')>-1) ? mime.split('/')[1] : 'jpg';
  return { blob: Utilities.newBlob(bytes, mime, 'upload.'+ext), ext };
}
function makeFilename_(d, ext){
  const ts = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss');
  const action = (d.action || 'action').toLowerCase();
  const nid = (d.nationalId || 'id').replace(/\D+/g,'');
  return `sanam_${action}_${nid}_${ts}.${ext||'jpg'}`;
}

/** يكتب العناوين في الصف 1 إذا كانت مختلفة أو ناقصة */
function ensureHeadersTop_(sheet){
  const need =
    sheet.getLastRow() === 0 ||
    !rangeEquals_(sheet.getRange(1,1,1,HEADERS.length).getValues()[0], HEADERS);
  if (need){
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
}
function rangeEquals_(arr, ref){
  const a = (arr || []).map(v => String(v||'').trim());
  const b = (ref || []).map(v => String(v||'').trim());
  if (a.length < b.length) return false;
  for (let i=0;i<b.length;i++){ if (a[i] !== b[i]) return false; }
  return true;
}