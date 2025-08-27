/* ===================== الإعدادات ===================== */
const apiKey = "AIzaSyBYPUAnYE8GC4Vx32cDYSb8UH6YV-VWmEA";
const CASES_SHEET_ID = "1k6BSYyEGiezQqubRUDbOFiFy8k34OjZO8ZnNbYi751I";
/* headers on row 2, data from row 3+ */
const CASES_RANGE = "نموذج بلاغات!A2:J";

/* ===================== أدوات مساعدة ===================== */
async function fetchSheetData(sheetId, range){
  try{
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(range)}?key=${apiKey}`;
    const res = await fetch(url);
    const data = await res.json();
    return data.values && data.values.length ? data.values : [];
  }catch(e){
    return [];
  }
}

function parseDateCell(v){
  const m = (v||'').match(/\d{4}[-\/]\d{2}[-\/]\d{2}|\d{2}[-\/]\d{2}[-\/]\d{4}/);
  return m ? new Date(m[0].replace(/-/g,'/')) : null;
}

/* تاريخ فقط للبطاقات */
function formatDateOnly(v){
  const d = parseDateCell(v);
  if(!d) return (v ?? '-') || '-';
  const y = d.getFullYear();
  const m = String(d.getMonth()+1).padStart(2,'0');
  const dd = String(d.getDate()).padStart(2,'0');
  return `${y}-${m}-${dd}`;
}

function normalizeSite(s){
  if (s == null) return '';
  const arabicDigits = {'٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'};
  const replaced = String(s).replace(/[٠-٩]/g, d => arabicDigits[d] || d);
  return replaced.trim().replace(/\s+/g,' ').toLowerCase();
}

function filterByDateRange(values, fromStr, toStr, dateCol){
  if(!values.length || !fromStr || !toStr) return values;
  const f = new Date(fromStr), t = new Date(toStr);
  if(isNaN(f)||isNaN(t)) return values;
  const out = [values[0]];
  for(let i=1;i<values.length;i++){
    const d = parseDateCell(values[i][dateCol]||"");
    if(d && d>=f && d<=t) out.push(values[i]);
  }
  return out;
}

function filterBySite(values, siteName, siteCol){
  if(!values.length || !siteName) return values;
  const wanted = normalizeSite(siteName);
  const out = [values[0]];
  for(let i=1;i<values.length;i++){
    const cell = normalizeSite(values[i][siteCol] ?? '');
    if(cell === wanted) out.push(values[i]);
  }
  return out;
}

/* تحويل أوقات 24h داخل أي خلية إلى 12h (ص/م) */
function convertTimesInText(text){
  if (text == null) return text;
  let s = String(text);
  const hasAmPm = /\b(AM|PM|am|pm|ص|م)\b/.test(s);
  const re = /(\b[01]?\d|2[0-3]):([0-5]\d)(?::([0-5]\d))?/g;
  return s.replace(re, (full, hh, mm, ss) => {
    if (hasAmPm) return full;
    const h = parseInt(hh,10);
    const period = (h < 12) ? 'ص' : 'م';
    let h12 = h % 12; if (h12 === 0) h12 = 12;
    return `${h12}:${mm}${ss ? ':'+ss : ''} ${period}`;
  });
}

/* ========= Linkify: URLs & emails become clickable (handles bare domains) ========= */
function escapeHTML(s){
  return String(s ?? '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

function linkifyHTML(text){
  let s = escapeHTML(text || '');

  // 1) Emails
  const emailRe = /([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g;
  s = s.replace(emailRe, '<a href="mailto:$1">$1</a>');

  // 2) URLs (http/https or bare domains), skip inside existing anchors
  const parts = s.split(/(<a[^>]*>.*?<\/a>)/g);
  const urlRe = /((?:https?:\/\/)?(?:[a-z0-9-]+\.)+[a-z]{2,}(?:\/[^\s<]*)?)/gi;

  for (let i = 0; i < parts.length; i++){
    const chunk = parts[i];
    if (!chunk || /^<a[^>]*>/.test(chunk)) continue;

    parts[i] = chunk.replace(urlRe, (match) => {
      // Trim trailing punctuation
      const trailRe = /[.,;:!?)،\]\}»]+$/;
      let trail = '';
      if (trailRe.test(match)){
        trail = match.match(trailRe)[0];
        match = match.slice(0, -trail.length);
      }
      let href = match;
      if (!/^https?:\/\//i.test(href)) href = 'https://' + href;
      return `<a href="${href}" target="_blank" rel="noopener">${match}</a>${trail}`;
    });
  }

  s = parts.join('');

  // 3) Preserve line breaks (multiple links per cell)
  s = s.replace(/\r\n|\r|\n/g, '<br>');

  return s;
}

/* ===================== جدول الديسكتوب (إخفاء عمود الطابع الزمني) ===================== */
function buildTableHTML(values){
  if(!values || !values.length)
    return `<div class="empty">لا توجد بيانات</div>`;

  const headers = values[0] || [];
  // ignore column 0 (timestamp) in header and rows
  let html = `<table><thead><tr>${
    headers.slice(1).map(h=>`<th>${h||''}</th>`).join('')
  }</tr></thead><tbody>`;

  for(let r=1;r<values.length;r++){
    const row = values[r] || [];
    const padded = Array.from({length: headers.length}, (_, j) => row[j] ?? '');
    html += `<tr>${
      padded.slice(1).map((cell)=>{
        let val = (cell==null || String(cell).trim()==='') ? '-' : String(cell).trim();
        val = convertTimesInText(val);
        const valHTML = (val === '-') ? '-' : linkifyHTML(val);
        return `<td>${valHTML}</td>`;
      }).join('')
    }</tr>`;
  }
  html += `</tbody></table>`;
  return html;
}

/* ===================== اكتشاف الأعمدة من العناوين (أكثر صرامة) ===================== */
function normHeader(s){
  return String(s||'').toLowerCase()
    .replace(/\s+/g,'')
    .replace(/[._\-:/|()]+/g,'')
    .replace(/أ|إ|آ/g,'ا')
    .replace(/ى/g,'ي')
    .replace(/ة/g,'ه');
}
function findFirst(headers, include, exclude=[]){
  const H = headers.map(normHeader);
  const inc = include.map(normHeader);
  const exc = exclude.map(normHeader);
  for(let i=0;i<H.length;i++){
    const h = H[i];
    const okInc = inc.some(p => h.includes(p));
    const okExc = exc.some(p => h.includes(p));
    if(okInc && !okExc) return i;
  }
  return -1;
}
function looksLikeDateOrTimeOnly(s){
  const t = String(s||'').trim();
  if (!t) return false;
  if (/^\d{2,4}[\/\-]\d{1,2}[\/\-]\d{1,2}(\s+\d{1,2}:\d{2}(:\d{2})?\s*(AM|PM|am|pm|ص|م)?)?$/.test(t)) return true;
  if (/^\d{1,2}:\d{2}(:\d{2})?\s*(AM|PM|am|pm|ص|م)?$/.test(t)) return true;
  return false;
}

/* ===================== بطاقات الموبايل ===================== */
let DATE_COL = 0;  // set after load
let SITE_COL = 1;
let DESC_COL = 3;

function safeCell(row, idx){
  const v = (row?.[idx] ?? '').toString().trim();
  return v ? convertTimesInText(v) : '-';
}

/* اختيار أفضل وصف للبطاقة (أبدًا ليس وقت/تاريخ فقط) */
function getCardDescription(row, headers){
  // 1) حاول العمود المكتشف للوصف
  let v = (row?.[DESC_COL] ?? '').toString().trim();

  // 2) إن كان فارغًا أو وقت/تاريخ فقط، جرّب أعمدة وصف محتملة
  const descKeywords = [
    'وصفالحاله الامنيه','وصفالحالهالامنيه','وصفالحالةالامنية','وصفالحالة',
    'الوصف','وصف','تفاصيل','تفاصيلالحالة','تفاصيلالبلاغ','الوصفالتفصيلي',
    'ملخص','ملاحظات','description','details','notes','incident'
  ];

  if (!v || looksLikeDateOrTimeOnly(v)) {
    for (let i=1;i<headers.length;i++){
      if (i===0 || i===DATE_COL || i===SITE_COL) continue;
      const nh = normHeader(headers[i] || '');
      if (descKeywords.some(k => nh.includes(normHeader(k)))){
        const cand = (row?.[i] ?? '').toString().trim();
        if (cand && !looksLikeDateOrTimeOnly(cand)) { v = cand; break; }
      }
    }
  }
  // 3) احتياط نهائي: أول خلية غير فارغة ليست تاريخ/موقع
  if (!v) {
    for (let i=1;i<headers.length;i++){
      if (i===DATE_COL || i===SITE_COL) continue;
      const cand = (row?.[i] ?? '').toString().trim();
      if (cand && !looksLikeDateOrTimeOnly(cand)) { v = cand; break; }
    }
  }
  return v || '-';
}

function buildCardsHTML(values){
  if(!values || values.length <= 1){
    return '<p class="muted">لا توجد سجلات ضمن نطاق الفلاتر الحالي.</p>';
  }
  const headers = values[0];

  let html = '';
  for(let r=1; r<values.length; r++){
    const row = values[r] || [];
    const site = safeCell(row, SITE_COL);
    const descRaw = getCardDescription(row, headers);
    const descHTML = (descRaw === '-') ? '-' : linkifyHTML(convertTimesInText(descRaw));
    const date = formatDateOnly(row[DATE_COL]);

    // تفاصيل: تجاهل عمود 0 (timestamp) + التاريخ + الموقع + الوصف
    let body = '';
    for(let c=1; c<headers.length; c++){
      if(c===DATE_COL || c===SITE_COL || c===DESC_COL) continue;
      const label = headers[c] || `عمود ${c+1}`;
      const val = safeCell(row, c);
      const valHTML = (val === '-') ? '-' : linkifyHTML(val);
      body += `
        <div class="kv-label">${label}</div>
        <div class="kv-value">${valHTML}</div>
      `;
    }

    html += `
      <details class="case-card" dir="rtl">
        <summary>
          <div class="card-title">
            <span class="card-site">${site}</span>
            <span class="card-desc">${descHTML}</span>
            <span class="card-date">التاريخ: ${date}</span>
          </div>
          <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"
               stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <polyline points="18 15 12 9 6 15"></polyline>
          </svg>
        </summary>
        <div class="card-body">
          ${body}
        </div>
      </details>
    `;
  }
  return html;
}

/* ===================== ترتيب الصفوف حسب الشيت (آخر صف أولًا) ===================== */
function reverseRows(values){
  if(!values || values.length <= 2) return values;
  const header = values[0];
  const rows = values.slice(1).reverse();
  return [header, ...rows];
}

/* ===================== واجهة / حالة ===================== */
function uniqueSites(values, siteCol){
  const set = new Set();
  for(let i=1;i<values.length;i++){
    const v = values[i]?.[siteCol];
    if(v && String(v).trim()) set.add(String(v).trim());
  }
  return Array.from(set);
}

let rawCases = []; // [headers, ...rows]

async function loadData(){
  const data = await fetchSheetData(CASES_SHEET_ID, CASES_RANGE);
  rawCases = Array.isArray(data) && data.length ? data : [];

  if(rawCases.length){
    const headers = rawCases[0];

    /* التاريخ: فضّل (التاريخ|date) وتجنّب (timestamp|time|الوقت|الساعة) */
    DATE_COL = findFirst(headers, ['التاريخ','date'], ['timestamp','time','datetime','الوقت','الساعة']);
    if (DATE_COL === -1) DATE_COL = findFirst(headers, ['timestamp','time','datetime','الوقت','الساعة'], []);
    if (DATE_COL === -1) DATE_COL = 0;

    /* الموقع / المركز */
    SITE_COL = findFirst(headers, ['الموقع','site','location','المركز','center','الفرع','branch'], []);
    if (SITE_COL === -1) SITE_COL = 1;

    /* وصف الحالة الأمنية */
    DESC_COL = findFirst(
      headers,
      ['وصف الحالة الأمنية','وصفالحاله الامنيه','وصفالحالهالامنيه','وصفالحالةالامنية','وصفالحالة','وصفالحاله',
       'الوصف','وصف','تفاصيل','تفاصيل الحالة','تفاصيلالبلاغ','الوصف التفصيلي','ملخص','ملاحظات',
       'description','details','notes','incident'],
      ['التاريخ','date','timestamp','time','الوقت','الساعة','الموقع','site','location','المركز','center','الفرع','branch']
    );
    if (DESC_COL === -1 || DESC_COL === DATE_COL || DESC_COL === SITE_COL || DESC_COL === 0){
      for (let i=1;i<headers.length;i++){
        if (i===DATE_COL || i===SITE_COL) continue;
        const nh = normHeader(headers[i]||'');
        if (!/(date|time|timestamp|التاريخ|الوقت|الساعة)/.test(nh)){ DESC_COL = i; break; }
      }
      if (DESC_COL === -1) DESC_COL = 1;
    }
  }

  // تعبئة قائمة المواقع
  const siteSelect = document.getElementById('site');
  if(rawCases.length){
    const sites = uniqueSites(rawCases, SITE_COL);
    siteSelect.innerHTML = `<option value="">الكل</option>` + sites.map(s=>`<option value="${s}">${s}</option>`).join('');
  }else{
    siteSelect.innerHTML = `<option value="">الكل</option>`;
  }
}

function render(view){
  const isMobile = window.innerWidth <= 640;
  const tableWrap = document.getElementById('tableContainer');
  const cardsWrap = document.getElementById('cardsContainer');

  if(isMobile){
    tableWrap.style.display = 'none';
    cardsWrap.style.display = 'grid';
    cardsWrap.innerHTML = buildCardsHTML(view);
  }else{
    cardsWrap.style.display = 'none';
    tableWrap.style.display = 'block';
    tableWrap.innerHTML = buildTableHTML(view);
  }
}

function applyFiltersAndRender(){
  const fromStr = document.getElementById('from').value;
  const toStr   = document.getElementById('to').value;
  const siteVal = document.getElementById('site').value;

  let view = rawCases.slice();
  if(view.length){
    view = filterByDateRange(view, fromStr, toStr, DATE_COL);
    view = filterBySite(view, siteVal, SITE_COL);
    view = reverseRows(view);   // آخر صف بالشيت يظهر أولاً
  }

  render(view);

  const countP = document.getElementById('count');
  const rowsCount = view.length ? Math.max(0, view.length - 1) : 0;
  countP.textContent = rowsCount ? `عدد السجلات: ${rowsCount}` : 'لا توجد سجلات ضمن نطاق الفلاتر الحالي.';
}

/* ===================== تهيئة ===================== */
window.addEventListener('load', async ()=>{
  // آخر 30 يوم افتراضيًا
  const to   = new Date();
  const from = new Date(); from.setDate(to.getDate() - 30);
  const fmt = d => d.toISOString().slice(0,10);
  document.getElementById('to').value   = fmt(to);
  document.getElementById('from').value = fmt(from);

  await loadData();
  applyFiltersAndRender();

  const applyBtn = document.getElementById('apply');
  if (applyBtn) applyBtn.addEventListener('click', applyFiltersAndRender);
  document.getElementById('site').addEventListener('change', applyFiltersAndRender);
  document.getElementById('from').addEventListener('change', applyFiltersAndRender);
  document.getElementById('to').addEventListener('change', applyFiltersAndRender);
});

/* إعادة العرض عند تغيير المقاس/الاتجاه (مع إلغاء الارتداد) */
let resizeTimer;
window.addEventListener('resize', ()=>{
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(applyFiltersAndRender, 120);
});

/* اجعل الروابط داخل <summary> قابلة للنقر بدون تبديل الكارت */
document.addEventListener('click', (e)=>{
  const a = e.target.closest('summary a');
  if(a){ e.stopPropagation(); }
});
