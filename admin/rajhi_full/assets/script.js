/* ===================== البيانات ===================== */
const apiKey = "AIzaSyBYPUAnYE8GC4Vx32cDYSb8UH6YV-VWmEA";

// جلب البيانات من Google Sheets
async function fetchSheetData(sheetId, range){
  try{
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(range)}?key=${apiKey}`;
    const res = await fetch(url);
    const data = await res.json();
    return data.values && data.values.length ? data.values : [["لا توجد بيانات"]];
  }catch(e){
    return [["خطأ في الاتصال بالجدول"]];
  }
}

// محاولة استخراج تاريخ من خلية نصية
function parseDateCell(v){
  const m=(v||'').match(/\d{4}[-\/]\d{2}[-\/]\d{2}|\d{2}[-\/]\d{2}[-\/]\d{4}/);
  return m ? new Date(m[0].replace(/-/g,'/')) : null;
}

// فلترة الصفوف حسب نطاق التاريخ
function filterByDateRange(data, fromDateStr, toDateStr, colIndex=0){
  const f=new Date(fromDateStr), t=new Date(toDateStr);
  if(isNaN(f)||isNaN(t)) return data;
  return data.filter((row,idx)=>{
    if(idx===0) return true; // احتفظ بالرأس
    const d=parseDateCell(row[colIndex]||"");
    return d && d>=f && d<=t;
  });
}

/* ===================== فلترة الموقع ===================== */
// تطبيع مبسط: إزالة مضاعفات المسافات + تحويل أرقام عربية إلى لاتينية + lower
function normalizeSite(s){
  if (s == null) return '';
  const arabicDigits = {'٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'};
  const replaced = String(s).replace(/[٠-٩]/g, d => arabicDigits[d] || d);
  return replaced.trim().replace(/\s+/g,' ').toLowerCase();
}

// فلترة بحسب الموقع بالاعتماد على عمود محدد
function filterBySiteFixedColumn(data, siteName, colIndex){
  if (!siteName) return data;               // 👈 لو اختير "الكل" يرجع كل الصفوف
  if (!data || !data.length) return data;
  const wanted = normalizeSite(siteName);
  return data.filter((row, r) => {
    if (r === 0) return true;               // احتفظ بالهيدر
    const cell = normalizeSite(row[colIndex] ?? '');
    return cell === wanted;
  });
}

/* ===================== تحويل الوقت إلى 12 ساعة (ص/م) ===================== */
// يحوّل أي وقت بالشكل HH:MM(:SS) داخل النص إلى 12 ساعة مع لاحقة ص/م
function convertTimesInText(text){
  if (text == null) return text;
  let s = String(text);

  // إذا كان الوقت مكتوبًا مسبقًا مع AM/PM أو ص/م نتركه كما هو
  const hasAmPm = /\b(AM|PM|am|pm|ص|م)\b/.test(s);
  // نبحث عن كل أجزاء الوقت 24 ساعة بالشكل HH:MM أو HH:MM:SS
  const re = /(\b[01]?\d|2[0-3]):([0-5]\d)(?::([0-5]\d))?/g;

  return s.replace(re, (full, hh, mm, ss) => {
    if (hasAmPm) return full;
    const h = parseInt(hh,10);
    const period = (h < 12) ? ;
    let h12 = h % 12; if (h12 === 0) h12 = 12;
    return `${h12}:${mm}${ss ? ':'+ss : ''} ${period}`;
  });
}

/* ===================== تجميع/ترتيب حسب (التاريخ + الهوية) ===================== */
// كشف تلقائي لعمود الهوية من الهيدر
function detectIdColumn(headers){
  if (!headers || !headers.length) return -1;
  const normalized = headers.map(h => normalizeHeader(String(h||'')));
  const candidates = [
    'رقمالهوية','هوية','id','employeeid','badge','البطاقة','المعرف','رقمالموظف'
  ];
  for (let i=0;i<normalized.length;i++){
    const h = normalized[i];
    if (candidates.some(c => h.includes(c))) return i;
  }
  return -1;
}

function normalizeHeader(s){
  // إزالة مسافات/رموز شائعة + تحويل أرقام عربية + تحويل لحروف صغيرة لاتينية/عربية
  const arabicDigits = {'٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'};
  let t = s.replace(/[٠-٩]/g, d => arabicDigits[d] || d);
  t = t.replace(/\s+|[_\-:/|]+/g,'').toLowerCase();
  return t;
}

// يرتّب الصفوف بحيث تتجاور السجلات ذات نفس (التاريخ + الهوية)
function groupSortByDateThenId(values, dateColIndex=0){
  if (!values || values.length<=2) return values;
  const headers = values[0];
  const rows = values.slice(1);

  const idColIndex = detectIdColumn(headers);
  if (idColIndex === -1) {
    // لو لم نجد عمود الهوية، نرتّب فقط حسب التاريخ للمحاذاة التقريبية
    rows.sort((a,b)=>{
      const da=parseDateCell(a[dateColIndex]||'');
      const db=parseDateCell(b[dateColIndex]||'');
      if (da && db) return da - db;
      if (da && !db) return -1;
      if (!da && db) return 1;
      return 0;
    });
    return [headers, ...rows];
  }

  rows.sort((a,b)=>{
    const da=parseDateCell(a[dateColIndex]||'');
    const db=parseDateCell(b[dateColIndex]||'');
    if (da && db && da - db !== 0) return da - db;     // التاريخ أولاً
    if (da && !db) return -1;
    if (!da && db) return 1;

    const ia=(a[idColIndex]??'').toString();
    const ib=(b[idColIndex]??'').toString();
    // ثم الهوية (مقارنة باللغة العربية/اللاتينية)
    return ia.localeCompare(ib,'ar',{numeric:true,sensitivity:'base'});
  });

  return [headers, ...rows];
}

// نمرر className لعنصر <table> حتى تعمل قواعد الطباعة الخاصة بكل جدول
// نحول القيم الفارغة إلى '-'، ونحوّل الأوقات إلى 12 ساعة
// ونملأ الأعمدة 9 و10 بـ "O" تلقائيًا في جدول الحالات الأمنية
function buildTableHTML(values, className){
  if (!values || !values.length) return '<p>لا توجد بيانات</p>';

  const headers = values[0];
  const cls = className ? ` class="${className}"` : '';

  let h = `<table${cls}><thead><tr>` +
          headers.map(h => `<th>${h}</th>`).join('') +
          `</tr></thead><tbody>`;

  for (let i = 1; i < values.length; i++) {
    // 👈 مهم: نعمل padding للصف لطول العناوين لأن Sheets يحذف الذيل الفارغ
    const row = values[i] || [];
    const rowPadded = Array.from({length: headers.length}, (_, j) => row[j] ?? '');

    h += '<tr>' + rowPadded.map((cell, colIdx) => {
      let val = (cell == null || String(cell).trim() === '') ? '' : String(cell).trim();

      // 👇 خاص بالجدول الأول (الحالات الأمنية): الأعمدة 9 و10 (index 8,9) إذا فاضية → "O"
      if (className === 'cases-table' && (colIdx === 8 || colIdx === 9)) {
        if (!val) val = '-';
      }

      // باقي الأعمدة: لو بقيت فاضية → '-'
      if (!val) return `<td>-</td>`;

      // تحويل أوقات 24h داخل النص إلى 12h (ص/م)
      const shown = convertTimesInText(val);
      return `<td>${shown}</td>`;
    }).join('') + '</tr>';
  }

  h += '</tbody></table>';
  return h;
}


/* ===================== العرض والشاشة ===================== */
const fields=["date","from","to","subject","pages","company","site","coverText","letterBody","thanksLine","footerBlock"];

// تحديث عنوان صفحة المتصفح حسب الفترة والموضوع
function updateDocumentTitleWithDateRange(){
  const f=document.getElementById('from').value;
  const t=document.getElementById('to').value;
  const s=document.getElementById('subject').value||'تقرير أمني';
  document.title=(f&&t)?`${s} - من ${f} إلى ${t}`:s;
}

// تحميل البيانات، بناء معاينات الشاشة، ثم بناء صفحات الطباعة
async function loadAndRender(){
  const fromDate=document.getElementById('from').value;
  const toDate  =document.getElementById('to').value;
  const siteSel =document.getElementById('site').value; // "النسيم 1" / "النسيم 4" / ""

  const sheet1Raw=await fetchSheetData('1k6BSYyEGiezQqubRUDbOFiFy8k34OjZO8ZnNbYi751I','نموذج بلاغات!A2:J');
  const sheet2Raw=await fetchSheetData('1LMeDt4PSeUBUA1IFhBoHSIgnTK-XzLcxIbN4hT6Y7fc','البيانات!A2:L');

  // فلترة التاريخ أولًا
  const dated1=(fromDate&&toDate)?filterByDateRange(sheet1Raw,fromDate,toDate,0):sheet1Raw;
  const dated2=(fromDate&&toDate)?filterByDateRange(sheet2Raw,fromDate,toDate,0):sheet2Raw;

  // ثم فلترة حسب الموقع بأعمدة ثابتة:
  // الجدول الأول (الحالات): العمود الثاني => index 1
  // الجدول الثاني (الحضور): العمود الخامس => index 4
  const bySite1 = filterBySiteFixedColumn(dated1, siteSel, 1);
  // 2.5) ملء قائمة الأسماء حسب البيانات الحالية
  populateEmployeeFilter(bySite1);

  // 2.6) فلترة حسب اسم الموظف (إن اختير)
  const employeeName = document.getElementById('employeeFilter')?.value || '';
  const byEmployee = filterByEmployeeName(bySite1, employeeName);

  const bySite2 = filterBySiteFixedColumn(dated2, siteSel, 4);

  // 👇 الترتيب/التجميع بحيث تتجاور (التاريخ + الهوية)
  const grouped1 = groupSortByDateThenId(bySite1, 0); // التاريخ في العمود 0
  const grouped2 = groupSortByDateThenId(bySite2, 0); // التاريخ في العمود 0

  // معاينة على الشاشة
  document.getElementById('securityCasesPreview').innerHTML = buildTableHTML(grouped1, 'cases-table');
  document.getElementById('attendanceRecordsPreview').innerHTML = buildTableHTML(grouped2, 'attendance-table');

  // بناء صفحات الطباعة الفعلية
  buildPrintPages(grouped1,grouped2);

  // عدّ الصفحات الحقيقية بعد البناء ثم مزامنة العرض
  updatePagesFieldFromDOM();
  syncPrintFields();
}

// تقدير تقريبي (اختياري)
function estimatePagesRough(v1,v2,rowsPerPage){
  const p1=Math.max(1, Math.ceil(Math.max(0,(v1.length-1))/rowsPerPage));
  const p2=Math.max(1, Math.ceil(Math.max(0,(v2.length-1))/rowsPerPage));
  return 2 + p1 + p2; // غلاف + صفحة معلومات + صفحات الجداول
}

/* ===================== عدّ الصفحات وكتابتها في الحقل ===================== */
function updatePagesFieldFromDOM(){
  const root  = document.getElementById('printRoot');
  const field = document.getElementById('pages');
  const span  = document.getElementById('print_pages');

  if (!root || !field) return;

  const count = root.querySelectorAll('.print-page').length;

  field.value = count;
  if (span) span.textContent = count;
}

/* ===================== بناء صفحات الطباعة ===================== */
function buildPrintPages(values1, values2){
  const root=document.getElementById('printRoot');
  root.innerHTML='';

  const coverTextSaved=document.getElementById('coverText').value||'';

  // (1) الغلاف
  const cover=document.createElement('section');
  cover.className='print-page cover-page';

  // نص الغلاف
  const cText=document.createElement('div');
  Object.assign(cText.style,{
    position:'absolute', bottom:'90px', left:'30px', color:'#fff',
    fontSize:'16px', fontWeight:'700', whiteSpace:'pre-wrap', width:'300px'
  });
  cText.textContent=coverTextSaved;
  cover.appendChild(cText);
  root.appendChild(cover);

  // (2) صفحة المعلومات
  const info=document.createElement('section');
  info.className='print-page second-page';
  info.innerHTML=`
    <div class="page-content">
      <div class="header">
         <img src="logo_sanam.png" alt="شعار سنام الأمن" style="height: 85px;" />
         <img src="logo_rajhi.png" alt="شعار الراجحي" />
      </div>
      <h1>تقرير الحالة الأمنية</h1>
      <div style="font-size:14px;line-height:1.8">
        <p><strong>التاريخ:</strong> <span id="print_date"></span></p>
        <p><strong>الفترة:</strong> من <span id="print_from"></span> إلى <span id="print_to"></span></p>
        <p><strong>الموضوع:</strong> <span id="print_subject"></span></p>
        <p><strong>عدد الصفحات:</strong> <span id="print_pages"></span></p>
        <p><strong>اسم العميل:</strong> <span id="print_company"></span></p>
        <p><strong>الموقع:</strong> <span id="print_site"></span></p>
        <div id="print_letterBlock" style="margin-top:12px"></div>
      </div>
    </div>`;
  root.appendChild(info);

  // (3) صفحات الجداول (قياس ديناميكي)
  appendPaginatedTableMeasured(root, values1, 'ملخص الحالات الأمنية', 20, 'cases-table');
  appendPaginatedTableMeasured(root, values2, 'ملخص الحضور والانصراف', 20, 'attendance-table');
}

// بناء صفحات الجداول مع قياس الصفوف التي تلائم الصفحة فعلاً
function appendPaginatedTableMeasured(root, values, title, extraBottomMm=20, className){
  const pages = paginateValuesMeasured(values, title, extraBottomMm);

  if (!pages.length) {
    const empty = document.createElement('section');
    empty.className = 'print-page table-page';
    empty.innerHTML = `<div class="page-content"><h2>${title}</h2><p>لا توجد بيانات.</p></div>`;
    root.appendChild(empty);
    return;
  }

  pages.forEach((vals, idx) => {
    const sec = document.createElement('section');
    sec.className = 'print-page table-page';
    const tableHTML = buildTableHTML(vals, className);
    sec.innerHTML = `<div class="page-content">${idx === 0 ? ('<h2>' + title + '</h2>') : ''}${tableHTML}</div>`;
    root.appendChild(sec);
  });
}

// تقسيم البيانات إلى صفحات حسب الارتفاع المتاح الحقيقي
function paginateValuesMeasured(allValues, title, extraBottomMm = 20){
  if (!allValues || !allValues.length) return [];
  const headers = allValues[0];
  const rows = allValues.slice(1);

  const chunks = [];
  let i = 0;
  let firstPage = true;

  while (i < rows.length) {
    const fit = measureRowsFit(headers, rows.slice(i), firstPage, title, extraBottomMm);
    const rowsThisPage = Math.max(1, fit);
    chunks.push([headers, ...rows.slice(i, i + rowsThisPage)]);
    i += rowsThisPage;
    firstPage = false;
  }
  return chunks;
}

// قياس كم صف يلائم صفحة A4 مع الحواشي
function measureRowsFit(headers, candidateRows, includeTitle, title, extraBottomMm){
  // صفحة قياس مخفية بنفس أبعاد A4 وحواشي داخلية
  const page = document.createElement('div');
  page.style.cssText = `
    position:absolute; left:-10000px; top:0; width:210mm; height:297mm;
    box-sizing:border-box; background:#fff; overflow:hidden;
    font-family:'Cairo', sans-serif;
  `;

  const pagePaddingStr = getComputedStyle(document.documentElement).getPropertyValue('--page-padding').trim() || '20mm';
  const padMm = parseFloat(pagePaddingStr) || 20;
  const padPx = mm2px(padMm);
  const extraPx = mm2px(extraBottomMm);

  const content = document.createElement('div');
  content.style.cssText = `
    box-sizing:border-box; width:100%; height:100%;
    padding:${padPx}px; padding-bottom:${padPx + extraPx}px;
  `;
  page.appendChild(content);

  if (includeTitle) {
    const h2 = document.createElement('h2');
    h2.style.cssText = `margin:8px 0; text-align:center; font-size:16px;`;
    h2.textContent = title || '';
    content.appendChild(h2);
  }

  const table = document.createElement('table');
  table.style.cssText = 'width:100%; border-collapse:collapse; table-layout:fixed;';
  const thead = document.createElement('thead');
  const hr = document.createElement('tr');
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h ?? '';
    th.style.cssText = 'border:1px solid #ccc; padding:6px; font-size:12px; word-break:break-word; white-space:pre-wrap;';
    hr.appendChild(th);
  });
  thead.appendChild(hr);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  table.appendChild(tbody);
  content.appendChild(table);

  document.body.appendChild(page);

  const safeHeight = page.clientHeight;

  let fit = 0;
  for (let r = 0; r < candidateRows.length; r++) {
    const tr = document.createElement('tr');
    candidateRows[r].forEach(cell => {
      const td = document.createElement('td');
      td.textContent = (cell ?? '').toString();
      td.style.cssText = 'border:1px solid #ccc; padding:6px; font-size:12px; word-break:break-word; white-space:pre-wrap;';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);

    if (page.scrollHeight > safeHeight) {
      tbody.removeChild(tr);
      break;
    } else {
      fit++;
    }
  }

  document.body.removeChild(page);
  return fit;
}

// تحويل mm إلى px (على افتراض 96dpi)
function mm2px(mm){ return mm * (96 / 25.4); }

/* ===================== مزامنة الحقول (بما فيها نص الخطاب) ===================== */
function syncPrintFields(){
  // نسخ الحقول البسيطة إلى الصفحة الثانية
  ['date','from','to','subject','pages','company','site'].forEach(id=>{
    const el=document.getElementById(id);
    const span=document.getElementById('print_'+id);
    if(el && span) span.textContent=el.value||'';
  });

  // تركيب نص الخطاب من الحقول الثلاثة: النص الأساسي + الشكر + التذييل
  const out = document.getElementById('print_letterBlock');
  const bodyEl   = document.getElementById('letterBody');
  const thanksEl = document.getElementById('thanksLine');
  const footEl   = document.getElementById('footerBlock');
  if(!out) return;

  const fromVal=document.getElementById('from').value||'____';
  const toVal  =document.getElementById('to').value  ||'____';

  // 1) النص الرئيسي مع استبدال {from}/{to}
  const bodyRaw = (bodyEl?.value || '')
    .replaceAll('{from}', fromVal)
    .replaceAll('{to}', toVal);

  // مساعد لتحويل الأسطر إلى فقرات <p>
  const toParas = (txt, align, bold=false) => {
    return (txt || '')
      .split(/\r?\n/)
      .map(line=>{
        const t=line.trim();
        if(!t) return `<p dir="rtl">&nbsp;</p>`;
        const alignCss  = align ? `text-align:${align};` : '';
        const weightCss = bold ? 'font-weight:700;' : '';
        return `<p dir="rtl" style="${alignCss}${weightCss}">${t}</p>`;
      })
      .join('');
  };

  let html = '';
  // (أ) النص الرئيسي
  html += toParas(bodyRaw, '');

  // (ب) سطر الشكر
  const thanks = (thanksEl?.value || '').trim();
  if(thanks){
    html += `<p dir="rtl" style="text-align:center;">${thanks}</p>`;
  }

  // (ج) التذييل
  const footer = footEl?.value || '';
  if(footer){
    html += toParas(footer, 'left', true);
  }

  out.innerHTML = html;
}

// هروب HTML (احتياط)
function escapeHTML(str){
  return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function nl2br(str){
  return (str||'').split(/\r?\n/).map(line=>line.trim()?line:'\u00A0').join('<br>');
}

/* ===================== إنهاء صفحات الطباعة (تنظيف قبل الطباعة) ===================== */
function finalizePrintPages(){
  const root = document.getElementById('printRoot');
  const pages = Array.from(root.querySelectorAll('.print-page'));

  // (1) تأكيد عدم كسر بعد الصفحة الأخيرة
  if (pages.length) {
    const last = pages[pages.length - 1];
    last.style.breakAfter = 'auto';
    last.style.pageBreakAfter = 'auto';
  }

  // (2) حذف أي صفحة لا تحتوي محتوى فعلي
  pages.forEach(p => {
    const hasContent =
      p.classList.contains('cover-page') ||
      p.classList.contains('second-page') ||
      p.querySelector('table, img, h1, h2, p, .page-content *');
    if (!hasContent) p.remove();
  });
}

/* ===================== التهيئة ===================== */
window.addEventListener('load', ()=>{
  // تحميل القيم المحفوظة ومزامنة فورية
  fields.forEach(id=>{
    const el=document.getElementById(id);
    const saved=localStorage.getItem('report_'+id);
    if(saved && el) el.value=saved;
    if(el){
      el.addEventListener('input', ()=>{
        localStorage.setItem('report_'+id, el.value);
        updateDocumentTitleWithDateRange();
        syncPrintFields();
      });
      if(id==='from' || id==='to'){
        el.addEventListener('change', ()=>{
          loadAndRender(); // سيعيد البناء والحساب
          updateDocumentTitleWithDateRange();
        });
      }
    }
  });

  // إعادة التحميل عند تغيير الموقع
  const siteSel = document.getElementById('site');
  if (siteSel){
    siteSel.addEventListener('change', ()=>{
      loadAndRender();
      syncPrintFields();
    });
  }

  // أول تحميل
  loadAndRender();
  syncPrintFields();
  updateDocumentTitleWithDateRange();
});

/* ===================== الطباعة ===================== */
// زر الطباعة — إعادة بناء + تنظيف + عدّ الصفحات ثم طباعة
document.getElementById('printBtn').addEventListener('click', ()=>{
  buildPrintPages(
    document.querySelector('#securityCasesPreview table') ? tableToValues(document.querySelector('#securityCasesPreview table')) : [],
    document.querySelector('#attendanceRecordsPreview table') ? tableToValues(document.querySelector('#attendanceRecordsPreview table')) : []
  );
  finalizePrintPages();
  updatePagesFieldFromDOM(); // عدّ الصفحات الفعلي بعد التنظيف
  syncPrintFields();         // دفع العدد إلى صفحة المعلومات
  void document.body.offsetHeight; // إعادة تدفق لتفادي صفحات وهمية في كروم
  window.print();
});

// دعم Ctrl+P / System Print — نفس منطق الزر
window.addEventListener('beforeprint', ()=>{
  buildPrintPages(
    document.querySelector('#securityCasesPreview table') ? tableToValues(document.querySelector('#securityCasesPreview table')) : [],
    document.querySelector('#attendanceRecordsPreview table') ? tableToValues(document.querySelector('#attendanceRecordsPreview table')) : []
  );
  finalizePrintPages();
  updatePagesFieldFromDOM();
  syncPrintFields();
  void document.body.offsetHeight;
});

// تحويل جدول المعاينة إلى values (استخدامه قبل الطباعة مباشرة)
function tableToValues(table){
  const values=[];
  const headers=Array.from(table.tHead?.rows?.[0]?.cells||[]).map(th=>th.textContent.trim());
  if(headers.length) values.push(headers);
  const rows=Array.from(table.tBodies?.[0]?.rows||[]);
  rows.forEach(tr=>{
    values.push(Array.from(tr.cells).map(td=>td.textContent));
  });
  return values;
}


/* ===== العثور على عمود الاسم واستخراج الأسماء الفريدة ===== */
function findNameColumn(headers){
  const patterns = [/\bالاسم\b/i,/اسم الموظف/i,/الاسم الرباعي/i];
  for (let i=0;i<headers.length;i++){
    const h = String(headers[i]||'').trim();
    if (patterns.some(re => re.test(h))) return i;
  }
  return 1;
}
function uniqueEmployeeNames(values){
  if (!values || values.length<2) return [];
  const headers = values[0];
  const nameIdx = findNameColumn(headers);
  const set = new Set();
  for (let r=1;r<values.length;r++){
    const name = String((values[r]||[])[nameIdx]||'').trim();
    if (name) set.add(name);
  }
  return Array.from(set).sort((a,b)=>a.localeCompare(b,'ar'));
}
function filterByEmployeeName(values, selectedName){
  if (!selectedName) return values;
  if (!values || values.length<2) return values;
  const headers = values[0];
  const nameIdx = findNameColumn(headers);
  const result = [headers];
  for (let r=1;r<values.length;r++){
    const row = values[r]||[];
    if (String(row[nameIdx]||'').trim() === selectedName) result.push(row);
  }
  return result;
}
function populateEmployeeFilter(values){
  const sel = document.getElementById('employeeFilter');
  if (!sel) return;
  const chosen = sel.value;
  const names = uniqueEmployeeNames(values);
  sel.innerHTML = '<option value="">الكل</option>' + names.map(n=>`<option value="${n}">${n}</option>`).join('');
  if (chosen && names.includes(chosen)) sel.value = chosen;
}
/* ===== إعادة العرض عند تغيير اختيار اسم الموظف ===== */
document.addEventListener('change', (e)=>{
  if (e.target && e.target.id === 'employeeFilter'){
    try { loadAndRender(); } catch(_) {}
    try { loadAndRenderCases && loadAndRenderCases(); } catch(_) {}
  }
});

