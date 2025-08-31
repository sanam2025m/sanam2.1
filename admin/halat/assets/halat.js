/* ===== إعدادات عامة ===== */
const apiKey = "AIzaSyBYPUAnYE8GC4Vx32cDYSb8UH6YV-VWmEA";
const fields = ["date","from","to","subject","pages","company","site","coverText","letterBody","thanksLine","footerBlock"];

/* ===== جلب بيانات Google Sheets (الحالات الأمنية) ===== */
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

/* ===== أدوات التاريخ ===== */
function parseDateCell(v){
  const m=(v||'').match(/\d{4}[-\/]\d{2}[-\/]\d{2}|\d{2}[-\/]\d{2}[-\/]\d{4}/);
  return m ? new Date(m[0].replace(/-/g,'/')) : null;
}
function filterByDateRange(data, fromDateStr, toDateStr, colIndex=0){
  const f=new Date(fromDateStr), t=new Date(toDateStr);
  if(isNaN(f)||isNaN(t)) return data;
  return data.filter((row,idx)=>{
    if(idx===0) return true;
    const d=parseDateCell(row[colIndex]||"");
    return d && d>=f && d<=t;
  });
}

/* ===== فلترة الموقع ===== */
// تطبيع النص + تحويل أرقام عربية
function normalizeSite(s){
  if (s == null) return '';
  const arabicDigits = {'٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'};
  const replaced = String(s).replace(/[٠-٩]/g, d => arabicDigits[d] || d);
  return replaced.trim().replace(/\s+/g,' ').toLowerCase();
}
// فلترة حسب عمود الموقع (ثابت على العمود الثاني B => index 1)
function filterBySiteFixedColumn(data, siteName, colIndex=1){
  if (!siteName) return data;               // لو الحقل فاضي → لا فلترة
  if (!data || !data.length) return data;
  const wanted = normalizeSite(siteName);
  return data.filter((row, r) => {
    if (r === 0) return true;               // احتفظ بالهيدر
    const cell = normalizeSite(row[colIndex] ?? '');
    return cell === wanted;
  });
}

/* ===== تحويل الوقت إلى 12 ساعة (ص/م) داخل النصوص ===== */
function convertTimesInText(text){
  if (text == null) return text;
  let s = String(text);
  // لو السطر يحوي AM/PM أو ص/م مسبقًا نتركه
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

/* ===== بناء HTML للجداول + ميزات (O للأعمدة 9/10) ===== */
function buildTableHTML(values, className){
  if (!values || !values.length) return '<p>لا توجد بيانات</p>';

  const headers = values[0];
  const cls = className ? ` class="${className}"` : '';
  let h = `<table${cls}><thead><tr>` +
          headers.map(h => `<th>${h}</th>`).join('') +
          `</tr></thead><tbody>`;

  for (let i = 1; i < values.length; i++) {
    // نعمل padding للصف لطول العناوين لأن Sheets يحذف الذيل الفارغ
    const row = values[i] || [];
    const rowPadded = Array.from({length: headers.length}, (_, j) => row[j] ?? '');

    h += '<tr>' + rowPadded.map((cell, colIdx) => {
      let val = (cell == null || String(cell).trim() === '') ? '' : String(cell).trim();

      // خاص بجدول الحالات الأمنية: الأعمدة 9 و10 (index 8,9) إذا فاضية → "O"
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

/* ===== عنوان الصفحة ===== */
function updateDocumentTitleWithDateRange(){
  const f=document.getElementById('from').value;
  const t=document.getElementById('to').value;
  const s=document.getElementById('subject').value||'تقرير الحالات الأمنية';
  document.title=(f&&t)?`${s} - من ${f} إلى ${t}`:s;
}

/* ===== التحميل والعرض ===== */
async function loadAndRender(){
  const fromDate=document.getElementById('from').value;
  const toDate  =document.getElementById('to').value;
  const siteVal = document.getElementById('site').value; // إدخال حر

  const raw = await fetchSheetData('1k6BSYyEGiezQqubRUDbOFiFy8k34OjZO8ZnNbYi751I','نموذج بلاغات!A2:J');

  // 1) فلترة تاريخ
  const dated = (fromDate&&toDate)?filterByDateRange(raw,fromDate,toDate,0):raw;

  // 2) فلترة موقع (عمود B = index 1، عدّل الرقم لو عمود الموقع مختلف)
  const bySite = filterBySiteFixedColumn(dated, siteVal, 1);
  // 2.5) ملء قائمة الأسماء حسب البيانات الحالية
  populateEmployeeFilter(bySite);

  // 2.6) فلترة حسب اسم الموظف (إن اختير)
  const employeeName = document.getElementById('employeeFilter')?.value || '';
  const byEmployee = filterByEmployeeName(bySite, employeeName);


  // 3) العرض
  document.getElementById('securityCasesPreview').innerHTML = buildTableHTML(byEmployee, 'cases-table');

  // 4) الطباعة
  buildPrintPages(byEmployee);
  updatePagesFieldFromDOM();
  syncPrintFields();
}

/* ===== عدّ الصفحات ===== */
function updatePagesFieldFromDOM(){
  const root  = document.getElementById('printRoot');
  const field = document.getElementById('pages');
  const span  = document.getElementById('print_pages');
  if (!root || !field) return;
  const count = root.querySelectorAll('.print-page').length;
  field.value = count; if (span) span.textContent = count;
}

/* ===== بناء صفحات الطباعة (صفحات الغلاف+المعلومات+جدول واحد) ===== */
function buildPrintPages(values){
  const root=document.getElementById('printRoot');
  root.innerHTML='';

  const coverTextSaved=document.getElementById('coverText').value||'';

  // الغلاف
  const cover=document.createElement('section');
  cover.className='print-page cover-page';
  const cText=document.createElement('div');
  Object.assign(cText.style,{position:'absolute', bottom:'90px', left:'30px', color:'#fff', fontSize:'16px', fontWeight:'700', whiteSpace:'pre-wrap', width:'300px'});
  cText.textContent=coverTextSaved; cover.appendChild(cText); root.appendChild(cover);

  // صفحة المعلومات
  const info=document.createElement('section');
  info.className='print-page second-page';
  info.innerHTML=`
    <div class="page-content">
      <div class="header">
         <img src="assets/logo_sanam.png" alt="شعار سنام الأمن" style="height: 85px;" />
         <img src="assets/logo_rajhi.png" alt="شعار الراجحي" />
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

  // صفحات الجدول
  appendPaginatedTableMeasured(root, values, 'ملخص الحالات الأمنية', 20, 'cases-table');
}

/* ===== تقسيم الجدول إلى صفحات بطباعة مقاسة ===== */
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

function paginateValuesMeasured(allValues, title, extraBottomMm = 20){
  if (!allValues || !allValues.length) return [];
  const headers = allValues[0];
  const rows = allValues.slice(1);

  const chunks = [];
  let i = 0, firstPage = true;

  while (i < rows.length) {
    const fit = measureRowsFit(headers, rows.slice(i), firstPage, title, extraBottomMm);
    const rowsThisPage = Math.max(1, fit);
    chunks.push([headers, ...rows.slice(i, i + rowsThisPage)]);
    i += rowsThisPage;
    firstPage = false;
  }
  return chunks;
}

function measureRowsFit(headers, candidateRows, includeTitle, title, extraBottomMm){
  const page = document.createElement('div');
  page.style.cssText = `position:absolute; left:-10000px; top:0; width:210mm; height:297mm; box-sizing:border-box; background:#fff; overflow:hidden; font-family:'Cairo', sans-serif;`;

  const pagePaddingStr = getComputedStyle(document.documentElement).getPropertyValue('--page-padding').trim() || '20mm';
  const padMm = parseFloat(pagePaddingStr) || 20;
  const padPx = mm2px(padMm);
  const extraPx = mm2px(extraBottomMm);

  const content = document.createElement('div');
  content.style.cssText = `box-sizing:border-box; width:100%; height:100%; padding:${padPx}px; padding-bottom:${padPx + extraPx}px;`;
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
  thead.appendChild(hr); table.appendChild(thead);
  const tbody = document.createElement('tbody'); table.appendChild(tbody);
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

    if (page.scrollHeight > safeHeight) { tbody.removeChild(tr); break; } else { fit++; }
  }

  document.body.removeChild(page);
  return fit;
}

function mm2px(mm){ return mm * (96 / 25.4); }

/* ===== مزامنة الحقول ===== */
function syncPrintFields(){
  ['date','from','to','subject','pages','company','site'].forEach(id=>{
    const el=document.getElementById(id);
    const span=document.getElementById('print_'+id);
    if(el && span) span.textContent=el.value||'';
  });

  const out = document.getElementById('print_letterBlock');
  const bodyEl   = document.getElementById('letterBody');
  const thanksEl = document.getElementById('thanksLine');
  const footEl   = document.getElementById('footerBlock');
  if(!out) return;

  const fromVal=document.getElementById('from').value||'____';
  const toVal  =document.getElementById('to').value  ||'____';

  const bodyRaw = (bodyEl?.value || '').replaceAll('{from}', fromVal).replaceAll('{to}', toVal);

  const toParas = (txt, align, bold=false) =>
    (txt || '').split(/\r?\n/).map(line=>{
      const t=line.trim(); if(!t) return `<p dir="rtl">&nbsp;</p>`;
      const alignCss  = align ? `text-align:${align};` : '';
      const weightCss = bold ? 'font-weight:700;' : '';
      return `<p dir="rtl" style="${alignCss}${weightCss}">${t}</p>`;
    }).join('');

  let html = '';
  html += toParas(bodyRaw, '');
  const thanks = (thanksEl?.value || '').trim(); if(thanks){ html += `<p dir="rtl" style="text-align:center;">${thanks}</p>`; }
  const footer = footEl?.value || ''; if(footer){ html += toParas(footer, 'left', true); }
  out.innerHTML = html;
}

/* ===== تهيئة ===== */
window.addEventListener('load', ()=>{
  fields.forEach(id=>{
    const el=document.getElementById(id);
    const saved=localStorage.getItem('report_cases_'+id);
    if(saved && el) el.value=saved;
    if(el){
      el.addEventListener('input', ()=>{
        localStorage.setItem('report_cases_'+id, el.value);
        updateDocumentTitleWithDateRange();
        syncPrintFields();
      });
      if(id==='from' || id==='to'){
        el.addEventListener('change', ()=>{
          loadAndRender();
          updateDocumentTitleWithDateRange();
        });
      }
      if(id==='site'){
        el.addEventListener('input', ()=>{ // فلترة مباشرة عند تعديل الموقع
          loadAndRender();
          syncPrintFields();
        });
      }
    }
  });

  loadAndRender();
  syncPrintFields();
  updateDocumentTitleWithDateRange();
});

/* ===== طباعة ===== */
document.getElementById('printBtn').addEventListener('click', ()=>{
  buildPrintPages(
    document.querySelector('#securityCasesPreview table') ? tableToValues(document.querySelector('#securityCasesPreview table')) : []
  );
  finalizePrintPages();
  updatePagesFieldFromDOM();
  syncPrintFields();
  void document.body.offsetHeight;
  window.print();
});

window.addEventListener('beforeprint', ()=>{
  buildPrintPages(
    document.querySelector('#securityCasesPreview table') ? tableToValues(document.querySelector('#securityCasesPreview table')) : []
  );
  finalizePrintPages();
  updatePagesFieldFromDOM();
  syncPrintFields();
  void document.body.offsetHeight;
});

/* ===== تحويل جدول المعاينة إلى values ===== */
function tableToValues(table){
  const values=[];
  const headers=Array.from(table.tHead?.rows?.[0]?.cells||[]).map(th=>th.textContent.trim());
  if(headers.length) values.push(headers);
  const rows=Array.from(table.tBodies?.[0]?.rows||[]);
  rows.forEach(tr=>{ values.push(Array.from(tr.cells).map(td=>td.textContent)); });
  return values;
}

/* ===== تنظيف صفحات الطباعة ===== */
function finalizePrintPages(){
  const root = document.getElementById('printRoot');
  const pages = Array.from(root.querySelectorAll('.print-page'));
  if (pages.length) { const last = pages[pages.length - 1]; last.style.breakAfter = 'auto'; last.style.pageBreakAfter = 'auto'; }
  pages.forEach(p => {
    const hasContent = p.classList.contains('cover-page') || p.classList.contains('second-page') || p.querySelector('table, img, h1, h2, p, .page-content *');
    if (!hasContent) p.remove();
  });
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
