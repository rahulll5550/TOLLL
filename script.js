document.addEventListener('DOMContentLoaded', () => {
  const $ = id => document.getElementById(id);
  const excelInput = $('excelInput');
  const btnSampleData = $('btnSampleData'), btnClearInput = $('btnClearInput');
  const btnProcess = $('btnProcess'), btnClearAll = $('btnClearAll');
  const btnCopyExcel = $('btnCopyExcel'), btnDownloadExcel = $('btnDownloadExcel');
  const chkMergeTrainLoco = $('chkMergeTrainLoco'), lblMergeTrainLoco = $('lblMergeTrainLoco');
  const chkCleanPrefix = $('chkCleanPrefix'), chkTrimText = $('chkTrimText'), chkPreserveBlank = $('chkPreserveBlank');
  const colSelect = $('colSelect'), patternSelect = $('patternSelect');
  const tableSearch = $('tableSearch'), tableInfo = $('tableInfo'), formatHint = $('formatHint'), pasteSubtext = $('pasteSubtext');
  const emptyState = $('emptyState'), tableWrap = $('tableWrap'), tableHead = $('tableHead'), tableBody = $('tableBody'), toastContainer = $('toastContainer');
  const metricInputRows = $('metricInputRows'), metricOutputRows = $('metricOutputRows'), metricSplitCount = $('metricSplitCount'), metricRatio = $('metricRatio');

  const dashboardSection = $('dashboardSection');
  const dashboardTableBody = $('dashboardTableBody');
  const btnCopyDashboard = $('btnCopyDashboard');

  const MAIN_SOURCE_HEADERS = ['SR NO.','Date','Section','Train No','Loco No','LOCO FIT / UNFIT','LOCO OEM','UP / DN','Reason','Failure type','Station','Failure Description'];
  const NMS_HEADERS = ['SR NO.','DATE','STATION NAME','LOCO NO','Train Number','LOCO OEM','Direction (R/N)','Time of Fault','Fault Type','NMS Observations','Preliminary Cause Analysis as per NMS Data'];
  const POST_COMM_HEADERS = ['Section','Date','Train No/Loco No','Remarks'];

  const FORMATS = {
    postCommissioning: {
      name: 'Post Commissioning',
      headers: POST_COMM_HEADERS,
      sourceHeaders: ['Date', 'Section', 'Train No', 'Loco No', 'Remarks'],
      issueRegex: /remarks?|issues?|description|details|fault|defect/i,
      issueDefault: 4,
      hint: 'Post Commissioning format: converts input into Section, Date, Train No/Loco No and Remarks; merges Train & Loco numbers, and splits numbered entries.'
    },
    nmsData: {
      name: 'NMS Data Format',
      headers: NMS_HEADERS,
      sourceHeaders: MAIN_SOURCE_HEADERS,
      issueRegex: /failure\s*description|description/i,
      issueDefault: 11,
      hint: 'NMS Data format: converts Main Data sheet into 11-column NMS report; splits numbered observations and matches Station, Fault Type & Cause.'
    }
  };

  let currentFormat = 'postCommissioning';
  let rawRows = [], sourceHeaders = [], headerless = false, processedData = [], filteredData = [], isCopying = false, isExporting = false;

  function norm(v){ return String(v ?? '').trim().toLowerCase().replace(/\s+/g,' '); }
  function isDateLike(v){ return /^(?:\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}|\d{4}[-\/]\d{1,2}[-\/]\d{1,2})$/.test(String(v ?? '').trim()); }

  function looksLikeHeader(row){
    const joined = row.map(norm).join(' | ');
    if (currentFormat === 'nmsData') {
      return /\bsr\s*no\b/.test(joined) && /date/.test(joined) && /failure\s*description/.test(joined);
    }
    return /\bdate\b/.test(joined) && (/section/.test(joined) || /train/.test(joined) || /loco/.test(joined) || /remark/.test(joined));
  }

  function detectInputLayout(rows) {
    if (!rows.length) return { hasHeader: false, order: 'date-section-train-loco-remarks' };
    const first = rows[0];
    if (looksLikeHeader(first)) return { hasHeader: true };

    if (first.length >= 5 && isDateLike(first[0]) && !isDateLike(first[1])) {
      return { hasHeader: false, order: 'date-section-train-loco-remarks' };
    }
    if (first.length >= 5 && !isDateLike(first[0]) && isDateLike(first[1])) {
      return { hasHeader: false, order: 'section-date-train-loco-remarks' };
    }
    return { hasHeader: false, order: 'date-section-train-loco-remarks' };
  }

  function parseTSV(text){
    const rows=[]; let row=[], cell='', quoted=false;
    const input=String(text??'').replace(/\r\n/g,'\n').replace(/\r/g,'\n');
    for(let i=0;i<input.length;i++){
      const ch=input[i];
      if(ch==='"'){
        if(quoted && input[i+1]==='"'){cell+='"';i++;} else quoted=!quoted;
      } else if(ch==='\t'&&!quoted){row.push(cell);cell='';}
      else if(ch==='\n'&&!quoted){row.push(cell);cell='';if(row.some(v=>String(v).trim()!==''))rows.push(row);row=[];}
      else cell+=ch;
    }
    row.push(cell); if(row.some(v=>String(v).trim()!==''))rows.push(row); return rows;
  }

  function updateColumnDropdown(headers){
    const old=colSelect.value;
    colSelect.innerHTML='<option value="__AUTO__">Auto-Detect Issue Column</option><option value="__NONE__">No Issue Column (Do not split)</option>';
    headers.forEach((h,i)=>{const o=document.createElement('option');o.value=String(i);o.textContent=`${h||`Column ${i+1}`} (Col ${i+1})`;colSelect.appendChild(o);});
    if([...colSelect.options].some(o=>o.value===old))colSelect.value=old;
  }

  function parseInputText(){
    const text=excelInput.value.trim();
    if(!text){rawRows=[];sourceHeaders=[];headerless=false;updateColumnDropdown([]);return;}
    const rows=parseTSV(text), layout=detectInputLayout(rows);
    if(layout.hasHeader){
      headerless=false; sourceHeaders=rows[0].map((v,i)=>String(v??'').trim()||`Column ${i+1}`); rawRows=rows.slice(1);
    }else{
      headerless=true; sourceHeaders=FORMATS[currentFormat].sourceHeaders; rawRows=rows;
    }
    updateColumnDropdown(sourceHeaders);
    return layout;
  }

  function detectIssueColumn(headers,rows){
    for(let i=0;i<headers.length;i++) if(FORMATS[currentFormat].issueRegex.test(headers[i])) return i;
    let best=-1,score=0;
    rows.forEach(r=>r.forEach((cell,i)=>{
      const count=(String(cell??'').match(/(?:^|[\n]|[.!?;:]\s)\d{1,3}(?:\.|\)|\s*-\s+)\s+/g)||[]).length;
      if(count>score){score=count;best=i;}
    }));
    return best>=0?best:Math.max(0, headers.length - 1);
  }

  function cleanRemark(v){return String(v??'').replace(/[ \t]*\n[ \t]*/g,' ').replace(/[ \t]{2,}/g,' ').trim();}

  function splitIssues(text,pattern,removeNumbering){
    const value=String(text??'').replace(/\r\n/g,'\n').replace(/\r/g,'\n').trim();
    if(!value)return [''];

    let re;
    if(pattern==='dot') re=/(?:^|\n|\s)(\d{1,3})\.(?=\s+|[A-Za-z])/g;
    else if(pattern==='parenthesis') re=/(?:^|\n|\s)(\d{1,3})\)(?=\s+|[A-Za-z])/g;
    else if(pattern==='dash') re=/(?:^|\n|\s)(\d{1,3})\s*-\s+/g;
    else re=/(?:^|\n|\s)(\d{1,3})(?:\.|\))(?=\s+|[A-Za-z])|(?:^|\n|\s)(\d{1,3})\s*-\s+/g;

    const candidates=[]; let m;
    while((m=re.exec(value))!==null){
      const numStr = m[1] || m[2];
      const num = parseInt(numStr, 10);

      if(isNaN(num) || num > 30 || num < 1) continue;

      const fullMatch = m[0];
      const numPos = m.index + fullMatch.indexOf(numStr);
      const beforeText = value.slice(0, numPos);

      const openCount = (beforeText.match(/\(/g) || []).length;
      const closeCount = (beforeText.match(/\)/g) || []).length;
      if(openCount > closeCount) continue;

      const preceding15 = beforeText.slice(-15);
      if(/\b(?:Tag|R|Location|Loco|Speed|PSR|Abs)\s*[-:]?\s*$/i.test(preceding15)) continue;

      candidates.push({start:numPos, number:num});
    }

    if(!candidates.length)return [cleanRemark(value)];

    const result=[];
    for(let i=0;i<candidates.length;i++){
      const start=candidates[i].start;
      const end=i+1<candidates.length?candidates[i+1].start:value.length;
      let item=value.slice(start,end).trim();
      if(removeNumbering)item=item.replace(/^\d{1,3}(?:\.|\)|\s*-\s+)\s*/,'').trim();
      item=cleanRemark(item);
      if(item)result.push(item);
    }
    return result.length?result:[cleanRemark(value)];
  }

  function headerIndex(headers,re){return headers.findIndex(h=>re.test(norm(h)));}

  function getHeaderIndexes(headers) {
    const find = re => headers.findIndex(h => re.test(norm(h)));
    return {
      section: find(/^section$/),
      date: find(/^date$/),
      train: find(/^train(?:\s*no|\s*number|\s*#)?$/),
      loco: find(/^loco(?:\s*no|\s*number|\s*#)?$/),
      combined: find(/^train\s*no\s*\/\s*loco\s*no$/),
      remarks: find(/^(?:remarks?|issues?|description|details|fault|defect)$/)
    };
  }

  function makePostCommissioning(row, rowIndex, layout, issueColumn) {
    let date = '', section = '', train = '', loco = '', remarks = '';

    if (headerless) {
      if (layout && layout.order === 'section-date-train-loco-remarks') {
        [section, date, train, loco, remarks] = row.map(v => String(v ?? ''));
      } else {
        [date, section, train, loco, remarks] = row.map(v => String(v ?? ''));
      }
    } else {
      const idx = getHeaderIndexes(sourceHeaders);
      date = idx.date >= 0 ? String(row[idx.date] ?? '') : '';
      section = idx.section >= 0 ? String(row[idx.section] ?? '') : '';
      train = idx.train >= 0 ? String(row[idx.train] ?? '') : '';
      loco = idx.loco >= 0 ? String(row[idx.loco] ?? '') : '';
      if (idx.combined >= 0) {
        const combined = String(row[idx.combined] ?? '').trim();
        if (!train && !loco) { train = combined; loco = ''; }
      }
      remarks = issueColumn >= 0 ? String(row[issueColumn] ?? '') : '';
    }

    if (!date && row.length >= 5 && isDateLike(row[0])) date = String(row[0]);
    if (!section && row.length >= 5 && !isDateLike(row[0]) && isDateLike(row[1])) section = String(row[0]);
    if (!train && row.length >= 5) train = String(headerless ? row[2] : train);
    if (!loco && row.length >= 5) loco = String(headerless ? row[3] : loco);
    if (!remarks && issueColumn >= 0) remarks = String(row[issueColumn] ?? '');

    if (chkTrimText.checked) {
      date = date.trim(); section = section.trim(); train = train.trim(); loco = loco.trim(); remarks = remarks.trim();
    }

    const mergeCheck = chkMergeTrainLoco ? chkMergeTrainLoco.checked : true;
    const trainLoco = mergeCheck ? (train && loco ? `${train} / ${loco}` : (train || loco)) : train;
    return {
      Section: section,
      Date: date,
      'Train No/Loco No': trainLoco,
      Remarks: remarks,
      __origRowIndex: rowIndex
    };
  }

  function extractTime(text){
    const m=String(text??'').match(/\b(?:[01]?\d|2[0-3]):[0-5]\d(?::[0-5]\d)?\b/);
    return m ? m[0] : '';
  }

  function cleanLeadingNumber(text){
    return String(text??'').replace(/^\s*\d{1,3}\s*[.)-]\s*/,'').trim();
  }

  function cleanNumberedValue(text){
    let v=String(text??'').trim();
    v=v.replace(/^\s*\d{1,3}\s*[.)]\s*/,'');
    v=v.replace(/\s+\d{1,3}\s*[.)]\s*$/,'');
    return v.trim();
  }

  function parseNumberedField(text){
    const raw=String(text??'').replace(/\r\n/g,'\n').replace(/\r/g,'\n').trim();
    if(!raw) return [];
    let re=/(?:^|\n)\s*(\d{1,3})\s*[.)]\s*/g;
    let marks=[],m;
    while((m=re.exec(raw))!==null) marks.push({start:m.index+m[0].length,n:String(m[1])});
    if(marks.length<2){
      re=/(?:^|\s)(\d{1,3})\s*[.)](?=\s+|[A-Za-z])/g;
      marks=[];
      while((m=re.exec(raw))!==null) marks.push({start:m.index+m[0].length,n:String(m[1])});
    }
    if(!marks.length) return [{n:'1',v:cleanLeadingNumber(cleanRemark(raw))}];
    const out=[];
    for(let i=0;i<marks.length;i++){
      const end=i+1<marks.length?marks[i+1].start:raw.length;
      const v=cleanRemark(raw.slice(marks[i].start,end));
      if(v) out.push({n:marks[i].n,v:cleanLeadingNumber(v)});
    }
    return out;
  }

  function getIssueItems(text){
    return parseNumberedField(text).map(x=>({n:String(x.n),v:cleanNumberedValue(x.v)}));
  }

  function normalizeLocoFaultType(text){
    const v=cleanNumberedValue(text);
    if(/^FS\s*-\s*SR$/i.test(v)) return 'SR';
    if(/^FS\s*-\s*LS$/i.test(v)) return 'LS';
    return v;
  }

  function getDescriptionIssueItems(text, allowedNumbers){
    const raw=String(text??'').replace(/\r\n/g,'\n').replace(/\r/g,'\n').trim();
    if(!raw) return [];

    const allowed = allowedNumbers && allowedNumbers.size
      ? new Set([...allowedNumbers].map(String))
      : null;

    const marks=[];
    const seen=new Set();
    const addMark=(start,n)=>{
      n=String(n);
      if(allowed && !allowed.has(n)) return;
      const key=`${start}|${n}`;
      if(!seen.has(key)){seen.add(key);marks.push({start,n});}
    };

    let re=/(?:^|\n)[ \t]*(\d{1,3})\s*[.)][ \t]*/g, m;
    while((m=re.exec(raw))!==null) addMark(m.index+m[0].length,String(m[1]));

    re=/(?:^|[ \t])(\d{1,3})\s*\.[ \t]*/g;
    while((m=re.exec(raw))!==null) addMark(m.index+m[0].length,String(m[1]));

    marks.sort((a,b)=>a.start-b.start);

    if(!marks.length){
      let v=raw.replace(/^\s*In\s+Train\s+No?\.?\s*[^\n]*?(?:,|;|\n)\s*/i,'').trim();
      return [{n:'1',v:cleanNumberedValue(v)}];
    }

    const out=[];
    for(let i=0;i<marks.length;i++){
      const start=marks[i].start;
      const end=i+1<marks.length ? marks[i+1].start : raw.length;
      let v=raw.slice(start,end).trim();
      v=cleanRemark(v);
      if(v) out.push({n:marks[i].n,v});
    }
    return out;
  }

  function getMatchingDescription(text, issueNumber, allowedNumbers, fallback=''){
    const raw=String(text??'').replace(/\r\n/g,'\n').replace(/\r/g,'\n').trim();
    if(!raw) return '';

    const target=String(issueNumber);
    const known = allowedNumbers && allowedNumbers.size
      ? new Set([...allowedNumbers].map(String)) : null;
    const markerRe=/(?:^|[\n\t ]+)(\d{1,3})\s*\.(?=\s|[A-Za-z])/g;
    const marks=[]; let m;
    while((m=markerRe.exec(raw))!==null){
      const n=String(m[1]);
      if(known && !known.has(n)) continue;
      const full=m[0];
      const dotPos=m.index + full.lastIndexOf('.') + 1;
      marks.push({n,start:dotPos});
    }

    if(marks.length){
      marks.sort((a,b)=>a.start-b.start);
      const hitIndex=marks.findIndex(x=>x.n===target);
      if(hitIndex>=0){
        const next=hitIndex+1<marks.length ? marks[hitIndex+1].start : raw.length;
        let value=raw.slice(marks[hitIndex].start,next).trim();
        value=value.replace(/^\.\s*/,'').trim();
        value=value.replace(/^\s*\d{1,3}\s*\.\s*/,'').trim();
        return cleanRemark(value);
      }
    }

    if(target==='1' && !/(?:^|[\n\t ])\d{1,3}\s*\.(?=\s|[A-Za-z])/.test(raw)){
      const value=raw.replace(/^\s*In\s+Train\s+No?\.?\s+[^,;\n]+[,;]?\s*/i,'').trim();
      return cleanRemark(value);
    }

    const directRe=new RegExp('(?:^|[\\n\\t ])'+target.replace(/[.*+?^${}()|[\\]\\\\]/g,'\\$&')+'\\s*\\.(?=\\s|[A-Za-z])','g');
    const direct=directRe.exec(raw);
    if(direct){
      const startPos=direct.index + direct[0].lastIndexOf('.') + 1;
      const rest=raw.slice(startPos).trim();
      return cleanRemark(rest.replace(/^\d{1,3}\s*\.\s*/,'').trim());
    }
    return String(fallback??'').trim();
  }

  function getIssueValue(text, issueNumber, issueIndex, fallback=''){
    const raw=String(text??'').trim();
    if(!raw) return '';

    const items=getIssueItems(raw);
    if(!items.length) return '';

    const hasExplicitNumbering=/^\s*\d{1,3}\s*[.)]\s*/.test(raw) ||
      /(?:^|\n|\s)\d{1,3}\s*[.)](?=\s+|[A-Za-z])/.test(raw);

    if(!hasExplicitNumbering){
      return String(issueNumber)==='1' ? cleanNumberedValue(raw) : '';
    }

    const hit=items.find(x=>x.n===String(issueNumber));
    return hit ? hit.v : '';
  }

  function makeNMS(row,rowIndex,issueColumn){
    const get = re => { const i=headerIndex(sourceHeaders,re); return i>=0?String(row[i]??''):''; };
    let sr='',date='',station='',loco='',train='',oem='',direction='',reason='',failureType='',desc='';
    if(headerless){
      const first=String(row[0]??'').trim(), second=String(row[1]??'').trim();
      const srMissing=isDateLike(first)&&!isDateLike(second), off=srMissing?-1:0;
      const at=i=>String(row[i+off]??'');
      sr=srMissing?'':String(row[0]??''); date=srMissing?String(row[0]??''):String(row[1]??'');
      train=at(3); loco=at(4); oem=at(6); direction=at(7); reason=at(8); failureType=at(9); station=at(10); desc=at(11);
    } else {
      sr=get(/^sr\s*no/); date=get(/^date$/); station=get(/^station$/); loco=get(/^loco\s*no|^loco\s*number|^loco\s*#$/); train=get(/^train\s*no|^train\s*number|^train\s*#$/); oem=get(/^loco\s*oem$/); direction=get(/^up\s*\/\s*dn$|^direction/); reason=get(/^reason$/); failureType=get(/^failure\s*type$/); desc=get(/failure\s*description|^description$/);
      if(issueColumn>=0) desc=row.slice(issueColumn).map(v=>String(v??'')).join('\t');
    }
    return { 'SR NO.':sr,'DATE':date,'STATION NAME':cleanNumberedValue(station),'LOCO NO':loco,'Train Number':train,'LOCO OEM':oem,'Direction (R/N)':cleanNumberedValue(direction),'Time of Fault':'','Fault Type':'','NMS Observations':'','Preliminary Cause Analysis as per NMS Data':'', __origRowIndex:rowIndex, __failureTypeRaw:failureType,__reasonRaw:reason,__stationRaw:station,__descriptionRaw:desc };
  }

  function generateDashboardData() {
    if (!processedData.length) {
      dashboardSection.classList.add('hidden');
      return;
    }

    const byDate = {};
    const allUniqueLocos = new Set();
    let totalFSB = 0, totalEB = 0, totalSR = 0, totalLS = 0;

    processedData.forEach(row => {
      const rawDate = String(row.Date || row.DATE || '').trim() || 'Unspecified';

      let locoId = '';
      if (row['Train No/Loco No']) {
        const parts = row['Train No/Loco No'].split('/');
        locoId = (parts[1] || parts[0] || '').trim();
      } else {
        locoId = String(row['LOCO NO'] || row['LOCO ID'] || row['Loco No'] || '').trim();
      }
      if (locoId) allUniqueLocos.add(locoId);

      const combinedText = [
        row.Remarks,
        row['NMS Observations'],
        row['Fault Type'],
        row['Loco Fault'],
        row['Preliminary Cause Analysis as per NMS Data'],
        row['Preliminary Cause']
      ].map(v => String(v ?? '')).join(' ').toUpperCase();

      const isFSB = /\bFSB\b|FULL\s*SERVICE\s*BRAKE/.test(combinedText);
      const isEB = /\bEB\b|EMERGENCY\s*BRAKE/.test(combinedText);
      const isSR = /\bSR\b|FS\s*-\s*SR|SPEED\s*RESTRICTION/.test(combinedText);
      const isLS = /\bLS\b|FS\s*-\s*LS|LOCO\s*STOP/.test(combinedText);

      if (!byDate[rawDate]) {
        byDate[rawDate] = { locos: new Set(), fsb: 0, eb: 0, sr: 0, ls: 0, total: 0 };
      }

      if (locoId) byDate[rawDate].locos.add(locoId);
      if (isFSB) { byDate[rawDate].fsb++; totalFSB++; }
      if (isEB) { byDate[rawDate].eb++; totalEB++; }
      if (isSR) { byDate[rawDate].sr++; totalSR++; }
      if (isLS) { byDate[rawDate].ls++; totalLS++; }
      byDate[rawDate].total++;
    });

    $('dashTotalLoco').textContent = allUniqueLocos.size;
    $('dashTotalFSB').textContent = totalFSB;
    $('dashTotalEB').textContent = totalEB;
    $('dashTotalSR').textContent = totalSR;
    $('dashTotalLS').textContent = totalLS;

    dashboardTableBody.innerHTML = '';
    Object.keys(byDate).forEach(date => {
      const d = byDate[date];
      const tr = document.createElement('tr');

      const formatBadge = (val, type) => val > 0 ? `<span class="badge-count badge-${type}">${val}</span>` : `<span class="badge-count badge-zero">0</span>`;
      const locoList = [...d.locos].slice(0, 3).join(', ') + (d.locos.size > 3 ? '...' : '');

      tr.innerHTML = `
        <td><strong>${date}</strong></td>
        <td>${formatBadge(d.locos.size, 'loco')} ${locoList ? `<span style="font-size:11px;color:#94a3b8;margin-left:4px;">(${locoList})</span>` : ''}</td>
        <td>${formatBadge(d.fsb, 'fsb')}</td>
        <td>${formatBadge(d.eb, 'eb')}</td>
        <td>${formatBadge(d.sr, 'sr')}</td>
        <td>${formatBadge(d.ls, 'ls')}</td>
        <td><strong>${d.total}</strong></td>
      `;
      dashboardTableBody.appendChild(tr);
    });

    dashboardSection.classList.remove('hidden');
  }

  function processData(){
    const text=excelInput.value.replace(/\r/g,'').trim(); if(!text){showToast('Please paste Excel data first.','warning');return;}
    const layout = parseInputText(); if(!rawRows.length){showToast('No data rows found.','warning');return;}

    let issueColumn;
    if (headerless && currentFormat === 'postCommissioning') issueColumn = 4;
    else if (colSelect.value === '__NONE__') issueColumn = -1;
    else if (colSelect.value === '__AUTO__') issueColumn = detectIssueColumn(sourceHeaders, rawRows);
    else issueColumn = parseInt(colSelect.value, 10);

    processedData=[]; let inputCount=0,splitCount=0;

    rawRows.forEach((row,rowIndex)=>{
      if(row.every(c=>!String(c??'').trim())&&!chkPreserveBlank.checked)return; inputCount++;

      if(currentFormat==='postCommissioning'){
        const base=makePostCommissioning(row,rowIndex,layout,issueColumn);
        const issueText=issueColumn>=0?String(base.Remarks??''):'';
        const issues=issueColumn>=0&&issueText?splitIssues(issueText,patternSelect.value,chkCleanPrefix.checked):[''];
        if(issues.length>1)splitCount+=issues.length-1;
        issues.forEach((issue,i)=>{
          const item={
            Section: base.Section,
            Date: base.Date,
            'Train No/Loco No': base['Train No/Loco No'],
            Remarks: chkTrimText.checked ? String(issue).trim() : String(issue),
            __origRowIndex: rowIndex,
            __isSplit: issues.length > 1,
            __splitIndex: i + 1,
            __splitTotal: issues.length
          };
          processedData.push(item);
        });
        return;
      }

      if(currentFormat==='nmsData'){
        const base=makeNMS(row,rowIndex,issueColumn);
        const failureItems=getIssueItems(base.__failureTypeRaw);
        const stationItems=getIssueItems(base.__stationRaw);
        const reasonItems=getIssueItems(base.__reasonRaw);
        const descItems=getDescriptionIssueItems(base.__descriptionRaw, new Set([...failureItems,...stationItems,...reasonItems].map(x=>String(x.n))));

        const issueItems=failureItems.length>1?failureItems:(failureItems.length===1?failureItems:(descItems.length?descItems:(reasonItems.length?reasonItems:[{n:'1',v:''}])));
        if(issueItems.length>1) splitCount+=issueItems.length-1;
        issueItems.forEach((issueItem,i)=>{
          const item={...base}, num=String(issueItem.n);
          const description=getMatchingDescription(base.__descriptionRaw,num,new Set([...failureItems,...stationItems,...reasonItems].map(x=>String(x.n)),...descItems.map(x=>String(x.n))),issueItem.v||'');
          const failure=getIssueValue(base.__failureTypeRaw,num,i,'');
          const st=getIssueValue(base.__stationRaw,num,i,'');
          const cause=getIssueValue(base.__reasonRaw,num,i,'');
          item['STATION NAME']=cleanNumberedValue(st);
          item['Fault Type']=normalizeLocoFaultType(failure);
          item['NMS Observations']=cleanRemark(description);
          item['Time of Fault']=extractTime(item['NMS Observations']);
          item['Preliminary Cause Analysis as per NMS Data']=cleanNumberedValue(cause);
          item.__isSplit=issueItems.length>1; item.__splitIndex=i+1; item.__splitTotal=issueItems.length;
          delete item.__failureTypeRaw; delete item.__reasonRaw; delete item.__stationRaw; delete item.__descriptionRaw;
          processedData.push(item);
        });
        return;
      }
    });

    filteredData=[...processedData]; updateMetrics(inputCount,processedData.length,splitCount); renderOutputTable(); generateDashboardData();
    btnCopyExcel.disabled=btnDownloadExcel.disabled=processedData.length===0;tableSearch.disabled=processedData.length===0;
    showToast(`Processed: ${inputCount} received, ${processedData.length} generated (${splitCount} issues separated)`,'success');
  }

  function updateMetrics(a,b,c){metricInputRows.textContent=a;metricOutputRows.textContent=b;metricSplitCount.textContent=c;metricRatio.textContent=`${a?(b/a).toFixed(2):'1.00'}x`;}

  function renderOutputTable(){
    const headers=FORMATS[currentFormat].headers;
    if(!filteredData.length){emptyState.classList.remove('hidden');tableWrap.classList.add('hidden');tableInfo.textContent='Showing 0 rows';return;}
    emptyState.classList.add('hidden');tableWrap.classList.remove('hidden');tableHead.innerHTML='';tableBody.innerHTML='';
    const trh=document.createElement('tr'); const thn=document.createElement('th');thn.textContent='#';trh.appendChild(thn);
    headers.forEach(h=>{const th=document.createElement('th');th.textContent=h;trh.appendChild(th);});tableHead.appendChild(trh);
    filteredData.forEach((row,idx)=>{const tr=document.createElement('tr');const n=document.createElement('td');n.innerHTML=row.__isSplit?`<span class="split-tag">⚡ ${row.__splitIndex}/${row.__splitTotal}</span>`:String(idx+1);tr.appendChild(n);headers.forEach(h=>{const td=document.createElement('td');td.textContent=String(row[h]??'');tr.appendChild(td);});tableBody.appendChild(tr);});
    tableInfo.textContent=`Showing ${filteredData.length} of ${processedData.length} rows`;
  }

  function copyText(text){
    if(navigator.clipboard&&window.isSecureContext)return navigator.clipboard.writeText(text);
    return new Promise((resolve,reject)=>{const ta=document.createElement('textarea');ta.value=text;ta.style.position='fixed';ta.style.left='-9999px';document.body.appendChild(ta);ta.select();try{document.execCommand('copy')?resolve():reject(new Error('Copy failed'));}catch(e){reject(e)}finally{ta.remove()}});
  }

  async function copyToExcel(){
    if(isCopying)return;if(!processedData.length){showToast('No data available to copy.','warning');return;}isCopying=true;
    try{const headers=FORMATS[currentFormat].headers;const lines=[headers.join('\t')];processedData.forEach(r=>lines.push(headers.map(h=>String(r[h]??'').replace(/\r?\n/g,' ')).join('\t')));await copyText(lines.join('\n'));showToast(`${processedData.length} rows copied. Ready to paste into Excel.`,'success');}
    catch(e){console.error(e);showToast('Copy failed.','warning')}finally{setTimeout(()=>isCopying=false,500)}
  }

  function exportExcel(){
    if(isExporting)return;if(!processedData.length){showToast('No data available to export.','warning');return;}if(typeof XLSX==='undefined'){showToast('Excel library failed to load. Refresh the page and try again.','warning');return;}
    isExporting=true;const old=btnDownloadExcel.textContent;btnDownloadExcel.disabled=true;btnDownloadExcel.textContent='Exporting...';
    try{
      const headers=FORMATS[currentFormat].headers;
      const exportRows=processedData.map(r=>headers.map(h=>String(r[h]??'')));
      const ws=XLSX.utils.aoa_to_sheet([headers,...exportRows]);
      const border={top:{style:'thin',color:{rgb:'FFB7B7B7'}},bottom:{style:'thin',color:{rgb:'FFB7B7B7'}},left:{style:'thin',color:{rgb:'FFB7B7B7'}},right:{style:'thin',color:{rgb:'FFB7B7B7'}}};

      if(currentFormat==='postCommissioning'){
        const headerStyle={font:{name:'Calibri',sz:11,bold:true,color:{rgb:'FF000000'}},fill:{fgColor:{rgb:'FFFFD966'}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border};
        const dataStyle=c=>({font:{name:'Calibri',sz:11,bold:c===2,color:{rgb:'FF000000'}},fill:c===2?{fgColor:{rgb:'FF00FF00'}}:undefined,alignment:{horizontal:c===3?'left':'center',vertical:c===3?'top':'center',wrapText:true},border});

        for(let R=0;R<=exportRows.length;R++)for(let C=0;C<4;C++){const ref=XLSX.utils.encode_cell({r:R,c:C});const cell=ws[ref]||(ws[ref]={});cell.t='s';cell.v=String(cell.v??'');cell.s=R===0?headerStyle:dataStyle(C);cell.z='@';}

        const merges=[];let start=0;
        while(start<processedData.length){
          const source=processedData[start].__origRowIndex;let end=start+1;
          while(end<processedData.length&&processedData[end].__origRowIndex===source)end++;
          if(end-start>1)[0,1,2].forEach(C=>merges.push({s:{r:start+1,c:C},e:{r:end,c:C}}));
          start=end;
        }
        ws['!merges']=merges;
        ws['!cols']=[{wch:18},{wch:15},{wch:25},{wch:90}];
        ws['!rows']=[{hpt:28}];
        exportRows.forEach(r=>ws['!rows'].push({hpt:Math.min(150,Math.max(30,Math.ceil(Math.max(1,r[3].length/78))*18))}));
      }else{
        const headerStyle={font:{name:'Calibri',sz:11,bold:true,color:{rgb:'FF000000'}},fill:{fgColor:{rgb:'FFD9EAD3'}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border};
        const dataStyle={font:{name:'Calibri',sz:11,color:{rgb:'FF000000'}},alignment:{horizontal:'left',vertical:'top',wrapText:true},border};
        for(let R=0;R<=exportRows.length;R++)for(let C=0;C<headers.length;C++){const ref=XLSX.utils.encode_cell({r:R,c:C});const cell=ws[ref]||(ws[ref]={});cell.t='s';cell.v=String(cell.v??'');cell.s=R===0?headerStyle:dataStyle;cell.z='@';}
        ws['!cols']=[{wch:9},{wch:14},{wch:20},{wch:20},{wch:14},{wch:16},{wch:14},{wch:18},{wch:16},{wch:60},{wch:45}];
        ws['!rows']=[{hpt:32}];
        exportRows.forEach(r=>ws['!rows'].push({hpt:Math.min(180,Math.max(24,Math.ceil(Math.max(1,Math.max(...r.map(x=>x.length)) / 75))*18))}));
      }

      const wb=XLSX.utils.book_new();
      const sheetName=currentFormat==='postCommissioning'?'Transformed Data':'NMS Data Report';
      XLSX.utils.book_append_sheet(wb,ws,sheetName);
      const now=new Date();
      const pad=n=>String(n).padStart(2,'0');
      const todayStr=`${pad(now.getDate())}-${pad(now.getMonth()+1)}-${now.getFullYear()}`;
      const formatLabel=currentFormat==='postCommissioning'?'Post_Commissioning':'NMS_Data';
      const filename=`${formatLabel}_${todayStr}.xlsx`;
      XLSX.writeFile(wb,filename,{bookType:'xlsx',compression:true});showToast(`Downloaded ${filename} (${exportRows.length} rows).`,'success');
    }catch(e){console.error(e);showToast(`Export failed: ${e.message||'Please try again.'}`,'warning')}finally{setTimeout(()=>{isExporting=false;btnDownloadExcel.disabled=processedData.length===0;btnDownloadExcel.textContent=old},700)}
  }

  function showToast(message,type='info'){const t=document.createElement('div');t.className=`toast toast-${type}`;t.textContent=message;toastContainer.appendChild(t);setTimeout(()=>{t.style.opacity='0';t.style.transition='opacity .25s';setTimeout(()=>t.remove(),250)},3000)}

  function setFormat(format){
    currentFormat=format;
    $('cardPostCommissioning').classList.toggle('active', format==='postCommissioning');
    $('cardNMSData').classList.toggle('active', format==='nmsData');
    document.querySelectorAll('input[name="dataFormat"]').forEach(r=>r.checked = (r.value === format));

    if (lblMergeTrainLoco) lblMergeTrainLoco.style.display = format === 'postCommissioning' ? 'inline-flex' : 'none';

    formatHint.textContent = FORMATS[format].hint;
    pasteSubtext.textContent = format === 'postCommissioning'
      ? 'Paste tab-separated data from Excel (Date, Section, Train No, Loco No, Remarks).'
      : 'Paste your Main Data sheet directly from Excel.';

    processedData=[];filteredData=[];rawRows=[];sourceHeaders=[];headerless=false;excelInput.value='';tableSearch.value='';tableSearch.disabled=true;btnCopyExcel.disabled=btnDownloadExcel.disabled=true;updateMetrics(0,0,0);updateColumnDropdown([]);renderOutputTable();dashboardSection.classList.add('hidden');
  }

  document.querySelectorAll('input[name="dataFormat"]').forEach(r=>r.addEventListener('change',()=>setFormat(r.value)));
  $('cardPostCommissioning').addEventListener('click', ()=>{ setFormat('postCommissioning'); });
  $('cardNMSData').addEventListener('click', ()=>{ setFormat('nmsData'); });

  excelInput.addEventListener('input',parseInputText);
  btnProcess.addEventListener('click',e=>{e.preventDefault();processData()});
  btnCopyExcel.addEventListener('click',e=>{e.preventDefault();copyToExcel()});
  btnDownloadExcel.addEventListener('click',e=>{e.preventDefault();exportExcel()});
  btnClearInput.addEventListener('click',e=>{e.preventDefault();excelInput.value='';parseInputText();showToast('Input cleared.','info')});
  btnClearAll.addEventListener('click',e=>{e.preventDefault();setFormat(currentFormat);showToast('All fields reset.','info')});

  btnCopyDashboard.addEventListener('click', async e => {
    e.preventDefault();
    if (!processedData.length) return;
    const rows = [['Date', 'Locos (Unique Count)', 'FSB Count', 'EB Count', 'SR Count', 'LS Count', 'Total Faults']];
    document.querySelectorAll('#dashboardTableBody tr').forEach(tr => {
      const cells = [...tr.querySelectorAll('td')].map(td => td.textContent.replace(/\s*\([^)]*\)/, '').trim());
      rows.push(cells);
    });
    const tsv = rows.map(r => r.join('\t')).join('\n');
    await copyText(tsv);
    showToast('Dashboard summary copied to clipboard! Ready to paste into Excel.', 'success');
  });

  btnSampleData.addEventListener('click',e=>{
    e.preventDefault();
    if (currentFormat === 'postCommissioning') {
      excelInput.value=[
        'Date\tSection\tTrain No\tLoco No\tRemarks',
        '01-09-2026\tBJW-ADI\t22955\t37330\t1. Emergency Brake (EB) was Applied at BJW at Time - 05:39:06 at Abs Location 405174, due to Loco Reader Issue, R-573 was missed by the loco 2. FSB was Applied at BJW at Time - 06:11:28 at Abs Location 402926. Due to Speed Control (PSR - 15, SPEED - 20). 3. Emergency Brake was Applied at KKEC at Time - 07:43:32 at Abs Location 494243. Due to Speed Control (PSR - 15, SPEED - 22).',
        '01-09-2026\tANND-ADI\t19028\t37492\t1. LOCO READER ISSUE (FSB) 2. S2S ISSUE (SR)',
        '02-09-2026\tADI-GER\t12901\t37111\t1. EB APPLIED AT GER 2. FS - LS OCCURRED'
      ].join('\n');
    } else {
      excelInput.value=[
        'SR NO.\tDate\tSection\tTrain No\tLoco No\tLOCO FIT / UNFIT\tLOCO OEM\tUP / DN\tReason\tFailure type\tStation\tFailure Description',
        '1\t01-09-2026\tANND-ADI\t\t37818\tFIT\tMEDHA\tDN\tDATA IS NOT GENERATED BEFORE VATVA\tFSB\tNIL\t1. FSB APPLIED AT VATVA DUE TO NMS DOWN',
        '2\t01-09-2026\tADI-ANND\t19028\t37492\tFIT\tMEDHA\tUP\t1. BJW-LC-258 OFFLINE 2. GER-BJD S2S ISSUE 3. LOCO READER ISSUE\t3.FS - SR\t3.ND\t1. IN TRAIN NO 19028 / LOCO NO 37492 2. BOTH TAG MISS (EB APPLIED) AT ND AT TIME 08:12:00 AT ABS LOCATION 450772 DUE TO TAG (R-299) MISSED BY LOCO (LOCO READER ISSUE) 3. DATA IS NOT GENERATED BEFORE GER DUE TO NMS DOWN (FS - LS)'
      ].join('\n');
    }
    parseInputText(); processData();
  });

  tableSearch.addEventListener('input',()=>{
    const q=tableSearch.value.toLowerCase().trim(),headers=FORMATS[currentFormat].headers;
    filteredData=q?processedData.filter(r=>headers.some(h=>String(r[h]??'').toLowerCase().includes(q))):[...processedData];
    renderOutputTable();
  });

  setFormat('postCommissioning');
});
