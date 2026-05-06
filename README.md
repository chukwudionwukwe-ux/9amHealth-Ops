[9amHealth_VP_Email_Template.html](https://github.com/user-attachments/files/27450677/9amHealth_VP_Email_Template.html)
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>9amHealth — VP Email Generator</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#f0f2f5;color:#111;font-size:14px}

/* ── BUILDER ── */
.app{display:grid;grid-template-columns:340px 1fr;height:100vh;overflow:hidden}
.sidebar{background:#0f1923;display:flex;flex-direction:column;overflow-y:auto}
.sb-logo{padding:18px 16px 14px;border-bottom:1px solid rgba(255,255,255,.1)}
.sb-logo-title{font-size:14px;font-weight:700;color:#fff}
.sb-logo-sub{font-size:10px;color:rgba(255,255,255,.4);margin-top:2px;letter-spacing:.04em;text-transform:uppercase}
.sb-section{padding:12px 16px;border-bottom:1px solid rgba(255,255,255,.07)}
.sb-section-title{font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:rgba(255,255,255,.35);margin-bottom:10px}
.field{margin-bottom:10px}
.field label{display:block;font-size:11px;font-weight:600;color:rgba(255,255,255,.5);margin-bottom:3px}
.field input,.field select,.field textarea{width:100%;padding:7px 9px;background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.12);color:#fff;font-size:12px;font-family:inherit;outline:none}
.field input:focus,.field select:focus,.field textarea:focus{border-color:rgba(255,255,255,.3);background:rgba(255,255,255,.1)}
.field select option{background:#1a2332;color:#fff}
.field textarea{height:52px;resize:vertical}
.field-row{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.tp-row{display:flex;gap:6px;margin-bottom:6px}
.tp-row input{flex:1;padding:5px 7px;background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.12);color:#fff;font-size:11px;font-family:inherit}
.tp-row button{background:rgba(255,0,0,.2);border:none;color:#fca5a5;padding:4px 8px;cursor:pointer;font-size:12px}
.add-tp{background:none;border:1px dashed rgba(255,255,255,.2);color:rgba(255,255,255,.5);padding:5px 10px;font-size:11px;cursor:pointer;width:100%;margin-top:4px;font-family:inherit}
.add-tp:hover{border-color:rgba(255,255,255,.4);color:#fff}

/* Status selector */
.status-btns{display:flex;gap:4px}
.status-btn{flex:1;padding:5px 4px;border:none;font-size:11px;font-weight:700;cursor:pointer;opacity:.4;font-family:inherit}
.status-btn.green{background:#14532d;color:#86efac}
.status-btn.amber{background:#78350f;color:#fde68a}
.status-btn.red{background:#7f1d1d;color:#fca5a5}
.status-btn.active{opacity:1}

/* Import zone */
.import-zone{border:1px dashed rgba(255,255,255,.2);padding:10px;text-align:center;cursor:pointer;position:relative;margin-bottom:8px}
.import-zone:hover{border-color:rgba(255,255,255,.4);background:rgba(255,255,255,.04)}
.import-zone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.import-zone-text{font-size:11px;color:rgba(255,255,255,.45)}
.import-status{font-size:11px;padding:4px 8px;margin-top:6px;display:none}
.import-status.ok{display:block;background:rgba(134,239,172,.15);color:#86efac}
.import-status.err{display:block;background:rgba(252,165,165,.15);color:#fca5a5}

/* Action buttons */
.sb-actions{padding:12px 16px;margin-top:auto;border-top:1px solid rgba(255,255,255,.1)}
.btn-gen{width:100%;background:#1d4ed8;color:#fff;border:none;padding:10px;font-size:13px;font-weight:700;cursor:pointer;font-family:inherit;margin-bottom:8px}
.btn-gen:hover{background:#1e40af}
.btn-copy{width:100%;background:#166534;color:#fff;border:none;padding:9px;font-size:13px;font-weight:600;cursor:pointer;font-family:inherit;margin-bottom:6px}
.btn-copy:hover{background:#14532d}
.btn-reset{width:100%;background:rgba(255,255,255,.06);color:rgba(255,255,255,.5);border:1px solid rgba(255,255,255,.1);padding:7px;font-size:12px;cursor:pointer;font-family:inherit}

/* Preview pane */
.preview-pane{background:#e5e7eb;overflow-y:auto;padding:20px}
.preview-pane-header{display:flex;align-items:center;justify-content:space-between;background:#fff;border:1px solid #e5e7eb;padding:10px 16px;margin-bottom:16px}
.preview-label{font-size:12px;font-weight:600;color:#374151}
.preview-subject{font-size:11px;color:#6b7280;margin-top:2px}
.mobile-toggle{display:flex;gap:6px}
.mt-btn{padding:5px 12px;border:1px solid #e5e7eb;background:#fff;font-size:11px;cursor:pointer;font-family:inherit;color:#374151}
.mt-btn.active{background:#1d4ed8;color:#fff;border-color:#1d4ed8}
.preview-wrap{transition:max-width .2s}
.preview-wrap.mobile{max-width:390px;margin:0 auto}

/* ─────────────────────────────────────────────────
   EMAIL STYLES — no rounded edges (Taire)
───────────────────────────────────────────────── */
.em-wrap{background:#f4f4f4;padding:16px 0;font-family:'Segoe UI',Arial,sans-serif}
.em-box{max-width:620px;margin:0 auto;background:#fff;border:1px solid #d1d5db}
.em-head{background:#0f1923;padding:18px 22px}
.em-head-row{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px}
.em-co{font-size:11px;font-weight:700;color:rgba(255,255,255,.5);letter-spacing:.08em;text-transform:uppercase}
.em-prog{font-size:10px;color:rgba(255,255,255,.35);letter-spacing:.05em;text-transform:uppercase}
.em-subject{font-size:17px;font-weight:700;color:#fff;line-height:1.3;margin-bottom:3px}
.em-meta{font-size:11px;color:rgba(255,255,255,.45)}
.em-body{padding:18px 22px}
.em-intro{font-size:13px;color:#374151;line-height:1.65;margin-bottom:14px}

/* Section headers — different colors per section (Taire) */
.sh{padding:7px 11px;font-size:9px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;color:#fff;margin-bottom:0;margin-top:2px}
.sh.blue{background:#1a3a5c}.sh.green{background:#14532d}
.sh.amber{background:#78350f}.sh.red{background:#7f1d1d}
.sh.gray{background:#374151}.sh.navy{background:#1e3a5f}

/* Tables — NO rounded edges (Taire) */
.et{width:100%;border-collapse:collapse;margin-bottom:14px}
.et th{background:#f3f4f6;padding:8px 11px;font-size:10px;font-weight:700;text-align:left;border:1px solid #e5e7eb;color:#374151;text-transform:uppercase;letter-spacing:.04em;white-space:nowrap}
.et th.wk{font-weight:900;color:#111} /* Week bold (Taire) */
.et td{padding:8px 11px;font-size:13px;border:1px solid #e5e7eb;color:#374151;vertical-align:top}
.et tr:nth-child(even) td{background:#fafafa}
.et td.wk{font-weight:700;color:#111} /* Week bold (Taire) */

/* Traffic lights */
.tg{background:#dcfce7;color:#14532d;font-weight:700;border:1px solid #86efac}
.ta{background:#fef9c3;color:#713f12;font-weight:700;border:1px solid #fde047}
.tr{background:#fee2e2;color:#7f1d1d;font-weight:700;border:1px solid #fca5a5}

/* Cell comments at bottom of cell (Taire) */
.cc{display:block;font-size:10px;font-weight:400;font-style:italic;color:#6b7280;margin-top:3px;line-height:1.4}

/* Confirmed gap = RED block, warning = YELLOW block (Taire) */
.gap-r{background:#fee2e2;border-left:3px solid #dc2626;padding:8px 11px;margin-bottom:10px}
.gap-r .gl{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#991b1b;margin-bottom:3px}
.gap-r .gt{font-size:12px;color:#7f1d1d;line-height:1.5}
.gap-y{background:#fef9c3;border-left:3px solid #ca8a04;padding:8px 11px;margin-bottom:10px}
.gap-y .gl{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#713f12;margin-bottom:3px}
.gap-y .gt{font-size:12px;color:#78350f;line-height:1.5}

.em-div{height:1px;background:#e5e7eb;margin:14px 0}
.em-foot{background:#f9fafb;border-top:1px solid #e5e7eb;padding:12px 22px;font-size:11px;color:#6b7280}
.em-foot strong{color:#111;font-size:12px;display:block;margin-bottom:2px}
.em-foot a{color:#1d4ed8}
.em-conf{color:#9ca3af;font-size:10px;margin-top:6px}

/* Week selector for history view */
.week-tabs{display:flex;gap:0;border-bottom:2px solid #e5e7eb;margin-bottom:16px;flex-wrap:wrap}
.wt{padding:6px 14px;border:none;background:none;font-size:12px;font-weight:500;cursor:pointer;color:#6b7280;border-bottom:2px solid transparent;margin-bottom:-2px;font-family:inherit}
.wt.active{color:#1d4ed8;border-bottom-color:#1d4ed8;font-weight:700}

.import-file-btn{display:block;border:1px dashed rgba(255,255,255,.2);padding:8px 10px;text-align:center;cursor:pointer;color:rgba(255,255,255,.5);font-size:11px;font-family:inherit}
.import-file-btn:hover{border-color:rgba(255,255,255,.4);color:#fff;background:rgba(255,255,255,.04)}
@media(max-width:700px){.app{grid-template-columns:1fr;height:auto}.sidebar{max-height:50vh}.preview-pane{padding:10px}}
</style>
</head>
<body>
<div class="app">

<!-- ── SIDEBAR ── -->
<div class="sidebar">
  <div class="sb-logo">
    <div class="sb-logo-title">9amHealth Email Generator</div>
    <div class="sb-logo-sub">VP Dashboard &middot; Hugo Technologies</div>
  </div>

  <!-- IMPORT -->
  <div class="sb-section">
    <div class="sb-section-title">Auto-import — upload each file separately</div>
    <div style="font-size:10px;color:rgba(255,255,255,.3);margin-bottom:10px;line-height:1.5">Each metric comes from a different export. Upload them one at a time — fields fill in automatically.</div>

    <div style="margin-bottom:7px">
      <div style="font-size:10px;font-weight:600;color:rgba(255,255,255,.4);margin-bottom:4px;text-transform:uppercase;letter-spacing:.05em">&#128200; Answer Rate CSV</div>
      <label class="import-file-btn" title="Your WFM answer rate export — reads answer rate, inbound, voicemail, wait time">
        <input type="file" accept=".csv,.xlsx,.xls" onchange="importAnswerRate(this)" style="display:none">
        <span>&#128196; Upload answer rate file</span>
      </label>
      <div class="import-status" id="import-st-ans"></div>
    </div>

    <div style="margin-bottom:7px">
      <div style="font-size:10px;font-weight:600;color:rgba(255,255,255,.4);margin-bottom:4px;text-transform:uppercase;letter-spacing:.05em">&#128222; Call Volume &amp; Duration CSV</div>
      <label class="import-file-btn" title="Your daily/weekly call volume export — reads total calls and avg duration">
        <input type="file" accept=".csv,.xlsx,.xls" onchange="importCallVolume(this)" style="display:none">
        <span>&#128196; Upload call volume file</span>
      </label>
      <div class="import-status" id="import-st-vol"></div>
    </div>

    <div style="margin-bottom:4px">
      <div style="font-size:10px;font-weight:600;color:rgba(255,255,255,.4);margin-bottom:4px;text-transform:uppercase;letter-spacing:.05em">&#128203; Any Report (auto-detect)</div>
      <label class="import-file-btn" title="Try uploading any CSV and the importer will attempt to read whatever columns it finds">
        <input type="file" accept=".csv,.xlsx,.xls" onchange="importReport(this)" style="display:none">
        <span>&#128196; Upload any report file</span>
      </label>
      <div class="import-status" id="import-st"></div>
    </div>
  </div>

  <!-- HEADER -->
  <div class="sb-section">
    <div class="sb-section-title">Header</div>
    <div class="field-row">
      <div class="field"><label>Report type</label>
        <select id="f-type" onchange="generate()">
          <option>Daily Performance Update</option>
          <option>Weekly Performance Summary</option>
          <option>Weekly Program Update</option>
        </select>
      </div>
      <div class="field"><label>Week #</label><input id="f-week" value="Week 19" oninput="generate()"></div>
    </div>
    <div class="field"><label>Date / Date range</label><input id="f-date" value="05/06/2026" oninput="generate()"></div>
    <div class="field"><label>Opening line</label><textarea id="f-opening" oninput="generate()">Below is today's program performance update. The program continues into Week 19 with metrics within target.</textarea></div>
  </div>

  <!-- PRODUCTION -->
  <div class="sb-section">
    <div class="sb-section-title">Production</div>
    <div class="field-row">
      <div class="field"><label>Total calls</label><input id="f-calls" value="212" oninput="generate()"></div>
      <div class="field"><label>Avg duration (min)</label><input id="f-dur" value="10.34" oninput="generate()"></div>
    </div>
    <div class="field"><label>Duration status</label>
      <div class="status-btns">
        <button class="status-btn green active" onclick="setStatus('dur','green',this)">&#10003; At goal</button>
        <button class="status-btn amber" onclick="setStatus('dur','amber',this)">&#9888; Near limit</button>
        <button class="status-btn red" onclick="setStatus('dur','red',this)">&#10007; Above target</button>
      </div>
    </div>
    <div class="field"><label>Duration comment (optional)</label><input id="f-dur-comment" placeholder="e.g. Slight uptick on complex calls" oninput="generate()"></div>
  </div>

  <!-- ANSWER RATE -->
  <div class="sb-section">
    <div class="sb-section-title">Answer Rate</div>
    <div class="field-row">
      <div class="field"><label>Inbound answered</label><input id="f-inbound" value="235" oninput="generate()"></div>
      <div class="field"><label>Voicemails</label><input id="f-vm" value="1" oninput="generate()"></div>
    </div>
    <div class="field-row">
      <div class="field"><label>Total relevant inbound</label><input id="f-total" value="238" oninput="generate()"></div>
      <div class="field"><label>Wait time (sec)</label><input id="f-wait" value="26" oninput="generate()"></div>
    </div>
    <div class="field-row">
      <div class="field"><label>Answer rate (%)</label><input id="f-ans" value="98.74" oninput="generate()"></div>
    </div>
    <div class="field"><label>Answer rate comment (optional)</label><input id="f-ans-comment" placeholder="e.g. W17 dip was 9amHealth staffing gap" oninput="generate()"></div>
  </div>

  <!-- FLAGS -->
  <div class="sb-section">
    <div class="sb-section-title">Flags (leave blank if none)</div>
    <div class="field"><label style="color:#fca5a5">&#9632; Confirmed gap (RED)</label>
      <textarea id="f-gap-red" placeholder="e.g. Live call monitoring not available in Zendesk Talk. All QA post-call." oninput="generate()"></textarea>
    </div>
    <div class="field"><label style="color:#fde68a">&#9632; Watch item (YELLOW)</label>
      <textarea id="f-gap-yellow" placeholder="e.g. Adherence reporting gap — WFM identified as long-term solution." oninput="generate()"></textarea>
    </div>
  </div>

  <!-- OPS NOTE -->
  <div class="sb-section">
    <div class="sb-section-title">Ops Note (optional)</div>
    <div class="field"><textarea id="f-ops" placeholder="e.g. Wave 2 Uptraining Day 2 of 3 complete. Four agents progressing through blended model ticket training with Jalana." oninput="generate()"></textarea></div>
  </div>

  <!-- TOUCHPOINTS -->
  <div class="sb-section">
    <div class="sb-section-title">Upcoming Touchpoints</div>
    <div id="tp-list"></div>
    <button class="add-tp" onclick="addTP()">+ Add touchpoint</button>
  </div>

  <!-- SENDER -->
  <div class="sb-section">
    <div class="sb-section-title">Sender</div>
    <div class="field-row">
      <div class="field"><label>Name</label><input id="f-sender" value="Chukwudi Onwukwe" oninput="generate()"></div>
      <div class="field"><label>Title</label><input id="f-stitle" value="Program Manager" oninput="generate()"></div>
    </div>
    <div class="field-row">
      <div class="field"><label>Phone</label><input id="f-phone" value="+1 (540) 319-0277" oninput="generate()"></div>
      <div class="field"><label>Email</label><input id="f-email" value="chukwudi.onwukwe@hugotech.co" oninput="generate()"></div>
    </div>
  </div>

  <!-- ACTIONS -->
  <div class="sb-actions">
    <button class="btn-gen" onclick="generate()">&#9654; Generate Email</button>
    <button class="btn-copy" onclick="copyEmail()">&#128203; Copy HTML to Clipboard</button>
    <button class="btn-reset" onclick="showHistory()">&#128337; View Week History</button>
  </div>
</div>

<!-- ── PREVIEW PANE ── -->
<div class="preview-pane">
  <div class="preview-pane-header">
    <div>
      <div class="preview-label">Email Preview</div>
      <div class="preview-subject" id="preview-subject">Fill in fields and click Generate</div>
    </div>
    <div class="mobile-toggle">
      <button class="mt-btn active" onclick="setView('desktop',this)">&#128196; Desktop</button>
      <button class="mt-btn" onclick="setView('mobile',this)">&#128241; Mobile</button>
      <button class="mt-btn" onclick="showHistory()">&#128337; History</button>
    </div>
  </div>
  <div class="preview-wrap" id="preview-wrap">
    <div id="preview-output">
      <div style="text-align:center;padding:60px 20px;color:#9ca3af">
        <div style="font-size:40px;margin-bottom:12px">&#9993;</div>
        <div style="font-size:14px;font-weight:600;margin-bottom:6px">Fill in the fields on the left</div>
        <div style="font-size:12px">The email will generate automatically as you type, or click Generate.</div>
      </div>
    </div>
  </div>
</div>
</div>

<!-- HISTORY MODAL -->
<div id="history-modal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:100;overflow-y:auto;padding:20px">
  <div style="max-width:900px;margin:0 auto;background:#fff;border:1px solid #e5e7eb">
    <div style="background:#0f1923;padding:14px 18px;display:flex;align-items:center;justify-content:space-between">
      <div style="color:#fff;font-size:14px;font-weight:700">Weekly Performance History</div>
      <button onclick="document.getElementById('history-modal').style.display='none'" style="background:none;border:none;color:rgba(255,255,255,.5);font-size:18px;cursor:pointer">&#215;</button>
    </div>
    <div style="padding:16px">
      <div id="history-week-tabs" class="week-tabs"></div>
      <div id="history-content"></div>
    </div>
  </div>
</div>

<script>
// ── STATE ────────────────────────────────────────────────────────────────────
var durStatus='green';
var history=JSON.parse(localStorage.getItem('9am-email-history')||'[]');
var defaultTPs=[
  {date:'05/06/2026',act:'SHBP Ticket Uptraining — Wave 2, Day 3 (Final)',obj:'Four agents complete blended model ticket training with Jalana'},
  {date:'Week of 05/11',act:'Benchmark Discussion with 9amHealth Leadership',obj:'Evaluate first 8 days of ticket data to inform blended model benchmark'}
];

// ── HELPERS ──────────────────────────────────────────────────────────────────
function v(id){return(document.getElementById(id)||{value:''}).value||''}
function setStatus(which,color,btn){
  durStatus=color;
  btn.parentElement.querySelectorAll('.status-btn').forEach(function(b){b.classList.remove('active')});
  btn.classList.add('active');
  generate();
}
function ansSt(r){var n=parseFloat(r);return n>=98?'g':n>=95?'a':'r'}
function waitSt(s){var n=parseInt(s);return n<=20?'g':n<=30?'a':'r'}
function stIcon(s){return s==='g'?'&#10003;':s==='a'?'&#9888;':'&#10007;'}
function stClass(s){return s==='g'?'tg':s==='a'?'ta':'tr'}
function setView(v,btn){
  document.querySelectorAll('.mt-btn').forEach(function(b){b.classList.remove('active')});
  btn.classList.add('active');
  document.getElementById('preview-wrap').className='preview-wrap'+(v==='mobile'?' mobile':'');
}

// ── TOUCHPOINTS ──────────────────────────────────────────────────────────────
function addTP(date,act,obj){
  var list=document.getElementById('tp-list');
  var idx=list.children.length;
  var row=document.createElement('div');row.className='tp-row';
  row.innerHTML=
    '<input placeholder="Date" value="'+(date||'')+'" oninput="generate()">'+
    '<input placeholder="Activity" style="flex:2" value="'+(act||'')+'" oninput="generate()">'+
    '<input placeholder="Objective" style="flex:2" value="'+(obj||'')+'" oninput="generate()">'+
    '<button onclick="this.parentElement.remove();generate()">&#215;</button>';
  list.appendChild(row);
}
function getTPs(){
  var rows=document.getElementById('tp-list').children;
  var out=[];
  for(var i=0;i<rows.length;i++){
    var ins=rows[i].querySelectorAll('input');
    if(ins[0]&&ins[0].value)out.push({date:ins[0].value,act:ins[1]?ins[1].value:'',obj:ins[2]?ins[2].value:''});
  }
  return out;
}

// ── ANSWER RATE IMPORT ───────────────────────────────────────────────────────
// Reads: answer rate %, inbound answered, voicemail, total inbound, wait time
// Matches your WFM export: "Inbound Answered (No Voicemail)", "Call Answer Rate", etc.
function importAnswerRate(input){
  var file=input.files[0];if(!file)return;
  var st=document.getElementById('import-st-ans');
  st.className='import-status info';st.style.display='block';st.textContent='Reading answer rate file...';

  function apply(rows){
    var filled=0;
    rows.forEach(function(r){
      var keys=Object.keys(r).map(function(k){return k.toLowerCase();});
      var vals=Object.values(r);
      function get(patterns){
        for(var i=0;i<keys.length;i++){
          for(var j=0;j<patterns.length;j++){
            if(keys[i].includes(patterns[j].toLowerCase()))return(vals[i]||'').toString().trim();
          }
        }
        return '';
      }
      // Answer rate
      var ans=get(['call answer rate','answer rate']);
      if(ans){var n=parseFloat(ans.replace('%',''));if(n>0&&n<=100){document.getElementById('f-ans').value=n>1?n.toFixed(2):(n*100).toFixed(2);filled++;}}
      // Inbound answered
      var inb=get(['inbound answered (no voicemail)','inbound answered','inbound answer']);
      if(inb&&parseInt(inb)>0){document.getElementById('f-inbound').value=parseInt(inb);filled++;}
      // Voicemail
      var vm=get(['voicemail calls','voicemail']);
      if(vm&&parseInt(vm)>=0){document.getElementById('f-vm').value=parseInt(vm);filled++;}
      // Total inbound
      var tot=get(['total relevant inbound calls','total relevant inbound','total inbound']);
      if(tot&&parseInt(tot)>0){document.getElementById('f-total').value=parseInt(tot);filled++;}
      // Wait time — could be "13 sec" or just "13"
      var wt=get(['call wait time','avg call wait','wait time','queue wait']);
      if(wt){var ws=parseFloat(wt.replace(/[^0-9.]/g,''));if(ws>0){document.getElementById('f-wait').value=Math.round(ws);filled++;}}
    });
    if(filled>0){
      st.className='import-status ok';
      st.textContent='✓ Answer rate imported — '+filled+' fields filled';
      generate();
    } else {
      st.className='import-status err';
      st.textContent='No answer rate columns found. Check column names match your WFM export.';
    }
    input.value='';
  }
  readFile(file,apply,st);
}

// ── CALL VOLUME / DURATION IMPORT ────────────────────────────────────────────
// Reads: total calls completed, avg call duration
// Matches your call system export
function importCallVolume(input){
  var file=input.files[0];if(!file)return;
  var st=document.getElementById('import-st-vol');
  st.className='import-status info';st.style.display='block';st.textContent='Reading call volume file...';

  function apply(rows){
    var filled=0;
    // Aggregate totals across all rows (daily data = sum; weekly data = single row)
    var totalCalls=0;var durationSum=0;var durationCount=0;

    rows.forEach(function(r){
      var keys=Object.keys(r).map(function(k){return k.toLowerCase();});
      var vals=Object.values(r);
      function get(patterns){
        for(var i=0;i<keys.length;i++){
          for(var j=0;j<patterns.length;j++){
            if(keys[i].includes(patterns[j].toLowerCase()))return(vals[i]||'').toString().trim();
          }
        }
        return '';
      }
      // Calls
      var inb=parseInt(get(['accepted call legs','inbound answered','total calls completed','total calls'])||0);
      var out=parseInt(get(['completed outbound','outbound calls'])||0);
      if(inb>0||out>0){totalCalls+=inb+out;filled++;}
      // Duration
      var dur=parseFloat((get(['avg call duration','call duration','talk time (min)','avg duration'])||'').replace(/[^0-9.]/g,''));
      if(dur>0){durationSum+=dur;durationCount++;}
    });

    if(totalCalls>0){document.getElementById('f-calls').value=totalCalls;}
    if(durationCount>0){document.getElementById('f-dur').value=(durationSum/durationCount).toFixed(2);}

    if(filled>0||durationCount>0){
      st.className='import-status ok';
      st.textContent='✓ Call volume imported — '+totalCalls+' total calls'+(durationCount?' · avg duration '+(durationSum/durationCount).toFixed(1)+'m':'');
      generate();
    } else {
      st.className='import-status err';
      st.textContent='No call volume columns found. Expected "Accepted call legs" or "Total calls".';
    }
    input.value='';
  }
  readFile(file,apply,st);
}

// ── SHARED FILE READER ────────────────────────────────────────────────────────
function readFile(file,applyFn,statusEl){
  var ext=file.name.split('.').pop().toLowerCase();
  if(ext==='csv'){
    var reader=new FileReader();
    reader.onload=function(e){
      try{applyFn(parseCSV(e.target.result));}
      catch(ex){statusEl.className='import-status err';statusEl.textContent='Error: '+ex.message;}
    };
    reader.readAsText(file);
  } else {
    var reader=new FileReader();
    reader.onload=function(e){
      try{
        var wb=XLSX.read(e.target.result,{type:'array'});
        var ws=wb.Sheets[wb.SheetNames[0]];
        applyFn(XLSX.utils.sheet_to_json(ws,{defval:''}));
      }catch(ex){statusEl.className='import-status err';statusEl.textContent='Error: '+ex.message;}
    };
    reader.readAsArrayBuffer(file);
  }
}

function parseCSV(text){
  var lines=text.trim().split('\n');
  var headers=lines[0].split(',').map(function(h){return h.replace(/"/g,'').trim();});
  return lines.slice(1).map(function(line){
    var vals=[];var cur='';var inQ=false;
    for(var i=0;i<line.length;i++){
      if(line[i]==='"'){inQ=!inQ;}
      else if(line[i]===','&&!inQ){vals.push(cur.trim());cur='';}
      else cur+=line[i];
    }
    vals.push(cur.trim());
    var obj={};headers.forEach(function(h,i){obj[h]=vals[i]||'';});
    return obj;
  }).filter(function(r){return Object.values(r).some(function(v){return v;});});
}

// ── AUTO-IMPORT (any file) ────────────────────────────────────────────────────
function importReport(input){
  var file=input.files[0];if(!file)return;
  var st=document.getElementById('import-st');
  st.className='import-status info';st.style.display='block';st.textContent='Reading file...';

  function applyData(rows){
    // Normalize keys to lowercase for matching
    rows=rows.map(function(r){
      var out={};
      Object.keys(r).forEach(function(k){out[k.toLowerCase()]=r[k];});
      return out;
    });
    // Try to detect what kind of report this is
    var filled=0;

    // Find answer rate row
    rows.forEach(function(r){
      var keys=Object.keys(r);
      var vals=Object.values(r).join(' ').toLowerCase();

      // Answer rate
      if(vals.includes('answer rate')||keys.some(function(k){return k.includes('answer rate')})){
        var ans=r['call answer rate']||r['answer rate']||'';
        if(ans){var n=ans.replace('%','').trim();if(parseFloat(n)>0){document.getElementById('f-ans').value=parseFloat(n).toFixed(2);filled++;}}
      }
      // Inbound answered
      var inb=r['inbound answered (no voicemail)']||r['inbound answered']||r['accepted call legs']||'';
      if(inb&&parseInt(inb)>0){document.getElementById('f-inbound').value=parseInt(inb);filled++;}
      // Voicemail
      var vm2=r['voicemail calls']||r['voicemail']||'';
      if(vm2&&parseInt(vm2)>=0){document.getElementById('f-vm').value=parseInt(vm2);filled++;}
      // Total inbound
      var tot=r['total relevant inbound calls']||r['total relevant inbound']||'';
      if(tot&&parseInt(tot)>0){document.getElementById('f-total').value=parseInt(tot);filled++;}
      // Wait time
      var wt=r['call wait time (sec)']||r['call wait time']||r['avg call wait time']||'';
      if(wt){var ws=wt.toString().replace(/[^0-9.]/g,'');if(parseFloat(ws)>0){document.getElementById('f-wait').value=Math.round(parseFloat(ws));filled++;}}
      // Call volume
      var cv=r['completed outbound calls']||r['total calls']||r['accepted call legs']||'';
      // Duration
      var dur=r['call talk time (min)']||r['avg call duration']||r['call duration']||'';
      if(dur){var dn=dur.toString().replace(/[^0-9.]/g,'');if(parseFloat(dn)>0){document.getElementById('f-dur').value=parseFloat(dn).toFixed(2);filled++;}}
    });

    if(filled>0){
      st.className='import-status ok';
      st.textContent='✓ Imported '+filled+' values — review and generate';
      generate();
    } else {
      st.className='import-status err';
      st.textContent='Could not auto-detect columns. Fill in manually.';
    }
    input.value='';
  }

  readFile(file,applyData,st);
}

// ── EMAIL BUILDER ─────────────────────────────────────────────────────────────
function buildEmail(data){
  data=data||{
    type:v('f-type'),week:v('f-week'),date:v('f-date'),opening:v('f-opening'),
    calls:v('f-calls'),dur:v('f-dur'),durStatus:durStatus,durComment:v('f-dur-comment'),
    inbound:v('f-inbound'),vm:v('f-vm'),total:v('f-total'),wait:v('f-wait'),
    ans:v('f-ans'),ansComment:v('f-ans-comment'),
    gapRed:v('f-gap-red'),gapYellow:v('f-gap-yellow'),ops:v('f-ops'),
    sender:v('f-sender'),stitle:v('f-stitle'),phone:v('f-phone'),email:v('f-email'),
    tps:getTPs()
  };

  var ac=ansSt(data.ans);var wc=waitSt(data.wait);
  var dc=data.durStatus==='green'?'g':data.durStatus==='amber'?'a':'r';
  var ansLabel=parseFloat(data.ans)>=95?'Exceeds 95% Target':'Below 95% Target';
  var durLabel=data.durStatus==='green'?'At goal (target: &le;15 min)':data.durStatus==='amber'?'Approaching limit (target: &le;15 min)':'Above target (target: &le;15 min)';

  var h='<div class="em-wrap"><div class="em-box">';

  // Header
  h+='<div class="em-head">'+
    '<div class="em-head-row"><div class="em-co">Hugo Technologies</div><div class="em-prog">9amHealth Program</div></div>'+
    '<div class="em-subject">'+data.type+'</div>'+
    '<div class="em-meta">'+data.week+' &nbsp;&middot;&nbsp; '+data.date+'</div>'+
  '</div>';

  h+='<div class="em-body">';
  h+='<div class="em-intro">Hi Team,<br><br>'+data.opening+'</div>';

  // Flags
  if(data.gapRed)h+='<div class="gap-r"><div class="gl">&#9888; Confirmed Gap</div><div class="gt">'+data.gapRed+'</div></div>';
  if(data.gapYellow)h+='<div class="gap-y"><div class="gl">&#9888; Watch Item</div><div class="gt">'+data.gapYellow+'</div></div>';

  // Production
  h+='<div class="sh blue">Production Update &mdash; '+data.date+'</div>';
  h+='<table class="et"><thead><tr><th>Metric</th><th class="wk">'+data.week+'</th></tr></thead><tbody>'+
    '<tr><td>Total calls completed</td><td class="wk">'+data.calls+'</td></tr>'+
    '<tr><td>Avg call duration (talk + wrap-up)</td><td class="wk">'+data.dur+' min</td></tr>'+
    '<tr><td>Call duration vs target</td><td class="'+stClass(dc)+'">'+stIcon(dc)+' '+durLabel+(data.durComment?'<span class="cc">'+data.durComment+'</span>':'')+'</td></tr>'+
  '</tbody></table>';

  // Answer rate
  h+='<div class="sh green">Answer Rate &mdash; '+data.date+'</div>';
  h+='<table class="et"><thead><tr><th>Metric</th><th class="wk">'+data.week+'</th></tr></thead><tbody>'+
    '<tr><td>Inbound answered (no voicemail)</td><td class="wk">'+data.inbound+'</td></tr>'+
    '<tr><td>Voicemail calls</td><td class="wk">'+data.vm+'</td></tr>'+
    '<tr><td>Total relevant inbound calls</td><td class="wk">'+data.total+'</td></tr>'+
    '<tr><td>Call wait time</td><td class="'+stClass(wc)+'">'+stIcon(wc)+' '+data.wait+' sec<span class="cc">'+(parseInt(data.wait)<=30?'Within 30-second threshold':'Above 30-second threshold')+'</span></td></tr>'+
    '<tr><td>Call answer rate</td><td class="'+stClass(ac)+'">'+stIcon(ac)+' '+data.ans+'%<span class="cc">'+(data.ansComment||ansLabel)+'</span></td></tr>'+
  '</tbody></table>';

  // Ops note
  if(data.ops){
    h+='<div class="sh gray">Ops Note</div>';
    h+='<table class="et"><tbody><tr><td style="font-size:13px;line-height:1.65;color:#374151">'+data.ops+'</td></tr></tbody></table>';
  }

  // Touchpoints
  if(data.tps&&data.tps.length){
    h+='<div class="sh navy">Upcoming Touchpoints</div>';
    h+='<table class="et"><thead><tr><th class="wk">Date</th><th>Activity</th><th>Objective</th></tr></thead><tbody>';
    data.tps.forEach(function(tp){
      h+='<tr><td class="wk">'+tp.date+'</td><td><strong>'+tp.act+'</strong></td><td style="color:#6b7280;font-size:12px">'+tp.obj+'</td></tr>';
    });
    h+='</tbody></table>';
  }

  h+='<div class="em-div"></div>';
  h+='<div style="font-size:13px;color:#374151;line-height:1.65">We look forward to continued progress and will keep you updated as the program advances.</div>';
  h+='</div>'; // em-body

  // Footer
  h+='<div class="em-foot">'+
    '<strong>'+data.sender+'</strong>'+
    data.stitle+'<br>'+data.phone+'<br>'+
    '<a href="mailto:'+data.email+'">'+data.email+'</a> &nbsp;|&nbsp; <a href="http://hugotech.co">hugoinc.com</a>'+
    '<div class="em-conf">Hugo Technologies &nbsp;&middot;&nbsp; 9amHealth Program &nbsp;&middot;&nbsp; Confidential</div>'+
  '</div>';

  h+='</div></div>';
  return h;
}

// ── GENERATE & PREVIEW ────────────────────────────────────────────────────────
var generateTimer=null;
function generate(){
  clearTimeout(generateTimer);
  generateTimer=setTimeout(function(){
    var sub='9amHealth Program Update — '+v('f-week')+' | '+v('f-type')+' | '+v('f-date');
    document.getElementById('preview-subject').textContent='Subject: '+sub;
    document.getElementById('preview-output').innerHTML=buildEmail();
  },150);
}

// ── COPY HTML ────────────────────────────────────────────────────────────────
function copyEmail(){
  var sub='9amHealth Program Update — '+v('f-week')+' | '+v('f-type')+' | '+v('f-date');
  var emailHtml=buildEmail();
  var full='<!-- Subject: '+sub+' -->\n'+emailHtml;

  // Save to history
  var entry={
    week:v('f-week'),date:v('f-date'),type:v('f-type'),
    ans:v('f-ans'),wait:v('f-wait'),calls:v('f-calls'),dur:v('f-dur'),
    inbound:v('f-inbound'),vm:v('f-vm'),total:v('f-total'),
    gapRed:v('f-gap-red'),gapYellow:v('f-gap-yellow'),ops:v('f-ops'),
    tps:getTPs(),durStatus:durStatus,durComment:v('f-dur-comment'),
    ansComment:v('f-ans-comment'),opening:v('f-opening'),
    sender:v('f-sender'),stitle:v('f-stitle'),phone:v('f-phone'),email:v('f-email'),
    savedAt:new Date().toISOString()
  };
  history.unshift(entry);
  if(history.length>20)history=history.slice(0,20);
  try{localStorage.setItem('9am-email-history',JSON.stringify(history));}catch(e){}

  // Show copy modal — works reliably in all browsers including local files
  showCopyModal(sub, full, emailHtml);
}

function showCopyModal(subject, fullHtml, emailHtml){
  var existing=document.getElementById("copy-modal");
  if(existing)existing.remove();

  window._copyHtml=fullHtml;
  window._copyEmailHtml=emailHtml;
  window._copySubject=subject;

  var overlay=document.createElement("div");
  overlay.id="copy-modal";
  overlay.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:9999;display:flex;align-items:center;justify-content:center;padding:20px;font-family:inherit";

  var box=document.createElement("div");
  box.style.cssText="background:#0f1923;border:1px solid rgba(255,255,255,.15);max-width:680px;width:100%;max-height:90vh;display:flex;flex-direction:column;overflow:hidden";

  // Header row
  var hdr=document.createElement("div");
  hdr.style.cssText="padding:14px 18px;border-bottom:1px solid rgba(255,255,255,.1);display:flex;align-items:flex-start;justify-content:space-between;gap:10px";
  var hdrLeft=document.createElement("div");
  var hdrTitle=document.createElement("div");
  hdrTitle.textContent="Email Ready to Send";
  hdrTitle.style.cssText="color:#fff;font-size:14px;font-weight:700;margin-bottom:3px";
  var hdrSub=document.createElement("div");
  hdrSub.textContent="Subject: "+subject;
  hdrSub.style.cssText="color:rgba(255,255,255,.4);font-size:11px";
  hdrLeft.appendChild(hdrTitle);
  hdrLeft.appendChild(hdrSub);
  var closeBtn=document.createElement("button");
  closeBtn.textContent="×";
  closeBtn.style.cssText="background:none;border:none;color:rgba(255,255,255,.5);font-size:22px;cursor:pointer;padding:0 4px;line-height:1;flex-shrink:0";
  closeBtn.onclick=function(){overlay.remove();};
  hdr.appendChild(hdrLeft);
  hdr.appendChild(closeBtn);

  // Button area
  var btnArea=document.createElement("div");
  btnArea.style.cssText="padding:14px 18px;border-bottom:1px solid rgba(255,255,255,.1)";

  var instr=document.createElement("div");
  instr.style.cssText="color:rgba(255,255,255,.7);font-size:12px;margin-bottom:12px;line-height:1.7";
  var instrTitle=document.createElement("strong");
  instrTitle.textContent="Choose how to send:";
  instrTitle.style.cssText="color:#fff;display:block;margin-bottom:4px";
  instr.appendChild(instrTitle);
  instr.appendChild(document.createTextNode("① Best: Click Copy HTML Code → open Gmail → Compose → click HTMaiL icon → paste"));
  instr.appendChild(document.createElement("br"));
  instr.appendChild(document.createTextNode("② Easiest: Click Open in New Window → Ctrl+A → Ctrl+C → paste into Gmail compose"));
  instr.appendChild(document.createElement("br"));
  instr.appendChild(document.createTextNode("③ Manual: Click the code box below → Ctrl+A → Ctrl+C"));

  var btnRow=document.createElement("div");
  btnRow.style.cssText="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px";

  function makeBtn(label,bg,fn){
    var b=document.createElement("button");
    b.textContent=label;
    b.style.cssText="background:"+bg+";color:#fff;border:none;padding:8px 14px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit";
    b.onclick=fn;
    return b;
  }
  btnRow.appendChild(makeBtn("Copy HTML Code","#1d4ed8",function(){copyRawHTML();}));
  btnRow.appendChild(makeBtn("Open in New Window","#374151",function(){openPreviewWindow();}));
  btnRow.appendChild(makeBtn("Copy Rendered Email","#166534",function(){copyRendered();}));

  var statusDiv=document.createElement("div");
  statusDiv.id="copy-status";
  statusDiv.style.cssText="font-size:12px;color:#86efac;display:none;margin-top:4px;line-height:1.5";

  btnArea.appendChild(instr);
  btnArea.appendChild(btnRow);
  btnArea.appendChild(statusDiv);

  // Code textarea
  var codeWrap=document.createElement("div");
  codeWrap.style.cssText="padding:12px 18px;flex:1;overflow:auto;min-height:0";
  var codeLabel=document.createElement("div");
  codeLabel.textContent="HTML Source — click to select all, then Ctrl+C to copy";
  codeLabel.style.cssText="color:rgba(255,255,255,.35);font-size:10px;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px";
  var ta=document.createElement("textarea");
  ta.id="copy-textarea";
  ta.value=fullHtml;
  ta.readOnly=true;
  ta.style.cssText="width:100%;height:180px;background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.1);color:rgba(255,255,255,.6);font-size:11px;font-family:monospace;padding:10px;resize:vertical;outline:none;box-sizing:border-box";
  ta.onclick=function(){this.select();};
  codeWrap.appendChild(codeLabel);
  codeWrap.appendChild(ta);

  box.appendChild(hdr);
  box.appendChild(btnArea);
  box.appendChild(codeWrap);
  overlay.appendChild(box);
  document.body.appendChild(overlay);
  setTimeout(function(){ta.focus();ta.select();},100);
}
function copyRawHTML(){
  var ta=document.getElementById('copy-textarea');
  if(ta){ta.value=window._copyHtml;ta.select();try{document.execCommand('copy');}catch(e){}}
  // Also try modern API
  if(navigator.clipboard){navigator.clipboard.writeText(window._copyHtml).catch(function(){});}
  var st=document.getElementById('copy-status');
  if(st){st.style.display='block';st.textContent='✓ HTML code copied! Paste into HTMaiL extension in Gmail compose window.';}
}

function copyRendered(){
  // Create a hidden iframe with the email, select all and copy
  var iframe=document.createElement('iframe');
  iframe.style.cssText='position:fixed;left:-9999px;top:0;width:800px;height:600px;border:none';
  document.body.appendChild(iframe);
  iframe.contentDocument.open();
  iframe.contentDocument.write(window._copyEmailHtml);
  iframe.contentDocument.close();
  setTimeout(function(){
    try{
      iframe.contentWindow.focus();
      iframe.contentDocument.execCommand('selectAll');
      iframe.contentDocument.execCommand('copy');
      var st=document.getElementById('copy-status');
      if(st){st.style.display='block';st.textContent='✓ Rendered email copied! Open Gmail → Compose → paste (Ctrl+V). The formatted email will appear.';}
    }catch(ex){
      var st=document.getElementById('copy-status');
      if(st){st.style.display='block';st.style.color='#fca5a5';st.textContent='Could not copy rendered version. Use "Open in New Window" instead — then Ctrl+A, Ctrl+C from the preview.';}
    }
    document.body.removeChild(iframe);
  },300);
}

function openPreviewWindow(){
  var win=window.open('','_blank','width=700,height=800');
  if(win){
    win.document.open();
    win.document.write(window._copyEmailHtml);
    win.document.close();
    var st=document.getElementById('copy-status');
    if(st){st.style.display='block';st.textContent='Email opened in new window. Press Ctrl+A then Ctrl+C to copy, then paste into Gmail compose.';}
  }
}

// ── HISTORY ──────────────────────────────────────────────────────────────────
function showHistory(){
  var modal=document.getElementById('history-modal');
  modal.style.display='block';
  var tabs=document.getElementById('history-week-tabs');
  var content=document.getElementById('history-content');

  if(!history.length){
    tabs.innerHTML='';
    content.innerHTML='<div style="text-align:center;padding:40px;color:#9ca3af">No saved emails yet. Generate and copy an email to save it to history.</div>';
    return;
  }

  tabs.innerHTML=history.map(function(e,i){
    return '<button class="wt'+(i===0?' active':'')+'" onclick="showHistoryItem('+i+',this)">'+e.week+' &middot; '+e.date+'</button>';
  }).join('');
  showHistoryItem(0,tabs.children[0]);
}

function showHistoryItem(idx,btn){
  document.querySelectorAll('.wt').forEach(function(b){b.classList.remove('active')});
  if(btn)btn.classList.add('active');
  var e=history[idx];
  if(!e)return;
  document.getElementById('history-content').innerHTML=
    '<div style="margin-bottom:10px;font-size:11px;color:#6b7280">Saved '+new Date(e.savedAt).toLocaleString()+'</div>'+
    '<button onclick="loadHistoryItem('+idx+')" style="background:#1d4ed8;color:#fff;border:none;padding:7px 16px;font-size:12px;cursor:pointer;font-family:inherit;margin-bottom:14px">&#8635; Load this entry to edit</button>'+
    buildEmail(e);
}

function loadHistoryItem(idx){
  var e=history[idx];if(!e)return;
  document.getElementById('f-type').value=e.type||'Daily Performance Update';
  document.getElementById('f-week').value=e.week||'';
  document.getElementById('f-date').value=e.date||'';
  document.getElementById('f-opening').value=e.opening||'';
  document.getElementById('f-calls').value=e.calls||'';
  document.getElementById('f-dur').value=e.dur||'';
  document.getElementById('f-dur-comment').value=e.durComment||'';
  document.getElementById('f-inbound').value=e.inbound||'';
  document.getElementById('f-vm').value=e.vm||'';
  document.getElementById('f-total').value=e.total||'';
  document.getElementById('f-wait').value=e.wait||'';
  document.getElementById('f-ans').value=e.ans||'';
  document.getElementById('f-ans-comment').value=e.ansComment||'';
  document.getElementById('f-gap-red').value=e.gapRed||'';
  document.getElementById('f-gap-yellow').value=e.gapYellow||'';
  document.getElementById('f-ops').value=e.ops||'';
  document.getElementById('f-sender').value=e.sender||'';
  document.getElementById('f-stitle').value=e.stitle||'';
  document.getElementById('f-phone').value=e.phone||'';
  document.getElementById('f-email').value=e.email||'';
  durStatus=e.durStatus||'green';
  document.querySelectorAll('.status-btn').forEach(function(b){
    b.classList.toggle('active',b.classList.contains(durStatus));
  });
  document.getElementById('tp-list').innerHTML='';
  (e.tps||[]).forEach(function(tp){addTP(tp.date,tp.act,tp.obj);});
  document.getElementById('history-modal').style.display='none';
  generate();
}

// ── INIT ─────────────────────────────────────────────────────────────────────
defaultTPs.forEach(function(tp){addTP(tp.date,tp.act,tp.obj);});
generate();
</script>
</body>
</html>
