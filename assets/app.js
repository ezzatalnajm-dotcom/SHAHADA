/* ============ تطبيق "شهادة" - المنطق الرئيسي ============ */

const LS_STUDENTS = "shahada_students_v1";
const LS_SETTINGS = "shahada_settings_v1";
const LS_SENTLOG  = "shahada_sentlog_v1";

let STATE = {
  students: [],
  settings: { teacherName: "عزت شعبان", sheetUrl: "", theme: "dark" },
  sentLog: {},          // { studentId: true }  -> أُرسلت شهادته
  currentView: "home",
  cert: { studentId: null, picks: {}, versions: {} }, // picks[crit] = level or null, versions[crit] = idx
  queue: { list: [], idx: 0, active: false, target: "student" },
  broadcast: { text: "", statusFilter: "2", group: "", location: "", target: "student", list: [], idx: 0, active: false }
};

/* ---------- تخزين ---------- */
function loadAll(){
  try{ STATE.students = JSON.parse(localStorage.getItem(LS_STUDENTS) || "[]"); }catch(e){ STATE.students = []; }
  try{ STATE.settings = Object.assign({teacherName:"عزت شعبان", sheetUrl:"", theme:"dark"}, JSON.parse(localStorage.getItem(LS_SETTINGS) || "{}")); }catch(e){}
  try{ STATE.sentLog = JSON.parse(localStorage.getItem(LS_SENTLOG) || "{}"); }catch(e){ STATE.sentLog = {}; }
}
function saveStudents(){ localStorage.setItem(LS_STUDENTS, JSON.stringify(STATE.students)); }
function saveSettings(){ localStorage.setItem(LS_SETTINGS, JSON.stringify(STATE.settings)); }
function saveSentLog(){ localStorage.setItem(LS_SENTLOG, JSON.stringify(STATE.sentLog)); }

/* ---------- أدوات نص ---------- */
function normalizeAr(s){
  if(!s) return "";
  return String(s)
    .replace(/[إأآا]/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/ة/g, "ه")
    .replace(/ؤ/g, "و")
    .replace(/ئ/g, "ي")
    .replace(/[ًٌٍَُِّْ]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}
function toast(msg){
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(toast._tm);
  toast._tm = setTimeout(()=> t.classList.remove("show"), 2200);
}
function statusBadge(status){
  const s = String(status ?? "").trim();
  if(s === "2") return '<span class="badge s2">حالة 2</span>';
  if(s === "1") return '<span class="badge s1">حالة 1</span>';
  if(s === "0") return '<span class="badge s0">حالة 0</span>';
  return '<span class="badge s0">بدون حالة</span>';
}

/* ---------- رفع الإكسيل ---------- */
function handleFile(file){
  const reader = new FileReader();
  reader.onload = (e)=>{
    try{
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, {type:"array"});
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, {header:1, raw:false, defval:""});
      if(!rows.length){ toast("⚠️ الملف فارغ"); return; }
      // نتخطى صف العناوين (أول صف)
      const dataRows = rows.slice(1).filter(r => r.some(c => String(c).trim() !== ""));
      const students = dataRows.map((r, i)=>{
        const obj = { id: "s"+i };
        COLUMNS.forEach((key, idx)=>{ obj[key] = (r[idx] ?? "").toString().trim(); });
        return obj;
      }).filter(s => s.name); // لازم اسم طالب على الأقل
      STATE.students = students;
      saveStudents();
      toast(`✅ تم رفع ${students.length} طالب/ة بنجاح`);
      renderHome();
      renderCertStudentPicker();
    }catch(err){
      console.error(err);
      toast("❌ تعذّرت قراءة الملف، تأكد إنه بصيغة xlsx");
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ---------- مزامنة مع Google Sheet (اختياري) ---------- */
async function syncFromSheet(){
  // لو المستخدم لسه واقف في شاشة الإعدادات وكاتب رابط بس مادوسش حفظ، ناخده مباشرة من الحقل
  const fieldEl = document.getElementById("sheetUrl");
  if(fieldEl && fieldEl.value.trim()){
    STATE.settings.sheetUrl = fieldEl.value.trim();
    saveSettings();
  }
  const url = (STATE.settings.sheetUrl || "").trim();
  if(!url){ toast("⚠️ حط رابط ربط الشيت الأول من الإعدادات"); return false; }
  try{
    toast("⏳ بيتزامن مع الشيت...");
    const res = await fetch(url, { method:"GET" });
    const json = await res.json();
    if(!json.ok){ toast("❌ " + (json.error || "فشلت المزامنة")); return false; }
    STATE.students = json.students.map(s=>({
      id: "r" + s.rowIndex,
      rowIndex: s.rowIndex,
      phoneStudent: s.phoneStudent, location: s.location, group: s.group,
      booking: s.booking, guardianRel: s.guardianRel, gender: s.gender,
      level: s.level, status: s.status, phoneGuardian: s.phoneGuardian, name: s.name,
      sentStudent: s.sentStudent, sentGuardian: s.sentGuardian
    }));
    saveStudents();
    toast(`✅ اتزامن ${STATE.students.length} طالب/ة من الشيت`);
    return true;
  }catch(err){
    console.error(err);
    toast("❌ تعذّر الاتصال بالشيت، تأكد من الرابط والنشر");
    return false;
  }
}

function markSentOnSheet(student, target){
  const url = (STATE.settings.sheetUrl || "").trim();
  if(!url || !student.rowIndex) return; // مفيش ربط شيت، أو الطالب من رفع إكسيل محلي مش من الشيت
  fetch(url, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action:"markSent", rowIndex: student.rowIndex, target })
  }).catch(()=>{ /* أفضل جهد بس، متعملش مشكلة لو فشلت */ });
}

/* ---------- بناء الرسالة ---------- */
function pickRandomVersionIdx(exclude){
  const n = 4;
  if(exclude === undefined) return Math.floor(Math.random()*n);
  let idx;
  do{ idx = Math.floor(Math.random()*n); }while(idx === exclude && n > 1);
  return idx;
}

function buildCertMessage(student, picks, versions){
  const name = student.name || "";
  let msg = `بسم الله معك أ.${STATE.settings.teacherName} معلم الأحياء💫\nبالنسبة للطالب/ة: ${name}\n`;
  CRITERIA_ORDER.forEach(critKey=>{
    const levelKey = picks[critKey];
    if(!levelKey) return; // تخطّاه المعلم
    const crit = CRITERIA[critKey];
    const level = crit.levels.find(l => l.key === Number(levelKey));
    if(!level) return;
    const vIdx = versions[critKey] ?? 0;
    const phrase = level.versions[vIdx] || level.versions[0];
    msg += `\n📌 ${crit.label}\n${phrase} ${level.emoji}\n`;
  });
  return msg.trim();
}

function waLink(phone, text){
  const clean = String(phone || "").replace(/[^\d]/g,"");
  return `https://wa.me/${clean}?text=${encodeURIComponent(text)}`;
}

/* ============ التنقل بين الشاشات ============ */
function switchView(name){
  STATE.currentView = name;
  document.querySelectorAll(".view").forEach(v=> v.classList.toggle("active", v.id === "view-"+name));
  document.querySelectorAll(".tab").forEach(t=> t.classList.toggle("active", t.dataset.view === name));
  if(name === "home") renderHome();
  if(name === "cert") renderCertStudentPicker();
  if(name === "broadcast") renderBroadcast();
  if(name === "settings") renderSettings();
}

/* ============ الرئيسية ============ */
function renderHome(){
  const el = document.getElementById("view-home");
  const total = STATE.students.length;
  const eligible = STATE.students.filter(s => String(s.status).trim() === "2").length;
  const sentCount = Object.keys(STATE.sentLog).length;
  el.innerHTML = `
    <div class="card">
      <h2>📤 رفع بيانات الطلاب</h2>
      <p class="hint">ارفع ملف الإكسيل بترتيب الأعمدة (A:J): هاتف الطالب، مكان السكن، اسم المجموعة، موعد الحجز، صفة ولي الأمر، النوع، المستوى، الحالة، هاتف ولي الأمر، اسم الطالب.</p>
      <input type="file" id="fileInput" accept=".xlsx,.xls,.xlsm" style="display:none">
      <button class="btn btn-gold" id="btnUpload">📁 اختيار ملف الإكسيل</button>
      ${STATE.settings.sheetUrl ? `<button class="btn btn-outline" style="margin-top:10px" onclick="syncFromSheet().then(()=>renderHome())">🔄 أو زامن من Google Sheet</button>` : `<p class="hint" style="margin-top:10px">تقدر كمان تربط Google Sheet من الإعدادات ⚙️ عشان تشتغل من أكتر من جهاز بنفس البيانات.</p>`}
    </div>
    <div class="card">
      <h2>📊 نظرة سريعة</h2>
      <div class="kv"><span>إجمالي الطلاب المرفوعين</span><b>${total}</b></div>
      <div class="kv"><span>مؤهلون لإرسال الشهادة (حالة 2)</span><b>${eligible}</b></div>
      <div class="kv"><span>شهادات تم إرسالها فعلاً</span><b>${sentCount}</b></div>
    </div>
    ${total===0 ? `<div class="empty"><span class="em-ic">🗂️</span>مفيش بيانات لسه.. ارفع ملف الإكسيل عشان تبدأ</div>` : `
    <div class="card">
      <h2>⚡ إجراء سريع</h2>
      <div class="row">
        <button class="btn btn-outline" onclick="switchView('cert')">✍️ كتابة شهادة</button>
        <button class="btn btn-outline" onclick="switchView('broadcast')">📢 رسالة عامة</button>
      </div>
    </div>`}
  `;
  document.getElementById("btnUpload").onclick = ()=> document.getElementById("fileInput").click();
  document.getElementById("fileInput").onchange = (e)=>{ if(e.target.files[0]) handleFile(e.target.files[0]); };
}

/* ============ شاشة الشهادات ============ */
function renderCertStudentPicker(){
  const el = document.getElementById("view-cert");
  if(!STATE.students.length){
    el.innerHTML = `<div class="empty"><span class="em-ic">📄</span>لازم ترفع ملف الإكسيل الأول من شاشة "الرئيسية".</div>`;
    return;
  }
  if(STATE.queue.active){ renderQueueMode(); return; }
  if(STATE.cert.studentId){ renderCertEditor(); return; }

  el.innerHTML = `
    <div class="card">
      <h2>✍️ إنشاء شهادة</h2>
      <div class="field">
        <input type="search" id="certSearch" placeholder="🔎 دوّر باسم الطالب...">
      </div>
      <div class="row" style="margin-bottom:12px">
        <button class="btn btn-green btn-sm" onclick="startQueueMode()">▶️ وضع الإرسال المتتالي (حالة 2)</button>
      </div>
      <div id="certList"></div>
    </div>
  `;
  const listEl = document.getElementById("certList");
  function draw(filter){
    const f = normalizeAr(filter||"");
    const items = STATE.students.filter(s => !f || normalizeAr(s.name).includes(f));
    if(!items.length){ listEl.innerHTML = `<div class="empty">مفيش نتائج</div>`; return; }
    listEl.innerHTML = items.slice(0,150).map(s=>`
      <div class="student-item" onclick="openCert('${s.id}')">
        <div><div class="nm">${s.name}</div><div class="meta">${s.group || "بدون مجموعة"}</div></div>
        ${statusBadge(s.status)}
      </div>`).join("");
  }
  draw("");
  document.getElementById("certSearch").oninput = (e)=> draw(e.target.value);
}

function openCert(id){
  STATE.cert.studentId = id;
  STATE.cert.picks = {}; STATE.cert.versions = {};
  renderCertEditor();
}

function renderCertEditor(){
  const el = document.getElementById("view-cert");
  const student = STATE.students.find(s=> s.id === STATE.cert.studentId);
  if(!student){ STATE.cert.studentId = null; renderCertStudentPicker(); return; }

  const blocksHtml = CRITERIA_ORDER.map(critKey=>{
    const crit = CRITERIA[critKey];
    const picked = STATE.cert.picks[critKey] || "";
    const options = crit.levels.map(l=> `<option value="${l.key}" ${String(picked)===String(l.key)?"selected":""}>${l.emoji} ${l.tag}</option>`).join("");
    return `
      <div class="crit-block">
        <div class="crit-head">
          <b>${crit.icon} ${crit.label}</b>
          <div class="crit-actions">
            ${picked ? `<span class="chip" onclick="shuffleVersion('${critKey}')">🔀 نسخة تانية</span>` : ""}
          </div>
        </div>
        <select onchange="setCritLevel('${critKey}', this.value)">
          <option value="">— تخطّي (لا يظهر بالرسالة) —</option>
          ${options}
        </select>
      </div>`;
  }).join("");

  const msg = buildCertMessage(student, STATE.cert.picks, STATE.cert.versions);
  const alreadySent = !!STATE.sentLog[student.id];

  el.innerHTML = `
    <div class="card">
      <div class="row" style="margin-bottom:10px">
        <button class="btn btn-outline btn-sm" onclick="closeCert()">↩️ رجوع للقائمة</button>
      </div>
      <h2>👤 ${student.name} ${alreadySent ? "· <small style=\"color:#81c784\">تم الإرسال قبل كده ✓</small>" : ""}</h2>
      ${blocksHtml}
    </div>
    <div class="card">
      <h2>💬 معاينة الرسالة</h2>
      <div class="wa-preview">${escapeHtml(msg).replace(/\n/g,"<br>")}</div>
    </div>
    <div class="card">
      <h2>📤 إرسال</h2>
      <p class="hint">هيفتح واتساب بالرسالة جاهزة، انت بس تدوس إرسال جوه واتساب.</p>
      <div class="row">
        <button class="btn btn-green" ${!student.phoneStudent?"disabled":""} onclick="sendCert('student')">👦 إرسال للطالب</button>
        <button class="btn btn-blue" ${!student.phoneGuardian?"disabled":""} onclick="sendCert('guardian')">👪 إرسال لولي الأمر</button>
      </div>
    </div>
  `;
}
function escapeHtml(s){ return s.replace(/[&<>]/g, c=> ({"&":"&amp;","<":"&lt;",">":"&gt;"}[c])); }

function setCritLevel(critKey, val){
  if(!val){ delete STATE.cert.picks[critKey]; delete STATE.cert.versions[critKey]; }
  else{
    STATE.cert.picks[critKey] = val;
    if(STATE.cert.versions[critKey] === undefined) STATE.cert.versions[critKey] = pickRandomVersionIdx();
  }
  renderCertEditor();
}
function shuffleVersion(critKey){
  STATE.cert.versions[critKey] = pickRandomVersionIdx(STATE.cert.versions[critKey]);
  renderCertEditor();
}
function closeCert(){ STATE.cert.studentId = null; renderCertStudentPicker(); }

function sendCert(target){
  const student = STATE.students.find(s=> s.id === STATE.cert.studentId);
  if(!student) return;
  const msg = buildCertMessage(student, STATE.cert.picks, STATE.cert.versions);
  const phone = target === "student" ? student.phoneStudent : student.phoneGuardian;
  if(!phone){ toast("⚠️ مفيش رقم متاح"); return; }
  window.open(waLink(phone, msg), "_blank");
  STATE.sentLog[student.id] = true;
  saveSentLog();
  markSentOnSheet(student, target);
  toast("✅ اتفتحت الرسالة في واتساب");
  renderCertEditor();
}

/* ============ وضع الإرسال المتتالي ============ */
function startQueueMode(){
  const list = STATE.students.filter(s => String(s.status).trim() === "2" && !STATE.sentLog[s.id] && !s.sentStudent);
  if(!list.length){ toast("🎉 كل طلاب الحالة 2 اتبعتلهم الشهادة بالفعل"); return; }
  STATE.queue = { list, idx: 0, active: true };
  STATE.cert.studentId = list[0].id; STATE.cert.picks = {}; STATE.cert.versions = {};
  renderQueueMode();
}
function stopQueueMode(){ STATE.queue.active = false; renderCertStudentPicker(); }

function renderQueueMode(){
  const el = document.getElementById("view-cert");
  const { list, idx } = STATE.queue;
  if(idx >= list.length){
    el.innerHTML = `<div class="card"><h2>🎉 خلصنا!</h2><p class="hint">تم المرور على كل الطلاب المستحقين (حالة 2).</p>
      <button class="btn btn-gold" onclick="stopQueueMode()">رجوع</button></div>`;
    return;
  }
  const student = list[idx];
  STATE.cert.studentId = student.id;
  const pct = Math.round((idx/list.length)*100);

  const blocksHtml = CRITERIA_ORDER.map(critKey=>{
    const crit = CRITERIA[critKey];
    const picked = STATE.cert.picks[critKey] || "";
    const options = crit.levels.map(l=> `<option value="${l.key}" ${String(picked)===String(l.key)?"selected":""}>${l.emoji} ${l.tag}</option>`).join("");
    return `
      <div class="crit-block">
        <div class="crit-head">
          <b>${crit.icon} ${crit.label}</b>
          ${picked ? `<span class="chip" onclick="shuffleVersion('${critKey}')">🔀 نسخة تانية</span>` : ""}
        </div>
        <select onchange="setCritLevel('${critKey}', this.value)">
          <option value="">— تخطّي —</option>
          ${options}
        </select>
      </div>`;
  }).join("");

  const msg = buildCertMessage(student, STATE.cert.picks, STATE.cert.versions);

  el.innerHTML = `
    <div class="card">
      <div class="progress-wrap"><div class="progress-bar" style="width:${pct}%"></div></div>
      <div class="progress-txt">طالب ${idx+1} من ${list.length}</div>
    </div>
    <div class="card">
      <h2>👤 ${student.name}</h2>
      ${blocksHtml}
    </div>
    <div class="card">
      <h2>💬 معاينة</h2>
      <div class="wa-preview">${escapeHtml(msg).replace(/\n/g,"<br>")}</div>
    </div>
    <div class="card">
      <div class="row" style="margin-bottom:10px">
        <button class="btn btn-green" ${!student.phoneStudent?"disabled":""} onclick="queueSend('student')">👦 إرسال للطالب والتالي</button>
        <button class="btn btn-blue" ${!student.phoneGuardian?"disabled":""} onclick="queueSend('guardian')">👪 إرسال لولي الأمر والتالي</button>
      </div>
      <div class="row">
        <button class="btn btn-outline btn-sm" onclick="queueSkip()">⏭️ تخطّي بدون إرسال</button>
        <button class="btn btn-red btn-sm" onclick="stopQueueMode()">⏸️ إيقاف مؤقت</button>
      </div>
    </div>
  `;
}
function queueAdvance(){
  STATE.queue.idx += 1;
  STATE.cert.picks = {}; STATE.cert.versions = {};
  renderQueueMode();
}
function queueSend(target){
  const student = STATE.students.find(s=> s.id === STATE.cert.studentId);
  const phone = target === "student" ? student.phoneStudent : student.phoneGuardian;
  if(!phone){ toast("⚠️ مفيش رقم متاح، تقدر تتخطاه"); return; }
  const msg = buildCertMessage(student, STATE.cert.picks, STATE.cert.versions);
  window.open(waLink(phone, msg), "_blank");
  STATE.sentLog[student.id] = true; saveSentLog();
  markSentOnSheet(student, target);
  queueAdvance();
}
function queueSkip(){ queueAdvance(); }

/* ============ رسالة عامة (بث) ============ */
function computeBroadcastList(){
  const b = STATE.broadcast;
  const locNorm = normalizeAr(b.location);
  return STATE.students.filter(s=>{
    if(b.statusFilter !== "all" && String(s.status).trim() !== b.statusFilter) return false;
    if(b.group && s.group !== b.group) return false;
    if(locNorm && !normalizeAr(s.location).includes(locNorm)) return false;
    const phone = b.target === "student" ? s.phoneStudent : s.phoneGuardian;
    return !!phone;
  });
}

function renderBroadcast(){
  const el = document.getElementById("view-broadcast");
  if(!STATE.students.length){
    el.innerHTML = `<div class="empty"><span class="em-ic">📢</span>لازم ترفع ملف الإكسيل الأول.</div>`;
    return;
  }
  if(STATE.broadcast.active){ renderBroadcastQueue(); return; }

  const groups = [...new Set(STATE.students.map(s=> s.group).filter(Boolean))];
  const b = STATE.broadcast;
  const count = computeBroadcastList().length;

  el.innerHTML = `
    <div class="card">
      <h2>📢 رسالة عامة (مش شهادات)</h2>
      <div class="field">
        <label class="field-label">نص الرسالة</label>
        <textarea id="bcText" rows="6" placeholder="اكتب رسالتك هنا...">${b.text}</textarea>
      </div>
      <div class="field">
        <label class="field-label">إرسال إلى</label>
        <div class="filter-chips">
          <button class="fchip ${b.target==='student'?'active':''}" onclick="setBcTarget('student')">👦 الطالب</button>
          <button class="fchip ${b.target==='guardian'?'active':''}" onclick="setBcTarget('guardian')">👪 ولي الأمر</button>
        </div>
      </div>
      <div class="field">
        <label class="field-label">فلترة بالحالة</label>
        <div class="filter-chips">
          <button class="fchip ${b.statusFilter==='all'?'active':''}" onclick="setBcStatus('all')">الكل</button>
          <button class="fchip ${b.statusFilter==='0'?'active':''}" onclick="setBcStatus('0')">حالة 0</button>
          <button class="fchip ${b.statusFilter==='1'?'active':''}" onclick="setBcStatus('1')">حالة 1</button>
          <button class="fchip ${b.statusFilter==='2'?'active':''}" onclick="setBcStatus('2')">حالة 2</button>
        </div>
      </div>
      <div class="field">
        <label class="field-label">مجموعة معيّنة (اختياري)</label>
        <select id="bcGroup" onchange="setBcGroup(this.value)">
          <option value="">كل المجموعات</option>
          ${groups.map(g=> `<option value="${g}" ${b.group===g?"selected":""}>${g}</option>`).join("")}
        </select>
      </div>
      <div class="field">
        <label class="field-label">مكان السكن يحتوي على (اختياري، بيتجاوز الهمزات تلقائي)</label>
        <input type="text" id="bcLoc" value="${b.location}" placeholder="مثال: الريان">
      </div>
      <div class="divider"></div>
      <div class="kv"><span>عدد المستلمين المطابقين</span><b>${count}</b></div>
      <button class="btn btn-gold" style="margin-top:10px" onclick="startBroadcastQueue()">▶️ بدء الإرسال</button>
    </div>
  `;
  document.getElementById("bcText").oninput = (e)=>{ STATE.broadcast.text = e.target.value; };
  document.getElementById("bcLoc").oninput = (e)=>{ STATE.broadcast.location = e.target.value; renderBroadcast(); };
}
function setBcTarget(t){ STATE.broadcast.target = t; renderBroadcast(); }
function setBcStatus(s){ STATE.broadcast.statusFilter = s; renderBroadcast(); }
function setBcGroup(g){ STATE.broadcast.group = g; renderBroadcast(); }

function startBroadcastQueue(){
  const text = (document.getElementById("bcText")?.value || STATE.broadcast.text || "").trim();
  if(!text){ toast("⚠️ اكتب نص الرسالة الأول"); return; }
  STATE.broadcast.text = text;
  const list = computeBroadcastList();
  if(!list.length){ toast("⚠️ مفيش مستلمين مطابقين للفلترة"); return; }
  STATE.broadcast.list = list;
  STATE.broadcast.idx = 0;
  STATE.broadcast.active = true;
  renderBroadcastQueue();
}
function stopBroadcastQueue(){ STATE.broadcast.active = false; renderBroadcast(); }

function renderBroadcastQueue(){
  const el = document.getElementById("view-broadcast");
  const b = STATE.broadcast;
  if(b.idx >= b.list.length){
    el.innerHTML = `<div class="card"><h2>🎉 خلصنا الإرسال</h2><button class="btn btn-gold" onclick="stopBroadcastQueue()">رجوع</button></div>`;
    return;
  }
  const student = b.list[b.idx];
  const phone = b.target === "student" ? student.phoneStudent : student.phoneGuardian;
  const pct = Math.round((b.idx/b.list.length)*100);
  el.innerHTML = `
    <div class="card">
      <div class="progress-wrap"><div class="progress-bar" style="width:${pct}%"></div></div>
      <div class="progress-txt">${b.idx+1} من ${b.list.length}</div>
    </div>
    <div class="card">
      <h2>👤 ${student.name}</h2>
      <div class="kv"><span>المرسل إليه</span><b>${b.target==='student'?'الطالب':'ولي الأمر'}</b></div>
      <div class="wa-preview">${escapeHtml(b.text).replace(/\n/g,"<br>")}</div>
    </div>
    <div class="card">
      <div class="row" style="margin-bottom:10px">
        <button class="btn btn-green" ${!phone?"disabled":""} onclick="broadcastSend()">📤 إرسال والتالي</button>
      </div>
      <div class="row">
        <button class="btn btn-outline btn-sm" onclick="broadcastSkip()">⏭️ تخطّي</button>
        <button class="btn btn-red btn-sm" onclick="stopBroadcastQueue()">⏸️ إيقاف مؤقت</button>
      </div>
    </div>
  `;
}
function broadcastSend(){
  const b = STATE.broadcast;
  const student = b.list[b.idx];
  const phone = b.target === "student" ? student.phoneStudent : student.phoneGuardian;
  window.open(waLink(phone, b.text), "_blank");
  b.idx += 1;
  renderBroadcastQueue();
}
function broadcastSkip(){ STATE.broadcast.idx += 1; renderBroadcastQueue(); }

/* ============ الإعدادات ============ */
let deferredInstallPrompt = null;
function renderSettings(){
  const el = document.getElementById("view-settings");
  el.innerHTML = `
    <div class="card">
      <h2>👨‍🏫 اسم المعلم</h2>
      <p class="hint">هيظهر في بداية كل رسالة: "بسم الله معك أ.[الاسم] معلم الأحياء"</p>
      <input type="text" id="teacherName" value="${STATE.settings.teacherName}">
      <button class="btn btn-gold" style="margin-top:10px" onclick="saveTeacherName()">💾 حفظ الاسم</button>
    </div>
    <div class="card">
      <h2>🔗 ربط Google Sheet (اختياري)</h2>
      <p class="hint">لو ربطت رابط النشر بتاع الشيت، هتقدر تزامن نفس البيانات بين أكتر من جهاز (لاب توب وموبايل). سيبه فاضي لو هتشتغل بالإكسيل المحلي بس.</p>
      <input type="text" id="sheetUrl" value="${STATE.settings.sheetUrl || ""}" placeholder="https://script.google.com/macros/s/.../exec">
      <div class="row" style="margin-top:10px">
        <button class="btn btn-outline btn-sm" onclick="saveSheetUrl()">💾 حفظ الرابط</button>
        <button class="btn btn-green btn-sm" onclick="syncFromSheet()">🔄 مزامنة الآن</button>
      </div>
    </div>
    <div class="card">
      <h2>📲 التطبيق</h2>
      <div class="row" style="margin-bottom:10px">
        <button class="btn btn-outline" id="btnInstall">⬇️ تثبيت التطبيق</button>
        <button class="btn btn-outline" id="btnShare">🔗 مشاركة التطبيق</button>
      </div>
      <div class="row">
        <button class="btn btn-outline btn-sm" id="btnSync">📁 رفع إكسيل محلي</button>
        <button class="btn btn-red btn-sm" id="btnReset">🗑️ تصفير التطبيق</button>
      </div>
    </div>
    <div class="card">
      <h2>ℹ️ معلومات</h2>
      <div class="kv"><span>عدد الطلاب المخزّنين</span><b>${STATE.students.length}</b></div>
      <div class="kv"><span>عدد الشهادات المُرسلة</span><b>${Object.keys(STATE.sentLog).length}</b></div>
    </div>
  `;
  document.getElementById("btnInstall").onclick = doInstall;
  document.getElementById("btnShare").onclick = doShare;
  document.getElementById("btnSync").onclick = ()=> switchView("home");
  document.getElementById("btnReset").onclick = doReset;
  // حفظ تلقائي للرابط أول ما يكتب، عشان مايحتاجش يدوس زرار حفظ منفصل
  document.getElementById("sheetUrl").oninput = (e)=>{
    STATE.settings.sheetUrl = e.target.value.trim();
    saveSettings();
  };
}
function saveSheetUrl(){
  const v = document.getElementById("sheetUrl").value.trim();
  STATE.settings.sheetUrl = v;
  saveSettings();
  toast("✅ اتحفظ رابط الشيت");
}
function saveTeacherName(){
  const v = document.getElementById("teacherName").value.trim();
  if(!v) return;
  STATE.settings.teacherName = v;
  saveSettings();
  toast("✅ اتحفظ اسم المعلم");
}
function doInstall(){
  if(deferredInstallPrompt){
    deferredInstallPrompt.prompt();
    deferredInstallPrompt.userChoice.then(()=>{ deferredInstallPrompt = null; });
  }else{
    toast("📲 من متصفحك: قائمة المتصفح ← إضافة إلى الشاشة الرئيسية");
  }
}
function doShare(){
  const url = location.href.split("#")[0];
  if(navigator.share){
    navigator.share({ title:"شهادة - تطبيق أ.عزت شعبان", text:"تطبيق إرسال شهادات منتصف الفصل", url });
  }else{
    navigator.clipboard?.writeText(url);
    toast("🔗 اتنسخ رابط التطبيق");
  }
}
function doReset(){
  if(!confirm("متأكد إنك عايز تصفّر كل البيانات (الطلاب وسجل الإرسال)؟ الخطوة دي مش هترجع.")) return;
  localStorage.removeItem(LS_STUDENTS);
  localStorage.removeItem(LS_SENTLOG);
  STATE.students = []; STATE.sentLog = {};
  toast("✅ اتصفّر التطبيق");
  switchView("home");
}

/* ---------- المود (نهاري/ليلي) ---------- */
function applyTheme(){
  const isLight = STATE.settings.theme === "light";
  document.body.classList.toggle("light-mode", isLight);
  const btn = document.getElementById("btnTheme");
  if(btn) btn.textContent = isLight ? "☀️" : "🌙";
}
function toggleTheme(){
  STATE.settings.theme = STATE.settings.theme === "light" ? "dark" : "light";
  saveSettings();
  applyTheme();
}

/* ============ تشغيل ============ */
window.addEventListener("beforeinstallprompt", (e)=>{
  e.preventDefault();
  deferredInstallPrompt = e;
});

document.addEventListener("DOMContentLoaded", ()=>{
  loadAll();
  applyTheme();
  document.getElementById("btnTheme").onclick = toggleTheme;
  document.querySelectorAll(".tab").forEach(t=> t.onclick = ()=> switchView(t.dataset.view));
  switchView("home");
  if("serviceWorker" in navigator){
    navigator.serviceWorker.register("./sw.js").catch(()=>{});
  }
});
