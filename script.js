"use strict";

/* ============================================================
   BITS DIGITAL CODEFORGE V1.0
   Single-file grading console
   ============================================================ */

const GRADES = ["A","A-","B","B-","C","C-","D","E"];
const DEFAULTS = {
  "A":[80,100], "A-":[70,79], "B":[60,69], "B-":[50,59],
  "C":[40,49], "C-":[30,39], "D":[20,29], "E":[0,19]
};

let data = [];
let selectedCourse = "";
let finalizeCount = 0;
let gradingStartTime = Date.now();
let timerInterval = null;
let previousSummary = {};
let lastFileName = "";
let searchTerm = "";

const $ = id => document.getElementById(id);

function escapeHTML(value){
  return String(value ?? "").replace(/[&<>"']/g, ch => ({
    "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"
  }[ch]));
}

function normalizeHeader(value){
  return String(value ?? "").trim().toLowerCase().replace(/\s+/g," ");
}

function showStatus(message,type="success"){
  const el = $("status");
  el.className = "status show " + type;
  el.textContent = message;
}

function hideStatus(){
  $("status").className = "status";
  $("status").textContent = "";
}

function toast(message){
  const el = $("toast");
  el.textContent = message;
  el.classList.add("show");
  clearTimeout(toast._t);
  toast._t = setTimeout(()=>el.classList.remove("show"),2800);
}

function csvEscape(value){
  const s = String(value ?? "");
  return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g,'""') + '"' : s;
}

function downloadBlob(content, filename, type){
  const blob = new Blob([content],{type});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
}

function resetCourseOptions(message="Upload a workbook to populate courses"){
  const course = $("course");
  course.innerHTML = `<option value="">${escapeHTML(message)}</option>`;
  course.disabled = false;
  selectedCourse = "";
  if($("courseHelp")) $("courseHelp").textContent = "Course options appear automatically after a valid Excel file is loaded.";
}

function getCourses(){
  return [...new Set(data.map(d=>d.Course).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
}

function validateRows(rows){
  if(!rows.length) return {ok:false,message:"The workbook contains no data rows."};

  const keys = Object.keys(rows[0] || {});
  const normalized = new Map(keys.map(k=>[normalizeHeader(k),k]));
  const required = ["bits id","course","total marks"];
  const missing = required.filter(k=>!normalized.has(k));
  if(missing.length){
    return {ok:false,message:"Missing required column(s): " + missing.map(x=>x==="bits id"?"BITS ID":x==="course"?"Course":"Total Marks").join(", ") + "."};
  }

  const idKey = normalized.get("bits id");
  const courseKey = normalized.get("course");
  const marksKey = normalized.get("total marks");

  const cleaned = [];
  const seen = new Set();

  for(let i=0;i<rows.length;i++){
    const raw = rows[i];
    const rowNo = i + 2;
    const id = String(raw[idKey] ?? "").trim();
    const course = String(raw[courseKey] ?? "").trim();
    const rawMarks = raw[marksKey];

    if(!id) return {ok:false,message:`Row ${rowNo}: BITS ID is blank.`};
    if(!course) return {ok:false,message:`Row ${rowNo}: Course is blank.`};
    if(rawMarks === "" || rawMarks === null || rawMarks === undefined) {
      return {ok:false,message:`Row ${rowNo}: Total Marks is blank.`};
    }

    const marks = Number(rawMarks);
    if(!Number.isFinite(marks)) return {ok:false,message:`Row ${rowNo}: Total Marks must be numeric.`};
    if(!Number.isInteger(marks)) return {ok:false,message:`Row ${rowNo}: Total Marks must be a whole number.`};
    if(marks < 0 || marks > 100) return {ok:false,message:`Row ${rowNo}: Total Marks must be between 0 and 100.`};

    const duplicateKey = id.toUpperCase() + "||" + course.toUpperCase();
    if(seen.has(duplicateKey)){
      return {ok:false,message:`Row ${rowNo}: duplicate BITS ID "${id}" found for course "${course}".`};
    }
    seen.add(duplicateKey);

    cleaned.push({ "BITS ID":id, Course:course, "Total Marks":marks });
  }

  return {ok:true,rows:cleaned};
}

async function handleFile(file){
  if(!file) return;
  const extension = file.name.toLowerCase().split(".").pop();
  if(!["xlsx","xls"].includes(extension)){
    showStatus("Please upload an Excel workbook with a .xlsx or .xls extension.","error");
    return;
  }

  try{
    hideStatus();
    $("fileName").textContent = file.name;
    $("datasetBadge").textContent = "Reading workbook…";

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer,{type:"array"});
    if(!workbook.SheetNames.length) throw new Error("The workbook has no worksheets.");

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet,{defval:"",raw:true});
    const result = validateRows(rows);

    if(!result.ok){
      data = [];
      resetCourseOptions();
      $("studentSearch").disabled = true;
      $("studentSearch").value = "";
      searchTerm = "";
      $("datasetBadge").textContent = "Invalid dataset";
      showStatus(result.message,"error");
      renderEmptyState();
      return;
    }

    data = result.rows;
    lastFileName = file.name;
    resetCourseOptions();

    const courseSelect = $("course");
    const courses = getCourses();
    courses.forEach(c => courseSelect.add(new Option(c,c)));
    courseSelect.disabled = false;
    $("studentSearch").disabled = false;
    $("datasetBadge").textContent = `${data.length} records • ${courses.length} course${courses.length===1?"":"s"}`;

    showStatus(`✓ Loaded ${data.length} valid student records across ${courses.length} course${courses.length===1?"":"s"}. Select a course to continue.`,"success");
    renderEmptyState();
    updateAll();
  }catch(err){
    data = [];
    resetCourseOptions();
    $("datasetBadge").textContent = "Upload failed";
    showStatus("Could not read this workbook. Please verify it is a valid Excel file.","error");
    console.error(err);
  }
}

function getCourseRows(){
  return data.filter(d=>d.Course === selectedCourse);
}

function gradeForMark(mark){
  for(const g of GRADES){
    const [min,max] = getRange(g);
    if(mark >= min && mark <= max) return g;
  }
  return "";
}

function getRange(g){
  return [
    Number($(g+"min")?.value),
    Number($(g+"max")?.value)
  ];
}

function validateRanges(){
  if(!data.length || !selectedCourse){
    return {ok:true,message:""};
  }

  for(let i=0;i<GRADES.length;i++){
    const g = GRADES[i];
    const [min,max] = getRange(g);
    if(!Number.isInteger(min) || !Number.isInteger(max)) return {ok:false,message:`${g}: range values must be whole numbers.`};
    if(min < 0 || max > 100) return {ok:false,message:`${g}: range must stay between 0 and 100.`};
    if(min > max) return {ok:false,message:`${g}: minimum cannot be greater than maximum.`};
    if(i === 0 && max !== 100) return {ok:false,message:"The A range must end at 100."};
    if(i === GRADES.length-1 && min !== 0) return {ok:false,message:"The E range must start at 0."};
    if(i > 0){
      const prevMin = getRange(GRADES[i-1])[0];
      if(max !== prevMin - 1){
        return {ok:false,message:"Grade ranges must be continuous with no gaps or overlaps."};
      }
    }
  }
  return {ok:true,message:"✓ Grade ranges are continuous, non-overlapping and cover 0–100."};
}

function buildGradeUI(){
  const root = $("grades");
  root.innerHTML = "";

  GRADES.forEach((g,idx)=>{
    const card = document.createElement("div");
    card.className = "grade-card";
    card.innerHTML = `
      <div class="grade-top">
        <div class="grade-name">${g}</div>
        <div class="range-preview" id="${g}preview">${DEFAULTS[g][0]}–${DEFAULTS[g][1]}</div>
      </div>
      <div class="range-row">
        <label>Minimum<select class="select" id="${g}min"></select></label>
        <label>Maximum<select class="select" id="${g}max"></select></label>
      </div>
    `;
    root.appendChild(card);

    const minSel = $(g+"min");
    const maxSel = $(g+"max");
    for(let i=0;i<=100;i++){
      minSel.add(new Option(i,i,i===DEFAULTS[g][0],i===DEFAULTS[g][0]));
      maxSel.add(new Option(i,i,i===DEFAULTS[g][1],i===DEFAULTS[g][1]));
    }

    const changed = ()=>{
      card.classList.add("changed");
      setTimeout(()=>card.classList.remove("changed"),220);
      updateRangePreview(g);
      updateAll();
    };
    minSel.addEventListener("change",changed);
    maxSel.addEventListener("change",changed);
  });
}

function updateRangePreview(g){
  const [min,max] = getRange(g);
  const el = $(g+"preview");
  if(el) el.textContent = `${min}–${max}`;
}

function resetRanges(){
  GRADES.forEach(g=>{
    if($(g+"min")) $(g+"min").value = DEFAULTS[g][0];
    if($(g+"max")) $(g+"max").value = DEFAULTS[g][1];
    updateRangePreview(g);
  });
  updateAll();
  toast("Grade ranges reset to defaults.");
}

function updateRangeUI(validation){
  const err = $("rangeError");
  const ok = $("rangeOk");
  if(validation.ok){
    err.classList.add("hidden");
    ok.classList.remove("hidden");
    ok.textContent = validation.message || "✓ Grade ranges are valid.";
    $("rangeBadge").textContent = "Ranges valid";
  }else{
    ok.classList.add("hidden");
    err.classList.remove("hidden");
    err.textContent = "⚠ " + validation.message;
    $("rangeBadge").textContent = "Fix ranges";
  }
}

function updateAll(){
  const validation = validateRanges();
  updateRangeUI(validation);
  renderAnalytics();
  renderStudents();
  const canDownload = !!(data.length && selectedCourse && $("instructor").value.trim() && validation.ok);
  $("download").disabled = !canDownload;
}

function renderAnalytics(){
  const rows = getCourseRows();
  $("analyticsSubtitle").textContent = selectedCourse
    ? `${selectedCourse} • ${rows.length} student${rows.length===1?"":"s"}`
    : "Upload a workbook and select a course to begin.";

  if(!rows.length){
    ["studentsKpi","avgKpi","medianKpi","maxKpi","passKpi"].forEach(id=>$(id).textContent="—");
    $("gradeSummary").innerHTML = "";
    drawChart([]);
    return;
  }

  const marks = rows.map(r=>r["Total Marks"]).sort((a,b)=>a-b);
  const sum = marks.reduce((a,b)=>a+b,0);
  const avg = sum / marks.length;
  const median = marks.length % 2
    ? marks[Math.floor(marks.length/2)]
    : (marks[marks.length/2-1] + marks[marks.length/2]) / 2;

  const pass = marks.filter(m=>m >= 40).length / marks.length * 100;
  $("studentsKpi").textContent = marks.length;
  $("avgKpi").textContent = avg.toFixed(2);
  $("medianKpi").textContent = median.toFixed(2);
  $("maxKpi").textContent = marks[marks.length-1];
  $("passKpi").textContent = pass.toFixed(1) + "%";

  drawChart(marks);
  renderGradeSummary(rows);
}

function renderGradeSummary(rows){
  const counts = Object.fromEntries(GRADES.map(g=>[g,0]));
  rows.forEach(r=>{
    const g = gradeForMark(r["Total Marks"]);
    if(g) counts[g]++;
  });

  const root = $("gradeSummary");
  root.innerHTML = "";
  GRADES.forEach(g=>{
    const pill = document.createElement("div");
    pill.className = "grade-pill";
    if(previousSummary[g] !== undefined && previousSummary[g] !== counts[g]){
      pill.style.boxShadow = "0 0 0 3px #f1efff";
      setTimeout(()=>pill.style.boxShadow="",350);
    }
    pill.innerHTML = `<b>${g}</b><span>${counts[g]} student${counts[g]===1?"":"s"}</span>`;
    root.appendChild(pill);
    previousSummary[g] = counts[g];
  });
}

function drawChart(marks){
  const canvas = $("hist");
  const ctx = canvas.getContext("2d");
  const rect = canvas.getBoundingClientRect();
  const dpr = window.devicePixelRatio || 1;
  const w = Math.max(320,Math.floor(rect.width));
  const h = 255;
  canvas.width = w*dpr;
  canvas.height = h*dpr;
  ctx.setTransform(dpr,0,0,dpr,0,0);
  ctx.clearRect(0,0,w,h);

  const pad = {l:38,r:14,t:18,b:38};
  const chartW = w-pad.l-pad.r;
  const chartH = h-pad.t-pad.b;

  ctx.strokeStyle="#e7eaf0";
  ctx.lineWidth=1;
  for(let i=0;i<=4;i++){
    const y=pad.t+chartH-(chartH*i/4);
    ctx.beginPath();ctx.moveTo(pad.l,y);ctx.lineTo(w-pad.r,y);ctx.stroke();
  }

  if(!marks.length){
    ctx.fillStyle="#98a2b3";
    ctx.font="12px system-ui";
    ctx.textAlign="center";
    ctx.fillText("No data to visualize yet",w/2,h/2);
    return;
  }

  const bins=Array(10).fill(0);
  marks.forEach(m=>bins[Math.min(9,Math.floor(m/10))]++);
  const maxBin=Math.max(...bins,1);
  const gap=7;
  const barW=(chartW-gap*9)/10;

  bins.forEach((count,i)=>{
    const bh=(count/maxBin)*(chartH-20);
    const x=pad.l+i*(barW+gap);
    const y=pad.t+chartH-bh;
    const grad=ctx.createLinearGradient(0,y,0,pad.t+chartH);
    grad.addColorStop(0,"#7c6cf0");
    grad.addColorStop(1,"#5b4bdb");
    ctx.fillStyle=grad;
    roundRect(ctx,x,y,barW,bh,7);
    ctx.fill();

    ctx.fillStyle="#667085";
    ctx.font="9px system-ui";
    ctx.textAlign="center";
    ctx.fillText(`${i*10}`,x+barW/2,h-19);

    if(count){
      ctx.fillStyle="#344054";
      ctx.font="10px system-ui";
      ctx.fillText(String(count),x+barW/2,y-6);
    }
  });

  ctx.fillStyle="#98a2b3";
  ctx.font="9px system-ui";
  ctx.textAlign="left";
  ctx.fillText("Students",pad.l,10);
}

function roundRect(ctx,x,y,w,h,r){
  if(h<=0) return;
  const rr=Math.min(r,w/2,h);
  ctx.beginPath();
  ctx.moveTo(x+rr,y);
  ctx.arcTo(x+w,y,x+w,y+h,rr);
  ctx.arcTo(x+w,y+h,x,y+h,rr);
  ctx.arcTo(x,y+h,x,y,rr);
  ctx.arcTo(x,y,x+w,y,rr);
  ctx.closePath();
}

function renderStudents(){
  const tbody = $("studentBody");
  const term = searchTerm.trim().toLowerCase();
  const baseRows = selectedCourse ? getCourseRows() : data;

  if(!data.length){
    tbody.innerHTML='<tr><td colspan="4"><div class="empty">Upload a valid Excel workbook to preview students.</div></td></tr>';
    return;
  }

  if(!selectedCourse && !term){
    tbody.innerHTML='<tr><td colspan="4"><div class="empty">Choose a course above, or type in the search box to search all uploaded students.</div></td></tr>';
    return;
  }

  const filtered = baseRows.filter(r=>{
    if(!term) return true;
    const grade=gradeForMark(r["Total Marks"]);
    return `${r["BITS ID"]} ${r.Course} ${r["Total Marks"]} ${grade}`.toLowerCase().includes(term);
  });

  if(!filtered.length){
    tbody.innerHTML='<tr><td colspan="4"><div class="empty">No students match your search.</div></td></tr>';
    return;
  }

  tbody.innerHTML=filtered.slice(0,250).map(r=>{
    const grade=gradeForMark(r["Total Marks"]) || "—";
    return `<tr>
      <td><strong>${escapeHTML(r["BITS ID"])}</strong></td>
      <td>${escapeHTML(r.Course)}</td>
      <td>${r["Total Marks"]}</td>
      <td><span class="grade-tag">${escapeHTML(grade)}</span></td>
    </tr>`;
  }).join("");

  if(filtered.length>250){
    tbody.innerHTML += `<tr><td colspan="4" class="muted">Showing first 250 matching students.</td></tr>`;
  }
}

function renderEmptyState(){
  selectedCourse="";
  $("analyticsSubtitle").textContent="Upload a workbook and select a course to begin.";
  ["studentsKpi","avgKpi","medianKpi","maxKpi","passKpi"].forEach(id=>$(id).textContent="—");
  $("gradeSummary").innerHTML="";
  $("studentBody").innerHTML='<tr><td colspan="4"><div class="empty">No course selected.</div></td></tr>';
  drawChart([]);
  updateAll();
}

function exportGrades(){
  const instructor = $("instructor").value.trim();
  if(!instructor){showStatus("Please enter the instructor name before finalizing.","error");$("instructor").focus();return;}
  if(!data.length){showStatus("Please upload a valid Excel workbook first.","error");return;}
  if(!selectedCourse){showStatus("Please select a course before finalizing.","error");return;}

  const validation=validateRanges();
  if(!validation.ok){showStatus(validation.message,"error");return;}

  const rows=getCourseRows();
  if(!rows.length){showStatus("The selected course has no valid student records.","error");return;}

  finalizeCount++;
  const elapsed=Date.now()-gradingStartTime;
  const minT=Math.floor(elapsed/60000);
  const secT=Math.floor((elapsed%60000)/1000);

  const output=[
    ["Instructor",instructor],
    ["Course",selectedCourse],
    ["Generated At",new Date().toLocaleString()],
    ["Student Count",rows.length],
    [],
    ["BITS ID","Course","Total Marks","Grade"]
  ];

  rows.forEach(r=>output.push([r["BITS ID"],r.Course,r["Total Marks"],gradeForMark(r["Total Marks"])]));

  const csv=output.map(row=>row.map(csvEscape).join(",")).join("\r\n");
  const safeCourse=selectedCourse.replace(/[^a-z0-9]+/gi,"-").replace(/^-|-$/g,"").toLowerCase() || "course";
  downloadBlob(csv,`grades-${safeCourse}.csv`,"text/csv;charset=utf-8");

  showStatus(`✓ Finalized ${rows.length} grades for ${selectedCourse}. CSV downloaded successfully.`,"success");
  toast(`Export complete • ${minT}m ${String(secT).padStart(2,"0")}s`);
}

function resetApp(){
  if(!confirm("Reset the current grading session and clear the loaded dataset?")) return;
  data=[];
  selectedCourse="";
  finalizeCount=0;
  previousSummary={};
  searchTerm="";
  $("file").value="";
  $("fileName").textContent="No file selected";
  $("instructor").value="";
  $("studentSearch").value="";
  $("studentSearch").disabled=true;
  $("datasetBadge").textContent="No dataset loaded";
  resetCourseOptions();
  hideStatus();
  renderEmptyState();
  resetRanges();
  gradingStartTime=Date.now();
  toast("Application reset.");
}

function updateTimer(){
  const elapsed=Date.now()-gradingStartTime;
  const min=Math.floor(elapsed/60000);
  const sec=Math.floor((elapsed%60000)/1000);
  $("timerBadge").textContent=`⏱ ${String(min).padStart(2,"0")}:${String(sec).padStart(2,"0")}`;
}
function startTimer(){
  updateTimer();
  clearInterval(timerInterval);
  timerInterval=setInterval(updateTimer,1000);
}

const BUG_LOG = [
  ["1","Course list could retain duplicate/stale options after another upload.","Uploaded two workbooks sequentially.","Existing options were not cleared.","Selector rebuilt from latest unique course list.","Repeated uploads; verified latest unique list."],
  ["2","Statistics could become NaN/undefined for an empty course.","Selected a course with no usable records.","Calculations assumed at least one mark.","Added safe empty-state handling.","Tested empty/invalid-course states."],
  ["3","Weak workbook/data validation.","Uploaded missing columns, blanks, invalid marks and duplicate IDs.","Raw sheet JSON was consumed without schema checks.","Added schema, range, blank and duplicate validation.","Tested each invalid input class."],
  ["4","File picker advertised only .xls.","Tried selecting .xlsx.","Original accept attribute was .xls only.","Supports .xlsx and .xls.","Tested both extensions."],
  ["5","Range validation did not enforce global 0–100 boundaries.","Changed edge ranges to invalid coverage.","Only local continuity was checked.","First band must end at 100; final band must start at 0.","Tested gaps, overlaps, reversed and edge ranges."],
  ["6","CSV fields were not escaped.","Used punctuation in instructor/course/IDs and exported.","CSV was built by raw string concatenation.","Added CSV escaping and safe filenames.","Opened exports containing commas/quotes."],
  ["7","Min/Max KPI labels were reversed.","Compared UI labels with actual values.","Original markup bound max under Min and min under Max.","Corrected KPI mapping and added students/pass rate.","Verified against source marks."]
];

function downloadBugLog(){
  const header=["#","Bug / Issue Identified","How You Reproduced It","Root Cause","Fix Implemented","How You Tested the Fix"];
  const csv=[header,...BUG_LOG].map(row=>row.map(csvEscape).join(",")).join("\r\n");
  downloadBlob(csv,"BITS-CodeForge-Bug-Fix-Log.csv","text/csv;charset=utf-8");
  toast("Bug Fix Log downloaded.");
}

const enhancementSummary =
`BITS Digital CodeForge V1.0 – Enhancement Summary

1. Safer validation:
- Validates required Excel columns.
- Rejects blank, non-numeric, fractional and out-of-range marks.
- Detects duplicate BITS ID + course records.
- Validates grade ranges across the complete 0–100 scale.

2. Actionable analytics:
- Student count, average, median, highest mark and pass rate.
- Grade distribution and marks histogram.
- Course-specific analytics update instantly.

3. Faster review:
- Searchable student preview table.
- Grade is calculated from the active ranges before export.

4. Responsive modern UI:
- Mobile-friendly layout, clearer hierarchy, accessible controls and empty states.

5. Safer export:
- Correct CSV escaping.
- Course-specific filename.
- Export is disabled until all required validation passes.

6. Local-first processing:
- Excel processing happens in the browser without a backend.`;

function copySummary(){
  navigator.clipboard?.writeText(enhancementSummary).then(
    ()=>toast("Enhancement summary copied."),
    ()=>toast("Copy failed; select the text manually.")
  );
}

function loadDemoData(){
  const demoRows = [
    {"BITS ID":"2024A001","Course":"Course A","Total Marks":92},
    {"BITS ID":"2024A002","Course":"Course A","Total Marks":84},
    {"BITS ID":"2024A003","Course":"Course A","Total Marks":77},
    {"BITS ID":"2024A004","Course":"Course A","Total Marks":68},
    {"BITS ID":"2024A005","Course":"Course A","Total Marks":55},
    {"BITS ID":"2024A006","Course":"Course B","Total Marks":88},
    {"BITS ID":"2024A007","Course":"Course B","Total Marks":73},
    {"BITS ID":"2024A008","Course":"Course B","Total Marks":61},
    {"BITS ID":"2024A009","Course":"Course B","Total Marks":46}
  ];
  data = demoRows;
  lastFileName = "CodeForge-demo-data.xlsx";
  resetCourseOptions("Select a course");
  const courseSelect = $("course");
  getCourses().forEach(c=>courseSelect.add(new Option(c,c)));
  courseSelect.disabled=false;
  $("studentSearch").disabled=false;
  $("studentSearch").placeholder="Search all students…";
  $("fileName").textContent="Demo dataset loaded";
  $("datasetBadge").textContent=`${data.length} records • ${getCourses().length} courses`;
  if($("courseHelp")) $("courseHelp").textContent="Demo data is loaded locally. You can replace it by uploading the challenge Excel file.";
  showStatus("✓ Demo dataset loaded. Try the course dropdown and student search now.","success");
  renderEmptyState();
  updateAll();
}

/* Events */
$("file").addEventListener("change",e=>handleFile(e.target.files[0]));
$("demoBtn")?.addEventListener("click",loadDemoData);

$("dropzone").addEventListener("dragover",e=>{
  e.preventDefault();$("dropzone").classList.add("dragover");
});
$("dropzone").addEventListener("dragleave",()=>$("dropzone").classList.remove("dragover"));
$("dropzone").addEventListener("drop",e=>{
  e.preventDefault();$("dropzone").classList.remove("dragover");
  const file=e.dataTransfer.files?.[0];
  if(file) handleFile(file);
});

$("instructor").addEventListener("input",()=>{
  if(selectedCourse){
    $("welcome")?.remove?.();
  }
  updateAll();
});

$("course").addEventListener("change",()=>{
  selectedCourse=$("course").value;
  previousSummary={};
  $("studentSearch").value="";
  searchTerm="";
  $("studentSearch").placeholder = selectedCourse ? `Search ${selectedCourse} students…` : "Search all students…";
  if(selectedCourse){
    showStatus(`✓ ${selectedCourse} selected. Review the analytics and student preview below.`,"success");
  }else if(data.length){
    showStatus("Choose a course to focus analytics, or use the search box to search all uploaded students.","warning");
  }else{
    hideStatus();
  }
  updateAll();
});

$("studentSearch").addEventListener("input",e=>{
  searchTerm=e.target.value;
  renderStudents();
});

$("resetRanges").addEventListener("click",()=>{
  if(confirm("Reset all grade ranges to the default A–E ranges?")) resetRanges();
});

$("download").addEventListener("click",exportGrades);
$("resetAppBtn").addEventListener("click",resetApp);

$("logBtn").addEventListener("click",()=> $("modalBackdrop").classList.add("open"));
$("submissionBtn").addEventListener("click",e=>{e.preventDefault();$("modalBackdrop").classList.add("open")});
$("closeModal").addEventListener("click",()=> $("modalBackdrop").classList.remove("open"));
$("modalBackdrop").addEventListener("click",e=>{if(e.target===$("modalBackdrop")) $("modalBackdrop").classList.remove("open")});
$("downloadLog")?.addEventListener("click",downloadBugLog);
$("copySummary")?.addEventListener("click",copySummary);
document.addEventListener("keydown",e=>{if(e.key==="Escape") $("modalBackdrop").classList.remove("open")});
window.addEventListener("resize",()=>drawChart(getCourseRows().map(r=>r["Total Marks"])));

buildGradeUI();
startTimer();
renderEmptyState();
