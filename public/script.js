// Helper to create demo entries for other regulations (keeps same subjects/filepaths with reg replaced)
function createRegCopy(reg) {
  // deep copy R15 but replace first-level file paths strings with reg
  const copy = JSON.parse(JSON.stringify(portalData.R15));
  // update file paths: replace "R15" with reg in all string values
  function replaceStrings(obj) {
    if (!obj || typeof obj !== "object") return;
    Object.keys(obj).forEach(k=>{
      const v = obj[k];
      if (typeof v === "string") {
        obj[k] = v.replace(/R15/g, reg);
      } else if (typeof v === "object") {
        replaceStrings(v);
      }
    });
  }
  replaceStrings(copy);
  return copy;
}

// -----------------------------------------------------
// DOM references
// -----------------------------------------------------
const regulationSelect = document.getElementById("regulation");
const branchSelect = document.getElementById("branch");
const yearSelect = document.getElementById("year");
const semSelect = document.getElementById("sem");
const materialTypeSelect = document.getElementById("materialType");
const subjectSelect = document.getElementById("subject");
const materialsListDiv = document.getElementById("materialsList");
const searchBtn = document.getElementById("searchBtn");

// -----------------------------------------------------
// Populate branch dropdown when regulation changes
// -----------------------------------------------------
function populateBranches() {
  const reg = regulationSelect.value;
  branchSelect.innerHTML = '<option value="">Select Branch</option>';
  if (reg && portalData[reg]) {
    Object.keys(portalData[reg]).forEach(branch => {
      const opt = document.createElement("option");
      opt.value = branch;
      opt.textContent = branch;
      branchSelect.appendChild(opt);
    });
  }
  // reset subject and materials
  subjectSelect.innerHTML = "<option value=''>Select Subject</option>";
  materialsListDiv.textContent = "Please select options above and click Search.";
}
regulationSelect.addEventListener("change", populateBranches);

// -----------------------------------------------------
// Update subject dropdown based on reg/branch/year/sem/materialType
// -----------------------------------------------------
function updateSubjectsAuto() {
  const reg = regulationSelect.value;
  const branch = branchSelect.value;
  const year = yearSelect.value;
  const sem = semSelect.value;
  const type = materialTypeSelect.value; // may be empty

  subjectSelect.innerHTML = "<option value=''>Select Subject</option>";

  if (!reg || !branch || !year || !sem) {
    // nothing to show until these are selected
    return;
  }

  const semData = portalData?.[reg]?.[branch]?.[year]?.[sem];
  if (!semData) return;

  // semData is object where keys are subject names
  Object.entries(semData).forEach(([subName, files]) => {
    // ensure we only add if files object exists and if selected material type exists (or if no type selected)
    if (typeof files === "object") {
      if (!type || files[type]) {
        const opt = document.createElement("option");
        opt.value = subName;
        opt.textContent = subName;
        subjectSelect.appendChild(opt);
      }
    }
  });
}

// Attach listeners to update subject list when any relevant select changes
[branchSelect, yearSelect, semSelect, materialTypeSelect].forEach(el =>
  el.addEventListener("change", updateSubjectsAuto)
);

// If regulation changes, branches will change, then user picks branch -> trigger updateSubjectsAuto on branch change
branchSelect.addEventListener("change", () => {
  // ensure subject list updates based on new branch/year/sem/type
  updateSubjectsAuto();
});

// -----------------------------------------------------
// Utility: create a download anchor element
// -----------------------------------------------------
function makeDownloadLink(url, text) {
  const a = document.createElement("a");
  a.href = url;
  a.textContent = "⬇ " + text;
  a.style.display = "block";
  a.style.margin = "6px 0";
  
  a.rel = "noopener";
  // If same-origin relative path, set download attribute for direct download
  if (!/^https?:\/\//i.test(url)) {
    // Only set download for same-origin / relative resources
    a.setAttribute("download", "");
  }
  return a;
}

// -----------------------------------------------------
// Show downloads for a subject
// -----------------------------------------------------
function showDownloadsForSubject(reg, branch, year, sem, subjectName) {
  materialsListDiv.innerHTML = "";
  const subj = portalData?.[reg]?.[branch]?.[year]?.[sem]?.[subjectName];
  if (!subj) {
    materialsListDiv.textContent = "No data found for this subject.";
    return;
  }

  // If syllabus (single file)
  if (subj.syllabus) {
    materialsListDiv.appendChild(makeDownloadLink(subj.syllabus, `${subjectName} - syllabus`));
  }

  // For unit-wise categories
  ["materialsPdf", "handwriting", "previousPapers", "importantQuestions"].forEach(cat => {
    if (subj[cat] && typeof subj[cat] === "object") {
      const h = document.createElement("h3");
      h.textContent = `${subjectName} — ${cat}`;
      materialsListDiv.appendChild(h);

      Object.entries(subj[cat]).forEach(([unit, url]) => {
        materialsListDiv.appendChild(makeDownloadLink(url, `${unit}`));
      });
    }
  });
}

// -----------------------------------------------------
// When search button clicked
// - If subject selected → show downloads for that subject
// - If subject NOT selected → list matched subjects as clickable items
// -----------------------------------------------------
searchBtn.addEventListener("click", () => {
  const reg = regulationSelect.value;
  const branch = branchSelect.value;
  const year = yearSelect.value;
  const sem = semSelect.value;
  const type = materialTypeSelect.value;
  const subject = subjectSelect.value;

  materialsListDiv.innerHTML = "";

  if (!reg || !branch || !year || !sem || !type) {
    materialsListDiv.textContent = "Please select regulation, branch, year, sem and material type.";
    return;
  }

  const semData = portalData?.[reg]?.[branch]?.[year]?.[sem];
  if (!semData) {
    materialsListDiv.textContent = "No subjects found for this selection.";
    return;
  }

  // ✔ If specific subject selected → Show only that subject’s materials
  if (subject) {
    const subjObj = semData[subject];
    if (!subjObj) {
      materialsListDiv.textContent = "Selected subject not found.";
      return;
    }

    // Syllabus (single file)
    if (type === "syllabus") {
      if (subjObj.syllabus) {
        materialsListDiv.appendChild(makeDownloadLink(subjObj.syllabus, `${subject} - syllabus`));
      } else {
        materialsListDiv.textContent = "Syllabus not available for this subject.";
      }
      return;
    }

    // Other categories (unit-wise)
    const cat = subjObj[type];
    if (cat && typeof cat === "object") {
      const header = document.createElement("h3");
      header.textContent =`${subject} — ${type}`;
      materialsListDiv.appendChild(header);

      Object.entries(cat).forEach(([unit, url]) => {
        materialsListDiv.appendChild(makeDownloadLink(url, unit));
      });
      return;
    }

    if (typeof cat === "string") {
      materialsListDiv.appendChild(makeDownloadLink(cat, `${subject} - ${type}`));
      return;
    }

    materialsListDiv.textContent = `No ${type} available for this subject.`;
    return;
  }

  // ✔ No subject selected → Show all subjects that match material type
  const matched = Object.entries(semData).filter(([_, files]) => {
    if (!files) return false;
    if (type === "syllabus") return !!files.syllabus;
    return !!files[type];
  });

  if (matched.length === 0) {
    materialsListDiv.textContent = "No subjects matched the selected material type.";
    return;
  }

  const info = document.createElement("p");
  info.textContent = "Matched subjects (click a subject to view downloads):";
  materialsListDiv.appendChild(info);

  matched.forEach(([subName]) => {
    const btn = document.createElement("button");
    btn.textContent = subName;
    btn.style.display = "block";
    btn.style.margin = "8px 0";
    btn.addEventListener("click", () => {
      showDownloadsForSubject(reg, branch, year, sem, subName);
    });
    materialsListDiv.appendChild(btn);
  });
});

// -----------------------------------------------------
// Keep subjects selected after download / when user returns
// No page reload behavior is required; downloads open in new tab.
// But we also want the subject dropdown to be pre-populated when page loads if selects already have values.
// -----------------------------------------------------
function initAutoPopulateOnLoad() {
  populateBranches();       // populate branches if reg already preselected
  updateSubjectsAuto();     // populate subject options if branch/year/sem already selected
}

// run on load
window.addEventListener("load", initAutoPopulateOnLoad);


async function login(){
    const email=document.getElementById('loginEmail').value;
    const password=document.getElementById('loginPassword').value;
    const error=document.getElementById('loginError');
    error.style.display='none';

    if(!email || !password){
        error.textContent='Fill all fields';
        error.style.display='block';
        return;
    }

    const res = await fetch('/login',{
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body:JSON.stringify({email,password})
    });

    const data = await res.json();

    if(res.ok){
        window.location.href='/main-portal.html';
    } else {
        error.textContent=data.message;
        error.style.display='block';
    }
}

async function register(){
    const data={
        name: document.getElementById('name').value,
        regulation: document.getElementById('regulation').value,
        email: document.getElementById('email').value,
        rollNumber: document.getElementById('rollNumber').value,
        yearofstudy: document.getElementById('yearofstudy').value,
        branch: document.getElementById('branch').value,
        phone: document.getElementById('phone').value,
        password: document.getElementById('password').value,
        confirmPassword: document.getElementById('confirmPassword').value
    };

    const error=document.getElementById('registerError');
    error.style.display='none';

    if(Object.values(data).some(v=>!v)){
        error.textContent='Fill all fields';
        error.style.display='block';
        return;
    }

    if(data.password!==data.confirmPassword){
        error.textContent='Passwords do not match';
        error.style.display='block';
        return;
    }

    const res = await fetch('/register',{
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body:JSON.stringify(data)
    });

    const result = await res.json();

    if(res.ok){
        showLogin();
    } else {
        error.textContent=result.message;
        error.style.display='block';
    }
}
