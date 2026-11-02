/* ==========================================================================
   Result Analysis Dashboard - Core Application Engine
   Comprehensive VTU / BMSIT Result Analysis & Faculty Management System
   ========================================================================== */

// --- Global Application State ---
let rawRows = [];               // Filtered rows for current active scope
let allRows = [];               // All loaded rows from file
let savedSections = [];         // [{ label, semester, start, end, customUsns, combined, isAll }]
let activeSectionIndex = -1;    // -1 = All students, >=0 = index in savedSections
let subjectSummary = [];        // [{ code, title, total, appeared, passed, failed, pct, w, f, ab, dx, x, i }]
let sortKey = 'pct', sortDir = 1;
let facultySortKey = 'pct', facultySortDir = -1;
let chartInstance = null;
let facultyChartInstance = null;

// Overrides & Allocations
// subjectOverrides: { [code]: { staff, classStrength, withheld, scope } }
let subjectOverrides = {};
// facultyMap: { [scopeKey]: { [courseCode]: facultyName } } where scopeKey is '-1' (global) or section index
let facultyMap = { '-1': {} };

let overallWithheld = 0;
let overallWithheldMap = {};    // batch -> manually entered withheld count
let overallWithheldScope = null;
let sigNames = { classTeacher: '', coordinator: '', divisionHead: '', hod: '' };

let categoryMap = null;         // USN -> Category
let categoryFileName = '';
let revalLog = [];              // [{ usn, name, code, title, oldGrade, newGrade, oldTotal, newTotal, statusChange }]

// Branch Name Mapping
const BRANCH_NAME_MAP = {
  'CS': 'Computer Science and Engineering',
  'CSE': 'Computer Science and Engineering',
  'IS': 'Information Science and Engineering',
  'ISE': 'Information Science and Engineering',
  'EC': 'Electronics and Communication Engineering',
  'ECE': 'Electronics and Communication Engineering',
  'EE': 'Electrical and Electronics Engineering',
  'EEE': 'Electrical and Electronics Engineering',
  'ME': 'Mechanical Engineering',
  'MECH': 'Mechanical Engineering',
  'CV': 'Civil Engineering',
  'CIV': 'Civil Engineering',
  'CT': 'Computer Science and Technology',
  'CD': 'Computer Science and Design',
  'AI': 'Artificial Intelligence and Machine Learning',
  'AIML': 'Artificial Intelligence and Machine Learning',
  'CY': 'Computer Science and Engineering (Cyber Security)',
  'CSBS': 'Computer Science and Business Systems',
  'ET': 'Electronics and Telecommunication Engineering',
  'BT': 'Biotechnology',
  'CH': 'Chemical Engineering',
  'IM': 'Industrial Engineering and Management',
  'AE': 'Aeronautical Engineering'
};

// Standard Grade Codes
const STANDARD_GRADE_CODES = ['O', 'A+', 'A', 'B+', 'B', 'C', 'P', 'F', 'AB', 'DX', 'X', 'I', 'W'];

function normalize(s) {
  return String(s || '').trim().toLowerCase();
}

// --- Faculty Helper Functions ---
function getAssignedFaculty(code, scopeIdx = activeSectionIndex) {
  // Check section-specific first
  const scopeKey = String(scopeIdx);
  if (facultyMap[scopeKey] && facultyMap[scopeKey][code]) {
    return facultyMap[scopeKey][code];
  }
  // Check global ('-1')
  if (facultyMap['-1'] && facultyMap['-1'][code]) {
    return facultyMap['-1'][code];
  }
  // Check subjectOverrides
  if (subjectOverrides[code] && subjectOverrides[code].staff) {
    return subjectOverrides[code].staff;
  }
  return '';
}

function setAssignedFaculty(code, facultyName, scopeIdx = activeSectionIndex) {
  const scopeKey = String(scopeIdx);
  if (!facultyMap[scopeKey]) facultyMap[scopeKey] = {};
  facultyMap[scopeKey][code] = facultyName.trim();

  // Also keep subjectOverrides in sync
  if (!subjectOverrides[code]) {
    subjectOverrides[code] = { staff: facultyName.trim(), classStrength: 0, withheld: 0, scope: scopeIdx };
  } else {
    subjectOverrides[code].staff = facultyName.trim();
  }
}

// --- Branch & Header Detection ---
function extractBranchCode(usn) {
  const m = String(usn || '').trim().toUpperCase().match(/^\d[A-Z0-9]{2,4}\d{2}([A-Z]{2,4})\d+$/);
  return m ? m[1] : null;
}

function detectDominantBranch(rows) {
  const counts = {};
  const usns = new Set(rows.map(r => r.usn));
  usns.forEach(u => {
    const code = extractBranchCode(u);
    if (code) counts[code] = (counts[code] || 0) + 1;
  });
  const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  if (entries.length === 0) return null;
  const [code, count] = entries[0];
  return { code, count, total: usns.size, name: BRANCH_NAME_MAP[code] || null };
}

function applyDetectedBranch() {
  const detected = detectDominantBranch(allRows);
  if (!detected) return;
  const deptEl = document.getElementById('metaDept');
  if (deptEl) {
    const label = detected.name || `${detected.code} Engineering`;
    deptEl.value = `Department of ${label}`;
    renderReportHeader();
  }
}

// --- Result Grid Parser (Excel / CSV) ---
function parseResultRows(grid) {
  let headerRowIdx = -1, colMap = {};
  for (let r = 0; r < Math.min(grid.length, 30); r++) {
    const row = grid[r].map(normalize);
    const hasUsn = row.some(c => c.includes('usn'));
    const hasGrade = row.some(c => c === 'grade');
    if (hasUsn && hasGrade) { headerRowIdx = r; break; }
  }
  if (headerRowIdx === -1) throw new Error('Could not find a header row with USN and Grade columns');

  const header = grid[headerRowIdx].map(normalize);
  const findCol = (...names) => {
    for (const n of names) {
      const idx = header.findIndex(h => h === n);
      if (idx !== -1) return idx;
    }
    for (const n of names) {
      const idx = header.findIndex(h => h.includes(n));
      if (idx !== -1) return idx;
    }
    return -1;
  };

  colMap = {
    usn: findCol('usn'),
    name: findCol('name'),
    code: findCol('course code', 'coursecode', 'sub code', 'subject code'),
    title: findCol('course title', 'coursetitle', 'sub title', 'subject title', 'subject name'),
    grade: findCol('grade'),
    sem: findCol('semester', 'sem'),
    total: findCol('total', 'total marks'),
    credits: findCol('credits registered', 'creditsregistered', 'credits'),
    sgpa: findCol('sgpa'),
    cgpa: findCol('cgpa')
  };

  const rows = [];
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r];
    if (!row || row.length === 0) continue;
    const usn = colMap.usn !== -1 ? String(row[colMap.usn] || '').trim().toUpperCase() : '';
    const code = colMap.code !== -1 ? String(row[colMap.code] || '').trim() : '';
    const grade = colMap.grade !== -1 ? String(row[colMap.grade] || '').trim().toUpperCase() : '';
    if (!usn || !code) continue;
    rows.push({
      usn,
      name: colMap.name !== -1 ? String(row[colMap.name] || '').trim() : '',
      code,
      title: colMap.title !== -1 ? String(row[colMap.title] || '').trim() : code,
      grade,
      sem: colMap.sem !== -1 ? row[colMap.sem] : '',
      total: colMap.total !== -1 ? Number(row[colMap.total]) || 0 : 0,
      credits: colMap.credits !== -1 ? Number(row[colMap.credits]) || 0 : 0,
      sgpa: colMap.sgpa !== -1 ? Number(row[colMap.sgpa]) || 0 : 0,
      cgpa: colMap.cgpa !== -1 ? Number(row[colMap.cgpa]) || 0 : 0
    });
  }
  if (rows.length === 0) throw new Error('Header row found, but no data rows parsed underneath it');
  return rows;
}

function parseGrid(grid) {
  allRows = parseResultRows(grid);
  subjectOverrides = {};
  facultyMap = { '-1': {} };
  categoryMap = null;
  savedSections = [];
  activeSectionIndex = -1;
  revalLog = [];

  const catStatus = document.getElementById('categoryStatus');
  if (catStatus) catStatus.textContent = '';
  const catSec = document.getElementById('categorySection');
  if (catSec) catSec.style.display = 'none';
  const secStatus = document.getElementById('sectionStatus');
  if (secStatus) secStatus.textContent = '';
  const revStatus = document.getElementById('revalStatus');
  if (revStatus) revStatus.textContent = '';
  const revWrap = document.getElementById('revalLogWrap');
  if (revWrap) revWrap.style.display = 'none';

  renderSectionSelect();
  rawRows = allRows.slice();
  applyDetectedBranch();
}

function handleFile(file) {
  const meta = document.getElementById('fileMeta');
  if (meta) {
    meta.style.display = 'inline-flex';
    meta.textContent = `Loading ${file.name}…`;
  }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      parseGrid(grid);

      buildGradeAssignTable();
      buildSummary();

      document.getElementById('emptyHint').style.display = 'none';
      document.getElementById('results').style.display = 'block';
      const landing = document.getElementById('landingSavedReports');
      if (landing) landing.style.display = 'none';

      render();
      renderFacultyTab();
      renderReport();

      if (meta) {
        const studentCount = new Set(allRows.map(r => r.usn)).size;
        const subCount = new Set(allRows.map(r => r.code)).size;
        meta.textContent = `✓ Loaded ${file.name} — ${studentCount} students, ${subCount} subjects, ${allRows.length} total records.`;
      }
    } catch (err) {
      alert('Error parsing file: ' + err.message);
      if (meta) meta.textContent = 'Failed to parse file: ' + err.message;
    }
  };
  reader.readAsArrayBuffer(file);
}

// --- Grade Classification Settings ---
function getFailSet() {
  const input = document.getElementById('failGradesInput');
  return (input ? input.value : 'F, DX, X, I').split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
}

function getExcludeSet() {
  const input = document.getElementById('excludeGradesInput');
  return (input ? input.value : 'AB, W').split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
}

function getWithheldSet() {
  const input = document.getElementById('withheldGradesInput');
  return (input ? input.value : 'W').split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
}

function buildGradeAssignTable() {
  const detected = new Set(rawRows.map(r => r.grade).filter(Boolean));
  STANDARD_GRADE_CODES.forEach(g => detected.add(g));
  const sorted = Array.from(detected).sort();

  const failSet = new Set(getFailSet());
  const excludeSet = new Set(getExcludeSet());

  let html = '';
  sorted.forEach(g => {
    const isFail = failSet.has(g);
    const isExclude = excludeSet.has(g);
    html += `<tr>
      <td class="grade-code">${g}</td>
      <td><input type="radio" name="grade_${g}" value="fail" ${isFail ? 'checked' : ''}></td>
      <td><input type="radio" name="grade_${g}" value="exclude" ${isExclude ? 'checked' : ''}></td>
      <td><input type="radio" name="grade_${g}" value="pass" ${(!isFail && !isExclude) ? 'checked' : ''}></td>
    </tr>`;
  });
  const tbody = document.getElementById('gradeAssignBody');
  if (tbody) {
    tbody.innerHTML = html;
    tbody.querySelectorAll('input[type=radio]').forEach(r => {
      r.addEventListener('change', () => {
        syncGradeAssignmentInputsOnly();
        buildSummary();
        render();
        renderFacultyTab();
        renderReport();
      });
    });
  }
}

function syncGradeAssignmentInputsOnly() {
  const fails = [], excludes = [];
  document.querySelectorAll('#gradeAssignBody tr').forEach(tr => {
    const code = tr.querySelector('.grade-code').textContent.trim();
    const checked = tr.querySelector('input[type=radio]:checked');
    if (!checked) return;
    if (checked.value === 'fail') fails.push(code);
    else if (checked.value === 'exclude') excludes.push(code);
  });
  document.getElementById('failGradesInput').value = fails.join(', ');
  document.getElementById('excludeGradesInput').value = excludes.join(', ');
}

// --- Summary Calculation ---
function buildSummary() {
  const failSet = new Set(getFailSet());
  const excludeSet = new Set(getExcludeSet());
  const bySubject = {};

  for (const row of rawRows) {
    const key = row.code + '||' + row.title;
    if (!bySubject[key]) {
      bySubject[key] = {
        code: row.code,
        title: row.title,
        total: 0,
        appeared: 0,
        failed: 0,
        w: 0,
        f: 0,
        ab: 0,
        dx: 0,
        x: 0,
        i: 0
      };
    }
    bySubject[key].total++;
    if (row.grade === 'F') bySubject[key].f++;
    if (row.grade === 'AB') bySubject[key].ab++;
    if (row.grade === 'DX') bySubject[key].dx++;
    if (row.grade === 'X') bySubject[key].x++;
    if (row.grade === 'I') bySubject[key].i++;
    if (row.grade === 'W') { bySubject[key].w++; continue; }
    if (excludeSet.has(row.grade)) continue;

    bySubject[key].appeared++;
    if (failSet.has(row.grade)) bySubject[key].failed++;
  }

  subjectSummary = Object.values(bySubject).map(s => {
    const passed = s.appeared - s.failed;
    const pct = s.appeared ? +(passed / s.appeared * 100).toFixed(2) : 0;
    return {
      ...s,
      passed,
      pct,
      faculty: getAssignedFaculty(s.code)
    };
  });
}

function pctClass(pct) {
  if (pct >= 85) return 'high';
  if (pct >= 60) return 'mid';
  return 'low';
}

function barColor(pct) {
  if (pct >= 85) return '#1f7a4d';
  if (pct >= 60) return '#b8862c';
  return '#b3402a';
}

// --- Dashboard Renderers ---
function render() {
  renderCards();
  renderTable();
  renderChart();
  renderLegend();
  const titleEl = document.getElementById('titleLine');
  if (titleEl) {
    titleEl.textContent = `Result Summary — ${subjectSummary.length} subjects, ${rawRows.length} records`;
  }
}

function renderCards() {
  const totalRecords = rawRows.length;
  const failSet = new Set(getFailSet());
  const totalFail = rawRows.filter(r => failSet.has(r.grade)).length;
  const overallPct = totalRecords ? ((totalRecords - totalFail) / totalRecords * 100).toFixed(1) : 0;
  const students = new Set(rawRows.map(r => r.usn)).size;
  const sorted = subjectSummary.slice().sort((a, b) => a.pct - b.pct);
  const worst = sorted[0];
  const best = sorted[sorted.length - 1];

  const cards = [
    { lbl: 'Students', num: students, cls: 'accent' },
    { lbl: 'Subjects', num: subjectSummary.length, cls: '' },
    { lbl: 'Records analysed', num: totalRecords, cls: '' },
    { lbl: 'Overall pass %', num: overallPct + '%', cls: 'pass' },
    { lbl: 'Lowest Pass Subject', num: worst ? `${worst.code} (${worst.pct}%)` : '—', cls: 'fail' },
    { lbl: 'Highest Pass Subject', num: best ? `${best.code} (${best.pct}%)` : '—', cls: 'pass' }
  ];

  const container = document.getElementById('summaryCards');
  if (container) {
    container.innerHTML = cards.map(c => `
      <div class="card ${c.cls}">
        <div class="num">${c.num}</div>
        <div class="lbl">${c.lbl}</div>
      </div>
    `).join('');
  }
}

function renderTable() {
  const query = (document.getElementById('subjectSearchInput') ? document.getElementById('subjectSearchInput').value : '').toLowerCase().trim();

  let filtered = subjectSummary.slice();
  if (query) {
    filtered = filtered.filter(s =>
      s.code.toLowerCase().includes(query) ||
      s.title.toLowerCase().includes(query) ||
      (getAssignedFaculty(s.code) || '').toLowerCase().includes(query)
    );
  }

  const sorted = filtered.sort((a, b) => {
    let av = a[sortKey], bv = b[sortKey];
    if (sortKey === 'faculty') {
      av = getAssignedFaculty(a.code) || '';
      bv = getAssignedFaculty(b.code) || '';
    }
    if (typeof av === 'string') { av = av.toLowerCase(); bv = bv.toLowerCase(); }
    if (av < bv) return -1 * sortDir;
    if (av > bv) return 1 * sortDir;
    return 0;
  });

  const tbody = document.getElementById('tableBody');
  if (!tbody) return;

  if (sorted.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; padding:20px; color:var(--muted);">No matching subjects found.</td></tr>`;
    return;
  }

  tbody.innerHTML = sorted.map(s => {
    const faculty = getAssignedFaculty(s.code);
    const facultyDisplay = faculty
      ? `<span class="faculty-badge" title="Click to change faculty" onclick="openFacultyModalForCode('${s.code}')">👤 ${faculty}</span>`
      : `<button class="btn ghost sm" onclick="openFacultyModalForCode('${s.code}')">+ Assign Faculty</button>`;

    return `
      <tr>
        <td class="code">${s.code}</td>
        <td><b>${s.title}</b></td>
        <td>${facultyDisplay}</td>
        <td class="num-cell">${s.total}</td>
        <td class="num-cell">${s.appeared}</td>
        <td class="num-cell" style="color:var(--pass); font-weight:700;">${s.passed}</td>
        <td class="num-cell" style="color:var(--fail); font-weight:700;">${s.failed}</td>
        <td>
          <div class="bar-cell">
            <div class="bar-track">
              <div class="bar-fill" style="width:${s.pct}%; background:${barColor(s.pct)}"></div>
            </div>
            <span class="pct-text pill ${pctClass(s.pct)}">${s.pct}%</span>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

function renderChart() {
  const canvas = document.getElementById('passChart');
  if (!canvas) return;

  const sorted = subjectSummary.slice().sort((a, b) => a.code.localeCompare(b.code));
  const labels = sorted.map(s => s.code);
  const data = sorted.map(s => s.pct);
  const colors = data.map(pct => barColor(pct));

  if (chartInstance) chartInstance.destroy();

  chartInstance = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Pass %',
        data,
        backgroundColor: colors,
        borderRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: {
          min: 0,
          max: 100,
          title: { display: true, text: 'Pass Percentage (%)' },
          grid: { color: '#e2e8f0' }
        },
        x: {
          grid: { display: false }
        }
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            title: items => {
              const idx = items[0].dataIndex;
              return `${sorted[idx].code}: ${sorted[idx].title}`;
            },
            label: item => {
              const idx = item.dataIndex;
              const fac = getAssignedFaculty(sorted[idx].code);
              return [
                `Pass: ${item.formattedValue}% (${sorted[idx].passed} / ${sorted[idx].appeared})`,
                fac ? `Faculty: ${fac}` : 'Faculty: Not assigned'
              ];
            }
          }
        }
      }
    }
  });
}

function renderLegend() {
  const leg = document.getElementById('gradeLegend');
  if (!leg) return;
  const failGrades = getFailSet().join(', ') || 'None';
  const excludeGrades = getExcludeSet().join(', ') || 'None';
  const withheldGrades = getWithheldSet().join(', ') || 'None';

  leg.innerHTML = `
    <span><b>Current Classification Rules:</b></span>
    <span>Fail: <span class="grade-tag fail">${failGrades}</span></span>
    <span>Exclude: <span class="grade-tag excluded">${excludeGrades}</span></span>
    <span>Withheld: <span class="grade-tag withheld">${withheldGrades}</span></span>
  `;
}

// --- FACULTY ANALYSIS & PERFORMANCE SYSTEM ---
function calculateFacultyPerformance() {
  const failSet = new Set(getFailSet());
  const excludeSet = new Set(getExcludeSet());

  // Aggregate by faculty name
  const facultyRecords = {}; // facultyName -> { subjects: Set, appeared: 0, passed: 0, failed: 0, grades: {}, students: [] }

  rawRows.forEach(row => {
    const fac = getAssignedFaculty(row.code) || 'Unassigned / Multiple';
    if (!facultyRecords[fac]) {
      facultyRecords[fac] = {
        name: fac,
        subjects: new Set(),
        subjectTitles: {},
        appeared: 0,
        passed: 0,
        failed: 0,
        grades: { 'O': 0, 'A+': 0, 'A': 0, 'B+': 0, 'B': 0, 'C': 0, 'P': 0, 'F': 0 },
        students: []
      };
    }
    const rec = facultyRecords[fac];
    rec.subjects.add(row.code);
    rec.subjectTitles[row.code] = row.title;

    if (row.grade === 'W' || excludeSet.has(row.grade)) return;

    rec.appeared++;
    if (failSet.has(row.grade)) {
      rec.failed++;
    } else {
      rec.passed++;
    }

    if (rec.grades[row.grade] !== undefined) {
      rec.grades[row.grade]++;
    }

    rec.students.push({
      usn: row.usn,
      name: row.name,
      total: row.total,
      code: row.code
    });
  });

  return Object.values(facultyRecords).map(f => {
    const pct = f.appeared ? +(f.passed / f.appeared * 100).toFixed(2) : 0;
    // Find highest scoring student in this faculty's subjects
    const topper = f.students.slice().sort((a, b) => b.total - a.total)[0] || null;
    return {
      name: f.name,
      subjectCodes: Array.from(f.subjects).join(', '),
      subjectCount: f.subjects.size,
      appeared: f.appeared,
      passed: f.passed,
      failed: f.failed,
      pct,
      grades: f.grades,
      topper: topper ? `${topper.name} (${topper.usn} — ${topper.total} marks)` : '—'
    };
  });
}

function renderFacultyTab() {
  const perfList = calculateFacultyPerformance();
  if (perfList.length === 0) return;

  // Render Faculty Metric Cards
  const assignedList = perfList.filter(f => f.name !== 'Unassigned / Multiple');
  const totalFaculty = assignedList.length;
  const avgPass = assignedList.length
    ? +(assignedList.reduce((sum, f) => sum + f.pct, 0) / assignedList.length).toFixed(1)
    : 0;

  const topFaculty = assignedList.slice().sort((a, b) => b.pct - a.pct)[0];
  const ninetyPlus = assignedList.filter(f => f.pct >= 90).length;

  const cardsContainer = document.getElementById('facultySummaryCards');
  if (cardsContainer) {
    cardsContainer.innerHTML = `
      <div class="card accent">
        <div class="num">${totalFaculty}</div>
        <div class="lbl">Faculty Allocated</div>
      </div>
      <div class="card pass">
        <div class="num">${avgPass}%</div>
        <div class="lbl">Average Faculty Pass %</div>
      </div>
      <div class="card info">
        <div class="num">${topFaculty ? topFaculty.name.split(' ').slice(0, 2).join(' ') : '—'}</div>
        <div class="lbl">Top Faculty (${topFaculty ? topFaculty.pct + '%' : '—'})</div>
      </div>
      <div class="card pass">
        <div class="num">${ninetyPlus}</div>
        <div class="lbl">Subjects with ≥90% Pass</div>
      </div>
    `;
  }

  // Render Faculty Performance Table
  const tbody = document.getElementById('facultyTableBody');
  if (tbody) {
    const sorted = perfList.sort((a, b) => {
      let av = a[facultySortKey], bv = b[facultySortKey];
      if (typeof av === 'string') { av = av.toLowerCase(); bv = bv.toLowerCase(); }
      if (av < bv) return -1 * facultySortDir;
      if (av > bv) return 1 * facultySortDir;
      return 0;
    });

    tbody.innerHTML = sorted.map((f, idx) => `
      <tr>
        <td style="text-align:center; font-family:var(--font-mono);">${idx + 1}</td>
        <td>
          <div style="font-weight:700; color:var(--navy-900);">${f.name}</div>
        </td>
        <td><span class="code">${f.subjectCodes}</span></td>
        <td class="num-cell">${f.appeared}</td>
        <td class="num-cell" style="color:var(--pass); font-weight:700;">${f.passed}</td>
        <td class="num-cell" style="color:var(--fail); font-weight:700;">${f.failed}</td>
        <td>
          <div class="bar-cell">
            <div class="bar-track">
              <div class="bar-fill" style="width:${f.pct}%; background:${barColor(f.pct)}"></div>
            </div>
            <span class="pct-text pill ${pctClass(f.pct)}">${f.pct}%</span>
          </div>
        </td>
        <td style="font-size:11.5px; font-family:var(--font-mono); text-align:center;">
          <span title="Distinction (O/A+)" style="color:var(--emerald-600); font-weight:700;">${(f.grades['O']||0)+(f.grades['A+']||0)}</span> /
          <span title="First Class (A/B+)" style="color:var(--navy-800);">${(f.grades['A']||0)+(f.grades['B+']||0)}</span> /
          <span title="Pass (B/C/P)" style="color:var(--slate-500);">${(f.grades['B']||0)+(f.grades['C']||0)+(f.grades['P']||0)}</span>
        </td>
        <td style="font-size:12px; color:var(--slate-600);">${f.topper}</td>
      </tr>
    `).join('');
  }

  // Render Faculty Comparison Chart
  const chartCanvas = document.getElementById('facultyChart');
  if (chartCanvas && assignedList.length > 0) {
    if (facultyChartInstance) facultyChartInstance.destroy();

    const chartLabels = assignedList.map(f => f.name.replace(/^(Dr\.|Prof\.)\s*/, ''));
    const chartData = assignedList.map(f => f.pct);
    const chartColors = chartData.map(pct => barColor(pct));

    facultyChartInstance = new Chart(chartCanvas, {
      type: 'bar',
      data: {
        labels: chartLabels,
        datasets: [{
          label: 'Faculty Pass %',
          data: chartData,
          backgroundColor: chartColors,
          borderRadius: 4
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: {
            min: 0,
            max: 100,
            title: { display: true, text: 'Pass Percentage (%)' },
            grid: { color: '#e2e8f0' }
          },
          x: {
            grid: { display: false }
          }
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              title: items => assignedList[items[0].dataIndex].name,
              label: item => {
                const f = assignedList[item.dataIndex];
                return [
                  `Pass: ${f.pct}% (${f.passed}/${f.appeared} students)`,
                  `Subjects: ${f.subjectCodes}`
                ];
              }
            }
          }
        }
      }
    });
  }
}

// --- Faculty Allocation Modal Handler ---
function openFacultyModal() {
  const modal = document.getElementById('facultyModal');
  const tbody = document.getElementById('facultyAllocationBody');
  if (!modal || !tbody) return;

  const distinctSubjects = [];
  const seen = new Set();
  allRows.forEach(r => {
    if (!seen.has(r.code)) {
      seen.add(r.code);
      distinctSubjects.push({ code: r.code, title: r.title, sem: r.sem });
    }
  });
  distinctSubjects.sort((a, b) => a.code.localeCompare(b.code));

  // Existing faculty names for autocomplete suggestions
  const existingNames = new Set();
  Object.values(facultyMap).forEach(m => Object.values(m).forEach(n => { if (n) existingNames.add(n); }));

  let rowsHtml = '';
  distinctSubjects.forEach(s => {
    const currentFaculty = getAssignedFaculty(s.code);
    rowsHtml += `
      <tr>
        <td class="code" style="font-weight:700;">${s.code}</td>
        <td><b>${s.title}</b></td>
        <td>
          <input type="text"
                 class="faculty-input"
                 data-code="${s.code}"
                 value="${currentFaculty}"
                 placeholder="e.g. Dr. S. N. Prasad"
                 list="facultySuggestions"
                 style="width:100%; padding:6px 10px; border:1px solid var(--slate-300); border-radius:var(--radius-sm); font-size:13px;">
        </td>
      </tr>
    `;
  });

  tbody.innerHTML = rowsHtml;

  // Add datalist for suggestions
  let datalist = document.getElementById('facultySuggestions');
  if (!datalist) {
    datalist = document.createElement('datalist');
    datalist.id = 'facultySuggestions';
    document.body.appendChild(datalist);
  }
  datalist.innerHTML = Array.from(existingNames).map(n => `<option value="${n}"></option>`).join('');

  modal.style.display = 'flex';
}

function openFacultyModalForCode(code) {
  openFacultyModal();
  setTimeout(() => {
    const input = document.querySelector(`.faculty-input[data-code="${code}"]`);
    if (input) {
      input.focus();
      input.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  }, 100);
}

function closeFacultyModal() {
  const modal = document.getElementById('facultyModal');
  if (modal) modal.style.display = 'none';
}

function saveFacultyAllocations() {
  document.querySelectorAll('.faculty-input').forEach(input => {
    const code = input.dataset.code;
    const name = input.value.trim();
    setAssignedFaculty(code, name, activeSectionIndex);
  });

  closeFacultyModal();
  buildSummary();
  render();
  renderFacultyTab();
  renderReport();

  showToast('✓ Faculty allocations saved and synchronized successfully!');
}

function handleFacultyFileUpload(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      let codeCol = -1, nameCol = -1;
      for (let r = 0; r < Math.min(grid.length, 10); r++) {
        const row = grid[r].map(normalize);
        const cIdx = row.findIndex(c => c.includes('code') || c.includes('subject'));
        const fIdx = row.findIndex(c => c.includes('faculty') || c.includes('staff') || c.includes('teacher') || c.includes('name'));
        if (cIdx !== -1 && fIdx !== -1 && cIdx !== fIdx) {
          codeCol = cIdx;
          nameCol = fIdx;
          break;
        }
      }

      if (codeCol === -1 || nameCol === -1) {
        throw new Error('Could not identify "Course Code" and "Faculty Name" columns in the file.');
      }

      let count = 0;
      for (let r = 1; r < grid.length; r++) {
        const row = grid[r];
        if (!row || !row[codeCol]) continue;
        const code = String(row[codeCol]).trim();
        const fac = String(row[nameCol] || '').trim();
        if (code && fac) {
          setAssignedFaculty(code, fac, activeSectionIndex);
          count++;
        }
      }

      buildSummary();
      render();
      renderFacultyTab();
      renderReport();

      showToast(`✓ Successfully mapped ${count} subjects to faculty members!`);
      const status = document.getElementById('facultyFileStatus');
      if (status) status.textContent = `Mapped ${count} subjects.`;
    } catch (err) {
      alert('Error parsing faculty mapping: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function downloadFacultyTemplate() {
  const distinctSubjects = [];
  const seen = new Set();
  allRows.forEach(r => {
    if (!seen.has(r.code)) {
      seen.add(r.code);
      distinctSubjects.push([r.code, r.title, getAssignedFaculty(r.code)]);
    }
  });

  const header = ['Course Code', 'Course Title', 'Faculty Name'];
  const rows = [header, ...distinctSubjects];
  const csv = rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  triggerDownload(blob, 'faculty_allocation_template.csv', 'text/csv');
}

// --- STUDENT SEARCH & SCORECARD SYSTEM ---
function searchStudent(query) {
  query = (query || '').toLowerCase().trim();
  const dropdown = document.getElementById('studentSearchResults');
  if (!dropdown) return;

  if (!query || query.length < 2) {
    dropdown.style.display = 'none';
    return;
  }

  const usnMap = new Map();
  allRows.forEach(r => {
    if (!usnMap.has(r.usn)) {
      usnMap.set(r.usn, { usn: r.usn, name: r.name, cgpa: r.cgpa, sgpa: r.sgpa });
    }
  });

  const matches = Array.from(usnMap.values()).filter(s =>
    s.usn.toLowerCase().includes(query) || s.name.toLowerCase().includes(query)
  ).slice(0, 8);

  if (matches.length === 0) {
    dropdown.innerHTML = `<div style="padding:10px 14px; font-size:12.5px; color:var(--muted);">No students matching "${query}"</div>`;
    dropdown.style.display = 'block';
    return;
  }

  dropdown.innerHTML = matches.map(s => `
    <div class="search-result-item" onclick="openStudentScorecard('${s.usn}')" style="padding:8px 14px; cursor:pointer; border-bottom:1px solid var(--slate-100); display:flex; justify-content:space-between; align-items:center;">
      <div>
        <div style="font-weight:700; color:var(--navy-900);">${s.name}</div>
        <div style="font-family:var(--font-mono); font-size:11.5px; color:var(--muted);">${s.usn}</div>
      </div>
      <div style="font-family:var(--font-mono); font-size:12px; font-weight:700; color:var(--gold);">
        SGPA: ${s.sgpa || '—'}
      </div>
    </div>
  `).join('');
  dropdown.style.display = 'block';
}

function openStudentScorecard(usn) {
  const dropdown = document.getElementById('studentSearchResults');
  if (dropdown) dropdown.style.display = 'none';

  const modal = document.getElementById('studentModal');
  if (!modal) return;

  const studentRows = allRows.filter(r => r.usn === usn);
  if (studentRows.length === 0) return;

  const first = studentRows[0];
  const failSet = new Set(getFailSet());
  const failedSubs = studentRows.filter(r => failSet.has(r.grade));
  const hasPassed = failedSubs.length === 0;

  document.getElementById('scStudentName').textContent = first.name || 'Student';
  document.getElementById('scStudentUsn').textContent = first.usn;
  document.getElementById('scStudentSgpa').textContent = first.sgpa ? first.sgpa.toFixed(2) : '—';
  document.getElementById('scStudentCgpa').textContent = first.cgpa ? first.cgpa.toFixed(2) : '—';
  document.getElementById('scStudentStatus').innerHTML = hasPassed
    ? `<span class="pill high">ALL CLEAR (PASSED)</span>`
    : `<span class="pill low">${failedSubs.length} BACKLOG(S)</span>`;

  const tbody = document.getElementById('scSubjectBody');
  tbody.innerHTML = studentRows.map(r => {
    const isFail = failSet.has(r.grade);
    const faculty = getAssignedFaculty(r.code);
    return `
      <tr>
        <td class="code">${r.code}</td>
        <td><b>${r.title}</b></td>
        <td style="font-size:12px; color:var(--slate-600);">${faculty || '—'}</td>
        <td class="num-cell">${r.total}</td>
        <td class="num-cell">${r.credits}</td>
        <td style="text-align:center; font-weight:700; color:${isFail ? 'var(--fail)' : 'var(--pass)'};">${r.grade}</td>
        <td style="text-align:center;">${isFail ? '<span class="pill low">FAIL</span>' : '<span class="pill high">PASS</span>'}</td>
      </tr>
    `;
  }).join('');

  modal.style.display = 'flex';
}

function closeStudentModal() {
  const modal = document.getElementById('studentModal');
  if (modal) modal.style.display = 'none';
}

// --- Section Management & Filter ---
function looksLikeUsn(v) {
  if (!v) return false;
  const s = String(v).trim().toUpperCase();
  return /^\d[A-Z0-9]{2,4}\d{2}[A-Z]{2,4}\d{3,4}$/.test(s) || /^[0-9A-Z]{8,12}$/.test(s);
}

function autoDetectContiguousRanges(usnList) {
  if (!usnList || usnList.length === 0) return { ranges: [], leftover: [] };

  const byPrefix = {};
  const leftover = [];

  usnList.forEach(u => {
    const m = u.match(/^(\d[A-Z0-9]{2,4}\d{2}[A-Z]{2,4})(\d+)$/);
    if (m) {
      const pfx = m[1];
      const num = parseInt(m[2], 10);
      const padLen = m[2].length;
      if (!byPrefix[pfx]) byPrefix[pfx] = [];
      byPrefix[pfx].push({ usn: u, num, padLen });
    } else {
      leftover.push(u);
    }
  });

  const ranges = [];
  Object.keys(byPrefix).forEach(pfx => {
    const list = byPrefix[pfx].sort((a, b) => a.num - b.num);
    let startIdx = 0;
    for (let i = 0; i < list.length; i++) {
      const isEnd = (i === list.length - 1) || (list[i + 1].num !== list[i].num + 1);
      if (isEnd) {
        const runLen = i - startIdx + 1;
        if (runLen >= 3) {
          ranges.push({ start: list[startIdx].usn, end: list[i].usn });
        } else {
          for (let j = startIdx; j <= i; j++) leftover.push(list[j].usn);
        }
        startIdx = i + 1;
      }
    }
  });

  return { ranges, leftover };
}

function handleSectionListFile(file) {
  const status = document.getElementById('sectionListStatus');
  if (status) status.textContent = `Reading ${file.name}…`;

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      if (grid.length < 2) throw new Error('File has no data rows.');

      // Check header row (within first 10 rows)
      let headerIdx = -1, usnCol = -1, secCol = -1, semCol = -1;
      for (let r = 0; r < Math.min(grid.length, 10); r++) {
        const row = grid[r].map(normalize);
        const u = row.findIndex(c => c.includes('usn') || c.includes('roll') || c.includes('reg'));
        const s = row.findIndex(c => c.includes('sec') || c.includes('class') || c.includes('division'));
        const sm = row.findIndex(c => c.includes('sem'));
        if (u !== -1) {
          headerIdx = r;
          usnCol = u;
          secCol = s;
          semCol = sm;
          break;
        }
      }

      // If no explicit header with 'usn' found, search for column with USN-like strings
      if (usnCol === -1) {
        for (let c = 0; c < (grid[0] || []).length; c++) {
          let usnMatchCount = 0;
          for (let r = 0; r < Math.min(grid.length, 20); r++) {
            if (looksLikeUsn(grid[r][c])) usnMatchCount++;
          }
          if (usnMatchCount >= 2) {
            usnCol = c;
            headerIdx = 0;
            break;
          }
        }
      }

      if (usnCol === -1) {
        throw new Error('Could not find a USN column in the uploaded file.');
      }

      // Case 1: Master file with Section column -> creates multiple sections
      if (secCol !== -1) {
        const bySec = {}; // secLabel -> { usns: [], sem: '' }
        for (let r = headerIdx + 1; r < grid.length; r++) {
          const row = grid[r];
          if (!row) continue;
          const usn = String(row[usnCol] || '').trim().toUpperCase();
          const sec = String(row[secCol] || '').trim();
          const sem = semCol !== -1 ? String(row[semCol] || '').trim() : '';
          if (!usn || !sec) continue;

          const secLabel = sec.toLowerCase().includes('sec') ? sec : `${sec} Section`;
          if (!bySec[secLabel]) bySec[secLabel] = { usns: [], sem: sem || '6' };
          bySec[secLabel].usns.push(usn);
        }

        const secEntries = Object.entries(bySec);
        if (secEntries.length === 0) throw new Error('No valid section rows found.');

        savedSections = [];
        secEntries.sort((a, b) => a[0].localeCompare(b[0])).forEach(([secLabel, data]) => {
          const uniqueUsns = Array.from(new Set(data.usns));
          const { ranges, leftover } = autoDetectContiguousRanges(uniqueUsns);
          savedSections.push({
            label: secLabel,
            semester: data.sem,
            ranges,
            customUsns: new Set(leftover)
          });
        });

        activeSectionIndex = 0;
        renderSectionSelect();
        applyActiveSection();

        const summaryText = secEntries.map(([k, v]) => `${k} (${v.usns.length} students)`).join(', ');
        if (status) status.textContent = `✓ Created ${secEntries.length} sections: ${summaryText}`;
        showToast(`✓ Uploaded ${secEntries.length} sections from ${file.name}`);
        return;
      }

      // Case 2: File has just USNs -> add to current section or create Section along with given ranges
      const extractedUsns = [];
      for (let r = headerIdx + 1; r < grid.length; r++) {
        const u = String(grid[r][usnCol] || '').trim().toUpperCase();
        if (u) extractedUsns.push(u);
      }

      if (extractedUsns.length === 0) throw new Error('No USNs found in file.');

      const givenLabel = (document.getElementById('secLabel').value || '').trim() || (file.name.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ') + ' Section');
      const givenSem = (document.getElementById('secSemester').value || '').trim() || '6';
      const givenRangesText = document.getElementById('secRanges').value;
      const givenRanges = parseRangesText(givenRangesText);

      // Populate into section form custom USNs
      const customTextarea = document.getElementById('secCustom');
      const existingCustom = (customTextarea.value || '').split(/[,\n]/).map(s => s.trim().toUpperCase()).filter(Boolean);
      const mergedCustom = Array.from(new Set([...existingCustom, ...extractedUsns]));
      customTextarea.value = mergedCustom.join(', ');

      // Also create/update the section
      savedSections.push({
        label: givenLabel,
        semester: givenSem,
        ranges: givenRanges,
        customUsns: new Set(mergedCustom)
      });
      activeSectionIndex = savedSections.length - 1;
      renderSectionSelect();
      applyActiveSection();

      if (status) status.textContent = `✓ Added ${extractedUsns.length} USNs from ${file.name} to section "${givenLabel}" (combined with ${givenRanges.length} range(s)).`;
      showToast(`✓ Section "${givenLabel}" created with ${extractedUsns.length} USNs and ranges.`);
    } catch (err) {
      alert('Error parsing section file: ' + err.message);
      if (status) status.textContent = 'Error: ' + err.message;
    }
  };
  reader.readAsArrayBuffer(file);
}

function handleSectionFormExcelUpload(file) {
  const status = document.getElementById('sectionFormExcelStatus');
  if (status) status.textContent = `Reading ${file.name}…`;

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      let usnCol = -1, headerIdx = 0;
      for (let r = 0; r < Math.min(grid.length, 10); r++) {
        const row = grid[r].map(normalize);
        const idx = row.findIndex(c => c.includes('usn') || c.includes('roll') || c.includes('reg'));
        if (idx !== -1) { usnCol = idx; headerIdx = r; break; }
      }
      if (usnCol === -1) {
        for (let c = 0; c < (grid[0] || []).length; c++) {
          let count = 0;
          for (let r = 0; r < Math.min(grid.length, 20); r++) {
            if (looksLikeUsn(grid[r][c])) count++;
          }
          if (count >= 2) { usnCol = c; break; }
        }
      }
      if (usnCol === -1) throw new Error('Could not find USN column in file');

      const usns = [];
      for (let r = headerIdx + 1; r < grid.length; r++) {
        const u = String(grid[r][usnCol] || '').trim().toUpperCase();
        if (u) usns.push(u);
      }

      const customTextarea = document.getElementById('secCustom');
      const existing = (customTextarea.value || '').split(/[,\n]/).map(s => s.trim().toUpperCase()).filter(Boolean);
      const combined = Array.from(new Set([...existing, ...usns]));
      customTextarea.value = combined.join(', ');

      if (status) status.textContent = `✓ Loaded ${usns.length} USNs from ${file.name}. Combined with given USN ranges.`;
      showToast(`✓ Loaded ${usns.length} USNs from Excel into custom USNs.`);
    } catch (err) {
      alert('Error: ' + err.message);
      if (status) status.textContent = 'Error: ' + err.message;
    }
  };
  reader.readAsArrayBuffer(file);
}

function downloadSectionTemplate() {
  const header = ['USN', 'Section', 'Semester', 'Student Name'];
  const sampleRows = [
    header,
    ['1BY22IS001', 'A Section', '6', 'Aarav Kumar'],
    ['1BY22IS002', 'A Section', '6', 'Aditi Sharma'],
    ['1BY22IS035', 'B Section', '6', 'Chetan Patil'],
    ['1BY22IS036', 'B Section', '6', 'Deepika Gowda'],
    ['1BY23IS401', 'B Section', '6', 'Karthik Rao (Lateral Entry)']
  ];
  const csv = sampleRows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  triggerDownload(blob, 'section_mapping_template.csv', 'text/csv');
}

function usnInRange(usn, start, end) {
  const numPart = u => {
    const m = u.match(/(\d+)$/);
    return m ? parseInt(m[1], 10) : NaN;
  };
  const prefix = u => u.replace(/\d+$/, '');
  if (prefix(usn) !== prefix(start)) return false;
  const uN = numPart(usn), sN = numPart(start), eN = numPart(end);
  if (isNaN(uN) || isNaN(sN) || isNaN(eN)) return false;
  return uN >= sN && uN <= eN;
}

function parseRangesText(text) {
  const ranges = [];
  (text || '').split('\n').map(l => l.trim()).filter(Boolean).forEach(line => {
    const parts = line.split('-').map(p => p.trim().toUpperCase());
    if (parts.length === 2 && parts[0] && parts[1]) {
      ranges.push({ start: parts[0], end: parts[1] });
    }
  });
  return ranges;
}

function rowsForSection(sec) {
  if (sec.isAll) return allRows.slice();
  if (sec.combined && sec.combined.length) {
    const combinedSets = sec.combined.map(idx => new Set(rowsForSection(savedSections[idx]).map(r => r.usn)));
    const mergedUsns = new Set();
    combinedSets.forEach(s => s.forEach(u => mergedUsns.add(u)));
    return allRows.filter(r => mergedUsns.has(r.usn));
  }
  return allRows.filter(r => {
    if (sec.customUsns && sec.customUsns.has(r.usn)) return true;
    if (sec.ranges && sec.ranges.some(rg => usnInRange(r.usn, rg.start, rg.end))) return true;
    return false;
  });
}

function applyActiveSection() {
  if (activeSectionIndex === -1) {
    rawRows = allRows.slice();
  } else if (savedSections[activeSectionIndex]) {
    rawRows = rowsForSection(savedSections[activeSectionIndex]);
  }
  buildSummary();
  render();
  renderFacultyTab();
  renderReport();
  updateMetaSemFromSection(activeSectionIndex);
}

function renderSectionSelect() {
  const sel = document.getElementById('sectionSelect');
  if (!sel) return;
  let html = '<option value="-1">All students (' + new Set(allRows.map(r => r.usn)).size + ' students)</option>';
  savedSections.forEach((s, idx) => {
    const cnt = new Set(rowsForSection(s).map(r => r.usn)).size;
    html += `<option value="${idx}" ${idx === activeSectionIndex ? 'selected' : ''}>${s.label} (${cnt} students)</option>`;
  });
  sel.innerHTML = html;

  const delBtn = document.getElementById('deleteSectionBtn');
  if (delBtn) delBtn.style.display = activeSectionIndex >= 0 ? 'inline-block' : 'none';
  renderCombineCheckboxes();
}

function renderCombineCheckboxes() {
  const wrap = document.getElementById('combineSectionsPanel');
  const box = document.getElementById('combineCheckboxes');
  if (!wrap || !box) return;

  const valid = savedSections.filter(s => !s.combined && !s.isAll);
  if (valid.length < 2) {
    wrap.style.display = 'none';
    return;
  }
  wrap.style.display = 'flex';
  box.innerHTML = valid.map((s, idx) => `
    <label style="font-size:12.5px; display:inline-flex; align-items:center; gap:4px;">
      <input type="checkbox" value="${idx}" class="combine-sec-cb"> ${s.label}
    </label>
  `).join('');
}

function updateMetaSemFromSection(index) {
  const metaSem = document.getElementById('metaSem');
  if (!metaSem) return;
  if (index === -1) {
    metaSem.value = 'Semester – 6 (All Sections)';
  } else if (savedSections[index]) {
    const sec = savedSections[index];
    metaSem.value = `Semester – ${sec.semester || '6'} (${sec.label})`;
  }
}

// --- Official Report Renderers ---
function populateAcademicYearOptions() {
  const sel = document.getElementById('metaAcademicYear');
  if (!sel) return;
  const curYear = new Date().getFullYear();
  let html = '';
  for (let y = curYear + 1; y >= curYear - 4; y--) {
    const opt = `${y - 1} - ${y}`;
    html += `<option value="${opt}">${opt}</option>`;
  }
  sel.innerHTML = html;
}

function renderReportHeader() {
  const inst = document.getElementById('metaInst').value;
  const dept = document.getElementById('metaDept').value;
  const title = document.getElementById('metaTitle').value;
  const sem = document.getElementById('metaSem').value;
  const ay = document.getElementById('metaAcademicYear').value;

  document.getElementById('rInst').textContent = inst;
  document.getElementById('rDept').textContent = dept;
  document.getElementById('rTitle').textContent = title;
  document.getElementById('rAcademicYear').textContent = `Academic Year: ${ay}`;
  document.getElementById('rSem').textContent = sem;
}

function extractBatch(usn) {
  const m = usn.match(/^\d[A-Z0-9]{2}(\d{2})/);
  if (!m) return 'Unknown';
  return '20' + m[1];
}

function studentBuckets(usn, failSet) {
  const sRows = rawRows.filter(r => r.usn === usn);
  const grades = sRows.map(r => r.grade);
  const excludeSet = new Set(getExcludeSet());

  const hasFail = grades.some(g => failSet.has(g));
  const hasAbsent = grades.some(g => g === 'AB');
  const hasWithdrawn = grades.some(g => g === 'W');
  const isExcluded = grades.every(g => excludeSet.has(g));

  return { hasFail, hasAbsent, hasWithdrawn, isExcluded };
}

function renderOverallTable() {
  const failSet = new Set(getFailSet());
  const byBatch = {};

  const usns = Array.from(new Set(rawRows.map(r => r.usn)));
  usns.forEach(u => {
    const b = extractBatch(u);
    if (!byBatch[b]) byBatch[b] = { batch: b, reg: 0, app: 0, pass: 0, fail: 0, ab: 0, w: 0 };
    byBatch[b].reg++;
    const { hasFail, hasAbsent, hasWithdrawn, isExcluded } = studentBuckets(u, failSet);
    if (!isExcluded) byBatch[b].app++;
    if (!hasFail && !isExcluded) byBatch[b].pass++;
    if (hasFail) byBatch[b].fail++;
    if (hasAbsent) byBatch[b].ab++;
    if (hasWithdrawn) byBatch[b].w++;
  });

  const batches = Object.values(byBatch).sort((a, b) => a.batch.localeCompare(b.batch));
  let rows = `<tr>
    <th>Sl.No</th><th>Batch</th><th>Class Strength</th><th>Appeared</th><th>Passed</th><th>Failed</th><th>Absent</th><th>Withdrawn</th><th>% of Pass</th>
  </tr>`;

  let gReg = 0, gApp = 0, gPass = 0, gFail = 0, gAb = 0, gW = 0;
  batches.forEach((b, idx) => {
    const pct = b.app ? (b.pass / b.app * 100).toFixed(2) : '0.00';
    gReg += b.reg; gApp += b.app; gPass += b.pass; gFail += b.fail; gAb += b.ab; gW += b.w;
    rows += `<tr>
      <td>${idx + 1}</td>
      <td>${b.batch}</td>
      <td>${b.reg}</td>
      <td>${b.app}</td>
      <td>${b.pass}</td>
      <td>${b.fail}</td>
      <td>${b.ab}</td>
      <td>${b.w}</td>
      <td>${pct}%</td>
    </tr>`;
  });

  const totalPct = gApp ? (gPass / gApp * 100).toFixed(2) : '0.00';
  rows += `<tr style="font-weight:700; background:#f7f9fc;">
    <td colspan="2">Total</td>
    <td>${gReg}</td><td>${gApp}</td><td>${gPass}</td><td>${gFail}</td><td>${gAb}</td><td>${gW}</td><td>${totalPct}%</td>
  </tr>`;

  document.getElementById('overallTable').innerHTML = rows;
}

function renderSubjectTable() {
  const sorted = subjectSummary.slice().sort((a, b) => a.code.localeCompare(b.code));
  let rows = `<tr>
    <th>Sl.No</th>
    <th>Subject Title</th>
    <th>Subject Code</th>
    <th>Staff Handling</th>
    <th>Class Strength</th>
    <th>Appeared</th>
    <th>Passed</th>
    <th>Failed</th>
    <th>F</th>
    <th>AB</th>
    <th>DX</th>
    <th>X</th>
    <th>I</th>
    <th>W</th>
    <th>% of Pass</th>
  </tr>`;

  sorted.forEach((s, idx) => {
    const faculty = getAssignedFaculty(s.code);
    rows += `<tr>
      <td>${idx + 1}</td>
      <td class="left"><b>${s.title}</b></td>
      <td class="code">${s.code}</td>
      <td class="left">
        <textarea data-code="${s.code}" data-field="staff" rows="1" placeholder="Staff Name">${faculty}</textarea>
      </td>
      <td>${s.total}</td>
      <td>${s.appeared}</td>
      <td>${s.passed}</td>
      <td>${s.failed}</td>
      <td>${s.f}</td>
      <td>${s.ab}</td>
      <td>${s.dx}</td>
      <td>${s.x}</td>
      <td>${s.i}</td>
      <td>${s.w}</td>
      <td style="font-weight:700;">${s.pct}%</td>
    </tr>`;
  });

  const totalW = sorted.reduce((sum, s) => sum + (s.w || 0), 0);
  rows += `<tr style="font-weight:700; background:#f7f9fc;">
    <td colspan="13" style="text-align:right;">Total Withdrawn (W):</td>
    <td style="font-weight:700; color:var(--gold);">${totalW}</td>
    <td></td>
  </tr>`;

  const subTable = document.getElementById('subjectTable');
  if (subTable) {
    subTable.innerHTML = rows;
    subTable.querySelectorAll('textarea[data-field="staff"]').forEach(ta => {
      ta.addEventListener('input', () => {
        setAssignedFaculty(ta.dataset.code, ta.value, activeSectionIndex);
        renderTable();
        renderFacultyTab();
      });
    });
  }
}

function renderReportFacultySection() {
  const perfList = calculateFacultyPerformance();
  const assignedList = perfList.filter(f => f.name !== 'Unassigned / Multiple');
  const table = document.getElementById('reportFacultyTable');
  if (!table) return;

  if (assignedList.length === 0) {
    document.getElementById('reportFacultySection').style.display = 'none';
    return;
  }
  document.getElementById('reportFacultySection').style.display = 'block';

  let rows = `<tr>
    <th>Sl.No</th>
    <th>Faculty Name</th>
    <th>Subject(s) Handled</th>
    <th>Students Appeared</th>
    <th>Students Passed</th>
    <th>Students Failed</th>
    <th>Pass %</th>
    <th>Distinction (O / A+)</th>
    <th>Subject Topper</th>
  </tr>`;

  assignedList.sort((a, b) => b.pct - a.pct).forEach((f, idx) => {
    rows += `<tr>
      <td>${idx + 1}</td>
      <td class="left"><b>${f.name}</b></td>
      <td class="left">${f.subjectCodes}</td>
      <td>${f.appeared}</td>
      <td>${f.passed}</td>
      <td>${f.failed}</td>
      <td style="font-weight:700; color:${f.pct >= 85 ? 'var(--pass)' : f.pct >= 60 ? 'var(--gold)' : 'var(--fail)'}">${f.pct}%</td>
      <td>${(f.grades['O'] || 0) + (f.grades['A+'] || 0)}</td>
      <td class="left" style="font-size:11px;">${f.topper}</td>
    </tr>`;
  });

  table.innerHTML = rows;
}

function renderTopperTable() {
  const usnMap = new Map();
  rawRows.forEach(r => {
    if (!usnMap.has(r.usn)) {
      usnMap.set(r.usn, { usn: r.usn, name: r.name, sgpa: r.sgpa || 0, cgpa: r.cgpa || 0, totalMarks: 0, coursesCount: 0 });
    }
    const st = usnMap.get(r.usn);
    if (r.total) {
      st.totalMarks += r.total;
      st.coursesCount++;
    }
  });

  const students = Array.from(usnMap.values()).map(s => ({
    ...s,
    agg: s.coursesCount ? +(s.totalMarks / s.coursesCount).toFixed(2) : 0
  }));

  // Rank by SGPA desc, then agg desc
  students.sort((a, b) => (b.sgpa - a.sgpa) || (b.agg - a.agg));
  const top10 = students.slice(0, 10);

  let rows = `<tr>
    <th>Rank</th><th>USN</th><th>Name of Student</th><th>SGPA</th><th>Aggregate %</th>
  </tr>`;

  top10.forEach((s, idx) => {
    rows += `<tr>
      <td><b>#${idx + 1}</b></td>
      <td class="code">${s.usn}</td>
      <td class="left"><b>${s.name}</b></td>
      <td style="font-weight:700; color:var(--navy-900);">${s.sgpa.toFixed(2)}</td>
      <td>${s.agg}%</td>
    </tr>`;
  });

  document.getElementById('topperTable').innerHTML = rows;
}

function renderFailedList() {
  const failSet = new Set(getFailSet());
  const failedMap = new Map();

  rawRows.forEach(r => {
    if (failSet.has(r.grade)) {
      if (!failedMap.has(r.usn)) {
        failedMap.set(r.usn, { usn: r.usn, name: r.name, subjects: [] });
      }
      failedMap.get(r.usn).subjects.push(`${r.code} (${r.title}) [Grade: ${r.grade}]`);
    }
  });

  const list = Array.from(failedMap.values()).sort((a, b) => a.usn.localeCompare(b.usn));
  const table = document.getElementById('failedListTable');
  if (!table) return;

  if (list.length === 0) {
    table.innerHTML = `<tr><td colspan="4" style="padding:14px; text-align:center; color:var(--pass); font-weight:600;">No failed students in this scope! (100% Pass)</td></tr>`;
    return;
  }

  let rows = `<tr>
    <th>Sl.No</th><th>USN</th><th>Name of Student</th><th>Failed Subject(s)</th>
  </tr>`;

  list.forEach((s, idx) => {
    rows += `<tr>
      <td>${idx + 1}</td>
      <td class="code">${s.usn}</td>
      <td class="left"><b>${s.name}</b></td>
      <td class="left" style="color:var(--fail);">${s.subjects.join('<br>')}</td>
    </tr>`;
  });

  table.innerHTML = rows;
}

function renderSubjectFailList() {
  const failSet = new Set(getFailSet());
  const bySub = new Map();

  rawRows.forEach(r => {
    if (failSet.has(r.grade)) {
      const key = `${r.code} - ${r.title}`;
      if (!bySub.has(key)) bySub.set(key, []);
      bySub.get(key).push({ usn: r.usn, name: r.name, grade: r.grade });
    }
  });

  const table = document.getElementById('subjectFailTable');
  if (!table) return;

  if (bySub.size === 0) {
    table.innerHTML = `<tr><td colspan="4" style="padding:14px; text-align:center; color:var(--pass); font-weight:600;">No subject failures!</td></tr>`;
    return;
  }

  let html = '';
  Array.from(bySub.keys()).sort().forEach(subKey => {
    const students = bySub.get(subKey).sort((a, b) => a.usn.localeCompare(b.usn));
    html += `
      <tr style="background:#eef2f8; font-weight:700;"><td colspan="4" class="left">${subKey} (${students.length} students failed)</td></tr>
      <tr><th>Sl.No</th><th>USN</th><th>Name of Student</th><th>Grade</th></tr>
    `;
    students.forEach((st, idx) => {
      html += `<tr>
        <td>${idx + 1}</td>
        <td class="code">${st.usn}</td>
        <td class="left">${st.name}</td>
        <td style="color:var(--fail); font-weight:700;">${st.grade}</td>
      </tr>`;
    });
  });

  table.innerHTML = html;
}

function renderClassificationTable() {
  const failSet = new Set(getFailSet());
  const usnSet = new Set(rawRows.map(r => r.usn));

  let fcd = 0, fc = 0, sc = 0, passClass = 0, fail = 0;

  usnSet.forEach(u => {
    const sRows = rawRows.filter(r => r.usn === u);
    const hasFail = sRows.some(r => failSet.has(r.grade) || r.grade === 'AB');
    if (hasFail) {
      fail++;
      return;
    }
    const cgpa = sRows[0].cgpa || sRows[0].sgpa || 0;
    const m = cgpa * 10;
    if (m >= 70) fcd++;
    else if (m >= 60) fc++;
    else if (m >= 50) sc++;
    else if (m >= 40) passClass++;
    else fail++;
  });

  const total = usnSet.size;
  const pct = cnt => total ? (cnt / total * 100).toFixed(2) + '%' : '0.00%';

  document.getElementById('classificationTable').innerHTML = `
    <tr><th>Classification</th><th>Criteria</th><th>No. of Students</th><th>Percentage</th></tr>
    <tr><td class="left"><b>First Class with Distinction (FCD)</b></td><td>Percentage ≥ 70%</td><td>${fcd}</td><td>${pct(fcd)}</td></tr>
    <tr><td class="left"><b>First Class (FC)</b></td><td>60% ≤ Percentage &lt; 70%</td><td>${fc}</td><td>${pct(fc)}</td></tr>
    <tr><td class="left"><b>Second Class (SC)</b></td><td>50% ≤ Percentage &lt; 60%</td><td>${sc}</td><td>${pct(sc)}</td></tr>
    <tr><td class="left"><b>Pass Class</b></td><td>40% ≤ Percentage &lt; 50%</td><td>${passClass}</td><td>${pct(passClass)}</td></tr>
    <tr><td class="left" style="color:var(--fail);"><b>Fail</b></td><td>Backlog / Absent</td><td>${fail}</td><td>${pct(fail)}</td></tr>
    <tr style="font-weight:700; background:#f7f9fc;"><td colspan="2" class="left">Total</td><td>${total}</td><td>100.00%</td></tr>
  `;
}

function renderSpecialGradesList() {
  const specialList = [];
  rawRows.forEach(r => {
    if (['AB', 'I', 'W', 'DX', 'X'].includes(r.grade)) {
      specialList.push(r);
    }
  });

  const table = document.getElementById('specialGradesTable');
  if (!table) return;

  if (specialList.length === 0) {
    table.innerHTML = `<tr><td colspan="5" style="padding:14px; text-align:center; color:var(--muted);">No special grades (AB, I, W, DX, X) found in this scope.</td></tr>`;
    return;
  }

  let rows = `<tr>
    <th>Sl.No</th><th>USN</th><th>Name of Student</th><th>Subject</th><th>Grade</th>
  </tr>`;

  specialList.forEach((r, idx) => {
    rows += `<tr>
      <td>${idx + 1}</td>
      <td class="code">${r.usn}</td>
      <td class="left"><b>${r.name}</b></td>
      <td class="left">${r.code} - ${r.title}</td>
      <td><span class="grade-tag ${r.grade === 'AB' ? 'fail' : 'withheld'}">${r.grade}</span></td>
    </tr>`;
  });

  table.innerHTML = rows;
}

function parseQuotaGroups() {
  const el = document.getElementById('quotaGroupInput');
  const text = el ? el.value : '';
  const groups = [];
  text.split('\n').forEach(line => {
    const idx = line.indexOf('=');
    if (idx === -1) return;
    const label = line.slice(0, idx).trim();
    const raws = line.slice(idx + 1).split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
    if (label && raws.length) groups.push({ label, raws: new Set(raws) });
  });
  return groups;
}

function groupForCategory(cat, groups) {
  const norm = String(cat).trim().toUpperCase();
  for (const g of groups) if (g.raws.has(norm)) return g.label;
  return 'Other';
}

function renderQuotaTable() {
  const sec = document.getElementById('quotaSection');
  const tbl = document.getElementById('quotaTable');
  if (!sec || !tbl) return;
  if (!categoryMap) { sec.style.display = 'none'; return; }

  const failSet = new Set(getFailSet());
  const groups = parseQuotaGroups();
  const usns = [...new Set(rawRows.map(r => r.usn))];
  const byQuota = {};

  usns.forEach(u => {
    const rawCat = categoryMap[u] || 'Unmapped';
    const label = rawCat === 'Unmapped' ? 'Unmapped' : groupForCategory(rawCat, groups);
    if (!byQuota[label]) byQuota[label] = { pass: 0, fail: 0, absent: 0, withdrawn: 0, total: 0 };
    const buckets = studentBuckets(u, failSet);
    byQuota[label].total++;
    if (!buckets.hasFail && !buckets.isExcluded) byQuota[label].pass++;
    if (buckets.hasFail) byQuota[label].fail++;
    if (buckets.hasAbsent) byQuota[label].absent++;
    if (buckets.hasWithdrawn) byQuota[label].withdrawn++;
  });

  const orderedLabels = [...groups.map(g => g.label), 'Other', 'Unmapped'].filter(l => byQuota[l]);
  let rows = '<tr><th>Quota</th><th>Pass</th><th>Fail</th><th>Absent</th><th>Withdrawn</th><th>Total</th></tr>';
  let tp = 0, tf = 0, ta = 0, tw = 0, tt = 0;
  orderedLabels.forEach(label => {
    const d = byQuota[label];
    tp += d.pass; tf += d.fail; ta += d.absent; tw += d.withdrawn; tt += d.total;
    rows += `<tr><td class="left"><b>${label}</b></td><td>${d.pass}</td><td>${d.fail}</td><td>${d.absent}</td><td>${d.withdrawn}</td><td>${d.total}</td></tr>`;
  });
  rows += `<tr style="font-weight:700; background:#f7f9fc;"><td class="left">Total</td><td>${tp}</td><td>${tf}</td><td>${ta}</td><td>${tw}</td><td>${tt}</td></tr>`;
  tbl.innerHTML = rows;
  sec.style.display = 'block';
  const qBox = document.getElementById('quotaGroupBox');
  if (qBox) qBox.style.display = 'block';
}

function renderCategoryTable() {
  const sec = document.getElementById('categorySection');
  const tbl = document.getElementById('categoryTable');
  if (!sec || !tbl) return;
  if (!categoryMap) { sec.style.display = 'none'; return; }

  const failSet = new Set(getFailSet());
  const usns = [...new Set(rawRows.map(r => r.usn))];
  const byCat = {};
  usns.forEach(u => {
    const cat = categoryMap[u] || 'Unmapped';
    if (!byCat[cat]) byCat[cat] = { appeared: 0, pass: 0, fail: 0, absent: 0, withdrawn: 0 };
    const buckets = studentBuckets(u, failSet);
    byCat[cat].appeared++;
    if (!buckets.hasFail && !buckets.isExcluded) byCat[cat].pass++;
    if (buckets.hasFail) byCat[cat].fail++;
    if (buckets.hasAbsent) byCat[cat].absent++;
    if (buckets.hasWithdrawn) byCat[cat].withdrawn++;
  });

  const cats = Object.keys(byCat).sort();
  let rows = '<tr><th>Category</th>' + cats.map(c => `<th>${c}</th>`).join('') + '<th>Total</th></tr>';
  const metric = (label, fn) => {
    let tot = 0;
    const cells = cats.map(c => { const v = fn(byCat[c]); tot += v; return `<td>${v}</td>`; }).join('');
    return `<tr><td class="left">${label}</td>${cells}<td>${tot}</td></tr>`;
  };
  rows += metric('Appeared', d => d.appeared);
  rows += metric('Passed', d => d.pass);
  rows += metric('Failed', d => d.fail);
  rows += metric('Absent', d => d.absent);
  rows += metric('Withdrawn', d => d.withdrawn);
  rows += `<tr><td class="left"><b>% of Pass</b></td>${cats.map(c => { const d = byCat[c]; return `<td><b>${d.appeared ? (d.pass / d.appeared * 100).toFixed(1) : '0.0'}%</b></td>`; }).join('')}<td><b>${(Object.values(byCat).reduce((a, d) => a + d.pass, 0) / (Object.values(byCat).reduce((a, d) => a + d.appeared, 0) || 1) * 100).toFixed(1)}%</b></td></tr>`;
  tbl.innerHTML = rows;
  sec.style.display = 'block';
}

function handleCategoryFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  e.target.value = '';
  categoryFileName = file.name;
  const status = document.getElementById('categoryStatus');
  if (status) status.textContent = 'Reading…';
  const reader = new FileReader();
  reader.onload = evt => {
    try {
      const data = new Uint8Array(evt.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
      const header = grid[0].map(h => String(h).trim().toLowerCase());
      const usnCol = header.findIndex(h => h.includes('usn'));
      const catCol = header.findIndex(h => h.includes('categ'));
      if (usnCol === -1 || catCol === -1) throw new Error('Need columns named USN and Category');
      categoryMap = {};
      for (let i = 1; i < grid.length; i++) {
        const row = grid[i];
        if (!row[usnCol]) continue;
        categoryMap[String(row[usnCol]).trim().toUpperCase()] = String(row[catCol]).trim();
      }
      if (status) status.textContent = `✓ Loaded ${Object.keys(categoryMap).length} USN→category mappings.`;
      renderQuotaTable();
      renderCategoryTable();
      showToast('✓ Category mapping applied successfully!');
    } catch (err) {
      if (status) status.textContent = 'Error: ' + err.message;
    }
  };
  reader.readAsArrayBuffer(file);
}

function renderReport() {
  renderReportHeader();
  renderOverallTable();
  renderSubjectTable();
  renderReportFacultySection();
  renderTopperTable();
  renderQuotaTable();
  renderCategoryTable();
  renderFailedList();
  renderSubjectFailList();
  renderClassificationTable();
  renderSpecialGradesList();
}

// --- Revaluation Processing ---
function applyRevaluation(revalRows, failSet) {
  let matched = 0, unmatched = 0;
  const changes = [];
  const usnNewScores = {};

  revalRows.forEach(rv => {
    const idx = allRows.findIndex(r => r.usn === rv.usn && r.code === rv.code);
    if (idx === -1) { unmatched++; return; }
    const old = allRows[idx];
    const gradeChanged = old.grade !== rv.grade;
    const totalChanged = old.total !== rv.total;
    if (gradeChanged || totalChanged) {
      const wasFail = failSet.has(old.grade);
      const isFail = failSet.has(rv.grade);
      let statusChange = '—';
      if (wasFail && !isFail) statusChange = 'Fail → Pass';
      else if (!wasFail && isFail) statusChange = 'Pass → Fail';
      changes.push({
        usn: rv.usn, name: old.name || rv.name, code: rv.code, title: old.title || rv.title,
        oldGrade: old.grade, newGrade: rv.grade, oldTotal: old.total, newTotal: rv.total, statusChange
      });
    }
    allRows[idx] = { ...old, grade: rv.grade, total: rv.total || old.total, credits: rv.credits || old.credits };
    matched++;
    if (rv.sgpa) usnNewScores[rv.usn] = { sgpa: rv.sgpa, cgpa: rv.cgpa };
  });

  Object.keys(usnNewScores).forEach(u => {
    allRows.forEach(r => { if (r.usn === u) { r.sgpa = usnNewScores[u].sgpa; r.cgpa = usnNewScores[u].cgpa; } });
  });

  return { matched, unmatched, changes };
}

function renderRevalLog() {
  const wrap = document.getElementById('revalLogWrap');
  const tbl = document.getElementById('revalLogTable');
  if (!wrap || !tbl) return;
  if (revalLog.length === 0) { wrap.style.display = 'none'; return; }

  wrap.style.display = 'block';
  let html = `<tr><th>USN</th><th>Name</th><th>Subject</th><th>Old Grade</th><th>New Grade</th><th>Old Total</th><th>New Total</th><th>Status Change</th></tr>`;
  revalLog.forEach(c => {
    const isGood = c.statusChange.includes('Pass');
    html += `<tr>
      <td class="code">${c.usn}</td>
      <td class="left">${c.name}</td>
      <td class="left">${c.code} - ${c.title}</td>
      <td>${c.oldGrade}</td>
      <td style="font-weight:700;">${c.newGrade}</td>
      <td>${c.oldTotal}</td>
      <td style="font-weight:700;">${c.newTotal}</td>
      <td style="font-weight:700; color:${isGood ? 'var(--pass)' : 'var(--fail)'};">${c.statusChange}</td>
    </tr>`;
  });
  tbl.innerHTML = html;
}

// --- Word (.docx) Native Export ---
function base64ToUint8Array(base64) {
  const binaryString = window.atob(base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) bytes[i] = binaryString.charCodeAt(i);
  return bytes;
}

function docxCellText(cell) {
  const input = cell.querySelector('input');
  if (input) return input.value || '';
  const textarea = cell.querySelector('textarea');
  if (textarea) return textarea.value || '';
  return cell.textContent.replace(/\s+/g, ' ').trim();
}

function domTableToDocxTable(tableEl) {
  const { Table, TableRow, TableCell, Paragraph, TextRun, AlignmentType, WidthType } = docx;
  const rows = [...tableEl.rows].map(tr => {
    const cells = [...tr.cells].map(cell => {
      const isHeader = cell.tagName === 'TH';
      const colSpan = Number(cell.getAttribute('colspan')) || 1;
      const isLeft = cell.classList.contains('left') || /text-align:\s*left/.test(cell.getAttribute('style') || '');
      const isGroupHeader = /background/.test(cell.getAttribute('style') || '') && colSpan > 1;
      const bold = isHeader || isGroupHeader || /font-weight:\s*(bold|700)/.test((tr.getAttribute('style') || '') + (cell.getAttribute('style') || ''));
      const shaded = isHeader || isGroupHeader || /background/.test(tr.getAttribute('style') || '');
      return new TableCell({
        columnSpan: colSpan > 1 ? colSpan : undefined,
        shading: shaded ? { fill: 'EEF2F8' } : undefined,
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        children: [new Paragraph({
          alignment: isLeft ? AlignmentType.LEFT : AlignmentType.CENTER,
          children: [new TextRun({ text: docxCellText(cell), bold, size: 18 })]
        })]
      });
    });
    return new TableRow({ children: cells });
  });
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows });
}

function docxSigCell(name, role) {
  const { TableCell, Paragraph, TextRun, AlignmentType, BorderStyle } = docx;
  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return new TableCell({
    borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
    margins: { top: 100 },
    children: [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400 }, children: [new TextRun({ text: name || '', bold: true, size: 18 })] }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        indent: { left: 400, right: 400 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '0B1F3A', space: 2 } },
        children: [new TextRun('')]
      }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60 }, children: [new TextRun({ text: role, bold: true, size: 16 })] })
    ]
  });
}

async function buildDocxBlob() {
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, ImageRun, Table, TableRow, BorderStyle } = docx;

  const dept = document.getElementById('metaDept').value;
  const title = document.getElementById('metaTitle').value;
  const sem = document.getElementById('metaSem').value;
  const academicYear = document.getElementById('metaAcademicYear').value;

  const children = [];

  try {
    const bytes = base64ToUint8Array(BMSIT_LOGO_DATA_URI.split(',')[1]);
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new ImageRun({ data: bytes, transformation: { width: 440, height: Math.round(440 * 206 / 816) } })]
    }));
  } catch (e) { /* continue without image if decoding fails */ }

  const headLine = (text, runOpts, paraOpts) => new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 40, ...(paraOpts || {}) },
    children: [new TextRun({ text, ...(runOpts || {}) })]
  });

  children.push(headLine(dept, { bold: true, size: 26 }));
  children.push(headLine(title, { bold: true, size: 22, color: '13315C' }));
  children.push(headLine(`Academic Year: ${academicYear}`, { size: 20, color: '5B6B85' }));
  children.push(headLine(sem, { bold: true, size: 20 }, { spacing: { after: 200 } }));

  function addSection(titleText, tableEl) {
    if (!tableEl) return;
    children.push(new Paragraph({
      text: titleText,
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 300, after: 120 }
    }));
    children.push(domTableToDocxTable(tableEl));
  }

  addSection('1. Overall Result Analysis', document.getElementById('overallTable'));
  addSection('2. Subject-wise Result Analysis', document.getElementById('subjectTable'));

  const rFac = document.getElementById('reportFacultyTable');
  if (rFac && rFac.rows.length > 1) {
    addSection('3. Faculty-wise Result Analysis', rFac);
  }

  addSection('4. Topper\'s List', document.getElementById('topperTable'));
  addSection('5. List of Failed Students', document.getElementById('failedListTable'));
  addSection('6. Subject-wise List of Failed Students', document.getElementById('subjectFailTable'));
  addSection('7. Classification of Results', document.getElementById('classificationTable'));
  addSection('8. Students with Special Grades (AB, I, W, DX)', document.getElementById('specialGradesTable'));

  // Signatures Table
  const sigRow = new TableRow({
    children: [
      docxSigCell(sigNames.classTeacher, 'Signature of Class Teacher'),
      docxSigCell(sigNames.coordinator, 'Result Analysis Coordinator'),
      docxSigCell(sigNames.divisionHead, 'Division Head'),
      docxSigCell(sigNames.hod, 'HoD')
    ]
  });
  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  children.push(new Paragraph({ spacing: { before: 400 } }));
  children.push(new Table({
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
    borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
    rows: [sigRow]
  }));

  const doc = new Document({
    sections: [{
      properties: { page: { margin: { top: 720, bottom: 720, left: 720, right: 720 } } },
      children
    }]
  });

  return await Packer.toBlob(doc);
}

// --- Toast & Download Helpers ---
function showToast(msg) {
  const toast = document.getElementById('downloadToast');
  if (!toast) return;
  toast.innerHTML = msg;
  toast.style.display = 'block';
  clearTimeout(window.__toastTimer);
  window.__toastTimer = setTimeout(() => { toast.style.display = 'none'; }, 5000);
}

async function triggerDownload(blob, filename, mimeType) {
  if (typeof window.showSaveFilePicker === 'function') {
    try {
      const ext = '.' + filename.split('.').pop();
      const handle = await window.showSaveFilePicker({
        suggestedName: filename,
        types: [{ description: 'File', accept: { [mimeType || 'application/octet-stream']: [ext] } }]
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return 'picked';
    } catch (err) {
      if (err && err.name === 'AbortError') return 'cancelled';
    }
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  return 'downloaded';
}

// --- LocalStorage Persistence ---
function snapshotState() {
  return {
    savedAt: new Date().toISOString(),
    title: document.getElementById('metaTitle').value || 'Untitled report',
    meta: {
      inst: document.getElementById('metaInst').value,
      dept: document.getElementById('metaDept').value,
      title: document.getElementById('metaTitle').value,
      sem: document.getElementById('metaSem').value,
      academicYear: document.getElementById('metaAcademicYear').value
    },
    allRows,
    savedSections,
    activeSectionIndex,
    subjectOverrides,
    facultyMap,
    overallWithheldMap,
    sigNames,
    categoryMap,
    categoryFileName,
    revalLog,
    failGrades: document.getElementById('failGradesInput').value,
    excludeGrades: document.getElementById('excludeGradesInput').value
  };
}

function saveReportToStorage() {
  try {
    const snap = snapshotState();
    const key = 'ra_report_' + Date.now();
    localStorage.setItem(key, JSON.stringify(snap));
    renderSavedReportsList();
    showToast(`✓ Saved report "${snap.title}" successfully.`);
  } catch (err) {
    alert('Could not save to localStorage: ' + err.message);
  }
}

function renderSavedReportsList() {
  const containers = [
    document.getElementById('savedReportsList'),
    document.getElementById('savedReportsListOuter')
  ].filter(Boolean);

  const items = [];
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key && key.startsWith('ra_report_')) {
      try {
        const snap = JSON.parse(localStorage.getItem(key));
        items.push({
          key,
          title: snap.title || 'Untitled Report',
          savedAt: snap.savedAt,
          studentCount: snap.allRows ? new Set(snap.allRows.map(r => r.usn)).size : 0
        });
      } catch (e) {}
    }
  }

  items.sort((a, b) => (b.savedAt || '').localeCompare(a.savedAt || ''));

  if (items.length === 0) {
    containers.forEach(c => c.innerHTML = '<div style="font-size:12.5px; color:var(--muted); padding:6px 0;">No saved reports in browser storage.</div>');
    return;
  }

  const html = items.map(it => `
    <div class="saved-report-item" style="display:flex; justify-content:space-between; align-items:center; padding:8px 12px; border:1px solid var(--slate-200); border-radius:var(--radius-sm); margin-bottom:6px; background:#fff;">
      <div>
        <div style="font-weight:700; color:var(--navy-900); font-size:13px;">${it.title}</div>
        <div style="font-size:11px; color:var(--muted); font-family:var(--font-mono);">${new Date(it.savedAt).toLocaleString()} · ${it.studentCount} students</div>
      </div>
      <div style="display:flex; gap:6px;">
        <button class="btn sm" onclick="loadReportFromStorage('${it.key}')">Load</button>
        <button class="btn ghost sm danger" onclick="deleteReportFromStorage('${it.key}')">Delete</button>
      </div>
    </div>
  `).join('');

  containers.forEach(c => c.innerHTML = html);
}

function loadReportFromStorage(key) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return;
    const snap = JSON.parse(raw);
    applySnapshot(snap, '(loaded from browser storage)');
    showToast(`✓ Loaded "${snap.title}" successfully.`);
  } catch (err) {
    alert('Could not load report: ' + err.message);
  }
}

function deleteReportFromStorage(key) {
  if (confirm('Delete this saved report?')) {
    localStorage.removeItem(key);
    renderSavedReportsList();
  }
}

function applySnapshot(snap, sourceLabel) {
  allRows = snap.allRows || [];
  savedSections = snap.savedSections || [];
  activeSectionIndex = snap.activeSectionIndex !== undefined ? snap.activeSectionIndex : -1;
  subjectOverrides = snap.subjectOverrides || {};
  facultyMap = snap.facultyMap || { '-1': {} };
  overallWithheldMap = snap.overallWithheldMap || {};
  sigNames = snap.sigNames || { classTeacher: '', coordinator: '', divisionHead: '', hod: '' };
  categoryMap = snap.categoryMap || null;
  revalLog = snap.revalLog || [];

  if (snap.meta) {
    if (snap.meta.inst) document.getElementById('metaInst').value = snap.meta.inst;
    if (snap.meta.dept) document.getElementById('metaDept').value = snap.meta.dept;
    if (snap.meta.title) document.getElementById('metaTitle').value = snap.meta.title;
    if (snap.meta.sem) document.getElementById('metaSem').value = snap.meta.sem;
    if (snap.meta.academicYear) document.getElementById('metaAcademicYear').value = snap.meta.academicYear;
  }

  if (snap.failGrades) document.getElementById('failGradesInput').value = snap.failGrades;
  if (snap.excludeGrades) document.getElementById('excludeGradesInput').value = snap.excludeGrades;

  Object.keys(sigNames).forEach(role => {
    const el = document.querySelector(`.sig-name[data-role="${role}"]`);
    if (el) el.value = sigNames[role];
  });

  rawRows = allRows.slice();
  buildGradeAssignTable();
  buildSummary();

  document.getElementById('emptyHint').style.display = 'none';
  document.getElementById('results').style.display = 'block';
  const landing = document.getElementById('landingSavedReports');
  if (landing) landing.style.display = 'none';

  renderSectionSelect();
  render();
  renderFacultyTab();
  renderReport();
}

function exportSnapshotFile() {
  const snap = snapshotState();
  const json = JSON.stringify(snap, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const fname = (snap.title || 'result_analysis').replace(/[^a-zA-Z0-9_-]/g, '_') + '.rsnap.json';
  triggerDownload(blob, fname, 'application/json');
}

function importSnapshotFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const snap = JSON.parse(e.target.result);
      if (!snap.allRows) throw new Error('Invalid snapshot file format');
      applySnapshot(snap, file.name);
      showToast(`✓ Restored snapshot from ${file.name}`);
    } catch (err) {
      alert('Error loading snapshot: ' + err.message);
    }
  };
  reader.readAsText(file);
}

// --- Sample / Demo Data Loader ---
function loadSampleDataset() {
  const sample = generateSampleResultData();
  allRows = sample.rows;
  rawRows = allRows.slice();

  // Populate faculty mappings
  facultyMap = { '-1': { ...sample.facultyMap } };
  subjectOverrides = {};
  Object.entries(sample.facultyMap).forEach(([code, name]) => {
    subjectOverrides[code] = { staff: name, classStrength: 0, withheld: 0, scope: -1 };
  });

  savedSections = [
    { label: 'A Section', semester: '6', ranges: [{ start: '1BY22IS001', end: '1BY22IS030' }], customUsns: new Set() },
    { label: 'B Section', semester: '6', ranges: [{ start: '1BY22IS031', end: '1BY22IS060' }], customUsns: new Set(['1BY23IS401', '1BY23IS402']) }
  ];
  activeSectionIndex = -1;

  document.getElementById('metaDept').value = 'Department of Information Science & Engineering';
  document.getElementById('metaTitle').value = 'Result Analysis (Even Semester 2025-26)';
  document.getElementById('metaSem').value = 'Semester – 6 (A & B Sections)';

  buildGradeAssignTable();
  buildSummary();

  document.getElementById('emptyHint').style.display = 'none';
  document.getElementById('results').style.display = 'block';
  const landing = document.getElementById('landingSavedReports');
  if (landing) landing.style.display = 'none';

  renderSectionSelect();
  render();
  renderFacultyTab();
  renderReport();

  const meta = document.getElementById('fileMeta');
  if (meta) {
    meta.style.display = 'inline-flex';
    meta.textContent = '✓ Loaded Sample VTU Result Dataset (60 students, 8 subjects with Faculty Allocations).';
  }

  showToast('✓ Loaded Sample VTU Result Sheet with faculty assignments!');
}

// --- Application Initializer & DOM Events ---
document.addEventListener('DOMContentLoaded', () => {
  // Setup Welcome & Masthead Backgrounds
  if (typeof CAMPUS_HERO_DATA_URI !== 'undefined') {
    const wBg = document.getElementById('welcomeBg');
    if (wBg) wBg.style.backgroundImage = `url('${CAMPUS_HERO_DATA_URI}')`;
    const mBg = document.getElementById('mastheadBg');
    if (mBg) mBg.style.backgroundImage = `url('${CAMPUS_HERO_DATA_URI}')`;
  }
  if (typeof BMSIT_LOGO_DATA_URI !== 'undefined') {
    const wLogo = document.getElementById('welcomeLogo');
    if (wLogo) wLogo.src = BMSIT_LOGO_DATA_URI;
    const rLogo = document.getElementById('rLogo');
    if (rLogo) rLogo.src = BMSIT_LOGO_DATA_URI;
  }
  if (typeof WATERMARK_DATA_URI !== 'undefined') {
    const wMark = document.getElementById('reportWatermark');
    if (wMark) wMark.src = WATERMARK_DATA_URI;
  }

  // Welcome Screen Dismiss
  const enterBtn = document.getElementById('enterAppBtn');
  if (enterBtn) {
    enterBtn.addEventListener('click', () => {
      document.getElementById('welcomeOverlay').style.display = 'none';
    });
  }

  // Populate Academic Year Dropdown
  populateAcademicYearOptions();

  // Drag & Drop File Zone
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');

  if (dropZone && fileInput) {
    dropZone.addEventListener('click', e => {
      if (e.target.tagName !== 'BUTTON' && !e.target.closest('button')) {
        fileInput.click();
      }
    });
    fileInput.addEventListener('change', e => {
      if (e.target.files.length) handleFile(e.target.files[0]);
    });
    dropZone.addEventListener('dragover', e => {
      e.preventDefault();
      dropZone.classList.add('drag');
    });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag'));
    dropZone.addEventListener('drop', e => {
      e.preventDefault();
      dropZone.classList.remove('drag');
      if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    });
  }

  // Sample Data Button
  const demoBtn = document.getElementById('loadDemoBtn');
  if (demoBtn) demoBtn.addEventListener('click', loadSampleDataset);
  const welcomeDemoBtn = document.getElementById('welcomeDemoBtn');
  if (welcomeDemoBtn) {
    welcomeDemoBtn.addEventListener('click', () => {
      document.getElementById('welcomeOverlay').style.display = 'none';
      loadSampleDataset();
    });
  }

  // Tab Navigation
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-view').forEach(v => v.classList.remove('active'));
      btn.classList.add('active');
      const target = document.getElementById('tab-' + btn.dataset.tab);
      if (target) target.classList.add('active');

      if (btn.dataset.tab === 'report') renderReport();
      if (btn.dataset.tab === 'faculty') renderFacultyTab();
    });
  });

  // Search input in Subject Table
  const searchInput = document.getElementById('subjectSearchInput');
  if (searchInput) {
    searchInput.addEventListener('input', () => renderTable());
  }

  // Table Sorting on Result Table Headers
  document.querySelectorAll('#resultTable th[data-key]').forEach(th => {
    th.addEventListener('click', () => {
      const key = th.dataset.key;
      if (sortKey === key) sortDir *= -1;
      else { sortKey = key; sortDir = 1; }
      document.querySelectorAll('#resultTable th .arrow').forEach(a => a.textContent = '');
      const arrow = th.querySelector('.arrow');
      if (arrow) arrow.textContent = sortDir === 1 ? '▲' : '▼';
      renderTable();
    });
  });

  // Table Sorting on Faculty Table Headers
  document.querySelectorAll('#facultyTable th[data-key]').forEach(th => {
    th.addEventListener('click', () => {
      const key = th.dataset.key;
      if (facultySortKey === key) facultySortDir *= -1;
      else { facultySortKey = key; facultySortDir = 1; }
      document.querySelectorAll('#facultyTable th .arrow').forEach(a => a.textContent = '');
      const arrow = th.querySelector('.arrow');
      if (arrow) arrow.textContent = facultySortDir === 1 ? '▲' : '▼';
      renderFacultyTab();
    });
  });

  // Student Search input
  const stSearchInput = document.getElementById('studentSearchInput');
  if (stSearchInput) {
    stSearchInput.addEventListener('input', e => searchStudent(e.target.value));
    stSearchInput.addEventListener('focus', e => searchStudent(e.target.value));
  }

  // Section Selector
  const secSelect = document.getElementById('sectionSelect');
  if (secSelect) {
    secSelect.addEventListener('change', e => {
      activeSectionIndex = parseInt(e.target.value, 10);
      applyActiveSection();
      const delBtn = document.getElementById('deleteSectionBtn');
      if (delBtn) delBtn.style.display = activeSectionIndex >= 0 ? 'inline-block' : 'none';
    });
  }

  // Section Save Button
  const saveSecBtn = document.getElementById('saveSectionBtn');
  if (saveSecBtn) {
    saveSecBtn.addEventListener('click', () => {
      const label = (document.getElementById('secLabel').value || '').trim();
      const sem = (document.getElementById('secSemester').value || '').trim();
      const rangesText = document.getElementById('secRanges').value;
      const customText = document.getElementById('secCustom').value;

      if (!label) { alert('Please enter a Section label'); return; }

      const ranges = parseRangesText(rangesText);
      const customUsns = new Set(
        (customText || '').split(/[,\n]/).map(u => u.trim().toUpperCase()).filter(Boolean)
      );

      savedSections.push({ label, semester: sem, ranges, customUsns });
      activeSectionIndex = savedSections.length - 1;
      renderSectionSelect();
      applyActiveSection();
      showToast(`✓ Section "${label}" created successfully.`);
    });
  }

  // Section Delete Button
  const delSecBtn = document.getElementById('deleteSectionBtn');
  if (delSecBtn) {
    delSecBtn.addEventListener('click', () => {
      if (activeSectionIndex >= 0 && confirm('Delete this section view?')) {
        savedSections.splice(activeSectionIndex, 1);
        activeSectionIndex = -1;
        renderSectionSelect();
        applyActiveSection();
      }
    });
  }

  // Combine Sections Button
  const combineBtn = document.getElementById('combineSectionsBtn');
  if (combineBtn) {
    combineBtn.addEventListener('click', () => {
      const cbs = Array.from(document.querySelectorAll('.combine-sec-cb:checked')).map(cb => parseInt(cb.value, 10));
      const label = (document.getElementById('combineLabel').value || '').trim();
      if (cbs.length < 2) { alert('Select at least 2 sections to combine.'); return; }
      if (!label) { alert('Please enter a label for the combined section.'); return; }

      savedSections.push({ label, combined: cbs });
      activeSectionIndex = savedSections.length - 1;
      renderSectionSelect();
      applyActiveSection();
      showToast(`✓ Combined section "${label}" created.`);
    });
  }

  // Revaluation File Upload
  const revalInput = document.getElementById('revalFile');
  if (revalInput) {
    revalInput.addEventListener('change', e => {
      if (!e.target.files.length) return;
      const reader = new FileReader();
      reader.onload = ev => {
        try {
          const data = new Uint8Array(ev.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
          const revalRows = parseResultRows(grid);

          const failSet = new Set(getFailSet());
          const res = applyRevaluation(revalRows, failSet);
          revalLog = res.changes;

          buildSummary();
          render();
          renderFacultyTab();
          renderReport();
          renderRevalLog();

          const revStatus = document.getElementById('revalStatus');
          if (revStatus) {
            revStatus.textContent = `Applied revaluation: ${res.matched} scores updated, ${res.changes.length} grade changes detected.`;
          }
          showToast(`✓ Revaluation applied: ${res.changes.length} grades updated.`);
        } catch (err) {
          alert('Error processing revaluation: ' + err.message);
        }
      };
      reader.readAsArrayBuffer(e.target.files[0]);
    });
  }

  // Mark "After Revaluation" Title button
  const revalTitleBtn = document.getElementById('markAfterRevalBtn');
  if (revalTitleBtn) {
    revalTitleBtn.addEventListener('click', () => {
      document.getElementById('metaTitle').value = 'Result Analysis (After Revaluation)';
      renderReportHeader();
      showToast('✓ Report title set to "After Revaluation".');
    });
  }

  // Signatures Binding
  document.querySelectorAll('.sig-name').forEach(el => {
    el.addEventListener('input', () => { sigNames[el.dataset.role] = el.value; });
  });

  // Report Metadata input events
  ['metaInst', 'metaDept', 'metaTitle', 'metaSem'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', renderReportHeader);
  });
  const aySel = document.getElementById('metaAcademicYear');
  if (aySel) aySel.addEventListener('change', renderReportHeader);

  // Print Button
  const printBtn = document.getElementById('printBtn');
  if (printBtn) printBtn.addEventListener('click', () => window.print());

  // Export Word (.docx) Button
  const exportWordBtn = document.getElementById('exportWordBtn');
  if (exportWordBtn) {
    exportWordBtn.addEventListener('click', async () => {
      const status = document.getElementById('wordExportStatus');
      if (status) status.textContent = 'Generating .docx…';
      try {
        const blob = await buildDocxBlob();
        const dept = document.getElementById('metaDept').value || 'Report';
        const filename = sanitizeFilename(`${dept}_Result_Analysis`) + '.docx';
        const res = await triggerDownload(blob, filename, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        if (status) status.textContent = res === 'picked' ? 'Saved.' : 'Downloaded.';
        showToast('✓ Word report (.docx) generated and saved.');
      } catch (err) {
        alert('Word export error: ' + err.message);
        if (status) status.textContent = 'Export error';
      }
    });
  }

  // CSV Export Buttons
  const downloadCsvBtn = document.getElementById('downloadCsv');
  if (downloadCsvBtn) {
    downloadCsvBtn.addEventListener('click', async () => {
      const header = ['Course Code', 'Course Title', 'Faculty Name', 'Class Strength', 'Appeared', 'Passed', 'Failed', 'Pass %'];
      const rows = [header];
      subjectSummary.slice().sort((a, b) => b.pct - a.pct).forEach(s => {
        rows.push([s.code, s.title, getAssignedFaculty(s.code), s.total, s.appeared, s.passed, s.failed, s.pct + '%']);
      });
      const csv = rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
      const blob = new Blob([csv], { type: 'text/csv' });
      await triggerDownload(blob, 'subject_wise_pass_percentage.csv', 'text/csv');
      showToast('✓ Exported Subject Results CSV.');
    });
  }

  const exportFacultyCsvBtn = document.getElementById('exportFacultyCsv');
  if (exportFacultyCsvBtn) {
    exportFacultyCsvBtn.addEventListener('click', async () => {
      const perfList = calculateFacultyPerformance();
      const header = ['Faculty Name', 'Subjects Handled', 'Appeared', 'Passed', 'Failed', 'Pass %', 'Topper'];
      const rows = [header];
      perfList.forEach(f => {
        rows.push([f.name, f.subjectCodes, f.appeared, f.passed, f.failed, f.pct + '%', f.topper]);
      });
      const csv = rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
      const blob = new Blob([csv], { type: 'text/csv' });
      await triggerDownload(blob, 'faculty_performance_report.csv', 'text/csv');
      showToast('✓ Exported Faculty Performance CSV.');
    });
  }

  // Reset / Load Another File
  const resetBtn = document.getElementById('resetBtn');
  if (resetBtn) {
    resetBtn.addEventListener('click', () => {
      document.getElementById('fileInput').value = '';
      document.getElementById('dropZone').scrollIntoView({ behavior: 'smooth' });
    });
  }

  // Save Report & Snapshots
  const saveRepBtn = document.getElementById('saveReportBtn');
  if (saveRepBtn) saveRepBtn.addEventListener('click', saveReportToStorage);

  const exportSnapBtn = document.getElementById('exportSnapshotBtn');
  if (exportSnapBtn) exportSnapBtn.addEventListener('click', exportSnapshotFile);

  const importSnapInner = document.getElementById('importSnapshotInner');
  if (importSnapInner) {
    importSnapInner.addEventListener('change', e => {
      if (e.target.files.length) importSnapshotFile(e.target.files[0]);
    });
  }
  const importSnapOuter = document.getElementById('importSnapshotOuter');
  if (importSnapOuter) {
    importSnapOuter.addEventListener('change', e => {
      if (e.target.files.length) importSnapshotFile(e.target.files[0]);
    });
  }

  // Section List File Upload Listener
  const secListInput = document.getElementById('sectionListFile');
  if (secListInput) {
    secListInput.addEventListener('change', e => {
      if (e.target.files.length) handleSectionListFile(e.target.files[0]);
    });
  }

  // Section Form Excel Upload Listener (for specific section along with range)
  const secFormExcelInput = document.getElementById('sectionFormExcelFile');
  if (secFormExcelInput) {
    secFormExcelInput.addEventListener('change', e => {
      if (e.target.files.length) handleSectionFormExcelUpload(e.target.files[0]);
    });
  }

  // Category File and Quota Recalculate Listeners
  const catFileInput = document.getElementById('categoryFile');
  if (catFileInput) {
    catFileInput.addEventListener('change', handleCategoryFile);
  }
  const recalcQuotaBtn = document.getElementById('recalcQuotaBtn');
  if (recalcQuotaBtn) {
    recalcQuotaBtn.addEventListener('click', () => {
      renderQuotaTable();
      showToast('✓ Quota groups recalculated.');
    });
  }

  // Render initial saved reports from localStorage
  renderSavedReportsList();
});

function sanitizeFilename(name) {
  return String(name || 'report').replace(/[\\/:*?"<>|]+/g, '_').trim();
}
