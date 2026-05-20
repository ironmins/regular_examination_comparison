/* =========================================================
   2026학년도 내신 성적 통합 분석 대시보드
   - 직전 학기 대비 성장 분석 + 담임 상담 지원 기능
   ========================================================= */

let db2025 = {};
let db2026 = {};
let subjects2026 = [];
let subjects2025 = [];

let charts = {};
let currentOptionList = [];

/* 교과군 분류 매핑 — 과목명 → 교과군 */
const SUBJ_TO_GROUP = {
    '공통국어':'국어군','공통국어1':'국어군','공통국어2':'국어군','문학':'국어군','화법과 작문':'국어군','독서':'국어군','언어와 매체':'국어군',
    '공통수학':'수학군','공통수학1':'수학군','공통수학2':'수학군','대수':'수학군','미적분Ⅰ':'수학군','미적분Ⅱ':'수학군','확률과 통계':'수학군','기하':'수학군',
    '공통영어':'영어군','공통영어1':'영어군','공통영어2':'영어군','영어Ⅰ':'영어군','영어Ⅱ':'영어군','영어회화':'영어군',
    '통합사회':'사회군','통합사회1':'사회군','통합사회2':'사회군','세계시민과 지리':'사회군','세계사':'사회군','사회와 문화':'사회군','정치':'사회군','현대사회와 윤리':'사회군','한국사':'사회군','한국사1':'사회군','한국사2':'사회군','동아시아사':'사회군','경제':'사회군','윤리와 사상':'사회군','한국지리':'사회군','세계지리':'사회군','생활과 윤리':'사회군',
    '통합과학':'과학군','통합과학1':'과학군','통합과학2':'과학군','물리학':'과학군','화학':'과학군','생명과학':'과학군','지구과학':'과학군','물리학Ⅱ':'과학군','화학Ⅱ':'과학군','생명과학Ⅱ':'과학군','지구과학Ⅱ':'과학군',
    '중국어Ⅰ':'제2외국어군','일본어Ⅰ':'제2외국어군','중국어Ⅱ':'제2외국어군','일본어Ⅱ':'제2외국어군'
};

function getGroup(subj) {
    if(SUBJ_TO_GROUP[subj]) return SUBJ_TO_GROUP[subj];
    // 유연 매칭: 공백 제거 후 비교
    let clean = subj.replace(/ /g, '');
    for(let k in SUBJ_TO_GROUP) {
        if(k.replace(/ /g, '') === clean) return SUBJ_TO_GROUP[k];
    }
    return '기타';
}

/* DOMAINS 동적 생성 — 업로드된 두 파일의 실제 과목명 기반 */
let DOMAINS = [];

function buildDomains() {
    DOMAINS = [];

    // 직전/이번 DB에서 실제 과목명 수집
    let prevSubjects = new Set();
    let currSubjects = new Set();
    Object.values(db2025).forEach(st => {
        Object.keys(st.scores).forEach(s => { if(st.scores[s].score !== null) prevSubjects.add(s); });
    });
    Object.values(db2026).forEach(st => {
        Object.keys(st.scores).forEach(s => { if(st.scores[s].score !== null) currSubjects.add(s); });
    });

    subjects2025 = Array.from(prevSubjects);
    subjects2026 = Array.from(currSubjects);

    // 교과군별로 그룹핑
    let prevByGroup = {};
    let currByGroup = {};
    subjects2025.forEach(s => {
        let g = getGroup(s);
        if(!prevByGroup[g]) prevByGroup[g] = [];
        prevByGroup[g].push(s);
    });
    subjects2026.forEach(s => {
        let g = getGroup(s);
        if(!currByGroup[g]) currByGroup[g] = [];
        currByGroup[g].push(s);
    });

    // 모든 교과군 합집합
    let allGroups = new Set([...Object.keys(prevByGroup), ...Object.keys(currByGroup)]);

    allGroups.forEach(g => {
        let prev = prevByGroup[g] || [];
        let curr = currByGroup[g] || [];

        // 같은 과목명이 양쪽에 모두 있으면 reliable (1:1 매칭 가능)
        let commonSubjects = prev.filter(s => curr.some(c => matchSubjectName(s, c)));

        let reliable;
        if(commonSubjects.length > 0 && commonSubjects.length === prev.length) {
            // 직전 과목이 모두 이번에도 있음 → 완전 1:1 매칭
            reliable = true;
        } else if(prev.length === 0 || curr.length === 0) {
            // 한쪽만 있음 → 비교 불가지만 표시는 함
            reliable = true;
        } else if(prev.length <= 2 && curr.length <= 2 && commonSubjects.length === 0) {
            // 통합→선택 전환 (공통국어→문학 등)
            reliable = prev.length === 1;
        } else {
            reliable = false;
        }

        DOMAINS.push({
            name: g,
            prev: prev,
            curr: curr,
            reliable: reliable
        });
    });

    // 교과군 순서: 국어→수학→영어→사회→과학→제2외국어→기타
    let order = ['국어군','수학군','영어군','사회군','과학군','제2외국어군','기타'];
    DOMAINS.sort((a,b) => {
        let ia = order.indexOf(a.name); if(ia === -1) ia = 99;
        let ib = order.indexOf(b.name); if(ib === -1) ib = 99;
        return ia - ib;
    });
}

/* 과목명 유연 비교 (공백/특수문자 무시) */
function matchSubjectName(a, b) {
    if(a === b) return true;
    let ca = a.replace(/ /g, '');
    let cb = b.replace(/ /g, '');
    return ca === cb || ca.includes(cb) || cb.includes(ca);
}

/* =========================================================
   엑셀 읽기 / DB 구성  (기존 로직 유지)
   ========================================================= */
function readExcelAsync(file) {
    return new Promise((resolve, reject) => {
        let reader = new FileReader();
        reader.onload = (e) => {
            try {
                let data = new Uint8Array(e.target.result);
                let workbook = XLSX.read(data, {type: 'array'});

                let targetSheetName = workbook.SheetNames.find(name => {
                    let sheet = workbook.Sheets[name];
                    let json = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ""});
                    return json.some(row => {
                        let str = row.map(c => String(c)).join("").replace(/ /g, "");
                        return str.includes("학번") && str.includes("이름");
                    });
                });

                let sheetName = targetSheetName || workbook.SheetNames[0];
                let firstSheet = workbook.Sheets[sheetName];
                let jsonArray = XLSX.utils.sheet_to_json(firstSheet, {header: 1, defval: ""});
                let stringArray = jsonArray.map(row =>
                    row.map(cell => (cell !== null && cell !== undefined) ? String(cell).trim() : "")
                );
                resolve(stringArray);
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function buildDB(excelData) {
    let db = {};
    if(!excelData || excelData.length < 3) return db;

    let fieldRowIdx = -1;
    for(let i=0; i<excelData.length; i++) {
        let rowStr = excelData[i].join("").replace(/ /g, "");
        if(rowStr.includes("학번") && rowStr.includes("이름")) {
            fieldRowIdx = i;
            break;
        }
    }
    if(fieldRowIdx === -1) return db;

    let subjRowIdx = fieldRowIdx - 1;
    let colMap = {};
    if (subjRowIdx >= 0) {
        let subRow = excelData[subjRowIdx];
        for(let i=0; i<subRow.length; i++) {
            let val = subRow[i];
            if(val && val.length >= 1) colMap[val] = i;
        }
    }

    let fieldRow = excelData[fieldRowIdx];
    let hakbunIdx=-1, nameIdx=-1, clsIdx=-1, numIdx=-1;
    for(let i=0; i<fieldRow.length; i++) {
        let val = fieldRow[i].replace(/ /g, "");
        if(val.includes("학번")) hakbunIdx = i;
        else if(val.includes("이름")) nameIdx = i;
        else if(val === "반") clsIdx = i;
        else if(val === "번호") numIdx = i;
    }

    let startRow = fieldRowIdx + 1;

    for(let r=startRow; r<excelData.length; r++) {
        const row = excelData[r];
        if(row.length <= Math.max(hakbunIdx, nameIdx)) continue;

        const hakbun = row[hakbunIdx];
        const name = row[nameIdx];

        if(!hakbun || !name || hakbun === "학번" || hakbun === "신학번" || hakbun.includes("통계") || hakbun.includes("평균")) continue;

        let student = { hakbun, name, cls: row[clsIdx], num: row[numIdx], scores: {} };

        for(const [subj, idx] of Object.entries(colMap)) {
            if(row.length > idx+3) {
                let rawScore = row[idx];
                let scoreVal = parseFloat(rawScore);
                // 빈 셀(미응시)은 null 로 처리하여 통계에서 제외
                student.scores[subj] = {
                    score: (rawScore === "" || isNaN(scoreVal)) ? null : scoreVal,
                    rank: row[idx+1],
                    grade: parseFloat(row[idx+2]) || null,
                    pct: parseFloat(row[idx+3]) || null
                };
            }
        }
        db[hakbun] = student;
    }
    return db;
}

async function processData() {
    const f2025 = document.getElementById('file2025').files[0];
    const f2026 = document.getElementById('file2026').files[0];

    if(!f2025 || !f2026) {
        alert("2025학년도 파일과 2026학년도 파일을 모두 업로드해주세요.");
        return;
    }

    document.getElementById('loader').style.display = 'block';

    try {
        const data25 = await readExcelAsync(f2025);
        const data26 = await readExcelAsync(f2026);

        db2025 = buildDB(data25);
        db2026 = buildDB(data26);

        if(Object.keys(db2026).length === 0) {
            alert("2026학년도 데이터에서 학생 정보를 찾지 못했습니다.");
        }

        buildDomains();
        populateDropdowns();
        buildGlobalStats();

        // 업로드된 파일명에서 시험 정보 추출
        updateExamLabel();

        document.getElementById('loader').style.display = 'none';
        document.getElementById('workspace').style.display = 'block';
        document.getElementById('btnSaveHtml').disabled = false;

    } catch (e) {
        alert("엑셀 파일 처리 중 오류가 발생했습니다: " + e.message);
        document.getElementById('loader').style.display = 'none';
    }
}

/* =========================================================
   학생 선택 드롭다운 / 학급 필터
   ========================================================= */
function populateDropdowns() {
    filterStudents();
}

function filterStudents() {
    const clsFilter = document.getElementById('classSelect').value;
    const sel = document.getElementById('studentSelect');
    sel.innerHTML = '<option value="">학생을 선택하세요</option>';
    currentOptionList = [];

    const students = Object.values(db2026).sort((a,b) => a.hakbun.localeCompare(b.hakbun));
    students.forEach(st => {
        if(clsFilter === 'all' || String(st.cls) === String(clsFilter)) {
            let opt = document.createElement('option');
            opt.value = st.hakbun;
            opt.textContent = `${st.hakbun} ${st.name} (${st.cls}반)`;
            sel.appendChild(opt);
            currentOptionList.push(st.hakbun);
        }
    });

    if(sel.options.length > 1) {
        sel.selectedIndex = 1;
        updateDashboard();
    } else {
        sel.selectedIndex = 0;
    }
}

/* 이전 / 다음 학생 이동 */
function navStudent(dir) {
    const sel = document.getElementById('studentSelect');
    const cur = sel.value;
    let idx = currentOptionList.indexOf(cur);
    if(idx === -1) idx = 0;
    let next = idx + dir;
    if(next < 0 || next >= currentOptionList.length) return;
    sel.value = currentOptionList[next];
    updateDashboard();
}

function updateNavButtons() {
    const sel = document.getElementById('studentSelect');
    const idx = currentOptionList.indexOf(sel.value);
    document.getElementById('prevBtn').disabled = (idx <= 0);
    document.getElementById('nextBtn').disabled = (idx === -1 || idx >= currentOptionList.length - 1);
}

/* =========================================================
   학생 이름 검색창
   ========================================================= */
let searchHighlightIdx = -1;

function onSearchInput() {
    const q = document.getElementById('studentSearch').value.trim();
    const box = document.getElementById('searchResults');
    searchHighlightIdx = -1;

    if(!q) { box.style.display = 'none'; return; }

    const matches = Object.values(db2026)
        .filter(st => st.name.includes(q) || st.hakbun.includes(q))
        .sort((a,b) => a.hakbun.localeCompare(b.hakbun))
        .slice(0, 30);

    if(matches.length === 0) {
        box.innerHTML = '<div style="color:#999;cursor:default;">검색 결과가 없습니다</div>';
        box.style.display = 'block';
        return;
    }

    box.innerHTML = matches.map(st =>
        `<div onmousedown="pickSearch('${st.hakbun}')">${st.name} <span style="color:#888;">${st.hakbun} · ${st.cls}반 ${st.num}번</span></div>`
    ).join('');
    box.style.display = 'block';
}

function pickSearch(hakbun) {
    const st = db2026[hakbun];
    if(!st) return;
    // 학급 필터가 다른 반이면 전체 학급으로 전환
    const clsSel = document.getElementById('classSelect');
    if(clsSel.value !== 'all' && String(clsSel.value) !== String(st.cls)) {
        clsSel.value = 'all';
        filterStudents();
    }
    const sel = document.getElementById('studentSelect');
    sel.value = hakbun;
    document.getElementById('studentSearch').value = '';
    document.getElementById('searchResults').style.display = 'none';
    updateDashboard();
}

function onSearchKey(e) {
    const box = document.getElementById('searchResults');
    const items = box.querySelectorAll('div');
    if(box.style.display === 'none' || items.length === 0) return;

    if(e.key === 'ArrowDown') {
        e.preventDefault();
        searchHighlightIdx = Math.min(searchHighlightIdx + 1, items.length - 1);
    } else if(e.key === 'ArrowUp') {
        e.preventDefault();
        searchHighlightIdx = Math.max(searchHighlightIdx - 1, 0);
    } else if(e.key === 'Enter') {
        e.preventDefault();
        if(searchHighlightIdx >= 0 && items[searchHighlightIdx]) {
            items[searchHighlightIdx].dispatchEvent(new MouseEvent('mousedown'));
        }
        return;
    } else if(e.key === 'Escape') {
        box.style.display = 'none';
        return;
    } else {
        return;
    }
    items.forEach((it, i) => it.classList.toggle('highlight', i === searchHighlightIdx));
}

document.addEventListener('click', (e) => {
    if(!e.target.closest('.search-wrap')) {
        const box = document.getElementById('searchResults');
        if(box) box.style.display = 'none';
    }
});

/* =========================================================
   탭 전환
   ========================================================= */
function switchTab(evt, tabId) {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    if(evt && evt.target) evt.target.classList.add('active');
    document.getElementById(tabId).classList.add('active');
}

/* =========================================================
   대시보드 갱신
   ========================================================= */
function updateDashboard() {
    const hakbun = document.getElementById('studentSelect').value;
    if(!hakbun || !db2026[hakbun]) return;

    const s26 = db2026[hakbun];
    let s25 = db2025[hakbun];
    if(!s25) {
        s25 = Object.values(db2025).find(st => st.name === s26.name);
    }

    renderTab1(s26);
    renderTab2(s25, s26);
    renderTab3(s26);
    updateNavButtons();
}

/* 과목명 유연 매칭 */
function matchSubject(studentScores, targetSubj) {
    let cleanTarget = targetSubj.replace(/ /g, '');
    if(studentScores[targetSubj]) return targetSubj;
    let found = Object.keys(studentScores).find(k => {
        let cleanK = k.replace(/ /g, '');
        return cleanK.includes(cleanTarget) || cleanTarget.includes(cleanK) || (cleanTarget === '사회와문화' && cleanK === '사회와문와');
    });
    return found || null;
}

/* 등급 배지 HTML */
function gradeBadge(grade) {
    if(!grade) return '<span class="grade-badge grade-none">-</span>';
    let g = Math.round(grade);
    if(g < 1) g = 1; if(g > 5) g = 5;
    return `<span class="grade-badge grade-${g}">${grade}등급</span>`;
}

/* 석차백분율 게이지 막대 HTML  (백분율이 낮을수록 = 상위권 = 막대 길고 초록) */
function pctGauge(pct) {
    if(pct === null || pct === undefined || isNaN(pct)) {
        return '<span style="color:#999;">-</span>';
    }
    let standing = 100 - pct;         // 상위 standing%
    let color = '#15803D';
    if(pct > 23) color = '#4D7C0F';
    if(pct > 40) color = '#D97706';
    if(pct > 60) color = '#EA580C';
    if(pct > 77) color = '#DC2626';
    return `<div class="pct-gauge">
        <div class="bar-track"><div class="bar-fill" style="width:${standing.toFixed(0)}%;background:${color};"></div></div>
        <span class="bar-label">${pct}%</span>
    </div>`;
}

/* 성취도 배지 (원점수 기준)  A 90↑ / B 80↑ / C 70↑ / D 60↑ / E 60미만 */
function achvBadge(score) {
    if(score === null || score === undefined || isNaN(score)) {
        return '<span class="achv-badge achv-none">-</span>';
    }
    let g = 'E';
    if(score >= 90) g = 'A';
    else if(score >= 80) g = 'B';
    else if(score >= 70) g = 'C';
    else if(score >= 60) g = 'D';
    return `<span class="achv-badge achv-${g}">${g}</span>`;
}

/* 석차백분율(pct) → 백분위 환산  (백분위 = 100 - 석차백분율, 클수록 우수) */
function toPercentile(pct) {
    if(pct === null || pct === undefined || isNaN(pct)) return null;
    return 100 - pct;
}

/* 학생 1명의 전과목 평균 백분위 */
function avgPercentile(student) {
    let sum = 0, cnt = 0;
    Object.values(student.scores).forEach(sc => {
        if(sc.score !== null && sc.pct !== null && sc.pct !== undefined && !isNaN(sc.pct)) {
            sum += (100 - sc.pct); cnt++;
        }
    });
    return cnt > 0 ? sum / cnt : null;
}

/* 등급 순위 계산 — 평균등급이 낮을수록(우수) 상위.
   scope: 'class' 면 같은 반, 'total' 이면 전체.
   반환: { rank, total } 또는 null */
function gradeRank(s26, scope) {
    let pool = Object.values(db2026);
    if(scope === 'class') pool = pool.filter(st => String(st.cls) === String(s26.cls));

    // 평균등급이 산출되는 학생만 대상
    let ranked = pool
        .map(st => ({ hakbun: st.hakbun, avg: avgGrade(st) }))
        .filter(x => x.avg !== null)
        .sort((a, b) => a.avg - b.avg);

    let myAvg = avgGrade(s26);
    if(myAvg === null || ranked.length === 0) return null;

    // 동점은 같은 등수(공동 순위)로 처리
    let rank = 1 + ranked.filter(x => x.avg < myAvg).length;
    return { rank, total: ranked.length };
}

/* =========================================================
   탭1 : 개인별 성적 종합 조회  (헤더 + 학생정보 + 레이더 + 응시전과목 표)
   ========================================================= */
function renderTab1(s26) {
    // ----- 응시 과목 수집 (subjects2026 순서 유지) -----
    let rows = [];
    subjects2026.forEach(subj => {
        let key = matchSubject(s26.scores, subj);
        if(key && s26.scores[key] && s26.scores[key].score !== null) {
            let sc = s26.scores[key];
            rows.push({ subj, score: sc.score, grade: sc.grade, rank: sc.rank, pct: sc.pct });
        }
    });

    // ----- 헤더 카드 -----
    const avgG = avgGrade(s26);
    document.getElementById('shName').textContent = s26.name;
    document.getElementById('shClass').textContent =
        `2학년 ${s26.cls}반 ${s26.num}번`;
    document.getElementById('shAvgGrade').textContent = avgG !== null ? avgG.toFixed(1) : '-';
    document.getElementById('shTotal').textContent = Object.keys(db2026).length + '명';

    // ----- 주의 학생 플래그 -----
    const flag = computeFlag(s26);
    document.getElementById('flagBadge1').innerHTML = flagBadgeHtml(flag);

    // ----- 학생 정보 패널 -----
    const avgP = avgPercentile(s26);
    const cRank = gradeRank(s26, 'class');
    const tRank = gradeRank(s26, 'total');
    document.getElementById('infoAvgGrade').textContent = avgG !== null ? avgG.toFixed(2) + ' 등급' : '-';
    document.getElementById('infoAvgPct').textContent = avgP !== null ? avgP.toFixed(1) : '-';
    document.getElementById('infoClassRank').textContent = cRank ? `${cRank.rank} / ${cRank.total} 위` : '-';
    document.getElementById('infoTotalRank').textContent = tRank ? `${tRank.rank} / ${tRank.total} 위` : '-';
    document.getElementById('infoSubjCount').textContent = rows.length + ' 과목';

    // 직전 학기 대비 추세 (교과군 향상/하락 다수 판정)
    let s25 = db2025[s26.hakbun] || Object.values(db2025).find(st => st.name === s26.name);
    const trendEl = document.getElementById('infoTrend');
    if(s25) {
        let up = 0, down = 0;
        DOMAINS.forEach(d => {
            let vp = avgMetric(s25, d.prev, 'percentile');
            let vc = avgMetric(s26, d.curr, 'percentile');
            if(vp !== null && vc !== null && Math.abs(vc - vp) > 1) {
                if(vc < vp) up++; else down++;
            }
        });
        if(up > down)      trendEl.innerHTML = '<span class="delta-up">전반적 향상</span>';
        else if(down > up) trendEl.innerHTML = '<span class="delta-down">일부 하락</span>';
        else               trendEl.innerHTML = '<span class="delta-flat">유지</span>';
    } else {
        trendEl.innerHTML = '<span style="color:#999;">직전 데이터 없음</span>';
    }

    // ----- 과목별 상세 분석 표 (응시 전과목 단일 표) -----
    document.getElementById('subjCountLabel').textContent = `응시 ${rows.length}과목`;
    const tb = document.querySelector('#reportCardTable tbody');
    tb.innerHTML = '';
    rows.forEach(r => {
        let percentile = toPercentile(r.pct);
        let pctColor = percentile === null ? '#999'
                     : (percentile >= 77 ? 'var(--success, #15803D)'
                     : (percentile >= 40 ? 'var(--warning, #D97706)' : '#DC2626'));
        tb.innerHTML += `<tr>
            <td style="font-weight:bold;">${r.subj}</td>
            <td>${r.score}</td>
            <td>${achvBadge(r.score)}</td>
            <td>${gradeBadge(r.grade)}</td>
            <td>${r.rank || '-'}</td>
            <td style="color:${pctColor}; font-weight:bold;">${percentile !== null ? percentile.toFixed(1) : '-'}</td>
        </tr>`;
    });
    if(rows.length === 0) {
        tb.innerHTML = '<tr><td colspan="6" style="color:#999;">응시 과목 데이터가 없습니다.</td></tr>';
    }

    // ----- 레이더 차트 (응시 전과목 · 백분위 기준) -----
    let radarLabels = rows.map(r => r.subj);
    let radarData = rows.map(r => { let p = toPercentile(r.pct); return p !== null ? p : 0; });

    if(charts.radar) charts.radar.destroy();
    const ctx = document.getElementById('radarChart').getContext('2d');
    charts.radar = new Chart(ctx, {
        type: 'radar',
        data: {
            labels: radarLabels,
            datasets: [{
                label: '백분위 (바깥쪽일수록 상위권)',
                data: radarData,
                backgroundColor: 'rgba(95, 74, 139, 0.2)',
                borderColor: 'rgba(95, 74, 139, 1)',
                borderWidth: 2,
                pointBackgroundColor: 'rgba(95, 74, 139, 1)'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            aspectRatio: 1,
            layout: { padding: { top: 10, bottom: 10, left: 10, right: 10 } },
            scales: {
                r: {
                    min: 0, max: 100,
                    ticks: { stepSize: 25, display: true, backdropColor: 'transparent', color: '#aaa', font: { size: 9 } },
                    pointLabels: {
                        font: { size: 11, family: "'Pretendard', sans-serif" },
                        padding: 14,
                        callback: function(label) {
                            if(label.length > 5) return label.substring(0, 5) + '..';
                            return label;
                        }
                    }
                }
            },
            plugins: { legend: { display: false } }
        }
    });
}

/* 주의 학생 자동 판정 — 담임 상담 우선순위 3단계 분류
   level: 'high'  최우선 상담  / 'mid' 관심 필요 / 'ok' 안정적
   판정 점수가 누적될수록 단계가 올라감 */
function computeFlag(s26) {
    let reasons = [];
    let severity = 0;   // 누적 심각도 점수

    // 1) 응시 과목 대비 하위 등급(4·5등급) 비율
    let lowGrades = 0, graded = 0;
    subjects2026.forEach(subj => {
        let k = matchSubject(s26.scores, subj);
        if(k && s26.scores[k] && s26.scores[k].score !== null && s26.scores[k].grade) {
            graded++;
            if(s26.scores[k].grade >= 4) lowGrades++;
        }
    });
    if(graded >= 3) {
        let ratio = lowGrades / graded;
        if(ratio >= 0.7) { reasons.push(`응시 ${graded}과목 중 ${lowGrades}과목이 4·5등급`); severity += 2; }
        else if(ratio >= 0.5) { reasons.push(`4·5등급 과목 다수(${lowGrades}/${graded})`); severity += 1; }
    }

    // 2) 직전 대비 하락 교과군 (석차백분율 상승 = 석차 하락)
    let s25 = db2025[s26.hakbun] || Object.values(db2025).find(st => st.name === s26.name);
    if(s25) {
        let downCnt = 0, maxDrop = 0, maxDropName = '';
        DOMAINS.forEach(d => {
            let vPrev = avgMetric(s25, d.prev, 'percentile');
            let vCurr = avgMetric(s26, d.curr, 'percentile');
            if(vPrev !== null && vCurr !== null) {
                let drop = vCurr - vPrev;
                if(drop > 10) downCnt++;
                if(drop > maxDrop) { maxDrop = drop; maxDropName = d.name; }
            }
        });
        if(downCnt >= 3) { reasons.push(`${downCnt}개 교과군 동반 하락`); severity += 2; }
        else if(downCnt >= 2) { reasons.push(`${downCnt}개 교과군 석차 하락`); severity += 1; }
        if(maxDrop > 25) { reasons.push(`${maxDropName} 급락(${maxDrop.toFixed(0)}%p)`); severity += 1; }
    }

    // 3) 전과목 평균 등급
    let avgG = avgGrade(s26);
    if(avgG !== null) {
        if(avgG >= 4.3) { reasons.push(`평균 ${avgG.toFixed(1)}등급`); severity += 2; }
        else if(avgG >= 3.7) { reasons.push(`평균 ${avgG.toFixed(1)}등급`); severity += 1; }
    }

    let level = 'ok';
    if(severity >= 3) level = 'high';
    else if(severity >= 1) level = 'mid';

    return { level, alert: level !== 'ok', reasons };
}

/* 플래그 단계 → 화면 배지 HTML */
function flagBadgeHtml(flag) {
    if(flag.level === 'high') return `<span class="flag-badge high">🔔 최우선 상담</span>`;
    if(flag.level === 'mid')  return `<span class="flag-badge">⚠️ 관심 필요</span>`;
    return `<span class="flag-badge ok">✓ 안정적</span>`;
}

/* 특정 과목 묶음의 평균 (지표: percentile / score) */
function avgMetric(student, subjList, metric) {
    let sum = 0, valid = 0;
    subjList.forEach(sj => {
        let k = matchSubject(student.scores, sj);
        if(k && student.scores[k]) {
            let v = metric === 'percentile' ? student.scores[k].pct : student.scores[k].score;
            if(v !== null && v !== undefined && !isNaN(v)) { sum += v; valid++; }
        }
    });
    return valid > 0 ? sum / valid : null;
}

/* 전과목 평균 등급 */
function avgGrade(student) {
    let sum = 0, cnt = 0;
    Object.values(student.scores).forEach(sc => {
        if(sc.score !== null && sc.grade) { sum += sc.grade; cnt++; }
    });
    return cnt > 0 ? sum / cnt : null;
}

/* 응시 과목 수 */
function countSubjects(student) {
    return Object.values(student.scores).filter(sc => sc.score !== null).length;
}

/* 강점 / 보완 과목 Top3  (석차백분율 기준) */
function strengthsWeaknesses(s26) {
    let arr = [];
    subjects2026.forEach(subj => {
        let k = matchSubject(s26.scores, subj);
        if(k && s26.scores[k] && s26.scores[k].score !== null) {
            let sc = s26.scores[k];
            arr.push({ subj, score: sc.score, grade: sc.grade, pct: sc.pct });
        }
    });
    let withPct = arr.filter(a => a.pct !== null && a.pct !== undefined);
    let sorted = [...withPct].sort((a,b) => a.pct - b.pct);
    return {
        strong: sorted.slice(0, 3),
        weak: [...sorted].reverse().slice(0, 3)
    };
}

/* =========================================================
   탭2 : 직전 대비 향상/하락 분석  (교과군 → 과목 아코디언)
   ========================================================= */
function renderTab2(s25, s26) {
    const metric = document.getElementById('metricToggle').value;
    const accordion = document.getElementById('domainAccordion');
    accordion.innerHTML = '';
    const narrativeBox = document.getElementById('narrativeBox');

    // ---------- 0) 상단 학생 정보 카드 채우기 (s25 유무와 무관) ----------
    fillTab2InfoCard(s25, s26);

    if(!s25) {
        narrativeBox.innerHTML = "해당 학생의 2025학년도(직전 학기) 데이터를 찾을 수 없어 향상도 분석이 불가합니다. (동명이인이 아닌 전학생일 가능성)";
        if(charts.bar) charts.bar.destroy();
        return;
    }

    // ---------- 1) 차트 + 종합 narrative 데이터 ----------
    let barLabels = [];
    let barDeltas = [];
    let upDomains = [], downDomains = [];
    let hasUnreliable = false;

    DOMAINS.forEach(d => {
        let vPrev = avgMetric(s25, d.prev, metric);
        let vCurr = avgMetric(s26, d.curr, metric);
        if(vPrev === null || vCurr === null) return;

        let delta = vCurr - vPrev;
        // percentile(석차백분율)은 작을수록 우수 → 감소가 향상
        let isPositive = metric === 'percentile' ? (delta < 0) : (delta > 0);
        if(Math.abs(delta) > 1.0) {
            if(isPositive) upDomains.push(d.name);
            else downDomains.push(d.name);
            barDeltas.push(metric === 'percentile' ? -delta : delta);
        } else {
            barDeltas.push(0);
        }
        barLabels.push(d.name);
        if(!d.reliable) hasUnreliable = true;
    });

    // narrative
    let ntext = `<strong>종합 분석:</strong> ${s26.name} 학생은 직전 학기 대비 `;
    if(upDomains.length > downDomains.length) {
        ntext += `<span class="delta-up">전반적으로 향상</span>되는 추세입니다. `;
        if(upDomains.length > 0) ntext += `특히 <strong>${upDomains.join(', ')}</strong>에서 눈에 띄는 성장이 있었습니다.`;
    } else if (downDomains.length > upDomains.length) {
        ntext += `<span class="delta-down">일부 교과에서 어려움</span>을 겪고 있는 것으로 진단됩니다. `;
        if(downDomains.length > 0) ntext += `<strong>${downDomains.join(', ')}</strong>의 학습법 점검이 필요해 보입니다.`;
    } else {
        ntext += `전체적으로 <strong>비슷한 성취도</strong>를 유지하고 있습니다.`;
    }
    if(hasUnreliable) {
        ntext += ` <span style="font-size:0.82rem;color:#D97706;">※ 사·과군은 응시 집단이 달라 참고용</span>`;
    }
    narrativeBox.innerHTML = ntext;

    // ---------- 2) 막대 차트 ----------
    if(charts.bar) charts.bar.destroy();
    const ctxBar = document.getElementById('barChart').getContext('2d');
    charts.bar = new Chart(ctxBar, {
        type: 'bar',
        data: {
            labels: barLabels,
            datasets: [{
                label: metric === 'percentile' ? '백분위 변화 (+가 향상, %p)' : '원점수 변화 (점)',
                data: barDeltas,
                backgroundColor: barDeltas.map(v =>
                    v > 0 ? 'rgba(21, 128, 61, 0.75)'
                          : (v < 0 ? 'rgba(217, 119, 6, 0.75)' : 'rgba(160, 160, 160, 0.4)'))
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: { y: { beginAtZero: true } },
            plugins: { legend: { display: false } }
        }
    });

    // ---------- 3) 교과군 아코디언 ----------
    DOMAINS.forEach((d, idx) => {
        let vPrevPct = avgMetric(s25, d.prev, 'percentile');  // 석차백분율
        let vCurrPct = avgMetric(s26, d.curr, 'percentile');
        if(vPrevPct === null && vCurrPct === null) return;     // 양쪽 다 미응시 → 표시 안함

        // 백분위(=100-석차백분율) 환산
        let prevPercentile = vPrevPct !== null ? (100 - vPrevPct) : null;
        let currPercentile = vCurrPct !== null ? (100 - vCurrPct) : null;
        let pctDelta = (prevPercentile !== null && currPercentile !== null)
            ? (currPercentile - prevPercentile) : null;

        // 군별 향상/하락 배지
        let badgeHtml = '';
        if(pctDelta === null) {
            badgeHtml = `<span class="domain-badge flat">비교 불가</span>`;
        } else if(Math.abs(pctDelta) <= 1.0) {
            badgeHtml = `<span class="domain-badge flat">─ 유지</span>`;
        } else if(pctDelta > 0) {
            badgeHtml = `<span class="domain-badge up">▲ ${pctDelta.toFixed(1)} 향상</span>`;
        } else {
            badgeHtml = `<span class="domain-badge down">▼ ${Math.abs(pctDelta).toFixed(1)} 하락</span>`;
        }

        // 군 헤더 요약 ("백분위 83.0 → 98.6")
        let summaryHtml = '';
        if(prevPercentile !== null && currPercentile !== null) {
            summaryHtml = `<span class="pct-change">백분위 <b>${prevPercentile.toFixed(1)}</b> → <b>${currPercentile.toFixed(1)}</b></span>`;
        } else if(currPercentile !== null) {
            summaryHtml = `<span class="pct-change">이번 백분위 <b>${currPercentile.toFixed(1)}</b> <small>(직전 미응시)</small></span>`;
        } else if(prevPercentile !== null) {
            summaryHtml = `<span class="pct-change">이번 학기 미응시 <small>(직전 ${prevPercentile.toFixed(1)})</small></span>`;
        }

        let refTag = d.reliable ? '' : '<span class="ref-tag">참고</span>';

        // 본문: 신뢰 가능 → 1:1 표 / 불가 → 좌우 카드
        let bodyHtml = d.reliable
            ? buildReliableBody(s25, s26, d)
            : buildUnreliableBody(s25, s26, d);

        // (사용자 답변에 따라) 처음부터 모두 펼쳐서 보여주기
        accordion.insertAdjacentHTML('beforeend', `
            <div class="domain-card" data-domain-idx="${idx}">
                <div class="domain-header" onclick="toggleDomainCard(this)">
                    <span class="domain-toggle">▼</span>
                    <span class="domain-name">${d.name}${refTag}</span>
                    <span class="domain-summary">${summaryHtml}${badgeHtml}</span>
                </div>
                <div class="domain-body">${bodyHtml}</div>
            </div>
        `);
    });

    if(!accordion.innerHTML) {
        accordion.innerHTML = '<p style="color:#999; padding:20px; text-align:center;">표시할 교과군 데이터가 없습니다.</p>';
    }
}

/* 신뢰 가능한 교과군 (국·수·영) — 1:1 비교 표
   각 curr 과목을 동일한 prev 과목과 1:1 매칭하여 한 행으로 표시 */
function buildReliableBody(s25, s26, d) {
    let rowsHtml = `
        <div class="subj-row head">
            <div>과목 (직전 → 이번)</div>
            <div>원점수</div>
            <div>석차 · 등급</div>
            <div>백분위</div>
        </div>
    `;

    let drawnCount = 0;

    d.curr.forEach(currSubj => {
        let currKey = matchSubject(s26.scores, currSubj);
        let currSc = (currKey && s26.scores[currKey] && s26.scores[currKey].score !== null)
            ? s26.scores[currKey] : null;
        if(!currSc) return;

        // 직전 과목 매칭: 같은 이름이 있으면 1:1, 없으면 교과군 대표(prev[0])
        let prevSubj = null;
        let prevSc = null;

        // 1) 동일 과목명 매칭 시도
        let sameKey = matchSubject(s25.scores, currSubj);
        if(sameKey && s25.scores[sameKey] && s25.scores[sameKey].score !== null) {
            prevSubj = currSubj;
            prevSc = s25.scores[sameKey];
        }

        // 2) 동일 과목 없으면 교과군 prev 과목에서 찾기
        if(!prevSc) {
            for(let i = 0; i < d.prev.length; i++) {
                let pk = matchSubject(s25.scores, d.prev[i]);
                if(pk && s25.scores[pk] && s25.scores[pk].score !== null) {
                    prevSubj = d.prev[i];
                    prevSc = s25.scores[pk];
                    break;
                }
            }
        }

        if(!prevSc) {
            // 직전 데이터 없음 — 이번 과목만 표시
            rowsHtml += `<div class="subj-row">
                <div><span class="subj-name">직전 미응시 <span class="arrow">→</span> ${currSubj}</span></div>
                <div><span class="metric-main">- → ${currSc.score}</span></div>
                <div><span class="rank-line">-</span></div>
                <div><span class="metric-main">${currSc.pct !== null ? (100 - currSc.pct).toFixed(1) : '-'}</span></div>
            </div>`;
            drawnCount++;
            return;
        }

        // 같은 과목이면 "문학" 한 번만, 다르면 "공통국어 → 문학"
        let displayPrev = prevSubj;
        let displayCurr = currSubj;
        if(matchSubjectName(prevSubj, currSubj)) {
            displayPrev = currSubj; // 같은 이름
        }

        rowsHtml += renderSubjectRow(displayPrev, prevSc, displayCurr, currSc);
        drawnCount++;
    });

    if(drawnCount === 0) {
        return `<div class="unmatched-empty">이번 학기에 ${d.name} 과목을 응시하지 않았습니다.</div>`;
    }
    return rowsHtml;
}

/* 과목별 1:1 비교 행 — 원점수 / 석차·등급 / 백분위
   요청 양식: 중복 라벨 제거, 석차·등급은 "52위 → 1위 ▲51 / 2등급 → 1등급 ▲1" 형식으로 볼드+컬러 적용 */
function renderSubjectRow(prevSubj, prevSc, currSubj, currSc) {
    // ---- 원점수 ----
    let scoreDelta = currSc.score - prevSc.score;
    let scoreDeltaHtml = '';
    if(Math.abs(scoreDelta) >= 0.05) {
        let cls = scoreDelta > 0 ? 'up' : 'down';
        let arrow = scoreDelta > 0 ? '▲' : '▼';
        scoreDeltaHtml = `<span class="delta-inline ${cls}">${arrow}${Math.abs(scoreDelta).toFixed(1)}</span>`;
    }

    // ---- 석차 (낮을수록 우수) ----
    let pRankNum = parseInt(prevSc.rank, 10);
    let cRankNum = parseInt(currSc.rank, 10);
    let rankPrevTxt = !isNaN(pRankNum) ? `${pRankNum}위` : '-';
    let rankCurrTxt = !isNaN(cRankNum) ? `${cRankNum}위` : '-';
    let rankDeltaHtml = '';
    if(!isNaN(pRankNum) && !isNaN(cRankNum)) {
        let diff = cRankNum - pRankNum;            // +면 등수 떨어짐(하락)
        if(diff !== 0) {
            let cls = diff < 0 ? 'up' : 'down';
            let arrow = diff < 0 ? '▲' : '▼';
            rankDeltaHtml = `<span class="delta-inline ${cls}">${arrow}${Math.abs(diff)}</span>`;
        }
    }

    // ---- 등급 (낮을수록 우수) ----
    let pGr = (prevSc.grade !== null && prevSc.grade !== undefined && !isNaN(prevSc.grade)) ? prevSc.grade : null;
    let cGr = (currSc.grade !== null && currSc.grade !== undefined && !isNaN(currSc.grade)) ? currSc.grade : null;
    let gradePrevTxt = pGr !== null ? `${pGr}등급` : '-';
    let gradeCurrTxt = cGr !== null ? `${cGr}등급` : '-';
    let gradeDeltaHtml = '';
    if(pGr !== null && cGr !== null) {
        let diff = cGr - pGr;                       // +면 등급 올라감(하락)
        if(Math.abs(diff) >= 0.05) {
            let cls = diff < 0 ? 'up' : 'down';
            let arrow = diff < 0 ? '▲' : '▼';
            let absStr = Number.isInteger(diff) ? Math.abs(diff).toString() : Math.abs(diff).toFixed(1);
            gradeDeltaHtml = `<span class="delta-inline ${cls}">${arrow}${absStr}</span>`;
        }
    }

    // ---- 백분위 (= 100 - 석차백분율, 클수록 우수) ----
    let pctPrev = prevSc.pct !== null ? (100 - prevSc.pct) : null;
    let pctCurr = currSc.pct !== null ? (100 - currSc.pct) : null;
    let pctDelta = (pctPrev !== null && pctCurr !== null) ? (pctCurr - pctPrev) : null;
    let pctDeltaHtml = '';
    if(pctDelta !== null && Math.abs(pctDelta) >= 0.05) {
        let cls = pctDelta > 0 ? 'up' : 'down';
        let arrow = pctDelta > 0 ? '▲' : '▼';
        pctDeltaHtml = `<span class="delta-inline ${cls}">${arrow}${Math.abs(pctDelta).toFixed(1)}</span>`;
    }
    let pctMainHtml = (pctPrev !== null && pctCurr !== null)
        ? `<span class="metric-main">${pctPrev.toFixed(1)} → ${pctCurr.toFixed(1)}</span>${pctDeltaHtml}`
        : '-';

    // 같은 과목이면 "문학", 다르면 "공통국어 → 문학"
    let subjLabel = matchSubjectName(prevSubj, currSubj)
        ? currSubj
        : prevSubj + '<span class="arrow">→</span>' + currSubj;

    return `
        <div class="subj-row">
            <div class="subj-name">${subjLabel}</div>
            <div><span class="col-label">원점수 </span><span class="metric-main">${prevSc.score} → ${currSc.score}</span>${scoreDeltaHtml}</div>
            <div>
                <span class="rank-line"><span class="col-label">석차 </span><span class="metric-main">${rankPrevTxt} → ${rankCurrTxt}</span>${rankDeltaHtml}</span>
                <span class="grade-line"><span class="col-label">등급 </span><span class="metric-main">${gradePrevTxt} → ${gradeCurrTxt}</span>${gradeDeltaHtml}</span>
            </div>
            <div><span class="col-label">백분위 </span>${pctMainHtml}</div>
        </div>
    `;
}

/* 신뢰 불가 교과군 (사·과) — 좌우 분리 패널, 각 패널은 4열 미니 표
   왼쪽: 직전 통합과목들 / 오른쪽: 이번 학기 선택과목들 (직전·이번 비교는 1:1 불가하므로 화살표 없음) */
function buildUnreliableBody(s25, s26, d) {
    // 좌측 (2025 직전 통합과목)
    let leftRows = '';
    d.prev.forEach(prevSubj => {
        let key = matchSubject(s25.scores, prevSubj);
        if(key && s25.scores[key] && s25.scores[key].score !== null) {
            leftRows += renderSplitRow(prevSubj, s25.scores[key]);
        }
    });
    let leftBody;
    if(leftRows) {
        leftBody = `
            <div class="split-row head">
                <div>과목</div><div>원점수</div><div>석차·등급</div><div>백분위</div>
            </div>
            ${leftRows}
        `;
    } else {
        leftBody = '<div class="unmatched-empty">직전 학기 응시 과목 없음</div>';
    }

    // 우측 (2026 이번 선택과목)
    let rightRows = '';
    d.curr.forEach(currSubj => {
        let key = matchSubject(s26.scores, currSubj);
        if(key && s26.scores[key] && s26.scores[key].score !== null) {
            rightRows += renderSplitRow(currSubj, s26.scores[key]);
        }
    });
    let rightBody;
    if(rightRows) {
        rightBody = `
            <div class="split-row head">
                <div>과목</div><div>원점수</div><div>석차·등급</div><div>백분위</div>
            </div>
            ${rightRows}
        `;
    } else {
        rightBody = '<div class="unmatched-empty">이번 학기 응시 과목 없음</div>';
    }

    return `
        <div class="unmatched-note">
            직전 통합과목 → 이번 학기 선택과목으로 분리되어, 좌(2025 직전)·우(2026 이번) 패널로 나누어 표시합니다.
            응시 집단(전체 vs 선택자)도 다르므로 <b>참고용으로만</b> 해석하세요.
        </div>
        <div class="unmatched-grid">
            <div class="unmatched-col">
                <h5>📘 2025 직전 (통합과목)</h5>
                ${leftBody}
            </div>
            <div class="unmatched-col">
                <h5>📗 2026 이번 (선택과목)</h5>
                ${rightBody}
            </div>
        </div>
    `;
}

/* split 패널 안의 한 행 — 과목명 / 원점수 / 석차·등급 / 백분위 (화살표·델타 없음) */
function renderSplitRow(subj, sc) {
    let rankNum = parseInt(sc.rank, 10);
    let rank = !isNaN(rankNum) ? `${rankNum}위` : '-';
    let grade = (sc.grade !== null && sc.grade !== undefined && !isNaN(sc.grade)) ? `${sc.grade}등급` : '-';
    let pct = sc.pct !== null ? (100 - sc.pct).toFixed(1) : '-';
    let scoreTxt = (sc.score !== null && sc.score !== undefined) ? sc.score : '-';
    return `
        <div class="split-row">
            <div class="s-name"><span class="col-label">과목 </span>${subj}</div>
            <div class="s-main"><span class="col-label">원점수 </span>${scoreTxt}</div>
            <div class="s-sub"><span class="col-label">석차·등급 </span>${rank} · ${grade}</div>
            <div class="s-main"><span class="col-label">백분위 </span>${pct}</div>
        </div>
    `;
}

/* 탭2 상단 학생 정보 카드 채우기 (s25 유무와 무관하게 이번 학기 데이터로 항상 채움) */
function fillTab2InfoCard(s25, s26) {
    // 헤더
    document.getElementById('t2Name').textContent = s26.name;
    document.getElementById('t2Meta').textContent = `2학년 ${s26.cls}반 ${s26.num}번`;

    const avgG = avgGrade(s26);
    document.getElementById('t2AvgGrade').textContent = avgG !== null ? avgG.toFixed(1) : '-';

    // 플래그 배지
    const flag = computeFlag(s26);
    document.getElementById('t2FlagBadge').innerHTML = flagBadgeHtml(flag);

    // 응시 과목 수
    let subjCount = 0;
    subjects2026.forEach(subj => {
        let key = matchSubject(s26.scores, subj);
        if(key && s26.scores[key] && s26.scores[key].score !== null) subjCount++;
    });

    // 평균 백분위 / 순위
    const avgP = avgPercentile(s26);
    const cRank = gradeRank(s26, 'class');
    const tRank = gradeRank(s26, 'total');
    document.getElementById('t2InfoAvgGrade').textContent = avgG !== null ? avgG.toFixed(2) + ' 등급' : '-';
    document.getElementById('t2InfoAvgPct').textContent = avgP !== null ? avgP.toFixed(1) : '-';
    document.getElementById('t2InfoClassRank').textContent = cRank ? `${cRank.rank} / ${cRank.total} 위` : '-';
    document.getElementById('t2InfoTotalRank').textContent = tRank ? `${tRank.rank} / ${tRank.total} 위` : '-';
    document.getElementById('t2InfoSubjCount').textContent = subjCount + ' 과목';

    // 직전 학기 대비 추세
    const trendEl = document.getElementById('t2InfoTrend');
    if(s25) {
        let up = 0, down = 0;
        DOMAINS.forEach(d => {
            let vp = avgMetric(s25, d.prev, 'percentile');
            let vc = avgMetric(s26, d.curr, 'percentile');
            if(vp !== null && vc !== null && Math.abs(vc - vp) > 1) {
                if(vc < vp) up++; else down++;
            }
        });
        if(up > down)      trendEl.innerHTML = '<span class="delta-up">전반적 향상</span>';
        else if(down > up) trendEl.innerHTML = '<span class="delta-down">일부 하락</span>';
        else               trendEl.innerHTML = '<span class="delta-flat">유지</span>';
    } else {
        trendEl.innerHTML = '<span style="color:#999;">직전 데이터 없음</span>';
    }
}

/* 아코디언 펼침/접기 토글 */
function toggleDomainCard(headerEl) {
    headerEl.closest('.domain-card').classList.toggle('collapsed');
}

/* =========================================================
   탭3 : 학급/과목별 통계 요약  (기존 로직 유지 + 미응시 제외)
   ========================================================= */
function buildGlobalStats() {
    let stats = {};

    Object.values(db2026).forEach(st => {
        Object.entries(st.scores).forEach(([subj, data]) => {
            if(!stats[subj]) stats[subj] = { sum: 0, count: 0, scores: [] };
            if(data.score !== null) {
                stats[subj].sum += data.score;
                stats[subj].count++;
                stats[subj].scores.push(data.score);
            }
        });
    });

    const tb = document.querySelector('#statsTable tbody');
    tb.innerHTML = '';

    Object.keys(stats).forEach(subj => {
        if(stats[subj].count < 1) return;
        let avg = stats[subj].sum / stats[subj].count;
        let sorted = stats[subj].scores.slice().sort((a,b) => b-a);
        let len = sorted.length;

        const getCut = (pct) => {
            let idx = Math.max(0, Math.ceil(len * pct) - 1);
            return sorted[idx] !== undefined ? sorted[idx].toFixed(1) : "-";
        };

        tb.innerHTML += `<tr>
            <td style="font-weight:bold;">${subj}</td>
            <td>${avg.toFixed(1)}점</td>
            <td>${stats[subj].count}명</td>
            <td style="color:#15803D; font-weight:bold;">${getCut(0.10)}점</td>
            <td style="color:#2563EB; font-weight:bold;">${getCut(0.34)}점</td>
            <td style="color:#D97706; font-weight:bold;">${getCut(0.66)}점</td>
            <td style="color:#9333EA; font-weight:bold;">${getCut(0.90)}점</td>
        </tr>`;
    });

    if(!tb.innerHTML) {
        tb.innerHTML = '<tr><td colspan="7" style="color:#999;">통계 데이터가 없습니다.</td></tr>';
    }
}

/* =========================================================
   탭3 : 과목별 분포 및 통계 — 전체 응시 인원 기준 반투명 점 분포
   ========================================================= */
function renderTab3(s26) {
    const container = document.getElementById('distContainer');
    container.innerHTML = '';

    const cohort = Object.values(db2026);
    document.getElementById('distSubtitle').textContent = `${s26.name} · 전체 응시 인원 기준`;

    subjects2026.forEach(subj => {
        let myKey = matchSubject(s26.scores, subj);
        if(!myKey || !s26.scores[myKey] || s26.scores[myKey].score === null) return;
        let myScore = s26.scores[myKey].score;

        let peerScores = [];
        cohort.forEach(st => {
            let k = matchSubject(st.scores, subj);
            if(k && st.scores[k] && st.scores[k].score !== null) {
                peerScores.push({ hakbun: st.hakbun, score: st.scores[k].score });
            }
        });
        if(peerScores.length === 0) return;

        let avg = peerScores.reduce((a,b) => a + b.score, 0) / peerScores.length;
        let rank = 1 + peerScores.filter(p => p.score > myScore).length;

        // 다른 학생 점은 반투명(겹칠수록 진하게), 선택 학생은 맨 나중에 그려 위에 표시
        let dots = '';
        peerScores.forEach(p => {
            if(p.hakbun === s26.hakbun) return;
            let left = Math.max(0, Math.min(100, p.score));
            dots += `<div class="dist-dot" style="left:${left}%;"></div>`;
        });
        let meLeft = Math.max(0, Math.min(100, myScore));
        dots += `<div class="dist-dot me" style="left:${meLeft}%;" title="${myScore}점"></div>`;

        container.insertAdjacentHTML('beforeend', `
            <div class="dist-row">
                <div class="dist-head">
                    <span class="d-subj">${subj}</span>
                    <span class="d-info">내 점수 <strong style="color:var(--primary);">${myScore}점</strong>
                        · 전체 ${peerScores.length}명 중 <strong>${rank}등</strong>
                        · 평균 ${avg.toFixed(1)}점</span>
                </div>
                <div class="dist-strip">
                    <div class="dist-avg" style="left:${Math.max(0,Math.min(100,avg))}%;" title="평균 ${avg.toFixed(1)}점"></div>
                    ${dots}
                </div>
                <div class="dist-scale"><span>0점</span><span>50점</span><span>100점</span></div>
            </div>
        `);
    });

    if(!container.innerHTML) {
        container.innerHTML = '<p style="color:#999;">표시할 과목 데이터가 없습니다.</p>';
    }
}

/* =========================================================
   상담 카드 출력
   ========================================================= */
function printCounselCard() {
    const hakbun = document.getElementById('studentSelect').value;
    if(!hakbun || !db2026[hakbun]) {
        alert("먼저 학생을 선택해주세요.");
        return;
    }

    // 탭1의 레이더 차트를 고해상도로 캡처
    var srcCanvas = document.getElementById('radarChart');
    var radarImgData = '';
    if(srcCanvas && charts.radar) {
        try {
            // 2배 해상도 캔버스에 다시 그려서 선명하게 캡처
            var scale = 2;
            var w = srcCanvas.width;
            var h = srcCanvas.height;
            var hiResCanvas = document.createElement('canvas');
            hiResCanvas.width = w * scale;
            hiResCanvas.height = h * scale;
            var hiCtx = hiResCanvas.getContext('2d');
            hiCtx.scale(scale, scale);
            hiCtx.drawImage(srcCanvas, 0, 0, w, h);
            radarImgData = hiResCanvas.toDataURL('image/png');
        } catch(e) {
            // 고해상도 실패 시 원본 캡처
            try { radarImgData = srcCanvas.toDataURL('image/png'); } catch(e2) {}
        }
    }

    // 상담 카드 생성
    buildCounselCard();

    // 캡처한 이미지를 상담 카드의 img에 삽입
    var ccImg = document.getElementById('ccRadarImg');
    if(ccImg && radarImgData) {
        ccImg.src = radarImgData;
    }

    // 인쇄
    setTimeout(function() {
        window.print();
    }, 200);
}

function buildCounselCard() {
    const hakbun = document.getElementById('studentSelect').value;
    const s26 = db2026[hakbun];
    let s25 = db2025[hakbun] || Object.values(db2025).find(st => st.name === s26.name);

    // --- 헤더 ---
    document.getElementById('ccTitle').textContent = `${s26.name} 학생 상담 카드`;
    let prevFile = document.getElementById('file2025');
    let currFile = document.getElementById('file2026');
    let prevLabel = (prevFile && prevFile.files && prevFile.files[0]) ? prevFile.files[0].name.replace(/\.(xlsx|xls|csv)$/i,'').replace(/_/g,' ') : '직전 시험';
    let currLabel = (currFile && currFile.files && currFile.files[0]) ? currFile.files[0].name.replace(/\.(xlsx|xls|csv)$/i,'').replace(/_/g,' ') : '이번 시험';
    document.getElementById('ccSubtitle').textContent =
        `2학년 ${s26.cls}반 ${s26.num}번 | ${currLabel} (직전: ${prevLabel})`;
    const flag = computeFlag(s26);
    document.getElementById('ccFlag').innerHTML = flagBadgeHtml(flag);

    // --- 1. 개인별 성적 종합 (탭1과 동일) ---
    const avgG = avgGrade(s26);

    // 학생 정보 패널
    const avgP = avgPercentile(s26);
    const cRank = gradeRank(s26, 'class');
    const tRank = gradeRank(s26, 'total');
    const cnt = countSubjects(s26);
    let infoHtml = '';
    infoHtml += '<tr><td class="info-key">평균등급</td><td class="info-val">' + (avgG !== null ? avgG.toFixed(2) + ' 등급' : '-') + '</td></tr>';
    infoHtml += '<tr><td class="info-key">평균 백분위</td><td class="info-val">' + (avgP !== null ? avgP.toFixed(1) : '-') + '</td></tr>';
    infoHtml += '<tr><td class="info-key">학급 내 등급 순위</td><td class="info-val">' + (cRank ? cRank.rank + ' / ' + cRank.total + ' 위' : '-') + '</td></tr>';
    infoHtml += '<tr><td class="info-key">전체 등급 순위</td><td class="info-val">' + (tRank ? tRank.rank + ' / ' + tRank.total + ' 위' : '-') + '</td></tr>';
    infoHtml += '<tr><td class="info-key">응시 과목 수</td><td class="info-val">' + cnt + ' 과목</td></tr>';
    // 직전 학기 대비 추세
    let trendHtml = '-';
    if(s25) {
        let up = 0, down = 0;
        DOMAINS.forEach(function(d) {
            let vp = avgMetric(s25, d.prev, 'percentile');
            let vc = avgMetric(s26, d.curr, 'percentile');
            if(vp !== null && vc !== null && Math.abs(vc - vp) > 1) {
                if(vc < vp) up++; else down++;
            }
        });
        if(up > down) trendHtml = '<span class="delta-up">전반적 향상</span>';
        else if(down > up) trendHtml = '<span class="delta-down">일부 하락</span>';
        else trendHtml = '<span class="delta-flat">유지</span>';
    } else {
        trendHtml = '<span style="color:#999;">직전 데이터 없음</span>';
    }
    infoHtml += '<tr><td class="info-key">직전 학기 대비</td><td class="info-val">' + trendHtml + '</td></tr>';
    document.getElementById('ccInfoBody').innerHTML = infoHtml;

    // 레이더 차트
    let rows = [];
    subjects2026.forEach(function(subj) {
        let key = matchSubject(s26.scores, subj);
        if(key && s26.scores[key] && s26.scores[key].score !== null) {
            rows.push({ subj: subj, score: s26.scores[key].score, grade: s26.scores[key].grade, rank: s26.scores[key].rank, pct: s26.scores[key].pct });
        }
    });
    let radarLabels = rows.map(function(r) { return r.subj; });
    let radarData = rows.map(function(r) { return r.pct !== null ? (100 - r.pct) : 0; });

    // 레이더 차트는 printCounselCard에서 탭1 canvas를 캡처해서 img로 넣음

    // 과목별 상세 분석 표
    document.getElementById('ccSubjCountLabel').textContent = '응시 ' + rows.length + '과목';
    const stb = document.querySelector('#ccSubjectTable tbody');
    stb.innerHTML = '';
    rows.forEach(function(r) {
        let pctile = r.pct !== null ? (100 - r.pct).toFixed(1) : '-';
        let pctColor = r.pct !== null ? ((100 - r.pct) >= 70 ? '#15803D' : ((100 - r.pct) >= 40 ? '#D97706' : '#DC2626')) : '#999';
        stb.innerHTML += '<tr>'
            + '<td>' + r.subj + '</td>'
            + '<td>' + r.score + '</td>'
            + '<td>' + achvBadge(r.score) + '</td>'
            + '<td>' + gradeBadge(r.grade) + '</td>'
            + '<td>' + (r.rank || '-') + '</td>'
            + '<td style="color:' + pctColor + '; font-weight:bold;">' + pctile + '</td>'
            + '</tr>';
    });
    if(!stb.innerHTML) stb.innerHTML = '<tr><td colspan="6" style="color:#999;">응시 과목 데이터가 없습니다.</td></tr>';

    // --- 2. 교과별 직전 대비 성장 추이 ---
    const gtb = document.querySelector('#ccGrowthTable tbody');
    gtb.innerHTML = '';

    if(!s25) {
        gtb.innerHTML = '<tr><td colspan="13" style="color:#999;">직전 시험 데이터 없음</td></tr>';
    } else {
        DOMAINS.forEach(function(d) {
            let currSubjects = [];
            d.curr.forEach(function(currSubj) {
                let ck = matchSubject(s26.scores, currSubj);
                let csc = (ck && s26.scores[ck] && s26.scores[ck].score !== null) ? s26.scores[ck] : null;
                if(!csc) return;

                // 직전 매칭
                let psc = null, pSubj = currSubj;
                let sk = matchSubject(s25.scores, currSubj);
                if(sk && s25.scores[sk] && s25.scores[sk].score !== null) {
                    psc = s25.scores[sk];
                } else {
                    for(let i = 0; i < d.prev.length; i++) {
                        let pk = matchSubject(s25.scores, d.prev[i]);
                        if(pk && s25.scores[pk] && s25.scores[pk].score !== null) {
                            psc = s25.scores[pk]; pSubj = d.prev[i]; break;
                        }
                    }
                }
                currSubjects.push({ currSubj: currSubj, csc: csc, psc: psc, pSubj: pSubj });
            });

            if(currSubjects.length === 0) return;

            currSubjects.forEach(function(item, idx) {
                let row = '<tr>';
                // 과목명 (가운데 정렬)
                let label = matchSubjectName(item.pSubj, item.currSubj) ? item.currSubj : item.pSubj + '→' + item.currSubj;
                row += '<td style="text-align:center;">' + label + '</td>';

                if(item.psc) {
                    let pp = item.psc.pct !== null ? (100 - item.psc.pct).toFixed(1) : '-';
                    let cp = item.csc.pct !== null ? (100 - item.csc.pct).toFixed(1) : '-';
                    let pRank = parseInt(item.psc.rank, 10);
                    let cRank = parseInt(item.csc.rank, 10);

                    // 직전
                    row += '<td class="cc-prev-bg">' + item.psc.score + '</td>';
                    row += '<td class="cc-prev-bg">' + (!isNaN(pRank) ? pRank : '-') + '</td>';
                    row += '<td class="cc-prev-bg">' + (item.psc.grade || '-') + '</td>';
                    row += '<td class="cc-prev-bg">' + pp + '</td>';

                    // 이번
                    row += '<td>' + item.csc.score + '</td>';
                    row += '<td>' + (!isNaN(cRank) ? cRank : '-') + '</td>';
                    row += '<td>' + (item.csc.grade || '-') + '</td>';
                    row += '<td>' + cp + '</td>';

                    // 변화 4칼럼
                    var scoreDiff = item.csc.score - item.psc.score;
                    row += '<td class="cc-delta-bg ' + (scoreDiff > 0 ? 'cc-up' : (scoreDiff < 0 ? 'cc-down' : 'cc-flat')) + '">'
                        + (scoreDiff > 0 ? '▲' : (scoreDiff < 0 ? '▼' : '-')) + (scoreDiff !== 0 ? Math.abs(scoreDiff).toFixed(1) : '') + '</td>';

                    var rankDiff = (!isNaN(pRank) && !isNaN(cRank)) ? (cRank - pRank) : null;
                    if(rankDiff !== null && rankDiff !== 0) {
                        row += '<td class="cc-delta-bg ' + (rankDiff < 0 ? 'cc-up' : 'cc-down') + '">'
                            + (rankDiff < 0 ? '▲' : '▼') + Math.abs(rankDiff) + '</td>';
                    } else {
                        row += '<td class="cc-delta-bg cc-flat">-</td>';
                    }

                    var gPrev = item.psc.grade, gCurr = item.csc.grade;
                    if(gPrev && gCurr && gPrev !== gCurr) {
                        var gDiff = gCurr - gPrev;
                        row += '<td class="cc-delta-bg ' + (gDiff < 0 ? 'cc-up' : 'cc-down') + '">'
                            + (gDiff < 0 ? '▲' : '▼') + Math.abs(gDiff) + '</td>';
                    } else {
                        row += '<td class="cc-delta-bg cc-flat">-</td>';
                    }

                    var ppN = parseFloat(pp), cpN = parseFloat(cp);
                    if(!isNaN(ppN) && !isNaN(cpN)) {
                        var pDiff = cpN - ppN;
                        row += '<td class="cc-delta-bg ' + (pDiff > 0 ? 'cc-up' : (pDiff < 0 ? 'cc-down' : 'cc-flat')) + '">'
                            + (pDiff > 0 ? '▲' : (pDiff < 0 ? '▼' : '-')) + (pDiff !== 0 ? Math.abs(pDiff).toFixed(1) : '') + '</td>';
                    } else {
                        row += '<td class="cc-delta-bg cc-flat">-</td>';
                    }
                } else {
                    // 직전 없음
                    row += '<td class="cc-prev-bg" colspan="4" style="color:#999;">직전 미응시</td>';
                    row += '<td>' + item.csc.score + '</td>';
                    row += '<td>' + (parseInt(item.csc.rank,10) || '-') + '</td>';
                    row += '<td>' + (item.csc.grade || '-') + '</td>';
                    row += '<td>' + (item.csc.pct !== null ? (100 - item.csc.pct).toFixed(1) : '-') + '</td>';
                    row += '<td class="cc-delta-bg" colspan="4" style="color:#999;">비교 불가</td>';
                }
                row += '</tr>';
                gtb.innerHTML += row;
            });
        });
    }
    if(!gtb.innerHTML) gtb.innerHTML = '<tr><td colspan="13" style="color:#999;">표시할 데이터가 없습니다.</td></tr>';

    // --- 3. 과목별 분포 위치 ---
    var distC = document.getElementById('ccDistContainer');
    distC.innerHTML = '';
    var cohort = Object.values(db2026);

    subjects2026.forEach(function(subj) {
        var myKey = matchSubject(s26.scores, subj);
        if(!myKey || !s26.scores[myKey] || s26.scores[myKey].score === null) return;
        var myScore = s26.scores[myKey].score;

        var peerScores = [];
        cohort.forEach(function(st) {
            var k = matchSubject(st.scores, subj);
            if(k && st.scores[k] && st.scores[k].score !== null) peerScores.push(st.scores[k].score);
        });
        if(peerScores.length === 0) return;

        var avg = peerScores.reduce(function(a, b) { return a + b; }, 0) / peerScores.length;
        var rank = 1 + peerScores.filter(function(p) { return p > myScore; }).length;
        var meLeft = Math.max(0, Math.min(100, myScore));
        var avgLeft = Math.max(0, Math.min(100, avg));

        distC.innerHTML += '<div class="cc-dist-row">'
            + '<div class="cc-dist-head"><span style="font-weight:bold;">' + subj + '</span>'
            + '<span style="color:#666;">내 점수 <b style="color:var(--primary);">' + myScore + '점</b>'
            + ' · 전체 ' + peerScores.length + '명 중 <b>' + rank + '등</b>'
            + ' · 평균 ' + avg.toFixed(1) + '점</span></div>'
            + '<div class="cc-dist-strip">'
            + '<div class="cc-dist-avg" style="left:' + avgLeft + '%;"></div>'
            + '<div class="cc-dist-me" style="left:' + meLeft + '%;"></div>'
            + '</div>'
            + '<div class="cc-dist-scale"><span>0점</span><span>50점</span><span>100점</span></div>'
            + '</div>';
    });
    if(!distC.innerHTML) distC.innerHTML = '<p style="color:#999;">표시할 과목 데이터가 없습니다.</p>';
}

/* =========================================================
   파일 업로드 UI — 파일명 표시 + 드래그앤드롭
   ========================================================= */
/* 업로드된 파일명에서 시험 정보 추출하여 컨트롤바 + 통계 제목에 표시 */
function updateExamLabel() {
    let prevFile = document.getElementById('file2025');
    let currFile = document.getElementById('file2026');

    let prevName = (prevFile && prevFile.files && prevFile.files[0]) ? prevFile.files[0].name : '직전 시험';
    let currName = (currFile && currFile.files && currFile.files[0]) ? currFile.files[0].name : '이번 시험';

    // 파일명에서 확장자 제거 + 언더스코어를 공백으로
    function cleanName(fn) {
        return fn.replace(/\.(xlsx|xls|csv)$/i, '').replace(/_/g, ' ').trim();
    }

    let prevLabel = cleanName(prevName);
    let currLabel = cleanName(currName);

    // 컨트롤바 시험 표시
    let examEl = document.getElementById('examLabel');
    if(examEl) {
        examEl.textContent = '📋 ' + currLabel;
        examEl.title = '직전: ' + prevLabel + ' → 이번: ' + currLabel;
    }

    // 등급 컷 통계 제목
    let statsLabel = document.getElementById('statsExamLabel');
    if(statsLabel) {
        statsLabel.textContent = currLabel;
    }
}

function showFileName(input, dropId) {
    let drop = document.getElementById(dropId);
    let existing = drop.querySelector('.upload-file-name');
    if(existing) existing.remove();
    if(input.files && input.files[0]) {
        let nameEl = document.createElement('div');
        nameEl.className = 'upload-file-name';
        nameEl.textContent = '✅ ' + input.files[0].name;
        drop.appendChild(nameEl);
    }
}

document.addEventListener('DOMContentLoaded', function() {
    ['drop2025', 'drop2026'].forEach(function(dropId) {
        let drop = document.getElementById(dropId);
        if(!drop) return;
        let fileInput = drop.querySelector('input[type="file"]');

        drop.addEventListener('dragover', function(e) {
            e.preventDefault();
            drop.style.borderColor = '#5F4A8B';
            drop.style.background = '#f0eef5';
        });
        drop.addEventListener('dragleave', function(e) {
            e.preventDefault();
            drop.style.borderColor = '#ccc';
            drop.style.background = '#fafafa';
        });
        drop.addEventListener('drop', function(e) {
            e.preventDefault();
            drop.style.borderColor = '#ccc';
            drop.style.background = '#fafafa';
            if(e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                fileInput.files = e.dataTransfer.files;
                showFileName(fileInput, dropId);
            }
        });
    });
});

/* =========================================================
   분석결과 HTML 저장 — 전체 학생 데이터 포함 self-contained HTML
   ========================================================= */
function saveFullHTML() {
    if(Object.keys(db2026).length === 0) {
        alert('먼저 데이터를 업로드하고 분석을 실행해주세요.');
        return;
    }

    // 1) 현재 <head> 안의 <style> 블록 수집 (인라인 CSS)
    let headStyles = '';
    document.querySelectorAll('head style').forEach(function(s) {
        headStyles += '<style>' + s.textContent + '</style>\n';
    });

    // 2) 외부 style.css의 규칙을 인라인으로 수집
    let externalCSS = '';
    for(let i = 0; i < document.styleSheets.length; i++) {
        try {
            let sheet = document.styleSheets[i];
            if(sheet.href && sheet.href.includes('style.css')) {
                let rules = sheet.cssRules || sheet.rules;
                let css = '';
                for(let j = 0; j < rules.length; j++) css += rules[j].cssText + '\n';
                externalCSS = '<style>/* style.css 인라인 */\n' + css + '</style>\n';
            }
        } catch(e) {}
    }

    if (externalCSS === '') {
        externalCSS = `
        <link rel="stylesheet" href="./style.css">
        <style>
            /* 로컬 환경 CSS 수집 실패 대비 안전장치 (탭 디자인 & 차트 크기 제한) */
            .tabs { display: flex; width: 100%; }
            .tab-btn { flex: 1; }
            
            /* [수정] 탭1 레이더 차트는 원 목적대로 중앙 정렬 정사각형 크기로 제약 */
            #radarChart { max-width: 320px !important; max-height: 320px !important; margin: 0 auto; }
            
            /* [수정] 탭2 막대 차트는 원본 style.css 규격대로 가로 제약 없이 높이만 부모 영역에 맞춤 */
            #barChart { max-height: 240px !important; width: 100% !important; }
        </style>\n`;
    }

    // 3) 외부 라이브러리 스크립트 태그 (CDN 참조 유지)
    let libScripts = '';
    document.querySelectorAll('head script[src]').forEach(function(s) {
        libScripts += '<script src="' + s.src + '"><\/script>\n';
    });

    // 4) Pretendard 폰트 링크
    let fontLink = '';
    document.querySelectorAll('head link[href*="pretendard"]').forEach(function(l) {
        fontLink += '<link rel="stylesheet" crossorigin href="' + l.href + '" />\n';
    });

    // 5) workspace 영역의 현재 innerHTML (분석 결과 포함)
    let workspace = document.getElementById('workspace');
    let wsHTML = workspace.outerHTML.replace('display: none', 'display: block').replace('display:none', 'display:block');
    // workspace가 숨겨져 있을 수 있으므로 강제 표시
    wsHTML = wsHTML.replace('id="workspace"', 'id="workspace" style="display:block;"');

    // 6) 상단 헤더 (업로드 섹션 제외)
    let header = document.querySelector('.main-header');
    let headerHTML = '';
    if(header) {
        let clone = header.cloneNode(true);
        let p = clone.querySelector('p');
        if(p) p.textContent = '분석결과 저장본 — 오프라인에서 열람 가능';
        headerHTML = clone.outerHTML;
    }

    // 7) DB 데이터를 JSON으로 직렬화
    let dbJSON = 'var db2025=' + JSON.stringify(db2025) + ';\n'
        + 'var db2026=' + JSON.stringify(db2026) + ';\n'
        + 'var subjects2025=' + JSON.stringify(subjects2025) + ';\n'
        + 'var subjects2026=' + JSON.stringify(subjects2026) + ';\n';

    // 8) app.js 전체 함수를 문자열로 수집 (현재 정의된 모든 함수)
    let fnNames = [
        'readExcelAsync','buildDB','processData','populateDropdowns','filterStudents',
        'navStudent','updateNavButtons','onSearchInput','pickSearch','onSearchKey',
        'switchTab','updateDashboard','matchSubject','matchSubjectName','getGroup','buildDomains',
        'gradeBadge','pctGauge',
        'achvBadge','toPercentile','avgPercentile','gradeRank',
        'renderTab1','computeFlag','flagBadgeHtml','avgMetric','avgGrade','countSubjects',
        'strengthsWeaknesses','renderTab2','toggleDomainCard','fillTab2InfoCard',
        'buildReliableBody','renderSubjectRow','buildUnreliableBody','renderSplitRow',
        'buildGlobalStats','renderTab3',
        'printCounselCard','buildCounselCard',
        'showFileName','saveFullHTML','updateExamLabel'
    ];
    let fnSource = '';
    fnNames.forEach(function(name) {
        try {
            let fn = window[name];
            if(typeof fn === 'function') {
                fnSource += fn.toString() + '\n\n';
            }
        } catch(e) {}
    });

    // 전역 변수 선언 (charts 객체 등)
    let globalVars = 'var charts = {};\nvar currentOptionList = [];\n'
        + 'var DOMAINS = ' + JSON.stringify(DOMAINS) + ';\n'
        + 'var SUBJ_TO_GROUP = ' + JSON.stringify(SUBJ_TO_GROUP) + ';\n';

    // DOMContentLoaded 초기화
    let initScript = `
document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('workspace').style.display = 'block';
    buildDomains();
    populateDropdowns();
    buildGlobalStats();
});
document.addEventListener('click', function(e) {
    if(!e.target.closest('.search-wrap')) {
        var box = document.getElementById('searchResults');
        if(box) box.style.display = 'none';
    }
});
`;

    // 9) 완성된 HTML 조합
    let fullHTML = `<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>내신 성적 분석 결과</title>
    ${fontLink}
    ${libScripts}
    ${externalCSS}
    ${headStyles}
</head>
<body>
    <div class="dashboard-container">
        ${headerHTML}
        ${wsHTML}
    </div>
    <script>
    ${dbJSON}
    ${globalVars}
    ${fnSource}
    ${initScript}
    <\/script>
</body>
</html>`;

    // 10) 다운로드
    let blob = new Blob([fullHTML], { type: 'text/html; charset=utf-8' });
    let url = URL.createObjectURL(blob);
    let a = document.createElement('a');
    let now = new Date();
    let dateStr = now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0');
    a.href = url;
    a.download = '내신분석_' + dateStr + '.html';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}
