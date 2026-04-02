/**
 * Tender Intelligence Dashboard - app.js (Pro v3)
 * Full Modular Implementation: Parser, Comparator, Filter, Renderer, Charts, UI
 */

const App = (() => {
    // --- State ---
    let db = JSON.parse(localStorage.getItem('tender_db_v3')) || [];
    let state = {
        filters: {
            search: '',
            country: '',
            status: '',
            deadline: '',
            employer: '',
            tab: 'all',
            includeHidden: false
        },
        sort: {
            field: 'deadline',
            order: 'desc' // asc, desc
        },
        selectedId: null,
        theme: localStorage.getItem('tender_theme') || 'light'
    };

    // --- Configuration ---
    const CONFIG = {
        sheetTarget: '2026',
        headerKeywords: ["project name", "country", "employer", "사업명", "발주처", "project", "code"],
        essentialFields: ['projectName', 'employer', 'participation', 'award', 'deadlineRaw', 'budget', 'krw', 'currentStatus', 'jv', 'evaluation'],
        // ID 생성 로직 (중복 방지 및 정합성 보장)
        norm: (s) => String(s || "").replace(/\s+/g, '').toLowerCase(),
        cleanCode: (c) => {
            if (typeof c === 'object' && c !== null && 'cleaned' in c) return c;
            const str = String(c || "").trim();
            let type = null;
            let cleaned = str;
            const upper = str.toUpperCase();
            if (upper.endsWith("_P")) { type = "P"; cleaned = str.slice(0, -2); }
            else if (upper.endsWith("_N")) { type = "N"; cleaned = str.slice(0, -2); }
            else if (upper.endsWith("_PB")) { type = "PB"; cleaned = str.slice(0, -3); }
            else if (upper.endsWith("_DB")) { type = "PB"; cleaned = str.slice(0, -3); } // _DB -> _PB (D&B)
            return { cleaned: cleaned.trim(), type };
        },
        generateId: (code, pName, emp) => {
            const { cleaned } = CONFIG.cleanValue(code) === 'EMPTY' ? { cleaned: "" } : CONFIG.cleanCode(code);
            const n = CONFIG.norm;
            return cleaned ? n(cleaned) : `${n(pName)}_${n(emp)}`;
        },
        cleanValue: (v) => {
            const s = String(v || "").trim().toUpperCase();
            if (s === "X" || s === "-" || s === "" || s === "(EMPTY)") return "EMPTY";
            return s;
        },
        rates: { PLN: 405 } // 1 PLN = 405 KRW
    };

    const Helpers = {
        extractNum(str) {
            const m = String(str).match(/[\d,.]+/);
            return m ? parseFloat(m[0].replace(/,/g, '')) : NaN;
        },
        calculateKRW(budgetStr, existingKRW) {
            const str = String(budgetStr || "");
            
            // 1. 사업명/예산 칸에 이미 'XX억원' 식이 명시되어 있다면 그걸 우선 사용 (사용자 요청)
            const krwMatch = str.match(/([\d,.]+)\s*억원/);
            if (krwMatch) {
                return `${krwMatch[1]} 억원`;
            }

            // 2. 별도의 한화(KRW) 컬럼에 값이 있는 경우 그걸 사용
            if (existingKRW && String(existingKRW).trim() !== "") return String(existingKRW).trim();
            
            // 3. 없으면 폴란드 환율(405원) 기반 자동 환산
            const num = this.extractNum(str);
            if (str.includes('~')) {
                const parts = str.split('~').map(p => {
                    const n = this.extractNum(p);
                    return !isNaN(n) ? ((n * CONFIG.rates.PLN) / 100000000).toFixed(1) : null;
                }).filter(v => v !== null);
                return parts.length > 0 ? `${parts.join(' ~ ')} 억원` : "";
            } else {
                if (!isNaN(num) && num > 100) {
                    const won = (num * CONFIG.rates.PLN) / 100000000;
                    return `${won.toFixed(1)} 억원`;
                }
            }
            return "";
        }
    };



    // --- Modules ---

    const DateUtils = {
        parseDate(str) {
            if (!str) return null;
            const s = String(str).trim();
            // Supports YYYY.MM.DD, YYYY-MM-DD, YYYY/MM/DD
            const match = s.match(/(\d{4})[./-](\d{2})[./-](\d{2})/);
            if (match) return new Date(`${match[1]}-${match[2]}-${match[3]}`);
            return null;
        },
        getDDay(date) {
            if (!date) return null;
            const now = new Date();
            now.setHours(0, 0, 0, 0);
            const diff = date - now;
            return Math.ceil(diff / (1000 * 60 * 60 * 24));
        }
    };

    const ExcelParser = {
        async read(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array', cellStyles: true });
                    const targetSheet = workbook.SheetNames.find(n => n.includes(CONFIG.sheetTarget)) || workbook.SheetNames[0];
                    const ws = workbook.Sheets[targetSheet];
                    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
                    resolve({ rows, ws, sheetName: targetSheet });
                };
                reader.onerror = reject;
                reader.readAsArrayBuffer(file);
            });
        },

        process(rows, ws) {
            const data = [];
            let currentHeaders = null;

            rows.forEach((row, rowIndex) => {
                const rowStr = row.map(c => String(c).trim().toLowerCase());
                const isHeader = CONFIG.headerKeywords.some(k => rowStr.includes(k));

                if (isHeader) {
                    currentHeaders = row.map(c => String(c).trim().replace(/\s+/g, ' '));
                    return;
                }

                if (currentHeaders && row.length > 0) {
                    const obj = {};
                    currentHeaders.forEach((h, i) => { if (h) obj[h] = row[i]; });

                    // 엑셀 시각적 속성 감지 (숨김 행) - 절대 인덱스 계산 반영
                    const rangeIdx = XLSX.utils.decode_range(ws['!ref'] || "A1:Z500");
                    const absoluteRowIndex = rangeIdx.s.r + rowIndex;
                    
                    let isRowHidden = false;
                    const rowMeta = ws['!rows'] ? ws['!rows'][absoluteRowIndex] : null;
                    if (rowMeta && (rowMeta.hidden || rowMeta.hpt === 0 || rowMeta.hpx === 0)) {
                        isRowHidden = true;
                    }

                    const getRawCell = (keys) => {
                        for (const pk of keys) {
                            const idx = currentHeaders.findIndex(h => h.toLowerCase() === pk.toLowerCase());
                            if (idx !== -1) {
                                const cellAddr = XLSX.utils.encode_cell({ r: rowIndex, c: idx });
                                return ws[cellAddr];
                            }
                        }
                        return null;
                    };

                    const pnCell = getRawCell(["Project Name", "사업명", "project"]);
                    const isStriked = (pnCell && pnCell.s && pnCell.s.font && pnCell.s.font.strike);

                    const get = (keys) => {
                        for (const pk of keys) {
                            const foundKey = Object.keys(obj).find(ok => ok.toLowerCase() === pk.toLowerCase());
                            if (foundKey) return obj[foundKey];
                        }
                        return "";
                    };

                    const projectName = String(get(["Project Name", "사업명", "project"])).trim();
                    const no = String(get(["No", "no."])).trim();
                    const employer = get(["Employer", "발주처"]);
                    const { cleaned: codeStr, type: workType } = CONFIG.cleanCode(get(["Code", "번호", "입찰번호"]));
                    const id = CONFIG.generateId(codeStr, projectName, employer);

                    // Filtering non-data rows
                    if (!projectName || projectName.toLowerCase() === "project name" || projectName.toLowerCase().includes("country")) return;
                    if (no === "" && code === "") return;

                    const deadline = DateUtils.parseDate(get(["Deadline", "제출마감일", "마감일"]));


                    data.push({
                        id: id,
                        code: codeStr,
                        workType: workType,
                        no: no,
                        projectName,
                        employer: employer || "N/A",
                        country: get(["Country", "국가"]) || "N/A",
                        noticeDate: get(["Notice Date", "공고일"]),
                        deadlineRaw: String(get(["Deadline", "마감일"]) || "").split('(')[0].trim(),
                        deadlineObj: deadline,
                        dDay: DateUtils.getDDay(deadline),
                        participation: String(get(["Participation", "참여여부"]) || "").toUpperCase().trim(),
                        award: (() => {
                            const s = String(get(["수주", "Award"]) || "").toUpperCase().trim();
                            return (s === 'O' || s === 'WINNER' || s === 'WON' || s === 'SELECTED') ? 'O' : 'X';
                        })(),
                        budget: get(["Final Budget", "Estimated Budget", "Budget"]),
                        krw: Helpers.calculateKRW(get(["Final Budget", "Estimated Budget", "Budget"]), get(["KRW", "원화"])),
                        currentStatus: get(["Current status", "진행현황", "상태"]),
                        jv: get(["JV", "컨소시엄"]),
                        evaluation: get(["Evaluation", "평가방식", "Evaluation Criteria"]),
                        tenderLink: get(["Site Link", "Link", "링크"]),
                        category1: get(["code1", "분류1"]),
                        category2: get(["code2", "분류2"]),
                        isHidden: isRowHidden || isStriked || false,
                        updateType: 'UNCHANGED',
                        changedFields: []
                    });
                }
            });
            return data;
        }
    };

    const Comparator = {
        diff(oldData, newData) {
            const now = new Date().toISOString();
            const finalMap = new Map();
            
            // 1. 기존 데이터 로드 + ID 마이그레이션 (중요: 과거 데이터의 ID를 새로운 규격으로 강제 업데이트)
            oldData.forEach(item => {
                const migratedId = CONFIG.generateId(item.code, item.projectName, item.employer);
                const updated = { ...item, id: migratedId };
                
                if (updated.updateType === 'NEW' || updated.updateType === 'UPDATED') {
                    updated.updateType = 'UNCHANGED';
                }
                updated.foundInCurrentSync = false; 
                finalMap.set(updated.id, updated);
            });

            // 2. 신규 데이터 비교 및 병합
            newData.forEach(row => {
                let existing = finalMap.get(row.id);

                // [지능형 매칭] 직접적인 ID(코드) 일치가 실패했을 때: 사업명(Project Name)이 동일한 것이 있는지 2차 대조
                if (!existing) {
                    const normName = CONFIG.norm(row.projectName);
                    existing = Array.from(finalMap.values()).find(v => 
                        !v.foundInCurrentSync && CONFIG.norm(v.projectName) === normName
                    );
                    
                    if (existing) {
                        // 이름이 같은 과거 항목을 찾았다면, 새로운 ID(바뀐 코드)로 업데이트하고 매칭을 이어갑니다.
                        finalMap.delete(existing.id);
                        existing.id = row.id; 
                        finalMap.set(row.id, existing);
                    }
                }

                if (row.isHidden) {
                    // 숨김/취소선 항목은 명시적으로 'CLOSED' 상태로 처리
                    finalMap.set(row.id, { ...row, updateType: 'CLOSED', lastFound: now, foundInCurrentSync: true });
                } else if (!existing) {
                    // 완전히 처음 보는 사업
                    finalMap.set(row.id, { ...row, updateType: 'NEW', lastFound: now, foundInCurrentSync: true });
                } else {
                    // 동일 사업 발견 (ID 또는 사업명 매칭 성공)
                    const diffs = [];
                    // 주요 필드들 정밀 대조 (값이 의미상 같은지 체크)
                    CONFIG.essentialFields.forEach(f => {
                        const ov = existing[f];
                        const nv = row[f];
                        if (CONFIG.cleanValue(ov) !== CONFIG.cleanValue(nv)) {
                            // 실제 변화가 있는 경우만 diff 추가
                            diffs.push({ field: f, old: ov || "(Empty)", new: nv || "(Empty)" });
                        }
                    });

                    if (diffs.length > 0) {
                        // 내용이 변경된 건
                        finalMap.set(row.id, { 
                            ...existing, 
                            ...row, 
                            updateType: 'UPDATED', 
                            changedFields: diffs,
                            lastFound: now,
                            foundInCurrentSync: true
                        });
                    } else {
                        // 동일한 내용
                        finalMap.set(row.id, { ...existing, lastFound: now, foundInCurrentSync: true });
                    }
                }
            });

            // 3. 누각 항목 CLS 처리
            finalMap.forEach(item => {
                if (!item.foundInCurrentSync && item.updateType !== 'CLOSED') {
                    item.updateType = 'CLOSED';
                }
            });

            return Array.from(finalMap.values());
        }
    };

    const UI = {
        init() {
            this.migrate();
            this.bindEvents();
            this.render();
            if (window.lucide) lucide.createIcons();
        },

        migrate() {
            let changed = false;
            db = db.map(d => {
                let updatedObj = { ...d };
                
                // 1. 환율 재계산
                const newKRW = Helpers.calculateKRW(d.budget, d.krw);
                if (newKRW !== d.krw) {
                    updatedObj.krw = newKRW;
                    changed = true;
                }
                
                // 2. 마감일 간소화 (괄호 내역 제거)
                if (d.deadlineRaw && String(d.deadlineRaw).includes('(')) {
                    updatedObj.deadlineRaw = String(d.deadlineRaw).split('(')[0].trim();
                    changed = true;
                }

                // 3. 코드 접미사(_P, _N, _PB) 제거 및 워크타입 추출 마이그레이션 (+ [object Object] 복구)
                let currentCode = updatedObj.code;
                if (typeof currentCode === 'object' && currentCode !== null) {
                    currentCode = currentCode.cleaned || ""; 
                    changed = true;
                }

                if (currentCode) {
                    const { cleaned, type } = CONFIG.cleanCode(currentCode);
                    // 이미 정제된 코드인 경우 라벨(type)이 null로 반환되어 기존 라벨이 사라지는 것 방지
                    if (cleaned !== currentCode) {
                        updatedObj.code = cleaned;
                        changed = true;
                    }
                    if (type && type !== d.workType) {
                        updatedObj.workType = type;
                        changed = true;
                    }
                    
                    if (changed) {
                        updatedObj.id = CONFIG.generateId(cleaned, d.projectName, d.employer);
                        changed = true;
                    }
                }

                // 4. 가짜 변경 알림 (X, -, 공백 간 변경) 청소
                if (updatedObj.updateType === 'UPDATED' && updatedObj.changedFields) {
                    const realDiffs = updatedObj.changedFields.filter(f => {
                        return CONFIG.cleanValue(f.old) !== CONFIG.cleanValue(f.new);
                    });
                    if (realDiffs.length !== updatedObj.changedFields.length) {
                        updatedObj.changedFields = realDiffs;
                        changed = true;
                    }
                    if (updatedObj.changedFields.length === 0) {
                        updatedObj.updateType = 'UNCHANGED';
                        changed = true;
                    }
                }

                return updatedObj;
            });
            if (changed) localStorage.setItem('tender_db_v3', JSON.stringify(db));
        },

        bindEvents() {
            // File Upload & Drag/Drop
            const inp = document.getElementById('excel-upload');
            const dropZone = document.getElementById('drop-zone');

            if (inp) {
                document.getElementById('upload-btn').onclick = () => inp.click();
                inp.onchange = (e) => this.handleUpload(e.target.files[0]);
            }

            if (dropZone) {
                dropZone.ondragover = (e) => { e.preventDefault(); dropZone.classList.add('dragover'); };
                dropZone.ondragleave = () => dropZone.classList.remove('dragover');
                dropZone.ondrop = (e) => {
                    e.preventDefault();
                    dropZone.classList.remove('dragover');
                    if (e.dataTransfer.files.length) this.handleUpload(e.dataTransfer.files[0]);
                };
            }

            document.getElementById('toggle-hidden').onchange = (e) => {
                state.filters.includeHidden = e.target.checked;
                this.render();
            };

            // Search & Global Filters
            document.getElementById('global-search').oninput = (e) => { state.filters.search = e.target.value; this.render(); };
            document.getElementById('filter-country').onchange = (e) => { state.filters.country = e.target.value; this.render(); };
            document.getElementById('filter-employer').onchange = (e) => { state.filters.employer = e.target.value; this.render(); };
            document.getElementById('filter-status').onchange = (e) => { state.filters.status = e.target.value; this.render(); };
            document.getElementById('filter-deadline').onchange = (e) => { state.filters.deadline = e.target.value; this.render(); };

            document.getElementById('reset-filters').onclick = () => {
                state.filters = { search: '', country: '', status: '', deadline: '', employer: '', tab: 'all', includeHidden: false };
                state.sort = { field: 'deadline', order: 'desc' };
                document.getElementById('global-search').value = '';
                document.getElementById('filter-country').value = '';
                document.getElementById('filter-employer').value = '';
                document.getElementById('filter-status').value = '';
                document.getElementById('filter-deadline').value = '';
                document.getElementById('toggle-hidden').checked = false;
                this.render(); // Changed from renderAll()
            };

            // Card Filters
            document.querySelectorAll('.kpi-card').forEach(card => {
                card.onclick = () => {
                    const f = card.dataset.filter;
                    const tab = document.querySelector(`.tab-btn[data-tab="${f}"]`);
                    if (tab) tab.click();
                    else { state.filters.tab = f; this.render(); }
                };
            });

            // Table Sorting
            document.querySelectorAll('#main-table th').forEach(th => {
                const field = th.dataset.sort;
                if (!field) return;
                th.style.cursor = 'pointer';
                th.onclick = () => {
                    if (state.sort.field === field) {
                        state.sort.order = state.sort.order === 'asc' ? 'desc' : 'asc';
                    } else {
                        state.sort.field = field;
                        state.sort.order = 'asc';
                    }
                    this.render();
                };
            });

            // Theme
            document.getElementById('theme-toggle').onclick = () => {
                state.theme = state.theme === 'dark' ? 'light' : 'dark';
                document.body.className = state.theme + '-theme';
                localStorage.setItem('tender_theme', state.theme);
                this.renderCharts(); // Redraw for contrast
            };

            // Side Panel Close
            document.querySelector('.close-panel').onclick = () => document.getElementById('side-panel').classList.remove('active');
            
            // Reset DB
            document.getElementById('reset-db').onclick = () => {
                if(confirm("모든 데이터를 초기화하시겠습니까?")) {
                    localStorage.removeItem('tender_db_v3');
                    db = [];
                    this.render();
                }
            };

            // Export
            document.getElementById('export-excel').onclick = () => this.export();

            // Backup & Restore (JSON)
            const backupBtn = document.getElementById('backup-db');
            if (backupBtn) backupBtn.onclick = () => this.backupJSON();
            
            const jsonInp = document.getElementById('json-upload');
            if (jsonInp) {
                document.getElementById('restore-db').onclick = () => jsonInp.click();
                jsonInp.onchange = (e) => this.restoreJSON(e.target.files[0]);
            }
            
            // Copy
            document.getElementById('copy-summary').onclick = () => this.copyToClipboard();
        },

        async handleUpload(file) {
            if (!file) return;
            const loader = document.getElementById('loading-overlay');
            if (loader) loader.classList.remove('hidden');
            try {
                const { rows, ws, sheetName } = await ExcelParser.read(file);
                const newData = ExcelParser.process(rows, ws);
                db = Comparator.diff(db, newData);
                localStorage.setItem('tender_db_v3', JSON.stringify(db));
                this.render();
                alert(`[${sheetName}] 시트 분석 완료`);
            } catch (err) {
                console.error(err);
                alert("엑셀 처리 중 오류 발생");
            } finally {
                if (loader) loader.classList.add('hidden');
            }
        },

        render() {
            try {
                this.syncTabs();
                this.populateSelects();
                this.renderKPIs();
                this.renderTable();
                this.renderSummary();
                this.renderCharts();
                if (window.lucide) lucide.createIcons();
                
                const timeStr = new Date().toLocaleTimeString();
                const updateEl = document.getElementById('last-update-time');
                if (updateEl) updateEl.innerText = `Update: ${timeStr}`;
            } catch (err) {
                console.error("Render Error:", err);
            }
        },

        syncTabs() {
            document.querySelectorAll('.tab-btn').forEach(btn => {
                btn.classList.toggle('active', btn.dataset.tab === state.filters.tab);
            });
        },

        populateSelects() {
            const countries = [...new Set(db.map(d => d.country))].filter(Boolean).sort();
            const statuses = [...new Set(db.map(d => d.updateType))].filter(Boolean);
            
            // 발주처 약어 추출 (첫 단어만)
            const employerShorts = [...new Set(db.map(d => String(d.employer || "").split(' ')[0]))].filter(Boolean).sort();
            
            const countrySelect = document.getElementById('filter-country');
            const employerSelect = document.getElementById('filter-employer');
            const statusSelect = document.getElementById('filter-status');

            // Keep only first (All)
            countrySelect.innerHTML = '<option value="">All Countries</option>';
            countries.forEach(c => countrySelect.innerHTML += `<option value="${c}">${c}</option>`);

            employerSelect.innerHTML = '<option value="">All Employers</option>';
            employerShorts.forEach(e => employerSelect.innerHTML += `<option value="${e}">${e}</option>`);

            statusSelect.innerHTML = '<option value="">All Status</option>';
            statuses.forEach(s => statusSelect.innerHTML += `<option value="${s}">${s}</option>`);
        },

        renderKPIs() {
            const activeDb = db.filter(d => state.filters.includeHidden || !d.isHidden);
            document.getElementById('kpi-total').innerText = activeDb.length;
            document.getElementById('kpi-participation').innerText = activeDb.filter(d => d.participation === 'O').length;
            document.getElementById('kpi-award').innerText = activeDb.filter(d => d.award === 'O').length;
            document.getElementById('kpi-new').innerText = activeDb.filter(d => d.updateType === 'NEW').length;
            document.getElementById('kpi-updated').innerText = activeDb.filter(d => d.updateType === 'UPDATED').length;
            document.getElementById('kpi-risk').innerText = activeDb.filter(d => d.dDay !== null && d.dDay >= 0 && d.dDay <= 7).length;
        },

        renderTable() {
            const { search, country, status, deadline, tab, includeHidden } = state.filters;
            const { field, order } = state.sort;
            const tableBody = document.getElementById('table-body');

            const filtered = db.filter(d => {
                // 0. 취소/숨김 항목 필터
                if (!includeHidden && d.isHidden) return false;

                const pName = (d.projectName || "").toLowerCase();
                const emp = (d.employer || "").toLowerCase();
                const code = (d.code || "").toLowerCase();
                const q = (search || "").toLowerCase();

                const matchSearch = !search || pName.includes(q) || emp.includes(q) || code.includes(q);
                const matchCountry = !country || d.country === country;
                const matchStatus = !status || d.updateType === status;
                const matchDeadline = !deadline || (d.dDay !== null && d.dDay <= parseInt(deadline));
                const matchEmployer = !state.filters.employer || emp.includes(state.filters.employer.toLowerCase());
                
                let matchTab = true;
                if (tab === 'new-only') matchTab = d.updateType === 'NEW';
                if (tab === 'updated-only') matchTab = d.updateType === 'UPDATED';
                if (tab === 'deadline-risk') matchTab = (d.dDay !== null && d.dDay >= 0 && d.dDay <= 7);
                if (tab === 'participation-o') matchTab = d.participation === 'O';
                if (tab === 'award-o') matchTab = d.award === 'O';

                return matchSearch && matchCountry && matchStatus && matchDeadline && matchEmployer && matchTab;
            });

            // Sort Logic
            filtered.sort((a, b) => {
                let valA, valB;
                
                if (field === 'deadline') {
                    // Sort by dDay (null = end)
                    valA = (a.dDay !== null && a.dDay !== undefined) ? a.dDay : (order === 'asc' ? 99999 : -99999);
                    valB = (b.dDay !== null && b.dDay !== undefined) ? b.dDay : (order === 'asc' ? 99999 : -99999);
                } else if (field === 'projectName' || field === 'employer' || field === 'code') {
                    valA = String(a[field] || "").toLowerCase();
                    valB = String(b[field] || "").toLowerCase();
                } else if (field === 'status') {
                    valA = String(a.updateType || "").toLowerCase();
                    valB = String(b.updateType || "").toLowerCase();
                } else {
                    valA = a[field] || 0;
                    valB = b[field] || 0;
                }

                if (valA < valB) return order === 'asc' ? -1 : 1;
                if (valA > valB) return order === 'asc' ? 1 : -1;
                return 0;
            });

            tableBody.innerHTML = filtered.length ? '' : '<tr><td colspan="10" class="empty-msg">No results found for current filters.</td></tr>';
            
            // Update Sort Icons in Header
            document.querySelectorAll('#main-table th').forEach(th => {
                const f = th.dataset.sort;
                let icon = th.querySelector('.sort-icon');
                if (icon) icon.remove();
                if (f === field) {
                    const arrow = order === 'asc' ? '↑' : '↓';
                    const span = document.createElement('span');
                    span.className = 'sort-icon';
                    span.innerText = arrow;
                    span.style.color = 'var(--accent-primary)';
                    th.appendChild(span);
                }

                // Resizer Ensure
                if (!th.querySelector('.resizer')) {
                    const resizer = document.createElement('div');
                    resizer.className = 'resizer';
                    th.appendChild(resizer);
                    this.initResizable(th);
                }
            });

            filtered.forEach(d => {
                try {
                    const tr = document.createElement('tr');
                    if (d.isHidden) tr.classList.add('row-cancelled');
                    const dDay = d.dDay;
                    let ddayText = '-';
                    let ddayClass = '';

                    if (dDay !== null && dDay !== undefined) {
                        if (dDay < 0) {
                            ddayText = `D+${Math.abs(dDay)}`;
                            ddayClass = 'dd-gray';
                        } else if (dDay === 0) {
                            ddayText = 'D-Day';
                            ddayClass = 'dd-red';
                        } else {
                            ddayText = `D-${dDay}`;
                            ddayClass = dDay <= 3 ? 'dd-red' : (dDay <= 7 ? 'dd-orange' : '');
                        }
                    }

                    const ut = d.updateType || 'UNCHANGED';
                    const displayType = ut === 'UPDATED' ? 'UPD' : (ut === 'UNCHANGED' ? 'SAME' : (ut === 'CLOSED' ? 'CLS' : ut));
                    
                    tr.innerHTML = `
                        <td><div class="cell-content"><span class="badge badge-${ut.toLowerCase()}">${displayType}</span></div></td>
                        <td><div class="cell-content code-text">
                            ${d.code || '-'}
                            ${d.workType ? `<span class="work-badge work-${d.workType.toLowerCase()}">${d.workType}</span>` : ''}
                        </div></td>
                        <td><div class="cell-content projectName-wrap" title="${d.projectName || ''}">${d.projectName || '-'}</div></td>
                        <td><div class="cell-content employer-wrap">${d.employer || '-'}</div></td>
                        <td class="align-center"><div class="cell-content part-${d.participation === 'O' ? 'o' : 'x'}">${d.participation || 'X'}</div></td>
                        <td class="align-center"><div class="cell-content part-${d.award === 'O' ? 'o' : 'x'}">${d.award || 'X'}</div></td>
                        <td><div class="cell-content deadline-wrap"><div class="${ddayClass}">${ddayText}</div><small>${d.deadlineRaw || '-'}</small></div></td>
                        <td><div class="cell-content jv-wrap">${d.jv || '-'}</div></td>
                        <td><div class="cell-content krw-text">${d.krw || '-'}</div></td>
                        <td><div class="cell-content action-btn"><button class="btn-icon sm info-btn" title="상세보기"><i data-lucide="info"></i></button></div></td>
                    `;
                    const btn = tr.querySelector('.info-btn');
                    if (btn) btn.onclick = () => this.showDetail(d);
                    tableBody.appendChild(tr);
                } catch (e) {
                    console.error("Row Render Error:", e, d);
                }
            });
            lucide.createIcons();
        },

        showDetail(d) {
            state.selectedId = d.id;
            const panel = document.getElementById('side-panel');
            panel.classList.add('active');

            document.getElementById('p-badge').innerText = d.code || 'NO CODE';
            document.getElementById('p-project-name').innerText = d.projectName;

            const mapList = (target, items) => {
                document.getElementById(target).innerHTML = items.map(i => `
                    <div class="info-cell"><span>${i.l}</span><span>${i.v || '-'}</span></div>
                `).join('');
            };

            mapList('p-essential', [{l:'Country', v:d.country}, {l:'Employer', v:d.employer}, {l:'Participation', v:d.participation}, {l:'Award', v:d.award}]);
            mapList('p-budget', [{l:'Budget (Raw)', v:d.budget}, {l:'KRW', v:d.krw}, {l:'Evaluation', v:d.evaluation}, {l:'JV', v:d.jv}]);
            mapList('p-dates', [
                {l:'Notice Date', v:d.noticeDate}, 
                {l:'Deadline', v:d.deadlineRaw}, 
                {l:'D-Day', v: d.dDay !== null && d.dDay !== undefined ? (d.dDay < 0 ? `D+${Math.abs(d.dDay)}` : (d.dDay === 0 ? 'D-Day' : `D-${d.dDay}`)) : '-'}
            ]);
            
            document.getElementById('p-current-status').innerText = d.currentStatus || 'N/A';
            
            // Diff
            const diffEl = document.getElementById('diff-section');
            if (d.updateType === 'UPDATED' && d.changedFields.length > 0) {
                diffEl.classList.remove('hidden');
                document.getElementById('diff-list').innerHTML = d.changedFields.map(f => `
                    <div class="diff-item">
                        <span style="font-weight:700">${f.field.toUpperCase()}</span>
                        <div style="display:flex; align-items:center; gap:8px">
                            <span class="old-val">${f.old || '(Empty)'}</span>
                            <i data-lucide="arrow-right" style="width:14px"></i>
                            <span class="new-val">${f.new || '(Empty)'}</span>
                        </div>
                    </div>
                `).join('');
                lucide.createIcons();
            } else {
                diffEl.classList.add('hidden');
            }

            const link = document.getElementById('tender-link');
            if (d.tenderLink) { 
                link.href = d.tenderLink; 
                link.classList.remove('hidden'); 
            } else { 
                link.classList.add('hidden'); 
            }
        },

        renderSummary() {
            const activeDb = db.filter(d => state.filters.includeHidden || !d.isHidden);
            const counts = {
                new: activeDb.filter(d => d.updateType === 'NEW').length,
                updated: activeDb.filter(d => d.updateType === 'UPDATED').length,
                risk: activeDb.filter(d => d.dDay !== null && d.dDay >= 0 && d.dDay <= 7).length,
                part: activeDb.filter(d => d.participation === 'O').length
            };
            const text = activeDb.length ? `[요약] 신규 ${counts.new}건 • 변경 ${counts.updated}건 • 마감임박 ${counts.risk}건 • 실참여 ${counts.part}건` : "파일을 업로드하면 보고서용 요약 문장이 자동 생성됩니다.";
            document.getElementById('summary-sentence').innerText = text;
        },

        renderCharts() {
            const activeDb = db.filter(d => state.filters.includeHidden || !d.isHidden);
            if (!activeDb.length) {
                ['chart-status', 'chart-award', 'chart-country'].forEach(id => {
                    if (this[id]) { this[id].destroy(); this[id] = null; }
                    const ctx = document.getElementById(id).getContext('2d');
                    ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
                });
                return;
            }
            const isDark = state.theme === 'dark';
            const colors = { text: isDark ? '#94a3b8' : '#64748b' };
            
            // Status Chart
            this.drawPie('chart-status', ['NEW', 'UPDATED', 'UNCHANGED'], 
                [activeDb.filter(d => d.updateType === 'NEW').length, activeDb.filter(d => d.updateType === 'UPDATED').length, activeDb.filter(d => d.updateType === 'UNCHANGED').length],
                ['#10b981', '#f59e0b', '#64748b'], colors.text);
            
            // Part/Award
            this.drawPie('chart-award', ['Award O', 'Participate O', 'Others'],
                [activeDb.filter(d => d.award === 'O').length, activeDb.filter(d => d.participation === 'O').length, activeDb.length],
                ['#10b981', '#3b82f6', '#334155'], colors.text);

            // Country Bar
            const cmap = {};
            activeDb.forEach(d => cmap[d.country] = (cmap[d.country] || 0) + 1);
            const topC = Object.entries(cmap).sort((a,b) => b[1]-a[1]).slice(0, 5);
            this.drawBar('chart-country', topC.map(c => c[0]), topC.map(c => c[1]), '#6366f1', colors.text);
        },

        drawPie(id, labels, data, colors, textColor) {
            const ctx = document.getElementById(id).getContext('2d');
            if (this[id]) this[id].destroy();
            this[id] = new Chart(ctx, {
                type: 'doughnut',
                data: { labels, datasets: [{ data, backgroundColor: colors, borderWidth: 0 }] },
                options: { responsive: true, maintainAspectRatio: false, cutout: '70%', plugins: { legend: { position: 'bottom', labels: { color: textColor, font: { size: 10 } } } } }
            });
        },

        drawBar(id, labels, data, color, textColor) {
            const ctx = document.getElementById(id).getContext('2d');
            if (this[id]) this[id].destroy();
            this[id] = new Chart(ctx, {
                type: 'bar',
                data: { labels, datasets: [{ data, backgroundColor: color, borderRadius: 5 }] },
                options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, ticks: { color: textColor } }, x: { ticks: { color: textColor } } }, plugins: { legend: { display:false } } }
            });
        },

        copyToClipboard() {
            const isDetailOpen = document.getElementById('side-panel').classList.contains('active');
            if (isDetailOpen && state.selectedId) {
                const d = db.find(item => item.id === state.selectedId);
                if (d) {
                    const text = `[사업상세] ${d.projectName}\n- 발주처: ${d.employer}\n- 현재상태: ${d.currentStatus || '-'}\n- 마감일: ${d.deadlineRaw || '-'}\n- 예산: ${d.krw || '-'}\n- JV구성: ${d.jv || '-'}`;
                    navigator.clipboard.writeText(text).then(() => alert("사업 상세 정보가 복사되었습니다."));
                    return;
                }
            }
            const sum = document.getElementById('summary-sentence').innerText;
            navigator.clipboard.writeText(sum).then(() => alert("보고서 요약이 복사되었습니다."));
        },

        export() {
            const sheetData = db.map(d => {
                const { id, deadlineObj, dDay, updateType, changedFields, ...rest } = d;
                return rest;
            });
            const ws = XLSX.utils.json_to_sheet(sheetData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "TenderData");
            XLSX.writeFile(wb, `Tender_Intelligence_${new Date().toISOString().slice(0,10)}.xlsx`);
        },

        backupJSON() {
            const blob = new Blob([JSON.stringify(db, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Tender_DB_Backup_${new Date().toISOString().slice(0,10)}.json`;
            a.click();
            URL.revokeObjectURL(url);
        },

        restoreJSON(file) {
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = JSON.parse(e.target.result);
                    if (Array.isArray(data)) {
                        if (confirm(`백업 파일에서 ${data.length}개의 항목을 불러오시겠습니까? 기존 데이터는 유지되며 중복은 자동으로 처리됩니다.`)) {
                            db = Comparator.diff(db, data); // Use diff logic to merge
                            localStorage.setItem('tender_db_v3', JSON.stringify(db));
                            this.render();
                            alert("복구 완료");
                        }
                    } else alert("올바른 백업 파일 형식이 아닙니다.");
                } catch (err) { alert("파일 읽기 오류"); }
            };
            reader.readAsText(file);
        },

        initResizable(th) {
            const resizer = th.querySelector('.resizer');
            let startX, startWidth;

            resizer.addEventListener('mousedown', (e) => {
                e.stopPropagation(); // Prevent sorting
                startX = e.pageX;
                startWidth = th.offsetWidth;
                th.classList.add('resizing');
                
                const onMouseMove = (e) => {
                    const width = startWidth + (e.pageX - startX);
                    if (width > 50) { 
                        th.style.width = width + 'px';
                        th.style.minWidth = width + 'px';
                    }
                };

                const onMouseUp = () => {
                    th.classList.remove('resizing');
                    document.removeEventListener('mousemove', onMouseMove);
                    document.removeEventListener('mouseup', onMouseUp);
                };

                document.addEventListener('mousemove', onMouseMove);
                document.addEventListener('mouseup', onMouseUp);
            });
        }
    };

    return { init: () => {
        UI.init();
        setTimeout(() => { if(window.lucide) lucide.createIcons(); }, 150);
    }};
})();

// Re-run icons on full load just in case
window.addEventListener('load', () => { if(window.lucide) lucide.createIcons(); });

App.init();
