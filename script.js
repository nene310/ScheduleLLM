/**
 * ScheduleLLM - Core Logic
 */

// Global State
let workbook = null;
let rawScheduleData = [];
let defaultTimeSlots = [
    { start: '08:20', end: '09:05' }, // 1
    { start: '09:15', end: '10:00' }, // 2
    { start: '10:20', end: '11:05' }, // 3
    { start: '11:15', end: '12:00' }, // 4
    { start: '14:30', end: '15:15' }, // 5
    { start: '15:25', end: '16:10' }, // 6
    { start: '16:30', end: '17:15' }, // 7
    { start: '17:15', end: '18:00' }, // 8
    { start: '19:10', end: '19:55' }, // 9
    { start: '19:55', end: '20:40' }  // 10
];

// Initialization
document.addEventListener('DOMContentLoaded', () => {
    initTimeSettings();
    initTimeSettingsCollapsible();

    // Load Configuration from config.js if available
    if (window.AppConfig) {
        const baseUrlInput = document.getElementById('llmBaseUrl');
        const modelInput = document.getElementById('llmModel');

        const llmApiUrl = window.AppConfig.llmApiUrl || window.AppConfig.backendUrl;
        if (llmApiUrl) baseUrlInput.value = llmApiUrl;
        if (window.AppConfig.model) modelInput.value = window.AppConfig.model;

        console.log("Environment configuration loaded.");
    }

    document.getElementById('fileUpload').addEventListener('change', handleFileUpload);
    document.getElementById('btnGenerate').addEventListener('click', generateSchedule);

    const retryA = document.getElementById('llmProgressRetry');
    const retryB = document.getElementById('llmProgressErrorRetry');
    if (retryA) retryA.addEventListener('click', () => generateSchedule());
    if (retryB) retryB.addEventListener('click', () => generateSchedule());

    const list = document.getElementById('llmRecognizedList');
    if (list) {
        list.addEventListener('click', (e) => {
            const btn = e.target && e.target.closest ? e.target.closest('button[data-idx]') : null;
            if (!btn) return;
            const idx = parseInt(btn.getAttribute('data-idx'), 10);
            if (!Number.isFinite(idx)) return;
            scheduleLLMProgressShowDetail(idx);
        });
    }

    const courseListPanel = document.getElementById('courseListPanel');
    const courseListToggle = document.getElementById('courseListToggle');
    if (courseListPanel && courseListToggle) {
        courseListPanel.classList.toggle('is-open', false);
        courseListToggle.setAttribute('aria-expanded', 'false');
        courseListToggle.addEventListener('click', () => {
            const open = !courseListPanel.classList.contains('is-open');
            courseListPanel.classList.toggle('is-open', open);
            courseListToggle.setAttribute('aria-expanded', open ? 'true' : 'false');
        });
    }

    if (typeof scheduleLLMSetCourseListVisible === 'function') {
        scheduleLLMSetCourseListVisible(false);
    } else if (courseListPanel) {
        courseListPanel.style.display = 'none';
    }

    // LLM Toggle listener
    const useLLMCheckbox = document.getElementById('useLLM');
    const llmConfigFields = document.getElementById('llmConfigFields');
    useLLMCheckbox.addEventListener('change', () => {
        llmConfigFields.style.display = useLLMCheckbox.checked ? 'block' : 'none';
    });
});

function initTimeSettings() {
    const legacy = document.getElementById('timeSettings');
    const always = document.getElementById('timeSettingsAlways');
    const extra = document.getElementById('timeSettingsExtra');

    const clear = (el) => {
        if (!el) return;
        while (el.firstChild) el.removeChild(el.firstChild);
    };

    if (always || extra) {
        clear(always);
        clear(extra);
    } else {
        clear(legacy);
    }

    const addRow = (container, slot, index) => {
        if (!container) return;
        const row = document.createElement('div');
        row.style.display = 'contents';
        row.innerHTML = `
            <span>${index + 1}</span>
            <input type="time" value="${slot.start}" data-idx="${index}" data-type="start">
            <input type="time" value="${slot.end}" data-idx="${index}" data-type="end">
        `;
        container.appendChild(row);
    };

    defaultTimeSlots.forEach((slot, index) => {
        if (always || extra) {
            if (index === 0) addRow(always, slot, index);
            else addRow(extra, slot, index);
            return;
        }
        addRow(legacy, slot, index);
    });

    scheduleLLMUpdateTimeSettingsSummary();
    const start = document.querySelector('input[type="time"][data-idx="0"][data-type="start"]');
    const end = document.querySelector('input[type="time"][data-idx="0"][data-type="end"]');
    if (start) start.addEventListener('input', scheduleLLMUpdateTimeSettingsSummary);
    if (end) end.addEventListener('input', scheduleLLMUpdateTimeSettingsSummary);
}

function scheduleLLMUpdateTimeSettingsSummary() {
    const start = document.querySelector('input[type="time"][data-idx="0"][data-type="start"]');
    const end = document.querySelector('input[type="time"][data-idx="0"][data-type="end"]');
    const el = document.getElementById('timeSettingsSummary');
    if (!el) return;
    const s = start && start.value ? start.value : '';
    const e = end && end.value ? end.value : '';
    const mid = (s || e) ? (s + (s && e ? '-' : '') + e) : '';
    el.textContent = '第1节' + (mid ? ' ' + mid : '');
}

function initTimeSettingsCollapsible() {
    const panel = document.getElementById('timeSettingsPanel');
    if (!panel) return;

    const header = document.getElementById('timeSettingsHeader');
    const icon = document.getElementById('timeSettingsToggleIcon');
    const hint = document.getElementById('timeSettingsHint');
    const extraWrap = document.getElementById('timeSettingsExtraWrap');

    const setExpanded = (expanded) => {
        panel.classList.toggle('expanded', expanded);
        if (header) header.setAttribute('aria-expanded', expanded ? 'true' : 'false');
        if (extraWrap) extraWrap.setAttribute('aria-hidden', expanded ? 'false' : 'true');
        if (icon) icon.textContent = expanded ? '∧' : '∨';
        if (hint) hint.textContent = expanded ? '点击收起' : '点击展开查看完整课表时间';
    };

    const toggle = () => {
        const expanded = panel.classList.contains('expanded');
        setExpanded(!expanded);
    };

    setExpanded(false);

    const onPanelClick = (e) => {
        const t = e && e.target;
        if (t && t.closest && t.closest('input, select, textarea, button, a, label')) return;
        toggle();
    };

    panel.addEventListener('click', onPanelClick);

    if (header) {
        header.addEventListener('keydown', (e) => {
            const k = e && e.key;
            if (k === 'Enter' || k === ' ') {
                e.preventDefault();
                toggle();
            }
        });
    }
}

// File Handling
// File Handling
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById('fileName').textContent = "正在读取: " + file.name;

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);

            if (typeof XLSX === 'undefined') {
                throw new Error("XLSX 库未加载，请检查网络或刷新页面");
            }

            workbook = XLSX.read(data, { type: 'array' });

            if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
                throw new Error("文件解析失败或无工作表");
            }

            // Assume first sheet
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            // Convert to JSON (Array of Arrays) to easier handling of messy headers
            rawScheduleData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (!rawScheduleData || rawScheduleData.length === 0) {
                throw new Error("工作表为空");
            }

            document.getElementById('fileName').textContent = "已加载: " + file.name;
            console.log("File loaded. Rows:", rawScheduleData.length);

            const calendarArea = document.getElementById('calendarArea');
            if (calendarArea) {
                const placeholder = calendarArea.querySelector('.placeholder-text');
                if (placeholder) {
                    placeholder.textContent = "课表文件已上传，请设定开学第一天日期，并设定节次时间。";
                    placeholder.classList.add('placeholder-uploaded');
                }

                const msg = "已成功上传课表[" + file.name + "]";
                let toast = calendarArea.querySelector('.calendar-toast');
                if (!toast) {
                    toast = document.createElement('div');
                    toast.className = 'calendar-toast no-print';
                    calendarArea.prepend(toast);
                }
                toast.textContent = msg;
                toast.classList.remove('calendar-toast-hide');
                if (window.scheduleLLMUploadMessageTimer) {
                    clearTimeout(window.scheduleLLMUploadMessageTimer);
                }
                window.scheduleLLMUploadMessageTimer = setTimeout(() => {
                    toast.classList.add('calendar-toast-hide');
                }, 5000);
            }

        } catch (err) {
            console.error(err);
            document.getElementById('fileName').textContent = "读取失败: " + err.message;
            alert("读取文件出错: " + err.message);
            rawScheduleData = []; // Reset on error
        }
    };
    reader.readAsArrayBuffer(file);
}

// Core Parsing Logic
function parseCourseString(cellContent) {
    if (!cellContent || typeof cellContent !== 'string') return [];

    // Pre-process: normalize OCR text and delimiters
    let cleanContent = normalizeOCRText(cellContent);
    cleanContent = cleanContent
        .replace(/◇/g, ' / ')
        .replace(/[《〈]/g, '(')
        .replace(/[》〉]/g, ')')
        .replace(/(\d+\s*[-~]\s*\d+|\d+)\s*[\r\n]+\s*周/g, '$1周');

    const independentCourses = [];

    // Phase 1: build week index table (debug / segmentation)
    const weekIndex = [];
    const weekGlobal = /(\d+\s*[-~]\s*\d+|\d+)\s*周/g;
    let wm;
    while ((wm = weekGlobal.exec(cleanContent)) !== null) {
        weekIndex.push({ idx: wm.index, text: wm[0] });
    }

    // Phase 2: segment courses by locating repeated "Name/...周" entry starts
    const entryRe = /(^|[\r\n]+)\s*([^\/\r\n]{2,}?)\s*\/\s*(?:\d{6,}\s*\/\s*)?(?:[\(（]?\s*\d+\s*[-~]\s*\d+\s*节[\)）]?\s*)?(?:\d+\s*[-~]\s*\d+|\d+)\s*周/g;
    const starts = [];
    let em;
    while ((em = entryRe.exec(cleanContent)) !== null) {
        starts.push(em.index + (em[1] ? em[1].length : 0));
    }

    if (starts.length > 1) {
        for (let i = 0; i < starts.length; i++) {
            const seg = cleanContent.slice(starts[i], starts[i + 1] || cleanContent.length).trim();
            if (seg) independentCourses.push(seg);
        }
    } else {
        // Fallback: line-buffer segmentation when entry starts are unclear
        const rawLines = cleanContent.split(/\r?\n/).map(s => s.trim()).filter(s => s.length > 0);
        let buffer = "";
        let bufferHasWeek = false;
        const hasWeekInfo = (str) => /(\d+[-~]\d+|\d+)周/.test(str);
        rawLines.forEach(line => {
            const lineHasWeek = hasWeekInfo(line);
            if (lineHasWeek && bufferHasWeek) {
                independentCourses.push(buffer);
                buffer = line;
                bufferHasWeek = true;
                return;
            }
            buffer = buffer.trim();
            if (buffer && !buffer.endsWith('/') && !buffer.endsWith('\n')) buffer += " " + line;
            else buffer = buffer + line;
            if (lineHasWeek) bufferHasWeek = true;
        });
        if (buffer) independentCourses.push(buffer);
    }

    const parsedCourses = [];

    independentCourses.forEach(courseStr => {
        // Smart Parsing: Handle variable formats
        // Format A: Name/Code/Weeks/Location/... (Standard)
        // Format B: Name/Weeks/Location (Simplified)
        // Format C: Missing slashes? (Not handled yet, assuming at least some delimiters)

        const courseStrClean = String(courseStr).replace(/[\r\n]+/g, "");

        // Pre-split handling: if "/" is missing but newlines were there (now spaces), 
        // we might have "Name 1-16周 Location". 
        // Let's ensure slashes exist around week info if missing.
        let normalizedStr = courseStrClean;
        const weekRegex = /(\d+[-~]\d+|\d+)周/;
        if (!normalizedStr.includes('/') && weekRegex.test(normalizedStr)) {
            normalizedStr = normalizedStr.replace(weekRegex, (match) => ` / ${match} / `);
        }

        // Also handle colons ':' and semicolons ';' as separators if they appear between what looks like class names
        // But simply replacing all colons might break times like 12:00 (though rare in course cells)
        // For now, let's treat colons and semicolons as potential delimiters during split or inner loop processing.
        // Actually, easiest is to replace colons/semicolons with slashes BEFORE split, 
        // IF they are not part of "人数:30" pattern.
        normalizedStr = normalizedStr.replace(/[:：;；]/g, '/'); 

        const parts = normalizedStr.split('/').map(s => s.trim()).filter(s => s.length > 0);

        let name = parts[0];
        let weeks = [];
        let location = "";
        let className = "";
        let periodRange = "";
        let weeksRaw = "";

        // Strategy: Find "Week" part specifically
        // It usually contains digit + "周"
        let weekPartIdx = -1;
        for (let i = 0; i < parts.length; i++) {
            if (weekRegex.test(parts[i])) {
                weekPartIdx = i;
                break;
            }
        }

        if (weekPartIdx !== -1) {
            // Found Weeks
            weeksRaw = parts[weekPartIdx];
            weeks = parseWeekString(parts[weekPartIdx]);

            // Name discovery:
            if (weekPartIdx > 0) {
                name = parts[0];
            } else if (weekPartIdx === 0) {
                // Case: "Software Engineering 1-16周 / Location"
                // Split parts[0] by the week regex
                const weekMatch = parts[0].match(weekRegex);
                if (weekMatch) {
                    const weekStr = weekMatch[0];
                    const splitPos = parts[0].indexOf(weekStr);
                    name = parts[0].substring(0, splitPos).trim() || "未知课程";
                } else {
                    name = "未知课程";
                }
            }

            // Location, Class Name, and "Other" info discovery
            const locs = [];
            const others = [];
            let prevWasBuildingOnly = false;

            for (let i = weekPartIdx + 1; i < parts.length; i++) {
                const p = parts[i];
                if (!p) continue;

                // 1. Identify Class Name
                // Rule: Contains digits/major + "班"/"级"/"专业", excludes "人数"
                const isPeopleCount = /人(数)?[:：°\s]*\d+|\d+\s*人/.test(p);
                // Enhanced class detection from instruction
                const isClassLike = (/((\d+|专业)[\s\S]*?[班级])/.test(p) || /^[A-Za-z0-9\u4e00-\u9fa5]+班$/.test(p)) && !isPeopleCount;

                if (isClassLike) {
                    if (className) {
                        className += "," + p;
                    } else {
                        className = p;
                    }
                    prevWasBuildingOnly = false;
                } else {
                    // 2. Identify Location
                    const token = String(p).trim();
                    const isLocationKeyword = /[楼室馆区教厅场苑基地中心工程]/.test(token);
                    const blacklistRegex = /(专业|导论|概论|基础|原理|必修|选修|考查|考试|讲课)/;
                    const hitsBlacklist = blacklistRegex.test(token);

                    const hasDigit = /\d/.test(token);
                    const hasLetter = /[A-Za-z]/.test(token);
                    const isPureDigit = /^\d+$/.test(token);

                    // Only treat pure-digit room numbers as location when they immediately follow a building-only token.
                    const isStandaloneRoomAfterBuilding = prevWasBuildingOnly && /^\d{3,4}$/.test(token);

                    const endsWithStrongSuffix = /[楼室馆区教厅场苑基地中心]$/.test(token);

                    let isLocation = false;
                    if (hasLetter && hasDigit) isLocation = true; // e.g. S103 / N608 / A101
                    else if (isStandaloneRoomAfterBuilding) isLocation = true; // e.g. 北苑电影大楼 + 414
                    else if (endsWithStrongSuffix) isLocation = true;
                    else if (isLocationKeyword && !hitsBlacklist) isLocation = true;

                    if (isPureDigit && !isStandaloneRoomAfterBuilding) {
                        // Likely teacher id / sequence / credit / other metadata (e.g. 426/0)
                        isLocation = false;
                    }

                    if (isLocation) {
                        locs.push(token);
                        prevWasBuildingOnly = isLocationKeyword && !hasDigit;
                    } else {
                        others.push(token);
                        prevWasBuildingOnly = false;
                    }
                }
            }
            location = locs.join(" ");
            // Note: 'others' array is available if we want to extract teacher later. 
            // e.g. const teacher = others.join(" ");

            // Try to extract period info from the week string BEFORE it was parsed/cleaned?
            // Actually parseWeekString cleaned it. 
            // We should check parts[weekPartIdx] for period info like "(1-2节)"
            const weekStrRaw = parts[weekPartIdx];
            const pMatch = weekStrRaw.match(/(\d+)\s*[-~]\s*(\d+)\s*节/);
            if (pMatch) {
                periodRange = `${pMatch[1]}-${pMatch[2]}`;
            } else {
                const pMatch2 = weekStrRaw.match(/\(([^)]*?)节\)/);
                if (pMatch2) periodRange = pMatch2[1];
            }
        } else {
            // FALLBACK: If no "X周" found
            console.warn("No weeks found for course:", courseStr);
            weeks = [];
            location = parts[1] || ""; 
            className = parts[2] || "";
        }

        // 2. Standardize all fields (Location, Name, ClassName, etc.)
        // This replaces individual simplify calls and ensures consistency
        const baseCourse = {
            rawName: name,
            displayName: name,
            weeks: Array.isArray(weeks) ? weeks : parseWeekString(weeks || ""),
            weeksRaw: weeksRaw,
            location: location,
            className: className,
            periodRange: periodRange,
            teacher: "",
            rawStr: courseStr
        };

        baseCourse.confidence = (
            (baseCourse.displayName && baseCourse.displayName !== "未知课程" ? 0.3 : 0) +
            (baseCourse.weeks && baseCourse.weeks.length ? 0.3 : 0) +
            (baseCourse.location ? 0.2 : 0) +
            (baseCourse.className ? 0.1 : 0) +
            (baseCourse.periodRange ? 0.1 : 0)
        );

        const stdCourse = standardizeCourseData(baseCourse);
        if (typeof window !== 'undefined' && window.__SCHEDULELLM_DEBUG_PARSE) {
            stdCourse._debug = { weekIndex: weekIndex, segmentCount: independentCourses.length };
        }

        parsedCourses.push(stdCourse);
    });

    return parsedCourses;
}

function standardizeCourseData(course) {
    // 1. Name: Remove all whitespace
    if (course.displayName) {
        course.displayName = simplifyName(course.displayName).replace(/\s+/g, "");
    }
    
    // 2. ClassName: Remove all whitespace, uppercase
    if (course.className) {
        course.className = course.className.replace(/^[\(（]/, '').replace(/[\)）]$/, ''); // Remove parens first
        course.className = course.className.replace(/\s+/g, "").toUpperCase();
    }

    // 3. Location, Building, Room
    // Extract components and reconstruct standardized location
    const locInfo = standardizeLocation(course.location);
    course.location = locInfo.location;
    course.building = locInfo.building;
    course.room = locInfo.room;

    // 4. Teacher: Remove all whitespace
    if (course.teacher) {
        course.teacher = course.teacher.replace(/\s+/g, "");
    } else {
        course.teacher = ""; // Ensure field exists
    }

    // Logging for audit
    // console.log("[Standardization]", course.displayName, course.location, course.className);

    return course;
}

function standardizeLocation(loc) {
    if (!loc) return { location: "待通知", building: "", room: "" };

    let s = loc;
    
    // 1. Basic Cleaning
    s = s.replace(/实验实训中心/g, "实训楼");
    s = s.replace(/(校区|场地|地点|场所)[：:]\s*/g, "");

    // 2. Remove Campus Noise
    const campusNoise = ["桂林洋", "府城", "龙昆南", "校区"];
    campusNoise.forEach(noise => {
        s = s.replace(new RegExp(noise + "(校区)?", 'g'), "");
    });
    s = s.replace(/校区[：:]?/g, "");

    // 3. Remove ALL whitespace to ensure clean parsing
    s = s.replace(/\s+/g, "");

    // 4. Split Building and Room
    const buildingSuffixes = "楼|教|馆|室|厅|部|大楼|场|苑|中心|程|基地";

    const candidates = [];
    const pushCandidates = (re, kind, baseScore) => {
        let m;
        re.lastIndex = 0;
        while ((m = re.exec(s)) !== null) {
            const v = m[1];
            if (!v) continue;
            const idx = m.index;
            if (v.length > 10) continue;
            if (/^\d+$/.test(v) && v.length < 3) continue;
            let score = baseScore;
            if (/^[A-Za-z]/.test(v)) score += 3;
            if (/\d{3,4}$/.test(v)) score += 1;
            if (/\d{2}[\u4e00-\u9fa5]/.test(s.slice(idx + v.length, idx + v.length + 3))) score += 4;
            candidates.push({ idx, v, kind, score });
        }
    };

    pushCandidates(/([A-Za-z]{1,3}\d{2,4})(?=\d{2}[\u4e00-\u9fa5])/g, 'alphaNum_yearMajor', 30);
    pushCandidates(/(\d{3,4})(?=\d{2}[\u4e00-\u9fa5])/g, 'num_yearMajor', 24);
    pushCandidates(/([A-Za-z]{1,3}\d{2,4})(?!\d)/g, 'alphaNum', 18);
    pushCandidates(/(\d{3,4})(?!\d)/g, 'num', 14);
    pushCandidates(/(\d{1,4}[A-Za-z]{1,2})(?=\D|$)/g, 'numAlpha', 12);

    let best = null;
    for (const c of candidates) {
        if (!best) {
            best = c;
            continue;
        }
        if (c.score > best.score) best = c;
        else if (c.score === best.score && c.idx < best.idx) best = c;
    }

    let building = "";
    let room = "";
    let truncatedSuffix = "";

    if (best) {
        room = best.v;
        const roomEndIdx = best.idx + room.length;
        building = s.substring(0, best.idx);
        truncatedSuffix = s.substring(roomEndIdx);

        if (truncatedSuffix && /^\d{2}[\u4e00-\u9fa5]/.test(truncatedSuffix)) {
            if (typeof window !== 'undefined' && window.__SCHEDULELLM_DEBUG_PARSE) {
                console.warn('[LocationTruncateAfterRoom]', { input: loc, raw: s, building, room, truncatedSuffix });
            }
        }
    } else {
        building = s;
    }

    // Further clean building
    // If building ended with "楼" or similar, it's good.
    // If building is empty but room exists? (e.g. "101") -> Building unknown.
    
    const buildingRoom = building + room;
    let fullLocation = buildingRoom || "待通知";

    if (building && room && building.endsWith(room)) {
        if (typeof window !== 'undefined' && window.__SCHEDULELLM_DEBUG_PARSE) {
            console.warn("[LocationDupStandardize]", { input: loc, building, room, full: buildingRoom });
        }
        fullLocation = building;
    }

    return {
        location: fullLocation,
        building: building,
        room: room,
        _truncated: truncatedSuffix
    };
}

function mergeBuildingRoom(building, room) {
    let b = String(building || "").replace(/\s+/g, "");
    let r = String(room || "").replace(/\s+/g, "");

    if (!b || !r) return b + r;

    r = r.replace(/^([A-Za-z])\1(\d)/, "$1$2");

    const m = r.match(/^([A-Za-z])\d/);
    if (m && b.endsWith(m[1])) {
        b = b.slice(0, -1);
    }

    if (b && r && b.endsWith(r)) {
        const merged = b;
        if (typeof window !== 'undefined' && window.__SCHEDULELLM_DEBUG_PARSE) {
            console.warn("[LocationDupMerge]", { building: b, room: r, merged });
        }
        return merged;
    }

    return b + r;
}

function simplifyLocation(loc) {
    // Legacy wrapper for compatibility
    return standardizeLocation(loc).location;
}

function parseWeekString(str) {
    // Example: "(1-2节)2-6周,8-12周(双)"
    // Or just "2-6周"
    // Normalize first
    let cleanStr = normalizeOCRText(str);
    
    // Remove anything inside parens that looks like period "1-2节"
    cleanStr = cleanStr.replace(/\([^)]*节\)/g, "");

    // Logic: Split by comma
    const parts = cleanStr.split(/[,，]/); // Handle Chinese comma too
    let weekSet = new Set();

    parts.forEach(part => {
        // Match patterns: "2-6周", "8-12周(双)", "5周"
        // Also support missing "周" if it's clearly a range like "1-16"
        // Scan all candidates inside the same part, because part may contain course codes like (43011091)
        const weekRe = /(\d+)(?:-(\d+))?(?:周|W|w)?(?:\((单|双)\))?/g;
        let match;

        while ((match = weekRe.exec(part)) !== null) {
            if (!match[0]) continue;

            // If it doesn't explicitly look like week info (no 周/W/w and no range), skip it.
            // This prevents course codes / counts from blocking later valid week ranges.
            const token = match[0];
            const hasWeekMark = /[周Ww]/.test(token);
            const hasRange = !!match[2];
            if (!hasWeekMark && !hasRange) continue;

            const start = parseInt(match[1], 10);
            const end = match[2] ? parseInt(match[2], 10) : start;
            const type = match[3];

            // Sanity check: weeks shouldn't be > 30 usually
            // IMPORTANT: don't return; just skip this candidate and keep scanning.
            if (!Number.isFinite(start) || start <= 0 || start > 30) continue;
            if (!Number.isFinite(end) || end <= 0 || end > 30) continue;

            for (let i = start; i <= end; i++) {
                if (type === '单' && i % 2 === 0) continue;
                if (type === '双' && i % 2 !== 0) continue;
                weekSet.add(i);
            }
        }
    });

    return Array.from(weekSet).sort((a, b) => a - b);
}

function simplifyName(name) {
    if (!name) return "";

    // 1. Find the first occurrence of balanced brackets (English or Chinese)
    // We want to keep everything from the start until the end of the FIRST bracketed pair.
    const match = name.match(/^(.*?[[\(（][^()（）]*[\)）])/);

    let s = name;
    if (match && match[1]) {
        s = match[1];
    }

    // 2. Clean up trailing spaces or non-word delimiters
    s = s.replace(/[\s\-_/]+$/, "");

    // 3. Remove ALL whitespace for standardization
    return s.replace(/\s+/g, "");
}

function normalizeOCRText(str) {
    if (!str) return "";
    return str
        .replace(/[０-９]/g, d => String.fromCharCode(d.charCodeAt(0) - 65248))
        .replace(/[Ａ-Ｚａ-ｚ]/g, s => String.fromCharCode(s.charCodeAt(0) - 65248))
        .replace(/（/g, "(").replace(/）/g, ")")
        .replace(/：/g, ":")
        .replace(/—/g, "-")
        .replace(/－/g, "-")
        .replace(/(\d+)\s*[\n\r]*[-~～]\s*[\n\r]*\s*(\d+)/g, "$1-$2")
        .replace(/(\d+)\s*[\n\r]+\s*(\d+)/g, "$1$2")
        .replace(/～/g, "-")
        .trim();
}

function canonicalizeLLMCellKey(cell) {
    const s = normalizeOCRText(String(cell || "").trim());
    return s.replace(/◇/g, ' / ').replace(/[:：;；]/g, '/');
}

const SCHEDULELLM_LOG_STORAGE_KEY = "schedulellm_logs_v1";
const SCHEDULELLM_LOG_MAX = 500;

function scheduleLLMHash(str) {
    const s = String(str || "");
    let h = 2166136261;
    for (let i = 0; i < s.length; i++) {
        h ^= s.charCodeAt(i);
        h = Math.imul(h, 16777619);
    }
    return (h >>> 0).toString(16);
}

function scheduleLLMGetLogs() {
    try {
        const raw = localStorage.getItem(SCHEDULELLM_LOG_STORAGE_KEY);
        const arr = raw ? JSON.parse(raw) : [];
        return Array.isArray(arr) ? arr : [];
    } catch (_) {
        return [];
    }
}

function scheduleLLMSetLogs(logs) {
    try {
        localStorage.setItem(SCHEDULELLM_LOG_STORAGE_KEY, JSON.stringify(logs));
    } catch (_) {
    }
}

function scheduleLLMLog(entry) {
    const safe = entry && typeof entry === 'object' ? { ...entry } : { msg: String(entry) };
    if (safe.apiKey) delete safe.apiKey;
    safe.ts = safe.ts || new Date().toISOString();
    const logs = scheduleLLMGetLogs();
    logs.push(safe);
    if (logs.length > SCHEDULELLM_LOG_MAX) logs.splice(0, logs.length - SCHEDULELLM_LOG_MAX);
    scheduleLLMSetLogs(logs);
}

function scheduleLLMSummarizeLogs() {
    const logs = scheduleLLMGetLogs();
    const byType = {};
    const byReason = {};
    for (const l of logs) {
        const t = l.type || "unknown";
        byType[t] = (byType[t] || 0) + 1;
        if (l.reason) byReason[l.reason] = (byReason[l.reason] || 0) + 1;
    }
    return { total: logs.length, byType, byReason };
}

function scheduleLLMExportLogs() {
    const logs = scheduleLLMGetLogs();
    const blob = new Blob([JSON.stringify({ exportedAt: new Date().toISOString(), logs }, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `schedulellm-logs-${Date.now()}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
}

function scheduleLLMClearLogs() {
    try { localStorage.removeItem(SCHEDULELLM_LOG_STORAGE_KEY); } catch (_) {}
}

if (typeof window !== 'undefined') {
    window.scheduleLLMGetLogs = scheduleLLMGetLogs;
    window.scheduleLLMSummarizeLogs = scheduleLLMSummarizeLogs;
    window.scheduleLLMExportLogs = scheduleLLMExportLogs;
    window.scheduleLLMClearLogs = scheduleLLMClearLogs;
}

let scheduleLLMRecognizedItems = [];

function scheduleLLMProgressEls() {
    const host = document.getElementById('llmProgressHost');
    if (!host) return null;
    return {
        host,
        fill: document.getElementById('llmProgressFill'),
        pct: document.getElementById('llmProgressPct'),
        count: document.getElementById('llmProgressCount'),
        text: document.getElementById('llmProgressText'),
        icon: document.getElementById('llmProgressIcon'),
        retry: document.getElementById('llmProgressRetry'),
        err: document.getElementById('llmProgressError'),
        errText: document.getElementById('llmProgressErrorText'),
        errRetry: document.getElementById('llmProgressErrorRetry'),
        panel: document.getElementById('llmRecognizedPanel'),
        summary: document.getElementById('llmRecognizedSummary'),
        list: document.getElementById('llmRecognizedList'),
        detail: document.getElementById('llmCourseDetail')
    };
}

function scheduleLLMProgressSetVisible(visible) {
    const els = scheduleLLMProgressEls();
    if (!els) return;
    els.host.style.display = visible ? 'block' : 'none';
}

function scheduleLLMProgressSetRunning(running) {
    const els = scheduleLLMProgressEls();
    if (!els) return;
    els.host.classList.toggle('running', !!running);
    els.host.classList.toggle('done', false);
}

function scheduleLLMProgressSetIcon(kind, text) {
    const els = scheduleLLMProgressEls();
    if (!els || !els.icon) return;
    if (!kind) {
        els.icon.style.display = 'none';
        els.icon.className = 'llm-progress-icon';
        els.icon.textContent = '';
        return;
    }
    els.icon.style.display = 'inline-flex';
    els.icon.className = `llm-progress-icon ${kind}`;
    els.icon.textContent = text || (kind === 'done' ? '✓' : '!');
}

function scheduleLLMProgressSetText(text) {
    const els = scheduleLLMProgressEls();
    if (!els || !els.text) return;
    els.text.textContent = String(text || '');
}

function scheduleLLMProgressSetProgress(processed, total, extractedCourses) {
    const els = scheduleLLMProgressEls();
    if (!els) return;
    const t = Math.max(0, parseInt(total || 0, 10));
    const p = Math.max(0, Math.min(t || 0, parseInt(processed || 0, 10)));
    const pct = t > 0 ? Math.round((p / t) * 100) : 0;
    if (els.fill) els.fill.style.width = `${pct}%`;
    if (els.pct) els.pct.textContent = `${pct}%`;
    if (els.count) {
        const courses = Math.max(0, parseInt(extractedCourses || 0, 10));
        els.count.textContent = `${p}/${t} · ${courses}门课`;
    }
}

function scheduleLLMProgressReset() {
    const els = scheduleLLMProgressEls();
    if (!els) return;
    scheduleLLMRecognizedItems = [];
    if (els.list) els.list.innerHTML = '';
    if (els.detail) {
        els.detail.innerHTML = '';
        els.detail.style.display = 'none';
    }
    if (els.summary) els.summary.textContent = '0';
    if (els.err) els.err.style.display = 'none';
    scheduleLLMProgressSetIcon(null);
    scheduleLLMProgressSetText('准备中');
    scheduleLLMProgressSetProgress(0, 0, 0);
    els.host.classList.toggle('done', false);
}

function scheduleLLMProgressShowError(message) {
    const els = scheduleLLMProgressEls();
    if (!els || !els.err || !els.errText) return;
    els.errText.textContent = String(message || '识别失败');
    els.err.style.display = 'flex';
    scheduleLLMProgressSetIcon('err', '!');
}

function scheduleLLMProgressHideError() {
    const els = scheduleLLMProgressEls();
    if (!els || !els.err) return;
    els.err.style.display = 'none';
}

function scheduleLLMResetLayoutMode() {
    const els = scheduleLLMProgressEls();
    if (!els) return;

    // Restore Panel to Host if needed (revert move)
    if (els.panel && els.host && els.panel.parentElement !== els.host) {
        els.host.appendChild(els.panel);
        els.panel.classList.remove('moved-below-calendar');
    }

    // Restore Progress Visibility
    els.host.classList.remove('fade-out');
    els.host.style.display = '';
    els.host.style.opacity = '';

    els.host.classList.remove('sidebar-mode', 'drawer-mode', 'post-mode');
    const card = els.host.querySelector('.llm-progress-card');
    if (card) card.style.display = '';
    if (els.panel) els.panel.open = true;
}

function scheduleLLMEnterPostMode(enabled) {
    const els = scheduleLLMProgressEls();
    if (!els) return;
    scheduleLLMResetLayoutMode();
    if (!enabled) return;
    els.host.classList.add('post-mode');
    const isMobile = window.innerWidth <= 768;
    if (isMobile) {
        els.host.classList.add('drawer-mode');
        if (els.panel) els.panel.open = false;
    } else {
        els.host.classList.add('sidebar-mode');
        if (els.panel) els.panel.open = true;
    }
    const card = els.host.querySelector('.llm-progress-card');
    if (card) card.style.display = 'none';
}

function scheduleLLMOnCalendarRendered(useLLM) {
    const els = scheduleLLMProgressEls();
    if (!els) return;
    if (!useLLM) {
        scheduleLLMProgressSetVisible(false);
        return;
    }
    scheduleLLMResetLayoutMode();
}

function scheduleLLMFormatCourseTime(c) {
    const pr = c && (c.periodRange || c.period);
    const wk = c && (c.raw_weeks || c.weeksRaw);
    const parts = [];
    if (pr) parts.push(`${String(pr).replace(/节/g, '')}节`);
    if (wk) parts.push(String(wk));
    return parts.length ? parts.join(' ') : '—';
}

function scheduleLLMProgressAddCourses(courses, source) {
    const els = scheduleLLMProgressEls();
    if (!els || !els.list) return;
    const arr = Array.isArray(courses) ? courses : [];
    arr.forEach(c => {
        const name = c && (c.name || c.displayName || c.rawName) ? String(c.name || c.displayName || c.rawName) : '未命名课程';
        const time = scheduleLLMFormatCourseTime(c);
        const item = {
            name,
            time,
            teacher: c && c.teacher ? String(c.teacher) : '',
            className: c && c.className ? String(c.className) : '',
            location: c && c.location ? String(c.location) : '',
            building: c && c.building ? String(c.building) : '',
            room: c && c.room ? String(c.room) : '',
            raw_weeks: c && c.raw_weeks ? String(c.raw_weeks) : (c && c.weeksRaw ? String(c.weeksRaw) : ''),
            periodRange: c && c.periodRange ? String(c.periodRange) : '',
            source: source ? String(source) : ''
        };
        const idx = scheduleLLMRecognizedItems.push(item) - 1;
        const li = document.createElement('li');
        li.innerHTML = `<button type="button" class="llm-course-item" data-idx="${idx}"><span class="llm-course-item-name"></span><span class="llm-course-item-time"></span></button>`;
        const btn = li.querySelector('button');
        btn.querySelector('.llm-course-item-name').textContent = name;
        btn.querySelector('.llm-course-item-time').textContent = time;
        els.list.appendChild(li);
    });
    if (els.summary) els.summary.textContent = String(scheduleLLMRecognizedItems.length);
}

function scheduleLLMProgressShowDetail(idx) {
    const els = scheduleLLMProgressEls();
    if (!els || !els.detail) return;
    const item = scheduleLLMRecognizedItems[idx];
    if (!item) return;

    const loc = item.location || ((item.building || item.room) ? `${item.building || ''}${item.room || ''}` : '');

    els.detail.innerHTML = `
        <div class="llm-course-detail-title">${item.name}</div>
        <div class="llm-course-detail-grid">
            <strong>时间</strong><span>${item.time}</span>
            <strong>地点</strong><span>${loc || '—'}</span>
            <strong>教师</strong><span>${item.teacher || '—'}</span>
            <strong>班级</strong><span>${item.className || '—'}</span>
            <strong>来源</strong><span>${item.source || '—'}</span>
        </div>
    `;
    els.detail.style.display = 'block';
}

function scheduleLLMResetLayoutMode() {
    const els = scheduleLLMProgressEls();
    if (!els) return;
    els.host.classList.remove('sidebar-mode', 'drawer-mode', 'post-mode', 'fade-out');
    els.host.style.display = '';
    
    if (els.panel && els.panel.parentElement !== els.host) {
        const card = els.host.querySelector('.llm-progress-card');
        if (card) {
            card.insertAdjacentElement('afterend', els.panel);
        } else {
            els.host.appendChild(els.panel);
        }
        els.panel.classList.remove('moved-below-calendar');
    }

    const card = els.host.querySelector('.llm-progress-card');
    if (card) {
        card.classList.remove('hidden');
        card.style.display = '';
    }
    if (els.panel) els.panel.open = true;
    const label = document.getElementById('llmRecognizedLabel');
    if (label) label.textContent = '已识别课程';
    const main = document.querySelector('.main-content');
    if (main) main.classList.remove('llm-two-col');
}

function scheduleLLMEnterPostMode(enabled) {
    const els = scheduleLLMProgressEls();
    if (!els) return;
    scheduleLLMResetLayoutMode();
    if (!enabled) return;

    els.host.classList.add('post-mode');
    const isMobile = window.innerWidth <= 768;
    if (isMobile) {
        els.host.classList.add('drawer-mode');
        const label = document.getElementById('llmRecognizedLabel');
        if (label) label.textContent = '查看课程';
    } else {
        els.host.classList.add('sidebar-mode');
        const label = document.getElementById('llmRecognizedLabel');
        if (label) label.textContent = '课程';
        const main = document.querySelector('.main-content');
        if (main) main.classList.add('llm-two-col');
    }

    if (els.panel) els.panel.open = false;

    const card = els.host.querySelector('.llm-progress-card');
    if (card) {
        card.classList.add('hidden');
        window.setTimeout(() => {
            card.style.display = 'none';
        }, 220);
    }
}

function scheduleLLMOnCalendarRendered(useLLM) {
    const els = scheduleLLMProgressEls();
    if (!els) return;
    if (!useLLM) {
        scheduleLLMProgressSetVisible(false);
        return;
    }
    scheduleLLMResetLayoutMode();

    // Move Panel below calendar
    const calendarArea = document.getElementById('calendarArea');
    if (els.panel && calendarArea) {
        calendarArea.insertAdjacentElement('afterend', els.panel);
        els.panel.classList.add('moved-below-calendar');
        els.panel.open = true;
    }

    // Fade out progress host
    if (els.host) {
        els.host.classList.add('fade-out');
        setTimeout(() => {
            if (els.host.classList.contains('fade-out')) {
                els.host.style.display = 'none';
            }
        }, 300);
    }

    window.requestAnimationFrame(() => {
        window.dispatchEvent(new Event('resize'));
    });
}

function scheduleLLMHideHintOnGenerate() {
    const el = document.getElementById('llmHint');
    if (!el) return;
    if (el.style.display === 'none') return;
    if (el.classList.contains('llm-hint-fade-out')) return;

    const reduce = typeof window !== 'undefined' && window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (reduce) {
        el.style.display = 'none';
        return;
    }

    el.addEventListener('animationend', () => {
        el.style.display = 'none';
        el.classList.remove('llm-hint-fade-out');
    }, { once: true });

    el.classList.add('llm-hint-fade-out');
}

let generatedEvents = [];
let currentCalendarDate = new Date();
let scheduleLLMMinMonth = null;
let scheduleLLMMaxMonth = null;

async function generateSchedule() {
    try {
        console.log("Starting generation...");
        if (rawScheduleData.length === 0) {
            alert("请先上传课表文件");
            return;
        }

        scheduleLLMHideHintOnGenerate();

        const btnGen = document.getElementById('btnGenerate');
        const originalBtnText = btnGen.textContent;
        let useLLM = document.getElementById('useLLM').checked; // Changed from const to let
        const llmCache = new Map(); // Cache for unique cell results

        // Regex patterns to identify non-course cells (headers, metadata, etc.)
        // Defined here to be used both for LLM filtering and main loop skipping
        const ignorePatterns = [
            /(星期|周)[\s\n]*[一二三四五六日天]/, // Days: 星期一, 周一 (Allow whitespace)
            // Periods: 1-2节, 第九+节, 第-二节 (Allow hyphen after 第)
            /第\s*[-]*\s*[一二三四五六七八九十\d]+\s*[-~+～至,\s]*\s*[一二三四五六七八九十\d]*\s*节/, 
            /学年|学期|课表|教工号|打印时间|注一|内容顺序/, // Titles & Metadata
            /^(上|下|晚|早|午)[\s\n]*(午|晚|晨|间|上)$/, // Time of day: 上午, 晚上, 早上 (Updated to include '上' in second group)
            /^节次$/ // Header label
        ];

        if (typeof scheduleLLMResetLayoutMode === 'function') {
            scheduleLLMResetLayoutMode();
        }

        // 未启用 LLM：隐藏右侧进度提示（如果存在）
        if (!useLLM) {
            if (typeof scheduleLLMProgressSetVisible === 'function') {
                scheduleLLMProgressSetVisible(false);
            }
        }

        if (useLLM) {
            const config = {
                baseUrl: document.getElementById('llmBaseUrl').value,
                apiKey: document.getElementById('llmApiKey').value,
                model: document.getElementById('llmModel').value
            };

            const logCtx = {
                baseUrl: config.baseUrl,
                model: config.model,
                ua: (typeof navigator !== 'undefined' ? navigator.userAgent : undefined)
            };

            scheduleLLMLog({ type: 'run_start', ...logCtx, rawRows: rawScheduleData.length });

            const isProxy = /\/api\/llm\/?$/.test((config.baseUrl || '').trim());
            if (!isProxy && !config.apiKey) {
                alert("直连模式需要填写 API Key；生产环境建议使用后端 /api/llm 代理");
                return;
            }

            // Collect all unique non-empty cells
            const uniqueCells = new Set();
            
            rawScheduleData.forEach(row => {
                row.forEach(cell => {
                    if (cell && typeof cell === 'string' && cell.trim()) {
                        // Apply normalizeOCRText BEFORE sending to LLM or Regex
                        // This ensures LLM sees "4-15" instead of "4\n15"
                        const val = canonicalizeLLMCellKey(cell);
                        
                        // Skip if matches any ignore pattern
                        if (ignorePatterns.some(p => p.test(val))) {
                            return;
                        }
                        uniqueCells.add(val);
                    }
                });
            });

            btnGen.disabled = true;
            btnGen.textContent = "LLM 语义识别中...";

            scheduleLLMProgressSetVisible(true);
            scheduleLLMProgressReset();
            scheduleLLMProgressSetRunning(true);
            scheduleLLMProgressHideError();
            scheduleLLMProgressSetText('准备识别…');

            const cellsToProcess = Array.from(uniqueCells);
            let processedCells = 0;
            let extractedCourses = 0;
            let hadException = false;
            let slowTimer = null;

            scheduleLLMProgressSetProgress(0, cellsToProcess.length, 0);

            const service = window.llmService || (typeof llmService !== 'undefined' ? llmService : null);
            if (!service) {
                console.error("LLM Service not found in window or global scope. Check llm_parser.js loading.");
                scheduleLLMProgressShowError("LLM组件加载失败，已降级使用普通解析");
                alert("LLM组件加载失败，无法使用智能识别功能。将降级使用普通解析。");
                useLLM = false;
                scheduleLLMProgressSetRunning(false);
            } else {
                service.updateConfig(config.baseUrl, config.apiKey, config.model);

                for (let i = 0; i < cellsToProcess.length; i++) {
                    const cell = cellsToProcess[i];
                    btnGen.textContent = `识别中 (${i + 1}/${cellsToProcess.length})`;
                    scheduleLLMProgressSetText(`识别中 ${i + 1}/${cellsToProcess.length}`);

                    if (slowTimer) clearTimeout(slowTimer);
                    slowTimer = setTimeout(() => {
                        scheduleLLMProgressSetIcon('warn', '!');
                        scheduleLLMProgressSetText('正在努力识别中...');
                    }, 3000);

                    try {
                        const result = await service.parseCourse(cell);
                        if (slowTimer) {
                            clearTimeout(slowTimer);
                            slowTimer = null;
                        }
                        scheduleLLMProgressSetIcon(null);

                        processedCells++;

                        if (result && !result.error && result.courses && result.courses.length > 0) {
                            llmCache.set(cell, result.courses);
                            extractedCourses += result.courses.length;
                            scheduleLLMProgressAddCourses(result.courses, 'LLM');
                            scheduleLLMLog({
                                type: 'llm_success',
                                ...logCtx,
                                cellHash: scheduleLLMHash(cell),
                                cellLen: String(cell).length,
                                courses: result.courses.length
                            });
                        } else {
                            const reason = result ? (result.error || "Empty courses array") : "Null result";
                            const cleanCell = cell.replace(/\n/g, '\\n');
                            console.warn(`[LLM Failure] Cell: "${cleanCell}" - Reason: ${reason}. System will attempt Regex fallback.`);

                            const regexFallback = parseCourseString(cell);
                            const fbCount = regexFallback && regexFallback.length ? regexFallback.length : 0;
                            extractedCourses += fbCount;
                            if (fbCount > 0) scheduleLLMProgressAddCourses(regexFallback, '正则');

                            scheduleLLMLog({
                                type: 'llm_failure',
                                ...logCtx,
                                reason,
                                cellHash: scheduleLLMHash(cell),
                                cellLen: String(cell).length,
                                llmCourses: (result && result.courses && result.courses.length) ? result.courses.length : 0,
                                regexFallback: fbCount
                            });

                            if (fbCount > 0) {
                                console.info(`[Fallback Success] Regex parser successfully identified ${fbCount} courses from "${cleanCell}".`);
                            } else {
                                console.warn(`[Parsing Warning] Both LLM and Regex failed to extract content from: "${cleanCell}". This may be a header or unrecognized format.`);
                            }
                        }

                        scheduleLLMProgressSetProgress(processedCells, cellsToProcess.length, extractedCourses);

                    } catch (err) {
                        if (slowTimer) {
                            clearTimeout(slowTimer);
                            slowTimer = null;
                        }

                        processedCells++;
                        hadException = true;
                        scheduleLLMProgressSetIcon('err', '!');
                        scheduleLLMProgressShowError((err && err.message) ? `识别出错：${err.message}` : '识别出错');

                        const cleanCell = cell.replace(/\n/g, '\\n');
                        console.error(`[LLM Exception] Error processing cell: "${cleanCell}"`, err);

                        const regexFallback = parseCourseString(cell);
                        const fbCount = regexFallback && regexFallback.length ? regexFallback.length : 0;
                        extractedCourses += fbCount;
                        if (fbCount > 0) scheduleLLMProgressAddCourses(regexFallback, '正则');

                        scheduleLLMLog({
                            type: 'llm_exception',
                            ...logCtx,
                            reason: (err && err.name ? err.name : 'Error') + (err && err.message ? `: ${err.message}` : ''),
                            cellHash: scheduleLLMHash(cell),
                            cellLen: String(cell).length,
                            regexFallback: fbCount
                        });

                        if (fbCount > 0) {
                            console.info(`[Fallback Success] Regex parser successfully identified ${fbCount} courses from "${cleanCell}".`);
                        } else {
                            console.warn(`[Parsing Warning] Regex fallback also returned no results for: "${cleanCell}".`);
                        }

                        scheduleLLMProgressSetProgress(processedCells, cellsToProcess.length, extractedCourses);
                    }
                }

                scheduleLLMProgressSetRunning(false);
                if (!hadException) {
                    const els = scheduleLLMProgressEls();
                    if (els) els.host.classList.toggle('done', true);
                    scheduleLLMProgressSetIcon('done', '✓');
                    scheduleLLMProgressHideError();
                    scheduleLLMProgressSetText('识别完成');
                } else {
                    scheduleLLMProgressSetText('识别完成（存在错误，可重试）');
                }
            }

            btnGen.textContent = originalBtnText;
            btnGen.disabled = false;
        }

        const startDateInput = document.getElementById('semesterStart').value;
        if (!startDateInput) return;
        const semesterStart = new Date(startDateInput);

        // Find Header Row
        let headerRowIdx = -1;
        for (let r = 0; r < rawScheduleData.length; r++) {
            const row = rawScheduleData[r];
            // Support "星期一" or "周一"
            if (row.some(c => c && typeof c === 'string' && /(星期|周)一/.test(c))) {
                headerRowIdx = r;
                break;
            }
        }

        if (headerRowIdx === -1) {
            alert("未识别到'星期一'或'周一'表头，请检查文件格式");
            return;
        }

        const headerRow = rawScheduleData[headerRowIdx];
        const colToDayIdx = {}; // col -> 1(Mon)..7(Sun)
        headerRow.forEach((cell, idx) => {
            if (!cell || typeof cell !== 'string') return;
            if (/(星期|周)一/.test(cell)) colToDayIdx[idx] = 1;
            if (/(星期|周)二/.test(cell)) colToDayIdx[idx] = 2;
            if (/(星期|周)三/.test(cell)) colToDayIdx[idx] = 3;
            if (/(星期|周)四/.test(cell)) colToDayIdx[idx] = 4;
            if (/(星期|周)五/.test(cell)) colToDayIdx[idx] = 5;
            if (/(星期|周)六/.test(cell)) colToDayIdx[idx] = 6;
            if (/(星期|周)(日|天)/.test(cell)) colToDayIdx[idx] = 7;
        });

        console.log("Day Map:", colToDayIdx);
        // Iterate rows below header
        const events = [];
        const weeklessBuffer = []; // Buffer for courses without specific weeks

        // Read current time settings from UI
        // Fix: Removed unused timeInputs variable
        const currentSlots = [];
        for (let i = 0; i < defaultTimeSlots.length; i++) {
            const startInput = document.querySelector(`input[data-idx="${i}"][data-type="start"]`);
            const endInput = document.querySelector(`input[data-idx="${i}"][data-type="end"]`);
            if (startInput && endInput) {
                currentSlots.push({ start: startInput.value, end: endInput.value });
            } else {
                currentSlots.push(defaultTimeSlots[i]);
            }
        }

        for (let r = headerRowIdx + 1; r < rawScheduleData.length; r++) {
            const row = rawScheduleData[r];
            if (!row || row.length === 0) continue;

            let periodNum = -1;

            // Helper to parse "第一节", "二", "3", etc.
            const parsePeriodCell = (cell) => {
                if (!cell) return -1;
                const s = String(cell).trim();

                // 1. Check for standard digits
                const digitMatch = s.match(/^(\d+)/);
                if (digitMatch) return parseInt(digitMatch[1]);

                // 2. Check for Chinese numerals
                const cnNums = {
                    '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
                    '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
                    '十一': 11, '十二': 12
                };

                // Look for any key in string
                for (const [k, v] of Object.entries(cnNums)) {
                    if (s.includes(k)) return v;
                }

                return -1;
            };
            
            // Try to find period number from the first few columns
            // Heuristic: Check first 3 columns for a valid period number
            periodNum = -1;
            for (let c = 0; c < Math.min(row.length, 3); c++) {
                const p = parsePeriodCell(row[c]);
                if (p !== -1) {
                    periodNum = p;
                    break; // Found it
                }
            }

            // Also check if the row is purely metadata (like "Lunch Break")
            // If periodNum is still -1, verify if we should skip
            if (periodNum === -1 || periodNum > 12) {
                // Debug log for skipped rows if needed
                // console.log(`Skipping row ${r} due to invalid period:`, row);
                continue;
            }

            const timeSlot = currentSlots[periodNum - 1];
            if (!timeSlot) continue;

            // Iterate Columns
            for (const [colIdx, dayIdx] of Object.entries(colToDayIdx)) {
                const cellContent = row[colIdx];
                if (!cellContent || typeof cellContent !== 'string' || !cellContent.trim()) continue;

                // Skip if matches any ignore pattern (Headers, Time slots, etc.)
                if (ignorePatterns.some(p => p.test(cellContent.trim()))) {
                    continue;
                }

                // 1. Parse Courses from Cell
                let courses = [];
                
                const cacheKey = canonicalizeLLMCellKey(cellContent);

                if (useLLM && llmCache.has(cacheKey)) {
                    const cachedCourses = llmCache.get(cacheKey);
                    // Re-hydrate objects (ensure structure)
                    courses = cachedCourses.map(c => {
                        const locSeed = (c && c.building && c.room) ? mergeBuildingRoom(c.building, c.room) : (c && c.location ? String(c.location) : "");
                        const locInfo = standardizeLocation(locSeed);

                        if (typeof window !== 'undefined' && window.__SCHEDULELLM_DEBUG_PARSE && /([A-Za-z])\1\d/.test(locInfo.room || "")) {
                            console.warn("[LocationDupLetter]", { locSeed, locInfo, llm: c });
                        }

                        return {
                            rawName: c.name,
                            displayName: simplifyName(c.name),
                            weeks: Array.isArray(c.weeks) ? c.weeks : parseWeekString(c.raw_weeks || c.weeks),
                            location: locInfo.location || "待通知",
                            className: c.className ? c.className.replace(/^[\(（]/, '').replace(/[\)）]$/, '') : "",
                            periodRange: c.periodRange || "",
                            rawStr: cellContent,
                            building: locInfo.building,
                            room: locInfo.room
                        };
                    });
                } else {
                    courses = parseCourseString(cellContent);
                }

                // 2. Generate Events
                courses.forEach(course => {
                    // [DEBUG] Logging for Week Extraction Diagnosis
                    // Logs detailed info for all courses, with specific focus on Higher Math as requested
                    const debugDayName = ['?', '一', '二', '三', '四', '五', '六', '日'][dayIdx] || dayIdx;
                    const isTarget = course.displayName.includes("高等数学");
                    // Use a distinctive prefix for easy filtering
                    const logPrefix = isTarget ? "[DEBUG-TARGET]" : "[DEBUG]"; 
                    
                    console.log(`${logPrefix}周次识别结果：
  课程: ${course.displayName}
  时间: 周${debugDayName} (DayIdx: ${dayIdx})
  原始文本: "${(course.rawStr || cellContent || "").replace(/\n/g, '\\n')}"
  识别周次: [${course.weeks.join(', ')}]
  识别班级: ${course.className || "无"}
  识别地点: ${course.location || "无"}
  来源: ${useLLM && llmCache.has(cacheKey) ? "LLM缓存" : "正则解析"}`);

                    const courseWeeks = course.weeks;
                    
                    if (courseWeeks.length === 0) {
                         // No weeks? Buffer it? 
                         // Or just add to "All Semester" (bad idea).
                         // Let's warn and skip for now, or add to a "Fix Me" list.
                         // But for simplicity, we skip generation but maybe show error.
                         return;
                    }

                    courseWeeks.forEach(weekNum => {
                        // Calculate Date
                        // Date = StartDate + (Week-1)*7 days + (DayIdx-1) days
                        const daysToAdd = (weekNum - 1) * 7 + (dayIdx - 1);
                        const targetDate = new Date(semesterStart);
                        targetDate.setDate(semesterStart.getDate() + daysToAdd);

                        // Determine Time of Day for Color Coding
                        // Period 1-4: Morning
                        // Period 5-8: Afternoon
                        // Period 9+: Evening
                        let timeOfDay = 'morning';
                        if (periodNum >= 5 && periodNum <= 8) timeOfDay = 'afternoon';
                        if (periodNum >= 9) timeOfDay = 'evening';

                        events.push({
                            title: course.displayName,
                            rawTitle: course.rawName,
                            location: course.location,
                            className: course.className,
                            weeks: course.weeks, // Pass all weeks
                            periodRange: course.periodRange, // Pass period info
                            startTime: timeSlot.start, // HH:mm
                            endTime: timeSlot.end,
                            date: targetDate, // Date Object
                            week: weekNum,
                            period: periodNum,
                            dayOfWeek: dayIdx,
                            timeOfDay: timeOfDay,
                            description: `课程: ${course.rawName}\n地点: ${course.location}\n周次: ${weekNum}周\n班级: ${course.className}`
                        });
                    });
                });
            }
        }

        generatedEvents = events;
        console.log("Events generated:", events.length);

        if (events.length === 0) {
            // Diagnostic Alert
            let msg = "未生成任何日程。\n诊断信息：\n";
            msg += `1. 读取总行数: ${rawScheduleData.length}\n`;
            msg += `2. 表头行索引: ${headerRowIdx} (列数: ${Object.keys(colToDayIdx).length})\n`;
            const effectiveCourseRowCount = Math.max(0, rawScheduleData.length - headerRowIdx - 1);
            msg += `3. 有效课程行数(估算): ${effectiveCourseRowCount}\n`;
            msg += "可能原因：\n- 无法识别节次列（请确保第一列或前几列包含'1','2','一','二'等数字）\n- 课程单元格为空或无法解析\n- 课程周次格式不标准";
            alert(msg);
        }

        renderCalendar(events);

        if (typeof scheduleLLMOnCalendarRendered === 'function') {
            scheduleLLMOnCalendarRendered(useLLM && events.length > 0);
        }
        
        // Restore button state (It was modified at start of generateSchedule)
        // btnGen is already defined in the outer scope of generateSchedule, reuse it.
        btnGen.textContent = "生成月历";
        btnGen.disabled = false;

    } catch (e) {
        console.error(e);
        alert("生成月历出错: " + e.message);
        
        // Restore button state on error too
        const btnGenRetry = document.getElementById('btnGenerate'); // Use different name if strictly needed, but reusing outer var is better if scope allows. 
        // Actually, btnGen is defined at top of function. We can just use it if we are in same scope.
        // But 'catch' block has its own scope? No, 'btnGen' from top of function is accessible in catch.
        // BUT, the previous code re-declared it with 'const btnGen = ...' inside the try block or if block?
        // Let's check where it was defined.
        // It was defined at: const btnGen = document.getElementById('btnGenerate'); at line 418 (approx).
        
        // So we should NOT redeclare it.
        if (typeof btnGen !== 'undefined') {
            btnGen.textContent = "生成月历";
            btnGen.disabled = false;
        } else {
             document.getElementById('btnGenerate').textContent = "生成月历";
             document.getElementById('btnGenerate').disabled = false;
        }
    }
}

// Rendering Logic
function scheduleLLMMonthStart(d) {
    return new Date(d.getFullYear(), d.getMonth(), 1);
}

function scheduleLLMCompareMonth(a, b) {
    return (a.getFullYear() - b.getFullYear()) || (a.getMonth() - b.getMonth());
}

function scheduleLLMSetMonthRangeFromEvents(events) {
    if (!Array.isArray(events) || events.length === 0) {
        scheduleLLMMinMonth = null;
        scheduleLLMMaxMonth = null;
        return;
    }
    let minTime = Infinity;
    let maxTime = -Infinity;
    events.forEach(e => {
        const t = e && e.date ? e.date.getTime() : NaN;
        if (!Number.isFinite(t)) return;
        if (t < minTime) minTime = t;
        if (t > maxTime) maxTime = t;
    });
    if (!Number.isFinite(minTime) || !Number.isFinite(maxTime)) {
        scheduleLLMMinMonth = null;
        scheduleLLMMaxMonth = null;
        return;
    }
    scheduleLLMMinMonth = scheduleLLMMonthStart(new Date(minTime));
    scheduleLLMMaxMonth = scheduleLLMMonthStart(new Date(maxTime));
}

function scheduleLLMClampMonthToRange(d) {
    const m = scheduleLLMMonthStart(d);
    if (scheduleLLMMinMonth && scheduleLLMCompareMonth(m, scheduleLLMMinMonth) < 0) return new Date(scheduleLLMMinMonth);
    if (scheduleLLMMaxMonth && scheduleLLMCompareMonth(m, scheduleLLMMaxMonth) > 0) return new Date(scheduleLLMMaxMonth);
    return m;
}

function scheduleLLMFormatMonthTitle(d) {
    return `${d.getFullYear()}年${d.getMonth() + 1}月`;
}

function scheduleLLMCourseListEls() {
    const panel = document.getElementById('courseListPanel');
    if (!panel) return null;
    return {
        panel,
        toggle: document.getElementById('courseListToggle'),
        summary: document.getElementById('courseListSummary'),
        body: document.getElementById('courseListBody'),
        content: document.getElementById('courseListContent')
    };
}

function scheduleLLMSetCourseListVisible(visible) {
    const els = scheduleLLMCourseListEls();
    if (!els || !els.panel) return;

    const on = !!visible;
    els.panel.style.display = on ? '' : 'none';

    if (!on) {
        els.panel.classList.remove('is-open');
        if (els.toggle) els.toggle.setAttribute('aria-expanded', 'false');
    }
}

function scheduleLLMUpdateCourseListForMonth(monthDate) {
    const els = scheduleLLMCourseListEls();
    if (!els || !els.content) return;

    const y = monthDate.getFullYear();
    const m = monthDate.getMonth();

    const monthEvents = generatedEvents
        .filter(e => e && e.date && e.date.getFullYear() === y && e.date.getMonth() === m)
        .slice()
        .sort((a, b) => {
            const ta = a.date.getTime();
            const tb = b.date.getTime();
            if (ta !== tb) return ta - tb;
            if ((a.period || 0) !== (b.period || 0)) return (a.period || 0) - (b.period || 0);
            return String(a.title || '').localeCompare(String(b.title || ''), 'zh');
        });

    if (els.summary) els.summary.textContent = String(monthEvents.length);

    if (monthEvents.length === 0) {
        els.content.innerHTML = '<div class="course-list-empty">本月无课程</div>';
        return;
    }

    const byDay = new Map();
    monthEvents.forEach(ev => {
        const k = ev.date.toISOString().slice(0, 10);
        if (!byDay.has(k)) byDay.set(k, []);
        byDay.get(k).push(ev);
    });

    const dayNames = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];

    const parts = [];
    Array.from(byDay.keys()).sort().forEach(k => {
        const list = byDay.get(k) || [];
        const d = list[0] && list[0].date ? list[0].date : new Date(k);
        const title = `${d.getMonth() + 1}月${d.getDate()}日 ${dayNames[d.getDay()] || ''}`;
        parts.push(`<div class="course-list-group"><div class="course-list-group-title"><span>${title}</span><span>${list.length}门</span></div><div class="course-list-items">`);
        list.forEach(ev => {
            const pRange = ev.periodRange ? ev.periodRange : ev.period;
            const time = ev.startTime && ev.endTime ? `${ev.startTime}-${ev.endTime}` : '';
            const loc = ev.location ? String(ev.location) : '—';
            const wk = ev.week ? `第${ev.week}周` : '';
            parts.push(
                `<div class="course-list-item">` +
                `<div class="course-list-item-head"><div class="course-list-item-name">${String(ev.title || '未命名课程')}</div><div class="course-list-item-meta"><span>${wk}</span><span>第${String(pRange)}节</span></div></div>` +
                `<div class="course-list-item-meta"><span>${time}</span><span>${loc}</span></div>` +
                `</div>`
            );
        });
        parts.push('</div></div>');
    });

    els.content.innerHTML = parts.join('');
}

function scheduleLLMEnsureCalendarScaffold() {
    const container = document.getElementById('calendarArea');
    if (!container) return null;

    let nav = container.querySelector('.calendar-nav');
    let viewport = container.querySelector('.calendar-month-viewport');

    if (!nav || !viewport) {
        container.innerHTML = '';

        nav = document.createElement('div');
        nav.className = 'calendar-nav no-print';
        nav.innerHTML = `
            <button type="button" class="calendar-nav-btn prev">前一个月</button>
            <div class="calendar-nav-title"></div>
            <button type="button" class="calendar-nav-btn next">后一个月</button>
        `;

        viewport = document.createElement('div');
        viewport.className = 'calendar-month-viewport';

        container.appendChild(nav);
        container.appendChild(viewport);

        const prevBtn = nav.querySelector('button.prev');
        const nextBtn = nav.querySelector('button.next');
        if (prevBtn) prevBtn.addEventListener('click', () => scheduleLLMChangeMonth(-1));
        if (nextBtn) nextBtn.addEventListener('click', () => scheduleLLMChangeMonth(1));
    }

    return {
        container,
        nav,
        viewport,
        titleEl: nav.querySelector('.calendar-nav-title'),
        prevBtn: nav.querySelector('button.prev'),
        nextBtn: nav.querySelector('button.next')
    };
}

function scheduleLLMRenderMonth(date, direction) {
    const els = scheduleLLMEnsureCalendarScaffold();
    if (!els) return;

    const monthDate = scheduleLLMMonthStart(date);
    currentCalendarDate = monthDate;

    if (els.titleEl) els.titleEl.textContent = scheduleLLMFormatMonthTitle(monthDate);

    const prevTarget = scheduleLLMMonthStart(new Date(monthDate.getFullYear(), monthDate.getMonth() - 1, 1));
    const nextTarget = scheduleLLMMonthStart(new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 1));

    if (els.prevBtn) els.prevBtn.disabled = !!scheduleLLMMinMonth && scheduleLLMCompareMonth(prevTarget, scheduleLLMMinMonth) < 0;
    if (els.nextBtn) els.nextBtn.disabled = !!scheduleLLMMaxMonth && scheduleLLMCompareMonth(nextTarget, scheduleLLMMaxMonth) > 0;

    const render = () => {
        els.viewport.innerHTML = '';
        els.viewport.appendChild(createMonthCalendarElement(monthDate, { includeTitle: false }));
    };

    const reduce = typeof window !== 'undefined' && window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    const canAnim = !reduce && els.viewport && typeof els.viewport.animate === 'function';

    if (!canAnim || !direction) {
        render();
        scheduleLLMUpdateCourseListForMonth(monthDate);
        return;
    }

    const outDx = direction > 0 ? -12 : 12;
    const inDx = direction > 0 ? 12 : -12;

    els.viewport
        .animate([{ opacity: 1, transform: 'translateX(0)' }, { opacity: 0, transform: `translateX(${outDx}px)` }], { duration: 160, easing: 'ease' })
        .finished
        .catch(() => {})
        .then(() => {
            render();
            scheduleLLMUpdateCourseListForMonth(monthDate);
            els.viewport.animate([{ opacity: 0, transform: `translateX(${inDx}px)` }, { opacity: 1, transform: 'translateX(0)' }], { duration: 200, easing: 'ease' });
        });
}

function scheduleLLMChangeMonth(delta) {
    const base = scheduleLLMMonthStart(currentCalendarDate || new Date());
    const target = new Date(base.getFullYear(), base.getMonth() + delta, 1);
    const clamped = scheduleLLMClampMonthToRange(target);
    if (scheduleLLMCompareMonth(base, clamped) === 0) return;
    scheduleLLMRenderMonth(clamped, delta);
}

function renderCalendar(events) {
    scheduleLLMSetCourseListVisible(Array.isArray(events) && events.length > 0);
    scheduleLLMSetMonthRangeFromEvents(events);
    const target = (Array.isArray(events) && events.length > 0 && scheduleLLMMinMonth) ? new Date(scheduleLLMMinMonth) : scheduleLLMMonthStart(new Date());
    scheduleLLMRenderMonth(target, 0);
}

function formatWeekRanges(weeks) {
    if (!weeks || weeks.length === 0) return "";
    // Ensure sorted and unique
    const uniqueWeeks = Array.from(new Set(weeks)).sort((a, b) => a - b);
    const ranges = [];
    let start = uniqueWeeks[0];
    let end = uniqueWeeks[0];
    
    for (let i = 1; i < uniqueWeeks.length; i++) {
        if (uniqueWeeks[i] === end + 1) {
            end = uniqueWeeks[i];
        } else {
            ranges.push(start === end ? `${start}` : `${start}-${end}`);
            start = uniqueWeeks[i];
            end = uniqueWeeks[i];
        }
    }
    ranges.push(start === end ? `${start}` : `${start}-${end}`);
    return `第${ranges.join(',')}周`;
}

function createMonthCalendarElement(date, options) {
    const year = date.getFullYear();
    const month = date.getMonth();

    const includeTitle = !(options && options.includeTitle === false);
    const printMode = !!(options && options.printMode);

    const monthContainer = document.createElement('div');
    monthContainer.className = 'month-container';

    if (includeTitle) {
        const title = document.createElement('div');
        title.className = 'month-title';
        title.textContent = `${year}年 ${month + 1}月`;
        monthContainer.appendChild(title);
    }

    const grid = document.createElement('div');
    grid.className = 'calendar-grid';

    let hasWeekendCourses = false;

    // Header
    const days = ['周一', '周二', '周三', '周四', '周五', '周六', '周日'];
    days.forEach(d => {
        const dh = document.createElement('div');
        dh.className = 'calendar-header-cell';
        dh.textContent = d;
        grid.appendChild(dh);
    });

    // Days
    // Calculate first day of month
    const firstDayOfMonth = new Date(year, month, 1);
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    
    // Adjust logic: Week starts on Monday (1) -> Sunday (7)
    // getDay(): 0(Sun), 1(Mon)...
    let startDayOfWeek = firstDayOfMonth.getDay(); 
    if (startDayOfWeek === 0) startDayOfWeek = 7; // Convert Sun 0 to 7

    // Empty slots before 1st
    for (let i = 1; i < startDayOfWeek; i++) {
        const empty = document.createElement('div');
        empty.className = 'calendar-day empty';
        grid.appendChild(empty);
    }

    // Days
    for (let d = 1; d <= daysInMonth; d++) {
        const currentDayDate = new Date(year, month, d);
        const dayEl = document.createElement('div');
        dayEl.className = 'calendar-day';
        
        // Find events for this day
        const dayEvents = generatedEvents.filter(e => 
            e.date.getFullYear() === year && 
            e.date.getMonth() === month && 
            e.date.getDate() === d
        );

        // Sort by start time
        dayEvents.sort((a, b) => a.period - b.period);

        // Header: Day Num + Week Badge (if Monday)
        const dayNumRow = document.createElement('div');
        dayNumRow.className = 'day-number';
        
        const daySpan = document.createElement('span');
        daySpan.textContent = d;
        dayNumRow.appendChild(daySpan);

        // Add Week Badge if it's Monday or 1st of month (to show context)
        // Or just show for any day that has events? No, consistency.
        // Let's show for every Monday and the 1st day of month.
        let dayOfWeek = currentDayDate.getDay();
        if (dayOfWeek === 0) dayOfWeek = 7;

        if (dayOfWeek === 1 || d === 1) {
            // Find week number from the first event of this day?
            // Or calculate from semester start? 
            // Since events have 'week' property, we can pick one.
            // If no events, we can't easily know week unless we pass semesterStart globally.
            // But we can peek at any event in this week?
            // Simplest: If dayEvents has items, take first one's week.
            if (dayEvents.length > 0) {
                const badge = document.createElement('span');
                badge.className = 'week-badge';
                badge.textContent = `第${dayEvents[0].week}周`;
                dayNumRow.appendChild(badge);
            }
        }

        dayEl.appendChild(dayNumRow);

        if (printMode && (dayOfWeek === 6 || dayOfWeek === 7)) {
            dayEl.classList.add('weekend');
        }

        const eventsWrap = document.createElement('div');
        eventsWrap.className = 'day-events';

        // Render Events
        // 1. Group Events for Joint Classes (Combined Display)
        // Key: Time + Location + Course Name
        const groupedEvents = new Map();

        dayEvents.forEach(ev => {
            const key = `${ev.period}-${ev.location}-${ev.title}`;
            if (!groupedEvents.has(key)) {
                groupedEvents.set(key, {
                    ...ev,
                    classNames: [ev.className] // Initialize list
                });
            } else {
                const existing = groupedEvents.get(key);
                if (ev.className && !existing.classNames.includes(ev.className)) {
                    existing.classNames.push(ev.className);
                }
            }
        });

        // Convert Map back to Array
        const displayEvents = Array.from(groupedEvents.values());
        const limitedEvents = printMode ? displayEvents.slice(0, 5) : displayEvents;

        if (printMode && dayEl.classList.contains('weekend')) {
            if (limitedEvents.length === 0) {
                dayEl.classList.add('print-narrow');
            } else {
                hasWeekendCourses = true;
            }
        }

        if (printMode && !dayEl.classList.contains('print-narrow') && limitedEvents.length > 0) {
            const marker = document.createElement('span');
            marker.className = 'day-marker';
            marker.textContent = `${limitedEvents.length}课`;
            dayNumRow.appendChild(marker);
        }

        limitedEvents.forEach(ev => {
            const evEl = document.createElement('div');
            evEl.className = `event-item type-${ev.timeOfDay}`;

            const pRange = ev.periodRange ? ev.periodRange : ev.period;

            if (printMode) {
                const loc = ev.location ? String(ev.location) : '';
                const title = ev.title ? String(ev.title) : '';
                evEl.innerHTML = `
                    <div class="ev-print-line1">${String(pRange)}节 ${loc}</div>
                    <div class="ev-print-line2">${title}</div>
                `;
            } else {
                // Tooltip Content
                const weeksText = formatWeekRanges(ev.weeks);

                // Combine Class Names: "Class1/Class2"
                // Filter empty names and clean them
                const cleanNames = ev.classNames
                    .filter(n => n)
                    .map(n => n.replace(/^[\(（]/, '').replace(/[\)）]$/, ''));

                const tooltipClassText = cleanNames.length > 0 ? cleanNames.join('/') : "";

                evEl.innerHTML = `
                    <div class="ev-header">
                        <span class="ev-period">第${pRange}节</span>
                        <span class="ev-location-separator">@</span>
                        <span class="ev-location">${ev.location}</span>
                    </div>
                    <div class="ev-course-name">
                        ${ev.title}
                        <div class="ev-tooltip">
                            ${tooltipClassText ? `<div>${tooltipClassText}</div>` : ''}
                            <div>${weeksText}</div>
                        </div>
                    </div>
                `;
            }

            eventsWrap.appendChild(evEl);
        });

        dayEl.appendChild(eventsWrap);
        grid.appendChild(dayEl);
    }

    if (printMode && hasWeekendCourses) {
        grid.classList.add('print-weekend-has-courses');
    }

    monthContainer.appendChild(grid);
    return monthContainer;
}

// Print Handler
const btnPrint = document.getElementById('btnPrint');
if (btnPrint) {
    btnPrint.addEventListener('click', () => {
        if (generatedEvents.length === 0) {
            alert("无日程数据");
            return;
        }

        let minTime = Infinity;
        let maxTime = -Infinity;
        generatedEvents.forEach(e => {
            const t = e.date.getTime();
            if (t < minTime) minTime = t;
            if (t > maxTime) maxTime = t;
        });

        const startDate = new Date(minTime);
        const startYear = startDate.getFullYear();
        const startMonth = startDate.getMonth();

        const endDate = new Date(maxTime);
        const endYear = endDate.getFullYear();
        const endMonth = endDate.getMonth();

        const area = document.getElementById('calendarArea');
        if (!area) {
            alert("找不到月历容器");
            return;
        }

        const OriginalHTML = area.innerHTML;

        const restore = () => {
            area.innerHTML = OriginalHTML;

            const nav = area.querySelector('.calendar-nav');
            if (nav) {
                const prevBtn = nav.querySelector('button.prev');
                const nextBtn = nav.querySelector('button.next');
                if (prevBtn) prevBtn.addEventListener('click', () => scheduleLLMChangeMonth(-1));
                if (nextBtn) nextBtn.addEventListener('click', () => scheduleLLMChangeMonth(1));
            }

            window.dispatchEvent(new Event('resize'));
        };

        area.innerHTML = '';

        const bar = document.createElement('div');
        bar.className = 'no-print';
        bar.style.position = 'sticky';
        bar.style.top = '0';
        bar.style.zIndex = '9999';
        bar.style.display = 'flex';
        bar.style.justifyContent = 'space-between';
        bar.style.alignItems = 'center';
        bar.style.gap = '10px';
        bar.style.padding = '10px 12px';
        bar.style.marginBottom = '10px';
        bar.style.border = '1px solid #e2e8f0';
        bar.style.borderRadius = '10px';
        bar.style.background = '#ffffff';
        bar.style.boxShadow = '0 1px 2px 0 rgb(0 0 0 / 0.05)';
        bar.innerHTML = `<div style="font-weight:800; color:#0f172a;">打印预览</div><div style="display:flex; gap:10px;"><button type="button" id="scheduleLLMPrintDo" style="padding:8px 12px; border-radius:10px; border:1px solid #e2e8f0; background:#2563eb; color:#fff; font-weight:800; cursor:pointer;">打印</button><button type="button" id="scheduleLLMPrintCancel" style="padding:8px 12px; border-radius:10px; border:1px solid #e2e8f0; background:#fff; color:#0f172a; font-weight:800; cursor:pointer;">返回</button></div>`;

        const printContainer = document.createElement('div');
        printContainer.className = 'print-all-container';

        let iterDate = new Date(startYear, startMonth, 1);
        while (iterDate.getFullYear() < endYear || (iterDate.getFullYear() === endYear && iterDate.getMonth() <= endMonth)) {
            const monthEl = createMonthCalendarElement(new Date(iterDate), { printMode: true });
            printContainer.appendChild(monthEl);
            iterDate.setMonth(iterDate.getMonth() + 1);
        }

        area.appendChild(bar);
        area.appendChild(printContainer);

        const doBtn = document.getElementById('scheduleLLMPrintDo');
        const cancelBtn = document.getElementById('scheduleLLMPrintCancel');

        if (cancelBtn) {
            cancelBtn.addEventListener('click', restore);
        }

        if (doBtn) {
            doBtn.addEventListener('click', () => {
                doBtn.disabled = true;
                doBtn.style.opacity = '0.8';
                window.print();
                setTimeout(restore, 1000);
            });
        }
    });
}

// Export Logic
document.getElementById('btnExport').addEventListener('click', () => {
    if (generatedEvents.length === 0) {
        alert("无日程数据");
        return;
    }

    // Generate ICS content
    let device = document.getElementById('exportTarget').value;
    let prodId = "-//ScheduleLLM//CN";
    if (device === 'windows') prodId = "-//Microsoft Corporation//Outlook 16.0 MIMEDIR//EN";
    if (device === 'ios') prodId = "-//Apple Inc.//iOS 15.0//EN";

    let icsContent = `BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:${prodId}\r\nCALSCALE:GREGORIAN\r\nMETHOD:PUBLISH\r\n`;

    // Windows Outlook: Add TimeZone Definition? 
    // Simplify for now, usually VEVENT stats are enough.

    if (device === 'vcard') {
        // Export as vCalendar 1.0 (.vcs) which is often compatible with older systems or "Contact Schedules"
        let vcsContent = `BEGIN:VCALENDAR\r\nVERSION:1.0\r\nPRODID:-//ScheduleLLM//CN\r\nTZ:-08\r\n`;

        generatedEvents.forEach(ev => {
            const dayStr = ev.date.toISOString().split('T')[0].replace(/-/g, '');
            const startStr = `${dayStr}T${ev.startTime.replace(/:/g, '')}00`;
            const endStr = `${dayStr}T${ev.endTime.replace(/:/g, '')}00`;

            vcsContent += "BEGIN:VEVENT\r\n";
            vcsContent += `SUMMARY:${ev.title}\r\n`;
            vcsContent += `DTSTART:${startStr}\r\n`;
            vcsContent += `DTEND:${endStr}\r\n`;
            vcsContent += `LOCATION:${ev.location}\r\n`;
            vcsContent += `DESCRIPTION:${ev.description}\r\n`;
            vcsContent += "END:VEVENT\r\n";
        });

        vcsContent += "END:VCALENDAR";

        const blob = new Blob([vcsContent], { type: 'text/x-vcalendar;charset=utf-8' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `schedule_export.vcs`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        return;
    }

    generatedEvents.forEach(ev => {
        // Format Date: YYYYMMDDTHHMMSS
        const dayStr = ev.date.toISOString().split('T')[0].replace(/-/g, '');
        const startStr = `${dayStr}T${ev.startTime.replace(/:/g, '')}00`;
        const endStr = `${dayStr}T${ev.endTime.replace(/:/g, '')}00`;

        let description = ev.description;
        if (device === 'ios') {
            // iOS sometimes likes cleaner description
        }

        icsContent += "BEGIN:VEVENT\r\n";
        icsContent += `UID:${Date.now()}-${Math.random()}@schedulellm\r\n`;
        icsContent += `DTSTAMP:${new Date().toISOString().replace(/[-:]/g, '').split('.')[0]}Z\r\n`;
        icsContent += `DTSTART;TZID=Asia/Shanghai:${startStr}\r\n`;
        icsContent += `DTEND;TZID=Asia/Shanghai:${endStr}\r\n`;
        icsContent += `SUMMARY:${ev.title}\r\n`;
        icsContent += `LOCATION:${ev.location}\r\n`;
        icsContent += `DESCRIPTION:${description}\r\n`;

        // Alarms
        if (device === 'ios' || device === 'android') {
            // 15 min reminder
            icsContent += "BEGIN:VALARM\r\nTRIGGER:-PT15M\r\nACTION:DISPLAY\r\nDESCRIPTION:Reminder\r\nEND:VALARM\r\n";
        }

        // Windows Outlook specific categories?
        if (device === 'windows') {
            const cat = ev.timeOfDay === 'morning' ? 'Blue Category' : (ev.timeOfDay === 'afternoon' ? 'Orange Category' : 'Purple Category');
            // icsContent += `CATEGORIES:${cat}\r\n`; // Outlook might need Master List, but safe to add
            icsContent += `X-MICROSOFT-CDO-BUSYSTATUS:BUSY\r\n`;
        }

        icsContent += "END:VEVENT\r\n";
    });

    icsContent += "END:VCALENDAR";

    const blob = new Blob([icsContent], { type: 'text/calendar;charset=utf-8' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `schedule_${device}.ics`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
});

// HTML Export
document.getElementById('btnSaveHtml').addEventListener('click', () => {
    if (generatedEvents.length === 0) {
        alert("无日程数据");
        return;
    }

    // 1. Calculate Date Range (Copy from Print logic)
    let minTime = Infinity;
    let maxTime = -Infinity;
    generatedEvents.forEach(e => {
        const t = e.date.getTime();
        if (t < minTime) minTime = t;
        if (t > maxTime) maxTime = t;
    });

    const startDate = new Date(minTime);
    const startYear = startDate.getFullYear();
    const startMonth = startDate.getMonth();

    const endDate = new Date(maxTime);
    const endYear = endDate.getFullYear();
    const endMonth = endDate.getMonth();

    // 2. Generate Content
    const container = document.createElement('div');
    container.className = 'print-all-container'; // Reuse print container class for layout

    let iterDate = new Date(startYear, startMonth, 1);
    while (iterDate.getFullYear() < endYear || (iterDate.getFullYear() === endYear && iterDate.getMonth() <= endMonth)) {
        const monthEl = createMonthCalendarElement(new Date(iterDate));
        container.appendChild(monthEl);
        iterDate.setMonth(iterDate.getMonth() + 1);
    }

    // 3. Define a vibrant and modern stylesheet for the export
    const cssText = `
        :root {
            --primary: #3b82f6;
            --bg: #f1f5f9;
            --card-bg: #ffffff;
            --text-main: #1e293b;
            --text-muted: #64748b;
            --border: #e2e8f0;
            --morning: #10b981;
            --afternoon: #f59e0b;
            --evening: #6366f1;
            --radius: 12px;
            --shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1);
        }

        * { box-sizing: border-box; margin: 0; padding: 0; }

        body { 
            background: var(--bg); 
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; 
            color: var(--text-main);
            line-height: 1.5;
            padding: 40px 20px;
            overflow-y: auto;
            height: auto;
        }

        .export-header {
            max-width: 1200px;
            margin: 0 auto 40px auto;
            text-align: center;
        }

        .export-header h1 {
            font-size: 2.5rem;
            font-weight: 800;
            color: var(--primary);
            margin-bottom: 8px;
            letter-spacing: -0.025em;
        }

        .export-header p {
            color: var(--text-muted);
            font-size: 1.1rem;
        }

        .content-wrapper {
            max-width: 1200px;
            margin: 0 auto;
        }

        .month-container { 
            background: var(--card-bg);
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            padding: 24px;
            margin-bottom: 40px; 
            border: 1px solid var(--border);
        }

        .month-title { 
            font-size: 1.5rem;
            font-weight: 700;
            color: var(--text-main);
            margin-bottom: 20px;
            text-align: left;
            padding-left: 8px;
            border-left: 4px solid var(--primary);
        }

        .calendar-grid { 
            display: grid; 
            grid-template-columns: repeat(7, 1fr); 
            gap: 12px; 
        }

        .calendar-header-cell { 
            text-align: center; 
            color: var(--text-muted);
            font-size: 0.875rem;
            font-weight: 600;
            padding: 8px;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        .calendar-day { 
            background: #f8fafc;
            border: 1px solid var(--border);
            border-radius: 8px;
            min-height: 120px; 
            padding: 8px; 
            display: flex;
            flex-direction: column;
            gap: 6px;
            transition: transform 0.2s, box-shadow 0.2s;
        }

        .calendar-day:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.1);
            background: #fff;
        }

        .calendar-day.empty { 
            background: transparent;
            border: 1px dashed var(--border);
        }

        .day-number { 
            font-size: 0.875rem; 
            font-weight: 700;
            display: flex; 
            justify-content: space-between; 
            align-items: center;
            margin-bottom: 4px;
            color: var(--text-muted);
        }

        .week-badge {
            font-size: 0.7rem;
            background: #eff6ff;
            color: var(--primary);
            padding: 2px 6px;
            border-radius: 4px;
            font-weight: 600;
        }

        .event-item { 
            padding: 6px 8px; 
            border-radius: 6px;
            font-size: 0.75rem; 
            font-weight: 500;
            line-height: 1.3;
            border-left: 3px solid transparent;
        }

        .type-morning { 
            background: #ecfdf5; 
            color: #065f46;
            border-left-color: var(--morning); 
        }

        .type-afternoon { 
            background: #fffbeb; 
            color: #92400e;
            border-left-color: var(--afternoon); 
        }

        .type-evening { 
            background: #f5f3ff; 
            color: #3730a3;
            border-left-color: var(--evening); 
        }

        .ev-time { 
            font-weight: 700; 
            display: block;
            font-size: 0.7rem;
            opacity: 0.8;
            margin-bottom: 2px;
        }

        .ev-location { 
            display: block;
            font-size: 0.7rem;
            font-style: italic;
            margin-top: 2px;
            opacity: 0.8;
            margin-bottom: 2px;
        }

        .ev-title {
            word-break: break-word;
        }
        
        /* Tooltip Styles for Export */
        .ev-header {
            font-size: 12px;
            font-weight: 400;
            color: var(--text-muted);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            display: flex;
            align-items: center;
            gap: 2px;
        }

        .ev-period {
            font-size: 11px;
            font-family: "Inter", "Microsoft YaHei", sans-serif;
            color: #2B579A;
            background-color: #dbeafe;
            padding: 1px 5px;
            border-radius: 4px;
            font-weight: 600;
            letter-spacing: 0.5px;
        }

        .ev-location-separator {
             color: #cbd5e1;
             font-size: 11px;
             margin: 0 2px;
        }

        .ev-location {
            font-size: 13px;
            font-family: "Inter", "Microsoft YaHei", sans-serif;
            color: #333333;
            font-weight: 500;
        }

        .ev-course-name {
            font-size: 14px;
            font-weight: 700;
            color: var(--text-main);
            margin-top: 2px;
            position: relative;
            cursor: pointer;
            line-height: 1.4;
        }

        .ev-tooltip {
            visibility: hidden;
            opacity: 0;
            position: absolute;
            background: #333;
            color: white;
            padding: 8px 12px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 400;
            width: max-content;
            max-width: 200px;
            top: 100%;
            left: 0;
            z-index: 1000;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
            transition: opacity 0.2s ease-in-out, visibility 0.2s ease-in-out;
            transition-delay: 0.2s;
            pointer-events: none;
        }

        .ev-course-name:hover .ev-tooltip {
            visibility: visible;
            opacity: 1;
            transition-delay: 0.4s;
        }

        @media (hover: none) {
            .ev-course-name:active .ev-tooltip {
                visibility: visible;
                opacity: 1;
                transition-delay: 0s;
            }
        }

        @media print {
            body { background: white; padding: 0; }
            .month-container { box-shadow: none; border-color: #eee; page-break-inside: avoid; }
        }
    `;

    const fullHtml = `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>我的课表 - ScheduleLLM</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap" rel="stylesheet">
    <style>
        ${cssText}
    </style>
</head>
<body>
    <header class="export-header">
        <h1>我的课程表</h1>
        <p>由 ScheduleLLM 自动生成</p>
    </header>
    <div class="content-wrapper">
        ${container.innerHTML}
    </div>
</body>
</html>
    `;

    // 4. Download
    const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `schedule_export.html`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
});
