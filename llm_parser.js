/**
 * LLM Parser Module for ScheduleLLM
 * Integrates with Alibaba Cloud Qwen (Tongyi Qianwen) and other models.
 */
window.__SCHEDULELLM_DEBUG_LLM = true  // 开启 ScheduleLLM 的 LLM 调试日志开关（true 时会在控制台输出详细请求与响应信息）
class LLMService {
    constructor() {
        this.config = {
            baseUrl: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
            apiKey: '',
            model: 'qwen-flash'
        };
        this.cache = new Map();
    }

    _debugEnabled() {
        return typeof window !== 'undefined' && !!window.__SCHEDULELLM_DEBUG_LLM;
    }

    _clip(str, maxLen = 1200) {
        const s = String(str || "");
        if (s.length <= maxLen) return s;
        return s.slice(0, maxLen) + `…[+${s.length - maxLen}]`;
    }

    updateConfig(baseUrl, apiKey, model) {
        this.config.baseUrl = baseUrl.replace(/\/$/, ''); // Remove trailing slash
        this.config.apiKey = apiKey;
        this.config.model = model;
    }

    /**
     * Standardized Request to LLM
     * @param {string} rawText - The cell content from Excel
     * @param {object} context - Additional context (e.g., current day, time)
     * @returns {Promise<object>} - Structured course data
     */
    async parseCourse(rawText, context = {}) {
        if (!rawText || !rawText.trim()) {
            return { courses: [], confidence: 1.0, error: null };
        }

        const debug = this._debugEnabled();
        const t0 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();

        const cacheKey = `${this.config.model}:${rawText}`;
        if (this.cache.has(cacheKey)) {
            if (debug) {
                console.groupCollapsed(`[LLM][cache_hit] model=${this.config.model}`);
                console.debug("input", this._clip(rawText, 600));
                console.groupEnd();
            }
            return this.cache.get(cacheKey);
        }

        const prompt = this.constructPrompt(rawText);

        if (debug) {
            console.groupCollapsed(`[LLM][request] model=${this.config.model}`);
            console.debug("baseUrl", this.config.baseUrl);
            console.debug("input", this._clip(rawText, 1200));
            console.debug("prompt", this._clip(prompt, 2000));
        }

        try {
            const tFetch0 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
            const response = await fetch(`${this.config.baseUrl}/chat/completions`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${this.config.apiKey}`
                },
                body: JSON.stringify({
                    model: this.config.model,
                    messages: [
                        {
                            role: "system",
                            content: `你是一个专业的课程表解析助手。
你的任务是从原始文本中高精度地提取课程信息，特别是班级名称和上课地点。
输出必须是符合以下结构的有效 JSON 对象：
{
    "courses": [
        {
            "name": "课程名称", // 仅包含学科名，移除班级或地点信息
            "weeks": [1, 2, 3], // 整数周次数组。必须展开范围！
            "location": "上课地点", // 完整的原始地点字符串：校区 + 楼栋 + 教室。保留中文。
            "building": "楼栋名称", // 提取的楼栋名。保留中文。
            "room": "教室号", // 提取的教室编号（如 "203", "A105"）
            "className": "班级名称", // 标准化班级名（如 "软件2023-1", "计科1班"）
            "periodRange": "1-2", // 节次范围（如果指定，如 "1-2节"）
            "teacher": "教师姓名",
            "raw_weeks": "1-16周" // 原始周次字符串
        }
    ],
    "confidence": 0.9 // 置信度 0-1
}

地点提取规则 ("location", "building", "room")：
1. "location"：必须严格保持原始输入中的地点名称格式（按输入原文截取），禁止任何形式的补全、扩展、规范化、翻译。
   - 例如：输入包含 "一教"，则输出必须仍为 "一教"，不得输出 "第一教学楼"。
   - 例如：输入 "桂林洋一教" -> 输出 "桂林洋一教"，不得输出 "桂林洋校区第一教学楼"。
2. 仅允许进行必要的空格与标点符号修正（例如移除多余空格、统一全角/半角标点），不得添加输入中不存在的词（如 "校区"、"第一"、"公共教学楼" 等）。
3. "building"：从输入中提取楼栋名称，必须保持输入原文写法（含简称）。如果无法可靠区分楼栋与教室号，可令 building 为空，但 location 仍必须保真。
4. "room"：严格提取教室编号，必须包含数字（例如 "203", "B105", "S103"）。
5. 忽略 "多媒体教室"、"实验室"、"室" 等描述性词语，除非它们原本就是地点名称的一部分（仍需保持原文）。

换行中断修复规则（课程名/地点名）:
1. 输入将以 JSON 字符串提供：包含 original（原文，含 \n）、marked（把 \n 标记为 ⏎）、preprocessed（用于识别的轻度合并版本）、lineBreaks（\n 的索引数组）。
2. 课程名称（name）：
   - 如果识别到课程名被 ⏎ 分割，允许在不新增字符的前提下合并片段（本质是移除换行导致的断裂）。
   - 合并前后需做一致性检查：片段均应符合课程名常见形态（连续中文/英文/数字/括号/点号），且合并后不应跨越明显字段边界（如 "/"、"周"、"节"、"班"、"教室" 等）。


   - 特殊高频形态：如果在周次/节次之前出现了“X专业⏎导论/概论/基础/原理/实验/实训”等断裂，应优先合并为完整课程名（例如“电气工程及其自动化专业⏎导论” -> name="电气工程及其自动化专业导论"）。
   - 输出 repairs 记录该修复，并给出 confidence；低于 0.8 时不要修复。
3. 地点名称（location/building/room）：
   - 识别可能被 ⏎ 分割的地点片段，允许合并以恢复连续地点（例如 "桂林\n洋工程S308" -> "桂林洋工程S308"）。
   - 合并需满足：合并后能匹配房间号形态（必须包含数字，如 203/A105/S103），且不跨越班级/周次/节次等字段边界。
   - 禁止借助任何知识库进行地点名称扩展或规范化。
4. 位置与标注：
   - 需要在输出中提供 nameSpan 与 locationSpan（在 original 字符串中的 [start,end) 索引），便于人工验证。
   - repairs 数组中明确标注哪些字段经过换行修复（from/to/reason/confidence/spans）。

周次提取规则 ("weeks")：
1. **必须**将周次字符串解析为整数数组。
2. 处理范围： "1-16周" -> [1, 2, ..., 16]。
3. 处理单双周：
   - "1-16周(单)" 或 "1-16单" -> [1, 3, 5, ..., 15]
   - "2-16周(双)" 或 "2-16双" -> [2, 4, 6, ..., 16]
4. 处理多段周次："1-8, 11-16周" -> [1..8, 11..16]。
5. 如果隐含或明确指出 "每周" 且包含范围，则包含范围内的所有周次。

班级名称提取规则 ("className")：
1. **仅当**字符串明确描述“学生群体/班级”时才输出为 className：通常包含 "班"、"级"、"届"、年级数字、班号等。
2. 重要：仅出现“专业”并不等价于班级（例如“电气工程及其自动化专业导论”里的“电气工程及其自动化专业”通常是课程名的一部分）。
3. 如果字符串以“专业”结尾但不包含年级/班号/届/班级标识，则默认不要作为 className，优先与相邻片段合并用于 name。
4. **格式**：年级 + 专业 + 班号（例如 "21软件1班"、"2023级计科2班"）。
5. **移除**提取出的 className 中的括号（例如：如果文本是 "(软件2101)"，则提取 "软件2101"）。
6. **处理合班情况**：如果多个班级名称连接在一起（例如 "软件2101软件2102"、"1班;2班"），必须进行**拆分**。
   - 如果多个班级共用同一课程/时间，请将它们合并为一个字符串，并用**逗号**分隔（例如 "软件2101, 软件2102"）。
   - 识别分隔符，如分号 (;)、空格或隐式边界（例如 "...1班...2班"）。

通用规则：
1. 如果文本中包含多门课程，请列出所有课程。
2. 如果未发现课程，返回空数组。
3. 处理简化格式（例如 "数学 1-16周 101室"）。
4. 处理 '◇' (菱形) 作为字段分隔符的情况（例如 "课程◇周次◇地点◇..."）。
5. 如果提及具体节次范围，请提取（例如 "(1-2节)"）。
6. **不要**包含 Markdown 格式（如 \`\`\`json）。仅返回纯 JSON 字符串。`
                        },
                        {
                            role: "user",
                            content: prompt
                        }
                    ],
                    temperature: 0.1
                })
            });

            if (!response.ok) {
                const errData = await response.json();
                throw new Error(`API Error: ${response.status} - ${JSON.stringify(errData)}`);
            }

            const tFetch1 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
            const tJson0 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
            const data = await response.json();
            const tJson1 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();

            const rawAssistant = (data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) ? data.choices[0].message.content : "";
            const content = String(rawAssistant).trim().replace(/^```json/, '').replace(/```$/, '');

            if (debug) {
                console.debug("http", { ok: response.ok, status: response.status, fetchMs: +(tFetch1 - tFetch0).toFixed(1), jsonMs: +(tJson1 - tJson0).toFixed(1) });
                console.debug("raw_response", this._clip(rawAssistant, 4000));
            }

            let result;
            try {
                result = JSON.parse(content);
            } catch (e) {
                if (debug) {
                    console.debug("json_parse_failed", this._clip(content, 4000));
                }
                console.warn("LLM returned invalid JSON", content);
                throw new Error("Invalid JSON response from LLM");
            }

            // Validate structure
            if (!result.courses || !Array.isArray(result.courses)) {
                result.courses = [];
            }

            const tPost0 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();

            const normalizePunctLight = (str) => String(str || '')
                .replace(/[\u3000\u00A0]/g, ' ')
                .replace(/[（]/g, '(')
                .replace(/[）]/g, ')')
                .replace(/[：]/g, ':');

            const normalizeSpaces = (str) => normalizePunctLight(str)
                .replace(/\s+/g, ' ')
                .trim();

            const normalizeNoSpaces = (str) => normalizePunctLight(str)
                .replace(/[\s]+/g, '')
                .trim();

            result.courses.forEach(course => {
                if (course.raw_weeks) {
                    const calculatedWeeks = this.parseWeekString(course.raw_weeks);
                    if (calculatedWeeks.length > 0) {
                        course.weeks = calculatedWeeks;
                    }
                }

                if (course && typeof course === 'object') {
                    if (typeof course.className === 'string') {
                        const before = course.className;
                        const after = normalizeNoSpaces(before).replace(/^[\(（]/, '').replace(/[\)）]$/, '');
                        if (before !== after) {
                            if (debug) console.debug('field_fix', { field: 'className', before, after });
                            course.className = after;
                        }
                    }

                    if (typeof course.room === 'string') {
                        const before = course.room;
                        const after = normalizeNoSpaces(before);
                        if (before !== after) {
                            if (debug) console.debug('field_fix', { field: 'room', before, after });
                            course.room = after;
                        }
                    }

                    if (typeof course.building === 'string') {
                        const before = course.building;
                        const after = normalizeSpaces(before);
                        if (before !== after) {
                            if (debug) console.debug('field_fix', { field: 'building', before, after });
                            course.building = after;
                        }
                    }

                    if (typeof course.location === 'string') {
                        const before = course.location;
                        const after = normalizeSpaces(before);
                        if (before !== after) {
                            if (debug) console.debug('field_fix', { field: 'location', before, after });
                            course.location = after;
                        }
                    }

                    if (typeof course.name === 'string') {
                        const before = course.name;
                        const after = normalizeSpaces(before);
                        if (before !== after) {
                            if (debug) console.debug('field_fix', { field: 'name', before, after });
                            course.name = after;
                        }
                    }

                    if (typeof course.teacher === 'string') {
                        const before = course.teacher;
                        const after = normalizeSpaces(before);
                        if (before !== after) {
                            if (debug) console.debug('field_fix', { field: 'teacher', before, after });
                            course.teacher = after;
                        }
                    }

                    if (typeof course.name === 'string' && typeof course.className === 'string') {
                        const n0 = course.name || '';
                        const c0 = course.className || '';
                        const genericName = /^(导论|概论|基础|原理|实验|实训|课程设计)$/;
                        const looksLikeMajorOnly = /专业$/.test(c0) && !/[班级届]/.test(c0) && !/\d/.test(c0);

                        if (genericName.test(n0) && looksLikeMajorOnly) {
                            const merged = c0 + n0;
                            const orig = String(rawText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
                            const flat = orig.replace(/\n/g, '');

                            if (flat.includes(merged)) {
                                const idxC = orig.indexOf(c0);
                                const idxN = idxC >= 0 ? orig.indexOf(n0, idxC) : orig.indexOf(n0);

                                if (debug) {
                                    console.debug('name_merge_fix', {
                                        name0: n0,
                                        className0: c0,
                                        name1: merged,
                                        className1: '',
                                        idxC,
                                        idxN
                                    });
                                }

                                course.name = merged;
                                course.className = '';

                                if (idxC >= 0 && idxN >= 0) {
                                    course.nameSpan = [idxC, idxN + n0.length];
                                }

                                if (result && typeof result === 'object') {
                                    if (!Array.isArray(result.repairs)) result.repairs = [];
                                    result.repairs.push({
                                        from: 'post',
                                        to: 'post',
                                        reason: '修复“专业”前缀被误判为班级，合并为完整课程名',
                                        confidence: 0.9,
                                        spans: (idxC >= 0 && idxN >= 0) ? [[idxC, idxC + c0.length], [idxN, idxN + n0.length]] : []
                                    });
                                }
                            }
                        }
                    }

                    if (typeof course.building === 'string' && typeof course.room === 'string') {
                        const b0 = course.building || '';
                        const r0 = course.room || '';
                        if (b0 && r0) {
                            let b1 = b0;
                            let r1 = r0;

                            r1 = r1.replace(/^([A-Za-z])\1(\d)/, '$1$2');

                            const m = r1.match(/^([A-Za-z])\d/);
                            if (m && b1.endsWith(m[1])) {
                                b1 = b1.slice(0, -1);
                            }

                            if (b1 !== b0 || r1 !== r0) {
                                if (debug) {
                                    console.debug('location_fix', {
                                        building0: b0,
                                        room0: r0,
                                        building1: b1,
                                        room1: r1,
                                        location: course.location
                                    });
                                }
                                course.building = b1;
                                course.room = r1;
                            }
                        }
                    }
                }
            });

            const tPost1 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();

            if (debug) {
                const t1 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
                console.debug("result", {
                    courses: Array.isArray(result.courses) ? result.courses.length : 0,
                    confidence: result.confidence,
                    postMs: +(tPost1 - tPost0).toFixed(1),
                    totalMs: +(t1 - t0).toFixed(1)
                });
                console.groupEnd();
            }

            this.cache.set(cacheKey, result);
            return result;

        } catch (error) {
            if (debug) {
                const t1 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
                console.debug("error", { message: error && error.message, totalMs: +(t1 - t0).toFixed(1) });
                console.groupEnd();
            }
            console.error("LLM Parsing Failed:", error);
            return {
                courses: [],
                confidence: 0,
                error: error.message
            };
        }
    }

    /**
     * Robust Week String Parser
     * Handles: "1-16", "1-16(单)", "1-3,5-8", "1-16周"
     * @param {string} str 
     * @returns {number[]} Sorted array of week numbers
     */
    parseWeekString(str) {
        if (!str) return [];

        // 1. Normalize Full-width characters to Half-width
        let s = str.replace(/[\uff01-\uff5e]/g, function(ch) {
            return String.fromCharCode(ch.charCodeAt(0) - 0xfee0);
        });
        
        // 2. Remove spaces and "周"
        s = s.replace(/\s+/g, '').replace(/周/g, '');

        // 3. Split by comma/semicolon
        const parts = s.split(/[,;，；]/);
        const weekSet = new Set();

        parts.forEach(part => {
            if (!part) return;

            // Match: Start[-End] ... [Odd/Even]
            // Regex: (\d+)(?:-(\d+))?
            const rangeMatch = part.match(/(\d+)(?:-(\d+))?/);
            if (!rangeMatch) return;

            const start = parseInt(rangeMatch[1]);
            const end = rangeMatch[2] ? parseInt(rangeMatch[2]) : start;
            
            // Check for Odd/Even markers
            const isOdd = part.includes('单');
            const isEven = part.includes('双');

            // Sanity check
            if (start > 50 || end > 50) return;

            for (let i = start; i <= end; i++) {
                if (isOdd && i % 2 === 0) continue;
                if (isEven && i % 2 !== 0) continue;
                weekSet.add(i);
            }
        });

        return Array.from(weekSet).sort((a, b) => a - b);
    }

    constructPrompt(rawText) {
        const original0 = String(rawText || "");
        const original = original0.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

        const lineBreaks = [];
        for (let i = 0; i < original.length; i++) {
            if (original[i] === "\n") lineBreaks.push(i);
        }

        const marked = original.replace(/\n/g, "⏎");

        const preprocessed = original
            .replace(/([A-Za-z0-9\u4e00-\u9fff])\n(?=[A-Za-z0-9\u4e00-\u9fff])/g, "$1")
            .replace(/\n+/g, " / ")
            .replace(/\s*\/\s*/g, " / ")
            .trim();

        return JSON.stringify({
            original,
            marked,
            preprocessed,
            lineBreaks
        });
    }

    clearCache() {
        this.cache.clear();
    }

    /**
     * Health check to verify service availability
     * @returns {Promise<boolean>}
     */
    async checkHealth() {
        // Simple check: do we have an API key?
        // In real scenario, might ping an endpoint.
        if (!this.config.apiKey) return false;
        return true;
    }
}

// Export singleton
const llmService = new LLMService();
window.llmService = llmService;
console.log("LLM Service Initialized");
