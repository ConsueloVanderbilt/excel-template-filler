(() => {
  "use strict";

  const state = { sourceFiles: [], target: null, result: null };
  const $ = id => document.getElementById(id);
  const els = {
    sourceFile: $("sourceFile"), targetFile: $("targetFile"), sourceCard: $("sourceCard"), targetCard: $("targetCard"),
    sourceSummary: $("sourceSummary"), sourceList: $("sourceList"), targetInfo: $("targetInfo"),
    targetDepth: $("targetHeaderDepth"), targetDepthWrap: $("targetDepthWrap"), sourceMatch: $("sourceMatchField"),
    targetMatch: $("targetMatchField"), sourceData: $("sourceDataField"), targetFill: $("targetFillField"),
    fillButton: $("fillButton"), exportButton: $("exportButton"), message: $("globalMessage"),
    results: $("results"), preview: $("previewTable"), resultNotice: $("resultNotice")
  };

  function showMessage(text, type = "warning") {
    els.message.textContent = text;
    els.message.className = `message ${type === "error" ? "error" : ""}`;
  }
  function clearMessage() { els.message.className = "message hidden"; els.message.textContent = ""; }
  function cellText(cell) { return cell ? XLSX.utils.format_cell(cell) : ""; }
  function address(row, col) { return XLSX.utils.encode_cell({ r: row, c: col }); }
  function escapeHtml(value) { return String(value ?? "").replace(/[&<>'"]/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[ch])); }
  function makeId() { return globalThis.crypto?.randomUUID?.() || `file-${Date.now()}-${Math.random().toString(16).slice(2)}`; }

  // 读取一个工作簿。原始文件和模板文件都复用此函数，并保留原始二进制。
  function parseWorkbook(buffer, fileName) {
    const workbook = XLSX.read(buffer, { type: "array", cellStyles: true, cellDates: true, cellNF: true, bookVBA: true });
    if (!workbook.SheetNames.length) throw new Error("Excel 中没有可读取的工作表");
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet || !worksheet["!ref"]) throw new Error("第一个工作表为空");
    const range = XLSX.utils.decode_range(worksheet["!ref"]);
    if (range.e.r <= range.s.r) throw new Error("工作表没有数据行");
    return { buffer, fileName, workbook, sheetName, worksheet, range };
  }

  function detectHeaderDepth(worksheet, range) {
    const merges = worksheet["!merges"] || [];
    let depth = 1;
    merges.forEach(merge => {
      if (merge.s.r < range.s.r || merge.s.r > range.s.r + 6) return;
      depth = Math.max(depth, merge.e.r - range.s.r + 1);
      if (merge.e.c > merge.s.c) depth = Math.max(depth, merge.s.r - range.s.r + 2);
    });
    return Math.min(8, Math.max(1, depth));
  }

  // 解析多行和合并表头：合并范围内每个列位置都继承左上角表头文字。
  function parseMergedHeaders(worksheet, range, headerDepth) {
    const merges = worksheet["!merges"] || [];
    const matrix = [];
    for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + headerDepth - 1); r++) {
      const row = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const merge = merges.find(m => r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c);
        row.push(cellText(merge ? worksheet[address(merge.s.r, merge.s.c)] : worksheet[address(r, c)]));
      }
      matrix.push(row);
    }
    return matrix;
  }

  // fieldKey 仅由完整表头路径生成，用来跨文件识别同一字段；columnIndex 仍由每个文件独立保存。
  function buildColumnDefinitions(worksheet, range, headerDepth) {
    const headerMatrix = parseMergedHeaders(worksheet, range, headerDepth);
    const definitions = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const values = headerMatrix.map(row => row[c - range.s.c]).filter(Boolean);
      const headerPath = values.filter((value, index) => index === 0 || value !== values[index - 1]);
      if (!headerPath.length) continue;
      definitions.push({
        id: `column-${c}`,
        fieldKey: headerPath.join("\u001f"),
        columnIndex: c,
        columnLetter: XLSX.utils.encode_col(c),
        headerPath,
        displayName: headerPath.join(" > ")
      });
    }
    return definitions;
  }

  function refreshWorkbookStructure(data, headerDepth) {
    data.headerDepth = headerDepth;
    data.columnDefinitions = buildColumnDefinitions(data.worksheet, data.range, headerDepth);
    data.columns = data.columnDefinitions;
    data.dataStartRow = data.range.s.r + headerDepth;
    data.rowCount = Math.max(0, data.range.e.r - data.dataStartRow + 1);
    data.dataRowCount = data.rowCount;
    if (!data.columnDefinitions.length) throw new Error("没有识别到可用字段，请调整表头行数");
    return data;
  }

  // 每个原始文件独立解析；单个失败会保留错误状态，不影响其他文件。
  async function parseSourceFiles(files) {
    const existingKeys = new Set(state.sourceFiles.map(item => item.fileKey));
    const tasks = [...files].map(async file => {
      const fileKey = `${file.name}|${file.size}|${file.lastModified}`;
      if (existingKeys.has(fileKey)) return null;
      existingKeys.add(fileKey);
      const base = { id: makeId(), file, fileKey, fileName: file.name, status: "loading" };
      if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) return { ...base, status: "error", error: "文件格式不支持" };
      try {
        const buffer = await file.arrayBuffer();
        const parsed = parseWorkbook(buffer, file.name);
        parsed.id = base.id; parsed.file = file; parsed.fileKey = fileKey; parsed.status = "success";
        refreshWorkbookStructure(parsed, detectHeaderDepth(parsed.worksheet, parsed.range));
        return parsed;
      } catch (error) {
        return { ...base, status: "error", error: error.message };
      }
    });
    const parsed = (await Promise.all(tasks)).filter(Boolean);
    state.sourceFiles.push(...parsed);
    state.result = null;
    renderSourceFiles(); updateFieldSelectors(); els.results.classList.add("hidden");
    const failed = parsed.filter(item => item.status === "error").length;
    if (failed) showMessage(`${failed} 个原始文件解析失败，其他成功文件仍可继续使用。`, "error");
    else clearMessage();
  }

  async function parseTargetFile(file) {
    clearMessage();
    if (!file || !/\.(xlsx|xlsm|xls)$/i.test(file.name)) { showMessage("模板文件格式错误，请选择 Excel 文件。", "error"); return; }
    try {
      const buffer = await file.arrayBuffer();
      const parsed = parseWorkbook(buffer, file.name);
      refreshWorkbookStructure(parsed, detectHeaderDepth(parsed.worksheet, parsed.range));
      state.target = parsed; state.result = null;
      renderTargetFile(); updateFieldSelectors(); els.results.classList.add("hidden"); clearMessage();
    } catch (error) {
      state.target = null; renderTargetFile(); updateFieldSelectors();
      showMessage(`目标模板表读取失败：${error.message}`, "error");
    }
  }

  function headerDepthOptions(file) {
    const max = Math.min(8, file.range.e.r - file.range.s.r);
    let html = "";
    for (let i = 1; i <= max; i++) html += `<option value="${i}"${i === file.headerDepth ? " selected" : ""}>${i} 行</option>`;
    return html;
  }

  function renderSourceFiles() {
    const total = state.sourceFiles.length;
    const successful = state.sourceFiles.filter(file => file.status === "success");
    const failed = total - successful.length;
    els.sourceCard.classList.toggle("has-file", total > 0);
    els.sourceSummary.classList.toggle("hidden", total === 0);
    els.sourceList.classList.toggle("hidden", total === 0);
    els.sourceSummary.innerHTML = `已导入 <strong>${total}</strong> 个文件 · <span class="ok-text">${successful.length} 个成功</span>${failed ? ` · <span class="error-text">${failed} 个失败</span>` : ""}`;
    els.sourceList.innerHTML = state.sourceFiles.map((file, index) => {
      if (file.status === "error") return `<article class="source-item source-error"><span class="file-status">!</span><div class="source-meta"><strong>${index + 1}. ${escapeHtml(file.fileName)}</strong><small>解析失败 · ${escapeHtml(file.error)}</small></div><button class="remove-file" data-remove-source="${file.id}">移除</button></article>`;
      return `<article class="source-item"><span class="file-status">✓</span><div class="source-meta"><strong>${index + 1}. ${escapeHtml(file.fileName)}</strong><small>${escapeHtml(file.sheetName)} · ${file.rowCount} 条数据 · 已识别 ${file.columnDefinitions.length} 个字段</small></div><label class="inline-depth">表头 <select data-source-depth="${file.id}">${headerDepthOptions(file)}</select></label><button class="remove-file" data-remove-source="${file.id}">移除</button></article>`;
    }).join("");
  }

  function renderTargetFile() {
    if (!state.target) {
      els.targetInfo.classList.add("hidden"); els.targetDepthWrap.classList.add("hidden"); els.targetCard.classList.remove("has-file"); return;
    }
    const target = state.target;
    els.targetInfo.innerHTML = `<span>${escapeHtml(target.fileName)}</span><b>${target.rowCount} 行 · ${target.columnDefinitions.length} 个字段</b>`;
    els.targetInfo.classList.remove("hidden"); els.targetDepthWrap.classList.remove("hidden"); els.targetCard.classList.add("has-file");
    els.targetDepth.innerHTML = headerDepthOptions(target);
  }

  // 统计字段在多少个成功原始文件中“唯一存在”，共同字段排在最前面。
  function getSharedSourceFields(sourceFiles) {
    const readyFiles = sourceFiles.filter(file => file.status === "success");
    const fieldMap = new Map();
    readyFiles.forEach(file => {
      const local = new Map();
      file.columnDefinitions.forEach(def => {
        if (!local.has(def.fieldKey)) local.set(def.fieldKey, []);
        local.get(def.fieldKey).push(def);
      });
      local.forEach((defs, fieldKey) => {
        if (defs.length !== 1) return;
        if (!fieldMap.has(fieldKey)) fieldMap.set(fieldKey, { fieldKey, displayName: defs[0].displayName, fileCount: 0, totalFiles: readyFiles.length });
        fieldMap.get(fieldKey).fileCount++;
      });
    });
    return [...fieldMap.values()].sort((a, b) => b.fileCount - a.fileCount || a.displayName.localeCompare(b.displayName, "zh-CN"));
  }

  function fillSourceSelect(select, fields) {
    const previous = select.value;
    select.innerHTML = ""; select.append(new Option("请选择字段", ""));
    fields.forEach(field => {
      const availability = field.fileCount === field.totalFiles ? `全部 ${field.totalFiles}/${field.totalFiles} 个文件` : `仅 ${field.fileCount}/${field.totalFiles} 个文件`;
      select.append(new Option(`${field.displayName}  [${availability}]`, field.fieldKey));
    });
    select.disabled = !fields.length;
    if ([...select.options].some(option => option.value === previous)) select.value = previous;
  }

  function fillTargetSelect(select, columns) {
    const previous = select.value;
    select.innerHTML = ""; select.append(new Option("请选择字段", ""));
    columns.forEach(col => select.append(new Option(`${col.displayName}  [${col.columnLetter}列]`, String(col.columnIndex))));
    select.disabled = false;
    if ([...select.options].some(option => option.value === previous)) select.value = previous;
  }

  function updateFieldSelectors() {
    const readyFiles = state.sourceFiles.filter(file => file.status === "success");
    const sharedFields = getSharedSourceFields(readyFiles);
    if (readyFiles.length) {
      fillSourceSelect(els.sourceMatch, sharedFields); fillSourceSelect(els.sourceData, sharedFields);
    } else {
      [els.sourceMatch, els.sourceData].forEach(select => { select.innerHTML = '<option value="">请先上传原始文件</option>'; select.disabled = true; });
    }
    if (state.target) {
      fillTargetSelect(els.targetMatch, state.target.columnDefinitions); fillTargetSelect(els.targetFill, state.target.columnDefinitions);
    } else {
      [els.targetMatch, els.targetFill].forEach(select => { select.innerHTML = '<option value="">请先上传模板</option>'; select.disabled = true; });
    }
    updateActionState();
  }

  function updateActionState() {
    const hasSource = state.sourceFiles.some(file => file.status === "success");
    els.fillButton.disabled = !(hasSource && state.target && els.sourceMatch.value && els.targetMatch.value && els.sourceData.value && els.targetFill.value);
  }

  function getTargetColumn(data, select) {
    const columnIndex = Number(select.value);
    return data.columnDefinitions.find(col => col.columnIndex === columnIndex);
  }

  // 在当前文件自己的字段定义中查找真实列位置；不共享 columnIndex。
  function resolveSourceField(sourceFile, selectedField) {
    const matches = sourceFile.columnDefinitions.filter(def => def.fieldKey === selectedField);
    return matches.length === 1 ? matches[0] : null;
  }

  // 为所有成功文件建立统一索引，每条记录保留文件、Sheet、行号和本文件字段定义。
  function buildGlobalMatchIndex(sourceFiles, matchField) {
    const index = new Map();
    const missingFiles = [];
    sourceFiles.filter(file => file.status === "success").forEach(file => {
      const matchColumn = resolveSourceField(file, matchField);
      if (!matchColumn) { missingFiles.push(file); return; }
      for (let rowIndex = file.dataStartRow; rowIndex <= file.range.e.r; rowIndex++) {
        const matchValue = cellText(file.worksheet[address(rowIndex, matchColumn.columnIndex)]);
        if (matchValue === "") continue;
        const record = { fileId: file.id, fileName: file.fileName, sheetName: file.sheetName, rowIndex, sourceFile: file, columnDefinitions: file.columnDefinitions };
        if (!index.has(matchValue)) index.set(matchValue, []);
        index.get(matchValue).push(record);
      }
    });
    return { index, missingFiles };
  }

  // 空字符串和只包含空格的字符串都视为空；0、"0" 和 false 均为有效数据。
  function isEmptyValue(value) {
    return value === undefined || value === null || (typeof value === "string" && value.trim() === "");
  }

  // 生成用于基础一致性判断的标准值：数字字符串与数字可比较，日期统一为 ISO，普通文本去除首尾空格。
  function normalizeComparableValue(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return `date:${value.toISOString()}`;
    if (typeof value === "number") return `number:${Object.is(value, -0) ? 0 : value}`;
    if (typeof value === "boolean") return `boolean:${value}`;
    if (typeof value === "string") {
      const trimmed = value.trim();
      if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(trimmed)) {
        const numeric = Number(trimmed);
        if (Number.isFinite(numeric)) return `number:${Object.is(numeric, -0) ? 0 : numeric}`;
      }
      if (/^\d{4}-\d{1,2}-\d{1,2}(?:[ T].*)?$/.test(trimmed)) {
        const date = new Date(trimmed);
        if (!Number.isNaN(date.getTime())) return `date:${date.toISOString()}`;
      }
      return `string:${trimmed}`;
    }
    return `${typeof value}:${String(value)}`;
  }

  function sourceLocation(record) { return `${record.fileName} · ${record.sheetName} · 第 ${record.rowIndex + 1} 行`; }

  // 多条匹配记录先按来源字段是否有值筛选；只有不同的非空值才是真正的数据冲突。
  function resolveMatchedValue(matches, sourceField) {
    const candidates = matches.map(record => {
      const sourceColumn = resolveSourceField(record.sourceFile, sourceField);
      const sourceCell = sourceColumn ? record.sourceFile.worksheet[address(record.rowIndex, sourceColumn.columnIndex)] : undefined;
      const value = sourceCell?.v;
      return {
        record,
        sourceCell,
        hasField: Boolean(sourceColumn),
        value,
        displayValue: isEmptyValue(value) ? "" : cellText(sourceCell),
        empty: isEmptyValue(value),
        comparableValue: isEmptyValue(value) ? null : normalizeComparableValue(value)
      };
    });
    const nonEmpty = candidates.filter(candidate => !candidate.empty);
    if (!nonEmpty.length) {
      return { status: "empty", value: null, sources: matches.map(sourceLocation), candidates, nonEmpty };
    }
    const uniqueComparableValues = new Set(nonEmpty.map(candidate => candidate.comparableValue));
    if (uniqueComparableValues.size > 1) {
      return { status: "conflict", value: null, sources: nonEmpty.map(candidate => sourceLocation(candidate.record)), candidates, nonEmpty };
    }
    const chosen = nonEmpty[0];
    const resolutionType = matches.length === 1 ? "single" : nonEmpty.length === 1 ? "unique_non_empty" : "consistent";
    return {
      status: "filled",
      value: chosen.value,
      sourceCell: chosen.sourceCell,
      sources: nonEmpty.map(candidate => sourceLocation(candidate.record)),
      candidates,
      nonEmpty,
      resolutionType
    };
  }

  function writeCellValue(worksheet, row, col, sourceCell) {
    const cellAddress = address(row, col);
    const targetCell = worksheet[cellAddress] || {};
    targetCell.v = sourceCell.v;
    targetCell.t = sourceCell.t || (typeof sourceCell.v === "number" ? "n" : "s");
    delete targetCell.f; delete targetCell.F; delete targetCell.w; delete targetCell.h; delete targetCell.r;
    worksheet[cellAddress] = targetCell;
  }

  // 遍历模板并查询全局索引。重复记录会交给 resolveMatchedValue 按非空值优先级处理。
  function fillTemplate(sourceFiles, target, mapping) {
    const globalIndex = buildGlobalMatchIndex(sourceFiles, mapping.sourceMatchField);
    const missingDataFiles = sourceFiles.filter(file => file.status === "success" && !resolveSourceField(file, mapping.fillMappings[0].sourceField));
    const records = [];
    let filled = 0, singleSource = 0, autoSelected = 0, multiSourceConsistent = 0;
    let unmatched = 0, dataConflicts = 0, allSourcesEmpty = 0, targetEmpty = 0;

    for (let targetRow = target.dataStartRow; targetRow <= target.range.e.r; targetRow++) {
      const matchValue = cellText(target.worksheet[address(targetRow, mapping.targetMatch.columnIndex)]);
      const preview = { targetRow, matchValue, status: "", sourceValue: "", writtenValue: "", sourceInfo: "" };
      if (matchValue === "") {
        preview.status = "匹配字段为空"; targetEmpty++;
      } else {
        const matches = globalIndex.index.get(matchValue) || [];
        if (!matches.length) {
          preview.status = "未匹配"; unmatched++;
        } else {
          const resolved = resolveMatchedValue(matches, mapping.fillMappings[0].sourceField);
          if (resolved.status === "empty") {
            preview.status = "来源数据为空"; allSourcesEmpty++;
            preview.sourceInfo = `${matches.length} 个来源：${resolved.sources.join("；")}`;
          } else if (resolved.status === "conflict") {
            preview.status = "数据冲突"; dataConflicts++;
            preview.sourceInfo = resolved.nonEmpty.map(candidate => `${sourceLocation(candidate.record)} → ${candidate.displayValue}`).join("；");
          } else {
            writeCellValue(target.worksheet, targetRow, mapping.fillMappings[0].targetFill.columnIndex, resolved.sourceCell);
            preview.sourceValue = cellText(resolved.sourceCell);
            preview.writtenValue = cellText(target.worksheet[address(targetRow, mapping.fillMappings[0].targetFill.columnIndex)]);
            preview.sourceInfo = resolved.sources.join("；");
            filled++;
            if (resolved.resolutionType === "single") {
              preview.status = "已填充"; singleSource++;
            } else if (resolved.resolutionType === "unique_non_empty") {
              preview.status = "已填充（自动选择唯一有值来源）"; autoSelected++;
            } else {
              preview.status = "已填充（多来源一致）"; multiSourceConsistent++;
            }
          }
        }
      }
      records.push(preview);
    }
    return { records, filled, singleSource, autoSelected, multiSourceConsistent, unmatched, dataConflicts, allSourcesEmpty, targetEmpty, missingMatchFiles: globalIndex.missingFiles, missingDataFiles };
  }

  function startFill() {
    clearMessage();
    if (els.fillButton.disabled) { showMessage("请完整选择匹配字段和填充字段。", "error"); return; }
    try {
      const cleanTarget = parseWorkbook(state.target.buffer, state.target.fileName);
      refreshWorkbookStructure(cleanTarget, state.target.headerDepth);
      const mapping = {
        sourceMatchField: els.sourceMatch.value,
        targetMatch: getTargetColumn(cleanTarget, els.targetMatch),
        fillMappings: [{ sourceField: els.sourceData.value, targetFill: getTargetColumn(cleanTarget, els.targetFill) }]
      };
      const summary = fillTemplate(state.sourceFiles, cleanTarget, mapping);
      state.result = { workbook: cleanTarget.workbook, target: cleanTarget, mapping, ...summary };
      renderResults();
    } catch (error) { showMessage(`处理失败：${error.message}`, "error"); }
  }

  function renderResults() {
    const result = state.result;
    const readyFiles = state.sourceFiles.filter(file => file.status === "success");
    $("sourceFileCount").textContent = readyFiles.length;
    $("sourceRowCount").textContent = readyFiles.reduce((sum, file) => sum + file.rowCount, 0);
    $("templateCount").textContent = result.records.length;
    $("filledCount").textContent = result.filled;
    $("singleSourceCount").textContent = result.singleSource;
    $("autoSelectedCount").textContent = result.autoSelected;
    $("consistentCount").textContent = result.multiSourceConsistent;
    $("unmatchedCount").textContent = result.unmatched;
    $("conflictCount").textContent = result.dataConflicts;
    $("sourceEmptyCount").textContent = result.allSourcesEmpty;
    const exceptions = result.records.length - result.filled;
    const badge = $("resultBadge");
    badge.textContent = exceptions ? `已完成，${exceptions} 行未填写` : "全部填写完成";
    badge.className = exceptions ? "badge warn" : "badge";

    const notices = [];
    if (result.missingMatchFiles.length) notices.push(`${result.missingMatchFiles.length} 个原始文件不存在所选匹配字段：${result.missingMatchFiles.map(f => f.fileName).join("、")}`);
    if (result.missingDataFiles.length) notices.push(`${result.missingDataFiles.length} 个原始文件不存在所选来源字段：${result.missingDataFiles.map(f => f.fileName).join("、")}`);
    if (result.targetEmpty) notices.push(`${result.targetEmpty} 个模板行的匹配字段为空，已跳过`);
    els.resultNotice.classList.toggle("hidden", !notices.length);
    els.resultNotice.innerHTML = notices.map(text => `<p>${escapeHtml(text)}</p>`).join("");

    els.exportButton.disabled = false;
    renderPreview(result.records);
    els.results.classList.remove("hidden");
    els.results.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function renderPreview(records) {
    const rows = records.slice(0, 100).map(record => {
      const statusClass = record.status.startsWith("已填充") ? "status-success" : record.status === "数据冲突" ? "status-duplicate" : "status-unmatched";
      const finalValue = record.writtenValue === "" ? "—" : escapeHtml(record.writtenValue);
      return `<tr><td>${escapeHtml(record.matchValue || "（空）")}</td><td class="${statusClass}">${escapeHtml(record.status)}</td><td>${finalValue}</td><td class="source-trace">${record.sourceInfo ? escapeHtml(record.sourceInfo) : "—"}</td></tr>`;
    }).join("");
    els.preview.innerHTML = `<thead><tr><th>匹配值</th><th>状态</th><th>最终采用值</th><th>来源</th></tr></thead><tbody>${rows}</tbody>`;
  }

  function exportWorkbook() {
    if (!state.result) return;
    const base = state.target.fileName.replace(/\.(xlsx|xlsm|xls)$/i, "");
    const ext = /\.xlsm$/i.test(state.target.fileName) ? "xlsm" : "xlsx";
    XLSX.writeFile(state.result.workbook, `${base}_已填写.${ext}`, { bookType: ext, cellStyles: true, bookVBA: ext === "xlsm" });
  }

  function removeSourceFile(id) {
    state.sourceFiles = state.sourceFiles.filter(file => file.id !== id);
    state.result = null; els.results.classList.add("hidden");
    renderSourceFiles(); updateFieldSelectors(); clearMessage();
  }

  function changeSourceHeaderDepth(id, value) {
    const file = state.sourceFiles.find(item => item.id === id && item.status === "success");
    if (!file) return;
    try {
      refreshWorkbookStructure(file, Number(value)); state.result = null; els.results.classList.add("hidden");
      renderSourceFiles(); updateFieldSelectors(); clearMessage();
    } catch (error) { showMessage(`${file.fileName}：${error.message}`, "error"); }
  }

  function changeTargetHeaderDepth(value) {
    if (!state.target) return;
    try {
      refreshWorkbookStructure(state.target, Number(value)); state.result = null; els.results.classList.add("hidden");
      renderTargetFile(); updateFieldSelectors(); clearMessage();
    } catch (error) { showMessage(error.message, "error"); }
  }

  window.ExcelTemplateTool = Object.freeze({ parseWorkbook, parseSourceFiles, parseMergedHeaders, buildColumnDefinitions, getSharedSourceFields, buildGlobalMatchIndex, resolveSourceField, isEmptyValue, normalizeComparableValue, resolveMatchedValue, fillTemplate, exportWorkbook });

  els.sourceFile.addEventListener("change", async () => { await parseSourceFiles(els.sourceFile.files); els.sourceFile.value = ""; });
  els.targetFile.addEventListener("change", () => parseTargetFile(els.targetFile.files[0]));
  [["source", $("sourceDrop")], ["target", $("targetDrop")]].forEach(([kind, drop]) => {
    ["dragenter", "dragover"].forEach(event => drop.addEventListener(event, e => { e.preventDefault(); drop.classList.add("drag"); }));
    ["dragleave", "drop"].forEach(event => drop.addEventListener(event, e => { e.preventDefault(); drop.classList.remove("drag"); }));
    drop.addEventListener("drop", e => kind === "source" ? parseSourceFiles(e.dataTransfer.files) : parseTargetFile(e.dataTransfer.files[0]));
  });
  els.sourceList.addEventListener("click", event => { const button = event.target.closest("[data-remove-source]"); if (button) removeSourceFile(button.dataset.removeSource); });
  els.sourceList.addEventListener("change", event => { const select = event.target.closest("[data-source-depth]"); if (select) changeSourceHeaderDepth(select.dataset.sourceDepth, select.value); });
  els.targetDepth.addEventListener("change", () => changeTargetHeaderDepth(els.targetDepth.value));
  [els.sourceMatch, els.targetMatch, els.sourceData, els.targetFill].forEach(select => select.addEventListener("change", updateActionState));
  els.fillButton.addEventListener("click", startFill);
  els.exportButton.addEventListener("click", exportWorkbook);
})();
