(() => {
  "use strict";

  const MAX_MAPPINGS = 20;
  const state = {
    sourceFiles: [],
    target: null,
    result: null,
    fieldMappings: [{ id: makeId("mapping"), sourceFieldKey: "", targetFieldKey: "" }]
  };
  const $ = id => document.getElementById(id);
  const els = {
    sourceFile: $("sourceFile"), targetFile: $("targetFile"), sourceCard: $("sourceCard"), targetCard: $("targetCard"),
    sourceSummary: $("sourceSummary"), sourceList: $("sourceList"), targetInfo: $("targetInfo"),
    targetDepth: $("targetHeaderDepth"), targetDepthWrap: $("targetDepthWrap"), sourceMatch: $("sourceMatchField"),
    targetMatch: $("targetMatchField"), mappingList: $("mappingList"), addMapping: $("addMappingButton"),
    fillButton: $("fillButton"), exportButton: $("exportButton"), message: $("globalMessage"), results: $("results"),
    preview: $("previewTable"), resultNotice: $("resultNotice"), mappingStats: $("mappingStatsTable")
  };

  function makeId(prefix = "file") { return globalThis.crypto?.randomUUID?.() || `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`; }
  function showMessage(text, type = "warning") { els.message.textContent = text; els.message.className = `message ${type === "error" ? "error" : ""}`; }
  function clearMessage() { els.message.className = "message hidden"; els.message.textContent = ""; }
  function cellText(cell) { return cell ? XLSX.utils.format_cell(cell) : ""; }
  function address(row, col) { return XLSX.utils.encode_cell({ r: row, c: col }); }
  function escapeHtml(value) { return String(value ?? "").replace(/[&<>'"]/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[ch])); }
  function isEtFile(fileName) { return /\.et$/i.test(fileName); }
  function fileTypeLabel(fileName) { return isEtFile(fileName) ? "ET 文件" : /\.xlsm$/i.test(fileName) ? "XLSM 文件" : /\.xls$/i.test(fileName) ? "XLS 文件" : "XLSX 文件"; }
  function parseErrorMessage(fileName, error) {
    console.error(`文件解析失败：${fileName}`, error);
    return isEtFile(fileName) ? "该 ET 文件当前无法直接解析，请在 WPS 中另存为 XLSX 后重新上传" : error.message;
  }

  // 不依赖扩展名判断内容格式：包括 ET 在内的文件都会实际交给 SheetJS 尝试读取。
  function parseWorkbook(buffer, fileName) {
    const workbook = XLSX.read(buffer, { type: "array", cellStyles: true, cellDates: true, cellNF: true, bookVBA: true });
    if (!workbook.SheetNames.length) throw new Error("文件中没有可读取的工作表");
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

  // 解析多行和合并表头：合并区域内的列会继承左上角单元格文字。
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

  // fieldKey 用于跨文件识别字段；真实 columnIndex 仍由每个文件独立保存。
  function buildColumnDefinitions(worksheet, range, headerDepth) {
    const headerMatrix = parseMergedHeaders(worksheet, range, headerDepth);
    const definitions = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const values = headerMatrix.map(row => row[c - range.s.c]).filter(Boolean);
      const headerPath = values.filter((value, index) => index === 0 || value !== values[index - 1]);
      if (!headerPath.length) continue;
      definitions.push({
        id: `column-${c}`,
        fieldKey: JSON.stringify(headerPath),
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
    if (!data.columnDefinitions.length) throw new Error("没有识别到可用字段，请调整表头行数");
    return data;
  }

  // 批量解析原始文件。某一个文件失败时只记录该文件错误，其他文件继续使用。
  async function parseSourceFiles(files) {
    const existingKeys = new Set(state.sourceFiles.map(item => item.fileKey));
    const tasks = [...files].map(async file => {
      const fileKey = `${file.name}|${file.size}|${file.lastModified}`;
      if (existingKeys.has(fileKey)) return null;
      existingKeys.add(fileKey);
      const base = { id: makeId(), file, fileKey, fileName: file.name, status: "loading", fileType: fileTypeLabel(file.name) };
      if (!/\.(xlsx|xlsm|xls|et)$/i.test(file.name)) return { ...base, status: "error", error: "文件格式不支持" };
      try {
        const parsed = parseWorkbook(await file.arrayBuffer(), file.name);
        Object.assign(parsed, base, { status: "success" });
        refreshWorkbookStructure(parsed, detectHeaderDepth(parsed.worksheet, parsed.range));
        return parsed;
      } catch (error) {
        return { ...base, status: "error", error: parseErrorMessage(file.name, error) };
      }
    });
    const parsed = (await Promise.all(tasks)).filter(Boolean);
    state.sourceFiles.push(...parsed); state.result = null;
    renderSourceFiles(); updateFieldSelectors(); els.results.classList.add("hidden");
    const failed = parsed.filter(item => item.status === "error").length;
    if (failed) showMessage(`${failed} 个原始文件解析失败，其他成功文件仍可继续使用。`, "error"); else clearMessage();
  }

  async function parseTargetFile(file) {
    clearMessage();
    if (!file || !/\.(xlsx|xlsm|xls|et)$/i.test(file.name)) { showMessage("模板文件格式错误，请选择 .xlsx、.xls、.xlsm 或 .et 文件。", "error"); return; }
    try {
      const parsed = parseWorkbook(await file.arrayBuffer(), file.name);
      parsed.fileType = fileTypeLabel(file.name);
      refreshWorkbookStructure(parsed, detectHeaderDepth(parsed.worksheet, parsed.range));
      state.target = parsed; state.result = null;
      renderTargetFile(); updateFieldSelectors(); els.results.classList.add("hidden"); clearMessage();
    } catch (error) {
      state.target = null; renderTargetFile(); updateFieldSelectors();
      showMessage(`目标模板表读取失败：${parseErrorMessage(file.name, error)}`, "error");
    }
  }

  function headerDepthOptions(file) {
    let html = "";
    for (let i = 1; i <= Math.min(8, file.range.e.r - file.range.s.r); i++) html += `<option value="${i}"${i === file.headerDepth ? " selected" : ""}>${i} 行</option>`;
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
      if (file.status === "error") return `<article class="source-item source-error"><span class="file-status">!</span><div class="source-meta"><strong>${index + 1}. ${escapeHtml(file.fileName)}</strong><small>${escapeHtml(file.fileType)} · 解析失败 · ${escapeHtml(file.error)}</small></div><button class="remove-file" data-remove-source="${file.id}">移除</button></article>`;
      return `<article class="source-item"><span class="file-status">✓</span><div class="source-meta"><strong>${index + 1}. ${escapeHtml(file.fileName)}</strong><small>${escapeHtml(file.fileType)} · 解析成功 · ${escapeHtml(file.sheetName)} · ${file.rowCount} 条数据</small></div><label class="inline-depth">表头 <select data-source-depth="${file.id}">${headerDepthOptions(file)}</select></label><button class="remove-file" data-remove-source="${file.id}">移除</button></article>`;
    }).join("");
  }

  function renderTargetFile() {
    if (!state.target) {
      els.targetInfo.classList.add("hidden"); els.targetDepthWrap.classList.add("hidden"); els.targetCard.classList.remove("has-file"); return;
    }
    const target = state.target;
    els.targetInfo.innerHTML = `<span>${escapeHtml(target.fileName)} · ${escapeHtml(target.fileType)}</span><b>${target.rowCount} 行 · ${target.columnDefinitions.length} 个字段</b>`;
    els.targetInfo.classList.remove("hidden"); els.targetDepthWrap.classList.remove("hidden"); els.targetCard.classList.add("has-file");
    els.targetDepth.innerHTML = headerDepthOptions(target);
  }

  // 统计字段在多少个成功原始文件中唯一存在，共同字段优先显示。
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

  function sourceOptions(fields, selected = "") {
    return `<option value="">请选择原始字段</option>` + fields.map(field => {
      const availability = field.fileCount === field.totalFiles ? `全部 ${field.totalFiles}/${field.totalFiles}` : `仅 ${field.fileCount}/${field.totalFiles}`;
      return `<option value="${escapeHtml(field.fieldKey)}"${field.fieldKey === selected ? " selected" : ""}>${escapeHtml(field.displayName)} [${availability}]</option>`;
    }).join("");
  }

  function targetOptions(columns, selected = "") {
    return `<option value="">请选择模板字段</option>` + columns.map(col => `<option value="${col.id}"${col.id === selected ? " selected" : ""}>${escapeHtml(col.displayName)} [${col.columnLetter}列]</option>`).join("");
  }

  function fillSourceSelect(select, fields) {
    const previous = select.value;
    select.innerHTML = sourceOptions(fields, previous); select.disabled = !fields.length;
  }

  function fillTargetSelect(select, columns) {
    const previous = select.value;
    select.innerHTML = `<option value="">请选择字段</option>` + columns.map(col => `<option value="${col.id}"${col.id === previous ? " selected" : ""}>${escapeHtml(col.displayName)} [${col.columnLetter}列]</option>`).join("");
    select.disabled = false;
  }

  function renderFieldMappings() {
    const fields = getSharedSourceFields(state.sourceFiles);
    const targetColumns = state.target?.columnDefinitions || [];
    const sourceKeys = new Set(fields.map(field => field.fieldKey));
    const targetKeys = new Set(targetColumns.map(column => column.id));
    state.fieldMappings.forEach(mapping => {
      if (!sourceKeys.has(mapping.sourceFieldKey)) mapping.sourceFieldKey = "";
      if (!targetKeys.has(mapping.targetFieldKey)) mapping.targetFieldKey = "";
    });
    els.mappingList.innerHTML = state.fieldMappings.map((mapping, index) => `
      <article class="mapping-row" data-mapping-id="${mapping.id}">
        <span class="mapping-number">第 ${index + 1} 组</span>
        <label>原始字段<select data-mapping-source="${mapping.id}"${fields.length ? "" : " disabled"}>${sourceOptions(fields, mapping.sourceFieldKey)}</select></label>
        <span class="mapping-arrow">→</span>
        <label>模板字段<select data-mapping-target="${mapping.id}"${targetColumns.length ? "" : " disabled"}>${targetOptions(targetColumns, mapping.targetFieldKey)}</select></label>
        <button type="button" class="delete-mapping" data-delete-mapping="${mapping.id}"${state.fieldMappings.length === 1 ? " disabled" : ""}>删除</button>
      </article>`).join("");
    els.addMapping.disabled = state.fieldMappings.length >= MAX_MAPPINGS;
    els.addMapping.textContent = state.fieldMappings.length >= MAX_MAPPINGS ? "已达到 20 组上限" : "＋ 添加一组填充映射";
  }

  function updateFieldSelectors() {
    const readyFiles = state.sourceFiles.filter(file => file.status === "success");
    const fields = getSharedSourceFields(readyFiles);
    if (readyFiles.length) fillSourceSelect(els.sourceMatch, fields);
    else { els.sourceMatch.innerHTML = '<option value="">请先上传原始文件</option>'; els.sourceMatch.disabled = true; }
    if (state.target) fillTargetSelect(els.targetMatch, state.target.columnDefinitions);
    else { els.targetMatch.innerHTML = '<option value="">请先上传模板</option>'; els.targetMatch.disabled = true; }
    renderFieldMappings(); updateActionState();
  }

  function updateActionState() {
    const hasSource = state.sourceFiles.some(file => file.status === "success");
    els.fillButton.disabled = !(hasSource && state.target && els.sourceMatch.value && els.targetMatch.value && state.fieldMappings.length);
  }

  function addFieldMapping() {
    if (state.fieldMappings.length >= MAX_MAPPINGS) { showMessage("最多只能添加 20 组填充映射。", "error"); return; }
    state.fieldMappings.push({ id: makeId("mapping"), sourceFieldKey: "", targetFieldKey: "" });
    renderFieldMappings(); clearMessage();
  }

  function deleteFieldMapping(id) {
    if (state.fieldMappings.length <= 1) { showMessage("至少需要保留 1 组填充映射。", "error"); return; }
    state.fieldMappings = state.fieldMappings.filter(mapping => mapping.id !== id);
    state.result = null; els.results.classList.add("hidden"); renderFieldMappings(); clearMessage();
  }

  // 开始处理前检查每组是否完整，并禁止多个来源字段写入同一个模板目标列。
  function validateFieldMappings(fieldMappings) {
    if (!fieldMappings.length) return { valid: false, error: "至少需要 1 组填充映射。" };
    for (let i = 0; i < fieldMappings.length; i++) {
      if (!fieldMappings[i].sourceFieldKey || !fieldMappings[i].targetFieldKey) return { valid: false, error: `第 ${i + 1} 组填充映射未完整选择原始字段和模板字段。` };
    }
    const targets = fieldMappings.map(mapping => mapping.targetFieldKey);
    if (new Set(targets).size !== targets.length) return { valid: false, error: "模板目标字段被重复使用，请为每组映射选择不同的模板字段。" };
    return { valid: true, error: "" };
  }

  function getTargetColumnByKey(data, key) { return data.columnDefinitions.find(col => col.id === key); }

  // 在每个文件自己的字段定义中寻找真实列位置，不共享 columnIndex。
  function resolveSourceField(sourceFile, selectedField) {
    const matches = sourceFile.columnDefinitions.filter(def => def.fieldKey === selectedField);
    return matches.length === 1 ? matches[0] : null;
  }

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

  function isEmptyValue(value) { return value === undefined || value === null || (typeof value === "string" && value.trim() === ""); }

  function normalizeComparableValue(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return `date:${value.toISOString()}`;
    if (typeof value === "number") return `number:${Object.is(value, -0) ? 0 : value}`;
    if (typeof value === "boolean") return `boolean:${value}`;
    if (typeof value === "string") {
      const trimmed = value.trim();
      if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(trimmed) && Number.isFinite(Number(trimmed))) return `number:${Object.is(Number(trimmed), -0) ? 0 : Number(trimmed)}`;
      if (/^\d{4}-\d{1,2}-\d{1,2}(?:[ T].*)?$/.test(trimmed)) {
        const date = new Date(trimmed); if (!Number.isNaN(date.getTime())) return `date:${date.toISOString()}`;
      }
      return `string:${trimmed}`;
    }
    return `${typeof value}:${String(value)}`;
  }

  function sourceLocation(record) { return `${record.fileName} · ${record.sheetName} · 第 ${record.rowIndex + 1} 行`; }

  // 针对单个来源字段独立解析多个候选，因此同一对象的不同字段可以来自不同文件。
  function resolveFieldValue(matches, sourceFieldKey) {
    const candidates = matches.map(record => {
      const sourceColumn = resolveSourceField(record.sourceFile, sourceFieldKey);
      const sourceCell = sourceColumn ? record.sourceFile.worksheet[address(record.rowIndex, sourceColumn.columnIndex)] : undefined;
      const value = sourceCell?.v;
      return { record, sourceCell, hasField: Boolean(sourceColumn), value, displayValue: isEmptyValue(value) ? "" : cellText(sourceCell), empty: isEmptyValue(value), comparableValue: isEmptyValue(value) ? null : normalizeComparableValue(value) };
    });
    const nonEmpty = candidates.filter(candidate => !candidate.empty);
    if (!nonEmpty.length) return { status: "empty", value: null, sources: matches.map(sourceLocation), candidates, nonEmpty };
    if (new Set(nonEmpty.map(candidate => candidate.comparableValue)).size > 1) return { status: "conflict", value: null, sources: nonEmpty.map(candidate => sourceLocation(candidate.record)), candidates, nonEmpty };
    const chosen = nonEmpty[0];
    return {
      status: "filled",
      value: chosen.value,
      sourceCell: chosen.sourceCell,
      sources: nonEmpty.map(candidate => sourceLocation(candidate.record)),
      candidates,
      nonEmpty,
      resolutionType: matches.length === 1 ? "single" : nonEmpty.length === 1 ? "unique_non_empty" : "consistent"
    };
  }
  const resolveMatchedValue = resolveFieldValue;

  function writeCellValue(worksheet, row, col, sourceCell) {
    const cellAddress = address(row, col);
    const targetCell = worksheet[cellAddress] || {};
    targetCell.v = sourceCell.v;
    targetCell.t = sourceCell.t || (typeof sourceCell.v === "number" ? "n" : "s");
    delete targetCell.f; delete targetCell.F; delete targetCell.w; delete targetCell.h; delete targetCell.r;
    worksheet[cellAddress] = targetCell;
  }

  function fieldResultBase(mapping, sourceFieldName) {
    return { mappingId: mapping.id, sourceField: sourceFieldName, targetField: mapping.targetFill.displayName, status: "", value: "", sources: [], sourceInfo: "", resolutionType: "" };
  }

  // 对模板中的一行逐组处理。每个映射都单独解析来源值，某一字段冲突不会阻止其他字段写入。
  function fillTemplateFields(templateRow, matches, fieldMappings, target, sourceFieldNames) {
    return fieldMappings.map(mapping => {
      const result = fieldResultBase(mapping, sourceFieldNames.get(mapping.sourceFieldKey) || mapping.sourceFieldKey);
      const resolved = resolveFieldValue(matches, mapping.sourceFieldKey);
      if (resolved.status === "empty") {
        result.status = "来源数据为空"; result.sources = resolved.sources; result.sourceInfo = `${matches.length} 个来源：${resolved.sources.join("；")}`;
      } else if (resolved.status === "conflict") {
        result.status = "数据冲突"; result.sources = resolved.sources;
        result.sourceInfo = resolved.nonEmpty.map(candidate => `${sourceLocation(candidate.record)} → ${candidate.displayValue}`).join("；");
      } else {
        writeCellValue(target.worksheet, templateRow, mapping.targetFill.columnIndex, resolved.sourceCell);
        result.value = cellText(target.worksheet[address(templateRow, mapping.targetFill.columnIndex)]);
        result.sources = resolved.sources; result.sourceInfo = resolved.sources.join("；"); result.resolutionType = resolved.resolutionType;
        result.status = resolved.resolutionType === "single" ? "已填充" : resolved.resolutionType === "unique_non_empty" ? "已填充（自动选择唯一有值来源）" : "已填充（多来源一致）";
      }
      return result;
    });
  }

  function emptyFieldResults(fieldMappings, sourceFieldNames, status) {
    return fieldMappings.map(mapping => ({ ...fieldResultBase(mapping, sourceFieldNames.get(mapping.sourceFieldKey) || mapping.sourceFieldKey), status }));
  }

  function createMappingStats(fieldMappings, sourceFieldNames) {
    return new Map(fieldMappings.map(mapping => [mapping.id, { mappingId: mapping.id, sourceField: sourceFieldNames.get(mapping.sourceFieldKey) || mapping.sourceFieldKey, targetField: mapping.targetFill.displayName, filled: 0, empty: 0, conflict: 0, unmatched: 0, skipped: 0 }]));
  }

  function fillTemplate(sourceFiles, target, config) {
    const globalIndex = buildGlobalMatchIndex(sourceFiles, config.sourceMatchField);
    const sourceFields = getSharedSourceFields(sourceFiles);
    const sourceFieldNames = new Map(sourceFields.map(field => [field.fieldKey, field.displayName]));
    const mappingStats = createMappingStats(config.fieldMappings, sourceFieldNames);
    const records = [];
    let filled = 0, singleSource = 0, autoSelected = 0, multiSourceConsistent = 0;
    let allSourcesEmpty = 0, dataConflicts = 0, unmatched = 0, targetEmpty = 0;

    for (let targetRow = target.dataStartRow; targetRow <= target.range.e.r; targetRow++) {
      const matchValue = cellText(target.worksheet[address(targetRow, config.targetMatch.columnIndex)]);
      let fieldResults;
      if (matchValue === "") {
        targetEmpty++;
        fieldResults = emptyFieldResults(config.fieldMappings, sourceFieldNames, "匹配字段为空");
      } else {
        const matches = globalIndex.index.get(matchValue) || [];
        if (!matches.length) {
          unmatched++;
          fieldResults = emptyFieldResults(config.fieldMappings, sourceFieldNames, "未匹配");
        } else {
          fieldResults = fillTemplateFields(targetRow, matches, config.fieldMappings, target, sourceFieldNames);
        }
      }

      fieldResults.forEach(result => {
        const stats = mappingStats.get(result.mappingId);
        if (result.status.startsWith("已填充")) {
          filled++; stats.filled++;
          if (result.resolutionType === "single") singleSource++;
          else if (result.resolutionType === "unique_non_empty") autoSelected++;
          else if (result.resolutionType === "consistent") multiSourceConsistent++;
        } else if (result.status === "来源数据为空") { allSourcesEmpty++; stats.empty++; }
        else if (result.status === "数据冲突") { dataConflicts++; stats.conflict++; }
        else if (result.status === "未匹配") stats.unmatched++;
        else stats.skipped++;
      });
      records.push({ targetRow, matchValue, fieldResults });
    }

    const missingDataByMapping = config.fieldMappings.map(mapping => ({
      mapping,
      sourceField: sourceFieldNames.get(mapping.sourceFieldKey) || mapping.sourceFieldKey,
      files: sourceFiles.filter(file => file.status === "success" && !resolveSourceField(file, mapping.sourceFieldKey))
    })).filter(item => item.files.length);

    return {
      records, mappingStats: [...mappingStats.values()], filled, singleSource, autoSelected, multiSourceConsistent,
      allSourcesEmpty, dataConflicts, unmatched, targetEmpty, missingMatchFiles: globalIndex.missingFiles, missingDataByMapping
    };
  }

  function startFill() {
    clearMessage();
    if (els.fillButton.disabled) { showMessage("请先上传文件并选择匹配字段。", "error"); return; }
    const validation = validateFieldMappings(state.fieldMappings);
    if (!validation.valid) { showMessage(validation.error, "error"); return; }
    try {
      const cleanTarget = parseWorkbook(state.target.buffer, state.target.fileName);
      refreshWorkbookStructure(cleanTarget, state.target.headerDepth);
      const fieldMappings = state.fieldMappings.map(mapping => ({
        id: mapping.id,
        sourceFieldKey: mapping.sourceFieldKey,
        targetFieldKey: mapping.targetFieldKey,
        targetFill: getTargetColumnByKey(cleanTarget, mapping.targetFieldKey)
      }));
      const config = { sourceMatchField: els.sourceMatch.value, targetMatch: getTargetColumnByKey(cleanTarget, els.targetMatch.value), fieldMappings };
      const summary = fillTemplate(state.sourceFiles, cleanTarget, config);
      state.result = { workbook: cleanTarget.workbook, target: cleanTarget, config, ...summary };
      renderResults();
    } catch (error) { showMessage(`处理失败：${error.message}`, "error"); }
  }

  function renderResults() {
    const result = state.result;
    const readyFiles = state.sourceFiles.filter(file => file.status === "success");
    $("sourceFileCount").textContent = readyFiles.length;
    $("sourceRowCount").textContent = readyFiles.reduce((sum, file) => sum + file.rowCount, 0);
    $("templateCount").textContent = result.records.length;
    $("mappingCount").textContent = result.config.fieldMappings.length;
    $("theoreticalCount").textContent = result.records.length * result.config.fieldMappings.length;
    $("filledCount").textContent = result.filled;
    $("singleSourceCount").textContent = result.singleSource;
    $("autoSelectedCount").textContent = result.autoSelected;
    $("consistentCount").textContent = result.multiSourceConsistent;
    $("unmatchedCount").textContent = result.unmatched;
    $("conflictCount").textContent = result.dataConflicts;
    $("sourceEmptyCount").textContent = result.allSourcesEmpty;
    const unfilledCells = result.records.length * result.config.fieldMappings.length - result.filled;
    const badge = $("resultBadge");
    badge.textContent = unfilledCells ? `已完成，${unfilledCells} 格未填写` : "全部填写完成";
    badge.className = unfilledCells ? "badge warn" : "badge";

    const notices = [];
    if (result.missingMatchFiles.length) notices.push(`${result.missingMatchFiles.length} 个原始文件不存在所选匹配字段：${result.missingMatchFiles.map(f => f.fileName).join("、")}`);
    result.missingDataByMapping.forEach(item => notices.push(`来源字段“${item.sourceField}”在 ${item.files.length} 个文件中不存在：${item.files.map(f => f.fileName).join("、")}`));
    if (result.targetEmpty) notices.push(`${result.targetEmpty} 个模板行的匹配字段为空，已跳过所有映射`);
    els.resultNotice.classList.toggle("hidden", !notices.length);
    els.resultNotice.innerHTML = notices.map(text => `<p>${escapeHtml(text)}</p>`).join("");
    renderMappingStats(result.mappingStats); renderPreview(result.records);
    els.exportButton.disabled = false; els.results.classList.remove("hidden");
    els.results.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function renderMappingStats(stats) {
    const rows = stats.map(item => `<tr><td>${escapeHtml(item.sourceField)} → ${escapeHtml(item.targetField)}</td><td class="status-success">${item.filled}</td><td>${item.empty}</td><td class="status-duplicate">${item.conflict}</td><td>${item.unmatched}</td></tr>`).join("");
    els.mappingStats.innerHTML = `<thead><tr><th>填充映射</th><th>成功</th><th>为空</th><th>冲突</th><th>未匹配</th></tr></thead><tbody>${rows}</tbody>`;
  }

  function renderPreview(records) {
    const flat = records.flatMap(record => record.fieldResults.map(result => ({ matchValue: record.matchValue, ...result }))).slice(0, 300);
    const rows = flat.map(result => {
      const statusClass = result.status.startsWith("已填充") ? "status-success" : result.status === "数据冲突" ? "status-duplicate" : "status-unmatched";
      return `<tr><td>${escapeHtml(result.matchValue || "（空）")}</td><td>${escapeHtml(result.sourceField)} → ${escapeHtml(result.targetField)}</td><td>${result.value === "" ? "—" : escapeHtml(result.value)}</td><td class="${statusClass}">${escapeHtml(result.status)}</td><td class="source-trace">${result.sourceInfo ? escapeHtml(result.sourceInfo) : "—"}</td></tr>`;
    }).join("");
    els.preview.innerHTML = `<thead><tr><th>匹配值</th><th>填充映射</th><th>最终值</th><th>状态</th><th>来源</th></tr></thead><tbody>${rows}</tbody>`;
  }

  function exportWorkbook() {
    if (!state.result) return;
    const base = state.target.fileName.replace(/\.(xlsx|xlsm|xls|et)$/i, "");
    const ext = /\.xlsm$/i.test(state.target.fileName) ? "xlsm" : "xlsx";
    XLSX.writeFile(state.result.workbook, `${base}_已填写.${ext}`, { bookType: ext, cellStyles: true, bookVBA: ext === "xlsm" });
  }

  function removeSourceFile(id) {
    state.sourceFiles = state.sourceFiles.filter(file => file.id !== id); state.result = null;
    els.results.classList.add("hidden"); renderSourceFiles(); updateFieldSelectors(); clearMessage();
  }

  function changeSourceHeaderDepth(id, value) {
    const file = state.sourceFiles.find(item => item.id === id && item.status === "success");
    if (!file) return;
    try { refreshWorkbookStructure(file, Number(value)); state.result = null; els.results.classList.add("hidden"); renderSourceFiles(); updateFieldSelectors(); clearMessage(); }
    catch (error) { showMessage(`${file.fileName}：${error.message}`, "error"); }
  }

  function changeTargetHeaderDepth(value) {
    if (!state.target) return;
    try { refreshWorkbookStructure(state.target, Number(value)); state.result = null; els.results.classList.add("hidden"); renderTargetFile(); updateFieldSelectors(); clearMessage(); }
    catch (error) { showMessage(error.message, "error"); }
  }

  window.ExcelTemplateTool = Object.freeze({
    parseWorkbook, parseSourceFiles, parseMergedHeaders, buildColumnDefinitions, getSharedSourceFields,
    buildGlobalMatchIndex, resolveSourceField, isEmptyValue, normalizeComparableValue, resolveFieldValue,
    resolveMatchedValue, validateFieldMappings, fillTemplateFields, fillTemplate, exportWorkbook
  });

  els.sourceFile.addEventListener("change", async () => { await parseSourceFiles(els.sourceFile.files); els.sourceFile.value = ""; });
  els.targetFile.addEventListener("change", () => parseTargetFile(els.targetFile.files[0]));
  [["source", $("sourceDrop")], ["target", $("targetDrop")]].forEach(([kind, drop]) => {
    ["dragenter", "dragover"].forEach(event => drop.addEventListener(event, e => { e.preventDefault(); drop.classList.add("drag"); }));
    ["dragleave", "drop"].forEach(event => drop.addEventListener(event, e => { e.preventDefault(); drop.classList.remove("drag"); }));
    drop.addEventListener("drop", e => kind === "source" ? parseSourceFiles(e.dataTransfer.files) : parseTargetFile(e.dataTransfer.files[0]));
  });
  els.sourceList.addEventListener("click", event => { const button = event.target.closest("[data-remove-source]"); if (button) removeSourceFile(button.dataset.removeSource); });
  els.sourceList.addEventListener("change", event => { const select = event.target.closest("[data-source-depth]"); if (select) changeSourceHeaderDepth(select.dataset.sourceDepth, select.value); });
  els.mappingList.addEventListener("change", event => {
    const source = event.target.closest("[data-mapping-source]");
    const target = event.target.closest("[data-mapping-target]");
    const id = source?.dataset.mappingSource || target?.dataset.mappingTarget;
    const mapping = state.fieldMappings.find(item => item.id === id);
    if (mapping && source) mapping.sourceFieldKey = source.value;
    if (mapping && target) mapping.targetFieldKey = target.value;
    state.result = null; els.results.classList.add("hidden"); clearMessage();
  });
  els.mappingList.addEventListener("click", event => { const button = event.target.closest("[data-delete-mapping]"); if (button) deleteFieldMapping(button.dataset.deleteMapping); });
  els.addMapping.addEventListener("click", addFieldMapping);
  els.targetDepth.addEventListener("change", () => changeTargetHeaderDepth(els.targetDepth.value));
  [els.sourceMatch, els.targetMatch].forEach(select => select.addEventListener("change", updateActionState));
  els.fillButton.addEventListener("click", startFill);
  els.exportButton.addEventListener("click", exportWorkbook);
  renderFieldMappings();
})();
