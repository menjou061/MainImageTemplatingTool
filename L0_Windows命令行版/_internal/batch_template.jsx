#target photoshop

/*
 * Run from Photoshop: File > Scripts > Browse.
 * Keep the renamed PSD template open and active before running this script.
 */

var REPORT_NAME = "结果报告.csv";
// Keep the product inside the designer's smart-object frame.  The 94% inset
// leaves the PSD safety margin visible while still filling the usable area.
var PRODUCT_SAFE_SCALE = 0.94;
var CONTINUE_WITH_PREFLIGHT_ISSUES = !!($.global.__BATCH_INPUTS__ && $.global.__BATCH_INPUTS__.continueWithPreflightIssues);
var CHANNEL_PROFILE = ($.global.__BATCH_INPUTS__ && $.global.__BATCH_INPUTS__.profile) || null;
var ACTIVE_LAYOUT_GROUP = null;
var SWITCH_STATE = {};

function trimText(value) {
    return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function startsWith(value, prefix) {
    return String(value).indexOf(prefix) === 0;
}

function isBlank(value) {
    return trimText(value) === "";
}

function asNameList(value) {
    if (value == null || value === "") {
        return [];
    }
    return typeof value === "string" ? [value] : value;
}

function isDisabledImageValue(value) {
    var normalized = trimText(value).toLowerCase();
    return normalized === "无" || normalized === "无.png" || normalized === "none" || normalized === "null";
}

function isOptionalProfileKey(key) {
    if (!CHANNEL_PROFILE || !CHANNEL_PROFILE.optional_psd_variables) {
        return false;
    }
    for (var index = 0; index < CHANNEL_PROFILE.optional_psd_variables.length; index++) {
        if (String(CHANNEL_PROFILE.optional_psd_variables[index]) === String(key)) {
            return true;
        }
    }
    return false;
}

function dataFieldsOptional() {
    return !!(CHANNEL_PROFILE && CHANNEL_PROFILE.data_fields_optional);
}

function isProfileTextKey(key) {
    if (CHANNEL_PROFILE && CHANNEL_PROFILE.required_psd_variables) {
        for (var index = 0; index < CHANNEL_PROFILE.required_psd_variables.length; index++) {
            var required = CHANNEL_PROFILE.required_psd_variables[index];
            if (required.type === "text" && String(required.name) === String(key)) {
                return true;
            }
        }
        return false;
    }
    return key === "折扣" || key === "券名" || key === "券门槛" ||
        key === "活动时间" || key === "到手" || key === "价格1" ||
        key === "价格2" || key === "卖点" || key === "规格";
}

function normalizeBindingKey(value) {
    var text = trimText(value);
    while (text.charAt(0) === "@" || text.charAt(0) === "!" || text.charAt(0) === "#") {
        text = text.substring(1);
    }
    return text.replace(/[\s_\-:\/\\（）()【】\[\]{}<>]/g, "").toLowerCase();
}

function pushUnique(values, value) {
    if (isBlank(value)) { return; }
    for (var index = 0; index < values.length; index++) {
        if (String(values[index]) === String(value)) { return; }
    }
    values.push(String(value));
}

function profileFieldForKey(key) {
    if (!CHANNEL_PROFILE || !CHANNEL_PROFILE.fields) { return null; }
    var normalized = normalizeBindingKey(key);
    for (var index = 0; index < CHANNEL_PROFILE.fields.length; index++) {
        var field = CHANNEL_PROFILE.fields[index];
        var aliases = field.aliases || [];
        var candidates = [field.field_id, field.label, field.output_key];
        for (var aliasIndex = 0; aliasIndex < aliases.length; aliasIndex++) {
            candidates.push(aliases[aliasIndex]);
        }
        for (var candidateIndex = 0; candidateIndex < candidates.length; candidateIndex++) {
            if (!isBlank(candidates[candidateIndex]) && normalizeBindingKey(candidates[candidateIndex]) === normalized) {
                return field;
            }
        }
    }
    return null;
}

function isMainImageKey(key) {
    var field = profileFieldForKey(key);
    if (field && String(field.field_id) === "main_image") {
        return true;
    }
    var normalized = normalizeBindingKey(key);
    return normalized === "商品图" || normalized === "产品" ||
        normalized === "产品图" || normalized === "堆图";
}

function bindingNames(key) {
    var names = [];
    pushUnique(names, key);
    // Visibility groups use #字段名. Older tables commonly label the same
    // control column 字段名开关, so expose the canonical group name as a
    // generic compatibility alias rather than maintaining channel rules.
    var textKey = trimText(key);
    if (textKey.length > 2 && textKey.substring(textKey.length - 2) === "开关") {
        pushUnique(names, textKey.substring(0, textKey.length - 2));
    }
    var field = profileFieldForKey(key);
    if (field) {
        pushUnique(names, field.field_id);
        pushUnique(names, field.label);
        pushUnique(names, field.output_key);
        var aliases = field.aliases || [];
        for (var index = 0; index < aliases.length; index++) {
            pushUnique(names, aliases[index]);
        }
    }
    if (CHANNEL_PROFILE && CHANNEL_PROFILE.mapping) {
        for (var source in CHANNEL_PROFILE.mapping) {
            if (!CHANNEL_PROFILE.mapping.hasOwnProperty(source)) { continue; }
            if (normalizeBindingKey(CHANNEL_PROFILE.mapping[source]) === normalizeBindingKey(key)) {
                pushUnique(names, source);
            }
        }
    }
    return names;
}

function resolveRecordKey(key, record) {
    if (record && record.hasOwnProperty(key)) { return key; }
    var names = bindingNames(key);
    for (var index = 0; index < names.length; index++) {
        if (record && record.hasOwnProperty(names[index])) {
            return names[index];
        }
    }
    var normalizedKey = normalizeBindingKey(key);
    if (record) {
        for (var recordKey in record) {
            if (record.hasOwnProperty(recordKey) && normalizeBindingKey(recordKey) === normalizedKey) {
                return recordKey;
            }
        }
    }
    return key;
}

function recordValue(record, key) {
    var resolved = resolveRecordKey(key, record);
    return record && record.hasOwnProperty(resolved) ? record[resolved] : "";
}

function isRequiredTextKey(key) {
    if (isOptionalProfileKey(key) || dataFieldsOptional()) {
        return false;
    }
    return isProfileTextKey(key);
}

function getBasename(value) {
    var normalized = String(value == null ? "" : value).replace(/\\/g, "/");
    var pieces = normalized.split("/");
    return pieces[pieces.length - 1];
}

function decodedName(value) {
    try {
        return decodeURI(String(value));
    } catch (error) {
        return String(value);
    }
}

function removeExtension(value) {
    return String(value).replace(/\.[^.]+$/, "");
}

function safeOutputName(value) {
    var name = removeExtension(trimText(value));
    var illegal = "\\/:*?\"<>|";
    var result = "";
    for (var index = 0; index < name.length; index++) {
        var character = name.charAt(index);
        result += (illegal.indexOf(character) >= 0 || character === "\r" || character === "\n") ? "_" : character;
    }
    return result;
}

function csvEscape(value) {
    var text = String(value == null ? "" : value);
    return '"' + text.replace(/"/g, '""') + '"';
}

function parseCsv(text) {
    var rows = [];
    var row = [];
    var value = "";
    var quoted = false;
    var index;
    for (index = 0; index < text.length; index++) {
        var character = text.charAt(index);
        if (quoted) {
            if (character === '"') {
                if (text.charAt(index + 1) === '"') {
                    value += '"';
                    index++;
                } else {
                    quoted = false;
                }
            } else {
                value += character;
            }
        } else if (character === '"') {
            quoted = true;
        } else if (character === ",") {
            row.push(value);
            value = "";
        } else if (character === "\n") {
            row.push(value.replace(/\r$/, ""));
            rows.push(row);
            row = [];
            value = "";
        } else {
            value += character;
        }
    }
    if (value !== "" || row.length > 0) {
        row.push(value.replace(/\r$/, ""));
        rows.push(row);
    }
    if (!rows.length) {
        return [];
    }
    var headers = rows[0];
    if (headers[0] && headers[0].charCodeAt(0) === 65279) {
        headers[0] = headers[0].substring(1);
    }
    var hasProductName = false;
    for (var headerIndex = 0; headerIndex < headers.length; headerIndex++) {
        if (headers[headerIndex] === "商品文件名") {
            hasProductName = true;
            break;
        }
    }
    if (!hasProductName) {
        throw new Error("所选文件不是 data.csv：缺少【商品文件名】列表头。");
    }
    var data = [];
    for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
        if (rows[rowIndex].length === 1 && isBlank(rows[rowIndex][0])) {
            continue;
        }
        var record = {};
        for (var columnIndex = 0; columnIndex < headers.length; columnIndex++) {
            record[headers[columnIndex]] = rows[rowIndex][columnIndex] || "";
        }
        data.push(record);
    }
    return data;
}

function readCsv(file) {
    file.encoding = "UTF8";
    if (!file.open("r")) {
        throw new Error("无法读取 data.csv：" + file.fsName);
    }
    var text = file.read();
    file.close();
    return parseCsv(text);
}

var materialIndexSkipCount = 0;

function addFilesToIndex(folder, index) {
    var items;
    try {
        items = folder.getFiles();
    } catch (error) {
        materialIndexSkipCount++;
        return;
    }
    for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
        var item = items[itemIndex];
        try {
            if (item instanceof Folder) {
                addFilesToIndex(item, index);
            } else if (item instanceof File) {
                var rawKey = String(item.name).toLowerCase();
                var decodedKey = decodedName(item.name).toLowerCase();
                if (!index[rawKey]) {
                    index[rawKey] = item;
                }
                if (!index[decodedKey]) {
                    index[decodedKey] = item;
                }
            }
        } catch (itemError) {
            materialIndexSkipCount++;
        }
    }
}

function findMaterial(value, materialIndex) {
    if (isBlank(value)) {
        return null;
    }
    var rawValue = String(value);
    // data.csv preserves the absolute UNC path stored in Excel.  Prefer it
    // directly so the task never walks an entire shared drive to find a file.
    try {
        var rawFile = File(String(value));
        if (rawFile.exists) {
            return rawFile;
        }
        var normalizedFile = File(String(value).replace(/\\/g, "/"));
        if (normalizedFile.exists) {
            return normalizedFile;
        }
    } catch (directError) {}
    // An absolute path that cannot be opened must not fall back to a
    // same-named file from a material folder. That fallback can silently
    // replace a requested stack image with another product asset.
    if (isAbsoluteMaterialPath(rawValue)) {
        return null;
    }
    var key = decodedName(getBasename(value)).toLowerCase();
    return materialIndex ? (materialIndex[key] || null) : null;
}

var stagedMaterialFiles = {};

function pathHash(value) {
    var hash = 0;
    var text = String(value == null ? "" : value);
    for (var index = 0; index < text.length; index++) {
        hash = ((hash * 31) + text.charCodeAt(index)) & 0x7fffffff;
    }
    return hash.toString(16);
}

function needsLocalMaterialStaging(file) {
    if (!file) {
        return false;
    }
    var path = String(file.fsName || file.fullName || file);
    return path.indexOf("\\\\") === 0 || path.indexOf("//") === 0;
}

function stageMaterialForPhotoshop(file) {
    if (!needsLocalMaterialStaging(file)) {
        return file;
    }
    var sourcePath = String(file.fsName || file.fullName || file);
    if (stagedMaterialFiles[sourcePath]) {
        return stagedMaterialFiles[sourcePath];
    }
    var cacheFolder = new Folder(Folder.temp.fsName + "/l0_material_cache");
    if (!cacheFolder.exists && !cacheFolder.create()) {
        throw new Error("无法创建素材本地缓存目录：" + cacheFolder.fsName);
    }
    var baseName = getBasename(sourcePath);
    var target = new File(cacheFolder.fsName + "/" + pathHash(sourcePath) + "_" + baseName);
    if (!target.exists && !file.copy(target)) {
        throw new Error("网络素材无法复制到 Photoshop 本地缓存：" + sourcePath);
    }
    stagedMaterialFiles[sourcePath] = target;
    return target;
}

function isAbsoluteMaterialPath(value) {
    var text = trimText(value);
    if (text.length === 0) {
        return false;
    }
    var first = text.charAt(0);
    var second = text.charAt(1);
    var third = text.charAt(2);
    if (first === "/" || (first === "\\" && second === "\\")) {
        return true;
    }
    var upper = first.toUpperCase();
    return upper >= "A" && upper <= "Z" && second === ":" && (third === "\\" || third === "/");
}

function addIssue(result, issue) {
    result.issues.push(issue);
}

function addDataPrecheckWarning(result, issue) {
    result.preflightIssue = true;
    addCode(result, "W_DATA_PRECHECK");
    addIssue(result, issue + "；已按表格内容继续生成");
}

function addCode(result, code) {
    for (var index = 0; index < result.codes.length; index++) {
        if (result.codes[index] === code) {
            return;
        }
    }
    result.codes.push(code);
}

function containsIssue(result, issue) {
    for (var index = 0; index < result.issues.length; index++) {
        if (result.issues[index] === issue) {
            return true;
        }
    }
    return false;
}

function makeResult(productName, record) {
    return {
        productName: productName,
        sku: trimText(record && record["货号"]) || productName,
        issues: [],
        codes: [],
        missingImage: false,
        emptyField: false,
        textOverflow: false,
        optionalImageMissing: false,
        optionalImageReplaced: {},
        preflightIssue: false,
        validationFailed: false,
        templateInvalid: false,
        processingFailed: false,
        outputFile: "",
        psdOutputFile: "",
        variableBindingStatus: CHANNEL_PROFILE && CHANNEL_PROFILE.execution_mode === "photoshop_variables" ? "photoshop_variables" : "legacy_prefix",
        profileId: record && record.profile_id || (CHANNEL_PROFILE && CHANNEL_PROFILE.profile_id) || "legacy-v1",
        profileVersion: record && record.profile_version || (CHANNEL_PROFILE && CHANNEL_PROFILE.profile_version) || "1.0.0"
    };
}

function usesPhotoshopVariables(profile) {
    return !!(profile && profile.execution_mode === "photoshop_variables");
}

function usesStaticProductArt(profile) {
    return !!(profile && profile.static_product_art);
}

function usesStaticSupportArt(profile) {
    return !!(profile && (profile.static_support_art || profile.static_product_art));
}

function expectedVariableKind(type) {
    return type === "text" ? VariableKind.TEXT : VariableKind.PIXELREPLACEMENT;
}

function findDocumentVariable(document, name) {
    if (!document.variables) { return null; }
    for (var index = 0; index < document.variables.length; index++) {
        if (document.variables[index].name === name) {
            return document.variables[index];
        }
    }
    return null;
}

function isOptionalProfileVariable(profile, required) {
    if (!profile || !profile.optional_psd_variables) {
        return false;
    }
    var requiredName = required && required.name;
    for (var index = 0; index < profile.optional_psd_variables.length; index++) {
        if (String(profile.optional_psd_variables[index]) === String(requiredName)) {
            return true;
        }
    }
    return false;
}

function priceValidationError(record) {
    if (dataFieldsOptional() || (CHANNEL_PROFILE && CHANNEL_PROFILE.layout === "record_rows")) {
        return "";
    }
    var price1 = trimText(record["价格1"]);
    var price2 = trimText(record["价格2"]);
    if (isBlank(price1) || isBlank(price2)) {
        return "价格1、价格2必须完整填写";
    }
    return "";
}

function collectRecordLayoutMatches(container, configured, matches) {
    for (var layerIndex = 0; layerIndex < container.layers.length; layerIndex++) {
        var layer = container.layers[layerIndex];
        for (var groupIndex = 0; groupIndex < configured.length; groupIndex++) {
            if (layer.name === configured[groupIndex]) {
                matches[groupIndex].push(layer);
            }
        }
        if (layer.typename === "LayerSet") {
            collectRecordLayoutMatches(layer, configured, matches);
        }
    }
}

function collectProfileRecordLayoutMatches(container, configured, matches, profile) {
    if (!usesStaticProductArt(profile)) {
        collectRecordLayoutMatches(container, configured, matches);
        return;
    }
    // A fixed product composition may contain nested groups with the same SKU
    // name. Only direct document-level layout groups are selectable.
    for (var layerIndex = 0; layerIndex < container.layers.length; layerIndex++) {
        var layer = container.layers[layerIndex];
        for (var groupIndex = 0; groupIndex < configured.length; groupIndex++) {
            if (layer.name === configured[groupIndex]) {
                matches[groupIndex].push(layer);
            }
        }
    }
}

function selectRecordLayout(document, record) {
    if (!CHANNEL_PROFILE || CHANNEL_PROFILE.layout !== "record_rows") {
        ACTIVE_LAYOUT_GROUP = null;
        return;
    }
    var configured = asNameList(CHANNEL_PROFILE.record_layout && CHANNEL_PROFILE.record_layout.groups);
    var selectedName = trimText(record["版式组"]);
    ACTIVE_LAYOUT_GROUP = null;
    if (!configured || configured.length === 0 || !selectedName) {
        throw new Error("E_CONFIG_MISMATCH: 卫品记录缺少版式组");
    }

    var selectedIndex = -1;
    var configuredMatches = [];
    for (var configIndex = 0; configIndex < configured.length; configIndex++) {
        configuredMatches.push([]);
        if (trimText(configured[configIndex]) === selectedName) {
            if (selectedIndex >= 0) {
                throw new Error("E_CONFIG_MISMATCH: 渠道配置中版式组重复 " + selectedName);
            }
            selectedIndex = configIndex;
        }
    }
    if (selectedIndex < 0) {
        throw new Error("E_CONFIG_MISMATCH: 版式组未在渠道配置中声明 " + selectedName);
    }

    collectProfileRecordLayoutMatches(document, configured, configuredMatches, CHANNEL_PROFILE);
    var selectedMatches = configuredMatches[selectedIndex];
    if (selectedMatches.length === 0) {
        throw new Error("E_CONFIG_MISMATCH: 模板没有版式组 " + configured[selectedIndex]);
    }
    if (selectedMatches.length > 1) {
        throw new Error("E_CONFIG_MISMATCH: 模板版式组重复 " + configured[selectedIndex] + "（找到 " + selectedMatches.length + " 个）");
    }
    if (selectedMatches[0].typename !== "LayerSet") {
        throw new Error("E_CONFIG_MISMATCH: 模板版式组不是图层组 " + configured[selectedIndex]);
    }

    for (var matchIndex = 0; matchIndex < configuredMatches.length; matchIndex++) {
        for (var layerMatchIndex = 0; layerMatchIndex < configuredMatches[matchIndex].length; layerMatchIndex++) {
            if (configuredMatches[matchIndex][layerMatchIndex].typename === "LayerSet") {
                configuredMatches[matchIndex][layerMatchIndex].visible = false;
            }
        }
    }
    selectedMatches[0].visible = true;
    ACTIVE_LAYOUT_GROUP = selectedMatches[0];
}

function applyPreflightIssue(record, result) {
    var errors = trimText(record["预检异常"]);
    var warnings = trimText(record["预检提醒"]);
    if (!isBlank(errors)) {
        result.preflightIssue = true;
        addCode(result, "W_DATA_PRECHECK");
        addIssue(result, "清洗预检异常：" + errors);
    }
    if (!isBlank(warnings)) {
        result.preflightIssue = true;
        addCode(result, "W_DATA_PRECHECK");
        addIssue(result, "清洗预检提醒：" + warnings);
    }
}

function addLayers(container, layerIndex) {
    for (var index = 0; index < container.layers.length; index++) {
        var layer = container.layers[index];
        addLayerToIndex(layer, layerIndex);
        if (layer.typename === "LayerSet") {
            addLayers(layer, layerIndex);
        }
    }
}

function addLayerToIndex(layer, layerIndex) {
    var name = layer.name;
    if (startsWith(name, "@")) {
        var textKey = name.substring(1);
        if (!layerIndex.text[textKey]) {
            layerIndex.text[textKey] = [];
        }
        layerIndex.text[textKey].push(layer);
    } else if (startsWith(name, "!")) {
        var imageKey = name.substring(1);
        if (!layerIndex.image[imageKey]) {
            layerIndex.image[imageKey] = [];
        }
        layerIndex.image[imageKey].push(layer);
    } else if (startsWith(name, "#")) {
        var switchKey = name.substring(1);
        if (!layerIndex.switches[switchKey]) {
            layerIndex.switches[switchKey] = [];
        }
        layerIndex.switches[switchKey].push(layer);
    }
}

function isRecordLayoutGroup(layer) {
    if (!CHANNEL_PROFILE || CHANNEL_PROFILE.layout !== "record_rows" || layer.typename !== "LayerSet") {
        return false;
    }
    var configured = asNameList(CHANNEL_PROFILE.record_layout && CHANNEL_PROFILE.record_layout.groups);
    for (var index = 0; index < configured.length; index++) {
        if (layer.name === configured[index]) {
            return true;
        }
    }
    return false;
}

function addGlobalLayersOutsideRecordLayouts(container, layerIndex) {
    for (var index = 0; index < container.layers.length; index++) {
        var layer = container.layers[index];
        // Each selectable layout is mutually exclusive. Do not traverse it from
        // the document root, otherwise hidden alternative layouts would receive
        // the current row's values. Public layers outside those groups remain
        // dynamic and are applied for every selected layout.
        if (isRecordLayoutGroup(layer)) {
            continue;
        }
        addLayerToIndex(layer, layerIndex);
        if (layer.typename === "LayerSet") {
            addGlobalLayersOutsideRecordLayouts(layer, layerIndex);
        }
    }
}

function buildRecordLayerIndex(document) {
    var layerIndex = { text: {}, image: {}, switches: {} };
    if (ACTIVE_LAYOUT_GROUP) {
        addLayers(ACTIVE_LAYOUT_GROUP, layerIndex);
        addGlobalLayersOutsideRecordLayouts(document, layerIndex);
    } else {
        addLayers(document, layerIndex);
    }
    return layerIndex;
}

function keyCount(collection) {
    var count = 0;
    for (var key in collection) {
        if (collection.hasOwnProperty(key)) {
            count++;
        }
    }
    return count;
}

function requiredLayerMatches(layerIndex, required) {
    var collection = required.type === "text" ? layerIndex.text : layerIndex.image;
    var names = bindingNames(required.name);
    var normalizedNames = [];
    for (var normalizedIndex = 0; normalizedIndex < names.length; normalizedIndex++) {
        var normalizedName = normalizeBindingKey(names[normalizedIndex]);
        var alreadyIncluded = false;
        for (var includedIndex = 0; includedIndex < normalizedNames.length; includedIndex++) {
            if (normalizedNames[includedIndex] === normalizedName) {
                alreadyIncluded = true;
                break;
            }
        }
        if (!alreadyIncluded) {
            normalizedNames.push(normalizedName);
        }
    }
    var matches = [];
    for (var collectionName in collection) {
        if (!collection.hasOwnProperty(collectionName)) { continue; }
        var candidateName = normalizeBindingKey(collectionName);
        var nameMatched = false;
        for (var nameIndex = 0; nameIndex < normalizedNames.length; nameIndex++) {
            if (candidateName === normalizedNames[nameIndex]) {
                nameMatched = true;
                break;
            }
        }
        if (!nameMatched) { continue; }
        var layers = collection[collectionName] || [];
        for (var layerIndexValue = 0; layerIndexValue < layers.length; layerIndexValue++) {
            var exists = false;
            for (var matchIndex = 0; matchIndex < matches.length; matchIndex++) {
                if (matches[matchIndex] === layers[layerIndexValue]) { exists = true; break; }
            }
            if (!exists) { matches.push(layers[layerIndexValue]); }
        }
    }
    return matches;
}

function profileBindingErrors(template, profile) {
    if (!profile || !profile.required_psd_variables) { return []; }
    if (usesPhotoshopVariables(profile)) {
        var variableErrors = [];
        for (var variableIndex = 0; variableIndex < profile.required_psd_variables.length; variableIndex++) {
            var variableRequired = profile.required_psd_variables[variableIndex];
            var documentVariable = findDocumentVariable(template, variableRequired.name);
            if (!documentVariable) {
                variableErrors.push("E_VAR_MISSING: " + variableRequired.name);
            } else if (documentVariable.kind !== expectedVariableKind(variableRequired.type)) {
                variableErrors.push("E_VAR_TYPE_MISMATCH: " + variableRequired.name);
            }
        }
        return variableErrors;
    }
    if (profile.layout === "record_rows") {
        return recordLayoutBindingErrors(template, profile);
    }
    return requiredBindingErrors(template, profile, "");
}

function requiredBindingErrors(container, profile, scope) {
    var layerIndex = { text: {}, image: {}, switches: {} };
    addLayers(container, layerIndex);
    var errors = [];
    for (var index = 0; index < profile.required_psd_variables.length; index++) {
        var required = profile.required_psd_variables[index];
        var matches = requiredLayerMatches(layerIndex, required);
        if (!matches.length) {
            if (isOptionalProfileVariable(profile, required)) {
                continue;
            }
            errors.push("E_VAR_UNBOUND: " + scope + required.name + " 未绑定到 " + (required.type === "text" ? "@文本层" : "!智能对象"));
            continue;
        }
        for (var layerIndexValue = 0; layerIndexValue < matches.length; layerIndexValue++) {
            var layer = matches[layerIndexValue];
            if ((required.type === "text" && layer.kind !== LayerKind.TEXT) ||
                (required.type === "smart_object" && layer.kind !== LayerKind.SMARTOBJECT)) {
                errors.push("E_VAR_TYPE_MISMATCH: " + scope + required.name);
                break;
            }
        }
    }
    return errors;
}

function recordLayoutBindingErrors(template, profile) {
    var errors = [];
    var configured = asNameList(profile.record_layout && profile.record_layout.groups);
    if (!configured || configured.length === 0) {
        return ["E_CONFIG_MISMATCH: 卫品渠道缺少版式组配置"];
    }
    var active = asNameList(profile.active_layout_groups);
    var requested = active.length > 0 ? active : configured;
    var matches = [];
    for (var index = 0; index < requested.length; index++) {
        matches.push([]);
    }
    collectProfileRecordLayoutMatches(template, requested, matches, profile);
    for (var layoutIndex = 0; layoutIndex < requested.length; layoutIndex++) {
        var layoutName = requested[layoutIndex];
        var layoutMatches = matches[layoutIndex];
        if (layoutMatches.length !== 1 || layoutMatches[0].typename !== "LayerSet") {
            errors.push("E_CONFIG_MISMATCH: 版式组 " + layoutName + " 缺失、重复或不是图层组");
            continue;
        }
        var scope = layoutName + "/";
        var layoutErrors = requiredBindingErrors(layoutMatches[0], profile, scope);
        for (var errorIndex = 0; errorIndex < layoutErrors.length; errorIndex++) {
            errors.push(layoutErrors[errorIndex]);
        }
        var layoutIndexData = { text: {}, image: {}, switches: {} };
        addLayers(layoutMatches[0], layoutIndexData);
        var giftsAreOptional = isOptionalProfileKey("赠品图1") && isOptionalProfileKey("赠品文案1");
        if (!usesStaticSupportArt(profile) && !giftsAreOptional && !layoutIndexData.switches["赠品区域"]) {
            errors.push("E_VAR_UNBOUND: " + scope + "赠品区域开关组未绑定");
        }
    }
    return errors;
}

function validateTemplate(template) {
    if (usesPhotoshopVariables(CHANNEL_PROFILE)) {
        var variableBindingErrors = profileBindingErrors(template, CHANNEL_PROFILE);
        if (variableBindingErrors.length) {
            throw new Error(variableBindingErrors.join("；"));
        }
        return;
    }
    var layerIndex = { text: {}, image: {}, switches: {} };
    addLayers(template, layerIndex);
    if (keyCount(layerIndex.text) === 0 || (!usesStaticProductArt(CHANNEL_PROFILE) && keyCount(layerIndex.image) === 0)) {
        throw new Error(usesStaticProductArt(CHANNEL_PROFILE) ?
            "当前固定商品组合 PSD 尚未按规范完成改造：至少需要一个 @文本层。" :
            "当前 PSD 尚未按规范完成改造：至少需要一个 @文本层和一个 !智能对象层。");
    }
    var bindingErrors = profileBindingErrors(template, CHANNEL_PROFILE);
    if (bindingErrors.length) {
        throw new Error(bindingErrors.join("；"));
    }
}

function isYes(value) {
    return trimText(value) === "是";
}

function isNo(value) {
    return trimText(value) === "否";
}

function setSwitches(layerIndex, record, result) {
    for (var key in layerIndex.switches) {
        if (!layerIndex.switches.hasOwnProperty(key)) {
            continue;
        }
        var value;
        if (record && record.hasOwnProperty(key) && (isYes(record[key]) || isNo(record[key]))) {
            // An exact #字段名 column is an explicit designer/operator override.
            value = record[key];
        } else {
            // A switch group is only a visibility container. It follows
            // dynamic children, so new fields need no channel-specific rule.
            value = switchGroupsHaveBoundValue(layerIndex.switches[key], record) ? "是" : "否";
        }
        SWITCH_STATE[key] = value;
        var visible = false;
        if (isYes(value)) {
            visible = true;
        } else if (isNo(value)) {
            visible = false;
        } else {
            // Optional groups never block a task. An unrecognised value means hidden.
            visible = false;
        }
        var layers = layerIndex.switches[key];
        for (var index = 0; index < layers.length; index++) {
            layers[index].visible = visible;
        }
    }
}

function hasBindingForKey(layerIndex, key) {
    var names = bindingNames(key);
    for (var nameIndex = 0; nameIndex < names.length; nameIndex++) {
        var normalized = normalizeBindingKey(names[nameIndex]);
        for (var textKey in layerIndex.text) {
            if (layerIndex.text.hasOwnProperty(textKey) && normalizeBindingKey(textKey) === normalized) { return true; }
        }
        for (var imageKey in layerIndex.image) {
            if (layerIndex.image.hasOwnProperty(imageKey) && normalizeBindingKey(imageKey) === normalized) { return true; }
        }
        for (var switchKey in layerIndex.switches) {
            if (layerIndex.switches.hasOwnProperty(switchKey) && normalizeBindingKey(switchKey) === normalized) { return true; }
        }
    }
    return false;
}

function dynamicBindingWarnings(layerIndex, record) {
    if (!CHANNEL_PROFILE || !CHANNEL_PROFILE.matching || !CHANNEL_PROFILE.matching.allow_unknown_fields) {
        return [];
    }
    var warnings = [];
    for (var key in record) {
        if (!record.hasOwnProperty(key) || isBlank(record[key])) { continue; }
        if (key === "商品文件名" || key === "货号" || key === "profile_id" || key === "profile_version" ||
            key === "variant" || key === "预检异常" || key === "预检提醒" || key === "版式组") { continue; }
        var field = profileFieldForKey(key);
        if (field && String(field.field_id) === "product_name") { continue; }
        if (!hasBindingForKey(layerIndex, key)) {
            warnings.push("W_BINDING_TABLE_ONLY: CSV 字段【" + key + "】有值，但 PSD 没有同名 @文本层或 !智能对象");
        }
    }
    for (var textKey in layerIndex.text) {
        if (layerIndex.text.hasOwnProperty(textKey) && layerIndex.image.hasOwnProperty(textKey)) {
            warnings.push("W_BINDING_TYPE_MISMATCH: " + textKey + " 同时存在 @文本层和 !智能对象");
        }
    }
    return warnings;
}

function isGiftSwitchKey(key) {
    return key === "赠品顶部" || key === "赠品区域";
}

function isPriceSwitchKey(key) {
    return key === "价格优惠券" || key === "价格立减";
}

function switchGroupsHaveBoundValue(groups, record) {
    for (var index = 0; index < groups.length; index++) {
        if (groupHasBoundValue(groups[index], record)) {
            return true;
        }
    }
    return false;
}

function groupHasBoundValue(group, record) {
    if (!group || !group.layers) {
        return false;
    }
    for (var index = 0; index < group.layers.length; index++) {
        var child = group.layers[index];
        if (startsWith(child.name, "@")) {
            if (!isBlank(recordValue(record, child.name.substring(1)))) {
                return true;
            }
        } else if (startsWith(child.name, "!")) {
            var imageValue = recordValue(record, child.name.substring(1));
            if (!isBlank(imageValue) && !isDisabledImageValue(imageValue)) {
                return true;
            }
        } else if (child.typename === "LayerSet" && groupHasBoundValue(child, record)) {
            return true;
        }
    }
    return false;
}

function groupHasAllImageValues(group, record) {
    var foundImage = false;
    for (var index = 0; index < group.layers.length; index++) {
        var child = group.layers[index];
        if (startsWith(child.name, "!")) {
            foundImage = true;
            var imageValue = recordValue(record, child.name.substring(1));
            if (isBlank(imageValue) || isDisabledImageValue(imageValue)) {
                return false;
            }
        } else if (child.typename === "LayerSet" && !startsWith(child.name, "#") && groupContainsImageLayer(child)) {
            if (!groupHasAllImageValues(child, record)) {
                return false;
            }
            foundImage = true;
        }
    }
    return foundImage;
}

function groupHasAnyImageValue(group, record) {
    for (var index = 0; index < group.layers.length; index++) {
        var child = group.layers[index];
        if (startsWith(child.name, "!")) {
            var imageValue = recordValue(record, child.name.substring(1));
            if (!isBlank(imageValue) && !isDisabledImageValue(imageValue)) {
                return true;
            }
        } else if (child.typename === "LayerSet" && !startsWith(child.name, "#") && groupHasAnyImageValue(child, record)) {
            return true;
        }
    }
    return false;
}

function groupContainsImageLayer(group) {
    for (var index = 0; index < group.layers.length; index++) {
        var child = group.layers[index];
        if (startsWith(child.name, "!")) {
            return true;
        }
        if (child.typename === "LayerSet" && !startsWith(child.name, "#") && groupContainsImageLayer(child)) {
            return true;
        }
    }
    return false;
}

function groupHasAllReplacedImageLayers(group, result) {
    var foundImage = false;
    for (var index = 0; index < group.layers.length; index++) {
        var child = group.layers[index];
        if (startsWith(child.name, "!")) {
            foundImage = true;
            if (!result.optionalImageReplaced[child.name.substring(1)]) {
                return false;
            }
        } else if (child.typename === "LayerSet" && !startsWith(child.name, "#") && groupContainsImageLayer(child)) {
            if (!groupHasAllReplacedImageLayers(child, result)) {
                return false;
            }
            foundImage = true;
        }
    }
    return foundImage;
}

function groupHasAnyReplacedImageLayer(group, result) {
    for (var index = 0; index < group.layers.length; index++) {
        var child = group.layers[index];
        if (startsWith(child.name, "!") && result.optionalImageReplaced[child.name.substring(1)]) {
            return true;
        }
        if (child.typename === "LayerSet" && !startsWith(child.name, "#") && groupHasAnyReplacedImageLayer(child, result)) {
            return true;
        }
    }
    return false;
}

function reconcileOptionalGroups(layerIndex, record, result) {
    for (var key in layerIndex.switches) {
        if (!layerIndex.switches.hasOwnProperty(key)) {
            continue;
        }
        var groups = layerIndex.switches[key];
        var requested = isYes(switchStateValue(record, key));
        for (var groupIndex = 0; groupIndex < groups.length; groupIndex++) {
            groups[groupIndex].visible = requested;
        }
        if (requested) {
            addIssue(result, "可选图层已展示：#" + key);
        } else {
            addIssue(result, "可选图层已隐藏：#" + key);
        }
    }
}

function collectNamedGroups(container, name, output) {
    if (!container || !container.layers) {
        return;
    }
    for (var index = 0; index < container.layers.length; index++) {
        var layer = container.layers[index];
        if (layer.typename !== "LayerSet") {
            continue;
        }
        if (layer.name === name) {
            output.push(layer);
        }
        collectNamedGroups(layer, name, output);
    }
}

function reconcileGiftSlotVisibility(container, record, result) {
    // The member strip is one switch group, but each of its three columns is
    // an independent slot. Hide only the empty columns so one/two gift rows
    // do not leave stale master artwork in the remaining columns.
    for (var slotIndex = 1; slotIndex <= 3; slotIndex++) {
        var groups = [];
        collectNamedGroups(container, "赠品槽" + slotIndex, groups);
        var imageKey = "赠品图" + slotIndex;
        var copyKey = "赠品文案" + slotIndex;
        var imageValue = recordValue(record, imageKey);
        var copyValue = recordValue(record, copyKey);
        var replaced = result && result.optionalImageReplaced &&
            result.optionalImageReplaced[imageKey] === true;
        var visible = !isBlank(imageValue) && !isDisabledImageValue(imageValue) &&
            !isBlank(copyValue) && replaced;
        for (var groupIndex = 0; groupIndex < groups.length; groupIndex++) {
            groups[groupIndex].visible = visible;
        }
    }
}

function insideDisabledSwitch(layer, record) {
    var parent = layer.parent;
    while (parent && parent.typename !== "Document") {
        if (startsWith(parent.name, "#") && !isYes(switchStateValue(record, parent.name.substring(1)))) {
            return true;
        }
        parent = parent.parent;
    }
    return false;
}

function switchStateValue(record, key) {
    if (SWITCH_STATE.hasOwnProperty(key)) {
        return SWITCH_STATE[key];
    }
    if (record && record.hasOwnProperty(key) && (isYes(record[key]) || isNo(record[key]))) {
        return record[key];
    }
    return "否";
}

function layerWidth(layer) {
    var bounds = layer.bounds;
    return bounds[2].as("px") - bounds[0].as("px");
}

function paragraphTextWidth(layer) {
    try {
        if (layer.textItem.kind === TextType.PARAGRAPHTEXT) {
            return layer.textItem.width.as("px");
        }
    } catch (error) {}
    return null;
}

function rectWidth(rect) {
    return rect.right - rect.left;
}

function rectHeight(rect) {
    return rect.bottom - rect.top;
}

function textFitConfig(key) {
    if (!CHANNEL_PROFILE || !CHANNEL_PROFILE.text_fit || !CHANNEL_PROFILE.text_fit.frame_by_key) {
        return null;
    }
    var configured = CHANNEL_PROFILE.text_fit.frame_by_key[key];
    return configured || null;
}

function textFrameForKey(key, originalRect) {
    var configured = textFitConfig(key);
    if (!configured || !app.activeDocument) {
        return null;
    }
    var canvasWidth = app.activeDocument.width.as("px");
    var canvasHeight = app.activeDocument.height.as("px");
    var left = typeof configured.left_px === "number" ? configured.left_px : canvasWidth * Number(configured.left_ratio || 0);
    var right = typeof configured.right_px === "number" ? configured.right_px : canvasWidth * Number(configured.right_ratio || 1);
    var top = typeof configured.top_px === "number" ? configured.top_px : originalRect.top;
    var bottom = typeof configured.bottom_px === "number" ? configured.bottom_px : Math.min(originalRect.bottom, canvasHeight);
    if (!(right > left) || !(bottom > top)) {
        return null;
    }
    return { left: left, top: top, right: right, bottom: bottom };
}

function siblingLeftBoundary(layer, siblingNames, originalRect) {
    var boundary = null;
    var siblings = layer.parent.layers;
    for (var index = 0; index < siblings.length; index++) {
        var sibling = siblings[index];
        for (var nameIndex = 0; nameIndex < siblingNames.length; nameIndex++) {
            if (sibling.name === siblingNames[nameIndex]) {
                var siblingRect = layerRect(sibling);
                if (siblingRect.left > originalRect.left + 1 &&
                    siblingRect.bottom > originalRect.top &&
                    siblingRect.top < originalRect.bottom) {
                    boundary = boundary === null ? siblingRect.left : Math.min(boundary, siblingRect.left);
                }
            }
        }
    }
    return boundary;
}

function textMaxWidth(layer, key, originalRect) {
    var maxWidth = paragraphTextWidth(layer);
    if (maxWidth === null) {
        maxWidth = rectWidth(originalRect) * 0.98;
    }
    var configuredFrame = textFrameForKey(key, originalRect);
    if (configuredFrame) {
        maxWidth = Math.min(maxWidth, rectWidth(configuredFrame));
    }
    var adjacentBoundary = null;
    if (key === "折扣") {
        adjacentBoundary = siblingLeftBoundary(layer, ["折"], originalRect);
    } else if (key === "价格1") {
        adjacentBoundary = siblingLeftBoundary(layer, ["@价格2"], originalRect);
    } else if (key === "利益点1") {
        // The two promotional benefits share one visual line. Keep the first
        // value inside the space before the second value instead of allowing
        // its point text to overlap when a longer offer is supplied.
        adjacentBoundary = siblingLeftBoundary(layer, ["@利益点2"], originalRect);
    }
    if (adjacentBoundary !== null) {
        var safetyGap = key === "利益点1" ? 20 : 4;
        maxWidth = Math.min(maxWidth, Math.max(1, adjacentBoundary - originalRect.left - safetyGap));
    }
    return Math.max(1, maxWidth);
}

function autoFitMinimumScale(key) {
    if (!CHANNEL_PROFILE || !CHANNEL_PROFILE.text_fit || !CHANNEL_PROFILE.text_fit.keys) {
        return null;
    }
    for (var index = 0; index < CHANNEL_PROFILE.text_fit.keys.length; index++) {
        if (CHANNEL_PROFILE.text_fit.keys[index] === key) {
            var configured = CHANNEL_PROFILE.text_fit.minimum_scale;
            if (CHANNEL_PROFILE.text_fit.minimum_scale_by_key &&
                typeof CHANNEL_PROFILE.text_fit.minimum_scale_by_key[key] !== "undefined") {
                configured = CHANNEL_PROFILE.text_fit.minimum_scale_by_key[key];
            }
            var scale = Number(configured);
            return scale > 0 && scale <= 1 ? scale : 0.85;
        }
    }
    return null;
}

function fitTextToOriginalFrame(layer, key, originalRect, minimumScale) {
    var maxWidth = textMaxWidth(layer, key, originalRect);
    var maxHeight = Math.max(1, rectHeight(originalRect) * 0.98);
    var configuredFrame = textFrameForKey(key, originalRect);
    if (configuredFrame) {
        maxHeight = Math.min(maxHeight, rectHeight(configuredFrame));
    }
    var appliedScale = 1;
    var fitted = false;
    for (var attempt = 0; attempt < 12; attempt++) {
        var currentRect = layerRect(layer);
        var currentWidth = rectWidth(currentRect);
        var currentHeight = rectHeight(currentRect);
        if (currentWidth <= 0 || currentHeight <= 0) {
            break;
        }
        if (currentWidth <= maxWidth + 0.5 && currentHeight <= maxHeight + 0.5) {
            fitted = true;
            break;
        }
        var factor = Math.min(maxWidth / currentWidth, maxHeight / currentHeight) * 0.98;
        if (factor >= 0.999) {
            break;
        }
        if (minimumScale !== null && appliedScale * factor < minimumScale) {
            factor = minimumScale / appliedScale;
        }
        if (factor >= 0.999) {
            break;
        }
        // Some supplied PSD text layers have a designer-applied transform.
        // Scaling the layer keeps that transform and its visual font size in
        // sync, unlike assigning textItem.size directly.
        layer.resize(factor * 100, factor * 100, AnchorPosition.TOPLEFT);
        appliedScale *= factor;
    }

    // Preserve the template's visual anchor after a point-text resize. The
    // selling-point headline and final-price number are centered; anchoring
    // either by left edge would make shorter values drift inside their badge.
    var finalRect = layerRect(layer);
    var shiftX;
    if (configuredFrame && key === "价格1") {
        // Price numbers share a fixed unit layer (@价格2) on their right.
        // Keep the numeric layer left-anchored so normal two-digit values do
        // not drift, while long values can shrink only inside the declared
        // price frame.
        shiftX = configuredFrame.left - finalRect.left;
    } else if (configuredFrame) {
        var frameCenterX = (configuredFrame.left + configuredFrame.right) / 2;
        var configuredFinalCenterX = (finalRect.left + finalRect.right) / 2;
        shiftX = frameCenterX - configuredFinalCenterX;
    } else if (key === "卖点" || key === "到手") {
        var originalCenterX = (originalRect.left + originalRect.right) / 2;
        var finalCenterX = (finalRect.left + finalRect.right) / 2;
        shiftX = originalCenterX - finalCenterX;
    } else {
        shiftX = originalRect.left - finalRect.left;
    }
    layer.translate(
        UnitValue(shiftX, "px"),
        UnitValue(originalRect.top - finalRect.top, "px")
    );
    finalRect = layerRect(layer);
    if (configuredFrame) {
        var frameShiftX = 0;
        if (finalRect.left < configuredFrame.left) {
            frameShiftX = configuredFrame.left - finalRect.left;
        } else if (finalRect.right > configuredFrame.right) {
            frameShiftX = configuredFrame.right - finalRect.right;
        }
        if (Math.abs(frameShiftX) > 0.5) {
            layer.translate(UnitValue(frameShiftX, "px"), UnitValue(0, "px"));
            finalRect = layerRect(layer);
        }
    }
    return {
        fitted: rectWidth(finalRect) <= maxWidth + 0.5 && rectHeight(finalRect) <= maxHeight + 0.5 &&
            (!configuredFrame || (finalRect.left >= configuredFrame.left - 0.5 && finalRect.right <= configuredFrame.right + 0.5)),
        maxWidth: maxWidth,
        finalRect: finalRect
    };
}

function photoshopTextValue(value) {
    return String(value == null ? "" : value).replace(/\r\n/g, "\r").replace(/\n/g, "\r");
}

function textStyleRangePlan(key, value, sourceCount) {
    var visibleText = photoshopTextValue(value);
    var totalLength = visibleText.length + 1;
    var ranges = [];
    var markerIndex;
    var decimalIndex;
    var integerStart;

    if (key === "卖点") {
        markerIndex = visibleText.indexOf("*");
        var fullWidthMarkerIndex = visibleText.indexOf("＊");
        if (markerIndex < 0 || (fullWidthMarkerIndex >= 0 && fullWidthMarkerIndex < markerIndex)) {
            markerIndex = fullWidthMarkerIndex;
        }
        if (markerIndex < 0) {
            return sourceCount > 0 ? [{ from: 0, to: totalLength, source: 0 }] : null;
        }
        if (sourceCount < 3 || markerIndex === 0) {
            return null;
        }
        ranges.push({ from: 0, to: markerIndex, source: 0 });
        ranges.push({ from: markerIndex, to: markerIndex + 1, source: 1 });
        ranges.push({ from: markerIndex + 1, to: totalLength, source: 2 });
        return ranges;
    }

    if (key === "到手") {
        integerStart = /^[¥￥$]/.test(visibleText) ? 1 : 0;
        decimalIndex = visibleText.indexOf(".", integerStart);
        if (integerStart > 0) {
            if (sourceCount < 2 || visibleText.length <= integerStart) {
                return null;
            }
            ranges.push({ from: 0, to: integerStart, source: 0 });
            if (decimalIndex >= 0) {
                if (sourceCount < 3 || decimalIndex === integerStart) {
                    return null;
                }
                ranges.push({ from: integerStart, to: decimalIndex, source: 1 });
                ranges.push({ from: decimalIndex, to: totalLength, source: 2 });
            } else {
                ranges.push({ from: integerStart, to: totalLength, source: 1 });
            }
            return ranges;
        }
        if (decimalIndex >= 0) {
            if (sourceCount < 3 || decimalIndex === 0) {
                return null;
            }
            ranges.push({ from: 0, to: decimalIndex, source: 1 });
            ranges.push({ from: decimalIndex, to: totalLength, source: 2 });
            return ranges;
        }
        return sourceCount > 1 ? [{ from: 0, to: totalLength, source: 1 }] : null;
    }

    if (key === "价格活动价" || key === "价格优惠券" || key === "价格立减") {
        decimalIndex = visibleText.indexOf(".");
        if (decimalIndex >= 0) {
            if (sourceCount < 2 || decimalIndex === 0) {
                return null;
            }
            ranges.push({ from: 0, to: decimalIndex, source: 0 });
            ranges.push({ from: decimalIndex, to: totalLength, source: 1 });
            return ranges;
        }
        return sourceCount > 0 ? [{ from: 0, to: totalLength, source: 0 }] : null;
    }

    return null;
}

function paragraphRangePlan(text, sourceCount) {
    var ranges = [];
    var start = 0;
    for (var index = 0; index < text.length; index++) {
        if (text.charAt(index) === "\r") {
            ranges.push({
                from: start,
                to: index + 1,
                source: Math.min(ranges.length, sourceCount - 1)
            });
            start = index + 1;
        }
    }
    ranges.push({
        from: start,
        to: text.length + 1,
        source: Math.min(ranges.length, sourceCount - 1)
    });
    return ranges;
}

function replaceDescriptorRangeList(descriptor, rangeKey, ranges) {
    var sourceList = descriptor.getList(rangeKey);
    var replacement = new ActionList();
    var fromKey = stringIDToTypeID("from");
    var toKey = stringIDToTypeID("to");
    for (var index = 0; index < ranges.length; index++) {
        if (ranges[index].source < 0 || ranges[index].source >= sourceList.count || ranges[index].to <= ranges[index].from) {
            return false;
        }
        var rangeDescriptor = sourceList.getObjectValue(ranges[index].source);
        rangeDescriptor.putInteger(fromKey, ranges[index].from);
        rangeDescriptor.putInteger(toKey, ranges[index].to);
        replacement.putObject(sourceList.getObjectType(ranges[index].source), rangeDescriptor);
    }
    descriptor.putList(rangeKey, replacement);
    return true;
}

function setTextContentsPreservingStyle(layer, value, key) {
    try {
        app.activeDocument.activeLayer = layer;
        var targetReference = new ActionReference();
        targetReference.putEnumerated(charIDToTypeID("Lyr "), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
        var layerDescriptor = executeActionGet(targetReference);
        var textKey = stringIDToTypeID("textKey");
        var styleRangeKey = stringIDToTypeID("textStyleRange");
        var paragraphRangeKey = stringIDToTypeID("paragraphStyleRange");
        if (!layerDescriptor.hasKey(textKey)) { return false; }
        var textDescriptor = layerDescriptor.getObjectValue(textKey);
        var oldValue = textDescriptor.getString(textKey);
        var targetValue = photoshopTextValue(value);
        if (!textDescriptor.hasKey(styleRangeKey)) { return false; }
        var sourceStyleCount = textDescriptor.getList(styleRangeKey).count;
        var styleRanges = textStyleRangePlan(key, value, sourceStyleCount);
        if (!styleRanges && oldValue.length !== targetValue.length) { return false; }
        if (styleRanges && !replaceDescriptorRangeList(textDescriptor, styleRangeKey, styleRanges)) { return false; }
        if (textDescriptor.hasKey(paragraphRangeKey)) {
            var paragraphCount = textDescriptor.getList(paragraphRangeKey).count;
            if (paragraphCount < 1 || !replaceDescriptorRangeList(textDescriptor, paragraphRangeKey, paragraphRangePlan(targetValue, paragraphCount))) {
                return false;
            }
        }
        textDescriptor.putString(textKey, targetValue);
        var setDescriptor = new ActionDescriptor();
        var setReference = new ActionReference();
        setReference.putEnumerated(charIDToTypeID("TxLr"), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
        setDescriptor.putReference(charIDToTypeID("null"), setReference);
        setDescriptor.putObject(charIDToTypeID("T   "), charIDToTypeID("TxLr"), textDescriptor);
        executeAction(charIDToTypeID("setd"), setDescriptor, DialogModes.NO);
        return true;
    } catch (error) {
        return false;
    }
}

function setTextLayer(layer, value, key, record, result) {
    if (insideDisabledSwitch(layer, record)) {
        return;
    }
    if (isBlank(value)) {
        layer.textItem.contents = "";
        if (key === "备注") {
            layer.visible = false;
        }
        if (dataFieldsOptional() && isProfileTextKey(key)) {
            addDataPrecheckWarning(result, "字段为空：" + key);
        } else if (isRequiredTextKey(key)) {
            result.emptyField = true;
            addCode(result, "E_EMPTY_FIELD");
            addIssue(result, "字段为空：" + key);
        }
        return;
    }
    if (key === "备注") {
        layer.visible = true;
    }
    var originalRect = layerRect(layer);
    if (!setTextContentsPreservingStyle(layer, value, key)) {
        layer.textItem.contents = String(value);
    }
    // These point-text layers sit inside fixed visual frames. Keep their
    // paragraph justification centered while preserving the PSD designer's
    // original font size and transform.
    if (key === "卖点" || key === "到手") {
        try { layer.textItem.justification = Justification.CENTER; } catch (ignoreJustification) {}
    }
    var textRect = layerRect(layer);
    var configuredFrame = textFrameForKey(key, originalRect);
    // Price fitting is opt-in per channel/variant. The profile must declare a
    // concrete frame, so no unrelated text layer can be resized implicitly.
    var priceFitMinimumScale = autoFitMinimumScale(key);
    if (key === "价格1" && configuredFrame && priceFitMinimumScale !== null) {
        var priceFit = fitTextToOriginalFrame(layer, key, originalRect, priceFitMinimumScale);
        textRect = priceFit.finalRect;
    }
    var canvas = app.activeDocument;
    var outOfCanvas = textRect.left < -0.5 || textRect.top < -0.5 ||
        textRect.right > canvas.width.as("px") + 0.5 ||
        textRect.bottom > canvas.height.as("px") + 0.5;
    var outOfConfiguredFrame = configuredFrame && (
        textRect.left < configuredFrame.left - 0.5 ||
        textRect.right > configuredFrame.right + 0.5 ||
        textRect.top < configuredFrame.top - 0.5 ||
        textRect.bottom > configuredFrame.bottom + 0.5
    );
    // Point-text bounds naturally change with the value. The original bounds
    // are only a measurement of the placeholder, not a designer-defined text
    // frame, so comparing against them produces false "overflow" warnings for
    // otherwise correct PSD-native text. Only flag a real canvas escape or an
    // explicit profile frame that the template declares as constrained.
    if (outOfCanvas || outOfConfiguredFrame) {
        result.textOverflow = true;
        addCode(result, "W_TEXT_OVERFLOW");
        addIssue(result, "文案超框：" + key);
    }
}

function replaceSmartObject(layer, imageFile) {
    app.activeDocument.activeLayer = layer;
    var descriptor = new ActionDescriptor();
    descriptor.putPath(charIDToTypeID("null"), imageFile);
    executeAction(stringIDToTypeID("placedLayerReplaceContents"), descriptor, DialogModes.NO);
}

function isolateSmartObjectForReplacement(layer) {
    // Photoshop smart-object instances can share one embedded object. Replacing
    // one shared instance then changes every linked instance, including
    // decorative layers that are not table-driven. "New Smart Object via Copy"
    // creates an independent embedded object while retaining the layer frame.
    if (!layer || layer.typename !== "ArtLayer" || layer.kind !== LayerKind.SMARTOBJECT) {
        return { layer: layer, copied: false };
    }
    var document = app.activeDocument;
    var original = layer;
    var originalName = String(layer.name);
    document.activeLayer = original;
    executeAction(stringIDToTypeID("placedLayerMakeCopy"), undefined, DialogModes.NO);
    var copy = document.activeLayer;
    if (!copy || copy === original || copy.typename !== "ArtLayer" || copy.kind !== LayerKind.SMARTOBJECT) {
        throw new Error("无法建立独立智能对象副本：" + originalName);
    }
    copy.name = originalName;
    try {
        original.remove();
    } catch (removeError) {
        try { copy.remove(); } catch (cleanupError) {}
        throw new Error("无法替换共享智能对象副本：" + originalName + "（" + removeError.message + "）");
    }
    return { layer: copy, copied: true };
}

function layerRect(layer) {
    var bounds = layer.bounds;
    return {
        left: bounds[0].as("px"),
        top: bounds[1].as("px"),
        right: bounds[2].as("px"),
        bottom: bounds[3].as("px")
    };
}

function visiblePixelRect(layer) {
    // A placed PNG smart object keeps its source canvas bounds, including any
    // transparent padding. Rasterise a disposable duplicate so Photoshop gives
    // us the actual non-transparent product bounds in document coordinates.
    // Never rasterise or otherwise alter the template's real smart object.
    var probe = null;
    try {
        probe = layer.duplicate();
        probe.rasterize(RasterizeType.ENTIRELAYER);
        var rect = layerRect(probe);
        probe.remove();
        probe = null;
        // Some nested smart objects rasterise to an empty probe even though
        // their original placeholder bounds are valid (notably gift slot 1
        // in the hygiene member template). Keep the compatibility path by
        // falling back to the original smart-object frame in that case.
        if (rect.right > rect.left && rect.bottom > rect.top) {
            return rect;
        }
        return layerRect(layer);
    } catch (error) {
        if (probe) {
            try {
                probe.remove();
            } catch (ignore) {
            }
        }
        // Keep the previous behaviour for unusual smart objects that Photoshop
        // cannot rasterise, rather than stopping an otherwise valid export.
        return layerRect(layer);
    }
}

function fitProductToTemplateFrame(layer, targetRect) {
    // Use visible product pixels so transparent padding in a source PNG does
    // not make the product appear smaller than the PSD designer intended.
    var currentRect = visiblePixelRect(layer);
    var currentWidth = currentRect.right - currentRect.left;
    var currentHeight = currentRect.bottom - currentRect.top;
    var frameWidth = targetRect.right - targetRect.left;
    var frameHeight = targetRect.bottom - targetRect.top;
    var targetWidth = frameWidth * PRODUCT_SAFE_SCALE;
    var targetHeight = frameHeight * PRODUCT_SAFE_SCALE;
    if (currentWidth <= 0 || currentHeight <= 0 || targetWidth <= 0 || targetHeight <= 0) {
        throw new Error("商品图展示框尺寸无效");
    }
    var scale = Math.min(targetWidth / currentWidth, targetHeight / currentHeight);
    if (Math.abs(scale - 1) > 0.001) {
        layer.resize(scale * 100, scale * 100, AnchorPosition.MIDDLECENTER);
    }
    currentRect = visiblePixelRect(layer);
    var currentCenterX = (currentRect.left + currentRect.right) / 2;
    var currentCenterY = (currentRect.top + currentRect.bottom) / 2;
    var targetCenterX = (targetRect.left + targetRect.right) / 2;
    var targetCenterY = (targetRect.top + targetRect.bottom) / 2;
    var shiftX = targetCenterX - currentCenterX;
    var shiftY = targetCenterY - currentCenterY;
    layer.translate(UnitValue(shiftX, "px"), UnitValue(shiftY, "px"));
    return Math.abs(scale - 1) > 0.001 || Math.abs(shiftX) > 0.5 || Math.abs(shiftY) > 0.5;
}

function fitOptionalImageToTemplateFrame(layer, targetRect) {
    // Optional new/old-package layers must preserve the exact placement made
    // by the PSD designer. "Replace Contents" can reset a layer transform
    // when the replacement asset has a different canvas size, so restore the
    // original placeholder frame after every successful replacement. The
    // source layer bounds are authoritative here: do not rasterise nested
    // smart objects or use their transparent-pixel bounds as a second frame.
    var currentRect = layerRect(layer);
    var currentWidth = currentRect.right - currentRect.left;
    var currentHeight = currentRect.bottom - currentRect.top;
    var targetWidth = targetRect.right - targetRect.left;
    var targetHeight = targetRect.bottom - targetRect.top;
    if (currentWidth <= 0 || currentHeight <= 0 || targetWidth <= 0 || targetHeight <= 0) {
        throw new Error("可选图层占位框尺寸无效");
    }
    var scale = Math.min(targetWidth / currentWidth, targetHeight / currentHeight);
    if (Math.abs(scale - 1) > 0.001) {
        layer.resize(scale * 100, scale * 100, AnchorPosition.MIDDLECENTER);
    }
    currentRect = layerRect(layer);
    var shiftX = (targetRect.left + targetRect.right - currentRect.left - currentRect.right) / 2;
    var shiftY = (targetRect.top + targetRect.bottom - currentRect.top - currentRect.bottom) / 2;
    layer.translate(UnitValue(shiftX, "px"), UnitValue(shiftY, "px"));
}

function markOptionalImageReplacement(result, key, replaced) {
    if (typeof result.optionalImageReplaced[key] === "undefined") {
        result.optionalImageReplaced[key] = replaced;
    } else {
        result.optionalImageReplaced[key] = result.optionalImageReplaced[key] && replaced;
    }
}

function setImageLayer(layer, value, key, record, materialIndex, result) {
    if (insideDisabledSwitch(layer, record)) {
        return;
    }
    if (isBlank(value) || isDisabledImageValue(value)) {
        if (isMainImageKey(key)) {
            if (dataFieldsOptional()) {
                addDataPrecheckWarning(result, "字段为空：商品图");
            } else {
                result.emptyField = true;
                addCode(result, "E_EMPTY_FIELD");
                addIssue(result, "字段为空：商品图");
            }
            layer.visible = false;
        } else {
            layer.visible = false;
            markOptionalImageReplacement(result, key, false);
        }
        return;
    }
    if (layer.typename !== "ArtLayer" || layer.kind !== LayerKind.SMARTOBJECT) {
        result.templateInvalid = true;
        addCode(result, "E_TEMPLATE_INVALID");
        addIssue(result, "模板错误：!" + key + " 不是智能对象");
        return;
    }
    var imageFile = findMaterial(value, materialIndex);
    if (!imageFile) {
        if (isMainImageKey(key)) {
            if (dataFieldsOptional()) {
                addDataPrecheckWarning(result, "缺图：商品图=" + value);
            } else {
                result.missingImage = true;
                addCode(result, "E_MISSING_IMAGE");
                addIssue(result, "缺图：商品图=" + value);
            }
            layer.visible = false;
        } else {
            layer.visible = false;
            result.optionalImageMissing = true;
            markOptionalImageReplacement(result, key, false);
            addCode(result, "W_OPTIONAL_IMAGE_MISSING");
            addIssue(result, "可选素材无法读取：" + key + "=" + value);
        }
        return;
    }
    try {
        var targetRect = layerRect(layer);
        var replacementFile = stageMaterialForPhotoshop(imageFile);
        var isolated = isolateSmartObjectForReplacement(layer);
        layer = isolated.layer;
        if (isolated.copied) {
            addIssue(result, "已隔离共享智能对象：!" + key);
        }
        replaceSmartObject(layer, replacementFile);
        if (isMainImageKey(key)) {
            var resized = fitProductToTemplateFrame(layer, targetRect);
            addIssue(result, resized ? "商品图已按 PSD 展示框定位" : "商品图已按 PSD 展示框居中");
        } else {
            fitOptionalImageToTemplateFrame(layer, targetRect);
            markOptionalImageReplacement(result, key, true);
            addIssue(result, "可选素材已替换：!" + key + "=" + imageFile.name);
        }
    } catch (error) {
        if (isMainImageKey(key)) {
            if (dataFieldsOptional()) {
                addDataPrecheckWarning(result, "替换失败：商品图=" + imageFile.name + "（" + error.message + "）");
                layer.visible = false;
            } else {
                result.missingImage = true;
                addCode(result, "E_MISSING_IMAGE");
                addIssue(result, "替换失败：商品图=" + imageFile.name + "（" + error.message + "）");
            }
        } else {
            layer.visible = false;
            result.optionalImageMissing = true;
            markOptionalImageReplacement(result, key, false);
            addCode(result, "W_OPTIONAL_IMAGE_MISSING");
                addIssue(result, "可选素材替换失败：" + key + "=" + imageFile.name + "（" + error.message + "）");
        }
    }
}

function applyRecord(document, record, materialIndex, result) {
    if (usesPhotoshopVariables(CHANNEL_PROFILE)) {
        applyPhotoshopVariables(document, record, materialIndex, result);
        applyPreflightIssue(record, result);
        return;
    }
    SWITCH_STATE = {};
    selectRecordLayout(document, record);
    var layerIndex = buildRecordLayerIndex(document);
    var dynamicWarnings = dynamicBindingWarnings(layerIndex, record);
    for (var dynamicWarningIndex = 0; dynamicWarningIndex < dynamicWarnings.length; dynamicWarningIndex++) {
        addCode(result, "W_BINDING_TABLE_ONLY");
        addIssue(result, dynamicWarnings[dynamicWarningIndex]);
    }
    var bindingErrors = recordBindingErrors(layerIndex, record);
    if (bindingErrors.length) {
        result.templateInvalid = true;
        addCode(result, "E_VAR_UNBOUND");
        for (var bindingIndex = 0; bindingIndex < bindingErrors.length; bindingIndex++) {
            addIssue(result, bindingErrors[bindingIndex]);
        }
        applyPreflightIssue(record, result);
        return;
    }
    setSwitches(layerIndex, record, result);

    for (var textKey in layerIndex.text) {
        if (!layerIndex.text.hasOwnProperty(textKey)) {
            continue;
        }
        var textLayers = layerIndex.text[textKey];
        for (var textIndex = 0; textIndex < textLayers.length; textIndex++) {
            var resolvedTextKey = resolveRecordKey(textKey, record);
            setTextLayer(textLayers[textIndex], recordValue(record, textKey), resolvedTextKey, record, result);
        }
    }
    for (var imageKey in layerIndex.image) {
        if (!layerIndex.image.hasOwnProperty(imageKey)) {
            continue;
        }
        var imageLayers = layerIndex.image[imageKey];
        for (var imageIndex = 0; imageIndex < imageLayers.length; imageIndex++) {
            var resolvedImageKey = resolveRecordKey(imageKey, record);
            setImageLayer(imageLayers[imageIndex], recordValue(record, imageKey), resolvedImageKey, record, materialIndex, result);
        }
    }
    reconcileOptionalGroups(layerIndex, record, result);
    reconcileGiftSlotVisibility(ACTIVE_LAYOUT_GROUP || document, record, result);
    applyPreflightIssue(record, result);
}

function recordBindingErrors(layerIndex, record) {
    if (!CHANNEL_PROFILE || !CHANNEL_PROFILE.required_psd_variables) {
        return [];
    }
    var errors = [];
    for (var index = 0; index < CHANNEL_PROFILE.required_psd_variables.length; index++) {
        var required = CHANNEL_PROFILE.required_psd_variables[index];
        var matches = requiredLayerMatches(layerIndex, required);
        if (matches.length > 0) {
            continue;
        }
        if ((isOptionalProfileVariable(CHANNEL_PROFILE, required) || dataFieldsOptional()) && isBlank(recordValue(record, required.name))) {
            continue;
        }
        errors.push("PSD 变量缺失：" + required.name + "（当前商品有对应数据）");
    }
    return errors;
}

function applyPhotoshopVariables(document, record, materialIndex, result) {
    var variables = [];
    var values = [];
    for (var index = 0; index < CHANNEL_PROFILE.required_psd_variables.length; index++) {
        var required = CHANNEL_PROFILE.required_psd_variables[index];
        var value = record[required.name];
        var variable = findDocumentVariable(document, required.name);
        if (!variable) {
            result.templateInvalid = true;
            addCode(result, "E_VAR_MISSING");
            addIssue(result, "PSD 变量缺失：" + required.name);
            continue;
        }
        if (variable.kind !== expectedVariableKind(required.type)) {
            result.templateInvalid = true;
            addCode(result, "E_VAR_TYPE_MISMATCH");
            addIssue(result, "PSD 变量类型不匹配：" + required.name);
            continue;
        }
        if (required.type === "text") {
            if (isBlank(value)) {
                if (dataFieldsOptional()) {
                    addDataPrecheckWarning(result, "字段为空：" + required.name);
                } else {
                    result.emptyField = true;
                    addCode(result, "E_EMPTY_FIELD");
                    addIssue(result, "字段为空：" + required.name);
                }
                continue;
            }
            variables.push(variable);
            values.push(String(value));
        } else {
            var imageFile = findMaterial(value, materialIndex);
            if (!imageFile) {
                if (dataFieldsOptional()) {
                    addDataPrecheckWarning(result, isBlank(value) ? "字段为空：" + required.name : "缺图：" + required.name + "=" + value);
                } else {
                    result.missingImage = true;
                    addCode(result, isBlank(value) ? "E_EMPTY_FIELD" : "E_MISSING_IMAGE");
                    addIssue(result, isBlank(value) ? "字段为空：" + required.name : "缺图：" + required.name + "=" + value);
                }
                continue;
            }
            variables.push(variable);
            values.push(stageMaterialForPhotoshop(imageFile));
        }
    }
    if (hasErrorCode(result)) { return; }
    try {
        // Data Set names are deliberately generated per run. Existing PSD data-set
        // names are reference content only and never participate in SKU matching.
        var dataSet = document.dataSets.add("__l0_" + safeOutputName(record["商品文件名"]) + "_" + (new Date().getTime()), variables, values);
        dataSet.apply();
        result.variableBindingStatus = "photoshop_variables_applied";
    } catch (error) {
        result.templateInvalid = true;
        addCode(result, "E_VAR_UNBOUND");
        addIssue(result, "PSD Variables 绑定或应用失败：" + error.message);
    }
}

function validateTargetSize(document) {
    var width = Math.round(document.width.as("px"));
    var height = Math.round(document.height.as("px"));
    var target = CHANNEL_PROFILE && CHANNEL_PROFILE.target_size;
    var acceptedSizes = CHANNEL_PROFILE && CHANNEL_PROFILE.accepted_template_sizes;
    if (!target && CHANNEL_PROFILE && CHANNEL_PROFILE.variants) {
        var variantName = CHANNEL_PROFILE.default_variant;
        if (variantName && CHANNEL_PROFILE.variants[variantName]) {
            target = CHANNEL_PROFILE.variants[variantName];
        }
    }
    if (acceptedSizes && acceptedSizes.length > 0) {
        var acceptedLabels = [];
        for (var acceptedIndex = 0; acceptedIndex < acceptedSizes.length; acceptedIndex++) {
            var accepted = acceptedSizes[acceptedIndex];
            if (!accepted || !accepted.width || !accepted.height) { continue; }
            acceptedLabels.push(accepted.width + "x" + accepted.height);
            if (width === accepted.width && height === accepted.height) {
                return;
            }
        }
        throw new Error("E_SIZE_MISMATCH: 模板尺寸 " + width + "x" + height + "，profile 允许 " + acceptedLabels.join("、"));
    }
    if (target && (width !== target.width || height !== target.height)) {
        throw new Error("E_SIZE_MISMATCH: 模板尺寸 " + width + "x" + height + "，profile 要求 " + target.width + "x" + target.height);
    }
    if (!target && width !== height) {
        throw new Error("E_SIZE_MISMATCH: 模板不是正方形，未强制拉伸：" + width + "x" + height);
    }
}

function exportJpeg(document, outputFolder, productName) {
    validateTargetSize(document);
    if (document.mode !== DocumentMode.RGB) {
        document.changeMode(ChangeMode.RGB);
    }
    if (document.bitsPerChannel !== BitsPerChannelType.EIGHT) {
        document.bitsPerChannel = BitsPerChannelType.EIGHT;
    }
    var outputFile = uniqueOutputFile(outputFolder, safeOutputName(productName), ".jpg");
    var options = new JPEGSaveOptions();
    options.quality = 10;
    options.embedColorProfile = true;
    options.formatOptions = FormatOptions.STANDARDBASELINE;
    options.matte = MatteType.NONE;
    document.saveAs(outputFile, options, true, Extension.LOWERCASE);
    return outputFile;
}

function exportPsd(document, outputFolder, productName) {
    var outputFile = uniqueOutputFile(outputFolder, safeOutputName(productName), ".psd");
    var options = new PhotoshopSaveOptions();
    document.saveAs(outputFile, options, true, Extension.LOWERCASE);
    return outputFile;
}

function uniqueOutputFile(outputFolder, baseName, extension) {
    var cleanBase = isBlank(baseName) ? "未命名商品" : baseName;
    var cleanExtension = extension || ".jpg";
    var candidate = File(outputFolder.fsName + "/" + cleanBase + cleanExtension);
    var suffix = 1;
    while (candidate.exists) {
        candidate = File(outputFolder.fsName + "/" + cleanBase + "_" + suffix + cleanExtension);
        suffix++;
    }
    return candidate;
}

function statusFor(result) {
    if (result.validationFailed) {
        return "数据需核对";
    }
    if (result.processingFailed) {
        return "处理失败";
    }
    if (result.templateInvalid) {
        return "模板错误";
    }
    if (result.missingImage) {
        return "缺图";
    }
    if (result.emptyField) {
        return "字段为空";
    }
    if (result.preflightIssue) {
        return "需复核";
    }
    if (result.optionalImageMissing) {
        return "需复核";
    }
    if (result.textOverflow) {
        return "文案超框";
    }
    return "成功";
}

function hasErrorCode(result) {
    for (var index = 0; index < result.codes.length; index++) {
        if (startsWith(result.codes[index], "E_")) { return true; }
    }
    return false;
}

function severityFor(result) {
    if (hasErrorCode(result) || result.processingFailed || result.templateInvalid || result.validationFailed) { return "E"; }
    if (result.codes.length || result.preflightIssue || result.textOverflow || result.optionalImageMissing) { return "W"; }
    return "OK";
}

function writeReport(outputFolder, results) {
    var reportFile = File(outputFolder.fsName + "/" + REPORT_NAME);
    reportFile.encoding = "UTF8";
    if (!reportFile.open("w")) {
        throw new Error("无法写入结果报告：" + reportFile.fsName);
    }
    reportFile.write("\uFEFF货号,商品文件名,profile_id,profile_version,severity,variable_binding_status,状态,错误码,中文说明,建议动作,输出文件,输出PSD,未导出\n");
    for (var index = 0; index < results.length; index++) {
        var result = results[index];
        reportFile.write(
            csvEscape(result.sku) + "," +
            csvEscape(result.productName) + "," +
            csvEscape(result.profileId) + "," +
            csvEscape(result.profileVersion) + "," +
            csvEscape(severityFor(result)) + "," +
            csvEscape(result.variableBindingStatus) + "," +
            csvEscape(statusFor(result)) + "," +
            csvEscape(result.codes.join(";")) + "," +
            csvEscape(result.issues.join("；")) + "," +
            csvEscape(suggestAction(result)) + "," +
            csvEscape(result.outputFile) + "," +
            csvEscape(result.psdOutputFile) + "," +
            csvEscape(isBlank(result.outputFile) && isBlank(result.psdOutputFile) ? "是" : "否") + "\n"
        );
    }
    reportFile.close();
    return reportFile;
}

function suggestAction(result) {
    if (result.validationFailed) {
        return "核对 Excel 价格字段后重新生成 data.csv";
    }
    if (result.templateInvalid) {
        return "检查 PSD 图层命名和智能对象规范";
    }
    if (result.processingFailed) {
        return "查看中文说明并联系开发排查 Photoshop 执行错误";
    }
    if (result.missingImage) {
        return "核对 Excel 图片字段的完整路径和文件可访问性后复跑失败行";
    }
    if (result.emptyField) {
        return "补齐 Excel 对应字段，或确认开关字段为是/否";
    }
    if (result.preflightIssue) {
        return "已按表格内容继续生成；请核对清洗预检说明后决定是否修正并复跑";
    }
    if (result.optionalImageMissing) {
        return "可选素材未展示：核对 Excel 图片路径和 PSD 对应智能对象";
    }
    if (result.textOverflow) {
        return "设计师检查 PSD 文本框或调整文案长度";
    }
    return "无需处理";
}

function processRecord(template, record, materialIndex, outputFolder, psdOutputFolder) {
    var productName = trimText(record["商品文件名"]);
    var result = makeResult(productName || "未命名商品", record);
    if (isBlank(productName)) {
        result.emptyField = true;
        addCode(result, "E_EMPTY_FIELD");
        addIssue(result, "字段为空：商品文件名");
        return result;
    }
    var priceError = priceValidationError(record);
    if (priceError) {
        if (CONTINUE_WITH_PREFLIGHT_ISSUES) {
            result.preflightIssue = true;
            addCode(result, "W_DATA_PRECHECK");
            addIssue(result, "价格格式异常：" + priceError + "；已按表格内容继续生成");
        } else {
            result.validationFailed = true;
            addCode(result, "E_PRICE_INVALID");
            addIssue(result, "价格格式异常：" + priceError + "；已阻止输出");
            return result;
        }
    }
    var copy = null;
    try {
        copy = template.duplicate();
        app.activeDocument = copy;
        applyRecord(copy, record, materialIndex, result);
        if (hasErrorCode(result)) {
            return result;
        }
        validateTargetSize(copy);
        var psdOutputFile = exportPsd(copy, psdOutputFolder, productName);
        result.psdOutputFile = psdOutputFile.fsName;
        var outputFile = exportJpeg(copy, outputFolder, productName);
        result.outputFile = outputFile.fsName;
    } catch (error) {
        result.processingFailed = true;
        addCode(result, String(error.message).indexOf("E_SIZE_MISMATCH") >= 0 ? "E_SIZE_MISMATCH" : "E_TEMPLATE_INVALID");
        addIssue(result, "处理失败：" + error.message);
    } finally {
        if (copy) {
            try {
                copy.close(SaveOptions.DONOTSAVECHANGES);
            } catch (closeError) {}
        }
    }
    return result;
}

var runProgressWindow = null;
var runProgressLabel = null;
var runProgressBar = null;

function startRunProgress(total) {
    if (!automatedRun) {
        return;
    }
    try {
        runProgressWindow = new Window("palette", "电商主图批量套版正在执行");
        runProgressWindow.orientation = "column";
        runProgressWindow.alignChildren = "fill";
        runProgressWindow.margins = 18;
        runProgressWindow.add("statictext", undefined, "Photoshop 正在处理，请勿关闭 Photoshop 或任务窗口。");
        runProgressLabel = runProgressWindow.add("statictext", undefined, "准备导出 0 / " + total);
        runProgressLabel.preferredSize.width = 440;
        runProgressBar = runProgressWindow.add("progressbar", undefined, 0, total);
        runProgressBar.preferredSize.width = 440;
        runProgressWindow.show();
        app.refresh();
        $.sleep(50);
    } catch (progressError) {
        runProgressWindow = null;
    }
}

function updateRunProgress(current, total, productName) {
    if (!runProgressWindow) {
        return;
    }
    try {
        runProgressLabel.text = "正在导出 " + current + " / " + total + "：" + productName;
        runProgressBar.value = current;
        runProgressWindow.layout.layout(true);
        app.refresh();
        $.sleep(30);
    } catch (progressError) {}
}

function closeRunProgress() {
    if (!runProgressWindow) {
        return;
    }
    try {
        runProgressWindow.close();
    } catch (progressError) {}
    runProgressWindow = null;
}

function main() {
    if (app.documents.length === 0) {
        throw new Error("请先打开并激活已按 @、!、# 规范命名的 PSD 模板。");
    }
    var template = app.activeDocument;
    validateTemplate(template);
    var automatedInputs = (typeof $.global.__BATCH_INPUTS__ !== "undefined") ? $.global.__BATCH_INPUTS__ : null;
    var csvFile = automatedInputs ? File(automatedInputs.csv) : File.openDialog("请选择 data.csv", "*.csv");
    if (!csvFile) {
        return;
    }
    // Automated L0 jobs use the complete image paths held in Excel.  The
    // manual folder picker remains only as a compatibility fallback.
    var materialFolder = automatedInputs ? null : Folder.selectDialog("请选择本地素材文件夹");
    if (!automatedInputs && !materialFolder) {
        return;
    }
    var outputFolder = automatedInputs ? Folder(automatedInputs.output) : Folder.selectDialog("请选择输出文件夹");
    if (!outputFolder) {
        return;
    }
    var psdOutputFolder = automatedInputs && automatedInputs.psdOutput ? Folder(automatedInputs.psdOutput) : Folder(outputFolder.fsName + "/成品PSD");
    if (!psdOutputFolder.exists && !psdOutputFolder.create()) {
        throw new Error("无法创建 PSD 输出文件夹：" + psdOutputFolder.fsName);
    }
    if (!csvFile.exists || !outputFolder.exists || !psdOutputFolder.exists || (materialFolder && !materialFolder.exists)) {
        throw new Error("自动化输入路径不存在，请检查 CSV、JPG 输出文件夹、PSD 输出文件夹，以及（手动模式下）素材文件夹。");
    }
    var records = readCsv(csvFile);
    if (!records.length) {
        throw new Error("data.csv 没有可处理记录。");
    }
    var materialIndex = null;
    materialIndexSkipCount = 0;
    if (materialFolder) {
        materialIndex = {};
        addFilesToIndex(materialFolder, materialIndex);
    }

    var results = [];
    var acceptedCount = 0;
    var exportedCount = 0;
    var psdExportedCount = 0;
    startRunProgress(records.length);
    for (var index = 0; index < records.length; index++) {
        updateRunProgress(index + 1, records.length, trimText(records[index]["商品文件名"]));
        var result = processRecord(template, records[index], materialIndex, outputFolder, psdOutputFolder);
        results.push(result);
        // Persist after every row so a later Photoshop error does not lose earlier results.
        writeReport(outputFolder, results);
        if (!isBlank(result.outputFile)) {
            exportedCount++;
        }
        if (!isBlank(result.psdOutputFile)) {
            psdExportedCount++;
        }
        if (statusFor(result) === "成功") {
            acceptedCount++;
        }
    }
    var reportFile = writeReport(outputFolder, results);
    var summary = "已导出 " + exportedCount + " 张 JPG / " + psdExportedCount + " 份 PSD / 合格 " + acceptedCount + " 张 / 需复核 " + (results.length - acceptedCount) + " 条\n结果报告：" + reportFile.fsName;
    if (materialIndexSkipCount > 0) {
        summary += "\n素材索引提示：已跳过 " + materialIndexSkipCount + " 个不可访问目录/文件，其他素材继续处理。";
    }
    if (automatedInputs) {
        $.global.__BATCH_RESULT__ = summary;
    } else {
        alert("批量处理完成\n" + summary);
    }
    return summary;
}

var previousDialogs = app.displayDialogs;
var automatedRun = (typeof $.global.__BATCH_INPUTS__ !== "undefined");
try {
    app.displayDialogs = DialogModes.NO;
    main();
} catch (error) {
    if (automatedRun) {
        $.global.__BATCH_RESULT__ = "批量套版未完成：" + error.message;
        throw error;
    }
    alert("批量套版未完成：" + error.message);
} finally {
    closeRunProgress();
    app.displayDialogs = previousDialogs;
}
