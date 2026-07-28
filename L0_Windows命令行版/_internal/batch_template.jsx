#target photoshop

/*
 * Run from Photoshop: File > Scripts > Browse.
 * Keep the renamed PSD template open and active before running this script.
 */

var REPORT_NAME = "结果报告.csv";
var PRODUCT_VERTICAL_OFFSET_PX = 32;
var CONTINUE_WITH_PREFLIGHT_ISSUES = !!($.global.__BATCH_INPUTS__ && $.global.__BATCH_INPUTS__.continueWithPreflightIssues);
var CHANNEL_PROFILE = ($.global.__BATCH_INPUTS__ && $.global.__BATCH_INPUTS__.profile) || null;

function trimText(value) {
    return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function startsWith(value, prefix) {
    return String(value).indexOf(prefix) === 0;
}

function isBlank(value) {
    return trimText(value) === "";
}

function isDisabledImageValue(value) {
    var normalized = trimText(value).toLowerCase();
    return normalized === "无" || normalized === "无.png" || normalized === "none" || normalized === "null";
}

function isRequiredTextKey(key) {
    return key === "折扣" || key === "券名" || key === "券门槛" ||
        key === "活动时间" || key === "到手" || key === "价格1" ||
        key === "价格2" || key === "卖点" || key === "规格";
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
    var key = decodedName(getBasename(value)).toLowerCase();
    return materialIndex ? (materialIndex[key] || null) : null;
}

function addIssue(result, issue) {
    result.issues.push(issue);
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

function priceValidationError(record) {
    var price1 = trimText(record["价格1"]);
    var price2 = trimText(record["价格2"]);
    if (isBlank(price1) || isBlank(price2)) {
        return "价格1、价格2必须完整填写";
    }
    return "";
}

function applyPreflightIssue(record, result) {
    var issue = trimText(record["预检异常"]);
    if (isBlank(issue)) {
        return;
    }
    result.preflightIssue = true;
    addCode(result, "W_DATA_PRECHECK");
    addIssue(result, "清洗预检：" + issue);
}

function addLayers(container, layerIndex) {
    for (var index = 0; index < container.layers.length; index++) {
        var layer = container.layers[index];
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
        if (layer.typename === "LayerSet") {
            addLayers(layer, layerIndex);
        }
    }
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
    var layerIndex = { text: {}, image: {}, switches: {} };
    addLayers(template, layerIndex);
    var errors = [];
    for (var index = 0; index < profile.required_psd_variables.length; index++) {
        var required = profile.required_psd_variables[index];
        var collection = required.type === "text" ? layerIndex.text : layerIndex.image;
        if (!collection[required.name] || collection[required.name].length === 0) {
            errors.push("E_VAR_UNBOUND: " + required.name + " 未绑定到 " + (required.type === "text" ? "@文本层" : "!智能对象"));
            continue;
        }
        for (var layerIndexValue = 0; layerIndexValue < collection[required.name].length; layerIndexValue++) {
            var layer = collection[required.name][layerIndexValue];
            if ((required.type === "text" && layer.kind !== LayerKind.TEXT) ||
                (required.type === "smart_object" && layer.kind !== LayerKind.SMARTOBJECT)) {
                errors.push("E_VAR_TYPE_MISMATCH: " + required.name);
                break;
            }
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
    if (keyCount(layerIndex.text) === 0 || keyCount(layerIndex.image) === 0) {
        throw new Error("当前 PSD 尚未按规范完成改造：至少需要一个 @文本层和一个 !智能对象层。");
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
        if (key === "优惠券" || key === "优惠券开关") {
            value = switchValue(record, key);
        } else {
            // Show the group long enough to try each child smart-object
            // replacement. File accessibility is confirmed by setImageLayer;
            // checking an SMB path here used to silently hide an entire group.
            value = optionalGroupHasValues(layerIndex.switches[key], record) ? "是" : "否";
            record[key] = value;
        }
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

function optionalGroupHasValues(groups, record) {
    for (var index = 0; index < groups.length; index++) {
        if (groups[index].typename === "LayerSet" && groupHasAllImageValues(groups[index], record)) {
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
            var imageValue = record[child.name.substring(1)];
            if (isBlank(imageValue) || isDisabledImageValue(imageValue)) {
                return false;
            }
        } else if (child.typename === "LayerSet" && !startsWith(child.name, "#")) {
            if (!groupHasAllImageValues(child, record)) {
                return false;
            }
            foundImage = true;
        }
    }
    return foundImage;
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
        } else if (child.typename === "LayerSet" && !startsWith(child.name, "#")) {
            if (!groupHasAllReplacedImageLayers(child, result)) {
                return false;
            }
            foundImage = true;
        }
    }
    return foundImage;
}

function reconcileOptionalGroups(layerIndex, record, result) {
    for (var key in layerIndex.switches) {
        if (!layerIndex.switches.hasOwnProperty(key) || key === "优惠券" || key === "优惠券开关") {
            continue;
        }
        var groups = layerIndex.switches[key];
        var requested = isYes(record[key]);
        if (!requested) {
            for (var hiddenIndex = 0; hiddenIndex < groups.length; hiddenIndex++) {
                groups[hiddenIndex].visible = false;
            }
            record[key] = "否";
            continue;
        }
        var allReplaced = false;
        for (var index = 0; index < groups.length; index++) {
            if (groups[index].typename === "LayerSet" && groupHasAllReplacedImageLayers(groups[index], result)) {
                allReplaced = true;
                break;
            }
        }
        for (var groupIndex = 0; groupIndex < groups.length; groupIndex++) {
            groups[groupIndex].visible = allReplaced;
        }
        record[key] = allReplaced ? "是" : "否";
        if (allReplaced) {
            addIssue(result, "可选图层已展示：#" + key);
        } else if (requested) {
            addIssue(result, "可选图层已隐藏：#" + key);
        }
    }
}

function insideDisabledSwitch(layer, record) {
    var parent = layer.parent;
    while (parent && parent.typename !== "Document") {
        if (startsWith(parent.name, "#") && !isYes(switchValue(record, parent.name.substring(1)))) {
            return true;
        }
        parent = parent.parent;
    }
    return false;
}

function switchValue(record, key) {
    var value = record[key];
    if (key === "优惠券" && isBlank(value) && !isBlank(record["优惠券开关"])) {
        return record["优惠券开关"];
    }
    return value;
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
    var adjacentBoundary = null;
    if (key === "折扣") {
        adjacentBoundary = siblingLeftBoundary(layer, ["折"], originalRect);
    } else if (key === "价格1") {
        adjacentBoundary = siblingLeftBoundary(layer, ["@价格2"], originalRect);
    }
    if (adjacentBoundary !== null) {
        maxWidth = Math.min(maxWidth, Math.max(1, adjacentBoundary - originalRect.left - 4));
    }
    return Math.max(1, maxWidth);
}

function fitTextToOriginalFrame(layer, key, originalRect) {
    var maxWidth = textMaxWidth(layer, key, originalRect);
    var maxHeight = Math.max(1, rectHeight(originalRect) * 0.98);
    var minimumSize = 6;
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
        var currentSize;
        try {
            currentSize = layer.textItem.size.as("pt");
        } catch (sizeError) {
            break;
        }
        var nextSize = Math.max(minimumSize, currentSize * factor);
        if (nextSize >= currentSize - 0.01) {
            break;
        }
        layer.textItem.size = UnitValue(nextSize, "pt");
    }

    // Preserve the template's visual anchor after a point-text resize.
    var finalRect = layerRect(layer);
    layer.translate(
        UnitValue(originalRect.left - finalRect.left, "px"),
        UnitValue(originalRect.top - finalRect.top, "px")
    );
    finalRect = layerRect(layer);
    return {
        fitted: rectWidth(finalRect) <= maxWidth + 0.5 && rectHeight(finalRect) <= maxHeight + 0.5,
        maxWidth: maxWidth,
        finalRect: finalRect
    };
}

function setTextLayer(layer, value, key, record, result) {
    if (insideDisabledSwitch(layer, record)) {
        return;
    }
    if (isBlank(value)) {
        layer.textItem.contents = "";
        if (isRequiredTextKey(key)) {
            result.emptyField = true;
            addCode(result, "E_EMPTY_FIELD");
            addIssue(result, "字段为空：" + key);
        }
        return;
    }
    var originalRect = layerRect(layer);
    layer.textItem.contents = String(value);
    // Keep the template's font, size, style and original anchor unchanged.
    // Uneven automated shrinking made the same price/copy fields look wildly
    // different between products. A long value is reported for design review
    // instead of changing visual hierarchy during export.
    var textRect = layerRect(layer);
    var maxWidth = textMaxWidth(layer, key, originalRect);
    var maxHeight = Math.max(1, rectHeight(originalRect) * 0.98);
    var canvas = app.activeDocument;
    var outOfCanvas = textRect.left < -0.5 || textRect.top < -0.5 ||
        textRect.right > canvas.width.as("px") + 0.5 ||
        textRect.bottom > canvas.height.as("px") + 0.5;
    if (rectWidth(textRect) > maxWidth + 0.5 || rectHeight(textRect) > maxHeight + 0.5 || outOfCanvas) {
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
        return rect;
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
    var targetWidth = targetRect.right - targetRect.left;
    var targetHeight = targetRect.bottom - targetRect.top;
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
    var shiftY = targetCenterY - currentCenterY + PRODUCT_VERTICAL_OFFSET_PX;
    layer.translate(UnitValue(shiftX, "px"), UnitValue(shiftY, "px"));
    return Math.abs(scale - 1) > 0.001 || Math.abs(shiftX) > 0.5 || Math.abs(shiftY) > 0.5;
}

function fitOptionalImageToTemplateFrame(layer, targetRect) {
    // Optional new/old-package layers must preserve the exact placement made
    // by the PSD designer. "Replace Contents" can reset a layer transform
    // when the replacement asset has a different canvas size, so restore the
    // original placeholder frame after every successful replacement.
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
        if (key === "商品图") {
            result.emptyField = true;
            addCode(result, "E_EMPTY_FIELD");
            addIssue(result, "字段为空：商品图");
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
        if (key === "商品图") {
            result.missingImage = true;
            addCode(result, "E_MISSING_IMAGE");
            addIssue(result, "缺图：商品图=" + value);
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
        replaceSmartObject(layer, imageFile);
        if (key === "商品图") {
            var resized = fitProductToTemplateFrame(layer, targetRect);
            addIssue(result, resized ? "商品图已按 PSD 展示框定位" : "商品图已按 PSD 展示框居中");
        } else {
            fitOptionalImageToTemplateFrame(layer, targetRect);
            markOptionalImageReplacement(result, key, true);
            addIssue(result, "可选素材已替换：!" + key + "=" + imageFile.name);
        }
    } catch (error) {
        if (key === "商品图") {
            result.missingImage = true;
            addCode(result, "E_MISSING_IMAGE");
            addIssue(result, "替换失败：商品图=" + imageFile.name + "（" + error.message + "）");
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
    var layerIndex = { text: {}, image: {}, switches: {} };
    addLayers(document, layerIndex);
    setSwitches(layerIndex, record, result);

    for (var textKey in layerIndex.text) {
        if (!layerIndex.text.hasOwnProperty(textKey)) {
            continue;
        }
        var textLayers = layerIndex.text[textKey];
        for (var textIndex = 0; textIndex < textLayers.length; textIndex++) {
            setTextLayer(textLayers[textIndex], record[textKey], textKey, record, result);
        }
    }
    for (var imageKey in layerIndex.image) {
        if (!layerIndex.image.hasOwnProperty(imageKey)) {
            continue;
        }
        var imageLayers = layerIndex.image[imageKey];
        for (var imageIndex = 0; imageIndex < imageLayers.length; imageIndex++) {
            setImageLayer(imageLayers[imageIndex], record[imageKey], imageKey, record, materialIndex, result);
        }
    }
    reconcileOptionalGroups(layerIndex, record, result);
    applyPreflightIssue(record, result);
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
                result.emptyField = true;
                addCode(result, "E_EMPTY_FIELD");
                addIssue(result, "字段为空：" + required.name);
                continue;
            }
            variables.push(variable);
            values.push(String(value));
        } else {
            var imageFile = findMaterial(value, materialIndex);
            if (!imageFile) {
                result.missingImage = true;
                addCode(result, isBlank(value) ? "E_EMPTY_FIELD" : "E_MISSING_IMAGE");
                addIssue(result, isBlank(value) ? "字段为空：" + required.name : "缺图：" + required.name + "=" + value);
                continue;
            }
            variables.push(variable);
            values.push(imageFile);
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
