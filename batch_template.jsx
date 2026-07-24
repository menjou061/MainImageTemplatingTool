#target photoshop

/*
 * Run from Photoshop: File > Scripts > Browse.
 * Keep the renamed 800x800 PSD open and active before running this script.
 */

var REPORT_NAME = "结果报告.csv";
var REQUIRED_SIZE = 800;
var PRODUCT_SAFE_SCALE = 0.94;

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
        result += illegal.indexOf(character) >= 0 ? "_" : character;
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
        throw new Error("所选文件不是 data.csv：缺少“商品文件名”列表头。");
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

function addFilesToIndex(folder, index) {
    var items = folder.getFiles();
    for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
        var item = items[itemIndex];
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
    }
}

function findMaterial(value, materialIndex) {
    if (isBlank(value)) {
        return null;
    }
    var key = decodedName(getBasename(value)).toLowerCase();
    return materialIndex[key] || null;
}

function addIssue(result, issue) {
    result.issues.push(issue);
}

function containsIssue(result, issue) {
    for (var index = 0; index < result.issues.length; index++) {
        if (result.issues[index] === issue) {
            return true;
        }
    }
    return false;
}

function makeResult(productName) {
    return { productName: productName, issues: [], missingImage: false, emptyField: false, textOverflow: false, validationFailed: false, processingFailed: false, outputFile: "" };
}

function priceValidationError(record) {
    var price1 = trimText(record["价格1"]);
    var price2 = trimText(record["价格2"]);
    if (isBlank(price1) || isBlank(price2)) {
        return "价格1、价格2必须完整填写";
    }
    return "";
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

function validateTemplate(template) {
    var layerIndex = { text: {}, image: {}, switches: {} };
    addLayers(template, layerIndex);
    if (keyCount(layerIndex.text) === 0 || keyCount(layerIndex.image) === 0) {
        throw new Error("当前 PSD 尚未按规范完成改造：至少需要一个 @文本层和一个 !智能对象层。");
    }
}

function isYes(value) {
    var normalized = trimText(value).toLowerCase();
    return normalized === "是" || normalized === "yes" || normalized === "true" || normalized === "1" || normalized === "y";
}

function isNo(value) {
    var normalized = trimText(value).toLowerCase();
    return normalized === "否" || normalized === "no" || normalized === "false" || normalized === "0" || normalized === "n";
}

function setSwitches(layerIndex, record, result) {
    for (var key in layerIndex.switches) {
        if (!layerIndex.switches.hasOwnProperty(key)) {
            continue;
        }
        var value = record[key];
        var hasConfiguredSwitch = record.hasOwnProperty(key);
        var visible = false;
        if (isYes(value)) {
            visible = true;
        } else if (isNo(value)) {
            visible = false;
        } else if (!hasConfiguredSwitch) {
            // Optional # groups not supplied by this sheet are hidden by default.
            visible = false;
        } else {
            result.emptyField = true;
            addIssue(result, "开关字段为空或无效：" + key);
        }
        var layers = layerIndex.switches[key];
        for (var index = 0; index < layers.length; index++) {
            layers[index].visible = visible;
        }
    }
}

function insideDisabledSwitch(layer, record) {
    var parent = layer.parent;
    while (parent && parent.typename !== "Document") {
        if (startsWith(parent.name, "#") && !isYes(record[parent.name.substring(1)])) {
            return true;
        }
        parent = parent.parent;
    }
    return false;
}

function layerWidth(layer) {
    var bounds = layer.bounds;
    return bounds[2].as("px") - bounds[0].as("px");
}

function fitPointTextToCanvas(layer) {
    if (layer.textItem.kind !== TextType.POINTTEXT) {
        return { fitted: false, overflow: false };
    }
    var canvasWidth = app.activeDocument.width.as("px");
    var margin = 8;
    var initialSize = layer.textItem.size.as("pt");
    var minimumSize = initialSize * 0.55;
    var currentSize = initialSize;
    var fitted = false;
    var attempts = 0;
    var bounds = layer.bounds;
    while ((bounds[0].as("px") < margin || bounds[2].as("px") > canvasWidth - margin) && currentSize > minimumSize && attempts < 24) {
        currentSize = Math.max(minimumSize, currentSize * 0.92);
        layer.textItem.size = UnitValue(currentSize, "pt");
        fitted = true;
        attempts++;
        bounds = layer.bounds;
    }
    return {
        fitted: fitted,
        overflow: bounds[0].as("px") < margin || bounds[2].as("px") > canvasWidth - margin
    };
}

function paragraphTextWidth(layer) {
    try {
        if (layer.textItem.kind === TextType.PARAGRAPHTEXT) {
            return layer.textItem.width.as("px");
        }
    } catch (error) {}
    return null;
}

function setTextLayer(layer, value, key, record, result) {
    if (insideDisabledSwitch(layer, record)) {
        return;
    }
    if (isBlank(value)) {
        layer.textItem.contents = "";
        if (isRequiredTextKey(key)) {
            result.emptyField = true;
            addIssue(result, "字段为空：" + key);
        }
        return;
    }
    layer.textItem.contents = String(value);
    var textBoxWidth = paragraphTextWidth(layer);
    var fitResult = fitPointTextToCanvas(layer);
    if (fitResult.fitted) {
        addIssue(result, "文字已自动缩小：" + key);
    }
    if (fitResult.overflow || (textBoxWidth !== null && layerWidth(layer) > textBoxWidth + 0.5)) {
        result.textOverflow = true;
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

function fitProductToSafeFrame(layer, targetRect) {
    var currentRect = layerRect(layer);
    var currentWidth = currentRect.right - currentRect.left;
    var currentHeight = currentRect.bottom - currentRect.top;
    var targetWidth = (targetRect.right - targetRect.left) * PRODUCT_SAFE_SCALE;
    var targetHeight = (targetRect.bottom - targetRect.top) * PRODUCT_SAFE_SCALE;
    if (currentWidth <= 0 || currentHeight <= 0 || targetWidth <= 0 || targetHeight <= 0) {
        throw new Error("商品图安全框尺寸无效");
    }
    var scale = Math.min(1, targetWidth / currentWidth, targetHeight / currentHeight);
    if (scale < 0.999) {
        layer.resize(scale * 100, scale * 100, AnchorPosition.MIDDLECENTER);
    }
    currentRect = layerRect(layer);
    var currentCenterX = (currentRect.left + currentRect.right) / 2;
    var currentCenterY = (currentRect.top + currentRect.bottom) / 2;
    var targetCenterX = (targetRect.left + targetRect.right) / 2;
    var targetCenterY = (targetRect.top + targetRect.bottom) / 2;
    layer.translate(UnitValue(targetCenterX - currentCenterX, "px"), UnitValue(targetCenterY - currentCenterY, "px"));
    return scale < 0.999;
}

function setImageLayer(layer, value, key, record, materialIndex, result) {
    if (insideDisabledSwitch(layer, record)) {
        return;
    }
    if (isDisabledImageValue(value)) {
        return;
    }
    if (isBlank(value)) {
        result.emptyField = true;
        addIssue(result, "字段为空：" + key);
        return;
    }
    if (layer.typename !== "ArtLayer" || layer.kind !== LayerKind.SMARTOBJECT) {
        result.emptyField = true;
        addIssue(result, "模板错误：!" + key + " 不是智能对象");
        return;
    }
    var imageFile = findMaterial(value, materialIndex);
    if (!imageFile) {
        result.missingImage = true;
        addIssue(result, "缺图：" + key + "=" + value);
        return;
    }
    try {
        var targetRect = key === "商品图" ? layerRect(layer) : null;
        replaceSmartObject(layer, imageFile);
        if (targetRect) {
            var resized = fitProductToSafeFrame(layer, targetRect);
            addIssue(result, resized ? "商品图已缩小并居中" : "商品图已居中");
        }
    } catch (error) {
        result.missingImage = true;
        addIssue(result, "替换失败：" + key + "=" + imageFile.name + "（" + error.message + "）");
    }
}

function applyRecord(document, record, materialIndex, result) {
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
    if (String(record["预检异常"] || "").indexOf("字段为空") !== -1) {
        result.emptyField = true;
        addIssue(result, "清洗预检：字段为空");
    }
}

function exportJpeg(document, outputFolder, productName) {
    var width = Math.round(document.width.as("px"));
    var height = Math.round(document.height.as("px"));
    if (width !== height) {
        throw new Error("模板不是正方形，未强制拉伸：" + width + "x" + height);
    }
    if (width !== REQUIRED_SIZE || height !== REQUIRED_SIZE) {
        document.resizeImage(UnitValue(REQUIRED_SIZE, "px"), UnitValue(REQUIRED_SIZE, "px"), null, ResampleMethod.BICUBIC);
    }
    if (document.mode !== DocumentMode.RGB) {
        document.changeMode(ChangeMode.RGB);
    }
    if (document.bitsPerChannel !== BitsPerChannelType.EIGHT) {
        document.bitsPerChannel = BitsPerChannelType.EIGHT;
    }
    var outputFile = File(outputFolder.fsName + "/" + safeOutputName(productName) + ".jpg");
    var options = new ExportOptionsSaveForWeb();
    options.format = SaveDocumentType.JPEG;
    options.quality = 90;
    options.optimized = true;
    options.includeProfile = true;
    document.exportDocument(outputFile, ExportType.SAVEFORWEB, options);
    return outputFile;
}

function statusFor(result) {
    if (result.validationFailed) {
        return "数据需核对";
    }
    if (result.processingFailed) {
        return "处理失败";
    }
    if (result.missingImage) {
        return "缺图";
    }
    if (result.emptyField) {
        return "字段为空";
    }
    if (result.textOverflow) {
        return "文案超框";
    }
    return "成功";
}

function writeReport(outputFolder, results) {
    var reportFile = File(outputFolder.fsName + "/" + REPORT_NAME);
    reportFile.encoding = "UTF8";
    if (!reportFile.open("w")) {
        throw new Error("无法写入结果报告：" + reportFile.fsName);
    }
    reportFile.write("\uFEFF商品文件名,状态,详情,输出文件\n");
    for (var index = 0; index < results.length; index++) {
        var result = results[index];
        reportFile.write(
            csvEscape(result.productName) + "," +
            csvEscape(statusFor(result)) + "," +
            csvEscape(result.issues.join("；")) + "," +
            csvEscape(result.outputFile) + "\n"
        );
    }
    reportFile.close();
    return reportFile;
}

function processRecord(template, record, materialIndex, outputFolder) {
    var productName = trimText(record["商品文件名"]);
    var result = makeResult(productName || "未命名商品");
    if (isBlank(productName)) {
        result.emptyField = true;
        addIssue(result, "字段为空：商品文件名");
        return result;
    }
    var priceError = priceValidationError(record);
    if (priceError) {
        result.validationFailed = true;
        addIssue(result, "价格格式异常：" + priceError + "；已阻止输出");
        return result;
    }
    var copy = null;
    try {
        copy = template.duplicate();
        app.activeDocument = copy;
        applyRecord(copy, record, materialIndex, result);
        var outputFile = exportJpeg(copy, outputFolder, productName);
        result.outputFile = outputFile.fsName;
    } catch (error) {
        result.processingFailed = true;
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
    var materialFolder = automatedInputs ? Folder(automatedInputs.materials) : Folder.selectDialog("请选择本地素材文件夹");
    if (!materialFolder) {
        return;
    }
    var outputFolder = automatedInputs ? Folder(automatedInputs.output) : Folder.selectDialog("请选择输出文件夹");
    if (!outputFolder) {
        return;
    }
    if (!csvFile.exists || !materialFolder.exists || !outputFolder.exists) {
        throw new Error("自动化输入路径不存在，请检查 CSV、素材文件夹和输出文件夹。");
    }
    var records = readCsv(csvFile);
    if (!records.length) {
        throw new Error("data.csv 没有可处理记录。");
    }
    var materialIndex = {};
    addFilesToIndex(materialFolder, materialIndex);

    var results = [];
    var acceptedCount = 0;
    var exportedCount = 0;
    for (var index = 0; index < records.length; index++) {
        var result = processRecord(template, records[index], materialIndex, outputFolder);
        results.push(result);
        // Persist after every row so a later Photoshop error does not lose earlier results.
        writeReport(outputFolder, results);
        if (!isBlank(result.outputFile)) {
            exportedCount++;
        }
        if (statusFor(result) === "成功") {
            acceptedCount++;
        }
    }
    var reportFile = writeReport(outputFolder, results);
    var summary = "已导出 " + exportedCount + " 张 / 合格 " + acceptedCount + " 张 / 需复核 " + (results.length - acceptedCount) + " 条\n结果报告：" + reportFile.fsName;
    if (automatedInputs) {
        $.global.__BATCH_RESULT__ = summary;
    } else {
        alert("批量处理完成\n" + summary);
        outputFolder.execute();
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
    app.displayDialogs = previousDialogs;
}
