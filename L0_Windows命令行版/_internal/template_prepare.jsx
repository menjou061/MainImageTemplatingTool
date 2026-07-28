#target photoshop

/*
 * PSD template preflight and one-time copy-based preparation.
 * PowerShell opens the selected PSD, sets __TEMPLATE_PREP_INPUTS__, and reads
 * __TEMPLATE_PREP_RESULT__. The source PSD is never overwritten.
 */

var REQUIRED_TEXT_KEYS = ["卖点", "规格", "到手", "价格1", "价格2", "券名", "折扣", "券门槛", "活动时间"];
var CHANNEL_PROFILE = $.global.__TEMPLATE_PREP_INPUTS__ && $.global.__TEMPLATE_PREP_INPUTS__.profile;
var TEXT_NAME_MAP = {
    "卖点": "@卖点",
    "规格": "@规格",
    "到手": "@到手",
    "价格1": "@价格1",
    "价格2": "@价格2",
    "券名": "@券名",
    "折扣": "@折扣",
    "满129可用": "@券门槛",
    "券门槛": "@券门槛",
    "时间": "@活动时间",
    "活动时间": "@活动时间"
};
var PRODUCT_LAYER_NAMES = {
    "!商品图": true,
    "商品图": true,
    "堆图": true,
    "产品图": true,
    "产品堆图": true,
    "主产品": true
};

function trimText(value) {
    return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function addAllLayers(container, items) {
    for (var index = 0; index < container.layers.length; index++) {
        var layer = container.layers[index];
        items.push(layer);
        if (layer.typename === "LayerSet") {
            addAllLayers(layer, items);
        }
    }
}

function isTextLayer(layer) {
    return layer.typename === "ArtLayer" && layer.kind === LayerKind.TEXT;
}

function isSmartObject(layer) {
    return layer.typename === "ArtLayer" && layer.kind === LayerKind.SMARTOBJECT;
}

function findNamedLayers(document, name) {
    var all = [];
    var matches = [];
    addAllLayers(document, all);
    for (var index = 0; index < all.length; index++) {
        if (all[index].name === name) {
            matches.push(all[index]);
        }
    }
    return matches;
}

function usesPhotoshopVariables() {
    return !!(CHANNEL_PROFILE && CHANNEL_PROFILE.execution_mode === "photoshop_variables");
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

function templateProblems(document) {
    var issues = [];
    if (usesPhotoshopVariables()) {
        for (var variableIndex = 0; variableIndex < CHANNEL_PROFILE.required_psd_variables.length; variableIndex++) {
            var variableRequired = CHANNEL_PROFILE.required_psd_variables[variableIndex];
            var documentVariable = findDocumentVariable(document, variableRequired.name);
            if (!documentVariable) {
                issues.push("E_VAR_MISSING: " + variableRequired.name);
            } else if (documentVariable.kind !== expectedVariableKind(variableRequired.type)) {
                issues.push("E_VAR_TYPE_MISMATCH: " + variableRequired.name);
            }
        }
        return issues;
    }
    if (CHANNEL_PROFILE && CHANNEL_PROFILE.required_psd_variables) {
        for (var profileIndex = 0; profileIndex < CHANNEL_PROFILE.required_psd_variables.length; profileIndex++) {
            var required = CHANNEL_PROFILE.required_psd_variables[profileIndex];
            var expectedName = (required.type === "text" ? "@" : "!") + required.name;
            var bound = findNamedLayers(document, expectedName);
            if (bound.length === 0) {
                issues.push("E_VAR_UNBOUND: " + expectedName);
                continue;
            }
            for (var boundIndex = 0; boundIndex < bound.length; boundIndex++) {
                if ((required.type === "text" && !isTextLayer(bound[boundIndex])) ||
                    (required.type === "smart_object" && !isSmartObject(bound[boundIndex]))) {
                    issues.push("E_VAR_TYPE_MISMATCH: " + expectedName);
                    break;
                }
            }
        }
        return issues;
    }
    for (var index = 0; index < REQUIRED_TEXT_KEYS.length; index++) {
        var key = REQUIRED_TEXT_KEYS[index];
        var matches = findNamedLayers(document, "@" + key);
        if (matches.length === 0) {
            issues.push("缺少 @" + key + " 文本层");
        } else {
            for (var matchIndex = 0; matchIndex < matches.length; matchIndex++) {
                if (!isTextLayer(matches[matchIndex])) {
                    issues.push("@" + key + " 不是文本层");
                    break;
                }
            }
        }
    }
    var productMatches = findNamedLayers(document, "!商品图");
    if (productMatches.length === 0) {
        issues.push("缺少 !商品图 智能对象");
    } else {
        for (var productIndex = 0; productIndex < productMatches.length; productIndex++) {
            if (!isSmartObject(productMatches[productIndex])) {
                issues.push("!商品图 不是智能对象");
                break;
            }
        }
    }
    return issues;
}

function findProductCandidates(document) {
    var all = [];
    var candidates = [];
    addAllLayers(document, all);
    for (var index = 0; index < all.length; index++) {
        var layer = all[index];
        if (layer.typename === "ArtLayer" && PRODUCT_LAYER_NAMES[layer.name]) {
            candidates.push(layer);
        }
    }
    return candidates;
}

function hasLegacyChannelDesignSignals(document) {
    // Some channel teams deliver a finished design composition whose layers
    // are named by visual section (产品/时间/价格), not by the fields that the
    // batch engine can safely replace.  Treat it as an explicit mapping task
    // instead of guessing which of several text layers should receive data.
    var signals = ["产品", "时间", "价格"];
    var matches = 0;
    for (var index = 0; index < signals.length; index++) {
        if (findNamedLayers(document, signals[index]).length > 0) {
            matches++;
        }
    }
    return matches >= 2;
}

function inspectPreparation(document) {
    var existing = templateProblems(document);
    if (existing.length === 0) {
        return { status: "READY", message: usesPhotoshopVariables() ? "模板 PSD Variables 名称、类型和绑定关系已通过体检。" : "模板已符合 @文本、!智能对象命名规范。" };
    }

    if (usesPhotoshopVariables()) {
        return { status: "AMBIGUOUS", message: "PSD Variables 体检未通过：" + existing.join("；") + "。请模板制作人员恢复变量绑定后重试。" };
    }

    if (hasLegacyChannelDesignSignals(document)) {
        return {
            status: "AMBIGUOUS",
            message: "检测到渠道设计稿图层（产品/时间/价格），但未按工具字段命名。请模板制作人员先将动态文字改为 @字段、商品图改为 !商品图智能对象后再运行。"
        };
    }

    var candidates = findProductCandidates(document);
    if (findNamedLayers(document, "!商品图").length === 0) {
        if (candidates.length === 0) {
            return { status: "AMBIGUOUS", message: "未能识别商品图层。请将商品像素层命名为【商品图】、【堆图】或【产品图】后重试。" };
        }
        if (candidates.length > 1) {
            var names = [];
            for (var i = 0; i < candidates.length; i++) { names.push(candidates[i].name); }
            return { status: "AMBIGUOUS", message: "识别到多个可能的商品图层：" + names.join("、") + "。为避免改错模板，未自动处理。请保留一个商品图层候选后重试。" };
        }
    }

    var all = [];
    addAllLayers(document, all);
    for (var index = 0; index < all.length; index++) {
        var layer = all[index];
        if (TEXT_NAME_MAP[layer.name] && !isTextLayer(layer)) {
            return { status: "AMBIGUOUS", message: "图层【" + layer.name + "】不是文本层，不能自动改为 " + TEXT_NAME_MAP[layer.name] + "。" };
        }
    }
    return { status: "NEEDS_PREP", message: existing.join("；") };
}

function convertToSmartObject(document, layer) {
    document.activeLayer = layer;
    if (!isSmartObject(layer)) {
        executeAction(stringIDToTypeID("newPlacedLayer"), undefined, DialogModes.NO);
    }
    if (!isSmartObject(document.activeLayer)) {
        throw new Error("商品图层转换智能对象失败");
    }
    document.activeLayer.name = "!商品图";
}

function prepareTemplate(document) {
    var inspection = inspectPreparation(document);
    if (inspection.status === "READY") {
        return inspection;
    }
    if (inspection.status === "AMBIGUOUS") {
        return inspection;
    }

    var all = [];
    addAllLayers(document, all);
    for (var index = 0; index < all.length; index++) {
        var layer = all[index];
        if (TEXT_NAME_MAP[layer.name] && isTextLayer(layer)) {
            layer.name = TEXT_NAME_MAP[layer.name];
        }
    }

    var product = findNamedLayers(document, "!商品图");
    if (product.length === 1 && !isSmartObject(product[0])) {
        convertToSmartObject(document, product[0]);
    } else if (product.length === 0) {
        var candidates = findProductCandidates(document);
        if (candidates.length !== 1) {
            throw new Error("商品图层识别状态在改造时发生变化");
        }
        convertToSmartObject(document, candidates[0]);
    }

    // Preserve optional elements but give two common groups their standard switch names.
    var couponGroups = findNamedLayers(document, "优惠券");
    for (var couponIndex = 0; couponIndex < couponGroups.length; couponIndex++) {
        if (couponGroups[couponIndex].typename === "LayerSet") {
            couponGroups[couponIndex].name = "#优惠券开关";
        }
    }
    var packageGroups = findNamedLayers(document, "新旧包装").concat(findNamedLayers(document, "新旧"));
    for (var packageIndex = 0; packageIndex < packageGroups.length; packageIndex++) {
        if (packageGroups[packageIndex].typename === "LayerSet") {
            packageGroups[packageIndex].name = "#展示新旧包装";
        }
    }

    var problems = templateProblems(document);
    if (problems.length > 0) {
        throw new Error("自动改造后仍不完整：" + problems.join("；"));
    }
    return { status: "PREPARED", message: "已完成标准图层命名和商品智能对象转换。" };
}

function uniquePreparedFile(sourceFile) {
    var sourceName = sourceFile.name.replace(/\.psd$/i, "");
    var folder = sourceFile.parent;
    var candidate = File(folder.fsName + "/" + sourceName + "_套版模板.psd");
    var suffix = 1;
    while (candidate.exists) {
        candidate = File(folder.fsName + "/" + sourceName + "_套版模板_" + suffix + ".psd");
        suffix++;
    }
    return candidate;
}

function writePreparationReport(document, outputFile, message) {
    var report = File(outputFile.fsName.replace(/\.psd$/i, "_模板改造报告.txt"));
    var lines = [
        "PSD 模板自动改造报告",
        "模板副本：" + outputFile.fsName,
        "结果：" + message,
        "",
        "图层清单："
    ];
    var all = [];
    addAllLayers(document, all);
    for (var index = 0; index < all.length; index++) {
        var layer = all[index];
        var type = layer.typename === "LayerSet" ? "组" : (isTextLayer(layer) ? "文本" : (isSmartObject(layer) ? "智能对象" : "像素层"));
        lines.push("- " + layer.name + " [" + type + "]");
    }
    report.encoding = "UTF8";
    if (report.open("w")) {
        report.write("\uFEFF" + lines.join("\n"));
        report.close();
    }
}

function main() {
    if (app.documents.length === 0) {
        throw new Error("没有打开 PSD 模板");
    }
    var document = app.activeDocument;
    var inputs = $.global.__TEMPLATE_PREP_INPUTS__ || {};
    var mode = inputs.mode || "check";
    var inspection = inspectPreparation(document);
    if (mode !== "prepare" || inspection.status === "READY" || inspection.status === "AMBIGUOUS") {
        return inspection.status + "|" + inspection.message + "|" + document.fullName.fsName;
    }

    var result = prepareTemplate(document);
    var outputFile = uniquePreparedFile(File(document.fullName));
    var saveOptions = new PhotoshopSaveOptions();
    saveOptions.layers = true;
    saveOptions.embedColorProfile = true;
    document.saveAs(outputFile, saveOptions, true, Extension.LOWERCASE);
    writePreparationReport(document, outputFile, result.message);
    return result.status + "|" + result.message + "|" + outputFile.fsName;
}

var oldDialogs = app.displayDialogs;
var templatePrepFinalResult = "ERROR|未执行模板检测|";
try {
    app.displayDialogs = DialogModes.NO;
    templatePrepFinalResult = main();
} catch (error) {
    templatePrepFinalResult = "ERROR|" + error.message + "|";
} finally {
    app.displayDialogs = oldDialogs;
}
// Keep the explicit result as the last expression. Photoshop COM otherwise
// returns the DialogModes value assigned in the finally block.
templatePrepFinalResult;
