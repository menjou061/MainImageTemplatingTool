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

function buildLayerIndex(document) {
    var index = { all: [], named: {} };
    for (var layerIndex = 0; layerIndex < document.layers.length; layerIndex++) {
        indexLayer(document.layers[layerIndex], index);
    }
    return index;
}

function indexLayer(layer, index) {
    index.all.push(layer);
    if (!index.named[layer.name]) {
        index.named[layer.name] = [];
    }
    index.named[layer.name].push(layer);
    if (layer.typename === "LayerSet") {
        for (var childIndex = 0; childIndex < layer.layers.length; childIndex++) {
            indexLayer(layer.layers[childIndex], index);
        }
    }
}

function isLayerWithin(layer, container) {
    var current = layer;
    while (current && current.parent) {
        if (current.parent === container) {
            return true;
        }
        current = current.parent;
    }
    return false;
}

function isTextLayer(layer) {
    return layer.typename === "ArtLayer" && layer.kind === LayerKind.TEXT;
}

function isSmartObject(layer) {
    return layer.typename === "ArtLayer" && layer.kind === LayerKind.SMARTOBJECT;
}

function findNamedLayers(document, name, layerIndex) {
    if (layerIndex) {
        return layerIndex.named[name] || [];
    }
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

function findNamedLayersWithin(container, name, layerIndex) {
    if (layerIndex) {
        var indexed = layerIndex.named[name] || [];
        var scoped = [];
        for (var indexedLayer = 0; indexedLayer < indexed.length; indexedLayer++) {
            if (isLayerWithin(indexed[indexedLayer], container)) {
                scoped.push(indexed[indexedLayer]);
            }
        }
        return scoped;
    }
    var all = [];
    var matches = [];
    descendants(container, all);
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

function isHygieneRecordProfile() {
    return !!(CHANNEL_PROFILE && CHANNEL_PROFILE.profile_id === "hygiene-tmall-v1.2");
}

function descendants(container, output) {
    for (var index = 0; index < container.layers.length; index++) {
        var layer = container.layers[index];
        output.push(layer);
        if (layer.typename === "LayerSet") {
            descendants(layer, output);
        }
    }
}

function findFirstGroup(container, name, layerIndex) {
    if (layerIndex) {
        var indexed = layerIndex.named[name] || [];
        for (var indexedLayer = 0; indexedLayer < indexed.length; indexedLayer++) {
            if (indexed[indexedLayer].typename === "LayerSet" && isLayerWithin(indexed[indexedLayer], container)) {
                return indexed[indexedLayer];
            }
        }
        return null;
    }
    var all = [];
    descendants(container, all);
    for (var index = 0; index < all.length; index++) {
        if (all[index].typename === "LayerSet" && all[index].name === name) {
            return all[index];
        }
    }
    return null;
}

function findDirectGroup(container, name) {
    for (var index = 0; index < container.layers.length; index++) {
        var layer = container.layers[index];
        if (layer.typename === "LayerSet" && layer.name === name) {
            return layer;
        }
    }
    return null;
}

function findDirectGroupsByTrimmedName(container, name) {
    var matches = [];
    for (var index = 0; index < container.layers.length; index++) {
        var layer = container.layers[index];
        if (layer.typename === "LayerSet" && trimText(layer.name) === name) {
            matches.push(layer);
        }
    }
    return matches;
}

function textLayersWithin(container) {
    var all = [];
    var text = [];
    descendants(container, all);
    for (var index = 0; index < all.length; index++) {
        if (isTextLayer(all[index])) {
            text.push(all[index]);
        }
    }
    return text;
}

function artLayersWithin(container) {
    var all = [];
    var art = [];
    descendants(container, all);
    for (var index = 0; index < all.length; index++) {
        if (all[index].typename === "ArtLayer") {
            art.push(all[index]);
        }
    }
    return art;
}

function smartObjectLayersWithin(container) {
    var layers = artLayersWithin(container);
    var smartObjects = [];
    for (var index = 0; index < layers.length; index++) {
        if (isSmartObject(layers[index])) {
            smartObjects.push(layers[index]);
        }
    }
    return smartObjects;
}

function sortLayersByVisualPosition(layers) {
    layers.sort(function (left, right) {
        var leftTop = left.bounds[1].as("px");
        var rightTop = right.bounds[1].as("px");
        if (Math.abs(leftTop - rightTop) > 0.5) {
            return leftTop - rightTop;
        }
        return left.bounds[0].as("px") - right.bounds[0].as("px");
    });
    return layers;
}

function hygieneLayoutGroupNames() {
    var configured = CHANNEL_PROFILE && CHANNEL_PROFILE.record_layout && CHANNEL_PROFILE.record_layout.groups;
    var active = CHANNEL_PROFILE && CHANNEL_PROFILE.active_layout_groups;
    if (active && active.length > 0) {
        return active;
    }
    return configured || [];
}

function hygieneGiftSwitchNames() {
    var configured = CHANNEL_PROFILE && CHANNEL_PROFILE.record_layout && CHANNEL_PROFILE.record_layout.gift_switches;
    return configured && configured.length ? configured : ["赠品顶部", "赠品区域"];
}

function hygieneGiftSlotCount() {
    var configured = CHANNEL_PROFILE && CHANNEL_PROFILE.record_layout && Number(CHANNEL_PROFILE.record_layout.gift_slots);
    return configured > 0 ? configured : 3;
}

function inactiveHygieneLayouts(document, layerIndex) {
    var configured = CHANNEL_PROFILE && CHANNEL_PROFILE.record_layout && CHANNEL_PROFILE.record_layout.groups;
    var active = hygieneLayoutGroupNames();
    var activeNames = {};
    var inactive = [];
    if (!configured || configured.length === 0 || !active || active.length === 0) {
        return inactive;
    }
    for (var activeIndex = 0; activeIndex < active.length; activeIndex++) {
        activeNames[active[activeIndex]] = true;
    }
    for (var configuredIndex = 0; configuredIndex < configured.length; configuredIndex++) {
        var layoutName = configured[configuredIndex];
        if (activeNames[layoutName]) {
            continue;
        }
        var matches = findNamedLayers(document, layoutName, layerIndex);
        for (var matchIndex = 0; matchIndex < matches.length; matchIndex++) {
            if (matches[matchIndex].typename === "LayerSet") {
                inactive.push({ name: layoutName, layer: matches[matchIndex] });
            }
        }
    }
    return inactive;
}

function hygieneProblems(document, layerIndex) {
    var expected = [];
    var required = CHANNEL_PROFILE && CHANNEL_PROFILE.required_psd_variables || [];
    for (var requiredIndex = 0; requiredIndex < required.length; requiredIndex++) {
        expected.push((required[requiredIndex].type === "text" ? "@" : "!") + required[requiredIndex].name);
    }
    var configuredSwitches = hygieneGiftSwitchNames();
    var expectedGroups = [];
    for (var switchNameIndex = 0; switchNameIndex < configuredSwitches.length; switchNameIndex++) {
        expectedGroups.push("#" + configuredSwitches[switchNameIndex]);
    }
    var problems = [];
    var configured = hygieneLayoutGroupNames();
    for (var layoutIndex = 0; configured && layoutIndex < configured.length; layoutIndex++) {
        var layoutMatches = findNamedLayers(document, configured[layoutIndex], layerIndex);
        if (layoutMatches.length !== 1 || layoutMatches[0].typename !== "LayerSet") {
            problems.push("E_CONFIG_MISMATCH: 版式组 " + configured[layoutIndex] + " 缺失、重复或不是图层组");
            continue;
        }
        var layout = layoutMatches[0];
        for (var index = 0; index < expected.length; index++) {
            if (findNamedLayersWithin(layout, expected[index], layerIndex).length === 0) {
                problems.push("E_VAR_UNBOUND: " + configured[layoutIndex] + "/" + expected[index]);
            }
        }
        for (var groupIndex = 0; groupIndex < expectedGroups.length; groupIndex++) {
            if (findNamedLayersWithin(layout, expectedGroups[groupIndex], layerIndex).length === 0) {
                problems.push("E_GROUP_UNBOUND: " + configured[layoutIndex] + "/" + expectedGroups[groupIndex]);
            }
        }
    }
    if (!configured) {
        for (var expectedIndex = 0; expectedIndex < expected.length; expectedIndex++) {
            if (findNamedLayers(document, expected[expectedIndex], layerIndex).length === 0) {
                problems.push("E_VAR_UNBOUND: " + expected[expectedIndex]);
            }
        }
    }
    var structureProblems = hygieneStructureProblems(document, layerIndex);
    for (var structureIndex = 0; structureIndex < structureProblems.length; structureIndex++) {
        problems.push(structureProblems[structureIndex]);
    }
    return problems;
}

function hygieneStructureProblems(document, layerIndex) {
    var problems = [];
    var groups = hygieneLayoutGroupNames();
    if (!groups) { return problems; }
    for (var layoutIndex = 0; layoutIndex < groups.length; layoutIndex++) {
        var layouts = findNamedLayers(document, groups[layoutIndex], layerIndex);
        if (layouts.length !== 1 || layouts[0].typename !== "LayerSet") { continue; }
        var layout = layouts[0];
        if (hygieneGiftSlotCount() === 2) {
            var switchNames = hygieneGiftSwitchNames();
            for (var switchIndex = 0; switchIndex < switchNames.length; switchIndex++) {
                var card = findFirstGroup(layout, "#" + switchNames[switchIndex], layerIndex);
                var slotName = "!赠品图" + (switchIndex + 1);
                // A business card can retain decorative smart objects. Only
                // the explicitly bound material field controls replacement.
                if (card && findNamedLayersWithin(card, slotName, layerIndex).length !== 1) {
                    problems.push("E_GIFT_SLOT_MULTIPLE: " + groups[layoutIndex] + "/#" + switchNames[switchIndex]);
                }
            }
            continue;
        }
        var topGift = findDirectGroup(layout, "#赠品顶部");
        if (topGift && smartObjectLayersWithin(topGift).length !== 1) {
            problems.push("E_GIFT_SLOT_MULTIPLE: " + groups[layoutIndex] + "/#赠品顶部");
        }
        var bottom = findFirstGroup(layout, "下帖", layerIndex);
        var bottomGift = bottom ? findDirectGroup(bottom, "#赠品区域") : null;
        var assets = bottomGift ? findDirectGroup(bottomGift, "组 382") : null;
        if (!assets) { continue; }
        var slotGroups = [];
        for (var assetIndex = 0; assetIndex < assets.layers.length; assetIndex++) {
            if (assets.layers[assetIndex].typename === "LayerSet") {
                slotGroups.push(assets.layers[assetIndex]);
            }
        }
        for (var slotIndex = 0; slotIndex < 2 && slotIndex < slotGroups.length; slotIndex++) {
            var slotObjects = smartObjectLayersWithin(slotGroups[slotIndex]);
            if (slotObjects.length !== 1) {
                problems.push("E_GIFT_SLOT_MULTIPLE: " + groups[layoutIndex] + "/赠品图" + (slotIndex + 2));
            }
        }
    }
    return problems;
}

function hygieneSourceHasMultipleLayouts(document) {
    // Business PSDs commonly keep one top-level group per SKU/layout. Do not
    // attempt automatic mapping on that source; only a fully standardized
    // sibling copy may pass the normal READY path above.
    var topLevelGroups = 0;
    for (var index = 0; index < document.layers.length; index++) {
        if (document.layers[index].typename === "LayerSet") {
            topLevelGroups++;
        }
    }
    return topLevelGroups > 1;
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

function profileRequiredVariable(name) {
    if (!CHANNEL_PROFILE || !CHANNEL_PROFILE.required_psd_variables) { return null; }
    for (var index = 0; index < CHANNEL_PROFILE.required_psd_variables.length; index++) {
        if (CHANNEL_PROFILE.required_psd_variables[index].name === name) {
            return CHANNEL_PROFILE.required_psd_variables[index];
        }
    }
    return null;
}

function profileBindingsAreUsable(document) {
    if (!CHANNEL_PROFILE || !CHANNEL_PROFILE.template_bindings) { return false; }
    for (var target in CHANNEL_PROFILE.template_bindings) {
        if (!CHANNEL_PROFILE.template_bindings.hasOwnProperty(target)) { continue; }
        var sourceLayers = findNamedLayers(document, CHANNEL_PROFILE.template_bindings[target]);
        var required = profileRequiredVariable(target);
        if (!required || sourceLayers.length !== 1) { return false; }
        if (required.type === "text" && !isTextLayer(sourceLayers[0])) { return false; }
        if (required.type === "smart_object" && sourceLayers[0].typename !== "ArtLayer") { return false; }
    }
    return true;
}

function applyProfileBindings(document) {
    if (!CHANNEL_PROFILE || !CHANNEL_PROFILE.template_bindings) { return; }
    for (var target in CHANNEL_PROFILE.template_bindings) {
        if (!CHANNEL_PROFILE.template_bindings.hasOwnProperty(target)) { continue; }
        var sourceLayers = findNamedLayers(document, CHANNEL_PROFILE.template_bindings[target]);
        var required = profileRequiredVariable(target);
        if (!required || sourceLayers.length !== 1) {
            throw new Error("模板图层映射失效：" + target);
        }
        var layer = sourceLayers[0];
        if (required.type === "text") {
            if (!isTextLayer(layer)) { throw new Error("映射图层不是文本层：" + target); }
            layer.name = "@" + target;
        } else {
            convertToSmartObject(document, layer, target);
        }
    }
}

function inspectPreparation(document, layerIndex) {
    if (isHygieneRecordProfile()) {
        var scopedIndex = layerIndex || buildLayerIndex(document);
        var hygieneIssues = hygieneProblems(document, scopedIndex);
        var inactiveLayouts = inactiveHygieneLayouts(document, scopedIndex);
        if (hygieneIssues.length === 0) {
            var layoutMessage = inactiveLayouts.length > 0 ?
                "模板包含 " + inactiveLayouts.length + " 个未使用版式，本次只读取选中版式。" :
                "模板只包含本次任务版式。";
            return { status: "READY", message: "卫品天猫官旗 750 字段已通过体检；" + layoutMessage };
        }
        if (inactiveLayouts.length > 0 || hygieneSourceHasMultipleLayouts(document)) {
            return {
                status: "BLOCKED_PREP_REQUIRED",
                message: "E_TEMPLATE_PREP_REQUIRED：模板包含多套版式且字段未标准化。请由业务方先生成并确认标准副本（文件名建议为原文件名_套版模板.psd），设计师不要在跑批时自动改造原始 PSD。"
            };
        }
        return {
            status: "NEEDS_PREP",
            message: "已识别单版式卫品天猫官旗 750 模板，仅会在副本中建立本次任务字段映射。"
        };
    }
    var existing = templateProblems(document);
    if (existing.length === 0) {
        return { status: "READY", message: usesPhotoshopVariables() ? "模板 PSD Variables 名称、类型和绑定关系已通过体检。" : "模板已符合 @文本、!智能对象命名规范。" };
    }

    if (usesPhotoshopVariables()) {
        return { status: "AMBIGUOUS", message: "PSD Variables 体检未通过：" + existing.join("；") + "。请模板制作人员恢复变量绑定后重试。" };
    }

    if (profileBindingsAreUsable(document)) {
        return { status: "NEEDS_PREP", message: "已识别当前渠道模板图层映射，将生成套版模板副本。" };
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

function convertToSmartObject(document, layer, targetName) {
    document.activeLayer = layer;
    if (!isSmartObject(layer)) {
        executeAction(stringIDToTypeID("newPlacedLayer"), undefined, DialogModes.NO);
    }
    if (!isSmartObject(document.activeLayer)) {
        throw new Error("商品图层转换智能对象失败");
    }
    document.activeLayer.name = "!" + (targetName || "商品图");
}

function renameFirstText(layers, targetName) {
    if (layers.length === 0) { return false; }
    layers[0].name = "@" + targetName;
    return true;
}

function textLayerHeight(layer) {
    try {
        return layer.bounds[3].as("px") - layer.bounds[1].as("px");
    } catch (error) {
        return 0;
    }
}

function renameLargeSmallText(layers, largeName, smallName) {
    if (layers.length === 0) { return false; }
    if (layers.length === 1) {
        layers[0].name = "@" + largeName;
        return true;
    }
    var large = layers[0];
    var small = layers[1];
    if (textLayerHeight(small) > textLayerHeight(large)) {
        large = layers[1];
        small = layers[0];
    }
    large.name = "@" + largeName;
    small.name = "@" + smallName;
    return true;
}

function renameGiftImage(document, layer, targetName) {
    if (!layer) { return false; }
    convertToSmartObject(document, layer, targetName);
    return true;
}

function normalizeGiftSlot(document, slotGroup, targetName) {
    if (!slotGroup) {
        throw new Error("E_CONFIG_MISMATCH: 缺少 " + targetName + " 赠品槽位");
    }
    var sourceObjects = smartObjectLayersWithin(slotGroup);
    if (sourceObjects.length === 0) {
        throw new Error("E_CONFIG_MISMATCH: " + targetName + " 槽位没有可替换素材");
    }
    // A supplied gift path represents one finished gift composition. Some
    // legacy PSDs build one card from several original product objects; turn
    // that complete card into one smart object before replacement so the old
    // objects cannot remain underneath or receive the same path repeatedly.
    convertToSmartObject(document, slotGroup, targetName);
}

function prepareHygieneTwoGiftCards(document, layoutGroup) {
    var cards = findDirectGroupsByTrimmedName(layoutGroup, "赠品");
    sortLayersByVisualPosition(cards);
    if (cards.length !== 2) {
        throw new Error("E_CONFIG_MISMATCH: 800 模板需要 2 个独立顶部赠品卡片，当前有 " + cards.length + " 个");
    }
    for (var cardIndex = 0; cardIndex < cards.length; cardIndex++) {
        var card = cards[cardIndex];
        var copy = [];
        var text = textLayersWithin(card);
        for (var textIndex = 0; textIndex < text.length; textIndex++) {
            var textName = trimText(text[textIndex].name);
            if (textName !== "买2赠" && textName !== "买就送") {
                copy.push(text[textIndex]);
            }
        }
        if (copy.length !== 1) {
            throw new Error("E_CONFIG_MISMATCH: 800 赠品卡片 " + (cardIndex + 1) + " 需要 1 个动态文案图层，当前有 " + copy.length + " 个");
        }
        renameFirstText(copy, "赠品文案" + (cardIndex + 1));
        var assetGroups = [];
        for (var childIndex = 0; childIndex < card.layers.length; childIndex++) {
            if (card.layers[childIndex].typename === "LayerSet") {
                assetGroups.push(card.layers[childIndex]);
            }
        }
        if (assetGroups.length !== 1) {
            throw new Error("E_CONFIG_MISMATCH: 800 赠品卡片 " + (cardIndex + 1) + " 需要 1 个可替换素材组，当前有 " + assetGroups.length + " 个");
        }
        normalizeGiftSlot(document, assetGroups[0], "赠品图" + (cardIndex + 1));
        card.name = "#赠品槽位" + (cardIndex + 1);
    }
}

function prepareHygieneTopAndBottomGiftCards(document, layoutGroup, layoutIndex) {
    var topCards = findDirectGroupsByTrimmedName(layoutGroup, "赠品");
    if (topCards.length !== 1) {
        throw new Error("E_CONFIG_MISMATCH: 750 模板需要 1 个顶部赠品卡片，当前有 " + topCards.length + " 个");
    }
    var topCard = topCards[0];
    var topCopy = [];
    var topText = textLayersWithin(topCard);
    for (var topTextIndex = 0; topTextIndex < topText.length; topTextIndex++) {
        if (trimText(topText[topTextIndex].name) !== "买就送") {
            topCopy.push(topText[topTextIndex]);
        }
    }
    if (topCopy.length !== 1) {
        throw new Error("E_CONFIG_MISMATCH: 750 顶部赠品需要 1 个动态文案图层，当前有 " + topCopy.length + " 个");
    }
    renameFirstText(topCopy, "赠品文案1");
    // The approved 750 PSD has decoration groups alongside the actual asset.
    // Bind its named material group rather than treating every nested group as
    // a candidate. This preserves the card frame and prevents stale material.
    var topAsset = findDirectGroup(topCard, "组 381");
    if (!topAsset) {
        throw new Error("E_CONFIG_MISMATCH: 750 顶部赠品缺少素材组“组 381”");
    }
    normalizeGiftSlot(document, topAsset, "赠品图1");
    topCard.name = "#赠品顶部";

    var bottom = findFirstGroup(layoutGroup, "下帖", layoutIndex);
    var bottomCard = bottom ? findDirectGroup(bottom, "赠品") : null;
    if (!bottomCard) {
        throw new Error("E_CONFIG_MISMATCH: 750 模板缺少下方组合赠品卡片");
    }
    var bottomCopy = [];
    var bottomText = textLayersWithin(bottomCard);
    for (var bottomTextIndex = 0; bottomTextIndex < bottomText.length; bottomTextIndex++) {
        var bottomName = trimText(bottomText[bottomTextIndex].name);
        if (bottomName !== "买2赠" && bottomName.charAt(0) !== "*") {
            bottomCopy.push(bottomText[bottomTextIndex]);
        }
    }
    if (bottomCopy.length !== 1) {
        throw new Error("E_CONFIG_MISMATCH: 750 下方组合赠品需要 1 个动态文案图层，当前有 " + bottomCopy.length + " 个");
    }
    renameFirstText(bottomCopy, "赠品文案2");
    var bottomAsset = findDirectGroup(bottomCard, "3QFC8202+QFC8802");
    if (!bottomAsset) {
        throw new Error("E_CONFIG_MISMATCH: 750 下方组合赠品缺少组合素材组“3QFC8202+QFC8802”");
    }
    normalizeGiftSlot(document, bottomAsset, "赠品图2");
    bottomCard.name = "#赠品区域";
}

function prepareHygieneLayoutGroup(document, layoutGroup) {
    var layoutIndex = buildLayerIndex(layoutGroup);
    var product = findFirstGroup(layoutGroup, "产品", layoutIndex);
    var productFrame = product ? findFirstGroup(product, "组 7", layoutIndex) : null;
    var productLayers = productFrame ? artLayersWithin(productFrame) : [];
    var productCandidates = [];
    for (var productIndex = 0; productIndex < productLayers.length; productIndex++) {
        var productLayer = productLayers[productIndex];
        var productName = trimText(productLayer.name);
        if (!isSmartObject(productLayer) || productName.charAt(0) === "#") {
            continue;
        }
        if (PRODUCT_LAYER_NAMES[productName] || /商品|产品|堆品|堆图/.test(productName)) {
            productCandidates.push(productLayer);
        }
    }
    if (productCandidates.length === 0) {
        var smartProducts = smartObjectLayersWithin(productFrame || layoutGroup);
        if (smartProducts.length === 1) {
            productCandidates.push(smartProducts[0]);
        }
    }
    if (productCandidates.length !== 1) {
        throw new Error("E_CONFIG_MISMATCH: 产品/组 7 必须只有一个商品图智能对象，未按默认图层猜测");
    }
    convertToSmartObject(document, productCandidates[0], "商品图");

    var title = findFirstGroup(layoutGroup, "标题", layoutIndex);
    var titleTexts = title ? textLayersWithin(title) : [];
    var mainTitle = [];
    var note = [];
    for (var titleIndex = 0; titleIndex < titleTexts.length; titleIndex++) {
        var titleLayer = titleTexts[titleIndex];
        if (titleLayer.name.charAt(0) === "#") { continue; }
        if (titleLayer.name.charAt(0) === "*") {
            note.push(titleLayer);
        } else {
            mainTitle.push(titleLayer);
        }
    }
    renameFirstText(mainTitle, "卖点");
    renameFirstText(note, "备注");

    var pieces = findFirstGroup(layoutGroup, "片数", layoutIndex);
    var pieceText = pieces ? textLayersWithin(pieces) : [];
    renameLargeSmallText(pieceText, "片数数量", "片数套");

    var bottom = findFirstGroup(layoutGroup, "下帖", layoutIndex);
    var formula = bottom ? findFirstGroup(bottom, "公式", layoutIndex) : null;
    var formulaKeys = ["价格活动价", "价格优惠券", "价格立减"];
    var formulaSubKeys = ["价格活动价副标", "价格优惠券副标", null];
    for (var formulaIndex = 0; formulaIndex < 3; formulaIndex++) {
        var formulaPart = formula ? findDirectGroup(formula, String(formulaIndex + 1)) : null;
        var formulaText = formulaPart ? textLayersWithin(formulaPart) : [];
        if (formulaSubKeys[formulaIndex]) {
            renameLargeSmallText(formulaText, formulaKeys[formulaIndex], formulaSubKeys[formulaIndex]);
        } else {
            renameFirstText(formulaText, formulaKeys[formulaIndex]);
        }
    }
    var priceBadge = bottom ? findFirstGroup(bottom, "组 359", layoutIndex) : null;
    var badgeText = priceBadge ? textLayersWithin(priceBadge) : [];
    renameLargeSmallText(badgeText, "到手", "到手标签");

    if (hygieneGiftSlotCount() === 2) {
        var giftLayout = CHANNEL_PROFILE && CHANNEL_PROFILE.record_layout && CHANNEL_PROFILE.record_layout.gift_layout;
        if (giftLayout === "two_top_cards") {
            prepareHygieneTwoGiftCards(document, layoutGroup);
        } else {
            prepareHygieneTopAndBottomGiftCards(document, layoutGroup, layoutIndex);
        }
        var ambassador800 = findFirstGroup(layoutGroup, "代言人", layoutIndex);
        if (ambassador800) {
            var ambassadorLayers800 = artLayersWithin(ambassador800);
            if (ambassadorLayers800.length > 0) {
                renameGiftImage(document, ambassadorLayers800[0], "代言IP");
                ambassador800.name = "#代言IP";
            }
        }
        return;
    }

    var topGift = findDirectGroup(layoutGroup, "赠品");
    var topGiftLayers = topGift ? smartObjectLayersWithin(topGift) : [];
    if (topGiftLayers.length !== 1) {
        throw new Error("E_CONFIG_MISMATCH: 顶部赠品槽位必须只有一个智能对象，当前有 " + topGiftLayers.length + " 个");
    }
    renameGiftImage(document, topGiftLayers[0], "赠品图1");
    var topGiftText = textLayersWithin(topGift);
    var topGiftCopy = [];
    for (var topTextIndex = 0; topTextIndex < topGiftText.length; topTextIndex++) {
        if (trimText(topGiftText[topTextIndex].name) !== "买就送") {
            topGiftCopy.push(topGiftText[topTextIndex]);
        }
    }
    if (topGiftCopy.length !== 1) {
        throw new Error("E_CONFIG_MISMATCH: 顶部赠品需要一个独立动态文案图层，当前有 " + topGiftCopy.length + " 个");
    }
    renameFirstText(topGiftCopy, "赠品文案1");
    // Photoshop 2026 drops a nested !赠品图1 name when its parent has
    // the same #赠品图1 switch name. Keep the switch semantic while
    // using a distinct internal group name so the binding survives save.
    topGift.name = "#赠品顶部";
    var bottomGift = bottom ? findDirectGroup(bottom, "赠品") : null;
    if (bottomGift) {
        var giftCopy = findFirstGroup(bottomGift, "文案 拷贝", layoutIndex);
        var giftText = giftCopy ? textLayersWithin(giftCopy) : [];
        var dynamicGiftText = [];
        for (var giftTextIndex = 0; giftTextIndex < giftText.length; giftTextIndex++) {
            var candidateText = giftText[giftTextIndex];
            var candidateName = trimText(candidateText.name);
            // The footer disclaimer and the fixed button are not per-gift
            // fields. Binding either one would overwrite template copy.
            if (candidateName === "拍即赠" || candidateName.charAt(0) === "*") {
                continue;
            }
            if (textLayerHeight(candidateText) >= 14) {
                dynamicGiftText.push(candidateText);
            }
        }
        sortLayersByVisualPosition(dynamicGiftText);
        if (dynamicGiftText.length !== 2) {
            throw new Error("E_CONFIG_MISMATCH: 下方赠品区域需要 2 个独立动态文案图层（对应赠品文案2-3），当前有 " + dynamicGiftText.length + " 个");
        }
        for (var copyIndex = 0; copyIndex < dynamicGiftText.length; copyIndex++) {
            dynamicGiftText[copyIndex].name = "@赠品文案" + (copyIndex + 2);
        }
        var giftAssets = findFirstGroup(bottomGift, "组 382", layoutIndex);
        var giftGroups = giftAssets ? [] : null;
        if (giftAssets) {
            for (var groupIndex = 0; groupIndex < giftAssets.layers.length; groupIndex++) {
                if (giftAssets.layers[groupIndex].typename === "LayerSet") {
                    giftGroups.push(giftAssets.layers[groupIndex]);
                }
            }
            sortLayersByVisualPosition(giftGroups);
            if (giftGroups.length > 0) {
                normalizeGiftSlot(document, giftGroups[0], "赠品图2");
            }
            if (giftGroups.length > 1) {
                normalizeGiftSlot(document, giftGroups[1], "赠品图3");
            }
        }
        // Keep the optional switch key separate from the !赠品图2 image
        // field; otherwise setSwitches would overwrite the material path
        // with the switch value "是" before replacement.
        bottomGift.name = "#赠品区域";
    }

    var ambassador = findFirstGroup(layoutGroup, "代言人", layoutIndex);
    if (ambassador) {
        var ambassadorLayers = artLayersWithin(ambassador);
        if (ambassadorLayers.length > 0) {
            renameGiftImage(document, ambassadorLayers[0], "代言IP");
            ambassador.name = "#代言IP";
        }
    }
}

function prepareHygieneTemplate(document, layerIndex) {
    var groups = hygieneLayoutGroupNames();
    if (!groups) { throw new Error("卫品模板缺少版式组配置"); }
    for (var index = 0; index < groups.length; index++) {
        var matches = findNamedLayers(document, groups[index], layerIndex);
        if (matches.length !== 1 || matches[0].typename !== "LayerSet") {
            throw new Error("卫品模板版式组映射失效：" + groups[index]);
        }
        prepareHygieneLayoutGroup(document, matches[0]);
    }
}

function prepareTemplate(document, inspection, layerIndex) {
    if (inspection.status === "READY") {
        return inspection;
    }
    if (inspection.status === "AMBIGUOUS") {
        return inspection;
    }
    if (inspection.status === "BLOCKED_PREP_REQUIRED") {
        return inspection;
    }

    if (isHygieneRecordProfile()) {
        prepareHygieneTemplate(document, layerIndex || buildLayerIndex(document));
        var hygieneIssues = hygieneProblems(document, buildLayerIndex(document));
        if (hygieneIssues.length > 0) {
            throw new Error("卫品模板自动映射后仍不完整：" + hygieneIssues.join("；"));
        }
        return { status: "PREPARED", message: "已为卫品天猫官旗 750 单版式模板建立字段映射副本。" };
    }

    var all = [];
    addAllLayers(document, all);
    applyProfileBindings(document);
    if (CHANNEL_PROFILE && CHANNEL_PROFILE.template_bindings) {
        var profileProblems = templateProblems(document);
        if (profileProblems.length > 0) {
            throw new Error("自动改造后仍不完整：" + profileProblems.join("；"));
        }
        return { status: "PREPARED", message: "已按当前渠道映射完成图层命名和商品智能对象转换。" };
    }
    all = [];
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
        "结果：" + message
    ];
    if (isHygieneRecordProfile()) {
        lines.push("本次映射版式：" + hygieneLayoutGroupNames().join("、"));
        lines.push("说明：仅处理本次任务版式的字段图层；多版式原始稿须先由业务方生成标准副本，原始 PSD 未修改。");
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
    var layerIndex = buildLayerIndex(document);
    var inspection = inspectPreparation(document, layerIndex);
    if (mode !== "prepare" || inspection.status === "READY" || inspection.status === "AMBIGUOUS") {
        return inspection.status + "|" + inspection.message + "|" + document.fullName.fsName;
    }

    var result = prepareTemplate(document, inspection, layerIndex);
    if (result.status !== "PREPARED") {
        return result.status + "|" + result.message + "|" + document.fullName.fsName;
    }
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
