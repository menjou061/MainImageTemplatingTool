#target photoshop

/*
 * One-time converter for the approved kitchen-paper pilot template.
 * The source PSD is never overwritten.
 */

var SOURCE_PSD = "/Users/liuyizhen/Desktop/需求内容/需求2-同规格电商主图套版/纸品案例/2现有素材/主图背书厨房.psd";
var OUTPUT_PSD = "/Users/liuyizhen/Desktop/需求内容/需求2-同规格电商主图套版/纸品案例/2现有素材/主图背书厨房_模板v2.psd";
var LAYER_REPORT = "/Users/liuyizhen/Documents/工作/电商主图套版工具/主图背书厨房_模板v2_图层清单.txt";

function findDirectLayer(container, name) {
    for (var index = 0; index < container.layers.length; index++) {
        if (container.layers[index].name === name) {
            return container.layers[index];
        }
    }
    return null;
}

function requireLayer(container, name, pathLabel) {
    var layer = findDirectLayer(container, name);
    if (!layer) {
        throw new Error("找不到图层：" + pathLabel);
    }
    return layer;
}

function requirePath(document, names) {
    var container = document;
    var layer = null;
    var pathParts = [];
    for (var index = 0; index < names.length; index++) {
        pathParts.push(names[index]);
        layer = requireLayer(container, names[index], pathParts.join("/"));
        if (index < names.length - 1) {
            if (layer.typename !== "LayerSet") {
                throw new Error("目标不是图层组：" + pathParts.join("/"));
            }
            container = layer;
        }
    }
    return layer;
}

function renamePath(document, names, newName) {
    var layer = requirePath(document, names);
    layer.name = newName;
    return layer;
}

function convertPathToSmartObject(document, names, newName) {
    var targetLayer = requirePath(document, names);
    if (targetLayer.typename !== "ArtLayer") {
        throw new Error(names.join("/") + " 不是可转换的普通图层。");
    }
    document.activeLayer = targetLayer;
    if (targetLayer.kind !== LayerKind.SMARTOBJECT) {
        executeAction(stringIDToTypeID("newPlacedLayer"), undefined, DialogModes.NO);
    }
    document.activeLayer.name = newName;
    if (document.activeLayer.kind !== LayerKind.SMARTOBJECT) {
        throw new Error(names.join("/") + " 转换智能对象失败。");
    }
}

function layerType(layer) {
    if (layer.typename === "LayerSet") {
        return "group";
    }
    if (layer.kind === LayerKind.TEXT) {
        return "text";
    }
    if (layer.kind === LayerKind.SMARTOBJECT) {
        return "smartobject";
    }
    return "artlayer";
}

function appendLayerTree(container, depth, lines) {
    for (var index = 0; index < container.layers.length; index++) {
        var layer = container.layers[index];
        var indent = "";
        for (var pad = 0; pad < depth; pad++) {
            indent += "  ";
        }
        var line = indent + "- " + layer.name + " [" + layerType(layer) + "]";
        if (!layer.visible) {
            line += " [隐藏]";
        }
        if (layer.typename === "ArtLayer" && layer.kind === LayerKind.TEXT) {
            line += " 文本=\"" + String(layer.textItem.contents).replace(/[\r\n]+/g, " ") + "\"";
        }
        lines.push(line);
        if (layer.typename === "LayerSet") {
            appendLayerTree(layer, depth + 1, lines);
        }
    }
}

function writeLayerReport(document) {
    var lines = [
        "主图背书厨房_模板v2.psd 图层清单",
        "尺寸：" + Math.round(document.width.as("px")) + "x" + Math.round(document.height.as("px")),
        ""
    ];
    appendLayerTree(document, 0, lines);
    var reportFile = File(LAYER_REPORT);
    reportFile.encoding = "UTF8";
    if (!reportFile.open("w")) {
        throw new Error("无法写入图层清单：" + LAYER_REPORT);
    }
    reportFile.write("\uFEFF" + lines.join("\n"));
    reportFile.close();
}

function main() {
    var sourceFile = File(SOURCE_PSD);
    var outputFile = File(OUTPUT_PSD);
    if (!sourceFile.exists) {
        throw new Error("找不到源模板：" + SOURCE_PSD);
    }
    if (outputFile.exists) {
        throw new Error("目标模板已存在，为避免覆盖已停止：" + OUTPUT_PSD);
    }

    var document = app.open(sourceFile);
    app.activeDocument = document;
    convertPathToSmartObject(document, ["产品", "堆图"], "!商品图");
    convertPathToSmartObject(document, ["新旧", "DT17090-24旧"], "!旧包装图");
    convertPathToSmartObject(document, ["新旧", "正式618新旧包装底"], "!新旧包装底图");
    convertPathToSmartObject(document, ["组 32", "大尺寸"], "!大尺寸图");

    renamePath(document, ["组 7 拷贝", "卖点"], "@卖点");
    renamePath(document, ["组 7 拷贝", "规格"], "@规格");
    renamePath(document, ["组 7 拷贝", "到手"], "@到手");
    renamePath(document, ["组 7 拷贝", "价格1"], "@价格1");
    renamePath(document, ["组 7 拷贝", "价格2"], "@价格2");
    renamePath(document, ["优惠券", "券名"], "@券名");
    renamePath(document, ["优惠券", "折扣"], "@折扣");
    renamePath(document, ["优惠券", "满129可用"], "@券门槛");
    renamePath(document, ["时间"], "@活动时间");

    // The switch name intentionally matches clean_data.py's derived CSV column.
    renamePath(document, ["优惠券"], "#优惠券开关");
    renamePath(document, ["新旧"], "#展示新旧包装");
    renamePath(document, ["组 32"], "#展示大尺寸");

    var saveOptions = new PhotoshopSaveOptions();
    saveOptions.layers = true;
    saveOptions.embedColorProfile = true;
    document.saveAs(outputFile, saveOptions, true, Extension.LOWERCASE);
    writeLayerReport(document);
    document.close(SaveOptions.DONOTSAVECHANGES);
}

var previousDialogs = app.displayDialogs;
try {
    app.displayDialogs = DialogModes.NO;
    main();
} catch (error) {
    throw new Error("模板改造失败：" + error.message);
} finally {
    app.displayDialogs = previousDialogs;
}
