const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const repositoryRoot = path.resolve(__dirname, "..");
const scriptPath = path.join(
    repositoryRoot,
    "L0_Windows命令行版",
    "_internal",
    "batch_template.jsx"
);
const source = fs.readFileSync(scriptPath, "utf8");
const start = source.indexOf("var REPORT_NAME");
const end = source.indexOf("function keyCount");
assert.notEqual(start, -1, "batch_template.jsx test prelude was not found");
assert.notEqual(end, -1, "batch_template.jsx record-layout indexing boundary was not found");

const profile = {
    layout: "record_rows",
    record_layout: { groups: ["版式甲", "版式乙 "] }
};
const sandbox = {
    $: { global: { __BATCH_INPUTS__: { profile } } }
};
vm.createContext(sandbox);
vm.runInContext(source.slice(start, end), sandbox, { filename: scriptPath });

function layer(name) {
    return { name, typename: "ArtLayer", visible: true };
}

function group(name, layers = []) {
    return { name, typename: "LayerSet", visible: true, layers };
}

const first = group("版式甲");
const selected = group("版式乙 ");
const document = {
    layers: [group("外层", [first, group("更深一层", [selected])])]
};

sandbox.selectRecordLayout(document, { "版式组": "版式乙" });
assert.equal(first.visible, false, "unselected configured layouts must be hidden");
assert.equal(selected.visible, true, "the recursively resolved layout must be visible");
assert.equal(sandbox.ACTIVE_LAYOUT_GROUP, selected);

const publicFloat = layer("!赠品浮窗");
const inactiveFloat = layer("!赠品浮窗");
const activeGift = layer("!赠品图1");
const publicSwitch = group("#价格优惠券");
const publicDocument = {
    layers: [publicFloat, publicSwitch, group("版式甲", [inactiveFloat]), selected]
};
selected.layers = [activeGift];
sandbox.selectRecordLayout(publicDocument, { "版式组": "版式乙" });
const dynamicIndex = sandbox.buildRecordLayerIndex(publicDocument);
assert.equal(dynamicIndex.image["赠品浮窗"].length, 1, "document-level smart objects must be dynamic for the selected layout");
assert.equal(dynamicIndex.image["赠品浮窗"][0], publicFloat, "inactive layout smart objects must not be updated");
assert.equal(dynamicIndex.image["赠品图1"][0], activeGift, "selected layout smart objects must stay dynamic");
assert.equal(dynamicIndex.switches["价格优惠券"][0], publicSwitch, "document-level switch groups must remain available");

const switchEnd = source.indexOf("function layerWidth");
assert.notEqual(switchEnd, -1, "batch_template.jsx switch boundary was not found");
const switchSandbox = { $: { global: { __BATCH_INPUTS__: { profile: {} } } } };
vm.createContext(switchSandbox);
vm.runInContext(source.slice(start, switchEnd), switchSandbox, { filename: scriptPath });
assert.equal(switchSandbox.hasBindingForKey(dynamicIndex, "价格优惠券开关"), true, "legacy switch columns must match their canonical # field name");
const genericSwitch = group("#浮窗区域", [layer("@浮窗文案"), layer("!赠品浮窗")]);
const genericIndex = { text: {}, image: {}, switches: { "浮窗区域": [genericSwitch] } };
const genericRecord = { "浮窗文案": "会员0元试用", "赠品浮窗": "伞2.png" };
switchSandbox.setSwitches(genericIndex, genericRecord, { issues: [] });
assert.equal(genericSwitch.visible, true, "a switch must show when any child dynamic field has a value");
assert.equal(genericRecord["浮窗区域"], undefined, "derived switch state must not alter the table record");
assert.equal(switchSandbox.switchStateValue(genericRecord, "浮窗区域"), "是", "derived switch state must be retained outside the table record");
const emptyGenericRecord = { "浮窗文案": "", "赠品浮窗": "" };
switchSandbox.SWITCH_STATE = {};
switchSandbox.setSwitches(genericIndex, emptyGenericRecord, { issues: [] });
assert.equal(genericSwitch.visible, false, "a switch must hide when all child dynamic fields are empty");
const priceSwitch = group("#价格优惠券", [layer("@价格优惠券")]);
const priceRecord = { "价格优惠券": "2" };
switchSandbox.SWITCH_STATE = {};
switchSandbox.setSwitches({ text: {}, image: {}, switches: { "价格优惠券": [priceSwitch] } }, priceRecord, { issues: [] });
assert.equal(priceSwitch.visible, true, "a price switch must follow the populated child text layer");
assert.equal(priceRecord["价格优惠券"], "2", "a derived switch must never overwrite same-name text data");

assert.throws(
    () => sandbox.selectRecordLayout({ layers: [first] }, { "版式组": "版式乙" }),
    /E_CONFIG_MISMATCH: 模板没有版式组/
);

const duplicate = {
    layers: [group("外层甲", [group("版式乙 ")]), group("外层乙", [group("版式乙 ")])]
};
assert.throws(
    () => sandbox.selectRecordLayout(duplicate, { "版式组": "版式乙" }),
    /E_CONFIG_MISMATCH: 模板版式组重复/
);

assert.throws(
    () => sandbox.selectRecordLayout({ layers: [layer("版式乙 ")] }, { "版式组": "版式乙" }),
    /E_CONFIG_MISMATCH: 模板版式组不是图层组/
);

const fitProfile = {
    text_fit: {
        keys: ["卖点"],
        frame_by_key: { "卖点": { left_ratio: 0.05, right_ratio: 0.95 } }
    }
};
const fitSandbox = {
    $: { global: { __BATCH_INPUTS__: { profile: fitProfile } } },
    app: {
        activeDocument: {
            width: { as: () => 800 },
            height: { as: () => 800 }
        }
    }
};
vm.createContext(fitSandbox);
const fitEnd = source.indexOf("function setTextContentsPreservingStyle");
assert.notEqual(fitEnd, -1, "batch_template.jsx text-fit boundary was not found");
vm.runInContext(source.slice(start, fitEnd), fitSandbox, { filename: scriptPath });
const sellingPointRect = { left: -363, top: 702, right: 1431, bottom: 753 };
const sellingPointFrame = fitSandbox.textFrameForKey("卖点", sellingPointRect);
assert.equal(JSON.stringify(sellingPointFrame), JSON.stringify({ left: 40, top: 702, right: 760, bottom: 753 }));
assert.equal(
    fitSandbox.textMaxWidth({ textItem: {} }, "卖点", sellingPointRect),
    720,
    "JD selling-point fitting must use the configured in-canvas frame"
);
const setTextStart = source.indexOf("function setTextLayer");
const setTextEnd = source.indexOf("function replaceSmartObject", setTextStart);
assert.notEqual(setTextStart, -1, "setTextLayer definition was not found");
assert.notEqual(setTextEnd, -1, "setTextLayer boundary was not found");
const setTextBody = source.slice(setTextStart, setTextEnd);
assert.match(setTextBody, /key === "价格1" && configuredFrame && priceFitMinimumScale !== null/, "only price1 layers with an explicit profile frame may auto-fit");
assert.match(setTextBody, /fitTextToOriginalFrame\(layer, key, originalRect, priceFitMinimumScale\)/, "price fitting must use the PSD-native layer transform");
assert.equal(setTextBody.includes("文案已自动缩字适配"), false, "non-price text overflow must remain visible instead of silently shrinking");

const optionalFitStart = source.indexOf("function fitOptionalImageToTemplateFrame");
const optionalFitEnd = source.indexOf("function markOptionalImageReplacement", optionalFitStart);
assert.notEqual(optionalFitStart, -1, "optional-image fitting definition was not found");
assert.notEqual(optionalFitEnd, -1, "optional-image fitting boundary was not found");
const optionalFitBody = source.slice(optionalFitStart, optionalFitEnd);
assert.equal(optionalFitBody.includes("visiblePixelRect"), false, "optional nested images must use layer bounds, not raster probes");
assert.match(optionalFitBody, /var\s+currentRect\s*=\s*layerRect\(layer\)/, "optional images must start from replacement layer bounds");
assert.match(optionalFitBody, /Math\.min\(targetWidth \/ currentWidth, targetHeight \/ currentHeight\)/, "optional images must use a uniform contain scale");

const imageSetStart = source.indexOf("function setImageLayer");
const imageSetEnd = source.indexOf("function applyRecord", imageSetStart);
assert.notEqual(imageSetStart, -1, "setImageLayer definition was not found");
assert.notEqual(imageSetEnd, -1, "setImageLayer boundary was not found");
const imageSetBody = source.slice(imageSetStart, imageSetEnd);
assert.match(imageSetBody, /fitOptionalImageToTemplateFrame\(layer, targetRect\)/, "gift image smart objects must use the optional placeholder-frame fitter");
assert.match(imageSetBody, /可选素材已替换：!" \+ key/, "gift image replacement must remain visible in the output receipt");
assert.match(source, /function isolateSmartObjectForReplacement\(layer\)/, "image replacement must isolate shared smart objects");
assert.match(source, /stringIDToTypeID\("placedLayerMakeCopy"\)/, "shared smart objects must use New Smart Object via Copy");
assert.match(imageSetBody, /isolateSmartObjectForReplacement\(layer\)/, "all table-driven image replacements must pass through shared-object isolation");

const fieldProfile = {
    fields: [
        {
            field_id: "price_3",
            label: "价格3",
            output_key: "价格3",
            aliases: ["价格3", "price_3"]
        },
        {
            field_id: "main_image",
            label: "产品",
            output_key: "产品",
            type: "smart_object",
            aliases: ["产品", "商品图"]
        }
    ]
};
const fieldSandbox = {
    $: { global: { __BATCH_INPUTS__: { profile: fieldProfile } } }
};
vm.createContext(fieldSandbox);
const fieldEnd = source.indexOf("function profileBindingErrors");
assert.notEqual(fieldEnd, -1, "batch_template.jsx field-binding boundary was not found");
vm.runInContext(source.slice(start, fieldEnd), fieldSandbox, { filename: scriptPath });
const priceLayer = { name: "价格_3", typename: "ArtLayer" };
const productLayer = { name: "产品", typename: "ArtLayer" };
const fieldMatches = fieldSandbox.requiredLayerMatches(
    { text: { "价格_3": [priceLayer] }, image: { "产品": [productLayer] } },
    { name: "价格3", type: "text" }
);
assert.equal(fieldMatches.length, 1, "normalized PSD aliases must resolve punctuation variants");
assert.equal(fieldSandbox.isMainImageKey("产品"), true, "产品 must be treated as the main image field");

const statusSandbox = {
    $: {
        global: {
            __BATCH_INPUTS__: {
                profile: {
                    data_fields_optional: true,
                    required_psd_variables: [{ name: "利益点1", type: "text" }]
                }
            }
        }
    }
};
vm.createContext(statusSandbox);
const statusEnd = source.indexOf("function hasErrorCode");
assert.notEqual(statusEnd, -1, "batch_template.jsx status boundary was not found");
vm.runInContext(source.slice(start, statusEnd), statusSandbox, { filename: scriptPath });
assert.equal(statusSandbox.isProfileTextKey("卖点"), false, "unconfigured optional text must not warn on blank values");
assert.equal(statusSandbox.statusFor({ optionalImageMissing: true }), "需复核", "optional missing images must not report success");

console.log("batch layout selection regression: ok");
