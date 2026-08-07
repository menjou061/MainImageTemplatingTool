const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const root = path.resolve(__dirname, "..");
const scriptPath = path.join(root, "L0_Windows命令行版", "_internal", "template_prepare.jsx");
const source = fs.readFileSync(scriptPath, "utf8");
const boundary = source.indexOf("function main()");
assert.notEqual(boundary, -1, "template preparation boundary was not found");

const profile = {
    profile_id: "hygiene-tmall-v1.2",
    layout: "record_rows",
    static_product_art: false,
    static_support_art: true,
    record_layout: { groups: ["三栏会员"] },
    required_psd_variables: [
        { name: "商品图", type: "smart_object" },
        { name: "卖点", type: "text" },
        { name: "备注", type: "text" },
        { name: "片数套", type: "text" },
        { name: "片数数量", type: "text" },
        { name: "到手标签", type: "text" },
        { name: "到手", type: "text" },
        { name: "价格活动价", type: "text" }
    ]
};
const sandbox = {
    $: { global: { __TEMPLATE_PREP_INPUTS__: { profile } } },
    LayerKind: { TEXT: "TEXT", SMARTOBJECT: "SMARTOBJECT" }
};
vm.createContext(sandbox);
vm.runInContext(source.slice(0, boundary).replace(/^#target[^\n]*\n/, ""), sandbox, { filename: scriptPath });

function textLayer(name) {
    return { name, typename: "ArtLayer", kind: "TEXT", layers: [], parent: null };
}

function group(name, layers = []) {
    const result = { name, typename: "LayerSet", layers, parent: null };
    for (const child of layers) child.parent = result;
    return result;
}

function documentWith(layers) {
    const document = { layers, parent: null };
    for (const layer of layers) layer.parent = document;
    return document;
}

function skuGroup(name) {
    return group(name, [
        group("产品"),
        textLayer("@卖点"),
        textLayer("@备注"),
        textLayer("@片数套"),
        textLayer("@片数数量"),
        textLayer("@到手")
    ]);
}

const sourceDocument = documentWith([skuGroup("5QDAC618"), skuGroup("12SKU新组合")]);
const index = sandbox.buildLayerIndex(sourceDocument);
assert.equal(sandbox.isStaticMemberFieldsTemplate(sourceDocument, index), true);
assert.equal(sandbox.inspectPreparation(sourceDocument, index).status, "NEEDS_PREP");

const singleLayout = documentWith([group("三栏会员", [
    textLayer("@卖点"), textLayer("@备注"), textLayer("@片数套"),
    textLayer("@片数数量"), textLayer("@到手"),
    { name: "!商品图", typename: "ArtLayer", kind: "SMARTOBJECT", layers: [], parent: null }
])]);
assert.equal(sandbox.isStaticMemberFieldsTemplate(singleLayout, sandbox.buildLayerIndex(singleLayout)), false);

const noGiftProfile = {
    profile_id: "hygiene-tmall-v1.2",
    layout: "record_rows",
    static_product_art: false,
    static_support_art: false,
    record_layout: { groups: ["三栏会员"] },
    optional_psd_variables: ["备注", "赠品图1", "赠品图2", "赠品图3", "赠品文案1", "赠品文案2", "赠品文案3"],
    required_psd_variables: [
        { name: "商品图", type: "smart_object" },
        { name: "卖点", type: "text" },
        { name: "备注", type: "text" },
        { name: "片数套", type: "text" },
        { name: "片数数量", type: "text" },
        { name: "到手标签", type: "text" },
        { name: "到手", type: "text" },
        { name: "价格活动价", type: "text" },
        { name: "赠品图1", type: "smart_object" },
        { name: "赠品图2", type: "smart_object" },
        { name: "赠品图3", type: "smart_object" },
        { name: "赠品文案1", type: "text" },
        { name: "赠品文案2", type: "text" },
        { name: "赠品文案3", type: "text" }
    ]
};
sandbox.$.global.__TEMPLATE_PREP_INPUTS__.profile = noGiftProfile;
sandbox.CHANNEL_PROFILE = noGiftProfile;
const noGiftLayout = documentWith([group("三栏会员", [
    { name: "!商品图", typename: "ArtLayer", kind: "SMARTOBJECT", layers: [], parent: null },
    textLayer("@卖点"), textLayer("@片数套"), textLayer("@片数数量"),
    textLayer("@到手标签"), textLayer("@到手"), textLayer("@价格活动价")
])]);
assert.equal(sandbox.hygieneProblems(noGiftLayout, sandbox.buildLayerIndex(noGiftLayout)).length, 0);

console.log("static member detection regression: ok");
