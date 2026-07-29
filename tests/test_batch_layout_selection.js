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
const end = source.indexOf("function applyPreflightIssue");
assert.notEqual(start, -1, "batch_template.jsx test prelude was not found");
assert.notEqual(end, -1, "batch_template.jsx layout-selection boundary was not found");

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

console.log("batch layout selection regression: ok");
