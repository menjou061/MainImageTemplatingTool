const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const repositoryRoot = path.resolve(__dirname, "..");
const scriptPath = path.join(repositoryRoot, "L0_Windows命令行版", "_internal", "batch_template.jsx");
const source = fs.readFileSync(scriptPath, "utf8");
const start = source.indexOf("var REPORT_NAME");
const end = source.indexOf("function insideDisabledSwitch");
assert.notEqual(start, -1, "batch_template.jsx test prelude was not found");
assert.notEqual(end, -1, "batch_template.jsx gift-switch boundary was not found");

const sandbox = { $: { global: { __BATCH_INPUTS__: { profile: {} } } } };
vm.createContext(sandbox);
vm.runInContext(source.slice(start, end), sandbox, { filename: scriptPath });

function layer(name) {
    return { name, typename: "ArtLayer", visible: true };
}

function group(name, layers = []) {
    return { name, typename: "LayerSet", visible: true, layers };
}

const giftRegion = group("#赠品区域", [group("赠品槽位", [layer("!赠品图2"), layer("!赠品图3")])]);
const layerIndex = { switches: { "赠品区域": [giftRegion] } };
const record = { "赠品图2": "gift-2.png", "赠品图3": "" };
const result = { issues: [], codes: [], optionalImageReplaced: { "赠品图2": true } };

assert.equal(sandbox.optionalGroupHasAnyValues([giftRegion], record), true);
assert.equal(sandbox.optionalGroupHasValues([giftRegion], record), false);
sandbox.setSwitches(layerIndex, record, result);
assert.equal(record["赠品区域"], "是", "a two-gift row must keep the lower gift region enabled");
assert.equal(giftRegion.visible, true);
sandbox.reconcileOptionalGroups(layerIndex, record, result);
assert.equal(record["赠品区域"], "是", "a successfully replaced second gift must remain visible");
assert.equal(giftRegion.visible, true);

console.log("hygiene gift switch regression: ok");
