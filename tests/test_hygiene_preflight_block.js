const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const root = path.resolve(__dirname, "..");
const scriptPath = path.join(root, "L0_Windows命令行版", "_internal", "template_prepare.jsx");
const source = fs.readFileSync(scriptPath, "utf8");
const boundary = source.indexOf("function main()");
assert.notEqual(boundary, -1, "template preflight boundary was not found");

const sandbox = {
    $: {
        global: {
            __TEMPLATE_PREP_INPUTS__: {
                profile: {
                    profile_id: "hygiene-tmall-v1.2",
                    record_layout: { groups: ["版式甲", "版式乙"] },
                    active_layout_groups: ["版式甲"],
                    required_psd_variables: [],
                },
            },
        },
    },
};
vm.createContext(sandbox);
vm.runInContext(source.slice(0, boundary).replace(/^#target[^\n]*\n/, ""), sandbox, { filename: scriptPath });

function group(name, layers = []) {
    const result = { name, typename: "LayerSet", layers, parent: null };
    for (const child of layers) child.parent = result;
    return result;
}

const document = { layers: [], parent: null };
const active = group("版式甲", [group("产品"), group("标题"), group("片数"), group("下帖")]);
const inactive = group("版式乙", [group("产品")]);
active.parent = document;
inactive.parent = document;
document.layers.push(active, inactive);

const result = sandbox.inspectPreparation(document, sandbox.buildLayerIndex(document));
assert.equal(result.status, "BLOCKED_PREP_REQUIRED");
assert.match(result.message, /E_TEMPLATE_PREP_REQUIRED/);

console.log("hygiene multi-layout preflight block regression: ok");
