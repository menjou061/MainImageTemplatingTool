const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const script = fs.readFileSync(
    path.join(__dirname, "..", "L0_Windows命令行版", "_internal", "L0_Run.ps1"),
    "utf8"
);

const override = "$selectedProfile | Add-Member -NotePropertyName text_fit -NotePropertyValue $selectedVariant.text_fit -Force";
assert.equal(
    script.split(override).length - 1,
    2,
    "variant text_fit must be forwarded before the initial profile JSON and after sheet-driven variant selection"
);
console.log("L0 variant profile overrides: ok");
