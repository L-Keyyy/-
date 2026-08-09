import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const extensionRoot = new URL(
  "../tools/web-ai-editor-extension/",
  import.meta.url,
);

test("ships a reusable Web AI visual editor extension", async () => {
  const manifest = JSON.parse(
    await readFile(new URL("manifest.json", extensionRoot), "utf8"),
  );
  const background = await readFile(
    new URL("background.js", extensionRoot),
    "utf8",
  );
  const content = await readFile(
    new URL("content.js", extensionRoot),
    "utf8",
  );
  const styles = await readFile(
    new URL("content.css", extensionRoot),
    "utf8",
  );

  assert.equal(manifest.manifest_version, 3);
  assert.equal(manifest.background.service_worker, "background.js");
  assert.ok(manifest.permissions.includes("activeTab"));
  assert.ok(manifest.permissions.includes("scripting"));
  assert.match(background, /chrome\.scripting\.executeScript/);
  assert.match(content, /function undo\(\)/);
  assert.match(content, /function redo\(\)/);
  assert.match(content, /web-ai-change-set\/v1/);
  assert.match(content, /复制AI指令/);
  assert.match(content, /导出修改单/);
  assert.match(content, /insertPlaceholder/);
  assert.match(content, /getDescriptor/);
  assert.match(content, /function askUser/);
  assert.doesNotMatch(content, /window\.prompt/);
  assert.match(styles, /\.wae-page-selected/);
  assert.match(styles, /#wae-root/);
});
