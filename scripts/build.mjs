// 產生單一檔案版本：dist/proposal-generator.html（可直接開啟或發布為 Claude Artifact）
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
let html = readFileSync(join(root, "index.html"), "utf8");
for (const name of ["parse", "gen", "app"]) {
  const tag = `<script src="src/${name}.js"></script>`;
  const code = readFileSync(join(root, "src", `${name}.js`), "utf8");
  if (code.includes("</script")) throw new Error(`src/${name}.js 含有 </script，無法內嵌`);
  if (!html.includes(tag)) throw new Error(`index.html 找不到 ${tag}`);
  html = html.replace(tag, () => `<script>\n${code}\n</script>`);
}
mkdirSync(join(root, "dist"), { recursive: true });
writeFileSync(join(root, "dist", "proposal-generator.html"), html);
console.log(`dist/proposal-generator.html（${Math.round(html.length / 1024)} KB）`);
