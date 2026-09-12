import { readFileSync, writeFileSync, readdirSync } from "node:fs";
import { gunzipSync } from "node:zlib";

const parts = readdirSync("scripts")
  .filter(name => /^app-payload-\d+\.txt$/.test(name))
  .sort();
if (parts.length !== 5) throw new Error(`Expected 5 App payload parts, found ${parts.length}`);
const encoded = parts.map(name => readFileSync(`scripts/${name}`, "utf8").trim()).join("");
writeFileSync("src/App.jsx", gunzipSync(Buffer.from(encoded, "base64")));
