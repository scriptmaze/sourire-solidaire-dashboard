import { cp, mkdir, rm } from "node:fs/promises";

await rm("dist", { recursive: true, force: true });
await mkdir("dist/vendor", { recursive: true });
await cp("src", "dist", { recursive: true });
await cp("node_modules/xlsx/dist/xlsx.full.min.js", "dist/vendor/xlsx.full.min.js");
