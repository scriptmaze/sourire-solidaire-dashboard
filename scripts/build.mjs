import { cp, mkdir, rm, writeFile } from "node:fs/promises";

await rm("dist", { recursive: true, force: true });
await mkdir("dist/vendor", { recursive: true });
await cp("src", "dist", { recursive: true });
await cp("node_modules/xlsx/dist/xlsx.full.min.js", "dist/vendor/xlsx.full.min.js");
await writeFile("dist/CNAME", "tableau-de-bord-soso.aclclouds.online\n");
