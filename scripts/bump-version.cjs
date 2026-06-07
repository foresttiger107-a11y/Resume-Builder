"use strict";

/**
 * Bump semver patch in package.json (e.g. 4.2.0 → 4.2.1).
 * Used before packaging an app update: npm run release
 */
const fs = require("fs");
const path = require("path");

const pkgPath = path.join(__dirname, "..", "package.json");
const pkg = JSON.parse(fs.readFileSync(pkgPath, "utf8"));

const raw = String(pkg.version || "0.0.0").trim();
const parts = raw.split(".").map((n) => parseInt(n, 10));
while (parts.length < 3) {
  parts.push(0);
}
if (parts.some((n) => Number.isNaN(n))) {
  console.error(`Invalid version in package.json: ${raw}`);
  process.exit(1);
}

parts[2] += 1;
const next = parts.join(".");
pkg.version = next;

fs.writeFileSync(pkgPath, `${JSON.stringify(pkg, null, 2)}\n`, "utf8");
console.log(`Resume Builder version: ${raw} → ${next}`);
