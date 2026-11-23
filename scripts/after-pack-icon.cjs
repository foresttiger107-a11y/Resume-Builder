"use strict";

/**
 * Embed `assets/icon.png` into the Windows .exe as the application icon.
 * Needed when `win.signAndEditExecutable` is false (skips electron-builder's rcedit),
 * or when that step fails on some machines.
 */
module.exports = async function afterPackIcon(context) {
  if (context.electronPlatformName !== "win32") return;

  const path = require("path");
  const fs = require("fs");
  const os = require("os");

  const projectDir = context.packager.info.projectDir;
  const pngPath = path.join(projectDir, "assets", "icon.png");
  if (!fs.existsSync(pngPath)) {
    return;
  }

  const exePath = path.join(
    context.appOutDir,
    `${context.packager.appInfo.productFilename}.exe`,
  );
  if (!fs.existsSync(exePath)) {
    return;
  }

  const pngToIco = (await import("png-to-ico")).default;
  const { rcedit } = await import("rcedit");

  const icoBuf = await pngToIco(pngPath);
  const tmpIco = path.join(os.tmpdir(), `prompt-builder-icon-${process.pid}.ico`);
  fs.writeFileSync(tmpIco, icoBuf);
  try {
    await rcedit(exePath, { icon: tmpIco });
  } finally {
    try {
      fs.unlinkSync(tmpIco);
    } catch {
      /* ignore */
    }
  }
};
