const fs = require("fs");
const path = require("path");
const sharp = require("sharp");

const assetsDirectory = path.join(__dirname, "..", "appPackage", "assets");
const svgPath = path.join(assetsDirectory, "coreai-mark.svg");
const svg = fs.readFileSync(svgPath);
const outputs = [
  ["icon-16.png", 16],
  ["icon-32.png", 32],
  ["icon-64.png", 64],
  ["icon-80.png", 80],
  ["icon-128.png", 128],
  ["color.png", 192],
  ["outline.png", 32],
];

Promise.all(outputs.map(([filename, size]) =>
  sharp(svg).resize(size, size).png().toFile(path.join(assetsDirectory, filename)),
)).catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
