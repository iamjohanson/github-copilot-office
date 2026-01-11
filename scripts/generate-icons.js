/**
 * Generate Office Add-in icons from logo.svg
 * Run: node scripts/generate-icons.js
 */
const sharp = require('sharp');
const path = require('path');
const fs = require('fs');
const toIco = require('to-ico');

const sizes = [16, 32, 64, 80, 128, 256];
const inputFile = path.resolve(__dirname, '../logo.svg');
const outputDir = path.resolve(__dirname, '../src/ui/public');
const installerDir = path.resolve(__dirname, '../installer/windows');

async function generateIcons() {
  // Ensure output directories exist
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  console.log('Generating icons from logo.svg...');

  const pngPaths = [];
  for (const size of sizes) {
    const outputFile = path.join(outputDir, `icon-${size}.png`);
    
    await sharp(inputFile)
      .resize(size, size, {
        fit: 'contain',
        background: { r: 0, g: 0, b: 0, alpha: 0 }
      })
      .png()
      .toFile(outputFile);
    
    console.log(`  ✓ icon-${size}.png`);
    pngPaths.push(outputFile);
  }

  console.log(`\nPNG icons saved to: ${outputDir}`);

  // Generate Windows .ico file (contains multiple sizes)
  console.log('\nGenerating Windows .ico file...');
  try {
    const icoSizes = [16, 32, 64, 128, 256];
    const pngBuffers = icoSizes.map(s => fs.readFileSync(path.join(outputDir, `icon-${s}.png`)));
    const icoBuffer = await toIco(pngBuffers);
    const icoPath = path.join(installerDir, 'app.ico');
    fs.writeFileSync(icoPath, icoBuffer);
    console.log(`  ✓ app.ico saved to: ${icoPath}`);
  } catch (err) {
    console.log(`  ⚠ Could not generate .ico: ${err.message}`);
  }

  // Copy icon for macOS installer (uses PNG)
  const macosInstallerIcon = path.resolve(__dirname, '../installer/macos/icon.png');
  fs.copyFileSync(path.join(outputDir, 'icon-256.png'), macosInstallerIcon);
  console.log(`  ✓ icon.png copied to: ${macosInstallerIcon}`);
}

generateIcons().catch(err => {
  console.error('Error generating icons:', err);
  process.exit(1);
});
