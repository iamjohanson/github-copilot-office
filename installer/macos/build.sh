#!/bin/bash
# Build script for macOS installer (.pkg)
# Run from repository root: ./installer/macos/build.sh

set -e

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
ROOT_DIR="$SCRIPT_DIR/../.."
BUILD_DIR="$ROOT_DIR/build/macos"
APP_NAME="GitHub Copilot Office Add-in"
VERSION="1.0.0"
IDENTIFIER="com.github.copilot-office-addin"

echo "Building macOS installer..."

# Clean and create build directory
rm -rf "$BUILD_DIR"
mkdir -p "$BUILD_DIR/payload/Applications/$APP_NAME"
mkdir -p "$BUILD_DIR/scripts"

# Build the frontend
echo "Building frontend..."
cd "$ROOT_DIR"
npm run build

# Build the server executable with pkg
echo "Building server executable..."
npx @yao-pkg/pkg src/server-prod.js \
    --targets node18-macos-x64 \
    --output "$BUILD_DIR/payload/Applications/$APP_NAME/copilot-office-server" \
    --compress GZip

# Also build for ARM64 (Apple Silicon)
echo "Building ARM64 executable..."
npx @yao-pkg/pkg src/server-prod.js \
    --targets node18-macos-arm64 \
    --output "$BUILD_DIR/payload/Applications/$APP_NAME/copilot-office-server-arm64" \
    --compress GZip

# Create universal binary
echo "Creating universal binary..."
if command -v lipo &> /dev/null; then
    lipo -create \
        "$BUILD_DIR/payload/Applications/$APP_NAME/copilot-office-server" \
        "$BUILD_DIR/payload/Applications/$APP_NAME/copilot-office-server-arm64" \
        -output "$BUILD_DIR/payload/Applications/$APP_NAME/copilot-office-server-universal" 2>/dev/null || {
            echo "Note: Could not create universal binary, using x64 version"
            cp "$BUILD_DIR/payload/Applications/$APP_NAME/copilot-office-server" \
               "$BUILD_DIR/payload/Applications/$APP_NAME/copilot-office-server-universal"
        }
    mv "$BUILD_DIR/payload/Applications/$APP_NAME/copilot-office-server-universal" \
       "$BUILD_DIR/payload/Applications/$APP_NAME/copilot-office-server"
    rm -f "$BUILD_DIR/payload/Applications/$APP_NAME/copilot-office-server-arm64"
fi

# Copy required files
echo "Copying files..."
cp -r "$ROOT_DIR/dist" "$BUILD_DIR/payload/Applications/$APP_NAME/"
cp -r "$ROOT_DIR/certs" "$BUILD_DIR/payload/Applications/$APP_NAME/"
cp "$ROOT_DIR/manifest.xml" "$BUILD_DIR/payload/Applications/$APP_NAME/"
cp "$SCRIPT_DIR/launchagent/com.github.copilot-office-addin.plist" "$BUILD_DIR/payload/Applications/$APP_NAME/"

# Copy scripts
cp "$SCRIPT_DIR/scripts/preinstall" "$BUILD_DIR/scripts/"
cp "$SCRIPT_DIR/scripts/postinstall" "$BUILD_DIR/scripts/"
chmod +x "$BUILD_DIR/scripts/preinstall"
chmod +x "$BUILD_DIR/scripts/postinstall"

# Build the component package
echo "Building component package..."
pkgbuild \
    --root "$BUILD_DIR/payload" \
    --scripts "$BUILD_DIR/scripts" \
    --identifier "$IDENTIFIER" \
    --version "$VERSION" \
    --install-location "/" \
    "$BUILD_DIR/CopilotOfficeAddin-component.pkg"

# Create distribution XML
cat > "$BUILD_DIR/distribution.xml" << EOF
<?xml version="1.0" encoding="utf-8"?>
<installer-gui-script minSpecVersion="2">
    <title>$APP_NAME</title>
    <organization>$IDENTIFIER</organization>
    <domains enable_localSystem="true" enable_currentUserHome="false"/>
    <options customize="never" require-scripts="true" rootVolumeOnly="true"/>
    <welcome file="welcome.html"/>
    <conclusion file="conclusion.html"/>
    <pkg-ref id="$IDENTIFIER"/>
    <choices-outline>
        <line choice="default">
            <line choice="$IDENTIFIER"/>
        </line>
    </choices-outline>
    <choice id="default"/>
    <choice id="$IDENTIFIER" visible="false">
        <pkg-ref id="$IDENTIFIER"/>
    </choice>
    <pkg-ref id="$IDENTIFIER" version="$VERSION" onConclusion="none">CopilotOfficeAddin-component.pkg</pkg-ref>
</installer-gui-script>
EOF

# Create welcome and conclusion HTML
cat > "$BUILD_DIR/welcome.html" << 'EOF'
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, sans-serif; padding: 20px; }
        h1 { color: #24292f; }
        p { color: #57606a; line-height: 1.5; }
        ul { color: #57606a; }
    </style>
</head>
<body>
    <h1>GitHub Copilot Office Add-in</h1>
    <p>This installer will set up the GitHub Copilot Office Add-in on your Mac.</p>
    <p>The installer will:</p>
    <ul>
        <li>Install the add-in server application</li>
        <li>Register the add-in with Word, PowerPoint, and Excel</li>
        <li>Configure the service to start automatically</li>
        <li>Trust the required SSL certificate</li>
    </ul>
    <p>Click Continue to proceed with the installation.</p>
</body>
</html>
EOF

cat > "$BUILD_DIR/conclusion.html" << 'EOF'
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, sans-serif; padding: 20px; }
        h1 { color: #24292f; }
        p { color: #57606a; line-height: 1.5; }
        .success { color: #1a7f37; font-weight: 600; }
    </style>
</head>
<body>
    <h1>Installation Complete!</h1>
    <p class="success">✓ GitHub Copilot Office Add-in has been installed successfully.</p>
    <p>The add-in service is now running in the background.</p>
    <p><strong>Next steps:</strong></p>
    <ol>
        <li>Open Word, PowerPoint, or Excel</li>
        <li>Look for the "GitHub Copilot" button on the Home ribbon</li>
        <li>Click the button to open the Copilot panel</li>
    </ol>
    <p>The service will start automatically when you log in.</p>
</body>
</html>
EOF

# Build the final distribution package
echo "Building distribution package..."
productbuild \
    --distribution "$BUILD_DIR/distribution.xml" \
    --resources "$BUILD_DIR" \
    --package-path "$BUILD_DIR" \
    "$BUILD_DIR/CopilotOfficeAddin-$VERSION.pkg"

# Clean up intermediate files
rm -f "$BUILD_DIR/CopilotOfficeAddin-component.pkg"
rm -f "$BUILD_DIR/distribution.xml"
rm -f "$BUILD_DIR/welcome.html"
rm -f "$BUILD_DIR/conclusion.html"
rm -rf "$BUILD_DIR/payload"
rm -rf "$BUILD_DIR/scripts"

echo ""
echo "✓ macOS installer built successfully!"
echo "  Output: $BUILD_DIR/CopilotOfficeAddin-$VERSION.pkg"
echo ""
echo "To sign the package for distribution (optional):"
echo "  productsign --sign 'Developer ID Installer: Your Name' \\"
echo "    '$BUILD_DIR/CopilotOfficeAddin-$VERSION.pkg' \\"
echo "    '$BUILD_DIR/CopilotOfficeAddin-$VERSION-signed.pkg'"
