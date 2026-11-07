#!/bin/bash
# Builds OfficeLite and packages it as a double clickable OfficeLite.app in target/release.
set -euo pipefail
cd "$(dirname "${BASH_SOURCE[0]}")/.."

# Keep local paths, including the builder's username, out of the published binary.
export RUSTFLAGS="${RUSTFLAGS:-} --remap-path-prefix=$HOME=~ --remap-path-prefix=$PWD=."

cargo build --release

RELEASE_DIR="${CARGO_TARGET_DIR:-target}/release"
APP="$RELEASE_DIR/OfficeLite.app"
rm -rf "$APP"
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"
cp "$RELEASE_DIR/officelite" "$APP/Contents/MacOS/OfficeLite"
cp AppIcon.icns "$APP/Contents/Resources/AppIcon.icns"

cat > "$APP/Contents/Info.plist" <<'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>OfficeLite</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
    <key>CFBundleIdentifier</key>
    <string>com.officelite.app</string>
    <key>CFBundleName</key>
    <string>OfficeLite</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0.0</string>
    <key>CFBundleVersion</key>
    <string>1</string>
    <key>LSMinimumSystemVersion</key>
    <string>11.0</string>
</dict>
</plist>
PLIST

codesign --force --sign - "$APP"
echo "Built $APP"
