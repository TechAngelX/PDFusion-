#!/bin/bash

echo "====================================="
echo "  PDFusion macOS Build"
echo "====================================="
echo ""

# Detect architecture
ARCH=$(uname -m)
if [ "$ARCH" = "arm64" ]; then
    RUNTIME="osx-arm64"
    echo "Detected: Apple Silicon (arm64)"
else
    RUNTIME="osx-x64"
    echo "Detected: Intel (x64)"
fi
echo ""

# Clean previous builds
echo "[1/5] Cleaning previous builds..."
rm -rf bin obj
rm -rf ~/Desktop/PDFusion.app
echo "Done"
echo ""

# Restore and build
echo "[2/5] Building..."
dotnet restore PDFusion.csproj
dotnet build PDFusion.csproj -c Release
if [ $? -ne 0 ]; then
    echo "Build failed!"
    exit 1
fi
echo "Done"
echo ""

# Publish
echo "[3/5] Publishing..."
dotnet publish PDFusion.csproj -c Release -r $RUNTIME --self-contained \
    -p:PublishSingleFile=true \
    -p:IncludeNativeLibrariesForSelfExtract=true

if [ $? -ne 0 ]; then
    echo "Publish failed!"
    exit 1
fi
echo "Done"
echo ""

# Create .app bundle
echo "[4/5] Creating macOS app bundle..."
APP_NAME="PDFusion"
APP_BUNDLE=~/Desktop/${APP_NAME}.app
CONTENTS_DIR=${APP_BUNDLE}/Contents
MACOS_DIR=${CONTENTS_DIR}/MacOS
RESOURCES_DIR=${CONTENTS_DIR}/Resources

mkdir -p "${MACOS_DIR}"
mkdir -p "${RESOURCES_DIR}"

# Copy executable
cp "bin/Release/net10.0/${RUNTIME}/publish/PDFusion" "${MACOS_DIR}/"
chmod +x "${MACOS_DIR}/PDFusion"

# Create Info.plist
cat > "${CONTENTS_DIR}/Info.plist" << 'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key>
    <string>PDFusion</string>
    <key>CFBundleDisplayName</key>
    <string>PDFusion</string>
    <key>CFBundleIdentifier</key>
    <string>com.rickiangel.pdfusion</string>
    <key>CFBundleVersion</key>
    <string>1.0.0</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0.0</string>
    <key>CFBundleExecutable</key>
    <string>PDFusion</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>LSMinimumSystemVersion</key>
    <string>11.0</string>
    <key>NSHighResolutionCapable</key>
    <true/>
    <key>NSHumanReadableCopyright</key>
    <string>© 2025 Ricki Angel</string>
</dict>
</plist>
PLIST

# Create icon using sips (converts PNG to ICNS)
echo "[5/5] Creating app icon..."

# Create iconset directory
ICONSET_DIR="/tmp/PDFusion.iconset"
rm -rf "$ICONSET_DIR"
mkdir -p "$ICONSET_DIR"

# Create a simple PNG icon using Python (available on macOS)
python3 << 'PYTHON'
import struct
import zlib
import os

def create_png(width, height, filename):
    # Create pixel data for a purple gradient with documents
    pixels = []

    for y in range(height):
        row = [0]  # Filter byte
        for x in range(width):
            # Background: rounded purple rectangle
            margin = width // 16
            corner_radius = width // 5

            in_rect = margin <= x < width - margin and margin <= y < height - margin
            in_corner = False

            if in_rect:
                # Check corners for rounded effect
                corners = [
                    (margin + corner_radius, margin + corner_radius),
                    (width - margin - corner_radius - 1, margin + corner_radius),
                    (margin + corner_radius, height - margin - corner_radius - 1),
                    (width - margin - corner_radius - 1, height - margin - corner_radius - 1)
                ]
                for cx, cy in corners:
                    dx, dy = x - cx, y - cy
                    if ((x < margin + corner_radius and y < margin + corner_radius) or
                        (x > width - margin - corner_radius and y < margin + corner_radius) or
                        (x < margin + corner_radius and y > height - margin - corner_radius) or
                        (x > width - margin - corner_radius and y > height - margin - corner_radius)):
                        if dx*dx + dy*dy > corner_radius*corner_radius:
                            in_corner = True
                            break

            if in_rect and not in_corner:
                # Purple gradient
                t = (x + y) / (width + height)
                r = int(139 + t * (124 - 139))
                g = int(92 + t * (58 - 92))
                b = int(246 + t * (237 - 246))
                a = 255
            else:
                r, g, b, a = 0, 0, 0, 0

            # Draw white document (left)
            doc_left = int(width * 0.18)
            doc_top = int(height * 0.22)
            doc_w = int(width * 0.28)
            doc_h = int(height * 0.56)
            fold = int(width * 0.08)

            if doc_left <= x < doc_left + doc_w and doc_top <= y < doc_top + doc_h:
                if x >= doc_left + doc_w - fold and y < doc_top + fold:
                    if (x - (doc_left + doc_w - fold)) + (y - doc_top) < fold:
                        r, g, b, a = 255, 255, 255, 255
                else:
                    r, g, b, a = 255, 255, 255, 255

            # Draw arrow
            arrow_x = int(width * 0.48)
            arrow_y = int(height * 0.44)
            arrow_w = int(width * 0.12)
            arrow_h = int(height * 0.12)

            ax, ay = x - arrow_x, y - arrow_y
            if 0 <= ax < arrow_w * 0.6 and arrow_h * 0.3 <= ay < arrow_h * 0.7:
                r, g, b, a = 16, 185, 129, 255
            if arrow_w * 0.4 <= ax < arrow_w and 0 <= ay < arrow_h:
                tip_x = arrow_w - 1
                if abs(ay - arrow_h/2) <= (tip_x - ax) * 0.6:
                    r, g, b, a = 16, 185, 129, 255

            # Draw second document (right)
            doc2_left = int(width * 0.58)
            if doc2_left <= x < doc2_left + doc_w and doc_top <= y < doc_top + doc_h:
                if x >= doc2_left + doc_w - fold and y < doc_top + fold:
                    if (x - (doc2_left + doc_w - fold)) + (y - doc_top) < fold:
                        r, g, b, a = 255, 255, 255, 255
                else:
                    r, g, b, a = 255, 255, 255, 255

            # Green checkmark circle
            cx, cy = int(width * 0.72), int(height * 0.44)
            cr = int(width * 0.12)
            if (x - cx)**2 + (y - cy)**2 <= cr**2:
                r, g, b, a = 16, 185, 129, 255

            row.extend([r, g, b, a])
        pixels.append(bytes(row))

    def make_png(w, h, rows):
        def crc32(data):
            return zlib.crc32(data) & 0xffffffff

        signature = b'\x89PNG\r\n\x1a\n'

        # IHDR
        ihdr_data = struct.pack('>IIBBBBB', w, h, 8, 6, 0, 0, 0)
        ihdr = struct.pack('>I', 13) + b'IHDR' + ihdr_data + struct.pack('>I', crc32(b'IHDR' + ihdr_data))

        # IDAT
        raw_data = b''.join(rows)
        compressed = zlib.compress(raw_data, 9)
        idat = struct.pack('>I', len(compressed)) + b'IDAT' + compressed + struct.pack('>I', crc32(b'IDAT' + compressed))

        # IEND
        iend = struct.pack('>I', 0) + b'IEND' + struct.pack('>I', crc32(b'IEND'))

        return signature + ihdr + idat + iend

    png_data = make_png(width, height, pixels)
    with open(filename, 'wb') as f:
        f.write(png_data)

# Create icons at different sizes
sizes = [16, 32, 64, 128, 256, 512, 1024]
iconset_dir = "/tmp/PDFusion.iconset"

for size in sizes:
    create_png(size, size, f"{iconset_dir}/icon_{size}x{size}.png")
    if size <= 512:
        create_png(size * 2, size * 2, f"{iconset_dir}/icon_{size}x{size}@2x.png")

print("PNG icons created")
PYTHON

# Convert iconset to icns
iconutil -c icns "$ICONSET_DIR" -o "${RESOURCES_DIR}/AppIcon.icns"

if [ $? -eq 0 ]; then
    echo "Icon created successfully"
else
    echo "Warning: Could not create icon (iconutil failed)"
fi

# Clean up
rm -rf "$ICONSET_DIR"

echo ""
echo "====================================="
echo "  Build Complete!"
echo "====================================="
echo ""
echo "App location: ~/Desktop/PDFusion.app"
echo ""
echo "Double-click PDFusion.app to run!"
echo ""

# Ask to open folder
read -p "Open Desktop? (y/n) " -n 1 -r
echo
if [[ $REPLY =~ ^[Yy]$ ]]; then
    open ~/Desktop
fi
