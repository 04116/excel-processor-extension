#!/bin/bash

# PowerBI Data Export Extension - Release Builder
# This script creates a release package for the Chrome extension

echo "🔨 Building PowerBI Data Export Extension release..."

# Get version from manifest.json
VERSION=$(grep '"version"' manifest.json | sed 's/.*"version": "\([^"]*\)".*/\1/')
echo "📦 Version: $VERSION"

# Create release directory
RELEASE_DIR="release"
PACKAGE_NAME="powerbi-data-export-extension-v${VERSION}"
PACKAGE_DIR="$RELEASE_DIR/$PACKAGE_NAME"

echo "🗂️  Creating release directory: $PACKAGE_DIR"
rm -rf "$RELEASE_DIR"
mkdir -p "$PACKAGE_DIR"

# Copy essential files
echo "📄 Copying extension files..."
cp manifest.json "$PACKAGE_DIR/"
cp popup.html "$PACKAGE_DIR/"
cp popup.js "$PACKAGE_DIR/"
cp download.js "$PACKAGE_DIR/"
cp excel-processor.js "$PACKAGE_DIR/"
cp readme.md "$PACKAGE_DIR/"

# Copy libs directory
echo "📚 Copying libraries..."
cp -r libs "$PACKAGE_DIR/"

# Create zip file
echo "🗜️  Creating zip package..."
cd "$RELEASE_DIR"
zip -r "${PACKAGE_NAME}.zip" "$PACKAGE_NAME"
cd ..

# Show results
echo "✅ Release package created successfully!"
echo "📍 Location: $RELEASE_DIR/${PACKAGE_NAME}.zip"
echo "📊 Package contents:"
unzip -l "$RELEASE_DIR/${PACKAGE_NAME}.zip"

echo ""
echo "🚀 Ready for distribution!"
echo "💡 To install: Unzip the file and load the folder in Chrome Developer Mode" 