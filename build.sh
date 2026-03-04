#!/bin/bash

echo "🚀 Building Health & Schedule Reminder VS Code Extension..."
echo ""

# Step 1: Install dependencies
echo "📦 Step 1: Installing dependencies..."
npm install
echo "✅ Dependencies installed!"
echo ""

# Step 2: Install vsce globally
echo "🔧 Step 2: Installing vsce (VS Code Extension packager)..."
npm install -g @vscode/vsce
echo "✅ vsce installed!"
echo ""

# Step 3: Compile TypeScript
echo "⚙️  Step 3: Compiling TypeScript..."
npm run compile
echo "✅ TypeScript compiled!"
echo ""

# Step 4: Package VSIX
echo "📦 Step 4: Packaging VSIX..."
vsce package --no-dependencies
echo ""
echo "🎉 SUCCESS! Your VSIX file is ready!"
echo ""
echo "📌 To install in VS Code:"
echo "   code --install-extension health-schedule-reminder-1.0.0.vsix"
echo ""
echo "   OR: VS Code → Extensions → '...' → Install from VSIX"
