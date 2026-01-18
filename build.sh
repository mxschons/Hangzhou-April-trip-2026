#!/bin/bash
# Converts the markdown file back to docx and pdf
# Using Jiangnan Style Guide - Hangzhou architectural tradition
# Colors: Ink #1C1917, Cinnabar #C2675B (links), Celadon #6B8F82 (urls)

cd "$(dirname "$0")"
eval "$(/usr/libexec/path_helper)" 2>/dev/null

echo ""
echo "╔══════════════════════════════════════════════╗"
echo "║  MxSchons Tours - Document Builder           ║"
echo "║  Jiangnan Style Guide                        ║"
echo "╚══════════════════════════════════════════════╝"
echo ""

# Build DOCX with original styles
echo "→ Creating Word document..."
pandoc "Hangzhou_Family_Trip_April_2026_Final.md" \
  --reference-doc="Hangzhou_Family_Trip_April_2026_Final.docx" \
  -o "Hangzhou_Family_Trip_April_2026_Final_NEW.docx"
echo "  ✓ Hangzhou_Family_Trip_April_2026_Final_NEW.docx"

# Build PDF with Jiangnan styling
echo "→ Creating styled PDF..."
pandoc "Hangzhou_Family_Trip_April_2026_Final.md" \
  -o "Hangzhou_Family_Trip_April_2026_Final.pdf" \
  --pdf-engine=xelatex \
  -V geometry:margin=1in \
  -V mainfont="PingFang SC" \
  -V fontsize=11pt \
  -V linkcolor="[HTML]{C2675B}" \
  -V urlcolor="[HTML]{6B8F82}" \
  -V linestretch=1.4 \
  2>&1 | grep -v "^$" | head -3

if [ -f "Hangzhou_Family_Trip_April_2026_Final.pdf" ]; then
  SIZE=$(ls -lh "Hangzhou_Family_Trip_April_2026_Final.pdf" | awk '{print $5}')
  echo "  ✓ Hangzhou_Family_Trip_April_2026_Final.pdf ($SIZE)"
else
  echo "  ✗ PDF generation failed"
fi

echo ""
echo "══════════════════════════════════════════════"
echo "  Build complete! 一路平安"
echo "══════════════════════════════════════════════"
echo ""
