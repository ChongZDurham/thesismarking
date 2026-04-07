#!/bin/bash
set -euo pipefail

cd "$(dirname "$0")"

python3 -m pip install -r thesis_reviewer_requirements.txt

python3 -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name ThesisReviewerMac \
  --workpath build_thesis_reviewer_mac \
  --distpath dist_thesis_reviewer_mac \
  --osx-bundle-identifier cn.socialwork.thesisreviewer \
  --collect-data docx \
  thesis_reviewer_app.py

echo
echo "Built macOS app bundle at:"
echo "  $(pwd)/dist_thesis_reviewer_mac/ThesisReviewerMac.app"
echo
echo "Note: inline annotated Word copies remain Windows-only."
