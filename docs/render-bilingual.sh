#!/usr/bin/env bash
# Bilingual docs render helper.
#
# 순서가 중요: 영문 먼저 (docs/) → 한국어 (docs/ko/).
# 영문 render 가 docs/_site/ 를 cleanup 하므로 한국어를 나중에
# 돌려야 docs/_site/ko/ 가 살아남는다.
#
# 사용:
#   bash docs/render-bilingual.sh
#
# 빌드 후 commit + push 하면 deploy-docs.yaml 이 docs/_site/ 를
# Pages artifact 로 자동 업로드한다.

set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "[1/2] Rendering English (docs/) → docs/_site/"
( cd "$ROOT" && quarto render )

echo "[2/2] Rendering Korean (docs/ko/) → docs/_site/ko/"
( cd "$ROOT/ko" && quarto render )

echo ""
echo "✓ Both sites rendered."
echo "  - English: docs/_site/index.html"
echo "  - Korean : docs/_site/ko/index.html"
echo ""
echo "Next: git add docs/_site && git commit && git push"
