# Bilingual docs render helper.
#
# 순서가 중요: 영문 먼저 (`docs/`) → 한국어 (`docs/ko/`).
# 영문 render 가 `docs/_site/` 를 cleanup 하므로 한국어를 나중에
# 돌려야 `docs/_site/ko/` 가 살아남는다.
#
# 사용:
#   pwsh docs/render-bilingual.ps1
#
# 빌드 후 commit + push 하면 deploy-docs.yaml 이 docs/_site/ 를
# Pages artifact 로 자동 업로드한다.

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "[1/2] Rendering English (docs/) → docs/_site/" -ForegroundColor Cyan
Push-Location $root
try {
    quarto render
    if ($LASTEXITCODE -ne 0) {
        throw "English render failed (exit $LASTEXITCODE)"
    }
} finally {
    Pop-Location
}

Write-Host "[2/2] Rendering Korean (docs/ko/) → docs/_site/ko/" -ForegroundColor Cyan
Push-Location (Join-Path $root 'ko')
try {
    quarto render
    if ($LASTEXITCODE -ne 0) {
        throw "Korean render failed (exit $LASTEXITCODE)"
    }
} finally {
    Pop-Location
}

Write-Host ""
Write-Host "✓ Both sites rendered." -ForegroundColor Green
Write-Host "  - English: docs/_site/index.html"
Write-Host "  - Korean : docs/_site/ko/index.html"
Write-Host ""
Write-Host "Next: git add docs/_site && git commit && git push"
