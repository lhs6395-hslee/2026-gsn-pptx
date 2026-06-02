#!/usr/bin/env bash
# PPT 생성 스킬 설치 스크립트
# 한 번만 실행. 이후엔 Claude Code에서 바로 사용.
set -e

SKILL_DIR="$HOME/.ppt-skill"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "================================================"
echo " PPT 생성 스킬 설치"
echo " 설치 위치: $SKILL_DIR"
echo "================================================"

# ── 1. 디렉토리 구조 생성 ─────────────────────────────────
echo ""
echo "[1/6] 디렉토리 생성..."
mkdir -p "$SKILL_DIR"/{harness,scripts/office/helpers,scripts/office/validators,\
scripts/office/schemas,ahe_tools,templates,traces,evolution,runs}
echo "  ✓"

# ── 2. 스킬 파일 복사 ─────────────────────────────────────
echo "[2/6] 스킬 파일 복사..."
cp -r "$SCRIPT_DIR/harness/"*     "$SKILL_DIR/harness/"
cp -r "$SCRIPT_DIR/scripts/"*     "$SKILL_DIR/scripts/"
cp -r "$SCRIPT_DIR/ahe_tools/"*   "$SKILL_DIR/ahe_tools/" 2>/dev/null || true
cp    "$SCRIPT_DIR/SKILL.md"      "$SKILL_DIR/"
echo "  ✓"

# ── 3. Python 패키지 설치 ─────────────────────────────────
echo "[3/6] Python 패키지 설치..."
pip install anthropic defusedxml Pillow python-pptx -q
echo "  ✓ anthropic, defusedxml, Pillow, python-pptx"

# ── 4. extract-text 설치 ──────────────────────────────────
echo "[4/6] extract-text 설치..."
EXTRACT_TEXT_PATH="$(python3 -c 'import sys; print(sys.prefix)')/bin/extract-text"
cat > "$EXTRACT_TEXT_PATH" << 'PYEOF'
#!/usr/bin/env python3
"""extract-text: pptx 슬라이드별 텍스트 추출"""
import sys
from pptx import Presentation

if len(sys.argv) < 2:
    print("Usage: extract-text <file.pptx>", file=sys.stderr)
    sys.exit(1)

prs = Presentation(sys.argv[1])
for i, slide in enumerate(prs.slides, 1):
    print(f"\n## Slide {i}\n")
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                if text:
                    print(text)
PYEOF
chmod +x "$EXTRACT_TEXT_PATH"
echo "  ✓ extract-text → $EXTRACT_TEXT_PATH"

# ── 5. 시스템 의존성 확인 ─────────────────────────────────
echo "[5/6] 시스템 의존성 확인..."

check_install() {
    local cmd="$1" brew_pkg="$2" apt_pkg="$3"
    if command -v "$cmd" &>/dev/null; then
        echo "  ✓ $cmd 이미 설치됨"
    elif [[ "$OSTYPE" == "darwin"* ]]; then
        echo "  → macOS: brew install $brew_pkg"
        brew install "$brew_pkg" 2>/dev/null || echo "  ⚠️  수동 설치 필요: brew install $brew_pkg"
    else
        echo "  → Linux: sudo apt install $apt_pkg"
        sudo apt-get install -y "$apt_pkg" 2>/dev/null || echo "  ⚠️  수동: sudo apt install $apt_pkg"
    fi
}

check_install soffice libreoffice libreoffice
check_install pdftoppm poppler poppler-utils

# ── 6. git 초기화 (harness 변경 추적용) ──────────────────
echo "[6/6] git 초기화..."
cd "$SKILL_DIR"
if [ ! -d .git ]; then
    git init -q
    git config user.email "ppt-skill@local"
    git config user.name "PPT Skill"
fi

cat > .gitignore << 'EOF'
runs/
traces/
*.pdf
slide-*.jpg
thumbs*.jpg
unpacked/
__pycache__/
*.pyc
EOF

git add harness/ SKILL.md 2>/dev/null || true
git commit -m "seed harness v0.0" -q 2>/dev/null || \
    git commit --allow-empty -m "seed harness v0.0" -q 2>/dev/null || true
echo "  ✓"

# ── 완료 메시지 ──────────────────────────────────────────
echo ""
echo "================================================"
echo " 설치 완료!"
echo "================================================"
echo ""
echo "다음 단계:"
echo ""
echo "1. 템플릿 배치:"
echo "   cp your_template.pptx ~/.ppt-skill/templates/default.pptx"
echo ""
echo "2. Claude API 백엔드 설정 (~/.zshrc 또는 ~/.bashrc에 추가):"
echo ""
echo "   # Team Plan (Vertex AI — 권장):"
echo "   export ANTHROPIC_VERTEX_PROJECT_ID='your-gcp-project-id'"
echo "   export CLOUD_ML_REGION='global'  # 또는 us-east5"
echo "   export ANTHROPIC_DEFAULT_SONNET_MODEL='claude-sonnet-4-6'"
echo ""
echo "   # Team Plan (Bedrock):"
echo "   export AWS_REGION='us-east-1'"
echo "   export ANTHROPIC_DEFAULT_SONNET_MODEL='us.anthropic.claude-sonnet-4-5-20250929-v1:0'"
echo ""
echo "   # Anthropic API (개인 키 보유 시):"
echo "   export ANTHROPIC_API_KEY='sk-ant-...'"
echo ""
echo "   ※ ~/.claude/settings.json의 env 값이 있으면 자동으로 감지됩니다."
echo ""
echo "3. Claude Code에서 사용:"
echo "   > Kafka 아키텍처 PPT 10장 만들어줘"
echo ""
echo "4. 터미널에서 직접 사용 (선택):"
echo "   alias ppt='~/.ppt-skill/bin/ppt'"
echo "   ppt \"Kafka 아키텍처\""
echo ""
