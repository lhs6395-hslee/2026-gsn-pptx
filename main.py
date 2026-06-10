#!/usr/bin/env python3
"""
PPT 생성 CLI 진입점
Usage: python main.py --topic "Kafka 아키텍처" [--slides 10] [--audience 전문가] [--evolve]
"""
import argparse
import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

SKILL_DIR = Path.home() / ".ppt-skill"


def _inject_claude_settings_env() -> None:
    """
    ~/.claude/settings.json의 env 블록에 있는 API 관련 변수를
    현재 프로세스 환경에 주입한다 (이미 설정된 값은 덮어쓰지 않음).
    """
    settings = Path.home() / ".claude" / "settings.json"
    if not settings.exists():
        return
    try:
        data = json.loads(settings.read_text())
        for key, val in data.get("env", {}).items():
            if key not in os.environ:
                os.environ[key] = str(val)
    except Exception:
        pass
DEFAULT_TEMPLATE = SKILL_DIR / "templates" / "default.pptx"
RUNS_DIR = SKILL_DIR / "runs"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="PPT 자동 생성")
    parser.add_argument("--topic", required=True, help="발표 주제")
    parser.add_argument("--slides", type=int, default=10, help="슬라이드 수 (기본: 10)")
    parser.add_argument("--audience", default="전문가", help="대상 청중 (기본: 전문가)")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="템플릿 경로")
    parser.add_argument("--output", default=str(Path.home() / "Desktop"), help="출력 디렉토리")
    parser.add_argument("--evolve", action="store_true", help="AHE 진화 루프 활성화")
    parser.add_argument(
        "--approve-evolve",
        action="store_true",
        help="QA 결함 감지 시 진화 루프 자동 승인 (기본: OFF — human-in-loop §1). "
        "미지정 시 결함을 감지해도 제안만 출력하고 실행하지 않는다.",
    )
    parser.add_argument(
        "--backend",
        choices=["vertex", "bedrock", "anthropic", "auto"],
        default="auto",
        help="Claude API 백엔드 (기본: auto — settings.json 우선순위에 따라 자동 선택)",
    )
    return parser.parse_args()


def main() -> None:
    _inject_claude_settings_env()
    args = parse_args()
    template = Path(args.template)

    if not template.exists():
        print(f"✗ 템플릿 없음: {template}", file=sys.stderr)
        print("  cp your_template.pptx ~/.ppt-skill/templates/default.pptx", file=sys.stderr)
        sys.exit(1)

    # 작업 디렉토리
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    slug = args.topic.replace(" ", "_")[:40]
    work_dir = RUNS_DIR / f"{stamp}_{slug}"
    work_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, work_dir / "template.pptx")

    # --backend 명시 시 환경변수로 전달해 generate_plan_with_claude가 읽도록 함
    if args.backend != "auto":
        os.environ["PPT_SKILL_BACKEND"] = args.backend

    print(f"작업 디렉토리: {work_dir}")

    # PPT 생성
    from ppt_generator import run_ppt_generation
    output, vision_issues = run_ppt_generation(
        topic=args.topic,
        template_path=work_dir / "template.pptx",
        work_dir=work_dir,
        audience=args.audience,
        n_slides=args.slides,
    )

    # 결과 복사
    dest_dir = Path(args.output)
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / f"{slug}.pptx"
    shutil.copy2(output, dest)
    print(f"\n완료: {dest}")

    # AHE 진화 루프 — 게이트 경유 트리거 (#12 evolve-manual-trigger-gap)
    # 명시적 --evolve, 인라인 Vision 이슈, 또는 독립 QA 결함(qa_ok=False)을 한 경로로 통합한다.
    # human-in-loop §1: --evolve(명시적 사람 의도)만 자동 실행으로 간주하고,
    # 그 외(QA 결함 등)는 --approve-evolve로 명시 승인하지 않는 한 제안만 출력한다.
    from ahe_loop import maybe_run_evolve_loop
    qa_ok = False if vision_issues > 0 else None
    ran = maybe_run_evolve_loop(
        work_dir=work_dir,
        topic=args.topic,
        qa_ok=qa_ok,
        vision_issues=vision_issues,
        explicit=args.evolve,
        approved=args.approve_evolve,
    )
    if not ran and not (args.evolve or vision_issues > 0):
        print("\n[Auto Evolve] 이슈 없음 — Evolve 건너뜀 (다음 실행 빠름)")


if __name__ == "__main__":
    main()
