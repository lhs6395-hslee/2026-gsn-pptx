#!/usr/bin/env python3
"""
AHE Experience Observability 도구
원시 트레이스를 Evolve Agent가 읽을 수 있는 digest로 압축
"""
import argparse, json, datetime
from pathlib import Path


def distill(run_id: str, traces_dir: str, output_dir: str) -> dict:
    traces_path = Path(traces_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    traces = [
        json.loads(p.read_text())
        for p in sorted(traces_path.glob(f"{run_id}_*.json"))
    ]
    if not traces:
        print(f"No traces found for run_id={run_id}")
        return {}

    failures = [t for t in traces if t.get("status") == "FAIL"]
    passes   = [t for t in traces if t.get("status") == "PASS"]

    # 실패 패턴 집계
    fp: dict[str, int] = {}
    for t in failures:
        for issue in t.get("issues", []):
            key = issue.get("type", "unknown")
            fp[key] = fp.get(key, 0) + 1

    # 컴포넌트별 귀인
    ca = {
        "middleware": [t["slide_id"] for t in failures
                       if any(i["type"] in ("xml_parse_error","unescaped_ampersand","unescaped_lt")
                              for i in t.get("issues",[]))],
        "verifier_rules": [t["slide_id"] for t in failures
                           if any(i["type"] == "placeholder_remaining"
                                  for i in t.get("issues",[]))],
        "tools": [t["slide_id"] for t in failures
                  if any(i["type"] == "visual_overflow"
                         for i in t.get("issues",[]))],
        "system_prompt": [t["slide_id"] for t in failures
                          if any(i["type"] in ("file_not_found","plan_mismatch")
                                 for i in t.get("issues",[]))],
    }

    digest = {
        "run_id":   run_id,
        "created":  datetime.datetime.utcnow().isoformat(),
        "summary": {
            "total":     len(traces),
            "pass":      len(passes),
            "fail":      len(failures),
            "pass_rate": round(len(passes) / len(traces), 3) if traces else 0,
        },
        "failure_patterns":      fp,
        "component_attribution": {k: v for k, v in ca.items() if v},
    }

    out = output_path / f"{run_id}_digest.json"
    out.write_text(json.dumps(digest, ensure_ascii=False, indent=2))
    print(f"Digest written: {out}")
    print(f"pass_rate: {digest['summary']['pass_rate']:.1%}  "
          f"({digest['summary']['pass']}/{digest['summary']['total']})")
    return digest


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--run-id",     required=True)
    ap.add_argument("--traces-dir", default=str(Path.home() / ".ppt-skill/traces"))
    ap.add_argument("--output-dir", default=str(Path.home() / ".ppt-skill/evolution"))
    args = ap.parse_args()
    distill(args.run_id, args.traces_dir, args.output_dir)
