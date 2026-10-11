"""Run the unchanged v4 decisions and bind data into the frozen dashboard.

Only NOTICE_AUTOMATION is replaced. Original CSVs and preview CSS are read-only.
HWPX uses its own embedded template, not the HTML page's CSS.
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import html
import io
import json
import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

import notice_core_v4 as core

ROOT = Path(__file__).resolve().parents[1]
SOURCE_REPO = "Yu-Hyemin/culture_safe_map"
FILES = {
    "disease.csv": "disease_df.csv",
    "ili_region.csv": "ili_region_df_260923.csv",
    "performance_master.csv": "performance_master_df.csv",
}
DATA_RE = re.compile(r"const NOTICE_AUTOMATION=(\{.*?\});\n// 공문 자동화 계산일", re.S)
ELIGIBLE = {"우선 대비", "대비 필요", "일반 관리"}


def download_sources(destination: Path) -> str:
    session = requests.Session()
    session.headers["User-Agent"] = "CultureSafeMap-dashboard-build"
    session.mount("https://", HTTPAdapter(max_retries=Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])))
    # Read a single immutable commit so the three input files cannot be mixed.
    response = session.get(f"https://api.github.com/repos/{SOURCE_REPO}/commits/main", timeout=60)
    response.raise_for_status()
    commit = response.json()["sha"]
    if not re.fullmatch(r"[0-9a-f]{40}", commit):
        raise RuntimeError("원본 저장소 main 커밋을 확인하지 못했습니다.")
    for remote, local in FILES.items():
        response = session.get(f"https://raw.githubusercontent.com/{SOURCE_REPO}/{commit}/output/{remote}", timeout=60)
        response.raise_for_status()
        if not response.content:
            raise RuntimeError(f"원본 CSV가 비어 있습니다: {remote}")
        (destination / local).write_bytes(response.content)
        print(f"CSV 읽기 완료: {remote} ({len(response.content):,} bytes)")
    return commit


def notice_payload(source: Path, output: Path, as_of: str) -> dict:
    summary = pd.read_csv(output / "행사별_공문대상_요약.csv", encoding="utf-8-sig", dtype={"공연ID": str}).fillna("")
    matches = pd.read_csv(output / "행사_감염병_전체판단.csv", encoding="utf-8-sig", dtype={"공연ID": str}).fillna("")
    disease = pd.read_csv(source / FILES["disease.csv"], encoding="utf-8-sig")
    latest = disease.sort_values(["연도", "주차"]).iloc[-1]
    rows_by_id = {}
    for event_id, group in matches.groupby("공연ID", sort=False):
        chosen = group[group["최종업무상태"].isin(ELIGIBLE)].sort_values(["최종상태순위", "질병명"], ascending=[False, True]).head(3)
        rows_by_id[str(event_id)] = [{
            "name": str(row.질병명), "status": str(row.최종업무상태), "pattern": str(row.과거반복성),
            "years": int(row.비교가능연도수) if str(row.비교가능연도수) else 0,
            "highYears": int(row.높은수준연도수) if str(row.높은수준연도수) else 0,
            "recent": str(row.최근유행상태), "recentApplied": str(row.최근상태적용여부),
            "latestEnd": str(row.최신자료종료일)[:10],
        } for row in chosen.itertuples(index=False)]
    data = {"asOf": as_of, "latestYear": int(latest["연도"]), "latestWeek": int(latest["주차"]), "events": {}}
    for row in summary.itertuples(index=False):
        event_id = str(row.공연ID)
        if event_id in data["events"]:
            raise RuntimeError(f"중복 공연ID 판단 결과: {event_id}")
        data["events"][event_id] = {
            "status": str(row.최고업무상태), "diseases": str(row.대비대상감염병), "task": str(row.자동화작업),
            "recipient": str(row.담당기관), "department": str(row.담당부서), "rows": rows_by_id.get(event_id, []),
        }
    return data


def bind_payload(template: str, data: dict) -> str:
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"), allow_nan=False).replace("<", "\\u003c")
    result, count = DATA_RE.subn(lambda _: "const NOTICE_AUTOMATION=" + payload + ";\n// 공문 자동화 계산일", template, count=1)
    if count != 1:
        raise RuntimeError("고정 템플릿 공문 데이터 영역을 찾지 못했습니다.")
    # Prove that all other bytes, including every preview CSS rule, are untouched.
    if DATA_RE.sub("__DATA__", template) != DATA_RE.sub("__DATA__", result):
        raise RuntimeError("데이터 영역 외 고정 템플릿이 변경되었습니다.")
    if "downloadSelectedNoticeHwpx" not in result or "cache:'no-store'" not in result:
        raise RuntimeError("한글 내려받기 또는 페이지 접속 시 CSV 로딩 연결이 없습니다.")
    return result


def validate_hwpx_bytes(payload: bytes, *, filled: bool) -> None:
    with zipfile.ZipFile(io.BytesIO(payload)) as archive:
        if archive.testzip() is not None:
            raise RuntimeError("HWPX ZIP 검증 실패")
        mime = archive.infolist()[0]
        if mime.filename != "mimetype" or mime.compress_type != zipfile.ZIP_STORED:
            raise RuntimeError("HWPX mimetype 구조 오류")
        if archive.read("mimetype").strip() != b"application/hwp+zip":
            raise RuntimeError("HWPX 문서 형식 오류")
        for name in archive.namelist():
            if name.endswith(".xml"):
                ET.fromstring(archive.read(name))
            if filled and re.fullmatch(r"Contents/section\d+\.xml", name):
                if re.search(r"◆[A-Z0-9_]+◆", archive.read(name).decode("utf-8")):
                    raise RuntimeError("HWPX 미치환 데이터 발견")


def sample_hwpx(template: str, data: dict, output: Path) -> Path | None:
    """One test notice only; browser generates each chosen event on demand."""
    embedded = re.search(r"const NOTICE_HWPX_TEMPLATE_B64='([^']+)';", template)
    if embedded is None:
        raise RuntimeError("고정 HWPX 템플릿이 없습니다.")
    blank = base64.b64decode(embedded.group(1), validate=True)
    validate_hwpx_bytes(blank, filled=False)
    summary = pd.read_csv(output / "행사별_공문대상_요약.csv", encoding="utf-8-sig", dtype={"공연ID": str}).fillna("")
    candidates = summary[summary["최고업무상태"].isin(["우선 대비", "대비 필요"])].sort_values(["공연시작일", "공연ID"])
    if candidates.empty:
        return None
    ev = candidates.iloc[0]
    result = data["events"][str(ev["공연ID"]) ]
    def text(key, fallback="미확인"):
        return str(ev.get(key, "")).strip() or fallback
    def date(value):
        d = pd.Timestamp(value)
        return f"{d.year}. {d.month}. {d.day}."
    def number(key, unit):
        try:
            n = int(float(ev.get(key, "")))
            return f"{n:,}{unit}" if n > 0 else "미확인"
        except (TypeError, ValueError):
            return "미확인"
    start, end = date(ev["공연시작일"]), date(ev["공연종료일"])
    values = {
        "EVENT_ID": text("공연ID"), "EVENT_NAME": text("공연명"), "RECIPIENT": result["recipient"] or "수신기관 미확인",
        "DEPARTMENT": result["department"] or "담당부서 미확인", "EMAIL": "담당자 이메일 미입력", "START_DATE_KO": start,
        "PERIOD": start + ("∼" + end if start != end else ""), "STATUS": result["status"], "VENUE": text("공연장명"),
        "REGION": text("지역"), "GENRE": text("공연장르"), "INDOOR": text("실내야외"), "SEATS": number("공연장좌석수", "석"),
        "EXPOSURE": number("행사노출단계", "단계"), "TASK": result["task"] or "담당자 검토", "REVIEW_STATUS": "검토대기",
        "DISEASE_NAMES": ", ".join(r["name"] for r in result["rows"]) or "검토 대상 질병 없음",
        "LATEST_INFO": f"{data['asOf']} · {data['latestYear']}년 {data['latestWeek']}주차",
    }
    for i in range(3):
        row = result["rows"][i] if i < len(result["rows"]) else {}
        n = i + 1
        values.update({
            f"D{n}_NAME": row.get("name", ""), f"D{n}_PATTERN": row.get("pattern", ""),
            f"D{n}_RECENT": row.get("recent", "") if row.get("recentApplied") == "적용" else ("행사 직전 재조회" if row else ""),
            f"D{n}_STATUS": row.get("status", ""),
            f"D{n}_EVIDENCE": f"{row.get('years', 0)}개년 중 {row.get('highYears', 0)}개년 높은 주차" if row else "",
        })
    buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blank)) as original, zipfile.ZipFile(buffer, "w") as generated:
        for entry in original.infolist():
            payload = original.read(entry.filename)
            if re.fullmatch(r"Contents/section\d+\.xml", entry.filename):
                xml = payload.decode("utf-8")
                for key, value in values.items():
                    xml = xml.replace(f"◆{key}◆", html.escape(str(value), quote=True))
                if re.search(r"◆[A-Z0-9_]+◆", xml):
                    raise RuntimeError("한글 템플릿 데이터 키가 누락됐습니다.")
                payload = xml.encode("utf-8")
            generated.writestr(entry.filename, payload, compress_type=zipfile.ZIP_STORED if entry.filename == "mimetype" else zipfile.ZIP_DEFLATED)
    validate_hwpx_bytes(buffer.getvalue(), filled=True)
    path = output / "sample_notice.hwpx"
    path.write_bytes(buffer.getvalue())
    return path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-dir", type=Path, help="Read-only cached inputs for offline verification")
    parser.add_argument("--source-commit", default="")
    parser.add_argument("--work-dir", type=Path, default=ROOT / ".culture-safe-map-build")
    parser.add_argument("--as-of", help="Reproducible verification date; production defaults to Korean date")
    args = parser.parse_args()
    source = args.work_dir / "source"
    output = args.work_dir / "output"
    source.mkdir(parents=True, exist_ok=True)
    output.mkdir(parents=True, exist_ok=True)
    if args.source_dir:
        for local in FILES.values():
            shutil.copy2(args.source_dir / local, source / local)
        source_commit = args.source_commit or "offline-verification"
    else:
        source_commit = download_sources(source)
    for local in FILES.values():
        if pd.read_csv(source / local, encoding="utf-8-sig", nrows=1).empty:
            raise RuntimeError(f"데이터가 없는 입력 CSV: {local}")
    core.BASE_DIR, core.SOURCE_DIR, core.OUTPUT_DIR = args.work_dir, source, output
    core.DOCS_DIR = output / "공문예시"
    core.MAX_DOCS = 0
    core.AS_OF_DATE = pd.Timestamp(args.as_of).normalize() if args.as_of else pd.Timestamp.now(tz="Asia/Seoul").tz_localize(None).normalize()
    core.main()
    as_of = core.AS_OF_DATE.strftime("%Y-%m-%d")
    data = notice_payload(source, output, as_of)
    template = (ROOT / "automation/templates/dashboard.html").read_text(encoding="utf-8")
    page = bind_payload(template, data)
    generated_html = output / "index.html"
    generated_html.write_text(page, encoding="utf-8")
    example = sample_hwpx(template, data, output)
    (output / "notice_payload.json").write_text(json.dumps(data, ensure_ascii=False, separators=(",", ":"), allow_nan=False), encoding="utf-8")
    report = {
        "sourceRepository": SOURCE_REPO, "sourceBranch": "main", "sourceCommit": source_commit,
        "calculatedAtKorea": as_of, "latestDiseaseYear": data["latestYear"], "latestDiseaseWeek": data["latestWeek"],
        "noticeEvents": len(data["events"]), "htmlBytes": generated_html.stat().st_size,
        "templateSha256": hashlib.sha256(template.encode("utf-8")).hexdigest(),
        "sampleHwpx": example.name if example else None, "previewCssChanged": False,
        "sourceSha256": {remote: hashlib.sha256((source / local).read_bytes()).hexdigest() for remote, local in FILES.items()},
        "limitations": ["HWPX 문서는 한글에서 최종 글꼴·쪽 배치 확인 필요", "전자결재·메일 실제 발송 없음"],
    }
    (output / "build_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    # Replace the deploy file only after every build check succeeds.
    pending = ROOT / "index.html.pending"
    pending.write_text(page, encoding="utf-8")
    pending.replace(ROOT / "index.html")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
