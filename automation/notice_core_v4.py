from __future__ import annotations

import json
import hashlib
import math
import os
import re
from pathlib import Path

import numpy as np
import pandas as pd
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor


BASE_DIR = Path("/content/문화행사_감염병_자동화")
SOURCE_DIR = BASE_DIR / "source"
OUTPUT_DIR = BASE_DIR / "output"
DOCS_DIR = OUTPUT_DIR / "공문예시"
AS_OF_DATE = pd.Timestamp.now(tz='Asia/Seoul').tz_localize(None).normalize()
FUTURE_DAYS = 90
MAX_DOCS = 0
FONT = "NanumGothic"

KDCA_ARI_URL = "https://dportal.kdca.go.kr/pot/is/st/ari.do"
KDCA_INFLU_URL = "https://dportal.kdca.go.kr/pot/is/st/influ.do"
KDCA_PREVENT_URL = "https://health.kdca.go.kr/healthinfo/biz/health/gnrlzHealthInfo/gnrlzHealthInfo/gnrlzHealthInfoView.do?cntnts_sn=6680"

ALLOWED_DISEASES = [
    "마이코플라즈마균",
    "클라미디아균",
    "아데노바이러스",
    "사람보카바이러스",
    "파라인플루엔자바이러스",
    "호흡기세포융합바이러스",
    "리노바이러스",
    "사람메타뉴모바이러스",
    "사람코로나바이러스",
    "인플루엔자바이러스",
    "코로나19",
]

STATUS_RANK = {"해당 없음": 0, "일반 관리": 1, "대비 필요": 2, "우선 대비": 3}
RANK_STATUS = {v: k for k, v in STATUS_RANK.items()}

PERFORMANCE_REQUIRED = [
    "공연ID", "공연명", "공연장르", "공연시작일", "공연종료일", "공연시설주소", "공연장명",
]
PERFORMANCE_CURRENT = [
    "공연ID", "공연시설ID", "공연장ID", "공연명", "공연장르", "공연시작일", "공연종료일",
    "공연시간", "공연시간_분", "공연일정", "공연시설위도", "공연시설경도", "공연시설주소",
    "공연장명", "공연장좌석수", "지역_시도", "지역_구군", "좌석단계", "시간단계", "행사노출단계",
]
PERFORMANCE_OPTIONAL = [
    "실내야외", "예상참가자수", "담당기관", "담당부서", "담당자", "담당자이메일",
    "데이터출처", "확인상태", "비고",
]


def portal_week(date_value: pd.Timestamp) -> int:
    date_value = pd.Timestamp(date_value).normalize()
    jan1 = pd.Timestamp(year=date_value.year, month=1, day=1)
    days_since_sunday = (jan1.weekday() + 1) % 7
    week1_start = jan1 - pd.Timedelta(days=days_since_sunday)
    return int(((date_value - week1_start).days // 7) + 1)


def clean_text(value, fallback="미확인"):
    if pd.isna(value):
        return fallback
    value = str(value).strip()
    return value if value else fallback


def safe_int(value):
    if pd.isna(value):
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def first_existing(*names):
    for name in names:
        path = SOURCE_DIR / name
        if path.exists():
            return path
    raise FileNotFoundError(f"입력 파일이 없습니다: {', '.join(names)}")


def optional_existing(*names):
    for name in names:
        path = SOURCE_DIR / name
        if path.exists():
            return path
    return None


def file_fingerprint(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load_inputs():
    performance_path = first_existing("performance_master_df_260923.csv", "performance_master_df.csv")
    disease_path = first_existing("disease_df_260923.csv", "disease_df.csv")
    risk_path = optional_existing("risk_df_260923.csv", "risk_df.csv")
    ili_path = SOURCE_DIR / "ili_region_df_260923.csv"

    performance = pd.read_csv(performance_path, encoding="utf-8-sig")
    disease = pd.read_csv(disease_path, encoding="utf-8-sig")
    risk_reference = pd.read_csv(risk_path, encoding="utf-8-sig") if risk_path else pd.DataFrame()
    ili_region = pd.read_csv(ili_path, encoding="utf-8-sig") if ili_path.exists() else pd.DataFrame()

    for col in ["공연시작일", "공연종료일"]:
        performance[col] = pd.to_datetime(performance[col], errors="coerce")

    for col in ["주차시작일", "주차종료일"]:
        disease[col] = pd.to_datetime(disease[col], errors="coerce")

    for col in ["공연장좌석수", "공연시간_분", "공연시설위도", "공연시설경도", "예상참가자수", "좌석단계", "시간단계", "행사노출단계"]:
        if col not in performance.columns:
            performance[col] = np.nan
        performance[col] = pd.to_numeric(performance[col], errors="coerce")

    for col in PERFORMANCE_OPTIONAL:
        if col not in performance.columns:
            performance[col] = np.nan

    disease["연도"] = pd.to_numeric(disease["연도"], errors="coerce").astype("Int64")
    disease["주차"] = pd.to_numeric(disease["주차"], errors="coerce").astype("Int64")
    disease["환자수"] = pd.to_numeric(disease["환자수"], errors="coerce")

    source_info = {
        "performance": {"파일명": performance_path.name, "sha256": file_fingerprint(performance_path)},
        "disease": {"파일명": disease_path.name, "sha256": file_fingerprint(disease_path)},
        "risk_reference": {"파일명": risk_path.name, "sha256": file_fingerprint(risk_path)} if risk_path else None,
        "ili_region": {"파일명": ili_path.name, "sha256": file_fingerprint(ili_path)} if ili_path.exists() else None,
    }
    return performance, disease, risk_reference, ili_region, source_info


def seat_stage(value):
    value = safe_int(value)
    if value is None or value <= 0:
        return np.nan
    if value < 300:
        return 1
    if value < 500:
        return 2
    if value < 1000:
        return 3
    return 4


def time_stage(value):
    value = safe_int(value)
    if value is None or value <= 0:
        return np.nan
    if value < 60:
        return 1
    if value < 120:
        return 2
    if value < 180:
        return 3
    return 4


def exposure_stage(seat_value, time_value):
    if pd.isna(seat_value) or pd.isna(time_value):
        return np.nan
    return int((int(seat_value) + int(time_value)) // 2)


def normalize_performance(performance):
    out = performance.copy()
    missing_required = [c for c in PERFORMANCE_REQUIRED if c not in out.columns]
    if missing_required:
        raise ValueError(f"공연 데이터 필수 컬럼 누락: {missing_required}")
    for col in PERFORMANCE_CURRENT + PERFORMANCE_OPTIONAL:
        if col not in out.columns:
            out[col] = np.nan
    out["지역_시도"] = out["지역_시도"].fillna(
        out["공연시설주소"].fillna("").astype(str).str.split().str[0]
        .replace({"서울특별시": "서울", "부산광역시": "부산", "대구광역시": "대구", "인천광역시": "인천", "광주광역시": "광주", "대전광역시": "대전", "울산광역시": "울산", "세종특별자치시": "세종", "제주특별자치도": "제주", "경기도": "경기", "강원특별자치도": "강원", "충청북도": "충북", "충청남도": "충남", "전북특별자치도": "전북", "전라남도": "전남", "경상북도": "경북", "경상남도": "경남"})
    )
    computed_seat = out["공연장좌석수"].apply(seat_stage)
    computed_time = out["공연시간_분"].apply(time_stage)
    out["좌석단계"] = out["좌석단계"].where(out["좌석단계"].between(1, 4), computed_seat)
    out["시간단계"] = out["시간단계"].where(out["시간단계"].between(1, 4), computed_time)
    computed_exposure = [exposure_stage(s, t) for s, t in zip(out["좌석단계"], out["시간단계"])]
    out["행사노출단계"] = out["행사노출단계"].where(out["행사노출단계"].between(1, 4), computed_exposure)
    out["실내야외"] = out["실내야외"].fillna("미확인")
    out["데이터출처"] = out["데이터출처"].fillna("KOPIS")
    out["확인상태"] = out["확인상태"].fillna("미확인")
    return out


def build_validation_issues(performance):
    issues = []
    def add(mask, code, severity, field, message):
        for idx in performance.index[mask]:
            issues.append({"행번호": int(idx) + 2, "공연ID": clean_text(performance.at[idx, "공연ID"], ""), "심각도": severity, "코드": code, "컬럼": field, "내용": message})
    add(performance["공연ID"].isna() | performance["공연ID"].astype(str).str.strip().eq(""), "E001", "오류", "공연ID", "필수값 누락")
    add(performance["공연ID"].duplicated(keep=False), "E002", "오류", "공연ID", "중복 공연ID")
    add(performance["공연시작일"].isna(), "E003", "오류", "공연시작일", "날짜 형식 오류 또는 누락")
    add(performance["공연종료일"].isna(), "E004", "오류", "공연종료일", "날짜 형식 오류 또는 누락")
    add(performance["공연종료일"] < performance["공연시작일"], "E005", "오류", "공연종료일", "종료일이 시작일보다 빠름")
    add(performance["공연장좌석수"].notna() & (performance["공연장좌석수"] < 0), "E006", "오류", "공연장좌석수", "음수 입력")
    add(performance["공연시간_분"].notna() & (performance["공연시간_분"] < 0), "E007", "오류", "공연시간_분", "음수 입력")
    add(performance["공연시설주소"].isna() | performance["공연시설주소"].astype(str).str.strip().eq(""), "W001", "경고", "공연시설주소", "주소 누락")
    add(performance["행사노출단계"].isna(), "W002", "경고", "행사노출단계", "좌석수 또는 공연시간 부족으로 계산 불가")
    return pd.DataFrame(issues, columns=["행번호", "공연ID", "심각도", "코드", "컬럼", "내용"])


def validate_inputs(performance, disease, risk_reference, ili_region, source_info):
    disease_set = set(disease["질병명"].dropna().unique())
    missing = sorted(set(ALLOWED_DISEASES) - disease_set)
    extra = sorted(disease_set - set(ALLOWED_DISEASES))

    issues = build_validation_issues(performance)
    shared_ids = set(performance["공연ID"].dropna()) & set(risk_reference.get("공연ID", pd.Series(dtype=str)).dropna())
    stage_mismatch = 0
    if shared_ids and all(c in risk_reference.columns for c in ["공연ID", "좌석단계", "시간단계", "행사노출단계"]):
        left = performance[performance["공연ID"].isin(shared_ids)][["공연ID", "좌석단계", "시간단계", "행사노출단계"]]
        right = risk_reference[risk_reference["공연ID"].isin(shared_ids)][["공연ID", "좌석단계", "시간단계", "행사노출단계"]].drop_duplicates("공연ID")
        check = left.merge(right, on="공연ID", suffixes=("_행사", "_위험"))
        stage_mismatch = int((
            check["좌석단계_행사"].fillna(-1).ne(check["좌석단계_위험"].fillna(-1))
            | check["시간단계_행사"].fillna(-1).ne(check["시간단계_위험"].fillna(-1))
            | check["행사노출단계_행사"].fillna(-1).ne(check["행사노출단계_위험"].fillna(-1))
        ).sum())

    report = {
        "기준일": AS_OF_DATE.strftime("%Y-%m-%d"),
        "원본파일": source_info,
        "performance_master_df": {
            "행수": int(len(performance)),
            "정규화후열수": int(len(performance.columns)),
            "지도원본열수": 20,
            "전체중복": int(performance.duplicated().sum()),
            "공연ID중복": int(performance.duplicated(subset=["공연ID"]).sum()),
            "공연시간결측": int(performance["공연시간_분"].isna().sum()),
            "좌석수0또는결측": int(((performance["공연장좌석수"].isna()) | (performance["공연장좌석수"] <= 0)).sum()),
            "실제컬럼수": int(len([c for c in PERFORMANCE_CURRENT if c in performance.columns])),
            "검사오류": int((issues["심각도"] == "오류").sum()) if len(issues) else 0,
            "검사경고": int((issues["심각도"] == "경고").sum()) if len(issues) else 0,
        },
        "disease_df": {
            "행수": int(len(disease)),
            "열수": int(len(disease.columns)),
            "질병수": int(disease["질병명"].nunique()),
            "최소연도": int(disease["연도"].min()),
            "최대연도": int(disease["연도"].max()),
            "2026최신주차": int(disease.loc[disease["연도"] == 2026, "주차"].max()),
            "누락질병": missing,
            "추가질병": extra,
            "전체중복": int(disease.duplicated().sum()),
        },
        "risk_df_참고용": {
            "행수": int(len(risk_reference)),
            "열수": int(len(risk_reference.columns)),
            "자동화직접입력": False,
            "사유": "주차별 대표 질병 1개만 연결된 지도 대시보드용 결과",
            "공통공연ID수": int(len(shared_ids)),
            "공통ID_단계불일치수": stage_mismatch,
        },
        "ili_region_df_참고용": {"행수": int(len(ili_region)), "자동화직접입력": False, "사유": "지역 묶음 단위 참고지표이며 개인 공문 우선순위에는 직접 가산하지 않음"},
    }

    if missing:
        raise ValueError(f"자동화 대상 질병 누락: {missing}")

    if len(issues) and (issues["심각도"] == "오류").any():
        fatal = issues[issues["심각도"] == "오류"].head(10).to_dict("records")
        raise ValueError(f"공연 입력 오류가 있어 실행을 중단합니다: {fatal}")

    return report, issues


def build_pattern_db(disease):
    hist = disease[
        (disease["연도"] >= 2016)
        & (disease["연도"] <= 2025)
        & (disease["질병명"].isin(ALLOWED_DISEASES))
    ].copy()

    # 코로나19는 실제 확보된 2024~2025만 남는다.
    hist = hist[(hist["질병명"] != "코로나19") | (hist["연도"] >= 2024)].copy()

    hist["연도내75백분위"] = hist.groupby(["질병명", "연도"])["환자수"].transform(
        lambda s: s.quantile(0.75)
    )
    hist["연도내높은주차"] = hist["환자수"] >= hist["연도내75백분위"]
    hist["연도중앙값"] = hist.groupby(["질병명", "연도"])["환자수"].transform("median")
    hist["연도내계절지수"] = np.where(
        hist["연도중앙값"] > 0,
        hist["환자수"] / hist["연도중앙값"],
        np.nan,
    )

    grouped = hist.groupby(["질병명", "주차"], as_index=False)
    pattern = grouped.agg(
        비교가능연도수=("연도", "nunique"),
        높은수준연도수=("연도내높은주차", "sum"),
        과거평균환자수=("환자수", "mean"),
        과거중앙값=("환자수", "median"),
        과거75백분위=("환자수", lambda s: s.quantile(0.75)),
        과거90백분위=("환자수", lambda s: s.quantile(0.90)),
        계절지수중앙값=("연도내계절지수", "median"),
    )

    pattern["반복비율"] = pattern["높은수준연도수"] / pattern["비교가능연도수"]

    def repeat_label(row):
        if row["비교가능연도수"] < 3:
            return "자료기간 짧음"
        if row["반복비율"] >= 0.60:
            return "높음"
        if row["반복비율"] >= 0.40:
            return "보통"
        return "낮음"

    pattern["과거반복성"] = pattern.apply(repeat_label, axis=1)
    pattern["자료기간"] = np.where(
        pattern["질병명"].eq("코로나19"),
        "2024~2025 (참고용)",
        "2016~2025",
    )
    pattern["공식출처"] = np.where(
        pattern["질병명"].eq("인플루엔자바이러스"),
        KDCA_INFLU_URL,
        KDCA_ARI_URL,
    )

    numeric_cols = [
        "과거평균환자수", "과거중앙값", "과거75백분위", "과거90백분위",
        "계절지수중앙값", "반복비율",
    ]
    pattern[numeric_cols] = pattern[numeric_cols].round(3)

    return pattern.sort_values(["질병명", "주차"]).reset_index(drop=True)


def percentile_strict(series, value):
    s = pd.to_numeric(series, errors="coerce").dropna()
    if len(s) == 0 or pd.isna(value):
        return np.nan
    return float((s < value).sum() / len(s) * 100)


def build_recent_status(disease):
    recent_source = disease[(disease["연도"] == 2026) & disease["질병명"].isin(ALLOWED_DISEASES)].copy()
    hist_source = disease[(disease["연도"] >= 2016) & (disease["연도"] <= 2025)].copy()

    rows = []
    for disease_name in ALLOWED_DISEASES:
        d = recent_source[recent_source["질병명"] == disease_name].sort_values("주차").tail(4)
        if d.empty:
            continue

        values = d["환자수"].astype(float).tolist()
        latest = float(values[-1])
        previous_mean = float(np.mean(values[:-1])) if len(values) > 1 else np.nan
        change_rate = ((latest / previous_mean) - 1) * 100 if previous_mean and previous_mean > 0 else np.nan
        slope = float(np.polyfit(np.arange(len(values)), values, 1)[0]) if len(values) >= 2 else 0.0

        baseline = hist_source[hist_source["질병명"] == disease_name]
        if disease_name == "코로나19":
            baseline = baseline[baseline["연도"] >= 2024]
        pct = percentile_strict(baseline["환자수"], latest)

        if not pd.isna(pct) and pct >= 75 and slope > 0:
            status = "증가·높음"
        elif not pd.isna(pct) and pct >= 90:
            status = "높음"
        elif slope > 0:
            status = "증가 관찰"
        else:
            status = "안정"

        rows.append({
            "질병명": disease_name,
            "최근4주": " → ".join(str(int(v)) for v in values),
            "최근4주시작주차": int(d["주차"].min()),
            "최신주차": int(d["주차"].max()),
            "최신자료종료일": d["주차종료일"].max(),
            "최신환자수": int(latest),
            "직전3주평균": round(previous_mean, 2) if not pd.isna(previous_mean) else np.nan,
            "최근증감률": round(change_rate, 2) if not pd.isna(change_rate) else np.nan,
            "4주추세기울기": round(slope, 2),
            "과거기준백분위": round(pct, 2) if not pd.isna(pct) else np.nan,
            "최근유행상태": status,
            "공식출처": KDCA_INFLU_URL if disease_name == "인플루엔자바이러스" else KDCA_ARI_URL,
        })

    return pd.DataFrame(rows)


def event_stage(row):
    exposure = safe_int(row.get("행사노출단계"))
    if exposure is not None and exposure >= 3:
        return 2
    if exposure == 2:
        return 1
    return 0


def event_scale_label(seats):
    seats = safe_int(seats)
    if seats is None or seats <= 0:
        return "미확인"
    if seats < 300:
        return "소규모"
    if seats < 500:
        return "중소규모"
    if seats < 1000:
        return "중규모"
    return "대규모"


def build_matches(performance, pattern, recent):
    events = performance[
        performance["공연시작일"].between(
            AS_OF_DATE,
            AS_OF_DATE + pd.Timedelta(days=FUTURE_DAYS),
            inclusive="both",
        )
    ].copy()

    events = events.drop_duplicates(subset=["공연ID"]).copy()
    events["행사주차"] = events["공연시작일"].apply(portal_week)
    events["기준일까지남은일수"] = (events["공연시작일"] - AS_OF_DATE).dt.days
    events["행사보조단계"] = events.apply(event_stage, axis=1)
    events["좌석규모"] = events["공연장좌석수"].apply(event_scale_label)
    events["지역"] = events["지역_시도"].fillna("미확인").astype(str) + " " + events["지역_구군"].fillna("").astype(str)
    events["지역"] = events["지역"].str.strip().replace("", "미확인")

    event_cols = [
        "공연ID", "공연명", "공연장르", "공연시작일", "공연종료일", "지역", "지역_시도", "지역_구군", "공연시설주소",
        "공연장명", "공연장좌석수", "좌석규모", "공연시간_분", "공연일정", "좌석단계", "시간단계", "행사노출단계",
        "실내야외", "예상참가자수", "담당기관", "담당부서", "담당자", "담당자이메일", "데이터출처", "확인상태", "비고",
        "행사주차", "기준일까지남은일수", "행사보조단계",
    ]
    event_base = events[event_cols].copy()

    disease_weeks = pattern[[
        "질병명", "주차", "비교가능연도수", "높은수준연도수", "반복비율", "과거반복성",
        "과거평균환자수", "과거중앙값", "과거75백분위", "과거90백분위", "계절지수중앙값",
        "자료기간", "공식출처",
    ]].rename(columns={"주차": "행사주차", "공식출처": "과거자료출처"})

    matches = event_base.merge(disease_weeks, on="행사주차", how="left")
    matches = matches.merge(
        recent[[
            "질병명", "최근4주", "최신주차", "최신자료종료일", "최근증감률",
            "과거기준백분위", "최근유행상태", "공식출처",
        ]].rename(columns={"공식출처": "최근자료출처"}),
        on="질병명",
        how="left",
    )

    def pattern_points(row):
        if row["비교가능연도수"] < 3:
            return 0
        if row["과거반복성"] == "높음":
            return 2
        if row["과거반복성"] == "보통":
            return 1
        return 0

    def recent_points(row):
        if row["기준일까지남은일수"] > 30:
            return 0
        if row["최근유행상태"] in ["증가·높음", "높음"]:
            return 2
        if row["최근유행상태"] == "증가 관찰":
            return 1
        return 0

    matches["과거패턴점수"] = matches.apply(pattern_points, axis=1)
    matches["최근상태적용여부"] = np.where(matches["기준일까지남은일수"] <= 30, "적용", "행사 직전 재조회 필요")
    matches["최근유행점수"] = matches.apply(recent_points, axis=1)
    matches["감염병신호점수"] = matches["과거패턴점수"] + matches["최근유행점수"]

    def final_rank(row):
        signal = int(row["감염병신호점수"])
        if signal >= 3:
            rank = 3
        elif signal == 2:
            rank = 2
        elif signal == 1:
            rank = 1
        else:
            rank = 0
        if signal > 0 and int(row["행사보조단계"]) == 2:
            rank = min(3, rank + 1)
        return rank

    matches["최종상태순위"] = matches.apply(final_rank, axis=1)
    matches["최종업무상태"] = matches["최종상태순위"].map(RANK_STATUS)

    def evidence(row):
        n = safe_int(row["비교가능연도수"]) or 0
        m = safe_int(row["높은수준연도수"]) or 0
        seats = safe_int(row["공연장좌석수"])
        runtime = safe_int(row["공연시간_분"])
        seat_text = f"{seats:,}석" if seats and seats > 0 else "미확인"
        time_text = f"{runtime}분" if runtime and runtime > 0 else "미확인"
        exposure_text = safe_int(row.get("행사노출단계"))
        exposure_text = f"{exposure_text}단계" if exposure_text else "계산 불가"
        recent_text = row["최근유행상태"] if row["최근상태적용여부"] == "적용" else "행사 직전 재조회 필요"
        text = (
            f"과거 비교 가능한 {n}개 연도 중 {m}개 연도에서 행사 주차가 연도 내 높은 수준으로 확인되어 "
            f"반복성은 {row['과거반복성']}입니다. 최근 4주 상태는 {recent_text}이며, "
            f"공연장 좌석수는 {seat_text}, 공연시간은 {time_text}, 행사노출단계는 {exposure_text}로 확인되었습니다. "
            f"최종 업무상태는 {row['최종업무상태']}입니다."
        )
        if row["질병명"] == "코로나19":
            text += " 코로나19는 비교 가능한 과거 자료가 2024~2025년으로 짧아 장기 계절성 판단이 아닌 참고용 패턴으로 사용했습니다."
        return text

    matches["판단사유"] = matches.apply(evidence, axis=1)

    output_cols = [
        "공연ID", "공연명", "공연장르", "공연시작일", "공연종료일", "지역", "지역_시도", "지역_구군", "공연시설주소", "공연장명",
        "공연장좌석수", "좌석규모", "공연시간_분", "좌석단계", "시간단계", "행사노출단계", "실내야외", "예상참가자수",
        "담당기관", "담당부서", "담당자", "담당자이메일", "데이터출처", "확인상태", "비고",
        "행사주차", "기준일까지남은일수", "질병명", "자료기간", "비교가능연도수",
        "높은수준연도수", "반복비율", "과거반복성", "최근4주", "최신주차",
        "최신자료종료일", "최근증감률", "과거기준백분위", "최근유행상태",
        "최근상태적용여부", "과거패턴점수", "최근유행점수", "행사보조단계",
        "최종상태순위", "최종업무상태", "판단사유", "과거자료출처", "최근자료출처",
    ]
    return matches[output_cols].sort_values(
        ["공연시작일", "공연ID", "최종상태순위", "질병명"],
        ascending=[True, True, False, True],
    ).reset_index(drop=True)


def build_event_summary(matches):
    target = matches[matches["최종업무상태"].isin(["우선 대비", "대비 필요"])].copy()

    def join_names(g):
        ordered = g.sort_values(["최종상태순위", "질병명"], ascending=[False, True])
        return ", ".join(ordered["질병명"].astype(str).tolist())

    rows = []
    for event_id, group in matches.groupby("공연ID", sort=False):
        base = group.iloc[0]
        tg = target[target["공연ID"] == event_id]
        highest = group["최종상태순위"].max()
        days = int(base["기준일까지남은일수"])
        if days <= 3:
            task = "D-3 최신자료 재확인 및 공문 검토"
        elif days <= 14:
            task = "D-14 공문 초안 생성·검토"
        elif days <= 30:
            task = "D-30 패턴 기반 후보 사전확인"
        else:
            task = "일정 대기"
        rows.append({
            "공연ID": event_id,
            "공연명": base["공연명"],
            "공연장르": base["공연장르"],
            "공연시작일": base["공연시작일"],
            "공연종료일": base["공연종료일"],
            "지역": base["지역"],
            "공연장명": base["공연장명"],
            "공연장좌석수": base["공연장좌석수"],
            "행사노출단계": base["행사노출단계"],
            "실내야외": base["실내야외"],
            "담당기관": base["담당기관"],
            "담당부서": base["담당부서"],
            "담당자이메일": base["담당자이메일"],
            "기준일까지남은일수": days,
            "최고업무상태": RANK_STATUS[int(highest)],
            "대비대상감염병": join_names(tg) if len(tg) else "없음",
            "자동화작업": task,
            "검토상태": "검토대기",
        })
    return pd.DataFrame(rows).sort_values(["공연시작일", "공연ID"]).reset_index(drop=True)


def build_prevention_db():
    common_participant = "흐르는 물에 비누로 손 씻기; 기침예절 준수; 호흡기 증상이 있을 때 마스크 착용 및 참여 자제 안내"
    common_operator = "환기 가능 여부와 혼잡구간 사전 점검; 손 위생 안내물 비치; 운영인력 예방수칙 및 내부 연락절차 공유"
    rows = []
    for disease_name in ALLOWED_DISEASES:
        extra = ""
        if disease_name in ["인플루엔자바이러스", "코로나19", "호흡기세포융합바이러스"]:
            extra = "공식 예방접종 대상과 최신 사업기준은 질병관리청 자료를 별도 확인"
        rows.append({
            "감염병명": disease_name,
            "감염병유형": "호흡기·비말 관련",
            "참가자안내사항": common_participant,
            "행사운영자조치": common_operator,
            "추가확인사항": extra or "질병별 추가 조치는 공식자료 확인 후 적용",
            "공식출처": KDCA_PREVENT_URL,
        })
    return pd.DataFrame(rows)


def build_send_log(matches, selected_event_ids):
    selected = matches[
        matches["공연ID"].isin(selected_event_ids)
        & matches["최종업무상태"].isin(["우선 대비", "대비 필요"])
    ].copy()
    if selected.empty:
        return pd.DataFrame(columns=[
            "중복키", "행사ID", "행사명", "감염병명", "패턴판단결과", "최근유행상태",
            "공문생성일", "검토상태", "승인일", "발송일", "수신기관", "수신이메일", "발송상태",
        ])
    out = pd.DataFrame({
        "행사ID": selected["공연ID"],
        "행사명": selected["공연명"],
        "감염병명": selected["질병명"],
        "패턴판단결과": selected["최종업무상태"],
        "최근유행상태": selected["최근유행상태"],
        "공문생성일": AS_OF_DATE.strftime("%Y-%m-%d"),
        "검토상태": "검토대기",
        "승인일": "",
        "발송일": "",
        "수신기관": selected["담당기관"].apply(lambda v: clean_text(v, "수신기관 미확인")),
        "수신이메일": selected["담당자이메일"].apply(lambda v: clean_text(v, "")),
        "발송상태": "미발송",
    })
    out.insert(
        0,
        "중복키",
        out["행사ID"].astype(str) + "|" + out["감염병명"].astype(str) + "|예방협조|" + AS_OF_DATE.strftime("%Y%m%d"),
    )
    return out.drop_duplicates(subset=["중복키"]).reset_index(drop=True)


# ---------- DOCX helpers ----------

def set_run_font(run, size=10.5, bold=False, color="000000"):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{attr}"), FONT)


def set_cell_shading(cell, fill="F2F2F2"):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_borders(cell, color="777777", size=4):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "bottom", "start", "end"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), str(size))
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)


def write_cell(cell, text, size=8.5, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, shade=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.05
    r = p.add_run(str(text))
    set_run_font(r, size=size, bold=bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_borders(cell)
    if shade:
        set_cell_shading(cell)


def add_para(doc, text="", size=10.5, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=0, after=0, line=1.35):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = line
    r = p.add_run(text)
    set_run_font(r, size=size, bold=bold)
    return p


def safe_filename(text):
    text = re.sub(r"[\\/:*?\"<>|]", "_", str(text))
    return re.sub(r"\s+", "_", text).strip("_")[:70]


def create_official_doc(event_row, disease_rows, prevention_db, index):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    sec.top_margin = Mm(13)
    sec.bottom_margin = Mm(13)
    sec.left_margin = Mm(18)
    sec.right_margin = Mm(18)

    header = sec.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hr = hp.add_run("시연용 자동생성 초안 · 외부발송 금지")
    set_run_font(hr, size=8, bold=True, color="444444")

    add_para(doc, "○ ○ 시  감 염 병 관 리 과", size=22, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, before=2, after=10, line=1.0)

    top = doc.add_table(rows=3, cols=2)
    top.alignment = WD_TABLE_ALIGNMENT.CENTER
    labels = ["수신", "(경유)", "제목"]
    values = [
        clean_text(event_row.get("담당기관"), "수신기관 미확인"),
        clean_text(event_row.get("담당부서"), "담당부서 미확인"),
        f"{event_row['공연명']} 감염병 예방관리 협조 요청",
    ]
    for i in range(3):
        write_cell(top.cell(i, 0), labels[i], size=9.5, bold=(i != 1), align=WD_ALIGN_PARAGRAPH.CENTER, shade=(i == 2))
        write_cell(top.cell(i, 1), values[i], size=9.5 if i < 2 else 10.5, bold=(i == 2), shade=(i == 2))

    disease_names = ", ".join(disease_rows["질병명"].head(3).tolist())
    highest = disease_rows.sort_values("최종상태순위", ascending=False).iloc[0]
    add_para(doc, "1. 귀 기관의 무궁한 발전을 기원합니다.", before=7)
    add_para(
        doc,
        f"2. {event_row['공연시작일']:%Y년 %m월 %d일} 개최 예정인 ‘{event_row['공연명']}’과 관련하여, 과거 공식 감염병 감시자료의 반복 패턴과 최근 동향을 검토한 결과 {disease_names}에 대한 예방관리 안내가 필요한 것으로 판단됩니다.",
    )
    add_para(
        doc,
        f"3. 자동 판단의 최고 업무상태는 ‘{highest['최종업무상태']}’이며, 이는 감염확률이나 의학적 진단이 아니라 담당자 검토를 위한 업무 우선순위입니다.",
    )
    add_para(doc, "4. 행사 운영 시 다음 사항을 반영하여 주시기 바랍니다.")
    add_para(doc, "  가. 행사장 환기 가능 여부와 혼잡 예상구간 사전 점검", line=1.25)
    add_para(doc, "  나. 손 위생·기침예절 등 공식 예방수칙 안내")
    add_para(doc, "  다. 호흡기 증상이 있는 참가자·종사자 대상 사전 안내")
    add_para(doc, "  라. 행사 직전 최신 감시자료 재확인 및 필요 시 안내수준 조정")
    add_para(doc, "붙임  1. 행사별 감염병 판단 및 예방관리 요약 1부.  끝.", before=4)

    add_para(doc, "○ ○ 시  감 염 병 관 리 과 장", size=19, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, before=6, after=1, line=1.0)
    add_para(doc, "(담당자 검토·결재 후 기관 명의로 발송)", size=8, align=WD_ALIGN_PARAGRAPH.CENTER, after=3, line=1.0)

    approval = doc.add_table(rows=2, cols=4)
    approval.alignment = WD_TABLE_ALIGNMENT.RIGHT
    for j, label in enumerate(["담당", "검토", "과장", "기관장"]):
        write_cell(approval.cell(0, j), label, size=8.2, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, shade=True)
        write_cell(approval.cell(1, j), "결재란", size=8.2, align=WD_ALIGN_PARAGRAPH.CENTER)

    info = doc.add_table(rows=2, cols=4)
    info.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_values = [
        ("문서번호", "자동화시연-미부여", "시행", "승인 후 부여"),
        ("주소·전화", "기관 주소·연락처 미입력", "전자우편", "담당자 이메일 미입력"),
    ]
    for i, row_values in enumerate(info_values):
        for j, value in enumerate(row_values):
            write_cell(info.cell(i, j), value, size=7.7, bold=j in (0, 2), align=WD_ALIGN_PARAGRAPH.CENTER if j in (0, 2) else WD_ALIGN_PARAGRAPH.LEFT, shade=j in (0, 2))

    add_para(doc, "※ 본 문서는 프로젝트 시연용 초안이며 담당자 승인 전에는 외부로 발송할 수 없습니다.", size=7.8, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, before=3, line=1.0)

    # Attachment starts on page 2.
    p = add_para(doc, "붙임 1", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, after=2)
    p.paragraph_format.page_break_before = True
    add_para(doc, "행사별 감염병 판단 및 예방관리 요약", size=17, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=8, line=1.0)

    add_para(doc, "1. 행사 기본정보", size=11.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, before=4, after=3, line=1.0)
    event_table = doc.add_table(rows=4, cols=4)
    event_values = [
        ("행사명", event_row["공연명"], "행사기간", f"{event_row['공연시작일']:%Y-%m-%d} ~ {event_row['공연종료일']:%Y-%m-%d}"),
        ("장르", clean_text(event_row.get("공연장르")), "지역", clean_text(event_row["지역"])),
        ("장소", clean_text(event_row["공연장명"]), "노출단계", f"{safe_int(event_row.get('행사노출단계'))}단계" if safe_int(event_row.get("행사노출단계")) else "계산 불가"),
        ("좌석수", f"{safe_int(event_row['공연장좌석수']):,}석" if safe_int(event_row["공연장좌석수"]) else "미확인", "실내/야외", clean_text(event_row.get("실내야외"))),
    ]
    for i, vals in enumerate(event_values):
        for j, val in enumerate(vals):
            write_cell(event_table.cell(i, j), val, size=8.0, bold=j in (0, 2), align=WD_ALIGN_PARAGRAPH.CENTER if j in (0, 2) else WD_ALIGN_PARAGRAPH.LEFT, shade=j in (0, 2))

    add_para(doc, "2. 자동 판단 결과", size=11.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, before=7, after=3, line=1.0)
    top_rows = disease_rows.sort_values(["최종상태순위", "질병명"], ascending=[False, True]).head(3)
    judgement = doc.add_table(rows=1 + len(top_rows), cols=5)
    for j, h in enumerate(["감염병", "과거 반복성", "최근 상태", "최종 상태", "판단 요약"]):
        write_cell(judgement.cell(0, j), h, size=7.6, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, shade=True)
    for i, (_, row) in enumerate(top_rows.iterrows(), start=1):
        vals = [
            row["질병명"], row["과거반복성"],
            row["최근유행상태"] if row["최근상태적용여부"] == "적용" else "행사 직전 재조회",
            row["최종업무상태"],
            f"{safe_int(row['비교가능연도수']) or 0}개년 중 {safe_int(row['높은수준연도수']) or 0}개년 높은 주차",
        ]
        for j, val in enumerate(vals):
            write_cell(judgement.cell(i, j), val, size=7.3, bold=j == 3, align=WD_ALIGN_PARAGRAPH.CENTER if j < 4 else WD_ALIGN_PARAGRAPH.LEFT)

    add_para(doc, "3. 예방관리 협조사항", size=11.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, before=7, after=3, line=1.0)
    actions = [
        ("행사장", "환기 가능 여부와 혼잡 예상구간을 사전 점검하고 운영계획에 반영"),
        ("참가자", "손 씻기·기침예절 및 호흡기 증상 시 마스크 착용·참여 자제 안내"),
        ("운영인력", "예방수칙과 유증상자 발생 시 내부 연락절차 사전 공유"),
        ("행사 직전", "최신 감시자료 재확인 후 필요 시 안내수준 조정"),
    ]
    action_table = doc.add_table(rows=1 + len(actions), cols=2)
    write_cell(action_table.cell(0, 0), "적용 대상", size=8, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, shade=True)
    write_cell(action_table.cell(0, 1), "협조사항", size=8, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, shade=True)
    for i, (target, action) in enumerate(actions, start=1):
        write_cell(action_table.cell(i, 0), target, size=7.8, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        write_cell(action_table.cell(i, 1), action, size=7.8)

    add_para(doc, "4. 자료출처 및 유의사항", size=11.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, before=7, after=3, line=1.0)
    sources = doc.add_table(rows=3, cols=2)
    source_values = [
        ("과거·최근 감시자료", "질병관리청 감염병포털 급성호흡기감염증·인플루엔자 통계"),
        ("예방수칙", "질병관리청 국가건강정보포털 급성호흡기바이러스감염증 예방정보"),
        ("행사정보", "KOPIS 공연·공연시설 데이터"),
    ]
    for i, (label, value) in enumerate(source_values):
        write_cell(sources.cell(i, 0), label, size=7.7, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, shade=True)
        write_cell(sources.cell(i, 1), value, size=7.7)

    add_para(doc, "본 붙임은 과거 반복 패턴과 행사 특성을 바탕으로 담당자 검토를 지원하는 행정 참고자료이며, 인과관계·개인별 감염위험·의료적 진단을 제시하지 않습니다.", size=7.8, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, before=5, line=1.15)

    filename = f"공문예시_{index:02d}_{safe_filename(event_row['공연명'])}.docx"
    out_path = DOCS_DIR / filename
    doc.save(out_path)
    return out_path


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    DOCS_DIR.mkdir(parents=True, exist_ok=True)

    performance, disease, risk_reference, ili_region, source_info = load_inputs()
    performance = normalize_performance(performance)
    validation, validation_issues = validate_inputs(performance, disease, risk_reference, ili_region, source_info)
    pattern = build_pattern_db(disease)
    recent = build_recent_status(disease)
    matches = build_matches(performance, pattern, recent)
    event_summary = build_event_summary(matches)
    prevention_db = build_prevention_db()

    candidates = event_summary[
        (event_summary["기준일까지남은일수"] >= 0)
        & (event_summary["기준일까지남은일수"] <= 14)
        & event_summary["최고업무상태"].isin(["우선 대비", "대비 필요"])
    ].copy()
    if len(candidates) < MAX_DOCS:
        candidates = event_summary[
            (event_summary["기준일까지남은일수"] >= 0)
            & (event_summary["기준일까지남은일수"] <= 30)
            & event_summary["최고업무상태"].isin(["우선 대비", "대비 필요"])
        ].copy()
    candidates["상태순위"] = candidates["최고업무상태"].map(STATUS_RANK)
    candidates = candidates.sort_values(["상태순위", "공연장좌석수", "공연시작일"], ascending=[False, False, True]).head(MAX_DOCS)

    doc_paths = []
    for idx, (_, event) in enumerate(candidates.iterrows(), start=1):
        disease_rows = matches[
            (matches["공연ID"] == event["공연ID"])
            & matches["최종업무상태"].isin(["우선 대비", "대비 필요", "일반 관리"])
        ].copy()
        doc_paths.append(str(create_official_doc(event, disease_rows, prevention_db, idx)))

    send_log = build_send_log(matches, candidates["공연ID"].tolist())

    pattern.to_csv(OUTPUT_DIR / "감염병_주차별_패턴DB.csv", index=False, encoding="utf-8-sig")
    recent.to_csv(OUTPUT_DIR / "2026_최근4주_유행상태.csv", index=False, encoding="utf-8-sig")
    matches.to_csv(OUTPUT_DIR / "행사_감염병_전체판단.csv", index=False, encoding="utf-8-sig")
    event_summary.to_csv(OUTPUT_DIR / "행사별_공문대상_요약.csv", index=False, encoding="utf-8-sig")
    prevention_db.to_csv(OUTPUT_DIR / "감염병별_공식예방수칙_DB.csv", index=False, encoding="utf-8-sig")
    send_log.to_csv(OUTPUT_DIR / "발송이력_관리표.csv", index=False, encoding="utf-8-sig")
    validation_issues.to_csv(OUTPUT_DIR / "입력데이터_검사결과.csv", index=False, encoding="utf-8-sig")
    if not ili_region.empty:
        ili_region.to_csv(OUTPUT_DIR / "지역별_ILI_참고자료.csv", index=False, encoding="utf-8-sig")

    summary = {
        "검증": validation,
        "패턴DB행수": int(len(pattern)),
        "최근유행상태행수": int(len(recent)),
        "미래행사수": int(matches["공연ID"].nunique()),
        "행사감염병판단행수": int(len(matches)),
        "상태분포": matches["최종업무상태"].value_counts().to_dict(),
        "공문생성행사": candidates[["공연ID", "공연명", "공연시작일", "최고업무상태"]].astype(str).to_dict("records"),
        "공문파일": doc_paths,
        "현재완료": [
            "실제 공연·감염병 데이터 점검",
            "실제 지도 데이터 20개 컬럼 보존 및 파생 3단계 누락값 자동 계산",
            "입력 오류·경고 검사표와 원본 파일 해시 생성",
            "2016~2025 주차별 반복 패턴 생성",
            "2026 최근 4주 보정",
            "행사-감염병 자동 매칭",
            "판단사유 자동 생성",
            "기관용 HWPX 공문은 고정 템플릿으로 선택 공연별 생성",
            "발송이력 초기화 및 중복키 생성",
        ],
        "미연결": [
            "상시 예약실행",
            "행사 직전 질병관리청 데이터 자동 다운로드",
            "기관 전자결재·메일 실제 발송",
            "KOPIS에 없는 실내·야외 정보 자동 보완",
        ],
    }
    (OUTPUT_DIR / "입력데이터_검사요약.json").write_text(
        json.dumps(validation, ensure_ascii=False, indent=2, default=str), encoding="utf-8"
    )
    (OUTPUT_DIR / "자동화_실행요약.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2, default=str))