"""
광고 리포트 자동화 웹앱 (Streamlit)
- 브랜드별 '프로필'을 선택 -> 필요한 파일 업로드 -> 자동 채움 -> 결과 다운로드
- 듀코몰 프로필이 첫 번째로 구현됨. 다른 브랜드는 PROFILES에 추가하면 됨.
실행:  streamlit run app.py
"""
import os, tempfile, traceback, datetime
import streamlit as st
import build_report
import sns_auto

st.set_page_config(page_title="리포트 자동화", page_icon="📊", layout="centered")

# ──────────────────────────── 비밀번호 ────────────────────────────
def check_password():
    if st.session_state.get("auth_ok"):
        return True
    st.title("📊 광고 리포트 자동화")
    pw = st.text_input("접속 비밀번호", type="password")
    if pw:
        if pw == st.secrets.get("APP_PASSWORD", ""):
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
    st.stop()

# ──────────────────────────── 듀코몰 처리 ────────────────────────────
def process_dukomol(files, account_totals, only_dates=None):
    """files: dict(report=, tonghap=, meta=)  ->  결과 파일경로, 요약메시지"""
    import gc
    tmp = tempfile.mkdtemp()
    paths = {}
    for k, up in files.items():
        if up is None:
            paths[k] = None; continue
        p = os.path.join(tmp, f"{k}.xlsx")
        with open(p, "wb") as o:
            o.write(up.getbuffer())
        paths[k] = p

    summary = []
    cur = paths["report"]

    # 1) 엠케이로드 통합보고서 -> 매체 시트 (only_dates 지정 시 그 날짜만)
    if paths["tonghap"]:
        step1 = os.path.join(tmp, "step1.xlsx")
        build_report.run(paths["tonghap"], cur, step1, verify=False, only_dates=only_dates)
        cur = step1
        gc.collect()
        if only_dates:
            ds = ", ".join(sorted(str(d) for d in only_dates))
            summary.append(f"✅ 매체 시트 자동 입력 완료 (날짜: {ds})")
        else:
            summary.append("✅ 매체 시트(N검색·파워링크·쇼핑·구글·모비온·GFA·GDN) 자동 입력 완료 (전체 날짜)")

    # 2) 메타 raw -> SNS 소재별 효율 + 메모
    new_creatives = []
    if paths["meta"]:
        step2 = os.path.join(tmp, "final.xlsx")
        new_creatives = sns_auto.run(paths["meta"], cur, step2, account_totals=account_totals)
        cur = step2
        gc.collect()
        summary.append("✅ SNS 소재별 효율 + 전환 분해 메모 자동 입력 완료")

    # 3) 추가 시트: [페이스북] 일별 + [듀퐁소품] 4표 + GFA 브로이어 카탈로그
    if paths["tonghap"] and paths["meta"]:
        import dukomol_extra
        step3 = os.path.join(tmp, "final2.xlsx")
        dukomol_extra.run(cur, paths["tonghap"], paths["meta"], step3, only_dates=only_dates)
        cur = step3
        gc.collect()
        summary.append("✅ [페이스북]·[듀퐁소품]·GFA 브로이어 카탈로그 자동 입력 완료")

    return cur, summary, new_creatives


# ──────────────────────────── NSR 처리 ────────────────────────────
def process_nsr(files, account_totals, only_dates=None):
    """NSR: 단일 월시트(가로블록) + 메타 소재별 효율.
    files: dict(report=, tonghap=, meta=) -> 결과 파일경로, 요약, 신규소재"""
    import gc, nsr_media
    tmp = tempfile.mkdtemp()
    paths = {}
    for k, up in files.items():
        if up is None:
            paths[k] = None; continue
        p = os.path.join(tmp, f"{k}.xlsx")
        with open(p, "wb") as o:
            o.write(up.getbuffer())
        paths[k] = p

    if not paths.get("tonghap") or not paths.get("meta"):
        raise ValueError("NSR은 '엠케이로드 통합보고서'와 '메타 소재별 raw' 두 파일이 모두 필요합니다.")

    out = os.path.join(tmp, "nsr_final.xlsx")
    new = nsr_media.run(paths["tonghap"], paths["meta"], paths["report"], out,
                        only_dates=only_dates)
    gc.collect()

    summary = []
    if only_dates:
        ds = ", ".join(sorted(str(d) for d in only_dates))
        summary.append(f"✅ 매체 광고영역(네이버·GFA·크리테오) + META 일별 자동 입력 완료 (날짜: {ds})")
    else:
        summary.append("✅ 매체 광고영역(네이버·GFA·크리테오) + META 일별 자동 입력 완료 (raw의 모든 날짜)")
    summary.append("✅ 메타 소재별 효율(해당 주차, ASC+전환 합산) 자동 입력 완료")
    # app 표시 형식((브랜드,기획전명), 지표)에 맞춰 변환
    new2 = [(("NSR", k), m) for (k, m) in new]
    return out, summary, new2


# ──────────────────────────── 아틀라스 처리 ────────────────────────────
def process_atlas(files, account_totals, only_dates=None):
    """아틀라스(NSR형): 단일 월시트(파워링크 PC/MO·GFA·SNS) + 메타 소재별 효율_*.
    files: dict(report=, tonghap=, meta=) -> 결과 파일경로, 요약, 신규소재"""
    import gc, atlas_media
    tmp = tempfile.mkdtemp()
    paths = {}
    for k, up in files.items():
        if up is None:
            paths[k] = None; continue
        p = os.path.join(tmp, f"{k}.xlsx")
        with open(p, "wb") as o:
            o.write(up.getbuffer())
        paths[k] = p

    if not paths.get("tonghap") or not paths.get("meta"):
        raise ValueError("아틀라스는 '엠케이로드 통합보고서'와 '메타 소재별 raw' 두 파일이 모두 필요합니다.")

    out = os.path.join(tmp, "atlas_final.xlsx")
    new = atlas_media.run(paths["tonghap"], paths["meta"], paths["report"], out,
                          only_dates=only_dates)
    gc.collect()

    summary = []
    if only_dates:
        ds = ", ".join(sorted(str(d) for d in only_dates))
        summary.append(f"✅ 매체 광고영역(파워링크 PC·MO·GFA) + META 일별 자동 입력 완료 (날짜: {ds})")
    else:
        summary.append("✅ 매체 광고영역(파워링크 PC·MO·GFA) + META 일별 자동 입력 완료 (raw의 모든 날짜)")
    summary.append("✅ 메타 소재별 효율(해당 주차, 기획전명+번호 매칭, 전환+트래픽 합산) 자동 입력 완료")
    new2 = [(("아틀라스", k), m) for (k, m) in new]
    return out, summary, new2


# ──────────────────────────── 챌린저골프웨어 처리 ────────────────────────────
def process_challenger(files, account_totals, only_dates=None):
    """챌린저: 매체 리포트(+META) -> 통합보고서 일간리포트(자사몰·브랜드스토어) 채움.
    결과물 1개(=채워진 통합보고서). files: dict(report=통합보고서, source=매체리포트, meta=)"""
    import gc, challenger_media
    tmp = tempfile.mkdtemp()
    paths = {}
    for k, up in files.items():
        if up is None:
            paths[k] = None; continue
        p = os.path.join(tmp, f"{k}.xlsx")
        with open(p, "wb") as o:
            o.write(up.getbuffer())
        paths[k] = p

    if not paths.get("source"):
        raise ValueError("챌린저는 '챌린저 매체 리포트'(검색광고·브랜드검색·GFA·모비온·구글)가 필요합니다.")

    out = os.path.join(tmp, "challenger_final.xlsx")
    meta_missing = challenger_media.run(paths["source"], paths.get("meta"), paths["report"], out,
                                        only_dates=only_dates)
    gc.collect()

    summary = []
    if only_dates:
        ds = ", ".join(sorted(str(d) for d in only_dates))
        summary.append(f"✅ 자사몰(브랜드검색·파워링크·GFA·모비온·구글) + 브랜드스토어(GFA) 자동 입력 완료 (날짜: {ds})")
    else:
        summary.append("✅ 자사몰(브랜드검색·파워링크·GFA·모비온·구글) + 브랜드스토어(GFA) 자동 입력 완료 (매체 리포트의 모든 날짜)")
    if paths.get("meta"):
        summary.append("✅ META(SNS 광고) 일별 자동 입력 완료 (지출×1.1÷0.85)")
        summary.append("✅ META 소재별 효율 주차 블록 재합산 완료 (노출·클릭·도달·전환·장바구니·빈도·매출·광고비) — 빈도=노출/도달, 블록 맨 마지막 합계행은 수기 유지")
        if meta_missing:
            lst = ", ".join(f"{a} {b}({c:,}원)" for a, b, c in meta_missing)
            summary.append(f"⚠ META 소재별효율: 시트에 행이 없는 소재 {len(meta_missing)}건 — 직접 행 추가 필요: {lst}")
    return out, summary, []


# ──────────────────────────── 프로필 정의 ────────────────────────────
# 새 브랜드 추가 시 여기에 프로필을 추가하면 됩니다.
PROFILES = {
    "듀코몰 (DUKO MALL)": {
        "ready": True,
        "inputs": [
            ("report",  "최종 리포트 (종합매출 리포트)", "현재까지 작성된 리포트 .xlsx"),
            ("tonghap", "엠케이로드 통합보고서",          "매체 데이터 .xlsx"),
            ("meta",    "메타 소재별 Raw 보고서",         "SNS 소재 데이터 .xlsx"),
        ],
        "has_sns_totals": True,
        "process": process_dukomol,
    },
    # "다른브랜드 (BRAND B)": {"ready": False},
    "NSR": {
        "ready": True,
        "inputs": [
            ("report",  "최종 리포트 (NSR 통합보고서)", "현재까지 작성된 리포트 .xlsx"),
            ("tonghap", "엠케이로드 통합보고서",         "매체 데이터 .xlsx (브검·파워·쇼핑·GFA·크리테오)"),
            ("meta",    "메타 소재별 Raw 보고서",        "SNS 소재 데이터 .xlsx ('일' 날짜열 포함)"),
        ],
        "has_sns_totals": False,
        "process": process_nsr,
    },
    "아틀라스 (ATLAS)": {
        "ready": True,
        "inputs": [
            ("report",  "최종 리포트 (아틀라스 통합보고서)", "현재까지 작성된 리포트 .xlsx"),
            ("tonghap", "엠케이로드 통합보고서",            "매체 데이터 .xlsx (파워링크·GFA)"),
            ("meta",    "메타 소재별 Raw 보고서",           "SNS 소재 데이터 .xlsx ('일' 날짜열 포함)"),
        ],
        "has_sns_totals": False,
        "process": process_atlas,
    },
    "챌린저골프웨어 (CHALLENGER)": {
        "ready": True,
        "inputs": [
            ("report",  "최종 리포트 (챌린저 통합보고서)", "일간리포트(자사몰·브랜드스토어)를 채울 통합보고서 .xlsx ← 결과물"),
            ("source",  "챌린저 매체 리포트",              "검색광고·브랜드검색·네이버GFA·모비온·구글 시트 .xlsx ← 소스"),
            ("meta",    "메타 소재별 Raw 보고서",          "SNS 소재 데이터 .xlsx ('일' 날짜열 포함)"),
        ],
        "has_sns_totals": False,
        "process": process_challenger,
    },
}

# ──────────────────────────── 메인 UI ────────────────────────────
def main():
    check_password()

    st.title("📊 광고 리포트 자동화")
    st.caption("브랜드를 고르고 파일을 올리면, 리포트가 자동으로 채워집니다.")

    brand = st.selectbox("브랜드 / 리포트 선택", list(PROFILES.keys()))
    prof = PROFILES[brand]

    if not prof.get("ready"):
        st.info("이 브랜드 프로필은 아직 준비 중입니다. 해당 브랜드의 리포트 양식을 등록하면 활성화됩니다.")
        st.stop()

    # ── 기입할 날짜 선택 (매체 시트) ──
    only_dates = None
    date_mode = st.radio(
        "기입할 날짜 (매체 시트)",
        ["통합보고서의 모든 날짜", "특정 날짜만"],
        horizontal=True,
    )
    if date_mode == "특정 날짜만":
        today = datetime.date.today()
        picked = st.date_input(
            "채울 날짜 (하루 또는 기간 선택)",
            value=(today, today),
            help="선택한 날짜만 매체 시트에 기입합니다. 리포트와 통합보고서 양쪽에 그 날짜가 있어야 채워집니다.",
        )
        if isinstance(picked, (list, tuple)):
            if len(picked) == 2:
                d0, d1 = picked
                only_dates = {d0 + datetime.timedelta(days=i) for i in range((d1 - d0).days + 1)}
            elif len(picked) == 1:
                only_dates = {picked[0]}
        else:
            only_dates = {picked}
        if only_dates:
            st.caption("선택됨: " + ", ".join(sorted(str(d) for d in only_dates)))

    st.divider()
    files = {}
    for key, label, help_ in prof["inputs"]:
        files[key] = st.file_uploader(label, type=["xlsx"], help=help_, key=f"u_{key}")

    # SNS 주차 TOTAL(메타 계정 총결과) — 선택 입력
    account_totals = None
    if prof.get("has_sns_totals"):
        with st.expander("⚙️ (선택) SNS 주차 TOTAL = 메타 '총 결과' 값 입력"):
            st.caption("도달·빈도는 메타 계정 기준 중복제거 값이라 합산이 안 됩니다. "
                       "메타 '총 결과' 행 값을 넣으면 해당 주차 TOTAL 행에 그대로 기입합니다. "
                       "비워두면 기존 SUM 수식을 유지합니다.")
            c1, c2, c3 = st.columns(3)
            노출 = c1.number_input("노출", min_value=0, value=0, step=1)
            클릭 = c2.number_input("링크 클릭", min_value=0, value=0, step=1)
            도달 = c3.number_input("도달", min_value=0, value=0, step=1)
            c4, c5, c6 = st.columns(3)
            전환 = c4.number_input("구매(전환)", min_value=0, value=0, step=1)
            장바구니 = c5.number_input("장바구니", min_value=0, value=0, step=1)
            빈도 = c6.number_input("빈도", min_value=0.0, value=0.0, step=0.01, format="%.2f")
            if any([노출, 클릭, 도달, 전환, 장바구니, 빈도]):
                account_totals = dict(노출=노출, 클릭=클릭, 도달=도달,
                                      전환=전환, 장바구니=장바구니, 빈도=빈도)

    st.divider()
    if st.button("🚀 자동 입력 실행", type="primary", use_container_width=True):
        if not files.get("report"):
            st.error("최종 리포트 파일을 올려주세요.")
            st.stop()
        if not files.get("tonghap") and not files.get("source") and not files.get("meta"):
            st.error("매체 데이터(통합보고서·매체 리포트) 또는 메타 보고서 중 최소 하나는 올려주세요.")
            st.stop()
        try:
            with st.spinner("처리 중입니다… (차트·서식 보존하며 값만 채웁니다)"):
                out_path, summary, new_creatives = prof["process"](files, account_totals, only_dates)
            st.success("완료되었습니다!")
            for s in summary:
                st.write(s)

            if new_creatives:
                st.warning(f"시트에 없는 신규 소재 {len(new_creatives)}건 — 행을 직접 추가하세요(썸네일 포함).")
                rows = []
                for (b, jp), m in new_creatives:
                    freq = round(m["노출"] / m["도달"], 2) if m["도달"] else 0
                    rows.append({"브랜드": b, "기획전명": jp, "노출": int(m["노출"]),
                                 "클릭": int(m["클릭"]), "도달": int(m["도달"]),
                                 "전환": int(m["전환"]), "장바구니": int(m["장바구니"]), "빈도": freq})
                st.dataframe(rows, use_container_width=True)

            with open(out_path, "rb") as f:
                st.download_button("📥 채워진 리포트 다운로드", f.read(),
                                   file_name=f"{brand.split(' ')[0]}_리포트_{datetime.date.today()}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
        except Exception as e:
            st.error(f"처리 중 오류가 발생했습니다: {e}")
            st.code(traceback.format_exc())


if __name__ == "__main__":
    main()
