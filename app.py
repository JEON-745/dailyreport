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
def process_dukomol(files, account_totals):
    """files: dict(report=, tonghap=, meta=)  ->  결과 파일경로, 요약메시지"""
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

    # 1) 엠케이로드 통합보고서 -> 매체 시트
    if paths["tonghap"]:
        step1 = os.path.join(tmp, "step1.xlsx")
        build_report.run(paths["tonghap"], cur, step1, verify=False)
        cur = step1
        summary.append("✅ 매체 시트(N검색·파워링크·쇼핑·구글·모비온·GFA·GDN) 자동 입력 완료")

    # 2) 메타 raw -> SNS 소재별 효율 + 메모
    new_creatives = []
    if paths["meta"]:
        step2 = os.path.join(tmp, "final.xlsx")
        new_creatives = sns_auto.run(paths["meta"], cur, step2, account_totals=account_totals)
        cur = step2
        summary.append("✅ SNS 소재별 효율 + 전환 분해 메모 자동 입력 완료")

    return cur, summary, new_creatives


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
        if not files.get("tonghap") and not files.get("meta"):
            st.error("통합보고서 또는 메타 보고서 중 최소 하나는 올려주세요.")
            st.stop()
        try:
            with st.spinner("처리 중입니다… (차트·서식 보존하며 값만 채웁니다)"):
                out_path, summary, new_creatives = prof["process"](files, account_totals)
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
