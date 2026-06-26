import streamlit as st
import pandas as pd
from core import DEFAULT_API_KEY, SEARCH_MODES, process_dataframe, to_excel_bytes

st.set_page_config(page_title="ISBN 자동 입력기", page_icon="📚", layout="centered")

st.title("📚 ISBN 자동 입력기")
st.caption("도서명 + 출판사 → 카카오 책 검색 API → ISBN 13자리 자동 완성")

with st.sidebar:
    st.header("⚙️ 설정")
    api_key = st.text_input(
        "카카오 REST API 키",
        value=DEFAULT_API_KEY,
        type="password",
        placeholder=".env 파일에 KAKAO_API_KEY를 설정하거나 여기에 직접 입력",
        help="https://developers.kakao.com 에서 발급",
    )
    st.divider()
    search_mode_label = st.radio(
        "검색 방식",
        options=list(SEARCH_MODES.keys()),
        index=0,
        help="ISBN을 찾을 때 카카오 API에 어떤 검색어를 보낼지 선택합니다.",
    )
    search_mode = SEARCH_MODES[search_mode_label]

uploaded = st.file_uploader(
    "엑셀 파일 업로드 (.xlsx)",
    type=["xlsx"],
    help="'도서명', '출판사', 'ISBN' 컬럼이 있어야 합니다.",
)

if uploaded:
    try:
        df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"파일을 읽을 수 없습니다: {e}")
        st.stop()

    required = {'도서명', '출판사'}
    if not required.issubset(df.columns):
        missing = required - set(df.columns)
        st.error(f"필수 컬럼이 없습니다: {missing}")
        st.stop()

    already = int(
        df.get('ISBN', pd.Series()).fillna('').astype(str)
        .apply(lambda x: len(x.strip()) >= 10).sum()
    )
    need = len(df) - already

    col1, col2, col3 = st.columns(3)
    col1.metric("전체 도서", len(df))
    col2.metric("ISBN 있음", already)
    col3.metric("검색 필요", need)

    st.dataframe(df.head(5), use_container_width=True)

    if st.button("🔍 ISBN 검색 시작", type="primary", disabled=(not api_key)):
        progress_bar = st.progress(0.0, text="준비 중...")
        status_box = st.empty()
        log_box = st.empty()

        stats = {"success": 0, "fail": 0}
        log_lines = []

        def on_progress(current, total, title, isbn, success, skipped):
            pct = current / total
            progress_bar.progress(pct, text=f"{current} / {total} 처리 중...")

            if skipped:
                icon = "⏭"
            elif success:
                icon = "✅"
                stats["success"] += 1
            else:
                icon = "❌"
                stats["fail"] += 1

            short = title[:25] + ("…" if len(title) > 25 else "")
            log_lines.append(f"{icon} [{current}] {short}  →  {isbn or '실패'}")
            log_box.text("\n".join(log_lines[-12:]))
            status_box.markdown(
                f"**✅ 성공 {stats['success']}** &nbsp; **❌ 실패 {stats['fail']}** &nbsp; (전체 {total}개)"
            )

        with st.spinner("카카오 API 호출 중..."):
            result_df = process_dataframe(df, api_key, search_mode, on_progress)

        progress_bar.progress(1.0, text="완료!")
        total_found = stats["success"] + already
        st.success(f"완료! {len(result_df)}개 중 **{total_found}개** ISBN 확보")

        st.subheader("결과 미리보기")
        st.dataframe(result_df, use_container_width=True)

        excel_bytes = to_excel_bytes(result_df)
        st.download_button(
            label="📥 결과 엑셀 다운로드",
            data=excel_bytes,
            file_name="ISBN_완료.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
