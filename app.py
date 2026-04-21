"""
율이공방 — 토요상설공연 운영 체크리스트 (Streamlit 웹앱)
구글 스프레드시트 연동
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime

from data_manager import (
    ChecklistManager, DEFAULT_ITEMS,
    STAGE_LABELS, GENRE_LIST, WEATHER_LIST, EVAL_LIST,
)

# ── 페이지 설정 ──
st.set_page_config(
    page_title="토요상설공연 운영 체크리스트",
    page_icon="🎵",
    layout="wide",
)

# ── 스타일 ──
STAGE_COLORS = {
    "A": "#E8F5E9",
    "B": "#E3F2FD",
    "C": "#FFF9C4",
    "D": "#FCE4EC",
    "E": "#F3E5F5",
}
STAGE_COLORS_DARK = {
    "A": "#4CAF50",
    "B": "#2196F3",
    "C": "#FFC107",
    "D": "#E91E63",
    "E": "#9C27B0",
}

st.markdown("""
<style>
    .main .block-container { max-width: 1200px; padding-top: 1rem; }
    .stage-header {
        padding: 8px 16px; border-radius: 6px; margin: 8px 0 4px 0;
        font-weight: bold; font-size: 1.05rem;
    }
    .check-done { color: #4CAF50; font-weight: bold; }
    .check-na { color: #9E9E9E; }
    .check-undone { color: #E53935; }
    .metric-card {
        background: #313244; border-radius: 8px; padding: 16px;
        text-align: center; margin: 4px;
    }
    .metric-value { font-size: 2rem; font-weight: bold; color: #cdd6f4; }
    .metric-label { font-size: 0.85rem; color: #a6adc8; }
    div[data-testid="stExpander"] { border: 1px solid #45475a; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)


# ── 구글 시트 연결 ──
@st.cache_resource
def get_mgr():
    """ChecklistManager 싱글톤 (구글 시트 연동)"""
    gsheet = None
    try:
        from gsheet_sync import GoogleSheetSync
        if "gcp_service_account" in st.secrets:
            gsheet = GoogleSheetSync(
                credentials_dict=dict(st.secrets["gcp_service_account"]),
                spreadsheet_id=st.secrets["spreadsheet"]["spreadsheet_id"],
            )
    except Exception as e:
        st.sidebar.warning(f"구글 시트 연결 실패: {e}")
    return ChecklistManager(gsheet_sync=gsheet)


def reload_mgr():
    get_mgr.clear()
    st.rerun()


# ══════════════════════════════════════════════════════════════
#  탭 1: 회차별 운영 체크
# ══════════════════════════════════════════════════════════════

def render_tab_check():
    mgr = get_mgr()

    # ── 사이드바: 회차 선택/등록/삭제 ──
    st.sidebar.markdown("### 회차 관리")

    round_list = mgr.round_list
    if round_list:
        rnd_options = [f"{r}회" for r in round_list]
        sel = st.sidebar.selectbox("회차 선택", rnd_options, key="sel_round")
        cur_rnd = int(sel.replace("회", ""))
    else:
        cur_rnd = None
        st.sidebar.info("등록된 회차가 없습니다.")

    # 회차 등록
    with st.sidebar.expander("회차 등록", expanded=not round_list):
        new_rnd = st.number_input("회차 번호", min_value=1, max_value=50,
                                   value=max(round_list) + 1 if round_list else 1,
                                   key="new_rnd")
        new_date = st.text_input("공연일 (YYYY-MM-DD)", key="new_date")
        new_group = st.text_input("출연단체", key="new_group")
        new_genre = st.selectbox("장르", [""] + GENRE_LIST, key="new_genre")
        new_time = st.text_input("공연시간", key="new_time",
                                  placeholder="예: 14:00~15:30")
        new_weather = st.selectbox("날씨", [""] + WEATHER_LIST, key="new_weather")
        new_staff = st.text_input("담당자", key="new_staff")

        if st.button("회차 등록", key="btn_add_round", use_container_width=True):
            info = {
                "공연일": new_date, "출연단체": new_group,
                "장르": new_genre, "공연시간": new_time,
                "날씨": new_weather, "담당자": new_staff,
            }
            mgr.add_round(int(new_rnd), info)
            st.success(f"{int(new_rnd)}회차 등록 완료!")
            st.rerun()

    # 회차 삭제
    if cur_rnd is not None:
        with st.sidebar.expander("회차 삭제"):
            st.warning(f"{cur_rnd}회차를 삭제하면 모든 체크 데이터가 사라집니다.")
            if st.button("삭제 실행", key="btn_del_round"):
                mgr.delete_round(cur_rnd)
                st.success("삭제 완료!")
                st.rerun()

    st.sidebar.markdown("---")
    if st.sidebar.button("구글 시트 새로고침", key="btn_reload",
                          use_container_width=True):
        reload_mgr()

    # ── 메인 영역 ──
    if cur_rnd is None:
        st.info("좌측에서 회차를 등록해 주세요.")
        return

    info = mgr.round_info.get(cur_rnd, {})
    rate = mgr.get_round_rate(cur_rnd)
    status = mgr.get_round_status(cur_rnd)

    # 공연 정보 헤더 — 1행: 회차 + 상태
    h1, h2 = st.columns([4, 1])
    with h1:
        st.markdown(f"### {cur_rnd}회차 — {info.get('출연단체', '미정')}")
    with h2:
        color = "#4CAF50" if status == "완료" else "#FFC107" if status == "진행중" else "#9E9E9E"
        st.markdown(f"<div style='text-align:right; padding-top:8px;'>"
                    f"<span style='color:{color}; font-weight:bold; font-size:1.2rem;'>"
                    f"● {status} ({rate}%)</span></div>", unsafe_allow_html=True)

    # 2행: 공연 정보 (한 줄)
    info_parts = []
    for label, key in [("공연일", "공연일"), ("장르", "장르"), ("시간", "공연시간"),
                        ("날씨", "날씨"), ("담당", "담당자")]:
        val = info.get(key, "")
        if val:
            info_parts.append(f"**{label}** {val}")
    if info_parts:
        st.markdown(" &nbsp;·&nbsp; ".join(info_parts), unsafe_allow_html=True)

    st.progress(rate / 100)

    # ── 섹션별 이전 회차 복사 (B·C·D·E만) ──
    copy_stages = {"B": "공연 당일 — 현장 세팅", "C": "공연 당일 — 리허설",
                   "D": "공연 중", "E": "공연 후 — 마무리"}
    copy_cols = st.columns(len(copy_stages))
    for i, (stg, stg_name) in enumerate(copy_stages.items()):
        with copy_cols[i]:
            if st.button(f"📋 {stg_name}", key=f"copy_stage_{stg}_{cur_rnd}",
                         use_container_width=True):
                if mgr.copy_prev_stage(cur_rnd, stg):
                    for key in list(st.session_state.keys()):
                        if key.startswith(("st_", "tm_", "sf_", "mm_")) and f"_{cur_rnd}_" in key:
                            code_part = key.split(f"_{cur_rnd}_")[-1]
                            if any(code_part == c for s, c, _ in mgr.items if s == stg):
                                del st.session_state[key]
                    st.success(f"{stg_name} 섹션 복사 완료!")
                    st.rerun()
                else:
                    st.warning("이전 회차 데이터가 없습니다.")

    # ── 체크리스트 폼 ──
    with st.form(f"check_form_{cur_rnd}", border=False):
        grouped = {}
        for stage, code, name in mgr.items:
            grouped.setdefault(stage, []).append((code, name))

        for stage in STAGE_LABELS:
            items_in_stage = grouped.get(stage, [])
            if not items_in_stage:
                continue

            stage_rate = mgr.get_stage_rate(cur_rnd, stage)
            sc = STAGE_COLORS.get(stage, "#333")
            scd = STAGE_COLORS_DARK.get(stage, "#666")

            st.markdown(
                f'<div class="stage-header" style="background:{sc}; color:#333;">'
                f'{STAGE_LABELS[stage]} — {stage_rate}%</div>',
                unsafe_allow_html=True
            )

            for code, name in items_in_stage:
                cd = mgr.get_check(cur_rnd, code)
                cols = st.columns([0.8, 3, 1, 1, 2])
                status_options = ["미완료", "완료", "해당없음"]
                cur_status = cd["상태"] if cd["상태"] in status_options else "미완료"
                cur_idx = status_options.index(cur_status)

                with cols[0]:
                    st.selectbox(
                        "상태", status_options, index=cur_idx,
                        key=f"st_{cur_rnd}_{code}", label_visibility="collapsed"
                    )
                with cols[1]:
                    st.markdown(f"**{code}** {name}")
                with cols[2]:
                    st.text_input(
                        "완료시간", value=cd["완료시간"],
                        key=f"tm_{cur_rnd}_{code}", label_visibility="collapsed",
                        placeholder="HH:MM"
                    )
                with cols[3]:
                    st.text_input(
                        "담당", value=cd["담당"],
                        key=f"sf_{cur_rnd}_{code}", label_visibility="collapsed",
                        placeholder="담당자"
                    )
                with cols[4]:
                    st.text_input(
                        "메모", value=cd["메모"],
                        key=f"mm_{cur_rnd}_{code}", label_visibility="collapsed",
                        placeholder="메모"
                    )

        # 운영 총평 영역
        st.markdown("---")
        st.markdown("#### 운영 총평")
        rv = mgr.reviews.get(cur_rnd, {})
        rc1, rc2 = st.columns(2)
        with rc1:
            st.text_input("관객수", value=rv.get("예상관객수", ""),
                          key=f"rv_aud_{cur_rnd}", placeholder="숫자")
            st.selectbox("공연평가", [""] + EVAL_LIST,
                         index=(EVAL_LIST.index(rv["공연평가"]) + 1
                                if rv.get("공연평가") in EVAL_LIST else 0),
                         key=f"rv_eval_{cur_rnd}")
        with rc2:
            st.text_area("총평", value=rv.get("총평", ""),
                         key=f"rv_review_{cur_rnd}", height=80)
            st.text_area("개선사항", value=rv.get("개선사항", ""),
                         key=f"rv_improve_{cur_rnd}", height=80)

        # 버튼 행
        bc1, bc2, bc3 = st.columns(3)
        with bc1:
            submitted = st.form_submit_button("💾 저장", use_container_width=True,
                                               type="primary")
        with bc2:
            copy_prev = st.form_submit_button("📋 이전 회차 복사",
                                               use_container_width=True)
        with bc3:
            reset = st.form_submit_button("🔄 초기화", use_container_width=True)

    # ── 폼 제출 처리 ──
    if submitted:
        checks_data = {}
        for _, code, _ in mgr.items:
            s = st.session_state.get(f"st_{cur_rnd}_{code}", "미완료")
            tm = st.session_state.get(f"tm_{cur_rnd}_{code}", "")
            sf = st.session_state.get(f"sf_{cur_rnd}_{code}", "")
            mm = st.session_state.get(f"mm_{cur_rnd}_{code}", "")
            if s == "완료" and not tm:
                tm = datetime.now().strftime("%H:%M")
            checks_data[code] = {
                "상태": s, "완료시간": tm, "담당": sf, "메모": mm
            }
        review_data = {
            "예상관객수": st.session_state.get(f"rv_aud_{cur_rnd}", ""),
            "공연평가": st.session_state.get(f"rv_eval_{cur_rnd}", ""),
            "총평": st.session_state.get(f"rv_review_{cur_rnd}", ""),
            "개선사항": st.session_state.get(f"rv_improve_{cur_rnd}", ""),
        }
        mgr.save_checks(cur_rnd, checks_data, review_data)
        st.success("저장 완료!")
        st.rerun()

    if copy_prev:
        if mgr.copy_prev_checks(cur_rnd):
            for key in list(st.session_state.keys()):
                if key.startswith(("st_", "tm_", "sf_", "mm_")) and f"_{cur_rnd}_" in key:
                    del st.session_state[key]
            st.success("이전 회차 체크를 복사했습니다.")
            st.rerun()
        else:
            st.warning("이전 회차 데이터가 없습니다.")

    if reset:
        mgr.reset_checks(cur_rnd)
        st.success("초기화 완료!")
        st.rerun()


# ══════════════════════════════════════════════════════════════
#  탭 2: 연간 현황 대시보드
# ══════════════════════════════════════════════════════════════

def render_tab_dashboard():
    mgr = get_mgr()
    rounds = mgr.round_list
    if not rounds:
        st.info("등록된 회차가 없습니다.")
        return

    # ── 요약 카드 ──
    total = len(rounds)
    rates = [mgr.get_round_rate(r) for r in rounds]
    avg_rate = round(sum(rates) / len(rates), 1) if rates else 0

    audiences = []
    evals = {"매우좋음": 0, "좋음": 0, "보통": 0, "미흡": 0, "매우미흡": 0}
    for r in rounds:
        rv = mgr.reviews.get(r, {})
        try:
            audiences.append(int(rv.get("예상관객수", 0) or 0))
        except (ValueError, TypeError):
            audiences.append(0)
        ev = rv.get("공연평가", "")
        if ev in evals:
            evals[ev] += 1

    total_aud = sum(audiences)

    mc = st.columns(4)
    cards = [
        ("총 운영 회차", f"{total}회"),
        ("평균 완료율", f"{avg_rate}%"),
        ("누적 관객수", f"{total_aud:,}명"),
        ("최다 평가", max(evals, key=evals.get) if any(evals.values()) else "-"),
    ]
    for col, (label, value) in zip(mc, cards):
        col.markdown(
            f'<div class="metric-card">'
            f'<div class="metric-value">{value}</div>'
            f'<div class="metric-label">{label}</div></div>',
            unsafe_allow_html=True
        )

    st.markdown("")

    # ── 차트 ──
    ch1, ch2 = st.columns(2)

    with ch1:
        labels = [f"{r}회" for r in rounds]
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=labels, y=audiences, name="관객수",
            marker_color="#90CAF9", opacity=0.7
        ))
        fig.add_trace(go.Scatter(
            x=labels, y=rates, name="완료율(%)",
            mode="lines+markers", yaxis="y2",
            line=dict(color="#2B3A67", width=2),
            marker=dict(size=8)
        ))
        fig.update_layout(
            title="회차별 완료율 · 관객수",
            yaxis=dict(title="관객수 (명)"),
            yaxis2=dict(title="완료율 (%)", overlaying="y", side="right",
                        range=[0, 110]),
            template="plotly_dark",
            height=400,
            legend=dict(orientation="h", y=1.12),
        )
        st.plotly_chart(fig, use_container_width=True)

    with ch2:
        # 단계별 평균 완료율
        stage_avg = {}
        for stage in STAGE_LABELS:
            stage_rates = [mgr.get_stage_rate(r, stage) for r in rounds]
            stage_avg[stage] = round(sum(stage_rates) / len(stage_rates), 1)

        fig2 = go.Figure()
        stage_labels = [STAGE_LABELS[s].split("—")[0].strip() for s in STAGE_LABELS]
        stage_values = [stage_avg[s] for s in STAGE_LABELS]
        stage_clrs = [STAGE_COLORS_DARK[s] for s in STAGE_LABELS]

        fig2.add_trace(go.Bar(
            x=stage_labels, y=stage_values,
            marker_color=stage_clrs, text=[f"{v}%" for v in stage_values],
            textposition="auto"
        ))
        fig2.update_layout(
            title="단계별 평균 완료율",
            yaxis=dict(range=[0, 110], title="완료율 (%)"),
            template="plotly_dark",
            height=400,
        )
        st.plotly_chart(fig2, use_container_width=True)

    # ── 회차별 요약 테이블 ──
    st.markdown("#### 회차별 요약")
    rows = []
    for r in rounds:
        info = mgr.round_info.get(r, {})
        rv = mgr.reviews.get(r, {})
        rows.append({
            "회차": f"{r}회",
            "공연일": info.get("공연일", ""),
            "출연단체": info.get("출연단체", ""),
            "장르": info.get("장르", ""),
            "완료율": f"{mgr.get_round_rate(r)}%",
            "상태": mgr.get_round_status(r),
            "공연평가": rv.get("공연평가", ""),
            "관객수": rv.get("예상관객수", ""),
        })
    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════
#  탭 3: 항목별 통계
# ══════════════════════════════════════════════════════════════

def render_tab_stats():
    mgr = get_mgr()
    stats, total_rounds = mgr.get_item_stats()
    if total_rounds == 0:
        st.info("등록된 회차가 없습니다.")
        return

    st.markdown(f"#### 항목별 이행 통계 (전체 {total_rounds}회차)")

    rows = []
    for stage, code, name in mgr.items:
        s = stats.get(code, {"완료수": 0, "미완료수": 0, "해당없음수": 0})
        done = s["완료수"]
        na = s["해당없음수"]
        undone = s["미완료수"]
        denom = done + undone + na
        rate = round((done + na) / denom * 100, 1) if denom else 0
        rows.append({
            "단계": stage,
            "코드": code,
            "항목명": name,
            "완료": done,
            "미완료": undone,
            "해당없음": na,
            "이행률": f"{rate}%",
        })

    df = pd.DataFrame(rows)

    # 단계별 필터
    stage_filter = st.selectbox("단계 필터", ["전체"] + list(STAGE_LABELS.keys()),
                                 key="stats_filter")
    if stage_filter != "전체":
        df = df[df["단계"] == stage_filter]

    st.dataframe(df, use_container_width=True, hide_index=True)

    # 이행률 낮은 항목 하이라이트
    low_items = [r for r in rows if float(r["이행률"].replace("%", "")) < 80]
    if low_items:
        st.markdown("#### ⚠ 이행률 80% 미만 항목")
        for item in sorted(low_items, key=lambda x: float(x["이행률"].replace("%", ""))):
            st.markdown(f"- **{item['코드']}** {item['항목명']} — {item['이행률']}")

    # 항목별 이행률 차트
    if len(rows) > 0:
        fig = px.bar(
            pd.DataFrame(rows),
            x="코드", y=[float(r["이행률"].replace("%", "")) for r in rows],
            color="단계",
            color_discrete_map=STAGE_COLORS_DARK,
            title="항목별 이행률",
            labels={"y": "이행률 (%)", "코드": "항목 코드"},
        )
        fig.update_layout(
            template="plotly_dark",
            yaxis=dict(range=[0, 110]),
            height=400,
        )
        st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════════════════════════
#  탭 4: 항목 관리
# ══════════════════════════════════════════════════════════════

def render_tab_items():
    mgr = get_mgr()

    st.markdown("#### 체크리스트 항목 관리")

    # 현재 항목 목록
    rows = []
    for i, (stage, code, name) in enumerate(mgr.items):
        rows.append({"#": i + 1, "단계": stage, "코드": code, "항목명": name})

    if rows:
        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("---")

    # ── 항목 추가 ──
    with st.expander("항목 추가"):
        with st.form("add_item_form"):
            ac1, ac2 = st.columns([1, 3])
            with ac1:
                add_stage = st.selectbox("단계", list(STAGE_LABELS.keys()),
                                          key="add_stage")
            with ac2:
                add_name = st.text_input("항목명", key="add_name")
            if st.form_submit_button("추가", use_container_width=True):
                if add_name.strip():
                    new_code = mgr.gen_code(add_stage)
                    mgr.add_item(add_stage, new_code, add_name.strip())
                    st.success(f"[{new_code}] {add_name.strip()} 추가 완료!")
                    st.rerun()
                else:
                    st.warning("항목명을 입력하세요.")

    # ── 항목 수정 ──
    if st.session_state.get("item_edit_mode"):
        target_code = st.session_state.get("item_edit_code")
        target_item = None
        for s, c, n in mgr.items:
            if c == target_code:
                target_item = (s, c, n)
                break
        if target_item is None:
            st.session_state["item_edit_mode"] = False
            st.rerun()
        else:
            st.markdown(f"##### 항목 수정 — {target_code}")
            with st.form(f"edit_item_{target_code}"):
                ec1, ec2 = st.columns([1, 3])
                with ec1:
                    edit_stage = st.selectbox(
                        "단계", list(STAGE_LABELS.keys()),
                        index=list(STAGE_LABELS.keys()).index(target_item[0]),
                        key=f"edit_stage_{target_code}")
                with ec2:
                    edit_name = st.text_input("항목명", value=target_item[2],
                                               key=f"edit_name_{target_code}")
                ebc1, ebc2 = st.columns(2)
                with ebc1:
                    if st.form_submit_button("수정 저장", use_container_width=True,
                                              type="primary"):
                        new_code = f"{edit_stage}-{target_code.split('-')[1]}"
                        mgr.update_item(target_code, edit_stage,
                                         new_code, edit_name.strip())
                        st.session_state["item_edit_mode"] = False
                        st.success("수정 완료!")
                        st.rerun()
                with ebc2:
                    if st.form_submit_button("취소", use_container_width=True):
                        st.session_state["item_edit_mode"] = False
                        st.rerun()
    else:
        with st.expander("항목 수정 / 삭제"):
            if not mgr.items:
                st.info("항목이 없습니다.")
            else:
                item_options = [f"{code} — {name}" for _, code, name in mgr.items]
                sel_item = st.selectbox("항목 선택", item_options, key="sel_item_edit")
                sel_code = sel_item.split(" — ")[0]

                mc1, mc2 = st.columns(2)
                with mc1:
                    if st.button("수정", key="btn_edit_item", use_container_width=True):
                        st.session_state["item_edit_mode"] = True
                        st.session_state["item_edit_code"] = sel_code
                        st.rerun()
                with mc2:
                    if st.button("삭제", key="btn_del_item", use_container_width=True,
                                  type="primary"):
                        mgr.delete_item(sel_code)
                        st.success(f"[{sel_code}] 삭제 완료!")
                        st.rerun()


# ══════════════════════════════════════════════════════════════
#  메인
# ══════════════════════════════════════════════════════════════

def main():
    st.markdown("## 🎵 토요상설공연 운영 체크리스트")

    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 회차별 운영 체크",
        "📊 연간 현황 대시보드",
        "📈 항목별 통계",
        "⚙ 항목 관리",
    ])

    with tab1:
        render_tab_check()
    with tab2:
        render_tab_dashboard()
    with tab3:
        render_tab_stats()
    with tab4:
        render_tab_items()


if __name__ == "__main__":
    main()
