"""KBondWatcher public demo — display-only trader UI for Streamlit Cloud."""

from __future__ import annotations

import html
import re
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parent
VIDEO_PATH = ROOT / "video" / "2026-08-20_before_close_cut.mp4"
CSS_PATH = ROOT / "static" / "styles.css"

ENGINE_VERSION = "0.3.1"
DEMO_MSG = "※ 본 웹은 데모 시연용 웹입니다. 다운로드 및 설치 문의 부탁드립니다."
COORD_IDLE = "'Set Click Position' 버튼 클릭 후 입력 좌표를 설정하세요."
CELL_RE = re.compile(r"^[A-Za-z]{1,3}\d{1,7}$")
ALLOWED = ("25-10", "25-4", "25-8", "25-5", "25-11")
MODE_OPTIONS = (
    "1 - KBond / KBond",
    "2 - KBond / Notepad",
    "3 - Forest / Notepad",
)

st.set_page_config(
    page_title="KBondWatcher",
    layout="centered",
    initial_sidebar_state="collapsed",
)


def _defaults() -> dict:
    return {
        "active_tab": "Demo",
        "profile_name": "",
        "kbond_chat_title": "[채권] 블커본드",
        "excel_workbook": "",
        "excel_sheet": "",
        "mode": MODE_OPTIONS[1],
        "sent_after": "exit",
        "instrument": "25-10",
        "required_qty": 100,
        "looking_for": "BID",
        "threshold_op": "<=",
        "threshold": 0,
        "yield_input_cell": "A1",
        "pnl_cell": "B1",
        "yield_prefix": "3",
        "send_input_x": 0.50,
        "send_input_y": 0.90,
        "msg_profile": "",
        "msg_profile_err": False,
        "msg_status": "",
        "msg_status_err": False,
        "msg_settings": "",
        "msg_settings_err": False,
        "msg_coord": COORD_IDLE,
        "msg_coord_err": False,
    }


def init_state() -> None:
    for key, value in _defaults().items():
        if key not in st.session_state:
            st.session_state[key] = value


def inject_css() -> None:
    css = _css_text()
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


@st.cache_resource
def _css_text() -> str:
    return CSS_PATH.read_text(encoding="utf-8")


def flash(slot: str, msg: str, err: bool = True) -> None:
    st.session_state[f"msg_{slot}"] = msg
    st.session_state[f"msg_{slot}_err"] = err


def render_msg(slot: str) -> None:
    msg = st.session_state.get(f"msg_{slot}", "") or ""
    err = bool(st.session_state.get(f"msg_{slot}_err", False))
    cls = "form-msg err" if msg and err else "form-msg"
    st.markdown(f'<p class="{cls}">{html.escape(msg)}</p>', unsafe_allow_html=True)


def field_label(text: str, required: bool = False) -> None:
    req = ' <span class="req">*</span>' if required else ""
    st.markdown(f'<div class="field-label">{html.escape(text)}{req}</div>', unsafe_allow_html=True)


def mode_number() -> int:
    mode = str(st.session_state.get("mode", MODE_OPTIONS[1]))
    try:
        return int(mode[0])
    except (TypeError, ValueError, IndexError):
        return 2


def click_window_label() -> str:
    n = mode_number()
    if n == 1:
        return (st.session_state.get("kbond_chat_title") or "").strip() or "KBond"
    if n in (2, 3):
        return "Notepad"
    return "—"


def fmt_threshold(n) -> str:
    try:
        return f"{int(round(float(n))):,}"
    except (TypeError, ValueError):
        return "…"


def position_check_text() -> str:
    cell = (str(st.session_state.get("pnl_cell") or "")).strip() or "…"
    op = st.session_state.get("threshold_op", "<=")
    thr = fmt_threshold(st.session_state.get("threshold", 0))
    inst = st.session_state.get("instrument") or "…"
    qty = st.session_state.get("required_qty") or "…"
    side = "매수" if st.session_state.get("looking_for") == "OFFER" else "매도"
    cmp = "이상" if op == ">=" else "이하"
    return f"만약 {cell}이/가 {thr} {cmp}이면, {inst} {qty}억 {side}"


def validate_profile() -> str:
    s = st.session_state
    if not str(s.get("profile_name", "")).strip():
        return "profile_name is required"
    if s.get("instrument") not in ALLOWED:
        return "instrument must be one of 25-10, 25-4, 25-8, 25-5, 25-11"
    if s.get("looking_for") not in ("BID", "OFFER"):
        return "looking_for must be BID or OFFER"
    try:
        qty = float(s.get("required_qty"))
    except (TypeError, ValueError):
        qty = 0
    if not (qty > 0):
        return "required_qty must be > 0"
    try:
        thr = float(s.get("threshold"))
        if thr != thr:
            return "threshold must be numeric"
    except (TypeError, ValueError):
        return "threshold must be numeric"
    if s.get("threshold_op") not in ("<=", ">="):
        return "threshold_op must be <= or >="
    wb = str(s.get("excel_workbook") or "").strip()
    if not wb:
        return "excel_workbook FullName is required"
    if "\\" not in wb and "/" not in wb:
        return "excel_workbook must be a FullName path, not a bare Name"
    if not str(s.get("excel_sheet") or "").strip():
        return "excel_sheet is required"
    y_cell = str(s.get("yield_input_cell") or "").strip()
    p_cell = str(s.get("pnl_cell") or "").strip()
    if not CELL_RE.test(y_cell):
        return "yield_input_cell must be a cell like D19"
    if not CELL_RE.test(p_cell):
        return "pnl_cell must be a cell like D19"
    if str(s.get("yield_prefix")) not in ("3", "4"):
        return "yield_prefix must be 3 or 4"
    mode = mode_number()
    if mode not in (1, 2, 3):
        return "mode must be 1, 2, or 3"
    if mode in (1, 2) and not str(s.get("kbond_chat_title") or "").strip():
        return "kbond_chat_title is required for MODE 1/2"
    loop = "exit" if mode == 1 else s.get("sent_after")
    if loop not in ("exit", "loop"):
        return "sent_after must be exit or loop"
    if mode == 1 and loop != "exit":
        return "sent_after must be exit when mode is 1"
    return ""


def validate_settings() -> str:
    s = st.session_state
    if s.get("instrument") not in ALLOWED:
        return "instrument must be one of 25-10, 25-4, 25-8, 25-5, 25-11"
    if s.get("looking_for") not in ("BID", "OFFER"):
        return "looking_for must be BID or OFFER"
    try:
        qty = float(s.get("required_qty"))
    except (TypeError, ValueError):
        qty = 0
    if not (qty > 0):
        return "required_qty must be > 0"
    try:
        thr = float(s.get("threshold"))
        if thr != thr:
            return "threshold must be numeric"
    except (TypeError, ValueError):
        return "threshold must be numeric"
    if s.get("threshold_op") not in ("<=", ">="):
        return "threshold_op must be <= or >="
    if str(s.get("yield_prefix")) not in ("3", "4"):
        return "yield_prefix must be 3 or 4"
    y_cell = str(s.get("yield_input_cell") or "").strip()
    p_cell = str(s.get("pnl_cell") or "").strip()
    if not CELL_RE.test(y_cell):
        return "yield_input_cell must be a cell like D19"
    if not CELL_RE.test(p_cell):
        return "pnl_cell must be a cell like D19"
    return ""


def render_demo() -> None:
    with st.container(border=True):
        st.markdown(
            '<div class="section-head"><h2>Demo</h2>'
            '<span class="demo-kicker">Walkthrough · 2026-08-20</span></div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<p class="muted">장외 호가 감시 → 조건 판정 → 확정 메시지 전송 루프를 '
            "장 마감 전 환경에서 기록한 시연입니다. 자동매매가 아닙니다.</p>",
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div class="demo-meta">'
            '<div><div class="lbl">Engine</div><div class="val">0.3.1</div></div>'
            '<div><div class="lbl">Loop</div><div class="val">Confirm-Judge-Act</div></div>'
            '<div><div class="lbl">Latency</div><div class="val">~1s loop</div></div>'
            "</div>",
            unsafe_allow_html=True,
        )
        if VIDEO_PATH.is_file():
            st.video(str(VIDEO_PATH), autoplay=True, loop=True, muted=True)
        else:
            st.markdown(
                f'<p class="form-msg err">{html.escape("demo video not found")}</p>',
                unsafe_allow_html=True,
            )
        st.markdown(
            '<p class="demo-caption">Before close · desktop control</p>',
            unsafe_allow_html=True,
        )


def render_profile() -> None:
    if mode_number() == 1:
        st.session_state.sent_after = "exit"

    with st.container(border=True):
        head_l, head_r = st.columns([4, 1])
        with head_l:
            st.markdown(
                '<div class="section-head"><h2>Profile - '
                '<span class="auth-label">Not Authorized</span></h2></div>',
                unsafe_allow_html=True,
            )
        with head_r:
            st.markdown('<div class="revert-anchor" id="revert-profile-anchor"></div>', unsafe_allow_html=True)
            if st.button("↺ Revert", key="btn_profile_revert"):
                flash("profile", DEMO_MSG, True)

        c1, c2 = st.columns(2)
        with c1:
            field_label("Name", True)
            st.text_input("profile_name", key="profile_name", label_visibility="collapsed")
        with c2:
            field_label("KBond Chat Title", True)
            st.text_input(
                "kbond_chat_title", key="kbond_chat_title", label_visibility="collapsed"
            )

        c3, c4 = st.columns(2)
        with c3:
            field_label("Excel Directory", True)
            wb, find = st.columns([4, 1])
            with wb:
                st.text_input(
                    "excel_workbook",
                    key="excel_workbook",
                    label_visibility="collapsed",
                    placeholder="FullName path",
                )
            with find:
                if st.button("Find", key="btn_find"):
                    flash("profile", DEMO_MSG, True)
        with c4:
            field_label("Sheet Name", True)
            st.text_input("excel_sheet", key="excel_sheet", label_visibility="collapsed")

        c5, c6 = st.columns(2)
        with c5:
            field_label("Mode", True)
            st.selectbox("mode", MODE_OPTIONS, key="mode", label_visibility="collapsed")
        with c6:
            field_label("Loop", True)
            st.selectbox(
                "sent_after",
                ("exit", "loop"),
                key="sent_after",
                label_visibility="collapsed",
                disabled=(mode_number() == 1),
                help="Mode 1 uses exit only",
            )

        foot_msg, foot_btn = st.columns([3, 2])
        with foot_btn:
            b1, b2 = st.columns(2)
            with b1:
                save = st.button("Save", key="btn_save_profile")
            with b2:
                submit = st.button("Submit", key="btn_submit_profile")
        if save or submit:
            err = validate_profile()
            flash("profile", err or DEMO_MSG, True)
        with foot_msg:
            render_msg("profile")


def render_watcher() -> None:
    st.markdown(
        '<p class="flash msg">탭을 닫거나 새로고침하면 감시가 중단됩니다.</p>',
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        st.markdown("<h2>Status</h2>", unsafe_allow_html=True)
        b1, b2, _ = st.columns([1, 1, 4])
        with b1:
            st.markdown('<div id="start-btn-anchor"></div>', unsafe_allow_html=True)
            if st.button("START", key="btn_start"):
                flash("status", DEMO_MSG, True)
        with b2:
            st.markdown('<div id="stop-btn-anchor"></div>', unsafe_allow_html=True)
            if st.button("STOP", key="btn_stop"):
                flash("status", DEMO_MSG, True)
        st.markdown(
            '<div class="status-grid">'
            '<div class="lbl">Target</div><div class="val">—</div>'
            '<div class="lbl">Threshold</div><div class="val">—</div>'
            '<div class="lbl">Status</div><div class="val">—</div>'
            '<div class="lbl">Looking For</div><div class="val">—</div>'
            '<div class="lbl">Last Quote</div><div class="val">—</div>'
            '<div class="lbl">Last Calculation</div><div class="val">—</div>'
            '<div class="lbl">Last Action</div><div class="val">—</div>'
            "</div>",
            unsafe_allow_html=True,
        )
        render_msg("status")

    with st.container(border=True):
        head_l, head_r = st.columns([4, 1])
        with head_l:
            st.markdown(
                '<div class="section-head"><h2>Settings</h2></div>',
                unsafe_allow_html=True,
            )
        with head_r:
            st.markdown('<div class="revert-anchor" id="revert-settings-anchor"></div>', unsafe_allow_html=True)
            if st.button("↺ Revert", key="btn_settings_revert"):
                flash("settings", DEMO_MSG, True)

        s1, s2 = st.columns(2)
        with s1:
            field_label("Target", True)
            st.selectbox("instrument", ALLOWED, key="instrument", label_visibility="collapsed")
        with s2:
            field_label("Quantity", True)
            st.number_input(
                "required_qty",
                min_value=1,
                step=1,
                format="%d",
                key="required_qty",
                label_visibility="collapsed",
            )

        s3, s4 = st.columns(2)
        with s3:
            field_label("Looking For", True)
            st.selectbox(
                "looking_for",
                ("BID", "OFFER"),
                key="looking_for",
                label_visibility="collapsed",
            )
        with s4:
            field_label("Threshold", True)
            op, thr = st.columns([1, 2])
            with op:
                st.selectbox(
                    "threshold_op",
                    ("<=", ">="),
                    key="threshold_op",
                    label_visibility="collapsed",
                )
            with thr:
                st.number_input(
                    "threshold",
                    step=1,
                    format="%d",
                    key="threshold",
                    label_visibility="collapsed",
                )

        s5, s6 = st.columns(2)
        with s5:
            field_label("Input Cell", True)
            st.text_input(
                "yield_input_cell", key="yield_input_cell", label_visibility="collapsed"
            )
        with s6:
            field_label("Output Cell", True)
            st.text_input("pnl_cell", key="pnl_cell", label_visibility="collapsed")

        s7, s8 = st.columns(2)
        with s7:
            field_label("Yield Prefix", True)
            st.selectbox(
                "yield_prefix", ("3", "4"), key="yield_prefix", label_visibility="collapsed"
            )
        with s8:
            field_label("Position Check")
            st.markdown(
                f'<div class="readonly-input">{html.escape(position_check_text())}</div>',
                unsafe_allow_html=True,
            )

        foot_msg, foot_btn = st.columns([3, 1])
        with foot_btn:
            if st.button("Save", key="btn_save_settings"):
                err = validate_settings()
                flash("settings", err or DEMO_MSG, True)
        with foot_msg:
            render_msg("settings")

    with st.container(border=True):
        st.markdown("<h2>Coordinate</h2>", unsafe_allow_html=True)
        d1, d2 = st.columns(2)
        with d1:
            field_label("Mode")
            st.markdown(
                f'<div class="readonly-input">{html.escape(str(mode_number()))}</div>',
                unsafe_allow_html=True,
            )
        with d2:
            field_label("Click Window")
            st.markdown(
                f'<div class="readonly-input">{html.escape(click_window_label())}</div>',
                unsafe_allow_html=True,
            )
        d3, d4 = st.columns(2)
        with d3:
            field_label("X")
            st.number_input(
                "send_input_x",
                min_value=0.0,
                max_value=1.0,
                step=0.01,
                format="%.2f",
                key="send_input_x",
                label_visibility="collapsed",
            )
        with d4:
            field_label("Y")
            st.number_input(
                "send_input_y",
                min_value=0.0,
                max_value=1.0,
                step=0.01,
                format="%.2f",
                key="send_input_y",
                label_visibility="collapsed",
            )

        foot_msg, foot_btn = st.columns([3, 2])
        with foot_btn:
            b1, b2 = st.columns(2)
            with b1:
                if st.button("Set Click Position", key="btn_calibrate"):
                    flash("coord", DEMO_MSG, True)
            with b2:
                if st.button("Test Click", key="btn_test_click"):
                    flash("coord", DEMO_MSG, True)
        with foot_msg:
            render_msg("coord")


def main() -> None:
    init_state()
    inject_css()

    st.markdown("<h1>KBondWatcher</h1>", unsafe_allow_html=True)
    st.markdown(
        '<p class="flash msg">KBondWatcher는 장외 호가를 감시하고 특정 조건에 부합할 시 '
        "확정 메시지를 전송하는 엔진입니다.</p>",
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="flash msg">자동매매 시스템이 아닌, 사람이 하던 확인-판정-액션 루프를 '
        "약 1초 이내로 줄여주는 데스크톱 제어 시스템입니다.</p>",
        unsafe_allow_html=True,
    )

    tab = st.radio(
        "section",
        ("Demo", "Profile", "Watcher"),
        horizontal=True,
        label_visibility="collapsed",
        key="active_tab",
    )
    if tab == "Demo":
        render_demo()
    elif tab == "Profile":
        render_profile()
    else:
        render_watcher()

    st.markdown(
        f'<div class="footer-meta">KBondWatcher (2026) · Engine Version {ENGINE_VERSION} · by jhacom02</div>',
        unsafe_allow_html=True,
    )


main()
