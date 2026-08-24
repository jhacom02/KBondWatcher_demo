"""KBondWatcher public demo — display-only trader UI for Streamlit Cloud."""

from __future__ import annotations

import html
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parent
ATTACHMENT = ROOT / "attachment"
VIDEO_PATH = ATTACHMENT / "video" / "2026-08-20_before_close_cut.mp4"
CSS_PATH = ROOT / "static" / "styles.css"
REPORT_NAME = "KBondWatcher_Technical_Report_v1.5.pdf"
REPORT_PATH = ATTACHMENT / "report" / REPORT_NAME
DOWNLOAD_NAME = "KBondWatcher-0.3.2.zip"
DOWNLOAD_PATH = ATTACHMENT / "download" / DOWNLOAD_NAME

ENGINE_VERSION = "0.3.2"
DEMO_MSG = "본 웹은 데모 시연용 웹입니다."
DEMO_BANNER = "※ 본 웹은 데모 시연용 웹입니다. 다운로드 및 설치 후 실사용 가능합니다."
COORD_IDLE = "'Set Click Position' 버튼 클릭 후 입력 좌표를 설정하세요."
ALLOWED = ("25-10", "25-4", "25-8", "25-5", "25-11")
MODE_OPTIONS = (
    "1 - KBond / KBond",
    "2 - KBond / Notepad",
    "3 - Forest / Notepad",
)
PROFILE_KEYS = (
    "profile_name",
    "kbond_chat_title",
    "excel_workbook",
    "excel_sheet",
    "mode",
    "sent_after",
)
SETTINGS_KEYS = (
    "instrument",
    "required_qty",
    "looking_for",
    "threshold_op",
    "threshold",
    "yield_input_cell",
    "pnl_cell",
    "yield_prefix",
)

st.set_page_config(
    page_title="KBondWatcher Demo",
    layout="centered",
    initial_sidebar_state="collapsed",
)


def _defaults() -> dict:
    return {
        "profile_name": "Hannah",
        "kbond_chat_title": "[채권] 블커본드",
        "excel_workbook": "C:/Users/user/Trading/Trading.xlsm",
        "excel_sheet": "Main",
        "mode": MODE_OPTIONS[1],
        "sent_after": "one-shot",
        "instrument": "25-10",
        "required_qty": 100,
        "looking_for": "BID",
        "threshold_op": "<=",
        "threshold": -1000000,
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
    st.markdown(f"<style>{CSS_PATH.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)


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


def form_footer(
    slot: str,
    buttons: list[tuple[str, str]],
    btn_w: float = 1,
) -> list[bool]:
    n = len(buttons)
    cols = st.columns([max(3, 8 - int(n * btn_w))] + [btn_w] * n)
    clicks: list[bool] = []
    for i, (label, key) in enumerate(buttons):
        with cols[i + 1]:
            clicks.append(st.button(label, key=key, use_container_width=True))
    if any(clicks):
        flash(slot, DEMO_MSG, True)
    with cols[0]:
        st.markdown('<div class="form-footer-row"></div>', unsafe_allow_html=True)
        render_msg(slot)
    return clicks


def section_head(title_html: str, revert_key: str | None = None, btn_cols: int = 2) -> bool:
    if revert_key is None:
        st.markdown(
            f'<div class="section-head"><h2>{title_html}</h2></div>',
            unsafe_allow_html=True,
        )
        return False
    cols = st.columns([7, 1])
    with cols[0]:
        st.markdown(
            f'<div class="section-head"><h2>{title_html}</h2></div>',
            unsafe_allow_html=True,
        )
    with cols[1]:
        st.markdown('<div class="revert-anchor"></div>', unsafe_allow_html=True)
        return st.button("↺ Revert", key=revert_key)


def reset_keys(keys: tuple[str, ...]) -> None:
    defaults = _defaults()
    for key in keys:
        st.session_state[key] = defaults[key]


def reset_profile() -> None:
    reset_keys(PROFILE_KEYS)


def reset_settings() -> None:
    reset_keys(SETTINGS_KEYS)


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


def render_file_section(
    title: str,
    path: Path | None = None,
    label: str | None = None,
    key: str = "",
    mime: str = "application/pdf",
    kicker: str | None = None,
    note: str | None = None,
) -> None:
    with st.container(border=True):
        if kicker:
            st.markdown(
                f'<div class="section-head"><h2>{html.escape(title)}</h2>'
                f'<span class="demo-kicker">{html.escape(kicker)}</span></div>',
                unsafe_allow_html=True,
            )
        else:
            section_head(title)
        if path is not None and path.is_file():
            st.markdown('<div class="report-dl-anchor"></div>', unsafe_allow_html=True)
            st.download_button(
                label=label or path.name,
                data=path.read_bytes(),
                file_name=path.name,
                mime=mime,
                key=key,
            )
        else:
            st.markdown(
                '<p class="file-unavailable">Not Available</p>',
                unsafe_allow_html=True,
            )
        if note:
            st.markdown(
                f'<p class="muted">{html.escape(note)}</p>',
                unsafe_allow_html=True,
            )


def render_demo() -> None:
    with st.container(border=True):
        st.markdown(
            '<div class="section-head"><h2>Watch Demo</h2>'
            '<span class="demo-kicker">Walkthrough · 2026-08-20</span></div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<p class="muted">본 영상은 웹이 아닌 엑셀을 인터페이스로 한 것이며, 루프 반복 실행한 결과입니다. (Mode: 3, Loop: loop)</p>',
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
            '<div class="demo-meta">'
            '<div><div class="lbl">Total Latency</div><div class="val">0.93s</div></div>'
            '<div><div class="lbl">Sending Latency</div><div class="val">0.75s</div></div>'
            '<div><div class="lbl">Excel Calculation Latency</div><div class="val">0.10s</div></div>'
            "</div>",
            unsafe_allow_html=True,
        )

    render_file_section(
        "Technical Report",
        REPORT_PATH,
        REPORT_NAME,
        "btn_dl_report",
        kicker="Last Update 2026-08-24",
    )
    render_file_section(
        "Download",
        DOWNLOAD_PATH,
        DOWNLOAD_NAME,
        "btn_dl_install",
        "application/zip",
        kicker="Last Update 2026-08-24",
        note="위 배포 버전은 데모 버전으로, admin으로부터 별도 key를 받아 실행해야 합니다. 문의 부탁드립니다.",
    )


def render_profile() -> None:
    if mode_number() == 1:
        st.session_state.sent_after = "one-shot"

    with st.container(border=True):
        if section_head(
            'Profile - <span class="auth-label">Not Authorized</span>',
            "btn_profile_revert",
        ):
            reset_profile()
            flash("profile", DEMO_MSG, True)
            st.rerun()

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
            st.text_input(
                "excel_workbook",
                key="excel_workbook",
                label_visibility="collapsed",
                placeholder="FullName path",
            )
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
                ("one-shot", "loop"),
                key="sent_after",
                label_visibility="collapsed",
                disabled=(mode_number() == 1),
                help="Mode 1 uses one-shot only",
            )

        form_footer(
            "profile",
            [("Save", "btn_save_profile"), ("Submit", "btn_submit_profile")],
        )


def render_watcher() -> None:
    thr = fmt_threshold(st.session_state.get("threshold", -1000000))
    qty = st.session_state.get("required_qty", 100)
    last_calc = fmt_threshold(-1458910)

    with st.container(border=True):
        section_head("Status")
        b1, b2, msg = st.columns([1, 1, 4])
        with b1:
            st.markdown('<div id="start-btn-anchor"></div>', unsafe_allow_html=True)
            if st.button("START", key="btn_start", use_container_width=True):
                flash("status", DEMO_MSG, True)
        with b2:
            st.markdown('<div id="stop-btn-anchor"></div>', unsafe_allow_html=True)
            if st.button("STOP", key="btn_stop", use_container_width=True):
                flash("status", DEMO_MSG, True)
        with msg:
            st.markdown('<div class="form-footer-row"></div>', unsafe_allow_html=True)
            render_msg("status")
        st.markdown(
            '<div class="status-grid">'
            '<div class="lbl">Status</div><div class="val">STOPPED</div>'
            f'<div class="lbl">Target</div><div class="val">{html.escape(str(st.session_state.get("instrument", "25-10")))}</div>'
            f'<div class="lbl">Quantity</div><div class="val">{html.escape(str(qty))}</div>'
            f'<div class="lbl">Looking For</div><div class="val">{html.escape(str(st.session_state.get("looking_for", "BID")))}</div>'
            f'<div class="lbl">Threshold</div><div class="val">{html.escape(thr)}</div>'
            '<div class="lbl">Last Quote</div><div class="val">25-10 74+</div>'
            f'<div class="lbl">Last Calculation</div><div class="val">{html.escape(last_calc)}</div>'
            '<div class="lbl">Last Action</div><div class="val">(16:05:16) Message Sent: 25-10 74- ㅎㅈ</div>'
            "</div>",
            unsafe_allow_html=True,
        )
    with st.container(border=True):
        if section_head("Settings", "btn_settings_revert"):
            reset_settings()
            flash("settings", DEMO_MSG, True)
            st.rerun()

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
            op, thr_col = st.columns([1, 2])
            with op:
                st.selectbox(
                    "threshold_op",
                    ("<=", ">="),
                    key="threshold_op",
                    label_visibility="collapsed",
                )
            with thr_col:
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

        form_footer("settings", [("Save", "btn_save_settings")])

    with st.container(border=True):
        section_head("Calibration")
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

        form_footer(
            "coord",
            [
                ("Set Click Position", "btn_calibrate"),
                ("Test Click", "btn_test_click"),
            ],
            btn_w=1.9,
        )


def main() -> None:
    init_state()
    inject_css()

    st.markdown("<h1>KBondWatcher Demo</h1>", unsafe_allow_html=True)
    st.markdown(
        '<p class="flash msg">KBondWatcher는 장외 호가를 감시하고 특정 조건에 부합할 시 확정 메시지를 전송하는 엔진입니다.</p>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="flash msg">자동매매 시스템이 아닌, 사람이 하던 확인-판정-액션 루프를 약 1초 이내로 줄여주는 데스크톱 제어 시스템입니다.</p>',
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<p class="flash msg err">{html.escape(DEMO_BANNER)}</p>',
        unsafe_allow_html=True,
    )

    demo_tab, profile_tab, watcher_tab = st.tabs(["Demo", "Profile", "Watcher"])
    with demo_tab:
        render_demo()
    with profile_tab:
        render_profile()
    with watcher_tab:
        render_watcher()

    st.markdown(
        f'<div class="footer-meta">KBondWatcher (2026) · Engine Version {ENGINE_VERSION} · by jhacom02</div>',
        unsafe_allow_html=True,
    )


main()
