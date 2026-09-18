import streamlit as st
import pandas as pd
import smtplib
import ssl
import json
import time
import re
import os
import sys
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

# ----------------------------------------------------------------------
# 기본 설정
# ----------------------------------------------------------------------
def get_persistent_dir():
    """
    저장 파일(양식/계정 정보)을 둘 위치를 정한다.
    - 파이썬으로 직접 실행할 때: 이 스크립트가 있는 폴더
    - PyInstaller로 만든 exe/app 안에서 실행할 때: 실행 파일 내부의 임시 폴더가 아니라,
      운영체제의 표준 사용자 데이터 폴더(재실행해도 지워지지 않는 곳)에 저장한다.
    """
    if getattr(sys, "frozen", False):
        if sys.platform == "win32":
            base = os.getenv("APPDATA") or os.path.expanduser("~")
        elif sys.platform == "darwin":
            base = os.path.expanduser("~/Library/Application Support")
        else:
            base = os.path.expanduser("~/.local/share")
        path = os.path.join(base, "MailMergeApp")
        os.makedirs(path, exist_ok=True)
        return path
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_persistent_dir()
TEMPLATE_FILE = os.path.join(APP_DIR, "templates", "templates.json")
CONFIG_FILE = os.path.join(APP_DIR, "smtp_config.json")

os.makedirs(os.path.dirname(TEMPLATE_FILE), exist_ok=True)

st.set_page_config(page_title="메일 대량 발송 프로그램", layout="wide")

PLACEHOLDER_PATTERN = re.compile(r"\{\{(.*?)\}\}")

# ----------------------------------------------------------------------
# 디자인 테마 (파랑 / 검정 / 흰색) - .streamlit/config.toml 의 색상 테마와 짝을 이룹니다.
# ----------------------------------------------------------------------
BLUE = "#2A55FF"
INK = "#0B0B0D"

st.markdown(
    f"""
    <style>
    h1, h2, h3 {{ font-weight: 800 !important; letter-spacing: -0.02em; }}
    .step-badge {{
        display:inline-flex; align-items:center; justify-content:center;
        width:28px; height:28px; border-radius:50%;
        background:{INK}; color:#fff; font-weight:800; font-size:13px;
        margin-right:10px; flex-shrink:0;
    }}
    .step-heading {{ display:flex; align-items:center; margin: 6px 0 2px 0; }}
    .step-heading h2 {{ margin:0; font-size:22px; }}
    .hero-tag {{
        display:inline-block; background: rgba(42,85,255,0.1); color:{BLUE};
        font-weight:800; font-size:11px; letter-spacing:0.06em;
        padding:5px 12px; border-radius:999px; margin-bottom:10px;
    }}
    div[data-testid="stButton"] > button[kind="primary"] {{
        background:{BLUE}; border-color:{BLUE}; border-radius:999px; font-weight:700;
    }}
    div[data-testid="stButton"] > button[kind="secondary"] {{
        border-radius:999px; font-weight:700; border-color:{INK}; color:{INK};
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


def step_heading(number, text):
    st.markdown(
        f'<div class="step-heading"><span class="step-badge">{number}</span><h2>{text}</h2></div>',
        unsafe_allow_html=True,
    )


# ----------------------------------------------------------------------
# 템플릿 저장/불러오기
# ----------------------------------------------------------------------
def load_templates():
    if os.path.exists(TEMPLATE_FILE):
        with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_templates(templates):
    with open(TEMPLATE_FILE, "w", encoding="utf-8") as f:
        json.dump(templates, f, ensure_ascii=False, indent=2)


def load_smtp_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_smtp_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ----------------------------------------------------------------------
# 치환 로직
# ----------------------------------------------------------------------
def render_text(template_str, row_dict):
    """{{컬럼명}} 형태의 자리표시자를 엑셀 값으로 치환"""

    def repl(m):
        key = m.group(1).strip()
        return str(row_dict.get(key, m.group(0)))

    return PLACEHOLDER_PATTERN.sub(repl, template_str)


def find_placeholders(template_str):
    return sorted(set(m.strip() for m in PLACEHOLDER_PATTERN.findall(template_str)))


# ----------------------------------------------------------------------
# 메일 생성/발송
# ----------------------------------------------------------------------
def build_message(from_addr, to_addr, subject, body_text, inline_images, attachments):
    """
    본문 안에 사진을 보이게(인라인) + 별도 첨부파일 둘 다 지원하는 메시지 생성.
    body_text 안의 줄바꿈은 <br>로 변환하고, 인라인 이미지는
    본문에 {{사진1}}, {{사진2}}... 형태의 토큰이 있으면 그 위치에,
    없으면 본문 맨 아래에 순서대로 삽입한다.
    """
    msg_root = MIMEMultipart("mixed")
    msg_root["Subject"] = subject
    msg_root["From"] = from_addr
    msg_root["To"] = to_addr

    msg_related = MIMEMultipart("related")
    msg_root.attach(msg_related)

    msg_alt = MIMEMultipart("alternative")
    msg_related.attach(msg_alt)

    # 일반 텍스트(HTML을 못 여는 클라이언트를 위한 대체본)
    plain_text = re.sub(r"\{\{사진\d*\}\}", "", body_text)
    msg_alt.attach(MIMEText(plain_text, "plain", "utf-8"))

    html_body = body_text.replace("\n", "<br>\n")

    cid_map = {}
    for idx, img in enumerate(inline_images, start=1):
        cid = f"inline_image_{idx}"
        cid_map[idx] = cid

    # 본문에 {{사진N}} 토큰이 있으면 그 자리에, 없으면 본문 맨 위에 전부 붙인다
    used_token = False
    for idx, cid in cid_map.items():
        token_variants = [f"{{{{사진{idx}}}}}", "{{사진}}"]
        for token in token_variants:
            if token in html_body:
                html_body = html_body.replace(
                    token, f'<br><img src="cid:{cid}" style="max-width:500px;"><br>', 1
                )
                used_token = True
                break

    if not used_token and inline_images:
        prefix = ""
        for cid in cid_map.values():
            prefix += f'<img src="cid:{cid}" style="max-width:500px;"><br>'
        html_body = prefix + "<br>" + html_body

    msg_alt.attach(MIMEText(html_body, "html", "utf-8"))

    # 인라인 이미지 실제 삽입
    for idx, img in enumerate(inline_images, start=1):
        cid = cid_map[idx]
        img_data = img["data"]
        mime_img = MIMEImage(img_data)
        mime_img.add_header("Content-ID", f"<{cid}>")
        mime_img.add_header("Content-Disposition", "inline", filename=img["name"])
        msg_related.attach(mime_img)

    # 일반 첨부파일
    for att in attachments:
        ctype, encoding = mimetypes.guess_type(att["name"])
        if ctype is None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        part = MIMEApplication(att["data"], _subtype=subtype)
        part.add_header("Content-Disposition", "attachment", filename=att["name"])
        msg_root.attach(part)

    return msg_root


def send_via_smtp(smtp_host, smtp_port, login_user, login_pass, from_addr, to_addr, msg):
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context, timeout=20) as server:
        server.login(login_user, login_pass)
        server.sendmail(from_addr, [to_addr], msg.as_string())


# ----------------------------------------------------------------------
# 사이드바 : 발송 계정(SMTP) 설정
# ----------------------------------------------------------------------
st.sidebar.markdown(
    '<div style="font-size:11px;font-weight:800;letter-spacing:0.08em;'
    'text-transform:uppercase;color:#74767E;margin-bottom:6px;">'
    '01 — 보내는 메일 계정</div>',
    unsafe_allow_html=True,
)

saved_cfg = load_smtp_config()

smtp_host = st.sidebar.text_input("SMTP 서버", value=saved_cfg.get("smtp_host", "smtp.mailplug.co.kr"))
smtp_port = st.sidebar.number_input("SMTP 포트", value=saved_cfg.get("smtp_port", 465))
from_addr = st.sidebar.text_input("보내는 이메일 주소", value=saved_cfg.get("from_addr", ""))
login_user = st.sidebar.text_input(
    "로그인 아이디(보통 이메일 주소와 동일)", value=saved_cfg.get("login_user", "")
)
app_password = st.sidebar.text_input("앱 비밀번호(App Password)", type="password")

remember = st.sidebar.checkbox("이메일 주소/아이디 이 컴퓨터에 저장", value=bool(saved_cfg))
st.sidebar.caption(
    "⚠️ 비밀번호는 저장하지 않습니다. 메일플러그 로그인 보안 설정에서 발급한 "
    "'앱 비밀번호'를 매번 입력해 주세요. (일반 로그인 비밀번호가 아닙니다)"
)

if remember and from_addr:
    save_smtp_config(
        {
            "smtp_host": smtp_host,
            "smtp_port": smtp_port,
            "from_addr": from_addr,
            "login_user": login_user,
        }
    )

with st.sidebar.expander("메일플러그 사전 설정 방법 (최초 1회)"):
    st.markdown(
        """
1. https://login.mailplug.com 에 로그인
2. 환경설정에서 **'IMAP/SMTP 사용'** 또는 **'POP3/SMTP 사용'**을 **'사용'**으로 변경
3. 로그인 보안 설정에서 **앱 비밀번호(App Password)** 발급
4. 위 칸에 이메일 주소 + 발급받은 앱 비밀번호 입력
5. 회사(그룹웨어) 계정이라면 관리자에게 외부 SMTP 발송 허용 여부를 먼저 확인하세요.
"""
    )

st.markdown('<span class="hero-tag">MAIL MERGE TOOL</span>', unsafe_allow_html=True)
st.title("📧 메일 대량 발송 프로그램")
st.caption("엑셀 명단을 올리면 받는 사람마다 이름·회사 등을 자동으로 바꿔 넣어 개인화된 메일을 한 번에 발송합니다.")

# ----------------------------------------------------------------------
# 1단계 : 받는 사람 명단 업로드
# ----------------------------------------------------------------------
step_heading("02", "받는 사람 명단 (엑셀)")

with st.container(border=True):
    uploaded_excel = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx", "xls"])

    df = None
    if uploaded_excel is not None:
        df = pd.read_excel(uploaded_excel)
        st.write("업로드된 명단 미리보기")

        col1, col2 = st.columns(2)
        with col1:
            name_col = st.selectbox("이름 컬럼", options=df.columns, index=0)
        with col2:
            email_col = st.selectbox(
                "이메일 컬럼", options=df.columns, index=min(1, len(df.columns) - 1)
            )

        df["_선택"] = True
        edited_df = st.data_editor(
            df,
            column_config={"_선택": st.column_config.CheckboxColumn("보낼 대상")},
            use_container_width=True,
            num_rows="fixed",
        )
        selected_df = edited_df[edited_df["_선택"] == True].drop(columns=["_선택"])
        st.caption(f"선택된 받는 사람: {len(selected_df)}명 / 전체 {len(df)}명")
    else:
        selected_df = None

# ----------------------------------------------------------------------
# 2단계 : 제목/본문 양식 (저장/불러오기)
# ----------------------------------------------------------------------
step_heading("03", "메일 제목 / 본문 양식")

with st.container(border=True):
    templates = load_templates()
    template_names = list(templates.keys())

    colA, colB = st.columns([2, 1])
    with colB:
        pick = st.selectbox("저장된 양식 불러오기", options=["(선택 안 함)"] + template_names)
        if st.button("불러오기") and pick != "(선택 안 함)":
            st.session_state["subject"] = templates[pick]["subject"]
            st.session_state["body"] = templates[pick]["body"]
            st.rerun()
        if pick != "(선택 안 함)" and st.button("이 양식 삭제"):
            del templates[pick]
            save_templates(templates)
            st.rerun()

    with colA:
        if df is not None:
            cols_hint = ", ".join(f"{{{{{c}}}}}" for c in df.columns if c != "_선택")
            st.caption(f"엑셀 컬럼을 자리표시자로 쓸 수 있어요 → {cols_hint}")
        subject = st.text_input(
            "메일 제목", value=st.session_state.get("subject", "안녕하세요, {{이름}}님")
        )
        body = st.text_area(
            "메일 본문 (예: {{이름}}, {{회사}} 등 엑셀 컬럼명을 {{ }}로 감싸서 사용)",
            value=st.session_state.get(
                "body",
                "{{이름}}님, 안녕하세요.\n\n본문 내용을 입력하세요.\n\n{{사진1}}\n\n감사합니다.",
            ),
            height=220,
        )
        st.session_state["subject"] = subject
        st.session_state["body"] = body

        new_tpl_name = st.text_input("이 양식을 저장할 이름", value=pick if pick != "(선택 안 함)" else "")
        if st.button("현재 제목/본문 양식 저장", type="primary"):
            if new_tpl_name.strip():
                templates[new_tpl_name.strip()] = {"subject": subject, "body": body}
                save_templates(templates)
                st.success(f"'{new_tpl_name}' 양식을 저장했습니다. 다음에도 불러올 수 있어요.")
            else:
                st.warning("저장할 양식 이름을 입력해주세요.")

# ----------------------------------------------------------------------
# 3단계 : 사진(인라인) / 첨부파일
# ----------------------------------------------------------------------
step_heading("04", "사진 및 첨부파일")

with st.container(border=True):
    col1, col2 = st.columns(2)
    with col1:
        st.caption("본문 안에 바로 보이게 넣을 사진 (본문에 {{사진1}}, {{사진2}}... 토큰을 넣으면 그 자리에 삽입됩니다. 토큰이 없으면 본문 맨 위에 자동으로 붙습니다.)")
        inline_files = st.file_uploader(
            "본문 삽입용 사진", type=["png", "jpg", "jpeg", "gif"], accept_multiple_files=True
        )
    with col2:
        st.caption("메일에 파일로만 첨부할 항목 (문서, PDF, 표 등 아무 형식이나 가능)")
        attach_files = st.file_uploader("첨부파일", accept_multiple_files=True, key="attach")

inline_images = []
if inline_files:
    for f in inline_files:
        inline_images.append({"name": f.name, "data": f.getvalue()})

attachments = []
if attach_files:
    for f in attach_files:
        attachments.append({"name": f.name, "data": f.getvalue()})

# ----------------------------------------------------------------------
# 4단계 : 미리보기
# ----------------------------------------------------------------------
step_heading("05", "미리보기")

with st.container(border=True):
    if selected_df is not None and len(selected_df) > 0:
        preview_idx = st.selectbox(
            "미리볼 받는 사람 선택", options=list(range(len(selected_df))),
            format_func=lambda i: str(selected_df.iloc[i][name_col]),
        )
        row_dict = selected_df.iloc[preview_idx].to_dict()
        rendered_subject = render_text(subject, row_dict)
        rendered_body = render_text(body, row_dict)

        st.markdown(f"**받는사람:** {row_dict.get(email_col, '')}")
        st.markdown(f"**제목:** {rendered_subject}")
        preview_html = rendered_body.replace("\n", "<br>")
        for idx in range(1, len(inline_images) + 1):
            preview_html = preview_html.replace(f"{{{{사진{idx}}}}}", f"[사진{idx} 위치]")
        st.markdown("**본문 미리보기:**")
        st.markdown(
            f"<div style='background:{INK};color:#fff;border-radius:14px;padding:16px 18px;'>{preview_html}</div>",
            unsafe_allow_html=True,
        )
        if attachments:
            st.caption("첨부파일: " + ", ".join(a["name"] for a in attachments))
    else:
        st.info("엑셀 명단을 업로드하면 미리보기를 볼 수 있습니다.")

# ----------------------------------------------------------------------
# 5단계 : 발송
# ----------------------------------------------------------------------
step_heading("06", "발송")

with st.container(border=True):
    delay_sec = st.slider("메일 사이 발송 간격(초) — 스팸 차단 방지용", 0.0, 10.0, 2.0, 0.5)

    col1, col2 = st.columns(2)
    test_clicked = col1.button("✉️ 테스트 발송 (나에게만 1통)")
    send_clicked = col2.button("🚀 전체 발송", type="primary")

log_rows = []

def do_send(targets):
    if not from_addr or not app_password or not login_user:
        st.error("왼쪽 사이드바에서 보내는 메일 계정 정보(이메일/아이디/앱 비밀번호)를 먼저 입력해주세요.")
        return
    progress = st.progress(0)
    status = st.empty()
    total = len(targets)
    for i, row_dict in enumerate(targets, start=1):
        to_addr = row_dict.get(email_col, "").strip()
        rendered_subject = render_text(subject, row_dict)
        rendered_body = render_text(body, row_dict)
        try:
            msg = build_message(from_addr, to_addr, rendered_subject, rendered_body, inline_images, attachments)
            send_via_smtp(smtp_host, int(smtp_port), login_user, app_password, from_addr, to_addr, msg)
            log_rows.append({"받는사람": to_addr, "결과": "성공", "메시지": ""})
        except Exception as e:
            log_rows.append({"받는사람": to_addr, "결과": "실패", "메시지": str(e)})
        status.text(f"{i}/{total} 발송 처리 중... ({to_addr})")
        progress.progress(i / total)
        if i < total:
            time.sleep(delay_sec)

    result_df = pd.DataFrame(log_rows)
    st.subheader("발송 결과")
    st.dataframe(result_df, use_container_width=True)
    st.download_button(
        "결과 CSV 다운로드",
        result_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="발송결과.csv",
    )

if test_clicked:
    if not from_addr:
        st.error("사이드바에서 보내는 이메일 주소를 먼저 입력해주세요.")
    else:
        test_row = {}
        if selected_df is not None and len(selected_df) > 0:
            test_row = selected_df.iloc[0].to_dict()
        test_row[email_col if selected_df is not None else "이메일"] = from_addr
        if selected_df is None:
            test_row = {"이름": "테스트", "이메일": from_addr}
            name_col_local, email_col_local = "이름", "이메일"
        else:
            name_col_local, email_col_local = name_col, email_col
        globals()["email_col"] = email_col_local
        do_send([test_row])

if send_clicked:
    if selected_df is None or len(selected_df) == 0:
        st.error("엑셀 명단을 업로드하고 받는 사람을 선택해주세요.")
    else:
        targets = [selected_df.iloc[i].to_dict() for i in range(len(selected_df))]
        do_send(targets)
