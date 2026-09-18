"""
파이썬이 설치되지 않은 컴퓨터에서도 더블클릭만으로 실행되는 exe/app을 만들기 위한
진입점(entry point)입니다. 평소에 `streamlit run app.py`로 직접 실행할 때는 이 파일을
쓸 필요가 없고, GitHub Actions가 PyInstaller로 exe/app을 빌드할 때만 사용됩니다.

streamlit-desktop-app 패키지가 내부적으로 Streamlit 서버를 띄우고, 브라우저 대신
독립된 데스크톱 창(webview)으로 열어줍니다.
"""

import os
import sys

from streamlit_desktop_app import start_desktop_app


def resource_path(relative_path: str) -> str:
    """일반 실행과 PyInstaller onefile 실행 양쪽에서 올바른 파일 경로를 찾는다."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


if __name__ == "__main__":
    start_desktop_app(
        resource_path("app.py"),
        title="메일 대량 발송 프로그램",
        width=1400,
        height=940,
        options={
            "theme.primaryColor": "#2A55FF",
            "theme.backgroundColor": "#FFFFFF",
            "theme.secondaryBackgroundColor": "#F2F3F6",
            "theme.textColor": "#0B0B0D",
        },
    )
