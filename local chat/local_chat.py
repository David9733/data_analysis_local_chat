"""
Windows용 파이썬 GUI 챗봇 - 데이터 실무 코치
- 기본 모듈(tkinter, urllib, json, threading, csv) 및 pypdf, openpyxl 지원
- Ollama 로컬 API (http://localhost:11434/api/chat) 연동
- PDF, CSV, Excel, 텍스트/코드 파일 첨부 및 분석 지원
"""

import csv
import json
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import urllib.error
import urllib.request

# 선택적 외부 라이브러리 (설치 시 확장 지원)
try:
    import pypdf
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# ==========================================
# 1. 상단 설정 변수
# ==========================================

# 사용할 Ollama 모델명
MODEL = "gemma4:e2b"

# Ollama 로컬 API 엔드포인트 URL
OLLAMA_API_URL = "http://localhost:11434/api/chat"

# 시스템 프롬프트 (챗봇의 역할 및 행동 지침)
SYSTEM = """# 시스템 프롬프트

## 1) 누가 사용하는 챗봇인가?

데이터 분석을 공부하고 있으며, 향후 데이터 분석가 또는 데이터를 활용하는 사무직 업무를 준비하는 취업준비생이 사용합니다.

## 2) 챗봇의 역할은?

당신은 데이터를 실제 업무에 적용할 수 있도록 돕는 실무형 데이터 분석 컨설턴트입니다.
사용자가 업무 상황이나 문제를 제시하면 어떤 데이터를 확인해야 하는지, 어떤 KPI와 분석 방법을 사용할지, 분석 결과를 업무 의사결정에 어떻게 연결할지 안내합니다.
단순히 Python이나 SQL 문법을 알려주는 것보다 "왜 이 분석을 하는가 → 무엇을 확인하는가 → 결과를 어떻게 활용하는가"를 중심으로 설명합니다.
가능하면 실제 기업의 업무에서 발생할 법한 상황을 기준으로 설명합니다.

## 3) 말투

정중하고 명확한 "~합니다"체를 사용합니다.
어려운 데이터 분석 용어는 실무자가 이해할 수 있도록 쉽게 설명합니다.
불필요하게 어려운 전문용어를 사용하지 않습니다.

## 4) 답변 길이

기본적으로 5문장 이내로 간결하게 답변합니다.
다만 코드, 분석 절차, 표, 예시 등이 필요한 경우에는 이해에 필요한 만큼 추가할 수 있습니다.
항상 핵심 내용을 먼저 제시합니다.

## 5) 절대 하지 말 것

* 근거 없이 수치, 기업 사례, 업계 동향, 분석 결과를 만들어내지 않습니다.
* 사용자가 제공하지 않은 데이터를 실제로 분석했다고 가정하지 않습니다.
* 존재하지 않는 데이터나 결과를 만들어내지 않습니다.
* 단순히 Python/SQL 문법만 설명하고 끝내지 않습니다.
* 공부를 위한 이론 설명에만 머무르지 않고 실제 업무 적용 방법을 함께 제시합니다.
* 확실하지 않은 내용을 사실처럼 단정하지 않습니다.

## 6) 모르는 내용일 시에는?

확실하지 않은 경우 먼저 "확실하지 않습니다" 또는 "현재 제공된 정보만으로는 판단하기 어렵습니다"라고 밝힙니다.
추측이 필요한 경우 반드시 "추측입니다"라고 명확하게 표시합니다.
추가 정보가 필요하다면 어떤 정보가 필요한지 구체적으로 질문합니다.
최신 정보가 필요한 경우에는 최신 자료를 확인해야 한다고 안내합니다.

## 7) 챗봇 이름

데이터 실무 코치

## 8) 첫 인사

안녕하세요. 데이터 실무 코치입니다.
데이터 분석을 실제 업무에 어떻게 적용할지 함께 고민해드립니다.
업무 상황이나 데이터가 있다면 "이 상황에서 어떤 분석을 해야 하지?"처럼 편하게 질문해주세요.
PDF, CSV, Excel, 텍스트 문서가 있다면 [파일 첨부] 버튼으로 언제든 전달해주세요.

## 핵심 원칙

모든 답변은 가능하면 다음 흐름으로 생각합니다.

업무 문제 → 필요한 데이터 → KPI/지표 → 분석 방법 → 분석 결과 → 업무 의사결정

데이터 분석 자체가 목적이 아니라, 데이터를 활용해 실제 업무의 문제를 해결하고 의사결정을 지원하는 것을 목표로 합니다."""

# 첫 인사 메시지
GREETING_MESSAGE = (
    "안녕하세요. 데이터 실무 코치입니다.\n"
    "데이터 분석을 실제 업무에 어떻게 적용할지 함께 고민해드립니다.\n"
    "업무 상황이나 데이터가 있다면 \"이 상황에서 어떤 분석을 해야 하지?\"처럼 편하게 질문해주세요.\n"
    "💡 PDF 문서, CSV/Excel 데이터, 코드 파일은 [파일 첨부] 버튼으로 읽힐 수 있습니다."
)


# ==========================================
# 2. 파일 텍스트 추출 헬퍼 함수
# ==========================================

def extract_file_content(file_path: str, max_chars: int = 15000) -> tuple[str, str]:
    """
    다양한 파일 형식(PDF, CSV, Excel, TXT, 코드 등)에서 텍스트를 추출합니다.
    반환: (추출된 본문 텍스트, 사용자 표시용 메타 정보 문자열)
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError("지정된 파일을 찾을 수 없습니다.")

    filename = os.path.basename(file_path)
    file_size = os.path.getsize(file_path)
    ext = os.path.splitext(filename)[1].lower()

    # 사람이 읽기 쉬운 파일 크기 표시 (KB, MB)
    if file_size < 1024:
        size_str = f"{file_size} B"
    elif file_size < 1024 * 1024:
        size_str = f"{file_size / 1024:.1f} KB"
    else:
        size_str = f"{file_size / (1024 * 1024):.1f} MB"

    # 1. PDF 파일 (.pdf)
    if ext == ".pdf":
        if not PYPDF_AVAILABLE:
            raise RuntimeError("PDF 파일을 읽으려면 'pip install pypdf' 설치가 필요합니다.")

        reader = pypdf.PdfReader(file_path)
        total_pages = len(reader.pages)
        pages_content = []
        char_count = 0
        truncated = False

        for i, page in enumerate(reader.pages):
            page_text = page.extract_text() or ""
            clean_text = page_text.strip()
            if clean_text:
                pages_content.append(f"--- [페이지 {i+1}/{total_pages}] ---\n{clean_text}")
                char_count += len(clean_text)
                if char_count >= max_chars:
                    truncated = True
                    pages_content.append(f"\n(※ 전체 {total_pages}페이지 중 용량 제한으로 인해 앞부분 약 {i+1}페이지까지 발췌되었습니다.)")
                    break

        if not pages_content:
            return "(PDF에서 추출 가능한 텍스트가 없습니다. 스캔된 이미지 문서일 수 있습니다.)", f"{filename} ({size_str}, {total_pages}페이지)"

        full_text = "\n\n".join(pages_content)
        meta = f"{filename} ({size_str}, 총 {total_pages}페이지" + (", 일부 발췌)" if truncated else ")")
        return full_text, meta

    # 2. CSV / TSV 데이터 파일 (.csv, .tsv)
    elif ext in [".csv", ".tsv"]:
        delimiter = "\t" if ext == ".tsv" else ","
        raw_rows = []

        # UTF-8, CP949(한국어 윈도우 기본), EUC-KR 순차 시도
        for enc in ["utf-8-sig", "utf-8", "cp949", "euc-kr"]:
            try:
                with open(file_path, "r", encoding=enc, newline="") as f:
                    r = csv.reader(f, delimiter=delimiter)
                    raw_rows = [row for row in r]
                break
            except (UnicodeDecodeError, Exception):
                continue

        if not raw_rows:
            return "(데이터 파일이 비어 있거나 읽을 수 없습니다.)", f"{filename} ({size_str})"

        header = raw_rows[0] if raw_rows else []
        total_rows = max(0, len(raw_rows) - 1)
        sample_rows = raw_rows[1:41]  # 최대 상위 40행 미리보기

        table_lines = [
            f"[CSV 데이터 요약] 전체 행 수: {total_rows:,}개, 컬럼 수: {len(header)}개",
            f"컬럼 목록: {', '.join(header)}",
            "\n[데이터 미리보기 (상위 최대 40행)]",
            " | ".join(header),
            "-" * min(80, max(20, len(" | ".join(header))))
        ]
        for row in sample_rows:
            table_lines.append(" | ".join(row))
        if total_rows > 40:
            table_lines.append(f"... (외 {total_rows - 40:,}개 행 생략)")

        meta = f"{filename} ({size_str}, {total_rows:,}행 {len(header)}열)"
        return "\n".join(table_lines), meta

    # 3. Excel 파일 (.xlsx, .xls)
    elif ext in [".xlsx", ".xls"]:
        if not OPENPYXL_AVAILABLE:
            raise RuntimeError("Excel 파일을 읽으려면 'pip install openpyxl' 설치가 필요합니다.")

        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheets_output = []

        for sname in wb.sheetnames[:3]:  # 최대 3개 시트 분석
            sheet = wb[sname]
            all_rows = list(sheet.iter_rows(values_only=True))
            if not all_rows:
                continue

            valid_rows = [row for row in all_rows if any(cell is not None for cell in row)]
            if not valid_rows:
                continue

            header = [str(c) if c is not None else "" for c in valid_rows[0]]
            data_rows = valid_rows[1:31]  # 시트당 상위 최대 30행
            total_r = max(0, len(valid_rows) - 1)

            sheet_text = [
                f"[시트: {sname}] (총 {total_r:,}행, {len(header)}열)",
                f"컬럼: {', '.join(header)}",
                " | ".join(header),
                "-" * min(80, max(20, len(" | ".join(header))))
            ]
            for row in data_rows:
                row_str = [str(c) if c is not None else "" for c in row]
                sheet_text.append(" | ".join(row_str))
            if total_r > 30:
                sheet_text.append(f"... (외 {total_r - 30:,}개 행 생략)")

            sheets_output.append("\n".join(sheet_text))

        meta = f"{filename} ({size_str}, 시트 {len(wb.sheetnames)}개)"
        return "\n\n".join(sheets_output) if sheets_output else "(시트에 데이터가 없습니다.)", meta

    # 4. 일반 텍스트 및 코드/설정 파일 (.txt, .md, .py, .sql, .json, .log 등)
    else:
        text_content = ""
        for enc in ["utf-8-sig", "utf-8", "cp949", "euc-kr"]:
            try:
                with open(file_path, "r", encoding=enc) as f:
                    text_content = f.read()
                break
            except (UnicodeDecodeError, Exception):
                continue

        if not text_content:
            return "(파일 내용이 비어 있거나 지원하지 않는 인코딩입니다.)", f"{filename} ({size_str})"

        truncated = False
        if len(text_content) > max_chars:
            text_content = text_content[:max_chars] + f"\n\n(※ 파일 내용이 방대하여 앞부분 약 {max_chars:,}자까지만 발췌되었습니다.)"
            truncated = True

        meta = f"{filename} ({size_str}" + (", 일부 발췌)" if truncated else ")")
        return text_content, meta


# ==========================================
# 3. GUI 애플리케이션 클래스
# ==========================================

class LocalChatApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("데이터 실무 코치 - 로컬 챗봇")
        self.root.geometry("740x820")
        self.root.minsize(580, 560)

        # 현재 첨부된 파일 정보 (경로, 본문 텍스트, 메타 라벨)
        self.attached_file_path = None
        self.attached_content = ""
        self.attached_meta = ""

        # 대화 이력(messages) 초기화 - 시스템 프롬프트 포함
        self.messages = []
        self._reset_messages()

        # UI 스타일 및 컴포넌트 구성
        self._setup_style()
        self._create_widgets()

        # 시작 시 첫 인사말 출력
        self._display_greeting()

    def _reset_messages(self):
        """대화 기록을 시스템 프롬프트만 있는 상태로 초기화합니다."""
        self.messages = [{"role": "system", "content": SYSTEM}]

    def _setup_style(self):
        """기본 폰트 및 테마 스타일을 설정합니다."""
        self.font_main = ("맑은 고딕", 10)
        self.font_bold = ("맑은 고딕", 10, "bold")
        self.font_title = ("맑은 고딕", 11, "bold")
        self.font_status = ("맑은 고딕", 9)
        self.root.configure(bg="#f4f5f7")

    def _create_widgets(self):
        """윈도우 위젯들을 배치합니다."""
        # 1. 상단 정보 헤더 프레임
        header_frame = tk.Frame(self.root, bg="#ffffff", bd=1, relief=tk.SOLID)
        header_frame.pack(fill=tk.X, padx=12, pady=(12, 6))

        title_label = tk.Label(
            header_frame,
            text="📊 데이터 실무 코치 (로컬 AI 챗봇)",
            font=self.font_title,
            bg="#ffffff",
            fg="#202124",
            anchor="w",
            padx=12,
            pady=10,
        )
        title_label.pack(side=tk.LEFT)

        model_badge = tk.Label(
            header_frame,
            text=f"모델: {MODEL}",
            font=self.font_status,
            bg="#e8f0fe",
            fg="#1a73e8",
            padx=8,
            pady=4,
            relief=tk.GROOVE,
        )
        model_badge.pack(side=tk.RIGHT, padx=12, pady=10)

        # 2. 대화 기록 표시 창 (ScrolledText)
        chat_frame = tk.Frame(self.root, bg="#f4f5f7")
        chat_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=6)

        self.chat_area = scrolledtext.ScrolledText(
            chat_frame,
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=self.font_main,
            bg="#ffffff",
            fg="#202124",
            relief=tk.SOLID,
            bd=1,
            padx=12,
            pady=12,
        )
        self.chat_area.pack(fill=tk.BOTH, expand=True)

        # 텍스트 태그 스타일 설정
        self.chat_area.tag_config("user_header", font=self.font_bold, foreground="#1a73e8")
        self.chat_area.tag_config("user_attachment", font=self.font_bold, foreground="#0b57d0", lmargin1=10, lmargin2=10)
        self.chat_area.tag_config("user_body", font=self.font_main, foreground="#202124", lmargin1=10, lmargin2=10)
        self.chat_area.tag_config("coach_header", font=self.font_bold, foreground="#137333")
        self.chat_area.tag_config("coach_body", font=self.font_main, foreground="#202124", lmargin1=10, lmargin2=10)
        self.chat_area.tag_config("error_msg", font=self.font_bold, foreground="#d93025", lmargin1=10, lmargin2=10)
        self.chat_area.tag_config("system_notice", font=self.font_status, foreground="#5f6368", justify=tk.CENTER)

        # 3. 파일 첨부 상태 바 (파일 첨부 시에만 활성화)
        self.attachment_bar = tk.Frame(self.root, bg="#e8f0fe", bd=1, relief=tk.SOLID)

        self.attachment_label = tk.Label(
            self.attachment_bar,
            text="",
            font=self.font_status,
            bg="#e8f0fe",
            fg="#174ea6",
            anchor="w",
            padx=8,
            pady=4,
        )
        self.attachment_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.remove_file_btn = tk.Button(
            self.attachment_bar,
            text="✕ 첨부 취소",
            command=self.clear_attachment,
            font=self.font_status,
            bg="#ffffff",
            fg="#d93025",
            activebackground="#fce8e6",
            relief=tk.FLAT,
            padx=6,
            pady=2,
            cursor="hand2",
        )
        self.remove_file_btn.pack(side=tk.RIGHT, padx=6, pady=3)

        # 4. 사용자 입력 및 버튼 영역
        input_container = tk.Frame(self.root, bg="#f4f5f7")
        input_container.pack(fill=tk.X, padx=12, pady=(4, 6))

        # 입력창 (다중 행 입력 지원, Enter로 전송, Shift+Enter로 줄바꿈)
        self.input_text = tk.Text(
            input_container,
            height=4,
            wrap=tk.WORD,
            font=self.font_main,
            bg="#ffffff",
            fg="#202124",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=8,
        )
        self.input_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
        self.input_text.bind("<Return>", self._on_enter_press)
        self.input_text.bind("<Shift-Return>", self._on_shift_enter_press)
        self.input_text.focus_set()

        # 버튼 프레임
        button_frame = tk.Frame(input_container, bg="#f4f5f7")
        button_frame.pack(side=tk.RIGHT, fill=tk.Y)

        self.send_button = tk.Button(
            button_frame,
            text="전송\n(Enter)",
            command=self.send_message,
            font=self.font_bold,
            bg="#1a73e8",
            fg="#ffffff",
            activebackground="#1557b0",
            activeforeground="#ffffff",
            relief=tk.FLAT,
            padx=14,
            pady=5,
            cursor="hand2",
            width=11,
        )
        self.send_button.pack(fill=tk.X, pady=(0, 3))

        self.attach_button = tk.Button(
            button_frame,
            text="📁 파일 첨부",
            command=self.attach_file,
            font=self.font_main,
            bg="#ffffff",
            fg="#1a73e8",
            activebackground="#f1f3f4",
            activeforeground="#1557b0",
            relief=tk.SOLID,
            bd=1,
            padx=14,
            pady=3,
            cursor="hand2",
            width=11,
        )
        self.attach_button.pack(fill=tk.X, pady=(0, 3))

        self.new_chat_button = tk.Button(
            button_frame,
            text="새 대화",
            command=self.new_chat,
            font=self.font_main,
            bg="#e8eaed",
            fg="#3c4043",
            activebackground="#dadce0",
            activeforeground="#202124",
            relief=tk.FLAT,
            padx=14,
            pady=3,
            cursor="hand2",
            width=11,
        )
        self.new_chat_button.pack(fill=tk.X)

        # 5. 하단 상태 표시줄
        self.status_var = tk.StringVar(value="준비 완료")
        status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            font=self.font_status,
            bg="#e8eaed",
            fg="#5f6368",
            anchor="w",
            padx=12,
            pady=4,
        )
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def attach_file(self):
        """파일 선택 다이얼로그를 열고 텍스트를 추출해 준비합니다."""
        file_types = [
            ("모든 지원 파일", "*.pdf;*.txt;*.csv;*.tsv;*.xlsx;*.xls;*.py;*.sql;*.json;*.md;*.log"),
            ("PDF 문서 (*.pdf)", "*.pdf"),
            ("데이터 파일 (*.csv, *.tsv, *.xlsx, *.xls)", "*.csv;*.tsv;*.xlsx;*.xls"),
            ("텍스트/코드 파일 (*.txt, *.md, *.py, *.sql, *.json)", "*.txt;*.md;*.py;*.sql;*.json;*.log"),
            ("모든 파일 (*.*)", "*.*"),
        ]

        chosen_path = filedialog.askopenfilename(
            title="챗봇에 전달할 파일 선택 (PDF, CSV, Excel, TXT 등)",
            filetypes=file_types,
        )
        if not chosen_path:
            return

        try:
            self.status_var.set("파일 텍스트 추출 중...")
            content, meta = extract_file_content(chosen_path)

            self.attached_file_path = chosen_path
            self.attached_content = content
            self.attached_meta = meta

            # 첨부 바 표시
            self.attachment_label.configure(text=f"📎 첨부됨: {meta}")
            self.attachment_bar.pack(fill=tk.X, padx=12, pady=(0, 4), before=self.input_text.master)
            self.status_var.set(f"파일 첨부 완료: {meta}")
            self.input_text.focus_set()

        except Exception as e:
            messagebox.showerror("파일 읽기 오류", f"파일을 읽는 중 오류가 발생했습니다:\n{str(e)}")
            self.clear_attachment()
            self.status_var.set("파일 읽기 실패")

    def clear_attachment(self):
        """현재 선택된 첨부 파일을 취소합니다."""
        self.attached_file_path = None
        self.attached_content = ""
        self.attached_meta = ""
        self.attachment_bar.pack_forget()
        self.status_var.set("첨부가 취소되었습니다.")

    def _append_to_chat(self, text: str, tag: str = None):
        """대화창에 텍스트를 추가하고 맨 아래로 스크롤합니다."""
        self.chat_area.configure(state=tk.NORMAL)
        if tag:
            self.chat_area.insert(tk.END, text, tag)
        else:
            self.chat_area.insert(tk.END, text)
        self.chat_area.configure(state=tk.DISABLED)
        self.chat_area.see(tk.END)

    def _display_greeting(self):
        """첫 인사를 대화창에 출력합니다."""
        self._append_to_chat("\n[데이터 실무 코치]\n", "coach_header")
        self._append_to_chat(f"{GREETING_MESSAGE}\n\n", "coach_body")

    def _on_enter_press(self, event):
        """Enter 키 입력 시 메시지를 전송합니다."""
        if not event.state & 0x0001:  # Shift 키가 눌려있지 않은 경우
            self.send_message()
            return "break"  # 기본 줄바꿈 동작 방지

    def _on_shift_enter_press(self, event):
        """Shift + Enter 키 입력 시 줄바꿈을 허용합니다."""
        return None

    def send_message(self):
        """사용자 입력 메시지와 첨부 파일을 처리하고 Ollama API를 비동기로 호출합니다."""
        user_input = self.input_text.get("1.0", tk.END).strip()

        # 파일도 없고 입력 텍스트도 없는 경우 무시
        if not user_input and not self.attached_file_path:
            return

        # 파일만 첨부하고 질문을 비워둔 경우 기본 분석 요청 문구 부여
        if self.attached_file_path and not user_input:
            user_input = "이 파일의 핵심 내용과 데이터 분석 관점에서의 주요 시사점 및 업무 활용 방안을 정리해주세요."

        # 입력창 비우기
        self.input_text.delete("1.0", tk.END)

        # 1. UI 대화창 표시
        self._append_to_chat("[사용자]\n", "user_header")
        if self.attached_file_path:
            self._append_to_chat(f"📎 [첨부 파일] {self.attached_meta}\n", "user_attachment")
        self._append_to_chat(f"{user_input}\n\n", "user_body")

        # 2. 모델에 전달할 프롬프트 구성
        if self.attached_file_path:
            prompt_for_llm = (
                f"[첨부 파일 정보: {self.attached_meta}]\n"
                f"--- [파일 내용 시작] ---\n"
                f"{self.attached_content}\n"
                f"--- [파일 내용 끝] ---\n\n"
                f"[사용자 질문]\n"
                f"{user_input}"
            )
        else:
            prompt_for_llm = user_input

        # 대화 이력에 사용자 메시지 추가
        self.messages.append({"role": "user", "content": prompt_for_llm})

        # 첨부 상태 정리
        self.clear_attachment()

        # UI 비활성화 및 상태 변경
        self.send_button.configure(state=tk.DISABLED, bg="#dadce0", text="생성 중...")
        self.attach_button.configure(state=tk.DISABLED)
        self.status_var.set(f"모델({MODEL}) 응답 대기 중...")

        # 백그라운드 스레드에서 API 호출
        threading.Thread(target=self._call_ollama_api, daemon=True).start()

    def _call_ollama_api(self):
        """Ollama 로컬 API (/api/chat)를 호출하는 백그라운드 작업 함수입니다."""
        payload = {
            "model": MODEL,
            "messages": self.messages,
            "stream": False,
            "think": False,
        }

        try:
            req_data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            req = urllib.request.Request(
                OLLAMA_API_URL,
                data=req_data,
                headers={"Content-Type": "application/json; charset=utf-8"},
                method="POST",
            )

            with urllib.request.urlopen(req, timeout=180) as response:
                status_code = response.getcode()
                raw_response = response.read().decode("utf-8")

                if status_code != 200:
                    raise Exception(f"HTTP 상태 코드 {status_code} 오류가 반환되었습니다.")

                data = json.loads(raw_response)
                reply_content = data.get("message", {}).get("content", "").strip()

                if not reply_content:
                    reply_content = "(모델로부터 전달받은 답변 내용이 없습니다.)"

                self.root.after(0, self._on_api_success, reply_content)

        except urllib.error.URLError:
            self.root.after(0, self._on_api_error, "Ollama가 실행 중인지 확인하세요")
        except Exception as e:
            self.root.after(0, self._on_api_error, f"오류가 발생했습니다: {str(e)}")

    def _on_api_success(self, reply: str):
        """API 호출 성공 시 UI 업데이트 및 대화 이력 저장"""
        self._append_to_chat("[데이터 실무 코치]\n", "coach_header")
        self._append_to_chat(f"{reply}\n\n", "coach_body")

        # 어시스턴트 답변 이력 누적
        self.messages.append({"role": "assistant", "content": reply})

        self._restore_ui_state("준비 완료")

    def _on_api_error(self, error_message: str):
        """API 호출 실패 시 에러 메시지 표시 및 대화 이력 복구"""
        self._append_to_chat(f"[오류 안내] {error_message}\n\n", "error_msg")

        # 실패한 사용자 메시지 롤백
        if self.messages and self.messages[-1]["role"] == "user":
            self.messages.pop()

        self._restore_ui_state(f"오류: {error_message}")

    def _restore_ui_state(self, status_text: str):
        """버튼 활성화 및 상태 텍스트를 복구합니다."""
        self.send_button.configure(
            state=tk.NORMAL,
            bg="#1a73e8",
            text="전송\n(Enter)",
        )
        self.attach_button.configure(state=tk.NORMAL)
        self.status_var.set(status_text)
        self.input_text.focus_set()

    def new_chat(self):
        """대화 기록을 비우고 새로운 대화를 시작합니다."""
        self._reset_messages()
        self.clear_attachment()

        # 대화창 클리어
        self.chat_area.configure(state=tk.NORMAL)
        self.chat_area.delete("1.0", tk.END)
        self.chat_area.configure(state=tk.DISABLED)

        # 새 대화 구분선 및 첫 인사말 출력
        self._append_to_chat("--- 새 대화가 시작되었습니다 ---\n", "system_notice")
        self._display_greeting()

        # 입력창 초기화
        self.input_text.delete("1.0", tk.END)
        self.status_var.set("새 대화가 시작되었습니다.")
        self.input_text.focus_set()


# ==========================================
# 4. 프로그램 진입점
# ==========================================

if __name__ == "__main__":
    root = tk.Tk()
    app = LocalChatApp(root)
    root.mainloop()
