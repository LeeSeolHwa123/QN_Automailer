import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from restored import AutomailerMain
from settings_manager import (
    DEFAULT_SETTINGS_FILE,
    load_settings, save_settings, validate_settings,
    get_effective_execution_date, get_effective_date_range
)

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

# 진행 단계별 상세 설명
STEP_DETAILS = {
    0: "SAP 트랜잭션에서 오늘 날짜 기준 전체 QN 데이터를 조회하고 있습니다.",
    1: "기존에 발송했던 보고서 메일을 검색 중입니다...",
    2: "새로 등록된 C-Type QN 항목을 분석 중입니다...",
    3: "삭제된 C-Type QN 항목을 분석 중입니다...",
    4: "새로 등록된 Q-Type QN 항목을 분석 중입니다...",
    5: "삭제된 Q-Type QN 항목을 분석 중입니다...",
    6: "각 QN의 대응 기한과 현재 진행 상태(D3/D5/D8)를 확인 중입니다...",
    7: "KPI 점수를 계산 중입니다...",
    8: "최신 정보로 보고서 메일 본문을 생성 중입니다...",
    9: "Outlook에서 보고서를 열고 있습니다. 확인 후 발송해주세요."
}


class QNAutomailerGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("QN_AutoMailer")
        self.geometry("1200x850")
        self.minsize(960, 640)
        self.resizable(True, True)
        self.configure(fg_color="#000000")
        self.base_width = 1200
        self.base_height = 850
        self.widget_scaling = 1.0
        
        self.settings = load_settings()
        self.current_step = -1
        self.step_states = {}
        self.progress_track_width = 0
        self.progress_ratio = 0.0
        self.spacing_scaling = 1.0
        
        # Grid configuration
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Scrollable viewport to prevent clipping on any monitor
        container = ctk.CTkFrame(self, fg_color="#000000")
        container.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        container.grid_rowconfigure(1, weight=1)
        container.grid_columnconfigure(0, weight=1)
        self.container = container
        
        # Top: Input section (full width)
        self.input_frame = ctk.CTkFrame(container, fg_color="transparent")
        self.input_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.input_frame.grid_columnconfigure(0, weight=1)
        
        # Bottom: Split layout (progress left, buttons right)
        split_frame = ctk.CTkFrame(container, fg_color="transparent")
        split_frame.grid(row=1, column=0, sticky="nsew")
        split_frame.grid_rowconfigure(0, weight=1)
        split_frame.grid_columnconfigure((0, 1), weight=1)
        
        # Left: Progress
        self.progress_frame = ctk.CTkFrame(split_frame, fg_color="transparent")
        self.progress_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.progress_frame.grid_rowconfigure(3, weight=1)
        self.progress_frame.grid_columnconfigure(0, weight=1)
        
        # Right: Buttons
        self.button_frame = ctk.CTkFrame(split_frame, fg_color="transparent")
        self.button_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        self.button_frame.grid_rowconfigure(7, weight=1)
        self.button_frame.grid_columnconfigure(0, weight=1)
        
        self.build_ui()
        self.bind("<Configure>", self._on_window_resize)
    
    def build_ui(self):
        """UI 구성"""
        self.build_input_section()
        self.build_progress_section()
        self.build_buttons()

    def _build_runner(self):
        """Create an AutomailerMain instance wired to GUI callbacks."""
        def handle_progress(step_idx):
            def _update():
                self.update_step_progress(step_idx)
                if not self.step_messages.get(step_idx):
                    default_msg = STEP_DETAILS.get(step_idx, "Processing...")
                    self.add_step_message(step_idx, default_msg)

            self.after(0, _update)

        def handle_message(step_idx, message):
            def _update():
                self.update_step_progress(step_idx)
                text = message or STEP_DETAILS.get(step_idx, "Processing...")
                self.add_step_message(step_idx, text)

            self.after(0, _update)

        return AutomailerMain(
            settings_path=Path(DEFAULT_SETTINGS_FILE),
            progress_callback=handle_progress,
            message_callback=handle_message,
        )

    def _run_single_export(self, target: str, success_message: str, context: str):
        """Execute a single SAP helper through the orchestrator."""
        try:
            runner = self._build_runner()
            runner.run_single_export(target)
            self.after(0, lambda: messagebox.showinfo("성공", success_message))
        except Exception as exc:
            import traceback

            error_msg = f"{context} 실행 중 오류 발생:\n{exc}"
            traceback.print_exc()
            self.after(0, lambda: messagebox.showerror("오류", error_msg))
    
    def build_input_section(self):
        """입력 섹션"""
        # Title - centered
        title = ctk.CTkLabel(
            self.input_frame, text="QN_AutoMailer", text_color="white",
            font=("Inter", 20, "bold")
        )
        title.pack(pady=(0, 6), anchor="center")
        
        # Input container
        container = ctk.CTkFrame(self.input_frame, fg_color="#323232", corner_radius=2)
        container.pack(fill="x", pady=(0, 4))
        container.grid_columnconfigure(1, weight=1)
        
        # 실행 날짜
        ctk.CTkLabel(container, text="실행 날짜", text_color="white", font=("Inter", 11, "bold")).grid(row=0, column=0, sticky="w", padx=12, pady=(8, 2))
        ctk.CTkLabel(container, text="(YYYYMMDD) *미입력시 오늘 날짜", text_color="#999999", font=("Inter", 9)).grid(row=1, column=0, sticky="w", padx=12, pady=(0, 4))
        
        self.exec_date_entry = ctk.CTkEntry(
            container, height=24, fg_color="#1a1a1a",
            border_color="#F4F6FA", border_width=1, text_color="white", font=("Inter", 10)
        )
        self.exec_date_entry.grid(row=0, column=1, rowspan=2, sticky="ew", padx=12, pady=(8, 6))
        if self.settings.get("execution_date"):
            self.exec_date_entry.insert(0, self.settings["execution_date"])
        
        # 날짜 범위
        ctk.CTkLabel(container, text="날짜 조회 범위", text_color="white", font=("Inter", 11, "bold")).grid(row=2, column=0, sticky="w", padx=12, pady=(6, 2))
        ctk.CTkLabel(container, text="*Default: 당해 1/1-오늘", text_color="#999999", font=("Inter", 9)).grid(row=3, column=0, sticky="w", padx=12, pady=(0, 4))
        
        date_frame = ctk.CTkFrame(container, fg_color="transparent")
        date_frame.grid(row=2, column=1, rowspan=2, sticky="ew", padx=12, pady=(6, 6))
        date_frame.grid_columnconfigure((0, 2), weight=1)
        
        self.start_date_entry = ctk.CTkEntry(
            date_frame, width=90, height=24, fg_color="#F4F6FA",
            border_color="#999999", border_width=1, text_color="black", font=("Inter", 10)
        )
        self.start_date_entry.grid(row=0, column=0, sticky="ew")
        
        ctk.CTkLabel(date_frame, text="~", text_color="white", font=("Inter", 11, "bold")).grid(row=0, column=1, padx=8)
        
        self.end_date_entry = ctk.CTkEntry(
            date_frame, width=90, height=24, fg_color="#F4F6FA",
            border_color="#999999", border_width=1, text_color="black", font=("Inter", 10)
        )
        self.end_date_entry.grid(row=0, column=2, sticky="ew")
        
        start_date, end_date = get_effective_date_range(self.settings)
        self.start_date_entry.insert(0, start_date.strftime("%Y%m%d"))
        self.end_date_entry.insert(0, end_date.strftime("%Y%m%d"))
        
        # 파일 저장 경로
        ctk.CTkLabel(container, text="*파일 저장 경로", text_color="white", font=("Inter", 11, "bold")).grid(row=4, column=0, sticky="w", padx=12, pady=(6, 2))
        ctk.CTkLabel(container, text="SAP 자료 저장 위치", text_color="#999999", font=("Inter", 9)).grid(row=5, column=0, sticky="w", padx=12, pady=(0, 4))
        
        path_frame = ctk.CTkFrame(container, fg_color="transparent")
        path_frame.grid(row=4, column=1, rowspan=2, sticky="ew", padx=12, pady=(6, 6))
        path_frame.grid_columnconfigure(0, weight=1)
        
        self.file_path_entry = ctk.CTkEntry(
            path_frame, height=24, fg_color="#1a1a1a",
            border_color="#F4F6FA", border_width=1, text_color="white", font=("Inter", 10)
        )
        self.file_path_entry.pack(side="left", fill="both", expand=True, padx=(0, 6))
        if self.settings.get("export_dir"):
            self.file_path_entry.insert(0, self.settings["export_dir"])
        
        browse_btn = ctk.CTkButton(
            path_frame, text="경로선택", width=70, height=24,
            fg_color="#FF8574", hover_color="#E63D28", text_color="white",
            font=("Inter", 10, "bold"), command=self.browse_directory
        )
        browse_btn.pack(side="left")
        
        # 메일 수신자
        ctk.CTkLabel(container, text="*메일 수신자", text_color="white", font=("Inter", 11, "bold")).grid(row=6, column=0, sticky="w", padx=12, pady=(6, 2))
        ctk.CTkLabel(container, text="여러 사용자: ;로 구분", text_color="#999999", font=("Inter", 9)).grid(row=7, column=0, sticky="w", padx=12, pady=(0, 4))
        
        self.mail_entry = ctk.CTkEntry(
            container, height=24, fg_color="#1a1a1a",
            border_color="#F4F6FA", border_width=1, text_color="white", font=("Inter", 10)
        )
        self.mail_entry.grid(row=6, column=1, rowspan=2, sticky="ew", padx=12, pady=(6, 8))
        if self.settings.get("mail_recipients"):
            self.mail_entry.insert(0, self.settings["mail_recipients"])
        
        # Save button - fixed width like browse button
        save_btn = ctk.CTkButton(
            container, text="설정 저장", width=70, height=24,
            fg_color="#FF8574", hover_color="#E63D28", text_color="white",
            font=("Inter", 10, "bold"), command=self.save_settings
        )
        save_btn.grid(row=8, column=0, columnspan=2, sticky="", padx=12, pady=(0, 8))
    
    def build_progress_section(self):
        """진행 단계 섹션 (좌측)"""
        # Title
        title = ctk.CTkLabel(
            self.progress_frame, text="진행 단계", text_color="white",
            font=("Inter", 13, "bold")
        )
        title.grid(row=0, column=0, sticky="w", pady=(0, 4))
        
        # Progress bar
        progress_bar_frame = ctk.CTkFrame(
            self.progress_frame, fg_color="#333333", height=24, corner_radius=2
        )
        progress_bar_frame.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        progress_bar_frame.pack_propagate(False)
        progress_bar_frame.grid_columnconfigure(1, weight=1)
        
        self.progress_label = ctk.CTkLabel(
            progress_bar_frame, text="진행률: 0%", text_color="white",
            font=("Inter", 9), fg_color="transparent"
        )
        self.progress_label.grid(row=0, column=0, sticky="w", padx=10, pady=3)
        
        # Visual progress bar
        progress_visual = ctk.CTkFrame(
            self.progress_frame, fg_color="#555555", height=4, corner_radius=2
        )
        progress_visual.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        progress_visual.pack_propagate(False)
        self.progress_visual = progress_visual
        self.progress_visual.bind("<Configure>", self._on_progress_track_resize)

        self.progress_bar = ctk.CTkFrame(
            progress_visual, fg_color="#FF8574", height=4, corner_radius=2, width=0
        )
        self.progress_bar.pack(side="left")
        
        # Steps scrollable container
        self.steps_container = ctk.CTkScrollableFrame(
            self.progress_frame, fg_color="transparent", bg_color="transparent"
        )
        self.steps_container.grid(row=3, column=0, sticky="nsew")
        self.steps_container.grid_columnconfigure(0, weight=1)
        
        self.steps = [
            "SAP QN 전체 데이터 조회",
            "기존 보고서 메일 검색",
            "C-Type 신규 QN 분석",
            "C-Type 삭제 QN 분석",
            "Q-Type 신규 QN 분석",
            "Q-Type 삭제 QN 분석",
            "대응 상태(D3/D5/D8) 분석",
            "KPI 계산",
            "보고서 메일 생성",
            "Outlook에서 보고서 열기"
        ]
        
        self.step_frames = {}
        self.step_messages = {}
        
        for i, step_name in enumerate(self.steps):
            self.create_step_accordion(i, step_name)
            self.step_states[i] = {'expanded': False, 'messages': set()}
    
    def create_step_accordion(self, step_index, step_name):
        """아코디언 단계 생성"""
        # Step header (clickable)
        header_frame = ctk.CTkFrame(
            self.steps_container, fg_color="#4a4a4a", corner_radius=2, height=30
        )
        header_frame.pack(fill="x", pady=(0, 1))
        header_frame.pack_propagate(False)
        header_frame.grid_columnconfigure(2, weight=1)
        
        # Indicator
        indicator = ctk.CTkLabel(
            header_frame, text="◯", text_color="#B0B0B0", font=("Inter", 12),
            fg_color="transparent"
        )
        indicator.grid(row=0, column=0, sticky="w", padx=10, pady=4)
        
        # Step number
        step_num = ctk.CTkLabel(
            header_frame, text=f"Step {step_index}", text_color="#888888",
            font=("Inter", 9), fg_color="transparent"
        )
        step_num.grid(row=0, column=1, sticky="w", padx=0, pady=4)
        
        # Step name
        label = ctk.CTkLabel(
            header_frame, text=step_name, text_color="white",
            font=("Inter", 10, "bold"), fg_color="transparent"
        )
        label.grid(row=0, column=2, sticky="w", padx=8, pady=4)
        
        # Message box (hidden by default)
        msg_frame = ctk.CTkFrame(
            self.steps_container, fg_color="#2a2a2a", corner_radius=0
        )
        msg_frame.pack(fill="x", pady=(0, 1))
        msg_frame.pack_forget()
        
        msg_box = ctk.CTkTextbox(
            msg_frame, height=80, fg_color="#1a1a1a",
            text_color="#D0D0D0", font=("Inter", 11),
            border_color="#444444", border_width=1
        )
        msg_box.pack(fill="both", expand=True, padx=8, pady=8)
        msg_box.configure(state="disabled")
        
        # Store references
        self.step_frames[step_index] = {
            'header': header_frame,
            'indicator': indicator,
            'message_frame': msg_frame,
            'message_box': msg_box,
            'label': label
        }
        self.step_messages[step_index] = []
        
        # Bind click event to header and label
        for widget in [header_frame, indicator, label, step_num]:
            widget.bind("<Button-1>", lambda e, idx=step_index: self.toggle_step(idx))
    
    def toggle_step(self, step_index):
        """아코디언 토글"""
        state = self.step_states[step_index]
        msg_frame = self.step_frames[step_index]['message_frame']
        
        if state['expanded']:
            msg_frame.pack_forget()
            state['expanded'] = False
        else:
            msg_frame.pack(fill="x", pady=(0, 1), after=self.step_frames[step_index]['header'])
            state['expanded'] = True
    
    def add_step_message(self, step_index, message):
        """단계별 메시지 추가"""
        if step_index not in self.step_messages:
            return
        
        message_text = (message or "").strip()
        if not message_text:
            message_text = STEP_DETAILS.get(step_index, "Processing...")

        # 동일 메시지는 한 번만 표시
        if message_text in self.step_states[step_index]['messages']:
            return

        self.step_states[step_index]['messages'].add(message_text)
        self.step_messages[step_index].append(message_text)
        msg_box = self.step_frames[step_index]['message_box']
        msg_frame = self.step_frames[step_index]['message_frame']
        
        msg_box.configure(state="normal")
        msg_box.insert("end", f"• {message_text}\n")
        msg_box.see("end")
        msg_box.configure(state="disabled")
        
        # Auto-expand when message added (if not already expanded)
        if not self.step_states[step_index]['expanded']:
            msg_frame.pack(fill="x", pady=(0, 1), after=self.step_frames[step_index]['header'])
            self.step_states[step_index]['expanded'] = True

    def _on_progress_track_resize(self, event):
        """Store latest track width to keep the bar in sync with container size."""
        self.progress_track_width = max(event.width, 0)
        self._apply_progress_bar_width()

    def _apply_progress_bar_width(self, ratio=None):
        """Recalculate the visual width based on the current container size."""
        if not hasattr(self, "progress_bar"):
            return

        ratio = self.progress_ratio if ratio is None else max(0.0, min(1.0, ratio))
        self.progress_ratio = ratio

        self.progress_visual.update_idletasks()
        container_width = self.progress_track_width or self.progress_visual.winfo_width()
        if container_width <= 1:
            container_width = self.progress_frame.winfo_width()

        if ratio <= 0:
            bar_width = 0
        else:
            bar_width = max(2, int(container_width * ratio))

        self.progress_bar.configure(width=bar_width)

    def _on_window_resize(self, event):
        """Dynamically scale widgets and ensure scroll area expands with window."""
        if event.widget is not self:
            return

        width_ratio = event.width / self.base_width if self.base_width else 1
        height_ratio = event.height / self.base_height if self.base_height else 1

        self.update_idletasks()
        req_height = self.container.winfo_reqheight() if hasattr(self, "container") else self.base_height
        height_fit_ratio = event.height / req_height if req_height else 1

        # 화면이 작을 때 더 과감하게 축소(최소 45%), 크게 넓어져도 1.0 이상 확대하지 않음
        target_scale = min(1.0, max(0.45, min(width_ratio, height_ratio, height_fit_ratio)))

        if abs(target_scale - self.widget_scaling) >= 0.03:
            self.widget_scaling = target_scale
            self.spacing_scaling = max(0.5, min(1.0, target_scale + 0.1))
            ctk.set_widget_scaling(self.widget_scaling)
            ctk.set_spacing_scaling(self.spacing_scaling)
            self.after(0, self._apply_progress_bar_width)
    
    def update_step_progress(self, step_index):
        """단계 진행 상태 업데이트"""
        if step_index not in self.step_frames:
            return
        
        indicator = self.step_frames[step_index]['indicator']
        header_frame = self.step_frames[step_index]['header']
        
        # Update indicator
        indicator.configure(text_color="#FF8574", text="●")
        header_frame.configure(fg_color="#555555")
        
        # Update progress value and resize bar proportionally
        total = len(self.steps)
        progress_ratio = (step_index + 1) / total if total else 0
        progress = int(progress_ratio * 100)
        self.progress_label.configure(text=f"진행률: {progress}%")
        self._apply_progress_bar_width(progress_ratio)
        
        self.current_step = step_index

        # 메시지 프레임을 항상 펼쳐서 진행 상황을 바로 보이게 함
        msg_frame = self.step_frames[step_index]['message_frame']
        if not self.step_states[step_index]['expanded']:
            msg_frame.pack(fill="x", pady=(0, 1), after=self.step_frames[step_index]['header'])
            self.step_states[step_index]['expanded'] = True
    
    def build_buttons(self):
        """버튼 섹션 (우측)"""
        # Title
        title = ctk.CTkLabel(
            self.button_frame, text="실행", text_color="white",
            font=("Inter", 13, "bold")
        )
        title.grid(row=0, column=0, sticky="w", pady=(0, 8))
        
        # Main button
        main_btn = ctk.CTkButton(
            self.button_frame, text="전체 통합 실행", height=50,
            fg_color="#FF4F3A", hover_color="#E63024", text_color="white",
            font=("Inter", 12, "bold"), command=self.run_automailer
        )
        main_btn.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        
        # Individual execution section
        ctk.CTkLabel(
            self.button_frame, text="개별 실행", text_color="white",
            font=("Inter", 11, "bold")
        ).grid(row=2, column=0, sticky="w", pady=(0, 8))
        
        # Sub buttons vertically stacked
        sub_buttons_data = [
            ("ZQCN > C1-CK 개별 조회", self.run_c_type),
            ("ZQCN > Q1-Q9 개별 조회", self.run_q_type),
            ("ZNOT > C1-CK 개별 조회", self.run_znot)
        ]
        
        for idx, (text, cmd) in enumerate(sub_buttons_data):
            btn = ctk.CTkButton(
                self.button_frame, text=text, height=36,
                fg_color="#F8F8F8", hover_color="#E0E0E0", text_color="#323232",
                font=("Inter", 10, "bold"), command=cmd
            )
            btn.grid(row=3+idx, column=0, sticky="ew", pady=(0, 12))
        
        # Settings info placed right below individual buttons with slightly larger gap
        info_frame = ctk.CTkFrame(self.button_frame, fg_color="#2a2a2a", corner_radius=2)
        info_frame.grid(row=6, column=0, sticky="ew", pady=(8, 8))
        info_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            info_frame, text="현재 설정", text_color="white",
            font=("Inter", 11, "bold")
        ).grid(row=0, column=0, sticky="w", pady=(8, 4), padx=8)
        
        self.info_label = ctk.CTkLabel(
            info_frame, text="", text_color="#888888",
            font=("Inter", 11), justify="left"
        )
        self.info_label.grid(row=1, column=0, sticky="w", pady=(0, 8), padx=8)
        self.update_info_display()
        
        # Footer buttons - align under info section
        footer_frame = ctk.CTkFrame(self.button_frame, fg_color="transparent")
        footer_frame.grid(row=7, column=0, sticky="e", pady=(0, 0))

        readme_btn = ctk.CTkButton(
            footer_frame, text="도움말", width=70, height=24,
            fg_color="#555555", hover_color="#666666", text_color="white",
            font=("Inter", 9), command=self.show_help
        )
        readme_btn.pack(side="left", padx=(0, 6))

        quit_btn = ctk.CTkButton(
            footer_frame, text="종료", width=70, height=24,
            fg_color="#555555", hover_color="#666666", text_color="white",
            font=("Inter", 9), command=self.quit
        )
        quit_btn.pack(side="left")
    
    def update_info_display(self):
        """설정 정보 표시"""
        export_dir = self.file_path_entry.get().strip()
        mail = self.mail_entry.get().strip()
        
        # Display full paths without truncation
        if export_dir and mail:
            info_text = f"경로:\n{export_dir}\n\n메일:\n{mail}"
        else:
            info_text = "설정 필요"
        self.info_label.configure(text=info_text)
    
    def show_help(self):
        """도움말 표시"""
        help_text = """QN_AutoMailer v1.0

1. 사용자 입력 부분에서 필요한 정보를 입력하세요
2. '설정 저장' 버튼으로 설정을 저장하세요
3. '전체 통합 실행'으로 전체 프로세스를 실행하거나
   '개별 실행'에서 특정 항목만 실행할 수 있습니다
4. 진행 단계에서 각 단계의 상세 내용을 확인하세요

필수 입력: 파일 저장 경로, 메일 수신자"""
        messagebox.showinfo("도움말", help_text)
    
    def browse_directory(self):
        """디렉토리 선택"""
        directory = filedialog.askdirectory(title="SAP 파일 저장 경로 선택")
        if directory:
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, directory)
            self.update_info_display()

    def _collect_settings_from_inputs(self):
        """현재 입력 필드 값을 설정 dict로 수집"""
        export_dir = self.file_path_entry.get().strip()
        mail_recipients = self.mail_entry.get().strip()
        execution_date = self.exec_date_entry.get().strip() or None
        start_date = self.start_date_entry.get().strip() or None
        end_date = self.end_date_entry.get().strip() or None

        # 실행 날짜만 입력되고 종료일이 비어 있으면 종료일 = 실행 날짜
        if execution_date and not end_date:
            end_date = execution_date

        return {
            "export_dir": export_dir,
            "mail_recipients": mail_recipients,
            "execution_date": execution_date,
            "date_range_start": start_date,
            "date_range_end": end_date,
            "plant": "2201",
        }
    
    def save_settings(self):
        """설정 저장"""
        settings = self._collect_settings_from_inputs()

        is_valid, error_msg = validate_settings(settings)
        if not is_valid:
            messagebox.showerror("설정 오류", error_msg)
            return
        
        save_settings(settings)
        messagebox.showinfo("성공", "설정이 저장되었습니다.")
        self.settings = settings
        self.update_info_display()
    
    def run_automailer(self):
        """자동화 프로그램 전체 실행"""
        if not self.validate_before_run():
            return
        
        self.reset_progress()
        thread = threading.Thread(target=self._execute_full, daemon=True)
        thread.start()
    
    def run_c_type(self):
        """C-Type (C1-CK) 조회 실행"""
        if not self.validate_before_run():
            return
        
        thread = threading.Thread(
            target=lambda: self._run_single_export("sap-c", "C1-CK 조회가 완료되었습니다.", "C1-CK"),
            daemon=True,
        )
        thread.start()
    
    def run_q_type(self):
        """Q-Type (Q1-Q9) 조회 실행"""
        if not self.validate_before_run():
            return
        
        thread = threading.Thread(
            target=lambda: self._run_single_export("sap-q", "Q1-Q9 조회가 완료되었습니다.", "Q1-Q9"),
            daemon=True,
        )
        thread.start()
    
    def run_znot(self):
        """ZNOT 조회 실행"""
        if not self.validate_before_run():
            return
        
        thread = threading.Thread(
            target=lambda: self._run_single_export("sap-z", "ZNOT 조회가 완료되었습니다.", "ZNOT"),
            daemon=True,
        )
        thread.start()
    
    def reset_progress(self):
        """진행 상태 초기화"""
        for i in range(len(self.steps)):
            self.step_states[i] = {'expanded': False, 'messages': set()}
            self.step_frames[i]['indicator'].configure(text_color="#B0B0B0", text="◯")
            self.step_frames[i]['header'].configure(fg_color="#4a4a4a")
            self.step_frames[i]['message_box'].configure(state="normal")
            self.step_frames[i]['message_box'].delete("1.0", "end")
            self.step_frames[i]['message_box'].configure(state="disabled")
            self.step_frames[i]['message_frame'].pack_forget()
        
        self.progress_label.configure(text="진행률: 0%")
        self.progress_bar.configure(width=0)
    
    def _execute_full(self):
        """백그라운드에서 전체 실행"""
        try:
            overrides = self._collect_settings_from_inputs()
            runner = self._build_runner()

            # 초기 스텝 바로 표시 (파이프라인 콜백 전에 시각화)
            self.after(0, lambda: self.update_step_progress(0))

            print("[GUI_DEBUG] Starting AutomailerMain pipeline")
            final_settings = runner.run_full_pipeline(overrides, persist=True)
            self.settings = final_settings
            print("[GUI_DEBUG] AutomailerMain pipeline completed")
            
            # 완료 메시지
            def show_complete():
                messagebox.showinfo("Complete", "Task completed successfully.")
            self.after(0, show_complete)
            
        except Exception as e:
            import traceback
            error_msg = f"Error occurred:\n{str(e)}"
            def show_error():
                messagebox.showerror("Error", error_msg)
            self.after(0, show_error)
            traceback.print_exc()
    
    def validate_before_run(self):
        """실행 전 설정 검증"""
        settings = self._collect_settings_from_inputs()

        if not settings.get("export_dir") or not settings.get("mail_recipients"):
            messagebox.showerror("오류", "필수 입력항목을 확인하세요:\n✓ 파일 저장 경로\n✓ 메일 수신자")
            return False

        is_valid, error_msg = validate_settings(settings)
        if not is_valid:
            messagebox.showerror("설정 오류", error_msg)
            return False
        
        save_settings(settings)
        self.settings = settings
        return True


if __name__ == "__main__":
    app = QNAutomailerGUI()
    app.mainloop()
