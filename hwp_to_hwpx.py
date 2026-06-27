"""
HWP → HWPX 일괄 변환기
한컴오피스 2022가 설치된 Windows 환경에서 실행

사용법:
    python hwp_to_hwpx.py

필요 패키지:
    pip install pywin32
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext

try:
    import win32com.client
    WIN32_OK = True
except ImportError:
    WIN32_OK = False


def find_hwp_files(folder: str) -> list[str]:
    """폴더 내 모든 .hwp 파일 경로 반환 (재귀 탐색)"""
    result = []
    for root, _, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(".hwp") and not f.lower().endswith(".hwpx"):
                result.append(os.path.join(root, f))
    return sorted(result)


def _get_hwp_instance(log_fn):
    """
    HWP COM 인스턴스 생성 및 보안 모듈 등록.
    보안 팝업("접근 허용") 없이 파일을 열기 위해 FilePathCheckerModule을 등록한다.
    """
    if not WIN32_OK:
        log_fn("오류: pywin32가 설치되지 않았습니다. 'pip install pywin32' 실행 후 재시도하세요.")
        return None

    hwp = win32com.client.Dispatch("HWPFrame.HwpObject")

    # 보안 모듈 등록 — 파일 접근 허용 팝업 차단
    # 한컴오피스 설치 경로에서 모듈을 찾아 등록한다
    checker_paths = [
        r"C:\Program Files (x86)\HNC\Hwp80\FilePathCheckerModule.dll",
        r"C:\Program Files\HNC\Hwp80\FilePathCheckerModule.dll",
        r"C:\Program Files (x86)\Hnc\Office 2022\HOffice110\Bin\FilePathCheckerModule.dll",
        r"C:\Program Files\Hnc\Office 2022\HOffice110\Bin\FilePathCheckerModule.dll",
        r"C:\Program Files (x86)\Hnc\Office NEO\HOffice96\Bin\FilePathCheckerModule.dll",
    ]
    registered = False
    for dll_path in checker_paths:
        if os.path.exists(dll_path):
            try:
                hwp.RegisterModule("FilePathCheckDLL", dll_path)
                registered = True
                break
            except Exception:
                pass

    # 경로 탐색 실패 시 모듈명으로 시도 (PATH에 있을 경우)
    if not registered:
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            log_fn("  경고: 보안 모듈 등록 실패 — 파일 접근 허용 팝업이 나타날 수 있습니다.")

    return hwp


def batch_convert(files: list[str], output_dir: str | None, log_fn, progress_fn=None) -> tuple[int, int]:
    """
    HWP 파일 목록을 HWPX로 일괄 변환.
    HWP 인스턴스를 한 번만 생성해 재사용한다 (속도 향상, 팝업 최소화).
    반환값: (성공 수, 실패 수)
    """
    hwp = _get_hwp_instance(log_fn)
    if hwp is None:
        return 0, len(files)

    ok = fail = 0
    try:
        for i, hwp_path in enumerate(files):
            src = os.path.abspath(hwp_path)
            basename = os.path.splitext(os.path.basename(src))[0]
            dest_folder = output_dir if output_dir else os.path.dirname(src)
            dest = os.path.join(dest_folder, basename + ".hwpx")

            if os.path.exists(dest):
                log_fn(f"  건너뜀 (이미 존재): {os.path.basename(dest)}")
                ok += 1
            else:
                try:
                    # Open: (경로, 형식, 옵션)  ← 3개 인자 필수 (한컴 2022)
                    if not hwp.Open(src, "HWP", "AutoOpen"):
                        raise Exception("Open() 반환값 False — 파일 열기 실패")
                    # SaveAs: (경로, 형식, 옵션) ← 3개 인자 필수 (한컴 2022)
                    hwp.SaveAs(dest, "HWPX", "")
                    log_fn(f"  완료: {os.path.basename(src)} → {os.path.basename(dest)}")
                    ok += 1
                except Exception as e:
                    log_fn(f"  실패: {os.path.basename(src)} — {e}")
                    fail += 1

            if progress_fn:
                progress_fn(i + 1)
    finally:
        try:
            hwp.Quit()
        except Exception:
            pass

    return ok, fail


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HWP → HWPX 일괄 변환기")
        self.resizable(False, False)
        self._build_ui()

    # ── UI 구성 ──────────────────────────────────────────────
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # 소스 폴더
        src_frame = tk.LabelFrame(self, text="HWP 파일이 있는 폴더", **pad)
        src_frame.grid(row=0, column=0, columnspan=2, sticky="ew", **pad)
        src_frame.columnconfigure(0, weight=1)

        self.src_var = tk.StringVar()
        tk.Entry(src_frame, textvariable=self.src_var, width=55).grid(row=0, column=0, padx=4)
        tk.Button(src_frame, text="찾아보기", command=self._pick_src).grid(row=0, column=1, padx=4)

        # 출력 폴더
        out_frame = tk.LabelFrame(self, text="저장 폴더 (비워두면 원본 위치에 저장)", **pad)
        out_frame.grid(row=1, column=0, columnspan=2, sticky="ew", **pad)
        out_frame.columnconfigure(0, weight=1)

        self.out_var = tk.StringVar()
        tk.Entry(out_frame, textvariable=self.out_var, width=55).grid(row=0, column=0, padx=4)
        tk.Button(out_frame, text="찾아보기", command=self._pick_out).grid(row=0, column=1, padx=4)

        # 옵션
        opt_frame = tk.Frame(self)
        opt_frame.grid(row=2, column=0, columnspan=2, sticky="w", padx=10)
        self.recursive_var = tk.BooleanVar(value=True)
        tk.Checkbutton(opt_frame, text="하위 폴더 포함", variable=self.recursive_var).pack(side="left")

        # 파일 목록 미리보기
        preview_frame = tk.LabelFrame(self, text="변환 대상 파일", **pad)
        preview_frame.grid(row=3, column=0, columnspan=2, sticky="ew", **pad)
        self.preview = scrolledtext.ScrolledText(preview_frame, width=70, height=8, state="disabled", font=("맑은 고딕", 9))
        self.preview.pack()
        tk.Button(preview_frame, text="목록 새로고침", command=self._refresh_preview).pack(pady=4)

        # 진행 상황
        self.progress = ttk.Progressbar(self, length=480, mode="determinate")
        self.progress.grid(row=4, column=0, columnspan=2, **pad)

        # 로그
        log_frame = tk.LabelFrame(self, text="변환 로그", **pad)
        log_frame.grid(row=5, column=0, columnspan=2, sticky="ew", **pad)
        self.log_box = scrolledtext.ScrolledText(log_frame, width=70, height=10, state="disabled", font=("맑은 고딕", 9))
        self.log_box.pack()

        # 버튼
        btn_frame = tk.Frame(self)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=8)
        self.convert_btn = tk.Button(btn_frame, text="변환 시작", width=16, bg="#0078D4", fg="white",
                                     font=("맑은 고딕", 10, "bold"), command=self._start_convert)
        self.convert_btn.pack(side="left", padx=6)
        tk.Button(btn_frame, text="닫기", width=10, command=self.destroy).pack(side="left", padx=6)

        # 상태바
        self.status_var = tk.StringVar(value="폴더를 선택하세요.")
        tk.Label(self, textvariable=self.status_var, anchor="w", fg="gray").grid(
            row=7, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 6))

    # ── 이벤트 ───────────────────────────────────────────────
    def _pick_src(self):
        folder = filedialog.askdirectory(title="HWP 파일 폴더 선택")
        if folder:
            self.src_var.set(folder)
            self._refresh_preview()

    def _pick_out(self):
        folder = filedialog.askdirectory(title="저장 폴더 선택")
        if folder:
            self.out_var.set(folder)

    def _refresh_preview(self):
        src = self.src_var.get().strip()
        if not src or not os.path.isdir(src):
            return
        files = find_hwp_files(src)
        self.preview.config(state="normal")
        self.preview.delete("1.0", "end")
        for f in files:
            self.preview.insert("end", f + "\n")
        self.preview.config(state="disabled")
        self.status_var.set(f"{len(files)}개 HWP 파일 발견.")

    def _log(self, msg: str):
        self.log_box.config(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")
        self.update_idletasks()

    def _start_convert(self):
        src = self.src_var.get().strip()
        if not src or not os.path.isdir(src):
            self.status_var.set("올바른 소스 폴더를 선택하세요.")
            return

        out = self.out_var.get().strip() or None
        if out and not os.path.isdir(out):
            os.makedirs(out, exist_ok=True)

        files = find_hwp_files(src)
        if not files:
            self.status_var.set("변환할 HWP 파일이 없습니다.")
            return

        self.convert_btn.config(state="disabled")
        threading.Thread(target=self._run_convert, args=(files, out), daemon=True).start()

    def _run_convert(self, files: list[str], out_dir: str | None):
        total = len(files)
        self._log(f"── 변환 시작: 총 {total}개 ──")
        self.progress["maximum"] = total

        def on_progress(n: int):
            self.progress["value"] = n

        # 파일별 번호 출력을 위해 log_fn 래핑
        idx = [0]
        original_log = self._log

        def numbered_log(msg: str):
            if msg.startswith("  "):  # 파일별 결과 메시지
                original_log(msg)
            else:
                original_log(msg)

        # 파일별 헤더 출력 후 batch_convert로 위임
        # batch_convert 내부에서 직접 출력하므로 파일 번호를 앞에 추가
        log_with_index = []

        def log_fn(msg: str):
            if not msg.startswith("  ") and not msg.startswith("경고"):
                pass  # batch_convert 내부 로그는 그대로 전달
            self._log(msg)

        # 파일별 번호 헤더 출력을 위해 custom wrapper 사용
        file_iter = iter(enumerate(files, 1))

        def tracking_log(msg: str):
            self._log(msg)

        # 파일별 번호 출력: batch_convert 호출 전 각 파일 진입 시점에 출력
        # 단순화: batch_convert에 넘기기 전에 파일 헤더를 붙인 로그 함수 사용
        current = [0]

        def log_with_header(msg: str):
            # "완료:" 또는 "실패:" 또는 "건너뜀"으로 시작하면 파일 결과
            self._log(msg)

        def progress_with_header(n: int):
            # n번째 파일 처리 완료 시점 — 헤더는 이미 출력됨
            self.progress["value"] = n

        # 파일별 헤더 출력을 위한 최종 접근: 파일마다 직접 루프
        ok_total = fail_total = 0

        hwp_instance_log = []

        def pre_log(msg):
            self._log(msg)

        # HWP 인스턴스 준비 로그
        hwp = _get_hwp_instance(pre_log)
        if hwp is None:
            self._log("HWP 인스턴스 생성 실패. 종료합니다.")
            self.convert_btn.config(state="normal")
            return

        try:
            for i, hwp_path in enumerate(files, 1):
                self._log(f"[{i}/{total}] {os.path.basename(hwp_path)}")
                src = os.path.abspath(hwp_path)
                basename = os.path.splitext(os.path.basename(src))[0]
                dest_folder = out_dir if out_dir else os.path.dirname(src)
                dest = os.path.join(dest_folder, basename + ".hwpx")

                if os.path.exists(dest):
                    self._log(f"  건너뜀 (이미 존재): {os.path.basename(dest)}")
                    ok_total += 1
                else:
                    try:
                        if not hwp.Open(src, "HWP", "AutoOpen"):
                            raise Exception("Open() 반환값 False — 파일 열기 실패")
                        hwp.SaveAs(dest, "HWPX", "")
                        self._log(f"  완료: {os.path.basename(src)} → {os.path.basename(dest)}")
                        ok_total += 1
                    except Exception as e:
                        self._log(f"  실패: {os.path.basename(src)} — {e}")
                        fail_total += 1

                self.progress["value"] = i
        finally:
            try:
                hwp.Quit()
            except Exception:
                pass

        summary = f"완료: {ok_total}개 성공 / {fail_total}개 실패 (전체 {total}개)"
        self._log(f"\n── {summary} ──")
        self.status_var.set(summary)
        self.convert_btn.config(state="normal")


if __name__ == "__main__":
    app = App()
    app.mainloop()
