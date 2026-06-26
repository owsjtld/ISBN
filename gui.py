import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from dotenv import load_dotenv

load_dotenv()

from core import SEARCH_MODES, process_dataframe, to_excel_bytes
import pandas as pd

SEARCH_MODE_LABELS = list(SEARCH_MODES.keys())


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ISBN 자동 입력기")
        self.resizable(False, False)
        self._build_ui()
        self._df = None
        self._result_df = None

    def _build_ui(self):
        pad = dict(padx=12, pady=6)

        # API 키
        frm_key = ttk.LabelFrame(self, text="카카오 REST API 키")
        frm_key.pack(fill="x", **pad)
        self._api_key = tk.StringVar(value=os.getenv("KAKAO_API_KEY", ""))
        ttk.Entry(frm_key, textvariable=self._api_key, width=52, show="*").pack(
            fill="x", padx=8, pady=6
        )

        # 파일 선택
        frm_file = ttk.LabelFrame(self, text="엑셀 파일")
        frm_file.pack(fill="x", **pad)
        self._file_path = tk.StringVar()
        file_row = ttk.Frame(frm_file)
        file_row.pack(fill="x", padx=8, pady=6)
        ttk.Entry(file_row, textvariable=self._file_path, width=42).pack(side="left")
        ttk.Button(file_row, text="찾아보기", command=self._browse).pack(side="left", padx=(6, 0))

        # 검색 방식
        frm_mode = ttk.LabelFrame(self, text="검색 방식")
        frm_mode.pack(fill="x", **pad)
        self._mode_var = tk.StringVar(value=SEARCH_MODE_LABELS[0])
        for label in SEARCH_MODE_LABELS:
            ttk.Radiobutton(frm_mode, text=label, variable=self._mode_var, value=label).pack(
                anchor="w", padx=8, pady=2
            )

        # 진행 상황
        frm_progress = ttk.LabelFrame(self, text="진행 상황")
        frm_progress.pack(fill="x", **pad)
        self._progress_var = tk.DoubleVar()
        self._progress_bar = ttk.Progressbar(
            frm_progress, variable=self._progress_var, maximum=100
        )
        self._progress_bar.pack(fill="x", padx=8, pady=(6, 2))
        self._status_var = tk.StringVar(value="파일을 선택하세요.")
        ttk.Label(frm_progress, textvariable=self._status_var).pack(anchor="w", padx=8, pady=(0, 2))

        # 로그
        frm_log = ttk.LabelFrame(self, text="로그")
        frm_log.pack(fill="both", expand=True, **pad)
        self._log = tk.Text(frm_log, height=12, width=62, state="disabled", font=("Consolas", 9))
        scroll = ttk.Scrollbar(frm_log, command=self._log.yview)
        self._log.configure(yscrollcommand=scroll.set)
        self._log.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=6)
        scroll.pack(side="right", fill="y", pady=6, padx=(0, 4))

        # 버튼
        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill="x", **pad)
        self._btn_start = ttk.Button(frm_btn, text="ISBN 검색 시작", command=self._start)
        self._btn_start.pack(side="left")
        self._btn_save = ttk.Button(
            frm_btn, text="결과 저장", command=self._save, state="disabled"
        )
        self._btn_save.pack(side="left", padx=(8, 0))

    def _browse(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 파일", "*.xlsx")])
        if path:
            self._file_path.set(path)
            self._load_file(path)

    def _load_file(self, path):
        try:
            try:
                df = pd.read_excel(path)
            except Exception:
                df = pd.read_excel(path, engine='calamine')
            self._df = df
            already = int(
                df.get('ISBN', pd.Series()).fillna('').astype(str)
                .apply(lambda x: len(x.strip()) >= 10).sum()
            )
            self._log_write(f"파일 로드: {len(df)}개 도서 (ISBN 기존 {already}개)\n")
            self._status_var.set(f"총 {len(df)}개 도서 로드됨")
        except Exception as e:
            messagebox.showerror("오류", f"파일을 읽을 수 없습니다.\n{e}")

    def _start(self):
        if not self._api_key.get().strip():
            messagebox.showwarning("API 키 없음", "카카오 REST API 키를 입력하세요.")
            return
        if self._df is None:
            messagebox.showwarning("파일 없음", "엑셀 파일을 선택하세요.")
            return

        self._btn_start.config(state="disabled")
        self._btn_save.config(state="disabled")
        self._result_df = None
        self._progress_var.set(0)

        mode = SEARCH_MODES[self._mode_var.get()]
        api_key = self._api_key.get().strip()
        df = self._df.copy()

        def run():
            stats = {"success": 0, "fail": 0}

            def on_progress(current, total, title, isbn, success, skipped):
                self._progress_var.set(current / total * 100)
                if skipped:
                    icon = "⏭"
                elif success:
                    icon = "✅"
                    stats["success"] += 1
                else:
                    icon = "❌"
                    stats["fail"] += 1

                short = title[:28] + ("…" if len(title) > 28 else "")
                self._log_write(f"{icon} [{current}/{total}] {short} → {isbn or '실패'}\n")
                self._status_var.set(
                    f"✅ 성공 {stats['success']}  ❌ 실패 {stats['fail']}  ({current}/{total})"
                )

            result = process_dataframe(df, api_key, mode, on_progress)
            self._result_df = result
            self._progress_var.set(100)
            self._log_write(f"\n완료! {len(result)}개 중 성공 {stats['success']}개  실패 {stats['fail']}개\n")
            self._status_var.set(f"완료!  성공 {stats['success']}개  /  실패 {stats['fail']}개")
            self._btn_start.config(state="normal")
            self._btn_save.config(state="normal")

        threading.Thread(target=run, daemon=True).start()

    def _save(self):
        if self._result_df is None:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            initialfile="ISBN_완료.xlsx",
        )
        if not path:
            return
        with open(path, "wb") as f:
            f.write(to_excel_bytes(self._result_df))
        messagebox.showinfo("저장 완료", f"저장됐습니다!\n{path}")

    def _log_write(self, text: str):
        self._log.config(state="normal")
        self._log.insert("end", text)
        self._log.see("end")
        self._log.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()
