from __future__ import annotations

import logging
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

import config
import create_excel
from autofill import VocabularyAutofiller


APP_TITLE = "German Vocabulary Manager"


class QueueLogHandler(logging.Handler):
    """Send log records into a thread-safe queue for display in the GUI."""

    def __init__(self, log_queue: queue.Queue[str]) -> None:
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record: logging.LogRecord) -> None:
        message = self.format(record)
        self.log_queue.put(message)


class GermanVocabularyApp(tk.Tk):
    """Small Windows desktop UI for creating and autofilling the workbook."""

    def __init__(self) -> None:
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("820x560")
        self.minsize(760, 500)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self._configure_logging()
        self._build_ui()
        self._poll_log_queue()

    def _configure_logging(self) -> None:
        gui_handler = QueueLogHandler(self.log_queue)
        gui_handler.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)s [%(name)s] %(message)s")
        )

        root_logger = logging.getLogger()
        root_logger.addHandler(gui_handler)
        root_logger.setLevel(logging.INFO)

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        header = ttk.Frame(self, padding=(20, 18, 20, 8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        title = ttk.Label(
            header,
            text="German Vocabulary Manager",
            font=("Segoe UI", 20, "bold"),
        )
        title.grid(row=0, column=0, sticky="w")

        subtitle = ttk.Label(
            header,
            text="Create your workbook, autofill German words, and open Excel from one place.",
            font=("Segoe UI", 10),
        )
        subtitle.grid(row=1, column=0, sticky="w", pady=(4, 0))

        button_bar = ttk.Frame(self, padding=(20, 8, 20, 12))
        button_bar.grid(row=1, column=0, sticky="ew")

        self.create_button = ttk.Button(
            button_bar,
            text="Create / Reset Workbook",
            command=self.create_workbook,
        )
        self.create_button.grid(row=0, column=0, padx=(0, 10), ipadx=8, ipady=6)

        self.autofill_button = ttk.Button(
            button_bar,
            text="Autofill Vocabulary",
            command=self.autofill_workbook,
        )
        self.autofill_button.grid(row=0, column=1, padx=(0, 10), ipadx=8, ipady=6)

        self.open_excel_button = ttk.Button(
            button_bar,
            text="Open Workbook",
            command=self.open_workbook,
        )
        self.open_excel_button.grid(row=0, column=2, padx=(0, 10), ipadx=8, ipady=6)

        self.open_folder_button = ttk.Button(
            button_bar,
            text="Open Project Folder",
            command=self.open_project_folder,
        )
        self.open_folder_button.grid(row=0, column=3, ipadx=8, ipady=6)

        content = ttk.Frame(self, padding=(20, 0, 20, 20))
        content.grid(row=2, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        self.status_label = ttk.Label(
            content,
            text=self._status_text(),
            font=("Segoe UI", 10),
        )
        self.status_label.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        self.log_box = scrolledtext.ScrolledText(
            content,
            wrap=tk.WORD,
            height=18,
            font=("Consolas", 9),
        )
        self.log_box.grid(row=1, column=0, sticky="nsew")
        self.log_box.configure(state="disabled")

    def _status_text(self) -> str:
        workbook_status = (
            "Workbook found"
            if config.WORKBOOK_PATH.exists()
            else "Workbook not created yet"
        )
        return f"{workbook_status}: {config.WORKBOOK_PATH}"

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.create_button.configure(state=state)
        self.autofill_button.configure(state=state)
        self.open_excel_button.configure(state=state)
        self.open_folder_button.configure(state=state)

    def _run_background(self, title: str, task) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo(APP_TITLE, "Another task is already running.")
            return

        self._append_log(f"\n--- {title} ---")
        self._set_busy(True)

        def runner() -> None:
            try:
                task()
                self.log_queue.put(f"{title} finished successfully.")
            except Exception as exc:
                logging.getLogger(__name__).exception("%s failed.", title)
                self.log_queue.put(f"{title} failed: {exc}")
            finally:
                self.after(0, self._task_finished)

        self.worker_thread = threading.Thread(target=runner, daemon=True)
        self.worker_thread.start()

    def _task_finished(self) -> None:
        self._set_busy(False)
        self.status_label.configure(text=self._status_text())

    def _append_log(self, text: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert(tk.END, text + "\n")
        self.log_box.see(tk.END)
        self.log_box.configure(state="disabled")

    def _poll_log_queue(self) -> None:
        while True:
            try:
                message = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self._append_log(message)

        self.after(150, self._poll_log_queue)

    def _workbook_lock_file_exists(self) -> bool:
        lock_file = config.WORKBOOK_PATH.with_name(f"~${config.WORKBOOK_PATH.name}")
        return lock_file.exists()

    def create_workbook(self) -> None:
        if self._workbook_lock_file_exists():
            messagebox.showwarning(
                APP_TITLE,
                "Please close vocabulary.xlsx in Excel before creating the workbook.",
            )
            return

        confirmed = messagebox.askyesno(
            APP_TITLE,
            "This will create/reset vocabulary.xlsx. Continue?",
        )
        if not confirmed:
            return

        self._run_background("Create workbook", create_excel.main)

    def autofill_workbook(self) -> None:
        if self._workbook_lock_file_exists():
            messagebox.showwarning(
                APP_TITLE,
                "Please close vocabulary.xlsx in Excel before running autofill.",
            )
            return

        if not config.WORKBOOK_PATH.exists():
            messagebox.showwarning(
                APP_TITLE,
                "vocabulary.xlsx does not exist yet. Click 'Create / Reset Workbook' first.",
            )
            return

        self._run_background("Autofill vocabulary", lambda: VocabularyAutofiller().run())

    def open_workbook(self) -> None:
        if not config.WORKBOOK_PATH.exists():
            messagebox.showwarning(
                APP_TITLE,
                "vocabulary.xlsx does not exist yet. Create it first.",
            )
            return

        self._open_path(config.WORKBOOK_PATH)

    def open_project_folder(self) -> None:
        self._open_path(config.BASE_DIR)

    @staticmethod
    def _open_path(path: Path) -> None:
        if sys.platform.startswith("win"):
            subprocess.Popen(["explorer", str(path)])
        else:
            subprocess.Popen(["open", str(path)])


def main() -> None:
    app = GermanVocabularyApp()
    app.mainloop()


if __name__ == "__main__":
    main()