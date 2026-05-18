from __future__ import annotations

import os
import subprocess
import sys
import threading
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter.scrolledtext import ScrolledText
    TK_IMPORT_ERROR: ModuleNotFoundError | None = None
except ModuleNotFoundError as exc:
    if exc.name not in {"_tkinter", "tkinter"}:
        raise
    tk = None  # type: ignore[assignment]
    filedialog = None  # type: ignore[assignment]
    messagebox = None  # type: ignore[assignment]
    ttk = None  # type: ignore[assignment]
    ScrolledText = None  # type: ignore[assignment]
    TK_IMPORT_ERROR = exc


APP_PATH = Path(__file__).with_name("app.py")
OPENAI_IMAGE_API_KEY_ENV = "OPENAI_API_KEY_FOR_IMAGE"
LEGACY_OPENAI_API_KEY_ENV = "OPENAI_API_KEY"
IMAGE_SIZES = ["1024x1024", "1536x1024", "1024x1536", "auto"]
VIDEO_RESOLUTIONS = ["512P", "720P", "768P", "1080P"]
AUDIO_FORMATS = ["mp3", "wav", "opus", "flac", "pcm", "pcmu_raw", "pcmu_wav"]


def get_tk_install_hint() -> str:
    return "\n".join(
        [
            "無法啟動 GUI：目前的 Python 缺少 tkinter / _tkinter 支援。",
            "",
            "這通常發生在 macOS 的 Homebrew Python 尚未安裝對應 Tk 套件時。",
            "可先執行 CLI 版本：python app.py --help",
            "",
            "建議處理方式：",
            "1. brew install python-tk@3.13",
            "2. 重新建立虛擬環境後再執行 python gui.py",
            "",
            "如果仍失敗，請改用 python.org 官方 Python，或確認目前 venv 使用的是支援 Tk 的 Python。",
        ]
    )


class PptFilmMakerGui:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PowerPoint 轉影片工具")
        self.root.geometry("1040x900")
        self.is_running = False

        self.openai_key_var = tk.StringVar(
            value=os.getenv(OPENAI_IMAGE_API_KEY_ENV) or os.getenv(LEGACY_OPENAI_API_KEY_ENV, "")
        )
        self.minimax_key_var = tk.StringVar(value=os.getenv("MINIMAX_API_KEY", ""))
        self.output_var = tk.StringVar()

        self.template_model_var = tk.StringVar(value="gpt-image-2")
        self.template_size_var = tk.StringVar(value="1024x1024")

        self.scene_model_var = tk.StringVar(value="gpt-image-2")
        self.scene_size_var = tk.StringVar(value="1024x1024")
        self.scene_ppt_var = tk.StringVar()
        self.scene_template_var = tk.StringVar()
        self.scene_number_var = tk.StringVar()

        self.tts_model_var = tk.StringVar(value="speech-2.8-hd")
        self.tts_voice_var = tk.StringVar(value="English_Graceful_Lady")
        self.tts_format_var = tk.StringVar(value="mp3")
        self.tts_speed_var = tk.StringVar(value="1.0")
        self.tts_pitch_var = tk.StringVar(value="0")
        self.tts_volume_var = tk.StringVar(value="1.0")
        self.tts_text_file_var = tk.StringVar()

        self.video_image_var = tk.StringVar()
        self.video_end_image_var = tk.StringVar()
        self.video_audio_var = tk.StringVar()
        self.video_model_var = tk.StringVar(value="MiniMax-Hailuo-02")
        self.video_duration_var = tk.StringVar(value="6")
        self.video_resolution_var = tk.StringVar(value="1080P")
        self.video_poll_interval_var = tk.StringVar(value="10")
        self.video_timeout_var = tk.StringVar(value="1800")

        self.subtitle_video_var = tk.StringVar()
        self.subtitle_file_var = tk.StringVar()

        self.compose_video_var = tk.StringVar()
        self.compose_audio_var = tk.StringVar()

        self.concat_input_dir_var = tk.StringVar()

        self._build_ui()

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        container = ttk.Frame(self.root, padding=14)
        container.grid(sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)
        container.rowconfigure(2, weight=1)

        common_frame = ttk.LabelFrame(container, text="共用設定", padding=10)
        common_frame.grid(row=0, column=0, sticky="ew")
        common_frame.columnconfigure(1, weight=1)
        self._add_labeled_entry(common_frame, 0, OPENAI_IMAGE_API_KEY_ENV, self.openai_key_var, show="*")
        self._add_labeled_entry(common_frame, 1, "MINIMAX_API_KEY", self.minimax_key_var, show="*")
        self._add_path_row(common_frame, 2, "輸出路徑", self.output_var, save=True)

        self.notebook = ttk.Notebook(container)
        self.notebook.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        self.task_frames: dict[str, ttk.Frame] = {}
        self._build_template_tab()
        self._build_scene_tab()
        self._build_tts_tab()
        self._build_video_tab()
        self._build_subtitle_tab()
        self._build_compose_tab()
        self._build_concat_tab()

        log_frame = ttk.LabelFrame(container, text="執行結果", padding=10)
        log_frame.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.output_text = ScrolledText(log_frame, wrap=tk.WORD, height=13)
        self.output_text.grid(row=0, column=0, sticky="nsew")

        action_frame = ttk.Frame(container)
        action_frame.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        action_frame.columnconfigure(0, weight=1)
        self.run_button = ttk.Button(action_frame, text="執行", command=self._run_task)
        self.run_button.grid(row=0, column=1, sticky="e")

    def _new_tab(self, task: str, title: str) -> ttk.Frame:
        frame = ttk.Frame(self.notebook, padding=12)
        frame.columnconfigure(1, weight=1)
        self.notebook.add(frame, text=title)
        self.task_frames[task] = frame
        return frame

    def _build_template_tab(self) -> None:
        frame = self._new_tab("template", "模板")
        self._add_labeled_entry(frame, 0, "模型", self.template_model_var)
        self._add_labeled_combobox(frame, 1, "尺寸", self.template_size_var, IMAGE_SIZES)
        self.template_refs_text = self._add_multi_path_row(frame, 2, "參考圖片", height=4)
        ttk.Label(frame, text="Prompt").grid(row=3, column=0, sticky="nw", pady=(6, 0))
        self.template_prompt_text = ScrolledText(frame, wrap=tk.WORD, height=10)
        self.template_prompt_text.grid(row=3, column=1, sticky="nsew", pady=(6, 0))
        ttk.Label(frame, text="可一次選多張，或手動每行輸入一個圖片路徑。", wraplength=760).grid(
            row=4, column=1, sticky="w", pady=(6, 0)
        )
        frame.rowconfigure(3, weight=1)

    def _build_scene_tab(self) -> None:
        frame = self._new_tab("scene", "場景")
        self._add_path_row(frame, 0, "PPT 圖片", self.scene_ppt_var)
        self._add_path_row(frame, 1, "模板圖片", self.scene_template_var)
        self.scene_refs_text = self._add_multi_path_row(frame, 2, "額外參考圖片", height=4)
        self._add_labeled_entry(frame, 3, "場景編號", self.scene_number_var)
        self._add_labeled_entry(frame, 4, "模型", self.scene_model_var)
        self._add_labeled_combobox(frame, 5, "尺寸", self.scene_size_var, IMAGE_SIZES)
        ttk.Label(frame, text="Prompt").grid(row=6, column=0, sticky="nw", pady=(6, 0))
        self.scene_prompt_text = ScrolledText(frame, wrap=tk.WORD, height=9)
        self.scene_prompt_text.grid(row=6, column=1, sticky="nsew", pady=(6, 0))
        ttk.Label(frame, text="可一次選多張，或手動每行輸入一個圖片路徑。", wraplength=760).grid(
            row=7, column=1, sticky="w", pady=(6, 0)
        )
        frame.rowconfigure(6, weight=1)

    def _build_tts_tab(self) -> None:
        frame = self._new_tab("tts", "語音")
        self._add_path_row(frame, 0, "文稿檔", self.tts_text_file_var)
        self._add_labeled_entry(frame, 1, "模型", self.tts_model_var)
        self._add_labeled_entry(frame, 2, "聲線", self.tts_voice_var)
        self._add_labeled_combobox(frame, 3, "音訊格式", self.tts_format_var, AUDIO_FORMATS)
        self._add_labeled_entry(frame, 4, "語速", self.tts_speed_var)
        self._add_labeled_entry(frame, 5, "音調", self.tts_pitch_var)
        self._add_labeled_entry(frame, 6, "音量", self.tts_volume_var)
        ttk.Label(frame, text="文稿內容").grid(row=7, column=0, sticky="nw", pady=(6, 0))
        self.tts_text = ScrolledText(frame, wrap=tk.WORD, height=9)
        self.tts_text.grid(row=7, column=1, sticky="nsew", pady=(6, 0))
        frame.rowconfigure(7, weight=1)

    def _build_video_tab(self) -> None:
        frame = self._new_tab("video", "影片")
        self._add_path_row(frame, 0, "起始圖片", self.video_image_var)
        self._add_path_row(frame, 1, "結束圖片", self.video_end_image_var)
        self._add_path_row(frame, 2, "語音音檔", self.video_audio_var)
        self._add_labeled_entry(frame, 3, "模型", self.video_model_var)
        self._add_labeled_entry(frame, 4, "秒數", self.video_duration_var)
        self._add_labeled_combobox(frame, 5, "解析度", self.video_resolution_var, VIDEO_RESOLUTIONS)
        self._add_labeled_entry(frame, 6, "輪詢間隔", self.video_poll_interval_var)
        self._add_labeled_entry(frame, 7, "Timeout", self.video_timeout_var)
        ttk.Label(frame, text="Prompt").grid(row=8, column=0, sticky="nw", pady=(6, 0))
        self.video_prompt_text = ScrolledText(frame, wrap=tk.WORD, height=8)
        self.video_prompt_text.grid(row=8, column=1, sticky="nsew", pady=(6, 0))
        frame.rowconfigure(8, weight=1)

    def _build_subtitle_tab(self) -> None:
        frame = self._new_tab("subtitle", "字幕")
        self._add_path_row(frame, 0, "影片", self.subtitle_video_var)
        self._add_path_row(frame, 1, "字幕檔", self.subtitle_file_var)
        ttk.Label(frame, text="說明").grid(row=2, column=0, sticky="nw", pady=(8, 0))
        ttk.Label(frame, text="使用 ffmpeg subtitles filter 將 .srt 等字幕檔燒入影片。", wraplength=760).grid(
            row=2, column=1, sticky="w", pady=(8, 0)
        )

    def _build_compose_tab(self) -> None:
        frame = self._new_tab("compose", "影音合成")
        self._add_path_row(frame, 0, "影片", self.compose_video_var)
        self._add_path_row(frame, 1, "語音音檔", self.compose_audio_var)
        ttk.Label(frame, text="說明").grid(row=2, column=0, sticky="nw", pady=(8, 0))
        ttk.Label(frame, text="將影片畫面與外部語音音檔合成，輸出長度取較短者。", wraplength=760).grid(
            row=2, column=1, sticky="w", pady=(8, 0)
        )

    def _build_concat_tab(self) -> None:
        frame = self._new_tab("concat", "影片串接")
        self._add_path_row(frame, 0, "輸入資料夾", self.concat_input_dir_var, directory=True)
        ttk.Label(frame, text="影片清單").grid(row=1, column=0, sticky="nw", pady=(6, 0))
        self.concat_inputs_text = ScrolledText(frame, wrap=tk.WORD, height=12)
        self.concat_inputs_text.grid(row=1, column=1, sticky="nsew", pady=(6, 0))
        ttk.Label(frame, text="每行一個影片路徑；可只填資料夾，會自動讀取 sence*.mp4。", wraplength=760).grid(
            row=2, column=1, sticky="w", pady=(6, 0)
        )
        frame.rowconfigure(1, weight=1)

    def _add_labeled_entry(self, parent: ttk.Frame, row: int, label: str, variable: tk.StringVar, show: str | None = None) -> ttk.Entry:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=(6, 0))
        entry = ttk.Entry(parent, textvariable=variable, show=show)
        entry.grid(row=row, column=1, sticky="ew", pady=(6, 0))
        return entry

    def _add_labeled_combobox(self, parent: ttk.Frame, row: int, label: str, variable: tk.StringVar, values: list[str]) -> ttk.Combobox:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=(6, 0))
        combo = ttk.Combobox(parent, textvariable=variable, values=values, state="normal")
        combo.grid(row=row, column=1, sticky="ew", pady=(6, 0))
        return combo

    def _add_path_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        save: bool = False,
        directory: bool = False,
        multi: bool = False,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=(6, 0))
        row_frame = ttk.Frame(parent)
        row_frame.grid(row=row, column=1, sticky="ew", pady=(6, 0))
        row_frame.columnconfigure(0, weight=1)
        entry = ttk.Entry(row_frame, textvariable=variable)
        entry.grid(row=0, column=0, sticky="ew")
        button = ttk.Button(
            row_frame,
            text="瀏覽",
            command=lambda: self._browse_path(variable, save=save, directory=directory, multi=multi),
        )
        button.grid(row=0, column=1, padx=(8, 0))

    def _add_multi_path_row(self, parent: ttk.Frame, row: int, label: str, height: int = 4) -> ScrolledText:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="nw", pady=(6, 0))
        row_frame = ttk.Frame(parent)
        row_frame.grid(row=row, column=1, sticky="ew", pady=(6, 0))
        row_frame.columnconfigure(0, weight=1)
        text_widget = ScrolledText(row_frame, wrap=tk.WORD, height=height)
        text_widget.grid(row=0, column=0, sticky="ew")
        button_frame = ttk.Frame(row_frame)
        button_frame.grid(row=0, column=1, sticky="ns", padx=(8, 0))
        ttk.Button(button_frame, text="瀏覽", command=lambda: self._append_multi_paths(text_widget)).grid(
            row=0, column=0, sticky="ew"
        )
        ttk.Button(button_frame, text="清空", command=lambda: text_widget.delete("1.0", tk.END)).grid(
            row=1, column=0, sticky="ew", pady=(8, 0)
        )
        return text_widget

    def _browse_path(self, variable: tk.StringVar, save: bool = False, directory: bool = False, multi: bool = False) -> None:
        if directory:
            path = filedialog.askdirectory(initialdir=str(Path.cwd()))
            if path:
                variable.set(path)
            return
        if save:
            path = filedialog.asksaveasfilename(initialdir=str(Path.cwd()))
        elif multi:
            paths = filedialog.askopenfilenames(initialdir=str(Path.cwd()))
            path = "\n".join(paths)
        else:
            path = filedialog.askopenfilename(initialdir=str(Path.cwd()))
        if path:
            variable.set(path)

    def _append_multi_paths(self, text_widget: ScrolledText) -> None:
        paths = filedialog.askopenfilenames(initialdir=str(Path.cwd()))
        if not paths:
            return
        existing = [line.strip() for line in text_widget.get("1.0", tk.END).splitlines() if line.strip()]
        for path in paths:
            if path not in existing:
                existing.append(path)
        text_widget.delete("1.0", tk.END)
        text_widget.insert("1.0", "\n".join(existing))

    def _current_task(self) -> str:
        selected = self.notebook.select()
        for task, frame in self.task_frames.items():
            if str(frame) == selected:
                return task
        return "template"

    def _common_command(self, task: str) -> list[str]:
        command = [sys.executable, str(APP_PATH), "--task", task]
        if self.output_var.get().strip():
            command.extend(["--output", self.output_var.get().strip()])
        return command

    def _add_references(self, command: list[str], value: str) -> None:
        for line in value.splitlines():
            path = line.strip()
            if path:
                command.extend(["--reference", path])

    def _build_command(self) -> list[str] | None:
        task = self._current_task()
        command = self._common_command(task)
        if task == "template":
            prompt = self.template_prompt_text.get("1.0", tk.END).strip()
            if not prompt:
                messagebox.showerror("缺少欄位", "請輸入模板 prompt。")
                return None
            command.extend(["--model", self.template_model_var.get().strip(), "--size", self.template_size_var.get().strip(), "--prompt", prompt])
            self._add_references(command, self.template_refs_text.get("1.0", tk.END))
            return command
        if task == "scene":
            prompt = self.scene_prompt_text.get("1.0", tk.END).strip()
            if not self.scene_ppt_var.get().strip():
                messagebox.showerror("缺少欄位", "請選擇 PPT 圖片。")
                return None
            if not prompt:
                messagebox.showerror("缺少欄位", "請輸入場景 prompt。")
                return None
            command.extend(["--ppt-image", self.scene_ppt_var.get().strip(), "--model", self.scene_model_var.get().strip(), "--size", self.scene_size_var.get().strip(), "--prompt", prompt])
            if self.scene_template_var.get().strip():
                command.extend(["--template-image", self.scene_template_var.get().strip()])
            if self.scene_number_var.get().strip():
                command.extend(["--scene-number", self.scene_number_var.get().strip()])
            self._add_references(command, self.scene_refs_text.get("1.0", tk.END))
            return command
        if task == "tts":
            text = self.tts_text.get("1.0", tk.END).strip()
            if not text and not self.tts_text_file_var.get().strip():
                messagebox.showerror("缺少欄位", "請輸入文稿內容或選擇文稿檔。")
                return None
            command.extend([
                "--model", self.tts_model_var.get().strip(),
                "--voice", self.tts_voice_var.get().strip(),
                "--audio-format", self.tts_format_var.get().strip(),
                "--speed", self.tts_speed_var.get().strip(),
                "--pitch", self.tts_pitch_var.get().strip(),
                "--volume", self.tts_volume_var.get().strip(),
            ])
            if self.tts_text_file_var.get().strip():
                command.extend(["--text-file", self.tts_text_file_var.get().strip()])
            else:
                command.extend(["--text", text])
            return command
        if task == "video":
            prompt = self.video_prompt_text.get("1.0", tk.END).strip()
            if not self.video_image_var.get().strip():
                messagebox.showerror("缺少欄位", "請選擇起始圖片。")
                return None
            if not prompt:
                messagebox.showerror("缺少欄位", "請輸入影片 prompt。")
                return None
            command.extend([
                "--image", self.video_image_var.get().strip(),
                "--model", self.video_model_var.get().strip(),
                "--resolution", self.video_resolution_var.get().strip(),
                "--poll-interval", self.video_poll_interval_var.get().strip(),
                "--timeout", self.video_timeout_var.get().strip(),
                "--prompt", prompt,
            ])
            if self.video_end_image_var.get().strip():
                command.extend(["--end-image", self.video_end_image_var.get().strip()])
            if self.video_audio_var.get().strip():
                command.extend(["--audio-input", self.video_audio_var.get().strip()])
            if self.video_duration_var.get().strip():
                command.extend(["--duration", self.video_duration_var.get().strip()])
            return command
        if task == "subtitle":
            if not self.subtitle_video_var.get().strip() or not self.subtitle_file_var.get().strip():
                messagebox.showerror("缺少欄位", "請選擇影片與字幕檔。")
                return None
            command.extend(["--video-input", self.subtitle_video_var.get().strip(), "--subtitle-input", self.subtitle_file_var.get().strip()])
            return command
        if task == "compose":
            if not self.compose_video_var.get().strip() or not self.compose_audio_var.get().strip():
                messagebox.showerror("缺少欄位", "請選擇影片與語音音檔。")
                return None
            command.extend(["--video-input", self.compose_video_var.get().strip(), "--audio-input", self.compose_audio_var.get().strip()])
            return command
        if task == "concat":
            if self.concat_input_dir_var.get().strip():
                command.extend(["--input-dir", self.concat_input_dir_var.get().strip()])
            for line in self.concat_inputs_text.get("1.0", tk.END).splitlines():
                path = line.strip()
                if path:
                    command.extend(["--input", path])
            if "--input-dir" not in command and "--input" not in command:
                messagebox.showerror("缺少欄位", "請選擇輸入資料夾或填寫影片清單。")
                return None
            return command
        return None

    def _append_output(self, text: str) -> None:
        self.output_text.insert(tk.END, text)
        self.output_text.see(tk.END)

    def _run_task(self) -> None:
        if self.is_running:
            return
        command = self._build_command()
        if command is None:
            return
        self.is_running = True
        self.run_button.state(["disabled"])
        self.output_text.delete("1.0", tk.END)
        self._append_output("執行命令:\n")
        self._append_output(" ".join(command) + "\n\n")
        threading.Thread(target=self._execute_command, args=(command,), daemon=True).start()

    def _execute_command(self, command: list[str]) -> None:
        try:
            env = os.environ.copy()
            if self.openai_key_var.get().strip():
                env[OPENAI_IMAGE_API_KEY_ENV] = self.openai_key_var.get().strip()
            if self.minimax_key_var.get().strip():
                env["MINIMAX_API_KEY"] = self.minimax_key_var.get().strip()
            result = subprocess.run(command, cwd=str(APP_PATH.parent), capture_output=True, text=True, env=env)
        except Exception as exc:
            self.root.after(0, lambda: self._finish_run(False, f"執行失敗: {exc}\n"))
            return
        output_parts = []
        if result.stdout:
            output_parts.append(result.stdout)
        if result.stderr:
            output_parts.append(result.stderr)
        combined_output = "\n".join(part.strip() for part in output_parts if part.strip())
        if combined_output:
            combined_output += "\n"
        self.root.after(0, lambda: self._finish_run(result.returncode == 0, combined_output))

    def _finish_run(self, success: bool, output: str) -> None:
        if output:
            self._append_output(output)
        self._append_output("\n完成。\n" if success else "\n失敗。\n")
        self.is_running = False
        self.run_button.state(["!disabled"])


def main() -> None:
    if TK_IMPORT_ERROR is not None:
        print(get_tk_install_hint(), file=sys.stderr)
        raise SystemExit(1)
    root = tk.Tk()
    root.minsize(920, 760)
    PptFilmMakerGui(root)
    root.mainloop()


if __name__ == "__main__":
    main()
