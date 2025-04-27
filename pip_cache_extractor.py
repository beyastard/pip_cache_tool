import os
import struct
import threading
import zipfile
import time
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, font, ttk

# Constants
MAX_FILE_SIZE_MB = 384
CHUNK_SIZE = 1024 * 1024
READ_IN_SIZE = 32
OUTPUT_DIR = Path("output")
LOG_DIR = Path("logs")

LIGHT_THEME = {"bg": "#FFFFFF", "fg": "#000000"}
DARK_THEME = {"bg": "#2E2E2E", "fg": "#FFFFFF"}

def get_default_http_cache_root() -> Path:
    return Path(os.environ.get("LOCALAPPDATA", "")) / "pip" / "cache" / "http"

def is_cachecontrol_v4(file_path: Path) -> bool:
    try:
        with open(file_path, "rb") as f:
            return f.read(5) == b"cc=4,"
    except Exception:
        return False

def reconstruct_whl_filename(zip_path: Path) -> Optional[str]:
    with zipfile.ZipFile(zip_path, 'r') as archive:
        dist_info_folders = {name.split('/')[0] for name in archive.namelist() if name.endswith('.dist-info/WHEEL')}
        if not dist_info_folders:
            raise FileNotFoundError("No .dist-info/WHEEL file found in archive.")

        dist_info_folder = dist_info_folders.pop()
        variable_name = dist_info_folder.replace('.dist-info', '')

        wheel_path = f"{dist_info_folder}/WHEEL"
        with archive.open(wheel_path) as wheel_file:
            for line_bytes in wheel_file:
                line = line_bytes.decode('utf-8').strip()
                if line.startswith("Tag:"):
                    tag_name = line[len("Tag:"):].strip()
                    return f"{variable_name}-{tag_name}.whl"
    return None

class CacheExtractorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Pip Cache Extractor")
        self.center_window(625, 510)
        self.root.resizable(False, False)

        self.cache_folder = tk.StringVar(value=str(get_default_http_cache_root()))
        self.output_folder = tk.StringVar(value=str(OUTPUT_DIR.resolve()))

        self.file_list = []
        self.is_dark_mode = False
        self.abort_flag = False

        self.log_file = self.create_log_file()
        self.setup_ui()

        self.write_log("Program started.")

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_log_file(self) -> Path:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        timestamp = int(time.time())
        log_path = LOG_DIR / f"log-{timestamp}.txt"
        return log_path

    def write_log(self, message: str) -> None:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        with open(self.log_file, "a", encoding="utf-8") as log:
            log.write(f"[{timestamp}] {message}\n")

    def center_window(self, width: int, height: int) -> None:
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def setup_ui(self) -> None:
        selected_font = font.Font(family='Arial', size=9)

        tk.Label(self.root, font=selected_font, text='Cache Folder:').place(x=18, y=14)
        tk.Entry(self.root, textvariable=self.cache_folder, width=65).place(x=102, y=18)
        tk.Button(self.root, text='Browse', command=self.browse_cache).place(x=515, y=14, width=90)

        tk.Label(self.root, font=selected_font, text='Output Folder:').place(x=18, y=44)
        tk.Entry(self.root, textvariable=self.output_folder, width=65).place(x=102, y=48)
        tk.Button(self.root, text='Browse', command=self.browse_output).place(x=515, y=44, width=90)

        self.toggle_button = tk.Button(self.root, text='Toggle Dark Mode', command=self.toggle_dark_mode)
        self.toggle_button.place(x=495, y=74, width=110)

        self.status_label = tk.Label(self.root, font=selected_font, text='')
        self.status_label.place(x=18, y=78)

        tk.Button(self.root, text='Load Files', command=self.load_files).place(x=268, y=74, width=90)

        frame = tk.Frame(self.root)
        frame.place(x=5, y=110, width=620, height=280)

        self.listbox = tk.Listbox(frame, selectmode=tk.SINGLE, width=95, height=16)
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        self.listbox.config(yscrollcommand=scrollbar.set)

        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        tk.Button(self.root, text='Extract Selected', command=self.extract_selected_thread).place(x=78, y=400, width=110)
        tk.Button(self.root, text='Extract All', command=self.extract_all_thread).place(x=430, y=400, width=110)

        self.abort_button = tk.Button(self.root, text='Abort', command=self.abort_extraction)
        self.abort_button.place(x=265, y=400, width=90)
        self.abort_button.config(state=tk.DISABLED)

        self.progress_label = tk.Label(self.root, font=selected_font, text='')
        self.progress_label.place(x=10, y=440)

        self.progress = ttk.Progressbar(self.root, orient='horizontal', length=600, mode='determinate')
        self.progress.place(x=10, y=470)

    def toggle_dark_mode(self) -> None:
        self.is_dark_mode = not self.is_dark_mode
        theme = DARK_THEME if self.is_dark_mode else LIGHT_THEME
        self.root.configure(bg=theme["bg"])
        for widget in self.root.winfo_children():
            if isinstance(widget, (tk.Label, tk.Button, tk.Entry, tk.Listbox, tk.Frame)):
                try:
                    widget.configure(bg=theme["bg"], fg=theme["fg"])
                except Exception:
                    pass
        if self.is_dark_mode:
            self.toggle_button.config(text='Toggle Light Mode')
        else:
            self.toggle_button.config(text='Toggle Dark Mode')

    def browse_cache(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.cache_folder.set(folder)

    def browse_output(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder.set(folder)

    def load_files(self) -> None:
        self.listbox.delete(0, tk.END)
        self.file_list.clear()

        base = Path(self.cache_folder.get())
        for file in base.rglob("*"):
            if file.is_file() and is_cachecontrol_v4(file):
                self.file_list.append(file)
                self.listbox.insert(tk.END, str(file.relative_to(base)))

        self.status_label.config(text=f"Loaded {len(self.file_list)} files.")

    def extract_selected_thread(self) -> None:
        threading.Thread(target=self.extract_selected, daemon=True).start()

    def extract_all_thread(self) -> None:
        threading.Thread(target=self.extract_all, daemon=True).start()

    def abort_extraction(self) -> None:
        self.abort_flag = True
        self.write_log("User requested abort.")

    def extract_selected(self) -> None:
        index = self.listbox.curselection()
        if not index:
            messagebox.showwarning("No Selection", "Please select a file to extract.")
            return

        file = self.file_list[index[0]]
        output_file = self.extract_file(file)

        if output_file and output_file.suffix.lower() == ".whl":
            try:
                reconstructed_name = reconstruct_whl_filename(output_file)
                if reconstructed_name:
                    final_path = Path(self.output_folder.get()) / reconstructed_name
                    output_file.rename(final_path)
                    self.write_log(f"{file} -> {final_path}")
            except Exception as e:
                print(f"Failed to reconstruct .whl name: {e}")
        if output_file:
            self.write_log(f"{file} -> {output_file}")

    def extract_all(self) -> None:
        self.abort_flag = False
        total = len(self.file_list)
        self.progress.config(maximum=total, value=0)
        self.abort_button.config(state=tk.NORMAL)

        extracted = 0
        for idx, file in enumerate(self.file_list):
            if self.abort_flag:
                break
            output_file = self.extract_file(file)
            if output_file is None:
                continue
            if output_file.suffix.lower() == ".whl":
                try:
                    reconstructed_name = reconstruct_whl_filename(output_file)
                    if reconstructed_name:
                        final_path = Path(self.output_folder.get()) / reconstructed_name
                        output_file.rename(final_path)
                        self.write_log(f"{file} -> {final_path}")
                        extracted += 1
                        self.progress['value'] = idx + 1
                        self.progress_label.config(text=f"Extracting {idx + 1} of {total} files...")
                        self.root.update_idletasks()
                        continue
                except Exception as e:
                    print(f"Failed to reconstruct .whl name: {e}")
            self.write_log(f"{file} -> {output_file}")
            extracted += 1
            self.progress['value'] = idx + 1
            self.progress_label.config(text=f"Extracting {idx + 1} of {total} files...")
            self.root.update_idletasks()

        self.abort_button.config(state=tk.DISABLED)
        if self.abort_flag:
            self.progress_label.config(text=f"User aborted after {extracted}/{total} files.")
            self.progress['value'] = 0
            self.write_log(f"User aborted after {extracted}/{total} files.")
        else:
            messagebox.showinfo("Done", f"Extracted {len(self.file_list)} files.")
            self.progress_label.config(text=f"Extracted {len(self.file_list)} files.")
            self.write_log(f"Extracted {len(self.file_list)} files.")

    def extract_file(self, file: Path) -> Optional[Path]:
        try:
            if file.stat().st_size > MAX_FILE_SIZE_MB * 1024 * 1024:
                proceed = messagebox.askyesno("Large File", f"{file.name} is large. Extract anyway?")
                if not proceed:
                    return None

            with open(file, "rb") as f:
                header = f.read(READ_IN_SIZE)
                if not header.startswith(b"cc=4"):
                    return None

                indicator = header[0x15]
                if indicator == 0xC5:
                    f.seek(0x16)
                    body_length = struct.unpack(">H", f.read(2))[0]
                    body_offset = 0x18
                elif indicator == 0xC6:
                    f.seek(0x16)
                    body_length = struct.unpack(">I", f.read(4))[0]
                    body_offset = 0x1A
                else:
                    print(f"Unknown format in {file.name}")
                    return None

                f.seek(body_offset)
                body = f.read(body_length)

            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            default_name = file.name.replace(os.sep, "-")
            final_name = self.detect_file_type(body, default_name)
            out_path = Path(self.output_folder.get()) / final_name

            with open(out_path, "wb") as out_file:
                out_file.write(body)

            return out_path

        except Exception as e:
            print(f"Failed to extract {file}: {e}")
            return None

    def detect_file_type(self, body: bytes, default_name: str) -> str:
        try:
            if body.startswith(b'PK\x03\x04'):
                return default_name + ".whl"
            elif body.startswith(b'\x1f\x8b\x08\x00'):
                return default_name + ".gz"
            elif body.startswith(b'\x1f\x8b\x08\x08'):
                return default_name + ".tgz"
            elif body.startswith(b'Metadata-Version'):
                text = body.decode("utf-8")
                lines = text.splitlines()
                name, version, python_version = None, None, None

                for line in lines:
                    if line.startswith("Name:"):
                        name = line.split(":", 1)[1].strip()
                    elif line.startswith("Version:"):
                        version = line.split(":", 1)[1].strip()
                    elif line.startswith("Classifier: Programming Language :: Python ::"):
                        python_version = line.split("::")[-1].strip()

                if name and version and python_version:
                    return f"{name}-{version}-py{python_version}.metadata.txt"
                elif name and version:
                    return f"{name}-{version}-py3-none-any.metadata.txt"
        except Exception as e:
            print(f"Metadata parsing error: {e}")

        return default_name + ".dat"

    def on_closing(self) -> None:
        self.write_log("Program closed.")
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = CacheExtractorApp(root)
    root.mainloop()
