from __future__ import annotations

import hashlib
import math
import os
import random
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import cv2
except Exception:
    cv2 = None

try:
    import fitz
except Exception:
    fitz = None

try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk
except Exception:
    Image = ImageDraw = ImageFont = ImageTk = None


APP_NAME = "Fileviewer Pro"
SLIDE_SECONDS = 15

IMAGE_EXTENSIONS = {
    ".jpg",
    ".jpeg",
    ".png",
    ".bmp",
    ".gif",
    ".tif",
    ".tiff",
    ".webp",
}
VIDEO_EXTENSIONS = {
    ".mp4",
    ".mkv",
    ".mov",
    ".avi",
    ".wmv",
    ".m4v",
    ".webm",
}
PDF_EXTENSIONS = {".pdf"}
WORD_EXTENSIONS = {".doc", ".docx"}
POWERPOINT_EXTENSIONS = {".ppt", ".pptx"}
DOCUMENT_EXTENSIONS = PDF_EXTENSIONS | WORD_EXTENSIONS | POWERPOINT_EXTENSIONS


@dataclass(frozen=True)
class Slide:
    kind: str
    source: Path
    label: str
    duration: Optional[float] = SLIDE_SECONDS
    category: Optional[str] = None
    page_index: Optional[int] = None
    page_count: Optional[int] = None
    text: Optional[str] = None
    bg: Optional[str] = None
    fg: str = "#ffffff"


class FileScanner:
    def __init__(self, folder: Path, cache_dir: Path, log):
        self.folder = folder
        self.cache_dir = cache_dir
        self.log = log

    def scan(self) -> list[Slide]:
        slides: list[Slide] = []
        if not self.folder.exists():
            return slides

        files = [p for p in self.folder.iterdir() if p.is_file()]
        for path in files:
            suffix = path.suffix.lower()
            if path.name.lower() == "quotes.txt":
                slides.extend(self._quote_slides(path))
            elif suffix in IMAGE_EXTENSIONS:
                slides.append(
                    Slide(
                        "image",
                        path,
                        path.name,
                        duration=SLIDE_SECONDS,
                        category="image",
                    )
                )
            elif suffix in VIDEO_EXTENSIONS:
                slides.append(
                    Slide("video", path, path.name, duration=None, category="video")
                )
            elif suffix in DOCUMENT_EXTENSIONS:
                slides.extend(self._document_slides(path))
            else:
                self.log(f"Overgeslagen: {path.name} (niet ondersteund)")

        return slides

    def _quote_slides(self, path: Path) -> list[Slide]:
        try:
            text = path.read_text(encoding="utf-8-sig")
        except UnicodeDecodeError:
            text = path.read_text(encoding="cp1252", errors="replace")

        blocks = []
        current = []
        for line in text.splitlines():
            if line.strip():
                current.append(line.rstrip())
            elif current:
                blocks.append("\n".join(current).strip())
                current = []
        if current:
            blocks.append("\n".join(current).strip())

        slides = []
        for i, block in enumerate(blocks, start=1):
            bg, fg = choose_quote_colors(block)
            slides.append(
                Slide(
                    "quote",
                    path,
                    f"Quote {i}",
                    duration=SLIDE_SECONDS,
                    category="quote",
                    text=block,
                    bg=bg,
                    fg=fg,
                )
            )
        self.log(f"Quotes geladen: {len(slides)} blokken")
        return slides

    def _document_slides(self, path: Path) -> list[Slide]:
        pdf_path = path
        suffix = path.suffix.lower()
        category = "pdf"
        if suffix in WORD_EXTENSIONS | POWERPOINT_EXTENSIONS:
            category = "word" if suffix in WORD_EXTENSIONS else "presentation"
            pdf_path = self._convert_office_to_pdf(path)
            if pdf_path is None:
                self.log(
                    f"Overgeslagen: {path.name} (Office-conversie naar PDF niet gelukt)"
                )
                return []

        if fitz is None:
            self.log(f"Overgeslagen: {path.name} (PyMuPDF ontbreekt)")
            return []

        try:
            with fitz.open(str(pdf_path)) as doc:
                page_count = doc.page_count
        except Exception as exc:
            self.log(f"Overgeslagen: {path.name} ({exc})")
            return []

        slides = []
        for page_index in range(page_count):
            slides.append(
                Slide(
                    "document",
                    pdf_path,
                    f"{path.name} - pagina {page_index + 1}/{page_count}",
                    duration=SLIDE_SECONDS,
                    category=category,
                    page_index=page_index,
                    page_count=page_count,
                )
            )
        self.log(f"Document geladen: {path.name} ({page_count} pagina's)")
        return slides

    def _convert_office_to_pdf(self, path: Path) -> Optional[Path]:
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        key = hashlib.sha1(
            f"{path.resolve()}|{path.stat().st_mtime_ns}|{path.stat().st_size}".encode(
                "utf-8", errors="ignore"
            )
        ).hexdigest()
        pdf_path = self.cache_dir / f"{key}.pdf"
        if pdf_path.exists():
            return pdf_path

        suffix = path.suffix.lower()
        try:
            if suffix in WORD_EXTENSIONS:
                return convert_word_to_pdf(path, pdf_path, self.log)
            if suffix in POWERPOINT_EXTENSIONS:
                return convert_powerpoint_to_pdf(path, pdf_path, self.log)
        except Exception as exc:
            self.log(f"Conversie mislukt voor {path.name}: {exc}")
        return None


def convert_word_to_pdf(path: Path, pdf_path: Path, log) -> Optional[Path]:
    try:
        import pythoncom
        import win32com.client
    except Exception as exc:
        log(f"pywin32 is nodig voor Word-conversie: {exc}")
        return None

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
        doc.ExportAsFixedFormat(str(pdf_path.resolve()), 17)
        return pdf_path if pdf_path.exists() else None
    finally:
        if doc is not None:
            doc.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


def convert_powerpoint_to_pdf(path: Path, pdf_path: Path, log) -> Optional[Path]:
    try:
        import pythoncom
        import win32com.client
    except Exception as exc:
        log(f"pywin32 is nodig voor PowerPoint-conversie: {exc}")
        return None

    pythoncom.CoInitialize()
    app = None
    presentation = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        presentation = app.Presentations.Open(
            str(path.resolve()), WithWindow=False, ReadOnly=True
        )
        presentation.SaveAs(str(pdf_path.resolve()), 32)
        return pdf_path if pdf_path.exists() else None
    finally:
        if presentation is not None:
            presentation.Close()
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()


def choose_quote_colors(text: str) -> tuple[str, str]:
    palettes = {
        "calm": [("#12343b", "#f5f0e6"), ("#2f4858", "#f7f7f2"), ("#2e4f4f", "#f4f1de")],
        "warm": [("#8f3f2f", "#fff4e6"), ("#6c3428", "#fff1d6"), ("#9a3412", "#fff7ed")],
        "bright": [("#0f766e", "#f0fdfa"), ("#5b21b6", "#f5f3ff"), ("#1d4ed8", "#eff6ff")],
        "soft": [("#3f3cbb", "#eef2ff"), ("#7c2d12", "#ffedd5"), ("#166534", "#f0fdf4")],
    }
    lowered = text.lower()
    if any(word in lowered for word in ["rust", "stil", "adem", "water", "nacht"]):
        group = "calm"
    elif any(word in lowered for word in ["liefde", "hart", "warm", "zon", "dank"]):
        group = "warm"
    elif any(word in lowered for word in ["kracht", "durf", "win", "actie", "energie"]):
        group = "bright"
    else:
        group = random.choice(list(palettes))
    return random.choice(palettes[group])


class SlideDeck:
    CATEGORY_WEIGHTS = {
        "quote": 4,
        "image": 4,
        "pdf": 1,
        "word": 1,
        "presentation": 1,
        "video": 1,
    }

    def __init__(self, slides: list[Slide]):
        self.slides = slides
        self.history: list[Slide] = []
        self.current: Optional[Slide] = None
        self.forward_stack: list[Slide] = []
        self.groups: dict[Path, list[Slide]] = {}
        self.source_categories: dict[Path, str] = {}
        for slide in slides:
            self.groups.setdefault(slide.source, []).append(slide)
            self.source_categories.setdefault(slide.source, slide.category or slide.kind)
        self.sources_by_category: dict[str, list[Path]] = {}
        for source, category in self.source_categories.items():
            self.sources_by_category.setdefault(category, []).append(source)
        self.category_bag: list[str] = []
        self.source_bags: dict[str, list[Path]] = {}

    def next(self, manual: bool = False) -> Optional[Slide]:
        if manual and self.forward_stack:
            return self._set_current(self.forward_stack.pop())

        doc_next = self._next_document_page()
        if doc_next is not None:
            return self._set_current(doc_next)

        return self._set_current(self._random_slide())

    def previous(self) -> Optional[Slide]:
        if len(self.history) < 2:
            return self.current
        if self.current is not None:
            self.forward_stack.append(self.current)
        self.history.pop()
        self.current = self.history[-1]
        return self.current

    def _set_current(self, slide: Optional[Slide]) -> Optional[Slide]:
        if slide is None:
            return None
        self.current = slide
        if not self.history or self.history[-1] != slide:
            self.history.append(slide)
        return slide

    def _next_document_page(self) -> Optional[Slide]:
        if self.current is None or self.current.kind != "document":
            return None
        if self.current.page_index is None or self.current.page_count is None:
            return None
        next_index = self.current.page_index + 1
        if next_index >= self.current.page_count:
            return None
        for slide in self.slides:
            if (
                slide.kind == "document"
                and slide.source == self.current.source
                and slide.page_index == next_index
            ):
                return slide
        return None

    def _random_slide(self) -> Optional[Slide]:
        if not self.sources_by_category:
            return None
        category = self._next_category()
        if category is None:
            return None
        source = self._next_source(category)
        if source is None:
            return None

        choices = self.groups[source]
        first = choices[0]
        if first.kind == "quote":
            return random.choice(choices)
        if first.kind == "document":
            return next((slide for slide in choices if slide.page_index == 0), first)
        return first

    def _next_category(self) -> Optional[str]:
        if not self.category_bag:
            self._refill_category_bag()

        current_category = self.current.category if self.current is not None else None
        if current_category is not None and len(set(self.category_bag)) > 1:
            for index in range(len(self.category_bag) - 1, -1, -1):
                if self.category_bag[index] != current_category:
                    return self.category_bag.pop(index)
        if (
            current_category is not None
            and len(self.sources_by_category) > 1
            and self.category_bag
            and set(self.category_bag) == {current_category}
        ):
            self._refill_category_bag()
            for index in range(len(self.category_bag) - 1, -1, -1):
                if self.category_bag[index] != current_category:
                    return self.category_bag.pop(index)
        return self.category_bag.pop() if self.category_bag else None

    def _refill_category_bag(self):
        self.category_bag = []
        for category in self.sources_by_category:
            weight = self.CATEGORY_WEIGHTS.get(category, 1)
            self.category_bag.extend([category] * weight)
        random.shuffle(self.category_bag)

    def _next_source(self, category: str) -> Optional[Path]:
        sources = self.sources_by_category.get(category, [])
        if not sources:
            return None
        bag = self.source_bags.setdefault(category, [])
        if not bag:
            bag.extend(sources)
            random.shuffle(bag)

        current_source = self.current.source if self.current is not None else None
        if current_source is not None and len(bag) > 1:
            for index in range(len(bag) - 1, -1, -1):
                if bag[index] != current_source:
                    return bag.pop(index)
        return bag.pop()


class FullscreenViewer(tk.Toplevel):
    def __init__(self, master: tk.Tk, slides: list[Slide], on_exit):
        super().__init__(master)
        self.on_exit = on_exit
        self.deck = SlideDeck(slides)
        self.paused = False
        self.after_id: Optional[str] = None
        self.video_after_id: Optional[str] = None
        self.video_capture = None
        self.video_photo = None
        self.current_photo = None
        self.current_slide: Optional[Slide] = None
        self.slide_started_at = 0.0
        self.remaining = SLIDE_SECONDS

        self.configure(bg="black")
        self.attributes("-fullscreen", True)
        self.bind("<Right>", lambda _event: self.show_next(manual=True))
        self.bind("<Left>", lambda _event: self.show_previous())
        self.bind("<Down>", lambda _event: self.pause())
        self.bind("<Up>", lambda _event: self.resume())
        self.bind("<Return>", lambda _event: self.exit_viewer())
        self.bind("<Escape>", lambda _event: self.exit_viewer())

        self.display = tk.Label(self, bg="black", fg="white", bd=0)
        self.display.pack(fill="both", expand=True)
        self.focus_set()

        self.show_next()

    def show_next(self, manual: bool = False):
        self._clear_timers()
        self._stop_video()
        slide = self.deck.next(manual=manual)
        self._show_slide(slide)

    def show_previous(self):
        self._clear_timers()
        self._stop_video()
        slide = self.deck.previous()
        self._show_slide(slide)

    def pause(self):
        if self.paused:
            return
        self.paused = True
        self._clear_timers()
        if self.current_slide and self.current_slide.duration:
            elapsed = time.monotonic() - self.slide_started_at
            self.remaining = max(0.1, self.current_slide.duration - elapsed)

    def resume(self):
        if not self.paused:
            return
        self.paused = False
        if self.current_slide and self.current_slide.kind == "video":
            self._play_video_frame()
        elif self.current_slide:
            self.slide_started_at = time.monotonic()
            self.after_id = self.after(int(self.remaining * 1000), self.show_next)
        self._set_status(self.current_slide)

    def exit_viewer(self):
        self._clear_timers()
        self._stop_video()
        self.destroy()
        self.on_exit()

    def _show_slide(self, slide: Optional[Slide]):
        self.current_slide = slide
        if slide is None:
            self.display.configure(
                text="Geen bestanden gevonden", image="", bg="black", fg="white"
            )
            return

        self._set_status(slide)
        self.slide_started_at = time.monotonic()
        self.remaining = slide.duration or SLIDE_SECONDS

        if slide.kind == "image":
            self._show_image(slide.source)
            self._schedule(slide.duration)
        elif slide.kind == "document":
            self._show_pdf_page(slide)
            self._schedule(slide.duration)
        elif slide.kind == "quote":
            self._show_quote(slide)
            self._schedule(slide.duration)
        elif slide.kind == "video":
            self._show_video(slide.source)
        else:
            self.show_next()

    def _schedule(self, duration: Optional[float]):
        if duration and not self.paused:
            self.after_id = self.after(int(duration * 1000), self.show_next)

    def _show_image(self, path: Path):
        if Image is None:
            self.display.configure(text=f"Pillow ontbreekt:\n{path.name}", image="")
            return
        try:
            with Image.open(path) as img:
                self._display_pil_image(img.convert("RGB"), bg="black")
        except Exception as exc:
            self.display.configure(text=f"Kan afbeelding niet openen:\n{path.name}\n{exc}", image="")

    def _show_pdf_page(self, slide: Slide):
        if fitz is None or Image is None:
            self.display.configure(text="PDF-rendering is niet beschikbaar", image="")
            return
        try:
            width = max(1, self.winfo_screenwidth())
            height = max(1, self.winfo_screenheight())
            with fitz.open(str(slide.source)) as doc:
                page = doc.load_page(slide.page_index or 0)
                rect = page.rect
                zoom = min(width / rect.width, height / rect.height) * 1.9
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            self._display_pil_image(img, bg="#111111")
        except Exception as exc:
            self.display.configure(
                text=f"Kan pagina niet tonen:\n{slide.label}\n{exc}", image=""
            )

    def _show_quote(self, slide: Slide):
        self.configure(bg=slide.bg or "#111111")
        self.display.configure(bg=slide.bg or "#111111", image="", text="")
        width = max(800, self.winfo_screenwidth())
        height = max(600, self.winfo_screenheight())
        img = Image.new("RGB", (width, height), slide.bg or "#111111")
        draw = ImageDraw.Draw(img)
        font = load_font(max(34, min(74, width // 22)), bold=False)
        text = slide.text or ""
        wrapped = wrap_text(draw, text, font, int(width * 0.74))
        bbox = draw.multiline_textbbox((0, 0), wrapped, font=font, spacing=16)
        x = (width - (bbox[2] - bbox[0])) / 2
        y = (height - (bbox[3] - bbox[1])) / 2
        draw.multiline_text(
            (x, y),
            wrapped,
            font=font,
            fill=slide.fg,
            spacing=16,
            align="center",
        )
        self._display_pil_image(img, bg=slide.bg or "#111111", exact=True)

    def _show_video(self, path: Path):
        if cv2 is None:
            self.display.configure(text=f"OpenCV ontbreekt:\n{path.name}", image="")
            self._schedule(SLIDE_SECONDS)
            return
        self.video_capture = cv2.VideoCapture(str(path))
        if not self.video_capture.isOpened():
            self.display.configure(text=f"Kan video niet openen:\n{path.name}", image="")
            self._schedule(SLIDE_SECONDS)
            return
        self._play_video_frame()

    def _play_video_frame(self):
        if self.paused or self.video_capture is None:
            return
        ok, frame = self.video_capture.read()
        if not ok:
            self.show_next()
            return
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(frame)
        self._display_pil_image(img, bg="black")
        fps = self.video_capture.get(cv2.CAP_PROP_FPS) or 25
        delay = max(1, int(1000 / fps))
        self.video_after_id = self.after(delay, self._play_video_frame)

    def _display_pil_image(self, img, bg: str, exact: bool = False):
        self.configure(bg=bg)
        self.display.configure(bg=bg, text="")
        screen_w = max(1, self.winfo_screenwidth())
        screen_h = max(1, self.winfo_screenheight())
        if not exact:
            img = fit_image_to_screen(img, screen_w, screen_h, bg)
        self.current_photo = ImageTk.PhotoImage(img)
        self.display.configure(image=self.current_photo)

    def _set_status(self, slide: Optional[Slide]):
        return

    def _clear_timers(self):
        for timer in (self.after_id, self.video_after_id):
            if timer is not None:
                try:
                    self.after_cancel(timer)
                except Exception:
                    pass
        self.after_id = None
        self.video_after_id = None

    def _stop_video(self):
        if self.video_capture is not None:
            try:
                self.video_capture.release()
            except Exception:
                pass
        self.video_capture = None


def fit_image_to_screen(img, width: int, height: int, bg: str):
    canvas = Image.new("RGB", (width, height), bg)
    working = img.copy()
    working.thumbnail((width, height), Image.Resampling.LANCZOS)
    x = (width - working.width) // 2
    y = (height - working.height) // 2
    canvas.paste(working, (x, y))
    return canvas


def load_font(size: int, bold: bool = False):
    if ImageFont is None:
        return None
    candidates = [
        "C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size=size)
    return ImageFont.load_default()


def wrap_text(draw, text: str, font, max_width: int) -> str:
    lines = []
    for paragraph in text.splitlines():
        words = paragraph.split()
        if not words:
            lines.append("")
            continue
        line = ""
        for word in words:
            candidate = word if not line else f"{line} {word}"
            bbox = draw.textbbox((0, 0), candidate, font=font)
            if bbox[2] - bbox[0] <= max_width:
                line = candidate
            else:
                if line:
                    lines.append(line)
                line = word
        if line:
            lines.append(line)
    return "\n".join(lines)


class FileviewerProApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("760x520")
        self.root.minsize(680, 440)

        self.folder_var = tk.StringVar(value=str(Path.cwd() / "Data"))
        self.status_var = tk.StringVar(value="Kies een map en start de viewer.")
        self.cache_dir = Path.cwd() / ".fileviewer_cache"

        self._build_gui()
        self._check_dependencies()

    def _build_gui(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", padding=(12, 8))
        style.configure("Title.TLabel", font=("Segoe UI", 24, "bold"))
        style.configure("Hint.TLabel", font=("Segoe UI", 10), foreground="#555555")

        outer = ttk.Frame(self.root, padding=24)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Fullscreen random viewer voor afbeeldingen, video, PDF, Word, PowerPoint en quotes.",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(4, 24))

        folder_frame = ttk.Frame(outer)
        folder_frame.pack(fill="x")
        ttk.Label(folder_frame, text="Map").pack(anchor="w")
        row = ttk.Frame(folder_frame)
        row.pack(fill="x", pady=(6, 0))
        ttk.Entry(row, textvariable=self.folder_var).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(row, text="Kiezen", command=self.choose_folder).pack(
            side="left", padx=(10, 0)
        )

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=22)
        ttk.Button(actions, text="Start viewer", command=self.start_viewer).pack(
            side="left"
        )
        ttk.Button(actions, text="Scan map", command=self.scan_only).pack(
            side="left", padx=(10, 0)
        )

        help_box = ttk.LabelFrame(outer, text="Bediening fullscreen", padding=14)
        help_box.pack(fill="x", pady=(0, 16))
        ttk.Label(
            help_box,
            text=(
                "Rechts: volgende | Links: vorige | Omlaag: pauze | "
                "Omhoog: verder | Enter/Escape: terug naar dit scherm"
            ),
        ).pack(anchor="w")

        log_frame = ttk.LabelFrame(outer, text="Status", padding=10)
        log_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_frame, height=10, wrap="word")
        self.log_text.pack(fill="both", expand=True)
        self.log_text.configure(state="disabled")

        ttk.Label(outer, textvariable=self.status_var, style="Hint.TLabel").pack(
            anchor="w", pady=(10, 0)
        )

    def choose_folder(self):
        folder = filedialog.askdirectory(initialdir=self.folder_var.get() or str(Path.cwd()))
        if folder:
            self.folder_var.set(folder)
            self.scan_only()

    def scan_only(self):
        slides = self._scan()
        self.status_var.set(f"{len(slides)} slides gevonden.")

    def start_viewer(self):
        slides = self._scan()
        if not slides:
            messagebox.showwarning(APP_NAME, "Geen ondersteunde bestanden gevonden.")
            return
        random.shuffle(slides)
        self.root.withdraw()
        FullscreenViewer(self.root, slides, self._return_to_gui)

    def _return_to_gui(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def _scan(self) -> list[Slide]:
        self.clear_log()
        folder = Path(self.folder_var.get()).expanduser()
        scanner = FileScanner(folder, self.cache_dir, self.log)
        slides = scanner.scan()
        counts = {}
        for slide in slides:
            counts[slide.kind] = counts.get(slide.kind, 0) + 1
        summary = ", ".join(f"{kind}: {count}" for kind, count in sorted(counts.items()))
        self.log(summary or "Geen ondersteunde bestanden gevonden.")
        return slides

    def _check_dependencies(self):
        missing = []
        if Image is None or ImageTk is None:
            missing.append("Pillow")
        if fitz is None:
            missing.append("PyMuPDF")
        if cv2 is None:
            missing.append("opencv-python")
        if missing:
            self.log("Ontbrekende pakketten: " + ", ".join(missing))

    def log(self, message: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")


def main():
    root = tk.Tk()
    app = FileviewerProApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
