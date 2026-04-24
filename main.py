import re
import threading
import time
import tkinter as tk
from dataclasses import dataclass
from tkinter import filedialog, messagebox, ttk

import pyttsx3


WORD_RE = re.compile(r"[A-Za-zА-Яа-яЁё0-9']+")
CYR_RE = re.compile(r"[А-Яа-яЁё]")
LAT_RE = re.compile(r"[A-Za-z]")


def extract_words(text: str) -> list[str]:
    return WORD_RE.findall(text)


def split_sentences(text: str) -> list[str]:
    normalized = re.sub(r"\s+", " ", text.strip())
    if not normalized:
        return []
    parts = re.split(r"(?<=[.!?…])\s+", normalized)
    return [p.strip() for p in parts if p.strip()]


def guess_gender_from_voice(v) -> str | None:
    hay = " ".join(
        [
            str(getattr(v, "name", "")),
            str(getattr(v, "id", "")),
            str(getattr(v, "languages", "")),
        ]
    ).lower()
    if any(x in hay for x in ("female", "жен", "zira", "irina", "svetlana", "tatyana", "anna")):
        return "female"
    if any(x in hay for x in ("male", "муж", "david", "pavel", "ivan", "alex", "sergey")):
        return "male"
    return None


def detect_language(text: str) -> str:
    c = len(CYR_RE.findall(text))
    l = len(LAT_RE.findall(text))
    if c == 0 and l == 0:
        return "ru"
    return "ru" if c >= l else "en"


def voice_matches_language(v, lang: str) -> bool:
    lang = lang.lower()
    hay = " ".join(
        [
            str(getattr(v, "name", "")),
            str(getattr(v, "id", "")),
            str(getattr(v, "languages", "")),
        ]
    ).lower()
    if lang == "ru":
        return any(x in hay for x in ("ru", "rus", "russian", "рус", "рос"))
    if lang == "en":
        return any(x in hay for x in ("en", "eng", "english", "enu", "en-us", "en_gb", "en-"))
    return True


@dataclass
class SpeakSettings:
    rate: int
    pause_s: float
    voice_id: str | None
    lang: str  # auto | ru | en


class DictationApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Диктант (TTS)")
        self.root.geometry("980x720")

        self.engine = pyttsx3.init()
        self.engine_lock = threading.Lock()

        self.voices = list(self.engine.getProperty("voices") or [])
        self.voice_name_by_id: dict[str, str] = {
            getattr(v, "id", ""): getattr(v, "name", getattr(v, "id", "")) for v in self.voices
        }
        self.voice_ids = [getattr(v, "id", "") for v in self.voices if getattr(v, "id", "")]

        self.auto_thread: threading.Thread | None = None
        self.stop_event = threading.Event()
        self.mode_var = tk.StringVar(value="auto")
        self.auto_unit_var = tk.StringVar(value="word")  # word | sentence

        self.rate_var = tk.IntVar(value=175)
        self.pause_var = tk.DoubleVar(value=0.35)
        self.lang_var = tk.StringVar(value="auto")  # auto | ru | en
        self.gender_pref_var = tk.StringVar(value="any")  # any | female | male
        self.voice_var = tk.StringVar(value=self.voice_ids[0] if self.voice_ids else "")

        self.words: list[str] = []
        self.sentences: list[str] = []
        self.prepared_lang = "ru"
        self.word_index = 0

        self._build_ui()
        self._apply_gender_preference()
        self._refresh_voice_label()
        self._update_buttons()

        self.root.bind("<space>", lambda _e: self.on_next_word())
        self.root.bind("<Button-1>", self._maybe_click_next)

    def _reinit_engine(self) -> None:
        prev_voice_id = self.voice_var.get()
        with self.engine_lock:
            try:
                self.engine.stop()
            except Exception:
                pass
            try:
                self.engine = pyttsx3.init()
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("TTS", f"Не удалось инициализировать TTS:\n{e}"))
                return

            self.voices = list(self.engine.getProperty("voices") or [])
            self.voice_name_by_id = {
                getattr(v, "id", ""): getattr(v, "name", getattr(v, "id", "")) for v in self.voices
            }
            self.voice_ids = [getattr(v, "id", "") for v in self.voices if getattr(v, "id", "")]

        if self.voice_ids:
            self.voice_var.set(prev_voice_id if prev_voice_id in self.voice_ids else self.voice_ids[0])

        if hasattr(self, "voice_combo"):
            self.voice_combo["values"] = [self.voice_name_by_id.get(vid, vid) for vid in self.voice_ids]
            self._refresh_voice_label()

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x")

        ttk.Label(top, text="Текст для диктанта").pack(anchor="w")
        self.input_text = tk.Text(top, height=10, wrap="word")
        self.input_text.pack(fill="x", expand=False, pady=(6, 0))

        file_row = ttk.Frame(top)
        file_row.pack(fill="x", pady=(8, 0))
        ttk.Button(file_row, text="Загрузить .txt…", command=self.on_load_file).pack(side="left")
        ttk.Button(file_row, text="Очистить", command=self.on_clear).pack(side="left", padx=(8, 0))
        ttk.Button(file_row, text="Подготовить (разбить на слова)", command=self.on_prepare).pack(
            side="left", padx=(8, 0)
        )

        mid = ttk.Frame(outer)
        mid.pack(fill="x", pady=(14, 0))

        modes = ttk.LabelFrame(mid, text="Режим", padding=10)
        modes.pack(side="left", fill="x", expand=True)
        ttk.Radiobutton(modes, text="Автоматический", value="auto", variable=self.mode_var, command=self._update_buttons).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Radiobutton(modes, text="По клику (по одному слову)", value="click", variable=self.mode_var, command=self._update_buttons).grid(
            row=1, column=0, sticky="w", pady=(6, 0)
        )

        unit = ttk.LabelFrame(mid, text="В авто-режиме читать", padding=10)
        unit.pack(side="left", fill="x", expand=True, padx=(12, 0))
        ttk.Radiobutton(unit, text="по словам", value="word", variable=self.auto_unit_var).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Radiobutton(unit, text="по предложениям", value="sentence", variable=self.auto_unit_var).grid(
            row=1, column=0, sticky="w", pady=(6, 0)
        )

        controls = ttk.LabelFrame(outer, text="Управление", padding=10)
        controls.pack(fill="x", pady=(14, 0))

        ttk.Button(controls, text="Старт авто", command=self.on_start_auto).grid(row=0, column=0, sticky="w")
        ttk.Button(controls, text="Стоп", command=self.on_stop).grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Button(controls, text="Следующее слово (клик/пробел)", command=self.on_next_word).grid(
            row=0, column=2, sticky="w", padx=(8, 0)
        )
        ttk.Button(controls, text="Сначала", command=self.on_reset).grid(row=0, column=3, sticky="w", padx=(8, 0))

        settings = ttk.LabelFrame(outer, text="Настройки TTS", padding=10)
        settings.pack(fill="x", pady=(14, 0))

        ttk.Label(settings, text="Скорость").grid(row=0, column=0, sticky="w")
        self.rate_scale = ttk.Scale(
            settings,
            from_=80,
            to=260,
            orient="horizontal",
            command=self._on_rate_scale,
        )
        self.rate_scale.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.rate_label = ttk.Label(settings, text=str(self.rate_var.get()))
        self.rate_label.grid(row=0, column=2, sticky="w", padx=(8, 0))
        self.rate_scale.set(self.rate_var.get())

        ttk.Label(settings, text="Пауза (сек)").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.pause_spin = ttk.Spinbox(
            settings,
            from_=0.0,
            to=5.0,
            increment=0.05,
            textvariable=self.pause_var,
            width=8,
        )
        self.pause_spin.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(10, 0))

        ttk.Label(settings, text="Язык").grid(row=2, column=0, sticky="w", pady=(10, 0))
        lang_row = ttk.Frame(settings)
        lang_row.grid(row=2, column=1, sticky="w", padx=(8, 0), pady=(10, 0))
        ttk.Radiobutton(lang_row, text="Авто", value="auto", variable=self.lang_var).pack(side="left")
        ttk.Radiobutton(lang_row, text="RU", value="ru", variable=self.lang_var).pack(side="left", padx=(10, 0))
        ttk.Radiobutton(lang_row, text="EN", value="en", variable=self.lang_var).pack(side="left", padx=(10, 0))

        ttk.Label(settings, text="Предпочтительный голос").grid(row=3, column=0, sticky="w", pady=(10, 0))
        gender_row = ttk.Frame(settings)
        gender_row.grid(row=3, column=1, sticky="w", padx=(8, 0), pady=(10, 0))
        ttk.Radiobutton(
            gender_row,
            text="Любой",
            value="any",
            variable=self.gender_pref_var,
            command=self._apply_gender_preference,
        ).pack(side="left")
        ttk.Radiobutton(
            gender_row,
            text="Женский",
            value="female",
            variable=self.gender_pref_var,
            command=self._apply_gender_preference,
        ).pack(side="left", padx=(10, 0))
        ttk.Radiobutton(
            gender_row,
            text="Мужской",
            value="male",
            variable=self.gender_pref_var,
            command=self._apply_gender_preference,
        ).pack(side="left", padx=(10, 0))

        ttk.Label(settings, text="Выбор голоса (если доступно)").grid(row=4, column=0, sticky="w", pady=(10, 0))
        self.voice_combo = ttk.Combobox(
            settings,
            state="readonly",
            values=[self.voice_name_by_id.get(vid, vid) for vid in self.voice_ids],
            width=52,
        )
        self.voice_combo.grid(row=4, column=1, sticky="w", padx=(8, 0), pady=(10, 0))
        self.voice_combo.bind("<<ComboboxSelected>>", self._on_voice_selected)
        self.voice_hint = ttk.Label(settings, text="")
        self.voice_hint.grid(row=4, column=2, sticky="w", padx=(8, 0), pady=(10, 0))

        settings.columnconfigure(1, weight=1)

        view = ttk.LabelFrame(outer, text="Индикатор текущего слова", padding=10)
        view.pack(fill="both", expand=True, pady=(14, 0))

        self.words_view = tk.Text(view, height=10, wrap="word")
        self.words_view.pack(fill="both", expand=True)
        self.words_view.tag_configure("current", background="#fff2a8")
        self.words_view.tag_configure("base", foreground="#111")
        self.words_view.config(state="disabled")

        self.status_var = tk.StringVar(value="Вставьте текст и нажмите «Подготовить».")
        ttk.Label(outer, textvariable=self.status_var).pack(anchor="w", pady=(10, 0))

    def _maybe_click_next(self, _e) -> None:
        if self.mode_var.get() == "click":
            self.on_next_word()

    def _on_rate_scale(self, _value: str) -> None:
        self.rate_var.set(int(float(self.rate_scale.get())))
        if hasattr(self, "rate_label"):
            self.rate_label.config(text=str(self.rate_var.get()))

    def _on_voice_selected(self, _e=None) -> None:
        idx = self.voice_combo.current()
        if idx is None or idx < 0 or idx >= len(self.voice_ids):
            return
        self.voice_var.set(self.voice_ids[idx])
        self._refresh_voice_label()

    def _refresh_voice_label(self) -> None:
        vid = self.voice_var.get()
        name = self.voice_name_by_id.get(vid, vid)
        self.voice_hint.config(text=name if name else "—")

        if self.voice_ids:
            try:
                idx = self.voice_ids.index(vid)
                self.voice_combo.current(idx)
            except ValueError:
                self.voice_combo.current(0)

    def _apply_gender_preference(self) -> None:
        pref = self.gender_pref_var.get()
        if not self.voices or pref == "any":
            if self.voice_ids:
                self.voice_var.set(self.voice_ids[0])
                self._refresh_voice_label()
            return

        chosen_id = None
        for v in self.voices:
            g = guess_gender_from_voice(v)
            if g == pref and getattr(v, "id", ""):
                chosen_id = v.id
                break
        if chosen_id:
            self.voice_var.set(chosen_id)
            self._refresh_voice_label()

    def _update_buttons(self) -> None:
        mode = self.mode_var.get()
        can_next = mode == "click"
        self.status_var.set(
            "Режим по клику: нажимайте «Следующее слово» / пробел / клик по окну."
            if mode == "click"
            else "Авто-режим: нажмите «Старт авто»."
        )

        for child in self.root.winfo_children():
            _ = child
        # (кнопки остаются активными; логика блокировки ниже)

        if not can_next:
            # всё равно разрешаем кнопку "следующее" как быстрый тест, но в статусе подсказываем
            pass

    def on_load_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите текстовый файл",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = f.read()
        except UnicodeDecodeError:
            with open(path, "r", encoding="cp1251", errors="replace") as f:
                data = f.read()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать файл:\n{e}")
            return

        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", data)
        self.status_var.set("Текст загружен. Нажмите «Подготовить».")

    def on_clear(self) -> None:
        self.on_stop()
        self.input_text.delete("1.0", "end")
        self.words = []
        self.sentences = []
        self.word_index = 0
        self._render_words()
        self.status_var.set("Очищено. Вставьте текст и нажмите «Подготовить».")

    def on_prepare(self) -> None:
        self.on_stop()
        text = self.input_text.get("1.0", "end").strip()
        if not text:
            messagebox.showinfo("Текст пустой", "Вставьте или загрузите текст.")
            return

        self.words = extract_words(text)
        self.sentences = split_sentences(text)
        self.prepared_lang = detect_language(text)
        self.word_index = 0

        if not self.words:
            messagebox.showinfo("Нет слов", "Не удалось извлечь слова из текста.")
            return

        self._render_words()
        self._highlight_current_word()
        lang_txt = "RU" if self.prepared_lang == "ru" else "EN"
        self.status_var.set(f"Готово: {len(self.words)} слов. Язык: {lang_txt}. Текущее: 1/{len(self.words)}.")

    def on_reset(self) -> None:
        self.on_stop()
        if not self.words:
            self.status_var.set("Сначала подготовьте текст.")
            return
        self.word_index = 0
        self._highlight_current_word()
        self.status_var.set(f"Сброшено. Текущее: 1/{len(self.words)}.")

    def on_start_auto(self) -> None:
        if not self.words:
            self.on_prepare()
            if not self.words:
                return

        if self.word_index >= len(self.words):
            self.word_index = 0
            self._highlight_current_word()

        self.mode_var.set("auto")
        self._update_buttons()
        self.on_stop()
        self.stop_event.clear()

        self.auto_thread = threading.Thread(target=self._auto_worker, daemon=True)
        self.auto_thread.start()
        self.status_var.set("Авто-режим запущен…")

    def on_stop(self) -> None:
        self.stop_event.set()
        try:
            with self.engine_lock:
                self.engine.stop()
        except Exception:
            pass
        self._reinit_engine()

    def on_next_word(self) -> None:
        if not self.words:
            self.on_prepare()
            if not self.words:
                return

        self.mode_var.set("click")
        self._update_buttons()

        if self.word_index >= len(self.words):
            self.status_var.set("Конец текста. Нажмите «Сначала» чтобы начать заново.")
            return

        w = self.words[self.word_index]
        idx = self.word_index
        self.word_index += 1
        self._highlight_word_index(idx)
        self._speak(w, lang_hint=detect_language(w))
        if self.word_index <= len(self.words):
            self.status_var.set(f"Текущее: {min(self.word_index+1, len(self.words))}/{len(self.words)}.")

    def _settings_snapshot(self) -> SpeakSettings:
        return SpeakSettings(
            rate=int(self.rate_var.get()),
            pause_s=float(self.pause_var.get()),
            voice_id=self.voice_var.get() or None,
            lang=self.lang_var.get(),
        )

    def _choose_voice_for(self, s: SpeakSettings, lang_hint: str | None) -> str | None:
        preferred_lang = s.lang
        lang = lang_hint or self.prepared_lang
        if preferred_lang in ("ru", "en"):
            lang = preferred_lang

        gender_pref = self.gender_pref_var.get()

        # 1) Если пользователь явно выбрал голос в комбобоксе — уважаем выбор.
        if s.voice_id:
            return s.voice_id

        # 2) Подбор голоса по языку/полу.
        best_any = None
        best_gender = None
        for v in self.voices:
            if not getattr(v, "id", ""):
                continue
            if not voice_matches_language(v, lang):
                continue
            best_any = best_any or v.id
            if gender_pref != "any" and guess_gender_from_voice(v) == gender_pref:
                best_gender = best_gender or v.id
        return best_gender or best_any

    def _configure_engine(self, s: SpeakSettings) -> None:
        try:
            self.engine.setProperty("rate", s.rate)
        except Exception:
            pass
        vid = s.voice_id
        if vid:
            try:
                self.engine.setProperty("voice", vid)
            except Exception:
                pass

    def _speak(self, text: str, lang_hint: str | None = None) -> None:
        s = self._settings_snapshot()
        s.voice_id = self._choose_voice_for(s, lang_hint)
        with self.engine_lock:
            self._configure_engine(s)
            self.engine.say(text)
            self.engine.runAndWait()

    def _auto_worker(self) -> None:
        unit = self.auto_unit_var.get()
        s = self._settings_snapshot()

        if unit == "sentence":
            items = self.sentences[:] if self.sentences else []
            if not items:
                items = [" ".join(self.words)]
            for i, sent in enumerate(items):
                if self.stop_event.is_set():
                    break
                self.root.after(0, lambda i=i, n=len(items): self.status_var.set(f"Предложение {i+1}/{n}…"))
                with self.engine_lock:
                    s.voice_id = self._choose_voice_for(s, detect_language(sent))
                    self._configure_engine(s)
                    self.engine.say(sent)
                    self.engine.runAndWait()
                if self.stop_event.is_set():
                    break
                time.sleep(max(0.0, s.pause_s))

            self.root.after(0, lambda: self.status_var.set("Авто-режим завершён."))
            return

        # unit == "word"
        while self.word_index < len(self.words) and not self.stop_event.is_set():
            idx = self.word_index
            w = self.words[idx]
            self.word_index += 1

            self.root.after(0, lambda idx=idx: self._highlight_word_index(idx))
            self.root.after(
                0, lambda idx=idx, n=len(self.words): self.status_var.set(f"Слово {idx+1}/{n}: {self.words[idx]}")
            )

            with self.engine_lock:
                s.voice_id = self._choose_voice_for(s, detect_language(w))
                self._configure_engine(s)
                self.engine.say(w)
                self.engine.runAndWait()

            if self.stop_event.is_set():
                break
            time.sleep(max(0.0, s.pause_s))

        self.root.after(0, lambda: self.status_var.set("Авто-режим завершён."))

    def _render_words(self) -> None:
        self.words_view.config(state="normal")
        self.words_view.delete("1.0", "end")
        if self.words:
            self.words_view.insert("1.0", " ".join(self.words), ("base",))
        self.words_view.config(state="disabled")

    def _highlight_current_word(self) -> None:
        if not self.words:
            return
        idx = min(self.word_index, len(self.words) - 1)
        self._highlight_word_index(idx)

    def _highlight_word_index(self, idx: int) -> None:
        if not self.words:
            return
        idx = max(0, min(idx, len(self.words) - 1))

        self.words_view.config(state="normal")
        self.words_view.tag_remove("current", "1.0", "end")

        # Находим границы слова в тексте "слово слово слово" без пунктуации.
        before = " ".join(self.words[:idx])
        start_char = len(before) + (1 if idx > 0 else 0)
        word_len = len(self.words[idx])

        start = f"1.0+{start_char}c"
        end = f"1.0+{start_char + word_len}c"
        self.words_view.tag_add("current", start, end)
        self.words_view.see(start)
        self.words_view.config(state="disabled")


def main() -> None:
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    DictationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

