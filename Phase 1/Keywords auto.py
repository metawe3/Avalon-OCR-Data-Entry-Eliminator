import tkinter as tk
from tkinter import ttk
import ctypes
import re
import time
import threading

import pyautogui
import win32gui
import win32api
import win32con

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0

# ---- Strata Type buttons ---------------------------------------------------
STRATA_GROUPS = {
    "Most common": [
        "SAND", "Silty SAND", "Clayey SAND", "SAND&CLAY", "Clay", "Sandy CLAY",
        "Silty CLAY", "CLAY&SAND", "Halite", "Fill", "Silt", "Sandstone",
        "GRAVEL", "limestone",
    ],
    "Other / less common": [
        "Sandstone", "Basalt", "Conglomerate", "Silt", "Silt Sandy",
        "Backfilling", "Clay stone", "Clay & Silt", "Peat & Organic",
        "Gravel & Sand", "Limestone fragments", "SAND&SILT", "SAND & Gravel",
        "Gravelly SAND"
    ],
}

def _norm(s):
    """Normalize a strata label for lookup - case and space insensitive."""
    return re.sub(r"\s+", "", s).upper()

STRATA_MODE_CANON = {
    "SAND": "sand", "SILTYSAND": "sand", "SAND&CLAY": "sand", "GRAVEL": "sand",
    "CLAYEYSAND": "sand", "GRAVEL&SAND": "sand", "SAND&SILT": "sand",
    "SILT": "sand", "SILTSANDY": "sand",
    "SILTYCLAY": "clay", "CLAY&SAND": "clay", "CLAY": "clay",
    "SANDYCLAY": "clay", "CLAYSTONE": "clay", "CLAY&SILT": "clay",
    "PEAT&ORGANIC": "clay",
    "LIMESTONE": "rock", "LIMESTONEFRAGMENTS": "rock", "SANDSTONE": "rock",
    "BASALT": "rock", "CONGLOMERATE": "rock",
    "HALITE": "special", "FILL": "special", "LOSSSAMPLE": "special",
    "BACKFILLING": "special",
}

STRATA_DESC_STARTER_CANON = {
    "SAND": "SAND,",
    "SILTYSAND": "SAND,",
    "CLAYEYSAND": "SAND,",
    "SAND&CLAY": "SAND & CLAY,",
    "CLAY": "CLAY,",
    "SANDYCLAY": "CLAY,",
    "SILTYCLAY": "CLAY,",
    "CLAY&SAND": "CLAY & SAND,",
    "HALITE": "HALITE,",
    "FILL": "FILL:",
    "SILT": "SILT,",
    "SANDSTONE": "SANDSTONE,",
    "GRAVEL": "GRAVEL & SAND,",
    "GRAVEL&SAND": "GRAVEL & SAND,",
    "LIMESTONE": "LIMESTONE,",
    "LOSSSAMPLE": "LOSS SAMPLE",
}

def get_mode(stratum_label):
    return STRATA_MODE_CANON.get(_norm(stratum_label), "sand")

def get_starter(stratum_label):
    key = _norm(stratum_label)
    if key in STRATA_DESC_STARTER_CANON:
        return STRATA_DESC_STARTER_CANON[key]
    return stratum_label.upper() + ","

# ---- Universal Description word groups -------------------------------------
UNIVERSAL_GROUPS = {
    "Colors": [
        "Light", "Dark", "Pale", "Yellow", "White", "Pink", "Red", "Brown",
        "Olive", "Green", "Blue", "Black", "Grey", "Reddish", "Pinkish",
        "Yellowish", "Brownish", "Greenish", "Bluish", "Grayish",
    ],
    "Texture": ["Silty", "Sandy", "Gravelly", "Clayey", "Calcareous"],
    "Grain size": ["Fine", "to", "Medium", "Coarse", "Grained"],
    "Connectors": ["Traces of", "Some", "With", "Without", "&"],
    "Inclusions & Materials": [
        "Sand", "Gravel", "Clay", "Silt", "Rock Fragments", "Limestone Fragments",
        "Limestone Pieces", "Sandstone Fragments", "Sandstone Pieces",
        "Cemented Sand", "Organic Sand", "Organic Matter", "Shell Fragments",
        "Clay Pocket", "Iron Oxides", "Calcareous Materials", "Calcareous Pebbles",
        "Mica", "Gypsum", "Rock Salt", "Fossils", "Vugs",
    ],
}

STRUCTURE_WORDS = ["Fissured", "Soapy Touch", "Interbedded by", "Intercalation of"]
QUICK_PHRASE_COMBOS = ["Traces of Iron Oxides", "Iron Oxides & Calcareous Materials"]

MODE_GROUPS = {
    "sand": ("Density / Grading", ["Very", "Loose", "Medium", "Dense", "Poorly Graded"]),
    "clay": ("Consistency", ["Very", "Soft", "Medium", "Stiff", "Hard"]),
    "rock": ("Rock Quality", [
        "Highly Weathered", "Moderately Weathered", "Slightly Weathered",
        "Fragmented", "Weak", "Strong", "Vuggy", "Fossiliferous",
    ]),
    "special": ("", []),
}

STEP_KEYS = ["colors", "texture", "grain", "connectors"]
STEP_TITLES = {
    "colors": "Step 1/4: Colors",
    "texture": "Step 2/4: Texture",
    "grain": "Step 3/4: Grain Size",
    "connectors": "Step 4/4: Connectors & Inclusions",
}
# ---------------------------------------------------------------------------

kernel32 = ctypes.windll.kernel32


class ScrollableFrame(ttk.Frame):
    """A frame with a vertical scrollbar."""

    def __init__(self, parent):
        super().__init__(parent)
        canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)

        self.inner.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)


class EntryHelper(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Quick Entry Helper")
        self.attributes("-topmost", True)
        self.resizable(True, True)
        self.geometry("700x680")
        self.minsize(380, 420)

        self.current_mode = "sand"
        self.step_index = 0
        self.other_visible = False
        self.pending_starter = None

        self._build_ui()

        # Apply WS_EX_NOACTIVATE after the window is fully drawn so that
        # clicking any button never steals focus from Strater.
        self.after(100, self._make_non_activating)

    # -- Make the window non-focus-stealing ----------------------------------
    def _make_non_activating(self):
        """Set WS_EX_NOACTIVATE on the helper window so button clicks never
        pull focus away from Strater's active cell."""
        self.update_idletasks()
        hwnd = win32gui.FindWindow(None, "Quick Entry Helper")
        if hwnd:
            style = win32api.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            win32api.SetWindowLong(
                hwnd,
                win32con.GWL_EXSTYLE,
                style | win32con.WS_EX_NOACTIVATE | win32con.WS_EX_TOOLWINDOW,
            )
            self.status.config(
                text="Ready. Click into Strater, then pick a stratum here.",
                foreground="green",
            )

    # -- UI ------------------------------------------------------------------
    def _build_ui(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(padx=6, pady=6, fill="both", expand=True)

        self.strata_frame = ttk.Frame(self.notebook)
        self.desc_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.strata_frame, text="Strata Type")
        self.notebook.add(self.desc_frame, text="Description")

        self._build_strata_tab()
        self._build_description_tab()

        self.status = ttk.Label(
            self,
            text="Starting up...",
            foreground="gray",
        )
        self.status.pack(padx=6, pady=(0, 2), anchor="w")

        bottom_bar = ttk.Frame(self)
        bottom_bar.pack(side="bottom", fill="x", padx=6, pady=6)

        self.desc_extra_frame = ttk.Frame(bottom_bar)
        self.desc_extra_frame.pack(side="right", padx=(6, 0))
        tk.Button(
            self.desc_extra_frame, text=",", font=("Segoe UI", 10, "bold"), width=3,
            command=self.on_comma_click,
        ).pack(side="left", padx=(0, 3))
        tk.Button(
            self.desc_extra_frame, text="\u2715", font=("Segoe UI", 10, "bold"), width=3, fg="red",
            command=self.on_null_click,
        ).pack(side="left")
        self.desc_extra_frame.pack_forget()

        tk.Button(bottom_bar, text="Next", font=("Segoe UI", 10, "bold"), command=self.on_next_click)\
            .pack(side="left", expand=True, fill="x", padx=(0, 3))
        tk.Button(bottom_bar, text="End", font=("Segoe UI", 10, "bold"), command=self.on_end_click)\
            .pack(side="left", expand=True, fill="x", padx=(3, 0))

        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _grid_buttons(self, parent, start_row, items, cols=3, width=15, command_factory=None):
        for i, word in enumerate(items):
            btn = tk.Button(
                parent, text=word, width=width, font=("Segoe UI", 9),
                command=command_factory(word),
            )
            btn.grid(row=start_row + i // cols, column=i % cols, padx=4, pady=3)
        return start_row + (len(items) + cols - 1) // cols

    # -- Strata tab ----------------------------------------------------------
    def _build_strata_tab(self):
        scroll = ScrollableFrame(self.strata_frame)
        scroll.pack(fill="both", expand=True)
        row = 0

        ttk.Label(scroll.inner, text="Most common", font=("Segoe UI", 10, "bold"))\
            .grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=(8, 2))
        row += 1
        row = self._grid_buttons(
            scroll.inner, row, STRATA_GROUPS["Most common"], cols=3, width=15,
            command_factory=lambda w: (lambda: self.on_strata_click(w)),
        )

        self.other_toggle_btn = tk.Button(
            scroll.inner, text="Other \u25be (show)", font=("Segoe UI", 9),
            command=self._toggle_other,
        )
        self.other_toggle_btn.grid(row=row, column=0, columnspan=3, sticky="we", padx=6, pady=(10, 4))
        row += 1

        self.other_frame = ttk.Frame(scroll.inner)
        self.other_frame.grid(row=row, column=0, columnspan=3, sticky="w")
        self._grid_buttons(
            self.other_frame, 0, STRATA_GROUPS["Other / less common"], cols=3, width=15,
            command_factory=lambda w: (lambda: self.on_strata_click(w)),
        )
        self.other_frame.grid_remove()

    def _toggle_other(self):
        self.other_visible = not self.other_visible
        if self.other_visible:
            self.other_frame.grid()
            self.other_toggle_btn.config(text="Other \u25b4 (hide)")
        else:
            self.other_frame.grid_remove()
            self.other_toggle_btn.config(text="Other \u25be (show)")

    # -- Description tab -----------------------------------------------------
    def _build_description_tab(self):
        scroll = ScrollableFrame(self.desc_frame)
        scroll.pack(fill="both", expand=True)
        self.desc_inner = scroll.inner

        tk.Button(
            scroll.inner, text="Same As Above But", font=("Segoe UI", 9, "bold"),
            command=lambda: self.on_word_click("Same As Above But"),
        ).grid(row=0, column=0, columnspan=4, sticky="we", padx=6, pady=(6, 8))

        self.step_label = ttk.Label(scroll.inner, text="", font=("Segoe UI", 10, "bold"))
        self.step_label.grid(row=1, column=0, columnspan=4, sticky="w", padx=6, pady=(4, 4))

        self.step_content_frame = ttk.Frame(scroll.inner)
        self.step_content_frame.grid(row=2, column=0, columnspan=4, sticky="w")

        self._render_step()

    def _step_sections(self, step_key):
        if step_key == "colors":
            sections = [("Colors", UNIVERSAL_GROUPS["Colors"])]
        elif step_key == "texture":
            mode_title, mode_words = MODE_GROUPS[self.current_mode]
            sections = [("Texture", UNIVERSAL_GROUPS["Texture"])]
            if mode_words:
                sections.append((mode_title, mode_words))
            sections.append(("Structure", STRUCTURE_WORDS))
        elif step_key == "grain":
            sections = [("Grain size", UNIVERSAL_GROUPS["Grain size"])]
        elif step_key == "connectors":
            # Connectors already live here as part of the step itself, so
            # no extra copy is appended below.
            return [
                ("Connectors", UNIVERSAL_GROUPS["Connectors"]),
                ("Inclusions & Materials", UNIVERSAL_GROUPS["Inclusions & Materials"] + QUICK_PHRASE_COMBOS),
            ]
        else:
            sections = []

        # Connectors are available on every step (not just the last one) so
        # you can drop in "With", "Some", "&", etc. without jumping ahead.
        sections.append(("Connectors", UNIVERSAL_GROUPS["Connectors"]))
        return sections

    def _render_step(self):
        for widget in self.step_content_frame.winfo_children():
            widget.destroy()

        step_key = STEP_KEYS[self.step_index]
        self.step_label.config(text=STEP_TITLES[step_key])

        row = 0
        for subheading, words in self._step_sections(step_key):
            ttk.Label(self.step_content_frame, text=subheading, font=("Segoe UI", 9, "italic"))\
                .grid(row=row, column=0, columnspan=4, sticky="w", padx=6, pady=(6, 2))
            row += 1
            row = self._grid_buttons(
                self.step_content_frame, row, words, cols=4, width=15,
                command_factory=lambda w: (lambda: self.on_word_click(w)),
            )

    def _on_tab_changed(self, event=None):
        active_tab = self.notebook.tab(self.notebook.select(), "text")
        if active_tab == "Description":
            self.desc_extra_frame.pack(side="right", padx=(6, 0))
        else:
            self.desc_extra_frame.pack_forget()

    # -- Strata click --------------------------------------------------------
    def on_strata_click(self, stratum):
        self.current_mode = get_mode(stratum)
        self.pending_starter = get_starter(stratum)
        self.on_word_click(stratum)

    # -- Word click ----------------------------------------------------------
    def on_word_click(self, word):
        # Strater keeps focus (WS_EX_NOACTIVATE), so we type directly.
        threading.Thread(target=self._type_word, args=(word,), daemon=True).start()

    def _type_word(self, word):
        # Small delay so the button click event fully resolves before typing.
        time.sleep(0.05)
        pyautogui.typewrite(word + " ", interval=0.01)

    # -- Next click ----------------------------------------------------------
    def on_next_click(self):
        active_tab = self.notebook.tab(self.notebook.select(), "text")
        if active_tab == "Strata Type":
            threading.Thread(target=self._next_from_strata, daemon=True).start()
        else:
            threading.Thread(target=self._next_from_description, daemon=True).start()

    def _next_from_strata(self):
        starter = self.pending_starter
        time.sleep(0.05)
        pyautogui.press("tab")
        if starter:
            time.sleep(0.03)
            pyautogui.typewrite(starter + " ", interval=0.01)
        self.step_index = 0
        self.after(0, self._switch_to_description)

    def _next_from_description(self):
        time.sleep(0.05)
        pyautogui.typewrite(", ", interval=0.01)
        if self.step_index < len(STEP_KEYS) - 1:
            self.step_index += 1
        self.after(0, self._render_step)

    # -- Comma / Null mini-buttons -------------------------------------------
    def on_comma_click(self):
        threading.Thread(target=self._type_comma_only, daemon=True).start()

    def _type_comma_only(self):
        time.sleep(0.05)
        pyautogui.typewrite(", ", interval=0.01)

    def on_null_click(self):
        if self.step_index < len(STEP_KEYS) - 1:
            self.step_index += 1
        self._render_step()

    def _switch_to_description(self):
        self._render_step()
        self.notebook.select(self.desc_frame)
        self.status.config(text="Click description words, Next moves to the next step, then End.")

    # -- End click -----------------------------------------------------------
    def on_end_click(self):
        threading.Thread(target=self._type_end_sequence, daemon=True).start()

    def _type_end_sequence(self):
        time.sleep(0.05)
        pyautogui.press("backspace")
        pyautogui.typewrite(".", interval=0.01)
        pyautogui.press("enter")
        pyautogui.press("left")
        self.step_index = 0
        self.after(0, self._switch_to_strata)

    def _switch_to_strata(self):
        self.notebook.select(self.strata_frame)
        self.status.config(text="Click into the Strata column, then pick a stratum.")


if __name__ == "__main__":
    app = EntryHelper()
    app.mainloop()