"""
Shelf — File Management Toolkit
Requires: pip install py7zr rarfile send2trash
"""

import os, re, shutil, hashlib, zipfile, tarfile, threading
from datetime import datetime
from pathlib import Path
from collections import defaultdict

# ── Optional libraries ────────────────────────────────────────────────────────
try:
    import py7zr;    HAS_7Z    = True
except ImportError:  HAS_7Z    = False
try:
    import rarfile;  HAS_RAR   = True
except ImportError:  HAS_RAR   = False
try:
    import send2trash; HAS_TRASH = True
except ImportError:    HAS_TRASH = False

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ══════════════════════════════════════════════════════════════════════════════
#  THEME
# ══════════════════════════════════════════════════════════════════════════════
C = dict(
    bg       = "#0D0F18",
    sidebar  = "#111420",
    card     = "#161925",
    border   = "#252836",
    accent   = "#6C63FF",
    accent2  = "#FF6584",
    green    = "#43D98C",
    amber    = "#F5A623",
    red      = "#FF5C5C",
    fg       = "#E2E4F0",
    fg2      = "#9395A5",
    fg3      = "#555868",
    logbg    = "#0A0C14",
    selbg    = "#1E2035",
)

FONT_TITLE  = ("Trebuchet MS", 13, "bold")
FONT_LABEL  = ("Trebuchet MS", 8, "bold")
FONT_BODY   = ("Trebuchet MS", 9)
FONT_MONO   = ("Consolas", 9)
FONT_LOGO   = ("Trebuchet MS", 17, "bold")
FONT_NAV    = ("Trebuchet MS", 10)
FONT_BTN    = ("Trebuchet MS", 9, "bold")

# ══════════════════════════════════════════════════════════════════════════════
#  CORE LOGIC
# ══════════════════════════════════════════════════════════════════════════════
FILE_TYPE_MAP = {
    "Images"     : [".jpg",".jpeg",".png",".gif",".bmp",".svg",".webp",".tiff",".ico",".heic"],
    "Videos"     : [".mp4",".mkv",".avi",".mov",".wmv",".flv",".webm",".m4v",".mpg"],
    "Audio"      : [".mp3",".wav",".flac",".aac",".ogg",".wma",".m4a",".opus"],
    "Documents"  : [".pdf",".doc",".docx",".xls",".xlsx",".ppt",".pptx",".txt",".rtf",".odt",".csv",".md"],
    "Archives"   : [".zip",".rar",".7z",".tar",".gz",".bz2",".xz",".tgz"],
    "Code"       : [".py",".js",".ts",".html",".css",".java",".cpp",".c",".h",
                    ".json",".xml",".yaml",".yml",".sh",".bat",".go",".rs",".rb"],
    "Executables": [".exe",".msi",".dmg",".deb",".rpm",".apk"],
}

ARCHIVE_EXTS = {".zip",".tar",".gz",".bz2",".xz",".tgz",".tar.gz",".tar.bz2",".tar.xz"}
if HAS_7Z:  ARCHIVE_EXTS.add(".7z")
if HAS_RAR: ARCHIVE_EXTS.add(".rar")

# ── helpers ───────────────────────────────────────────────────────────────────
def _safe_name(folder: Path, name: str) -> Path:
    """Return a non-colliding path by appending _1, _2 … if needed."""
    dest = folder / name
    if not dest.exists():
        return dest
    stem, suffix = Path(name).stem, Path(name).suffix
    i = 1
    while True:
        dest = folder / f"{stem}_{i}{suffix}"
        if not dest.exists():
            return dest
        i += 1

def _extract_archive(arc: Path, dest: Path, log):
    n = arc.name.lower()
    try:
        if n.endswith(".zip"):
            with zipfile.ZipFile(arc) as z: z.extractall(dest)
        elif HAS_7Z and n.endswith(".7z"):
            with py7zr.SevenZipFile(arc) as z: z.extractall(dest)
        elif HAS_RAR and n.endswith(".rar"):
            with rarfile.RarFile(arc) as z: z.extractall(dest)
        elif any(n.endswith(e) for e in (".tar",".tar.gz",".tgz",".tar.bz2",".tar.xz",".gz",".bz2",".xz")):
            with tarfile.open(arc) as z: z.extractall(dest)
        else:
            log(f"⚠  Unsupported format: {arc.name}"); return False
    except Exception as e:
        log(f"❌  Failed to extract {arc.name}: {e}"); return False
    return True

# ── tools ─────────────────────────────────────────────────────────────────────
def run_smart_extract(folder, rename_map, log, done_cb):
    folder = Path(folder)
    if not folder.exists(): log("❌  Folder not found."); done_cb(); return
    archives = [f for f in folder.iterdir()
                if f.is_file() and any(f.name.lower().endswith(e) for e in ARCHIVE_EXTS)]
    if not archives: log("⚠  No supported archives found."); done_cb(); return

    for arc in sorted(archives):
        stem = arc.stem
        if stem.endswith(".tar"): stem = stem[:-4]
        final_name = rename_map.get(stem, stem).strip() or stem
        # Sanitise name — strip illegal Windows chars
        final_name = re.sub(r'[\\/:*?"<>|]', "_", final_name)

        tmp = folder / f"__shelf_tmp_{stem}"
        if tmp.exists(): shutil.rmtree(tmp)
        tmp.mkdir()

        log(f"📦  Extracting  {arc.name} …")
        if not _extract_archive(arc, tmp, log):
            shutil.rmtree(tmp, ignore_errors=True); continue

        contents = list(tmp.iterdir())
        # promote single inner folder (common zip pattern)
        if len(contents) == 1 and contents[0].is_dir():
            inner = contents[0]
            dest = _safe_name(folder, final_name)
            if dest != folder / final_name:
                log(f"⚠  Name collision — saving as  {dest.name}")
            shutil.move(str(inner), str(dest))
            shutil.rmtree(tmp, ignore_errors=True)
        else:
            dest = _safe_name(folder, final_name)
            if dest != folder / final_name:
                log(f"⚠  Name collision — saving as  {dest.name}")
            tmp.rename(dest)

        log(f"✅  {arc.name}  →  {dest.name}/")
        try: arc.unlink(); log(f"🗑  Removed  {arc.name}")
        except Exception as e: log(f"⚠  Could not delete {arc.name}: {e}")

    log("\n🎉  All done!"); done_cb()


def run_bulk_rename(folder, find, replace, regex, preview_only, log, done_cb):
    folder = Path(folder)
    if not folder.exists(): log("❌  Folder not found."); done_cb(); return
    if not find: log("⚠  Find field is empty."); done_cb(); return
    renamed = 0
    for item in sorted(folder.iterdir()):
        if not item.is_dir(): continue
        old = item.name
        try:
            new = re.sub(find, replace, old) if regex else old.replace(find, replace)
        except re.error as e:
            log(f"❌  Regex error: {e}"); done_cb(); return
        if new == old: continue
        new = re.sub(r'[\\/:*?"<>|]', "_", new).strip()  # sanitise
        if not new: log(f"⚠  Skipped '{old}' — result would be empty"); continue
        dest = folder / new
        if preview_only:
            log(f"👁  {old}  →  {new}"); renamed += 1; continue
        if dest.exists():
            log(f"⚠  Skipped '{old}' — '{new}' already exists")
        else:
            item.rename(dest); log(f"✅  {old}  →  {new}"); renamed += 1
    action = "previewed" if preview_only else "renamed"
    log(f"\n✔  {renamed} folder(s) {action}."); done_cb()


def run_find_duplicates(folder, log, done_cb):
    folder = Path(folder)
    if not folder.exists(): log("❌  Folder not found."); done_cb(); return
    hashes = defaultdict(list)
    log("🔍  Hashing files …")
    count = 0
    for f in folder.rglob("*"):
        if f.is_file():
            try:
                # Use first 64 KB for speed on large files, then full hash
                head = f.read_bytes()[:65536]
                h = hashlib.md5(head + str(f.stat().st_size).encode()).hexdigest()
                hashes[h].append(f); count += 1
            except (PermissionError, OSError) as e:
                log(f"⚠  Skipped {f.name}: {e}")
    log(f"   Scanned {count} file(s).")
    dupes = {h: ps for h, ps in hashes.items() if len(ps) > 1}
    if not dupes: log("✅  No duplicates found!"); done_cb(); return
    total = 0
    for h, paths in dupes.items():
        log(f"\n🔁  Group [{h[:10]}…]  ({len(paths)} copies)")
        for p in paths: log(f"    {p.relative_to(folder)}")
        for p in paths[1:]:
            try:
                if HAS_TRASH:
                    send2trash.send2trash(str(p)); log(f"  🗑  Trashed: {p.name}")
                else:
                    p.unlink(); log(f"  🗑  Deleted: {p.name}")
                total += 1
            except Exception as e: log(f"  ❌  Could not remove {p.name}: {e}")
    log(f"\n✔  Removed {total} duplicate(s)."); done_cb()


def run_flatten(folder, log, done_cb):
    folder = Path(folder)
    if not folder.exists(): log("❌  Folder not found."); done_cb(); return
    moved = 0
    for f in list(folder.rglob("*")):
        if f.is_file() and f.parent != folder:
            dest = _safe_name(folder, f.name)
            try:
                shutil.move(str(f), str(dest))
                log(f"✅  {f.relative_to(folder)}  →  {dest.name}")
                moved += 1
            except Exception as e: log(f"⚠  Skipped {f.name}: {e}")
    # remove empty dirs (deepest first)
    for d in sorted(folder.rglob("*"), key=lambda p: len(p.parts), reverse=True):
        if d.is_dir():
            try: d.rmdir()
            except OSError: pass
    log(f"\n✔  {moved} file(s) moved to root."); done_cb()


def run_sort_by_type(folder, log, done_cb):
    folder = Path(folder)
    if not folder.exists(): log("❌  Folder not found."); done_cb(); return
    moved = 0
    for f in list(folder.iterdir()):
        if not f.is_file(): continue
        ext = f.suffix.lower()
        cat = next((c for c, exts in FILE_TYPE_MAP.items() if ext in exts), "Other")
        dest_dir = folder / cat; dest_dir.mkdir(exist_ok=True)
        dest = _safe_name(dest_dir, f.name)
        try:
            shutil.move(str(f), str(dest))
            log(f"✅  {f.name}  →  {cat}/"); moved += 1
        except Exception as e: log(f"⚠  Skipped {f.name}: {e}")
    log(f"\n✔  {moved} file(s) sorted."); done_cb()


def run_size_report(folder, log, done_cb):
    folder = Path(folder)
    if not folder.exists(): log("❌  Folder not found."); done_cb(); return
    items = []
    for item in folder.iterdir():
        try:
            size = sum(f.stat().st_size for f in item.rglob("*") if f.is_file()) \
                   if item.is_dir() else item.stat().st_size
            items.append((size, item.name, "📁" if item.is_dir() else "📄"))
        except Exception: pass
    items.sort(reverse=True)
    total = sum(s for s,_,_ in items)
    log(f"{'SIZE':>12}   {'%':>5}   ITEM")
    log("─" * 55)
    for size, name, icon in items:
        pct = f"{size/total*100:.1f}%" if total else "—"
        if   size >= 1_073_741_824: s = f"{size/1_073_741_824:.2f} GB"
        elif size >= 1_048_576:     s = f"{size/1_048_576:.2f} MB"
        elif size >= 1024:          s = f"{size/1024:.2f} KB"
        else:                       s = f"{size} B"
        log(f"{s:>12}   {pct:>5}   {icon} {name}")
    log("─" * 55)
    if   total >= 1_073_741_824: ts = f"{total/1_073_741_824:.2f} GB"
    elif total >= 1_048_576:     ts = f"{total/1_048_576:.2f} MB"
    else:                        ts = f"{total/1024:.2f} KB"
    log(f"{'TOTAL':>12}          {ts}")
    done_cb()


def run_merge_folders(src_list, dest, log, done_cb):
    dest = Path(dest)
    if not dest.exists(): dest.mkdir(parents=True)
    moved = 0
    for src in src_list:
        src = Path(src)
        if not src.is_dir(): log(f"⚠  Not a folder: {src}"); continue
        if src.resolve() == dest.resolve(): log(f"⚠  Source = destination, skipping"); continue
        for f in src.rglob("*"):
            if f.is_file():
                rel = f.relative_to(src)
                target = dest / rel; target.parent.mkdir(parents=True, exist_ok=True)
                target = _safe_name(target.parent, target.name)
                try:
                    shutil.copy2(str(f), str(target))
                    log(f"✅  {src.name}/{rel}"); moved += 1
                except Exception as e: log(f"⚠  {f.name}: {e}")
    log(f"\n✔  {moved} file(s) merged."); done_cb()


def run_organize_by_date(folder, log, done_cb):
    folder = Path(folder)
    if not folder.exists(): log("❌  Folder not found."); done_cb(); return
    moved = 0
    for f in list(folder.iterdir()):
        if not f.is_file(): continue
        try: mtime = datetime.fromtimestamp(f.stat().st_mtime)
        except Exception: log(f"⚠  Could not read date for {f.name}"); continue
        dir_name = mtime.strftime("%Y-%m")
        dest_dir = folder / dir_name; dest_dir.mkdir(exist_ok=True)
        dest = _safe_name(dest_dir, f.name)
        try:
            shutil.move(str(f), str(dest))
            log(f"✅  {f.name}  →  {dir_name}/"); moved += 1
        except Exception as e: log(f"⚠  Skipped {f.name}: {e}")
    log(f"\n✔  {moved} file(s) organised."); done_cb()


# ══════════════════════════════════════════════════════════════════════════════
#  UI HELPERS
# ══════════════════════════════════════════════════════════════════════════════
TOOLS = [
    ("📦", "Smart Extract",       "Extract ZIP · RAR · 7Z · TAR with optional batch rename"),
    ("✏️", "Bulk Rename",          "Find & replace text in folder names"),
    ("🔁", "Find Duplicates",     "Scan & remove duplicate files"),
    ("📂", "Flatten Folders",     "Pull all nested files to root"),
    ("🗂", "Sort by Type",        "Organise files into category folders"),
    ("📊", "Size Report",         "Visualise folder & file sizes"),
    ("🔀", "Merge Folders",       "Combine multiple folders into one"),
    ("📅", "Organise by Date",    "Move files into YYYY-MM folders"),
]


class _StyledEntry(tk.Entry):
    def __init__(self, master, var, **kw):
        super().__init__(master, textvariable=var,
                         font=FONT_MONO,
                         bg=C["card"], fg=C["fg"],
                         insertbackground=C["fg"],
                         relief="flat", bd=0,
                         highlightthickness=1,
                         highlightbackground=C["border"],
                         highlightcolor=C["accent"], **kw)


class _Card(tk.Frame):
    def __init__(self, master, **kw):
        super().__init__(master, bg=C["card"],
                         highlightthickness=1,
                         highlightbackground=C["border"], **kw)


class ShelfApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Shelf")
        self.geometry("1100x700")
        self.minsize(960, 620)
        self.configure(bg=C["bg"])
        self._active = -1
        self._setup_treeview_style()
        self._build()
        self._switch(0)

    # ── treeview style ────────────────────────────────────────────────────
    def _setup_treeview_style(self):
        s = ttk.Style()
        s.theme_use("default")
        s.configure("Shelf.Treeview",
                    background=C["card"], foreground=C["fg"],
                    fieldbackground=C["card"], rowheight=28,
                    font=FONT_MONO, borderwidth=0)
        s.configure("Shelf.Treeview.Heading",
                    background=C["sidebar"], foreground=C["fg2"],
                    font=FONT_LABEL, relief="flat")
        s.map("Shelf.Treeview",
              background=[("selected", C["selbg"])],
              foreground=[("selected", C["fg"])])

    # ── master layout ─────────────────────────────────────────────────────
    def _build(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── sidebar ──
        self._sb = tk.Frame(self, bg=C["sidebar"], width=220)
        self._sb.grid(row=0, column=0, sticky="nsew")
        self._sb.grid_propagate(False)
        self._sb.grid_rowconfigure(99, weight=1)

        # logo
        logo_frame = tk.Frame(self._sb, bg=C["sidebar"])
        logo_frame.grid(row=0, column=0, sticky="ew", padx=18, pady=(22, 6))
        tk.Label(logo_frame, text="⬡", font=("Trebuchet MS", 22),
                 bg=C["sidebar"], fg=C["accent"]).pack(side="left")
        tk.Label(logo_frame, text=" Shelf", font=FONT_LOGO,
                 bg=C["sidebar"], fg=C["fg"]).pack(side="left")

        tk.Frame(self._sb, bg=C["border"], height=1).grid(
            row=1, column=0, sticky="ew", padx=14, pady=(0, 8))

        # nav buttons
        self._nav = []
        for i, (icon, label, _) in enumerate(TOOLS):
            f = tk.Frame(self._sb, bg=C["sidebar"], cursor="hand2")
            f.grid(row=i+2, column=0, sticky="ew", padx=8, pady=2)
            f.grid_columnconfigure(1, weight=1)

            pill = tk.Frame(f, bg=C["sidebar"], width=3)
            pill.grid(row=0, column=0, sticky="ns", padx=(0, 0))

            icon_lbl = tk.Label(f, text=icon, font=("Trebuchet MS", 12),
                                bg=C["sidebar"], fg=C["fg2"], width=2)
            icon_lbl.grid(row=0, column=1, padx=(8, 4), pady=8)

            txt_lbl = tk.Label(f, text=label, font=FONT_NAV,
                               bg=C["sidebar"], fg=C["fg2"], anchor="w")
            txt_lbl.grid(row=0, column=2, sticky="ew")

            for w in (f, icon_lbl, txt_lbl, pill):
                w.bind("<Button-1>", lambda e, idx=i: self._switch(idx))
                w.bind("<Enter>",    lambda e, idx=i: self._nav_hover(idx, True))
                w.bind("<Leave>",    lambda e, idx=i: self._nav_hover(idx, False))

            self._nav.append((f, pill, icon_lbl, txt_lbl))

        # version label at bottom
        tk.Label(self._sb, text="Shelf  v1.0", font=("Trebuchet MS", 8),
                 bg=C["sidebar"], fg=C["fg3"]).grid(
            row=99, column=0, pady=12)

        # ── right panel ──
        right = tk.Frame(self, bg=C["bg"])
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=1)

        # topbar
        self._topbar = tk.Frame(right, bg=C["card"], height=58)
        self._topbar.grid(row=0, column=0, sticky="ew")
        self._topbar.grid_propagate(False)
        self._topbar.grid_columnconfigure(0, weight=1)

        self._top_icon  = tk.Label(self._topbar, text="", font=("Trebuchet MS", 16),
                                   bg=C["card"], fg=C["accent"])
        self._top_icon.place(x=22, rely=0.5, anchor="w")

        self._top_title = tk.Label(self._topbar, text="", font=FONT_TITLE,
                                   bg=C["card"], fg=C["fg"])
        self._top_title.place(x=52, rely=0.35, anchor="w")

        self._top_desc  = tk.Label(self._topbar, text="", font=FONT_BODY,
                                   bg=C["card"], fg=C["fg2"])
        self._top_desc.place(x=53, rely=0.72, anchor="w")

        # content area
        self._content = tk.Frame(right, bg=C["bg"])
        self._content.grid(row=1, column=0, sticky="nsew", padx=24, pady=20)
        self._content.grid_columnconfigure(0, weight=1)
        self._content.grid_rowconfigure(99, weight=1)

    # ── nav state ────────────────────────────────────────────────────────
    def _nav_hover(self, idx, on):
        if idx == self._active: return
        f, pill, il, tl = self._nav[idx]
        col = C["fg"] if on else C["fg2"]
        bg  = C["selbg"] if on else C["sidebar"]
        for w in (f, il, tl, pill): w.configure(bg=bg)
        il.configure(fg=col); tl.configure(fg=col)

    def _switch(self, idx):
        if self._active >= 0:
            f, pill, il, tl = self._nav[self._active]
            for w in (f, il, tl, pill): w.configure(bg=C["sidebar"])
            il.configure(fg=C["fg2"]); tl.configure(fg=C["fg2"])
            pill.configure(bg=C["sidebar"])

        self._active = idx
        f, pill, il, tl = self._nav[idx]
        for w in (f, il, tl): w.configure(bg=C["selbg"])
        il.configure(fg=C["accent"], bg=C["selbg"])
        tl.configure(fg=C["fg"],     bg=C["selbg"], font=("Trebuchet MS", 10, "bold"))
        pill.configure(bg=C["accent"])

        icon, label, desc = TOOLS[idx]
        self._top_icon.configure(text=icon)
        self._top_title.configure(text=label)
        self._top_desc.configure(text=desc)

        for w in self._content.winfo_children(): w.destroy()
        [self._pg_extract, self._pg_rename, self._pg_dupes,
         self._pg_flatten, self._pg_sort,   self._pg_size,
         self._pg_merge,   self._pg_date][idx]()

    # ── shared widgets ────────────────────────────────────────────────────
    def _section_label(self, parent, text, row, col=0, padtop=14):
        tk.Label(parent, text=text, font=FONT_LABEL,
                 bg=C["bg"], fg=C["fg2"]).grid(
            row=row, column=col, sticky="w", pady=(padtop, 3))

    def _path_row(self, parent, label, var, row, multi=False):
        self._section_label(parent, label, row)
        f = tk.Frame(parent, bg=C["bg"])
        f.grid(row=row+1, column=0, columnspan=2, sticky="ew")
        f.grid_columnconfigure(0, weight=1)
        e = _StyledEntry(f, var)
        e.grid(row=0, column=0, sticky="ew", ipady=8, padx=(0, 8))

        def browse():
            if multi:
                paths = []
                while True:
                    p = filedialog.askdirectory(title="Add folder (Cancel when done)")
                    if not p: break
                    paths.append(p)
                if paths:
                    cur = var.get()
                    var.set((cur + ";" if cur else "") + ";".join(paths))
            else:
                p = filedialog.askdirectory()
                if p: var.set(p)

        self._btn(f, "Browse", browse, row=0, col=1, accent=False)
        return row + 2

    def _btn(self, parent, text, cmd, row=0, col=0, accent=True,
             sticky="w", pady=0, padx=0, use_pack=False):
        bg = C["accent"] if accent else C["border"]
        b = tk.Label(parent, text=text, font=FONT_BTN,
                     bg=bg, fg=C["fg"], padx=14, pady=8,
                     cursor="hand2")
        if use_pack:
            px = padx if isinstance(padx, tuple) else (padx, 0)
            py = pady if isinstance(pady, tuple) else (pady, 0)
            b.pack(side="left", padx=px, pady=py)
        else:
            b.grid(row=row, column=col, sticky=sticky,
                   pady=(pady, 0) if isinstance(pady, int) else pady,
                   padx=(padx, 0) if isinstance(padx, int) else padx)
        b.bind("<Button-1>", lambda e: cmd())
        b.bind("<Enter>", lambda e: b.configure(bg=C["accent2"] if accent else C["selbg"]))
        b.bind("<Leave>", lambda e: b.configure(bg=bg))
        return b

    def _log_panel(self, parent, row, col=0, colspan=2):
        self._section_label(parent, "OUTPUT LOG", row)
        outer = tk.Frame(parent, bg=C["border"], padx=1, pady=1)
        outer.grid(row=row+1, column=col, columnspan=colspan, sticky="nsew")
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(0, weight=1)
        inner = tk.Frame(outer, bg=C["logbg"])
        inner.grid(sticky="nsew"); inner.grid_columnconfigure(0, weight=1); inner.grid_rowconfigure(0, weight=1)
        txt = tk.Text(inner, font=FONT_MONO, bg=C["logbg"], fg=C["green"],
                      relief="flat", bd=0, wrap="word",
                      insertbackground=C["fg"], padx=12, pady=10,
                      state="disabled")
        sb = tk.Scrollbar(inner, command=txt.yview,
                          bg=C["card"], troughcolor=C["logbg"],
                          activebackground=C["accent"], relief="flat")
        txt.configure(yscrollcommand=sb.set)
        txt.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        return txt

    def _log(self, w, msg):
        w.configure(state="normal")
        w.insert("end", msg + "\n")
        w.see("end")
        w.configure(state="disabled")
        self.update_idletasks()

    def _clear(self, w):
        w.configure(state="normal"); w.delete("1.0","end"); w.configure(state="disabled")

    def _field(self, parent, label, var, row, col=0, width=None, padtop=14):
        self._section_label(parent, label, row, col=col, padtop=padtop)
        e = _StyledEntry(parent, var, width=width or 0)
        e.grid(row=row+1, column=col, sticky="ew" if not width else "", ipady=8,
               padx=(0, 12 if col == 0 else 0))
        return row + 2

    # ══════════════════════════════════════════════════════════════════════
    #  PAGE: SMART EXTRACT
    # ══════════════════════════════════════════════════════════════════════
    def _pg_extract(self):
        c = self._content
        c.grid_rowconfigure(7, weight=1)
        self._ex_path = tk.StringVar()
        nrow = self._path_row(c, "FOLDER CONTAINING ARCHIVES", self._ex_path, 0)

        # info pills for supported formats
        pill_frame = tk.Frame(c, bg=C["bg"])
        pill_frame.grid(row=nrow, column=0, sticky="w", pady=(4, 0))
        supported = ["ZIP", "TAR", "GZ", "BZ2", "XZ"]
        if HAS_7Z:  supported.append("7Z")
        if HAS_RAR: supported.append("RAR")
        for fmt in supported:
            tk.Label(pill_frame, text=f" {fmt} ", font=("Trebuchet MS", 8, "bold"),
                     bg=C["selbg"], fg=C["accent"], padx=4, pady=2
                     ).pack(side="left", padx=(0, 4))
        nrow += 1

        # rename table header
        self._section_label(c, "BATCH RENAME  —  double-click New Name to edit", nrow)
        nrow += 1

        tbl_outer = tk.Frame(c, bg=C["border"], padx=1, pady=1)
        tbl_outer.grid(row=nrow, column=0, sticky="ew")
        tbl_outer.grid_columnconfigure(0, weight=1)
        cols = ("Archive File", "New Name  (blank = keep original)")
        self._ex_tree = ttk.Treeview(tbl_outer, columns=cols,
                                     show="headings", height=5,
                                     style="Shelf.Treeview")
        self._ex_tree.heading(cols[0], text=cols[0])
        self._ex_tree.heading(cols[1], text=cols[1])
        self._ex_tree.column(cols[0], width=260, stretch=False)
        self._ex_tree.column(cols[1], width=300, stretch=True)
        self._ex_tree.pack(fill="x")
        self._ex_tree.bind("<Double-1>", self._ex_edit_cell)
        nrow += 1

        btn_row = tk.Frame(c, bg=C["bg"])
        btn_row.grid(row=nrow, column=0, sticky="w", pady=(8, 0))
        self._btn(btn_row, "🔄  Scan Folder", self._ex_scan,
                  row=0, col=0, accent=False)
        self._btn(btn_row, "▶  Extract All", self._ex_run,
                  row=0, col=1, accent=True, padx=10)
        nrow += 1

        self._ex_log = self._log_panel(c, nrow)

    def _ex_scan(self):
        folder = self._ex_path.get().strip()
        if not folder or not Path(folder).exists():
            messagebox.showerror("Shelf", "Select a valid folder first."); return
        for r in self._ex_tree.get_children(): self._ex_tree.delete(r)
        found = 0
        for f in sorted(Path(folder).iterdir()):
            if f.is_file() and any(f.name.lower().endswith(e) for e in ARCHIVE_EXTS):
                self._ex_tree.insert("", "end", values=(f.name, "")); found += 1
        self._clear(self._ex_log)
        self._log(self._ex_log,
                  f"Found {found} archive(s).  Double-click the right column to rename before extracting.")

    def _ex_edit_cell(self, event):
        item = self._ex_tree.identify_row(event.y)
        col  = self._ex_tree.identify_column(event.x)
        if not item or col != "#2": return
        x, y, w, h = self._ex_tree.bbox(item, col)
        val = self._ex_tree.item(item, "values")[1]
        v = tk.StringVar(value=val)
        e = tk.Entry(self._ex_tree, textvariable=v, font=FONT_MONO,
                     bg=C["selbg"], fg=C["fg"], insertbackground=C["fg"],
                     relief="flat", bd=0, highlightthickness=1,
                     highlightbackground=C["accent"], highlightcolor=C["accent"])
        e.place(x=x, y=y, width=w, height=h); e.focus(); e.select_range(0, "end")
        def commit(_=None):
            vals = list(self._ex_tree.item(item, "values"))
            vals[1] = v.get(); self._ex_tree.item(item, values=vals); e.destroy()
        e.bind("<Return>", commit); e.bind("<Escape>", lambda _: e.destroy())
        e.bind("<FocusOut>", commit)

    def _ex_run(self):
        folder = self._ex_path.get().strip()
        if not folder: messagebox.showerror("Shelf", "Select a folder."); return
        rename_map = {}
        for row in self._ex_tree.get_children():
            arc, new = self._ex_tree.item(row, "values")
            stem = Path(arc).stem
            if stem.endswith(".tar"): stem = stem[:-4]
            if new.strip(): rename_map[stem] = new.strip()
        self._clear(self._ex_log)
        threading.Thread(target=run_smart_extract,
                         args=(folder, rename_map,
                               lambda m: self._log(self._ex_log, m),
                               lambda: None), daemon=True).start()

    # ══════════════════════════════════════════════════════════════════════
    #  PAGE: BULK RENAME
    # ══════════════════════════════════════════════════════════════════════
    def _pg_rename(self):
        c = self._content
        c.grid_columnconfigure(0, weight=1)
        c.grid_rowconfigure(8, weight=1)
        self._rn_path = tk.StringVar()
        nrow = self._path_row(c, "TARGET FOLDER", self._rn_path, 0)

        # find / replace side by side
        fr = tk.Frame(c, bg=C["bg"])
        fr.grid(row=nrow, column=0, sticky="ew", pady=(4, 0))
        fr.grid_columnconfigure(0, weight=1)
        fr.grid_columnconfigure(1, weight=1)
        self._rn_find    = tk.StringVar(value="-2023")
        self._rn_replace = tk.StringVar(value="")
        self._rn_regex   = tk.BooleanVar(value=False)
        self._field(fr, "FIND", self._rn_find, 0, col=0)
        self._field(fr, "REPLACE WITH  (blank = remove)", self._rn_replace, 0, col=1)
        nrow += 1

        # checkbox + buttons on same row — all use pack
        ctrl = tk.Frame(c, bg=C["bg"])
        ctrl.grid(row=nrow, column=0, sticky="w", pady=(10, 0))
        cb = tk.Checkbutton(ctrl, text="  Use Regex", variable=self._rn_regex,
                            font=FONT_BODY, bg=C["bg"], fg=C["fg2"],
                            selectcolor=C["card"], activebackground=C["bg"],
                            activeforeground=C["fg"], cursor="hand2")
        cb.pack(side="left", padx=(0, 16))
        self._btn(ctrl, "👁  Preview", lambda: self._rn_run(preview=True),
                  accent=False, use_pack=True, padx=(0, 8))
        self._btn(ctrl, "▶  Rename", lambda: self._rn_run(preview=False),
                  accent=True, use_pack=True)
        nrow += 1

        self._rn_log = self._log_panel(c, nrow)

    def _rn_run(self, preview):
        folder = self._rn_path.get().strip()
        if not folder: messagebox.showerror("Shelf", "Select a folder."); return
        self._clear(self._rn_log)
        if preview:
            self._log(self._rn_log, "— PREVIEW MODE — no changes will be made\n")
        threading.Thread(target=run_bulk_rename,
                         args=(folder, self._rn_find.get(), self._rn_replace.get(),
                               self._rn_regex.get(), preview,
                               lambda m: self._log(self._rn_log, m),
                               lambda: None), daemon=True).start()

    # ══════════════════════════════════════════════════════════════════════
    #  PAGE: FIND DUPLICATES
    # ══════════════════════════════════════════════════════════════════════
    def _pg_dupes(self):
        c = self._content; c.grid_rowconfigure(4, weight=1)
        self._dup_path = tk.StringVar()
        nrow = self._path_row(c, "SCAN FOLDER  (scans recursively)", self._dup_path, 0)
        notice = tk.Label(c,
            text="ℹ  Keeps the first copy found. Extras are sent to the Recycle Bin (or deleted if unavailable).",
            font=FONT_BODY, bg=C["bg"], fg=C["amber"], wraplength=780, justify="left")
        notice.grid(row=nrow, column=0, sticky="w", pady=(6, 0))
        nrow += 1
        self._btn(c, "🔍  Scan & Remove Duplicates", self._dup_run,
                  row=nrow, sticky="w", pady=10)
        nrow += 1
        self._dup_log = self._log_panel(c, nrow)

    def _dup_run(self):
        folder = self._dup_path.get().strip()
        if not folder: messagebox.showerror("Shelf", "Select a folder."); return
        if not messagebox.askyesno("Shelf — Confirm",
                "This will permanently remove duplicate files (keeps one copy).\n\nContinue?"):
            return
        self._clear(self._dup_log)
        threading.Thread(target=run_find_duplicates,
                         args=(folder, lambda m: self._log(self._dup_log, m),
                               lambda: None), daemon=True).start()

    # ══════════════════════════════════════════════════════════════════════
    #  PAGE: FLATTEN
    # ══════════════════════════════════════════════════════════════════════
    def _pg_flatten(self):
        c = self._content; c.grid_rowconfigure(4, weight=1)
        self._fl_path = tk.StringVar()
        nrow = self._path_row(c, "ROOT FOLDER", self._fl_path, 0)
        tk.Label(c, text="ℹ  All files in sub-folders will be moved to the root. Empty folders are deleted.",
                 font=FONT_BODY, bg=C["bg"], fg=C["amber"], wraplength=780
                 ).grid(row=nrow, column=0, sticky="w", pady=(6, 0)); nrow += 1
        self._btn(c, "📂  Flatten", self._fl_run, row=nrow, sticky="w", pady=10)
        nrow += 1
        self._fl_log = self._log_panel(c, nrow)

    def _fl_run(self):
        folder = self._fl_path.get().strip()
        if not folder: messagebox.showerror("Shelf", "Select a folder."); return
        self._clear(self._fl_log)
        threading.Thread(target=run_flatten,
                         args=(folder, lambda m: self._log(self._fl_log, m),
                               lambda: None), daemon=True).start()

    # ══════════════════════════════════════════════════════════════════════
    #  PAGE: SORT BY TYPE
    # ══════════════════════════════════════════════════════════════════════
    def _pg_sort(self):
        c = self._content; c.grid_rowconfigure(5, weight=1)
        self._st_path = tk.StringVar()
        nrow = self._path_row(c, "FOLDER TO SORT", self._st_path, 0)

        # category chips
        chip_outer = tk.Frame(c, bg=C["bg"])
        chip_outer.grid(row=nrow, column=0, sticky="w", pady=(8, 0)); nrow += 1
        chips = [(cat, C["accent"]) for cat in FILE_TYPE_MAP] + [("Other", C["fg3"])]
        for i, (cat, col) in enumerate(chips):
            tk.Label(chip_outer, text=f" {cat} ", font=("Trebuchet MS", 8),
                     bg=C["selbg"], fg=col, padx=6, pady=3
                     ).grid(row=0, column=i, padx=(0, 4))

        self._btn(c, "🗂  Sort Files", self._st_run, row=nrow, sticky="w", pady=10)
        nrow += 1
        self._st_log = self._log_panel(c, nrow)

    def _st_run(self):
        folder = self._st_path.get().strip()
        if not folder: messagebox.showerror("Shelf", "Select a folder."); return
        self._clear(self._st_log)
        threading.Thread(target=run_sort_by_type,
                         args=(folder, lambda m: self._log(self._st_log, m),
                               lambda: None), daemon=True).start()

    # ══════════════════════════════════════════════════════════════════════
    #  PAGE: SIZE REPORT
    # ══════════════════════════════════════════════════════════════════════
    def _pg_size(self):
        c = self._content; c.grid_rowconfigure(4, weight=1)
        self._sz_path = tk.StringVar()
        nrow = self._path_row(c, "FOLDER TO ANALYSE", self._sz_path, 0)
        self._btn(c, "📊  Generate Report", self._sz_run, row=nrow, sticky="w", pady=10)
        nrow += 1
        self._sz_log = self._log_panel(c, nrow)

    def _sz_run(self):
        folder = self._sz_path.get().strip()
        if not folder: messagebox.showerror("Shelf", "Select a folder."); return
        self._clear(self._sz_log)
        threading.Thread(target=run_size_report,
                         args=(folder, lambda m: self._log(self._sz_log, m),
                               lambda: None), daemon=True).start()

    # ══════════════════════════════════════════════════════════════════════
    #  PAGE: MERGE FOLDERS
    # ══════════════════════════════════════════════════════════════════════
    def _pg_merge(self):
        c = self._content; c.grid_rowconfigure(6, weight=1)
        self._mg_src  = tk.StringVar()
        self._mg_dest = tk.StringVar()
        nrow = self._path_row(c, "SOURCE FOLDERS  —  click Browse repeatedly to add multiple",
                              self._mg_src, 0, multi=True)
        nrow = self._path_row(c, "DESTINATION FOLDER", self._mg_dest, nrow)
        self._btn(c, "🔀  Merge", self._mg_run, row=nrow, sticky="w", pady=10)
        nrow += 1
        self._mg_log = self._log_panel(c, nrow)

    def _mg_run(self):
        src_raw = self._mg_src.get().strip()
        dest    = self._mg_dest.get().strip()
        if not src_raw or not dest:
            messagebox.showerror("Shelf", "Select source and destination folders."); return
        srcs = [s.strip() for s in src_raw.split(";") if s.strip()]
        self._clear(self._mg_log)
        threading.Thread(target=run_merge_folders,
                         args=(srcs, dest, lambda m: self._log(self._mg_log, m),
                               lambda: None), daemon=True).start()

    # ══════════════════════════════════════════════════════════════════════
    #  PAGE: ORGANISE BY DATE
    # ══════════════════════════════════════════════════════════════════════
    def _pg_date(self):
        c = self._content; c.grid_rowconfigure(4, weight=1)
        self._dt_path = tk.StringVar()
        nrow = self._path_row(c, "FOLDER TO ORGANISE", self._dt_path, 0)
        tk.Label(c, text="Files are moved into sub-folders named  YYYY-MM  based on last-modified date.",
                 font=FONT_BODY, bg=C["bg"], fg=C["fg2"]
                 ).grid(row=nrow, column=0, sticky="w", pady=(4, 0)); nrow += 1
        self._btn(c, "📅  Organise by Date", self._dt_run, row=nrow, sticky="w", pady=10)
        nrow += 1
        self._dt_log = self._log_panel(c, nrow)

    def _dt_run(self):
        folder = self._dt_path.get().strip()
        if not folder: messagebox.showerror("Shelf", "Select a folder."); return
        self._clear(self._dt_log)
        threading.Thread(target=run_organize_by_date,
                         args=(folder, lambda m: self._log(self._dt_log, m),
                               lambda: None), daemon=True).start()


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = ShelfApp()
    app.mainloop()
