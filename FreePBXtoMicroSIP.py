"""
FreePBXtoMicroSIP Converter
Конвертер справочников АТС → MicroSIP
Поддержка нескольких АТС, префиксов, фильтрации, ручного добавления и импорта XML.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
import os
import json
import re
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET


# ─── Конфигурация ───────────────────────────────────────────────

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pbx2microsip.json")

DEFAULT_CONFIG = {
    "delimiter": "comma",
    "encoding": "utf-8",
    "name_format": "[{prefix}] {name} ({ext})",
    "sort_by": "extension",
    "skip_empty_names": True,
    "skip_numeric_names": True,
    "last_output_dir": "",
    "pbx_sources": [],
    "manual_contacts": [],
    "excluded_contacts": [],
    "window_width": 1100,
    "window_height": 750,
}


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                cfg = DEFAULT_CONFIG.copy()
                cfg.update(saved)
                return cfg
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()


def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ─── Модель данных ──────────────────────────────────────────────

DELIMITERS = {"tab": "\t", "comma": ",", "semicolon": ";"}
ENCODINGS = ["utf-8", "utf-8-sig", "cp1251", "cp866", "latin-1"]
SORT_OPTIONS = {"extension": "По номеру", "name": "По имени", "source": "По АТС"}
NAME_FORMATS = {
    "[{prefix}] {name} ({ext})": "[Офис] Иванов Иван (101)",
    "{name} ({ext})": "Иванов Иван (101)",
    "{ext} - {name}": "101 - Иванов Иван",
    "[{prefix}] {ext} - {name}": "[Офис] 101 - Иванов Иван",
    "{name}": "Иванов Иван",
}


class PBXSource:
    """Один CSV-файл = одна АТС"""

    def __init__(self, filepath="", prefix="", enabled=True):
        self.filepath = filepath
        self.prefix = prefix
        self.enabled = enabled
        self.contacts = []
        self.error = ""

    def load(self, delimiter="\t", encoding="utf-8"):
        self.contacts = []
        self.error = ""
        if not self.filepath or not os.path.exists(self.filepath):
            self.error = "Файл не найден"
            return

        # Определяем тип файла
        ext_lower = os.path.splitext(self.filepath)[1].lower()
        if ext_lower == ".xml":
            self._load_xml()
        else:
            self._load_csv(delimiter, encoding)

    def _load_xml(self):
        """Загрузка контактов из Contacts.xml"""
        try:
            tree = ET.parse(self.filepath)
            root = tree.getroot()

            for contact in root.findall("contact"):
                name = contact.get("name", "").strip()
                number = contact.get("number", "").strip()
                if not number:
                    continue
                if not name:
                    name = number
                self.contacts.append((number, name))

        except ET.ParseError:
            self.error = "Ошибка разбора XML"
        except Exception as e:
            self.error = str(e)

    def _load_csv(self, delimiter, encoding):
        """Загрузка контактов из CSV"""
        try:
            with open(self.filepath, encoding=encoding) as f:
                reader = csv.reader(f, delimiter=delimiter)
                rows = list(reader)

            if not rows:
                self.error = "Файл пуст"
                return

            first_row = rows[0]
            has_header = False
            ext_col = 0
            name_col = 2

            header_keywords_ext = {"extension", "ext", "number", "номер", "экстеншен", "внутренний"}
            header_keywords_name = {"name", "имя", "фио", "название", "callerid", "caller_id"}

            for i, val in enumerate(first_row):
                vl = val.strip().lower()
                if vl in header_keywords_ext:
                    ext_col = i
                    has_header = True
                if vl in header_keywords_name:
                    name_col = i
                    has_header = True

            if has_header and name_col == 2:
                name_col = ext_col + 1 if ext_col + 1 < len(first_row) else ext_col

            data_rows = rows[1:] if has_header else rows

            for row in data_rows:
                if len(row) <= ext_col:
                    continue
                ext = row[ext_col].strip()
                name = row[name_col].strip() if name_col < len(row) else ""

                if not ext:
                    continue

                callerid_match = re.match(r'^(.+?)\s*<\d+>$', name)
                if callerid_match:
                    name = callerid_match.group(1).strip()

                if name.lower() == "device" or name.lower().startswith("device <"):
                    name = ""

                self.contacts.append((ext, name))

        except UnicodeDecodeError:
            self.error = f"Ошибка кодировки. Попробуйте другую (текущая: {encoding})"
        except Exception as e:
            self.error = str(e)


# ─── GUI ────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.config = load_config()
        self.sources: list[PBXSource] = []
        self.manual_contacts: list[dict] = list(self.config.get("manual_contacts", []))
        self.excluded_contacts: set = set(self.config.get("excluded_contacts", []))

        self.title("FreePBXtoMicroSIP Converter")
        self.geometry(f"{self.config['window_width']}x{self.config['window_height']}")
        self.minsize(900, 550)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Status.TLabel", font=("Segoe UI", 9))

        self._build_ui()
        self._restore_sources()

    def _build_ui(self):
        # ─── Верхняя панель ─────────────────────────────
        top = ttk.Frame(self, padding=10)
        top.pack(fill=tk.X)

        ttk.Label(top, text="FreePBXtoMicroSIP Converter", style="Header.TLabel").pack(side=tk.LEFT)

        tk.Button(top, text="Экспорт XML", command=self._export,
                  font=("Segoe UI", 10, "bold"), bg="#4a90d9", fg="white").pack(side=tk.RIGHT, padx=4)

        # ─── Основная область ───────────────────────────
        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # ════════════════════════════════════════════════
        # Левая панель
        # ════════════════════════════════════════════════
        left = ttk.Frame(paned)
        paned.add(left, weight=1)

        # --- Источники ---
        src_frame = ttk.LabelFrame(left, text="Источники (CSV / XML файлы)", padding=8)
        src_frame.pack(fill=tk.BOTH, expand=True)

        self.sources_list = tk.Listbox(src_frame, font=("Segoe UI", 10),
                                        selectmode=tk.SINGLE, activestyle="none")
        self.sources_list.pack(fill=tk.BOTH, expand=True)

        btn_row = ttk.Frame(src_frame)
        btn_row.pack(fill=tk.X, pady=(6, 0))
        tk.Button(btn_row, text="+ Добавить файлы", command=self._add_files).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_row, text="Удалить", command=self._remove_source).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_row, text="Префикс", command=self._edit_prefix).pack(side=tk.LEFT, padx=2)

        # --- Ручное добавление ---
        manual_frame = ttk.LabelFrame(left, text="Добавить контакт вручную", padding=8)
        manual_frame.pack(fill=tk.X, pady=(8, 0))

        ttk.Label(manual_frame, text="Номер:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.manual_ext_var = tk.StringVar()
        ttk.Entry(manual_frame, textvariable=self.manual_ext_var, width=12).grid(
            row=0, column=1, sticky=tk.W, pady=2, padx=6)

        ttk.Label(manual_frame, text="Имя:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.manual_name_var = tk.StringVar()
        ttk.Entry(manual_frame, textvariable=self.manual_name_var, width=25).grid(
            row=1, column=1, sticky=tk.W, pady=2, padx=6)

        ttk.Label(manual_frame, text="АТС:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.manual_prefix_var = tk.StringVar(value="Вручную")
        ttk.Entry(manual_frame, textvariable=self.manual_prefix_var, width=15).grid(
            row=2, column=1, sticky=tk.W, pady=2, padx=6)

        manual_btn_row = ttk.Frame(manual_frame)
        manual_btn_row.grid(row=3, column=0, columnspan=2, pady=(6, 0))
        tk.Button(manual_btn_row, text="Добавить", command=self._add_manual_contact).pack(side=tk.LEFT, padx=2)

        # --- Настройки ---
        settings_frame = ttk.LabelFrame(left, text="Настройки", padding=8)
        settings_frame.pack(fill=tk.X, pady=(8, 0))

        row = 0
        ttk.Label(settings_frame, text="Разделитель CSV:").grid(row=row, column=0, sticky=tk.W, pady=3)
        self.delim_var = tk.StringVar(value=self.config["delimiter"])
        ttk.Combobox(settings_frame, textvariable=self.delim_var,
                     values=["tab", "comma", "semicolon"], state="readonly", width=12).grid(
            row=row, column=1, sticky=tk.W, pady=3, padx=6)
        row += 1

        ttk.Label(settings_frame, text="Кодировка CSV:").grid(row=row, column=0, sticky=tk.W, pady=3)
        self.enc_var = tk.StringVar(value=self.config["encoding"])
        ttk.Combobox(settings_frame, textvariable=self.enc_var,
                     values=ENCODINGS, state="readonly", width=12).grid(
            row=row, column=1, sticky=tk.W, pady=3, padx=6)
        row += 1

        ttk.Label(settings_frame, text="Формат имени:").grid(row=row, column=0, sticky=tk.W, pady=3)
        self.fmt_var = tk.StringVar(value=self.config["name_format"])
        ttk.Combobox(settings_frame, textvariable=self.fmt_var,
                     values=list(NAME_FORMATS.keys()), width=28).grid(
            row=row, column=1, sticky=tk.W, pady=3, padx=6)
        row += 1

        ttk.Label(settings_frame, text="Сортировка:").grid(row=row, column=0, sticky=tk.W, pady=3)
        self.sort_var = tk.StringVar(value=self.config["sort_by"])
        ttk.Combobox(settings_frame, textvariable=self.sort_var,
                     values=list(SORT_OPTIONS.keys()), state="readonly", width=12).grid(
            row=row, column=1, sticky=tk.W, pady=3, padx=6)
        row += 1

        self.skip_empty_var = tk.BooleanVar(value=self.config["skip_empty_names"])
        ttk.Checkbutton(settings_frame, text="Пустые имена → номер",
                         variable=self.skip_empty_var).grid(
            row=row, column=0, columnspan=2, sticky=tk.W, pady=2)
        row += 1

        self.skip_num_var = tk.BooleanVar(value=self.config["skip_numeric_names"])
        ttk.Checkbutton(settings_frame, text="Пропускать где имя = число",
                         variable=self.skip_num_var).grid(
            row=row, column=0, columnspan=2, sticky=tk.W, pady=2)
        row += 1

        settings_btn_row = tk.Frame(settings_frame)
        settings_btn_row.grid(row=row, column=0, columnspan=2, pady=(8, 2))
        tk.Button(settings_btn_row, text="Применить", command=self._on_settings_changed,
                  bg="#5cb85c", fg="white").pack(side=tk.LEFT, padx=4)
        tk.Button(settings_btn_row, text="Сбросить всё", command=self._reset_config,
                  bg="#d9534f", fg="white").pack(side=tk.LEFT, padx=4)

        # ════════════════════════════════════════════════
        # Правая панель — предпросмотр
        # ════════════════════════════════════════════════
        right = ttk.LabelFrame(paned, text="Предпросмотр (как будет в MicroSIP)", padding=8)
        paned.add(right, weight=2)

        search_frame = ttk.Frame(right)
        search_frame.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(search_frame, text="Поиск:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._refresh_preview())
        ttk.Entry(search_frame, textvariable=self.search_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)

        self.count_label = ttk.Label(search_frame, text="0 контактов", style="Status.TLabel")
        self.count_label.pack(side=tk.RIGHT)

        columns = ("name", "number")
        self.tree = ttk.Treeview(right, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("name", text="Name")
        self.tree.heading("number", text="Number")
        self.tree.column("name", width=350, minwidth=200)
        self.tree.column("number", width=80, minwidth=60)

        scrollbar = ttk.Scrollbar(right, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        tree_frame = ttk.Frame(right)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, in_=tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, in_=tree_frame)

        # Кнопки управления контактами
        preview_btn_row = ttk.Frame(right)
        preview_btn_row.pack(fill=tk.X, pady=(6, 0))
        tk.Button(preview_btn_row, text="Удалить выбранные", command=self._delete_selected_contacts).pack(side=tk.LEFT, padx=2)
        tk.Button(preview_btn_row, text="Восстановить удалённые", command=self._restore_excluded).pack(side=tk.LEFT, padx=2)
        self.excluded_label = ttk.Label(preview_btn_row, text="", style="Status.TLabel")
        self.excluded_label.pack(side=tk.RIGHT, padx=4)

        # ─── Статусбар ─────────────────────────────────
        self.status_var = tk.StringVar(value="Добавьте CSV или XML файлы")
        ttk.Label(self, textvariable=self.status_var, style="Status.TLabel",
                  padding=(10, 4)).pack(fill=tk.X, side=tk.BOTTOM)

    # ─── Настройки ──────────────────────────────────────

    def _on_settings_changed(self):
        self.config["delimiter"] = self.delim_var.get()
        self.config["encoding"] = self.enc_var.get()
        self.config["name_format"] = self.fmt_var.get()
        self.config["sort_by"] = self.sort_var.get()
        self.config["skip_empty_names"] = self.skip_empty_var.get()
        self.config["skip_numeric_names"] = self.skip_num_var.get()
        save_config(self.config)
        self._reload_all()
        self.status_var.set("Настройки применены")

    def _reset_config(self):
        if not messagebox.askyesno("Сброс настроек",
                "Сбросить все настройки, источники и ручные контакты?\n\n"
                "Это удалит файл конфигурации и перезапустит программу."):
            return

        # Удаляем JSON
        if os.path.exists(CONFIG_FILE):
            os.remove(CONFIG_FILE)

        # Сбрасываем состояние
        self.config = DEFAULT_CONFIG.copy()
        self.sources.clear()
        self.manual_contacts.clear()
        self.excluded_contacts.clear()

        # Обновляем виджеты настроек
        self.delim_var.set(self.config["delimiter"])
        self.enc_var.set(self.config["encoding"])
        self.fmt_var.set(self.config["name_format"])
        self.sort_var.set(self.config["sort_by"])
        self.skip_empty_var.set(self.config["skip_empty_names"])
        self.skip_num_var.set(self.config["skip_numeric_names"])

        self._reload_all()
        self.status_var.set("Настройки сброшены до заводских")

    # ─── Ручные контакты ────────────────────────────────

    def _add_manual_contact(self):
        ext = self.manual_ext_var.get().strip()
        name = self.manual_name_var.get().strip()
        prefix = self.manual_prefix_var.get().strip() or "Вручную"

        if not ext:
            messagebox.showwarning("Ошибка", "Укажите номер.")
            return
        if not name:
            name = ext

        self.manual_contacts.append({"ext": ext, "name": name, "prefix": prefix})
        self.config["manual_contacts"] = self.manual_contacts
        save_config(self.config)

        self.manual_ext_var.set("")
        self.manual_name_var.set("")

        self._refresh_preview()
        self.status_var.set(f"Добавлен контакт: {name} ({ext})")

    def _delete_selected_contacts(self):
        """Удалить выбранные контакты (один или несколько)"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Подсказка", "Выберите контакты в таблице.\n"
                "Для выбора нескольких: Ctrl+клик или Shift+клик.")
            return

        count = len(selected)
        if count > 5:
            if not messagebox.askyesno("Подтверждение",
                    f"Удалить {count} контактов из справочника?"):
                return

        removed = 0
        for item_id in selected:
            values = self.tree.item(item_id)["values"]
            display_name = str(values[0])
            ext = str(values[1])

            # Сначала ищем среди ручных контактов — удаляем полностью
            found_manual = None
            for i, mc in enumerate(self.manual_contacts):
                fmt = self.config["name_format"]
                mc_display = fmt.format(name=mc["name"], ext=mc["ext"], prefix=mc["prefix"])
                if mc_display == display_name and mc["ext"] == ext:
                    found_manual = i
                    break

            if found_manual is not None:
                self.manual_contacts.pop(found_manual)
            else:
                # Для контактов из файлов — добавляем в исключения
                exclude_key = f"{display_name}||{ext}"
                self.excluded_contacts.add(exclude_key)

            removed += 1

        self.config["manual_contacts"] = self.manual_contacts
        self.config["excluded_contacts"] = list(self.excluded_contacts)
        save_config(self.config)
        self._refresh_preview()
        self.status_var.set(f"Удалено {removed} контактов")

    def _restore_excluded(self):
        """Восстановить все ранее удалённые контакты из файлов"""
        if not self.excluded_contacts:
            messagebox.showinfo("Нет удалённых", "Нет удалённых контактов для восстановления.")
            return

        count = len(self.excluded_contacts)
        if not messagebox.askyesno("Восстановление",
                f"Восстановить {count} удалённых контактов?"):
            return

        self.excluded_contacts.clear()
        self.config["excluded_contacts"] = []
        save_config(self.config)
        self._refresh_preview()
        self.status_var.set(f"Восстановлено {count} контактов")

    # ─── Источники (CSV + XML) ──────────────────────────

    def _restore_sources(self):
        for src_data in self.config.get("pbx_sources", []):
            src = PBXSource(
                filepath=src_data.get("filepath", ""),
                prefix=src_data.get("prefix", ""),
                enabled=src_data.get("enabled", True),
            )
            if os.path.exists(src.filepath):
                self.sources.append(src)
        self._reload_all()

    def _save_sources(self):
        self.config["pbx_sources"] = [
            {"filepath": s.filepath, "prefix": s.prefix, "enabled": s.enabled}
            for s in self.sources
        ]
        save_config(self.config)

    def _add_files(self):
        files = filedialog.askopenfilenames(
            title="Выберите CSV или XML файлы",
            filetypes=[
                ("CSV и XML файлы", "*.csv *.tsv *.txt *.xml"),
                ("CSV файлы", "*.csv *.tsv *.txt"),
                ("XML файлы", "*.xml"),
                ("Все файлы", "*.*"),
            ],
        )
        if not files:
            return

        for fpath in files:
            if any(s.filepath == fpath for s in self.sources):
                continue
            basename = os.path.splitext(os.path.basename(fpath))[0]
            prefix = self._ask_prefix(basename)
            if prefix is None:
                continue
            src = PBXSource(filepath=fpath, prefix=prefix)
            self.sources.append(src)

        self._reload_all()
        self._save_sources()

    def _ask_prefix(self, default=""):
        dialog = tk.Toplevel(self)
        dialog.title("Название АТС")
        dialog.geometry("350x140")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        ttk.Label(dialog, text="Введите название АТС (префикс):").pack(padx=20, pady=(15, 5), anchor=tk.W)
        ttk.Label(dialog, text="Например: Офис, Склад, Филиал МСК").pack(padx=20, pady=(0, 5), anchor=tk.W)

        entry = ttk.Entry(dialog, width=35)
        entry.pack(padx=20, pady=5)
        entry.insert(0, default)
        entry.select_range(0, tk.END)
        entry.focus_set()

        result = [None]

        def on_ok(event=None):
            result[0] = entry.get().strip()
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        entry.bind("<Return>", on_ok)
        entry.bind("<Escape>", lambda e: on_cancel())

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=8)
        tk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=4)
        tk.Button(btn_frame, text="Отмена", command=on_cancel, width=10).pack(side=tk.LEFT, padx=4)

        self.wait_window(dialog)
        return result[0]

    def _remove_source(self):
        sel = self.sources_list.curselection()
        if not sel:
            return
        idx = sel[0]
        del self.sources[idx]
        self._reload_all()
        self._save_sources()

    def _edit_prefix(self):
        sel = self.sources_list.curselection()
        if not sel:
            return
        idx = sel[0]
        src = self.sources[idx]
        new_prefix = self._ask_prefix(src.prefix)
        if new_prefix is not None:
            src.prefix = new_prefix
            self._reload_all()
            self._save_sources()

    # ─── Данные и отображение ───────────────────────────

    def _reload_all(self):
        delim = DELIMITERS.get(self.config["delimiter"], "\t")
        enc = self.config["encoding"]

        for src in self.sources:
            src.load(delimiter=delim, encoding=enc)

        self.sources_list.delete(0, tk.END)
        for src in self.sources:
            ext_type = "XML" if src.filepath.lower().endswith(".xml") else "CSV"
            status = f"✓ {len(src.contacts)}" if not src.error else "✗ ошибка"
            label = f"[{src.prefix}] {os.path.basename(src.filepath)} ({ext_type}) — {status}"
            self.sources_list.insert(tk.END, label)
            if src.error:
                self.sources_list.itemconfig(self.sources_list.size() - 1, fg="red")

        errors = [f"[{s.prefix}] {s.error}" for s in self.sources if s.error]
        if errors:
            self.status_var.set("Ошибки: " + "; ".join(errors))
        else:
            total = sum(len(s.contacts) for s in self.sources) + len(self.manual_contacts)
            src_count = len(self.sources)
            manual_count = len(self.manual_contacts)
            parts = []
            if src_count:
                parts.append(f"{src_count} файлов")
            if manual_count:
                parts.append(f"{manual_count} вручную")
            self.status_var.set(f"Загружено {total} контактов ({', '.join(parts)})" if parts else "Нет контактов")

        self._refresh_preview()

    def _get_all_contacts(self):
        contacts = []

        for src in self.sources:
            if not src.enabled:
                continue
            for ext, name in src.contacts:
                if self.config["skip_empty_names"] and not name:
                    name = ext
                if self.config["skip_numeric_names"] and name.isdigit():
                    continue

                fmt = self.config["name_format"]
                display = fmt.format(name=name, ext=ext, prefix=src.prefix)

                # Проверяем исключения
                exclude_key = f"{display}||{ext}"
                if exclude_key in self.excluded_contacts:
                    continue

                contacts.append({
                    "source": src.prefix,
                    "ext": ext,
                    "name": name,
                    "display": display,
                })

        for mc in self.manual_contacts:
            fmt = self.config["name_format"]
            display = fmt.format(name=mc["name"], ext=mc["ext"], prefix=mc["prefix"])
            contacts.append({
                "source": mc["prefix"],
                "ext": mc["ext"],
                "name": mc["name"],
                "display": display,
            })

        sort_key = self.config["sort_by"]
        if sort_key == "extension":
            contacts.sort(key=lambda c: (c["source"], int(c["ext"]) if c["ext"].isdigit() else 999999))
        elif sort_key == "name":
            contacts.sort(key=lambda c: (c["source"], c["name"].lower()))
        else:
            contacts.sort(key=lambda c: (c["source"], int(c["ext"]) if c["ext"].isdigit() else 999999))

        return contacts

    def _refresh_preview(self):
        self.tree.delete(*self.tree.get_children())
        contacts = self._get_all_contacts()

        search = self.search_var.get().lower().strip()
        filtered = [
            c for c in contacts
            if not search or search in c["name"].lower() or search in c["ext"] or search in c["source"].lower()
        ]

        for c in filtered:
            self.tree.insert("", tk.END, values=(c["display"], c["ext"]))

        self.count_label.config(text=f"{len(filtered)} из {len(contacts)} контактов")

        # Показать количество удалённых
        excl_count = len(self.excluded_contacts)
        if excl_count:
            self.excluded_label.config(text=f"Скрыто: {excl_count}")
        else:
            self.excluded_label.config(text="")

    # ─── Экспорт ────────────────────────────────────────

    def _export(self):
        contacts = self._get_all_contacts()
        if not contacts:
            messagebox.showwarning("Нет данных", "Нет контактов для экспорта.")
            return

        initial_dir = self.config.get("last_output_dir", "") or os.path.expanduser("~")
        filepath = filedialog.asksaveasfilename(
            title="Сохранить Contacts.xml",
            defaultextension=".xml",
            initialfile="Contacts.xml",
            initialdir=initial_dir,
            filetypes=[("XML файлы", "*.xml"), ("Все файлы", "*.*")],
        )
        if not filepath:
            return

        self.config["last_output_dir"] = os.path.dirname(filepath)
        save_config(self.config)

        try:
            lines = ['<?xml version="1.0" encoding="utf-8"?>', "<contacts>"]
            for c in contacts:
                name_escaped = escape(c["display"])
                lines.append(f'  <contact name="{name_escaped}" number="{c["ext"]}" />')
            lines.append("</contacts>")

            with open(filepath, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))

            self.status_var.set(f"Экспортировано {len(contacts)} контактов → {filepath}")
            messagebox.showinfo(
                "Готово!",
                f"Сохранено {len(contacts)} контактов.\n\n"
                f"Скопируйте файл в папку с MicroSIP\nи перезапустите программу.",
            )
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")

    def destroy(self):
        self.config["window_width"] = self.winfo_width()
        self.config["window_height"] = self.winfo_height()
        save_config(self.config)
        super().destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
