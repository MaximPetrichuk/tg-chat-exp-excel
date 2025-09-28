#!/usr/bin/env python3
"""
=== Telegram Chat Exporter to Excel file v1.1 GUI ===
Клиент Telegram для экспорта содержимого чата в файл Excel за указанный период времени.
Версия GUI — графический интерфейс пользователя.
"""

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import asyncio
from datetime import datetime
import core


class ChatExporterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{core.PROGRAM_NAME} {core.PROGRAM_VERSION}")
        self.root.position = "+200+100"

        # --- переменные ---
        self.api_id_var = tk.StringVar()
        self.api_hash_var = tk.StringVar()
        self.phone_var = tk.StringVar()
        self.year_var = tk.StringVar()
        self.month_var = tk.StringVar()
        self.chat_var = tk.StringVar()
        self.log_text = None
        self.progress = None
        self.open_btn = None

        # --- загрузка env ---
        env = core.load_env_vars()
        self.api_id_var.set(env.get("API_ID", ""))
        self.api_hash_var.set(env.get("API_HASH", ""))
        self.phone_var.set(env.get("PHONE", ""))
        self.year_var.set(env.get("YEAR_DEFAULT", ""))
        self.month_var.set(env.get("MONTH_DEFAULT", ""))

        # --- интерфейс ---
        self._build_ui()

    def _build_ui(self):
        # --- Параметры авторизации ---
        frm_auth = ttk.LabelFrame(self.root, text="Авторизация")
        frm_auth.pack(fill="x", padx=10, pady=5)

        ttk.Label(frm_auth, text="API ID:").grid(row=0, column=0, sticky="e")
        ttk.Entry(frm_auth, textvariable=self.api_id_var, width=30).grid(row=0, column=1, sticky="ew")

        ttk.Label(frm_auth, text="API Hash:").grid(row=1, column=0, sticky="e")
        ttk.Entry(frm_auth, textvariable=self.api_hash_var, width=50).grid(row=1, column=1, sticky="ew")

        ttk.Label(frm_auth, text="Телефон:").grid(row=2, column=0, sticky="e")
        ttk.Entry(frm_auth, textvariable=self.phone_var, width=30).grid(row=2, column=1, sticky="ew")

        # --- Чаты ---
        frm_chats = ttk.LabelFrame(self.root, text="Чаты")
        frm_chats.pack(fill="both", expand=True, padx=10, pady=5)

        # Кнопка загрузки чатов
        self.load_btn = ttk.Button(frm_chats, text="Загрузить чаты", command=self.load_chats)
        self.load_btn.grid(row=5, column=0, columnspan=2, pady=5)

        # Список чатов
        ttk.Label(frm_chats, text="Выберите чат:").grid(row=6, column=0, sticky="e")
        self.chat_combo = ttk.Combobox(frm_chats, textvariable=self.chat_var, state="readonly", width=50)
        self.chat_combo.grid(row=6, column=1, sticky="ew")

        # --- Экспорт ---
        frm_export = ttk.LabelFrame(self.root, text="Экспорт сообщений")
        frm_export.pack(fill="x", padx=10, pady=5)

        ttk.Label(frm_export, text="Год:").grid(row=3, column=0, sticky="e")
        ttk.Entry(frm_export, textvariable=self.year_var).grid(row=3, column=1, sticky="ew")

        ttk.Label(frm_export, text="Месяц:").grid(row=4, column=0, sticky="e")
        ttk.Entry(frm_export, textvariable=self.month_var).grid(row=4, column=1, sticky="ew")


        # Кнопка экспорта
        self.export_btn = ttk.Button(frm_export, text="Экспортировать", command=self.export_messages, state="disabled")
        self.export_btn.grid(row=7, column=0, columnspan=2, pady=5)


        # --- Лог ---
        frm_log = ttk.LabelFrame(self.root, text="Лог")
        frm_log.pack(fill="both", expand=True, padx=10, pady=5)

        # Лог
        self.log_text = tk.Text(frm_log, width=80, height=20)
        self.log_text.grid(row=9, column=0, columnspan=2, pady=5, sticky="nsew")

        # Кнопка открыть папку (появляется только после экспорта)
        self.open_btn = ttk.Button(frm_log, text="Открыть папку с файлом", command=self.open_folder, state="disabled")
        self.open_btn.grid(row=10, column=0, columnspan=2, pady=5)

        # Прогресс
        self.progress = ttk.Progressbar(frm_log, mode="indeterminate")
        self.progress.grid(row=8, column=0, columnspan=2, pady=5, sticky="nsew")
        self.progress.grid_remove()  # скрываем до начала экспорта

    def log(self, text):
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, ts + " — " + text + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def save_env(self):
        """Сохраняем введённые параметры в .env"""
        core.save_env_vars(
            self.api_id_var.get(),
            self.api_hash_var.get(),
            self.phone_var.get(),
            self.year_var.get(),
            self.month_var.get()
        )

    def load_chats(self):
        self.save_env()
        self.load_btn.config(state="disabled")
        threading.Thread(target=self._load_chats_thread).start()

    def _load_chats_thread(self):
        asyncio.set_event_loop(asyncio.new_event_loop())  # ✅ фиксим отсутствие event loop
        chats = core.list_chats(
            self.api_id_var.get(),
            self.api_hash_var.get(),
            core.SESSION_NAME,
            self.phone_var.get(),
            log_callback=self.log
        )
        chat_names = [f"{name} ({cid})" for name, cid in chats]
        self.root.after(0, lambda: self._update_chat_list(chat_names))

    def _update_chat_list(self, chat_names):
        self.chat_combo["values"] = chat_names
        if chat_names:
            self.chat_combo.current(0)
            self.export_btn.config(state="normal")
        self.load_btn.config(state="normal")

    def export_messages(self):
        self.save_env()
        chat_text = self.chat_var.get()
        if not chat_text:
            messagebox.showerror("Ошибка", "Не выбран чат")
            return

        try:
            chat_id = int(chat_text.split("(")[-1][:-1])
        except Exception:
            messagebox.showerror("Ошибка", "Не удалось определить ID чата")
            return

        try:
            year = int(self.year_var.get())
            month = int(self.month_var.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректный год и месяц")
            return

        self.export_btn.config(state="disabled")
        self.progress.grid()
        self.progress.start()
        threading.Thread(target=self._export_thread, args=(chat_id, year, month)).start()

    def _export_thread(self, chat_id, year, month):
        asyncio.set_event_loop(asyncio.new_event_loop())  # ✅ фиксим отсутствие event loop
        result = core.export_messages(
            self.api_id_var.get(),
            self.api_hash_var.get(),
            core.SESSION_NAME,
            self.phone_var.get(),
            chat_id,
            year,
            month,
            log_callback=self.log
        )
        self.root.after(0, lambda: self._export_done(result))

    def _export_done(self, result):
        self.progress.stop()
        self.progress.grid_remove()
        self.export_btn.config(state="normal")
        if result.get("success"):
            self.log(f"Экспорт завершён. Файл: {result['filename']}")
            self.open_btn.config(state="normal")
            self.last_file = result["filename"]
        else:
            messagebox.showerror("Ошибка экспорта", result.get("message", "Неизвестная ошибка"))

    def open_folder(self):
        if hasattr(self, "last_file"):
            folder = os.path.dirname(self.last_file)
            if os.name == "nt":
                os.startfile(folder)
            elif os.name == "posix":
                import subprocess
                subprocess.Popen(["xdg-open", folder])

def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")

# Решение для Ctrl+C, Ctrl+V, Ctrl+X в Tkinter
#  https://ru.stackoverflow.com/questions/722885/%D0%92%D1%81%D1%82%D0%B0%D0%B2%D0%BA%D0%B0-%D0%B8%D0%B7-%D0%B1%D1%83%D1%84%D0%B5%D1%80%D0%B0-%D0%B2-tkinter-%D0%B3%D0%BE%D1%80%D1%8F%D1%87%D0%B8%D0%BC%D0%B8-%D0%BA%D0%BB%D0%B0%D0%B2%D0%B8%D1%88%D0%B0%D0%BC%D0%B8-%D0%B2-%D1%80%D1%83%D1%81%D1%81%D0%BA%D0%BE%D0%B9-%D1%80%D0%B0%D1%81%D0%BA%D0%BB%D0%B0%D0%B4%D0%BA%D0%B5
#@staticmethod #можно опустить, если функция вне класса
def CopyPaste(e):
    if e.keycode == 86 and e.keysym != 'v':
        e.widget.event_generate('<<Paste>>')
    elif e.keycode == 67 and e.keysym != 'c':
        e.widget.event_generate('<<Copy>>')
    elif e.keycode == 88 and e.keysym != 'x':
        e.widget.event_generate('<<Cut>>')


def run_gui():
    root = tk.Tk()
    root.bind("<Control-Key>", CopyPaste)
    app = ChatExporterGUI(root)
    center_window(root)

    ok, missing = core.check_env_vars()
    if not ok:
        messagebox.showerror("Ошибка", f"Отсутствуют параметры в .env: {', '.join(missing)}. Откорректируйте файл .env или введите параметры в окне.")
    
    root.mainloop()

if __name__ == "__main__":
    run_gui()