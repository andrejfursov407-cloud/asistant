# -*- coding: utf-8 -*-
"""
Помощник AI — многофункциональный настольный помощник.
Требования: Python 3.10+, customtkinter, openai, PyPDF2, python-docx.
Веб-поиск использует стандартную библиотеку urllib, без зависимости requests.
"""
import os
import sys
import json
import threading
import urllib.parse
import urllib.request
import urllib.error
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from openai import OpenAI
import PyPDF2
import docx

APP_NAME = "Помощник AI"
CONFIG_FILE = Path.home() / ".pomoshchnik_ai_config.json"
MAX_CONTEXT_CHARS = 70000

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


PROVIDERS = {
    "Google Gemini": {
        "url": "https://generativelanguage.googleapis.com/v1beta/openai/",
        "model": "gemini-2.5-flash",
        "help": "Создайте API-ключ в Google AI Studio. Доступность бесплатного тарифа зависит от региона и текущих условий."
    },
    "Groq": {
        "url": "https://api.groq.com/openai/v1",
        "model": "llama-3.3-70b-versatile",
        "help": "Создайте API-ключ в консоли Groq. Модели и лимиты могут меняться."
    },
    "OpenRouter (free models)": {
        "url": "https://openrouter.ai/api/v1",
        "model": "meta-llama/llama-3.3-70b-instruct:free",
        "help": "Нужен API-ключ OpenRouter. Бесплатные модели и их доступность меняются."
    },
}


class AssistantApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1240x790")
        self.minsize(980, 650)
        self.protocol("WM_DELETE_WINDOW", self.close_app)

        self.document_text = ""
        self.document_names = []
        self.busy = False
        self.history = []
        self._load_config()
        self._build_ui()
        self._show_welcome()

    def _load_config(self):
        self.saved_provider = "Google Gemini"
        self.saved_key = ""
        try:
            if CONFIG_FILE.exists():
                data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
                self.saved_provider = data.get("provider", self.saved_provider)
                self.saved_key = data.get("api_key", "")
        except Exception:
            pass

    def _save_config(self, silent=False):
        try:
            CONFIG_FILE.write_text(json.dumps({
                "provider": self.provider_menu.get(),
                "api_key": self.api_entry.get().strip()
            }, ensure_ascii=False, indent=2), encoding="utf-8")
            if not silent:
                self._status("Настройки сохранены")
                messagebox.showinfo(APP_NAME, "Настройки сохранены в профиле текущего пользователя.")
        except Exception as e:
            if not silent:
                messagebox.showerror("Ошибка сохранения", str(e))

    def close_app(self):
        self._save_config(silent=True)
        self.destroy()

    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Sidebar
        self.sidebar = ctk.CTkFrame(self, width=270, corner_radius=0, fg_color="#101827")
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)
        self.sidebar.grid_columnconfigure(0, weight=1)
        self.sidebar.grid_rowconfigure(12, weight=1)

        ctk.CTkLabel(self.sidebar, text="✦ Помощник AI",
                     font=("Segoe UI", 23, "bold"), text_color="#67e8f9").grid(
            row=0, column=0, padx=20, pady=(22, 4), sticky="w")
        ctk.CTkLabel(self.sidebar, text="ВАШ РАБОЧИЙ AI-ИНСТРУМЕНТ",
                     font=("Segoe UI", 9, "bold"), text_color="#94a3b8").grid(
            row=1, column=0, padx=21, pady=(0, 18), sticky="w")

        ctk.CTkLabel(self.sidebar, text="ИНСТРУМЕНТЫ", font=("Segoe UI", 10, "bold"),
                     text_color="#64748b").grid(row=2, column=0, padx=20, sticky="w")
        self.tool_buttons = {}
        tools = [
            ("✎  Написать текст", "Написать текст"),
            ("⌕  Поиск в интернете", "Поиск в интернете"),
            ("▤  Анализ документов", "Анализ документов"),
            ("⇄  Объединить файлы", "Объединить файлы"),
            ("✧  Универсальный чат", "Обычный ответ"),
        ]
        for i, (label, mode) in enumerate(tools, start=3):
            b = ctk.CTkButton(self.sidebar, text=label, anchor="w", height=39,
                              corner_radius=8, fg_color="#1e293b", hover_color="#334155",
                              command=lambda m=mode: self.select_tool(m))
            b.grid(row=i, column=0, padx=14, pady=4, sticky="ew")
            self.tool_buttons[mode] = b

        ctk.CTkFrame(self.sidebar, height=1, fg_color="#334155").grid(
            row=8, column=0, padx=18, pady=14, sticky="ew")
        ctk.CTkLabel(self.sidebar, text="МОДЕЛЬ ИИ", font=("Segoe UI", 10, "bold"),
                     text_color="#64748b").grid(row=9, column=0, padx=20, sticky="w")
        self.provider_menu = ctk.CTkOptionMenu(
            self.sidebar, values=list(PROVIDERS.keys()), command=self.provider_changed,
            fg_color="#263449", button_color="#0e7490", button_hover_color="#155e75")
        self.provider_menu.grid(row=10, column=0, padx=18, pady=(7, 8), sticky="ew")
        if self.saved_provider in PROVIDERS:
            self.provider_menu.set(self.saved_provider)
        self.api_entry = ctk.CTkEntry(self.sidebar, show="*", placeholder_text="API-ключ")
        self.api_entry.grid(row=11, column=0, padx=18, pady=(0, 8), sticky="ew")
        if self.saved_key:
            self.api_entry.insert(0, self.saved_key)
        ctk.CTkButton(self.sidebar, text="Сохранить настройки", command=self._save_config,
                      fg_color="#334155", hover_color="#475569").grid(
            row=12, column=0, padx=18, pady=(2, 8), sticky="new")
        self.key_help = ctk.CTkLabel(self.sidebar, text="", wraplength=225, justify="left",
                                     font=("Segoe UI", 10), text_color="#94a3b8")
        self.key_help.grid(row=13, column=0, padx=18, pady=(0, 18), sticky="sew")
        self.provider_changed(self.provider_menu.get())

        # Main panel
        self.main = ctk.CTkFrame(self, corner_radius=0, fg_color="#0b1220")
        self.main.grid(row=0, column=1, sticky="nsew")
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(1, weight=1)

        self.header = ctk.CTkFrame(self.main, height=66, corner_radius=0, fg_color="#111c2e")
        self.header.grid(row=0, column=0, sticky="ew")
        self.header.grid_columnconfigure(0, weight=1)
        self.page_title = ctk.CTkLabel(self.header, text="Универсальный чат",
                                       font=("Segoe UI", 18, "bold"))
        self.page_title.grid(row=0, column=0, padx=24, pady=(12, 0), sticky="w")
        self.page_subtitle = ctk.CTkLabel(self.header, text="Тексты, документы и поиск в одном окне",
                                          font=("Segoe UI", 11), text_color="#94a3b8")
        self.page_subtitle.grid(row=1, column=0, padx=24, pady=(0, 10), sticky="w")
        ctk.CTkButton(self.header, text="＋ Добавить файлы", width=130,
                      command=self.add_documents).grid(row=0, column=1, rowspan=2, padx=20)

        self.chat = ctk.CTkTextbox(self.main, wrap="word", font=("Segoe UI", 13),
                                   fg_color="#0b1220", border_width=0)
        self.chat.grid(row=1, column=0, padx=22, pady=(15, 8), sticky="nsew")
        self.chat.configure(state="disabled")

        self.docs_bar = ctk.CTkFrame(self.main, fg_color="#111c2e", corner_radius=9)
        self.docs_bar.grid(row=2, column=0, padx=22, pady=(0, 8), sticky="ew")
        self.docs_label = ctk.CTkLabel(self.docs_bar, text="Документы не загружены",
                                       anchor="w", text_color="#94a3b8")
        self.docs_label.pack(side="left", padx=12, pady=7, fill="x", expand=True)
        ctk.CTkButton(self.docs_bar, text="Сохранить объединённый DOCX",
                      width=210, command=self.export_docx, fg_color="#0f766e").pack(
            side="right", padx=8, pady=5)
        ctk.CTkButton(self.docs_bar, text="Сохранить TXT", width=125,
                      command=self.export_txt, fg_color="#0f766e").pack(side="right", padx=4, pady=5)

        self.mode_menu = ctk.CTkOptionMenu(
            self.main, values=[x[1] for x in tools], command=self.select_tool, width=210,
            fg_color="#263449", button_color="#0e7490", button_hover_color="#155e75")
        self.mode_menu.set("Обычный ответ")
        self.mode_menu.grid(row=3, column=0, padx=22, pady=(4, 7), sticky="w")

        self.input = ctk.CTkTextbox(self.main, height=95, wrap="word", font=("Segoe UI", 13),
                                    fg_color="#111c2e", border_color="#334155", border_width=1)
        self.input.grid(row=4, column=0, padx=22, pady=(0, 9), sticky="ew")
        self.input.bind("<Control-Return>", lambda e: self.send())
        self.footer = ctk.CTkFrame(self.main, fg_color="transparent")
        self.footer.grid(row=5, column=0, padx=22, pady=(0, 15), sticky="ew")
        self.status_label = ctk.CTkLabel(self.footer, text="Готов к работе",
                                         text_color="#94a3b8", anchor="w")
        self.status_label.pack(side="left", fill="x", expand=True)
        self.send_btn = ctk.CTkButton(self.footer, text="Отправить  ➤", width=145,
                                      height=39, command=self.send, corner_radius=9)
        self.send_btn.pack(side="right")

    def provider_changed(self, provider):
        info = PROVIDERS.get(provider, {})
        self.key_help.configure(text=info.get("help", ""))

    def select_tool(self, mode):
        self.mode_menu.set(mode)
        titles = {
            "Написать текст": ("Написать текст", "Укажите тип текста, цель, аудиторию и желаемый стиль"),
            "Поиск в интернете": ("Поиск в интернете", "Найду результаты, затем попрошу ИИ составить ответ"),
            "Анализ документов": ("Анализ документов", "Сначала добавьте документы кнопкой в верхней панели"),
            "Объединить файлы": ("Объединить файлы", "Добавьте несколько файлов и сохраните общий TXT или DOCX"),
            "Обычный ответ": ("Универсальный чат", "Задайте вопрос или поручите задачу помощнику"),
        }
        title, subtitle = titles.get(mode, (mode, ""))
        self.page_title.configure(text=title)
        self.page_subtitle.configure(text=subtitle)
        for key, button in self.tool_buttons.items():
            button.configure(fg_color="#0e7490" if key == mode else "#1e293b")

    def _show_welcome(self):
        self._append("Помощник AI", "Здравствуйте! Я помогу написать текст, найти информацию в интернете, проанализировать документы и объединить несколько файлов.\n\n1. Выберите инструмент слева.\n2. При необходимости добавьте PDF, DOCX или TXT.\n3. Введите API-ключ выбранного провайдера и отправьте запрос.\n\nУ бесплатных API есть ограничения, а доступность моделей может меняться. Не отправляйте конфиденциальные документы во внешние сервисы без разрешения вашей организации.")

    def _append(self, who, text):
        self.chat.configure(state="normal")
        self.chat.insert("end", f"\n{who}\n{'─' * 48}\n{text}\n\n")
        self.chat.configure(state="disabled")
        self.chat.see("end")

    def _status(self, text):
        self.after(0, lambda: self.status_label.configure(text=text))

    def _read_file(self, path):
        ext = path.suffix.lower()
        if ext == ".pdf":
            with path.open("rb") as f:
                reader = PyPDF2.PdfReader(f)
                return "\n".join(page.extract_text() or "" for page in reader.pages)
        if ext == ".docx":
            document = docx.Document(str(path))
            chunks = [p.text for p in document.paragraphs]
            for table in document.tables:
                for row in table.rows:
                    chunks.append(" | ".join(cell.text for cell in row.cells))
            return "\n".join(chunks)
        if ext == ".txt":
            return path.read_text(encoding="utf-8-sig")
        raise ValueError("Поддерживаются только PDF, DOCX и TXT.")

    def add_documents(self):
        paths = filedialog.askopenfilenames(
            title="Добавить документы",
            filetypes=[("Документы", "*.pdf *.docx *.txt"), ("Все файлы", "*.*")])
        if not paths:
            return
        names, chunks, errors = [], [], []
        for raw in paths:
            path = Path(raw)
            try:
                content = self._read_file(path).strip()
                if not content:
                    errors.append(f"{path.name}: текст не извлечён (PDF может быть сканом).")
                    continue
                names.append(path.name)
                chunks.append(f"===== ДОКУМЕНТ: {path.name} =====\n{content}")
            except Exception as exc:
                errors.append(f"{path.name}: {exc}")
        if chunks:
            self.document_names = names
            self.document_text = "\n\n".join(chunks)
            self.docs_label.configure(text=f"Загружено документов: {len(names)} · " + ", ".join(names)[:180])
            self._append("Система", f"Загружено документов: {len(names)}.\n" + "\n".join(names))
            self.select_tool("Анализ документов")
        if errors:
            messagebox.showwarning("Не все файлы обработаны", "\n".join(errors[:8]))

    def export_txt(self):
        if not self.document_text:
            messagebox.showinfo(APP_NAME, "Сначала добавьте документы.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".txt",
            initialfile="объединённые_документы.txt", filetypes=[("Текстовый файл", "*.txt")])
        if path:
            Path(path).write_text(self.document_text, encoding="utf-8")
            self._status("TXT-файл сохранён")

    def export_docx(self):
        if not self.document_text:
            messagebox.showinfo(APP_NAME, "Сначала добавьте документы.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".docx",
            initialfile="объединённые_документы.docx", filetypes=[("Word", "*.docx")])
        if path:
            out = docx.Document()
            out.add_heading("Объединённые документы", 0)
            for i, name in enumerate(self.document_names):
                out.add_heading(name, level=1)
                marker = f"===== ДОКУМЕНТ: {name} =====\n"
                content = self.document_text.split(marker, 1)[1] if marker in self.document_text else ""
                # Keep only the portion belonging to this file.
                next_marker = "\n\n===== ДОКУМЕНТ: "
                content = content.split(next_marker, 1)[0]
                for paragraph in content.splitlines():
                    if paragraph.strip():
                        out.add_paragraph(paragraph)
            out.save(path)
            self._status("DOCX-файл сохранён")

    def send(self):
        if self.busy:
            return
        prompt = self.input.get("1.0", "end").strip()
        mode = self.mode_menu.get()
        if mode == "Объединить файлы":
            self.export_docx()
            return
        if mode == "Анализ документов" and not self.document_text:
            messagebox.showwarning(APP_NAME, "Сначала добавьте документы.")
            return
        if not prompt and mode == "Анализ документов":
            prompt = "Подробно проанализируй загруженные документы: краткое содержание, ключевые положения, противоречия, риски, выводы и практические рекомендации. Указывай названия источников."
        if not prompt:
            return
        api_key = self.api_entry.get().strip()
        if not api_key:
            messagebox.showwarning("Нужен API-ключ", "Введите API-ключ выбранного провайдера. Бесплатный доступ зависит от условий сервиса.")
            return
        provider = self.provider_menu.get()
        config = PROVIDERS[provider]
        self.input.delete("1.0", "end")
        self._append("Вы", prompt)
        self.busy = True
        self.send_btn.configure(state="disabled", text="Работаю…")
        self._status("Выполняется запрос…")
        threading.Thread(target=self._worker, args=(prompt, mode, api_key, provider, config), daemon=True).start()

    def _web_search(self, query):
        url = "https://html.duckduckgo.com/html/?" + urllib.parse.urlencode({"q": query})
        request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(request, timeout=20) as response:
            page = response.read().decode("utf-8", errors="replace")
        # Extract result blocks without adding a third-party HTML parser dependency.
        import re
        links = re.findall(r'<a[^>]+class="result__a"[^>]+href="([^"]+)"[^>]*>(.*?)</a>', page, re.I | re.S)
        snippets = re.findall(r'<a[^>]+class="result__snippet"[^>]*>(.*?)</a>|<div[^>]+class="result__snippet"[^>]*>(.*?)</div>', page, re.I | re.S)
        clean = lambda s: re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", s)).strip()
        results = []
        for i, (href, title) in enumerate(links[:7]):
            snippet = ""
            if i < len(snippets):
                snippet = clean(snippets[i][0] or snippets[i][1])
            results.append(f"- {clean(title)}\n  {href}\n  {snippet}")
        return "\n".join(results) if results else "Поиск не вернул результатов. Попробуйте изменить запрос."

    def _worker(self, prompt, mode, api_key, provider, config):
        try:
            web_context = ""
            if mode == "Поиск в интернете":
                self._status("Поиск в интернете…")
                web_context = "\n\nРЕЗУЛЬТАТЫ ПОИСКА:\n" + self._web_search(prompt)
            system = (
                "Ты — Помощник AI, русскоязычный помощник по написанию текстов и работе с документами. "
                "Пиши понятно, структурированно и профессионально. Не выдумывай источники и факты. "
                "Отделяй подтверждённую информацию от предположений."
            )
            if mode == "Написать текст":
                system += " Подготовь готовый текст с учётом цели, аудитории, формата и стиля из запроса."
            elif mode == "Анализ документов":
                system += " Проанализируй предоставленные документы, указывай названия файлов, находи выводы и расхождения. Не додумывай отсутствующие данные."
            if self.document_text:
                system += "\n\nТЕКСТ ДОКУМЕНТОВ:\n" + self.document_text[:MAX_CONTEXT_CHARS]
                if len(self.document_text) > MAX_CONTEXT_CHARS:
                    system += "\n[Контекст ограничен: не весь текст был передан модели.]"
            if web_context:
                system += web_context + "\nУкажи, что результаты получены поиском, и не считай сниппеты полной проверкой источников."
            self._status("Обращение к ИИ…")
            client = OpenAI(api_key=api_key, base_url=config["url"], timeout=60, max_retries=1)
            response = client.chat.completions.create(
                model=config["model"],
                messages=[{"role": "system", "content": system},
                          {"role": "user", "content": prompt}],
                temperature=0.3)
            answer = response.choices[0].message.content or "Модель вернула пустой ответ."
            self.after(0, self._append, "Помощник AI · " + provider, answer)
            self._status("Готово")
        except urllib.error.URLError as exc:
            self.after(0, self._append, "Ошибка сети", f"Не удалось выполнить сетевой запрос: {exc}")
            self._status("Ошибка сети")
        except Exception as exc:
            self.after(0, self._append, "Ошибка", f"{exc}\n\nПроверьте API-ключ, имя модели, лимиты провайдера и доступ к интернету.")
            self._status("Ошибка запроса")
        finally:
            self.busy = False
            self.after(0, lambda: self.send_btn.configure(state="normal", text="Отправить  ➤"))


if __name__ == "__main__":
    app = AssistantApp()
    app.mainloop()
