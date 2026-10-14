import os
import threading
import json
import customtkinter as ctk
from tkinter import filedialog, messagebox
import PyPDF2
import docx
from openai import OpenAI

# Настройка темы
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

CONFIG_FILE = "aichat_config.json"

class AichatStyleAssistant(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Aichatting — Document Assistant")
        self.geometry("1050x700")
        self.minsize(900, 600)
        
        self.document_text = ""
        self.load_config()

        # Главный контейнер (сетка из двух колонок: слева меню, справа чат)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ==================== ЛЕВАЯ ПАНЕЛЬ (Sidebar) ====================
        self.sidebar_frame = ctk.CTkFrame(self, width=280, corner_radius=0, fg_color=("#1e293b", "#0f172a"))
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1)

        # Логотип / Название
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="💬 Aichatting", font=("Segoe UI Bold", 20), text_color="#38bdf8")
        self.logo_label.grid(row=0, column=0, padx=20, pady=(25, 15), sticky="w")

        # Выбор провайдера ИИ
        self.prov_title = ctk.CTkLabel(self.sidebar_frame, text="Модель ИИ:", font=("Segoe UI", 12, "bold"), text_color="#94a3b8")
        self.prov_title.grid(row=1, column=0, padx=20, sticky="w")

        self.providers = {
            "Google Gemini (Flash)": {
                "url": "https://generativelanguage.googleapis.com/v1beta/openai/",
                "model": "gemini-2.0-flash"
            },
            "Groq (Llama 3.3)": {
                "url": "https://api.groq.com/openai/v1",
                "model": "llama-3.3-70b-versatile"
            }
        }
        self.provider_menu = ctk.CTkOptionMenu(
            self.sidebar_frame, 
            values=list(self.providers.keys()), 
            command=self.on_provider_change,
            fg_color="#334155", button_color="#475569", button_hover_color="#64748b"
        )
        self.provider_menu.grid(row=2, column=0, padx=20, pady=(5, 15), sticky="ew")
        if self.saved_provider in self.providers:
            self.provider_menu.set(self.saved_provider)

        # Поле API Ключа
        self.api_title = ctk.CTkLabel(self.sidebar_frame, text="API Key:", font=("Segoe UI", 12, "bold"), text_color="#94a3b8")
        self.api_title.grid(row=3, column=0, padx=20, sticky="w")

        self.api_entry = ctk.CTkEntry(self.sidebar_frame, show="*", placeholder_text="Вставьте ключ...", fg_color="#334155", border_color="#475569")
        self.api_entry.grid(row=4, column=0, padx=20, pady=(5, 15), sticky="ew")
        if self.saved_api_key:
            self.api_entry.insert(0, self.saved_api_key)

        # Блок управления документом (карточка как на скрипте Recent Chat)
        self.doc_card = ctk.CTkFrame(self.sidebar_frame, fg_color=("#334155", "#1e293b"), corner_radius=12)
        self.doc_card.grid(row=5, column=0, padx=15, pady=10, sticky="new")
        self.doc_card.grid_columnconfigure(0, weight=1)

        self.doc_header = ctk.CTkLabel(self.doc_card, text="📁 Документ для анализа", font=("Segoe UI", 13, "bold"), text_color="#f1f5f9")
        self.doc_header.grid(row=0, column=0, padx=15, pady=(12, 5), sticky="w")

        self.doc_status_label = ctk.CTkLabel(self.doc_card, text="Файл не выбран", font=("Segoe UI", 11), text_color="#94a3b8", wraplength=220, justify="left")
        self.doc_status_label.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="w")

        self.upload_btn = ctk.CTkButton(self.doc_card, text="Загрузить файл", command=self.load_document, fg_color="#0ea5e9", hover_color="#0284c7", height=32)
        self.upload_btn.grid(row=2, column=0, padx=15, pady=(0, 15), sticky="ew")

        # Кнопка сохранения настроек внизу сайдбара
        self.save_btn = ctk.CTkButton(self.sidebar_frame, text="Сохранить настройки", command=self.save_config, fg_color="transparent", border_color="#475569", border_width=1, hover_color="#334155", text_color="#cbd5e1")
        self.save_btn.grid(row=6, column=0, padx=20, pady=20, sticky="ew")


        # ==================== ПРАВАЯ ПАНЕЛЬ (Чат область) ====================
        self.chat_main_frame = ctk.CTkFrame(self, fg_color=("#f8fafc", "#0b0f19"), corner_radius=0)
        self.chat_main_frame.grid(row=0, column=1, sticky="nsew")
        self.chat_main_frame.grid_columnconfigure(0, weight=1)
        self.chat_main_frame.grid_rowconfigure(1, weight=1)

        # Верхняя панель чата (Заголовок как на Aichatting)
        self.chat_top_bar = ctk.CTkFrame(self.chat_main_frame, height=60, fg_color=("#ffffff", "#111827"), corner_radius=0)
        self.chat_top_bar.grid(row=0, column=0, sticky="ew")
        
        self.chat_title = ctk.CTkLabel(self.chat_top_bar, text="🤖 ChatBot (Документ активен)", font=("Segoe UI Bold", 15), text_color=("#0f172a", "#f8fafc"))
        self.chat_title.pack(side="left", padx=25, pady=15)

        # Область сообщений (Чат-бокс)
        self.chat_box = ctk.CTkTextbox(self.chat_main_frame, state="disabled", wrap="word", font=("Segoe UI", 13), fg_color=("#f8fafc", "#0b0f19"), text_color=("#1e293b", "#f1f5f9"))
        self.chat_box.grid(row=1, column=0, padx=20, pady=15, sticky="nsew")

        # Нижняя панель ввода (Плавающее поле ввода как на референсе с кнопкой отправки)
        self.input_container = ctk.CTkFrame(self.chat_main_frame, fg_color="transparent", height=80)
        self.input_container.grid(row=2, column=0, padx=25, pady=(0, 20), sticky="ew")
        self.input_container.grid_columnconfigure(0, weight=1)

        self.user_input = ctk.CTkEntry(
            self.input_container, 
            placeholder_text="Enter your message...", 
            height=50, 
            corner_radius=25,
            font=("Segoe UI", 13),
            fg_color=("#ffffff", "#1e293b"),
            border_color=("#cbd5e1", "#334155"),
            border_width=1
        )
        self.user_input.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.user_input.bind("<Return>", lambda event: self.send_message())

        self.send_btn = ctk.CTkButton(
            self.input_container, 
            text="➤", 
            width=50, 
            height=50, 
            corner_radius=25,
            font=("Segoe UI Bold", 16),
            fg_color="#f97316", # Оранжевый акцент кнопки отправки как на референсе
            hover_color="#ea580c",
            command=self.send_message
        )
        self.send_btn.grid(row=0, column=1)

        # Приветственное сообщение
        self.append_to_chat("Система", "Добро пожаловать в Aichatting Assistant! Загрузите документ слева, введите API-ключ и задавайте вопросы.")

    def on_provider_change(self, choice):
        pass

    def load_config(self):
        self.saved_api_key = ""
        self.saved_provider = "Google Gemini (Flash)"
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.saved_api_key = data.get("api_key", "")
                    self.saved_provider = data.get("provider", "Google Gemini (Flash)")
            except:
                pass

    def save_config(self):
        data = {
            "api_key": self.api_entry.get().strip(),
            "provider": self.provider_menu.get()
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f)
            messagebox.showinfo("Успех", "Настройки сохранены локально.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить: {e}")

    def load_document(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Documents", "*.pdf *.docx *.txt"), ("All Files", "*.*")]
        )
        if not filepath:
            return

        filename = os.path.basename(filepath)
        self.doc_status_label.configure(text=f"Загрузка: {filename}...", text_color="#fbbf24")
        self.update()

        try:
            ext = os.path.splitext(filepath)[1].lower()
            text = ""
            if ext == ".pdf":
                with open(filepath, "rb") as f:
                    reader = PyPDF2.PdfReader(f)
                    for page in reader.pages:
                        extracted = page.extract_text()
                        if extracted:
                            text += extracted + "\n"
            elif ext == ".docx":
                doc = docx.Document(filepath)
                for para in doc.paragraphs:
                    text += para.text + "\n"
            elif ext == ".txt":
                with open(filepath, "r", encoding="utf-8") as f:
                    text = f.read()
            else:
                messagebox.showerror("Ошибка", "Формат не поддерживается.")
                self.doc_status_label.configure(text="Файл не выбран", text_color="#94a3b8")
                return

            self.document_text = text.strip()
            self.doc_status_label.configure(text=f"✓ {filename}", text_color="#4ade80")
            self.chat_title.configure(text=f"🤖 ChatBot ({filename})")
            self.append_to_chat("Система", f"Документ '{filename}' успешно загружен и готов к анализу.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка чтения файла:\n{e}")
            self.doc_status_label.configure(text="Ошибка загрузки", text_color="#f87171")

    def append_to_chat(self, sender, message):
        self.chat_box.configure(state="normal")
        self.chat_box.insert("end", f"[{sender}]\n{message}\n\n")
        self.chat_box.configure(state="disabled")
        self.chat_box.yview("end")

    def send_message(self):
        question = self.user_input.get().strip()
        api_key = self.api_entry.get().strip()
        
        selected_provider = self.provider_menu.get()
        prov_data = self.providers[selected_provider]

        if not api_key:
            messagebox.showwarning("Внимание", "Введите API Key!")
            return
        if not question:
            return

        self.user_input.delete(0, "end")
        self.append_to_chat("Вы", question)

        self.send_btn.configure(state="disabled", text="...")
        threading.Thread(target=self.ask_ai_masked, args=(prov_data["url"], prov_data["model"], api_key, question, selected_provider), daemon=True).start()

    def ask_ai_masked(self, base_url, model_name, api_key, question, provider_name):
        try:
            # 🛡️ Маскировка под стандартный браузер Google Chrome
            browser_headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
                "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
                "Referer": "https://www.google.com/",
                "Origin": "https://www.google.com",
                "Sec-Ch-Ua": '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
                "Sec-Ch-Ua-Mobile": "?0",
                "Sec-Ch-Ua-Platform": '"Windows"',
                "Sec-Fetch-Dest": "document",
                "Sec-Fetch-Mode": "navigate",
                "Sec-Fetch-Site": "cross-site"
            }

            client = OpenAI(
                api_key=api_key, 
                base_url=base_url,
                default_headers=browser_headers
            )

            system_prompt = "Ты — корпоративный помощник по работе с документами. Отвечай кратко и строго по делу на русском языке."
            if self.document_text:
                system_prompt += f"\n\nКонтекст документа:\n{self.document_text[:40000]}"

            response = client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": question}
                ],
                temperature=0.2,
            )

            answer = response.choices[0].message.content
            self.after(0, self.append_to_chat, f"AI ({provider_name.split()[0]})", answer)
            
        except Exception as e:
            self.after(0, self.append_to_chat, "Ошибка", f"Сбой защищенного соединения:\n{str(e)}")
            
        finally:
            self.after(0, lambda: self.send_btn.configure(state="normal", text="➤"))

if __name__ == "__main__":
    app = AichatStyleAssistant()
    app.mainloop()
