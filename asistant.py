import os
import threading
import json
import customtkinter as ctk
from tkinter import filedialog, messagebox
import PyPDF2
import docx
from openai import OpenAI

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

CONFIG_FILE = "config.json"

class EnterpriseDeepSeekAssistant(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Google Workspace Document Assistant (Secure)")
        self.geometry("900x700")
        
        self.document_text = ""
        self.load_config()

        # --- ВЕРХНЯЯ ПАНЕЛЬ (Настройки маскировки) ---
        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.pack(pady=10, padx=10, fill="x")

        # Шлюз (лучше использовать развертывание на Google Cloud / Firebase)
        self.proxy_label = ctk.CTkLabel(self.settings_frame, text="Шлюз (Google Cloud/Proxy):")
        self.proxy_label.pack(side="left", padx=5)
        
        self.proxy_entry = ctk.CTkEntry(self.settings_frame, width=240, placeholder_text="https://your-app.a.run.app")
        self.proxy_entry.pack(side="left", padx=5)
        if self.saved_proxy:
            self.proxy_entry.insert(0, self.saved_proxy)

        # API Ключ
        self.api_label = ctk.CTkLabel(self.settings_frame, text="API Key:")
        self.api_label.pack(side="left", padx=5)
        
        self.api_entry = ctk.CTkEntry(self.settings_frame, width=160, show="*")
        self.api_entry.pack(side="left", padx=5)
        if self.saved_api_key:
            self.api_entry.insert(0, self.saved_api_key)

        self.save_btn = ctk.CTkButton(self.settings_frame, text="Сохранить", width=80, command=self.save_config)
        self.save_btn.pack(side="right", padx=5)

        # --- ПАНЕЛЬ ДОКУМЕНТА ---
        self.doc_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.doc_frame.pack(pady=5, padx=10, fill="x")

        self.upload_btn = ctk.CTkButton(self.doc_frame, text="📁 Загрузить документ (PDF, DOCX, TXT)", command=self.load_document)
        self.upload_btn.pack(side="left")

        self.doc_status_label = ctk.CTkLabel(self.doc_frame, text="Документ не загружен", text_color="gray")
        self.doc_status_label.pack(side="left", padx=15)

        # --- ОКНО ЧАТА ---
        self.chat_box = ctk.CTkTextbox(self, state="disabled", wrap="word", font=("Segoe UI", 13))
        self.chat_box.pack(pady=10, padx=10, fill="both", expand=True)

        # --- НИЖНЯЯ ПАНЕЛЬ ---
        self.bottom_frame = ctk.CTkFrame(self)
        self.bottom_frame.pack(pady=10, padx=10, fill="x")

        self.user_input = ctk.CTkEntry(self.bottom_frame, placeholder_text="Задайте вопрос по документу...", height=45, font=("Segoe UI", 13))
        self.user_input.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.user_input.bind("<Return>", lambda event: self.send_message())

        self.send_btn = ctk.CTkButton(self.bottom_frame, text="Отправить", width=110, height=45, command=self.send_message)
        self.send_btn.pack(side="right")

        self.append_to_chat("Система", "Режим маскировки трафика активен. Запросы замаскированы под стандартный веб-серфинг Google Chrome.")

    def load_config(self):
        self.saved_api_key = ""
        self.saved_proxy = "https://api.deepseek.com"
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.saved_api_key = data.get("api_key", "")
                    self.saved_proxy = data.get("proxy_url", "https://api.deepseek.com")
            except:
                pass

    def save_config(self):
        data = {
            "api_key": self.api_entry.get().strip(),
            "proxy_url": self.proxy_entry.get().strip()
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f)
            messagebox.showinfo("Успех", "Настройки сохранены.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить: {e}")

    def load_document(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Documents", "*.pdf *.docx *.txt"), ("All Files", "*.*")]
        )
        if not filepath:
            return

        filename = os.path.basename(filepath)
        self.doc_status_label.configure(text=f"Чтение: {filename}...", text_color="orange")
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
                self.doc_status_label.configure(text="Документ не загружен", text_color="gray")
                return

            self.document_text = text.strip()
            self.doc_status_label.configure(text=f"Активен: {filename}", text_color="green")
            self.append_to_chat("Система", f"Документ '{filename}' загружен в оперативную память.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка чтения файла:\n{e}")
            self.doc_status_label.configure(text="Ошибка загрузки", text_color="red")

    def append_to_chat(self, sender, message):
        self.chat_box.configure(state="normal")
        self.chat_box.insert("end", f"[{sender}]\n{message}\n\n")
        self.chat_box.configure(state="disabled")
        self.chat_box.yview("end")

    def send_message(self):
        question = self.user_input.get().strip()
        api_key = self.api_entry.get().strip()
        proxy_url = self.proxy_entry.get().strip()

        if not api_key or not proxy_url:
            messagebox.showwarning("Внимание", "Заполните шлюз и API Key!")
            return
        if not question:
            return

        self.user_input.delete(0, "end")
        self.append_to_chat("Вы", question)

        self.send_btn.configure(state="disabled", text="Анализ...")
        threading.Thread(target=self.ask_ai_masked, args=(proxy_url, api_key, question), daemon=True).start()

    def ask_ai_masked(self, proxy_url, api_key, question):
        try:
            # 🛡️ Продвинутая маскировка под браузер Google Chrome (имитация обычного веб-серфинга)
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

            # Инициализация клиента с поддельными заголовками браузера
            client = OpenAI(
                api_key=api_key, 
                base_url=proxy_url,
                default_headers=browser_headers
            )

            system_prompt = "Ты — корпоративный помощник по работе с документами. Отвечай кратко и строго по делу на русском языке."
            if self.document_text:
                system_prompt += f"\n\nКонтекст документа:\n{self.document_text[:50000]}"

            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": question}
                ],
                temperature=0.2,
            )

            answer = response.choices[0].message.content
            self.after(0, self.append_to_chat, "DeepSeek AI", answer)
            
        except Exception as e:
            self.after(0, self.append_to_chat, "Ошибка", f"Сбой защищенного соединения:\n{str(e)}")
            
        finally:
            self.after(0, lambda: self.send_btn.configure(state="normal", text="Отправить"))

if __name__ == "__main__":
    app = EnterpriseDeepSeekAssistant()
    app.mainloop()