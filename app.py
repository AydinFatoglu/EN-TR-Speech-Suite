import ctypes
import json
import queue
import re
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import pyperclip
import sounddevice as sd
import vosk
import win32com.client


# ================================================================
# AYARLAR
# ================================================================

# Onefile: modeller EXE'nin yanında; kütüphaneler PyInstaller paketindedir.
# Çalışma dizini veya kısayolun "Başlangıç yeri" model yolunu değiştirmez.
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    try:
        BASE_DIR = Path(__file__).resolve().parent
    except NameError:
        # Thonny: editör içeriğini __file__ olmadan çalıştırma desteği.
        BASE_DIR = Path.cwd()

MODEL_TR_PATH = BASE_DIR / "model" / "tr"
MODEL_EN_PATH = BASE_DIR / "model" / "en"
SAMPLE_RATE = 16000

# Kullanıcının mevcut SAPI ses sıralaması korunmuştur.
ENGLISH_VOICE_INDEX = 0  # Microsoft David Desktop
TURKISH_VOICE_INDEX = 3  # Microsoft Tolga

SVS_FLAGS_ASYNC = 1
SVSFP_PURGE_BEFORE_SPEAK = 2
ASYNC_PURGE = SVS_FLAGS_ASYNC | SVSFP_PURGE_BEFORE_SPEAK

VK_CONTROL = 0x11
VK_V = 0x56
KEYEVENTF_KEYUP = 0x0002

TRIGGERS_TO_EN = {"ingilizce", "english", "ingiliz", "ingilizceye"}
TRIGGERS_TO_TR = {"turkce", "turkish", "turk", "turkceye"}


def clean_text(text):
    text = text.lower()
    replacements = {
        "i̇": "i",
        "ı": "i",
        "ç": "c",
        "ğ": "g",
        "ö": "o",
        "ş": "s",
        "ü": "u",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return re.sub(r"[^a-z\s]", "", text).strip()


class SpeechApp:
    def __init__(self, root):
        self.root = root
        self.running = True

        # Ortak dil: hem SAPI sesi hem de Vosk modeli bunu kullanır.
        self.current_lang = "TR"
        self.selected_voice_index = TURKISH_VOICE_INDEX

        # TTS durumu
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.tts_active = False

        # Vosk / mikrofon durumu
        self.models_ready = False
        self.listening_enabled = False
        self.recognizers = {}
        self.audio_queue = queue.Queue(maxsize=50)
        self.ui_queue = queue.Queue()
        self.reset_recognizers = threading.Event()
        self.stop_event = threading.Event()

        # Pano koruması
        self.monitor_enabled = True
        self.clipboard_guard_until = 0.0
        try:
            self.last_clipboard_text = pyperclip.paste().strip()
        except Exception:
            self.last_clipboard_text = ""

        self.build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.root.after(100, self.process_ui_queue)
        self.root.after(250, self.check_speaking_status)
        self.root.after(500, self.monitor_clipboard)

        threading.Thread(
            target=self.load_models_and_run_audio,
            name="VoskWorker",
            daemon=True,
        ).start()

    # ============================================================
    # ARAYÜZ
    # ============================================================

    def build_ui(self):
        self.root.title("Konuş ve Dinle")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.configure(bg="#F0F2F5")
        self.center_window(430, 430)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TCheckbutton", background="#F0F2F5", font=("Segoe UI", 10))
        style.configure("TScale", background="#F0F2F5")

        tk.Label(
            self.root,
            text="Okuma ve Sesle Yazma",
            bg="#F0F2F5",
            fg="#222222",
            font=("Segoe UI", 15, "bold"),
        ).pack(pady=(16, 6))

        tk.Label(
            self.root,
            text="Dil hem konuşma hem dinleme için birlikte değişir",
            bg="#F0F2F5",
            fg="#666666",
            font=("Segoe UI", 9),
        ).pack()

        lang_frame = tk.Frame(self.root, bg="#F0F2F5")
        lang_frame.pack(pady=12)

        self.tr_btn = tk.Button(
            lang_frame,
            text="TR  TÜRKÇE",
            command=lambda: self.set_language("TR"),
            width=14,
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            cursor="hand2",
        )
        self.tr_btn.pack(side="left", padx=5)

        self.en_btn = tk.Button(
            lang_frame,
            text="EN  ENGLISH",
            command=lambda: self.set_language("EN"),
            width=14,
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            cursor="hand2",
        )
        self.en_btn.pack(side="left", padx=5)

        self.lang_label = tk.Label(
            self.root,
            text="Geçerli Dil: TÜRKÇE (Tolga + Vosk TR)",
            bg="#F0F2F5",
            font=("Segoe UI", 10, "bold"),
        )
        self.lang_label.pack(pady=(0, 8))

        self.speed_display = tk.Label(
            self.root,
            text="Okuma Hızı: %140",
            bg="#F0F2F5",
            font=("Segoe UI", 10),
        )
        self.speed_display.pack(pady=(4, 2))

        self.speed_var = tk.IntVar(value=140)
        ttk.Scale(
            self.root,
            from_=50,
            to=200,
            variable=self.speed_var,
            orient="horizontal",
            command=self.update_speed_label,
            length=310,
        ).pack(pady=4)

        self.monitor_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            self.root,
            text="Panoyu otomatik oku",
            variable=self.monitor_var,
            command=self.toggle_monitor,
        ).pack(pady=8)

        button_frame = tk.Frame(self.root, bg="#F0F2F5")
        button_frame.pack(pady=8)

        self.speak_btn = tk.Button(
            button_frame,
            text="▶ SESLİ OKU",
            command=self.speak_clipboard,
            font=("Segoe UI", 11, "bold"),
            bg="#4CAF50",
            fg="white",
            activebackground="#45A049",
            activeforeground="white",
            width=12,
            height=2,
            relief="flat",
            cursor="hand2",
        )
        self.speak_btn.pack(side="left", padx=5)

        self.listen_btn = tk.Button(
            button_frame,
            text="● SESLİ YAZ",
            command=self.toggle_listening,
            font=("Segoe UI", 11, "bold"),
            bg="#1976D2",
            fg="white",
            activebackground="#1565C0",
            activeforeground="white",
            width=12,
            height=2,
            relief="flat",
            cursor="hand2",
            state="disabled",
        )
        self.listen_btn.pack(side="left", padx=5)

        self.stop_btn = tk.Button(
            button_frame,
            text="■ DURDUR",
            command=self.stop_audio,
            font=("Segoe UI", 11, "bold"),
            bg="#F44336",
            fg="white",
            activebackground="#E53935",
            activeforeground="white",
            width=12,
            height=2,
            relief="flat",
            cursor="hand2",
        )
        self.stop_btn.pack(side="left", padx=5)

        self.listen_info = tk.Label(
            self.root,
            text="Vosk modelleri yükleniyor...",
            bg="#F0F2F5",
            fg="#555555",
            font=("Segoe UI", 9),
        )
        self.listen_info.pack(pady=(5, 0))

        self.status_label = tk.Label(
            self.root,
            text="Hazırlanıyor",
            bg="#E3E7EB",
            fg="#333333",
            font=("Segoe UI", 9, "italic"),
            anchor="w",
            padx=10,
        )
        self.status_label.pack(side="bottom", fill="x", ipady=6)

        self.refresh_language_ui()

    def center_window(self, width, height):
        x = int((self.root.winfo_screenwidth() - width) / 2)
        y = int((self.root.winfo_screenheight() - height) / 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def set_status(self, text):
        if self.running:
            self.status_label.config(text=text)

    def update_speed_label(self, value):
        self.speed_display.config(text=f"Okuma Hızı: %{int(float(value))}")

    # ============================================================
    # DİL YÖNETİMİ
    # ============================================================

    def set_language(self, language):
        if language not in {"TR", "EN"}:
            return

        self.stop_audio(show_status=False)
        self.current_lang = language
        self.selected_voice_index = (
            TURKISH_VOICE_INDEX if language == "TR" else ENGLISH_VOICE_INDEX
        )
        self.clear_audio_queue()
        self.reset_recognizers.set()
        self.refresh_language_ui()

        language_name = "Türkçe" if language == "TR" else "İngilizce"
        if self.listening_enabled:
            self.set_status(f"Dinleniyor: {language_name}")
        else:
            self.set_status(f"Dil değiştirildi: {language_name}")

    def refresh_language_ui(self):
        if self.current_lang == "TR":
            self.tr_btn.config(bg="#D32F2F", fg="white")
            self.en_btn.config(bg="#DDE3E8", fg="#222222")
            self.lang_label.config(text="Geçerli Dil: TÜRKÇE (Tolga + Vosk TR)")
        else:
            self.en_btn.config(bg="#1976D2", fg="white")
            self.tr_btn.config(bg="#DDE3E8", fg="#222222")
            self.lang_label.config(text="Current Language: ENGLISH (David + Vosk EN)")

    # ============================================================
    # METİNDEN SESE: WINDOWS SAPI
    # ============================================================

    def speak_clipboard(self):
        try:
            text = pyperclip.paste().strip()
        except Exception:
            text = ""

        if not text:
            messagebox.showwarning("Uyarı", "Panoda okunacak metin yok.")
            return

        self.start_reading(text)

    def start_reading(self, text):
        text = text.strip()
        if not text:
            return

        # Hoparlör sesi Vosk tarafından tekrar yazılmasın.
        self.tts_active = True
        self.clear_audio_queue()
        self.reset_recognizers.set()

        try:
            self.speaker.Voice = self.speaker.GetVoices().Item(
                self.selected_voice_index
            )
        except Exception as exc:
            print(f"Ses seçimi hatası: {exc}")

        rate = int((self.speed_var.get() - 100) / 10)
        self.speaker.Rate = max(-10, min(10, rate))

        self.set_status(f"Okunuyor... (Hız: %{self.speed_var.get()})")
        self.speak_btn.config(state="disabled")
        self.speaker.Speak(text, ASYNC_PURGE)

    def check_speaking_status(self):
        if not self.running:
            return

        try:
            if self.tts_active and self.speaker.Status.RunningState == 1:
                self.tts_active = False
                self.clear_audio_queue()
                self.reset_recognizers.set()
                self.speak_btn.config(state="normal")
                self.show_idle_status()
        except Exception as exc:
            print(f"SAPI durum hatası: {exc}")

        self.root.after(250, self.check_speaking_status)

    def stop_audio(self, show_status=True):
        try:
            self.speaker.Speak("", ASYNC_PURGE)
        except Exception:
            pass

        self.tts_active = False
        self.clear_audio_queue()
        self.reset_recognizers.set()
        self.speak_btn.config(state="normal")
        if show_status:
            self.show_idle_status()

    # ============================================================
    # SESTEN METNE: VOSK
    # ============================================================

    def load_models_and_run_audio(self):
        try:
            missing_models = [
                str(path) for path in (MODEL_TR_PATH, MODEL_EN_PATH)
                if not path.is_dir()
            ]
            if missing_models:
                raise FileNotFoundError(
                    "Model klasörü bulunamadı:\n"
                    + "\n".join(missing_models)
                    + "\n\nmodel/tr ve model/en klasörlerini READ.exe ile "
                    "aynı ana klasöre yerleştirin."
                )
            model_tr = vosk.Model(str(MODEL_TR_PATH))
            model_en = vosk.Model(str(MODEL_EN_PATH))
            self.recognizers = {
                "TR": vosk.KaldiRecognizer(model_tr, SAMPLE_RATE),
                "EN": vosk.KaldiRecognizer(model_en, SAMPLE_RATE),
            }
            self.models_ready = True
            self.ui_queue.put(("models_ready",))
        except Exception as exc:
            self.ui_queue.put(("fatal_error", f"Vosk modelleri yüklenemedi:\n{exc}"))
            return

        try:
            with sd.RawInputStream(
                samplerate=SAMPLE_RATE,
                blocksize=8000,
                device=None,
                dtype="int16",
                channels=1,
                callback=self.microphone_callback,
            ):
                while not self.stop_event.is_set():
                    try:
                        data = self.audio_queue.get(timeout=0.2)
                    except queue.Empty:
                        continue

                    if not self.listening_enabled or self.tts_active:
                        continue

                    if self.reset_recognizers.is_set():
                        for recognizer in self.recognizers.values():
                            recognizer.Reset()
                        self.reset_recognizers.clear()

                    language = self.current_lang
                    recognizer = self.recognizers[language]

                    if recognizer.AcceptWaveform(data):
                        result = json.loads(recognizer.Result())
                        text = result.get("text", "").strip()
                        if text:
                            self.ui_queue.put(("recognized", text, language))

        except Exception as exc:
            self.ui_queue.put(("audio_error", str(exc)))

    def microphone_callback(self, indata, frames, audio_time, status):
        if not self.running or not self.listening_enabled or self.tts_active:
            return

        try:
            self.audio_queue.put_nowait(bytes(indata))
        except queue.Full:
            # Eski ses parçalarını büyütmek yerine o bloğu atla.
            pass

    def toggle_listening(self):
        if not self.models_ready:
            messagebox.showinfo("Bilgi", "Vosk modelleri henüz yükleniyor.")
            return

        self.listening_enabled = not self.listening_enabled
        self.clear_audio_queue()
        self.reset_recognizers.set()

        if self.listening_enabled:
            self.listen_btn.config(text="■ BİTİR", bg="#EF6C00")
            self.listen_info.config(
                text="Dinleme açık — şimdi yazı alanına tıklayıp konuşun"
            )
        else:
            self.listen_btn.config(text="● SESLİ YAZ", bg="#1976D2")
            self.listen_info.config(text="Sesle yazma kapalı")

        self.show_idle_status()

    def handle_recognized_text(self, text, recognized_language):
        if not self.running or not self.listening_enabled or self.tts_active:
            return

        cleaned = clean_text(text)
        words = set(cleaned.split())

        if "kendini kapat" in cleaned or "kill the program" in cleaned:
            self.on_close()
            return

        if recognized_language == "TR" and words.intersection(TRIGGERS_TO_EN):
            self.set_language("EN")
            return

        if recognized_language == "EN" and words.intersection(TRIGGERS_TO_TR):
            self.set_language("TR")
            return

        self.paste_at_cursor(text)

    def clear_audio_queue(self):
        while True:
            try:
                self.audio_queue.get_nowait()
            except queue.Empty:
                break

    # ============================================================
    # PANO VE AKTİF İMLECE YAPIŞTIRMA
    # ============================================================

    def paste_at_cursor(self, text):
        # Her Vosk sonucu önceki sonuçla birleşmesin.
        payload = f" {text.strip()} "

        try:
            previous_text = pyperclip.paste()
        except Exception:
            previous_text = None

        # Pano okuyucusunun bu dahili değişikliği seslendirmesini engelle.
        self.clipboard_guard_until = time.monotonic() + 1.2

        try:
            pyperclip.copy(payload)
            self.last_clipboard_text = payload.strip()
            internal_sequence = ctypes.windll.user32.GetClipboardSequenceNumber()

            ctypes.windll.user32.keybd_event(VK_CONTROL, 0, 0, 0)
            ctypes.windll.user32.keybd_event(VK_V, 0, 0, 0)
            ctypes.windll.user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)
            ctypes.windll.user32.keybd_event(
                VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
            )

            self.root.after(
                150,
                lambda: self.restore_clipboard(previous_text, internal_sequence),
            )
        except Exception as exc:
            self.set_status(f"Yapıştırma hatası: {exc}")

    def restore_clipboard(self, previous_text, internal_sequence):
        try:
            current_sequence = ctypes.windll.user32.GetClipboardSequenceNumber()

            # Arada kullanıcı yeni bir şey kopyaladıysa onun panosuna dokunma.
            if current_sequence != internal_sequence:
                try:
                    self.last_clipboard_text = pyperclip.paste().strip()
                except Exception:
                    pass
                return

            if previous_text is not None:
                pyperclip.copy(previous_text)
                self.last_clipboard_text = previous_text.strip()
        except Exception:
            pass

    def toggle_monitor(self):
        self.monitor_enabled = self.monitor_var.get()
        self.set_status(
            "Pano izleme açık" if self.monitor_enabled else "Pano izleme kapalı"
        )

    def monitor_clipboard(self):
        if not self.running:
            return

        try:
            current_text = pyperclip.paste().strip()
        except Exception:
            current_text = ""

        # Vosk'un dahili Ctrl+V işlemi okuma başlatmaz.
        if time.monotonic() < self.clipboard_guard_until:
            self.last_clipboard_text = current_text
        elif (
            self.monitor_enabled
            and current_text
            and current_text != self.last_clipboard_text
        ):
            self.last_clipboard_text = current_text
            self.start_reading(current_text)

        self.root.after(500, self.monitor_clipboard)

    # ============================================================
    # THREAD -> TKINTER OLAYLARI
    # ============================================================

    def process_ui_queue(self):
        if not self.running:
            return

        while True:
            try:
                event = self.ui_queue.get_nowait()
            except queue.Empty:
                break

            event_name = event[0]

            if event_name == "models_ready":
                self.listen_btn.config(state="normal")
                self.listen_info.config(text="Sesle yazma hazır")
                self.show_idle_status()

            elif event_name == "recognized":
                self.handle_recognized_text(event[1], event[2])

            elif event_name == "audio_error":
                self.listening_enabled = False
                self.listen_btn.config(state="disabled")
                self.listen_info.config(text="Mikrofon açılamadı")
                messagebox.showerror("Mikrofon Hatası", event[1])

            elif event_name == "fatal_error":
                self.listen_btn.config(state="disabled")
                self.listen_info.config(text="Vosk modelleri yüklenemedi")
                messagebox.showerror("Model Hatası", event[1])

        self.root.after(100, self.process_ui_queue)

    def show_idle_status(self):
        if self.tts_active:
            return

        if self.listening_enabled:
            language_name = "Türkçe" if self.current_lang == "TR" else "İngilizce"
            self.set_status(f"Dinleniyor: {language_name}")
        elif self.models_ready:
            self.set_status("Hazır")
        else:
            self.set_status("Vosk modelleri yükleniyor...")

    # ============================================================
    # KAPATMA
    # ============================================================

    def on_close(self):
        if not self.running:
            return

        self.running = False
        self.listening_enabled = False
        self.stop_event.set()
        try:
            self.speaker.Speak("", ASYNC_PURGE)
        except Exception:
            pass
        self.root.destroy()


def main():
    root = tk.Tk()
    SpeechApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

