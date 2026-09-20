# -*- coding: utf-8 -*-
"""Джарвис — голосовой помощник для Windows (push-to-talk).

Запуск:    py -3 jarvis.py          (режим микрофона: зажми Ctrl+Windows и говори)
           py -3 jarvis.py -t       (режим ввода с клавиатуры)
"""
import ctypes
import datetime
import glob
import json
import os
import platform
import random
import re
import shutil
import socket
import string
import subprocess
import sys
import threading
import time
import urllib.parse
import urllib.request
import webbrowser
import winreg
import winsound

try:
    import win32com.client
except ImportError:
    win32com = None

try:
    import psutil
except ImportError:
    psutil = None

try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    from ctypes import cast, POINTER
except ImportError:
    AudioUtilities = None

try:
    import speech_recognition as sr
except ImportError:
    sr = None

try:
    import pyaudio
except ImportError:
    pyaudio = None

try:
    from pynput import keyboard
except ImportError:
    keyboard = None

TEXT_MODE = "-t" in sys.argv or "--text" in sys.argv

STOP_WORDS = r"\b(выключись|отключись|стоп|до свидания|до свиданья|пока|отбой|заверши работу)\b"

WEEKDAYS = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье"]

SAMPLE_RATE = 16000

APPS = {
    "калькулятор": "calc.exe",
    "блокнот": "notepad.exe",
    "проводник": "explorer.exe",
    "диспетчер задач": "taskmgr.exe",
    "диспетчер": "taskmgr.exe",
    "paint": "mspaint.exe",
    "wordpad": "write.exe",
    "панель управления": "control.exe",
    "командная строка": "cmd.exe",
    "терминал": "cmd.exe",
    "powershell": "powershell.exe",
}

APP_ALIASES = {
    "хром": "chrome", "гугл хром": "chrome", "google chrome": "chrome",
    "дискорд": "discord", "стим": "steam",
    "спотифай": "spotify", "телеграм": "telegram", "тг": "telegram",
    "фотошоп": "photoshop", "ворд": "word", "эксель": "excel",
    "паверпоинт": "powerpoint", "зум": "zoom",
    "параметры": "settings", "настройки": "settings",
    "вс код": "visual studio code", "вижуал студио": "visual studio",
    "роблокс": "roblox",
}

SITES = {
    "youtube": "https://www.youtube.com",
    "ютуб": "https://www.youtube.com",
    "вк": "https://vk.com",
    "вконтакте": "https://vk.com",
    "mail": "https://mail.ru",
    "почта": "https://mail.ru",
    "mail.ru": "https://mail.ru",
    "github": "https://github.com",
    "гитхаб": "https://github.com",
    "яндекс": "https://ya.ru",
    "гугл": "https://www.google.com",
    "википедия": "https://ru.wikipedia.org",
    "хабр": "https://habr.com",
    "музыка": "https://music.youtube.com",
    "госуслуги": "https://www.gosuslugi.ru",
    "вотсап": "https://web.whatsapp.com",
    "whatsapp": "https://web.whatsapp.com",
    "веб телеграм": "https://web.telegram.org",
    "мессенджер": "https://web.telegram.org",
}

START_MENU_DIRS = [
    os.path.join(os.environ.get("ProgramData", r"C:\ProgramData"),
                 r"Microsoft\Windows\Start Menu\Programs"),
    os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs"),
]

JOKES = [
    "Вчера пытался открыть холодильник с помощью голоса. Попросил считать одним из своих программистов.",
    "Почему программисты путают Хеллоуин и Рождество? Потому что 25 декабря равно 31 окт, а в восьмеричной системе всё наоборот.",
    "Два байта идут по улице. Один говорит: — Слышь, а чё ты грустный? — Да ёмкость в 128 мегабайт, все время одна и та же история...",
    "Захожу в лифт, говорю: «в триста тридцать шестую». Лифт повёз молча. Обиделся, наверное.",
]

WHETHER_CODES = {
    0: "ясно", 1: "преимущественно ясно", 2: "переменная облачность", 3: "пасмурно",
    45: "туман", 48: "изморозь", 51: "мелкая морось", 53: "морось", 55: "сильная морось",
    61: "небольшой дождь", 63: "дождь", 65: "сильный дождь",
    71: "небольшой снег", 73: "снег", 75: "сильный снег", 77: "ледяная крупа",
    80: "небольшой ливень", 81: "ливень", 82: "сильный ливень",
    95: "гроза", 96: "гроза с градом", 99: "сильная гроза с градом",
}


def _after(text, key):
    idx = text.find(key)
    return text[idx + len(key):].strip(" ,.!?;:")


def _clean_search(text, keys):
    for k in keys:
        if k in text:
            return _after(text, k)
    return ""


NUMBERS = {
    "ноль": 0, "один": 1, "одна": 1, "два": 2, "две": 2, "три": 3,
    "четыре": 4, "пять": 5, "шесть": 6, "семь": 7, "восемь": 8, "девять": 9,
    "десять": 10, "одиннадцать": 11, "двенадцать": 12, "тринадцать": 13,
    "четырнадцать": 14, "пятнадцать": 15, "шестнадцать": 16,
    "семнадцать": 17, "восемнадцать": 18, "девятнадцать": 19,
    "двадцать": 20, "тридцать": 30, "сорок": 40, "пятьдесят": 50,
    "шестьдесят": 60, "семьдесят": 70, "восемьдесят": 80, "девяносто": 90,
    "сто": 100, "двести": 200, "триста": 300, "четыреста": 400,
    "пятьсот": 500, "шестьсот": 600, "семьсот": 700, "восемьсот": 800,
    "девятьсот": 900, "тысяча": 1000, "тысячи": 1000, "миллион": 1000000, "миллиона": 1000000,
    "ста": 100, "двухсот": 200, "трехсот": 300, "трёхсот": 300,
    "четырехсот": 400, "четырёхсот": 400, "пятисот": 500, "шестисот": 600,
    "семисот": 700, "восьмисот": 800, "девятисот": 900,
    "двадцати": 20, "тридцати": 30, "пятидесяти": 50, "шестидесяти": 60,
    "семидесяти": 70, "восьмидесяти": 80,
    "десяти": 10, "одиннадцати": 11, "двенадцати": 12, "тринадцати": 13,
    "четырнадцати": 14, "пятнадцати": 15, "шестнадцати": 16,
    "семнадцати": 17, "восемнадцати": 18, "девятнадцати": 19,
}


def words_to_digits(text):
    digits = []
    total = None
    for w in text.split():
        if w in NUMBERS:
            total = NUMBERS[w] if total is None else total + NUMBERS[w]
        else:
            if total is not None:
                digits.append(str(total))
                total = None
            digits.append(w)
    if total is not None:
        digits.append(str(total))
    return " ".join(digits)


def words_to_int(text):
    total = None
    for w in text.split():
        if w in NUMBERS:
            total = NUMBERS[w] if total is None else total + NUMBERS[w]
    if total is None:
        nums = re.findall(r"\d+", text)
        total = int(nums[-1]) if nums else None
    return total


def _ru_plural(n, one, few, many):
    n10 = n % 10
    n100 = n % 100
    if n10 == 1 and n100 != 11:
        return one
    if 2 <= n10 <= 4 and not (12 <= n100 <= 14):
        return few
    return many


def safe_url(url, timeout=12):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def beep(up):
    try:
        if up:
            winsound.Beep(880, 110)
            winsound.Beep(1320, 110)
        else:
            winsound.Beep(660, 110)
    except Exception:
        pass


def normalize(name):
    """Нормализация имени: нижний регистр без пробелов и символов."""
    return re.sub(r"[^a-z0-9а-яё]", "", name.lower())


def resolve(host):
    """Проверка, существует ли домен (DNS-запрос с коротким таймаутом)."""
    try:
        socket.setdefaulttimeout(2)
        socket.gethostbyname(host)
        return True
    except Exception:
        return False


NEURAL_VOICES = ["ru-RU-DmitryNeural", "ru-RU-SvetlanaNeural"]
NEURAL_RATE = "-5%"


class Speech:
    """Озвучка ответов: нейросетевой голос edge-tts (интернет) + SAPI как запасной."""

    def __init__(self):
        self.busy = False
        self._sapi = None
        self._edge = None
        self._voice_i = 0
        try:
            self._sapi = win32com.client.Dispatch("SAPI.SpVoice")
            for v in self._sapi.GetVoices():
                desc = v.GetDescription().lower()
                if "ru" in desc or "russian" in desc:
                    self._sapi.Voice = v
                    break
        except Exception:
            self._sapi = None
        try:
            import edge_tts
            self._edge = edge_tts
        except ImportError:
            self._edge = None
            print("edge-tts недоступен. Установите: py -3 -m pip install edge-tts")

    @property
    def voice_names(self):
        return ["Дмитрий", "Светлана"]

    def voice_label(self):
        return self.voice_names[self._voice_i]

    def switch_voice(self, idx=None):
        if not self._edge:
            return
        if idx is None:
            self._voice_i = (self._voice_i + 1) % len(NEURAL_VOICES)
        else:
            self._voice_i = idx % len(NEURAL_VOICES)

    async def _synth(self, text, path):
        await self._edge.Communicate(text, NEURAL_VOICES[self._voice_i],
                                     rate=NEURAL_RATE).save(path)

    def _play_mp3(self, path):
        winmm = ctypes.windll.winmm
        alias = "jarvisvoice"
        winmm.mciSendStringW("close %s" % alias, None, 0, 0)
        if winmm.mciSendStringW('open "%s" type mpegvideo alias %s' % (path, alias), None, 0, 0) != 0:
            return False
        ok = winmm.mciSendStringW("play %s wait" % alias, None, 0, 0) == 0
        winmm.mciSendStringW("close %s" % alias, None, 0, 0)
        return ok

    def _speak_neural(self, text):
        import asyncio
        import tempfile
        path = os.path.join(tempfile.gettempdir(), "jarvis_tts.mp3")
        try:
            asyncio.run(self._synth(text, path))
        except Exception as e:
            print("edge-tts ошибка:", e)
            return False
        if not os.path.exists(path) or os.path.getsize(path) < 1024:
            return False
        try:
            return self._play_mp3(path)
        except Exception as e:
            print("Плеер mp3 ошибка:", e)
            return False

    def _speak_sapi(self, text):
        if self._sapi is not None:
            try:
                self._sapi.Speak(text)
                return True
            except Exception:
                return False
        return False

    def say(self, text):
        print("Джарвис:", text)
        while self.busy:
            threading.Event().wait(0.2)
        self.busy = True
        try:
            if self._edge is not None:
                if self._speak_neural(text):
                    return
            self._speak_sapi(text)
        finally:
            self.busy = False


class PushToTalk:
    """Глобальный хоткей: слушать, пока зажаты Ctrl+Windows."""

    CTRL_KEYS = (keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r) if keyboard else ()
    WIN_KEYS = (keyboard.Key.cmd, keyboard.Key.cmd_r) if keyboard else ()

    def __init__(self):
        self.go = threading.Event()
        self.stop = threading.Event()
        self._pressed = set()
        self._was_active = False
        self._listener = None

    @staticmethod
    def _is_mod(k, group):
        for ref in group:
            if k == ref:
                return True
        return False

    def _on_press(self, key):
        try:
            self._pressed.add(key)
        except TypeError:
            pass
        active = self._active()
        if active and not self._was_active:
            self.go.set()
        self._was_active = active

    def _on_release(self, key):
        try:
            self._pressed.discard(key)
        except TypeError:
            pass
        active = self._active()
        if self._was_active and not active:
            self.stop.set()
        self._was_active = active

    def _active(self):
        ctrl = any(self._is_mod(k, self.CTRL_KEYS) for k in self._pressed)
        win = any(self._is_mod(k, self.WIN_KEYS) for k in self._pressed)
        return ctrl and win

    def start(self):
        if keyboard is None:
            return False
        self._listener = keyboard.Listener(on_press=self._on_press, on_release=self._on_release)
        self._listener.start()
        return True

    def wait_for_push(self):
        self.go.wait()
        self.go.clear()

    def wait_for_release(self):
        self.stop.wait()
        self.stop.clear()

    def stop_listener(self):
        if self._listener:
            self._listener.stop()


class Recorder:
    """Непрерывная запись с микрофона (не зависит от энергетического порога)."""

    def __init__(self):
        self._frames = []
        self._running = False
        self._thread = None
        self._stream = None
        self._p = None

    def start(self):
        if pyaudio is None:
            return False
        self._p = pyaudio.PyAudio()
        self._frames = []
        self._stream = self._p.open(format=pyaudio.paInt16, channels=1,
                                    rate=SAMPLE_RATE, input=True,
                                    frames_per_buffer=1024)
        self._running = True
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()
        return True

    def _loop(self):
        while self._running:
            try:
                data = self._stream.read(1024, exception_on_overflow=False)
                self._frames.append(data)
            except Exception:
                break

    def stop(self):
        self._running = False
        if self._thread:
            self._thread.join(timeout=2)
        try:
            self._stream.stop_stream()
            self._stream.close()
        except Exception:
            pass
        try:
            self._p.terminate()
        except Exception:
            pass

    def audio(self):
        if not self._frames or sr is None:
            return None
        return sr.AudioData(b"".join(self._frames), SAMPLE_RATE, 2)


def find_app(name):
    """Поиск установленного приложения по имени в ярлыках Пуска, реестре и PATH."""
    target = normalize(APP_ALIASES.get(name, name))
    if not target:
        return None

    exe = shutil.which(name) or shutil.which(name + ".exe")
    if exe:
        return exe

    best = None
    for base in START_MENU_DIRS:
        if not os.path.isdir(base):
            continue
        for lnk in glob.glob(base + r"\**\*.lnk", recursive=True):
            stem = normalize(os.path.splitext(os.path.basename(lnk))[0])
            if not stem:
                continue
            if stem == target:
                return lnk
            if best is None and (target in stem or stem in target):
                best = lnk
    if best:
        return best

    uninstall = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            key = winreg.OpenKey(root, uninstall)
            nkeys = winreg.QueryInfoKey(key)[0]
            for i in range(nkeys):
                try:
                    sub = winreg.EnumKey(key, i)
                except OSError:
                    break
                try:
                    sub_key = winreg.OpenKey(key, sub)
                    display = winreg.QueryValueEx(sub_key, "DisplayName")[0]
                except OSError:
                    continue
                if normalize(display) == target:
                    for value in ("DisplayIcon", "InstallLocation"):
                        try:
                            val = winreg.QueryValueEx(sub_key, value)[0]
                        except OSError:
                            continue
                        if val and val not in ("1", display):
                            if os.path.isdir(val):
                                val = os.path.join(val, str(display).replace("  ", " ") + ".exe")
                            if os.path.isfile(val) or val.lower().endswith(".exe"):
                                return val
            winreg.CloseKey(key)
        except OSError:
            continue
    return None


NOTES_FILE = os.path.join(os.path.expanduser("~"), "jarvis_notes.txt")

SPECIAL_FOLDERS = {
    "загрузки": "Downloads", "документы": "Documents",
    "картинки": "Pictures", "изображения": "Pictures",
    "музыку": "Music", "музыка": "Music",
    "видео": "Videos", "рабочий стол": "Desktop", "рабочий стол": "Desktop",
    "рабочего стола": "Desktop", "загрузок": "Downloads",
}

POWER_ACTIONS = {
    "выключи компьютер": ("s", 20), "выключить компьютер": ("s", 20),
    "перезагрузи": ("r", 15), "перезагруз": ("r", 15), "перезагрузи компьютер": ("r", 15),
    "сон": (None, 0), "спящий режим": (None, 0), "гибернация": ("h", 5), "гибернацию": ("h", 5),
    "выйти из системы": ("l", 10), "завершить сеанс": ("l", 10),
}

SETTINGS_URIS = {
    "конфиденциальность": "ms-settings:privacy",
    "специальные возможности": "ms-settings:easeofaccess",
    "учетные записи": "ms-settings:accounts", "учётные записи": "ms-settings:accounts",
    "панель задач": "ms-settings:taskbar",
    "аккаунт": "ms-settings:accounts",
    "доступность": "ms-settings:easeofaccess",
    "хранилищ": "ms-settings:storagesense",
    "обновлен": "ms-settings:windowsupdate",
    "клавиатур": "ms-settings:keyboard",
    "язык": "ms-settings:keyboard",
    "дат": "ms-settings:dateandtime",
    "фон": "ms-settings:personalization-background",
    "персонал": "ms-settings:personalization",
    "прокси": "ms-settings:network-proxy",
    "приложен": "ms-settings:appsfeatures",
    "производительн": "ms-settings:display",
    "питани": "ms-settings:powersleep",
    "блютус": "ms-settings:bluetooth", "блюта": "ms-settings:bluetooth", "bluetooth": "ms-settings:bluetooth",
    "вай фай": "ms-settings:network-wifi", "wi fi": "ms-settings:network-wifi", "wifi": "ms-settings:network-wifi",
    "звук": "ms-settings:sound",
    "экран": "ms-settings:display", "дисплей": "ms-settings:display", "видео": "ms-settings:display",
    "сеть": "ms-settings:network",
}

KILL_ALIASES = {
    "блокнот": "notepad.exe", "ноутпад": "notepad.exe",
    "калькулятор": "calc.exe", "проводник": "explorer.exe",
    "хром": "chrome.exe", "дискорд": "discord.exe", "стим": "steam.exe",
    "телеграм": "telegram.exe", "спотифай": "spotify.exe",
    "ворд": "winword.exe", "word": "winword.exe", "эксель": "excel.exe",
    "пейнт": "mspaint.exe", "paint": "mspaint.exe",
}


def human_gb(size):
    return round(size / (1024 ** 3), 1)


def _ps(script, timeout=10):
    """Запуск PowerShell-скрипта без окна. Возвращает stdout+stderr текстом."""
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                           capture_output=True, timeout=timeout, creationflags=subprocess.CREATE_NO_WINDOW)
        return (r.stdout + b"\n" + r.stderr).decode("utf-8", errors="replace")
    except Exception:
        return ""


def temp_readings():
    """Возвращает список строк с температурами (CPU/GPU) или пустой список."""
    out = []
    txt = _ps("(Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace root/wmi -ErrorAction SilentlyContinue "
              "| ForEach-Object { [math]::Round(($_.CurrentTemperature/10) - 273.15, 1) }) -join ','", timeout=8)
    n = 0
    for v in txt.split(","):
        v = v.strip()
        if re.match(r"^-?\d+(\.\d+)?$", v):
            n += 1
            out.append("%d датчик: %s градусов" % (n, v.replace(".", ",")))
    gpu = _ps("$g = Get-Command nvidia-smi -ErrorAction SilentlyContinue; "
              "if ($g) { nvidia-smi --query-gpu=temperature.gpu --format=csv,noheader,nounits }", timeout=8).strip()
    for line in gpu.splitlines():
        line = line.strip()
        if re.match(r"^\d+$", line):
            out.append("видеокарта: %s градусов" % line)
    return out


def screenshot(path):
    try:
        import win32api
        import win32con
        import win32gui
        import win32ui
    except ImportError:
        return False
    try:
        width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
        height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
        left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
        top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
        hwnd_dc = win32gui.GetWindowDC(0)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(mfc_dc, width, height)
        save_dc.SelectObject(bmp)
        save_dc.BitBlt((0, 0), (width, height), mfc_dc, (left, top), win32con.SRCCOPY)
        bmp.SaveBitmapFile(save_dc, path)
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.DeleteObject(bmp.GetHandle())
        return os.path.exists(path)
    except Exception:
        return False


def get_volume():
    if AudioUtilities is None:
        return None
    try:
        from comtypes import CLSCTX_ALL
        dev = AudioUtilities.GetSpeakers()
        volume = cast(dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None), POINTER(IAudioEndpointVolume))
        return round(volume.GetMasterVolumeLevelScalar() * 100)
    except Exception:
        return None


def set_volume(percent):
    if AudioUtilities is None:
        return None
    try:
        from comtypes import CLSCTX_ALL
        dev = AudioUtilities.GetSpeakers()
        volume = cast(dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None), POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(max(0, min(100, percent)) / 100.0, None)
        return round(percent)
    except Exception:
        return None


class Jarvis:
    def __init__(self):
        self._setup_stdout()
        self.speech = Speech()
        self.ptt = None
        self.ptt_ok = False
        if not TEXT_MODE:
            self.ptt = PushToTalk()
            self.ptt_ok = self.ptt.start() and sr is not None
        else:
            self.ptt_ok = False

    # ---------- инициализация ----------
    @staticmethod
    def _setup_stdout():
        for name in ("stdout", "stderr"):
            stream = getattr(sys, name, None)
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass

    # ---------- вывод ----------
    def say(self, text):
        self.speech.say(text)

    @staticmethod
    def say_text_only(text):
        print("Джарвис:", text)

    # ---------- ввод ----------
    def recognize(self, audio):
        try:
            return self.recognizer.recognize_google(audio, language="ru-RU").lower().strip()
        except sr.UnknownValueError:
            return "@unknown"
        except sr.RequestError:
            return "!!network"

    def capture_ptt(self):
        """Записать голос, пока зажата комбинация, и вернуть распознанную фразу."""
        if not self.ptt_ok or pyaudio is None or sr is None:
            return ""
        self.ptt.wait_for_push()
        beep(True)
        recorder = Recorder()
        recorder.start()
        try:
            self.ptt.wait_for_release()
        finally:
            recorder.stop()
        beep(False)
        audio = recorder.audio()
        if audio is None:
            return "@empty"
        return self.recognize(audio)

    def ask(self, prompt):
        self.say(prompt)
        return self.capture_ptt()

    # ---------- команды ----------
    def cmd_time(self, text):
        now = datetime.datetime.now()
        self.say("Сейчас %d часов %d минут." % (now.hour, now.minute))

    def cmd_date(self, text):
        now = datetime.datetime.now()
        months = ["января", "февраля", "марта", "апреля", "мая", "июня",
                  "июля", "августа", "сентября", "октября", "ноября", "декабря"]
        self.say("Сегодня %s, %d %s %d года." % (WEEKDAYS[now.weekday()], now.day, months[now.month - 1], now.year))

    def cmd_hello(self, text):
        replies = ["Здравствуйте, сэр. Все системы в норме.",
                   "Добрый день. Чем могу помочь?",
                   "Приветствую, командир. Я готов к работе."]
        self.say(random.choice(replies))

    def cmd_mood(self, text):
        self.say("Как и всегда: на связи и готов помочь. Разве что процессор у меня потихоньку шумит от энтузиазма.")

    def cmd_who(self, text):
        self.say("Я Джарвис — ваш голосовой помощник. Зажмите Ctrl и Windows и скажите команду. "
                 "Умею говорить время и дату, открывать любое приложение и любой сайт, искать в интернете, "
                 "считать с корнями, степенями и процентами, показывать погоду, курс валют, новости. "
                 "Записываю заметки, ставлю таймеры и напоминания, делаю скриншоты, генерирую пароли, "
                 "подсказываю заряд батареи и загрузку компьютера, меняю громкость и рассказываю анекдоты.")

    @staticmethod
    def _open_target(text):
        for kw in ("открой сайт", "открой страницу", "запусти приложение", "открой приложение",
                   "запусти программу", "открой программу", "запусти", "открой"):
            if kw in text:
                return text.split(kw, 1)[1].strip(" ,.!?;:")
        return ""

    def _open_site(self, name):
        host = name.replace("точка ру", ".ру").replace("точка ком", ".ком") \
            .replace("точка", ".").replace(".ру", ".ru").replace(".ком", ".com") \
            .replace(".нет", ".net").strip(" ,.!?;:")
        if " " in host and "." not in host:
            host = re.sub(r"\s+", "", host)
        host = host.lower()
        if host.startswith("http://") or host.startswith("https://"):
            webbrowser.open(host)
            self.say("Открываю сайт.")
            return True
        if "." not in host:
            for suffix in (".ru", ".com", ".net"):
                if resolve(host + suffix):
                    webbrowser.open("https://" + host + suffix)
                    self.say("Открываю сайт: %s." % host)
                    return True
        elif resolve(host) or resolve(host.split("/")[0]):
            webbrowser.open("https://" + host)
            self.say("Открываю сайт: %s." % host)
            return True
        return False

    def try_open_app(self, name):
        for alias, exe in APPS.items():
            if alias in name:
                try:
                    os.startfile(exe)
                except Exception:
                    return False
                self.say("Открываю %s." % alias)
                return True
        found = find_app(name)
        if found:
            try:
                os.startfile(found)
            except Exception:
                return False
            self.say("Открываю приложение %s." % name)
            return True
        return False

    def _cmd_open(self, text):
        name = self._open_target(text)
        for alias, url in SITES.items():
            if alias in text:
                webbrowser.open(url)
                self.say("Открываю %s." % alias)
                return
        force_app = any(k in text for k in ("приложение", "программу", "запусти", "программа"))
        force_web = any(k in text for k in ("сайт", "страницу", "веб", "интернете"))
        if not name:
            self.say("Что открыть? Например, «открой калькулятор» или «открой сайт ютуб».")
            return
        if not force_web:
            if self.try_open_app(name):
                return
        if force_app:
            self.say("Приложение «%s» не найдено." % name)
            return
        if not self._open_site(name):
            self.say("Не нашёл приложение или сайт «%s». Открываю поиск Яндекса." % name)
            webbrowser.open("https://yandex.ru/search/?text=" + urllib.parse.quote(name))

    def cmd_search(self, text):
        query = _clean_search(text, ["поищи", "найди", "погугли", "поиск", "найди в интернете"])
        if not query:
            query = _clean_search(text, ["что такое", "кто такой", "кто такая"])
        if not query:
            self.say("Не услышал, что искать. Повторите запрос.")
            return
        engine = "https://www.google.com/search?q=" if "погугли" in text else "https://yandex.ru/search/?text="
        webbrowser.open(engine + urllib.parse.quote(query))
        self.say("Ищу в интернете: %s." % query)

    def cmd_joke(self, text):
        self.say(random.choice(JOKES))

    def cmd_calc(self, text):
        pct = re.search(r"([а-яё0-9 ]+?)\s+процент\w*\s+от\s+([а-яё0-9 ]+)$", text)
        if pct:
            a = words_to_int(pct.group(1))
            b = words_to_int(pct.group(2))
            if a is not None and b:
                self.say("Результат: %s." % str(a / 100 * b).replace(".", ","))
                return
        expr = words_to_digits(text)
        expr = re.sub(r"(\d+)\s*в\s+квадрате", r"(\1)**2", expr)
        expr = re.sub(r"(\d+)\s*в\s+кубе", r"(\1)**3", expr)
        expr = re.sub(r"(\d+)\s*в\s+степени\s+(\d+)", r"(\1)**(\2)", expr)
        expr = re.sub(r"корень\s+из\s+(\d+)", r"(\1)**0.5", expr)
        expr = expr.replace("умножить на", "*").replace("умножить", "*").replace("разделить на", "/") \
            .replace("разделить", "/").replace("делить на", "/").replace("плюс", "+") \
            .replace("минус", "-").replace("прибавить", "+").replace("вычесть", "-") \
            .replace("знак равно", "=").replace("равно", "").replace("сколько будет", "") \
            .replace("х", "*").replace("на", "")
        expr = re.sub(r"[^0-9+\-*/().\s]", "", expr)
        expr = expr.replace(" ", "")
        if not expr:
            self.say("Не понял пример. Скажите, например, «сколько будет два плюс два».")
            return
        try:
            result = eval(expr)  # выражение прошло фильтр допустимых символов
            self.say("Результат: %s." % (str(result).replace(".", ",")))
        except Exception:
            self.say("Не смог вычислить. Попробуйте проще.")

    def cmd_weather(self, text):
        city = None
        for key in ("в городе", "погода в"):
            if key in text:
                city = _after(text, key).strip()
                break
        if not city:
            city = self.ask("В каком городе узнать погоду?")
            if not city or city in ("@empty", "!!network", "@unknown"):
                self.say("Не услышал город. Попробуйте еще раз.")
                return
        try:
            candidates = [city, city.capitalize()]
            c = city.rstrip(".,;:!?")
            if c.endswith("е"):
                candidates.append(c[:-1])
                candidates.append(c[:-1] + "а")
            if c.endswith("и"):
                candidates.append(c[:-1])
                candidates.append(c[:-1] + "ь")
            if c.endswith("у"):
                candidates.append(c[:-1])
            if c.endswith("е") and len(c) > 4:
                candidates.append(c[:-1] + "а")
            place = None
            for cand in dict.fromkeys(candidates):
                geo = safe_url("https://geocoding-api.open-meteo.com/v1/search"
                               "?name=%s&count=1&language=ru&format=json" % urllib.parse.quote(cand))
                if geo.get("results"):
                    place, city = geo["results"][0], cand
                    break
            if place is None:
                self.say("Город «%s» не найден." % city)
                return
            lat, lon = place["latitude"], place["longitude"]
            named = place.get("name", city)
            wx = safe_url("https://api.open-meteo.com/v1/forecast?latitude=%f&longitude=%f&current_weather=true"
                          % (lat, lon))
            cw = wx.get("current_weather", {})
            temp = round(cw.get("temperature", 0))
            wind = round(cw.get("windspeed", 0))
            desc = WHETHER_CODES.get(cw.get("weathercode", 0), "неизвестные условия")
            self.say("В городе %s сейчас %s, около %d градусов, ветер %d километров в час." %
                     (named, desc, temp, wind))
        except Exception as e:
            self.say("Не удалось получить погоду. Проверьте интернет. Ошибка: %s" % e)

    def cmd_volume(self, text):
        expr = words_to_digits(text)
        m = re.search(r"(\d+)", expr)
        if m:
            pct = max(0, min(100, int(m.group(1))))
            cur = set_volume(pct)
            if cur is None:
                self.say("Не удалось изменить громкость.")
            else:
                self.say("Громкость установлена на %d процентов." % pct)
            return
        if "какая громкость" in text or "уровень громкости" in text or "сейчас громкость" in text:
            cur = get_volume()
            if cur is None:
                self.say("Не удалось узнать уровень громкости.")
            else:
                self.say("Сейчас громкость %d процентов." % cur)
            return
        keys = {"увеличь": 0xAF, "громче": 0xAF, "прибавь": 0xAF,
                "уменьш": 0xAE, "тише": 0xAE, "убавь": 0xAE}
        if "звук" in text and "увеличь" not in text and "уменьш" not in text:
            key = 0xAD
        else:
            key = None
            for k, v in keys.items():
                if k in text:
                    key = v
                    break
        if key is None:
            self.say("Скажите «увеличь громкость», «уменьши громкость», «выключи звук» или «громкость на 50».")
            return
        for _ in range(3):
            ctypes.windll.user32.keybd_event(key, 0, 0, 0)
            ctypes.windll.user32.keybd_event(key, 0, 2, 0)
        words = {0xAF: "Увеличиваю громкость.", 0xAE: "Уменьшаю громкость.",
                 0xAD: "Отключаю звук."}
        self.say(words.get(key, "Готово."))

    def cmd_system(self, text):
        try:
            class MEM(ctypes.Structure):
                _fields_ = [("dwLength", ctypes.c_uint), ("dwMemoryLoad", ctypes.c_uint),
                            ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                            ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong),
                            ("ullTotalVirtual", ctypes.c_ulonglong), ("ullAvailVirtual", ctypes.c_ulonglong),
                            ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]
            m = MEM()
            m.dwLength = ctypes.sizeof(MEM)
            ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(m))
            total_gb = m.ullTotalPhys / (1024 ** 3)
        except Exception:
            total_gb = None
        proc = platform.processor() or "неизвестный процессор"
        self.say("Имя компьютера: %s. Процессор: %s. Разрядность: %s. " %
                 (platform.node(), proc, platform.machine()))
        if total_gb:
            self.say("Оперативной памяти: примерно %d гигабайт." % round(total_gb))

    def cmd_thanks(self, text):
        self.say("Всегда пожалуйста, сэр.")

    def cmd_meow(self, text):
        self.say("Мяу!")

    # ---------- заметки ----------
    def cmd_note(self, text):
        note = ""
        for kw in ("запомни", "запоминай", "запиши", "сохрани", "заметка", "занеси", "отметь"):
            if kw in text:
                note = _after(text, kw).strip()
                break
        if not note:
            self.say("Что запомнить?")
            return
        stamp = datetime.datetime.now().strftime("%d.%m %H:%M")
        with open(NOTES_FILE, "a", encoding="utf-8") as f:
            f.write("[%s] %s\n" % (stamp, note))
        self.say("Запомнил: %s." % note)

    def cmd_read_notes(self, text):
        if not os.path.exists(NOTES_FILE):
            self.say("Заметок пока нет.")
            return
        with open(NOTES_FILE, encoding="utf-8") as f:
            lines = [l.strip() for l in f if l.strip()]
        if not lines:
            self.say("Заметок пока нет.")
            return
        if len(lines) > 5:
            hdr = "У вас %d заметок. Последние:" % len(lines)
        else:
            hdr = "Ваши заметки:"
        self.say(hdr)
        for line in lines[-5:]:
            self.say(line)

    def cmd_clear_notes(self, text):
        if os.path.exists(NOTES_FILE):
            os.remove(NOTES_FILE)
        self.say("Заметки удалены.")

    # ---------- скриншот ----------
    def cmd_screenshot(self, text):
        folder = os.path.join(os.path.expanduser("~"), "Pictures", "Screenshots")
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, "jarvis_%s.png" % datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
        if screenshot(path):
            self.say("Скриншот сохранён в папку изображений: %s." % path)
        else:
            self.say("Не удалось сделать скриншот.")

    # ---------- блокировка и питание ----------
    @staticmethod
    def _exec_power(arg, delay):
        if arg == "s":
            subprocess.Popen(["shutdown", "/s", "/t", str(delay)], shell=(os.name == "nt"))
        elif arg == "r":
            subprocess.Popen(["shutdown", "/r", "/t", str(delay)])
        elif arg == "h":
            subprocess.Popen(["shutdown", "/h"])
        elif arg == "l":
            subprocess.Popen(["shutdown", "/l"])
        else:
            try:
                ctypes.windll.powrprof.SetSuspendState(0, 1, 0)
            except Exception:
                pass

    def confirm(self, prompt):
        self.say(prompt)
        ans = self.capture_ptt()
        if ans and re.search(r"\b(да|подтверждаю|ага|угу|точно|го|давай)\b", ans):
            return True
        self.say("Отменяю.")
        return False

    def cmd_power(self, text):
        for match in POWER_ACTIONS:
            if re.search(r"\b%s\b" % match, text):
                arg, delay = POWER_ACTIONS[match]
                if not self.confirm("Подтвердите: %s?" % match):
                    return
                self._exec_power(arg, delay)
                self.say("Выполняю.")
                return
        self.say("Что сделать? Скажите «выключи компьютер», «перезагрузи», «сон» или «гибернация».")

    def cmd_lock(self, text):
        try:
            ctypes.windll.user32.LockWorkStation()
        except Exception:
            self.say("Не удалось заблокировать.")

    def cmd_monitor_off(self, text):
        try:
            ctypes.windll.user32.SendMessageW(0xFFFF, 0x0112, 0xF170, 2)
        except Exception:
            self.say("Не удалось выключить монитор.")

    def cmd_empty_bin(self, text):
        try:
            ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, 7)
            self.say("Корзина очищена.")
        except Exception:
            self.say("Не удалось очистить корзину.")

    # ---------- системный мониторинг ----------
    def cmd_battery(self, text):
        if psutil is None:
            self.say("Библиотека psutil не установлена.")
            return
        try:
            b = psutil.sensors_battery()
        except Exception:
            b = None
        if not b:
            self.say("Батарея не найдена. Похоже, у компьютера нет аккумулятора.")
            return
        status = "от сети" if b.power_plugged else "от батареи"
        self.say("Заряд батареи: %d процентов, питание %s." % (b.percent, status))

    def cmd_usage(self, text):
        if psutil is None:
            self.say("Библиотека psutil не установлена.")
            return
        cpu = psutil.cpu_percent(interval=1)
        ram = psutil.virtual_memory().percent
        self.say("Процессор загружен на %d %s, память занята на %d %s." %
                 (cpu, _ru_plural(cpu, "процент", "процента", "процентов"),
                  ram, _ru_plural(ram, "процент", "процента", "процентов")))

    def cmd_disk(self, text):
        if psutil is None:
            self.say("Библиотека psutil не установлена.")
            return
        try:
            disk = psutil.disk_usage(os.environ.get("SystemDrive", "C:"))
            self.say("На системном диске свободно %.1f гигабайт из %.1f. Занято %d %s." %
                     (human_gb(disk.free), human_gb(disk.total), disk.percent,
                      _ru_plural(disk.percent, "процент", "процента", "процентов")))
        except Exception:
            self.say("Не удалось узнать место на диске.")

    def cmd_uptime(self, text):
        if psutil is None:
            self.say("Библиотека psutil не установлена.")
            return
        sec = int(time.time() - psutil.boot_time())
        h, m = divmod(sec // 60, 60)
        self.say("Компьютер работает %d %s %d %s." %
                 (h, _ru_plural(h, "час", "часа", "часов"),
                  m, _ru_plural(m, "минуту", "минуты", "минут")))

    # ---------- таймер и напоминания ----------
    @staticmethod
    def _parse_dur(text, expr):
        """Возвращает (секунды, подпись) по фразам вида «через 30 минут», «на час», «через полчаса»."""
        m = re.search(r"(?:через|на|потом|позже)\s+(\d+)\s*(секунд\w*|минут\w*|час\w*)", expr)
        if m:
            value = int(m.group(1))
            unit = m.group(2)
            if unit.startswith("ча"):
                return value * 3600, "%d %s" % (value, _ru_plural(value, "час", "часа", "часов"))
            if unit.startswith("мину"):
                return value * 60, "%d %s" % (value, _ru_plural(value, "минуту", "минуты", "минут"))
            return value, "%d %s" % (value, _ru_plural(value, "секунду", "секунды", "секунд"))
        if re.search(r"полчас|пол часа", text):
            return 1800, "30 минут"
        if re.search(r"пол минут|пол минутк", text):
            return 30, "30 секунд"
        if re.search(r"час\b|часик|часок", text):
            return 3600, "1 час"
        if re.search(r"минут\w*|минутк\w*", text):
            return 60, "1 минуту"
        if re.search(r"секунд\w*", text):
            return 1, "1 секунду"
        return None, None

    def _after_delay(self, seconds, message):
        def run():
            time.sleep(seconds)
            self.say(message)
        threading.Thread(target=run, daemon=True).start()

    def cmd_timer(self, text):
        sec, label = self._parse_dur(text, words_to_digits(text))
        if sec is None:
            self.say("Скажите, например, «таймер на 5 минут» или «таймер на час».")
            return
        self._after_delay(sec, "Таймер завершён.")
        self.say("Таймер на %s запущен." % label)

    def cmd_remind(self, text):
        expr = words_to_digits(text)
        sec, label = self._parse_dur(text, expr)
        if sec is None:
            self.say("Скажите, например, «напомни мне через 5 минут выключить чайник».")
            return
        rest = re.sub(r"напомни.*?через\s+(?:\d+\s*)?[а-яё]+", "", text).strip(" ,.!?;:")
        memo = rest
        for w in ["мне что", "мне", "что"]:
            if memo.startswith(w + " "):
                memo = memo[len(w):].strip(" ,.!?;:")
                break
        memo = memo or "напоминание"
        self._after_delay(sec, "Напоминание: %s." % memo)
        self.say("Напомню через %s." % label)

    # ---------- случайности ----------
    def cmd_random(self, text):
        expr = words_to_digits(text)
        if "монет" in text:
            self.say("Выпало: %s." % random.choice(["орёл", "решка"]))
            return
        if re.search(r"\b(кубик|кость|кост)\b", text):
            self.say("На кубике выпало: %d." % random.randint(1, 6))
            return
        m = re.search(r"от\s+(\d+).*?(\d+)", expr)
        if m:
            lo, hi = int(m.group(1)), int(m.group(2))
            if lo > hi:
                lo, hi = hi, lo
            self.say("Случайное число: %d." % random.randint(lo, hi))
            return
        self.say("Случайное число: %d." % random.randint(1, 100))

    def cmd_password(self, text):
        expr = words_to_digits(text)
        m = re.search(r"(\d+)", expr)
        n = min(32, max(6, int(m.group(1)) if m else 12))
        chars = string.ascii_letters + string.digits + "@#$%&*!?:"
        pw = "".join(random.choice(chars) for _ in range(n))
        try:
            subprocess.run(["clip"], input=("\ufeff" + pw).encode("utf-16"), check=False)
            self.say("Пароль готов и скопирован в буфер обмена.")
        except Exception:
            self.say("Пароль: %s." % pw)

    # ---------- перевод и валюты ----------
    def cmd_translate(self, text):
        langs = {"английск": "en", "русск": "ru", "немецк": "de", "французск": "fr",
                 "испанск": "es", "итальянск": "it", "китайск": "zh", "японск": "ja"}
        target = None
        phrase = None
        m = re.match(r"(?:переведи|перевод)\s+(.+?)(?:\s+на\s+([а-яё]+))?\s*$", text)
        if m:
            phrase, lang_word = m.group(1), m.group(2)
            if lang_word:
                for lang, code in langs.items():
                    if lang in lang_word:
                        target = code
                        break
        else:
            phrase = _after(text, "перевести")
        if not phrase:
            phrase = text.replace("переведи", "").replace("перевод", "").strip(" ,.!?;:")
        phrase = phrase.strip(" ,.!?;:")
        if not phrase:
            self.say("Что перевести?")
            return
        if target is None:
            target = "ru" if re.search(r"[a-zA-Z]", phrase) else "en"
        url = ("https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=%s&dt=t&q=%s"
               % (target, urllib.parse.quote(phrase)))
        try:
            data = safe_url(url, timeout=10)
            result = "".join(seg[0] for seg in data[0] if seg[0])
            self.say("Перевод: %s." % result)
        except Exception:
            self.say("Не удалось перевести. Проверьте интернет.")

    def cmd_currency(self, text):
        currencies = {"доллар": "USD", "евро": "EUR", "фунт": "GBP", "юань": "CNY",
                      "йен": "JPY", "йена": "JPY", "тенг": "KZT"}
        code = None
        for name, c in currencies.items():
            if name in text:
                code = c
                break
        if code is None:
            self.say("Скажите, например, «курс доллара» или «курс евро».")
            return
        try:
            data = safe_url("https://www.cbr-xml-daily.ru/daily_json.js")["Valute"]
            rate = data[code]["Value"]
            name = next((n for n, c in currencies.items() if c == code), code)
            self.say("%s стоит примерно %d %s." % (name.capitalize(), round(rate),
                     _ru_plural(int(round(rate)), "рубль", "рубля", "рублей")))
        except Exception:
            self.say("Не удалось получить курс. Проверьте интернет.")

    def cmd_newyear(self, text):
        now = datetime.date.today()
        ny = datetime.date(now.year + 1, 1, 1)
        days = (ny - now).days
        self.say("До Нового года осталось %d %s." % (days, "день" if days % 10 == 1 and days % 100 != 11 else "дня" if days % 10 in (2, 3, 4) and days % 100 not in (12, 13, 14) else "дней"))

    def cmd_netcheck(self, text):
        if resolve("ya.ru"):
            self.say("Интернет работает, сэр.")
        else:
            self.say("Интернет недоступен.")

    # ---------- новые системные и сетевые команды ----------
    def _netsh(self, args):
        try:
            r = subprocess.run(["netsh", "wlan"] + args.split(),
                               capture_output=True, timeout=10, creationflags=subprocess.CREATE_NO_WINDOW)
            return (r.stdout + r.stderr).decode("utf-8", errors="replace")
        except Exception:
            return ""

    def _last_wifi_profile(self):
        out = self._netsh("show profiles")
        for m in re.finditer(r"(?:Профиль|Profile)\s*:\s*(.+)$", out, re.M):
            return m.group(1).strip()
        return ""

    def cmd_netinfo(self, text):
        if psutil is None:
            self.say("Не установлен psutil. Установите его через pip.")
            return
        try:
            a = psutil.net_io_counters()
            time.sleep(1)
            b = psutil.net_io_counters()
            down = max(0, b.bytes_recv - a.bytes_recv) / 1024.0
            up = max(0, b.bytes_sent - a.bytes_sent) / 1024.0
            active = [i for i, s in psutil.net_if_stats().items() if s.isup]
            extra = ", активные интерфейсы: %s" % ", ".join(active) if active else ""
            self.say("Скорость загрузки %.0f килобайт в секунду, отдачи %.0f%s." % (down, up, extra))
        except Exception:
            self.say("Не удалось измерить скорость сети.")

    def cmd_wifi(self, text):
        off = any(k in text for k in ("выключ", "отключ", "убери"))
        out = self._netsh("show interfaces")
        lower = out.lower()
        connected = "connected" in lower or "подключено" in lower
        if off:
            self._netsh("disconnect")
            self.say("Wi-Fi отключен.")
        elif connected or any(k in text for k in ("статус", "работает", "сигнал", "включ")):
            if connected:
                self.say("Wi-Fi включен и подключен к сети.")
            else:
                self.say("Wi-Fi отключен. Скажите «включи вай фай», и я подключусь к последней сети.")
        else:
            profile = self._last_wifi_profile()
            if profile:
                self._netsh('connect name="%s"' % profile.replace('"', "'"))
                self.say("Подключаюсь к сети %s." % profile)
            else:
                self.say("Сохранённых сетей не нашёл.")

    def cmd_bluetooth(self, text):
        off = any(k in text for k in ("выключ", "отключ"))
        verb = "Disable" if off else "Enable"
        script = (
            "$d = Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue "
            "| Where-Object { $_.FriendlyName -match 'bluetooth|блюта|блютус|адаптер|radio' } "
            "| Select-Object -First 1; "
            "if ($d) { %s -InstanceId $d.InstanceId -Confirm:$false; 'OK' } else { 'NO_DEVICE' }" % verb
        )
        res = _ps(script, 15)
        if "NO_DEVICE" in res:
            self.say("Bluetooth-адаптер не найден.")
        elif "OK" in res and not re.search(r"(denied|отказано|недостаточно|permission|доступ)", res, re.I):
            self.say("Bluetooth %s." % ("выключен" if off else "включен"))
        else:
            self.say("Не удалось переключить Bluetooth. Попробуйте запустить Джарвиса от имени администратора.")

    def cmd_winupdate(self, text):
        try:
            date = _ps("(Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1)"
                       ".InstalledOn.ToString('dd.MM.yyyy')", 10).strip()
        except Exception:
            date = ""
        try:
            os.startfile("ms-settings:windowsupdate")
        except Exception:
            pass
        if re.match(r"^\d{2}\.\d{2}\.\d{4}$", date):
            self.say("Последнее обновление установлено %s. Открываю центр обновлений Windows." % date)
        else:
            self.say("Открываю центр обновлений Windows.")

    def cmd_shutdown_later(self, text):
        sec, label = self._parse_dur(text, words_to_digits(text))
        if sec is None:
            self.say("Скажите, например, «выключи компьютер через 30 минут» или «перезагрузи через час».")
            return
        restart = "перезагрузи" in text or "перезагрузить" in text or "перезагруз" in text
        flag = "/r" if restart else "/s"
        action = "перезагружу" if restart else "выключу"
        if not self.confirm("Через %s %s компьютер?" % (label, "перезагрузить" if restart else "выключить")):
            return
        try:
            subprocess.Popen(["shutdown", flag, "/t", str(int(sec))], shell=(os.name == "nt"))
            self.say("Хорошо, %s компьютер через %s." % (action, label))
        except Exception:
            self.say("Не удалось запланировать выключение.")

    def cmd_shutdown_cancel(self, text):
        try:
            r = subprocess.run(["shutdown", "/a"], capture_output=True, timeout=10,
                               creationflags=subprocess.CREATE_NO_WINDOW)
            if r.returncode == 0:
                self.say("Запланированное выключение отменено.")
            else:
                self.say("Запланированного выключения не было.")
        except Exception:
            self.say("Не удалось отменить выключение.")

    def cmd_wallpaper(self, text):
        folders = [os.path.join(os.path.expanduser("~"), "Pictures"),
                   os.path.join(os.path.expanduser("~"), "Pictures", "Wallpapers"),
                   os.path.join(os.path.expanduser("~"), "Pictures", "Обои")]
        imgs = []
        for f in folders:
            if os.path.isdir(f):
                for ext in ("*.jpg", "*.jpeg", "*.png", "*.bmp"):
                    imgs.extend(glob.glob(os.path.join(f, "**", ext), recursive=True))
        if not imgs:
            self.say("Не нашёл картинок в папке «Изображения». Добавьте туда обои.")
            return
        path = random.choice(imgs)
        try:
            if ctypes.windll.user32.SystemParametersInfoW(0x0014, 0, path, 0x01 | 0x02):
                self.say("Обои обновлены.")
            else:
                self.say("Не удалось сменить обои.")
        except Exception:
            self.say("Не удалось сменить обои.")

    def cmd_clean_temp(self, text):
        targets = [os.environ.get("TEMP"), os.environ.get("TMP"),
                   os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "Temp")]
        freed = 0
        removed = 0
        for folder in {t for t in targets if t and os.path.isdir(t)}:
            for root, dirs, files in os.walk(folder, topdown=False):
                for fn in files:
                    p = os.path.join(root, fn)
                    try:
                        freed += os.path.getsize(p)
                        os.remove(p)
                        removed += 1
                    except Exception:
                        pass
                for dn in dirs:
                    try:
                        shutil.rmtree(os.path.join(root, dn))
                    except Exception:
                        pass
        if removed == 0:
            self.say("Временные файлы уже чистые.")
            return
        mb = freed / 1048576.0
        self.say("Очищено %d временных файлов, освобождено примерно %.1f мегабайт." % (removed, mb))

    def cmd_temps(self, text):
        readings = temp_readings()
        if readings:
            self.say("Температуры системы: %s." % ", ".join(readings))
            return
        if psutil is None:
            self.say("Не удалось получить температуру на этом оборудовании.")
            return
        cpu = psutil.cpu_percent(interval=1)
        self.say("Температурные датчики недоступны на этом оборудовании, "
                 "но процессор сейчас загружен на %d процентов." % cpu)

    def cmd_kill(self, text):
        name = ""
        for kw in ("убей процесс", "заверши процесс", "завершить процесс", "снять задачу",
                   "сними задачу", "закрой процесс", "убей", "убери"):
            if kw in text:
                name = _after(text, kw).strip()
                break
        if not name:
            if not any(k in text for k in ("убей", "убейте", "убить", "уничтожь",
                                           "заверши", "закрой", "процесс", "задач")):
                self.say("Какой процесс завершить? Например, «убей блокнот».")
                return
            words = re.split(r"\s+", text.strip())
            name = words[-1] if words else ""
        name = name.strip(" ,.!?;:")
        if not name:
            self.say("Какой процесс завершить? Например, «убей блокнот».")
            return
        base = KILL_ALIASES.get(name.lower())
        if not base:
            base = name if name.lower().endswith(".exe") else name + ".exe"
        if psutil is not None:
            running = [p.info.get("name") or "" for p in psutil.process_iter(["name"])]
            lowered = [r.lower() for r in running]
            if base.lower() not in lowered:
                for r in running:
                    if r.lower().startswith(name.lower()):
                        base = r
                        break
        try:
            r = subprocess.run(["taskkill", "/F", "/IM", base], capture_output=True, timeout=10,
                               creationflags=subprocess.CREATE_NO_WINDOW)
            if r.returncode == 0:
                self.say("Процесс %s завершён." % base)
            else:
                self.say("Процесс %s не найден или уже закрыт." % base)
        except Exception:
            self.say("Не удалось завершить процесс.")

    def cmd_layout(self, text):
        try:
            ctypes.windll.user32.keybd_event(0x5B, 0, 0, 0)
            ctypes.windll.user32.keybd_event(0x20, 0, 0, 0)
            ctypes.windll.user32.keybd_event(0x20, 0, 2, 0)
            ctypes.windll.user32.keybd_event(0x5B, 0, 2, 0)
            self.say("Раскладка переключена.")
        except Exception:
            self.say("Не удалось переключить раскладку.")

    def cmd_start_menu(self, text):
        try:
            ctypes.windll.user32.keybd_event(0x5B, 0, 0, 0)
            ctypes.windll.user32.keybd_event(0x5B, 0, 2, 0)
        except Exception:
            self.say("Не удалось открыть меню «Пуск».")

    def cmd_settings(self, text):
        for k, uri in SETTINGS_URIS.items():
            if k in text:
                try:
                    os.startfile(uri)
                    self.say("Открываю параметры: %s." % k)
                except Exception:
                    self.say("Не удалось открыть параметры.")
                return
        try:
            os.startfile("ms-settings:")
            self.say("Открываю параметры Windows.")
        except Exception:
            self.say("Не удалось открыть параметры.")

    # ---------- окна и папки ----------
    def cmd_show_desktop(self, text):
        try:
            ctypes.windll.user32.keybd_event(0x5B, 0, 0, 0)
            ctypes.windll.user32.keybd_event(ord("D"), 0, 0, 0)
            ctypes.windll.user32.keybd_event(ord("D"), 0, 2, 0)
            ctypes.windll.user32.keybd_event(0x5B, 0, 2, 0)
        except Exception:
            pass

    def cmd_open_folder(self, text):
        for alias, folder in SPECIAL_FOLDERS.items():
            if alias in text:
                path = os.path.join(os.path.expanduser("~"), folder)
                if os.path.isdir(path):
                    os.startfile(path)
                    self.say("Открываю папку %s." % folder)
                else:
                    self.say("Папка не найдена.")
                return True
        return False

    def cmd_voice(self, text):
        if "женский" in text:
            self.speech.switch_voice(1)
        elif "мужской" in text:
            self.speech.switch_voice(0)
        else:
            self.speech.switch_voice()
        self.say("Мой голос теперь %s." % self.speech.voice_label())

    def cmd_roblox(self, text):
        path = self._roblox_path()
        if not path:
            self.say("Roblox не найден на компьютере.")
            return
        try:
            os.startfile(path)
            self.say("Запускаю Roblox.")
        except Exception as e:
            self.say("Не удалось запустить Roblox: %s" % e)

    @staticmethod
    def _roblox_path():
        candidates = []
        versions = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Roblox", "Versions")
        try:
            if os.path.isdir(versions):
                for d in os.listdir(versions):
                    launcher = os.path.join(versions, d, "RobloxPlayerLauncher.exe")
                    if os.path.isfile(launcher):
                        candidates.append(launcher)
        except Exception:
            pass
        candidates += [
            r"C:\XboxGames\Roblox\Content\RobloxPlayerBeta.exe",
            os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files")),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Roblox\Versions\RobloxPlayerLauncher.exe"),
        ]
        user_lnk = os.path.join(os.environ.get("APPDATA", ""),
                                r"Microsoft\Windows\Start Menu\Programs\Roblox.lnk")
        common_lnk = os.path.join(os.environ.get("ProgramData", r"C:\ProgramData"),
                                  r"Microsoft\Windows\Start Menu\Programs\Roblox.lnk")
        candidates += [user_lnk, common_lnk,
                       r"C:\Program Files\Roblox\Versions\RobloxPlayerLauncher.exe"]
        for c in candidates:
            if c and os.path.isfile(c):
                return c
        return None

    # ---------- диспетчер ----------
    def process(self, text):
        rules = [
            (r"(переведи|перевод)", self.cmd_translate),
            (r"\b(привет|здравств|добрый день|доброе утро|добрый вечер|здорово|приветствую)\b", self.cmd_hello),
            (r"\b(который час|сколько время|сколько времени|какое время|что за время)\b", self.cmd_time),
            (r"\b(какое сегодня число|какая сегодня дата|какой сегодня день|день недели)\b", self.cmd_date),
            (r"\b(как дела|как настроение|как ты)\b", self.cmd_mood),
            (r"\b(кто ты|что ты умеешь|чем ты можешь помочь|помощь|команды)\b", self.cmd_who),
            (r"погод|температура на улице|градусов на улице|градусов за окном", self.cmd_weather),
            (r"(поищи|найди|погугли|поиск|что такое|кто такой|кто такая)", self.cmd_search),
            (r"(анекдот|шутку|рассмеши|пошути)", self.cmd_joke),
            (r"(курс|сколько стоит|цена)", self.cmd_currency),
            (r"(до нового года|до нг|новый год через|сколько.{0,6}новый год)", self.cmd_newyear),
            (r"(сколько будет|вычисли|посчитай|корень из|квадрат|в кубе|степени|процент\w*\s+от)", self.cmd_calc),
            (r"(проверь интернет|интернет работает|проверь сеть|сеть есть|доступ в интернет)", self.cmd_netcheck),
            (r"(скорость интернета|скорость сети|сетевые подключения|сетевое соединение)", self.cmd_netinfo),
            (r"(вай фай|вайфай|wi[\s-]*fi|беспроводн)", self.cmd_wifi),
            (r"(включ\w*|выключ\w*|отключ\w*)\s+(блютус|блютуз|blue?tooth)", self.cmd_bluetooth),
            (r"(проверь обновления|обновления (windows|виндовс)|обнови (windows|виндовс|систему)|центр обновлений)", self.cmd_winupdate),
            (r"(покажи заметки|что ты помнишь|какие заметки|прочитай заметки)", self.cmd_read_notes),
            (r"(очисти|удали).{0,10}замет|забудь всё|забудь все", self.cmd_clear_notes),
            (r"(запомни|запоминай|запиши|сохрани|заметка|занеси|отметь)", self.cmd_note),
            (r"напомни|напоминани|напомню", self.cmd_remind),
            (r"(таймер|засеки|отсчет|секундомер)", self.cmd_timer),
            (r"(покажи рабочий стол|сверни все окна|сверни окна|покажи стол)", self.cmd_show_desktop),
            (r"(открой|покажи).{0,4}(загрузк|документ|музык|картинк|изображени|видео|рабочий стол|рабочего стола|папк)", self.cmd_open_folder),
            (r"(заряд батареи|заряжена|батаре|аккумулятор|сколько заряда)", self.cmd_battery),
            (r"(громкость|громче|тише|увеличь|уменьш|выключи звук|прибавь|убавь)", self.cmd_volume),
            (r"(отмени|отмена|отмену).{0,8}(выключен|перезагруз|завершен)", self.cmd_shutdown_cancel),
            (r"(выключи|выключить|перезагрузи|перезагрузить|перезагрузк).{0,24}(через|потом|позже)", self.cmd_shutdown_later),
            (r"убей процесс|заверши(ть)? процесс|снять задачу|сними задачу|закрой процесс|\bубей\b", self.cmd_kill),
            (r"(смени обои|поменяй обои|поменять обои|новые обои|обнови обои|обои)", self.cmd_wallpaper),
            (r"(выключи компьютер|выключить компьютер|выключи пк|перезагруз|спящий режим|гибернац|выйти из системы|завершить сеанс|\bсон\b)", self.cmd_power),
            (r"(загрузк\w*.{0,6}(процессор|память|компьютер|система)|нагрузк)", self.cmd_usage),
            (r"(свободн|сколько места|место на диске|заполнен диск)", self.cmd_disk),
            (r"(как давно.{0,6}включен|аптайм|сколько.{0,4}работает|давно работает|время работы)", self.cmd_uptime),
            (r"\b(кубик|кость|орел|решка)\b|монет|случайное число|рандом(ное)? число|случайную цифру", self.cmd_random),
            (r"(пароль|генератор паролей|сгенерируй)", self.cmd_password),
            (r"(скриншот|скрин|сделай.{0,8}(экран|скрин)|фото экрана|сними экран)", self.cmd_screenshot),
            (r"(заблокируй|заблокировать|блокировка)", self.cmd_lock),
            (r"(выключи монитор|выключи экран|погаси экран|гаси монитор|погаси монитор)", self.cmd_monitor_off),
            (r"(очисти корзину|очисть корзину|пуста.{0,8}корзина)", self.cmd_empty_bin),
            (r"(температур|перегрев|нагрет|сколько градусов)", self.cmd_temps),
            (r"(характеристик|информаци|о компьютере|о системе|железо|процессор|память)", self.cmd_system),
            (r"(спасибо|благодар|неплохо|молодец)", self.cmd_thanks),
            (r"(мяу|мур)", self.cmd_meow),
            (r"(смени|поменяй|другой).{0,4}голос|женский голос|мужской голос", self.cmd_voice),
            (r"(deploy|деплой|диплой|включи|запуск|запусти|запустить|открой|открыть).{0,6}(roblox|роблокс)", self.cmd_roblox),
            (r"(раскладк|переключи язык|смени язык|другую раскладку)", self.cmd_layout),
            (r"(меню пуск|кнопка пуск|покажи пуск|открой пуск)", self.cmd_start_menu),
            (r"(параметры|настройки)", self.cmd_settings),
            (r"(очисти|почисти|очисть).{0,12}(временн|кэш)|чистка временных|очисти кэш", self.cmd_clean_temp),
            (r"(открой|запусти|открыть|запустить)", self._cmd_open),
            (r"(.+)", lambda t: self.say("Не совсем понял, сэр. Скажите «помощь», чтобы узнать команды.")),
        ]
        for pattern, handler in rules:
            if re.search(pattern, text):
                handler(text)
                return

    def is_stop(self, text):
        return bool(re.search(STOP_WORDS, text))

    def handle(self, phrase):
        if phrase in ("@empty", "", "@error"):
            return
        if phrase == "@unknown":
            self.say("Не расслышал. Повторите еще раз.")
            return
        if phrase == "!!network":
            self.say("Распознавание недоступно без интернета. Проверьте соединение.")
            return
        if self.is_stop(phrase):
            self.say("До встречи, сэр. Системы переводятся в спящий режим.")
            raise SystemExit
        print("-> %s" % phrase)
        self.process(phrase)

    # ---------- режимы ----------
    def _run_textmode(self):
        print("Текстовый режим. Напишите команду или «стоп».")
        while True:
            try:
                phrase = input("Вы: ").strip().lower()
            except EOFError:
                break
            if not phrase:
                continue
            try:
                self.handle(phrase)
            except SystemExit:
                break

    def _run_fallback(self):
        self.say("Режим нажатия недоступен: не установлены библиотеки pynput или SpeechRecognition. "
                 "Установите: py -3 -m pip install pynput SpeechRecognition pyaudio")
        print("(переключаюсь в текстовый режим)")
        self._run_textmode()

    def run(self):
        if TEXT_MODE:
            self._run_textmode()
            return
        if not self.ptt_ok or sr is None:
            self._run_fallback()
            return
        self.recognizer = sr.Recognizer()
        self.say("Джарвис активирован. Зажмите Ctrl и Windows и говорите команду — отпустите, когда закончите. "
                 "Скажите «стоп», чтобы выключить меня.")
        try:
            while True:
                phrase = self.capture_ptt()
                self.handle(phrase)
        except SystemExit:
            pass
        finally:
            self.ptt.stop_listener()


if __name__ == "__main__":
    jarvis = Jarvis()
    jarvis.run()