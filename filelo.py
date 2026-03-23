"""
Filelo v6.1 — PyQt6  (Windows)
IOS 26 + Toss 디자인, 다크/라이트 테마, 검색, 패키지 자동설치
설치: pip install PyQt6
"""
import os, re, json, shutil, threading, datetime, base64, uuid, sys
from pathlib import Path

# ── 패키지 자동 설치 (GUI 진행창) ───────────────────
def _auto_install():
    import subprocess, importlib

    PACKAGES = [
        ("deepl",        "deepl",          "DeepL 번역"),
        ("google.genai", "google-genai",   "Google Gemini AI"),
        ("docx",         "python-docx",    "Word 문서 처리"),
        ("fitz",         "pymupdf",        "PDF 처리"),
        ("PIL",          "Pillow",         "이미지 처리"),
        ("openpyxl",     "openpyxl",       "엑셀 처리"),
        ("cryptography", "cryptography",   "암호화"),
    ]

    missing = []
    for import_name, pip_name, label in PACKAGES:
        try:
            importlib.import_module(import_name)
        except ImportError:
            missing.append((pip_name, label))

    if not missing:
        return  # 다 설치돼 있으면 바로 실행

    # ── tkinter GUI 설치창 (PyQt6 없어도 동작하는 내장 모듈)
    import tkinter as tk
    from tkinter import ttk

    root = tk.Tk()
    root.title("Filelo — 초기 설정")
    root.geometry("460x260")
    root.resizable(False, False)
    root.configure(bg="#12141C")

    # 창 중앙 배치
    root.update_idletasks()
    x = (root.winfo_screenwidth()  - 460) // 2
    y = (root.winfo_screenheight() - 260) // 2
    root.geometry(f"460x260+{x}+{y}")

    tk.Label(root, text="Filelo", font=("Malgun Gothic", 18, "bold"),
             bg="#12141C", fg="#F2F2F7").pack(pady=(24, 2))
    msg = "필요한 패키지를 설치하고 있습니다.\n잠시만 기다려 주세요..."
    tk.Label(root, text=msg,
             font=("Malgun Gothic", 12), bg="#12141C", fg="#8E8E93",
             justify="center").pack(pady=(0, 16))

    status_var = tk.StringVar(value="준비 중...")
    tk.Label(root, textvariable=status_var,
             font=("Malgun Gothic", 11), bg="#12141C", fg="#3182F6").pack()

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Filelo.Horizontal.TProgressbar",
                    troughcolor="#1E2030", background="#3182F6",
                    bordercolor="#12141C", lightcolor="#3182F6", darkcolor="#3182F6")
    pb = ttk.Progressbar(root, style="Filelo.Horizontal.TProgressbar",
                         orient="horizontal", length=380, mode="determinate")
    pb.pack(pady=14)
    pb["maximum"] = len(missing)

    detail_var = tk.StringVar(value="")
    tk.Label(root, textvariable=detail_var,
             font=("Malgun Gothic", 10), bg="#12141C", fg="#636366").pack()

    root.update()

    # 실제 설치
    for i, (pip_name, label) in enumerate(missing):
        status_var.set(f"설치 중... ({i+1}/{len(missing)})")
        detail_var.set(f" {label} ({pip_name})")
        pb["value"] = i
        root.update()

        cmd = [sys.executable, "-m", "pip", "install", "--quiet", pip_name]
        try:
            subprocess.check_call(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except subprocess.CalledProcessError:
            try:
                subprocess.check_call(
                    cmd + ["--break-system-packages"],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
                )
            except subprocess.CalledProcessError:
                detail_var.set(f"️ {label} 설치 실패 (건너뜀)")
                root.update()

    pb["value"] = len(missing)
    status_var.set("✓ 설치 완료! Filelo를 시작합니다...")
    detail_var.set("")
    root.update()
    root.after(1200, root.destroy)
    root.mainloop()

_auto_install()
# ─────────────────────────────────────────────────────

try:    import deepl;                                HAS_DEEPL  = True
except: HAS_DEEPL  = False
try:    from google import genai; from google.genai import types; HAS_GEMINI = True
except: HAS_GEMINI = False
try:    from docx import Document;                   HAS_DOCX   = True
except: HAS_DOCX   = False
try:    import fitz;                                 HAS_FITZ   = True
except: HAS_FITZ   = False
try:    from PIL import Image, ImageDraw, ImageFont; HAS_PIL    = True
except: HAS_PIL    = False
try:    import openpyxl;                             HAS_XL     = True
except: HAS_XL     = False
try:    from rembg import remove as rembg_remove;   HAS_REMBG  = True
except: HAS_REMBG  = False
try:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    HAS_CRYPTO = True
except: HAS_CRYPTO = False

TARGET_MODEL = "gemini-2.0-flash"
# ── 앱 아이콘 (base64 embed) ─────────────────────────
_ICON_B64 = (
    "AAABAAYAEBAAAAAAIAD/AAAAZgAAACAgAAAAACAAYwEAAGUBAAAwMAAAAAAgAMIBAADIAgAAQEAAAAAAIAAaAgAAigQAAICAAAAAACAA0wMAAKQGAAAAAAAAAAAgAMMHAAB3CgAAiVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAxklEQVR4nGNkgAItXYceBhLAtcsHShgYGBhYYJofP7xSTIoBWroODNcuHyhhJEczDMjK6/SyIAs8ffKYoCZpGVkUPhM5NiMDFmyC6LYQbYD9hP8MDAwMDColjzAU3umRw2oAdb1wsICRgYGBNC/QJhBxgbAZL06eu/PLnIEBESZEuyBsxouTDAwMDEYqbCcZGBABjWKAtIwsTv/DbIbRMEBSGKBrZmBgYGC6dvlAiay8Ti8hzejp4Odqr95rlw+UMMIEyM3OAHNKPr7CYraOAAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAABKklEQVR4nO1XPQ6CMBh9JcRLYGVxMrC4OMopOIeJHkMSR4+jI4uLxolF0cQzuOhgqhX7Q9HWxPgWSCm89/3RPAIBenEyFa2/i+16Mamu+SLicrcZ2xDQi5MXIYQnt0VcBQ2jjInwXJMDtwyzbHuuSGUgrqPnQcMo83WbjoeyMUHQpto9yhK8Q173/a/3gFUBdUqg7QHTD5rit0tQB9oSDGeX+313slfuLaYdYwHKDPDkdaATaCzABZQCliOievyCJiXQ9gAv4j+GfwE2YHQWmCKdn/JVcR4A8glRZoDvetMJSOenHAD63VYOyH9S2hIEbdpo/Fjk7CqD1R7QkVsVUK25rAcI4N4XAA9z8vUx9ICbV6NhlLki5a3Z03Fn25yyIIXmVCTk0xDZ8yt7HmuL+5WKlgAAAABJRU5ErkJggolQTkcNChoKAAAADUlIRFIAAAAwAAAAMAgGAAAAVwL5hwAAAYlJREFUeJztmjFugzAUhn9QlUu0lIUpIkuWjM0pco5K5Rggdcxx0pElS6pMLG0SKWfI0g6REXWc4OeY+lnytwFG/n/ee9hgR7jBeDIvb13/L7abVXHt2oPqpBC++/p8G0oUhfFkDkBtJFI0LrkIl0nSvJJNxN0DzuKBc0bIad0a4C5eIJuIAX/EC7om4r7G3Il8e/pdkjSvvI+Achy4xmG/G0oHAODxKSHfox2BocWb9uF9CgUDrmFlwKSISW8hW53ahFUETAgGXOO9AVIRv7z/XJzLim9Sh035TGrfh3YEVOJNoBruw/sUCgZco23g4/XiD4wRtouY9BZSmQhTiTsJBlwTDLgmGHDN3Z+UNlksj/W6Oc0A/QFPOwKqAcvmILZYHmsAmGajGtCftZIiMOSou25Os2k2aiOgC6saoIoHGBmQc163BlgVsclMlU0ETIm3m1WRpHnlWggVseTqfwSA8wq4T1HoLni3EfDFhLxa/yeFuJtQbTVQfqlz2+whHqrWZo8uPmy3+QVVb42rLIMn9wAAAABJRU5ErkJggolQTkcNChoKAAAADUlIRFIAAABAAAAAQAgGAAAAqmlx3gAAAeFJREFUeJztm7FugzAQhg9U5SVaysIUJUuWjM0jVSqPAVIfKR2zZEmViSVJ+xJZ2gGBLAKG+Bx+Gd83BcVY//3cYUNyAd3BfLnJ7hmP4njYpkPHPvUNUIO+nL4/TEWNyXy5qT/3mRH0TJS5EnQXUbzIdSa0GlBdddeDr4jiRU7Ung03BkzhqnfRlg2hejDl4InKjG7eyMOuwb5QGzD1q1/RzIKQyJ/gK1QTpATQAtAEvqW/ShQv8t6tsI7fn4stLUY8v0TsOYxLAB28LQ3e3wPEALQANGIAWgAapw2wsQyy9gE6bIgbA6czwAZiAFoAGjEALQCN9wawlsG3z7/O75L0zJmaiuyVdf5QjDNAF7wNuAYOxfsSEAPQAtCIAWgBaIwN+HrX/rWAzVjLIGsfoDNBHocdQQxAC0AjBqAFoBED0ALQiAFoAWge9sPIGKgvTUy3zs5mQJKeaZXMduqxCcYG6Pb6j34OUIPnmsAqAeQDz764rlfJbLcvrmvOPM6WAFFpAncOJw3ouuGZ3AidNIDoNljTVcDpZdDGW6PweNimVUeFT1TNE86WgC3EAKKymcinMlB7h+oM8MWEZuOUlIB6MPUsaGubk8ZJ3YlT6CYxap1VcbF5Wi1jVvN0kym2z/8DGqOw7BqLkVkAAAAASUVORK5CYIKJUE5HDQoaCgAAAA1JSERSAAAAgAAAAIAIBgAAAMM+YcsAAAOaSURBVHic7Z2xjtpAFEUfqyg/kbA0WyFo0mwZPikSfAZI+SRSpkkDoqIhJD+RJmkyKzBgPGY8M557jrTFsqwx3PPeDAy2B9Yh4+ls2eX2Vdht1ouutv0u5MaqgR8P23nI7asyns7Ofg8pxCDERlzwBB6H4WiyMgsjwkMCEHxaQojQSgCCz4tHRPAWYDydLQk+T4ajycpXgiefOxN+3hwP27nvO69GHYCW3y98hoS7AlD1/aXJkFA7BBB+v2kyJHjNAaA8bgpA9ZfBvS5wVQDCL4s6CS4EIPwyuSUBcwBxzgSg+svmWhegA4jzJgDVr0G1C9ABxHkyo/rVOO0CdABxEEAcBBBnwPivy3A0WdEBxAl6XIAPv38dUz10Fnz4OEy9C2aWaA6gHr5ZPq9BdAFyeeI5kMNrwRxAHAQQBwHEQQBxEEAcBBAHAcRBAHEQQBwEEAcBEpLDghACJCKH8M0SLge3JZcXrhToAOIggDgIIA4CiIMA4iCAOAggDgKIgwDiIIA4CCAOAoiDAOIggDjJloM/f/3b6v9eFj8D74k/++Vz6l0IRpIO0Db8XMhBwlBEF6Dv4TtKkYA5gDgIIA4CiIMA4iCAOAggDgKIgwDiIIA4CCBOdAG+ffG+Yn2WlLIglKQD9F2CUsI3S7gc3FYCjg4OC3MAcRBAHAQQBwHEQQBxEEAcBBAHAcRBAHEQQBwEEAcBxEEAcRBAHAQQBwHEQQBxEEAcBBCnd1cMKYlrJ5mI/YVTBEhA3dlF3N9iicAQEJmmp5aJdQoaBIhINdRPL++/u58m9+8CBEhENfRbEnRNdAEeObCjzweFnFbzrbCv3d51F0jSAdoE2efwcybZuwACzQPmAAn4sf/z6nN7lyBAIqphpwjfjA+CorFfPl9M6JqE3vUHQnQAcRAgIr7VHOPjYASITNNQY60FMAdIgAs3h9XAgZnZeDpbHg/bedRHhqQMR5PVbrNeMASIgwDiIIA4T2Zmu816MRxNVql3BuLgxn8zOoA8bwLQBTQ4rX4zOoA8ZwLQBcqmWv1mdAB5LgSgC5TJteo3u9EBkKAsboVvVjMEIEEZ1IVvxhxAnloB6AL95l71m/1fDr7HeDpbmpmxZNwPXNHeC9+soQAOvjeQP02q/hSvOQBDQt74hm/m2QEcDAl54dPyqzx0/TZESMsjwTuCXMAPEeISInhH0Cs4OhEcCBGG6rwrRPCOTi/hWRUC2hEy8Cr/APZQVkHkecMCAAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAAQAAAAEACAYAAABccqhmAAAHiklEQVR4nO3dsY7bVhqG4TNGsDdhT9y4Cpxmm5TxJQXIXIYN7CV5y222ceDKjeP4JrbZFAYTjYaSSInkOTzf81TJZDIi4Plf/qQk66504Icf37ytfQzk+fjh/UPtY7jVXe0DmMuw07K9RaH5ABh49qz1IDQZgLlD/+Xzb7+udSxwyv3L1+/mfH+LMWgqAFMG37DTsilRaCkE1QNwaegNPHt2KQi1Y1AtAOcG39DTo3MxqBWCKgE4NfwGnwSnQlAjApsGwODD31oIwSYBMPhwWs0QrB6AseE3+PDUWAjWjsCqATgefoMPlx2HYM0IrBIAZ324zVbbwLOlf6Dhh9uNzcwaL4tfdAOw8sPy1rwkWGwDMPywjuNZWnITWCQAhh/WtVYEbg6A4YdtrBGBmwJg+GFbS0fg6gAYfqhjyQhcFQDDD3UtFYHZATD80IYlInDTPQDDD3XdOoOzAnBYGMMPbTicxblbwOQA+Nt5YR/mzOqkALjuh7Zdez9g9j0Aww9tumY2LwbAdT/sx9z7AYu/HRjYj7MBcPaH/ZmzBZwMgLv+0IdzszzpEsDZH/Zl6syOBsDZH/pyaqYvbgDO/rBPU2bXswAQ7EkA3PmHflx6RsAGAMEeBcDZH/pzbguwAUAwAYBgfwXA+g/9OnUZYAOAYAIAwQQAgj0rxfU/JBi7D2ADgGACAMEEAIIJAAS7cwMQsty/fP1u+GcbAAT7rvYB1PL1jy+1D4EKnr+4r30ITYkLgMHPNvz5C8E3UZcAhp+B34VvYgLgD5xjfieCAgA8FREApeeU9N+NiAAA4wQAggkABBMACCYAEEwAIJgAQDABgGACAMEEAIIJAAQTAAgmABBMACCYAEAwAYBgAgDBBACCCQAEEwCipX8+gABAMAEgVvrZvxQBIJTh/ybuswHJZvAfE4CV+YWjZS4BIJgAQDABgGACAMEEAIIJAAQTAAgmABBMACCYAEAwAYBgAgDBBACCCQAEEwAIJgAQTAAgmABAMAGAYAIAwQQAggkABBMACCYAEEwAIJgAQDABgGACAMEEAIIJAASL/Xjwn//1/00e59XD75s8Tus+vf2+9iEwIi4AWw0+jw0hFIK2RF0CGP76bERtiQmA4W+HCLQjJgDAUxEBcPZvjy2gDREBAMYJAAQTAAgmABBMACCYAEAwAYBgAgDBBACCCQAEEwAIJgAQTAAgmABAMAGAYAIAwQQAggkABBMACBYRgH//clf7EDji8wHaEBEAYFxMAGwB7XD2b0dMAEoRgRYY/rbEfTbgEAGfFbAtg9+muAAMttoGnr+43+Rx4BpRlwDAYwIAwQQAggkABBMACCYAEEwAIJgAQDABgGACAMEEAIIJAAQTAAgmABBMACCYAEAwAYBgAgDBBACCCQAEEwAIJgAQTAAgmABAMAGAYAIAwQQAggkABBMACCYAEEwAIJgAQDABgGACAMEEAIIJAAQTAAj2Xe0DgFpePfx+8r99evv9hkdSjwAQ5dzQn/q+nmPgEoAYU4d/qf9vD2wAdG+JAR5+Rm/bgA2Ari199u5tGxAAurXWsPYUAQGAYO4B0KUpZ+l/vvrHf46/9t9P//tp6s/v4X6ADYDuXDv8575+7eO0TgCIc2nI50Rg7yIC8PzFfdTjctrU4U6JQEQAyHFuLZ871FO+f++XATEB2Pps7OzPHsQEoJTthtLwsxdxTwMOw/n1jy+r/WzYi7gADAwrhF0CkG3qi3yu/f49EgCiTB3qhOEvRQDozJSX514a7jnDv/eXAwsAkU4NecqZfxB7ExDShn2MDYDubLWW7339L0UA6NTaw9nD8JciABBNAOjWWmfpXs7+pQgAnVt6WHsa/lI8C0CAYWhveetub4M/ePbxw/uH4V/uX75+V/NgYE3XDnFPw3844x8/vH+wARDlcJh9NqBLAIKlDPk5bgJCMAGAYM9K+XYzYPiCG4HQp+MbgKXYACCaAEAwAYBgfwXAfQDo19j1fyk2AIgmABDsUQBcBkB/Tq3/pdgAINqTANgCoB/nzv6l2AAg2sUA2AJgn6bM7mgAxlYFYL9OzfSkSwBbAOzL1Jk9GQBbAPTh3Cyf3QA8IwD7c+nO/yHPAkCwiwGwBcB+zDn7l3LFBiAC0KZrZnNSAI5LIgLQluOZnHoTf/IG4FkB2Ic5szrrEsD9AGjP3Ov+Qzc9CyACUNetMzg7AO4HQBuuve4/dNUGIAJQ1xLDX8oNlwAiAHUsNfyl3HgPQARgW0sOfykLvBRYBGAbSw9/KQu9F0AEYF1rDH8pC74ZSARgHWsNfyml3C31gwY//Pjm7fHXvnz+7delHwd6N3YSXfoVuYu/HXjsAG0DMM8Ww1/KChvAoeNtwCYAl6258h9bNQCluCSAqbY66x9aPQCljEegFCGAUk5fIm/xDtxNAjAQAvhbzcEfbBqAgRCQrIXBH1QJQCmnI1CKENCnc8+G1foLd6oFYHAuBKWIAft26Snw2n/TVvUAHLoUg1IEgbZNec1L7aE/1FQABlNCcEgUqGHuC9xaGvxBkwE4NDcG0JIWh/5Q8wE4Jgi0rPWBP7a7AIwRBWrY27CP+RM60Kt1byiyhwAAAABJRU5ErkJggg=="
)

def _get_icon():
    """ICO 데이터를 QIcon으로 변환"""
    import base64, tempfile
    try:
        from PyQt6.QtGui import QIcon as _QI
        data = base64.b64decode(_ICON_B64)
        tf = tempfile.NamedTemporaryFile(suffix=".ico", delete=False)
        tf.write(data); tf.flush(); tf.close()
        return _QI(tf.name)
    except Exception:
        return None
# ─────────────────────────────────────────────────────

CONFIG_FILE  = os.path.join(os.path.expanduser("~"), ".filelo_secure.json")
SALT_FILE    = os.path.join(os.path.expanduser("~"), ".filelo_salt")
DATA_FILE    = os.path.join(os.path.expanduser("~"), ".filelo_tasks.json")
CONSENT_FILE = os.path.join(os.path.expanduser("~"), ".filelo_consent.json")
CONSENT_VERSION = "1.0"   # 약관 버전 — 변경 시 재동의 요청

# ── 앱 버전 & 업데이트 설정 ──────────────────────────────
APP_VERSION     = "6.1.0"          # 현재 버전 (Major.Minor.Patch)
GITHUB_OWNER    = "joys000"
GITHUB_REPO     = "Filelo"
GITHUB_API_URL  = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
GITHUB_REL_URL  = f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
UPDATE_CHECK_FILE = os.path.join(os.path.expanduser("~"), ".filelo_update.json")

def _machine_id():
    import platform
    # macOS: ioreg 명령어로 하드웨어 UUID 읽기
    if platform.system() == "Darwin":
        try:
            import subprocess
            result = subprocess.check_output(
                ["ioreg", "-rd1", "-c", "IOPlatformExpertDevice"], text=True
            )
            for line in result.splitlines():
                if "IOPlatformUUID" in line:
                    return line.split('"')[-2]
        except: pass
    # Windows: 레지스트리에서 MachineGuid 읽기
    try:
        import winreg
        k = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography")
        return winreg.QueryValueEx(k, "MachineGuid")[0]
    except: pass
    # Linux
    try:
        with open("/etc/machine-id") as f: return f.read().strip()
    except: pass
    # 최후 fallback
    fb = os.path.join(os.path.expanduser("~"), ".filelo_id")
    if os.path.exists(fb):
        with open(fb) as f: return f.read().strip()
    v = str(uuid.uuid4())
    with open(fb, "w") as f: f.write(v)
    return v

def _salt():
    if os.path.exists(SALT_FILE):
        with open(SALT_FILE, "rb") as f: return f.read()
    s = os.urandom(32); open(SALT_FILE, "wb").write(s); return s

_KEY_CACHE = None
def _key():
    global _KEY_CACHE
    if _KEY_CACHE is not None: return _KEY_CACHE
    if not HAS_CRYPTO: return b""
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=_salt(), iterations=480_000)
    _KEY_CACHE = kdf.derive(("Filelo:" + _machine_id() + ":v1").encode())
    return _KEY_CACHE

def encrypt(pt):
    if not HAS_CRYPTO or not pt: return ""
    n = os.urandom(12)
    return base64.b64encode(n + AESGCM(_key()).encrypt(n, pt.encode(), None)).decode()

def decrypt(enc):
    if not HAS_CRYPTO or not enc: return ""
    try:
        raw = base64.b64decode(enc.encode())
        return AESGCM(_key()).decrypt(raw[:12], raw[12:], None).decode()
    except: return ""

def load_cfg():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return {k: decrypt(v) for k, v in json.load(f).items()}
    except: pass
    return {}

def save_cfg(cfg):
    if not HAS_CRYPTO: return
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({k: encrypt(v) for k, v in cfg.items() if v}, f)
    try: os.chmod(CONFIG_FILE, 0o600); os.chmod(SALT_FILE, 0o600)
    except: pass

def _consent_given() -> bool:
    """이미 동의했는지 확인"""
    try:
        if os.path.exists(CONSENT_FILE):
            with open(CONSENT_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
            return d.get("version") == CONSENT_VERSION and d.get("agreed") is True
    except Exception:
        pass
    return False

def _save_consent():
    """동의 정보 저장"""
    try:
        with open(CONSENT_FILE, "w", encoding="utf-8") as f:
            json.dump({
                "agreed":    True,
                "version":   CONSENT_VERSION,
                "timestamp": datetime.datetime.now().isoformat(),
                "locale":    "ko-KR"
            }, f, ensure_ascii=False)
    except Exception:
        pass


# ── 업데이트 체크 ────────────────────────────────────────
def _parse_version(v: str) -> tuple:
    """'v6.1.0' 또는 '6.1.0' → (6, 1, 0)"""
    v = v.lstrip("vV").strip()
    try:
        parts = [int(x) for x in v.split(".")]
        while len(parts) < 3:
            parts.append(0)
        return tuple(parts[:3])
    except Exception:
        return (0, 0, 0)

def _should_check_update() -> bool:
    """마지막 체크로부터 24시간 지났는지 확인 (너무 잦은 API 호출 방지)"""
    try:
        if os.path.exists(UPDATE_CHECK_FILE):
            with open(UPDATE_CHECK_FILE, "r") as f:
                d = json.load(f)
            last = datetime.datetime.fromisoformat(d.get("last_check", "2000-01-01"))
            if (datetime.datetime.now() - last).total_seconds() < 86400:   # 24시간
                return False
    except Exception:
        pass
    return True

def _save_update_check(latest_ver: str, has_update: bool):
    try:
        with open(UPDATE_CHECK_FILE, "w") as f:
            json.dump({
                "last_check":  datetime.datetime.now().isoformat(),
                "latest":      latest_ver,
                "has_update":  has_update,
            }, f)
    except Exception:
        pass

def _check_update_async(callback):
    """백그라운드 스레드에서 GitHub API 호출 → callback(latest_ver, release_url, has_update)"""
    import threading, urllib.request, urllib.error

    def _worker():
        try:
            req = urllib.request.Request(
                GITHUB_API_URL,
                headers={"User-Agent": f"Filelo/{APP_VERSION}",
                         "Accept":     "application/vnd.github+json"}
            )
            with urllib.request.urlopen(req, timeout=5) as resp:
                data = json.loads(resp.read().decode())

            latest_tag  = data.get("tag_name", "")
            release_url = data.get("html_url", GITHUB_REL_URL)
            latest_ver  = _parse_version(latest_tag)
            current_ver = _parse_version(APP_VERSION)
            has_update  = latest_ver > current_ver

            _save_update_check(latest_tag, has_update)
            callback(latest_tag, release_url, has_update)
        except Exception:
            pass   # 네트워크 없어도 조용히 실패

    threading.Thread(target=_worker, daemon=True).start()


def load_tasks():
    try:
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, "r", encoding="utf-8") as f: return json.load(f)
    except: pass
    return []

def save_tasks(t):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(t, f, ensure_ascii=False, indent=2)

USAGE_FILE = os.path.join(os.path.expanduser("~"), ".filelo_usage.json")

def load_usage():
    try:
        if os.path.exists(USAGE_FILE):
            with open(USAGE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except: pass
    return {}

def save_usage(u):
    try:
        with open(USAGE_FILE, "w", encoding="utf-8") as f:
            json.dump(u, f, ensure_ascii=False, indent=2)
    except: pass

def record_usage(key):
    u = load_usage()
    u[key] = u.get(key, 0) + 1
    save_usage(u)

def get_frequent_features(n=8):
    """사용 빈도 상위 n개 반환. 데이터 없으면 기본값."""
    DEFAULT = [
        ("PDF 변환",    "pdf"),
        ("이미지 처리", "image"),
        ("AI 요약",     "summary"),
        ("번역",        "translate"),
        ("배경 제거",   "rembg"),
        ("엑셀",        "excel"),
        ("OCR",         "ocr"),
        ("트래커",      "tracker"),
    ]
    NAME_MAP = {
        "translate": "번역",      "folder":    "폴더 정리",
        "rename":    "파일명 변경","task_dir":  "과제 폴더",
        "pdf":       "PDF 변환",   "pdfmerge":  "PDF 병합",
        "pdfpwd":    "PDF 암호",   "meta":      "메타 삭제",
        "table2xl":  "표→엑셀",    "image":     "이미지 처리",
        "imgpdf":    "이미지→PDF", "imgext":    "이미지 추출",
        "watermark": "워터마크",   "ocr":       "OCR",
        "rembg":     "배경 제거",  "summary":   "AI 요약",
        "draft":     "AI 초안",    "citation":  "참고문헌",
        "excel":     "엑셀",       "tracker":   "트래커",
        "settings":  "API 설정",
    }
    u = load_usage()
    if not u:
        return DEFAULT[:n]
    sorted_keys = sorted(u.keys(), key=lambda k: -u[k])
    result = [(NAME_MAP.get(k, k), k) for k in sorted_keys[:n]]
    # n개 미만이면 DEFAULT로 채움
    existing_keys = {k for _, k in result}
    for label, key in DEFAULT:
        if len(result) >= n:
            break
        if key not in existing_keys:
            result.append((label, key))
    return result[:n]

_cfg = load_cfg()
DEEPL_KEY  = _cfg.get("deepl_key", "")
GEMINI_KEY = _cfg.get("gemini_key", "")
ai_client  = None

def init_gemini():
    global ai_client
    if HAS_GEMINI and GEMINI_KEY:
        try: ai_client = genai.Client(api_key=GEMINI_KEY); return True
        except: pass
    return False
init_gemini()

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QStackedWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout, QFormLayout,
    QLabel, QPushButton, QLineEdit, QTextEdit,
    QRadioButton, QCheckBox, QComboBox,
    QProgressBar, QScrollArea, QFrame,
    QDialog, QFileDialog, QMessageBox, QListWidget, QListWidgetItem,
    QTreeWidget, QTreeWidgetItem, QHeaderView,
)
from PyQt6.QtCore import (Qt, QThread, pyqtSignal, QTimer,
    QPropertyAnimation, QVariantAnimation, QEasingCurve, QPoint, QRect, QSize)
from PyQt6.QtCore import pyqtProperty
from PyQt6.QtGui import QColor, QPainter, QPainterPath, QBrush, QPen, QFont
from PyQt6.QtWidgets import QGraphicsOpacityEffect

# ── 다크 팔레트
DARK = {
    "bg": "#08090E", "side": "#0D0E14", "card": "#12141C",
    "card2": "#191B25", "input": "#0E1018", "glass": "#1A1C28",
    "overlay": "#14162088",
    "accent": "#3182F6", "accent_h": "#1C6FE8", "accent_s": "#1155CC",
    "accent2": "#6366F1", "accent2_h": "#4F46E5",
    "success": "#05C072", "success_h": "#04A861",
    "warning": "#FF9500", "warning_h": "#E08500",
    "danger": "#FF3B30", "danger_h": "#E03228",
    "text": "#F2F2F7", "text2": "#EBEBF0",
    "sub": "#8E8E93", "sub2": "#636366",
    "border": "#1E2030", "border2": "#2A2D40", "sep": "#1C1E2B",
    "hover": "#1A1D2A", "active": "#1C2540", "pressed": "#111420",
}

# ── 라이트 팔레트
P = DARK.copy()

def make_qss(P):
    return f"""
/* ── 전역 ─────────────────────────────────────────── */
QMainWindow, QWidget {{
    background-color: {P['bg']};
    color: {P['text']};
    font-family: "Segoe UI Emoji", "Pretendard", "Apple SD Gothic Neo",
                 "Malgun Gothic", "맑은 고딕", sans-serif;
    font-size: 13px;
    selection-background-color: {P['accent']};
    selection-color: #ffffff;
}}

/* ── 툴팁 ────────────────────────────────────────────── */
QToolTip {{
    background-color: {P['glass']};
    color: {P['text']};
    border: 1px solid {P['border2']};
    border-radius: 6px;
    padding: 5px 10px;
    font-size: 12px;
}}

/* ── 스크롤바 — 미니멀 ───────────────────────────── */
QScrollArea {{ border: none; background: transparent; }}
QScrollArea > QWidget > QWidget {{ background: transparent; }}
QScrollBar:vertical {{
    background: transparent;
    width: 6px; border: none; margin: 3px 0;
}}
QScrollBar::handle:vertical {{
    background: {P['border2']};
    border-radius: 3px;
    min-height: 32px;
}}
QScrollBar::handle:vertical:hover {{ background: {P['sub']}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: none; }}

/* ── 버튼 — 기본 (Toss Blue 필 모양) ─────────────── */
QPushButton {{
    background-color: {P['accent']};
    color: #ffffff;
    border: none;
    border-radius: 10px;
    padding: 9px 20px;
    font-weight: 600;
    font-size: 13px;
    letter-spacing: -0.1px;
    cursor: pointer;
}}
QPushButton:hover  {{ background-color: {P['accent_h']}; }}
QPushButton:pressed {{ background-color: {P['accent_s']}; }}
QPushButton:disabled {{
    background-color: {P['border']};
    color: {P['sub']};
    border: 1px solid {P['border2']};
    font-weight: 500;
}}

/* ── 버튼 변형 ────────────────────────────────────── */
QPushButton.success {{
    background-color: {P['success']};
    color: #ffffff;
}}
QPushButton.success:hover  {{ background-color: {P['success_h']}; }}
QPushButton.success:pressed {{ background-color: #038050; }}

QPushButton.danger {{
    background-color: {P['danger']};
    color: #ffffff;
}}
QPushButton.danger:hover  {{ background-color: {P['danger_h']}; }}
QPushButton.danger:pressed {{ background-color: #C02820; }}

QPushButton.warning {{
    background-color: {P['warning']};
    color: #ffffff;
}}
QPushButton.warning:hover {{ background-color: {P['warning_h']}; }}

QPushButton.accent2 {{
    background-color: {P['accent2']};
    color: #ffffff;
}}
QPushButton.accent2:hover {{ background-color: {P['accent2_h']}; }}

QPushButton.ghost {{
    background-color: {P['card']};
    color: {P['text']};
    border: 1.5px solid {P['border2']};
}}
QPushButton.ghost:hover {{
    background-color: {P['glass']};
    border-color: {P['accent']}66;
    color: {P['text']};
}}

/* 작은 버튼 */
QPushButton.sm {{
    padding: 6px 14px;
    font-size: 12px;
    border-radius: 8px;
}}
QPushButton.sm.success {{
    background-color: {P['success']};
    color: #ffffff;
    padding: 6px 14px;
    font-size: 12px;
    border-radius: 8px;
}}
QPushButton.sm.success:hover {{ background-color: {P['success_h']}; }}
QPushButton.sm.danger {{
    background-color: {P['danger']};
    color: #ffffff;
    padding: 6px 14px;
    font-size: 12px;
    border-radius: 8px;
}}
QPushButton.sm.danger:hover {{ background-color: {P['danger_h']}; }}

/* ── 입력 필드 ────────────────────────────────────── */
QLineEdit {{
    background-color: {P['input']};
    border: 1.5px solid {P['border2']};
    border-radius: 10px;
    padding: 9px 14px;
    color: {P['text']};
    font-size: 13px;
}}
QLineEdit:focus {{
    border-color: {P['accent']};
    background-color: {P['card2']};
    outline: none;
}}
QLineEdit:disabled {{
    color: {P['sub']};
    background-color: {P['border']};
}}
QLineEdit::placeholder {{
    color: {P['sub2']};
}}

/* ── 텍스트에디터 ─────────────────────────────────── */
QTextEdit {{
    background-color: {P['input']};
    border: 1.5px solid {P['border2']};
    border-radius: 10px;
    padding: 10px;
    color: {P['text']};
    font-size: 13px;
    line-height: 1.6;
}}
QTextEdit:focus {{ border-color: {P['accent']}; }}

/* ── 콤보박스 ─────────────────────────────────────── */
QComboBox {{
    background-color: {P['input']};
    border: 1px solid {P['border2']};
    border-radius: 10px;
    padding: 8px 14px;
    color: {P['text']};
    font-size: 13px;
}}
QComboBox:focus {{ border-color: {P['accent']}; }}
QComboBox::drop-down {{ border: none; width: 28px; }}
QComboBox::down-arrow {{ width: 10px; height: 10px; }}
QComboBox QAbstractItemView {{
    background-color: {P['glass']};
    border: 1px solid {P['border2']};
    border-radius: 10px;
    color: {P['text']};
    selection-background-color: {P['accent']}33;
    selection-color: {P['text']};
    padding: 4px;
    outline: none;
}}

/* ── 레이블 ───────────────────────────────────────── */
QLabel {{
    color: {P['text']};
    font-size: 13px;
    background: transparent;
    border: none;
    padding: 0;
}}
QFormLayout QLabel {{
    color: {P['sub']};
    font-size: 12px;
    font-weight: 500;
    background: transparent;
    border: none;
    padding: 2px 0;
}}

/* ── ScrollArea 내부 widget 배경 강제 ─────────────── */
QScrollArea > QWidget {{
    background: transparent;
}}
QScrollArea > QWidget > QWidget {{
    background: transparent;
}}

/* ── 라디오 / 체크박스 ────────────────────────────── */
QRadioButton, QCheckBox {{
    color: {P['text']};
    spacing: 8px;
    background: transparent;
    font-size: 13px;
    font-weight: 500;
}}
QRadioButton::indicator {{
    width: 18px; height: 18px;
    border: 2px solid {P['border2']};
    border-radius: 9px;
    background: {P['input']};
}}
QRadioButton::indicator:checked {{
    background: {P['accent']};
    border-color: {P['accent']};
    image: none;
}}
QCheckBox::indicator {{
    width: 18px; height: 18px;
    border: 2px solid {P['border2']};
    border-radius: 5px;
    background: {P['input']};
}}
QCheckBox::indicator:checked {{
    background: {P['accent']};
    border-color: {P['accent']};
}}

/* ── 프로그레스바 ─────────────────────────────────── */
QProgressBar {{
    background-color: {P['border2']};
    border: none;
    border-radius: 4px;
    height: 8px;
    color: transparent;
    text-align: center;
}}
QProgressBar::chunk {{
    background: qlineargradient(
        x1:0, y1:0, x2:1, y2:0,
        stop:0 {P['accent']}, stop:1 {P['accent2']}
    );
    border-radius: 4px;
}}

/* ── 트리 / 테이블 ────────────────────────────────── */
QTreeWidget {{
    background-color: {P['card']};
    border: 1px solid {P['border']};
    border-radius: 12px;
    color: {P['text']};
    outline: none;
    font-size: 13px;
}}
QTreeWidget::item {{
    padding: 8px 12px;
    border: none;
    border-radius: 6px;
    margin: 1px 4px;
}}
QTreeWidget::item:selected {{
    background-color: {P['accent']}22;
    color: {P['text']};
}}
QTreeWidget::item:hover {{
    background-color: {P['hover']};
}}
QHeaderView::section {{
    background-color: {P['card']};
    color: {P['sub']};
    font-weight: 600;
    border: none;
    border-bottom: 1px solid {P['border']};
    padding: 10px 12px;
    font-size: 11px;
    letter-spacing: 0.5px;
    text-transform: uppercase;
}}

/* ── QFrame 내부 QLabel — 카드 배경 투과 ──────────── */
QFrame QLabel {{
    background: transparent;
}}

/* ── QFrame 라인 — 기본 border 없앰 ───────────────── */
QFrame[frameShape="4"] {{   /* HLine */
    background: {P['sep']};
    border: none;
    max-height: 1px;
}}
QFrame[frameShape="5"] {{   /* VLine */
    background: {P['sep']};
    border: none;
    max-width: 1px;
}}

/* ── 리스트 위젯 ──────────────────────────────────── */
QListWidget {{
    background-color: {P['card']};
    border: 1px solid {P['border']};
    border-radius: 10px;
    color: {P['text']};
    outline: none;
    font-size: 13px;
}}
QListWidget::item {{
    padding: 5px 10px;
    border-radius: 6px;
    margin: 1px 4px;
}}
QListWidget::item:selected {{
    background-color: {P['accent']}22;
    color: {P['text']};
}}
QListWidget::item:hover {{
    background-color: {P['hover']};
}}

/* ── 메뉴 ─────────────────────────────────────────── */
QMenu {{
    background-color: {P['glass']};
    border: 1px solid {P['border2']};
    border-radius: 12px;
    padding: 6px;
    color: {P['text']};
}}
QMenu::item {{
    padding: 8px 20px;
    border-radius: 8px;
    font-size: 13px;
}}
QMenu::item:selected {{ background-color: {P['hover']}; }}
QMenu::separator {{
    height: 1px;
    background: {P['border']};
    margin: 4px 0;
}}

/* ── setProperty class 매칭 ─────────────────────────── */
QPushButton[class="sm success"] {{
    background-color: #05C072; color: #ffffff;
    padding: 6px 14px; font-size: 12px; border-radius: 8px; font-weight: 600;
}}
QPushButton[class="sm success"]:hover {{ background-color: #04A861; }}
QPushButton[class="sm danger"] {{
    background-color: #FF3B30; color: #ffffff;
    padding: 6px 14px; font-size: 12px; border-radius: 8px; font-weight: 600;
}}
QPushButton[class="sm danger"]:hover {{ background-color: #E03228; }}
QPushButton[class="accent2"] {{
    background-color: #6366F1; color: #ffffff; padding: 9px 20px; font-weight: 600;
}}
QPushButton[class="accent2"]:hover {{ background-color: #4F46E5; }}
QPushButton[class="success"] {{
    background-color: #05C072; color: #ffffff; padding: 9px 20px; font-weight: 600;
}}
QPushButton[class="success"]:hover {{ background-color: #04A861; }}
QPushButton[class="warning"] {{
    background-color: #FF9500; color: #ffffff; padding: 9px 20px; font-weight: 600;
}}
QPushButton[class="warning"]:hover {{ background-color: #E08500; }}
"""

QSS = make_qss(P)  # 다크 모드 고정

# ── 유체적 모션 버튼 베이스 ──────────────────────────────
class FluidButton(QPushButton):
    """hover 색상 페이드 + press 스케일 바운드 버튼"""

    # (bg_normal, bg_hover, bg_press, fg, font_weight, pad_v, pad_h, radius)
    _PRESETS = {
        "accent":     ("#3182F6","#1C6FE8","#1155CC","#ffffff",600,9,20,10),
        "success":    ("#05C072","#04A861","#038050","#ffffff",600,9,20,10),
        "danger":     ("#FF3B30","#E03228","#C02820","#ffffff",600,9,20,10),
        "warning":    ("#FF9500","#E08500","#CC7700","#ffffff",600,9,20,10),
        "accent2":    ("#6366F1","#4F46E5","#3730A3","#ffffff",600,9,20,10),
        "ghost":      (P["card"],P["glass"],P["hover"],P["text"],400,9,20,10),
        "sm_success": ("#05C072","#04A861","#038050","#ffffff",600,6,14,8),
        "sm_danger":  ("#FF3B30","#E03228","#C02820","#ffffff",600,6,14,8),
    }

    def __init__(self, text="", preset="accent", parent=None):
        super().__init__(text, parent)
        st = self._PRESETS.get(preset, self._PRESETS["accent"])
        self._cn, self._ch, self._cp, self._fg, fw, pv, ph, rx = st
        self._rx = rx; self._fw = fw; self._pv = pv; self._ph = ph
        self._cur  = QColor(self._cn)   # 현재 배경색
        self._sc   = 1.0                 # 스케일

        # ── 호버 색상 애니메이션
        self._hanim = QVariantAnimation(self)
        self._hanim.setDuration(180)
        self._hanim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._hanim.valueChanged.connect(self._on_color)

        # ── 스케일 (press bounce)
        self._sanim = QPropertyAnimation(self, b"_sc_prop")

        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setStyleSheet("QPushButton{border:none;background:transparent;color:transparent;}")

    def _on_color(self, c):
        self._cur = c; self.update()

    def _anim_color(self, target, ms=180, ease=QEasingCurve.Type.OutCubic):
        self._hanim.stop()
        self._hanim.setDuration(ms)
        self._hanim.setEasingCurve(ease)
        self._hanim.setStartValue(self._cur)
        self._hanim.setEndValue(QColor(target))
        self._hanim.start()

    @pyqtProperty(float)
    def _sc_prop(self): return self._sc
    @_sc_prop.setter
    def _sc_prop(self, v): self._sc = v; self.update()

    def _anim_scale(self, target, ms, ease):
        self._sanim.stop()
        self._sanim.setDuration(ms)
        self._sanim.setEasingCurve(ease)
        self._sanim.setStartValue(self._sc)
        self._sanim.setEndValue(float(target))
        self._sanim.start()

    def enterEvent(self, e):
        super().enterEvent(e)
        if self.isEnabled(): self._anim_color(self._ch, 180)

    def leaveEvent(self, e):
        super().leaveEvent(e)
        self._anim_color(self._cn, 220)

    def mousePressEvent(self, e):
        super().mousePressEvent(e)
        if not self.isEnabled(): return
        self._hanim.stop(); self._cur = QColor(self._cp); self.update()
        self._anim_scale(0.94, 70, QEasingCurve.Type.OutQuad)

    def mouseReleaseEvent(self, e):
        super().mouseReleaseEvent(e)
        target = self._ch if self.underMouse() else self._cn
        self._anim_color(target, 260, QEasingCurve.Type.OutCubic)
        self._anim_scale(1.0, 320, QEasingCurve.Type.OutBack)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setRenderHint(QPainter.RenderHint.TextAntialiasing)

        w, h = float(self.width()), float(self.height())
        cx, cy = w / 2, h / 2

        if abs(self._sc - 1.0) > 0.001:
            p.translate(cx, cy)
            p.scale(self._sc, self._sc)
            p.translate(-cx, -cy)

        bg = self._cur if self.isEnabled() else QColor(P["border"])
        path = QPainterPath()
        path.addRoundedRect(0, 0, w, h, self._rx, self._rx)
        p.fillPath(path, bg)

        fg = QColor(self._fg) if self.isEnabled() else QColor(P["sub2"])
        p.setPen(fg)
        f = QFont(self.font())
        f.setWeight(QFont.Weight(self._fw))
        p.setFont(f)
        from PyQt6.QtCore import QRectF
        p.drawText(QRectF(0, 0, w, h), Qt.AlignmentFlag.AlignCenter, self.text())
        p.end()

    def sizeHint(self):
        fm = self.fontMetrics()
        tw = fm.horizontalAdvance(self.text())
        return __import__("PyQt6.QtCore", fromlist=["QSize"]).QSize(
            tw + self._ph * 2 + 8,
            fm.height() + self._pv * 2
        )


class Worker(QThread):
    log_sig  = pyqtSignal(str)
    done_sig = pyqtSignal(str, str)
    prog_sig = pyqtSignal(int)
    def __init__(self, fn, *args):
        super().__init__(); self._fn = fn; self._args = args
        self._emitted = False   # done_sig 발행 여부 추적
    def run(self):
        try:
            self._fn(self, *self._args)
        except Exception as e:
            self.log_sig.emit(f"❌ 오류: {e}")
        finally:
            # _fn이 done_sig를 발행하지 않고 return한 경우 자동 발행
            # → spin 버튼이 절대 멈추지 않는 상태 방지
            if not self._emitted:
                self.done_sig.emit("", "err")
    def emit_done(self, msg="", kind="ok"):
        """done_sig 중복 발행 방지"""
        self._emitted = True
        self.done_sig.emit(msg, kind)


class SpinBtn(FluidButton):
    """실행 중 스피너 + 부드러운 hover/press 애니메이션"""
    FRAMES = ["⠋","⠙","⠹","⠸","⠼","⠴","⠦","⠧","⠇","⠏"]

    def __init__(self, text, parent=None):
        super().__init__(text, preset="accent", parent=parent)
        self._orig = text
        self._timer = QTimer(self)
        self._timer.setInterval(80)
        self._timer.timeout.connect(self._spin)
        self._frame = 0

    def start_spin(self):
        self.setEnabled(False); self._frame = 0; self._timer.start()

    def stop_spin(self):
        self._timer.stop(); self.setText(self._orig); self.setEnabled(True)

    def set_enabled_state(self, v):
        self.setEnabled(v)

    def _spin(self):
        f = self.FRAMES[self._frame % len(self.FRAMES)]
        self._frame += 1
        self.setText(f" {f} 처리 중...")



class NavBtn(QPushButton):
    """사이드바 버튼 — hover/active 부드러운 전환"""

    def __init__(self, icon, text, key, parent=None):
        super().__init__(text, parent)
        self.key = key
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedHeight(36)
        self._active = False
        self._hover_p = 0.0   # 0.0=off, 1.0=hover

        # 배경 hover 애니메이션
        self._hanim = QVariantAnimation(self)
        self._hanim.setDuration(160)
        self._hanim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._hanim.valueChanged.connect(lambda v: setattr(self,"_hover_p",v) or self.update())

        # 텍스트 색상 애니메이션
        self._canim = QVariantAnimation(self)
        self._canim.setDuration(160)
        self._canim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._canim.valueChanged.connect(lambda c: setattr(self,"_cur_fg",c) or self.update())

        self._cur_fg = QColor(P["sub"])
        self._cur_acc_w = 0.0   # 왼쪽 액센트 바 너비 (0→2)

        # 액센트 바 너비 애니메이션
        self._aacc = QVariantAnimation(self)
        self._aacc.setDuration(200)
        self._aacc.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._aacc.valueChanged.connect(lambda v: setattr(self,"_cur_acc_w",v) or self.update())

        self._set(False)

    def _set(self, active):
        self._active = active
        # 액센트 바
        target_aw = 2.0 if active else 0.0
        self._aacc.stop()
        self._aacc.setStartValue(self._cur_acc_w)
        self._aacc.setEndValue(target_aw)
        self._aacc.start()
        # 글자 색
        self._canim.stop()
        self._canim.setStartValue(self._cur_fg)
        self._canim.setEndValue(QColor(P["text"] if active else P["sub"]))
        self._canim.start()
        # 배경 (active면 즉시 고정)
        if active:
            self._hanim.stop(); self._hover_p = 1.0; self.update()
        else:
            self._hanim.stop(); self._hover_p = 0.0; self.update()

    def set_active(self, v): self._set(v)

    def enterEvent(self, e):
        super().enterEvent(e)
        if not self._active:
            self._hanim.stop()
            self._hanim.setStartValue(self._hover_p)
            self._hanim.setEndValue(1.0)
            self._hanim.start()
            self._canim.stop()
            self._canim.setStartValue(self._cur_fg)
            self._canim.setEndValue(QColor(P["text"]))
            self._canim.start()

    def leaveEvent(self, e):
        super().leaveEvent(e)
        if not self._active:
            self._hanim.stop()
            self._hanim.setStartValue(self._hover_p)
            self._hanim.setEndValue(0.0)
            self._hanim.start()
            self._canim.stop()
            self._canim.setStartValue(self._cur_fg)
            self._canim.setEndValue(QColor(P["sub"]))
            self._canim.start()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setRenderHint(QPainter.RenderHint.TextAntialiasing)

        w, h = float(self.width()), float(self.height())
        from PyQt6.QtCore import QRectF

        # 배경 (active = 파란 틴트, hover = 미묘한 하이라이트)
        if self._active:
            bg = QColor(P["accent"]); bg.setAlpha(22)
        else:
            bg = QColor(P["hover"]); bg.setAlpha(int(self._hover_p * 200))
        p.fillRect(QRectF(0, 0, w, h), bg)

        # 왼쪽 액센트 바
        if self._cur_acc_w > 0.1:
            acc = QColor(P["accent"])
            p.fillRect(QRectF(0, 0, self._cur_acc_w, h), acc)

        # 텍스트
        p.setPen(self._cur_fg)
        f = QFont(self.font())
        f.setPixelSize(13)
        f.setWeight(QFont.Weight.DemiBold if self._active else QFont.Weight.Normal)
        p.setFont(f)
        p.drawText(QRectF(16, 0, w-28, h), Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, self.text())
        p.end()

def _divider():
    f = QFrame(); f.setFrameShape(QFrame.Shape.HLine)
    f.setStyleSheet(f"background:{P['border']};border:none;max-height:1px;"); return f

def _tip(lines):
    """박스 없는 미니멀 안내 텍스트 — 왼쪽 컬러 바만"""
    w = QWidget()
    w.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
    w.setStyleSheet("QWidget{background:transparent;}")
    h = QHBoxLayout(w); h.setContentsMargins(0, 0, 0, 0); h.setSpacing(0)

    # 왼쪽 컬러 바 (2px)
    bar = QFrame()
    bar.setFixedWidth(2)
    bar.setStyleSheet(f"QFrame{{background:{P['accent']}55;border:none;border-radius:1px;}}")
    h.addWidget(bar)

    # 텍스트 영역
    txt_w = QWidget()
    txt_w.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
    txt_w.setStyleSheet("QWidget{background:transparent;}")
    tv = QVBoxLayout(txt_w); tv.setContentsMargins(12, 0, 0, 0); tv.setSpacing(4)
    for line in lines:
        l = QLabel(line)
        l.setWordWrap(True)
        l.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        l.setStyleSheet(
            f"QLabel{{color:{P['sub']};font-size:12px;"
            f"background:transparent;border:none;padding:0;}}"
        )
        tv.addWidget(l)
    h.addWidget(txt_w, 1)
    return w

def _btn(text, cls="", small=False):
    """FluidButton 헬퍼 — 색상·스케일 애니메이션 내장"""
    preset = cls or "accent"
    if small and cls:
        preset = f"sm_{cls}" if f"sm_{cls}" in FluidButton._PRESETS else cls
    return FluidButton(text, preset=preset)


def _lerp_color(h1: str, h2: str, t: float) -> str:
    """두 #rrggbb 색상을 t(0~1)로 선형 보간"""
    t = max(0.0, min(1.0, t))
    def _p(h): h=h.lstrip('#'); return int(h[0:2],16),int(h[2:4],16),int(h[4:6],16)
    r1,g1,b1=_p(h1); r2,g2,b2=_p(h2)
    return f"#{int(r1+(r2-r1)*t):02x}{int(g1+(g2-g1)*t):02x}{int(b1+(b2-b1)*t):02x}"


def _card():
    f = QFrame()
    f.setStyleSheet(
        f"QFrame{{"
        f"background:{P['card']};"
        f"border:none;"
        f"border-radius:16px;"
        f"}}"
    )
    return f

class DropList(QListWidget):
    """드래그 앤 드롭 지원 파일 목록"""
    files_dropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setMaximumHeight(110)
        self.setStyleSheet(
            f"QListWidget{{background:{P['card']};border:1px solid {P['border2']};"
            f"border-radius:10px;font-size:12px;padding:4px;}}"
            f"QListWidget::item{{padding:4px 10px;border-radius:6px;color:{P['sub']};}}"
            f"QListWidget::item:hover{{background:{P['hover']};}}"
        )

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            # 드래그 진입 — 부드러운 테두리 강조
            self.setStyleSheet(
                f"QListWidget{{background:{P['active']};border:2px solid {P['accent']};"
                f"border-radius:10px;font-size:12px;padding:4px;"
                f"transition:all 150ms;}}"
                f"QListWidget::item{{padding:4px 10px;border-radius:6px;}}"
            )

    def dragLeaveEvent(self, e):
        self.setStyleSheet(
            f"QListWidget{{background:{P['card']};border:1px solid {P['border2']};"
            f"border-radius:10px;font-size:12px;padding:4px;}}"
            f"QListWidget::item{{padding:4px 10px;border-radius:6px;color:{P['sub']};}}"
        )

    def dropEvent(self, e):
        self.dragLeaveEvent(e)
        paths = [u.toLocalFile() for u in e.mimeData().urls()
                 if u.toLocalFile() and os.path.isfile(u.toLocalFile())]
        if paths:
            self.files_dropped.emit(paths)
        e.acceptProposedAction()


def _mk_filelist():
    lw = DropList()
    sel = []
    _ph = QListWidgetItem(" 파일을 추가하거나 이 영역에 드래그하세요")
    _ph.setForeground(QColor(P['sub'])); lw.addItem(_ph)

    def add(paths):
        for p in paths:
            if p not in sel: sel.append(p)
        _ref()

    def clear(): sel.clear(); _ref()

    def _ref():
        lw.clear()
        if sel:
            for p in sel:
                it = QListWidgetItem(f" • {os.path.basename(p)}")
                it.setForeground(QColor(P["text2"])); lw.addItem(it)
        else:
            ph = QListWidgetItem(" 파일을 추가하거나 여기에 드래그하세요")
            ph.setForeground(QColor(P['sub'])); lw.addItem(ph)

    lw.files_dropped.connect(add)

    # 우클릭 → 개별 파일 제거
    lw.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
    def _ctx_menu(pos):
        from PyQt6.QtWidgets import QMenu
        row = lw.indexAt(pos).row()
        if row < 0 or not sel: return
        menu = QMenu(lw)
        menu.setStyleSheet(
            f"QMenu{{background:{P['card']};color:{P['text']};border:1px solid {P['border']};border-radius:8px;padding:4px;}}"
            f"QMenu::item{{padding:7px 20px;border-radius:6px;}}"
            f"QMenu::item:selected{{background:{P['hover']};}}"
        )
        a_del = menu.addAction("✕  이 파일 제거")
        a_clr = menu.addAction("🗑  전체 초기화")
        act = menu.exec(lw.mapToGlobal(pos))
        if act == a_del and row < len(sel):
            sel.pop(row); _ref()
        elif act == a_clr:
            clear()
    lw.customContextMenuRequested.connect(_ctx_menu)

    return lw, sel, add, clear

def _log_color(m):
    return (P["success"] if m.startswith(("","")) else
            P["danger"]  if m.startswith("") else
            P["warning"] if m.startswith("️") else
            P["accent"]  if m.startswith(("","")) else P["sub"])

class Page(QScrollArea):
    def __init__(self, p=None):
        super().__init__(p); self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._r = QWidget()
        self._r.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self._r.setStyleSheet(f"QWidget{{background:{P['bg']};}}")
        self._v = QVBoxLayout(self._r); self._v.setContentsMargins(0,0,0,0); self._v.setSpacing(0)
        self.setWidget(self._r)

    def hdr(self, icon, title, sub):
        w = QWidget()
        w.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        w.setStyleSheet(f"QWidget{{background:{P['bg']};}}")
        v = QVBoxLayout(w); v.setContentsMargins(28,24,28,8); v.setSpacing(4)
        from PyQt6.QtGui import QFont as _QF
        # 아이콘 + 타이틀 행
        hr = QHBoxLayout(); hr.setSpacing(10); hr.setContentsMargins(0,0,0,0)
        ic = QLabel(icon)
        ic.setStyleSheet(f"QLabel{{background:transparent;color:{P['text']};border:none;}}")
        tl = QLabel(title)
        tl.setStyleSheet(
            f"QLabel{{font-size:20px;font-weight:700;color:{P['text']};"
            f"background:transparent;border:none;letter-spacing:-0.5px;}}"
        )
        hr.addWidget(ic); hr.addWidget(tl); hr.addStretch()
        v.addLayout(hr)
        # 서브타이틀
        s = QLabel(sub)
        s.setStyleSheet(
            f"QLabel{{color:{P['sub']};font-size:13px;background:transparent;"
            f"border:none;letter-spacing:-0.1px;}}"
        )
        v.addWidget(s)
        self._v.addWidget(w)

    def tip(self, lines):
        w = QWidget()
        w.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        w.setStyleSheet(f"QWidget{{background:{P['bg']};}}")
        v = QVBoxLayout(w); v.setContentsMargins(28,4,28,14); v.addWidget(_tip(lines))
        self._v.addWidget(w)

    def card(self):
        w = QWidget()
        w.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        w.setStyleSheet(f"QWidget{{background:{P['bg']};}}")
        v = QVBoxLayout(w); v.setContentsMargins(28,0,28,10)
        c = _card(); v.addWidget(c); self._v.addWidget(w)
        inn = QVBoxLayout(c); inn.setContentsMargins(20,16,20,16); inn.setSpacing(10)
        return inn

    def logbox(self, hint="실행 결과가 여기에 표시됩니다"):
        w = QWidget(); w.setStyleSheet(f"background:{P['bg']};")
        v = QVBoxLayout(w); v.setContentsMargins(28,0,28,20)
        t = QTextEdit(); t.setReadOnly(True); t.setMinimumHeight(180)
        t.setPlaceholderText(hint)
        t.setSizePolicy(t.sizePolicy().horizontalPolicy(),
                        __import__("PyQt6.QtWidgets",fromlist=["QSizePolicy"]).QSizePolicy.Policy.Expanding)
        t.setStyleSheet(
            f"QTextEdit{{"
            f"background:{P['card']};"
            f"border:none;"
            f"border-radius:14px;"
            f"font-family:'JetBrains Mono','Cascadia Code','Consolas',monospace;"
            f"font-size:12px;"
            f"color:{P['text']};padding:12px;line-height:1.6;}}"
            f"QTextEdit:focus{{border-color:{P['accent']}44;}}"
        )
        # 우클릭 메뉴: 복사 / 지우기
        t.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        def _ctx(pos):
            from PyQt6.QtWidgets import QMenu
            menu = QMenu(t)
            menu.setStyleSheet(
                f"QMenu{{background:{P['card']};color:{P['text']};border:1px solid {P['border']};border-radius:6px;padding:4px;}}"
                f"QMenu::item{{padding:6px 20px;border-radius:4px;}}"
                f"QMenu::item:selected{{background:{P['hover']};}}"
            )
            a_copy  = menu.addAction("📋  전체 복사")
            a_clear = menu.addAction("🗑  지우기")
            a_save  = menu.addAction("💾  파일로 저장")
            act = menu.exec(t.mapToGlobal(pos))
            if act == a_copy:
                __import__("PyQt6.QtWidgets",fromlist=["QApplication"]).QApplication.clipboard().setText(t.toPlainText())
            elif act == a_clear:
                t.clear()
            elif act == a_save:
                from PyQt6.QtWidgets import QFileDialog as _QFD
                path, _ = _QFD.getSaveFileName(t, "로그 저장", "", "텍스트 (*.txt)")
                if path:
                    with open(path, "w", encoding="utf-8") as _f: _f.write(t.toPlainText())
        t.customContextMenuRequested.connect(_ctx)
        v.addWidget(t,1); self._v.addWidget(w,1); return t

    def result(self, h=200):
        w = QWidget(); w.setStyleSheet(f"background:{P['bg']};")
        v = QVBoxLayout(w); v.setContentsMargins(28,0,28,20)
        t = QTextEdit(); t.setMinimumHeight(h)
        t.setSizePolicy(t.sizePolicy().horizontalPolicy(),
                        __import__("PyQt6.QtWidgets",fromlist=["QSizePolicy"]).QSizePolicy.Policy.Expanding)
        t.setStyleSheet(
            f"QTextEdit{{background:{P['card']};border:none;"
            f"border-radius:14px;"
            f"font-size:13px;color:{P['text']};padding:14px;line-height:1.7;}}"
        )
        v.addWidget(t, 1); self._v.addWidget(w, 1); return t

    def filelist(self, ph="파일을 추가하세요"):
        w = QWidget(); w.setStyleSheet(f"background:{P['bg']};")
        o = QVBoxLayout(w); o.setContentsMargins(28,0,28,8); o.setSpacing(5)
        br = QHBoxLayout(); br.setSpacing(6); br.addStretch()
        lw, sel, add, clear = _mk_filelist(); o.addLayout(br); o.addWidget(lw)
        self._v.addWidget(w); return br, sel, add, clear

    def stretch(self): self._v.addStretch()

    def _guard(self, checks: list) -> bool:
        """실행 전 유효성 검사. 실패 시 toast 표시 후 False 반환.
        checks = [(조건식, "오류 메시지"), ...]  — 조건이 False면 오류"""
        for cond, msg in checks:
            if not cond:
                self.app.toast(msg, "err")
                return False
        return True

    @staticmethod
    def lbl(text):
        """카드 안 행 레이블 — sub 색상"""
        l = QLabel(text)
        l.setStyleSheet(f"color:{P['sub']};font-size:12px;font-weight:500;background:transparent;")
        return l

    def log(self, tb, m):
        tb.append(f'<span style="color:{_log_color(m)}">{m}</span>')
        tb.verticalScrollBar().setValue(tb.verticalScrollBar().maximum())

    def run_worker(self, fn, log_tb, done_cb, btn=None, btn_text=None):
        orig = btn.text() if btn else ""
        if btn: btn.setEnabled(False); btn.setText(" 처리 중...")
        w = Worker(fn)
        w.log_sig.connect(lambda m: self.log(log_tb, m))
        def _done(m, k):
            done_cb(m, k)
            if btn: btn.setEnabled(True); btn.setText(btn_text or orig)
        w.done_sig.connect(_done); w.start(); return w

# ── 홈
class HomePage(QWidget):
    nav_req = pyqtSignal(str)

    def __init__(self, app=None, p=None):
        super().__init__(p)
        self.app = app
        self.setStyleSheet(f"background:{P['bg']};")

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # 수직 중앙 정렬을 위한 stretch
        root.addStretch(2)

        # ── 중앙 컨텐츠
        center = QWidget()
        center.setStyleSheet("background:transparent;")
        cv = QVBoxLayout(center)
        cv.setContentsMargins(80, 0, 80, 0)
        cv.setSpacing(0)
        cv.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # 타이틀
        title = QLabel("무엇을 찾으십니까?")
        title.setStyleSheet(
            f"font-size:36px;font-weight:900;color:{P['text']};"
            f"background:transparent;letter-spacing:-1.5px;"
        )
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cv.addWidget(title)
        cv.addSpacing(8)

        subtitle = QLabel("기능 이름, 파일 형식, 작업 내용을 입력하세요")
        subtitle.setStyleSheet(
            f"font-size:14px;color:{P['sub']};background:transparent;letter-spacing:-0.2px;"
        )
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cv.addWidget(subtitle)
        cv.addSpacing(32)

        # ── 검색 입력창
        search_wrap = QWidget()
        search_wrap.setStyleSheet("background:transparent;")
        sw = QHBoxLayout(search_wrap)
        sw.setContentsMargins(0, 0, 0, 0)
        sw.setSpacing(0)

        self._home_search = QLineEdit()
        self._home_search.setPlaceholderText("PDF 변환, 이미지 배경 제거, 파일명 번역 ...")
        self._home_search.setFixedHeight(52)
        self._home_search.setStyleSheet(
            f"QLineEdit{{"
            f"background:{P['card2']};"
            f"border:2px solid {P['border2']};"
            f"border-radius:14px;"
            f"padding:0 20px;"
            f"font-size:15px;"
            f"color:{P['text']};"
            f"}}"
            f"QLineEdit:focus{{"
            f"border-color:{P['accent']};"
            f"background:{P['card']};"
            f"}}"
        )
        sw.addWidget(self._home_search)
        cv.addWidget(search_wrap)
        cv.addSpacing(20)

        # ── 자주 쓰는 기능 태그 (사용 빈도 기반, 동적 업데이트)
        freq_label = QLabel("자주 사용")
        freq_label.setStyleSheet(
            f"font-size:11px;color:{P['sub2']};background:transparent;"
            f"letter-spacing:0.5px;"
        )
        freq_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cv.addWidget(freq_label)
        cv.addSpacing(8)

        self._tags_wrap = QWidget()
        self._tags_wrap.setStyleSheet("background:transparent;")
        self._tags_layout = QHBoxLayout(self._tags_wrap)
        self._tags_layout.setContentsMargins(0, 0, 0, 0)
        self._tags_layout.setSpacing(8)
        self._tags_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._refresh_tags()
        cv.addWidget(self._tags_wrap)

        root.addWidget(center)
        root.addStretch(3)

        # ── 하단 기능 수 안내
        bottom = QLabel(f"총 21개 기능  ·  파일 관리 · 문서 처리 · 이미지 · AI 도구 · 데이터 · 학습 관리")
        bottom.setStyleSheet(
            f"font-size:11px;color:{P['sub2']};background:transparent;"
        )
        bottom.setAlignment(Qt.AlignmentFlag.AlignCenter)
        bw = QWidget(); bw.setStyleSheet("background:transparent;")
        bwl = QVBoxLayout(bw); bwl.setContentsMargins(0, 0, 0, 24)
        bwl.addWidget(bottom)
        root.addWidget(bw)

    def _refresh_tags(self):
        """사용 빈도 기반으로 태그 버튼을 다시 그림"""
        # 기존 버튼 제거
        while self._tags_layout.count():
            item = self._tags_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for label, key in get_frequent_features(8):
            tag = QPushButton(label)
            tag.setCursor(Qt.CursorShape.PointingHandCursor)
            tag.setFixedHeight(30)
            tag.setStyleSheet(
                f"QPushButton{{"
                f"background:{P['card']};"
                f"color:{P['sub']};"
                f"border:1px solid {P['border2']};"
                f"border-radius:15px;"
                f"padding:0 14px;"
                f"font-size:12px;font-weight:500;"
                f"}}"
                f"QPushButton:hover{{"
                f"background:{P['glass']};"
                f"color:{P['text']};"
                f"border-color:{P['accent']}55;"
                f"}}"
            )
            tag.clicked.connect(lambda _, k=key: self.nav_req.emit(k))
            self._tags_layout.addWidget(tag)

    def get_home_search(self):
        return self._home_search

    def get_tags_widget(self):
        return self._tags_wrap


# ── 파일명 번역
class TranslatePage(Page):
    def __init__(self, app, p=None):
        super().__init__(p); self.app=app
        self.hdr("","파일명 번역","영문 파일명을 DeepL로 한글 변환합니다")
        self.tip([
    "영문으로 된 파일명을 DeepL AI가 자연스러운 한국어로 번역해 줍니다.",
    "① 개별 파일을 직접 추가하거나, 폴더를 선택하면 그 안의 모든 파일을 한 번에 처리합니다.",
    "② 번역은 파일명만 바꾸며 파일 내용은 건드리지 않습니다.",
    "③ 원본 파일명은 복구할 수 없으므로 중요한 파일은 반드시 미리 백업하세요.",
    "④ DeepL API 키가 필요합니다 — ️ 설정에서 등록하세요.",
])
        br,self.files,self._add,self._clr=self.filelist()
        ba=_btn(" 파일 추가","success",True)
        ba.clicked.connect(lambda: self._add(QFileDialog.getOpenFileNames(self,""," ","모든 파일 (*.*)")[0]))
        bc=_btn(" 초기화","danger",True)
        bc.clicked.connect(self._clr); br.addWidget(ba); br.addWidget(bc)
        ci=self.card()
        fr=QHBoxLayout(); fr.addWidget(QLabel("또는 폴더 전체:"))
        self._fl=QLabel("선택된 폴더 없음"); self._fl.setStyleSheet(f"color:{P['sub']};font-size:12px;")
        fr.addWidget(self._fl,1)
        bf=QPushButton(" 폴더 선택"); bf.clicked.connect(self._pf); fr.addWidget(bf); ci.addLayout(fr)
        prog_row = QHBoxLayout()
        self._prog=QProgressBar(); self._prog.setValue(0)
        self._prog_lbl = QLabel("0%")
        self._prog_lbl.setStyleSheet(f"color:{P['sub']};font-size:11px;min-width:32px;background:transparent;")
        self._prog_lbl.setAlignment(Qt.AlignmentFlag.AlignRight)
        prog_row.addWidget(self._prog,1); prog_row.addWidget(self._prog_lbl)
        ci.addLayout(prog_row)
        self._btn=SpinBtn(" 번역 시작"); self._btn.clicked.connect(self._run); ci.addWidget(self._btn)
        self._log=self.logbox(); self.stretch(); self._fd=""

    def _pf(self):
        d=QFileDialog.getExistingDirectory(self,"폴더 선택")
        if d: self._fd=d; self._fl.setText(d)

    def _run(self):
        if not self._guard([
            (HAS_DEEPL,       "DeepL 패키지가 설치되지 않았습니다"),
            (bool(DEEPL_KEY), "DeepL API 키가 없습니다 — ⚙️ 설정에서 등록해 주세요"),
            (bool(self.files) or bool(self._fd),
                              "번역할 파일 또는 폴더를 선택해 주세요"),
        ]): return
        self._btn.start_spin()
        w=Worker(self._exec)
        w.log_sig.connect(lambda m: self.log(self._log,m))
        w.prog_sig.connect(self._prog.setValue)
        w.prog_sig.connect(lambda v: self._prog_lbl.setText(f"{v}%"))
        w.done_sig.connect(lambda m,k: (self.app.toast(m,k),self._btn.stop_spin(),self._prog_lbl.setText('✓ 완료')))
        w.start(); self._w=w

    def _exec(self,w):
        if not HAS_DEEPL: w.log_sig.emit(" DeepL 패키지 없음"); return
        if not DEEPL_KEY: w.log_sig.emit(" API 키 없음 — ️ 설정에서 등록"); return
        tr=deepl.Translator(DEEPL_KEY)
        if self.files:
            targets=[(os.path.dirname(f),os.path.basename(f)) for f in self.files if re.search(r'[a-zA-Z]',os.path.basename(f))]
        elif self._fd:
            targets=[(self._fd,fn) for fn in os.listdir(self._fd)
                     if os.path.isfile(os.path.join(self._fd,fn)) and re.search(r'[a-zA-Z]',fn) and not fn.startswith('.')]
        else: w.log_sig.emit("️ 파일 또는 폴더 선택 필요"); return
        if not targets: w.log_sig.emit("️ 번역할 영문 파일명 없음"); return
        for i,(folder,fn) in enumerate(targets):
            name,ext=os.path.splitext(fn)
            clean=re.sub(r'[\\/*?:"<>|]',"",tr.translate_text(name,target_lang="KO").text).strip()
            os.rename(os.path.join(folder,fn),os.path.join(folder,clean+ext))
            w.log_sig.emit(f" {fn} → {clean+ext}")
            w.prog_sig.emit(int((i+1)/len(targets)*100))
        w.done_sig.emit(f"파일명 번역 완료! ({len(targets)}개)","ok")

# ── 폴더 자동 정리
class FolderPage(Page):
    EM={"이미지":[".jpg",".jpeg",".png",".gif",".bmp",".webp",".svg",".ico"],
        "문서":[".pdf",".docx",".doc",".hwp",".hwpx",".txt",".pptx",".xlsx",".csv"],
        "영상":[".mp4",".mov",".avi",".mkv",".wmv",".flv"],
        "음악":[".mp3",".wav",".flac",".aac",".ogg"],
        "압축파일":[".zip",".rar",".7z",".tar",".gz"],
        "코드":[".py",".js",".ts",".html",".css",".java",".c",".cpp"],"기타":[]}
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("","폴더 자동 정리","확장자별로 파일을 자동 분류합니다")
        self.tip([
    "폴더 안에 뒤섞인 파일들을 확장자 기준으로 자동 분류합니다.",
    "① 정리할 폴더를 선택하면 하위에 카테고리 폴더가 자동 생성됩니다.",
    "② 분류 기준: 이미지(jpg·png 등) / 문서(pdf·docx·hwp 등) / 영상(mp4·mkv 등)",
    " 음악(mp3·wav 등) / 압축파일(zip·rar 등) / 코드(py·js·html 등) / 기타",
    "③ 이미 같은 이름의 파일이 있으면 덮어쓸 수 있으니 주의하세요.",
])
        ci=self.card()
        fr=QHBoxLayout()
        self._lbl=QLabel("정리할 폴더를 선택하세요"); self._lbl.setStyleSheet(f"color:{P['sub']};font-size:12px;")
        fr.addWidget(self._lbl,1); b=QPushButton(" 폴더 선택"); b.clicked.connect(self._pick); fr.addWidget(b); ci.addLayout(fr)
        self._btn=SpinBtn(" 정리 시작"); self._btn.set_enabled_state(False); self._btn.clicked.connect(self._run); ci.addWidget(self._btn)
        self._log=self.logbox(); self.stretch(); self._d=""
    def _pick(self):
        d=QFileDialog.getExistingDirectory(self,"폴더 선택")
        if d:
            self._d=d
            try:
                cnt=len([f for f in os.listdir(d) if os.path.isfile(os.path.join(d,f)) and not f.startswith(".")])
                self._lbl.setText(f"{d}  ({cnt}개 파일)")
            except:
                self._lbl.setText(d)
            self._btn.set_enabled_state(True)
    def _run(self):
        if not self._guard([
            (bool(self._d), "정리할 폴더를 먼저 선택해 주세요"),
        ]): return
        self._btn.start_spin()
        w=Worker(self._exec)
        w.log_sig.connect(lambda m: self.log(self._log,m))
        w.done_sig.connect(lambda m,k: (self.app.toast(m,k),self._btn.stop_spin()))
        w.start(); self._w=w
    def _exec(self,w):
        moved=0
        for fn in os.listdir(self._d):
            if not os.path.isfile(os.path.join(self._d,fn)) or fn.startswith('.'): continue
            ext=Path(fn).suffix.lower(); cat="기타"
            for c,exts in self.EM.items():
                if ext in exts: cat=c; break
            dest=os.path.join(self._d,cat); os.makedirs(dest,exist_ok=True)
            shutil.move(os.path.join(self._d,fn),os.path.join(dest,fn))
            w.log_sig.emit(f" [{cat}] {fn}"); moved+=1
        w.log_sig.emit(f" 완료! 총 {moved}개"); w.done_sig.emit(f"{moved}개 파일 정리 완료!","ok")

# ── 파일명 규칙 변경
class RenamePage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("️","파일명 규칙 일괄 변경","폴더 안의 파일명을 정해진 규칙으로 일괄 변경합니다")
        self.tip([
            "폴더 안의 파일명을 내가 정한 규칙에 맞춰 한 번에 바꿔줍니다.",
            "① 규칙 패턴에 쓸 수 있는 변수:",
            " {date} → 오늘 날짜(20240115) {num} → 순서 번호 {name} → 원본 파일명",
            " 예) {date}_{num:03d}_{name} 결과: 20240115_001_보고서.pdf",
            "② 확장자 필터를 입력하면 특정 형식(.jpg, .pdf 등)만 골라서 변경합니다.",
            "③ 적용 전 반드시 미리보기로 결과를 확인하세요. 되돌리기가 어렵습니다.",
        ])
        ci=self.card(); g=QFormLayout(); g.setSpacing(8)
        pr=QHBoxLayout(); self._path=QLineEdit(); self._path.setPlaceholderText("폴더 경로"); pr.addWidget(self._path,1)
        bp=QPushButton("선택"); bp.setFixedWidth(60); bp.clicked.connect(self._pick); pr.addWidget(bp); g.addRow("폴더 선택:",pr)
        self._pat=QLineEdit("{date}_{num:03d}_{name}"); g.addRow("규칙 패턴:",self._pat)
        hr=QHBoxLayout(); self._st=QLineEdit("1"); self._st.setFixedWidth(60); hr.addWidget(self._st)
        hr.addWidget(QLabel(" 확장자 필터:")); self._ext=QLineEdit(); self._ext.setPlaceholderText(".jpg"); self._ext.setFixedWidth(80); hr.addWidget(self._ext); hr.addStretch(); g.addRow("시작 번호:",hr)
        ci.addLayout(g); br=QHBoxLayout()
        b1=_btn(" 미리보기","accent2"); b1.clicked.connect(self._prev); br.addWidget(b1)
        b2=_btn(" 적용","success"); b2.clicked.connect(self._apply); br.addWidget(b2); br.addStretch(); ci.addLayout(br)
        self._log=self.logbox(); self.stretch(); self._folder=""
    def _pick(self):
        d=QFileDialog.getExistingDirectory(self,"폴더 선택")
        if d: self._folder=d; self._path.setText(d)
    def _pairs(self):
        if not self._folder: return None,[]
        ef=self._ext.text().strip().lower()
        files=sorted([f for f in os.listdir(self._folder) if os.path.isfile(os.path.join(self._folder,f)) and not f.startswith('.') and (not ef or Path(f).suffix.lower()==ef)])
        pat=self._pat.text().strip(); today=datetime.date.today().strftime("%Y%m%d")
        try: start=int(self._st.text().strip())
        except: start=1
        return self._folder,[(fn,(lambda fn=fn,i=i: pat.format(num=start+i,name=Path(fn).stem,date=today)+Path(fn).suffix if pat else fn)()) for i,fn in enumerate(files)]
    def _prev(self):
        folder,pairs=self._pairs()
        if not folder: QMessageBox.warning(self,"안내","폴더를 먼저 선택하세요."); return
        self._log.clear(); self.log(self._log,"[ 미리보기 — 실제 변경 없음 ]")
        for o,n in pairs: self.log(self._log,f" {o} → {n}")
    def _apply(self):
        folder,pairs=self._pairs()
        if not folder: QMessageBox.warning(self,"안내","폴더를 먼저 선택하세요."); return
        if QMessageBox.question(self,"확인",f"총 {len(pairs)}개 파일명을 변경합니다.")!=QMessageBox.StandardButton.Yes: return
        self._log.clear()
        for o,n in pairs:
            try: os.rename(os.path.join(folder,o),os.path.join(folder,n)); self.log(self._log,f" {o} → {n}")
            except Exception as e: self.log(self._log,f" {o}: {e}")
        self.app.toast("파일명 일괄 변경 완료!","ok")

# ── 과제 폴더 생성
class TaskDirPage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("","과제 폴더 생성","과목명 입력 시 날짜·제출본·참고자료 구조 자동 생성")
        self.tip([
    "과목별로 체계적인 과제 폴더 구조를 자동으로 만들어줍니다.",
    "① 저장할 위치를 선택하고 과목명을 쉼표(,)로 구분해 입력하세요.",
    " 예) 소방학개론, 위험물안전관리, 화재조사론",
    "② 각 과목 아래에 날짜별 / 제출본 / 참고자료 / 필기 폴더가 자동 생성됩니다.",
    "③ 날짜별 폴더 안에는 오늘 날짜(YYYY-MM-DD) 폴더도 함께 만들어집니다.",
])
        ci=self.card(); g=QFormLayout(); g.setSpacing(8)
        br=QHBoxLayout(); self._base=QLineEdit(); self._base.setPlaceholderText("폴더 경로"); br.addWidget(self._base,1)
        bp=QPushButton("선택"); bp.setFixedWidth(60); bp.clicked.connect(lambda:(lambda d:self._base.setText(d) if d else None)(QFileDialog.getExistingDirectory(self,"저장 위치"))); br.addWidget(bp); g.addRow("저장 위치:",br)
        self._subj=QLineEdit(); self._subj.setPlaceholderText("소방학개론, 위험물, 화재조사론"); g.addRow("과목명:",self._subj); ci.addLayout(g)
        rb=_btn(" 폴더 생성","success"); rb.clicked.connect(self._create); ci.addWidget(rb)
        self._log=self.logbox(); self.stretch()
    def _create(self):
        base=self._base.text().strip(); raw=self._subj.text().strip()
        if not base or not raw: QMessageBox.warning(self,"입력 오류","저장 위치와 과목명을 모두 입력하세요."); return
        today=datetime.datetime.now().strftime("%Y-%m-%d")
        for subj in [s.strip() for s in raw.split(",") if s.strip()]:
            for sd in ["날짜별","제출본","참고자료","필기"]: os.makedirs(os.path.join(base,subj,sd),exist_ok=True)
            os.makedirs(os.path.join(base,subj,"날짜별",today),exist_ok=True)
            self.log(self._log,f" [{subj}] 폴더 생성 완료")
        self.app.toast("과제 폴더 생성 완료!","ok")

# ── PDF 변환 & 추출
class PdfPage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("","PDF 변환 & 추출","PDF에서 텍스트 또는 이미지를 추출합니다")
        self.tip([
    "PDF 파일에서 텍스트나 이미지를 꺼내 별도 파일로 저장합니다.",
    "① 텍스트 추출: PDF의 모든 글자를 .txt 파일로 저장합니다.",
    " 스캔 PDF(사진으로 만든 PDF)는 텍스트 추출이 안 될 수 있습니다.",
    "② 이미지 추출: PDF 안에 삽입된 그림·사진을 PNG/JPG 파일로 저장합니다.",
    " 결과 파일은 원본 PDF와 같은 폴더에 저장됩니다.",
])
        br,self.files,self._add,self._clr=self.filelist()
        ba=_btn(" 추가","success",True)
        ba.clicked.connect(lambda: self._add(QFileDialog.getOpenFileNames(self,""," ","PDF (*.pdf)")[0]))
        bc=_btn(" 초기화","danger",True)
        bc.clicked.connect(self._clr); br.addWidget(ba); br.addWidget(bc)
        ci=self.card()
        mr=QHBoxLayout(); mr.addWidget(QLabel("추출 모드:"))
        self._txt=QRadioButton("텍스트 (TXT)"); self._txt.setChecked(True)
        self._img=QRadioButton("이미지 (PNG/JPG)")
        mr.addWidget(self._txt); mr.addWidget(self._img); mr.addStretch(); ci.addLayout(mr)
        self._btn=SpinBtn(" 추출 시작"); self._btn.clicked.connect(self._run); ci.addWidget(self._btn)
        self._log=self.logbox(); self.stretch()
    def _run(self):
        w=Worker(self._exec)
        w.log_sig.connect(lambda m: self.log(self._log,m))
        w.done_sig.connect(lambda m,k: (self.app.toast(m,k),self._btn.stop_spin()))
        if not self._guard([
            (HAS_FITZ,         "PyMuPDF(pymupdf) 패키지가 필요합니다"),
            (bool(self.files), "처리할 PDF/DOCX 파일을 추가해 주세요"),
        ]): return
        self._btn.start_spin(); w.start(); self._w=w
    def _exec(self,w):
        if not HAS_FITZ: w.log_sig.emit(" pymupdf 필요"); return
        if not self.files: w.log_sig.emit("️ 파일 추가 먼저"); return
        mode="text" if self._txt.isChecked() else "image"
        for path in self.files:
            w.log_sig.emit(f" {os.path.basename(path)}")
            doc=fitz.open(path); base=os.path.splitext(path)[0]
            if mode=="text":
                with open(base+"_추출텍스트.txt","w",encoding="utf-8") as f: f.write("".join(pg.get_text() for pg in doc))
                w.log_sig.emit(" 텍스트 저장")
            else:
                d=base+"_이미지들"; os.makedirs(d,exist_ok=True); cnt=0
                for i,pg in enumerate(doc):
                    for j,img in enumerate(pg.get_images(full=True)):
                        bi=doc.extract_image(img[0])
                        with open(os.path.join(d,f"p{i+1}_{j+1}.{bi['ext']}"),"wb") as f: f.write(bi["image"]); cnt+=1
                w.log_sig.emit(f" 이미지 {cnt}개 추출")
        w.done_sig.emit("PDF 추출 완료!","ok")

# ── PDF 합치기/쪼개기
class PdfMergePage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("","PDF 합치기 / 쪼개기","여러 PDF 병합 또는 원하는 페이지만 분리")
        self.tip([
    "여러 PDF를 하나로 합치거나, 필요한 페이지만 골라 새 PDF로 저장합니다.",
    " 합치기: 파일을 순서대로 추가하면 추가된 순서 그대로 합쳐집니다.",
    " 저장 파일명을 입력하지 않으면 합본.pdf로 저장됩니다.",
    " 쪼개기: 원하는 페이지 범위를 입력합니다.",
    " 연속: 1-5 (1~5페이지) 개별: 1,3,7 (1·3·7페이지만) 혼합: 1-3,7,10-12",
    " 결과 파일은 원본 파일명에 _분리.pdf가 붙어 저장됩니다.",
])
        mc=self.card(); mc.addWidget(QLabel(" PDF 합치기",styleSheet=f"font-weight:700;font-size:14px;color:{P['text']};background:transparent;"))
        br,self.mf,self._ma,self._mc=self.filelist()
        ba=_btn(" 추가","success",True)
        ba.clicked.connect(lambda: self._ma(QFileDialog.getOpenFileNames(self,""," ","PDF (*.pdf)")[0]))
        bcl=_btn(" 초기화","danger",True)
        bcl.clicked.connect(self._mc); br.addWidget(ba); br.addWidget(bcl)
        nr=QHBoxLayout(); nr.addWidget(QLabel("저장 파일명:"))
        self._mn=QLineEdit("최종제출본"); self._mn.setFixedWidth(200); nr.addWidget(self._mn); nr.addWidget(QLabel(".pdf")); nr.addStretch(); mc.addLayout(nr)
        bm=SpinBtn(" 합치기 실행"); bm.clicked.connect(self._mrg); self._bm=bm; mc.addWidget(bm)
        sc=self.card(); sc.addWidget(QLabel("️ PDF 쪼개기",styleSheet=f"font-weight:700;font-size:14px;color:{P['text']};background:transparent;"))
        sr=QHBoxLayout(); self._sl=QLabel("쪼갤 PDF 파일 선택"); self._sl.setStyleSheet(f"color:{P['sub']};"); sr.addWidget(self._sl,1)
        bs=QPushButton(" 파일 선택"); bs.clicked.connect(self._sp); sr.addWidget(bs); sc.addLayout(sr)
        rr=QHBoxLayout(); rr.addWidget(QLabel("페이지 범위:")); self._sr=QLineEdit(); self._sr.setPlaceholderText("예) 1-5 또는 1,3,5"); rr.addWidget(self._sr,1); sc.addLayout(rr)
        bse=SpinBtn("️ 쪼개기 실행"); bse.setStyleSheet(f"QPushButton{{background:{P['warning']};color:#fff;border:none;border-radius:10px;padding:9px 20px;font-weight:600;}}QPushButton:hover{{background:{P['warning_h']};}}"); bse.clicked.connect(self._spl); self._bse=bse; sc.addWidget(bse)
        self._log=self.logbox(); self.stretch(); self._sf=""
    def _sp(self):
        f,_=QFileDialog.getOpenFileName(self,""," ","PDF (*.pdf)")
        if f: self._sf=f; self._sl.setText(os.path.basename(f))
    def _lm(self,m): self.log(self._log,m)
    def _mrg(self):
        if not self._guard([
            (HAS_FITZ,          "PyMuPDF 패키지가 필요합니다"),
            (len(self.mf) >= 2, "PDF 파일을 2개 이상 추가해 주세요"),
        ]): return
        w=Worker(self._me); w.log_sig.connect(self._lm); w.done_sig.connect(lambda m,k: self.app.toast(m,k)); w.start(); self._w=w
    def _me(self,w):
        if not HAS_FITZ: w.log_sig.emit(" pymupdf 필요"); return
        if len(self.mf)<2: w.log_sig.emit("️ PDF 2개 이상 추가"); return
        mg=fitz.open()
        for f in self.mf: d=fitz.open(f); mg.insert_pdf(d); w.log_sig.emit(f" {os.path.basename(f)} ({d.page_count}p)")
        out=self._mn.text().strip() or "합본"; op=os.path.join(os.path.dirname(self.mf[0]),out+".pdf")
        mg.save(op); w.log_sig.emit(f" {out}.pdf (총 {mg.page_count}p)"); w.done_sig.emit(f"{out}.pdf 저장 완료!","ok")
    def _spl(self):
        if not self._guard([
            (HAS_FITZ,                      "PyMuPDF 패키지가 필요합니다"),
            (bool(self._sf),                "쪼갤 PDF 파일을 선택해 주세요"),
            (bool(self._sr.text().strip()), "페이지 범위를 입력해 주세요 (예: 1-3,7)"),
        ]): return
        w=Worker(self._se); w.log_sig.connect(self._lm); w.done_sig.connect(lambda m,k: self.app.toast(m,k)); w.start(); self._w2=w
    def _se(self,w):
        if not HAS_FITZ or not self._sf: return
        rs=self._sr.text().strip()
        if not rs: w.log_sig.emit("️ 페이지 범위 입력"); return
        pages=set()
        for pt in rs.split(","):
            pt=pt.strip()
            if "-" in pt: s,e=pt.split("-"); pages.update(range(int(s)-1,int(e)))
            else: pages.add(int(pt)-1)
        src=fitz.open(self._sf); out=fitz.open()
        for pg in sorted(pages):
            if 0<=pg<src.page_count: out.insert_pdf(src,from_page=pg,to_page=pg)
        op=os.path.splitext(self._sf)[0]+"_분리.pdf"; out.save(op)
        w.log_sig.emit(f" {os.path.basename(op)} ({len(pages)}p)"); w.done_sig.emit("PDF 쪼개기 완료!","ok")

# ── 나머지 단순 페이지들을 위한 팩토리
def _simple(app, icon, title, sub, tips, ft, exec_fn, btn_text="실행"):
    class _P(Page):
        def __init__(self,p=None):
            super().__init__(p); self.app=app
            self.hdr(icon,title,sub); self.tip(tips)
            self.br,self.files,self._add,self._clr=self.filelist()
            ba=_btn(" 추가","success",True)
            ba.clicked.connect(lambda: self._add(QFileDialog.getOpenFileNames(self,""," ",ft)[0]))
            bc=_btn(" 초기화","danger",True)
            bc.clicked.connect(self._clr); self.br.addWidget(ba); self.br.addWidget(bc)
            ci=self.card(); self._btn=SpinBtn(btn_text); self._btn.clicked.connect(self._run); ci.addWidget(self._btn)
            self._log=self.logbox(); self.stretch(); self._extra_init(ci)
        def _extra_init(self,ci): pass
        def _run(self):
            if not self._guard([
                (bool(self.files), "처리할 파일을 먼저 추가해 주세요"),
            ]): return
            self._btn.start_spin()
            w=Worker(exec_fn,self)
            w.log_sig.connect(lambda m: self.log(self._log,m))
            w.done_sig.connect(lambda m,k: (self.app.toast(m,k) if m else None,self._btn.stop_spin()))
            w.start(); self._w=w
    return _P()

# ── 메타데이터 삭제
def _meta_exec(w,page):
    if not page.files: w.log_sig.emit("️ 파일 추가 먼저"); return
    for fp in page.files:
        ext=Path(fp).suffix.lower()
        try:
            if ext==".pdf" and HAS_FITZ:
                doc=fitz.open(fp); doc.set_metadata({}); op=os.path.splitext(fp)[0]+"_clean.pdf"; doc.save(op); w.log_sig.emit(f" PDF: {os.path.basename(op)}")
            elif ext==".docx" and HAS_DOCX:
                doc=Document(fp); cp=doc.core_properties
                for a in ["author","last_modified_by","title","subject","keywords","comments"]:
                    try: setattr(cp,a,"")
                    except: pass
                op=os.path.splitext(fp)[0]+"_clean.docx"; doc.save(op); w.log_sig.emit(f" DOCX: {os.path.basename(op)}")
            else: w.log_sig.emit(f"️ {os.path.basename(fp)}: 지원 안 됨")
        except Exception as e: w.log_sig.emit(f" {os.path.basename(fp)}: {e}")
    w.done_sig.emit("메타데이터 삭제 완료!","ok")

# ── 표 → 엑셀
def _t2xl_exec(w,page):
    if not HAS_XL: w.log_sig.emit(" openpyxl 필요"); return
    if not page.files: w.log_sig.emit("️ 파일 추가 먼저"); return
    for fp in page.files:
        ext=Path(fp).suffix.lower()
        try:
            wb=openpyxl.Workbook(); wb.remove(wb.active); idx=0
            if ext==".docx" and HAS_DOCX:
                for ti,tbl in enumerate(Document(fp).tables):
                    ws=wb.create_sheet(f"표{ti+1}"); idx+=1
                    for row in tbl.rows: ws.append([c.text.strip() for c in row.cells])
            elif ext==".pdf" and HAS_FITZ:
                doc=fitz.open(fp)
                for pi,pg in enumerate(doc):
                    for ti,tab in enumerate(pg.find_tables()):
                        ws=wb.create_sheet(f"p{pi+1}_표{ti+1}"); idx+=1
                        for row in tab.extract(): ws.append([str(c) if c else "" for c in row])
            if idx==0: w.log_sig.emit(f"️ {os.path.basename(fp)}: 표 없음"); continue
            op=os.path.splitext(fp)[0]+"_표변환.xlsx"; wb.save(op); w.log_sig.emit(f" {os.path.basename(op)} ({idx}시트)")
        except Exception as e: w.log_sig.emit(f" {os.path.basename(fp)}: {e}")
    w.done_sig.emit("표 → 엑셀 변환 완료!","ok")

# ── 이미지 일괄 추출
def _imgext_exec(w,page):
    if not page.files: w.log_sig.emit("️ 파일 추가 먼저"); return
    for fp in page.files:
        ext=Path(fp).suffix.lower(); out=os.path.splitext(fp)[0]+"_이미지들"; os.makedirs(out,exist_ok=True); cnt=0
        try:
            if ext==".pdf" and HAS_FITZ:
                doc=fitz.open(fp)
                for i,pg in enumerate(doc):
                    for j,img in enumerate(pg.get_images(full=True)):
                        bi=doc.extract_image(img[0])
                        with open(os.path.join(out,f"p{i+1}_{j+1}.{bi['ext']}"),"wb") as f: f.write(bi["image"]); cnt+=1
            elif ext==".docx" and HAS_DOCX:
                doc=Document(fp)
                for rel in doc.part.rels.values():
                    if "image" in rel.reltype:
                        ip=rel.target_part; ie=ip.content_type.split("/")[-1]
                        with open(os.path.join(out,f"img_{cnt+1}.{ie}"),"wb") as f: f.write(ip.blob); cnt+=1
            w.log_sig.emit(f" {os.path.basename(fp)}: {cnt}개 → {os.path.basename(out)}")
        except Exception as e: w.log_sig.emit(f" {os.path.basename(fp)}: {e}")
    w.done_sig.emit("이미지 추출 완료!","ok")


# ── 이미지 일괄 처리
class ImagePage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("️","이미지 일괄 처리","폴더 전체 또는 선택한 이미지를 리사이즈·포맷 변환")
        self.tip([
    "이미지 크기 조절과 파일 형식 변환을 한 번에 일괄 처리합니다.",
    "① 최대 너비(px): 입력한 너비보다 큰 이미지만 줄이며 가로세로 비율은 유지됩니다.",
    " 비워두면 크기 변경 없이 포맷 변환만 합니다.",
    "② 포맷 변환: JPEG(압축률 높음) / PNG(투명 지원) / WEBP(최신 고효율 형식)",
    "③ JPEG 품질: 1~100, 숫자가 클수록 고화질·용량 큼 (기본값 85 권장)",
    "④ 결과는 원본 폴더 안 _처리완료 폴더에 저장됩니다. 원본은 유지됩니다.",
])
        br,self.files,self._add,self._clr=self.filelist()
        ba=_btn(" 추가","success",True)
        ba.clicked.connect(lambda: self._add(QFileDialog.getOpenFileNames(self,""," ","이미지 (*.jpg *.jpeg *.png *.bmp *.webp *.tiff)")[0]))
        bc=_btn(" 초기화","danger",True); bc.clicked.connect(self._clr)
        br.addWidget(ba); br.addWidget(bc)
        ci=self.card()
        fr=QHBoxLayout(); self._fl=QLabel("또는 폴더 전체 선택"); self._fl.setStyleSheet(f"color:{P['sub']};font-size:12px;"); fr.addWidget(self._fl,1)
        bf=QPushButton(" 폴더 선택"); bf.clicked.connect(self._pf); fr.addWidget(bf); ci.addLayout(fr)
        opt=QHBoxLayout(); opt.addWidget(QLabel("최대 너비(px):"))
        self._w=QLineEdit("1280"); self._w.setFixedWidth(80); opt.addWidget(self._w)
        opt.addWidget(QLabel(" 포맷:")); self._fmt=QComboBox(); self._fmt.addItems(["원본 유지","JPEG","PNG","WEBP"]); opt.addWidget(self._fmt)
        opt.addWidget(QLabel(" 품질:")); self._q=QLineEdit("85"); self._q.setFixedWidth(50); opt.addWidget(self._q); opt.addStretch(); ci.addLayout(opt)
        self._btn=SpinBtn("️ 처리 시작"); self._btn.clicked.connect(self._run); ci.addWidget(self._btn)
        self._log=self.logbox(); self.stretch(); self._fol=""
    def _pf(self):
        d=QFileDialog.getExistingDirectory(self,"폴더 선택")
        if d: self._fol=d; self._fl.setText(d)
    def _run(self):
        if not self._guard([
            (HAS_PIL, "Pillow 패키지가 필요합니다"),
            (bool(self.files) or bool(self._fol),
                      "이미지 파일 또는 폴더를 선택해 주세요"),
        ]): return
        self._btn.start_spin()
        w=Worker(self._exec)
        w.log_sig.connect(lambda m: self.log(self._log,m))
        w.done_sig.connect(lambda m,k: (self.app.toast(m,k),self._btn.stop_spin()))
        w.start(); self._w=w
    def _exec(self,w):
        if not HAS_PIL: w.log_sig.emit(" Pillow 필요"); return
        EXTS={".jpg",".jpeg",".png",".bmp",".webp",".tiff"}
        mw=self._w.text().strip(); maxw=int(mw) if mw.isdigit() else None
        try: q=int(self._q.text().strip())
        except: q=85
        fc=self._fmt.currentText()
        if self.files: files=self.files; base=os.path.dirname(self.files[0])
        elif self._fol: files=[os.path.join(self._fol,f) for f in os.listdir(self._fol) if Path(f).suffix.lower() in EXTS]; base=self._fol
        else: w.log_sig.emit("️ 파일 또는 폴더 선택 먼저"); return
        out=os.path.join(base,"_처리완료"); os.makedirs(out,exist_ok=True)
        for i,src in enumerate(files):
            img=Image.open(src)
            if maxw and img.width>maxw: img=img.resize((maxw,int(img.height*maxw/img.width)),Image.LANCZOS)
            oe=("."+fc.lower()) if fc!="원본 유지" else Path(src).suffix.lower()
            of="JPEG" if oe in (".jpg",".jpeg") else (fc if fc!="원본 유지" else oe[1:].upper())
            op=os.path.join(out,Path(src).stem+oe)
            if of=="JPEG": img=img.convert("RGB")
            img.save(op,format=of,**{"quality":q} if of=="JPEG" else {})
            w.log_sig.emit(f" ({i+1}/{len(files)}) {os.path.basename(src)}")
        w.done_sig.emit(f"{len(files)}개 처리 완료!","ok")

# ── 이미지 → PDF
class ImgPdfPage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("","이미지 → PDF","선택한 이미지들을 순서대로 PDF 한 파일로 묶습니다")
        self.tip([
    "여러 이미지 파일을 순서대로 묶어 하나의 PDF로 만들어줍니다.",
    "① 이미지를 추가하는 순서가 PDF의 페이지 순서가 됩니다.",
    " 순서가 중요하다면 파일명에 번호를 붙여두고 추가하세요.",
    "② 스캔한 문서 사진을 PDF로 제출하거나, 사진들을 하나의 파일로 묶을 때 유용합니다.",
    "③ 출력 파일명과 저장 폴더를 지정하지 않으면 첫 번째 이미지와 같은 폴더에 저장됩니다.",
])
        br,self.files,self._add,self._clr=self.filelist()
        ba=_btn(" 추가","success",True)
        ba.clicked.connect(lambda: self._add(QFileDialog.getOpenFileNames(self,""," ","이미지 (*.jpg *.jpeg *.png *.bmp *.webp)")[0]))
        bc=_btn(" 초기화","danger",True); bc.clicked.connect(self._clr)
        br.addWidget(ba); br.addWidget(bc)
        ci=self.card()
        nr=QHBoxLayout(); nr.addWidget(QLabel("출력 파일명:"))
        self._n=QLineEdit("과제제출본"); nr.addWidget(self._n); nr.addWidget(QLabel(".pdf")); nr.addStretch(); ci.addLayout(nr)
        sr=QHBoxLayout(); self._sl=QLabel("저장 폴더 선택"); self._sl.setStyleSheet(f"color:{P['sub']};font-size:12px;"); sr.addWidget(self._sl,1)
        bs=QPushButton(" 선택"); bs.clicked.connect(self._ps); sr.addWidget(bs); ci.addLayout(sr)
        self._btn=SpinBtn(" PDF 생성"); self._btn.clicked.connect(self._run); ci.addWidget(self._btn)
        self._log=self.logbox(); self.stretch(); self._sv=""
    def _ps(self):
        d=QFileDialog.getExistingDirectory(self,"저장 폴더")
        if d: self._sv=d; self._sl.setText(d)
    def _run(self):
        if not self._guard([
            (HAS_PIL,          "Pillow 패키지가 필요합니다"),
            (bool(self.files), "PDF로 묶을 이미지를 추가해 주세요"),
        ]): return
        self._btn.start_spin()
        w=Worker(self._exec); w.log_sig.connect(lambda m: self.log(self._log,m))
        w.done_sig.connect(lambda m,k: (self.app.toast(m,k),self._btn.stop_spin())); w.start(); self._w=w
    def _exec(self,w):
        imgs=[Image.open(f).convert("RGB") for f in self.files]
        for f in self.files: w.log_sig.emit(f" {os.path.basename(f)}")
        name=self._n.text().strip() or "output"; folder=self._sv or os.path.dirname(self.files[0])
        op=os.path.join(folder,name+".pdf"); imgs[0].save(op,save_all=True,append_images=imgs[1:])
        w.log_sig.emit(f" {name}.pdf ({len(imgs)}장)"); w.done_sig.emit(f"{name}.pdf 생성 완료!","ok")

# ── 워터마크
class WatermarkPage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("","워터마크 삽입","이미지에 텍스트 워터마크를 일괄 삽입합니다")
        self.tip([
    "이미지에 텍스트 워터마크를 넣어 무단 도용을 방지합니다.",
    "① 워터마크 텍스트: 이름, 회사명, '대외비', '© 2025' 등 자유롭게 입력하세요.",
    "② 위치: 우하단(기본) / 우상단 / 중앙 / 좌하단 / 좌상단 중 선택",
    "③ 투명도: 0(완전 투명) ~ 255(완전 불투명). 120~160 정도가 자연스럽습니다.",
    "④ 결과는 원본 폴더 안 _워터마크 폴더에 저장됩니다. 원본은 유지됩니다.",
])
        br,self.files,self._add,self._clr=self.filelist()
        ba=_btn(" 추가","success",True)
        ba.clicked.connect(lambda: self._add(QFileDialog.getOpenFileNames(self,""," ","이미지 (*.jpg *.jpeg *.png *.webp *.bmp)")[0]))
        bc=_btn(" 초기화","danger",True); bc.clicked.connect(self._clr)
        br.addWidget(ba); br.addWidget(bc)
        ci=self.card(); g=QFormLayout(); g.setSpacing(8)
        self._wt=QLineEdit(); self._wt.setPlaceholderText("예) 대외비 / © 회사명"); g.addRow("워터마크 텍스트:",self._wt)
        self._wp=QComboBox(); self._wp.addItems(["우하단","우상단","중앙","좌하단","좌상단"]); g.addRow("위치:",self._wp)
        self._wa=QLineEdit("120"); self._wa.setFixedWidth(60); g.addRow("투명도(0~255):",self._wa); ci.addLayout(g)
        self._btn=SpinBtn(" 워터마크 삽입"); self._btn.clicked.connect(self._run); ci.addWidget(self._btn)
        self._log=self.logbox(); self.stretch()
    def _run(self):
        if not self._guard([
            (HAS_PIL,                          "Pillow 패키지가 필요합니다"),
            (bool(self.files),                 "워터마크를 삽입할 이미지를 추가해 주세요"),
            (bool(self._wt.text().strip()),    "워터마크 텍스트를 입력해 주세요"),
        ]): return
        self._btn.start_spin()
        w=Worker(self._exec); w.log_sig.connect(lambda m: self.log(self._log,m))
        w.done_sig.connect(lambda m,k: (self.app.toast(m,k),self._btn.stop_spin())); w.start(); self._w=w
    def _exec(self,w):
        text=self._wt.text().strip()
        if not text: w.log_sig.emit("️ 워터마크 텍스트 입력 필요"); return
        pos=self._wp.currentText()
        try: alpha=int(self._wa.text())
        except: alpha=120
        out=os.path.join(os.path.dirname(self.files[0]),"_워터마크"); os.makedirs(out,exist_ok=True)
        for fp in self.files:
            try:
                img=Image.open(fp).convert("RGBA"); ov=Image.new("RGBA",img.size,(0,0,0,0)); draw=ImageDraw.Draw(ov)
                try: font=ImageFont.truetype("malgun.ttf",max(20,img.width//20))
                except: font=ImageFont.load_default()
                bb=draw.textbbox((0,0),text,font=font); tw,th=bb[2]-bb[0],bb[3]-bb[1]; mg=20
                pm={"우하단":(img.width-tw-mg,img.height-th-mg),"우상단":(img.width-tw-mg,mg),
                     "중앙":((img.width-tw)//2,(img.height-th)//2),"좌하단":(mg,img.height-th-mg),"좌상단":(mg,mg)}
                draw.text(pm.get(pos,(mg,mg)),text,font=font,fill=(255,255,255,alpha))
                Image.alpha_composite(img,ov).convert("RGB").save(os.path.join(out,os.path.basename(fp)))
                w.log_sig.emit(f" {os.path.basename(fp)}")
            except Exception as e: w.log_sig.emit(f" {os.path.basename(fp)}: {e}")
        w.done_sig.emit("워터마크 삽입 완료!","ok")

# ── 배경 제거
class RembgPage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("️","배경 제거","AI가 이미지 배경을 자동으로 제거합니다 (로컬 처리)")
        self.tip([
    "AI(rembg)가 이미지의 배경을 자동으로 인식하고 제거합니다. 인터넷 불필요.",
    "① 인물, 제품, 동물 사진 등 다양한 이미지에서 배경을 지울 수 있습니다.",
    "② 결과는 투명 배경의 PNG 파일로 저장됩니다 (_nobg.png).",
    "③ 처음 실행 시 AI 모델을 다운로드해 속도가 느릴 수 있습니다. 이후엔 빠릅니다.",
    "④ 복잡한 배경이나 머리카락이 많은 경우 결과가 완벽하지 않을 수 있습니다.",
])
        br,self.files,self._add,self._clr=self.filelist()
        ba=_btn(" 파일 추가","success",True)
        ba.clicked.connect(lambda: self._add(QFileDialog.getOpenFileNames(self,""," ","이미지 (*.jpg *.jpeg *.png *.webp *.bmp)")[0]))
        bc=_btn(" 초기화","danger",True); bc.clicked.connect(self._clr)
        br.addWidget(ba); br.addWidget(bc)
        ci=self.card()
        rp_row=QHBoxLayout(); self._prog=QProgressBar(); self._prog.setValue(0)
        self._prog_lbl=QLabel("0%")
        self._prog_lbl.setStyleSheet(f"color:{P['sub']};font-size:11px;min-width:40px;background:transparent;")
        self._prog_lbl.setAlignment(Qt.AlignmentFlag.AlignRight)
        rp_row.addWidget(self._prog,1); rp_row.addWidget(self._prog_lbl); ci.addLayout(rp_row)
        self._btn=SpinBtn("️ 배경 제거 시작"); self._btn.clicked.connect(self._run); ci.addWidget(self._btn)
        self._log=self.logbox(); self.stretch()
    def _run(self):
        if not self._guard([
            (HAS_REMBG,        "rembg 패키지가 필요합니다 (pip install rembg)"),
            (bool(self.files), "배경을 제거할 이미지를 추가해 주세요"),
        ]): return
        self._btn.start_spin()
        w=Worker(self._exec); w.log_sig.connect(lambda m: self.log(self._log,m))
        w.prog_sig.connect(self._prog.setValue)
        w.prog_sig.connect(lambda v: self._prog_lbl.setText(f"{v}%"))
        w.prog_sig.connect(lambda v: self._prog_lbl.setText(f"{v}%"))
        w.done_sig.connect(lambda m,k: (self.app.toast(m,k),self._btn.stop_spin(),self._prog_lbl.setText("✓ 완료"),self._prog_lbl.setText("완료")))
        w.start(); self._w=w
    def _exec(self,w):
        if not HAS_REMBG: w.log_sig.emit(" rembg 필요"); return
        if not self.files: w.log_sig.emit("️ 이미지 추가 먼저"); return
        out=os.path.join(os.path.dirname(self.files[0]),"_배경제거"); os.makedirs(out,exist_ok=True)
        for i,fp in enumerate(self.files):
            with open(fp,"rb") as f: inp=f.read()
            op=os.path.join(out,Path(fp).stem+"_nobg.png")
            with open(op,"wb") as f: f.write(rembg_remove(inp))
            w.log_sig.emit(f" {os.path.basename(fp)} → {os.path.basename(op)}")
            w.prog_sig.emit(int((i+1)/len(self.files)*100))
        w.done_sig.emit(f"배경 제거 완료! ({len(self.files)}개)","ok")

# ── OCR
class OcrPage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("","이미지 OCR","사진·스캔 이미지 속 텍스트를 Gemini AI가 추출합니다")
        self.tip([
    "사진이나 스캔한 이미지 안의 텍스트를 Gemini AI가 읽어 텍스트 파일로 저장합니다.",
    "① 책 페이지 사진, 명함, 영수증, 칠판 사진 등 다양하게 활용할 수 있습니다.",
    "② 원본 그대로: 이미지에 보이는 언어 그대로 추출합니다.",
    " 한국어로 번역: 추출과 동시에 한국어로 번역합니다.",
    "③ 추출 결과는 화면에 표시되고, 원본 파일명에 _OCR.txt가 붙어 자동 저장됩니다.",
    "④ Gemini API 키가 필요합니다 — ️ 설정에서 등록하세요.",
])
        br,self.files,self._add,self._clr=self.filelist()
        ba=_btn(" 추가","success",True)
        ba.clicked.connect(lambda: self._add(QFileDialog.getOpenFileNames(self,""," ","이미지 (*.jpg *.jpeg *.png *.webp *.bmp *.tiff)")[0]))
        bc=_btn(" 초기화","danger",True); bc.clicked.connect(self._clr)
        br.addWidget(ba); br.addWidget(bc)
        ci=self.card()
        lr=QHBoxLayout(); lr.addWidget(QLabel("출력 언어:"))
        self._lo=QRadioButton("원본 그대로"); self._lo.setChecked(True); self._lk=QRadioButton("한국어로 번역")
        lr.addWidget(self._lo); lr.addWidget(self._lk); lr.addStretch(); ci.addLayout(lr)
        self._btn=SpinBtn(" 텍스트 추출"); self._btn.clicked.connect(self._run); ci.addWidget(self._btn)
        self._res=self.result(200); self.stretch()
    def _run(self):
        if not self._guard([
            (bool(ai_client),  "Gemini API 키가 없습니다 — ⚙️ 설정에서 등록해 주세요"),
            (bool(self.files), "OCR 처리할 이미지를 추가해 주세요"),
        ]): return
        self._btn.start_spin()
        w=Worker(self._exec); w.log_sig.connect(self._res.append)
        w.done_sig.connect(lambda m,k: (self.app.toast(m,k),self._btn.stop_spin())); w.start(); self._w=w
    def _exec(self,w):
        mm={"jpg":"image/jpeg","jpeg":"image/jpeg","png":"image/png","webp":"image/webp","bmp":"image/bmp","tiff":"image/tiff"}
        for fp in self.files:
            w.log_sig.emit(f"━━━ {os.path.basename(fp)} ━━━")
            with open(fp,"rb") as f: data=f.read()
            mime=mm.get(Path(fp).suffix.lower().lstrip("."),"image/jpeg")
            prompt="이미지에서 모든 텍스트를 추출하고 한국어로 번역해줘. 텍스트만 반환해." if self._lk.isChecked() else "이미지에서 모든 텍스트를 그대로 추출해줘. 텍스트만 반환해."
            try:
                res=ai_client.models.generate_content(model=TARGET_MODEL,contents=[types.Part.from_bytes(data=data,mime_type=mime),prompt])
                txt=str(res.text).strip(); w.log_sig.emit(txt)
                with open(os.path.splitext(fp)[0]+"_OCR.txt","w",encoding="utf-8") as f: f.write(txt)
                w.log_sig.emit(" 저장됨\n")
            except Exception as e: w.log_sig.emit(f" {e}\n")
        w.done_sig.emit("OCR 완료!","ok")

# ── AI 요약
class SummaryPage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("","AI 문서 요약","PDF·TXT·DOCX를 Gemini AI가 핵심만 요약합니다")
        self.tip([
    "긴 문서를 Gemini AI가 읽고 핵심 내용만 골라 요약해줍니다.",
    "① 핵심 요점: 중요한 내용을 5~10개의 불릿 포인트로 정리합니다.",
    " 한 단락: 전체 내용을 3~5문장으로 압축합니다.",
    " 시험 Q&A: 시험에 나올 법한 질문과 답변 5쌍을 만들어 줍니다.",
    "② PDF·TXT·DOCX 파일을 지원합니다. 스캔 PDF(이미지 PDF)는 텍스트 추출이 안 됩니다.",
    "③ 결과는 화면에 표시되고 _AI요약.txt로 자동 저장됩니다.",
    "④ Gemini API 키가 필요합니다 — ️ 설정에서 등록하세요.",
])
        br,self.files,self._add,self._clr=self.filelist()
        ba=_btn(" 추가","success",True)
        ba.clicked.connect(lambda: self._add(QFileDialog.getOpenFileNames(self,""," ","문서 (*.pdf *.txt *.docx)")[0]))
        bc=_btn(" 초기화","danger",True); bc.clicked.connect(self._clr)
        br.addWidget(ba); br.addWidget(bc)
        ci=self.card()
        sr=QHBoxLayout(); sr.addWidget(QLabel("요약 스타일:"))
        self._sb=QRadioButton("핵심 요점"); self._sb.setChecked(True); self._sp=QRadioButton("한 단락"); self._sq=QRadioButton("시험 Q&A")
        for r in (self._sb,self._sp,self._sq): sr.addWidget(r)
        sr.addStretch(); ci.addLayout(sr)
        self._btn=SpinBtn(" 요약 시작"); self._btn.clicked.connect(self._run); ci.addWidget(self._btn)
        self._res=self.result(220)
        self._char_lbl=QLabel("")
        self._char_lbl.setStyleSheet(f"color:{P['sub2']};font-size:11px;background:transparent;margin:0 28px 4px;")
        self._char_lbl.setAlignment(Qt.AlignmentFlag.AlignRight)
        self._v.addWidget(self._char_lbl)
        self._res.textChanged.connect(lambda: self._char_lbl.setText(f"{len(self._res.toPlainText()):,}자"))
        self.stretch()
    def _run(self):
        if not self._guard([
            (bool(ai_client),  "Gemini API 키가 없습니다 — ⚙️ 설정에서 등록해 주세요"),
            (bool(self.files), "요약할 파일(PDF·TXT·DOCX)을 추가해 주세요"),
        ]): return
        self._btn.start_spin()
        w=Worker(self._exec); w.log_sig.connect(self._res.append)
        w.done_sig.connect(lambda m,k: (self.app.toast(m,k),self._btn.stop_spin())); w.start(); self._w=w
    def _exec(self,w):
        style="bullet" if self._sb.isChecked() else "paragraph" if self._sp.isChecked() else "qa"
        prompts={"bullet":"핵심 요점 5~10개로 불릿 요약해줘 (한국어):\n\n","paragraph":"3~5문장 단락으로 요약해줘 (한국어):\n\n","qa":"시험 Q&A 5개 만들어줘 (한국어, Q:/A: 형식):\n\n"}
        for fp in self.files:
            ext=Path(fp).suffix.lower(); content=""
            try:
                if ext==".txt":
                    with open(fp,"r",encoding="utf-8") as f: content=f.read()
                elif ext==".pdf" and HAS_FITZ: content="".join(pg.get_text() for pg in fitz.open(fp))
                elif ext==".docx" and HAS_DOCX: content="\n".join(p.text for p in Document(fp).paragraphs)
                else: w.log_sig.emit(f"️ {os.path.basename(fp)}: 지원 불가"); continue
                res=ai_client.models.generate_content(model=TARGET_MODEL,contents=prompts[style]+content[:8000])
                txt=str(res.text).strip(); w.log_sig.emit(f"━━━ {os.path.basename(fp)} ━━━\n{txt}\n")
                with open(os.path.splitext(fp)[0]+"_AI요약.txt","w",encoding="utf-8") as f: f.write(txt)
                w.log_sig.emit(" 저장됨\n")
            except Exception as e: w.log_sig.emit(f" {os.path.basename(fp)}: {e}\n")
        w.done_sig.emit("요약 완료!","ok")

# ── AI 초안
class DraftPage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("️","AI 문서 초안 작성","주제와 조건을 입력하면 Gemini AI가 초안을 자동 생성합니다")
        self.tip([
    "주제와 조건을 입력하면 Gemini AI가 완성도 있는 문서 초안을 작성해줍니다.",
    "① 문서 종류: 보고서·자기소개서·이메일·기획서·레포트·회의록 등 선택",
    "② 추가 조건 예시: 'A4 2장 분량', '격식체', '소방법 관련 규정 포함'",
    " 조건이 구체적일수록 원하는 결과에 가까워집니다.",
    "③ 생성된 초안은 직접 편집할 수 있으며, ' TXT 저장'으로 파일로 내보낼 수 있습니다.",
    "④ Gemini API 키가 필요합니다 — ️ 설정에서 등록하세요.",
])
        ci=self.card(); g=QFormLayout(); g.setSpacing(8)
        self._ty=QComboBox(); self._ty.addItems(["보고서","자기소개서","이메일","독후감","기획서","레포트","발표 스크립트","회의록","공지문","기타"]); g.addRow("문서 종류:",self._ty)
        self._tp=QLineEdit(); self._tp.setPlaceholderText("예) 소방시설 설치 기준"); g.addRow("주제/키워드:",self._tp)
        self._cd=QLineEdit(); self._cd.setPlaceholderText("예) A4 1장, 공식 문체, 3단락"); g.addRow("추가 조건:",self._cd)
        lr=QHBoxLayout(); self._ko=QRadioButton("한국어"); self._ko.setChecked(True); self._en=QRadioButton("English"); lr.addWidget(self._ko); lr.addWidget(self._en); lr.addStretch(); g.addRow("언어:",lr); ci.addLayout(g)
        br=QHBoxLayout()
        self._btn=SpinBtn("️ 초안 생성"); self._btn.clicked.connect(self._run); br.addWidget(self._btn)
        bs=_btn(" TXT 저장","success"); bs.clicked.connect(self._save); br.addWidget(bs)
        bc=_btn(" 초기화","danger"); bc.clicked.connect(self._clr); br.addWidget(bc); br.addStretch(); ci.addLayout(br)
        self._res=self.result(240)
        self._char_lbl=QLabel("")
        self._char_lbl.setStyleSheet(f"color:{P['sub2']};font-size:11px;text-align:right;background:transparent;margin:0 28px 4px;")
        self._char_lbl.setAlignment(Qt.AlignmentFlag.AlignRight)
        self._v.addWidget(self._char_lbl)
        self._res.textChanged.connect(lambda: self._char_lbl.setText(f"{len(self._res.toPlainText()):,}자"))
        self.stretch()
    def _run(self):
        tp=self._tp.text().strip()
        if not self._guard([
            (bool(ai_client), "Gemini API 키가 없습니다 — ⚙️ 설정에서 등록해 주세요"),
            (bool(tp),        "주제/키워드를 입력해 주세요"),
        ]): return
        self._btn.start_spin(); self._res.clear(); self._res.append("✍️ 작성 중...")
        cd=self._cd.text().strip(); lang="한국어" if self._ko.isChecked() else "English"
        cond_str = ("조건: " + cd + "\n") if cd else ""
        prompt=f"{self._ty.currentText()} 초안을 작성해줘.\n주제: {tp}\n{cond_str}언어: {lang}\n제목 포함, 완성도 있게."
        def fn(w):
            if not ai_client: w.log_sig.emit(" Gemini API 키 없음"); w.done_sig.emit("","err"); return
            try:
                res=ai_client.models.generate_content(model=TARGET_MODEL,contents=prompt)
                w.log_sig.emit(str(res.text).strip()); w.done_sig.emit("초안 생성 완료!","ok")
            except Exception as e: w.log_sig.emit(f" {e}"); w.done_sig.emit("","err")
        worker=Worker(fn)
        def _sr(m): self._res.clear(); self._res.append(m)
        worker.log_sig.connect(_sr)
        worker.done_sig.connect(lambda m,k: (self.app.toast(m,k) if m else None,self._btn.stop_spin()))
        worker.start(); self._w=worker
    def _save(self):
        c=self._res.toPlainText().strip()
        if not c: return
        path,_=QFileDialog.getSaveFileName(self,"저장","","텍스트 (*.txt)")
        if path:
            with open(path,"w",encoding="utf-8") as f: f.write(c)
            self.app.toast(f"저장: {os.path.basename(path)}","ok")
    def _clr(self): self._res.clear(); self._tp.clear(); self._cd.clear(); self._char_lbl.setText('')

# ── 참고문헌
class CitationPage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("","참고문헌 자동 정리","논문·책·URL 정보를 AI가 인용 형식으로 자동 정리합니다")
        self.tip([
    "논문·책·웹사이트의 출처 정보를 정해진 인용 형식에 맞게 자동으로 변환합니다.",
    "① 정보를 줄바꿈으로 구분해 자유롭게 입력하세요. 형식이 불완전해도 괜찮습니다.",
    " 예) 김민수 외 (2023). 소방시설 설치기준. 한국소방학회지, 45(2), 123-145.",
    "② 지원 형식: APA 7판 / MLA 9판 / 시카고 스타일 / 한국어 논문(KCI)",
    "③ 논문 제출 전 참고문헌 목록을 빠르게 정리할 때 특히 유용합니다.",
    "④ Gemini API 키가 필요합니다 — ️ 설정에서 등록하세요.",
])
        ci=self.card()
        sr=QHBoxLayout(); sr.addWidget(QLabel("인용 형식:"))
        self._st=QComboBox(); self._st.addItems(["APA 7판","MLA 9판","시카고 스타일","한국어 논문(KCI)"]); sr.addWidget(self._st); sr.addStretch(); ci.addLayout(sr)
        ci.addWidget(QLabel("참고문헌 정보 (줄바꿈으로 구분):"))
        self._inp=QTextEdit(); self._inp.setMaximumHeight(130)
        self._inp.setPlaceholderText("예시:\n김민수 (2023). 인공지능과 교육. 한국교육학회지, 45(2), 123-145.\nSmith, J. (2022). Deep Learning. MIT Press.\nhttps://www.example.com (접속일: 2024.01.15)")
        ci.addWidget(self._inp)
        self._btn=SpinBtn(" 참고문헌 정리"); self._btn.clicked.connect(self._run); ci.addWidget(self._btn)
        self._res=self.result(200)
        self._char_lbl=QLabel("")
        self._char_lbl.setStyleSheet(f"color:{P['sub2']};font-size:11px;background:transparent;margin:0 28px 4px;")
        self._char_lbl.setAlignment(Qt.AlignmentFlag.AlignRight)
        self._v.addWidget(self._char_lbl)
        self._res.textChanged.connect(lambda: self._char_lbl.setText(f"{len(self._res.toPlainText()):,}자"))
        self.stretch()
    def _run(self):
        raw=self._inp.toPlainText().strip()
        if not self._guard([
            (bool(ai_client), "Gemini API 키가 없습니다 — ⚙️ 설정에서 등록해 주세요"),
            (bool(raw),       "정리할 참고문헌 정보를 입력해 주세요"),
        ]): return
        self._btn.start_spin(); self._res.clear(); self._res.append("🔄 AI가 정리 중...")
        style=self._st.currentText()
        def fn(w):
            if not ai_client: w.log_sig.emit(" Gemini API 키 없음"); w.done_sig.emit("","err"); return
            try:
                prompt=f"아래 참고문헌 정보들을 {style} 형식으로 정확하게 정리해줘.\n각 항목을 번호 없이 한 줄씩 변환하고 결과만 반환해.\n\n{raw}"
                res=ai_client.models.generate_content(model=TARGET_MODEL,contents=prompt)
                w.log_sig.emit(str(res.text).strip()); w.done_sig.emit("참고문헌 정리 완료!","ok")
            except Exception as e: w.log_sig.emit(f" {e}"); w.done_sig.emit("","err")
        worker=Worker(fn)
        def _sr(m): self._res.clear(); self._res.append(m)
        worker.log_sig.connect(_sr)
        worker.done_sig.connect(lambda m,k: (self.app.toast(m,k) if m else None,self._btn.stop_spin()))
        worker.start(); self._w=worker

# ── 엑셀
class ExcelPage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("","엑셀 자동화","엑셀·CSV 파일의 시트 합치기, 중복 제거, 포맷 변환을 자동화합니다")
        self.tip([
    "엑셀과 CSV 파일의 반복 작업을 자동화합니다.",
    "① 시트 합치기: 여러 엑셀 파일의 시트를 하나의 파일(_합본.xlsx)로 모읍니다.",
    " 여러 달치 데이터나 부서별 파일을 합칠 때 유용합니다.",
    "② 중복 행 제거: 완전히 동일한 행을 자동으로 찾아 제거합니다.",
    " 결과는 _중복제거.xlsx(.csv)로 저장됩니다.",
    "③ CSV → XLSX: CSV 파일을 엑셀 형식(.xlsx)으로 변환합니다.",
])
        br,self.files,self._add,self._clr=self.filelist()
        ba=_btn(" 파일 추가","success",True)
        ba.clicked.connect(lambda: self._add(QFileDialog.getOpenFileNames(self,""," ","엑셀/CSV (*.xlsx *.xls *.csv)")[0]))
        bc=_btn(" 초기화","danger",True); bc.clicked.connect(self._clr)
        br.addWidget(ba); br.addWidget(bc)
        ci=self.card()
        mr=QHBoxLayout(); mr.addWidget(QLabel("작업 선택:"))
        self._mg=QRadioButton("시트 합치기"); self._mg.setChecked(True); self._dd=QRadioButton("중복 행 제거"); self._c2=QRadioButton("CSV → XLSX")
        for r in (self._mg,self._dd,self._c2): mr.addWidget(r)
        mr.addStretch(); ci.addLayout(mr)
        self._btn=SpinBtn(" 실행"); self._btn.clicked.connect(self._run); ci.addWidget(self._btn)
        self._log=self.logbox(); self.stretch()
    def _run(self):
        if not self._guard([
            (HAS_XL,           "openpyxl 패키지가 필요합니다"),
            (bool(self.files), "처리할 엑셀/CSV 파일을 추가해 주세요"),
        ]): return
        self._btn.start_spin()
        w=Worker(self._exec); w.log_sig.connect(lambda m: self.log(self._log,m))
        w.done_sig.connect(lambda m,k: (self.app.toast(m,k),self._btn.stop_spin()))
        w.start(); self._w=w
    def _exec(self,w):
        import csv as _csv
        if not HAS_XL: w.log_sig.emit(" openpyxl 필요"); return
        if not self.files: w.log_sig.emit("️ 파일 추가 먼저"); return
        mode="merge" if self._mg.isChecked() else "dedup" if self._dd.isChecked() else "csv2xl"
        if mode=="merge":
            mg=openpyxl.Workbook(); mg.remove(mg.active)
            for fp in self.files:
                wb=openpyxl.load_workbook(fp)
                for sn in wb.sheetnames:
                    ws=mg.create_sheet(title=(Path(fp).stem[:10]+"_"+sn)[:31])
                    for row in wb[sn].iter_rows(values_only=True): ws.append(list(row))
                w.log_sig.emit(f" {os.path.basename(fp)} ({len(wb.sheetnames)}시트)")
            op=os.path.join(os.path.dirname(self.files[0]),"_합본.xlsx"); mg.save(op)
            w.log_sig.emit(f" 저장: {os.path.basename(op)}"); w.done_sig.emit(f"시트 합치기 완료! ({len(mg.sheetnames)}시트)","ok")
        elif mode=="dedup":
            for fp in self.files:
                ext=Path(fp).suffix.lower()
                if ext==".csv":
                    with open(fp,"r",encoding="utf-8-sig") as f: rows=list(_csv.reader(f))
                    hd=rows[0] if rows else []; seen=set(); uniq=[hd]
                    for row in rows[1:]:
                        k=tuple(row)
                        if k not in seen: seen.add(k); uniq.append(row)
                    op=os.path.splitext(fp)[0]+"_중복제거.csv"
                    with open(op,"w",encoding="utf-8-sig",newline="") as f: _csv.writer(f).writerows(uniq)
                else:
                    wb=openpyxl.load_workbook(fp); ws=wb.active
                    rows=list(ws.iter_rows(values_only=True)); hd=rows[0] if rows else ()
                    seen=set(); uniq=[hd]
                    for row in rows[1:]:
                        if row not in seen: seen.add(row); uniq.append(row)
                    wb2=openpyxl.Workbook(); ws2=wb2.active
                    for row in uniq: ws2.append(list(row))
                    op=os.path.splitext(fp)[0]+"_중복제거.xlsx"; wb2.save(op)
                w.log_sig.emit(f" {os.path.basename(fp)} → 중복 {len(rows)-len(uniq)}행 제거")
            w.done_sig.emit("중복 행 제거 완료!","ok")
        else:
            for fp in self.files:
                if Path(fp).suffix.lower()!=".csv": w.log_sig.emit(f"️ {os.path.basename(fp)}: CSV 아님"); continue
                wb=openpyxl.Workbook(); ws=wb.active
                with open(fp,"r",encoding="utf-8-sig") as f:
                    for row in _csv.reader(f): ws.append(row)
                op=os.path.splitext(fp)[0]+".xlsx"; wb.save(op); w.log_sig.emit(f" {os.path.basename(fp)} → {Path(op).name}")
            w.done_sig.emit("CSV → XLSX 변환 완료!","ok")

# ── PDF 비밀번호
class PdfPwdPage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("","PDF 비밀번호 설정 / 해제","PDF에 비밀번호를 걸거나 일괄 해제합니다")
        self.tip([
    "PDF에 비밀번호를 걸어 열람을 제한하거나, 기존 비밀번호를 제거합니다.",
    " 비밀번호 설정: AES-256 암호화로 잠근 _잠금.pdf 파일이 생성됩니다.",
    " 원본 파일은 그대로 유지되며 잠긴 파일이 새로 만들어집니다.",
    " 비밀번호 해제: 기존 비밀번호를 입력하면 잠금이 풀린 _해제.pdf가 생성됩니다.",
    "① 여러 파일을 한 번에 추가하면 동일한 비밀번호로 일괄 처리됩니다.",
])
        br,self.files,self._add,self._clr=self.filelist()
        ba=_btn(" 추가","success",True)
        ba.clicked.connect(lambda: self._add(QFileDialog.getOpenFileNames(self,""," ","PDF (*.pdf)")[0]))
        bc=_btn(" 초기화","danger",True); bc.clicked.connect(self._clr)
        br.addWidget(ba); br.addWidget(bc)
        ci=self.card()
        pr=QHBoxLayout(); pr.addWidget(QLabel("비밀번호:"))
        self._pw=QLineEdit(); self._pw.setEchoMode(QLineEdit.EchoMode.Password); self._pw.setPlaceholderText("비밀번호 입력"); pr.addWidget(self._pw,1); ci.addLayout(pr)
        bb=QHBoxLayout()
        b1=_btn(" 비밀번호 설정","danger"); b1.clicked.connect(lambda: self._exec("set")); bb.addWidget(b1)
        b2=_btn(" 비밀번호 해제","success"); b2.clicked.connect(lambda: self._exec("rm")); bb.addWidget(b2); bb.addStretch(); ci.addLayout(bb)
        self._log=self.logbox(); self.stretch()
    def _exec(self,mode):
        if not self._guard([
            (HAS_FITZ,                     "PyMuPDF(pymupdf) 패키지가 필요합니다"),
            (bool(self.files),             "처리할 PDF 파일을 추가해 주세요"),
            (bool(self._pw.text().strip()), "비밀번호를 입력해 주세요"),
        ]): return
        def fn(w):
            for fp in self.files:
                try:
                    doc=fitz.open(fp)
                    if mode=="set":
                        op=os.path.splitext(fp)[0]+"_잠금.pdf"; doc.save(op,encryption=fitz.PDF_ENCRYPT_AES_256,owner_pw=pw,user_pw=pw); w.log_sig.emit(f" 설정: {os.path.basename(op)}")
                    else:
                        if doc.is_encrypted: doc.authenticate(pw)
                        op=os.path.splitext(fp)[0]+"_해제.pdf"; doc.save(op,encryption=fitz.PDF_ENCRYPT_NONE); w.log_sig.emit(f" 해제: {os.path.basename(op)}")
                except Exception as e: w.log_sig.emit(f" {os.path.basename(fp)}: {e}")
            w.done_sig.emit("비밀번호 처리 완료!","ok")
        worker=Worker(fn); worker.log_sig.connect(lambda m: self.log(self._log,m))
        worker.done_sig.connect(lambda m,k: self.app.toast(m,k)); worker.start(); self._w=worker

# ── 과제 트래커
class TrackerPage(QWidget):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app; self.tasks=load_tasks(); self._sel=None
        self.setStyleSheet(f"background:{P['bg']};")
        vl=QVBoxLayout(self); vl.setContentsMargins(0,0,0,0); vl.setSpacing(0)
        hw=QWidget(); hw.setStyleSheet(f"background:{P['bg']};"); hv=QVBoxLayout(hw); hv.setContentsMargins(28,18,28,8); hv.setSpacing(3)
        t=QLabel(" 과제 마감 트래커"); t.setStyleSheet(f"font-size:17px;font-weight:bold;color:{P['text']};background:transparent;")
        s=QLabel("과목별 과제 등록 및 D-day 관리"); s.setStyleSheet(f"color:{P['sub']};font-size:12px;background:transparent;")
        hv.addWidget(t); hv.addWidget(s); vl.addWidget(hw)
        tw=QWidget(); tw.setStyleSheet(f"background:{P['bg']};"); tv=QVBoxLayout(tw); tv.setContentsMargins(28,6,28,14); tv.addWidget(_tip(["① 과목명·과제명·마감일(YYYY-MM-DD) 입력 후 추가.","② 목록에서 행 클릭 → 완료 또는 삭제.","③ D-3 이하 주황, 당일·초과 빨강 표시."])); vl.addWidget(tw)
        fw=QWidget(); fw.setStyleSheet(f"background:{P['bg']};"); fg=QGridLayout(fw); fg.setContentsMargins(28,0,28,8); fg.setSpacing(8)
        fg.addWidget(QLabel("과목명:"),0,0); self._sj=QLineEdit(); self._sj.setPlaceholderText("예) 소방학개론"); fg.addWidget(self._sj,0,1)
        fg.addWidget(QLabel("과제명:"),0,2); self._ti=QLineEdit(); self._ti.setPlaceholderText("예) 3장 요약"); fg.addWidget(self._ti,0,3)
        fg.addWidget(QLabel("마감일:"),1,0); self._du=QLineEdit(datetime.date.today().strftime("%Y-%m-%d")); fg.addWidget(self._du,1,1)
        # 진입 시 오늘 날짜 자동 갱신
        self._auto_date = True
        fg.addWidget(QLabel("메모:"),1,2); self._mo=QLineEdit(); self._mo.setPlaceholderText("선택 입력"); fg.addWidget(self._mo,1,3)
        fg.setColumnStretch(1,1); fg.setColumnStretch(3,1); vl.addWidget(fw)
        bw=QWidget(); bw.setStyleSheet(f"background:{P['bg']};"); bh=QHBoxLayout(bw); bh.setContentsMargins(28,0,28,10)
        b1=_btn(" 추가","success"); b1.clicked.connect(self._add); bh.addWidget(b1)
        b2=_btn(" 완료","warning"); b2.clicked.connect(self._done); bh.addWidget(b2)
        b2u=_btn(" 취소","ghost"); b2u.clicked.connect(self._undone); bh.addWidget(b2u)
        b3=_btn("️ 삭제","danger"); b3.clicked.connect(self._del); bh.addWidget(b3)
        bh.addStretch()
        b4=_btn(" 새로고침","ghost")
        b4.clicked.connect(self._refresh); bh.addWidget(b4); vl.addWidget(bw)
        rw=QWidget(); rw.setStyleSheet(f"background:{P['bg']};"); rv=QVBoxLayout(rw); rv.setContentsMargins(28,0,28,20)
        self._tree=QTreeWidget(); self._tree.setHeaderLabels(["과목","과제명","마감일","D-day","상태"])
        self._tree.setColumnWidth(0,110); self._tree.setColumnWidth(1,200); self._tree.setColumnWidth(2,100); self._tree.setColumnWidth(3,90); self._tree.setColumnWidth(4,80)
        self._tree.itemClicked.connect(lambda it: setattr(self,"_sel",it.data(0,Qt.ItemDataRole.UserRole)))
        self._tree.setRootIsDecorated(False); rv.addWidget(self._tree); vl.addWidget(rw,1)
        self._refresh()
    def _add(self):
        sj=self._sj.text().strip(); ti=self._ti.text().strip(); du=self._du.text().strip(); mo=self._mo.text().strip()
        if not sj or not ti or not du: self.app.toast("과목명, 과제명, 마감일은 필수입니다","err"); return
        try: datetime.datetime.strptime(du,"%Y-%m-%d")
        except: self.app.toast("마감일 형식: YYYY-MM-DD  (예: 2025-12-31)","err"); return
        self.tasks.append({"subject":sj,"title":ti,"due":du,"memo":mo,"done":False}); save_tasks(self.tasks)
        self._sj.clear(); self._ti.clear(); self._mo.clear()
        self._du.setText(datetime.date.today().strftime("%Y-%m-%d"))
        self._refresh()
    def _done(self):
        if self._sel is None: self.app.toast("목록에서 과제를 먼저 클릭해 선택하세요","warn"); return
        self.tasks[self._sel]["done"]=True; save_tasks(self.tasks); self._refresh()
    def _undone(self):
        if self._sel is None: QMessageBox.information(self,"안내","과제를 먼저 클릭해 선택하세요."); return
        self.tasks[self._sel]["done"]=False; save_tasks(self.tasks); self._refresh()
    def _del(self):
        if self._sel is None: QMessageBox.information(self,"안내","과제를 먼저 클릭해 선택하세요."); return
        self.tasks.pop(self._sel); save_tasks(self.tasks); self._sel=None; self._refresh()
    def showEvent(self, e):
        super().showEvent(e)
        if getattr(self, "_auto_date", False):
            self._du.setText(datetime.date.today().strftime("%Y-%m-%d"))

    def _refresh(self):
        self._sel=None; self._tree.clear(); today=datetime.date.today()
        for task in sorted(self.tasks,key=lambda t:(t.get("done",False),t["due"])):
            ridx=self.tasks.index(task); delta=(datetime.datetime.strptime(task["due"],"%Y-%m-%d").date()-today).days; done=task.get("done",False)
            if done: dt,col="완료 ",P["success"]
            elif delta<0: dt,col=f"D+{abs(delta)} 초과",P["danger"]
            elif delta==0: dt,col="D-day ",P["danger"]
            elif delta<=3: dt,col=f"D-{delta} ️",P["warning"]
            else: dt,col=f"D-{delta}",P["text"]
            it=QTreeWidgetItem([task["subject"],task["title"],task["due"],dt,"완료" if done else "진행중"])
            it.setData(0,Qt.ItemDataRole.UserRole,ridx); it.setForeground(3,QColor(col))
            if done: it.setForeground(4,QColor(P["success"]))
            self._tree.addTopLevelItem(it)

# ── 설정

# ── 도움말
class HelpPage(QScrollArea):
    def __init__(self, app=None, p=None):
        super().__init__(p)
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        root = QWidget()
        root.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        root.setStyleSheet(f"QWidget{{background:{P['bg']};}}")
        v = QVBoxLayout(root); v.setContentsMargins(0,0,0,0); v.setSpacing(0)
        self.setWidget(root)

        # 헤더
        hdr_w = QWidget()
        hdr_w.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        hdr_w.setStyleSheet(f"QWidget{{background:{P['bg']};}}")
        hv = QVBoxLayout(hdr_w); hv.setContentsMargins(28,24,28,8); hv.setSpacing(4)
        ht = QLabel('도움말 — API 키 발급 가이드')
        ht.setStyleSheet(f"font-size:20px;font-weight:700;color:{P['text']};background:transparent;letter-spacing:-0.5px;")
        hs = QLabel('DeepL · Gemini API 키를 무료로 발급받는 방법을 안내합니다')
        hs.setStyleSheet(f"color:{P['sub']};font-size:13px;background:transparent;")
        hv.addWidget(ht); hv.addWidget(hs); v.addWidget(hdr_w)

        def _sec(pv, title, color, items):
            sw = QWidget()
            sw.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
            sw.setStyleSheet(f"QWidget{{background:{P['bg']};}}")
            sv = QVBoxLayout(sw); sv.setContentsMargins(28,0,28,16)
            card = QFrame()
            card.setStyleSheet(f"QFrame{{background:{P['card']};border:none;border-radius:16px;}}")
            cv = QVBoxLayout(card); cv.setContentsMargins(24,20,24,22); cv.setSpacing(0)
            # 섹션 타이틀
            tr = QHBoxLayout(); tr.setSpacing(10)
            bar = QFrame(); bar.setFixedSize(4,20)
            bar.setStyleSheet(f"QFrame{{background:{color};border:none;border-radius:2px;}}")
            tl = QLabel(title)
            tl.setStyleSheet(
                f"QLabel{{font-size:14px;font-weight:700;"
                f"color:{P['text']};background:transparent;border:none;padding:0;}}"
            )
            tr.addWidget(bar); tr.addWidget(tl); tr.addStretch()
            cv.addLayout(tr); cv.addSpacing(18)
            for i,(st,sb,link) in enumerate(items):
                if i > 0: cv.addSpacing(14)
                nr = QHBoxLayout(); nr.setSpacing(12); nr.setContentsMargins(0,0,0,0)
                nm = QLabel(str(i+1)); nm.setFixedSize(24,24)
                nm.setAlignment(Qt.AlignmentFlag.AlignCenter)
                nm.setStyleSheet(
                    f"QLabel{{background:{color};color:#fff;"
                    f"font-size:11px;font-weight:700;"
                    f"border-radius:12px;border:none;padding:0;}}"
                )
                sl = QLabel(st)
                sl.setStyleSheet(
                    f"QLabel{{font-size:13px;font-weight:600;"
                    f"color:{P['text']};background:transparent;"
                    f"border:none;padding:0;}}"
                )
                nr.addWidget(nm); nr.addWidget(sl, 1)
                cv.addLayout(nr); cv.addSpacing(5)
                dl = QLabel(sb); dl.setWordWrap(True)
                dl.setStyleSheet(
                    f"QLabel{{font-size:12px;color:{P['sub2']};"
                    f"background:transparent;border:none;"
                    f"padding:0 0 0 36px;}}"
                )
                cv.addWidget(dl)
                if link:
                    import webbrowser as _wb
                    lrow = QHBoxLayout(); lrow.setContentsMargins(36,4,0,0)
                    lb = QPushButton(link)
                    lb.setStyleSheet(
                        f"QPushButton{{background:transparent;border:none;"
                        f"color:{color};font-size:11px;padding:0;text-align:left;}}"
                        f"QPushButton:hover{{color:{P['text']};}}"
                    )
                    lb.clicked.connect(lambda _,u=link: _wb.open(u))
                    lrow.addWidget(lb); lrow.addStretch(); cv.addLayout(lrow)
            sv.addWidget(card); pv.addWidget(sw)

        _sec(v, 'DeepL API 키 발급 (파일명 번역)', P['accent'], [
            ('DeepL 계정 만들기',
             'DeepL 공식 사이트에 접속해 우측 상단 가입 버튼으로 계정을 만드세요. Google 또는 이메일로 가입 가능합니다.',
             'https://www.deepl.com'),
            ('API Free 플랜 선택',
             '로그인 후 우측 상단 프로필 -> API 메뉴로 이동합니다. DeepL API Free 플랜을 선택하세요. 신용카드가 필요하지만 월 50만 글자까지 무료입니다.',
             'https://www.deepl.com/pro-api'),
            ('API 키 복사',
             'API 대시보드 하단 Authentication Key 항목에서 키를 복사합니다. 형식 예: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx:fx',
             ''),
            ('Filelo에 등록',
             '복사한 키를 API 키 설정 탭의 DeepL API 키 입력란에 붙여넣고 저장 버튼을 누르세요.',
             ''),
        ])

        _sec(v, 'Gemini API 키 발급 (AI 요약·초안·OCR)', P['success'], [
            ('Google AI Studio 접속',
             '기존 Google 계정(Gmail)으로 바로 로그인됩니다. 별도 가입이 필요 없습니다.',
             'https://aistudio.google.com'),
            ('Get API key 클릭',
             'AI Studio 좌측 메뉴에서 Get API key 버튼을 클릭합니다.',
             'https://aistudio.google.com/apikey'),
            ('새 키 생성',
             'Create API key 버튼을 클릭하고 프로젝트를 선택합니다. Gemini 2.0 Flash 모델은 무료 플랜에서 사용 가능합니다.',
             ''),
            ('Filelo에 등록',
             '생성된 키(AIzaSy... 형태)를 복사해 API 키 설정 탭의 Gemini 입력란에 붙여넣고 저장하세요.',
             ''),
        ])

        # 주의사항
        nw = QWidget(); nw.setStyleSheet(f"background:{P['bg']};")
        nvl = QVBoxLayout(nw); nvl.setContentsMargins(28,0,28,20)
        nc = QFrame()
        nc.setStyleSheet(f"QFrame{{background:{P['card2']};border:none;border-radius:12px;}}")
        ncv = QVBoxLayout(nc); ncv.setContentsMargins(20,16,20,16); ncv.setSpacing(6)
        nt = QLabel('  API 키 보안 주의사항')
        nt.setStyleSheet(f"QLabel{{font-size:13px;font-weight:700;color:{P['warning']};background:transparent;border:none;padding:0;}}")
        ncv.addWidget(nt)
        for ln in [
            'API 키는 비밀번호와 같습니다. 타인과 공유하지 마세요.',
            'Filelo는 AES-256-GCM 암호화로 이 PC에만 저장하며 외부로 전송하지 않습니다.',
            '키가 유출됐다면 즉시 해당 서비스 대시보드에서 삭제하고 새로 발급받으세요.',
        ]:
            nl = QLabel('•  ' + ln); nl.setWordWrap(True)
            nl.setStyleSheet(f"QLabel{{font-size:12px;color:{P['sub']};background:transparent;border:none;padding:0;}}")
            ncv.addWidget(nl)
        nvl.addWidget(nc); v.addWidget(nw)

        # Discord 지원 카드
        _dw = QWidget(); _dw.setStyleSheet(f"background:{P['bg']};")
        _dvl = QVBoxLayout(_dw); _dvl.setContentsMargins(28,0,28,28)
        _dc = QFrame()
        _dc.setStyleSheet(
            "QFrame{background:#5865F211;border:1px solid #5865F244;border-radius:14px;}"
        )
        _dcv = QHBoxLayout(_dc); _dcv.setContentsMargins(24,18,24,18); _dcv.setSpacing(16)
        _dl = QVBoxLayout(); _dl.setSpacing(5)
        _dt = QLabel("💬  고객센터 — Discord")
        _dt.setStyleSheet("font-size:14px;font-weight:700;color:#5865F2;background:transparent;")
        _dd = QLabel("사용 중 문제, 버그 신고, 기능 제안은 공식 Discord 서버로 문의해 주세요.\n개발자가 직접 응답합니다.")
        _dd.setWordWrap(True)
        _dd.setStyleSheet(f"font-size:12px;color:{P['sub']};background:transparent;")
        _du = QLabel("https://discord.gg/7agPwy9KRb")
        _du.setStyleSheet("font-size:11px;color:#7289DA;background:transparent;")
        _dl.addWidget(_dt); _dl.addWidget(_dd); _dl.addWidget(_du)
        _dcv.addLayout(_dl, 1)
        import webbrowser as _wb_hp
        _db = QPushButton("서버 참여하기  →")
        _db.setFixedHeight(38); _db.setMinimumWidth(140)
        _db.setCursor(Qt.CursorShape.PointingHandCursor)
        _db.setStyleSheet(
            "QPushButton{background:#5865F2;color:#fff;border:none;"
            "border-radius:9px;font-size:12px;font-weight:700;padding:0 18px;}"
            "QPushButton:hover{background:#4752C4;}"
        )
        _db.clicked.connect(lambda: _wb_hp.open("https://discord.gg/7agPwy9KRb"))
        _dcv.addWidget(_db)
        _dvl.addWidget(_dc); v.addWidget(_dw)
        v.addStretch()

class SettingsPage(Page):
    def __init__(self,app,p=None):
        super().__init__(p); self.app=app
        self.hdr("️","API 키 설정","DeepL · Gemini API 키를 안전하게 로컬에 저장합니다")
        self.tip([
            "API 키를 이 PC에만 AES-256-GCM 암호화로 안전하게 저장합니다.",
            "DeepL API 키: 파일명 번역 기능에 사용됩니다. (무료 플랜: 월 50만 글자)",
            "Gemini API 키: AI 요약·초안·OCR·참고문헌 정리에 사용됩니다. (무료 플랜 지원)",
            "키 파일은 이 PC에만 저장되며 외부로 전송되거나 다른 PC에서 복호화할 수 없습니다.",
        ])
        ci=self.card()

        def _api_row(label, placeholder):
            """API 키 입력 행 생성"""
            section_lbl = QLabel(label)
            section_lbl.setStyleSheet(
                f"font-size:12px;font-weight:600;color:{P['sub']};"
                f"background:transparent;border:none;"
            )
            edit = QLineEdit()
            edit.setEchoMode(QLineEdit.EchoMode.Password)
            edit.setPlaceholderText(placeholder)
            edit.setFixedHeight(38)
            edit.setStyleSheet(
                f"QLineEdit{{background:{P['input']};border:1.5px solid {P['border2']};"
                f"border-radius:9px;padding:0 14px;font-size:13px;color:{P['text']};}}"
                f"QLineEdit:focus{{border-color:{P['accent']};}}"
            )
            show_btn = FluidButton("보기", preset="ghost")
            show_btn.setFixedHeight(34)
            show_btn.setMinimumWidth(62)
            show_btn.setCheckable(True)
            show_btn.setStyleSheet(
                f"QPushButton{{background:{P['card2']};border:1px solid {P['border2']};"
                f"border-radius:8px;font-size:12px;font-weight:600;color:{P['sub']};"
                f"padding:0 10px;}}"
                f"QPushButton:checked{{background:{P['accent']}18;border-color:{P['accent']};"
                f"color:{P['accent']};}}"  
                f"QPushButton:hover{{background:{P['glass']};color:{P['text']};}}"  
            )
            show_btn.toggled.connect(
                lambda v, e=edit: e.setEchoMode(
                    QLineEdit.EchoMode.Normal if v else QLineEdit.EchoMode.Password
                )
            )
            # 클립보드 복사 버튼 (QTimer 객체를 버튼에 저장해 GC 방지)
            copy_btn = QPushButton("복사")
            copy_btn.setFixedHeight(34)
            copy_btn.setMinimumWidth(62)
            copy_btn.setStyleSheet(
                f"QPushButton{{background:{P['card2']};border:1px solid {P['border2']};"
                f"border-radius:8px;font-size:12px;font-weight:600;color:{P['sub']};"
                f"padding:0 10px;}}"
                f"QPushButton:hover{{background:{P['glass']};color:{P['text']};}}"  
            )
            _t = QTimer(); _t.setSingleShot(True); _t.setInterval(1400)
            copy_btn._rst = _t  # GC 방지
            _t.timeout.connect(lambda b=copy_btn: b.setText("복사") if b.isVisible() else None)
            def _do_copy(checked=False, e=edit, b=copy_btn, t=_t):
                txt = e.text().strip()
                if not txt: return
                QApplication.clipboard().setText(txt)
                b.setText("✓ 복사됨"); t.start()
            copy_btn.clicked.connect(_do_copy)
            row = QHBoxLayout(); row.setSpacing(6); row.setContentsMargins(0,0,0,0)
            row.addWidget(edit, 1); row.addWidget(show_btn); row.addWidget(copy_btn)
            return section_lbl, edit, row

        deepl_lbl, self._dk, deepl_row = _api_row(
            "DeepL API 키  —  파일명 번역 기능에 사용",
            "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx:fx"
        )
        gemini_lbl, self._gk, gemini_row = _api_row(
            "Gemini API 키  —  AI 요약·초안·OCR·참고문헌에 사용",
            "AIzaSy..."
        )
        ci.addWidget(deepl_lbl); ci.addLayout(deepl_row)
        ci.addSpacing(10)
        ci.addWidget(gemini_lbl); ci.addLayout(gemini_row)
        self._sl = QLabel()
        self._sl.setStyleSheet(
            f"QLabel{{color:{P['sub']};font-size:11px;background:transparent;border:none;padding:0;}}"
        )
        self._sl.setWordWrap(True)
        ci.addSpacing(4)
        ci.addWidget(self._sl)
        self._rs()
        br=QHBoxLayout()
        bs=_btn(" 저장","success"); bs.clicked.connect(self._save); br.addWidget(bs)
        bc=_btn("️ 키 초기화","danger"); bc.clicked.connect(self._clr); br.addWidget(bc); br.addStretch(); ci.addLayout(br)
        if DEEPL_KEY: self._dk.setText(DEEPL_KEY)
        if GEMINI_KEY: self._gk.setText(GEMINI_KEY)

        # ── 패키지 상태
        pkg_ci = self.card()
        pkg_ci.addWidget(QLabel("패키지 설치 상태", styleSheet=f"font-size:12px;font-weight:700;color:{P['text']};background:transparent;"))
        PKGS = [
            ("PyQt6",       True,         "UI 프레임워크"),
            ("deepl",       HAS_DEEPL,    "파일명 번역"),
            ("pymupdf",     HAS_FITZ,     "PDF 처리"),
            ("Pillow",      HAS_PIL,      "이미지 처리"),
            ("python-docx", HAS_DOCX,     "Word 문서"),
            ("openpyxl",    HAS_XL,       "엑셀 처리"),
            ("google-genai",HAS_GEMINI,   "AI 기능"),
            ("rembg",       HAS_REMBG,    "배경 제거 AI"),
            ("cryptography",HAS_CRYPTO,   "API 키 암호화"),
        ]
        pkg_grid = QGridLayout(); pkg_grid.setSpacing(6); pkg_grid.setContentsMargins(0,4,0,0)
        for i, (name, ok, desc) in enumerate(PKGS):
            dot_c = "#05C072" if ok else "#FF3B30"
            dot = QLabel("●")
            dot.setStyleSheet(
                f"QLabel{{color:{dot_c};font-size:12px;"
                f"background:transparent;border:none;padding:0;}}"
            )
            nm  = QLabel(name)
            nm.setStyleSheet(f"QLabel{{font-size:12px;font-weight:600;color:{P['text']};background:transparent;border:none;padding:0;}}")
            ds  = QLabel(desc)
            ds.setStyleSheet(f"QLabel{{font-size:11px;color:{P['sub']};background:transparent;border:none;padding:0;}}")
            row, col = divmod(i, 3)
            pkg_grid.addWidget(dot, row, col*3+0)
            pkg_grid.addWidget(nm,  row, col*3+1)
            pkg_grid.addWidget(ds,  row, col*3+2)
        pkg_ci.addLayout(pkg_grid)
        # Discord 문의 (설정 하단)
        _cw = QWidget(); _cw.setStyleSheet(f"background:{P['bg']};")
        _cvl = QVBoxLayout(_cw); _cvl.setContentsMargins(28,4,28,4)
        _cc = QFrame()
        _cc.setStyleSheet(
            f"QFrame{{background:{P['card']};border:1px solid {P['sep']};border-radius:12px;}}"
        )
        _ccv = QHBoxLayout(_cc); _ccv.setContentsMargins(18,12,18,12); _ccv.setSpacing(12)
        _ci = QLabel("💬"); _ci.setStyleSheet("font-size:18px;background:transparent;")
        _ct = QVBoxLayout(); _ct.setSpacing(2)
        _ctl = QLabel("문의 및 지원"); _ctl.setStyleSheet(f"font-size:13px;font-weight:600;color:{P['text']};background:transparent;")
        _csl = QLabel("버그 신고·기능 제안·사용법 문의는 Discord로"); _csl.setStyleSheet(f"font-size:11px;color:{P['sub']};background:transparent;")
        _ct.addWidget(_ctl); _ct.addWidget(_csl)
        _ccv.addWidget(_ci); _ccv.addLayout(_ct, 1)
        import webbrowser as _wb_cfg
        _cb = QPushButton("Discord 열기  →")
        _cb.setFixedHeight(32); _cb.setMinimumWidth(120)
        _cb.setCursor(Qt.CursorShape.PointingHandCursor)
        _cb.setStyleSheet(
            "QPushButton{background:#5865F2;color:#fff;border:none;"
            "border-radius:8px;font-size:11px;font-weight:700;padding:0 14px;}"
            "QPushButton:hover{background:#4752C4;}"
        )
        _cb.clicked.connect(lambda: _wb_cfg.open("https://discord.gg/7agPwy9KRb"))
        _ccv.addWidget(_cb)
        _cvl.addWidget(_cc)
        self._v.addWidget(_cw)
        self.stretch()
    def _rs(self):
        d = "저장됨 ✓" if DEEPL_KEY else "미등록"
        g = "저장됨 ✓" if GEMINI_KEY else "미등록"
        c = "AES-256-GCM 암호화" if HAS_CRYPTO else "⚠ cryptography 없음"
        self._sl.setText(
            f"DeepL: {d}  |  Gemini: {g}\n"
            f"보안: {c}\n"
            f"저장 위치: {CONFIG_FILE}"
        )
    def _save(self):
        global DEEPL_KEY,GEMINI_KEY,ai_client
        if not HAS_CRYPTO: self.app.toast("cryptography 패키지가 설치되지 않았습니다","err"); return
        nd=self._dk.text().strip(); ng=self._gk.text().strip()
        if not nd and not ng: self.app.toast("DeepL 또는 Gemini API 키를 입력해 주세요","err"); return
        cfg={}
        if nd: cfg["deepl_key"]=nd
        if ng: cfg["gemini_key"]=ng
        save_cfg(cfg); DEEPL_KEY=nd or DEEPL_KEY; GEMINI_KEY=ng or GEMINI_KEY
        if ng: init_gemini()
        self._rs()
        if hasattr(self.app,"_refresh_badges"): self.app._refresh_badges()
        self.app.toast("AES-256 암호화로 안전하게 저장되었습니다!","ok")
    def _clr(self):
        if QMessageBox.question(self,"확인","저장된 API 키를 모두 삭제하시겠습니까?")!=QMessageBox.StandardButton.Yes: return
        global DEEPL_KEY,GEMINI_KEY,ai_client
        try:
            if os.path.exists(CONFIG_FILE): os.remove(CONFIG_FILE)
        except: pass
        DEEPL_KEY=GEMINI_KEY=""; ai_client=None; self._dk.clear(); self._gk.clear(); self._rs()
        self.app._refresh_badges()
        self.app.toast("API 키가 초기화되었습니다.","warn")



# ══════════════════════════════════════════════════════════════
# 검색 오버레이 위젯
# ══════════════════════════════════════════════════════════════
class SearchOverlay(QFrame):
    """
    MainWindow centralWidget 위에 떠있는 검색 드롭다운.
    - 퍼지 매칭 (오타 허용)
    - 초성 검색
    - 카테고리 뱃지
    - 키보드 탐색 (↑↓ Enter Esc)
    - 최근 검색어 (최대 5개)
    - 매칭 텍스트 하이라이트
    """

    # 기능 DB: (이름, 설명, key, 카테고리, 태그목록)
    # 태그는 최대한 풍부하게 — 오타·유의어·영문·동사형 모두 포함
    FEATURES = [
        ("파일명 번역", "영문 파일명을 한글로 자동 번역", "translate", "파일 관리",
         ["번역","translate","영문","영어","한글","파일명","deepl","이름","name","파일","file","언어","english","korean"]),
        ("폴더 자동 정리", "확장자별 폴더 자동 분류", "folder", "파일 관리",
         ["폴더","folder","정리","분류","확장자","자동","organize","sort","파일","file","ext","extension","디렉토리","directory"]),
        ("파일명 일괄 변경", "날짜·번호 규칙으로 일괄 리네임", "rename", "파일 관리",
         ["파일명","이름","변경","rename","리네임","규칙","날짜","번호","일괄","batch","바꾸기","수정","rule"]),
        ("과제 폴더 생성", "과목별 폴더 구조 자동 생성", "task_dir", "파일 관리",
         ["과제","폴더","생성","만들기","과목","학교","구조","자동","task","dir","directory","수업","강의"]),
        ("PDF 변환 & 추출", "PDF에서 텍스트·이미지 추출", "pdf", "문서 처리",
         ["pdf","변환","추출","텍스트","이미지","extract","convert","문서","읽기","text","content"]),
        ("PDF 병합 / 분리", "여러 PDF 합치기·페이지 분리", "pdfmerge", "문서 처리",
         ["pdf","합치기","분리","병합","merge","split","페이지","page","붙이기","나누기","쪼개기","combine"]),
        ("PDF 비밀번호", "PDF 암호 설정 및 해제", "pdfpwd", "문서 처리",
         ["pdf","비밀번호","암호","password","잠금","lock","unlock","보안","security","설정","해제"]),
        ("메타데이터 삭제", "작성자·수정 이력 완전 제거", "meta", "문서 처리",
         ["메타","metadata","작성자","이력","개인정보","지문","삭제","제거","author","history","숨기기","익명"]),
        ("표 → 엑셀 변환", "문서 내 표를 xlsx로 추출", "table2xl", "문서 처리",
         ["표","table","엑셀","excel","xlsx","변환","추출","문서","워드","word","pdf","docx"]),
        ("이미지 일괄 처리", "리사이즈·포맷 변환·압축", "image", "이미지",
         ["이미지","사진","image","photo","크기","리사이즈","resize","포맷","변환","jpg","png","webp","압축","batch","일괄"]),
        ("이미지 → PDF", "여러 이미지를 PDF로 묶기", "imgpdf", "이미지",
         ["이미지","사진","pdf","변환","묶기","합치기","combine","image","photo","jpg","png"]),
        ("이미지 일괄 추출", "문서에서 이미지 추출", "imgext", "이미지",
         ["이미지","추출","extract","문서","docx","pdf","word","사진","꺼내기","저장"]),
        ("워터마크 삽입", "텍스트 워터마크 일괄 적용", "watermark", "이미지",
         ["워터마크","watermark","도장","텍스트","삽입","이미지","사진","복사방지","저작권","copyright"]),
        ("이미지 OCR", "AI 텍스트 인식 및 추출", "ocr", "이미지",
         ["ocr","텍스트","인식","추출","사진","스캔","scan","이미지","글자","읽기","image","ai"]),
        ("배경 제거", "AI 자동 배경 제거", "rembg", "이미지",
         ["배경","제거","remove","background","rembg","ai","이미지","사진","누끼","투명","transparent"]),
        ("AI 문서 요약", "Gemini로 핵심 내용 요약", "summary", "AI 도구",
         ["요약","summary","ai","gemini","문서","핵심","줄이기","정리","summarize","요점","brief"]),
        ("AI 문서 초안", "주제 입력 → 초안 자동 작성", "draft", "AI 도구",
         ["초안","draft","ai","작성","gemini","자동","생성","글쓰기","문서","보고서","이메일","report"]),
        ("참고문헌 정리", "APA·MLA 인용 형식 자동 변환", "citation", "AI 도구",
         ["참고문헌","인용","citation","apa","mla","논문","레퍼런스","reference","bibliography","형식"]),
        ("엑셀 자동화", "시트 합치기·중복 제거·변환", "excel", "데이터",
         ["엑셀","excel","xlsx","xls","시트","합치기","중복","제거","csv","변환","자동화","데이터","spreadsheet"]),
        ("마감 트래커", "과제 D-day 등록 및 관리", "tracker", "학습 관리",
         ["마감","트래커","tracker","과제","dday","일정","deadline","관리","학교","수업","달력","schedule"]),
        ("API 키 설정", "DeepL·Gemini API 키 등록", "settings", "설정",
         ["api","키","설정","deepl","gemini","key","등록","token","인증","configuration"]),
        ("도움말", "API 키 발급 방법 안내", "help", "설정",
         ["도움말","help","발급","사용법","가이드","guide","deepl","gemini","방법"]),
        ("고객센터 Discord", "문의 · 버그 · 기능제안", "discord", "설정",
         ["문의","discord","고객센터","지원","버그","신고","제안","디스코드","채팅"]),
    ]

    CAT_COLORS = {
        "파일 관리": "#3182F6",
        "문서 처리": "#FF9500",
        "이미지":    "#6366F1",
        "AI 도구":   "#3182F6",
        "데이터":    "#05C072",
        "학습 관리": "#FF3B30",
        "설정":      "#8E8E93",
    }

    def __init__(self, parent, search_input, nav_cb, refresh_cb=None):
        super().__init__(parent)
        self._search = search_input
        self._nav_cb = nav_cb
        self._refresh_cb = refresh_cb
        self._cur_idx = -1
        self._btns = []
        self._history = []  # 최근 검색어

        self.setStyleSheet(
            f"QFrame#searchDrop{{background:{P['glass']};border:1px solid {P['border2']};"
            f"border-radius:12px;}}"
            f"QFrame{{border:none;background:transparent;}}"
            f"QLabel{{border:none;background:transparent;}}"
            f"QPushButton{{border:none;}}"
        )
        self.setObjectName("searchDrop")
        self.hide()

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # 스크롤 영역
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._scroll.setStyleSheet(
            "QScrollArea{border:none;background:transparent;}"
            "QScrollBar:vertical{width:4px;background:transparent;border:none;}"
            f"QScrollBar::handle:vertical{{background:{P['border2']};border-radius:2px;}}"
            "QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical{height:0;}"
        )
        self._inner = QWidget()
        self._inner.setStyleSheet("background:transparent;")
        self._layout = QVBoxLayout(self._inner)
        self._layout.setContentsMargins(8, 8, 8, 8)
        self._layout.setSpacing(2)
        self._scroll.setWidget(self._inner)
        outer.addWidget(self._scroll)

        # 하단 힌트
        hint = QLabel("↑↓ 탐색    Enter 선택    Esc 닫기")
        hint.setStyleSheet(
            f"color:{P['sub2']};font-size:10px;padding:6px 14px;"
            f"border-top:1px solid {P['border']};background:transparent;"
        )
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        outer.addWidget(hint)

    # ── 강화된 스코어 계산
    @staticmethod
    def _score(query, name, desc, tags):
        """
        다층 스코어링:
        1. 완전 일치 (최우선)
        2. 이름 시작 일치
        3. 이름/설명/태그 포함
        4. 단어 단위 매칭 (공백 분리)
        5. Subsequence (오타 허용)
        6. 초성/접두어 부스트
        """
        q = query.lower().strip()
        if not q:
            return 0
        n = name.lower()
        d = desc.lower()
        tag_str = " ".join(tags).lower()

        score = 0
        words = q.split()

        # ① 완전 일치
        if q == n:                          score += 200
        # ② 이름 시작
        elif n.startswith(q):               score += 100
        # ③ 이름 포함
        elif q in n:                        score += 70
        # ④ 설명 포함
        if q in d:                          score += 35
        # ⑤ 태그 완전 일치 (단어 단위)
        for tag in tags:
            tl = tag.lower()
            if q == tl:                     score += 90
            elif tl.startswith(q):          score += 50
            elif q in tl:                   score += 25
        # ⑥ 다중 단어 — 각 단어가 이름/설명/태그에 있으면 가산
        if len(words) > 1:
            matched = sum(
                1 for w in words
                if w in n or w in d or w in tag_str
            )
            score += matched * 20
            if matched == len(words):       score += 30  # 전부 매칭 보너스
        else:
            # 단일 단어 부분 일치
            for tag in tags:
                if words[0] in tag.lower(): score += 15
        # ⑦ Subsequence (오타/축약 허용)
        target = n + " " + tag_str
        ti, sub_score = 0, 0
        for ch in q:
            idx = target.find(ch, ti)
            if idx != -1:
                sub_score += 1
                ti = idx + 1
        if sub_score == len(q):             score += 10  # 완전 subsequence
        elif sub_score >= len(q) * 0.8:     score += 5
        # ⑧ 1글자 검색 — 이름/설명/태그 어디든 포함되면 표시
        if len(q) == 1:
            if q in n:                              score += 30
            if q in d:                              score += 15
            for tag in tags:
                if q in tag.lower():                score += 10; break
        return score

    # ── 매칭 텍스트 하이라이트 (완전한 HTML)
    @staticmethod
    def _highlight(text, query, hl_color):
        import re, html as _h
        safe = _h.escape(text)
        q = re.escape(_h.escape(query))
        result = re.sub(
            f"({q})",
            lambda m: f'<span style="color:{hl_color};font-weight:700;">{m.group(0)}</span>',
            safe, flags=re.IGNORECASE
        )
        tc = P["text"]
        return f'<span style="color:{tc};font-size:13px;font-weight:600;">{result}</span>' 

    def update_results(self, query):
        # 기존 버튼 제거
        while self._layout.count():
            item = self._layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._btns.clear()
        self._cur_idx = -1

        query = query.strip()

        # 빈 쿼리 → 최근 검색어 표시
        if not query:
            if self._history:
                hdr = QLabel("최근 검색")
                hdr.setStyleSheet(
                    f"color:{P['sub2']};font-size:10px;font-weight:700;"
                    f"letter-spacing:0.8px;padding:4px 8px 2px 8px;background:transparent;"
                )
                self._layout.addWidget(hdr)
                for h_text, h_key in self._history[-5:][::-1]:
                    self._add_row(h_text, "최근 검색", h_key, P["sub"], query="")
                self._layout.addStretch()
                self._inner.adjustSize()
                self._inner.updateGeometry()
                self._reposition()
                self.show(); self.raise_()
            else:
                self.hide()
            return

        # 스코어 기반 정렬
        scored = []
        for name, desc, key, cat, tags in self.FEATURES:
            s = self._score(query, name, desc, tags)
            if s > 0:
                scored.append((s, name, desc, key, cat))
            # 빈 쿼리가 아닌데 0점이어도 이름에 한 글자라도 있으면 후보
            elif len(query) >= 2 and any(query[i:i+2] in (name+desc).lower() for i in range(len(query)-1)):
                scored.append((3, name, desc, key, cat))
        scored.sort(key=lambda x: -x[0])

        if not scored:
            no = QLabel(f'"{query}" 에 해당하는 기능이 없습니다')
            no.setStyleSheet(
                f"color:{P['sub']};font-size:12px;padding:12px 10px;background:transparent;"
            )
            no.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self._layout.addWidget(no)
        else:
            for score, name, desc, key, cat in scored[:10]:
                cat_color = self.CAT_COLORS.get(cat, P["accent"])
                self._add_row(name, cat, key, cat_color, query, desc)

        self._layout.addStretch()
        # inner 위젯 크기 강제 업데이트 → 스크롤 영역이 올바른 크기를 인식
        self._inner.adjustSize()
        self._inner.updateGeometry()
        self._reposition()
        self.show()
        self.raise_()

    def _add_row(self, name, cat, key, cat_color, query="", desc=""):
        btn = QPushButton()
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setMinimumHeight(48)
        btn.setStyleSheet(
            f"QPushButton{{background:transparent;border:none;border-radius:8px;"
            f"text-align:left;padding:0;}}"
            f"QPushButton:hover{{background:{P['hover']};}}"
        )

        bl = QHBoxLayout(btn)
        bl.setContentsMargins(14, 0, 14, 0)
        bl.setSpacing(10)

        # 텍스트 컨테이너
        txt = QWidget()
        txt.setStyleSheet("background:transparent;border:none;")
        txt.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        tv = QVBoxLayout(txt)
        tv.setContentsMargins(0, 0, 0, 0)
        tv.setSpacing(2)

        # 기능명 — 잘리지 않게 WordWrap 적용
        name_lbl = QLabel()
        name_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        name_lbl.setWordWrap(False)
        name_lbl.setTextFormat(Qt.TextFormat.RichText)
        name_lbl.setSizePolicy(
            __import__("PyQt6.QtWidgets", fromlist=["QSizePolicy"]).QSizePolicy.Policy.Expanding,
            __import__("PyQt6.QtWidgets", fromlist=["QSizePolicy"]).QSizePolicy.Policy.Preferred,
        )
        name_lbl.setStyleSheet(
            f"font-size:13px;font-weight:600;color:{P['text']};"
            f"background:transparent;border:none;"
        )
        if query:
            name_lbl.setText(self._highlight(name, query, P["accent"]))
        else:
            name_lbl.setText(name)
        tv.addWidget(name_lbl)

        # 설명
        if desc:
            desc_lbl = QLabel(desc)
            desc_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
            desc_lbl.setStyleSheet(
                f"font-size:10px;color:{P['sub']};background:transparent;border:none;"
            )
            tv.addWidget(desc_lbl)

        bl.addWidget(txt, 1)

        # 카테고리 텍스트 (배경 없이 색상만)
        cat_lbl = QLabel(cat)
        cat_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        cat_lbl.setStyleSheet(
            f"color:{cat_color};font-size:10px;font-weight:500;"
            f"background:transparent;border:none;"
        )
        bl.addWidget(cat_lbl, 0)

        btn.clicked.connect(lambda _, k=key, nm=name: self._select(k, nm))
        self._layout.addWidget(btn)
        self._btns.append(btn)

    def _select(self, key, name):
        # 사용 빈도 기록
        record_usage(key)
        # 최근 검색어 저장
        entry = (name, key)
        if entry in self._history:
            self._history.remove(entry)
        self._history.append(entry)
        if len(self._history) > 8:
            self._history.pop(0)
        self._search.clear()
        self.hide()
        if key == "discord":
            import webbrowser as _wb_sel
            _wb_sel.open("https://discord.gg/7agPwy9KRb")
        else:
            self._nav_cb(key)
        # 홈 태그 갱신 콜백
        if self._refresh_cb:
            self._refresh_cb()

    def move_cursor(self, delta):
        if not self._btns:
            return
        prev = self._cur_idx
        self._cur_idx = max(0, min(len(self._btns) - 1, self._cur_idx + delta))
        if prev >= 0 and prev < len(self._btns):
            self._btns[prev].setStyleSheet(
                f"QWidget{{background:transparent;border-radius:8px;}}"
                f"QWidget:hover{{background:{P['hover']};}}"
            )
        self._btns[self._cur_idx].setStyleSheet(
            f"QWidget{{background:{P['active']};border-radius:8px;"
            f"border-left:2px solid {P['accent']};}}"
        )
        self._scroll.ensureWidgetVisible(self._btns[self._cur_idx])

    def select_current(self):
        if 0 <= self._cur_idx < len(self._btns):
            self._btns[self._cur_idx].click()

    def _reposition(self):
        parent = self.parent()
        if parent is None or not self._search.isVisible():
            return
        # 검색창의 좌하단/우하단 좌표를 parent 기준으로 변환
        global_bl = self._search.mapToGlobal(self._search.rect().bottomLeft())
        global_br = self._search.mapToGlobal(self._search.rect().bottomRight())
        local_bl  = parent.mapFromGlobal(global_bl)
        local_br  = parent.mapFromGlobal(global_br)

        # 드롭다운 너비 = 검색창 너비와 완전히 동일
        drop_w = local_br.x() - local_bl.x()
        x = local_bl.x()
        y = local_bl.y() + 4
        max_h = min(480, parent.height() - y - 20)

        self.setFixedWidth(max(drop_w, 200))
        self.setMinimumHeight(min(160, max_h))
        self.setMaximumHeight(max_h)
        self.move(x, y)
        self.raise_()

    def reposition(self):
        self._reposition()

# ══════════════════════════════════════════════════════════════
# 메인 윈도우
# ══════════════════════════════════════════════════════════════
# ── 이용 약관 및 법적 동의 창 ────────────────────────────
class ConsentDialog(QDialog):
    """최초 실행 시 1회만 표시되는 법적 동의 창"""

    # ── 약관 섹션 데이터 ──────────────────────────────────
    SECTIONS = [
        {
            "id":    "privacy",
            "icon":  "🔒",
            "title": "개인정보 처리방침",
            "law":   "개인정보 보호법 제30조",
            "color": "#3182F6",
            "body":  (
                "■ 로컬 처리 원칙\n"
                "Filelo는 사용자의 파일을 외부 서버로 전송하지 않습니다. "
                "파일 정리·이름 변경·PDF 처리·이미지 편집 등 모든 핵심 기능은 "
                "사용자의 PC 내에서만 처리됩니다.\n\n"
                "■ 외부 API 전송 항목\n"
                "아래 기능은 사용자가 직접 API 키를 등록한 경우에 한해 "
                "해당 파일 내용이 외부 서버로 전송됩니다.\n"
                "  • DeepL API — 파일명 번역 시 파일명(텍스트)이 DeepL 서버로 전송됩니다.\n"
                "  • Google Gemini API — AI 요약·초안 생성·OCR·참고문헌 정리 시 "
                "문서 내용 또는 이미지가 Google 서버로 전송됩니다.\n\n"
                "■ 외부 서비스 책임\n"
                "외부 API 사용 시 각 서비스의 개인정보 처리방침이 적용되며, "
                "해당 서비스로 전송된 데이터에 관한 책임은 각 서비스 공급자에게 있습니다.\n"
                "  • DeepL 개인정보처리방침: https://www.deepl.com/ko/privacy\n"
                "  • Google 개인정보처리방침: https://policies.google.com/privacy\n\n"
                "■ API 키 보안\n"
                "등록된 API 키는 AES-256-GCM 암호화 및 PBKDF2(480,000회) 해싱을 통해 "
                "이 PC에만 저장되며, 외부로 전송되거나 제3자에게 제공되지 않습니다.\n\n"
                "■ 수집 데이터\n"
                "본 프로그램은 기능 사용 빈도(기능명만, 파일 내용 미포함)를 "
                "로컬 파일(.filelo_usage.json)에 저장하여 '자주 쓰는 기능' 표시에 활용합니다. "
                "이 데이터는 외부로 전송되지 않습니다."
            )
        },
        {
            "id":    "disclaimer",
            "icon":  "⚠️",
            "title": "면책 조항",
            "law":   "민법 제750조 / 소프트웨어 관련 법령",
            "color": "#FF9500",
            "body":  (
                "■ 무보증 원칙 (AS-IS)\n"
                "본 소프트웨어는 '있는 그대로(AS-IS)' 제공되며, 명시적 또는 "
                "묵시적 보증 없이 제공됩니다. 특정 목적에의 적합성, 오류 없음, "
                "중단 없는 서비스를 보증하지 않습니다.\n\n"
                "■ 데이터 손실 면책\n"
                "본 소프트웨어 사용으로 인해 발생하는 파일 손상, 데이터 손실, "
                "시스템 오류, 업무 중단 등 직접적·간접적·우발적·특수적 손해에 대해 "
                "개발자는 어떠한 법적 책임도 지지 않습니다.\n\n"
                "■ 백업 권장\n"
                "파일명 변경, PDF 병합/분리, 메타데이터 삭제 등 파일을 직접 수정하는 "
                "기능 사용 전 반드시 원본 파일의 백업을 권장합니다. "
                "일부 작업(파일명 변경 등)은 되돌리기가 불가능합니다.\n\n"
                "■ AI 생성 콘텐츠\n"
                "AI 문서 요약, 초안 생성, 참고문헌 정리 등 AI 기반 기능의 결과물은 "
                "참고용으로만 사용하십시오. AI 생성 결과의 정확성, 완전성, 적법성에 대해 "
                "개발자는 책임지지 않으며, 최종 확인 및 검토는 사용자의 책임입니다.\n\n"
                "■ 책임 한도\n"
                "어떠한 경우에도 개발자의 최대 배상 책임은 사용자가 본 소프트웨어에 대해 "
                "지불한 금액(무료 배포 시 KRW 0)을 초과하지 않습니다."
            )
        },
        {
            "id":    "opensource",
            "icon":  "📦",
            "title": "오픈소스 라이선스",
            "law":   "저작권법 제37조",
            "color": "#05C072",
            "body":  (
                "본 프로그램은 아래 오픈소스 소프트웨어를 사용합니다.\n"
                "각 패키지의 라이선스 전문은 해당 프로젝트 저장소에서 확인할 수 있습니다.\n\n"
                "┌─────────────────┬──────────┬─────────────────────────────┐\n"
                "│ 패키지           │ 라이선스 │ 용도                         │\n"
                "├─────────────────┼──────────┼─────────────────────────────┤\n"
                "│ PyQt6           │ GPL v3   │ GUI 프레임워크                │\n"
                "│ Pillow          │ HPND     │ 이미지 처리                  │\n"
                "│ PyMuPDF (fitz)  │ AGPL v3  │ PDF 처리                    │\n"
                "│ python-docx     │ MIT      │ Word 문서 처리               │\n"
                "│ openpyxl        │ MIT      │ 엑셀 처리                   │\n"
                "│ deepl           │ MIT      │ DeepL 번역 API               │\n"
                "│ google-genai    │ Apache 2 │ Google Gemini AI             │\n"
                "│ rembg           │ MIT      │ AI 배경 제거                 │\n"
                "│ cryptography    │ Apache 2 │ AES-256 암호화               │\n"
                "└─────────────────┴──────────┴─────────────────────────────┘\n\n"
                "■ 폰트 관련\n"
                "본 프로그램은 시스템 기본 폰트(Malgun Gothic 등)를 렌더링에 사용하나, "
                "폰트 파일을 프로그램 내 포함(embedding)하거나 배포하지 않습니다. "
                "폰트 저작권은 각 폰트 제작사에 있습니다.\n\n"
                "■ PyQt6 및 PyMuPDF 라이선스 안내\n"
                "PyQt6(GPL v3)과 PyMuPDF(AGPL v3)는 강한 카피레프트 라이선스를 채택합니다. "
                "본 프로그램을 수정·배포할 경우 해당 라이선스 조건을 준수해야 합니다."
            )
        },
        {
            "id":    "security",
            "icon":  "🛡️",
            "title": "보안 및 무결성",
            "law":   "정보통신망법 제48조",
            "color": "#6366F1",
            "body":  (
                "■ 악성코드 관련\n"
                "Filelo는 악성코드를 포함하지 않습니다. 그러나 서명되지 않은 실행 파일의 특성상 "
                "일부 백신 소프트웨어가 오탐(False Positive)할 수 있습니다.\n\n"
                "■ 파일 무결성 확인\n"
                "배포된 파일의 SHA-256 해시값을 공개하여 다운로드한 파일이 "
                "변조되지 않았음을 확인할 수 있습니다.\n"
                "  • 공식 배포처: https://github.com/joys000/Filelo\n"
                "  • 해시 확인 방법 (PowerShell):\n"
                "    Get-FileHash filelo.py -Algorithm SHA256\n\n"
                "■ 코드 공개\n"
                "본 프로그램의 소스 코드는 GitHub에 공개되어 있어 "
                "누구든 코드를 직접 검토할 수 있습니다.\n\n"
                "■ 자동 업데이트 없음\n"
                "본 프로그램은 자동 업데이트 기능을 포함하지 않으며, "
                "사용자 동의 없이 어떠한 파일도 다운로드하거나 실행하지 않습니다.\n\n"
                "■ 신고 및 문의\n"
                "보안 취약점 발견 시 GitHub Issues를 통해 신고해 주시면\n"
                "신속하게 대응하겠습니다: https://github.com/joys000/Filelo/issues"
            )
        },
        {
            "id":    "terms",
            "icon":  "📋",
            "title": "이용 약관",
            "law":   "전자상거래법 / 약관규제법",
            "color": "#8E8E93",
            "body":  (
                "■ 서비스 제공\n"
                "본 소프트웨어는 개인 및 소규모 업무 사용을 목적으로 "
                "무료로 제공됩니다.\n\n"
                "■ 허용 사용\n"
                "  • 개인적 용도의 파일 관리 및 자동화\n"
                "  • 비영리 목적의 사용 및 내부 업무 지원\n"
                "  • 오픈소스 라이선스 조건 준수 하의 수정·배포\n\n"
                "■ 금지 사용\n"
                "  • 불법적인 파일 처리 또는 타인의 저작권 침해 목적 사용\n"
                "  • 본 프로그램을 이용한 개인정보 무단 수집·가공\n"
                "  • 악의적 목적의 자동화 사용\n\n"
                "■ 약관 변경\n"
                "약관이 변경될 경우 프로그램 업데이트 시 재동의를 요청합니다. "
                f"현재 약관 버전: {CONSENT_VERSION} (2025년 기준)\n\n"
                "■ 준거법 및 관할\n"
                "본 약관은 대한민국 법률에 따라 해석되며, "
                "분쟁 발생 시 서울중앙지방법원을 제1심 관할 법원으로 합니다.\n\n"
                "■ 개발자 정보\n"
                "본 프로그램은 개인 개발자가 제공하는 소프트웨어입니다.\n"
                "문의: https://discord.gg/7agPwy9KRb  |  https://github.com/joys000/Filelo"
            )
        },
    ]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Filelo — 이용 약관 및 개인정보 처리방침")
        self.setWindowFlags(
            Qt.WindowType.Dialog |
            Qt.WindowType.WindowTitleHint |
            Qt.WindowType.WindowCloseButtonHint
        )
        self.setModal(True)
        self.setMinimumSize(800, 640)
        self.resize(860, 700)
        self.setStyleSheet(f"background:{P['bg']};color:{P['text']};")

        # ── 화면 중앙 배치
        from PyQt6.QtWidgets import QApplication as _QA
        scr = _QA.primaryScreen().geometry()
        self.move((scr.width()-self.width())//2, (scr.height()-self.height())//2)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── 헤더
        hdr = QWidget(); hdr.setStyleSheet(f"background:{P['side']};border-bottom:1px solid {P['sep']};")
        hdr.setFixedHeight(72)
        hl = QHBoxLayout(hdr); hl.setContentsMargins(32, 0, 32, 0); hl.setSpacing(14)
        icon_lbl = QLabel("⚖️"); icon_lbl.setStyleSheet("font-size:28px;background:transparent;")
        ttl = QLabel("이용 약관 및 법적 고지")
        ttl.setStyleSheet(f"font-size:17px;font-weight:700;color:{P['text']};background:transparent;letter-spacing:-0.4px;")
        sub = QLabel("Filelo를 사용하기 전에 아래 약관을 모두 읽어주세요")
        sub.setStyleSheet(f"font-size:12px;color:{P['sub']};background:transparent;")
        txt_col = QVBoxLayout(); txt_col.setSpacing(2)
        txt_col.addWidget(ttl); txt_col.addWidget(sub)
        hl.addWidget(icon_lbl); hl.addLayout(txt_col); hl.addStretch()
        ver_lbl = QLabel(f"약관 버전 {CONSENT_VERSION}")
        ver_lbl.setStyleSheet(f"font-size:10px;color:{P['sub2']};background:transparent;")
        hl.addWidget(ver_lbl)
        root.addWidget(hdr)

        # ── 본문 (탭 + 내용)
        body = QWidget(); body.setStyleSheet(f"background:{P['bg']};")
        bl = QHBoxLayout(body); bl.setContentsMargins(0,0,0,0); bl.setSpacing(0)

        # 왼쪽 탭 목록
        tab_panel = QWidget(); tab_panel.setFixedWidth(196)
        tab_panel.setStyleSheet(f"background:{P['side']};border-right:1px solid {P['sep']};")
        tpl = QVBoxLayout(tab_panel); tpl.setContentsMargins(0,12,0,12); tpl.setSpacing(2)

        self._tab_btns = []
        for i, sec in enumerate(self.SECTIONS):
            tb = QPushButton(f"  {sec['icon']}  {sec['title']}")
            tb.setFixedHeight(40)
            tb.setCursor(Qt.CursorShape.PointingHandCursor)
            tb.setCheckable(True)
            tb.clicked.connect(lambda _, idx=i: self._switch_tab(idx))
            tb.setStyleSheet(
                f"QPushButton{{background:transparent;border:none;border-left:3px solid transparent;"
                f"text-align:left;padding:0 12px;font-size:12px;color:{P['sub']};}}"
                f"QPushButton:checked{{background:{P['accent']}15;border-left:3px solid {P['accent']};"
                f"color:{P['text']};font-weight:600;}}"
                f"QPushButton:hover:!checked{{background:{P['hover']};color:{P['text']};}}"
            )
            self._tab_btns.append(tb)
            tpl.addWidget(tb)

        # 읽음 표시 (각 섹션 읽으면 체크)
        self._read = set()
        tpl.addStretch()

        # 읽은 섹션 수 표시
        self._read_lbl = QLabel("0 / 5 섹션 확인")
        self._read_lbl.setStyleSheet(f"color:{P['sub2']};font-size:10px;padding:8px 16px;background:transparent;")
        tpl.addWidget(self._read_lbl)
        bl.addWidget(tab_panel)

        # 오른쪽 콘텐츠
        content_wrap = QWidget(); content_wrap.setStyleSheet(f"background:{P['bg']};")
        cl = QVBoxLayout(content_wrap); cl.setContentsMargins(0,0,0,0); cl.setSpacing(0)

        self._stack = QStackedWidget()
        self._stack.setStyleSheet(f"background:{P['bg']};")
        for sec in self.SECTIONS:
            page = self._make_section_page(sec)
            self._stack.addWidget(page)
        cl.addWidget(self._stack, 1)
        bl.addWidget(content_wrap, 1)
        root.addWidget(body, 1)

        # ── 하단 동의 영역
        ftr = QWidget()
        ftr.setStyleSheet(f"background:{P['side']};border-top:1px solid {P['sep']};")
        fl = QVBoxLayout(ftr); fl.setContentsMargins(32, 16, 32, 20); fl.setSpacing(10)

        # 체크박스들
        chk_style = (
            f"QCheckBox{{color:{P['text']};font-size:13px;spacing:10px;background:transparent;}}"
            f"QCheckBox::indicator{{width:18px;height:18px;border:2px solid {P['border2']};"
            f"border-radius:5px;background:{P['input']};}}"
            f"QCheckBox::indicator:checked{{background:{P['accent']};border-color:{P['accent']};}}"
        )
        self._chk1 = QCheckBox("위의 이용 약관 및 개인정보 처리방침을 모두 읽었으며 이에 동의합니다. (필수)")
        self._chk2 = QCheckBox("외부 API(DeepL, Google Gemini) 사용 시 데이터가 외부로 전송될 수 있음을 이해합니다. (필수)")
        self._chk3 = QCheckBox("본 소프트웨어 사용으로 인한 데이터 손실 등의 책임이 사용자에게 있음을 확인합니다. (필수)")
        for chk in [self._chk1, self._chk2, self._chk3]:
            chk.setStyleSheet(chk_style)
            chk.stateChanged.connect(self._update_agree_btn)
            fl.addWidget(chk)

        # 버튼 행
        btn_row = QHBoxLayout(); btn_row.setSpacing(10)
        self._all_chk = QCheckBox("위 항목 모두 동의")
        self._all_chk.setStyleSheet(chk_style)
        self._all_chk.stateChanged.connect(self._toggle_all)
        btn_row.addWidget(self._all_chk)
        btn_row.addStretch()

        refuse_btn = QPushButton("동의하지 않음 (종료)")
        refuse_btn.setFixedHeight(40)
        refuse_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        refuse_btn.setStyleSheet(
            f"QPushButton{{background:transparent;border:1px solid {P['border2']};"
            f"border-radius:10px;color:{P['sub']};font-size:13px;padding:0 20px;}}"
            f"QPushButton:hover{{border-color:{P['danger']};color:{P['danger']};}}"
        )
        refuse_btn.clicked.connect(self._refuse)

        self._agree_btn = QPushButton("동의하고 시작하기  →")
        self._agree_btn.setFixedHeight(40)
        self._agree_btn.setMinimumWidth(200)
        self._agree_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._agree_btn.setEnabled(False)
        self._agree_btn.setStyleSheet(
            f"QPushButton{{background:{P['accent']};color:#fff;border:none;"
            f"border-radius:10px;font-size:13px;font-weight:700;padding:0 24px;}}"
            f"QPushButton:hover{{background:{P['accent_h']};}}"
            f"QPushButton:disabled{{background:{P['border']};color:{P['sub2']};}}"
        )
        self._agree_btn.clicked.connect(self._agree)

        btn_row.addWidget(refuse_btn)
        btn_row.addWidget(self._agree_btn)
        fl.addLayout(btn_row)
        root.addWidget(ftr)

        # 초기 탭
        self._switch_tab(0)

    def _make_section_page(self, sec: dict) -> QWidget:
        page = QScrollArea()
        page.setWidgetResizable(True)
        page.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        page.setStyleSheet("QScrollArea{border:none;background:transparent;}")

        inner = QWidget(); inner.setStyleSheet(f"background:{P['bg']};")
        v = QVBoxLayout(inner); v.setContentsMargins(32, 28, 32, 28); v.setSpacing(0)

        # 섹션 헤더
        accent = QFrame(); accent.setFixedHeight(3)
        accent.setStyleSheet(f"background:{sec['color']};border-radius:1px;")
        v.addWidget(accent); v.addSpacing(16)

        title_row = QHBoxLayout(); title_row.setSpacing(10)
        ic = QLabel(sec["icon"]); ic.setStyleSheet("font-size:22px;background:transparent;")
        tl = QLabel(sec["title"])
        tl.setStyleSheet(f"font-size:17px;font-weight:700;color:{P['text']};background:transparent;letter-spacing:-0.3px;")
        law = QLabel(f"근거: {sec['law']}")
        law.setStyleSheet(f"font-size:11px;color:{P['sub2']};background:transparent;")
        tc = QVBoxLayout(); tc.setSpacing(2); tc.addWidget(tl); tc.addWidget(law)
        title_row.addWidget(ic); title_row.addLayout(tc); title_row.addStretch()
        v.addLayout(title_row); v.addSpacing(20)

        # 본문 (QLabel로 스크롤)
        body_lbl = QLabel(sec["body"])
        body_lbl.setWordWrap(True)
        body_lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        body_lbl.setStyleSheet(
            f"font-size:12px;color:{P['text']};background:transparent;"
            f"line-height:1.8;letter-spacing:-0.1px;"
        )
        v.addWidget(body_lbl)
        v.addStretch()
        page.setWidget(inner)

        # 스크롤 시 '읽음' 마킹
        sb = page.verticalScrollBar()
        sec_id = sec["id"]
        def _on_scroll(val, _id=sec_id, _sb=sb):
            if val >= _sb.maximum() * 0.85 or _sb.maximum() == 0:
                self._mark_read(_id)
        sb.valueChanged.connect(_on_scroll)
        # 짧으면 즉시 읽음
        QTimer.singleShot(300, lambda _id=sec_id, _sb=sb: (
            self._mark_read(_id) if _sb.maximum() < 10 else None
        ))
        return page

    def _mark_read(self, sec_id: str):
        if sec_id not in self._read:
            self._read.add(sec_id)
            self._read_lbl.setText(f"{len(self._read)} / {len(self.SECTIONS)} 섹션 확인")
            # 탭 버튼에 체크 표시
            for i, sec in enumerate(self.SECTIONS):
                if sec["id"] == sec_id:
                    btn = self._tab_btns[i]
                    if "✓" not in btn.text():
                        btn.setText(btn.text() + "  ✓")
                    break

    def _switch_tab(self, idx: int):
        self._stack.setCurrentIndex(idx)
        for i, btn in enumerate(self._tab_btns):
            btn.setChecked(i == idx)
        # 짧은 콘텐츠면 즉시 읽음 처리
        sec_id = self.SECTIONS[idx]["id"]
        QTimer.singleShot(400, lambda: self._mark_read(sec_id)
            if self._stack.currentWidget().verticalScrollBar
            and callable(getattr(self._stack.currentWidget(), "verticalScrollBar", None))
            and self._stack.currentWidget().verticalScrollBar().maximum() < 10
            else None
        )

    def _update_agree_btn(self):
        all_checked = (self._chk1.isChecked() and
                       self._chk2.isChecked() and
                       self._chk3.isChecked())
        self._agree_btn.setEnabled(all_checked)
        # 전체동의 체크박스 동기화
        self._all_chk.blockSignals(True)
        self._all_chk.setChecked(all_checked)
        self._all_chk.blockSignals(False)

    def _toggle_all(self, state):
        checked = state == 2  # Qt.CheckState.Checked
        for chk in [self._chk1, self._chk2, self._chk3]:
            chk.blockSignals(True)
            chk.setChecked(checked)
            chk.blockSignals(False)
        self._agree_btn.setEnabled(checked)

    def _agree(self):
        _save_consent()
        self.accept()

    def _refuse(self):
        from PyQt6.QtWidgets import QApplication as _QA
        self.reject()
        _QA.quit()

    def closeEvent(self, e):
        # X 버튼 → 종료
        from PyQt6.QtWidgets import QApplication as _QA
        _QA.quit()


# ── 스플래시 스크린 ────────────────────────────────────
class SplashScreen(QWidget):
    """앱 시작 시 3초간 표시되는 스플래시 스크린"""

    def __init__(self, icon_b64: str):
        super().__init__()
        from PyQt6.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve
        from PyQt6.QtGui import QFont, QColor, QPainter, QPainterPath, QLinearGradient, QPixmap, QImage
        from PyQt6.QtWidgets import QGraphicsOpacityEffect
        import base64

        # ── 창 설정 (프레임 없음, 반투명 배경)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self.setFixedSize(360, 280)

        # 화면 중앙 배치
        from PyQt6.QtWidgets import QApplication as _QA
        screen = _QA.primaryScreen().geometry()
        self.move(
            (screen.width()  - self.width())  // 2,
            (screen.height() - self.height()) // 2
        )

        # ── 아이콘 이미지 로드 (ICO base64 → PIL → QPixmap)
        self._icon_px = None
        try:
            from PIL import Image as _PI
            import io as _io
            raw = base64.b64decode(icon_b64)
            pil_img = _PI.open(_io.BytesIO(raw)).convert("RGBA")
            pil_img = pil_img.resize((100, 100), _PI.LANCZOS)
            data = pil_img.tobytes("raw", "RGBA")
            qi = QImage(data, 100, 100, QImage.Format.Format_RGBA8888)
            self._icon_px = QPixmap.fromImage(qi)
        except Exception:
            self._icon_px = None

        # ── 진행 상태
        self._progress   = 0      # 0~100
        self._step_idx   = 0
        self._steps = [
            (0,  "패키지 확인 중..."),
            (25, "API 설정 로드 중..."),
            (55, "UI 초기화 중..."),
            (80, "AI 클라이언트 연결 중..."),
            (100,"준비 완료!"),
        ]

        # ── 페이드 인
        self._eff = QGraphicsOpacityEffect(self)
        self.setGraphicsEffect(self._eff)
        self._eff.setOpacity(0.0)
        self._fade_in = QPropertyAnimation(self._eff, b"opacity")
        self._fade_in.setDuration(350)
        self._fade_in.setStartValue(0.0)
        self._fade_in.setEndValue(1.0)
        self._fade_in.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._fade_in.start()

        # ── 진행 타이머 (60fps → ~3.2초)
        self._tick = 0
        self._timer = QTimer(self)
        self._timer.setInterval(16)   # ~60fps
        self._timer.timeout.connect(self._update)
        self._timer.start()

    def _update(self):
        self._tick += 1
        # 0~200 ticks = ~3.2초 (처음 20tick는 느리게, 이후 가속, 끝에서 잠깐 멈춤)
        total = 200
        t = self._tick / total
        # easeInOutCubic
        if t < 0.5:
            self._progress = int(4 * t * t * t * 100)
        else:
            f = -2 * t + 2
            self._progress = int((1 - f * f * f / 2) * 100)
        self._progress = min(100, self._progress)

        # 스텝 레이블 업데이트
        for threshold, label in reversed(self._steps):
            if self._progress >= threshold:
                self._step_idx = self._steps.index((threshold, label))
                break

        self.update()   # repaint

        if self._tick >= total + 20:  # 완료 후 0.3초 더 표시
            self._timer.stop()
            self._close_anim()

    def _close_anim(self):
        from PyQt6.QtCore import QPropertyAnimation, QEasingCurve
        anim = QPropertyAnimation(self._eff, b"opacity")
        anim.setDuration(280)
        anim.setStartValue(1.0)
        anim.setEndValue(0.0)
        anim.setEasingCurve(QEasingCurve.Type.InCubic)
        anim.finished.connect(self.close)
        anim.start()
        self._fade_out = anim   # GC 방지

    def paintEvent(self, event):
        from PyQt6.QtGui import QPainter, QPainterPath, QColor, QFont, QLinearGradient, QBrush, QPen
        from PyQt6.QtCore import QRectF, Qt

        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)

        W, H = self.width(), self.height()

        # ── 카드 배경 (둥글게)
        path = QPainterPath()
        path.addRoundedRect(QRectF(0, 0, W, H), 24, 24)
        p.setClipPath(path)

        # 배경색 (다크)
        p.fillPath(path, QColor("#0D0E14"))

        # 미묘한 테두리
        pen = QPen(QColor("#2A2D40"))
        pen.setWidthF(1.2)
        p.setPen(pen)
        p.drawRoundedRect(QRectF(0.6, 0.6, W-1.2, H-1.2), 23.5, 23.5)

        # 상단 파란 그라디언트 후광
        grd = QLinearGradient(0, 0, W, 0)
        grd.setColorAt(0.0, QColor("#3182F6") )
        grd.setColorAt(0.5, QColor("#6366F1"))
        grd.setColorAt(1.0, QColor("#3182F6"))
        accent_path = QPainterPath()
        accent_path.addRoundedRect(QRectF(0, 0, W, 3), 1.5, 1.5)
        p.fillPath(accent_path, QBrush(grd))

        # ── 아이콘
        if self._icon_px and not self._icon_px.isNull():
            ix = (W - 100) // 2
            p.drawPixmap(ix, 38, self._icon_px)
        else:
            # 폴백: 텍스트 "F"
            p.setPen(QColor("#3182F6"))
            f = QFont("Malgun Gothic", 48, QFont.Weight.Bold)
            p.setFont(f)
            p.drawText(QRectF(0, 30, W, 100), Qt.AlignmentFlag.AlignCenter, "F")

        # ── 앱 이름
        p.setPen(QColor("#F2F2F7"))
        fn = QFont("Malgun Gothic", 22, QFont.Weight.Bold)
        fn.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, -0.8)
        p.setFont(fn)
        p.drawText(QRectF(0, 150, W, 32), Qt.AlignmentFlag.AlignCenter, "Filelo")

        # 버전
        p.setPen(QColor("#636366"))
        fs = QFont("Malgun Gothic", 11)
        p.setFont(fs)
        p.drawText(QRectF(0, 181, W, 20), Qt.AlignmentFlag.AlignCenter, f"v{APP_VERSION}  ·  파일 · 문서 · AI 자동화")

        # ── 프로그레스 바 배경
        bar_x = 40
        bar_y = 220
        bar_w = W - 80
        bar_h = 4
        bar_r = 2.0

        bar_bg = QPainterPath()
        bar_bg.addRoundedRect(QRectF(bar_x, bar_y, bar_w, bar_h), bar_r, bar_r)
        p.fillPath(bar_bg, QColor("#1E2030"))

        # ── 프로그레스 바 채움 (그라디언트)
        fill_w = max(bar_r * 2, bar_w * self._progress / 100)
        if fill_w > 0:
            fill_grd = QLinearGradient(bar_x, 0, bar_x + fill_w, 0)
            fill_grd.setColorAt(0.0, QColor("#3182F6"))
            fill_grd.setColorAt(1.0, QColor("#6366F1"))
            bar_fill = QPainterPath()
            bar_fill.addRoundedRect(QRectF(bar_x, bar_y, fill_w, bar_h), bar_r, bar_r)
            p.fillPath(bar_fill, QBrush(fill_grd))

        # ── 스텝 레이블
        label = self._steps[self._step_idx][1]
        p.setPen(QColor("#8E8E93"))
        fl = QFont("Malgun Gothic", 10)
        p.setFont(fl)
        p.drawText(QRectF(0, 233, W, 20), Qt.AlignmentFlag.AlignCenter, label)

        p.end()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Filelo  v{APP_VERSION}")
        self.resize(1200,780); self.setMinimumSize(900,600)
        self.setStyleSheet(QSS); self._build()
        from PyQt6.QtGui import QShortcut, QKeySequence
        QShortcut(QKeySequence("Ctrl+K"), self).activated.connect(self._focus_search)
        QShortcut(QKeySequence("Ctrl+H"), self).activated.connect(lambda: self._show("home"))
        # 업데이트 체크 (24시간마다 1회, 백그라운드)
        if _should_check_update():
            QTimer.singleShot(3000, self._start_update_check)

    def _focus_search(self):
        if hasattr(self, "_search_input"):
            self._search_input.setFocus(); self._search_input.selectAll()

    def _start_update_check(self):
        def _on_result(latest_tag, release_url, has_update):
            if has_update:
                # UI 업데이트는 메인 스레드에서 — QTimer.singleShot 사용
                QTimer.singleShot(0, lambda: self._show_update_banner(latest_tag, release_url))
        _check_update_async(_on_result)

    def _show_update_banner(self, latest_tag: str, release_url: str):
        """상단에 업데이트 알림 배너 표시"""
        import webbrowser as _wb
        if hasattr(self, "_update_banner") and self._update_banner:
            return   # 이미 표시 중

        banner = QWidget(self)
        banner.setStyleSheet(
            f"QWidget{{background:{P['accent']};border:none;}}"
        )
        banner.setFixedHeight(38)
        banner.resize(self.width(), 38)
        banner.move(0, 0)

        bl = QHBoxLayout(banner)
        bl.setContentsMargins(16, 0, 12, 0)
        bl.setSpacing(12)

        ico = QLabel("🆕")
        ico.setStyleSheet("QLabel{background:transparent;font-size:14px;border:none;padding:0;}")

        msg = QLabel(
            f"새 버전 {latest_tag} 이 출시되었습니다!  "
            f"현재 버전: v{APP_VERSION}"
        )
        msg.setStyleSheet(
            "QLabel{background:transparent;color:#ffffff;"
            "font-size:12px;font-weight:600;border:none;padding:0;}"
        )

        dl_btn = QPushButton("지금 다운로드  →")
        dl_btn.setFixedHeight(26)
        dl_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        dl_btn.setStyleSheet(
            "QPushButton{background:#ffffff;color:#1C6FE8;border:none;"
            "border-radius:6px;font-size:11px;font-weight:700;padding:0 12px;}"
            "QPushButton:hover{background:#E8F0FF;}"
        )
        dl_btn.clicked.connect(lambda: _wb.open(release_url))

        close_btn = QPushButton("✕")
        close_btn.setFixedSize(26, 26)
        close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        close_btn.setStyleSheet(
            "QPushButton{background:transparent;color:#ffffff;"
            "border:none;font-size:13px;font-weight:700;}"
            "QPushButton:hover{background:#ffffff33;border-radius:4px;}"
        )
        def _close_banner():
            banner.hide()
            self._update_banner = None
            # 중앙 위젯을 원래 위치로
            if self.centralWidget():
                self.centralWidget().move(0, 0)
                self.centralWidget().resize(self.width(), self.height())
        close_btn.clicked.connect(_close_banner)

        bl.addWidget(ico)
        bl.addWidget(msg, 1)
        bl.addWidget(dl_btn)
        bl.addWidget(close_btn)

        # 중앙 위젯을 38px 아래로 밀기
        cw = self.centralWidget()
        if cw:
            cw.move(0, 38)
            cw.resize(self.width(), self.height() - 38)

        banner.show()
        banner.raise_()
        self._update_banner = banner

        # 페이드 인
        eff = QGraphicsOpacityEffect(banner)
        banner.setGraphicsEffect(eff)
        anim = QPropertyAnimation(eff, b"opacity")
        anim.setDuration(300)
        anim.setStartValue(0.0)
        anim.setEndValue(1.0)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        anim.start()
        self._update_anim = anim  # GC 방지

    def _build(self):
        root=QWidget(); self.setCentralWidget(root)
        rh=QHBoxLayout(root); rh.setContentsMargins(0,0,0,0); rh.setSpacing(0)

        # 사이드바
        sb=QWidget(); sb.setFixedWidth(220)
        sb.setStyleSheet(f"background:{P['side']};border-right:none;")
        sv=QVBoxLayout(sb); sv.setContentsMargins(0,0,0,0); sv.setSpacing(0)
        lw=QWidget(); lw.setStyleSheet(f"background:{P['side']};")
        lv=QVBoxLayout(lw); lv.setContentsMargins(20,24,20,4); lv.setSpacing(3)
        # 로고 행
        lr=QHBoxLayout(); lr.setSpacing(10); lr.setContentsMargins(0,0,0,0)
        from PyQt6.QtGui import QFont as _QFl
        # 로고 좌측 액센트 바
        accent_bar = QFrame()
        accent_bar.setFixedSize(3, 18)
        accent_bar.setStyleSheet(
            f"background:{P['accent']};border:none;border-radius:1px;"
        )
        l1 = QLabel("Filelo")
        l1.setStyleSheet(
            f"font-size:22px;font-weight:900;color:{P['text']};"
            f"background:transparent;letter-spacing:-1px;"
        )
        lr.addWidget(accent_bar); lr.addSpacing(10); lr.addWidget(l1); lr.addStretch()
        lv.addLayout(lr)
        l2 = QLabel("파일 · 문서 · AI 자동화")
        l2.setStyleSheet(
            f"font-size:10px;color:{P['sub2']};background:transparent;"
            f"letter-spacing:0.3px;margin-top:2px;"
        )
        lv.addWidget(l2); sv.addWidget(lw)
        # 구분선
        sep_w=QWidget(); sep_w.setStyleSheet(f"background:{P['side']};")
        sep_l=QVBoxLayout(sep_w); sep_l.setContentsMargins(16,8,16,4)
        sep_f=QFrame(); sep_f.setFrameShape(QFrame.Shape.HLine)
        sep_f.setStyleSheet(f"background:{P['border']};border:none;max-height:1px;")
        sep_l.addWidget(sep_f); sv.addWidget(sep_w)
        sc=QScrollArea(); sc.setWidgetResizable(True)
        sc.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        sc.setStyleSheet(
            f"QScrollArea{{background:{P['side']};border:none;}}"
            f"QScrollBar:vertical{{background:transparent;width:3px;border:none;}}"
            f"QScrollBar::handle:vertical{{background:{P['border']};border-radius:1px;}}"
            f"QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical{{height:0;}}"
        )
        mw=QWidget(); mw.setStyleSheet(f"background:{P['side']};")
        mv=QVBoxLayout(mw); mv.setContentsMargins(0,4,0,4); mv.setSpacing(0)
        CATS=[
            ("", [
                ("home", "홈"),
            ]),
            ("파일 관리", [
                ("translate", "파일명 번역"),
                ("folder",    "폴더 자동 정리"),
                ("rename",    "파일명 일괄 변경"),
                ("task_dir",  "과제 폴더 생성"),
            ]),
            ("문서 처리", [
                ("pdf",      "PDF 변환 & 추출"),
                ("pdfmerge", "PDF 병합 / 분리"),
                ("pdfpwd",   "PDF 비밀번호"),
                ("meta",     "메타데이터 삭제"),
                ("table2xl", "표 → 엑셀 변환"),
            ]),
            ("이미지", [
                ("image",     "이미지 일괄 처리"),
                ("imgpdf",    "이미지 → PDF"),
                ("imgext",    "이미지 일괄 추출"),
                ("watermark", "워터마크 삽입"),
                ("ocr",       "이미지 OCR"),
                ("rembg",     "배경 제거"),
            ]),
            ("AI 도구", [
                ("summary",  "AI 문서 요약"),
                ("draft",    "AI 문서 초안"),
                ("citation", "참고문헌 정리"),
            ]),
            ("데이터", [
                ("excel", "엑셀 자동화"),
            ]),
            ("학습 관리", [
                ("tracker", "마감 트래커"),
            ]),
            ("설정", [
                ("settings", "API 키 설정"),
                ("help",     "도움말"),
            ]),
        ]
        self._nb={}
        for cat, items in CATS:
            if cat:
                cl = QLabel(cat.upper())
                cl.setStyleSheet(
                    f"color:{P['sub2']};font-size:9px;font-weight:700;"
                    f"padding:16px 16px 4px 16px;background:transparent;"
                    f"letter-spacing:1.2px;"
                )
                mv.addWidget(cl)
            for key, lbl in items:
                nb = NavBtn(None, lbl, key)
                nb.clicked.connect(lambda _, k=key: self._show(k))
                self._nb[key] = nb
                mv.addWidget(nb)
        mv.addStretch()

        # Discord 문의 버튼
        _dc_sep = QFrame(); _dc_sep.setFixedHeight(1)
        _dc_sep.setStyleSheet(f"background:{P['sep']};border:none;")
        sv.addWidget(_dc_sep)
        _dc_btn = QPushButton("  💬  문의하기  —  Discord")
        _dc_btn.setFixedHeight(44)
        _dc_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        _dc_btn.setStyleSheet(
            "QPushButton{background:transparent;border:none;text-align:left;"
            "padding:0 16px;font-size:12px;font-weight:500;color:#5865F2;}"
            "QPushButton:hover{background:#5865F211;color:#7289DA;}"
        )
        import webbrowser as _wb_dc
        _dc_btn.clicked.connect(lambda: _wb_dc.open("https://discord.gg/7agPwy9KRb"))
        sv.addWidget(_dc_btn)

        sc.setWidget(mw); sv.addWidget(sc,1); rh.addWidget(sb)

        # 오른쪽
        right=QWidget(); right.setStyleSheet(f"background:{P['bg']};")
        rv=QVBoxLayout(right); rv.setContentsMargins(0,0,0,0); rv.setSpacing(0)
        tb=QWidget(); tb.setFixedHeight(52); tb.setStyleSheet(f"background:{P['side']};border-bottom:1px solid {P['sep']};")
        th=QHBoxLayout(tb); th.setContentsMargins(24,0,16,0); th.setSpacing(8)
        self._pl = QLabel("홈 대시보드")
        self._pl.setStyleSheet(
            f"font-size:14px;font-weight:700;color:{P['text']};"
            f"background:transparent;letter-spacing:-0.3px;"
        )
        th.addWidget(self._pl)
        th.addStretch()

        # ── 검색창 (topbar 중앙~우측)
        self._search_input = QLineEdit()
        self._search_input.setPlaceholderText("기능 검색  (PDF, 번역, OCR ...)")
        self._search_input.setFixedWidth(280)
        self._search_input.setFixedHeight(32)
        self._search_input.setStyleSheet(
            f"QLineEdit{{background:{P['input']};border:1px solid {P['border2']};"
            f"border-radius:8px;padding:0 14px;font-size:13px;color:{P['text']};}}"
            f"QLineEdit:focus{{border-color:{P['accent']};}}"
        )
        th.addWidget(self._search_input)
        th.addSpacing(8)
        # ── API 키 상태 통합 카드 (한 줄 — 클릭 시 설정으로)
        self._api_status_btn = QPushButton()
        self._api_status_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._api_status_btn.setFixedHeight(34)
        self._api_status_btn.setMinimumWidth(170)  # 축소 방지
        self._api_status_btn.setStyleSheet(
            f"QPushButton{{background:{P['card']};border:1px solid {P['border2']};"
            f"border-radius:9px;padding:0;}}"
            f"QPushButton:hover{{background:{P['glass']};}}"
        )
        self._api_status_btn.clicked.connect(lambda: self._show("settings"))

        api_hl = QHBoxLayout(self._api_status_btn)
        api_hl.setContentsMargins(10, 0, 12, 0)
        api_hl.setSpacing(5)

        # "API 연결" 레이블 (짧게)
        api_title = QLabel("API")
        api_title.setStyleSheet(
            f"font-size:11px;font-weight:600;color:{P['sub']};"
            f"background:transparent;border:none;"
        )
        api_title.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        api_hl.addWidget(api_title)

        # 구분선
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.VLine)
        sep.setFixedSize(1, 14)
        sep.setStyleSheet(f"background:{P['border2']};border:none;")
        sep.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        api_hl.addWidget(sep)

        # Gemini 상태
        self._g_dot = QLabel("●")
        self._g_dot.setStyleSheet(
            f"font-size:10px;color:{'#05C072' if GEMINI_KEY else '#FF3B30'};"
            f"background:transparent;border:none;"
        )
        self._g_dot.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self._g_lbl = QLabel("Gemini")
        self._g_lbl.setStyleSheet(
            f"font-size:12px;font-weight:600;"
            f"color:{P['text'] if GEMINI_KEY else P['sub2']};"
            f"background:transparent;border:none;"
        )
        self._g_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        api_hl.addWidget(self._g_dot)
        api_hl.addWidget(self._g_lbl)
        api_hl.addSpacing(6)

        # 구분 슬래시
        slash = QLabel("/")
        slash.setStyleSheet(f"color:{P['sub2']};background:transparent;border:none;font-size:11px;")
        slash.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        api_hl.addWidget(slash)
        api_hl.addSpacing(6)

        # DeepL 상태
        self._d_dot = QLabel("●")
        self._d_dot.setStyleSheet(
            f"font-size:10px;color:{'#05C072' if DEEPL_KEY else '#FF3B30'};"
            f"background:transparent;border:none;"
        )
        self._d_dot.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self._d_lbl = QLabel("DeepL")
        self._d_lbl.setStyleSheet(
            f"font-size:12px;font-weight:600;"
            f"color:{P['text'] if DEEPL_KEY else P['sub2']};"
            f"background:transparent;border:none;"
        )
        self._d_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        api_hl.addWidget(self._d_dot)
        api_hl.addWidget(self._d_lbl)

        th.addWidget(self._api_status_btn)
        th.addSpacing(4)

        # 더미 변수 (refresh_badges 호환)
        self._deepl_card = self._api_status_btn
        self._gemini_card = self._api_status_btn
        self._d_status = self._d_lbl
        self._g_status = self._g_lbl
        self._d_use = self._d_dot
        self._g_use = self._g_dot
        self._d_ind = self._d_dot
        self._g_ind = self._g_dot
        rv.addWidget(tb)
        self._st=QStackedWidget(); rv.addWidget(self._st,1); rh.addWidget(right,1)

        # 페이지 등록
        pg=HomePage(app=self)
        pg.nav_req.connect(self._show)
        meta_pg=_simple(self,"","메타데이터 완전 삭제","문서의 작성자·수정 이력 등 디지털 지문을 제거합니다",[
    "문서 파일에 숨어있는 작성자 정보, 수정 이력, PC 이름 등 개인정보를 완전히 지웁니다.",
    "① 외부에 제출하는 파일에서 내 이름이나 회사명 등이 노출되는 것을 방지합니다.",
    "② PDF: 작성자·제목·키워드·소프트웨어 정보 등이 제거됩니다.",
    "③ DOCX: 작성자·마지막 수정자·제목·주석 정보 등이 제거됩니다.",
    "④ 원본은 그대로 유지되고 _clean이 붙은 새 파일이 같은 폴더에 저장됩니다.",
],"문서 (*.pdf *.docx)",_meta_exec," 메타데이터 삭제")
        t2xl_pg=_simple(self,"","표 → 엑셀 변환","PDF 또는 DOCX 안의 표를 엑셀 시트로 자동 변환합니다",[
    "PDF나 Word 문서 안에 있는 표를 자동으로 인식해 엑셀 파일로 변환합니다.",
    "① 문서 안에 표가 여러 개 있으면 각각 별도의 시트로 분리 저장됩니다.",
    "② DOCX: 문서 내 모든 표를 순서대로 추출합니다.",
    "③ PDF: 페이지별로 표를 자동 감지해 추출합니다. (스캔 PDF는 지원 안 됨)",
    "④ 결과 파일은 원본 파일명에 _표변환.xlsx가 붙어 저장됩니다.",
],"문서 (*.pdf *.docx)",_t2xl_exec," 변환 시작")
        ie_pg=_simple(self,"️","이미지 일괄 추출","PDF·DOCX 문서들에서 이미지만 한번에 추출합니다",[
    "PDF나 Word 문서 안에 삽입된 이미지를 한 번에 모두 꺼냅니다.",
    "① 보고서나 교재에서 그림만 따로 모아야 할 때 유용합니다.",
    "② 추출된 이미지는 원본 파일명_이미지들 폴더 안에 저장됩니다.",
    "③ 파일명 규칙: PDF는 p(페이지번호)_(이미지번호).확장자 형식입니다.",
    "④ 스캔된 PDF의 경우 PDF 자체가 이미지이므로 내부 이미지는 추출되지 않습니다.",
],"문서 (*.pdf *.docx)",_imgext_exec,"️ 추출 시작")
        self._pages={
            "home":pg, "translate":TranslatePage(self), "folder":FolderPage(self),
            "rename":RenamePage(self), "task_dir":TaskDirPage(self),
            "pdf":PdfPage(self), "pdfmerge":PdfMergePage(self), "pdfpwd":PdfPwdPage(self),
            "meta":meta_pg, "table2xl":t2xl_pg,
            "image":ImagePage(self), "imgpdf":ImgPdfPage(self), "imgext":ie_pg,
            "watermark":WatermarkPage(self), "ocr":OcrPage(self),
            "summary":SummaryPage(self), "draft":DraftPage(self),
            "rembg":RembgPage(self), "citation":CitationPage(self),
            "excel":ExcelPage(self), "tracker":TrackerPage(self), "settings":SettingsPage(self), "help":HelpPage(self),
        }
        for p in self._pages.values(): self._st.addWidget(p)
        self._show("home")

        # ── SearchOverlay 초기화
        self._search_overlay = SearchOverlay(
            parent=self,
            search_input=self._search_input,
            nav_cb=self._show,
            refresh_cb=lambda: (
                self._pages["home"]._refresh_tags()
                if hasattr(self._pages.get("home"), "_refresh_tags") else None
            ),
        )
        self._search_input.textChanged.connect(self._on_search_text)
        self._search_input.returnPressed.connect(lambda: (
            self._search_overlay.select_current()
            if hasattr(self, "_search_overlay") and self._search_overlay.isVisible() else None
        ))

        # 홈 페이지 내 검색창도 연결
        home_pg = self._pages.get("home")
        if home_pg and hasattr(home_pg, "get_home_search"):
            home_search = home_pg.get_home_search()
            # 홈 검색창용 별도 오버레이 (MainWindow 기준 좌표)
            self._home_overlay = SearchOverlay(
                parent=self,
                search_input=home_search,
                nav_cb=self._show,
                refresh_cb=home_pg._refresh_tags,
            )
            home_search.textChanged.connect(
                lambda t: self._home_overlay.update_results(t)
            )
            home_search.returnPressed.connect(self._home_overlay.select_current)

            from PyQt6.QtCore import QObject as _QObj2, QEvent as _QEv2

            # 홈 검색창 키보드 필터
            class _HKF(_QObj2):
                def __init__(self, overlay, parent=None):
                    super().__init__(parent)
                    self._ov = overlay
                def eventFilter(self, obj, event):
                    if event.type() == _QEv2.Type.KeyPress:
                        from PyQt6.QtCore import Qt as _Qt
                        k = event.key()
                        if k == _Qt.Key.Key_Down:   self._ov.move_cursor(1);  return True
                        if k == _Qt.Key.Key_Up:     self._ov.move_cursor(-1); return True
                        if k == _Qt.Key.Key_Escape: self._ov.hide(); return True
                    return False
            self._hkf = _HKF(self._home_overlay)
            home_search.installEventFilter(self._hkf)

            # 포커스 시 오버레이 표시
            class _HFF(_QObj2):
                def __init__(self, overlay, parent=None):
                    super().__init__(parent)
                    self._ov = overlay
                def eventFilter(self, obj, event):
                    if event.type() == _QEv2.Type.FocusIn:
                        if not self._ov._search.text():
                            self._ov.update_results("")
                    if event.type() == _QEv2.Type.FocusOut:
                        from PyQt6.QtCore import QTimer as _QT
                        _QT.singleShot(150, self._ov.hide)
                    return False
            self._hff = _HFF(self._home_overlay)
            home_search.installEventFilter(self._hff)

        # 키보드 이벤트 필터 (↑↓ Esc)
        from PyQt6.QtCore import QObject, QEvent
        class _KeyFilter(QObject):
            def __init__(self, overlay, parent=None):
                super().__init__(parent)
                self._ov = overlay
            def eventFilter(self, obj, event):
                if event.type() == QEvent.Type.KeyPress:
                    from PyQt6.QtCore import Qt as _Qt
                    k = event.key()
                    if k == _Qt.Key.Key_Down:
                        self._ov.move_cursor(1); return True
                    if k == _Qt.Key.Key_Up:
                        self._ov.move_cursor(-1); return True
                    if k == _Qt.Key.Key_Escape:
                        self._ov.hide()
                        self._ov._search.clearFocus(); return True
                return False
        self._kf = _KeyFilter(self._search_overlay)
        self._search_input.installEventFilter(self._kf)

        # 검색창 포커스 시 최근 검색어 표시
        from PyQt6.QtCore import QObject as _QO, QEvent as _QE
        class _FocusFilter(_QO):
            def __init__(self, overlay, parent=None):
                super().__init__(parent)
                self._ov = overlay
            def eventFilter(self, obj, event):
                if event.type() == _QE.Type.FocusIn:
                    if not self._ov._search.text():
                        self._ov.update_results("")
                if event.type() == _QE.Type.FocusOut:
                    from PyQt6.QtCore import QTimer as _QT
                    _QT.singleShot(150, self._ov.hide)
                return False
        self._ff = _FocusFilter(self._search_overlay)
        self._search_input.installEventFilter(self._ff)

    def _on_search_text(self, text):
        self._search_overlay.update_results(text)

    def _show(self,key):
        # 페이지 이동 시 오버레이 닫기
        if hasattr(self, "_search_overlay"):
            self._search_overlay.hide()
        if hasattr(self, "_home_overlay"):
            self._home_overlay.hide()
        if hasattr(self, "_search_input") and self._search_input.text():
            self._search_input.clear()
        # 홈에서는 상단 검색창 숨김
        if hasattr(self, "_search_input"):
            self._search_input.setVisible(key != "home")
        if key not in self._pages: return
        # 사용 기록 (홈/설정 제외)
        if key not in ("home", "settings"):
            record_usage(key)
            home_pg = self._pages.get("home")
            if home_pg and hasattr(home_pg, "_refresh_tags"):
                home_pg._refresh_tags()
             # 홈 태그 갱신
            home_pg = self._pages.get("home")
            if home_pg and hasattr(home_pg, "_refresh_tags"):
                home_pg._refresh_tags()
        for k,nb in self._nb.items(): nb.set_active(k==key)
        # 페이드 + 슬라이드Y 전환
        new_page = self._pages[key]
        self._st.setCurrentWidget(new_page)

        # ── opacity effect
        if not hasattr(new_page, "_fade_eff"):
            new_page._fade_eff = QGraphicsOpacityEffect(new_page)
            new_page.setGraphicsEffect(new_page._fade_eff)
        eff = new_page._fade_eff; eff.setOpacity(0.0)

        # ── opacity 애니메이션
        if not hasattr(self,"_tr_op") or self._tr_op is None:
            self._tr_op = QPropertyAnimation(eff, b"opacity")
            self._tr_op.setDuration(220)
            self._tr_op.setEasingCurve(QEasingCurve.Type.OutCubic)
        else:
            self._tr_op.stop(); self._tr_op.setTargetObject(eff)
        self._tr_op.setStartValue(0.0); self._tr_op.setEndValue(1.0)
        self._tr_op.start()

        # ── 슬라이드Y 애니메이션 (12px 아래서 올라오는 느낌)
        orig_pos = new_page.pos()
        new_page.move(orig_pos.x(), orig_pos.y() + 12)
        if not hasattr(self,"_tr_pos"):
            self._tr_pos = QPropertyAnimation(new_page, b"pos")
            self._tr_pos.setDuration(260)
            self._tr_pos.setEasingCurve(QEasingCurve.Type.OutCubic)
        else:
            self._tr_pos.stop(); self._tr_pos.setTargetObject(new_page)
        self._tr_pos.setStartValue(new_page.pos())
        self._tr_pos.setEndValue(orig_pos)
        self._tr_pos.start()
        titles = {
            "home":     "홈",
            "translate":"파일명 번역",
            "folder":   "폴더 자동 정리",
            "rename":   "파일명 일괄 변경",
            "task_dir": "과제 폴더 생성",
            "pdf":      "PDF 변환 & 추출",
            "pdfmerge": "PDF 병합 / 분리",
            "pdfpwd":   "PDF 비밀번호",
            "meta":     "메타데이터 삭제",
            "table2xl": "표 → 엑셀 변환",
            "image":    "이미지 일괄 처리",
            "imgpdf":   "이미지 → PDF",
            "imgext":   "이미지 일괄 추출",
            "watermark":"워터마크 삽입",
            "ocr":      "이미지 OCR",
            "summary":  "AI 문서 요약",
            "draft":    "AI 문서 초안",
            "rembg":    "배경 제거",
            "citation": "참고문헌 정리",
            "excel":    "엑셀 자동화",
            "tracker":  "마감 트래커",
            "settings": "API 키 설정",
            "help":     "도움말 — API 키 발급 가이드",
        }
        self._pl.setText(titles.get(key,key))
        self._refresh_badges()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        if hasattr(self, "_search_overlay") and self._search_overlay.isVisible():
            self._search_overlay.reposition()
        if hasattr(self, "_home_overlay") and self._home_overlay.isVisible():
            self._home_overlay.reposition()
        # 업데이트 배너 너비 동기화
        if getattr(self, "_update_banner", None):
            self._update_banner.resize(self.width(), 38)
            cw = self.centralWidget()
            if cw:
                cw.move(0, 38)
                cw.resize(self.width(), self.height() - 38)

    def _refresh_badges(self):
        def _dot(ok): return f"font-size:10px;color:{'#05C072' if ok else '#FF3B30'};background:transparent;border:none;"
        def _lbl(ok): return f"font-size:12px;font-weight:600;color:{P['text'] if ok else P['sub2']};background:transparent;border:none;"
        self._g_dot.setStyleSheet(_dot(bool(GEMINI_KEY)))
        self._g_lbl.setStyleSheet(_lbl(bool(GEMINI_KEY)))
        self._d_dot.setStyleSheet(_dot(bool(DEEPL_KEY)))
        self._d_lbl.setStyleSheet(_lbl(bool(DEEPL_KEY)))

    def toast(self,msg,kind="ok"):
        cols={"ok":(P["success"],P["glass"]),"err":(P["danger"],P["glass"]),"warn":(P["warning"],P["glass"]),"info":(P["accent"],P["glass"])}
        fg,bg=cols.get(kind,(P["success"],"#0D2018"))
        icons={"ok":"","err":"","warn":"️","info":"ℹ️"}
        from PyQt6.QtGui import QFont as _QFt
        t=QLabel(f" {icons.get(kind,'')} {msg} ",self)
        t.setStyleSheet(
            f"QLabel{{"
            f"background:{P['glass']};"
            f"color:{fg};"
            f"border:1px solid {fg}44;"
            f"border-radius:12px;"
            f"font-size:13px;font-weight:600;"
            f"padding:8px 4px;"
            f"letter-spacing:-0.2px;"
            f"}}"
        )
        t.adjustSize()
        cw = self.centralWidget()
        end_x = cw.width() - t.width() - 20
        end_y = cw.height() - t.height() - 20
        # 아래에서 슬라이드 인
        t.move(end_x, end_y + 28)
        t.show(); t.raise_()

        # 슬라이드 인 애니메이션
        anim_in = QPropertyAnimation(t, b"pos")
        anim_in.setDuration(280)
        anim_in.setStartValue(QPoint(end_x, end_y + 28))
        anim_in.setEndValue(QPoint(end_x, end_y))
        anim_in.setEasingCurve(QEasingCurve.Type.OutBack)
        anim_in.start()

        # 페이드 아웃 후 삭제
        def _fade_out():
            effect = QGraphicsOpacityEffect(t)
            t.setGraphicsEffect(effect)
            anim_out = QPropertyAnimation(effect, b"opacity")
            anim_out.setDuration(300)
            anim_out.setStartValue(1.0)
            anim_out.setEndValue(0.0)
            anim_out.setEasingCurve(QEasingCurve.Type.InCubic)
            anim_out.finished.connect(t.deleteLater)
            anim_out.start()
            t._fade = anim_out  # GC 방지

        t._anim_in = anim_in  # GC 방지
        QTimer.singleShot(2500, _fade_out)

if __name__=="__main__":
    app=QApplication(sys.argv)
    app.setStyle("Fusion")
    icon = _get_icon()
    if icon:
        app.setWindowIcon(icon)

    # ── 1단계: 약관 동의 (최초 1회)
    if not _consent_given():
        dlg = ConsentDialog()
        if icon:
            dlg.setWindowIcon(icon)
        result = dlg.exec()
        if result != QDialog.DialogCode.Accepted:
            sys.exit(0)   # 동의 거부 → 종료

    # ── 2단계: 스플래시 스크린
    splash = None
    try:
        splash = SplashScreen(_ICON_B64)
        splash.show()
        app.processEvents()
    except Exception:
        splash = None

    # ── 3단계: 메인 윈도우 (스플래시 종료 후)
    def _launch():
        win = MainWindow()
        if icon:
            win.setWindowIcon(icon)
        win.show()
        if splash and splash.isVisible():
            splash.close()
        app._main_win = win   # GC 방지

    QTimer.singleShot(3500 if splash else 0, _launch)
    sys.exit(app.exec())
