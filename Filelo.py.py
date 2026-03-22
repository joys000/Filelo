import os, re, time, json, shutil, threading, datetime, base64, uuid
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path

try:    import deepl;                                HAS_DEEPL  = True
except Exception:                                    HAS_DEEPL  = False
try:
    from google import genai; from google.genai import types; HAS_GEMINI = True
except Exception:                                    HAS_GEMINI = False
try:    from docx import Document;                   HAS_DOCX   = True
except Exception:                                    HAS_DOCX   = False
try:    import fitz;                                 HAS_FITZ   = True
except Exception:                                    HAS_FITZ   = False
try:    from PIL import Image, ImageDraw, ImageFont; HAS_PIL    = True
except Exception:                                    HAS_PIL    = False
try:    import openpyxl;                             HAS_XL     = True
except Exception:                                    HAS_XL     = False
try:    from rembg import remove as rembg_remove;   HAS_REMBG  = True
except Exception:                                    HAS_REMBG  = False

TARGET_MODEL = "gemini-2.0-flash"
CONFIG_FILE  = os.path.join(os.path.expanduser("~"), ".filelo_secure.json")
SALT_FILE    = os.path.join(os.path.expanduser("~"), ".filelo_salt")

try:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    HAS_CRYPTO = True
except Exception:
    HAS_CRYPTO = False

def _get_machine_id():
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\Microsoft\\Cryptography")
        val, _ = winreg.QueryValueEx(key, "MachineGuid")
        return val
    except Exception: pass
    try:
        with open("/etc/machine-id","r") as f: return f.read().strip()
    except Exception: pass
    fb = os.path.join(os.path.expanduser("~"), ".filelo_id")
    if os.path.exists(fb):
        with open(fb,"r") as f: return f.read().strip()
    nid = str(uuid.uuid4())
    with open(fb,"w") as f: f.write(nid)
    return nid

def _derive_key(salt):
    if not HAS_CRYPTO: return b""
    pw = ("Filelo:" + _get_machine_id() + ":v1").encode()
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=480_000)
    return kdf.derive(pw)

def _get_or_create_salt():
    if os.path.exists(SALT_FILE):
        with open(SALT_FILE,"rb") as f: return f.read()
    salt = os.urandom(32)
    with open(SALT_FILE,"wb") as f: f.write(salt)
    return salt

def _encrypt(plaintext):
    if not HAS_CRYPTO or not plaintext: return ""
    salt = _get_or_create_salt(); key = _derive_key(salt)
    nonce = os.urandom(12)
    ct = AESGCM(key).encrypt(nonce, plaintext.encode(), None)
    return base64.b64encode(nonce + ct).decode()

def _decrypt(encoded):
    if not HAS_CRYPTO or not encoded: return ""
    try:
        salt = _get_or_create_salt(); key = _derive_key(salt)
        raw = base64.b64decode(encoded.encode())
        return AESGCM(key).decrypt(raw[:12], raw[12:], None).decode()
    except Exception: return ""

def load_config():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE,"r",encoding="utf-8") as f: raw = json.load(f)
            return {k: _decrypt(v) for k,v in raw.items()}
    except Exception: pass
    return {}

def save_config(cfg):
    if not HAS_CRYPTO: return
    encoded = {k: _encrypt(v) for k,v in cfg.items() if v}
    with open(CONFIG_FILE,"w",encoding="utf-8") as f: json.dump(encoded,f)
    try: os.chmod(CONFIG_FILE,0o600); os.chmod(SALT_FILE,0o600)
    except Exception: pass

DATA_FILE = os.path.join(os.path.expanduser("~"), ".filelo_tasks.json")
def load_tasks():
    try:
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE,"r",encoding="utf-8") as f: return json.load(f)
    except Exception: pass
    return []
def save_tasks(tasks):
    with open(DATA_FILE,"w",encoding="utf-8") as f:
        json.dump(tasks,f,ensure_ascii=False,indent=2)

_cfg           = load_config()
DEEPL_API_KEY  = _cfg.get("deepl_key","")
GEMINI_API_KEY = _cfg.get("gemini_key","")
ai_client      = None
def _init_gemini():
    global ai_client
    if HAS_GEMINI and GEMINI_API_KEY:
        try: ai_client = genai.Client(api_key=GEMINI_API_KEY); return True
        except Exception: pass
    return False
_init_gemini()

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")
BG_MAIN="#13151F"; BG_SIDE="#181C2A"; BG_CARD="#1C2035"; BG_DARK="#0E1018"
BG_TIP="#161D33"; ACCENT="#5B9CF6"; ACCENT2="#7B6CF6"
SUCCESS="#34D9A0"; WARNING="#F9C74F"; DANGER="#F47171"
TEXT_PRI="#E8EAF2"; TEXT_SUB="#7A84A0"; BORDER="#272D45"

_FONT_CACHE = {}
def F(s,w="normal"):
    k=(s,w,"d")
    if k not in _FONT_CACHE: _FONT_CACHE[k]=ctk.CTkFont(size=s,weight=w)
    return _FONT_CACHE[k]
def FBOLD(s):
    k=(s,"bold","d")
    if k not in _FONT_CACHE: _FONT_CACHE[k]=ctk.CTkFont(size=s,weight="bold")
    return _FONT_CACHE[k]
def FMONO(s):
    k=(s,"n","m")
    if k not in _FONT_CACHE: _FONT_CACHE[k]=ctk.CTkFont(family="Consolas",size=s)
    return _FONT_CACHE[k]

def _press_color(c):
    M={"#5B9CF6":"#2A4A9A","#34D9A0":"#1E7A60","#F47171":"#9A4A2A",
       "#F9C74F":"#9A7A20","#7B6CF6":"#4A3A9A"}
    return M.get(c,"#2A3A7A")

def make_animated_btn(parent, text, command, fg_color=None, **kw):
    color = fg_color or ACCENT
    press = _press_color(color)
    def _cmd():
        def _fade(step):
            try:
                btn.configure(fg_color=press if step<=3 else color)
                if step<6: btn.after(30,lambda:_fade(step+1))
            except Exception: pass
        _fade(0)
        if command: btn.after(90,command)
    btn = ctk.CTkButton(parent,text=text,fg_color=color,command=_cmd,**kw)
    return btn

class GradientCanvas(tk.Canvas):
    def __init__(self,parent,**kw):
        super().__init__(parent,highlightthickness=0,bd=0,**kw)
        self._cache_size=(0,0); self._photo=None; self._after_id=None
        self.bind("<Configure>",self._on_configure)
    def _on_configure(self,e=None):
        if self._after_id: self.after_cancel(self._after_id)
        self._after_id=self.after(100,self._draw)
    def _draw(self):
        self._after_id=None; w=self.winfo_width(); h=self.winfo_height()
        if w<2 or h<2: return
        if (w,h)==self._cache_size and self._photo: return
        if HAS_PIL: self._draw_pil(w,h)
        else: self._draw_fallback(w,h)
        self._cache_size=(w,h)
    def _draw_pil(self,w,h):
        img=Image.new("RGB",(w,h)); px=img.load()
        for y in range(h):
            t=y/max(h-1,1)
            r=int(0x1A+(0x13-0x1A)*t); g=int(0x20+(0x15-0x20)*t); b=int(0x40+(0x1F-0x40)*t)
            for x in range(w): px[x,y]=(r,g,b)
        rad=min(w,h)//3
        for y in range(min(rad*2,h)):
            for x in range(min(rad*2,w)):
                d=(x*x+y*y)**0.5
                if d<rad:
                    s=(1-d/rad)*0.35; pr,pg,pb=px[x,y]
                    px[x,y]=(min(255,int(pr+0x28*s)),min(255,int(pg+0x10*s)),min(255,int(pb+0x30*s)))
        from PIL import ImageTk
        self._photo=ImageTk.PhotoImage(img); self.delete("all")
        self.create_image(0,0,anchor="nw",image=self._photo)
    def _draw_fallback(self,w,h):
        self.delete("all"); self.configure(bg="#181C2A")

class MinjuToolkitApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Filelo  v5.0")
        self.geometry("1180x780"); self.minsize(960,640)
        self.configure(fg_color=BG_MAIN)
        self.tasks=load_tasks()
        self._cat_open={}; self.nav_btns={}; self.cat_frames={}
        self._active_tab=None; self.tabs={}
        self._build_layout()
        self._show_tab("home")

    CATEGORIES = [
        ("📁  파일 관리","file",[
            ("translate","🚀","파일명 번역"),
            ("folder","📁","폴더 자동 정리"),
            ("rename","🏷️","파일명 규칙 변경"),
            ("task_dir","📂","과제 폴더 생성"),
        ]),
        ("📄  문서 처리","doc",[
            ("pdf","📄","PDF 변환 & 추출"),
            ("pdfmerge","🔗","PDF 합치기/쪼개기"),
            ("pdfpwd","🔒","PDF 비밀번호 설정"),
            ("meta","🧹","메타데이터 삭제"),
            ("table2xl","📊","표 → 엑셀 변환"),
        ]),
        ("🖼️  이미지","img",[
            ("image","🖼️","이미지 일괄 처리"),
            ("imgpdf","📎","이미지 → PDF"),
            ("imgext","🖼️","이미지 일괄 추출"),
            ("watermark","💧","워터마크 삽입"),
            ("ocr","📸","이미지 OCR"),
            ("rembg","✂️","배경 제거"),
        ]),
        ("🤖  AI 도구","ai",[
            ("summary","🤖","AI 문서 요약"),
            ("draft","✏️","AI 문서 초안"),
            ("citation","📚","참고문헌 자동 정리"),
        ]),
        ("📊  데이터","data",[
            ("excel","📊","엑셀 자동화"),
        ]),
        ("📅  학습 관리","study",[
            ("tracker","📅","과제 마감 트래커"),
        ]),
        ("⚙️  설정","settings",[
            ("settings","⚙️","API 키 설정"),
        ]),
    ]

    def _build_layout(self):
        self.sidebar=ctk.CTkFrame(self,width=230,fg_color=BG_SIDE,corner_radius=0)
        self.sidebar.pack(side="left",fill="y"); self.sidebar.pack_propagate(False)
        GradientCanvas(self.sidebar,bg=BG_SIDE).place(x=0,y=0,relwidth=1,relheight=1)

        top_f=ctk.CTkFrame(self.sidebar,fg_color="transparent"); top_f.pack(fill="x")
        ctk.CTkLabel(top_f,text="Filelo",font=FBOLD(21),text_color=ACCENT).pack(pady=(28,2),padx=20,anchor="w")
        ctk.CTkLabel(top_f,text="파일·문서·AI 자동화 툴킷",font=F(11),text_color=TEXT_SUB).pack(padx=20,anchor="w")
        ctk.CTkFrame(top_f,height=1,fg_color=BORDER).pack(fill="x",padx=20,pady=(12,0))
        home_btn=make_animated_btn(top_f,text="  🏠  홈 대시보드",
            command=lambda:self._show_tab("home"),fg_color="transparent",
            hover_color="#232840",text_color=TEXT_PRI,anchor="w",height=40,
            corner_radius=8,font=FBOLD(13))
        home_btn.pack(fill="x",padx=10,pady=(6,0))
        self.nav_btns["home"]=home_btn
        ctk.CTkFrame(top_f,height=1,fg_color=BORDER).pack(fill="x",padx=20,pady=(6,0))

        nav_scroll=ctk.CTkScrollableFrame(self.sidebar,fg_color="transparent",
            scrollbar_button_color=BORDER,scrollbar_button_hover_color=ACCENT)
        nav_scroll.pack(fill="both",expand=True)

        for _,cat_key,items in self.CATEGORIES:
            self._cat_open[cat_key]=True
        for cat_label,cat_key,items in self.CATEGORIES:
            short=cat_label.split("  ",1)[-1] if "  " in cat_label else cat_label
            cat_btn=ctk.CTkButton(nav_scroll,text="  ▾  "+short,anchor="w",
                font=ctk.CTkFont(size=11,weight="bold"),height=30,fg_color="#1A2040",
                hover_color="#222848",text_color=ACCENT,corner_radius=6,
                command=lambda k=cat_key:self._toggle_cat(k))
            cat_btn.pack(fill="x",padx=8,pady=(10,0))
            item_frame=ctk.CTkFrame(nav_scroll,fg_color="transparent")
            item_frame.pack(fill="x",padx=0,pady=0)
            self.cat_frames[cat_key]={"btn":cat_btn,"item_frame":item_frame,"items":items}
            for key,icon,label in items:
                btn=make_animated_btn(item_frame,text="    "+icon+"  "+label,
                    command=lambda k=key:self._show_tab(k),
                    fg_color="transparent",hover_color="#232840",
                    text_color="#C8D0E8",anchor="w",height=36,corner_radius=8,font=F(12))
                btn.pack(fill="x",padx=6,pady=1)
                self.nav_btns[key]=btn

        self.content=ctk.CTkFrame(self,fg_color=BG_MAIN,corner_radius=0)
        self.content.pack(side="left",fill="both",expand=True)

        builders={
            "home":self._build_home,"translate":self._build_translate,
            "folder":self._build_folder,"rename":self._build_rename,
            "task_dir":self._build_task_dir,"pdf":self._build_pdf,
            "pdfmerge":self._build_pdfmerge,"pdfpwd":self._build_pdfpwd,
            "meta":self._build_meta,"table2xl":self._build_table2xl,
            "image":self._build_image,"imgpdf":self._build_imgpdf,
            "imgext":self._build_imgext,"watermark":self._build_watermark,
            "ocr":self._build_ocr,"summary":self._build_summary,
            "draft":self._build_draft,"tracker":self._build_tracker,
            "rembg":self._build_rembg,"citation":self._build_citation,
            "excel":self._build_excel,"settings":self._build_settings,
        }
        self._tab_builders=builders; self._tab_built=set()
        self._no_scroll_tabs={"tracker","home"}
        self._build_tab("home")

    def _toggle_cat(self,cat_key):
        is_open=self._cat_open[cat_key]
        item_frame=self.cat_frames[cat_key]["item_frame"]
        btn=self.cat_frames[cat_key]["btn"]
        cat_label=[c[0] for c in self.CATEGORIES if c[1]==cat_key][0]
        short=cat_label.split("  ",1)[-1] if "  " in cat_label else cat_label
        if is_open:
            for child in item_frame.winfo_children(): child.pack_forget()
            btn.configure(text="  ▸  "+short,text_color=TEXT_SUB,fg_color="transparent")
            self._cat_open[cat_key]=False
        else:
            for key,icon,label in self.cat_frames[cat_key]["items"]:
                b=self.nav_btns.get(key)
                if b: b.pack(fill="x",padx=6,pady=1)
            btn.configure(text="  ▾  "+short,text_color=ACCENT,fg_color="#1A2040")
            self._cat_open[cat_key]=True

    def _build_tab(self,key):
        if key in self._tab_built: return
        builder=self._tab_builders.get(key)
        if not builder: return
        if key in self._no_scroll_tabs:
            outer=ctk.CTkFrame(self.content,fg_color=BG_MAIN,corner_radius=0)
        else:
            outer=ctk.CTkScrollableFrame(self.content,fg_color=BG_MAIN,corner_radius=0,
                scrollbar_button_color=BORDER,scrollbar_button_hover_color=ACCENT)
        builder(outer); self.tabs[key]=outer; self._tab_built.add(key)

    def _show_tab(self,key):
        self._build_tab(key)
        prev=getattr(self,"_active_tab",None)
        if prev and prev in self.tabs: self.tabs[prev].place_forget()
        self.tabs[key].place(relx=0,rely=0,relwidth=1,relheight=1)
        self.tabs[key].lift(); self._active_tab=key
        for k in ({prev,key}-{None}):
            btn=self.nav_btns.get(k)
            if not btn: continue
            if k==key:
                btn.configure(fg_color=ACCENT,text_color="white",
                              font=FBOLD(13) if k=="home" else FBOLD(12))
            else:
                btn.configure(fg_color="transparent",text_color=TEXT_PRI,
                              font=FBOLD(13) if k=="home" else F(12))

    # ── 공통 헬퍼 ────────────────────────────────────────────────────
    def _card(self,parent,**kw):
        return ctk.CTkFrame(parent,fg_color=BG_CARD,corner_radius=16,
                            border_width=1,border_color=BORDER,**kw)
    def _header(self,parent,icon,title,sub):
        hf=ctk.CTkFrame(parent,fg_color="transparent"); hf.pack(fill="x",padx=28,pady=(24,6))
        ctk.CTkLabel(hf,text=icon+"  "+title,font=FBOLD(22),text_color=TEXT_PRI).pack(anchor="w")
        ctk.CTkLabel(hf,text=sub,font=F(13),text_color=TEXT_SUB).pack(anchor="w",pady=(3,0))
    def _tip_box(self,parent,lines):
        box=ctk.CTkFrame(parent,fg_color=BG_TIP,corner_radius=12,border_width=1,border_color=BORDER)
        box.pack(fill="x",padx=28,pady=(0,10))
        ctk.CTkLabel(box,text="💡  사용법",font=FBOLD(12),text_color=ACCENT).pack(anchor="w",padx=16,pady=(10,4))
        for line in lines:
            ctk.CTkLabel(box,text=line,font=F(12),text_color=TEXT_SUB,
                         anchor="w",wraplength=720).pack(anchor="w",padx=22,pady=2)
        ctk.CTkFrame(box,height=8,fg_color="transparent").pack()
    def _logbox(self,parent):
        tb=ctk.CTkTextbox(parent,font=FMONO(12),fg_color=BG_DARK,text_color="#B8C4DC",
                          corner_radius=12,border_width=1,border_color=BORDER)
        def _setup(tb=tb):
            t=tb._textbox
            t.tag_config("ts",foreground="#4A5568"); t.tag_config("ok",foreground="#34D9A0")
            t.tag_config("err",foreground="#F47171"); t.tag_config("warn",foreground="#F9C74F")
            t.tag_config("info",foreground="#5B9CF6"); t.tag_config("body",foreground="#B8C4DC")
        tb.after(50,_setup); return tb
    def _log(self,box,msg):
        def _do():
            try:
                ts="["+time.strftime('%H:%M:%S')+"]"
                t=box._textbox; t.configure(state="normal")
                t.insert("end",ts,"ts"); t.insert("end","  ")
                if msg.startswith("✅") or msg.startswith("🎉"): tag="ok"
                elif msg.startswith("❌"): tag="err"
                elif msg.startswith("⚠️"): tag="warn"
                elif msg.startswith("▶") or msg.startswith("➕"): tag="info"
                else: tag="body"
                t.insert("end",msg+"\n",tag); t.configure(state="disabled"); box.see("end")
            except Exception: pass
        self.after(0,_do)
    def _abtn(self,parent,text,command,fg_color=None,**kw):
        kw.setdefault("font",F(13)); kw.setdefault("corner_radius",10)
        return make_animated_btn(parent,text=text,command=command,fg_color=fg_color or ACCENT,**kw)
    def _btn_working(self,btn,txt="⏳ 처리 중..."):
        btn._original_text=btn.cget("text"); btn._original_color=btn.cget("fg_color")
        btn.configure(state="disabled",text=txt,fg_color="#2A3450")
    def _btn_done(self,btn):
        ot=getattr(btn,"_original_text",btn.cget("text"))
        oc=getattr(btn,"_original_color",ACCENT)
        btn.configure(state="normal",text=ot,fg_color=oc)
    def _toast(self,msg,kind="ok",duration=2800):
        colors={"ok":("#1C3A2A","#34D9A0"),"err":("#3A1C1C","#F47171"),
                "warn":("#3A2E0A","#F9C74F"),"info":("#1C2A3A","#5B9CF6")}
        bg,fg=colors.get(kind,colors["ok"])
        toast=ctk.CTkFrame(self,fg_color=bg,corner_radius=12,border_width=1,border_color=fg)
        icon={"ok":"✅","err":"❌","warn":"⚠️","info":"ℹ️"}.get(kind,"✅")
        ctk.CTkLabel(toast,text=icon+"  "+msg,font=FBOLD(13),text_color=fg).pack(padx=18,pady=12)
        toast.place(relx=1.0,rely=1.0,anchor="se",x=-20,y=-20); toast.lift()
        self.after(duration,lambda:toast.place_forget() or toast.destroy())
    def _add_tooltip(self,widget,text):
        tip=[None]
        def _show(e):
            x=widget.winfo_rootx()+10; y=widget.winfo_rooty()+widget.winfo_height()+4
            tw=tk.Toplevel(widget); tw.wm_overrideredirect(True)
            tw.wm_geometry("+"+str(x)+"+"+str(y)); tw.configure(bg="#1A2040")
            tk.Label(tw,text=text,bg="#1A2040",fg="#C8D0E8",
                     font=("Consolas",11),padx=10,pady=5).pack()
            tip[0]=tw
        def _hide(e):
            if tip[0]:
                try: tip[0].destroy()
                except Exception: pass
                tip[0]=None
        widget.bind("<Enter>",_show); widget.bind("<Leave>",_hide)
    def _file_list_widget(self,parent,label="파일을 추가하세요"):
        frame=self._card(parent); frame.pack(fill="x",padx=28,pady=(0,10))
        top=ctk.CTkFrame(frame,fg_color="transparent"); top.pack(fill="x",padx=14,pady=(12,6))
        ctk.CTkLabel(top,text="📋  파일 목록",font=FBOLD(13),text_color=TEXT_PRI).pack(side="left")
        btn_row=ctk.CTkFrame(top,fg_color="transparent"); btn_row.pack(side="right")
        lb=ctk.CTkTextbox(frame,height=88,font=FMONO(12),fg_color=BG_DARK,
                          text_color="#B8C4DC",corner_radius=8,border_width=1,border_color=BORDER)
        lb.pack(fill="x",padx=14,pady=(0,12))
        lb.insert("0.0","  "+label); lb.configure(state="disabled")
        sel=[]
        def add(ft):
            for p in filedialog.askopenfilenames(filetypes=ft):
                if p not in sel: sel.append(p)
            _ref()
        def clear(): sel.clear(); _ref()
        def _ref():
            lb.configure(state="normal"); lb.delete("0.0","end")
            if sel:
                for f in sel: lb.insert("end","  • "+os.path.basename(f)+"\n")
            else: lb.insert("0.0","  "+label)
            lb.configure(state="disabled")
        return frame,btn_row,sel,add,clear

    # ══════════════════════════════════════════
    # 🏠  홈
    # ══════════════════════════════════════════
    def _build_home(self,parent):
        banner=GradientCanvas(parent,height=100,bg=BG_MAIN); banner.pack(fill="x")
        ctk.CTkLabel(banner,text="Filelo — 파일·문서·AI 자동화",
                     font=FBOLD(25),text_color=TEXT_PRI).place(x=32,y=18)
        ctk.CTkLabel(banner,text="파일 · 문서 · 이미지 · AI  —  22가지 자동화 기능을 한 곳에서",
                     font=F(13),text_color=TEXT_SUB).place(x=34,y=58)
        scroll=ctk.CTkScrollableFrame(parent,fg_color=BG_MAIN,
            scrollbar_button_color=BORDER,scrollbar_button_hover_color=ACCENT)
        scroll.pack(fill="both",expand=True)
        HOME_CATS=[
            ("📁 파일 관리","file",[
                ("🚀","파일명 번역","DeepL 영문→한글","translate",ACCENT),
                ("📁","폴더 자동 정리","확장자별 분류","folder",ACCENT2),
                ("🏷️","파일명 규칙 변경","일괄 리네임","rename",SUCCESS),
                ("📂","과제 폴더 생성","과목별 구조","task_dir",WARNING),
            ]),
            ("📄 문서 처리","doc",[
                ("📄","PDF 변환&추출","텍스트·이미지","pdf",WARNING),
                ("🔗","PDF 합치기/쪼개기","병합·분리","pdfmerge",ACCENT),
                ("🔒","PDF 비밀번호","설정·해제","pdfpwd",DANGER),
                ("🧹","메타데이터 삭제","디지털 지문 제거","meta",TEXT_SUB),
                ("📊","표 → 엑셀","표 자동 변환","table2xl",SUCCESS),
            ]),
            ("🖼️ 이미지","img",[
                ("🖼️","이미지 일괄처리","리사이즈·변환","image",ACCENT2),
                ("📎","이미지→PDF","묶어 PDF로","imgpdf",SUCCESS),
                ("🖼️","이미지 일괄추출","문서에서 추출","imgext",WARNING),
                ("💧","워터마크","삽입","watermark",ACCENT),
                ("📸","이미지 OCR","텍스트 추출","ocr",ACCENT2),
                ("✂️","배경 제거","AI 배경 제거","rembg",DANGER),
            ]),
            ("🤖 AI 도구","ai",[
                ("🤖","AI 문서 요약","Gemini 요약","summary",ACCENT),
                ("✏️","AI 문서 초안","초안 자동 작성","draft",ACCENT2),
                ("📚","참고문헌 정리","APA·MLA 자동화","citation",ACCENT2),
            ]),
            ("📊 데이터","data",[
                ("📊","엑셀 자동화","시트·중복·변환","excel",SUCCESS),
            ]),
            ("📅 학습 관리","study",[
                ("📅","마감 트래커","D-day 관리","tracker",DANGER),
            ]),
            ("⚙️ 설정","settings",[
                ("⚙️","API 키 설정","DeepL · Gemini 키 관리","settings",TEXT_SUB),
            ]),
        ]
        for cat_label,_,items in HOME_CATS:
            sec=ctk.CTkFrame(scroll,fg_color="transparent"); sec.pack(fill="x",padx=24,pady=(18,6))
            ctk.CTkLabel(sec,text=cat_label,font=FBOLD(16),text_color=TEXT_PRI).pack(side="left")
            ctk.CTkFrame(sec,height=1,fg_color=BORDER).pack(side="left",fill="x",expand=True,padx=(12,0),pady=8)
            grid=ctk.CTkFrame(scroll,fg_color="transparent"); grid.pack(fill="x",padx=24,pady=(0,4))
            cols=4
            for c in range(cols): grid.columnconfigure(c,weight=1,minsize=160)
            for i,(icon,title,desc,key,color) in enumerate(items):
                col=i%cols; row=i//cols; grid.rowconfigure(row,weight=0)
                card=ctk.CTkFrame(grid,fg_color=BG_CARD,corner_radius=14,border_width=1,border_color=BORDER)
                card.grid(row=row,column=col,padx=5,pady=5,sticky="nsew")
                def _on_enter(e,c=card,cl=color): c.configure(fg_color="#232840",border_color=cl)
                def _on_leave(e,c=card): c.configure(fg_color=BG_CARD,border_color=BORDER)
                card.bind("<Enter>",_on_enter); card.bind("<Leave>",_on_leave)
                card.bind("<Button-1>",lambda e,k=key:self._show_tab(k))
                ctk.CTkFrame(card,height=3,fg_color=color,corner_radius=0).pack(fill="x")
                inner=ctk.CTkFrame(card,fg_color="transparent"); inner.pack(fill="both",expand=True,padx=10,pady=8)
                il=ctk.CTkLabel(inner,text=icon,font=F(26)); il.pack(anchor="w")
                tl=ctk.CTkLabel(inner,text=title,font=FBOLD(13),text_color=TEXT_PRI,wraplength=140,justify="left"); tl.pack(anchor="w",pady=(2,0))
                dl=ctk.CTkLabel(inner,text=desc,font=F(11),text_color=TEXT_SUB,wraplength=140,justify="left"); dl.pack(anchor="w",pady=(1,6))
                ob=self._abtn(inner,"열기 →",lambda k=key:self._show_tab(k),fg_color=color,height=28,font=F(11)); ob.pack(fill="x")
                for w in (inner,il,tl,dl):
                    w.bind("<Enter>",_on_enter); w.bind("<Leave>",_on_leave)
                    w.bind("<Button-1>",lambda e,k=key:self._show_tab(k))
                self._add_tooltip(card,title+" — "+desc)

    # ══════════════════════════════════════════
    # 🚀  파일명 번역
    # ══════════════════════════════════════════
    def _build_translate(self,parent):
        self._header(parent,"🚀","파일명 번역","영문 파일명을 DeepL로 한글 변환합니다")
        self._tip_box(parent,[
            "① 파일을 직접 골라서 추가하거나, 폴더를 선택해 전체를 한번에 번역할 수 있습니다.",
            "② 파일 추가와 폴더 선택을 함께 쓸 경우 파일 추가가 우선 적용됩니다.",
            "③ 원본 파일명은 사라지므로 중요한 파일은 미리 백업하세요.",
        ])
        fw,btn_row,self.translate_files,tr_add,tr_clear=self._file_list_widget(
            parent,"번역할 파일을 직접 추가하세요 (여러 개 선택 가능)")
        self._abtn(btn_row,"➕ 파일 추가",lambda:tr_add([("모든 파일","*.*")]),fg_color=SUCCESS,width=100,height=28).pack(side="left",padx=(0,6))
        self._abtn(btn_row,"🗑 초기화",tr_clear,fg_color=DANGER,width=80,height=28).pack(side="left")
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        folder_row=ctk.CTkFrame(card,fg_color="transparent"); folder_row.pack(fill="x",padx=16,pady=(14,6))
        ctk.CTkLabel(folder_row,text="또는 폴더 전체 선택:",font=FBOLD(13),text_color=TEXT_PRI).pack(side="left",padx=(0,10))
        self.translate_path=ctk.StringVar(value="선택된 폴더 없음")
        ctk.CTkLabel(folder_row,textvariable=self.translate_path,font=F(12),text_color=TEXT_SUB,wraplength=400).pack(side="left")
        self._abtn(folder_row,"📂 폴더 선택",self._translate_pick_folder,width=120,height=36).pack(side="right")
        self.translate_progress=ctk.CTkProgressBar(card,progress_color=ACCENT)
        self.translate_progress.set(0); self.translate_progress.pack(fill="x",padx=16,pady=6)
        self.btn_run_translate=self._abtn(card,"🚀 번역 시작",self._translate_start,height=42)
        self.btn_run_translate.pack(padx=16,pady=(4,14))
        self.translate_log=self._logbox(parent)
        self.translate_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
        self._selected_translate_folder=""

    def _translate_pick_folder(self):
        p=filedialog.askdirectory()
        if p: self._selected_translate_folder=p; self.translate_path.set(p)
    def _translate_start(self):
        threading.Thread(target=self._translate_run,daemon=True).start()
    def _translate_run(self):
        if not HAS_DEEPL: self._log(self.translate_log,"❌ DeepL 연결 실패. API 키를 확인하세요."); return
        if not DEEPL_API_KEY: self._log(self.translate_log,"❌ DeepL API 키가 없습니다. ⚙️ 설정 탭에서 등록하세요."); return
        self._btn_working(self.btn_run_translate,"⏳ 번역 중...")
        try:
            tr=deepl.Translator(DEEPL_API_KEY)
            if self.translate_files:
                targets=[(os.path.dirname(fp),os.path.basename(fp)) for fp in self.translate_files if re.search('[a-zA-Z]',os.path.basename(fp))]
            elif self._selected_translate_folder:
                folder=self._selected_translate_folder
                targets=[(folder,fn) for fn in os.listdir(folder) if os.path.isfile(os.path.join(folder,fn)) and re.search('[a-zA-Z]',fn) and not fn.startswith('.')]
            else:
                self._log(self.translate_log,"⚠️ 파일을 추가하거나 폴더를 선택하세요."); return
            if not targets: self._log(self.translate_log,"⚠️ 번역할 영문 파일명이 없습니다."); return
            for i,(folder,fn) in enumerate(targets):
                name,ext=os.path.splitext(fn)
                result=tr.translate_text(name,target_lang="KO")
                clean=re.sub(r'[\\/*?:"<>|]',"",result.text).strip()
                os.rename(os.path.join(folder,fn),os.path.join(folder,clean+ext))
                self._log(self.translate_log,"✅  "+fn+"  →  "+clean+ext)
                self.translate_progress.set((i+1)/len(targets))
            self._toast("파일명 번역 완료! ("+str(len(targets))+"개)","ok")
        except Exception as e: self._log(self.translate_log,"❌ "+str(e))
        finally: self._btn_done(self.btn_run_translate)

    # ══════════════════════════════════════════
    # 📁  폴더 자동 정리
    # ══════════════════════════════════════════
    def _build_folder(self,parent):
        self._header(parent,"📁","폴더 자동 정리","확장자별로 파일을 자동 분류합니다")
        self._tip_box(parent,["① 정리할 폴더를 선택하세요.","② 정리 시작을 누르면 이미지·문서·영상·음악·압축파일·코드 폴더로 자동 분류됩니다."])
        EXT_MAP={"이미지":[".jpg",".jpeg",".png",".gif",".bmp",".webp",".svg",".ico"],
                 "문서":[".pdf",".docx",".doc",".hwp",".hwpx",".txt",".pptx",".xlsx",".csv"],
                 "영상":[".mp4",".mov",".avi",".mkv",".wmv",".flv"],
                 "음악":[".mp3",".wav",".flac",".aac",".ogg"],
                 "압축파일":[".zip",".rar",".7z",".tar",".gz"],
                 "코드":[".py",".js",".ts",".html",".css",".java",".c",".cpp"],"기타":[]}
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        self.folder_path_var=ctk.StringVar(value="정리할 폴더를 선택하세요")
        row=ctk.CTkFrame(card,fg_color="transparent"); row.pack(fill="x",padx=16,pady=(14,8))
        ctk.CTkLabel(row,textvariable=self.folder_path_var,font=F(12),text_color=TEXT_SUB,wraplength=520).pack(side="left")
        self._abtn(row,"📂 폴더 선택",lambda:self._folder_pick(self.folder_path_var),width=120,height=36).pack(side="right")
        rule_row=ctk.CTkFrame(card,fg_color="transparent"); rule_row.pack(fill="x",padx=16,pady=(0,8))
        for (cat,_),color in zip(EXT_MAP.items(),[ACCENT,SUCCESS,WARNING,DANGER,"#A78BFA","#FB923C",TEXT_SUB]):
            ctk.CTkLabel(rule_row,text=cat,font=F(11),text_color=color).pack(side="left",padx=6)
        self.btn_run_folder=self._abtn(card,"📁 정리 시작",lambda:threading.Thread(target=self._run_folder_sort,args=(EXT_MAP,),daemon=True).start(),height=42)
        self.btn_run_folder.configure(state="disabled"); self.btn_run_folder.pack(padx=16,pady=(0,14))
        self.folder_log=self._logbox(parent); self.folder_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
        self._selected_folder_path=""
    def _folder_pick(self,var):
        p=filedialog.askdirectory()
        if p: self._selected_folder_path=p; var.set(p); self.btn_run_folder.configure(state="normal")
    def _run_folder_sort(self,ext_map):
        self._btn_working(self.btn_run_folder,"⏳ 정리 중...")
        base=self._selected_folder_path
        files=[f for f in os.listdir(base) if os.path.isfile(os.path.join(base,f)) and not f.startswith('.')]
        moved=0
        for fn in files:
            ext=Path(fn).suffix.lower(); dest_cat="기타"
            for cat,exts in ext_map.items():
                if ext in exts: dest_cat=cat; break
            dest=os.path.join(base,dest_cat); os.makedirs(dest,exist_ok=True)
            shutil.move(os.path.join(base,fn),os.path.join(dest,fn))
            self._log(self.folder_log,"✅  ["+dest_cat+"]  "+fn); moved+=1
        self._log(self.folder_log,"🎉  완료! 총 "+str(moved)+"개")
        self._toast(str(moved)+"개 파일 정리 완료!","ok"); self._btn_done(self.btn_run_folder)

    # ══════════════════════════════════════════
    # 🏷️  파일명 규칙 변경
    # ══════════════════════════════════════════
    def _build_rename(self,parent):
        self._header(parent,"🏷️","파일명 규칙 일괄 변경","폴더 안의 파일명을 정해진 규칙으로 일괄 변경합니다")
        self._tip_box(parent,["① 폴더를 선택하세요.","② 규칙 패턴: {num}=번호, {name}=원본이름, {date}=오늘날짜",
            "   예) {date}_{num:03d}_{name}  →  20240115_001_파일명.jpg",
            "③ 시작 번호와 확장자 필터(비워두면 전체)를 설정하고 미리보기로 먼저 확인하세요."])
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        inner=ctk.CTkFrame(card,fg_color="transparent"); inner.pack(fill="x",padx=16,pady=14)
        inner.columnconfigure(1,weight=1)
        ctk.CTkLabel(inner,text="폴더 선택:",font=FBOLD(13),text_color=TEXT_PRI).grid(row=0,column=0,padx=(0,10),pady=7,sticky="w")
        self.rename_path_var=ctk.StringVar(value="")
        ctk.CTkEntry(inner,textvariable=self.rename_path_var,font=F(12),placeholder_text="폴더 경로").grid(row=0,column=1,sticky="ew",pady=7)
        self._abtn(inner,"선택",self._rename_pick,width=70,height=32).grid(row=0,column=2,padx=(8,0))
        ctk.CTkLabel(inner,text="규칙 패턴:",font=FBOLD(13),text_color=TEXT_PRI).grid(row=1,column=0,padx=(0,10),pady=7,sticky="w")
        self.rename_pattern=ctk.CTkEntry(inner,font=F(12),placeholder_text="예) {date}_{num:03d}_{name}")
        self.rename_pattern.grid(row=1,column=1,columnspan=2,sticky="ew",pady=7)
        row2=ctk.CTkFrame(inner,fg_color="transparent"); row2.grid(row=2,column=0,columnspan=3,sticky="ew",pady=7)
        ctk.CTkLabel(row2,text="시작 번호:",font=FBOLD(13),text_color=TEXT_PRI).pack(side="left")
        self.rename_start=ctk.CTkEntry(row2,width=80,font=F(12),placeholder_text="1"); self.rename_start.pack(side="left",padx=(6,20))
        ctk.CTkLabel(row2,text="확장자 필터:",font=FBOLD(13),text_color=TEXT_PRI).pack(side="left")
        self.rename_ext_filter=ctk.CTkEntry(row2,width=120,font=F(12),placeholder_text="예) .jpg"); self.rename_ext_filter.pack(side="left",padx=6)
        btn_row=ctk.CTkFrame(card,fg_color="transparent"); btn_row.pack(fill="x",padx=16,pady=(0,14))
        self._abtn(btn_row,"👁 미리보기",self._rename_preview,fg_color=ACCENT2,height=40,width=140).pack(side="left",padx=(0,8))
        self._abtn(btn_row,"✅ 적용",self._rename_apply,fg_color=SUCCESS,height=40,width=120).pack(side="left")
        self.rename_log=self._logbox(parent); self.rename_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
        self._rename_folder=""
    def _rename_pick(self):
        p=filedialog.askdirectory()
        if p: self._rename_folder=p; self.rename_path_var.set(p)
    def _get_rename_list(self):
        folder=self._rename_folder
        if not folder: return None,[]
        ext_f=self.rename_ext_filter.get().strip().lower()
        files=sorted([f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder,f)) and not f.startswith('.') and (not ext_f or Path(f).suffix.lower()==ext_f)])
        pattern=self.rename_pattern.get().strip(); start=int(self.rename_start.get().strip() or "1")
        today=datetime.date.today().strftime("%Y%m%d"); pairs=[]
        for i,fn in enumerate(files):
            stem=Path(fn).stem; ext=Path(fn).suffix
            try: new_stem=pattern.format(num=start+i,name=stem,date=today); new_fn=new_stem+ext
            except Exception: new_fn=fn
            pairs.append((fn,new_fn))
        return folder,pairs
    def _rename_preview(self):
        folder,pairs=self._get_rename_list()
        if not folder: messagebox.showwarning("안내","폴더를 먼저 선택하세요."); return
        self.rename_log.delete("0.0","end"); self._log(self.rename_log,"[ 미리보기 — 실제 변경되지 않습니다 ]")
        for old,new in pairs: self._log(self.rename_log,"  "+old+"  →  "+new)
    def _rename_apply(self):
        folder,pairs=self._get_rename_list()
        if not folder: messagebox.showwarning("안내","폴더를 먼저 선택하세요."); return
        if not messagebox.askyesno("확인","총 "+str(len(pairs))+"개 파일명을 변경합니다.\n계속하시겠습니까?"): return
        self.rename_log.delete("0.0","end")
        for old,new in pairs:
            try: os.rename(os.path.join(folder,old),os.path.join(folder,new)); self._log(self.rename_log,"✅  "+old+"  →  "+new)
            except Exception as e: self._log(self.rename_log,"❌  "+old+": "+str(e))
        self._toast("파일명 일괄 변경 완료!","ok")

    # ══════════════════════════════════════════
    # 📂  과제 폴더 생성
    # ══════════════════════════════════════════
    def _build_task_dir(self,parent):
        self._header(parent,"📂","과제 폴더 생성","과목명 입력 시 날짜·제출본·참고자료 구조 자동 생성")
        self._tip_box(parent,["① 저장 위치를 선택하세요.","② 과목명을 쉼표로 구분해서 입력하세요. 예) 소방학개론, 위험물","③ 각 과목별로 날짜별·제출본·참고자료·필기 폴더가 생성됩니다."])
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        inner=ctk.CTkFrame(card,fg_color="transparent"); inner.pack(fill="x",padx=16,pady=14); inner.columnconfigure(1,weight=1)
        ctk.CTkLabel(inner,text="저장 위치",font=FBOLD(13),text_color=TEXT_PRI).grid(row=0,column=0,padx=(0,10),pady=8,sticky="w")
        self.taskdir_base_var=ctk.StringVar()
        ctk.CTkEntry(inner,textvariable=self.taskdir_base_var,placeholder_text="폴더 경로",font=F(12)).grid(row=0,column=1,sticky="ew",pady=8)
        self._abtn(inner,"선택",self._taskdir_pick_base,width=72,height=34).grid(row=0,column=2,padx=(8,0))
        ctk.CTkLabel(inner,text="과목명\n(쉼표 구분)",font=FBOLD(13),text_color=TEXT_PRI).grid(row=1,column=0,padx=(0,10),pady=8,sticky="w")
        self.taskdir_subjects=ctk.CTkEntry(inner,placeholder_text="예) 소방학개론, 위험물, 화재조사론",font=F(12))
        self.taskdir_subjects.grid(row=1,column=1,columnspan=2,sticky="ew",pady=8)
        self._abtn(card,"📂 폴더 생성",self._taskdir_create,fg_color=SUCCESS,height=42).pack(padx=16,pady=(6,14))
        self.taskdir_log=self._logbox(parent); self.taskdir_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
    def _taskdir_pick_base(self):
        p=filedialog.askdirectory()
        if p: self.taskdir_base_var.set(p)
    def _taskdir_create(self):
        base=self.taskdir_base_var.get().strip(); raw=self.taskdir_subjects.get().strip()
        if not base or not raw: messagebox.showwarning("입력 오류","저장 위치와 과목명을 모두 입력하세요."); return
        subjects=[s.strip() for s in raw.split(",") if s.strip()]
        today=datetime.datetime.now().strftime("%Y-%m-%d")
        for subj in subjects:
            for sd in ["날짜별","제출본","참고자료","필기"]: os.makedirs(os.path.join(base,subj,sd),exist_ok=True)
            os.makedirs(os.path.join(base,subj,"날짜별",today),exist_ok=True)
            self._log(self.taskdir_log,"✅  ["+subj+"] 폴더 생성 완료")
        self._toast(str(len(subjects))+"개 과목 폴더 생성!","ok")

    # ══════════════════════════════════════════
    # 📄  PDF 변환 & 추출
    # ══════════════════════════════════════════
    def _build_pdf(self,parent):
        self._header(parent,"📄","PDF 변환 & 추출","PDF에서 텍스트 또는 이미지를 추출합니다")
        self._tip_box(parent,["① 파일 추가 버튼으로 PDF를 여러 개 선택할 수 있습니다.","② 텍스트 추출: 각 PDF와 같은 위치에 _추출텍스트.txt 가 저장됩니다.","③ 이미지 추출: 각 PDF와 같은 위치에 _이미지들 폴더가 생성됩니다."])
        fw,btn_row,self.pdf_files,pdf_add,pdf_clear=self._file_list_widget(parent,"PDF 파일을 추가하세요")
        self._abtn(btn_row,"➕ 추가",lambda:pdf_add([("PDF","*.pdf")]),fg_color=SUCCESS,width=90,height=28).pack(side="left",padx=(0,6))
        self._abtn(btn_row,"🗑 초기화",pdf_clear,fg_color=DANGER,width=80,height=28).pack(side="left")
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        mr=ctk.CTkFrame(card,fg_color="transparent"); mr.pack(fill="x",padx=16,pady=(12,8))
        ctk.CTkLabel(mr,text="추출 모드:",font=FBOLD(13),text_color=TEXT_PRI).pack(side="left",padx=(0,12))
        self.pdf_mode=ctk.StringVar(value="text")
        ctk.CTkRadioButton(mr,text="텍스트 (TXT)",variable=self.pdf_mode,value="text",font=F(13)).pack(side="left",padx=8)
        ctk.CTkRadioButton(mr,text="이미지 (PNG/JPG)",variable=self.pdf_mode,value="image",font=F(13)).pack(side="left",padx=8)
        self._abtn(card,"📄 추출 시작",lambda:threading.Thread(target=self._pdf_run,daemon=True).start(),height=42).pack(padx=16,pady=(0,14))
        self.pdf_log=self._logbox(parent); self.pdf_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
    def _pdf_run(self):
        if not HAS_FITZ: self._log(self.pdf_log,"❌ pymupdf 패키지가 필요합니다."); return
        if not self.pdf_files: self._log(self.pdf_log,"⚠️ 파일을 먼저 추가하세요."); return
        try:
            for pdf_path in self.pdf_files:
                self._log(self.pdf_log,"▶  "+os.path.basename(pdf_path))
                doc=fitz.open(pdf_path); base=os.path.splitext(pdf_path)[0]
                if self.pdf_mode.get()=="text":
                    txt="".join(p.get_text() for p in doc)
                    with open(base+"_추출텍스트.txt","w",encoding="utf-8") as f: f.write(txt)
                    self._log(self.pdf_log,"  ✅  텍스트 저장")
                else:
                    img_dir=base+"_이미지들"; os.makedirs(img_dir,exist_ok=True); cnt=0
                    for i,pg in enumerate(doc):
                        for j,img in enumerate(pg.get_images(full=True)):
                            bi=doc.extract_image(img[0])
                            with open(os.path.join(img_dir,"p"+str(i+1)+"_"+str(j+1)+"."+bi['ext']),"wb") as f: f.write(bi["image"])
                            cnt+=1
                    self._log(self.pdf_log,"  ✅  이미지 "+str(cnt)+"개 추출")
            self._toast("PDF 추출 완료!","ok")
        except Exception as e: self._log(self.pdf_log,"❌ "+str(e))

    # ══════════════════════════════════════════
    # 🔗  PDF 합치기/쪼개기
    # ══════════════════════════════════════════
    def _build_pdfmerge(self,parent):
        self._header(parent,"🔗","PDF 합치기 / 쪼개기","여러 PDF 병합 또는 원하는 페이지만 분리")
        self._tip_box(parent,["① [합치기] 파일 추가 → 저장 파일명 입력 → 합치기 실행","② [쪼개기] PDF 파일 선택 → 페이지 범위 입력 (예: 1-3 또는 1,3,5)"])
        mc=self._card(parent); mc.pack(fill="x",padx=28,pady=(0,8))
        ctk.CTkLabel(mc,text="🔗  PDF 합치기",font=FBOLD(14),text_color=TEXT_PRI).pack(anchor="w",padx=16,pady=(12,4))
        mw,mbr,self.merge_files,madd,mclr=self._file_list_widget(mc,"합칠 PDF 순서대로 추가")
        mw.configure(fg_color="transparent",border_width=0); mw.pack_configure(padx=0,pady=0)
        self._abtn(mbr,"➕ 추가",lambda:madd([("PDF","*.pdf")]),fg_color=SUCCESS,width=90,height=28).pack(side="left",padx=(0,6))
        self._abtn(mbr,"🗑 초기화",mclr,fg_color=DANGER,width=80,height=28).pack(side="left")
        mnr=ctk.CTkFrame(mc,fg_color="transparent"); mnr.pack(fill="x",padx=16,pady=(4,4))
        ctk.CTkLabel(mnr,text="저장 파일명:",font=FBOLD(13),text_color=TEXT_PRI).pack(side="left",padx=(0,8))
        self.merge_out_name=ctk.CTkEntry(mnr,placeholder_text="예) 최종제출본",font=F(12),width=200); self.merge_out_name.pack(side="left")
        ctk.CTkLabel(mnr,text=".pdf",font=F(12),text_color=TEXT_SUB).pack(side="left",padx=4)
        self._abtn(mc,"🔗 합치기 실행",lambda:threading.Thread(target=self._merge_run,daemon=True).start(),height=38).pack(padx=16,pady=(4,12))
        sc=self._card(parent); sc.pack(fill="x",padx=28,pady=(0,10))
        ctk.CTkLabel(sc,text="✂️  PDF 쪼개기",font=FBOLD(14),text_color=TEXT_PRI).pack(anchor="w",padx=16,pady=(12,4))
        sr=ctk.CTkFrame(sc,fg_color="transparent"); sr.pack(fill="x",padx=16,pady=(0,6))
        self.split_path_var=ctk.StringVar(value="쪼갤 PDF 파일 선택")
        ctk.CTkLabel(sr,textvariable=self.split_path_var,font=F(12),text_color=TEXT_SUB,wraplength=480).pack(side="left")
        self._abtn(sr,"📂 파일 선택",self._split_pick,width=110,height=36).pack(side="right")
        rr=ctk.CTkFrame(sc,fg_color="transparent"); rr.pack(fill="x",padx=16,pady=(0,4))
        ctk.CTkLabel(rr,text="페이지 범위:",font=FBOLD(13),text_color=TEXT_PRI).pack(side="left",padx=(0,8))
        self.split_range=ctk.CTkEntry(rr,placeholder_text="예) 1-5  또는  1,3,5",font=F(12),width=280); self.split_range.pack(side="left")
        self._abtn(sc,"✂️ 쪼개기 실행",lambda:threading.Thread(target=self._split_run,daemon=True).start(),fg_color=WARNING,height=38).pack(padx=16,pady=(4,12))
        self.pdfmerge_log=self._logbox(parent); self.pdfmerge_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
        self._split_pdf_path=""
    def _split_pick(self):
        p=filedialog.askopenfilename(filetypes=[("PDF","*.pdf")])
        if p: self._split_pdf_path=p; self.split_path_var.set(os.path.basename(p))
    def _merge_run(self):
        if not HAS_FITZ: self._log(self.pdfmerge_log,"❌ pymupdf 패키지가 필요합니다."); return
        if len(self.merge_files)<2: self._log(self.pdfmerge_log,"⚠️ PDF 2개 이상 추가하세요."); return
        try:
            merged=fitz.open()
            for p in self.merge_files:
                doc=fitz.open(p); merged.insert_pdf(doc)
                self._log(self.pdfmerge_log,"➕  "+os.path.basename(p)+" ("+str(doc.page_count)+"p)")
            out=self.merge_out_name.get().strip() or "합본"
            op=os.path.join(os.path.dirname(self.merge_files[0]),out+".pdf")
            merged.save(op); self._log(self.pdfmerge_log,"✅  "+out+".pdf (총 "+str(merged.page_count)+"p)")
            self._toast(out+".pdf 저장 완료!","ok")
        except Exception as e: self._log(self.pdfmerge_log,"❌ "+str(e))
    def _split_run(self):
        if not HAS_FITZ: self._log(self.pdfmerge_log,"❌ pymupdf 패키지가 필요합니다."); return
        if not self._split_pdf_path: self._log(self.pdfmerge_log,"⚠️ PDF 파일 선택 먼저"); return
        rs=self.split_range.get().strip()
        if not rs: self._log(self.pdfmerge_log,"⚠️ 페이지 범위 입력하세요."); return
        try:
            pages=set()
            for part in rs.split(","):
                part=part.strip()
                if "-" in part:
                    s,e=part.split("-"); pages.update(range(int(s)-1,int(e)))
                else: pages.add(int(part)-1)
            pages=sorted(pages); src=fitz.open(self._split_pdf_path); out=fitz.open()
            for p in pages:
                if 0<=p<src.page_count: out.insert_pdf(src,from_page=p,to_page=p)
            op=os.path.splitext(self._split_pdf_path)[0]+"_분리.pdf"
            out.save(op); self._log(self.pdfmerge_log,"✅  분리 완료: "+os.path.basename(op)+" ("+str(len(pages))+"p)")
            self._toast("PDF 쪼개기 완료!","ok")
        except Exception as e: self._log(self.pdfmerge_log,"❌ "+str(e))

    # ══════════════════════════════════════════
    # 🔒  PDF 비밀번호
    # ══════════════════════════════════════════
    def _build_pdfpwd(self,parent):
        self._header(parent,"🔒","PDF 비밀번호 설정 / 해제","PDF에 비밀번호를 걸거나 일괄 해제합니다")
        self._tip_box(parent,["① 파일 추가 → 비밀번호 입력 → 비밀번호 설정 또는 해제를 선택하세요.","② 설정: 각 파일 옆에 _잠금.pdf 로 저장됩니다.","③ 해제: 기존 비밀번호를 입력하면 _해제.pdf 로 저장됩니다."])
        fw,btn_row,self.pwd_files,pwd_add,pwd_clear=self._file_list_widget(parent,"비밀번호를 설정/해제할 PDF 추가")
        self._abtn(btn_row,"➕ 추가",lambda:pwd_add([("PDF","*.pdf")]),fg_color=SUCCESS,width=90,height=28).pack(side="left",padx=(0,6))
        self._abtn(btn_row,"🗑 초기화",pwd_clear,fg_color=DANGER,width=80,height=28).pack(side="left")
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        pr=ctk.CTkFrame(card,fg_color="transparent"); pr.pack(fill="x",padx=16,pady=(12,8))
        ctk.CTkLabel(pr,text="비밀번호:",font=FBOLD(13),text_color=TEXT_PRI).pack(side="left",padx=(0,8))
        self.pdf_pwd_entry=ctk.CTkEntry(pr,placeholder_text="비밀번호 입력",show="●",font=F(12),width=200); self.pdf_pwd_entry.pack(side="left")
        br=ctk.CTkFrame(card,fg_color="transparent"); br.pack(fill="x",padx=16,pady=(0,14))
        self._abtn(br,"🔒 비밀번호 설정",lambda:threading.Thread(target=lambda:self._pdf_pwd_run("set"),daemon=True).start(),fg_color=DANGER,height=40,width=160).pack(side="left",padx=(0,10))
        self._abtn(br,"🔓 비밀번호 해제",lambda:threading.Thread(target=lambda:self._pdf_pwd_run("remove"),daemon=True).start(),fg_color=SUCCESS,height=40,width=160).pack(side="left")
        self.pwd_log=self._logbox(parent); self.pwd_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
    def _pdf_pwd_run(self,mode):
        if not HAS_FITZ: self._log(self.pwd_log,"❌ pymupdf 패키지가 필요합니다."); return
        if not self.pwd_files: self._log(self.pwd_log,"⚠️ 파일을 먼저 추가하세요."); return
        pwd=self.pdf_pwd_entry.get().strip()
        if not pwd: self._log(self.pwd_log,"⚠️ 비밀번호를 입력하세요."); return
        for fp in self.pwd_files:
            try:
                doc=fitz.open(fp)
                if mode=="set":
                    op=os.path.splitext(fp)[0]+"_잠금.pdf"
                    doc.save(op,encryption=fitz.PDF_ENCRYPT_AES_256,owner_pw=pwd,user_pw=pwd)
                    self._log(self.pwd_log,"✅  설정 완료: "+os.path.basename(op))
                else:
                    if doc.is_encrypted: doc.authenticate(pwd)
                    op=os.path.splitext(fp)[0]+"_해제.pdf"
                    doc.save(op,encryption=fitz.PDF_ENCRYPT_NONE)
                    self._log(self.pwd_log,"✅  해제 완료: "+os.path.basename(op))
            except Exception as e: self._log(self.pwd_log,"❌  "+os.path.basename(fp)+": "+str(e))
        self._toast("비밀번호 처리 완료!","ok")

    # ══════════════════════════════════════════
    # 🧹  메타데이터 삭제
    # ══════════════════════════════════════════
    def _build_meta(self,parent):
        self._header(parent,"🧹","메타데이터 완전 삭제","문서의 작성자·수정 이력·숨겨진 메모 등 디지털 지문을 제거합니다")
        self._tip_box(parent,["① 파일을 추가하세요 (PDF·DOCX 지원).","② 메타데이터 삭제를 누르면 작성자, 제목, 수정 시간 등이 모두 제거된 _clean 파일이 생성됩니다."])
        fw,btn_row,self.meta_files,meta_add,meta_clear=self._file_list_widget(parent,"PDF 또는 DOCX 파일을 추가하세요")
        self._abtn(btn_row,"➕ 추가",lambda:meta_add([("문서","*.pdf *.docx")]),fg_color=SUCCESS,width=90,height=28).pack(side="left",padx=(0,6))
        self._abtn(btn_row,"🗑 초기화",meta_clear,fg_color=DANGER,width=80,height=28).pack(side="left")
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        self._abtn(card,"🧹 메타데이터 삭제",lambda:threading.Thread(target=self._meta_run,daemon=True).start(),height=42).pack(padx=16,pady=14)
        self.meta_log=self._logbox(parent); self.meta_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
    def _meta_run(self):
        if not self.meta_files: self._log(self.meta_log,"⚠️ 파일을 먼저 추가하세요."); return
        for fp in self.meta_files:
            ext=Path(fp).suffix.lower()
            try:
                if ext==".pdf" and HAS_FITZ:
                    doc=fitz.open(fp); doc.set_metadata({})
                    op=os.path.splitext(fp)[0]+"_clean.pdf"; doc.save(op)
                    self._log(self.meta_log,"✅  PDF 메타 삭제: "+os.path.basename(op))
                elif ext==".docx" and HAS_DOCX:
                    doc=Document(fp); cp=doc.core_properties
                    for attr in ["author","last_modified_by","title","subject","keywords","comments"]:
                        try: setattr(cp,attr,"")
                        except Exception: pass
                    op=os.path.splitext(fp)[0]+"_clean.docx"; doc.save(op)
                    self._log(self.meta_log,"✅  DOCX 메타 삭제: "+os.path.basename(op))
                else: self._log(self.meta_log,"⚠️  "+os.path.basename(fp)+": 지원하지 않는 형식")
            except Exception as e: self._log(self.meta_log,"❌  "+os.path.basename(fp)+": "+str(e))
        self._toast("메타데이터 삭제 완료!","ok")

    # ══════════════════════════════════════════
    # 📊  표 → 엑셀 변환
    # ══════════════════════════════════════════
    def _build_table2xl(self,parent):
        self._header(parent,"📊","표 → 엑셀 변환","PDF 또는 DOCX 안의 표를 엑셀 시트로 자동 변환합니다")
        self._tip_box(parent,["① 파일을 추가하세요 (PDF·DOCX 지원).","② 변환 시작을 누르면 문서 안의 표가 각 시트로 분리된 엑셀 파일이 생성됩니다."])
        fw,btn_row,self.t2xl_files,t2xl_add,t2xl_clear=self._file_list_widget(parent,"PDF 또는 DOCX 파일 추가")
        self._abtn(btn_row,"➕ 추가",lambda:t2xl_add([("문서","*.pdf *.docx")]),fg_color=SUCCESS,width=90,height=28).pack(side="left",padx=(0,6))
        self._abtn(btn_row,"🗑 초기화",t2xl_clear,fg_color=DANGER,width=80,height=28).pack(side="left")
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        self._abtn(card,"📊 변환 시작",lambda:threading.Thread(target=self._table2xl_run,daemon=True).start(),height=42).pack(padx=16,pady=14)
        self.t2xl_log=self._logbox(parent); self.t2xl_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
    def _table2xl_run(self):
        if not HAS_XL: self._log(self.t2xl_log,"❌ openpyxl 패키지가 필요합니다."); return
        if not self.t2xl_files: self._log(self.t2xl_log,"⚠️ 파일을 먼저 추가하세요."); return
        for fp in self.t2xl_files:
            ext=Path(fp).suffix.lower()
            try:
                wb=openpyxl.Workbook(); wb.remove(wb.active); sheet_idx=0
                if ext==".docx" and HAS_DOCX:
                    doc=Document(fp)
                    for ti,table in enumerate(doc.tables):
                        ws=wb.create_sheet(title="표"+str(ti+1)); sheet_idx+=1
                        for row in table.rows: ws.append([cell.text.strip() for cell in row.cells])
                        self._log(self.t2xl_log,"  📋  표"+str(ti+1)+": "+str(len(table.rows))+"행 추출")
                elif ext==".pdf" and HAS_FITZ:
                    doc=fitz.open(fp)
                    for pi,page in enumerate(doc):
                        tabs=page.find_tables()
                        for ti,tab in enumerate(tabs):
                            ws=wb.create_sheet(title="p"+str(pi+1)+"_표"+str(ti+1)); sheet_idx+=1
                            for row in tab.extract(): ws.append([str(c) if c else "" for c in row])
                            self._log(self.t2xl_log,"  📋  "+str(pi+1)+"페이지 표"+str(ti+1)+" 추출")
                if sheet_idx==0: self._log(self.t2xl_log,"⚠️  "+os.path.basename(fp)+": 표를 찾지 못했습니다."); continue
                op=os.path.splitext(fp)[0]+"_표변환.xlsx"; wb.save(op)
                self._log(self.t2xl_log,"✅  "+os.path.basename(op)+" 저장 ("+str(sheet_idx)+"개 시트)")
            except Exception as e: self._log(self.t2xl_log,"❌  "+os.path.basename(fp)+": "+str(e))
        self._toast("표 → 엑셀 변환 완료!","ok")

    # ══════════════════════════════════════════
    # 🖼️  이미지 일괄 처리
    # ══════════════════════════════════════════
    def _build_image(self,parent):
        self._header(parent,"🖼️","이미지 일괄 처리","폴더 전체 또는 선택한 이미지 리사이즈·포맷 변환")
        self._tip_box(parent,["① 파일 추가 또는 폴더 선택 중 하나를 사용하세요.","② 최대 너비(px) 입력 시 비율 유지하며 리사이즈. 비워두면 원본 크기 유지.","③ 처리된 파일은 원본 위치의 '_처리완료' 폴더에 저장됩니다."])
        fw,btn_row,self.image_files,img_add,img_clear=self._file_list_widget(parent,"이미지를 추가하거나 아래 폴더 선택")
        self._abtn(btn_row,"➕ 추가",lambda:img_add([("이미지","*.jpg *.jpeg *.png *.bmp *.webp *.tiff")]),fg_color=SUCCESS,width=90,height=28).pack(side="left",padx=(0,6))
        self._abtn(btn_row,"🗑 초기화",img_clear,fg_color=DANGER,width=80,height=28).pack(side="left")
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        fr=ctk.CTkFrame(card,fg_color="transparent"); fr.pack(fill="x",padx=16,pady=(12,6))
        self.img_path_var=ctk.StringVar(value="또는 폴더 전체 선택")
        ctk.CTkLabel(fr,textvariable=self.img_path_var,font=F(12),text_color=TEXT_SUB,wraplength=480).pack(side="left")
        self._abtn(fr,"📂 폴더 선택",self._image_pick_folder,width=120,height=36).pack(side="right")
        opt=ctk.CTkFrame(card,fg_color="transparent"); opt.pack(fill="x",padx=16,pady=6); opt.columnconfigure((1,3),weight=1)
        ctk.CTkLabel(opt,text="최대 너비(px):",font=FBOLD(13),text_color=TEXT_PRI).grid(row=0,column=0,padx=(0,8),pady=6,sticky="w")
        self.img_width=ctk.CTkEntry(opt,placeholder_text="예) 1280",font=F(12),width=180); self.img_width.grid(row=0,column=1,sticky="w")
        ctk.CTkLabel(opt,text="출력 포맷:",font=FBOLD(13),text_color=TEXT_PRI).grid(row=0,column=2,padx=(20,8),sticky="w")
        self.img_format=ctk.CTkOptionMenu(opt,values=["원본 유지","JPEG","PNG","WEBP"],width=130,font=F(12)); self.img_format.grid(row=0,column=3,sticky="w")
        ctk.CTkLabel(opt,text="JPEG 품질:",font=FBOLD(13),text_color=TEXT_PRI).grid(row=1,column=0,padx=(0,8),pady=6,sticky="w")
        self.img_quality=ctk.CTkEntry(opt,placeholder_text="예) 85",font=F(12),width=180); self.img_quality.grid(row=1,column=1,sticky="w")
        self._abtn(card,"🖼️ 처리 시작",lambda:threading.Thread(target=self._image_run,daemon=True).start(),height=42).pack(padx=16,pady=(6,14))
        self.image_log=self._logbox(parent); self.image_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
        self._selected_img_folder=""
    def _image_pick_folder(self):
        p=filedialog.askdirectory()
        if p: self._selected_img_folder=p; self.img_path_var.set(p)
    def _image_run(self):
        if not HAS_PIL: self._log(self.image_log,"❌ Pillow 패키지가 필요합니다."); return
        try:
            IMG_EXTS={".jpg",".jpeg",".png",".bmp",".webp",".tiff"}
            mw=self.img_width.get().strip(); max_w=int(mw) if mw.isdigit() else None
            quality=int(self.img_quality.get().strip() or "85"); fmt_choice=self.img_format.get()
            if self.image_files:
                files=self.image_files; out_base=os.path.dirname(self.image_files[0])
            elif self._selected_img_folder:
                files=[os.path.join(self._selected_img_folder,f) for f in os.listdir(self._selected_img_folder) if Path(f).suffix.lower() in IMG_EXTS]
                out_base=self._selected_img_folder
            else: self._log(self.image_log,"⚠️ 파일 또는 폴더 선택 먼저"); return
            out_dir=os.path.join(out_base,"_처리완료"); os.makedirs(out_dir,exist_ok=True)
            for i,src in enumerate(files):
                img=Image.open(src)
                if max_w and img.width>max_w: img=img.resize((max_w,int(img.height*max_w/img.width)),Image.LANCZOS)
                stem=Path(src).stem
                out_ext=("."+fmt_choice.lower()) if fmt_choice!="원본 유지" else Path(src).suffix.lower()
                out_fmt="JPEG" if out_ext in (".jpg",".jpeg") else (fmt_choice if fmt_choice!="원본 유지" else out_ext[1:].upper())
                op=os.path.join(out_dir,stem+out_ext)
                if out_fmt=="JPEG": img=img.convert("RGB")
                img.save(op,format=out_fmt,**{"quality":quality} if out_fmt=="JPEG" else {})
                self._log(self.image_log,"✅  ("+str(i+1)+"/"+str(len(files))+") "+os.path.basename(src))
            self._toast(str(len(files))+"개 처리 완료!","ok")
        except Exception as e: self._log(self.image_log,"❌ "+str(e))

    # ══════════════════════════════════════════
    # 📎  이미지 → PDF
    # ══════════════════════════════════════════
    def _build_imgpdf(self,parent):
        self._header(parent,"📎","이미지 → PDF","선택한 이미지들을 순서대로 PDF 한 파일로 묶습니다")
        self._tip_box(parent,["① 파일 추가 버튼으로 이미지를 선택하세요. 추가 순서 = PDF 페이지 순서.","② 파일명을 입력하고 저장 폴더를 선택하세요."])
        fw,btn_row,self.imgpdf_files,ipadd,ipclear=self._file_list_widget(parent,"이미지 추가 (순서 = PDF 페이지 순서)")
        self._abtn(btn_row,"➕ 추가",lambda:ipadd([("이미지","*.jpg *.jpeg *.png *.bmp *.webp")]),fg_color=SUCCESS,width=90,height=28).pack(side="left",padx=(0,6))
        self._abtn(btn_row,"🗑 초기화",ipclear,fg_color=DANGER,width=80,height=28).pack(side="left")
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        nr=ctk.CTkFrame(card,fg_color="transparent"); nr.pack(fill="x",padx=16,pady=(12,6))
        ctk.CTkLabel(nr,text="출력 파일명:",font=FBOLD(13),text_color=TEXT_PRI).pack(side="left",padx=(0,8))
        self.imgpdf_name=ctk.CTkEntry(nr,placeholder_text="예) 과제제출본",font=F(12),width=220); self.imgpdf_name.pack(side="left")
        ctk.CTkLabel(nr,text=".pdf",font=F(12),text_color=TEXT_SUB).pack(side="left",padx=4)
        svr=ctk.CTkFrame(card,fg_color="transparent"); svr.pack(fill="x",padx=16,pady=6)
        self.imgpdf_save_var=ctk.StringVar(value="저장 폴더 선택")
        ctk.CTkLabel(svr,textvariable=self.imgpdf_save_var,font=F(12),text_color=TEXT_SUB,wraplength=400).pack(side="left")
        self._abtn(svr,"📂 선택",self._imgpdf_pick_save,width=80,height=36).pack(side="right")
        self._abtn(card,"📎 PDF 생성",lambda:threading.Thread(target=self._imgpdf_run,daemon=True).start(),height=42).pack(padx=16,pady=(6,14))
        self.imgpdf_log=self._logbox(parent); self.imgpdf_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
        self._imgpdf_save_folder=""
    def _imgpdf_pick_save(self):
        p=filedialog.askdirectory()
        if p: self._imgpdf_save_folder=p; self.imgpdf_save_var.set(p)
    def _imgpdf_run(self):
        if not HAS_PIL: self._log(self.imgpdf_log,"❌ Pillow 패키지가 필요합니다."); return
        if not self.imgpdf_files: self._log(self.imgpdf_log,"⚠️ 이미지 추가 먼저"); return
        try:
            imgs=[Image.open(f).convert("RGB") for f in self.imgpdf_files]
            for f in self.imgpdf_files: self._log(self.imgpdf_log,"📷  "+os.path.basename(f))
            name=self.imgpdf_name.get().strip() or "output"
            folder=self._imgpdf_save_folder or os.path.dirname(self.imgpdf_files[0])
            op=os.path.join(folder,name+".pdf")
            imgs[0].save(op,save_all=True,append_images=imgs[1:])
            self._log(self.imgpdf_log,"✅  "+name+".pdf ("+str(len(imgs))+"장)")
            self._toast(name+".pdf 생성 완료!","ok")
        except Exception as e: self._log(self.imgpdf_log,"❌ "+str(e))

    # ══════════════════════════════════════════
    # 🖼️  이미지 일괄 추출
    # ══════════════════════════════════════════
    def _build_imgext(self,parent):
        self._header(parent,"🖼️","이미지 일괄 추출","PDF·DOCX 문서들에서 이미지만 한번에 추출합니다")
        self._tip_box(parent,["① 문서 파일을 여러 개 추가하세요 (PDF·DOCX 지원).","② 추출 시작을 누르면 각 문서 옆에 _이미지들 폴더가 생성되고 이미지가 저장됩니다."])
        fw,btn_row,self.imgext_files,ie_add,ie_clear=self._file_list_widget(parent,"PDF 또는 DOCX 파일 추가")
        self._abtn(btn_row,"➕ 추가",lambda:ie_add([("문서","*.pdf *.docx")]),fg_color=SUCCESS,width=90,height=28).pack(side="left",padx=(0,6))
        self._abtn(btn_row,"🗑 초기화",ie_clear,fg_color=DANGER,width=80,height=28).pack(side="left")
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        self._abtn(card,"🖼️ 추출 시작",lambda:threading.Thread(target=self._imgext_run,daemon=True).start(),height=42).pack(padx=16,pady=14)
        self.imgext_log=self._logbox(parent); self.imgext_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
    def _imgext_run(self):
        if not self.imgext_files: self._log(self.imgext_log,"⚠️ 파일을 먼저 추가하세요."); return
        for fp in self.imgext_files:
            ext=Path(fp).suffix.lower(); out_dir=os.path.splitext(fp)[0]+"_이미지들"; os.makedirs(out_dir,exist_ok=True); count=0
            try:
                if ext==".pdf" and HAS_FITZ:
                    doc=fitz.open(fp)
                    for i,pg in enumerate(doc):
                        for j,img in enumerate(pg.get_images(full=True)):
                            bi=doc.extract_image(img[0])
                            with open(os.path.join(out_dir,"p"+str(i+1)+"_"+str(j+1)+"."+bi['ext']),"wb") as f: f.write(bi["image"])
                            count+=1
                elif ext==".docx" and HAS_DOCX:
                    doc=Document(fp)
                    for rel in doc.part.rels.values():
                        if "image" in rel.reltype:
                            ip=rel.target_part; ie=ip.content_type.split("/")[-1]
                            with open(os.path.join(out_dir,"img_"+str(count+1)+"."+ie),"wb") as f: f.write(ip.blob)
                            count+=1
                self._log(self.imgext_log,"✅  "+os.path.basename(fp)+": 이미지 "+str(count)+"개 → "+os.path.basename(out_dir))
            except Exception as e: self._log(self.imgext_log,"❌  "+os.path.basename(fp)+": "+str(e))
        self._toast("이미지 추출 완료!","ok")

    # ══════════════════════════════════════════
    # 💧  워터마크 삽입
    # ══════════════════════════════════════════
    def _build_watermark(self,parent):
        self._header(parent,"💧","워터마크 삽입","이미지에 텍스트 워터마크를 일괄 삽입합니다")
        self._tip_box(parent,["① 이미지 파일을 여러 개 추가하세요.","② 워터마크 텍스트를 입력하고 위치·투명도를 설정 후 삽입 실행을 누르세요.","③ 결과는 원본 폴더의 '_워터마크' 폴더에 저장됩니다."])
        fw,btn_row,self.wm_files,wm_add,wm_clear=self._file_list_widget(parent,"워터마크를 적용할 이미지 추가")
        self._abtn(btn_row,"➕ 추가",lambda:wm_add([("이미지","*.jpg *.jpeg *.png *.webp *.bmp")]),fg_color=SUCCESS,width=90,height=28).pack(side="left",padx=(0,6))
        self._abtn(btn_row,"🗑 초기화",wm_clear,fg_color=DANGER,width=80,height=28).pack(side="left")
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        inner=ctk.CTkFrame(card,fg_color="transparent"); inner.pack(fill="x",padx=16,pady=12); inner.columnconfigure(1,weight=1)
        ctk.CTkLabel(inner,text="워터마크 텍스트:",font=FBOLD(13),text_color=TEXT_PRI).grid(row=0,column=0,padx=(0,10),pady=7,sticky="w")
        self.wm_text=ctk.CTkEntry(inner,placeholder_text="예) 대외비  /  © 회사명",font=F(12)); self.wm_text.grid(row=0,column=1,columnspan=2,sticky="ew",pady=7)
        ctk.CTkLabel(inner,text="위치:",font=FBOLD(13),text_color=TEXT_PRI).grid(row=1,column=0,padx=(0,10),pady=7,sticky="w")
        self.wm_pos=ctk.CTkOptionMenu(inner,values=["우하단","우상단","중앙","좌하단","좌상단"],width=140,font=F(12)); self.wm_pos.grid(row=1,column=1,sticky="w",pady=7)
        ctk.CTkLabel(inner,text="투명도(0~255):",font=FBOLD(13),text_color=TEXT_PRI).grid(row=2,column=0,padx=(0,10),pady=7,sticky="w")
        self.wm_alpha=ctk.CTkEntry(inner,placeholder_text="예) 120",font=F(12),width=100); self.wm_alpha.grid(row=2,column=1,sticky="w",pady=7)
        self._abtn(card,"💧 워터마크 삽입",lambda:threading.Thread(target=self._wm_run,daemon=True).start(),height=42).pack(padx=16,pady=(0,14))
        self.wm_log=self._logbox(parent); self.wm_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
    def _wm_run(self):
        if not HAS_PIL: self._log(self.wm_log,"❌ Pillow 패키지가 필요합니다."); return
        if not self.wm_files: self._log(self.wm_log,"⚠️ 이미지 추가 먼저"); return
        text=self.wm_text.get().strip()
        if not text: self._log(self.wm_log,"⚠️ 워터마크 텍스트 입력하세요."); return
        pos=self.wm_pos.get()
        try: alpha=int(self.wm_alpha.get().strip() or "120")
        except Exception: alpha=120
        out_base=os.path.dirname(self.wm_files[0]); out_dir=os.path.join(out_base,"_워터마크"); os.makedirs(out_dir,exist_ok=True)
        for fp in self.wm_files:
            try:
                img=Image.open(fp).convert("RGBA"); overlay=Image.new("RGBA",img.size,(0,0,0,0))
                draw=ImageDraw.Draw(overlay)
                try: font=ImageFont.truetype("malgun.ttf",max(20,img.width//20))
                except Exception: font=ImageFont.load_default()
                bbox=draw.textbbox((0,0),text,font=font); tw,th=bbox[2]-bbox[0],bbox[3]-bbox[1]; margin=20
                pos_map={"우하단":(img.width-tw-margin,img.height-th-margin),"우상단":(img.width-tw-margin,margin),"중앙":((img.width-tw)//2,(img.height-th)//2),"좌하단":(margin,img.height-th-margin),"좌상단":(margin,margin)}
                x,y=pos_map.get(pos,(margin,margin)); draw.text((x,y),text,font=font,fill=(255,255,255,alpha))
                result=Image.alpha_composite(img,overlay).convert("RGB")
                op=os.path.join(out_dir,os.path.basename(fp)); result.save(op)
                self._log(self.wm_log,"✅  "+os.path.basename(fp))
            except Exception as e: self._log(self.wm_log,"❌  "+os.path.basename(fp)+": "+str(e))
        self._toast("워터마크 삽입 완료!","ok")

    # ══════════════════════════════════════════
    # 📸  이미지 OCR
    # ══════════════════════════════════════════
    def _build_ocr(self,parent):
        self._header(parent,"📸","이미지 OCR","사진·스캔 이미지 속 텍스트를 Gemini AI가 추출합니다")
        self._tip_box(parent,["① 이미지 파일을 여러 개 추가하세요.","② 출력 언어를 선택하고 텍스트 추출 시작을 누르세요.","③ 결과는 이미지와 같은 위치에 _OCR.txt 로 저장됩니다."])
        fw,btn_row,self.ocr_files,ocr_add,ocr_clear=self._file_list_widget(parent,"이미지 파일을 추가하세요")
        self._abtn(btn_row,"➕ 추가",lambda:ocr_add([("이미지","*.jpg *.jpeg *.png *.webp *.bmp *.tiff")]),fg_color=SUCCESS,width=90,height=28).pack(side="left",padx=(0,6))
        self._abtn(btn_row,"🗑 초기화",ocr_clear,fg_color=DANGER,width=80,height=28).pack(side="left")
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        lr=ctk.CTkFrame(card,fg_color="transparent"); lr.pack(fill="x",padx=16,pady=(12,8))
        ctk.CTkLabel(lr,text="출력 언어:",font=FBOLD(13),text_color=TEXT_PRI).pack(side="left",padx=(0,12))
        self.ocr_lang=ctk.StringVar(value="원본 그대로")
        ctk.CTkRadioButton(lr,text="원본 그대로",variable=self.ocr_lang,value="원본 그대로",font=F(13)).pack(side="left",padx=8)
        ctk.CTkRadioButton(lr,text="한국어로 번역",variable=self.ocr_lang,value="한국어",font=F(13)).pack(side="left",padx=8)
        self._abtn(card,"📸 텍스트 추출 시작",lambda:threading.Thread(target=self._ocr_run,daemon=True).start(),height=42).pack(padx=16,pady=(0,14))
        self.ocr_result=ctk.CTkTextbox(parent,font=F(13),fg_color=BG_DARK,text_color=TEXT_PRI,corner_radius=12,border_width=1,border_color=BORDER)
        self.ocr_result.pack(fill="both",expand=True,padx=28,pady=(0,20))
    def _ocr_run(self):
        if not HAS_GEMINI: self.ocr_result.insert("end","❌ Gemini 패키지가 없습니다.\n"); return
        if not GEMINI_API_KEY: self.ocr_result.insert("end","❌ API 키가 없습니다. ⚙️ 설정 탭에서 등록하세요.\n"); return
        if not self.ocr_files: self.ocr_result.insert("end","⚠️ 이미지 추가 먼저\n"); return
        self.ocr_result.delete("0.0","end")
        mime_map={"jpg":"image/jpeg","jpeg":"image/jpeg","png":"image/png","webp":"image/webp","bmp":"image/bmp","tiff":"image/tiff"}
        for fp in self.ocr_files:
            self.ocr_result.insert("end","━━━  "+os.path.basename(fp)+"  ━━━\n")
            with open(fp,"rb") as f: data=f.read()
            mime=mime_map.get(Path(fp).suffix.lower().lstrip("."),"image/jpeg")
            prompt="이 이미지에서 모든 텍스트를 추출하고 한국어로 번역해줘. 텍스트만 반환해." if self.ocr_lang.get()=="한국어" else "이 이미지에서 모든 텍스트를 그대로 추출해줘. 텍스트만 반환해."
            try:
                res=ai_client.models.generate_content(model=TARGET_MODEL,contents=[types.Part.from_bytes(data=data,mime_type=mime),prompt])
                txt=str(res.text).strip(); self.ocr_result.insert("end",txt+"\n\n")
                with open(os.path.splitext(fp)[0]+"_OCR.txt","w",encoding="utf-8") as f: f.write(txt)
                self.ocr_result.insert("end","💾  저장됨\n\n")
            except Exception as e: self.ocr_result.insert("end","❌ "+str(e)+"\n\n")
        self._toast("OCR 완료!","ok")

    # ══════════════════════════════════════════
    # 🤖  AI 문서 요약
    # ══════════════════════════════════════════
    def _build_summary(self,parent):
        self._header(parent,"🤖","AI 문서 요약","PDF·TXT·DOCX를 Gemini AI가 핵심만 요약합니다")
        self._tip_box(parent,["① 파일을 여러 개 추가하세요 (PDF·TXT·DOCX).","② 요약 스타일을 선택하세요: 핵심 요점 / 한 단락 / 시험 Q&A","③ 결과는 원본 파일과 같은 위치에 _AI요약.txt 로 저장됩니다."])
        fw,btn_row,self.summary_files,sum_add,sum_clear=self._file_list_widget(parent,"PDF·TXT·DOCX 파일 추가")
        self._abtn(btn_row,"➕ 추가",lambda:sum_add([("문서","*.pdf *.txt *.docx")]),fg_color=SUCCESS,width=90,height=28).pack(side="left",padx=(0,6))
        self._abtn(btn_row,"🗑 초기화",sum_clear,fg_color=DANGER,width=80,height=28).pack(side="left")
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        sr=ctk.CTkFrame(card,fg_color="transparent"); sr.pack(fill="x",padx=16,pady=(12,8))
        ctk.CTkLabel(sr,text="요약 스타일:",font=FBOLD(13),text_color=TEXT_PRI).pack(side="left",padx=(0,12))
        self.summary_style=ctk.StringVar(value="bullet")
        for txt,val in [("핵심 요점","bullet"),("한 단락","paragraph"),("시험 Q&A","qa")]:
            ctk.CTkRadioButton(sr,text=txt,variable=self.summary_style,value=val,font=F(13)).pack(side="left",padx=8)
        self._abtn(card,"🤖 요약 시작",lambda:threading.Thread(target=self._summary_run,daemon=True).start(),height=42).pack(padx=16,pady=(0,14))
        self.summary_result=ctk.CTkTextbox(parent,font=F(13),fg_color=BG_DARK,text_color=TEXT_PRI,corner_radius=12,border_width=1,border_color=BORDER)
        self.summary_result.pack(fill="both",expand=True,padx=28,pady=(0,20))
    def _summary_run(self):
        if not HAS_GEMINI: self.summary_result.insert("end","❌ Gemini 패키지가 없습니다.\n"); return
        if not GEMINI_API_KEY: self.summary_result.insert("end","❌ API 키가 없습니다. ⚙️ 설정 탭에서 등록하세요.\n"); return
        if not self.summary_files: self.summary_result.insert("end","⚠️ 파일 추가 먼저\n"); return
        self.summary_result.delete("0.0","end"); style=self.summary_style.get()
        for fp in self.summary_files:
            ext=Path(fp).suffix.lower(); content=""
            try:
                if ext==".txt":
                    with open(fp,"r",encoding="utf-8") as f: content=f.read()
                elif ext==".pdf" and HAS_FITZ: content="".join(p.get_text() for p in fitz.open(fp))
                elif ext==".docx" and HAS_DOCX: content="\n".join(p.text for p in Document(fp).paragraphs)
                else: self.summary_result.insert("end","⚠️ "+os.path.basename(fp)+" 지원 불가\n"); continue
                content=content[:8000]
                prompts={"bullet":"핵심 요점 5~10개로 불릿 요약해줘 (한국어):\n\n"+content,
                         "paragraph":"3~5문장 단락으로 요약해줘 (한국어):\n\n"+content,
                         "qa":"시험 Q&A 5개 만들어줘 (한국어):\n\n"+content}
                res=ai_client.models.generate_content(model=TARGET_MODEL,contents=prompts[style])
                txt=str(res.text).strip()
                self.summary_result.insert("end","━━━  "+os.path.basename(fp)+"  ━━━\n"+txt+"\n\n")
                with open(os.path.splitext(fp)[0]+"_AI요약.txt","w",encoding="utf-8") as f: f.write(txt)
                self.summary_result.insert("end","💾  저장됨\n\n")
            except Exception as e: self.summary_result.insert("end","❌ "+os.path.basename(fp)+": "+str(e)+"\n\n")
        self._toast("요약 완료!","ok")

    # ══════════════════════════════════════════
    # ✏️  AI 문서 초안
    # ══════════════════════════════════════════
    def _build_draft(self,parent):
        self._header(parent,"✏️","AI 문서 초안 작성","주제와 조건을 입력하면 Gemini AI가 초안을 자동 생성합니다")
        self._tip_box(parent,["① 문서 종류를 선택하고 주제/키워드를 입력하세요.","② 추가 조건(분량, 문체 등)을 입력하면 더 정확한 초안이 생성됩니다.","③ 초안 생성 후 결과를 확인하고 TXT 저장 버튼으로 저장하세요."])
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        inner=ctk.CTkFrame(card,fg_color="transparent"); inner.pack(fill="x",padx=16,pady=12); inner.columnconfigure(1,weight=1)
        ctk.CTkLabel(inner,text="문서 종류:",font=FBOLD(13),text_color=TEXT_PRI).grid(row=0,column=0,padx=(0,10),pady=8,sticky="w")
        self.draft_type=ctk.CTkOptionMenu(inner,values=["보고서","자기소개서","이메일","독후감","기획서","레포트","발표 스크립트","회의록","공지문","기타"],width=160,font=F(12))
        self.draft_type.grid(row=0,column=1,sticky="w",pady=8)
        ctk.CTkLabel(inner,text="주제/키워드:",font=FBOLD(13),text_color=TEXT_PRI).grid(row=1,column=0,padx=(0,10),pady=8,sticky="w")
        self.draft_topic=ctk.CTkEntry(inner,placeholder_text="예) 소방시설 설치 기준",font=F(12)); self.draft_topic.grid(row=1,column=1,sticky="ew",pady=8)
        ctk.CTkLabel(inner,text="추가 조건:",font=FBOLD(13),text_color=TEXT_PRI).grid(row=2,column=0,padx=(0,10),pady=8,sticky="w")
        self.draft_condition=ctk.CTkEntry(inner,placeholder_text="예) A4 1장, 공식 문체, 3단락",font=F(12)); self.draft_condition.grid(row=2,column=1,sticky="ew",pady=8)
        ctk.CTkLabel(inner,text="언어:",font=FBOLD(13),text_color=TEXT_PRI).grid(row=3,column=0,padx=(0,10),pady=8,sticky="w")
        self.draft_lang=ctk.StringVar(value="한국어")
        lf=ctk.CTkFrame(inner,fg_color="transparent"); lf.grid(row=3,column=1,sticky="w")
        ctk.CTkRadioButton(lf,text="한국어",variable=self.draft_lang,value="한국어",font=F(13)).pack(side="left",padx=(0,16))
        ctk.CTkRadioButton(lf,text="English",variable=self.draft_lang,value="English",font=F(13)).pack(side="left")
        br=ctk.CTkFrame(card,fg_color="transparent"); br.pack(fill="x",padx=16,pady=(0,12))
        self._abtn(br,"✏️ 초안 생성",lambda:threading.Thread(target=self._draft_run,daemon=True).start(),height=42,width=150).pack(side="left",padx=(0,10))
        self._abtn(br,"💾 TXT 저장",self._draft_save,fg_color=SUCCESS,height=42,width=120).pack(side="left",padx=(0,10))
        self._abtn(br,"🗑 초기화",self._draft_clear,fg_color=DANGER,height=42,width=100).pack(side="left")
        self.draft_result=ctk.CTkTextbox(parent,font=F(13),fg_color=BG_DARK,text_color=TEXT_PRI,corner_radius=12,border_width=1,border_color=BORDER)
        self.draft_result.pack(fill="both",expand=True,padx=28,pady=(0,20))
    def _draft_run(self):
        if not HAS_GEMINI: self.draft_result.insert("end","❌ Gemini 패키지가 없습니다.\n"); return
        if not GEMINI_API_KEY: self.draft_result.insert("end","❌ API 키가 없습니다. ⚙️ 설정 탭에서 등록하세요.\n"); return
        topic=self.draft_topic.get().strip()
        if not topic: messagebox.showwarning("입력 오류","주제를 입력하세요."); return
        self.draft_result.delete("0.0","end"); self.draft_result.insert("end","⏳  작성 중...\n\n")
        try:
            cond=self.draft_condition.get().strip()
            prompt=(self.draft_type.get()+" 초안을 작성해줘.\n주제: "+topic+"\n"+("조건: "+cond+"\n" if cond else "")+"언어: "+self.draft_lang.get()+"\n제목 포함, 완성도 있게.")
            res=ai_client.models.generate_content(model=TARGET_MODEL,contents=prompt)
            self.draft_result.delete("0.0","end"); self.draft_result.insert("0.0",str(res.text).strip())
        except Exception as e: self.draft_result.delete("0.0","end"); self.draft_result.insert("end","❌ "+str(e)+"\n")
    def _draft_save(self):
        content=self.draft_result.get("0.0","end").strip()
        if not content or "작성 중" in content: messagebox.showwarning("안내","먼저 초안을 생성하세요."); return
        p=filedialog.asksaveasfilename(defaultextension=".txt",filetypes=[("텍스트","*.txt")],initialfile=self.draft_type.get()+"_초안.txt")
        if p:
            with open(p,"w",encoding="utf-8") as f: f.write(content)
            self._toast("저장 완료: "+os.path.basename(p),"ok")
    def _draft_clear(self):
        self.draft_result.delete("0.0","end"); self.draft_topic.delete(0,"end"); self.draft_condition.delete(0,"end")

    # ══════════════════════════════════════════
    # ✂️  배경 제거
    # ══════════════════════════════════════════
    def _build_rembg(self,parent):
        self._header(parent,"✂️","배경 제거","이미지의 배경을 AI가 자동으로 제거합니다")
        self._tip_box(parent,["① 배경을 제거할 이미지를 여러 개 추가하세요 (JPG·PNG·WEBP).","② 배경 제거 시작을 누르면 AI가 자동으로 배경을 제거합니다.","③ 결과는 원본 폴더의 '_배경제거' 폴더에 PNG로 저장됩니다 (투명 배경)."])
        fw,btn_row,self.rembg_files,rb_add,rb_clear=self._file_list_widget(parent,"배경 제거할 이미지를 추가하세요")
        self._abtn(btn_row,"➕ 파일 추가",lambda:rb_add([("이미지","*.jpg *.jpeg *.png *.webp *.bmp")]),fg_color=SUCCESS,width=100,height=28).pack(side="left",padx=(0,6))
        self._abtn(btn_row,"🗑 초기화",rb_clear,fg_color=DANGER,width=80,height=28).pack(side="left")
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        self.rembg_progress=ctk.CTkProgressBar(card,progress_color=ACCENT)
        self.rembg_progress.set(0); self.rembg_progress.pack(fill="x",padx=16,pady=(14,6))
        self.btn_run_rembg=self._abtn(card,"✂️ 배경 제거 시작",lambda:threading.Thread(target=self._rembg_run,daemon=True).start(),height=42)
        self.btn_run_rembg.pack(padx=16,pady=(0,14))
        self.rembg_log=self._logbox(parent); self.rembg_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
    def _rembg_run(self):
        if not HAS_REMBG: self._log(self.rembg_log,"❌ rembg 패키지가 없습니다."); return
        if not HAS_PIL: self._log(self.rembg_log,"❌ Pillow 패키지가 없습니다."); return
        if not self.rembg_files: self._log(self.rembg_log,"⚠️ 이미지를 먼저 추가하세요."); return
        self._btn_working(self.btn_run_rembg,"⏳ 배경 제거 중...")
        try:
            out_dir=os.path.join(os.path.dirname(self.rembg_files[0]),"_배경제거"); os.makedirs(out_dir,exist_ok=True)
            for i,fp in enumerate(self.rembg_files):
                with open(fp,"rb") as f: inp=f.read()
                out_data=rembg_remove(inp)
                out_path=os.path.join(out_dir,Path(fp).stem+"_nobg.png")
                with open(out_path,"wb") as f: f.write(out_data)
                self._log(self.rembg_log,"✅  "+os.path.basename(fp)+"  →  "+os.path.basename(out_path))
                self.rembg_progress.set((i+1)/len(self.rembg_files))
            self._toast("배경 제거 완료! ("+str(len(self.rembg_files))+"개)","ok")
        except Exception as e: self._log(self.rembg_log,"❌ "+str(e))
        finally: self._btn_done(self.btn_run_rembg)

    # ══════════════════════════════════════════
    # 📚  참고문헌 자동 정리
    # ══════════════════════════════════════════
    def _build_citation(self,parent):
        self._header(parent,"📚","참고문헌 자동 정리","논문·책·웹사이트 정보를 입력하면 AI가 인용 형식으로 자동 정리합니다")
        self._tip_box(parent,["① 참고문헌으로 쓸 정보를 자유롭게 입력하세요 (제목, 저자, 연도, URL 등).","② 원하는 인용 형식을 선택하세요: APA·MLA·시카고·한국어 논문 형식.","③ 여러 항목은 줄바꿈으로 구분하세요."])
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        inner=ctk.CTkFrame(card,fg_color="transparent"); inner.pack(fill="x",padx=16,pady=12); inner.columnconfigure(1,weight=1)
        ctk.CTkLabel(inner,text="인용 형식:",font=FBOLD(13),text_color=TEXT_PRI).grid(row=0,column=0,padx=(0,10),pady=8,sticky="w")
        self.citation_style=ctk.CTkOptionMenu(inner,values=["APA 7판","MLA 9판","시카고 스타일","한국어 논문(KCI)"],width=180,font=F(12))
        self.citation_style.grid(row=0,column=1,sticky="w",pady=8)
        ctk.CTkLabel(inner,text="참고문헌 정보:",font=FBOLD(13),text_color=TEXT_PRI).grid(row=1,column=0,padx=(0,10),pady=8,sticky="nw")
        self.citation_input=ctk.CTkTextbox(inner,height=140,font=F(12),fg_color=BG_DARK,text_color=TEXT_PRI,corner_radius=8,border_width=1,border_color=BORDER)
        self.citation_input.grid(row=1,column=1,sticky="ew",pady=8)
        self.citation_input.insert("0.0","예시:\n김민수, 이영희 (2023). 인공지능과 교육의 미래. 한국교육학회지, 45(2), 123-145.\nSmith, J. (2022). Deep Learning. MIT Press.\nhttps://www.example.com (접속일: 2024.01.15)")
        self._abtn(card,"📚 참고문헌 정리",lambda:threading.Thread(target=self._citation_run,daemon=True).start(),height=42).pack(padx=16,pady=(0,14))
        self.citation_result=ctk.CTkTextbox(parent,font=F(13),fg_color=BG_DARK,text_color=TEXT_PRI,corner_radius=12,border_width=1,border_color=BORDER)
        self.citation_result.pack(fill="both",expand=True,padx=28,pady=(0,20))
    def _citation_run(self):
        if not HAS_GEMINI: self.citation_result.delete("0.0","end"); self.citation_result.insert("end","❌ Gemini 패키지가 없습니다.\n"); return
        if not GEMINI_API_KEY: self.citation_result.delete("0.0","end"); self.citation_result.insert("end","❌ API 키가 없습니다. ⚙️ 설정 탭에서 등록하세요.\n"); return
        raw=self.citation_input.get("0.0","end").strip()
        if not raw: self.citation_result.delete("0.0","end"); self.citation_result.insert("end","⚠️ 참고문헌 정보를 입력하세요.\n"); return
        style=self.citation_style.get()
        self.citation_result.delete("0.0","end"); self.citation_result.insert("end","⏳ AI가 참고문헌을 정리 중입니다...\n\n")
        try:
            prompt="아래 참고문헌 정보들을 "+style+" 형식에 맞게 정확하게 정리해줘.\n각 항목을 번호 없이 한 줄씩 올바른 형식으로 변환하고, 결과만 반환해.\n\n"+raw
            res=ai_client.models.generate_content(model=TARGET_MODEL,contents=prompt)
            result=str(res.text).strip()
            self.citation_result.delete("0.0","end"); self.citation_result.insert("0.0",result)
            self._toast("참고문헌 정리 완료!","ok")
        except Exception as e: self.citation_result.delete("0.0","end"); self.citation_result.insert("end","❌ "+str(e)+"\n")

    # ══════════════════════════════════════════
    # 📊  엑셀 자동화
    # ══════════════════════════════════════════
    def _build_excel(self,parent):
        self._header(parent,"📊","엑셀 자동화","엑셀·CSV 파일의 시트 합치기, 중복 제거, 포맷 변환을 자동화합니다")
        self._tip_box(parent,["① 처리할 엑셀(.xlsx) 또는 CSV 파일을 추가하세요.","② [시트 합치기]: 여러 엑셀 파일의 모든 시트를 하나의 파일로 합칩니다.","③ [중복 행 제거]: 완전히 동일한 행을 자동으로 제거합니다.","④ [CSV → XLSX]: CSV 파일을 엑셀 파일로 변환합니다."])
        fw,btn_row,self.excel_files,xl_add,xl_clear=self._file_list_widget(parent,"엑셀(.xlsx) 또는 CSV 파일을 추가하세요")
        self._abtn(btn_row,"➕ 파일 추가",lambda:xl_add([("엑셀/CSV","*.xlsx *.xls *.csv")]),fg_color=SUCCESS,width=100,height=28).pack(side="left",padx=(0,6))
        self._abtn(btn_row,"🗑 초기화",xl_clear,fg_color=DANGER,width=80,height=28).pack(side="left")
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        mode_row=ctk.CTkFrame(card,fg_color="transparent"); mode_row.pack(fill="x",padx=16,pady=(12,8))
        ctk.CTkLabel(mode_row,text="작업 선택:",font=FBOLD(13),text_color=TEXT_PRI).pack(side="left",padx=(0,12))
        self.excel_mode=ctk.StringVar(value="merge")
        for txt,val in [("시트 합치기","merge"),("중복 행 제거","dedup"),("CSV → XLSX","csv2xl")]:
            ctk.CTkRadioButton(mode_row,text=txt,variable=self.excel_mode,value=val,font=F(13)).pack(side="left",padx=8)
        self.btn_run_excel=self._abtn(card,"📊 실행",lambda:threading.Thread(target=self._excel_run,daemon=True).start(),height=42)
        self.btn_run_excel.pack(padx=16,pady=(0,14))
        self.excel_log=self._logbox(parent); self.excel_log.pack(fill="both",expand=True,padx=28,pady=(0,20))
    def _excel_run(self):
        if not HAS_XL: self._log(self.excel_log,"❌ openpyxl 패키지가 필요합니다."); return
        if not self.excel_files: self._log(self.excel_log,"⚠️ 파일을 먼저 추가하세요."); return
        self._btn_working(self.btn_run_excel,"⏳ 처리 중..."); mode=self.excel_mode.get()
        try:
            if mode=="merge":
                merged=openpyxl.Workbook(); merged.remove(merged.active)
                for fp in self.excel_files:
                    wb=openpyxl.load_workbook(fp)
                    for sname in wb.sheetnames:
                        title=(Path(fp).stem[:10]+"_"+sname)[:31]
                        ws_dst=merged.create_sheet(title=title)
                        for row in wb[sname].iter_rows(values_only=True): ws_dst.append(list(row))
                    self._log(self.excel_log,"✅  "+os.path.basename(fp)+" ("+str(len(wb.sheetnames))+"개 시트)")
                out=os.path.join(os.path.dirname(self.excel_files[0]),"_합본.xlsx"); merged.save(out)
                self._log(self.excel_log,"▶  저장: "+os.path.basename(out))
                self._toast("시트 합치기 완료! ("+str(len(merged.sheetnames))+"개 시트)","ok")
            elif mode=="dedup":
                import csv as _csv
                for fp in self.excel_files:
                    ext=Path(fp).suffix.lower()
                    if ext==".csv":
                        with open(fp,"r",encoding="utf-8-sig") as f: rows=list(_csv.reader(f))
                        header=rows[0] if rows else []; seen=set(); unique=[header]
                        for row in rows[1:]:
                            key=tuple(row)
                            if key not in seen: seen.add(key); unique.append(row)
                        removed=len(rows)-len(unique)
                        out=os.path.splitext(fp)[0]+"_중복제거.csv"
                        with open(out,"w",encoding="utf-8-sig",newline="") as f: _csv.writer(f).writerows(unique)
                    else:
                        wb=openpyxl.load_workbook(fp); ws=wb.active
                        rows=list(ws.iter_rows(values_only=True)); header=rows[0] if rows else ()
                        seen=set(); unique=[header]
                        for row in rows[1:]:
                            if row not in seen: seen.add(row); unique.append(row)
                        removed=len(rows)-len(unique)
                        wb_new=openpyxl.Workbook(); ws_new=wb_new.active
                        for row in unique: ws_new.append(list(row))
                        out=os.path.splitext(fp)[0]+"_중복제거.xlsx"; wb_new.save(out)
                    self._log(self.excel_log,"✅  "+os.path.basename(fp)+" → 중복 "+str(removed)+"행 제거")
                self._toast("중복 행 제거 완료!","ok")
            elif mode=="csv2xl":
                import csv as _csv
                for fp in self.excel_files:
                    if Path(fp).suffix.lower()!=".csv": self._log(self.excel_log,"⚠️  "+os.path.basename(fp)+": CSV 파일이 아닙니다."); continue
                    wb=openpyxl.Workbook(); ws=wb.active
                    with open(fp,"r",encoding="utf-8-sig") as f:
                        for row in _csv.reader(f): ws.append(row)
                    out=os.path.splitext(fp)[0]+".xlsx"; wb.save(out)
                    self._log(self.excel_log,"✅  "+os.path.basename(fp)+"  →  "+Path(out).name)
                self._toast("CSV → XLSX 변환 완료!","ok")
        except Exception as e: self._log(self.excel_log,"❌ "+str(e))
        finally: self._btn_done(self.btn_run_excel)

    # ══════════════════════════════════════════
    # 📅  과제 마감 트래커
    # ══════════════════════════════════════════
    def _build_tracker(self,parent):
        self._header(parent,"📅","과제 마감 트래커","과목별 과제 등록 및 D-day 관리")
        self._tip_box(parent,["① 과목명·과제명·마감일(YYYY-MM-DD)을 입력하고 추가 버튼을 누르세요.","② 목록에서 과제를 클릭하면 선택됩니다. 완료 처리 또는 삭제 버튼으로 관리하세요.","③ D-day 3일 이하 주황색, 당일·초과 빨간색으로 표시됩니다."])
        top=ctk.CTkFrame(parent,fg_color="transparent"); top.pack(fill="x",padx=28,pady=(0,8)); top.columnconfigure((1,3),weight=1)
        for label,row,col,attr,ph in [("과목명:",0,0,"tracker_subject","예) 소방학개론"),("과제명:",0,2,"tracker_title","예) 3장 요약"),("마감일:",1,0,"tracker_due",datetime.date.today().strftime("%Y-%m-%d")),("메모:",1,2,"tracker_memo","선택 입력")]:
            ctk.CTkLabel(top,text=label,font=FBOLD(13),text_color=TEXT_PRI).grid(row=row,column=col,padx=(0,6),pady=6,sticky="w")
            e=ctk.CTkEntry(top,placeholder_text=ph,font=F(12)); e.grid(row=row,column=col+1,sticky="ew",padx=(0,12 if col==0 else 0),pady=6)
            setattr(self,attr,e)
        br=ctk.CTkFrame(parent,fg_color="transparent"); br.pack(fill="x",padx=28,pady=(0,10))
        for txt,color,cmd in [("➕ 추가",SUCCESS,self._tracker_add),("✅ 완료",WARNING,self._tracker_done),("🗑️ 삭제",DANGER,self._tracker_delete)]:
            self._abtn(br,txt,cmd,fg_color=color,width=110,height=38).pack(side="left",padx=(0,8))
        self._abtn(br,"🔄 새로고침",self._tracker_refresh,width=110,height=38).pack(side="right")
        self.tracker_scroll=ctk.CTkScrollableFrame(parent,fg_color=BG_DARK,corner_radius=12,border_width=1,border_color=BORDER)
        self.tracker_scroll.pack(fill="both",expand=True,padx=28,pady=(0,20))
        self.tracker_scroll.columnconfigure((0,1,2,3,4),weight=1)
        for col,h in enumerate(["과목","과제명","마감일","D-day","상태"]):
            ctk.CTkLabel(self.tracker_scroll,text=h,font=FBOLD(12),text_color=ACCENT).grid(row=0,column=col,padx=8,pady=(8,4),sticky="w")
        self._tracker_selected=None; self._tracker_refresh()
    def _tracker_add(self):
        subj=self.tracker_subject.get().strip(); title=self.tracker_title.get().strip()
        due=self.tracker_due.get().strip(); memo=self.tracker_memo.get().strip()
        if not subj or not title or not due: messagebox.showwarning("입력 오류","과목명, 과제명, 마감일은 필수입니다."); return
        try: datetime.datetime.strptime(due,"%Y-%m-%d")
        except Exception: messagebox.showwarning("형식 오류","YYYY-MM-DD 형식으로 입력하세요."); return
        self.tasks.append({"subject":subj,"title":title,"due":due,"memo":memo,"done":False}); save_tasks(self.tasks)
        for e in [self.tracker_subject,self.tracker_title,self.tracker_due,self.tracker_memo]: e.delete(0,"end")
        self._tracker_refresh()
    def _tracker_done(self):
        if self._tracker_selected is None: messagebox.showinfo("안내","과제를 먼저 클릭해서 선택하세요."); return
        self.tasks[self._tracker_selected]["done"]=True; save_tasks(self.tasks); self._tracker_refresh()
    def _tracker_delete(self):
        if self._tracker_selected is None: messagebox.showinfo("안내","과제를 먼저 클릭해서 선택하세요."); return
        self.tasks.pop(self._tracker_selected); save_tasks(self.tasks); self._tracker_selected=None; self._tracker_refresh()
    def _tracker_refresh(self):
        self._tracker_selected=None
        children=self.tracker_scroll.winfo_children()
        for w in children:
            try:
                if int(w.grid_info().get("row",0))>0: w.destroy()
            except Exception: pass
        today=datetime.date.today()
        for i,task in enumerate(sorted(self.tasks,key=lambda t:(t.get("done",False),t["due"]))):
            ridx=self.tasks.index(task); delta=(datetime.datetime.strptime(task["due"],"%Y-%m-%d").date()-today).days; done=task.get("done",False)
            if done: dt,dc="완료 ✅",SUCCESS
            elif delta<0: dt,dc="D+"+str(abs(delta))+" 초과",DANGER
            elif delta==0: dt,dc="D-day ❗",DANGER
            elif delta<=3: dt,dc="D-"+str(delta)+" ⚠️",WARNING
            else: dt,dc="D-"+str(delta),"white"
            rbg="#1A1D2E" if i%2==0 else BG_DARK
            for col,val in enumerate([task["subject"],task["title"],task["due"],dt,"완료" if done else "진행중"]):
                color=dc if col==3 else (TEXT_SUB if done else TEXT_PRI)
                lbl=ctk.CTkLabel(self.tracker_scroll,text=val,font=FBOLD(12) if col==3 else F(12),text_color=color,fg_color=rbg,corner_radius=0)
                lbl.grid(row=i+1,column=col,padx=8,pady=3,sticky="ew")
                lbl.bind("<Button-1>",lambda e,idx=ridx:self._tracker_select(idx))
    def _tracker_select(self,idx):
        self._tracker_selected=idx; task=self.tasks[idx]
        messagebox.showinfo("선택된 과제","과목: "+task['subject']+"\n과제: "+task['title']+"\n마감: "+task['due']+"\n메모: "+task.get('memo','없음')+"\n\n완료 처리 또는 삭제 버튼을 눌러 관리하세요.")

    # ══════════════════════════════════════════
    # ⚙️  설정
    # ══════════════════════════════════════════
    def _build_settings(self,parent):
        self._header(parent,"⚙️","API 키 설정","DeepL · Gemini API 키를 안전하게 로컬에 저장합니다")
        self._tip_box(parent,["① 아래에 API 키를 입력하고 '저장' 버튼을 누르세요.","② 키는 AES-256-GCM + PBKDF2(480,000회) 로 암호화되어 이 PC에만 저장됩니다.","③ 암호화된 파일을 다른 PC로 복사해도 복호화가 불가능합니다.","④ DeepL 키: https://www.deepl.com/pro-api  /  Gemini 키: https://aistudio.google.com/apikey"])
        card=self._card(parent); card.pack(fill="x",padx=28,pady=(0,10))
        inner=ctk.CTkFrame(card,fg_color="transparent"); inner.pack(fill="x",padx=16,pady=16); inner.columnconfigure(1,weight=1)
        ctk.CTkLabel(inner,text="DeepL API 키:",font=FBOLD(13),text_color=TEXT_PRI).grid(row=0,column=0,padx=(0,12),pady=8,sticky="w")
        self.settings_deepl=ctk.CTkEntry(inner,placeholder_text="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx:fx",font=F(12),show="●")
        self.settings_deepl.grid(row=0,column=1,sticky="ew",pady=8)
        show_deepl_var=ctk.BooleanVar(value=False)
        def _toggle_deepl(): self.settings_deepl.configure(show="" if show_deepl_var.get() else "●")
        ctk.CTkCheckBox(inner,text="표시",variable=show_deepl_var,command=_toggle_deepl,font=F(11),width=60).grid(row=0,column=2,padx=(8,0))
        ctk.CTkLabel(inner,text="Gemini API 키:",font=FBOLD(13),text_color=TEXT_PRI).grid(row=1,column=0,padx=(0,12),pady=8,sticky="w")
        self.settings_gemini=ctk.CTkEntry(inner,placeholder_text="AIzaSy...",font=F(12),show="●")
        self.settings_gemini.grid(row=1,column=1,sticky="ew",pady=8)
        show_gemini_var=ctk.BooleanVar(value=False)
        def _toggle_gemini(): self.settings_gemini.configure(show="" if show_gemini_var.get() else "●")
        ctk.CTkCheckBox(inner,text="표시",variable=show_gemini_var,command=_toggle_gemini,font=F(11),width=60).grid(row=1,column=2,padx=(8,0))
        status_frame=ctk.CTkFrame(card,fg_color=BG_TIP,corner_radius=10); status_frame.pack(fill="x",padx=16,pady=(0,8))
        self._settings_status=ctk.CTkLabel(status_frame,text="",font=F(12),text_color=TEXT_SUB)
        self._settings_status.pack(padx=14,pady=8,anchor="w"); self._refresh_settings_status()
        btn_row=ctk.CTkFrame(card,fg_color="transparent"); btn_row.pack(fill="x",padx=16,pady=(0,14))
        self._abtn(btn_row,"💾 저장",self._settings_save,fg_color=SUCCESS,height=42,width=120).pack(side="left",padx=(0,10))
        self._abtn(btn_row,"🗑️ 키 초기화",self._settings_clear,fg_color=DANGER,height=42,width=130).pack(side="left")
        if DEEPL_API_KEY: self.settings_deepl.insert(0,DEEPL_API_KEY)
        if GEMINI_API_KEY: self.settings_gemini.insert(0,GEMINI_API_KEY)
    def _refresh_settings_status(self):
        deepl_ok="✅ 저장됨" if DEEPL_API_KEY else "❌ 없음"
        gemini_ok="✅ 저장됨" if GEMINI_API_KEY else "❌ 없음"
        crypto_ok="🔒 AES-256-GCM" if HAS_CRYPTO else "⚠️ cryptography 없음"
        self._settings_status.configure(text="DeepL: "+deepl_ok+"   |   Gemini: "+gemini_ok+"\n암호화: "+crypto_ok+"   |   저장 위치: "+CONFIG_FILE)
    def _settings_save(self):
        global DEEPL_API_KEY,GEMINI_API_KEY,ai_client
        if not HAS_CRYPTO:
            messagebox.showerror("보안 패키지 없음","API 키 저장을 위해 cryptography 패키지가 필요합니다.\n설치 방법은 공식 문서를 참고하세요."); return
        new_deepl=self.settings_deepl.get().strip(); new_gemini=self.settings_gemini.get().strip()
        if not new_deepl and not new_gemini: messagebox.showwarning("입력 오류","최소 하나의 API 키를 입력하세요."); return
        cfg={}
        if new_deepl: cfg["deepl_key"]=new_deepl
        if new_gemini: cfg["gemini_key"]=new_gemini
        save_config(cfg)
        DEEPL_API_KEY=new_deepl or DEEPL_API_KEY; GEMINI_API_KEY=new_gemini or GEMINI_API_KEY
        if new_gemini: _init_gemini()
        self._refresh_settings_status(); self._toast("AES-256 암호화로 안전하게 저장되었습니다!","ok")
    def _settings_clear(self):
        if not messagebox.askyesno("확인","저장된 API 키를 모두 삭제하시겠습니까?"): return
        global DEEPL_API_KEY,GEMINI_API_KEY,ai_client
        try:
            if os.path.exists(CONFIG_FILE): os.remove(CONFIG_FILE)
        except Exception: pass
        DEEPL_API_KEY=GEMINI_API_KEY=""; ai_client=None
        self.settings_deepl.delete(0,"end"); self.settings_gemini.delete(0,"end")
        self._refresh_settings_status(); self._toast("API 키가 초기화되었습니다.","warn")


if __name__=="__main__":
    app=MinjuToolkitApp()
    app.mainloop()
