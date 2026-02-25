import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu
from PIL import Image, ImageOps
import os
from tkinterdnd2 import DND_FILES, TkinterDnD
import sys
import pillow_heif

# HEIFフォーマットをPillowに登録
pillow_heif.register_heif_opener()

# ==============================
# リソースパス取得関数 (exe化対応)
# ==============================
def resource_path(relative_path):
    """実行時（exe）でも通常時（.pyw）でも正しいファイルパスを取得する"""
    try:
        # PyInstallerで展開された時の一時フォルダパス
        base_path = sys._MEIPASS
    except Exception:
        # 通常のPythonスクリプトとして実行した時のパス
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class ImageQualityApp:
    # バージョン定数
    VERSION = "v1.2.4"
    
    # 変更履歴
    CHANGELOG = f"""
    【変更履歴】
    {VERSION}
     - 任意のフォルダ保存キャンセル時に完了メッセージを非表示にするよう修正
     - 「画像を選択」の参照ダイアログの初期表示を「PC（Cドライブ直下）」に変更
     - 品質の選択肢に「元の画質 (変更なし)」を追加
     - 保存するファイル形式（元の形式、JPEG、PNG、WebP、HEIF）を選択できる機能を追加
     - HEIF/HEIC形式の読み込みおよび保存に対応（pillow-heif追加）

    v1.2.3
     - ヘルプメニューに「使い方 (Readme)」を追加

    v1.2.2
     - 完了通知ウィンドウのOKボタンを削除（自動クローズのみに変更）

    v1.2.1
     - 処理完了時のメッセージを3秒後に自動で閉じるように変更

    v1.2.0
     - タイトル名称を ImageResizeMiya に変更
     - デザインを青ベースのモダンなテーマに変更
     - バージョン情報と更新履歴の表示機能を追加

    v1.1.0
     - 複数ファイルのドラッグ＆ドロップと一括処理に対応
     - 保存先設定の柔軟化（フォルダ指定による一括保存）

    v1.0.0
     - 画像品質変更ツールとして新規リリース
    """

    # Readme情報
    README_TEXT = """【ImageResizeMiya 使い方】

■ 概要
画像の品質（圧縮率）やファイル形式を一括で変更するためのツールです。

■ 基本的な使い方
1. 画像の選択
   「参照」ボタン、またはウィンドウ内へのドラッグ＆ドロップで画像を選択します。複数選択も可能です。
   ※参照ダイアログは常に「PC（Cドライブ直下）」を初期位置として開きます。

2. 品質の選択
   リストから希望の画質を選びます（初期設定は「低画質(50)」）。
   「元の画質 (変更なし)」を選ぶと、極力元の画質を維持したまま保存・変換します。
   ※PNG形式は元々圧縮効果が薄いため、画質設定の影響を受けにくい場合があります。

3. 保存形式の選択
   変換先のファイル形式を選びます。「元の形式」を選ぶと、元画像と同じ拡張子で保存されます。

4. 保存先の選択
   - 「同じフォルダ」: 元画像と同じ場所に「_q50」（元の画質の場合は「_org」）等の名前で保存されます。
   - 「任意のフォルダ」: 指定したフォルダにまとめて保存されます。

5. 実行
   「品質変更を実行」ボタンを押すと処理が始まります。
   完了メッセージは3秒後に自動で閉じます。"""

    def __init__(self, root):
        self.root = root
        # タイトル
        self.root.title(f"ImageResizeMiya {self.VERSION}")
        self.root.geometry("450x460") 

        # --- カラーパレット定義 ---
        self.colors = {
            "bg_window": "#F0F8FF",      # 全体の背景（AliceBlue）
            "bg_card": "#FFFFFF",        # 各セクションの背景（白）
            "accent": "#0078D7",         # アクセントカラー（鮮やかな青）
            "accent_hover": "#005A9E",   # ボタンホバー色
            "text_main": "#333333",      # 主な文字色
            "text_sub": "#666666",       # 補足文字色
            "text_white": "#FFFFFF"      # 白文字
        }
        
        # フォント設定
        self.font_bold = ("Yu Gothic UI", 9, "bold")
        self.font_norm = ("Yu Gothic UI", 9)
        self.font_small = ("Yu Gothic UI", 8)

        # 全体の背景色設定
        self.root.configure(bg=self.colors["bg_window"])

        # 変数
        self.file_path = tk.StringVar()
        self.save_mode = tk.StringVar(value="same") # 初期値: 同じフォルダ
        self.selected_files = [] 
        
        # 品質設定 (キー: 表示名, 値: Quality値。'keep'は特別な値として処理)
        self.quality_options = {
            "元の画質 (変更なし)": "keep",
            "最高画質 (Quality: 95)": 95,
            "高画質 (Quality: 85)": 85,
            "中画質 (Quality: 65)": 65,
            "低画質 (Quality: 50)": 50,
            "最低画質 (Quality: 30)": 30,
            "超低画質 (Quality: 10)": 10,
            "限界画質 (Quality: 1)": 1
        }

        # 保存形式設定
        self.format_options = {
            "元の形式のまま": None,
            "JPEG (.jpg) に変換": ".jpg",
            "PNG (.png) に変換": ".png",
            "WebP (.webp) に変換": ".webp",
            "HEIF (.heic) に変換": ".heic"
        }

        # メニューバー作成
        self._create_menu()
        
        # ウィジェット作成
        self._create_widgets()

    def _create_menu(self):
        menubar = Menu(self.root)
        
        # ヘルプメニュー
        help_menu = Menu(menubar, tearoff=0)
        # Readmeをメニューに追加
        help_menu.add_command(label="使い方 (Readme)", command=self.show_readme)
        help_menu.add_separator() # 区切り線
        help_menu.add_command(label="バージョン情報", command=self.show_version_info)
        
        menubar.add_cascade(label="ヘルプ", menu=help_menu)
        self.root.config(menu=menubar)

    def show_readme(self):
        messagebox.showinfo("使い方 (Readme)", self.README_TEXT)

    def show_version_info(self):
        info_text = f"ImageResizeMiya\nVersion: {self.VERSION}\n\n{self.CHANGELOG}"
        messagebox.showinfo("バージョン情報", info_text)

    def _create_widgets(self):
        # --- 1. ファイル選択エリア ---
        frame_file = tk.LabelFrame(
            self.root, 
            text=" 1. 画像を選択 (複数可・D&D可) ",
            bg=self.colors["bg_card"],
            fg=self.colors["accent"],
            font=self.font_bold,
            bd=1, relief="solid"
        )
        frame_file.pack(fill="x", padx=10, pady=(10, 5))

        inner_file = tk.Frame(frame_file, bg=self.colors["bg_card"])
        inner_file.pack(fill="x", padx=10, pady=10)

        self.entry_file = tk.Entry(inner_file, textvariable=self.file_path, font=self.font_norm, bd=1, relief="solid")
        self.entry_file.pack(side="left", fill="x", expand=True, padx=(0, 5), ipady=3)
        
        # DND設定
        self.entry_file.drop_target_register(DND_FILES)
        self.entry_file.dnd_bind('<<Drop>>', self.on_drop)

        btn_browse = tk.Button(
            inner_file, text="参照", command=self.select_file,
            bg="#E1E1E1", fg=self.colors["text_main"], font=self.font_norm,
            relief="flat", padx=10
        )
        btn_browse.pack(side="left")

        # --- 2. 品質選択エリア ---
        frame_quality = tk.LabelFrame(
            self.root, 
            text=" 2. 品質(圧縮率)を選択 ", 
            bg=self.colors["bg_card"],
            fg=self.colors["accent"],
            font=self.font_bold,
            bd=1, relief="solid"
        )
        frame_quality.pack(fill="x", padx=10, pady=5)

        inner_quality = tk.Frame(frame_quality, bg=self.colors["bg_card"])
        inner_quality.pack(fill="x", padx=10, pady=5)

        self.combo_quality = ttk.Combobox(inner_quality, values=list(self.quality_options.keys()), state="readonly", font=self.font_norm)
        # デフォルト: 低画質(50)
        self.combo_quality.current(4) 
        self.combo_quality.pack(fill="x")

        # --- 3. 保存形式エリア ---
        frame_format = tk.LabelFrame(
            self.root, 
            text=" 3. 保存形式を選択 ", 
            bg=self.colors["bg_card"],
            fg=self.colors["accent"],
            font=self.font_bold,
            bd=1, relief="solid"
        )
        frame_format.pack(fill="x", padx=10, pady=5)

        inner_format = tk.Frame(frame_format, bg=self.colors["bg_card"])
        inner_format.pack(fill="x", padx=10, pady=5)

        self.combo_format = ttk.Combobox(inner_format, values=list(self.format_options.keys()), state="readonly", font=self.font_norm)
        self.combo_format.current(0) # デフォルト: 元の形式のまま
        self.combo_format.pack(fill="x")

        # --- 4. 保存先設定エリア ---
        frame_save = tk.LabelFrame(
            self.root, 
            text=" 4. 保存先 ", 
            bg=self.colors["bg_card"],
            fg=self.colors["accent"],
            font=self.font_bold,
            bd=1, relief="solid"
        )
        frame_save.pack(fill="x", padx=10, pady=5)

        inner_save = tk.Frame(frame_save, bg=self.colors["bg_card"])
        inner_save.pack(fill="x", padx=10, pady=10)

        # ラジオボタンのスタイル
        rb_opts = {
            'bg': self.colors["bg_card"], 
            'fg': self.colors["text_main"], 
            'font': self.font_norm, 
            'activebackground': self.colors["bg_card"],
            'selectcolor': self.colors["bg_card"]
        }

        tk.Radiobutton(inner_save, text="同じフォルダに保存 (ファイル名_quality)", variable=self.save_mode, value="same", **rb_opts).pack(anchor="w")
        tk.Radiobutton(inner_save, text="任意のフォルダを指定 (保存時に選択)", variable=self.save_mode, value="arbitrary", **rb_opts).pack(anchor="w", pady=(5,0))

        # --- 5. 実行ボタン ---
        btn_run = tk.Button(
            self.root, 
            text="品質変更を実行", 
            command=self.change_quality, 
            bg=self.colors["accent"], 
            fg=self.colors["text_white"], 
            font=("Yu Gothic UI", 11, "bold"),
            relief="flat",
            activebackground=self.colors["accent_hover"],
            activeforeground=self.colors["text_white"],
            cursor="hand2"
        )
        btn_run.pack(fill="x", padx=10, pady=10, ipady=5)

    def on_drop(self, event):
        try:
            files = self.root.tk.splitlist(event.data)
            self.update_file_list(files)
        except Exception:
            self.file_path.set(event.data)
            self.selected_files = [event.data]

    def select_file(self):
        filetypes = [("画像ファイル", "*.jpg;*.jpeg;*.png;*.webp;*.heic;*.heif"), ("すべてのファイル", "*.*")]
        
        # Tkinterの仕様上、仮想フォルダ(PC)は弾かれるため実在する最上位階層(C:\)を指定
        initial_dir = "C:\\" if sys.platform == "win32" else "/"
        
        files = filedialog.askopenfilenames(filetypes=filetypes, initialdir=initial_dir)
        if files:
            self.update_file_list(files)

    def update_file_list(self, files):
        self.selected_files = files
        display_str = ";".join(files)
        self.file_path.set(display_str)

    def show_auto_close_message(self, title, message, duration=3000):
        popup = tk.Toplevel(self.root)
        popup.title(title)
        
        width, height = 300, 120 
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (width // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (height // 2)
        popup.geometry(f"{width}x{height}+{x}+{y}")
        
        popup.configure(bg=self.colors["bg_window"])
        
        lbl = tk.Label(
            popup, 
            text=message, 
            bg=self.colors["bg_window"], 
            fg=self.colors["text_main"], 
            font=self.font_norm, 
            justify="center",
            wraplength=260
        )
        lbl.pack(expand=True, fill="both", padx=20, pady=20)
        
        popup.transient(self.root)
        popup.grab_set()
        self.root.wait_visibility(popup)
        
        popup.after(duration, popup.destroy)

    def change_quality(self):
        raw_str = self.file_path.get()
        if not raw_str:
            messagebox.showerror("エラー", "画像ファイルを選択してください。")
            return
        
        target_files = [f.strip() for f in raw_str.split(';') if f.strip()]
        if not target_files:
            return

        quality_val = self.quality_options[self.combo_quality.get()]
        selected_format_ext = self.format_options[self.combo_format.get()]
        
        target_dir = None
        # 複数ファイル選択時に「任意のフォルダ」がキャンセルされたらそのまま終了
        if self.save_mode.get() == "arbitrary" and len(target_files) > 1:
            target_dir = filedialog.askdirectory(title="保存先フォルダを選択してください")
            if not target_dir: return

        success_count = 0
        error_count = 0
        
        self.root.config(cursor="watch")
        self.root.update()

        try:
            for src_path in target_files:
                if not os.path.exists(src_path):
                    continue

                try:
                    with Image.open(src_path) as img:
                        # EXIF補正
                        img = ImageOps.exif_transpose(img)

                        file_dir = os.path.dirname(src_path)
                        file_name, original_ext = os.path.splitext(os.path.basename(src_path))
                        
                        # 変換先拡張子の決定（指定がなければ元の拡張子）
                        target_ext = selected_format_ext if selected_format_ext else original_ext
                        
                        # 接尾辞の決定（元の画質の場合は _org とする）
                        suffix = "org" if quality_val == "keep" else f"q{quality_val}"
                        default_name = f"{file_name}_{suffix}{target_ext}"
                        save_path = ""

                        if self.save_mode.get() == "same":
                            save_path = os.path.join(file_dir, default_name)
                        elif self.save_mode.get() == "arbitrary":
                            if len(target_files) > 1:
                                save_path = os.path.join(target_dir, default_name)
                            else:
                                # 単一ファイル選択時に「任意のフォルダ」がキャンセルされた場合はスキップ
                                save_path = filedialog.asksaveasfilename(
                                    initialdir=file_dir,
                                    initialfile=default_name,
                                    defaultextension=target_ext,
                                    filetypes=[("Target Type", f"*{target_ext}")] if selected_format_ext else [("Original Type", original_ext), ("JPEG", "*.jpg"), ("PNG", "*.png"), ("WebP", "*.webp"), ("HEIF", "*.heic")]
                                )
                                if not save_path: continue

                        if save_path:
                            # JPEG保存時のアルファチャンネルエラー対策
                            if save_path.lower().endswith((".jpg", ".jpeg")) and img.mode in ("RGBA", "P"):
                                img = img.convert("RGB")
                            
                            # 「元の画質」を選んだ場合の処理
                            if quality_val == "keep":
                                # 元もJPEGで保存先もJPEGなら品質情報を極力維持 (keep)
                                if original_ext.lower() in (".jpg", ".jpeg") and save_path.lower().endswith((".jpg", ".jpeg")):
                                    img.save(save_path, quality="keep", optimize=True)
                                elif save_path.lower().endswith((".heic", ".heif")):
                                    img.save(save_path, format="HEIF", quality=100)
                                else:
                                    img.save(save_path, quality=100, optimize=True)
                                    
                            # 「圧縮率」を選んだ場合の通常の処理
                            else:
                                if save_path.lower().endswith((".heic", ".heif")):
                                    img.save(save_path, format="HEIF", quality=quality_val)
                                else:
                                    img.save(save_path, quality=quality_val, optimize=True)
                                
                            success_count += 1

                except Exception as e:
                    print(f"Error processing {src_path}: {e}")
                    error_count += 1

            self.root.config(cursor="")
            
            # --- キャンセル等で1件も処理されなかった場合はメッセージを出さずに終了 ---
            if success_count == 0 and error_count == 0:
                return
            
            msg = f"処理が完了しました。\n成功: {success_count}件"
            if error_count > 0:
                msg += f"\n失敗: {error_count}件"
            
            self.show_auto_close_message("完了", msg, duration=3000)

        except Exception as e:
            self.root.config(cursor="")
            messagebox.showerror("エラー", f"予期せぬエラーが発生しました:\n{e}")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ImageQualityApp(root)
    root.mainloop()