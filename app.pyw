import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu
from PIL import Image, ImageOps
import os
from tkinterdnd2 import DND_FILES, TkinterDnD

class ImageQualityApp:
    # バージョン定数
    VERSION = "v1.2.3"
    
    # 変更履歴
    CHANGELOG = f"""
    【変更履歴】
    {VERSION}
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

    v1.0.5
     - 画質設定に「超低画質(10)」「限界画質(1)」を追加
     - デフォルト画質を「低画質(50)」に変更

    v1.0.2
     - 写真の向き（EXIF回転情報）を自動補正する機能を追加

    v1.0.0
     - 画像品質変更ツールとして新規リリース
    """

    # Readme情報
    README_TEXT = """【ImageResizeMiya 使い方】

■ 概要
画像の品質（圧縮率）を一括で変更するためのツールです。
JPEGやWebPなどのファイルサイズを削減したい場合に便利です。

■ 基本的な使い方
1. 画像の選択
   「参照」ボタン、またはウィンドウ内へのドラッグ＆ドロップで画像を選択します。複数選択も可能です。

2. 品質の選択
   リストから希望の画質を選びます（初期設定は「低画質(50)」）。

3. 保存先の選択
   - 「同じフォルダ」: 元画像と同じ場所に「_q50」等の名前で保存されます。
   - 「任意のフォルダ」: 指定したフォルダにまとめて保存されます。

4. 実行
   「品質変更を実行」ボタンを押すと処理が始まります。
   完了メッセージは3秒後に自動で閉じます。

■ 注意事項
- PNG形式も処理可能ですが、品質設定による圧縮効果はJPEG/WebPほど高くありません。
- 透明度(アルファチャンネル)を含む画像をJPEGで保存する場合、自動的に背景が黒(または白)に変換されます。
- スマホ等で撮影した写真の向き（EXIF回転情報）は自動的に補正して保存されます。"""

    def __init__(self, root):
        self.root = root
        # タイトル
        self.root.title(f"ImageResizeMiya {self.VERSION}")
        self.root.geometry("450x380") 

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
        
        # 品質設定
        self.quality_options = {
            "最高画質 (Quality: 95)": 95,
            "高画質 (Quality: 85)": 85,
            "中画質 (Quality: 65)": 65,
            "低画質 (Quality: 50)": 50,
            "最低画質 (Quality: 30)": 30,
            "超低画質 (Quality: 10)": 10,
            "限界画質 (Quality: 1)": 1
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

        # 内部フレームでレイアウト調整
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
        inner_quality.pack(fill="x", padx=10, pady=10)

        self.combo_quality = ttk.Combobox(inner_quality, values=list(self.quality_options.keys()), state="readonly", font=self.font_norm)
        self.combo_quality.current(3) # デフォルト: 低画質(50)
        self.combo_quality.pack(fill="x")
        
        lbl_note = tk.Label(
            inner_quality, 
            text="※JPEG/WebP形式で有効です。PNGは圧縮効果が薄い場合があります。", 
            font=self.font_small, 
            fg=self.colors["text_sub"],
            bg=self.colors["bg_card"]
        )
        lbl_note.pack(anchor="w", pady=(5, 0))

        # --- 3. 保存先設定エリア ---
        frame_save = tk.LabelFrame(
            self.root, 
            text=" 3. 保存先 ", 
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

        # --- 4. 実行ボタン ---
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
        btn_run.pack(fill="x", padx=10, pady=15, ipady=5)

    def on_drop(self, event):
        try:
            files = self.root.tk.splitlist(event.data)
            self.update_file_list(files)
        except Exception:
            self.file_path.set(event.data)
            self.selected_files = [event.data]

    def select_file(self):
        filetypes = [("画像ファイル", "*.jpg;*.jpeg;*.png;*.webp"), ("すべてのファイル", "*.*")]
        files = filedialog.askopenfilenames(filetypes=filetypes)
        if files:
            self.update_file_list(files)

    def update_file_list(self, files):
        self.selected_files = files
        display_str = "; ".join(files)
        self.file_path.set(display_str)

    def show_auto_close_message(self, title, message, duration=3000):
        """
        指定時間(ミリ秒)後に自動で閉じるメッセージウィンドウを表示
        OKボタンは無し
        """
        popup = tk.Toplevel(self.root)
        popup.title(title)
        
        # ウィンドウサイズと位置
        width, height = 300, 120 # ボタンがない分少し高さを減らす
        # 親ウィンドウの中心付近に表示するための計算
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (width // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (height // 2)
        popup.geometry(f"{width}x{height}+{x}+{y}")
        
        popup.configure(bg=self.colors["bg_window"])
        
        # メッセージラベル
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
        
        # 最前面に表示
        popup.transient(self.root)
        popup.grab_set()
        self.root.wait_visibility(popup)
        
        # タイマーセット (durationミリ秒後に閉じる)
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
        
        target_dir = None
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
                        file_name, ext = os.path.splitext(os.path.basename(src_path))
                        default_name = f"{file_name}_q{quality_val}{ext}"
                        save_path = ""

                        if self.save_mode.get() == "same":
                            save_path = os.path.join(file_dir, default_name)
                        elif self.save_mode.get() == "arbitrary":
                            if len(target_files) > 1:
                                save_path = os.path.join(target_dir, default_name)
                            else:
                                save_path = filedialog.asksaveasfilename(
                                    initialdir=file_dir,
                                    initialfile=default_name,
                                    defaultextension=ext,
                                    filetypes=[("Original Type", ext), ("JPEG", "*.jpg"), ("WebP", "*.webp")]
                                )
                                if not save_path: continue

                        if save_path:
                            if save_path.lower().endswith((".jpg", ".jpeg")) and img.mode in ("RGBA", "P"):
                                img = img.convert("RGB")
                            
                            img.save(save_path, quality=quality_val, optimize=True)
                            success_count += 1

                except Exception as e:
                    print(f"Error processing {src_path}: {e}")
                    error_count += 1

            self.root.config(cursor="")
            
            msg = f"処理が完了しました。\n成功: {success_count}件"
            if error_count > 0:
                msg += f"\n失敗: {error_count}件"
            
            # 自動で閉じるメッセージを表示 (3秒)
            self.show_auto_close_message("完了", msg, duration=3000)

        except Exception as e:
            self.root.config(cursor="")
            messagebox.showerror("エラー", f"予期せぬエラーが発生しました:\n{e}")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ImageQualityApp(root)
    root.mainloop()