import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image
import os
from tkinterdnd2 import DND_FILES, TkinterDnD

class ImageResizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ImageResizeMiya")
        self.root.geometry("450x350")

        # 変数
        self.file_path = tk.StringVar()
        self.save_mode = tk.StringVar(value="same") # 初期値: 同じフォルダ
        
        # 一般的なサイズ定義 (幅, 高さ)
        self.size_options = {
            "Full HD (1920 x 1080)": (1920, 1080),
            "HD (1280 x 720)": (1280, 720),
            "XGA (1024 x 768)": (1024, 768),
            "Instagram Square (1080 x 1080)": (1080, 1080),
            "SVGA (800 x 600)": (800, 600),
            "VGA (640 x 480)": (640, 480),
            "Icon (256 x 256)": (256, 256)
        }

        self._create_widgets()

    def _create_widgets(self):
        # 1. ファイル選択 (ドラッグ&ドロップ対応)
        frame_file = tk.LabelFrame(self.root, text="1. 画像を選択 (ドラッグ&ドロップ可)", padx=10, pady=10)
        frame_file.pack(fill="x", padx=10, pady=5)

        self.entry_file = tk.Entry(frame_file, textvariable=self.file_path)
        self.entry_file.pack(side="left", fill="x", expand=True, padx=5)
        
        # DND設定
        self.entry_file.drop_target_register(DND_FILES)
        self.entry_file.dnd_bind('<<Drop>>', self.on_drop)

        tk.Button(frame_file, text="参照", command=self.select_file).pack(side="left")

        # 2. サイズ選択
        frame_size = tk.LabelFrame(self.root, text="2. サイズを選択", padx=10, pady=10)
        frame_size.pack(fill="x", padx=10, pady=5)

        self.combo_size = ttk.Combobox(frame_size, values=list(self.size_options.keys()), state="readonly")
        self.combo_size.current(0)
        self.combo_size.pack(fill="x", padx=5)

        # 3. 保存先設定
        frame_save = tk.LabelFrame(self.root, text="3. 保存先", padx=10, pady=10)
        frame_save.pack(fill="x", padx=10, pady=5)

        # ラジオボタン (デフォルトは「同じフォルダ」)
        tk.Radiobutton(frame_save, text="同じフォルダに保存 (ファイル名_resized)", variable=self.save_mode, value="same").pack(anchor="w")
        tk.Radiobutton(frame_save, text="任意のフォルダを指定 (保存時に選択)", variable=self.save_mode, value="arbitrary").pack(anchor="w")

        # 4. 実行ボタン
        btn_resize = tk.Button(self.root, text="リサイズ実行", command=self.resize_image, bg="#ddd", height=2)
        btn_resize.pack(fill="x", padx=10, pady=10)

    def on_drop(self, event):
        # ドラッグ&ドロップされたファイルのパスを取得
        path = event.data
        # Windows等でパスにスペースが含まれる場合、{}で囲まれることがあるため除去
        if path.startswith('{') and path.endswith('}'):
            path = path[1:-1]
        self.file_path.set(path)

    def select_file(self):
        filetypes = [("画像ファイル", "*.jpg;*.jpeg;*.png;*.bmp;*.webp"), ("すべてのファイル", "*.*")]
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            self.file_path.set(path)

    def resize_image(self):
        src_path = self.file_path.get()
        if not src_path or not os.path.exists(src_path):
            messagebox.showerror("エラー", "有効な画像ファイルを選択してください。")
            return

        target_size = self.size_options[self.combo_size.get()]
        
        try:
            # 画像処理
            with Image.open(src_path) as img:
                resized_img = img.resize(target_size, Image.Resampling.LANCZOS)
                
                file_dir = os.path.dirname(src_path)
                file_name, ext = os.path.splitext(os.path.basename(src_path))
                default_name = f"{file_name}_resized{ext}"
                save_path = ""

                # 保存先の分岐
                if self.save_mode.get() == "same":
                    save_path = os.path.join(file_dir, default_name)
                else:
                    save_path = filedialog.asksaveasfilename(
                        initialdir=file_dir,
                        initialfile=default_name,
                        defaultextension=ext,
                        filetypes=[("Original Type", ext)]
                    )

                # 保存実行
                if save_path:
                    resized_img.save(save_path)
                    messagebox.showinfo("成功", f"保存しました:\n{save_path}")

        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")

if __name__ == "__main__":
    # TkinterDnD.Tkを使用する
    root = TkinterDnD.Tk()
    app = ImageResizerApp(root)
    root.mainloop()