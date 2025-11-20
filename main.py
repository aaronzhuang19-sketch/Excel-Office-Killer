import os
import io
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD
from PIL import Image

from ui.views import SplitterView, MergerView


def main():
    root = TkinterDnD.Tk()
    root.title("Office Killer Toolkit V1")
    root.geometry("900x600")
    try:
        ico_path = os.path.join(os.getcwd(), "Devsoul3.ico")
        if os.path.exists(ico_path):
            root.iconbitmap(ico_path)
    except Exception:
        pass

    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")

    header = ctk.CTkFrame(root)
    header.pack(fill="x", padx=12, pady=8)
    png_path = os.path.join(os.getcwd(), "Devsoul3.png")
    logo_img = None
    if os.path.exists(png_path):
        try:
            img = Image.open(png_path)
            logo_img = ctk.CTkImage(light_image=img, dark_image=img, size=(24, 24))
        except Exception:
            logo_img = None
    if logo_img:
        logo = ctk.CTkLabel(header, image=logo_img, text="")
        logo.pack(side="left", padx=8)
    title = ctk.CTkLabel(header, text="Office Killer Toolkit V1")
    title.pack(side="left", padx=8)

    tab = ctk.CTkTabview(root)
    tab.pack(fill="both", expand=True, padx=12, pady=12)

    tab.add("拆分")
    tab.add("合并")

    split_view = SplitterView(tab.tab("拆分"))
    split_view.pack(fill="both", expand=True)

    merge_view = MergerView(tab.tab("合并"))
    merge_view.pack(fill="both", expand=True)

    footer = ctk.CTkFrame(root, fg_color=("#ffffff", "#0f172a"))
    footer.pack(side="bottom", fill="x", padx=0, pady=0)
    f1 = ctk.CTkLabel(footer, text="Devsoul ", text_color=("#2563eb", "#60a5fa"))
    f1.pack(side="left", padx=12, pady=8)
    f2 = ctk.CTkLabel(footer, text="OfficeKiller ", text_color=("#ef4444", "#f87171"))
    f2.pack(side="left", padx=0, pady=8)
    f3 = ctk.CTkLabel(footer, text="E-mail：", text_color=("#f59e0b", "#fbbf24"))
    f3.pack(side="left", padx=6, pady=8)
    f4 = ctk.CTkLabel(footer, text="cs@devsoul-ai.com", text_color=("#10b981", "#34d399"))
    f4.pack(side="left", padx=0, pady=8)

    root.mainloop()


if __name__ == "__main__":
    main()