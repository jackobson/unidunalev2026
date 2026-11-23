import tkinter as tk
from tkinter import ttk
from am_eszkozok import am_eszkozok_megjelenitese

root = tk.Tk()
root.title("AM - Leltárkezelő rendszer")
root.geometry("1280x800")
root.minsize(1000, 600)

app = ttk.Frame(root)
app.pack(fill="both", expand=True)

menu = tk.Frame(app, bg="#243e57", width=220)
menu.pack(side="left", fill="y")
menu.pack_propagate(False)

content = tk.Frame(app, bg="#f3f6fa")
content.pack(side="right", fill="both", expand=True)


def am_oldal_valtas(oldal):
    for widget in content.winfo_children():
        widget.destroy()

    tk.Label(
        content,
        text=oldal,
        font=("Segoe UI", 24, "bold"),
        bg="#f3f6fa",
        fg="#243e57"
    ).pack(anchor="w", padx=30, pady=25)

    if oldal == "Eszköznyilvántartás":
        am_eszkozok_megjelenitese(content)


tk.Label(
    menu,
    text="AM Leltárkezelő",
    font=("Segoe UI", 16, "bold"),
    bg="#243e57",
    fg="white"
).pack(pady=25)

menu_items = [
    "Áttekintés",
    "Eszköznyilvántartás",
    "Leltározás",
    "Eszköz áthelyezése",
    "Helyiségek",
    "Felelősök",
    "Statisztikák",
    "Excel-importálás",
    "Excel-exportálás",
    "Kilépés"
]

for item in menu_items:
    tk.Button(
        menu,
        text=item,
        font=("Segoe UI", 11),
        bg="#243e57",
        fg="white",
        activebackground="#3a6285",
        activeforeground="white",
        relief="flat",
        anchor="w",
        padx=20,
        pady=12,
        command=root.destroy if item == "Kilépés"
        else lambda oldal=item: am_oldal_valtas(oldal)
    ).pack(fill="x")

am_oldal_valtas("Áttekintés")

root.mainloop()