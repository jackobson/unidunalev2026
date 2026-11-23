
import tkinter as tk
from tkinter import ttk, messagebox
from am_database import AMAdatbazis
from am_eszkozok import am_eszkozok_megjelenitese
from am_leltarozas import am_leltarozas_megjelenitese


root = tk.Tk()
root.title("AM - Leltárkezelő rendszer")
root.geometry("1280x800")
root.minsize(1100, 700)
root.configure(bg="#F1F4F8")

stilus = ttk.Style()
stilus.theme_use("clam")

stilus.configure(
    "TFrame",
    background="#F1F4F8"
)

stilus.configure(
    "TLabel",
    background="#F1F4F8",
    foreground="#243E57",
    font=("Segoe UI", 10)
)

stilus.configure(
    "TButton",
    font=("Segoe UI", 9),
    padding=(9, 5)
)

stilus.configure(
    "Treeview",
    font=("Segoe UI", 9),
    rowheight=25
)

stilus.configure(
    "Treeview.Heading",
    font=("Segoe UI", 9, "bold")
)

fo_keret = tk.Frame(
    root,
    bg="#F1F4F8"
)
fo_keret.pack(
    fill="both",
    expand=True
)

menu = tk.Frame(
    fo_keret,
    bg="#243E57",
    width=215
)
menu.pack(
    side="left",
    fill="y"
)
menu.pack_propagate(False)

tk.Label(
    menu,
    text="AM Leltárkezelő",
    bg="#243E57",
    fg="white",
    font=("Segoe UI", 16, "bold"),
    anchor="w"
).pack(
    fill="x",
    padx=20,
    pady=(30, 35)
)

jobb_oldal = tk.Frame(
    fo_keret,
    bg="#F1F4F8"
)
jobb_oldal.pack(
    side="left",
    fill="both",
    expand=True
)

fejlec = tk.Frame(
    jobb_oldal,
    bg="#F1F4F8",
    height=105
)
fejlec.pack(
    fill="x"
)
fejlec.pack_propagate(False)

oldal_cim = tk.Label(
    fejlec,
    text="",
    bg="#F1F4F8",
    fg="#243E57",
    font=("Segoe UI", 23, "bold"),
    anchor="w"
)
oldal_cim.pack(
    fill="both",
    expand=True,
    padx=30
)

content = ttk.Frame(jobb_oldal)
content.pack(
    fill="both",
    expand=True
)

app = content


def am_tartalom_torlese():
    for widget in content.winfo_children():
        widget.destroy()

    content._am_eszkoz_aktiv_oldal = None


def am_attekintes():
    db = AMAdatbazis()

    try:
        eszkozok = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM eszkozok
            WHERE archivalt = 0
        """).fetchone()[0]

        helyisegek = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM helyisegek
        """).fetchone()[0]

        felelosok = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM felelosok
        """).fetchone()[0]

    finally:
        db.am_bezaras()

    keret = ttk.Frame(
        content,
        padding=30
    )
    keret.pack(
        fill="both",
        expand=True
    )

    ttk.Label(
        keret,
        text="Üdvözöl az AM Leltárkezelő rendszer!",
        font=("Segoe UI", 16, "bold")
    ).pack(
        anchor="w",
        pady=(0, 25)
    )

    ttk.Label(
        keret,
        text=f"Nyilvántartott aktív eszközök: {eszkozok}",
        font=("Segoe UI", 12)
    ).pack(
        anchor="w",
        pady=10
    )

    ttk.Label(
        keret,
        text=f"Helyiségek száma: {helyisegek}",
        font=("Segoe UI", 12)
    ).pack(
        anchor="w",
        pady=10
    )

    ttk.Label(
        keret,
        text=f"Felelősök száma: {felelosok}",
        font=("Segoe UI", 12)
    ).pack(
        anchor="w",
        pady=10
    )

    ttk.Label(
        keret,
        text=(
            "A bal oldali menüben választhatod ki "
            "a használni kívánt funkciót."
        )
    ).pack(
        anchor="w",
        pady=(35, 0)
    )


def am_helyisegek_oldal():
    db = AMAdatbazis()

    try:
        adatok = db.am_helyisegek_lekerese()
    finally:
        db.am_bezaras()

    keret = ttk.Frame(
        content,
        padding=25
    )
    keret.pack(
        fill="both",
        expand=True
    )

    tabla = ttk.Treeview(
        keret,
        columns=(
            "telephely",
            "helyiseg",
            "megnevezes"
        ),
        show="headings"
    )

    tabla.heading(
        "telephely",
        text="Telephely"
    )
    tabla.heading(
        "helyiseg",
        text="Helyiség"
    )
    tabla.heading(
        "megnevezes",
        text="Megnevezés"
    )

    tabla.column("telephely", width=150)
    tabla.column("helyiseg", width=130)
    tabla.column("megnevezes", width=400)

    gorgeto = ttk.Scrollbar(
        keret,
        command=tabla.yview
    )

    tabla.configure(
        yscrollcommand=gorgeto.set
    )

    gorgeto.pack(
        side="right",
        fill="y"
    )
    tabla.pack(
        side="left",
        fill="both",
        expand=True
    )

    for sor in adatok:
        tabla.insert(
            "",
            "end",
            values=sor[1:]
        )


def am_felelosok_oldal():
    db = AMAdatbazis()

    try:
        adatok = db.am_felelosok_lekerese()
    finally:
        db.am_bezaras()

    keret = ttk.Frame(
        content,
        padding=25
    )
    keret.pack(
        fill="both",
        expand=True
    )

    tabla = ttk.Treeview(
        keret,
        columns=("nev", "sztsz"),
        show="headings"
    )

    tabla.heading(
        "nev",
        text="Felelős neve"
    )
    tabla.heading(
        "sztsz",
        text="Személyi törzsszám"
    )

    tabla.column("nev", width=350)
    tabla.column("sztsz", width=170)

    tabla.pack(
        fill="both",
        expand=True
    )

    for sor in adatok:
        tabla.insert(
            "",
            "end",
            values=(sor[1], sor[2])
        )


def am_helykitolto(szoveg):
    keret = ttk.Frame(
        content,
        padding=30
    )
    keret.pack(
        fill="both",
        expand=True
    )

    ttk.Label(
        keret,
        text=szoveg,
        font=("Segoe UI", 12)
    ).pack(
        anchor="w",
        pady=10
    )


def am_oldal_valtas(oldal):
    if oldal == "Kilépés":
        root.destroy()
        return

    am_tartalom_torlese()
    oldal_cim.config(text=oldal)

    if oldal == "Áttekintés":
        am_attekintes()

    elif oldal == "Eszköznyilvántartás":
        am_eszkozok_megjelenitese(content)

    elif oldal == "Leltározás":
        am_leltarozas_megjelenitese(content)

    elif oldal == "Helyiségek":
        am_helyisegek_oldal()

    elif oldal == "Felelősök":
        am_felelosok_oldal()

    elif oldal == "Eszköz áthelyezése":
        am_helykitolto(
            "Az eszközök áthelyezésének kezelése."
        )

    elif oldal == "Statisztikák":
        am_helykitolto(
            "A leltári statisztikák megjelenítése."
        )

    elif oldal == "Excel-importálás":
        am_helykitolto(
            "Az Excel-adatok importálásának kezelése."
        )

    elif oldal == "Excel-exportálás":
        am_helykitolto(
            "Az eszköznyilvántartás exportálása."
        )


menupontok = [
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

for menupont in menupontok:
    gomb = tk.Button(
        menu,
        text=menupont,
        bg="#243E57",
        fg="white",
        activebackground="#365571",
        activeforeground="white",
        borderwidth=0,
        relief="flat",
        font=("Segoe UI", 10),
        anchor="w",
        padx=20,
        pady=11,
        cursor="hand2",
        command=lambda nev=menupont: am_oldal_valtas(nev)
    )

    gomb.pack(fill="x")


am_oldal_valtas("Áttekintés")

root.mainloop()
