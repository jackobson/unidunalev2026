
import tkinter as tk
from tkinter import ttk, messagebox
from am_database import AMAdatbazis
from am_eszkozok import am_eszkozok_megjelenitese
from am_leltarozas import am_leltarozas_megjelenitese
from am_athelyezes import am_athelyezes_megjelenitese
from am_selejtezes import am_selejtezes_tabla
from am_kimutatasok import (
    am_helyisegek_megjelenitese,
    am_felelosok_megjelenitese
)
from am_statisztikak import am_statisztikak_megjelenitese
from am_nyilvantartasok import (
    AMNyilvantartasKezelo,
    am_nyilvantartas_valaszto
)
from am_excel import am_excel_megjelenitese


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
fo_keret.pack(fill="both", expand=True)

menu = tk.Frame(
    fo_keret,
    bg="#243E57",
    width=225
)
menu.pack(side="left", fill="y")
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
    pady=(25, 12)
)

tk.Label(
    menu,
    text="Aktív nyilvántartás:",
    bg="#243E57",
    fg="#BACAD8",
    font=("Segoe UI", 9),
    anchor="w"
).pack(
    fill="x",
    padx=20,
    pady=(0, 3)
)

aktiv_nev = tk.StringVar(
    value="Nincs kiválasztva"
)

tk.Label(
    menu,
    textvariable=aktiv_nev,
    bg="#243E57",
    fg="white",
    font=("Segoe UI", 10, "bold"),
    anchor="w",
    justify="left",
    wraplength=190
).pack(
    fill="x",
    padx=20,
    pady=(0, 17)
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
fejlec.pack(fill="x")
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
content.pack(fill="both", expand=True)

app = content

am_kezelo = None
am_aktualis_rekord = None


def am_tartalom_torlese():
    for widget in content.winfo_children():
        widget.destroy()

    content._am_eszkoz_aktiv_oldal = None


def am_attekintes():
    am_selejtezes_tabla()

    db = AMAdatbazis()

    try:
        eszkozok = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM eszkozok e
            WHERE e.archivalt = 0
                AND NOT EXISTS (
                    SELECT 1
                    FROM selejtezesek s
                    WHERE s.eszkoz_id = e.id
                        AND s.allapot = 'elszallitva'
                )
        """).fetchone()[0]

        helyisegek = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM helyisegek
        """).fetchone()[0]

        felelosok = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM felelosok
        """).fetchone()[0]

        athelyezesek = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM athelyezesek
        """).fetchone()[0]

        leltarozasok = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM leltarozas
        """).fetchone()[0]

    finally:
        db.am_bezaras()

    keret = ttk.Frame(
        content,
        padding=30
    )
    keret.pack(fill="both", expand=True)

    ttk.Label(
        keret,
        text="Üdvözöl az AM Leltárkezelő rendszer!",
        font=("Segoe UI", 16, "bold")
    ).pack(
        anchor="w",
        pady=(0, 12)
    )

    ttk.Label(
        keret,
        text=f"Nyilvántartás: {aktiv_nev.get()}",
        font=("Segoe UI", 12, "bold")
    ).pack(
        anchor="w",
        pady=(0, 25)
    )

    adatok = [
        (
            "Nyilvántartott aktív eszközök",
            eszkozok
        ),
        (
            "Helyiségek száma",
            helyisegek
        ),
        (
            "Felelősök száma",
            felelosok
        ),
        (
            "Rögzített áthelyezések",
            athelyezesek
        ),
        (
            "Leltározási bejegyzések",
            leltarozasok
        )
    ]

    for felirat, ertek in adatok:
        ttk.Label(
            keret,
            text=f"{felirat}: {ertek}",
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


def am_oldal_valtas(oldal):
    if oldal == "Kilépés":
        root.destroy()
        return

    if oldal == "Nyilvántartás váltása":
        am_valaszto_megnyitasa()
        return

    if am_aktualis_rekord is None:
        return

    am_tartalom_torlese()
    oldal_cim.config(text=oldal)

    if oldal == "Áttekintés":
        am_attekintes()

    elif oldal == "Eszköznyilvántartás":
        am_eszkozok_megjelenitese(content)

    elif oldal == "Leltározás":
        am_leltarozas_megjelenitese(content)

    elif oldal == "Eszköz áthelyezése":
        am_athelyezes_megjelenitese(content)

    elif oldal == "Helyiségek":
        am_helyisegek_megjelenitese(content)

    elif oldal == "Felelősök":
        am_felelosok_megjelenitese(content)

    elif oldal == "Statisztikák":
        am_statisztikak_megjelenitese(content)

    elif oldal == "Excel-importálás":
        am_excel_megjelenitese(
            content,
            "import",
            aktiv_nev.get()
        )

    elif oldal == "Excel-exportálás":
        am_excel_megjelenitese(
            content,
            "export",
            aktiv_nev.get()
        )


def am_nyilvantartas_megnyitasa(rekord):
    global am_aktualis_rekord

    regi_utvonal = (
        AMAdatbazis.am_aktiv_utvonal_lekerese()
    )

    try:
        utvonal = am_kezelo.am_utvonal(rekord)

        AMAdatbazis.am_aktiv_utvonal_beallitasa(
            utvonal
        )

        am_selejtezes_tabla()

    except Exception as hiba:
        AMAdatbazis.am_aktiv_utvonal_beallitasa(
            regi_utvonal
        )

        messagebox.showerror(
            "Nyilvántartás megnyitási hiba",
            str(hiba)
        )
        return False

    am_aktualis_rekord = rekord
    aktiv_nev.set(rekord["nev"])

    root.title(
        f"AM - Leltárkezelő rendszer - {rekord['nev']}"
    )

    root.deiconify()
    am_oldal_valtas("Áttekintés")

    return True


def am_valaszto_megnyitasa(inditas=False):
    aktualis_id = None

    if am_aktualis_rekord is not None:
        aktualis_id = am_aktualis_rekord["id"]

    am_nyilvantartas_valaszto(
        root,
        am_kezelo,
        am_nyilvantartas_megnyitasa,
        aktualis_id=aktualis_id,
        inditas=inditas
    )


menupontok = [
    "Nyilvántartás váltása",
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
        bg=(
            "#365571"
            if menupont == "Nyilvántartás váltása"
            else "#243E57"
        ),
        fg="white",
        activebackground="#486986",
        activeforeground="white",
        borderwidth=0,
        relief="flat",
        font=("Segoe UI", 10),
        anchor="w",
        padx=20,
        pady=10,
        cursor="hand2",
        command=lambda nev=menupont: am_oldal_valtas(nev)
    )

    gomb.pack(fill="x")


root.withdraw()

try:
    am_kezelo = AMNyilvantartasKezelo()
    am_valaszto_megnyitasa(inditas=True)

except Exception as hiba:
    messagebox.showerror(
        "Indítási hiba",
        str(hiba)
    )
    root.destroy()

root.mainloop()
