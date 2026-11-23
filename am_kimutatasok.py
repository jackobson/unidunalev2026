
import html
import os
import tempfile
import webbrowser
import tkinter as tk
from pathlib import Path
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from am_database import AMAdatbazis
from am_selejtezes import am_selejtezes_tabla


AM_OSZLOPOK = [
    "Eszközazonosító",
    "Megnevezés",
    "Gyártási szám",
    "Helyiség",
    "Felelős",
    "Mennyiség",
    "Mértékegység",
    "Bruttó érték"
]


def am_penz(ertek):
    if ertek is None:
        return ""
    return f"{ertek:,.0f} Ft".replace(",", " ")


def am_lista_lekerese(tipus):
    am_selejtezes_tabla()
    db = AMAdatbazis()

    try:
        if tipus == "helyiseg":
            return db.kurzor.execute("""
                SELECT
                    h.id,
                    h.telephely,
                    h.helyiseg_szam,
                    h.megnevezes,
                    COUNT(e.id),
                    COALESCE(SUM(e.brutto_ertek), 0)
                FROM helyisegek h
                LEFT JOIN eszkozok e
                    ON e.helyiseg_id = h.id
                    AND e.archivalt = 0
                    AND NOT EXISTS (
                        SELECT 1
                        FROM selejtezesek s
                        WHERE s.eszkoz_id = e.id
                            AND s.allapot = 'elszallitva'
                    )
                GROUP BY
                    h.id,
                    h.telephely,
                    h.helyiseg_szam,
                    h.megnevezes
                ORDER BY h.telephely, h.helyiseg_szam
            """).fetchall()

        return db.kurzor.execute("""
            SELECT
                f.id,
                f.sztsz,
                f.teljes_nev,
                f.osztaly,
                COUNT(e.id),
                COALESCE(SUM(e.brutto_ertek), 0)
            FROM felelosok f
            LEFT JOIN eszkozok e
                ON e.felelos_id = f.id
                AND e.archivalt = 0
                AND NOT EXISTS (
                    SELECT 1
                    FROM selejtezesek s
                    WHERE s.eszkoz_id = e.id
                        AND s.allapot = 'elszallitva'
                )
            GROUP BY
                f.id, f.sztsz, f.teljes_nev, f.osztaly
            ORDER BY f.teljes_nev, f.sztsz
        """).fetchall()

    finally:
        db.am_bezaras()


def am_eszkozlista_lekerese(tipus, rekord_id):
    db = AMAdatbazis()

    try:
        oszlop = (
            "e.helyiseg_id"
            if tipus == "helyiseg"
            else "e.felelos_id"
        )

        return db.kurzor.execute(f"""
            SELECT
                e.eszkozazonosito,
                e.megnevezes,
                e.gyartasi_szam,
                h.helyiseg_szam,
                f.teljes_nev,
                e.mennyiseg,
                e.mertekegyseg,
                e.brutto_ertek
            FROM eszkozok e
            LEFT JOIN helyisegek h
                ON h.id = e.helyiseg_id
            LEFT JOIN felelosok f
                ON f.id = e.felelos_id
            WHERE {oszlop} = ?
                AND e.archivalt = 0
                AND NOT EXISTS (
                    SELECT 1
                    FROM selejtezesek s
                    WHERE s.eszkoz_id = e.id
                        AND s.allapot = 'elszallitva'
                )
            ORDER BY e.eszkozazonosito, e.id
        """, (rekord_id,)).fetchall()

    finally:
        db.am_bezaras()


def am_uj_rekord_ablak(szulo, tipus, frissites):
    helyiseg = tipus == "helyiseg"

    cim = (
        "Új helyiség felvétele"
        if helyiseg else "Új felelős felvétele"
    )

    ablak = tk.Toplevel(szulo)
    ablak.title(f"AM – {cim}")
    ablak.geometry("580x340")
    ablak.resizable(False, False)
    ablak.transient(szulo.winfo_toplevel())
    ablak.grab_set()

    keret = ttk.Frame(ablak, padding=25)
    keret.pack(fill="both", expand=True)

    ttk.Label(
        keret,
        text=cim,
        font=("Segoe UI", 16, "bold")
    ).pack(anchor="w", pady=(0, 22))

    if helyiseg:
        mezok = [
            ("telephely", "Telephely"),
            ("helyiseg_szam", "Helyiség száma"),
            ("megnevezes", "Megnevezés")
        ]
    else:
        mezok = [
            ("sztsz", "Törzsszám"),
            ("teljes_nev", "Teljes név"),
            ("osztaly", "Osztály")
        ]

    valtozok = {}
    adatkeret = ttk.Frame(keret)
    adatkeret.pack(fill="x")

    elso_mezo = None

    for index, (kulcs, felirat) in enumerate(mezok):
        valtozok[kulcs] = tk.StringVar()

        ttk.Label(
            adatkeret,
            text=f"{felirat}:"
        ).grid(
            row=index,
            column=0,
            sticky="w",
            padx=(0, 16),
            pady=9
        )

        beviteli_mezo = ttk.Entry(
            adatkeret,
            textvariable=valtozok[kulcs],
            width=43
        )
        beviteli_mezo.grid(
            row=index,
            column=1,
            sticky="ew",
            pady=9
        )

        if index == 0:
            elso_mezo = beviteli_mezo

    adatkeret.columnconfigure(1, weight=1)

    def am_mentes():
        adatok = {
            kulcs: valtozo.get().strip()
            for kulcs, valtozo in valtozok.items()
        }

        if helyiseg:
            if not all(adatok.values()):
                messagebox.showwarning(
                    "Hiányzó adat",
                    "A telephely, helyiségszám és "
                    "megnevezés megadása kötelező!",
                    parent=ablak
                )
                return
        else:
            if not adatok["sztsz"] or not adatok["teljes_nev"]:
                messagebox.showwarning(
                    "Hiányzó adat",
                    "A törzsszám és a teljes név kötelező!",
                    parent=ablak
                )
                return

        db = None

        try:
            db = AMAdatbazis()
            db.kurzor.execute("BEGIN IMMEDIATE")

            if helyiseg:
                letezik = db.kurzor.execute("""
                    SELECT id
                    FROM helyisegek
                    WHERE LOWER(TRIM(telephely)) =
                          LOWER(TRIM(?))
                      AND LOWER(TRIM(helyiseg_szam)) =
                          LOWER(TRIM(?))
                """, (
                    adatok["telephely"],
                    adatok["helyiseg_szam"]
                )).fetchone()

                if letezik:
                    raise ValueError(
                        "Ezen a telephelyen már létezik "
                        "ilyen helyiségszám!"
                    )

                db.kurzor.execute("""
                    INSERT INTO helyisegek (
                        telephely,
                        helyiseg_szam,
                        megnevezes
                    )
                    VALUES (?, ?, ?)
                """, (
                    adatok["telephely"],
                    adatok["helyiseg_szam"],
                    adatok["megnevezes"]
                ))

            else:
                letezik = db.kurzor.execute("""
                    SELECT id
                    FROM felelosok
                    WHERE TRIM(sztsz) = ?
                """, (
                    adatok["sztsz"],
                )).fetchone()

                if letezik:
                    raise ValueError(
                        "Ez a törzsszám már szerepel "
                        "a nyilvántartásban!"
                    )

                db.kurzor.execute("""
                    INSERT INTO felelosok (
                        sztsz,
                        teljes_nev,
                        osztaly
                    )
                    VALUES (?, ?, ?)
                """, (
                    adatok["sztsz"],
                    adatok["teljes_nev"],
                    adatok["osztaly"]
                ))

            db.kapcsolat.commit()

        except Exception as hiba:
            if db is not None:
                db.kapcsolat.rollback()

            messagebox.showerror(
                "Mentési hiba",
                str(hiba),
                parent=ablak
            )
            return

        finally:
            if db is not None:
                db.am_bezaras()

        ablak.destroy()
        frissites()

        messagebox.showinfo(
            "Sikeres mentés",
            (
                "Az új helyiség sikeresen rögzítve!"
                if helyiseg
                else "Az új felelős sikeresen rögzítve!"
            ),
            parent=szulo.winfo_toplevel()
        )

    gombkeret = ttk.Frame(keret)
    gombkeret.pack(
        side="bottom",
        fill="x",
        pady=(20, 0)
    )

    ttk.Button(
        gombkeret,
        text="Mentés",
        command=am_mentes
    ).pack(side="left", padx=(0, 10))

    ttk.Button(
        gombkeret,
        text="Mégse",
        command=ablak.destroy
    ).pack(side="left")

    if elso_mezo is not None:
        elso_mezo.focus_set()


def am_excel_export(cim, adatok, fajlnev):
    utvonal = filedialog.asksaveasfilename(
        title="Eszközlista exportálása",
        initialfile=fajlnev,
        defaultextension=".xlsx",
        filetypes=[("Excel-munkafüzet", "*.xlsx")]
    )

    if not utvonal:
        return

    munkafuzet = Workbook()
    lap = munkafuzet.active
    lap.title = "Eszközlista"

    lap.merge_cells("A1:H1")
    lap["A1"] = cim
    lap["A1"].font = Font(
        bold=True,
        size=16,
        color="243E57"
    )
    lap["A1"].alignment = Alignment(
        vertical="center"
    )
    lap.row_dimensions[1].height = 30

    lap.merge_cells("A2:H2")
    lap["A2"] = (
        "Kimutatás készítve: "
        + datetime.now().strftime("%Y-%m-%d %H:%M")
    )

    for index, felirat in enumerate(AM_OSZLOPOK, 1):
        cella = lap.cell(row=4, column=index)
        cella.value = felirat
        cella.font = Font(
            bold=True,
            color="FFFFFF"
        )
        cella.fill = PatternFill(
            "solid",
            fgColor="243E57"
        )
        cella.alignment = Alignment(wrap_text=True)

    for index, sor in enumerate(adatok, 5):
        for oszlop, ertek in enumerate(sor, 1):
            cella = lap.cell(
                row=index,
                column=oszlop
            )

            cella.value = (
                "" if ertek is None else ertek
            )

            if oszlop == 8 and ertek is not None:
                cella.number_format = '#,##0 "Ft"'

    osszes = sum(
        sor[7] or 0
        for sor in adatok
    )

    utolso_sor = 5 + len(adatok)

    lap.cell(
        row=utolso_sor,
        column=2,
        value="Összesen"
    ).font = Font(bold=True)

    lap.cell(
        row=utolso_sor,
        column=6,
        value=len(adatok)
    ).font = Font(bold=True)

    lap.cell(
        row=utolso_sor,
        column=8,
        value=osszes
    ).font = Font(bold=True)

    lap.cell(
        row=utolso_sor,
        column=8
    ).number_format = '#,##0 "Ft"'

    szelessegek = [
        20, 43, 24, 17,
        28, 15, 18, 22
    ]

    for index, szelesseg in enumerate(szelessegek, 1):
        lap.column_dimensions[
            get_column_letter(index)
        ].width = szelesseg

    lap.freeze_panes = "A5"
    lap.auto_filter.ref = f"A4:H{4 + len(adatok)}"

    lap.sheet_properties.pageSetUpPr.fitToPage = True
    lap.page_setup.orientation = "landscape"
    lap.page_setup.paperSize = lap.PAPERSIZE_A3
    lap.page_setup.fitToWidth = 1
    lap.page_setup.fitToHeight = 0

    try:
        munkafuzet.save(utvonal)

    except Exception as hiba:
        messagebox.showerror(
            "Exportálási hiba",
            str(hiba)
        )
        return

    messagebox.showinfo(
        "Sikeres exportálás",
        f"Az eszközlista elkészült:\n{utvonal}"
    )


def am_nyomtatas(cim, adatok):
    osszes = sum(
        sor[7] or 0
        for sor in adatok
    )

    fejlecek = "".join(
        f"<th>{html.escape(felirat)}</th>"
        for felirat in AM_OSZLOPOK
    )

    sorok_html = ""

    for sor in adatok:
        cellak = ""

        for index, ertek in enumerate(sor):
            if index == 7:
                szoveg = am_penz(ertek)
            else:
                szoveg = (
                    "" if ertek is None else str(ertek)
                )

            cellak += (
                "<td>"
                + html.escape(szoveg)
                + "</td>"
            )

        sorok_html += f"<tr>{cellak}</tr>"

    dokumentum = f"""<!DOCTYPE html>
<html lang="hu">
<head>
<meta charset="UTF-8">
<title>{html.escape(cim)}</title>
<style>
@page {{
    size: A4 landscape;
    margin: 12mm;
}}
body {{
    font-family: Arial, sans-serif;
    font-size: 10px;
    color: #222;
}}
h1 {{
    color: #243E57;
    font-size: 20px;
}}
p {{
    font-size: 11px;
}}
table {{
    width: 100%;
    border-collapse: collapse;
}}
th, td {{
    border: 1px solid #999;
    padding: 5px;
    text-align: left;
    overflow-wrap: anywhere;
}}
th {{
    background: #e4e9ef;
}}
thead {{
    display: table-header-group;
}}
tr {{
    break-inside: avoid;
}}
.osszesites {{
    margin-top: 18px;
    font-weight: bold;
}}
</style>
</head>
<body>
<h1>{html.escape(cim)}</h1>
<p>Kimutatás készítve:
{datetime.now().strftime("%Y-%m-%d %H:%M")}</p>
<table>
<thead><tr>{fejlecek}</tr></thead>
<tbody>{sorok_html}</tbody>
</table>
<p class="osszesites">
Eszközök száma: {len(adatok)}
&nbsp;&nbsp;&nbsp;
Összes bruttó érték: {html.escape(am_penz(osszes))}
</p>
</body>
</html>"""

    try:
        azonosito, ideiglenes = tempfile.mkstemp(
            prefix="AM_kimutatas_",
            suffix=".html"
        )

        os.close(azonosito)

        fajl = Path(ideiglenes)
        fajl.write_text(
            dokumentum,
            encoding="utf-8"
        )

        megnyitva = webbrowser.open(fajl.as_uri())

        if megnyitva:
            messagebox.showinfo(
                "Nyomtatás",
                "A kimutatás megnyílt a böngészőben.\n\n"
                "Nyomtatás: Ctrl + P\n"
                "PDF mentése: Mentés PDF-ként."
            )
        else:
            messagebox.showinfo(
                "Nyomtatás",
                "A kimutatás elkészült, de a böngésző "
                "nem nyílt meg automatikusan.\n\n"
                f"Fájl: {fajl}"
            )

    except Exception as hiba:
        messagebox.showerror(
            "Nyomtatási hiba",
            str(hiba)
        )


def am_reszletes_ablak(szulo, tipus, rekord):
    rekord_id = rekord[0]

    adatok = am_eszkozlista_lekerese(
        tipus,
        rekord_id
    )

    if tipus == "helyiseg":
        cim = (
            f"Helyiség: {rekord[2]} – "
            f"{rekord[3] or ''}"
        )
        alcim = f"Telephely: {rekord[1] or ''}"
        fajlnev = (
            f"Helyiseg_{rekord[2] or rekord_id}.xlsx"
        )
    else:
        cim = f"Felelős: {rekord[2]}"
        alcim = (
            f"Törzsszám: {rekord[1] or ''}"
        )
        fajlnev = f"Felelos_{rekord_id}.xlsx"

    ablak = tk.Toplevel(szulo)
    ablak.title(f"AM – {cim}")
    ablak.geometry("1100x680")
    ablak.minsize(850, 520)
    ablak.transient(szulo.winfo_toplevel())
    ablak.grab_set()

    alap = ttk.Frame(ablak, padding=15)
    alap.pack(fill="both", expand=True)

    gombkeret = ttk.Frame(alap)
    gombkeret.pack(
        side="bottom",
        fill="x",
        pady=(12, 0)
    )

    ttk.Label(
        alap,
        text=cim,
        font=("Segoe UI", 17, "bold")
    ).pack(
        anchor="w",
        pady=(0, 8)
    )

    ttk.Label(
        alap,
        text=alcim
    ).pack(
        anchor="w",
        pady=(0, 14)
    )

    osszes = sum(
        sor[7] or 0
        for sor in adatok
    )

    ttk.Label(
        alap,
        text=(
            f"Eszközök száma: {len(adatok)}"
            f"        Összes bruttó érték: {am_penz(osszes)}"
        ),
        font=("Segoe UI", 11, "bold")
    ).pack(
        anchor="w",
        pady=(0, 12)
    )

    kereses = tk.StringVar()

    keresokeret = ttk.Frame(alap)
    keresokeret.pack(
        fill="x",
        pady=(0, 12)
    )

    ttk.Label(
        keresokeret,
        text="Keresés:"
    ).pack(
        side="left",
        padx=(0, 8)
    )

    ttk.Entry(
        keresokeret,
        textvariable=kereses,
        width=45
    ).pack(side="left")

    tablakeret = ttk.Frame(alap)
    tablakeret.pack(
        fill="both",
        expand=True
    )

    oszlopok = (
        "azonosito",
        "nev",
        "gyartasi",
        "helyiseg",
        "felelos",
        "mennyiseg",
        "mertekegyseg",
        "ertek"
    )

    tabla = ttk.Treeview(
        tablakeret,
        columns=oszlopok,
        show="headings",
        selectmode="browse"
    )

    szelessegek = [
        140, 300, 150, 90,
        160, 90, 100, 135
    ]

    for kulcs, felirat, szelesseg in zip(
        oszlopok,
        AM_OSZLOPOK,
        szelessegek
    ):
        tabla.heading(
            kulcs,
            text=felirat
        )
        tabla.column(
            kulcs,
            width=szelesseg
        )

    tabla.column("ertek", anchor="e")

    gorgeto_y = ttk.Scrollbar(
        tablakeret,
        orient="vertical",
        command=tabla.yview
    )

    gorgeto_x = ttk.Scrollbar(
        tablakeret,
        orient="horizontal",
        command=tabla.xview
    )

    tabla.configure(
        yscrollcommand=gorgeto_y.set,
        xscrollcommand=gorgeto_x.set
    )

    gorgeto_y.pack(side="right", fill="y")
    gorgeto_x.pack(side="bottom", fill="x")

    tabla.pack(
        side="left",
        fill="both",
        expand=True
    )

    talalatok = tk.StringVar()

    ttk.Label(
        alap,
        textvariable=talalatok
    ).pack(
        anchor="w",
        pady=(8, 0)
    )

    def am_szures(*args):
        for elem in tabla.get_children():
            tabla.delete(elem)

        keresett = (
            kereses.get().strip().casefold()
        )
        darab = 0

        for index, sor in enumerate(adatok):
            szoveg = " ".join(
                str(ertek or "")
                for ertek in sor[:5]
            ).casefold()

            if keresett not in szoveg:
                continue

            ertekek = [
                "" if ertek is None else ertek
                for ertek in sor
            ]

            ertekek[7] = am_penz(sor[7])

            tabla.insert(
                "",
                "end",
                iid=str(index),
                values=ertekek
            )

            darab += 1

        talalatok.set(
            f"Találatok száma: {darab} / {len(adatok)}"
        )

    kereses.trace_add("write", am_szures)

    ttk.Button(
        gombkeret,
        text="Excel-export",
        command=lambda: am_excel_export(
            cim,
            adatok,
            fajlnev
        )
    ).pack(side="left", padx=5)

    ttk.Button(
        gombkeret,
        text="Nyomtatás / PDF",
        command=lambda: am_nyomtatas(
            cim,
            adatok
        )
    ).pack(side="left", padx=5)

    ttk.Button(
        gombkeret,
        text="Bezárás",
        command=ablak.destroy
    ).pack(side="right", padx=5)

    am_szures()


def am_kimutatas_megjelenitese(content, tipus):
    lista = am_lista_lekerese(tipus)

    oldal = ttk.Frame(content)
    oldal.pack(fill="both", expand=True)

    if tipus == "helyiseg":
        oszlopok = (
            "telephely",
            "helyiseg",
            "megnevezes",
            "darab",
            "ertek"
        )

        fejlecek = [
            ("telephely", "Telephely", 140),
            ("helyiseg", "Helyiség", 110),
            ("megnevezes", "Megnevezés", 330),
            ("darab", "Eszközök száma", 135),
            ("ertek", "Bruttó érték", 160)
        ]

        uj_gomb = "Új helyiség"
    else:
        oszlopok = (
            "torzsszam",
            "nev",
            "osztaly",
            "darab",
            "ertek"
        )

        fejlecek = [
            ("torzsszam", "Törzsszám", 140),
            ("nev", "Felelős neve", 240),
            ("osztaly", "Osztály", 250),
            ("darab", "Eszközök száma", 135),
            ("ertek", "Bruttó érték", 160)
        ]

        uj_gomb = "Új felelős"

    ttk.Label(
        oldal,
        text=(
            "Kattints duplán egy sorra "
            "a hozzá tartozó eszközök megtekintéséhez."
        )
    ).pack(
        anchor="w",
        padx=25,
        pady=(10, 16)
    )

    felso_gombok = ttk.Frame(oldal)
    felso_gombok.pack(
        fill="x",
        padx=25,
        pady=(0, 12)
    )

    ttk.Button(
        felso_gombok,
        text=uj_gomb,
        command=lambda: am_uj_rekord_ablak(
            content,
            tipus,
            lambda: am_kimutatas_frissites(
                content, tipus
            )
        )
    ).pack(side="left", padx=(0, 10))

    kereses = tk.StringVar()

    keresokeret = ttk.Frame(oldal)
    keresokeret.pack(
        fill="x",
        padx=25,
        pady=(0, 12)
    )

    ttk.Label(
        keresokeret,
        text="Keresés:"
    ).pack(
        side="left",
        padx=(0, 8)
    )

    ttk.Entry(
        keresokeret,
        textvariable=kereses,
        width=42
    ).pack(side="left")

    tablakeret = ttk.Frame(oldal)
    tablakeret.pack(
        fill="both",
        expand=True,
        padx=25,
        pady=10
    )

    tabla = ttk.Treeview(
        tablakeret,
        columns=oszlopok,
        show="headings",
        selectmode="browse"
    )

    for kulcs, felirat, szelesseg in fejlecek:
        tabla.heading(
            kulcs,
            text=felirat
        )
        tabla.column(
            kulcs,
            width=szelesseg
        )

    tabla.column("darab", anchor="center")
    tabla.column("ertek", anchor="e")

    gorgeto = ttk.Scrollbar(
        tablakeret,
        orient="vertical",
        command=tabla.yview
    )

    tabla.configure(
        yscrollcommand=gorgeto.set
    )

    gorgeto.pack(side="right", fill="y")

    tabla.pack(
        side="left",
        fill="both",
        expand=True
    )

    rekordok = {
        sor[0]: sor
        for sor in lista
    }

    talalatok = tk.StringVar()

    ttk.Label(
        oldal,
        textvariable=talalatok
    ).pack(
        anchor="w",
        padx=25,
        pady=5
    )

    def am_szures(*args):
        for elem in tabla.get_children():
            tabla.delete(elem)

        keresett = (
            kereses.get().strip().casefold()
        )
        darab = 0

        for sor in lista:
            szoveg = " ".join(
                str(ertek or "")
                for ertek in sor[1:4]
            ).casefold()

            if keresett not in szoveg:
                continue

            tabla.insert(
                "",
                "end",
                iid=str(sor[0]),
                values=(
                    sor[1] or "",
                    sor[2] or "",
                    sor[3] or "",
                    sor[4],
                    am_penz(sor[5])
                )
            )

            darab += 1

        talalatok.set(
            f"Találatok: {darab} / {len(lista)}"
        )

    def am_reszletek(event=None):
        kijeloles = tabla.selection()

        if not kijeloles:
            messagebox.showwarning(
                "Nincs kijelölés",
                "Először válassz ki egy sort!"
            )
            return

        rekord = rekordok.get(
            int(kijeloles[0])
        )

        if rekord is None:
            messagebox.showerror(
                "Hiba",
                "A kiválasztott adat nem található!"
            )
            return

        am_reszletes_ablak(
            content,
            tipus,
            rekord
        )

    tabla.bind(
        "<Double-1>",
        am_reszletek
    )

    kereses.trace_add(
        "write",
        am_szures
    )

    gombkeret = ttk.Frame(oldal)
    gombkeret.pack(
        fill="x",
        padx=25,
        pady=(8, 15)
    )

    ttk.Button(
        gombkeret,
        text="Kijelölt sor eszközei",
        command=am_reszletek
    ).pack(side="left", padx=5)

    ttk.Button(
        gombkeret,
        text="Frissítés",
        command=lambda: am_kimutatas_frissites(
            content, tipus
        )
    ).pack(side="left", padx=5)

    am_szures()


def am_kimutatas_frissites(content, tipus):
    if not content.winfo_exists():
        return

    for widget in content.winfo_children():
        widget.destroy()

    am_kimutatas_megjelenitese(
        content,
        tipus
    )


def am_helyisegek_megjelenitese(content):
    am_kimutatas_megjelenitese(
        content,
        "helyiseg"
    )


def am_felelosok_megjelenitese(content):
    am_kimutatas_megjelenitese(
        content,
        "felelos"
    )
