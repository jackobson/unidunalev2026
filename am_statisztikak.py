
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from am_database import AMAdatbazis
from am_selejtezes import am_selejtezes_tabla


AM_AKTIV_FELTETEL = """
    e.archivalt = 0
    AND NOT EXISTS (
        SELECT 1
        FROM selejtezesek s
        WHERE s.eszkoz_id = e.id
          AND s.allapot = 'elszallitva'
    )
"""

AM_SZINEK = [
    "#3274A1",
    "#3F927B",
    "#D29B45",
    "#916CB0",
    "#C36A68",
    "#548AAB",
    "#6D9873",
    "#AE8559",
    "#746AA4",
    "#648B91"
]


def am_statisztika_adatok(datum):
    datum = date.fromisoformat(datum).isoformat()
    am_selejtezes_tabla()

    db = AMAdatbazis()

    try:
        aktiv = db.kurzor.execute(f"""
            SELECT
                COUNT(*),
                COALESCE(SUM(e.brutto_ertek), 0),
                COALESCE(SUM(e.konyv_szerinti_ertek), 0)
            FROM eszkozok e
            WHERE {AM_AKTIV_FELTETEL}
        """).fetchone()

        archivalt = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM eszkozok e
            WHERE e.archivalt = 1
              AND NOT EXISTS (
                  SELECT 1
                  FROM selejtezesek s
                  WHERE s.eszkoz_id = e.id
                    AND s.allapot = 'elszallitva'
              )
        """).fetchone()[0]

        elszallitott = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM selejtezesek
            WHERE allapot = 'elszallitva'
        """).fetchone()[0]

        osszes = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM eszkozok
        """).fetchone()[0]

        helyisegek = db.kurzor.execute(f"""
            SELECT
                COALESCE(h.telephely, '') || ' / ' ||
                COALESCE(h.helyiseg_szam, 'Nincs helyiség'),
                COUNT(e.id),
                COALESCE(SUM(e.brutto_ertek), 0)
            FROM eszkozok e
            LEFT JOIN helyisegek h
                ON h.id = e.helyiseg_id
            WHERE {AM_AKTIV_FELTETEL}
            GROUP BY e.helyiseg_id
            ORDER BY COUNT(e.id) DESC, 1
        """).fetchall()

        felelosok = db.kurzor.execute(f"""
            SELECT
                COALESCE(f.teljes_nev, 'Nincs felelős') ||
                CASE
                    WHEN f.sztsz IS NOT NULL
                    THEN ' (' || f.sztsz || ')'
                    ELSE ''
                END,
                COUNT(e.id),
                COALESCE(SUM(e.brutto_ertek), 0)
            FROM eszkozok e
            LEFT JOIN felelosok f
                ON f.id = e.felelos_id
            WHERE {AM_AKTIV_FELTETEL}
            GROUP BY e.felelos_id
            ORDER BY COUNT(e.id) DESC, 1
        """).fetchall()

        ellenorzesek = db.kurzor.execute(f"""
            SELECT
                COALESCE(l.allapot, 'Nem ellenőrzött'),
                COUNT(*)
            FROM eszkozok e
            LEFT JOIN leltarozas l
                ON l.id = (
                    SELECT MAX(l2.id)
                    FROM leltarozas l2
                    WHERE l2.eszkoz_id = e.id
                      AND l2.datum = ?
                )
            WHERE {AM_AKTIV_FELTETEL}
            GROUP BY COALESCE(l.allapot, 'Nem ellenőrzött')
        """, (datum,)).fetchall()

        selejtezesek = db.kurzor.execute("""
            SELECT allapot, COUNT(*)
            FROM selejtezesek
            GROUP BY allapot
        """).fetchall()

        athelyezesek = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM athelyezesek
        """).fetchone()[0]

        leltar_bejegyzesek = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM leltarozas
        """).fetchone()[0]

        helyisegek_szama = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM helyisegek
        """).fetchone()[0]

        felelosok_szama = db.kurzor.execute("""
            SELECT COUNT(*)
            FROM felelosok
        """).fetchone()[0]

    finally:
        db.am_bezaras()

    allapotok = {
        "Nem ellenőrzött": 0,
        "Megvan": 0,
        "Hiányzik": 0,
        "Sérült": 0
    }

    for allapot, darab in ellenorzesek:
        allapotok[allapot] = darab

    selejtezes_allapotok = {
        "jelolt": 0,
        "jovahagyott": 0,
        "elszallitva": 0,
        "visszavont": 0
    }

    for allapot, darab in selejtezesek:
        selejtezes_allapotok[allapot] = darab

    return {
        "datum": datum,
        "aktiv": aktiv[0],
        "brutto": aktiv[1],
        "konyv": aktiv[2],
        "archivalt": archivalt,
        "elszallitott": elszallitott,
        "osszes": osszes,
        "helyisegek": helyisegek,
        "felelosok": felelosok,
        "leltar": allapotok,
        "selejtezes": selejtezes_allapotok,
        "athelyezesek": athelyezesek,
        "leltar_bejegyzesek": leltar_bejegyzesek,
        "helyisegek_szama": helyisegek_szama,
        "felelosok_szama": felelosok_szama
    }


def am_forint(ertek):
    return f"{ertek:,.0f} Ft".replace(",", " ")


def am_diagram(szulo, adatok, cim):
    keret = ttk.Frame(szulo, padding=10)
    keret.pack(fill="both", expand=True)

    ttk.Label(
        keret,
        text=cim,
        font=("Segoe UI", 13, "bold")
    ).pack(anchor="w", pady=(0, 8))

    vaszon = tk.Canvas(
        keret,
        bg="white",
        highlightthickness=1,
        highlightbackground="#D8E0E8",
        height=360
    )
    vaszon.pack(fill="both", expand=True)

    def am_rajzolas(event=None):
        vaszon.delete("all")

        if not adatok:
            vaszon.create_text(
                200, 100,
                text="Nincs megjeleníthető adat.",
                fill="#526578",
                font=("Segoe UI", 12)
            )
            return

        szelesseg = max(vaszon.winfo_width(), 400)
        magassag = max(vaszon.winfo_height(), 250)

        bal = min(240, int(szelesseg * 0.35))
        jobb = 75
        felso = 25
        also = 20

        maximum = max(ertek for _, ertek in adatok)
        maximum = max(maximum, 1)

        sor_magassag = (
            (magassag - felso - also) / len(adatok)
        )

        sav_hossz = max(szelesseg - bal - jobb, 50)

        for index, (nev, ertek) in enumerate(adatok):
            y = felso + index * sor_magassag
            kozep = y + sor_magassag / 2

            felirat = str(nev)
            if len(felirat) > 31:
                felirat = felirat[:28] + "..."

            vaszon.create_text(
                bal - 12,
                kozep,
                text=felirat,
                anchor="e",
                fill="#243E57",
                font=("Segoe UI", 9)
            )

            hosszusag = sav_hossz * ertek / maximum

            vastagsag = min(
                25,
                max(sor_magassag * 0.55, 6)
            )

            vaszon.create_rectangle(
                bal,
                kozep - vastagsag / 2,
                bal + hosszusag,
                kozep + vastagsag / 2,
                fill=AM_SZINEK[index % len(AM_SZINEK)],
                outline=""
            )

            vaszon.create_text(
                bal + hosszusag + 8,
                kozep,
                text=str(ertek),
                anchor="w",
                fill="#243E57",
                font=("Segoe UI", 9, "bold")
            )

    vaszon.bind("<Configure>", am_rajzolas)

    return keret


def am_statisztika_tabla(szulo, fejlecek, sorok):
    keret = ttk.Frame(szulo)
    keret.pack(
        fill="both",
        expand=True,
        padx=10,
        pady=10
    )

    oszlopok = [
        f"oszlop_{i}" for i in range(len(fejlecek))
    ]

    tabla = ttk.Treeview(
        keret,
        columns=oszlopok,
        show="headings"
    )

    for index, (kulcs, felirat) in enumerate(
        zip(oszlopok, fejlecek)
    ):
        tabla.heading(kulcs, text=felirat)
        tabla.column(
            kulcs,
            width=350 if index == 0 else 150,
            anchor="w" if index == 0 else "e"
        )

    gorgeto = ttk.Scrollbar(
        keret,
        orient="vertical",
        command=tabla.yview
    )
    tabla.configure(yscrollcommand=gorgeto.set)

    gorgeto.pack(side="right", fill="y")
    tabla.pack(side="left", fill="both", expand=True)

    for sor in sorok:
        tabla.insert("", "end", values=sor)


def am_statisztika_excel(adatok):
    utvonal = filedialog.asksaveasfilename(
        title="Statisztikák exportálása",
        defaultextension=".xlsx",
        initialfile=(
            f"AM_Statisztikak_{adatok['datum']}.xlsx"
        ),
        filetypes=[("Excel-munkafüzet", "*.xlsx")]
    )

    if not utvonal:
        return

    munkafuzet = Workbook()
    attekintes = munkafuzet.active
    attekintes.title = "Áttekintés"

    attekintes.append(["Mutató", "Érték"])

    mutatok = [
        ("Kimutatás dátuma", adatok["datum"]),
        ("Összes adatbázisrekord", adatok["osszes"]),
        ("Aktív eszközök", adatok["aktiv"]),
        ("Archivált eszközök", adatok["archivalt"]),
        ("Elszállított eszközök", adatok["elszallitott"]),
        ("Aktív állomány bruttó értéke", adatok["brutto"]),
        ("Aktív állomány könyv szerinti értéke", adatok["konyv"]),
        ("Helyiségek száma", adatok["helyisegek_szama"]),
        ("Felelősök száma", adatok["felelosok_szama"]),
        ("Áthelyezések száma", adatok["athelyezesek"]),
        ("Leltározási bejegyzések", adatok["leltar_bejegyzesek"])
    ]

    for sor in mutatok:
        attekintes.append(sor)

    helyisegek = munkafuzet.create_sheet("Helyiségek")
    helyisegek.append([
        "Helyiség", "Eszközök száma", "Bruttó érték"
    ])

    for sor in adatok["helyisegek"]:
        helyisegek.append(sor)

    felelosok = munkafuzet.create_sheet("Felelősök")
    felelosok.append([
        "Felelős", "Eszközök száma", "Bruttó érték"
    ])

    for sor in adatok["felelosok"]:
        felelosok.append(sor)

    leltar = munkafuzet.create_sheet("Leltározás")
    leltar.append(["Dátum", adatok["datum"]])
    leltar.append(["Állapot", "Eszközök száma"])

    for allapot, darab in adatok["leltar"].items():
        leltar.append([allapot, darab])

    muveletek = munkafuzet.create_sheet("Műveletek")
    muveletek.append(["Mutató", "Darabszám"])
    muveletek.append([
        "Áthelyezések", adatok["athelyezesek"]
    ])
    muveletek.append([
        "Leltározási bejegyzések",
        adatok["leltar_bejegyzesek"]
    ])

    allapot_nevek = {
        "jelolt": "Selejtezésre jelölt",
        "jovahagyott": "Jóváhagyott",
        "elszallitva": "Elszállított",
        "visszavont": "Visszavont"
    }

    for kod, darab in adatok["selejtezes"].items():
        muveletek.append([
            allapot_nevek.get(kod, kod),
            darab
        ])

    for lap in munkafuzet.worksheets:
        fejlec_sor = 2 if lap.title == "Leltározás" else 1

        for cella in lap[fejlec_sor]:
            cella.font = Font(
                bold=True,
                color="FFFFFF"
            )
            cella.fill = PatternFill(
                "solid",
                fgColor="243E57"
            )
            cella.alignment = Alignment(
                wrap_text=True
            )

        lap.freeze_panes = (
            "A3" if lap.title == "Leltározás" else "A2"
        )

        lap.column_dimensions["A"].width = 43
        lap.column_dimensions["B"].width = 22

        if lap.max_column >= 3:
            lap.column_dimensions["C"].width = 24

        lap.sheet_properties.pageSetUpPr.fitToPage = True
        lap.page_setup.orientation = "landscape"
        lap.page_setup.fitToWidth = 1
        lap.page_setup.fitToHeight = 0

    for lap in (helyisegek, felelosok):
        for sor in lap.iter_rows(
            min_row=2, min_col=3, max_col=3
        ):
            sor[0].number_format = '#,##0 "Ft"'

    attekintes["B7"].number_format = '#,##0 "Ft"'
    attekintes["B8"].number_format = '#,##0 "Ft"'

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
        f"A statisztikai Excel-fájl elkészült:\n{utvonal}"
    )


def am_statisztikak_megjelenitese(content):
    oldal = ttk.Frame(content)
    oldal.pack(fill="both", expand=True)

    felso = ttk.Frame(oldal, padding=(22, 10))
    felso.pack(fill="x")

    ttk.Label(
        felso,
        text="Statisztikai áttekintés",
        font=("Segoe UI", 16, "bold")
    ).pack(anchor="w")

    vezerles = ttk.Frame(oldal, padding=(22, 5))
    vezerles.pack(fill="x")

    ttk.Label(
        vezerles,
        text="Leltározás dátuma:"
    ).pack(side="left", padx=(0, 8))

    datum_valtozo = tk.StringVar(
        value=date.today().isoformat()
    )

    ttk.Entry(
        vezerles,
        textvariable=datum_valtozo,
        width=15
    ).pack(side="left", padx=(0, 10))

    adatok = None
    ful_keret = ttk.Frame(oldal)
    ful_keret.pack(
        fill="both",
        expand=True,
        padx=20,
        pady=10
    )

    def am_ujratoltes():
        nonlocal adatok

        try:
            datum = date.fromisoformat(
                datum_valtozo.get().strip()
            ).isoformat()

            uj_adatok = am_statisztika_adatok(datum)

        except Exception as hiba:
            messagebox.showerror(
                "Statisztikai hiba",
                str(hiba)
            )
            return

        adatok = uj_adatok

        for widget in ful_keret.winfo_children():
            widget.destroy()

        fulek = ttk.Notebook(ful_keret)
        fulek.pack(fill="both", expand=True)

        osszesito = ttk.Frame(fulek)
        helyiseg_ful = ttk.Frame(fulek)
        felelos_ful = ttk.Frame(fulek)
        leltar_ful = ttk.Frame(fulek)
        muvelet_ful = ttk.Frame(fulek)

        fulek.add(osszesito, text="Áttekintés")
        fulek.add(helyiseg_ful, text="Helyiségek")
        fulek.add(felelos_ful, text="Felelősök")
        fulek.add(leltar_ful, text="Leltározás")
        fulek.add(muvelet_ful, text="Műveletek")

        kartyak = ttk.Frame(osszesito, padding=15)
        kartyak.pack(fill="x")

        mutatok = [
            ("Aktív eszköz", str(adatok["aktiv"])),
            ("Archivált", str(adatok["archivalt"])),
            ("Elszállított", str(adatok["elszallitott"])),
            ("Bruttó érték", am_forint(adatok["brutto"])),
            ("Könyv szerinti érték", am_forint(adatok["konyv"])),
            ("Összes rekord", str(adatok["osszes"]))
        ]

        for index, (felirat, ertek) in enumerate(mutatok):
            keret = ttk.LabelFrame(
                kartyak,
                text=felirat,
                padding=12
            )
            keret.grid(
                row=index // 3,
                column=index % 3,
                sticky="nsew",
                padx=6,
                pady=6
            )

            ttk.Label(
                keret,
                text=ertek,
                font=("Segoe UI", 15, "bold")
            ).pack(pady=8)

        for index in range(3):
            kartyak.columnconfigure(index, weight=1)

        aktiv_statuszok = [
            ("Aktív", adatok["aktiv"]),
            ("Archivált", adatok["archivalt"]),
            ("Elszállított", adatok["elszallitott"])
        ]

        am_diagram(
            osszesito,
            aktiv_statuszok,
            "Eszközök aktuális nyilvántartási állapota"
        )

        am_diagram(
            helyiseg_ful,
            [
                (sor[0], sor[1])
                for sor in adatok["helyisegek"][:10]
            ],
            "A 10 legtöbb eszközzel rendelkező helyiség"
        )

        am_diagram(
            felelos_ful,
            [
                (sor[0], sor[1])
                for sor in adatok["felelosok"][:10]
            ],
            "A 10 legtöbb eszközzel rendelkező felelős"
        )

        am_diagram(
            leltar_ful,
            list(adatok["leltar"].items()),
            f"Leltározási állapotok – {adatok['datum']}"
        )

        muvelet_adatok = [
            ("Áthelyezések", adatok["athelyezesek"]),
            (
                "Leltározási bejegyzések",
                adatok["leltar_bejegyzesek"]
            ),
            (
                "Selejtezésre jelölt",
                adatok["selejtezes"]["jelolt"]
            ),
            (
                "Selejtezés jóváhagyva",
                adatok["selejtezes"]["jovahagyott"]
            ),
            (
                "Elszállított",
                adatok["selejtezes"]["elszallitva"]
            ),
            (
                "Selejtezés visszavonva",
                adatok["selejtezes"]["visszavont"]
            )
        ]

        am_statisztika_tabla(
            muvelet_ful,
            ("Művelet", "Darabszám"),
            muvelet_adatok
        )

        ttk.Label(
            osszesito,
            text=(
                f"Helyiségek: {adatok['helyisegek_szama']}    "
                f"Felelősök: {adatok['felelosok_szama']}    "
                f"Leltározási dátum: {adatok['datum']}"
            )
        ).pack(
            anchor="w",
            padx=20,
            pady=8
        )

    ttk.Button(
        vezerles,
        text="Statisztikák frissítése",
        command=am_ujratoltes
    ).pack(side="left", padx=5)

    def am_export():
        if adatok is None:
            messagebox.showwarning(
                "Nincs adat",
                "Először töltsd be a statisztikákat!"
            )
            return

        am_statisztika_excel(adatok)

    ttk.Button(
        vezerles,
        text="Excel-export",
        command=am_export
    ).pack(side="left", padx=5)

    ttk.Label(
        vezerles,
        text="Dátumformátum: ÉÉÉÉ-HH-NN"
    ).pack(side="right", padx=5)

    am_ujratoltes()
