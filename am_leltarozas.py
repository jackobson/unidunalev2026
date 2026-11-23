
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from am_database import AMAdatbazis
from am_selejtezes import am_selejtezes_tabla


AM_LELTAR_ALLAPOTOK = (
    "Nem ellenőrzött",
    "Megvan",
    "Hiányzik",
    "Sérült"
)


def am_leltar_tablak():
    am_selejtezes_tabla()

    db = AMAdatbazis()

    try:
        db.kurzor.execute("""
            CREATE TABLE IF NOT EXISTS leltarozas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                eszkoz_id INTEGER NOT NULL,
                allapot TEXT NOT NULL,
                datum TEXT NOT NULL,
                megjegyzes TEXT,
                FOREIGN KEY (eszkoz_id)
                    REFERENCES eszkozok(id)
            )
        """)

        db.kurzor.execute("""
            CREATE INDEX IF NOT EXISTS
                am_leltar_datum_eszkoz_idx
            ON leltarozas(datum, eszkoz_id, id)
        """)

        db.kapcsolat.commit()

    finally:
        db.am_bezaras()


def am_leltar_adatok(datum):
    db = AMAdatbazis()

    try:
        return db.kurzor.execute("""
            SELECT
                e.id,
                e.eszkozazonosito,
                e.megnevezes,
                e.gyartasi_szam,
                h.helyiseg_szam,
                f.teljes_nev,
                COALESCE(l.allapot, 'Nem ellenőrzött'),
                COALESCE(l.megjegyzes, ''),
                l.id,
                e.mennyiseg,
                e.mertekegyseg,
                e.brutto_ertek
            FROM eszkozok e
            LEFT JOIN helyisegek h
                ON h.id = e.helyiseg_id
            LEFT JOIN felelosok f
                ON f.id = e.felelos_id
            LEFT JOIN leltarozas l
                ON l.id = (
                    SELECT MAX(l2.id)
                    FROM leltarozas l2
                    WHERE l2.eszkoz_id = e.id
                        AND l2.datum = ?
                )
            WHERE e.archivalt = 0
                AND NOT EXISTS (
                    SELECT 1
                    FROM selejtezesek s
                    WHERE s.eszkoz_id = e.id
                        AND s.allapot = 'elszallitva'
                )
            ORDER BY e.eszkozazonosito, e.id
        """, (datum,)).fetchall()

    finally:
        db.am_bezaras()


def am_leltar_elozmenyek(eszkoz_id):
    db = AMAdatbazis()

    try:
        return db.kurzor.execute("""
            SELECT
                id,
                datum,
                allapot,
                megjegyzes
            FROM leltarozas
            WHERE eszkoz_id = ?
            ORDER BY datum DESC, id DESC
        """, (eszkoz_id,)).fetchall()

    finally:
        db.am_bezaras()


def am_leltar_rogzites(
    eszkoz_id, allapot, datum, megjegyzes
):
    if allapot not in AM_LELTAR_ALLAPOTOK[1:]:
        raise ValueError(
            "Érvénytelen leltározási állapot!"
        )

    datum = date.fromisoformat(datum).isoformat()

    db = AMAdatbazis()

    try:
        db.kurzor.execute("BEGIN IMMEDIATE")

        letezik = db.kurzor.execute("""
            SELECT e.id
            FROM eszkozok e
            WHERE e.id = ?
                AND e.archivalt = 0
                AND NOT EXISTS (
                    SELECT 1
                    FROM selejtezesek s
                    WHERE s.eszkoz_id = e.id
                        AND s.allapot = 'elszallitva'
                )
        """, (eszkoz_id,)).fetchone()

        if letezik is None:
            raise ValueError(
                "Az eszköz nem aktív vagy már elszállították!"
            )

        db.kurzor.execute("""
            INSERT INTO leltarozas (
                eszkoz_id,
                allapot,
                datum,
                megjegyzes
            )
            VALUES (?, ?, ?, ?)
        """, (
            eszkoz_id,
            allapot,
            datum,
            megjegyzes
        ))

        db.kapcsolat.commit()

    except Exception:
        db.kapcsolat.rollback()
        raise

    finally:
        db.am_bezaras()


def am_leltar_export(datum):
    try:
        datum = date.fromisoformat(datum).isoformat()
        adatok = am_leltar_adatok(datum)

        utvonal = filedialog.asksaveasfilename(
            title="Leltározási Excel mentése",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel-munkafüzet", "*.xlsx")
            ],
            initialfile=f"Leltarozas_{datum}.xlsx"
        )

        if not utvonal:
            return

        munkafuzet = Workbook()
        lap = munkafuzet.active
        lap.title = "Leltározás"

        lap.append([
            "Eszközazonosító",
            "Megnevezés",
            "Gyártási szám",
            "Helyiség",
            "Felelős",
            "Mennyiség",
            "Mértékegység",
            "Bruttó érték",
            "Leltározási állapot",
            "Megjegyzés",
            "Leltározás dátuma"
        ])

        for sor in adatok:
            lap.append([
                sor[1] or "",
                sor[2] or "",
                sor[3] or "",
                sor[4] or "",
                sor[5] or "",
                sor[9],
                sor[10] or "",
                sor[11],
                sor[6],
                sor[7],
                datum
            ])

        for cella in lap[1]:
            cella.font = Font(
                color="FFFFFF",
                bold=True
            )
            cella.fill = PatternFill(
                "solid",
                fgColor="243E57"
            )
            cella.alignment = Alignment(
                wrap_text=True
            )

        szelessegek = [
            20, 42, 22, 17, 25,
            14, 17, 19, 24, 45, 22
        ]

        for index, szelesseg in enumerate(
            szelessegek, 1
        ):
            lap.column_dimensions[
                get_column_letter(index)
            ].width = szelesseg

        lap.freeze_panes = "A2"
        lap.auto_filter.ref = lap.dimensions

        osszesito = munkafuzet.create_sheet(
            "Összesítés"
        )

        osszesito.append([
            "Leltározás dátuma", datum
        ])
        osszesito.append([])
        osszesito.append([
            "Állapot", "Darabszám"
        ])

        for allapot in AM_LELTAR_ALLAPOTOK:
            darab = sum(
                1 for sor in adatok
                if sor[6] == allapot
            )

            osszesito.append([
                allapot, darab
            ])

        osszesito.append([])
        osszesito.append([
            "Összes aktív eszköz",
            len(adatok)
        ])

        osszesito.column_dimensions["A"].width = 32
        osszesito.column_dimensions["B"].width = 22

        for cella in osszesito[3]:
            cella.font = Font(
                color="FFFFFF",
                bold=True
            )
            cella.fill = PatternFill(
                "solid",
                fgColor="243E57"
            )

        munkafuzet.save(utvonal)

        messagebox.showinfo(
            "Sikeres exportálás",
            "A leltározási Excel-fájl elkészült!"
        )

    except Exception as hiba:
        messagebox.showerror(
            "Exportálási hiba",
            str(hiba)
        )


def am_leltar_eszkoz_ablak(
    content, eszkoz, datum, frissites
):
    eszkoz_id = eszkoz[0]

    ablak = tk.Toplevel(content)
    ablak.title("AM – Eszköz leltározása")
    ablak.geometry("760x670")
    ablak.minsize(700, 600)
    ablak.transient(content.winfo_toplevel())
    ablak.grab_set()

    fo = ttk.Frame(
        ablak,
        padding=22
    )
    fo.pack(
        fill="both",
        expand=True
    )

    ttk.Label(
        fo,
        text="Eszköz leltározása",
        font=("Segoe UI", 17, "bold")
    ).pack(
        anchor="w",
        pady=(0, 15)
    )

    adatkeret = ttk.LabelFrame(
        fo,
        text="Eszköz adatai",
        padding=14
    )
    adatkeret.pack(
        fill="x",
        pady=(0, 15)
    )

    brutto = eszkoz[11]

    if isinstance(brutto, (int, float)):
        brutto_szoveg = (
            f"{brutto:,.0f} Ft".replace(",", " ")
        )
    else:
        brutto_szoveg = ""

    eszkozadatok = [
        ("Eszközazonosító", eszkoz[1]),
        ("Megnevezés", eszkoz[2]),
        ("Gyártási szám", eszkoz[3]),
        ("Helyiség", eszkoz[4]),
        ("Felelős", eszkoz[5]),
        (
            "Mennyiség",
            f"{eszkoz[9] or ''} {eszkoz[10] or ''}"
        ),
        ("Bruttó érték", brutto_szoveg),
        ("Leltározás dátuma", datum)
    ]

    for index, (felirat, ertek) in enumerate(
        eszkozadatok
    ):
        oszlop = (index // 4) * 2
        sor = index % 4

        ttk.Label(
            adatkeret,
            text=f"{felirat}:",
            font=("Segoe UI", 9, "bold")
        ).grid(
            row=sor,
            column=oszlop,
            sticky="w",
            padx=(5, 10),
            pady=7
        )

        ttk.Label(
            adatkeret,
            text="" if ertek is None else str(ertek),
            wraplength=225
        ).grid(
            row=sor,
            column=oszlop + 1,
            sticky="w",
            padx=(0, 15),
            pady=7
        )

    adatkeret.columnconfigure(1, weight=1)
    adatkeret.columnconfigure(3, weight=1)

    ellenorzeskeret = ttk.LabelFrame(
        fo,
        text="Leltározási eredmény",
        padding=14
    )
    ellenorzeskeret.pack(
        fill="x",
        pady=(0, 15)
    )

    aktualis = eszkoz[6]

    if aktualis in AM_LELTAR_ALLAPOTOK[1:]:
        alap_allapot = aktualis
    else:
        alap_allapot = "Megvan"

    allapot_valtozo = tk.StringVar(
        value=alap_allapot
    )

    for allapot, leiras in [
        (
            "Megvan",
            "Az eszköz megtalálható"
        ),
        (
            "Hiányzik",
            "Az eszköz nem található"
        ),
        (
            "Sérült",
            "Az eszköz sérült vagy hibás"
        )
    ]:
        ttk.Radiobutton(
            ellenorzeskeret,
            text=f"{allapot} – {leiras}",
            variable=allapot_valtozo,
            value=allapot
        ).pack(
            anchor="w",
            pady=4
        )

    ttk.Label(
        ellenorzeskeret,
        text="Megjegyzés:"
    ).pack(
        anchor="w",
        pady=(12, 4)
    )

    megjegyzes_valtozo = tk.StringVar(
        value=eszkoz[7] or ""
    )

    ttk.Entry(
        ellenorzeskeret,
        textvariable=megjegyzes_valtozo,
        width=75
    ).pack(
        fill="x",
        pady=(0, 4)
    )

    ttk.Label(
        ellenorzeskeret,
        text=(
            "Hiányzó vagy sérült eszköznél "
            "a megjegyzés kitöltése kötelező."
        )
    ).pack(
        anchor="w",
        pady=(3, 0)
    )

    elozmenykeret = ttk.LabelFrame(
        fo,
        text="Korábbi leltározási bejegyzések",
        padding=10
    )
    elozmenykeret.pack(
        fill="both",
        expand=True,
        pady=(0, 15)
    )

    oszlopok = (
        "datum",
        "allapot",
        "megjegyzes"
    )

    elozmeny_tabla = ttk.Treeview(
        elozmenykeret,
        columns=oszlopok,
        show="headings",
        height=5
    )

    for kulcs, felirat, szelesseg in [
        ("datum", "Dátum", 115),
        ("allapot", "Állapot", 130),
        ("megjegyzes", "Megjegyzés", 400)
    ]:
        elozmeny_tabla.heading(
            kulcs, text=felirat
        )
        elozmeny_tabla.column(
            kulcs, width=szelesseg
        )

    gorgeto = ttk.Scrollbar(
        elozmenykeret,
        orient="vertical",
        command=elozmeny_tabla.yview
    )

    elozmeny_tabla.configure(
        yscrollcommand=gorgeto.set
    )

    gorgeto.pack(
        side="right",
        fill="y"
    )
    elozmeny_tabla.pack(
        side="left",
        fill="both",
        expand=True
    )

    for bejegyzes in am_leltar_elozmenyek(
        eszkoz_id
    ):
        elozmeny_tabla.insert(
            "",
            "end",
            values=(
                bejegyzes[1],
                bejegyzes[2],
                bejegyzes[3] or ""
            )
        )

    def am_mentes():
        allapot = allapot_valtozo.get()
        megjegyzes = (
            megjegyzes_valtozo.get().strip()
        )

        if (
            allapot in ("Hiányzik", "Sérült")
            and not megjegyzes
        ):
            messagebox.showwarning(
                "Hiányzó megjegyzés",
                "Hiányzó vagy sérült eszköznél "
                "kötelező megjegyzést megadni!",
                parent=ablak
            )
            return

        if aktualis != "Nem ellenőrzött":
            if not messagebox.askyesno(
                "Korábbi ellenőrzés javítása",
                "Ezt az eszközt ezen a napon "
                "már ellenőrizted.\n\n"
                "Az új bejegyzés lesz érvényes, "
                "de a korábbi is megmarad "
                "az előzményekben.\n\n"
                "Folytatod?",
                parent=ablak
            ):
                return

        try:
            am_leltar_rogzites(
                eszkoz_id,
                allapot,
                datum,
                megjegyzes
            )

        except Exception as hiba:
            messagebox.showerror(
                "Mentési hiba",
                str(hiba),
                parent=ablak
            )
            return

        ablak.destroy()
        frissites()

        messagebox.showinfo(
            "Sikeres mentés",
            "A leltározási eredményt elmentettük!"
        )

    gombkeret = ttk.Frame(fo)
    gombkeret.pack(fill="x")

    ttk.Button(
        gombkeret,
        text="Leltározás mentése",
        command=am_mentes
    ).pack(
        side="left",
        padx=(0, 10)
    )

    ttk.Button(
        gombkeret,
        text="Mégse",
        command=ablak.destroy
    ).pack(side="left")


def am_leltarozas_megjelenitese(content):
    am_leltar_tablak()

    oldal = ttk.Frame(content)
    oldal.pack(
        fill="both",
        expand=True
    )

    datum_valtozo = tk.StringVar(
        value=date.today().isoformat()
    )

    aktiv_datum = date.today().isoformat()

    kereses = tk.StringVar()
    allapot_szuro = tk.StringVar(
        value="Összes"
    )
    helyiseg_szuro = tk.StringVar(
        value="Összes"
    )

    lista = []
    rekordok = {}

    fejlec = ttk.Frame(oldal)
    fejlec.pack(
        fill="x",
        padx=25,
        pady=10
    )

    ttk.Label(
        fejlec,
        text="Eszközök leltározása",
        font=("Segoe UI", 16, "bold")
    ).pack(anchor="w")

    ttk.Label(
        fejlec,
        text=(
            "Kattints duplán egy eszközre "
            "a leltározási ablak megnyitásához."
        )
    ).pack(
        anchor="w",
        pady=(5, 0)
    )

    szurokeret = ttk.Frame(oldal)
    szurokeret.pack(
        fill="x",
        padx=25,
        pady=10
    )

    ttk.Label(
        szurokeret,
        text="Leltározás dátuma:"
    ).grid(
        row=0,
        column=0,
        padx=5,
        pady=6,
        sticky="w"
    )

    ttk.Entry(
        szurokeret,
        textvariable=datum_valtozo,
        width=16
    ).grid(
        row=0,
        column=1,
        padx=5,
        pady=6,
        sticky="w"
    )

    ttk.Label(
        szurokeret,
        text="ÉÉÉÉ-HH-NN"
    ).grid(
        row=0,
        column=2,
        padx=5
    )

    ttk.Label(
        szurokeret,
        text="Keresés:"
    ).grid(
        row=1,
        column=0,
        padx=5,
        pady=6,
        sticky="w"
    )

    ttk.Entry(
        szurokeret,
        textvariable=kereses,
        width=35
    ).grid(
        row=1,
        column=1,
        columnspan=2,
        padx=5,
        pady=6,
        sticky="w"
    )

    ttk.Label(
        szurokeret,
        text="Állapot:"
    ).grid(
        row=1,
        column=3,
        padx=5,
        pady=6
    )

    ttk.Combobox(
        szurokeret,
        textvariable=allapot_szuro,
        values=(
            "Összes",
        ) + AM_LELTAR_ALLAPOTOK,
        state="readonly",
        width=20
    ).grid(
        row=1,
        column=4,
        padx=5,
        pady=6
    )

    ttk.Label(
        szurokeret,
        text="Helyiség:"
    ).grid(
        row=2,
        column=0,
        padx=5,
        pady=6,
        sticky="w"
    )

    helyiseg_combo = ttk.Combobox(
        szurokeret,
        textvariable=helyiseg_szuro,
        state="readonly",
        width=23
    )

    helyiseg_combo.grid(
        row=2,
        column=1,
        padx=5,
        pady=6,
        sticky="w"
    )

    osszesites = ttk.Frame(oldal)
    osszesites.pack(
        fill="x",
        padx=25,
        pady=12
    )

    szamlalok = {}

    for felirat in (
        "Összes",
        "Nem ellenőrzött",
        "Megvan",
        "Hiányzik",
        "Sérült"
    ):
        valtozo = tk.StringVar(
            value=f"{felirat}: 0"
        )

        ttk.Label(
            osszesites,
            textvariable=valtozo,
            padding=(8, 7)
        ).pack(
            side="left",
            padx=5
        )

        szamlalok[felirat] = valtozo

    tablakeret = ttk.Frame(oldal)
    tablakeret.pack(
        fill="both",
        expand=True,
        padx=25,
        pady=10
    )

    oszlopok = (
        "azonosito",
        "megnevezes",
        "gyartasi",
        "helyiseg",
        "felelos",
        "allapot"
    )

    tabla = ttk.Treeview(
        tablakeret,
        columns=oszlopok,
        show="headings",
        selectmode="browse"
    )

    fejlecek = [
        (
            "azonosito",
            "Eszközazonosító",
            125
        ),
        (
            "megnevezes",
            "Megnevezés",
            245
        ),
        (
            "gyartasi",
            "Gyártási szám",
            135
        ),
        (
            "helyiseg",
            "Helyiség",
            90
        ),
        (
            "felelos",
            "Felelős",
            145
        ),
        (
            "allapot",
            "Leltározási állapot",
            150
        )
    ]

    for kulcs, felirat, szelesseg in fejlecek:
        tabla.heading(
            kulcs, text=felirat
        )
        tabla.column(
            kulcs, width=szelesseg
        )

    gorgeto = ttk.Scrollbar(
        tablakeret,
        orient="vertical",
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
        kijeloles = tabla.selection()
        regi_id = (
            kijeloles[0]
            if kijeloles else None
        )

        for elem in tabla.get_children():
            tabla.delete(elem)

        keresett = (
            kereses.get().casefold().strip()
        )

        kert_allapot = allapot_szuro.get()
        kert_helyiseg = helyiseg_szuro.get()

        darab = 0

        for sor in lista:
            szoveg = " ".join(
                str(ertek or "")
                for ertek in sor[1:6]
            ).casefold()

            if keresett not in szoveg:
                continue

            if (
                kert_allapot != "Összes"
                and sor[6] != kert_allapot
            ):
                continue

            if (
                kert_helyiseg != "Összes"
                and str(sor[4] or "") != kert_helyiseg
            ):
                continue

            tabla.insert(
                "",
                "end",
                iid=str(sor[0]),
                values=(
                    sor[1] or "",
                    sor[2] or "",
                    sor[3] or "",
                    sor[4] or "",
                    sor[5] or "",
                    sor[6]
                )
            )

            darab += 1

        talalatok.set(
            f"Találatok: {darab} / {len(lista)}"
        )

        if regi_id and tabla.exists(regi_id):
            tabla.selection_set(regi_id)
            tabla.see(regi_id)

    def am_frissites():
        nonlocal lista, rekordok

        if not oldal.winfo_exists():
            return

        lista = am_leltar_adatok(aktiv_datum)

        rekordok = {
            sor[0]: sor
            for sor in lista
        }

        helyisegek = sorted({
            str(sor[4])
            for sor in lista
            if sor[4] is not None
        })

        helyiseg_ertekek = [
            "Összes"
        ] + helyisegek

        helyiseg_combo.configure(
            values=helyiseg_ertekek
        )

        if (
            helyiseg_szuro.get()
            not in helyiseg_ertekek
        ):
            helyiseg_szuro.set(
                "Összes"
            )

        szamlalok["Összes"].set(
            f"Összes: {len(lista)}"
        )

        for allapot in AM_LELTAR_ALLAPOTOK:
            darab = sum(
                1 for sor in lista
                if sor[6] == allapot
            )

            szamlalok[allapot].set(
                f"{allapot}: {darab}"
            )

        am_szures()

    def am_datum_betoltese():
        nonlocal aktiv_datum

        try:
            uj_datum = date.fromisoformat(
                datum_valtozo.get().strip()
            ).isoformat()

        except ValueError:
            messagebox.showerror(
                "Hibás dátum",
                "A dátum formátuma ÉÉÉÉ-HH-NN legyen!"
            )
            return

        aktiv_datum = uj_datum
        am_frissites()

    ttk.Button(
        szurokeret,
        text="Nap betöltése",
        command=am_datum_betoltese
    ).grid(
        row=0,
        column=3,
        padx=5,
        pady=6
    )

    def am_eszkoz_megnyitas(event=None):
        kijeloles = tabla.selection()

        if not kijeloles:
            messagebox.showwarning(
                "Nincs kiválasztás",
                "Először válassz ki egy eszközt!"
            )
            return

        if (
            datum_valtozo.get().strip()
            != aktiv_datum
        ):
            messagebox.showwarning(
                "Dátum megváltozott",
                "Először kattints a Nap betöltése gombra!"
            )
            return

        eszkoz_id = int(kijeloles[0])
        eszkoz = rekordok.get(eszkoz_id)

        if eszkoz is None:
            messagebox.showerror(
                "Hiba",
                "Az eszköz nem található!"
            )
            return

        am_leltar_eszkoz_ablak(
            content,
            eszkoz,
            aktiv_datum,
            am_frissites
        )

    tabla.bind(
        "<Double-1>",
        am_eszkoz_megnyitas
    )

    gombkeret = ttk.Frame(oldal)
    gombkeret.pack(
        fill="x",
        padx=25,
        pady=(5, 15)
    )

    ttk.Button(
        gombkeret,
        text="Kijelölt eszköz leltározása",
        command=am_eszkoz_megnyitas
    ).pack(
        side="left",
        padx=5
    )

    ttk.Button(
        gombkeret,
        text="Frissítés",
        command=am_frissites
    ).pack(
        side="left",
        padx=5
    )

    ttk.Button(
        gombkeret,
        text="Excel-export",
        command=lambda: am_leltar_export(
            aktiv_datum
        )
    ).pack(
        side="left",
        padx=5
    )

    kereses.trace_add(
        "write",
        am_szures
    )

    allapot_szuro.trace_add(
        "write",
        am_szures
    )

    helyiseg_szuro.trace_add(
        "write",
        am_szures
    )

    am_frissites()
