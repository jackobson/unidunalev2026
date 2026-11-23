
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from am_database import AMAdatbazis


AM_LELTAR_ALLAPOTOK = (
    "Nem ellenőrzött",
    "Megvan",
    "Hiányzik",
    "Sérült"
)


def am_leltar_tablak():
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
                l.id
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


def am_leltar_rogzites(eszkoz_id, allapot, datum, megjegyzes):
    if allapot not in AM_LELTAR_ALLAPOTOK[1:]:
        raise ValueError("Érvénytelen leltározási állapot!")

    datum = date.fromisoformat(datum).isoformat()

    db = AMAdatbazis()

    try:
        db.kurzor.execute("BEGIN IMMEDIATE")

        eszkoz = db.kurzor.execute("""
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

        if eszkoz is None:
            raise ValueError(
                "Az eszköz már nem aktív vagy nem található!"
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


def am_leltar_elozmenyek(eszkoz_id):
    db = AMAdatbazis()

    try:
        return db.kurzor.execute("""
            SELECT id, datum, allapot, megjegyzes
            FROM leltarozas
            WHERE eszkoz_id = ?
            ORDER BY datum DESC, id DESC
        """, (eszkoz_id,)).fetchall()

    finally:
        db.am_bezaras()


def am_leltar_export(datum):
    try:
        datum = date.fromisoformat(datum).isoformat()
        adatok = am_leltar_adatok(datum)

        utvonal = filedialog.asksaveasfilename(
            title="Leltározási eredmény mentése",
            initialfile=f"Leltarozas_{datum}.xlsx",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel-munkafüzet", "*.xlsx")
            ]
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
            "Állapot",
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
            20, 45, 22, 17, 27, 22, 50, 22
        ]

        for index, szelesseg in enumerate(szelessegek, 1):
            lap.column_dimensions[
                get_column_letter(index)
            ].width = szelesseg

        lap.freeze_panes = "A2"
        lap.auto_filter.ref = lap.dimensions

        osszesito = munkafuzet.create_sheet("Összesítés")

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

        osszesito.column_dimensions["A"].width = 30
        osszesito.column_dimensions["B"].width = 25

        for cella in osszesito[3]:
            cella.font = Font(
                bold=True,
                color="FFFFFF"
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


def am_leltar_elozmeny_ablak(szulo, eszkoz_id, eszkoz_nev):
    ablak = tk.Toplevel(szulo)
    ablak.title("Leltározási előzmények")
    ablak.geometry("850x450")
    ablak.transient(szulo)
    ablak.grab_set()

    ttk.Label(
        ablak,
        text=f"Leltározási előzmények – {eszkoz_nev}",
        font=("Segoe UI", 13, "bold")
    ).pack(
        anchor="w",
        padx=15,
        pady=15
    )

    keret = ttk.Frame(ablak, padding=10)
    keret.pack(fill="both", expand=True)

    tabla = ttk.Treeview(
        keret,
        columns=(
            "sorszam",
            "datum",
            "allapot",
            "megjegyzes"
        ),
        show="headings"
    )

    oszlopok = [
        ("sorszam", "Bejegyzés", 100),
        ("datum", "Dátum", 130),
        ("allapot", "Állapot", 150),
        ("megjegyzes", "Megjegyzés", 400)
    ]

    for kulcs, nev, szelesseg in oszlopok:
        tabla.heading(kulcs, text=nev)
        tabla.column(kulcs, width=szelesseg)

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

    for sor in am_leltar_elozmenyek(eszkoz_id):
        tabla.insert(
            "",
            "end",
            values=(
                sor[0],
                sor[1],
                sor[2],
                sor[3] or ""
            )
        )

    ttk.Button(
        ablak,
        text="Bezárás",
        command=ablak.destroy
    ).pack(pady=10)


def am_leltarozas_megjelenitese(content):
    am_leltar_tablak()

    oldal = ttk.Frame(content)
    oldal.pack(fill="both", expand=True)

    fejlec = ttk.Frame(oldal)
    fejlec.pack(
        fill="x",
        padx=25,
        pady=(5, 12)
    )

    ttk.Label(
        fejlec,
        text="Leltározás",
        font=("Segoe UI", 16, "bold")
    ).pack(side="left")

    datum_valtozo = tk.StringVar(
        value=date.today().isoformat()
    )

    aktiv_datum = date.today().isoformat()

    lista = []
    rekordok = {}

    szurokeret = ttk.Frame(oldal)
    szurokeret.pack(
        fill="x",
        padx=25,
        pady=6
    )

    ttk.Label(
        szurokeret,
        text="Leltározás dátuma:"
    ).grid(
        row=0,
        column=0,
        padx=5,
        pady=5,
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
        pady=5
    )

    ttk.Label(
        szurokeret,
        text="ÉÉÉÉ-HH-NN"
    ).grid(
        row=0,
        column=2,
        padx=5,
        pady=5
    )

    kereses = tk.StringVar()
    allapot_szuro = tk.StringVar(
        value="Összes"
    )
    helyiseg_szuro = tk.StringVar(
        value="Összes"
    )

    ttk.Label(
        szurokeret,
        text="Keresés:"
    ).grid(
        row=1,
        column=0,
        padx=5,
        pady=5,
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
        pady=5,
        sticky="w"
    )

    ttk.Label(
        szurokeret,
        text="Állapot:"
    ).grid(
        row=1,
        column=3,
        padx=5,
        pady=5
    )

    ttk.Combobox(
        szurokeret,
        textvariable=allapot_szuro,
        values=("Összes",) + AM_LELTAR_ALLAPOTOK,
        state="readonly",
        width=20
    ).grid(
        row=1,
        column=4,
        padx=5,
        pady=5
    )

    ttk.Label(
        szurokeret,
        text="Helyiség:"
    ).grid(
        row=2,
        column=0,
        padx=5,
        pady=5,
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
        pady=5,
        sticky="w"
    )

    osszesites = ttk.Frame(oldal)
    osszesites.pack(
        fill="x",
        padx=25,
        pady=12
    )

    szamlalok = {}

    for szoveg in (
        "Összes",
        "Nem ellenőrzött",
        "Megvan",
        "Hiányzik",
        "Sérült"
    ):
        valtozo = tk.StringVar(
            value=f"{szoveg}: 0"
        )

        ttk.Label(
            osszesites,
            textvariable=valtozo,
            padding=(10, 7)
        ).pack(side="left", padx=5)

        szamlalok[szoveg] = valtozo

    tablakeret = ttk.Frame(oldal)
    tablakeret.pack(
        fill="both",
        expand=True,
        padx=25,
        pady=5
    )

    oszlopok = (
        "azonosito",
        "nev",
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
        ("azonosito", "Eszközazonosító", 125),
        ("nev", "Megnevezés", 245),
        ("gyartasi", "Gyártási szám", 140),
        ("helyiseg", "Helyiség", 90),
        ("felelos", "Felelős", 150),
        ("allapot", "Leltározási állapot", 150)
    ]

    for kulcs, nev, szelesseg in fejlecek:
        tabla.heading(kulcs, text=nev)
        tabla.column(kulcs, width=szelesseg)

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

    adatkeret = ttk.LabelFrame(
        oldal,
        text="Kiválasztott eszköz ellenőrzése",
        padding=12
    )

    adatkeret.pack(
        fill="x",
        padx=25,
        pady=10
    )

    eszkoz_felirat = tk.StringVar(
        value="Nincs kiválasztott eszköz"
    )

    ttk.Label(
        adatkeret,
        textvariable=eszkoz_felirat
    ).grid(
        row=0,
        column=0,
        columnspan=4,
        sticky="w",
        pady=8
    )

    uj_allapot = tk.StringVar(
        value="Megvan"
    )

    ttk.Label(
        adatkeret,
        text="Állapot:"
    ).grid(
        row=1,
        column=0,
        padx=5,
        pady=5,
        sticky="w"
    )

    ttk.Combobox(
        adatkeret,
        textvariable=uj_allapot,
        values=AM_LELTAR_ALLAPOTOK[1:],
        state="readonly",
        width=20
    ).grid(
        row=1,
        column=1,
        padx=5,
        pady=5,
        sticky="w"
    )

    megjegyzes = tk.StringVar()

    ttk.Label(
        adatkeret,
        text="Megjegyzés:"
    ).grid(
        row=2,
        column=0,
        padx=5,
        pady=5,
        sticky="w"
    )

    ttk.Entry(
        adatkeret,
        textvariable=megjegyzes,
        width=65
    ).grid(
        row=2,
        column=1,
        columnspan=3,
        padx=5,
        pady=5,
        sticky="ew"
    )

    adatkeret.columnconfigure(3, weight=1)

    def am_kivalasztott():
        kijeloles = tabla.selection()

        if not kijeloles:
            return None

        return rekordok.get(
            int(kijeloles[0])
        )

    def am_kijeloles_valtozas(event=None):
        sor = am_kivalasztott()

        if sor is None:
            eszkoz_felirat.set(
                "Nincs kiválasztott eszköz"
            )
            megjegyzes.set("")
            uj_allapot.set("Megvan")
            return

        eszkoz_felirat.set(
            f"{sor[1]} – {sor[2]}"
        )

        if sor[6] in AM_LELTAR_ALLAPOTOK[1:]:
            uj_allapot.set(sor[6])
        else:
            uj_allapot.set("Megvan")

        megjegyzes.set(sor[7] or "")

    def am_szures(*args):
        kijeloles = tabla.selection()
        korabbi_id = kijeloles[0] if kijeloles else None

        for elem in tabla.get_children():
            tabla.delete(elem)

        keresett = kereses.get().casefold().strip()
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

        if korabbi_id and tabla.exists(korabbi_id):
            tabla.selection_set(korabbi_id)
            tabla.see(korabbi_id)
        else:
            am_kijeloles_valtozas()

    def am_adatok_frissitese():
        nonlocal lista, rekordok

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

        helyiseg_combo.configure(
            values=["Összes"] + helyisegek
        )

        if helyiseg_szuro.get() not in (
            ["Összes"] + helyisegek
        ):
            helyiseg_szuro.set("Összes")

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
        am_adatok_frissitese()

    ttk.Button(
        szurokeret,
        text="Nap betöltése",
        command=am_datum_betoltese
    ).grid(
        row=0,
        column=3,
        padx=5,
        pady=5
    )

    def am_mentes():
        sor = am_kivalasztott()

        if sor is None:
            messagebox.showwarning(
                "Nincs kiválasztás",
                "Először jelölj ki egy eszközt!"
            )
            return

        if datum_valtozo.get().strip() != aktiv_datum:
            messagebox.showwarning(
                "Dátum változott",
                "Először kattints a Nap betöltése gombra!"
            )
            return

        allapot = uj_allapot.get()
        szoveg = megjegyzes.get().strip()

        if allapot in ("Hiányzik", "Sérült") and not szoveg:
            messagebox.showwarning(
                "Hiányzó megjegyzés",
                "Hiányzó vagy sérült eszköznél "
                "kötelező megjegyzést írni!"
            )
            return

        if sor[6] != "Nem ellenőrzött":
            if not messagebox.askyesno(
                "Korábbi rögzítés javítása",
                "Ezt az eszközt ezen a napon már "
                "ellenőrizted.\n\n"
                "Az új bejegyzés lesz az érvényes, "
                "de a korábbi megmarad az előzményekben.\n\n"
                "Folytatod?"
            ):
                return

        try:
            am_leltar_rogzites(
                sor[0],
                allapot,
                aktiv_datum,
                szoveg
            )
        except Exception as hiba:
            messagebox.showerror(
                "Mentési hiba",
                str(hiba)
            )
            return

        am_adatok_frissitese()

        messagebox.showinfo(
            "Sikeres mentés",
            "A leltározási adatokat elmentettük!"
        )

    def am_elozmenyek():
        sor = am_kivalasztott()

        if sor is None:
            messagebox.showwarning(
                "Nincs kiválasztás",
                "Először jelölj ki egy eszközt!"
            )
            return

        am_leltar_elozmeny_ablak(
            content.winfo_toplevel(),
            sor[0],
            f"{sor[1]} – {sor[2]}"
        )

    gombkeret = ttk.Frame(oldal)
    gombkeret.pack(
        fill="x",
        padx=25,
        pady=(0, 15)
    )

    ttk.Button(
        gombkeret,
        text="Ellenőrzés mentése",
        command=am_mentes
    ).pack(
        side="left",
        padx=5
    )

    ttk.Button(
        gombkeret,
        text="Előzmények",
        command=am_elozmenyek
    ).pack(
        side="left",
        padx=5
    )

    ttk.Button(
        gombkeret,
        text="Frissítés",
        command=am_adatok_frissitese
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

    tabla.bind(
        "<<TreeviewSelect>>",
        am_kijeloles_valtozas
    )

    kereses.trace_add(
        "write", am_szures
    )

    allapot_szuro.trace_add(
        "write", am_szures
    )

    helyiseg_szuro.trace_add(
        "write", am_szures
    )

    am_adatok_frissitese()
