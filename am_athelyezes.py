
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from am_database import AMAdatbazis
from am_selejtezes import am_selejtezes_tabla


def am_athelyezes_elokeszites():
    am_selejtezes_tabla()

    db = AMAdatbazis()
    try:
        db.kurzor.execute("""
            CREATE TABLE IF NOT EXISTS athelyezesek (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                eszkoz_id INTEGER NOT NULL,
                regi_helyiseg_id INTEGER,
                uj_helyiseg_id INTEGER,
                regi_felelos_id INTEGER,
                uj_felelos_id INTEGER,
                datum TEXT NOT NULL,
                megjegyzes TEXT,
                FOREIGN KEY (eszkoz_id) REFERENCES eszkozok(id),
                FOREIGN KEY (regi_helyiseg_id) REFERENCES helyisegek(id),
                FOREIGN KEY (uj_helyiseg_id) REFERENCES helyisegek(id),
                FOREIGN KEY (regi_felelos_id) REFERENCES felelosok(id),
                FOREIGN KEY (uj_felelos_id) REFERENCES felelosok(id)
            )
        """)
        db.kapcsolat.commit()
    finally:
        db.am_bezaras()


def am_athelyezheto_eszkozok():
    db = AMAdatbazis()
    try:
        return db.kurzor.execute("""
            SELECT
                e.id,
                e.eszkozazonosito,
                e.megnevezes,
                e.gyartasi_szam,
                e.helyiseg_id,
                e.felelos_id,
                h.telephely,
                h.helyiseg_szam,
                h.megnevezes,
                f.teljes_nev,
                f.sztsz
            FROM eszkozok e
            LEFT JOIN helyisegek h
                ON h.id = e.helyiseg_id
            LEFT JOIN felelosok f
                ON f.id = e.felelos_id
            WHERE e.archivalt = 0
                AND NOT EXISTS (
                    SELECT 1
                    FROM selejtezesek s
                    WHERE s.eszkoz_id = e.id
                        AND s.allapot = 'elszallitva'
                )
            ORDER BY e.eszkozazonosito, e.id
        """).fetchall()
    finally:
        db.am_bezaras()


def am_athelyezes_elozmenyek(eszkoz_id):
    db = AMAdatbazis()
    try:
        return db.kurzor.execute("""
            SELECT
                a.id,
                a.datum,
                rh.helyiseg_szam,
                uh.helyiseg_szam,
                rf.teljes_nev,
                uf.teljes_nev,
                a.megjegyzes
            FROM athelyezesek a
            LEFT JOIN helyisegek rh
                ON rh.id = a.regi_helyiseg_id
            LEFT JOIN helyisegek uh
                ON uh.id = a.uj_helyiseg_id
            LEFT JOIN felelosok rf
                ON rf.id = a.regi_felelos_id
            LEFT JOIN felelosok uf
                ON uf.id = a.uj_felelos_id
            WHERE a.eszkoz_id = ?
            ORDER BY a.datum DESC, a.id DESC
        """, (eszkoz_id,)).fetchall()
    finally:
        db.am_bezaras()


def am_athelyezes_mentes(
    eszkoz_id, uj_helyiseg_id, uj_felelos_id,
    datum, megjegyzes
):
    datum = date.fromisoformat(datum).isoformat()

    db = AMAdatbazis()

    try:
        db.kurzor.execute("BEGIN IMMEDIATE")

        eszkoz = db.kurzor.execute("""
            SELECT
                e.helyiseg_id,
                e.felelos_id
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
                "Az eszköz nem található vagy nem aktív!"
            )

        regi_helyiseg_id, regi_felelos_id = eszkoz

        if (
            regi_helyiseg_id == uj_helyiseg_id
            and regi_felelos_id == uj_felelos_id
        ):
            raise ValueError(
                "Nem történt változás a helyiségben "
                "vagy a felelős személyében."
            )

        helyiseg = db.kurzor.execute("""
            SELECT telephely
            FROM helyisegek
            WHERE id = ?
        """, (uj_helyiseg_id,)).fetchone()

        felelos = db.kurzor.execute("""
            SELECT id
            FROM felelosok
            WHERE id = ?
        """, (uj_felelos_id,)).fetchone()

        if helyiseg is None or felelos is None:
            raise ValueError(
                "A kiválasztott helyiség vagy felelős nem létezik!"
            )

        db.kurzor.execute("""
            UPDATE eszkozok
            SET
                helyiseg_id = ?,
                felelos_id = ?,
                eszkoz_telephelye = ?
            WHERE id = ?
                AND archivalt = 0
        """, (
            uj_helyiseg_id,
            uj_felelos_id,
            helyiseg[0],
            eszkoz_id
        ))

        if db.kurzor.rowcount != 1:
            raise ValueError(
                "Az eszköz módosítása nem sikerült!"
            )

        db.kurzor.execute("""
            INSERT INTO athelyezesek (
                eszkoz_id,
                regi_helyiseg_id,
                uj_helyiseg_id,
                regi_felelos_id,
                uj_felelos_id,
                datum,
                megjegyzes
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            eszkoz_id,
            regi_helyiseg_id,
            uj_helyiseg_id,
            regi_felelos_id,
            uj_felelos_id,
            datum,
            megjegyzes
        ))

        db.kapcsolat.commit()

    except Exception:
        db.kapcsolat.rollback()
        raise

    finally:
        db.am_bezaras()


def am_athelyezes_excel():
    db = AMAdatbazis()

    try:
        adatok = db.kurzor.execute("""
            SELECT
                e.eszkozazonosito,
                e.megnevezes,
                e.gyartasi_szam,
                a.datum,
                rh.telephely,
                rh.helyiseg_szam,
                uh.telephely,
                uh.helyiseg_szam,
                rf.teljes_nev,
                uf.teljes_nev,
                a.megjegyzes
            FROM athelyezesek a
            JOIN eszkozok e
                ON e.id = a.eszkoz_id
            LEFT JOIN helyisegek rh
                ON rh.id = a.regi_helyiseg_id
            LEFT JOIN helyisegek uh
                ON uh.id = a.uj_helyiseg_id
            LEFT JOIN felelosok rf
                ON rf.id = a.regi_felelos_id
            LEFT JOIN felelosok uf
                ON uf.id = a.uj_felelos_id
            ORDER BY a.datum DESC, a.id DESC
        """).fetchall()

    finally:
        db.am_bezaras()

    if not adatok:
        messagebox.showinfo(
            "Excel-export",
            "Még nincs rögzített áthelyezés."
        )
        return

    utvonal = filedialog.asksaveasfilename(
        title="Áthelyezési napló mentése",
        defaultextension=".xlsx",
        filetypes=[("Excel-munkafüzet", "*.xlsx")],
        initialfile="Athelyezesi_naplo.xlsx"
    )

    if not utvonal:
        return

    munkafuzet = Workbook()
    lap = munkafuzet.active
    lap.title = "Áthelyezések"

    lap.append([
        "Eszközazonosító",
        "Megnevezés",
        "Gyártási szám",
        "Áthelyezés dátuma",
        "Régi telephely",
        "Régi helyiség",
        "Új telephely",
        "Új helyiség",
        "Régi felelős",
        "Új felelős",
        "Megjegyzés"
    ])

    for sor in adatok:
        lap.append([
            "" if ertek is None else ertek
            for ertek in sor
        ])

    for cella in lap[1]:
        cella.font = Font(
            bold=True,
            color="FFFFFF"
        )
        cella.fill = PatternFill(
            "solid",
            fgColor="243E57"
        )

    szelessegek = [
        20, 42, 20, 22, 18, 18,
        18, 18, 28, 28, 45
    ]

    for index, szelesseg in enumerate(szelessegek, 1):
        lap.column_dimensions[
            get_column_letter(index)
        ].width = szelesseg

    lap.freeze_panes = "A2"
    lap.auto_filter.ref = lap.dimensions

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
        "Az áthelyezési napló Excel-fájlja elkészült!"
    )


def am_athelyezes_ablak(szulo, eszkoz, frissites):
    db = AMAdatbazis()

    try:
        helyisegek = db.am_helyisegek_lekerese()
        felelosok = db.am_felelosok_lekerese()
    finally:
        db.am_bezaras()

    ablak = tk.Toplevel(szulo)
    ablak.title("AM – Eszköz áthelyezése")
    ablak.geometry("850x680")
    ablak.minsize(760, 600)
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

    tartalom = ttk.Frame(alap)
    tartalom.pack(
        side="top",
        fill="both",
        expand=True
    )

    ttk.Label(
        tartalom,
        text="Eszköz áthelyezése",
        font=("Segoe UI", 17, "bold")
    ).pack(
        anchor="w",
        pady=(0, 12)
    )

    adatkeret = ttk.LabelFrame(
        tartalom,
        text="Jelenlegi adatok",
        padding=10
    )
    adatkeret.pack(
        fill="x",
        pady=(0, 12)
    )

    jelenlegi = [
        ("Eszközazonosító", eszkoz[1]),
        ("Megnevezés", eszkoz[2]),
        ("Gyártási szám", eszkoz[3]),
        (
            "Helyiség",
            f"{eszkoz[6] or ''} - {eszkoz[7] or ''}"
        ),
        ("Felelős", eszkoz[9])
    ]

    for index, (felirat, ertek) in enumerate(jelenlegi):
        ttk.Label(
            adatkeret,
            text=f"{felirat}:",
            font=("Segoe UI", 9, "bold")
        ).grid(
            row=index,
            column=0,
            sticky="w",
            padx=8,
            pady=4
        )

        ttk.Label(
            adatkeret,
            text="" if ertek is None else str(ertek),
            wraplength=550
        ).grid(
            row=index,
            column=1,
            sticky="w",
            padx=8,
            pady=4
        )

    ujkeret = ttk.LabelFrame(
        tartalom,
        text="Új adatok",
        padding=10
    )
    ujkeret.pack(
        fill="x",
        pady=(0, 12)
    )

    helyiseg_nevek = {
        f"{h[1]} - {h[2]} - {h[3]}": h[0]
        for h in helyisegek
    }

    felelos_nevek = {
        f"{f[1]} ({f[2]})": f[0]
        for f in felelosok
    }

    uj_helyiseg = tk.StringVar()
    uj_felelos = tk.StringVar()

    for nev, azonosito in helyiseg_nevek.items():
        if azonosito == eszkoz[4]:
            uj_helyiseg.set(nev)
            break

    for nev, azonosito in felelos_nevek.items():
        if azonosito == eszkoz[5]:
            uj_felelos.set(nev)
            break

    datum_valtozo = tk.StringVar(
        value=date.today().isoformat()
    )
    megjegyzes = tk.StringVar()

    szerkeszto_mezok = [
        ("Új helyiség", uj_helyiseg),
        ("Új felelős", uj_felelos),
        ("Áthelyezés dátuma", datum_valtozo),
        ("Áthelyezés indoka", megjegyzes)
    ]

    for index, (felirat, valtozo) in enumerate(
        szerkeszto_mezok
    ):
        ttk.Label(
            ujkeret,
            text=f"{felirat}:"
        ).grid(
            row=index,
            column=0,
            sticky="w",
            padx=8,
            pady=5
        )

        if index == 0:
            widget = ttk.Combobox(
                ujkeret,
                textvariable=valtozo,
                values=list(helyiseg_nevek),
                state="readonly",
                width=55
            )

        elif index == 1:
            widget = ttk.Combobox(
                ujkeret,
                textvariable=valtozo,
                values=list(felelos_nevek),
                state="readonly",
                width=55
            )

        else:
            widget = ttk.Entry(
                ujkeret,
                textvariable=valtozo,
                width=58
            )

        widget.grid(
            row=index,
            column=1,
            sticky="ew",
            padx=8,
            pady=5
        )

    ujkeret.columnconfigure(1, weight=1)

    ttk.Label(
        tartalom,
        text="Korábbi áthelyezések",
        font=("Segoe UI", 11, "bold")
    ).pack(
        anchor="w",
        pady=(0, 6)
    )

    elozmenykeret = ttk.Frame(tartalom)
    elozmenykeret.pack(
        fill="both",
        expand=True
    )

    tabla = ttk.Treeview(
        elozmenykeret,
        columns=(
            "datum",
            "regi",
            "uj",
            "regi_felelos",
            "uj_felelos",
            "indok"
        ),
        show="headings",
        height=4
    )

    for kulcs, felirat, szelesseg in [
        ("datum", "Dátum", 100),
        ("regi", "Régi helyiség", 105),
        ("uj", "Új helyiség", 105),
        ("regi_felelos", "Régi felelős", 145),
        ("uj_felelos", "Új felelős", 145),
        ("indok", "Megjegyzés", 230)
    ]:
        tabla.heading(kulcs, text=felirat)
        tabla.column(kulcs, width=szelesseg)

    gorgeto = ttk.Scrollbar(
        elozmenykeret,
        command=tabla.yview
    )
    tabla.configure(yscrollcommand=gorgeto.set)

    gorgeto.pack(side="right", fill="y")
    tabla.pack(
        side="left",
        fill="both",
        expand=True
    )

    for sor in am_athelyezes_elozmenyek(eszkoz[0]):
        tabla.insert(
            "",
            "end",
            values=tuple(
                "" if ertek is None else ertek
                for ertek in sor[1:]
            )
        )

    def am_mentes():
        try:
            if uj_helyiseg.get() not in helyiseg_nevek:
                raise ValueError(
                    "Válassz új helyiséget!"
                )

            if uj_felelos.get() not in felelos_nevek:
                raise ValueError(
                    "Válassz felelőst!"
                )

            datum = date.fromisoformat(
                datum_valtozo.get().strip()
            ).isoformat()

            indok = megjegyzes.get().strip()

            if not indok:
                raise ValueError(
                    "Az áthelyezés indokát kötelező megadni!"
                )

            uj_helyiseg_id = helyiseg_nevek[
                uj_helyiseg.get()
            ]

            uj_felelos_id = felelos_nevek[
                uj_felelos.get()
            ]

            if (
                uj_helyiseg_id == eszkoz[4]
                and uj_felelos_id == eszkoz[5]
            ):
                raise ValueError(
                    "Változtasd meg a helyiséget "
                    "vagy a felelőst!"
                )

            megerosites = messagebox.askyesno(
                "Áthelyezés megerősítése",
                f"Biztosan áthelyezed az eszközt?\n\n"
                f"{eszkoz[1]} - {eszkoz[2]}\n\n"
                f"Új helyiség: {uj_helyiseg.get()}\n"
                f"Új felelős: {uj_felelos.get()}\n"
                f"Dátum: {datum}\n"
                f"Indok: {indok}",
                parent=ablak
            )

            if not megerosites:
                return

            am_athelyezes_mentes(
                eszkoz[0],
                uj_helyiseg_id,
                uj_felelos_id,
                datum,
                indok
            )

        except Exception as hiba:
            messagebox.showerror(
                "Áthelyezési hiba",
                str(hiba),
                parent=ablak
            )
            return

        ablak.destroy()
        frissites()

        messagebox.showinfo(
            "Sikeres áthelyezés",
            "Az eszköz áthelyezése sikeres!\n\n"
            "Az új helyiség és felelős rögzítésre került.\n"
            "A korábbi adatok az áthelyezési naplóban "
            "megmaradtak."
        )

    ttk.Button(
        gombkeret,
        text="Áthelyezés mentése",
        command=am_mentes
    ).pack(
        side="left",
        padx=5
    )

    ttk.Button(
        gombkeret,
        text="Mégse",
        command=ablak.destroy
    ).pack(
        side="left",
        padx=5
    )


def am_athelyezes_megjelenitese(content):
    am_athelyezes_elokeszites()

    oldal = ttk.Frame(content)
    oldal.pack(fill="both", expand=True)

    ttk.Label(
        oldal,
        text="Eszközök áthelyezése",
        font=("Segoe UI", 16, "bold")
    ).pack(
        anchor="w",
        padx=25,
        pady=(15, 7)
    )

    ttk.Label(
        oldal,
        text=(
            "Válassz ki egy eszközt, majd kattints rá "
            "duplán az áthelyezéshez."
        )
    ).pack(
        anchor="w",
        padx=25,
        pady=(0, 15)
    )

    kereses = tk.StringVar()
    helyiseg_szuro = tk.StringVar(value="Összes")
    felelos_szuro = tk.StringVar(value="Összes")

    szurokeret = ttk.Frame(oldal)
    szurokeret.pack(
        fill="x",
        padx=25,
        pady=10
    )

    ttk.Label(
        szurokeret,
        text="Keresés:"
    ).grid(
        row=0,
        column=0,
        padx=5,
        pady=6
    )

    ttk.Entry(
        szurokeret,
        textvariable=kereses,
        width=40
    ).grid(
        row=0,
        column=1,
        padx=5,
        pady=6
    )

    ttk.Label(
        szurokeret,
        text="Helyiség:"
    ).grid(
        row=1,
        column=0,
        padx=5,
        pady=6
    )

    helyiseg_combo = ttk.Combobox(
        szurokeret,
        textvariable=helyiseg_szuro,
        state="readonly",
        width=25
    )

    helyiseg_combo.grid(
        row=1,
        column=1,
        padx=5,
        pady=6
    )

    ttk.Label(
        szurokeret,
        text="Felelős:"
    ).grid(
        row=1,
        column=2,
        padx=5,
        pady=6
    )

    felelos_combo = ttk.Combobox(
        szurokeret,
        textvariable=felelos_szuro,
        state="readonly",
        width=30
    )

    felelos_combo.grid(
        row=1,
        column=3,
        padx=5,
        pady=6
    )

    tablakeret = ttk.Frame(oldal)
    tablakeret.pack(
        fill="both",
        expand=True,
        padx=25,
        pady=10
    )

    tabla = ttk.Treeview(
        tablakeret,
        columns=(
            "azonosito",
            "nev",
            "gyartasi",
            "helyiseg",
            "felelos"
        ),
        show="headings",
        selectmode="browse"
    )

    for kulcs, felirat, szelesseg in [
        ("azonosito", "Eszközazonosító", 140),
        ("nev", "Megnevezés", 285),
        ("gyartasi", "Gyártási szám", 155),
        ("helyiseg", "Helyiség", 130),
        ("felelos", "Felelős", 190)
    ]:
        tabla.heading(kulcs, text=felirat)
        tabla.column(kulcs, width=szelesseg)

    gorgeto = ttk.Scrollbar(
        tablakeret,
        command=tabla.yview
    )

    tabla.configure(yscrollcommand=gorgeto.set)
    gorgeto.pack(side="right", fill="y")

    tabla.pack(
        side="left",
        fill="both",
        expand=True
    )

    darabszam = tk.StringVar()

    ttk.Label(
        oldal,
        textvariable=darabszam
    ).pack(
        anchor="w",
        padx=25,
        pady=5
    )

    lista = []
    rekordok = {}

    def am_szures(*args):
        kijeloles = tabla.selection()
        regi_id = kijeloles[0] if kijeloles else None

        for elem in tabla.get_children():
            tabla.delete(elem)

        keresett = kereses.get().strip().casefold()
        kert_helyiseg = helyiseg_szuro.get()
        kert_felelos = felelos_szuro.get()

        darab = 0

        for sor in lista:
            szoveg = " ".join(
                str(ertek or "")
                for ertek in (
                    sor[1],
                    sor[2],
                    sor[3]
                )
            ).casefold()

            if keresett not in szoveg:
                continue

            if (
                kert_helyiseg != "Összes"
                and str(sor[7] or "") != kert_helyiseg
            ):
                continue

            if (
                kert_felelos != "Összes"
                and str(sor[9] or "") != kert_felelos
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
                    sor[7] or "",
                    sor[9] or ""
                )
            )

            darab += 1

        darabszam.set(
            f"Találatok: {darab} / {len(lista)}"
        )

        if regi_id and tabla.exists(regi_id):
            tabla.selection_set(regi_id)
            tabla.see(regi_id)

    def am_frissites():
        nonlocal lista, rekordok

        if not oldal.winfo_exists():
            return

        lista = am_athelyezheto_eszkozok()

        rekordok = {
            sor[0]: sor
            for sor in lista
        }

        helyisegek = ["Összes"] + sorted({
            str(sor[7])
            for sor in lista
            if sor[7] is not None
        })

        felelosok = ["Összes"] + sorted({
            str(sor[9])
            for sor in lista
            if sor[9] is not None
        })

        helyiseg_combo.configure(
            values=helyisegek
        )

        felelos_combo.configure(
            values=felelosok
        )

        if helyiseg_szuro.get() not in helyisegek:
            helyiseg_szuro.set("Összes")

        if felelos_szuro.get() not in felelosok:
            felelos_szuro.set("Összes")

        am_szures()

    def am_eszkoz_megnyitasa(event=None):
        kijeloles = tabla.selection()

        if not kijeloles:
            messagebox.showwarning(
                "Nincs kijelölés",
                "Válassz ki egy eszközt!"
            )
            return

        eszkoz = rekordok.get(
            int(kijeloles[0])
        )

        if eszkoz is None:
            messagebox.showerror(
                "Hiba",
                "Az eszköz nem található!"
            )
            return

        am_athelyezes_ablak(
            content,
            eszkoz,
            am_frissites
        )

    tabla.bind(
        "<Double-1>",
        am_eszkoz_megnyitasa
    )

    gombok = ttk.Frame(oldal)
    gombok.pack(
        fill="x",
        padx=25,
        pady=15
    )

    ttk.Button(
        gombok,
        text="Kijelölt eszköz áthelyezése",
        command=am_eszkoz_megnyitasa
    ).pack(
        side="left",
        padx=5
    )

    ttk.Button(
        gombok,
        text="Frissítés",
        command=am_frissites
    ).pack(
        side="left",
        padx=5
    )

    ttk.Button(
        gombok,
        text="Excel-export",
        command=am_athelyezes_excel
    ).pack(
        side="left",
        padx=5
    )

    kereses.trace_add("write", am_szures)
    helyiseg_szuro.trace_add("write", am_szures)
    felelos_szuro.trace_add("write", am_szures)

    am_frissites()
