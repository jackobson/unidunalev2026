
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from am_database import AMAdatbazis


AM_ALLAPOTOK = {
    "nincs": "Aktív / nem jelölt",
    "jelolt": "Selejtezésre jelölt",
    "jovahagyott": "Jóváhagyott",
    "elszallitva": "Elszállítva",
    "visszavont": "Aktív / nem jelölt"
}


def am_selejtezes_tabla():
    db = AMAdatbazis()
    try:
        db.kurzor.execute("""
            CREATE TABLE IF NOT EXISTS selejtezesek (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                eszkoz_id INTEGER NOT NULL UNIQUE,
                allapot TEXT NOT NULL DEFAULT 'jelolt',
                jeloles_datuma TEXT NOT NULL,
                selejtezes_datuma TEXT,
                indok TEXT NOT NULL,
                megjegyzes TEXT,
                jegyzokonyv_szam TEXT,
                elszallitas_datuma TEXT,
                elszallito TEXT,
                atvevo TEXT,
                szallitolevel_szam TEXT,
                FOREIGN KEY (eszkoz_id) REFERENCES eszkozok(id)
            )
        """)

        db.kurzor.execute("""
            CREATE TABLE IF NOT EXISTS am_selejtezes_naplo (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                eszkoz_id INTEGER NOT NULL,
                idopont TEXT NOT NULL,
                muvelet TEXT NOT NULL,
                elozo_allapot TEXT NOT NULL,
                uj_allapot TEXT NOT NULL,
                indok TEXT,
                korabbi_adatok TEXT,
                FOREIGN KEY (eszkoz_id) REFERENCES eszkozok(id)
            )
        """)

        db.kapcsolat.commit()
    finally:
        db.am_bezaras()


def am_selejtezes_lista():
    db = AMAdatbazis()
    try:
        return db.kurzor.execute("""
            SELECT
                e.id,
                e.eszkozazonosito,
                e.megnevezes,
                h.helyiseg_szam,
                COALESCE(s.allapot, 'nincs'),
                s.jeloles_datuma,
                s.selejtezes_datuma,
                s.indok,
                s.megjegyzes,
                s.jegyzokonyv_szam,
                s.elszallitas_datuma,
                s.elszallito,
                s.atvevo,
                s.szallitolevel_szam
            FROM eszkozok e
            LEFT JOIN helyisegek h ON h.id = e.helyiseg_id
            LEFT JOIN selejtezesek s ON s.eszkoz_id = e.id
            WHERE e.archivalt = 0
                OR s.allapot = 'elszallitva'
            ORDER BY e.eszkozazonosito, e.id
        """).fetchall()
    finally:
        db.am_bezaras()


def am_naplozas(db, eszkoz_id, muvelet, elozo, uj,
                indok="", korabbi=None):
    db.kurzor.execute("""
        INSERT INTO am_selejtezes_naplo (
            eszkoz_id, idopont, muvelet,
            elozo_allapot, uj_allapot, indok, korabbi_adatok
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        eszkoz_id,
        datetime.now().isoformat(timespec="seconds"),
        muvelet,
        elozo,
        uj,
        indok,
        json.dumps(korabbi, ensure_ascii=False)
        if korabbi is not None else None
    ))


def am_muvelet_mentese(eszkoz_id, elvart, uj, muvelet,
                       indok="", mezok=None, visszavonas=False):
    db = AMAdatbazis()
    try:
        db.kurzor.execute("BEGIN IMMEDIATE")

        rekord = db.kurzor.execute("""
            SELECT
                id, eszkoz_id, allapot, jeloles_datuma,
                selejtezes_datuma, indok, megjegyzes,
                jegyzokonyv_szam, elszallitas_datuma,
                elszallito, atvevo, szallitolevel_szam
            FROM selejtezesek
            WHERE eszkoz_id = ?
        """, (eszkoz_id,)).fetchone()

        elozo = rekord[2] if rekord else "nincs"

        if elozo != elvart:
            raise ValueError(
                "Az eszköz állapota időközben megváltozott. "
                "Frissítsd a listát!"
            )

        if rekord is None:
            if uj != "jelolt":
                raise ValueError("Először selejtezésre kell jelölni!")

            db.kurzor.execute("""
                INSERT INTO selejtezesek (
                    eszkoz_id, allapot, jeloles_datuma,
                    indok, megjegyzes
                )
                VALUES (?, 'jelolt', ?, ?, ?)
            """, (
                eszkoz_id,
                mezok["jeloles"],
                mezok["indok"],
                mezok["megjegyzes"]
            ))
        else:
            frissitendo = {"allapot": uj}

            if visszavonas:
                if elozo == "elszallitva":
                    frissitendo.update({
                        "elszallitas_datuma": None,
                        "elszallito": None,
                        "atvevo": None,
                        "szallitolevel_szam": None
                    })
                elif elozo == "jovahagyott":
                    frissitendo.update({
                        "selejtezes_datuma": None,
                        "jegyzokonyv_szam": None
                    })

            elif uj == "jelolt":
                frissitendo.update({
                    "jeloles_datuma": mezok["jeloles"],
                    "selejtezes_datuma": None,
                    "indok": mezok["indok"],
                    "megjegyzes": mezok["megjegyzes"],
                    "jegyzokonyv_szam": None,
                    "elszallitas_datuma": None,
                    "elszallito": None,
                    "atvevo": None,
                    "szallitolevel_szam": None
                })

            elif uj == "jovahagyott":
                frissitendo.update({
                    "selejtezes_datuma": mezok["selejtezes"],
                    "jegyzokonyv_szam": mezok["jegyzokonyv"],
                    "megjegyzes": mezok["megjegyzes"]
                })

            elif uj == "elszallitva":
                frissitendo.update({
                    "elszallitas_datuma": mezok["elszallitas"],
                    "elszallito": mezok["elszallito"],
                    "atvevo": mezok["atvevo"],
                    "szallitolevel_szam": mezok["szallitolevel"]
                })

            oszlopok = ", ".join(
                f"{nev} = ?" for nev in frissitendo
            )

            db.kurzor.execute(
                f"UPDATE selejtezesek SET {oszlopok} "
                "WHERE eszkoz_id = ?",
                (*frissitendo.values(), eszkoz_id)
            )

        am_naplozas(
            db, eszkoz_id, muvelet, elozo, uj,
            indok, rekord
        )

        db.kapcsolat.commit()

    except Exception:
        db.kapcsolat.rollback()
        raise
    finally:
        db.am_bezaras()


def am_selejtezes_export():
    db = AMAdatbazis()

    try:
        adatok = db.kurzor.execute("""
            SELECT
                e.eszkozazonosito,
                e.megnevezes,
                e.gyartasi_szam,
                h.helyiseg_szam,
                f.teljes_nev,
                s.allapot,
                s.jeloles_datuma,
                s.selejtezes_datuma,
                s.indok,
                s.megjegyzes,
                s.jegyzokonyv_szam,
                s.elszallitas_datuma,
                s.elszallito,
                s.atvevo,
                s.szallitolevel_szam
            FROM selejtezesek s
            JOIN eszkozok e ON e.id = s.eszkoz_id
            LEFT JOIN helyisegek h ON h.id = e.helyiseg_id
            LEFT JOIN felelosok f ON f.id = e.felelos_id
            ORDER BY s.jeloles_datuma, e.eszkozazonosito
        """).fetchall()
    finally:
        db.am_bezaras()

    if not adatok:
        messagebox.showinfo(
            "Exportálás", "Nincs exportálható selejtezési adat."
        )
        return

    utvonal = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel fájl", "*.xlsx")],
        initialfile="Selejtezesi_nyilvantartas.xlsx"
    )

    if not utvonal:
        return

    munkafuzet = Workbook()
    lap = munkafuzet.active
    lap.title = "Selejtezések"

    lap.append([
        "Eszközazonosító", "Megnevezés", "Gyártási szám",
        "Helyiség", "Felelős", "Állapot",
        "Jelölés dátuma", "Selejtezés dátuma",
        "Indok", "Megjegyzés", "Jegyzőkönyv száma",
        "Elszállítás dátuma", "Elszállító",
        "Átvevő", "Szállítólevél száma"
    ])

    for sor in adatok:
        ertekek = list(sor)
        ertekek[5] = AM_ALLAPOTOK.get(
            ertekek[5], ertekek[5]
        )
        lap.append([
            "" if ertek is None else ertek
            for ertek in ertekek
        ])

    for cella in lap[1]:
        cella.font = Font(bold=True, color="FFFFFF")
        cella.fill = PatternFill(
            "solid", fgColor="17365D"
        )

    for index in range(1, 16):
        lap.column_dimensions[
            get_column_letter(index)
        ].width = 24 if index != 2 else 42

    lap.freeze_panes = "A2"
    lap.auto_filter.ref = lap.dimensions

    try:
        munkafuzet.save(utvonal)
        messagebox.showinfo(
            "Siker", "Az Excel-fájl elkészült!"
        )
    except Exception as hiba:
        messagebox.showerror("Hiba", str(hiba))


def am_selejtezes_ablak(content, frissites):
    am_selejtezes_tabla()

    ablak = tk.Toplevel(content)
    ablak.title("AM – Selejtezési nyilvántartás")
    ablak.geometry("1150x790")
    ablak.minsize(960, 680)
    ablak.transient(content.winfo_toplevel())
    ablak.grab_set()

    fo = ttk.Frame(ablak, padding=15)
    fo.pack(fill="both", expand=True)

    ttk.Label(
        fo, text="Selejtezési nyilvántartás",
        font=("Segoe UI", 17, "bold")
    ).pack(anchor="w", pady=(0, 12))

    szurokeret = ttk.Frame(fo)
    szurokeret.pack(fill="x", pady=(0, 10))

    kereses = tk.StringVar()
    szures = tk.StringVar(value="Összes")

    ttk.Label(
        szurokeret, text="Keresés:"
    ).pack(side="left", padx=5)

    ttk.Entry(
        szurokeret, textvariable=kereses, width=40
    ).pack(side="left", padx=5)

    ttk.Label(
        szurokeret, text="Állapot:"
    ).pack(side="left", padx=5)

    szuroertekek = {"Összes": None}
    szuroertekek.update({
        nev: kod for kod, nev in AM_ALLAPOTOK.items()
    })

    ttk.Combobox(
        szurokeret,
        textvariable=szures,
        values=list(szuroertekek),
        state="readonly",
        width=27
    ).pack(side="left", padx=5)

    tablakeret = ttk.Frame(fo)
    tablakeret.pack(fill="both", expand=True)

    tabla = ttk.Treeview(
        tablakeret,
        columns=("azonosito", "nev", "helyiseg", "allapot"),
        show="headings",
        selectmode="browse"
    )

    for oszlop, felirat, szelesseg in [
        ("azonosito", "Eszközazonosító", 150),
        ("nev", "Megnevezés", 440),
        ("helyiseg", "Helyiség", 140),
        ("allapot", "Állapot", 230)
    ]:
        tabla.heading(oszlop, text=felirat)
        tabla.column(oszlop, width=szelesseg)

    gorgeto = ttk.Scrollbar(
        tablakeret, command=tabla.yview
    )
    tabla.configure(yscrollcommand=gorgeto.set)
    gorgeto.pack(side="right", fill="y")
    tabla.pack(side="left", fill="both", expand=True)

    szamlalo = tk.StringVar()
    ttk.Label(
        fo, textvariable=szamlalo
    ).pack(anchor="w", pady=7)

    adatkeret = ttk.LabelFrame(
        fo, text="Selejtezési és elszállítási adatok",
        padding=12
    )
    adatkeret.pack(fill="x", pady=10)

    mezok = [
        ("jeloles", "Jelölés dátuma"),
        ("selejtezes", "Selejtezés dátuma"),
        ("indok", "Selejtezés indoka"),
        ("jegyzokonyv", "Jegyzőkönyv száma"),
        ("elszallitas", "Elszállítás dátuma"),
        ("elszallito", "Elszállító neve / cég"),
        ("atvevo", "Átvevő neve"),
        ("szallitolevel", "Szállítólevél száma"),
        ("megjegyzes", "Megjegyzés")
    ]

    valtozok = {
        kulcs: tk.StringVar()
        for kulcs, _ in mezok
    }

    for index, (kulcs, felirat) in enumerate(mezok):
        csoport = index // 5
        sor = index % 5

        ttk.Label(
            adatkeret, text=felirat
        ).grid(
            row=sor, column=csoport * 2,
            sticky="w", padx=5, pady=5
        )

        ttk.Entry(
            adatkeret,
            textvariable=valtozok[kulcs],
            width=36
        ).grid(
            row=sor, column=csoport * 2 + 1,
            sticky="ew", padx=8, pady=5
        )

    adatkeret.columnconfigure(1, weight=1)
    adatkeret.columnconfigure(3, weight=1)

    lista = []
    rekordok = {}

    def am_kijelolt():
        kijeloles = tabla.selection()

        if not kijeloles:
            messagebox.showwarning(
                "Figyelmeztetés",
                "Válassz ki egy eszközt!",
                parent=ablak
            )
            return None

        return rekordok.get(int(kijeloles[0]))

    def am_datum(kulcs):
        ertek = valtozok[kulcs].get().strip()
        if not ertek:
            raise ValueError("A dátum megadása kötelező!")
        return date.fromisoformat(ertek).isoformat()

    def am_mezok_betoltese(event=None):
        for valtozo in valtozok.values():
            valtozo.set("")

        sor = am_kijelolt_csendben()
        if sor is None:
            return

        megfeleltetes = {
            "jeloles": 5,
            "selejtezes": 6,
            "indok": 7,
            "megjegyzes": 8,
            "jegyzokonyv": 9,
            "elszallitas": 10,
            "elszallito": 11,
            "atvevo": 12,
            "szallitolevel": 13
        }

        for kulcs, index in megfeleltetes.items():
            ertek = sor[index]
            valtozok[kulcs].set(
                "" if ertek is None else str(ertek)
            )

        for kulcs in ("jeloles", "selejtezes", "elszallitas"):
            if not valtozok[kulcs].get():
                valtozok[kulcs].set(date.today().isoformat())

    def am_kijelolt_csendben():
        kijeloles = tabla.selection()
        if not kijeloles:
            return None
        return rekordok.get(int(kijeloles[0]))

    def am_lista_megjelenites(*args):
        kijeloles = tabla.selection()
        korabbi_id = kijeloles[0] if kijeloles else None

        for elem in tabla.get_children():
            tabla.delete(elem)

        keresett = kereses.get().casefold().strip()
        kivant = szuroertekek.get(szures.get())

        darab = 0

        for sor in lista:
            szoveg = " ".join(
                str(ertek or "")
                for ertek in sor[1:4]
            ).casefold()

            if keresett not in szoveg:
                continue

            if kivant is not None:
                allapot = (
                    "nincs" if sor[4] == "visszavont"
                    else sor[4]
                )
                if kivant != allapot:
                    continue

            tabla.insert(
                "", "end", iid=str(sor[0]),
                values=(
                    sor[1] or "",
                    sor[2] or "",
                    sor[3] or "",
                    AM_ALLAPOTOK.get(sor[4], sor[4])
                )
            )
            darab += 1

        szamlalo.set(
            f"Találatok: {darab} / {len(lista)}"
        )

        if korabbi_id and tabla.exists(korabbi_id):
            tabla.selection_set(korabbi_id)
            tabla.see(korabbi_id)

    def am_friss_lista():
        nonlocal lista, rekordok
        lista = am_selejtezes_lista()
        rekordok = {
            sor[0]: sor for sor in lista
        }
        am_lista_megjelenites()

    def am_muvelet(uj):
        sor = am_kijelolt()
        if sor is None:
            return

        jelenlegi = sor[4]

        kovetkezo = {
            "jelolt": ("nincs", "visszavont"),
            "jovahagyott": ("jelolt",),
            "elszallitva": ("jovahagyott",)
        }

        if uj == "jelolt":
            megengedett = jelenlegi in ("nincs", "visszavont")
        else:
            megengedett = jelenlegi in kovetkezo.get(uj, ())

        if not megengedett:
            messagebox.showwarning(
                "Nem lehetséges",
                "A művelet nem végezhető el "
                "a jelenlegi állapotból.",
                parent=ablak
            )
            return

        try:
            adatok = {
                kulcs: valtozo.get().strip()
                for kulcs, valtozo in valtozok.items()
            }

            if uj == "jelolt":
                adatok["jeloles"] = am_datum("jeloles")
                if not adatok["indok"]:
                    raise ValueError(
                        "Add meg a selejtezés indokát!"
                    )

            elif uj == "jovahagyott":
                adatok["selejtezes"] = am_datum(
                    "selejtezes"
                )
                if adatok["selejtezes"] < sor[5]:
                    raise ValueError(
                        "A jóváhagyás nem előzheti meg "
                        "a jelölést!"
                    )

            elif uj == "elszallitva":
                adatok["elszallitas"] = am_datum(
                    "elszallitas"
                )
                if not adatok["elszallito"]:
                    raise ValueError(
                        "Add meg az elszállító nevét!"
                    )
                if adatok["elszallitas"] < sor[6]:
                    raise ValueError(
                        "Az elszállítás nem előzheti meg "
                        "a selejtezést!"
                    )

            if not messagebox.askyesno(
                "Megerősítés",
                f"Biztosan végrehajtod ezt a műveletet?\n\n"
                f"{sor[1]} - {sor[2]}\n"
                f"Új állapot: {AM_ALLAPOTOK[uj]}",
                parent=ablak
            ):
                return

            am_muvelet_mentese(
                sor[0], jelenlegi, uj,
                AM_ALLAPOTOK[uj],
                adatok.get("indok", ""),
                adatok
            )

        except Exception as hiba:
            messagebox.showerror(
                "Hiba", str(hiba), parent=ablak
            )
            return

        am_friss_lista()
        am_mezok_betoltese()
        frissites()

        messagebox.showinfo(
            "Siker",
            "A művelet sikeresen megtörtént!",
            parent=ablak
        )

    def am_visszavonas():
        sor = am_kijelolt()
        if sor is None:
            return

        elozo = sor[4]

        visszalepes = {
            "jelolt": "visszavont",
            "jovahagyott": "jelolt",
            "elszallitva": "jovahagyott"
        }

        if elozo not in visszalepes:
            messagebox.showwarning(
                "Nem lehetséges",
                "Ennél az eszköznél nincs "
                "visszavonható selejtezési lépés.",
                parent=ablak
            )
            return

        uj = visszalepes[elozo]

        indok = simpledialog.askstring(
            "Visszavonás indoka",
            "Miért szükséges visszavonni a műveletet?\n"
            "Az indok bekerül a műveleti naplóba.",
            parent=ablak
        )

        if indok is None:
            return

        indok = indok.strip()

        if not indok:
            messagebox.showwarning(
                "Hiányzó indok",
                "A visszavonás indoka kötelező!",
                parent=ablak
            )
            return

        if not messagebox.askyesno(
            "Visszavonás megerősítése",
            f"Eszköz: {sor[1]} - {sor[2]}\n\n"
            f"Jelenlegi állapot: {AM_ALLAPOTOK[elozo]}\n"
            f"Visszaállítás: {AM_ALLAPOTOK[uj]}\n\n"
            "Biztosan visszavonod?",
            parent=ablak
        ):
            return

        try:
            am_muvelet_mentese(
                sor[0], elozo, uj,
                "Selejtezési művelet visszavonása",
                indok,
                visszavonas=True
            )
        except Exception as hiba:
            messagebox.showerror(
                "Hiba", str(hiba), parent=ablak
            )
            return

        am_friss_lista()
        am_mezok_betoltese()
        frissites()

        messagebox.showinfo(
            "Sikeres visszavonás",
            "A műveletet visszavontuk.\n"
            "A korábbi állapot és az indok "
            "megmaradt a naplóban.",
            parent=ablak
        )

    def am_naplo_megnyitas():
        sor = am_kijelolt()
        if sor is None:
            return

        db = AMAdatbazis()
        try:
            naplo = db.kurzor.execute("""
                SELECT
                    idopont, muvelet, elozo_allapot,
                    uj_allapot, indok
                FROM am_selejtezes_naplo
                WHERE eszkoz_id = ?
                ORDER BY id DESC
            """, (sor[0],)).fetchall()
        finally:
            db.am_bezaras()

        naploablak = tk.Toplevel(ablak)
        naploablak.title("Selejtezési műveleti napló")
        naploablak.geometry("930x400")
        naploablak.transient(ablak)

        ttk.Label(
            naploablak,
            text=f"Műveleti napló: {sor[1]} - {sor[2]}",
            font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", padx=12, pady=12)

        keret = ttk.Frame(naploablak, padding=12)
        keret.pack(fill="both", expand=True)

        oszlopok = (
            "idopont", "muvelet", "elozo", "uj", "indok"
        )

        naplotabla = ttk.Treeview(
            keret, columns=oszlopok, show="headings"
        )

        for oszlop, nev, szelesseg in [
            ("idopont", "Időpont", 165),
            ("muvelet", "Művelet", 190),
            ("elozo", "Előző állapot", 145),
            ("uj", "Új állapot", 145),
            ("indok", "Indok", 300)
        ]:
            naplotabla.heading(oszlop, text=nev)
            naplotabla.column(oszlop, width=szelesseg)

        for bejegyzes in naplo:
            naplotabla.insert(
                "", "end", values=(
                    bejegyzes[0],
                    bejegyzes[1],
                    AM_ALLAPOTOK.get(
                        bejegyzes[2], bejegyzes[2]
                    ),
                    AM_ALLAPOTOK.get(
                        bejegyzes[3], bejegyzes[3]
                    ),
                    bejegyzes[4] or ""
                )
            )

        naplotabla.pack(fill="both", expand=True)

        if not naplo:
            ttk.Label(
                naploablak,
                text="Ehhez az eszközhöz még nincs naplóbejegyzés."
            ).pack(pady=8)

    tabla.bind(
        "<<TreeviewSelect>>", am_mezok_betoltese
    )
    kereses.trace_add(
        "write", am_lista_megjelenites
    )
    szures.trace_add(
        "write", am_lista_megjelenites
    )

    gombkeret = ttk.Frame(fo)
    gombkeret.pack(fill="x", pady=8)

    for felirat, muvelet in [
        ("1. Selejtezésre jelölés",
         lambda: am_muvelet("jelolt")),
        ("2. Jóváhagyás",
         lambda: am_muvelet("jovahagyott")),
        ("3. Elszállítás",
         lambda: am_muvelet("elszallitva")),
        ("Visszavonás", am_visszavonas),
        ("Műveleti napló", am_naplo_megnyitas),
        ("Excel-export", am_selejtezes_export)
    ]:
        ttk.Button(
            gombkeret, text=felirat,
            command=muvelet
        ).pack(side="left", padx=3)

    ttk.Button(
        gombkeret, text="Bezárás",
        command=ablak.destroy
    ).pack(side="right", padx=3)

    am_friss_lista()
