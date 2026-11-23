
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date
from am_database import AMAdatbazis
from am_selejtezes import (
    am_selejtezes_tabla,
    am_selejtezes_ablak
)


def am_eszkoz_ablak(content, frissites, eszkoz_id=None):
    db = AMAdatbazis()

    try:
        helyisegek = db.am_helyisegek_lekerese()
        felelosok = db.am_felelosok_lekerese()
        eszkoz = (
            db.am_eszkoz_lekerese(eszkoz_id)
            if eszkoz_id is not None else None
        )
    finally:
        db.am_bezaras()

    modositas = eszkoz_id is not None

    if modositas and eszkoz is None:
        messagebox.showerror(
            "Hiba",
            "Az eszköz nem található!"
        )
        return

    ablak = tk.Toplevel(content)
    ablak.title(
        "Eszköz módosítása"
        if modositas else "Új eszköz felvétele"
    )
    ablak.geometry("650x650")
    ablak.resizable(False, False)
    ablak.transient(content.winfo_toplevel())
    ablak.grab_set()

    keret = ttk.Frame(ablak, padding=25)
    keret.pack(fill="both", expand=True)

    mezok = [
        ("azonosito", "Eszközazonosító", ""),
        ("eszkoz_szam", "Eszköz száma", "0"),
        ("megnevezes", "Megnevezés", ""),
        ("gyartasi_szam", "Gyártási szám", ""),
        ("mennyiseg", "Mennyiség", "1"),
        ("mertekegyseg", "Mértékegység", "DB"),
        (
            "datum",
            "Aktiválás dátuma",
            date.today().isoformat()
        ),
        ("brutto", "Bruttó érték (Ft)", "0"),
        ("konyv", "Könyv szerinti érték (Ft)", "0"),
        ("megjegyzes", "Megjegyzés", "")
    ]

    valtozok = {}

    for i, (kulcs, felirat, alap) in enumerate(mezok):
        ttk.Label(
            keret,
            text=felirat
        ).grid(
            row=i,
            column=0,
            sticky="w",
            pady=6
        )

        valtozok[kulcs] = tk.StringVar(value=alap)

        ttk.Entry(
            keret,
            textvariable=valtozok[kulcs],
            width=42
        ).grid(
            row=i,
            column=1,
            padx=15,
            pady=6
        )

    helyiseg_lista = {
        f"{h[1]} - {h[2]} - {h[3]}": h
        for h in helyisegek
    }

    felelos_lista = {
        f"{f[1]} ({f[2]})": f
        for f in felelosok
    }

    helyiseg = tk.StringVar()
    felelos = tk.StringVar()

    ttk.Label(
        keret,
        text="Helyiség"
    ).grid(
        row=10,
        column=0,
        sticky="w",
        pady=6
    )

    ttk.Combobox(
        keret,
        textvariable=helyiseg,
        values=list(helyiseg_lista),
        state="readonly",
        width=39
    ).grid(
        row=10,
        column=1,
        pady=6
    )

    ttk.Label(
        keret,
        text="Felelős"
    ).grid(
        row=11,
        column=0,
        sticky="w",
        pady=6
    )

    ttk.Combobox(
        keret,
        textvariable=felelos,
        values=list(felelos_lista),
        state="readonly",
        width=39
    ).grid(
        row=11,
        column=1,
        pady=6
    )

    if modositas:
        megfeleltetes = {
            "azonosito": 1,
            "eszkoz_szam": 2,
            "megnevezes": 3,
            "gyartasi_szam": 4,
            "mennyiseg": 5,
            "mertekegyseg": 6,
            "datum": 11,
            "brutto": 12,
            "konyv": 14,
            "megjegyzes": 16
        }

        for kulcs, index in megfeleltetes.items():
            ertek = eszkoz[index]
            valtozok[kulcs].set(
                "" if ertek is None else str(ertek)
            )

        for nev, h in helyiseg_lista.items():
            if h[0] == eszkoz[9]:
                helyiseg.set(nev)
                break

        for nev, f in felelos_lista.items():
            if f[0] == eszkoz[10]:
                felelos.set(nev)
                break

    def am_mentes():
        kapcsolat = None

        try:
            azonosito = (
                valtozok["azonosito"].get().strip()
            )
            megnevezes = (
                valtozok["megnevezes"].get().strip()
            )

            if not azonosito or not megnevezes:
                raise ValueError(
                    "Az azonosító és a megnevezés kötelező!"
                )

            if helyiseg.get() not in helyiseg_lista:
                raise ValueError(
                    "Válassz helyiséget!"
                )

            if felelos.get() not in felelos_lista:
                raise ValueError(
                    "Válassz felelőst!"
                )

            h = helyiseg_lista[helyiseg.get()]
            f = felelos_lista[felelos.get()]

            mennyiseg = float(
                valtozok["mennyiseg"]
                .get()
                .replace(",", ".")
            )

            brutto = float(
                valtozok["brutto"]
                .get()
                .replace(",", ".")
            )

            konyv = float(
                valtozok["konyv"]
                .get()
                .replace(",", ".")
            )

            if mennyiseg <= 0:
                raise ValueError(
                    "A mennyiségnek pozitívnak kell lennie!"
                )

            if brutto < 0 or konyv < 0:
                raise ValueError(
                    "Az értékek nem lehetnek negatívak!"
                )

            datum_szoveg = (
                valtozok["datum"].get().strip()
            )

            datum = (
                date.fromisoformat(
                    datum_szoveg
                ).isoformat()
                if datum_szoveg else None
            )

            if modositas:
                telephely = eszkoz[7]
                eszkoz_telephelye = eszkoz[8]
                brutto_penznem = eszkoz[13] or "HUF"
                konyv_penznem = eszkoz[15] or "HUF"
            else:
                telephely = h[1]
                eszkoz_telephelye = h[1]
                brutto_penznem = "HUF"
                konyv_penznem = "HUF"

            adatok = (
                azonosito,
                valtozok["eszkoz_szam"].get().strip(),
                megnevezes,
                valtozok["gyartasi_szam"].get().strip(),
                mennyiseg,
                valtozok["mertekegyseg"].get().strip(),
                telephely,
                eszkoz_telephelye,
                h[0],
                f[0],
                datum,
                brutto,
                brutto_penznem,
                konyv,
                konyv_penznem,
                valtozok["megjegyzes"].get().strip()
            )

            kapcsolat = AMAdatbazis()

            if modositas:
                kapcsolat.am_eszkoz_modositasa(
                    eszkoz_id,
                    adatok
                )
            else:
                kapcsolat.am_eszkoz_felvetele(
                    adatok
                )

        except Exception as hiba:
            messagebox.showerror(
                "Hiba",
                str(hiba),
                parent=ablak
            )
            return

        finally:
            if kapcsolat is not None:
                kapcsolat.am_bezaras()

        messagebox.showinfo(
            "Sikeres mentés",
            (
                "Az eszköz módosítása sikeres!"
                if modositas
                else "Az új eszköz rögzítése sikeres!"
            ),
            parent=ablak
        )

        ablak.destroy()
        frissites()

    gombok = ttk.Frame(keret)
    gombok.grid(
        row=12,
        column=0,
        columnspan=2,
        pady=25
    )

    ttk.Button(
        gombok,
        text=(
            "Módosítás mentése"
            if modositas else "Mentés"
        ),
        command=am_mentes
    ).pack(
        side="left",
        padx=10
    )

    ttk.Button(
        gombok,
        text="Mégse",
        command=ablak.destroy
    ).pack(
        side="left",
        padx=10
    )


def am_uj_eszkoz(content, frissites):
    am_eszkoz_ablak(
        content,
        frissites
    )


def am_archivum_ablak(content, frissites):
    ablak = tk.Toplevel(content)
    ablak.title("Archivált eszközök")
    ablak.geometry("750x450")
    ablak.transient(content.winfo_toplevel())
    ablak.grab_set()

    keret = ttk.Frame(
        ablak,
        padding=15
    )
    keret.pack(
        fill="both",
        expand=True
    )

    ttk.Label(
        keret,
        text="Archivált eszközök",
        font=("Segoe UI", 16, "bold")
    ).pack(
        anchor="w",
        pady=(0, 15)
    )

    tablakeret = ttk.Frame(keret)
    tablakeret.pack(
        fill="both",
        expand=True
    )

    tabla = ttk.Treeview(
        tablakeret,
        columns=("azonosito", "megnevezes"),
        show="headings",
        selectmode="browse"
    )

    tabla.heading(
        "azonosito",
        text="Eszközazonosító"
    )
    tabla.heading(
        "megnevezes",
        text="Megnevezés"
    )

    tabla.column(
        "azonosito",
        width=160
    )
    tabla.column(
        "megnevezes",
        width=450
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

    def am_archivum_frissites():
        for elem in tabla.get_children():
            tabla.delete(elem)

        db = AMAdatbazis()

        try:
            eszkozok = db.kurzor.execute("""
                SELECT
                    e.id,
                    e.eszkozazonosito,
                    e.megnevezes
                FROM eszkozok e
                WHERE e.archivalt = 1
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

        for eszkoz in eszkozok:
            tabla.insert(
                "",
                tk.END,
                iid=str(eszkoz[0]),
                values=(
                    eszkoz[1] or "",
                    eszkoz[2] or ""
                )
            )

    def am_visszaallitas():
        kijeloles = tabla.selection()

        if not kijeloles:
            messagebox.showwarning(
                "Figyelmeztetés",
                "Válassz ki egy archivált eszközt!",
                parent=ablak
            )
            return

        adatok = tabla.item(
            kijeloles[0],
            "values"
        )

        kerdes = (
            "Visszaállítod az alábbi eszközt?\n\n"
            f"{adatok[0]} - {adatok[1]}"
        )

        if not messagebox.askyesno(
            "Visszaállítás megerősítése",
            kerdes,
            parent=ablak
        ):
            return

        db = AMAdatbazis()

        try:
            db.am_eszkoz_visszaallitasa(
                int(kijeloles[0])
            )
        except Exception as hiba:
            messagebox.showerror(
                "Hiba",
                str(hiba),
                parent=ablak
            )
            return
        finally:
            db.am_bezaras()

        am_archivum_frissites()
        frissites()

        messagebox.showinfo(
            "Siker",
            "Az eszköz visszaállítása sikeres!",
            parent=ablak
        )

    gombok = ttk.Frame(keret)
    gombok.pack(
        fill="x",
        pady=15
    )

    ttk.Button(
        gombok,
        text="Visszaállítás",
        command=am_visszaallitas
    ).pack(
        side="left",
        padx=5
    )

    ttk.Button(
        gombok,
        text="Bezárás",
        command=ablak.destroy
    ).pack(
        side="right",
        padx=5
    )

    am_archivum_frissites()


def am_eszkozok_megjelenitese(content):
    if not content.winfo_exists():
        return

    regi_oldal = getattr(
        content,
        "_am_eszkoz_aktiv_oldal",
        None
    )

    if (
        regi_oldal is not None
        and regi_oldal.winfo_exists()
    ):
        regi_oldal.destroy()

    am_selejtezes_tabla()

    db = AMAdatbazis()

    try:
        adatok = db.kurzor.execute("""
            SELECT
                e.id,
                e.eszkozazonosito,
                e.megnevezes,
                e.gyartasi_szam,
                h.helyiseg_szam,
                f.teljes_nev,
                e.brutto_ertek
            FROM eszkozok e
            LEFT JOIN helyisegek h
                ON e.helyiseg_id = h.id
            LEFT JOIN felelosok f
                ON e.felelos_id = f.id
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

    oldal = ttk.Frame(content)

    content._am_eszkoz_aktiv_oldal = oldal

    oldal.pack(
        fill="both",
        expand=True
    )

    kereses = tk.StringVar()
    helyiseg = tk.StringVar(value="Összes")
    felelos = tk.StringVar(value="Összes")
    talalatok = tk.StringVar()

    def am_frissites():
        if not content.winfo_exists():
            return

        aktiv_oldal = getattr(
            content,
            "_am_eszkoz_aktiv_oldal",
            None
        )

        if aktiv_oldal is not oldal:
            return

        am_eszkozok_megjelenitese(content)

    gombkeret = ttk.Frame(oldal)
    gombkeret.pack(
        fill="x",
        padx=30,
        pady=5
    )

    ttk.Button(
        gombkeret,
        text="Új eszköz",
        command=lambda: am_uj_eszkoz(
            content,
            am_frissites
        )
    ).pack(
        side="left",
        padx=4
    )

    def am_kijelolt_modositasa():
        kijeloles = tabla.selection()

        if not kijeloles:
            messagebox.showwarning(
                "Figyelmeztetés",
                "Először válassz ki egy eszközt!"
            )
            return

        am_eszkoz_ablak(
            content,
            am_frissites,
            int(kijeloles[0])
        )

    ttk.Button(
        gombkeret,
        text="Módosítás",
        command=am_kijelolt_modositasa
    ).pack(
        side="left",
        padx=4
    )

    def am_archivalas():
        kijeloles = tabla.selection()

        if not kijeloles:
            messagebox.showwarning(
                "Figyelmeztetés",
                "Válassz ki egy archiválandó eszközt!"
            )
            return

        adatok = tabla.item(
            kijeloles[0],
            "values"
        )

        kerdes = (
            "Biztosan archiválod ezt az eszközt?\n\n"
            f"{adatok[0]} - {adatok[1]}\n\n"
            "Az eszköz később visszaállítható."
        )

        if not messagebox.askyesno(
            "Archiválás megerősítése",
            kerdes
        ):
            return

        db = AMAdatbazis()

        try:
            db.am_eszkoz_archivalasa(
                int(kijeloles[0])
            )
        except Exception as hiba:
            messagebox.showerror(
                "Hiba",
                str(hiba)
            )
            return
        finally:
            db.am_bezaras()

        am_frissites()

        messagebox.showinfo(
            "Sikeres archiválás",
            "Az eszköz archiválása sikeres!"
        )

    ttk.Button(
        gombkeret,
        text="Törlés (archiválás)",
        command=am_archivalas
    ).pack(
        side="left",
        padx=4
    )

    ttk.Button(
        gombkeret,
        text="Archivált eszközök",
        command=lambda: am_archivum_ablak(
            content,
            am_frissites
        )
    ).pack(
        side="left",
        padx=4
    )

    ttk.Button(
        gombkeret,
        text="Selejtezés",
        command=lambda: am_selejtezes_ablak(
            content,
            am_frissites
        )
    ).pack(
        side="left",
        padx=4
    )

    ttk.Button(
        gombkeret,
        text="Frissítés",
        command=am_frissites
    ).pack(
        side="left",
        padx=4
    )

    szurokeret = ttk.Frame(oldal)
    szurokeret.pack(
        fill="x",
        padx=30,
        pady=10
    )

    ttk.Label(
        szurokeret,
        text="Keresés:"
    ).grid(
        row=0,
        column=0,
        padx=5,
        pady=5
    )

    ttk.Entry(
        szurokeret,
        textvariable=kereses,
        width=40
    ).grid(
        row=0,
        column=1,
        padx=5,
        pady=5
    )

    ttk.Label(
        szurokeret,
        text="Helyiség:"
    ).grid(
        row=1,
        column=0,
        padx=5,
        pady=5
    )

    helyisegek = ["Összes"] + sorted({
        str(sor[4])
        for sor in adatok
        if sor[4] is not None
    })

    ttk.Combobox(
        szurokeret,
        textvariable=helyiseg,
        values=helyisegek,
        state="readonly",
        width=20
    ).grid(
        row=1,
        column=1,
        padx=5,
        pady=5
    )

    ttk.Label(
        szurokeret,
        text="Felelős:"
    ).grid(
        row=1,
        column=2,
        padx=5,
        pady=5
    )

    felelosok = ["Összes"] + sorted({
        str(sor[5])
        for sor in adatok
        if sor[5] is not None
    })

    ttk.Combobox(
        szurokeret,
        textvariable=felelos,
        values=felelosok,
        state="readonly",
        width=25
    ).grid(
        row=1,
        column=3,
        padx=5,
        pady=5
    )

    def am_alaphelyzet():
        kereses.set("")
        helyiseg.set("Összes")
        felelos.set("Összes")

    ttk.Button(
        szurokeret,
        text="Alaphelyzet",
        command=am_alaphelyzet
    ).grid(
        row=0,
        column=3,
        padx=5,
        pady=5
    )

    tablakeret = ttk.Frame(oldal)
    tablakeret.pack(
        fill="both",
        expand=True,
        padx=30,
        pady=10
    )

    oszlopok = (
        "azonosito",
        "megnevezes",
        "gyartasi_szam",
        "helyiseg",
        "felelos",
        "ertek"
    )

    tabla = ttk.Treeview(
        tablakeret,
        columns=oszlopok,
        show="headings",
        selectmode="browse"
    )

    fejlecek = [
        ("azonosito", "Eszközazonosító", 130),
        ("megnevezes", "Megnevezés", 260),
        ("gyartasi_szam", "Gyártási szám", 150),
        ("helyiseg", "Helyiség", 100),
        ("felelos", "Felelős", 160),
        ("ertek", "Bruttó érték", 120)
    ]

    for oszlop, nev, szelesseg in fejlecek:
        tabla.heading(
            oszlop,
            text=nev
        )

        tabla.column(
            oszlop,
            width=szelesseg
        )

    tabla.column(
        "ertek",
        anchor="e"
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

    tabla.bind(
        "<Double-1>",
        lambda event: am_kijelolt_modositasa()
    )

    def am_szures(*args):
        for elem in tabla.get_children():
            tabla.delete(elem)

        keresett = (
            kereses.get().casefold().strip()
        )

        darab = 0

        for sor in adatok:
            szoveg = " ".join(
                str(ertek)
                for ertek in sor[1:4]
                if ertek is not None
            ).casefold()

            if keresett not in szoveg:
                continue

            if helyiseg.get() != "Összes":
                if str(sor[4]) != helyiseg.get():
                    continue

            if felelos.get() != "Összes":
                if str(sor[5]) != felelos.get():
                    continue

            megjelenitett = [
                "" if ertek is None else ertek
                for ertek in sor[1:]
            ]

            if isinstance(
                sor[6],
                (int, float)
            ):
                megjelenitett[5] = (
                    f"{sor[6]:,.0f} Ft"
                    .replace(",", " ")
                )

            tabla.insert(
                "",
                tk.END,
                iid=str(sor[0]),
                values=megjelenitett
            )

            darab += 1

        talalatok.set(
            f"Találatok száma: {darab} / {len(adatok)}"
        )

    ttk.Label(
        oldal,
        textvariable=talalatok
    ).pack(
        anchor="w",
        padx=30,
        pady=10
    )

    kereses.trace_add(
        "write",
        am_szures
    )

    helyiseg.trace_add(
        "write",
        am_szures
    )

    felelos.trace_add(
        "write",
        am_szures
    )

    am_szures()
