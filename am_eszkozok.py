import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date
from am_database import AMAdatbazis


def am_uj_eszkoz(content, frissites):
    adatbazis = AMAdatbazis()
    helyisegek = adatbazis.am_helyisegek_lekerese()
    felelosok = adatbazis.am_felelosok_lekerese()
    adatbazis.am_bezaras()

    ablak = tk.Toplevel(content)
    ablak.title("Új eszköz felvétele")
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
        ("datum", "Aktiválás dátuma", date.today().isoformat()),
        ("brutto", "Bruttó érték (Ft)", "0"),
        ("konyv", "Könyv szerinti érték (Ft)", "0"),
        ("megjegyzes", "Megjegyzés", "")
    ]

    valtozok = {}

    for i, (kulcs, felirat, alapertelmezett) in enumerate(mezok):
        ttk.Label(
            keret, text=felirat
        ).grid(row=i, column=0, sticky="w", pady=6)

        valtozok[kulcs] = tk.StringVar(
            value=alapertelmezett
        )

        ttk.Entry(
            keret,
            textvariable=valtozok[kulcs],
            width=42
        ).grid(row=i, column=1, padx=15, pady=6)

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
        keret, text="Helyiség"
    ).grid(row=10, column=0, sticky="w", pady=6)

    ttk.Combobox(
        keret,
        textvariable=helyiseg,
        values=list(helyiseg_lista),
        state="readonly",
        width=39
    ).grid(row=10, column=1, pady=6)

    ttk.Label(
        keret, text="Felelős"
    ).grid(row=11, column=0, sticky="w", pady=6)

    ttk.Combobox(
        keret,
        textvariable=felelos,
        values=list(felelos_lista),
        state="readonly",
        width=39
    ).grid(row=11, column=1, pady=6)

    def am_mentes():
        adatbazis = None

        try:
            azonosito = valtozok["azonosito"].get().strip()
            megnevezes = valtozok["megnevezes"].get().strip()

            if not azonosito or not megnevezes:
                raise ValueError(
                    "Az azonosító és a megnevezés kötelező!"
                )

            if not helyiseg.get() or not felelos.get():
                raise ValueError(
                    "Válassz helyiséget és felelőst!"
                )

            h = helyiseg_lista[helyiseg.get()]
            f = felelos_lista[felelos.get()]

            mennyiseg = float(
                valtozok["mennyiseg"].get().replace(",", ".")
            )

            brutto = float(
                valtozok["brutto"].get().replace(",", ".")
            )

            konyv = float(
                valtozok["konyv"].get().replace(",", ".")
            )

            if mennyiseg <= 0:
                raise ValueError(
                    "A mennyiségnek pozitívnak kell lennie!"
                )

            if brutto < 0 or konyv < 0:
                raise ValueError(
                    "Az értékek nem lehetnek negatívak!"
                )

            datum = date.fromisoformat(
                valtozok["datum"].get().strip()
            ).isoformat()

            adatok = (
                azonosito,
                valtozok["eszkoz_szam"].get(),
                megnevezes,
                valtozok["gyartasi_szam"].get(),
                mennyiseg,
                valtozok["mertekegyseg"].get(),
                h[1],
                h[1],
                h[0],
                f[0],
                datum,
                brutto,
                "HUF",
                konyv,
                "HUF",
                valtozok["megjegyzes"].get()
            )

            adatbazis = AMAdatbazis()
            adatbazis.am_eszkoz_felvetele(adatok)

        except Exception as hiba:
            messagebox.showerror(
                "Hiba", str(hiba), parent=ablak
            )
            return

        finally:
            if adatbazis is not None:
                adatbazis.am_bezaras()

        messagebox.showinfo(
            "Sikeres mentés",
            "Az új eszköz rögzítése sikeres!",
            parent=ablak
        )

        ablak.destroy()
        frissites()

    gombok = ttk.Frame(keret)
    gombok.grid(row=12, column=0, columnspan=2, pady=25)

    ttk.Button(
        gombok,
        text="Mentés",
        command=am_mentes
    ).pack(side="left", padx=10)

    ttk.Button(
        gombok,
        text="Mégse",
        command=ablak.destroy
    ).pack(side="left", padx=10)


def am_eszkozok_megjelenitese(content):
    adatbazis = AMAdatbazis()

    adatok = adatbazis.kurzor.execute("""
        SELECT
            e.eszkozazonosito,
            e.megnevezes,
            e.gyartasi_szam,
            h.helyiseg_szam,
            f.teljes_nev,
            e.brutto_ertek
        FROM eszkozok e
        LEFT JOIN helyisegek h ON e.helyiseg_id = h.id
        LEFT JOIN felelosok f ON e.felelos_id = f.id
        ORDER BY e.eszkozazonosito
    """).fetchall()

    adatbazis.am_bezaras()

    oldal = ttk.Frame(content)
    oldal.pack(fill="both", expand=True)

    kereses = tk.StringVar()
    helyiseg = tk.StringVar(value="Összes")
    felelos = tk.StringVar(value="Összes")
    talalatok = tk.StringVar()

    def am_frissites():
        oldal.destroy()
        am_eszkozok_megjelenitese(content)

    gombkeret = ttk.Frame(oldal)
    gombkeret.pack(fill="x", padx=30, pady=5)

    ttk.Button(
        gombkeret,
        text="Új eszköz",
        command=lambda: am_uj_eszkoz(content, am_frissites)
    ).pack(side="left", padx=5)

    ttk.Button(
        gombkeret,
        text="Frissítés",
        command=am_frissites
    ).pack(side="left", padx=5)

    szurokeret = ttk.Frame(oldal)
    szurokeret.pack(fill="x", padx=30, pady=10)

    ttk.Label(szurokeret, text="Keresés:").grid(
        row=0, column=0, padx=5, pady=5
    )

    ttk.Entry(
        szurokeret,
        textvariable=kereses,
        width=40
    ).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

    ttk.Label(szurokeret, text="Helyiség:").grid(
        row=1, column=0, padx=5, pady=5
    )

    helyisegek = ["Összes"] + sorted({
        str(sor[3]) for sor in adatok if sor[3] is not None
    })

    ttk.Combobox(
        szurokeret,
        textvariable=helyiseg,
        values=helyisegek,
        state="readonly",
        width=20
    ).grid(row=1, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(szurokeret, text="Felelős:").grid(
        row=1, column=2, padx=5, pady=5
    )

    felelosok = ["Összes"] + sorted({
        str(sor[4]) for sor in adatok if sor[4] is not None
    })

    ttk.Combobox(
        szurokeret,
        textvariable=felelos,
        values=felelosok,
        state="readonly",
        width=25
    ).grid(row=1, column=3, padx=5, pady=5)

    keret = ttk.Frame(oldal)
    keret.pack(fill="both", expand=True, padx=30, pady=10)

    oszlopok = (
        "azonosito",
        "megnevezes",
        "gyartasi_szam",
        "helyiseg",
        "felelos",
        "ertek"
    )

    tabla = ttk.Treeview(
        keret,
        columns=oszlopok,
        show="headings"
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
        tabla.heading(oszlop, text=nev)
        tabla.column(oszlop, width=szelesseg)

    tabla.column("ertek", anchor="e")

    gorgeto = ttk.Scrollbar(
        keret,
        orient="vertical",
        command=tabla.yview
    )

    tabla.configure(yscrollcommand=gorgeto.set)

    gorgeto.pack(side="right", fill="y")
    tabla.pack(side="left", fill="both", expand=True)

    def am_szures(*args):
        for elem in tabla.get_children():
            tabla.delete(elem)

        keresett = kereses.get().casefold().strip()
        darab = 0

        for sor in adatok:
            szoveg = " ".join(
                str(ertek) for ertek in sor[:3]
                if ertek is not None
            ).casefold()

            if keresett not in szoveg:
                continue

            if helyiseg.get() != "Összes":
                if str(sor[3]) != helyiseg.get():
                    continue

            if felelos.get() != "Összes":
                if str(sor[4]) != felelos.get():
                    continue

            megjelenitett = [
                "" if ertek is None else ertek
                for ertek in sor
            ]

            if isinstance(sor[5], (int, float)):
                megjelenitett[5] = (
                    f"{sor[5]:,.0f} Ft".replace(",", " ")
                )

            tabla.insert("", tk.END, values=megjelenitett)
            darab += 1

        talalatok.set(
            f"Találatok száma: {darab} / {len(adatok)}"
        )

    def am_alaphelyzet():
        kereses.set("")
        helyiseg.set("Összes")
        felelos.set("Összes")

    ttk.Button(
        szurokeret,
        text="Alaphelyzet",
        command=am_alaphelyzet
    ).grid(row=0, column=3, padx=5, pady=5)

    ttk.Label(
        oldal,
        textvariable=talalatok
    ).pack(anchor="w", padx=30, pady=10)

    kereses.trace_add("write", am_szures)
    helyiseg.trace_add("write", am_szures)
    felelos.trace_add("write", am_szures)

    am_szures()