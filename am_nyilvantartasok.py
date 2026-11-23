
import json
import uuid
import tkinter as tk
from pathlib import Path
from tkinter import ttk, messagebox
from am_database import AMAdatbazis


class AMNyilvantartasKezelo:
    def __init__(self):
        self.mappa = Path(__file__).resolve().parent
        self.adatmappa = self.mappa / "nyilvantartasok"
        self.index_utvonal = (
            self.mappa / "am_nyilvantartasok.json"
        )

        self.adatmappa.mkdir(
            parents=True,
            exist_ok=True
        )

        if not self.index_utvonal.exists():
            eredeti = self.mappa / "leltar.db"

            if not eredeti.is_file():
                raise FileNotFoundError(
                    "A meglévő leltar.db fájl nem található!\n"
                    "Az adatok védelme érdekében nem "
                    "hozunk létre helyette üres adatbázist."
                )

            self.am_index_mentes([
                {
                    "id": "eredeti",
                    "nev": "Eredeti eszköznyilvántartás",
                    "fajl": "leltar.db"
                }
            ])

    def am_index_lekerese(self):
        with self.index_utvonal.open(
            "r",
            encoding="utf-8"
        ) as fajl:
            adatok = json.load(fajl)

        if not isinstance(adatok, list):
            raise ValueError(
                "A nyilvántartások listája hibás!"
            )

        return adatok

    def am_index_mentes(self, adatok):
        atmeneti = self.index_utvonal.with_suffix(
            ".json.tmp"
        )

        with atmeneti.open(
            "w",
            encoding="utf-8"
        ) as fajl:
            json.dump(
                adatok,
                fajl,
                ensure_ascii=False,
                indent=4
            )

        atmeneti.replace(self.index_utvonal)

    def am_utvonal(self, rekord):
        fajlnev = rekord["fajl"]
        utvonal = (self.mappa / fajlnev).resolve()

        engedelyezett = (
            utvonal == (self.mappa / "leltar.db").resolve()
            or utvonal.parent == self.adatmappa.resolve()
        )

        if not engedelyezett:
            raise ValueError(
                "Nem engedélyezett adatbázis-útvonal!"
            )

        return utvonal

    def am_osszes(self):
        return self.am_index_lekerese()

    def am_uj(self, nev):
        nev = nev.strip()

        if not nev:
            raise ValueError(
                "Add meg az új nyilvántartás nevét!"
            )

        if len(nev) > 100:
            raise ValueError(
                "A nyilvántartás neve legfeljebb "
                "100 karakter lehet!"
            )

        adatok = self.am_index_lekerese()

        if any(
            rekord["nev"].casefold() == nev.casefold()
            for rekord in adatok
        ):
            raise ValueError(
                "Már létezik ilyen nevű nyilvántartás!"
            )

        azonosito = uuid.uuid4().hex
        fajlnev = (
            f"nyilvantartasok/{azonosito}.db"
        )
        utvonal = self.mappa / fajlnev

        adatbazis = None

        try:
            adatbazis = AMAdatbazis(
                utvonal=utvonal
            )
            adatbazis.am_bezaras()
            adatbazis = None

            rekord = {
                "id": azonosito,
                "nev": nev,
                "fajl": fajlnev
            }

            self.am_index_mentes(
                adatok + [rekord]
            )

        except Exception:
            if adatbazis is not None:
                adatbazis.am_bezaras()

            if utvonal.exists():
                utvonal.unlink()

            raise

        return rekord


def am_nyilvantartas_valaszto(
    root,
    kezelo,
    megnyitas,
    aktualis_id=None,
    inditas=False
):
    ablak = tk.Toplevel(root)
    ablak.title("AM – Nyilvántartások")
    ablak.geometry("710x540")
    ablak.minsize(620, 480)

    if not inditas:
        ablak.transient(root)

    ablak.grab_set()

    keret = ttk.Frame(
        ablak,
        padding=22
    )
    keret.pack(fill="both", expand=True)

    gombkeret = ttk.Frame(keret)
    gombkeret.pack(
        side="bottom",
        fill="x",
        pady=(15, 0)
    )

    ttk.Label(
        keret,
        text="Nyilvántartás kiválasztása",
        font=("Segoe UI", 18, "bold")
    ).pack(
        anchor="w",
        pady=(0, 8)
    )

    ttk.Label(
        keret,
        text=(
            "Válaszd ki, melyik önálló "
            "eszköznyilvántartással szeretnél dolgozni."
        )
    ).pack(
        anchor="w",
        pady=(0, 18)
    )

    tablakeret = ttk.Frame(keret)
    tablakeret.pack(
        fill="both",
        expand=True
    )

    tabla = ttk.Treeview(
        tablakeret,
        columns=("nev", "fajl"),
        show="headings",
        selectmode="browse",
        height=9
    )

    tabla.heading(
        "nev",
        text="Nyilvántartás neve"
    )
    tabla.heading(
        "fajl",
        text="Adatbázisfájl"
    )

    tabla.column("nev", width=310)
    tabla.column("fajl", width=330)

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

    ttk.Label(
        keret,
        text="Új nyilvántartás létrehozása",
        font=("Segoe UI", 11, "bold")
    ).pack(
        anchor="w",
        pady=(16, 8)
    )

    ujkeret = ttk.Frame(keret)
    ujkeret.pack(fill="x")

    uj_nev = tk.StringVar()

    ttk.Entry(
        ujkeret,
        textvariable=uj_nev,
        width=43
    ).pack(
        side="left",
        fill="x",
        expand=True,
        padx=(0, 10)
    )

    rekordok = {}

    def am_lista_frissites(kijelolt_id=None):
        nonlocal rekordok

        rekordok = {
            rekord["id"]: rekord
            for rekord in kezelo.am_osszes()
        }

        for elem in tabla.get_children():
            tabla.delete(elem)

        for rekord in rekordok.values():
            tabla.insert(
                "",
                "end",
                iid=rekord["id"],
                values=(
                    rekord["nev"],
                    rekord["fajl"]
                )
            )

        kivant_id = kijelolt_id or aktualis_id

        if (
            kivant_id is not None
            and tabla.exists(kivant_id)
        ):
            tabla.selection_set(kivant_id)
            tabla.focus(kivant_id)
            tabla.see(kivant_id)

        elif tabla.get_children():
            elso = tabla.get_children()[0]
            tabla.selection_set(elso)
            tabla.focus(elso)

    def am_uj_letrehozasa():
        try:
            rekord = kezelo.am_uj(
                uj_nev.get()
            )

            uj_nev.set("")
            am_lista_frissites(rekord["id"])

            messagebox.showinfo(
                "Sikeres létrehozás",
                f"Elkészült az új nyilvántartás:\n"
                f"{rekord['nev']}\n\n"
                "A nyilvántartás egyelőre üres.\n"
                "A Megnyitás gombbal válthatsz rá.",
                parent=ablak
            )

        except Exception as hiba:
            messagebox.showerror(
                "Létrehozási hiba",
                str(hiba),
                parent=ablak
            )

    ttk.Button(
        ujkeret,
        text="Új létrehozása",
        command=am_uj_letrehozasa
    ).pack(side="right")

    def am_megnyitas(event=None):
        kijeloles = tabla.selection()

        if not kijeloles:
            messagebox.showwarning(
                "Nincs kiválasztás",
                "Válassz egy nyilvántartást!",
                parent=ablak
            )
            return

        rekord = rekordok.get(kijeloles[0])

        if rekord is None:
            return

        sikeres = megnyitas(rekord)

        if sikeres:
            ablak.destroy()

    def am_bezaras():
        if inditas:
            root.destroy()
        else:
            ablak.destroy()

    ttk.Button(
        gombkeret,
        text="Kiválasztott nyilvántartás megnyitása",
        command=am_megnyitas
    ).pack(
        side="left",
        padx=(0, 10)
    )

    ttk.Button(
        gombkeret,
        text="Kilépés" if inditas else "Mégse",
        command=am_bezaras
    ).pack(side="right")

    tabla.bind(
        "<Double-1>",
        am_megnyitas
    )

    ablak.protocol(
        "WM_DELETE_WINDOW",
        am_bezaras
    )

    am_lista_frissites()
    return ablak
