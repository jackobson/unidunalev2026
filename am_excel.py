
from pathlib import Path
from datetime import date, datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from am_database import AMAdatbazis
from am_selejtezes import am_selejtezes_tabla


AM_EXCEL_OSZLOPOK = [
    "Telephely",
    "Eszköz telephelye",
    "Helyiségszám",
    "Helyiség megnevezése",
    "Felelős törzsszáma",
    "Felelős neve",
    "Felelős osztálya",
    "Aktiválás dátuma",
    "Főcsoport",
    "Eszközazonosító",
    "Eszközszám",
    "Megnevezés",
    "Mennyiség",
    "Mértékegység",
    "Gyártási szám",
    "Bruttó érték",
    "Bruttó pénznem",
    "Könyv szerinti érték",
    "Könyv szerinti pénznem",
    "Eredeti eszköz",
    "ASZ/BEB",
    "Leltárútmutatás",
    "KNYT",
    "Összevont"
]


def am_ertek(ertek):
    if isinstance(ertek, datetime):
        return ertek.date().isoformat()

    if isinstance(ertek, date):
        return ertek.isoformat()

    return ertek


def am_szoveg(ertek):
    return "" if ertek is None else str(ertek).strip()


def am_excel_beolvasas(fajl):
    munkafuzet = load_workbook(
        fajl,
        read_only=True,
        data_only=True
    )

    try:
        munkalap = munkafuzet.active

        if munkalap.max_column < 24:
            raise ValueError(
                "A fájl első munkalapján legalább "
                "24 oszlop szükséges."
            )

        adatok = []

        for sorszam, sor in enumerate(
            munkalap.iter_rows(
                min_row=2,
                min_col=1,
                max_col=24,
                values_only=True
            ),
            start=2
        ):
            ertekek = tuple(
                am_ertek(x) for x in sor
            )

            if all(
                am_szoveg(x) == ""
                for x in ertekek
            ):
                continue

            if (
                not am_szoveg(ertekek[9])
                or not am_szoveg(ertekek[11])
            ):
                raise ValueError(
                    f"A(z) {sorszam}. sorban nincs "
                    "eszközazonosító vagy megnevezés."
                )

            adatok.append(ertekek)

        if not adatok:
            raise ValueError(
                "A fájl nem tartalmaz importálható "
                "eszközöket."
            )

        return adatok

    finally:
        munkafuzet.close()


def am_excel_importalas(fajl):
    adatok = am_excel_beolvasas(fajl)
    adatbazis = AMAdatbazis()

    try:
        adatbazis.kurzor.execute(
            "BEGIN IMMEDIATE"
        )

        szam = adatbazis.kurzor.execute("""
            SELECT COUNT(*)
            FROM eszkozok
        """).fetchone()[0]

        if szam:
            raise ValueError(
                "Ebben a nyilvántartásban már "
                "vannak eszközök. "
                "A duplikációk elkerüléséért "
                "az importálás tiltott. "
                "Hozz létre egy új, üres "
                "nyilvántartást!"
            )

        helyisegek = {}
        felelosok = {}

        uj_helyisegek = 0
        uj_felelosok = 0

        meglevo_helyisegek = (
            adatbazis.kurzor.execute("""
                SELECT
                    id,
                    telephely,
                    helyiseg_szam
                FROM helyisegek
            """).fetchall()
        )

        for sor in meglevo_helyisegek:
            kulcs = (
                am_szoveg(sor[1]).casefold(),
                am_szoveg(sor[2]).casefold()
            )

            helyisegek[kulcs] = sor[0]

        meglevo_felelosok = (
            adatbazis.kurzor.execute("""
                SELECT
                    id,
                    sztsz,
                    teljes_nev,
                    osztaly
                FROM felelosok
            """).fetchall()
        )

        for sor in meglevo_felelosok:
            sztsz = am_szoveg(sor[1])

            if sztsz:
                kulcs = (
                    "sztsz",
                    sztsz.casefold()
                )
            else:
                kulcs = (
                    "nev",
                    am_szoveg(sor[2]).casefold(),
                    am_szoveg(sor[3]).casefold()
                )

            felelosok[kulcs] = sor[0]

        for sor in adatok:
            telephely = am_szoveg(sor[0])
            helyiseg_szam = am_szoveg(sor[2])
            helyiseg_nev = am_szoveg(sor[3])

            sztsz = am_szoveg(sor[4])
            teljes_nev = am_szoveg(sor[5])
            osztaly = am_szoveg(sor[6])

            helyiseg_id = None

            if (
                telephely
                or helyiseg_szam
                or helyiseg_nev
            ):
                kulcs = (
                    telephely.casefold(),
                    helyiseg_szam.casefold()
                )

                if kulcs not in helyisegek:
                    adatbazis.kurzor.execute("""
                        INSERT INTO helyisegek (
                            telephely,
                            helyiseg_szam,
                            megnevezes
                        )
                        VALUES (?, ?, ?)
                    """, (
                        telephely,
                        helyiseg_szam,
                        helyiseg_nev
                    ))

                    helyisegek[kulcs] = (
                        adatbazis.kurzor.lastrowid
                    )
                    uj_helyisegek += 1

                helyiseg_id = helyisegek[kulcs]

            felelos_id = None

            if sztsz or teljes_nev or osztaly:
                if sztsz:
                    kulcs = (
                        "sztsz",
                        sztsz.casefold()
                    )
                else:
                    kulcs = (
                        "nev",
                        teljes_nev.casefold(),
                        osztaly.casefold()
                    )

                if kulcs not in felelosok:
                    adatbazis.kurzor.execute("""
                        INSERT INTO felelosok (
                            sztsz,
                            teljes_nev,
                            osztaly
                        )
                        VALUES (?, ?, ?)
                    """, (
                        sztsz,
                        teljes_nev,
                        osztaly
                    ))

                    felelosok[kulcs] = (
                        adatbazis.kurzor.lastrowid
                    )
                    uj_felelosok += 1

                felelos_id = felelosok[kulcs]

            adatbazis.kurzor.execute("""
                INSERT INTO eszkozok (
                    telephely,
                    eszkoz_telephelye,
                    helyiseg_id,
                    felelos_id,
                    aktivalas_datuma,
                    focsoport,
                    eszkozazonosito,
                    eszkoz_szam,
                    megnevezes,
                    mennyiseg,
                    mertekegyseg,
                    gyartasi_szam,
                    brutto_ertek,
                    brutto_penznem,
                    konyv_szerinti_ertek,
                    konyv_penznem,
                    eredeti_eszkoz,
                    asz_beb,
                    leltarutmutatas,
                    knyt,
                    osszevont
                )
                VALUES (
                    ?, ?, ?, ?, ?, ?, ?,
                    ?, ?, ?, ?, ?, ?, ?,
                    ?, ?, ?, ?, ?, ?, ?
                )
            """, (
                sor[0],
                sor[1],
                helyiseg_id,
                felelos_id,
                *sor[7:24]
            ))

        adatbazis.kapcsolat.commit()

        return (
            len(adatok),
            uj_helyisegek,
            uj_felelosok
        )

    except Exception:
        adatbazis.kapcsolat.rollback()
        raise

    finally:
        adatbazis.am_bezaras()


def am_excel_exportalas(fajl):
    am_selejtezes_tabla()

    adatbazis = AMAdatbazis()

    try:
        adatok = adatbazis.kurzor.execute("""
            SELECT
                e.telephely,
                e.eszkoz_telephelye,
                h.helyiseg_szam,
                h.megnevezes,
                f.sztsz,
                f.teljes_nev,
                f.osztaly,
                e.aktivalas_datuma,
                e.focsoport,
                e.eszkozazonosito,
                e.eszkoz_szam,
                e.megnevezes,
                e.mennyiseg,
                e.mertekegyseg,
                e.gyartasi_szam,
                e.brutto_ertek,
                e.brutto_penznem,
                e.konyv_szerinti_ertek,
                e.konyv_penznem,
                e.eredeti_eszkoz,
                e.asz_beb,
                e.leltarutmutatas,
                e.knyt,
                e.osszevont
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
        adatbazis.am_bezaras()

    munkafuzet = Workbook()
    lap = munkafuzet.active
    lap.title = "Eszközök"

    lap.append(AM_EXCEL_OSZLOPOK)

    for sor in adatok:
        lap.append(list(sor))

    for cella in lap[1]:
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

    for index in range(1, 25):
        lap.column_dimensions[
            get_column_letter(index)
        ].width = (
            36 if index in (4, 6, 12) else 22
        )

    lap.row_dimensions[1].height = 32
    lap.freeze_panes = "A2"
    lap.auto_filter.ref = lap.dimensions

    for sor in lap.iter_rows(min_row=2):
        for index in (16, 18):
            sor[index - 1].number_format = (
                '#,##0.00'
            )

    osszesito = munkafuzet.create_sheet(
        "Összesítés"
    )

    osszesito.append([
        "Mutató",
        "Érték"
    ])

    osszesito.append([
        "Exportálás időpontja",
        datetime.now().strftime(
            "%Y-%m-%d %H:%M"
        )
    ])

    osszesito.append([
        "Aktív eszközrekordok",
        len(adatok)
    ])

    osszesito.append([
        "Bruttó érték összesen",
        sum(
            sor[15] or 0
            for sor in adatok
        )
    ])

    osszesito.column_dimensions["A"].width = 34
    osszesito.column_dimensions["B"].width = 26

    osszesito["B4"].number_format = (
        '#,##0.00'
    )

    for cella in osszesito[1]:
        cella.font = Font(
            bold=True,
            color="FFFFFF"
        )
        cella.fill = PatternFill(
            "solid",
            fgColor="243E57"
        )

    munkafuzet.save(fajl)

    return len(adatok)


def am_excel_megjelenitese(
    content,
    tipus,
    nyilvantartas_neve
):
    oldal = ttk.Frame(
        content,
        padding=25
    )
    oldal.pack(
        fill="both",
        expand=True
    )

    import_mod = tipus == "import"

    cim = (
        "Excel-importálás"
        if import_mod
        else "Excel-exportálás"
    )

    ttk.Label(
        oldal,
        text=cim,
        font=("Segoe UI", 17, "bold")
    ).pack(
        anchor="w",
        pady=(0, 12)
    )

    ttk.Label(
        oldal,
        text=(
            f"Aktív nyilvántartás: "
            f"{nyilvantartas_neve}"
        ),
        font=("Segoe UI", 11, "bold")
    ).pack(
        anchor="w",
        pady=(0, 20)
    )

    adatbazis = AMAdatbazis()

    try:
        darab = adatbazis.kurzor.execute("""
            SELECT COUNT(*)
            FROM eszkozok
        """).fetchone()[0]

    finally:
        adatbazis.am_bezaras()

    statusz = tk.StringVar(
        value=f"Jelenlegi eszközrekordok: {darab}"
    )

    ttk.Label(
        oldal,
        textvariable=statusz
    ).pack(
        anchor="w",
        pady=(0, 18)
    )

    utvonal = tk.StringVar()

    elonezet = tk.StringVar(
        value="Még nincs kiválasztott fájl."
    )

    def am_fajlvalasztas():
        if import_mod:
            fajl = filedialog.askopenfilename(
                title=(
                    "Importálandó Excel kiválasztása"
                ),
                filetypes=[
                    ("Excel-munkafüzet", "*.xlsx")
                ]
            )

        else:
            fajl = filedialog.asksaveasfilename(
                title=(
                    "Eszköznyilvántartás exportálása"
                ),
                defaultextension=".xlsx",
                initialfile=(
                    "AM_Eszkoznyilvantartas.xlsx"
                ),
                filetypes=[
                    ("Excel-munkafüzet", "*.xlsx")
                ]
            )

        if not fajl:
            return

        utvonal.set(fajl)

        if import_mod:
            try:
                adatok = am_excel_beolvasas(
                    fajl
                )

                helyisegek = len({
                    (
                        am_szoveg(sor[0]).casefold(),
                        am_szoveg(sor[2]).casefold()
                    )
                    for sor in adatok
                })

                felelosok = len({
                    (
                        am_szoveg(sor[4]).casefold()
                        or (
                            am_szoveg(sor[5]).casefold(),
                            am_szoveg(sor[6]).casefold()
                        )
                    )
                    for sor in adatok
                })

                elonezet.set(
                    f"Ellenőrzött sorok: "
                    f"{len(adatok)}\n"
                    f"Helyiség-kulcsok: "
                    f"{helyisegek}\n"
                    f"Felelős-kulcsok: "
                    f"{felelosok}\n"
                    "Betöltés csak üres "
                    "eszköznyilvántartásba "
                    "lehetséges."
                )

            except Exception as hiba:
                utvonal.set("")

                elonezet.set(
                    f"A fájl nem importálható: "
                    f"{hiba}"
                )

        else:
            elonezet.set(
                "Az aktív eszközök 24 oszlopos "
                "Excel-fájlba kerülnek."
            )

    ttk.Button(
        oldal,
        text=(
            "Excel-fájl kiválasztása"
            if import_mod
            else "Mentési hely kiválasztása"
        ),
        command=am_fajlvalasztas
    ).pack(
        anchor="w",
        pady=(0, 12)
    )

    ttk.Label(
        oldal,
        textvariable=utvonal,
        wraplength=800
    ).pack(
        anchor="w",
        pady=(0, 15)
    )

    ttk.Label(
        oldal,
        textvariable=elonezet,
        justify="left",
        wraplength=800
    ).pack(
        anchor="w",
        pady=(0, 22)
    )

    def am_vegrehajtas():
        if not utvonal.get():
            messagebox.showwarning(
                "Nincs fájl",
                "Először válassz Excel-fájlt!"
            )
            return

        if import_mod:
            db = AMAdatbazis()

            try:
                aktualis = db.kurzor.execute("""
                    SELECT COUNT(*)
                    FROM eszkozok
                """).fetchone()[0]

            finally:
                db.am_bezaras()

            if aktualis > 0:
                messagebox.showwarning(
                    "Az importálás letiltva",
                    "A nyilvántartás már tartalmaz "
                    "eszközöket. Hozz létre egy "
                    "új nyilvántartást az "
                    "importáláshoz!"
                )
                return

            if not messagebox.askyesno(
                "Importálás megerősítése",
                f"Nyilvántartás: "
                f"{nyilvantartas_neve}\n"
                f"Fájl: "
                f"{Path(utvonal.get()).name}\n\n"
                "Biztosan betöltöd az adatokat?"
            ):
                return

            try:
                eszkozok, helyisegek, felelosok = (
                    am_excel_importalas(
                        utvonal.get()
                    )
                )

                statusz.set(
                    f"Jelenlegi eszközrekordok: "
                    f"{eszkozok}"
                )

                messagebox.showinfo(
                    "Sikeres importálás",
                    f"Importált eszközök: "
                    f"{eszkozok}\n"
                    f"Új helyiségek: "
                    f"{helyisegek}\n"
                    f"Új felelősök: "
                    f"{felelosok}"
                )

            except Exception as hiba:
                messagebox.showerror(
                    "Importálási hiba",
                    str(hiba)
                )

        else:
            try:
                szam = am_excel_exportalas(
                    utvonal.get()
                )

                messagebox.showinfo(
                    "Sikeres exportálás",
                    f"Exportált aktív eszközök: "
                    f"{szam}\n"
                    f"Fájl: {utvonal.get()}"
                )

            except Exception as hiba:
                messagebox.showerror(
                    "Exportálási hiba",
                    str(hiba)
                )

    ttk.Button(
        oldal,
        text=(
            "Importálás indítása"
            if import_mod
            else "Excel-export indítása"
        ),
        command=am_vegrehajtas
    ).pack(
        anchor="w",
        pady=8
    )

    ttk.Label(
        oldal,
        text=(
            "Az importálás nem ír felül adatokat. "
            "A korábbi eseménynaplók nem "
            "kerülnek át."
            if import_mod
            else
            "Az export az aktív eszközöket "
            "tartalmazza; az archivált és "
            "elszállított tételeket nem."
        ),
        wraplength=800
    ).pack(
        anchor="w",
        pady=(20, 0)
    )
