from openpyxl import load_workbook
from pathlib import Path
from am_database import AMAdatbazis


def am_excel_beolvasas(fajl):
    munkafuzet = load_workbook(fajl, read_only=True, data_only=True)
    munkalap = munkafuzet.active

    adatok = list(munkalap.iter_rows(min_row=2, values_only=True))

    munkafuzet.close()

    return adatok


def am_excel_importalas(fajl):
    adatok = am_excel_beolvasas(fajl)
    adatbazis = AMAdatbazis()
    kurzor = adatbazis.kurzor

    try:
        kurzor.execute("SELECT COUNT(*) FROM eszkozok")

        if kurzor.fetchone()[0] > 0:
            print("Az adatbázis már tartalmaz eszközöket.")
            return

        helyisegek = {}
        felelosok = {}

        for sor in adatok:
            helyiseg = (sor[0], sor[2], sor[3])
            felelos = (sor[4], sor[5], sor[6])

            if helyiseg not in helyisegek:
                kurzor.execute("""
                    INSERT INTO helyisegek
                    (telephely, helyiseg_szam, megnevezes)
                    VALUES (?, ?, ?)
                """, helyiseg)

                helyisegek[helyiseg] = kurzor.lastrowid

            if felelos not in felelosok:
                kurzor.execute("""
                    INSERT INTO felelosok
                    (sztsz, teljes_nev, osztaly)
                    VALUES (?, ?, ?)
                """, felelos)

                felelosok[felelos] = kurzor.lastrowid

            kurzor.execute("""
                INSERT INTO eszkozok (
                    telephely, eszkoz_telephelye,
                    helyiseg_id, felelos_id,
                    aktivalas_datuma, focsoport,
                    eszkozazonosito, eszkoz_szam,
                    megnevezes, mennyiseg, mertekegyseg,
                    gyartasi_szam, brutto_ertek,
                    brutto_penznem, konyv_szerinti_ertek,
                    konyv_penznem, eredeti_eszkoz,
                    asz_beb, leltarutmutatas, knyt,
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
                helyisegek[helyiseg],
                felelosok[felelos],
                str(sor[7]) if sor[7] is not None else None,
                *sor[8:24]
            ))

        adatbazis.kapcsolat.commit()

        print("Az Excel importálása sikeres!")
        print("Importált eszközök:", len(adatok))
        print("Helyiségek száma:", len(helyisegek))
        print("Felelősök száma:", len(felelosok))

    except Exception as hiba:
        adatbazis.kapcsolat.rollback()
        print("Importálási hiba:", hiba)

    finally:
        adatbazis.am_bezaras()


if __name__ == "__main__":
    fajl = Path(__file__).parent / "Leltar_2026.xlsx"
    am_excel_importalas(fajl)