
import sqlite3
from pathlib import Path


class AMAdatbazis:
    am_alap_utvonal = Path(__file__).resolve().parent / "leltar.db"
    am_aktiv_utvonal = am_alap_utvonal

    @classmethod
    def am_aktiv_utvonal_beallitasa(cls, utvonal):
        uj_utvonal = Path(utvonal).resolve()

        if not uj_utvonal.is_file():
            raise FileNotFoundError(
                f"Az adatbázis nem található:\n{uj_utvonal}"
            )

        cls.am_aktiv_utvonal = uj_utvonal

    @classmethod
    def am_aktiv_utvonal_lekerese(cls):
        return cls.am_aktiv_utvonal

    def __init__(self, utvonal=None):
        if utvonal is None:
            self.utvonal = self.am_aktiv_utvonal
        else:
            self.utvonal = Path(utvonal).resolve()

        self.utvonal.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        self.kapcsolat = sqlite3.connect(self.utvonal)
        self.kapcsolat.execute("PRAGMA foreign_keys = ON")
        self.kurzor = self.kapcsolat.cursor()
        self.am_tablak_letrehozasa()

    def am_tablak_letrehozasa(self):
        self.kurzor.execute("""
            CREATE TABLE IF NOT EXISTS helyisegek (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                telephely TEXT,
                helyiseg_szam TEXT,
                megnevezes TEXT
            )
        """)

        self.kurzor.execute("""
            CREATE TABLE IF NOT EXISTS felelosok (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sztsz TEXT,
                teljes_nev TEXT,
                osztaly TEXT
            )
        """)

        self.kurzor.execute("""
            CREATE TABLE IF NOT EXISTS eszkozok (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                telephely TEXT,
                eszkoz_telephelye TEXT,
                helyiseg_id INTEGER,
                felelos_id INTEGER,
                aktivalas_datuma TEXT,
                focsoport TEXT,
                eszkozazonosito TEXT,
                eszkoz_szam TEXT,
                megnevezes TEXT,
                mennyiseg REAL,
                mertekegyseg TEXT,
                gyartasi_szam TEXT,
                brutto_ertek REAL,
                brutto_penznem TEXT,
                konyv_szerinti_ertek REAL,
                konyv_penznem TEXT,
                eredeti_eszkoz TEXT,
                asz_beb TEXT,
                leltarutmutatas TEXT,
                knyt TEXT,
                osszevont TEXT,
                megjegyzes TEXT,
                archivalt INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY (helyiseg_id)
                    REFERENCES helyisegek(id),
                FOREIGN KEY (felelos_id)
                    REFERENCES felelosok(id)
            )
        """)

        self.kurzor.execute("""
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

        self.kurzor.execute("""
            CREATE TABLE IF NOT EXISTS athelyezesek (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                eszkoz_id INTEGER NOT NULL,
                regi_helyiseg_id INTEGER,
                uj_helyiseg_id INTEGER,
                regi_felelos_id INTEGER,
                uj_felelos_id INTEGER,
                datum TEXT NOT NULL,
                megjegyzes TEXT,
                FOREIGN KEY (eszkoz_id)
                    REFERENCES eszkozok(id),
                FOREIGN KEY (regi_helyiseg_id)
                    REFERENCES helyisegek(id),
                FOREIGN KEY (uj_helyiseg_id)
                    REFERENCES helyisegek(id),
                FOREIGN KEY (regi_felelos_id)
                    REFERENCES felelosok(id),
                FOREIGN KEY (uj_felelos_id)
                    REFERENCES felelosok(id)
            )
        """)

        oszlopok = [
            sor[1]
            for sor in self.kurzor.execute(
                "PRAGMA table_info(eszkozok)"
            ).fetchall()
        ]

        if "archivalt" not in oszlopok:
            self.kurzor.execute("""
                ALTER TABLE eszkozok
                ADD COLUMN archivalt INTEGER NOT NULL DEFAULT 0
            """)

        self.kapcsolat.commit()

    def am_helyisegek_lekerese(self):
        return self.kurzor.execute("""
            SELECT
                id,
                telephely,
                helyiseg_szam,
                megnevezes
            FROM helyisegek
            ORDER BY telephely, helyiseg_szam
        """).fetchall()

    def am_felelosok_lekerese(self):
        return self.kurzor.execute("""
            SELECT id, teljes_nev, sztsz
            FROM felelosok
            ORDER BY teljes_nev
        """).fetchall()

    def am_eszkoz_felvetele(self, adatok):
        try:
            self.kurzor.execute("""
                INSERT INTO eszkozok (
                    eszkozazonosito,
                    eszkoz_szam,
                    megnevezes,
                    gyartasi_szam,
                    mennyiseg,
                    mertekegyseg,
                    telephely,
                    eszkoz_telephelye,
                    helyiseg_id,
                    felelos_id,
                    aktivalas_datuma,
                    brutto_ertek,
                    brutto_penznem,
                    konyv_szerinti_ertek,
                    konyv_penznem,
                    megjegyzes
                )
                VALUES (
                    ?, ?, ?, ?, ?, ?, ?, ?,
                    ?, ?, ?, ?, ?, ?, ?, ?
                )
            """, adatok)

            self.kapcsolat.commit()

        except Exception:
            self.kapcsolat.rollback()
            raise

    def am_eszkoz_lekerese(self, eszkoz_id):
        return self.kurzor.execute("""
            SELECT
                id,
                eszkozazonosito,
                eszkoz_szam,
                megnevezes,
                gyartasi_szam,
                mennyiseg,
                mertekegyseg,
                telephely,
                eszkoz_telephelye,
                helyiseg_id,
                felelos_id,
                aktivalas_datuma,
                brutto_ertek,
                brutto_penznem,
                konyv_szerinti_ertek,
                konyv_penznem,
                megjegyzes
            FROM eszkozok
            WHERE id = ?
        """, (eszkoz_id,)).fetchone()

    def am_eszkoz_modositasa(self, eszkoz_id, adatok):
        try:
            self.kurzor.execute("""
                UPDATE eszkozok
                SET
                    eszkozazonosito = ?,
                    eszkoz_szam = ?,
                    megnevezes = ?,
                    gyartasi_szam = ?,
                    mennyiseg = ?,
                    mertekegyseg = ?,
                    telephely = ?,
                    eszkoz_telephelye = ?,
                    helyiseg_id = ?,
                    felelos_id = ?,
                    aktivalas_datuma = ?,
                    brutto_ertek = ?,
                    brutto_penznem = ?,
                    konyv_szerinti_ertek = ?,
                    konyv_penznem = ?,
                    megjegyzes = ?
                WHERE id = ? AND archivalt = 0
            """, (*adatok, eszkoz_id))

            if self.kurzor.rowcount != 1:
                raise ValueError(
                    "A módosítandó aktív eszköz nem található!"
                )

            self.kapcsolat.commit()

        except Exception:
            self.kapcsolat.rollback()
            raise

    def am_eszkoz_archivalasa(self, eszkoz_id):
        try:
            self.kurzor.execute("""
                UPDATE eszkozok
                SET archivalt = 1
                WHERE id = ? AND archivalt = 0
            """, (eszkoz_id,))

            if self.kurzor.rowcount != 1:
                raise ValueError(
                    "Az aktív eszköz nem található!"
                )

            self.kapcsolat.commit()

        except Exception:
            self.kapcsolat.rollback()
            raise

    def am_eszkoz_visszaallitasa(self, eszkoz_id):
        try:
            self.kurzor.execute("""
                UPDATE eszkozok
                SET archivalt = 0
                WHERE id = ? AND archivalt = 1
            """, (eszkoz_id,))

            if self.kurzor.rowcount != 1:
                raise ValueError(
                    "Az archivált eszköz nem található!"
                )

            self.kapcsolat.commit()

        except Exception:
            self.kapcsolat.rollback()
            raise

    def am_archivalt_eszkozok_lekerese(self):
        return self.kurzor.execute("""
            SELECT id, eszkozazonosito, megnevezes
            FROM eszkozok
            WHERE archivalt = 1
            ORDER BY eszkozazonosito, id
        """).fetchall()

    def am_bezaras(self):
        self.kapcsolat.close()


if __name__ == "__main__":
    adatbazis = AMAdatbazis()
    print(
        "Az adatbázis frissítése sikeres!\n"
        f"Aktív adatbázis: {adatbazis.utvonal}"
    )
    adatbazis.am_bezaras()
