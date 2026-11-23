import sqlite3
from pathlib import Path


class AMAdatbazis:
    def __init__(self):
        self.utvonal = Path(__file__).parent / "leltar.db"
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
                FOREIGN KEY (helyiseg_id) REFERENCES helyisegek(id),
                FOREIGN KEY (felelos_id) REFERENCES felelosok(id)
            )
        """)

        self.kurzor.execute("""
            CREATE TABLE IF NOT EXISTS leltarozas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                eszkoz_id INTEGER NOT NULL,
                allapot TEXT NOT NULL,
                datum TEXT NOT NULL,
                megjegyzes TEXT,
                FOREIGN KEY (eszkoz_id) REFERENCES eszkozok(id)
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
                FOREIGN KEY (eszkoz_id) REFERENCES eszkozok(id),
                FOREIGN KEY (regi_helyiseg_id) REFERENCES helyisegek(id),
                FOREIGN KEY (uj_helyiseg_id) REFERENCES helyisegek(id),
                FOREIGN KEY (regi_felelos_id) REFERENCES felelosok(id),
                FOREIGN KEY (uj_felelos_id) REFERENCES felelosok(id)
            )
        """)

        self.kapcsolat.commit()

    def am_helyisegek_lekerese(self):
        return self.kurzor.execute("""
            SELECT id, telephely, helyiseg_szam, megnevezes
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

    def am_bezaras(self):
        self.kapcsolat.close()


if __name__ == "__main__":
    adatbazis = AMAdatbazis()
    print("Az adatbázistáblák sikeresen létrejöttek!")
    adatbazis.am_bezaras()