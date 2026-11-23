# unidunalev2026

Script nyelvek



&#x20;AM – Leltárkezelő rendszer



Készítette: Abaffy Miklós  

Intézmény: Dunaújvárosi Egyetem  

Tantárgy: Szkript nyelvek  

Programozási nyelv: Python



&#x20;1. A program bemutatása



Az AM Leltárkezelő rendszer egy Python nyelven fejlesztett, grafikus felhasználói felülettel rendelkező eszköznyilvántartó alkalmazás.



A program célja az eszközállomány rendszerezett kezelése, az eszközök mozgásának követése, a leltározási folyamat támogatása és különféle kimutatások készítése.



Az alkalmazás több, egymástól független nyilvántartást is képes kezelni. Minden nyilvántartás külön SQLite-adatbázissal rendelkezik.



A projekt demonstrációs célból fiktív adatokat tartalmaz.



&#x20;2. Felhasznált technológiák



\- Python – a program működésének megvalósítása

\- Tkinter és ttk – grafikus felhasználói felület

\- SQLite – adatbázis-kezelés

\- openpyxl – Excel-fájlok importálása és exportálása

\- pathlib – fájlok és könyvtárak kezelése

\- json – a nyilvántartások listájának tárolása

\- datetime – dátumok és időpontok kezelése

\- Git és GitHub – verziókezelés



&#x20;3. A program funkciói



&#x20;3.1. Nyilvántartások kezelése



Az alkalmazás indításakor kiválasztható a használni kívánt nyilvántartás. Új, önálló nyilvántartás is létrehozható.



&#x20;3.2. Eszköznyilvántartás



\- Új eszköz felvétele

\- Meglévő eszköz adatainak módosítása

\- Keresés és szűrés

\- Archiválás és visszaállítás

\- Selejtezési folyamat kezelése



&#x20;3.3. Leltározás



Az eszközök leltározási állapota külön ablakban rögzíthető.



Lehetséges állapotok:

\- Megvan

\- Hiányzik

\- Sérült



A korábbi leltározási bejegyzések megtekinthetők. A leltározási eredmények Excelbe exportálhatók.



&#x20;3.4. Eszközök áthelyezése



Az eszközök másik helyiséghez vagy felelőshöz rendelhetők.



A program nyilvántartja az áthelyezések dátumát, korábbi és új adatait, valamint az áthelyezés indokát.



&#x20;3.5. Helyiségek és felelősök



Új helyiségek és felelősök rögzíthetők.



A kiválasztott helyiséghez vagy felelőshöz tartozó eszközök dupla kattintással megjeleníthetők.



Az eszközlisták Excelbe exportálhatók vagy nyomtatásra előkészíthetők.



&#x20;3.6. Statisztikák



Az alkalmazás többek között az alábbi adatokat jeleníti meg:



\- Aktív és archivált eszközök száma

\- Elszállított eszközök száma

\- Eszközállomány bruttó értéke

\- Eszközök megoszlása helyiségenként és felelősönként

\- Leltározási eredmények

\- Selejtezési és áthelyezési adatok



A statisztikai eredmények grafikus diagramokon és Excel-kimutatásokban is megjeleníthetők.



&#x20;3.7. Excel-importálás és -exportálás



Az alkalmazás 24 oszlopos Excel-formátumból képes eszközadatokat importálni.



Az importálás csak olyan nyilvántartásba engedélyezett, amely még nem tartalmaz eszközöket.



Az aktív eszközállomány Excel-fájlba exportálható.



&#x20;4. Telepítés



A program futtatásához Python és az `openpyxl` csomag szükséges.



A projekt letöltése:



git clone https://github.com/jackobson/unidunalev2026.git

cd unidunalev2026





Virtuális környezet létrehozása:



python -m venv .venv



Windows alatt a virtuális környezet aktiválása:



.venv\\Scripts\\Activate.ps1



A szükséges csomag telepítése:



python -m pip install -r requirements.txt



A program indítása:



python main.py



&#x20;5. Használat



A program indításakor megjelenik a nyilvántartásválasztó ablak.



A felhasználó kiválaszthatja a meglévő nyilvántartást, vagy létrehozhat egy újat.



A nyilvántartás megnyitása után a bal oldali menüből érhetők el a különböző funkciók.



Az eszközök leltározása és áthelyezése külön adatbeviteli ablakokban történik.



&#x20;6. Adattárolás



Az alkalmazás SQLite-adatbázisokat használ.



A `leltar.db` tartalmazza az eredeti demonstrációs nyilvántartást.



Az újonnan létrehozott nyilvántartások adatbázisai a `nyilvantartasok` mappába kerülnek.



Az `am\_nyilvantartasok.json` fájl a nyilvántartások neveit és adatbázisfájljainak elérési útját tartalmazza.



&#x20;7. A projekt felépítése



| Fájl | Feladat |

|---|---|

| `main.py` | A grafikus főablak és a navigáció |

| `am\_database.py` | SQLite-adatbázis-kezelés |

| `am\_nyilvantartasok.py` | Több nyilvántartás kezelése |

| `am\_eszkozok.py` | Eszközök nyilvántartása |

| `am\_leltarozas.py` | Leltározási műveletek |

| `am\_athelyezes.py` | Eszközök áthelyezése |

| `am\_selejtezes.py` | Selejtezési folyamat |

| `am\_kimutatasok.py` | Helyiségek és felelősök kimutatásai |

| `am\_statisztikak.py` | Statisztikák és diagramok |

| `am\_excel.py` | Excel-adatok importálása és exportálása |



&#x20;8. Megjegyzések



Az alkalmazás oktatási célból készült.



A mellékelt demonstrációs nyilvántartás fiktív adatokat tartalmaz.



A projekt továbbfejlesztési lehetőségei közé tartozik a felhasználói jogosultságkezelés és a részletesebb riportkészítés.

