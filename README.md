# AM – Leltárkezelő rendszer

Készítette: Abaffy Miklós  
Intézmény: Dunaújvárosi Egyetem  
Tantárgy: Szkript nyelvek  
Programozási nyelv: Python

## 1. A feladat bemutatása

A beadandó feladatomhoz egy Python nyelven működő leltárkezelő programot készítettem.

A program célja, hogy egy eszközállomány adatait egyszerűen lehessen nyilvántartani és kezelni. Lehetőség van új eszközök felvételére, a meglévők módosítására, leltározására, áthelyezésére és selejtezésére.

A program grafikus felhasználói felülettel rendelkezik, így a legtöbb művelet gombokkal, táblázatokkal és külön adatbeviteli ablakokkal végezhető el.

Az alkalmazás több önálló nyilvántartást is képes kezelni. Ezek külön adatbázisban tárolják az eszközöket és a hozzájuk tartozó információkat.

A programhoz tartozó mintaadatok fiktívek, kizárólag a működés bemutatására szolgálnak.

## 2. Felhasznált technológiák

A program elkészítéséhez az alábbi modulokat és technológiákat használtam:

- Python: a program működésének megvalósítása.
- Tkinter és ttk: a grafikus felület kialakítása.
- SQLite (sqlite3): az eszközök és a hozzájuk kapcsolódó adatok tárolása.
- openpyxl: Excel-fájlok beolvasása és létrehozása.
- pathlib: fájlok és könyvtárak elérési útjainak kezelése.
- json: a különböző nyilvántartások adatainak kezelése.
- datetime: dátumok és időpontok kezelése.
- Git és GitHub: a forráskód verziókezelése.

## 3. A program funkciói

### 3.1. Nyilvántartások kezelése

A program indításakor ki lehet választani, hogy melyik nyilvántartással szeretnénk dolgozni.

Új nyilvántartás is létrehozható, amely saját SQLite-adatbázist kap. Az egyes nyilvántartások adatai egymástól függetlenek.

A nyilvántartások között a program futása közben is lehet váltani.

### 3.2. Eszköznyilvántartás

Az eszköznyilvántartásban megtekinthetők az aktuálisan nyilvántartott eszközök.

A következő műveletek végezhetők el:

- Új eszköz felvétele.
- Meglévő eszköz adatainak módosítása.
- Keresés eszközazonosító és egyéb adatok alapján.
- Szűrés helyiség és felelős szerint.
- Eszköz archiválása és visszaállítása.
- Eszköz selejtezése.

### 3.3. Leltározás

A leltározási menüpontban az eszközök ellenőrzése végezhető el.

A kiválasztott eszközre duplán kattintva külön ablak jelenik meg, amelyben rögzíthető a leltározás eredménye.

A lehetséges állapotok:

- Megvan
- Hiányzik
- Sérült

Az ellenőrzésekhez megjegyzés is megadható. A korábbi leltározási bejegyzések visszakereshetők.

A leltározás eredménye Excel-fájlba is exportálható.

### 3.4. Eszközök áthelyezése

Az eszközök másik helyiségbe helyezhetők át, illetve másik felelőshöz rendelhetők.

Az áthelyezés során megadható a dátum és az áthelyezés indoka.

A program elt