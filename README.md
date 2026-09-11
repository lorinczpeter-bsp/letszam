# Bábolna Sped – Sofőrlétszám

Önálló Streamlit-alkalmazás Agroorg XLS/XLSX/CSV exporthoz.

## Indítás

Python 3.11 vagy 3.12 környezetben, a kicsomagolt könyvtárból:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Streamlit Community Cloud esetén a teljes könyvtár tartalmát töltsd fel a saját repository-ba. A belépő fájl `app.py`. Az `assets` mappa és a két TTF betűfájl is szükséges a magyar ékezeteket helyesen megjelenítő PDF-hez. A csomag nem tartalmaz személyes bemeneti adatokat.

## Használat

1. Töltsd fel a változtatás nélküli Agroorg-exportot.
2. Ellenőrizd a fájlnév alapján felajánlott elszámolási hónapot.
3. Mindenki Nemzetközisként indul. A táblázat Besorolás oszlopában módosítsd a kivételeket. A forrás esetleges korábbi N/B/H/O oszlopát az alkalmazás szándékosan nem használja.
4. Korábbi alkalmazás-exportból opcionálisan töltsd vissza a besorolásokat. Ez az összes jelenlegi besorolást lecseréli a törzsszám szerint egyező mentett értékekre. Az új személyek Nemzetközisként indulnak.
5. Az eredmények azonnal frissülnek. A kihagyottak nem számítanak bele. A nullás létszámú besorolt személyek külön személyszámban követhetők.
6. Ellenőrzés után töltsd le az Excel- és PDF-kimutatást. Az Excel egyúttal a besorolások mentése is, később visszatölthető.

## Számítási szabályok

- Nemzetközis: alaplétszám × 1 a nemzetközi létszámba.
- Belföldes: alaplétszám × 1 a belföldi létszámba.
- Hibrid: alaplétszám × 0,8 nemzetközi és × 0,2 belföldi.
- Váltó / Ónódi: alaplétszám × 0,625 nemzetközi.
- Kihagyás: egyik létszámba sem kerül.

A súlyozott létszám nem a személyek darabszáma. A program az Agroorg `letsz` mezőjét veszi alapul, ezt nem számítja újra a belépési/kilépési dátumokból. Csak megjelenítéskor kerekít négy tizedesre.

Az import név szerint azonosítja a `torzsszam`, `nev`, `letsz` és opcionális `feorkód` oszlopokat. Hibás létszám, hiányzó név/törzsszám vagy ismétlődő törzsszám esetén megáll, nem hagy ki sorokat észrevétlenül. Több jogviszonyú személyek összevonása nem automatikus.

## Export és megőrzés

Az Excel Összesítés, Részletezés, Kihagyottak és Besorolások munkalapokat tartalmaz. A Besorolások lap szerkesztéséből képletekkel frissül a Részletezés és az Összesítés. A Kihagyottak lap exportáláskori pillanatkép. Az alkalmazásba visszatöltött Excelből csak a törzsszámhoz kapcsolt besorolást vesszük át; az alaplétszám mindig az aktuális Agroorg-exportból származik.

A PDF ismétlődő táblázatfejlécekkel és külön kihagyott-listával készül. Az exportok színei: #2A3756, #F2E47D, #D1D5DB, #FFFFFF. Excel: Calibri; PDF: beágyazott DejaVu Sans az ékezetek hordozható megjelenítéséhez.

Nincs automatikus szerveroldali tartós személyadat-tárolás. A böngésző munkamenetének elvesztésekor a nem exportált módosítások elveszhetnek. A havi Excel-exportokat őrizd meg, így a korábbi hónapok állapota külön fájlban megmarad.

Hivatalos API-dokumentáció: https://docs.streamlit.io/develop/api-reference/data/st.data_editor
