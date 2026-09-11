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
3. Kizárólag a 8417 FEOR-kódú dolgozók kerülnek a listába és az exportokba. A nem nulla keresetű dolgozók Nemzetközisként indulnak. A 0 keresetűek külön, nem szerkeszthető listában láthatók, kötelező Kihagyás besorolással. A táblázat Besorolás oszlopában módosítsd a kivételeket. A forrás esetleges korábbi N/B/H/O oszlopát az alkalmazás szándékosan nem használja.
4. Korábbi alkalmazás-exportból opcionálisan töltsd vissza a besorolásokat. Ez az összes jelenlegi besorolást lecseréli a törzsszám szerint egyező mentett értékekre. Az új, nem nulla keresetű személyek Nemzetközisként indulnak. Az aktuális havi nulla kereset felülírja a mentett besorolást. A korábbi automatikus nullás kizárás nem kerül át egy nem nullás hónapra.
5. Az eredmények azonnal frissülnek. A kihagyottak nem számítanak bele. A nullás létszámú besorolt személyek külön személyszámban követhetők.
6. Ellenőrzés után töltsd le az Excel- és PDF-kimutatást. Az Excel egyúttal a besorolások mentése is, később visszatölthető.

## Számítási szabályok

- Nemzetközis: alaplétszám × 1 a nemzetközi létszámba.
- Belföldes: alaplétszám × 1 a belföldi létszámba.
- Hibrid: alaplétszám × 0,8 nemzetközi és × 0,2 belföldi.
- Váltó / Ónódi: alaplétszám × 0,625 nemzetközi.
- Kihagyás: egyik létszámba sem kerül.

A súlyozott létszám nem a személyek darabszáma. A program az Agroorg `letsz` mezőjét veszi alapul, ezt nem számítja újra a belépési/kilépési dátumokból. Csak megjelenítéskor kerekít négy tizedesre.

Az import név szerint azonosítja a `torzsszam`, `nev`, `letsz` , `feorkód` és `kereset` oszlopokat. Hiányzó vagy hibás kereset nem minősül nullának, feldolgozási hibát okoz. Hibás létszám, hiányzó név/törzsszám vagy ismétlődő törzsszám esetén megáll, nem hagy ki sorokat észrevétlenül. Több jogviszonyú személyek összevonása nem automatikus.

## Export és megőrzés

Az Excel Összesítés, Részletezés, Kihagyottak és Besorolások munkalapokat tartalmaz. A Besorolások lap szerkesztéséből képletekkel frissül a Részletezés és az Összesítés. A Kihagyottak lap exportáláskori pillanatkép. Az alkalmazásba visszatöltött Excelből csak a törzsszámhoz kapcsolt besorolást vesszük át; az alaplétszám mindig az aktuális Agroorg-exportból származik.

A PDF ismétlődő táblázatfejlécekkel és külön kihagyott-listával készül. Az exportok színei: #2A3756, #F2E47D, #D1D5DB, #FFFFFF. Excel: Calibri; PDF: beágyazott DejaVu Sans az ékezetek hordozható megjelenítéséhez.

Nincs automatikus szerveroldali tartós személyadat-tárolás. A böngésző munkamenetének elvesztésekor a nem exportált módosítások elveszhetnek. A havi Excel-exportokat őrizd meg, így a korábbi hónapok állapota külön fájlban megmarad.

Hivatalos API-dokumentáció: https://docs.streamlit.io/develop/api-reference/data/st.data_editor

## Kötelező kizárás

A 0 keresetű, 8417 FEOR-kódú dolgozók a felületen külön, zárolt listában, az Excel és PDF exportban a kihagyottak között is megjelennek, „0 kereset miatt kihagyva” megjelöléssel. A szabály a számítás és az export során, illetve régi besorolások visszatöltésekor is érvényes. Az Excel számítási képletei nulla kereset mellett nem veszik figyelembe az esetleg átírt besorolást. Más FEOR-kódú dolgozók egyik személylistában vagy exportban sem szerepelnek. A nulla keresetből nem állapítjuk meg a távollét okát.
