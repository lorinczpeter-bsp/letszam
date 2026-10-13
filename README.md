# Bábolna Sped – Sofőrlétszám

Önálló Streamlit-alkalmazás Agroorg XLS/XLSX/CSV exporthoz. Python 3.11 vagy 3.12 környezetben:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Streamlit Community Cloud esetén a teljes csomagot töltsd fel a repository-ba, belépő fájl: `app.py`. Az `assets` betűfájljai is szükségesek. A csomag nem tartalmaz személyes bemeneti adatokat.

## Használat

1. Töltsd fel a változtatás nélküli Agroorg-exportot és ellenőrizd a hónapot.
2. Minden dolgozó megjelenik. A 8417 FEOR-kódú, nem nulla keresetű személy alapbesorolása Nemzetközis. Más FEOR esetén Kihagyás: ezt módosíthatod, ha az illető sofőrként dolgozott.
3. A 0 keresetűek külön listában jelennek meg. Besorolásuk mindig „0 Ft kereset”, nem módosítható, nem számítanak bele a létszámba. A Megjegyzés mezőben rögzíthető a távollét oka, legfeljebb 500 karakterrel. Az alkalmazás nem következtet a távollét okára.
4. A korábbi Excel-exportból törzsszám alapján visszatölthetők a besorolások. A régi „Váltó / Ónódi” értéket „Ónódi”-ként veszi át. Az aktuális hónap nulla keresete mindig felülírja a visszatöltött besorolást. Korábbi automatikus nullás kizárás után az új hónap alapbesorolása a FEOR alapján áll vissza.
5. A megjegyzések az Excel-exportban megmaradnak, és ugyanarra a hónapra visszatölthetők. Más hónap megjegyzését nem veszi át a program.
6. A Havi belépők fülön vagy a csoportfüleken módosítsd a besorolásokat, majd kattints a Mentés és újraszámítás gombra. Az Excel- és PDF-export mindig a mentett állapotból készül.

## Számítás

- Nemzetközis: alaplétszám × 1 nemzetközi.
- Belföldes: alaplétszám × 1 belföldi.
- Hibrid: alaplétszám × 0,8 nemzetközi és × 0,2 belföldi.
- Ónódi: alaplétszám × 0,625 nemzetközi.
- 4-kezes: alaplétszám × 0,75 nemzetközi. Két teljes alaplétszámú sofőr együtt 1,5 fő.
- Kihagyás és 0 Ft kereset: nem számítanak bele.

A számítás alapja az Agroorg `letsz` mezője. A súlyozott létszám nem a személyek darabszáma. A köztes értékeket nem kerekíti a program; megjelenítéskor négy tizedest használ.

Kötelező oszlopok: `torzsszam`, `nev`, `letsz`, `feorkód`, `kereset`. Hibás létszám vagy kereset, hiányzó név/törzsszám és ismétlődő törzsszám esetén megáll a feldolgozás. A hiányzó kereset nem nulla. Az eredeti fájl kézi besorolási oszlopát nem használja.

## Export

Az Excel lapjai: Összesítés, Részletezés, Kihagyottak, Besorolások. A Besorolások szerkesztése képletekkel frissíti a Részletezést és az Összesítést. A nulla kereset mellett átírt besorolás nem változtatja meg a számítást. A Kihagyottak lap az exportáláskori pillanatkép. A megjegyzések az Excelben és a PDF-ben is megjelennek.

Minden névrész nagy kezdőbetűvel, a többi betű kisbetűvel jelenik meg. Az egyeztetés továbbra is törzsszám alapján történik.

Arculati színek: #2A3756, #F2E47D, #D1D5DB, #FFFFFF. Excel: Calibri. PDF: beágyazott DejaVu Sans.

Nincs automatikus tartós szerveroldali mentés. A havi Excel-exportot őrizd meg: ez tárolja a besorolásokat és a megjegyzéseket is.


## Hóközi be- és kilépés ellenőrzése

Az alkalmazás továbbra is az Agroorg `letsz` mezőjéből számol, nem osztja/szorozza azt újra a jogviszony napjaival. A `jogvkezd` és `jogvvég` oszlopokat, ha rendelkezésre állnak, ellenőrzéshez használja. Naptári összehasonlításkor a be- és kilépés napja is beleszámít. Ez összehasonlító módszer, nem az Agroorg számítási szabályának állítása. A részmunkaidő, távollétek és egyéb bérszámfejtési szabályok miatt az eltérés nem feltétlenül hiba.

Jelzések: 1 fő feletti alaplétszám; negatív kereset; nem nulla kereset nulla alaplétszámmal; hónapon kívüli jogviszony pozitív alaplétszámmal; hibás vagy fordított dátum; hóközi időarány eltérése. Az ellenőrzés nem írja át az alaplétszámot. A hiányzó belépési dátum korlátozza az ellenőrzést. A hónap átállítása önmagában nem számítja újra az Agroorg-adatokat.

## Csoportos nézet

Külön fül: Belföldes, Nemzetközis, Hibrid, Ónódi, 4-kezes, Nem számolt. A nulla alaplétszámúak is a Nem számolt csoportban szerepelnek, a besorolásuk megtartásával. Minden dolgozó pontosan egy csoportba kerül. A hibrid személy külön csoportban van, miközben a hozzájárulása 80–20%-ban a két fő összesítés része.

A PDF külön szakaszokat, az Excel külön csoportlapokat és Ellenőrzés munkalapot tartalmaz. A csoportlapok exportáláskori pillanatképek: Excelben végzett utólagos módosítás után frissítésükhöz visszatöltés és új export kell. A Részletezés és Összesítés képletekkel frissül.

A másik hónapból visszatöltött adatok nem tarthatják meg az előző megjegyzéseket. A mentett felületi módosítások törzsszámhoz kapcsolódnak, így átsoroláskor nem kerülnek át másik személyre.


## Havi belépők és nyitott jogviszony

Az üres, kötőjeles vagy „- -” kilépési mező nyitott jogviszonyt jelent, nem dátumhibát. Csak a ténylegesen kitöltött kilépési dátumot ellenőrizzük. A kiválasztott hónap belépői külön Havi belépők fülön jelennek meg, belépési dátum szerinti sorrendben. A nézet az Excelben külön munkalapként, a PDF-ben külön szakaszként is szerepel. Ez kiegészítő nézet: a dolgozók a kategóriájukban is megjelennek, de az összesítés egyszer számolja őket.

A Havi belépők fülön közvetlenül szerkeszthető a besorolás. A nulla keresetű belépőknél csak a megjegyzés módosítható. Mentés után a kategóriafülek és exportok is az új besorolást mutatják. Ha egy mezőt két fülön eltérően módosítasz, a program jelzi az ütközést; állítsd azonosra a két értéket és ments újra.
