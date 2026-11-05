# KerVez Akciók WebPart - Funkcionális Dokumentáció

Ez a dokumentum a valós projektháztartás alapján összegzi az összes megvalósított funkcionalitást.

## 1. Akciók Listázása és Megjelenítése

**Funkció:** Akciók kártyalayoutban történő megjelenítése

- Összes akció megtekintése kártyarendszerben
- Lapozás 20 tétel/oldal
- Kártyaadatok: cím, típus, dátumok, felelős személy, státusz
- Piszkozat akciók rejtése nem-szerkesztőktől
- Publikálási dátum ellenőrzése: jövőbeli publikálási dátummal rendelkező akciók nem látszanak nem-szerkesztőknek

**Megvalósítás:** AkciokList.tsx, pnpListServices.ts

## 2. Szűrési Lehetőségek

**Funkció:** Valós idejű szűrés több feltétel alapján

- **Akció típus szűrés** - rendelkezésre álló típusok közül választható
- **Dátum szűrés (6 opció):**
  - Nincs szűrés
  - Következő 1 hét
  - Következő 1 hónap
  - Következő 3 hónap
  - Előző 1 hónap
  - Előző 3 hónap
- **Lejárt akciók elrejtése** - toggle funkció
- **Többszörös szűrés** - az összes szűrő kombinálható
- Szűrő állapot az oldal frissítésénél elvész (session-scoped)

**Megvalósítás:** AkciokList.tsx, AkciokListHeader.tsx, GlobalStore.ts

## 3. Rendezési Opciók

**Funkció:** Multifunkcionális rendezési lehetőség

- Akció típusa (A-Z)
- Akció típusa (Z-A)
- Kezdési idő (legfrissebb elöl - alapértelmezett)
- Kezdési idő (legrégebbi elöl)
- Kapcsolattartó (A-Z)
- Kapcsolattartó (Z-A)
- Rendezés nem perzisztens az oldal frissítésénél

**Megvalósítás:** AkciokList.tsx, AkciokListHeader.tsx

## 4. Akciók Létrehozása

**Funkció:** Új akció létrehozása szerkesztőknek

- Akció cím megadása (kötelező)
- Akció típusa (7 előre definiált típus):
  - EcoFamily napok
  - Plakátos akció
  - Extra akció
  - Kuponos akció
  - Akciós újság
  - Online újság
  - Egyéb akció
- Kezdési és végdátum
- Publikálási dátum
- Felelős személy kijelölése
- Leírás szerkesztése (Rich Text Editor)
- Mechanizmus/működési módszer szerkesztése
- Bolt kiválasztása (3 mód)
- Termékek kiválasztása
- Képek beágyazása
- Kapcsolódó dokumentumok feltöltése
- Piszkozatként vagy közzétéveként mentés
- Csak szerkesztő csoport tagjai hozhatnak létre akciót
- Beágyazott képek mérete korlátozva (max 300KB/kép, max 900KB összesen base64-ben)

**Megvalósítás:** AkcioPanel.tsx, pnpListServices.ts

## 5. Akciók Szerkesztése és Megtekintése

**Funkció:** Akció megtekintése és szerkesztése módváltás

- Akció szerkesztése (szerkesztőknek)
- Akció megtekintése (csak olvasható mód)
- Módváltás (szerkesztő ↔ megtekintő)
- Összes mező szerkesztése
- Automatikus adat töltés a panelbe
- Nem szerkesztőknek csak a megtekintő mód érhető el

**Megvalósítás:** AkcioPanel.tsx, pnpListServices.ts

## 6. Akciók Közzététele

**Funkció:** Akció státuszváltás piszkozatból közzétett állapotba

- Közzététel - státusz: Piszkozat → Közzétéve
- Közzétételkor megerősítő dialógus
- Publikálási dátum megjegyzés
- Csak szerkesztőknek elérhető
- Az már közzétett akciók nem módosíthatók vissza piszkozat állapotba
- Státusz frissítés több lehetséges mezőnevet próbál ("Status", "Státusz", OData__Status, stb.)

**Megvalósítás:** AkcioPanel.tsx, pnpListServices.ts

## 7. Bolt Kiválasztás Funkcionalitás

**Funkció:** Akciók által érintett boltok kiválasztása

- **3 mód támogatása:**
  - "Összes bolt" - az összes bolt érintett
  - "Csoportosított boltok (K1-K7)" - csoportok kiválasztása
  - "Kiválasztott boltok" - egyedi boltok választása
- Keresés a kiválasztott boltok között
- Bolttípusok dinamikus listájából töltődnek
- Többszörös kiválasztás
- Maximum 5000 bolt lehet betölteni
- Kiválasztott boltok vesszővel elválasztott szöveggé konvertálódnak

**Megvalósítás:** TaskShopsSection.tsx, pnpListServices.ts

## 8. Termékkereső Modul

**Funkció:** EAN-számok alapján termékek keresése

- **Keresés típusok:**
  - EAN számok (8, 12, 13, 14 számjegy)
  - Cikkszámok
- **Bemenet módok:**
  - Kézi EAN szám bevitel (vesszővel, szóközzel, újsorral elválasztva)
  - Táblázatból való automatikus detektálás
- Keresési eredmények táblázatos megjelenítése
- Oszlopok: Cikkszám, EAN, Megnevezés, Ár, Akciós ár
- Eredmények másolása vágólapra (HTML + TEXT formátum)
- Termékadatok szinkronizálása a fő formba
- Maximálisan 5000 termék lehet egy lekérdezésben
- Automatikus detektálás csak táblázatokból működik
- Hibás EAN keresés fallback módban lineáris keresést végez

**Megvalósítás:** TiptapEditor.tsx, pnpListServices.ts

## 9. Rich Text Editor (Tiptap)

**Funkció:** Fejlett szöveg és tartalom szerkesztése

- **Szövegformázás:**
  - Félkövér, dőlt, aláhúzott, kódblokk
  - Címsorok (H1-H6)
  - Listák (számozás, felsorolás)
  - Idézetek
- **Tartalom típusok:**
  - Képek beágyazása (automatikus szélességmérés és 800px korlát)
  - Táblázatok (szerkeszthető, méretezhetó cellák)
  - Vízszintes elválasztó
- **Formatálás:**
  - Szövegigazítás (bal, közép, jobb)
  - Automatikus/manuális képméret szabályozás
- Beágyazott kép méret figyelmeztetés
- Termékkeresés integráció
- Base64-be kódolt képek maximum 300KB/kép méretkorlátozottsága
- Az összes beágyazott kép maximum 900KB-ot vehet fel
- Beágyazott képek nagyobb tárolási terheletet jelentenek

**Megvalósítás:** TiptapEditor.tsx

## 10. Kapcsolódó Dokumentumok Kezelése

**Funkció:** Akció-specifikus fájlok feltöltése és kezelése

- Fájl feltöltése filekickerrel
- Képek és dokumentumok külön megjelenítése
- Automatikus könyvtár struktura: `/[ActionId]/[FileName]`
- Fájlméret megjelenítése (B, KB, MB)
- Fájlletöltés
- Fájlok listázása és rendezése
- Töltési állapot jelzés
- **Támogatott képfájltípusok:** PNG, JPG, JPEG, GIF, WebP, SVG, BMP, TIFF, AVIF
- **Támogatott dokumentumtípusok:** PDF, DOC, DOCX, XLS, XLSX, PPT, PPTX, TXT, CSV, ZIP
- Dokumentumkönyvtár beállítása kötelező a webpart tulajdonságaiban
- Új akciónál (ID nélkül) nem lehet fájlokat feltölteni

**Megvalósítás:** RelatedFilesSection.tsx, pnpListServices.ts

## 11. Szerkesztői Jogosultság Kezelés

**Funkció:** Szerepalapú hozzáférés-vezérlés

- Szerkesztő csoport beállítása a webpart tulajdonságaiban
- Microsoft Graph API segítségével csoporthozzárendelés ellenőrzése
- Szerkesztőknek: teljes szerkesztési lehetőség
- Nem szerkesztőknek: csak megtekintés
- Piszkozatok rejtése nem szerkesztőktől
- Dinamikus jogosultság ellenőrzés
- Csak biztonsági csoportok (nem terjesztési listák) támogatottak
- Jogosultság ellenőrzése szinkron végrehajtódik az oldal betöltésénél

**Megvalósítás:** AkciokWebPart.ts, GlobalStore.ts, pnpListServices.ts

## 12. HTML Biztonsági Szanitizálás

**Funkció:** XSS-támadások megelőzése

- Script tagok eltávolítása
- Event handlerek (onclick, onload, stb.) eltávolítása
- javascript: protokoll linkek eltávolítása
- Style és iframe tagok eltávolítása
- Biztonságos megjelenítés
- Beágyazott kép ellenőrzés
- Az erőforrás-intenzív nagy HTML dokumentumok feldolgozása lassú lehet

**Megvalósítás:** AkcioPanel.tsx

## 13. Globális Állapotkezelés (Zustand)

**Funkció:** Centralizált state management

- WebPart kontextus
- SharePoint lista azonosítók (akciók, boltok, termékek, dokumentumok)
- Szerkesztő csoport információk
- Panel nyitottság és módja
- Kiválasztott akció ID
- Szűrő és rendezési beállítások
- Akció típusok gyűjteménye
- Bolt adatok cache
- Az állapot az oldal frissítésénél elvész
- Nem perzisztens az oldal bezárása után

**Megvalósítás:** GlobalStore.ts

## 14. WebPart Konfigurációs Lehetőségek

**Funkció:** SharePoint Property Pane beállítások

- **Akciók lista** (kötelező) - a SharePoint lista ID-ja
- **Boltok lista** (kötelező) - a bolt adatok forrása
- **Termékek lista** (opcionális) - az EAN-kereséshez
- **Kapcsolódó dokumentumok könyvtára** (opcionális) - fájlok tárolási helye
- **Feladatok oldala** (opcionális) - a feladatokra mutató link
- **Szerkesztő csoport** (opcionális) - szerkesztéshez jogosult csoport
- Az akciók lista és a boltok lista kiválasztása kötelező
- Property Pane értékek csak az oldal frissítése után lépnek életbe

**Megvalósítás:** AkciokWebPart.ts

## 15. Adatmodellek és Tipusdefiníciók

**Funkció:** TypeScript interfészek és típusok

- `IAkcio` - Az akció teljes adatstruktúrája
- `IAkcioListItem` - A listanézethez szükséges minimális adatok
- `IAkcioDetails` - A panel nézet teljes adatai
- `IShopListItem` - Bolt adatok
- `IRelatedFileItem` - Dokumentum adatok

**Megvalósítás:** IAkcio.ts, IShop.ts, IRelatedFile.ts

## 16. UI/UX Komponensek

**Funkció:** Fluent UI alapú felhasználói felület

- Panel, Dropdown, DatePicker, MessageBar komponensek
- Betöltési és hibajelzések minden fontosabb user flow-ban
- Validáció és felhasználói visszajelzés
- Intuitív navigáció és módváltás

**Megvalósítás:** Összes komponens (AkciokList.tsx, AkcioPanel.tsx, stb.)


