# Ania NSP - wersja przeglądarkowa

Ten folder zawiera samodzielną statyczną aplikację przeglądarkową. Możesz wgrać cały folder do repozytorium GitHub i hostować go przez GitHub Pages.

## Co robi aplikacja

- wczytuje lokalne dane granic gmin
- wczytuje lokalne dane granic powiatów
- wczytuje lokalne dane przepływów z `XLSX`, `CSV` lub `TXT`
- pozwala wybrać zakres gminy albo powiatu
- stosuje filtry przepływu, odległości i Top N
- eksportuje paczkę ZIP do pobrania zawierającą:
  - liniowy `GeoJSON`
  - poligonowy `GeoJSON`
  - podsumowanie `TXT`
  - wieloarkuszowy `XLSX`

## Ważne ograniczenie tej pierwszej wersji

Ta wersja działa całkowicie bez zależności zewnętrznych i wyłącznie w przeglądarce, dlatego na razie **nie** odczytuje ani nie zapisuje `GPKG`.

Zamiast tego oczekuje:

- granic gmin w formacie `GeoJSON`
- granic powiatów w formacie `GeoJSON`
- danych przepływów w formacie `XLSX`, `CSV` lub `TXT`

To uproszczenie pozwala zachować aplikację:

- hostowalną na GitHub Pages
- w pełni zamkniętą w jednym folderze
- bez CDN, kroków budowania i instalacji

## Oczekiwane pola wejściowe

### GeoJSON gmin

Właściwości obiektów powinny zawierać:

- `JPT_KOD_JE`
- `JPT_NAZWA_`

Obsługiwane są także zapasowe nazwy pól:

- `kod_teryt`
- `code`
- `name`
- `nazwa_gminy`

### GeoJSON powiatów

Właściwości obiektów powinny zawierać:

- `JPT_KOD_JE`
- `JPT_NAZWA_`

### Skoroszyt przepływów

Aplikacja używa pierwszego pasującego arkusza o nazwie zbliżonej do `Macierz przeplywow`, a jeśli nie znajdzie podobnego, bierze pierwszy arkusz.

Pierwsze pięć kolumn jest interpretowanych jako:

1. kod zamieszkania
2. nazwa zamieszkania
3. kod pracy
4. nazwa pracy
5. przepływy

## Hosting na GitHub Pages

1. Umieść ten folder w swoim repozytorium.
2. Upewnij się, że `index.html` znajduje się w publikowanym katalogu głównym albo podfolderze publikacji.
3. Włącz GitHub Pages w ustawieniach repozytorium.

## Następny krok rozwoju

Jeżeli chcesz uzyskać pełniejszą zgodność z wersją desktopową, kolejnym krokiem będzie dołączenie do tego folderu przeglądarkowego środowiska SQLite/WASM, aby aplikacja mogła również wczytywać i generować prawdziwe pliki `GPKG`.
