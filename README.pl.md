# Data Handover — wtyczka QGIS

[English](README.md) · **Polski**

Wtyczka QGIS, która eksportuje tabelę atrybutów warstwy wektorowej do
sformatowanego pliku Excela — z wyglądem ustawionym przez Ciebie, czytelnymi
nagłówkami i etykietami zamiast kodów.

## Co ustawiasz

| Zakładka | Co zawiera |
|---|---|
| **Dane** | warstwa, zakres (wszystkie / zaznaczone / w zasięgu mapy), kolumny i ich kolejność, nagłówek, jednostka i opis kolumny (zapamiętywane w projekcie), decyzja o danych osobowych |
| **Wygląd** | tabela Excela albo zwykły zakres (z filtrami lub bez); osobno dla nagłówków i zawartości: czcionka, rozmiar, kolor, pogrubienie, kursywa, podkreślenie, tło, wyrównanie, zawijanie; pasy co drugi wiersz; obramowanie (linia, kolor, wszystkie linie / tylko poziome / tylko obrys); tytuł nad danymi, zamrożony nagłówek, siatka, opis w dymku nagłówka, szerokość kolumn, orientacja wydruku, kolor karty; gotowe kolorystyki i podgląd na żywo |
| **Arkusze** | tytuł, język pliku (10 języków, domyślnie angielski), arkusze dodatkowe — domyślnie wyłączone — z wyborem zawartości |
| **Zmiany** | porównanie z poprzednim eksportem według kolumny-klucza |

## Co jest w pliku

| Arkusz | Co zawiera |
|---|---|
| **Dane** | dane sformatowane według ustawień; nagłówki z aliasów; etykiety zamiast kodów (value map, relacje); liczby z jednostką, które dalej są liczbami; identyfikatory (TERYT, nr działki) jako tekst; gotowe do wydruku |
| **Słownik** (opcja) | wybrane pozycje dla każdej kolumny: pole źródłowe, typ, jednostka, opis, kody i etykiety |
| **Metryka** (opcja) | wybrane pozycje: stan danych na dzień, warstwa, źródło, układ współrzędnych, zakres, filtr, liczba obiektów, autor, instrukcja czytania, uwagi do wydania |
| **Zmiany** (opcja) | obiekty dodane, usunięte i zmienione względem poprzedniego pliku, z wartościami „było / jest”; zmienione komórki podświetlone w danych |

Pola wyglądające na dane osobowe (nazwa pola albo wartości: PESEL z sumą
kontrolną, e-mail, telefon, numer księgi wieczystej) są domyślnie
odznaczone; można je zamaskować albo świadomie wyeksportować. Wykrywanie to
podpowiedź, nie gwarancja anonimizacji.

Eksport działa w tle: z QGIS można dalej korzystać, a pasek komunikatów
pokazuje postęp z przyciskiem **Anuluj**. Anulowanie zostawia plik
docelowy bez zmian. Przy bardzo dużych warstwach słownikowych (pola Value
Relation, Relation Reference) QGIS może chwilę stać przed startem eksportu,
gdy wczytuje etykiety.

Przycisk **Przywróć domyślne** na dole okna zeruje wygląd, arkusze i język
pliku (opisy kolumn zapisane w projekcie zostają). Okno i plik są dostępne
po angielsku, niemiecku, francusku, hiszpańsku, portugalsku, polsku,
rosyjsku, koreańsku, japońsku i chińsku. Język okna wybierasz na liście
„Język wtyczki” w stopce (domyślnie angielski), język treści pliku —
osobno w zakładce Arkusze. Nagłówki i wartości kolumn pochodzą z warstwy,
więc zostają w języku danych.

## Wymagania i instalacja

- QGIS 3.40–3.44 albo 4.x (sprawdzone: 3.40, 3.44.15, 4.2.3).
- Biblioteka Pythona **openpyxl** — jest w instalatorze QGIS dla Windows;
  w OSGeo4W dodaj pakiet `python3-openpyxl`, na Linuksie zainstaluj
  `python3-openpyxl`. Bez niej wtyczka pokaże instrukcję instalacji.

Instalacja z pliku: `scripts/package.sh` tworzy
`build/data_handover-<wersja>.zip`; w QGIS: *Wtyczki → Zarządzaj wtyczkami →
Zainstaluj z pliku ZIP*. Wtyczka pojawia się na pasku narzędzi, w menu
*Wtyczki* i w menu kontekstowym warstwy wektorowej.

## Rozwój

```bash
uvx ruff check .
uv run --no-project --python 3.12 --with mypy --with openpyxl --with types-openpyxl mypy data_handover/core
uv run --no-project --python 3.12 --with pytest --with openpyxl pytest -q
scripts/qgis-test.sh 4.2.3          # test w prawdziwym QGIS (Docker); CI także 3.40 i 3.44.15
scripts/render-xlsx.sh build/przyklad_qgis4.2.3.xlsx   # zrzuty arkuszy (LibreOffice)
```

Licencja:
GPL-2.0-or-later ([`LICENSE`](LICENSE)).
