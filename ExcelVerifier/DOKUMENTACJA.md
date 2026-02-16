# ExcelVerifier - Dokumentacja

## 📋 Spis Treści

1. [Opis ogólny](#1-opis-ogólny)
2. [Przetwarzane dane](#2-przetwarzane-dane)
3. [Moduły aplikacji](#3-moduły-aplikacji)
4. [Instalacja](#4-instalacja)
5. [Konfiguracja](#5-konfiguracja)
6. [Instrukcja użytkowania](#6-instrukcja-użytkowania)
7. [Struktura projektu](#7-struktura-projektu)
8. [Architektura techniczna](#8-architektura-techniczna)
9. [Walidacja danych](#9-walidacja-danych)
10. [FAQ i Rozwiązywanie problemów](#10-faq-i-rozwiązywanie-problemów)

---

## 1. Opis ogólny

### 🎯 Cel aplikacji

**ExcelVerifier** to desktopowa aplikacja do automatycznego przetwarzania zeskanowanych dokumentów magazynowych i logistycznych z wykorzystaniem sztucznej inteligencji. Aplikacja konwertuje obrazy papierowych dokumentów (faktury dostaw, protokoły zwrotów, dokumenty wymiany towarów) na pliki Excel z zachowaniem struktury tabelarycznej i automatyczną walidacją poprawności stanów magazynowych.

### 💼 Problemy biznesowe, które rozwiązuje

1. **Digitalizacja dokumentów papierowych** - Eliminuje konieczność ręcznego przepisywania danych ze zeskanowanych dokumentów magazynowych
2. **Oszczędność czasu** - Przetwarzanie dokumentu w 10-30 sekund zamiast 5-10 minut ręcznej pracy
3. **Redukcja błędów** - AI Google Gemini ekstraktuje dane z 95%+ dokładnością
4. **Automatyczna walidacja** - Weryfikacja poprawności obliczeń stanów magazynowych
5. **Ustrukturyzowane dane** - Gotowe raporty Excel do analizy i arhiwizacji
6. **Zarządzanie historią** - Automatyczna klasyfikacja po kontrahentach i datach

### 👥 Grupy użytkowników

- **Magazynierzy** - Pracownicy przyjmujący dostawy i zarządzający stanami
- **Logistyka** - Osoby odpowiedzialne za zwroty i wymianę towarów
- **Księgowość** - Weryfikacja dokumentów dostaw i zwrotów
- **Kontrola jakości** - Audyt poprawności stanów magazynowych
- **Menedżerowie** - Dostęp do raportów i statystyk

### 🔄 Krótki opis działania

Aplikacja implementuje trzystopniowy proces przetwarzania dokumentów:

**Transform → Verify → Generate Report**

Każdy dokument przechodzi przez pełny cykl: od skanowania, przez ekstrakcję za pomocą AI, weryfikację użytkownika, aż po zatwierdzenie i włączenie do raportów zbiorczych.

---

## 2. Przetwarzane dane

### 📄 Typ dokumentów

Aplikacja jest dedykowana do przetwarzania **dokumentów magazynowych/logistycznych**:
- Faktury dostaw towarów
- Protokoły zwrotów produktów
- Dokumenty wymiany/rotacji towarów
- Zestawienia stanów magazynowych
- Wszystkie dokumenty zawierające tabelę produktów ze stanami

### 🗂️ Ekstraktowane informacje

#### Metadane dokumentu (wiersz 1 w Excel)

| Pole | Opis | Przykład |
|------|------|----------|
| **Odbiorca** | Pełna nazwa firmy i adres | "ABCDE SP. Z O.O. KORCZOWA UL. KWIATEK 6" |
| **Nr dokumentu** | Numer faktury/protokołu | "FV/2026/01/001" |
| **Data wystawienia** | Data dokumentu (DD.MM.YYYY) | "19.01.2026" |

#### Tabela danych produktowych (od wiersza 3)

| Kolumna | Nazwa | Typ danych | Opis |
|---------|-------|------------|------|
| **A** | Lp | Liczba | Liczba porządkowa pozycji |
| **B** | Nazwa | Tekst | Nazwa produktu/towaru |
| **C** | Ilość | Liczba | Ilość zamówiona/dostarczona |
| **D** | Uwagi | Tekst | Dodatkowe uwagi, komentarze |
| **E** | Ilość | Liczba | Ilość zwrócona |
| **F** | Stan poprzedni | Liczba | Stan magazynowy przed operacją |
| **G** | Stan po wymianie | Liczba | Stan magazynowy po operacji |

### 🎯 Obsługiwane formaty

- **Wejście**: JPG, PNG, JPEG (skanowane obrazy dokumentów)
- **Wyjście**: XLSX (Excel), strukturalne pliki tekstowe

---

## 3. Moduły aplikacji

### 📸 Moduł 1: Transform - Transformacja obrazu

**Ścieżka:** `ui/TransformPicToExcelPage.py`

#### Funkcjonalność

1. **Ładowanie obrazów**
   - Przeciągnij i upuść (drag & drop)
   - Wybór z eksploratora plików
   - Obsługa wielu plików naraz

2. **Przetwarzanie wstępne**
   - Automatyczne wykrywanie krawędzi tabel
   - Wizualne podświetlenie strukture tabelarycznej
   - Usuwanie białych marginesów (trim whitespace)
   - Optymalizacja obrazu dla AI

3. **Operacje na obrazie**
   - Obrót w lewo (90°)
   - Obrót w prawo (90°)
   - Podgląd przed wysłaniem

4. **Ekstrakcja danych przez AI**
   - Wysyłanie do Google Gemini API
   - Połączone zapytanie (4 informacje w 1 wywołaniu = 4x szybciej)
   - Automatyczne fallbacki między modelami AI
   - Retry logic przy błędach 503

5. **Zapis wyników**
   - Generowanie pliku Excel z danymi
   - Automatyczna klasyfikacja do folderu kontrahenta
   - Kopiowanie obrazu źródłowego do folderu

#### Obsługiwane modele AI (w kolejności fallback)

1. `gemini-3-flash-preview` (domyślny - najszybszy)
2. `gemini-2.5-flash`
3. `gemini-2.5-pro` (najdokładniejszy, ale wolniejszy)

#### Algorytm ekstrakcji

```python
# Pojedyncze zapytanie do AI ekstraktuje wszystkie dane:
1. "ODBIORCA: [tekst]"          → Wiersz 1, Kolumna B
2. "Nr dokumentu: [tekst]"       → Wiersz 1, Kolumna D
3. "Data wystawienia: [tekst]"   → Wiersz 1, Kolumna F
4. Tabela w formacie pipe:       → Od wiersza 3
   |Lp|Nazwa|Ilość|Uwagi|Ilość|Stan poprzedni|Stan po wymianie|
```

### ✅ Moduł 2: Verify - Weryfikacja i edycja

**Ścieżka:** `ui/VerificationPage.py`

#### Funkcjonalność

1. **Lista dokumentów niezatwierdzonych**
   - Wyświetlanie wszystkich przetworzonych dokumentów
   - Status: Niezatwierdzone / Zatwierdzone
   - Filtrowanie i sortowanie

2. **Podgląd równoległy**
   - Oryginalny obraz dokumentu (lewa strona)
   - Wyekstraktowane dane w tabeli (prawa strona)
   - Synchroniczny podgląd dla łatwej weryfikacji

3. **Edycja danych**
   - Edycja bezpośrednio w tabeli (double-click)
   - Zachowanie typów danych (liczby, tekst)
   - Automatyczny zapis po zakończeniu edycji

4. **Walidacja matematyczna**
   - Automatyczne sprawdzanie zgodności stanów
   - Kolorowanie błędnych wierszy na czerwono
   - Formuła: `Stan po = Stan poprzedni + Dostawa - Zwrot`

5. **Akcje na dokumentach**
   - **Zatwierdź** - Przenosi do zatwierdzonych + aktualizuje raporty
   - **Usuń** - Trwałe usunięcie dokumentu i obrazu
   - **Przetwórz ponownie** - Ponowna ekstrakcja przez AI
   - **Zapisz** - Zapis zmian bez zatwierdzania

6. **Obsługa zatwierdzonych dokumentów**
   - Przeglądanie historii zatwierdzonych
   - Możliwość cofnięcia zatwierdzenia
   - Ponowna edycja i aktualizacja

#### Walidacja w czasie rzeczywistym

Po każdej zmianie danych:
1. Obliczany jest oczekiwany stan końcowy
2. Porównanie z rzeczywistym stanem
3. Podświetlenie na czerwono jeśli niezgodność
4. Przywrócenie oryginalnego formatu jeśli poprawne

### 📊 Moduł 3: Generate Report - Generowanie raportów

**Ścieżka:** `ui/GenerateReportPage.py`

#### Funkcjonalność

1. **Filtry raportów**
   - Wszystkie miesiące
   - Konkretny miesiąc
   - Zakres dat

2. **Generowanie raportu zbiorczego**
   - Agregacja wszystkich zatwierdzonych dokumentów
   - Zbiorczy plik Excel ze wszystkimi pozycjami
   - Sortowanie chronologiczne

3. **Zawartość raportu**
   - Wszystkie produkty ze wszystkich dokumentów
   - Pełna historia stanów magazynowych
   - Kolumny: Data | Odbiorca | Nazwa | Ilość zamówiona | Ilość zwrócona | Stan poprzedni | Stan po wymianie

4. **Eksport**
   - Plik zapisywany w głównym folderze `Reports/`
   - Nazwa: `Report_[miesiąc]_[rok].xlsx` lub `Report_All.xlsx`
   - Gotowy do importu do systemów ERP

---

## 4. Instalacja

### Wymagania systemowe

- **System operacyjny**: Windows 10/11, macOS, Linux
- **Python**: 3.11 lub nowszy
- **RAM**: Minimum 4 GB (zalecane 8 GB)
- **Miejsce na dysku**: 500 MB + miejsce na dokumenty
- **Połączenie internetowe**: Wymagane tylko przy ekstrakcji danych (API)

### Instalacja krok po kroku

#### 1. Zainstaluj Python 3.11+

Pobierz ze strony: https://www.python.org/downloads/

```bash
python --version  # Sprawdź wersję (powinno być 3.11+)
```

#### 2. Sklonuj/pobierz projekt

```bash
# Jeśli masz git:
git clone <repository-url>
cd ExcelVerifier

# Lub rozpakuj archiwum ZIP
```

#### 3. Utwórz środowisko wirtualne

```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# macOS/Linux
python3 -m venv .venv
source .venv/bin/activate
```

#### 4. Zainstaluj zależności

```bash
pip install -r ExcelVerifier/requirements.txt
```

**Główne zależności:**
```
PyQt5>=5.15.0
Pillow>=10.0.0
pandas>=2.0.0
openpyxl>=3.1.0
google-generativeai>=0.3.0
```

---

## 5. Konfiguracja

### 🔑 Klucz API Google Gemini

#### Uzyskanie klucza API

1. Przejdź do: https://makersuite.google.com/app/apikey
2. Zaloguj się kontem Google
3. Kliknij "Create API Key"
4. Skopiuj wygenerowany klucz

#### Konfiguracja klucza

**Metoda 1: Zmienna środowiskowa (zalecana)**

```bash
# Windows (PowerShell)
$env:GEMINI_API_KEY="twój-klucz-api"

# Windows (CMD)
set GEMINI_API_KEY=twój-klucz-api

# macOS/Linux
export GEMINI_API_KEY="twój-klucz-api"
```

**Metoda 2: Plik config.py**

Edytuj `ExcelVerifier/config.py`:

```python
GEMINI_API_KEY = "twój-klucz-api-tutaj"
```

### ⚙️ Ustawienia aplikacji

Plik: `settings.json` (tworzony automatycznie przy pierwszym uruchomieniu)

```json
{
  "api_key": "twój-klucz-api",
  "default_model": "gemini-3-flash-preview",
  "auto_trim_images": true,
  "reports_folder": "Reports"
}
```

---

## 6. Instrukcja użytkowania

### 🚀 Uruchomienie aplikacji

```bash
# Aktywuj środowisko wirtualne (jeśli nie aktywowane)
.venv\Scripts\activate  # Windows
source .venv/bin/activate  # macOS/Linux

# Uruchom aplikację
cd ExcelVerifier
python main.py
```

### 📖 Scenariusz użycia krok po kroku

#### Krok 1: Transformacja dokumentu

1. Otwórz aplikację i przejdź do zakładki **Transform**
2. Przeciągnij obraz dokumentu lub kliknij "Load Images"
3. Sprawdź podgląd - czy dokument jest prawidłowo zorientowany
4. Jeśli trzeba, użyj przycisków rotacji (⟲ ⟳)
5. Kliknij **"Transform to Excel"**
6. Poczekaj 10-30 sekund na przetworzenie
7. Dokument pojawi się na liście niezatwierdzonych

#### Krok 2: Weryfikacja danych

1. Przejdź do zakładki **Verify**
2. Wybierz dokument z listy niezatwierdzonych
3. Sprawdź wyekstraktowane dane:
   - **Lewa strona**: Oryginalny obraz
   - **Prawa strona**: Tabela z danymi
4. Zwróć uwagę na czerwone wiersze (błędy walidacji)
5. Edytuj błędne dane (double-click na komórce)
6. Kliknij **"Zapisz"** aby zapisać zmiany
7. Jeśli wszystko OK, kliknij **"Zatwierdź"**

#### Krok 3: Generowanie raportu

1. Przejdź do zakładki **Generate Report**
2. Wybierz miesiąc lub "Wszystkie miesiące"
3. Kliknij **"Generuj Raport"**
4. Raport zostanie zapisany w folderze `Reports/`
5. Otwórz plik Excel i sprawdź wyniki

---

## 7. Struktura projektu

```
ExcelVerifier/
│
├── ExcelVerifier/                    # Główny katalog aplikacji
│   ├── main.py                       # Punkt wejścia aplikacji
│   ├── config.py                     # Konfiguracja globalna
│   ├── requirements.txt              # Zależności Python
│   │
│   ├── core/                         # Logika biznesowa
│   │   ├── __init__.py
│   │   ├── excel_handler.py          # Obsługa plików Excel
│   │   ├── file_manager.py           # Zarządzanie plikami
│   │   └── image_transformer.py      # Ekstrakcja danych z obrazów (AI)
│   │
│   ├── ui/                           # Interfejs użytkownika (PyQt5)
│   │   ├── __init__.py
│   │   ├── main_window.py            # Główne okno aplikacji
│   │   ├── TransformPicToExcelPage.py    # Moduł Transform
│   │   ├── VerificationPage.py       # Moduł Verify
│   │   ├── GenerateReportPage.py     # Moduł Generate Report
│   │   ├── dialogs.py                # Dialogi pomocnicze
│   │   ├── settings_dialog.py        # Okno ustawień
│   │   ├── styles.py                 # Style CSS/QSS
│   │   └── utils.py                  # Funkcje pomocnicze UI
│   │
│   └── Reports/                      # Folder z danymi (generowany)
│       ├── [Nazwa Firmy]/            # Foldery per kontrahent
│       │   ├── 2026-01-15_Firma.xlsx
│       │   ├── 2026-01-15_Firma.jpg
│       │   └── ...
│       ├── Niezatwierdzone/          # Dokumenty do weryfikacji
│       └── Zatwierdzone/             # Zatwierdzone dokumenty
│
├── DOKUMENTACJA.md                   # Ten plik
├── build.py                          # Skrypt budowania exe
├── ExcelVerifier.spec                # Konfiguracja PyInstaller
├── settings.json                     # Ustawienia użytkownika
└── *.py                             # Skrypty pomocnicze (patch_*.py)
```

---

## 8. Architektura techniczna

### 🏗️ Wzorce projektowe

- **MVC (Model-View-Controller)**: Separacja logiki biznesowej (core) od UI
- **Factory Pattern**: Tworzenie instancji obiektów AI
- **Strategy Pattern**: Wybór modelu AI (fallback mechanism)
- **Observer Pattern**: Aktualizacja UI po zmianach danych

### 🔌 Główne komponenty

#### ImageTransformer (`core/image_transformer.py`)

**Odpowiedzialności:**
- Komunikacja z Google Gemini API
- Przetwarzanie obrazów przed wysłaniem
- Parsowanie odpowiedzi AI
- Zarządzanie fallbackami między modelami
- Retry logic przy błędach

**Kluczowe metody:**
```python
def query_gemini_combined(image_path, model) -> dict
    # Pojedyncze zapytanie ekstraktujące wszystkie dane

def process_image_file(image_path, base_folder) -> str
    # Pełny pipeline: obraz → AI → Excel

def parse_date_flexible(date_text) -> datetime
    # Parsowanie dat w różnych formatach
```

#### ExcelHandler (`core/excel_handler.py`)

**Odpowiedzialności:**
- Odczyt i zapis plików Excel
- Walidacja matematyczna stanów
- Zarządzanie formatowaniem (kolory, czcionki)
- Generowanie raportów zbiorczych
- Synchronizacja z plikami metadanych

**Kluczowe metody:**
```python
def load_file(file_path) -> DataFrame
    # Ładowanie Excel do DataFrame

def save_data(ui_table_data)
    # Zapis + walidacja + kolorowanie

def _apply_validation_coloring(worksheet)
    # Czerwone podświetlenie błędów

def approve_report(filename, date, company, path)
    # Zatwierdzenie dokumentu

def generate_report(filters) -> str
    # Generowanie raportu zbiorczego
```

#### FileManager (`core/file_manager.py`)

**Odpowiedzialności:**
- Organizacja struktury folderów
- Przenoszenie plików między statusami
- Tworzenie backupów
- Czyszczenie starych plików

### 🔄 Przepływ danych

```
┌─────────────┐
│   Obraz     │
│   (JPG/PNG) │
└──────┬──────┘
       │
       ▼
┌─────────────────┐
│  ImageTransform │──────┐
│  - trim_whitespace     │
│  - detect_edges        │
│  - optimize           │
└─────────┬──────────────┘
          │
          ▼
┌──────────────────┐
│  Gemini API      │
│  - gemini-3-flash│
│  - fallback models│
└─────────┬────────┘
          │
          ▼
┌──────────────────┐
│  Parse Response  │
│  - Odbiorca      │
│  - Nr dokumentu  │
│  - Data          │
│  - Tabela        │
└─────────┬────────┘
          │
          ▼
┌──────────────────┐
│  ExcelHandler    │
│  - Create XLSX   │
│  - Format cells  │
│  - Save to folder│
└─────────┬────────┘
          │
          ▼
┌──────────────────┐
│  UI Update       │
│  - Show in Verify│
│  - Enable editing│
└──────────────────┘
```

---

## 9. Walidacja danych

### 📐 Formuła walidacji stanów magazynowych

Dla każdego wiersza produktu (od wiersza 4 wzwyż):

```
Stan po wymianie (G) = Stan poprzedni (F) + Ilość dostarczona (C) - Ilość zwrócona (E)
```

### 🎨 Kolorowanie walidacyjne

| Kolor | Znaczenie | Warunek |
|-------|-----------|---------|
| 🔴 **Czerwony** (`#FF0000`) | Błąd obliczeń | `Stan po != (Stan poprz + Dostawa - Zwrot)` |
| ⚪ **Biały/Oryginalny** | Poprawne | `Stan po == (Stan poprz + Dostawa - Zwrot)` |

### 🔍 Przypadki szczególne

#### Przypadek 1: Brak dostaw i zwrotów
```python
if C == None and E == None:
    Expected = F  # Stan bez zmian
```

#### Przypadek 2: Tylko dostawa
```python
if C != None and E == None:
    Expected = F + C
```

#### Przypadek 3: Tylko zwrot
```python
if C == None and E != None:
    Expected = F - E
```

#### Przypadek 4: Dostawa i zwrot
```python
if C != None and E != None:
    Expected = F + C - E
```

### ⚙️ Implementacja techniczna

```python
def _apply_validation_coloring(self, ws):
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    
    for row_idx in range(4, ws.max_row + 1):
        c = self._to_num(ws.cell(row=row_idx, column=3).value)  # Dostawa
        e = self._to_num(ws.cell(row=row_idx, column=5).value)  # Zwrot
        f = self._to_num(ws.cell(row=row_idx, column=6).value)  # Stan poprz
        g = self._to_num(ws.cell(row=row_idx, column=7).value)  # Stan po
        
        safe_f = f if f is not None else 0.0
        expected = None
        
        # Logika obliczania expected (jak wyżej)
        
        # Porównanie z zaokrągleniem do 2 miejsc po przecinku
        is_error = False
        if expected is not None and g is not None:
            if round(float(g), 2) != round(float(expected), 2):
                is_error = True
        
        # Aplikacja koloru
        if is_error:
            ws.cell(row=row_idx, column=7).fill = red_fill
        else:
            # Restore original fill if it wasn't red
            ws.cell(row=row_idx, column=7).fill = original_fill
```

---

## 10. FAQ i Rozwiązywanie problemów

### ❓ Najczęstsze pytania

#### Q: Jak długo trwa przetworzenie jednego dokumentu?
**A:** Średnio 10-30 sekund, w zależności od:
- Rozmiaru obrazu (większe = wolniejsze)
- Złożoności tabeli
- Obciążenia API Google
- Modelu AI (gemini-3-flash jest najszybszy)

#### Q: Czy aplikacja działa offline?
**A:** Nie całkowicie. Wymaga połączenia internetowego tylko podczas ekstrakcji danych (moduł Transform). Moduły Verify i Generate Report działają offline.

#### Q: Ile kosztuje użycie Google Gemini API?
**A:** Google oferuje darmowy tier:
- 15 zapytań/minutę
- 1500 zapytań/dzień
- Całkowicie darmowe dla użytku standardowego

Więcej: https://ai.google.dev/pricing

#### Q: Czy dane są wysyłane do zewnętrznych serwerów?
**A:** Tak, ale tylko obrazy podczas ekstrakcji. Wszystkie pliki Excel i dane są przechowywane lokalnie.

#### Q: Czy mogę edytować zatwierdzone dokumenty?
**A:** Tak, w module Verify możesz otworzyć "Zatwierdzone dokumenty" i wprowadzić zmiany. System automatycznie zaktualizuje raporty.

#### Q: Co się stanie jeśli AI źle rozpozna dane?
**A:** Możesz:
1. Edytować dane ręcznie w module Verify
2. Kliknąć "Przetwórz ponownie" aby wysłać obraz ponownie do AI
3. Usunąć dokument i przetwarzać go od nowa

### 🐛 Rozwiązywanie problemów

#### Problem: "GEMINI_API_KEY not found"

**Rozwiązanie:**
```bash
# Ustaw zmienną środowiskową
export GEMINI_API_KEY="twój-klucz"

# Lub dodaj do config.py
GEMINI_API_KEY = "twój-klucz"
```

#### Problem: "503 Service Unavailable"

**Przyczyna:** API Google przeciążone lub niedostępne

**Rozwiązanie:**
- Aplikacja automatycznie spróbuje ponownie (5 prób z wykładniczym opóźnieniem)
- Następnie spróbuje innych modeli (fallback)
- Jeśli wszystkie modele zawiodą, dokument trafi do "Niezatwierdzonych" z komunikatem błędu

#### Problem: "Permission denied" przy zapisie

**Przyczyna:** Plik Excel jest otwarty w innej aplikacji

**Rozwiązanie:**
- Zamknij plik w programie Excel
- Spróbuj ponownie zapisać

#### Problem: Czerwone wiersze mimo poprawnych danych

**Przyczyna:** Zaokrąglenia lub formaty liczb

**Rozwiązanie:**
- Sprawdź czy wartości są liczbami (nie tekstem)
- System zaokrągla do 2 miejsc po przecinku
- Ręcznie popraw wartości jeśli potrzeba

#### Problem: AI nie rozpoznaje struktury tabeli

**Rozwiązanie:**
1. Upewnij się, że obraz jest wyraźny i dobrze oświetlony
2. Użyj przycisku "Pretreatment" przed transformacją
3. Obrób dokument jeśli jest przekrzywiony
4. Zwiększ rozdzielczość skanu (min. 300 DPI)

#### Problem: Aplikacja się nie uruchamia

**Rozwiązanie:**
```bash
# Sprawdź wersję Pythona
python --version  # Powinno być 3.11+

# Reinstaluj zależności
pip install --force-reinstall -r requirements.txt

# Sprawdź błędy
python ExcelVerifier/main.py
```

### 📞 Wsparcie techniczne

Jeśli problem nie został rozwiązany:

1. Sprawdź logi w terminalu
2. Zrób zrzut ekranu błędu
3. Przygotuj przykładowy obraz dokumentu (jeśli problem dotyczy ekstrakcji)
4. Skontaktuj się z zespołem deweloperskim

---

## 📝 Changelog

### Wersja aktualna
- ✅ Połączone zapytania AI (4x szybciej)
- ✅ Automatyczna walidacja matematyczna
- ✅ Dynamiczna reorganizacja plików po zmianie odbiorcy
- ✅ Obsługa wielu modeli AI z fallbackami
- ✅ Retry logic z wykładniczym opóźnieniem
- ✅ Trim whitespace dla lepszej ekstrakcji
- ✅ Wykrywanie i podświetlanie krawędzi tabel

### Planowane funkcje
- 🔜 Obsługa PDF
- 🔜 Batch processing (wiele plików naraz)
- 🔜 Export do CSV/JSON
- 🔜 Integracja z systemami ERP
- 🔜 Statystyki i wykresy
- 🔜 Wyszukiwanie pełnotekstowe

---

## 📄 Licencja

Projekt ExcelVerifier jest własnością...
(Dodaj informacje o licencji)

---

## 👨‍💻 Autorzy

Aplikacja rozwijana przez...
(Dodaj informacje o autorach)

---

**Ostatnia aktualizacja dokumentacji:** 8 lutego 2026
