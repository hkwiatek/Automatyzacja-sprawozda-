Automatyzacja Raportowania Finansowego (Excel -> Word)
* O projekcie
Narzędzie stworzone w środowisku Python automatyzujące proces przenoszenia, formatowania i linkowania sprawozdań finansowych pomiędzy MS Excel a MS Word. Projekt został wdrożony w środowisku produkcyjnym, eliminując powtarzalną, manualną pracę i drastycznie redukując ryzyko błędów ludzkich podczas tworzenia raportów.
Ze względu na to, że projekt jest wdrożonym narzędziem komercyjnym, repozytorium zawiera opis projektu README.
* Technologie
Python 3.x
pywin32 (win32com.client): Do bezpośredniego sterowania aplikacjami MS Office w tle (w trybie ukrytym).   
re (RegEx): Do zaawansowanego wyszukiwania numerów not finansowych.   
* Kluczowe funkcjonalności
Sterowanie w oparciu o arkusz konfiguracyjny: Skrypt działa w oparciu o dedykowany arkusz "Konfiguracja" w Excelu, w którym użytkownik definiuje nazwy zakładek docelowych, arkusze źródłowe i zakresy komórek.   
Algorytm inteligentnego linkowania (Noty): Skrypt skanuje pierwsze 5 wierszy każdej tabeli w poszukiwaniu kolumn z notami. Wykorzystując wyrażenia regularne, ucina wielopoziomowe końcówki (np. zamienia "10.1.3" na główną notę "10") i automatycznie tworzy wewnętrzne hiperłącza (niebieskie, niepodkreślone) do odpowiednich zakładek w Wordzie.   
Odtwarzanie komentarzy audytorów: Przed usunięciem starej tabeli w Wordzie, narzędzie skanuje ją pod kątem komentarzy, zapisuje ich dokładne współrzędne (wiersz, kolumna), autora oraz treść, a następnie precyzyjnie nanosi je na nowo wklejoną, zaktualizowaną tabelę.   
Zarządzanie formatowaniem: Narzędzie rozpoznaje formatowanie z Excela. Jeśli tekst jest wyrównany do prawej (np. długie kwoty finansowe), skrypt kategorycznie blokuje zawijanie tekstu w Wordzie, wymuszając poszerzenie kolumny, aby liczby nie zostały rozbite na dwie linijki. Dodatkowo szerokość wklejanych tabel jest automatycznie skalowana do 99% szerokości okna, aby uniknąć błędów marginesów.   
* Wyzwania i rozwiązania
Zarządzanie systemowym schowkiem: Wykorzystanie wbudowanej funkcji kopiuj/wklej systemu Windows wymagało zabezpieczenia skryptu przed ingerencją użytkownika. Stworzono instrukcję ("INSTRUKCJA_OBSŁUGI.docx") oraz mechanizmy wstrzymujące (np. pętle ponawiające wklejanie), aby uniknąć "wstrzyknięcia" przypadkowego tekstu w środek sprawozdania.   
Ciągłość numeracji i nagłówków: Zastosowano logikę zapewniającą ciągłość numeracji sekcji (wyłączenie restartu numeracji) oraz automatyczne powtarzanie nagłówków tabel przechodzących na kolejne strony (z uwzględnieniem scalonych komórek).   
* Dokumentacja
Jako część wdrożenia, przygotowałem pełną instrukcję dla użytkowników końcowych, opisującą zasady działania, najczęstsze błędy (np. chronione arkusze, przeniesienie numeracji po usunięciu noty) oraz sposób zarządzania sekcjami w Wordzie. 
