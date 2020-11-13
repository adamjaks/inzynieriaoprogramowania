# Inzynieria oprogramowania - PROJEKT

#### Skład grupy: ####
Adam Jakś, Bartosz Farbowski 
#### Temat projektu: ####
Aplikacja do zastępstw 
#### Technologie: ####
Język C# 
#### Opis projektu: ####
  Program na na celu zachowanie ciągłości pracy w firmie, tak, aby jak najoptymalniej wyszukać zastępstwo z listy osób, które na dane zastępstwo mogą przystać (na zasadzie podpowiedzi). 
  Główne założenie to pokrycie wszystkich dni roboczych przez osobę obsługującą program. Jej aktywność musi być rejestrowana w postaci wejścia/wyjscia z programu, (opcja logowania do rozważenia).

Główne funkcjonalności:
- pobieranie aktualnie zalogowanego uzytkownika (tworzenie logu z wejsciem/wyjściem)
- pobieranie daty aby zaczytać excele z odpowiedniego folderu na dysku (2020/10/*.xlsx)
- zaczytywanie danych z pliku xls/xlsx do programu
- opcja edytowania zawartości excela- oryginał przechodzi do archiwum, tworzona jest kopia (nazwa poprzedniego + data zmiany ddMM)
- wyszukiwanie "nieobsadzonych" dni
- podpowiadanie użytkowników "na zastępstwo" według klucza dostępność + preferowana lokalizacja
- wysylanie maila do wskazanych osob z aktualnym harmonogramem
