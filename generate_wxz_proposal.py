#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os

# Kolory
COLOR_HEADER = "1F4E78"
COLOR_HEADER_ALT = "D9D9D9"
COLOR_WARN = "FFEB9C"
COLOR_OK = "C6EFCE"
COLOR_DANGER = "FFC7CE"
COLOR_GRAY = "F2F2F2"
COLOR_GREEN = "E2EFDA"

# Style komponenty
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

wb = Workbook()
wb.remove(wb.active)

# ============================================================
# ARKUSZ 1: STRESZCZENIE
# ============================================================
ws_summary = wb.create_sheet("STRESZCZENIE", 0)

ws_summary['A1'] = "PROPOZYCJA: NOWY SYSTEM KLASYFIKACJI SKU ERGOM"
ws_summary['A1'].font = Font(bold=True, size=12)
ws_summary.merge_cells('A1:C1')

row = 3
summary_data = [
    ("Autor", "Rafał Tarczyński (APKA Prognozowanie)"),
    ("Data", "2026-09-02"),
    ("Status", "DRAFT do konsultacji"),
    ("", ""),
    ("=== PROBLEM ===", ""),
    ("Obecne pole WXZ", "22 wartości mieszają 4 wymiary biznesowe"),
    ("Skutek", "Niespójność, trudne filtrowanie, konflikt interesów użytkowników"),
    ("Zakres", "Cała baza (50 281 SKU)"),
    ("", ""),
    ("=== PROPOZYCJA ===", ""),
    ("Rozbić WXZ na", "4 pola niezależne:"),
    ("", "  • SPOSÓB_ZAOPATRZENIA (4 wartości)"),
    ("", "  • FAZA_ŻYCIA (5 wartości)"),
    ("", "  • TYP_POZYCJI (3 wartości)"),
    ("", "  • MOQ (3 wartości, tylko dla ZAMÓWIENIE)"),
    ("Redukcja", "22 → efektywnie ~15 sensownych kombinacji"),
    ("Nazewnictwo", "Po polsku"),
    ("", ""),
    ("=== KORZYŚCI ===", ""),
    ("", "  • Każda grupa użytkowników filtruje po swoim wymiarze"),
    ("", "  • Zmiana fazy życia nie wymaga zmiany innych atrybutów"),
    ("", "  • Możliwość dodania nowych kombinacji bez tworzenia kodu"),
    ("", "  • Prostsze raportowanie i segmentacja MODEL_V5"),
    ("", ""),
    ("=== UŻYTKOWNICY ===", ""),
    ("Handlowcy", "SPOSÓB_ZAOPATRZENIA + MOQ"),
    ("Zaopatrzenie", "SPOSÓB_ZAOPATRZENIA + TYP_POZYCJI + MOQ"),
    ("Planiści", "TYP_POZYCJI + FAZA_ŻYCIA"),
    ("S&OP", "FAZA_ŻYCIA + ABC × XYZ segmentacja"),
    ("Dział Rozwoju", "FAZA_ŻYCIA (owner słownika)"),
    ("", ""),
    ("=== PYTANIA OTWARTE ===", ""),
    ("1. Kod N", "Samo N — jak zmapować? Decyzja per SKU lub reguła."),
    ("2. Sprawdź dostępność", "Kolumna — osobny wymiar czy pozostawić bez zmian?"),
    ("3. PRZEPAK", "Czy potrzebne pole przepakowywany_z?"),
    ("4. BR (Badanie rynku)", "Czy faza czy flaga?"),
]

for label, value in summary_data:
    ws_summary[f'A{row}'] = label
    ws_summary[f'B{row}'] = value

    if label.startswith("==="):
        ws_summary[f'A{row}'].font = Font(bold=True, size=11)
        ws_summary[f'A{row}'].fill = PatternFill(start_color=COLOR_HEADER_ALT, end_color=COLOR_HEADER_ALT, fill_type='solid')
    elif label and not label.startswith("  "):
        ws_summary[f'A{row}'].font = Font(bold=True)

    row += 1

ws_summary.column_dimensions['A'].width = 30
ws_summary.column_dimensions['B'].width = 50
ws_summary.freeze_panes = 'A2'

# ============================================================
# ARKUSZ 2: WYMIARY
# ============================================================
ws_dims = wb.create_sheet("WYMIARY", 1)

row = 1

# SPOSÓB_ZAOPATRZENIA
ws_dims[f'A{row}'] = "SPOSÓB_ZAOPATRZENIA"
ws_dims[f'B{row}'] = "Główni użytkownicy: Handlowcy, Zaopatrzenie"
ws_dims[f'A{row}'].font = Font(bold=True, color="FFFFFF", size=11)
ws_dims[f'B{row}'].font = Font(bold=True, color="FFFFFF", size=11)
ws_dims[f'A{row}'].fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type='solid')
ws_dims[f'B{row}'].fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type='solid')
ws_dims.merge_cells(f'A{row}:C{row}')
row += 1

ws_dims[f'A{row}'] = "Definicja: W jaki sposób dostarczamy produkt klientowi"
row += 1

# Header
ws_dims[f'A{row}'] = "Kod"
ws_dims[f'B{row}'] = "Opis"
ws_dims[f'C{row}'] = "Znaczenie operacyjne"
for col in ['A', 'B', 'C']:
    ws_dims[f'{col}{row}'].font = Font(bold=True)
    ws_dims[f'{col}{row}'].fill = PatternFill(start_color=COLOR_HEADER_ALT, end_color=COLOR_HEADER_ALT, fill_type='solid')
row += 1

zaopatrz_data = [
    ("MAGAZYN", "Magazynowy", "Dostępny od ręki z zapasu"),
    ("MAGAZYN_WAR", "Magazynowy warunkowo", "Dostępny z ograniczeniami stanu"),
    ("ZAMÓWIENIE", "Na zamówienie", "Produkowany pod zlecenie"),
    ("KARTA_UZG", "Karta uzgodnień", "Wg umowy z konkretnym klientem"),
]

for kod, opis, znaczenie in zaopatrz_data:
    ws_dims[f'A{row}'] = kod
    ws_dims[f'B{row}'] = opis
    ws_dims[f'C{row}'] = znaczenie
    row += 1

row += 2

# FAZA_ŻYCIA
ws_dims[f'A{row}'] = "FAZA_ŻYCIA"
ws_dims[f'B{row}'] = "Główni użytkownicy: Dział Rozwoju, S&OP, Planiści"
ws_dims[f'A{row}'].font = Font(bold=True, color="FFFFFF", size=11)
ws_dims[f'B{row}'].font = Font(bold=True, color="FFFFFF", size=11)
ws_dims[f'A{row}'].fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type='solid')
ws_dims[f'B{row}'].fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type='solid')
ws_dims.merge_cells(f'A{row}:C{row}')
row += 1

ws_dims[f'A{row}'] = "Definicja: W jakiej fazie cyklu życia jest produkt"
row += 1

ws_dims[f'A{row}'] = "Kod"
ws_dims[f'B{row}'] = "Opis"
ws_dims[f'C{row}'] = "Znaczenie operacyjne"
for col in ['A', 'B', 'C']:
    ws_dims[f'{col}{row}'].font = Font(bold=True)
    ws_dims[f'{col}{row}'].fill = PatternFill(start_color=COLOR_HEADER_ALT, end_color=COLOR_HEADER_ALT, fill_type='solid')
row += 1

faza_data = [
    ("NOWY", "W trakcie wdrożenia", "Wprowadzany do oferty"),
    ("AKTYWNY", "Aktywna sprzedaż", "Standardowa sprzedaż"),
    ("WYGASZANY", "Wyprzedaż stanów", "Sprzedajemy do wyczyszczenia magazynu"),
    ("WYCOFANY", "Usunięty", "Zamknięty, tylko archiwum"),
    ("BADANIE", "Badanie rynku", "Prezentacja B2B, test popytu"),
]

for kod, opis, znaczenie in faza_data:
    ws_dims[f'A{row}'] = kod
    ws_dims[f'B{row}'] = opis
    ws_dims[f'C{row}'] = znaczenie
    row += 1

row += 2

# TYP_POZYCJI
ws_dims[f'A{row}'] = "TYP_POZYCJI"
ws_dims[f'B{row}'] = "Główni użytkownicy: Planiści, Zaopatrzenie, Marketing"
ws_dims[f'A{row}'].font = Font(bold=True, color="FFFFFF", size=11)
ws_dims[f'B{row}'].font = Font(bold=True, color="FFFFFF", size=11)
ws_dims[f'A{row}'].fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type='solid')
ws_dims[f'B{row}'].fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type='solid')
ws_dims.merge_cells(f'A{row}:C{row}')
row += 1

ws_dims[f'A{row}'] = "Definicja: Rodzaj artykułu technicznie"
row += 1

ws_dims[f'A{row}'] = "Kod"
ws_dims[f'B{row}'] = "Opis"
ws_dims[f'C{row}'] = "Znaczenie operacyjne"
for col in ['A', 'B', 'C']:
    ws_dims[f'{col}{row}'].font = Font(bold=True)
    ws_dims[f'{col}{row}'].fill = PatternFill(start_color=COLOR_HEADER_ALT, end_color=COLOR_HEADER_ALT, fill_type='solid')
row += 1

typ_data = [
    ("WYRÓB", "Wyrób gotowy", "Produkt sprzedażowy (domyślnie)"),
    ("SUROWIEC", "Surowiec / półfabrykat", "Składnik do produkcji wyrobu"),
    ("PRZEPAK", "Do przepakowania", "Techniczny wariant sprzedażowego SKU"),
]

for kod, opis, znaczenie in typ_data:
    ws_dims[f'A{row}'] = kod
    ws_dims[f'B{row}'] = opis
    ws_dims[f'C{row}'] = znaczenie
    row += 1

row += 2

# MOQ
ws_dims[f'A{row}'] = "MOQ"
ws_dims[f'B{row}'] = "Główni użytkownicy: Handlowcy | Zastosowanie: TYLKO dla SPOSÓB_ZAOPATRZENIA = ZAMÓWIENIE"
ws_dims[f'A{row}'].font = Font(bold=True, color="FFFFFF", size=11)
ws_dims[f'B{row}'].font = Font(bold=True, color="FFFFFF", size=11)
ws_dims[f'A{row}'].fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type='solid')
ws_dims[f'B{row}'].fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type='solid')
ws_dims.merge_cells(f'A{row}:C{row}')
row += 1

ws_dims[f'A{row}'] = "Kod"
ws_dims[f'B{row}'] = "Opis"
ws_dims[f'C{row}'] = "Znaczenie operacyjne"
for col in ['A', 'B', 'C']:
    ws_dims[f'{col}{row}'].font = Font(bold=True)
    ws_dims[f'{col}{row}'].fill = PatternFill(start_color=COLOR_HEADER_ALT, end_color=COLOR_HEADER_ALT, fill_type='solid')
row += 1

moq_data = [
    ("MOQ_SERIA", "MOQ = minimalna seria", "Klient musi zamówić całą serię"),
    ("MOQ_JEDN", "MOQ = 1 jednostka mag.", "Klient może zamówić 1 sztukę"),
    ("MOQ_CZĘŚĆ", "MOQ mniej niż seria", "Klient może mniej, produkujemy pełną serię"),
]

for kod, opis, znaczenie in moq_data:
    ws_dims[f'A{row}'] = kod
    ws_dims[f'B{row}'] = opis
    ws_dims[f'C{row}'] = znaczenie
    row += 1

ws_dims.column_dimensions['A'].width = 20
ws_dims.column_dimensions['B'].width = 25
ws_dims.column_dimensions['C'].width = 50

# ============================================================
# ARKUSZ 3: MAPPING
# ============================================================
ws_map = wb.create_sheet("MAPPING", 2)

row = 1
headers = ["WXZ", "Opis obecny", "SPOSÓB_ZAOPATRZENIA", "FAZA_ŻYCIA", "TYP_POZYCJI", "MOQ", "Komentarz"]
for col, header in enumerate(headers, 1):
    cell = ws_map.cell(row, col, header)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type='solid')
    cell.border = thin_border

row += 1

mapping_data = [
    ("W", "MAGAZYNOWY", "MAGAZYN", "AKTYWNY", "WYRÓB", "-", "Domyślny wariant magazynowy"),
    ("X", "MAGAZYNOWY WARUNKOWO", "MAGAZYN_WAR", "AKTYWNY", "WYRÓB", "-", "Warunkowa dostępność"),
    ("XS", "SUROWIEC MAGAZYNOWY", "MAGAZYN", "AKTYWNY", "SUROWIEC", "-", ""),
    ("Z", "NA ZAMÓWIENIE - MOQ = MIN. SERIA PROD", "ZAMÓWIENIE", "AKTYWNY", "WYRÓB", "MOQ_SERIA", "Klasyczny MTO"),
    ("Z1", "NA ZAMÓWIENIE, MOQ = 1 jedn. mag", "ZAMÓWIENIE", "AKTYWNY", "WYRÓB", "MOQ_JEDN", ""),
    ("Zc", "NA ZAMÓWIENIE, MOQ < MIN. SERIA PROD", "ZAMÓWIENIE", "AKTYWNY", "WYRÓB", "MOQ_CZĘŚĆ", ""),
    ("ZS", "SUROWIEC NA WYRÓB Z", "ZAMÓWIENIE", "AKTYWNY", "SUROWIEC", "MOQ_SERIA", ""),
    ("N", "W TRAKCIE WDROŻENIA", "DECYZJA", "NOWY", "WYRÓB", "-", "⚠️ Wymaga decyzji per SKU"),
    ("NX", "NOWY WYRÓB MAGAZYNOWY", "MAGAZYN", "NOWY", "WYRÓB", "-", ""),
    ("NZ", "NOWY WYRÓB NA ZAMÓWIENIE", "ZAMÓWIENIE", "NOWY", "WYRÓB", "MOQ_SERIA", ""),
    ("KU", "ZAMÓWIENIA PRZEZ KARTĘ UZGODNIEŃ", "KARTA_UZG", "AKTYWNY", "WYRÓB", "-", ""),
    ("MZ", "WYPRZEDAŻ STANÓW MAG.", "MAGAZYN", "WYGASZANY", "WYRÓB", "-", ""),
    ("MZKU", "po wyprzedaży → KU", "KARTA_UZG", "WYGASZANY", "WYRÓB", "-", ""),
    ("MZP", "WYPRZEDAŻ + PRODUKCJA do likwidacji", "ZAMÓWIENIE", "WYGASZANY", "WYRÓB", "MOQ_SERIA", ""),
    ("U", "USUNIĘTY", "-", "WYCOFANY", "WYRÓB", "-", "Historyczny, brak nowych zamówień"),
    ("US", "SUROWIEC USUNIĘTY", "-", "WYCOFANY", "SUROWIEC", "-", ""),
    ("BR", "BADANIE RYNKU / WYCOFANE do B2B", "-", "BADANIE", "WYRÓB", "-", "⚠️ Do rozstrzygnięcia"),
    ("WW", "INDEKS DO PRZEPAKOWANIA W", "MAGAZYN", "AKTYWNY", "PRZEPAK", "-", ""),
    ("XX", "INDEKS DO PRZEPAKOWANIA X", "MAGAZYN_WAR", "AKTYWNY", "PRZEPAK", "-", ""),
    ("ZZ", "INDEKS DO PRZEPAKOWANIA Z", "ZAMÓWIENIE", "AKTYWNY", "PRZEPAK", "MOQ_SERIA", ""),
    ("ZCZC", "INDEKS DO PRZEPAKOWANIA ZC", "ZAMÓWIENIE", "AKTYWNY", "PRZEPAK", "MOQ_CZĘŚĆ", ""),
    ("(nowa kombinacja)", "Dostępne dla przyszłych potrzeb", "", "", "", "", "Elastyczność systemu"),
]

for wxz, opis, zaopatrz, faza, typ, moq, kom in mapping_data:
    for col, val in enumerate([wxz, opis, zaopatrz, faza, typ, moq, kom], 1):
        cell = ws_map.cell(row, col, val)
        cell.border = thin_border

        # Podświetlenie wierszy problemowych
        if "⚠️" in kom or zaopatrz == "DECYZJA":
            cell.fill = PatternFill(start_color=COLOR_WARN, end_color=COLOR_WARN, fill_type='solid')
        elif zaopatrz == "-" or (zaopatrz == "" and wxz == "(nowa kombinacja)"):
            cell.fill = PatternFill(start_color=COLOR_GRAY, end_color=COLOR_GRAY, fill_type='solid')

    row += 1

ws_map.column_dimensions['A'].width = 10
ws_map.column_dimensions['B'].width = 45
ws_map.column_dimensions['C'].width = 18
ws_map.column_dimensions['D'].width = 15
ws_map.column_dimensions['E'].width = 12
ws_map.column_dimensions['F'].width = 12
ws_map.column_dimensions['G'].width = 45
ws_map.freeze_panes = 'A2'

# Auto-filter
ws_map.auto_filter.ref = f'A1:G{row-1}'

# ============================================================
# ARKUSZ 4: MATRIX
# ============================================================
ws_matrix = wb.create_sheet("MATRIX", 3)

row = 1
ws_matrix['A1'] = "MATRYCA KOMBINACJI SENSOWNYCH: SPOSÓB_ZAOPATRZENIA × FAZA_ŻYCIA"
ws_matrix['A1'].font = Font(bold=True, size=11)
ws_matrix.merge_cells('A1:F1')

row = 3
headers_matrix = ["SPOSÓB_ZAOPATRZENIA", "NOWY", "AKTYWNY", "WYGASZANY", "WYCOFANY", "BADANIE"]
for col, header in enumerate(headers_matrix, 1):
    cell = ws_matrix.cell(row, col, header)
    cell.font = Font(bold=True)
    cell.fill = PatternFill(start_color=COLOR_HEADER_ALT, end_color=COLOR_HEADER_ALT, fill_type='solid')
    cell.border = thin_border

row += 1

matrix_data = [
    ("MAGAZYN", "✓", "✓", "✓", "-", "-"),
    ("MAGAZYN_WAR", "✓", "✓", "✓", "-", "-"),
    ("ZAMÓWIENIE", "✓", "✓", "✓", "-", "-"),
    ("KARTA_UZG", "?", "✓", "✓", "-", "-"),
    ("(brak)", "-", "-", "-", "✓", "✓"),
]

for sposob, nowy, akt, wyg, wyc, bad in matrix_data:
    col_num = 1
    for val in [sposob, nowy, akt, wyg, wyc, bad]:
        cell = ws_matrix.cell(row, col_num, val)
        cell.border = thin_border

        if val == "✓":
            cell.fill = PatternFill(start_color=COLOR_OK, end_color=COLOR_OK, fill_type='solid')
            cell.font = Font(bold=True, color="008000")
        elif val == "?":
            cell.fill = PatternFill(start_color=COLOR_WARN, end_color=COLOR_WARN, fill_type='solid')
            cell.font = Font(bold=True)
        elif val == "-":
            cell.fill = PatternFill(start_color=COLOR_GRAY, end_color=COLOR_GRAY, fill_type='solid')

        col_num += 1

    row += 1

row += 2
ws_matrix[f'A{row}'] = "LEGENDA:"
ws_matrix[f'A{row}'].font = Font(bold=True)
row += 1
ws_matrix[f'A{row}'] = "✓ = kombinacja typowa"
row += 1
ws_matrix[f'A{row}'] = "? = kombinacja rzadka, wymaga decyzji"
row += 1
ws_matrix[f'A{row}'] = "- = kombinacja nie ma sensu"
row += 1
ws_matrix[f'A{row}'] = ""
row += 1
ws_matrix[f'A{row}'] = "NOTATKI:"
ws_matrix[f'A{row}'].font = Font(bold=True)
row += 1
ws_matrix[f'A{row}'] = "• WYCOFANY nie ma SPOSÓB_ZAOPATRZENIA (już nie sprzedajemy)"
row += 1
ws_matrix[f'A{row}'] = "• BADANIE ma osobny status, poza standardowym zaopatrzeniem"
row += 1
ws_matrix[f'A{row}'] = "• MOQ dotyczy tylko ZAMÓWIENIE — dla innych puste"

ws_matrix.column_dimensions['A'].width = 20
ws_matrix.column_dimensions['B'].width = 12
ws_matrix.column_dimensions['C'].width = 12
ws_matrix.column_dimensions['D'].width = 12
ws_matrix.column_dimensions['E'].width = 12
ws_matrix.column_dimensions['F'].width = 12

# ============================================================
# ARKUSZ 5: UŻYTKOWNICY
# ============================================================
ws_users = wb.create_sheet("UŻYTKOWNICY", 4)

row = 1
headers_users = ["Grupa użytkowników", "Cel filtrowania", "Wymiary", "Przykład użycia"]
for col, header in enumerate(headers_users, 1):
    cell = ws_users.cell(row, col, header)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type='solid')
    cell.border = thin_border

row += 1

users_data = [
    ("Handlowcy", "Co obiecać klientowi", "ZAOPATRZ + MOQ", "Filtr: ZAOPATRZ=MAGAZYN → dostępne od ręki\nFiltr: ZAOPATRZ=ZAMÓWIENIE + MOQ_SERIA → wyjaśnienie o serii"),
    ("Zaopatrzenie", "Co uruchomić w produkcji", "ZAOPATRZ + TYP + MOQ", "Filtr: TYP=SUROWIEC + ZAOPATRZ=ZAMÓWIENIE → lista do zamówienia u dostawców"),
    ("Planiści", "Co planować", "TYP + FAZA", "Filtr: TYP=WYRÓB + FAZA=AKTYWNY → które wyroby wymagają planu produkcji"),
    ("S&OP", "Segmentacja forecast", "FAZA (zamiast WXZ)", "MODEL_V5: ABC × XYZ × FAZA\nNOWY → bootstrap, WYGASZANY → decay, WYCOFANY → poza scope"),
    ("Dział Rozwoju", "Co jest w wdrożeniu (owner słownika)", "FAZA", "Filtr: FAZA=NOWY → lista wdrożeń w toku\nFiltr: FAZA=BADANIE → prezentacje B2B"),
    ("Marketing", "Co pokazać w katalogu", "TYP + FAZA", "Filtr: TYP=WYRÓB + FAZA≠WYCOFANY"),
]

for grupa, cel, wymiary, przyklad in users_data:
    cell_grupa = ws_users.cell(row, 1, grupa)
    cell_cel = ws_users.cell(row, 2, cel)
    cell_wymiary = ws_users.cell(row, 3, wymiary)
    cell_przyklad = ws_users.cell(row, 4, przyklad)

    # Alternating colors
    fill_color = COLOR_GRAY if row % 2 == 0 else None
    if fill_color:
        for cell in [cell_grupa, cell_cel, cell_wymiary, cell_przyklad]:
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

    for cell in [cell_grupa, cell_cel, cell_wymiary, cell_przyklad]:
        cell.border = thin_border
        cell.alignment = Alignment(wrap_text=True, vertical='top')

    ws_users.row_dimensions[row].height = 40
    row += 1

ws_users.column_dimensions['A'].width = 20
ws_users.column_dimensions['B'].width = 30
ws_users.column_dimensions['C'].width = 25
ws_users.column_dimensions['D'].width = 45

# ============================================================
# ARKUSZ 6: PYTANIA_OTWARTE
# ============================================================
ws_questions = wb.create_sheet("PYTANIA_OTWARTE", 5)

row = 1
headers_q = ["#", "Pytanie", "Propozycja rozwiązania", "Decyzja", "Owner"]
for col, header in enumerate(headers_q, 1):
    cell = ws_questions.cell(row, col, header)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type='solid')
    cell.border = thin_border

row += 1

questions_data = [
    ("1", "Kod N (samo N) — niejednoznaczny", "Decyzja per SKU: prawdopodobnie dla większości → SPOSÓB=MAGAZYN lub ZAMÓWIENIE zależy od intencji", "", "Dział Rozwoju"),
    ("2", "Kolumna 'Sprawdź dostępność powyżej' (0, 1, 4)", "Opcja A: osobny wymiar SPRAWDZ_KASKADĘ\nOpcja B: pozostawić bez zmian", "", "IT/AX + S&OP"),
    ("3", "PRZEPAK — czy referencja do źródła?", "Dodać pole 'przepakowywany_z' (link do SKU źródłowego)\nlub: tylko flaga TYP=PRZEPAK", "", "Dział Rozwoju"),
    ("4", "BR (Badanie rynku) — faza czy flaga?", "Opcja A: FAZA=BADANIE (osobna faza)\nOpcja B: flaga PREZENTACJA_B2B + FAZA=AKTYWNY", "", "Dział Rozwoju"),
    ("5", "Backward compatibility WXZ", "Opcja A: WXZ obliczane z 4 pól\nOpcja B: wygasić całkowicie\nOpcja C: archive read-only", "", "IT/AX + S&OP"),
    ("6", "Mechanizm AX 2012 R3", "Do sprawdzenia z IT:\nA) Product dimensions\nB) Item attribute values\nC) Custom fields", "", "IT/AX"),
    ("7", "Wpływ na MODEL_V5 (68 grup)", "Zmiana segmentacji z ABC×XYZ×WXZ na ABC×XYZ×FAZA_ŻYCIA — konsultacja z owner MODEL_V5", "", "S&OP"),
    ("8", "Nazewnictwo polskie", "Zaproponowane po polsku (MAGAZYN, ZAMÓWIENIE...)\n— walidacja u operacyjnych", "", "Handlowcy + Zaopatrzenie"),
    ("9", "Strategia wdrożenia", "Rekomendacja: pilot 1 grupa (np. E11), potem cała baza", "", "Rafał + S&OP"),
    ("10", "Timeline", "Rekomendacja: 4-6 miesięcy\n(2-3 tyg konsultacje + 1 miesiąc pilot + 2-3 miesiące full migration)", "", "Rafał"),
]

for num, pytanie, propozycja, decyzja, owner in questions_data:
    for col, val in enumerate([num, pytanie, propozycja, decyzja, owner], 1):
        cell = ws_questions.cell(row, col, val)
        cell.border = thin_border
        cell.alignment = Alignment(wrap_text=True, vertical='top')

        if col == 4:  # Kolumna Decyzja
            cell.fill = PatternFill(start_color=COLOR_WARN, end_color=COLOR_WARN, fill_type='solid')

    ws_questions.row_dimensions[row].height = 60
    row += 1

ws_questions.column_dimensions['A'].width = 5
ws_questions.column_dimensions['B'].width = 45
ws_questions.column_dimensions['C'].width = 40
ws_questions.column_dimensions['D'].width = 25
ws_questions.column_dimensions['E'].width = 25

# ============================================================
# ARKUSZ 7: WDROŻENIE
# ============================================================
ws_impl = wb.create_sheet("WDROŻENIE", 6)

row = 1
headers_impl = ["Faza", "Czas", "Zakres", "Owner", "Deliverables", "Status"]
for col, header in enumerate(headers_impl, 1):
    cell = ws_impl.cell(row, col, header)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type='solid')
    cell.border = thin_border

row += 1

impl_data = [
    ("Faza 1: Konsultacje", "2-3 tygodnie", "Rozmowy z IT/AX, S&OP, Rozwój, użytkownicy operacyjni", "Rafał (APKA)", "• Odpowiedzi na 10 pytań otwartych\n• Wybór mechanizmu AX\n• Wybór strategii backward compat WXZ", "PENDING"),
    ("Faza 2: Decyzja techniczna", "1 tydzień", "Zatwierdzenie mechanizmu AX, plan migracji", "IT/AX + Dział Rozwoju", "• Schemat bazy zaakceptowany\n• Skrypt migracji WXZ → 4 wymiary", "PENDING"),
    ("Faza 3: Pilot", "1 miesiąc", "Migracja grupy pilotażowej (np. E11 Kabelowe)", "Dział Rozwoju + APKA", "• Migracja wszystkich SKU tej grupy\n• Test raportów, procesów AX, SSAS, APKA\n• Walidacja użytkowników", "PENDING"),
    ("Faza 4: Full migration", "2-3 miesiące", "Migracja pozostałych ~40 000 SKU", "IT/AX + Dział Rozwoju", "• Batch migracja\n• Aktualizacja wszystkich raportów\n• Szkolenia użytkowników", "PENDING"),
    ("Faza 5: Wygaszenie WXZ", "3-6 miesięcy po Fazie 4", "Decyzja o pozostawieniu obliczanego WXZ vs usunięcie", "S&OP + Dział Rozwoju", "• Decyzja finalna\n• Wygaszenie pola (jeśli decyzja)", "PENDING"),
]

for faza, czas, zakres, owner, deliverables, status in impl_data:
    for col, val in enumerate([faza, czas, zakres, owner, deliverables, status], 1):
        cell = ws_impl.cell(row, col, val)
        cell.border = thin_border
        cell.alignment = Alignment(wrap_text=True, vertical='top')

        if faza.startswith("Faza"):
            cell.fill = PatternFill(start_color=COLOR_HEADER_ALT, end_color=COLOR_HEADER_ALT, fill_type='solid')
            cell.font = Font(bold=True)

    ws_impl.row_dimensions[row].height = 70
    row += 1

ws_impl.column_dimensions['A'].width = 25
ws_impl.column_dimensions['B'].width = 15
ws_impl.column_dimensions['C'].width = 35
ws_impl.column_dimensions['D'].width = 25
ws_impl.column_dimensions['E'].width = 35
ws_impl.column_dimensions['F'].width = 12

# ============================================================
# ARKUSZ 8: RYZYKA
# ============================================================
ws_risks = wb.create_sheet("RYZYKA", 7)

row = 1
headers_risks = ["#", "Ryzyko", "Prawdopodobieństwo", "Impact", "Priorytet", "Mitygacja", "Owner"]
for col, header in enumerate(headers_risks, 1):
    cell = ws_risks.cell(row, col, header)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type='solid')
    cell.border = thin_border

row += 1

risks_data = [
    ("1", "AX 2012 R3 nie obsługuje 4 wymiarów w preferowany sposób", "Średnie", "Wysoki", "WYSOKI", "Rozmowa z IT/AX na starcie przed pilotem", "IT/AX"),
    ("2", "Migracja psuje istniejące raporty (SSAS/AX/APKA)", "Wysokie", "Wysoki", "KRYT.", "Backward compat WXZ obliczane z 4 pól", "IT/AX + S&OP"),
    ("3", "MODEL_V5 wymaga refaktoringu segmentacji", "Średnie", "Średni", "ŚREDNI", "Zachować WXZ obliczane początkowo", "S&OP"),
    ("4", "Użytkownicy nie akceptują polskiego nazewnictwa", "Niskie", "Średni", "NISKI", "Warsztaty konsultacyjne", "Rafał"),
    ("5", "Kod N i Sprawdź dostępność nie dają się jednoznacznie zmapować", "Wysokie", "Niski", "ŚREDNI", "Decyzja per SKU przez Dział Rozwoju", "Dział Rozwoju"),
    ("6", "Timeline przekroczony (>6 miesięcy)", "Średnie", "Średni", "ŚREDNI", "Ścisły monitor per faza, escalacja po każdej fazie", "Rafał"),
    ("7", "Konflikt priorytetów z bieżącymi zadaniami APKA", "Wysokie", "Średni", "WYSOKI", "Uzgodnienie z zarządem (S&OP + Dział Rozwoju)", "S&OP + Dział Rozwoju"),
]

for num, ryzyko, prawd, impact, priorytet, mitygacja, owner in risks_data:
    for col, val in enumerate([num, ryzyko, prawd, impact, priorytet, mitygacja, owner], 1):
        cell = ws_risks.cell(row, col, val)
        cell.border = thin_border
        cell.alignment = Alignment(wrap_text=True, vertical='top')

        # Kolorowanie priorytetu
        if priorytet == "KRYT.":
            cell.fill = PatternFill(start_color=COLOR_DANGER, end_color=COLOR_DANGER, fill_type='solid')
            if col == 5:
                cell.font = Font(bold=True, color="FFFFFF")
        elif priorytet == "WYSOKI":
            cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type='solid')
            if col == 5:
                cell.font = Font(bold=True)
        elif priorytet == "ŚREDNI":
            cell.fill = PatternFill(start_color=COLOR_WARN, end_color=COLOR_WARN, fill_type='solid')
            if col == 5:
                cell.font = Font(bold=True)
        elif priorytet == "NISKI":
            cell.fill = PatternFill(start_color=COLOR_GREEN, end_color=COLOR_GREEN, fill_type='solid')
            if col == 5:
                cell.font = Font(bold=True)

    ws_risks.row_dimensions[row].height = 50
    row += 1

ws_risks.column_dimensions['A'].width = 5
ws_risks.column_dimensions['B'].width = 45
ws_risks.column_dimensions['C'].width = 15
ws_risks.column_dimensions['D'].width = 10
ws_risks.column_dimensions['E'].width = 12
ws_risks.column_dimensions['F'].width = 40
ws_risks.column_dimensions['G'].width = 25

# ============================================================
# ZAPIS
# ============================================================
output_path = "output/spreadsheet/agent_sop_exports/propozycja_wxz_refactor_202609.xlsx"
wb.save(output_path)

print(f"✓ Plik zapisany: {output_path}")
print(f"✓ Rozmiar: {os.path.getsize(output_path) / 1024:.1f} KB")
print(f"✓ Arkusze: {len(wb.sheetnames)}")
for i, sheet in enumerate(wb.sheetnames, 1):
    print(f"  {i}. {sheet}")
