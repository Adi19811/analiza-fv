#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Prosty parser faktur (PDF/JPG/PNG) wyciągający podstawowe dane do Excela:
- data (próbuje: Data wystawienia / Data sprzedaży / Data)
- opis usługi/towaru (pierwsza sensowna linia po "Opis"/"Nazwa towaru/usługi" itp.)
- stawka VAT (np. 23%) oraz kwota podatku (jeśli znajdzie)
- kwoty: netto, brutto (jeśli znajdzie)

Wejście: folder z plikami faktur.
Wyjście: plik Excel "wyniki_faktur.xlsx" w bieżącym katalogu.

Instalacja zależności (przykład):
    pip install pdfplumber pymupdf pillow pytesseract pandas openpyxl

Uwaga (OCR):
- Jeśli PDF nie ma warstwy tekstowej, skrypt renderuje strony do obrazów (PyMuPDF)
  i używa Tesseract OCR. Zainstaluj Tesseract w systemie (np. Windows: choco install tesseract,
  macOS: brew install tesseract, Linux: apt-get install tesseract-ocr) i upewnij się, że
  jest w PATH. Dla polskich znaków można doinstalować model języka pl.

Uruchomienie:
    # typowe uruchomienie (wskaż katalog z fakturami)
    python parser_faktur_do_excela.py --input ./faktury --excel wyniki_faktur.xlsx

Tryb autotestu (bez plików):
    python parser_faktur_do_excela.py --selftest

Zachowanie gdy nie podasz --input:
- Jeżeli istnieje folder "./faktury" – zostanie użyty automatycznie.
- W przeciwnym razie – użyty będzie bieżący katalog roboczy (cwd).
  Skrypt wypisze o tym informację.

Autor: ChatGPT (GPT-5 Thinking)
"""
from __future__ import annotations

import argparse
import io
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

# --- Opcjonalne biblioteki do ekstrakcji tekstu ---
try:
    import pdfplumber  # do tekstu z PDF
except Exception:  # pragma: no cover
    pdfplumber = None

try:
    import fitz  # PyMuPDF – renderowanie PDF do obrazów
except Exception:  # pragma: no cover
    fitz = None

try:
    from PIL import Image
except Exception:  # pragma: no cover
    Image = None

try:
    import pytesseract  # OCR
except Exception:  # pragma: no cover
    pytesseract = None


DATE_PATTERNS = [
    r"\b(\d{2}[./-]\d{2}[./-]\d{4})\b",  # 31.12.2025 / 31-12-2025 / 31/12/2025
    r"\b(\d{4}[./-]\d{2}[./-]\d{2})\b",  # 2025-12-31
]

MONEY_NUM = r"\d{1,3}(?:[ .]\d{3})*(?:[.,]\d{2})"

# wzorce z etykietami po PL
AMOUNT_PATTERNS = [
    ("kwota_brutto", rf"(?i)\bkwota\s*brutto\s*[:=]?\s*({MONEY_NUM})"),
    ("kwota_netto", rf"(?i)\bkwota\s*netto\s*[:=]?\s*({MONEY_NUM})"),
    ("brutto", rf"(?i)\bbrutto\s*[:=]?\s*({MONEY_NUM})"),
    ("netto", rf"(?i)\bnetto\s*[:=]?\s*({MONEY_NUM})"),
    ("podatek_vat_kwota", rf"(?i)\b(podatek\s*vat|vat)\s*(?:kwota|suma|razem)?\s*[:=]?\s*({MONEY_NUM})"),
]

VAT_RATE_PATTERNS = [
    r"(?i)\bvat\s*[:=]?\s*(\d{1,2})(?:\,\d+)?\s*%",
    r"(?i)\b(zw|np|oo)\b\s*vat",  # zwolniony/nie podlega/odwrotne obciążenie
]

DESCRIPTION_HINTS = [
    r"(?i)\bopis\b",
    r"(?i)nazwa\s+towaru\s*/?\s*us\w+",
    r"(?i)\bpozycj[ae]\b",
    r"(?i)\btytu[łl]\b",
    r"(?i)\bprzedmiot\s+sprzeda[żz]y\b",
]

DATE_LABELS = [
    r"(?i)data\s+wystawienia",
    r"(?i)data\s+sprzeda[żz]y",
    r"(?i)data\b",
]


@dataclass
class InvoiceData:
    plik: str
    data: Optional[str]
    opis: Optional[str]
    stawka_vat: Optional[str]
    podatek_vat_kwota: Optional[str]
    kwota_netto: Optional[str]
    kwota_brutto: Optional[str]

    def to_row(self) -> Dict[str, Optional[str]]:
        return {
            "plik": self.plik,
            "data": self.data,
            "opis": self.opis,
            "stawka_vat": self.stawka_vat,
            "podatek_vat_kwota": self.podatek_vat_kwota,
            "kwota_netto": self.kwota_netto,
            "kwota_brutto": self.kwota_brutto,
        }


# --- Ekstrakcja tekstu z plików ---

def extract_text_from_pdf(path: Path) -> str:
    text_parts: List[str] = []

    # 1) Spróbuj pdfplumberem
    if pdfplumber is not None:
        try:
            with pdfplumber.open(str(path)) as pdf:
                for page in pdf.pages:
                    t = page.extract_text() or ""
                    if t.strip():
                        text_parts.append(t)
        except Exception:
            pass

    text = "\n".join(text_parts).strip()

    # 2) Jeśli pusto, spróbuj OCR z renderu
    if not text and fitz is not None and pytesseract is not None and Image is not None:
        try:
            doc = fitz.open(str(path))
            for page in doc:
                # Render w wyższej rozdzielczości i jawny format PNG (stabilniejsze dla PIL)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                ocr_text = pytesseract.image_to_string(img, lang="pol+eng")
                text_parts.append(ocr_text)
            text = "\n".join(text_parts)
        except Exception:
            pass

    return text


def extract_text_from_image(path: Path) -> str:
    if pytesseract is None or Image is None:
        return ""
    try:
        img = Image.open(path)
        return pytesseract.image_to_string(img, lang="pol+eng")
    except Exception:
        return ""


# --- Parsowanie treści ---

def _search_first(patterns: List[str], text: str) -> Optional[re.Match]:
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            return m
    return None


def find_date(text: str) -> Optional[str]:
    # Preferuj datę po etykietach, w przeciwnym razie pierwszą datę w tekście
    for label in DATE_LABELS:
        lab = re.search(label + r".*?\b(\d{2}[./-]\d{2}[./-]\d{4}|\d{4}[./-]\d{2}[./-]\d{2})\b", text, re.DOTALL)
        if lab:
            return lab.group(1)
    for pat in DATE_PATTERNS:
        m = re.search(pat, text)
        if m:
            return m.group(1)
    return None


def normalize_money(val: Optional[str]) -> Optional[str]:
    if not val:
        return None
    v = val.replace(" ", "").replace(".", "").replace(",", ".")
    # pozostaw tylko numery i kropkę
    m = re.match(r"\d+(?:\.\d{2})?$", v)
    return v if m else val  # jeśli nietypowe formatowanie – zwróć oryginał


def find_amounts(text: str) -> Dict[str, Optional[str]]:
    found: Dict[str, Optional[str]] = {
        "kwota_brutto": None,
        "kwota_netto": None,
        "podatek_vat_kwota": None,
    }
    for key, pat in AMOUNT_PATTERNS:
        m = re.search(pat, text)
        if m:
            if key == "podatek_vat_kwota":
                found[key] = normalize_money(m.group(2)) if m.lastindex and m.lastindex >= 2 else None
            else:
                found[key if key in found else ("kwota_" + key)] = normalize_money(m.group(1))

    # Uzupełnij brakujące z alternatywnych tagów "brutto/netto"
    if not found.get("kwota_brutto") and (m := re.search(rf"(?i)\bbrutto\b.*?({MONEY_NUM})", text)):
        found["kwota_brutto"] = normalize_money(m.group(1))
    if not found.get("kwota_netto") and (m := re.search(rf"(?i)\bnetto\b.*?({MONEY_NUM})", text)):
        found["kwota_netto"] = normalize_money(m.group(1))

    return found


def find_vat_rate(text: str) -> Optional[str]:
    for pat in VAT_RATE_PATTERNS:
        m = re.search(pat, text)
        if m:
            # np/zw/oo
            if m.lastindex and m.group(1) and m.group(1).lower() in {"zw", "np", "oo"}:
                return m.group(1).upper()
            return (m.group(1) + "%") if m.lastindex else m.group(0)
    # fallback: pierwsze wystąpienie "23%" itp. w pobliżu VAT
    m = re.search(r"(?i)vat[^\n%]{0,30}(\d{1,2}\s*%)", text)
    if m:
        return m.group(1).replace(" ", "")
    return None


def find_description(text: str) -> Optional[str]:
    lines = [l.strip() for l in text.splitlines()]
    # spróbuj po słowach-kluczach
    for i, line in enumerate(lines):
        for hint in DESCRIPTION_HINTS:
            if re.search(hint, line):
                # weź następną niepustą linię jako opis
                for j in range(i + 1, min(i + 6, len(lines))):
                    if lines[j] and not re.match(r"(?i)\b(szt\.|ilo[sś]c|jm|netto|brutto|vat|stawka)\b", lines[j]):
                        return lines[j][:200]
    # fallback: poszukaj linii z myślnikiem/kropką – wygląda na tytuł pozycji
    for line in lines:
        if len(line) > 5 and (" - " in line or re.search(r"\b(us\w+|abonament|usługa|serwis|licencja|naprawa|sprzedaż)\b", line, re.I)):
            return line[:200]
    return None


def parse_invoice(text: str, filename: str) -> InvoiceData:
    data = find_date(text)
    amounts = find_amounts(text)
    vat_rate = find_vat_rate(text)
    desc = find_description(text)

    return InvoiceData(
        plik=filename,
        data=data,
        opis=desc,
        stawka_vat=vat_rate,
        podatek_vat_kwota=amounts.get("podatek_vat_kwota"),
        kwota_netto=amounts.get("kwota_netto"),
        kwota_brutto=amounts.get("kwota_brutto"),
    )


def extract_from_file(path: Path) -> InvoiceData:
    ext = path.suffix.lower()
    text = ""
    if ext == ".pdf":
        text = extract_text_from_pdf(path)
    elif ext in {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}:
        text = extract_text_from_image(path)
    else:
        try:
            text = path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            text = ""

    return parse_invoice(text, path.name)


def run(input_dir: Path, excel_path: Path) -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Podany katalog nie istnieje lub nie jest katalogiem: {input_dir}")
        return

    files = [p for p in input_dir.glob("**/*") if p.is_file() and p.suffix.lower() in {".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}]
    rows: List[Dict[str, Optional[str]]] = []

    for f in sorted(files):
        inv = extract_from_file(f)
        rows.append(inv.to_row())

    if not rows:
        print("Nie znaleziono faktur w podanym katalogu.")
        return

    df = pd.DataFrame(rows, columns=[
        "plik", "data", "opis", "stawka_vat", "podatek_vat_kwota", "kwota_netto", "kwota_brutto"
    ])

    # upewnij się, że rozszerzenie to .xlsx
    excel_path = excel_path if str(excel_path).lower().endswith(".xlsx") else excel_path.with_suffix(".xlsx")

    # zapisz do Excela
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="faktury")
    print(f"Zapisano wyniki do: {excel_path}")


# --- AUTOTESTY --------------------------------------------------------------

def _selftest_samples() -> List[Dict[str, str]]:
    """Zwraca przykładowe wycinki faktur do testów parsera."""
    return [
        {
            "name": "fv_prosta_vat23",
            "text": (
                "Faktura VAT\nData wystawienia: 12-09-2025\n\nNazwa towaru/usługi\nAbonament serwisowy PRO\n\nNetto: 1 000,00 PLN\nVAT 23%\nPodatek VAT: 230,00 PLN\nBrutto: 1 230,00 PLN\n"
            ),
            "expect": {
                "data": "12-09-2025",
                "opis": "Abonament serwisowy PRO",
                "stawka_vat": "23%",
                "podatek_vat_kwota": "230.00",
                "kwota_netto": "1000.00",
                "kwota_brutto": "1230.00",
            },
        },
        {
            "name": "fv_zwolniona",
            "text": (
                "Faktura\nData sprzedaży 2025/09/30\nPrzedmiot sprzedaży\nUsługa konsultacyjna\n\nKwota netto: 500,00 PLN\nVAT: zw\nKwota brutto: 500,00 PLN\n"
            ),
            "expect": {
                "data": "2025/09/30",
                "opis": "Usługa konsultacyjna",
                "stawka_vat": "ZW",
                "podatek_vat_kwota": None,
                "kwota_netto": "500.00",
                "kwota_brutto": "500.00",
            },
        },
        {
            "name": "fv_minimalna",
            "text": (
                "FAKTURA\nData: 31.01.2025\nOpis\nLicencja oprogramowania - roczna\nNetto 200,00 PLN\nBrutto 246,00 PLN\nVAT 23%\n"
            ),
            "expect": {
                "data": "31.01.2025",
                "opis": "Licencja oprogramowania - roczna",
                "stawka_vat": "23%",
                "podatek_vat_kwota": None,
                "kwota_netto": "200.00",
                "kwota_brutto": "246.00",
            },
        },
        {
            "name": "fv_vat8",
            "text": (
                "Faktura VAT\nData wystawienia 2025-10-05\nNazwa towaru/usługi\nUsługa gastronomiczna\nNetto: 1.500,00 PLN\nVAT 8%\nBrutto: 1.620,00 PLN\n"
            ),
            "expect": {
                "data": "2025-10-05",
                "opis": "Usługa gastronomiczna",
                "stawka_vat": "8%",
                "podatek_vat_kwota": None,
                "kwota_netto": "1500.00",
                "kwota_brutto": "1620.00",
            },
        },
        {
            "name": "fv_np",
            "text": (
                "Faktura\nData wystawienia: 01/11/2025\nOpis\nSzkolenie wewnętrzne\nVAT NP\nKwota netto 300,00 PLN\nKwota brutto 300,00 PLN\n"
            ),
            "expect": {
                "data": "01/11/2025",
                "opis": "Szkolenie wewnętrzne",
                "stawka_vat": "NP",
                "podatek_vat_kwota": None,
                "kwota_netto": "300.00",
                "kwota_brutto": "300.00",
            },
        },
    ]


def run_selftests() -> None:
    """Uruchamia zestaw testów parsera na wbudowanych próbkach tekstu."""
    samples = _selftest_samples()
    ok = 0
    for s in samples:
        inv = parse_invoice(s["text"], s["name"])
        row = inv.to_row()
        exp = s["expect"]
        # porównania (łagodne: normalizujemy kwoty)
        def _n(x):
            if x is None:
                return None
            x = x.replace(" ", "").replace(".", "").replace(",", ".")
            return x
        checks = [
            (row["data"], exp["data"]),
            (row["opis"], exp["opis"]),
            (row["stawka_vat"], exp["stawka_vat"]),
            (_n(row["podatek_vat_kwota"]) if row["podatek_vat_kwota"] else None, exp["podatek_vat_kwota"]),
            (_n(row["kwota_netto"]) if row["kwota_netto"] else None, exp["kwota_netto"]),
            (_n(row["kwota_brutto"]) if row["kwota_brutto"] else None, exp["kwota_brutto"]),
        ]
        failed = [(i, a, b) for i, (a, b) in enumerate(checks, 1) if a != b]
        if failed:
            print(f"[FAIL] {s['name']}")
            for i, a, b in failed:
                print(f"  field#{i}: got={a!r} expect={b!r}")
        else:
            print(f"[OK]   {s['name']}")
            ok += 1
    print(f"Wynik: {ok}/{len(samples)} testów zaliczonych.")


def build_argparser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="Parser faktur do Excela")
    ap.add_argument("--input", type=Path, help="Folder z plikami faktur (PDF/JPG/PNG)")
    ap.add_argument("--excel", default=Path("wyniki_faktur.xlsx"), type=Path, help="Ścieżka do pliku .xlsx (domyślnie: wyniki_faktur.xlsx)")
    ap.add_argument("--selftest", action="store_true", help="Uruchom wbudowane testy parsera")
    return ap


if __name__ == "__main__":
    args = build_argparser().parse_args()

    if args.selftest:
        run_selftests()
    else:
        # Inteligentny fallback: jeśli nie podano --input, spróbuj ./faktury, a jak nie ma – użyj cwd
        input_dir = args.input
        if input_dir is None:
            default_dir = Path("faktury") if Path("faktury").exists() else Path.cwd()
            print(f"[INFO] Nie podano --input. Używam: {default_dir}")
            input_dir = default_dir
        run(input_dir, args.excel)
