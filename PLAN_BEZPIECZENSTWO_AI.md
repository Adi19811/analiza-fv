# Plan: bezpieczne wdrażanie AI w firmach (VPS Hostinger „KVM 2")

Data: 2026-11-25 · Status: plan roboczy, dokument do odhaczania

## 1. Cel
Zostać specjalistą od **bezpiecznego wdrażania AI w firmach**. Brak zaufania do
bezpieczeństwa danych to główna bolączka i bariera firm przy wdrożeniach AI.
Cel końcowy: usługi dla firm (audyt gotowości, architektura, wdrożenie, dokumentacja).

## 2. Punkt wyjścia
- VPS na Hostingerze, plan „KVM 2" (do sprawdzenia w panelu: ok. 2 vCPU, 8 GB RAM).
- Na VPS-ie działa już jedna aplikacja (wieloagentowa, dla firmy rekrutacyjnej,
  tworzona z Claude Code). **Stan zabezpieczeń nieznany**, serwer prawdopodobnie
  nieutwardzony.
- Poziom techniczny: początkujący (nauka od zera, krok po kroku).
- Narzędzia: Claude Code (agent), terminal w przeglądarce w hPanelu.

## 3. Ustalenia (z rozmowy)
| Temat | Wniosek |
|---|---|
| Claude Code na własnym VPS | Tak, ale na utwardzonym serwerze, użytkownik non-root, tryb z pytaniem o zgodę |
| Izolacja | Kontener to za mało. microVM (Firecracker/Kata) to obecny standard, ale wymaga KVM wewnątrz VPS-a, co na tanim VPS-ie zwykle nie działa. **Realny wariant: Docker + gVisor** |
| Dane firmy | Izolacja nie chroni przed tym, co wysyłasz do modelu. Liczy się wybór dostawcy: Anthropic Enterprise (ZDR po kwalifikacji), Bedrock/Vertex w UE przez prywatną sieć, albo model lokalny |
| Model lokalny | Daje brak wysyłki danych na zewnątrz, ale słabsza jakość, potrzebny sprzęt (GPU), a ryzyka agenta zostają. 2 vCPU/8 GB to tylko do nauki |
| Chmura z umową | Umowa (DPA, ZDR) to podstawa prawna, **nie gwarancja** bezpieczeństwa. Potrzebne warstwy techniczne |
| Zabezpieczony serwer ≠ bezpieczna aplikacja | Serwer chroni infrastrukturę. Aplikacja wymaga własnych zabezpieczeń (kod, zależności, sekrety, uprawnienia) |

Uwaga: ustalenia o izolacji i dostawcach pochodzą głównie z blogów producentów.
Przed decyzjami biznesowymi zweryfikuj w dokumentacji Anthropic, AWS i Google.

## 4. Plan nauki (ok. 6 miesięcy przy 6-8 h/tydz.; pierwsze efekty w tydzień)
1. **Fundamenty (miesiąc 1):** Linux, SSH, zapora, Docker, model zagrożeń.
2. **Izolacja agentów (miesiąc 2):** gVisor, allowlista ruchu wychodzącego, sekrety, Claude Code w izolacji, `settings.json`.
3. **Bezpieczeństwo AI (miesiąc 3):** OWASP Top 10 for LLM, MITRE ATLAS, NIST AI RMF, prompt injection, red teaming.
4. **Dostawcy i modele lokalne (miesiąc 4):** ZDR, Bedrock/Vertex (region UE, prywatna sieć), Ollama/vLLM.
5. **Zgodność i praktyka (miesiące 5-6):** RODO (DPA, DPIA, transfer danych), EU AI Act, ISO 27001/42001, polityka użycia AI, projekt końcowy.

Do konsultacji z prawnikiem lub IOD: DPA, transfery danych, klasyfikacja ryzyka AI Act.

## 5. Najbliższe kroki (tydzień 1)
- [ ] **Snapshot VPS-a** w hPanelu (kopia przed zmianami).
- [ ] Zalogować się przez Browser terminal (`whoami`, `cat /etc/os-release`).
- [ ] Audyt stanu istniejącej aplikacji (tylko odczyt): `docker ps`, `ss -tulpn`, `ufw status`, `grep -E '^(PermitRootLogin|PasswordAuthentication)' /etc/ssh/sshd_config`.
- [ ] Sprawdzić `ls -l /dev/kvm` i `systemd-detect-virt` (czy możliwe microVM).
- [ ] Utwardzić serwer: aktualizacje, użytkownik non-root z `sudo`, klucze SSH, zapora (ufw + zapora w panelu), fail2ban, snapshoty.
- [ ] Zamknąć porty bazy i usług wewnętrznych dla internetu.
- [ ] Reverse proxy z HTTPS (Caddy/Nginx), aplikacja w kontenerze jako non-root.
- [ ] Zainstalować Claude Code na utwardzonym serwerze (tryb z pytaniem o zgodę).
- Zasada: ćwiczenia na danych sztucznych, nigdy na danych klienta.

## 6. Aplikacja rekrutacyjna (priorytet)
Przetwarza dane osobowe kandydatów, a AI w rekrutacji może być systemem
wysokiego ryzyka według EU AI Act. Przed prawdziwymi danymi:
- [ ] Zakres danych i podstawa prawna (RODO), DPA z dostawcą modelu, DPIA.
- [ ] Minimalizacja i maskowanie danych przed wysłaniem do modelu.
- [ ] Izolacja agentów, sekrety poza kodem, logi i audyt.
- [ ] Człowiek zatwierdza decyzje o kandydatach (nadzór ludzki).
- [ ] Przegląd architektury (wymaga dostępu do repozytorium tej aplikacji).

## 7. Przegląd aplikacji dostępnej w tej sesji (`analiza-fv`)
Aplikacja rekrutacyjna **nie** znajduje się w tym repozytorium; dostępne tu jest
tylko `analiza-fv`: jeden plik `index.py`, parser faktur (PDF/JPG/PNG → Excel, OCR).

Ustalenia (przegląd statyczny, bez uruchamiania):
- Brak kluczy, haseł i wywołań sieciowych. Plik się kompiluje (`py_compile` OK).
- Nie używa `eval`, `exec`, `subprocess`, `pickle`; dane tylko lokalnie. Niskie ryzyko.
- Przetwarza faktury (dane firm, ewentualnie osobowe): uruchamiaj na zaufanych plikach; pliki PDF/obrazy z nieznanych źródeł mogą wykorzystywać luki w bibliotekach (pdfplumber, PyMuPDF, Pillow, Tesseract), więc trzymaj je zaktualizowane.
- Wyszukuje pliki rekurencyjnie (`**/*`); wskazuj konkretny folder, nie cały katalog domowy.
- W nagłówku nazwa pliku to `parser_faktur_do_excela.py`, a plik nazywa się `index.py` (po zmianie nazwy). Dokumentacja jest nieaktualna.
- Brak `requirements.txt` z przypiętymi wersjami. Dodaj i skanuj (`pip-audit`).
- Wynik `.xlsx` zawiera dane z faktur: nie zostawiaj go w publicznym folderze.

## 8. Lista do sprawdzenia dla każdej aplikacji
- [ ] Porty: tylko 80/443 na zewnątrz, baza bez publicznego portu.
- [ ] HTTPS, uwierzytelnianie, 2FA, rate limiting.
- [ ] Kontener non-root, `read-only`, ograniczone zasoby.
- [ ] Sekrety poza repozytorium, rotacja kluczy.
- [ ] Zależności: przypięte wersje, skanowanie (`pip-audit`, `npm audit`, Trivy).
- [ ] Logi, monitoring, kopie z testem przywracania.
- [ ] Agenci AI: najmniejsze uprawnienia, brak sekretów, allowlista ruchu wychodzącego, zatwierdzanie ryzykownych akcji.
- [ ] Dane: klasyfikacja, minimalizacja, wiadomo, co trafia do modelu.

## 9. Co potrzebuję od Ciebie
1. Nazwa repozytorium aplikacji rekrutacyjnej (dodam je do sesji do przeglądu).
2. Wyniki komend audytu z pkt 5 (bez haseł i kluczy).
3. Ile godzin tygodniowo możesz się uczyć.
