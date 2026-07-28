# Dashboard vyvoja balenia

Jednoduchy report pre sledovanie vyvoja balenia z Excel vstupov.

Hlavne KPI:

```text
AB Eliminovane / Celkovy pocet
```

Ciel je 75 %.

## Kam nahrat data

Excel subory vkladaj do priecinka:

```text
input/
```

Skript berie najnovsi Excel subor s koncovkou `.xlsx`, `.xls` alebo `.xlsm`.
Ak je subor otvoreny v Exceli alebo zamknuty OneDrive syncom, skript si ho skopiruje do temp priecinka a cita kopiu.

## Co report zobrazuje

- hlavny podiel `AB Eliminovane / Celkovy pocet`
- rozdiel do ciela 75 %
- trend po dnoch
- filtre pre datum, geosize, stanicu balenia a baliacu skupinu
- filtr pre `detail dopravy` a jej podiel na celkovych SJLs
- top hodnoty v geosize, staniciach a skupinach

## Oczakavane stlpce

Skript sa snazi rozpoznat aj menej upratane Exceli.
Najlepsie funguje s poliami:

- datum
- geosize
- stanica balenia
- baliaca skupina
- detail dopravy
- AB Eliminovane
- Celkovy pocet

Ak Excel nie je v klasickej tabulke, ale ako pivot output, skript ho vie rozbalit do pouzitelneho datoveho formatu.

## Spustenie

Ak mas aktivne Python prostredie:

```powershell
pip install -r requirements.txt
python scripts/build_dashboard.py
```

## Vystup

Po spusteni vznikne:

```text
output/packaging_dashboard.html
```

Tento subor otvor v prehliadaci. Obsahuje uz vlozene data, takze nepotrebuje server.

Pre input `balikovka_den_CZLC4.xlsx` je k dispozicii aj specialny dashboard:

```text
output/balikovka_den_CZLC4_dashboard.html
```

Vygenerujes ho cez:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/build_balikovka_dashboard.ps1
```

## GitHub Pages

Ak chces, aby si dashboard vedel otvorit cez verejny link, pouzi GitHub Pages.

Hlavna vstupna stranka je teraz `Balení dashboard` a detailny dashboard `Vývoj balenia`
sa otvara ako rozbaľovací blok. Tato vrstva je pripravena na pridavanie dalsich kariet bez
zasahu do detailu.

Najjednoduchsie je publikovat obsah priecinka `docs/`, ktory uz teraz obsahuje samostatny `index.html`.

Kroky pre zaciatocnika:

1. V GitHube si vytvor novy repository.
2. Nahraj do neho tento projekt.
3. V GitHube otvor `Settings`.
4. V lavom menu klikni na `Pages`.
5. V casti `Build and deployment` zvol `Deploy from a branch`.
6. Ako branch vyber `main` a ako folder vyber `/docs`.
7. Uloz nastavenie a pockaj, kym GitHub Pages spravi deploy.
8. Po deployi dostanes link typu `https://tvoje-meno.github.io/nazov-repa/`.

V tomto repozitari je pre Pages pripravene:

- `docs/index.html` - nadradena vrstva `Balení dashboard`
- `docs/vyvoj-balenia.html` - stabilny snapshot detailu `Vývoj balenia`
- `docs/.nojekyll` - vypnutie Jekyll spracovania

Ak budes dashboard aktualizovat, po novom build-e len prepises `docs/vyvoj-balenia.html`
novym exportom z `output/packaging_dashboard.html`.

Na to sluzi aj skript:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/publish_to_docs.ps1
```

## Streamlit a n8n

Ak chces, aby sa data v Streamlit appke menili automaticky po kazdom update Excelu,
nenacitaj ich z `docs/` snapshotu. Namiesto toho nastav v Streamlit Cloud secrets:

```toml
EXCEL_SOURCE_URL = "https://.../tvoj-staly-excel.xlsx"
```

Potom n8n workflow spravi iba toto:

1. nahra novy Excel na tu istu stabilnu URL
2. alebo prepis ten isty subor v GitHub repozitari a pushni commit
3. Streamlit appka si pri otvoreni nacita aktualnu verziu z `EXCEL_SOURCE_URL`

Pre lokalne testovanie mozes dat:

```powershell
$env:EXCEL_SOURCE_URL="https://.../tvoj-staly-excel.xlsx"
streamlit run scripts/build_dashboard.py
```

Ak `EXCEL_SOURCE_URL` nie je nastavene, appka fallbackne pouzije lokalny snapshot z `docs/`.

## Doplnkovy skript

Povodny technicky prieskum Excelu je stale k dispozicii:

```powershell
python scripts/analyze_excel.py
```
