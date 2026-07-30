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

Skript berie najnovsi packaging Excel subor s koncovkou `.xlsx`, `.xls` alebo `.xlsm`.
Balikovka workbooky (`balikovka...`) su pre tento dashboard blokovane, aby sa nikdy omylom nedostali do produkcie.
Ak je subor otvoreny v Exceli alebo zamknuty OneDrive syncom, skript si ho skopiruje do temp priecinka a cita kopiu.

## Co report zobrazuje

- hlavny podiel `AB Eliminovane / Celkovy pocet`
- rozdiel do ciela 75 %
- trend po dnoch
- filtre pre datum, geosize, stanicu balenia a baliacu skupinu
- filtr pre `detail dopravy` a jej podiel na celkovych SJLs
- top hodnoty v geosize, staniciach a skupinach

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

## GitHub Pages

Ak chces verejny link, pouzi GitHub Pages.

Hlavna vstupna stranka je `Baleni dashboard` a detailny dashboard `Vyvoj balenia` sa otvara ako rozbalovaci blok. Tato vrstva je pripravena na pridavanie dalsich kariet bez zasahu do detailu.

Najjednoduchsie je publikovat obsah priecinka `docs/`, ktory uz teraz obsahuje samostatny `index.html`.

Kroky:

1. V GitHube si vytvor repository.
2. Nahraj do neho tento projekt.
3. V GitHube otvor `Settings`.
4. V lavom menu klikni na `Pages`.
5. V casti `Build and deployment` zvol `Deploy from a branch`.
6. Ako branch vyber `main` a ako folder vyber `/docs`.
7. Uloz nastavenie a pockaj na deploy.

V tomto repozitari je pre Pages pripravene:

- `docs/index.html` - nadradena vrstva `Baleni dashboard`
- `docs/vyvoj-balenia.html` - stabilny snapshot detailu `Vyvoj balenia`
- `docs/.nojekyll` - vypnutie Jekyll spracovania

Ak dashboard aktualizujes, po novom build-e len prepises `docs/vyvoj-balenia.html`
novym exportom z `output/packaging_dashboard.html`.

Na to sluzi aj skript:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/publish_to_docs.ps1
```

## Streamlit a n8n

Ak chces, aby sa data v Streamlit appke menili automaticky po kazdom update Excelu, mas dve moznosti:

1. `EXCEL_SOURCE_URL` - Streamlit si pri otvoreni natiahne zivy Excel zo stabilnej URL.
2. `docs/vyvoj-balenia.html` - Streamlit pouzije lokalny snapshot z repozitara.

Pre zivy zdroj nastav v Streamlit Cloud secrets:

```toml
EXCEL_SOURCE_URL = "https://.../tvoj-staly-excel.xlsx"
```

Ak `EXCEL_SOURCE_URL` nie je nastavene, appka fallbackne pouzije lokalny snapshot z `docs/`.

### Verzia 2

Verzia 2 je postavena tak, aby cloud n8n nemusel spustat lokalny PowerShell priamo:

1. n8n o 11:00 prepise trigger subor `n8n/refresh.request.json` v GitHube.
2. tvoj PC bezi s watcherom `scripts/watch_n8n_trigger.ps1`.
3. watcher stiahne zmenu, spusti `scripts/n8n_refresh_and_push.ps1` a pushne vysledok spat do GitHubu.
4. n8n si po krátkom waiti stiahne `output/daily_kpi.json` z GitHubu a cez Outlook node odošle mail s KPI za včerajšok.

Lokalny watcher spustis takto:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/watch_n8n_trigger.ps1
```

Ak chces iba jednorazovy test:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/watch_n8n_trigger.ps1 -RunOnce
```

Lokalny refresh/push skript zostava:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/n8n_refresh_and_push.ps1
```

Importovatelny n8n workflow pre verziu 2 je tu:

```text
n8n/daily_refresh_v2.json
```

Trigger subor, ktory n8n prepise, je:

```text
n8n/refresh.request.json
```

Jeho vzor je pripraveny aj ako:

```text
n8n/refresh.request.example.json
```

Ak chces workflow testovat rucne, staci zmenit hodnotu `requested_at` v `n8n/refresh.request.json`.

### Email notifikacia

Mail už posiela n8n workflow cez Microsoft Outlook node.

Ak chceš workflow testovať ručne, stačí spustiť `n8n/daily_refresh_v2.json` v n8n. Workflow:

1. zapíše trigger do GitHubu,
2. počká na lokálny refresh,
3. načíta `output/daily_kpi.json`,
4. odošle mail na `peter.kadlec@alza.sk`.

## Watcher na PC

Na tvojom PC je watcher pripraveny tu:

```text
scripts/watch_n8n_trigger.ps1
```

Ak ho chceš spúšťať automaticky po prihlásení do Windowsu, nainštaluj plánovanú úlohu:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/install_n8n_watcher_startup.ps1
```

Tento príkaz vytvorí task, ktorý spustí watcher skrytý na pozadí pri logone.

Watcher používa samostatný čistý clone v:

```text
%TEMP%\\baleni-dashboard-n8n-watcher\\repo
```

To znamená, že tvoje rozpracované zmeny v hlavnom repozitári mu nevadia.

Ak budeš chcieť task neskôr odstrániť:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/uninstall_n8n_watcher_startup.ps1
```

Ručné spustenie watcheru:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/watch_n8n_trigger.ps1
```
