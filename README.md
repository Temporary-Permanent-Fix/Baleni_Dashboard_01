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

## Doplnkovy skript

Povodny technicky prieskum Excelu je stale k dispozicii:

```powershell
python scripts/analyze_excel.py
```
