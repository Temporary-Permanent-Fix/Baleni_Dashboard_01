param(
    [string]$InputFile = ""
)

$ErrorActionPreference = "Stop"

$ProjectDir = Resolve-Path (Join-Path $PSScriptRoot "..")
$InputDir = Join-Path $ProjectDir "input"
$OutputDir = Join-Path $ProjectDir "output"
$OutputFile = Join-Path $OutputDir "balikovka_den_CZLC4_dashboard.html"

function Get-LatestExcelFile {
    param([string]$Folder)

    $files = Get-ChildItem -LiteralPath $Folder -File |
        Where-Object { $_.Extension -in ".xlsx", ".xls", ".xlsm" -and -not $_.Name.StartsWith("~$") } |
        Sort-Object LastWriteTime -Descending

    if (-not $files) {
        throw "V zlozke input nie je ziaden Excel subor."
    }

    $preferred = $files | Where-Object { $_.Name -match "balikovka_den_CZLC4" } | Select-Object -First 1
    if ($preferred) {
        return $preferred.FullName
    }

    return $files[0].FullName
}

function Copy-ExcelForReading {
    param([string]$Path)

    $copyPath = Join-Path $env:TEMP ("balikovka_dashboard_{0}.xlsx" -f ([IO.Path]::GetFileNameWithoutExtension($Path)))
    Copy-Item -LiteralPath $Path -Destination $copyPath -Force
    return $copyPath
}

function Get-SharedStrings {
    param(
        [System.IO.Compression.ZipArchive]$Zip
    )

    $entry = $Zip.GetEntry("xl/sharedStrings.xml")
    if (-not $entry) {
        return @()
    }

    $reader = New-Object System.IO.StreamReader($entry.Open())
    try {
        [xml]$xml = $reader.ReadToEnd()
    } finally {
        $reader.Close()
    }

    $values = New-Object System.Collections.Generic.List[string]
    foreach ($si in $xml.sst.si) {
        $values.Add($si.InnerText)
    }
    return $values
}

function Get-WorksheetXml {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$Path
    )

    $entry = $Zip.GetEntry($Path)
    if (-not $entry) {
        throw "Nenasiel som sheet XML: $Path"
    }

    $reader = New-Object System.IO.StreamReader($entry.Open())
    try {
        [xml]$xml = $reader.ReadToEnd()
    } finally {
        $reader.Close()
    }
    return $xml
}

function Get-CellValue {
    param(
        [System.Xml.XmlNode]$Row,
        [string]$Column,
        [System.Xml.XmlNamespaceManager]$Ns,
        [string[]]$SharedStrings
    )

    $rowNumber = $Row.GetAttribute("r")
    $ref = "$Column$rowNumber"
    $node = $Row.SelectSingleNode("x:c[@r='$ref']", $Ns)
    if (-not $node) {
        return ""
    }

    $type = $node.GetAttribute("t")
    $v = $node.SelectSingleNode("x:v", $Ns)
    $raw = if ($v) { $v.InnerText } else { "" }

    switch ($type) {
        "s" {
            if ($raw -eq "") { return "" }
            return $SharedStrings[[int]$raw]
        }
        "inlineStr" {
            return $node.InnerText
        }
        default {
            return $raw
        }
    }
}

function Round-ToInt {
    param([object]$Value)
    if ($null -eq $Value -or $Value -eq "") { return 0 }
    return [int][math]::Round([double]$Value, 0)
}

function Format-Int {
    param([int]$Value)
    return [string]::Format([System.Globalization.CultureInfo]::GetCultureInfo("sk-SK"), "{0:N0}", $Value)
}

function Format-Pct {
    param([double]$Value)
    return [string]::Format([System.Globalization.CultureInfo]::GetCultureInfo("sk-SK"), "{0:N1} %", ($Value * 100))
}

function Normalize-BoolText {
    param([object]$Value)

    $text = ([string]$Value).Trim().ToLowerInvariant()
    switch ($text) {
        "true" { return "true" }
        "1" { return "true" }
        "yes" { return "true" }
        "false" { return "false" }
        "0" { return "false" }
        "no" { return "false" }
        default { return "" }
    }
}

if ([string]::IsNullOrWhiteSpace($InputFile)) {
    $InputFile = Get-LatestExcelFile -Folder $InputDir
}

if (-not (Test-Path -LiteralPath $InputFile)) {
    throw "Excel subor neexistuje: $InputFile"
}

$copyPath = Copy-ExcelForReading -Path $InputFile

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($copyPath)
    try {
        $sharedStrings = Get-SharedStrings -Zip $zip
        $sheetXml = Get-WorksheetXml -Zip $zip -Path "xl/worksheets/sheet1.xml"
        $ns = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable)
        $ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

        $rows = $sheetXml.SelectNodes("//x:sheetData/x:row", $ns)
        $dataRows = @()

        foreach ($row in $rows) {
            if ($row.GetAttribute("r") -eq "1") {
                continue
            }

        $dataRows += [pscustomobject]@{
            RowNumber = [int]$row.GetAttribute("r")
            Transport = Get-CellValue -Row $row -Column "BM" -Ns $ns -SharedStrings $sharedStrings
            GeoSize   = Get-CellValue -Row $row -Column "X"  -Ns $ns -SharedStrings $sharedStrings
            Product   = Get-CellValue -Row $row -Column "V"  -Ns $ns -SharedStrings $sharedStrings
            Dobalovat = Normalize-BoolText (Get-CellValue -Row $row -Column "AX" -Ns $ns -SharedStrings $sharedStrings)
            ProductId = Get-CellValue -Row $row -Column "S"  -Ns $ns -SharedStrings $sharedStrings
        }
        }

        $totalRows = $dataRows.Count
        $geoSizes = @($dataRows | ForEach-Object { $_.GeoSize } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

        $transportSummary =
            $dataRows |
            Group-Object Transport |
            Sort-Object Count -Descending |
            ForEach-Object {
                [pscustomobject]@{
                    label = if ([string]::IsNullOrWhiteSpace($_.Name)) { "Nezadane" } else { $_.Name }
                    count = $_.Count
                    share = if ($totalRows) { $_.Count / $totalRows } else { 0 }
                }
            }

        $geoSummary =
            $dataRows |
            Group-Object GeoSize |
            Sort-Object Count -Descending |
            ForEach-Object {
                [pscustomobject]@{
                    label = if ([string]::IsNullOrWhiteSpace($_.Name)) { "Nezadane" } else { $_.Name }
                    count = $_.Count
                    share = if ($totalRows) { $_.Count / $totalRows } else { 0 }
                }
            }

        $productSummary =
            $dataRows |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.Product) } |
            Group-Object Product |
            Sort-Object @{ Expression = "Count"; Descending = $true }, @{ Expression = "Name"; Descending = $false } |
            Select-Object -First 10 |
            ForEach-Object {
                [pscustomobject]@{
                    label = $_.Name
                    count = $_.Count
                    share = if ($totalRows) { $_.Count / $totalRows } else { 0 }
                }
            }

        $payload = [ordered]@{
            metadata = [ordered]@{
                source_file = Split-Path $InputFile -Leaf
                generated_at = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                total_rows = $totalRows
                transport_types = $transportSummary.Count
                geo_sizes = $geoSummary.Count
                geo_size_values = @($geoSizes)
                top_products = $productSummary.Count
            }
            rows = @($dataRows)
            transport = @($transportSummary)
            geosize = @($geoSummary)
            products = @($productSummary)
        }

        $json = $payload | ConvertTo-Json -Depth 6 -Compress

        $html = @'
<!doctype html>
<html lang="sk">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Balikovka CZLC4 dashboard</title>
  <style>
    :root {
      --bg: #0f172a;
      --bg2: #111827;
      --panel: rgba(255,255,255,.92);
      --panel-strong: #ffffff;
      --ink: #0f172a;
      --muted: #64748b;
      --line: rgba(148,163,184,.28);
      --accent: #ffb000;
      --accent-2: #0f9f6e;
      --accent-3: #2563eb;
      --shadow: 0 18px 50px rgba(15, 23, 42, .18);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      color: var(--ink);
      font-family: "Segoe UI", Arial, sans-serif;
      background:
        radial-gradient(circle at top left, rgba(255,176,0,.18), transparent 32%),
        radial-gradient(circle at top right, rgba(37,99,235,.18), transparent 28%),
        linear-gradient(180deg, #f8fafc 0%, #eef2ff 100%);
    }
    header {
      position: relative;
      overflow: hidden;
      padding: 28px clamp(16px, 3vw, 36px);
      color: white;
      background:
        linear-gradient(135deg, #0f172a 0%, #172554 45%, #0f9f6e 100%);
      box-shadow: var(--shadow);
    }
    header::after {
      content: "";
      position: absolute;
      inset: auto -15% -42% auto;
      width: 320px;
      height: 320px;
      border-radius: 50%;
      background: radial-gradient(circle, rgba(255,255,255,.18), transparent 70%);
      pointer-events: none;
    }
    header h1 {
      margin: 0 0 8px;
      font-size: clamp(28px, 3.6vw, 48px);
      letter-spacing: -.03em;
    }
    header p {
      margin: 0;
      max-width: 70ch;
      color: rgba(255,255,255,.82);
      font-size: 15px;
      line-height: 1.5;
    }
    main {
      width: min(1450px, 100%);
      margin: 0 auto;
      padding: 22px clamp(12px, 2.5vw, 32px) 36px;
    }
    .meta {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 18px;
    }
    .pill {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 7px 12px;
      background: rgba(255,255,255,.72);
      color: var(--muted);
      font-size: 13px;
      backdrop-filter: blur(8px);
    }
    .cards {
      display: grid;
      grid-template-columns: repeat(4, minmax(180px, 1fr));
      gap: 14px;
      margin-bottom: 18px;
    }
    .card, .panel {
      border: 1px solid var(--line);
      border-radius: 18px;
      background: var(--panel);
      box-shadow: var(--shadow);
    }
    .card {
      padding: 18px 18px 16px;
      min-height: 118px;
    }
    .card span {
      display: inline-block;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
      letter-spacing: .08em;
      text-transform: uppercase;
    }
    .card strong {
      display: block;
      margin-top: 10px;
      font-size: clamp(28px, 3vw, 40px);
      line-height: 1;
    }
    .card small {
      display: block;
      margin-top: 8px;
      color: var(--muted);
      font-size: 13px;
    }
    .grid {
      display: grid;
      grid-template-columns: 1.1fr .9fr;
      gap: 16px;
      align-items: start;
    }
    .panel {
      padding: 18px;
    }
    .panel h2 {
      margin: 0 0 12px;
      font-size: 20px;
      letter-spacing: -.02em;
    }
    .panel p {
      margin: -2px 0 14px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.5;
    }
    .list {
      display: grid;
      gap: 10px;
    }
    .bar-row {
      display: grid;
      grid-template-columns: minmax(140px, 1.3fr) minmax(90px, .6fr) 3fr;
      gap: 12px;
      align-items: center;
    }
    .bar-row .label {
      font-weight: 600;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .bar-row .value {
      text-align: right;
      color: var(--muted);
      font-variant-numeric: tabular-nums;
    }
    .bar-track {
      height: 14px;
      border-radius: 999px;
      background: #e2e8f0;
      overflow: hidden;
    }
    .bar-fill {
      height: 100%;
      border-radius: inherit;
      background: linear-gradient(90deg, var(--accent-3), var(--accent));
      width: 0;
      transition: width .25s ease;
    }
    .bar-fill.geo {
      background: linear-gradient(90deg, var(--accent-2), #34d399);
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
    }
    th, td {
      padding: 11px 8px;
      border-bottom: 1px solid var(--line);
      text-align: right;
      vertical-align: top;
    }
    th:first-child, td:first-child { text-align: left; }
    th {
      color: var(--muted);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: .06em;
    }
    .table-wrap { overflow: auto; }
    .tag {
      display: inline-block;
      padding: 4px 8px;
      border-radius: 999px;
      background: rgba(37,99,235,.08);
      color: #1d4ed8;
      font-size: 12px;
      font-weight: 700;
    }
    .empty {
      padding: 24px;
      color: var(--muted);
      border: 1px dashed var(--line);
      border-radius: 14px;
      background: rgba(255,255,255,.7);
      text-align: center;
    }
    @media (max-width: 1080px) {
      .cards, .grid { grid-template-columns: 1fr 1fr; }
      .grid .panel:first-child { grid-column: 1 / -1; }
    }
    @media (max-width: 720px) {
      .cards, .grid, .bar-row { grid-template-columns: 1fr; }
      .bar-row .value { text-align: left; }
      .bar-row .label, .bar-row .value { white-space: normal; overflow: visible; text-overflow: clip; }
    }
  </style>
</head>
<body>
  <header>
    <h1>Balikovka CZLC4 dashboard</h1>
    <p>Prehlad z inputu <strong>balikovka_den_CZLC4.xlsx</strong>. Vidis tu pocet StoreJobLines podla typu dopravy, rozdelenie per geosize a top 10 opakujucich sa produktov.</p>
  </header>
  <main>
    <div class="meta" id="meta"></div>
    <section class="cards">
      <div class="card">
        <span>StoreJobLines</span>
        <strong id="totalRows">0</strong>
        <small>Celkovy pocet riadkov v inpute</small>
      </div>
      <div class="card">
        <span>Typy dopravy</span>
        <strong id="transportTypes">0</strong>
        <small>Unikatne hodnoty v stlpci <code>Detail dopravy</code></small>
      </div>
      <div class="card">
        <span>Geosize</span>
        <strong id="geoSizes">0</strong>
        <small>Rozdelenie cez <code>Geo Size produktu</code></small>
      </div>
      <div class="card">
        <span>Top produkt</span>
        <strong id="topProductCount">0</strong>
        <small id="topProductName">Najcastejsi produkt</small>
      </div>
    </section>
    <section class="grid">
      <div class="panel">
        <h2>Typ dopravy</h2>
        <p>Pocet StoreJobLines podla detailu dopravy.</p>
        <div class="list" id="transportBars"></div>
      </div>
      <div class="panel">
        <h2>Geosize</h2>
        <p>Rozdelenie StoreJobLines podla geosize produktu.</p>
        <div class="list" id="geoBars"></div>
      </div>
      <div class="panel" style="grid-column: 1 / -1;">
        <h2>Produkty podla geosize</h2>
        <p>Vyber si geosize. Predvolene je nastavene <strong>BPO</strong>, a produktový zoznam sa zoradi od najopakovanejsieho po najmenej opakovany.</p>
        <div class="meta" style="margin: 0 0 12px;">
          <label class="pill" style="gap:10px; cursor: default;">
            Geosize
            <select id="geoSelect" style="border:0; background:transparent; font:inherit; color:inherit; padding:0; outline:none;">
            </select>
          </label>
          <label class="pill" style="gap:10px; cursor: default;">
            Dobalovat
            <select id="dobalovatSelect" style="border:0; background:transparent; font:inherit; color:inherit; padding:0; outline:none;">
            </select>
          </label>
        </div>
        <div class="table-wrap">
          <table id="productTable"></table>
        </div>
      </div>
    </section>
  </main>
  <script id="payload" type="application/json">__PAYLOAD__</script>
  <script>
    const data = JSON.parse(document.getElementById('payload').textContent);
    const fmtInt = new Intl.NumberFormat('sk-SK', { maximumFractionDigits: 0 });
    const fmtPct = new Intl.NumberFormat('sk-SK', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
    const totalRows = Number(data.metadata?.total_rows || 0);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    const geoValues = ['BPO', ...((data.metadata?.geo_size_values || []).filter(value => value !== 'BPO'))];
    const geoSelect = document.getElementById('geoSelect');
    const dobalovatSelect = document.getElementById('dobalovatSelect');
    geoSelect.innerHTML = ['BPO', 'All', ...geoValues.filter(value => value !== 'BPO' && value !== 'All')].map(value => {
      const label = value === 'All' ? 'Všetky geosize' : value;
      const selected = value === 'BPO' ? ' selected' : '';
      return `<option value="${escapeHtml(value)}"${selected}>${escapeHtml(label)}</option>`;
    }).join('');
    geoSelect.value = 'BPO';
    dobalovatSelect.innerHTML = [
      { value: 'all', label: 'Všetko' },
      { value: 'true', label: 'true' },
      { value: 'false', label: 'false' },
      { value: 'empty', label: 'Nezadané' },
    ].map(option => `<option value="${escapeHtml(option.value)}">${escapeHtml(option.label)}</option>`).join('');
    dobalovatSelect.value = 'all';

    function escapeHtml(value) {
      return String(value ?? '').replace(/[&<>"']/g, char => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
      }[char]));
    }

    function renderBars(holderId, rows, fillClass) {
      const holder = document.getElementById(holderId);
      if (!rows.length) {
        holder.innerHTML = '<div class="empty">No data.</div>';
        return;
      }
      const max = Math.max(...rows.map(row => Number(row.count) || 0), 1);
      holder.innerHTML = rows.map(row => {
        const width = Math.max(2, (Number(row.count) || 0) / max * 100);
        return `
          <div class="bar-row">
            <div class="label" title="${escapeHtml(row.label)}">${escapeHtml(row.label)}</div>
            <div class="value">${fmtInt.format(Number(row.count) || 0)} (${fmtPct.format((Number(row.share) || 0) * 100)} %)</div>
            <div class="bar-track"><div class="bar-fill ${fillClass}" style="width:${width}%"></div></div>
          </div>
        `;
      }).join('');
    }

    function renderProducts(rows) {
      const table = document.getElementById('productTable');
      if (!rows.length) {
        table.innerHTML = '<tbody><tr><td class="empty">No repeated products found.</td></tr></tbody>';
        return;
      }
      table.innerHTML = `
        <thead>
          <tr>
            <th>#</th>
            <th>Produkt</th>
            <th>Dobalovat</th>
            <th>Pocet</th>
            <th>Podiel</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map((row, index) => `
            <tr>
              <td>${index + 1}</td>
              <td>${escapeHtml(row.label)}</td>
              <td>${escapeHtml(row.dobalovat || '')}</td>
              <td>${fmtInt.format(Number(row.count) || 0)}</td>
              <td><span class="tag">${fmtPct.format((Number(row.share) || 0) * 100)} %</span></td>
            </tr>
          `).join('')}
        </tbody>
      `;
    }

    function productsForGeosize(selectedGeo, selectedDobalovat) {
      const scopedRows = selectedGeo === 'All'
        ? rows
        : rows.filter(row => String(row.GeoSize || row.geosize || '').trim() === selectedGeo);
      const filteredRows = scopedRows.filter(row => {
        const dobalovat = String(row.Dobalovat || row.dobalovat || '').trim().toLowerCase();
        if (selectedDobalovat === 'all') return true;
        if (selectedDobalovat === 'empty') return !dobalovat;
        return dobalovat === selectedDobalovat;
      });
      const total = filteredRows.length || 0;
      const counts = new Map();
      for (const row of filteredRows) {
        const product = String(row.Product || row.product || '').trim();
        if (!product) continue;
        const current = counts.get(product) || { count: 0, values: new Set() };
        current.count += 1;
        const dobalovat = String(row.Dobalovat || row.dobalovat || '').trim().toLowerCase();
        if (dobalovat) current.values.add(dobalovat);
        counts.set(product, current);
      }
      return [...counts.entries()]
        .map(([label, item]) => {
          const dobalovatValues = [...item.values];
          const dobalovat = dobalovatValues.length === 1 ? dobalovatValues[0] : (dobalovatValues.length > 1 ? 'mixed' : '');
          return { label, count: item.count, share: total ? item.count / total : 0, dobalovat };
        })
        .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label, 'sk'));
    }

    document.getElementById('meta').innerHTML = [
      `Source: ${escapeHtml(data.metadata?.source_file || 'unknown')}`,
      `Generated: ${escapeHtml(data.metadata?.generated_at || '')}`
    ].map(text => `<span class="pill">${text}</span>`).join('');

    document.getElementById('totalRows').textContent = fmtInt.format(totalRows);
    document.getElementById('transportTypes').textContent = fmtInt.format(Number(data.metadata?.transport_types || 0));
    document.getElementById('geoSizes').textContent = fmtInt.format(Number(data.metadata?.geo_sizes || 0));
    if (data.products?.length) {
      document.getElementById('topProductCount').textContent = fmtInt.format(Number(data.products[0].count) || 0);
      document.getElementById('topProductName').textContent = data.products[0].label;
    } else {
      document.getElementById('topProductName').textContent = 'Bez opakovani';
    }

    renderBars('transportBars', data.transport || [], '');
    renderBars('geoBars', data.geosize || [], 'geo');
    renderProducts(productsForGeosize(geoSelect.value, dobalovatSelect.value));

    geoSelect.addEventListener('change', () => {
      renderProducts(productsForGeosize(geoSelect.value, dobalovatSelect.value));
    });

    dobalovatSelect.addEventListener('change', () => {
      renderProducts(productsForGeosize(geoSelect.value, dobalovatSelect.value));
    });
  </script>
</body>
</html>
'@ -replace "__PAYLOAD__", $json

        New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
        Set-Content -LiteralPath $OutputFile -Value $html -Encoding UTF8
    } finally {
        $zip.Dispose()
    }
} finally {
    Remove-Item -LiteralPath $copyPath -Force -ErrorAction SilentlyContinue
}

Write-Host "Hotovo."
Write-Host "Dashboard ulozeny v: $OutputFile"
