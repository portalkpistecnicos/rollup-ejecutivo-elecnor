param(
  [switch]$NoPush
)
$ErrorActionPreference = "Stop"
$RepoDir = "G:\Mi unidad\Respaldo\Documents\KPI's\Rollup Ejecutivo"
$LogFile = Join-Path $RepoDir "Log_Actualizacion.txt"
function Log($msg) {
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg"
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

function Norm-Agencia($raw) {
    if (-not $raw) { return $null }
    $s = "$raw".ToUpper()
    if ($s -match "SAN ANTONIO") { return "San Antonio" }
    if ($s -match "VALPARA") { return "Valparaíso" }
    if ($s -match "VINA DEL MAR" -or $s -match "VIÑA DEL MAR" -or $s -match "VIï¿½A DEL MAR") { return "Viña del Mar" }
    return $null
}

try {

Log "=== Inicio actualizacion Rollup Ejecutivo ==="
$hoyStr = Get-Date -Format "dd-MM-yyyy"

# ---------- 1) AI% desde Informe_Gerencial_Infancia.html ----------
$infanciaHtml = Get-Content "G:\Mi unidad\Respaldo\Documents\KPI's\Averías Infancia\Informe_Gerencial_Infancia.html" -Raw -Encoding UTF8
$aiPct = @{}
$rowMatches = [regex]::Matches($infanciaHtml, '<td>(San Antonio|Valparaíso|Viña del Mar)</td>\s*<td>[^<]*</td>\s*<td>[^<]*</td>\s*<td>([\d,]+)\s*%</td>')
foreach ($m in $rowMatches) {
    $ag = $m.Groups[1].Value
    $pct = [double]($m.Groups[2].Value -replace ',', '.')
    $aiPct[$ag] = $pct / 100.0
}
Log "AI%: $($aiPct.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, $_.Value } | Out-String)"

# ---------- 2) RR% desde historico_p23_averias_reiteradas.csv (fecha mas reciente) ----------
$p23 = Import-Csv "G:\Mi unidad\Respaldo\Documents\KPI's\Averías Reiteradas\Historico\historico_p23_averias_reiteradas.csv" -Encoding UTF8
function ParseDMY($s) { try { [datetime]::ParseExact($s, "dd-MM-yyyy", $null) } catch { [datetime]::MinValue } }
$maxFechaP23 = ($p23 | ForEach-Object { ParseDMY $_.FechaActualizacion } | Measure-Object -Maximum).Maximum
$p23Latest = $p23 | Where-Object { (ParseDMY $_.FechaActualizacion) -eq $maxFechaP23 }
$rrPct = @{}
foreach ($row in $p23Latest) {
    $ag = Norm-Agencia $row.Agencia
    if ($ag -and $row.TasaReiteracionPct) {
        $rrPct[$ag] = [double]($row.TasaReiteracionPct -replace ',', '.') / 100.0
    }
}
Log "RR% (fecha $($maxFechaP23.ToString('dd-MM-yyyy'))): $($rrPct.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, $_.Value } | Out-String)"

# ---------- 3) PCITA% desde Primera Cita gh-pages-repo/index.html ----------
$pcitaHtml = Get-Content "G:\Mi unidad\Respaldo\Documents\KPI's\Primera Cita\gh-pages-repo\index.html" -Raw -Encoding UTF8
$pcitaPct = @{}
$cardMatches = [regex]::Matches($pcitaHtml, '<h2>(San Antonio|Valparaíso|Viña del Mar)</h2>.*?card__figure">([\d,]+)<small>%', [System.Text.RegularExpressions.RegexOptions]::Singleline)
foreach ($m in $cardMatches) {
    $ag = $m.Groups[1].Value
    $pcitaPct[$ag] = [double]($m.Groups[2].Value -replace ',', '.') / 100.0
}
Log "PCITA%: $($pcitaPct.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, $_.Value } | Out-String)"

# ---------- 4) VEL_1D% desde historico_p31_velocidad_reparaciones.csv (fecha mas reciente) ----------
$p31 = Import-Csv "G:\Mi unidad\Respaldo\Documents\KPI's\Velocidad Reparaciones\BBDD\historico_p31_velocidad_reparaciones.csv" -Encoding UTF8
$maxFechaP31 = ($p31 | ForEach-Object { ParseDMY $_.FechaActualizacion } | Measure-Object -Maximum).Maximum
$p31Latest = $p31 | Where-Object { (ParseDMY $_.FechaActualizacion) -eq $maxFechaP31 }
$velPct = @{}
foreach ($row in $p31Latest) {
    $ag = Norm-Agencia $row.Agencia
    if ($ag -and $row.PctCumplimiento) {
        $velPct[$ag] = [double]($row.PctCumplimiento -replace ',', '.') / 100.0
    }
}
Log "VEL_1D% (fecha $($maxFechaP31.ToString('dd-MM-yyyy'))): $($velPct.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, $_.Value } | Out-String)"

# ---------- 5) Plantel_Elecnor: Rut -> Agencia ----------
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wbP = $excel.Workbooks.Open("G:\Mi unidad\Respaldo\Documents\KPI's\Producción\Plantel_Elecnor.xlsx", 0, $true)
$wsP = $wbP.Worksheets.Item("Plantel")
$lastRowP = $wsP.Cells.Item($wsP.Rows.Count,2).End(-4162).Row
$valsP = $wsP.Range($wsP.Cells.Item(1,1), $wsP.Cells.Item($lastRowP,6)).Value2
$rutToAgencia = @{}
for ($r=4; $r -le $lastRowP; $r++) {
    $rut = $valsP[$r,2]
    $agRaw = $valsP[$r,6]
    if ($rut -and $agRaw) {
        $ag = Norm-Agencia $agRaw
        if ($ag) { $rutToAgencia["$rut"] = $ag }
    }
}
$wbP.Close($false)
Log "Plantel: $($rutToAgencia.Count) ruts mapeados"

# ---------- 6) Producción.xlsx BBDD: Baremo_Total, MODELO, Rut_Tac ----------
$wbProd = $excel.Workbooks.Open("G:\Mi unidad\Respaldo\Documents\KPI's\Producción\Producción.xlsx", 0, $true)
$wsProd = $wbProd.Worksheets.Item("BBDD")
$lastRowProd = $wsProd.Cells.Item($wsProd.Rows.Count,1).End(-4162).Row
$valsProd = $wsProd.Range($wsProd.Cells.Item(1,1), $wsProd.Cells.Item($lastRowProd,25)).Value2
$agBaremo = @{ "San Antonio"=0.0; "Valparaíso"=0.0; "Viña del Mar"=0.0 }
$agBaremoInstala = @{ "San Antonio"=0.0; "Valparaíso"=0.0; "Viña del Mar"=0.0 }
$agBaremoReparaPxQ = @{ "San Antonio"=0.0; "Valparaíso"=0.0; "Viña del Mar"=0.0 }
$agPersonal = @{ "San Antonio"=@{}; "Valparaíso"=@{}; "Viña del Mar"=@{} }
for ($r=2; $r -le $lastRowProd; $r++) {
    $rut = $valsProd[$r,1]
    if (-not $rut) { continue }
    $ag = $rutToAgencia["$rut"]
    if (-not $ag) { continue }
    $baremo = $valsProd[$r,22]
    if ($baremo -eq $null) { $baremo = 0 }
    $modelo = $valsProd[$r,11]
    $agBaremo[$ag] += [double]$baremo
    $agPersonal[$ag]["$rut"] = $true
    if ($modelo -eq "INSTALA CUFH") { $agBaremoInstala[$ag] += [double]$baremo }
    elseif ($modelo -eq "REPARA PXQ CUFH") { $agBaremoReparaPxQ[$ag] += [double]$baremo }
}
$wbProd.Close($false)
Log "Producción: $lastRowProd filas procesadas"

# ---------- 7) Resumen Quiebre.xlsx: per-tecnico -> promedio por agencia ----------
$wbQ = $excel.Workbooks.Open("G:\Mi unidad\Respaldo\Documents\KPI's\Producción\Resumen Quiebre.xlsx", 0, $true)
$wsQ = $wbQ.Worksheets.Item("Resumen Quiebre")
$lastRowQ = $wsQ.Cells.Item($wsQ.Rows.Count,1).End(-4162).Row
$valsQ = $wsQ.Range($wsQ.Cells.Item(1,1), $wsQ.Cells.Item($lastRowQ,5)).Value2
$agQuiebreTot = @{ "San Antonio"=@(); "Valparaíso"=@(); "Viña del Mar"=@() }
for ($r=2; $r -le $lastRowQ; $r++) {
    $rut = $valsQ[$r,1]
    if (-not $rut) { continue }
    $ag = $rutToAgencia["$rut"]
    if (-not $ag) { continue }
    $tot = $valsQ[$r,5]
    if ($tot -ne $null) { $agQuiebreTot[$ag] += [double]$tot }
}
$wbQ.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
function Avg($arr) { if ($arr.Count -eq 0) { return 0 } else { return ($arr | Measure-Object -Average).Average } }

# ---------- 8) Backlog Instala (DATA_START/DATA_END blob) ----------
$biHtml = Get-Content "G:\Mi unidad\Respaldo\Documents\KPI's\Backlog Instala\index.html" -Raw -Encoding UTF8
$biJson = $null
if ($biHtml -match '/\*DATA_START\*/(.*?)/\*DATA_END\*/') { $biJson = $matches[1] | ConvertFrom-Json }
$backlogInstala = @{}
foreach ($a in $biJson.agencies) {
    $ratio = if (($a.cancelada + $a.terminada) -gt 0) { $a.enProceso / (($a.cancelada + $a.terminada)/6.0) } else { 0 }
    $backlogInstala[$a.name] = [math]::Round($ratio,3)
}
Log "Backlog Instala: $($backlogInstala.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, $_.Value } | Out-String)"

# ---------- 9) Backlog Repara (p67 CSV crudo, ultimo partition_date) ----------
$p67Files = Get-ChildItem "G:\Mi unidad\Respaldo\Documents\KPI's\Backlog Repara\p67_base_backlog_reparaciones_mod-detalle_*.csv" | Sort-Object LastWriteTime -Descending
$p67 = Import-Csv -Path $p67Files[0].FullName -Delimiter ";" -Encoding UTF8
$maxDateP67 = ($p67 | ForEach-Object { ParseDMY $_.partition_date } | Measure-Object -Maximum).Maximum
$maxDateStrP67 = $maxDateP67.ToString("dd-MM-yyyy")
$p67Latest = $p67 | Where-Object { $_.partition_date -eq $maxDateStrP67 }
$backlogRepara = @{}
foreach ($agName in @("San Antonio","Valparaíso","Viña del Mar")) {
    $rows = $p67Latest | Where-Object { (Norm-Agencia $_.rdy_cod_territorio) -eq $agName }
    $enProceso = ($rows | Where-Object { $_.rdy_estado -match "(?i)pendiente" }).Count
    $cancelado = ($rows | Where-Object { $_.rdy_estado -match "(?i)cancelad" }).Count
    $cerrado   = ($rows | Where-Object { $_.rdy_estado -match "(?i)cerrado" }).Count
    $denom = ($cancelado + $cerrado) / 6.0
    $ratio = if ($denom -gt 0) { $enProceso / $denom } else { 0 }
    $backlogRepara[$agName] = [math]::Round($ratio,3)
}
Log "Backlog Repara (fuente $($p67Files[0].Name), fecha $maxDateStrP67): $($backlogRepara.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, $_.Value } | Out-String)"

# ---------- 10) FM_* piecewise (decoded de la DAX del pbix) + Factor Desempeño (pesos custom 15/30/15/40) ----------
function Scale($x, $goodThresh, $badThresh, $goodIsLow) {
    if ($goodIsLow) {
        if ($x -lt $goodThresh) { return 1.1 }
        if ($x -gt $badThresh) { return 0.9 }
        return ((1.1-0.9)*($x-$badThresh))/($goodThresh-$badThresh)+0.9
    } else {
        if ($x -lt $badThresh) { return 0.9 }
        if ($x -gt $goodThresh) { return 1.1 }
        return ((1.1-0.9)*($x-$badThresh))/($goodThresh-$badThresh)+0.9
    }
}
$aiThresh = @{ "San Antonio"=@(0.0108,0.0252); "Valparaíso"=@(0.0108,0.0252); "Viña del Mar"=@(0.0108,0.0252) }
$rrThresh = @{ "San Antonio"=@(0.0227,0.0471); "Valparaíso"=@(0.0221,0.0459); "Viña del Mar"=@(0.0183,0.0379) }
$pcitaThresh = @{ "San Antonio"=@(0.7056,0.7798); "Valparaíso"=@(0.6650,0.7350); "Viña del Mar"=@(0.6318,0.6983) }
$velThresh = @{ "San Antonio"=@(0.76,0.84); "Valparaíso"=@(0.76,0.84); "Viña del Mar"=@(0.76,0.84) }

$wAI = 0.15; $wRR = 0.30; $wPCITA = 0.15; $wVEL = 0.40
$valorPB = 23632.53 * 1.0065

$mesesEs = @{1="Enero";2="Febrero";3="Marzo";4="Abril";5="Mayo";6="Junio";7="Julio";8="Agosto";9="Septiembre";10="Octubre";11="Noviembre";12="Diciembre"}
$periodo = "$($mesesEs[(Get-Date).Month]) $((Get-Date).Year)"

$agencies = @()
$historicoRows = @()
foreach ($ag in @("San Antonio","Valparaíso","Viña del Mar")) {
    $ai = $aiPct[$ag]; $rr = $rrPct[$ag]; $pcita = $pcitaPct[$ag]; $vel = $velPct[$ag]
    if ($ai -eq $null -or $rr -eq $null -or $pcita -eq $null -or $vel -eq $null) {
        Log "ADVERTENCIA: faltan datos de entrada para $ag (AI=$ai RR=$rr PCITA=$pcita VEL=$vel) - se omite del reporte"
        continue
    }
    $fmAI    = Scale $ai    $aiThresh[$ag][0]    $aiThresh[$ag][1]    $true
    $fmRR    = Scale $rr    $rrThresh[$ag][0]    $rrThresh[$ag][1]    $true
    $fmPCITA = Scale $pcita $pcitaThresh[$ag][1] $pcitaThresh[$ag][0] $false
    $fmVEL   = Scale $vel   $velThresh[$ag][1]   $velThresh[$ag][0]   $false
    $factorDesempeno = $wAI*$fmAI + $wRR*$fmRR + $wPCITA*$fmPCITA + $wVEL*$fmVEL

    $importePB = ($agBaremoInstala[$ag] + $agBaremoReparaPxQ[$ag]) * $valorPB * $factorDesempeno

    $obj = [PSCustomObject]@{
        Agencia = $ag
        AI_pct = [math]::Round($ai*100,2)
        RR_pct = [math]::Round($rr*100,2)
        PCITA_pct = [math]::Round($pcita*100,2)
        VEL_pct = [math]::Round($vel*100,2)
        FM_AI = [math]::Round($fmAI,4)
        FM_RR = [math]::Round($fmRR,4)
        FM_PCITA = [math]::Round($fmPCITA,4)
        FM_VEL = [math]::Round($fmVEL,4)
        FactorDesempeno = [math]::Round($factorDesempeno,4)
        Personal = $agPersonal[$ag].Count
        Peticiones = 0
        BaremoTotal = [math]::Round($agBaremo[$ag],1)
        ImportePB = [math]::Round($importePB,0)
        QuiebreComercial_pct = 0
        QuiebreTecnico_pct = 0
        QuiebreTotal_pct = [math]::Round((Avg $agQuiebreTot[$ag])*100,2)
        BacklogInstala = [PSCustomObject]@{ Ratio = $backlogInstala[$ag] }
        BacklogRepara = [PSCustomObject]@{ Ratio = $backlogRepara[$ag] }
    }
    $agencies += $obj

    $historicoRows += [PSCustomObject]@{
        FechaActualizacion = $hoyStr
        Agencia = $ag
        AI_pct = $obj.AI_pct
        RR_pct = $obj.RR_pct
        PCITA_pct = $obj.PCITA_pct
        VEL_pct = $obj.VEL_pct
        FactorDesempeno = $obj.FactorDesempeno
        ImportePB = $obj.ImportePB
        Personal = $obj.Personal
        QuiebreTotal_pct = $obj.QuiebreTotal_pct
        BacklogInstala_dias = $backlogInstala[$ag]
        BacklogRepara_dias = $backlogRepara[$ag]
    }
}

if ($agencies.Count -eq 0) { throw "No se pudo calcular ninguna agencia - abortando sin tocar index.html" }

$meta = [PSCustomObject]@{
    generatedAt = "$hoyStr $(Get-Date -Format 'HH:mm')"
    periodo = $periodo
    agencies = $agencies
}
$newJson = ($meta | ConvertTo-Json -Depth 6 -Compress)

# ---------- 11) Inyectar en index.html ----------
$htmlPath = Join-Path $RepoDir "index.html"
$html = Get-Content $htmlPath -Raw -Encoding UTF8
$pattern = '(?s)(<script id="rollup-data" type="application/json">).*?(</script>)'
$html = [regex]::Replace($html, $pattern, { param($m) $m.Groups[1].Value + $newJson.Trim() + $m.Groups[2].Value })
Set-Content -Path $htmlPath -Value $html -Encoding UTF8
Log "index.html actualizado"

# ---------- 12) Historico (idempotente por fecha) ----------
$histPath = Join-Path $RepoDir "Historico_Rollup_Ejecutivo.csv"
$existing = @()
if (Test-Path $histPath) {
    $existing = Import-Csv $histPath -Encoding UTF8 | Where-Object { $_.FechaActualizacion -ne $hoyStr }
}
$allRows = @($existing) + @($historicoRows)
$allRows | Export-Csv -Path $histPath -NoTypeInformation -Encoding UTF8
Log "Historico actualizado: $($allRows.Count) filas totales ($($historicoRows.Count) de hoy)"

# ---------- 13) Copia fechada (snapshot) ----------
$histDir = Join-Path $RepoDir "historico\html"
New-Item -ItemType Directory -Path $histDir -Force | Out-Null
Copy-Item $htmlPath -Destination (Join-Path $histDir "$hoyStr.html") -Force

# ---------- 14) Git commit + push ----------
$prevEAP = $ErrorActionPreference
$ErrorActionPreference = "Continue"
Set-Location $RepoDir
git add index.html Historico_Rollup_Ejecutivo.csv "historico/html/$hoyStr.html" *> $null
$changes = git status --porcelain
if ($changes) {
    git -c user.email="xarancibia@local" -c user.name="xarancibia" commit -m "Actualizacion diaria $hoyStr" *> $null
    Log "Commit creado"
    if (-not $NoPush) {
        $pushResult = git push origin main 2>&1 | Out-String
        Log "Push: $($pushResult.Trim())"
    }
} else {
    Log "Sin cambios que commitear"
}
$ErrorActionPreference = $prevEAP

Log "=== Fin actualizacion Rollup Ejecutivo (OK) ==="

} catch {
    Log "ERROR: $($_.Exception.Message)"
    throw
}
