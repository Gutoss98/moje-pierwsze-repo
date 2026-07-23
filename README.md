$folder = "D:\Projekty\Procesy-Funkcje-Systemy"

$systemFile  = Join-Path $folder "System.xlsx"
$processFile = Join-Path $folder "Proces.xlsx"
$functionFile = Join-Path $folder "Funkcja.xlsx"
$outputFile  = Join-Path $folder "Wynik.xlsx"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {

    ####################################################
    # Wczytanie System.xlsx
    ####################################################

    Write-Host "Czytam System.xlsx..."

    $wbSys = $excel.Workbooks.Open($systemFile)
    $wsSys = $wbSys.Worksheets.Item(1)

    $systems = @()

    $row = 2
    while($true){

        $name = $wsSys.Cells.Item($row,1).Text.Trim()

        if([string]::IsNullOrWhiteSpace($name)){
            break
        }

        $systems += $name
        $row++
    }

    $wbSys.Close($false)

    ####################################################
    # Wczytanie Proces.xlsx
    ####################################################

    Write-Host "Czytam Proces.xlsx..."

    $wbProc = $excel.Workbooks.Open($processFile)
    $wsProc = $wbProc.Worksheets.Item(1)

    $lastProcRow = $wsProc.UsedRange.Rows.Count
    $lastProcCol = $wsProc.UsedRange.Columns.Count

    ####################################################
    # Wczytanie Funkcja.xlsx do pamięci
    ####################################################

    Write-Host "Czytam Funkcja.xlsx..."

    $wbFun = $excel.Workbooks.Open($functionFile)
    $wsFun = $wbFun.Worksheets.Item(1)

    $lastFunRow = $wsFun.UsedRange.Rows.Count

    $functionMap = @{}

    for($r=2;$r -le $lastFunRow;$r++){

        $code = $wsFun.Cells.Item($r,4).Text.Trim()

        if($code -eq ""){
            continue
        }

        if(-not $functionMap.ContainsKey($code)){

            $functionMap[$code] = @{
                Meta = $wsFun.Cells.Item($r,1).Text.Trim()
                Func = $wsFun.Cells.Item($r,2).Text.Trim()
            }

        }

    }

    ####################################################
    # Wynik
    ####################################################

   ####################################################
# Wynik
####################################################

Write-Host "Tworzę Wynik.xlsx..."

$wbOut = $excel.Workbooks.Add()
$wsOut = $wbOut.Worksheets.Item(1)

$wsOut.Cells.Item(1,1)="System"
$wsOut.Cells.Item(1,2)="Procesy"
$wsOut.Cells.Item(1,3)="MetaFunkcje"
$wsOut.Cells.Item(1,4)="Funkcje"

$outRow = 2
$systemCounter = 1

foreach($system in $systems){

    Write-Host "[$systemCounter/$($systems.Count)] $system"

    $procesy = @()
    $metaFunkcje = @()
    $funkcje = @()

    for($col=2;$col -le $lastProcCol;$col++){

        $header = $wsProc.Cells.Item(1,$col).Text

        if($header.ToLower().Contains($system.ToLower())){

            for($r=2;$r -le $lastProcRow;$r++){

                $waga = $wsProc.Cells.Item($r,$col).Text.Trim()

                if([string]::IsNullOrWhiteSpace($waga)){
                    continue
                }

                $proces = $wsProc.Cells.Item($r,1).Text.Trim()

                # Wyciągnięcie kodu procesu
                if($proces -match '^([A-Za-z]+-\d+)'){
                    $kod = $matches[1]
                }
                else{
                    $kod = $proces
                }

                $meta = "BRAK"
                $funkcja = "BRAK"

                if($functionMap.ContainsKey($kod)){
                    $meta = $functionMap[$kod].Meta
                    $funkcja = $functionMap[$kod].Func
                }

                # Zamiana 0.25 -> 25%
                $wagaTekst = $waga
                try{
                    $wagaTekst = "{0:P0}" -f ([double]$waga.Replace(",","."))
                } catch {}

                $procesy += "$proces $wagaTekst"
                $metaFunkcje += $meta
                $funkcje += $funkcja
            }
        }
    }

    $wsOut.Cells.Item($outRow,1) = $system
    $wsOut.Cells.Item($outRow,2) = ($procesy -join "; ")
    $wsOut.Cells.Item($outRow,3) = ($metaFunkcje -join "; ")
    $wsOut.Cells.Item($outRow,4) = ($funkcje -join "; ")

    $outRow++
    $systemCounter++
}
    ####################################################
    # Formatowanie
    ####################################################

    $wsOut.Rows.Item(1).Font.Bold = $true
    $wsOut.UsedRange.Columns.AutoFit() | Out-Null

    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true

    $wbOut.SaveAs($outputFile)

    Write-Host ""
    Write-Host "==================================="
    Write-Host "Gotowe!"
    Write-Host "Wynik zapisano:"
    Write-Host $outputFile
    Write-Host "==================================="

}
finally{

    if($wbOut){$wbOut.Close($true)}
    if($wbFun){$wbFun.Close($false)}
    if($wbProc){$wbProc.Close($false)}

    $excel.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsSys) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsProc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsFun) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsOut) | Out-Null

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbSys) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbProc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbFun) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbOut) | Out-Null

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

}
