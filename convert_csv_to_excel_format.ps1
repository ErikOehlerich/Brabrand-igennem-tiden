# PowerShell script til at konvertere CSV-filer fra $ separator til Excel (.xlsx) filer
# Dette opretter rigtige Excel-filer med kolonner som du kan redigere

# Tjek om ImportExcel modulet er installeret
if (!(Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installerer ImportExcel modul..."
    try {
        Install-Module -Name ImportExcel -Force -Scope CurrentUser
        Write-Host "ImportExcel modul installeret succesfuldt"
    }
    catch {
        Write-Host "Kunne ikke installere ImportExcel modul. Prøver alternativ metode..."
    }
}

# Definer input og output mapper
$inputFolder = ".\Folketællinger"
$outputFolder = ".\Folketællinger_Excel"

# Opret output mappe hvis den ikke eksisterer
if (!(Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder
    Write-Host "Oprettet mappe: $outputFolder"
}

# Find alle CSV-filer i input mappen
$csvFiles = Get-ChildItem -Path $inputFolder -Filter "*.csv"

foreach ($file in $csvFiles) {
    Write-Host "Behandler: $($file.Name)"
    
    try {
        # Læs filen
        $content = Get-Content $file.FullName -Raw -Encoding UTF8
        
        # Parse CSV med $ som separator
        $lines = $content -split "`r?`n" | Where-Object { $_.Trim() -ne "" }
        $headers = $lines[0] -split '\$'
        
        # Opretter PowerShell objekter for hver række
        $data = @()
        for ($i = 1; $i < $lines.Count; $i++) {
            $values = $lines[$i] -split '\$'
            $obj = New-Object PSObject
            
            for ($j = 0; $j < $headers.Count; $j++) {
                if ($j -lt $values.Count) {
                    $obj | Add-Member -MemberType NoteProperty -Name $headers[$j] -Value $values[$j]
                } else {
                    $obj | Add-Member -MemberType NoteProperty -Name $headers[$j] -Value ""
                }
            }
            $data += $obj
        }
        
        # Opret output fil navn (ændrer .csv til .xlsx)
        $outputFile = Join-Path $outputFolder ($file.BaseName + ".xlsx")
        
        # Prøv at eksportere til Excel med ImportExcel modul
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module ImportExcel
            $data | Export-Excel -Path $outputFile -AutoSize -BoldTopRow -FreezeTopRow -TableStyle Medium2
            Write-Host "Konverteret til Excel: $outputFile"
        } else {
            # Fallback: eksporter til CSV med komma som separator
            $csvOutput = Join-Path $outputFolder ($file.BaseName + "_Excel_format.csv")
            $data | Export-Csv -Path $csvOutput -NoTypeInformation -Delimiter "," -Encoding UTF8
            Write-Host "Konverteret til CSV (komma-separeret): $csvOutput"
            Write-Host "Denne fil kan åbnes i Excel og gemmes som .xlsx"
        }
        
    } catch {
        Write-Host "Fejl ved behandling af $($file.Name): $($_.Exception.Message)"
    }
}

Write-Host ""
Write-Host "Konvertering færdig!"
Write-Host "Alle filer er gemt i mappen: $outputFolder"
Write-Host ""
Write-Host "Du kan nu:"
Write-Host "1. Åbne Excel-filerne (.xlsx) direkte"
Write-Host "2. Tilføje en 'Koordinater' kolonne med lat,lng værdier"
Write-Host "3. Eller tilføje separate 'Lat' og 'Lng' kolonner"
Write-Host ""
Write-Host "Originale filer er bevaret i: $inputFolder"