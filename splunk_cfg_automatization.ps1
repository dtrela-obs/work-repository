param( 
    [Parameter(Mandatory=$true)] [string]$ExcelPath,
    [Parameter(Mandatory=$true)] [string]$SheetName,
    [Parameter(Mandatory=$true)] [string]$ConfPath,
    [string]$IndexTimeMode = "current",
    [string]$Mode = "batch",
    [string]$Disabled = "0"

)

#Set-ExecutionPolicy -ExecutionPolicy Unrestricted - Scope Process -Force
$cfgFilePattern = "db_inputs_{dc}.conf"


function NormalizeString {
    param($s)
    if ($null -eq $s -or [string]::IsNullOrWhiteSpace($s)) { return ""}
    $val = $s.ToString()
    $val = $val -replace '[\x00-\x1F]', ''
    $val = $val.Replace([char]0x00A0, ' ')
    $val = ($val -replace "\s+", " ").Trim()
    return $val
}

function MakeSafeFileNamePart {
    param([string]$s)
    $safe = NormalizeString $s
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($ch in $invalid) { $safe = $safe.Replace($ch, "_")}
    $safe = $safe.Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) { return "Unnamed file. "}
    $limit = [Math]::Min($safe.Length, 64)
    return $safe.Substring(0, $limit)
    
}

function FormatSQLQuery {
    param([string]$sql)
    if ([string]::IsNullOrWhiteSpace($sql)) { return "" }

    $lines = @($sql -replace "(\r\n|\r)", "`n" -split "`n" | ForEach-Object { $_.TrimEnd() } | Where-Object { $_ })

    $res = @()
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = [string]$lines[$i]
        $isLast = ($i -eq ($lines.Count - 1))

        if ($islast) {
            if ($line -notmatch ";\s*$") { $line = "$line;" }
            if ($i -eq 0) { $res += $line } else { $res += " $line" }
        } else {
            $line = $line -replace ";\s*$", ""
            if ($i -eq 0) { $res += "$line \" } else { $res += " $line \" }
        }
    }
return  ($res -join "`r`n")
}

function NewStanzaBlock {
    param(
        $row,
        $IndexTimeMode,
        $Mode,
        $Disabled
    )

$stanza = ($row.StanzaName -replace "^\w\.\-:]+", "_").Trim()
$cron = NormalizeString $row.IntervalCRON
$desc = (NormalizeString $row.Description) -replace "(\r\n|\n|\r)", " "
$q = FormatSQLQuery $row.SQLQuery

$block = "[$stanza]`r`n"
$block += "connection = $($row.Connection)`r`n"
$block += "description = $desc`r`n"
$block += "disabled = $Disabled`r`n"
$block += "host = $($row.Host)`r`n"
$block += "index = $($row.Index)`r`n"
$block += "index_time_mode = $IndexTimeMode`r`n"
$block += "interval = $cron`r`n"
$block += "mode = $Mode`r`n"
$block += "query = $q`r`n"
$block += "sourcetype = $($row.Sourcetype)`r`n`r`n"

return $block
}

function SetConfAutoSection {
    param($Path, $GeneratedText)
    $begin = "# BEGIN GENERATED - DO NOT EDIT"
    $end = "# END GENERATED"
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Host "Configuration file not found: $Path"
        return
    }
    
    $content = [System.IO.File]::ReadAllText($Path)
    $startIdx = $content.IndexOf($begin)
    $stopIdx = $content.IndexOf($end)


    if ($StartIdx -lt 0 -or $stopIdx -lt 0) {
        Write-Host "Markers not found in configuration file: $Path, add the following lines to the file: $Path"
        return
    }
$before = $content.Substring(0, $StartIdx)
$after = $content.Substring($stopIdx + $end.Length)

$newContent = $before + $begin + "`r`n" + $GeneratedText + $end + $after
$utf8WithoutBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($Path, $newContent, $utf8WithoutBom)
}


 function Get-ExcelRows { 

param($Path, $Sheet)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null

try {
    $wb = $excel.Workbooks.Open($Path)
    $ws = $wb.Worksheets.Item($Sheet)
    $used = $ws.UsedRange
    

    $expected = @("stanzaname","sqlquery","intervalcron","index","connection","description","host","sourcetype","datacenter")
    $map = @{}

    for ($c = 1; $c -le $used.Columns.Count; $c++) {
        $val = [string]$ws.Cells.Item($used.Row, $c).Value2
        if (-not [string]::IsNullOrWhiteSpace($val)) {
            $clean = $val.ToLower().Replace("_","").Replace(" ","")
            if ($expected -contains $clean) { $map[$clean] = $c }
        }
    }

    foreach ($header in $expected) {
    if (-not $map.ContainsKey($header)) {
        throw "Missing required column in Excel: '$header'."
    }
    }

    $data = @()
    for ($r = 2; $r -le $used.Rows.Count; $r++) {
    $nameVal = [string]$ws.Cells.Item($r, $map["stanzaname"]).Value2
    if ([string]::IsNullOrWhiteSpace($nameVal)) { continue }

    $data += [PSCustomObject]@{
        StanzaName = $nameVal
        SQLQuery = [string]$ws.Cells.Item($r, $map["sqlquery"]).Value2
        IntervalCRON = [string]$ws.Cells.Item($r, $map["intervalcron"]).Value2
        Index = [string]$ws.Cells.Item($r, $map["index"]).Value2
        Connection = [string]$ws.Cells.Item($r, $map["connection"]).Value2
        Description = [string]$ws.Cells.Item($r, $map["description"]).Value2
        Host = [string]$ws.Cells.Item($r, $map["host"]).Value2
        Sourcetype = [string]$ws.Cells.Item($r, $map["sourcetype"]).Value2
        Datacenter = [string]$ws.Cells.Item($r, $map["datacenter"]).Value2
        RowNum = $r

     }
    }

    Write-Host "Loaded $($data.Count) data rows"
    return $data

}
finally {
    if ($wb) {$wb.Close($false)}
    $excel.Quit()
     [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
}



try {
    $rows = Get-ExcelRows -Path $ExcelPath -Sheet $SheetName
    if ($null -eq $rows -or $rows.Count -eq 0) {
        Write-Warning "No rows found to process"
        return
    }



    $confFolder = Split-Path -Path $ConfPath -Parent


    $groups = $rows | Group-Object { [string]$_.Datacenter.Trim().ToLower() }
    foreach ( $group in $groups ) {
    $dcName = $group.Group[0].Datacenter
    if ([string]::IsNullOrWhiteSpace($dcName)) { continue }

    $fileName = $cfgFilePattern.Replace("{dc}",(MakeSafeFileNamePart $dcName))
    $targetFile = Join-Path -Path $confFolder -ChildPath $fileName

    if (Test-Path -LiteralPath $targetFile ) {
        Write-Host "Processing DC: $dcName... "
        $generated = ""
        foreach ($row_item in $group.Group) {
            $generated += NewStanzaBlock -row $row_item -IndexTimeMode $IndexTimeMode -Mode $Mode -Disabled $Disabled
        }
        SetConfAutoSection -Path $targetFile -GeneratedText $generated
        Start-Process "notepad.exe" "`"$targetFile`""
    } else {
        Write-Warning "Skipped: $targetFile does not exists"
    }
}
} catch {
    Write-Error "Error: $($_.Exception.Message)"
}





























 