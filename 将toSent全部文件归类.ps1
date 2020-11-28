$BasePath = "D:\Projects\invoice"
$toSent = $BasePath + "\toSent"
$logcsv = $BasePath + "\归类记录.csv"
Set-Location -Path $BasePath
class clsLog {
    [System.String]$Date
    [System.String]$Time
    [System.String]$Path
    [System.String]$Name
}
$logs = New-Object -TypeName System.Collections.ArrayList
$start = Get-Date
$fl = Get-ChildItem -Path $toSent -File | Where-Object -In -Property Extension -Value (".pdf", ".ofd")
$fl | ForEach-Object -Process {
    $kemu = $BasePath + "\" + $_.BaseName.Substring(6,2)
    if (-not (Test-Path -Path $kemu)) {New-Item -Path $kemu -ItemType Directory | Out-Null}
    $km = Get-Item -Path $kemu
    Move-Item -Path $_.FullName -Destination $km.FullName
    $log = New-Object -TypeName clsLog
    $log.Date = Get-Date -Format "yyyy-MM-dd"
    $log.Time = Get-Date -Format "HH:mm:ss"
    $log.Path = $km.FullName
    $log.Name = $_.Name
    $logs.Add($log) | Out-Null
}
$logs | Export-Csv -Path $logcsv -Append -Encoding UTF8 -NoTypeInformation
$logs | Format-Table -AutoSize  -HideTableHeaders
$diff= (Get-Date) - $start
Write-Host -Object ("归类" + $fl.Length + "个文件，耗时" + $diff.Milliseconds.ToString() + "毫秒。")
Write-Host -Object ("按回车退出……") -NoNewline
Read-Host
