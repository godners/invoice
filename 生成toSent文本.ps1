$toSent = "D:\Projects\invoice\toSent"
$BureauCode = $toSent + "\BureauCode.csv"
Set-Location -Path $toSent
$toSent = Get-Location
$price = [System.Decimal]0.00
$total = [System.Decimal]0.00
$oup = [System.String]""
$bc = Import-Csv -Path $BureauCode -Encoding UTF8
$fl = Get-ChildItem -Path $toSent -File | Where-Object -In -Property Extension -Value (".pdf", ".ofd")
$fl | Sort-Object -Property Name | ForEach-Object -Process {
    $bn = [System.String]$_.BaseName
    $idx = [System.Int16] $bn.Length - 1
    for ($i = 8; $i -lt $bn.Length; $i++) {
        if ($bn.Substring($i, 1) -notmatch "[0-9]") {$idx = [System.Math]::Min($idx, $i)}
    }
    $price = [System.Convert]::ToDecimal($bn.Substring(8, $idx -8))/100
    $total += $price
    $oup += "20" + $bn.Substring(0,2) + "-" + $bn.Substring(2,2) + "-" + $bn.Substring(4,2) + " （"
    $oup += ($bc | Where-Object -EQ -Property Code -Value $bn.Substring($idx)).Bureau + "） "
    $oup += $bn.Substring(6,2) + " ￥" +  ("{0:0.00}" -f $price) + "`r`n"
}
$oup = "电子发票（" + $fl.Length +"张，￥" + ("{0:0.00}" -f $total) + "）`r`n`r`n" + $oup
Write-Host -Object $oup
$key = Read-Host -Prompt "复制到剪贴板请输入[Y/y/1]"
if ($key -in ("Y", "y", "1")) {
    $oup | Set-Clipboard
    Write-Host -Object ("已复制到剪贴板，按回车退出……") -NoNewline
} else {
    Write-Host -Object ("取消复制，按回车退出……") -NoNewline
}
Read-Host
