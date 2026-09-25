$excelPath = (Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe").'(default)'
$xllPath = (Resolve-Path "bin\LeeExcel64.xll").Path

Write-Host "正在启动 Excel 并加载 LeeExcel AI 原生插件..." -ForegroundColor Cyan
Write-Host "Excel 路径: $excelPath"
Write-Host "插件路径: $xllPath"

Start-Process -FilePath $excelPath -ArgumentList "`"$xllPath`""
