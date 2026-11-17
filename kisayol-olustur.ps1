# Masaüstüne "Yazgaç" kısayolu oluşturur (simgeli). Çift tıklayınca uygulama açılır.
$proje = $PSScriptRoot
$electron = Join-Path $proje "node_modules\electron\dist\electron.exe"
if (-not (Test-Path $electron)) {
  Write-Host "Electron bulunamadi. Once proje klasorunde 'npm install' calistirin." -ForegroundColor Yellow
  exit 1
}
$masaustu = [Environment]::GetFolderPath("Desktop")
$lnk = Join-Path $masaustu "Yazgaç.lnk"
$shell = New-Object -ComObject WScript.Shell
$sc = $shell.CreateShortcut($lnk)
$sc.TargetPath = $electron
$sc.Arguments = "."
$sc.WorkingDirectory = $proje
$sc.IconLocation = (Join-Path $proje "build\icon.ico") + ",0"
$sc.Description = "Yazgaç - Word ve Excel editörü"
$sc.Save()
Write-Host "Kisayol olusturuldu: $lnk"
