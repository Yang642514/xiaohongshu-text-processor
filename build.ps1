param(
  [string]$Name = "晴天的小红书文案转模板助手",
  [string]$IconPath = "app/static/logo/logo.ico",
  [switch]$Clean
)

Write-Host "Building $Name with PyInstaller..."

if ($Clean) {
  Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
}

$addData = @(
  "稿定设计-数据上传模板.xlsx;.",
  "app/config/default_settings.json;app/config",
  "app/gui/style.qss;app/gui",
  "app/static/logo/logo.jpg;app/static/logo",
  "app/static/logo/logo.png;app/static/logo",
  "app/static/logo/logo.ico;app/static/logo"
)

$dataArgs = $addData | ForEach-Object { "--add-data \"$_\"" } | Join-String -Separator " "

$iconArg = ""
if ($IconPath -and (Test-Path $IconPath)) { $iconArg = "--icon `"$IconPath`"" }

$cmd = "pyinstaller --noconfirm --windowed --name `"$Name`" $iconArg $dataArgs main.py"
# 如需避免控制台窗口，可改为：run.pyw
# $cmd = "pyinstaller --noconfirm --windowed --name `"$Name`" $iconArg $dataArgs run.pyw"
Write-Host $cmd
Invoke-Expression $cmd

Write-Host "Done. Check dist/ folder."