param(
  [string]$Name = "晴天的小红书文案助手",
  [string]$IconPath = "",
  [switch]$Clean
)

Write-Host "Building $Name with PyInstaller..."

if ($Clean) {
  Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
}

$addData = @(
  "稿定设计-数据上传模板.xlsx;.",
  "app/config/default_settings.json;app/config",
  "app/gui/style.qss;app/gui"
)

$dataArgs = $addData | ForEach-Object { "--add-data \"$_\"" } | Join-String -Separator " "

$iconArg = ""
if ($IconPath -and (Test-Path $IconPath)) { $iconArg = "--icon `"$IconPath`"" }

$cmd = "pyinstaller --noconfirm --windowed --name `"$Name`" $iconArg $dataArgs main.py"
Write-Host $cmd
Invoke-Expression $cmd

Write-Host "Done. Check dist/ folder."