# 做「CT 自動合併」免安裝版：python.org 官方的 embeddable Python + ct_auto.py + ct_merge.py
#
# 為什麼不用 exe：之前分院同事電腦上，Nuitka 編的 exe 都被防毒擋掉（沒簽章的新 exe）。
# python.org 官方的 python.exe / pythonw.exe 有 Python Software Foundation 的數位簽章，
# 程式本身是純文字的 .py 檔。（跟 ris-schedule-tool 的 build_portable.ps1 同一個做法）
#
# 產出：portable_build\CT自動合併\  整個資料夾壓成 zip，拿到 CT 檢查室的電腦解壓縮即可。
#   CT自動合併.bat           雙擊啟動
#   建立桌面捷徑.bat         在那台電腦的桌面建捷徑（捷徑要絕對路徑，只能到現場建）
#   ct_auto.py、ct_merge.py  程式，改版時只要換這兩個
#   ct_merge_no_merge.txt    不用合併的醫令（預設 *(HMC)）
#   python\                  免安裝 Python（含 tkinter 和 pywinauto 等套件）
#
# 執行： powershell -ExecutionPolicy Bypass -File build_portable.ps1

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here
$name = "CT自動合併"

# 版本要跟本機 Python 一樣，因為 tkinter 是從本機那一份複製過去的
$ver = (python -c "import platform; print(platform.python_version())").Trim()
$localPy = (python -c "import sys; print(sys.base_prefix)").Trim()
$tag = "python" + ($ver.Split(".")[0..1] -join "")          # 例：python312

$out = Join-Path $here "portable_build\$name"
$py = Join-Path $out "python"
$cache = Join-Path $here "portable_build\_cache"
New-Item -ItemType Directory -Force $cache | Out-Null
if (Test-Path $out) { Remove-Item -Recurse -Force $out }
New-Item -ItemType Directory -Force $py | Out-Null

# 1. 官方 embeddable Python
$zipName = "python-$ver-embed-amd64.zip"
$zip = Join-Path $cache $zipName
if (-not (Test-Path $zip)) {
    Write-Host "下載 $zipName ..."
    Invoke-WebRequest "https://www.python.org/ftp/python/$ver/$zipName" -OutFile $zip
}
Expand-Archive $zip -DestinationPath $py

# 2. tkinter（embeddable 版沒有附，視窗需要）
Copy-Item -Recurse (Join-Path $localPy "Lib\tkinter") (Join-Path $py "tkinter")
# zlib1.dll 是 tcl86t.dll 要的，漏了會「DLL load failed while importing _tkinter」
foreach ($f in "_tkinter.pyd", "tcl86t.dll", "tk86t.dll", "zlib1.dll") {
    Copy-Item (Join-Path $localPy "DLLs\$f") $py
}
New-Item -ItemType Directory -Force (Join-Path $py "tcl") | Out-Null
foreach ($d in "tcl8.6", "tk8.6", "tcl8") {
    Copy-Item -Recurse (Join-Path $localPy "tcl\$d") (Join-Path $py "tcl\$d")
}

# 3. 套件：版本鎖死成本機實測過的那一組
$site = Join-Path $py "Lib\site-packages"
python -m pip install --quiet --disable-pip-version-check --no-deps `
    --target $site --only-binary=:all: `
    --platform win_amd64 --python-version $ver --implementation cp `
    pywinauto==0.6.9 comtypes==1.4.16 pywin32==312 six==1.17.0
if ($LASTEXITCODE -ne 0) { throw "pip 安裝套件失敗" }

# 4. 讓 embeddable Python 看得到 site-packages（pywin32 靠 .pth 設 DLL 路徑，import site 要打開）
$pth = Join-Path $py "$tag._pth"
@("$tag.zip", ".", "Lib\site-packages", "import site") |
    Set-Content -Encoding ascii $pth

# 4a. 刪掉 pywin32 附的、用不到的執行檔（沒有數位簽章，防毒最容易盯上 exe）。.pyd 模組全部保留。
foreach ($rel in "bin", "isapi", "pythonwin\Pythonwin.exe", "win32\pythonservice.exe") {
    $p = Join-Path $site $rel
    if (Test-Path $p) { Remove-Item -Recurse -Force $p }
}

# 4b. 先把套件編譯成 .pyc（unchecked-hash：zip 解壓後時間對不上也不會重編）
& (Join-Path $py "python.exe") -W ignore -m compileall -q -f --invalidation-mode unchecked-hash $site | Out-Null

# 5. 程式
foreach ($f in "ct_auto.py", "ct_merge.py", "ct_merge_no_merge.txt") {
    Copy-Item (Join-Path $here $f) $out
}
Copy-Item (Join-Path $here "portable_make_shortcut.py") (Join-Path $out "make_shortcut.py")

# 6. 啟動檔。bat 裡只放英文和 %~dp0（中文會隨命令列編碼變亂碼，2026-09-29 實測），
#    要中文的事交給 Python 做。整包搬到哪裡都能跑。
$launch = @"
@echo off
start "" "%~dp0python\pythonw.exe" "%~dp0ct_auto.py"
"@
$mkLink = @"
@echo off
"%~dp0python\python.exe" "%~dp0make_shortcut.py"
pause
"@
foreach ($pair in @(@("$name.bat", $launch), @("建立桌面捷徑.bat", $mkLink))) {
    $text = $pair[1].Replace("`r`n", "`n").Replace("`n", "`r`n") + "`r`n"
    [System.IO.File]::WriteAllText((Join-Path $out $pair[0]), $text,
                                   [System.Text.Encoding]::ASCII)
}

# 7. 壓成 zip（用 Python 的 zipfile：中文檔名會標 UTF-8，別台電腦解壓不會亂碼）
$outZip = Join-Path $here "portable_build\$name`_免安裝版.zip"
if (Test-Path $outZip) { Remove-Item -Force $outZip }
python -c "import shutil, sys; shutil.make_archive(sys.argv[1][:-4], 'zip', sys.argv[2], sys.argv[3])" $outZip (Split-Path $out) $name
if ($LASTEXITCODE -ne 0) { throw "壓縮失敗" }

$size = "{0:N1}" -f ((Get-Item $outZip).Length / 1MB)
Write-Host ""
Write-Host "完成：$out"
Write-Host "      $outZip （$size MB）"
Write-Host ""
Write-Host "放到 CT 檢查室："
Write-Host "  1. 解壓縮到那台電腦（例如 D:\$name），雙擊「建立桌面捷徑.bat」一次"
Write-Host "  2. RIS 開「登錄作業清單」、選這間檢查室，用桌面捷徑開，按「開始」"
