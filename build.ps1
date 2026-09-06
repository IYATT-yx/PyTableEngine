$startTime = Get-Date

python -m venv venv
.\venv\Scripts\Activate.ps1
python.exe -m pip install --upgrade pip
pip install nuitka==4.2
pip install -r requirements.txt

python .\savebuildtime.py

$env:CL = "/utf-8"

nuitka --standalone `
--windows-console-mode=disable `
--lto=yes `
--no-deployment-flag=self-contained `
--enable-plugin=tk-inter `
--include-package=win32com `
--include-package=win32comext `
--include-package=win32api `
--include-package=win32gui `
--include-package=win32con `
--include-package=pythoncom `
--include-package=pywintypes `
--noinclude-unittest-mode=nofollow `
--windows-company-name="IYATT-yx" `
--windows-product-name="思能快表引擎" `
--windows-file-description="思能快表引擎" `
--windows-product-version="1.0.0.0" `
--windows-file-version="1.0.0.0" `
--copyright="Copyright (C) 2026 IYATT-yx. All Rights Reserved." `
--windows-icon-from-ico=.\icon.ico `
--include-data-file=.\icon.ico=.\ `
--include-data-files=.\extensions\IYATT_DemoPlugin\*=extensions/IYATT_DemoPlugin/ `
--include-data-files=.\pytableenginesdk\*=pytableenginesdk/ `
--noinclude-data-files="*__pycache__*/*" `
--noinclude-data-files="*.pyc" `
--noinclude-data-files="*.pyo" `
--output-dir=dist `
--output-filename=PyTableEngine_win_amd64 `
.\PyTableEngine.py

$endTime = Get-Date
$elapsedTime = New-TimeSpan -Start $startTime -End $endTime
Write-Output "程序构建用时：$($elapsedTime.TotalSeconds) 秒"