<#
file: build.ps1
description: PyTableEngine 构建脚本。
            PowerShell 7.x 验证可用。
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
#>
[CmdletBinding()]
param (
    [switch]$DebugBuild
)

$startTime = Get-Date

# 根据是否传入 -DebugBuild 参数动态设置 LTO 和编译选项
if ($DebugBuild) {
    Write-Host "[BUILD] 正在以 Debug 模式构建 (已禁用 LTO 以加速编译)..." -ForegroundColor Yellow
    $ltoOption = "--lto=no"
} else {
    Write-Host "[BUILD] 正在以 Release 模式构建 (已启用 LTO 极致优化)..." -ForegroundColor Cyan
    $ltoOption = "--lto=yes"
}

if (-not (Test-Path "venv")) {
    python -m venv venv;
}
.\venv\Scripts\Activate.ps1
python.exe -m pip install --upgrade pip
pip install nuitka==4.2.1
pip install -r requirements.txt

python .\savebuildtime.py

$embedScript = ".\SetupEmbedPython.ps1"

if (-not (Test-Path $embedScript)) {
    Write-Host "[FATAL] 未找到 Embed 环境配置脚本: $embedScript，终止构建！" -ForegroundColor Red
    exit 1
}
Write-Host "[BUILD] 正在执行 SetupEmbedPython.ps1 部署内嵌 Runtime..." -ForegroundColor Cyan
& $embedScript
if ($LASTEXITCODE -ne 0) {
    Write-Host "[FATAL] SetupEmbedPython.ps1 执行失败，代码: $LASTEXITCODE，终止构建！" -ForegroundColor Red
    exit $LASTEXITCODE
}
# 校验 runtime\python.exe 和 pip 可用性
$embedPython = ".\runtime\python.exe"
if (-not (Test-Path $embedPython)) {
    Write-Host "[FATAL] 未生成 runtime\python.exe，终止构建！" -ForegroundColor Red
    exit 1
}
& $embedPython -m pip --version 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[FATAL] 内嵌 pip 功能校验失败，终止构建！" -ForegroundColor Red
    exit 1
}
Write-Host "[SUCCESS] Embed Runtime 准备就绪，开始执行 Nuitka 打包..." -ForegroundColor Green
# ------------------------------------------------------------------

if (Test-Path ".\dist\PyTableEngine.dist") {
    Remove-Item -Path ".\dist\PyTableEngine.dist" -Recurse -Force
}

$env:CL = "/utf-8"

nuitka --standalone `
--windows-console-mode=disable `
$ltoOption `
--no-deployment-flag=self-contained `
--enable-plugin=tk-inter `
--include-package=tkinter `
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
--include-data-files=.\extensions\README.md=extensions/ `
--include-data-files=.\extensions\BuildPlugin.ps1=extensions/ `
--include-data-files=.\pytableenginesdk\*=pytableenginesdk/ `
--noinclude-data-files="*__pycache__*/*" `
--output-dir=dist `
--output-filename=PyTableEngine_win_amd64 `
.\PyTableEngine.py

$expectedExe = ".\dist\PyTableEngine.dist\PyTableEngine_win_amd64.exe"
# 双重重校验：退出码 + 文件存在性
if ($LASTEXITCODE -ne 0 -or -not (Test-Path $expectedExe)) {
    Write-Host "[FATAL] Nuitka 打包失败！未生成预期文件: $expectedExe" -ForegroundColor Red
    exit 1
}

# --- 后处理：复制完整的 Embed Runtime 文件夹---
$distRuntimeDir = ".\dist\PyTableEngine.dist\runtime"
Write-Host "[BUILD] 正在部署完整的 Embed Runtime 到 dist 目录..." -ForegroundColor Cyan

if (Test-Path $distRuntimeDir) {
    Remove-Item -Path $distRuntimeDir -Recurse -Force
}

# 复制目录树
Copy-Item -Path ".\runtime" -Destination $distRuntimeDir -Recurse -Force

# 校验 runtime\python.exe 是否成功复制
if (Test-Path "$distRuntimeDir\python.exe") {
    Write-Host "[SUCCESS] Embed Runtime 已完整部署到打包输出目录！" -ForegroundColor Green
} else {
    Write-Host "[ERROR] 复制 Embed Runtime 失败！" -ForegroundColor Red
    exit 1
}

$endTime = Get-Date
$elapsedTime = New-TimeSpan -Start $startTime -End $endTime
Write-Output "程序构建用时：$($elapsedTime.TotalSeconds) 秒"