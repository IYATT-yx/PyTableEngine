<#
file: BuildPlugin.ps1
description: 插件打包脚本。
            PowerShell 7.x 验证可用。
author: IYATT-yx
copyright:  Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
#>
[CmdletBinding()]
param (
    [switch]$DebugBuild
);

$ErrorActionPreference = 'Stop';

$ScriptDir =$PSScriptRoot;
$RootPath = (Resolve-Path "$ScriptDir\..").Path;

Set-Location $ScriptDir;
$startTime = Get-Date;

if ($DebugBuild) {
    Write-Host '[BUILD] 正在以 Debug 模式编译插件 (已手动禁用 LTO)...' -ForegroundColor Yellow;
    $ltoOption = '--lto=no';
} else {
    Write-Host '[BUILD] 正在以 Release 模式编译插件 (已默认启用 LTO 极致优化)...' -ForegroundColor Cyan;
    $ltoOption = '--lto=yes';
}

# 虚拟环境与 Nuitka 检查
$parentPython = Join-Path $RootPath 'venv' 'Scripts' 'python.exe';
$localVenv    = Join-Path $ScriptDir 'venv';
$localPython  = Join-Path $localVenv 'Scripts' 'python.exe';

$targetPython =$null;

if (Test-Path $parentPython) {
    Write-Host '[INFO] 检测到上一级目录虚拟环境，正在验证 Nuitka...' -ForegroundColor Cyan;
    & $parentPython -c 'import nuitka' 2>&1 | Out-Null;
    if ($LASTEXITCODE -eq 0) {
        $targetPython =$parentPython;
        Write-Host '[SUCCESS] 成功复用上一级虚拟环境的 Nuitka！' -ForegroundColor Green;
    }
}

if (-not $targetPython -and (Test-Path $localPython)) {
    Write-Host '[INFO] 检测到 extensions 本地虚拟环境，正在验证 Nuitka...' -ForegroundColor Cyan;
    & $localPython -c 'import nuitka' 2>&1 | Out-Null;
    if ($LASTEXITCODE -eq 0) {
        $targetPython =$localPython;
        Write-Host '[SUCCESS] 成功复用 extensions 本地虚拟环境的 Nuitka！' -ForegroundColor Green;
    }
}

if (-not $targetPython) {
    Write-Host '[INFO] 未检测到有效的 Nuitka 环境，准备在 extensions 目录下初始化...' -ForegroundColor Yellow;
    if (-not (Test-Path $localVenv)) {
        python -m venv $localVenv;
    }
    $targetPython =$localPython;
    & $targetPython -m pip install --upgrade pip;
    & $targetPython -m pip install nuitka==4.2.1;
}

# 选择要打包的目录
Add-Type -AssemblyName System.Windows.Forms;
$folderDialog = [System.Windows.Forms.FolderBrowserDialog]::new();
$folderDialog.Description = '请选择要打包的插件目录（或 extensions 根目录）';

$folderDialog.SelectedPath = "$ScriptDir\"; 

$folderDialog.ShowNewFolderButton = $false;

if ($folderDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host '[INFO] 用户取消了操作。' -ForegroundColor Yellow;
    exit 0;
}

$selectedDir =$folderDialog.SelectedPath;
Write-Host "[INFO] 已选择路径: $selectedDir" -ForegroundColor Cyan;

$targetPlugins = [System.Collections.Generic.List[PSCustomObject]]::new();
$folderName = Split-Path -Path $selectedDir -Leaf;
$potentialMainPy = Join-Path $selectedDir "$folderName.py";

if (Test-Path $potentialMainPy) {$targetPlugins.Add([PSCustomObject]@{
        Dir  = $selectedDir;
        File = $potentialMainPy;
    });
} else {
    Get-ChildItem -Path $selectedDir -Directory | Where-Object { $_.Name -ne 'venv' -and -not $_.Name.EndsWith('.build') } | ForEach-Object {
        $subPy = Join-Path $_.FullName "$($_.Name).py";
        if (Test-Path $subPy) {$targetPlugins.Add([PSCustomObject]@{
                Dir  = $_.FullName;
                File = $subPy;
            });
        }
    };
}

if ($targetPlugins.Count -eq 0) {
    Write-Host '[WARN] 未在所选目录中找到符合 [插件目录名/插件目录名.py] 命名规则的插件文件！' -ForegroundColor Yellow;
    exit 1;
}

$env:CL = '/utf-8';

# Nuitka 打包
foreach ($plugin in $targetPlugins) {
    Write-Host "`n[BUILD] 开始打包插件: $($plugin.File)" -ForegroundColor Yellow;

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($plugin.File);
    $targetPydName = "$baseName.pyd";
    $finalPydPath = Join-Path $plugin.Dir $targetPydName;

    & $targetPython -m nuitka --module `
        $ltoOption `
        --nofollow-import-to=pytableenginesdk `
        --output-dir="$($plugin.Dir)" `
        "$($plugin.File)";

    if ($LASTEXITCODE -eq 0) {
        Write-Host '[SUCCESS] 插件打包完成 -> 正在处理文件重命名...' -ForegroundColor Green;

        # 查找 Nuitka 生成带 ABI 后缀的 .pyd 文件（如 xxx.cp314-win_amd64.pyd）
        $rawPyd = Get-ChildItem -Path $plugin.Dir -Filter "$baseName*.pyd" | Where-Object { $_.Name -ne $targetPydName } | Select-Object -First 1;

        if ($rawPyd) {
            # 如果存在旧的同名干净文件，先移除防止覆盖报错
            if (Test-Path $finalPydPath) {
                Remove-Item -Path $finalPydPath -Force;
            }
            # 重命名为标准的纯净文件名 <插件名>.pyd
            Rename-Item -Path $rawPyd.FullName -NewName $targetPydName -Force;

            $buildCacheDir = Join-Path $plugin.Dir "$baseName.build";
            if (Test-Path $buildCacheDir) {
                Remove-Item -Path $buildCacheDir -Recurse -Force;
            }
        }

        # 打开文件夹并高亮选中最终生成的 .pyd 文件
        if (Test-Path $finalPydPath) {
            Start-Process "explorer.exe" -ArgumentList "/select,`"$finalPydPath`"";
        }
    } else {
        Write-Host "[ERROR] 插件打包失败: $($plugin.File)" -ForegroundColor Red;
        exit 1;
    }
}

$elapsedTime = (Get-Date) - $startTime;
Write-Host "`n[COMPLETE] 编译流程结束，总用时：$([math]::Round($elapsedTime.TotalSeconds, 2)) 秒" -ForegroundColor Green;