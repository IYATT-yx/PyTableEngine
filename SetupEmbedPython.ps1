<#
    file: SetupEmbedPython.ps1
    description: 自动化下载 Python Embed，配置 .pth 文件并安装 pip
    author: IYATT-yx
    copyright:   Copyright (c) 2026 IYATT-yx.
                Licensed under the MIT License. See LICENSE file in the project root for full license information.
#>
# --- 路径与下载地址变量集中声明 ---
Param(
    [string]$TargetDir = ".\runtime",                                                             # 目标下载路径
    [string]$PythonUrl = "https://www.python.org/ftp/python/3.14.5/python-3.14.5-embed-amd64.zip", # Python Embed 下载地址
    [string]$GetPipUrl = "https://bootstrap.pypa.io/get-pip.py"                                   # get-pip.py 下载地址
)

$ErrorActionPreference = "Stop"

# ---  内部衍生路径变量派生 ---
$zipPath = Join-Path $TargetDir "python-embed.zip"
$getPipPath = Join-Path $TargetDir "get-pip.py"
$embedPython = Join-Path $TargetDir "python.exe"

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Python Embed 自动化部署与 Pip 环境配置" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# --- 代理配置 ---
$timeoutSeconds = 10  # 倒计时等待时间（秒）

Write-Host "[INFO] 网络代理配置提示：" -ForegroundColor Cyan

$startTime = Get-Date
$keyPressed = $false

# 倒计时并原地刷新显示
while (($remainingSeconds = [math]::Max(0, [math]::Ceiling($timeoutSeconds - ((Get-Date) - $startTime).TotalSeconds))) -gt 0) {
    if ([System.Console]::KeyAvailable) {
        $keyPressed = $true
        break
    }
    # \r 回退光标至行首，实现原地刷新显示
    Write-Host "`r请输入代理地址 (按下任意键开始输入，${remainingSeconds}秒无操作自动跳过): " -NoNewline -ForegroundColor Yellow
    Start-Sleep -Milliseconds 50
}

if ($keyPressed) {
    # 捕获按键并开始正常接收用户输入
    Write-Host "`r请输入代理地址 (直接回车不使用代理):                             `n> " -NoNewline -ForegroundColor Cyan
    $proxyInput = Read-Host
    if (-not [string]::IsNullOrWhitespace($proxyInput)) {
        $proxyUrl =$proxyInput.Trim()
        $env:HTTP_PROXY =$proxyUrl
        $env:HTTPS_PROXY =$proxyUrl
        Write-Host "[INFO] 已设置进程代理地址: $proxyUrl" -ForegroundColor Yellow
    } else {
        Write-Host "[INFO] 未设置代理，将直接连接下载。" -ForegroundColor Yellow
    }
} else {
    # 超时自动跳过
    Write-Host "`r[INFO] 超时 ${timeoutSeconds} 秒未响应，自动跳过代理配置，直接连接下载。                     " -ForegroundColor Yellow
    $proxyInput = ""
}

$iwrArgs = @{}
if (-not [string]::IsNullOrWhitespace($proxyInput)) {
    $iwrArgs["Proxy"] = $proxyInput.Trim()
}

# 清理并创建目标目录
if (Test-Path $TargetDir) {
    Write-Host "[INFO] 清理旧的 runtime 目录..." -ForegroundColor Yellow
    Remove-Item -Path $TargetDir -Recurse -Force
}
New-Item -ItemType Directory -Path $TargetDir | Out-Null

# 下载 Python Embed Zip
Write-Host "[INFO] 正在下载 Python Embed 压缩包..." -ForegroundColor Green
Invoke-WebRequest -Uri $PythonUrl -OutFile $zipPath @iwrArgs

# 解压压缩包
Write-Host "[INFO] 正在解压至 $TargetDir ..." -ForegroundColor Green
Expand-Archive -Path $zipPath -DestinationPath $TargetDir -Force
Remove-Item -Path $zipPath -Force

# 关键配置：修正 python*._pth 文件
# Embed 版本的 .pth 默认屏蔽了 site-packages 和 import site
$pthFile = Get-ChildItem -Path $TargetDir -Filter "python*._pth" | Select-Object -First 1

if ($pthFile) {
    Write-Host "[INFO] 正在修正 $($pthFile.Name) 配置，取消 import site 注释..." -ForegroundColor Green
    $content = Get-Content -Path $pthFile.FullName
    # 取消 'import site' 前面的注释符号 '#'
    $updatedContent = $content -replace '^\s*#\s*import site', 'import site'
    # 追加 current directory 和 site-packages 路径保障
    if (-not ($updatedContent -contains "Lib\site-packages")) {
        $updatedContent += "`r`nLib\site-packages"
    }
    Set-Content -Path $pthFile.FullName -Value $updatedContent -Encoding UTF8
} else {
    Write-Host "[ERROR] 未找到 .pth 配置文件，请检查解压目录！" -ForegroundColor Red
    exit 1
}

# 下载并运行 get-pip.py 安装 pip
Write-Host "[INFO] 正在下载 get-pip.py 引导脚本..." -ForegroundColor Green
Invoke-WebRequest -Uri $GetPipUrl -OutFile $getPipPath @iwrArgs

Write-Host "[INFO] 正在使用 Embed Python 执行 pip 引导安装..." -ForegroundColor Green

# 构建 get-pip.py 参数
$pipArgs = "`"$getPipPath`" --no-warn-script-location"
if (-not [string]::IsNullOrWhitespace($proxyInput)) {
    $pipArgs += " --proxy `"$($proxyInput.Trim())`""
}

# 执行 get-pip.py，把 pip 安装到 runtime 内部
$pipProcess = Start-Process -FilePath $embedPython -ArgumentList $pipArgs -Wait -NoNewWindow -PassThru

if ($pipProcess.ExitCode -eq 0) {
    Write-Host "[SUCCESS] pip 安装完成！" -ForegroundColor Green
    Remove-Item -Path $getPipPath -Force
} else {
    Write-Host "[ERROR] pip 安装失败，返回码: $($pipProcess.ExitCode)" -ForegroundColor Red
    exit 1
}

# 验证安装结果
Write-Host "[INFO] 正在验证嵌入式 pip 版本..." -ForegroundColor Green
& $embedPython -m pip --version

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "[SUCCESS] Embed Runtime 配置完成，可直接用于打包！" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Cyan