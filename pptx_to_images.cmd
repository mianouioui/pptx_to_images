@echo off
rem ============================================================
rem  PPTX 转图片器 V1.0.1 - Windows 启动器
rem  直接双击或拖拽 .pptx 到本文件即可运行
rem ============================================================
setlocal
chcp 65001 >nul
set "PPTX2IMG_SCRIPT=%~f0"
set "PPTX2IMG_INPUT=%*"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -Command "$script=[IO.File]::ReadAllText($env:PPTX2IMG_SCRIPT,[Text.Encoding]::UTF8); $parts=[regex]::Split($script,'(?m)^# POWERSHELL_START\r?$',2); if($parts.Count -lt 2){throw 'PowerShell section missing'}; iex $parts[1]"
set "exitCode=%ERRORLEVEL%"
if not "%PPTX2IMG_NO_PAUSE%"=="1" pause
exit /b %exitCode%
# POWERSHELL_START
$ErrorActionPreference = "Stop"
$Utf8NoBom = New-Object System.Text.UTF8Encoding $false

try {
    [Console]::OutputEncoding = $Utf8NoBom
    [Console]::InputEncoding = $Utf8NoBom
    $OutputEncoding = $Utf8NoBom
}
catch {}

$Version = "1.0.1"
$Width = 1920

function Fail {
    param([string]$Message)
    Write-Host ""
    Write-Host $Message
    exit 1
}

function Get-FirstDroppedPath {
    param([string]$Line)
    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $null
    }
    $trimmed = $Line.Trim()
    $quote = $trimmed[0]
    if ($quote -eq '"' -or $quote -eq "'") {
        $end = $trimmed.IndexOf($quote, 1)
        if ($end -gt 0) {
            return $trimmed.Substring(1, $end - 1)
        }
    }
    return ($trimmed -split "\s+", 2)[0]
}

function Split-Arguments {
    param([string]$Line)
    $results = @()
    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $results
    }
    foreach ($match in [regex]::Matches($Line, '"([^"]*)"|''([^'']*)''|(\S+)')) {
        if ($match.Groups[1].Success) {
            $results += $match.Groups[1].Value
        }
        elseif ($match.Groups[2].Success) {
            $results += $match.Groups[2].Value
        }
        else {
            $results += $match.Groups[3].Value
        }
    }
    return $results
}

function Resolve-InputPresentations {
    param([string]$RawPath)

    if ([string]::IsNullOrWhiteSpace($RawPath)) {
        return @()
    }

    $expanded = [Environment]::ExpandEnvironmentVariables($RawPath.Trim())
    try {
        $full = [System.IO.Path]::GetFullPath($expanded)
    }
    catch {
        return @()
    }

    if (Test-Path -LiteralPath $full -PathType Leaf) {
        $ext = [System.IO.Path]::GetExtension($full).ToLowerInvariant()
        if ($ext -eq ".pptx" -or $ext -eq ".ppt") {
            return @((Get-Item -LiteralPath $full).FullName)
        }
        return @()
    }

    if (Test-Path -LiteralPath $full -PathType Container) {
        return @(Get-ChildItem -LiteralPath $full -File |
            Where-Object { $_.Extension.ToLowerInvariant() -in @(".pptx", ".ppt") } |
            Sort-Object Name |
            ForEach-Object { $_.FullName })
    }

    return @()
}

function Get-UniqueOutputDir {
    param([string]$PresentationPath)

    $parent = Split-Path -Path $PresentationPath -Parent
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($PresentationPath)
    $output = Join-Path $parent ($stem + "_images")
    if (-not (Test-Path -LiteralPath $output)) {
        New-Item -ItemType Directory -Path $output -Force | Out-Null
        return $output
    }

    for ($i = 2; $i -lt 1000; $i++) {
        $candidate = Join-Path $parent ("{0}_images_{1}" -f $stem, $i)
        if (-not (Test-Path -LiteralPath $candidate)) {
            New-Item -ItemType Directory -Path $candidate -Force | Out-Null
            return $candidate
        }
    }

    Fail "输出文件夹已存在，且无法生成新的文件夹名。"
}

function Show-Banner {
    Write-Host "============================================================"
    Write-Host "  PPTX 转图片   v$Version"
    Write-Host "============================================================"
    Write-Host "  把 PPTX/PPT 演示文稿的每一页导出为 PNG 图片。"
    Write-Host ""
    Write-Host "  怎么用："
    Write-Host "    1. 把一个 PPTX 文件拖到这个窗口里（拖文件夹也行）"
    Write-Host "    2. 按回车"
    Write-Host ""
    Write-Host "  输出："
    Write-Host "    · 默认生成到同目录的 “文件名_images” 文件夹"
    Write-Host "    · 图片命名为 page001.png、page002.png ..."
    Write-Host "    · 拖入文件夹时，会转换里面的 PPTX/PPT 文件"
    Write-Host "------------------------------------------------------------"
}

function Read-InteractivePresentations {
    Show-Banner
    while ($true) {
        [Console]::Write("把 PPTX 文件拖到这里，然后按回车：")
        $line = [Console]::ReadLine()
        if ($null -eq $line) {
            return @()
        }
        $raw = Get-FirstDroppedPath $line
        $items = Resolve-InputPresentations $raw
        if ($items.Count -eq 0) {
            Write-Host "  没有找到 PPTX/PPT 文件，请重新拖入。"
            Write-Host ""
            continue
        }
        return $items
    }
}

function Rename-ExportedSlides {
    param([string]$RawDir, [string]$OutputDir)

    $files = @(Get-ChildItem -LiteralPath $RawDir -File -Filter "*.png" |
        Sort-Object {
            if ($_.BaseName -match "(\d+)$") { [int]$Matches[1] } else { 0 }
        }, Name)

    if ($files.Count -eq 0) {
        Fail "PowerPoint 没有生成 PNG 图片。"
    }

    $index = 1
    foreach ($file in $files) {
        $target = Join-Path $OutputDir ("page{0:D3}.png" -f $index)
        if (Test-Path -LiteralPath $target) {
            Remove-Item -LiteralPath $target -Force
        }
        Move-Item -LiteralPath $file.FullName -Destination $target
        $index++
    }

    return $files.Count
}

function Export-WithPowerPoint {
    param([string]$PresentationPath, [string]$OutputDir)

    $rawDir = Join-Path $OutputDir ".raw_powerpoint"
    if (Test-Path -LiteralPath $rawDir) {
        Remove-Item -LiteralPath $rawDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $rawDir -Force | Out-Null

    $powerPoint = $null
    $presentation = $null
    try {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $powerPoint.Visible = -1
        $presentation = $powerPoint.Presentations.Open($PresentationPath, $true, $false, $false)
        $ratio = [double]$presentation.PageSetup.SlideHeight / [double]$presentation.PageSetup.SlideWidth
        $height = [int][Math]::Round($Width * $ratio)
        $presentation.Export($rawDir, "PNG", $Width, $height)
    }
    finally {
        if ($presentation -ne $null) {
            try { $presentation.Close() } catch {}
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
        }
        if ($powerPoint -ne $null) {
            try { $powerPoint.Quit() } catch {}
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    $count = Rename-ExportedSlides -RawDir $rawDir -OutputDir $OutputDir
    Remove-Item -LiteralPath $rawDir -Recurse -Force -ErrorAction SilentlyContinue
    return $count
}

try {
    $items = @()
    if ([string]::IsNullOrWhiteSpace($env:PPTX2IMG_INPUT)) {
        $items = Read-InteractivePresentations
    }
    else {
        foreach ($argPath in (Split-Arguments $env:PPTX2IMG_INPUT)) {
            $items += Resolve-InputPresentations $argPath
        }
        $items = @($items | Select-Object -Unique)
    }

    if ($items.Count -eq 0) {
        Fail "没有找到 PPTX/PPT 文件。"
    }

    Write-Host "待转换文件："
    $n = 1
    foreach ($item in $items) {
        Write-Host ("  {0:D3}. {1}" -f $n, $item)
        $n++
    }

    foreach ($item in $items) {
        $output = Get-UniqueOutputDir $item
        Write-Host ""
        Write-Host "正在导出：$([System.IO.Path]::GetFileName($item))"
        Write-Host "输出文件夹：$output"
        $count = Export-WithPowerPoint -PresentationPath $item -OutputDir $output
        Write-Host "完成：$count 张 PNG"
        Start-Process -FilePath $output
    }
}
catch {
    Fail $_.Exception.Message
}
