<#
.SYNOPSIS
    整车测试报告一键生成 - 从MF4到Word报告的完整流程

.DESCRIPTION
    完整流程：MF4文件 → 信号提取 → 指标计算 → JSON回填 → Word报告生成

    步骤：
    1. 解析MF4文件，提取CAN信号
    2. 自动检测制动事件
    3. 计算测试指标（平均减速度、制动距离等）
    4. 输出中间JSON数据
    5. 生成Word格式测试报告

.PARAMETER Mf4File
    MF4数据文件路径

.PARAMETER DataFile
    已有的JSON数据文件路径（跳过MF4解析，直接生成报告）

.PARAMETER OutputDir
    输出目录（默认: output/）

.PARAMETER SignalConfig
    信号映射配置文件路径（默认: config/signal_mapping.json）

.PARAMETER SkipParse
    跳过MF4解析步骤，直接用现有JSON生成报告

.EXAMPLE
    .\run.ps1 -Mf4File "C:\test_data\2026-04-16_run1.mf4"
    .\run.ps1 -DataFile "output/parsed_data.json" -SkipParse
    .\run.ps1 -Mf4File "data.mf4" -SignalConfig "my_signals.json"
#>

param(
    [Parameter(HelpMessage="MF4数据文件路径")]
    [string]$Mf4File,

    [Parameter(HelpMessage="已有的JSON数据文件")]
    [string]$DataFile,

    [Parameter(HelpMessage="输出目录")]
    [string]$OutputDir = (Join-Path $PSScriptRoot "..\output"),

    [Parameter(HelpMessage="信号映射配置文件")]
    [string]$SignalConfig = (Join-Path $PSScriptRoot "..\config\signal_mapping.json"),

    [Parameter(HelpMessage="跳过MF4解析，直接生成报告")]
    [switch]$SkipParse,

    [Parameter(HelpMessage="仅解析MF4，不生成报告")]
    [switch]$ParseOnly,

    [Parameter(HelpMessage="仅列出MF4中的可用信号")]
    [switch]$ListSignals
)

$ErrorActionPreference = "Stop"

# 路径解析
$SkillDir = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$ConfigDir = Join-Path $SkillDir "config"
$ScriptsDir = Join-Path $SkillDir "scripts"
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)

# 确保输出目录存在
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$parsedJsonPath = Join-Path $OutputDir "parsed_data_$timestamp.json"
$reportPath = Join-Path $OutputDir "vehicle_test_report_$timestamp.docx"

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  整车功能性能测试报告 - 一键生成" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# ============================================================
# 步骤1：列出MF4信号（可选）
# ============================================================
if ($ListSignals) {
    if (-not $Mf4File) {
        Write-Error "请指定MF4文件: -Mf4File <路径>"
        exit 1
    }
    Write-Host "[步骤0] 列出MF4文件中的可用信号..." -ForegroundColor Yellow
    python (Join-Path $ScriptsDir "mf4_parser.py") --input $Mf4File --list-signals
    exit 0
}

# ============================================================
# 步骤1：解析MF4文件（如未跳过）
# ============================================================
if (-not $SkipParse) {
    if (-not $Mf4File) {
        Write-Error "请指定MF4文件: -Mf4File <路径>`n或使用 -SkipParse -DataFile <JSON路径> 跳过解析"
        exit 1
    }

    if (-not (Test-Path $Mf4File)) {
        Write-Error "MF4文件不存在: $Mf4File"
        exit 1
    }

    Write-Host "[步骤1] 解析MF4文件..." -ForegroundColor Yellow
    Write-Host "  输入: $Mf4File"
    Write-Host "  信号配置: $SignalConfig"
    Write-Host "  输出: $parsedJsonPath"
    Write-Host ""

    $requirementsPath = Join-Path $ConfigDir "test_requirements.json"

    python (Join-Path $ScriptsDir "mf4_parser.py") `
        --input $Mf4File `
        --config $SignalConfig `
        --output $parsedJsonPath `
        --requirements $requirementsPath

    if ($LASTEXITCODE -ne 0) {
        Write-Error "MF4解析失败！请检查Python环境和asammdf库。"
        Write-Host "安装依赖: pip install asammdf numpy pandas" -ForegroundColor Yellow
        exit 1
    }

    Write-Host ""
    Write-Host "[步骤1] ✓ MF4解析完成" -ForegroundColor Green
    Write-Host "  中间数据: $parsedJsonPath" -ForegroundColor Gray

    if ($ParseOnly) {
        Write-Host ""
        Write-Host "仅解析模式，跳过报告生成。" -ForegroundColor Yellow
        exit 0
    }

    $DataFile = $parsedJsonPath
}

# ============================================================
# 步骤2：生成Word报告
# ============================================================
if (-not $DataFile -or -not (Test-Path $DataFile)) {
    Write-Error "数据文件不存在: $DataFile"
    exit 1
}

Write-Host ""
Write-Host "[步骤2] 生成Word报告..." -ForegroundColor Yellow
Write-Host "  数据文件: $DataFile"
Write-Host "  输出: $reportPath"
Write-Host ""

& (Join-Path $ScriptsDir "generate_report.ps1") -DataFile $DataFile -OutputFile $reportPath

if ($LASTEXITCODE -ne 0) {
    Write-Error "报告生成失败！"
    exit 1
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  ✓ 报告生成完成！" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "  报告文件: $reportPath" -ForegroundColor White
if (-not $SkipParse) {
    Write-Host "  中间数据: $parsedJsonPath" -ForegroundColor Gray
}
Write-Host ""

# 打开报告
$openReport = Read-Host "是否打开报告？(Y/n)"
if ($openReport -ne "n") {
    Start-Process $reportPath
}
