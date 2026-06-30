<#
.SYNOPSIS
    Install word-mcp-server and configure Claude Code.
.DESCRIPTION
    Installs word-mcp-server from GitHub via pip and writes the MCP
    server entry into the appropriate settings.local.json for Claude Code.
#>

$ErrorActionPreference = "Stop"
$RepoUrl = "https://github.com/wpz666hhh/word-mcp-server.git"

Write-Host "=== Word MCP Server Installer ===" -ForegroundColor Cyan
Write-Host ""

# --- Step 1: Check Python ---
try {
    $py = Get-Command python -ErrorAction Stop
    $ver = & python --version 2>&1
    Write-Host "[1/4] Python 已找到: $ver" -ForegroundColor Green
} catch {
    Write-Host "[1/4] 未找到 Python。请先安装 Python 3.10+。" -ForegroundColor Red
    Write-Host "      下载地址: https://www.python.org/downloads/"
    exit 1
}

# --- Step 2: Install package ---
Write-Host "[2/4] 正在安装 word-mcp-server 从 GitHub ..." -ForegroundColor Yellow
try {
    & python -m pip install "git+$RepoUrl" --quiet
    Write-Host "      安装完成" -ForegroundColor Green
} catch {
    Write-Host "      安装失败: $_" -ForegroundColor Red
    exit 1
}

# --- Step 3: Write Claude Code config ---
$claudeDir = "$env:USERPROFILE\.claude"
$localSettings = "$claudeDir\settings.local.json"
$projectSettings = ""

# Try project-level .claude first
$cwd = (Get-Location).Path
$projectClaude = Join-Path $cwd ".claude"
if (Test-Path "$projectClaude\settings.local.json") {
    $targetFile = "$projectClaude\settings.local.json"
    Write-Host "[3/4] 检测到项目级配置: $targetFile" -ForegroundColor Green
} else {
    # Fall back to user-level config
    if (-not (Test-Path $claudeDir)) {
        New-Item -ItemType Directory -Path $claudeDir -Force | Out-Null
    }
    $targetFile = $localSettings
    Write-Host "[3/4] 使用用户级配置: $targetFile" -ForegroundColor Green
}

# Read existing or create new settings
$settings = @{}
if (Test-Path $targetFile) {
    try {
        $settings = Get-Content $targetFile -Raw -Encoding UTF8 | ConvertFrom-Json -AsHashtable
    } catch {
        $settings = @{}
    }
}

# Merge MCP server entry
if (-not $settings.ContainsKey("mcpServers")) {
    $settings["mcpServers"] = @{}
}
$settings["mcpServers"]["word"] = @{
    command = "python"
    args    = @("-m", "word_mcp")
    type    = "local"
}

$settingsJson = $settings | ConvertTo-Json -Depth 4
$settingsJson | Out-File $targetFile -Encoding UTF8
Write-Host "      已写入 MCP 配置到: $targetFile" -ForegroundColor Green

# --- Step 4: Done ---
Write-Host ""
Write-Host "[4/4] ✓ 安装完成！" -ForegroundColor Green
Write-Host ""
Write-Host "请重启 Claude Code 使配置生效。" -ForegroundColor Yellow
Write-Host ""
Write-Host "验证方式: 在 Claude Code 中发送：" -ForegroundColor Cyan
Write-Host '  /word_get_active_document' -ForegroundColor White
Write-Host ""
Write-Host "如果遇到问题，请提交 Issue:" -ForegroundColor Gray
Write-Host "  https://github.com/wpz666hhh/word-mcp-server/issues" -ForegroundColor Gray
