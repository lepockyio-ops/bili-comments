# BiliComments PowerShell 包装
# 用法：
#   .\run.ps1 BV1FDNJ6BE4j
#   .\run.ps1 https://www.bilibili.com/video/BV1FDNJ6BE4j --include-replies
#   .\run.ps1 BV1FDNJ6BE4j --sort new --max-pages 10

$ErrorActionPreference = "Stop"
$env:PYTHONIOENCODING = "utf-8"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 清除可能被其他工具（如 Claude Cowork 的 Windows-MCP）污染的 venv 环境变量
# 避免 py 被劫持到别的 Python 环境（导致 pip install 失败）
Remove-Item Env:VIRTUAL_ENV -ErrorAction SilentlyContinue
Remove-Item Env:PYTHONHOME -ErrorAction SilentlyContinue
Remove-Item Env:PYTHONPATH -ErrorAction SilentlyContinue

Set-Location $PSScriptRoot

# 依赖检测
$deps = py -c "import httpx, openpyxl" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "首次运行，正在安装依赖..." -ForegroundColor Yellow
    py -m pip install -r requirements.txt --quiet --disable-pip-version-check
}

if ($args.Count -eq 0) {
    Write-Host "用法：.\run.ps1 <bvid_or_url> [--out xxx.xlsx] [--include-replies] [--sort hot|new] [--max-pages N]" -ForegroundColor Yellow
    exit 1
}

py -X utf8 collect_comments.py @args
exit $LASTEXITCODE
