# manage-task.ps1
# RestoreCustomFonts Scheduled Task Manager
# 检查是否管理员
$currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {

    Write-Host "正在请求管理员权限..." -ForegroundColor Yellow

    $script = $MyInvocation.MyCommand.Definition

    Start-Process powershell `
        -Verb RunAs `
        -ArgumentList "-ExecutionPolicy Bypass -File `"$script`""

    exit
}
$TaskName = "RestoreCustomFonts"
$XmlPath = Join-Path $PSScriptRoot "restore-fonts-task.xml"
$RestoreScriptPath = Join-Path $PSScriptRoot "restore-fonts.ps1"

function New-TaskXmlForCurrentLocation {
    if (-not (Test-Path $XmlPath)) {
        Write-Host "XML not found: $XmlPath" -ForegroundColor Red
        return $null
    }

    if (-not (Test-Path $RestoreScriptPath)) {
        Write-Host "Script not found: $RestoreScriptPath" -ForegroundColor Red
        return $null
    }

    [xml]$taskXml = Get-Content -Raw -Path $XmlPath
    $ns = New-Object System.Xml.XmlNamespaceManager($taskXml.NameTable)
    $ns.AddNamespace("ts", "http://schemas.microsoft.com/windows/2004/02/mit/task")

    $commandNode = $taskXml.SelectSingleNode("//ts:Actions/ts:Exec/ts:Command", $ns)
    $argumentsNode = $taskXml.SelectSingleNode("//ts:Actions/ts:Exec/ts:Arguments", $ns)
    $workingDirectoryNode = $taskXml.SelectSingleNode("//ts:Actions/ts:Exec/ts:WorkingDirectory", $ns)

    if ($null -eq $commandNode -or $null -eq $argumentsNode -or $null -eq $workingDirectoryNode) {
        Write-Host "Invalid task XML: missing Exec nodes." -ForegroundColor Red
        return $null
    }

    $commandNode.InnerText = "powershell.exe"
    $argumentsNode.InnerText = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$RestoreScriptPath`""
    $workingDirectoryNode.InnerText = $PSScriptRoot

    $tempXmlPath = Join-Path $env:TEMP ("restore-fonts-task-{0}.xml" -f ([guid]::NewGuid().ToString("N")))
    $taskXml.Save($tempXmlPath)
    return $tempXmlPath
}

function Task-Exists {
    $null = schtasks /query /tn $TaskName 2>$null
    return $LASTEXITCODE -eq 0
}

function Show-Status {
    if (Task-Exists) {
        Write-Host "  当前状态：" -NoNewline
        Write-Host "已安装" -ForegroundColor Green
    } else {
        Write-Host "  当前状态：" -NoNewline
        Write-Host "未安装" -ForegroundColor Red
    }
    Write-Host ""
}

function Do-Add {
    if (Task-Exists) {
        Write-Host "任务已存在，请使用 [2] 更新。" -ForegroundColor Red
        return
    }

    $taskXmlPath = New-TaskXmlForCurrentLocation
    if ($null -eq $taskXmlPath) {
        return
    }

    try {
        Write-Host "脚本路径：$RestoreScriptPath"
        $output = schtasks /create /tn $TaskName /xml $taskXmlPath 2>&1
        $exitCode = $LASTEXITCODE
    }
    finally {
        Remove-Item -Path $taskXmlPath -Force -ErrorAction SilentlyContinue
    }

    if ($exitCode -eq 0) {
        Write-Host "任务创建成功。" -ForegroundColor Green
    } else {
        Write-Host "任务创建失败。" -ForegroundColor Red
        Write-Host "退出码：$exitCode" -ForegroundColor Yellow
        Write-Host "错误信息：" -ForegroundColor Yellow
        Write-Host $output
    }
}

function Do-Update {
    $taskXmlPath = New-TaskXmlForCurrentLocation
    if ($null -eq $taskXmlPath) {
        return
    }

    try {
        if (Task-Exists) {
            schtasks /delete /tn $TaskName /f >$null 2>&1
        }
        Write-Host "脚本路径：$RestoreScriptPath"
        schtasks /create /tn $TaskName /xml $taskXmlPath >$null 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "任务更新成功。" -ForegroundColor Green
        } else {
            Write-Host "任务更新失败。" -ForegroundColor Red
        }
    }
    finally {
        Remove-Item -Path $taskXmlPath -Force -ErrorAction SilentlyContinue
    }
}

function Do-Delete {
    if (-not (Task-Exists)) {
        Write-Host "任务不存在。" -ForegroundColor Red
        return
    }
    schtasks /delete /tn $TaskName /f >$null 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "任务删除成功。" -ForegroundColor Green
    } else {
        Write-Host "任务删除失败。" -ForegroundColor Red
    }
}

function Do-RunNow {
    if (-not (Test-Path $RestoreScriptPath)) {
        Write-Host "脚本不存在：$RestoreScriptPath" -ForegroundColor Red
        return
    }
    Write-Host "正在执行 restore-fonts.ps1 ..." -ForegroundColor Cyan
    & powershell.exe -ExecutionPolicy Bypass -File $RestoreScriptPath
    if ($LASTEXITCODE -eq 0) {
        Write-Host "执行成功。" -ForegroundColor Green
    } else {
        Write-Host "执行失败，退出码：$LASTEXITCODE" -ForegroundColor Red
    }
}

while ($true) {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  字体恢复任务管理器" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Show-Status
    Write-Host "  [1] 添加任务"
    Write-Host "  [2] 更新任务"
    Write-Host "  [3] 删除任务"
    Write-Host "  [4] 立即执行"
    Write-Host "  [0] 退出"
    Write-Host ""
    $key = Read-Host "  请选择"

    Write-Host ""
    switch ($key) {
        "1" { Do-Add }
        "2" { Do-Update }
        "3" { Do-Delete }
        "4" { Do-RunNow }
        "0" { exit 0 }
        default { Write-Host "无效选项。" -ForegroundColor Yellow }
    }
    Write-Host ""
    Write-Host "按任意键继续..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
