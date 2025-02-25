# 获取 .xlsx 文件关联的 ProgID
$progID   = (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\.xlsx").'(default)'
# 获取 GOPATH 环境变量
$gopath   = & go env GOPATH
$programPath  = Join-Path $gopath "bin\addBOM.exe"
# 添加右键菜单
if ($progID) {
    $regPath  = "HKEY_CLASSES_ROOT\$progID\shell\addBOM\command"
    $command  = "`"$programPath`" -i `"%1`""
    # 检查是否以管理员身份运行
#    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
#        Write-Host "Requesting elevated permissions..."
#        $scriptPath = $MyInvocation.MyCommand.Path
#        Start-Process powershell -ArgumentList "-File `"$scriptPath`"" -Verb RunAs
#        exit
#    } else {
        Write-Host "reg add $regPath /ve /d '$command' /f"
        reg add $regPath /ve /d "$command" /f
        # Write-Host "Right-click menu added for .xlsx files."
        Read-Host -Prompt "Press Enter to exit"
#    }
} else {
    Write-Host "Failed to find ProgID for .xlsx files."
    Read-Host -Prompt "Press Enter to exit"
}