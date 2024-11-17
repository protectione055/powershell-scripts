<#
.SYNOPSIS
此脚本删除 Windows 最近文件夹中指向指定目标路径的快捷方式文件。

.DESCRIPTION
脚本遍历当前用户最近文件夹中的所有快捷方式 (.lnk) 文件，并删除那些指向 $TargetPaths 数组中目标路径的快捷方式。它使用 WScript.Shell COM 对象解析每个快捷方式的目标路径。

.PARAMETER $CurrentUser
从环境变量 $env:USERNAME 获取的当前用户名。

.PARAMETER $RecentDir
当前用户最近文件夹的路径。

.PARAMETER $TargetPaths
目标路径的数组。指向这些路径的快捷方式将被删除。

.EXAMPLE
# 示例：删除所有指向 D:\Downloads 的快捷方式
$TargetPaths = @("D:\Downloads")

.NOTES
- 确保以适当的权限运行脚本，以访问和删除最近文件夹中的文件。
- 脚本使用 WScript.Shell COM 对象解析快捷方式目标路径。

#>
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$CurrentUser = $env:USERNAME
$RecentDir = "C:\Users\$CurrentUser\AppData\Roaming\Microsoft\Windows\Recent"
$TargetPaths = @(
    # 例子：删除所有指向 D:\Downloads 的快捷方式
    # "D:\Downloads",
)


# 获取最近目录下的所有快捷方式文件
$Shortcuts = Get-ChildItem -Path $RecentDir -Filter *.lnk

foreach ($Shortcut in $Shortcuts) {
    # 解析快捷方式的目标路径
    $WScriptShell = New-Object -ComObject WScript.Shell
    $ShortcutObject = $WScriptShell.CreateShortcut($Shortcut.FullName)
    $ShortcutPath = $ShortcutObject.TargetPath

    # 检查目标路径是否为空
    if (![string]::IsNullOrWhiteSpace($ShortcutPath)) {
        # 检查目标路径是否包含在列表中的任何路径
        foreach ($TargetPath in $TargetPaths) {
            Write-Host "Checking shortcut: $Shortcut.FullName -> $ShortcutPath"
            if ($ShortcutPath.StartsWith($TargetPath)) {
                Write-Host "Deleting shortcut: $Shortcut.FullName -> $ShortcutPath"
                Remove-Item -Path $Shortcut.FullName -Force
                break
            }
            else {
                Write-Host "Skipping $Shortcut.FullName -> $ShortcutPath for not matching $TargetPath"
            }
        }
    }
    else {
        Write-Host "Skipping shortcut with no target: $Shortcut.FullName"
    }
}
