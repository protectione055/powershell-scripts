# PowerShell-Scripts

## 文件说明

这个仓库提供一些 PowerShell 脚本：

- **`CheckAndDeleteRecent.ps1`**：用于检查并清理指定目录下最近访问的记录。

## CleanRecent

Windows 11 最近访问经常会出现最近访问过的学习资料，但是又不提供排除某个目录的功能。因此这个脚本用于清除 Windows 系统中最近打开文件中，指定目录下的记录。通过自定义的目标路径列表，脚本会检查 Recent 文件夹下的快捷方式文件是否指向指定路径的内容，并自动删除符合条件的快捷方式。

### 功能特性

- **自动清理最近记录**：扫描 `Recent` 文件夹下的所有快捷方式，并根据目标路径列表删除指定路径的快捷方式。
- **支持中文路径**：正确处理中文路径匹配。
- **前缀匹配**：快捷方式的目标路径以目标路径列表中的某一项为前缀时，快捷方式会被删除。
- **计划任务支持**：可通过 Windows 计划任务后台自动执行，无窗口弹出。

---

### 使用方法

#### 1. 配置目标路径列表

在 `CheckAndDeleteRecent.ps1` 脚本中，找到以下部分，并将目标路径替换为你需要清理的路径列表：

```powershell
$TargetPaths = @(
    "F:\path1",      # 示例路径 1
    "D:\path2"       # 示例路径 2
)
```

### 2. 打开终端测试脚本是否正常工作

将脚本文件 `CheckAndDeleteRecent.ps1` 复制到指定目录（如 `C:\Scripts`）

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\CheckAndDeleteRecent.ps1"
```

### 3. 配置计划任务

- 打开 **任务计划程序**，创建一个新的基本任务命名为 `CleanRecent`。
- 在 **触发器** 中设置任务的触发条件。
- 在 **操作** 中设置启动程序 `pwsh`，添加参数 `-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Scripts\CheckAndDeleteRecent.ps1"`
- 在 **条件** 中取消勾选 `启动任务时，计算机是否已连接到电源`，以及 `启动任务时，只有计算机已空闲`。
- 完成后，任务会在被触发时自动执行。
