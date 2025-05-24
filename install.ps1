# Office-OneClick GUI
# 带用户界面的 Microsoft Office 安装和激活脚本
# 作者: 二进制(sindri) | 博客: blog.nbvil.com
# GitHub: https://github.com/sindricn/OfficeOneClick

# 获取程序路径
try {
    if ($MyInvocation.MyCommand.Path) {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    } elseif ($PSScriptRoot) {
        $scriptDir = $PSScriptRoot
    } else {
        $scriptDir = [System.IO.Path]::GetDirectoryName([System.Reflection.Assembly]::GetExecutingAssembly().Location)
    }
} catch {
    $scriptDir = Get-Location
}

# 如果路径为空，使用当前目录
if ([string]::IsNullOrEmpty($scriptDir)) {
    $scriptDir = Get-Location
}

# 管理员权限检查
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "请以管理员身份运行此脚本！"
    Write-Host "按任意键退出..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# 添加 Windows Forms 支持
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 设置工作目录
$setupUrl = "https://raw.githubusercontent.com/sindricn/OfficeOneClick/main/setup.exe"
$configUrl = "https://raw.githubusercontent.com/sindricn/OfficeOneClick/main/config.xml"
$workDir = Join-Path $env:TEMP "OfficeInstall"
if (!(Test-Path $workDir)) {
    New-Item -Path $workDir -ItemType Directory -Force | Out-Null
}
$setupPath = Join-Path $workDir "setup.exe"
$configPath = Join-Path $workDir "config.xml"
$officePath = ""

# 定义函数

# 下载安装文件
function DownloadFiles {
    param (
        [System.Windows.Forms.RichTextBox]$logBox
    )
    
    # 检查文件是否已存在
    if ((Test-Path $setupPath) -and (Test-Path $configPath)) {
        $logBox.AppendText("安装文件已存在于临时文件夹中。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.SelectionColor = [System.Drawing.Color]::Green
        $logBox.AppendText("可以直接进行下一步操作。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.AppendText("文件位置: $workDir")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return $true
    }
    
    $logBox.AppendText("正在从 GitHub 仓库下载安装文件...")
    $logBox.AppendText([Environment]::NewLine)
    $logBox.ScrollToCaret()
    
    try {
        Invoke-WebRequest -Uri $setupUrl -OutFile $setupPath
        Invoke-WebRequest -Uri $configUrl -OutFile $configPath
        $logBox.SelectionColor = [System.Drawing.Color]::Green
        $logBox.AppendText("安装文件下载成功。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.AppendText("文件已保存至: $workDir")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return $true
    } catch {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("下载安装文件失败！错误详情: $_")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return $false
    }
}

# 下载 Office
function DownloadOffice {
    param (
        [System.Windows.Forms.RichTextBox]$logBox
    )
    
    if (!(Test-Path $setupPath) -or !(Test-Path $configPath)) {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("缺少必要的安装文件，请先下载安装文件。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return $false
    }
    
    # 检查是否已下载 Office 安装包
    $officeFilesExist = $false
    if (Test-Path "$workDir\Office") {
        $officeFiles = Get-ChildItem -Path "$workDir\Office" -Recurse
        if ($officeFiles.Count -gt 5) {  # 假设有超过5个文件表示已下载
            $officeFilesExist = $true
        }
    }
    
    if ($officeFilesExist) {
        $logBox.AppendText("Office 安装包已存在于临时文件夹中。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.SelectionColor = [System.Drawing.Color]::Green
        $logBox.AppendText("可以直接进行下一步安装。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return $true
    }
    
    $logBox.AppendText("正在下载 Office 安装包（这可能需要几分钟）...")
    $logBox.AppendText([Environment]::NewLine)
    $logBox.ScrollToCaret()
    
    try {
        # 确保当前目录是工作目录，以便所有下载的文件都保存在此
        $originalLocation = Get-Location
        Set-Location $workDir
        
        $process = Start-Process -FilePath $setupPath -ArgumentList "/download", $configPath -PassThru -Wait
        if ($process.ExitCode -ne 0) {
            $logBox.SelectionColor = [System.Drawing.Color]::Red
            $logBox.AppendText("Office 安装包下载失败！错误代码: $($process.ExitCode)")
            $logBox.AppendText([Environment]::NewLine)
            $logBox.ScrollToCaret()
            Set-Location $originalLocation
            return $false
        }
        
        Set-Location $originalLocation
        $logBox.SelectionColor = [System.Drawing.Color]::Green
        $logBox.AppendText("Office 安装包下载成功。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.AppendText("所有文件已保存至: $workDir")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return $true
    } catch {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("下载 Office 安装包时出错: $_")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        Set-Location $originalLocation
        return $false
    }
}

# 安装 Office
function InstallOffice {
    param (
        [System.Windows.Forms.RichTextBox]$logBox
    )
    
    if (!(Test-Path $setupPath) -or !(Test-Path $configPath)) {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("缺少必要的安装文件，请先下载安装文件。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return $false
    }
    
    # 检查是否已安装 Office
    $officeInstalled = $false
    if (Test-Path "C:\Program Files\Microsoft Office\Office16") {
        $officeInstalled = $true
        $officePath = "C:\Program Files\Microsoft Office\Office16"
    } elseif (Test-Path "C:\Program Files (x86)\Microsoft Office\Office16") {
        $officeInstalled = $true
        $officePath = "C:\Program Files (x86)\Microsoft Office\Office16"
    }
    
    if ($officeInstalled) {
        $logBox.AppendText("检测到 Office 已安装在系统中。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.SelectionColor = [System.Drawing.Color]::Green
        $logBox.AppendText("安装路径: $officePath")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.AppendText("可以直接进行下一步激活。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return $true
    }
    
    $logBox.AppendText("正在安装 Office（这可能需要几分钟）...")
    $logBox.AppendText([Environment]::NewLine)
    $logBox.ScrollToCaret()
    
    try {
        # 确保当前目录是工作目录
        $originalLocation = Get-Location
        Set-Location $workDir
        
        $process = Start-Process -FilePath $setupPath -ArgumentList "/configure", $configPath -PassThru -Wait
        
        # 检查安装结果
        $installSuccess = $false
        
        # 检查 Office 是否已安装
        Start-Sleep -Seconds 5  # 等待安装完成
        if ((Test-Path "C:\Program Files\Microsoft Office\Office16") -or (Test-Path "C:\Program Files (x86)\Microsoft Office\Office16")) {
            $installSuccess = $true
        }
        
        Set-Location $originalLocation
        
        if ($installSuccess) {
            $logBox.SelectionColor = [System.Drawing.Color]::Green
            $logBox.AppendText("Office 安装成功。")
            $logBox.AppendText([Environment]::NewLine)
            return $true
        } else {
            $logBox.SelectionColor = [System.Drawing.Color]::Red
            $logBox.AppendText("Office 安装失败！错误代码: $($process.ExitCode)")
            $logBox.AppendText([Environment]::NewLine)
            $logBox.ScrollToCaret()
            return $false
        }
    } catch {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("安装 Office 时出错: $_")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        Set-Location $originalLocation
        return $false
    }
}

# 激活 Office
function ActivateOffice {
    param (
        [System.Windows.Forms.RichTextBox]$logBox
    )
    
    $logBox.AppendText("正在查找 Office 安装路径...")
    $logBox.AppendText([Environment]::NewLine)
    $logBox.ScrollToCaret()
    
    $officePath = ""
    if (Test-Path "C:\Program Files\Microsoft Office\Office16") {
        $officePath = "C:\Program Files\Microsoft Office\Office16"
    } elseif (Test-Path "C:\Program Files (x86)\Microsoft Office\Office16") {
        $officePath = "C:\Program Files (x86)\Microsoft Office\Office16"
    } else {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("无法找到 Office 安装路径。请确保 Office 已正确安装。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return $false
    }
    
    # 检查是否已激活
    $isActivated = $false
    try {
        $originalLocation = Get-Location
        Set-Location $officePath
        
        $activationStatus = & cscript ospp.vbs /dstatus 2>&1
        if ($activationStatus -match "-------LICENSED-------") {
            $isActivated = $true
        }
        
        Set-Location $originalLocation
    } catch {
        # 忽略错误，继续尝试激活
        Set-Location $originalLocation
    }
    
    if ($isActivated) {
        $logBox.AppendText("找到 Office 安装路径: $officePath")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.SelectionColor = [System.Drawing.Color]::Green
        $logBox.AppendText("Office 已激活，无需再次激活。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return $true
    }
    
    $logBox.AppendText("找到 Office 安装路径: $officePath")
    $logBox.AppendText([Environment]::NewLine)
    $logBox.AppendText("正在通过 KMS 激活 Office...")
    $logBox.AppendText([Environment]::NewLine)
    $logBox.ScrollToCaret()
    
    try {
        $originalLocation = Get-Location
        Set-Location $officePath
        
        $output = & cscript ospp.vbs /sethst:kms.03k.org 2>&1
        $logBox.AppendText($output)
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        
        $output = & cscript ospp.vbs /act 2>&1
        $logBox.AppendText($output)
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        
        Set-Location $originalLocation
        
        $logBox.SelectionColor = [System.Drawing.Color]::Green
        $logBox.AppendText("Office 激活过程完成。")
        $logBox.AppendText([Environment]::NewLine)
        return $true
    } catch {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("激活 Office 时出错: $_")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        Set-Location $originalLocation
        return $false
    }
}

# 打开临时文件夹
function OpenWorkDirectory {
    param (
        [System.Windows.Forms.RichTextBox]$logBox
    )
    
    if (!(Test-Path $workDir)) {
        New-Item -Path $workDir -ItemType Directory -Force | Out-Null
    }
    
    $logBox.AppendText("正在打开临时文件夹...")
    $logBox.AppendText([Environment]::NewLine)
    
    try {
        Start-Process -FilePath "explorer.exe" -ArgumentList $workDir
        $logBox.SelectionColor = [System.Drawing.Color]::Green
        $logBox.AppendText("已打开临时文件夹: $workDir")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return $true
    } catch {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("打开临时文件夹失败: $_")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return $false
    }
}

# 检查临时文件夹内容
function CheckWorkDirectory {
    param (
        [System.Windows.Forms.RichTextBox]$logBox
    )
    
    $logBox.AppendText("正在检查临时文件夹内容...")
    $logBox.AppendText([Environment]::NewLine)
    $logBox.ScrollToCaret()
    
    if (!(Test-Path $workDir)) {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("临时文件夹不存在。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return
    }
    
    try {
        $files = Get-ChildItem -Path $workDir -Recurse
        
        if ($files.Count -eq 0) {
            $logBox.AppendText("临时文件夹是空的。")
            $logBox.AppendText([Environment]::NewLine)
            $logBox.ScrollToCaret()
            return
        }
        
        $logBox.AppendText("临时文件夹包含以下文件：")
        $logBox.AppendText([Environment]::NewLine)
        $totalSize = 0
        
        foreach ($file in $files) {
            if (!$file.PSIsContainer) {
                $size = $file.Length
                $totalSize += $size
                $sizeStr = ""
                
                if ($size -lt 1KB) {
                    $sizeStr = "$size B"
                } elseif ($size -lt 1MB) {
                    $sizeStr = "{0:N2} KB" -f ($size / 1KB)
                } elseif ($size -lt 1GB) {
                    $sizeStr = "{0:N2} MB" -f ($size / 1MB)
                } else {
                    $sizeStr = "{0:N2} GB" -f ($size / 1GB)
                }
                
                $relativePath = $file.FullName.Substring($workDir.Length + 1)
                $logBox.AppendText("- $relativePath ($sizeStr)")
                $logBox.AppendText([Environment]::NewLine)
            }
        }
        
        $totalSizeStr = ""
        if ($totalSize -lt 1KB) {
            $totalSizeStr = "$totalSize B"
        } elseif ($totalSize -lt 1MB) {
            $totalSizeStr = "{0:N2} KB" -f ($totalSize / 1KB)
        } elseif ($totalSize -lt 1GB) {
            $totalSizeStr = "{0:N2} MB" -f ($totalSize / 1MB)
        } else {
            $totalSizeStr = "{0:N2} GB" -f ($totalSize / 1GB)
        }
        
        $logBox.AppendText([Environment]::NewLine)
        $logBox.AppendText("总计: $($files.Count) 个文件，占用空间 $totalSizeStr")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
    } catch {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("检查临时文件夹内容时出错: $_")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
    }
}

# 离线安装函数
function InstallOfficeOffline {
    param (
        [System.Windows.Forms.RichTextBox]$logBox
    )
    
    $logBox.Clear()
    $logBox.AppendText("正在准备离线安装...")
    $logBox.AppendText([Environment]::NewLine)
    
    # 检查配置文件
    $configPath = "$env:TEMP\OfficeInstall\config.xml"
    if (!(Test-Path $configPath)) {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("错误: 未找到配置文件。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.AppendText("请先运行离线安装配置工具生成配置文件。")
        return $false
    }
    
    # 检查离线安装包
    $offlinePackagePath = "$PSScriptRoot\Office2024"
    if (!(Test-Path $offlinePackagePath)) {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("错误: 未找到离线安装包。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.AppendText("请确保 'Office2024' 文件夹位于与此工具相同的目录中。")
        return $false
    }
    
    # 检查setup.exe
    $setupPath = "$offlinePackagePath\setup.exe"
    if (!(Test-Path $setupPath)) {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("错误: 未找到安装程序。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.AppendText("请确保 'setup.exe' 位于 'Office2024' 文件夹中。")
        return $false
    }
    
    try {
        # 确保当前目录是工作目录
        $originalLocation = Get-Location
        Set-Location $offlinePackagePath
        
        $logBox.AppendText("正在安装 Office...")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.AppendText("使用最新配置文件进行安装...")
        $logBox.AppendText([Environment]::NewLine)
        
        # 使用 CMD 以管理员权限运行安装命令
        $cmdCommand = "cmd.exe /c `"$setupPath`" /configure `"$configPath`""
        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c", "`"$setupPath`"", "/configure", "`"$configPath`"" -Verb RunAs -PassThru -Wait
        
        if ($process.ExitCode -ne 0) {
            $logBox.SelectionColor = [System.Drawing.Color]::Red
            $logBox.AppendText("Office 安装失败！错误代码: $($process.ExitCode)")
            $logBox.AppendText([Environment]::NewLine)
            $logBox.ScrollToCaret()
            Set-Location $originalLocation
            return $false
        }
        
        Set-Location $originalLocation
        $logBox.SelectionColor = [System.Drawing.Color]::Green
        $logBox.AppendText("Office 安装成功。")
        $logBox.AppendText([Environment]::NewLine)
        return $true
    } catch {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("安装 Office 时出错: $_")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        Set-Location $originalLocation
        return $false
    }
}

# 离线安装配置工具
function ShowOfflineConfigUI {
    # 创建配置窗口
    $configForm = New-Object System.Windows.Forms.Form
    $configForm.Text = "Office 离线安装配置"
    $configForm.Size = New-Object System.Drawing.Size(500, 600)
    $configForm.StartPosition = "CenterScreen"
    $configForm.FormBorderStyle = "FixedSingle"
    $configForm.MaximizeBox = $false
    $configForm.BackColor = [System.Drawing.Color]::WhiteSmoke

    # 创建标题标签
    $configTitleLabel = New-Object System.Windows.Forms.Label
    $configTitleLabel.Text = "Office 离线安装配置"
    $configTitleLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 14, [System.Drawing.FontStyle]::Bold)
    $configTitleLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $configTitleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $configTitleLabel.Size = New-Object System.Drawing.Size(450, 30)
    $configForm.Controls.Add($configTitleLabel)

    # 创建说明标签
    $configDescLabel = New-Object System.Windows.Forms.Label
    $configDescLabel.Text = "请选择要安装的 Office 组件："
    $configDescLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $configDescLabel.Location = New-Object System.Drawing.Point(20, 60)
    $configDescLabel.Size = New-Object System.Drawing.Size(450, 20)
    $configForm.Controls.Add($configDescLabel)

    # 创建组件选择组
    $componentsGroup = New-Object System.Windows.Forms.GroupBox
    $componentsGroup.Text = "Office 组件"
    $componentsGroup.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $componentsGroup.Location = New-Object System.Drawing.Point(20, 90)
    $componentsGroup.Size = New-Object System.Drawing.Size(450, 300)
    $configForm.Controls.Add($componentsGroup)

    # 组件说明标签
    $componentsDescLabel = New-Object System.Windows.Forms.Label
    $componentsDescLabel.Text = "请选择您需要安装的 Office 组件（取消勾选将不会安装该组件）"
    $componentsDescLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8.5)
    $componentsDescLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $componentsDescLabel.Location = New-Object System.Drawing.Point(150, 22)
    $componentsDescLabel.Size = New-Object System.Drawing.Size(440, 20)
    $componentsDescLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $componentsGroup.Controls.Add($componentsDescLabel)

    # 定义组件
    $components = @{
        "Word" = @{ Text = "Word (文字处理)"; Checked = $true; Enabled = $true }
        "Excel" = @{ Text = "Excel (电子表格)"; Checked = $true; Enabled = $true }
        "PowerPoint" = @{ Text = "PowerPoint (演示文稿)"; Checked = $true; Enabled = $true }
        "Outlook" = @{ Text = "Outlook (邮件客户端)"; Checked = $false; Enabled = $true }
        "Access" = @{ Text = "Access (数据库)"; Checked = $false; Enabled = $true }
        "Publisher" = @{ Text = "Publisher (不适用于2024版本)"; Checked = $false; Enabled = $false }
        "OneNote" = @{ Text = "OneNote (笔记本)"; Checked = $false; Enabled = $true }
        "Lync" = @{ Text = "Skype for Business"; Checked = $false; Enabled = $true }
        "OneDrive" = @{ Text = "OneDrive (云存储)"; Checked = $false; Enabled = $true }
        "Groove" = @{ Text = "Groove (OneDrive同步)"; Checked = $false; Enabled = $true }
    }

    # 添加"全选"复选框
    $selectAllCheckBox = New-Object System.Windows.Forms.CheckBox
    $selectAllCheckBox.Text = "全选/取消全选"
    $selectAllCheckBox.Location = New-Object System.Drawing.Point(20, 25)
    $selectAllCheckBox.Size = New-Object System.Drawing.Size(120, 20)
    $selectAllCheckBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8, [System.Drawing.FontStyle]::Bold)
    $selectAllCheckBox.Checked = $false
    $componentsGroup.Controls.Add($selectAllCheckBox)

    # 创建组件复选框
    $checkBoxes = @{}
    $leftColumn = 20
    $rightColumn = 320

    # 左右两列显示
    $leftComponents = @("Word", "Excel", "PowerPoint", "Outlook", "Access")
    $rightComponents = @("Publisher", "OneNote", "Lync", "OneDrive", "Groove")

    $yStart = 50
    $yIncrement = 25

    # 创建左侧组件
    for ($i = 0; $i -lt $leftComponents.Count; $i++) {
        $component = $leftComponents[$i]
        if ($components.ContainsKey($component)) {
            $checkBox = New-Object System.Windows.Forms.CheckBox
            $checkBox.Text = $components[$component].Text
            $checkBox.Location = New-Object System.Drawing.Point($leftColumn, ($yStart + $i * $yIncrement))
            $checkBox.Size = New-Object System.Drawing.Size(280, 20)
            $checkBox.Checked = $components[$component].Checked
            $checkBox.Enabled = $components[$component].Enabled
            $checkBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
            $componentsGroup.Controls.Add($checkBox)
            $checkBoxes[$component] = $checkBox
        }
    }

    # 创建右侧组件
    for ($i = 0; $i -lt $rightComponents.Count; $i++) {
        $component = $rightComponents[$i]
        if ($components.ContainsKey($component)) {
            $checkBox = New-Object System.Windows.Forms.CheckBox
            $checkBox.Text = $components[$component].Text
            $checkBox.Location = New-Object System.Drawing.Point($rightColumn, ($yStart + $i * $yIncrement))
            $checkBox.Size = New-Object System.Drawing.Size(280, 20)
            $checkBox.Checked = $components[$component].Checked
            $checkBox.Enabled = $components[$component].Enabled
            $checkBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
            $componentsGroup.Controls.Add($checkBox)
            $checkBoxes[$component] = $checkBox
        }
    }

    # 添加全选/取消全选功能
    $selectAllCheckBox.Add_CheckedChanged({
        $isChecked = $selectAllCheckBox.Checked
        foreach ($component in $components.Keys) {
            if ($checkBoxes.ContainsKey($component)) {
                # 只对启用的组件进行全选/取消全选操作
                if ($components[$component].Enabled) {
                    $checkBoxes[$component].Checked = $isChecked
                }
            }
        }
    })

    # 创建语言选择组
    $languageGroup = New-Object System.Windows.Forms.GroupBox
    $languageGroup.Text = "安装语言"
    $languageGroup.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $languageGroup.Location = New-Object System.Drawing.Point(20, 400)
    $languageGroup.Size = New-Object System.Drawing.Size(450, 80)
    $configForm.Controls.Add($languageGroup)

    # 创建语言选择下拉框
    $languageCombo = New-Object System.Windows.Forms.ComboBox
    $languageCombo.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $languageCombo.Location = New-Object System.Drawing.Point(20, 25)
    $languageCombo.Size = New-Object System.Drawing.Size(410, 25)
    $languageCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $languageGroup.Controls.Add($languageCombo)

    # 添加语言选项
    $languages = @(
        @{Name = "简体中文"; ID = "zh-cn"},
        @{Name = "繁体中文"; ID = "zh-tw"},
        @{Name = "英语"; ID = "en-us"},
        @{Name = "日语"; ID = "ja-jp"},
        @{Name = "韩语"; ID = "ko-kr"}
    )

    foreach ($lang in $languages) {
        $languageCombo.Items.Add($lang.Name)
    }
    $languageCombo.SelectedIndex = 0

    # 创建生成配置按钮
    $generateButton = New-Object System.Windows.Forms.Button
    $generateButton.Text = "生成配置文件"
    $generateButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10, [System.Drawing.FontStyle]::Bold)
    $generateButton.Location = New-Object System.Drawing.Point(20, 490)
    $generateButton.Size = New-Object System.Drawing.Size(450, 40)
    $generateButton.BackColor = [System.Drawing.Color]::DodgerBlue
    $generateButton.ForeColor = [System.Drawing.Color]::White
    $generateButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $configForm.Controls.Add($generateButton)

    # 添加生成配置按钮事件
    $generateButton.Add_Click({
        $selectedComponents = @()
        foreach ($component in $components.Keys) {
            if ($checkBoxes.ContainsKey($component) -and $checkBoxes[$component].Checked -and $components[$component].Enabled) {
                $selectedComponents += $component
            }
        }

        if ($selectedComponents.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("请至少选择一个 Office 组件！", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # 获取选中的语言ID
        $selectedLanguageIndex = $languageCombo.SelectedIndex
        $selectedLanguage = $languages[$selectedLanguageIndex].ID

        # 生成配置文件
        $configContent = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume">
      <Language ID="$selectedLanguage" />
      $(foreach ($component in $selectedComponents) {
        "<ExcludeApp ID=`"$component`" />"
      })
    </Product>
  </Add>
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>
  <Property Name="SharedComputerLicensing" Value="0"/>
  <Property Name="PinIconsToTaskbar" Value="TRUE"/>
  <Property Name="SCLCacheOverride" Value="0"/>
  <Property Name="AUTOACTIVATE" Value="1"/>
  <Property Name="FORCEREBOOT" Value="FALSE"/>
  <Property Name="Display" Value="None"/>
  <Property Name="AcceptEULA" Value="TRUE"/>
  <Property Name="AutoUpgrade" Value="FALSE"/>
</Configuration>
"@

        # 保存配置文件
        $configPath = "$env:TEMP\OfficeInstall\config.xml"
        if (!(Test-Path "$env:TEMP\OfficeInstall")) {
            New-Item -Path "$env:TEMP\OfficeInstall" -ItemType Directory -Force | Out-Null
        }
        $configContent | Out-File -FilePath $configPath -Encoding UTF8

        [System.Windows.Forms.MessageBox]::Show("配置文件已生成！", "成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $configForm.Close()
    })

    # 显示配置窗口
    $configForm.ShowDialog()
}

# 完整安装流程
function CompleteInstallation {
    param (
        [System.Windows.Forms.RichTextBox]$logBox
    )
    
    $logBox.Clear()
    $logBox.AppendText("开始执行完整安装流程...")
    $logBox.AppendText([Environment]::NewLine)
    $logBox.ScrollToCaret()
    
    # 第一步：下载安装文件
    $downloadSuccess = DownloadFiles -logBox $logBox
    if (-not $downloadSuccess) {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("下载安装文件失败，安装流程终止。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return
    }
    
    # 第二步：安装 Office
    $installSuccess = InstallOffice -logBox $logBox
    if (-not $installSuccess) {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("安装 Office 失败，安装流程终止。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return
    }
    
    # 第三步：激活 Office
    $activateSuccess = ActivateOffice -logBox $logBox
    if (-not $activateSuccess) {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("激活 Office 失败，请尝试手动激活。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return
    }
    
    $logBox.SelectionColor = [System.Drawing.Color]::Green
    $logBox.AppendText([Environment]::NewLine)
    $logBox.AppendText("Office 安装和激活已成功完成！")
    $logBox.AppendText([Environment]::NewLine)
    $logBox.AppendText("您现在可以使用 Office 了。")
    $logBox.AppendText([Environment]::NewLine)
    $logBox.ScrollToCaret()
}

# 创建主窗口
$form = New-Object System.Windows.Forms.Form
$form.Text = "Office 一键安装工具 - 作者: 二进制(sindri)"
$form.Size = New-Object System.Drawing.Size(600, 500)  # 调整窗口高度
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::WhiteSmoke

# 创建标题标签
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Microsoft Office 安装工具"
$titleLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::DarkBlue
$titleLabel.Location = New-Object System.Drawing.Point(20, 10)
$titleLabel.Size = New-Object System.Drawing.Size(550, 25)
$form.Controls.Add($titleLabel)

# 创建副标题标签
$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = "作者: 二进制(sindri) | 项目地址: github.com/sindricn/OfficeOneClick"
$subtitleLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8)
$subtitleLabel.ForeColor = [System.Drawing.Color]::Gray
$subtitleLabel.Location = New-Object System.Drawing.Point(20, 35)
$subtitleLabel.Size = New-Object System.Drawing.Size(550, 15)
$form.Controls.Add($subtitleLabel)

# 创建说明标签
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "本工具可以帮助您自动化安装和激活 Microsoft Office。请选择下面的操作："
$descriptionLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$descriptionLabel.Location = New-Object System.Drawing.Point(20, 55)
$descriptionLabel.Size = New-Object System.Drawing.Size(550, 20)
$form.Controls.Add($descriptionLabel)

# 创建一键安装按钮
$completeButton = New-Object System.Windows.Forms.Button
$completeButton.Text = "一键安装&激活`nOffice LTSC 专业增强版 2024"
$completeButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10, [System.Drawing.FontStyle]::Bold)
$completeButton.Location = New-Object System.Drawing.Point(20, 80)
$completeButton.Size = New-Object System.Drawing.Size(270, 45)  # 增加高度以适应三行文本
$completeButton.BackColor = [System.Drawing.Color]::DodgerBlue
$completeButton.ForeColor = [System.Drawing.Color]::White
$completeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($completeButton)

# 创建离线安装按钮
$offlineInstallButton = New-Object System.Windows.Forms.Button
$offlineInstallButton.Text = "离线安装`nOffice LTSC 专业增强版 2024"
$offlineInstallButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10, [System.Drawing.FontStyle]::Bold)
$offlineInstallButton.Location = New-Object System.Drawing.Point(300, 80)
$offlineInstallButton.Size = New-Object System.Drawing.Size(270, 45)  # 调整高度以适应两行文本
$offlineInstallButton.BackColor = [System.Drawing.Color]::LightGreen
$offlineInstallButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($offlineInstallButton)

# 创建分步操作组
$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Text = "分步操作"
$groupBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$groupBox.Location = New-Object System.Drawing.Point(20, 135)  # 调整位置
$groupBox.Size = New-Object System.Drawing.Size(550, 100)
$form.Controls.Add($groupBox)

# 创建下载文件按钮
$downloadFilesButton = New-Object System.Windows.Forms.Button
$downloadFilesButton.Text = "1. 下载安装文件"
$downloadFilesButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$downloadFilesButton.Location = New-Object System.Drawing.Point(20, 20)
$downloadFilesButton.Size = New-Object System.Drawing.Size(250, 25)
$groupBox.Controls.Add($downloadFilesButton)

# 创建安装按钮
$installButton = New-Object System.Windows.Forms.Button
$installButton.Text = "2. 安装 Office"
$installButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$installButton.Location = New-Object System.Drawing.Point(280, 20)
$installButton.Size = New-Object System.Drawing.Size(250, 25)
$groupBox.Controls.Add($installButton)

# 创建激活按钮
$activateButton = New-Object System.Windows.Forms.Button
$activateButton.Text = "3. 激活 Office"
$activateButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$activateButton.Location = New-Object System.Drawing.Point(20, 55)
$activateButton.Size = New-Object System.Drawing.Size(250, 25)
$groupBox.Controls.Add($activateButton)

# 创建自定义配置按钮
$customConfigButton = New-Object System.Windows.Forms.Button
$customConfigButton.Text = "4. 自定义配置"
$customConfigButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$customConfigButton.Location = New-Object System.Drawing.Point(280, 55)
$customConfigButton.Size = New-Object System.Drawing.Size(250, 25)
$groupBox.Controls.Add($customConfigButton)

# 创建文件管理组
$fileGroupBox = New-Object System.Windows.Forms.GroupBox
$fileGroupBox.Text = "临时文件管理"
$fileGroupBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$fileGroupBox.Location = New-Object System.Drawing.Point(20, 245)  # 调整位置
$fileGroupBox.Size = New-Object System.Drawing.Size(550, 50)
$form.Controls.Add($fileGroupBox)

# 创建查看文件列表按钮
$listFilesButton = New-Object System.Windows.Forms.Button
$listFilesButton.Text = "列出临时文件"
$listFilesButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$listFilesButton.Location = New-Object System.Drawing.Point(20, 20)
$listFilesButton.Size = New-Object System.Drawing.Size(170, 25)
$fileGroupBox.Controls.Add($listFilesButton)

# 创建查看临时文件夹按钮
$openWorkDirButton = New-Object System.Windows.Forms.Button
$openWorkDirButton.Text = "查看临时文件夹"
$openWorkDirButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$openWorkDirButton.Location = New-Object System.Drawing.Point(200, 20)
$openWorkDirButton.Size = New-Object System.Drawing.Size(170, 25)
$fileGroupBox.Controls.Add($openWorkDirButton)

# 创建清理按钮
$cleanupButton = New-Object System.Windows.Forms.Button
$cleanupButton.Text = "清理临时文件"
$cleanupButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$cleanupButton.Location = New-Object System.Drawing.Point(380, 20)
$cleanupButton.Size = New-Object System.Drawing.Size(150, 25)
$fileGroupBox.Controls.Add($cleanupButton)

# 创建紧急终止按钮
$emergencyStopButton = New-Object System.Windows.Forms.Button
$emergencyStopButton.Text = "紧急终止"
$emergencyStopButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9, [System.Drawing.FontStyle]::Bold)
$emergencyStopButton.Location = New-Object System.Drawing.Point(20, 305)  # 调整位置
$emergencyStopButton.Size = New-Object System.Drawing.Size(270, 30)
$emergencyStopButton.BackColor = [System.Drawing.Color]::Red
$emergencyStopButton.ForeColor = [System.Drawing.Color]::White
$emergencyStopButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($emergencyStopButton)

# 创建退出按钮
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "退出程序"
$exitButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$exitButton.Location = New-Object System.Drawing.Point(300, 305)  # 调整位置
$exitButton.Size = New-Object System.Drawing.Size(270, 30)
$exitButton.BackColor = [System.Drawing.Color]::LightGray
$form.Controls.Add($exitButton)

# 创建日志框
$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = New-Object System.Drawing.Point(20, 345)  # 调整位置
$logBox.Size = New-Object System.Drawing.Size(550, 80)
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logBox.BackColor = [System.Drawing.Color]::Black
$logBox.ForeColor = [System.Drawing.Color]::White
$logBox.Multiline = $true
$logBox.ReadOnly = $true
$logBox.ScrollBars = "Vertical"
$form.Controls.Add($logBox)

# 设置版权标签
$copyrightLabel = New-Object System.Windows.Forms.Label
$copyrightLabel.Text = "© 2025 Office OneClick Tool | 作者: 二进制(sindri) | 博客: blog.nbvil.com"
$copyrightLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8)
$copyrightLabel.ForeColor = [System.Drawing.Color]::Gray
$copyrightLabel.Location = New-Object System.Drawing.Point(20, 430)  # 调整位置
$copyrightLabel.Size = New-Object System.Drawing.Size(550, 20)
$copyrightLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($copyrightLabel)

# 添加按钮事件
# 一键完成安装
$completeButton.Add_Click({
    CompleteInstallation -logBox $logBox
})

# 离线安装
$offlineInstallButton.Add_Click({
    # 检查离线安装包是否存在
    $offlinePackagePath = "$PSScriptRoot\Office2024"
    if (!(Test-Path $offlinePackagePath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "未找到离线安装包！请确保Office2024文件夹存在。",
            "错误",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # 检查setup.exe
    $setupPath = "$offlinePackagePath\setup.exe"
    if (!(Test-Path $setupPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "未找到安装程序！请确保setup.exe位于Office2024文件夹中。",
            "错误",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    try {
        # 直接在当前窗口启动配置工具
        & "$PSScriptRoot\OfflineConfig.ps1"
    } catch {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("启动离线安装配置工具时出错: $_")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
    }
})

# 自定义配置
$customConfigButton.Add_Click({
    $logBox.Clear()
    $logBox.AppendText("正在启动自定义配置工具...")
    $logBox.AppendText([Environment]::NewLine)
    
    # 从 GitHub 下载 CustomConfig.ps1
    $customConfigUrl = "https://raw.githubusercontent.com/sindricn/OfficeOneClick/main/CustomConfig.ps1"
    $customConfigPath = Join-Path $workDir "CustomConfig.ps1"
    
    try {
        # 确保临时目录存在
        if (!(Test-Path $workDir)) {
            New-Item -Path $workDir -ItemType Directory -Force | Out-Null
        }
        
        # 下载 CustomConfig.ps1
        $logBox.AppendText("正在从 GitHub 下载自定义配置工具...")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        
        # 使用 WebClient 下载文件
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($customConfigUrl, $customConfigPath)
        
        # 验证文件是否下载成功
        if (!(Test-Path $customConfigPath)) {
            throw "下载自定义配置工具失败"
        }
        
        $logBox.SelectionColor = [System.Drawing.Color]::Green
        $logBox.AppendText("自定义配置工具下载成功。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.AppendText("正在启动配置工具...")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        
        # 启动自定义配置工具
        $arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$customConfigPath`""
        Start-Process PowerShell -ArgumentList $arguments -WindowStyle Hidden
        
        $logBox.SelectionColor = [System.Drawing.Color]::Green
        $logBox.AppendText("自定义配置工具已启动。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.AppendText("请在配置工具中选择所需选项，然后生成配置文件。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.AppendText("配置文件将保存在: $workDir")
        $logBox.ScrollToCaret()
    } catch {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("启动自定义配置工具时出错: $_")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.AppendText("请检查网络连接并重试。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
    }
})

# 下载安装文件
$downloadFilesButton.Add_Click({
    $logBox.Clear()
    DownloadFiles -logBox $logBox
})

# 安装 Office
$installButton.Add_Click({
    $logBox.Clear()
    InstallOffice -logBox $logBox
})

# 激活 Office
$activateButton.Add_Click({
    $logBox.Clear()
    ActivateOffice -logBox $logBox
})

# 显示临时文件夹路径
$openWorkDirButton.Add_Click({
    $logBox.Clear()
    OpenWorkDirectory -logBox $logBox
})

# 列出临时文件
$listFilesButton.Add_Click({
    $logBox.Clear()
    CheckWorkDirectory -logBox $logBox
})

# 清理临时文件
$cleanupButton.Add_Click({
    $logBox.Clear()
    $logBox.AppendText("正在清理临时文件...")
    $logBox.AppendText([Environment]::NewLine)
    Remove-Item -Path $workDir\* -Recurse -Force -ErrorAction SilentlyContinue
    $logBox.SelectionColor = [System.Drawing.Color]::Green
    $logBox.AppendText("清理完成。文件夹结构保留，但内容已删除。")
    $logBox.AppendText([Environment]::NewLine)
})

# 退出程序
$exitButton.Add_Click({
    $form.Close()
})

# 修改紧急终止按钮事件
$emergencyStopButton.Add_Click({
    # 获取所有需要终止的进程
    $processesToStop = @(
        Get-Process | Where-Object { $_.ProcessName -match 'setup|office|msiexec' -or $_.MainWindowTitle -match 'Office|Microsoft' }
        Get-Process | Where-Object { $_.ProcessName -match 'powershell' -and $_.MainWindowTitle -match 'Office' }
    ) | Select-Object -Unique

    if ($processesToStop.Count -gt 0) {
        $processNames = $processesToStop | ForEach-Object { $_.ProcessName }
        $processList = $processNames -join "`n"
        
        $result = [System.Windows.Forms.MessageBox]::Show(
            "检测到以下Office相关进程正在运行：`n`n$processList`n`n是否要终止这些进程？",
            "确认终止",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $logBox.AppendText("正在终止Office相关进程...`n")
            
            foreach ($process in $processesToStop) {
                try {
                    $processName = $process.ProcessName
                    $processId = $process.Id
                    $logBox.AppendText("正在终止进程: $processName (PID: $processId)`n")
                    
                    # 尝试正常终止进程
                    $process.CloseMainWindow()
                    Start-Sleep -Milliseconds 500
                    
                    # 如果进程仍在运行，强制终止
                    if (!$process.HasExited) {
                        $process | Stop-Process -Force
                        $logBox.AppendText("已强制终止进程: $processName`n")
                    } else {
                        $logBox.AppendText("已终止进程: $processName`n")
                    }
                } catch {
                    $logBox.AppendText("终止进程 $($process.ProcessName) 时出错: $_`n")
                }
            }
            
            # 清理临时文件
            try {
                if (Test-Path "$env:TEMP\OfficeInstall") {
                    Remove-Item "$env:TEMP\OfficeInstall" -Recurse -Force
                    $logBox.AppendText("已清理临时文件`n")
                }
            } catch {
                $logBox.AppendText("清理临时文件时出错: $_`n")
            }
            
            $logBox.AppendText("所有Office相关进程已终止`n")
            [System.Windows.Forms.MessageBox]::Show(
                "所有Office相关进程已终止。",
                "操作完成",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "未检测到正在运行的Office相关进程。",
            "提示",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

# 显示欢迎信息
$logBox.AppendText("欢迎使用 Office 一键安装工具！")
$logBox.AppendText([Environment]::NewLine)
$logBox.AppendText("请选择上方的操作按钮开始安装。")
$logBox.AppendText([Environment]::NewLine)
$logBox.AppendText("推荐使用【一键完成安装】以自动执行所有步骤。")
$logBox.AppendText([Environment]::NewLine)
$logBox.AppendText("临时文件将保存在: $workDir")
$logBox.AppendText([Environment]::NewLine)

# 显示窗口
$form.ShowDialog()

