# Office-OneClick GUI
# 带用户界面的 Microsoft Office 安装和激活脚本
# 作者: 二进制(sindri) | 博客: blog.nbvil.com
# GitHub: https://github.com/sindricn/OfficeOneClick

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
$workDir = "$env:TEMP\OfficeInstall"
if (!(Test-Path $workDir)) {
    New-Item -Path $workDir -ItemType Directory -Force | Out-Null
}

# 定义变量
$setupUrl = "https://raw.githubusercontent.com/sindricn/OfficeOneClick/main/setup.exe"
$configUrl = "https://raw.githubusercontent.com/sindricn/OfficeOneClick/main/config.xml"
$setupPath = "$workDir\setup.exe"
$configPath = "$workDir\config.xml"
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
        $logBox.ScrollToCaret()
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
        $logBox.ScrollToCaret()
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

# 完整安装流程
function CompleteInstallation {
    param (
        [System.Windows.Forms.RichTextBox]$logBox
    )
    
    $logBox.Clear()
    $logBox.AppendText("开始执行完整安装流程...")
    $logBox.AppendText([Environment]::NewLine)
    $logBox.ScrollToCaret()
    
    $downloadFilesSuccess = DownloadFiles -logBox $logBox
    if (-not $downloadFilesSuccess) {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("下载安装文件失败，安装过程终止。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return
    }
    
    $downloadOfficeSuccess = DownloadOffice -logBox $logBox
    if (-not $downloadOfficeSuccess) {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("下载 Office 安装包失败，安装过程终止。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return
    }
    
    $installOfficeSuccess = InstallOffice -logBox $logBox
    if (-not $installOfficeSuccess) {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("安装 Office 失败，安装过程终止。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return
    }
    
    $activateOfficeSuccess = ActivateOffice -logBox $logBox
    if (-not $activateOfficeSuccess) {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.AppendText("激活 Office 失败，请尝试手动激活。")
        $logBox.AppendText([Environment]::NewLine)
        $logBox.ScrollToCaret()
        return
    }
    
    # 不自动清理，方便用户查看文件
    $logBox.SelectionColor = [System.Drawing.Color]::Green
    $logBox.AppendText([Environment]::NewLine)
    $logBox.AppendText("Office 安装和激活已成功完成！您现在可以使用 Microsoft Office 产品了。")
    $logBox.AppendText([Environment]::NewLine)
    $logBox.AppendText("临时安装文件保留在: $workDir，可以通过「查看临时文件夹」按钮查看。")
    $logBox.AppendText([Environment]::NewLine)
    $logBox.ScrollToCaret()
}

# 创建主窗口
$form = New-Object System.Windows.Forms.Form
$form.Text = "Office 一键安装工具 - 作者: 二进制(sindri)"
$form.Size = New-Object System.Drawing.Size(600, 520)  # 稍微增加窗口高度
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::WhiteSmoke

# 创建标题标签
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Microsoft Office 安装工具"
$titleLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::DarkBlue
$titleLabel.Location = New-Object System.Drawing.Point(20, 20)
$titleLabel.Size = New-Object System.Drawing.Size(550, 30)
$form.Controls.Add($titleLabel)

# 创建副标题标签
$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = "作者: 二进制(sindri) | 项目地址: github.com/sindricn/OfficeOneClick"
$subtitleLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8)
$subtitleLabel.ForeColor = [System.Drawing.Color]::Gray
$subtitleLabel.Location = New-Object System.Drawing.Point(20, 50)
$subtitleLabel.Size = New-Object System.Drawing.Size(550, 20)
$form.Controls.Add($subtitleLabel)

# 创建说明标签
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "本工具可以帮助您自动化安装和激活 Microsoft Office。请选择下面的操作："
$descriptionLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$descriptionLabel.Location = New-Object System.Drawing.Point(20, 70)
$descriptionLabel.Size = New-Object System.Drawing.Size(550, 20)
$form.Controls.Add($descriptionLabel)

# 创建一键安装按钮
$completeButton = New-Object System.Windows.Forms.Button
$completeButton.Text = "一键完成安装"
$completeButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10, [System.Drawing.FontStyle]::Bold)
$completeButton.Location = New-Object System.Drawing.Point(20, 100)
$completeButton.Size = New-Object System.Drawing.Size(550, 40)
$completeButton.BackColor = [System.Drawing.Color]::DodgerBlue
$completeButton.ForeColor = [System.Drawing.Color]::White
$completeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($completeButton)

# 创建分步操作组
$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Text = "分步操作"
$groupBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$groupBox.Location = New-Object System.Drawing.Point(20, 150)
$groupBox.Size = New-Object System.Drawing.Size(550, 150)
$form.Controls.Add($groupBox)

# 创建下载文件按钮
$downloadFilesButton = New-Object System.Windows.Forms.Button
$downloadFilesButton.Text = "1. 下载安装文件"
$downloadFilesButton.Location = New-Object System.Drawing.Point(20, 30)
$downloadFilesButton.Size = New-Object System.Drawing.Size(150, 30)
$groupBox.Controls.Add($downloadFilesButton)

# 创建下载 Office 按钮
$downloadOfficeButton = New-Object System.Windows.Forms.Button
$downloadOfficeButton.Text = "2. 下载 Office"
$downloadOfficeButton.Location = New-Object System.Drawing.Point(20, 70)
$downloadOfficeButton.Size = New-Object System.Drawing.Size(150, 30)
$groupBox.Controls.Add($downloadOfficeButton)

# 创建安装按钮
$installButton = New-Object System.Windows.Forms.Button
$installButton.Text = "3. 安装 Office"
$installButton.Location = New-Object System.Drawing.Point(190, 30)
$installButton.Size = New-Object System.Drawing.Size(150, 30)
$groupBox.Controls.Add($installButton)

# 创建激活按钮
$activateButton = New-Object System.Windows.Forms.Button
$activateButton.Text = "4. 激活 Office"
$activateButton.Location = New-Object System.Drawing.Point(190, 70)
$activateButton.Size = New-Object System.Drawing.Size(150, 30)
$groupBox.Controls.Add($activateButton)

# 创建临时文件夹操作按钮
$openWorkDirButton = New-Object System.Windows.Forms.Button
$openWorkDirButton.Text = "查看临时文件夹"
$openWorkDirButton.Location = New-Object System.Drawing.Point(360, 30)
$openWorkDirButton.Size = New-Object System.Drawing.Size(150, 30)
$groupBox.Controls.Add($openWorkDirButton)

# 创建清理按钮
$cleanupButton = New-Object System.Windows.Forms.Button
$cleanupButton.Text = "5. 清理临时文件"
$cleanupButton.Location = New-Object System.Drawing.Point(360, 70)
$cleanupButton.Size = New-Object System.Drawing.Size(150, 30)
$groupBox.Controls.Add($cleanupButton)

# 创建文件管理组
$fileGroupBox = New-Object System.Windows.Forms.GroupBox
$fileGroupBox.Text = "临时文件管理"
$fileGroupBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$fileGroupBox.Location = New-Object System.Drawing.Point(20, 310)
$fileGroupBox.Size = New-Object System.Drawing.Size(550, 50)
$form.Controls.Add($fileGroupBox)

# 创建查看文件列表按钮
$listFilesButton = New-Object System.Windows.Forms.Button
$listFilesButton.Text = "列出临时文件"
$listFilesButton.Location = New-Object System.Drawing.Point(20, 20)
$listFilesButton.Size = New-Object System.Drawing.Size(150, 25)
$fileGroupBox.Controls.Add($listFilesButton)

# 创建临时文件夹路径按钮
$showPathButton = New-Object System.Windows.Forms.Button
$showPathButton.Text = "显示临时文件夹路径"
$showPathButton.Location = New-Object System.Drawing.Point(190, 20)
$showPathButton.Size = New-Object System.Drawing.Size(150, 25)
$fileGroupBox.Controls.Add($showPathButton)

# 创建退出按钮
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "退出程序"
$exitButton.Location = New-Object System.Drawing.Point(360, 20)
$exitButton.Size = New-Object System.Drawing.Size(150, 25)
$exitButton.BackColor = [System.Drawing.Color]::LightGray
$fileGroupBox.Controls.Add($exitButton)

# 创建日志框
$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = New-Object System.Drawing.Point(20, 370)
$logBox.Size = New-Object System.Drawing.Size(550, 90)  # 增加日志框高度
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
$copyrightLabel.Location = New-Object System.Drawing.Point(20, 440)  # 调整位置
$copyrightLabel.Size = New-Object System.Drawing.Size(550, 20)
$copyrightLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($copyrightLabel)

# 添加按钮事件

# 一键完成安装
$completeButton.Add_Click({
    CompleteInstallation -logBox $logBox
})

# 下载安装文件
$downloadFilesButton.Add_Click({
    $logBox.Clear()
    DownloadFiles -logBox $logBox
})

# 下载 Office
$downloadOfficeButton.Add_Click({
    $logBox.Clear()
    DownloadOffice -logBox $logBox
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

# 打开临时文件夹
$openWorkDirButton.Add_Click({
    $logBox.Clear()
    OpenWorkDirectory -logBox $logBox
})

# 显示临时文件夹路径
$showPathButton.Add_Click({
    $logBox.Clear()
    $logBox.AppendText("临时文件夹路径: $workDir\n")
    
    # 将路径复制到剪贴板
    [System.Windows.Forms.Clipboard]::SetText($workDir)
    $logBox.AppendText("已将路径复制到剪贴板。\n")
})

# 列出临时文件
$listFilesButton.Add_Click({
    $logBox.Clear()
    CheckWorkDirectory -logBox $logBox
})

# 清理临时文件
$cleanupButton.Add_Click({
    $logBox.Clear()
    $logBox.AppendText("正在清理临时文件...\n")
    Remove-Item -Path $workDir\* -Recurse -Force -ErrorAction SilentlyContinue
    $logBox.SelectionColor = [System.Drawing.Color]::Green
    $logBox.AppendText("清理完成。文件夹结构保留，但内容已删除。\n")
})

# 退出程序
$exitButton.Add_Click({
    $form.Close()
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
