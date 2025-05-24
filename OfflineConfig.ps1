# Office-OneClick 离线安装配置工具
# 允许用户选择语言和组件，然后生成自定义config.xml
# 适用于 OfficeOneClick 工具的离线安装

# 管理员权限检查
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "请以管理员身份运行此脚本！"
    Write-Host "按任意键退出..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# 添加Windows Forms支持
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 创建主窗体
$form = New-Object System.Windows.Forms.Form
$form.Text = "Office 离线安装配置工具"
$form.Size = New-Object System.Drawing.Size(650, 550)  # 减小窗口高度
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.BackColor = [System.Drawing.Color]::White

# 添加标题
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Office 离线安装配置工具"
$titleLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.Location = New-Object System.Drawing.Point(20, 15)  # 减小上边距
$titleLabel.Size = New-Object System.Drawing.Size(600, 25)  # 减小高度
$form.Controls.Add($titleLabel)

# 添加描述
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "请选择您需要的语言及组件，生成自定义配置文件"
$descriptionLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$descriptionLabel.Location = New-Object System.Drawing.Point(20, 45)  # 调整位置
$descriptionLabel.Size = New-Object System.Drawing.Size(600, 20)
$form.Controls.Add($descriptionLabel)

# 添加版本提示标签
$versionLabel = New-Object System.Windows.Forms.Label
$versionLabel.Text = "当前版本: Office LTSC 专业增强版 2024"
$versionLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9, [System.Drawing.FontStyle]::Bold)
$versionLabel.ForeColor = [System.Drawing.Color]::DarkBlue
$versionLabel.Location = New-Object System.Drawing.Point(20, 20)
$versionLabel.Size = New-Object System.Drawing.Size(600, 20)
$form.Controls.Add($versionLabel)

# 添加语言选项
$languages = [ordered]@{
    "简体中文" = "zh-cn"
    "繁体中文(台湾)" = "zh-tw"
    "英语(美国)" = "en-us"
    "英语(英国)" = "en-gb"
    "日语" = "ja-jp"
    "韩语" = "ko-kr"
}

# 创建组件复选框
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

# 创建语言选择组
$languageGroupBox = New-Object System.Windows.Forms.GroupBox
$languageGroupBox.Text = "语言选项"
$languageGroupBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$languageGroupBox.Location = New-Object System.Drawing.Point(20, 70)
$languageGroupBox.Size = New-Object System.Drawing.Size(600, 180)
$form.Controls.Add($languageGroupBox)

# 主要语言标签
$mainLangLabel = New-Object System.Windows.Forms.Label
$mainLangLabel.Text = "主要语言:"
$mainLangLabel.Location = New-Object System.Drawing.Point(10, 25)
$mainLangLabel.Size = New-Object System.Drawing.Size(120, 20)
$languageGroupBox.Controls.Add($mainLangLabel)

# 主语言选择下拉框
$mainLangComboBox = New-Object System.Windows.Forms.ComboBox
$mainLangComboBox.Location = New-Object System.Drawing.Point(130, 23)
$mainLangComboBox.Size = New-Object System.Drawing.Size(150, 25)
$mainLangComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$mainLangComboBox.DropDownWidth = 180
$languageGroupBox.Controls.Add($mainLangComboBox)

# 添加语言选项到下拉框
$mainLangComboBox.Items.Clear()
foreach ($language in $languages.Keys) {
    $mainLangComboBox.Items.Add($language)
}
# 设置简体中文为默认语言
$mainLangComboBox.SelectedItem = "简体中文"

# 添加主语言选择事件
$mainLangComboBox.Add_SelectedIndexChanged({
    $selectedLanguage = $mainLangComboBox.SelectedItem
    if ($selectedLanguage) {
        # 禁用其他语言中与主语言相同的选项
        foreach ($langKey in $langCheckBoxes.Keys) {
            if ($langKey -eq $selectedLanguage) {
                $langCheckBoxes[$langKey].Checked = $false
                $langCheckBoxes[$langKey].Enabled = $false
            } else {
                $langCheckBoxes[$langKey].Enabled = $true
            }
        }
    }
})

# 其他语言标签
$otherLangLabel = New-Object System.Windows.Forms.Label
$otherLangLabel.Text = "其他语言:"
$otherLangLabel.Location = New-Object System.Drawing.Point(10, 55)
$otherLangLabel.Size = New-Object System.Drawing.Size(120, 20)
$languageGroupBox.Controls.Add($otherLangLabel)

# 添加语言复选框
$langCheckBoxes = @{}
$langYStart = 80
$langYIncrement = 25
$langsPerColumn = 4
$langColumnWidth = 140

$y = $langYStart
$column = 0
foreach ($lang in $languages.Keys) {
    $langCheckBox = New-Object System.Windows.Forms.CheckBox
    $langCheckBox.Text = $lang
    $langCheckBox.Location = New-Object System.Drawing.Point((10 + $column * $langColumnWidth), $y)
    $langCheckBox.Size = New-Object System.Drawing.Size(130, 20)
    $langCheckBox.Checked = $false
    $langCheckBox.Tag = $languages[$lang]
    $langCheckBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8)
    $languageGroupBox.Controls.Add($langCheckBox)
    $langCheckBoxes[$lang] = $langCheckBox
    
    $y += $langYIncrement
    if ($y -ge ($langYStart + $langsPerColumn * $langYIncrement)) {
        $y = $langYStart
        $column++
    }
}

# 组件选择组
$componentsGroupBox = New-Object System.Windows.Forms.GroupBox
$componentsGroupBox.Text = "Office 组件"
$componentsGroupBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$componentsGroupBox.Location = New-Object System.Drawing.Point(20, 230)  # 调整位置
$componentsGroupBox.Size = New-Object System.Drawing.Size(600, 200)  # 减小高度
$form.Controls.Add($componentsGroupBox)

# 组件说明标签
$componentsDescLabel = New-Object System.Windows.Forms.Label
$componentsDescLabel.Text = "请选择您需要安装的 Office 组件（取消勾选将不会安装该组件）"
$componentsDescLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8.5)
$componentsDescLabel.ForeColor = [System.Drawing.Color]::DarkBlue
$componentsDescLabel.Location = New-Object System.Drawing.Point(150, 22)
$componentsDescLabel.Size = New-Object System.Drawing.Size(440, 20)
$componentsDescLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$componentsGroupBox.Controls.Add($componentsDescLabel)

# 创建组件复选框
$checkBoxes = @{}
$leftColumn = 20
$rightColumn = 320

# 添加"全选"复选框
$selectAllCheckBox = New-Object System.Windows.Forms.CheckBox
$selectAllCheckBox.Text = "全选/取消全选"
$selectAllCheckBox.Location = New-Object System.Drawing.Point(20, 25)
$selectAllCheckBox.Size = New-Object System.Drawing.Size(120, 20)
$selectAllCheckBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8, [System.Drawing.FontStyle]::Bold)
$selectAllCheckBox.Checked = $false
$componentsGroupBox.Controls.Add($selectAllCheckBox)

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
        $componentsGroupBox.Controls.Add($checkBox)
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
        $componentsGroupBox.Controls.Add($checkBox)
        $checkBoxes[$component] = $checkBox
    }
}

# 添加全选/取消全选功能
$selectAllCheckBox.Add_CheckedChanged({
    foreach ($component in $components.Keys) {
        if ($checkBoxes.ContainsKey($component)) {
            # 检查组件是否可用
            if ($components[$component].Enabled) {
                $checkBoxes[$component].Checked = $selectAllCheckBox.Checked
            }
        }
    }
})

# 按钮区域
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(20, 440)  # 调整位置
$buttonPanel.Size = New-Object System.Drawing.Size(600, 50)
$form.Controls.Add($buttonPanel)

# 生成配置按钮
$generateButton = New-Object System.Windows.Forms.Button
$generateButton.Text = "生成配置文件"
$generateButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10, [System.Drawing.FontStyle]::Bold)
$generateButton.Location = New-Object System.Drawing.Point(0, 5)
$generateButton.Size = New-Object System.Drawing.Size(195, 40)  # 减小宽度
$generateButton.BackColor = [System.Drawing.Color]::DodgerBlue
$generateButton.ForeColor = [System.Drawing.Color]::White
$generateButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonPanel.Controls.Add($generateButton)

# 开始安装按钮
$installButton = New-Object System.Windows.Forms.Button
$installButton.Text = "开始离线安装"
$installButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10, [System.Drawing.FontStyle]::Bold)
$installButton.Location = New-Object System.Drawing.Point(205, 5)
$installButton.Size = New-Object System.Drawing.Size(195, 40)  # 减小宽度
$installButton.BackColor = [System.Drawing.Color]::ForestGreen
$installButton.ForeColor = [System.Drawing.Color]::White
$installButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonPanel.Controls.Add($installButton)

# 退出按钮
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "退出离线安装"
$exitButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
$exitButton.Location = New-Object System.Drawing.Point(410, 5)
$exitButton.Size = New-Object System.Drawing.Size(190, 40)  # 减小宽度
$exitButton.BackColor = [System.Drawing.Color]::LightGray
$exitButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonPanel.Controls.Add($exitButton)

# 添加状态栏
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "准备就绪"
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# 修改生成配置文件函数
function GenerateConfigFile {
    # 获取选择的主语言
    $selectedLanguage = $mainLangComboBox.SelectedItem
    if ($null -eq $selectedLanguage) {
        [System.Windows.Forms.MessageBox]::Show(
            "请选择主要语言！",
            "错误",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
    $mainLanguageValue = $languages[$selectedLanguage]
    
    # 获取选中的其他语言
    $additionalLanguages = @()
    foreach ($langKey in $langCheckBoxes.Keys) {
        if ($langCheckBoxes[$langKey].Checked) {
            $langValue = $languages[$langKey]
            # 避免添加重复的语言
            if ($langValue -ne $mainLanguageValue) {
                $additionalLanguages += $langValue
            }
        }
    }
    
    # 检查离线安装包
    $offlinePackagePath = "$PSScriptRoot\Office2024"
    if (!(Test-Path $offlinePackagePath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "未找到离线安装包！请确保Office2024文件夹存在。",
            "错误",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
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
        return $false
    }
    
    # 创建XML文档
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDeclaration = $xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", $null)
    $xmlDoc.AppendChild($xmlDeclaration) | Out-Null
    
    # 创建根节点
    $configNode = $xmlDoc.CreateElement("Configuration")
    $configNode.SetAttribute("ID", [guid]::NewGuid().ToString())
    $xmlDoc.AppendChild($configNode) | Out-Null
    
    # 添加Add节点
    $addNode = $xmlDoc.CreateElement("Add")
    $addNode.SetAttribute("OfficeClientEdition", "64")
    $addNode.SetAttribute("Channel", "PerpetualVL2024")
    $addNode.SetAttribute("SourcePath", $offlinePackagePath)
    $addNode.SetAttribute("AllowCdnFallback", "FALSE")
    $configNode.AppendChild($addNode) | Out-Null
    
    # 添加Product节点
    $productNode = $xmlDoc.CreateElement("Product")
    $productNode.SetAttribute("ID", "ProPlus2024Volume")
    $productNode.SetAttribute("PIDKEY", "XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB")
    $addNode.AppendChild($productNode) | Out-Null
    
    # 添加主语言节点
    $languageNode = $xmlDoc.CreateElement("Language")
    $languageNode.SetAttribute("ID", $mainLanguageValue)
    $productNode.AppendChild($languageNode) | Out-Null
    
    # 添加其他语言节点
    foreach ($langCode in $additionalLanguages) {
        $addLangNode = $xmlDoc.CreateElement("Language")
        $addLangNode.SetAttribute("ID", $langCode)
        $productNode.AppendChild($addLangNode) | Out-Null
    }
    
    # 添加排除的应用程序
    foreach ($component in $components.Keys) {
        if ($checkBoxes.ContainsKey($component) -and -not $checkBoxes[$component].Checked -and $components[$component].Enabled) {
            $excludeNode = $xmlDoc.CreateElement("ExcludeApp")
            $excludeNode.SetAttribute("ID", $component)
            $productNode.AppendChild($excludeNode) | Out-Null
        }
    }
    
    # 添加属性节点
    $properties = @{
        "SharedComputerLicensing" = "0"
        "DeviceBasedLicensing" = "0"
        "SCLCacheOverride" = "0"
        "FORCEAPPSHUTDOWN" = "TRUE"
        "AUTOACTIVATE" = "1"
        "FORCEREBOOT" = "FALSE"
        "Display" = "None"
        "AcceptEULA" = "TRUE"
        "AutoUpgrade" = "FALSE"
    }
    
    foreach ($prop in $properties.Keys) {
        $propNode = $xmlDoc.CreateElement("Property")
        $propNode.SetAttribute("Name", $prop)
        $propNode.SetAttribute("Value", $properties[$prop])
        $configNode.AppendChild($propNode) | Out-Null
    }
    
    # 添加更新节点
    $updatesNode = $xmlDoc.CreateElement("Updates")
    $updatesNode.SetAttribute("Enabled", "FALSE")
    $configNode.AppendChild($updatesNode) | Out-Null
    
    # 添加RemoveMSI节点
    $removeMsiNode = $xmlDoc.CreateElement("RemoveMSI")
    $configNode.AppendChild($removeMsiNode) | Out-Null
    
    # 保存XML文件
    try {
        # 直接保存到Office2024文件夹
        $configPath = "$offlinePackagePath\config.xml"
        $xmlDoc.Save($configPath)
        $statusLabel.Text = "配置文件已生成: $configPath"
        
        # 提示用户
        [System.Windows.Forms.MessageBox]::Show(
            "配置文件已成功生成！`n`n路径: $configPath`n`n请点击'开始离线安装'按钮开始安装。",
            "成功",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        return $true
    }
    catch {
        $statusLabel.Text = "生成配置文件时出错: $_"
        
        [System.Windows.Forms.MessageBox]::Show(
            "生成配置文件时出错: $_",
            "错误",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        
        return $false
    }
}

# 修改按钮事件
$generateButton.Add_Click({
    GenerateConfigFile
})

# 修改安装按钮事件
$installButton.Add_Click({
    # 检查离线安装包
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

    # 检查配置文件
    $configPath = "$offlinePackagePath\config.xml"
    if (!(Test-Path $configPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "请先生成配置文件！",
            "提示",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    try {
        # 启动安装程序
        $process = Start-Process -FilePath $setupPath -ArgumentList "/configure `"$configPath`"" -Verb RunAs -PassThru
        
        # 等待进程启动
        Start-Sleep -Seconds 2
        
        if ($process.HasExited) {
            throw "安装程序启动失败，退出代码: $($process.ExitCode)"
        }
        
        # 显示成功消息
        [System.Windows.Forms.MessageBox]::Show(
            "Office 安装程序已成功启动！`n`n安装过程可能需要一些时间，请耐心等待。",
            "安装已启动",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        # 关闭配置窗口
        $form.Close()
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "启动安装程序时出错:`n`n$_",
            "错误",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

$exitButton.Add_Click({
    $form.Close()
})

# 隐藏调试信息输出
$Host.UI.RawUI.WindowTitle = "Office 离线安装配置工具"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$host.UI.RawUI.BackgroundColor = "Black"
$host.UI.RawUI.ForegroundColor = "White"
Clear-Host

# 显示窗口
[void]$form.ShowDialog() 