# Office-OneClick 自定义配置工具
# 允许用户选择Office版本、语言和组件，然后生成自定义config.xml
# 适用于 OfficeOneClick 工具

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
$form.Text = "Office 自定义配置工具"
$form.Size = New-Object System.Drawing.Size(650, 650)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.BackColor = [System.Drawing.Color]::White

# 添加标题
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Office 自定义配置工具"
$titleLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.Location = New-Object System.Drawing.Point(20, 20)
$titleLabel.Size = New-Object System.Drawing.Size(600, 30)
$form.Controls.Add($titleLabel)

# 添加描述
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "请选择您需要的 Office 版本、语言及组件，生成自定义配置文件"
$descriptionLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$descriptionLabel.Location = New-Object System.Drawing.Point(20, 60)
$descriptionLabel.Size = New-Object System.Drawing.Size(600, 20)
$form.Controls.Add($descriptionLabel)

# 版本选择组
$versionGroupBox = New-Object System.Windows.Forms.GroupBox
$versionGroupBox.Text = "Office 版本"
$versionGroupBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$versionGroupBox.Location = New-Object System.Drawing.Point(20, 90)
$versionGroupBox.Size = New-Object System.Drawing.Size(290, 170)
$form.Controls.Add($versionGroupBox)

# 版本选择下拉框标签
$versionLabel = New-Object System.Windows.Forms.Label
$versionLabel.Text = "选择 Office 版本:"
$versionLabel.Location = New-Object System.Drawing.Point(15, 25)
$versionLabel.Size = New-Object System.Drawing.Size(120, 20)
$versionGroupBox.Controls.Add($versionLabel)

# 版本选择下拉框
$versionComboBox = New-Object System.Windows.Forms.ComboBox
$versionComboBox.Location = New-Object System.Drawing.Point(135, 23)
$versionComboBox.Size = New-Object System.Drawing.Size(145, 25)
$versionComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$versionComboBox.DropDownWidth = 300
$versionComboBox.BackColor = [System.Drawing.SystemColors]::Window
$versionComboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
$versionGroupBox.Controls.Add($versionComboBox)

# 添加Office版本
$versions = @{
    "Office LTSC 专业增强版 2024" = @{Channel = "PerpetualVL2024"; ProductID = "ProPlus2024Volume"; GVLK = "XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB"}
    "Office LTSC 标准版 2024" = @{Channel = "PerpetualVL2024"; ProductID = "Standard2024Volume"; GVLK = "V28N4-JG22K-W66P8-VTMGK-H6HGR"}
    "Office LTSC 专业增强版 2021" = @{Channel = "PerpetualVL2021"; ProductID = "ProPlus2021Volume"; GVLK = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"}
    "Office LTSC 标准版 2021" = @{Channel = "PerpetualVL2021"; ProductID = "Standard2021Volume"; GVLK = "KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3"}
    "Office 专业增强版 2019" = @{Channel = "PerpetualVL2019"; ProductID = "ProPlus2019Volume"; GVLK = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"}
    "Office 标准版 2019" = @{Channel = "PerpetualVL2019"; ProductID = "Standard2019Volume"; GVLK = "6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK"}
    "Office 专业增强版 2016" = @{Channel = "PerpetualVL2016"; ProductID = "ProPlusVolume"; GVLK = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"}
    "Office 标准版 2016" = @{Channel = "PerpetualVL2016"; ProductID = "StandardVolume"; GVLK = "JNRGM-WHDWX-FJJG3-K47QV-DRTFM"}
}

# 按年份降序添加版本
$orderedVersions = @(
    "Office LTSC 专业增强版 2024",
    "Office LTSC 标准版 2024",
    "Office LTSC 专业增强版 2021",
    "Office LTSC 标准版 2021",
    "Office 专业增强版 2019",
    "Office 标准版 2019",
    "Office 专业增强版 2016",
    "Office 标准版 2016"
)

foreach ($version in $orderedVersions) {
    $versionComboBox.Items.Add($version)
}
$versionComboBox.SelectedIndex = 0

# Office版本架构
$architectureLabel = New-Object System.Windows.Forms.Label
$architectureLabel.Text = "系统架构:"
$architectureLabel.Location = New-Object System.Drawing.Point(10, 60)
$architectureLabel.Size = New-Object System.Drawing.Size(120, 20)
$versionGroupBox.Controls.Add($architectureLabel)

$arch64RadioButton = New-Object System.Windows.Forms.RadioButton
$arch64RadioButton.Text = "64位 (推荐)"
$arch64RadioButton.Location = New-Object System.Drawing.Point(140, 60)
$arch64RadioButton.Size = New-Object System.Drawing.Size(120, 20)
$arch64RadioButton.Checked = $true
$versionGroupBox.Controls.Add($arch64RadioButton)

$arch32RadioButton = New-Object System.Windows.Forms.RadioButton
$arch32RadioButton.Text = "32位"
$arch32RadioButton.Location = New-Object System.Drawing.Point(140, 85)
$arch32RadioButton.Size = New-Object System.Drawing.Size(120, 20)
$versionGroupBox.Controls.Add($arch32RadioButton)

# 产品密钥
$keyLabel = New-Object System.Windows.Forms.Label
$keyLabel.Text = "产品密钥:"
$keyLabel.Location = New-Object System.Drawing.Point(10, 120)
$keyLabel.Size = New-Object System.Drawing.Size(120, 20)
$versionGroupBox.Controls.Add($keyLabel)

$keyTextBox = New-Object System.Windows.Forms.TextBox
$keyTextBox.Location = New-Object System.Drawing.Point(140, 118)
$keyTextBox.Size = New-Object System.Drawing.Size(140, 25)
$keyTextBox.Text = "XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB"
$versionGroupBox.Controls.Add($keyTextBox)

# 语言选择组
$languageGroupBox = New-Object System.Windows.Forms.GroupBox
$languageGroupBox.Text = "语言选项"
$languageGroupBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$languageGroupBox.Location = New-Object System.Drawing.Point(330, 90)
$languageGroupBox.Size = New-Object System.Drawing.Size(290, 170)
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

# 添加语言选项
$languages = @{
    "简体中文" = "zh-cn"
    "繁体中文(台湾)" = "zh-tw"
    "繁体中文(香港)" = "zh-hk"
    "英语(美国)" = "en-us"
    "英语(英国)" = "en-gb"
    "日语" = "ja-jp"
    "韩语" = "ko-kr"
    "法语" = "fr-fr"
    "德语" = "de-de"
    "西班牙语" = "es-es"
    "俄语" = "ru-ru"
}

foreach ($language in $languages.Keys) {
    $mainLangComboBox.Items.Add($language)
}
# 设置简体中文为默认语言
$mainLangComboBox.SelectedItem = "简体中文"

# 其他语言标签
$otherLangLabel = New-Object System.Windows.Forms.Label
$otherLangLabel.Text = "其他语言(可选):"
$otherLangLabel.Location = New-Object System.Drawing.Point(10, 55)
$otherLangLabel.Size = New-Object System.Drawing.Size(120, 20)
$languageGroupBox.Controls.Add($otherLangLabel)

# 添加自定义语言按钮
$addLangButton = New-Object System.Windows.Forms.Button
$addLangButton.Text = "更多语言..."
$addLangButton.Location = New-Object System.Drawing.Point(200, 53)
$addLangButton.Size = New-Object System.Drawing.Size(80, 23)
$addLangButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8)
$addLangButton.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$languageGroupBox.Controls.Add($addLangButton)

# 添加常用语言复选框
$langCheckBoxes = @{}
$langYStart = 80
$langYIncrement = 20
$langsPerColumn = 4
$langColumnWidth = 140

$commonLanguages = @(
    @{Key = "英语(美国)"; Value = "en-us"},
    @{Key = "英语(英国)"; Value = "en-gb"},
    @{Key = "日语"; Value = "ja-jp"},
    @{Key = "韩语"; Value = "ko-kr"},
    @{Key = "法语"; Value = "fr-fr"},
    @{Key = "德语"; Value = "de-de"},
    @{Key = "西班牙语"; Value = "es-es"},
    @{Key = "俄语"; Value = "ru-ru"}
)

for ($i = 0; $i -lt $commonLanguages.Count; $i++) {
    $langKey = $commonLanguages[$i].Key
    $langValue = $commonLanguages[$i].Value
    
    $column = [Math]::Floor($i / $langsPerColumn)
    $row = $i % $langsPerColumn
    
    $langCheckBox = New-Object System.Windows.Forms.CheckBox
    $langCheckBox.Text = $langKey
    $langCheckBox.Location = New-Object System.Drawing.Point((10 + $column * $langColumnWidth), ($langYStart + $row * $langYIncrement))
    $langCheckBox.Size = New-Object System.Drawing.Size(130, 20)
    $langCheckBox.Checked = $false
    $langCheckBox.Tag = $langValue
    $langCheckBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8)
    $languageGroupBox.Controls.Add($langCheckBox)
    $langCheckBoxes[$langKey] = $langCheckBox
}

# 添加主语言变更事件
$mainLangComboBox.Add_SelectedIndexChanged({
    $selectedLang = $mainLangComboBox.SelectedItem
    
    # 如果有相同的语言在其他语言复选框中，则取消选中
    if ($langCheckBoxes.ContainsKey($selectedLang)) {
        $langCheckBoxes[$selectedLang].Checked = $false
    }
})

# 添加更多语言事件
$addLangButton.Add_Click({
    # 创建语言选择对话框
    $langForm = New-Object System.Windows.Forms.Form
    $langForm.Text = "添加更多语言"
    $langForm.Size = New-Object System.Drawing.Size(350, 400)
    $langForm.StartPosition = "CenterParent"
    $langForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $langForm.MaximizeBox = $false
    $langForm.MinimizeBox = $false
    
    # 添加语言列表
    $langListBox = New-Object System.Windows.Forms.CheckedListBox
    $langListBox.Location = New-Object System.Drawing.Point(15, 15)
    $langListBox.Size = New-Object System.Drawing.Size(310, 300)
    $langListBox.CheckOnClick = $true
    $langForm.Controls.Add($langListBox)
    
    # 创建已选中语言的列表
    $selectedLanguages = @{}
    foreach ($langKey in $langCheckBoxes.Keys) {
        if ($langCheckBoxes[$langKey].Checked) {
            $selectedLanguages[$langKey] = $true
        }
    }
    
    # 添加所有可用语言，并标记已选中的语言
    foreach ($language in $languages.Keys) {
        # 跳过主语言
        if ($language -ne $mainLangComboBox.SelectedItem) {
            # 添加到列表
            $index = $langListBox.Items.Add($language)
            
            # 如果该语言已被选中，则在列表中标记为选中状态
            if ($selectedLanguages.ContainsKey($language)) {
                $langListBox.SetItemChecked($index, $true)
            }
        }
    }
    
    # 添加按钮
    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Text = "添加选中语言"
    $addButton.Location = New-Object System.Drawing.Point(15, 325)
    $addButton.Size = New-Object System.Drawing.Size(150, 30)
    $langForm.Controls.Add($addButton)
    
    # 取消按钮
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "取消"
    $cancelButton.Location = New-Object System.Drawing.Point(175, 325)
    $cancelButton.Size = New-Object System.Drawing.Size(150, 30)
    $langForm.Controls.Add($cancelButton)
    
    # 添加按钮事件
    $addButton.Add_Click({
        # 保存当前已选中的语言
        $currentSelectedLanguages = @{}
        foreach ($langKey in $langCheckBoxes.Keys) {
            if ($langCheckBoxes[$langKey].Checked) {
                $currentSelectedLanguages[$langKey] = $true
            }
        }
        
        # 处理取消选中的语言复选框
        foreach ($langKey in @($langCheckBoxes.Keys)) {
            $index = $langListBox.Items.IndexOf($langKey)
            if ($index -ge 0 -and -not $langListBox.GetItemChecked($index)) {
                # 如果语言在列表中且未选中，则移除对应的复选框
                if ($langCheckBoxes.ContainsKey($langKey)) {
                    $languageGroupBox.Controls.Remove($langCheckBoxes[$langKey])
                    $langCheckBoxes.Remove($langKey)
                }
            }
        }
        
        # 为每个选中的语言添加或更新复选框
        for ($i = 0; $i -lt $langListBox.Items.Count; $i++) {
            if ($langListBox.GetItemChecked($i)) {
                $selectedLang = $langListBox.Items[$i]
                $langCode = $languages[$selectedLang]
                
                # 检查是否已经存在该语言的复选框
                if (-not $langCheckBoxes.ContainsKey($selectedLang)) {
                    # 计算新位置
                    $newIndex = $langCheckBoxes.Count
                    $column = [Math]::Floor($newIndex / $langsPerColumn)
                    $row = $newIndex % $langsPerColumn
                    
                    $customLangCheckBox = New-Object System.Windows.Forms.CheckBox
                    $customLangCheckBox.Text = $selectedLang
                    $customLangCheckBox.Location = New-Object System.Drawing.Point((10 + $column * $langColumnWidth), ($langYStart + $row * $langYIncrement))
                    $customLangCheckBox.Size = New-Object System.Drawing.Size(130, 20)
                    $customLangCheckBox.Checked = $true
                    $customLangCheckBox.Tag = $langCode
                    $customLangCheckBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8)
                    $languageGroupBox.Controls.Add($customLangCheckBox)
                    $langCheckBoxes[$selectedLang] = $customLangCheckBox
                } else {
                    # 如果已存在，确保为选中状态
                    $langCheckBoxes[$selectedLang].Checked = $true
                }
            }
        }
        
        $langForm.Close()
    })
    
    # 取消按钮事件
    $cancelButton.Add_Click({
        $langForm.Close()
    })
    
    # 显示窗体
    $langForm.ShowDialog() | Out-Null
})

# 组件选择组
$componentsGroupBox = New-Object System.Windows.Forms.GroupBox
$componentsGroupBox.Text = "Office 组件"
$componentsGroupBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$componentsGroupBox.Location = New-Object System.Drawing.Point(20, 270)
$componentsGroupBox.Size = New-Object System.Drawing.Size(600, 200)
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
$components = @{
    "Word" = @{ Text = "Word (文字处理)"; Checked = $true }
    "Excel" = @{ Text = "Excel (电子表格)"; Checked = $true }
    "PowerPoint" = @{ Text = "PowerPoint (演示文稿)"; Checked = $true }
    "Outlook" = @{ Text = "Outlook (邮件客户端)"; Checked = $false }
    "Access" = @{ Text = "Access (数据库)"; Checked = $false }
    "Publisher" = @{ Text = "Publisher (桌面排版)"; Checked = $false }
    "OneNote" = @{ Text = "OneNote (笔记本)"; Checked = $false }
    "Lync" = @{ Text = "Skype for Business"; Checked = $false }
    "OneDrive" = @{ Text = "OneDrive (云存储)"; Checked = $false }
    "Groove" = @{ Text = "Groove (OneDrive同步)"; Checked = $false }
}

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
        $checkBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
        $componentsGroupBox.Controls.Add($checkBox)
        $checkBoxes[$component] = $checkBox
    }
}

# 添加全选/取消全选功能
$selectAllCheckBox.Add_CheckedChanged({
    foreach ($component in $components.Keys) {
        if ($checkBoxes.ContainsKey($component)) {
            $checkBoxes[$component].Checked = $selectAllCheckBox.Checked
        }
    }
})

# 按钮区域
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(20, 550)
$buttonPanel.Size = New-Object System.Drawing.Size(600, 70)
$form.Controls.Add($buttonPanel)

# 生成配置按钮
$generateButton = New-Object System.Windows.Forms.Button
$generateButton.Text = "生成配置文件"
$generateButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10, [System.Drawing.FontStyle]::Bold)
$generateButton.Location = New-Object System.Drawing.Point(0, 10)
$generateButton.Size = New-Object System.Drawing.Size(295, 40)
$generateButton.BackColor = [System.Drawing.Color]::DodgerBlue
$generateButton.ForeColor = [System.Drawing.Color]::White
$generateButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonPanel.Controls.Add($generateButton)

# 取消按钮
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "取消"
$cancelButton.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
$cancelButton.Location = New-Object System.Drawing.Point(305, 10)
$cancelButton.Size = New-Object System.Drawing.Size(295, 40)
$cancelButton.BackColor = [System.Drawing.Color]::LightGray
$cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonPanel.Controls.Add($cancelButton)

# 添加状态栏
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "准备就绪"
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# 添加更新和激活选项组
$optionsGroupBox = New-Object System.Windows.Forms.GroupBox
$optionsGroupBox.Text = "安装选项"
$optionsGroupBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$optionsGroupBox.Location = New-Object System.Drawing.Point(20, 480)
$optionsGroupBox.Size = New-Object System.Drawing.Size(600, 60)
$form.Controls.Add($optionsGroupBox)

# 启用更新选项
$updatesCheckbox = New-Object System.Windows.Forms.CheckBox
$updatesCheckbox.Text = "启用Office更新"
$updatesCheckbox.Location = New-Object System.Drawing.Point(20, 25)
$updatesCheckbox.Size = New-Object System.Drawing.Size(150, 20)
$updatesCheckbox.Checked = $true
$optionsGroupBox.Controls.Add($updatesCheckbox)

# 自动激活选项
$autoActivateCheckbox = New-Object System.Windows.Forms.CheckBox
$autoActivateCheckbox.Text = "安装后自动激活"
$autoActivateCheckbox.Location = New-Object System.Drawing.Point(190, 25)
$autoActivateCheckbox.Size = New-Object System.Drawing.Size(150, 20)
$autoActivateCheckbox.Checked = $true
$optionsGroupBox.Controls.Add($autoActivateCheckbox)

# 强制关闭应用程序选项
$forceAppShutdownCheckbox = New-Object System.Windows.Forms.CheckBox
$forceAppShutdownCheckbox.Text = "强制关闭运行的Office应用"
$forceAppShutdownCheckbox.Location = New-Object System.Drawing.Point(360, 25)
$forceAppShutdownCheckbox.Size = New-Object System.Drawing.Size(220, 20)
$forceAppShutdownCheckbox.Checked = $true
$optionsGroupBox.Controls.Add($forceAppShutdownCheckbox)

# 生成配置文件函数
function GenerateConfigFile {
    # 获取选择的版本信息
    $selectedVersion = $versionComboBox.SelectedItem
    $channelValue = $versions[$selectedVersion].Channel
    $productID = $versions[$selectedVersion].ProductID
    
    $architecture = "64"
    if ($arch32RadioButton.Checked) {
        $architecture = "32"
    }
    
    # 获取选择的主语言
    $selectedLanguage = $mainLangComboBox.SelectedItem
    $mainLanguageValue = $languages[$selectedLanguage]
    
    # 获取选中的其他语言
    $additionalLanguages = @()
    foreach ($langKey in $langCheckBoxes.Keys) {
        if ($langCheckBoxes[$langKey].Checked -and $languages.ContainsKey($langKey)) {
            $additionalLanguages += $languages[$langKey]
        }
    }
    
    # 获取产品密钥
    $pidKey = $keyTextBox.Text.Trim()
    
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
    $addNode.SetAttribute("OfficeClientEdition", $architecture)
    $addNode.SetAttribute("Channel", $channelValue)
    $configNode.AppendChild($addNode) | Out-Null
    
    # 添加Product节点
    $productNode = $xmlDoc.CreateElement("Product")
    $productNode.SetAttribute("ID", $productID)
    
    if (-not [string]::IsNullOrWhiteSpace($pidKey)) {
        $productNode.SetAttribute("PIDKEY", $pidKey)
    }
    
    $addNode.AppendChild($productNode) | Out-Null
    
    # 添加主语言节点
    $languageNode = $xmlDoc.CreateElement("Language")
    $languageNode.SetAttribute("ID", $mainLanguageValue)
    $productNode.AppendChild($languageNode) | Out-Null
    
    # 添加其他语言节点
    foreach ($langCode in $additionalLanguages) {
        # 避免添加重复的语言
        if ($langCode -ne $mainLanguageValue) {
            $addLangNode = $xmlDoc.CreateElement("Language")
            $addLangNode.SetAttribute("ID", $langCode)
            $productNode.AppendChild($addLangNode) | Out-Null
        }
    }
    
    # 添加排除的应用程序
    foreach ($component in $components.Keys) {
        if ($checkBoxes.ContainsKey($component) -and -not $checkBoxes[$component].Checked) {
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
    }
    
    # 根据复选框设置属性值
    if ($forceAppShutdownCheckbox.Checked) {
        $properties["FORCEAPPSHUTDOWN"] = "TRUE"
    } else {
        $properties["FORCEAPPSHUTDOWN"] = "FALSE"
    }
    
    if ($autoActivateCheckbox.Checked) {
        $properties["AUTOACTIVATE"] = "1"
    } else {
        $properties["AUTOACTIVATE"] = "0"
    }
    
    foreach ($prop in $properties.Keys) {
        $propNode = $xmlDoc.CreateElement("Property")
        $propNode.SetAttribute("Name", $prop)
        $propNode.SetAttribute("Value", $properties[$prop])
        $configNode.AppendChild($propNode) | Out-Null
    }
    
    # 添加更新节点
    $updatesNode = $xmlDoc.CreateElement("Updates")
    if ($updatesCheckbox.Checked) {
        $updatesNode.SetAttribute("Enabled", "TRUE")
    } else {
        $updatesNode.SetAttribute("Enabled", "FALSE")
    }
    $configNode.AppendChild($updatesNode) | Out-Null
    
    # 添加RemoveMSI节点
    $removeMsiNode = $xmlDoc.CreateElement("RemoveMSI")
    $configNode.AppendChild($removeMsiNode) | Out-Null
    
    # 保存XML文件
    try {
        # 使用临时文件夹路径，与主程序保持一致
        $tempFolder = "$env:TEMP\OfficeInstall"
        if (!(Test-Path $tempFolder)) {
            New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null
        }
        
        $configPath = "$tempFolder\config.xml"
        $xmlDoc.Save($configPath)
        $statusLabel.Text = "配置文件已生成: $configPath"
        
        # 提示用户
        [System.Windows.Forms.MessageBox]::Show(
            "配置文件已成功生成！`n`n路径: $configPath`n`n现在您可以执行Office下载和安装操作。",
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

# 版本选择事件处理
function UpdateProductKey {
    $selectedVersion = $versionComboBox.SelectedItem
    if (-not [string]::IsNullOrEmpty($selectedVersion)) {
        # 获取选择的版本对应的产品密钥
        $gvlk = $versions[$selectedVersion].GVLK
        if (-not [string]::IsNullOrEmpty($gvlk)) {
            $keyTextBox.Text = $gvlk
        }
        
        # 判断版本是否包含Publisher (2024不包含)
        $versionYear = ""
        if ($selectedVersion -match "2024") {
            $versionYear = "2024"
        } elseif ($selectedVersion -match "2021") {
            $versionYear = "2021"
        } elseif ($selectedVersion -match "2019") {
            $versionYear = "2019"
        } elseif ($selectedVersion -match "2016") {
            $versionYear = "2016"
        }
        
        # 根据版本设置组件可用性
        if ($versionYear -ne "") {
            # 设置Publisher的可用性（2024不可用）
            if ($versionYear -eq "2024" -and $checkBoxes.ContainsKey("Publisher")) {
                $checkBoxes["Publisher"].Enabled = $false
                $checkBoxes["Publisher"].Checked = $false
                $checkBoxes["Publisher"].Text = "Publisher (2024版本不可用)"
            } elseif ($checkBoxes.ContainsKey("Publisher")) {
                $checkBoxes["Publisher"].Enabled = $true
                $checkBoxes["Publisher"].Text = "Publisher (桌面排版)"
            }
            
            # 设置OneNote的可用性提示
            if (($versionYear -eq "2024" -or $versionYear -eq "2021" -or $versionYear -eq "2019") -and $checkBoxes.ContainsKey("OneNote")) {
                $checkBoxes["OneNote"].Text = "OneNote (使用Windows自带版本)"
            } elseif ($checkBoxes.ContainsKey("OneNote")) {
                $checkBoxes["OneNote"].Text = "OneNote (笔记本)"
            }
        }
    }
}

# 添加版本选择事件
$versionComboBox.Add_SelectedIndexChanged({
    UpdateProductKey
})

# 添加按钮事件
$generateButton.Add_Click({
    if (GenerateConfigFile) {
        $form.Close()
    }
})

$cancelButton.Add_Click({
    $form.Close()
})

# 初始化产品密钥和组件启用状态
UpdateProductKey

# 隐藏调试信息输出
$Host.UI.RawUI.WindowTitle = "Office 自定义配置工具"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$host.UI.RawUI.BackgroundColor = "Black"
$host.UI.RawUI.ForegroundColor = "White"
Clear-Host

# 显示窗口
[void]$form.ShowDialog() 