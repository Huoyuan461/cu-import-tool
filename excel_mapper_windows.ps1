Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ScriptDir "rules_config_windows.json"

function Normalize-Key($Value) {
    if ($null -eq $Value) { return "" }
    $text = [string]$Value
    $text = $text.Trim()
    $text = [regex]::Replace($text, "\s+", " ")
    return $text.ToUpperInvariant()
}

function Split-Keys($Value) {
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    if ($null -eq $Value) { return $set }
    $parts = ([string]$Value) -split "[,，;；`r`n]+"
    foreach ($part in $parts) {
        $key = Normalize-Key $part
        if ($key.Length -gt 0) { [void]$set.Add($key) }
    }
    return $set
}

function Is-Blank($Value) {
    return ($null -eq $Value) -or (([string]$Value).Trim().Length -eq 0)
}

function Get-ColumnNumber($Worksheet, $ColumnRef, [int]$HeaderRow) {
    $ref = ([string]$ColumnRef).Trim()
    if ($ref.Length -eq 0) { throw "列配置为空" }
    if ($ref -match "^\d+$") {
        $n = [int]$ref
        if ($n -le 0) { throw "列号必须大于0: $ref" }
        return $n
    }
    if ($ref -match "^[A-Za-z]{1,3}$") {
        $sum = 0
        foreach ($ch in $ref.ToUpperInvariant().ToCharArray()) {
            $sum = $sum * 26 + ([int][char]$ch - [int][char]'A' + 1)
        }
        return $sum
    }
    $wanted = Normalize-Key $ref
    $lastCol = $Worksheet.UsedRange.Columns.Count
    for ($col = 1; $col -le $lastCol; $col++) {
        if ((Normalize-Key $Worksheet.Cells.Item($HeaderRow, $col).Value2) -eq $wanted) {
            return $col
        }
    }
    throw "找不到表头列: $ref"
}

function Get-WorksheetByName($Workbook, [string]$Name) {
    foreach ($ws in @($Workbook.Worksheets)) {
        if ($ws.Name -eq $Name) { return $ws }
    }
    return $null
}

function Get-GridRows($Grid) {
    $rows = @()
    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) { continue }
        $obj = [ordered]@{}
        $hasData = $false
        foreach ($col in $Grid.Columns) {
            $value = $row.Cells[$col.Name].Value
            if ($null -ne $value -and ([string]$value).Trim().Length -gt 0) { $hasData = $true }
            $obj[$col.Name] = if ($null -eq $value) { "" } else { [string]$value }
        }
        if ($hasData) { $rows += [pscustomobject]$obj }
    }
    return $rows
}

function Add-GridColumns($Grid, $Columns) {
    $Grid.Columns.Clear()
    foreach ($name in $Columns) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $name
        $col.HeaderText = $name
        $col.Width = 130
        [void]$Grid.Columns.Add($col)
    }
}

function Set-GridRows($Grid, $Rows) {
    $Grid.Rows.Clear()
    foreach ($row in $Rows) {
        $index = $Grid.Rows.Add()
        foreach ($col in $Grid.Columns) {
            $Grid.Rows[$index].Cells[$col.Name].Value = $row.$($col.Name)
        }
    }
}

function New-DefaultConfig {
    return [pscustomobject]@{
        products = @(
            [pscustomobject]@{
                product_name = "产品A"
                customer_value = "客户名称"
                part_no_value = "零件号"
                project_no_value = "项目号"
            }
        )
        mappings = @(
            [pscustomobject]@{
                rule_name = "示例_取数"
                product_name = "产品A"
                source_sheet = "Sheet1"
                customer_column = "客户"
                part_no_column = "零件号"
                project_no_column = "项目号"
                operation = "copy_value"
                source_value_column = "金额"
                target_sheet = "Sheet1"
                target_cell = "B2"
                empty_policy = "skip"
                header_row = "1"
            }
        )
    }
}

function Load-Config {
    if (Test-Path $ConfigPath) {
        return Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    return New-DefaultConfig
}

function Save-Config($Products, $Mappings) {
    $config = [pscustomobject]@{ products = $Products; mappings = $Mappings }
    $config | ConvertTo-Json -Depth 5 | Set-Content $ConfigPath -Encoding UTF8
}

function Add-Log($Logs, [string]$Level, [string]$RuleName, [string]$ProductName, [string]$Message, [string]$SourceFile="", [string]$Target="", $Value=$null) {
    $Logs.Add([pscustomobject]@{
        level = $Level
        rule_name = $RuleName
        product_name = $ProductName
        source_file = $SourceFile
        target = $Target
        value = $Value
        message = $Message
    }) | Out-Null
}

function Product-Matches($Worksheet, [int]$RowIndex, $Product, $Mapping, [int]$HeaderRow, $Logs, [string]$SourceFile) {
    $checks = @(
        @{ column = $Mapping.customer_column; keys = (Split-Keys $Product.customer_value); label = "客户列" },
        @{ column = $Mapping.part_no_column; keys = (Split-Keys $Product.part_no_value); label = "零件号列" },
        @{ column = $Mapping.project_no_column; keys = (Split-Keys $Product.project_no_value); label = "项目号列" }
    )
    foreach ($check in $checks) {
        if (([string]$check.column).Trim().Length -eq 0 -or $check.keys.Count -eq 0) { continue }
        try {
            $col = Get-ColumnNumber $Worksheet $check.column $HeaderRow
            $cellKey = Normalize-Key $Worksheet.Cells.Item($RowIndex, $col).Value2
            if ($check.keys.Contains($cellKey)) { return $true }
        } catch {
            Add-Log $Logs "WARN" $Mapping.rule_name $Mapping.product_name "$($check.label)不可用：$($_.Exception.Message)" $SourceFile
        }
    }
    return $false
}

function Evaluate-Rules($SourceFiles, $Products, $Mappings) {
    $logs = New-Object System.Collections.ArrayList
    $results = New-Object System.Collections.ArrayList
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        foreach ($mapping in $Mappings) {
            if (([string]$mapping.rule_name).Trim().Length -eq 0) { continue }
            $product = $Products | Where-Object { $_.product_name -eq $mapping.product_name } | Select-Object -First 1
            if ($null -eq $product) {
                Add-Log $logs "ERROR" $mapping.rule_name $mapping.product_name "找不到对应的产品规则"
                continue
            }
            $headerRow = 1
            [void][int]::TryParse([string]$mapping.header_row, [ref]$headerRow)
            if ($headerRow -le 0) { $headerRow = 1 }
            $matched = 0
            $copyValue = $null
            $sum = 0.0
            $sumSeen = $false
            foreach ($sourceFile in $SourceFiles) {
                $wb = $null
                try {
                    $wb = $excel.Workbooks.Open($sourceFile, $false, $true)
                    $ws = Get-WorksheetByName $wb $mapping.source_sheet
                    if ($null -eq $ws) {
                        Add-Log $logs "ERROR" $mapping.rule_name $mapping.product_name "源sheet不存在：$($mapping.source_sheet)" $sourceFile
                        continue
                    }
                    try {
                        $valueCol = Get-ColumnNumber $ws $mapping.source_value_column $headerRow
                    } catch {
                        Add-Log $logs "ERROR" $mapping.rule_name $mapping.product_name $_.Exception.Message $sourceFile
                        continue
                    }
                    $lastRow = $ws.UsedRange.Rows.Count
                    for ($r = $headerRow + 1; $r -le $lastRow; $r++) {
                        if (-not (Product-Matches $ws $r $product $mapping $headerRow $logs $sourceFile)) { continue }
                        $matched += 1
                        $raw = $ws.Cells.Item($r, $valueCol).Value2
                        if ($mapping.operation -eq "copy_value") {
                            if ($null -eq $copyValue -and -not (Is-Blank $raw)) {
                                $copyValue = $raw
                                Add-Log $logs "INFO" $mapping.rule_name $mapping.product_name "命中并取第一个非空值，源行 $r" $sourceFile "" $raw
                            }
                        } elseif ($mapping.operation -eq "sum_column") {
                            $num = 0.0
                            if ([double]::TryParse(([string]$raw).Replace(",", ""), [ref]$num)) {
                                $sum += $num
                                $sumSeen = $true
                            } elseif (-not (Is-Blank $raw)) {
                                Add-Log $logs "WARN" $mapping.rule_name $mapping.product_name "求和时忽略非数字值，源行 $r" $sourceFile "" $raw
                            }
                        } else {
                            Add-Log $logs "ERROR" $mapping.rule_name $mapping.product_name "不支持的操作：$($mapping.operation)" $sourceFile
                        }
                    }
                } catch {
                    Add-Log $logs "ERROR" $mapping.rule_name $mapping.product_name "源文件处理失败：$($_.Exception.Message)" $sourceFile
                } finally {
                    if ($null -ne $wb) { $wb.Close($false) | Out-Null }
                }
            }
            $shouldWrite = $false
            $value = $null
            if ($matched -eq 0) {
                Add-Log $logs "WARN" $mapping.rule_name $mapping.product_name "未匹配到产品数据"
            } elseif ($mapping.operation -eq "copy_value" -and $null -ne $copyValue) {
                $value = $copyValue
                $shouldWrite = $true
            } elseif ($mapping.operation -eq "sum_column" -and $sumSeen) {
                $value = $sum
                $shouldWrite = $true
            } else {
                Add-Log $logs "WARN" $mapping.rule_name $mapping.product_name "匹配到产品，但没有可写入的值"
            }
            $results.Add([pscustomobject]@{
                mapping = $mapping
                value = $value
                matched_count = $matched
                should_write = $shouldWrite
            }) | Out-Null
        }
    } finally {
        $excel.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    return [pscustomobject]@{ results = $results; logs = $logs }
}

function Write-Target($TargetPath, $Evaluation) {
    $logs = $Evaluation.logs
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $dir = Split-Path -Parent $TargetPath
    $name = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
    $ext = [System.IO.Path]::GetExtension($TargetPath)
    $backup = Join-Path $dir "$name.backup_$timestamp$ext"
    Copy-Item $TargetPath $backup -Force

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $null
    try {
        $wb = $excel.Workbooks.Open($TargetPath)
        foreach ($result in $Evaluation.results) {
            if (-not $result.should_write) { continue }
            $mapping = $result.mapping
            $ws = Get-WorksheetByName $wb $mapping.target_sheet
            if ($null -eq $ws) {
                Add-Log $logs "ERROR" $mapping.rule_name $mapping.product_name "目标sheet不存在：$($mapping.target_sheet)" "" "$($mapping.target_sheet)!$($mapping.target_cell)" $result.value
                continue
            }
            try {
                $ws.Range($mapping.target_cell).Value2 = $result.value
                Add-Log $logs "INFO" $mapping.rule_name $mapping.product_name "已写入目标单元格" "" "$($mapping.target_sheet)!$($mapping.target_cell)" $result.value
            } catch {
                Add-Log $logs "ERROR" $mapping.rule_name $mapping.product_name "目标单元格写入失败：$($_.Exception.Message)" "" "$($mapping.target_sheet)!$($mapping.target_cell)" $result.value
            }
        }
        $wb.Save()
    } finally {
        if ($null -ne $wb) { $wb.Close($true) | Out-Null }
        $excel.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    return $backup
}

function Show-Logs($Grid, $Logs) {
    $Grid.Rows.Clear()
    foreach ($log in $Logs) {
        $idx = $Grid.Rows.Add()
        $Grid.Rows[$idx].Cells["level"].Value = $log.level
        $Grid.Rows[$idx].Cells["rule_name"].Value = $log.rule_name
        $Grid.Rows[$idx].Cells["product_name"].Value = $log.product_name
        $Grid.Rows[$idx].Cells["source_file"].Value = $log.source_file
        $Grid.Rows[$idx].Cells["target"].Value = $log.target
        $Grid.Rows[$idx].Cells["value"].Value = $log.value
        $Grid.Rows[$idx].Cells["message"].Value = $log.message
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Excel产品数据自动提取与写入工具 - Windows免安装版"
$form.Size = New-Object System.Drawing.Size(1180, 760)
$form.StartPosition = "CenterScreen"

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = "Fill"
$form.Controls.Add($tabs)

$tabSetup = New-Object System.Windows.Forms.TabPage
$tabSetup.Text = "文件与规则"
$tabs.TabPages.Add($tabSetup)

$tabLogs = New-Object System.Windows.Forms.TabPage
$tabLogs.Text = "预览与日志"
$tabs.TabPages.Add($tabLogs)

$lblSources = New-Object System.Windows.Forms.Label
$lblSources.Text = "源Excel文件（可多选）"
$lblSources.Location = New-Object System.Drawing.Point(15, 15)
$lblSources.Size = New-Object System.Drawing.Size(180, 22)
$tabSetup.Controls.Add($lblSources)

$txtSources = New-Object System.Windows.Forms.TextBox
$txtSources.Location = New-Object System.Drawing.Point(15, 40)
$txtSources.Size = New-Object System.Drawing.Size(920, 55)
$txtSources.Multiline = $true
$txtSources.ScrollBars = "Vertical"
$tabSetup.Controls.Add($txtSources)

$btnSources = New-Object System.Windows.Forms.Button
$btnSources.Text = "选择源文件"
$btnSources.Location = New-Object System.Drawing.Point(950, 40)
$btnSources.Size = New-Object System.Drawing.Size(150, 30)
$tabSetup.Controls.Add($btnSources)

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = "目标Excel文件（会覆盖，执行前自动备份；请先关闭Excel里的目标文件）"
$lblTarget.Location = New-Object System.Drawing.Point(15, 105)
$lblTarget.Size = New-Object System.Drawing.Size(600, 22)
$tabSetup.Controls.Add($lblTarget)

$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = New-Object System.Drawing.Point(15, 130)
$txtTarget.Size = New-Object System.Drawing.Size(920, 24)
$tabSetup.Controls.Add($txtTarget)

$btnTarget = New-Object System.Windows.Forms.Button
$btnTarget.Text = "选择目标文件"
$btnTarget.Location = New-Object System.Drawing.Point(950, 128)
$btnTarget.Size = New-Object System.Drawing.Size(150, 30)
$tabSetup.Controls.Add($btnTarget)

$lblProducts = New-Object System.Windows.Forms.Label
$lblProducts.Text = "产品规则：多个关键值可用逗号、分号或换行分隔。客户/零件号/项目号任一命中即可识别。"
$lblProducts.Location = New-Object System.Drawing.Point(15, 170)
$lblProducts.Size = New-Object System.Drawing.Size(900, 22)
$tabSetup.Controls.Add($lblProducts)

$gridProducts = New-Object System.Windows.Forms.DataGridView
$gridProducts.Location = New-Object System.Drawing.Point(15, 195)
$gridProducts.Size = New-Object System.Drawing.Size(1090, 150)
$gridProducts.AllowUserToAddRows = $true
$gridProducts.AllowUserToDeleteRows = $true
$gridProducts.AutoSizeColumnsMode = "Fill"
Add-GridColumns $gridProducts @("product_name", "customer_value", "part_no_value", "project_no_value")
$tabSetup.Controls.Add($gridProducts)

$lblMappings = New-Object System.Windows.Forms.Label
$lblMappings.Text = "映射规则：operation填写 copy_value 或 sum_column；列可填表头名称、列字母或列号。"
$lblMappings.Location = New-Object System.Drawing.Point(15, 360)
$lblMappings.Size = New-Object System.Drawing.Size(900, 22)
$tabSetup.Controls.Add($lblMappings)

$gridMappings = New-Object System.Windows.Forms.DataGridView
$gridMappings.Location = New-Object System.Drawing.Point(15, 385)
$gridMappings.Size = New-Object System.Drawing.Size(1090, 230)
$gridMappings.AllowUserToAddRows = $true
$gridMappings.AllowUserToDeleteRows = $true
$gridMappings.AutoSizeColumnsMode = "DisplayedCells"
Add-GridColumns $gridMappings @("rule_name", "product_name", "source_sheet", "customer_column", "part_no_column", "project_no_column", "operation", "source_value_column", "target_sheet", "target_cell", "empty_policy", "header_row")
$tabSetup.Controls.Add($gridMappings)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "保存规则"
$btnSave.Location = New-Object System.Drawing.Point(15, 635)
$btnSave.Size = New-Object System.Drawing.Size(130, 34)
$tabSetup.Controls.Add($btnSave)

$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = "重新加载规则"
$btnLoad.Location = New-Object System.Drawing.Point(160, 635)
$btnLoad.Size = New-Object System.Drawing.Size(130, 34)
$tabSetup.Controls.Add($btnLoad)

$btnPreview = New-Object System.Windows.Forms.Button
$btnPreview.Text = "预览匹配"
$btnPreview.Location = New-Object System.Drawing.Point(815, 635)
$btnPreview.Size = New-Object System.Drawing.Size(130, 34)
$tabSetup.Controls.Add($btnPreview)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "执行写入"
$btnRun.Location = New-Object System.Drawing.Point(970, 635)
$btnRun.Size = New-Object System.Drawing.Size(130, 34)
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(0, 91, 187)
$btnRun.ForeColor = [System.Drawing.Color]::White
$tabSetup.Controls.Add($btnRun)

$gridLogs = New-Object System.Windows.Forms.DataGridView
$gridLogs.Location = New-Object System.Drawing.Point(15, 15)
$gridLogs.Size = New-Object System.Drawing.Size(1090, 610)
$gridLogs.AllowUserToAddRows = $false
$gridLogs.AllowUserToDeleteRows = $false
$gridLogs.ReadOnly = $true
$gridLogs.AutoSizeColumnsMode = "DisplayedCells"
Add-GridColumns $gridLogs @("level", "rule_name", "product_name", "source_file", "target", "value", "message")
$tabLogs.Controls.Add($gridLogs)

$btnSaveLog = New-Object System.Windows.Forms.Button
$btnSaveLog.Text = "导出日志CSV"
$btnSaveLog.Location = New-Object System.Drawing.Point(15, 640)
$btnSaveLog.Size = New-Object System.Drawing.Size(130, 34)
$tabLogs.Controls.Add($btnSaveLog)

function Load-Config-To-Grids {
    $config = Load-Config
    Set-GridRows $gridProducts @($config.products)
    Set-GridRows $gridMappings @($config.mappings)
}

function Get-SelectedSources {
    return @($txtSources.Lines | Where-Object { ([string]$_).Trim().Length -gt 0 })
}

function Validate-Inputs {
    if ((Get-SelectedSources).Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("请选择至少一个源Excel文件。"); return $false }
    if (-not (Test-Path $txtTarget.Text)) { [System.Windows.Forms.MessageBox]::Show("请选择有效的目标Excel文件。"); return $false }
    if ((Get-GridRows $gridProducts).Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("请至少配置一个产品规则。"); return $false }
    if ((Get-GridRows $gridMappings).Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("请至少配置一条映射规则。"); return $false }
    return $true
}

$btnSources.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Excel文件 (*.xlsx)|*.xlsx"
    $dialog.Multiselect = $true
    if ($dialog.ShowDialog() -eq "OK") {
        $txtSources.Lines = $dialog.FileNames
    }
})

$btnTarget.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Excel文件 (*.xlsx)|*.xlsx"
    $dialog.Multiselect = $false
    if ($dialog.ShowDialog() -eq "OK") {
        $txtTarget.Text = $dialog.FileName
    }
})

$btnSave.Add_Click({
    Save-Config (Get-GridRows $gridProducts) (Get-GridRows $gridMappings)
    [System.Windows.Forms.MessageBox]::Show("规则已保存到：$ConfigPath")
})

$btnLoad.Add_Click({
    Load-Config-To-Grids
})

$btnPreview.Add_Click({
    if (-not (Validate-Inputs)) { return }
    try {
        $evaluation = Evaluate-Rules (Get-SelectedSources) (Get-GridRows $gridProducts) (Get-GridRows $gridMappings)
        Show-Logs $gridLogs $evaluation.logs
        $tabs.SelectedTab = $tabLogs
    } catch {
        [System.Windows.Forms.MessageBox]::Show("预览失败：$($_.Exception.Message)")
    }
})

$btnRun.Add_Click({
    if (-not (Validate-Inputs)) { return }
    try {
        $evaluation = Evaluate-Rules (Get-SelectedSources) (Get-GridRows $gridProducts) (Get-GridRows $gridMappings)
        $backup = Write-Target $txtTarget.Text $evaluation
        Show-Logs $gridLogs $evaluation.logs
        $tabs.SelectedTab = $tabLogs
        [System.Windows.Forms.MessageBox]::Show("写入完成。备份文件：$backup")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("写入失败：$($_.Exception.Message)`n请确认目标文件没有在Excel中打开。")
    }
})

$btnSaveLog.Add_Click({
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "CSV文件 (*.csv)|*.csv"
    $dialog.FileName = "processing_log.csv"
    if ($dialog.ShowDialog() -eq "OK") {
        $lines = New-Object System.Collections.ArrayList
        $lines.Add("level,rule_name,product_name,source_file,target,value,message") | Out-Null
        foreach ($row in $gridLogs.Rows) {
            if ($row.IsNewRow) { continue }
            $values = @()
            foreach ($name in @("level", "rule_name", "product_name", "source_file", "target", "value", "message")) {
                $text = [string]$row.Cells[$name].Value
                $text = '"' + $text.Replace('"', '""') + '"'
                $values += $text
            }
            $lines.Add(($values -join ",")) | Out-Null
        }
        [System.IO.File]::WriteAllLines($dialog.FileName, [string[]]$lines, [System.Text.Encoding]::UTF8)
    }
})

Load-Config-To-Grids
[void]$form.ShowDialog()
