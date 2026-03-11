Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:RootDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:DataDir = Join-Path $script:RootDir 'data'
$script:ComputersFile = Join-Path $script:DataDir 'computers.json'
$script:ColleaguesFile = Join-Path $script:DataDir 'colleagues.json'
$script:Computers = @()
$script:Colleagues = @()
$script:CurrentComputerId = $null
$script:CurrentColleagueEditorId = $null
$script:CurrentInventoryComputerId = $null
$script:SelectedOwnerId = $null
$script:SuppressOwnerTextChange = $false
$script:SuppressColleagueAutoPinyin = $false
$script:ColleaguePinyinManuallyEdited = $false

function Ensure-DataFiles {
    if (-not (Test-Path $script:DataDir)) {
        New-Item -ItemType Directory -Path $script:DataDir | Out-Null
    }

    foreach ($file in @($script:ComputersFile, $script:ColleaguesFile)) {
        if (-not (Test-Path $file)) {
            '[]' | Set-Content -Path $file -Encoding UTF8
        }
    }
}

function Show-WarningMessage {
    param([string]$Message, [string]$Title = '提示')

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
}

function Show-InfoMessage {
    param([string]$Message, [string]$Title = '完成')

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Load-JsonArray {
    param([string]$Path)

    try {
        $raw = Get-Content -Path $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return @()
        }

        $data = $raw | ConvertFrom-Json
        if ($null -eq $data) {
            return @()
        }

        if ($data -is [System.Array]) {
            return @($data)
        }

        return @($data)
    } catch {
        Show-WarningMessage -Title '读取失败' -Message "数据文件读取失败，已使用空数据继续。`n`n$($_.Exception.Message)"
        return @()
    }
}

function Save-JsonArray {
    param(
        [string]$Path,
        [array]$Data
    )

    $json = ConvertTo-Json -InputObject @($Data) -Depth 6
    Set-Content -Path $Path -Value $json -Encoding UTF8
}

function Normalize-ColleagueRecord {
    param($Record)

    if (-not ($Record.PSObject.Properties.Name -contains 'id') -or [string]::IsNullOrWhiteSpace([string]$Record.id)) {
        Add-Member -InputObject $Record -MemberType NoteProperty -Name id -Value ([guid]::NewGuid().ToString()) -Force
    }
    if (-not ($Record.PSObject.Properties.Name -contains 'display_name')) {
        Add-Member -InputObject $Record -MemberType NoteProperty -Name display_name -Value '' -Force
    }
    if (-not ($Record.PSObject.Properties.Name -contains 'pinyin')) {
        Add-Member -InputObject $Record -MemberType NoteProperty -Name pinyin -Value '' -Force
    }
    if (-not ($Record.PSObject.Properties.Name -contains 'email') -or [string]::IsNullOrWhiteSpace([string]$Record.email)) {
        Add-Member -InputObject $Record -MemberType NoteProperty -Name email -Value '@itk-engineering.com' -Force
    }
    if (-not ($Record.PSObject.Properties.Name -contains 'department')) {
        Add-Member -InputObject $Record -MemberType NoteProperty -Name department -Value '' -Force
    }
    if (-not ($Record.PSObject.Properties.Name -contains 'employee_type') -or [string]::IsNullOrWhiteSpace([string]$Record.employee_type)) {
        Add-Member -InputObject $Record -MemberType NoteProperty -Name employee_type -Value '正式员工' -Force
    }
    if (-not ($Record.PSObject.Properties.Name -contains 'mentor_id')) {
        Add-Member -InputObject $Record -MemberType NoteProperty -Name mentor_id -Value '' -Force
    }

    if ([string]::IsNullOrWhiteSpace([string]$Record.email)) { $Record.email = '@itk-engineering.com' }
    if ([string]::IsNullOrWhiteSpace([string]$Record.employee_type)) { $Record.employee_type = '正式员工' }
    if ($Record.employee_type -ne '实习生') { $Record.mentor_id = '' }

    return $Record
}

function Get-ColleagueById {
    param([string]$Id)

    if ([string]::IsNullOrWhiteSpace($Id)) { return $null }
    return $script:Colleagues | Where-Object { $_.id -eq $Id } | Select-Object -First 1
}

function Get-OwnerDisplayName {
    param([string]$OwnerId)

    $owner = Get-ColleagueById -Id $OwnerId
    if ($null -eq $owner) { return '' }
    return [string]$owner.display_name
}

function Get-OwnerLabel {
    param([string]$OwnerId)

    if ([string]::IsNullOrWhiteSpace($OwnerId)) { return '库存' }
    $displayName = Get-OwnerDisplayName -OwnerId $OwnerId
    if ([string]::IsNullOrWhiteSpace($displayName)) { return '未知人员' }
    return $displayName
}

function Normalize-ComputerRecord {
    param($Record)

    if (-not ($Record.PSObject.Properties.Name -contains 'id') -or [string]::IsNullOrWhiteSpace([string]$Record.id)) {
        Add-Member -InputObject $Record -MemberType NoteProperty -Name id -Value ([guid]::NewGuid().ToString()) -Force
    }
    foreach ($name in 'computer_name','serial_number','asset_number','model','mac_address','owner_id','remark') {
        if (-not ($Record.PSObject.Properties.Name -contains $name)) {
            Add-Member -InputObject $Record -MemberType NoteProperty -Name $name -Value '' -Force
        }
    }
    if (-not ($Record.PSObject.Properties.Name -contains 'updated_at') -or [string]::IsNullOrWhiteSpace([string]$Record.updated_at)) {
        Add-Member -InputObject $Record -MemberType NoteProperty -Name updated_at -Value (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') -Force
    }

    if (-not ($Record.PSObject.Properties.Name -contains 'owner_history') -or $null -eq $Record.owner_history) {
        $history = @()
        if (-not [string]::IsNullOrWhiteSpace([string]$Record.owner_id)) {
            $history = @([PSCustomObject]@{
                changed_at = [string]$Record.updated_at
                old_owner_id = ''
                old_owner_name = '库存'
                new_owner_id = [string]$Record.owner_id
                new_owner_name = Get-OwnerLabel -OwnerId ([string]$Record.owner_id)
            })
        }
        Add-Member -InputObject $Record -MemberType NoteProperty -Name owner_history -Value $history -Force
    } else {
        $Record.owner_history = @($Record.owner_history)
    }

    return $Record
}

function Load-AllData {
    Ensure-DataFiles
    $script:Colleagues = @(Load-JsonArray -Path $script:ColleaguesFile | ForEach-Object { Normalize-ColleagueRecord -Record $_ })
    $script:Computers = @(Load-JsonArray -Path $script:ComputersFile | ForEach-Object { Normalize-ComputerRecord -Record $_ })
}

function Save-Computers { Save-JsonArray -Path $script:ComputersFile -Data $script:Computers }
function Save-Colleagues { Save-JsonArray -Path $script:ColleaguesFile -Data $script:Colleagues }

function Format-ColleagueOption {
    param($Colleague)

    if ($null -eq $Colleague) { return '' }
    return '{0} ({1}, {2})' -f [string]$Colleague.display_name, [string]$Colleague.pinyin, [string]$Colleague.department
}

function Resolve-HistoryOwnerName {
    param([string]$OwnerId, [string]$StoredName)

    if (-not [string]::IsNullOrWhiteSpace($StoredName)) { return $StoredName }
    return Get-OwnerLabel -OwnerId $OwnerId
}

function Add-ComputerOwnerHistoryEntry {
    param(
        [Parameter(Mandatory = $true)]$ComputerRecord,
        [string]$OldOwnerId,
        [string]$NewOwnerId,
        [string]$ChangedAt
    )

    $ComputerRecord.owner_history = @($ComputerRecord.owner_history) + [PSCustomObject]@{
        changed_at = $ChangedAt
        old_owner_id = [string]$OldOwnerId
        old_owner_name = Get-OwnerLabel -OwnerId ([string]$OldOwnerId)
        new_owner_id = [string]$NewOwnerId
        new_owner_name = Get-OwnerLabel -OwnerId ([string]$NewOwnerId)
    }
}

function Set-ComputerOwner {
    param(
        [Parameter(Mandatory = $true)]$ComputerRecord,
        [string]$NewOwnerId,
        [string]$ChangedAt
    )

    $oldOwnerId = [string]$ComputerRecord.owner_id
    $targetOwnerId = [string]$NewOwnerId
    if ($oldOwnerId -eq $targetOwnerId) { return $false }

    $ComputerRecord.owner_id = $targetOwnerId
    Add-ComputerOwnerHistoryEntry -ComputerRecord $ComputerRecord -OldOwnerId $oldOwnerId -NewOwnerId $targetOwnerId -ChangedAt $ChangedAt
    return $true
}

function Get-AutoPinyinText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

    $encoding = [System.Text.Encoding]::GetEncoding('GB2312')
    $builder = New-Object System.Text.StringBuilder

    foreach ($char in $Text.Trim().ToCharArray()) {
        $charText = [string]$char
        if ([string]::IsNullOrWhiteSpace($charText)) { continue }
        if ($charText -match '[A-Za-z0-9]') {
            [void]$builder.Append($charText.ToLowerInvariant())
            continue
        }

        $bytes = $encoding.GetBytes($charText)
        if ($bytes.Length -lt 2) { continue }

        $code = $bytes[0] * 256 + $bytes[1] - 65536
        $letter = switch ($code) {
            { $_ -ge -20319 -and $_ -le -20284 } { 'a'; break }
            { $_ -ge -20283 -and $_ -le -19776 } { 'b'; break }
            { $_ -ge -19775 -and $_ -le -19219 } { 'c'; break }
            { $_ -ge -19218 -and $_ -le -18711 } { 'd'; break }
            { $_ -ge -18710 -and $_ -le -18527 } { 'e'; break }
            { $_ -ge -18526 -and $_ -le -18240 } { 'f'; break }
            { $_ -ge -18239 -and $_ -le -17923 } { 'g'; break }
            { $_ -ge -17922 -and $_ -le -17418 } { 'h'; break }
            { $_ -ge -17417 -and $_ -le -16475 } { 'j'; break }
            { $_ -ge -16474 -and $_ -le -16213 } { 'k'; break }
            { $_ -ge -16212 -and $_ -le -15641 } { 'l'; break }
            { $_ -ge -15640 -and $_ -le -15166 } { 'm'; break }
            { $_ -ge -15165 -and $_ -le -14923 } { 'n'; break }
            { $_ -ge -14922 -and $_ -le -14915 } { 'o'; break }
            { $_ -ge -14914 -and $_ -le -14631 } { 'p'; break }
            { $_ -ge -14630 -and $_ -le -14150 } { 'q'; break }
            { $_ -ge -14149 -and $_ -le -14091 } { 'r'; break }
            { $_ -ge -14090 -and $_ -le -13319 } { 's'; break }
            { $_ -ge -13318 -and $_ -le -12839 } { 't'; break }
            { $_ -ge -12838 -and $_ -le -12557 } { 'w'; break }
            { $_ -ge -12556 -and $_ -le -11848 } { 'x'; break }
            { $_ -ge -11847 -and $_ -le -11056 } { 'y'; break }
            { $_ -ge -11055 -and $_ -le -10247 } { 'z'; break }
            default { '' }
        }

        if (-not [string]::IsNullOrWhiteSpace($letter)) { [void]$builder.Append($letter) }
    }

    return $builder.ToString().ToLowerInvariant()
}

function Request-EditAuthorization {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = '编辑确认'
    $dialog.StartPosition = 'CenterParent'
    $dialog.Size = New-Object System.Drawing.Size(360, 180)
    $dialog.MinimumSize = New-Object System.Drawing.Size(360, 180)
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.FormBorderStyle = 'FixedDialog'

    $label = New-Object System.Windows.Forms.Label
    $label.Text = '编辑已有电脑信息需要输入授权密码：'
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(300, 24)
    $dialog.Controls.Add($label)

    $passwordBox = New-Object System.Windows.Forms.TextBox
    $passwordBox.Location = New-Object System.Drawing.Point(20, 55)
    $passwordBox.Size = New-Object System.Drawing.Size(300, 28)
    $passwordBox.UseSystemPasswordChar = $true
    $dialog.Controls.Add($passwordBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = '确认'
    $okButton.Location = New-Object System.Drawing.Point(60, 95)
    $okButton.Size = New-Object System.Drawing.Size(100, 30)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dialog.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = '取消'
    $cancelButton.Location = New-Object System.Drawing.Point(190, 95)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dialog.Controls.Add($cancelButton)

    $dialog.AcceptButton = $okButton
    $dialog.CancelButton = $cancelButton

    $result = $dialog.ShowDialog($form)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $false }
    if ($passwordBox.Text -ne '123456') {
        Show-WarningMessage -Title '授权失败' -Message '密码不正确，未保存修改。'
        return $false
    }

    return $true
}

function Test-ComputerFieldValues {
    param([string]$Name, [string]$Serial, [string]$Asset, [string]$Mac)

    if ([string]::IsNullOrWhiteSpace($Name)) { Show-WarningMessage '请输入电脑名称。'; return $false }
    if ([string]::IsNullOrWhiteSpace($Serial)) { Show-WarningMessage '请输入序列号。'; return $false }
    if ([string]::IsNullOrWhiteSpace($Asset)) { Show-WarningMessage '请输入固定资产号。'; return $false }
    if ([string]::IsNullOrWhiteSpace($Mac)) { Show-WarningMessage '请输入 MAC 地址。'; return $false }
    return $true
}

function Refresh-ModelOptions {
    if ($null -eq $cmbModel) { return }

    $currentText = $cmbModel.Text
    $models = @($script:Computers | ForEach-Object { [string]$_.model } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

    $cmbModel.BeginUpdate()
    $cmbModel.Items.Clear()
    foreach ($modelName in $models) { [void]$cmbModel.Items.Add($modelName) }
    $cmbModel.Text = $currentText
    $cmbModel.EndUpdate()
}

function Set-SelectedOwner {
    param($Colleague)

    $script:SuppressOwnerTextChange = $true
    if ($null -eq $Colleague) {
        $script:SelectedOwnerId = $null
        $txtOwner.Text = ''
    } else {
        $script:SelectedOwnerId = [string]$Colleague.id
        $txtOwner.Text = Format-ColleagueOption -Colleague $Colleague
    }
    $script:SuppressOwnerTextChange = $false
    $lstOwnerSuggestions.Visible = $false
}

function Resolve-ColleagueFromOwnerInput {
    param([string]$InputText)

    $text = $InputText.Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    foreach ($colleague in $script:Colleagues) {
        if ((Format-ColleagueOption -Colleague $colleague) -eq $text) {
            return $colleague
        }
    }

    return $script:Colleagues | Where-Object {
        [string]$_.display_name -eq $text -or
        [string]$_.pinyin -eq $text -or
        [string]$_.email -eq $text
    } | Select-Object -First 1
}
function Refresh-OwnerSuggestions {
    if ($null -eq $lstOwnerSuggestions) { return }

    $keyword = $txtOwner.Text.Trim().ToLowerInvariant()
    $lstOwnerSuggestions.Items.Clear()
    if ([string]::IsNullOrWhiteSpace($keyword)) {
        $lstOwnerSuggestions.Visible = $false
        return
    }

    $matches = @($script:Colleagues | Where-Object {
        ([string]$_.pinyin).ToLowerInvariant().Contains($keyword) -or
        ([string]$_.display_name).ToLowerInvariant().Contains($keyword) -or
        ([string]$_.department).ToLowerInvariant().Contains($keyword)
    } | Sort-Object display_name, pinyin, department)

    foreach ($item in $matches) {
        [void]$lstOwnerSuggestions.Items.Add([PSCustomObject]@{
            Id = [string]$item.id
            Display = Format-ColleagueOption -Colleague $item
        })
    }

    $lstOwnerSuggestions.DisplayMember = 'Display'
    $lstOwnerSuggestions.ValueMember = 'Id'
    $lstOwnerSuggestions.Visible = $lstOwnerSuggestions.Items.Count -gt 0
    if ($lstOwnerSuggestions.Visible) {
        $lstOwnerSuggestions.Height = [Math]::Min(110, 24 * $lstOwnerSuggestions.Items.Count + 4)
    }
}

function Get-FilteredComputers {
    $keyword = $txtSearch.Text.Trim().ToLowerInvariant()
    $rows = foreach ($item in $script:Computers) {
        [PSCustomObject]@{
            id = [string]$item.id
            computer_name = [string]$item.computer_name
            serial_number = [string]$item.serial_number
            asset_number = [string]$item.asset_number
            model = [string]$item.model
            mac_address = [string]$item.mac_address
            owner_id = [string]$item.owner_id
            owner_name = Get-OwnerLabel -OwnerId ([string]$item.owner_id)
            remark = [string]$item.remark
            updated_at = [string]$item.updated_at
        }
    }

    if ([string]::IsNullOrWhiteSpace($keyword)) { return @($rows | Sort-Object computer_name, serial_number) }

    return @($rows | Where-Object {
        $_.computer_name.ToLowerInvariant().Contains($keyword) -or
        $_.serial_number.ToLowerInvariant().Contains($keyword) -or
        $_.asset_number.ToLowerInvariant().Contains($keyword) -or
        $_.model.ToLowerInvariant().Contains($keyword) -or
        $_.mac_address.ToLowerInvariant().Contains($keyword) -or
        $_.owner_name.ToLowerInvariant().Contains($keyword) -or
        $_.remark.ToLowerInvariant().Contains($keyword)
    } | Sort-Object computer_name, serial_number)
}

function Add-MainComputerGridColumns {
    param($TargetGrid)

    [void]$TargetGrid.Columns.Add('colName', '电脑名称')
    [void]$TargetGrid.Columns.Add('colSerial', '序列号')
    [void]$TargetGrid.Columns.Add('colAsset', '固定资产号')
    [void]$TargetGrid.Columns.Add('colModel', '型号')
    [void]$TargetGrid.Columns.Add('colMac', 'MAC 地址')
    [void]$TargetGrid.Columns.Add('colOwner', '归属人')
    [void]$TargetGrid.Columns.Add('colRemark', '备注')
    [void]$TargetGrid.Columns.Add('colUpdated', '更新时间')
    $TargetGrid.Columns['colUpdated'].FillWeight = 125
}

function Add-MainComputerGridRow {
    param(
        [Parameter(Mandatory = $true)]$TargetGrid,
        [Parameter(Mandatory = $true)]$RowData
    )

    $index = $TargetGrid.Rows.Add()
    $row = $TargetGrid.Rows[$index]
    $row.Tag = $RowData.id
    $row.Cells['colName'].Value = $RowData.computer_name
    $row.Cells['colSerial'].Value = $RowData.serial_number
    $row.Cells['colAsset'].Value = $RowData.asset_number
    $row.Cells['colModel'].Value = $RowData.model
    $row.Cells['colMac'].Value = $RowData.mac_address
    $row.Cells['colOwner'].Value = $RowData.owner_name
    $row.Cells['colRemark'].Value = $RowData.remark
    $row.Cells['colUpdated'].Value = $RowData.updated_at
}

function Get-MainSelectedComputerId {
    if (-not [string]::IsNullOrWhiteSpace([string]$script:CurrentComputerId)) {
        return [string]$script:CurrentComputerId
    }

    foreach ($targetGrid in @($gridInUse, $gridInventoryMain)) {
        if ($null -ne $targetGrid -and $targetGrid.SelectedRows.Count -gt 0) {
            return [string]$targetGrid.SelectedRows[0].Tag
        }
    }

    return ''
}

function Select-MainComputerRow {
    param([string]$ComputerId)

    if ([string]::IsNullOrWhiteSpace($ComputerId)) { return $false }

    foreach ($targetGrid in @($gridInUse, $gridInventoryMain)) {
        if ($null -eq $targetGrid) { continue }

        foreach ($row in $targetGrid.Rows) {
            if ([string]$row.Tag -eq $ComputerId) {
                if ($targetGrid -eq $gridInUse -and $null -ne $gridInventoryMain) { $gridInventoryMain.ClearSelection() }
                if ($targetGrid -eq $gridInventoryMain -and $null -ne $gridInUse) { $gridInUse.ClearSelection() }
                $targetGrid.ClearSelection()
                $row.Selected = $true
                $targetGrid.CurrentCell = $row.Cells['colName']
                Fill-ComputerForm -ComputerId $ComputerId
                return $true
            }
        }
    }

    return $false
}

function Refresh-ComputerGrid {
    $gridInUse.Rows.Clear()
    $gridInventoryMain.Rows.Clear()
    $rows = Get-FilteredComputers
    $inUseRows = @($rows | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.owner_id) })
    $inventoryRows = @($rows | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.owner_id) })

    foreach ($rowData in $inUseRows) {
        Add-MainComputerGridRow -TargetGrid $gridInUse -RowData $rowData
    }

    foreach ($rowData in $inventoryRows) {
        Add-MainComputerGridRow -TargetGrid $gridInventoryMain -RowData $rowData
    }

    $groupInUse.Text = "在用电脑（$($inUseRows.Count) 台）"
    $groupInventoryMain.Text = "库存电脑（$($inventoryRows.Count) 台）"
    $lblCount.Text = "在用 $($inUseRows.Count) 台，库存 $($inventoryRows.Count) 台，共 $($rows.Count) 台"
}

function Clear-ComputerForm {
    $script:CurrentComputerId = $null
    $txtName.Text = ''
    $txtSerial.Text = ''
    $txtAsset.Text = ''
    $cmbModel.Text = ''
    $txtMac.Text = ''
    $txtRemark.Text = ''
    Set-SelectedOwner -Colleague $null
    $lblMode.Text = '当前模式：新增'
    if ($null -ne $gridInUse) { $gridInUse.ClearSelection() }
    if ($null -ne $gridInventoryMain) { $gridInventoryMain.ClearSelection() }
}

function Fill-ComputerForm {
    param([string]$ComputerId)

    $record = $script:Computers | Where-Object { $_.id -eq $ComputerId } | Select-Object -First 1
    if ($null -eq $record) { return }

    $script:CurrentComputerId = [string]$record.id
    $txtName.Text = [string]$record.computer_name
    $txtSerial.Text = [string]$record.serial_number
    $txtAsset.Text = [string]$record.asset_number
    $cmbModel.Text = [string]$record.model
    $txtMac.Text = [string]$record.mac_address
    $txtRemark.Text = [string]$record.remark
    Set-SelectedOwner -Colleague (Get-ColleagueById -Id ([string]$record.owner_id))
    $lblMode.Text = '当前模式：编辑'
}

function Validate-ComputerInput {
    return Test-ComputerFieldValues -Name $txtName.Text -Serial $txtSerial.Text -Asset $txtAsset.Text -Mac $txtMac.Text
}

function Save-ComputerRecord {
    if (-not (Validate-ComputerInput)) { return }

    $name = $txtName.Text.Trim()
    $serial = $txtSerial.Text.Trim()
    $asset = $txtAsset.Text.Trim()
    $model = $cmbModel.Text.Trim()
    $mac = $txtMac.Text.Trim()
    $remark = $txtRemark.Text.Trim()
    $now = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

    $duplicate = $script:Computers | Where-Object {
        $_.id -ne $script:CurrentComputerId -and (
            [string]$_.serial_number -eq $serial -or
            [string]$_.asset_number -eq $asset
        )
    } | Select-Object -First 1

    if ($duplicate) {
        Show-WarningMessage -Title '重复数据' -Message '序列号或固定资产号已存在，请确认后再保存。'
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:CurrentComputerId)) {
        $history = @()
        if (-not [string]::IsNullOrWhiteSpace($script:SelectedOwnerId)) {
            $history = @([PSCustomObject]@{
                changed_at = $now
                old_owner_id = ''
                old_owner_name = '库存'
                new_owner_id = [string]$script:SelectedOwnerId
                new_owner_name = Get-OwnerLabel -OwnerId ([string]$script:SelectedOwnerId)
            })
        }

        $script:Computers = @($script:Computers) + [PSCustomObject]@{
            id = [guid]::NewGuid().ToString()
            computer_name = $name
            serial_number = $serial
            asset_number = $asset
            model = $model
            mac_address = $mac
            owner_id = [string]$script:SelectedOwnerId
            owner_history = $history
            remark = $remark
            updated_at = $now
        }
    } else {
        if (-not (Request-EditAuthorization)) { return }

        foreach ($item in $script:Computers) {
            if ($item.id -eq $script:CurrentComputerId) {
                $item.computer_name = $name
                $item.serial_number = $serial
                $item.asset_number = $asset
                $item.model = $model
                $item.mac_address = $mac
                [void](Set-ComputerOwner -ComputerRecord $item -NewOwnerId ([string]$script:SelectedOwnerId) -ChangedAt $now)
                $item.remark = $remark
                $item.updated_at = $now
                break
            }
        }
    }

    Save-Computers
    Refresh-ModelOptions
    Refresh-ComputerGrid
    Clear-ComputerForm
    Show-InfoMessage '保存成功。'
}

function Remove-SelectedComputer {
    $selectedId = Get-MainSelectedComputerId
    if ([string]::IsNullOrWhiteSpace($selectedId)) {
        Show-WarningMessage '请先选择要转入库存的电脑记录。'
        return
    }

    $record = $script:Computers | Where-Object { $_.id -eq $selectedId } | Select-Object -First 1
    if ($null -eq $record) {
        Show-WarningMessage '未找到选中的电脑记录。'
        return
    }

    if ([string]::IsNullOrWhiteSpace([string]$record.owner_id)) {
        Show-WarningMessage '这台电脑已经在库存中。如需彻底删除，请到库存电脑管理里删除。'
        return
    }

    $ownerLabel = Get-OwnerLabel -OwnerId ([string]$record.owner_id)
    $result = [System.Windows.Forms.MessageBox]::Show(
        "确定将这台电脑从 $ownerLabel 名下转入库存吗？",
        '确认转入库存',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $now = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    [void](Set-ComputerOwner -ComputerRecord $record -NewOwnerId '' -ChangedAt $now)
    $record.updated_at = $now

    Save-Computers
    Refresh-ModelOptions
    Refresh-ComputerGrid
    Clear-ComputerForm
    Show-InfoMessage '电脑已转入库存。若要彻底删除，请到库存电脑管理里删除。'
}
function Show-OwnerHistory {
    $computerId = Get-MainSelectedComputerId

    if ([string]::IsNullOrWhiteSpace($computerId)) {
        Show-WarningMessage '请先选择一台电脑，再查看归属历史。'
        return
    }

    $record = $script:Computers | Where-Object { $_.id -eq $computerId } | Select-Object -First 1
    if ($null -eq $record) {
        Show-WarningMessage '未找到这台电脑的记录。'
        return
    }

    $historyForm = New-Object System.Windows.Forms.Form
    $historyForm.Text = '归属人历史'
    $historyForm.StartPosition = 'CenterParent'
    $historyForm.Size = New-Object System.Drawing.Size(760, 420)
    $historyForm.MinimumSize = New-Object System.Drawing.Size(700, 360)

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = ('电脑：{0}    型号：{1}' -f [string]$record.computer_name, [string]$record.model)
    $titleLabel.Location = New-Object System.Drawing.Point(15, 15)
    $titleLabel.Size = New-Object System.Drawing.Size(700, 24)
    $historyForm.Controls.Add($titleLabel)

    $historyGrid = New-Object System.Windows.Forms.DataGridView
    $historyGrid.Location = New-Object System.Drawing.Point(15, 50)
    $historyGrid.Size = New-Object System.Drawing.Size(710, 310)
    $historyGrid.BackgroundColor = [System.Drawing.Color]::White
    $historyGrid.BorderStyle = 'FixedSingle'
    $historyGrid.AllowUserToAddRows = $false
    $historyGrid.AllowUserToDeleteRows = $false
    $historyGrid.AllowUserToResizeRows = $false
    $historyGrid.MultiSelect = $false
    $historyGrid.SelectionMode = 'FullRowSelect'
    $historyGrid.ReadOnly = $true
    $historyGrid.RowHeadersVisible = $false
    $historyGrid.AutoSizeColumnsMode = 'Fill'
    $historyForm.Controls.Add($historyGrid)

    [void]$historyGrid.Columns.Add('colChangedAt', '变更时间')
    [void]$historyGrid.Columns.Add('colOldOwner', '原归属人')
    [void]$historyGrid.Columns.Add('colNewOwner', '新归属人')
    $historyGrid.Columns['colChangedAt'].FillWeight = 120

    foreach ($entry in @($record.owner_history | Sort-Object changed_at)) {
        $index = $historyGrid.Rows.Add()
        $row = $historyGrid.Rows[$index]
        $row.Cells['colChangedAt'].Value = [string]$entry.changed_at
        $row.Cells['colOldOwner'].Value = Resolve-HistoryOwnerName -OwnerId ([string]$entry.old_owner_id) -StoredName ([string]$entry.old_owner_name)
        $row.Cells['colNewOwner'].Value = Resolve-HistoryOwnerName -OwnerId ([string]$entry.new_owner_id) -StoredName ([string]$entry.new_owner_name)
    }

    if ($historyGrid.Rows.Count -eq 0) {
        $index = $historyGrid.Rows.Add()
        $historyGrid.Rows[$index].Cells['colNewOwner'].Value = '暂无归属人变更记录'
    }

    [void]$historyForm.ShowDialog($form)
}
function Export-Computers {
    if ($script:Computers.Count -eq 0) {
        Show-WarningMessage '当前没有可导出的电脑数据。'
        return
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = 'CSV 文件 (*.csv)|*.csv'
    $dialog.FileName = '电脑信息_{0}.csv' -f (Get-Date).ToString('yyyyMMdd_HHmmss')

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $exportRows = foreach ($item in ($script:Computers | Sort-Object computer_name, serial_number)) {
        [PSCustomObject]@{
            '电脑名称' = [string]$item.computer_name
            '序列号' = [string]$item.serial_number
            '固定资产号' = [string]$item.asset_number
            '型号' = [string]$item.model
            'MAC地址' = [string]$item.mac_address
            '归属人' = Get-OwnerLabel -OwnerId ([string]$item.owner_id)
            '备注' = [string]$item.remark
            '更新时间' = [string]$item.updated_at
        }
    }

    $exportRows | Export-Csv -Path $dialog.FileName -NoTypeInformation -Encoding UTF8
    Show-InfoMessage "导出成功：`n$($dialog.FileName)"
}

function Set-SafeSplitterLayout {
    param(
        [Parameter(Mandatory = $true)]$SplitContainer,
        [int]$Panel2MinSize,
        [int]$DesiredSplitterDistance,
        [int]$Panel1MinSize = 120
    )

    $availableWidth = [int]$SplitContainer.ClientSize.Width
    if ($availableWidth -le 0) {
        return
    }

    $safePanel2Min = [Math]::Min($Panel2MinSize, [Math]::Max(0, $availableWidth - $Panel1MinSize))
    $SplitContainer.Panel1MinSize = $Panel1MinSize
    $SplitContainer.Panel2MinSize = $safePanel2Min

    $maxDistance = [Math]::Max($Panel1MinSize, $availableWidth - $safePanel2Min)
    $safeDistance = [Math]::Max($Panel1MinSize, [Math]::Min($DesiredSplitterDistance, $maxDistance))
    $SplitContainer.SplitterDistance = $safeDistance
}

function Set-SafeStackedSplitterLayout {
    param(
        [Parameter(Mandatory = $true)]$SplitContainer,
        [int]$Panel2MinSize,
        [int]$DesiredSplitterDistance,
        [int]$Panel1MinSize = 140
    )

    $availableHeight = [int]$SplitContainer.ClientSize.Height
    if ($availableHeight -le 0) {
        return
    }

    $safePanel2Min = [Math]::Min($Panel2MinSize, [Math]::Max(0, $availableHeight - $Panel1MinSize))
    $SplitContainer.Panel1MinSize = $Panel1MinSize
    $SplitContainer.Panel2MinSize = $safePanel2Min

    $maxDistance = [Math]::Max($Panel1MinSize, $availableHeight - $safePanel2Min)
    $safeDistance = [Math]::Max($Panel1MinSize, [Math]::Min($DesiredSplitterDistance, $maxDistance))
    $SplitContainer.SplitterDistance = $safeDistance
}
function Open-InventoryManager {
    $manager = New-Object System.Windows.Forms.Form
    $manager.Text = '库存电脑管理'
    $manager.StartPosition = 'CenterParent'
    $manager.Size = New-Object System.Drawing.Size(1320, 760)
    $manager.MinimumSize = New-Object System.Drawing.Size(1120, 680)
    $manager.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Dock = 'Fill'
    $split.FixedPanel = 'Panel2'
    $split.SplitterWidth = 8
    $manager.Controls.Add($split)

    $groupList = New-Object System.Windows.Forms.GroupBox
    $groupList.Text = '库存电脑列表'
    $groupList.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)
    $groupList.Dock = 'Fill'
    $split.Panel1.Controls.Add($groupList)

    $txtSearchInv = New-Object System.Windows.Forms.TextBox
    $txtSearchInv.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
    $txtSearchInv.Location = New-Object System.Drawing.Point(18, 35)
    $txtSearchInv.Size = New-Object System.Drawing.Size(260, 28)
    $txtSearchInv.Anchor = 'Top,Left,Right'
    $groupList.Controls.Add($txtSearchInv)

    $btnSearchInv = New-Object System.Windows.Forms.Button
    $btnSearchInv.Text = '搜索'
    $btnSearchInv.Size = New-Object System.Drawing.Size(78, 32)
    $btnSearchInv.FlatStyle = 'Flat'
    $btnSearchInv.Anchor = 'Top,Right'
    $groupList.Controls.Add($btnSearchInv)

    $lblCountInv = New-Object System.Windows.Forms.Label
    $lblCountInv.Text = '共 0 台库存电脑'
    $lblCountInv.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $lblCountInv.ForeColor = [System.Drawing.Color]::FromArgb(95, 99, 104)
    $lblCountInv.TextAlign = 'MiddleRight'
    $lblCountInv.Anchor = 'Top,Right'
    $groupList.Controls.Add($lblCountInv)

    $gridInv = New-Object System.Windows.Forms.DataGridView
    $gridInv.Location = New-Object System.Drawing.Point(18, 82)
    $gridInv.Size = New-Object System.Drawing.Size(780, 590)
    $gridInv.Anchor = 'Top,Bottom,Left,Right'
    $gridInv.BackgroundColor = [System.Drawing.Color]::White
    $gridInv.BorderStyle = 'FixedSingle'
    $gridInv.AllowUserToAddRows = $false
    $gridInv.AllowUserToDeleteRows = $false
    $gridInv.AllowUserToResizeRows = $false
    $gridInv.MultiSelect = $false
    $gridInv.SelectionMode = 'FullRowSelect'
    $gridInv.ReadOnly = $true
    $gridInv.RowHeadersVisible = $false
    $gridInv.AutoSizeColumnsMode = 'Fill'
    $gridInv.ColumnHeadersHeight = 34
    $groupList.Controls.Add($gridInv)

    [void]$gridInv.Columns.Add('colName', '电脑名称')
    [void]$gridInv.Columns.Add('colSerial', '序列号')
    [void]$gridInv.Columns.Add('colAsset', '固定资产号')
    [void]$gridInv.Columns.Add('colModel', '型号')
    [void]$gridInv.Columns.Add('colMac', 'MAC 地址')
    [void]$gridInv.Columns.Add('colRemark', '备注')
    [void]$gridInv.Columns.Add('colUpdated', '更新时间')

    $groupForm = New-Object System.Windows.Forms.GroupBox
    $groupForm.Text = '库存电脑登记'
    $groupForm.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)
    $groupForm.Dock = 'Fill'
    $split.Panel2.Controls.Add($groupForm)

    $lblModeInv = New-Object System.Windows.Forms.Label
    $lblModeInv.Text = '当前模式：新增'
    $lblModeInv.Location = New-Object System.Drawing.Point(20, 32)
    $lblModeInv.Size = New-Object System.Drawing.Size(160, 24)
    $groupForm.Controls.Add($lblModeInv)

    $lblHintInv = New-Object System.Windows.Forms.Label
    $lblHintInv.Text = '这里登记的是暂无归属同事的电脑；也可以在这里直接分配给指定同事。'
    $lblHintInv.Location = New-Object System.Drawing.Point(20, 58)
    $lblHintInv.Size = New-Object System.Drawing.Size(380, 42)
    $lblHintInv.ForeColor = [System.Drawing.Color]::FromArgb(95, 99, 104)
    $groupForm.Controls.Add($lblHintInv)

    $selectedInventoryOwnerId = ''
    $suppressInventoryOwnerTextChange = $false

    function New-InventoryLabel {
        param([string]$Text, [int]$Y)
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $Text
        $label.Location = New-Object System.Drawing.Point(20, $Y)
        $label.Size = New-Object System.Drawing.Size(260, 24)
        return $label
    }

    function New-InventoryTextbox {
        param([int]$Y)
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
        $textBox.Location = New-Object System.Drawing.Point(20, $Y)
        $textBox.Size = New-Object System.Drawing.Size(430, 28)
        $textBox.Anchor = 'Top,Left,Right'
        return $textBox
    }

    $groupForm.Controls.Add((New-InventoryLabel -Text '电脑名称' -Y 112))
    $txtNameInv = New-InventoryTextbox -Y 136
    $groupForm.Controls.Add($txtNameInv)
    $groupForm.Controls.Add((New-InventoryLabel -Text '序列号' -Y 168))
    $txtSerialInv = New-InventoryTextbox -Y 192
    $groupForm.Controls.Add($txtSerialInv)
    $groupForm.Controls.Add((New-InventoryLabel -Text '固定资产号' -Y 224))
    $txtAssetInv = New-InventoryTextbox -Y 248
    $groupForm.Controls.Add($txtAssetInv)
    $groupForm.Controls.Add((New-InventoryLabel -Text '型号（可选可填）' -Y 280))

    $cmbModelInv = New-Object System.Windows.Forms.ComboBox
    $cmbModelInv.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
    $cmbModelInv.Location = New-Object System.Drawing.Point(20, 304)
    $cmbModelInv.Size = New-Object System.Drawing.Size(430, 28)
    $cmbModelInv.Anchor = 'Top,Left,Right'
    $cmbModelInv.DropDownStyle = 'DropDown'
    $groupForm.Controls.Add($cmbModelInv)

    $groupForm.Controls.Add((New-InventoryLabel -Text 'MAC 地址' -Y 336))
    $txtMacInv = New-InventoryTextbox -Y 360
    $groupForm.Controls.Add($txtMacInv)
    $groupForm.Controls.Add((New-InventoryLabel -Text '归属同事（输入拼音匹配，可留空）' -Y 392))
    $txtOwnerInv = New-InventoryTextbox -Y 416
    $groupForm.Controls.Add($txtOwnerInv)

    $lstOwnerSuggestionsInv = New-Object System.Windows.Forms.ListBox
    $lstOwnerSuggestionsInv.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $lstOwnerSuggestionsInv.Location = New-Object System.Drawing.Point(20, 446)
    $lstOwnerSuggestionsInv.Size = New-Object System.Drawing.Size(430, 72)
    $lstOwnerSuggestionsInv.Anchor = 'Top,Left,Right'
    $lstOwnerSuggestionsInv.Visible = $false
    $groupForm.Controls.Add($lstOwnerSuggestionsInv)

    $groupForm.Controls.Add((New-InventoryLabel -Text '备注' -Y 526))
    $txtRemarkInv = New-Object System.Windows.Forms.TextBox
    $txtRemarkInv.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
    $txtRemarkInv.Location = New-Object System.Drawing.Point(20, 550)
    $txtRemarkInv.Size = New-Object System.Drawing.Size(430, 120)
    $txtRemarkInv.Multiline = $true
    $txtRemarkInv.ScrollBars = 'Vertical'
    $txtRemarkInv.Anchor = 'Top,Bottom,Left,Right'
    $groupForm.Controls.Add($txtRemarkInv)

    $btnSaveInv = New-Object System.Windows.Forms.Button
    $btnSaveInv.Text = '保存库存电脑'
    $btnSaveInv.Size = New-Object System.Drawing.Size(130, 34)
    $btnSaveInv.FlatStyle = 'Flat'
    $btnSaveInv.BackColor = [System.Drawing.Color]::FromArgb(30, 136, 229)
    $btnSaveInv.ForeColor = [System.Drawing.Color]::White
    $btnSaveInv.Anchor = 'Bottom,Left'
    $groupForm.Controls.Add($btnSaveInv)

    $btnClearInv = New-Object System.Windows.Forms.Button
    $btnClearInv.Text = '清空表单'
    $btnClearInv.Size = New-Object System.Drawing.Size(110, 34)
    $btnClearInv.FlatStyle = 'Flat'
    $btnClearInv.Anchor = 'Bottom,Left'
    $groupForm.Controls.Add($btnClearInv)

    $btnDeleteInv = New-Object System.Windows.Forms.Button
    $btnDeleteInv.Text = '删除库存电脑'
    $btnDeleteInv.Size = New-Object System.Drawing.Size(130, 34)
    $btnDeleteInv.FlatStyle = 'Flat'
    $btnDeleteInv.Anchor = 'Bottom,Right'
    $groupForm.Controls.Add($btnDeleteInv)
    function Refresh-InventoryModelOptions {
        $currentText = $cmbModelInv.Text
        $models = @($script:Computers | ForEach-Object { [string]$_.model } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
        $cmbModelInv.BeginUpdate()
        $cmbModelInv.Items.Clear()
        foreach ($modelName in $models) { [void]$cmbModelInv.Items.Add($modelName) }
        $cmbModelInv.Text = $currentText
        $cmbModelInv.EndUpdate()
    }

    function Set-InventorySelectedOwner {
        param($Colleague)

        $suppressInventoryOwnerTextChange = $true
        if ($null -eq $Colleague) {
            $selectedInventoryOwnerId = ''
            $txtOwnerInv.Text = ''
        } else {
            $selectedInventoryOwnerId = [string]$Colleague.id
            $txtOwnerInv.Text = Format-ColleagueOption -Colleague $Colleague
        }
        $suppressInventoryOwnerTextChange = $false
        $lstOwnerSuggestionsInv.Visible = $false
    }

    function Refresh-InventoryOwnerSuggestions {
        $keyword = $txtOwnerInv.Text.Trim().ToLowerInvariant()
        $lstOwnerSuggestionsInv.Items.Clear()

        if ([string]::IsNullOrWhiteSpace($keyword)) {
            $lstOwnerSuggestionsInv.Visible = $false
            return
        }

        $matches = @($script:Colleagues | Where-Object {
            ([string]$_.pinyin).ToLowerInvariant().Contains($keyword) -or
            ([string]$_.display_name).ToLowerInvariant().Contains($keyword) -or
            ([string]$_.department).ToLowerInvariant().Contains($keyword)
        } | Sort-Object display_name, pinyin, department)

        foreach ($item in $matches) {
            [void]$lstOwnerSuggestionsInv.Items.Add([PSCustomObject]@{
                Id = [string]$item.id
                Display = Format-ColleagueOption -Colleague $item
            })
        }

        $lstOwnerSuggestionsInv.DisplayMember = 'Display'
        $lstOwnerSuggestionsInv.ValueMember = 'Id'
        $lstOwnerSuggestionsInv.Visible = $lstOwnerSuggestionsInv.Items.Count -gt 0
        if ($lstOwnerSuggestionsInv.Visible) {
            $lstOwnerSuggestionsInv.Height = [Math]::Min(110, 24 * $lstOwnerSuggestionsInv.Items.Count + 4)
        }
    }

    function Get-FilteredInventoryComputers {
        $keyword = $txtSearchInv.Text.Trim().ToLowerInvariant()
        $rows = foreach ($item in ($script:Computers | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.owner_id) })) {
            [PSCustomObject]@{
                id = [string]$item.id
                computer_name = [string]$item.computer_name
                serial_number = [string]$item.serial_number
                asset_number = [string]$item.asset_number
                model = [string]$item.model
                mac_address = [string]$item.mac_address
                remark = [string]$item.remark
                updated_at = [string]$item.updated_at
            }
        }

        if ([string]::IsNullOrWhiteSpace($keyword)) { return @($rows | Sort-Object computer_name, serial_number) }

        return @($rows | Where-Object {
            $_.computer_name.ToLowerInvariant().Contains($keyword) -or
            $_.serial_number.ToLowerInvariant().Contains($keyword) -or
            $_.asset_number.ToLowerInvariant().Contains($keyword) -or
            $_.model.ToLowerInvariant().Contains($keyword) -or
            $_.mac_address.ToLowerInvariant().Contains($keyword) -or
            $_.remark.ToLowerInvariant().Contains($keyword)
        } | Sort-Object computer_name, serial_number)
    }

    function Refresh-InventoryGrid {
        $gridInv.Rows.Clear()
        $rows = Get-FilteredInventoryComputers
        foreach ($rowData in $rows) {
            $index = $gridInv.Rows.Add()
            $row = $gridInv.Rows[$index]
            $row.Tag = $rowData.id
            $row.Cells['colName'].Value = $rowData.computer_name
            $row.Cells['colSerial'].Value = $rowData.serial_number
            $row.Cells['colAsset'].Value = $rowData.asset_number
            $row.Cells['colModel'].Value = $rowData.model
            $row.Cells['colMac'].Value = $rowData.mac_address
            $row.Cells['colRemark'].Value = $rowData.remark
            $row.Cells['colUpdated'].Value = $rowData.updated_at
        }
        $lblCountInv.Text = "共 $($rows.Count) 台库存电脑"
    }

    function Clear-InventoryForm {
        $script:CurrentInventoryComputerId = $null
        $txtNameInv.Text = ''
        $txtSerialInv.Text = ''
        $txtAssetInv.Text = ''
        $cmbModelInv.Text = ''
        $txtMacInv.Text = ''
        Set-InventorySelectedOwner -Colleague $null
        $txtRemarkInv.Text = ''
        $lblModeInv.Text = '当前模式：新增'
        $gridInv.ClearSelection()
    }

    function Fill-InventoryForm {
        param([string]$ComputerId)
        $record = $script:Computers | Where-Object { $_.id -eq $ComputerId } | Select-Object -First 1
        if ($null -eq $record -or -not [string]::IsNullOrWhiteSpace([string]$record.owner_id)) { return }

        $script:CurrentInventoryComputerId = [string]$record.id
        $txtNameInv.Text = [string]$record.computer_name
        $txtSerialInv.Text = [string]$record.serial_number
        $txtAssetInv.Text = [string]$record.asset_number
        $cmbModelInv.Text = [string]$record.model
        $txtMacInv.Text = [string]$record.mac_address
        Set-InventorySelectedOwner -Colleague (Get-ColleagueById -Id ([string]$record.owner_id))
        $txtRemarkInv.Text = [string]$record.remark
        $lblModeInv.Text = '当前模式：编辑'
    }

    function Save-InventoryComputer {
        if (-not (Test-ComputerFieldValues -Name $txtNameInv.Text -Serial $txtSerialInv.Text -Asset $txtAssetInv.Text -Mac $txtMacInv.Text)) { return }

        $name = $txtNameInv.Text.Trim()
        $serial = $txtSerialInv.Text.Trim()
        $asset = $txtAssetInv.Text.Trim()
        $model = $cmbModelInv.Text.Trim()
        $mac = $txtMacInv.Text.Trim()
        $ownerText = $txtOwnerInv.Text.Trim()
        $ownerId = [string]$selectedInventoryOwnerId
        if (-not [string]::IsNullOrWhiteSpace($ownerText) -and [string]::IsNullOrWhiteSpace($ownerId)) {
            $resolvedOwner = Resolve-ColleagueFromOwnerInput -InputText $ownerText
            if ($null -ne $resolvedOwner) {
                $ownerId = [string]$resolvedOwner.id
                Set-InventorySelectedOwner -Colleague $resolvedOwner
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($ownerText) -and [string]::IsNullOrWhiteSpace($ownerId)) {
            Show-WarningMessage '归属同事未正确匹配，请从联想列表中重新选择一位人员。'
            return
        }
        $remark = $txtRemarkInv.Text.Trim()
        $now = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

        $targetInventoryId = [string]$script:CurrentInventoryComputerId
        if ([string]::IsNullOrWhiteSpace($targetInventoryId) -and $gridInv.SelectedRows.Count -gt 0) {
            $targetInventoryId = [string]$gridInv.SelectedRows[0].Tag
        }

        $duplicate = $script:Computers | Where-Object {
            $_.id -ne $targetInventoryId -and (
                [string]$_.serial_number -eq $serial -or
                [string]$_.asset_number -eq $asset
            )
        } | Select-Object -First 1

        if ($duplicate) {
            Show-WarningMessage -Title '重复数据' -Message '序列号或固定资产号已存在，请确认后再保存。'
            return
        }

        if ([string]::IsNullOrWhiteSpace($targetInventoryId)) {
            $script:Computers = @($script:Computers) + [PSCustomObject]@{
                id = [guid]::NewGuid().ToString()
                computer_name = $name
                serial_number = $serial
                asset_number = $asset
                model = $model
                mac_address = $mac
                owner_id = $ownerId
                owner_history = if ([string]::IsNullOrWhiteSpace($ownerId)) {
                    @()
                } else {
                    @([PSCustomObject]@{
                        changed_at = $now
                        old_owner_id = ''
                        old_owner_name = '库存'
                        new_owner_id = $ownerId
                        new_owner_name = Get-OwnerLabel -OwnerId $ownerId
                    })
                }
                remark = $remark
                updated_at = $now
            }
            $targetInventoryId = [string]$script:Computers[-1].id
        } else {
            if (-not (Request-EditAuthorization)) { return }
            foreach ($item in $script:Computers) {
                if ($item.id -eq $targetInventoryId) {
                    $item.computer_name = $name
                    $item.serial_number = $serial
                    $item.asset_number = $asset
                    $item.model = $model
                    $item.mac_address = $mac
                    [void](Set-ComputerOwner -ComputerRecord $item -NewOwnerId $ownerId -ChangedAt $now)
                    $item.remark = $remark
                    $item.updated_at = $now
                    break
                }
            }
        }

        $script:CurrentInventoryComputerId = $targetInventoryId

        Save-Computers
        Refresh-ModelOptions
        Refresh-InventoryModelOptions
        Refresh-ComputerGrid
        Refresh-InventoryGrid

        if ([string]::IsNullOrWhiteSpace($ownerId)) {
            Clear-InventoryForm
            Show-InfoMessage '库存电脑已保存。'
        } else {
            [void](Select-MainComputerRow -ComputerId $targetInventoryId)
            Clear-InventoryForm
            Show-InfoMessage '库存电脑已分配给指定同事，并已同步到主界面。'
        }
    }
    function Remove-InventoryComputer {
        if ($gridInv.SelectedRows.Count -eq 0) {
            Show-WarningMessage '请先选择要删除的库存电脑。'
            return
        }

        $selectedId = [string]$gridInv.SelectedRows[0].Tag
        $result = [System.Windows.Forms.MessageBox]::Show(
            '确定删除这条库存电脑记录吗？',
            '确认删除',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $script:Computers = @($script:Computers | Where-Object { $_.id -ne $selectedId })
        Save-Computers
        Refresh-ModelOptions
        Refresh-InventoryModelOptions
        Refresh-ComputerGrid
        Refresh-InventoryGrid
        Clear-InventoryForm
    }

    function Update-InventoryLayout {
        $listWidth = [int]$groupList.ClientRectangle.Width
        $listHeight = [int]$groupList.ClientRectangle.Height
        $panelWidth = [int]$groupForm.ClientRectangle.Width
        $panelHeight = [int]$groupForm.ClientRectangle.Height
        $editorWidth = $panelWidth - 40

        $btnSearchInv.Location = New-Object System.Drawing.Point(($listWidth - 96), 33)
        $txtSearchInv.Width = [Math]::Max(260, ($btnSearchInv.Left - 36))
        $lblCountInv.Location = New-Object System.Drawing.Point(($listWidth - 180), 68)
        $lblCountInv.Size = New-Object System.Drawing.Size(162, 20)
        $gridInv.Size = New-Object System.Drawing.Size(($listWidth - 36), ($listHeight - 100))

        $txtNameInv.Width = $editorWidth
        $txtSerialInv.Width = $editorWidth
        $txtAssetInv.Width = $editorWidth
        $cmbModelInv.Width = $editorWidth
        $txtMacInv.Width = $editorWidth
        $txtOwnerInv.Width = $editorWidth
        $lstOwnerSuggestionsInv.Width = $editorWidth
        $txtRemarkInv.Width = $editorWidth
        $txtRemarkInv.Height = [Math]::Max(100, ($panelHeight - 640))

        $buttonY = $panelHeight - 52
        $btnSaveInv.Location = New-Object System.Drawing.Point(20, $buttonY)
        $btnClearInv.Location = New-Object System.Drawing.Point(160, $buttonY)
        $btnDeleteInv.Location = New-Object System.Drawing.Point(($panelWidth - 150), $buttonY)
    }

    $btnSearchInv.Add_Click({ Refresh-InventoryGrid })
    $txtSearchInv.Add_TextChanged({ Refresh-InventoryGrid })
    $txtOwnerInv.Add_TextChanged({
        if ($suppressInventoryOwnerTextChange) { return }
        $selectedOwner = Get-ColleagueById -Id ([string]$selectedInventoryOwnerId)
        $selectedText = if ($null -eq $selectedOwner) { '' } else { Format-ColleagueOption -Colleague $selectedOwner }
        if ($txtOwnerInv.Text.Trim() -ne $selectedText) { $selectedInventoryOwnerId = '' }
        Refresh-InventoryOwnerSuggestions
    })
    $txtOwnerInv.Add_Leave({
        if (-not $lstOwnerSuggestionsInv.Focused) {
            $manager.BeginInvoke([Action]{ $lstOwnerSuggestionsInv.Visible = $false }) | Out-Null
        }
    })
    $lstOwnerSuggestionsInv.Add_DoubleClick({ if ($null -ne $lstOwnerSuggestionsInv.SelectedItem) { Set-InventorySelectedOwner -Colleague (Get-ColleagueById -Id ([string]$lstOwnerSuggestionsInv.SelectedItem.Id)) } })
    $lstOwnerSuggestionsInv.Add_Click({ if ($null -ne $lstOwnerSuggestionsInv.SelectedItem) { Set-InventorySelectedOwner -Colleague (Get-ColleagueById -Id ([string]$lstOwnerSuggestionsInv.SelectedItem.Id)) } })
    $lstOwnerSuggestionsInv.Add_KeyDown({ param($sender, $e) if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter -and $null -ne $lstOwnerSuggestionsInv.SelectedItem) { Set-InventorySelectedOwner -Colleague (Get-ColleagueById -Id ([string]$lstOwnerSuggestionsInv.SelectedItem.Id)); $e.Handled = $true } })
    $btnSaveInv.Add_Click({ Save-InventoryComputer })
    $btnClearInv.Add_Click({ Clear-InventoryForm })
    $btnDeleteInv.Add_Click({ Remove-InventoryComputer })
    $gridInv.Add_SelectionChanged({ if ($gridInv.SelectedRows.Count -gt 0) { Fill-InventoryForm -ComputerId ([string]$gridInv.SelectedRows[0].Tag) } })

    $manager.Add_Shown({ Set-SafeSplitterLayout -SplitContainer $split -Panel2MinSize 420 -DesiredSplitterDistance 840; Refresh-InventoryModelOptions; Refresh-InventoryGrid; Clear-InventoryForm; Update-InventoryLayout })
    $manager.Add_Resize({ Set-SafeSplitterLayout -SplitContainer $split -Panel2MinSize 420 -DesiredSplitterDistance 840; Update-InventoryLayout })

    [void]$manager.ShowDialog($form)
    Refresh-ModelOptions
    Refresh-ComputerGrid
    Clear-ComputerForm
}
function Open-ColleagueManager {
    $manager = New-Object System.Windows.Forms.Form
    $manager.Text = '人员名单管理'
    $manager.StartPosition = 'CenterParent'
    $manager.Size = New-Object System.Drawing.Size(1320, 760)
    $manager.MinimumSize = New-Object System.Drawing.Size(1120, 680)
    $manager.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Dock = 'Fill'
    $split.FixedPanel = 'Panel2'
    $split.SplitterWidth = 8
    $manager.Controls.Add($split)

    $groupList = New-Object System.Windows.Forms.GroupBox
    $groupList.Text = '人员列表'
    $groupList.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)
    $groupList.Dock = 'Fill'
    $split.Panel1.Controls.Add($groupList)

    $txtSearchCol = New-Object System.Windows.Forms.TextBox
    $txtSearchCol.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
    $txtSearchCol.Location = New-Object System.Drawing.Point(18, 35)
    $txtSearchCol.Size = New-Object System.Drawing.Size(260, 28)
    $txtSearchCol.Anchor = 'Top,Left,Right'
    $groupList.Controls.Add($txtSearchCol)

    $btnSearchCol = New-Object System.Windows.Forms.Button
    $btnSearchCol.Text = '搜索'
    $btnSearchCol.Size = New-Object System.Drawing.Size(78, 32)
    $btnSearchCol.FlatStyle = 'Flat'
    $btnSearchCol.Anchor = 'Top,Right'
    $groupList.Controls.Add($btnSearchCol)

    $lblCountCol = New-Object System.Windows.Forms.Label
    $lblCountCol.Text = '共 0 人'
    $lblCountCol.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $lblCountCol.ForeColor = [System.Drawing.Color]::FromArgb(95, 99, 104)
    $lblCountCol.TextAlign = 'MiddleRight'
    $lblCountCol.Anchor = 'Top,Right'
    $groupList.Controls.Add($lblCountCol)

    $gridCol = New-Object System.Windows.Forms.DataGridView
    $gridCol.Location = New-Object System.Drawing.Point(18, 82)
    $gridCol.Size = New-Object System.Drawing.Size(780, 590)
    $gridCol.Anchor = 'Top,Bottom,Left,Right'
    $gridCol.BackgroundColor = [System.Drawing.Color]::White
    $gridCol.BorderStyle = 'FixedSingle'
    $gridCol.AllowUserToAddRows = $false
    $gridCol.AllowUserToDeleteRows = $false
    $gridCol.AllowUserToResizeRows = $false
    $gridCol.MultiSelect = $false
    $gridCol.SelectionMode = 'FullRowSelect'
    $gridCol.ReadOnly = $true
    $gridCol.RowHeadersVisible = $false
    $gridCol.AutoSizeColumnsMode = 'Fill'
    $gridCol.ColumnHeadersHeight = 34
    $groupList.Controls.Add($gridCol)

    [void]$gridCol.Columns.Add('colDisplayName', '中文名')
    [void]$gridCol.Columns.Add('colPinyin', '拼音')
    [void]$gridCol.Columns.Add('colEmail', '邮箱')
    [void]$gridCol.Columns.Add('colDepartment', '部门')
    [void]$gridCol.Columns.Add('colEmployeeType', '类型')
    [void]$gridCol.Columns.Add('colMentor', 'Mentor')

    $groupForm = New-Object System.Windows.Forms.GroupBox
    $groupForm.Text = '人员信息编辑'
    $groupForm.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)
    $groupForm.Dock = 'Fill'
    $split.Panel2.Controls.Add($groupForm)

    $lblModeCol = New-Object System.Windows.Forms.Label
    $lblModeCol.Text = '当前模式：新增'
    $lblModeCol.Location = New-Object System.Drawing.Point(20, 32)
    $lblModeCol.Size = New-Object System.Drawing.Size(160, 24)
    $groupForm.Controls.Add($lblModeCol)

    function New-ColLabel {
        param([string]$Text, [int]$Y)
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $Text
        $label.Location = New-Object System.Drawing.Point(20, $Y)
        $label.Size = New-Object System.Drawing.Size(220, 24)
        return $label
    }

    function New-ColTextbox {
        param([int]$Y)
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
        $textBox.Location = New-Object System.Drawing.Point(20, $Y)
        $textBox.Size = New-Object System.Drawing.Size(350, 28)
        $textBox.Anchor = 'Top,Left,Right'
        return $textBox
    }

    $groupForm.Controls.Add((New-ColLabel -Text '中文名' -Y 68))
    $txtDisplayName = New-ColTextbox -Y 92
    $groupForm.Controls.Add($txtDisplayName)
    $groupForm.Controls.Add((New-ColLabel -Text '拼音' -Y 124))
    $txtPinyin = New-ColTextbox -Y 148
    $groupForm.Controls.Add($txtPinyin)

    $btnGeneratePinyin = New-Object System.Windows.Forms.Button
    $btnGeneratePinyin.Text = '重新生成'
    $btnGeneratePinyin.Size = New-Object System.Drawing.Size(90, 28)
    $btnGeneratePinyin.FlatStyle = 'Flat'
    $btnGeneratePinyin.Anchor = 'Top,Right'
    $groupForm.Controls.Add($btnGeneratePinyin)

    $groupForm.Controls.Add((New-ColLabel -Text '邮箱' -Y 180))
    $txtEmail = New-ColTextbox -Y 204
    $txtEmail.Text = '@itk-engineering.com'
    $groupForm.Controls.Add($txtEmail)
    $groupForm.Controls.Add((New-ColLabel -Text '部门' -Y 236))

    $cmbDepartment = New-Object System.Windows.Forms.ComboBox
    $cmbDepartment.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
    $cmbDepartment.Location = New-Object System.Drawing.Point(20, 260)
    $cmbDepartment.Size = New-Object System.Drawing.Size(350, 28)
    $cmbDepartment.Anchor = 'Top,Left,Right'
    $cmbDepartment.DropDownStyle = 'DropDown'
    $groupForm.Controls.Add($cmbDepartment)

    $groupForm.Controls.Add((New-ColLabel -Text '员工类型' -Y 292))
    $cmbEmployeeType = New-Object System.Windows.Forms.ComboBox
    $cmbEmployeeType.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
    $cmbEmployeeType.Location = New-Object System.Drawing.Point(20, 316)
    $cmbEmployeeType.Size = New-Object System.Drawing.Size(350, 28)
    $cmbEmployeeType.Anchor = 'Top,Left,Right'
    $cmbEmployeeType.DropDownStyle = 'DropDownList'
    [void]$cmbEmployeeType.Items.Add('正式员工')
    [void]$cmbEmployeeType.Items.Add('实习生')
    $groupForm.Controls.Add($cmbEmployeeType)

    $lblMentor = New-ColLabel -Text 'Mentor' -Y 348
    $groupForm.Controls.Add($lblMentor)
    $cmbMentor = New-Object System.Windows.Forms.ComboBox
    $cmbMentor.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
    $cmbMentor.Location = New-Object System.Drawing.Point(20, 372)
    $cmbMentor.Size = New-Object System.Drawing.Size(350, 28)
    $cmbMentor.Anchor = 'Top,Left,Right'
    $cmbMentor.DropDownStyle = 'DropDownList'
    $groupForm.Controls.Add($cmbMentor)

    $btnSaveCol = New-Object System.Windows.Forms.Button
    $btnSaveCol.Text = '保存人员'
    $btnSaveCol.Size = New-Object System.Drawing.Size(110, 34)
    $btnSaveCol.FlatStyle = 'Flat'
    $btnSaveCol.BackColor = [System.Drawing.Color]::FromArgb(30, 136, 229)
    $btnSaveCol.ForeColor = [System.Drawing.Color]::White
    $btnSaveCol.Anchor = 'Bottom,Left'
    $groupForm.Controls.Add($btnSaveCol)

    $btnClearCol = New-Object System.Windows.Forms.Button
    $btnClearCol.Text = '清空表单'
    $btnClearCol.Size = New-Object System.Drawing.Size(110, 34)
    $btnClearCol.FlatStyle = 'Flat'
    $btnClearCol.Anchor = 'Bottom,Left'
    $groupForm.Controls.Add($btnClearCol)

    $btnDeleteCol = New-Object System.Windows.Forms.Button
    $btnDeleteCol.Text = '删除选中人员'
    $btnDeleteCol.Size = New-Object System.Drawing.Size(110, 34)
    $btnDeleteCol.FlatStyle = 'Flat'
    $btnDeleteCol.Anchor = 'Bottom,Right'
    $groupForm.Controls.Add($btnDeleteCol)
    function Refresh-DepartmentOptions {
        $currentText = $cmbDepartment.Text
        $departments = @($script:Colleagues | ForEach-Object { [string]$_.department } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
        $cmbDepartment.BeginUpdate()
        $cmbDepartment.Items.Clear()
        foreach ($departmentName in $departments) { [void]$cmbDepartment.Items.Add($departmentName) }
        $cmbDepartment.Text = $currentText
        $cmbDepartment.EndUpdate()
    }

    function Refresh-MentorOptions {
        $selectedMentorId = if ($cmbMentor.SelectedItem -and $cmbMentor.SelectedItem.Tag) { [string]$cmbMentor.SelectedItem.Tag } else { '' }
        $mentorOptions = @($script:Colleagues | Where-Object { $_.employee_type -eq '正式员工' -and $_.id -ne $script:CurrentColleagueEditorId } | Sort-Object display_name, department)

        $cmbMentor.BeginUpdate()
        $cmbMentor.Items.Clear()
        [void]$cmbMentor.Items.Add([PSCustomObject]@{ Label = '请选择 Mentor'; Tag = '' })
        foreach ($mentor in $mentorOptions) {
            [void]$cmbMentor.Items.Add([PSCustomObject]@{ Label = Format-ColleagueOption -Colleague $mentor; Tag = [string]$mentor.id })
        }
        $cmbMentor.DisplayMember = 'Label'
        $cmbMentor.ValueMember = 'Tag'
        $cmbMentor.EndUpdate()

        $matchedItem = $null
        foreach ($item in $cmbMentor.Items) {
            if ([string]$item.Tag -eq $selectedMentorId) { $matchedItem = $item; break }
        }
        if ($null -eq $matchedItem -and $cmbMentor.Items.Count -gt 0) { $matchedItem = $cmbMentor.Items[0] }
        $cmbMentor.SelectedItem = $matchedItem
    }

    function Update-MentorState {
        $isIntern = ($cmbEmployeeType.SelectedItem -eq '实习生')
        $cmbMentor.Enabled = $isIntern
        $lblMentor.ForeColor = if ($isIntern) { [System.Drawing.Color]::Black } else { [System.Drawing.Color]::Gray }
        if (-not $isIntern -and $cmbMentor.Items.Count -gt 0) { $cmbMentor.SelectedIndex = 0 }
    }

    function Update-ColleaguePinyinFromName {
        if ($script:ColleaguePinyinManuallyEdited) { return }
        $script:SuppressColleagueAutoPinyin = $true
        $txtPinyin.Text = Get-AutoPinyinText -Text $txtDisplayName.Text
        $script:SuppressColleagueAutoPinyin = $false
    }

    function Clear-ColleagueForm {
        $script:CurrentColleagueEditorId = $null
        $script:ColleaguePinyinManuallyEdited = $false
        $script:SuppressColleagueAutoPinyin = $true
        $txtDisplayName.Text = ''
        $txtPinyin.Text = ''
        $txtEmail.Text = '@itk-engineering.com'
        $cmbDepartment.Text = ''
        $cmbEmployeeType.SelectedItem = '正式员工'
        Refresh-MentorOptions
        Update-MentorState
        $lblModeCol.Text = '当前模式：新增'
        $gridCol.ClearSelection()
        $script:SuppressColleagueAutoPinyin = $false
    }

    function Get-FilteredColleagues {
        $keyword = $txtSearchCol.Text.Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($keyword)) { return @($script:Colleagues | Sort-Object display_name, pinyin, department) }

        return @($script:Colleagues | Where-Object {
            $mentorName = Get-OwnerDisplayName -OwnerId ([string]$_.mentor_id)
            ([string]$_.display_name).ToLowerInvariant().Contains($keyword) -or
            ([string]$_.pinyin).ToLowerInvariant().Contains($keyword) -or
            ([string]$_.email).ToLowerInvariant().Contains($keyword) -or
            ([string]$_.department).ToLowerInvariant().Contains($keyword) -or
            ([string]$_.employee_type).ToLowerInvariant().Contains($keyword) -or
            $mentorName.ToLowerInvariant().Contains($keyword)
        } | Sort-Object display_name, pinyin, department)
    }

    function Refresh-ColleagueGrid {
        $gridCol.Rows.Clear()
        $rows = Get-FilteredColleagues
        foreach ($item in $rows) {
            $index = $gridCol.Rows.Add()
            $row = $gridCol.Rows[$index]
            $row.Tag = [string]$item.id
            $row.Cells['colDisplayName'].Value = [string]$item.display_name
            $row.Cells['colPinyin'].Value = [string]$item.pinyin
            $row.Cells['colEmail'].Value = [string]$item.email
            $row.Cells['colDepartment'].Value = [string]$item.department
            $row.Cells['colEmployeeType'].Value = [string]$item.employee_type
            $row.Cells['colMentor'].Value = Get-OwnerDisplayName -OwnerId ([string]$item.mentor_id)
        }
        $lblCountCol.Text = "共 $($rows.Count) 人"
    }

    function Fill-ColleagueForm {
        param([string]$Id)
        $item = $script:Colleagues | Where-Object { $_.id -eq $Id } | Select-Object -First 1
        if ($null -eq $item) { return }

        $script:CurrentColleagueEditorId = [string]$item.id
        $script:SuppressColleagueAutoPinyin = $true
        $txtDisplayName.Text = [string]$item.display_name
        $txtPinyin.Text = [string]$item.pinyin
        $txtEmail.Text = [string]$item.email
        $cmbDepartment.Text = [string]$item.department
        $cmbEmployeeType.SelectedItem = [string]$item.employee_type
        Refresh-MentorOptions

        $targetMentorId = [string]$item.mentor_id
        $matchedMentor = $null
        foreach ($mentorItem in $cmbMentor.Items) {
            if ([string]$mentorItem.Tag -eq $targetMentorId) { $matchedMentor = $mentorItem; break }
        }
        if ($null -eq $matchedMentor -and $cmbMentor.Items.Count -gt 0) { $matchedMentor = $cmbMentor.Items[0] }
        $cmbMentor.SelectedItem = $matchedMentor
        $script:ColleaguePinyinManuallyEdited = $true
        $script:SuppressColleagueAutoPinyin = $false
        Update-MentorState
        $lblModeCol.Text = '当前模式：编辑'
    }

    function Save-ColleagueRecord {
        $displayName = $txtDisplayName.Text.Trim()
        $pinyin = $txtPinyin.Text.Trim().ToLowerInvariant()
        $email = $txtEmail.Text.Trim()
        $department = $cmbDepartment.Text.Trim()
        $employeeType = [string]$cmbEmployeeType.SelectedItem
        $mentorId = if ($cmbMentor.SelectedItem) { [string]$cmbMentor.SelectedItem.Tag } else { '' }

        if ([string]::IsNullOrWhiteSpace($displayName)) { Show-WarningMessage '请输入人员中文名。'; return }
        if ([string]::IsNullOrWhiteSpace($pinyin)) { Show-WarningMessage '请输入人员拼音。'; return }
        if ([string]::IsNullOrWhiteSpace($email)) { Show-WarningMessage '请输入邮箱。'; return }
        if ($email -notmatch '@') { Show-WarningMessage '邮箱格式不正确，请至少包含 @。'; return }
        if ([string]::IsNullOrWhiteSpace($department)) { Show-WarningMessage '请输入部门。'; return }
        if ($employeeType -ne '正式员工' -and $employeeType -ne '实习生') { Show-WarningMessage '请选择员工类型。'; return }

        if ($employeeType -eq '实习生') {
            if ([string]::IsNullOrWhiteSpace($mentorId)) { Show-WarningMessage '实习生必须关联一位正式员工 Mentor。'; return }
            $mentorRecord = Get-ColleagueById -Id $mentorId
            if ($null -eq $mentorRecord -or [string]$mentorRecord.employee_type -ne '正式员工') {
                Show-WarningMessage 'Mentor 必须是已存在的正式员工。'
                return
            }
        } else {
            $mentorId = ''
        }

        if (-not [string]::IsNullOrWhiteSpace($script:CurrentColleagueEditorId)) {
            $hasDependents = $script:Colleagues | Where-Object { $_.mentor_id -eq $script:CurrentColleagueEditorId } | Select-Object -First 1
            if ($hasDependents -and $employeeType -ne '正式员工') {
                Show-WarningMessage '该人员当前已被其他实习生关联为 Mentor，请先调整相关 Mentor 关系。'
                return
            }
        }

        if ([string]::IsNullOrWhiteSpace($script:CurrentColleagueEditorId)) {
            $script:Colleagues = @($script:Colleagues) + [PSCustomObject]@{
                id = [guid]::NewGuid().ToString()
                display_name = $displayName
                pinyin = $pinyin
                email = $email
                department = $department
                employee_type = $employeeType
                mentor_id = $mentorId
            }
        } else {
            foreach ($colleague in $script:Colleagues) {
                if ($colleague.id -eq $script:CurrentColleagueEditorId) {
                    $colleague.display_name = $displayName
                    $colleague.pinyin = $pinyin
                    $colleague.email = $email
                    $colleague.department = $department
                    $colleague.employee_type = $employeeType
                    $colleague.mentor_id = $mentorId
                    break
                }
            }
        }

        Save-Colleagues
        Refresh-DepartmentOptions
        Refresh-MentorOptions
        Refresh-ColleagueGrid
        Refresh-ComputerGrid
        Refresh-OwnerSuggestions
        Clear-ColleagueForm
        Show-InfoMessage '人员信息已保存。'
    }

    function Remove-ColleagueRecord {
        if ($gridCol.SelectedRows.Count -eq 0) { Show-WarningMessage '请先选择要删除的人员。'; return }

        $selectedId = [string]$gridCol.SelectedRows[0].Tag
        $usedAsMentor = $script:Colleagues | Where-Object { $_.mentor_id -eq $selectedId } | Select-Object -First 1
        if ($usedAsMentor) {
            Show-WarningMessage '该人员仍被实习生关联为 Mentor，请先调整相关 Mentor 关系后再删除。'
            return
        }

        $ownedComputers = @($script:Computers | Where-Object { $_.owner_id -eq $selectedId })
        $ownedCount = $ownedComputers.Count
        $message = if ($ownedCount -gt 0) { "确定删除这位人员吗？`n`n该人员名下的 $ownedCount 台电脑会自动转入库存。" } else { '确定删除这位人员吗？' }

        $result = [System.Windows.Forms.MessageBox]::Show($message, '确认删除', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        if ($ownedCount -gt 0) {
            $now = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            foreach ($computer in $ownedComputers) {
                [void](Set-ComputerOwner -ComputerRecord $computer -NewOwnerId '' -ChangedAt $now)
                $computer.updated_at = $now
            }
            Save-Computers
        }

        $script:Colleagues = @($script:Colleagues | Where-Object { $_.id -ne $selectedId })
        Save-Colleagues
        Refresh-DepartmentOptions
        Refresh-MentorOptions
        Refresh-ColleagueGrid
        Refresh-ComputerGrid
        Refresh-OwnerSuggestions
        Refresh-ModelOptions
        Clear-ColleagueForm
        if ($ownedCount -gt 0) { Show-InfoMessage "人员已删除，原名下的 $ownedCount 台电脑已转入库存。" }
    }

    function Update-ColleagueLayout {
        $listWidth = [int]$groupList.ClientRectangle.Width
        $listHeight = [int]$groupList.ClientRectangle.Height
        $panelWidth = [int]$groupForm.ClientRectangle.Width
        $panelHeight = [int]$groupForm.ClientRectangle.Height
        $editorWidth = $panelWidth - 40

        $btnSearchCol.Location = New-Object System.Drawing.Point(($listWidth - 96), 33)
        $txtSearchCol.Width = [Math]::Max(260, ($btnSearchCol.Left - 36))
        $lblCountCol.Location = New-Object System.Drawing.Point(($listWidth - 138), 68)
        $lblCountCol.Size = New-Object System.Drawing.Size(120, 20)
        $gridCol.Size = New-Object System.Drawing.Size(($listWidth - 36), ($listHeight - 100))

        $btnGeneratePinyin.Location = New-Object System.Drawing.Point(($panelWidth - 110), 146)
        $txtDisplayName.Width = $editorWidth
        $txtPinyin.Width = $editorWidth - 100
        $txtEmail.Width = $editorWidth
        $cmbDepartment.Width = $editorWidth
        $cmbEmployeeType.Width = $editorWidth
        $cmbMentor.Width = $editorWidth

        $buttonY = $panelHeight - 52
        $btnSaveCol.Location = New-Object System.Drawing.Point(20, $buttonY)
        $btnClearCol.Location = New-Object System.Drawing.Point(150, $buttonY)
        $btnDeleteCol.Location = New-Object System.Drawing.Point(($panelWidth - 130), $buttonY)
    }

    $btnSearchCol.Add_Click({ Refresh-ColleagueGrid })
    $txtSearchCol.Add_TextChanged({ Refresh-ColleagueGrid })
    $cmbEmployeeType.Add_SelectedIndexChanged({ Update-MentorState })
    $txtDisplayName.Add_TextChanged({ if (-not $script:SuppressColleagueAutoPinyin) { Update-ColleaguePinyinFromName } })
    $txtPinyin.Add_TextChanged({ if (-not $script:SuppressColleagueAutoPinyin -and $txtPinyin.Focused) { $script:ColleaguePinyinManuallyEdited = $true } })
    $btnGeneratePinyin.Add_Click({ $script:ColleaguePinyinManuallyEdited = $false; Update-ColleaguePinyinFromName })
    $gridCol.Add_SelectionChanged({ if ($gridCol.SelectedRows.Count -gt 0) { Fill-ColleagueForm -Id ([string]$gridCol.SelectedRows[0].Tag) } })
    $btnSaveCol.Add_Click({ Save-ColleagueRecord })
    $btnClearCol.Add_Click({ Clear-ColleagueForm })
    $btnDeleteCol.Add_Click({ Remove-ColleagueRecord })

    $manager.Add_Shown({ Set-SafeSplitterLayout -SplitContainer $split -Panel2MinSize 420 -DesiredSplitterDistance 840; Refresh-DepartmentOptions; Refresh-MentorOptions; Refresh-ColleagueGrid; Clear-ColleagueForm; Update-ColleagueLayout })
    $manager.Add_Resize({ Set-SafeSplitterLayout -SplitContainer $split -Panel2MinSize 420 -DesiredSplitterDistance 840; Update-ColleagueLayout })

    [void]$manager.ShowDialog($form)
    Refresh-ComputerGrid
    Clear-ComputerForm
}
$form = New-Object System.Windows.Forms.Form
$form.Text = 'ITK China 电脑信息管理系统'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1520, 860)
$form.MinimumSize = New-Object System.Drawing.Size(1260, 720)
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

$logoPath = Join-Path $script:RootDir 'ITK_Logo_RGB.jpg'
if (Test-Path $logoPath) {
    $logoBox = New-Object System.Windows.Forms.PictureBox
    $logoBox.Location = New-Object System.Drawing.Point(20, 16)
    $logoBox.Size = New-Object System.Drawing.Size(120, 56)
    $logoBox.SizeMode = 'Zoom'
    $logoBox.Image = [System.Drawing.Image]::FromFile($logoPath)
    $form.Controls.Add($logoBox)
}

$title = New-Object System.Windows.Forms.Label
$title.Text = 'ITK China 电脑信息管理系统'
$title.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 18, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(152, 14)
$title.Size = New-Object System.Drawing.Size(520, 40)
$form.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = 'ITK China 设备资产、人员与库存电脑管理平台'
$subtitle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$subtitle.ForeColor = [System.Drawing.Color]::FromArgb(95, 99, 104)
$subtitle.Location = New-Object System.Drawing.Point(154, 50)
$subtitle.Size = New-Object System.Drawing.Size(520, 24)
$form.Controls.Add($subtitle)

$splitMain = New-Object System.Windows.Forms.SplitContainer
$splitMain.Location = New-Object System.Drawing.Point(20, 95)
$splitMain.Size = New-Object System.Drawing.Size(1460, 720)
$splitMain.Anchor = 'Top,Bottom,Left,Right'
$splitMain.FixedPanel = 'Panel2'
$splitMain.SplitterWidth = 8
$form.Controls.Add($splitMain)

$groupList = New-Object System.Windows.Forms.GroupBox
$groupList.Text = '电脑列表'
$groupList.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)
$groupList.Dock = 'Fill'
$splitMain.Panel1.Controls.Add($groupList)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
$txtSearch.Location = New-Object System.Drawing.Point(18, 35)
$txtSearch.Size = New-Object System.Drawing.Size(260, 28)
$txtSearch.Anchor = 'Top,Left,Right'
$groupList.Controls.Add($txtSearch)

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = '搜索'
$btnSearch.Size = New-Object System.Drawing.Size(78, 32)
$btnSearch.FlatStyle = 'Flat'
$btnSearch.Anchor = 'Top,Right'
$groupList.Controls.Add($btnSearch)

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = '转入库存'
$btnDelete.Size = New-Object System.Drawing.Size(96, 32)
$btnDelete.FlatStyle = 'Flat'
$btnDelete.Anchor = 'Top,Right'
$groupList.Controls.Add($btnDelete)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = '导出 CSV'
$btnExport.Size = New-Object System.Drawing.Size(96, 32)
$btnExport.FlatStyle = 'Flat'
$btnExport.Anchor = 'Top,Right'
$groupList.Controls.Add($btnExport)

$btnInventoryManager = New-Object System.Windows.Forms.Button
$btnInventoryManager.Text = '库存电脑管理'
$btnInventoryManager.Size = New-Object System.Drawing.Size(120, 32)
$btnInventoryManager.FlatStyle = 'Flat'
$btnInventoryManager.Anchor = 'Top,Right'
$groupList.Controls.Add($btnInventoryManager)

$btnColleagueManager = New-Object System.Windows.Forms.Button
$btnColleagueManager.Text = '人员名单管理'
$btnColleagueManager.Size = New-Object System.Drawing.Size(120, 32)
$btnColleagueManager.FlatStyle = 'Flat'
$btnColleagueManager.Anchor = 'Top,Right'
$groupList.Controls.Add($btnColleagueManager)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = '在用 0 台，库存 0 台，共 0 台'
$lblCount.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$lblCount.ForeColor = [System.Drawing.Color]::FromArgb(95, 99, 104)
$lblCount.TextAlign = 'MiddleRight'
$lblCount.Anchor = 'Top,Right'
$groupList.Controls.Add($lblCount)

$splitComputerLists = New-Object System.Windows.Forms.SplitContainer
$splitComputerLists.Location = New-Object System.Drawing.Point(18, 92)
$splitComputerLists.Size = New-Object System.Drawing.Size(880, 600)
$splitComputerLists.Anchor = 'Top,Bottom,Left,Right'
$splitComputerLists.Orientation = 'Horizontal'
$splitComputerLists.SplitterWidth = 8
$groupList.Controls.Add($splitComputerLists)

$groupInUse = New-Object System.Windows.Forms.GroupBox
$groupInUse.Text = '在用电脑（0 台）'
$groupInUse.Dock = 'Fill'
$groupInUse.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9, [System.Drawing.FontStyle]::Bold)
$splitComputerLists.Panel1.Controls.Add($groupInUse)

$groupInventoryMain = New-Object System.Windows.Forms.GroupBox
$groupInventoryMain.Text = '库存电脑（0 台）'
$groupInventoryMain.Dock = 'Fill'
$groupInventoryMain.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9, [System.Drawing.FontStyle]::Bold)
$splitComputerLists.Panel2.Controls.Add($groupInventoryMain)

$gridInUse = New-Object System.Windows.Forms.DataGridView
$gridInUse.Dock = 'Fill'
$gridInUse.BackgroundColor = [System.Drawing.Color]::White
$gridInUse.BorderStyle = 'FixedSingle'
$gridInUse.AllowUserToAddRows = $false
$gridInUse.AllowUserToDeleteRows = $false
$gridInUse.AllowUserToResizeRows = $false
$gridInUse.MultiSelect = $false
$gridInUse.SelectionMode = 'FullRowSelect'
$gridInUse.ReadOnly = $true
$gridInUse.RowHeadersVisible = $false
$gridInUse.AutoSizeColumnsMode = 'Fill'
$gridInUse.ColumnHeadersHeight = 34
$gridInUse.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(227, 242, 253)
$gridInUse.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
$groupInUse.Controls.Add($gridInUse)
Add-MainComputerGridColumns -TargetGrid $gridInUse

$gridInventoryMain = New-Object System.Windows.Forms.DataGridView
$gridInventoryMain.Dock = 'Fill'
$gridInventoryMain.BackgroundColor = [System.Drawing.Color]::White
$gridInventoryMain.BorderStyle = 'FixedSingle'
$gridInventoryMain.AllowUserToAddRows = $false
$gridInventoryMain.AllowUserToDeleteRows = $false
$gridInventoryMain.AllowUserToResizeRows = $false
$gridInventoryMain.MultiSelect = $false
$gridInventoryMain.SelectionMode = 'FullRowSelect'
$gridInventoryMain.ReadOnly = $true
$gridInventoryMain.RowHeadersVisible = $false
$gridInventoryMain.AutoSizeColumnsMode = 'Fill'
$gridInventoryMain.ColumnHeadersHeight = 34
$gridInventoryMain.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(227, 242, 253)
$gridInventoryMain.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
$groupInventoryMain.Controls.Add($gridInventoryMain)
Add-MainComputerGridColumns -TargetGrid $gridInventoryMain
$groupForm = New-Object System.Windows.Forms.GroupBox
$groupForm.Text = '电脑信息编辑'
$groupForm.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)
$groupForm.Dock = 'Fill'
$splitMain.Panel2.Controls.Add($groupForm)

$lblMode = New-Object System.Windows.Forms.Label
$lblMode.Text = '当前模式：新增'
$lblMode.Location = New-Object System.Drawing.Point(20, 32)
$lblMode.Size = New-Object System.Drawing.Size(160, 24)
$groupForm.Controls.Add($lblMode)

$btnOwnerHistory = New-Object System.Windows.Forms.Button
$btnOwnerHistory.Text = '归属历史'
$btnOwnerHistory.Size = New-Object System.Drawing.Size(120, 30)
$btnOwnerHistory.Anchor = 'Top,Right'
$btnOwnerHistory.FlatStyle = 'Flat'
$groupForm.Controls.Add($btnOwnerHistory)

function New-EditorLabel {
    param([string]$Text, [int]$Y)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point(20, $Y)
    $label.Size = New-Object System.Drawing.Size(220, 24)
    return $label
}

function New-EditorTextbox {
    param([int]$Y)
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
    $textBox.Location = New-Object System.Drawing.Point(20, $Y)
    $textBox.Size = New-Object System.Drawing.Size(430, 28)
    $textBox.Anchor = 'Top,Left,Right'
    return $textBox
}

$groupForm.Controls.Add((New-EditorLabel -Text '电脑名称' -Y 68))
$txtName = New-EditorTextbox -Y 92
$groupForm.Controls.Add($txtName)
$groupForm.Controls.Add((New-EditorLabel -Text '序列号' -Y 124))
$txtSerial = New-EditorTextbox -Y 148
$groupForm.Controls.Add($txtSerial)
$groupForm.Controls.Add((New-EditorLabel -Text '固定资产号' -Y 180))
$txtAsset = New-EditorTextbox -Y 204
$groupForm.Controls.Add($txtAsset)
$groupForm.Controls.Add((New-EditorLabel -Text '型号（可选可填）' -Y 236))

$cmbModel = New-Object System.Windows.Forms.ComboBox
$cmbModel.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
$cmbModel.Location = New-Object System.Drawing.Point(20, 260)
$cmbModel.Size = New-Object System.Drawing.Size(430, 28)
$cmbModel.Anchor = 'Top,Left,Right'
$cmbModel.DropDownStyle = 'DropDown'
$groupForm.Controls.Add($cmbModel)

$groupForm.Controls.Add((New-EditorLabel -Text 'MAC 地址' -Y 292))
$txtMac = New-EditorTextbox -Y 316
$groupForm.Controls.Add($txtMac)
$groupForm.Controls.Add((New-EditorLabel -Text '归属人（输入拼音匹配，可留空表示库存）' -Y 348))
$txtOwner = New-EditorTextbox -Y 372
$groupForm.Controls.Add($txtOwner)

$lstOwnerSuggestions = New-Object System.Windows.Forms.ListBox
$lstOwnerSuggestions.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$lstOwnerSuggestions.Location = New-Object System.Drawing.Point(20, 402)
$lstOwnerSuggestions.Size = New-Object System.Drawing.Size(430, 72)
$lstOwnerSuggestions.Anchor = 'Top,Left,Right'
$lstOwnerSuggestions.Visible = $false
$groupForm.Controls.Add($lstOwnerSuggestions)

$groupForm.Controls.Add((New-EditorLabel -Text '备注' -Y 482))
$txtRemark = New-Object System.Windows.Forms.TextBox
$txtRemark.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
$txtRemark.Location = New-Object System.Drawing.Point(20, 506)
$txtRemark.Size = New-Object System.Drawing.Size(430, 120)
$txtRemark.Multiline = $true
$txtRemark.ScrollBars = 'Vertical'
$txtRemark.Anchor = 'Top,Bottom,Left,Right'
$groupForm.Controls.Add($txtRemark)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = '保存信息'
$btnSave.Size = New-Object System.Drawing.Size(200, 32)
$btnSave.FlatStyle = 'Flat'
$btnSave.BackColor = [System.Drawing.Color]::FromArgb(30, 136, 229)
$btnSave.ForeColor = [System.Drawing.Color]::White
$btnSave.Anchor = 'Bottom,Left'
$groupForm.Controls.Add($btnSave)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = '清空表单'
$btnClear.Size = New-Object System.Drawing.Size(200, 32)
$btnClear.FlatStyle = 'Flat'
$btnClear.Anchor = 'Bottom,Right'
$groupForm.Controls.Add($btnClear)

function Update-MainLayout {
    $toolbarY = 33
    $listWidth = [int]$groupList.ClientRectangle.Width
    $listHeight = [int]$groupList.ClientRectangle.Height
    $panelWidth = [int]$groupForm.ClientRectangle.Width
    $panelHeight = [int]$groupForm.ClientRectangle.Height
    $editorWidth = $panelWidth - 40

    $btnColleagueManager.Location = New-Object System.Drawing.Point(($listWidth - 138), $toolbarY)
    $btnInventoryManager.Location = New-Object System.Drawing.Point(($btnColleagueManager.Left - 128), $toolbarY)
    $btnExport.Location = New-Object System.Drawing.Point(($btnInventoryManager.Left - 104), $toolbarY)
    $btnDelete.Location = New-Object System.Drawing.Point(($btnExport.Left - 118), $toolbarY)
    $btnSearch.Location = New-Object System.Drawing.Point(($btnDelete.Left - 90), $toolbarY)

    $txtSearch.Width = [Math]::Max(260, ($btnSearch.Left - 36))
    $lblCount.Location = New-Object System.Drawing.Point(($listWidth - 350), 68)
    $lblCount.Size = New-Object System.Drawing.Size(332, 20)
    $splitComputerLists.Size = New-Object System.Drawing.Size(($listWidth - 36), ($listHeight - 110))
    Set-SafeStackedSplitterLayout -SplitContainer $splitComputerLists -Panel2MinSize 160 -DesiredSplitterDistance ([Math]::Max(180, [int](($splitComputerLists.Height - 8) * 0.58))) -Panel1MinSize 180

    $btnOwnerHistory.Location = New-Object System.Drawing.Point(($panelWidth - 140), 28)
    $txtName.Width = $editorWidth
    $txtSerial.Width = $editorWidth
    $txtAsset.Width = $editorWidth
    $cmbModel.Width = $editorWidth
    $txtMac.Width = $editorWidth
    $txtOwner.Width = $editorWidth
    $lstOwnerSuggestions.Width = $editorWidth
    $txtRemark.Width = $editorWidth
    $txtRemark.Height = [Math]::Max(100, ($panelHeight - 590))

    $buttonY = $panelHeight - 50
    $buttonWidth = [int][Math]::Floor((($editorWidth - 20) / 2))
    $btnSave.Location = New-Object System.Drawing.Point(20, $buttonY)
    $btnSave.Size = New-Object System.Drawing.Size($buttonWidth, 32)
    $btnClear.Location = New-Object System.Drawing.Point(($btnSave.Right + 20), $buttonY)
    $btnClear.Size = New-Object System.Drawing.Size($buttonWidth, 32)
}

$btnSearch.Add_Click({ Refresh-ComputerGrid })
$txtSearch.Add_TextChanged({ Refresh-ComputerGrid })
$btnSave.Add_Click({ Save-ComputerRecord })
$btnClear.Add_Click({ Clear-ComputerForm })
$btnDelete.Add_Click({ Remove-SelectedComputer })
$btnExport.Add_Click({ Export-Computers })
$btnInventoryManager.Add_Click({ Open-InventoryManager })
$btnColleagueManager.Add_Click({ Open-ColleagueManager })
$btnOwnerHistory.Add_Click({ Show-OwnerHistory })
$gridInUse.Add_SelectionChanged({
    if ($gridInUse.SelectedRows.Count -gt 0) {
        $gridInventoryMain.ClearSelection()
        Fill-ComputerForm -ComputerId ([string]$gridInUse.SelectedRows[0].Tag)
    }
})
$gridInventoryMain.Add_SelectionChanged({
    if ($gridInventoryMain.SelectedRows.Count -gt 0) {
        $gridInUse.ClearSelection()
        Fill-ComputerForm -ComputerId ([string]$gridInventoryMain.SelectedRows[0].Tag)
    }
})
$txtOwner.Add_TextChanged({
    if ($script:SuppressOwnerTextChange) { return }
    $selectedOwner = Get-ColleagueById -Id ([string]$script:SelectedOwnerId)
    $selectedText = if ($null -eq $selectedOwner) { '' } else { Format-ColleagueOption -Colleague $selectedOwner }
    if ($txtOwner.Text.Trim() -ne $selectedText) { $script:SelectedOwnerId = $null }
    Refresh-OwnerSuggestions
})
$txtOwner.Add_Leave({ if (-not $lstOwnerSuggestions.Focused) { $form.BeginInvoke([Action]{ $lstOwnerSuggestions.Visible = $false }) | Out-Null } })
$lstOwnerSuggestions.Add_DoubleClick({ if ($null -ne $lstOwnerSuggestions.SelectedItem) { Set-SelectedOwner -Colleague (Get-ColleagueById -Id ([string]$lstOwnerSuggestions.SelectedItem.Id)) } })
$lstOwnerSuggestions.Add_Click({ if ($null -ne $lstOwnerSuggestions.SelectedItem) { Set-SelectedOwner -Colleague (Get-ColleagueById -Id ([string]$lstOwnerSuggestions.SelectedItem.Id)) } })
$lstOwnerSuggestions.Add_KeyDown({ param($sender, $e) if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter -and $null -ne $lstOwnerSuggestions.SelectedItem) { Set-SelectedOwner -Colleague (Get-ColleagueById -Id ([string]$lstOwnerSuggestions.SelectedItem.Id)); $e.Handled = $true } })

$form.Add_Shown({ Set-SafeSplitterLayout -SplitContainer $splitMain -Panel2MinSize 420 -DesiredSplitterDistance 950; Update-MainLayout; Load-AllData; Refresh-ModelOptions; Refresh-ComputerGrid; Clear-ComputerForm })
$form.Add_Resize({ Set-SafeSplitterLayout -SplitContainer $splitMain -Panel2MinSize 420 -DesiredSplitterDistance 950; Update-MainLayout })

[void]$form.ShowDialog()






