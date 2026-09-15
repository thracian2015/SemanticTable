param([string]$AssemblyPath = "$PSScriptRoot\..\src\SemanticTable\bin\x64\Release\net48\SemanticTable.dll")
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$flags = [Reflection.BindingFlags]'Instance,Public,NonPublic'
$static = [Reflection.BindingFlags]'Static,Public,NonPublic'
function New-Model($name) { [Activator]::CreateInstance($assembly.GetType("SemanticTable.$name"), $true) }
function Control($row, $name) { $row.GetType().GetField($name, $flags).GetValue($row) }
$builder = $assembly.GetType('SemanticTable.DaxQueryBuilder')
foreach ($scale in @(1,1.5,2,2.5)) {
    foreach ($type in @('String','Int64','DateTime','Boolean')) {
        $field = New-Model 'SemanticField'; $field.Table = 'Customer'; $field.Name = 'Phone'; $field.DataType = $type
        $filter = New-Model 'FieldFilter'; $filter.Field = $field; $filter.Mode = 'Advanced'
        $row = [Activator]::CreateInstance($assembly.GetType('SemanticTable.FilterRow'), $flags, $null, @($filter,$null,$null,$null,$null), $null)
        $row.Font = New-Object Drawing.Font('Microsoft Sans Serif', (8.25*$scale))
        $row.Width = [int](365*$scale)
        $hostForm = New-Object Windows.Forms.Form
        $hostForm.ShowInTaskbar = $false; $hostForm.Opacity = 0
        $hostForm.Controls.Add($row); $hostForm.ClientSize = $row.Size; $hostForm.Show(); $row.PerformLayout()
        $heights = @('_mode','_operator','_value1','_value2','_selectValues') | ForEach-Object { (Control $row $_).Height }
        if (@($heights | Select-Object -Unique).Count -ne 1) { throw "Unequal heights at scale ${scale}, ${type}: $heights" }
        $operator = Control $row '_operator'
        foreach ($op in @('Is Blank','Is Not Blank')) {
            $operator.SelectedItem = $op
            if ($filter.Operator -ne $op) { throw 'Operator missing' }
            foreach ($name in @('_value1','_value2','_selectValues')) {
                if ((Control $row $name).Enabled) { throw "$op enables $name" }
            }
            $expected = "ISBLANK('Customer'[Phone])"
            if ($op -eq 'Is Not Blank') { $expected = "NOT($expected)" }
            $signature = $builder.GetMethod('ConditionSignature',$static).Invoke($null,@($filter))
            if ($signature -ne "FILTER ( KEEPFILTERS ( VALUES ( 'Customer'[Phone] ) ), $expected )") { throw "Wrong DAX: $signature" }
            $mode = Control $row '_mode'
            $mode.SelectedItem = 'Basic'; $mode.SelectedItem = 'Advanced'
            if ($filter.Operator -ne $op -or (Control $row '_value1').Enabled) { throw 'Mode switch lost blank operator' }
        }
        $row.GetType().GetMethod('ClearSelection',$flags).Invoke($row,@())
        if ($builder.GetMethod('HasCondition',$static).Invoke($null,@($filter))) { throw 'Clear retained condition' }
        $operator.SelectedItem = 'Equals'
        if (-not (Control $row '_value1').Enabled) { throw 'Value input did not reenable' }
        if ($type -eq 'String') {
            $operator.SelectedItem = 'Is Blank'
            $output = Join-Path $PSScriptRoot '../artifacts/filter-layout'
            New-Item -ItemType Directory -Force $output | Out-Null
            $bitmap = New-Object Drawing.Bitmap($row.Width,$row.Height)
            $row.DrawToBitmap($bitmap,$row.ClientRectangle)
            $bitmap.Save((Join-Path $output "filter-$scale.png")); $bitmap.Dispose()
        }
        $hostForm.Dispose()
    }
    Write-Output "PASS blank DAX, disabled inputs, mode switching and clear at scale $scale for text, number, date and boolean"
}

