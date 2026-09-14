$ErrorActionPreference = 'Stop'
$root = (Split-Path $PSScriptRoot -Parent)
$assembly = [Reflection.Assembly]::LoadFrom("$root\src\SemanticTable\bin\x64\Release\net48\SemanticTable.dll")
$store = $assembly.GetType('SemanticTable.StateStore')
function Load-State($book, $table) { $store.GetMethod('Load').Invoke($null, @($book, $table.Name)) }
function Save-State($book, $state) { $store.GetMethod('Save').Invoke($null, @($book, $state)) }
function Check($book, $table, $label) {
    $state = Load-State $book $table
    if ($state.ExcelTableName -ne $table.Name -or $state.Filters.Count -ne 2 -or $state.Filters[0].Values.Count -ne 500 -or $state.Filters[1].Value2 -ne '200') { throw "Failed: $label" }
    if (-not $table.AlternativeText.StartsWith('User description')) { throw 'Description lost' }
    Write-Output "PASS $label"
}
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Range('A1').Value2 = 'Country'
    $sheet.Range('A2').Value2 = 'Test'
    $table = $sheet.ListObjects.Add(1, $sheet.Range('A1:A2'), $null, 1)
    $table.AlternativeText = 'User description'
    $state = [Activator]::CreateInstance($assembly.GetType('SemanticTable.TableDefinition'), $true)
    $state.ExcelTableName = $table.Name
    $state.DatasetId = 'test-model'
    $field = [Activator]::CreateInstance($assembly.GetType('SemanticTable.SemanticField'), $true)
    $field.Table = 'Sales'; $field.Name = 'Country'
    $state.Fields.Add($field)
    $filter = [Activator]::CreateInstance($assembly.GetType('SemanticTable.FieldFilter'), $true)
    $filter.Field = $field
    1..500 | ForEach-Object { $filter.Values.Add(('Value "quoted" ' + $_)) }
    $state.Filters.Add($filter)
    $advanced = [Activator]::CreateInstance($assembly.GetType('SemanticTable.FieldFilter'), $true)
    $advanced.Field = $field; $advanced.Mode = 'Advanced'; $advanced.Operator = 'Between'; $advanced.Value = '100'; $advanced.Value2 = '200'
    $state.Filters.Add($advanced)
    Save-State $book $state
    Check $book $table 'long definition round trip'
    $table.Range.Copy($sheet.Range('D1'))
    Check $book $sheet.ListObjects.Item(2) 'table range copy'
    $sheet.Copy($sheet)
    Check $book $book.Worksheets.Item(1).ListObjects.Item(1) 'worksheet copy'
    $other = $excel.Workbooks.Add()
    $table.Range.Copy($other.Worksheets.Item(1).Range('A1'))
    Check $other $other.Worksheets.Item(1).ListObjects.Item(1) 'cross-workbook copy'
    $other.Close($false)
    $table.Name = 'RenamedTable'
    Check $book $table 'rename'
    $path = "$root\artifacts\beta9\persistence-test.xlsx"
    New-Item -ItemType Directory -Force (Split-Path $path -Parent) | Out-Null
    $book.SaveAs($path, 51)
    $book.Close($false)
    $book = $excel.Workbooks.Open($path)
    foreach ($ws in $book.Worksheets) { foreach ($item in $ws.ListObjects) { Check $book $item 'save and reopen' } }
    $book.Close($false)
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}


