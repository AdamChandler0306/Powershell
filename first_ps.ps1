Get-VIServer yourservername

$Excel = New-Object -Com Excel.Application
$Excel.visible = $True
$Excel = $Excel.Workbooks.Add(1)

$Sheet = $Excel.WorkSheets.Item(1)
$Sheet.Cells.Item(1,1) = "Name"
$Sheet.Cells.Item(1,2) = "Power State"
$Sheet.Cells.Item(1,3) = "Description"
$Sheet.Cells.Item(1,4) = "Number of CPUs"
$Sheet.Cells.Item(1,5) = "Memory (MB)"

$WorkBook = $Sheet.UsedRange
$WorkBook.Font.Bold = $True

$intRow = 2
$colItems = Get-VM | Select-Object -property "Name","PowerState","Description","NumCPU","MemoryMB" 

foreach ($objItem in $colItems) {
    $Sheet.Cells.Item($intRow,1) = $objItem.Name
            
    $powerstate = $objItem.PowerState
    If ($PowerState -eq 1) {$power = "Powerd On"}
        Else {$power = "Powerd Off"}
        
    $Sheet.Cells.Item($intRow,2) = $power
    $Sheet.Cells.Item($intRow,3) = $objItem.Description
    $Sheet.Cells.Item($intRow,4) = $objItem.NumCPU
    $Sheet.Cells.Item($intRow,5) = $objItem.MemoryMB

$intRow = $intRow + 1

}
$WorkBook.EntireColumn.AutoFit()
Clear
test