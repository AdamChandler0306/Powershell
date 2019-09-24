Set-Location -Path "~\Desktop"

$file = .\data.xlsx
$xl=New-Object -ComObject "Excel.Application"
$wb=$xl.Workbooks.open($file)
$ws=$wb.ActiveSheet

$Row=2
 
do {
  $data=$ws.Range("A$Row").Text
...
if ($data) {
    Write-Verbose "Querying $data" 
      $ping=Test-Connection -ComputerName $data -Quiet
if ($Ping) {
        $OS=(Get-WmiObject -Class Win32_OperatingSystem -Property Caption -computer $data).Caption
      }
      else {
        $OS=$Null
New-Object -TypeName PSObject -Property @{
        Computername=$Data.ToUpper()
        OS=$OS
        Ping=$Ping
        Location=$ws.Range("B$Row").Text
        AssetAge=((Get-Date)-($ws.Range("D$Row").Text -as [datetime])).TotalDays -as [int]
      }
$Row++
} While ($data)
$xl.displayAlerts=$False
$wb.Close()
$xl.Application.Quit()