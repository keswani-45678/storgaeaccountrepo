

$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true


$InputFilename = Get-Content 'D:\test.csv'


$wb = $xl.Workbooks.Open("D:\test.csv")
$ws = $wb.Sheets.Item(1)

# $rows = $ws.UsedRange.Rows.Count
$rows = $InputFilename.Length


$columnCount = $ws.UsedRange.Columns.Count



$OutputFilenamePattern = "arc_netstat_report_202111_part_"

$line = 0
$file = 0
$start = 0

 
$myDataField1 = New-Object Collections.Generic.List[String]
$myDataField2 = New-Object Collections.Generic.List[String]
$myDataField3 = New-Object Collections.Generic.List[String]
$myDataField4 = New-Object Collections.Generic.List[String]
       $myDataField5  = New-Object Collections.Generic.List[String]
       $myDataField6  = New-Object Collections.Generic.List[String]
       $myDataField7  = New-Object Collections.Generic.List[String]
       $myDataField8  = New-Object Collections.Generic.List[String]
       $myDataField9  = New-Object Collections.Generic.List[String]
       $myDataField10  = New-Object Collections.Generic.List[String]
       $myDataField11  = New-Object Collections.Generic.List[String]
       $myDataField12  = New-Object Collections.Generic.List[String]
       $myDataField13  = New-Object Collections.Generic.List[String]
       $myDataField14  = New-Object Collections.Generic.List[String]
       $myDataField15  = New-Object Collections.Generic.List[String]
$myDataField16  = New-Object Collections.Generic.List[String]

$my1stColumn1=1
$my1stColumn2=2
$my1stColumn3=3
$my1stColumn4=4
$my1stColumn5=5
$my1stColumn6=6
$my1stColumn7=7
$my1stColumn8=8
$my1stColumn9=9
$my1stColumn10=10
$my1stColumn11=11
$my1stColumn12=12
$my1stColumn13=13
$my1stColumn14=14
$my1stColumn15=15





    #$myDataField1 | Out-File $Filename -Force

# # $ListItemCollection = @() 
# $myDataField16.Add($myDataField1)
# $myDataField16.Add($myDataField2)
# #     for ([int]$i = 0; $i -le $rows; $i++)
# #     # {
    
# $ListItemCollection += $myDataField16
#     $ListItemCollection | Out-File $Filename -Force

#     # }


for ([int]$i = 0; $i -le $rows; $i++) {
    $value = $ws.Cells.Item($row, 2).Text
    if($value -eq "12.COM"){
   
     
Write-Host "test" $value

#while($value -eq '12.COM'){
$file++
$Filename = "$OutputFilenamePattern$file.csv"
$InputFilename[$start..($line-1)] | Out-File $Filename -Force
$start = $line;
Write-Host "$Filename"
 $line++       
}

}

 

$wb.Close()
$xl.Quit()
