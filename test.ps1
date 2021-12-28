

$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true


$InputFilename = Get-Content 'D:\data.csv'


$wb = $xl.Workbooks.Open("D:\Trumpf\data.csv")
$ws = $wb.Sheets.Item(1)

$rows = $ws.UsedRange.Rows.Count
$rows

$OutputFilenamePattern = "arc_netstat_report_202111_part_"

$line = 0
$i = 0
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


for ($i = 1; $i -le $rows - 1; $i++)
    {
        $myDataField1.Add($ws.Cells.Item($r + $i, $my1stColumn1).text)
       $myDataField2.Add($ws.Cells.Item($r + $i, $my1stColumn2).text)
       $myDataField3.Add($ws.Cells.Item($r + $i, $my1stColumn3).text)
       $myDataField4.Add($ws.Cells.Item($r + $i, $my1stColumn4).text)
       $myDataField5.Add($ws.Cells.Item($r + $i, $my1stColumn5).text)
       $myDataField6.Add($ws.Cells.Item($r + $i, $my1stColumn6).text)
       $myDataField7.Add($ws.Cells.Item($r + $i, $my1stColumn7).text)
       $myDataField8.Add($ws.Cells.Item($r + $i, $my1stColumn8).text)
       $myDataField9.Add($ws.Cells.Item($r + $i, $my1stColumn9).text)
       $myDataField10.Add($ws.Cells.Item($r + $i, $my1stColumn10).text)
       $myDataField11.Add($ws.Cells.Item($r + $i, $my1stColumn11).text)
       $myDataField12.Add($ws.Cells.Item($r + $i, $my1stColumn12).text)
       $myDataField13.Add($ws.Cells.Item($r + $i, $my1stColumn13).text)
       $myDataField14.Add($ws.Cells.Item($r + $i, $my1stColumn14).text)
       $myDataField15.Add($ws.Cells.Item($r + $i, $my1stColumn15).text)
       
    }
$Filename = "test.csv"
   $myDataField1 | Out-File $Filename -Force







$wb.Close()
$xl.Quit()
