Param(
    $startDate = "01-01-2019",
    $endDate = "31-12-2019",
    $exportCsvLocation = "export_calendar.csv",
    $excelLocation = "registratie.xlsx"
)

function Get-WeekNumber([datetime]$DateTime = (Get-Date)) {
    $cultureInfo = [System.Globalization.CultureInfo]::CurrentCulture
    $cultureInfo.Calendar.GetWeekOfYear($DateTime,$cultureInfo.DateTimeFormat.CalendarWeekRule,$cultureInfo.DateTimeFormat.FirstDayOfWeek)
}

Write-Output "Open Outlook calendar"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')
$store = $namespace.DefaultStore
$calendar = $store.GetDefaultFolder(9)

$rangeFilter = "[Start] >= '+ $startDate +' AND [END] <= '+ $endDate +'"

Write-Output "Get items in range: " $rangeFilter
$items = $calendar.Items.Restrict($rangeFilter)
$items.Sort("[Start]") 

Write-Output "Export items to CSV file: " $exportCsvLocation
If (Test-Path $exportCsvLocation){
	Remove-Item $exportCsvLocation
}
$header = '"Subject";"Category";"Startdate";"Enddate";"Duration (in hours)";"Weeknumber"'
$header | Out-File $exportCsvLocation -Append -Encoding utf8

ForEach($item in $items) {
    $weekNumber = Get-WeekNumber -DateTime $item.Start
    $duration = If ($item.Duration -eq 1440) {8} Else {($item.Duration / 60)}
    $row = '"' + $item.Subject + '";"' + $item.Categories + '";"' + $item.Start + '";"' + $item.End + '";"' + $duration.ToString("#.##") + '";"' + $weekNumber + '"'
    $row | Out-File $exportCsvLocation -Append -Encoding utf8
}

Write-Output "Open Excelsheet: " $excelLocation
$excel = New-Object -Com Excel.Application
$workbook = $excel.Workbooks.Open($excelLocation) 

Write-Output "Update data connections"
$connections = $workbook.Connections  
foreach ($c in $connections) {  
    if ($null -ne $c.DataFeedConnection)  
    {  
            $c.DataFeedConnection.Refresh()
            while ($c.DataFeedConnection.Refreshing)
            {
                Start-Sleep -Seconds 1
            }
    }  
}  

Write-Output "Refresh All the pivot tables data." 
$workbook.RefreshAll()  

# Need to wait till workbook is refreshed before saving...
Start-Sleep -Seconds 30

Write-Output "Save & quit"
$workbook.Save()  
$workbook.Close()  
$excel.quit()