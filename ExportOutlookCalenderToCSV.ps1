Param(
    $startDate = "01-01-2019",
    $endDate = "31-12-2019",
    $exportCsvLocation = "export_calendar.csv"
)

function Get-WeekNumber([datetime]$DateTime = (Get-Date)) {
    $cultureInfo = [System.Globalization.CultureInfo]::CurrentCulture
    $cultureInfo.Calendar.GetWeekOfYear($DateTime,$cultureInfo.DateTimeFormat.CalendarWeekRule,$cultureInfo.DateTimeFormat.FirstDayOfWeek)
}

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')
$store = $namespace.DefaultStore
$calendar = $store.GetDefaultFolder(9)

$rangeFilter = "[Start] >= '+ $startDate +' AND [END] <= '+ $endDate +'"
$items = $calendar.Items.Restrict($rangeFilter)
$items.Sort("[Start]") 

If (Test-Path $exportCsvLocation){
	Remove-Item $exportCsvLocation
}
$header = '"Subject";"Category";"Startdate";"Enddate";"Duration (in hours)";"Weeknumber"'
$header | Out-File $exportCsvLocation -Append -Encoding utf8

ForEach($item in $items){
    $weekNumber = Get-WeekNumber -DateTime $item.Start
    $row = '"' + $item.Subject + '";"' + $item.Categories + '";"' + $item.Start + '";"' + $item.End + '";"' + ($item.Duration / 60).ToString("#.##") + '";"' + $weekNumber + '"'
    $row | Out-File $exportCsvLocation -Append -Encoding utf8
}