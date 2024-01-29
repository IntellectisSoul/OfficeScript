Today : formatDateTime(utcNow(), 'yyyy-MM-dd')
Differences : div(sub(ticks(formatDateTime(variables('LastDayOfService'), 'yyyy-MM-dd')), ticks(formatDateTime(outputs('TodayDate'), 'yyyy-MM-dd'))), 864000000000)
This Month : formatDateTime(outputs('TodayDate'), 'MM')

End Date for this Month : addDays(startOfMonth(addDays(startOfMonth(outputs('thisMonth_StartDate')), 31)), -1)


Reference Functions : https://learn.microsoft.com/en-us/power-automate/minit/date-and-time-operations