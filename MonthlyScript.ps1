# Locates the previous month pqc and moves into new folder on USB
Write-Host "Looking for files..."
Start-Sleep -Seconds 1

# Filepaths
$flowReports = "C:\BD Export\Setup\Reports"
$asDestination = "C:\BD Export\Setup\Reports\AS Reports"
$7zip = "C:\Program Files\7-Zip\7z.exe"
Set-Alias zipexe $7zip

# Dates

$prevMonth = (Get-Date).AddMonths(-1).ToString("MM")
$prevSingleDigitMonth = (Get-Date).AddMonths(-1).Month
$year = (Get-Date).ToString("yyyy")
#$prevYear = (Get-Date).AddYears(-1).Year
$folderMonth = (Get-Date).ToString("MMM").ToUpper()
$prevFolderMonth = (Get-Date).AddMonths(-1).ToString("MMM").ToUpper()
$lastDayOfPrevMonth = [datetime]::DaysInMonth($year, $prevSingleDigitMonth)
#$month = (Get-Date).Month

# Window for files within the month
$start = Get-Date -Year $year -Month $prevSingleDigitMonth -Day 1 -Hour 0 -Minute 0 -Second 0
$end = Get-Date -Year $year -Month $prevSingleDigitMonth -Day $lastDayOfPrevMonth -Hour 23 -Minute 59 -Second 59

# create folder on USB
New-Item -Path "D:\$year\PM PR 487131 $folderMonth $year" -ItemType Directory

# Filtered Files PQC for current year
$pqcToBeMoved = Get-ChildItem $flowReports -Filter *.pdf |  
    Where-Object {$_.LastWriteTime -gt $start -and $_.LastWriteTime -lt $end }

$pqcDestination = "D:\$year\PM PR 487131 $folderMonth $year"
# TODO, add edge case for January where previous month is last year


$restToBeMoved = Get-ChildItem $flowReports -Filter *pdf

# Move each PQC
foreach ($i in $pqcToBeMoved) {
    if ($i.Name -match "PQC") {
        Move-Item $i.FullName $pqcDestination
    }
}

# Zip the moved PQC in place, use actual filepaths!
$7zip = "C:\Program Files\7-Zip\7z.exe"
Set-Alias zipexe $7zip
zipexe a -mx=9 "D:\$year\PM PR 487131 $folderMonth $year\PQC $prevFolderMonth $year.zip" "D:\$year\PM PR 487131 $folderMonth $year\*" -sdel

# Move each AS and Lyse file TODO - move the CQC files
foreach ($i in $restToBeMoved) {
    if ($i.Name -match "LYSE" -or $i.Name -match "Lyse" -or $i.Name -match "lyse" -or $i.Name -match "WASH" -or $i.Name -match "Wash" -or $i.Name -match "wash") {
        Move-Item $i.FullName $pqcdestination
    }
    if ($i.Name -match "AS" -and $i.Name -match $year) {
        Move-Item $i.FullName $asdestination
    }
    if ($i.Name -match "CQC") {
        Move-Item $i.FullName $pqcdestination
    }
}
