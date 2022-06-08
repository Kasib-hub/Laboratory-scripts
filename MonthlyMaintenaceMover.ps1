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
$singleDigitMonth = (Get-Date).Month

$folderMonth = (Get-Date).ToString("MMM").ToUpper()
$prevFolderMonth = (Get-Date).AddMonths(-1).ToString("MMM").ToUpper()

$year = (Get-Date).ToString("yyyy")
$prevYear = (Get-Date).AddYears(-1).Year

$lastDayOfPrevMonth = [datetime]::DaysInMonth($year, $prevSingleDigitMonth)

#Code I am working on 20 APR 2022 KZA
if ((Get-Date).Month -lt (Get-Date).AddMonths(-1).Month) {
    # Window for files within the month
    $start = Get-Date -Year (Get-Date).AddYears(-1).Year -Month $prevSingleDigitMonth -Day 1 -Hour 0 -Minute 0 -Second 0
    $end = Get-Date -Year (Get-Date).AddYears(-1).Year -Month $prevSingleDigitMonth -Day $lastDayOfPrevMonth -Hour 23 -Minute 59 -Second 59
    
    $pqcToBeMoved = Get-ChildItem $flowReports -Filter *.pdf |  
        Where-Object {$_.LastWriteTime -gt $start -and $_.LastWriteTime -lt $end }
    
    New-Item -Path "D:\$year\PM PR 487131 $folderMonth $year" -ItemType Directory
    $pqcDestination = "D:\$year\PM PR 487131 $folderMonth $year"
    
    foreach ($i in $pqcToBeMoved) {
        if ($i.Name -match "PQC") {
            Move-Item $i.FullName $pqcDestination
        }
    }
    
    zipexe a -mx=9 "D:\$year\PM PR 487131 $folderMonth $year\PQC $prevFolderMonth $prevYear.zip" "D:\$year\PM PR 487131 $folderMonth $year\*" -sdel
}
else {
    # Window for files within the month
    $start = Get-Date -Year $year -Month $prevSingleDigitMonth -Day 1 -Hour 0 -Minute 0 -Second 0
    $end = Get-Date -Year $year -Month $prevSingleDigitMonth -Day $lastDayOfPrevMonth -Hour 23 -Minute 59 -Second 59
    
    # create folder on USB
    New-Item -Path "D:\$year\PM PR 487131 $folderMonth $year" -ItemType Directory
    
    # Filtered Files PQC for current year
    $pqcToBeMoved = Get-ChildItem $flowReports -Filter *.pdf |  
        Where-Object {$_.LastWriteTime -gt $start -and $_.LastWriteTime -lt $end }
    
    $pqcDestination = "D:\$year\PM PR 487131 $folderMonth $year"
    $restToBeMoved = Get-ChildItem $flowReports -Filter *pdf
    
    # Move each PQC
    foreach ($i in $pqcToBeMoved) {
        if ($i.Name -match "PQC") {
            Move-Item $i.FullName $pqcDestination
        }
    }
    
    # Zip the moved PQC in place, use actual filepaths!
    zipexe a -mx=9 "D:\$year\PM PR 487131 $folderMonth $year\PQC $prevFolderMonth $year.zip" "D:\$year\PM PR 487131 $folderMonth $year\*" -sdel
}




# Move each AS and Lyse file TODO - move the CQC files
foreach ($i in $restToBeMoved) {
    if ($i.Name -match "LYSE" -or $i.Name -match "Lyse" -or $i.Name -match "lyse" -or $i.Name -match "WASH" -or $i.Name -match "Wash" -or $i.Name -match "wash") {
        Move-Item $i.FullName $pqcdestination
    }
    if ($i.Name -match "AS Report") {
        Move-Item $i.FullName $asdestination
    }
    if ($i.Name -match "CQC") {
        Move-Item $i.FullName $pqcdestination
    }
}
