# Locates the previous month pqc and moves into new folder on USB
# Zips those files then deletes originals
# TODO: move AS files to AS folder, move LYSE and LYSENOWASH files, move CHARACTERIZATION (need its abbrev)

$source = "C:\BD Export\Setup\Reports"
$file = Get-ChildItem $source -Filter *pdf
$ASdestination = "C:\BD Export\Setup\Reports\AS Reports"

$prevmonth = (Get-Date).AddMonths(-1).ToString("MM")
$year = (Get-Date).ToString("yyyy")
$foldermonth = (Get-Date).ToString("MMM").ToUpper()
$prevfoldermonth = (Get-Date).ToString("MMM").ToUpper()

# create folder on USB
New-Item -Path "D:\$year\PM PR 487131 $foldermonth $year" -ItemType Directory

$PQCdestination = "D:\$year\PM PR 487131 $foldermonth $year"
$zip = "D:\$year\PM PR 487131 $foldermonth $year\PQC $prevfoldermonth $year.zip"

# Move each PQC file in FLOW to USB, move each AS file to AS folder
foreach ($i in $file) {
    if ($i.Name -match $prevmonth -and $i.Name -match "PQC") {
        Move-Item $i.FullName $PQCdestination
    }
    if ($i.Name -match "AS" -and $i.Name -match $year) {
        Move-Item $i.FullName $ASdestination
    }
}

# zip the moved files in USB
Add-Type -Assembly 'System.IO.Compression.Filesystem'
[System.IO.Compression.ZipFile]::CreateFromDirectory($PQCdestination, $zip)

foreach ($i in $file) {
    if ($i.Name -match "Lyse" -or $i.Name -match "LYSE" -or $i.Name -match "lyse") {
        Move-Item $i.FullName $PQCdestination
    }
}

# Not using the delete function right now
# Get-ChildItem $PQCdestination -Filter *.pdf | Remove-Item -Force

Write-Host `n`n`n`n
for ($counter = 1; $counter -le 100; $counter++) {
    Write-Progress -Activity "Update Progress" -Status "$counter% Complete:" -PercentComplete $counter;
}


# FLOW reports @ C:\BD Export\Setup\Reports
# USB @ D:\
# PQC destination D:\2022\PM PR 487131 MAR 2022 contains zip of PQC FEB 2022.zip, LyseNoWash03072022.pdf and lyse wash
# AS FILE destination C:\BD Export\Setup\Reports\AS Reports
#
# https://theposhwolf.com/howtos/PowerShell-and-Zip-Files/