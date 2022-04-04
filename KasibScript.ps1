# Locates the previous month pqc and moves into new folder on USB
# Zips those files then deletes originals
# TODO: move AS files to AS folder, move LYSE and LYSENOWASH files, move CHARACTERIZATION (need its abbrev)
# TODO: 

$source = "C:\BD Export\Setup\Reports"
$file = Get-ChildItem $source -Filter *pdf -Recurse

$prevmonth = (Get-Date).AddMonths(-1).ToString("MM")
$year = (Get-Date).ToString("yyyy")
$foldermonth = (Get-Date).ToString("MMM").ToUpper()

New-Item -Path "C:\Users\kasib\Documents\Work\Computer Science\Scripting\PQC $foldermonth" -ItemType Directory

$destination = "C:\Users\kasib\Documents\Work\Computer Science\Scripting\PQC $foldermonth"
$zip = "C:\Users\kasib\Documents\Work\Computer Science\Scripting\PQC $foldermonth\PQC $foldermonth.zip"

foreach ($i in $file) {
    if ($i.Name -match $prevmonth -and $i.Name -match $year) {
        Move-Item $i $destination
    }
}

Compress-Archive $destination\* $zip

#once the item is selected just pipe into your action
Get-ChildItem $destination -Filter *.pdf -Recurse | Remove-Item -Force

for ($counter = 1; $counter -le 10; $counter++) {
    Write-Progress -Activity "Update Progress" -Status "$counter% Complete:" -PercentComplete $counter;
}


Write-Host "`n`n`n`n`n`nCompleted"

for ($seconds=5; $seconds -gt -1; $seconds--) {
    
    Write-Host -NoNewline ("`rseconds remaining: " + ("{0:d4}" -f $seconds))
    Start-Sleep -Seconds 1
}

# FLOW reports @ C:\BD Export\Setup\Reports
# USB @ D:\
# PQC destination D:\2022\PM PR 487131 MAR 2022 contains zip of PQC FEB 2022.zip, LyseNoWash03072022.pdf and lyse wash
# AS FILE destination C:\BD Export\Setup\Reports\AS Reports
#
#