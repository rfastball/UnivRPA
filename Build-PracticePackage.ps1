# Run after vbaloop push. Ship the HTML next to the workbook in the ZIP.
$ErrorActionPreference = 'Stop'
Push-Location -LiteralPath $PSScriptRoot
try {
    $dest = New-Item -Force -ItemType Directory -Path 'dist/islands'
    $book = Join-Path $dest.FullName 'UnivRPA.xlsm'
    vbaloop dist --out $book --json
    if ($LASTEXITCODE -ne 0) { throw 'Practice workbook build failed.' }
    $room = Get-Item -LiteralPath '군도.html'
    Copy-Item -LiteralPath $room.FullName -Destination $dest.FullName -Force
    $zip = Join-Path $dest.Parent.FullName 'UnivRPA-islands.zip'
    Compress-Archive -LiteralPath $book,(Join-Path $dest.FullName $room.Name),'LICENSE' -DestinationPath $zip -Force
    Write-Output $zip
}
finally { Pop-Location }
