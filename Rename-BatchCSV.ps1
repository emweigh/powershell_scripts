# Rename-BatchCSV
# Renames set of documents using  CSV file to hold new filenames
# Ex: Rename-BatchCSV -c .\export.csv -f Files -i ('#', 'Document')

param(
[Parameter(Mandatory=$true, Position = 0, ValueFromPipeline=$true)]
[Alias("c")]
[string]$inputCSV,

[Parameter(Mandatory=$true, Position = 1)]
[Alias("f")]
[string]$fileColumn,

[Parameter(Mandatory=$true, Position = 2, ValueFromRemainingArguments=$true)]
[Alias("i")]
[string[]]$fields
)

$main = {
    if ($inputCSV -like $("")){
        throw "CSV param ('-c') required."
    }
    if ($fileColumn -like $("")){
        throw "fileColumn param ('-f') required."
    }
    if ($fields.count -le 0){
        throw "fileColumn param ('-f') required."
    }

    $csv = import-csv $inputCSV
    $originalFiles = ($csv | Select-Object -ExpandProperty $fileColumn)
    $renamedFolderName = "$(Get-Date -Format "yyyyMMdd_Hmm")_[RENAMED]_[$(Split-Path -Path (Get-Location) -Leaf)]"
    New-Item -ItemType "directory" -Path ".\$renamedFolderName" -ErrorAction:SilentlyContinue

    foreach($item in $originalFiles){
        if ($item -like $("")){}
        else {
            $pullInfo = $csv | Where-Object -property $fileColumn -eq $item
            $newFilename = ""
            $docExt = [System.IO.Path]::GetExtension($item)

            foreach($field in $fields){
                $addition = $pullInfo | Select-Object -ExpandProperty $field | Out-String -NoNewline
                $newFilename += " $($addition.Split([IO.Path]::GetInvalidFileNameChars()) -join '_')"
            }
            
            Copy-Item ".\$item" ".\$renamedFolderName\$($newFilename.Trim())$docExt" -WarningAction:SilentlyContinue
        }
    }
}
& $main