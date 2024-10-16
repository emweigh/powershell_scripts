# Add-PDFBookmarks
# 1 -- Get PDF
# 2 -- Add bookmark
# 3 -- Output bookmarked PDF
# Uses pdftk
# Added features to check if bookmarks already exist
# And if so: 1) increment level of all current bookmarks, and 2) insert default bookmark at top of bookmark list
param(
[Parameter(Mandatory=$true, ValueFromPipeline=$true)]
[Alias("p")]
[string]$PDF
)

$main = {
	if($PDF -like $("")){
		throw "PDF param ('-p') required."
	}
	else{
		$filename = "$([System.IO.Path]::GetFileNameWithoutExtension("$PDF"))"
		$defaultBookmark=@("BookmarkBegin","BookmarkTitle: $filename","BookmarkLevel: 1","BookmarkPageNumber: 1")
		$metadata = pdftk $PDF dump_data_utf8
		$newdata = New-Object System.Collections.Generic.List[System.Object]
		$tempfile = New-TemporaryFile
	}
	
	if (Test-ForPDFBookmarks $metadata){
		$newdata=Add-NewPDFBookmark $(Move-PDFBookmarksLevel $metadata) $defaultBookmark
		Write-Output $newdata | Out-File $tempfile -Encoding UTF8
		pdftk $PDF update_info_utf8 $tempfile output "$($filename)_bookmarked.pdf"; Remove-Item $tempfile
	}
	else{
		Write-Output $metadata $defaultBookmark | Out-File $tempfile -Encoding UTF8
		pdftk $PDF update_info_utf8 $tempfile output "$($filename)_bookmarked.pdf"; Remove-Item $tempfile
		}
}

function Test-ForPDFBookmarks ($eval_metadata) {
	foreach($element in $eval_metadata){
		if($element -like "BookmarkBegin"){
			return $true
			break
		}
}return $false
}

function Move-PDFBookmarksLevel ([string[]]$eval_metadata) {
	$returnArray = New-Object System.Collections.Generic.List[System.Object]
	foreach($element in $eval_metadata){
		if ($element -match "BookmarkLevel: (?<lvl>\d+)"){
			$returnArray.Add("BookmarkLevel: {0}" -f $([int]$matches.lvl+1))}
		else{
			$returnArray.Add($element)
		}
}
	return $returnArray.ToArray()
}

function Add-NewPDFBookmark ([string[]]$eval_metadata,[string[]]$bookmark) {
	$count = $eval_metadata.Length - 1
	$top = New-Object System.Collections.Generic.List[System.Object]
	$bottom = New-Object System.Collections.Generic.List[System.Object]
	$final = New-Object System.Collections.Generic.List[System.Object]
	for ($i=0; $i -le $count; $i++){
		if($eval_metadata[$i] -like "BookmarkBegin"){
			$top += $eval_metadata[0..$i]
			$bottom += $eval_metadata[$i..$count]
			$final = $top + $bookmark + $bottom
			return $final
		}
	}
}		
	
& $main | Out-Null
