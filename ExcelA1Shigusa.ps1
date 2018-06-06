Param([string] $originalPath)

$random = Get-Random
$dirPath = $originalPath + "_" + $random
$zipPath = $dirPath + ".zip"

Move-Item $originalPath $zipPath

Expand-Archive $zipPath $dirPath

Get-ChildItem $dirPath\xl\worksheets\*.xml | % {
  [xml]$xml = Get-Content $_ -Encoding UTF8
  $activeCell = $xml.GetElementsByTagName("worksheet")[0].GetElementsByTagName("selection")[0]
  if ($activeCell){
    # $activeCell.ParentNode.RemoveChild($activeCell)
    $activeCell.activeCell = "A1"
    $activeCell.sqref = "A1"
    $xml.Save($_)
  }
}

Compress-Archive -Force $dirPath"\*" $zipPath

Move-Item $zipPath $originalPath

Remove-Item -Recurse $dirPath