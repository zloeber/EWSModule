# Create embeddedable EWS DLL
$OutFile = 'EncodedFile.ps1'
Get-ChildItem -Path ..\ -Filter "*.dll" | Foreach {
    $FileName = $_.Name
    $VarName = '$decode_' + ($FileName -replace '.dll','' -replace '.','')
    $Content = Get-Content -Path $_.FullName -Encoding Byte
    
    $Base64 = [System.Convert]::ToBase64String($Content)
    $NewContent = $VarName +  '= "' + $Base64 + '"'
    $NewContent | Out-File $OutFile -Append
}
