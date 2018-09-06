function loadGACAssemblyLastVer($partialname){
    $(Get-ChildItem C:\Windows\assembly\GAC_MSIL -filter $($partialname + ".dll") -recurse ) | 
    Sort-Object -Descending | 
    Select-Object -First 1 |
    ForEach-Object {
        [System.Reflection.Assembly]::Load($($($partialname +  
         ", Version=" + $_.Directory.Name.Substring(0,$($_.Directory.Name.Length-18)) + 
         ", Culture=neutral, PublicKeyToken=" + $_.Directory.Name.Substring($($_.Directory.Name.Length-16),16))))
    }
}

#Add-Type -AssemblyName Microsoft.AnalysisServices

$serverName="ssasServer"
$xmlaPathDir = "c:\CreateScripts"


[System.Reflection.Assembly]::Load("Microsoft.AnalysisServices, Version=12.0.0.0, Culture=neutral, PublicKeyToken=89845dcd8080cc91")

#loadGACAssemblyLastVer("Microsoft.AnalysisServices")

$server = New-Object Microsoft.AnalysisServices.Server
$server.connect($ServerName)
if ($server.name -eq $null) {
 Write-Output ("Server '{0}' not found" -f $serverName)
 break
}


$server.Databases | ForEach-Object {

$name = $_.Name

$db = $server.Databases.FindByName($name)

Write-Host "Scripting database $($name)"
   
   $stringBuilder = New-Object System.Text.StringBuilder
   $stringWriter = New-Object System.IO.StringWriter($stringBuilder)
   $xmlOut = New-Object System.Xml.XmlTextWriter($stringWriter)
   $xmlOut.Formatting = [System.Xml.Formatting]::Indented
   $scriptObject = New-Object Microsoft.AnalysisServices.Scripter
   $MSASObject=[Microsoft.AnalysisServices.MajorObject[]] @($db)
   $ScriptObject.ScriptCreate($MSASObject,$xmlOut,$false)
   $stringbuilder.ToString() |out-file -filepath "$xmlaPathDir\$(get-date -Format yyyyMMdd)_$($name).xmla"
}

