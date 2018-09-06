# Kill all PowerBI "rude" queries against MOLAP SSAS - grand total issue as described by Chris Webb 
#####                    SET VARIABLES                #####
$SSASServerName = "ssasserver"
###########################################################

[Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices") | Out-Null;
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices.AdomdClient") | Out-Null; 


$SSASServer = New-Object Microsoft.AnalysisServices.Server
$SSASServer.Connect($SSASServerName)
$MySessionID = $SSASServer.SessionID
$MyConnID = $SSASServer.ConnectionInfo
Write-Host $MyConnID

$conn = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection; 
$conn.ConnectionString = "Data Source=$SSASServerName;SspropInitAppName=PowerShell SsasDiscoverCurrentProcesses;" 
$conn.Open(); 
$cmd = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdCommand; 
$cmd.Connection = $conn; 

$da = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdDataAdapter; 

[String] $mdx = "SELECT
[SESSION_CONNECTION_ID]
, [SESSION_USER_NAME]
, [SESSION_ID]
, [SESSION_LAST_COMMAND]
, [SESSION_ELAPSED_TIME_MS]
from
`$SYSTEM.DISCOVER_SESSIONS
where SESSION_ID <> '$MySessionID'
and SESSION_STATUS = 1
"

$cmd.CommandText = $mdx; 
$da.SelectCommand = $cmd; 

$victims = 0
$spreecount =0

while(1)
{
    Clear-Host
    $spreecount = $spreecount + 1
    Write-Host "  " -ForegroundColor Red
    Write-Host "  $(Get-Date)"
    write-host "  Killing spree #$spreecount"
    $connTbl = New-Object System.Data.DataTable; 
    $nil = $da.Fill($connTbl); 

    foreach ($connRow in $connTbl.Rows) {
        #$SSASServer.CancelSession($connRow.SESSION_ID)
        if ( $connRow.SESSION_LAST_COMMAND -like "EVALUATE*ROW(*") {
            if ($connRow.SESSION_ELAPSED_TIME_MS -gt 90000) {
                Write-Host -NoNewline "  Cancelling session ..."
                $SSASServer.CancelConnection($connRow.SESSION_CONNECTION_ID)
                $a = $connRow.SESSION_ID +";"+ $connRow.SESSION_USER_NAME+";"+$(Get-Date)
                Write-Host $a
                Out-File -FilePath kill.log -InputObject $a -Append -Encoding ascii 
                $victims = $victims + 1
            }
        }
    }
    Write-Host "  Victims Total: $victims"
    Start-Sleep -Seconds 10
}
