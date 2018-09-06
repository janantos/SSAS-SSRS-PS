[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices.AdomdClient") | Out-Null;
function Get-CubeValue
(
    [Parameter(Position=0, ParameterSetName="array")]
    [object[]]$Parameters
)
{
    $CubeConnStr = $Parameters[0]
    $Tuple = $Parameters[1..$($Parameters.Length-1)] -join ","
    $cube = $($connectionString -split ";"  | Where-Object { [regex]::matches($_.ToString(),"cube=","IgnoreCase") } | ForEach-Object {$_ -split "="})[1].Trim()

    $conn = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection; 
    $conn.ConnectionString = $connectionString
    $conn.Open(); 
    $cmd = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdCommand; 
    $cmd.Connection = $conn; 
    $da = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdDataAdapter; 
    $cmd.CommandText = "Select ($Tuple) on 0 from [$cube]"
    $da.SelectCommand = $cmd; 
    $data = New-Object System.Data.DataTable; 
    $nil = $da.Fill($data); 
    return $data.Rows[0].Item(0)
}



$connectionString = "DataSource=SSASServer; Catalog=Database; Provider=MSOLAP; Cube=Cube"
Get-CubeValue $connectionString, "[Measures].[Turnover]","[Brand].[Brand].[My Brand]"
