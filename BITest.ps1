Param(
    [string]$TestScenarioPath
)


$TestsObj =  Get-Content -Raw -Path $TestScenarioPath | ConvertFrom-Json

# Load "Microsoft.AnalysisServices.AdomdClient" needed for Get-CubeValue function
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices.AdomdClient") | Out-Null;


function Get-CubeValue
(
    [string]$CubeConnStr,
    [string]$Query
)
{

    $conn = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection; 
    $conn.ConnectionString = $CubeConnStr
    $conn.Open(); 
    $cmd = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdCommand; 
    $cmd.Connection = $conn; 
    $da = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdDataAdapter; 
    $cmd.CommandText = $Query
    $da.SelectCommand = $cmd; 
    $data = New-Object System.Data.DataTable; 
    $null = $da.Fill($data); 
    if ($data.Rows.Count -eq 1){
        $return = $data.Rows[0].Item(0)
    } else {
        $return = [System.DBNull]
    }
    $da.Dispose()
    $data.Dispose()
    $conn.Close()
    return $return
}

function Get-ODBCValue
(
    [string]$CubeConnStr,
    [string]$Query
)
{
    $conn = New-Object System.Data.Odbc.OdbcConnection
    $conn.ConnectionString = $CubeConnStr  #"DSN=odbcdsn);"
    $conn.open()
    $cmd = New-object System.Data.Odbc.OdbcCommand($Query,$conn)
    $ds = New-Object system.Data.DataSet
    (New-Object system.Data.odbc.odbcDataAdapter($cmd)).fill($ds) | out-null
    $conn.close()
    return $ds.Tables[0].Rows[0].Item(0)
}

$current  = 0
$failed = 0
$softfailed = 0
$passed = 0
$results = ""
$TestCount = $TestsObj.Tests.Count

$resumeResult = "
<span style='font-weight: bold;'>Total Tests: {0}</span> 
<span style='color:green'>Passed: {1}</span> 
<span style='color:red'>Hard Failed: {2}</span> 
<span style='color:orange'>(Soft Passed: {4})</span>
<span style='font-weight: bold;'>Hard Pass rate: {3}</span>
<span style='font-weight: bold;'>Soft Pass rate: {5}</span>
<button id=""btnExport"" onclick=""fnExcelReport();""> EXPORT Excel </button><hr>
"
$resultstable = "
<table id='resulttable'>
<thead>
<tr>
<th class='width400'>Cube Slice<br>Filter: <input type='text' onkeyup='ResultFilter()' id='filtertxt'></th>
<th class='width100'>Side A<br>{1}</th>
<th class='width100'>Side B<br>{2}</th>
<th class='width100'>Diff</th>
<th class='width100'>% SideA Diff</th>
<th class='width100'>Result Hard Fail<br><input type='checkbox' onchange='ResultFilter()' id='failedcb' checked/>Failed only</th>
<th class='width100'>Result Soft Fail<br><input type='checkbox' onchange='ResultFilter()' id='softfailedcb'/>Failed only</th>
</tr>
</thead>
<tbody>
{0}
</tbody>
</table>
"

$output = "<!DOCTYPE html>
<html>
<head>
<title>{0}</title>
<style>
table
{{
	width: 100%;
	border-collapse: collapse;
}}

thead
{{
	display: block;
	width: 100%;
	overflow: auto;
	color: #fff;
	background: #000;
}}

tbody
{{
	display: block;
	width: 100%;
	height: 600px;
	overflow: auto;
}}

th,td
{{
	padding: .5em 1em;
	text-align: left;
	vertical-align: top;
    border-left: 1px solid #fff;
}}

tr
{{
	border-bottom: 1px solid #fff;
}}

.width100 {{ width: 100px; }}
.width100r {{ width: 100px; background: salmon;}}
.width100g {{ width: 100px; background: lightgreen;}}
.width400 {{ width: 400px; }}
.width400r {{ width: 400px; background: salmon;}}
.width400g {{ width: 400px; background: lightgreen;}}
.width100y {{ width: 100px; background: orange;}}
.width400y {{ width: 400px; background: orange;}}
</style>

<script>
function ResultFilter() {{
    
    table = document.getElementById('resulttable');
    tr = table.getElementsByTagName('tr');
    failedCb = document.getElementById('failedcb');
    softfailedCb = document.getElementById('softfailedcb');
    filtertxt = document.getElementById('filtertxt');
    
    for(i = 0; i < tr.length; i++){{
        tr[i].style.display = '';
    }}

    if (!(filtertxt.value == '')) {{
        for(i = 0; i < tr.length; i++) {{
            td = tr[i].getElementsByTagName('td')[0];
            if(td) {{
                if(td.innerHTML.toUpperCase().indexOf(filtertxt.value.toUpperCase()) > -1) {{
                    
                }}
                else {{
                    tr[i].style.display = 'none';
                }}
            }}
        }}
    }} 
        
    if (failedCb.checked){{
        for(i = 0; i < tr.length; i++) {{
            td5 = tr[i].getElementsByTagName('td')[5];
            if(td5) {{
                if((td5.innerHTML.indexOf('Failed') > -1)) {{
                    
                }} else {{
                    tr[i].style.display = 'none';
                }}
            }}
        }}
    }}

    if (softfailedCb.checked){{
        for(i = 0; i < tr.length; i++) {{
            td5 = tr[i].getElementsByTagName('td')[6];
            if(td5) {{
                if((td5.innerHTML.indexOf('Failed') > -1)) {{
                    
                }} else {{
                    tr[i].style.display = 'none';
                }}
            }}
        }}
    }}

}}


function fnExcelReport()
{{
    var tab_text=""<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Comparison</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table border='2px'><tr bgcolor='#87AFC6'>"";
    var textRange; var j=0;
    tab = document.getElementById('resulttable'); // id of table

    for(j = 0 ; j < tab.rows.length ; j++) 
    {{     
        tab_text=tab_text+tab.rows[j].innerHTML+""</tr>"";
        //tab_text=tab_text+""</tr>"";
    }}

    tab_text=tab_text+""</table></body></html>"";
    //tab_text= tab_text.replace(/<A[^>]*>|<\/A>/g, """");//remove if u want links in your table
    //tab_text= tab_text.replace(/<img[^>]*>/gi,""""); // remove if u want images in your table
    tab_text= tab_text.replace(/<input[^>]*>|<\/input>/gi, """"); // reomves input params
    sa = window.open('data:application/vnd.ms-excel,' + encodeURIComponent(tab_text));  
    return (sa);
}}

</script>

</head>
<body  onload='ResultFilter()'>
<h3>{0}</h3>
{1}

{2}

<h4>Test Scenario Info:</h4>
Execution Time: {3}
<br>
SideA Parameters: {4}
<br>
SideB Parameters: {5}
<br>
Test Scenario File: {6}
<br>
SoftFail Treshold: {7}

<iframe id=""txtArea1"" style=""display:none""></iframe>


</body>
</html>
"


foreach ($test in $TestsObj.Tests) {
    $current = $current + 1
    $SideA_Query = $test.SideA_Query -f $TestsObj.SideA_Parameters
    $SideB_Query = $test.SideB_Query -f $TestsObj.SideB_Parameters

    Write-Progress -Activity "Running tests" -status "Progress: ($current/$TestCount)" -PercentComplete ($current/$TestCount*100) -CurrentOperation "$test"
    Write-Host $SideA_Query
    try {
        if ($test.SideA_Type -eq "SSAS") {
            $SideA_RetValue = Get-CubeValue -CubeConnStr $test.SideA_ConnStr -Query $SideA_Query
        } elseif ($test.SideA_Type -eq "ODBC"){
            $SideA_RetValue = Get-ODBCValue -CubeConnStr $test.SideA_ConnStr -Query $SideA_Query
        }

        if ($test.SideB_Type -eq "SSAS") {
            $SideB_RetValue = Get-CubeValue -CubeConnStr $test.SideB_ConnStr -Query $SideB_Query
        } elseif ($test.SideB_Type -eq "ODBC"){
            $SideB_RetValue = Get-ODBCValue -CubeConnStr $test.SideB_ConnStr -Query $SideB_Query
        }
    } catch {
        $SideA_RetValue = [System.DBNull]
        $SideB_RetValue = [System.DBNull]
        Write-Host "test broken"
    }
    if ($SideA_RetValue.GetType() -eq [System.DBNull]) {
        $SideA_RetValue2 = 0
    } else {
        $SideA_RetValue2 = $SideA_RetValue
    }
    
    if ($SideB_RetValue.GetType() -eq [System.DBNull]) {
        $SideB_RetValue2 = 0
    } else {
        $SideB_RetValue2 = $SideB_RetValue
    }
    
    $SideA_RetValueHtml = [decimal]$("{0:N2}" -f $SideA_RetValue2)
    $SideB_RetValueHtml = [decimal]$("{0:N2}" -f $SideB_RetValue2)

    $diff = [decimal]("{0:N2}" -f $($SideB_RetValue2 - $SideA_RetValue2))
    if ($SideA_RetValue2 -ne 0 ) {
        $diff_perc = "{0:N2}%" -f  $((($SideB_RetValue2 - $SideA_RetValue2)/$SideA_RetValue2)*100)
    } else {
        $diff_perc = $null
    }
    $HtmlTestDesc = "{4}) {5}<br>SourceA: {0}<br>SourceB: {1}<br>QueryA: {2}<br>QueryB: {3}" -f $test.SideA_ConnStr, $test.SideB_ConnStr, $($SideA_Query -replace "&","&amp;"), $($SideB_Query -replace "&","&amp;"), $current, $test.TestName
    
    if ($diff -ne 0) {
        $failed = $failed + 1
        if ($SideA_RetValue2 -eq 0  ){
            $softfailstatus = "Failed"
         } else {
            if ([decimal]$([Math]::abs((($SideB_RetValue2 - $SideA_RetValue2)/$SideA_RetValue2)*100)) -gt $test.SoftFailTreshold ) {
                $softfailstatus = "Failed"
            } else {
                $softfailstatus = "Passed"
            }  

         }

        if ($softfailstatus -eq "Passed"){
            $softfailed = $softfailed + 1
            $results = $results +  "<tr'><td class='width400y'>$HtmlTestDesc</td><td class='width100y'>$SideA_RetValueHtml</td><td class='width100y'>$SideB_RetValueHtml</td><td class='width100y'>$diff</td><td class='width100y'>$diff_perc<br>Treshold:$($test.SoftFailTreshold)</td><td class='width100y'>Failed</td><td class='width100y'>$softfailstatus</td></tr>`n"
        } else {
            $results = $results +  "<tr'><td class='width400r'>$HtmlTestDesc</td><td class='width100r'>$SideA_RetValueHtml</td><td class='width100r'>$SideB_RetValueHtml</td><td class='width100r'>$diff</td><td class='width100r'>$diff_perc<br>Treshold:$($test.SoftFailTreshold)</td><td class='width100r'>Failed</td><td class='width100r'>$softfailstatus</td></tr>`n"
        }

    } else {
        $passed = $passed + 1
        $results = $results +  "<tr><td class='width400g'>$HtmlTestDesc</td><td class='width100g'>$SideA_RetValueHtml</td><td class='width100g'>$SideB_RetValueHtml</td><td class='width100g'>$diff</td><td class='width100g'>$diff_perc<br>Treshold:$($test.SoftFailTreshold)</td><td class='width100g'>Passed</td><td class='width100g'>Passed</td></tr>`n"
    }    
}

Write-Progress -Completed -Activity "Running tests"
# Result Variables
$passrate = "{0:N2}%" -f $(($passed/$TestCount)*100)
$softpassrate = "{0:N2}%" -f ((($passed+$softfailed)/$TestCount)*100)
$resumeResult = $resumeResult -f $TestCount, $passed, $failed, $passrate, $softfailed, $softpassrate
$resultstable = $resultstable -f $results, $TestsObj.SideA_Name, $TestsObj.SideB_Name
#Generate Output
$output  -f $TestsObj.TestSetName, $resumeResult, $resultstable , $(Get-Date), $($TestsObj.SideA_Parameters -join ";"), $($TestsObj.SideB_Parameters -join ";"), $TestScenarioPath, $($TestsObj.TestSetSoftFailTreshold)  | Out-File "Result_$($TestsObj.TestSetName)_$(Get-Date -format yyyyMMddhhmmss).html"
Start-Process "Result_$($TestsObj.TestSetName)_$(Get-Date -format yyyyMMddhhmmss).html"


