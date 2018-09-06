$ReportServerUri  = "http://ssrsserver/reportserver/ReportService2010.asmx"
$proxy = New-WebServiceProxy -Uri $ReportServerUri -UseDefaultCredential -Namespace SSRS
$subscriptions = @(
    "29482546-f5c2-481e-8bef-14b83eff8028", #Report Subscription id
    "6037ad02-11ad-420b-acf9-fd2cc1ce39e1", #Report Subscription id
    "51f2701d-2104-4077-bb72-d21ccc3ddf0e", #Report Subscription id
    "7294f10e-63c3-406f-ba39-3aa7a21cf3c2", #Report Subscription id
    "de891def-78b3-4c5e-90b8-59ce81a71fc1", #Report Subscription id
    "fb3e8e6a-affd-4f18-b37d-105e83253ec6", #Report Subscription id
    "c1e169bc-990e-418b-aefd-d88f3bdda1a6"  #Report Subscription id
)
if([int]$(Get-Date -Format "yyyyMMdd") -gt 20171019){
    foreach ($subscription  in $subscriptions){
        $proxy.FireEvent("TimedSubscription",$subscription,$null) 
        "$subscription fired"
    }
} else {
    "..."
}
