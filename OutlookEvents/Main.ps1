Add-Type -Path @(
    "$PSScriptRoot\ApplicationEvent.cs",
    "$PSScriptRoot\ItemEvent.cs",
    "$PSScriptRoot\AppointmentItemEvent.cs",
    "$PSScriptRoot\ContactItemEvent.cs",
    "$PSScriptRoot\DistListItemEvent.cs",
    "$PSScriptRoot\DocumentItemEvent.cs",
    "$PSScriptRoot\JournalItemEvent.cs",
    "$PSScriptRoot\MailItemEvent.cs",
    "$PSScriptRoot\MeetingItemEvent.cs",
    "$PSScriptRoot\MobileItemEvent.cs"
    
) -ReferencedAssemblies @(
    'System',
    'System.Runtime',
    'Microsoft.CSharp',
    'System.Runtime.InteropServices',
    'Microsoft.Office.Interop.Outlook',
    'System.Management.Automation',
    'office'
)

Add-Type -AssemblyName 'Microsoft.Office.Interop.Outlook'

function New-OutlookApplicationEvent
{
    param(
        [Microsoft.Office.Interop.Outlook.Application] $application
    )

    return [PowershellExtensions.OutlookEvents.ApplicationEvent]::new($application)
}

function New-OutlookMailItemEvent
{
    param(
        [System.__ComObject] $mailItem
    )

    return [PowershellExtensions.OutlookEvents.MailItemEvent]::new($mailItem);
}

#<#
$oApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
$mItem = $oApp.CreateItem([Microsoft.Office.Interop.Outlook.OlItemType]::olMailItem)
$sendEvent = {
    Write-Host $this
    Write-Host "Sending From App"
}
$itemsendEvent = {
    Write-Host $this
    Write-Host "Sending From Item"
}

$oEvent = New-OutlookApplicationEvent $oApp
$oItemEvent = New-OutlookMailItemEvent $mItem
$oEvent.Add_ItemSend($sendEvent)
$oItemEvent.Add_Send($itemsendEvent)

$mItem.Display($true)

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($oApp)
#>