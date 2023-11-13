#------------------------------------------------------
# Name:        BuildSite
# Purpose:     Builds All VMs needed for RayStation Deployment based on Excel Values
# Author:      John Burriss
# Created:     10/25/2023  2:22 PM 
#------------------------------------------------------
#Requires -RunAsAdministrator

param(
  [Parameter(Mandatory=$false)]
  [string]$skipnetworkcheck = $False,
  [Parameter(Mandatory=$false)]
  [string]$skipexcelimport = $False,
  [Parameter(Mandatory=$false)]
  [string]$skipoverprovisioncheck = $False,
  [Parameter(Mandatory=$false)]
  [string]$skipadcheck = $False

)

$RunLocation = split-path -parent $MyInvocation.MyCommand.Definition -ErrorAction SilentlyContinue

$PowerCLI = Get-Module -ListAvailable -name vmware.powercli
$PSExcel = Get-Module -ListAvailable -name psexcel
$PEMEncrypt = Get-Module -ListAvailable -Name PEMEncrypt

if($Null -eq $PowerCLI){
Write-host "Importing PowerCLI Module" -ForegroundColor Green
Copy-Item "$RunLocation\Modules\PowerCLI\" "C:\Windows\System32\WindowsPowerShell\v1.0\Modules" -Recurse

}
if($Null -eq $PSExcel){
Write-host "Importing PSExcel Module" -ForegroundColor Green
Copy-Item "$RunLocation\Modules\PSExcel" "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\PSExcel" -Recurse

}
if($Null -eq $PEMEncrypt){
Write-host "Importing PEMEncrypt Module" -ForegroundColor Green
Copy-Item "$RunLocation\Modules\PEMEncrypt" "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\PEMEncrypt" -Recurse

}
Write-Host "Importing Modules, This may take a few min." -ForegroundColor Green
Import-Module vmware.powercli | out-null
Import-Module PSExcel
Import-Module PEMEncrypt


function WriteJobProgress
{
    param($Job,
          $Completed = $false,
          $Index,
          $currentOp)
 
    #Make sure the first child job exists
    if($Job.percentcomplete -ne $null)
    {
        #Extracts the latest progress of the job and writes the progress
        $name = $job.name
        $State = $Job.state
        $Percent = $job.percentcomplete
    
        #When adding multiple progress bars, a unique ID must be provided. Here I am providing the JobID as this
        if($Completed -eq $false){
        try{
        Write-Progress -Id $Index -Activity "$name $currentOp" -Status "$State $percent`%" -PercentComplete $job.percentcomplete;
        }
        catch{
        $_
        }
        }
        elseif($Completed -eq $true){
        Write-Progress -Id $Index -Activity $job.name -Completed
        }
    }
}

function Test-Cred {
           
    [CmdletBinding()]
    [OutputType([String])] 
       
    Param ( 
        [Parameter( 
            Mandatory = $false, 
            ValueFromPipeLine = $true, 
            ValueFromPipelineByPropertyName = $true
        )] 
        [Alias( 
            'PSCredential'
        )] 
        [ValidateNotNull()] 
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()] 
        $Credentials
    )
    $Domain = $null
    $Root = $null
    $Username = $null
    $Password = $null
      
    If($Credentials -eq $null)
    {
        Try
        {
            $Credentials = Get-Credential "domain\$env:username" -ErrorAction Stop
        }
        Catch
        {
            $ErrorMsg = $_.Exception.Message
            Write-Warning "Failed to validate credentials: $ErrorMsg "
            Pause
            Break
        }
    }
      
    # Checking module
    Try
    {
        # Split username and password
        $Username = $credentials.username
        $Password = $credentials.GetNetworkCredential().password
  
        # Get Domain
        $Root = "LDAP://" + ([ADSI]'').distinguishedName
        $Domain = New-Object System.DirectoryServices.DirectoryEntry($Root,$UserName,$Password)
    }
    Catch
    {
        $_.Exception.Message
        Continue
    }
  
    If(!$domain)
    {
        Write-Warning "Something went wrong"
    }
    Else
    {
        If ($domain.name -ne $null)
        {
            return "Authenticated"
        }
        Else
        {
            return "Not authenticated"
        }
    }
}

function Test-ADUser {
  param(
    [Parameter(Mandatory)]
    [String]
    $sAMAccountName
  )
  $null -ne ([ADSISearcher] "(sAMAccountName=$sAMAccountName)").FindOne()
}


Function Get-GPUProfile {
    Param ($vmhost)
    $VMhost = Get-VMhost $VMhost
    $vmhost.ExtensionData.Config.SharedPassthruGpuTypes
}
  
Function Get-vGPUDevice {
    Param ($vm)
    $VM = Get-VM $VM
    $vGPUDevice = $VM.ExtensionData.Config.hardware.Device | Where { $_.backing.vgpu}
    $vGPUDevice | Select Key, ControllerKey, Unitnumber, @{Name="Device";Expression={$_.DeviceInfo.Label}}, @{Name="Summary";Expression={$_.DeviceInfo.Summary}}
}
  
Function Remove-vGPU {
    Param (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true,Position=0)] $VM,
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true,Position=1)] $vGPUDevice
    )
  
    $ControllerKey = $vGPUDevice.controllerKey
    $key = $vGPUDevice.Key
    $UnitNumber = $vGPUDevice.UnitNumber
    $device = $vGPUDevice.device
    $Summary = $vGPUDevice.Summary
  
    $VM = Get-VM $VM
  
    $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
    $spec.deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec[] (1)
    $spec.deviceChange[0] = New-Object VMware.Vim.VirtualDeviceConfigSpec
    $spec.deviceChange[0].operation = 'remove'
    $spec.deviceChange[0].device = New-Object VMware.Vim.VirtualPCIPassthrough
    $spec.deviceChange[0].device.controllerKey = $controllerkey
    $spec.deviceChange[0].device.unitNumber = $unitnumber
    $spec.deviceChange[0].device.deviceInfo = New-Object VMware.Vim.Description
    $spec.deviceChange[0].device.deviceInfo.summary = $summary
    $spec.deviceChange[0].device.deviceInfo.label = $device
    $spec.deviceChange[0].device.key = $key
    $_this = $VM  | Get-View
    $nulloutput = $_this.ReconfigVM_Task($spec)
}
  
Function New-vGPU {
    Param ($VM, $vGPUProfile)
    $VM = Get-VM $VM
    $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
    $spec.deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec[] (1)
    $spec.deviceChange[0] = New-Object VMware.Vim.VirtualDeviceConfigSpec
    $spec.deviceChange[0].operation = 'add'
    $spec.deviceChange[0].device = New-Object VMware.Vim.VirtualPCIPassthrough
    $spec.deviceChange[0].device.deviceInfo = New-Object VMware.Vim.Description
    $spec.deviceChange[0].device.deviceInfo.summary = ''
    $spec.deviceChange[0].device.deviceInfo.label = 'New PCI device'
    $spec.deviceChange[0].device.backing = New-Object VMware.Vim.VirtualPCIPassthroughVmiopBackingInfo
    $spec.deviceChange[0].device.backing.vgpu = "$vGPUProfile"
    $vmobj = $VM | Get-View
    $reconfig = $vmobj.ReconfigVM_Task($spec)
    if ($reconfig) {
        $ChangedVM = Get-VM $VM
        $vGPUDevice = $ChangedVM.ExtensionData.Config.hardware.Device | Where { $_.backing.vgpu}
        $vGPUDevice | Select Key, ControllerKey, Unitnumber, @{Name="Device";Expression={$_.DeviceInfo.Label}}, @{Name="Summary";Expression={$_.DeviceInfo.Summary}}
  
    }   
}


Function Convert-ExcelSheetToJson{
[CmdletBinding()]
Param(
    [Parameter(
        ValueFromPipeline=$true,
        Mandatory=$true
        )]
    [Object]$InputFile,

    [Parameter()]
    [string]$OutputFileName,

    [Parameter()]
    [string]$SheetName
    )

if ($InputFile -is "System.IO.FileSystemInfo") {
    $InputFile = $InputFile.FullName.ToString()
}
# Make sure the input file path is fully qualified
$InputFile = [System.IO.Path]::GetFullPath($InputFile)
Write-Verbose "Converting '$InputFile' to JSON"

# If no OutputfileName was specified, make one up
if (-not $OutputFileName) {
    $OutputFileName = [System.IO.Path]::GetFileNameWithoutExtension($(Split-Path $InputFile -Leaf))
    $OutputFileName = Join-Path $pwd ($OutputFileName + ".json")
}
# Make sure the output file path is fully qualified
$OutputFileName = [System.IO.Path]::GetFullPath($OutputFileName)

# Instantiate Excel
#$excelApplication = New-Object -ComObject Excel.Application
#$excelApplication.DisplayAlerts = $false
#$Workbook = $excelApplication.Workbooks.Open($InputFile)

$Excel = New-Excel -Path $InputFile
$SheetName = $Excel | Get-Worksheet -Name $SheetName
$workbook = $Excel | Get-Workbook

# If SheetName wasn't specified, make sure there's only one sheet
if (-not $SheetName) {
    if ($workbook.Worksheets.count-eq 1) {
        $SheetName = @($Workbook.Worksheets)[0].Name
        Write-Verbose "SheetName was not specified, but only one sheet exists. Converting '$SheetName'"
    } else {
        throw "SheetName was not specified and more than one sheet exists."
    }
} else {
    # If SheetName was specified, make sure the sheet exists
    $theSheet = $Workbook.Worksheets | Where-Object {$_.Name -eq $SheetName}
    if (-not $theSheet) {
        throw "Could not locate SheetName '$SheetName' in the workbook"
    }
}
Write-Verbose "Outputting sheet '$SheetName' to '$OutputFileName'"
#endregion prep


# Grab the sheet to work with
$theSheet = $Workbook.Worksheets | Where-Object {$_.Name -eq $SheetName}

#region headers
# Get the row of headers
$Headers = @{}
$NumberOfColumns = 0
$FoundHeaderValue = $true
while ($FoundHeaderValue -eq $true) {
    $cellValue = $theSheet.Cells.Item(1, $NumberOfColumns+1).Text
    if ($cellValue.Trim().Length -eq 0) {
        $FoundHeaderValue = $false
    } else {
        $NumberOfColumns++
        $Headers.$NumberOfColumns = $cellValue
    }
}
#endregion headers

# Count the number of rows in use, ignore the header row
$rowsToIterate = $theSheet.Dimension.Rows

#region rows
$results = @()
foreach ($rowNumber in 2..$rowsToIterate+1) {
    if ($rowNumber -gt 1) {
        $result = @{}
        foreach ($columnNumber in $Headers.GetEnumerator()) {
            $ColumnName = $columnNumber.Value
            $CellValue = $theSheet.Cells.Item($rowNumber, $columnNumber.Name).Value
            $result.Add($ColumnName,$cellValue)
        }
        $results += $result
    }
}
#endregion rows


$results | ConvertTo-Json | Out-File -Encoding ASCII -FilePath $OutputFileName -Force

Get-Item $OutputFileName

# Close the Workbook
#$excelApplication.Workbooks.Close()
# Close Excel
#[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApplication)

}

function testport ($hostname='',$port=443,$timeout=100) {
  $requestCallback = $state = $null
  $client = New-Object System.Net.Sockets.TcpClient
  $beginConnect = $client.BeginConnect($hostname,$port,$requestCallback,$state)
  Start-Sleep -milli $timeOut
  if ($client.Connected) { $open = $true } else { $open = $false }
  $client.Close()
  [pscustomobject]@{hostname=$hostname;port=$port;open=$open}
}

function Convert-Size {            
    [cmdletbinding()]            
    param(            
        [validateset("Bytes","KB","MB","GB","TB")]            
        [string]$From,            
        [validateset("Bytes","KB","MB","GB","TB")]            
        [string]$To,            
        [Parameter(Mandatory=$true)]            
        [double]$Value,            
        [int]$Precision = 2            
    )            
    switch($From) {            
        "Bytes" {$value = $Value }            
        "KB" {$value = $Value * 1024 }            
        "MB" {$value = $Value * 1024 * 1024}            
        "GB" {$value = $Value * 1024 * 1024 * 1024}            
        "TB" {$value = $Value * 1024 * 1024 * 1024 * 1024}            
    }            
                
    switch ($To) {            
        "Bytes" {return $value}            
        "KB" {$Value = $Value/1KB}            
        "MB" {$Value = $Value/1MB}            
        "GB" {$Value = $Value/1GB}            
        "TB" {$Value = $Value/1TB}            
                
    }            
                
    return [Math]::Round($value,$Precision,[MidPointRounding]::AwayFromZero)            
                
    }            

function Confirm-RayStationObject {
PARAM
(
  [ValidateSet('2019','2022')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $OS
  ,
  [ValidateLength(3,15)]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $hostname
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [string] $ipaddress
  ,
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateRange(1,32)] 
  [int] $netmask
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $gateway
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $DNS1
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(ValueFromPipelineByPropertyName = $true)]
  [string] $DNS2
  ,
  [ValidatePattern('^((?!-)[A-Za-z0-9-]{1,63}(?<!-)\.)+[A-Za-z]{2,6}$')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $domain
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $hostIP
  ,
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $datastore
  ,
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $network
  ,
  [ValidateRange(75, [int]::MaxValue)]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [int] $diskGB
  ,
  [ValidateRange(1,[int]::MaxValue)]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [int] $vCPU
  ,
  [ValidateRange(1,[int]::MaxValue)]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [int] $RAM
  ,
  [AllowNull()]
  [ValidateSet('vGPU','Passthrough',$null)]
  [Parameter(ValueFromPipelineByPropertyName = $true)]
  [string] $GPUType
  ,
  [Parameter(ValueFromPipelineByPropertyName = $true)]
  [string] $GPU
  ,
  [Parameter(ValueFromPipelineByPropertyName = $true)]
  [string] $CardUUID
  ,
  [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
  $InputObject
)
  return $true;
}

function Confirm-NvidiaObject {
PARAM
(
  [ValidateLength(3,15)]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $VMName
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [string] $ipaddress
  ,
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateRange(1,32)] 
  [int] $netmask
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $gateway
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $DNS1
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(ValueFromPipelineByPropertyName = $true)]
  [string] $DNS2
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $hostIP
  ,
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $datastore
  ,
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $network
  ,
  [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
  $InputObject
)
  return $true;
}

function Confirm-SQLObject {
PARAM
(
  [ValidateSet('2019','2022')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $OS
  ,
  [ValidateLength(3,15)]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $hostname
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [string] $ipaddress
  ,
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateRange(1,32)] 
  [int] $netmask
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $gateway
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $DNS1
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(ValueFromPipelineByPropertyName = $true)]
  [string] $DNS2
  ,
  [ValidatePattern('^((?!-)[A-Za-z0-9-]{1,63}(?<!-)\.)+[A-Za-z]{2,6}$')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $domain
  ,
  [ValidatePattern('^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$')]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $hostIP
  ,
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $datastore
  ,
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [string] $network
  ,
  [ValidateRange(75, [int]::MaxValue)]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [int] $diskGB
  ,
  [ValidateRange(1,[int]::MaxValue)]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [int] $vCPU
  ,
  [ValidateRange(1,[int]::MaxValue)]
  [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
  [ValidateNotNullOrEmpty()]
  [int] $RAM
  ,
  [AllowNull()]
  [Parameter(ValueFromPipelineByPropertyName = $true)]
  [string] $AdditionalDiskDatastore
  ,
  [AllowNull()] 
  [Validatescript({
            if ($_ -ge "1" -or $_ -match $null) {$true}
            else { throw $false}})]
  [Parameter(ValueFromPipelineByPropertyName = $true)]
  [int] $DiskSize1
  ,
  [AllowNull()]
  [Validatescript({
            if ($_ -ge "1" -or $_ -match $null) {$true}
            else { throw $false}})]
  [Parameter(ValueFromPipelineByPropertyName = $true)]
  [int] $DiskSize2
  ,
  [AllowNull()] 
  [Validatescript({
            if ($_ -ge "1" -or $_ -match $null) {$true}
            else { throw $false}})]
  [Parameter(ValueFromPipelineByPropertyName = $true)]
  [int] $DiskSize3
  ,
  [AllowNull()] 
  [Validatescript({
            if ($_ -ge "1" -or $_ -match $null) {$true}
            else { throw $false}})]
  [Parameter(ValueFromPipelineByPropertyName = $true)]
  [int] $DiskSize4
  ,
  [AllowNull()] 
  [Validatescript({
            if ($_ -ge "1" -or $_ -match $null) {$true}
            else { throw $false}})]
  [Parameter(ValueFromPipelineByPropertyName = $true)]
  [int] $DiskSize5
  ,
  [AllowNull()] 
  [Validatescript({
            if ($_ -ge "1" -or $_ -match $null) {$true}
            else { throw $false}})]
  [Parameter(ValueFromPipelineByPropertyName = $true)]
  [int] $DiskSize6
  ,
  [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
  $InputObject
)
  return $true;

}

Function Confirm-UserFormat
{
PARAM
(
  [ValidatePattern('^((?!-)[A-Za-z0-9-_]{1,63}(?<!-)\\)+((?!-)[A-Za-z0-9-_]{1,63})$')]
  [Parameter(Mandatory = $true)]
  [string] $Username
  )
  return $true;
  }

#--------------------------------------------------------------------------------------------------------------------


#$RunLocation = split-path -parent $MyInvocation.MyCommand.Definition

Disconnect-VIServer -Confirm:$false -ErrorAction SilentlyContinue


if($skipexcelimport -eq $False){
write-host "Converting Excel to Json" -ForegroundColor Green

if(Test-Path -Path "$runLocation\VMs.xlsx"){
"$runLocation\VMs.xlsx" | Convert-ExcelSheetToJson -OutputFileName "$RunLocation\bin\RayStation.json" -SheetName "RayStation" | out-null
"$runLocation\VMs.xlsx" | Convert-ExcelSheetToJson -OutputFileName "$RunLocation\bin\Nvidia.json" -SheetName "Nvidia" | out-null
"$runLocation\VMs.xlsx" | Convert-ExcelSheetToJson -OutputFileName "$RunLocation\bin\SQL.json" -SheetName "SQL" | Out-Null
}
else{

write-host "Missing VMs.xlsx file" -ForegroundColor Red
break
}

}

write-host "Importing Json" -ForegroundColor Green

if(Test-Path -Path "$RunLocation\bin\RayStation.json"){
$Sheet = Get-Content "$RunLocation\bin\RayStation.json"
$VMs = $sheet | ConvertFrom-Json
}else{
   Write-Host "Missing RayStation.json" -ForegroundColor Red
   break
}
if(Test-Path -Path "$RunLocation\bin\Nvidia.json"){
$Sheet = Get-Content "$RunLocation\bin\Nvidia.json"
$NvidiaVMs = $sheet | ConvertFrom-Json
}else{
   Write-Host "Missing Nvidia.json" -ForegroundColor Red
   break
}
if(Test-Path -Path "$RunLocation\bin\SQL.json"){
$Sheet = Get-Content "$RunLocation\bin\SQL.json"
$SQLVMs = $sheet | ConvertFrom-Json
}else{
   Write-Host "Missing SQL.json" -ForegroundColor Red
   break
}

$TotalVMs = @()

$i = 0
$NoRSVMs = $False
$IsNull = ([System.Object[]] $VMs | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Definition).split('=') | Where-Object { $i % 2 -eq 1; $i++ }

if(($IsNull | Select-Object -Unique).count -eq "1"){

$NoRSVMs = $True

}
if($NoRSVMs -eq $False){
$TotalVMs += $VMs
}
$i = 0
$NoNvidiaVMs = $False
$IsNull = ([System.Object[]] $NvidiaVMs | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Definition).split('=') | Where-Object { $i % 2 -eq 1; $i++ }

if(($IsNull | Select-Object -Unique).count -eq "1"){

$NoNvidiaVMs = $True

}
$i = 0
$NoSQLVMs = $False
$IsNull = ([System.Object[]] $SQLVMs | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Definition).split('=') | Where-Object { $i % 2 -eq 1; $i++ }

if(($IsNull | Select-Object -Unique).count -eq "1"){

$NoSQLVMs = $True

}
if($NoSQLVMs -eq $False){
$TotalVMs += $SQLVMs
}



Write-Host "Validating Excel Values" -ForegroundColor Green

$IsValidated = $True
$ValidationErrors =@()
$ListedIPs = @()
$ListedHostnames = @()
$ListedVMNames = @()

if($NoRSVMs -ne $True){
foreach($VM in $VMs){

try{
$Validation = $VM | Confirm-RayStationObject -ErrorAction:Stop;
$ListedIPs += $vm.ipaddress
$ListedHostnames += $VM.hostname
$ListedVMNames += $VM.hostname
}
catch{

$ValidationErrors += $_.exception
$IsValidated = $False
}

}
}
if($NoNvidiaVMs -ne $True){
foreach($NvidiaVM in $NvidiaVMs){

try{
$Validation = $NvidiaVM | Confirm-NvidiaObject -ErrorAction:Stop;
$ListedIPs += $NvidiaVM.ipaddress
$ListedVMNames += $NvidiaVM.VMName

}
catch{

$ValidationErrors += $_.exception
$IsValidated = $False
}

}
}
if($NoSQLVMs -ne $True){
foreach($SQLVM in $SQLVMs){

try{

$Validation = $SQLVM | Confirm-SQLObject -ErrorAction:Stop;
$ListedIPs += $SQLVM.ipaddress
$ListedHostnames += $SQLVM.hostname
$ListedVMNames += $SQLVM.hostname
}
catch{

$ValidationErrors += $_.exception
$IsValidated = $False
}

}
}
if($IsValidated -eq $False){

write-host $ValidationErrors -ForegroundColor Red
break
}

if(($ListedIPs | Select-Object -Unique).count -ne ($ListedIPs).count){

write-host "Duplicate IP Addresses found. Please Adjust VMs.xlsx" -ForegroundColor Red
break

}

if(($ListedHostnames | Select-Object -Unique).count -ne ($ListedHostnames).count){

write-host "Duplicate Hostnames found. Please Adjust VMs.xlsx" -ForegroundColor Red
break

}

if(($ListedVMNames | Select-Object -Unique).count -ne ($ListedVMNames).count){

write-host "Duplicate VMNames found. Please Adjust VMs.xlsx" -ForegroundColor Red
break

}


if($skipnetworkcheck -eq $false){

write-host "Checking for in use IPs and Hostnames. This may take some time." -ForegroundColor Green

$Connection = $ListedIPs  | ForEach-Object { Test-Connection -ComputerName $_ -Count 1 -AsJob } | Get-Job | Receive-Job -Wait | Select-Object @{Name='ComputerName';Expression={$_.Address}},@{Name='Reachable';Expression={if ($_.StatusCode -eq 0) { $true } else { $false }}}

$Reachable = $Connection | Where-Object {$_.Reachable -Match "True"}

if($Reachable -ne $null){

$Reachable = $Reachable | Format-Table -AutoSize | Out-String

Write-host "The following IP addresses are in use.`n $Reachable `nPlease fix IP or remove before continuing."
Break

}

$Connection = $ListedHostnames  | ForEach-Object { Test-Connection -ComputerName $_ -Count 1 -AsJob } | Get-Job | Receive-Job -Wait | Select-Object @{Name='ComputerName';Expression={$_.Address}},@{Name='Reachable';Expression={if ($_.StatusCode -eq 0) { $true } else { $false }}}

$Reachable = $Connection | Where-Object {$_.Reachable -Match "True"}

if($Reachable -ne $null){

$Reachable = $Reachable | Format-Table -AutoSize | Out-String

Write-host "The following Hostnames are in use.`n $Reachable `nPlease fix Hostname or remove before continuing."
Break

}
}
elseif($skipnetworkcheck -eq $True){

Write-Host "Skipping Network Check" -ForegroundColor Yellow

}

$ListedHostnames | out-file "$RunLocation\RemoteMachines.txt" -force

$vcenterAddress = Read-Host "Enter the vCenter Address"
$vcenterUsername = Read-Host "Enter the vCenter Username"
$vcenterPassword = Read-Host "Enter the vCenter Password" -AsSecureString

$username = Read-Host "AD Join Username"
$PasswordSecure = Read-Host "AD Join Password" -AsSecureString
$RSAccount = Read-Host "RayStation Installation Account"


if($skipadcheck -ne $True){
Write-Host "Testing AD Creds" -ForegroundColor Green

try{
$Validation = Confirm-UserFormat $username -ErrorAction:Stop;
}
catch{
$_.exception
break
}
try{
$Validation = Confirm-UserFormat $RSAccount -ErrorAction:Stop;
}
catch{
$_.exception
break
}


$Remotecreds = New-Object System.Management.Automation.PSCredential ($Username, $PasswordSecure)

$UsernameFQDN = $Username
$Domain = $username.split("\")[0]
$usernameClean = $username.split("\")[1]

if($Domain -notmatch '`.' -or $Domain -notmatch 'localhost'){
         $CredCheck = $Remotecreds  | Test-Cred
         If($CredCheck -ne "Authenticated"){
           Write-Warning "AD Credential validation failed"
            Break
            }
        }

if((Test-ADuser $usernameClean) -eq $False){

Write-Host "Unable to find RayStation Instalation Account" -ForegroundColor Red
Break
}

}
else{
Write-Host "Skipping AD User Check." -ForegroundColor Yellow
}




$cred = New-Object System.Management.Automation.PSCredential ($vcenterUsername, $vcenterPassword)
	
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | out-null
Set-PowerCLIConfiguration -DefaultVIServerMode multiple -Confirm:$false | out-null
Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -Confirm:$false | out-null


$open = testport $vcenterAddress

if($open.open -eq $True){
try{
Write-Host "Connecting to vCenter" -ForegroundColor Green
Connect-VIServer $vcenterAddress -Credential $cred -ErrorAction Stop | Out-Null
}
Catch{
Write-host "Unable to Connect to vCenter at: $vcenterAddress. Please Check Credentials." -ForegroundColor Red
$_.Exception
break
}
}
if($open.open -eq $false){

Write-Host "Unable to Connect to vCenter: $vcenterAddress on Port 443. Please verify that vCenter is up and reachable." -ForegroundColor Red
break

}


$jsonBase = @()

Write-Host "Getting Host Inventory" -ForegroundColor Green
$Servers = get-vmhost
foreach($Server in $Servers){

$Hosts = [ordered]@{}

$ServerName = $server.Name
$CPUCapacityGHz = [math]::Round($server.CpuTotalMhz/1000,2)
$MemoryCapacityGB = [math]::Round($server.MemoryTotalGB,2)

$BasicInfo = [ordered]@{"HOSTNAME"="$ServerName";"CPUCapacityGHz"="$CPUCapacityGHz";"MemoryCapacity"="$MemoryCapacityGB"}

$Hosts.add("HOST",$BasicInfo) | out-null

$DataStoreList = New-Object System.Collections.ArrayList
$Datastores = $Server | Get-Datastore
foreach($Datastore in $Datastores){
$FreeSpace = Get-Datastore -Name $Datastore.name | Select-Object -ExpandProperty FreeSpaceGB
$CapacityGB = Get-Datastore -Name $Datastore.name | Select-Object -ExpandProperty CapacityGB
$DatastoreName = $Datastore.name 
$FreeSpace = [math]::truncate($FreeSpace)
[void]$DataStoreList.Add([ordered]@{"DataStoreName"="$DatastoreName";"FreeSpaceGB"="$FreeSpace";"TotalCapacityGB"="$CapacityGB";})

}
$Hosts.add("DataStores",$DataStoreList) | out-null

$Portgrouplist = New-Object System.Collections.ArrayList
$Portgroups = $Server |Get-VirtualPortgroup
foreach($portgroup in $Portgroups){
$PortGroupName = $portgroup.Name
$Portgrouplist += $PortGroupName
}
$Hosts.add("PortGroups",$Portgrouplist) | out-null

$vGPUList = New-Object System.Collections.ArrayList

$SupportedProfiles = $Server.ExtensionData.Config.SharedPassthruGpuTypes

$TotalMemoryKBArray = @()
$GPUs = $Server.ExtensionData.Config.GraphicsInfo | where-object {$_.GraphicsType -match "sharedDirect"}
foreach($GPU in $GPUs){

$TotalMemoryKBArray += $gpu.MemorySizeInKB

}
$TotalMemoryKB = ($TotalMemoryKBArray | Measure-Object -Sum).Sum
$TotalMemoryGB = Convert-Size -From KB -to GB -value $TotalMemoryKB

#$TotalMemoryGB

foreach($SupportedProfile in $SupportedProfiles){

$CardName = (($SupportedProfile -split "_")[1] -split "-")[0]

$ProfileMemory = ($SupportedProfile -split "-")[1] -replace "[^0-9]" , ''

$SupportedProfileCount = [math]::ceiling($TotalMemoryGB / $ProfileMemory)

[void]$vGPUList.add([ordered]@{"CardType"=$CardName;"ProfileName"=$SupportedProfile;"AvaiablePerHost"=$SupportedProfileCount})

}

$Hosts.add("vGPUs",$vGPUList) | out-null


$PassthroughGPUs = New-Object System.Collections.ArrayList
$PassthroughDevices = @()
$GPUs = $Server.ExtensionData.Config.GraphicsInfo

foreach($GPU in $GPUs){

if($GPU.GraphicsType -match "basic" -or $GPU.GraphicsType -eq "Direct"){

$BasicGPUs = $GPU

$PassthroughDevices += Get-PassthroughDevice -Type Pci -VMHost $Server | Where-Object {$_.VendorName -match "Nvidia" -and (($_.uid.split('-'))[3] -replace '/','') -match $BasicGPUs.PciID}
    }
}

foreach($PassthroughDevice in $PassthroughDevices){

$GPUName = $PassthroughDevice.Name
$GPUID = ($PassthroughDevice.uid.split('-'))[3] -replace '/',''
$ID = $PassthroughDevice.DeviceId

$PassthroughGPUs.add([ordered]@{"CardName"=$GPUName;"UUID"=$GPUID;"ID"=$ID}) | out-null
}

$Hosts.add("PassthroughGPUs",$PassthroughGPUs) | out-null


$jsonBase += $hosts

}

if($skipoverprovisioncheck -eq $False){
Write-Host "Validating that Environment will not be Overprovisioned" -ForegroundColor Green
$Hosts = $jsonBase
$Valid = $True
$PassthroughGPUIDs = @()
foreach ($VM in $TotalVMs) {

    foreach ($VMHost in $Hosts) {

        if ($VM.hostIP -match $VMHost.host.hostname) {
                $ESXiHostname = $vmhost.HOST.Hostname
                $VMPortGroup = $VM.network
                $VMHostname = $VM.hostname

            $newRam = $VMHost.host.MemoryCapacity - $VM.RAM
            $VMHost.host.MemoryCapacity = $newRam

            $Datastores = $VMHost.DataStores

            foreach ($Datastore in $datastores) {

                $DSNames = $Datastore.DataStoreName
                foreach ($DSName in $DSNames) {

                    if ($DSName -contains $VM.datastore) {

                        $NewFreeSpaceGB = $datastore.FreeSpaceGB - $VM.DiskGB

                        $Hosts.DataStores | Where { $_.DataStoreName -eq $DSName } | foreach { $_.FreeSpaceGB = $NewFreeSpaceGB }

                    }
                }

            }
             if((![string]::IsNullOrEmpty($VM.AdditionalDiskDatastore))){
                $TotalAdditionalStorage = 0
                
                if((![string]::IsNullOrEmpty($VM.DiskSize1))){

                $TotalAdditionalStorage = $TotalAdditionalStorage + $VM.DiskSize1
            }
                if((![string]::IsNullOrEmpty($VM.DiskSize2))){

                $TotalAdditionalStorage = $TotalAdditionalStorage + $VM.DiskSize2
            }
                if((![string]::IsNullOrEmpty($VM.DiskSize3))){

                $TotalAdditionalStorage = $TotalAdditionalStorage + $VM.DiskSize3
            }
                if((![string]::IsNullOrEmpty($VM.DiskSize4))){

                $TotalAdditionalStorage = $TotalAdditionalStorage + $VM.DiskSize4
            }
                if((![string]::IsNullOrEmpty($VM.DiskSize5))){

                $TotalAdditionalStorage = $TotalAdditionalStorage + $VM.DiskSize5
            }
                if((![string]::IsNullOrEmpty($VM.DiskSize6))){

                $TotalAdditionalStorage = $TotalAdditionalStorage + $VM.DiskSize6
            }

            foreach ($Datastore in $datastores) {

                $DSNames = $Datastore.DataStoreName
                foreach ($DSName in $DSNames) {

                    if ($DSName -contains $VM.AdditionalDiskDatastore) {

                        $NewFreeSpaceGB = $datastore.FreeSpaceGB - $TotalAdditionalStorage

                        $Hosts.DataStores | Where { $_.DataStoreName -eq $DSName } | foreach { $_.FreeSpaceGB = $NewFreeSpaceGB }
                        }
                       }
                      }
                    }

            $PortGroups = $VMHost.PortGroups
            if ($portGroups -notcontains $VM.network) {

                $Valid = $false
                Write-Host "ESXi Host: $ESXiHostname does not contain Portgroup: $VMPortGroup for VM: $VMHostname" -ForegroundColor Red

            }

            $IDs = $VMHost.PassthroughGPUs.UUID

            if((![string]::IsNullOrEmpty($VM.CardUUID)) -and $VM.GPUType -contains "passthrough"){

            if ($IDs -contains $VM.CardUUID -and $VM.GPUType -contains "Passthrough") {

                $PassthroughGPUIDs += $VM.CardUUID

            }
            else {
                $Valid = $false

                $CardUUID = $VM.CardUUID

                Write-Host "ESXi Host: $ESXiHostname does not contain GPU ID: $CardUUID for VM: $VMHostname" -ForegroundColor Red


            }
            }

            if($VM.GPUType -contains "vGPU"){

                if($VMHost.vgpus.ProfileName -contains $VM.GPU){

                    $Profiles = $VMHost.vgpus

                    $SelectedProfile = $Profiles | Where-Object {$_.ProfileName -match $VM.GPU}

                    $NewCount = $SelectedProfile.AvaiablePerHost - 1

                    $VMHost.vgpus | where {$_.ProfileName -eq $SelectedProfile.ProfileName} | foreach {$_.AvaiablePerHost = $NewCount}

                }

            }


        }

    }

}
if($NoNvidiaVMs -ne $True){

    foreach($VM in $NvidiaVMs){
            
            foreach ($VMHost in $Hosts) {
                        if ($VM.hostIP -match $VMHost.host.hostname) {
                            $ESXiHostname = $vmhost.HOST.Hostname
                            $VMPortGroup = $VM.network
                            $VMHostname = $VM.VMName

                            $newRam = $VMHost.host.MemoryCapacity - 8
                            $VMHost.host.MemoryCapacity = $newRam

                            $Datastores = $VMHost.DataStores

                            foreach ($Datastore in $datastores) {

                                $DSNames = $Datastore.DataStoreName
                                foreach ($DSName in $DSNames) {

                                    if ($DSName -contains $VM.datastore) {

                                        $NewFreeSpaceGB = $datastore.FreeSpaceGB - 10

                                        $Hosts.DataStores | Where { $_.DataStoreName -eq $DSName } | foreach { $_.FreeSpaceGB = $NewFreeSpaceGB }

                           }
                      }
                 }
            $PortGroups = $VMHost.PortGroups
            if ($portGroups -notcontains $VM.network) {

                $Valid = $false
                Write-Host "ESXi Host: $ESXiHostname does not contain Portgroup: $VMPortGroup for VM: $VMHostname" -ForegroundColor Red

            }

           }
        }

    }

}

if (($PassthroughGPUIDs | where-object { (!([string]::IsNullOrEmpty($_))) } | Select-Object -Unique).count -ne ($PassthroughGPUIDs | where-object { (!([string]::IsNullOrEmpty($_))) }).count) {
    $Valid = $false
    Write-Host "GPU IDs are not Unique for VMs" -ForegroundColor Red
}

 foreach ($VMHost in $Hosts) {
    $ESXiHostname = $vmhost.HOST.Hostname
    if($VMHost.HOST.MemoryCapacity -le 0){
        $Valid = $false
        Write-Host "$ESXiHostname has overprovisioned RAM. Please adjust VM Specs." -ForegroundColor Red
    }
       
    $Datastores = $VMHost.DataStores
    foreach ($Datastore in $datastores) {
        $DatastoreName = $Datastore.DataStoreName
        if($Datastore.FreeSpaceGB -le 0){
            $Valid = $false
            Write-Host "DataStore: $DatastoreName is overprovisioned on host: $ESXiHostname" -ForegroundColor Red
        }
    if((!([string]::IsNullOrEmpty($VMHost.vGPUs)))){
        foreach($VGPU in $VMHost.vGPUs){
            
            if($VGPU.AvaiablePerHost -lt 0){
                $Valid = $false
                Write-Host "vGPUs overprovisioned on host: $ESXiHostname" -ForegroundColor Red
            }

          }
    
        }

    }


 }

 if($Valid -eq $false){
    Write-Host "Host(s) will not be able to support VMs" -ForegroundColor Red
    Break
 }
}
elseif($skipoverprovisioncheck -eq $True){

Write-Host "Skipping Overprovisioning Check" -ForegroundColor Yellow
}

write-host "Checking for Duplicate VM Names" -ForegroundColor Green

$CurrentVMNames = (Get-VM).name
$DuplicateVMNames = @()

   foreach($ListedVMName in $ListedVMNames){
        if($CurrentVMNames -contains $ListedVMName){
            $DuplicateVMNames += $ListedVMName
        }
   }
    
    if($DuplicateVMNames.count -gt "0"){

        Write-Host "The following VMNames are already in use:`n$DuplicateVMNames"
        break
    }



    Write-Host "Attempting to Encrypt Password" -ForegroundColor Green
    try{
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PasswordSecure)
    $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    $Password = $Password | Protect-PEMString -PublicKey "$RunLocation\Bin\protected.pem"
    write-Host "Encrypted Password" -ForegroundColor Green
    }
    Catch{
    Write-host "Failed to Encrypt Password, Falling back to Clear Text" -ForegroundColor Red
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PasswordSecure)
    $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    }

$DeploymentTime = [diagnostics.stopwatch]::StartNew()

if($NoRSVMs -ne $True){
write-host "Deploying RayStation VMs" -ForegroundColor Green
foreach($VMJson in $VMs){

$OS =$vmJson.OS


if($OS -Match "2022"){
$ovfPath = "$RunLocation\OVAs\2022Template.ova"
}
if($OS -Match "2019"){
$ovfPath = "$RunLocation\OVAs\2019Template.ova"
}


$ovfConfig = Get-OvfConfiguration -Ovf $ovfPath


$dns1 = $vmJson.DNS1                  
$gateway = $vmJson.gateway     
$VMNetwork = $vmJson.network
$ipaddress = $vmJson.ipaddress             
$ad_username = $username
$IpProtocol = "IPv4"
$ad_password = $Password       
$netmask = $vmJson.netmask  
$rs_account = $RSAccount           
$domain = $vmJson.domain
$hostname = $vmJson.hostname  
$dns2 = $vmJson.DNS2
$ad_domain = $vmJson.domain

$ovfConfig.common.guestinfo.dns1.Value = $dns1                   
$ovfConfig.common.guestinfo.gateway.Value = $gateway                    
$ovfConfig.NetworkMapping.VM_Network.Value = $VMNetwork         
$ovfConfig.common.guestinfo.ipaddress.Value = $ipaddress 
$ovfConfig.common.guestinfo.ad_username.Value = $ad_username         
$ovfConfig.IpAssignment.IpProtocol.Value = $IpProtocol       
$ovfConfig.common.guestinfo.ad_password.Value = $ad_password     
$ovfConfig.common.guestinfo.netmask.Value = $netmask       
$ovfConfig.common.guestinfo.rs_account.Value =  $rs_account          
$ovfConfig.common.guestinfo.domain.Value = $domain        
$ovfConfig.common.guestinfo.hostname.Value =  $hostname               
$ovfConfig.common.guestinfo.dns2.Value =$dns2
$ovfConfig.common.guestinfo.ad_domain.Value = $ad_domain

$VMHost = $vmJson.hostIP
$Datastore = $vmJson.datastore
$DiskFormat = "Thin"


$null = Import-VApp -Source $ovfpath -OvfConfiguration $ovfConfig -Name $hostname -VMHost $VMHost -Datastore $Datastore -DiskStorageFormat $DiskFormat -RunAsync -Confirm:$false
}

}

if($NoSQLVMs -ne $True){
Write-Host "Deploying SQL VMs" -ForegroundColor Green

foreach($SQLVM in $SQLVMs){

    $OS =$SQLVM.OS
    
    
    if($OS -Match "2022"){
    $ovfPath = "$RunLocation\OVAs\2022Template.ova"
    }
    if($OS -Match "2019"){
    $ovfPath = "$RunLocation\OVAs\2019Template.ova"
    }
    
    
    $ovfConfig = Get-OvfConfiguration -Ovf $ovfPath
    
    
    $dns1 = $SQLVM.DNS1                  
    $gateway = $SQLVM.gateway     
    $VMNetwork = $SQLVM.network
    $ipaddress = $SQLVM.ipaddress             
    $ad_username = $username
    $IpProtocol = "IPv4"
    $ad_password = $Password       
    $netmask = $SQLVM.netmask  
    $rs_account = $RSAccount           
    $domain = $SQLVM.domain
    $hostname = $SQLVM.hostname  
    $dns2 = $SQLVM.DNS2
    $ad_domain = $SQLVM.domain
    
    $ovfConfig.common.guestinfo.dns1.Value = $dns1                   
    $ovfConfig.common.guestinfo.gateway.Value = $gateway                    
    $ovfConfig.NetworkMapping.VM_Network.Value = $VMNetwork         
    $ovfConfig.common.guestinfo.ipaddress.Value = $ipaddress 
    $ovfConfig.common.guestinfo.ad_username.Value = $ad_username         
    $ovfConfig.IpAssignment.IpProtocol.Value = $IpProtocol       
    $ovfConfig.common.guestinfo.ad_password.Value = $ad_password     
    $ovfConfig.common.guestinfo.netmask.Value = $netmask       
    $ovfConfig.common.guestinfo.rs_account.Value =  $rs_account          
    $ovfConfig.common.guestinfo.domain.Value = $domain        
    $ovfConfig.common.guestinfo.hostname.Value =  $hostname               
    $ovfConfig.common.guestinfo.dns2.Value =$dns2
    $ovfConfig.common.guestinfo.ad_domain.Value = $ad_domain
    
    $VMHost = $SQLVM.hostIP
    $Datastore = $SQLVM.datastore
    $DiskFormat = "Thin"
    
    
    $null = Import-VApp -Source $ovfpath -OvfConfiguration $ovfConfig -Name $hostname -VMHost $VMHost -Datastore $Datastore -DiskStorageFormat $DiskFormat -RunAsync -Confirm:$false
    }
}
if($NoNvidiaVMs -ne $True){
Write-Host "Deploying Nvidia License Appliance(s)" -ForegroundColor Green

foreach($NvidiaVM in $NvidiaVMs){


    $ovfPath = "$RunLocation\OVAs\nls-3.1.0-bios.ova"
 
    $ovfConfig = Get-OvfConfiguration -Ovf $ovfPath

    $hostname = $NvidiaVM.VMName
    $dns1 = $NvidiaVM.DNS1                  
    $gateway = $NvidiaVM.gateway     
    $VMNetwork = $NvidiaVM.network
    $ipaddress = $NvidiaVM.ipaddress             
    $IpProtocol = "IPv4"      
    $netmask = $NvidiaVM.netmask  
    $dns2 = $NvidiaVM.DNS2
           
    $ovfConfig.NetworkMapping.VM_Network.Value = $VMNetwork                  
    $ovfConfig.IpAssignment.IpProtocol.Value = $IpProtocol     
    $ovfConfig.NetworkProperty.ipaddress.Value = $ipaddress  
    $ovfConfig.NetworkProperty.netmask.Value = $netmask    
    $ovfConfig.NetworkProperty.gateway.Value =  $gateway       
    $ovfConfig.NetworkProperty.dns_server_one.Value = $dns1      
    $ovfConfig.NetworkProperty.dns_server_two.Value =  $dns2 
    
    $VMHost = $NvidiaVM.hostIP
    $Datastore = $NvidiaVM.datastore
    $DiskFormat = "Thin"
    
    
    $null = Import-VApp -Source $ovfpath -OvfConfiguration $ovfConfig -Name $hostname -VMHost $VMHost -Datastore $Datastore -DiskStorageFormat $DiskFormat -RunAsync -Confirm:$false
    }
}

while((Get-task| Where-Object {($_.State -eq "Running") -and $_.Name -Match "Deploy OVF template"}).count -gt 0 )
{   

   $i = 0  
    $tasks = (Get-task | Where-Object {$_.Name -Match "Deploy OVF template"})
    foreach($task in $Tasks){
    $TaskObject = [PSCustomObject] @{
                'Entity' = $Task.ExtensionData.Info.EntityName
                'Description' = $Task.Description
                'Status' = $Task.State
                'Progress' = $Task.PercentComplete
                'Username' = $Task.ExtensionData.Info.Reason.UserName
                'Message' = $Task.ExtensionData.Info.Description.Message
                'Id' = $Task.Id
                'StartTime' = $Task.ExtensionData.Info.StartTime
                'CompleteTime' = $Task.ExtensionData.Info.CompleteTime
                'IsCancellable' = $Task.ExtensionData.Info.Cancelable
                'Index' = $i
            }


            if($TaskObject.Status -eq "Running"){

                WriteJobProgress -job $task -Index $TaskObject.Index -currentOp $TaskObject.Entity;

             }
            if($TaskObject.Status -eq "Success"){
                WriteJobProgress -Job $task -Index $TaskObject.Index -Completed $true
             }
            if($TaskObject.Status -eq "Error"){
                WriteJobProgress -Job $task -Index $TaskObject.Index -Completed $true
             }

        $i++
    }

start-sleep -Seconds 2

}

$i = 0  
    $tasks = (Get-task | Where-Object {$_.Name -Match "Deploy OVF template"})
    foreach($task in $Tasks){
    $TaskObject = [PSCustomObject] @{
                'Entity' = $Task.ExtensionData.Info.EntityName
                'Description' = $Task.Description
                'Status' = $Task.State
                'Progress' = $Task.PercentComplete
                'Username' = $Task.ExtensionData.Info.Reason.UserName
                'Message' = $Task.ExtensionData.Info.Description.Message
                'Id' = $Task.Id
                'StartTime' = $Task.ExtensionData.Info.StartTime
                'CompleteTime' = $Task.ExtensionData.Info.CompleteTime
                'IsCancellable' = $Task.ExtensionData.Info.Cancelable
                'Index' = $i
            }
                if($TaskObject.Status -eq "Success"){
                WriteJobProgress -Job $task -Index $TaskObject.Index -Completed $true
             }
                if($TaskObject.Status -eq "Error"){
                WriteJobProgress -Job $task -Index $TaskObject.Index -Completed $true
             }
        $i++
    }


if((Get-task| Where-Object {($_.State -eq "Error") -and $_.Name -Match "Deploy OVF template"}).count -gt 0){

Write-host "Errors Found in Deployment." -ForegroundColor Red


}

Start-Sleep -Seconds 10

$VMNames = @()
if($NoRSVMs -ne $True){
Write-Host "Configuring RayStation VMs" -ForegroundColor Green
foreach($VMJson in $VMs){

$VM = get-vm $VMJson.hostname

$VMNames += $VM.Name

set-vm -VM $VM -MemoryGB $VMJson.RAM -confirm:$False | out-null
$spec = New-Object VMware.Vim.VirtualMachineConfigSpec
$spec.memoryReservationLockedToMax = $true
$vm.ExtensionData.ReconfigVM_Task($spec) | out-null


set-vm -vm $VM -NumCPU $VMJson.vCPU -confirm:$false | out-null


$vm | Get-HardDisk -Name 'Hard disk 1' | Set-HardDisk -CapacityGB $VMJson.DiskGB -Confirm:$false | out-null

if($VMJson.GPUType -match "VGPU"){

New-vGPU -VM $vm -vGPUProfile $VMJson.GPU | out-null

}

if($VMJson.GPUType -match "Passthrough"){

$GPUID = $VMJson.CardUUID

$PassthroughDevice = Get-PassthroughDevice -Type Pci -VMHost $VMJson.hostIP | Where-Object {(($_.uid.split('-'))[3] -replace '/','') -match $GPUID}

Add-PassthroughDevice -VM $VM -PassthroughDevice $PassthroughDevice | out-null

}

New-AdvancedSetting $VM -Name "pciPassthru.64bitMMIOSizeGB" -Value "128" -Confirm:$False | Out-Null
New-AdvancedSetting $VM -Name "pciPassthru.use64bitMMIO" -Value "TRUE" -Confirm:$False | Out-Null
}

Write-Host "Finished Configuring RayStation VMs" -ForegroundColor Green


foreach($VMJson in $VMs){

$VM = get-vm $VMJson.hostname

Start-VM -VM $VM -Confirm:$false -RunAsync | out-null

}

Write-Host "Finished Powering on RayStation VMs" -ForegroundColor Green
}


if($NoNvidiaVMs -ne $True){
write-Host "Configuring Nvidia VMs" -ForegroundColor Green
foreach($NvidiaVM in $NvidiaVMs){

    $VM = get-vm $NvidiaVM.VMName
    
    Start-VM -VM $VM -Confirm:$false -RunAsync | out-null
    
}
Write-Host "Finished Powering on Nvidia VMs" -ForegroundColor Green
}

if($NoSQLVMs -ne $True){
    write-host "Configuring SQL VMs" -ForegroundColor Green
    foreach($SQLVM in $SQLVMs){

        $VM = get-vm $SQLVM.hostname
        
        $VMNames += $VM.Name
        
        set-vm -VM $VM -MemoryGB $SQLVM.RAM -confirm:$False | out-null
        
        set-vm -vm $VM -NumCPU $SQLVM.vCPU -confirm:$false | out-null
        
        $vm | Get-HardDisk -Name 'Hard disk 1' | Set-HardDisk -CapacityGB $SQLVM.DiskGB -Confirm:$false | out-null

        if(![string]::IsNullOrEmpty($SQLVM.AdditionalDiskDatastore)){
            if(![string]::IsNullOrEmpty($SQLVM.DiskSize1)){
                New-HardDisk -VM $vm -CapacityGB $SQLVM.DiskSize1 -Datastore $SQLVM.AdditionalDiskDatastore -confirm:$false | out-null
        }
        if(![string]::IsNullOrEmpty($SQLVM.DiskSize2)){
            New-HardDisk -VM $vm -CapacityGB $SQLVM.DiskSize2 -Datastore $SQLVM.AdditionalDiskDatastore -confirm:$false | out-null
        }
        if(![string]::IsNullOrEmpty($SQLVM.DiskSize3)){
            New-HardDisk -VM $vm -CapacityGB $SQLVM.DiskSize3 -Datastore $SQLVM.AdditionalDiskDatastore -confirm:$false | out-null
        }
        if(![string]::IsNullOrEmpty($SQLVM.DiskSize4)){
            New-HardDisk -VM $vm -CapacityGB $SQLVM.DiskSize4 -Datastore $SQLVM.AdditionalDiskDatastore -confirm:$false | out-null
        }
        if(![string]::IsNullOrEmpty($SQLVM.DiskSize5)){
            New-HardDisk -VM $vm -CapacityGB $SQLVM.DiskSize5 -Datastore $SQLVM.AdditionalDiskDatastore -confirm:$false | out-null
        }
        if(![string]::IsNullOrEmpty($SQLVM.DiskSize6)){
            New-HardDisk -VM $vm -CapacityGB $SQLVM.DiskSize6 -Datastore $SQLVM.AdditionalDiskDatastore -confirm:$false | out-null
        }
    }
}

foreach($SQLVM in $SQLVMs){

    $VM = get-vm $SQLVM.hostname
    
    Start-VM -VM $VM -Confirm:$false -RunAsync | out-null
    
    }
    Write-Host "Finished Powering on SQL VMs" -ForegroundColor Green

  }


$DeploymentTime.Stop()
$DeploymentTimeMin = $DeploymentTime.Elapsed.Minutes
$DeploymentTimeSec = $DeploymentTime.Elapsed.Seconds
Write-Host "Deployment Complete. Time to complete: $DeploymentTimeMin Minutes $DeploymentTimeSec Seconds" -ForegroundColor Green