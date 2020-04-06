#
# Clean Script
#

function Get-ScriptDirectory
{
  try{
        $Invocation = (Get-Variable MyInvocation -Scope 1).Value
        Split-Path $Invocation.MyCommand.Path
    }catch{
        return "."
    }
}

Function LogLine($strLine)
{
	Write-Host $strLine
	$StrTime = Get-Date -Format "MM-dd-yyyy-HH-mm-ss-tt"
	"$StrTime - $strLine " | Out-file -FilePath $LogFile -Encoding ASCII -Append
}


#Script Setup
#===================================================================
$adminAddress = "xd1"
$domain = "citrix"
#note $adminaddress is ignored when using the cloud sdk
$CatalogNamePrefix = "MC_"
$DesktopGroupNamePrefix = "DG_"
$BranchNamePrefix = "Filiale "
$csvName = "input.txt"
$AddComputers = $true
$AddUsers = $true
$AddDeliveryGroup = $true
$AllowMultipleUsers = $false
$CheckResultsComputers = $true
$CheckResultsUsers = $true
$AddCatalogs= $true
#===================================================================

$ScriptSource = Get-ScriptDirectory
$ErrorActionPreference = 'stop'

#Create a log folder and file
$LogFolderName = Get-Date -Format "yyyyMMddHHmmss"
$LogTopFolder = "$ScriptSource\Logs"

If (!(Test-Path "$LogTopFolder"))
{
	mkdir "$LogTopFolder" >$null
}

$LogFolder = "$LogTopFolder\$LogFolderName"
mkdir "$LogFolder" >$null
$LogFile = "$LogFolder\RemotePC_Import_log.txt"

Logline "Running RemotePC Import Script"

$CsvFile = "$ScriptSource\$csvName"
if (Test-Path $CsvFile)
{
	Logline "Found CSV file will import"
	#Get Map csv file
	$MapUsers = Import-Csv -Path $CsvFile -Encoding ASCII
}


#Lets see if the Citrix Broker Admin snapin is loaded and if not load it.
$Snapins =  Get-PSSnapin

foreach ($Snapin in $Snapins) 
{
	If ($Snapin.Name -eq "Citrix.Broker.Admin.V2")
	{
		Logline "Snapin Citrix.Broker.Admin.V2 already loaded!"
		$SnapinLoaded = $True
		break
	}

}

if (!$SnapinLoaded)
{
	Logline "Loading Snapin Citrix.Broker.Admin.V2"
	asnp Citrix*
}

#Now lets make sure it loaded
$SnapinLoaded = $false
$Snapins2 =  Get-PSSnapin

foreach ($Snapin in $Snapins2) 
{
	If ($Snapin.Name -eq "Citrix.Broker.Admin.V2")
	{
		$SnapinLoaded = $True
		break
	}

}
if (!$SnapinLoaded)
{
		Logline "****Snapin [Citrix.Broker.Admin.V2] could not loaded - Exiting Script"
		Throw "Snapin Could not be loaded.  Exiting Script"
		break
}

$BranchList = ($MapUsers | Group-Object -Property CodFiliale,GruppoAD,Descrizione)

#
# FOR EACH BRANCH CLEANUP
#

foreach($Branch in $BranchList)
{
    $BranchCode = $Branch.Name.ToString().Split(",")[0].Trim()
    if($BranchCode.Length -eq 0)
    {
        Logline "ERROR: Branch Code is empty. Continue with next branch."
        continue
    }

    Logline "**** Starting Cleanup Action for Branch [$BranchCode] ****"

    $CatalogName = $CatalogNamePrefix+$BranchCode
    $DeliveryGroupName = $DesktopGroupNamePrefix+$BranchCode

    $MachineCatalog = Get-BrokerCatalog -AdminAddress $adminAddress -Name $CatalogName -ErrorAction SilentlyContinue
    if($MachineCatalog -ne $null)
    {
        #Get the list of machines and put them in MaintenaceMode
        $listOfMachines = Get-BrokerMachine -AdminAddress $adminAddress -CatalogName $CatalogName -ErrorAction SilentlyContinue
        if($listOfMachines -ne $null)
        {
            $listOfMachines | Set-BrokerMachine -InMaintenanceMode $true -AdminAddress $adminAddress
            Logline "Machines for Catalog [$CatalogName] set to MaintenanceMode ON"
        }
        else
        {
            Logline "No machines found for Catalog [$CatalogName]"
        }
    }

    $DesktopGroup = Get-BrokerDesktopGroup -AdminAddress $adminAddress -Name $DeliveryGroupName -ErrorAction SilentlyContinue
    if($DesktopGroup -ne $null)
    {
        #Delete Delivery Group
        $DesktopGroup | Remove-BrokerDesktopGroup -AdminAddress $adminAddress
        Logline "Delivery Group [$DesktopGroup] deleted"
    }
    else
    {
        Logline "No Delivery Group found for branch [$BranchCode]"
    }

    if($MachineCatalog -ne $null)
    {
        #Delete Machine Catalog
        $MachineCatalog | Remove-BrokerCatalog -AdminAddress $adminAddress
        Logline "Machine Catalog [$MachineCatalog] deleted"
    }
    else
    {
        Logline "No Machine Catalog found for branch [$BranchCode]"
    }
}