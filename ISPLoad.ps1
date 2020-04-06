#
# Load Script
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
# ADD NEW CATALOGS FOR EVERY BRANCHES
#
if($AddCatalogs)
{
    foreach($Branch in $BranchList)
    {
        Try {
                $BranchCode = $Branch.Name.ToString().Split(",")[0].Trim()
                if($BranchCode.Length -eq 0)
                {}

                $BranchGroupAD = $Branch.Name.ToString().Split(",")[1].Trim()
                if($BranchGroupAD.Lenght -eq 0)
                {}

                $BranchDescription = $Branch.Name.ToString().Split(",")[2].Trim()
                if($BranchDescription.Length -eq 0)
                {}

                $CatalogName = $CatalogNamePrefix+$BranchCode

                #Create catalogs if
                $RemotePCCatalog = Get-BrokerCatalog -Name $CatalogName -ErrorAction:SilentlyContinue
                if($RemotePCCatalog -eq $null)
                {
                    Logline "Creating Machine Catalog [$CatalogName]"
	                $RemotePCCatalog = New-BrokerCatalog -AdminAddress $adminAddress -AllocationType "Permanent" -IsRemotePC $False -MachinesArePhysical $True -MinimumFunctionalLevel "L7_9" -Name $CatalogName -PersistUserChanges "OnLocal" -ProvisioningType "Manual" -SessionSupport "SingleSession"
                }
                else
                {
                    Logline "Machine Catalog [$CatalogName] already exists."
                }

            }

        Catch {
	        Logline "Catalog could not be obtained.  Exiting script"
	        Throw "Catalog Could not be loaded.  Exiting Script"
	        break
        }
    }
      
}


#
# ADD MACHINE TO CATALOGS
#
if ($AddComputers)
{
	#First Loop through the list and add the computers to the catalog
	Logline "================================================================"
	logline "             Adding Computers to Catalog"
	Logline ""

	foreach ($UserMapping in $MapUsers)
	{

		$Machine = $UserMapping.NomeMacchina.Trim()
        $UserName = $UserMapping.Utente.Trim()
        $Branch = $UserMapping.CodFiliale.Trim()
        
        if ($Machine.length -eq 0)
		{
			Logline "**** Machine Name for User [$UserName] is null skipping to next machine"
			continue
		}

		Try {
			Logline "Adding Machine [$machine] to Catalog [$CatalogName]"
            $RemotePCCatalog = Get-BrokerCatalog -Name ($CatalogNamePrefix+$Branch)
			New-BrokerMachine -MachineName $Machine -CatalogUid $RemotePCCatalog.Uid -AdminAddress $adminAddress
        }
		Catch {
			$ErrorValue = $error[0] 
            if ($ErrorValue -like '*Machine is already allocated')
            {
                Logline "Machine [$machine] has already been added to Catalog [$CatalogName]."
            }
            else
            {
                Logline "=========================================================="
                Logline "**** Adding Machine [$machine] to Catalog [$CatalogName] FAILED"
				Logline $Error[0]
				Logline "=========================================================="
            }
		}
	}

	#Start-Sleep 60
}

# ADD NEW DELIVERY GROUP
#
if ($AddDeliveryGroup)
{
    #group the branches 
    $BranchList = ($MapUsers | Group-Object -Property CodFiliale,GruppoAD,Descrizione)
    
    foreach($Branch in $BranchList)
    {
        $BranchCode = $Branch.Name.ToString().Split(",")[0].Trim()
        if($BranchCode.Length -eq 0)
        {}

        $BranchGroupAD = $Branch.Name.ToString().Split(",")[1].Trim()
        if($BranchGroupAD.Lenght -eq 0)
        {}

        $BranchDescription = $Branch.Name.ToString().Split(",")[2].Trim()
        if($BranchDescription.Length -eq 0)
        {}

        #IF NOT EXISTS CREATE DELIVERY GROUP
        $DeliveryGroupName = ($DesktopGroupNamePrefix+$BranchCode)
        $DeliveryGroup = Get-BrokerDesktopGroup -Name $DeliveryGroupName -ErrorAction SilentlyContinue
        if( $DeliveryGroup -eq $null ){
            LogLine "Creating new Delivery group : $DeliveryGroupName"
            $DeliveryGroup = New-BrokerDesktopGroup -Name $DeliveryGroupName -DeliveryType DesktopsOnly -DesktopKind Private -MachineLogOnType ActiveDirectory -PublishedName ($BranchCode + "_" + $BranchDescription) -SessionSupport SingleSession
        }
        
        #IF AVAILABLE MACHINES > 0 ADD TO DELIVERY GROUP
        $MachineCatalog = Get-BrokerCatalog -Name ($CatalogNamePrefix+$BranchCode)
        if($MachineCatalog.AvailableCount -gt 0){
            LogLine "Adding catalog $($MachineCatalog.Name) to $DeliveryGroupName"
            Add-BrokerMachinesToDesktopGroup -Catalog $MachineCatalog -DesktopGroup $DeliveryGroup -Count $MachineCatalog.AvailableCount -AdminAddress $adminAddress
        }

        #CREATE ASSIGNMENT POLICY RULE  -> desktop assignment rule da Studio
        $fullBranchGroup = $domain+"\"+$BranchGroupAD
        $PolicyRule = Get-BrokerAssignmentPolicyRule -Name $BranchCode -AdminAddress $adminAddress -ErrorAction SilentlyContinue
        if($PolicyRule -eq $null){
            Logline "New Assignment Policy Rule : for $DeliveryGroupName"
            New-BrokerAssignmentPolicyRule -AdminAddress $adminAddress -DesktopGroupUid $DeliveryGroup.Uid -Description "KEYWORDS:Auto internetMFA LAN" -Name $BranchCode  -PublishedName ($BranchCode+"_"+$BranchDescription) -IncludedUserFilterEnabled $false
        }

        #CREATE ACCESS POLICY RULE -> da Users da Studio
        $PolicyAccessRule = Get-BrokerAccessPolicyRule -AdminAddress $adminAddress -Name $BranchCode -ErrorAction SilentlyContinue
        if( $PolicyAccessRule -eq $null ) {
            Logline "New Access Policy Rule : for $DeliveryGroupName"
            try{
                New-BrokerAccessPolicyRule -AdminAddress $adminAddress -Name ($BranchCode+"_Direct") -IncludedUsers $fullBranchGroup -DesktopGroupUid $DeliveryGroup.Uid -IncludedUserFilterEnabled $true -IncludedSmartAccessFilterEnabled $true -AllowedConnections NotViaAG
                New-BrokerAccessPolicyRule -AdminAddress $adminAddress -Name ($BranchCode+"_ViaAG") -IncludedUsers $fullBranchGroup -DesktopGroupUid $DeliveryGroup.Uid -IncludedUserFilterEnabled $true -IncludedSmartAccessFilterEnabled $true -AllowedConnections ViaAG
            }catch{
                LogLine "Error on AccessPolicyRule $BranchCode for $DeliveryGroupName"
            }
             
        }
       
    }
}

if ($AddUsers)
{
	#Now Loop through the list again and assign the users to the computers
	Logline "================================================================"
	logline "             Assigning Users to Computers"
	Logline ""

	foreach ($UserMapping in $MapUsers)
	{
		$Machine = $UserMapping.NomeMacchina.Trim()
        $fullMachineName = $domain+"\"+$Machine
		$UserName = $UserMapping.Utente.Trim()
        $fullUserName=$domain +"\"+$UserName
        $Branch = $UserMapping.CodFiliale
        $CatalogName = $CatalogNamePrefix + (Get-Date -Format "yyyyMMdd")

		if ($Machine.length -eq 0)
		{
			Logline "+++++ Desktop Name is blank for user [$UserName]. User will not be added"
			continue
		}
        if ($UserName.length -eq 0)
		{
			Logline "+++++ No User Defined for Desktop [$Machine]. User will not be added"
			continue
		}

		$GetDesktop = Get-BrokerMachine -MachineName $fullMachineName -AdminAddress $adminAddress -ErrorAction:SilentlyContinue
		if ($GetDesktop -isnot [Citrix.Broker.Admin.SDK.Machine])
		{
			Logline "**** Desktop [$machine] not found in catalog [$CatalogName]"
			Logline "**** Skipping to next user"
			continue
		}
		
		$GetAssignedUser = Get-BrokerUser -AdminAddress $adminAddress -MachineUid $GetDesktop.Uid
		if ($GetAssignedUser.Count -gt 1){$AssignedUserTest = $GetAssignedUser[0]}else{$AssignedUserTest = $GetAssignedUser}
		if ($AssignedUserTest -isnot [Citrix.Broker.Admin.SDK.User])
		{
			#We will assign the user
			Logline "Mapping user [$UserName] to Desktop [$Machine]"
			try {
				Add-BrokerUser -AdminAddress $adminAddress -Machine $fullMachineName -Name $fullUserName               
			}
			Catch {
				Logline "=========================================================="
				Logline "Error Adding User [$UserName] to Desktop [$Machine]"
				Logline $Error[0]
				Logline "=========================================================="
			}
			
		}
		elseif ($AllowMultipleUsers) ############DA VERIFICARE utilizzo del private desktop <##>
		{
			[System.Collections.ArrayList]$ArrAssUsers = @()
			foreach ($assUser in $GetAssignedUser)
			{
				$AssignedUser = $assUser.Name
				$ArrAssUsers.Add($AssignedUser)
			}
			
			if ( $ArrAssUsers -notcontains $UserName)
			{
				Logline "Adding additional User [$UserName] to Desktop [$Machine]"
				try {
					Add-BrokerUser -AdminAddress $adminAddress -PrivateDesktop $Machine -Name $fullUserName
				}
				Catch {
					Logline "=========================================================="
					Logline "Error Adding User [$UserName] to Desktop [$Machine]"
					Logline $Error[0]
					Logline "=========================================================="
				}
			}
		}
		else
		{
			$AssignedUser = $GetAssignedUser.Name
			Logline "+++ User already mapped to Desktop [$Machine] Mapped User [$UserName] Assigned User [$AssignedUser]"
		}
	 
	}
	#Start-Sleep 60
} # End AddUsers

#
