Write-Host "Importing Active Directory Module" -ForegroundColor 'Green'
Import-Module -Name ActiveDirectory

#region HTML Output Formatting

	$a = "<style>"
	$a = $a + "BODY{background-color:Lavender ;}"
	$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
	$a = $a + "TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:thistle}"
	$a = $a + "TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:PaleGoldenrod}"
	$a = $a + "</style>"
#endregion
#region Setting Variables

	$date = (Get-Date -Format d_MMMM_yyyy).toString()
	$filePATH = "$env:userprofile\Desktop\"
	$fileNAME = "AD_Info_" + $date + ".html"
	$file = $filePATH + $fileNAME

    # Active Directory Variables
	$adFOREST = Get-ADForest
	$adDOMAIN = Get-ADDomain

	# Forest Variables
	$adFORESTNAME = $adFOREST.Name
	$adFORESTMODE = $adFOREST.ForestMode
	$adFORESTDOMAIN = $adFOREST | select -ExpandProperty Domains
	$adFORESTROOTDOMAIN = $adFOREST.RootDomain
	$adFORESTSchemaMaster = $adFOREST.SchemaMaster
	$adFORESTNamingMaster = $adFOREST.DomainNamingMaster
	$adFORESTUPNSUFFIX = $adFOREST | select -ExpandProperty UPNSuffixes 
	$adFORESTSPNSUffix = $adFOREST | select -ExpandProperty SPNSuffixes
	$adFORESTGlobalCatalog = $adFOREST | select -ExpandProperty GlobalCatalogs
	$adFORESTSites = $adFOREST  |  select -ExpandProperty Sites
	

    #Domain Vaiables
	$adDomainName = $adDOMAIN.Name
	$adDOMAINNetBiosName = $adDOMAIN.NetBIOSName
	$adDOMAINDomainMode = $adDOMAIN.DomainMode
	$adDOMAINParentDomain = $adDOMAIN.ParentDomain
	$adDOMAINPDCEMu = $adDOMAIN.PDCEmulator
	$adDOMAINRIDMaster = $adDOMAIN.RIDMaster
	$adDOMAINInfra = $adDOMAIN.InfrastructureMaster
	$adDOMAINChildDomain = $adDOMAIN | select -ExpandProperty ChildDomains
	$adDOMAINReplica = $adDOMAIN | select -ExpandProperty ReplicaDirectoryServers
	$adDOMAINReadOnlyReplica = $adDOMAIN | select -ExpandProperty ReadOnlyReplicaDirectoryServers
#endregion	
#region delete old results file
if (Test-Path "$env:userprofile\Desktop\$filename" ) { 
	    "`n"
	    Write-Warning "file already exists, i am deleting it."
	    Remove-Item "$env:userprofile\Desktop\$filename" -Verbose -Force
	    "`n"
	    Write-Host "Creating a New file Named as $fileNAME" -ForegroundColor 'Green'
	    New-Item -Path $filePATH -Name $fileNAME -Type file | Out-Null
	} 
else {
	    "`n"
	    Write-Host "Creating a New file Named as $fileNAME" -ForegroundColor 'Green'
	    New-Item -Path $filePATH -Name $fileNAME -Type file | Out-Null
	    "`n"
	}
#endregion
#region HTML Output

ConvertTo-Html  -Head $a  -Title "ACtive Directory Information" -Body "<h1> Active Directory Information for :  $adFORESTNAME </h1>" > $file

ConvertTo-Html  -Head $a -Body "<h2> Active Directory Forest Information. </h2>"  >> $file 

ConvertTo-Html -Body "<table><tr><td> Forest Name: </td><td><b> $adFORESTNAME </b></td></tr> `
					  <tr><td> Forest Mode: </td><td><b> $adFORESTMODE </b></td></tr> `
					  <tr><td> Forest Domains: </td><td><b> $adFORESTDOMAIN </b></td></tr> `
					  <tr><td> Root Domain : </td><td><b> $adFORESTROOTDOMAIN </b></td></tr> `	
					  <tr><td> Domain Naming Master: </td><td><b> $adFORESTNamingMaster </b></td></tr> `	
					  <tr><td> Schema Master: </td><td><b> $adFORESTSchemaMaster </b></td></tr> `	
			 		  <tr><td> Domain SPNSuffixes : </td><td><b> $adFORESTSPNSUffix </b></td></tr> `
					  <tr><td> Domain UPNSuffixes : </td><td><b> $adFORESTUPNSUFFI </b></td></tr> `	
					  <tr><td> Global Catalog Servers : </td><td><b> $adFORESTGlobalCatalog </b></td></tr> `
					  <tr><td> Forest Domain Sites : </td><td><b> $adFORESTSites </b></td></tr></table>" >> $file 

ConvertTo-Html  -Head $a -Body "<h2> Active Directory Domain Information. </h2>"  >> $file 						
		
ConvertTo-Html -Body "<table><tr><td> Domain Name: </td><td><b> $adDomainName </b></td></tr> `
					  <tr><td> Domain NetBios Name: </td><td><b> $adDOMAINNetBiosName </b></td></tr> `
					  <tr><td> Domain Mode: </td><td><b> $adDOMAINDomainMode </b></td></tr> `
					  <tr><td> Parent Domain : </td><td><b> $adDOMAINParentDomain </b></td></tr> `	
					  <tr><td> Domain PDC Emulator : </td><td><b> $adDOMAINPDCEMu </b></td></tr> `	
					  <tr><td> Domain RID Master: </td><td><b> $adDOMAINRIDMaster </b></td></tr> `	
			 		  <tr><td> Domain InfraStructure Master : </td><td><b> $adDOMAINInfra </b></td></tr> `
					  <tr><td> Child Domains : </td><td><b> $adDOMAINChildDomain </b></td></tr> `	
					  <tr><td> Replicated Servers : </td><td><b> $adDOMAINReplica</b></td></tr> `
					  <tr><td> Read Only Replicated Server : </td><td><b> $adDOMAINReadOnlyReplica </b></td></tr></table>" >> $file 

$Report = "The Report is generated On  $(get-date) by $((Get-Item env:\username).Value) on computer $((Get-Item env:\Computername).Value)"
$Report  >> $file 

	
Invoke-Expression $file

#endregion