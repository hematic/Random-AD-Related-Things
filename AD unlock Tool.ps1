[void][Reflection.Assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][Reflection.Assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][Reflection.Assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')

function Show-MainForm_psf{

	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$FormMain = New-Object 'System.Windows.Forms.Form'
	$GroupBoxMainOutput = New-Object 'System.Windows.Forms.GroupBox'
	$RichTextBoxMainOutput = New-Object 'System.Windows.Forms.RichTextBox'
	$GroupBoxMainUnlockAccount = New-Object 'System.Windows.Forms.GroupBox'
	$ButtonMainUnlockAccountUnlockAccount = New-Object 'System.Windows.Forms.Button'
	$ButtonMainUnlockAccountQueryLockStatus = New-Object 'System.Windows.Forms.Button'
	$GroupBoxMainUserName = New-Object 'System.Windows.Forms.GroupBox'
	$TextBoxMainUserName = New-Object 'System.Windows.Forms.TextBox'
	$MenuStripMain = New-Object 'System.Windows.Forms.MenuStrip'
	$ToolStripMenuItemMainFile = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItemMainFileAbout = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItemMainFileExit = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	####################################################################################################
	### Begin: Main form ###############################################################################
	####################################################################################################
	
	$FormMain_Load ={
        # Set global ErrorActionPreference to Stop to ensure that error handling correctly works when using implicit remoting for the ActiveDirectory PowerShell module
		$global:ErrorActionPreference = "Stop"
	}
	
	$FormMain_Shown ={
		Import-SSAActiveDirectoryModule
		if ($SSAActiveDirectoryModuleLoaded -eq $false)
		{
			[System.Windows.Forms.MessageBox]::Show("The ActiveDirectory PowerShell module could not be loaded. You can review the log for more details. Please restart the application to try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
		}
	}
	
	$ToolStripMenuItemMainFileAbout_Click =	{
		Start-Process -FilePath "http://supersysadmin.com/100/powershell-gui-script-to-unlock-an-active-directory-users-account/"
	}
	
	$ToolStripMenuItemMainFileExit_Click =	{
		$FormMain.Close()
	}
	
	$ButtonMainUnlockAccountQueryLockStatus_Click =	{
		Add-SSAOutput -OutputText "Checking if ActiveDirectory PowerShell module is loaded."
		if ($SSAActiveDirectoryModuleLoaded -eq $true){
			Add-SSAOutput -OutputText "ActiveDirectory PowerShell module is loaded."
			Get-SSAUserName
			if ($SSAUserName){
				Add-SSAOutput -OutputText "Checking if user '$SSAUserName' is currently locked."
				try{
					$QueryADUser = Get-ADUser -Identity $SSAUserName -Properties LockedOut,lockoutTime -ErrorAction Stop
					if ($QueryADUser.LockedOut -eq $true){
						$QueryADUserLockoutTime = [datetime]::FromFileTime($($QueryADUser.lockoutTime)).ToString('yyyy-MM-dd HH:mm:ss')
						Add-SSAOutput -OutputText "User '$SSAUserName' is currently locked (since $QueryADUserLockoutTime)."
					}
					else{
						Add-SSAOutput -OutputText "User '$SSAUserName' is currently not locked."
					}
				}
				catch [exception]{
					Add-SSAOutput -OutputText "$_"
				}
			}
			else{
				Add-SSAOutput -OutputText "UserName field is empty, please review your input."
			}
		}
		else{
			Add-SSAOutput -OutputText "ActiveDirectory PowerShell module is currently not loaded, cannot proceed with the request. Restart the application to attempt load the module."
		}
		
	}
	
	$ButtonMainUnlockAccountUnlockAccount_Click ={
		Add-SSAOutput -OutputText "Checking if ActiveDirectory PowerShell module is loaded."
		if ($SSAActiveDirectoryModuleLoaded -eq $true){
			Add-SSAOutput -OutputText "ActiveDirectory PowerShell module is loaded."
			Get-SSAUserName
			if ($SSAUserName){
				Add-SSAOutput -OutputText "Checking if user '$SSAUserName' is currently locked."
				try{
					$QueryADUser = Get-ADUser -Identity $SSAUserName -Properties LockedOut, lockoutTime -ErrorAction Stop
					if ($QueryADUser.LockedOut -eq $true){
						$QueryADUserLockoutTime = [datetime]::FromFileTime($($QueryADUser.lockoutTime)).ToString('yyyy-MM-dd HH:mm:ss')
						Add-SSAOutput -OutputText "User '$SSAUserName' is currently locked (since $QueryADUserLockoutTime)."
						Add-SSAOutput -OutputText "Attempting to unlock '$SSAUserName'."
						Unlock-ADAccount -Identity $SSAUserName -ErrorAction Stop
						$QueryADUser = Get-ADUser -Identity $SSAUserName -Properties LockedOut, lockoutTime -ErrorAction Stop
						if ($QueryADUser.LockedOut -eq $false){
							Add-SSAOutput -OutputText "User '$SSAUserName' is now unlocked."
						}
						else{
							Add-SSAOutput -OutputText "User '$SSAUserName' could not be unlocked."
						}
					}
					else
					{
						Add-SSAOutput -OutputText "User '$SSAUserName' is currently not locked."
					}
				}
				catch [exception]{
					Add-SSAOutput -OutputText "$_"
				}
			}
			else{
				Add-SSAOutput -OutputText "UserName field is empty, please review your input."
			}
		}
		else{
			Add-SSAOutput -OutputText "ActiveDirectory PowerShell module is currently not loaded, cannot proceed with the request. Restart the application to attempt load the module."
		}
	}
	
	$RichTextBoxMainOutput_TextChanged ={
		$RichTextBoxMainOutput.SelectionStart = $RichTextBoxMainOutput.Text.Length
		$RichTextBoxMainOutput.ScrollToCaret()
	}
	
	####################################################################################################
	### End: Main form #################################################################################
	####################################################################################################
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$FormMain.WindowState = $InitialFormWindowState
	}
	
	$Form_StoreValues_Closing={
		#Store the control values
		$script:MainForm_RichTextBoxMainOutput = $RichTextBoxMainOutput.Text
		$script:MainForm_TextBoxMainUserName = $TextBoxMainUserName.Text
	}
	
	$Form_Cleanup_FormClosed={
		#Remove all event handlers from the controls
		try	{
			$RichTextBoxMainOutput.remove_TextChanged($RichTextBoxMainOutput_TextChanged)
			$ButtonMainUnlockAccountUnlockAccount.remove_Click($ButtonMainUnlockAccountUnlockAccount_Click)
			$ButtonMainUnlockAccountQueryLockStatus.remove_Click($ButtonMainUnlockAccountQueryLockStatus_Click)
			$FormMain.remove_Load($FormMain_Load)
			$FormMain.remove_Shown($FormMain_Shown)
			$ToolStripMenuItemMainFileAbout.remove_Click($ToolStripMenuItemMainFileAbout_Click)
			$ToolStripMenuItemMainFileExit.remove_Click($ToolStripMenuItemMainFileExit_Click)
			$FormMain.remove_Load($Form_StateCorrection_Load)
			$FormMain.remove_Closing($Form_StoreValues_Closing)
			$FormMain.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$FormMain.SuspendLayout()
	$GroupBoxMainOutput.SuspendLayout()
	$GroupBoxMainUnlockAccount.SuspendLayout()
	$GroupBoxMainUserName.SuspendLayout()
	$MenuStripMain.SuspendLayout()
	#
	# FormMain
	#
	$FormMain.Controls.Add($GroupBoxMainOutput)
	$FormMain.Controls.Add($GroupBoxMainUnlockAccount)
	$FormMain.Controls.Add($GroupBoxMainUserName)
	$FormMain.Controls.Add($MenuStripMain)
	$FormMain.AutoScaleDimensions = '6, 13'
	$FormMain.AutoScaleMode = 'Font'
	$FormMain.ClientSize = '633, 298'
	$FormMain.MainMenuStrip = $MenuStripMain
	$FormMain.MinimumSize = '649, 337'
	$FormMain.Name = 'FormMain'
	$FormMain.StartPosition = 'CenterScreen'
	$FormMain.Text = 'ADUserUnlock v1.0'
	$FormMain.add_Load($FormMain_Load)
	$FormMain.add_Shown($FormMain_Shown)
	#
	# GroupBoxMainOutput
	#
	$GroupBoxMainOutput.Controls.Add($RichTextBoxMainOutput)
	$GroupBoxMainOutput.Anchor = 'Top, Bottom, Left, Right'
	$GroupBoxMainOutput.Location = '13, 89'
	$GroupBoxMainOutput.Name = 'GroupBoxMainOutput'
	$GroupBoxMainOutput.Size = '609, 199'
	$GroupBoxMainOutput.TabIndex = 3
	$GroupBoxMainOutput.TabStop = $False
	$GroupBoxMainOutput.Text = 'Output'
	#
	# RichTextBoxMainOutput
	#
	$RichTextBoxMainOutput.Anchor = 'Top, Bottom, Left, Right'
	$RichTextBoxMainOutput.Font = 'Consolas, 8.25pt'
	$RichTextBoxMainOutput.Location = '7, 19'
	$RichTextBoxMainOutput.Name = 'RichTextBoxMainOutput'
	$RichTextBoxMainOutput.ScrollBars = 'ForcedVertical'
	$RichTextBoxMainOutput.Size = '596, 174'
	$RichTextBoxMainOutput.TabIndex = 0
	$RichTextBoxMainOutput.Text = ''
	$RichTextBoxMainOutput.add_TextChanged($RichTextBoxMainOutput_TextChanged)
	#
	# GroupBoxMainUnlockAccount
	#
	$GroupBoxMainUnlockAccount.Controls.Add($ButtonMainUnlockAccountUnlockAccount)
	$GroupBoxMainUnlockAccount.Controls.Add($ButtonMainUnlockAccountQueryLockStatus)
	$GroupBoxMainUnlockAccount.Location = '320, 28'
	$GroupBoxMainUnlockAccount.Name = 'GroupBoxMainUnlockAccount'
	$GroupBoxMainUnlockAccount.Size = '302, 54'
	$GroupBoxMainUnlockAccount.TabIndex = 2
	$GroupBoxMainUnlockAccount.TabStop = $False
	$GroupBoxMainUnlockAccount.Text = 'Unlock Account'
	#
	# ButtonMainUnlockAccountUnlockAccount
	#
	$ButtonMainUnlockAccountUnlockAccount.Location = '154, 20'
	$ButtonMainUnlockAccountUnlockAccount.Name = 'ButtonMainUnlockAccountUnlockAccount'
	$ButtonMainUnlockAccountUnlockAccount.Size = '142, 23'
	$ButtonMainUnlockAccountUnlockAccount.TabIndex = 1
	$ButtonMainUnlockAccountUnlockAccount.Text = 'Unlock Account'
	$ButtonMainUnlockAccountUnlockAccount.UseVisualStyleBackColor = $True
	$ButtonMainUnlockAccountUnlockAccount.add_Click($ButtonMainUnlockAccountUnlockAccount_Click)
	#
	# ButtonMainUnlockAccountQueryLockStatus
	#
	$ButtonMainUnlockAccountQueryLockStatus.Location = '7, 20'
	$ButtonMainUnlockAccountQueryLockStatus.Name = 'ButtonMainUnlockAccountQueryLockStatus'
	$ButtonMainUnlockAccountQueryLockStatus.Size = '141, 23'
	$ButtonMainUnlockAccountQueryLockStatus.TabIndex = 0
	$ButtonMainUnlockAccountQueryLockStatus.Text = 'Query Lock Status'
	$ButtonMainUnlockAccountQueryLockStatus.UseVisualStyleBackColor = $True
	$ButtonMainUnlockAccountQueryLockStatus.add_Click($ButtonMainUnlockAccountQueryLockStatus_Click)
	#
	# GroupBoxMainUserName
	#
	$GroupBoxMainUserName.Controls.Add($TextBoxMainUserName)
	$GroupBoxMainUserName.Location = '13, 28'
	$GroupBoxMainUserName.Name = 'GroupBoxMainUserName'
	$GroupBoxMainUserName.Size = '300, 54'
	$GroupBoxMainUserName.TabIndex = 1
	$GroupBoxMainUserName.TabStop = $False
	$GroupBoxMainUserName.Text = 'UserName (SamAccountName)'
	#
	# TextBoxMainUserName
	#
	$TextBoxMainUserName.Font = 'Consolas, 8.25pt'
	$TextBoxMainUserName.Location = '7, 20'
	$TextBoxMainUserName.Name = 'TextBoxMainUserName'
	$TextBoxMainUserName.Size = '287, 20'
	$TextBoxMainUserName.TabIndex = 0
	#
	# MenuStripMain
	#
	[void]$MenuStripMain.Items.Add($ToolStripMenuItemMainFile)
	$MenuStripMain.Location = '0, 0'
	$MenuStripMain.Name = 'MenuStripMain'
	$MenuStripMain.Size = '633, 24'
	$MenuStripMain.TabIndex = 0
	$MenuStripMain.Text = 'MenuStripMain'
	#
	# ToolStripMenuItemMainFile
	#
	[void]$ToolStripMenuItemMainFile.DropDownItems.Add($ToolStripMenuItemMainFileAbout)
	[void]$ToolStripMenuItemMainFile.DropDownItems.Add($ToolStripMenuItemMainFileExit)
	$ToolStripMenuItemMainFile.Name = 'ToolStripMenuItemMainFile'
	$ToolStripMenuItemMainFile.Size = '37, 20'
	$ToolStripMenuItemMainFile.Text = 'File'
	#
	# ToolStripMenuItemMainFileAbout
	#
	$ToolStripMenuItemMainFileAbout.Name = 'ToolStripMenuItemMainFileAbout'
	$ToolStripMenuItemMainFileAbout.Size = '107, 22'
	$ToolStripMenuItemMainFileAbout.Text = 'About'
	$ToolStripMenuItemMainFileAbout.add_Click($ToolStripMenuItemMainFileAbout_Click)
	#
	# ToolStripMenuItemMainFileExit
	#
	$ToolStripMenuItemMainFileExit.Name = 'ToolStripMenuItemMainFileExit'
	$ToolStripMenuItemMainFileExit.Size = '152, 22'
	$ToolStripMenuItemMainFileExit.Text = 'Exit'
	$ToolStripMenuItemMainFileExit.add_Click($ToolStripMenuItemMainFileExit_Click)
	$MenuStripMain.ResumeLayout()
	$GroupBoxMainUserName.ResumeLayout()
	$GroupBoxMainUnlockAccount.ResumeLayout()
	$GroupBoxMainOutput.ResumeLayout()
	$FormMain.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $FormMain.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$FormMain.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$FormMain.add_FormClosed($Form_Cleanup_FormClosed)
	#Store the control values when form is closing
	$FormMain.add_Closing($Form_StoreValues_Closing)
	#Show the Form
	return $FormMain.ShowDialog()

}

function Add-SSAOutput{
    [CmdletBinding()]
	Param
	(
	    [Parameter(Mandatory = $true)]$OutputText
	)
		
	$RichTextBoxMainOutput.Text += "$OutputText`n"
}

function Import-SSAActiveDirectoryModule{
    Add-SSAOutput -OutputText "Loading ActiveDirectory PowerShell module..."
    if ((Get-Module -name "ActiveDirectory") -eq $null){
	    Add-SSAOutput -OutputText "ActiveDirectory PowerShell module is currently not loaded."
		if (Get-Module -ListAvailable | Where-Object { $_.name -eq "ActiveDirectory" }){
		    Add-SSAOutput -OutputText "ActiveDirectory PowerShell module is available, importing module."
			Import-Module -Name "ActiveDirectory"
			if ((Get-Module -name "ActiveDirectory") -eq $null){
			    Add-SSAOutput -OutputText "ActiveDirectory PowerShell module could not be loaded."
				$global:SSAActiveDirectoryModuleLoaded = $false
			}
			else{
				Add-SSAOutput -OutputText "ActiveDirectory PowerShell module has been loaded."
				$global:SSAActiveDirectoryModuleLoaded = $true
			}
		}
		else{
		    Add-SSAOutput -OutputText "ActiveDirectory PowerShell module is not available on this computer, attempting to import it from a Domain Controller."
			try{
				$DomainDNSName = (Get-WmiObject -Class WIN32_ComputerSystem -ErrorAction Stop).Domain
				$DomainNetBiosName = (Get-WmiObject Win32_NTDomain -Filter "DnsForestName = '$((Get-WmiObject Win32_ComputerSystem).Domain)'" -ErrorAction Stop).DomainName
				$DomainControllerName = ((Get-WmiObject -Class WIN32_NTDomain -Filter "DomainName = '$DomainNetBiosName'" -ErrorAction Stop).DomainControllerName) -replace "\\", ""
				$DomainController = "$DomainControllerName.$DomainDNSName"
				$DomainControllerSession = New-PSSession -Computername $DomainController -ErrorAction Stop
				Invoke-Command -Command { Import-Module -Name "ActiveDirectory" } -Session $DomainControllerSession -ErrorAction Stop
				$ImportSession = Import-PSSession -Session $DomainControllerSession -Module ActiveDirectory -AllowClobber -ErrorAction Stop
				if ($ImportSession.Name){
					Add-SSAOutput -OutputText "ActiveDirectory PowerShell module has been imported from Domain Controller $DomainController."
					$global:SSAActiveDirectoryModuleLoaded = $true
				}
			}
			catch{
				Add-SSAOutput -OutputText "ActiveDirectory PowerShell module could not be imported. Possible reasons are: This workstation is not joined to the Active Directory domain, PowerShell remoting towards the Domain Controller does not work or is not setup, the current user does not have appropriate rights to open a PowerShell session to the Domain Controller."
				$global:SSAActiveDirectoryModuleLoaded = $false
			}
		}
	}
    else{
	    Add-SSAOutput -OutputText "ActiveDirectory PowerShell module is already loaded."
		$global:SSAActiveDirectoryModuleLoaded = $true
	}
}

function Get-SSAUserName{
    $global:SSAUserName = $TextBoxMainUserName.text
}

Main ($CommandLine)