<# 
.Synopsis 
   The purpose of this tool is to give you an easy front end for the management of Lync Common Area
   Phones.  It allows you to create, remove, move, set pins, and policies which you'll never even
   need yet are avaialable via Get-CsCommonAreaPhone.  :)  Feel free to hack them out.
 
.DESCRIPTION 
   PowerShell GUI script which allows for GUI management of Lync Common Area Phones
 
.Notes 
     NAME:      lync_common_area_phone_tool.ps1
     VERSION:   1.1
     AUTHOR:    C. Anthony Caragol 
     LASTEDIT:  08/12/2014 
      
   V 1.0 - August 12 2014 - Initial release 
   V 1.1 - August 15 2014 - Removed Persistent Chat, Mobility, VoiceRoutingPolicy for 2010 compatibility
   V 1.2 - August 19 2014 - Added multiselect and pool filtering
	    
.Link 
   Website: http://www.lyncfix.com
   Twitter: http://www.twitter.com/canthonycaragol
   LinkedIn: http://www.linkedin.com/in/canthonycaragol
 
.EXAMPLE 
   .\lync_common_area_phone_tool.ps1 

.TODO


.APOLOGY
  First, an apology to Greig.  I didn't realize I was reinventing your wheel until after I had built this tool
  and a friend (Michael LaMontagne) let me know I didn't search hard enough.  
  https://greiginsydney.com/madcap-ps1-a-gui-for-lync-analog-devices-common-area-phones/

  Please excuse the sloppy coding for now, I don't use a development environment, IDE or ISE.  I use notepad, 
  not even Notepad++, just notepad.  I am not a developer, just an enthusiast so some code may be redundant or
  inefficient.
#>


Function ShutDownForm()
{
	$objForm.Close()
}

Function Add-QuickOU-Node($Nodes, $Path)
{
	$OUArray=$Path.Split(",")
	[array]::Reverse($OuArray)
	$SelectPath=""

	$OuArray | %{
	if ($SelectPath.length -eq 0) {$SelectPath=$_} else {$SelectPath = $_ + "," + $SelectPath} 
	$FindIt = $Nodes.Find($_, $False)
	If ($FindIt.Count -eq 1)
	{
		$Nodes = $FindIt[0].Nodes
	}
	Else
	{
		$Node = New-Object Windows.Forms.TreeNode($_)
		$Node.Name = $_
		$Node.Tag = $SelectPath
		
		[void]$Nodes.Add($Node)  
		$FindIt = $Nodes.Find($_, $False)
		$Nodes = $FindIt[0].Nodes     
	}
 	}
}

Function Show-QuickOu-Form()
{

	#For Windows 2008 Support
	Import-Module ActiveDirectory

	$SelectOUForm = New-Object Windows.Forms.Form
	$SelectOUForm.Size = New-Object System.Drawing.Size(515,580) 
	$SelectOuForm.FormBorderStyle = 'Fixed3D'
	$SelectOuForm.MaximizeBox = $false
	$SelectOuForm.Text = "Please Select an Organizational Unit"
	$SelectOuForm.Icon = $Global:LyncFixIcon

	$OUTreeView = New-Object Windows.Forms.TreeView
	$OUTreeView.PathSeparator = ","
	$OUTreeView.Size = New-Object System.Drawing.Size(500,500) 
	$SelectOUForm.Controls.Add($OUTreeView)

	$objIPProperties = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()
	$strDNSDomain = $objIPProperties.DomainName.toLower()
	$strDomainDN = $strDNSDomain.toString().split('.'); foreach ($strVal in $strDomainDN) {$strTemp += "dc=$strVal,"}; $strDomainDN = $strTemp.TrimEnd(",").toLower()
	$AllOUs= Get-ADObject -Filter 'ObjectClass -eq "organizationalUnit"' -SearchScope SubTree -SearchBase $strDomainDN

	ForEach ($OU in $AllOUs)
	{
		$MyOU=$OU.DistinguishedName
		Add-QuickOU-Node $OUTreeView.Nodes $MyOU
	}
	#Add Users CN, I stripped CNs out because of a specific client, feel free to add them back instead of using next lines
	$UsersCN="CN=Users,"+$MyOu.substring($MyOu.Indexof("DC="),$MyOu.length - ($MyOu.Indexof("DC=")))
	Add-QuickOU-Node $OUTreeView.Nodes $UsersCN

	$SelectOUButton = New-Object System.Windows.Forms.Button
	$SelectOUButton.Location = New-Object System.Drawing.Size(25,510)
	$SelectOUButton.Size = New-Object System.Drawing.Size(200,25)
	$SelectOUButton.Text = "Select"
	$SelectOUButton.Add_Click({
	$Global:AddCAPselectedOU=$OUTreeView.SelectedNode.tag
	$SelectOUForm.Close()
	})
	$SelectOUButton.Anchor = 'Bottom, Left'
	$SelectOUForm.Controls.Add($SelectOUButton)
	
	$CancelOUButton = New-Object System.Windows.Forms.Button
	$CancelOUButton.Location = New-Object System.Drawing.Size(275,510)
	$CancelOUButton.Size = New-Object System.Drawing.Size(200,25)
	$CancelOUButton.Text = "Cancel"
	$CancelOUButton.Add_Click({
	$SelectOUForm.Close()})
	$CancelOUButton.Anchor = 'Bottom, Left'
	$SelectOUForm.Controls.Add($CancelOUButton)

	$SelectOUForm.ShowDialog()
	$SelectOUForm.Dispose()
}


Function AddCommonAreaPhoneForm()
{
	$AppCAPForm = New-Object System.Windows.Forms.Form 
	$AppCAPForm.Text = "Create Common Area Phone"
	$AppCAPForm.Size = New-Object System.Drawing.Size(590,520) 
	$AppCAPForm.FormBorderStyle = 'Fixed3D'
	$AppCAPForm.StartPosition = "CenterScreen"
	$AppCAPForm.KeyPreview = $True
	$AppCAPForm.Icon = $Global:LyncFixIcon

	$AddCAP_OULabel = New-Object System.Windows.Forms.Label
	$AddCAP_OULabel.Location = New-Object System.Drawing.Size(20,50) 
	$AddCAP_OULabel.Size = New-Object System.Drawing.Size(120,20) 
	$AddCAP_OULabel.Text = "OU:"
	$AppCAPForm.Controls.Add($AddCAP_OULabel) 

	$AddCAP_OUText = New-Object System.Windows.Forms.Textbox
	$AddCAP_OUText.Location = new-object System.Drawing.Size(140,50) 
	$AddCAP_OUText.Size = New-Object System.Drawing.Size(340,20) 
	$AddCAP_OUText.Text = ""
	$AddCAP_OUText.Anchor = 'Top, Left, Right'
	$AppCAPForm.Controls.Add($AddCAP_OUText) 

	$AddCAP_BrowseButton = New-Object System.Windows.Forms.Button
	$AddCAP_BrowseButton.Location = New-Object System.Drawing.Size(480,50)
	$AddCAP_BrowseButton.Size = New-Object System.Drawing.Size(60,20)
	$AddCAP_BrowseButton.Text = "Browse"
	$AddCAP_BrowseButton.Add_Click({
	$Global:AddCAPselectedOU=""
	Show-QuickOu-Form
	$AddCAP_OUText.Text=$Global:AddCAPselectedOU
	})
	$AddCAP_BrowseButton.Anchor = 'Top, Left, Right'
	$AppCAPForm.Controls.Add($AddCAP_BrowseButton)

	$AddCAP_PoolLabel = New-Object System.Windows.Forms.Label
	$AddCAP_PoolLabel.Location = New-Object System.Drawing.Size(20,80) 
	$AddCAP_PoolLabel.Size = New-Object System.Drawing.Size(120,20) 
	$AddCAP_PoolLabel.Text = "Pool:"
	$AppCAPForm.Controls.Add($AddCAP_PoolLabel) 

	$AddCAP_PoolDropDown = new-object System.Windows.Forms.ComboBox
	$AddCAP_PoolDropDown.Location = New-Object System.Drawing.Size(140,80) 
	$AddCAP_PoolDropDown.Size = New-Object System.Drawing.Size(400,20) 
	$AddCAP_PoolDropDown.Anchor = 'Top, Left, Right'
	$AddCAP_PoolDropDown.add_SelectedIndexChanged({
		$Global:AddCAPPoolFieldChange=$true
	})
	$AppCAPForm.Controls.Add($AddCAP_PoolDropDown) 

	$AddCAP_EnabledLabel = New-Object System.Windows.Forms.Label
	$AddCAP_EnabledLabel.Location = New-Object System.Drawing.Size(20,110) 
	$AddCAP_EnabledLabel.Size = New-Object System.Drawing.Size(120,20) 
	$AddCAP_EnabledLabel.Text = "Enabled:"
	$AppCAPForm.Controls.Add($AddCAP_EnabledLabel) 

	$AddCAP_EnabledDropDown = new-object System.Windows.Forms.ComboBox
	$AddCAP_EnabledDropDown.Location = new-object System.Drawing.Size(140,110) 
	$AddCAP_EnabledDropDown.Size = New-Object System.Drawing.Size(400,20) 
	$AddCAP_EnabledDropDown.Text="True"
	$AddCAP_EnabledDropDown.add_SelectedIndexChanged({
	$Global:AddCAPEnabledFieldChange=$true
	}) 
	$AddCAP_EnabledDropDown.Anchor = 'Top, Left, Right'
	$AppCAPForm.Controls.Add($AddCAP_EnabledDropDown) 

	$AddCAP_SipAddressLabel = New-Object System.Windows.Forms.Label
	$AddCAP_SipAddressLabel.Location = New-Object System.Drawing.Size(20,140) 
	$AddCAP_SipAddressLabel.Size = New-Object System.Drawing.Size(120,20) 
	$AddCAP_SipAddressLabel.Text = "Sip Address:"
	$AppCAPForm.Controls.Add($AddCAP_SipAddressLabel) 

	$AddCAP_SipAddressText = New-Object System.Windows.Forms.Textbox
	$AddCAP_SipAddressText.Location = new-object System.Drawing.Size(140,140) 
	$AddCAP_SipAddressText.Size = New-Object System.Drawing.Size(400,20) 
	$AddCAP_SipAddressText.Text = "<Automatic>"
	$AddCAP_SipAddressText.Anchor = 'Top, Left, Right'
	$AddCAP_SipAddressText.add_TextChanged({
	$Global:AddCAPSipAddressFieldChange=$true
	})
	$AppCAPForm.Controls.Add($AddCAP_SipAddressText) 

	$AddCAP_DialPlanLabel = New-Object System.Windows.Forms.Label
	$AddCAP_DialPlanLabel.Location = New-Object System.Drawing.Size(20,170) 
	$AddCAP_DialPlanLabel.Size = New-Object System.Drawing.Size(120,20) 
	$AddCAP_DialPlanLabel.Text = "DialPlan:"
	$AppCAPForm.Controls.Add($AddCAP_DialPlanLabel) 

	$AddCAP_DialPlanDropDown = new-object System.Windows.Forms.ComboBox
	$AddCAP_DialPlanDropDown.Location = new-object System.Drawing.Size(140,170) 
	$AddCAP_DialPlanDropDown.Size = New-Object System.Drawing.Size(400,20) 
	$AddCAP_DialPlanDropDown.Text="<Automatic>"
	$AddCAP_DialPlanDropDown.Anchor = 'Top, Left, Right'
	$AddCAP_DialPlanDropDown.add_SelectedIndexChanged({
	$Global:AddCAPDialPlanFieldChange=$true
	}) 
	$AppCAPForm.Controls.Add($AddCAP_DialPlanDropDown) 

	$AddCAP_ClientPolicyLabel = New-Object System.Windows.Forms.Label
	$AddCAP_ClientPolicyLabel.Location = New-Object System.Drawing.Size(20,200) 
	$AddCAP_ClientPolicyLabel.Size = New-Object System.Drawing.Size(120,20) 
	$AddCAP_ClientPolicyLabel.Text = "ClientPolicy:"
	$AppCAPForm.Controls.Add($AddCAP_ClientPolicyLabel) 

	$AddCAP_ClientPolicyDropDown = new-object System.Windows.Forms.ComboBox
	$AddCAP_ClientPolicyDropDown.Location = new-object System.Drawing.Size(140,200) 
	$AddCAP_ClientPolicyDropDown.Size = New-Object System.Drawing.Size(400,20) 
	$AddCAP_ClientPolicyDropDown.Anchor = 'Top, Left, Right'
	$AddCAP_ClientPolicyDropDown.Text="<Automatic>"
	$AddCAP_ClientPolicyDropDown.add_SelectedIndexChanged({
	$Global:AddCAPClientPolicyFieldChange=$true
	}) 
	$AppCAPForm.Controls.Add($AddCAP_ClientPolicyDropDown) 

	$AddCAP_PinPolicyLabel = New-Object System.Windows.Forms.Label
	$AddCAP_PinPolicyLabel.Location = New-Object System.Drawing.Size(20,230) 
	$AddCAP_PinPolicyLabel.Size = New-Object System.Drawing.Size(120,20) 
	$AddCAP_PinPolicyLabel.Text = "PinPolicy:"
	$AppCAPForm.Controls.Add($AddCAP_PinPolicyLabel) 

	$AddCAP_PinPolicyDropDown = new-object System.Windows.Forms.ComboBox
	$AddCAP_PinPolicyDropDown.Location = new-object System.Drawing.Size(140,230) 
	$AddCAP_PinPolicyDropDown.Size = New-Object System.Drawing.Size(400,20) 
	#$AddCAP_PinPolicyDropDown.add_SelectedIndexChanged($OnSelect_PinPolicyDropDown) 
	$AddCAP_PinPolicyDropDown.Text="<Automatic>"	
	$AddCAP_PinPolicyDropDown.Anchor = 'Top, Left, Right'
	$AddCAP_PinPolicyDropDown.add_SelectedIndexChanged({
	$Global:AddCAPPinPolicyFieldChange=$true
	}) 
	$AppCAPForm.Controls.Add($AddCAP_PinPolicyDropDown) 

	$AddCAP_VoicePolicyLabel = New-Object System.Windows.Forms.Label
	$AddCAP_VoicePolicyLabel.Location = New-Object System.Drawing.Size(20,260) 
	$AddCAP_VoicePolicyLabel.Size = New-Object System.Drawing.Size(120,20) 
	$AddCAP_VoicePolicyLabel.Text = "VoicePolicy:"
	$AppCAPForm.Controls.Add($AddCAP_VoicePolicyLabel) 

	$AddCAP_VoicePolicyDropDown = new-object System.Windows.Forms.ComboBox
	$AddCAP_VoicePolicyDropDown.Location = new-object System.Drawing.Size(140,260) 
	$AddCAP_VoicePolicyDropDown.Size = New-Object System.Drawing.Size(400,20) 
	#$AddCAP_VoicePolicyDropDown.add_SelectedIndexChanged($OnSelect_VoicePolicyDropDown) 
	$AddCAP_VoicePolicyDropDown.Text = "<Automatic>"
	$AddCAP_VoicePolicyDropDown.Anchor = 'Top, Left, Right'
	$AddCAP_VoicePolicyDropDown.add_SelectedIndexChanged({
	$Global:AddCAPVoicePolicyFieldChange=$True
	})
	$AppCAPForm.Controls.Add($AddCAP_VoicePolicyDropDown) 

	$AddCAP_ConferencingPolicyLabel = New-Object System.Windows.Forms.Label
	$AddCAP_ConferencingPolicyLabel.Location = New-Object System.Drawing.Size(20,290) 
	$AddCAP_ConferencingPolicyLabel.Size = New-Object System.Drawing.Size(120,20) 
	$AddCAP_ConferencingPolicyLabel.Text = "ConferencingPolicy:"
	$AppCAPForm.Controls.Add($AddCAP_ConferencingPolicyLabel) 

	$AddCAP_ConferencingPolicyDropDown = new-object System.Windows.Forms.ComboBox
	$AddCAP_ConferencingPolicyDropDown.Location = new-object System.Drawing.Size(140,290) 
	$AddCAP_ConferencingPolicyDropDown.Size = New-Object System.Drawing.Size(400,20) 
	$AddCAP_ConferencingPolicyDropDown.Text = "<Automatic>"
	$AddCAP_ConferencingPolicyDropDown.Anchor = 'Top, Left, Right'
	$AddCAP_ConferencingPolicyDropDown.add_SelectedIndexChanged({
	$Global:AddCAPConferencingPolicyFieldChange=$True
	})
	$AppCAPForm.Controls.Add($AddCAP_ConferencingPolicyDropDown) 

	$AddCAP_LineURILabel = New-Object System.Windows.Forms.Label
	$AddCAP_LineURILabel.Location = New-Object System.Drawing.Size(20,320) 
	$AddCAP_LineURILabel.Size = New-Object System.Drawing.Size(120,20) 
	$AddCAP_LineURILabel.Text = "LineURI:"
	$AppCAPForm.Controls.Add($AddCAP_LineURILabel) 

	$AddCAP_LineURIText = New-Object System.Windows.Forms.Textbox
	$AddCAP_LineURIText.Location = new-object System.Drawing.Size(140,320) 
	$AddCAP_LineURIText.Size = New-Object System.Drawing.Size(400,20) 
	$AddCAP_LineURIText.Text = ""
	$AddCAP_LineURIText.Anchor = 'Top, Left, Right'
	$AddCAP_LineURIText.add_TextChanged({
	$Global:AddCAPLineURIFieldChange=$true
	})
	$AppCAPForm.Controls.Add($AddCAP_LineURIText) 

	$AddCAP_DisplayNumberLabel = New-Object System.Windows.Forms.Label
	$AddCAP_DisplayNumberLabel.Location = New-Object System.Drawing.Size(20,350) 
	$AddCAP_DisplayNumberLabel.Size = New-Object System.Drawing.Size(120,20) 
	$AddCAP_DisplayNumberLabel.Text = "DisplayNumber:"
	$AppCAPForm.Controls.Add($AddCAP_DisplayNumberLabel) 

	$AddCAP_DisplayNumberText = New-Object System.Windows.Forms.Textbox
	$AddCAP_DisplayNumberText.Location = new-object System.Drawing.Size(140,350) 
	$AddCAP_DisplayNumberText.Size = New-Object System.Drawing.Size(400,20) 
	$AddCAP_DisplayNumberText.Text = ""
	$AddCAP_DisplayNumberText.Anchor = 'Top, Left, Right'
	$AddCAP_DisplayNumberText.add_TextChanged({
	$Global:AddCAPDisplayNumberFieldChange=$true
	})
	$AppCAPForm.Controls.Add($AddCAP_DisplayNumberText) 

	$AddCAP_DisplayNameLabel = New-Object System.Windows.Forms.Label
	$AddCAP_DisplayNameLabel.Location = New-Object System.Drawing.Size(20,380) 
	$AddCAP_DisplayNameLabel.Size = New-Object System.Drawing.Size(120,20) 
	$AddCAP_DisplayNameLabel.Text = "DisplayName:"
	$AppCAPForm.Controls.Add($AddCAP_DisplayNameLabel) 

	$AddCAP_DisplayNameText = New-Object System.Windows.Forms.Textbox
	$AddCAP_DisplayNameText.Location = new-object System.Drawing.Size(140,380) 
	$AddCAP_DisplayNameText.Size = New-Object System.Drawing.Size(400,20) 
	$AddCAP_DisplayNameText.Text = ""
	$AddCAP_DisplayNameText.Anchor = 'Top, Left, Right'
	$AddCAP_DisplayNameText.add_TextChanged({
	$Global:AddCAPDisplayNameFieldChange=$true
	})
	$AppCAPForm.Controls.Add($AddCAP_DisplayNameText) 

	$AddCAP_DescriptionLabel = New-Object System.Windows.Forms.Label
	$AddCAP_DescriptionLabel.Location = New-Object System.Drawing.Size(20,410) 
	$AddCAP_DescriptionLabel.Size = New-Object System.Drawing.Size(120,20) 
	$AddCAP_DescriptionLabel.Text = "Description:"
	$AppCAPForm.Controls.Add($AddCAP_DescriptionLabel) 

	$AddCAP_DescriptionText = New-Object System.Windows.Forms.Textbox
	$AddCAP_DescriptionText.Location = new-object System.Drawing.Size(140,410) 
	$AddCAP_DescriptionText.Size = New-Object System.Drawing.Size(400,20) 
	$AddCAP_DescriptionText.Text = ""
	$AddCAP_DescriptionText.Anchor = 'Top, Left, Right'
	$AddCAP_DescriptionText.add_TextChanged({
	$Global:AddCAPDescriptionFieldChange=$true
	})
	$AppCAPForm.Controls.Add($AddCAP_DescriptionText) 

	$AddCAP_PoolDropDown.Items.Clear()
	foreach ($x in $AddCAP_PoolArray) 
	{
		[void]$AddCAP_PoolDropDown.Items.Add($x)
	}
	
	[void]$AddCAP_EnabledDropDown.Items.Clear()
	[void]$AddCAP_EnabledDropDown.Items.Add("True")
	[void]$AddCAP_EnabledDropDown.Items.Add("False")

	$AddCAP_DialPlanDropDown.Items.Clear()
	[void]$AddCAP_DialPlanDropDown.Items.Add("<Automatic>")
	foreach ($x in $AddCAP_dialplanarray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$AddCAP_DialPlanDropDown.Items.Add($y)
	}

	$AddCAP_ClientPolicyDropDown.Items.Clear()
	[void]$AddCAP_ClientPolicyDropDown.Items.Add("<Automatic>")
	foreach ($x in $AddCAP_ClientPolicyarray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$AddCAP_ClientPolicyDropDown.Items.Add($y)
	}

	$AddCAP_PinPolicyDropDown.Items.Clear()
	[void]$AddCAP_PinPolicyDropDown.Items.Add("<Automatic>")
	foreach ($x in $AddCAP_PinPolicyarray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$AddCAP_PinPolicyDropDown.Items.Add($y)
	}
	$AddCAP_VoicePolicyDropDown.Items.Clear()
	[void]$AddCAP_VoicePolicyDropDown.Items.Add("<Automatic>")
	foreach ($x in $AddCAP_VoicePolicyarray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$AddCAP_VoicePolicyDropDown.Items.Add($y)
	}

	$AddCAP_ConferencingPolicyDropDown.Items.Clear()
	[void]$AddCAP_ConferencingPolicyDropDown.Items.Add("<Automatic>")
	foreach ($x in $AddCAP_ConfPolicyArray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$AddCAP_ConferencingPolicyDropDown.Items.Add($y)
	}

	$CreateCAPButton = New-Object System.Windows.Forms.Button
	$CreateCAPButton.Location = New-Object System.Drawing.Size(25,450)
	$CreateCAPButton.Size = New-Object System.Drawing.Size(250,25)
	$CreateCAPButton.Text = "Create"
	$CreateCAPButton.Add_Click({

		$CreateCAPButton.Text = "Creating..."
		$CreateCAPButton.Enabled=$False		
		$CancelCAPButton.Enabled=$False		

		#Let's perform some validation checks.
		$ReadyToCreate=0

		#Check OU
		if ($AddCAP_OUText.Text.length -lt 1) 
		{
			[Microsoft.VisualBasic.Interaction]::MsgBox("Cannot proceed, please fill in the OU box.",'OKOnly,Critical', "Not enough information!")
			$ReadyToCreate=1
		}

		#Check DisplayName
		if ($AddCAP_DisplayNameText.Text.length -lt 1) 
		{
			[Microsoft.VisualBasic.Interaction]::MsgBox("Cannot proceed, please fill in the Display Name.",'OKOnly,Critical', "Not enough information!")
			$ReadyToCreate=1
		}

		#Check Pool
		if ($AddCAP_PoolDropDown.Text.length -lt 1) 
		{
			[Microsoft.VisualBasic.Interaction]::MsgBox("Cannot proceed, please select a pool.",'OKOnly,Critical', "Not enough information!")
			$ReadyToCreate=1
		}

		$AddCAP_LineURIText.Text = $AddCAP_LineURIText.Text -replace "tel:",""
		if ($AddCAP_LineURIText.Text.length -lt 1)
		{
			[Microsoft.VisualBasic.Interaction]::MsgBox("Cannot proceed, no LineURI specified.",'OKOnly,Critical', "Not enough information!")
			$ReadyToCreate=1
		}
		$AddCAP_LineURIText.Text = "tel:" + $AddCAP_LineURIText.Text
	
		if ($ReadyToCreate -eq 0) 
		{
			New-CsCommonAreaPhone -LineUri $AddCAP_LineURIText.Text -RegistrarPool $AddCAP_PoolDropDown.Text -DisplayName $AddCAP_DisplayNameText.Text -OU $AddCAP_OUText.Text
	
			#Sleep up to 20 seconds to let pool catch up, noted in Enterprise pools, your milage may vary.
		
			for($i = 0; $i -lt 20; $i++) {
				$x = Get-CsCommonAreaPhone $AddCAP_DisplayNameText.Text -erroraction 'silentlycontinue'
				If ($x -eq $null)
				{
					write-host "Waiting for phone creation before applying policies"
					Start-Sleep -s 1
				}
				else
				{
					break	
				}
			}
	
			if ($AddCap_EnabledDropDown.Text="False") { Set-CsCommonAreaPhone $AddCAP_DisplayNameText.Text -Enabled:$false }
			if ($AddCap_SipAddressText.Text.contains("@")) { 
				#Just in case SIP is there, remove it because we're going to readd it to ensure it's there.
				$AddCap_SipAddressText.Text = $AddCap_SipAddressText.Text -replace "sip:",""
				$AddCap_SipAddressText.Text = "sip:" + $AddCap_SipAddressText.Text
				Set-CsCommonAreaPhone $AddCAP_DisplayNameText.Text -SipAddress $AddCap_SipAddressText.Text
			}
			if ($AddCap_DialPlanDropDown.Text -ne "<Automatic>") { Get-CsCommonAreaPhone $AddCAP_DisplayNameText.Text | Grant-CSDialPlan -Policy $AddCap_DialPlanDropDown.Text}
			if ($AddCap_ClientPolicyDropDown.Text -ne "<Automatic>") { Get-CsCommonAreaPhone $AddCAP_DisplayNameText.Text | Grant-CSClientPolicy -Policy $AddCap_ClientPolicyDropDown.Text}
			if ($AddCap_PinPolicyDropDown.Text -ne "<Automatic>") { Get-CsCommonAreaPhone $AddCAP_DisplayNameText.Text | Grant-CSPinPolicy -Policy $AddCap_PinPolicyDropDown.Text}
			if ($AddCap_VoicePolicyDropDown.Text -ne "<Automatic>") { Get-CsCommonAreaPhone $AddCAP_DisplayNameText.Text | Grant-CSVoicePolicy -Policy $AddCap_VoicePolicyDropDown.Text}
			if ($AddCap_ConferencingPolicyDropDown.Text -ne "<Automatic>") { Get-CsCommonAreaPhone $AddCAP_DisplayNameText.Text | Grant-CSConferencingPolicy -Policy $AddCap_ConferencingPolicyDropDown.Text}
			if ($AddCap_DisplayNumberText.Text -gt 0) { Set-CsCommonAreaPhone $AddCAP_DisplayNameText.Text -DisplayNumber $AddCap_DisplayNumberText.Text}
			if ($AddCap_DescriptionText.Text.length -gt 0) { Set-CsCommonAreaPhone $AddCAP_DisplayNameText.Text -Description $AddCap_DescriptionText.Text}

			$AppCAPForm.Close()
		}
		else
		{
			$CreateCAPButton.Text = "Create"
			$CreateCAPButton.Enabled=$True		
			$CancelCAPButton.Enabled=$True	
		}
	})
	$CreateCAPButton.Anchor = 'Bottom, Left'
	$AppCAPForm.Controls.Add($CreateCAPButton)
	
	$CancelCAPButton = New-Object System.Windows.Forms.Button
	$CancelCAPButton.Location = New-Object System.Drawing.Size(290,450)
	$CancelCAPButton.Size = New-Object System.Drawing.Size(250,25)
	$CancelCAPButton.Text = "Cancel"
	$CancelCAPButton.Add_Click({
	$AppCAPForm.Close()})
	$CancelCAPButton.Anchor = 'Bottom, Left'
	$AppCAPForm.Controls.Add($CancelCAPButton)

	$AppCAPForm.Add_Shown({$AppCAPForm.Activate()})

[void] $AppCAPForm.ShowDialog()

}

Function LoadPhones()
{

	if ($FilterDropDown.Text -eq "<All Pools>")
	{
		$allphones=Get-CSCommonAreaPhone
	}
	else
	{
		$allphones=Get-CSCommonAreaPhone | Where {$_.RegistrarPool -like $FilterDropDown.Text}
	}

	foreach ($phone in $allphones) 
	{
		[void] $objListBox.Items.Add($phone.displayname)
	}

	#########################################################
	#	Reset field colors, values and selections
	#########################################################

	$IdentityText.Text=""
	$IdentityText.ForeColor = [System.Drawing.Color]::Black
	$PoolDropDown.Text=""
	$PoolDropDown.ForeColor = [System.Drawing.Color]::Black
	$PoolDropDown.Items.Clear()
	$Global:PoolFieldChange=$False
	$SipAddressText.Text = ""
	$SipAddressText.ForeColor = [System.Drawing.Color]::Black
	$Global:SipAddressFieldChange=$False
	$EnabledDropDown.Text=""
	$EnabledDropDown.ForeColor = [System.Drawing.Color]::Black
	$EnabledDropDown.Items.Clear()
	$Global:EnabledFieldChange=$False
	$DialPLanDropDown.Text=""
	$DialPlanDropDown.ForeColor = [System.Drawing.Color]::Black
	$DialPlanDropDown.Items.Clear()
	$Global:DialPlanFieldChange=$False
	$ClientPolicyDropDown.Text=""
	$ClientPolicyDropDown.ForeColor = [System.Drawing.Color]::Black
	$ClientPolicyDropDown.Items.Clear()
	$Global:ClientPolicyFieldChange=$False
	$PinPolicyDropDown.Text=""
	$PinPolicyDropDown.ForeColor = [System.Drawing.Color]::Black
	$PinPolicyDropDown.Items.Clear()
	$Global:PinPolicyFieldChange=$False
	$VoicePolicyDropDown.Text=""
	$VoicePolicyDropDown.ForeColor = [System.Drawing.Color]::Black
	$VoicePolicyDropDown.Items.Clear()
	$Global:VoicePolicyFieldChange=$False
	$ConferencingPolicyDropDown.Text=""
	$ConferencingPolicyDropDown.ForeColor = [System.Drawing.Color]::Black
	$ConferencingPolicyDropDown.Items.Clear()
	$Global:ConferencingPolicyFieldChange=$False
	$LineURIText.Text = ""
	$LineURIText.ForeColor = [System.Drawing.Color]::Black
	$Global:LineURIFieldChange=$False
	$DisplayNumberText.Text = ""
	$DisplayNumberText.ForeColor = [System.Drawing.Color]::Black
	$Global:DisplayNumberFieldChange=$False
	$DisplayNameText.Text = ""
	$DisplayNameText.ForeColor = [System.Drawing.Color]::Black
	$Global:DisplayNameFieldChange=$False
	$DescriptionText.Text = ""
	$DescriptionText.ForeColor = [System.Drawing.Color]::Black
	$Global:DescriptionFieldChange=$False
	$ExUMEnabledText.Text = ""

}	

Function LoadPhoneInfo()
{

	$phone=Get-CSCommonAreaPhone $objListBox.SelectedItem.tostring() | select *


	$IdentityText.Text= $phone.Identity
	get-cspool | foreach {if ($_.Services -like "Registrar*") {$_.Identity}}

	$PoolDropDown.Items.Clear()
	$PoolDropDown.Text=$phone.registrarpool
	foreach ($x in $poolarray) 
	{
		[void]$PoolDropDown.Items.Add($x)
	}
	
	[void]$EnabledDropDown.Items.Clear()
	$EnabledDropDown.Text=$phone.enabled
	[void]$EnabledDropDown.Items.Add("True")
	[void]$EnabledDropDown.Items.Add("False")

	$SipAddressText.Text = $phone.SipAddress
	
	$DialPlanDropDown.Items.Clear()
	if ($phone.dialplan.friendlyname.length -lt 1) 
	{
		$DialPlanDropDown.Text="<Automatic>"
	}
	else
	{
		$DialPlanDropDown.Text=$phone.dialplan.friendlyname
	}
	[void]$DialPlanDropDown.Items.Add("<Automatic>")
	foreach ($x in $dialplanarray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$DialPlanDropDown.Items.Add($y)
	}


	$ClientPolicyDropDown.Items.Clear()
	if ($phone.ClientPolicy.FriendlyName.length -lt 1) 
	{
		$ClientPolicyDropDown.Text="<Automatic>"
	}
	else
	{
		$ClientPolicyDropDown.Text=$phone.clientpolicy.FriendlyName
	}
	[void]$ClientPolicyDropDown.Items.Add("<Automatic>")
	foreach ($x in $clientpolicyarray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$ClientPolicyDropDown.Items.Add($y)
	}

	$PinPolicyDropDown.Items.Clear()
	[void]$PinPolicyDropDown.Items.Add("<Automatic>")
	foreach ($x in $pinpolicyarray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$PinPolicyDropDown.Items.Add($y)
	}
	if ($phone.PinPolicy.friendlyname.length -lt 1) 
	{
		$PinPolicyDropDown.Text="<Automatic>"
	}
	else
	{
		$PinPolicyDropDown.Text=$phone.pinpolicy.friendlyname
	}

	$VoicePolicyDropDown.Items.Clear()
	[void]$VoicePolicyDropDown.Items.Add("<Automatic>")
	foreach ($x in $Voicepolicyarray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$VoicePolicyDropDown.Items.Add($y)
	}
	if ($phone.VoicePolicy.friendlyname.length -lt 1) 
	{
		$VoicePolicyDropDown.Text="<Automatic>"
	}
	else
	{
		$VoicePolicyDropDown.Text=$phone.Voicepolicy.friendlyname
	}


	$ConferencingPolicyDropDown.Items.Clear()
	[void]$ConferencingPolicyDropDown.Items.Add("<Automatic>")
	foreach ($x in $Confpolicyarray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$ConferencingPolicyDropDown.Items.Add($y)
	}
	if ($phone.ConferencingPolicy.friendlyname.length -lt 1) 
	{
		$ConferencingPolicyDropDown.Text="<Automatic>"
	}
	else
	{
		$ConferencingPolicyDropDown.Text=$phone.Conferencingpolicy.friendlyname
	}





	$LineURIText.text = $phone.LineURI
	$DisplayNumberText.text = $phone.DisplayNumber
	$DisplayNameText.text	= $phone.DisplayName
	$DescriptionText.text = $phone.Description
	$ExUMEnabledText.text = $phone.ExUmEnabled

	#########################################################
	#	Reset field colors and selections
	#########################################################
	$IdentityText.ForeColor = [System.Drawing.Color]::Black
	$PoolDropDown.ForeColor = [System.Drawing.Color]::Black
	$Global:PoolFieldChange=$False
	$SipAddressText.ForeColor = [System.Drawing.Color]::Black
	$Global:SipAddressFieldChange=$False
	$EnabledDropDown.ForeColor = [System.Drawing.Color]::Black
	$Global:EnabledFieldChange=$False
	$DialPlanDropDown.ForeColor = [System.Drawing.Color]::Black
	$Global:DialPlanFieldChange=$False
	$ClientPolicyDropDown.ForeColor = [System.Drawing.Color]::Black
	$Global:ClientPolicyFieldChange=$False
	$PinPolicyDropDown.ForeColor = [System.Drawing.Color]::Black
	$Global:PinPolicyFieldChange=$False
	$VoicePolicyDropDown.ForeColor = [System.Drawing.Color]::Black
	$Global:VoicePolicyFieldChange=$False
	$ConferencingPolicyDropDown.ForeColor = [System.Drawing.Color]::Black
	$Global:ConferencingPolicyFieldChange=$False
	$LineURIText.ForeColor = [System.Drawing.Color]::Black
	$Global:LineURIFieldChange=$False
	$DisplayNumberText.ForeColor = [System.Drawing.Color]::Black
	$Global:DisplayNumberFieldChange=$False
	$DisplayNameText.ForeColor = [System.Drawing.Color]::Black
	$Global:DisplayNameFieldChange=$False
	$DescriptionText.ForeColor = [System.Drawing.Color]::Black
	$Global:DescriptionFieldChange=$False


}	

Function LoadMultiSelectInfo()
{

	$IdentityText.Text=""
	get-cspool | foreach {if ($_.Services -like "Registrar*") {$_.Identity}}

	$PoolDropDown.Items.Clear()
	$PoolDropDown.Text=""
	foreach ($x in $poolarray) 
	{
		[void]$PoolDropDown.Items.Add($x)
	}
	
	[void]$EnabledDropDown.Items.Clear()
	$EnabledDropDown.Text=""
	[void]$EnabledDropDown.Items.Add("True")
	[void]$EnabledDropDown.Items.Add("False")

	$SipAddressText.Text = ""
	
	$DialPlanDropDown.Items.Clear()
	$DialPlanDropDown.Text=""
	[void]$DialPlanDropDown.Items.Add("<Automatic>")
	foreach ($x in $dialplanarray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$DialPlanDropDown.Items.Add($y)
	}


	$ClientPolicyDropDown.Items.Clear()
	$ClientPolicyDropDown.Text=""
	[void]$ClientPolicyDropDown.Items.Add("<Automatic>")
	foreach ($x in $clientpolicyarray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$ClientPolicyDropDown.Items.Add($y)
	}

	$PinPolicyDropDown.Items.Clear()
	$PinPolicyDropDown.Text=""
	[void]$PinPolicyDropDown.Items.Add("<Automatic>")
	foreach ($x in $pinpolicyarray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$PinPolicyDropDown.Items.Add($y)
	}

	$VoicePolicyDropDown.Items.Clear()
	$VoicePolicyDropDown.Text=""
	[void]$VoicePolicyDropDown.Items.Add("<Automatic>")
	foreach ($x in $Voicepolicyarray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$VoicePolicyDropDown.Items.Add($y)
	}

	$ConferencingPolicyDropDown.Items.Clear()
	$ConferencingPolicyDropDown.Text=""
	[void]$ConferencingPolicyDropDown.Items.Add("<Automatic>")
	foreach ($x in $Confpolicyarray) 
	{
		$y = $x.identity -replace "Tag:",""
		[void]$ConferencingPolicyDropDown.Items.Add($y)
	}

	$LineURIText.text = ""
	$DisplayNumberText.text = ""
	$DisplayNameText.text	= ""
	$DescriptionText.text = ""
	$ExUMEnabledText.text = ""

	#########################################################
	#	Reset field colors and selections
	#########################################################
	$IdentityText.ForeColor = [System.Drawing.Color]::Black
	$PoolDropDown.ForeColor = [System.Drawing.Color]::Black
	$Global:PoolFieldChange=$False
	$SipAddressText.ForeColor = [System.Drawing.Color]::Black
	$Global:SipAddressFieldChange=$False
	$EnabledDropDown.ForeColor = [System.Drawing.Color]::Black
	$Global:EnabledFieldChange=$False
	$DialPlanDropDown.ForeColor = [System.Drawing.Color]::Black
	$Global:DialPlanFieldChange=$False
	$ClientPolicyDropDown.ForeColor = [System.Drawing.Color]::Black
	$Global:ClientPolicyFieldChange=$False
	$PinPolicyDropDown.ForeColor = [System.Drawing.Color]::Black
	$Global:PinPolicyFieldChange=$False
	$VoicePolicyDropDown.ForeColor = [System.Drawing.Color]::Black
	$Global:VoicePolicyFieldChange=$False
	$ConferencingPolicyDropDown.ForeColor = [System.Drawing.Color]::Black
	$Global:ConferencingPolicyFieldChange=$False
	$LineURIText.ForeColor = [System.Drawing.Color]::Black
	$Global:LineURIFieldChange=$False
	$DisplayNumberText.ForeColor = [System.Drawing.Color]::Black
	$Global:DisplayNumberFieldChange=$False
	$DisplayNameText.ForeColor = [System.Drawing.Color]::Black
	$Global:DisplayNameFieldChange=$False
	$DescriptionText.ForeColor = [System.Drawing.Color]::Black
	$Global:DescriptionFieldChange=$False


}	



$CAC_FormSizeChanged = { 

$RefreshSelectedPhoneButton.Location = New-Object System.Drawing.Size((10 + (($objform.width-50) /7 * 1) ),($objform.height - 90))
$RefreshPhoneListButton.Location = New-Object System.Drawing.Size((10 + (($objform.width-50) /7 * 0) ),($objform.height - 90))
$SaveChangesButton.Location = New-Object System.Drawing.Size((10 + (($objform.width-50) /7 * 3) ),($objform.height - 90))
$RemovePhoneButton.Location = New-Object System.Drawing.Size((10 + (($objform.width-50) /7 * 4) ),($objform.height - 90))
$NewPhoneButton.Location = New-Object System.Drawing.Size((10 + (($objform.width-50) /7 * 5) ),($objform.height - 90))
$CancelButton.Location = New-Object System.Drawing.Size((10 + (($objform.width-50) /7 * 6) ),($objform.height - 90))
$SetPhonePinButton.Location = New-Object System.Drawing.Size((10 + (($objform.width-50) /7 * 2) ),($objform.height - 90))

} 
 
   
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void] [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

#For Windows 2008 support
import-module Lync

$Global:LyncFixIcon = [System.Convert]::FromBase64String('
AAABAAEAJiEAAAEACABYCgAAFgAAACgAAAAmAAAAQgAAAAEACAAAAAAAKAUAAAAAAAAAAAAAAAEAAAAB
AAAAAAAAMwAAAGYAAACZAAAAzAAAAP8AAAAAKwAAMysAAGYrAACZKwAAzCsAAP8rAAAAVQAAM1UAAGZV
AACZVQAAzFUAAP9VAAAAgAAAM4AAAGaAAACZgAAAzIAAAP+AAAAAqgAAM6oAAGaqAACZqgAAzKoAAP+q
AAAA1QAAM9UAAGbVAACZ1QAAzNUAAP/VAAAA/wAAM/8AAGb/AACZ/wAAzP8AAP//AAAAADMAMwAzAGYA
MwCZADMAzAAzAP8AMwAAKzMAMyszAGYrMwCZKzMAzCszAP8rMwAAVTMAM1UzAGZVMwCZVTMAzFUzAP9V
MwAAgDMAM4AzAGaAMwCZgDMAzIAzAP+AMwAAqjMAM6ozAGaqMwCZqjMAzKozAP+qMwAA1TMAM9UzAGbV
MwCZ1TMAzNUzAP/VMwAA/zMAM/8zAGb/MwCZ/zMAzP8zAP//MwAAAGYAMwBmAGYAZgCZAGYAzABmAP8A
ZgAAK2YAMytmAGYrZgCZK2YAzCtmAP8rZgAAVWYAM1VmAGZVZgCZVWYAzFVmAP9VZgAAgGYAM4BmAGaA
ZgCZgGYAzIBmAP+AZgAAqmYAM6pmAGaqZgCZqmYAzKpmAP+qZgAA1WYAM9VmAGbVZgCZ1WYAzNVmAP/V
ZgAA/2YAM/9mAGb/ZgCZ/2YAzP9mAP//ZgAAAJkAMwCZAGYAmQCZAJkAzACZAP8AmQAAK5kAMyuZAGYr
mQCZK5kAzCuZAP8rmQAAVZkAM1WZAGZVmQCZVZkAzFWZAP9VmQAAgJkAM4CZAGaAmQCZgJkAzICZAP+A
mQAAqpkAM6qZAGaqmQCZqpkAzKqZAP+qmQAA1ZkAM9WZAGbVmQCZ1ZkAzNWZAP/VmQAA/5kAM/+ZAGb/
mQCZ/5kAzP+ZAP//mQAAAMwAMwDMAGYAzACZAMwAzADMAP8AzAAAK8wAMyvMAGYrzACZK8wAzCvMAP8r
zAAAVcwAM1XMAGZVzACZVcwAzFXMAP9VzAAAgMwAM4DMAGaAzACZgMwAzIDMAP+AzAAAqswAM6rMAGaq
zACZqswAzKrMAP+qzAAA1cwAM9XMAGbVzACZ1cwAzNXMAP/VzAAA/8wAM//MAGb/zACZ/8wAzP/MAP//
zAAAAP8AMwD/AGYA/wCZAP8AzAD/AP8A/wAAK/8AMyv/AGYr/wCZK/8AzCv/AP8r/wAAVf8AM1X/AGZV
/wCZVf8AzFX/AP9V/wAAgP8AM4D/AGaA/wCZgP8AzID/AP+A/wAAqv8AM6r/AGaq/wCZqv8AzKr/AP+q
/wAA1f8AM9X/AGbV/wCZ1f8AzNX/AP/V/wAA//8AM///AGb//wCZ//8AzP//AP///wAAAAAAAAAAAAAA
AAAAAAAAHB0WHRwdHRwXHRwdHRwdHB0cFx0cHRwXHRwdHRwdHB0dFh0dHB0AAB0cHRwXHB0WHRwXHB0W
HRYdHB0cFxwXHB0WHRwXHBccHRwdFh0cAAAcHRYdHB0cHRwdHB0cHRwdHB0WHRwdHB0cHRwdHB0cHRwX
HB0cHQAAHRwdHBcdFh0cFxwdFh0XHB0dHB0WHRwXHBccFxwdFh0cHRwXHB0AAB0cFxwdHB0cHR0cHR0c
HRwdHBccHR0cHR0cHR0cHR0cHRccHRwdAAAdHB0dFh0cFxwXHBccFxwdFh0cHRYdFh0cFxwdFh0WHRwd
HBccHQAAHRwXHB0cHRwdHB0cHRwdHB0WHRwdHB0cHRwdHB0cHRwdFh0dHB0AAB0cHRwXHBccHRwXHB0c
Fx0cHRwXHB0cFxwXHRYdHBccHRwdFh0cAAAdHBccHRwdHBccHRwXHB0cHRYdHB0WHRwdHB0cHRwdHRwX
HB0cHQAAHRwdd/v7+/v7+/v7+/v7mh3R+/v7+/v1HBccHRYdHB0cHRwXHB0AAB0WHSL7+/v7+/v7+/v7
+/UcTfv7+/v7+3AdHB0dFh0WHRccHRwdAAAdHB0X0fv7+/v7+/v7+/v7mh3R+/v7+/uaHRYdHB0cHRwd
HBccHQAAFh0cHU37+/v7+/v7+/v7+/Udp/v7+/v7yxwdHB0cFxwdFh0dFh0AAB0cFxwd0fv7+/v7+/v7
+/v7cB37+/v7+/tAHRYdHB0dHB0cHRwdAAAdHB0cF6f7+/v7+8oXHB0cHR0c0fv7+/v7+/v7+/vEHRwd
Fh0cHQAAHRYdHRxN+/v7+/v1HB0XHB0WHXf7+/v7+/v7+/v79BccFxwdFxwAAB0cHRwXHNH7+/v7+3Ad
HB0cHRwd+/v7+/v7+/v7+/tHHB0cHRwdAAAdFh0cHR2n+/v7+/ubHB0WHRwdHdH7+/v7+/v7+/v7xB0c
HRYdHAAAHRwdHB0cTfv7+/v79B0cHRwdFh13+/v7+/v7+/v7+/UdFh0dHB0AAB0cHRccFx37+/v7+/tw
HRwXHB0cHfv7+/v7+0YdHB0cHRwcHRwXAAAcFxwdHB0c0fv7+/v7xRYdHRwdHB3R+/v7+/v7+/v7+/vF
HRYdHAAAHRwdFh0cHXf7+/v7+/tHHB0WHRccd/v7+/v7+/v7+/v79B0cHR0AABccHRwdFh0d+/v7+/v7
mhccHRwdHE37+/v7+/v7+/v7+/UdHBccAAAdHBcdHB0cHdH7+/v7+8scHRYdHB0d0fv7+/v7+/v7+/v7
cB0cHQAAHB0cHRwdFh13+/v7+/v7ah0dHB0WHaH7+/v7+/v7+/v7+8UcHR0AAB0cFxwXHB0cHRwdFh0c
HRwdHBccHRwdHB0WHRwXHB0cHRwXHBccAAAWHRwdHB0dHB0WHRwdHRwdFh0cHRwdHRYdHB0cHR0WHRYd
HB0cHQAAHRwXHB0cFxwXHB0WHRYdFh0cFxwXHBccHRYdFh0cHRwdHBccHRwAAB0cHR0WHRwdHRwdHRwd
HB0dHB0dHB0cHR0cHR0cHRwdFxwdFxwdAAAdFh0cHRwXHB0WHRwdFh0cFxwdHBccHRYdHB0WHRwXHB0c
HRwdHAAAHB0cHRYdHB0cHRwXHB0cHRwdFh0cHRwdHBccHRwdHB0WHRwdFh0AAB0cFxwdHBccHRYdHB0W
HRwXHB0cFxwdFh0cHRYdHBccHRwXHB0cAAAcHR0cHR0cHR0cHR0cHR0cHR0cHR0cHR0cHR0cHR0cHR0c
HR0cHQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
')

$poolarray=get-cspool | foreach {if ($_.Services -like "Registrar*") {$_.Identity}}
$dialplanarray=get-csdialplan | where {$_.Identity -like "Tag:*"}
$clientpolicyarray=get-csclientpolicy | where {$_.Identity -like "Tag:*"}
$pinpolicyarray=get-cspinpolicy | where {$_.Identity -like "Tag:*"}
$voicepolicyarray=get-csvoicepolicy | where {$_.Identity -like "Tag:*"}
$confpolicyarray=get-csconferencingpolicy | where {$_.Identity -like "Tag:*"}

$AddCAP_PoolArray=get-cspool | foreach {if ($_.Services -like "Registrar*") {$_.Identity}}
$AddCAP_DialPlanarray=get-csdialplan | where {$_.Identity -like "Tag:*"}
$AddCAP_ClientPolicyarray=get-csclientpolicy | where {$_.Identity -like "Tag:*"}
$AddCAP_PinPolicyarray=get-cspinpolicy | where {$_.Identity -like "Tag:*"}
$AddCAP_VoicePolicyarray=get-csvoicepolicy | where {$_.Identity -like "Tag:*"}
$AddCAP_ConfPolicyArray=get-csconferencingpolicy  | where {$_.Identity -like "Tag:*"}

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Lync Common Area Phone Tool v1.2"
$objForm.Size = New-Object System.Drawing.Size(880,600) 
$objForm.MinimumSize = New-Object System.Drawing.Size(880,600) 
$objForm.StartPosition = "CenterScreen"
$ObjForm.Add_SizeChanged($CAC_FormSizeChanged)
$objForm.KeyPreview = $True
$objForm.Icon = $Global:LyncFixIcon

$TitleLabel = New-Object System.Windows.Forms.Label
$TitleLabel.Location = New-Object System.Drawing.Size(10,10) 
$TitleLabel.Size = New-Object System.Drawing.Size(780,30) 
$TitleLabel.Text = "The purpose of this tool is to give you an easy front end for working with Common Area Phones.  Please use the Q/A section of the TechNet gallery to report bugs or suggest features you would like to see.  Use only at your own risk."
$objForm.Controls.Add($TitleLabel) 

$TitleLabel2 = New-Object System.Windows.Forms.Label
$TitleLabel2.Location = New-Object System.Drawing.Size(10,50) 
$TitleLabel2.Size = New-Object System.Drawing.Size(780,30) 
$TitleLabel2.ForeColor = [System.Drawing.Color]::Red
$TitleLabel2.Text = "Warning: Changes may take time to replicate.  You may need to use refresh more than once to see updates."
$objForm.Controls.Add($TitleLabel2) 

$FilterLabel = New-Object System.Windows.Forms.Label
$FilterLabel.Location = New-Object System.Drawing.Size(10,80) 
$FilterLabel.Size = New-Object System.Drawing.Size(40,20) 
$FilterLabel.Text = "Filter:"
$objForm.Controls.Add($FilterLabel) 

$FilterDropDown = new-object System.Windows.Forms.ComboBox
$FilterDropDown.Location = new-object System.Drawing.Size(50,80) 
$FilterDropDown.Size = New-Object System.Drawing.Size(260,20) 
$FilterDropDown.Anchor = 'Top, Left, Right'

#Load Pools Into Filter
[void]$FilterDropDown.Items.Add("<All Pools>")
foreach ($x in $poolarray) 
{
	[void]$FilterDropDown.Items.Add($x)
}
$FilterDropDown.SelectedIndex=0
$FilterDropDown.add_SelectedIndexChanged({
	$RefreshPhoneListButton.Enabled=$False
	$RefreshPhoneListButton.Text = "Loading..."
	$objListBox.Items.Clear()
	LoadPhones
	$RefreshPhoneListButton.Text = "Refresh Phone List"
	$RefreshPhoneListButton.Enabled=$True
})
$objForm.Controls.Add($FilterDropDown) 

$objListBox = New-Object System.Windows.Forms.ListBox 
$objListBox.Location = New-Object System.Drawing.Size(10,110) 
$objListBox.Size = New-Object System.Drawing.Size(300,390) 
$objListBox.Anchor = 'Top, Bottom,Left'
$objListBox.Sorted = $True
$objListbox.SelectionMode = "MultiExtended"
$objListBox.add_SelectedIndexChanged({
	if ($objListbox.SelectedItems.count -gt 1)
	{
		$SipAddressText.Enabled=$false
		$SipAddressLabel.Enabled=$false
		$LineURIText.Enabled=$false
		$LineURILabel.Enabled=$false
		$DisplayNumberText.Enabled=$false
		$DisplayNumberLabel.Enabled=$false
		$DisplayNameText.Enabled=$false
		$DisplayNameLabel.Enabled=$false
		$DescriptionText.Enabled=$false
		$DescriptionLabel.Enabled=$false
		$ExUMEnabledLabel.Enabled=$false
		$RefreshSelectedPhoneButton.Enabled=$false
		LoadMultiSelectInfo
		
	}
	else
	{

		$SipAddressText.Enabled=$true
		$SipAddressLabel.Enabled=$true
		$LineURIText.Enabled=$true
		$LineURILabel.Enabled=$true
		$DisplayNumberText.Enabled=$true
		$DisplayNumberLabel.Enabled=$true
		$DisplayNameText.Enabled=$true
		$DisplayNameLabel.Enabled=$true
		$DescriptionText.Enabled=$true
		$DescriptionLabel.Enabled=$true
		$ExUMEnabledLabel.Enabled=$true
		$RefreshSelectedPhoneButton.Enabled=$true
		LoadPhoneInfo
	}
})
$objForm.Controls.Add($objListBox) 

$IdentityLabel = New-Object System.Windows.Forms.Label
$IdentityLabel.Location = New-Object System.Drawing.Size(320,80) 
$IdentityLabel.Size = New-Object System.Drawing.Size(120,20) 
$IdentityLabel.Text = "Identity:"
$objForm.Controls.Add($IdentityLabel) 

$IdentityText = New-Object System.Windows.Forms.Textbox
$IdentityText.Location = new-object System.Drawing.Size(440,80) 
$IdentityText.Size = New-Object System.Drawing.Size(400,20) 
$IdentityText.Text = ""
$IdentityText.Anchor = 'Top, Left, Right'
$IdentityText.ReadOnly = "True"
$objForm.Controls.Add($IdentityText) 

$PoolLabel = New-Object System.Windows.Forms.Label
$PoolLabel.Location = New-Object System.Drawing.Size(320,110) 
$PoolLabel.Size = New-Object System.Drawing.Size(120,20) 
$PoolLabel.Text = "Pool:"
$objForm.Controls.Add($PoolLabel) 

$PoolDropDown = new-object System.Windows.Forms.ComboBox
$PoolDropDown.Location = new-object System.Drawing.Size(440,110) 
$PoolDropDown.Size = New-Object System.Drawing.Size(400,20) 
#$PoolDropDown.add_SelectedIndexChanged($OnSelect_PoolDropDown) 
$PoolDropDown.Anchor = 'Top, Left, Right'
$PoolDropDown.add_SelectedIndexChanged({
$PoolDropDown.ForeColor = [System.Drawing.Color]::Green
$Global:PoolFieldChange=$true
})
$objForm.Controls.Add($PoolDropDown) 

$EnabledLabel = New-Object System.Windows.Forms.Label
$EnabledLabel.Location = New-Object System.Drawing.Size(320,140) 
$EnabledLabel.Size = New-Object System.Drawing.Size(120,20) 
$EnabledLabel.Text = "Enabled:"
$objForm.Controls.Add($EnabledLabel) 

$EnabledDropDown = new-object System.Windows.Forms.ComboBox
$EnabledDropDown.Location = new-object System.Drawing.Size(440,140) 
$EnabledDropDown.Size = New-Object System.Drawing.Size(400,20) 
$EnabledDropDown.add_SelectedIndexChanged({
$EnabledDropDown.ForeColor = [System.Drawing.Color]::Green
$Global:EnabledFieldChange=$true
}) 
$EnabledDropDown.Anchor = 'Top, Left, Right'
$objForm.Controls.Add($EnabledDropDown) 

$SipAddressLabel = New-Object System.Windows.Forms.Label
$SipAddressLabel.Location = New-Object System.Drawing.Size(320,170) 
$SipAddressLabel.Size = New-Object System.Drawing.Size(120,20) 
$SipAddressLabel.Text = "Sip Address:"
$objForm.Controls.Add($SipAddressLabel) 

$SipAddressText = New-Object System.Windows.Forms.Textbox
$SipAddressText.Location = new-object System.Drawing.Size(440,170) 
$SipAddressText.Size = New-Object System.Drawing.Size(400,20) 
$SipAddressText.Text = ""
$SipAddressText.Anchor = 'Top, Left, Right'
$SipAddressText.add_TextChanged({
$SipAddressText.ForeColor = [System.Drawing.Color]::Green
$Global:SipAddressFieldChange=$true
})
$objForm.Controls.Add($SipAddressText) 

$DialPlanLabel = New-Object System.Windows.Forms.Label
$DialPlanLabel.Location = New-Object System.Drawing.Size(320,200) 
$DialPlanLabel.Size = New-Object System.Drawing.Size(120,20) 
$DialPlanLabel.Text = "DialPlan:"
$objForm.Controls.Add($DialPlanLabel) 

$DialPlanDropDown = new-object System.Windows.Forms.ComboBox
$DialPlanDropDown.Location = new-object System.Drawing.Size(440,200) 
$DialPlanDropDown.Size = New-Object System.Drawing.Size(400,20) 
#$DialPlanDropDown.add_SelectedIndexChanged($OnSelect_DialPlanDropDown) 
$DialPlanDropDown.Anchor = 'Top, Left, Right'
$DialPlanDropDown.add_SelectedIndexChanged({
$DialPlanDropDown.ForeColor = [System.Drawing.Color]::Green
$Global:DialPlanFieldChange=$true
}) 
$objForm.Controls.Add($DialPlanDropDown) 

$ClientPolicyLabel = New-Object System.Windows.Forms.Label
$ClientPolicyLabel.Location = New-Object System.Drawing.Size(320,230) 
$ClientPolicyLabel.Size = New-Object System.Drawing.Size(120,20) 
$ClientPolicyLabel.Text = "ClientPolicy:"
$objForm.Controls.Add($ClientPolicyLabel) 

$ClientPolicyDropDown = new-object System.Windows.Forms.ComboBox
$ClientPolicyDropDown.Location = new-object System.Drawing.Size(440,230) 
$ClientPolicyDropDown.Size = New-Object System.Drawing.Size(400,20) 
#$ClientPolicyDropDown.add_SelectedIndexChanged($OnSelect_ClientPolicyDropDown) 
$ClientPolicyDropDown.Anchor = 'Top, Left, Right'
$ClientPolicyDropDown.add_SelectedIndexChanged({
$ClientPolicyDropDown.ForeColor = [System.Drawing.Color]::Green
$Global:ClientPolicyFieldChange=$true
}) 
$objForm.Controls.Add($ClientPolicyDropDown) 

$PinPolicyLabel = New-Object System.Windows.Forms.Label
$PinPolicyLabel.Location = New-Object System.Drawing.Size(320,260) 
$PinPolicyLabel.Size = New-Object System.Drawing.Size(120,20) 
$PinPolicyLabel.Text = "PinPolicy:"
$objForm.Controls.Add($PinPolicyLabel) 

$PinPolicyDropDown = new-object System.Windows.Forms.ComboBox
$PinPolicyDropDown.Location = new-object System.Drawing.Size(440,260) 
$PinPolicyDropDown.Size = New-Object System.Drawing.Size(400,20) 
#$PinPolicyDropDown.add_SelectedIndexChanged($OnSelect_PinPolicyDropDown) 
$PinPolicyDropDown.Anchor = 'Top, Left, Right'
$PinPolicyDropDown.add_SelectedIndexChanged({
$PinPolicyDropDown.ForeColor = [System.Drawing.Color]::Green
$Global:PinPolicyFieldChange=$true
}) 
$objForm.Controls.Add($PinPolicyDropDown) 

$VoicePolicyLabel = New-Object System.Windows.Forms.Label
$VoicePolicyLabel.Location = New-Object System.Drawing.Size(320,290) 
$VoicePolicyLabel.Size = New-Object System.Drawing.Size(120,20) 
$VoicePolicyLabel.Text = "VoicePolicy:"
$objForm.Controls.Add($VoicePolicyLabel) 

$VoicePolicyDropDown = new-object System.Windows.Forms.ComboBox
$VoicePolicyDropDown.Location = new-object System.Drawing.Size(440,290) 
$VoicePolicyDropDown.Size = New-Object System.Drawing.Size(400,20) 
#$VoicePolicyDropDown.add_SelectedIndexChanged($OnSelect_VoicePolicyDropDown) 
$VoicePolicyDropDown.Anchor = 'Top, Left, Right'
$VoicePolicyDropDown.add_SelectedIndexChanged({
$VoicePolicyDropDown.ForeColor = [System.Drawing.Color]::Green
$Global:VoicePolicyFieldChange=$True
})
$objForm.Controls.Add($VoicePolicyDropDown) 


$ConferencingPolicyLabel = New-Object System.Windows.Forms.Label
$ConferencingPolicyLabel.Location = New-Object System.Drawing.Size(320,320) 
$ConferencingPolicyLabel.Size = New-Object System.Drawing.Size(120,20) 
$ConferencingPolicyLabel.Text = "ConferencingPolicy:"
$objForm.Controls.Add($ConferencingPolicyLabel) 

$ConferencingPolicyDropDown = new-object System.Windows.Forms.ComboBox
$ConferencingPolicyDropDown.Location = new-object System.Drawing.Size(440,320) 
$ConferencingPolicyDropDown.Size = New-Object System.Drawing.Size(400,20) 
#$ConferencingPolicyDropDown.add_SelectedIndexChanged($OnSelect_ConferencingPolicyDropDown) 
$ConferencingPolicyDropDown.Anchor = 'Top, Left, Right'
$ConferencingPolicyDropDown.add_SelectedIndexChanged({
$ConferencingPolicyDropDown.ForeColor = [System.Drawing.Color]::Green
$Global:ConferencingPolicyFieldChange=$True
})
$objForm.Controls.Add($ConferencingPolicyDropDown) 

$LineURILabel = New-Object System.Windows.Forms.Label
$LineURILabel.Location = New-Object System.Drawing.Size(320,350) 
$LineURILabel.Size = New-Object System.Drawing.Size(120,20) 
$LineURILabel.Text = "LineURI:"
$objForm.Controls.Add($LineURILabel) 

$LineURIText = New-Object System.Windows.Forms.Textbox
$LineURIText.Location = new-object System.Drawing.Size(440,350) 
$LineURIText.Size = New-Object System.Drawing.Size(400,20) 
$LineURIText.Text = ""
$LineURIText.Anchor = 'Top, Left, Right'
$LineURIText.add_TextChanged({
$LineURIText.ForeColor = [System.Drawing.Color]::Green
$Global:LineURIFieldChange=$true
})
$objForm.Controls.Add($LineURIText) 

$DisplayNumberLabel = New-Object System.Windows.Forms.Label
$DisplayNumberLabel.Location = New-Object System.Drawing.Size(320,380) 
$DisplayNumberLabel.Size = New-Object System.Drawing.Size(120,20) 
$DisplayNumberLabel.Text = "DisplayNumber:"
$objForm.Controls.Add($DisplayNumberLabel) 

$DisplayNumberText = New-Object System.Windows.Forms.Textbox
$DisplayNumberText.Location = new-object System.Drawing.Size(440,380) 
$DisplayNumberText.Size = New-Object System.Drawing.Size(400,20) 
$DisplayNumberText.Text = ""
$DisplayNumberText.Anchor = 'Top, Left, Right'
$DisplayNumberText.add_TextChanged({
$DisplayNumberText.ForeColor = [System.Drawing.Color]::Green
$Global:DisplayNumberFieldChange=$true
})
$objForm.Controls.Add($DisplayNumberText) 

$DisplayNameLabel = New-Object System.Windows.Forms.Label
$DisplayNameLabel.Location = New-Object System.Drawing.Size(320,410) 
$DisplayNameLabel.Size = New-Object System.Drawing.Size(120,20) 
$DisplayNameLabel.Text = "DisplayName:"
$objForm.Controls.Add($DisplayNameLabel) 

$DisplayNameText = New-Object System.Windows.Forms.Textbox
$DisplayNameText.Location = new-object System.Drawing.Size(440,410) 
$DisplayNameText.Size = New-Object System.Drawing.Size(400,20) 
$DisplayNameText.Text = ""
$DisplayNameText.Anchor = 'Top, Left, Right'
$DisplayNameText.add_TextChanged({
$DisplayNameText.ForeColor = [System.Drawing.Color]::Green
$Global:DisplayNameFieldChange=$true
})
$objForm.Controls.Add($DisplayNameText) 

$DescriptionLabel = New-Object System.Windows.Forms.Label
$DescriptionLabel.Location = New-Object System.Drawing.Size(320,440) 
$DescriptionLabel.Size = New-Object System.Drawing.Size(120,20) 
$DescriptionLabel.Text = "Description:"
$objForm.Controls.Add($DescriptionLabel) 

$DescriptionText = New-Object System.Windows.Forms.Textbox
$DescriptionText.Location = new-object System.Drawing.Size(440,440) 
$DescriptionText.Size = New-Object System.Drawing.Size(400,20) 
$DescriptionText.Text = ""
$DescriptionText.Anchor = 'Top, Left, Right'
$DescriptionText.add_TextChanged({
$DescriptionText.ForeColor = [System.Drawing.Color]::Green
$Global:DescriptionFieldChange=$true
})
$objForm.Controls.Add($DescriptionText) 

$ExUMEnabledLabel = New-Object System.Windows.Forms.Label
$ExUMEnabledLabel.Location = New-Object System.Drawing.Size(320,470) 
$ExUMEnabledLabel.Size = New-Object System.Drawing.Size(120,20) 
$ExUMEnabledLabel.Text = "ExUMEnabled:"
$objForm.Controls.Add($ExUMEnabledLabel) 

$ExUMEnabledText = New-Object System.Windows.Forms.Textbox
$ExUMEnabledText.Location = new-object System.Drawing.Size(440,470) 
$ExUMEnabledText.Size = New-Object System.Drawing.Size(400,20) 
$ExUMEnabledText.Text = ""
$ExUMEnabledText.Anchor = 'Top, Left, Right'
$ExUmEnabledText.Readonly = "True"
$objForm.Controls.Add($ExUMEnabledText) 

$RefreshPhoneListButton = New-Object System.Windows.Forms.Button
$RefreshPhoneListButton.Location = New-Object System.Drawing.Size((10 + (($objform.width-50) /7 * 0) ),($objform.height - 90))
$RefreshPhoneListButton.Size = New-Object System.Drawing.Size(115,25)
$RefreshPhoneListButton.Text = "Refresh Phone List"
$RefreshPhoneListButton.Add_Click({
	$RefreshPhoneListButton.Enabled=$False
	$RefreshPhoneListButton.Text = "Loading..."
	$objListBox.Items.Clear()
	LoadPhones
	$RefreshPhoneListButton.Text = "Refresh Phone List"
	$RefreshPhoneListButton.Enabled=$True
})
$RefreshPhoneListButton.Anchor = 'Bottom, Left'
$objForm.Controls.Add($RefreshPhoneListButton)

$RefreshSelectedPhoneButton = New-Object System.Windows.Forms.Button
$RefreshSelectedPhoneButton.Location = New-Object System.Drawing.Size((10 + (($objform.width-50) /7 * 1) ),($objform.height - 90))
$RefreshSelectedPhoneButton.Size = New-Object System.Drawing.Size(115,25)
$RefreshSelectedPhoneButton.Text = "Refresh Phone"
$RefreshSelectedPhoneButton.Add_Click({
LoadPhoneInfo
})
$RefreshSelectedPhoneButton.Anchor = 'Bottom, Left'
$objForm.Controls.Add($RefreshSelectedPhoneButton)


$SetPhonePinButton = New-Object System.Windows.Forms.Button
$SetPhonePinButton.Location = New-Object System.Drawing.Size((10 + (($objform.width-50) /7 * 2) ),($objform.height - 90))
$SetPhonePinButton.Size = New-Object System.Drawing.Size(115,25)
$SetPhonePinButton.Text = "Set Pin"
$SetPhonePinButton.Add_Click({
	$SetPhonePinButton.Text = "Setting"
	$SetPhonePinButton.Enabled=$False
	if ($objListBox.SelectedIndex -gt -1) 
	{
		$pinquery = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the new pin for the selected phone(s)", "Change Phone Pin", "")
		if ($pinquery.length -gt 0)
		{
			foreach ($PhoneName in $objListbox.SelectedItems)
			{      
				$error.clear()
				Set-CsClientPin $PhoneName -pin $pinquery
				if ($error.count -gt 0) 
				{
					$MyError=$Error[0].Exception
					[Microsoft.VisualBasic.Interaction]::MsgBox("$PhoneName $MyError",'OKOnly,Critical', "Pin Error")		
				}
			}
		}
	}
	else
	{
		[Microsoft.VisualBasic.Interaction]::MsgBox("Please select a phone first!",'OKOnly,Information', "Change Phone Pin")
	}
	$SetPhonePinButton.Text = "Set Pin"
	$SetPhonePinButton.Enabled=$True
})
$SetPhonePinButton.Anchor = 'Bottom, Left'
$objForm.Controls.Add($SetPhonePinButton)
$SaveChangesButton = New-Object System.Windows.Forms.Button
$SaveChangesButton.Location = New-Object System.Drawing.Size((10 + (($objform.width-50) /7 * 3) ),($objform.height - 90))
$SaveChangesButton.Size = New-Object System.Drawing.Size(115,25)
$SaveChangesButton.Text = "Save Changes"
$SaveChangesButton.Add_Click({
	$SaveChangesButton.Enabled=$False
	$SaveChangesButton.Text = "Working..."

	if ($objListBox.SelectedIndex -gt -1) 
	{
		foreach ($PhoneName in $objListbox.SelectedItems)
		{
			$phone=Get-CSCommonAreaPhone $PhoneName
	
			if ($Global:PoolFieldChange)
			{
				Move-CsCommonAreaPhone -Identity $phone.identity -Target $PoolDropDown.Text -Confirm:$False
			}
			if ($Global:EnabledFieldChange) 
			{
				if ($EnabledDropDown.Text -eq "False") {
					Set-CSCommonAreaPhone -Identity $phone.identity -Enabled $False
				}
				else
				{
					Set-CSCommonAreaPhone -Identity $phone.identity -Enabled $True
				}
			}
			if ($Global:SipAddressFieldChange) 	
			{
				#Just in case SIP is there, remove it because we're going to readd it to ensure it's there.
				$SipAddressText.Text = $SipAddressText.Text -replace "sip:",""
				$SipAddressText.Text = "sip:" + $SipAddressText.Text
				Set-CSCommonAreaPhone -Identity $phone.identity -SipAddress $SipAddressText.Text
			}
			if ($Global:DialPlanFieldChange) 
			{
				if ($DialPlanDropDown.Text -eq "<Automatic>") {
					Grant-CsDialPlan -Identity $phone.identity -PolicyName $Null
				}
				else
				{
					Grant-CsDialPlan -Identity $phone.identity -PolicyName $DialPlanDropDown.SelectedItem.tostring()
				}	
			}
			if ($Global:ClientPolicyFieldChange) 
			{
				if ($ClientPolicyDropDown.Text -eq "<Automatic>") {
					Grant-CsClientPolicy -Identity $phone.identity -PolicyName $Null
				}
				else
				{
					Grant-CsClientPolicy -Identity $phone.identity -PolicyName $ClientPolicyDropDown.SelectedItem.tostring()
				}	
			}
			if ($Global:PinPolicyFieldChange) 
			{
				if ($PinPolicyDropDown.Text -eq "<Automatic>") {
					Grant-CsPinPolicy -Identity $phone.identity -PolicyName $Null
				}
				else
				{
					Grant-CsPinPolicy -Identity $phone.identity -PolicyName $PinPolicyDropDown.SelectedItem.tostring()
				}		
			}
			if ($Global:VoicePolicyFieldChange) 
			{
				if ($VoicePolicyDropDown.Text -eq "<Automatic>") {
					Grant-CsVoicePolicy -Identity $phone.identity -PolicyName $Null
				}
				else
				{
					Grant-CsVoicePolicy -Identity $phone.identity -PolicyName $VoicePolicyDropDown.SelectedItem.tostring()
				}		
			}
			if ($Global:ConferencingPolicyFieldChange) 
			{
				if ($ConferencingPolicyDropDown.Text -eq "<Automatic>") {
					Grant-CsConferencingPolicy -Identity $phone.identity -PolicyName $Null
				}
				else
				{
					Grant-CsConferencingPolicy -Identity $phone.identity -PolicyName $ConferencingPolicyDropDown.SelectedItem.tostring()
				}		
			}
			if ($Global:LineURIFieldChange) 
			{
				#Just in case TEL is there, remove it because we're going to readd it to ensure it's there.
				$LineURIText.Text = $LineURIText.Text -replace "tel:",""
				$LineURIText.Text = "tel:" + $LineURIText.Text
				Set-CSCommonAreaPhone -Identity $phone.identity -LineURI $LineURIText.Text
			}
			if ($Global:DisplayNameFieldChange) 
			{
				Set-CSCommonAreaPhone -Identity $phone.identity -DisplayName $DisplayNameText.Text
				[Microsoft.VisualBasic.Interaction]::MsgBox("You have changed the display name of the phone.  This name is how the phone is recognized by this program.  Please click the Refresh Phone List button before any further changes are made to this phone.",'OKOnly,Information', "Phone Name Changed!")

			}
			if ($Global:DisplayNumberFieldChange) 
			{
				Set-CSCommonAreaPhone -Identity $phone.identity -DisplayNumber $DisplayNumberText.Text
			}
			if ($Global:DescriptionFieldChange) 
			{
				Set-CSCommonAreaPhone -Identity $phone.identity -Description $DescriptionText.Text
			}

		#End ForEach Phone
		}

	}


	$SaveChangesButton.Text = "Save Changes"
	$SaveChangesButton.Enabled=$True

})
$SaveChangesButton.Anchor = 'Bottom, Left'
$objForm.Controls.Add($SaveChangesButton)

$RemovePhoneButton = New-Object System.Windows.Forms.Button
$RemovePhoneButton.Location = New-Object System.Drawing.Size((10 + (($objform.width-50) /7 * 4) ),($objform.height - 90))
$RemovePhoneButton.Size = New-Object System.Drawing.Size(115,25)
$RemovePhoneButton.Text = "Remove Phone"
$RemovePhoneButton.Add_Click(
{
	$RemovePhoneButton.Text = "Removing"
	$RemovePhoneButton.Enabled=$False
	$RefreshPhoneListButton.Enabled=$False

	if ($objListBox.SelectedIndex -gt -1) 
	{
		$totalphones=$objListbox.SelectedItems.count
		$ConfirmIt=[Microsoft.VisualBasic.Interaction]::MsgBox("Are you sure you want to remove $totalphones phone(s)?",'YesNoCancel,Question', "Use at your own risk!")
		if ($ConfirmIt -eq "Yes") 
		{ 
			foreach ($PhoneName in $objListbox.SelectedItems)
			{
				Get-CsCommonAreaPhone $PhoneName | Remove-CsCommonAreaPhone
			}
			sleep 2
			$objListBox.Items.Clear()
			LoadPhones
			
		}
	}
	$RemovePhoneButton.Text = "Remove Phone"
	$RemovePhoneButton.Enabled=$True
	$RefreshPhoneListButton.Text = "Refresh Phone List"
	$RefreshPhoneListButton.Enabled=$True
	
})
$RemovePhoneButton.Anchor = 'Bottom, Left'
$objForm.Controls.Add($RemovePhoneButton)

$NewPhoneButton = New-Object System.Windows.Forms.Button
$NewPhoneButton.Location = New-Object System.Drawing.Size((10 + (($objform.width-50) /7 * 5) ),($objform.height - 90))
$NewPhoneButton.Size = New-Object System.Drawing.Size(115,25)
$NewPhoneButton.Text = "Create New Phone"
$NewPhoneButton.Add_Click(
{
	$NewPhoneButton.Enabled=$False
	AddCommonAreaPhoneForm	
	$NewPhoneButton.Enabled=$True
	$RefreshPhoneListButton.Enabled=$False
	$RefreshPhoneListButton.Text = "Loading..."
	$objListBox.Items.Clear()
	LoadPhones
	$RefreshPhoneListButton.Text = "Refresh Phone List"
	$RefreshPhoneListButton.Enabled=$True

})
$NewPhoneButton.Anchor = 'Bottom, Left'
$objForm.Controls.Add($NewPhoneButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size((10 + (($objform.width-50) /7 * 6) ),($objform.height - 90))
$CancelButton.Size = New-Object System.Drawing.Size(115,25)
$CancelButton.Text = "Quit"
$CancelButton.Add_Click({	$objForm.Close()})
$CancelButton.Anchor = 'Bottom, Left'
$objForm.Controls.Add($CancelButton)


#LyncFix LinkLabel
$LyncFixLinkLabel = New-Object System.Windows.Forms.LinkLabel
$LyncFixLinkLabel.Location = New-Object System.Drawing.Size(10,($objform.height - 60)) 
$LyncFixLinkLabel.Size = New-Object System.Drawing.Size(150,20)
$LyncFixLinkLabel.text = "http://www.lyncfix.com"
$LyncFixLinkLabel.add_Click({Start-Process $LyncFixLinkLabel.text})
$LyncFixLinkLabel.Anchor = 'Bottom, Left'
$objForm.Controls.Add($LyncFixLinkLabel)

$objForm.Add_Shown({$objForm.Activate()})

LoadPhones

[void] $objForm.ShowDialog()


