#-------------------------------------------------------------------------
# Module Hotmail
#-------------------------------------------------------------------------


#-------------------------------------------------------------------------
# Function Grant-WinLiveImplicitAuthorization
#
# Provied access token to access user informatin in the One Drive
# Implicit grant flow: ideal for a public environment where explicit user sign-in and consent is required
#-------------------------------------------------------------------------

Function Grant-WinLiveImplicitAuthorization {

	<#
		.SYNOPSIS
			Provide access token to access user informatin in Windows Live (Outlook, Hotmail, OneDrive, etc)
		.DESCRIPTION
			Our script needs to access user information; so first, we need to allow the user to sign in and give consent. We’ll do so by hosting a WebBrowser control to direct the user to the Microsoft authorization page. After the user successfully signs in and accepts the scope of information the script can access, we get an access token by reading the URI that the user is redirected to. More information about this process can be found on the Live SDK Core Concepts site. Depending on the scenario, one of the following OAuth 2.0 grant flows can be used:
			
			•Implicit grant flow: ideal for a public environment where explicit user sign-in and consent is required
		.PARAMETER ClientID
			ID of the application as registered with Microsoft.  
		.PARAMETER Scope_Permissions
			Specifies permissions the script has on the One Drive
			ReadOnly	= Read Only  (default)
			Update 		= Read/Write
		.Example
			Grant-WinLiveImplicitAuthorization -ClientID '00000000603E0BFE' -Scope_Permissions 'Update'
		.Link
			http://blogs.technet.com/b/heyscriptingguy/archive/2013/07/01/use-powershell-3-0-to-get-more-out-of-windows-live.aspx
	#>

	[CmdLetBinding()]
	param ( [String]$ClientID = "000000004811F237",
			[String]$Scope_Permissions = 'ReadOnly' )

	$RedirectUri = "https://login.live.com/oauth20_desktop.srf"
	$AuthorizeUri = "https://login.live.com/oauth20_authorize.srf"

	switch ( $Scope_Permissions ) {
		'ReadOnly' {
				$Scope = "wl.skydrive"
			}
		'Update' {
				$Scope = "wl.skydrive_update","wl.signin" -join "%20"
			}
	}

	# region - Implicit grant flow

	Add-Type -AssemblyName System.Windows.Forms

	$OnDocumentCompleted = {
	  	if($web.Url.AbsoluteUri -match "access_token=([^&]*)") {
			    $script:AccessToken = $Matches[1]
	    		if($web.Url.AbsoluteUri -match "expires_in=([^&]*)") {
					$script:ValidThru = (get-date).AddSeconds([int]$Matches[1])
			    }
			    $form.Close()
			  }
			elseif($web.Url.AbsoluteUri -match "error=") {
			    $form.Close()
		}
	}

	$web = new-object System.Windows.Forms.WebBrowser -Property @{Width=400;Height=500}
	$web.Add_DocumentCompleted($OnDocumentCompleted)
	$form = new-object System.Windows.Forms.Form -Property @{Width=400;Height=500}
	$form.Add_Shown({$form.Activate()})
	$form.Controls.Add($web)
	$web.Navigate("$AuthorizeUri`?client_id=$ClientID&scope=$Scope&response_type=token&redirect_uri=$RedirectUri")

	$null = $form.ShowDialog()

	# endregion
	
	Return $AccessToken

}

#-------------------------------------------------------------------------
# Function Grant-WinLiveCodeAuthorization
#
# Provied access token to access user informatin in Windows Live
# Authorization code grant flow: ideal for automation in a safe environment
#-------------------------------------------------------------------------

Function Grant-WinLiveCodeAuthorization {
	
	<#
		.SYNOPSIS
			Provied access token to access user informatin in the One Drive
		.DESCRIPTION
			Our script needs to access user information; so first, we need to allow the user to sign in and give consent. We’ll do so by hosting a WebBrowser control to direct the user to the Microsoft authorization page. After the user successfully signs in and accepts the scope of information the script can access, we get an access token by reading the URI that the user is redirected to. More information about this process can be found on the Live SDK Core Concepts site. Depending on the scenario, one of the following OAuth 2.0 grant flows can be used:
			
			•Authorization code grant flow: ideal for automation in a safe environment
		.PARAMETER ClientID
			ID of the application as registered with Microsoft.  
		.PARAMETER Secret
			
		.Example
			Grant-OneDriveCodeAuthorization -ClientID '00000000603E0BFE' 
		.Link
			http://blogs.technet.com/b/heyscriptingguy/archive/2013/07/01/use-powershell-3-0-to-get-more-out-of-windows-live.aspx
	#>
	
	[CmdLetBinding()]
	param( [String]$ClientID = "000000004811F237",
		   [String]$Secret = "8MEucksFZF99nopu3nw55RHSErU5ejQd" )
	
	$RedirectUri = "https://login.live.com/oauth20_desktop.srf"
	$AuthorizeUri = "https://login.live.com/oauth20_authorize.srf"

	$Scope = "wl.skydrive_update","wl.signin","wl.offline_access" -join "%20"

	#region - Authorization code grant flow...

	Add-Type -AssemblyName System.Windows.Forms
	$OnDocumentCompleted = {
		if($web.Url.AbsoluteUri -match "code=([^&]*)") {
			    $script:AuthCode = $Matches[1]
			    $form.Close()
			}
			elseif($web.Url.AbsoluteUri -match "error=") {
			    $form.Close()
	  }
	}

	$web = new-object System.Windows.Forms.WebBrowser -Property @{Width=400;Height=500}
	$web.Add_DocumentCompleted($OnDocumentCompleted)
	$form = new-object System.Windows.Forms.Form -Property @{Width=400;Height=500}
	$form.Add_Shown({$form.Activate()})
	$form.Controls.Add($web)

	# Request Authorization Code
	$web.Navigate("$AuthorizeUri`?client_id=$ClientID&scope=$Scope&response_type=code&redirect_uri=$RedirectUri")
	$null = $form.ShowDialog()

	# Request AccessToken
	$Response = Invoke-RestMethod -Uri "https://login.live.com/oauth20_token.srf" -Method Post -ContentType "application/x-www-form-urlencoded" -Body "client_id=$ClientID&redirect_uri=$RedirectUri&client_secret=$Secret&code=$AuthCode&grant_type=authorization_code"
	$AccessToken = $Response.access_token
	$ValidThru = (get-date).AddSeconds([int]$Response.expires_in)
	$RefreshToken = $Response.refresh_token

	#endregion

	#Apparently, the previous snippet discloses the client secret, so it should only be used in a secure environment. When the time comes to refresh the access token, another Rest method needs to be called:

	# Refresh AccessToken
	$Response = Invoke-RestMethod -Uri "https://login.live.com/oauth20_token.srf" -Method Post -ContentType "application/x-www-form-urlencoded" -Body "client_id=$ClientID&redirect_uri=$RedirectUri&grant_type=refresh_token&refresh_token=$RefreshToken"
	$AccessToken = $Response.access_token
	$ValidThru = (get-date).AddSeconds([int]$Response.expires_in)
	$RefreshToken = $Response.refresh_token

		

}

#-------------------------------------------------------------------------
# Function Close-WinLiveSession
#
# Closes the One Drive Session
#-------------------------------------------------------------------------

Function Close-WinLiveiveSession {

	<#
		.SYNOPSIS
			Closes the One Drive Session
		.DESCRIPTION
			 A script should properly sign out a user at the end of a session. Sending the following web request can help ensure that the user’s data isn’t left unattended
		.Example
			Close-OneDriveSession  
		.Link
			http://blogs.technet.com/b/heyscriptingguy/archive/2013/07/01/use-powershell-3-0-to-get-more-out-of-windows-live.aspx
	#>

	[CmdLetBinding()]
	param()
	
	Invoke-WebRequest "https://login.live.com/oauth20_logout.srf?client_id=$ClientID&redirect_uri=$RedirectUri"
}

#-------------------------------------------------------------------------
# Function Get-OneDriveItem
#
# List Items in a folder
#-------------------------------------------------------------------------

Function Get-OneDriveItem {

	<#
		.SYNOPSIS
			List items in a One Drive folder
		.DESCRIPTION
			List Items in a One Drive Folder
		.Example
			Get-OneDriveItem 
		.Link
			http://blogs.technet.com/b/heyscriptingguy/archive/2013/07/02/use-powershell-to-work-with-skydrive-for-powerful-automation.aspx
	#>

	[CmdLetBinding()]
	param( [String]$Uri,
		   $AccessToken )
	
#	Invoke-RestMethod -Uri "$ApiUri/me/skydrive/my_documents?access_token=$AccessToken"
	
}

#-------------------------------------------------------------------------

# TO DO ----  http://blogs.technet.com/b/heyscriptingguy/archive/2013/07/03/using-the-live-rest-api-with-powershell.aspx


#-------------------------------------------------------------------------
# Main
#-------------------------------------------------------------------------

Export-ModuleMember -Function Grant-WinLiveImplicitAuthorization, Grant-WinLiveCodeAuthorization, Close-WinLiveSession