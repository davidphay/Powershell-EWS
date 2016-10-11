
function Connect-EWS {
<#
	.SYNOPSIS
		Connects to a Exchange server throught EWS.
	
	.DESCRIPTION
		Connects to a Exchange server throught EWS.
	
	.PARAMETER UserMailAddress
		Specify the mailbox that you want to connect to a user account. You can use the followind values to identify the mailbox:
		email address
	
	.PARAMETER ModuleDllPath
		Define the path of EWS module.
		You must download and install him before
		
		Download link: https://www.microsoft.com/en-eg/download/details.aspx?id=42951
	
	.PARAMETER ServerVersion
		Define ServerVersion if AutoDetect failed.
		Possible value:
		Exchange2007_SP1
		Exchange2010
		Exchange2010_SP1
		Exchange2010_SP2
		Exchange2013
		Exchange2013_SP1
	
	.PARAMETER EwsUrl
		Define EWS Url if AutoDetect failed
	
	.PARAMETER Credential
		Define Credential who are authorized to read the mailbox
	
	.EXAMPLE
		PS C:\> $email = "myemail@mydomain.com"
		PS C:\> $credentials = @{
		"Username" = $AuthUser;
		"Password" = $password;
		"Domain" = $Domain
		}
		PS C:\> $cred = New-Object -TypeName PSObject -Property $credentials
		PS C:\> Connect-EWS -UserMailAddress $email -Credential $cred
		
		.VERSION
		1.0.0 - 2016.08.09
		Initial version
		
		1.1.0 - 2016.09.28
		Change download EWS url
		
		2.0.0 - 2016.10.06
		Credential param is now [PSCredential]
		To define Credential Just call Get-Credential
		If UserMailAddress 
		
		.VALIDATION
		Exchange 2013
	
	.OUTPUTS
		Microsoft.Exchange.WebServices.Data.ExchangeServiceBase
	
	.NOTES
		Additional information about the function.
#>
	
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true)]
		[ValidatePattern('^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}$|^AutoDetect$')]
		[Alias('MailAddress')]
		[String]$UserMailAddress,
		[Alias('DllPath')]
		[String]$ModuleDllPath = "$env:SystemDrive\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll",
		[ValidatePattern('^Exchange[0-9]{4}.{0,4}$|^AutoDetect$')]
		[String]$ServerVersion = 'AutoDetect',
		[ValidatePattern('^https?://[^/]*/ews/exchange.asmx$|^AutoDetect$')]
		[Alias('Url')]
		[String]$EwsUrl = "AutoDetect",
		[Parameter(Mandatory = $true)]
		[PSCredential]$Credential
	)
	
	Begin {
		Try {
			# Loading Module: Microsoft.Exchange.WebService ( if not loaded yet )
			If (-not (Get-Module -Name:Microsoft.Exchange.WebServices)) {
				Try {
					Write-Debug -Message "Import Module $ModuleDllPath"
					Import-Module -Name:$ModuleDllPath -ErrorAction:Stop
				} Catch [System.IO.FileNotFoundException] {
					Throw [System.IO.FileNotFoundException] "$_`nhttps://www.microsoft.com/en-eg/download/details.aspx?id=42951"
				} Catch {
					Throw [System.SystemException] "Loading module - $ErrorMessage"
				}
			}
			
			if (!$UserMailAddress) {
				$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
				$UserMailAddress = [ADSI]"LDAP://<SID=$sid>"
			}
			
			# initializing EWS ExchangeService
			If ($ServerVersion -eq "AutoDetect") {
				Write-Debug -Message "ServerVersion is AutoDetect"
				$ExchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
			} Else {
				Write-Debug -Message "ServerVersion is $ServerVersion"
				$ExchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ServerVersion)
			}
			
			Write-Debug -Message "RequestedServerVersion: $($ExchangeService.RequestedServerVersion)"
			Write-Debug -Message "UserAgent: $($ExchangeService.UserAgent)"
			
			# Define Credential
			$ExchangeService.Credentials = New-Object System.Net.NetworkCredential($Credential.UserName.ToString(), $Credential.GetNetworkCredential().password.ToString())
			Write-Debug -Message "Credentials: $($ExchangeService.Credentials.Credentials.UserName)"
		} Catch {
			Write-Warning $_.Exception.GetType().FullName;
			Throw [System.SystemException]$_.Exception.Message
		}
	}
	Process {
		Try {
			# Define Autodiscover url
			If ($EwsUrl -eq "AutoDetect") {
				Write-Debug -Message "AutodiscoverUrl($UserMailAddress)"
				$ExchangeService.AutodiscoverUrl($UserMailAddress)
			} Else {
				$ExchangeService.Url = New-Object Uri($EwsUrl)
			}
			Write-Debug -Message "EwsUrl: $($ExchangeService.Url)"
		} Catch {
			$ErrorMessage = $_.Exception.Message
			$message1 = 'Exception calling "AutodiscoverUrl" with "1" argument\(s\): "Autodiscover blocked a potentially insecure redirection to'
			$message2 = 'Exception calling "AutodiscoverUrl" with "1" argument\(s\): "The Autodiscover service returned an error."'
			
			If ($ErrorMessage -match $message1) {
				# When the credential is wrong (Username, Domain or password)
				Throw [System.Security.Authentication.InvalidCredentialException] "Credential error"
			} Elseif ($ErrorMessage -match $message2) {
				# When the email address does not exist
				Write-Error -Message "Email addresse does not exist($UserMailAddress)"
				Throw [System.InvalidOperationException] "Email address does not exist"
			} Else {
				Throw [System.SystemException]$_.Exception.Message
			}
		}
	}
	End {
		Try {
			# try to list anything in mailbox to validate access
			$rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $UserMailAddress)
			$folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind([Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService, $rootFolderId) | out-null
			
			Write-Debug -Message "Return ExchangeService"
			Return [Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$ExchangeService
		} Catch [System.Management.Automation.MethodInvocationException] {
			$ErrorMessage = $_.Exception.Message
			$message1 = "Exception calling `"Bind`" with `"2`" argument\(s\): `"The request failed. The remote name could not be resolved:"
			
			If ($ErrorMessage -match $message1) {
				# When the EWS url is wrong
				$url = ($ErrorMessage -split "'")[1]
				Throw [System.IO.IOException] "The remote name could not be resolved ($url)"
			} Else {
				Throw [System.SystemException]$_.Exception.Message
			}
		} Catch {
			Throw [System.SystemException]$_.Exception.Message
		}
	}
}

function Get-EWSFolder {
<#
	.SYNOPSIS
		Returns a Folder object corresponding to the folder in a specified path.
	
	.DESCRIPTION
		Returns a Folder object corresponding to the folder in a specified path.
	
	.PARAMETER MailboxName
		Specify the mailbox that you want to connect to a user account. You can use the followind values to identify the mailbox:
		email address
	
	.PARAMETER Path
		Enter the path of the folder
	
	.PARAMETER WellKnownFolderName
		Define the base search
	
	.PARAMETER Service
		Call Connect-EWS function
	
	.PARAMETER List
		A description of the List parameter.
	
	.PARAMETER Credential
		Define Credential who are authorized to read the mailbox (PSObject not PSCredential)
	
	.EXAMPLE
		PS C:\> $email = "myemail@mydomain.com"
		PS C:\> $credentials = @{
					"Username" = $AuthUser;
					"Password" = $password;
					"Domain" = $Domain
				}
		PS C:\> $cred = New-Object -TypeName PSObject -Property $credentials
		PS C:\> $service = Connect-EWS -UserMailAddress $email -Credential $cred -Verbose
		PS C:\> Get-FolderEWS -MailboxName $email -Path "parent\child1\child11" -Service $service
		
	.VERSION
		1.0.0 - 2016.08.09
			Initial Version
		
		1.1.0 - 2016.09.26
			Add List param to list all folder in define path
	
		1.2.0 - 2016.09.29
			Path param is not mandatory now, if is not setted you get WellKnown Folder Name
			Add folder's Size even if there are children folder
	
		1.3.0 - 2016.10.05
			Remove MailboxName param & update $rootFolderId
		
	.VALIDATION
		Exchange 2013
	
	.NOTES
		Next release : Return Folder ID
#>
	
	[CmdletBinding(ConfirmImpact = 'None')]
	param
	(
		[Parameter(Mandatory = $false)]
		[Alias('FullPath')]
		[String]$Path = $null,
		[ValidateSet('Calendar', 'Contacts', 'Inbox', 'SentItems', 'MsgFolderRoot', 'PublicFoldersRoot', 'Root', 'SearchFolders', 'ArchiveRoot', 'ArchiveMsgFolderRoot')]
		[String]$WellKnownFolderName = "Inbox",
		[Parameter(Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$Service,
		[Parameter(Mandatory = $false)]
		[Switch]$List
	)
	
	Begin {
		Try {
			if ($Path) {
				Write-Debug -Message "Try to search $Path"
				
				# Split the path to search recursively
				$arrPath = $Path.Split("\")
			} Else {
				Write-Debug -Message "Try to List Folder in $WellKnownFolderName"
			}
			
			# Initialize Variables
			New-Variable -Name concatpath -Value ""
			
			# Get all the folders in the message's root folder.
			$rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$WellKnownFolderName)
			$folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $rootFolderId)
		} Catch {
			Write-Error -Message $_.Exception.Message -ErrorAction Stop
		}
	}
	Process {
		Try {
			# Represents the view settings in a folder search operation: maximum of returned folders = 1
			$folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
			
			#Define Extended properties  
			$PR_FOLDER_PATH = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
			
			$PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
			$PropertySet.Add($PR_FOLDER_PATH);
			$folderView.PropertySet = $PropertySet;
			if ($Path) {
				if ($arrPath.count -eq 1) {
					# Represents a search filter that determines wheter a property is equal to a given value or other property.
					$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $Path)
					$findFolderResults = $service.FindFolders($folder.Id, $searchFilter, $folderView)
					
					If ($findFolderResults.TotalCount -gt 0) {
						$folder = $findFolderResults.Folders[0]
						$return = $true
						
						Write-Debug -Message "Folder \$($folder.DisplayName) was found"
					} Else {
						Throw [System.IO.IOException] "Folder Not found"
					}
				} Elseif ($arrPath.count -gt 1) {
					For ($i = 0; $i -lt $arrPath.count; $i++) {
						# Represents a search filter that determines wheter a property is equal to a given value or other property.
						$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $arrPath[$i])
						$findFolderResults = $service.FindFolders($folder.Id, $searchFilter, $folderView)
						
						If ($findFolderResults.TotalCount -gt 0) {
							$folder = $findFolderResults.Folders[0]
							$return = $true
							
							$concatpath += "\$($arrPath[$i])"
							
							Write-Debug -Message "Folder $concatpath was found"
						} Else {
							Throw [System.IO.IOException] "Folder Not found"
						}
					}
				}
				
				$folderPath = $null
				# Properties is in ExtendedProperties
				# To keep the value $item.ExtendedProperties.value
				Write-Debug "Add Folder Path in Extended Properties"
				$folder.TryGetProperty($PR_FOLDER_PATH, [ref]$folderPath) | Out-Null
			} Else {
				Write-Debug "The path is not define..."
			}
		} Catch [System.IO.IOException] {
			Throw [System.IO.IOException] "Folder Not found"
		} Catch {
			Throw [System.SystemException]$_.Exception.Message
		}
	}
	End {
		Try {
			if ($List) {
				$folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
				
				#Define Extended properties  
				$PR_MESSAGE_SIZE_EXTENDED = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
				$PR_FOLDER_PATH = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
				
				$psPropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
				$psPropertySet.Add($PR_MESSAGE_SIZE_EXTENDED);
				$psPropertySet.Add($PR_FOLDER_PATH);
				$folderView.PropertySet = $psPropertySet;
				
				Try {
					if ($Path) {
						$folders = $service.FindFolders($findFolderResults.Id, $folderView)
					} Else {
						$folders = $service.FindFolders($folder.Id, $folderView)
					}
					
					if (($folders | Measure-Object).count -gt 0) {
						$folders.Folders | ForEach-Object{
							[Int]$folderSize = 0
							
							#Deep Transval will ensure all folders in the search path are returned  
							$folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;
							$findFolderResults = $service.FindFolders($_.Id, $folderView)
							
							if (($findFolderResults | Measure-Object).count -eq 0) {
								$_.TryGetProperty($PR_MESSAGE_SIZE_EXTENDED, [ref]$folderSize) | Out-Null
							} Else {
								(($findFolderResults.ExtendedProperties | Where-Object{ $_.PropertyDefinition.Tag -eq 3592 }).Value) | ForEach-Object{
									$folderSize += $_
								}
							}
							
							# Add folder size in properties in MB
							$_ | Add-Member NoteProperty -Name "FolderSize(MB)" -Value ([Math]::Round($folderSize/1MB, 2, [MidPointRounding]::AwayFromZero)) -Force
						}
					}
				} Catch [System.Management.Automation.MethodInvocationException]{
					Write-Warning "Their is no folder"
				} Catch {
					Write-Error $_.Exception.Message
				}
				
				Return [Microsoft.Exchange.WebServices.Data.FindFoldersResults]$folders | Select-Object DisplayName, "FolderSize(MB)", ChildFolderCount, @{ Name = "MessageCount"; Expression = { $_.TotalCount } }
			} Else {
				Return [Microsoft.Exchange.WebServices.Data.Folder]$folder
			}
		} Catch {
			Throw [System.SystemException]$_.Exception.Message
		}
	}
}

function Get-EWSMail {
<#
	.SYNOPSIS
		Use Get-MailEWS to view mail user.
	
	.DESCRIPTION
		Use Get-MailEWS to view mail user.

	.PARAMETER GetFolder
		Call Get-FolderEWS function

	.EXAMPLE
		PS C:\> $folder = Get-FolderEWS -MailboxName $email -FullPath "parent\child1\child11" -Service $service
		PS C:\> Get-MailEWS -GetFolder $Folder
		
		
	.VERSION
		1.0.0 - 2016.08.09
			Initial version
		
	.VALIDATION
		Exchange 2013
	
	.OUTPUTS
		System.Object
	
	.NOTES
		Limitation at 2000 item
#>
	
	[OutputType([System.Object])]
	Param
	(
		[Parameter(Mandatory = $True)]
		[Microsoft.Exchange.WebServices.Data.ServiceObject]$Folder,
		[Switch]$Full,
		[Switch]$WithBody
	)
	
	Process {
		Try {
			Write-Debug -Message "$($Folder.DisplayName) - Retrieve mail list"
			
			#list the first 2000 items who match
			$mails = $Folder.FindItems(2000)
		} Catch {
			Write-Error -Message $_.Exception.Message
		}
	}
	End {
		Try {
			if ($mails -ne 0) {
				if ($WithBody) {
					Write-Debug "Load Body"
					$mails.load()
				}
				
				# Return mails
				Write-Debug -Message "Return email object"
				if ($Full) {
					Return [System.Object]$mails
				} Else {
					if ($WithBody) {
						Return [System.Object]$mails | Select-Object Subject, From, IsRead, IsAttachment, DateTimeReceived, DisplayTo, DisplayCC, @{ Name = "Body"; Expression = { $_.Body.text } } -
					} Else {
						Return [System.Object]$mails | Select-Object Subject, From, IsRead, IsAttachment, DateTimeReceived, DisplayTo, DisplayCC
					}
				}
			} Else {
				Throw [System.IO.IOException] "0 was founded"
			}
		} Catch [System.IO.IOException] {
			Write-Warning -Message "0 mail was founded" 
			Return [System.Object]$mails
		} Catch {
			Write-Error -Message $_.Exception.Message
		}
	}
}

function Move-EWSMail {
<#
	.SYNOPSIS
		Move an email to another folder (in same mailbox)
	
	.DESCRIPTION
		Move an email to another folder (in same mailbox)
	
	.PARAMETER Mail
		put the mail you want to move, it's an EWS Object [System.Object]
		ex: Get-MailEWS -GetFolder $Folder
	
	.PARAMETER Destination
		Define the destination folder, it's an EWS Object [Microsoft.Exchange.WebServices.Data.ServiceObject]
	
	.PARAMETER Service
		Call Connect-EWS function
	
	.PARAMETER Test
		Swith param define if Test mode is enable

	.EXAMPLE
		PS C:\> $EmailAddress = "myemail@mydomain.com"
		PS C:\> $service = Connect-EWS -UserMailAddress $EmailAddress -Credential $Credential
		PS C:\> $Source = Get-FolderEWS -MailboxName $EmailAddress -FolderPath $SourcePath -Service $service
		PS C:\> $mails = Get-MailEWS -GetFolder $Source
		PS C:\> $limithour = 7
		PS C:\> $Target = Get-FolderEWS -MailboxName $EmailAddress -FolderPath $DestPath -Service $service
		PS C:\> Move-MailEWS -Mail $mails -TargetFolder $Target -Service $service -Hours $limithour -Test:$Test
		
	.VERSION
		1.0.0 - 2016.08.09
			Initial version
		
	.VALIDATION
		Exchange 2013
	
	.NOTES
		Additional information about the function.
#>
	
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true)]
		[System.Object]$Mail,
		[Parameter(Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ServiceObject]$Destination,
		[Parameter(Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$Service,
		[Switch]$Test
	)
	
	Process {
		# If mail is not null, it means their is 1 or more mails
		If ($Mail) {
			$Mail | ForEach-Object{
				try {
					# Move the Message
					If (-not $Test) {
						Write-Debug -Message "Move $($_.Subject) to '$($Destination.FullPath)''"
						$_.Move($Destination.Id)
					} Else {
						Write-Debug -Message "TEST - Move $($_.Subject) to '$($Destination.FullPath)''"
					}
				} Catch {
					Write-Error -Message $_.Exception.Message -ErrorAction Continue
				}
			}
		} Else {
			Write-Warning -Message "No email are moved"
		}
	}
}

function Remove-EWSMail {
<#
	.SYNOPSIS
		Use Remove-MailEWS cmdlet to delete existing email.
	
	.DESCRIPTION
		Use Remove-MailEWS cmdlet to delete existing email.
	
	.PARAMETER Mail
		Enter the email's you want to remove, it an email object
		ex: Get-MailEWS -GetFolder $Folder
	
	.PARAMETER Service
		Call Connect-EWS function
	
	.PARAMETER DeleteMode
		Define the Delete mode.
		By default it's a MoveToDeletedItems, but you select : HardDelete,MoveToDeletedItems,SoftDelete:
			- The item or folder will be permanently deleted.
			- The item or folder will be moved to the mailbox's Deleted Items folder.
			- The item or folder will be moved to the dumpster. Items and folders in the dumpster can be recovered.
	
	.PARAMETER Test
		Swith param define if Test mode is enable
	
	.EXAMPLE
		PS C:\> Remove-MailEWS -Mail $mails -Service $service
	
	.EXAMPLE
		PS C:\> Remove-MailEWS -Mail $mails -Service $service -DeleteMode HardDelete
		
	.VERSION
		1.0.0 - 2016.08.09
		Initial version

	.VALIDATION
		Exchange 2013
	
	.NOTES
		Delete mode: https://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.deletemode(v=exchg.80).aspx
#>
	
	Param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true)]
		[psobject]$Mail,
		[Parameter(Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$Service,
		[ValidateSet('HardDelete', 'SoftDelete', 'MoveToDeletedItems')]
		[String]$DeleteMode = "MoveToDeletedItems",
		[Switch]$Test
	)
	
	Try {
		# If mail is not null, it means their is 1 or more mails
		If ($Mail) {
			# Delete message
			If (-not $test) {
				[void]$Mail.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::$DeleteMode)
				Write-Debug -Message "Delete $($Mail.Subject)"
			} Else {
				Write-Debug -Message "TEST - Delete $($Mail.Subject)"
			}
		} Else {
			Write-Debug -Message "No email must be deleted"
		}
	} Catch {
		Write-Error -Message $_.Exception.Message -ErrorAction Continue
	}
}

function Get-EWSCalendar {
<#
	.SYNOPSIS
		Retrieve a list of calendar from the GAL
	
	.DESCRIPTION
		Retrieve one or more calendar form the GAL (search with a wildcard...)
	
	.PARAMETER CalendarName
		Enter the name of the calendar. the value can be:
		- Mailbox Name
		- Mailbox Address
		- Display Name
	
	.PARAMETER Service
		Call Connect-EWS function
	
	.EXAMPLE
		PS C:\> Get-CalendarEWS -CalendarName 'John Doe' -Service $Service
		=> Retrieve all calendar who match with John Doe
		
	.VERSION
		1.0.0 - 2016.08.23
		Initial version
		
	.VALIDATION
		Exchange 2013
	
	.OUTPUTS
		Array
	
	.NOTES
		For Calendar Query:
		https://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.resolvenamesearchlocation(v=exchg.80).aspx
	
	.INPUTS
		String, Microsoft.Exchange.WebServices.Data.ExchangeServiceBase
#>
	
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true)]
		[Alias('Calendar')]
		[String]$CalendarName,
		[Parameter(Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$Service
	)
	
	Begin {
		Write-Debug -Message "Try to search $CalendarName"
		[Array]$calendarArray = @()
	}
	Process {
		Try {
			# Search calendar, in GAL (Directory Only), with full details ($true)
			$calendar = $Service.ResolveName($CalendarName, "DirectoryOnly", $true)
			
			If ($calendar.count -ne 0) {
				# If the number of calendar is greather than 0, we have more than 1 calendar
				Write-Debug -Message "CalendarEWS - $($calendar.count) founded"
				
				# Foreach calendar, add the mailbox Name & Address to the Contrat Properties
				$calendar | ForEach-Object {
					$Object = New-Object -TypeName PSObject
					
					$Object = $_.Contact | Select-Object DisplayName, GivenName, CompanyName, ContactSource, Department, JobTitle, Manager, OfficeLocation, Surname, Culture, ExtendedProperties, Categories
					$Object | Add-Member NoteProperty -Name MailboxName -Value $_.Mailbox.Name -Force
					$Object | Add-Member NoteProperty -Name MailboxAddress -Value $_.Mailbox.Address -Force
					
					$calendarArray += $Object
				}
			} Else {
				Throw [System.AccessViolationException] "0 Calendar was founded"
			}
		} Catch [System.AccessViolationException]{
			Throw [System.AccessViolationException] "0 Calendar was founded"
		} Catch [System.Management.Automation.MethodInvocationException]{
			$ErrorMessage = $_.Exception.Message
			$message1 = "Exception calling `"ResolveName`" with `"3`" argument\(s\): `"The request failed. The remote name could not be resolved:"
			
			If ($ErrorMessage -match $message1) {
				# When the EWS url is wrong
				$url = ($ErrorMessage -split "'")[1]
				Throw [System.IO.IOException] "The remote name could not be resolved ($url)"
			} Else {
				Throw [System.SystemException] $_.Exception.Message
			}
		} Catch {
			Write-Error -Message $_.Exception.Message
		}
	}
	End {
		Try {
			# Return an Array beacause we may have 1 or more calendar
			Return [Array]$calendarArray
		} Catch {
			Write-Error -Message $_.Exception.Message
		}
	}
}

function Get-EWSCalendarPermission {
<#
	.SYNOPSIS
		A brief description of the Get-EWSPermission function.
	
	.DESCRIPTION
		A detailed description of the Get-EWSPermission function.
	
	.PARAMETER Path
		Enter the path of the folder
	
	.PARAMETER WellKnownFolderName
		must remove some validateset ?
	
	.PARAMETER Details
		To show full details
	
	.PARAMETER Service
		Call Connect-EWS function
	
	.EXAMPLE
		PS C:\> Get-EWSPermission -Service $service
		
		Name                       DisplayPermissionLevel
		----                       ----------------------
		Default                                  Reviewer
		Anonymous                        FreeBusyTimeOnly
	
	.EXAMPLE
		PS C:\> Get-EWSPermission -Service $service -Details
		
		Name             : Default
		PermissionLevel  : Reviewer
		Read             : FullDetails
		Edit             : None
		CreateItems      : False
		CreateSubFolders : False
		DeleteItems      : None
		FolderOwner      : False
		FolderContact    : False
		FolderVisible    : True
		
		Name             : Anonymous
		PermissionLevel  : FreeBusyTimeOnly
		Read             : TimeOnly
		Edit             : None
		CreateItems      : False
		CreateSubFolders : False
		DeleteItems      : None
		FolderOwner      : False
		FolderContact    : False
		FolderVisible    : False
		
	.VERSION
		1.0.0 - First version
	
	.NOTES
		For more information about advanced functions, call Get-Help with any
		of the topics in the links listed below.
#>
	
	[CmdletBinding(ConfirmImpact = 'None')]
	param
	(
		[Parameter(Mandatory = $false)]
		[Switch]$Details,
		[Parameter(Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$Service
	)
	
	Begin {
		Try {
			$Permissions = @()
			# Get all the folders in the message's root folder.
			$rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
			$folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $rootFolderId)
			
		} Catch {
			Write-Error -Message $_.Exception.Message -ErrorAction Stop
		}
	}
	Process {
		Try {
			# On folder you can have many persmission
			$folder.Permissions | ForEach-Object{
				If ($_.UserId.PrimarySmtpAddress) {
					$Name = $_.UserId.PrimarySmtpAddress
				} else {
					$Name = $_.UserId.StandardUser.ToString()
				}
				$_ | Add-Member NoteProperty -Name Name -Value $Name -Force
				
				$Permissions += $_
			}
		} Catch {
			Throw [System.SystemException]$_.Exception.Message
		}
	}
	End {
		Try {
			if ($Details) {
				Return $Permissions
			} Else {
				Return $Permissions | Select-Object Name, DisplayPermissionLevel
			}
		} Catch {
			Throw [System.SystemException]$_.Exception.Message
		}
	}
}

function Set-EWSCalendarPermission {
<#
	.SYNOPSIS
		Define the level permission for a user or a group
	
	.DESCRIPTION
		Define the level permission for a user or a group
	
	.PARAMETER UserAdress
		ets the identifier of the user that the permission applies to. 
	
	.PARAMETER Permissionlevel
		Sets the permission level. 
	
	.PARAMETER Service
		Call Connect-EWS function
	
	.PARAMETER Force
		A description of the Force parameter.
	
	.PARAMETER WhatIf
		Shows what would happen if the cmdlet runs. The cmdlet is not run.
	
	.PARAMETER CanCreateItems
		Gets or sets a value that indicates whether the user can create new items. 
	
	.PARAMETER CanCreateSubFolders
		Sets a value that indicates whether the user can create subfolders. 
	
	.PARAMETER IsFolderOwner
		Sets a value that indicates whether the user owns the folder. 
	
	.PARAMETER IsFolderVisible
		Sets a value that indicates whether the folder is visible to the user. 
	
	.PARAMETER IsFolderContact
		Sets a value that indicates whether the user is a contact for the folder. 
	
	.PARAMETER EditItems
		Sets a value that indicates whether the user can edit existing items. 
	
	.PARAMETER DeleteItems
		Sets a value that indicates whether the user can delete existing items. 
	
	.PARAMETER ReadItems
		Sets the read items access permission. 
	
	.EXAMPLE
		PS C:\> Set-EWSCalendarPermission -UserAdress email@address.com -Permissionlevel Editor -Service $service
		
		.VERSION
		1.0.0 - First version
		1.1.0 - Add Custom Permission
	
	.NOTES
		https://msdn.microsoft.com/en-us/library/office/dn641962(v=exchg.150).aspx
		https://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.folderpermission_properties(v=exchg.80).aspx
#>
	
	[CmdletBinding(DefaultParameterSetName = 'PermissionLevel',
				   ConfirmImpact = 'Medium')]
	param
	(
		[Parameter(Mandatory = $true)]
		[System.String]$UserAdress,
		[Parameter(ParameterSetName = 'PermissionLevel',
				   Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.FolderPermissionLevel]$Permissionlevel,
		[Parameter(Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$Service,
		[Switch]$Force,
		[Switch]$WhatIf,
		[Parameter(ParameterSetName = 'CustomPermission')]
		[Boolean]$CanCreateItems,
		[Parameter(ParameterSetName = 'CustomPermission')]
		[Boolean]$CanCreateSubFolders,
		[Parameter(ParameterSetName = 'CustomPermission')]
		[Boolean]$IsFolderOwner,
		[Parameter(ParameterSetName = 'CustomPermission')]
		[Boolean]$IsFolderVisible,
		[Parameter(ParameterSetName = 'CustomPermission')]
		[Boolean]$IsFolderContact,
		[Parameter(ParameterSetName = 'CustomPermission')]
		[Microsoft.Exchange.WebServices.Data.PermissionScope]$EditItems,
		[Parameter(ParameterSetName = 'CustomPermission')]
		[Microsoft.Exchange.WebServices.Data.PermissionScope]$DeleteItems,
		[Parameter(ParameterSetName = 'CustomPermission')]
		[Microsoft.Exchange.WebServices.Data.FolderPermissionReadAccess]$ReadItems
	)
	
	Begin {
		Try {
			# Get all the folders in the message's root folder.
			$rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
			$folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $rootFolderId)
		} Catch {
			Write-Error -Message $_.Exception.Message -ErrorAction Stop
		}
	}
	Process {
		Try {
			$folder.Permissions | ForEach-Object{
				If ($_.UserId.PrimarySmtpAddress) {
					$Name = $_.UserId.PrimarySmtpAddress
				} else {
					$Name = $_.UserId.StandardUser.ToString()
				}
				
				if ($Name -eq $UserAdress) {
					if ($_.PermissionLevel -eq $Permissionlevel) {
						$Status = 0
						Return $Status | Out-Null
					} else {
						$Status = 1
						$perm = $_
						Return $Status | Out-Null
					}
				}
			}
		} Catch {
			Write-Error -Message $_.Exception.Message -ErrorAction Stop
		}
	}
	End {
		Try {
			# If Status :
			# = 0 => User and permission still exist
			# = 1 => User exist but the permission level is different
			# Default => User not exist
			switch ($status) {
				'0' {
					# the new permission level is the same
					Write-Warning "The Permission still exist for this user"
					Return
				}
				'1'{
					if (!$Force) {
						Write-Host "The Force parameter was not specified. If you continue, the permission level will be updated. Are you sure you want to continue?"
						do {
							$return = Read-Host "[Y] Yes  [N] No (default is 'Y')"
						} until ($return -eq "Y" -or $return -eq "N")
						
						
						switch ($return) {
							'N'{
								Return
							}
						}
					}
					
					# the new permission is different
					# Apply new permission
					Write-Warning "The existing permission will be overriding"
					Write-Debug "Remove existing Permission"
					$folder.Permissions.Remove($perm) | Out-Null
				}
				default {
					# User have no permission... Create new permission
					Write-Debug "Create the permission for $UserAdress"
				}
			}
			
			If (!$WhatIf) {
				if ($PSCmdlet.ParameterSetName -ne "CustomPermission") {
					$NewPermission = New-Object Microsoft.Exchange.WebServices.Data.FolderPermission($UserAddress, $Permissionlevel)
				} Else {
					$NewPermission = New-Object Microsoft.Exchange.WebServices.Data.FolderPermission
					$NewPermission.UserId = $UserAddress
					
					if ($CanCreateItems) {
						$NewPermission.CanCreateItems = $CanCreateItems
					}
					if ($CanCreateSubFolders) {
						$NewPermission.CanCreateSubFolders = $CanCreateSubFolders
					}
					if ($IsFolderOwner) {
						$NewPermission.IsFolderOwner = $IsFolderOwner
					}
					if ($IsFolderVisible) {
						$NewPermission.IsFolderVisible = $IsFolderVisible
					}
					if ($IsFolderContact) {
						$NewPermission.IsFolderContact = $IsFolderContact
					}
					if ($DeleteItems) {
						$NewPermission.DeleteItems = $DeleteItems
					}
					if ($ReadItems) {
						$NewPermission.ReadItems = $ReadItems
					}
					if ($EditItems) {
						$NewPermission.EditItems = $EditItems
					}
					
				}
				$folder.Permissions.Add($NewPermission)
				$folder.Update()
			} Else {
				Write-Host "What if: Performing the operation `"Set-EWSCalendarPermission`" for $UserAddress on current mailbox"
			}
		} Catch [System.Management.Automation.MethodInvocationException]{
			Throw [system.ArgumentException] "User was nor valid"
		} Catch {
			Write-Error -Message $_.Exception.Message -ErrorAction Stop
		}
	}
}

function Remove-EWSCalendarPermission {
<#
	.SYNOPSIS
		A brief description of the Remove-EWSCalendarPermission function.
	
	.DESCRIPTION
		A detailed description of the Remove-EWSCalendarPermission function.
	
	.PARAMETER UserAddress
		A description of the UserAddress parameter.
	
	.PARAMETER Service
		Call Connect-EWS function
	
	.PARAMETER Force
		Force the cmdlet
	
	.PARAMETER WhatIf
		Shows what would happen if the cmdlet runs. The cmdlet is not run.
	
	.PARAMETER UserAdress
		Define the user
	
	.EXAMPLE
		PS C:\> Remove-EWSCalendarPermission -UserAddress email@address.com -Service $service
	
	.EXAMPLE
		PS C:\> Remove-EWSCalendarPermission -UserAddress email@address.com -Service $service -Force
		
	.VERSION
		1.0.0 - First version
	
	.NOTES
		For more information about advanced functions, call Get-Help with any
		of the topics in the links listed below.
#>
	
	[CmdletBinding(ConfirmImpact = 'Medium')]
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$UserAddress,
		[Parameter(Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$Service,
		[Switch]$Force,
		[Switch]$WhatIf
	)
	
	Begin {
		Try {
			# Get all the folders in the message's root folder.
			$rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
			$folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $rootFolderId)
		} Catch {
			Write-Error -Message $_.Exception.Message -ErrorAction Stop
		}
	}
	Process {
		Try {
			$folder.Permissions | ForEach-Object{
				If ($_.UserId.PrimarySmtpAddress) {
					$Name = $_.UserId.PrimarySmtpAddress
				} else {
					$Name = $_.UserId.StandardUser.ToString()
				}
				
				if ($Name -eq $UserAdress) {
					Write-Host "User found"
					$perm = $_
				}
			}
		} Catch {
			Write-Error -Message $_.Exception.Message -ErrorAction Stop
		}
	}
	End {
		Try {
			if (!$Force) {
				Write-Host "The Force parameter was not specified. If you continue, the user permission will be removed. Are you sure you want to continue?"
				do {
					$return = Read-Host "[Y] Yes  [N] No (default is 'Y')"
				} until ($return -eq "Y" -or $return -eq "N")
				
				
				switch ($return) {
					'N' {
						Return
					}
				}
			}
			
			if (!$WhatIf) {
				Write-Debug "Remove User permission"
				$folder.Permissions.Remove($perm) | Out-Null
				$folder.update()
			} Else {
				Write-Host "What if: Performing the operation `"Remove-EWSCalendarPermission`" for $UserAddress on current mailbox"
			}
		} Catch {
			Write-Error -Message $_.Exception.Message -ErrorAction Stop
		}
	}
}

function Get-EWSMeeting {
<#
	.SYNOPSIS
		Retrieve all meeting for the define period
	
	.DESCRIPTION
		Retrieve all meeting for the define period
	
	.PARAMETER MailboxName
		specify the mailbox that you want to connect to a user account. You can use the following values to identify the mailbox:
		email address
	
	.PARAMETER Service
		Call Connect-EWS function
	
	.PARAMETER StartDate
		Define the start period must be checked
	
	.PARAMETER EndDate
		Define the end period must be checked
	
	.EXAMPLE
		PS C:\> Get-CalendarMeeting -MailboxName "John.Doe@mydomain.com" -Service $Service
	
	.EXAMPLE
		PS C:\> $Start = (get-Date)
		PS C:\> $End =  (get-Date)
		PS C:\> Get-CalendarMeeting -MailboxName "John.Doe@mydomain.com" -Service $Service -StartDate $Start -EndDate $End
		
	.VERSION
		1.0.0 - 23.08.2016
		Initial Version
		
	.VALIDATION
		Exchange 2013
	
	.OUTPUTS
		Array
	
	.NOTES
		For more information about advanced functions, call Get-Help with any
		of the topics in the links listed below.
	
	.INPUTS
		String, Microsoft.Exchange.WebServices.Data.ExchangeServiceBase, Datetime
#>
	
	[CmdletBinding()]
	[OutputType([psobject])]
	Param
	(
		[Parameter(Mandatory = $true)]
		[Alias('Mailbox')]
		[String]$MailboxName,
		[Parameter(Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$Service,
		[Alias('Start')]
		[Datetime]$StartDate = (Get-Date),
		[Alias('End')]
		[Datetime]$EndDate = (Get-Date)
	)
	
	Begin {
		# initialize Variable
		[Array]$meeting = @()
	}
	Process {
		Try {
			# Define the base of the research : Calendar in $MailboxName
			$folderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $MailboxName)
			
			# Bind the base of the research and the connection of EWS
			$CalendarFolder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($Service, $folderid)
			
			# Define the view, Start-End date and the number of max item (2000)
			$cvCalendarview = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($StartDate, $EndDate, 2000)
			
			# Define the returned properties (FirstClassProperties)
			$cvCalendarview.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
			
			# Find Appointements with the filter view
			$frCalendarResult = $CalendarFolder.FindAppointments($cvCalendarview)
			
			# Foreach appointment, load the appointement and retrieve the body in text format
			Foreach ($apApointment in $frCalendarResult.Items) {
				# Define the returned properties (FirstClassProperties)
				$psPropset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
				
				# Define the RequestedBodyType to text, by default HTML
				$psPropset.RequestedBodyType = "Text"
				
				# Load the appointement
				$apApointment.load($psPropset)
				
				Write-Debug -Message "Subject: $($apApointment.Subject)"
				
				$meeting += $apApointment
			}
		} Catch [System.Management.Automation.MethodInvocationException]{
			$ErrorMessage = $_.Exception.Message
			
			$message1 = "Exception calling `"Bind`" with `"2`" argument(s): `"The SMTP address has no mailbox associated with it.`""
			$message2 = "Exception calling `"Bind`" with `"2`" argument(s): `"The specified folder could not be found in the store.`""
			
			If ($ErrorMessage -eq $message1) {
				throw [System.IO.InvalidDataException] "The SMTP address has no mailbox associated with it."
			} Elseif ($ErrorMessage -eq $message2) {
				Throw [UnauthorizedAccessException] "Permission Denied, Cannot connect to $MailboxName"
			} Else {
				Write-Error $_.Exception.Message
			}
		} Catch {
			Write-Error $_.Exception.Message
		}
	}
	End {
		Try {
			Return $meeting
		} Catch {
			Write-Error $_.Exception.Message
		}
	}
}

function New-EWSMeeting {
<#
	.SYNOPSIS
		Create a meeting in calendar through EWS
	
	.DESCRIPTION
		A detailed description of the New-MeetingEWS function.
	
	.PARAMETER Service
		Call Connect-EWS function
	
	.PARAMETER Subject
		Define the subject appointement (mandatory)
	
	.PARAMETER Body
		Define the Body
	
	.PARAMETER StartDate
		Define the start period
	
	.PARAMETER EndDate
		Define the end period
	
	.PARAMETER Location
		Add a location to the meeting
	
	.PARAMETER RequiredAttendees
		Add required attendees
	
	.PARAMETER OptionalAttendees
		Add optional attendee
	
	.PARAMETER ImpersonateMailbox
		If you need to connect another calendar (if impersonation is enable, is different of Delegation right)
	
	.PARAMETER DelegationMailbox
		If you need to connect another calendar (you must have the correct right)
	
	.PARAMETER Frequency
		Define the frequency of the reccurence. 
	
	.PARAMETER Every
		Recurence param
	
	.PARAMETER EveryWeekDay
		Recurence param (can be used only with Daily frequency)
	
	.PARAMETER Day
		Recurence param (can be used only with Weekly, Monthly & Yearly frequency)
	
	.PARAMETER Week
		Recurence param (can be used only with Monthly & Yearly frequency)
	
	.PARAMETER DayNumber
		Recurence param (can be used only with Weekly, Monthly & Yearly frequency)
	
	.PARAMETER OfMonth
		Recurence param (can be used only with Yearly frequency)
	
	.PARAMETER RecurrenceStart
		Define the start of the recurrence
	
	.PARAMETER NoEndDate
		The series has no end.
	
	.PARAMETER EndAfterOccurence
		The value of this property or element specifies the number of occurrences.
	
	.PARAMETER EndBy
		The last occurrence in the series falls on or before the date specified by this property or element.
	
	.PARAMETER AllDay
		Use this param if the appointement during all day

	.EXAMPLE
		PS C:\> New-MeetingEWS -MailboxName "myemail@mydomain.com" -Service $Service
	
	.EXAMPLE
		PS C:\> $Start = (get-Date).AddDay(1)
		PS C:\> $End =  (get-Date).AddDay(1)
		PS C:\> New-MeetingEWS -MailboxName "myemail@mydomain.com" -Service $Service -StartDate $Start -EndDate $End
		
	.VERSION
		1.0.0 - Initial version
	
		1.1.0 - Change test for yearly frequency
				Change test to define if DayNumber param is called

		2.0.0 - Add DelegationMailbox to add meeting in another mailbox
	
		2.1.0 - Correct some bug with Week param
				Remove Reccurse Parm
		
	.VALIDATION
		Exchange 2013
	
	.NOTES
		https://msdn.microsoft.com/en-us/library/office/dd633694(v=exchg.80).aspx
		https://msdn.microsoft.com/en-us/library/office/dn727655(v=exchg.150).aspx
#>
	
	[CmdletBinding(DefaultParameterSetName = 'Yearly')]
	param
	(
		[Parameter(Mandatory = $true)]
		[Microsoft.Exchange.WebServices.Data.ExchangeServiceBase]$Service,
		[Parameter(Mandatory = $true)]
		[String]$Subject,
		[String]$Body,
		[Datetime]$StartDate = (Get-Date),
		[Datetime]$EndDate = (Get-Date -Date $StartDate).AddHours(1),
		[String]$Location,
		[Array]$RequiredAttendees,
		[Array]$OptionalAttendees,
		[System.Net.Mail.MailAddress]$ImpersonateMailbox,
		[System.Net.Mail.MailAddress]$DelegationMailbox,
		[ValidateSet('Daily', 'Weekly', 'Monthly', 'Yearly')]
		[String]$Frequency,
		[Parameter(ParameterSetName = 'Daily')]
		[Parameter(ParameterSetName = 'Weekly')]
		[Parameter(ParameterSetName = 'Monthly')]
		[Parameter(ParameterSetName = 'Yearly')]
		[Int]$Every = 1,
		[Parameter(ParameterSetName = 'Daily')]
		[Switch]$EveryWeekDay,
		[Parameter(ParameterSetName = 'Weekly')]
		[Parameter(ParameterSetName = 'Monthly')]
		[Parameter(ParameterSetName = 'Yearly')]
		[Microsoft.Exchange.WebServices.Data.DayOfTheWeek[]]$Day,
		[Parameter(ParameterSetName = 'Monthly')]
		[Parameter(ParameterSetName = 'Yearly')]
		[Microsoft.Exchange.WebServices.Data.DayOfTheWeekIndex]$Week,
		[Parameter(ParameterSetName = 'Monthly')]
		[Parameter(ParameterSetName = 'Yearly')]
		[ValidateRange(1, 31)]
		[Int]$DayNumber,
		[Parameter(ParameterSetName = 'Yearly')]
		[Microsoft.Exchange.WebServices.Data.Month]$OfMonth,
		[Parameter(ParameterSetName = 'Daily')]
		[Parameter(ParameterSetName = 'Weekly')]
		[Parameter(ParameterSetName = 'Monthly')]
		[Parameter(ParameterSetName = 'Yearly')]
		[Datetime]$RecurrenceStart = (Get-Date),
		[Parameter(ParameterSetName = 'Daily')]
		[Parameter(ParameterSetName = 'Weekly')]
		[Parameter(ParameterSetName = 'Monthly')]
		[Parameter(ParameterSetName = 'Yearly')]
		[Switch]$NoEndDate,
		[Parameter(ParameterSetName = 'Daily')]
		[Parameter(ParameterSetName = 'Weekly')]
		[Parameter(ParameterSetName = 'Monthly')]
		[Parameter(ParameterSetName = 'Yearly')]
		[ValidateRange(1, 999)]
		[Int]$EndAfterOccurence,
		[Parameter(ParameterSetName = 'Daily')]
		[Parameter(ParameterSetName = 'Weekly')]
		[Parameter(ParameterSetName = 'Monthly')]
		[Parameter(ParameterSetName = 'Yearly')]
		[Datetime]$EndBy,
		[Switch]$AllDay
	)
	
	Begin {
		Try {			
			# Test part
			switch ($Frequency) {
				'Daily' {
					if ($EveryWeekDay -and $Every -gt 1) {
						Throw [system.ArgumentException] "With Daily frequency you cannot set EveryWeekDay and Every param at the same time, choose only one"
					}
					if ($Day -or $Week -or $DayNumber -or $OfMonth) {
						Throw [system.ArgumentException] "With daily frequency, you cannot set one of the following param: Day, Week, DayNumber or OfMonth"
					}
				}
				'Weekly' {
					if ($EveryWeekDay) {
						Throw [system.ArgumentException] "You cannot use EveryWeekDay with Weekly frequency, use Day param and define 'Weekday'"
					} elseif ($Week -or ($DayNumber -gt 0) -or $OfMonth) {
						Throw [system.ArgumentException] "With weekly frequency, you cannot set one of the following param: Week, DayNumber or OfMonth"
					}
				}
				'Monthly' {
					if ($EveryWeekDay) {
						Throw [system.ArgumentException] "You cannot use EveryWeekDay with Monthly frequency"
					} elseif ($OfMonth) {
						Throw [system.ArgumentException] "With Monthly frequency, you cannot set one of the following param: OfMonth"
					}
					
					if ($Week) {
						if ($DayNumber -gt 0) {
							Throw [system.ArgumentException] "With Monthly frequency you cannot set Week and DayNumber param same time, choose only one"
						} elseif (!$Day) {
							Throw [system.ArgumentException] "With Week Param, you must define Day param"
						}
					} elseif ($DayNumber -gt 0) {
						if ($Week -or $Day) {
							Throw [system.ArgumentException] "With Monthly frequency you cannot set Week/Day and DayNumber param same time, choose only one"
						}
					}
				}
				'Yearly' {
					if (!$OfMonth) {
						Throw [system.ArgumentException] "OfMonth param is mandatory"
					} Elseif ($Week -or $Day) {
						if ($DayNumber -gt 0) {
							Throw [system.ArgumentException] "With Yearly frequency you cannot set Week and DayNumber param same time, choose only one"
						} elseif ([string]::IsNullOrEmpty($Day)) {
							Throw [system.ArgumentException] "With Week Param, you must define Day param"
						} elseif ([string]::IsNullOrEmpty($Week)) {
							Throw [system.ArgumentException] "With Day Param, you must define Week param"
						}
					} elseif ($DayNumber -gt 0) {
						if ($Week -or $Day) {
							Throw [system.ArgumentException] "With Yearly frequency you cannot set Week/Day and DayNumber param same time, choose only one"
						}
					}
				}
			}
			
			
			# If $ImpersonateMailbox is true, we create the appointement to another mailbox (the account must have the write access)
			if ($ImpersonateMailbox) {
				Write-Debug "Impersonate connexion"
				# Not tested...
				$ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress), $ImpersonateMailbox.Address
				$Service.ImpersonatedUserId = $ImpersonatedUserId
				$FolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
			} elseif ($DelegationMailbox) {
				Write-Debug "Delegation connexion"
				$FolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $DelegationMailbox.Address)
			} Else {
				Write-Debug "local connexion"
				$FolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
			}
			
			Write-Debug "Bind Calendar object"
			$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $FolderId)
		} Catch [system.ArgumentException]{
			Throw [system.ArgumentException]$_.Exception.Message
			Break
		} Catch {
			Write-Error $_.Exception.Message
		}
	}
	Process {
		# Define Appointement
		Try {
			Write-Debug "Create New appointement"
			# Mandatory Argument
			$appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment($service)
			$appointment.Subject = $Subject
			
			if ($Body) {
				$appointment.Body = $Body
			}
			
			# Duration Meeting
			$appointment.Start = $StartDate
			if ($AllDay) {
				$appointment.IsAllDayEvent = $True
			} Else {
				
				$appointment.End = $EndDate
			}
			
			# Optional Argument
			if ($Location) {
				$appointment.Location = $Location
			}
			if ($RequiredAttendees) {
				$RequiredAttendees | ForEach-Object{
					[void]$appointment.RequiredAttendees.Add($_)
				}
			}
			
			if ($OptionalAttendees) {
				$OptionalAttendees | ForEach-Object{
					[void]$appointment.RequiredAttendees.Add($_)
				}
			}
			
			# If reccurence
			If ([string]::IsNullOrEmpty($Frequency)) {
				# Set the appointement as reccurent event
				$appointment.IsRecurring
				
				Write-Debug "Frequency = $Frequency"
				Switch ($Frequency) {
					'Daily' {
						If ($EveryWeekDay) {
							# https://social.technet.microsoft.com/Forums/en-US/1ba5d635-2533-4270-bc07-92d6d0e324ef/daily-and-weekly-recurrence-pattern-trouble?forum=exchangesvrdevelopment
							Write-Debug "Recurrence = Every Week Day"
							$Day = [Microsoft.Exchange.WebServices.Data.DayOfTheWeek]::Weekday
							$appointment.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+WeeklyPattern($RecurrenceStart, $every, $Day)
						} Else {
							Write-Debug "Recurrence = Every $every Day"
							$appointment.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+DailyPattern($RecurrenceStart, $every)
						}
					}
					'Weekly' {
						# By default $every = 1
						Write-Debug "Recurrence: Recur every $Every Week(s) on: $($Day -as [String])"
						$appointment.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+WeeklyPattern($RecurrenceStart, $every, $Day)
					}
					'Monthly' {
						# 2 case:
						If ($DayNumber -gt 0) {
							# Day $DayNumber of every $every
							Write-Debug "Recurrence = Day $DayNumber of every $every Month"
							$appointment.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+MonthlyPattern($RecurrenceStart, $every, $DayNumber)
						} Elseif (-not [string]::IsNullOrEmpty($Week)) {
							# The $week $day of every $every
							Write-Debug "Recurrence = The $Week $Day of every $Every Month"
							$appointment.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+RelativeMonthlyPattern($RecurrenceStart, $Every, $Day, $Week)
						}
					}
					'Yearly' {
						#2 case:
						If ($DayNumber -gt 0) {
							# On: $DayNumber $OfMonth
							Write-Debug "Recurrence = On the $DayNumber of $OfMonth"
							$appointment.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+YearlyPattern($RecurrenceStart, $OfMonth, $DayNumber)
						} Elseif (-not [string]::IsNullOrEmpty($Week)) {
							# On the $Week $day of $ofmonth
							Write-Debug "Recurrence = On the $Week $Day of $OfMonth"
							
							$appointment.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+RelativeYearlyPattern($RecurrenceStart, $OfMonth, $Day, $Week)
						}
					}
					Default {
						Write-Host -ForegroundColor Red Switch default error	
					}
				} # End Switch
				
				
				# Range of recurrence
				Write-Debug "Define Range of recurrence"
				#Define Start of ocurrence
				$appointment.Recurrence.StartDate = $RecurrenceStart
				
				#End
				# 3 cases:
				If ($NoEndDate) {
					# No end date
					$appointment.Recurrence.NeverEnds()
				} Elseif ($EndAfterOccurence -ne 0) {
					# end after xx occurences
					$appointment.Recurrence.NumberOfOccurrences = $EndAfterOccurence
					$appointment.Recurrence.NumberOfOccurrences | Out-Host
				} Elseif ($EndBy) {
					# end by [datetime]
					$appointment.Recurrence.EndDate = $EndBy
				} Else {
					Write-Debug "End Recurrence isn't been defined. we used default value : NoEndDate"
					$appointment.Recurrence.NeverEnds()
				}
			}
		} Catch {
			Write-Error $_.Exception.Message
		}
	}
	End {
		Try {
			Write-Debug "Save appointement"
			if ($RequiredAttendees -or $OptionalAttendees) {
				$appointment.Save($Calendar.Id, [Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToAllAndSaveCopy)
			} Else {
				$appointment.Save($Calendar.Id, [Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
			}
		} Catch {
			Write-Error $_.Exception.Message
		}
	}
}


Export-ModuleMember -Function Connect-EWS,
					Get-EWSFolder,
					Get-EWSMail,
					Move-EWSMail,
					Remove-EWSMail,
					Get-EWSCalendar,
					Get-EWSCalendarPermission,
					Set-EWSCalendarPermission,
					Remove-EWSCalendarPermission,
					Get-EWSMeeting,
					New-EWSMeeting