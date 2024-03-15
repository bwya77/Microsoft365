<#
.SYNOPSIS
    Retrieves Intune device configuration scripts from Microsoft Graph.

.DESCRIPTION
    The Get-IntuneDeviceConfigurationScripts function connects to Microsoft Graph and retrieves Intune device configuration scripts.
    It can optionally export the scripts to a specified directory or return the results as an array of objects.

.PARAMETER URI
    The URI of the Microsoft Graph endpoint to retrieve the device management scripts.
    Default value is "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts".

.PARAMETER Scope
    The scope required to access the device management scripts.
    Default value is "DeviceManagementConfiguration.Read.All".

.PARAMETER exportDirectory
    The directory path where the scripts should be exported.
    If specified, the function will create the directory if it does not exist and export the scripts to that directory.
    If not specified, the function will return the scripts as an array of objects.

.EXAMPLE
    Get-IntuneDeviceConfigurationScripts -exportDirectory "C:\Scripts"

    This example retrieves the Intune device configuration scripts and exports them to the "C:\Scripts" directory.

.EXAMPLE
    Get-IntuneDeviceConfigurationScripts

    This example retrieves the Intune device configuration scripts using the default URI and scope and returns them as an array of objects.

#>
function Get-IntuneDeviceConfigurationScripts {
    Param(
        [Parameter()]
        [System.String]$Uri = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts",
        [Parameter()]
        [System.String]$Scope = "DeviceManagementConfiguration.Read.All",
        [Parameter()]
        [System.String]$exportDirectory
    )
    begin {
        Write-Verbose "Checking for Microsoft.Graph.Authentication module"
        if (-not (Get-Module Microsoft.Graph.Authentication -ListAvailable)) {
            Write-Verbose "Installing Microsoft.Graph.Authentication module"
            Try {
                $module = @{
                    Name        = 'Microsoft.Graph.Authentication'
                    Scope       = 'CurrentUser'
                    Force       = $true
                    ErrorAction = "Stop"
                }
                #Install the Microsoft.Graph.Authentication module
                Install-Module @module
                Write-Verbose "Importing Microsoft.Graph.Authentication module"
                #Import the Microsoft.Graph.Authentication module into memory
                Import-Module Microsoft.Graph.Authentication
            } 
            Catch {
                Write-Warning "Failed to install Microsoft.Graph.Authentication module. Please install the module manually and try again."
                break
            }
        }
        #Test the export folder if the variable is not null and create it if it does not exist
        if ($exportDirectory -and !(Test-Path -Path $exportDirectory)) {
            Write-Verbose "Creating Export Folder"
            $newDirectory = @{
                Path        = $exportDirectory
                ItemType    = "Directory"
                Force       = $true
                ErrorAction = "Stop"
            } 
            # Create the root export directory
            New-Item @newDirectory | Out-Null
        }
        [System.Array]$dataArray = @()
        [System.Array]$resultsArray = @()
    }
    process {
        Write-Verbose "Connecting to Microsoft Graph"
        try {
            $Request = @{
                Scope        = $Scope
                ContextScope = 'Process'
                NoWelcome    = $true
                ErrorAction  = 'Stop'
            }
            #Connect to Microsoft Graph REST API
            Connect-MgGraph @Request
        }
        catch {
            Write-Warning "Failed to connect to Microsoft Graph. Please check your credentials and try again."
            break
        }
        Write-Verbose "Checking for proper scope"
        #Get the scopes that we are connected to
        $AuthdScopes = (Get-MgContext | Select-Object -ExpandProperty Scopes)
        #If we are not connected to the correct scope, we need to throw an error and exit
        if ($Scope -notin $AuthdScopes) {
            Write-Warning "You are not connected with the correct scope. Please connect to the correct scope and try again."
            break
        }
        $Request = @{
            Uri    = $Uri
            Method = "GET"
        }
        #Invoke the Microsoft Graph request
        $data = Invoke-MGGraphRequest @Request
        while ($data.'@odata.nextLink') {
            $resultsArray += $data.value
            $Request.Uri = $data.'@odata.nextLink'
            $data = Invoke-MGGraphRequest @Request
            # to avoid throttling, the loop will sleep for 3 seconds
            Start-Sleep -Seconds 3
        }
        $resultsArray += $data.value
        #iterate through the results and return the script content
        foreach ($IntuneScript in $resultsArray) {
            Write-Verbose "Processing Script: $($IntuneScript.displayName)"
            $Request.Uri = "https://graph.microsoft.com/Beta/deviceManagement/deviceManagementScripts/$($IntuneScript.id)"
            # Get the script content which is saved in base64
            $IntuneScriptContent = (Invoke-MGGraphRequest @Request).scriptContent
            #Decode the base64 encoded script content
            $IntuneScriptContent = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($IntuneScriptContent))
            $resultsTable = [PSCustomObject]@{
                Name     = $IntuneScript.displayName
                FileName = $IntuneScript.fileName
                Script   = $IntuneScriptContent
            }
            $dataArray += $resultsTable
        }
    }
    end {
        if ($exportDirectory) {
            Write-Verbose "Exporting scripts to $exportDirectory"
            foreach ($i in $dataArray) {
                #Create a new folder for each policy
                [System.String]$policyFolder = "$exportDirectory/$($i.Name)"
                #Create a new folder that matches the policy name
                New-Item -ItemType Directory -Path $policyFolder -Force | Out-Null
                #Create the file path
                $FilePath = Join-Path -Path $policyFolder -ChildPath $i.fileName
                #Export the script to the policy folder
                $i.Script | Out-File -FilePath $FilePath -Force
                #Check that the file was created
                if ((Test-Path -Path $FilePath) -eq $false) {
                    Write-Warning "Failed to export script to $FilePath"
                }
            }
        }
        else {
            #Display the results
            $dataArray
        }
        Write-Verbose "Disconnecting from Microsoft Graph"
        #Disconnect from Microsoft Graph
        Disconnect-MgGraph | Out-Null
    }
}