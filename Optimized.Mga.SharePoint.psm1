#region functions
function Get-MgaSharePointFile {
    <#
    .SYNOPSIS
    With Get-MgaSharePointFile you can get a list of sharepoint files in a specific site. 

    .DESCRIPTION
    With Get-MgaSharePointFile you can get a list of sharepoint files in a specific site. 
    Keep it mind that it will return all files in the site, use Where-Object to filter results.

    .PARAMETER TenantName
    This is the URL to sharepoint, but the tenantname (before .onmicrosoft.com) is sufficient

    .PARAMETER Site
    This is the sitename where the files are stored in.

    .PARAMETER ChildFolders
    Add childfolders as an array firstfolder,subfolder,subsubfolder

    .EXAMPLE
    $SPItems = Get-MgaSharePointFile -TenantName 'BWIT.onmicrosoft.com' -Site 'Team_' -ChildFolders 'O365Reports'

    .NOTES
    Author: Bas Wijdenes

    .LINK
    https://baswijdenes.com/how-to-download-files-from-sharepoint-with-ms-graph-api-and-powershell/
    https://baswijdenes.com/how-to-upload-files-to-sharepoint-with-ms-graph-api-and-powershell/
    #>
    [CmdletBinding()]
    param (
        [Parameter(mandatory , HelpMessage = 'This is the URL to sharepoint, but the tenantname (before .onmicrosoft.com) is sufficient')]
        [string]
        $TenantName,
        [Parameter(mandatory, HelpMessage = 'Add the sitename')]
        [string]
        $Site,
        [Parameter(mandatory = $false, HelpMessage = 'Add childfolders as an array firstfolder,subfolder,subsubfolder')]
        [string[]]
        $ChildFolders
    )
    begin {
        if ($TenantName -like '*.*') {
            $TenantName = $TenantName.split('.')[0]
            Write-Verbose "Get-MgaSharePointFile: begin: Converted TenantName to $TenantName"
        }
        else {
            Write-Verbose "Get-MgaSharePointFile: begin: TenantName is $TenantName"
        }
        Write-Verbose "Get-MgaSharePointFile: begin: Site is $Sitename" 
        $SPURL = 'https://graph.microsoft.com/v1.0/sites/{0}.sharepoint.com:/sites/{1}/' -f $TenantName, $Site
        Write-Verbose "Get-MgaSharePointFile: begin: SPURL is $SPURL" 
        $SPChildrenURL = 'https://graph.microsoft.com/v1.0/sites/{0}/drive/items/root'
        $i = 1
        if ($ChildFolders) {
            $SPChildrenURL = 'https://graph.microsoft.com/v1.0/sites/{0}/drive/items/root:'
            foreach ($ChildFolder in $ChildFolders) {
                if ($i -eq $($ChildFolders).count) {
                    $SPChildrenURL = "$($SPChildrenURL)/$($ChildFolder):/children"
                }
                else {
                    $SPChildrenURL = "$($SPChildrenURL)/$($ChildFolder)"
                }
                $i++
            }
        } 
        else {
            $SPChildrenURL = "$($SPChildrenURL)/children"
        }
    }
    process {
        $SPsite = Get-Mga -URL $SPURL
        $SPItemsURL = $($SPChildrenURL) -f $SPSite.id
        Write-Verbose "Get-MgaSharePointFile: begin: SPItemsURL is $SPItemsURL" 
        $SPItems = Get-Mga -URL $SPItemsURL
    }
    end {
        return $SPItems
    }
}

function Download-MgaSharePointFile {
    <#
    .SYNOPSIS
    With Download-MgaSharePointFile you can download files from a SP Site.
    
    .DESCRIPTION
    Download-MgaSharePointFile will only work with Get-MgaSharePointFile return.
    
    .PARAMETER SPItem
    This Parameter needs the return from Get-MgaSharePointFile.
    
    .PARAMETER OutputFolder
    This is a FolderPath to where the files need to be exported
    
    .EXAMPLE
    foreach ($Item in $SPItems) {
        Download-MgaSharePointFile  -SPItem $Item -OutputFolder 'C:\temp\'
    }

    .NOTES
    Author: Bas Wijdenes

    .LINK
    https://baswijdenes.com/how-to-download-files-from-sharepoint-with-ms-graph-api-and-powershell/
    #>
    [CmdletBinding()]
    param (
        [parameter(mandatory)]
        $SPItem,
        [parameter(mandatory = $true, ParameterSetName = 'OutputFolder')]
        [string]
        $OutputFolder
    ) 
    begin {
        if (($OutputFolder) -and ($OutputFolder.Substring($OutputFolder.Length - 1, 1) -eq '\')) {
            Write-Verbose "Download-MgaSharePointFile: begin: $OutputFolder ends with a '\' script will trim the end"
            $OutputFolder = $OutputFolder.TrimEnd('\')
        }
    }   
    process {
        try {
            try {
                $ContentInBytes = Invoke-WebRequest -Uri $spitem.'@microsoft.graph.downloadUrl'
            }
            catch {
                try {
                    if ($_.Exception.Message -like '*Internet Explorer engine*') {
                        $ContentInBytes = Invoke-WebRequest -Uri $spitem.'@microsoft.graph.downloadUrl' -UseBasicParsing
                    }
                    else {
                        throw $_.Exception.Message
                    }
                }
                catch {
                    throw $_.Exception.Message
                }
            }
            Write-Verbose "Download-MgaSharePointFile: process: retrieved $($SPItem.Name) content"
            if ($OutputFolder) {
                Write-Verbose "Download-MgaSharePointFile: process: Exporting $($SPItem.Name) content"
                [System.IO.file]::WriteAllBytes("$OutputFolder\$($SPItem.Name)", $ContentInBytes.content)
                $Return = "Exported $($SPItem.Name) in $OutputFolder"
            } 
            else {
                Write-Verbose "Download-MgaSharePointFile: process: Converting $($SPItem.Name) content to UTF8 to return"
                $Return = ([System.Text.Encoding]::UTF8.GetString($ContentInBytes.content)).substring(2)
            }
        }
        catch {
            throw $_.Exception.Message
        }
    }    
    end {
        return $Return
    }
}

function Upload-MgaSharePointFile {
    <#
    .SYNOPSIS
    With Upload-MgaSharePointFiles you can upload files to SharePoint sites and OneDrive.

    .DESCRIPTION
    With Upload-MgaSharePointFiles you can upload files to SharePoint sites and OneDrive.

    .PARAMETER ItemPath
    ItemPath will accept a path string to the item you want to upload to SharePoint.

    .PARAMETER Item
    The parameter Item will accept an item of object [System.IO.FileSystemInfo]. 
    Which means that you first use Get-Item -Path C:\temp\blatemp.txt before you start Upload-MgaSharePointFile.

    .PARAMETER Type
    Type is SharePoint or OneDrive. 
    The default is SharePoint.

    .PARAMETER TenantName
    This is the URL to sharepoint, but the tenantname (before .onmicrosoft.com) is sufficient'

    .PARAMETER Site
    This is the sitename where the files are stored in.

    .PARAMETER ChildFolders
    Add childfolders as an array firstfolder,subfolder,subsubfolder

    .EXAMPLE
    Upload-MgaSharePointFiles -ItemPath "C:\temp\blatemp.txt" -TenantName 'm365x794103.onmicrosoft.com' -Site 'XXXX'

    .NOTES
    Author: Bas Wijdenes

    .LINK
    https://baswijdenes.com/how-to-upload-files-to-sharepoint-with-ms-graph-api-and-powershell/
#>
    [CmdletBinding()]
    param (
        [Parameter(mandatory = $true, ParameterSetName = 'ItemPath')]
        [string]
        $ItemPath,
        [Parameter(mandatory = $true, ParameterSetName = 'Item')]
        [System.IO.FileSystemInfo]
        $Item,
        [Parameter(mandatory = $false)]
        [ValidateSet('SharePoint', 'OneDrive')]
        $Type = 'SharePoint', 
        [Parameter(mandatory , HelpMessage = 'This is the URL to sharepoint, but the tenantname (before .onmicrosoft.com) is sufficient')]
        [string]
        $TenantName,
        [Parameter(mandatory, HelpMessage = 'Add the sitename')]
        [string]
        $Site,
        [Parameter(mandatory = $false, HelpMessage = 'Add childfolders as an array firstfolder,subfolder,subsubfolder')]
        [string[]]
        $ChildFolders
    )
    begin {
        if ($PSCmdlet.ParameterSetName -eq 'ItemPath') {
            if ((Test-Path $ItemPath) -eq $false) {
                throw "File $ItemPath cannot be found"
            } 
            else {
                $File = Get-Item $ItemPath
                $LocalFileBytes = [System.IO.File]::ReadAllBytes($File)
            }
        }
        else {
            $File = $Item
            $LocalFileBytes = [System.IO.File]::ReadAllBytes($File.FullName)
        }
        if ($TenantName -like '*.*') {
            $TenantName = $TenantName.split('.')[0]
            Write-Verbose "Upload-MgaSharePointFile: begin: Converted TenantName to $TenantName"
        }
        else {
            Write-Verbose "Upload-MgaSharePointFile: begin: TenantName is $TenantName"
        }
        Write-Verbose "Upload-MgaSharePointFile: begin: Site is $Site" 
        $SPURL = 'https://graph.microsoft.com/v1.0/sites/{0}.sharepoint.com:/sites/{1}/' -f $TenantName, $Site
        Write-Verbose "Upload-MgaSharePointFile: begin: SPURL is $SPURL" 
        $SPChildrenURL = 'https://graph.microsoft.com/v1.0/sites/{0}/drive/items/root:'
        $i = 1
        if ($ChildFolders) {
            foreach ($ChildFolder in $ChildFolders) {
                if ($i -eq $($ChildFolders).count) {
                    $SPChildrenURL = "$($SPChildrenURL)/$($ChildFolder)/{1}:/createUploadSession"
                    Write-Verbose "Upload-MgaSharePointFile: begin: ChildFolder URL is $SPChildrenURL"
                }
                else {
                    $SPChildrenURL = "$($SPChildrenURL)/$($ChildFolder)"
                }
                $i++
            }
        } 
        else {
            $SPChildrenURL = "$($SPChildrenURL)/{1}:/createUploadSession"

        }
        if ($Type -eq 'OneDrive') {
            $global:SPURL = $SPURL.Replace('/sites/', '/drives/')
            $global:SPChildrenURL = $SPChildrenURL.Replace('/sites/', '/drives/')
            $global:SPURL = $SPURL.Replace('/drive/', '')
            $global:SPChildrenURL = $SPChildrenURL.Replace('/drive/', '')
        } 
    }
    process {
        $SPsite = Get-Mga -URL $SPURL
        $SPItemsURL = $($SPChildrenURL) -f $SPSite.id, $File.Name
        Write-Verbose "Upload-MgaSharePointFile: begin: Upload URL is $SPItemsURL"
        $uploadUrlResponse = Post-Mga -URL $SPItemsURL
        $contentRange = [string]::Format('bytes 0-{0}/{1}', $($LocalFileBytes.Length - 1), $LocalFileBytes.Length)
        $Header = @{}
        $Header.Add('Content-Length', $LocalFileBytes.Length)
        $Header.Add('Content-Range', $contentRange)
        $Header.Add('Content-Type', 'octet/stream')
        Write-Verbose $contentRange
        $UploadResult = Put-Mga -URL $uploadUrlResponse.uploadUrl -InputObject $LocalFileBytes -CustomHeader $Header -Verbose
    }
    end {
        return $UploadResult
    }
}

function Get-MgaSharePointList {
    [CmdletBinding()]
    param (
        [Parameter(mandatory, HelpMessage = 'Add the sitename')]
        [string]
        $Site,
        [Parameter(mandatory, HelpMessage = 'Add the Listname')]
        [string]
        $List,
        [Parameter(mandatory , HelpMessage = 'This is the URL to sharepoint, but the tenantname (before .onmicrosoft.com) is sufficient')]
        [string]
        $TenantName
    )  
    begin {
        Write-Verbose "Get-MgaSharePointList: begin: site: $Site"
        Write-Verbose "Get-MgaSharePointList: begin: list: $List"
        if ($TenantName -like '*.*') {
            $TenantName = $TenantName.split('.')[0]
            Write-Verbose "Get-MgaSharePointList: begin: Converted TenantName to $TenantName"
        }
        else {
            Write-Verbose "Get-MgaSharePointList: begin: TenantName is $TenantName"
        }
    }
    process {        
        $SPSiteURL = 'https://graph.microsoft.com/v1.0/sites/{0}.sharepoint.com:/sites/{1}/' -f $TenantName, $Site
        $SPSite = Get-Mga -URL $SPSiteURL
        $SPListURL = 'https://graph.microsoft.com/v1.0/sites/{0}/lists/{1}/items?expand=fields' -f $SPSite.id, $List
        $Response = Get-Mga -URL $SPListURL 
    }    
    end {
        return $Response
    }
}
#endregion functions
#region aliases
New-Alias -Name 'Get-MgaSharePointFiles' -Value 'Get-MgaSharePointFile'
New-Alias -Name 'Download-MgaSharePointFiles' -Value 'Download-MgaSharePointFile'
New-Alias -Name 'Upload-MgaSharePointFiles' -Value 'Upload-MgaSharePointFile'
#endregion aliases