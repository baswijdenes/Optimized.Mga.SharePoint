#region functions
#region Azure AD activity reports
#region Directory audit
function Get-MgaReportdirectoryAudits {
    [CmdletBinding()]
    param (
    )
       begin {
             $null = Test-MgaRole -RoleType 'AzureAD'
        $URL = 'https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?$top=999'
        Write-Verbose "Get-MgaReportdirectoryAudits: begin: URL: $URL"
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportdirectoryAudits: Getting results | Result count: $($GetMga.Count)"
        $ReportsList = $GetMga
    }  
    end {
        return $ReportsList
    }
}

function Get-MgaReportdirectoryAudit {
    [CmdletBinding()]
    param (
        [parameter(mandatory = $true)]
        $AuditId
    )
       begin {
             $null = Test-MgaRole -RoleType 'AzureAD'
        $URL = 'https://graph.microsoft.com/v1.0/auditLogs/directoryAudits/{0}' -f $AuditId
        Write-Verbose "Get-MgaReportdirectoryAudit: begin: URL: $URL"
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportdirectoryAudit: Getting results | Result count: $($GetMga.Count)"
        $ReportsList = $GetMga
    }  
    end {
        return $ReportsList
    }
}
#endregion Directory audit
#region Sign-in
function Get-MgaReportsignIns {
    [CmdletBinding()]
    param (
    )
       begin {
             $null = Test-MgaRole -RoleType 'AzureAD'
        $URL = 'https://graph.microsoft.com/v1.0/auditLogs/signIns?$top=999'
        Write-Verbose "Get-MgaReportsignIns: begin: URL: $URL"
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportsignIns: Getting results | Result count: $($GetMga.Count)"
        $ReportsList = $GetMga
    }  
    end {
        return $ReportsList
    }
}

function Get-MgaReportsignIn {
    [CmdletBinding()]
    param (
        [parameter(mandatory = $true)]
        $SignInId
    )
       begin {
             $null = Test-MgaRole -RoleType 'AzureAD'
        $URL = 'https://graph.microsoft.com/v1.0/auditLogs/signIns/{0}' -f $SignInId
        Write-Verbose "Get-MgaReportsignIn: begin: URL: $URL"
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportsignIn: Getting results | Result count: $($GetMga.Count)"
        $ReportsList = $GetMga
    }  
    end {
        return $ReportsList
    }
}
#endregion
#region Provisioning
function Get-MgaReportprovisioningObjectSummary {
    [CmdletBinding()]
    param (
    )
       begin {
             $null = Test-MgaRole -RoleType 'AzureAD'
        $URL = 'https://graph.microsoft.com/v1.0/auditLogs/provisioning?$top=999'
        Write-Verbose "Get-MgaReportprovisioningObjectSummary: begin: URL: $URL"
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportprovisioningObjectSummary: Getting results | Result count: $($GetMga.Count)"
        $ReportsList = $GetMga
    }  
    end {
        return $ReportsList
    }
}
#endregion
#endregion AzureAD activity reports
#region  Microsoft 365 usage reports
#region Microsoft Teams device usage
function Get-MgaReportTeamsDeviceUsageUserDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportTeamsDeviceUsageUserDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getTeamsDeviceUsageUserDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportTeamsDeviceUsageUserDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportTeamsDeviceUsageUserDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportTeamsDeviceUsageUserDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                UserPrincipalName = $ReportItem.'User Principal Name'
                IsDeleted         = $ReportItem.'Is Deleted'
                LastActivityDate  = $ReportItem.'Last Activity Date'
                UsedWeb           = $ReportItem.'Used Web'
                UsedWindowsPhone  = $ReportItem.'Used Windows Phone'
                UsediOS           = $ReportItem.'Used iOS'
                UsedMac           = $ReportItem.'Used Mac'
                UsedAndroidPhone  = $ReportItem.'Used Android Phone'
                UsedWindows       = $ReportItem.'Used Windows'
            }
            if ($ReportItem.'Is Deleted' -eq $true) {
                $Object.DeletedDate = $ReportItem.'Deleted Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }  
    end {
        return $ReportsList
    }
}

function Get-MgaReportTeamsDeviceUsageUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportTeamsDeviceUsageUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getTeamsDeviceUsageUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportTeamsDeviceUsageUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportTeamsDeviceUsageUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportTeamsDeviceUsageUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                Web          = $ReportItem.Web
                WindowsPhone = $ReportItem.'Windows Phone'
                AndroidPhone = $ReportItem.'Android Phone'
                iOS          = $ReportItem.'iOS'
                Mac          = $ReportItem.'Mac'
                Windows      = $ReportItem.'Windows'
                ReportDate   = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }  
    end {
        return $ReportsList
    }
}

function Get-MgaReportTeamsDeviceUsageDistributionUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportTeamsDeviceUsageDistributionUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getTeamsDeviceUsageDistributionUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportTeamsDeviceUsageDistributionUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportTeamsDeviceUsageDistributionUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportTeamsDeviceUsageDistributionUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                Web          = $ReportItem.Web
                WindowsPhone = $ReportItem.'Windows Phone'
                AndroidPhone = $ReportItem.'Android Phone'
                iOS          = $ReportItem.'iOS'
                Mac          = $ReportItem.'Mac'
                Windows      = $ReportItem.'Windows'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }  
    end {
        return $ReportsList
    }
}
#endregion Microsoft Teams device usage
#region Microsoft Teams user activity
function Get-MgaReportTeamsUserActivityUserDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportTeamsUserActivityUserDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportTeamsUserActivityUserDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportTeamsUserActivityUserDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportTeamsUserActivityUserDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $AssignedProducts = $null
            $AssignedProducts = $ReportItem.'Assigned Products'.Split('+').Trim()
            $Object = @{
                UserPrincipalName       = $ReportItem.'User Principal Name'
                LastActivityDate        = $ReportItem.'Last Activity Date'
                IsDeleted               = $ReportItem.'Is Deleted'
                AssignedProducts        = $AssignedProducts
                TeamChatMessageCount    = $ReportItem.'Team Chat Message Count'
                PrivateChatMessageCount = $ReportItem.'Private Chat Message Count'
                CallCount               = $ReportItem.'Call Count'
                MeetingCount            = $ReportItem.'Meeting Count'           
                HasOtherAction          = $ReportItem.'Has Other Action'        
            }
            if ($ReportItem.'Is Deleted' -eq $true) {
                $Object.DeletedDate = $ReportItem.'Deleted Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }  
    end {
        return $ReportsList
    }
}

function Get-MgaReportTeamsUserActivityCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportTeamsUserActivityCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportTeamsUserActivityCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportTeamsUserActivityCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportTeamsUserActivityCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                ReportDate          = $ReportItem.'Report Date'
                TeamChatMessages    = $ReportItem.'Team Chat Messages'
                PrivateChatMessages = $ReportItem.'Private Chat Messages'
                Calls               = $ReportItem.Calls
                Meetings            = $ReportItem.Meetings
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }  
    end {
        return $ReportsList
    }
}

function Get-MgaReportTeamsUserActivityUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportTeamsUserActivityUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportTeamsUserActivityUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportTeamsUserActivityUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportTeamsUserActivityUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                ReportDate          = $ReportItem.'Report Date'
                TeamChatMessages    = $ReportItem.'Team Chat Messages'
                PrivateChatMessages = $ReportItem.'Private Chat Messages'
                Calls               = $ReportItem.Calls
                Meetings            = $ReportItem.Meetings
                OtherActions        = $ReportItem.'Other Actions'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }  
    end {
        return $ReportsList
    }
}
#endregion Microsoft Teams user activity
#region Outlook Activity
function Get-MgaReportEmailActivityUserDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportEmailActivityUserDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportEmailActivityUserDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportEmailActivityUserDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportEmailActivityUserDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $AssignedProducts = $null
            $AssignedProducts = $ReportItem.'Assigned Products'.Split('+').Trim()
            $Object = @{
                UserPrincipalName = $ReportItem.'User Principal Name'
                DisplayName       = $ReportItem.'Display Name'
                IsDeleted         = $ReportItem.'Is Deleted'    
                LastActivityDate  = $ReportItem.'Last Activity Date'
                SendCount         = $ReportItem.'Send Count'
                ReceiveCount      = $ReportItem.'Receive Count' 
                ReadCount         = $ReportItem.'Read Count'  
                AssignedProducts  = $AssignedProducts
            }           
            if ($ReportItem.'Is Deleted' -eq $true) {
                $Object.DeletedDate = $ReportItem.'Deleted Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }  
    end {
        return $ReportsList
    }
}

function Get-MgaReportEmailActivityCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportEmailActivityCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getEmailActivityCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportEmailActivityCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportEmailActivityCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportEmailActivityCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                Send       = $ReportItem.Send
                Receive    = $ReportItem.Receive
                Read       = $ReportItem.Read    
                ReportDate = $ReportItem.'Report Date'
            }           
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }  
    end {
        return $ReportsList
    }
}

function Get-MgaReportEmailActivityUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportEmailActivityUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportEmailActivityUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportEmailActivityUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportEmailActivityUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                Send       = $ReportItem.Send
                Receive    = $ReportItem.Receive
                Read       = $ReportItem.Read    
                ReportDate = $ReportItem.'Report Date'
            }           
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }  
    end {
        return $ReportsList
    }
}
#endregion Outlook Activity
#region Outlook app usage
function Get-MgaReportEmailAppUsageUserDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportEmailAppUsageUserDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getEmailAppUsageUserDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportEmailAppUsageUserDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportEmailAppUsageUserDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportEmailAppUsageUserDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                UserPrincipalName = $ReportItem.'User Principal Name'
                DisplayName       = $ReportItem.'Display Name'
                IsDeleted         = $ReportItem.'Is Deleted'
                LastActivityDate  = $ReportItem.'Last Activity Date'
                MailForMac        = $ReportItem.'Mail For Mac'
                OutlookForMac     = $ReportItem.'Outlook For Mac'
                OutlookForWindows = $ReportItem.'Outlook For Windows'
                OutlookForMobile  = $ReportItem.'Outlook For Mobile'
                OtherForMobile    = $ReportItem.'Other For Mobile'
                OutlookForWeb     = $ReportItem.'Outlook For Web'
                POP3App           = $ReportItem.'POP3 App'
                IMAP4App          = $ReportItem.'IMAP4 App'
                SMTPApp           = $ReportItem.'SMTP App'
            }
            if ($ReportItem.'Is Deleted' -eq $true) {
                $Object.DeletedDate = $ReportItem.'Deleted Date'
            }           
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }  
    end {
        return $ReportsList
    }
}

function Get-MgaReportEmailAppUsageAppsUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportEmailAppUsageAppsUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getEmailAppUsageAppsUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportEmailAppUsageAppsUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportEmailAppUsageAppsUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportEmailAppUsageAppsUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                MailForMac        = $ReportItem.'Mail For Mac'
                OutlookForMac     = $ReportItem.'Outlook For Mac'
                OutlookForWindows = $ReportItem.'Outlook For Windows'
                OutlookForMobile  = $ReportItem.'Outlook For Mobile'
                OtherForMobile    = $ReportItem.'Other For Mobile'
                OutlookForWeb     = $ReportItem.'Outlook For Web'
                POP3App           = $ReportItem.'POP3 App'
                IMAP4App          = $ReportItem.'IMAP4 App'
                SMTPApp           = $ReportItem.'SMTP App'
            }        
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }  
    end {
        return $ReportsList
    }
}

function Get-MgaReportEmailAppUsageUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportEmailAppUsageUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getEmailAppUsageUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportEmailAppUsageUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportEmailAppUsageUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportEmailAppUsageUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                MailForMac        = $ReportItem.'Mail For Mac'
                OutlookForMac     = $ReportItem.'Outlook For Mac'
                OutlookForWindows = $ReportItem.'Outlook For Windows'
                OutlookForMobile  = $ReportItem.'Outlook For Mobile'
                OtherForMobile    = $ReportItem.'Other For Mobile'
                OutlookForWeb     = $ReportItem.'Outlook For Web'
                POP3App           = $ReportItem.'POP3 App'
                IMAP4App          = $ReportItem.'IMAP4 App'
                SMTPApp           = $ReportItem.'SMTP App'
                ReportDate        = $ReportItem.'Report Date'
            }        
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }  
    end {
        return $ReportsList
    }
}

function Get-MgaReportEmailAppUsageVersionsUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportEmailAppUsageVersionsUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getEmailAppUsageVersionsUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportEmailAppUsageVersionsUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportEmailAppUsageVersionsUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportEmailAppUsageVersionsUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                Outlook2016  = $ReportItem.'Outlook 2016'
                Outlook2013  = $ReportItem.'Outlook 2013'
                Outlook2010  = $ReportItem.'Outlook 2010'
                Outlook2007  = $ReportItem.'Outlook 2007'
                Undetermined = $ReportItem.Undetermined
                OutlookM365  = $ReportItem.'Outlook M365'
                Outlook2019  = $ReportItem.'Outlook 2019'
            }        
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }  
    end {
        return $ReportsList
    }
}
#endregion Outlook app usage
#region Outlook mailbox usage
function Get-MgaReportMailboxUsageDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportMailboxUsageDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportMailboxUsageDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportMailboxUsageDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportMailboxUsageDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                DisplayName                  = $ReportItem.'Display Name'
                UserPrincipalName            = $ReportItem.'User Principal Name'
                SizeInGb                     = [math]::Round(((($ReportItem.'Storage Used (Byte)') / 1024 / 1024 / 1024)), 2)
                CreatedDate                  = $ReportItem.'Created Date'
                IsDeleted                    = $ReportItem.'Is Deleted'
                LastActivityDate             = $ReportItem.'Last Activity Date'
                ItemCount                    = $ReportItem.'Item Count'
                IssueWarningQuotaInGB        = [math]::Round(((($ReportItem.'Issue Warning Quota (Byte)') / 1024 / 1024 / 1024)), 2)
                ProhibitSendQuotaInGB        = [math]::Round(((($ReportItem.'Prohibit Send Quota (Byte)') / 1024 / 1024 / 1024)), 2)
                ProhibitSendReceiveQuotaInGb = [math]::Round(((($ReportItem.'Prohibit Send/Receive Quota (Byte)') / 1024 / 1024 / 1024)), 2)
                DeletedItemCount             = $ReportItem.'Deleted Item Count'
                DeletedItemSizeInGb          = [math]::Round(((($ReportItem.'Deleted Item Size (Byte)') / 1024 / 1024 / 1024)), 2)
                DeletedItemQuotaInGb         = [math]::Round(((($ReportItem.'Deleted Item Quota (Byte)') / 1024 / 1024 / 1024)), 2)
                HasArchive                   = $ReportItem.'Has Archive'
            }
            if ($ReportItem.'Is Deleted' -eq $true) {
                $Object.DeletedDate = $ReportItem.'Deleted Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}

function Get-MgaReportMailboxUsageMailboxCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportMailboxUsageMailboxCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageMailboxCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportMailboxUsageMailboxCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportMailboxUsageMailboxCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportMailboxUsageMailboxCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                Total      = $ReportItem.Total
                Active     = $ReportItem.Active
                ReportDate = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}

function Get-MgaReportMailboxUsageQuotaStatusMailboxCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportMailboxUsageQuotaStatusMailboxCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageQuotaStatusMailboxCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportMailboxUsageQuotaStatusMailboxCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportMailboxUsageQuotaStatusMailboxCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportMailboxUsageQuotaStatusMailboxCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{        
                UnderLimit            = $ReportItem.'Under Limit'
                WarningIssued         = $ReportItem.'Warning Issued'
                SendProhibited        = $ReportItem.'Send Prohibited'
                SendReceiveProhibited = $ReportItem.'Send/Receive Prohibited'
                Indeterminate         = $ReportItem.Indeterminate
                ReportDate            = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}

function Get-MgaReportMailboxUsageStorage {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportMailboxUsageStorage: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageStorage(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportMailboxUsageStorage: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportMailboxUsageStorage: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportMailboxUsageStorage: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{        
                StorageUsedInGb = [math]::Round(((($ReportItem.'Storage Used (Byte)') / 1024 / 1024 / 1024)), 2)
                ReportDate      = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}
#endregion Outlook mailbox usage
#region Microsoft 365 activations
function Get-MgaReportOffice365ActivationsUserDetail {
    [CmdletBinding()]
    param (
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        $URL = "https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail" 
        Write-Verbose "Get-MgaReportOffice365ActivationsUserDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOffice365ActivationsUserDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOffice365ActivationsUserDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{        
                UserPrincipalName         = $ReportItem.'User Principal Name'
                DisplayName               = $ReportItem.'Display Name'
                ProductType               = $ReportItem.'Product Type'
                LastActivatedDate         = $ReportItem.'Last Activated Date'
                Windows                   = $ReportItem.Windows
                Mac                       = $ReportItem.Mac
                Windows10Mobile           = $ReportItem.'Windows 10 Mobile'
                iOS                       = $ReportItem.iOS
                Android                   = $ReportItem.Android
                ActivatedOnSharedComputer = $ReportItem.'Activated On Shared Computer'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}

function Get-MgaReportOffice365ActivationCounts {
    [CmdletBinding()]
    param (
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        $URL = "https://graph.microsoft.com/v1.0/reports/getOffice365ActivationCounts" 
        Write-Verbose "Get-MgaReportOffice365ActivationCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOffice365ActivationCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOffice365ActivationCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{        
                ProductType     = $ReportItem.'Product Type'
                Windows         = $ReportItem.Windows
                Mac             = $ReportItem.Mac
                Windows10Mobile = $ReportItem.'Windows 10 Mobile'
                iOS             = $ReportItem.iOS
                Android         = $ReportItem.Android
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}

function Get-MgaReportOffice365ActivationsUserCounts {
    [CmdletBinding()]
    param (
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        $URL = "https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserCounts" 
        Write-Verbose "Get-MgaReportOffice365ActivationsUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOffice365ActivationsUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOffice365ActivationsUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{        
                ProductType              = $ReportItem.'Product Type'
                Assigned                 = $ReportItem.Assigned
                Activated                = $ReportItem.Activated
                SharedComputerActivation = $ReportItem.'Shared Computer Activation'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}
#endregion Microsoft 365 activations
#region Microsoft 365 groups activity
function Get-MgaReportOffice365GroupsActivityDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportOffice365GroupsActivityDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportOffice365GroupsActivityDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOffice365GroupsActivityDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOffice365GroupsActivityDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{  
                GroupDisplayName               = $ReportItem.'Group Display Name'
                IsDeleted                      = $ReportItem.'Is Deleted'
                OwnerPrincipalName             = $ReportItem.'Owner Principal Name'
                LastActivityDate               = $ReportItem.'Last Activity Date'
                GroupType                      = $ReportItem.'Group Type'
                MemberCount                    = $ReportItem.'Member Count'
                ExternalMemberCount            = $ReportItem.'External Member Count'
                ExchangeReceivedEmailCount     = $ReportItem.'Exchange Received Email Count'
                SharePointActiveFileCount      = $ReportItem.'SharePoint Active File Count'
                YammerPostedMessageCount       = $ReportItem.'Yammer Posted Message Count'
                YammerReadMessageCount         = $ReportItem.'Yammer Read Message Count'
                YammerLikedMessageCount        = $ReportItem.'Yammer Liked Message Count'
                ExchangeMailboxTotalItemCount  = $ReportItem.'Exchange Mailbox Total Item Count'
                ExchangeMailboxStorageUsedInGb = [math]::Round(((($ReportItem.'Exchange Mailbox Storage Used (Byte)') / 1024 / 1024 / 1024)), 2)
                SharePointTotalFileCount       = $ReportItem.'SharePoint Total File Count'
                SharePointSiteStorageUsedInGb  = [math]::Round(((($ReportItem.'SharePoint Site Storage Used (Byte)') / 1024 / 1024 / 1024)), 2)
                GroupId                        = $ReportItem.'Group Id'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}

function Get-MgaReportOffice365GroupsActivityCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportOffice365GroupsActivityCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportOffice365GroupsActivityCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOffice365GroupsActivityCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOffice365GroupsActivityCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{  
                ExchangeEmailsReceived = $ReportItem.'Exchange Emails Received'
                YammerMessagesPosted   = $ReportItem.'Yammer Messages Posted'
                YammerMessagesRead     = $ReportItem.'Yammer Messages Read'
                YammerMessagesLiked    = $ReportItem.'Yammer Messages Liked'
                ReportDate             = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}

function Get-MgaReportOffice365GroupsActivityGroupCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportOffice365GroupsActivityGroupCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityGroupCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportOffice365GroupsActivityGroupCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOffice365GroupsActivityGroupCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOffice365GroupsActivityGroupCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{  
                Total      = $ReportItem.Total
                Active     = $ReportItem.Active
                ReportDate = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}

function Get-MgaReportOffice365GroupsActivityStorage {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportOffice365GroupsActivityStorage: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityStorage(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportOffice365GroupsActivityStorage: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOffice365GroupsActivityStorage: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOffice365GroupsActivityStorage: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{  
                MailboxStorageUsedInGb = [math]::Round(((($ReportItem.'Mailbox Storage Used (Byte)') / 1024 / 1024 / 1024)), 2)
                SiteStorageUsedInGb    = [math]::Round(((($ReportItem.'Site Storage Used (Byte)') / 1024 / 1024 / 1024)), 2)
                ReportDate             = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}

function Get-MgaReportOffice365GroupsActivityFileCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportOffice365GroupsActivityFileCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityFileCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportOffice365GroupsActivityFileCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOffice365GroupsActivityFileCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOffice365GroupsActivityFileCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{  
                Total      = $ReportItem.Total
                Active     = $ReportItem.Active
                ReportDate = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}
#endregion Microsoft 365 groups activity
#region OneDrive activity
function Get-MgaReportOneDriveActivityUserDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportOneDriveActivityUserDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getOneDriveActivityUserDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportOneDriveActivityUserDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOneDriveActivityUserDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOneDriveActivityUserDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $AssignedProducts = $null
            $AssignedProducts = $ReportItem.'Assigned Products'.Split('+').Trim()
            $Object = @{  
                UserPrincipalName         = $ReportItem.'User Principal Name'
                IsDeleted                 = $ReportItem.'Is Deleted'
                LastActivityDate          = $ReportItem.'Last Activity Date'
                ViewedOrEditedFileCount   = $ReportItem.'Viewed Or Edited File Count'
                SyncedFileCount           = $ReportItem.'Synced File Count'
                SharedInternallyFileCount = $ReportItem.'Shared Internally File Count'
                SharedExternallyFileCount = $ReportItem.'Shared Externally File Count'
                AssignedProducts          = $AssignedProducts
            }
            if ($ReportItem.'Is Deleted' -eq $true) {
                $Object.DeletedDate = $ReportItem.'Deleted Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}

function Get-MgaReportOneDriveActivityUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportOneDriveActivityUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getOneDriveActivityUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportOneDriveActivityUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOneDriveActivityUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOneDriveActivityUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{  
                ViewedOrEdited   = $ReportItem.'Viewed Or Edited'
                Synced           = $ReportItem.'Synced'
                SharedInternally = $ReportItem.'Shared Internally'
                SharedExternally = $ReportItem.'Shared Externally'
                ReportDate       = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}

function Get-MgaReportOneDriveActivityFileCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )  
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportOneDriveActivityFileCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getOneDriveActivityFileCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportOneDriveActivityFileCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }  
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOneDriveActivityFileCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOneDriveActivityFileCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{  
                ViewedOrEdited   = $ReportItem.'Viewed Or Edited'
                Synced           = $ReportItem.'Synced'
                SharedInternally = $ReportItem.'Shared Internally'
                SharedExternally = $ReportItem.'Shared Externally'
                ReportDate       = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    }
    end {
        return $ReportsList
    }
}
#endregion OneDrive activity
#region OneDrive usage
function Get-MgaReportOneDriveUsageAccountDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportOneDriveUsageAccountDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportOneDriveUsageAccountDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOneDriveUsageAccountDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOneDriveUsageAccountDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $SizeInGb = $null
            $AllocatedInGb = $null
            $SizeInGb = [math]::Round(((($ReportItem.'Storage Used (Byte)') / 1024 / 1024 / 1024)), 2)
            $AllocatedInGb = [math]::Round(((($ReportItem.'Storage Allocated (Byte)') / 1024 / 1024 / 1024)), 2)
            $Object = @{
                DisplayName       = $ReportItem.'Owner Display Name'
                UserPrincipalName = $ReportItem.'Owner Principal Name'
                SiteUrl           = $ReportItem.'Site URL'
                SizeInGb          = $SizeInGb
                AllocatedSizeInGb = $AllocatedInGb
                FreeSizeInGb      = ($AllocatedInGb - $SizeInGb)
                IsDeleted         = $ReportItem.'Is Deleted'
                LastActivityDate  = $ReportItem.'Last Activity Date'
                FileCount         = $ReportItem.'File Count'
                ActiveFileCount   = $ReportItem.'Active File Count'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportOneDriveUsageAccountCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportOneDriveUsageAccountCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportOneDriveUsageAccountCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOneDriveUsageAccountCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOneDriveUsageAccountCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                SiteType   = $ReportItem.'Site Type'
                Total      = $ReportItem.Total
                Active     = $ReportItem.Active
                ReportDate = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportOneDriveUsageFileCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportOneDriveUsageFileCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageFileCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportOneDriveUsageFileCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOneDriveUsageFileCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOneDriveUsageFileCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                SiteType   = $ReportItem.'Site Type'
                Total      = $ReportItem.Total
                Active     = $ReportItem.Active
                ReportDate = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportOneDriveUsageStorage {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportOneDriveUsageStorage: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageStorage(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportOneDriveUsageStorage: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportOneDriveUsageStorage: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportOneDriveUsageStorage: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                SiteType        = $ReportItem.'Site Type'
                StorageUsedInGb = [math]::Round(((($ReportItem.'Storage Used (Byte)') / 1024 / 1024 / 1024)), 2)
                ReportDate      = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}
#endregion OneDrive usage
#region SharePoint activity
function Get-MgaReportSharePointActivityUserDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSharePointActivityUserDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSharePointActivityUserDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSharePointActivityUserDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSharePointActivityUserDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSharePointActivityUserDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $AssignedProducts = $null
            $AssignedProducts = $ReportItem.'Assigned Products'.Split('+').Trim()
            $Object = @{
                UserPrincipalName         = $ReportItem.'User Principal Name'
                IsDeleted                 = $ReportItem.'Is Deleted'
                LastActivityDate          = $ReportItem.'Last Activity Date'
                ViewedOrEditedFileCount   = $ReportItem.'Viewed Or Edited File Count'
                SyncedFileCount           = $ReportItem.'Synced File Count'
                SharedInternallyFileCount = $ReportItem.'Shared Internally File Count'
                SharedExternallyFileCount = $ReportItem.'Shared Externally File Count'
                AssignedProducts          = $AssignedProducts
            }
            if ($ReportItem.'Is Deleted' -eq $true) {
                $Object.DeletedDate = $ReportItem.'Deleted Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSharePointActivityFileCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSharePointActivityFileCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSharePointActivityFileCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSharePointActivityFileCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSharePointActivityFileCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSharePointActivityFileCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                ViewedOrEditedFileCount   = $ReportItem.'Viewed Or Edited'
                SyncedFileCount           = $ReportItem.'Synced'
                SharedInternallyFileCount = $ReportItem.'Shared Internally'
                SharedExternallyFileCount = $ReportItem.'Shared Externally'
                ReportDate                = $ReportItem.'report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSharePointActivityUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSharePointActivityUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSharePointActivityUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSharePointActivityUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSharePointActivityUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSharePointActivityUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                VisitedPage               = $ReportItem.'Visited Page'
                ViewedOrEditedFileCount   = $ReportItem.'Viewed Or Edited'
                SyncedFileCount           = $ReportItem.'Synced'
                SharedInternallyFileCount = $ReportItem.'Shared Internally'
                SharedExternallyFileCount = $ReportItem.'Shared Externally'
                ReportDate                = $ReportItem.'report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSharePointActivityPages {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSharePointActivityPages: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSharePointActivityPages(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSharePointActivityPages: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSharePointActivityPages: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSharePointActivityPages: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                VisitedPageCount = $ReportItem.'Visited Page Count'
                ReportDate       = $ReportItem.'report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}
#endregion SharePoint activity
#region SharePoint site usage
function Get-MgaReportSharePointSiteUsageDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSharePointSiteUsageDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSharePointSiteUsageDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSharePointSiteUsageDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSharePointSiteUsageDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $SizeInGb = $null
            $AllocatedInGb = $null
            $SizeInGb = [math]::Round(((($ReportItem.'Storage Used (Byte)') / 1024 / 1024 / 1024)), 2)
            $AllocatedInGb = [math]::Round(((($ReportItem.'Storage Allocated (Byte)') / 1024 / 1024 / 1024)), 2)
            $Object = @{
                SiteId               = $ReportItem.'Site Id'
                SiteURL              = $ReportItem.'Site URL'
                OwnerDisplayName     = $ReportItem.'Owner Display Name'
                IsDeleted            = $ReportItem.'Is Deleted'
                LastActivityDate     = $ReportItem.'Last Activity Date'
                FileCount            = $ReportItem.'File Count'
                ActiveFileCount      = $ReportItem.'Active File Count'
                PageViewCount        = $ReportItem.'Page View Count'
                VisitedPageCount     = $ReportItem.'Visited Page Count'
                StorageUsedInGb      = $SizeInGb
                StorageAllocatedInGb = $AllocatedInGb
                FreeSizeInGb         = ($AllocatedInGb - $SizeInGb)
                RootWebTemplate      = $ReportItem.'Root Web Template'
                OwnerPrincipalName   = $ReportItem.'Owner Principal Name'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSharePointSiteUsageFileCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSharePointSiteUsageFileCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageFileCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSharePointSiteUsageFileCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSharePointSiteUsageFileCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSharePointSiteUsageFileCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                SiteType   = $ReportItem.'Site Type'
                Total      = $ReportItem.Total
                Active     = $ReportItem.Active
                ReportDate = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSharePointSiteUsageSiteCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSharePointSiteUsageSiteCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageSiteCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSharePointSiteUsageSiteCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSharePointSiteUsageSiteCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSharePointSiteUsageSiteCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                SiteType   = $ReportItem.'Site Type'
                Total      = $ReportItem.Total
                Active     = $ReportItem.Active
                ReportDate = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSharePointSiteUsageStorage {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSharePointSiteUsageStorage: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageStorage(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSharePointSiteUsageStorage: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSharePointSiteUsageStorage: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSharePointSiteUsageStorage: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                SiteType        = $ReportItem.'Site Type'
                StorageUsedInGb = [math]::Round(((($ReportItem.'Storage Used (Byte)') / 1024 / 1024 / 1024)), 2)
                ReportDate      = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSharePointSiteUsagePages {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSharePointSiteUsagePages: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsagePages(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSharePointSiteUsagePages: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSharePointSiteUsagePages: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSharePointSiteUsagePages: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                SiteType      = $ReportItem.'Site Type'
                PageViewCount = $ReportItem.'Page View Count'
                ReportDate    = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}
#endregion SharePoint site usage
#region Skype for Business activity
function Get-MgaReportSkypeForBusinessActivityUserDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessActivityUserDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessActivityUserDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessActivityUserDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessActivityUserDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessActivityUserDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $AssignedProducts = $null
            $AssignedProducts = $ReportItem.'Assigned Products'.Split('+').Trim()
            $Object = @{
                UserPrincipalName                            = $ReportItem.'User Principal Name'
                IsDeleted                                    = $ReportItem.'Is Deleted'
                LastActivityDate                             = $ReportItem.'Last Activity Date'
                TotalPeertopeerSessionCount                  = $ReportItem.'Total Peer-to-peer Session Count'
                TotalOrganizedConferenceCount                = $ReportItem.'Total Organized Conference Count'
                TotalParticipatedConferenceCount             = $ReportItem.'Total Participated Conference Count'
                PeertopeerLastActivityDate                   = $ReportItem.'Peer-to-peer Last Activity Date'
                OrganizedConferenceLastActivityDate          = $ReportItem.'Organized Conference Last Activity Date'
                ParticipatedConferenceLastActivityDate       = $ReportItem.'Participated Conference Last Activity Date'
                PeertopeerIMCount                            = $ReportItem.'Peer-to-peer IM Count'
                PeertopeerAudioCount                         = $ReportItem.'Peer-to-peer Audio Count'
                PeertopeerAudioMinutes                       = $ReportItem.'Peer-to-peer Audio Minutes'
                PeertopeerVideoCount                         = $ReportItem.'Peer-to-peer Video Count'
                PeertopeerVideoMinutes                       = $ReportItem.'Peer-to-peer Video Minutes'
                PeertopeerAppSharingCount                    = $ReportItem.'Peer-to-peer App Sharing Count'
                PeertopeerFileTransferCount                  = $ReportItem.'Peer-to-peer File Transfer Count'
                OrganizedConferenceIMCount                   = $ReportItem.'Organized Conference IM Count'
                OrganizedConferenceAudioVideoCount           = $ReportItem.'Organized Conference Audio/Video Count'
                OrganizedConferenceAudioVideoMinutes         = $ReportItem.'Organized Conference Audio/Video Minutes'
                OrganizedConferenceAppSharingCount           = $ReportItem.'Organized Conference App Sharing Count'
                OrganizedConferenceWebCount                  = $ReportItem.'Organized Conference Web Count'
                OrganizedConferenceDialinout3rdPartyCount    = $ReportItem.'Organized Conference Dial-in/out 3rd Party Count'
                OrganizedConferenceDialinoutMicrosoftCount   = $ReportItem.'Organized Conference Dial-in/out Microsoft Count'
                OrganizedConferenceDialinMicrosoftMinutes    = $ReportItem.'Organized Conference Dial-in Microsoft Minutes'
                OrganizedConferenceDialoutMicrosoftMinutes   = $ReportItem.'Organized Conference Dial-out Microsoft Minutes'
                ParticipatedConferenceIMCount                = $ReportItem.'Participated Conference IM Count'
                ParticipatedConferenceAudioVideoCount        = $ReportItem.'Participated Conference Audio/Video Count'
                ParticipatedConferenceAudioVideoMinutes      = $ReportItem.'Participated Conference Audio/Video Minutes'
                ParticipatedConferenceAppSharingCount        = $ReportItem.'Participated Conference App Sharing Count'
                ParticipatedConferenceWebCount               = $ReportItem.'Participated Conference Web Count'
                ParticipatedConferenceDialinout3rdPartyCount = $ReportItem.'Participated Conference Dial-in/out 3rd Party Count'
                AssignedProducts                             = $AssignedProducts
            }
            if ($ReportItem.'Is Deleted' -eq $true) {
                $Object.DeletedDate = $ReportItem.'Deleted Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSkypeForBusinessActivityCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessActivityCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessActivityCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessActivityCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessActivityCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessActivityCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                ReportDate   = $ReportItem.'Report Date'
                Peertopeer   = $ReportItem.'Peer-to-peer'
                Organized    = $ReportItem.Organized
                Participated = $ReportItem.Participated
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSkypeForBusinessActivityUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessActivityUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessActivityUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessActivityUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessActivityUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessActivityUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                ReportDate   = $ReportItem.'Report Date'
                Peertopeer   = $ReportItem.'Peer-to-peer'
                Organized    = $ReportItem.Organized
                Participated = $ReportItem.Participated
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}
#endregion Skype for Business activity
#region Skype for Business device usage
function Get-MgaReportSkypeForBusinessDeviceUsageUserDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessDeviceUsageUserDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessDeviceUsageUserDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessDeviceUsageUserDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessDeviceUsageUserDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessDeviceUsageUserDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                UserPrincipalName = $ReportItem.'User Principal Name'
                LastActivityDate  = $ReportItem.'Last Activity Date'
                UsedWindows       = $ReportItem.'Used Windows'
                UsedWindowsPhone  = $ReportItem.'Used Windows Phone'
                UsedAndroidPhone  = $ReportItem.'Used Android Phone'
                UsediPhone        = $ReportItem.'Used iPhone'
                UsediPad          = $ReportItem.'Used iPad'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSkypeForBusinessDeviceUsageDistributionUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessDeviceUsageDistributionUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessDeviceUsageDistributionUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessDeviceUsageDistributionUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessDeviceUsageDistributionUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessDeviceUsageDistributionUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                Windows      = $ReportItem.Windows
                WindowsPhone = $ReportItem.'Windows Phone'
                AndroidPhone = $ReportItem.'Android Phone'
                iPhone       = $ReportItem.Iphone
                iPad         = $ReportItem.Ipad
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSkypeForBusinessDeviceUsageUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessDeviceUsageUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessDeviceUsageUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessDeviceUsageUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessDeviceUsageUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessDeviceUsageUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                Windows      = $ReportItem.Windows
                WindowsPhone = $ReportItem.'Windows Phone'
                AndroidPhone = $ReportItem.'Android Phone'
                iPhone       = $ReportItem.Iphone
                iPad         = $ReportItem.Ipad
                ReportDate   = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}
#endregion Skype for Business device usage
#region Skype for Business organizer activity
function Get-MgaReportSkypeForBusinessOrganizerActivityCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessOrganizerActivityCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessOrganizerActivityCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessOrganizerActivityCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessOrganizerActivityCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessOrganizerActivityCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                IM                 = $ReportItem.'IM'
                AudioVideo         = $ReportItem.'Audio/Video'
                AppSharing         = $ReportItem.'App Sharing'
                Web                = $ReportItem.'Web'
                Dialinout3rdParty  = $ReportItem.'Dial-in/out 3rd Party'
                DialinoutMicrosoft = $ReportItem.'Dial-in/out Microsoft'
                ReportDate         = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSkypeForBusinessOrganizerActivityUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessOrganizerActivityUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessOrganizerActivityUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessOrganizerActivityUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessOrganizerActivityUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessOrganizerActivityUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                IM                 = $ReportItem.'IM'
                AudioVideo         = $ReportItem.'Audio/Video'
                AppSharing         = $ReportItem.'App Sharing'
                Web                = $ReportItem.'Web'
                Dialinout3rdParty  = $ReportItem.'Dial-in/out 3rd Party'
                DialinoutMicrosoft = $ReportItem.'Dial-in/out Microsoft'
                ReportDate         = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSkypeForBusinessOrganizerActivityMinuteCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessOrganizerActivityMinuteCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessOrganizerActivityMinuteCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessOrganizerActivityMinuteCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessOrganizerActivityMinuteCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessOrganizerActivityMinuteCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                AudioVideo       = $ReportItem.'Audio/Video'
                ReportDate       = $ReportItem.'Report Date'
                DialinMicrosoft  = $ReportItem.'Dial-in Microsoft'
                DialoutMicrosoft = $ReportItem.'Dial-out Microsoft'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}
#endregion Skype for Business organizer activity
#region Skype for Business participant activity
function Get-MgaReportSkypeForBusinessParticipantActivityCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessParticipantActivityCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessParticipantActivityCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessParticipantActivityCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessParticipantActivityCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessParticipantActivityCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                IM                = $ReportItem.'IM'
                AudioVideo        = $ReportItem.'Audio/Video'
                AppSharing        = $ReportItem.'App Sharing'
                Web               = $ReportItem.'Web'
                Dialinout3rdParty = $ReportItem.'Dial-in/out 3rd Party'
                ReportDate        = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSkypeForBusinessParticipantActivityUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessParticipantActivityUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessParticipantActivityUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessParticipantActivityUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessParticipantActivityUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessParticipantActivityUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                IM                = $ReportItem.'IM'
                AudioVideo        = $ReportItem.'Audio/Video'
                AppSharing        = $ReportItem.'App Sharing'
                Web               = $ReportItem.'Web'
                Dialinout3rdParty = $ReportItem.'Dial-in/out 3rd Party'
                ReportDate        = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSkypeForBusinessParticipantActivityMinuteCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessParticipantActivityMinuteCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessParticipantActivityMinuteCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessParticipantActivityMinuteCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessParticipantActivityMinuteCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessParticipantActivityMinuteCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                AudioVideo = $ReportItem.'Audio/Video'
                ReportDate = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}
#endregion Skype for Business participant activity
#region Skype for Business peer-to-peer activity
function Get-MgaReportSkypeForBusinessPeerToPeerActivityCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessPeerToPeerActivityCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessPeerToPeerActivityCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessPeerToPeerActivityCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessPeerToPeerActivityCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessPeerToPeerActivityCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                IM           = $ReportItem.IM
                Audio        = $ReportItem.Audio
                Video        = $ReportItem.Video
                AppSharing   = $ReportItem.'App Sharing'
                FileTransfer = $ReportItem.'File Transfer'
                ReportDate   = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSkypeForBusinessPeerToPeerActivityUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessPeerToPeerActivityUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessPeerToPeerActivityUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessPeerToPeerActivityUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessPeerToPeerActivityUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessPeerToPeerActivityUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                IM           = $ReportItem.IM
                Audio        = $ReportItem.Audio
                Video        = $ReportItem.Video
                AppSharing   = $ReportItem.'App Sharing'
                FileTransfer = $ReportItem.'File Transfer'
                ReportDate   = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportSkypeForBusinessPeerToPeerActivityMinuteCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportSkypeForBusinessPeerToPeerActivityMinuteCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getSkypeForBusinessPeerToPeerActivityMinuteCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportSkypeForBusinessPeerToPeerActivityMinuteCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportSkypeForBusinessPeerToPeerActivityMinuteCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportSkypeForBusinessPeerToPeerActivityMinuteCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                Audio      = $ReportItem.Audio
                Video      = $ReportItem.Video
                ReportDate = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}
#endregion Skype for Business peer-to-peer activity
#region Yammer activity
function Get-MgaReportYammerActivityUserDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportYammerActivityUserDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getYammerActivityUserDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportYammerActivityUserDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportYammerActivityUserDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportYammerActivityUserDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $AssignedProducts = $null
            $AssignedProducts = $ReportItem.'Assigned Products'.Split('+').Trim()
            $Object = @{
                UserPrincipalName = $ReportItem.'User Principal Name'
                DisplayName       = $ReportItem.'Display Name'
                UserState         = $ReportItem.'User State'
                StateChangeDate   = $ReportItem.'State Change Date'
                LastActivityDate  = $ReportItem.'Last Activity Date'
                PostedCount       = $ReportItem.'Posted Count'
                ReadCount         = $ReportItem.'Read Count'
                LikedCount        = $ReportItem.'Liked Count'
                AssignedProducts  = $AssignedProducts
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportYammerActivityCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportYammerActivityCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getYammerActivityCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportYammerActivityCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportYammerActivityCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportYammerActivityCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = [PSCustomObject]@{
                Liked      = $ReportItem.Liked
                Posted     = $ReportItem.Posted
                Read       = $ReportItem.Read
                ReportDate = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportYammerActivityUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportYammerActivityUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getYammerActivityUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportYammerActivityUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportYammerActivityUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportYammerActivityUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                Liked      = $ReportItem.Liked
                Posted     = $ReportItem.Posted
                Read       = $ReportItem.Read
                ReportDate = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}
#endregion Yammer activity
#region Yammer device usage
function Get-MgaReportYammerDeviceUsageUserDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportYammerDeviceUsageUserDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getYammerDeviceUsageUserDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportYammerDeviceUsageUserDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportYammerDeviceUsageUserDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportYammerDeviceUsageUserDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                UserPrincipalName = $ReportItem.'User Principal Name'
                DisplayName       = $ReportItem.'Display Name'
                UserState         = $ReportItem.'User State'
                StateChangeDate   = $ReportItem.'State Change Date'
                LastActivityDate  = $ReportItem.'Last Activity Date'
                UsedWeb           = $ReportItem.'Used Web'
                UsedWindowsPhone  = $ReportItem.'Used Windows Phone'
                UsedAndroidPhone  = $ReportItem.'Used Android Phone'
                UsediPhone        = $ReportItem.'Used iPhone'
                UsediPad          = $ReportItem.'Used iPad'
                UsedOthers        = $ReportItem.'Used Others'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportYammerDeviceUsageDistributionUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportYammerDeviceUsageDistributionUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getYammerDeviceUsageDistributionUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportYammerDeviceUsageDistributionUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportYammerDeviceUsageDistributionUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportYammerDeviceUsageDistributionUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                Web          = $ReportItem.'Web'
                WindowsPhone = $ReportItem.'Windows Phone'
                AndroidPhone = $ReportItem.'Android Phone'
                iPhone       = $ReportItem.'iPhone'
                iPad         = $ReportItem.'iPad'
                Other        = $ReportItem.'Other'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportYammerDeviceUsageUserCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportYammerDeviceUsageUserCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getYammerDeviceUsageUserCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportYammerDeviceUsageUserCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportYammerDeviceUsageUserCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportYammerDeviceUsageUserCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                Web          = $ReportItem.'Web'
                WindowsPhone = $ReportItem.'Windows Phone'
                AndroidPhone = $ReportItem.'Android Phone'
                iPhone       = $ReportItem.'iPhone'
                iPad         = $ReportItem.'iPad'
                Other        = $ReportItem.'Other'
                ReportDate   = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}
#endregion Yammer device usage
#region Yammer groups activity
function Get-MgaReportYammerGroupsActivityDetail {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportYammerGroupsActivityDetail: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getYammerGroupsActivityDetail(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportYammerGroupsActivityDetail: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportYammerGroupsActivityDetail: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportYammerGroupsActivityDetail: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                GroupDisplayName   = $ReportItem.'Group Display Name'
                IsDeleted          = $ReportItem.'Is Deleted'
                OwnerPrincipalName = $ReportItem.'Owner Principal Name'
                LastActivityDate   = $ReportItem.'Last Activity Date'
                GroupType          = $ReportItem.'Group Type'
                Office365Connected = $ReportItem.'Office 365 Connected'
                MemberCount        = $ReportItem.'Member Count'
                PostedCount        = $ReportItem.'Posted Count'
                ReadCount          = $ReportItem.'Read Count'
                LikedCount         = $ReportItem.'Liked Count'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportYammerGroupsActivityGroupCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
       begin {
             $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportYammerGroupsActivityGroupCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getYammerGroupsActivityGroupCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportYammerGroupsActivityGroupCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportYammerGroupsActivityGroupCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportYammerGroupsActivityGroupCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                Total      = $ReportItem.Total
                Active     = $ReportItem.Active
                ReportDate = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}

function Get-MgaReportYammerGroupsActivityCounts {
    [CmdletBinding()]
    param (
        [ValidateSet("7", "30", "90", "180")]
        $Period = '7'
    )
    begin {
        $null = Test-MgaRole -RoleType 'O365'
        Write-Verbose "Get-MgaReportYammerGroupsActivityCounts: begin: Period: D$Period"
        $URL = "https://graph.microsoft.com/v1.0/reports/getYammerGroupsActivityCounts(period='D{0}')" -f $Period
        Write-Verbose "Get-MgaReportYammerGroupsActivityCounts: begin: URL: $URL"
        $ReportsList = [System.Collections.Generic.List[Object]]::new()
    }
    process {
        $GetMga = Get-Mga $URL
        Write-Verbose "Get-MgaReportYammerGroupsActivityCounts: Getting results | Result count: $($GetMga.Count)"
        Write-Verbose "Get-MgaReportYammerGroupsActivityCounts: Converting results"
        foreach ($ReportItem in $GetMga) {
            $Object = @{
                Liked      = $ReportItem.Liked
                Posted     = $ReportItem.Posted
                Read       = $ReportItem.Read
                ReportDate = $ReportItem.'Report Date'
            }
            $ReportsList.Add([PSCustomObject]$Object)
        }
    } 
    end {
        return $ReportsList
    }
}
#endregion Yammer groups activity
#endregion 365 usage reports
#endregion functions
#region internal functions
function Test-MgaRole {
    [CmdletBinding()]
    param (
        [parameter(mandatory = $true)]
        [ValidateSet('O365', 'AzureAD')]
        $RoleType
    )   
    begin {
        $SufficientPermissions = $false
        if ((($RoleType -eq 'AzureAD') -and ($global:MgaReportRoleAzureAD -eq $true)) -or (($RoleType -eq 'O365') -and ($global:MgaReportRoleO365 -eq $true))) {
            $SufficientPermissions = $true
        }
    }
    process {
        if ($SufficientPermissions -ne $true) {
            $Roles = Show-MgaAccessToken -Roles
            if ($RoleType -eq 'O365') {
                $ReportsLog = $false
                foreach ($Role in $Roles) {
                    if ($Role -eq 'Reports.Read.All') {
                        $ReportsLog = $true
                    }
                }
                if ($ReportsLog -eq $true) {
                    $global:MgaReportRoleO365 = $true
                }
                else {
                    throw "Missing Reports.Read.All permission for Microsoft Graph API to call Reports for $RoleType"
                }
            }
            elseif ($RoleType -eq 'AzureAD') {
                $AuditLog = $false
                $DirectoryLog = $false
                foreach ($Role in $Roles) {
                    if ($Role -eq 'AuditLog.Read.All') {
                        $AuditLog = $true
                    }
                    elseif (($Role -eq 'Directory.ReadWrite.All') -or ($Role -eq 'Directory.Read.All')) {
                        $DirectoryLog = $true
                    }
                }
                if (($AuditLog -eq $true) -and ($DirectoryLog -eq $true)) {
                    $global:MgaReportRoleAzureAD = $true
                }
                elseif (($AuditLog -eq $false) -and ($DirectoryLog -eq $false)) {
                    throw  "Missing AuditLog.Read.All and Directory.ReadWrite.All or Directory.Read.All permission for Microsoft Graph API to call Reports for $RoleType"
                }
                elseif (($AuditLog -eq $true) -and ($DirectoryLog -eq $false)) {
                    throw  "Missing Directory.ReadWrite.All or Directory.Read.All permission for Microsoft Graph API to call Reports for $RoleType"
                }
                elseif (($AuditLog -eq $false) -and ($DirectoryLog -eq $true)) {
                    throw  "Missing AuditLog.Read.All permission for Microsoft Graph API to call Reports for $RoleType"
                }
            }
        }
    }  
    end {
        if ((($RoleType -eq 'AzureAD') -and ($global:MgaReportRoleAzureAD -eq $true)) -or (($RoleType -eq 'O365') -and ($global:MgaReportRoleO365 -eq $true))) {
            return "Sufficient roles found for $RoleType reports"
        }
    }
}
#endregion internal functions