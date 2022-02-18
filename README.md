# Optimized.Mga.Report
This is a submodule for [Optimized.Mga](https://github.com/baswijdenes/Optimized.Mga).  
Optimized.Mga.Mail is dependant on Optimized.Mga

- I've added all the Report references in their own cmdlets 
- I updated several properties to contain arrays in stead of strings (Examples are Assigned Products)
- I removed spaces where possible in the property names
- For each cmdlet it will check if you have the right permissions and otherwise tell you which permission you need

## How can I interpret the cmdlets?
You can paste the title where the first UPPERCASE starts after `Get-MgaReport`.  
When we use (reportRoot: getMailboxUsageDetail)[https://docs.microsoft.com/en-us/graph/api/reportroot-getmailboxusagedetail?view=graph-rest-1.0] as an example, the cmdlet would be: `Get-MgaReportMailboxUsageDetail`.  
This cmdlet will return the content as a PSCustomObject.