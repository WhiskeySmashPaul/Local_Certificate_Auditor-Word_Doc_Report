#region Modules Reduired

Import-Module WebAdministration -ErrorAction SilentlyContinue
Import-Module PSWriteWord -force -ErrorAction SilentlyContinue

#endregion modules required

#region Adjustable Variables - Please adjust these variables to your enviorment to get the information you need from the report.


#Change the $DaystoSearch variable for how many days you would like to know in advance for the expiring certificates.
$DaystoSearch = "<Enter number of days to search for here>"

#Change the $Servers variable to a varaint of Get-ADComputers to ensure you get a list of servers you are inquiring upon. Can be more or less restrictive.
$servers = Get-ADComputer -Filter "OperatingSystem -like 'Windows Server*' -and Enabled -eq '$true'" | Sort-Object |Foreach {$_.Name}

#Change the $Path variable to the path where you would like Doc Report to be saved.
$Path = "<Enter a path to a central location for reports to be stored>"

#Change the $DOCReport variable to the naming convention of your choosing - As set below this will create a new folder for each Month/Year for historical purposes ie \\fileserver\2022-June\180 Day Report.docx
$DOCreport = "$Path\$YearMonth\$currentDay Certificate Expiration Report - $DaystoSearch Days.docx"


#endregion Adjustable Variables

#region Static Variables - These variables are not needed to be adjusted as they are just formatting for later in the script. Can adjust the Get-Date format if not liking MM/dd/yyyy format

#Sets variables for the enviorment
$today = ((Get-Date).AddDays($DaystoSearch)).tostring("MM-dd-yyyy")
$otoday = Get-Date -Format "MM/dd/yyy"
$currentDay = Get-Date -Format "MM-dd-yyyy"
$currentMonth = Get-Date -UFormat %m
$currentYear = Get-Date -Format "yyyy"
$currentMonth = (Get-Culture).DateTimeFormat.GetMonthName($currentMonth)
$MonthYear = $currentMonth + '-' + $currentYear
$YearMonth = $currentYear + '-' + $currentMonth

#endregion Static Variables

#region New Report

#Creates the directory for the report to be saved
New-Item -Path $Path -Name $YearMonth -ItemType Directory

#endregion of new report

#region Setup of New Report

$newdoc = New-WordDocument $DOCreport -Verbose

#region Add Header

<#
#Uncomment section if you would like to unclude a header image on the word doc report
$logoimage = "$dir\Header-Logo.png"
Add-WordPicture -WordDocument $newdoc -ImagePath $logoimage -Alignment right -ImageWidth 170 -ImageHeight 80
#>

#endregion of Add Header

#region Add Verbage to word doc

#This section adds all the text to the report. Verbage and styling can be changed as see fit.
Add-WordText -WordDocument $newdoc -Text "Report compiled on $otoday." -Alignment Right -FontSize 08 -FontFamily 'Calibri' -Color darkgray -HeadingType Heading1 
Add-wordtext -WordDocument $newdoc -Text "Certificates" -FontSize 18 -fontfamily 'Bahnschrift Condensed'  -Color Black
add-wordtext -WordDocument $newdoc -Text "The following report details all certificate that will expire before $today, please review and make relevant arrangements to replace, or renew these certificates to reduce the risk of a service outage." -FontSize 12 -fontfamily 'Bahnschrift Light SemiCondensed'

Add-WordParagraph -WordDocument $newdoc
Add-wordtext -WordDocument $newdoc -Text "Certificates Expiring Before $today" -FontSize 18 -fontfamily 'Bahnschrift Condensed' -Color Black

#endregion of Add Verbage to word doc

#region Add table to report

$T = New-wordtable -WordDocument $newdoc -NrColumns 6 -NrRows 1 
Add-WordTableRow -Table $t -Index 1
$expired = @()

#endregion of add table to report

#endregion of Setup of New Report

#region Get certificate information

foreach ($server in $servers) { #each server start
	$certs = Invoke-Command -ComputerName $server -ScriptBlock { Get-ChildItem -path Cert:\LocalMachine\My -Recurse -erroraction SilentlyContinue | select-Object Subject, Issuer, FriendlyName, Thumbprint, NotBefore, NotAfter  -ExcludeProperty PSComputerName, RunspaceId, PSSHowComputerName | Sort-Object Notafter -Descending } -ErrorAction SilentlyContinue
    
    foreach ($c in $certs) { #each cert start
        $subs = $c.subject.split(",")
        foreach ($sub in $subs) {
            if ($sub -like "*CN=*") {
                $sj = $sub -replace 'CN=', ''
                $SJ = $sj -replace '\s',''
                }
        }
        $edate = $c.NotAfter.ToString("MM-dd-yyyy")
        $ts = New-TimeSpan -Start $edate -End $today
        if ($ts.days -gt 0 -and $ts.Days -lt $DaystoSearch) {
            $item = [PSCustomobject]@{
                Name = $server
                Issuer = $c.Issuer
                Thumbprint = $c.Thumbprint
                'Issued Date' = $c.NotBefore.ToString("MM-dd-yyyy")
                'Expiration Date' = $c.NotAfter.tostring("MM-dd-yyyy")
            }
            $expired = $expired + $item
        }
    } #each cert end
} #each server end

#endregion Get certificate information

#region Add Server certificate information to report

$sortexpired = $expired | Sort-Object $expired.Name
$sortexpired | Format-Table -AutoSize

Add-WordTable -Table $t -DataTable $sortexpired -Design MediumShading1Accent2 -Alignment left -AutoFit Window -FontSize 10, 7 -ContinueFormatting 
Add-wordline -WordDocument $newdoc -LineColor darkgray -LineType single -LineSpace 0.2

#endregion of Add Server certificate information to report

#region Add footer image - uncomment this section if requiring footer images

<#
$footerimage = "$dir\Footer-Logo.png"
add your footer image
Add-WordPicture -WordDocument $newdoc -ImagePath $footerimage -Alignment right -ImageWidth 400 -ImageHeight 85
#>

#endregion Add footer image

#region Save File

#save the doc file
Save-WordDocument -WordDocument $newdoc #-KillWord

#endregion Save file

#region Email report

#Send report to someone

$From = "CertificateAuditor@DOMAIN.com"
$To = "<Enter email address of who to send when report finishes>"
$Attachment = "$DOCreport"
$Subject = "Certificates expiring within $DaystoSearch days"
$Body = "<h3><u>Certificates expiring within the next $DaystoSearch days.</u></h3><p>Please find attached the DOCX report detailing certificates that will expire within the next $DaystoSearch days<BR>Please review the report and make relevant arrangements to replace, or renew these certificates to reduce the risk of a service outage.  A copy of this report has alerady been saved centrally at '$DOCreport'</p></br/>Brought to you by Certificate Auditor Script"
$SMTPServer = "<Enter FQDN or IP of SMTP server>"
$SMTPPort = "25"
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -port $SMTPPort -Attachments $Attachment –DeliveryNotificationOption OnSuccess

#endregion Email Report
