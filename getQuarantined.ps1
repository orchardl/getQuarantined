# Get today's date
$today = Get-Date

# Get yesterday's date
$yesterday = $today.AddDays(-1)

# Get tomorrow's date
$tomorrow = $today.AddDays(1)

# Convert today's date into the format you specified for the filename
$todayString = $today.ToString('yyyyMMdd')

# Build the filename
$filename = "C:\Users\username\QuarantinedEmailsbyDateRange" + $todayString + ".csv"

# Connect Exchange Online
Connect-ExchangeOnline -UserPrincipalName your_o365_email@example.com

# Execute your command
Get-QuarantineMessage -StartReceivedDate $yesterday -EndReceivedDate $tomorrow | Select ReceivedTime,SenderAddress,RecipientAddress,Subject,MessageID,RecipientCount,QuarantineTypes | Export-Csv -Path $filename -NoTypeInformation -Append -Force

# Send email with Quarantined Messages
Send-MailMessage -From 'Quarantine <yomamma@example.com>' -SmtpServer 'smtp.example.com' -To 'Your Emails <your_email@example.com>' -Subject 'Test mail' -Attachments $filename