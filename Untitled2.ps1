# .LINK
#http://stackoverflow.com/questions/9959668/get-email-using-powershell

### Import the dll
[Reflection.Assembly]::LoadFile(“YourDirectory\imapx.dll”)
### Create a client object
$client = New-Object ImapX.ImapClient
###set the fetching mode to retrieve the part of message you want to retrieve, 
###the less the better
$client.Behavior.MessageFetchMode = "Full"
$client.Host = "imap.gmail.com"
$client.Port = 993
$client.UseSsl = $true
$client.Connect()
$user = "User"
$password = "Password"
$client.Login($user,$password)
$messages = $client.Folders.Inbox.Search("ALL", $client.Behavior.MessageFetchMode, 1000)
foreach($m in $messages){
$m.Subject
foreach($r in $m.Attachments){
$r | Out-File "Directory"
    }
 }
