Function Get-WebEmail {
    
# .Link
#     http://learningpcs.blogspot.com/2012/01/powershell-v2-read-gmail-more-proof-of.html
    
    [CmdletBinding()]
    Param()

    $EmailAccount = 'jeffbuenting@gmail.com'
    $EmailPassword = 'Branman1!'

     try {
        Write-Verbose "Creating new TcpClient."
        $tcpClient = New-Object -TypeName System.Net.Sockets.TcpClient

        # Connect to gmail
        $tcpClient.Connect("pop.gmail.com", 995)

        if($tcpClient.Connected) {
                Write-Verbose "You are connected to the host. Attempting to get SSL stream."

                 # Create new SSL Stream for tcpClient
                 Write-Verbose "Getting SSL stream."
                 [System.Net.Security.SslStream] $sslStream = $tcpClient.GetStream()

                 # Authenticating as client
                 Write-Verbose "Authenticating as client." 
                 $sslStream.AuthenticateAsClient("pop.gmail.com");

                if($sslStream.IsAuthenticated) {
                         Write-Verbose "You have authenticated. Attempting to login."
                         # Asssigned the writer to stream 
                         [System.IO.StreamWriter] $sw = $sslstream

                         # Assigned reader to stream
                         [System.IO.StreamReader] $reader = $sslstream

                         # refer POP rfc command, there very few around 6-9 command
                         $sw.WriteLine("USER $EmailAccount")

                         # sent to server
                         $sw.Flush()

                         # send pass
                         $sw.WriteLine("PASS $EmailPassword")
                         $sw.Flush()

                         # this will retrive your first email
                         $sw.WriteLine("RETR 1")
                         $sw.Flush()

                         $sw.WriteLine("Quit ")
                         $sw.Flush();
                            
                         [string]$strTemp = [String]::Empty

                         while (($strTemp = $reader.ReadLine() ) -ne $null ) {
                            # find the . character in line
                            if ($strTemp -eq ".") {
                                break
                            }
                            if ($strTemp.IndexOf("-ERR") -ne -1) {
                                break
                            }
                            $str += $strTemp
                        }

                        $STR
                        

                        
                    } else { 
                        Write-Error "You were not authenticated. Quitting."
               }

            } else {
                Write-Error "You are not connected to the host. Quitting"
        }
     }

     catch {
        $_
     }

     finally {
         Write-Output "Script complete."
     }

}

Get-WebEmail -Verbose