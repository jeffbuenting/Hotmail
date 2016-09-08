
http://powershellscripts.blogspot.com/2007/02/check-for-pop3-messages.html


#param (
    $Server = 'POP3.live.com'
    $Username = 'kwbrew@hotmail.com'
    $Password = '1qaz@WSX'
#)

    $UsageScript = "Usage: getMail.msh [Server] [Username] [Password]"
    $Port = 995



    $TCPConnection = new-object -TypeName System.Net.Sockets.TcpClient($Server, $Port)

    $NetStream = $TCPConnection.GetStream()
    $Reader = new-object -TypeName System.IO.StreamReader($NetStream)
    $Writer = new-object -TypeName System.IO.StreamWriter($NetStream)
    $Buffer = $Reader.ReadLine()
    $Writer.WriteLine("USER $Username")
    $Writer.Flush(); $Buffer = $Reader.ReadLine()
    $Writer.WriteLine("PASS $Password")
    $Writer.Flush();

    if ($Reader.ReadLine() -match "OK") {
            $writer.WriteLine("STAT"); $writer.Flush()
            $NumOfMessage = $Reader.ReadLine().SubString(4, 1)
            write-output "You Have $NumOfMessage Item(s) on $Server"
        }
        else {
             write-output "Authentication Error"
    }

    $Reader.Dispose()
    $Writer.Dispose()
    $NetStream.Dispose()
    $TCPConnection.Close() 
