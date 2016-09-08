 [String] $str = [String]::Empty
                         [String] $strTemp = [String]::Empty

                         Write-Verbose "`n"
                         while (($strTemp = $reader.ReadLine()) -ne $null) {
                             # find the . character in line
                             if($strTemp -eq '.') {
                                break;
                             }

                             if ($strTemp.IndexOf('-ERR') -ne -1) {
                                break;
                             }
                             Write-Verbose $StrTemp

                             $str += $strTemp
                         }

                         Write-Verbose "`n$Str"

                         # Return raw data
                         Write-Output "`nOutput email"
                         $str -match 'From:(.+?)?<(.+?)?>.' | Out-Null