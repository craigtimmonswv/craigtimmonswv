﻿cls
# Prompts user for where to store the file
$filelocation = Read-Host "Enter Location/filename to store output (i.e c:\scripts\test.csv)"
# Tests to see if the file currently exists.  It will stop if it does. 
if (Test-Path -Path $filelocation -PathType leaf)
    {
        Write-Host "File Exists Stopping"
    }
else 
    {
        # creates an Array called $details.  This will store all the information. 
        $erlocations = Get-CsTenantNetworkSite
        $details = @()
        foreach ($location in $erlocations)
        {
        
            $networks = Get-CsTenantNetworkSubnet | ? {$_.networksiteid -eq $location.NetworkSiteID}
            foreach ($net in $networks)

                {
                    $detail = New-Object PSObject
                    $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $location.Identity
                    $detail | add-Member -MemberType NoteProperty -Name "NetworkSiteID" -Value $net.NetworkSiteID
                    $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $net.Description
                    $detail | add-Member -MemberType NoteProperty -Name "SubnetID" -Value $net.SubnetID
                    $detail | add-Member -MemberType NoteProperty -Name "MaskBits" -Value $net.MaskBits
                    $detail | add-Member -MemberType NoteProperty -Name "EmergencyCallRoutingPolicy" -Value $location.EmergencyCallRoutingPolicy
                    $detail | add-Member -MemberType NoteProperty -Name "EmergencyCallingPolicy" -Value $location.EmergencyCallingPolicy
                    $details += $detail  
                }

        }
    }

    $details |Export-Csv $filelocation -Append -NoTypeInformation

# SIG # Begin signature block
# MIIVpgYJKoZIhvcNAQcCoIIVlzCCFZMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUuDjBwagIxnrehzj/iXq20sF7
# Md2gghIHMIIFbzCCBFegAwIBAgIQSPyTtGBVlI02p8mKidaUFjANBgkqhkiG9w0B
# AQwFADB7MQswCQYDVQQGEwJHQjEbMBkGA1UECAwSR3JlYXRlciBNYW5jaGVzdGVy
# MRAwDgYDVQQHDAdTYWxmb3JkMRowGAYDVQQKDBFDb21vZG8gQ0EgTGltaXRlZDEh
# MB8GA1UEAwwYQUFBIENlcnRpZmljYXRlIFNlcnZpY2VzMB4XDTIxMDUyNTAwMDAw
# MFoXDTI4MTIzMTIzNTk1OVowVjELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3Rp
# Z28gTGltaXRlZDEtMCsGA1UEAxMkU2VjdGlnbyBQdWJsaWMgQ29kZSBTaWduaW5n
# IFJvb3QgUjQ2MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAjeeUEiIE
# JHQu/xYjApKKtq42haxH1CORKz7cfeIxoFFvrISR41KKteKW3tCHYySJiv/vEpM7
# fbu2ir29BX8nm2tl06UMabG8STma8W1uquSggyfamg0rUOlLW7O4ZDakfko9qXGr
# YbNzszwLDO/bM1flvjQ345cbXf0fEj2CA3bm+z9m0pQxafptszSswXp43JJQ8mTH
# qi0Eq8Nq6uAvp6fcbtfo/9ohq0C/ue4NnsbZnpnvxt4fqQx2sycgoda6/YDnAdLv
# 64IplXCN/7sVz/7RDzaiLk8ykHRGa0c1E3cFM09jLrgt4b9lpwRrGNhx+swI8m2J
# mRCxrds+LOSqGLDGBwF1Z95t6WNjHjZ/aYm+qkU+blpfj6Fby50whjDoA7NAxg0P
# OM1nqFOI+rgwZfpvx+cdsYN0aT6sxGg7seZnM5q2COCABUhA7vaCZEao9XOwBpXy
# bGWfv1VbHJxXGsd4RnxwqpQbghesh+m2yQ6BHEDWFhcp/FycGCvqRfXvvdVnTyhe
# Be6QTHrnxvTQ/PrNPjJGEyA2igTqt6oHRpwNkzoJZplYXCmjuQymMDg80EY2NXyc
# uu7D1fkKdvp+BRtAypI16dV60bV/AK6pkKrFfwGcELEW/MxuGNxvYv6mUKe4e7id
# FT/+IAx1yCJaE5UZkADpGtXChvHjjuxf9OUCAwEAAaOCARIwggEOMB8GA1UdIwQY
# MBaAFKARCiM+lvEH7OKvKe+CpX/QMKS0MB0GA1UdDgQWBBQy65Ka/zWWSC8oQEJw
# IDaRXBeF5jAOBgNVHQ8BAf8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zATBgNVHSUE
# DDAKBggrBgEFBQcDAzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEMGA1Ud
# HwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwuY29tb2RvY2EuY29tL0FBQUNlcnRpZmlj
# YXRlU2VydmljZXMuY3JsMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuY29tb2RvY2EuY29tMA0GCSqGSIb3DQEBDAUAA4IBAQASv6Hvi3Sa
# mES4aUa1qyQKDKSKZ7g6gb9Fin1SB6iNH04hhTmja14tIIa/ELiueTtTzbT72ES+
# BtlcY2fUQBaHRIZyKtYyFfUSg8L54V0RQGf2QidyxSPiAjgaTCDi2wH3zUZPJqJ8
# ZsBRNraJAlTH/Fj7bADu/pimLpWhDFMpH2/YGaZPnvesCepdgsaLr4CnvYFIUoQx
# 2jLsFeSmTD1sOXPUC4U5IOCFGmjhp0g4qdE2JXfBjRkWxYhMZn0vY86Y6GnfrDyo
# XZ3JHFuu2PMvdM+4fvbXg50RlmKarkUT2n/cR/vfw1Kf5gZV6Z2M8jpiUbzsJA8p
# 1FiAhORFe1rYMIIGGjCCBAKgAwIBAgIQYh1tDFIBnjuQeRUgiSEcCjANBgkqhkiG
# 9w0BAQwFADBWMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVk
# MS0wKwYDVQQDEyRTZWN0aWdvIFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYw
# HhcNMjEwMzIyMDAwMDAwWhcNMzYwMzIxMjM1OTU5WjBUMQswCQYDVQQGEwJHQjEY
# MBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJTZWN0aWdvIFB1Ymxp
# YyBDb2RlIFNpZ25pbmcgQ0EgUjM2MIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIB
# igKCAYEAmyudU/o1P45gBkNqwM/1f/bIU1MYyM7TbH78WAeVF3llMwsRHgBGRmxD
# eEDIArCS2VCoVk4Y/8j6stIkmYV5Gej4NgNjVQ4BYoDjGMwdjioXan1hlaGFt4Wk
# 9vT0k2oWJMJjL9G//N523hAm4jF4UjrW2pvv9+hdPX8tbbAfI3v0VdJiJPFy/7Xw
# iunD7mBxNtecM6ytIdUlh08T2z7mJEXZD9OWcJkZk5wDuf2q52PN43jc4T9OkoXZ
# 0arWZVeffvMr/iiIROSCzKoDmWABDRzV/UiQ5vqsaeFaqQdzFf4ed8peNWh1OaZX
# nYvZQgWx/SXiJDRSAolRzZEZquE6cbcH747FHncs/Kzcn0Ccv2jrOW+LPmnOyB+t
# AfiWu01TPhCr9VrkxsHC5qFNxaThTG5j4/Kc+ODD2dX/fmBECELcvzUHf9shoFvr
# n35XGf2RPaNTO2uSZ6n9otv7jElspkfK9qEATHZcodp+R4q2OIypxR//YEb3fkDn
# 3UayWW9bAgMBAAGjggFkMIIBYDAfBgNVHSMEGDAWgBQy65Ka/zWWSC8oQEJwIDaR
# XBeF5jAdBgNVHQ4EFgQUDyrLIIcouOxvSK4rVKYpqhekzQwwDgYDVR0PAQH/BAQD
# AgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwGwYD
# VR0gBBQwEjAGBgRVHSAAMAgGBmeBDAEEATBLBgNVHR8ERDBCMECgPqA8hjpodHRw
# Oi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ1Jvb3RS
# NDYuY3JsMHsGCCsGAQUFBwEBBG8wbTBGBggrBgEFBQcwAoY6aHR0cDovL2NydC5z
# ZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdSb290UjQ2LnA3YzAj
# BggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20wDQYJKoZIhvcNAQEM
# BQADggIBAAb/guF3YzZue6EVIJsT/wT+mHVEYcNWlXHRkT+FoetAQLHI1uBy/YXK
# ZDk8+Y1LoNqHrp22AKMGxQtgCivnDHFyAQ9GXTmlk7MjcgQbDCx6mn7yIawsppWk
# vfPkKaAQsiqaT9DnMWBHVNIabGqgQSGTrQWo43MOfsPynhbz2Hyxf5XWKZpRvr3d
# MapandPfYgoZ8iDL2OR3sYztgJrbG6VZ9DoTXFm1g0Rf97Aaen1l4c+w3DC+IkwF
# kvjFV3jS49ZSc4lShKK6BrPTJYs4NG1DGzmpToTnwoqZ8fAmi2XlZnuchC4NPSZa
# PATHvNIzt+z1PHo35D/f7j2pO1S8BCysQDHCbM5Mnomnq5aYcKCsdbh0czchOm8b
# kinLrYrKpii+Tk7pwL7TjRKLXkomm5D1Umds++pip8wH2cQpf93at3VDcOK4N7Ew
# oIJB0kak6pSzEu4I64U6gZs7tS/dGNSljf2OSSnRr7KWzq03zl8l75jy+hOds9TW
# SenLbjBQUGR96cFr6lEUfAIEHVC1L68Y1GGxx4/eRI82ut83axHMViw1+sVpbPxg
# 51Tbnio1lB93079WPFnYaOvfGAA0e0zcfF/M9gXr+korwQTh2Prqooq2bYNMvUoU
# KD85gnJ+t0smrWrb8dee2CvYZXD5laGtaAxOfy/VKNmwuWuAh9kcMIIGcjCCBNqg
# AwIBAgIQXD41nnmZYnF2ThRsECu1mzANBgkqhkiG9w0BAQwFADBUMQswCQYDVQQG
# EwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJTZWN0aWdv
# IFB1YmxpYyBDb2RlIFNpZ25pbmcgQ0EgUjM2MB4XDTIyMDEwMzAwMDAwMFoXDTIz
# MDEwMzIzNTk1OVowVTELMAkGA1UEBhMCVVMxFjAUBgNVBAgMDVdlc3QgVmlyZ2lu
# aWExFjAUBgNVBAoMDUNyYWlnIFRpbW1vbnMxFjAUBgNVBAMMDUNyYWlnIFRpbW1v
# bnMwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCjjf7cVaOFnCw6/rdM
# p8XO7AlUq3mdX7Qj+9zYxetxT9r8fA+MlRcrztF12YY+VNAtMl2VsFk/t9rjbR0D
# 1VRpV+gqqpJ7a5EtrvYvOkqpLjlH6YuoXnsGCzMubgUjwyK1EPa4WYgyZTD/hIEW
# r3YtHNggAkMYpyxQxaamN0l2cGtH6IVZHBiAci8fYqcqoetyRTspZNeLRm5AkZBp
# 2m8frr5Ma/qsMI5azkGR4rb1NuvgohopXBeDeSDZMUWGkqANHJaI8THefoc/VvSB
# eU3cA5Na8LKiVIXldxbWIu/xoskWDDZbyLOtI4PohTAqo3/+AsO4ywsWauzwmr8j
# LnT8gB2I+w5VYrmGaFeeQTvOk0sN15gDL7CyFU3xA28jDwt4acJunbTr+mEI+LXy
# +cEqSkcmzF1ziHRLtkghjLOGsL/8VRLehIQj8QOzI4Ko+JEkFpNoQ4jTKFJOPPmS
# mEfVqRNwbP+jjUMLDPGu3YSH1R9hhD8E0UX89iFf9ySyHl8nNnRzRB0P0KakPk6l
# iJtme82KGAIBq471WSaC5NjjvnXTGzw2w3YSnFuzOq6KI1nE29hAWPQp359UqusE
# WH991EO5+FomUYbz/orGgrdMhKbs46CbTiWr3o1XRCB0x4MueeBWK/w8MdjE1l2z
# CkDNW6R6wVuazFYq8M/C+7FEAQIDAQABo4IBvTCCAbkwHwYDVR0jBBgwFoAUDyrL
# IIcouOxvSK4rVKYpqhekzQwwHQYDVR0OBBYEFALv9uiU65/zQs+lX7CUOVU1X3ai
# MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoGCCsGAQUF
# BwMDMBEGCWCGSAGG+EIBAQQEAwIEEDBKBgNVHSAEQzBBMDUGDCsGAQQBsjEBAgED
# AjAlMCMGCCsGAQUFBwIBFhdodHRwczovL3NlY3RpZ28uY29tL0NQUzAIBgZngQwB
# BAEwSQYDVR0fBEIwQDA+oDygOoY4aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0
# aWdvUHVibGljQ29kZVNpZ25pbmdDQVIzNi5jcmwweQYIKwYBBQUHAQEEbTBrMEQG
# CCsGAQUFBzAChjhodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWND
# b2RlU2lnbmluZ0NBUjM2LmNydDAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wHwYDVR0RBBgwFoEUY3JhaWdAdGhldGltbW9ucy5uZXQwDQYJKoZI
# hvcNAQEMBQADggGBAF9RTBcs4Sp7HUnX/Ry1iV96fYzTlMLt28RBxYPkruBwc3Lu
# ZO7UavdCvgZRs/OZ8sesr18jh30PQnwkqxRe8jQbBV9NoPmMwDm5g6rQmLE7or1f
# Yrn475sJQHSwK1jQCtrsbDfWGgfqqjmkRT8MAI+l2zvAC3LcSx70QCuq5yvUuwYt
# MnxzUYVVPBWJ913KJLthb5wxWRzoYpVpoZw33sZAXsJIi6Tvbd9hu2/7k/+xF1FL
# VqCnhIhcinz7I9eIMIW74DAlkjHdIimbLEDbCdHGcAzaB/1pnZ7omiPRWM7wVCHe
# Wc2MYfZoJQfKpaC87TQRaPW5+dG6Cu/nwZ4nF0IJ4LNCmaRC9XQvGjvIgpPd3c3q
# JnlZWYrbwF8FfVZVfTsYgUFUvGjuOSgzKXCn1xj9uc5Xxf8n4ONO6W094BKEEQJ1
# iIhuvuwCzSSEExA5nwjCEwUKzD4KhIIDmwYvaMEPP+vUBNxEXXJBhqPOhL8gKH/y
# 3fTIbUJsLk28MbBaZzGCAwkwggMFAgEBMGgwVDELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDErMCkGA1UEAxMiU2VjdGlnbyBQdWJsaWMgQ29k
# ZSBTaWduaW5nIENBIFIzNgIQXD41nnmZYnF2ThRsECu1mzAJBgUrDgMCGgUAoHgw
# GAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQx
# FgQUG2Ljjsk9XOo/sDdf3pLDity5/+kwDQYJKoZIhvcNAQEBBQAEggIAOBcrNLnZ
# VVxerxswgbjVefk+0A71a92IhwFQsPqUMoCun74j4Q5FhhhSjmzH9/81WRKVMAz5
# 9eXnYyOVmKUW1tuLPuLiRJ5sHgQkZJ/m121n71CHWfasG+Unal/WK9b8NMLRblkJ
# /3/Heh19CNlnA2Xnbf7Nv6whTzb79u/k5YrK1XyetcNljj8/Up+Vfoxr6xQdaoce
# fn25d6YLv2y+LGdVPH+6jQz0dujv6gAzDDIibJDH7QEcOGw7u0MkzuawbCkXpBXu
# p9hvy9iUNu3qScduG7TaddT//Kmkit2wTAwGpcsXtrd9LDpfFbajazird64wMj4N
# UcvBFV0OkR1pgaM++5piHYLmR4xm/zPkxToxleGaBrBxlSGHINZr2FoV4pEa52QY
# Xcex8UGsW03fmBJF77CNBrcGEPp9gtFJFJ/Z3zlZQbH11e+eLdO2uo5zCCkLnFM0
# 8aZ/jfqP2TjUpSpW6kOFFXr/483xG5FjyqSLU0suKB7Pn+aDYh7ZTLz92o4rRboE
# uSYflkjcBE0/xal93RKG4C5RN2EasdcoNDGzkqsaM6Z84CfFTNXT9+p/YI89TjXT
# rqO7KS7YeqXtK+tMEn7kVRamX+48nPCcOP6vevXSuA1y0mgiTjkEr+W+FPwd/1DI
# gFL1oaS/UM85LyrfGNLiCgqvE3UO/jKNrAI=
# SIG # End signature block
