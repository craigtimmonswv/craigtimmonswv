Function Write-DataToExcel
    {
        param ($filelocation, $details, $tabname)

        
        $excelpackage = Open-ExcelPackage -Path $filelocation 
        $ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName $tabname
        $details | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter
        Clv details 

    }


cls
Write-Host "This is will create an Excel Spreadsheet.  Make sure to enter the file name with .xlsx"
Write-Host "You will need to verify that you have installed the importexcel module"
$Details = @()



$filelocation = Read-Host "Enter Location/filename to store output (i.e c:\scripts\test.xlsx)"
CLS

# Extract PSTN Gateways
Write-Host 'Gathering Online PSTN Gateway Details'
$PSTNGWs = Get-CsOnlinePSTNGateway

foreach ($GW in $PSTNGWs)
    {       
        $detail = New-Object PSObject
        $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $GW.Identity
        $detail | add-Member -MemberType NoteProperty -Name "Fqdn" -Value $GW.Fqdn
        $detail | Add-Member -MemberType NoteProperty -Name "NumberPattern" -Value $GW.SipSignalingPort
        $detail | Add-Member -MemberType NoteProperty -Name "FailoverTimeSeconds" -Value $GW.FailoverTimeSeconds
        $detail | Add-Member -MemberType NoteProperty -Name "ForwardCallHistory" -Value $GW.ForwardCallHistory
        $detail | Add-Member -MemberType NoteProperty -Name "ForwardPai" -Value $GW.ForwardPai
        $detail | Add-Member -MemberType NoteProperty -Name "SendSipOptions" -Value $GW.SendSipOptions
        $detail | Add-Member -MemberType NoteProperty -Name "MaxConcurrentSessions" -Value $GW.MaxConcurrentSessions
        $detail | Add-Member -MemberType NoteProperty -Name "Enabled" -Value $GW.Enabled
        $detail | Add-Member -MemberType NoteProperty -Name "BypassMode" -Value $GW.BypassMode
        $detail | Add-Member -MemberType NoteProperty -Name "MediaBypass" -Value $GW.MediaBypass
        $detail | Add-Member -MemberType NoteProperty -Name "GatewaySiteId" -Value $GW.GatewaySiteId
        $detail | Add-Member -MemberType NoteProperty -Name "PidfLoSupported" -Value $GW.PidfLoSupported
        $detail | Add-Member -MemberType NoteProperty -Name "ProxySbc" -Value $GW.ProxySbc
        $detail | Add-Member -MemberType NoteProperty -Name "GatewaySiteLbrEnabled" -Value $GW.GatewaySiteLbrEnabled
        $detail | Add-Member -MemberType NoteProperty -Name "FailoverResponseCodes" -Value $GW.FailoverResponseCodes
        $Details += $detail
    }


$Details|Export-Excel -Path $filelocation -WorksheetName "PSTN Gateways" -AutoSize -AutoFilter
clv details

Write-Host 'Getting PSTN Usages'
$PSTNUSAGEs = Get-CsOnlinePstnUsage
$details =@()
foreach ($PSTNUsage in $PSTNUSAGEs)
    {   
        foreach ($u in $PSTNUSAGE.Usage)
        {
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $PSTNUSAGE.Identity
            $detail | add-Member -MemberType NoteProperty -Name "Usage" -Value $u
            $details += $detail
        }
    }
$tabname = 'PSTN Usages'
Write-DataToExcel $filelocation  $details $tabname



Write-Host 'Getting Voice Routes'
$Details = @()
$VRs = Get-CsOnlineVoiceRoute
foreach ($VR in $VRs)
    {   
            
        [string] $usage= $vr.OnlinePstnUsages
        [string] $pstngw =$vr.OnlinePstnGatewayList 
        $detail = New-Object PSObject
        $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $vr.Name
        $detail | Add-Member -MemberType NoteProperty -Name "NumberPattern" -Value $vr.NumberPattern
        $detail | Add-Member -MemberType NoteProperty -Name "OnlinePstnUsages" -Value $usage
        $detail | Add-Member -MemberType NoteProperty -Name "OnlinePstnGatewayList " -Value $pstngw
        $details += $detail
    }

$tabname = 'Voice Routes'
Write-DataToExcel $filelocation  $details $tabname

#Extracts Voice Routing Policies
Write-Host 'Getting Voice Routing Policies'
$Details = @()
$VRPs = Get-CsOnlineVoiceRoutingPolicy
foreach ($VRP in $VRPs)
    {       
        foreach ($usage in $vrp.OnlinePstnUsages)
        {
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $VRP.Identity
            $detail | Add-Member -MemberType NoteProperty -Name "Description" -Value $VRP.Description
            $detail | Add-Member -MemberType NoteProperty -Name "OnlinePstnUsages" -Value $usage
                    
            $Details += $detail
        }
    }

Write-DataToExcel $filelocation  $details $tabname


# Extracts Dial Plan info
Write-Host 'Getting Dial Plan Details'
$DPs=Get-CsTenantDialPlan
$Details = @()
foreach ($dp in $DPs)
    {   
        foreach ($rule in $dp.NormalizationRules)
            {
                # Creating an array to store the variables from the dial plans. 
                $detail = New-Object PSObject
                $detail | add-Member -MemberType NoteProperty -Name "Parent" -Value $dp.Identity.remove(0,4)
                $detail | Add-Member -MemberType NoteProperty -Name "Description" -Value $rule.Description
                $detail | Add-Member -MemberType NoteProperty -Name "Name" -Value $rule.Name
                $detail | Add-Member -MemberType NoteProperty -Name "Pattern" -Value $rule.Pattern
                $detail | Add-Member -MemberType NoteProperty -Name "Translation" -Value $rule.Translation
                $detail | Add-Member -MemberType NoteProperty -Name "IsInternalExtension" -Value $rule.IsInternalExtension
                    
                # Adding array from one dial plan to an array with all the dial plans. 
                $Details += $detail
                }
        }
$tabname = "Dial Plan"
Write-DataToExcel $filelocation  $details $tabname



# Extracts users enablement
Write-Host 'Getting Voice Enabled Users'
$Details = @()
$users =  Get-CsOnlineUser | ?  {$_.enterprisevoiceenabled -eq $true}
        $Userdetails = @()
        foreach ($user in $users)
            {
                # Creating an array to store the variables from the dial plans. 
                $detail = New-Object PSObject
                $detail | add-Member -MemberType NoteProperty -Name "Displayname" -Value $user.displayname
                $detail | add-Member -MemberType NoteProperty -Name "City" -Value $user.City
                $detail | add-Member -MemberType NoteProperty -Name "UPN" -Value $user.UserPrincipalName
                $detail | add-Member -MemberType NoteProperty -Name "Lineuri" -Value $user.LineUri
                $detail | add-Member -MemberType NoteProperty -Name "Dial Plan" -Value $user.TenantDialPlan
                $detail | add-Member -MemberType NoteProperty -Name "Voice Routing Policy" -Value $user.OnlineVoiceRoutingPolicy
                $detail | add-Member -MemberType NoteProperty -Name "EV Enabled" -Value $user.EnterpriseVoiceEnabled
                $detail | add-Member -MemberType NoteProperty -Name "Teams Upgrade Policy" -Value $user.TeamsUpgradePolicy
                $detail | add-Member -MemberType NoteProperty -Name "Teams Effective Mode" -Value $user.TeamsUpgradeEffectiveMode
                $details += $detail

            }
$tabname = "EV Users"
Write-DataToExcel $filelocation  $details $tabname



# Extract Emergency Calling Policies
Write-Host 'Getting Emergency Calling Policies'
$Details = @()
$ercallpolicies = Get-CsTeamsEmergencyCallingPolicy
    foreach ($ercp in $ercallpolicies)
        {
            $detail = New-Object PSObject
            $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $ercp.Identity
            $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $ercp.Description
            $detail | add-Member -MemberType NoteProperty -Name "NotificationGroup" -Value $ercp.NotificationGroup
            $detail | add-Member -MemberType NoteProperty -Name "ExternalLocationLookupMode" -Value $ercp.ExternalLocationLookupMode
            $detail | add-Member -MemberType NoteProperty -Name "NotificationDialOutNumber" -Value $ercp.NotificationDialOutNumber
            $detail | add-Member -MemberType NoteProperty -Name "NotificationMode" -Value $ercp.NotificationMode
            $details += $detail  
        }
$tabname = "Emergency Calling Policies"
Write-DataToExcel $filelocation  $details $tabname

# Extracts Emergency Call Routing Policy
Write-Host 'Getting Emergency Call Routing Policies'
$Details = @()
$ecrps = Get-CsTeamsEmergencyCallRoutingPolicy
foreach ($ecrp in $ecrps)
    {
        $numbers = Get-CsTeamsEmergencyCallRoutingPolicy -Identity $ecrp.identity
        foreach ($number in $numbers.EmergencyNumbers)
            {
                $detail = New-Object PSObject
                $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $ecrp.Identity
                $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $ecrp.Description
                $detail | add-Member -MemberType NoteProperty -Name "emergencydialstring" -Value $number.emergencydialstring
                $detail | add-Member -MemberType NoteProperty -Name "EmergencyDialMask" -Value $number.emergencydialmask
                $detail | add-Member -MemberType NoteProperty -Name "OnlinePSTNUsage" -Value $number.OnlinePSTNUsage
                $detail | add-Member -MemberType NoteProperty -Name "AllowEnhancedEmergencyServices" -Value $ecrp.AllowEnhancedEmergencyServices
                $details  += $detail  
            }
    }
$tabname = "Emergency Cal Routing Policies"
Write-DataToExcel $filelocation  $details $tabname

Write-Host 'Getting Tenant Network Site Details'
$Details = @()
$erlocations = Get-CsTenantNetworkSite
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
$tabname = "Tenant Network Site Details"
Write-DataToExcel $filelocation  $details $tabname

# Emergency Location information Services 
Write-Host 'Getting Emergency Location Information Services'
$locations = Get-CsOnlineLisLocation
$Details = @()
Foreach ($loc in $locations)
    {
        $detail = New-Object PSObject
        $detail | Add-Member NoteProperty -Name "CompanyName" -Value $loc.CompanyName
        $detail | Add-Member NoteProperty -Name "Civicaddressid" -Value $loc.civicaddressid
        $detail | Add-Member NoteProperty -Name "locationid" -Value $loc.LocationId
        $detail | Add-Member NoteProperty -Name "Description" -Value $loc.Description
        $detail | Add-Member NoteProperty -Name "location" -Value $loc.location
        $detail | Add-Member NoteProperty -Name "HouseNumber" -Value $loc.HouseNumber
        $detail | Add-Member NoteProperty -Name "HouseNumberSuffix" -Value $loc.HouseNumberSuffix
        $detail | Add-Member NoteProperty -Name "PreDirectional" -Value $loc.PreDirectional
        $detail | Add-Member NoteProperty -Name "StreetName" -Value $loc.StreetName
        $detail | Add-Member NoteProperty -Name "PostDirectional" -Value $loc.PostDirectional
        $detail | Add-Member NoteProperty -Name "StreetSuffix" -Value $loc.StreetSuffix
        $detail | Add-Member NoteProperty -Name "City" -Value $loc.City
        $detail | Add-Member NoteProperty -Name "StateOrProvince" -Value $loc.StateOrProvince
        $detail | Add-Member NoteProperty -Name "PostalCode" -Value $loc.PostalCode
        $detail | Add-Member NoteProperty -Name "Country" -Value $loc.CountryOrRegion
        $detail | Add-Member NoteProperty -Name "Latitude" -Value $loc.Latitude
        $detail | Add-Member NoteProperty -Name "Longitude" -Value $loc.Longitude
        $Details += $detail
        }
$tabname = "LIS Location"
Write-DataToExcel $filelocation  $details $tabname


#LIS Network information
Write-Host 'Getting LIS Network Information'
$subnets = Get-CsOnlineLisSubnet
$Details = @()
Foreach ($subnet in $subnets)
    {
        $detail = New-Object PSObject
        $detail | Add-Member NoteProperty -Name "Subnet" -Value $subnet.Subnet
        $detail | Add-Member NoteProperty -Name "Description" -Value $subnet.Description
        $subloc = Get-CsOnlineLisLocation -LocationId $subnet.LocationId
        $detail | Add-Member NoteProperty -Name "Location" -Value $subloc.location
        $detail | Add-Member NoteProperty -Name "City" -Value $subloc.city
        $Details += $detail
    }
$tabname = "LIS Network "
Write-DataToExcel $filelocation  $details $tabname


#LIS Wireless Access Point information
Write-Host 'Getting LIS WAP Information'
$WAPs = Get-CsOnlineLisWirelessAccessPoint
$Details = @()
Foreach ($WAP in $WAPs)
    {
        $detail = New-Object PSObject
        $detail | Add-Member NoteProperty -Name "BSSID" -Value $WAP.BSSID
        $detail | Add-Member NoteProperty -Name "Description" -Value $WAP.Description
        $WAPloc = Get-CsOnlineLisLocation -LocationId $WAP.LocationId
        $detail | Add-Member NoteProperty -Name "Location" -Value $WAPloc.location
        $detail | Add-Member NoteProperty -Name "City" -Value $WAPloc.city
        $Details += $detail
    }
$tabname = "LIS WAP"
Write-DataToExcel $filelocation  $details $tabname


#LIS Switch information
Write-Host 'Getting LIS SWitch information'
$Switches = Get-CsOnlineLisSwitch
$Details = @()
Foreach ($Switch in $Switches)
    {
        $detail = New-Object PSObject
        $detail | Add-Member NoteProperty -Name "ChassisID" -Value $Switch.ChassisID
        $detail | Add-Member NoteProperty -Name "Description" -Value $Switch.Description
        $Switchloc = Get-CsOnlineLisLocation -LocationId $Switch.LocationId
        $detail | Add-Member NoteProperty -Name "Location" -Value $Switchloc.location
        $detail | Add-Member NoteProperty -Name "City" -Value $Switchloc.city
        $Details += $detail
    }
$tabname = "LIS Switch"
Write-DataToExcel $filelocation  $details $tabname


#LIS Port information
Write-Host 'Getting LIS Port Information'
$Ports = Get-CsOnlineLisPort
$Details = @()
Foreach ($port in $ports)
    {
        $detail = New-Object PSObject
        $detail | Add-Member NoteProperty -Name "ChassisID" -Value $port.ChassisID
        $detail | Add-Member NoteProperty -Name "PortID" -Value $port.PortID
        $detail | Add-Member NoteProperty -Name "Description" -Value $port.Description
        $portloc = Get-CsOnlineLisLocation -LocationId $port.LocationId
        $detail | Add-Member NoteProperty -Name "Location" -Value $portloc.location
        $detail | Add-Member NoteProperty -Name "City" -Value $portloc.city
        $Details += $detail
    }
$tabname = "LIS Port"
Write-DataToExcel $filelocation  $details $tabname

foreach ($aa in $AAs)
    {
        foreach ($RA in $aa.ApplicationInstances)
            {
                $ResouceAct =Get-CsOnlineApplicationInstance -Identity $ra
                $operator = get-csonlineuser -identity $aa.Operator.Id
                $detail = New-Object PSObject
                $detail | Add-Member NoteProperty -Name "AAName" -Value $aa.name
                $detail | Add-Member NoteProperty -Name "Identity" -Value $aa.identity
                $detail | Add-Member NoteProperty -Name "Operator" -Value $operator.userprincipalname
                $detail | Add-Member NoteProperty -Name "Language" -Value $aa.LanguageId
                $detail | Add-Member NoteProperty -Name "TimeZone" -Value $aa.timezoneid
                $detail | Add-Member NoteProperty -Name "VoiceResponseEnabled" -Value $aa.VoiceresponseEnabled
                $cflows = @()
                foreach ($CF in $AA)
                    {
                        [string]$callflows = @((Get-CsAutoAttendant -Identity $aa.Identity | select callflows).callflows.name -join "," )
                        $cflows += $callflows
                    }
                $detail | Add-Member NoteProperty -Name "CallFlows" -Value $cflows
                $detail | Add-Member NoteProperty -Name "ResourceAccount" -Value $ResouceAct.UserPrincipalName
                $detail | Add-Member NoteProperty -Name "Phone Number" -Value $ResouceAct.PhoneNumber
                $details += $detail
                clv detail
            }
    }
$tabname = "Auto Attendant"
Write-DataToExcel $filelocation  $details $tabname

# Call Queues
$Details = @()
$CQs = Get-CsCallQueue 
foreach ($CQ in $CQs)
    { 
        #$ResouceAct =Get-CsOnlineApplicationInstance -Identity $ra
        $detail = New-Object PSObject
        $detail | Add-Member NoteProperty -Name "AAName" -Value $CQ.name
        $detail | Add-Member NoteProperty -Name "Identity" -Value $CQ.identity
        $detail | Add-Member NoteProperty -Name "RoutingMethod" -Value $CQ.RoutingMethod
        $agents=@()
        foreach ($agt in $CQ.Agents)
            {
                [string]$agtname = @((Get-CsOnlineUser -Identity $agt.ObjectId | select UserPrincipalName).UserPrincipalName  -join "," )
                $agents += $agtname
            }
                
        $detail | Add-Member NoteProperty -Name "Agents" -Value $Agents
        $detail | Add-Member NoteProperty -Name "ConferenceMode" -Value $CQ.ConferenceMode
        $detail | Add-Member NoteProperty -Name "PresenceBasedRouting" -Value $CQ.PresenceBasedRouting
        $detail | Add-Member NoteProperty -Name "AgentAlertTime" -Value $CQ.AgentAlertTime
        $detail | Add-Member NoteProperty -Name "OverflowThreshold" -Value $CQ.OverflowThreshold
        $detail | Add-Member NoteProperty -Name "OverflowAction" -Value $CQ.OverflowAction
        $ofatarget = @()
        foreach ($ofat in $CQ.OverflowActionTarget)
            {
                [string]$overflowtarget = @(((Get-CsCallQueue -Identity 90bee1bd-fb44-4d00-be47-eed5f6cc4a9b | select OverflowActionTarget).OverflowActionTarget).type  -join "," )
                $ofatarget += $overflowtarget
            }

        $detail | Add-Member NoteProperty -Name "OverflowActionTarget" -Value $ofatarget
        $detail | Add-Member NoteProperty -Name "OverflowSharedVoicemailTextToSpeechPrompt" -Value $CQ.OverflowSharedVoicemailTextToSpeechPrompt
        $detail | Add-Member NoteProperty -Name "TimeoutThreshold" -Value $CQ.TimeoutThreshold
        $detail | Add-Member NoteProperty -Name "TimeoutAction" -Value $CQ.TimeoutAction
        $detail | Add-Member NoteProperty -Name "TimeoutActionTarget" -Value $CQ.TimeoutActionTarget
        $detail | Add-Member NoteProperty -Name "TimeoutSharedVoicemailTextToSpeechPrompt" -Value $CQ.TimeoutSharedVoicemailTextToSpeechPrompt
        $detail | Add-Member NoteProperty -Name "EnableTimeoutSharedVoicemailTranscription" -Value $CQ.EnableTimeoutSharedVoicemailTranscription
        $details += $detail
    }
$tabname = "Call Queue"
Write-DataToExcel $filelocation  $details $tabname

# SIG # Begin signature block
# MIIVpgYJKoZIhvcNAQcCoIIVlzCCFZMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUKQ+YYiGINfXd8uLxn5UTVOKd
# LW2gghIHMIIFbzCCBFegAwIBAgIQSPyTtGBVlI02p8mKidaUFjANBgkqhkiG9w0B
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
# FgQUpdUpt4z550pnZANDR2/qUDxn03UwDQYJKoZIhvcNAQEBBQAEggIAEokKfRWg
# R1CW5G6ol3K4B8dcJvtqcKhBqOCK6/o8kHEqKihqJKylX/PyhMNTecvkb3juF5JS
# 8NHxYXVGaVLDrPEltLrTpAgq0Rafi+kn/v34XJthC1unFw8smw/gW6VrFDYR1AMk
# NaUf6lL5p1wodWs8WU5O5TvPCvN7BrT5Ln4qVdy0uSrGLYIwF4m6FaYezf8XiJ17
# Qs6z2IrPskDemjHUuKfzQ7+VZoJBgImdvdjgOyuY0dJnWgiLiWt4pFklYK6MTsKs
# 6lOueYxxNnAWallnMTeYQWZJ797UMnhFz2j49LSHtIpc6gwYEtzTNKSJAzKcYRuG
# OpwGmSy9kenPtvJR7DwONRPPh/pImDRRp36u0EuJla8NAMpMqEF/HamlhgQfgV//
# 874kJzkq+jzZ5rn/Zc/aQ3398FOKwADxu+3oIC1RNIEdevMr4HpmKkVR7h+be9Lk
# lkHYdkp/gCug3MBKCA1S5lWxqHskvOxYnfT1fvIiykjquqG3OzaGxgkearnoqX96
# ypKLQocwPxwrWBnY54SmBIRiM6+4Rlbe1XDn6Haf1a3Vc3YwlB29lCLBq0jEZcFK
# jGmBlRvIHyr24pJOKcwYvTqUN7d8W3YU5lNWPUAU/ygF0LJ4GvAN8oHeAHFiH5hf
# OiE6ciyWO1pZVZbR6D4i+pK4XkT8dZRTasg=
# SIG # End signature block
