cls
Write-Host "This is will create an Excel Spreadsheet.  Make sure to enter the file name with .xlsx"
Write-Host "You will need to verify that you have installed the exportexcel module"
$filelocation = Read-Host "Enter Location/filename to store output (i.e c:\scripts\test.xlsx)"

$VRs = Get-CsOnlineVoiceRoute
# Tests to see if the file currently exists.  It will stop if it does. 
$vrdetails = @()
        foreach ($VR in $VRs)
            {   
            
                [string] $usage= $vr.OnlinePstnUsages
                [string] $pstngw =$vr.OnlinePstnGatewayList 
                $detail = New-Object PSObject
                $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $vr.Name
                $detail | Add-Member -MemberType NoteProperty -Name "NumberPattern" -Value $vr.NumberPattern
                $detail | Add-Member -MemberType NoteProperty -Name "OnlinePstnUsages" -Value $usage
                $detail | Add-Member -MemberType NoteProperty -Name "OnlinePstnGatewayList " -Value $pstngw
                $VRdetails += $detail
            }

$VRdetails |Export-Excel -Path $filelocation -WorksheetName "Voice Routes" -AutoSize 

# Extract PSTN Gateways
$PSTNGWs = Get-CsOnlinePSTNGateway
$pstngwDetails = @()
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
                $pstngwDetails += $detail
            }
$excelpackage = Open-ExcelPackage -Path $filelocation 
$ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName "PSTN Gateways" 
$pstngwDetails | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter


# Extracts Dial Plan info
$DPs=Get-CsTenantDialPlan
$dpDetails = @()

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
                $dpDetails += $detail
                }
        }
$excelpackage = Open-ExcelPackage -Path $filelocation 
$ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName "Dial Plans" 
$DPDetails | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter

# Extracts Voice Routes
$PSTNUSAGEs = Get-CsOnlineVoiceRoute
$Usagedetails = @()
        foreach ($PSTNUsage in $PSTNUSAGEs)
            {   [string]$usage = $PSTNUsage.OnlinePstnUsages
                $detail = New-Object PSObject
                $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $PSTNUSAGE.Identity
                $detail | add-Member -MemberType NoteProperty -Name "Usage" -Value $usage
                $Usagedetails += $detail
            }
$excelpackage = Open-ExcelPackage -Path $filelocation 
$ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName "PSTN Usages" 
$Usagedetails | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter

# Extracts users enablement
$users =  Get-CsOnlineUser | ?  {$_.enterprisevoiceenabled -eq $true}
        $Userdetails = @()
        foreach ($user in $users)
            {
                # Creating an array to store the variables from the dial plans. 
                $detail = New-Object PSObject
                $detail | add-Member -MemberType NoteProperty -Name "Displayname" -Value $user.displayname
                $detail | add-Member -MemberType NoteProperty -Name "UPN" -Value $user.UserPrincipalName
                $detail | add-Member -MemberType NoteProperty -Name "Lineuri" -Value $user.LineUri
                $detail | add-Member -MemberType NoteProperty -Name "Dial Plan" -Value $user.TenantDialPlan
                $detail | add-Member -MemberType NoteProperty -Name "Voice Routing Policy" -Value $user.OnlineVoiceRoutingPolicy
                $detail | add-Member -MemberType NoteProperty -Name "EV Enabled" -Value $user.EnterpriseVoiceEnabled
                $detail | add-Member -MemberType NoteProperty -Name "Teams Upgrade Policy" -Value $user.TeamsUpgradePolicy
                $detail | add-Member -MemberType NoteProperty -Name "Teams Effective Mode" -Value $user.TeamsUpgradeEffectiveMode
                $Userdetails += $detail

            }
$excelpackage = Open-ExcelPackage -Path $filelocation 
$ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName "User Information" 
$Userdetails | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter

#Extracts Voice Routing Policies
$VRPs = Get-CsOnlineVoiceRoutingPolicy
$vrpDetails= @()
foreach ($VRP in $VRPs)
            {       
                foreach ($usage in $vrp.OnlinePstnUsages)
                {
                    $detail = New-Object PSObject
                    $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $VRP.Identity
                    $detail | Add-Member -MemberType NoteProperty -Name "Description" -Value $VRP.Description
                    $detail | Add-Member -MemberType NoteProperty -Name "OnlinePstnUsages" -Value $usage
                    
                    $vrpDetails += $detail
                }
            }

$excelpackage = Open-ExcelPackage -Path $filelocation 
$ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName "Voice Routing Policies" 
$vrpDetails | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter

# Extract Emergency Calling Policies
$ercallpolicies = Get-CsTeamsEmergencyCallingPolicy
        $ECPdetails = @()
        foreach ($ercp in $ercallpolicies)
            {
                
                        $detail = New-Object PSObject
                        $detail | add-Member -MemberType NoteProperty -Name "Identity" -Value $ercp.Identity
                        $detail | add-Member -MemberType NoteProperty -Name "Description" -Value $ercp.Description
                        $detail | add-Member -MemberType NoteProperty -Name "NotificationGroup" -Value $ercp.NotificationGroup
                        $detail | add-Member -MemberType NoteProperty -Name "ExternalLocationLookupMode" -Value $ercp.ExternalLocationLookupMode
                        $detail | add-Member -MemberType NoteProperty -Name "NotificationDialOutNumber" -Value $ercp.NotificationDialOutNumber
                        $detail | add-Member -MemberType NoteProperty -Name "NotificationMode" -Value $ercp.NotificationMode
                        $ECPdetails += $detail  
            }

$excelpackage = Open-ExcelPackage -Path $filelocation 
$ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName "Emergency Calling Policies" 
$ECPdetails | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter


# Extracts Emergency Call Routing Policy
$ecrps = Get-CsTeamsEmergencyCallRoutingPolicy
        $ECRPdetails = @()
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
                                $ECRPdetails  += $detail  
                            }

            }
$excelpackage = Open-ExcelPackage -Path $filelocation 
$ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName "Emergency Call Routing Policies" 
$ECRPdetails | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter


$erlocations = Get-CsTenantNetworkSite
$NetSitedetails = @()
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
                    $NetSitedetails += $detail  
                }

        }
$excelpackage = Open-ExcelPackage -Path $filelocation 
$ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName "Network Topology Site Detail" 
$NetSitedetails | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter

# Emergency Location information Services 

$locations = Get-CsOnlineLisLocation
$LISLocDetails = @()
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
        $LISLocDetails += $detail
        }
$excelpackage = Open-ExcelPackage -Path $filelocation 
$ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName "LIS Location " 
$LISLocDetails | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter

#LIS Network information
$subnets = Get-CsOnlineLisSubnet
$subnetDetails = @()
Foreach ($subnet in $subnets)
    {
        $detail = New-Object PSObject
        
        $detail | Add-Member NoteProperty -Name "Subnet" -Value $subnet.Subnet
        $detail | Add-Member NoteProperty -Name "Description" -Value $subnet.Description
        $subloc = Get-CsOnlineLisLocation -LocationId $subnet.LocationId
        $detail | Add-Member NoteProperty -Name "Location" -Value $subloc.location
        $detail | Add-Member NoteProperty -Name "City" -Value $subloc.city
        $subnetDetails += $detail
    }
$excelpackage = Open-ExcelPackage -Path $filelocation 
$ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName "LIS Subnet " 
$subnetDetails | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter

#LIS Wireless Access Point information
$WAPs = Get-CsOnlineLisWirelessAccessPoint
$WAPSDetails = @()
Foreach ($WAP in $WAPs)
    {
        $detail = New-Object PSObject
        $detail | Add-Member NoteProperty -Name "BSSID" -Value $WAP.BSSID
        $detail | Add-Member NoteProperty -Name "Description" -Value $WAP.Description
        $WAPloc = Get-CsOnlineLisLocation -LocationId $WAP.LocationId
        $detail | Add-Member NoteProperty -Name "Location" -Value $WAPloc.location
        $detail | Add-Member NoteProperty -Name "City" -Value $WAPloc.city
        $WAPSDetails += $detail
    }
$excelpackage = Open-ExcelPackage -Path $filelocation 
$ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName "LIS WAP " 
$WAPSDetails | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter

#LIS Switch information
$Switches = Get-CsOnlineLisSwitch
$SwitchDetails = @()
Foreach ($Switch in $Switches)
    {
        $detail = New-Object PSObject
        
        $detail | Add-Member NoteProperty -Name "ChassisID" -Value $Switch.ChassisID
        $detail | Add-Member NoteProperty -Name "Description" -Value $Switch.Description
        $Switchloc = Get-CsOnlineLisLocation -LocationId $Switch.LocationId
        $detail | Add-Member NoteProperty -Name "Location" -Value $Switchloc.location
        $detail | Add-Member NoteProperty -Name "City" -Value $Switchloc.city
        $SwitchDetails += $detail
    }
$excelpackage = Open-ExcelPackage -Path $filelocation 
$ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName "LIS Switch Details" 
$SwitchDetails| Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter

#LIS Port information
$Ports = Get-CsOnlineLisPort
$PortDetails = @()
Foreach ($port in $ports)
    {
        $detail = New-Object PSObject
        
        $detail | Add-Member NoteProperty -Name "ChassisID" -Value $port.ChassisID
        $detail | Add-Member NoteProperty -Name "PortID" -Value $port.PortID
        $detail | Add-Member NoteProperty -Name "Description" -Value $port.Description
        $portloc = Get-CsOnlineLisLocation -LocationId $port.LocationId
        $detail | Add-Member NoteProperty -Name "Location" -Value $portloc.location
        $detail | Add-Member NoteProperty -Name "City" -Value $portloc.city
        $PortDetails += $detail
    }
$excelpackage = Open-ExcelPackage -Path $filelocation 
$ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName "LIS Switch Port Details" 
$PortDetails| Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter