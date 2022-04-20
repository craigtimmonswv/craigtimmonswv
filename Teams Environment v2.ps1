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


$filelocation = 'c:\temp\test.xlsx'
#$filelocation = Read-Host "Enter Location/filename to store output (i.e c:\scripts\test.xlsx)"
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
tabname = "Call Queue"
Write-DataToExcel $filelocation  $details $tabname
