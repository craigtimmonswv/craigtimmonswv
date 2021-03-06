<#
You will need have the "ImportExcel" Module installed for this to properly run. 
You can get it here:
https://www.powershellgallery.com/packages/ImportExcel/7.4.1
To install it run: 
Install-Module -Name ImportExcel -RequiredVersion 7.4.1
Import-Module -Name ImportExcel
This will pull the basic environment from the Teams tenant. Items it gathers is:
PSTN Gateways
PSTN Usages
Voice Routes
Voice Routing Policies
Dial Plan
Voice enabled users - this might take a while depending upon number of users
Emergency Calling Policies
Emergency Call Routing Policies
Tenant Network Site Details
LIS Locations
LIS Network Information
LIS WAP Information
LIS SWitch information
LIS Port
Auto Attendant
Call Queue
It will place the Excel spreadsheet it in the location you enter when prompted. 
#>

Function Write-DataToExcel
    {
        param ($filelocation, $details, $tabname)

        
        $excelpackage = Open-ExcelPackage -Path $filelocation 
        $ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName $tabname
        $details | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter
        Clv details 

    }
Function Get-TeamsEnvironment
{
            param ($filelocation)
            $Details = @()
            $IncEmployees = Read-host "Include Enterprise Voice Users (y/n)"
            Write-Host "Running"
            # Extract PSTN Gateways
            Write-Host 'Getting Online PSTN Gateway Details'
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
            
            # Get PSTN Usages
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

            # Get Voice Routes
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

            # Get Voice Routing Policies
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
            $tabname = "Voice Routing Policies"
            Write-DataToExcel $filelocation  $details $tabname

            # Get Dial Plan info
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

            # Get users enablement
            If ($IncEmployees -eq "y" -or $IncEmployees -eq "Y")
                {
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
                }
            Else
                {
                    Write-Host "Skipping Enterprise Voice Users"
                }

            # Get Emergency Calling Policies
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

            # Get Emergency Call Routing Policy
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
            $tabname = "Emergency Call Routing Policies"
            Write-DataToExcel $filelocation  $details $tabname

            # Get Tenant Network Site Details
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

            # Get Emergency Location information Services 
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

            # Get LIS Network information
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

            #Get LIS Wireless Access Point information
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

            #Get LIS Switch information
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

            #Get LIS Port information
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
            
            #Get Auto Attendant Details
            Write-Host "Getting Auto Attendant Details"
$aas= Get-CsAutoAttendant
            $details =@()
            foreach ($aa in $AAs)
                {
                try {clv ResouceAct -ErrorAction SilentlyContinue } catch{}
                try {clv OperatorID -ErrorAction SilentlyContinue } catch{}
                try {clv Operator -ErrorAction SilentlyContinue } catch{}
                
                foreach ($RA in $aa.ApplicationInstances)
                    {
                        $ResouceAct = Get-CsOnlineApplicationInstance -Identity $ra
                        try {
                            $operatorID = ((Get-CsAutoAttendant -NameFilter $aa.Name | select operator).operator).id 
                            }
                        catch{$Error}
                        
                        if (!($operatorID))
                            {}
                            Else
                                {
                                
                                 $Operator = (Get-CsOnlineUser -Identity $operatorID -erroraction SilentlyContinue | select UserPrincipalName).UserPrincipalName
                                 
                                }
                        
                        $detail = New-Object PSObject
                        $detail | Add-Member NoteProperty -Name "AAName" -Value $aa.name
                        $detail | Add-Member NoteProperty -Name "Identity" -Value $aa.identity
                        $detail | Add-Member NoteProperty -Name "Language" -Value $aa.LanguageId
                        $detail | Add-Member NoteProperty -Name "TimeZone" -Value $aa.timezoneid
                        $detail | Add-Member NoteProperty -Name "Operator" -Value $Operator
                        $detail | Add-Member NoteProperty -Name "VoiceResponseEnabled" -Value $aa.VoiceresponseEnabled
                        $detail | Add-Member NoteProperty -Name "ResourceAccount" -Value $ResouceAct.UserPrincipalName
                        $detail | Add-Member NoteProperty -Name "Phone Number" -Value $ResouceAct.PhoneNumber
                        $details += $detail
                        clv detail
                    }
            }
            $tabname = "Auto Attendant"
            Write-DataToExcel $filelocation  $details $tabname

            # Get Call Queues Details
            Write-Host "Getting Call Queue Details"
            $Details = @()
           try { $CQs = Get-CsCallQueue -erroraction silentlycontinue -WarningAction silentlyContinue}
            Catch {Write-Warning }
            
            foreach ($CQ in $CQs)
                {
                    try {clv agent -ErrorAction SilentlyContinue } catch{}
                    try {clv agents -ErrorAction SilentlyContinue } catch{}
                    try {clv OFATarget -ErrorAction SilentlyContinue } catch{}
                    try {clv TOATarget -ErrorAction SilentlyContinue } catch{}
                    $detail = New-Object PSObject
                    $detail | Add-Member NoteProperty -Name "CQName" -Value $CQ.name
                    $detail | Add-Member NoteProperty -Name "Identity" -Value $CQ.identity
                    $detail | Add-Member NoteProperty -Name "RoutingMethod" -Value $CQ.RoutingMethod
                    $detail | Add-Member NoteProperty -Name "AllowOptOut" -Value $CQ.AllowOptOut                
                    $detail | Add-Member NoteProperty -Name "ConferenceMode" -Value $CQ.ConferenceMode
                    $detail | Add-Member NoteProperty -Name "PresenceBasedRouting" -Value $CQ.PresenceBasedRouting
                    foreach ($a in $cq.Agents.objectid)
                        {
                            try {$agent=(get-csonlineuser -Identity $a -erroraction SilentlyContinue| select UserPrincipalName).UserPrincipalName + ","}
                            Catch {}
                            $agents +=$agent
                        }
                    $detail | Add-Member NoteProperty -Name "Agents" -Value $agents
                    $detail | Add-Member NoteProperty -Name "AgentAlertTime" -Value $CQ.AgentAlertTime
                    $detail | Add-Member NoteProperty -Name "LanguageId" -Value $CQ.LanguageId
                    $detail | Add-Member NoteProperty -Name "OverflowThreshold" -Value $CQ.OverflowThreshold
                    $detail | Add-Member NoteProperty -Name "OverflowAction" -Value $CQ.OverflowAction
                    try 
                        {$OFATarget = ((Get-CsCallQueue -NameFilter $cq.Name| select OverflowActionTarget).OverflowActionTarget).id 
                        if ($OFATarget)
                        {
                            $OFATargetUser = (get-csonlineuser -Identity $OFATarget -erroraction SilentlyContinue| select UserPrincipalName).UserPrincipalName
                        }
                        
                        }
                    Catch {}
                    
                    $detail | Add-Member NoteProperty -Name "OverflowActionTarget" -Value $OFATargetUser
                    $detail | Add-Member NoteProperty -Name "OverflowSharedVoicemailTextToSpeechPrompt" -Value $CQ.OverflowSharedVoicemailTextToSpeechPrompt
                    $detail | Add-Member NoteProperty -Name "EnableOverflowSharedVoicemailTranscription" -Value $CQ.EnableOverflowSharedVoicemailTranscription
                    $detail | Add-Member NoteProperty -Name "TimeoutThreshold" -Value $CQ.TimeoutThreshold
                    $detail | Add-Member NoteProperty -Name "TimeoutAction" -Value $CQ.TimeoutAction
                    try 
                        {$TOATarget = ((Get-CsCallQueue -NameFilter $cq.Name| select TimeoutActionTarget).TimeoutActionTarget).id
                        if ($TOATarget)
                        {
                            $TOATargettUser = (get-csonlineuser -Identity $TOATarget -erroraction SilentlyContinue | select UserPrincipalName).UserPrincipalName
                        }
                        
                         }
                    Catch {}
                    
                    $detail | Add-Member NoteProperty -Name "TimeoutActionTarget" -Value $TOATargettUser
                    $detail | Add-Member NoteProperty -Name "TimeoutSharedVoicemailTextToSpeechPrompt" -Value $CQ.TimeoutSharedVoicemailTextToSpeechPrompt
                    $detail | Add-Member NoteProperty -Name "EnableTimeoutSharedVoicemailTranscription" -Value $CQ.EnableTimeoutSharedVoicemailTranscription
                    $details += $detail
                }
            $tabname = "Call Queue"
            Write-DataToExcel $filelocation  $details $tabname
}
cls
Write-Host "This is will create an Excel Spreadsheet.  Make sure to enter the file name with .xlsx"
Import-Module ImportExcel
$filelocation = Read-Host "Enter Location/filename to store output (i.e c:\scripts\test.xlsx)"

# Determine if ImportExcel module is loaded
$XLmodule = Get-Module -Name importexcel


if ($XLmodule )
    {
        If ( $connected=get-cstenant -ErrorAction SilentlyContinue)
        {
            write-host "Current Tenant:" $connected.displayname
            Get-TeamsEnvironment $filelocation
        }
                Else {Write-Host "Teams module isn't loaded.  Please load Teams Module (connect-microsoftteams)"  }
    }
    Else {Write-Host "ImportExcel module is not loaded"}
