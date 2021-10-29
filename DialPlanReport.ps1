######################################################################################
# This will export your dial plans and normalization rules into a CSV File           #
# You will be prompted where to store the file.                                      #
# When opening the file in Excel, remember that translations with just $1 will need  #
# to have tick mark ' added to in front of the the $1 prior to using.                #
#                                                                                    #
######################################################################################


cls

# Retrieves All the dial plans
$DPs=Get-CsTenantDialPlan

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

        $details=@()
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
                    $details += $detail
                }
        }
        # exporting the array with all the dial plan information to the file name designated earlier. 
        $details |Export-Csv $filelocation -Append -NoTypeInformation

    }

# Removing any remaining variables

Remove-Variable DP
Remove-Variable DPs
Remove-Variable rule
Remove-Variable detail
Remove-Variable details
Remove-Variable filelocation
