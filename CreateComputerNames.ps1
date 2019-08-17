# .\CreateComputerNames.ps1
# Sample usage:
# .\CreateComputerNames.ps1 MyComputer 11 "OU=MyOU,DC=mydomain,DC=com"
# will create 11 new computers MyComputer01 to MyComputer11 in the MyOU Organizational Unit for mydomain.com
# Input parameters $FirstName, $Totalnames, and $OUPath
# example OU path: "OU=Inward OU,OU=Next Level UP OU,OU=Third Level UP OU,OU=Top Level OU,DC=domain,DC=edu"
# simply copy and paste the full path to the OU from AD to the input parameter
# - Use the Get-ADObject to find the OU you wish to use e.g. Get-ADObject -Filter {name -like "*My Fancy OU*"}


Param([string]$FirstName=(Read-Host -Prompt 'Please enter the first part of the host name, total computers, and full OU path'),[int]$TotalNames,[string]$OUPath)

Clear-Content newcomps.txt

#Check to see if computer name exists in AD, and if so set the starting computer number to the last computer number plus one
$CompName="*$FirstName*"

$compExist=(Get-ADComputer -Filter {name -like $CompName} )
If ($compExist){$compstartnumber=$compExist.count + 1;$TotalNames=$TotalNames+$compstartnumber-1}


try {
    $ouExists=([adsi]::Exists("LDAP://$OUPath"))
}
catch {
    Write-Output "The OU you entered does NOT exist!"
}

If ($ouExists){ 
    $compstartnumber..$TotalNames |%  {if($_ -le 9){$_="0$_"};$ComputerName="$FirstName$_".ToUpper(); new-adcomputer -Name "$ComputerName" -SamAccountName "$ComputerName" -Path "$OUPath";Write-Output "New Computer names: $ComputerName, OU Path is $OUPath `n" >> newcomps.txt;}
    Invoke-Item newcomps.txt
}
Else {
    Write-Output "Please enter the correct OU in the format of: OU=myou,DC=domain,DC=com for the full OU path." 
    Write-Error "Invalid path, OU not found!"
}


