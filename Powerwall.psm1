function Get-FirewallRule {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False,Position=1)]
        [string]$Name
    )
    if($Name -eq ''){$Name='all'}
    netsh advfirewall firewall show rule name=$Name
}

function Update-FirewallRule {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=1)]
        [string]$Name,

        [Parameter(Mandatory=$True,Position=2)]
        [string]$RemoteIPs
    )
    
    Write-Verbose '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
    $RemoteIPField = (Get-FirewallRule $Name) | ? {$_.StartsWith('RemoteIP')}
    Write-Verbose $RemoteIPField
    Write-Verbose '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
    $RemoteIPList = $RemoteIPField.split()[-1]
    Write-Verbose $RemoteIPList

    if($RemoteIPList -eq 'Any') {
        Write-Verbose 'No Remote IPs previously set'
        $RemoteIPList = ''
        Write-Verbose 'Set field to ""'
    } else {
        $RemoteIPList += ','
    }
    
    Write-Verbose '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
    $NewRemoteIPList = $RemoteIPList + $RemoteIPs
    Write-Verbose $NewRemoteIPList
    Write-Verbose '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
    
    netsh advfirewall firewall set rule name=$Name new remoteip=$NewRemoteIPList
}