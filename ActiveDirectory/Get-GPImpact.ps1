function Get-GPImpact
{
    [CmdletBinding()]

    Param
    (
        [Parameter(ParameterSetName = 'Name', Mandatory)]
        [string]
        $Name,

        [Parameter(ParameterSetName = 'ID', Mandatory)]
        [string]
        $ID
    )

    Begin
    {

    }
    Process
    {
        Switch ($PSCmdlet.ParameterSetName)
        {
            'Name'
            {
                Write-Verbose 'using name'
                Try
                {
                    $gpo = Get-GPO -Name $Name
                }
                Catch
                {
                    Write-Error $error[0]
                    Return
                }
                $scope = Get-GPPermission -All -Name $Name | Where-Object { $_.Permission -eq 'GpoApply' }
                Write-Verbose "Scope: $($scope.trustee.name)"
            }
            'ID'
            {
                Write-Verbose 'using id'
                Try
                {
                    $gpo = Get-GPO -Id $ID
                }
                Catch
                {
                    Write-Error $error[0]
                    Return
                }
                $scope = Get-GPPermission -All -Id $ID | Where-Object { $_.Permission -eq 'GpoApply' }
                Write-Verbose "Scope: $($scope.trustee.name)"
            }
        }

        $GpoName = $Gpo.DisplayName

        $ResultHashtable = @{}
        $ResultHashtable.GpoName = $GpoName

        $ous = Get-ADOrganizationalUnit -Filter * -Properties gplink | Where-Object { $_.gplink -like "*$($gpo.id)*" }

        if ((Get-ADDomain).LinkedGroupPolicyObjects -like "*$($gpo.id)*")
        {
            $ous += (Get-ADDomain).DistinguishedName
        }
        
        if ($ous)
        {
            foreach ($OU in $OUs)
            {
                Write-Verbose "OU: $($OU.Name)"
            }
            if ($scope.Trustee)
            {
                if ($scope.Trustee.Name -contains 'Authenticated Users')
                {
                    Write-Verbose 'Authenticated Users'

                    foreach ($ou in $ous)
                    {
                        Write-Verbose $ou.distinguishedname

                        $ResultHashtable.AffectedObjects = Get-ADObject -SearchBase $ou.distinguishedname -Filter * |
                            Where-Object { @('user', 'computer') -contains $_.objectclass } |
                            Select-Object Name, ObjectClass
                    }
                }
                else
                {
                    Write-Verbose 'Groups'

                    #Get group members within scope
                    $GroupMembers = foreach ($Group in ($scope | Where-Object { $_.Trustee.SidType -eq 'Group' }))
                    {
                        Get-ADGroupMember $Group.Trustee.Name 
                    }
                    $GroupMembers = $GroupMembers | Select-Object -Unique

                    #Find members the GPO applies to
                    if ($groupMembers)
                    {
                        Write-Verbose "GroupMember count: $($groupMembers.count)"
                        $ResultHashtable.AffectedObjects = foreach ($OU in $OUs)
                        {
                            Get-ADObject -SearchBase $ou.DistinguishedName -Filter * |
                                Where-Object { @('user', 'computer') -contains $_.objectclass } |
                                Where-Object { $groupMembers.Name -contains $_.name } |
                                Select-Object Name, ObjectClass
                        }
                    }

                    #Find all users and computers explicitly within scope
                    Write-Verbose 'Users'
                    $ResultHashtable.AffectedObjects += ForEach ($object in ($scope | Where-Object { @('user', 'computer') -contains $_.Trustee.SidType }))
                    {
                        Write-Verbose $Object.Trustee.Name

                        switch ($object.Trustee.SidType)
                        {
                            'User' { $object = Get-ADUser $object.Trustee.Name }
                            'Computer' { $object = Get-ADComputer $object.Trustee.Name }
                        }
                        if ($OUs.DistinguishedName -contains ($object.distinguishedname -creplace "^[^,]*,", ""))
                        {
                            Get-ADObject $object.distinguishedname
                        }
                        else {
                            $object
                        }
                    }
                }
            }
            else
            {
                Write-Error 'GPO has no scope.'
            }
        }
        else
        {
            Write-Host "GPO '$GpoName' has no links"
        }
    }
    End
    {
        [PSCustomObject]$ResultHashtable |
            Add-Member -MemberType ScriptProperty -Name AffectedUsers -Value {$this.AffectedObjects.Where{$_.ObjectClass -eq 'user'}} -PassThru |
            Add-Member -MemberType ScriptProperty -Name AffectedUsersCount -Value {$this.AffectedUsers.Count} -PassThru |
            Add-Member -MemberType ScriptProperty -Name AffectedComputers -Value {$this.AffectedObjects.Where{$_.ObjectClass -eq 'computer'}} -PassThru |
            Add-Member -MemberType ScriptProperty -Name AffectedComputersCount -Value {$this.AffectedComputers.Count} -PassThru
    }
}
