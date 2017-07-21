# SCOM DB Ghost Object Grooming
# Cole McDonald - Beyond Impact, LLC
# 2017 Jul 21

# Built and tested on SCOM 2012 R2
# SQL commands came from somewhere else during an MS Premiere troubleshooting session

# !! Needs to be run under an account with access to the SCOM SQL from one of the Management servers

# Grab the OperationsManager DB info
$OpsMgrDB = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft Operations Manager\3.0\Setup').DatabaseServername -split "\\"

# Grab the SQL module from a known SQL server
if (!(get-command Invoke-Sqlcmd -ErrorAction SilentlyContinue)) {
    Write-Verbose "Importing SQLPS from $($OpsMgrDB[0])"
    $session = New-PSSession -ComputerName $OpsMgrDB[0]
    Import-PSSession -Module SQLPS -Session $session
} else {
    Write-verbose "SQLPS already exists on the system"
}

# Gather Agents, Magement Servers, and Gateways to match against
$ms = Get-SCOMManagementServer
$gw = Get-SCOMGatewayManagementServer
$agents = Get-SCOMAgent

# Gather computer instances
$insts = Get-SCOMClassInstance `
| where fullname -like "Microsoft.Windows.Computer*" `
| sort-object displayname

# Identify computer object instances that claim to be agent managed, but don't occur in the agents list
# Excepting computers registered as management servers or gateways
$ghosts = @()
foreach ($inst in $insts) {
    if (
        ($agents.displayName -notcontains $inst.DisplayName) `
        -and ($ms.DisplayName -notcontains $inst.DisplayName) `
        -and ($gw.DisplayName -notcontains $inst.displayname)
    ) {
        if ($inst.'[System.Mom.BackwardCompatibility.Computer].Management_Mode'.Value -EQ "Agent") {
            Write-Verbose "*** $($inst.displayname)`t`t$($inst.fullname)"
            $ghosts += @($inst)
        }
    }
}

# Offer gridview for the user to select the servers they'd like to remove from the DB
$selectedGhosts = $ghosts | Select-Object displayname, fullName, ID | Out-GridView -PassThru

# Loop through each of the selected instances and verifies, marks for deletion, and grooms them from the DB
foreach ($ghost in $selectedGhosts) {
    Write-Verbose "Verifying $($ghost.displayName)"
    $q_verify = "select * from MT_HealthService where isagent=1 and DisplayName = '$($ghost.displayName)';"
    $r_verify = Invoke-Sqlcmd -ServerInstance "$($OpsMgrDB[0])\$($OpsMgrDB[1])" -Database "OperationsManager" -Query $q_verify

    if ($r_verify.IsAgent) {
        Write-Verbose "-- Marking $($ghost.displayName) for deletion"
        $q_delete = @"
            -- Mark for removal
            DECLARE @EntityId uniqueidentifier;
            DECLARE @TimeGenerated datetime;

            -- change "GUID" to the ID of the invalid entity
            SET @EntityId = '$($ghost.id)';
            SET @TimeGenerated = getutcdate();

            -- Execute the transaction
            BEGIN TRANSACTION
                EXEC dbo.p_TypedManagedEntityDelete @EntityId, @TimeGenerated;
            COMMIT TRANSACTION
"@
        $r_delete = Invoke-Sqlcmd -ServerInstance "$($OpsMgrDB[0])\$($OpsMgrDB[1])" -Database "OperationsManager" -Query $q_delete

        Write-Verbose "-- Grooming $($ghost.displayName)"
        $q_groom = @"
            -- Groom Once marked for removal
            DECLARE @GroomingThresholdUTC datetime
            DECLARE @BatchSize int
            SET @GroomingThresholdUTC = getutcdate()
            SET @Batchsize = 250

            -- Execute the transaction
            BEGIN TRANSACTION
                EXEC dbo.p_DiscoveryDataPurgingByTypedManagedEntityInBatches @GroomingThresholdUTC, @BatchSize
                EXEC dbo.p_DiscoveryDataPurgingByRelationshipInBatches @GroomingThresholdUTC, @BatchSize
                EXEC dbo.p_DiscoveryDataPurgingByBaseManagedEntityInBatches @GroomingThresholdUTC, @BatchSize
            COMMIT TRANSACTION
"@
        $r_groom = Invoke-Sqlcmd -ServerInstance "$($OpsMgrDB[0])\$($OpsMgrDB[1])" -Database "OperationsManager" -Query $q_groom
    } else {
        Write-Verbose "-- $($ghost.displayname) not verified"
    }
}