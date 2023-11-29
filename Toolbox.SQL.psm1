Function Format-SqlHC {
    <#
    .SYNOPSIS
        Format a string to a readable SQL string and NULL to NULL.

    .DESCRIPTION
        Format a string to a readable SQL string and NULL to NULL. This is not
        usable for DateTime fields.

    .PARAMETER x
        String or NULL

    .EXAMPLE
        Set-Alias -Name FS -Value Format-SqlHC
        Invoke-Sqlcmd @SQLParams -Query "
            UPDATE $SQLTableTickets
            SET
            AssignToGroup = $(Format-SqlHC $Value),
            AssignToStaff = 'No'
            WHERE PSCode = '$PSCode'"

        If $Value contains 'Marie D'hondt' we will write 'Marie D''Hondt' to
        the SQL table. If $Value doesn't exist or is equal to NULL, we will
        write NULL to the SQL table.
    #>

    Param (
        $x
    )
    if (($x -eq $null) -or ($x -eq '')) {
        'NULL'
    }
    else {
        Switch -Wildcard ($x.GetType().Name) {
            '*Int*' {
                "'" + $x + "'"
            }
            'Boolean' {
                "'" + $x + "'"
            }
            'DateTime' {
                "'" + $x.ToString('yyyy-MM-dd HH:mm:ss') + "'"
            }
            Default {
                "'" + $x.Replace("'", "''") + "'"
            }
        }
    }
}
Function Save-TicketInSqlHC {
    <#
    .SYNOPSIS
        Save the ticket information in the SQL database.

    .DESCRIPTION
        Create a new row in the SQL table 'Tickets' for each created ticket
        containing all its details.

    .PARAMETER KeyValuePair
        One or more hash tables containing key value pair combinations to
        create tickets. The key represents the field name in a Cherwell ticket
        and the value represents the value for that specific field.

    .PARAMETER TicketNr
        Number of the ticket.

    .PARAMETER PSCode
        Unique identifier to match the row in the SQL tickets table with the
        row in the table for that specific script.

    .PARAMETER Force
        Overwrite a row when that PSCode is already present.
    #>

    [OutputType()]
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory)]
        [HashTable[]]$KeyValuePair,
        [Parameter(Mandatory)]
        [String]$TicketNr,
        [Parameter(Mandatory)]
        [String]$PSCode,
        [Parameter(Mandatory)]
        [String]$ScriptName,
        [String]$ComputerName = 'GRPSDFRAN0049',
        [String]$Database = 'PowerShell',
        [String]$Table = 'Tickets',
        [Switch]$Force
    )

    Process {
        Try {
            $SQLParams = @{
                ServerInstance         = $ComputerName
                Database               = $Database
                TrustServerCertificate = $true
                QueryTimeout           = '1000'
                ConnectionTimeout      = '20'
                ErrorAction            = 'Stop'
                Verbose                = $false
            }

            $Date = Get-Date

            $SQLResult = Invoke-Sqlcmd @SQLParams -Query "
                SELECT *
                FROM $Table
                WHERE PSCode = '$PSCode'"



            $InsertRow = {
                $null = Invoke-Sqlcmd @SQLParams -Query "
                    INSERT INTO $Table
                    (PSCode, ScriptName, ServiceCountryCode,
                    TicketRequestedDate, TicketCreationDate, TicketNr,
                    Requester, SubmittedBy,
                    OwnedByTeam, OwnedBy,
                    ShortDescription, Source, Service,
                    Category, SubCategory, IncidentType, Priority)
                    VALUES ($(FSQL $PSCode), $(FSQL $ScriptName), $(FSQL $KeyValuePair.ServiceCountryCode),
                    $(FSQL $Date), $(FSQL $Date), $(FSQL $TicketNr),
                    $(FSQL $KeyValuePair.Requester), $(FSQL $KeyValuePair.SubmittedBy),
                    $(FSQL $KeyValuePair.OwnedByTeam), $(FSQL $KeyValuePair.OwnedBy),
                    $(FSQL $KeyValuePair.ShortDescription),
                    $(FSQL $KeyValuePair.Source), $(FSQL $KeyValuePair.Service),
                    $(FSQL $KeyValuePair.Category), $(FSQL $KeyValuePair.SubCategory),
                    $(FSQL $KeyValuePair.IncidentType), $(FSQL $KeyValuePair.Priority))"
            }

            $UpdateRow = {
                $null = Invoke-Sqlcmd @SQLParams -Query "
                    UPDATE $Table SET
                    ScriptName = $(FSQL $ScriptName),
                    ServiceCountryCode = $(FSQL $KeyValuePair.ServiceCountryCode),
                    TicketRequestedDate = $(FSQL $Date),
                    TicketCreationDate = $(FSQL $Date),
                    TicketNr = $(FSQL $TicketNr),
                    Requester = $(FSQL $KeyValuePair.Requester),
                    SubmittedBy = $(FSQL $KeyValuePair.SubmittedBy),
                    OwnedByTeam = $(FSQL $KeyValuePair.OwnedByTeam),
                    OwnedBy = $(FSQL $KeyValuePair.OwnedBy),
                    ShortDescription = $(FSQL $KeyValuePair.ShortDescription),
                    Source = $(FSQL $KeyValuePair.Source),
                    Service = $(FSQL $KeyValuePair.Service),
                    Category = $(FSQL $KeyValuePair.Category),
                    SubCategory = $(FSQL $KeyValuePair.SubCategory),
                    IncidentType = $(FSQL $KeyValuePair.IncidentType),
                    Priority = $(FSQL $KeyValuePair.Priority)
                    WHERE PSCode = $(FSQL $PSCode)"
            }

            if ($Force) {
                if (-not $SQLResult) {
                    & $InsertRow
                }
                else {
                    & $UpdateRow
                }
            }
            else {
                if ($SQLResult) {
                    throw "A row with PSCode '$PSCode' is already present, please use the '-Force' parameter to overwrite the row."
                }

                & $InsertRow
            }
        }
        Catch {
            $P = $_
            $Global:Error.RemoveAt(0)
            throw "Failed saving ticket information for TicketNr '$TicketNr' in SQL table '$Table', Database '$Database', ComputerName '$ComputerName': $P"
        }
    }
}

New-Alias -Name FSQL -Value Format-SqlHC

Export-ModuleMember -Function * -Alias *