
Param(
    $a = (Get-Date (Get-Date   ).AddDays(-7) -Format 'yyyy-MM-dd'), 
    $b = (Get-Date (Get-Date $a).AddDays(45) -Format 'yyyy-MM-dd')
    )

# Upcoming auctions
#
# https://www.treasurydirect.gov/auctions/upcoming/

# $a = if ($a -eq $null) { Get-Date (Get-Date).AddDays(-7) -Format 'yyyy-MM-dd' } else { $a }

# $b = if ($b -eq $null) { Get-Date (Get-Date $a).AddDays(40) -Format 'yyyy-MM-dd' } else { $b }

# $a = '2023-01-26'
# $a = '2023-02-06'
# $b = '2023-03-13'
# $b = '2023-04-01'

$result_issued    = Invoke-RestMethod ('http://www.treasurydirect.gov/TA_WS/securities/search?issueDate={0},{1}&format=json'    -f $a, $b)
$result_maturing  = Invoke-RestMethod ('http://www.treasurydirect.gov/TA_WS/securities/search?maturityDate={0},{1}&format=json' -f $a, $b)
$result_auctioned = Invoke-RestMethod ('http://www.treasurydirect.gov/TA_WS/securities/search?auctionDate={0},{1}&format=json'  -f $a, $b)

# ----------------------------------------------------------------------
# workaround for newer versions of PowerShell
foreach ($row in $result_issued + $result_maturing + $result_auctioned)
{
    $row.issueDate    = Get-Date $row.issueDate    -Format 's'
    $row.maturityDate = Get-Date $row.maturityDate -Format 's'
    $row.auctionDate  = Get-Date $row.auctionDate  -Format 's'
}
# ----------------------------------------------------------------------

$auction_dates        = $result_auctioned | Group-Object auctionDate | ForEach-Object Name
$auction_issued_dates = $result_auctioned | Group-Object issueDate | ForEach-Object Name

function date-range ($a, $n)
{
    foreach ($i in (0..($n-1)))
    {
        Get-Date (Get-Date $a).AddDays($i) -Format 'yyyy-MM-dd'
    }
}



function get-sum ($ls)
{
    if ($ls -eq $null) { 0 } else { ($ls | Measure-Object -Property totalAccepted -Sum).Sum }
}


$days = ((Get-Date $b) - (Get-Date $a)).TotalDays

# $table = foreach ($date in (date-range '2023-01-26' 31))
$table = foreach ($date in (date-range $a $days))
{
    $issued   = $result_issued   | Where-Object issueDate    -Match $date
    $maturing = $result_maturing | Where-Object maturityDate -Match $date
    
    $issued_bills     = $result_issued   | Where-Object issueDate -Match $date    | Where-Object securityType -EQ Bill
    $issued_notes     = $result_issued   | Where-Object issueDate -Match $date    | Where-Object securityType -EQ Note
    $issued_bonds     = $result_issued   | Where-Object issueDate -Match $date    | Where-Object securityType -EQ Bond

    $maturing_bills   = $result_maturing | Where-Object maturityDate -Match $date | Where-Object securityType -EQ Bill
    $maturing_notes   = $result_maturing | Where-Object maturityDate -Match $date | Where-Object securityType -EQ Note
    $maturing_bonds   = $result_maturing | Where-Object maturityDate -Match $date | Where-Object securityType -EQ Bond    
    
    $issued_bills_sum = get-sum $issued_bills;   $maturing_bills_sum = get-sum $maturing_bills
    $issued_notes_sum = get-sum $issued_notes;   $maturing_notes_sum = get-sum $maturing_notes
    $issued_bonds_sum = get-sum $issued_bonds;   $maturing_bonds_sum = get-sum $maturing_bonds

    $issued_sum   = get-sum $issued
    $maturing_sum = get-sum $maturing

    [pscustomobject]@{
        date = $date
        issued_bills_sum = $issued_bills_sum;   maturing_bills_sum = $maturing_bills_sum;   bills_change = $issued_bills_sum - $maturing_bills_sum
        issued_notes_sum = $issued_notes_sum;   maturing_notes_sum = $maturing_notes_sum;   notes_change = $issued_notes_sum - $maturing_notes_sum
        issued_bonds_sum = $issued_bonds_sum;   maturing_bonds_sum = $maturing_bonds_sum;   bonds_change = $issued_bonds_sum - $maturing_bonds_sum
        
        issued = $issued_sum
        maturing = $maturing_sum
        change = $issued_sum - $maturing_sum

        auction         = if (($auction_dates        | Where-Object { $_ -match $date }) -eq $null) { '' } else { '*' }
        # auction_issuing = if (($auction_issued_dates | Where-Object { $_ -match $date }) -eq $null) { '' } else { '*' }

        auction_issuing = (($result_auctioned | Group-Object issueDate | Where-Object Name -Match $date).Group | ForEach-Object { if ($_ -ne $null) {$_.auctionDate.Substring(0,10) }} | Sort-Object -Unique) -join ' '

        offering_amount  = (($result_auctioned | Group-Object issueDate | Where-Object Name -Match $date).Group | Measure-Object offeringAmount -Sum).Sum

        # (($result_auctioned | Group-Object issueDate)[1].Group | Group-Object auctionDate | ForEach-Object Name)[0].Substring(0,10)
    }
}
# ----------------------------------------------------------------------
$weekend = 0

foreach ($row in $table)
{
    if ((Get-Date $row.date).DayOfWeek -in 'Saturday', 'Sunday')
    {
        $weekend = $weekend + $row.change
    }
    else
    {
        if ($weekend -ne 0)
        {
            $row | Add-Member -MemberType NoteProperty -Name change_with_weekend -Value ($row.change + $weekend)

            $weekend = 0
        }
    }
}
# ----------------------------------------------------------------------

# function field-format ($name)
# {
#     @{ Label = $name; Expression = { $_.$name.ToString('N0') }; Align = 'right' }
# }

# $table | Format-Table date, issued, maturing, @{ Label = 'change'; Expression = { $_.change.ToString('N0') }; Align = 'right' }

# $table | Format-Table (field-format 'date'), (field-format 'issued'), (field-format 'maturing'), (field-format 'change')

# function format-to-billions ($val)
# {
#     if ($val -eq 0)
#     {
#         ''
#     }
#     else
#     {
#         ($val / 1000 / 1000 / 1000).ToString('N0')
#     }    
# }

function format-to-billions ($val)
{
    ($val / 1000 / 1000 / 1000).ToString('N0')    
}

Write-Host '               BILLS                  NOTES                  BONDS                  TOTAL' -NoNewline
#           date           issued maturing change issued maturing change issued maturing change issued maturing change auction auction_issuing                  offering_amount

$fields = @(
    
    @{ Label = 'date'; Expression = { Get-Date $_.date -Format 'yyyy-MM-dd ddd' } } 
        
    @{ Label = 'issued';   Expression = { format-to-billions $_.issued_bills_sum   }; Align = 'right' }
    @{ Label = 'maturing'; Expression = { format-to-billions $_.maturing_bills_sum }; Align = 'right' }
    @{ Label = 'change';   Expression = { format-to-billions $_.bills_change       }; Align = 'right' }        
    @{ Label = 'issued';   Expression = { format-to-billions $_.issued_notes_sum   }; Align = 'right' }
    @{ Label = 'maturing'; Expression = { format-to-billions $_.maturing_notes_sum }; Align = 'right' }    
    @{ Label = 'change';   Expression = { format-to-billions $_.notes_change       }; Align = 'right' }
    @{ Label = 'issued';   Expression = { format-to-billions $_.issued_bonds_sum   }; Align = 'right' }
    @{ Label = 'maturing'; Expression = { format-to-billions $_.maturing_bonds_sum }; Align = 'right' }        
    @{ Label = 'change';   Expression = { format-to-billions $_.bonds_change       }; Align = 'right' }
    @{ Label = 'issued';   Expression = { format-to-billions $_.issued             }; Align = 'right' }
    @{ Label = 'maturing'; Expression = { format-to-billions $_.maturing           }; Align = 'right' }
    @{ Label = 'change';   Expression = { format-to-billions $_.change             }; Align = 'right' }    

    @{ Label = 'change_with_weekend'; Expression = { if ($_.change_with_weekend -ne $null) { format-to-billions $_.change_with_weekend } }; Align = 'right' }
    'auction'
    'auction_issuing'
    @{ Label = 'offering_amount';     Expression = { if ($_.offering_amount     -ne $null) { format-to-billions $_.offering_amount } };     Align = 'right' }
  
    
)
# ----------------------------------------------------------------------

# $table | Where-Object { 
#     if ((Get-Date $_.date).DayOfWeek -in 'Saturday', 'Sunday')
#     {
#         if ($_.change -ne 0) {
#             $_
#         }
#     }
#     else {
#         $_        
#     }
# } | Format-Table $fields

$table | Where-Object { 
    if ((Get-Date $_.date).DayOfWeek -cnotin 'Saturday', 'Sunday')
    {
        $_
    }
    elseif ($_.change -ne 0) {
        $_        
    }
    
} | Format-Table $fields
# ----------------------------------------------------------------------
# output checksum
# ----------------------------------------------------------------------
# $table | Where-Object { 
#     if ((Get-Date $_.date).DayOfWeek -cnotin 'Saturday', 'Sunday')
#     {
#         $_
#     }
#     elseif ($_.change -ne 0) {
#         $_        
#     }
    
# } | Format-Table $fields > c:\temp\out.txt

# Get-FileHash C:\Temp\out.txt
# ----------------------------------------------------------------------
exit
# ----------------------------------------------------------------------
# Example invocations

.\treasury-direct-issued-maturing.ps1 

.\treasury-direct-issued-maturing.ps1 2023-02-01 

.\treasury-direct-issued-maturing.ps1 2023-02-06 2023-03-13

# ----------------------------------------------------------------------

# $result_search = Invoke-RestMethod 'http://www.treasurydirect.gov/TA_WS/securities/search?auctionDate=2023-02-01,2023-03-01&format=json'



# $result_search    = Invoke-RestMethod ('http://www.treasurydirect.gov/TA_WS/securities/search?announcementDate={0},{1}&format=json'    -f '2023-02-13', '2023-03-01')

# $result_search    = Invoke-RestMethod ('http://www.treasurydirect.gov/TA_WS/securities/search?issueDate={0},{1}&format=json'    -f '2023-02-13', '2023-03-01')

# $result_search = Invoke-RestMethod ('http://www.treasurydirect.gov/TA_WS/securities/search?dateFieldName=issueDate&startDate=02/21/2023,02/23/2023&format=json')


# $date = '2023-02-21'

# ($result_auctioned | Group-Object issueDate | Where-Object Name -Match $date).Group[0].offeringAmount

# ($result_auctioned | Group-Object issueDate | Where-Object Name -Match $date).Group | Select-Object offeringAmount

# (($result_auctioned | Group-Object issueDate | Where-Object Name -Match $date).Group | Measure-Object offeringAmount -Sum).Sum



# (($result_auctioned | Group-Object issueDate | Where-Object Name -Match $date).Group | ForEach-Object { if ($_ -ne $null) {$_.auctionDate.Substring(0,10) }} | Sort-Object -Unique) -join ' '

# ----------------------------------------------------------------------

($result_maturing | Where-Object maturityDate -Match '2023-04-15')[-1]


$result_type_tips = $result_maturing | Where-Object type -EQ TIPS

$result_maturing[0..2] + $result_type_tips | Format-Table securityType, type


$result_maturing[0..2] + $result_type_tips | Export-Csv c:\temp\out.csv -NoTypeInformation


# ----------------------------------------------------------------------
$table = foreach ($date in (date-range $a $days))
{        
    $issued_bills_sum   = get-sum ($result_issued   | Where-Object issueDate    -Match $date | Where-Object securityType -EQ Bill)
    $issued_notes_sum   = get-sum ($result_issued   | Where-Object issueDate    -Match $date | Where-Object securityType -EQ Note)
    $issued_bonds_sum   = get-sum ($result_issued   | Where-Object issueDate    -Match $date | Where-Object securityType -EQ Bond)
    $maturing_bills_sum = get-sum ($result_maturing | Where-Object maturityDate -Match $date | Where-Object securityType -EQ Bill)
    $maturing_notes_sum = get-sum ($result_maturing | Where-Object maturityDate -Match $date | Where-Object securityType -EQ Note)
    $maturing_bonds_sum = get-sum ($result_maturing | Where-Object maturityDate -Match $date | Where-Object securityType -EQ Bond)

    $issued_sum   = get-sum ($result_issued   | Where-Object issueDate    -Match $date)
    $maturing_sum = get-sum ($result_maturing | Where-Object maturityDate -Match $date)

    [pscustomobject]@{
        date = $date
        issued_bills_sum = $issued_bills_sum;   maturing_bills_sum = $maturing_bills_sum;   bills_change = $issued_bills_sum - $maturing_bills_sum
        issued_notes_sum = $issued_notes_sum;   maturing_notes_sum = $maturing_notes_sum;   notes_change = $issued_notes_sum - $maturing_notes_sum
        issued_bonds_sum = $issued_bonds_sum;   maturing_bonds_sum = $maturing_bonds_sum;   bonds_change = $issued_bonds_sum - $maturing_bonds_sum
        
        issued = $issued_sum
        maturing = $maturing_sum
        change = $issued_sum - $maturing_sum

        auction         = if (($auction_dates        | Where-Object { $_ -match $date }) -eq $null) { '' } else { '*' }
        # auction_issuing = if (($auction_issued_dates | Where-Object { $_ -match $date }) -eq $null) { '' } else { '*' }

        auction_issuing = (($result_auctioned | Group-Object issueDate | Where-Object Name -Match $date).Group | ForEach-Object { if ($_ -ne $null) {$_.auctionDate.Substring(0,10) }} | Sort-Object -Unique) -join ' '

        offering_amount  = (($result_auctioned | Group-Object issueDate | Where-Object Name -Match $date).Group | Measure-Object offeringAmount -Sum).Sum

        # (($result_auctioned | Group-Object issueDate)[1].Group | Group-Object auctionDate | ForEach-Object Name)[0].Substring(0,10)
    }
}
# ----------------------------------------------------------------------
class Stats
{
    $issued
    $maturing
    $change

    Stats($result_issued, $result_maturing, $date, $type)
    {
        $this.issued   = get-sum ($result_issued   | ? issueDate    -Match $date | ? securityType -EQ $type)
        $this.maturing = get-sum ($result_maturing | ? maturityDate -Match $date | ? securityType -EQ $type)

        $this.change   = $this.issued - $this.maturing
    }
}


$table = foreach ($date in (date-range $a $days))
{        
    $issued_bills_sum   = get-sum ($result_issued   | Where-Object issueDate    -Match $date | Where-Object securityType -EQ Bill)
    $issued_notes_sum   = get-sum ($result_issued   | Where-Object issueDate    -Match $date | Where-Object securityType -EQ Note)
    $issued_bonds_sum   = get-sum ($result_issued   | Where-Object issueDate    -Match $date | Where-Object securityType -EQ Bond)
    $maturing_bills_sum = get-sum ($result_maturing | Where-Object maturityDate -Match $date | Where-Object securityType -EQ Bill)
    $maturing_notes_sum = get-sum ($result_maturing | Where-Object maturityDate -Match $date | Where-Object securityType -EQ Note)
    $maturing_bonds_sum = get-sum ($result_maturing | Where-Object maturityDate -Match $date | Where-Object securityType -EQ Bond)

    $issued_sum   = get-sum ($result_issued   | Where-Object issueDate    -Match $date)
    $maturing_sum = get-sum ($result_maturing | Where-Object maturityDate -Match $date)

    [pscustomobject]@{
        date = $date

        # bills = [Stats]::new($issued_bills_sum, $maturing_bills_sum)
        # notes = [Stats]::new($issued_notes_sum, $maturing_notes_sum)
        # bonds = [Stats]::new($issued_bonds_sum, $maturing_bonds_sum)



        # bills = [Stats]::new((get-sum ($result_issued   | ? issueDate    -Match $date | ? securityType -EQ Bill)), (get-sum ($result_maturing | ? maturityDate -Match $date | ? securityType -EQ Bill)))
        # notes = [Stats]::new((get-sum ($result_issued   | ? issueDate    -Match $date | ? securityType -EQ Note)), (get-sum ($result_maturing | ? maturityDate -Match $date | ? securityType -EQ Note)))
        # bonds = [Stats]::new((get-sum ($result_issued   | ? issueDate    -Match $date | ? securityType -EQ Bond)), (get-sum ($result_maturing | ? maturityDate -Match $date | ? securityType -EQ Bond)))        

        bills = [Stats]::new($result_issued, $result_maturing, $date, 'Bill')
        notes = [Stats]::new($result_issued, $result_maturing, $date, 'Note')
        bonds = [Stats]::new($result_issued, $result_maturing, $date, 'Bond')



        # issued_bills_sum = $issued_bills_sum;   maturing_bills_sum = $maturing_bills_sum;   bills_change = $issued_bills_sum - $maturing_bills_sum
        # issued_notes_sum = $issued_notes_sum;   maturing_notes_sum = $maturing_notes_sum;   notes_change = $issued_notes_sum - $maturing_notes_sum
        # issued_bonds_sum = $issued_bonds_sum;   maturing_bonds_sum = $maturing_bonds_sum;   bonds_change = $issued_bonds_sum - $maturing_bonds_sum
        
        issued = $issued_sum
        maturing = $maturing_sum
        change = $issued_sum - $maturing_sum

        auction         = if (($auction_dates        | Where-Object { $_ -match $date }) -eq $null) { '' } else { '*' }
        # auction_issuing = if (($auction_issued_dates | Where-Object { $_ -match $date }) -eq $null) { '' } else { '*' }

        auction_issuing = (($result_auctioned | Group-Object issueDate | Where-Object Name -Match $date).Group | ForEach-Object { if ($_ -ne $null) {$_.auctionDate.Substring(0,10) }} | Sort-Object -Unique) -join ' '

        offering_amount  = (($result_auctioned | Group-Object issueDate | Where-Object Name -Match $date).Group | Measure-Object offeringAmount -Sum).Sum

        # (($result_auctioned | Group-Object issueDate)[1].Group | Group-Object auctionDate | ForEach-Object Name)[0].Substring(0,10)
    }
}
# ----------------------------------------------------------------------
$table = foreach ($date in (date-range $a $days))
{        
    [pscustomobject]@{
        date = $date

        bills = [Stats]::new($result_issued, $result_maturing, $date, 'Bill')
        notes = [Stats]::new($result_issued, $result_maturing, $date, 'Note')
        bonds = [Stats]::new($result_issued, $result_maturing, $date, 'Bond')
        
        issued   = get-sum ($result_issued   | Where-Object issueDate    -Match $date)
        maturing = get-sum ($result_maturing | Where-Object maturityDate -Match $date)
        change   = $issued_sum - $maturing_sum

        auction         = if (($auction_dates        | Where-Object { $_ -match $date }) -eq $null) { '' } else { '*' }
        # auction_issuing = if (($auction_issued_dates | Where-Object { $_ -match $date }) -eq $null) { '' } else { '*' }

        auction_issuing  = (($result_auctioned | Group-Object issueDate | Where-Object Name -Match $date).Group | ForEach-Object { if ($_ -ne $null) {$_.auctionDate.Substring(0,10) }} | Sort-Object -Unique) -join ' '

        offering_amount  = (($result_auctioned | Group-Object issueDate | Where-Object Name -Match $date).Group | Measure-Object offeringAmount -Sum).Sum

        # (($result_auctioned | Group-Object issueDate)[1].Group | Group-Object auctionDate | ForEach-Object Name)[0].Substring(0,10)
    }
}
# ----------------------------------------------------------------------
$result_issued[0]