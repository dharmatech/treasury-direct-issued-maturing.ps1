
Param(
    $a = (Get-Date (Get-Date   ).AddDays(-7) -Format 'yyyy-MM-dd'), 
    $b = (Get-Date (Get-Date $a).AddDays(35) -Format 'yyyy-MM-dd')
    )

Write-Host $a

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

$auction_dates        = $result_auctioned | Group-Object auctionDate | ForEach-Object Name
$auction_issued_dates = $result_auctioned | Group-Object issueDate | ForEach-Object Name

# $result_maturing | ft *

# $issued   = $result_issued   | Where-Object issueDate    -Match 2023-01-31
# $maturing = $result_maturing | Where-Object maturityDate -Match 2023-01-31

# ($issued   | Measure-Object -Property totalAccepted -Sum).Sum
# ($maturing | Measure-Object -Property totalAccepted -Sum).Sum

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

    # $issued_sum   = if ($issued   -eq $null) { 0 } else { ($issued   | Measure-Object -Property totalAccepted -Sum).Sum }
    # $maturing_sum = if ($maturing -eq $null) { 0 } else { ($maturing | Measure-Object -Property totalAccepted -Sum).Sum }

    $issued_bills_sum = get-sum $issued_bills; $maturing_bills_sum = get-sum $maturing_bills
    $issued_notes_sum = get-sum $issued_notes; $maturing_notes_sum = get-sum $maturing_notes
    $issued_bonds_sum = get-sum $issued_bonds; $maturing_bonds_sum = get-sum $maturing_bonds

    $issued_sum   = get-sum $issued
    $maturing_sum = get-sum $maturing

    $issued_security_terms = ($issued | Group-Object securityTerm | ForEach-Object Name) -join ' '

    [pscustomobject]@{
        date = $date
        issued_bills_sum = $issued_bills_sum; maturing_bills_sum = $maturing_bills_sum
        issued_notes_sum = $issued_notes_sum; maturing_notes_sum = $maturing_notes_sum
        issued_bonds_sum = $issued_bonds_sum; maturing_bonds_sum = $maturing_bonds_sum

        # issued_notes = $issued_notes
        # maturing_notes = $maturing_notes

        # issued_bonds = $issued_bonds
        # maturing_bonds = $maturing_bonds

        issued = $issued_sum
        maturing = $maturing_sum
        change = $issued_sum - $maturing_sum

        auction         = if (($auction_dates        | Where-Object { $_ -match $date }) -eq $null) { '' } else { '*' }
        # auction_issuing = if (($auction_issued_dates | Where-Object { $_ -match $date }) -eq $null) { '' } else { '*' }

        auction_issuing = (($result_auctioned | Group-Object issueDate | Where-Object Name -Match $date).Group | ForEach-Object { if ($_ -ne $null) {$_.auctionDate.Substring(0,10) }} | Sort-Object -Unique) -join ' '

        issued_security_terms = $issued_security_terms

        # (($result_auctioned | Group-Object issueDate)[1].Group | Group-Object auctionDate | ForEach-Object Name)[0].Substring(0,10)
    }
}

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


Write-Host '           BILLS                  NOTES                  BONDS                  TOTAL' -NoNewline
#           date       issued maturing change issued maturing change issued maturing change issued maturing change auction auction_issuing

$fields = @(
    'date',

    # @{ Label = 'issued_bills_sum';     Expression = { $_.issued_bills_sum.ToString('N0')     }; Align = 'right' }
    # @{ Label = 'maturing_bills_sum';   Expression = { $_.maturing_bills_sum.ToString('N0')   }; Align = 'right' }
    
    # @{ Label = 'bills i';   Expression = { ($_.issued_bills_sum   / 1000 / 1000 / 1000).ToString('N0') }; Align = 'right' }
    # @{ Label = 'bills m';   Expression = { ($_.maturing_bills_sum / 1000 / 1000 / 1000).ToString('N0') }; Align = 'right' }
    # @{ Label = 'bills c';     Expression = { format-to-billions ($_.issued_bills_sum - $_.maturing_bills_sum) }; Align = 'right' }

    # @{ Label = 'bills issued';   Expression = { ($_.issued_bills_sum   / 1000 / 1000 / 1000).ToString('N0') }; Align = 'right' }
    # @{ Label = 'maturing';   Expression = { ($_.maturing_bills_sum / 1000 / 1000 / 1000).ToString('N0') }; Align = 'right' }
    # @{ Label = 'change';     Expression = { format-to-billions ($_.issued_bills_sum - $_.maturing_bills_sum) }; Align = 'right' }    

    @{ Label = 'issued';   Expression = { ($_.issued_bills_sum   / 1000 / 1000 / 1000).ToString('N0') }; Align = 'right' }
    @{ Label = 'maturing';   Expression = { ($_.maturing_bills_sum / 1000 / 1000 / 1000).ToString('N0') }; Align = 'right' }
    @{ Label = 'change';     Expression = { format-to-billions ($_.issued_bills_sum - $_.maturing_bills_sum) }; Align = 'right' }        

    # @{ Label = 'issued_notes_sum';     Expression = { $_.issued_notes_sum.ToString('N0')     }; Align = 'right' }
    # @{ Label = 'maturing_notes_sum';   Expression = { $_.maturing_notes_sum.ToString('N0')   }; Align = 'right' }    

    # @{ Label = 'notes i';   Expression = { format-to-billions $_.issued_notes_sum     }; Align = 'right' }
    # @{ Label = 'notes m';   Expression = { format-to-billions $_.maturing_notes_sum   }; Align = 'right' }    
    # @{ Label = 'notes c';     Expression = { format-to-billions ($_.issued_notes_sum - $_.maturing_notes_sum) }; Align = 'right' }

    @{ Label = 'issued';   Expression = { format-to-billions $_.issued_notes_sum     }; Align = 'right' }
    @{ Label = 'maturing';   Expression = { format-to-billions $_.maturing_notes_sum   }; Align = 'right' }    
    @{ Label = 'change';     Expression = { format-to-billions ($_.issued_notes_sum - $_.maturing_notes_sum) }; Align = 'right' }

    # @{ Label = 'issued_bonds_sum';     Expression = { $_.issued_bonds_sum.ToString('N0')     }; Align = 'right' }
    # @{ Label = 'maturing_bonds_sum';   Expression = { $_.maturing_bonds_sum.ToString('N0')   }; Align = 'right' }    

    @{ Label = 'issued';     Expression = { format-to-billions $_.issued_bonds_sum     }; Align = 'right' }
    @{ Label = 'maturing';   Expression = { format-to-billions $_.maturing_bonds_sum   }; Align = 'right' }        
    @{ Label = 'change';     Expression = { format-to-billions ($_.issued_bonds_sum - $_.maturing_bonds_sum) }; Align = 'right' }

    # @{ Label = 'issued';   Expression = { $_.issued.ToString('N0')   }; Align = 'right' }
    # @{ Label = 'maturing'; Expression = { $_.maturing.ToString('N0') }; Align = 'right' }
    # @{ Label = 'change';   Expression = { $_.change.ToString('N0')   }; Align = 'right' }

    @{ Label = 'issued';   Expression = { format-to-billions $_.issued   }; Align = 'right' }
    @{ Label = 'maturing'; Expression = { format-to-billions $_.maturing }; Align = 'right' }
    @{ Label = 'change';   Expression = { format-to-billions $_.change   }; Align = 'right' }    

    'auction'
    'auction_issuing'
    # 'issued_security_terms'
)

$table | Format-Table $fields

# ----------------------------------------------------------------------

function modified-duration ([decimal]$coupon, [decimal]$face_value, [decimal]$frequency, [decimal]$maturity, [decimal]$yield)
{
    $price = ($coupon * $face_value / $frequency) * (1 - (1 / [math]::Pow(1 + $yield / $frequency, $maturity * $frequency))) / ($yield / $frequency) + $face_value / [math]::Pow(1 + $yield / $frequency, $maturity * $frequency)

    (($coupon * $face_value / $frequency) / [math]::Pow($yield / $frequency, 2) *
    ((1 - 1 / [math]::Pow(1 + $yield / $frequency, $maturity * $frequency))) +
    ($maturity * $frequency * ($face_value - (($coupon * $face_value / 2) / ($yield / $frequency)))   /   [math]::Pow(1 + ($yield / $frequency), $maturity * $frequency + 1))) / 
    $price / $frequency    
}

# PARAMETER      
# $coupon        $security.interestRate / 100
# $face_value    $security.pricePer100
# $frequency     2 (always assuming semi-annual)
# $maturity      $security.term (parse years out)
# $yield         $security.interestRate / 100

foreach ($security in $result_issued)
{
    if ($security.interestRate.Length -gt 0)
    {
        $term = [decimal] ($security.term -replace '-.*', '')

        $interest_rate = $security.interestRate / 100

        # @{
        #     term = $term
        #     interestRate = $security.interestRate
        #     price = $security.pricePer100
        # }
        
        # [PSCustomObject]@{
        #     modified_duration = modified-duration $interest_rate $security.pricePer100 2 $term $interest_rate
        # }

        $security | Add-Member -MemberType NoteProperty -Name modified_duration -Value (modified-duration $interest_rate $security.pricePer100 2 $term $interest_rate)
    }
}

$result_issued | ft *

$result_issued | ft cusip, issueDate, securityType, securityTerm, maturityDate, interestRate, refCpiOnIssueDate, refCpiOnDatedDate, announcementDate, auctionDate, @{ Label = 'modified_duration'; Expression = { $_.modified_duration.ToString('N') } }

# [decimal] $coupon = $security.interestRate
# [decimal] $face_value = $security.pricePer100
# [decimal] $frequency = 2
# [decimal] $maturity = 30
# [decimal] $yield = $security.interestRate / 100


exit

# ----------------------------------------------------------------------
# Example invocations

.\treasury-direct-issued-maturing.ps1 

.\treasury-direct-issued-maturing.ps1 2023-02-01 

.\treasury-direct-issued-maturing.ps1 2023-02-06 2023-03-13




$result_issued | Format-Table issueDate, securityType, interestRate