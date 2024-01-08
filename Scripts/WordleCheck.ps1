[CmdletBinding( DefaultParameterSetName = 'Check' )]
param (
    [Parameter( ParameterSetName = 'Check',   Mandatory = $true,  Position = 1 )] [String]   $Word,
    [Parameter( ParameterSetName = 'Add',     Mandatory = $true                )] [Switch]   $Add        = $false,
    [Parameter( ParameterSetName = 'Add',     Mandatory = $true,  Position = 1 )] [String]   $Guess1,
    [Parameter( ParameterSetName = 'Add',     Mandatory = $false, Position = 2 )] [String]   $Guess2,
    [Parameter( ParameterSetName = 'Add',     Mandatory = $false, Position = 3 )] [String]   $Guess3,
    [Parameter( ParameterSetName = 'Add',     Mandatory = $false, Position = 4 )] [String]   $Guess4,
    [Parameter( ParameterSetName = 'Add',     Mandatory = $false, Position = 5 )] [String]   $Guess5,
    [Parameter( ParameterSetName = 'Add',     Mandatory = $false, Position = 6 )] [String]   $Guess6,
    [Parameter( ParameterSetName = 'Add',     Mandatory = $false               )] [String]   $Solution,
    [Parameter( ParameterSetName = 'Add',     Mandatory = $false               )] [Switch]   $Force      = $false,
    [Parameter( ParameterSetName = 'Add',     Mandatory = $false               )] [DateTime]
    [Parameter( ParameterSetName = 'Results', Mandatory = $false               )] [DateTime] $Date,
    [Parameter( ParameterSetName = 'List',    Mandatory = $true                )] [Switch]   $List       = $false,
    [Parameter( ParameterSetName = 'Results', Mandatory = $true                )] [Switch]   $Results    = $false,
    [Parameter( ParameterSetName = 'Results', Mandatory = $false               )] [Switch]   $Share      = $false,
    [Parameter( ParameterSetName = 'Stats',   Mandatory = $true                )] [Switch]   $Statistics = $false
)

function Add-Word {
    param (
        [Parameter( Mandatory = $true  )] [string] $Date,
        [Parameter( Mandatory = $false )] [switch] $Override = $false
    )

    if ( $Override ) {
        $wordListNew = $wordList | Where-Object Date -ne $Date
        Write-Output 'Overwriting with the following:'
    } else {
        $wordListNew = $wordList
        Write-Output 'Adding new word:'
    }

    $guesses = ( @( $Guess1 , $Guess2 , $Guess3 , $Guess4 , $Guess5 , $Guess6 ) -ne '' )
    
    # If no solution is specified, assume the last guess is the solution
    if ( [String]::IsNullOrEmpty( $Solution) ) {
        $Solution = $guesses[-1]
    }

    $newWord = [PSCustomObject] @{ 'Date' =  $Date ; 'Solution' = $Solution.ToUpper() ; 'Guesses' = $guesses.ToUpper() -join ' ' }
    $newWord | Format-Table

    $wordListNew = ( $wordListNew + $newWord | Sort-Object -Property Date | ConvertTo-Csv -UseQuotes AsNeeded | Select-Object -Skip 1 ) -join [Environment]::NewLine
    $wordList    = ( $wordList                                            | ConvertTo-Csv -UseQuotes AsNeeded | Select-Object -Skip 1 ) -join [Environment]::NewLine

    # Open script file
    $file = Get-Content -Path $PSCommandPath -Raw

    # Overwrite script file with updated word list
    Set-Content -Path $PSCommandPath -Value $file.Replace( $wordList, $wordListNew ) -NoNewline 
}

$wordList = @'
2024-01-06,CABLE,CABLE
'@ | ConvertFrom-Csv -Header 'Date' , 'Solution' , 'Guesses'

# Checks to see if the word has been a solution or a starting word before
if ( $PSCmdlet.ParameterSetName -eq 'Check' ) {
    $match = $wordList | Where-Object Solution -eq $Word
    if ( $match.Count -gt 0 ) {
        Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Red )$( $Word.ToUpper() ) was the solution on the following date(s):$( $PSStyle.Reset )"
        ( $match | Select-Object @{ Name = 'Date' ; Expression = { Get-Date -Date $_.Date -Format 'MMMM d, yyyy' } } ).Date
        Write-Output ''
    } else {
        Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )$( $Word.ToUpper() ) has not yet been a solution$( $PSStyle.Reset )"
    }

    $match = $wordList | Select-Object Date , @{ Name = 'FirstWord' ; Expression = { ( $_.Guesses -split ' ' )[0] } } | Where-Object FirstWord -eq $Word
    if ( $match.Count -gt 0 ) {
        Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Red )$( $Word.ToUpper() ) was used as a starting word on the following date(s):$( $PSStyle.Reset )"
        ( $match | Select-Object @{ Name = 'Date' ; Expression = { Get-Date -Date $_.Date -Format 'MMMM d, yyyy' } } ).Date
        Write-Output ''
    } else {
        Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )$( $Word.ToUpper() ) has not yet been used as a starting word$( $PSStyle.Reset + [Environment]::NewLine )"
    }
}

# Displays the full word list
if ( $PSCmdlet.ParameterSetName -eq 'List' ) {
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )Days played:  $( $PSStyle.Reset )$( $wordList.Count )"
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )Words in database: $( $PSStyle.Reset )$( ( $wordList.FirstWord + $wordList.Solution | Select-Object -Unique ).Count )"
    $wordList | Format-Table
}

# Displays statistics
if ( $PSCmdlet.ParameterSetName -eq 'Stats' ) {
    $wordList = $wordList | Select-Object Date , Solution , @{ Name = 'FirstWord' ; Expression = { ( $_.Guesses -split ' ' )[0] } } , @{ Name = 'Guesses' ; Expression = { ( $_.Guesses -split ' ' ) } }

    # Shows number of days played
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )Days played:  $( $PSStyle.Reset + $wordList.Count )"
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )Days won:     $( $PSStyle.Reset + ( $wordList | Where-Object { $_.Solution -eq $_.Guesses[-1] } ).Count )"
    Write-Output ''

    # Shows guess distribution
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )Guess Distribution$( $PSStyle.Reset )"
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )1:$( $PSStyle.Reset ) $( ( $wordList | Where-Object { $_.Guesses.Length -eq 1 } ).Count )"
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )2:$( $PSStyle.Reset ) $( ( $wordList | Where-Object { $_.Guesses.Length -eq 2 } ).Count )"
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )3:$( $PSStyle.Reset ) $( ( $wordList | Where-Object { $_.Guesses.Length -eq 3 } ).Count )"
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )4:$( $PSStyle.Reset ) $( ( $wordList | Where-Object { $_.Guesses.Length -eq 4 } ).Count )"
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )5:$( $PSStyle.Reset ) $( ( $wordList | Where-Object { $_.Guesses.Length -eq 5 } ).Count )"
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )6:$( $PSStyle.Reset ) $( ( $wordList | Where-Object { $_.Guesses.Length -eq 6 } ).Count )"
    Write-Output ''

    # Displays most used starting words
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )Most Used Starting Words:$( $PSStyle.Reset )"
    $wordList.FirstWord | Group-Object | Where-Object Count -gt 1 | Group-Object -Property Count | Sort-Object Name -Descending | ForEach-Object {
        Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green + $_.Name ): $( $PSStyle.Reset + ( $_.Group.Name -join ' ' -split "(.{96})" -ne '' -join [Environment]::NewLine + '   ' ) )"
    }
    Write-Output ''

    # Displays most used overall words
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )Most Used Words (Overall):$( $PSStyle.Reset )"
    $wordList.Guesses | Group-Object | Where-Object Count -gt 1 | Group-Object -Property Count | Sort-Object Name -Descending | ForEach-Object {
        Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green + $_.Name ): $( $PSStyle.Reset + ( $_.Group.Name -join ' ' -split "(.{96})" -ne '' -join [Environment]::NewLine + '   ' ) )"
    }
    Write-Output ''

    # Displays words guessed that have also been a solution
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )Words that have been both starting word and a solution:$( $PSStyle.Reset )"
    Compare-Object $wordList.Solution $wordList.FirstWord -IncludeEqual | Where-Object SideIndicator -eq '==' | Select-Object @{ Name = 'Word' ; Expression = { $_.InputObject } } , @{ Name = 'Date(s) of Solution' ; Expression = { ( $wordList | Where-Object Solution -eq $_.InputObject ).Date -join ', ' } } , @{ Name = 'Date(s) Guessed' ; Expression = { ( $wordList | Where-Object FirstWord -eq $_.InputObject ).Date -join ', ' } }
}

# Displays results for any given day
if ( $PSCmdlet.ParameterSetName -eq 'Results' ) {
    # Set date to today if not specified
    if ( !$Date ) {
        $Date = Get-Date
        $dateText = "today ($( Get-Date -Format "MMMM d" ))"
    } else {
        $dateText = Get-Date -Date $Date -Format "MMMM d, yyyy"
    }

    # Format date for JSON
    $dateJson = Get-Date -Date $Date -Format 'yyyy-MM-dd'

    #Checks if there is any result for the specified day
    if ( ( $WordList | Where-Object Date -eq $dateJson ).Count -lt 1 ) {
        if ( $dateText -like 'today*' ) {
            Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Red )You have not yet played Wordle $( $dateText )!"
        } else {
            Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Red )You did not play Wordle on $( $dateText )"
        }
        Write-Output $PSStyle.Reset
        exit
    }

    $guesses  = ( $WordList | Where-Object Date -eq $dateJson ).Guesses -split ' '
    $solution = ( $WordList | Where-Object Date -eq $dateJson ).Solution
    #$wordFormatted  = ''
    $wordFormatted  = $PSStyle.Bold + $PSStyle.Foreground.White
    $wordFormattedS = ''
    Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Green )Results from $( $dateText ):$( $PSStyle.Reset )"
    foreach ( $word in $guesses ) {
        for ( $i = 0 ; $i -lt 5 ; $i++ ) {
            if ( $solution[ $i ] -eq $word[ $i ] ) {
                #$wordFormatted  += $PSStyle.Foreground.FromRgb( '0x538D4E' )
                $wordFormatted  += $PSStyle.Background.FromRgb( '0x538D4E' )
                $wordFormattedS += '🟩'
            } elseif ( $solution.Contains( $word[ $i ] ) ) {
                #$wordFormatted  += $PSStyle.Foreground.FromRgb( '0xB59F3B' )
                $wordFormatted  += $PSStyle.Background.FromRgb( '0xB59F3B' )
                $wordFormattedS += '🟨'
            } else {
                # $wordFormatted  += $PSStyle.Foreground.FromRgb( '0xA0A0A6' )
                $wordFormatted  += $PSStyle.Background.FromRgb( '0x3A3A3C' )
                $wordFormattedS += '⬛'
            }
            # $wordFormatted += [char]::ConvertFromUtf32( '0x' + ( [Byte][Char]$word[ $i ] + 127279 ).ToString( 'X' ) ) + ' '
            $wordFormatted += " $( $word[ $i ] ) "
        }
        #$wordFormatted  +=  [Environment]::NewLine
        $wordFormatted  +=  $PSStyle.Reset + $PSStyle.Bold + $PSStyle.Foreground.White + [Environment]::NewLine
        $wordFormattedS += [Environment]::NewLine
    }

    if ( $Share ) {
        Write-Output ( $wordFormattedS + $PSStyle.Reset )
        ( "Results from $( $dateText ):" + [Environment]::NewLine + $wordFormattedS ) | Set-Clipboard
    } else {
        Write-Output ( $wordFormatted  + $PSStyle.Reset )
    }
}

# Adds new word to the list
if ( $PSCmdlet.ParameterSetName -eq 'Add' ) {
    # Verify that words are valid
    foreach ( $guess in @( $Guess1 , $Guess2 , $Guess3 , $Guess4 , $Guess5 , $Guess6 ) -ne '' ) {
        if ( -not ( $guess -match "^[A-Z]{5}$" ) ) {
            Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Red )$( $guess.ToUpper() ) is not a 5 letter word.$( $PSStyle.Reset ) + [Environment]::NewLine"
            exit
        }
    }

    # Verify solution is valid if specified
    if ( -not [String]::IsNullOrEmpty( $Solution) -and -not ( $Solution -match "^[A-Z]{5}$" ) ) {
        Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Red )$( $Solution.ToUpper() ) is not a 5 letter word.$( $PSStyle.Reset ) + [Environment]::NewLine"
        exit
    }
    
    # Set date to today if not specified
    if ( !$Date ) {
        $Date = Get-Date
    }

    # Format date for JSON
    $dateJson = Get-Date -Date $Date -Format 'yyyy-MM-dd'
    $dateText = Get-Date -Date $Date -Format "MMMM d, yyyy"

    # Check if entry already exists for specified date
    $match = $wordList | Where-Object Date -eq $dateJson
    if ( $match.Count -gt 0 ) {
        Write-Output "$( $PSStyle.Bold + $PSStyle.Foreground.Red )There is already an entry for $( $dateText ):$( $PSStyle.Reset )"
        $match | Format-Table

        # Overwrites the existing word if -Force is specified
        if ( $Force ) {
            Add-Word -Date $dateJson -Override
        }
    } else {
        Add-Word -Date $dateJson
    }
}
