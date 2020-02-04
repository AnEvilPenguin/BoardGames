Function newgame ($Game) {
    $BGG = "https://www.boardgamegeek.com/boardgame/$($Game.id)"
    $Check = ((((Invoke-RestMethod $BGG).Split([environment]::NewLine) | Select-String -SimpleMatch 'GEEK.geekitemPreload').tostring()).trimend(';') -split ' = ')
    If ($Check.Count -gt 2) {$Object = $Check[1..($Check.count - 1)] -join ' - ' | ConvertFrom-Json}
    Else {$Object = $Check[1] | ConvertFrom-Json}

    $Property = [ordered]@{
        Name = $Game.name.replace('??? ','')
        ID = $Game.id
        OwnedBy = $Game.OwnedBy
        WantsToPlay = $Game.wantstoplay
        ReleaseYear = $Game.Released
        MinPlayers = $Object.item.minplayers
        MaxPlayers = $Object.item.maxplayers
        Playtime = "$($Object.item.minplaytime) to $($Object.item.maxplaytime) mins"
        Rating = "$([math]::Round($Object.item.stats.average,2)) \ 10"
        Difficulty = "$([math]::Round($Object.item.stats.avgweight,2)) \ 5"
        BoardGameGeek = $BGG
        }

    New-Object -TypeName PSObject -Property $Property
}

Function bggapisearch ($Term) {
    Try {
        Invoke-RestMethod -Uri ('https://www.boardgamegeek.com/xmlapi2/search?type=boardgame&query=' + $Term.replace(' ','+').tolower()) -ErrorAction Stop
    }
    Catch{
        If ($_.exception -eq 'The remote server returned an error: (429) Too Many Requests.') {
            Start-Sleep -Seconds 2
            Invoke-RestMethod -Uri ($URL + $Search + $Query + $Object.name.replace('??? ','').replace(' ','+').tolower())
        }
        else {
            Throw $_
        }
    }
}

$CSV = Import-Csv -path .\games.csv
$Games = $CSV | Where-Object {[string]::IsNullOrEmpty($_.id)}


$Output = @()
$Output += Foreach ($Object in $Games) {
    $Result = $null
    $Result = bggapisearch ($Object.name.replace('??? ',''))

    If ($Result.items.total -eq 0){
        $Result = bggapisearch ($Object.name.replace('??? ','').Split('(')[0].trim()) #Try and deal with comments
    }

    If ($Result.items.total -eq 0){
        $Result = bggapisearch ($Object.name.replace('??? ','').Split('-')[0].trim()) #Try and deal with comments
    }

    If ($Result.items.total -eq 0){
        $Result = bggapisearch ($Object.name.replace('??? ','').Split(':')[0].trim()) #Try and deal with comments. We're grasping at straws at this point.
    }

    If ($Result.items.total -gt 1) {
        $MultipleResults = Foreach ($Item in $Result.items.item) {
            $Properties = @{
                ID = $Item.id
                Name = $Item.name.value
                Released = $Item.yearpublished.value
            }

            If ($Object.wantstoplay) {
                $Properties.Add('WantsToPlay',$Object.WantsToPlay)
            }

            If ($Object.OwnedBy) {
                $Properties.Add('OwnedBy',$Object.OwnedBy)
            }

            New-Object -TypeName psobject -Property $Properties
        }
        If ((($MultipleResults | Where-Object {$_.Name -eq $Object.name}) | Measure-Object).Count -eq 1){
            newgame ($MultipleResults | Where-Object {$_.Name -eq $Object.name})
        }
        elseif ((($MultipleResults | Where-Object {($_.Name -eq $Object.name) -and ($_.Released -eq $Object.ReleaseYear)}) | Measure-Object).Count -eq 1) {
            newgame ($MultipleResults | Where-Object {($_.Name -eq $Object.name) -and ($_.Released -eq $Object.ReleaseYear)})
        }
        elseif ($Object.ReleaseYear) {
            $Selected = $MultipleResults | Where-Object {($_.Released -eq $Object.ReleaseYear)} | Out-GridView -PassThru -Title "Which game(s) match $($Object.name)? Hint: You can visit https://www.boardgamegeek.com/boardgame/<ID>"
            $Selected | ForEach-Object {newgame $_}
        }
        else {
            $Selected = $MultipleResults | Out-GridView -PassThru -Title "Which game(s) match $($Object.name)? Hint: You can visit https://www.boardgamegeek.com/boardgame/<ID>"
            $Selected | ForEach-Object {newgame $_}
        }
    }
    elseif ($Result.items.total -lt 1) {
        Write-Warning "No results found for $($Object.name)"
        New-Object -TypeName PSObject -Property ([ordered]@{
            Name = $Object.name
            ID = $null
            OwnedBy = $Object.Ownedby
            WantsToPlay = $Object.WantstoPlay
            ReleaseYear = $Object.ReleaseYear
            MinPlayers = $null
            MaxPlayers = $null
            Playtime = $null
            Rating = $null
            Difficulty = $null
            BoardGameGeek = $null
            })
    }
    Else {
        newgame (New-Object -TypeName psobject -Property @{
            ID = $Result.items.item.id
            Name = $Result.items.item.name.value
            Released = $Result.items.item.yearpublished.value
            })
        }
        
    #Start-Sleep -Seconds 1
}

$Output += ($CSV | Where-Object {!([string]::IsNullOrEmpty($_.id))})
$Output | Sort-Object -Property Name -Unique | Select-Object -Property Name,ID,OwnedBy,wantstoplay,ReleaseYear,MinPlayers,MaxPlayers,Playtime,Rating,Difficulty,BoardGameGeek | Export-Csv -Path .\games.csv -NoTypeInformation -Force