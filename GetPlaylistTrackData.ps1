#Ideally you have Powershell 7 installed
#Be sure to fill in your own client ID and secret
  
<# 
    Function Load-Powershell_7{

    function New-OutOfProcRunspace {
        param($ProcessId)

        $connectionInfo = New-Object -TypeName System.Management.Automation.Runspaces.NamedPipeConnectionInfo -ArgumentList @($ProcessId)

        $TypeTable = [System.Management.Automation.Runspaces.TypeTable]::LoadDefaultTypeFiles()

        #$Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateOutOfProcessRunspace($connectionInfo,$Host,$TypeTable)
        $Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($connectionInfo,$Host,$TypeTable)

        $Runspace.Open()
        $Runspace
    }
    

    $Process = Start-Process PWSH -ArgumentList @("-NoExit") -PassThru -WindowStyle Hidden

    $Runspace = New-OutOfProcRunspace -ProcessId $Process.Id

    $Host.PushRunspace($Runspace)
} 
#>
     
     $tokenheader=@{
     'content-type'='application/x-www-form-urlencoded'

     }


    $Token = Invoke-RestMethod -Method Post -Uri "https://accounts.spotify.com/api/token" -Headers $tokenheader -body "grant_type=client_credentials&client_id=INSERT_CLIENT_ID&client_secret=INSERT_CLIENT_SECRET"
    $Token=$Token.access_token

     $headers = @{
    Authorization="Bearer $Token"
}

$Output=@()
$count=0
$PlaylistLink= Read-Host "Enter a public Spotify playlist link"
$string = $PlaylistLink
$startIndex = $string.LastIndexOf("/") + 1
$length = $string.IndexOf("?") - $startIndex
$PlaylistID= $string.Substring($startIndex, $length)

$Playlist=(Invoke-RestMethod -Method Get -Uri "https://api.spotify.com/v1/playlists/$PlaylistID" -Headers $headers)

$PlaylistName=$Playlist.Name

$TotalTracks= $Playlist.tracks.total
$TrackLoop=[math]::ceiling($TotalTracks / 100)

for($i=0;$i -lt $TrackLoop; $i++){

$Tracks=Invoke-RestMethod -Method Get -Uri "https://api.spotify.com/v1/playlists/$PlaylistID/tracks?limit=100&offset=$($i*100)" -Headers $headers


$TrackID=$Tracks.items.track.id

foreach($one in $TrackID){ 

$Track=Invoke-RestMethod -Method Get -Uri "https://api.spotify.com/v1/tracks/$one" -Headers $headers 

$Trackfeatures=Invoke-RestMethod -Method Get -Uri "https://api.spotify.com/v1/audio-features/$one" -Headers $headers  
    

     $Name=$Track.Name
     $Artist=$Track.artists.Name
     $Key=$Trackfeatures.key
     $Tempo=$Trackfeatures.tempo

     $ArtistID=$Track.artists.id
     $AlbumID=$track.album.id

     If($ArtistID.count -gt 1){$ArtistID=$ArtistID[0]}
     If($AlbumID.count -gt 1){$AlbumID=$AlbumID[0]}

     $AlbumGenres=(Invoke-RestMethod -Method Get -Uri "https://api.spotify.com/v1/albums/$AlbumID" -Headers $headers).genres 
     $ArtistGenres=(Invoke-RestMethod -Method Get -Uri "https://api.spotify.com/v1/artists/$ArtistID" -Headers $headers).genres 

     If([string]::IsNullOrEmpty($AlbumGenres)){$AlbumGenres="Unknown"}
     If([string]::IsNullOrEmpty($ArtistGenres)){$ArtistGenres="Unknown"}

     If($Artist.count -gt 1){ $Artist = ($Artist -join ",")}
     If($AlbumGenres.count -gt 1){$AlbumGenres = ($AlbumGenres -join ",")}
     If($ArtistGenres.count -gt 1){$ArtistGenres = ($ArtistGenres -join ",")}

     $Key=switch($Key){
    -1{"No Key Info"}
     0{"C"}
     1{"C#/Db"}
     2{"D"}
     3{"D#/Eb"}
     4{"E"}
     5{"F"}
     6{"F#/Gb"}
     7{"G"}
     8{"G#/Ab"}
     9{"A"}
     10{"A#/Bb"}
     11{"B"}
     default{"No Match Found"}
     }

$Output+= """$($Name)"",""$($Artist)"",""$($Key)"",""$($Tempo)"",""$($AlbumGenres)"",""$($Artistgenres)"""
 }
 }

# Set the CSV file path
$csvFilePath = "C:\TEMP\$($PlaylistName).csv"

If(!(test-path -path $csvfilepath)){

# Define headers
$headers = "Song Title", "Artist", "Key", "BPM", "Album Genre", "Artist Genre"

# Write the headers to the CSV file
$headers -join "," | Out-File -FilePath $csvFilePath -Encoding UTF8
}

Add-Content -Path $csvFilePath -Value $Output -Encoding UTF8

Write-Output "Exported all track data for Playlist: $($PlaylistName) to $CSVFilePath!"

#Still need to save each query as an object and then add sort by the properties (i.