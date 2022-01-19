# Get all deck lists from a user in moxfield
# Compile all cards from those decks
# Download all the card images locally
# Color the corners of the cards

# Example use:
# .\Get-MoxfiledCardImages.ps1 -username "Shaper" -path "Z:\mtg" -color "MidnightBlue"

################################################################################
###   Input Parameters / Globals
################################################################################
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, HelpMessage="Username of a Moxfield Account.")]
    [string]$username,
    [Parameter(Mandatory=$true, HelpMessage="Path of where images are to be saved")]
    [string]$path,
    [Parameter(Mandatory=$false, HelpMessage="Color to paint the corners of the cards (Default: Black)")]
    [string]$color="Black"
)
 
$primaryList = @{} # Main hashmap that contains all cards that will be downloaded K:ScryfallID,V:Name
 
################################################################################
###   Validate Chosen Color
################################################################################
 
try {
    $chosenColor = [System.Drawing.KnownColor]::$color # returns null if can't be found
    $chosenColor.ToString() | Out-Null # A hack to fail if chosenColor is Null
 
    [int]$red = [System.Drawing.Color]::FromName($chosenColor).R
    [int]$green = [System.Drawing.Color]::FromName($chosenColor).G
    [int]$blue = [System.Drawing.Color]::FromName($chosenColor).B
    Write-Host "Chosen Color:"
    Write-Host $("`t{0} R({1}),G({2}),B({3})" -f $chosenColor, $red, $green, $blue)
    Write-Host " "
}
Catch {
    Write-Host "Chosen Color does not exist ..." -NoNewline
    Write-Host "Choose color from https://developer.mozilla.org/en-US/docs/Web/CSS/color_value ... " -NoNewline
    Write-Host "Exiting." -ForegroundColor Red
    Exit
}
 
################################################################################
###   Get decks associated with the username
################################################################################

# TODO: Add multiple page gathering. Right now maxes out at just the default 12 that are returned.

try {
    Start-Sleep -Milliseconds 500 # A polite sleep
    $usernameResponse = $(Invoke-RestMethod -ErrorVariable err "https://api.moxfield.com/v2/users/$username/decks?pageNumber=1&pageSize=12" )
}
Catch {
   
    if ($($err.ErrorRecord | ConvertFrom-Json).status -eq 404) {
        Write-Host "Username does not exist ... " -NoNewline
        Write-Host "Exiting." -ForegroundColor Red
    } else {
        Write-Host "Unknown Error ... " -NoNewline
        Write-Host $("Got status ({0}) ... " -f $($err.ErrorRecord | ConvertFrom-Json).status) -NoNewline
        Write-Host "Exiting." -ForegroundColor Red
    }
    Exit
}
 
################################################################################
###   Get all cards from the decks
################################################################################
 
Write-Host $("{0} has {1} deck(s):" -f $username, $usernameResponse.data.Length)
 
foreach ($deck in $usernameResponse.data) {
   
    Write-Host $("`t{0}" -f $($deck.name)) # `t is tab in powershell
   
    # Get the details of the deck
    Start-Sleep -Milliseconds 500 # A polite sleep
    $deckResponse = Invoke-RestMethod "https://api.moxfield.com/v2/decks/all/$($deck.publicId)"
   
    # For the cards in the mainboard
    $deckResponse.mainboard | ForEach-Object {
        $_.psobject.properties | ForEach-Object {
 
            # Add them to the primary list if they're not already in there
            # Key = Scryfall's GUID of the card
            # Value = Name of card
            if ( $primaryList.ContainsKey($_.value.card.scryfall_id) -eq $false ){
                $primaryList.Add($_.value.card.scryfall_id, $_.name)
            }
        }
    }
}
Write-Host " "
 
################################################################################
###   Figure out which cards need to be downloaded
###   &
###   Purge cards on disk that aren't in moxfield lists
################################################################################

# TODO: fix it so that it only deletes files that are .*&GUID.jpg

$filesOnDisk = Get-ChildItem -Path $path
# Foreach of them remove them from the primary list
foreach ($file in $filesOnDisk){
    $fileSplits = $($file.Name.Split("&")[1])
    $fileWithNoExtension = $($fileSplits.Substring(0,$fileSplits.Length-4))
    # If file locally is found in primary list, remove from primary list
    # i.e. we don't need to download it, we've already got it
    if ( $primaryList.ContainsKey($fileWithNoExtension) ){
        $primaryList.Remove($fileWithNoExtension)
    } else {
        # else the file is not found in the primary list
        # meaning it's no longer in a deck online
        # purge it!
 
        Write-Host $("Removing {0} ... " -f $file.Name.Split("&")[0]) -NoNewline
        Remove-Item $file
        Write-Host $("Done!") -ForegroundColor Green
    }
}
 
################################################################################
###   Get images from scryfall & paint corners
################################################################################
foreach($scryfallID in $primaryList.Keys){
 
    Write-Host $("Fetching Data for {0} ... " -f $primaryList[$scryfallID] ) -NoNewline
 
    Start-Sleep -Milliseconds 200 # Scryfall's requested sleep time
    $scryfallResponse = Invoke-RestMethod "https://api.scryfall.com/cards/$scryfallID"
 
    # Only download if it's considered a high resolution image
    if ($scryfallResponse.highres_image){
 
        # There's double faced cards, so we need to know where to get the PNG url from within the response payload
        $imageurl = ""
        if ($null -ne $scryfallResponse.image_uris){
            if ($null -ne $scryfallResponse.image_uris.png){
                $imageurl = $($scryfallResponse.image_uris.png) # it's a front only card
            } else {
                # This shouldn't ever reach here ...
                Write-Host "SERIOUS ERROR parsing DFCs" -ForegroundColor White -BackgroundColor Red
                break
            }
        } else {
            $imageurl = $($scryfallResponse.card_faces[0].image_uris.png) # it's a DFC
        }
 
        Write-Host $("Fetching Image ... ") -NoNewline
   
        # Invoke-WebRequest downloads show a progress bar ... we don't want to see that.
        # Thus, we set ProgressPreference to shutup
        $progressPreferenceHolder = $progressVariable.Value
        $ProgressPreference = 'SilentlyContinue'
        Start-Sleep -Milliseconds 200 # Scryfall's requested sleep time
        Invoke-WebRequest $imageurl -OutFile "$path\tmp_$scryfallID.png" | Out-Null
        $ProgressPreference = $progressPreferenceHolder
 
        # remove invalid filename characters from name of card if they are there and replace with '-'
        $correctedName = $($primaryList[$scryfallID].Split([IO.Path]::GetInvalidFileNameChars()) -join '-')
   
        # Adjust background color from clear to MidnightBlue
        # Magic from Stackoverflow:
        $Source = "$path\tmp_$scryfallID.png"
        $test = $path
        $base= $($correctedName+" &"+$scryfallID+".jpg")
        $basedir = $test+"\"+$base
        Add-Type -AssemblyName system.drawing
        $imageFormat = "System.Drawing.Imaging.ImageFormat" -as [type]
        $image = [drawing.image]::FromFile($Source)
        $NewImage = [System.Drawing.Bitmap]::new($Image.Width,$Image.Height)
        $NewImage.SetResolution($Image.HorizontalResolution,$Image.VerticalResolution)
        $Graphics = [System.Drawing.Graphics]::FromImage($NewImage)
        $Graphics.Clear([System.Drawing.Color]::$chosenColor)
        $Graphics.DrawImageUnscaled($image,0,0)
        $NewImage.Save($basedir,$imageFormat::Jpeg)
        $image.Dispose()
        Remove-Item "$path\tmp_$scryfallID.png"
 
        Write-Host $("Done!") -ForegroundColor Green
    } else {
        Write-Host $("No High Resolution Image ... ") -NoNewline
        Write-Host $("Skipping.") -ForegroundColor Cyan
    }
}