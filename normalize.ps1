param (
    [string] $SourceDir,
    [string] $TargetDir,
    [bool] $Verbose = $False,
    [bool] $Mock = $False
)

# file types to treat as music.
$MusicExtensions  = @(".wma", ".mp3", ".m4a")

# Future: files to treat as video
$VideoExtensions = @(".wmv", ".mpg", ".avi", "mp4", "m4v")

function CountExtensions {
    param (
        $SourceFolder
    )

    $ExtensnionMap = @{}

    $TotalSongs = 0
    $TotalFiles = 0

    $AllFiles = Get-ChildItem $SourceFolder -Recurse -File

    foreach ($File in $AllFiles) {
        $ExtensnionMap[$File.Extension] = 1 + $ExtensnionMap[$File.Extension]
    }

    foreach ($Key in $ExtensnionMap.Keys) {
        "$($Key): $($ExtensnionMap[$Key])"

        if ($MusicExtensions.Contains($Key)) {
            $TotalSongs += $ExtensnionMap[$Key]
        }
        $TotalFiles += $ExtensnionMap[$Key]

    }

    "Total Songs found: $TotalSongs"
    "Total Files found: $TotalFiles"
}

Function Remove-InvalidFileNameChars([String]$Name) {
    $Result = $Name # [RegEx]::Replace($Name, "["+$([RegEx]::Escape('|<>!'))+"]", '')
    # $Result = $Result.Replace('!', "_")
    $Result = $Result.Replace('"', "_")
    # $Result = $Result.Replace("'", "_")
    $Result = $Result.Replace(':', "")
    $Result = $Result.Replace('*', "_")
    $Result = $Result.Replace('/', "-")
    $Result = $Result.Replace("\", "-")
    $Result = $Result.Replace('?', "")
    return $Result
}

# Returns a map of property names to identifiers.
function GetShellAttributes($NamespaceObject) {
    $PropertyMap = [Ordered]@{}

    for ($i = 0; $i -lt 256; $i++) {
        # Passing null retrieves the name of the property for the folder.
        $PropertyMap[$NamespaceObject.GetDetailsOf($null, $i)] = $i
    }

    return $PropertyMap
}

# precondition: all songs in a batch should be in the same folder.
function ExtractSongsFromFolder($Songs, [string] $DestinationFolder) {

    # output structure is 
    #   [artist]\[album]\[2-digit track number] [track name].[existing extension]
    #
    # We'll try to get the data from the metadata, with the following fallback behavior:
    #   If album or artist is unknown, try to guess from source folder
    #       If that fails, use 'unknown'
    #   If track number is unknown, exclude.
    #   If track name is unknown, prefer existing filename

    $shell =  New-Object -ComObject Shell.Application

    $SongSourceFolder = $shell.NameSpace("$($Songs[0].Directory)")
    $PropertyMap = GetShellAttributes($SongSourceFolder)

    # Show the property map (diagnostic)
    # ($PropertyMap.Keys | foreach { "$_ $($PropertyMap[$_])" }) -join " | "

    foreach ($Song in $Songs)
    {
        $SongSourceFile = $SongSourceFolder.parsename($Song.Name)
        $Title = $SongSourceFolder.GetDetailsOf($SongSourceFile, $PropertyMap["Title"])
        $Artist = $SongSourceFolder.GetDetailsOf($SongSourceFile, $PropertyMap["Album Artist"])
        $Album = $SongSourceFolder.GetDetailsOf($SongSourceFile, $PropertyMap["Album"])
        $Track = $SongSourceFolder.GetDetailsOf($SongSourceFile, $PropertyMap["#"])
        $Size = $Song.Length # use the file property instead of shell to avoid string formatting issues.
        $FileName = $SongSourceFolder.GetDetailsOf($SongSourceFile, $PropertyMap["Name"])

        # normalize track as 2-digit number, or delete if it can't parse
        if ($Track -ne "" -and [int]$Track -ne 0)
        {
            $Track = ([int]$Track).ToString('00')
        }

        # Choose the filename based on available informtion
        if ($Title -eq "") {
            $ResultFileName = $FileName
        } elseif ($Track -eq "") {
            $ResultFileName = "$Title$($Song.Extension)"
        } else {
            $ResultFileName = "$Track $Title$($Song.Extension)"
        }

        if ($Artist -eq "") {
            $Artist = "Unknown"
        }

        if ($Album -eq "") {
            $Album = "Unknown"
        }

        $Artist = Remove-InvalidFileNameChars($Artist)
        $ResultFileName = Remove-InvalidFileNameChars($ResultFileName)
        $Album = Remove-InvalidFileNameChars($Album)

        $ResultDirectory = "$DestinationFolder\$Artist\$Album"
        $ResultFullPath = "$ResultDirectory\$ResultFileName"

        if ($Verbose) {
            "Found Song: $Artist, $Album , $Track $Title "
            "   $FileName : $Size bytes)"
            "   Source: $($Song.Directory)\$($Song.Name)"
            "   Destination: $ResultFullPath"
        }

        $exists = Test-Path $ResultFullPath

        if ($Size -eq 0) {
            if ($exists) {
                "info: Skipping file, no data, (already exists): ($($Song.Directory)\$($Song.Name))"
            } else {
                "Error: Skipping file, no data, not copied ($($Song.Directory)\$($Song.Name))"
            }
            continue
        } elseif ($Size -lt 128000) {
            if ($exists) {
                "info: Small file, suspect data (already exists): ($($Song.Directory)\$($Song.Name))"
            } else {
                "Warning: Small file, suspect data: ($($Song.Directory)\$($Song.Name))"
            }
        }

        if ($exists) {
            # Get the target size, see if it mathes. If so, only log in verbose mode.
            $ExistingFile = Get-Item  $ResultFullPath
            $ExsitingSize = $ExistingFile.Length
            if ($ExsitingSize -eq $Size -and $Verbose)
            {
                "Info: Destination exists, same size: ($($Song.Directory)\$($Song.Name) not copied to $ResultFullPath"
            }

            if ($ExsitingSize -ne $Size) {
                "Error: Destination exists ($($Song.Directory)\$($Song.Name) not copied to $ResultFullPath"
            }
            continue
        }

        if ($Mock -ne $True)
        {
            # Create the destination if needed.
            # (Future: Doesn't create parent directories!)

            $ArtistFolder = "$DestinationFolder\$Artist"
            $AlbumFolder = "$DestinationFolder\$Artist\$Album"
            
            if (!(Test-Path -path $DestinationFolder)) {$null = New-Item $DestinationFolder -Type Directory}
            if (!(Test-Path -path $ArtistFolder)) {$null = New-Item $ArtistFolder -Type Directory}
            if (!(Test-Path -path $AlbumFolder)) {$null = New-Item $AlbumFolder -Type Directory}

            Copy-Item -LiteralPath "$($Song.Directory)\$($Song.Name)" -Destination $ResultFullPath

            # Also preserve any album art / metadata. Typically, we'll be mapping from one source 
            # folder to one destination folder. As long as that's the case, copy non-music files over
            # to each destination folder. When copying, if there's a conflict with an existing item, 
            # just log it.

            $SourceDirFiles = Get-ChildItem -Force -File $Song.Directory
            foreach ($ExtraFile in $SourceDirFiles) {
                $ExtraDestination = "$AlbumFolder\$($ExtraFile.Name)"
                if ($MusicExtensions.Contains($ExtraFile.Extension) -eq $True)
                {
                    continue
                }
                $exists = Test-Path $ExtraDestination

                if ($exists) {
                    continue
                }
                if ($ExtraFile.Length -eq 0) {
                    continue
                }

                if ($Verbose) {
                    "Copying related file $ExtraDestination"
                }


                Copy-Item  -LiteralPath "$($ExtraFile.Directory)\$($ExtraFile.Name)" -Destination $ExtraDestination
            }
        } else {
            "$($Song.Directory)\$($Song.Name) "
            "    -> $ResultFullPath"            
        }


        if ($Verbose) {
            "Copied: $($Song.Directory)\$($Song.Name) -> $ResultFullPath"
        }
    }


}

function ProcessDirectory ([string] $SourceFolder, [string] $DestinationFolder) {

    if ($Verbose) {
        "Processing $SourceFolder"
    }

    # find all sub-folders
    $Folders = Get-ChildItem $SourceFolder -Directory -Force

    $Songs = [System.Collections.ArrayList]@()
    foreach ($ext in $MusicExtensions) {
        $MoreSongs = Get-ChildItem -Force -File $SourceFolder -Filter "*$ext" 
        $Songs = $Songs + $MoreSongs
    }

    $HasSongs = ($Songs.Length -ne 0)

    if ($HasSongs) {
        ExtractSongsFromFolder -Songs $Songs -DestinationFolder $DestinationFolder
    }

    foreach ($Folder in $Folders)
    {
        ProcessDirectory -SourceFolder "$Folder" -DestinationFolder $DestinationFolder
    }

}


# First give a quick summary of what's in the folder.
CountExtensions($SourceDir)

# Now proces the files
ProcessDirectory -SourceFolder $SourceDir -DestinationFolder $TargetDir


# Test cases
# $Tests = { "Let's Get It Started", ' "Weird Al" Yankovic',  "04\17\2017 All options are on the table in North Korea." }
# foreach ($Test in $Tests) {
#     Remove-InvalidFileNameChars($Test)
# }

# $shell = new-object -com shell.application;  
# Foreach ($song in $songs) {  
#   $shellfolder = $shell.namespace ($song.Directory);  
#   $shellfile = $shellfolder.parsename ($song);   
#   $title = $shell.namespace ($song.Directory).getdetailsof($shellfile,21); 
# }


# For reference... the script reads these dynamically. Might not be constant.
# Name 0 | Size 1 | Item type 2 | Date modified 3 | Date created 4 | Date accessed 5 | Attributes 6 |
# Offline status 7 | Availability 8 | Perceived type 9 | Owner 10 | Kind 11 | Date taken 12 | 
# Contributing artists 13 | Album 14 | Year 15 | Genre 16 | Conductors 17 | Tags 18 | Rating 19 | 
# Authors 20 | Title 21 | Subject 22 | Categories 23 | Comments 24 | Copyright 25 | # 26 | Length 27 | 
# Bit rate 28 | Protected 29 | Camera model 30 | Dimensions 31 | Camera maker 32 | Company 33 | File description 34 | 
# Masters keywords 36 |  218 | Program name 42 | Duration 43 | Is online 44 | Is recurring 45 | Location 46 | 
# Optional attendee addresses 47 | Optional attendees 48 | Organizer address 49 | Organizer name 50 | Reminder time 51 | 
# Required attendee addresses 52 | Required attendees 53 | Resources 54 | Meeting status 55 | Free/busy status 56 | 
# Total size 57 | Account name 58 | Task status 60 | Computer 61 | Anniversary 62 | Assistant's name 63 | Assistant's phone 64 | 
# Birthday 65 | Business address 66 | Business city 67 | Business country/region 68 | Business P.O. box 69 | Business postal code 70 | 
# Business state or province 71 | Business street 72 | Business fax 73 | Business home page 74 | Business phone 75 | 
# Callback number 76 | Car phone 77 | Children 78 | Company main phone 79 | Department 80 | E-mail address 81 | E-mail2 82 | 
# E-mail3 83 | E-mail list 84 | E-mail display name 85 | File as 86 | First name 87 | Full name 88 | Gender 89 | 
# Given name 90 | Hobbies 91 | Home address 92 | Home city 93 | Home country/region 94 | Home P.O. box 95 | 
# Home postal code 96 | Home state or province 97 | Home street 98 | Home fax 99 | Home phone 100 | IM addresses 101 | 
# Initials 102 | Job title 103 | Label 104 | Last name 105 | Mailing address 106 | Middle name 107 | Cell phone 108 | 
# Nickname 109 | Office location 110 | Other address 111 | Other city 112 | Other country/region 113 | Other P.O. box 114 | 
# Other postal code 115 | Other state or province 116 | Other street 117 | Pager 118 | Personal title 119 | 
# City 120 | Country/region 121 | P.O. box 122 | Postal code 123 | State or province 124 | Street 125 | Primary e-mail 126 | 
# Primary phone 127 | Profession 128 | Spouse/Partner 129 | Suffix 130 | TTY/TTD phone 131 | Telex 132 | Webpage 133 | 
# Content status 134 | Content type 135 | Date acquired 136 | Date archived 137 | Date completed 138 | 
# Device category 139 | Connected 140 | Discovery method 141 | Friendly name 142 | Local computer 143 | Manufacturer 144 | 
# Model 145 | Paired 146 | Classification 147 | Status 149 | Client ID 150 | Contributors 151 | Content created 152 | 
# Last printed 153 | Date last saved 154 | Division 155 | Document ID 156 | Pages 157 | Slides 158 | Total editing time 159 | 
# Word count 160 | Due date 161 | End date 162 | File count 163 | File extension 164 | Filename 165 | File version 166 | 
# Flag color 167 | Flag status 168 | Space free 169 | Group 172 | Sharing type 173 | Bit depth 174 | Horizontal resolution 175 | 
# Width 176 | Vertical resolution 177 | Height 178 | Importance 179 | Is attachment 180 | Is deleted 181 | Encryption status 182 | 
# Has flag 183 | Is completed 184 | Incomplete 185 | Read status 186 | Shared 187 | Creators 188 | Date 189 | Folder name 190 | 
# File location 191 | Folder 192 | Participants 193 | Path 194 | By location 195 | Type 196 | Contact names 197 | Entry type 198 |
# Language 199 | Date visited 200 | Description 201 | Link status 202 | Link target 203 | URL 204 | Media created 208 | Date released 209 | 
# ...
