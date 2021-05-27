<#
    .SYNOPSIS
        Copy MSX ROM and disk images file with shortened files names.

    .DESCRIPTION
        This script copies ROM and disk image files onto a destined folder with
        shortened filenames to allow easier file listing for loading ROMs and disk images
        with an MegaFlash SD cart or SofaROM when using MSX-DOS. 
#>
param (
    # Set source folder that contains ROM and disk images
    [Parameter(Mandatory = $true)]
    [string]$Path,

    # Set destination folder
    [Paraneter(Mandatory = $true)]
    [string]$TargetDir,

    # Set file name for ROM file list text file
    [string]$ListName = 'ROMLIST.TXT',

    # Concatenate multiple disk images onto a single disk image file
    [switch]$Concatenate
)

function TrimFileName {
    param (
        [Parameter(ValueFromPipeline)]
        [string]$Filename,
        [int16]$Length = 8
    )

    process {
        $fileNum = $FileName -replace '\D+(\d+)','$1'

        # Short words if file name length greater than 8
        if(($Filename.Length -gt $Length)) {
            $file = $Filename -replace "er |er$", "r"
        } else {
            $file = $Filename
        }

        # Trim space and chars
        $shortName = $file -replace "([-' ',_'.])", ""

        if(($shortName.Length -gt $Length)) {
            if(($fileNum -match "^[\d\.]+$")) {
                $out = $shortName.Substring(0,($Length - $fileNum.Length)) + $fileNum
            } else {
                $out = $shortName.Substring(0,$Length)
            }            
        } else {
            $out = $shortName
        }

        $out
    }
}

function ShortName {
    param (
        [Parameter(ValueFromPipeline)]
        [string]$Filename
    )

    process {
        # Trim out disk number and bloated words
        $file = $Filename -replace "\Athe | disk [0-9]", ""

        $shortName = switch -wildcard ($file) {
            "Antarctic Adventure*"              { $file | SetFileName -NewName "articadv"; break }
            "Bomber Man Special*"               { $file | SetFileName -NewName "bmrmansp"; break }
            "Bomber Man*"                       { $file | SetFileName -NewName "bombrman"; break }            
            "Bubble Bobble*"                    { $file | SetFileName -NewName "bublbobl"; break } 
            "Metal Gear 2*"                     { $file | SetFileName -NewName "mgear2"; break }
            "Metal Gear*"                       { $file | SetFileName -NewName "mgear"; break }
            "Aleste Gaiden*"                    { $file | SetFileName -NewName "alesteg"; break }
            "Space Manbow*"                     { $file | SetFileName -NewName "smanbow"; break } 
            "SD Snatcher*"                      { $file | SetFileName -NewName "sdsntchr"; break }
            "Penguin Adventure*"                { $file | SetFileName -NewName "penguin"; break }
            "Salamander*"                       { $file | SetFileName -NewName "salamndr"; break }
            "Time Pilot*"                       { $file | SetFileName -NewName "tpilot"; break }
            "Super Cobra*"                      { $file | SetFileName -NewName "sprcobra"; break }
            "Super Deform Snatcher*"            { $file | SetFileName -NewName "sdsnter"; break }
            "Super Laydock*"                    { $file | SetFileName -NewName "slaydock"; break }
            "Road Fighter*"                     { $file | SetFileName -NewName "rfighter"; break }
            "Burger Time*"                      { $file | SetFileName -NewName "brgrtime"; break }
            "Arkanoid II*"                      { $file | SetFileName -NewName "arknoid2"; break }
            "Arkanoid*"                         { $file | SetFileName -NewName "arknoid"; break }
            "Jet Set Willy*"                    { $file | SetFileName -NewName "jswilly"; break }
            "King's Valley II*"                 { $file | SetFileName -NewName "kvalley2"; break }
            "King's Valley*"                    { $file | SetFileName -NewName "kvalley"; break }
            "Vampire Killer*"                   { $file | SetFileName -NewName "vamkillr"; break }
            "Night Knight*"                     { $file | SetFileName -NewName "nightknt"; break }
            "Knightmare III*"                   { $file | SetFileName -NewName "shalom"; break }
            "Knightmare II*"                    { $file | SetFileName -NewName "galious"; break }
            "Knightmare*"                       { $file | SetFileName -NewName "kntmare"; break }
            "Track & Field 2*"                  { $file | SetFileName -NewName "trknfld2"; break }
            "Track & Field 1*"                  { $file | SetFileName -NewName "trknfld1"; break }
            "Treasure of Usas*"                 { $file | SetFileName -NewName "usas"; break }
            "Konami Game Collection Special*"   { $file | SetFileName -NewName "kgcspec"; break }    
            "Konami Game Collection*"           { $file | SetFileName -NewName ("kgcvol" + ($file -replace '\D+(\d+)','$1')); break }
            "Konami*"                           { $file | SetFileName -NewName ($file -replace "^konami(''s)?",""); break }
            "MSX-DOS Tools*"                    { $file | SetFileName -NewName "dostools"; break }
            default                             { $file | SetFileName; break }
        }

        $shortName
    }
}

function SetDiskPath {
    param (
        [Parameter(ValueFromPipeline)]
        [string]$Filename
    )

    $dirName = $Filename | ShortName
    $diskNum = $Filename.Substring($Filename.Length - 3) -replace '\D+(\d+)','$1'

    $diskPath = ($dirName + "\DISK" + $diskNum.PadLeft(2,'0'))
    $diskPath
}

function ConvertToDosPath {
    param (
        [Parameter(ValueFromPipeline)]
        [string]$File
    )

    process {
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($File)
        $fileExt = [System.IO.Path]::GetExtension($File)
        $dosPath = $null
        $cleansedName = $fileName.ToLower() -replace "([-,\(\)])","" -replace "([_])"," "

        if($cleansedName -match 'disk [0-9]') {
            $dosPath = $cleansedName | SetDiskPath
        } else {
            $dosPath = $cleansedName | ShortName
        }

        $dosPath.ToUpper() + $fileExt.ToUpper()
    }
}

function ConvertFileNameToTitle {
    # Convert filename to game title
    param (
        [Parameter(ValueFromPipeline)]
        [string]$File        
    )

    process {
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($File)
        $title = $fileName -replace '_',' '
        $title = $title -replace ' - ',': '

        $title.ToUpper()
    }
}

function SetFileName {
    # Determine if rom file is a sequal or contains region tag
    # If present place at the end of the file name <e.g Arkanoid 2 (Japan) >> arknoi2j>
    param (
        [Parameter(ValueFromPipeline)]        
        [string]$OrigName,
        [string]$NewName = $null
    )

    process {
        $tag = switch -wildcard ($OrigName) {
            "*japan*"               {"j"; break}
            "*english*"             {"en"; break}
            "*spanish*"             {"es"; break}
            "*usa*"                 {"us"; break}
            "*europe*"              {"eu"; break}
            "*demo*"                {"dm"; break}
            "*side a*"              {"a"; break}
            "*side b*"              {"b"; break}
            "*beta*"                {"b"; break}
            "*enhanced version*"    {"e"; break}
            default                 {""; break}
        }

        if ($NewName) {
            $newName = $NewName | TrimFileName -Length (8 - $tag.Length)
        } else {
            $newName = $OrigName | TrimFileName -Length (8 - $tag.Length)
        }
        
        $out = $newName + $tag
        $out
    }
}

function SetSubFolder {
    # Determine media sub-folder based on file extension
    param (
        [Parameter(ValueFromPipeline)]
        [String]$FileName
    )

    process {
        $fileExt = [System.IO.Path]::GetExtension($FileName)
        $subFolder = switch -wildcard ($fileExt) {
            "*rom"      {"roms"; break}
            "*dsk"      {"disks"; break}
            "*img"      {"disks"; break}
            "*fd*"      {"disks"; break}
            "*asd"      {"disks"; break}
            "*ips"      {"patches"; break}
            "*cas"      {"tapes"; break}
            default     {"rom"; break}
        }

        $subFolder.ToUpper()
    }
}

function CatDisks {
    # Concatenate disk images
    param (
        [string]$DiskPath
    )

    $folders = Get-ChildItem -Path $DiskPath -Directory

    foreach ($d in $folders) {
        Push-Location -Path $d.fullName
        $fileName = $d.Name + ".DSK"
        Write-Host "Concatenate disk images from: $d"
        Get-Content -Encoding Byte DISK* | Add-Content -Encoding Byte $fileName
        Pop-Location    
    }
}

# Main area
$roms = @()
$files = Get-ChildItem -Path $Path -Recurse -File
$listFile = "$TargetDir\$ListName" 

foreach ($f in $files) {
    $name = $f.Name | ConvertFileNameToTitle
    $path = $f.fullName
    $dosName = $f.Name | ConvertToDosPath
    $subFolder = $f.Name | SetSubFolder
    $targetPath = "$TargetDir\$subFolder\$dosName"

    $fileMember = new-object PSObject -Property @{
        Name    = $name
        Path    = $path
        DOSPath = "$subFolder\$dosName"
        Target  = $targetPath
    }

    $roms += $fileMember
}

# Sort target files based on DOS target location
$sortedRoms = $roms | Sort-Object -Property DOSPath

$sortedRoms | ForEach-Object {
    $copyDir = Split-Path -Path $_.TargetPath -Parent

    if (!(Test-Path -path $copyDir)) {
        New-Item $copyDir -Type Directory
    }

    Write-Host $_.Target
    Copy-Item -Path $_.Path -Destination $_.Target -Recurse
}

# Write list of roms onto a text file
$sortedRoms | Select-Object -Property DOSPath,Name | Out-File -FilePath $listFile -Width 80 -Encoding 'ASCII'

if ($Concatenate) {
    # Concatenate multiple disk images onto single file
    CatDisks -DiskPath "$TargetDir\DISKS"
}
