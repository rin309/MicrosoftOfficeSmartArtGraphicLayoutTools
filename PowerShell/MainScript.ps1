# Microsoft Office SmartArt Graphic Layout (*.glox) Tool
# Version 0.1.20210526
# 
# 

Add-Type -AssemblyName WindowsBase

Set-Location (Split-Path $MyInvocation.MyCommand.Path -Parent)
$Global:TemporaryDirectoryRoot = (Join-Path (Get-Location) "Temp")
$Global:ExportDirectoryRoot = (Join-Path (Get-Location) "Export")

Function Global:Mount-OpenXmlFile($XmlPath, [Switch]$DiscardOnly){
    $TemporaryDirectory = (Join-Path $TemporaryDirectoryRoot (Split-Path $XmlPath -Leaf))
    Remove-Item $TemporaryDirectory -Force -ErrorAction SilentlyContinue | Out-Null
    If ($DiscardOnly){
        Write-Verbose "DiscardOnly"
        Return
    }
    
    If (Test-Path $XmlPath){
        [System.IO.Packaging.Package]$CurrentPackage = [System.IO.Packaging.Package]::Open($XmlPath,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read)
        $CurrentPackage.GetParts() | Where-Object Name -ne ".rels" | ForEach-Object{
            New-Item -Path (Join-Path $TemporaryDirectory (Split-Path $_.Uri -Parent)) -ItemType Directory -Force | Out-Null
            $ExportedFilePath = (Join-Path $TemporaryDirectory $_.Uri)
            [System.IO.Stream]$StreamReader = $_.GetStream()
            [System.IO.FileStream]$FileWriter = [System.IO.File]::OpenWrite($ExportedFilePath)
            $StreamReader.CopyTo($FileWriter)

            $StreamReader.Dispose()
            $FileWriter.Dispose()
        }
        $CurrentPackage.Close()
    }
    Else{
        Write-Verbose "Created new diagrams items."
        New-Item -Path (Join-Path $TemporaryDirectory "diagrams") -ItemType Directory -Force | Out-Null
        Copy-Item .\Assets\layout1.xml (Join-Path $TemporaryDirectory "diagrams\layout1.xml") -Force
        Copy-Item .\Assets\layoutHeader1.xml (Join-Path $TemporaryDirectory "diagrams\layoutHeader1.xml") -Force
    }
}

Function Global:DisMount-OpenXmlFile($XmlPath, [Switch]$CopyToTemplatesDirectory){
    $ContentTypeAndPartType = Import-Csv .\Assets\ContentTypeAndPartType.csv
    $TemporaryDirectory = (Join-Path $TemporaryDirectoryRoot (Split-Path $XmlPath -Leaf))
    If (!(Test-Path $TemporaryDirectory)){
        Return
    }
    $GloxPath = (Join-Path $ExportDirectoryRoot (Split-Path $XmlPath -Leaf))
    Remove-Item $GloxPath -Force -ErrorAction SilentlyContinue | Out-Null

    New-Item -Path $ExportDirectoryRoot -ItemType Directory -Force | Out-Null
    [System.IO.Packaging.Package]$CurrentPackage = [System.IO.Packaging.Package]::Open($GloxPath,[System.IO.FileMode]::Create,[System.IO.FileAccess]::ReadWrite)

    Get-ChildItem $TemporaryDirectory -Recurse -File | Where-Object Name -ne ".rels" | Sort-Object -Descending | ForEach-Object{
        $Path = $_.FullName.Replace("$TemporaryDirectory\","")
        $Parent = Split-Path $Path -Parent
        $FileName = Split-Path $Path -Leaf
        $FileNameWithoutExtension = ([IO.Path]::GetFileNameWithoutExtension($Path))
        $Path = "/$Parent/$FileName"
        $PathWithoutExtension = "$Parent/$FileNameWithoutExtension"
        $PathWithoutExtension
        
        $CurrentType = $ContentTypeAndPartType | Where-Object {$PathWithoutExtension -Match $_.TargetPath.Replace("?","\d{1}")}
        $Part = $CurrentPackage.CreatePart($Path, $CurrentType.PartContentType, [System.IO.Packaging.CompressionOption]::Maximum)
        
        [System.IO.FileStream]$FileReader = [System.IO.File]::OpenRead($_.FullName)
        [System.IO.Stream]$StreamReader = $Part.GetStream()
        $FileReader.CopyTo($StreamReader)
        $CurrentPackage.CreateRelationship($Path, [System.IO.Packaging.TargetMode]::Internal, $CurrentType.RelationshipType) | Out-Null

        $FileReader.Dispose()
        $StreamReader.Dispose()
    }
    $CurrentPackage.Flush()
    $CurrentPackage.Dispose()

    If ($CopyToTemplatesDirectory){
        Copy-Item $GloxPath "$env:APPDATA\Microsoft\Templates\SmartArt Graphics\"
    }
}


Function Global:Clear-OpenXmlCorruptMountPoint($GloxFileName){
    $TemporaryDirectory = (Join-Path $TemporaryDirectoryRoot (Split-Path $GloxFileName -Leaf))
    If (Test-Path $TemporaryDirectory){
        Remove-Item $TemporaryDirectory -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    }
}