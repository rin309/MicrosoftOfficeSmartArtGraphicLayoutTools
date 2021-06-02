# Microsoft Office SmartArt Graphic Layout (*.glox) Tool
# Version 0.2.20210602
# 
# 

Add-Type -AssemblyName WindowsBase , Microsoft.Office.Interop.Word, System.Windows.Forms, Microsoft.VisualBasic
[Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Word") | Out-Null

Set-Location (Split-Path $MyInvocation.MyCommand.Path -Parent)
$Global:TemporaryDirectoryRoot = (Join-Path (Get-Location) "Temp")
$Global:ExportDirectoryRoot = (Join-Path (Get-Location) "Export")
$Global:WordApplicationCaption = "PowerShellスクリプトで管理されたWord"
$Global:MsoSmartGraphicsDisplayNameTable = Import-Csv ".\Assets\MsoSmartArtGraphicsTable.1041.csv"

<#
    .SYNOPSIS
    OpenXML 形式のファイルを作成するための一時ファイルを展開します

    .PARAMETER Path
    フォルダー名を相対パスか絶対パスを指定します
    相対パスで指定された場合、このスクリプトがあるフォルダーの Temp フォルダー内に保存されます

    .PARAMETER OpenXmlType
    OpenXML 形式の種類を選択します

    .PARAMETER Force
    指定された Path が存在する場合でも続行します

    .EXAMPLE
    New-TemporaryOpenXmlFile -Path test.glox -OpenXmlType Diagrams

#>
Function Global:New-TemporaryOpenXmlFile([Parameter(Mandatory)]$Path, [Parameter(Mandatory)][OpenXmlType]$OpenXmlType, [Switch]$Force){
    Enum OpenXmlType {
        Diagrams
    }

    $SourceDirectory = ".\Assets\Templates\$($OpenXmlType.ToString().ToLower())"
    If (!(Test-Path $SourceDirectory -PathType Container)){
        Write-Error "コピー元のフォルダーが見つかりませんでした: $SourceDirectory"
    }
    
    $TemporaryDirectory = Get-OpenXmlTemporaryDirectory -Path $Path
    If (Test-Path $TemporaryDirectory){
        If ($Force){
            Remove-Item $TemporaryDirectory -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Else{
            Write-Error "指定されたパスはすでに存在します。このパスを削除するか、-Force オプションを使用してください: $TemporaryDirectory"
        }
    }

    #$TemporaryDirectory = (Join-Path $TemporaryDirectory ($OpenXmlType.ToString().ToLower()))
    Write-Verbose "Created new $OpenXmlType items."
    New-Item -Path $TemporaryDirectory -ItemType Directory -Force | Out-Null
    Copy-Item $SourceDirectory -Recurse $TemporaryDirectory -Force
}

<#
    .SYNOPSIS
    OpenXML 形式のファイルを展開します

    .PARAMETER Path
    フォルダー名を相対パスか絶対パスを指定します
    相対パスで指定された場合、このスクリプトがあるフォルダーの Temp フォルダー内に保存されます

    .PARAMETER XmlPath
    OpenXML 形式のファイルのパスを指定します

    .PARAMETER Force
    指定された Path が存在する場合でも続行します

    .EXAMPLE
    Mount-OpenXmlFile -Path test.glox -XmlPath "C:\Users\Owner\Documents\SmartArt Graphics\test.glox"

#>
Function Global:Mount-OpenXmlFile([Parameter(Mandatory)]$Path, [Parameter(Mandatory)]$XmlPath, [Switch]$Force){
    If (!(Test-Path $XmlPath -PathType Leaf)){
        Write-Error "ファイルが見つかりません: $XmlPath"
    }
    
    $TemporaryDirectory = Get-OpenXmlTemporaryDirectory -Path $Path
    If (Test-Path $TemporaryDirectory){
        If ($Force){
            Remove-Item $TemporaryDirectory -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Else{
            Write-Error "指定されたパスはすでに存在します。このパスを削除するか、-Force オプションを使用してください: $TemporaryDirectory"
        }
    }
    
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

<#
    .SYNOPSIS
    展開済みの OpenXML 形式のファイルを作成するための一時ファイルに値を設定します

    .PARAMETER Path
    フォルダー名を相対パスか絶対パスを指定します
    相対パスで指定された場合、このスクリプトがあるフォルダーの Temp フォルダー内に保存されます

    .PARAMETER OpenXmlType
    OpenXML 形式の種類を選択します

    .PARAMETER Title
    OpenXmlType で Diagrams を指定したとき: SmartArt グラフィックの選択画面に表示するタイトル

    .PARAMETER Description
    OpenXmlType で Diagrams を指定したとき: SmartArt グラフィックの選択画面に表示する説明

    .PARAMETER UniqueId
    既存の項目などと重複しないようにしてください
    OpenXmlType で Diagrams を指定したとき: 内部で使用される一意の値

    .PARAMETER Category
    Priority と共に使用します
    OpenXmlType で Diagrams を指定したとき: SmartArt グラフィックの選択画面に表示するカテゴリーの種類

    .PARAMETER Priority
    Category と共に使用します
    OpenXmlType で Diagrams を指定したとき: SmartArt グラフィックの選択画面に表示する優先順位の値

    .EXAMPLE
    Set-OpenXmlFileValue -Path test.glox -OpenXmlType Diagrams -Title "テスト" -Description "これはテストです" -Category List -Priority 1000 -UniqueId "urn:contoso.com/office/officeart/2021/6/layout/test"

#>
Function Global:Set-OpenXmlFileValue([Parameter(Mandatory)]$Path, [Parameter(Mandatory)][OpenXmlType]$OpenXmlType, $Title, $Description, $UniqueId, [IgxCategory]$Category, [Int]$Priority){
    Enum OpenXmlType {
        Diagrams
    }
    Enum IgxCategory{
        List
        Process
        Cycle
        Hierarchy
        Relationship
        Matrix
        Pyramid
        Picture
        Other
    }

    $Path = Get-OpenXmlTemporaryDirectory -Path $Path

    # 読み込み
    $Layout1Path = (Join-Path $Path "$($OpenXmlType.ToString().ToLower())\layout1.xml")
    $LayoutHeader1Path = (Join-Path $Path "$($OpenXmlType.ToString().ToLower())\layoutHeader1.xml")

    If (!(Test-Path $Layout1Path -PathType Leaf)){
        Write-Error "ファイルが見つかりません: $Layout1Path"
    }
    If (!(Test-Path $LayoutHeader1Path -PathType Leaf)){
        Write-Error "ファイルが見つかりません: $LayoutHeader1Path"
    }
    
    $Layout1IsChanged = $False
    $Layout1XmlDocument = [XML](Get-Content $Layout1Path)
    $LayoutHeader1IsChanged = $False
    $LayoutHeader1XmlDocument = [XML](Get-Content $LayoutHeader1Path)

    # 変更
    If ($Title -ne $null){
        $LayoutHeader1XmlDocument.layoutDefHdr.title.val = $Title
        $LayoutHeader1IsChanged = $True
    }
    If ($Description -ne $null){
        $LayoutHeader1XmlDocument.layoutDefHdr.desc.val = $Description
        $LayoutHeader1IsChanged = $True
    }
    If ($UniqueId -ne $null){
        $Layout1XmlDocument.layoutDef.uniqueId = $UniqueId
        $LayoutHeader1XmlDocument.layoutDefHdr.uniqueId = $UniqueId
        $Layout1IsChanged = $True
        $LayoutHeader1IsChanged = $True
    }
    If (($Category -ne $null) -or ($Priority -ne 0)){
        $catLstXPath = "//*[local-name()='layoutDefHdr']/*[local-name()='catLst']"
        If ($LayoutHeader1XmlDocument.SelectNodes($catLstXPath).Count -eq 0){
            $catLstXml = $LayoutHeader1XmlDocument.CreateElement("catLst", $LayoutHeader1XmlDocument.DocumentElement.NamespaceURI)
            $LayoutHeader1XmlDocument.layoutDefHdr.AppendChild($catLstXml) | Out-Null
        }
        ElseIf ($LayoutHeader1XmlDocument.SelectNodes($catLstXPath).Count -ne 1){
            Write-Warning "このファイルには、複数の catLst が定義されています: $Path"
        }
        $catLstXml = $LayoutHeader1XmlDocument.SelectNodes($catLstXPath)[0]
    
        $CatXml = $LayoutHeader1XmlDocument.CreateElement("cat", $LayoutHeader1XmlDocument.DocumentElement.NamespaceURI)
        $CatXml.SetAttribute("type", $Category.ToString().ToLower())
        $CatXml.SetAttribute("pri", $Priority)
        $catLstXml.AppendChild($CatXml) | Out-Null
        
        $LayoutHeader1IsChanged = $True
    }
    
    # 保存
    If ($Layout1IsChanged){
        $Layout1XmlDocument.Save($Layout1Path)
    }
    If ($LayoutHeader1IsChanged){
        $LayoutHeader1XmlDocument.Save($LayoutHeader1Path)
    }
    
}

<#
    .SYNOPSIS
    展開済みの OpenXML 形式のファイルを作成するための一時ファイルに値を設定します

    .PARAMETER Path
    フォルダー名を相対パスか絶対パスを指定します
    相対パスで指定された場合、このスクリプトがあるフォルダーの Temp フォルダー内に保存されます

    .PARAMETER OpenXmlType
    OpenXML 形式の種類を選択します

    .PARAMETER Category
    OpenXmlType で Diagrams を指定したとき: SmartArt グラフィックの選択画面に表示するカテゴリーをすべて削除

    .EXAMPLE
    Remove-OpenXmlFileValue -Path test.glox -OpenXmlType Diagrams -Category

#>
Function Global:Remove-OpenXmlFileValue([Parameter(Mandatory)]$Path, [Parameter(Mandatory)][OpenXmlType]$OpenXmlType, [Switch]$Category){
    Enum OpenXmlType {
        Diagrams
    }
    $Path = Get-OpenXmlTemporaryDirectory -Path $Path

    $LayoutHeader1IsChanged = $False
    $LayoutHeader1Path = (Join-Path $Path "$($OpenXmlType.ToString().ToLower())\layoutHeader1.xml")
    $LayoutHeader1XmlDocument = [XML](Get-Content $LayoutHeader1Path)

    If ($Category){
        $LayoutHeader1XmlDocument.layoutDefHdr.catLst | ForEach-Object {$LayoutHeader1XmlDocument.layoutDefHdr.RemoveChild($_)} | Out-Null
        $LayoutHeader1IsChanged = $True
    }

    If ($LayoutHeader1IsChanged){
        $LayoutHeader1XmlDocument.Save($LayoutHeader1Path)
    }
}


<#
    .SYNOPSIS
    展開済みの OpenXML 形式のファイルを作成するための一時ファイルを使用してOpenXML 形式のファイルを作成します

    .PARAMETER Path
    フォルダー名を相対パスか絶対パスを指定します
    相対パスで指定された場合、このスクリプトがあるフォルダーの Temp フォルダー内より参照されます

    .PARAMETER XmlPath
    OpenXML 形式のファイルのパスを相対パスか絶対パスを指定します
    相対パスで指定された場合、このスクリプトがあるフォルダーの Temp フォルダー内より参照されます

    .PARAMETER CopyToTemplatesDirectory
    Word
    %AppData%\Microsoft\Templates\SmartArt Graphics\

    .PARAMETER Force
    指定された Path が存在する場合でも続行します

    .EXAMPLE
    Save-OpenXmlFile -Path test.glox -XmlPath "C:\Users\Owner\Documents\SmartArt Graphics\test.glox"

#>
Function Global:Save-OpenXmlFile([Parameter(Mandatory)]$Path, [Parameter(Mandatory)]$XmlPath, [Switch]$CopyToTemplatesDirectory, [Switch]$Force){
    $ContentTypeAndPartType = Import-Csv .\Assets\ContentTypeAndPartType.csv
    $XmlPath = Get-OpenXmlExportDirectory -Path $XmlPath
    $TemporaryDirectory = Get-OpenXmlTemporaryDirectory -Path $Path
    
    If (!(Test-Path $TemporaryDirectory -PathType Container)){
        Write-Error "コピー元のフォルダーが見つかりませんでした: $TemporaryDirectory"
    }
    
    If (Test-Path $XmlPath){
        If ($Force){
            Remove-Item $XmlPath -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Else{
            Write-Error "指定されたパスはすでに存在します。このパスを削除するか、-Force オプションを使用してください: $XmlPath"
        }
    }

    New-Item -Path $ExportDirectoryRoot -ItemType Directory -Force | Out-Null
    [System.IO.Packaging.Package]$CurrentPackage = [System.IO.Packaging.Package]::Open($XmlPath,[System.IO.FileMode]::Create,[System.IO.FileAccess]::ReadWrite)

    Get-ChildItem $TemporaryDirectory -Recurse -File | Where-Object Name -ne ".rels" | Sort-Object -Descending | ForEach-Object{
        $Path = $_.FullName.Replace("$TemporaryDirectory\","")
        $Parent = Split-Path $Path -Parent
        $FileName = Split-Path $Path -Leaf
        $FileNameWithoutExtension = ([IO.Path]::GetFileNameWithoutExtension($Path))
        $Path = "/$Parent/$FileName"
        $PathWithoutExtension = "$Parent/$FileNameWithoutExtension"
        
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
        Copy-Item $XmlPath "$env:APPDATA\Microsoft\Templates\SmartArt Graphics\"

        $WordApplication = New-Object -ComObject Word.Application
        Write-Host "Wordを起動しています..."
        
        Start-SmartEditorPicker($WordApplication)

        Write-Host "Wordでの作業が終了しましたら何かキーを押してください"
        Read-Host
        Close-WordWindow($WordApplication)

    }
}

<#
    .SYNOPSIS

    .DESCRIPTION

    .EXAMPLE

#>
Function Global:Close-WordWindow($WordApplication){
    $WscriptShell = New-Object -ComObject Wscript.Shell
    If ($WscriptShell.AppActivate("SmartArt グラフィックの選択")){
        Start-Sleep 1
        $WscriptShell.SendKeys("{Escape}")
    }

    $WordApplication.ActiveDocument.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges) | Out-Null
    If ($WordApplication.ActiveDocument -ne $null){
        Write-Warning "ドキュメントが閉じられていないようです。開いているダイアログがあれば閉じてください。"
        Close-WordWindow($WordApplication, $WordDocument)
    }
    $WordApplication.Quit() | Out-Null
}


<#
    .SYNOPSIS

    .DESCRIPTION

    .EXAMPLE

#>
Function Global:Start-SmartEditorPicker($WordApplication){
    $WordApplication.Caption = $WordApplicationCaption
    $WordApplication.Application.Visible = $True
    $WordDocument = $WordApplication.Documents.Add()
    $WordRange = $WordDocument.Range(0, 0)
    $WordRange.Text = "SmartArtグラフィックの選択画面が表示されなかった場合は、""挿入"" > ""SmartArtグラフィックの挿入"" の順に押してください。`n`nこのWord文書は、PowerShellウィンドウで何かキーを押すと保存なしで終了されます。"
        
    $WscriptShell = New-Object -ComObject Wscript.Shell
    If ($WscriptShell.AppActivate($WordApplication.Caption)){
        $WscriptShell.SendKeys("%N")
        Start-Sleep 1
        $WscriptShell.SendKeys("{ESC}")
        Start-Sleep 1
        $WscriptShell.SendKeys("%N")
        Start-Sleep 1
        $WscriptShell.SendKeys("M")
    }
    Else{
        Write-Error "Wordが起動できなかった可能性があります。"
    }
}
<#
    .SYNOPSIS

    .DESCRIPTION

    .EXAMPLE

#>
Function Global:Clear-OpenXmlCorruptMountPoint($Path){
    $TemporaryDirectory = (Join-Path $TemporaryDirectoryRoot (Split-Path $Path -Leaf))
    If (Test-Path $TemporaryDirectory){
        Remove-Item $TemporaryDirectory -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    }
}


<#
    .SYNOPSIS

    .DESCRIPTION

    .EXAMPLE

#>
Function Global:Get-OpenXmlTemporaryDirectory([Parameter(Mandatory)]$Path){
    If (Test-Path $Path){
        Return $Path
    }ElseIf((Split-Path $Path -Leaf) -eq $Path){
        Return (Join-Path $TemporaryDirectoryRoot (Split-Path $Path -Leaf))
    }Else{
        Return $Path
    }
}
<#
    .SYNOPSIS

    .DESCRIPTION

    .EXAMPLE

#>
Function Global:Get-OpenXmlExportDirectory([Parameter(Mandatory)]$Path){
    If (Test-Path $Path){
        Return $Path
    }ElseIf((Split-Path $Path -Leaf) -eq $Path){
        Return (Join-Path $ExportDirectoryRoot (Split-Path $Path -Leaf))
    }Else{
        Return $Path
    }
}
