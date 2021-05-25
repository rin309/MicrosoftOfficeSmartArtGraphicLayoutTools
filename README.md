# MicrosoftOfficeSmartArtGraphicLayoutTools

## How to use
1. Load "MainScript.ps1".
```
& MainScript.ps1
```
2. Mount *.glox file or create from template.
```
Mount-OpenXmlFile -XmlPath "C:\Users\User\Downloads\MSDNExample.glox"
```
3. Edit Temp\*.glox\diagrams\layout1.xml and Temp\*.glox\diagrams\layoutHeader1.xml.

4. Mount *.glox file or create from template.
```
DisMount-OpenXmlFile -XmlPath "C:\Users\User\Downloads\MSDNExample.glox" -CopyToTemplatesDirectory
```
5. Debug SmartArt by Word.
```
winword
```
6. Clear temporary directory.
```
Clear-OpenXmlCorruptMountPoint -GloxFileName "MSDNExample.glox"
```

# Memo
- Welcome to the SmartArt Developer Reference: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/dd439463(v=office.12)
- Creating Document Themes with the Office Open XML Formats: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc964302(v=office.12)
- https://github.com/OfficeDev/Open-XML-SDK
