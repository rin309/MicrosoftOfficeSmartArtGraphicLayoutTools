# SmartArt Graphic Layout Tools

## How to use
1. Load "MainScript.ps1".
```
& MainScript.ps1
```
2. Mount *.glox file or create from template.
```
# Open
Mount-OpenXmlFile -XmlPath "C:\Users\User\Downloads\MSDNExample.glox"
# Create
Mount-OpenXmlFile -XmlPath "not-exists.glox"
```
3. Edit Temp\*.glox\diagrams\layout1.xml and Temp\*.glox\diagrams\layoutHeader1.xml.

4. Mount *.glox file or create from template.
```
DisMount-OpenXmlFile -XmlPath "C:\Users\User\Downloads\MSDNExample.glox" -CopyToTemplatesDirectory
# Look "Export" directory
```
5. Debug SmartArt by Word.
![SmartArt グラフィックの選択](https://user-images.githubusercontent.com/760251/119538371-03365a00-bdc6-11eb-9e27-2baecbf4faf9.png)

6. Clear temporary directory.
```
Clear-OpenXmlCorruptMountPoint -GloxFileName "MSDNExample.glox"
```

# Memo
- Welcome to the SmartArt Developer Reference: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/dd439463(v=office.12)
- Creating Document Themes with the Office Open XML Formats: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc964302(v=office.12)
- https://github.com/OfficeDev/Open-XML-SDK
