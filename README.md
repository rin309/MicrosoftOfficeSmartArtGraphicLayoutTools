# SmartArt Graphic Layout Tools

## How to use
1. Load "MainScript.ps1".
```
& MainScript.ps1
```
2. Open *.glox file / Create from template.
```
# Open
Mount-OpenXmlFile -Path test.glox -XmlPath "C:\Users\Owner\Documents\SmartArt Graphics\test.glox"
# Create
New-TemporaryOpenXmlFile -Path "not-exists.glox" -OpenXmlType Diagrams
```
3. Edit Temp\*.glox\diagrams\layoutHeader1.xml
```
Rmove-OpenXmlFileValue -Path test.glox -OpenXmlType Diagrams -Category
Set-OpenXmlFileValue -Path test.glox -OpenXmlType Diagrams -Title "テスト" -Description "これはテストです" -Category List -Priority 1000 -UniqueId "urn:contoso.com/office/officeart/2021/6/layout/test"
```
4. Edit Temp\*.glox\diagrams\layout1.xml by your text editor.

5. Save *.glox file.
```
Save-OpenXmlFile -Path test.glox -XmlPath test.glox
# Look "Export" directory
```
5. Debug SmartArt by Word.
![SmartArt グラフィックの選択](https://user-images.githubusercontent.com/760251/119538371-03365a00-bdc6-11eb-9e27-2baecbf4faf9.png)

6. Clear temporary directory.
```
Clear-OpenXmlCorruptMountPoint -Path test.glox
```

# Memo
- Welcome to the SmartArt Developer Reference: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/dd439463(v=office.12)
- Creating Document Themes with the Office Open XML Formats: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc964302(v=office.12)
- https://github.com/OfficeDev/Open-XML-SDK
