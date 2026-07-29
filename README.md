# VBAForm2PowerShell - VBA UserForm to PowerShell GUI (WinForms) Converter
:jp:[日本語の説明はこちら](https://github.com/GUI-Conversion-Tools/VBAForm2PowerShell/blob/main/README_ja.md)<br><br>
This program converts userforms created in Microsoft Office VBA into PowerShell (WinForms) code.<br>

## Example
<img width="681" height="1275" alt="Image" src="https://github.com/user-attachments/assets/51414770-91c2-44af-874c-d54134efae62" /><br>
<img width="704" height="695" alt="Image" src="https://github.com/user-attachments/assets/86ba4934-ed13-4871-adea-3b285f94f14d" /><br>

## System Requirements
- Supported OS: Windows XP or later
- Required Software: Microsoft Excel/Word/PowerPoint/Outlook 2000 or later
- Recommended Environment: Microsoft Excel 2016 or later

## Verified Operating Environments
- Windows XP(SP3)
- Windows 10/11
- Excel 2000(32bit)
- Excel 2010(32bit)
- Excel 2016(32bit)
- Excel 2019(64bit)
- Word/PowerPoint/Outlook 2000 (32bit)
- Word/PowerPoint/Outlook 2019 (64bit)

## Converted Elements
- Variable names (object names)
- Approximate layout and size of controls
- Control colors (foreground) (Excluding `MultiPage`, `ComboBox` [.Style = fmStyleDropDownList])
- Control colors (background) (Excluding `MultiPage`, `ComboBox` [.Style = fmStyleDropDownList], `ScrollBar`)
- Text display (`Label`, `CommandButton`, `CheckBox`, `ToggleButton`, `OptionButton`, `MultiPage`)
- Font (typeface, size, bold, italic)
- Borders (`Frame [without Caption]`, `TextBox`, `Label`, `ListBox`, `Image`)
- Mouse cursor
- Text alignment: left, center, right (`Label`, `TextBox`, `CheckBox`, `ToggleButton`, `OptionButton`,  `ListView`)
- Default values of `TextBox`, `ComboBox`
- Items set in `ComboBox`, `ListBox`, `ListView`, `TreeView`
- Selection state of `OptionButton`, `CheckBox` and `ToggleButton`
- Transparent background setting specified in `.BackStyle`(Excluding `ComboBox` [.Style = fmStyleDropDownList])
- Images Embedded in Controls (`Image`)
- `.Orientation`/`.Min`/`.Max` property (`ScrollBar`)
- `.Alignment` property (`CheckBox`. `OptionButton`)
- `.TabOrientation` property (`MultiPage`)
- `.Locked` property (`TextBox`, `ListBox`, `ComboBox`)
- `.PasswordChar` property (`TextBox`)
- `.ScrollBars` property (`TextBox`)
- `.WordWrap` property (`TextBox`)
- `.Style` property (`ComboBox`, `MultiPage`)
- `.MultiSelect` property (`ListBox`, `ListView`)
- `.PictureAlignment`/`.PictureSizeMode` property (`Image`)
- `.Scroll` property (`TreeView`)
- `.Expanded` property (`TreeView.Nodes`)

>Note:
>
>-   When `.BackStyle` is `fmBackStyleOpaque`, the control’s own `.BackColor` is used directly.
>-   When `.BackStyle` is `fmBackStyleTransparent`:
>        -   For controls that support transparency in WinForms (e.g., `Label`, `CommandButton`, `CheckBox`, `OptionButton`, `Image`, etc), `.BackColor` is set to `"Transparent"`.
>        -   For controls that **do not support transparency in WinForms** (`TextBox`, `ComboBox`, `ToggleButton`), the background is substituted:
>            -   If the parent control has a `.BackColor`, that color is used.
>            -   If the parent is a `Page` (which does not expose `.BackColor`), a system default color (`&H8000000F&`) is used as a fallback, which matches the visual background color of the `Page`.
>
>-   `.PictureSizeMode`/`.PictureAlignment` is mapped to the corresponding WinForms `.SizeMode` behavior:
>        -   `fmPictureSizeModeClip` → `"Normal"` or `"CenterImage"` (depends on `.PictureAlignment`)
>        -   `fmPictureSizeModeStretch` → `"StretchImage"`
>        -   `fmPictureSizeModeZoom` → `"Zoom"`
>
>        -   When `.PictureSizeMode = fmPictureSizeModeClip`:
>            -   `.PictureAlignment = fmPictureAlignmentCenter` → `"CenterImage"`
>            -   `.PictureAlignment = fmPictureAlignmentTopLeft` → `"Normal"`
>            -   Other alignment values are not supported in WinForms `PictureBox`, and are converted to `"Normal"` (top-left).
>        -   For `"StretchImage"` and `"Zoom"` modes, `.PictureAlignment` is ignored.
>-   MultiPage controls with `.TabOrientation` set to `fmTabOrientationLeft`  or  `fmTabOrientationRight` render tab text vertically in WinForms, unlike VBA, which keeps it horizontal.

## Supported Controls
| VBA Form Class | WinForms Class|
| ------ | ------ |
| `Label` | `Label` |
| `CommandButton` | `Button` |
| `Frame` (without Caption) | `Panel` |
| `Frame` (with any Caption) | `GroupBox` |
| `TextBox` | `TextBox` |
| `SpinButton` | `NumericUpDown` |
| `ListBox` | `ListBox` |
| `CheckBox` | `CheckBox` |
| `ToggleButton` | `CheckBox`<br>(`.Appearance = "Button"`) |
| `OptionButton` | `RadioButton` |
| `Image` | `PictureBox` |
| `ScrollBar` | `HScrollBar` / `VScrollBar` |
| `ComboBox` | `ComboBox` |
| `MultiPage` | `TabControl` |
| `ListView`(`.View=lvwReport`) | `ListView` |
| `TreeView` | `TreeView` |

> Note:
>- `SpinButton` behaves differently in VBA and WinForms, so appearance may vary depending on placement.<br>

If unsupported controls exist on the form, the conversion will fail. If that case, please remove those controls and run the conversion again.<br>



## Usage
Before using, prepare the Excel workbook containing the user form you want to convert.
Also, ensure that the Immediate Window is visible in the VBE (Visual Basic Editor).<br><br>
<img width="807" height="768" alt="Image" src="https://github.com/user-attachments/assets/b023597f-6f9e-4223-a9a4-1c7c499c194b" /><br><br>
1. Download the latest file from [here](https://github.com/GUI-Conversion-Tools/VBAForm2PowerShell/releases) and extract it. Use the `VBAForm2PowerShell.bas` file inside.<br>
2. In Excel, go to Developer -> Visual Basic to open VBE.<br>
3. Right-click your project and import the provided `.bas` file using Import File.<br>
4. In the Immediate Window, enter: `Call ConvertForm2PS(UserForm1)`<br>
```vb
Call ConvertForm2PS(UserForm1)
```
If you want to save it as a `.bat` file that can be executed by double-clicking, set the second argument to `True`.<br>
```vb
Call ConvertForm2PS(UserForm1, True)
```
   > Note: Replace UserForm1 with the object name of the form you want to convert.

5.  If conversion succeeds, a message will appear, and an `output.ps1`/`output.bat` file will be created.<br>
6.  After checking the GUI appearance, edit the `.ps1`/`.bat` file and, above `.ShowDialog()`, configure event handlers for controls (e.g., `Button.Add_Click({ FunctionName })`).<br>

## Output Directory

A dedicated `VBAForm2PS_output` folder is automatically created in the workbook directory, and all generated files are saved there:

### Excel and Word

When running from Excel or Word, the output folder is created in the same directory as the macro-enabled document.

-   **Excel**: Uses the workbook's directory (`ThisWorkbook.Path`)
-   **Word**: Uses the document's directory (`MacroContainer.Path`)

```
WorkbookFolder/
├─ MyWorkbook.xlsm
└─ VBAForm2PS_output/
   ├─ output.ps1
   ├─ image_base64.json
   └─ exported images...
```

### Other Office Applications
When running from other Office applications (such as PowerPoint, Outlook, etc.), or when the current Excel workbook or Word document has not yet been saved, the output folder is created in the user's **Documents** folder instead.

```
C:\Users\%USERNAME%\Documents\
└─ VBAForm2PS_output/
   ├─ output.ps1
   ├─ image_base64.json
   └─ exported images...
```

If the Documents folder cannot be resolved, the output folder will be created in the root of the **C:** drive as a final fallback.

## Parameters

`ConvertForm2PS` accepts the following parameters:

|**Parameter**|**Type**|**Description**                         |
|----------------|-------------------------------|-----------------------------|
|`frms` |`Variant`|**Required.**<br>Accepts a single `UserForm` object or an `Array` of `UserForm` objects to be converted.            |
|`saveAsBat` |`Boolean`|**Optional (Default: `False`).**<br>If set to `True`, the generated PowerShell script will be saved as a `.bat` file that can be executed by double-clicking.|
|`useCls`  |`Boolean` |**Optional (Default: `False`).**<br>If set to `True`, the generated PowerShell code will wrap each form in a PowerShell class structure. This is automatically set to `True` if `frms` is an array.|
|`noMainLoop`  |`Boolean`|**Optional (Default: `False`).**<br>If set to `True`, the `.ShowDialog()` call will be omitted from the end of the generated PowerShell script. When `useCls` is also `True`, this will additionally skip the code that creates the object instances (e.g., `$obj_UserForm1 = [UserForm1]::new()`).|
|`imageMode`  |`String` |**Optional (Default: `"file"`).**<br>Determines how image files used in the UserForm are handled during conversion. You can choose one of the following options:<br>• `"file"` (Default): Images are saved as separate external files in the output directory, and the generated code references these files.<br>• `"disabled"`: Image processing is disabled, and no image-related code is generated.<br>• `"reference-only"`: Similar to `"file"`, generates code that references image files, but does not export the image files. Useful when the image files already exist.<br>• `"base64"`: Images are embedded directly into the generated code as Base64-encoded strings, keeping everything in a single file.<br>• `"base64-dict"`: Images are embedded as Base64 strings within a `Hashtable` inside the generated code.<br>• `"base64-json"`: Images are stored in an external `image_base64.json` file as Base64 strings, and the generated code references the JSON file.<br>• `"base64-json-reference"`: Similar to `"base64-json"`, generates code that references `image_base64.json`, but does not export the JSON file. Useful when the JSON file already exists.|

You can execute the conversion by calling the `ConvertForm2PS` with a single UserForm object or an array of multiple UserForms.

```vb
' Example: Converting a single form
Call ConvertForm2PS(UserForm1)

' Example: Converting a single form (Class-based style)
Call ConvertForm2PS(UserForm1, useCls:=True)

' Example: Converting multiple forms (Automatically uses Class-based style)
Call ConvertForm2PS(Array(UserForm1, UserForm2))

' Example: Converting a single form (With image streams embedded directly as Base64 text)
Call ConvertForm2PS(UserForm1, imageMode:="base64")
```

## Control Order (for Controls Without Child Elements)
In WinForms, if you place one `Label` on top of another, the earlier control appears in front.<br>
However, in VBA, you can change front/back order, so the behavior differs.<br>
The program first reverses controls order and sorts controls by hierarchy level.<br>
Since VBA’s z-order (front/back) cannot currently be retrieved, some displays may not match VBA.<br>

To adjust:<br>
&nbsp;&nbsp;&nbsp;&nbsp;Edit the PowerShell code to use `.BringToFront()` or `.SendToBack()` to adjust the z-order.<br>
&nbsp;&nbsp;&nbsp;&nbsp;For new GUIs, instead of overlapping controls, it is recommended to use containers like `Frame`, which allow clear parent-child relationships.

