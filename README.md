# VBAForm2PowerShell - Excel VBA UserForm to PowerShell GUI (WinForms) Converter
:jp:[日本語の説明はこちら](https://github.com/GUI-Conversion-Tools/VBAForm2Tkinter/blob/main/README_ja.md)<br><br>
This program converts userforms created in Microsoft Excel VBA into PowerShell (WinForms) code.<br>

## Example
<img width="681" height="1275" alt="Image" src="https://github.com/user-attachments/assets/6a37e64a-ea99-4f90-92c5-b2b82794fcf8" /><br>
<img width="704" height="695" alt="Image" src="https://github.com/user-attachments/assets/515589e7-633d-4432-aae1-93c6457578c0" /><br>

## System Requirements
- Supported OS: Windows
- Required Software: Microsoft Excel

## Verified Operating Environments
= Windows 10/11
- Excel 2010(32bit)
- Excel 2016(32bit)
- Excel 2019(64bit)

## Converted Elements
- Variable names (object names)
- Approximate layout and size of controls
- Control colors (foreground, background)
- Text display (Label, CommandButton, CheckBox, ToggleButton, OptionButton, MultiPage)
- Font (typeface, size, bold, italic)
- Borders (Frame [without Caption], TextBox, Label, ListBox, Image)
- Mouse cursor
- Text alignment: left, center, right (Label, TextBox, CheckBox, ToggleButton, OptionButton)
- Default values of TextBox, ComboBox
- Items set in ComboBox, ListBox
- Selection state of OptionButton, CheckBox and ToggleButton
- Transparent background setting specified in BackStyle

## Supported Controls
| VBA Form Class | WinForms Class|
| ------ | ------ |
| Label | Label |
| CommandButton | Button |
| Frame (without Caption) | Panel |
| Frame (with any Caption) | GroupBox |
| TextBox | TextBox |
| SpinButton | NumericUpDown |
| ListBox | ListBox |
| CheckBox | CheckBox |
| ToggleButton | CheckBox<br>(Appearance = [System.Windows.Forms.Appearance]::Button) |
| OptionButton | RadioButton |
| Image | PictureBox |
| ScrollBar | HScrollBar / VScrollBar |
| ComboBox | ComboBox |
| MultiPage | TabControl |


> Note:
SpinButton behaves differently in VBA and WinForms, so appearance may vary depending on placement.<br>
If unsupported controls exist on the form, the conversion will fail. If that case, please remove those controls and run the conversion again.<br>



## Usage
Before using, prepare the Excel workbook containing the user form you want to convert.
Also, ensure that the Immediate Window is visible in the VBE (Visual Basic Editor).<br><br>
<img width="807" height="768" alt="Image" src="https://github.com/user-attachments/assets/b023597f-6f9e-4223-a9a4-1c7c499c194b" /><br><br>
1. Download the latest file from [here](https://github.com/GUI-Conversion-Tools/VBAForm2PowerShell/releases) and extract it. Use the VBAForm2PowerShell.bas file inside.<br>
2. In Excel, go to Developer -> Visual Basic to open VBE.<br>
3. Right-click your project and import the provided .bas file using Import File.<br>
4. In the Immediate Window, enter: Call ConvertForm2PS(UserForm1)<br>
```vb
Call ConvertForm2PS(UserForm1)
```
If you want to save it as a .bat file that can be executed by double-clicking, set the second argument to True.<br>
```vb
Call ConvertForm2PS(UserForm1, True)
```
   > Note: Replace UserForm1 with the object name of the form you want to convert.

5.  If conversion succeeds, a message will appear, and an output.ps1/output.bat file will be created in the same directory as your Excel workbook.<br>
6.  After checking the GUI appearance, edit the .ps1/.bat file and, above [System.Windows.Forms.Application]::EnableVisualStyles(), configure event handlers for controls (e.g., Button.Add_Click({ FunctionName })).<br>


## Control Order (for Controls Without Child Elements)
In WinForms, if you place one Label on top of another, the earlier control appears in front.<br>
However, in VBA, you can change front/back order, so the behavior differs.<br>
The program first reverses controls order and sorts controls by hierarchy level.<br>
Since VBA’s z-order (front/back) cannot currently be retrieved, some displays may not match VBA.<br>

To adjust:<br>
&nbsp;&nbsp;&nbsp;&nbsp;Edit the PowerShell code to use .BringToFront() or .SendToBack() to adjust the z-order.<br>
&nbsp;&nbsp;&nbsp;&nbsp;For new GUIs, instead of overlapping controls, it is recommended to use containers like Frame, which allow clear parent-child relationships.

## Notes on Usage
When using this program in a multi-monitor environment, please temporarily switch to a single monitor or ensure that all monitors have the same scaling percentage.
If monitors with different scaling percentages are mixed, the window size may not be calculated correctly.
