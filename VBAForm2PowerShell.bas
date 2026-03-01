Attribute VB_Name = "VBAForm2PowerShell"

' VBAForm2PowerShell v1.1.0
' https://github.com/GUI-Conversion-Tools/VBAForm2PowerShell
' Copyright (c) 2025-2026 ZeeZeX
' This software is released under the MIT License.
' https://opensource.org/licenses/MIT

Option Explicit


#If VBA7 Then
    ' 64bit Office / VBA7 or later
    Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
    Private Type RECT: Left As Long: Top As Long: Right As Long: Bottom As Long: End Type
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#Else
    ' 32bit Office
    Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
    Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Type RECT: Left As Long: Top As Long: Right As Long: Bottom As Long: End Type
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If

Private Const FORM_WINDOW_NAME As String = "window"

Sub TestRunConversion2PS()
    Call ConvertForm2PS(UserForm1)
End Sub

Sub TestRunConversion2PS_2()
    Call ConvertForm2PS(Array(UserForm1, UserForm2))
End Sub

Sub TestRunConversion2PS_3()
    Call ConvertForm2PS(UserForm1, saveAsBat:=True, useCls:=True)
End Sub

Sub ConvertForm2PS(ByVal frms As Variant, Optional ByVal saveAsBat As Boolean = False, Optional ByVal useCls As Boolean = False, Optional ByVal noMainLoop As Boolean = False)
    
    ' frms: Variant
    '   Accepts a single UserForm object or an Array of UserForm objects to be converted.
    ' saveAsBat: Boolean
    '   If set to True, the generated PowerShell script will be saved as a .bat file that can be executed by double-clicking.
    ' useCls: Boolean
    '   If set to True, the generated PowerShell code will wrap each form in a PowerShell class structure.
    '   This is automatically set to True if frms is an array.
    ' noMainLoop: Boolean
    '   If set to True, the .ShowDialog() call will be omitted from the end of the generated PowerShell script.
    
    Dim code As String
    Dim filePath As String
    Dim saveDir As String
    code = VBAForm2PSWinForms(frms, useCls, noMainLoop)
    If code <> "" Then
        If ThisWorkbook.Path = "" Then
            saveDir = "C:"
        Else
            saveDir = ThisWorkbook.Path
        End If
        If saveAsBat Then
            code = GenerateBatchCode() & vbLf & vbLf & code
            filePath = saveDir & "\output.bat"
            Call SaveUTF8Text_NoBOM(filePath, code) 'Batch does not support UTF-8(BOM)
        Else
            filePath = saveDir & "\output.ps1"
            Call SaveUTF8BOMText(filePath, code) ' In PowerShell 5.1, .ps1 does not support UTF-8(NoBOM)
        End If
        
        MsgBox "Saved: " & filePath
    Else
        MsgBox "Conversion failed."
    End If
    
End Sub


Function VBAForm2PSWinForms(ByVal frms As Variant, Optional ByVal useCls As Boolean = False, Optional ByVal noMainLoop As Boolean = False) As String
    Dim root As Variant
    Dim indent As String
    Dim prefix As String
    Dim clsNumber As Long
    Dim formName As String
    Dim controlName As String
    Dim parentName As String
    Dim childName As String
    Dim itemsListName As String
    Dim instanceName As String
    Dim toplevelInstanceName As String
    Dim unavailableNames() As Variant
    Dim ctrl As MSForms.Control
    Dim ctrls As Collection
    Dim item As Variant
    Dim r As String
    Const q As String = """"
    Dim fontStyle As String
    Dim fontOpts As String
    Dim widgetType As String
    Dim styleName As String
    Dim sizeFactorsAndOffsets() As Variant
    Dim sizeFactorX As Double
    Dim sizeFactorY As Double
    Dim pixelWidth As Long
    Dim pixelHeight As Long
    Dim pixelTop As Long
    Dim pixelLeft As Long
    Dim i As Long
    Dim orientation As String
    Dim cursorType As String
    Dim caption As String
    Dim dpis() As Variant
    Dim scaleFactorX As Double
    Dim scaleFactorY As Double
    Dim colorSetting As String
    
    r = ""
    
    If IsArray(frms) Then
        useCls = True
    Else
        frms = Array(frms)
    End If
    
    If useCls Then
        indent = "        "
        prefix = "$this."
    Else
        indent = ""
        prefix = "$"
    End If
    
    r = r & "Add-Type -AssemblyName System.Windows.Forms" & vbLf
    r = r & "Add-Type -AssemblyName System.Drawing" & vbLf
    r = r & "[System.Windows.Forms.Application]::EnableVisualStyles()" & vbLf
    r = r & vbLf
    
    For Each root In frms
        unavailableNames = VBA.Array("", "System")
        
        For i = LBound(unavailableNames) To UBound(unavailableNames)
            unavailableNames(i) = LCase(unavailableNames(i))
        Next
        
        If ContainsValue(unavailableNames, LCase(root.Name)) Then
            MsgBox GenerateUnavailableNameMessage(root)
            r = ""
            VBAForm2PSWinForms = r
            Exit Function
        End If
        unavailableNames(0) = LCase(FORM_WINDOW_NAME)
        
        dpis = GetPrimaryMonitorDPI
        scaleFactorX = dpis(0) / 96
        scaleFactorY = dpis(1) / 96
        
        ' Get factor for size conversion
        sizeFactorsAndOffsets = GetUserFormScaleFactorsAndOffsets(root)
        sizeFactorX = sizeFactorsAndOffsets(0)
        sizeFactorY = sizeFactorsAndOffsets(1)
        ' Convert UserForm's size to pixel size
        pixelWidth = UserFormSizeToPixel(root.Width, sizeFactorX)
        pixelHeight = UserFormSizeToPixel(root.Height, sizeFactorY)
        pixelWidth = pixelWidth - sizeFactorsAndOffsets(2)
        pixelHeight = pixelHeight - sizeFactorsAndOffsets(3)
        ' Divide window size by scaling factor
        pixelWidth = Round(pixelWidth / scaleFactorX)
        pixelHeight = Round(pixelHeight / scaleFactorY)
        
        formName = GetControlName(root, prefix, useCls)
        
        Set ctrls = New Collection
        For Each ctrl In root.Controls
            ctrls.Add ctrl
        Next ctrl
        Set ctrls = ReverseCollection(ctrls)
        Set ctrls = SortFormControlsByDepth(ctrls)
        
        If useCls Then
            r = r & "class " & root.Name & "{" & vbLf
            ' Declare instance variables
            r = r & "    " & "[object]$" & FORM_WINDOW_NAME & vbLf
            For Each ctrl In ctrls
                r = r & "    " & "[object]$" & ctrl.Name & vbLf
                If ContainsValue(Array("ComboBox", "ListBox"), TypeName(ctrl)) Then
                    itemsListName = ctrl.Name & "_items_value"
                    r = r & "    " & "[object]$" & itemsListName & vbLf
                End If
                
                If TypeName(ctrl) = "MultiPage" Then
                    For Each item In ctrl.Pages
                        r = r & "    " & "[object]$" & item.Name & vbLf
                    Next
                End If
            Next
            r = r & "    " & root.Name & "() {" & vbLf
        End If
        
        r = r & indent & formName & " = " & "New-Object System.Windows.Forms.Form" & vbLf
        
        caption = root.caption
        caption = Convert2PowerShellFormatText(caption)
        r = r & indent & formName & ".Text = " & q & caption & q & vbLf
        r = r & indent & formName & ".ClientSize = New-Object System.Drawing.Size(" & pixelWidth & ", " & pixelHeight & ")" & vbLf
        r = r & indent & formName & ".MaximizeBox = $false" & vbLf
        r = r & indent & formName & ".FormBorderStyle = " & DotNetTypeLiteral("System.Windows.Forms.FormBorderStyle", useCls) & "::FixedSingle" & vbLf ' Disable window resizing
        r = r & indent & formName & ".BackColor = " & q & FormColorToHex(root.BackColor) & q & vbLf
        r = r & indent & formName & ".AutoScaleMode = " & DotNetTypeLiteral("System.Windows.Forms.AutoScaleMode", useCls) & "::None" & vbLf
        
        
        cursorType = GetControlCursorType(root)
        If cursorType <> "" Then
            r = r & indent & formName & ".Cursor = " & DotNetTypeLiteral("System.Windows.Forms.Cursors", useCls) & "::" & cursorType & vbLf
        Else
            r = r & indent & formName & ".Cursor = $null" & vbLf
        End If
        
        r = r & vbLf

        For Each ctrl In ctrls
            If GetWinFormsControlName(ctrl) = "" Then
                MsgBox GenerateUnsupportedControlMessage(ctrl)
                r = ""
                VBAForm2PSWinForms = r
                Exit Function
            End If
            
            If ContainsValue(unavailableNames, LCase(ctrl.Name)) Then
                MsgBox GenerateUnavailableNameMessage(ctrl)
                r = ""
                VBAForm2PSWinForms = r
                Exit Function
            End If
            
            widgetType = "System.Windows.Forms." & GetWinFormsControlName(ctrl)
            
            pixelLeft = UserFormSizeToPixel(ctrl.Left, sizeFactorX)
            pixelTop = UserFormSizeToPixel(ctrl.Top, sizeFactorY)
            pixelWidth = UserFormSizeToPixel(ctrl.Width, sizeFactorX)
            pixelHeight = UserFormSizeToPixel(ctrl.Height, sizeFactorY)
            
            pixelLeft = Round(pixelLeft / scaleFactorX)
            pixelTop = Round(pixelTop / scaleFactorY)
            pixelWidth = Round(pixelWidth / scaleFactorX)
            pixelHeight = Round(pixelHeight / scaleFactorY)
            
            controlName = GetControlName(ctrl, prefix, useCls)
            parentName = GetParentName(ctrl, prefix, useCls)
            itemsListName = controlName & "_items_value"
            
            r = r & indent & controlName & " = " & "New-Object" & " " & widgetType & vbLf
            r = r & indent & parentName & ".Controls.Add(" & controlName & ")" & vbLf
            r = r & indent & controlName & ".Location = New-Object System.Drawing.Point(" & pixelLeft & ", " & pixelTop & ")" & vbLf
            r = r & indent & controlName & ".Size = New-Object System.Drawing.Size(" & pixelWidth & ", " & pixelHeight & ")" & vbLf
            
            If GetWinFormsControlName(ctrl) = "GroupBox" Or Not ContainsValue(Array("Frame", "Image", "ScrollBar", "MultiPage"), TypeName(ctrl)) Then
                ' Set ForeColor
                r = r & indent & controlName & ".ForeColor = " & q & FormColorToHex(ctrl.ForeColor) & q & vbLf
            End If
            
            If Not ContainsValue(Array("ScrollBar"), TypeName(ctrl)) Then
                ' Set BackColor
                colorSetting = q & FormColorToHex(ctrl.BackColor) & q
                If ContainsValue(Array("Label", "TextBox", "CommandButton", "CheckBox", "ToggleButton", "OptionButton", "Image", "ComboBox"), TypeName(ctrl)) Then
                    If ctrl.BackStyle = fmBackStyleTransparent Then
                        If Not ContainsValue(Array("TextBox", "ComboBox"), TypeName(ctrl)) Then
                            colorSetting = DotNetTypeLiteral("System.Drawing.Color", useCls) & "::Transparent"
                        Else
                            ' Apply the BackColor of the parent control because TextBox and ComboBox do not support [System.Drawing.Color]::Transparent
                            If TypeName(ctrl.Parent) <> "Page" Then
                                colorSetting = q & FormColorToHex(ctrl.Parent.BackColor) & q
                            Else
                                ' Because the Page control does not have a BackColor property, set the color to &H8000000F&, which matches the background color of the Page
                                colorSetting = q & FormColorToHex(&H8000000F) & q
                            End If
                        End If
                    End If
                End If
                r = r & indent & controlName & ".BackColor = " & colorSetting & vbLf
                
            End If
            
            
            If GetWinFormsControlName(ctrl) = "GroupBox" Or ContainsValue(Array("Label", "CommandButton", "CheckBox", "ToggleButton", "OptionButton"), TypeName(ctrl)) Then
                caption = ctrl.caption
                caption = Convert2PowerShellFormatText(caption)
                r = r & indent & controlName & ".Text = " & q & caption & q & vbLf
            End If
            
            If ContainsValue(Array("CheckBox", "OptionButton"), TypeName(ctrl)) Then
                If ctrl.Alignment = fmAlignmentLeft Then
                    r = r & indent & controlName & ".RightToLeft = " & DotNetTypeLiteral("System.Windows.Forms.RightToLeft", useCls) & "::Yes" & vbLf
                End If
            End If
            
            If TypeName(ctrl) = "ToggleButton" Then
                r = r & indent & controlName & ".Appearance = " & DotNetTypeLiteral("System.Windows.Forms.Appearance", useCls) & "::Button" & vbLf
                r = r & indent & controlName & ".FlatStyle = " & DotNetTypeLiteral("System.Windows.Forms.FlatStyle", useCls) & "::Flat" & vbLf
            End If
            
            If TypeName(ctrl) = "CommandButton" Then
                r = r & indent & controlName & ".FlatStyle = " & DotNetTypeLiteral("System.Windows.Forms.FlatStyle", useCls) & "::Popup" & vbLf
            End If
            
            If TypeName(ctrl) = "TextBox" Then
                r = r & indent & controlName & ".Text = " & q & Convert2PowerShellFormatText(ctrl.text) & q & vbLf
                r = r & indent & controlName & ".Multiline = " & "$" & LCase(CBool(ctrl.Multiline)) & vbLf
            End If
            
            If TypeName(ctrl) = "ComboBox" Then
                r = r & indent & itemsListName & " = " & GetListBoxValue(ctrl, indent) & vbLf
                r = r & indent & controlName & ".Items.AddRange(" & itemsListName & ")" & vbLf
                r = r & indent & controlName & ".Text = " & q & Convert2PowerShellFormatText(ctrl.text) & q & vbLf
            End If
            
            If TypeName(ctrl) = "ListBox" Then
                r = r & indent & itemsListName & " = " & GetListBoxValue(ctrl, indent) & vbLf
                r = r & indent & controlName & ".Items.AddRange(" & itemsListName & ")" & vbLf
            End If
            
            If TypeName(ctrl) = "ScrollBar" Then
                r = r & indent & controlName & ".Minimum = " & ctrl.Min & vbLf
                r = r & indent & controlName & ".Maximum = " & ctrl.Max & vbLf
            End If
            
            ' Set each Caption in MultiPage
            If TypeName(ctrl) = "MultiPage" Then
                For Each item In ctrl.Pages
                    childName = GetControlName(item, prefix, useCls)
                    caption = item.caption
                    caption = Convert2PowerShellFormatText(caption)
                    r = r & indent & childName & " = New-Object System.Windows.Forms.TabPage" & vbLf
                    r = r & indent & controlName & ".Controls.Add(" & childName & ")" & vbLf
                    r = r & indent & childName & ".Text = " & q & caption & q & vbLf
                Next
            End If
            
            ' Font size is rounded because VBA officially does not support decimal fraction in font settings
            If GetWinFormsControlName(ctrl) = "GroupBox" Or Not ContainsValue(Array("Frame", "ScrollBar", "Image", "SpinButton"), TypeName(ctrl)) Then
                fontStyle = ""
                'fontOpts = ""
                
                If ctrl.Font.Bold Then fontStyle = fontStyle & DotNetTypeLiteral("System.Drawing.FontStyle", useCls) & "::Bold"
                If ctrl.Font.Italic Then
                    If fontStyle <> "" Then fontStyle = fontStyle & " -bor "
                    fontStyle = fontStyle & DotNetTypeLiteral("System.Drawing.FontStyle", useCls) & "::Italic"
                End If
                If ctrl.Font.Underline Then
                    If fontStyle <> "" Then fontStyle = fontStyle & " -bor "
                    fontStyle = fontStyle & DotNetTypeLiteral("System.Drawing.FontStyle", useCls) & "::Underline"
                End If
                If ctrl.Font.Strikethrough Then
                    If fontStyle <> "" Then fontStyle = fontStyle & " -bor "
                    fontStyle = fontStyle & DotNetTypeLiteral("System.Drawing.FontStyle", useCls) & "::Strikeout"
                End If
                
                If fontStyle <> "" Then fontStyle = ", (" & fontStyle & ")"
                
                r = r & indent & controlName & ".Font = New-Object System.Drawing.Font(" & q & ctrl.Font.Name & q & ", " & Round(ctrl.Font.Size) & fontStyle & ")" & vbLf
            End If
            
            
            If GetWinFormsControlName(ctrl) <> "GroupBox" And ContainsValue(Array("Frame", "TextBox", "Label", "ListBox", "Image"), TypeName(ctrl)) Then
                ' WinForms' Combobox does not support customizing border style
                r = r & indent & controlName & GetBorderSetting(ctrl, useCls) & vbLf
            End If
            
            If ContainsValue(Array("Label", "TextBox", "CheckBox", "ToggleButton", "OptionButton"), TypeName(ctrl)) Then
                r = r & indent & controlName & GetTextAlignSetting(ctrl, useCls) & vbLf
            End If
            
            ' Set mouse cursor
            If TypeName(ctrl) <> "MultiPage" Then
                cursorType = GetControlCursorType(ctrl)
                If cursorType <> "" Then
                    r = r & indent & controlName & ".Cursor = " & DotNetTypeLiteral("System.Windows.Forms.Cursors", useCls) & "::" & cursorType & vbLf
                Else
                    r = r & indent & controlName & ".Cursor = $null" & vbLf
                End If
            End If
            
            If TypeName(ctrl) = "Image" Then
                r = r & indent & "#" & controlName & ".Image = " & DotNetTypeLiteral("System.Drawing.Image", useCls) & "::FromFile(" & q & "C:\path\to\your\image.png" & q & ")" & vbLf
                r = r & indent & "#" & controlName & ".SizeMode = " & DotNetTypeLiteral("System.Windows.Forms.PictureBoxSizeMode", useCls) & "::Normal" & vbLf
            End If
            
            r = r & vbLf
                
        Next ctrl
        r = r & SetWinFormsButtonValues(ctrls, indent, prefix, useCls) & vbLf
        If Not useCls And Not noMainLoop Then
            r = r & formName & ".ShowDialog() | Out-Null"
        End If
        
        If useCls Then
            r = r & indent & "}" & vbLf
        End If
        
        If useCls Then
            r = r & "}" & vbLf & vbLf
        End If
        
    Next root
    
    If useCls And Not noMainLoop Then
        clsNumber = 0
        For Each root In frms
            clsNumber = clsNumber + 1
            instanceName = "$obj_" & root.Name
            If clsNumber <= 1 Then
                r = r & instanceName & " = [" & root.Name & "]::new()" & vbLf
                toplevelInstanceName = instanceName
            Else
                r = r & instanceName & " = [" & root.Name & "]::new()" & vbLf
            End If
            r = r & instanceName & "." & FORM_WINDOW_NAME & ".ShowDialog() | Out-Null" & vbLf
        Next
    End If
    
    VBAForm2PSWinForms = r
End Function

Private Function GetParentName(ByVal ctrl As Object, ByVal prefix As String, ByVal useCls As Boolean) As String
    Dim parentName As String
    ' If the object name and type name of a control match, it is considered a UserForm.
    If ctrl.Parent.Name = TypeName(ctrl.Parent) And useCls Then
        parentName = prefix & FORM_WINDOW_NAME
    Else
        parentName = prefix & ctrl.Parent.Name
    End If
    GetParentName = parentName
End Function

Private Function GetControlName(ByVal ctrl As Object, ByVal prefix As String, ByVal useCls As Boolean) As String
    Dim controlName As String
    ' If the object name and type name of a control match, it is considered a UserForm.
    If ctrl.Name = TypeName(ctrl) And useCls Then
        controlName = prefix & FORM_WINDOW_NAME
    Else
        controlName = prefix & ctrl.Name
    End If
    GetControlName = controlName
End Function

Private Function DotNetTypeLiteral(ByVal dotNetTypeName As String, ByVal useCls As Boolean) As String
    ' Referencing a .NET assembly type such as [System.Windows.Forms.Cursors]
    ' inside a class definition causes an error.
    ' This happens because PowerShell classes are compiled before runtime code
    ' (e.g., Add-Type) is executed.
    ' Instead, use ("System.Windows.Forms.Cursors" -as [type]).
    
    ' "System.Windows.Forms.Cursors", useCls:=True -> ("System.Windows.Forms.Cursors" -as [type])
    ' "System.Windows.Forms.Cursors", useCls:=False -> [System.Windows.Forms.Cursors]
    Dim result As String
    If useCls Then
        result = "(" & """" & dotNetTypeName & """" & " -as [type])"
    Else
        result = "[" & dotNetTypeName & "]"
    End If
    DotNetTypeLiteral = result
End Function

Private Function GetBorderSetting(ByVal ctrl As Object, ByVal useCls As Boolean) As String
    Dim r As String
    Const q As String = """"
    Dim borderSetting As String
    borderSetting = "FixedSingle"

    Select Case ctrl.BorderStyle
        Case 1
            ' SpecialEffect is 0 if BorderStyle is 1
            borderSetting = "FixedSingle"
        Case 0
            Select Case ctrl.SpecialEffect
                Case 0
                    borderSetting = "None"
                Case 1
                    borderSetting = "Fixed3D"
                Case 2
                    borderSetting = "Fixed3D"
                Case 3
                    borderSetting = "FixedSingle"
                Case 6
                    borderSetting = "FixedSingle"
            End Select
    End Select

    r = ".BorderStyle = " & DotNetTypeLiteral("System.Windows.Forms.BorderStyle", useCls) & "::" & borderSetting
    GetBorderSetting = r
End Function

Private Function GetTextAlignSetting(ByVal ctrl As Object, ByVal useCls As Boolean) As String
   Dim r As String
   Const q As String = """"
   Dim position As String
   r = ""
   
   Select Case ctrl.TextAlign
        Case fmTextAlignLeft
            position = "Left"
        Case fmTextAlignCenter
            position = "Center"
        Case fmTextAlignRight
            position = "Right"
        Case Else
            position = "Center"
    End Select
    
    If TypeName(ctrl) = "TextBox" Then
        position = q & position & q
    Else
        position = DotNetTypeLiteral("System.Drawing.ContentAlignment", useCls) & "::Top" & position
    End If
    
    r = r & ".TextAlign = " & position
    GetTextAlignSetting = r
End Function

Private Function GetWinFormsControlName(ByVal ctrl As Object) As String
    Dim r As String
    Select Case TypeName(ctrl)
        Case "Label"
            r = "Label"
        Case "CommandButton"
            r = "Button"
        Case "Frame"
            If ctrl.caption = "" Then
                r = "Panel"
            Else
                r = "GroupBox"
            End If
        Case "TextBox"
            r = "TextBox"
        Case "SpinButton"
            r = "NumericUpDown"
        Case "ListBox"
            r = "ListBox"
        Case "CheckBox"
            r = "CheckBox"
        Case "ToggleButton"
            r = "CheckBox"
        Case "OptionButton"
            r = "RadioButton"
        Case "Image"
            r = "PictureBox"
        Case "ScrollBar"
            Select Case ctrl.orientation
                Case -1
                    If ctrl.Width > ctrl.Height Then
                        r = "HScrollBar"
                    Else
                        r = "VScrollBar"
                    End If
                    
                Case 0
                    r = "VScrollBar"
                Case 1
                    r = "HScrollBar"
                Case Else
                    r = "VScrollBar"
            End Select
        Case "ComboBox"
            r = "ComboBox"
        Case "MultiPage"
            r = "TabControl"
        Case Else
            r = ""
    End Select
    GetWinFormsControlName = r
End Function

Private Function GetControlCursorType(ByVal ctrl As Object) As String
    Dim cursorType As String
    Select Case ctrl.MousePointer
        Case fmMousePointerDefault
            cursorType = ""      ' Default cursor
        Case fmMousePointerArrow
            cursorType = "Arrow"        ' Arrow(normal)
        Case fmMousePointerCross
            cursorType = "Cross"        ' Cross
        Case fmMousePointerIBeam
            cursorType = "IBeam"        ' For inputting text
        Case fmMousePointerSizeNESW
            cursorType = "SizeNESW"     ' Arrow(NESW)
        Case fmMousePointerSizeNS
            cursorType = "SizeNS"       ' Arrow(NS)
        Case fmMousePointerSizeNWSE
            cursorType = "SizeNWSE"     ' Arrow(NWSE)
        Case fmMousePointerSizeWE
            cursorType = "SizeWE"       ' Arrow(WE)
        Case fmMousePointerUpArrow
            cursorType = "UpArrow"      ' Arrow(up)
        Case fmMousePointerHourGlass
            cursorType = "WaitCursor"   ' Busy(hourglass)
        Case fmMousePointerNoDrop
            cursorType = "No"           ' "Not allowed" synbol
        Case fmMousePointerAppStarting
            cursorType = "AppStarting"  ' Busy(AppStarting)
        Case fmMousePointerHelp
            cursorType = "Help"         ' Question arrow
        Case fmMousePointerSizeAll
            cursorType = "SizeAll"      ' Four headed Arrow
        Case Else
            cursorType = ""      ' Others are default cursor.
    End Select
    GetControlCursorType = cursorType
End Function


Private Function SetWinFormsButtonValues(ByVal ctrls As Variant, ByVal indent As String, ByVal prefix As String, ByVal useCls As Boolean) As String
    Dim ctrl As Variant
    Dim controlName As String
    Dim value As Boolean
    Dim r As String
    r = ""
    For Each ctrl In ctrls
        controlName = GetControlName(ctrl, prefix, useCls)
        If ContainsValue(Array("OptionButton", "CheckBox", "ToggleButton"), TypeName(ctrl)) Then
            r = r & indent & controlName & ".Checked = " & "$" & LCase(CBool(ctrl.value)) & vbLf
        End If
    Next
    SetWinFormsButtonValues = r
End Function

Private Function GetListBoxValue(ByVal ctrl As Object, ByVal indent As String) As String
    ' Retrieve the items of a ListBox or ComboBox as a string in the format @("1", "2", "3").
    Const q As String = """"
    Dim item As Variant
    Dim i As Long: i = 0
    Dim r As String
    Dim listIndent As String: listIndent = "    " & indent
    Const maxItemsPerLine As Long = 3
    r = ""
    If ctrl.ListCount > 0 Then
        If ctrl.ListCount > maxItemsPerLine Then r = r & vbLf & listIndent
        For Each item In ctrl.List
            i = i + 1
            r = r & q & Convert2PowerShellFormatText(item) & q
            If Not i = ctrl.ListCount Then
                r = r & ", "
                If i Mod maxItemsPerLine = 0 And ctrl.ListCount > maxItemsPerLine Then r = r & vbLf & listIndent
            Else
                If ctrl.ListCount > maxItemsPerLine Then r = r & vbLf
                Exit For
            End If
        Next item
    End If
    If ctrl.ListCount > maxItemsPerLine Then
        r = "@(" & r & indent & ")"
    Else
        r = "@(" & r & ")"
    End If
    GetListBoxValue = r
End Function

Private Function Convert2PowerShellFormatText(ByVal text As String) As String
    ' Escape special characters in the string
    Dim targetChars() As Variant
    Dim char As Variant
    targetChars = Array("`", """", "$", "{", "}")
    For Each char In targetChars
        text = VBA.Replace(text, char, "`" & char)
    Next
    ' Convert VBA line breaks to PowerShell format
    ' vbCrLf should be replaced first
    text = VBA.Replace(text, vbCrLf, vbLf)
    text = VBA.Replace(text, vbCr, vbLf)
    text = VBA.Replace(text, vbLf, "`r`n")
    Convert2PowerShellFormatText = text
End Function


Private Function GenerateBatchCode() As String
    ' Generate batch(.bat) code for running PowerShell code
    Const q As String = """"
    Dim code As String
    Dim codeArray() As Variant
    Dim i As Long
    Dim argsToPass As String
    argsToPass = ""
    Const loopCnt As Long = 9
    For i = 1 To loopCnt
        argsToPass = argsToPass & "\" & q & "%" & i & "\" & q
        If i <> loopCnt Then argsToPass = argsToPass & ","
    Next i
    codeArray = Array( _
    ":DUMMY for($i=1;$i -eq 0;$i++) {echo DUMMY} <#", _
    "", _
    "@echo off", _
    "chcp 65001 > nul", _
    "set " & q & "DirPath=%~dp0" & q, _
    "set " & q & "lastChar=%DirPath:~-1%" & q, _
    "if " & q & "%lastChar%" & q & "==" & q & "\" & q & " (", _
    "    set " & q & "DirPath=%DirPath:~0,-1%" & q, _
    ")", _
    "set ME=%~dpnx0", _
    "if /i CHK%1==CHK/C (", _
    "  set CHK=EXIT", _
    "  shift", _
    ") else (", _
    "  set CHK=PAUSE", _
    ")", _
    "powershell -ExecutionPolicy Unrestricted -Command " & q & "Set-Location \" & q & "%DirPath%\" & q & "; Invoke-Expression -Command (@('$parm=@(" & argsToPass & ")') + (Get-Content '%ME%' -Encoding UTF8) -join \" & q & "`n\" & q & ")" & q, _
    "", _
    "if /i %CHK%==EXIT exit /b", _
    "pause", _
    "exit /b", _
    "#>", _
    "# The following is PowerShell code." _
    )
    code = Join(codeArray, vbLf)
    GenerateBatchCode = code
End Function

Private Function FormColorToHex(ByVal clr As Long) As String
    Dim r As Long, g As Long, b As Long
    ' Convert a system color to its decimal color code when the parameter is a system color
    If 0 > clr Or clr >= 2147483648# Then
        clr = GetSysColor(clr And &HFF)
    End If
    ' Retrieve each component of the RGB color.
    r = clr And &HFF            ' Extract low-order 8 bits
    g = (clr \ &H100) And &HFF  ' Extract bits 8-15
    b = (clr \ &H10000) And &HFF ' Extract bits 16-23
    
    ' Convert the decimal RGB values to a #RRGGBB hex string and return it
    FormColorToHex = "#" & _
                     Right("0" & Hex(r), 2) & _
                     Right("0" & Hex(g), 2) & _
                     Right("0" & Hex(b), 2)
End Function


Private Function ContainsValue(ByVal itemList As Variant, ByVal value As Variant) As Boolean
    ' Check if a specific value exists in Array/Collection/Dictionary
    ' itemList - Array/Collection/Dictionary to search
    ' value - value to check
    ' Performs strict type comparison for non-numeric values
    ' Nested arrays are not supported. Objects are compared by reference
    ' Dependency: IsStrictlyEqual(helper function)
    Dim item As Variant
    Dim temp As Variant
    If LCase(TypeName(itemList)) = "dictionary" Then
        itemList = itemList.items
    End If
    If IsArray(itemList) Then
        On Error GoTo Finally
        ' Uninitialized Array -> False
        temp = LBound(itemList)
        On Error GoTo 0
    End If
    For Each item In itemList
    
        If IsStrictlyEqual(item, value) Then
            ContainsValue = True
            Exit Function
        End If
    Next
Finally:
    ContainsValue = False
    
End Function

Private Function IsStrictlyEqual(ByVal value1 As Variant, ByVal value2 As Variant) As Boolean
    ' Performs a strict equality comparison including data types.
    ' Numeric types (Integer, Long, Double, etc.) are treated as compatible.
    ' Boolean and Date types are NOT treated as numeric.
    Dim t1 As VbVarType, t2 As VbVarType
    t1 = VarType(value1)
    t2 = VarType(value2)
    
    ' Returns True if objects point to the same reference.
    ' Objects are evaluated first to prevent false matches (e.g., Empty vs empty Cells).
    ' (Also applies to variables holding both objects and other data types)
    If IsObject(value1) Or IsObject(value2) Then
        If IsObject(value1) And IsObject(value2) Then
            IsStrictlyEqual = (value1 Is value2)
        End If
        Exit Function
    End If
    
    ' Null / Empty
    If IsNull(value1) Or IsNull(value2) Then
        IsStrictlyEqual = (IsNull(value1) And IsNull(value2))
        Exit Function
    ElseIf IsEmpty(value1) Or IsEmpty(value2) Then
        IsStrictlyEqual = (IsEmpty(value1) And IsEmpty(value2))
        Exit Function
    End If
    
    
    ' Arrays are not supported (Extend if necessary).
    If IsArray(value1) Or IsArray(value2) Then
        IsStrictlyEqual = False
        Exit Function
    End If
    
    ' Error values
    If t1 = vbError Or t2 = vbError Then
        IsStrictlyEqual = (t1 = t2 And value1 = value2)
        Exit Function
    End If
    
    ' String, Date, Boolean
    If (t1 = vbString Or t2 = vbString) Or (t1 = vbDate Or t2 = vbDate) Or (t1 = vbBoolean Or t2 = vbBoolean) Then
        IsStrictlyEqual = (t1 = t2 And value1 = value2)
        Exit Function
    End If
    
    ' Other data types (e.g., Numeric)
    On Error Resume Next
    IsStrictlyEqual = (value1 = value2)
    Exit Function
    On Error GoTo 0
    IsStrictlyEqual = False
End Function

Private Function Win32_FindWindowW(ByVal className As String, ByVal windowTitle As String) As LongPtr
    ' Get the window's hwnd
    ' className: The window's class name (exact match). If not specified, provide "", Empty, or vbNullString
    ' windowTitle: The window's title (exact match). If not specified, provide "", Empty, or vbNullString
    ' Example: Get Excel's main window by specifying only the class name
    ' hwnd = Win32_FindWindowW("XLMAIN", Empty)
    Dim hwnd As LongPtr
    If className = "" Then className = vbNullString
    If windowTitle = "" Then windowTitle = vbNullString
    hwnd = FindWindowW(StrPtr(className), StrPtr(windowTitle))
    Win32_FindWindowW = hwnd
End Function

Private Function GetUserFormScaleFactorsAndOffsets(ByVal frm As Object) As Variant()
    ' Function to get the factors and offsets for converting a UserForm's size to pixel units
    ' Obtains the window size in pixels via Windows API and compares it with the UserForm's design size
    Dim clRect As RECT
    Dim winRect As RECT
    Dim pixClWidth As Long, pixClHeight As Long
    Dim pixWinWidth As Long, pixWinHeight As Long
    Dim pixWidthOffset As Long, pixHeightOffset As Long
    Dim scaleX As Double, scaleY As Double
    Dim hwnd As LongPtr
    Dim originalFrmTitle As String
    Dim tempFrmTitle As String
    Dim results(0 To 3) As Variant
    
    ' To avoid getting the handle of a window with the same name, temporarily change the title to a unique name when obtaining hwnd
    ' Restore the original title immediately after obtaining hwnd
    originalFrmTitle = frm.caption
    tempFrmTitle = "TempName_" & GenerateUUIDv4()
    frm.caption = tempFrmTitle
    hwnd = Win32_FindWindowW("", tempFrmTitle)
    frm.caption = originalFrmTitle
    
    If CLng(hwnd) = 0 Then
        Err.Raise Number:=513, Description:="Failed to get HWND."
    End If
    
    ' Get the actual client area size
    GetClientRect hwnd, clRect
    pixClWidth = clRect.Right - clRect.Left
    pixClHeight = clRect.Bottom - clRect.Top
    
    ' Get the difference in X and Y between the actual window size and the client area size
    GetWindowRect hwnd, winRect
    pixWinWidth = winRect.Right - winRect.Left
    pixWinHeight = winRect.Bottom - winRect.Top
    pixWidthOffset = pixWinWidth - pixClWidth
    pixHeightOffset = pixWinHeight - pixClHeight
    
    ' Twips -> pixel conversion factors
    scaleX = pixClWidth / frm.InsideWidth
    scaleY = pixClHeight / frm.InsideHeight
    
    ' If horizontal and vertical scales are almost the same, return the average
    If Abs(scaleX - scaleY) < 0.01 Then
        results(0) = (scaleX + scaleY) / 2
        results(1) = (scaleX + scaleY) / 2
    Else
        ' If there is a difference between horizontal and vertical scales
        results(0) = scaleX
        results(1) = scaleY
    End If
    results(2) = pixWidthOffset
    results(3) = pixHeightOffset
    GetUserFormScaleFactorsAndOffsets = results
End Function

Private Function UserFormSizeToPixel(ByVal ufSize As Double, ByVal factor As Double) As Long
    ' Function to convert the size of a UserForm or control to pixels
    UserFormSizeToPixel = Round(ufSize * factor)
End Function

Private Function GenerateUUIDv4() As String
    Dim i As Long
    Dim b(15) As Byte
    Dim s As String
    Dim hexStr As String
    
    ' Initialize random number generator
    Randomize
    
    ' Generate 16 bytes of random values
    For i = 0 To 15
        b(i) = Int(Rnd() * 256)
    Next i
    
    ' Set version (4) (set bits 7-4 to 0100)
    b(6) = (b(6) And &HF) Or &H40
    
    ' Set variant (10xx)
    b(8) = (b(8) And &H3F) Or &H80
    
    ' Convert the 16 bytes to a string (with hyphen format)
    hexStr = ""
    For i = 0 To 15
        hexStr = hexStr & Right$("0" & Hex(b(i)), 2)
        Select Case i
            Case 3, 5, 7, 9
                hexStr = hexStr & "-"
        End Select
    Next i
    
    GenerateUUIDv4 = LCase$(hexStr)
End Function

Private Sub SaveUTF8BOMText(ByVal filePath As String, ByVal textData As String)
    ' Save the specified string as UTF-8 without BOM
    Dim stream As Object
    Dim bytes() As Byte
    
    ' Normalize line endings
    textData = VBA.Replace(textData, vbCrLf, vbLf)
    textData = VBA.Replace(textData, vbCr, vbLf)
    textData = VBA.Replace(textData, vbLf, vbNewLine)
    
    ' Convert to UTF-8 and remove BOM
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text mode
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText textData
    stream.position = 0
    stream.Type = 1 ' Switch to binary mode
    bytes = stream.Read
    stream.Close
    Set stream = Nothing
    
    ' Save file in binary mode
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.SaveToFile filePath, 2
    stream.Close
    Set stream = Nothing
End Sub

Private Sub SaveUTF8Text_NoBOM(ByVal filePath As String, ByVal textData As String)
    ' Save the specified string as UTF-8 without BOM
    Dim stream As Object
    Dim bytes() As Byte
    
    ' Normalize line endings
    textData = VBA.Replace(textData, vbCrLf, vbLf)
    textData = VBA.Replace(textData, vbCr, vbLf)
    textData = VBA.Replace(textData, vbLf, vbNewLine)
    
    ' Convert to UTF-8 and remove BOM
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text mode
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText textData
    stream.position = 0
    stream.Type = 1 ' Switch to binary mode
    bytes = stream.Read
    stream.Close
    Set stream = Nothing
    
    ' Remove BOM if present
    If UBound(bytes) >= 2 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            bytes = MidB(bytes, 4) ' Remove BOM (EF BB BF)
        End If
    End If
    
    ' Save file in binary mode
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.SaveToFile filePath, 2
    stream.Close
    Set stream = Nothing
End Sub

Private Function GetPrimaryMonitorDPI() As Variant()
    Dim hdc As LongPtr
    Dim dpiX As Long, dpiY As Long
    Dim results(0 To 1) As Variant
    Const LOGPIXELSX As Long = 88 ' Horizontal DPI
    Const LOGPIXELSY As Long = 90 ' Vertical DPI
    
    ' Get device context for the entire screen
    hdc = GetDC(0)
    
    ' Get horizontal and vertical DPI
    dpiX = GetDeviceCaps(hdc, LOGPIXELSX)
    dpiY = GetDeviceCaps(hdc, LOGPIXELSY)
    
    ' Release the device context
    ReleaseDC 0, hdc
    
    results(0) = dpiX
    results(1) = dpiY
    
    ' Return DPI
    GetPrimaryMonitorDPI = results
End Function

Private Function GenerateUnsupportedControlMessage(ByVal ctrl As Object) As String
    Const q As String = """"
    GenerateUnsupportedControlMessage = "Control type " & q & TypeName(ctrl) & q & " is not supported."
End Function

Private Function GenerateUnavailableNameMessage(ByVal ctrl As Object) As String
    Const q As String = """"
    GenerateUnavailableNameMessage = "Object Name " & q & ctrl.Name & q & "is not available." & vbLf & "Please use a different name instead."
End Function

Private Function GetFormControlDepth(ByVal ctrl As Object) As Long
    ' Get the hierarchy depth of the control
    Dim depth As Long
    Dim temp As Variant
    depth = 0
    Set temp = ctrl
    Do While True
        If depth Mod 10 = 0 Then DoEvents
        On Error GoTo Finally
        Set temp = temp.Parent
        depth = depth + 1
        On Error GoTo 0
    Loop
Finally:
    
    If Err.Number <> 438 Then
        Err.Raise Number:=Err.Number
    End If
    
    GetFormControlDepth = depth
    
End Function

Private Function SortFormControlsByDepth(ByVal frmControls As Variant) As Collection
    ' Sort the list of UserForm controls in ascending order of hierarchy depth
    Dim tempColl As Collection
    Set tempColl = New Collection
    Dim sortedColl As Collection
    Set sortedColl = New Collection
    Dim ctrl As Variant
    Dim tempArray() As Variant
    Dim depth As Long
    Dim item As Variant
    For Each ctrl In frmControls
        depth = GetFormControlDepth(ctrl)
        tempColl.Add Array(depth, ctrl)
    Next ctrl
    If tempColl.Count > 0 Then
        tempArray = Collection2Array(tempColl)
        Call InsertionSortJaggedArray(tempArray, reverse:=False)
        For Each item In tempArray
            sortedColl.Add item(1)
        Next item
    End If
    Set SortFormControlsByDepth = sortedColl
End Function

Private Function Collection2Array(ByVal coll As Collection, Optional ByVal isStartIdx1 As Boolean = False) As Variant()
    ' Convert a Collection to an array
    ' If isStartIdx1 is True, create an array starting from index 1 (to match Collection numbering)
    Dim arr() As Variant
    Dim item As Variant
    Dim idx As Long
    If coll.Count > 0 Then
        If isStartIdx1 Then
            ReDim arr(1 To coll.Count)
        Else
            ReDim arr(0 To coll.Count - 1)
        End If
        idx = LBound(arr)
        For Each item In coll
            ' Use "Set" when assigning objects.
            If IsObject(item) Then
                Set arr(idx) = item
            Else
                arr(idx) = item
            End If
            idx = idx + 1
        Next
    Else
        arr = Array()
    End If
    Collection2Array = arr
End Function

Private Function Array2Collection(ByVal arr As Variant) As Collection
    ' Convert an array to a collection
    ' ArrayLength (Function) is dependency
    Dim coll As New Collection
    Dim i As Long
    
    If Not IsArray(arr) Then
        Err.Raise Number:=13
        Exit Function
    End If
    
    If ArrayLength(arr) > 0 Then
        For i = LBound(arr) To UBound(arr)
            coll.Add arr(i)
        Next i
    End If
    Set Array2Collection = coll
End Function

Private Function ArrayLength(ByVal arr As Variant) As Long
    ' Return the number of items in an array
    ' arr: Array to measure length
    ' if an array is empty (not initialized), return 0
    Dim temp As Variant
    If Not IsArray(arr) Then
        Err.Raise Number:=13
        Exit Function
    End If
    
    On Error GoTo Exception
    temp = LBound(arr)
    On Error GoTo 0
    
    ArrayLength = UBound(arr) + (1 - LBound(arr))
    Exit Function
Exception:
    ' Empty (not initialized) array
    If Err.Number <> 9 Then
        Err.Raise Number:=Err.Number
        Exit Function
    End If
    ArrayLength = 0
End Function

Private Sub InsertionSortJaggedArray(ByRef arr As Variant, _
    Optional ByVal reverse As Boolean = False, _
    Optional ByVal strSort As Boolean = False, _
    Optional ByVal ignoreCase As Boolean = True)
    
    ' Sorts a jagged array using the Insertion Sort algorithm based on the first element of each nested array.
    '   e.g., [[1, "A"], [3, "B"], [2, "C"]] -> [[1, "A"], [2, "C"], [3, "B"]]
    '   Does not affect the relative order of items with the same numeric value
    '   e.g., [[3, "C"], [3, "A"], [1, "A"], [3, "B"]] -> [[1, "A"], [3, "C"], [3, "A"], [3, "B"]]
    ' reverse: Set to True for descending order.
    '   e.g., [[1, "A"], [3, "B"], [2, "C"]] -> [[3, "B"], [2, "C"], [1, "A"]]
    ' strSort: Set to True for string-based comparison, False for numeric comparison.
    ' ignoreCase: Valid only when strSort is True. Set to True to perform case-insensitive comparison.
    ' Dependency: DynamicCompare
    If Not IsArray(arr) Then Err.Raise Number:=13
    Dim minIndex As Long
    Dim maxIndex As Long
    Dim idxToRef1 As Long
    Dim idxToRef2 As Long
    Dim op As String
    
    If reverse Then
        op = "<"
    Else
        op = ">"
    End If
    
    minIndex = LBound(arr)
    maxIndex = UBound(arr)
    Dim i As Long, j As Long
    Dim swap As Variant
    For i = minIndex + 1 To maxIndex
        swap = arr(i)
        For j = i - 1 To minIndex Step -1
            idxToRef1 = LBound(arr(j))
            idxToRef2 = LBound(swap)
            If DynamicCompare(arr(j)(idxToRef1), swap(idxToRef2), op, strSort, ignoreCase) Then
                arr(j + 1) = arr(j)
            Else
                Exit For
            End If
        Next
        arr(j + 1) = swap
    Next
End Sub


Private Function DynamicCompare(ByVal a As Variant, ByVal b As Variant, ByVal op As String, _
    Optional ByVal shouldStrComp As Boolean = False, Optional ByVal ignoreCase As Boolean = True) As Boolean
    ' Performs dynamic comparison using a string representation of an operator.
    ' a, b: Values to compare.
    ' op: Comparison operator as a string (">", ">=", "<", "<=", "=", "<>").
    ' shouldStrComp: Set to True for string comparison mode, False for numeric/default comparison.
    ' ignoreCase: Valid only when shouldStrComp is True. Set to True to ignore case sensitivity.
    Dim result As Boolean
    Dim compareMode As VbCompareMethod
    
    If shouldStrComp Then
        If ignoreCase Then
            compareMode = vbTextCompare
        Else
            compareMode = vbBinaryCompare
        End If
        
        Select Case op
            Case ">"
                result = StrComp(a, b, compareMode) > 0
            Case ">="
                result = StrComp(a, b, compareMode) >= 0
            Case "<"
                result = StrComp(a, b, compareMode) < 0
            Case "<="
                result = StrComp(a, b, compareMode) <= 0
            Case "="
                result = StrComp(a, b, compareMode) = 0
            Case "<>"
                result = StrComp(a, b, compareMode) <> 0
            Case Else
                Err.Raise vbObjectError, , "Unknown operator: " & op
        End Select
    Else
        Select Case op
            Case ">"
                result = (a > b)
            Case ">="
                result = (a >= b)
            Case "<"
                result = (a < b)
            Case "<="
                result = (a <= b)
            Case "="
                result = (a = b)
            Case "<>"
                result = (a <> b)
            Case Else
                Err.Raise vbObjectError, , "Unknown operator: " & op
        End Select
    End If
    DynamicCompare = result
End Function

Private Function CollContainsKey(ByVal coll As Collection, ByVal strKey As String) As Boolean
    ' Check if a specific key exists in the Collection
    CollContainsKey = False
    If coll Is Nothing Then Exit Function
    If coll.Count = 0 Then Exit Function
     
    On Error GoTo Exception
    Call coll.item(strKey)
    On Error GoTo 0
    CollContainsKey = True
    
    Exit Function
Exception:
    CollContainsKey = False
    Exit Function
End Function

Private Function ReverseArray(ByVal srcArr As Variant) As Variant
    Dim newArr As Variant: ReDim newArr(LBound(srcArr) To UBound(srcArr))
    Dim newIdx As Long: newIdx = LBound(newArr)
    Dim i As Long: For i = UBound(srcArr) To LBound(srcArr) Step -1
        If IsObject(srcArr(i)) Then
            Set newArr(newIdx) = srcArr(i)
        Else
            newArr(newIdx) = srcArr(i)
        End If
        newIdx = newIdx + 1
    Next
    ReverseArray = newArr
End Function


Private Function ReverseCollection(ByVal srcColl As Collection) As Collection
    Dim resultColl As Collection
    Dim arr() As Variant
    If srcColl.Count > 0 Then
        arr = Collection2Array(srcColl)
        arr = ReverseArray(arr)
        Set resultColl = Array2Collection(arr)
    Else
        Set resultColl = New Collection
    End If
    Set ReverseCollection = resultColl
End Function


