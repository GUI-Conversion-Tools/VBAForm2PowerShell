Attribute VB_Name = "VBAForm2PowerShell"

' VBAForm2PowerShell v1.2.0
' https://github.com/GUI-Conversion-Tools/VBAForm2PowerShell
' Copyright (c) 2025-2026 ZeeZeX
' This software is released under the MIT License.
' https://opensource.org/licenses/MIT

Option Explicit


#If VBA7 Then
    ' 64bit Office / VBA7 or later
    Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
#Else
    ' 32bit Office
    Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Const FORM_WINDOW_NAME As String = "window"

Public Sub TestRunConversion2PS()
    Call ConvertForm2PS(UserForm1)
End Sub

Public Sub TestRunConversion2PS_2()
    Call ConvertForm2PS(Array(UserForm1, UserForm2))
End Sub

Public Sub TestRunConversion2PS_3()
    Call ConvertForm2PS(UserForm1, saveAsBat:=True, useCls:=True)
End Sub

Public Sub ConvertForm2PS(ByVal frms As Variant, Optional ByVal saveAsBat As Boolean = False, Optional ByVal useCls As Boolean = False, Optional ByVal noMainLoop As Boolean = False)
    
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
    code = GeneratePSWinFormsCode(frms, useCls, noMainLoop)
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


Public Function GeneratePSWinFormsCode(ByVal frms As Variant, Optional ByVal useCls As Boolean = False, Optional ByVal noMainLoop As Boolean = False) As String
    Dim root As Variant
    Dim indent As String
    Dim prefix As String
    Dim clsNumber As Long
    Dim formName As String
    Dim controlVarName As String
    Dim parentVarName As String
    Dim childVarName As String
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
    Dim widgetType As String
    Dim styleName As String
    Dim pixelWidth As Long
    Dim pixelHeight As Long
    Dim pixelTop As Long
    Dim pixelLeft As Long
    Dim i As Long
    Dim orientation As String
    Dim cursorType As String
    Dim caption As String
    Dim colorSetting As String
    Dim picSizeMode As String
    Dim tabPosition As String
    Dim objHeaders As Object
    Dim treeviewNodes As Collection
    Dim node As Object
    Dim nodeDictName As String
    Dim nodeVarName As String
    Dim nodeParentVarName As String
    Dim enableScrollBar As Boolean
    
    r = ""
    
    If IsArray(frms) Then
        useCls = True
    Else
        frms = VBA.Array(frms)
    End If
    
    If useCls Then
        indent = "        "
        prefix = "$this."
    Else
        indent = ""
        prefix = "$"
    End If
    
    r = r & "$ErrorActionPreference = 'Stop'" & vbLf
    r = r & "Add-Type -AssemblyName System.Windows.Forms" & vbLf
    r = r & "Add-Type -AssemblyName System.Drawing" & vbLf
    r = r & "[System.Windows.Forms.Application]::EnableVisualStyles()" & vbLf
    r = r & vbLf
    
    For Each root In frms
        unavailableNames = VBA.Array("", "System", "item", "row")
        
        For i = LBound(unavailableNames) To UBound(unavailableNames)
            unavailableNames(i) = LCase(unavailableNames(i))
        Next
        
        If ContainsValue(unavailableNames, LCase(root.Name)) Then
            MsgBox GenerateUnavailableNameMessage(root)
            r = ""
            GeneratePSWinFormsCode = r
            Exit Function
        End If
        unavailableNames(0) = LCase(FORM_WINDOW_NAME)
        
        pixelWidth = UserFormSizeToPixel(root.InsideWidth)
        pixelHeight = UserFormSizeToPixel(root.InsideHeight)
        
        formName = GenerateCtrlVarName(root, prefix, useCls)
        
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
                r = r & "    " & "[object]" & GenerateCtrlVarName(ctrl, "$", False) & vbLf
                If ContainsValue(Array("ComboBox", "ListBox"), TypeName(ctrl)) Or IsListView(ctrl) Then
                    itemsListName = GenerateCtrlVarName(ctrl, "$", False) & "_items_value"
                    r = r & "    " & "[object]" & itemsListName & vbLf
                End If
                
                If TypeName(ctrl) = "MultiPage" Then
                    For Each item In ctrl.Pages
                        r = r & "    " & "[object]" & GenerateCtrlVarName(item, "$", False) & vbLf
                    Next
                End If
                
                If IsListView(ctrl) Then
                    Set objHeaders = ctrl.ColumnHeaders
                    i = 0
                    For Each item In objHeaders
                        i = i + 1
                        r = r & "    " & "[object]" & GenerateCtrlVarName(ctrl, "$", False) & "_col" & i & vbLf
                    Next
                End If
                
                If IsTreeView(ctrl) Then
                    r = r & "    " & "[object]" & GenerateCtrlVarName(ctrl, "$", False) & "_node_dict" & vbLf
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
        r = r & indent & formName & ".FormBorderStyle = " & q & "FixedSingle" & q & vbLf  ' Disable window resizing
        r = r & indent & formName & ".BackColor = " & q & FormColorToHex(root.BackColor) & q & vbLf
        r = r & indent & formName & ".AutoScaleMode = " & q & "None" & q & vbLf
        
        
        cursorType = GetControlCursorType(root)
        If cursorType <> "" Then
            r = r & indent & formName & ".Cursor = " & q & cursorType & q & vbLf
        End If
        
        r = r & vbLf

        For Each ctrl In ctrls
            If GetWinFormsControlName(ctrl) = "" Then
                MsgBox GenerateUnsupportedControlMessage(ctrl)
                r = ""
                GeneratePSWinFormsCode = r
                Exit Function
            End If
            
            If ContainsValue(unavailableNames, LCase(ctrl.Name)) Then
                MsgBox GenerateUnavailableNameMessage(ctrl)
                r = ""
                GeneratePSWinFormsCode = r
                Exit Function
            End If
            
            widgetType = "System.Windows.Forms." & GetWinFormsControlName(ctrl)
            
            pixelLeft = UserFormSizeToPixel(ctrl.Left)
            pixelTop = UserFormSizeToPixel(ctrl.Top)
            pixelWidth = UserFormSizeToPixel(ctrl.Width)
            pixelHeight = UserFormSizeToPixel(ctrl.Height)
            
            controlVarName = GenerateCtrlVarName(ctrl, prefix, useCls)
            parentVarName = GenerateCtrlVarName(ctrl.Parent, prefix, useCls)
            itemsListName = controlVarName & "_items_value"
            enableScrollBar = False
            
            r = r & indent & controlVarName & " = " & "New-Object" & " " & widgetType & vbLf
            r = r & indent & parentVarName & ".Controls.Add(" & controlVarName & ")" & vbLf
            r = r & indent & controlVarName & ".Location = New-Object System.Drawing.Point(" & pixelLeft & ", " & pixelTop & ")" & vbLf
            r = r & indent & controlVarName & ".Size = New-Object System.Drawing.Size(" & pixelWidth & ", " & pixelHeight & ")" & vbLf
            
            If GetWinFormsControlName(ctrl) = "GroupBox" Or ContainsValue(Array("Label", "CommandButton", "TextBox", "SpinButton", "ListBox", "CheckBox", "ToggleButton", "OptionButton", "ComboBox"), TypeName(ctrl)) Or IsListView(ctrl) Then
                ' Set ForeColor
                r = r & indent & controlVarName & ".ForeColor = " & q & FormColorToHex(ctrl.ForeColor) & q & vbLf
            End If
            
            If ContainsValue(Array("Label", "CommandButton", "Frame", "TextBox", "SpinButton", "ListBox", "CheckBox", "ToggleButton", "OptionButton", "Image", "ComboBox"), TypeName(ctrl)) Or IsListView(ctrl) Then
                ' Set BackColor
                colorSetting = q & FormColorToHex(ctrl.BackColor) & q
                If ContainsValue(Array("Label", "TextBox", "CommandButton", "CheckBox", "ToggleButton", "OptionButton", "Image", "ComboBox"), TypeName(ctrl)) Then
                    If ctrl.BackStyle = fmBackStyleTransparent Then
                        If Not ContainsValue(Array("TextBox", "ComboBox", "ToggleButton"), TypeName(ctrl)) Then
                            colorSetting = q & "Transparent" & q
                        Else
                            ' Apply the BackColor of the parent control because TextBox and ComboBox do not support "Transparent"
                            ' CheckBox with Appearance = "Button" also does not support "Transparent" when it is focused and clicked (pressed)
                            If TypeName(ctrl.Parent) <> "Page" Then
                                colorSetting = q & FormColorToHex(ctrl.Parent.BackColor) & q
                            Else
                                ' Because the Page control does not have a BackColor property, set the color to &H8000000F&, which matches the background color of the Page
                                colorSetting = q & FormColorToHex(&H8000000F) & q
                            End If
                        End If
                    End If
                End If
                r = r & indent & controlVarName & ".BackColor = " & colorSetting & vbLf
                
            End If
            
            
            If GetWinFormsControlName(ctrl) = "GroupBox" Or ContainsValue(Array("Label", "CommandButton", "CheckBox", "ToggleButton", "OptionButton"), TypeName(ctrl)) Then
                caption = ctrl.caption
                caption = Convert2PowerShellFormatText(caption)
                r = r & indent & controlVarName & ".Text = " & q & caption & q & vbLf
            End If
            
            If ContainsValue(Array("CheckBox", "OptionButton"), TypeName(ctrl)) Then
                If ctrl.Alignment = fmAlignmentLeft Then
                    r = r & indent & controlVarName & ".RightToLeft = " & q & "Yes" & q & vbLf
                End If
            End If
            
            If TypeName(ctrl) = "ToggleButton" Then
                r = r & indent & controlVarName & ".Appearance = " & q & "Button" & q & vbLf
                r = r & indent & controlVarName & ".FlatStyle = " & q & "Flat" & q & vbLf
            End If
            
            If TypeName(ctrl) = "CommandButton" Then
                r = r & indent & controlVarName & ".FlatStyle = " & q & "Popup" & q & vbLf
            End If
            
            If TypeName(ctrl) = "TextBox" Then
                r = r & indent & controlVarName & ".Text = " & q & Convert2PowerShellFormatText(ctrl.text) & q & vbLf
                r = r & indent & controlVarName & ".Multiline = " & "$" & LCase(CBool(ctrl.Multiline)) & vbLf
                r = r & indent & controlVarName & ".WordWrap = " & "$" & LCase(CBool(ctrl.WordWrap)) & vbLf
                
                If ctrl.PasswordChar <> "" Then
                    r = r & indent & controlVarName & ".PasswordChar = " & q & Left(ctrl.PasswordChar, 1) & q & vbLf
                End If
                
                If ctrl.Locked Then
                    r = r & indent & controlVarName & ".ReadOnly = " & "$true" & vbLf
                End If
                
                Select Case ctrl.ScrollBars
                    Case fmScrollBarsHorizontal
                        r = r & indent & controlVarName & ".ScrollBars = " & q & "Horizontal" & q & vbLf
                    Case fmScrollBarsVertical
                        r = r & indent & controlVarName & ".ScrollBars = " & q & "Vertical" & q & vbLf
                    Case fmScrollBarsBoth
                        r = r & indent & controlVarName & ".ScrollBars = " & q & "Both" & q & vbLf
                End Select
                
            End If
            
            If TypeName(ctrl) = "ComboBox" Then
                r = r & indent & itemsListName & " = " & GetListBoxValue(ctrl, indent) & vbLf
                r = r & indent & controlVarName & ".Items.AddRange(" & itemsListName & ")" & vbLf
                r = r & indent & controlVarName & ".Text = " & q & Convert2PowerShellFormatText(ctrl.text) & q & vbLf
                
                If ctrl.Style = fmStyleDropDownList Then
                    r = r & indent & controlVarName & ".DropDownStyle = " & q & "DropDownList" & q & vbLf
                End If
                
                If ctrl.Locked Then
                    r = r & indent & controlVarName & ".Enabled = " & "$false" & vbLf
                End If
                
            End If
            
            If TypeName(ctrl) = "ListBox" Then
                r = r & indent & itemsListName & " = " & GetListBoxValue(ctrl, indent) & vbLf
                r = r & indent & controlVarName & ".Items.AddRange(" & itemsListName & ")" & vbLf
                
                Select Case ctrl.MultiSelect
                    Case fmMultiSelectMulti
                        r = r & indent & controlVarName & ".SelectionMode = " & q & "MultiSimple" & q & vbLf
                    Case fmMultiSelectExtended
                        r = r & indent & controlVarName & ".SelectionMode = " & q & "MultiExtended" & q & vbLf
                End Select
                
                If ctrl.Locked Then
                    r = r & indent & controlVarName & ".Enabled = " & "$false" & vbLf
                End If
                
            End If
            
            If TypeName(ctrl) = "ScrollBar" Then
                r = r & indent & controlVarName & ".Minimum = " & ctrl.Min & vbLf
                r = r & indent & controlVarName & ".Maximum = " & ctrl.Max & vbLf
            End If
            
            
            ' Set each Caption in MultiPage
            If TypeName(ctrl) = "MultiPage" Then
                
                Select Case ctrl.TabOrientation
                    Case fmTabOrientationTop
                        tabPosition = "Top"
                    Case fmTabOrientationBottom
                        tabPosition = "Bottom"
                    Case fmTabOrientationLeft
                        tabPosition = "Left"
                    Case fmTabOrientationRight
                        tabPosition = "Right"
                    Case Else
                        tabPosition = "Top"
                End Select
                
                r = r & indent & controlVarName & ".Alignment = " & q & tabPosition & q & vbLf
                
                If ctrl.Style = fmTabStyleNone Then
                    r = r & indent & controlVarName & ".Appearance = " & q & "FlatButtons" & q & vbLf
                    r = r & indent & controlVarName & ".ItemSize = New-Object System.Drawing.Size(0, 1)" & vbLf
                    r = r & indent & controlVarName & ".SizeMode = " & q & "Fixed" & q & vbLf
                    r = r & indent & controlVarName & ".TabStop = " & "$false" & vbLf
                End If
                
                For Each item In ctrl.Pages
                    childVarName = GenerateCtrlVarName(item, prefix, useCls)
                    caption = item.caption
                    caption = Convert2PowerShellFormatText(caption)
                    r = r & indent & childVarName & " = New-Object System.Windows.Forms.TabPage" & vbLf
                    r = r & indent & controlVarName & ".Controls.Add(" & childVarName & ")" & vbLf
                    r = r & indent & childVarName & ".BackColor = " & q & FormColorToHex(&H8000000F) & q & vbLf
                    r = r & indent & childVarName & ".Text = " & q & caption & q & vbLf
                Next
            End If
            
            ' Font size is rounded because VBA officially does not support decimal fraction in font settings
            If GetWinFormsControlName(ctrl) = "GroupBox" Or ContainsValue(Array("Label", "CommandButton", "TextBox", "ListBox", "CheckBox", "ToggleButton", "OptionButton", "ComboBox", "MultiPage"), TypeName(ctrl)) Or IsListView(ctrl) Or IsTreeView(ctrl) Then
                fontStyle = ""
                
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
                
                r = r & indent & controlVarName & ".Font = New-Object System.Drawing.Font(" & q & ctrl.Font.Name & q & ", " & Round(ctrl.Font.Size) & fontStyle & ")" & vbLf
            End If
            
            
            If GetWinFormsControlName(ctrl) <> "GroupBox" And ContainsValue(Array("Frame", "TextBox", "Label", "ListBox", "Image"), TypeName(ctrl)) Then
                ' WinForms' Combobox does not support customizing border style
                r = r & indent & controlVarName & GetBorderSetting(ctrl, useCls) & vbLf
            End If
            
            If ContainsValue(Array("Label", "TextBox", "CheckBox", "ToggleButton", "OptionButton"), TypeName(ctrl)) Then
                r = r & indent & controlVarName & GetTextAlignSetting(ctrl, useCls) & vbLf
            End If
            
            ' Set mouse cursor
            If TypeName(ctrl) <> "MultiPage" Then
                cursorType = GetControlCursorType(ctrl)
                If cursorType <> "" Then
                    r = r & indent & controlVarName & ".Cursor = " & q & cursorType & q & vbLf
                End If
            End If
            
            
            If IsListView(ctrl) Then
                r = r & indent & controlVarName & ".View = " & q & "Details" & q & vbLf
                r = r & indent & controlVarName & ".FullRowSelect = " & "$true" & vbLf
                r = r & indent & controlVarName & ".GridLines = " & "$true" & vbLf
                r = r & indent & controlVarName & ".MultiSelect = " & "$" & LCase(CBool(ctrl.MultiSelect)) & vbLf
                r = r & DefineListViewColumns(ctrl, indent, prefix, useCls) & vbLf
                r = r & indent & itemsListName & " = " & GetListViewItems(ctrl, indent) & vbLf
                r = r & indent & "foreach ($row in " & itemsListName & ") {" & vbLf
                r = r & indent & "    $item = New-Object System.Windows.Forms.ListViewItem($row[0])" & vbLf
                r = r & indent & "    for ($i = 1; $i -lt $row.Length; $i++) {" & vbLf
                r = r & indent & "        [void]$item.SubItems.Add($row[$i])" & vbLf
                r = r & indent & "    }" & vbLf
                r = r & indent & "    [void]" & controlVarName & ".Items.Add($item)" & vbLf
                r = r & indent & "}" & vbLf
            End If
            
            If IsTreeView(ctrl) Then
                nodeDictName = controlVarName & "_node_dict"
                
                If HasScrollProperty(ctrl) Then
                    enableScrollBar = ctrl.Scroll
                Else
                    enableScrollBar = True
                End If
                r = r & indent & controlVarName & ".Scrollable = " & "$" & LCase(CBool(enableScrollBar)) & vbLf
                
                Set treeviewNodes = GetAllTreeViewNodesBfs(ctrl)
                r = r & indent & nodeDictName & " = @{}" & vbLf
                For Each item In treeviewNodes
                    Set node = item(0)
                    nodeVarName = nodeDictName & "[" & q & Convert2PowerShellFormatText(node.Key) & q & "]"
                    If node.Parent Is Nothing Then
                        nodeParentVarName = controlVarName
                    Else
                        nodeParentVarName = nodeDictName & "[" & q & Convert2PowerShellFormatText(node.Parent.Key) & q & "]"
                    End If
                    r = r & indent & nodeVarName & " = " & nodeParentVarName & ".Nodes.Add(" & q & Convert2PowerShellFormatText(node.text) & q & ")" & vbLf
                    If node.Expanded Then
                        r = r & indent & nodeVarName & ".Expand() | Out-Null" & vbLf
                    End If
                Next item
            End If
            
            If TypeName(ctrl) = "Image" Then
                
                Select Case ctrl.PictureSizeMode
                    Case fmPictureSizeModeClip
                        Select Case ctrl.PictureAlignment
                            Case fmPictureAlignmentCenter
                                picSizeMode = "CenterImage"
                            Case Else
                                picSizeMode = "Normal"
                        End Select
                    Case fmPictureSizeModeStretch
                        picSizeMode = "StretchImage"
                    Case fmPictureSizeModeZoom
                        picSizeMode = "Zoom"
                End Select
                
                r = r & indent & "#" & controlVarName & ".Image = " & DotNetTypeLiteral("System.Drawing.Image", useCls) & "::FromFile(" & q & "C:\path\to\your\image.png" & q & ")" & vbLf
                r = r & indent & "#" & controlVarName & ".SizeMode = " & q & picSizeMode & q & vbLf
            End If
            
            r = r & vbLf
                
        Next ctrl
        r = r & SetWinFormsButtonValues(ctrls, indent, prefix, useCls) & vbLf
        If Not useCls And Not noMainLoop Then
            r = r & formName & ".ShowDialog() | Out-Null"
        End If
        
        If useCls Then
            r = r & "    " & "}" & vbLf
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
    
    GeneratePSWinFormsCode = r
End Function

Private Function GenerateCtrlVarName(ByVal ctrl As Object, ByVal prefix As String, ByVal useCls As Boolean) As String
    ' Generates a valid, unique identifier for a control in the target language.
    Dim controlVarName As String
    If IsRootForm(ctrl) And useCls Then
        controlVarName = prefix & FORM_WINDOW_NAME
    Else
        If TypeName(ctrl) = "Page" Then
        ' VBA allows duplicate names for Page objects if they belong to different MultiPage controls.
        ' To ensure unique variable names in the target language (which typically uses a flat
        ' namespace), namespace the Page by prepending its parent MultiPage's name.
        ' Example: "Page1" inside "MultiPage1" becomes "MultiPage1_Page1"
            controlVarName = prefix & ctrl.Parent.Name & "_" & ctrl.Name
        Else
            controlVarName = prefix & ctrl.Name
        End If
    End If
    GenerateCtrlVarName = controlVarName
End Function

Private Function IsRootForm(ByVal ctrl As Object) As Boolean
    ' Determines whether the specified control is the root UserForm.
    '
    ' This function returns True only when:
    '   - The control is of type MSForms.UserForm, and
    '   - The control exists at the top level (i.e., its hierarchy depth is 0).
    '
    ' Note:
    '   Even if the control is of type MSForms.UserForm, this function will return False
    '   if the control is not the root window (for example, if it is nested or owned
    '   within another container or context).
    Dim result As Boolean
    If GetFormControlDepth(ctrl) = 0 And TypeOf ctrl Is MSForms.UserForm Then
        result = True
    Else
        result = False
    End If
    IsRootForm = result
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
        Case fmBorderStyleSingle
            ' SpecialEffect is fmSpecialEffectFlat if BorderStyle is fmBorderStyleSingle
            borderSetting = "FixedSingle"
        Case fmBorderStyleNone
            Select Case ctrl.SpecialEffect
                Case fmSpecialEffectFlat
                    borderSetting = "None"
                Case fmSpecialEffectRaised
                    borderSetting = "Fixed3D"
                Case fmSpecialEffectSunken
                    borderSetting = "Fixed3D"
                Case fmSpecialEffectEtched
                    borderSetting = "FixedSingle"
                Case fmSpecialEffectBump
                    borderSetting = "FixedSingle"
            End Select
    End Select

    r = ".BorderStyle = " & q & borderSetting & q
    GetBorderSetting = r
End Function

Private Function GetTextAlignSetting(ByVal ctrl As Object, ByVal useCls As Boolean) As String
    Dim r As String
    Const q As String = """"
    Dim position As String
    r = ""
    
    If TypeName(ctrl) = "TextBox" Then
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
    ElseIf ContainsValue(Array("CheckBox", "OptionButton", "ToggleButton"), TypeName(ctrl)) Then
        Select Case ctrl.TextAlign
            Case fmTextAlignLeft
                position = "MiddleLeft"
            Case fmTextAlignCenter
                position = "MiddleCenter"
            Case fmTextAlignRight
                position = "MiddleRight"
            Case Else
                position = "MiddleCenter"
        End Select
    Else
        Select Case ctrl.TextAlign
            Case fmTextAlignLeft
                position = "TopLeft"
            Case fmTextAlignCenter
                position = "TopCenter"
            Case fmTextAlignRight
                position = "TopRight"
            Case Else
                position = "TopCenter"
        End Select
    End If
    
    r = r & ".TextAlign = " & q & position & q
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
                Case fmOrientationAuto
                    If ctrl.Width > ctrl.Height Then
                        r = "HScrollBar"
                    Else
                        r = "VScrollBar"
                    End If
                    
                Case fmOrientationVertical
                    r = "VScrollBar"
                Case fmOrientationHorizontal
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
            
            If IsListView(ctrl) Then
                r = "ListView"
            End If
            
            If IsTreeView(ctrl) Then
                r = "TreeView"
            End If
            
    End Select
    GetWinFormsControlName = r
End Function

Private Function IsListView(ByVal ctrl As Object) As Boolean
    ' Since the class name of the ListView may vary depending on the version, so use InStr to check it.
    ' e.g ListView/ListView2/ListView3/ListView4
    If InStr(TypeName(ctrl), "ListView") = 1 Then
        IsListView = True
    Else
        IsListView = False
    End If
End Function

Private Function IsTreeView(ByVal ctrl As Object) As Boolean
    ' Since the class name of the TreeView may vary depending on the version, so use InStr to check it.
    ' e.g TreeView/TreeView2/TreeView3/TreeView4
    If InStr(TypeName(ctrl), "TreeView") = 1 Then
        IsTreeView = True
    Else
        IsTreeView = False
    End If
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
    Dim controlVarName As String
    Dim value As Boolean
    Dim r As String
    r = ""
    For Each ctrl In ctrls
        controlVarName = GenerateCtrlVarName(ctrl, prefix, useCls)
        If ContainsValue(Array("OptionButton", "CheckBox", "ToggleButton"), TypeName(ctrl)) Then
            r = r & indent & controlVarName & ".Checked = " & "$" & LCase(CBool(ctrl.value)) & vbLf
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

Private Function DefineListViewColumns(ByVal ctrl As Object, ByVal indent As String, ByVal prefix As String, ByVal useCls As Boolean) As String
    ' Generate code for the ListView headers.
    ' Example:
    ' $ListView1_col1 = New-Object System.Windows.Forms.ColumnHeader
    ' $ListView1_col1.Text = "Header1"
    ' $ListView1_col1.Width = 133
    ' $ListView1_col1.TextAlign = "Left"

    ' $ListView1_col2 = New-Object System.Windows.Forms.ColumnHeader
    ' $ListView1_col2.Text = "Header2"
    ' $ListView1_col2.Width = 133
    ' $ListView1_col2.TextAlign = "Left"

    ' $ListView1_col3 = New-Object System.Windows.Forms.ColumnHeader
    ' $ListView1_col3.Text = "Header3"
    ' $ListView1_col3.Width = 133
    ' $ListView1_col3.TextAlign = "Left"

    ' $listView1.Columns.Add($ListView1_col1) | Out-Null
    ' $listView1.Columns.Add($ListView1_col2) | Out-Null
    ' $listView1.Columns.Add($ListView1_col3) | Out-Null
    Dim objHeaders As Object
    Set objHeaders = ctrl.ColumnHeaders
    Dim controlVarName As String
    Dim i As Long
    Dim item As Variant
    Dim r As String
    Dim colVarName As String
    Dim headerText As String
    Dim colWidth As Long
    Dim colAlign As String
    Const q As String = """"
    controlVarName = GenerateCtrlVarName(ctrl, prefix, useCls)
    r = ""
    i = 0
    For Each item In objHeaders
        i = i + 1
        colVarName = controlVarName & "_col" & i
        headerText = Convert2PowerShellFormatText(item.text)
        colAlign = GetPSWinFormsListViewAlignment(item)
        colWidth = UserFormSizeToPixel(item.Width)
        r = r & indent & colVarName & " = " & "New-Object System.Windows.Forms.ColumnHeader" & vbLf
        r = r & indent & colVarName & ".Text = " & q & headerText & q & vbLf
        r = r & indent & colVarName & ".Width = " & colWidth & vbLf
        r = r & indent & colVarName & ".TextAlign = " & q & colAlign & q & vbLf
    Next item
    
    
    i = 0
    For Each item In objHeaders
        i = i + 1
        colVarName = controlVarName & "_col" & i
        r = r & indent & controlVarName & ".Columns.Add(" & colVarName & ") | Out-Null" & vbLf
    Next item
    
    DefineListViewColumns = r
End Function


Private Function GetPSWinFormsListViewAlignment(ByVal objLvHeader As Object) As String
    ' Header Alignment(ListView) -> ColumnHeader.TextAlign(PowerShell WinForms ListView)
    Const cnsLvwColumnLeft As Long = 0
    Const cnsLvwColumnRight As Long = 1
    Const cnsLvwColumnCenter As Long = 2
    Dim result As String
    Select Case objLvHeader.Alignment
        Case cnsLvwColumnLeft
            result = "Left"
        Case cnsLvwColumnRight
            result = "Right"
        Case cnsLvwColumnCenter
            result = "Center"
        Case Else
            result = "Left"
    End Select
    GetPSWinFormsListViewAlignment = result
End Function

Private Function GetListViewItems(ByVal ctrl As Object, ByVal indent As String) As String
    ' Retrieve the items of a ListView as a string in the format:
    ' @(
    '     ,@("Item1-1", "Item1-2", "Item1-3")
    '     ,@("Item2-1", "Item2-2", "Item2-3")
    '     ,@("Item3-1", "Item3-2", "Item3-3")
    ' )
    Dim item As Object
    Dim i As Long
    Dim coll As Collection
    Dim resultColl As New Collection
    Dim arr() As Variant
    Dim r As String
    Const arrayLiteralStart As String = "@("
    Const arrayLiteralEnd As String = ")"
    Const q As String = """"
    r = ""
    For Each item In ctrl.ListItems
        Set coll = New Collection
        coll.Add q & Convert2PowerShellFormatText(item.text) & q
        
        ' Because older versions of the ListView control do not support For Each for SubItems, use index-based access.
        For i = 1 To ctrl.ColumnHeaders.Count - 1
            coll.Add q & Convert2PowerShellFormatText(item.SubItems(i)) & q
        Next
        
        arr = Collection2Array(coll)
        resultColl.Add "," & arrayLiteralStart & Join(arr, ", ") & arrayLiteralEnd
        
    Next
    
    arr = Collection2Array(resultColl)
    If resultColl.Count > 0 Then
        r = r & arrayLiteralStart & vbLf & indent & "    " & Join(arr, "" & vbLf & indent & "    ") & vbLf & indent & arrayLiteralEnd
    Else
        r = r & arrayLiteralStart & arrayLiteralEnd
    End If
    GetListViewItems = r
End Function

Private Function GetAllTreeViewNodesBfs(ByVal treeviewCtrl As Object) As Collection
    ' This function performs a Breadth-First Search (BFS) on a TreeView control
    ' and returns a collection of nodes along with their hierarchy path.
    ' example:
    ' [[node, "1"], [node, "1-1"], [node, "1-2"], [node, "1-3"], [node, "1-4"], [node, "1-5"], [node, "1-6"], [node, "1-1-1"]]
    Dim queue As Collection
    Dim item As Variant
    Dim node As Object
    Dim child As Object
    Dim hierarchy As String
    Dim childIndex As Long
    Dim resultColl As Collection
    Set resultColl = New Collection
    Set queue = New Collection
    
    Dim nd As Object
    Dim rootIndex As Long
    rootIndex = 1
    
    ' Step 1: Add all root nodes (nodes without parents) to the queue
    ' Each root gets a hierarchy label like "1", "2", etc.
    For Each nd In treeviewCtrl.nodes
        If nd.Parent Is Nothing Then
            queue.Add VBA.Array(nd, CStr(rootIndex))
            rootIndex = rootIndex + 1
        End If
    Next nd
    
    ' Step 2: Perform BFS traversal
    Do While queue.Count > 0
        item = queue(1)
        queue.Remove 1
        
        Set node = item(0)
        hierarchy = item(1)
        ' Add current node and its hierarchy to the result collection
        resultColl.Add VBA.Array(node, hierarchy)
        ' Step 3: Enqueue all children of the current node
        Set child = node.child ' Get first child
        childIndex = 1
        
        Do While Not child Is Nothing
            ' Append child index to hierarchy (e.g., "1-2", "1-2-1")
            queue.Add VBA.Array(child, hierarchy & "-" & childIndex)
            childIndex = childIndex + 1
            Set child = child.Next
        Loop
    Loop
    ' Return the collection of (node, hierarchy) pairs
    Set GetAllTreeViewNodesBfs = resultColl
End Function

Private Function HasScrollProperty(ByVal ctrl As Object) As Boolean
    ' Since the Scroll property does not exist in older versions of TreeView, use this function to check for the property beforehand.
    Dim temp As Variant
    On Error GoTo Exception
    temp = VBA.Array(ctrl.Scroll)
    HasScrollProperty = True
    On Error GoTo 0
    Exit Function
Exception:
    HasScrollProperty = False
End Function

Private Function Convert2PowerShellFormatText(ByVal text As String) As String
    ' Escape special characters in the string
    Dim targetChars() As Variant
    Dim char As Variant
    targetChars = VBA.Array("`", """", "$", "{", "}")
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
    Dim codeColl As New Collection
    Dim codeArray() As Variant

    With codeColl
        .Add ":DUMMY for($i=1;$i -eq 0;$i++) {echo DUMMY} <#"
        .Add "@echo off"
        .Add "chcp 65001 > nul"
        .Add "set ""dirPath=%~dp0"""
        .Add "set ""ME=%~dpnx0"""
        .Add "powershell -ExecutionPolicy Unrestricted -Command ""Set-Location -LiteralPath $env:dirPath; $script = ((Get-Content -LiteralPath $env:ME -Encoding UTF8) -join \""`n\""); $sb = [ScriptBlock]::Create($script); & $sb; if (-not $?) { exit 1 }"""
        .Add "if %ERRORLEVEL% neq 0 ("
        .Add "    pause"
        .Add ") else ("
        .Add "    pause"
        .Add ")"
        .Add "exit /b"
        .Add "#>"
        .Add "# The following is PowerShell code."
    End With
    codeArray = Collection2Array(codeColl)
    code = Join(codeArray, vbLf)
    GenerateBatchCode = code
End Function

Private Function FormColorToHex(ByVal clr As Long) As String
    ' Example:
    ' 16777215 -> "#FFFFFF"
    ' 0 -> "#000000"
    ' &H000000FF& (255) -> "#FF0000"
    ' &H00B4769E& (11826846) -> "#9E76B4"
    ' &H8000000F& (-2147483633) -> "#F0F0F0"(Windows XP[Luna Theme]/10/11), "#D4D0C8"(Windows 2000/XP[Classic Theme])
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

Private Function UserFormSizeToPixel(ByVal ufSize As Double) As Long
    ' Function to convert the size of a UserForm or control to pixels
    ' Excel VBA UserForm dimensions are internally handled as
    ' DPI-independent logical points based on a fixed 96 DPI system.
    ' Therefore, point-to-pixel conversion can be calculated as:
    '     pixel = point * (96 / 72)
    ' and works consistently regardless of the monitor DPI setting.
    Dim pixelSize As Long
    pixelSize = Round(ufSize * (96 / 72))
    UserFormSizeToPixel = pixelSize
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


Private Function GenerateUnsupportedControlMessage(ByVal ctrl As Object) As String
    Const q As String = """"
    GenerateUnsupportedControlMessage = "Control type " & q & TypeName(ctrl) & q & " is not supported."
End Function

Private Function GenerateUnavailableNameMessage(ByVal ctrl As Object) As String
    Const q As String = """"
    GenerateUnavailableNameMessage = "Object Name " & q & ctrl.Name & q & " is not available." & vbLf & "Please use a different name instead."
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
        tempColl.Add VBA.Array(depth, ctrl)
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
        arr = VBA.Array()
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


