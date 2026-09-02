Option Explicit

Sub ButtonCancelTemp_Click()
    Unload SetTempForm
End Sub

Private Sub ButtonAbsTempPath_Click()
    AbsPathTextBox.Text = BrowseFolderPath(AbsPathTextBox.Text)
    AbsPathTextBox.SetFocus
End Sub

Private Sub ButtonEditorPath_Click()
    #If Mac Then
        TextBoxExternalEditor.Text = "open -b " & ShellEscape(MacChooseApp(TextBoxExternalEditor.Text))
    #Else
        TextBoxExternalEditor.Text = BrowseFilePath(TextBoxExternalEditor.Text, "All Files", "*.*")
    #End If
    TextBoxExternalEditor.SetFocus
End Sub

Private Sub ButtonExportToXML_Click()
    Dim FolderPath As String
    Dim FullFilePath As String
    FolderPath = BrowseFolderPath(GetTempPath())
    If Right$(FolderPath, 1) <> PathSep Then
        FolderPath = FolderPath & PathSep
    End If
    FullFilePath = InputBox("Choose file name with .xml extension to export settings under " & FolderPath & ":", _
                            "Export settings to XML", _
                            "IguanaTexSettings.xml")
    If FullFilePath <> vbNullString Then
        FullFilePath = FolderPath & FullFilePath
        WriteSettingsToFile FullFilePath
    End If
End Sub

Private Sub ButtonImportFromXML_Click()
    Dim FullFilePath As String
    MsgBox ("WARNING!! This will *overwrite your IguanaTex settings in the registry*!" & vbCrLf & _
             "To cancel, please click Cancel on the file selection screen.")
    FullFilePath = BrowseFilePath(GetTempPath(), "XML Files", "*.xml")
    If FullFilePath <> GetTempPath() Then
        ReadSettingsFromFileIntoRegistry FullFilePath
        ReadSavedSettings
        SetAbsRelDependencies
    End If
End Sub

Private Sub ButtonLaTeXiTPath_Click()
    TextBoxLaTeXiT.Text = BrowseFilePath(TextBoxLaTeXiT.Text, "All Files", "*.*")
    TextBoxLaTeXiT.SetFocus
End Sub

Private Sub SaveSettings()
    Dim res As String
    
    ' Temp folder
    SetITSetting "AbsOrRel", REG_DWORD, BoolToInt(AbsPathButton.value)
    SetITSetting "Abs Temp Dir", REG_SZ, CStr(AbsPathTextBox.Text)
    If Left$(RelPathTextBox.Text, 2) = "." & PathSep Then
        RelPathTextBox.Text = Mid$(RelPathTextBox.Text, 3, Len(RelPathTextBox.Text) - 2)
    End If
    SetITSetting "Rel Temp Dir", REG_SZ, CStr(RelPathTextBox.Text)
    
    If AbsPathButton.value = True Then
        res = AbsPathTextBox.Text
    Else
        res = "." & PathSep & RelPathTextBox.Text
    End If
    res = AddTrailingSlash(res)
    SetITSetting "Temp Dir", REG_SZ, CStr(res)
    
    ' UTF8
    'SetITSetting "UseUTF8", REG_DWORD, BoolToInt(CheckBoxUTF8.Value)
    
    ' Picture or Shape
    SetITSetting "BitmapVector", REG_DWORD, ComboBoxBitmapVector.ListIndex
    
    Dim VectorOutputTypeList As Variant
    VectorOutputTypeList = GetVectorOutputTypeList()
    Dim VectorOutputType As String
    VectorOutputType = VectorOutputTypeList(ComboBoxVectorOutputType.ListIndex)
    SetITSetting "VectorOutputTypeIdx", REG_DWORD, ComboBoxVectorOutputType.ListIndex
    SetITSetting "VectorOutputType", REG_SZ, CStr(VectorOutputType)
    
    Dim PictureOutputTypeList As Variant
    PictureOutputTypeList = GetPictureOutputTypeDisplayList()
    Dim PictureOutputType As String
    PictureOutputType = PictureOutputTypeList(ComboBoxPictureOutputType.ListIndex)
    SetITSetting "PictureOutputTypeIdx", REG_DWORD, ComboBoxPictureOutputType.ListIndex
    SetITSetting "PictureOutputType", REG_SZ, CStr(PictureOutputType)
    ' Path to External Editor
    res = RemoveQuotes(TextBoxExternalEditor.Text)
    SetITSetting "Editor", REG_SZ, CStr(res)
    ' Use External Editor by default
    SetITSetting "UseExternalEditor", REG_DWORD, BoolToInt(CheckBoxExternalEditor.value)
    ' Docker image used by every generated rendering path.
    res = Trim$(RemoveQuotes(TextBoxTeXExePath.Text))
    If res = vbNullString Then res = DEFAULT_DOCKER_IMAGE
    SetITSetting "DockerImage", REG_SZ, CStr(res)
    ' Path to LaTeXiT-metadata extractor
    res = RemoveQuotes(TextBoxLaTeXiT.Text)
    SetITSetting "LaTeXiT", REG_SZ, CStr(res)
    ' Magic scaling factor to fine-tune Picture displays
    SetITSetting "BitmapScalingX", REG_SZ, TextBoxBitmapScalingX.Text
    SetITSetting "BitmapScalingY", REG_SZ, TextBoxBitmapScalingY.Text
    
    ' Global dpi setting for latex output
    SetITSetting "OutputDpi", REG_DWORD, CLng(val(NormalizeDecimalNumber(TextBoxDpi.Text)))
    
    ' Time Out Interval for Processes
    SetITSetting "TimeOutTime", REG_DWORD, CLng(val(NormalizeDecimalNumber(TextBoxTimeOut.Text)))
    
    ' Font size for text in editor/template windows
    SetITSetting "EditorFontSize", REG_DWORD, CLng(val(NormalizeDecimalNumber(TextBoxFontSize.Text)))
    
    ' LaTeX Engine
    'SetITSetting "LaTeXEngine", REG_SZ, CStr(ComboBoxEngine.Text)
    SetITSetting "LaTeXEngineID", REG_DWORD, ComboBoxEngine.ListIndex

    ' Use Latexmk by default
    SetITSetting "UseLatexmk", REG_DWORD, BoolToInt(CheckBoxLatexmk.value)
    
    ' Add LaTeX source as Alt. text to display by default
    SetITSetting "AddAltText", REG_DWORD, BoolToInt(CheckBoxAltText.value)
    
    ' Keep Temporary files by default
    SetITSetting "KeepTempFiles", REG_DWORD, BoolToInt(CheckBoxKeepTempFiles.value)
    
    ' Height and Width of the Editor Window on Mac (remnant from when it wasn't resizable)
    #If Mac Then
        SetITSetting "LatexFormHeight", REG_DWORD, CLng(val(NormalizeDecimalNumber(TextBoxWindowHeight.Text)))
        SetITSetting "LatexFormWidth", REG_DWORD, CLng(val(NormalizeDecimalNumber(TextBoxWindowWidth.Text)))
    #End If
End Sub

Sub ButtonSetTemp_Click()
    
    SaveSettings
    Unload SetTempForm
End Sub

Private Sub AbsPathButton_Click()
    AbsPathButton.value = True
    SetAbsRelDependencies
End Sub

Private Sub LabelDLtexstudio_Click()
    OpenURL "http://www.texstudio.org/"
End Sub

Private Sub RelPathButton_Click()
    AbsPathButton.value = False
    SetAbsRelDependencies
End Sub

Private Sub SetAbsRelDependencies()
    RelPathButton.value = Not AbsPathButton.value
    AbsPathTextBox.Enabled = AbsPathButton.value
    RelPathTextBox.Enabled = RelPathButton.value
End Sub

Sub ButtonReset_Click()
    AbsPathButton.value = True
    AbsPathTextBox.Text = DEFAULT_TEMP_DIR
    
    'CheckBoxUTF8.Value = True
    
    CheckBoxExternalEditor.value = False
    
    CheckBoxLatexmk.value = False
    CheckBoxAltText.value = True
    CheckBoxKeepTempFiles.value = True
    ComboBoxBitmapVector.ListIndex = 0
    ComboBoxVectorOutputType.ListIndex = 0
    ComboBoxPictureOutputType.ListIndex = 0
    
    Dim UserProfile As String
    #If Mac Then
        UserProfile = vbNullString
    #Else
        UserProfile = Environ$("USERPROFILE")
    #End If
    TextBoxExternalEditor.Text = DEFAULT_EDITOR

    TextBoxTeXExePath.Text = DEFAULT_DOCKER_IMAGE
    
    TextBoxLaTeXiT.Text = Replace(DEFAULT_LATEXIT_METADATA_COMMAND, "%USERPROFILE%", UserProfile)
    
    TextBoxDpi.Text = "1200"
    
    TextBoxBitmapScalingX.Text = "1"
    TextBoxBitmapScalingY.Text = "1"
    
    TextBoxTimeOut.Text = "60"
    
    TextBoxFontSize.Text = "10"
    
    TextBoxWindowHeight.Text = "320"
    TextBoxWindowWidth.Text = "385"
    
    ComboBoxEngine.ListIndex = 0
    
    SetAbsRelDependencies
    
End Sub

Private Sub UserForm_Activate()
    #If Mac Then
        MacEnableCopyPaste Me
        MacEnableAccelerators Me
    #End If
End Sub

Private Sub SetUserFormLayout()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25

    #If Mac Then
        Me.LabelPictureOutputCreationMode.Visible = True
        Me.ComboBoxPictureOutputType.Visible = True
        Me.LabelWindowSize.Visible = True
        Me.LabelWindowHeight.Visible = True
        Me.LabelWindowWidth.Visible = True
        Me.TextBoxWindowHeight.Visible = True
        Me.TextBoxWindowWidth.Visible = True
        Me.LabelFontSize.Caption = "Font size="
        Me.LabelFontSize.Left = 220
        Me.LabelFontSize.Width = 52
        Me.LabelFontSize.Top = Me.LabelWindowSize.Top
        Me.TextBoxFontSize.Top = Me.TextBoxWindowHeight.Top
        Me.TextBoxFontSize.TabIndex = Me.TextBoxWindowWidth.TabIndex
        Me.ButtonExportToXML.Top = Me.LabelWindowSize.Top + 24
        Me.ButtonImportFromXML.Top = Me.ButtonExportToXML.Top
        Me.ButtonCancelTemp.Top = Me.ButtonExportToXML.Top + 34
        Me.ButtonSetTemp.Top = Me.ButtonCancelTemp.Top
        Me.ButtonReset.Top = Me.ButtonCancelTemp.Top
        Me.Height = Me.ButtonCancelTemp.Top + 58
        ResizeUserForm Me
    #End If

    ShowAcceleratorTip Me.ButtonSetTemp
    ShowAcceleratorTip Me.ButtonCancelTemp
    ShowAcceleratorTip Me.ButtonReset
    ShowAcceleratorTip Me.ButtonImportFromXML
    ShowAcceleratorTip Me.ButtonExportToXML
End Sub

Private Sub ReadSavedSettings()
    Dim res As String
    res = GetITSetting("Abs Temp Dir", DEFAULT_TEMP_DIR)
    res = AddTrailingSlash(res)
    AbsPathTextBox.Text = res
    
    RelPathTextBox.Text = GetITSetting("Rel Temp Dir", vbNullString)
    
    AbsPathButton.value = GetITSetting("AbsOrRel", True)
    
    TextBoxDpi.Text = GetITSetting("OutputDpi", "1200")
    
    TextBoxTimeOut.Text = GetITSetting("TimeOutTime", "60")
    
    TextBoxFontSize.Text = GetITSetting("EditorFontSize", "10")
    
    TextBoxBitmapScalingX.Text = GetITSetting("BitmapScalingX", "1")
    TextBoxBitmapScalingY.Text = GetITSetting("BitmapScalingY", "1")
    
    TextBoxExternalEditor.Text = GetITSetting("Editor", DEFAULT_EDITOR)
    CheckBoxExternalEditor.value = GetITSetting("UseExternalEditor", False)
    
    Dim UserProfile As String
    #If Mac Then
        UserProfile = vbNullString
    #Else
        UserProfile = Environ$("USERPROFILE")
    #End If
    TextBoxTeXExePath.Text = GetITSetting("DockerImage", DEFAULT_DOCKER_IMAGE)
    
    TextBoxLaTeXiT.Text = Replace(GetITSetting("LaTeXiT", DEFAULT_LATEXIT_METADATA_COMMAND), "%USERPROFILE%", UserProfile)
    'TextBoxLaTeXiT.Text = GetITSetting("LaTeXiT", DEFAULT_LATEXIT_METADATA_COMMAND)

    ComboBoxBitmapVector.List = GetBitmapVectorList()
    ComboBoxBitmapVector.ListIndex = GetITSetting("BitmapVector", 0)
    ComboBoxVectorOutputType.List = GetVectorOutputTypeDisplayList()
    Dim VectorOutputTypeIndex As Long
    VectorOutputTypeIndex = CLng(val(CStr(GetITSetting("VectorOutputTypeIdx", 0))))
    If VectorOutputTypeIndex < 0 Or _
       VectorOutputTypeIndex >= ComboBoxVectorOutputType.ListCount Then
        VectorOutputTypeIndex = 0
    End If
    ComboBoxVectorOutputType.ListIndex = VectorOutputTypeIndex
    ComboBoxPictureOutputType.List = GetPictureOutputTypeDisplayList()
    Dim PictureOutputTypeIndex As Long
    PictureOutputTypeIndex = CLng(val(CStr(GetITSetting("PictureOutputTypeIdx", 0))))
    If PictureOutputTypeIndex < 0 Or _
       PictureOutputTypeIndex >= ComboBoxPictureOutputType.ListCount Then
        PictureOutputTypeIndex = 0
    End If
    ComboBoxPictureOutputType.ListIndex = PictureOutputTypeIndex
    
    ComboBoxEngine.List = GetLaTexEngineDisplayList()
    ComboBoxEngine.ListIndex = GetITSetting("LaTeXEngineID", 0)
    CheckBoxLatexmk.value = GetITSetting("UseLatexmk", False)
    CheckBoxAltText.value = GetITSetting("AddAltText", True)
    CheckBoxKeepTempFiles.value = GetITSetting("KeepTempFiles", True)
    
    ' Latex editor window size on Mac
    TextBoxWindowHeight.Text = GetITSetting("LatexFormHeight", 320)
    TextBoxWindowWidth.Text = GetITSetting("LatexFormWidth", 385)
End Sub

Private Sub UserForm_Initialize()
    
    SetUserFormLayout

    ReadSavedSettings
    
    SetAbsRelDependencies
End Sub

Private Sub ReadSettingsFromFileIntoRegistry(FilePath As String)
    Dim SZSettingsKeys As Variant
    Dim DWORDSettingsKeys As Variant
    Dim RetiredSettingsKeys As Variant
    Dim SettingsKey As Variant
    Dim XMLText As String
    Dim XMLLines() As String
    Dim HasVectorOutputType As Boolean
    Dim HasVectorOutputTypeIndex As Boolean
    Dim HasPictureOutputType As Boolean
    Dim HasPictureOutputTypeIndex As Boolean
    Dim ImportedVectorOutputType As String
    Dim ImportedVectorOutputTypeIndex As Long
    Dim ImportedPictureOutputType As String
    Dim ImportedPictureOutputTypeIndex As Long
    
    SZSettingsKeys = Array("ColorHex", _
                        "LatexCode", _
                        "Multipage", _
                        "ReadFromFilePath", _
                        "LoadVectorFileScaling", _
                        "LoadVectorFileCalibrationX", "LoadVectorFileCalibrationY", _
                        "Abs Temp Dir", "Rel Temp Dir", "Temp Dir", _
                        "VectorOutputType", "PictureOutputType", _
                        "Editor", "DockerImage", "LaTeXiT", _
                        "BitmapScalingX", "BitmapScalingY", _
                        "TemplateSortedList", "TemplateNameSortedList")
    DWORDSettingsKeys = Array( _
                        "Debug", "AbsOrRel", _
                        "PointSize", _
                        "Transparent", _
                        "OutputDpi", _
                        "LatexCodeCursor", _
                        "EditorFontSize", _
                        "LatexFormWrap", _
                        "LatexFormHeight", "LatexFormWidth", _
                        "BitmapVector", _
                        "VectorOutputTypeIdx", "PictureOutputTypeIdx", _
                        "UseExternalEditor", _
                        "TimeOutTime", _
                        "LaTeXEngineID", _
                        "UseLatexmk", _
                        "AddAltText", _
                        "KeepTempFiles")
    RetiredSettingsKeys = Array( _
                        "GS Command", "IMconv", "TeX2img Command", _
                        "TeXExePath", "TeXExtraPath", "Libgs", _
                        "VectorScalingX", "VectorScalingY", _
                        "LoadVectorFileConvertLines", _
                        "LoadVectorFileOutputTypeIdx", _
                        "LoadVectorFileCleanUp", _
                        "EMFoutput", "UsePDF")
                        
    ' Read XML file
    If FileExists(FilePath) And GetExtension(FilePath) = "xml" Then
        XMLText = ReadAll(FilePath)
    Else
        MsgBox ("The file does not exist or is not an .xml file.")
        Exit Sub
    End If
    XMLLines = Split(XMLText, vbLf)
    Dim i As Integer
    Dim settingName As String
    Dim settingValue As String
    Dim thisXMLLine As String
    Dim inSetting As Boolean
    Dim completeSetting As Boolean
    inSetting = False
    completeSetting = False
    For i = LBound(XMLLines) To UBound(XMLLines)
        thisXMLLine = XMLLines(i)
        If InStr(thisXMLLine, "<Setting Name='") > 0 Then
            thisXMLLine = Right(thisXMLLine, Len(thisXMLLine) - InStr(thisXMLLine, "'"))
            settingName = Left(thisXMLLine, InStr(thisXMLLine, "'") - 1)
            thisXMLLine = Right(thisXMLLine, Len(thisXMLLine) - InStr(thisXMLLine, "'") - 1)
            If InStr(thisXMLLine, "</Setting>") > 0 Then
                inSetting = False
                completeSetting = True
                settingValue = Left(thisXMLLine, InStr(thisXMLLine, "</Setting>") - 1)
            Else
                settingValue = thisXMLLine
                inSetting = True
                completeSetting = False
            End If
        ElseIf inSetting Then
            ' Keep adding to the settingValue until we hit "</Setting>"
            If InStr(thisXMLLine, "</Setting>") > 0 Then
                inSetting = False
                completeSetting = True
                settingValue = settingValue & vbLf & Left(thisXMLLine, InStr(thisXMLLine, "</Setting>") - 1)
            Else
                settingValue = settingValue & vbLf & thisXMLLine
                completeSetting = False
            End If
        End If
        If completeSetting Then
            If IsInArray(RetiredSettingsKeys, settingName) Then
                ' Ignore settings that belonged to retired host renderers.
            ElseIf IsInArray(SZSettingsKeys, settingName) Then
                settingValue = NormalizeImportedSettingValue(settingName, settingValue)
                If settingName = "VectorOutputType" Then
                    HasVectorOutputType = True
                    ImportedVectorOutputType = settingValue
                ElseIf settingName = "PictureOutputType" Then
                    HasPictureOutputType = True
                    ImportedPictureOutputType = settingValue
                End If
                SetITSetting settingName, REG_SZ, settingValue
            ElseIf IsInArray(DWORDSettingsKeys, settingName) Then
                settingValue = NormalizeImportedSettingValue(settingName, settingValue)
                If settingName = "VectorOutputTypeIdx" Then
                    HasVectorOutputTypeIndex = True
                    ImportedVectorOutputTypeIndex = CLng(settingValue)
                ElseIf settingName = "PictureOutputTypeIdx" Then
                    HasPictureOutputTypeIndex = True
                    ImportedPictureOutputTypeIndex = CLng(settingValue)
                End If
                SetITSetting settingName, REG_DWORD, settingValue
            ElseIf InStr(settingName, "Template") Then
                If InStr(settingName, "TemplateCodeSelStart") _
                    Or InStr(settingName, "TemplateLaTeXEngineID") _
                    Or InStr(settingName, "TemplateBitmapVector") Then
                    SetITSetting settingName, REG_DWORD, settingValue
                ElseIf InStr(settingName, "TemplateCode") _
                    Or InStr(settingName, "TemplateTempFolder") _
                    Or InStr(settingName, "TemplateDPI") Then
                    SetITSetting settingName, REG_SZ, settingValue
                Else
                    MsgBox ("Unknown setting: " & settingName & " = " & settingValue)
                End If
            Else
                MsgBox ("Unknown setting: " & settingName & " = " & settingValue)
            End If
            
            completeSetting = False
        End If
    Next i

    SynchronizeImportedOutputType "VectorOutputType", "VectorOutputTypeIdx", _
        GetVectorOutputTypeList(), HasVectorOutputType, ImportedVectorOutputType, _
        HasVectorOutputTypeIndex, ImportedVectorOutputTypeIndex
    SynchronizeImportedOutputType "PictureOutputType", "PictureOutputTypeIdx", _
        GetPictureOutputTypeDisplayList(), HasPictureOutputType, ImportedPictureOutputType, _
        HasPictureOutputTypeIndex, ImportedPictureOutputTypeIndex
End Sub

Private Sub SynchronizeImportedOutputType(ByVal TypeSettingName As String, _
                                          ByVal IndexSettingName As String, _
                                          ByVal AllowedValues As Variant, _
                                          ByVal HasImportedType As Boolean, _
                                          ByVal ImportedType As String, _
                                          ByVal HasImportedIndex As Boolean, _
                                          ByVal ImportedIndex As Long)
    Dim i As Long

    If HasImportedType Then
        For i = LBound(AllowedValues) To UBound(AllowedValues)
            If CStr(AllowedValues(i)) = ImportedType Then
                SetITSetting IndexSettingName, REG_DWORD, i
                Exit For
            End If
        Next i
    ElseIf HasImportedIndex Then
        SetITSetting TypeSettingName, REG_SZ, CStr(AllowedValues(ImportedIndex))
    End If
End Sub

Private Function NormalizeImportedSettingValue(ByVal SettingName As String, _
                                               ByVal SettingValue As String) As String
    Dim AllowedValues As Variant
    Dim ImportedIndex As Long

    NormalizeImportedSettingValue = SettingValue

    Select Case SettingName
        Case "VectorOutputType"
            AllowedValues = GetVectorOutputTypeList()
            If Not IsInArray(AllowedValues, SettingValue) Then
                NormalizeImportedSettingValue = DEFAULT_VECTOR_OUTPUT_TYPE
            End If
        Case "PictureOutputType"
            AllowedValues = GetPictureOutputTypeDisplayList()
            If Not IsInArray(AllowedValues, SettingValue) Then
                NormalizeImportedSettingValue = DEFAULT_PICTURE_OUTPUT_TYPE
            End If
        Case "VectorOutputTypeIdx"
            AllowedValues = GetVectorOutputTypeList()
            ImportedIndex = CLng(val(SettingValue))
            If ImportedIndex < LBound(AllowedValues) Or _
               ImportedIndex > UBound(AllowedValues) Then
                ImportedIndex = 0
            End If
            NormalizeImportedSettingValue = CStr(ImportedIndex)
        Case "PictureOutputTypeIdx"
            AllowedValues = GetPictureOutputTypeDisplayList()
            ImportedIndex = CLng(val(SettingValue))
            If ImportedIndex < LBound(AllowedValues) Or _
               ImportedIndex > UBound(AllowedValues) Then
                ImportedIndex = 0
            End If
            NormalizeImportedSettingValue = CStr(ImportedIndex)
    End Select
End Function

Private Function MakeXMLString(SettingsKey As String, DefaultValue As Variant) As String
    Dim res As String
    res = CStr(GetITSetting(SettingsKey, DefaultValue))
    MakeXMLString = "<Setting Name='" & SettingsKey & "'>" & res & "</Setting>" & vbLf
End Function

Public Sub WriteSettingsToFile(FullFilePath As String)
    Dim xmlContent As String
    Dim SettingsKeys As Variant
    Dim SettingsKey As Variant
    Dim UserProfile As String
    Dim TemplateSortedList() As String
    Dim TemplateID As Long
    Dim RegStr As String
    Dim FolderPath As String
    Dim FilePath As String
    Dim Extension As String
    
    FolderPath = GetFolderFromPath(FullFilePath)
    FilePath = GetFileFromPath(FullFilePath)
    Extension = "." & GetExtension(FullFilePath)
    
    On Error GoTo FileNotWritable
    If IsPathWritable(FolderPath) And FilePath <> vbNullString And Extension = ".xml" Then
    
        ' We first save the settings to registry so that we know what is retrieved from registry is reasonable
        SaveSettings
        
        xmlContent = "<Settings>" & vbLf
        
        ' SetTempForm settings
        xmlContent = xmlContent & MakeXMLString("Abs Temp Dir", DEFAULT_TEMP_DIR)
        xmlContent = xmlContent & MakeXMLString("Rel Temp Dir", vbNullString)
        xmlContent = xmlContent & MakeXMLString("Temp Dir", DEFAULT_TEMP_DIR)
        xmlContent = xmlContent & MakeXMLString("OutputDpi", "1200")
        xmlContent = xmlContent & MakeXMLString("TimeOutTime", "60")
        xmlContent = xmlContent & MakeXMLString("EditorFontSize", "10")
        xmlContent = xmlContent & MakeXMLString("BitmapScalingX", "1")
        xmlContent = xmlContent & MakeXMLString("BitmapScalingY", "1")
        xmlContent = xmlContent & MakeXMLString("Editor", DEFAULT_EDITOR)
        xmlContent = xmlContent & MakeXMLString("DockerImage", DEFAULT_DOCKER_IMAGE)
        xmlContent = xmlContent & MakeXMLString("LatexFormHeight", 320)
        xmlContent = xmlContent & MakeXMLString("LatexFormWidth", 385)
        xmlContent = xmlContent & MakeXMLString("LaTeXiT", DEFAULT_LATEXIT_METADATA_COMMAND)
        xmlContent = xmlContent & MakeXMLString("AbsOrRel", 1)
        xmlContent = xmlContent & MakeXMLString("UseExternalEditor", 0)
        xmlContent = xmlContent & MakeXMLString("BitmapVector", 0)
        xmlContent = xmlContent & MakeXMLString("VectorOutputTypeIdx", 0)
        xmlContent = xmlContent & MakeXMLString("VectorOutputType", DEFAULT_VECTOR_OUTPUT_TYPE)
        xmlContent = xmlContent & MakeXMLString("PictureOutputTypeIdx", 0)
        xmlContent = xmlContent & MakeXMLString("PictureOutputType", DEFAULT_PICTURE_OUTPUT_TYPE)
        xmlContent = xmlContent & MakeXMLString("LaTeXEngineID", 0)
        xmlContent = xmlContent & MakeXMLString("UseLatexmk", 0)
        xmlContent = xmlContent & MakeXMLString("AddAltText", 1)
        xmlContent = xmlContent & MakeXMLString("KeepTempFiles", 1)
        ' LoadVectorGraphicsForm settings
        xmlContent = xmlContent & MakeXMLString("LoadVectorFileScaling", "1")
        xmlContent = xmlContent & MakeXMLString("LoadVectorFileCalibrationX", "1")
        xmlContent = xmlContent & MakeXMLString("LoadVectorFileCalibrationY", "1")
        ' LatexForm settings
        xmlContent = xmlContent & MakeXMLString("Transparent", 1)
        xmlContent = xmlContent & MakeXMLString("Debug", 0)
        xmlContent = xmlContent & MakeXMLString("ColorHex", "000000")
        xmlContent = xmlContent & MakeXMLString("PointSize", "20")
        xmlContent = xmlContent & MakeXMLString("LatexCode", DEFAULT_LATEX_CODE)
        xmlContent = xmlContent & MakeXMLString("LatexCodeCursor", 0)
        xmlContent = xmlContent & MakeXMLString("Multipage", 0)
        xmlContent = xmlContent & MakeXMLString("LatexFormWrap", 1)
        xmlContent = xmlContent & MakeXMLString("ReadFromFilePath", vbNullString)
        xmlContent = xmlContent & MakeXMLString("TemplateSortedList", "0")
        xmlContent = xmlContent & MakeXMLString("TemplateNameSortedList", "New Template")
        ' Template settings
        TemplateSortedList = Split(GetITSetting("TemplateSortedList", "0"), "|", , vbTextCompare)
        For TemplateID = LBound(TemplateSortedList) To UBound(TemplateSortedList) - 1
            xmlContent = xmlContent & MakeXMLString("TemplateCode" & TemplateID, vbNullString)
            xmlContent = xmlContent & MakeXMLString("TemplateCodeSelStart" & TemplateID, 0)
            xmlContent = xmlContent & MakeXMLString("TemplateLaTeXEngineID" & TemplateID, 0)
            xmlContent = xmlContent & MakeXMLString("TemplateBitmapVector" & TemplateID, 0)
            xmlContent = xmlContent & MakeXMLString("TemplateTempFolder" & TemplateID, vbNullString)
            xmlContent = xmlContent & MakeXMLString("TemplateDPI" & TemplateID, vbNullString)
        Next TemplateID
        
        xmlContent = xmlContent & "</Settings>"
    
        WriteToFile FolderPath, FilePath, Extension, xmlContent
        
        MsgBox ("Settings succesfully export to " & FullFilePath)
    Else
        MsgBox "Path is not writable or extension is not .xml."
        Exit Sub
    End If
    On Error GoTo 0

Exit Sub
   
FileNotWritable:
   MsgBox "An error occurred while trying to write the XML file."
End Sub
