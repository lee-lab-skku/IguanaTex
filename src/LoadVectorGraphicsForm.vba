Option Explicit

Sub CommandButtonSave_Click()
    SetITSetting "LoadVectorFileScaling", REG_SZ, textboxScalor.Text
    SetITSetting "LoadVectorFileCalibrationX", REG_SZ, TextBoxCalibrationX.Text
    SetITSetting "LoadVectorFileCalibrationY", REG_SZ, TextBoxCalibrationY.Text
End Sub

Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    Me.Height = 194
    Me.Width = 355
    Me.LabelInsertPath.Caption = "Insert path of .pdf/.dvi/.xdv/.ps/.eps/.svg file:"
    #If Mac Then
        ResizeUserForm Me
    #End If
    textboxScalor.Text = GetITSetting("LoadVectorFileScaling", "1")
    TextBoxCalibrationX.Text = GetITSetting("LoadVectorFileCalibrationX", "1")
    TextBoxCalibrationY.Text = GetITSetting("LoadVectorFileCalibrationY", "1")
    ComboBoxVectorOutputType.List = Array("SVG via Docker")
    ComboBoxVectorOutputType.ListIndex = 0
    ComboBoxVectorOutputType.Enabled = False
    CheckBoxCleanUp.value = False
    CheckBoxCleanUp.Enabled = False
    CheckBoxCleanUp.Visible = False
    CheckBoxConvertLines.value = False
    CheckBoxConvertLines.Enabled = False
    CheckBoxConvertLines.Visible = False
    ShowAcceleratorTip Me.ButtonLoadFile
    ShowAcceleratorTip Me.ButtonCancel
    ShowAcceleratorTip Me.CommandButtonSave
    
End Sub

Private Sub UserForm_Activate()
    #If Mac Then
        MacEnableAccelerators Me
    #End If
End Sub

Sub ButtonCancel_Click()
    Unload LoadVectorGraphicsForm
End Sub

Private Function isInsertableVectorFile(file As String) As Boolean
    Dim Ext As String
    Ext = GetExtension(file)
    isInsertableVectorFile = Ext = "pdf" Or Ext = "dvi" Or Ext = "xdv" Or _
        Ext = "ps" Or Ext = "eps" Or Ext = "svg"
End Function

Sub ButtonPath_Click()
    #If Mac Then
        TextBoxFile.Text = MacChooseFileOfType("pdf,dvi,xdv,ps,eps,svg")
    #Else
        TextBoxFile.Text = BrowseFilePath(TextBoxFile.Text, "Vector graphics files", _
            "*.pdf;*.dvi;*.xdv;*.ps;*.eps;*.svg", "&Select file")
    #End If
    TextBoxFile.SetFocus
End Sub

Private Sub TextBoxFile_Change()
    Dim path As String, Ext As String
    path = TextBoxFile.Text
    Ext = GetExtension(path)
    ButtonLoadFile.Enabled = FileExists(path) And isInsertableVectorFile(path)
    If Ext = "svg" Then
        ComboBoxVectorOutputType.ListIndex = 0
        ComboBoxVectorOutputType.Enabled = False
    End If
End Sub

Sub ButtonLoadFile_Click()
    DoInsertVectorGraphicsFile
    Unload LoadVectorGraphicsForm
End Sub


Private Sub DoInsertVectorGraphicsFile()
    Dim NewShape As Shape
    Dim TimeOutTimeString As String
    Dim TimeOutTime As Long
    TimeOutTimeString = GetITSetting("TimeOutTime", "20")
    TimeOutTime = val(NormalizeDecimalNumber(TimeOutTimeString)) * 1000

    Dim debugMode As Boolean
    debugMode = False
    #If Mac Then
        Dim fs As New MacFileSystemObject
    #Else
        Dim fs As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
    #End If

    Dim TempPath As String
    TempPath = CleanPath(GetTempPath())
    If Not IsPathWritable(TempPath) Then Exit Sub

    Dim posX As Single, posY As Single, ScalingX As Single, ScalingY As Single
    Dim Sel As Selection
    Set Sel = Application.ActiveWindow.Selection
    If Sel.Type = ppSelectionShapes Then
        posX = Sel.ShapeRange(1).Left
        posY = Sel.ShapeRange(1).Top
    Else
        posX = 200
        posY = 200
    End If
    ScalingX = val(NormalizeDecimalNumber(textboxScalor.value)) * _
        val(NormalizeDecimalNumber(TextBoxCalibrationX.value))
    ScalingY = val(NormalizeDecimalNumber(textboxScalor.value)) * _
        val(NormalizeDecimalNumber(TextBoxCalibrationY.value))

    Dim path As String
    Dim Ext As String
    path = TextBoxFile.Text
    Ext = GetExtension(path)

    If Ext = "svg" Then
        Set NewShape = AddDisplayShape(path, posX, posY)
    Else
        Dim FinalFilename As String
        Dim RunCommand As String
        Dim DockerFailureStage As String
        Dim PreserveDockerWorkspace As Boolean
        Dim RetVal As Long
        PreserveDockerWorkspace = debugMode Or _
            CBool(GetITSetting("KeepTempFiles", True))

        RetVal = ExecuteDockerVectorConversionJob(path, TempPath, debugMode, _
            TimeOutTime, PreserveDockerWorkspace, FinalFilename, RunCommand, _
            DockerFailureStage)
        If RetVal <> 0 Or Not fs.FileExists(TempPath & FinalFilename) Then
            ShowError DockerRenderErrorMessage(DockerFailureStage, _
                TimeOutTimeString), RunCommand
            Exit Sub
        End If

        Set NewShape = AddDisplayShape(TempPath & FinalFilename, posX, posY)
        If Not PreserveDockerWorkspace And _
           fs.FileExists(TempPath & FinalFilename) Then
            fs.DeleteFile TempPath & FinalFilename
        End If
    End If

    Set NewShape = convertSVG(NewShape, ScalingX, ScalingY, posX, posY)
    NewShape.Select
End Sub


