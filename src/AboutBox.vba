Option Explicit
Public Sub CloseAboutButton_Click()
    Unload AboutBox
End Sub


Private Sub LabelURL_Click()
    OpenURL "https://www.jonathanleroux.org/software/iguanatex/"
End Sub

Private Sub LabelGithub_Click()
    OpenURL "https://github.com/Jonathan-LeRoux/IguanaTex"
End Sub

Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    ShowAcceleratorTip Me.CloseAboutButton
    #If Mac Then
        ResizeUserForm Me
    #End If
End Sub

Private Sub UserForm_Activate()
    #If Mac Then
        MacEnableAccelerators Me
    #End If
End Sub

