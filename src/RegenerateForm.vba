Option Explicit

Private Sub UserForm_Initialize()
    Me.Top = Application.Top + 110
    Me.Left = Application.Left + 25
    #If Mac Then
        ResizeUserForm Me
    #End If
End Sub

Private Sub UserForm_Activate()
    #If Mac Then
        MacEnableAccelerators Me
    #End If
End Sub

Sub CommandButtonCancel_Click()
    RegenerateContinue = False
    Unload RegenerateForm
End Sub
