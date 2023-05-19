Attribute VB_Name = "SheetUIListenerModule"
Sub AddProject_Click()
    addProjectForm.Show
End Sub

Sub Start_Timer_Click()
    If IsNull(TimerSheet.currentProject) Then
        MsgBox ("Please Select Project")
    End If
End Sub
