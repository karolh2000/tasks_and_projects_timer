VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TaskComletedForm 
   Caption         =   "Task Complete"
   ClientHeight    =   5320
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   6986
   OleObjectBlob   =   "TaskComletedForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TaskComletedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub EndTaskButton_Click()
    TaskComletedForm.Hide
    Dim projectName As String
    projectName = TimerSheet.ProjectsComboBox.value
    Dim projectRow As Integer
    projectRow = UtilititesModule.GetProjectRow(projectName, 1)
    UtilititesModule.InsertTaskValues (projectRow)
End Sub

Private Sub UserForm_Activate()
    TaskComletedForm.ProjectTextBox.value = TimerSheet.ProjectsComboBox.value
    TaskComletedForm.DateCompletedTextBox.value = Now
    Dim currentMinutes As Double
    currentMinutes = UtilititesModule.GetTimeInMinutes
    Dim timeSpentMins As Double
    timeSpentMins = Round(currentMinutes - UtilititesModule.GetTaskStartTime, 0)
    TaskComletedForm.TotalTimeTextBox.value = timeSpentMins
    TaskComletedForm.NarrativeTextBox.value = ""
End Sub
