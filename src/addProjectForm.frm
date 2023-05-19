VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddProjectForm 
   Caption         =   "Add Project"
   ClientHeight    =   4060
   ClientLeft      =   49
   ClientTop       =   203
   ClientWidth     =   8120
   OleObjectBlob   =   "addProjectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "addProjectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub AddProjectButton_Click()
    Call UtilititesModule.InsertNewProjectIntoCells
    Call addProjectForm.Hide
    Call UtilititesModule.UpdateProjectComboBox(14, "PROJECT")
End Sub

Private Sub CancelButton_Click()
    addProjectForm.Hide
End Sub

