Attribute VB_Name = "UtilititesModule"


Sub SelectFirstEmptyCell(columnNumber)
    Dim ws As Worksheet
    Set ws = TimerSheet
    For Each cell In ws.Columns(columnNumber).Cells
        If IsEmpty(cell) = True Then cell.Select: Exit For
    Next cell
End Sub

Sub InsertDataIntoCells(row, column)
    SelectFirstEmptyCell (1)
    'Insert Project
    Call InsertNextValue(TimerSheet.Cells(row, column).value, 0, False)
    'Insert Client Name
    column = column + 1
    Call InsertNextValue(TimerSheet.Cells(row, column).value, 1, False)
    'Insert Matter
    column = column + 1
    Call InsertNextValue(TimerSheet.Cells(row, column), 1, False)
    'Insert Activity Code
    column = column + 1
    Call InsertNextValue(TimerSheet.Cells(row, column), 2, False)
    'Insert Country/State/City
    column = column + 1
    Call InsertNextValue(TimerSheet.Cells(row, column), 4, False)
   
End Sub

Sub InsertNewProjectIntoCells()
    SelectFirstEmptyCell (14)
    'Insert Project
    Call InsertNextValue(addProjectForm.ProjectTextBox.value, 0, False)
    'Insert Client Name
    Call InsertNextValue(addProjectForm.ClientTextBox.value, 1, False)
    'Insert Matter
    Call InsertNextValue(addProjectForm.MatterTextBox.value, 1, False)
    'Insert Matter
    Call InsertNextValue(addProjectForm.ActivityCodeTextBox.value, 1, False)
    'Insert Location
    Dim location As String
    location = addProjectForm.CityTextBox.value + "/" + addProjectForm.StateTextBox.value + "/" + addProjectForm.CountryTextBox.value
    Call InsertNextValue(location, 1, False)
End Sub

Function InsertNextValue(value As String, columnShift As Integer, accumulate As Boolean)
    Dim currentRow As Integer
    Dim currentColumn As Integer
    currentRow = ActiveCell.row
    currentColumn = ActiveCell.column + columnShift
    Call TimerSheet.Cells(currentRow, currentColumn).Select
    If accumulate = True Then
        value = TimerSheet.Cells(currentRow, currentColumn).value + value
    End If
    
    TimerSheet.Cells(currentRow, currentColumn).value = value
    InsertNextValue = value
End Function

Sub UpdateProjectComboBox(columnNumber As Integer, columnName As String)
    Dim ws As Worksheet
    Dim startRecording As Boolean
    Set ws = TimerSheet
    TimerSheet.ProjectsComboBox.Clear
    For Each cell In ws.Columns(columnNumber).Cells
        If IsEmpty(cell) = True Then Exit For
        If startRecording = True Then
            TimerSheet.ProjectsComboBox.AddItem (cell.value)
        End If
        If cell.value = columnName Then
            startRecording = True
        End If
    Next cell
End Sub

Sub SetActiveProject()
    ActiveSheet.Cells(1, 10).value = TimerSheet.ProjectsComboBox.value
    TimerSheet.currentProject = TimerSheet.ProjectsComboBox.value
End Sub


Sub GetActiveProject()
    TimerSheet.currentProject = ActiveSheet.Cells(1, 10).value
End Sub

Sub StartTime()
    Dim minutesStart As Double
    minutesStart = GetTimeInMinutes
    TimerSheet.Cells(1, 5).value = minutesStart
End Sub


Sub StopTime()
    TimerSheet.Cells(1, 5).value = ""
End Sub


Function GetTimeInMinutes()
    Dim milSecs As Double
    GetTimeInMinutes = Split((DateDiff("n", "01/01/1970", Date) + (Timer / 60)), ".")(0)
End Function

Function GetProjectRow(projectName, column)
    Dim ws As Worksheet
    Set ws = TimerSheet
    For Each cell In ws.Columns(column).Cells
        If IsEmpty(cell) = True Then Exit For
        If cell.value = projectName Then
            GetProjectRow = cell.row: Exit For
        End If
    Next cell
End Function

Sub InsertTaskValues(row As Integer)
    TimerSheet.Cells(row, 1).Select
    Dim taskSummary As String
    taskSummary = " - Date: " + TaskComletedForm.DateCompletedTextBox.value + ", Time: " + TaskComletedForm.TotalTimeTextBox + " minutes" + " (" + CStr(Round(TaskComletedForm.TotalTimeTextBox / 60, 2)) + " hours) " + ", Description: " + TaskComletedForm.NarrativeTextBox.value + vbLf
    Call InsertNextValue(taskSummary, 3, True)
    Dim totalTimeInMinutes As Integer
    totalTimeInMinutes = InsertNextValue(TaskComletedForm.TotalTimeTextBox, 2, True)
    Dim hoursMinutesTotal As String
    hoursMinutesTotal = [totalTimeInMinutes] \ 60 & Format([totalTimeInMinutes] Mod 60, "\:00")
    Call InsertNextValue(hoursMinutesTotal, 1, False)
    Dim hoursFractions As Double
    hoursFractions = totalTimeInMinutes / 60
    Call InsertNextValue(CStr(Round(hoursFractions, 2)), 1, False)
End Sub

Sub ClearAddProjectForm()
    addProjectForm.ProjectTextBox.value = ""
    addProjectForm.ClientTextBox.value = ""
    addProjectForm.MatterTextBox.value = ""
End Sub

Function GetTaskStartTime()
    GetTaskStartTime = TimerSheet.Cells(1, 5).value
End Function

Function GetSumTime(column)
    Dim ws As Worksheet
    Set ws = TimerSheet
    Dim sumTimeInMinutes As Double
    For Each cell In ws.Columns(column).Cells
        If IsEmpty(cell) = True Then Exit For
        If IsNumeric(cell.value) Then
            sumTimeInMinutes = cell.value + sumTimeInMinutes
        End If
    Next cell
    GetSumTime = sumTimeInMinutes
End Function




