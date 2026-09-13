Attribute VB_Name = "FeasibilityReport"
Option Explicit

' Generate the PDF feasibility report for the workbook that is currently open.
'
' Saves the workbook first (so the generator reads what is on screen), shells
' out to the Python generator, waits for it, and opens the result.
'
' Install: see excel_addin/README.md

Private Const SCRIPT_NAME As String = "build_report.py"
Private Const VENV_PYTHON As String = "\.venv\Scripts\pythonw.exe"


Public Sub GenerateFeasibilityReport()
    Dim wb As Workbook
    Dim sheetName As String
    Dim projectDir As String
    Dim scriptPath As String
    Dim pythonPath As String
    Dim command As String
    Dim shell As Object
    Dim exitCode As Long

    Set wb = ActiveWorkbook

    If wb Is Nothing Then
        MsgBox "No workbook is open.", vbExclamation, "Feasibility report"
        Exit Sub
    End If

    If Len(wb.Path) = 0 Then
        MsgBox "Save the workbook before generating a report.", _
               vbExclamation, "Feasibility report"
        Exit Sub
    End If

    sheetName = ActiveSheet.Name
    If sheetName <> "Residential" And sheetName <> "Townhouse" Then
        MsgBox "Select the Residential or Townhouse tab first." & vbCrLf & vbCrLf & _
               "The report covers one deal type at a time.", _
               vbExclamation, "Feasibility report"
        Exit Sub
    End If

    ' The workbook lives in <project>\output\, so the project is one level up.
    projectDir = ParentFolder(wb.Path)
    scriptPath = projectDir & "\" & SCRIPT_NAME

    If Len(Dir(scriptPath)) = 0 Then
        MsgBox "Could not find the generator:" & vbCrLf & scriptPath & vbCrLf & vbCrLf & _
               "This workbook must stay inside the project folder.", _
               vbCritical, "Feasibility report"
        Exit Sub
    End If

    pythonPath = ResolvePython(projectDir)
    If Len(pythonPath) = 0 Then
        MsgBox "Could not find a Python interpreter." & vbCrLf & vbCrLf & _
               "Expected a virtual environment at:" & vbCrLf & _
               projectDir & VENV_PYTHON, _
               vbCritical, "Feasibility report"
        Exit Sub
    End If

    ' Save so the generator reads the numbers currently on screen, then close:
    ' the generator opens the file itself and Excel will not hand over a file it
    ' still has locked for editing.
    Application.DisplayAlerts = False
    wb.Save
    Application.DisplayAlerts = True

    command = """" & pythonPath & """ """ & scriptPath & """" & _
              " --sheet " & sheetName & _
              " --workbook """ & wb.FullName & """" & _
              " --pdf --open"

    Application.StatusBar = "Generating " & sheetName & " report..."
    Application.Cursor = xlWait

    On Error GoTo Failed
    Set shell = CreateObject("WScript.Shell")
    exitCode = shell.Run(command, 0, True)   ' hidden window, wait for it
    On Error GoTo 0

    Application.Cursor = xlDefault
    Application.StatusBar = False

    If exitCode <> 0 Then
        MsgBox "The generator reported an error (exit code " & exitCode & ")." & vbCrLf & vbCrLf & _
               "Run it from a terminal to see why:" & vbCrLf & command, _
               vbCritical, "Feasibility report"
    End If
    Exit Sub

Failed:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    MsgBox "Could not start the generator." & vbCrLf & vbCrLf & Err.Description, _
           vbCritical, "Feasibility report"
End Sub


Private Function ParentFolder(ByVal path As String) As String
    Dim i As Long
    i = InStrRev(path, "\")
    If i > 0 Then
        ParentFolder = Left$(path, i - 1)
    Else
        ParentFolder = path
    End If
End Function


Private Function ResolvePython(ByVal projectDir As String) As String
    ' Prefer the project's own virtual environment. pythonw.exe rather than
    ' python.exe so no console window flashes up.
    Dim candidate As String

    candidate = projectDir & VENV_PYTHON
    If Len(Dir(candidate)) > 0 Then
        ResolvePython = candidate
        Exit Function
    End If

    candidate = Replace(candidate, "pythonw.exe", "python.exe")
    If Len(Dir(candidate)) > 0 Then
        ResolvePython = candidate
        Exit Function
    End If

    ResolvePython = ""
End Function
