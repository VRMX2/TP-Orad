VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton2_Click()

End Sub

Private Sub CommandButton1_Click()
    Dim WordApp As Word.Application
    Dim WordDoc As Word.Document
    Dim WordTable As Word.Table
    
    Dim ExcelApp As Excel.Application
    Dim ExcelSh As Excel.Worksheet
    Dim ExcelWb As Excel.Workbook
    
    Dim PathToExcel As String
    Dim i As Long, j As Long, LastRow As Long, LastCol As Long
    Dim isCreated As Boolean
    
    On Error GoTo ErrorHandler ' Changed from Resume Next to proper error handling
    
    ' Initialize Excel application
    Set ExcelApp = GetObject(, "Excel.Application")
    If ExcelApp Is Nothing Then
        Set ExcelApp = CreateObject("Excel.Application")
        isCreated = True
    End If
    
    PathToExcel = "C:\vrmx_khadmoney\Tp_Orad\Professional_Example.xlsx"
    
    ' Verify file exists before opening
    If Dir(PathToExcel) = "" Then
        MsgBox "Excel file not found at: " & PathToExcel, vbCritical, "File Error"
        GoTo CleanUp
    End If
    
    Set ExcelWb = ExcelApp.Workbooks.Open(PathToExcel)
    Set ExcelSh = ExcelWb.Sheets("Sheet1")
    ExcelApp.Visible = True
    
    ' Get data range dimensions
    With ExcelSh
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With
    
    ' Initialize Word objects
    Set WordDoc = ThisDocument
    
    ' Check if table exists
    If WordDoc.Tables.Count = 0 Then
        MsgBox "No table found in Word document", vbExclamation, "Error"
        GoTo CleanUp
    End If
    
    Set WordTable = WordDoc.Tables(1)
    
    Application.ScreenUpdating = False
    
    ' Resize and populate Word table
    With WordTable
        ' Add rows if needed (keeping header row)
        Do While .Rows.Count < LastRow
            .Rows.Add
        Loop
        
        ' Add columns if needed
        Do While .Columns.Count < LastCol
            .Columns.Add
        Loop
        
        ' Copy data (skip header if needed)
        For i = 1 To LastRow
            For j = 1 To LastCol
                .Cell(i, j).Range.Text = CStr(ExcelSh.Cells(i, j).Value)
            Next j
        Next i
    End With
    
CleanUp:
    Application.ScreenUpdating = True
    
    ' Proper cleanup sequence
    If Not ExcelSh Is Nothing Then Set ExcelSh = Nothing
    
    If Not ExcelWb Is Nothing Then
        ExcelWb.Close SaveChanges:=False
        Set ExcelWb = Nothing
    End If
    
    If isCreated Then
        If Not ExcelApp Is Nothing Then
            ExcelApp.Quit
            Set ExcelApp = Nothing
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Source: " & Err.Source, vbCritical, "Import Error"
    Resume CleanUp
End Sub
