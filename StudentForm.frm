VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StudentForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6630
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4870
   OleObjectBlob   =   "StudentForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StudentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Me.Hide
End Sub

Private Sub CommandButton2_Click()
   On Error GoTo ErrorHandler
   Const COL_STUDENT_CODE As Integer = 1
   Const COL_FIRST_NAME As Integer = 2
   Const COL_LAST_NAME As Integer = 3
   Const COL_TEST1 As Integer = 4
   Const COL_TEST2 As Integer = 5
   Const COL_TOTAL As Integer = 6
   
   Dim WordDoc As Document
   Dim wordTable As Table
   Dim newRow As Row
   Dim dblTest1 As Double
   Dim dblTest2 As Double
   Dim dblTotal As Double
   
   If InvalidForm() Then
      Exit Sub
   End If
   
   Set WordDoc = ThisDocument
   Set wordTable = WordDoc.Tables(1)
      
   If WordDoc.Tables.Count = 0 Then
      MsgBox "No table found in this document", vbExclamation, "Error"
      Exit Sub
   End If
   
   With wordTable
     If Trim(.Cell(.Rows.Count, COL_STUDENT_CODE).Range.Text) <> vbCr Then
        Set newRow = .Rows.Add
     Else
        Set newRow = .Rows(.Rows.Count)
     End If
   End With
   
   ' Calculate total
   CalculateTotal
   
   With newRow
     .Cells(COL_STUDENT_CODE).Range.Text = CStr(Student_Code.Value)
     .Cells(COL_FIRST_NAME).Range.Text = proparCase(CStr(First_Name.Value))
     .Cells(COL_LAST_NAME).Range.Text = proparCase(CStr(Last_Name.Value))
     
     If IsNumeric(Note1.Value) Then
        .Cells(COL_TEST1).Range.Text = Format(CDbl(Note1.Value), "0.00")
     Else
        .Cells(COL_TEST1).Range.Text = "D/N"
     End If
     
     If IsNumeric(Note2.Value) Then
        .Cells(COL_TEST2).Range.Text = Format(CDbl(Note2.Value), "0.00")
     Else
        .Cells(COL_TEST2).Range.Text = "D/N"
     End If
     
     .Cells(COL_TOTAL).Range.Text = Total.Value
   End With
   
   ClearForm
   Student_Code.SetFocus
   
   MsgBox "Student record added successfully!", vbInformation, "Success"
   
   Exit Sub
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected Error"
End Sub

Private Sub ClearForm()
    Student_Code.Value = ""
    First_Name.Value = ""
    Last_Name.Value = ""
    Note1.Value = ""
    Note2.Value = ""
    Total.Value = ""
End Sub

Private Function proparCase(ByVal Text As String) As String
    If Len(Text) > 0 Then
       proparCase = UCase(Left(Text, 1)) & LCase(Mid(Text, 2))
    Else
       proparCase = Text
    End If
End Function


Private Sub CalculateTotal()
    If IsNumeric(Note1.Value) And IsNumeric(Note2.Value) Then
        Total.Value = Format(CDbl(Note1.Value) + CDbl(Note2.Value), "0.00")
    Else
        Total.Value = "D/N"
    End If
End Sub

Private Function InvalidForm() As Boolean
    If Student_Code.Value = "" Then
       MsgBox "Student Code cannot be empty", vbExclamation, "Error"
       Student_Code.SetFocus
       InvalidForm = True
       Exit Function
    End If
    
    If First_Name.Value = "" Then
       MsgBox "First Name cannot be empty", vbExclamation, "Error"
       First_Name.SetFocus
       InvalidForm = True
       Exit Function
    End If
    
    If Last_Name.Value = "" Then
       MsgBox "Last Name cannot be empty", vbExclamation, "Error"
       Last_Name.SetFocus
       InvalidForm = True
       Exit Function
    End If
    
    If Note1.Value <> "" And Not IsNumeric(Note1.Value) Then
       MsgBox "Test1 must be numeric", vbExclamation, "Error"
       Test1.SetFocus
       InvalidForm = True
       Exit Function
    End If
    
    If Note2.Value <> "" And Not IsNumeric(Note2.Value) Then
       MsgBox "Test2 must be numeric", vbExclamation, "Error"
       Test2.SetFocus
       InvalidForm = True
       Exit Function
    End If
    
    InvalidForm = False
End Function

Private Sub Note1_Change()
  CalculateTotal
End Sub

Private Sub Note2_Change()
  CalculateTotal
End Sub

Private Sub Total_Change()

End Sub
