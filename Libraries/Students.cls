VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Students"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public sid As String
Public sln As String
Public sfn As String
Public sgender As String

Public Property Get mSID() As String
    mSID = sid
End Property

Public Property Let mSID(ByVal vNewValue As String)
    sid = vNewValue
End Property

Public Property Get mSLN() As String
    mSLN = sln
End Property

Public Property Let mSLN(ByVal vNewValue As String)
    sln = vNewValue
End Property

Public Property Get mSFN() As String
    mSFN = sfn
End Property

Public Property Let mSFN(ByVal vNewValue As String)
    sfn = vNewValue
End Property

Public Property Get mSGender() As String
    mSGender = sgender
End Property

Public Property Let mSGender(ByVal vNewValue As String)
    sgender = vNewValue
End Property


Public Function SaveStudent()
    If rs.State = 1 Then rs.Close
    
    sql = "SELECT * FROM students WHERE IDnumber='" & mSID & "'"
    rs.Open sql, con
    
    If rs.RecordCount >= 1 Then
        MsgBox "Student ID Number already exists!   ", vbExclamation, "Duplicate ID Number"
        Exit Function
    End If
    
    With rs
        .AddNew
            !idnumber = mSID
            !lastname = UCase(mSLN)
            !firstname = UCase(mSFN)
            !gender = UCase(mSGender)
        .Update
    End With
    
    MsgBox "New record was successfully saved!   ", vbInformation, "Save Success"
End Function

Public Function LoadStudents()
    If rs.State = 1 Then rs.Close
    
    sql = "SELECT * FROM students ORDER BY lastname"
    rs.Open sql, con
    
    Set LoadStudents = rs
End Function

Public Function DeleteStudent()
    If rs.State = 1 Then rs.Close
    
    sql = "DELETE * FROM students WHERE idnumber = '" & mSID & "'"
    con.Execute sql
End Function
