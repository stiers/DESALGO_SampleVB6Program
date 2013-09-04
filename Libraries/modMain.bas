Attribute VB_Name = "modMain"
Option Explicit

Public rs As ADODB.Recordset
Public con As ADODB.Connection
Public sql As String

Public Sub Main()
    Set rs = New ADODB.Recordset
    Set con = New ADODB.Connection
    
    With rs
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
    End With
    
    Dim dbPath As String
    
    dbPath = App.Path & "\data.mdb"
    
    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";Persist Security Info=false;"
End Sub
