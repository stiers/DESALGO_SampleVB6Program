VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Save-Update-Delete    -   Active-X DLL"
   ClientHeight    =   5190
   ClientLeft      =   2895
   ClientTop       =   2805
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   6840
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   495
      Left            =   5160
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.ComboBox cboGender 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1680
      List            =   "Form1.frx":000A
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtFN 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox txtLN 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   3360
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5530
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "Gender: "
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "First Name: "
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name: "
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ID Number: "
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objStudents As Students

Private Sub cmdDel_Click()
    Dim confirmDel As String
    
    confirmDel = MsgBox("Are you sure to delete the data?", vbQuestion + vbYesNo, "Confirm")
    
    If confirmDel = vbYes Then
    
        If objStudents.sid <> "" Then
            Call objStudents.DeleteStudent
            Call CustomizeGridHeader
            Call Display
        Else
            MsgBox "Please select a record to delete", vbExclamation, "Delete"
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    objStudents.sid = txtID.Text
    objStudents.SLN = txtLN.Text
    objStudents.sfn = txtFN.Text
    objStudents.sgender = cboGender.Text
    
    Call objStudents.SaveStudent
    Call CustomizeGridHeader
    Call Display
End Sub

Private Sub Form_Load()
    Set objStudents = New Students
    Call CustomizeGridHeader
    Call Display
End Sub

Private Sub CustomizeGridHeader()
    With grid
        .Rows = 1: .Cols = 5
        .Row = 0: .ColWidth(0) = 300
        .Redraw = False
            .Col = 1: .Text = "ID Number": .ColWidth(1) = 1200
            .Col = 2: .Text = "Last Name": .ColWidth(2) = 2000
            .Col = 3: .Text = "First Name": .ColWidth(3) = 2000
            .Col = 4: .Text = "Gender": .ColWidth(4) = 1000
        .Redraw = True
    End With
End Sub

Private Sub Display()
    Dim nRS As ADODB.Recordset
    Dim ctr As Integer
    
    Set nRS = objStudents.LoadStudents
    
    For ctr = 1 To nRS.RecordCount
    
    With grid
        .Rows = .Rows + 1
        .Row = ctr
            .Col = 1: .Text = nRS.Fields(1).Value
            .Col = 2: .Text = nRS.Fields(2).Value
            .Col = 3: .Text = nRS.Fields(3).Value
            .Col = 4: .Text = nRS.Fields(4).Value
        .Redraw = True
        nRS.MoveNext
    End With
    Next ctr
End Sub

Private Sub grid_Click()
    With grid
        .Col = 1
        objStudents.sid = .Text
    End With
End Sub
