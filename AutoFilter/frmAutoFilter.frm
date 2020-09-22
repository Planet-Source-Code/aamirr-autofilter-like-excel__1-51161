VERSION 5.00
Begin VB.Form frmAutoFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom AutoFilter"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framCols 
      Caption         =   "Math"
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7095
      Begin VB.OptionButton optAnd 
         Caption         =   "&Or"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   9
         Top             =   1200
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   8
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   7
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ComboBox cmbField2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   6
         Top             =   1680
         Width           =   3255
      End
      Begin VB.ComboBox cmbCriteria2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmAutoFilter.frx":0000
         Left            =   240
         List            =   "frmAutoFilter.frx":0016
         TabIndex        =   5
         Text            =   "Equals"
         Top             =   1680
         Width           =   3135
      End
      Begin VB.OptionButton optAnd 
         Caption         =   "&And"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox cmbField1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   3
         Top             =   600
         Width           =   3255
      End
      Begin VB.ComboBox cmbCriteria1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmAutoFilter.frx":0089
         Left            =   240
         List            =   "frmAutoFilter.frx":009F
         TabIndex        =   2
         Text            =   "Equals"
         Top             =   600
         Width           =   3135
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show rows where:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmAutoFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strLogicalOpterator1 As String
Dim strLogicalOpterator2 As String
Dim varField1 As Variant
Dim varField2 As Variant
Dim strAndOrOperator As String

Private Sub cmbCriteria1_Click()
strLogicalOpterator1 = cmbCriteria1.Text

End Sub

Private Sub cmbCriteria2_Click()
strLogicalOpterator2 = cmbCriteria2.Text
End Sub

Private Sub cmbField1_Click()
varField1 = cmbField1.Text
End Sub

Private Sub cmbField1_KeyUp(KeyCode As Integer, Shift As Integer)
varField1 = cmbField1.Text
End Sub

Private Sub cmbField2_Click()
varField2 = cmbField2.Text
End Sub

Private Sub cmbField2_KeyUp(KeyCode As Integer, Shift As Integer)
varField2 = cmbField2.Text
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo DataTypeErr
Dim rsAutoFilter As New ADODB.Recordset
Dim rsCounter As Integer, aLoop As Integer
If strLogicalOpterator1 = "Equals" Then strLogicalOpterator1 = "="
If strLogicalOpterator1 = "does Not equals" Then strLogicalOpterator1 = "<>"
If strLogicalOpterator1 = "is greate than" Then strLogicalOpterator1 = ">"
If strLogicalOpterator1 = "is greate than or equals to" Then strLogicalOpterator1 = ">="
If strLogicalOpterator1 = "is less than" Then strLogicalOpterator1 = "<"
If strLogicalOpterator1 = "is less than or equals to" Then strLogicalOpterator1 = "<="

If strLogicalOpterator2 = "Equals" Then strLogicalOpterator2 = "="
If strLogicalOpterator2 = "does Not equals" Then strLogicalOpterator2 = "<>"
If strLogicalOpterator2 = "is greate than" Then strLogicalOpterator2 = ">"
If strLogicalOpterator2 = "is greate than or equals to" Then strLogicalOpterator2 = ">="
If strLogicalOpterator2 = "is less than" Then strLogicalOpterator2 = "<"
If strLogicalOpterator2 = "is less than or equals to" Then strLogicalOpterator2 = "<="

If varField2 = "" Then
    rsAutoFilter.Open "SELECT * From Marksheet WHERE (((Marksheet." & framCols.Caption & ") " & strLogicalOpterator1 & " '" & varField1 & "'))   order by " & framCols.Caption & ";", cn, adOpenStatic, adLockPessimistic
Else
    rsAutoFilter.Open "SELECT * From Marksheet WHERE (((Marksheet." & framCols.Caption & ") " & strLogicalOpterator1 & " '" & varField1 & "')) " & strAndOrOperator & "  (((Marksheet." & framCols.Caption & ") " & strLogicalOpterator2 & " '" & varField2 & "')) order by " & framCols.Caption & ";", cn, adOpenStatic, adLockPessimistic
End If

DataTypeErr:
    If Err.Number = -2147217913 Then
        If varField2 = "" Then
            rsAutoFilter.Open "SELECT * From Marksheet WHERE (((Marksheet." & framCols.Caption & ") " & strLogicalOpterator1 & " " & varField1 & "))   order by " & framCols.Caption & ";", cn, adOpenStatic, adLockPessimistic
        Else
            rsAutoFilter.Open "SELECT * From Marksheet WHERE (((Marksheet." & framCols.Caption & ") " & strLogicalOpterator1 & " " & varField1 & ")) " & strAndOrOperator & "  (((Marksheet." & framCols.Caption & ") " & strLogicalOpterator2 & " " & varField2 & ")) order by " & framCols.Caption & ";", cn, adOpenStatic, adLockPessimistic
        End If
    End If

frmMain.Grid.Clear
For aLoop = 0 To frmMain.Grid.Cols - 1
    frmMain.Grid.Col = aLoop
    frmMain.Grid.Text = rsAutoFilter.Fields(aLoop).Name
Next aLoop

While Not rsAutoFilter.EOF
    rsCounter = rsCounter + 1
    frmMain.Grid.Row = rsCounter
    
        For aLoop = 0 To frmMain.Grid.Cols - 1
            frmMain.Grid.Col = aLoop
            frmMain.Grid.Text = rsAutoFilter.Fields(aLoop)
        Next aLoop
        
    rsAutoFilter.MoveNext
Wend
rsAutoFilter.Close
Set rsAutoFilter = Nothing
varField1 = ""
varField2 = ""
Unload Me

End Sub

Private Sub Form_Activate()
On Error GoTo errDB
Dim rsCols As New ADODB.Recordset
Dim strFrameCaption As String
strFrameCaption = framCols.Caption

rsCols.Open "SELECT distinct( " & strFrameCaption & " ) From Marksheet order by " & strFrameCaption & "", cn, adOpenStatic, adLockPessimistic

cmbField1.Clear
cmbField2.Clear

While Not rsCols.EOF
    cmbField1.AddItem rsCols.Fields("" & strFrameCaption & "")
    cmbField2.AddItem rsCols.Fields("" & strFrameCaption & "")
    rsCols.MoveNext
Wend
rsCols.Close
Set rsCols = Nothing
strLogicalOpterator1 = cmbCriteria1.Text
strLogicalOpterator2 = cmbCriteria2.Text
strAndOrOperator = "Or"
Exit Sub
errDB:
MsgBox Err.Description

End Sub



Private Sub optAnd_Click(Index As Integer)
strAndOrOperator = optAnd(Index).Caption
strAndOrOperator = Mid$(strAndOrOperator, 2, Len(strAndOrOperator) - 1)
End Sub
