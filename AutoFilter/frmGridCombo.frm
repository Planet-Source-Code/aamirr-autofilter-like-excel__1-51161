VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Main Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   3840
      TabIndex        =   1
      Text            =   "Hello"
      Top             =   1080
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6255
      Left            =   3720
      TabIndex        =   0
      Top             =   480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11033
      _Version        =   393216
      Rows            =   100
      Cols            =   6
      FixedCols       =   0
      HighLight       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click here to view AutoFilter >>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim strField As String

Private Sub cmdAdd_Click()
Dim rsAdd As New ADODB.Recordset
rsAdd.Open "select * from temp", cn, adOpenStatic, adLockPessimistic

    rsAdd.AddNew
        rsAdd("ID") = Text1
        rsAdd("Name") = Text2


End Sub

Private Sub Combo1_Click()
On Error GoTo DataTypeErr
Dim rsSearch As New ADODB.Recordset
Dim rsCounter As Integer
Dim strRecord As String
Dim intRecord As Integer

strRecord = Combo1.Text
intRecord = Val(Combo1.Text)
Grid.Row = 0

If strRecord = "( Custom )" Then
    frmAutoFilter.framCols.Caption = strField
    frmAutoFilter.Show
    Combo1.Visible = False
    Exit Sub
ElseIf strRecord = "( All )" Or strRecord = "( Custom )" Then
    Grid.Clear
    rsSearch.Open "SELECT * From Marksheet ", cn, adOpenStatic, adLockPessimistic
Else
    Grid.Clear
    rsSearch.Open "SELECT * From Marksheet WHERE (((Marksheet." & strField & ")='" & strRecord & "'));", cn, adOpenStatic, adLockPessimistic
End If


DataTypeErr:
    If Err.Number = -2147217913 Then
        rsSearch.Open "SELECT * From Marksheet WHERE (((Marksheet." & strField & ")=" & intRecord & "));", cn, adOpenStatic, adLockPessimistic
    End If



For aLoop = 0 To Grid.Cols - 1
    Grid.Col = aLoop
    Grid.Text = rsSearch.Fields(aLoop).Name
Next aLoop

While Not rsSearch.EOF
    rsCounter = rsCounter + 1
    Grid.Row = rsCounter
        For aLoop = 0 To Grid.Cols - 1
            Grid.Col = aLoop
            Grid.Text = rsSearch.Fields(aLoop)
        Next aLoop
        
    rsSearch.MoveNext
Wend
Grid.SetFocus
Combo1.Visible = False

End Sub

Private Sub Combo1_DropDown()
Dim rsDistinct As New ADODB.Recordset
strField = Combo1.Text
rsDistinct.Open "SELECT distinct( " & strField & " ) From Marksheet order by " & strField & "", cn, adOpenStatic, adLockPessimistic
Combo1.Clear
Combo1.AddItem "( All )"
Combo1.AddItem "( Custom )"

While Not rsDistinct.EOF
    Combo1.AddItem rsDistinct.Fields("" & strField & "")
    rsDistinct.MoveNext
Wend
rsDistinct.Close
Combo1.Text = strField
End Sub




Private Sub Form_Load()
Dim aLoop As Integer
Dim rsCounter As Integer


rs.Open "SELECT * FROM Marksheet;", cn, adOpenStatic, adLockPessimistic
Grid.Row = 0
Grid.RowHeight(0) = Combo1.Height

Grid.ColWidth(0) = 500
Grid.ColWidth(1) = 2000

For aLoop = 0 To Grid.Cols - 1
    
    Grid.Col = aLoop
    Grid.Text = rs.Fields(aLoop).Name
Next aLoop

While Not rs.EOF
    rsCounter = rsCounter + 1
    Grid.Row = rsCounter
        For aLoop = 0 To Grid.Cols - 1
            Grid.Col = aLoop
            Grid.Text = rs.Fields(aLoop)
        Next aLoop
        
    rs.MoveNext
Wend
            
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
Set rs = Nothing
cn.Close
End
End Sub

Private Sub Grid_Click()
If Grid.Row = 1 Then
Grid.Row = 0
    With Combo1
       .Top = Grid.CellTop + Grid.Top
        .Left = Grid.CellLeft + Grid.Left
        .Width = Grid.CellWidth
        .Text = Grid.Text
        .Visible = True
        .ZOrder
        .SetFocus
        .SelStart = Len(.Text)
    End With
Else
    Combo1.Visible = False
    
End If

End Sub
