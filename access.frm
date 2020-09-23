VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Query"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexecute 
      Caption         =   "Execute Query"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   960
      TabIndex        =   8
      Top             =   3990
      Width           =   3375
   End
   Begin VB.TextBox txtSQL 
      Height          =   495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "access.frx":0000
      Top             =   3480
      Width           =   5175
   End
   Begin VB.ListBox QryList 
      Height          =   1815
      Left            =   2520
      TabIndex        =   6
      Top             =   1200
      Width           =   2655
   End
   Begin VB.ListBox TblList 
      Height          =   1815
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "Open Database"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Queries"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Selected Query Definition"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tables"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   660
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdexecute_Click()
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim n As Integer
Dim m As Integer
Dim i As Integer
Dim j As Integer
Set rs = New ADODB.Recordset
rs.ActiveConnection = cn
rs.Open txtSQL.Text, cn, adOpenStatic, adLockOptimistic
n = rs.RecordCount
i = rs.Fields.Count
Set xlApp = New Excel.Application
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add
xlApp.Application.Visible = True
xlSheet.Cells(3, 1) = "Result"
While Not (rs.EOF)
For m = 0 To n - 1
    For j = 0 To i - 1
        xlSheet.Cells(m + 2, j + 1) = rs(j)
   Next
   rs.MoveNext
            If rs.EOF Then
                Exit For
            End If
Next
Wend
End Sub

Private Sub CmdOpen_Click()
Dim cat As New ADOX.Catalog
Dim cmd As New ADODB.Command
On Error GoTo NODATABASE
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Databases|*.mdb"
    CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
connectstring = CommonDialog1.FileName
End If
With cn
      .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & connectstring & ";Persist Security Info=False"
      .ConnectionTimeout = 10
      .Open
End With
Label1.Caption = connectstring
Set rs = cn.OpenSchema(adSchemaTables)
Do While Not (rs.EOF)
  If (rs!TABLE_TYPE) = "VIEW" Then
   QryList.AddItem rs!TABLE_NAME
  End If
   If rs!TABLE_TYPE = "TABLE" Then
      If Asc(Left(rs!TABLE_NAME, 1)) >= Asc("A") And Asc(Left(rs!TABLE_NAME, 1)) <= Asc("z") Then
      TblList.AddItem rs!TABLE_NAME
      End If
   End If
 rs.MoveNext
Loop
rs.Close
NODATABASE:
End Sub


Private Sub Command1_Click()

End Sub

Private Sub QryList_Click()
Dim cat As New ADOX.Catalog
Dim cmd As New ADODB.Command
Set cat.ActiveConnection = cn
Set cmd = cat.Views(QryList.List(QryList.ListIndex)).Command
txtSQL.Text = cmd.CommandText
End Sub

