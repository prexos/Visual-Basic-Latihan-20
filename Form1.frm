VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Kuliah\Semester 4\Pemrograman\Lat 19\Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TCUSTOMER"
      Top             =   8880
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   3375
      Left            =   2760
      OleObjectBlob   =   "Form1.frx":0014
      TabIndex        =   8
      Top             =   5400
      Width           =   10815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SELESAI"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11760
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11760
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   5880
      TabIndex        =   1
      Top             =   3840
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   5880
      TabIndex        =   0
      Top             =   2760
      Width           =   5655
   End
   Begin VB.Label Label4 
      Caption         =   "PERINTAH SQL:"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "label perintah SQL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   4680
      Width           =   8655
   End
   Begin VB.Label Label2 
      Caption         =   "KRITERIA DATA DITAMPILKAN"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "PERINTAH DASAR URUTAN"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   2760
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AWAL As String

Private Sub Command1_Click()
    On Error GoTo SALAH
    If Text2.Text = AWALKriteria And Text1.Text = AWALurut Then
        Command2.SetFocus
        Exit Sub
    Else
        If Text2.Text = "" Then
            sqltext = "SELECT * FROM TCUSTOMER"
        Else
            sqltext = "SELECT * FROM TCUSTOMER WHERE " + Text2.Text
        End If
        sqltext = sqltext + " order by " + Text1.Text
    End If
    
    Data1.RecordSource = sqltext
    Data1.Refresh
    DBGrid1.Refresh
    
    Label3.Caption = sqltext
    TEXTAWALurut = Text1.Text
    Exit Sub
    
SALAH:
    MsgBox "Data yang diinput tidak valid", vbExclamation, "Error"
    Text1.Text = ""
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Activate()
    AWALurut = ""
    Text1.Text = AWALurut
    AWALKriteria = ""
    Text2.Text = AWALKriteria
    Label3.Caption = ""
    
    Form1.WindowState = 2
End Sub

