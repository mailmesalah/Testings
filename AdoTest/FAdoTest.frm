VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form FAdoTest 
   Caption         =   "AdoTest"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CTest 
      Caption         =   "Test"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   5880
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adb 
      Height          =   495
      Left            =   720
      Top             =   2520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=adotest"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=adotest"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "FAdoTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim cn As adodb.Connection

Private Sub Form_Load()

Dim rs As adodb.Recordset

    'db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " _
    '           & "Data Source=C:\Program Files" _
    '                         & "\Microsoft Visual Studio" _
    '                         & "\VB98\Biblio.mdb"
    'Set db = Adb
    'Set db.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=adotest"
    'db.RecordSource = "Select * from Account Master"
    cn.ConnectionString = "DSN=adotest;Uid=;Pwd=12345abcde;"
    cn.Open
    rs.Open "Select * from Account Master"
    If rs.RecordCount > 0 Then
        MsgBox "hai"
    End If
    
End Sub
