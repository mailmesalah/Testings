VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8190
   LinkTopic       =   "Form2"
   ScaleHeight     =   5985
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtname 
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton cmdname 
      Caption         =   "Name"
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdid 
      Caption         =   "ID"
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtid 
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton cmdselectall 
      Caption         =   "Command1"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin Crystal.CrystalReport cr 
      Left            =   360
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label2 
      Caption         =   "by name"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "by ID"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdid_Click()
With cr
.ReportFileName = App.Path & "\report1.rpt"
.WindowState = crptMaximized
'if our field is a number we need to remove the pair of single qoute in our query ''
.ReplaceSelectionFormula " { table1.id} = " & txtid.Text & ""
.Destination = crptToWindow
.Action = 1


End With
End Sub

Private Sub cmdname_Click()
With cr
.ReportFileName = App.Path & "\report1.rpt"
.WindowState = crptMaximized

.ReplaceSelectionFormula " { table1.name} = '" & txtname.Text & "'"
.Destination = crptToWindow
.Action = 1


End With

End Sub

Private Sub cmdselectall_Click()
With cr


.ReportFileName = App.Path & "\report1.rpt"
.WindowState = crptMaximized
.Destination = crptToWindow
.Action = 1
'.ReportSource = objReport
'.ViewReport


End With
End Sub
