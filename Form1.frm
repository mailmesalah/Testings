VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   480
      Left            =   375
      TabIndex        =   0
      Top             =   1725
      Width           =   1350
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Long
    i = 0
    Open "LPT1" For Output As #1
            
    While i <= 255
        Print #1, Chr(27) & "!" & Chr(i) & "------------" & Chr(0) & " " & Chr(27) & "!" & Chr(i + 1) & "|----------|" & Chr(0) & " " & Chr(27) & "!" & Chr(i + 2) & "|----------|" & Chr(0) & " " & Chr(27) & "!" & Chr(i + 3) & "|----------|" & Chr(0)
        Print #1, Chr(27) & "!" & Chr(i) & "------------" & Chr(0) & " " & Chr(27) & "!" & Chr(i + 1) & "|----------|" & Chr(0) & " " & Chr(27) & "!" & Chr(i + 2) & "|----------|" & Chr(0) & " " & Chr(27) & "!" & Chr(i + 3) & "|----------|" & Chr(0)
        Print #1, Chr(27) & "!" & Chr(i) & "|" & Right(Space(3) & i, 3) & " DeXtop|" & Chr(0) & " " & Chr(27) & "!" & Chr(i + 1) & "|" & Right(Space(3) & i, 3) & " DeXtop|" & Chr(0) & " " & Chr(27) & "!" & Chr(i + 2) & "|" & Right(Space(3) & i, 3) & " DeXtop|" & Chr(0) & " " & Chr(27) & "!" & Chr(i + 3) & "|" & Right(Space(3) & i, 3) & " DeXtop|" & Chr(0)
        Print #1, Chr(27) & "!" & Chr(i) & "------------" & Chr(0) & " " & Chr(27) & "!" & Chr(i + 1) & "|----------|" & Chr(0); " " & Chr(27) & "!" & Chr(i + 2) & "|----------|" & Chr(0) & " " & Chr(27) & "!" & Chr(i + 3) & "|----------|" & Chr(0)
        i = i + 4
    Wend
    Close #1
End Sub
