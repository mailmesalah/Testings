VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   4425
   ClientTop       =   2985
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   5970
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.OLE OLE1 
      Height          =   975
      Left            =   1080
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit

  Dim oBook As Object
  Dim oSheet As Object

  Private Sub Command1_Click()
    On Error GoTo Err_Handler

  ' Create a new Excel worksheet...
    OLE1.CreateEmbed vbNullString, "Excel.Sheet"

  ' Now, pre-fill it with some data you
  ' can use. The OLE.Object property returns a
  ' workbook object, and you can use Sheets(1)
  ' to get the first sheet.
    Dim arrData(1 To 5, 1 To 5) As Variant
    Dim i As Long, j As Long

    Set oBook = OLE1.object
    Set oSheet = oBook.Sheets(1)

  ' It is much more efficient to use an array to
  ' pass data to Excel than to push data over
  ' cell-by-cell, so you can use an array.

  ' Add some column headers to the array...
    arrData(1, 2) = "April"
    arrData(1, 3) = "May"
    arrData(1, 4) = "June"
    arrData(1, 5) = "July"

  ' Add some row headers...
    arrData(2, 1) = "John"
    arrData(3, 1) = "Sally"
    arrData(4, 1) = "Charles"
    arrData(5, 1) = "Toni"

  ' Now add some data...
    For i = 2 To 5
       For j = 2 To 5
          arrData(i, j) = 350 + ((i + j) Mod 3)
       Next j
    Next i

  ' Assign the data to Excel...
    oSheet.Range("A3:E7").Value = arrData

    oSheet.Cells(1, 1).Value = "Test Data"
    oSheet.Range("B9:E9").FormulaR1C1 = "=SUM(R[-5]C:R[-2]C)"

  ' Do some auto formatting...
    oSheet.Range("A1:E9").Select
    oBook.Application.Selection.AutoFormat

    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = True
    Exit Sub

Err_Handler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
 End Sub

 Private Sub Command2_Click()
    On Error GoTo Err_Handler

  ' Create an embedded object using the data
  ' stored in Test.xls.<?xm-insertion_mark_start author="v-thomr" time="20070327T040420-0600"?> If this code is run in Microsoft Office
  ' Excel 2007, <?xm-insertion_mark_end?><?xm-deletion_mark author="v-thomr" time="20070327T040345-0600" data=".."?><?xm-insertion_mark_start author="v-thomr" time="20070327T040422-0600"?>change the file name to Test.xlsx.<?xm-insertion_mark_end?>
    OLE1.CreateEmbed "C:\Test.xls"

    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = True
    Exit Sub

Err_Handler:
    MsgBox "The file 'C:\Test.xls' does not exist" & _
           " or cannot be opened.", vbCritical
 End Sub

 Private Sub Command3_Click()
    On Error Resume Next

  ' Delete the existing test file (if any)...
    Kill "C:\Test.xls"

  ' Save the file as a native XLS file...
    oBook.SaveAs "C:\Test.xls"
    Set oBook = Nothing
    Set oSheet = Nothing

  ' Close the OLE object and remove it...
    OLE1.Close
    OLE1.Delete

       Command1.Enabled = True
       Command2.Enabled = True
      Command3.Enabled = False
    End Sub

    Private Sub Form_Load()
       Command1.Caption = "Create"
       Command2.Caption = "Open"
       Command3.Caption = "Save"
       Command3.Enabled = False
    End Sub


