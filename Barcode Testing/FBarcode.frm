VERSION 5.00
Begin VB.Form FBarcode 
   Caption         =   "Barcode Test"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2265
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CPrint 
      Caption         =   "Print Barcode"
      Height          =   495
      Left            =   2265
      TabIndex        =   0
      Top             =   1710
      Width           =   1890
   End
End
Attribute VB_Name = "FBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PrintBarCode()
 
 ' IMPORTANT: call Printer.Print Space(1) to initialize the hDC of the Printer
Printer.Print Space(1) ' initialize hDC of Printer object
 
 
 ' create barcode object
 Set bc = CreateObject("Bytescout.BarCode.Barcode")
 
 ' set symbology type
 bc.Symbology = 1 ' 1 = Code39
 
 ' set value to encode
 bc.Value = "012345"
 
 ' draw code 39 barcode on a page
 bc.DrawHDC Printer.hDC, 0, 0
 
 
 ' now drawing 2D Aztec barcode
 ' set symbology type
 bc.Symbology = 17 ' 17 = Aztec
 ' set value to encode
 
 bc.Value = "012345"
 
 ' draw Aztec 2D barcode on a page
 bc.DrawHDC Printer.hDC, 0, 300
 
 Printer.Print "This is a test barcode"
 
 ' finally send command to print the page
 Printer.EndDoc
 
 Set bc = Nothing
 
End Sub

Private Sub CPrint_Click()
    PrintBarCode
End Sub
