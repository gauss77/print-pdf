VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Print PDF from MSFlexGrid"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid fgMain 
      Height          =   6615
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11668
      _Version        =   393216
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "Print"
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   6960
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnPrint_Click()
    
    Dim prn As Printer
    Dim found As Boolean
    found = False
    For Each prn In Printers
        If InStr(prn.DeviceName, PDF) > 0 Then
            Set Printer = prn:
            found = True
            Exit For
        End If
    Next
    If found = False Then
        MsgBox "No PDF printer can be found."
        Exit Sub
    End If
    
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait
    Set Printer.font = Me.font
    Printer.font.Size = 14
        
    Printer.DrawWidth = 2
    Printer.ForeColor = RGB(0, 0, 0) ' Black color

    ' Loop through the rows and columns to print the table with borders
    Dim X As Integer, y As Integer
    Dim maringX As Integer, marginY As Integer
    Dim gap As Integer
    Dim imgCellX As Integer, imageCellY As Integer
    Dim pic As StdPicture
     
    marginX = Printer.ScaleWidth / 12
    gap = 100
        
    X = marginX + gap
    marginY = Printer.ScaleHeight / 20
    y = marginY
    
    imgCellX = marginX + Printer.ScaleWidth / 2
    ' Print main table
    For Row = 1 To fgMain.Rows
        If y >= Printer.ScaleHeight - marginY * 2 Or Row = fgMain.Rows Then
            Printer.Line (marginX, y)-(Printer.ScaleWidth - marginX, y)
    
            Printer.Line (marginX, marginY)-(marginX, y)
            Printer.Line (imgCellX, marginY)-(imgCellX, y)
            Printer.Line (Printer.ScaleWidth - marginX, marginY)-(Printer.ScaleWidth - marginX, y)
            y = marginY
            If Row = fgMain.Rows Then
                Exit For
            End If
            
            Printer.NewPage
        End If
        
        If y = marginY Then
            ' Print header
            Printer.font.Bold = True
            Printer.Line (marginX, y)-(Printer.ScaleWidth - marginX, y)
            Printer.CurrentX = marginX + gap
            Printer.CurrentY = y + gap
            Printer.Print "Item Description"
            Printer.CurrentX = imgCellX + gap
            Printer.CurrentY = y + gap
            Printer.Print "Image"
            Printer.font.Bold = False
            y = Printer.CurrentY + 100
        End If
        
        Printer.Line (marginX, y)-(Printer.ScaleWidth - marginX, y)
        
        Printer.CurrentX = marginX + gap
        Printer.CurrentY = y + gap
        Printer.Print fgMain.TextMatrix(Row, 1)
        
        Printer.CurrentX = marginX + gap
        Printer.CurrentY = Printer.CurrentY + gap
        Printer.Print fgMain.TextMatrix(Row, 2)
        
        Printer.CurrentX = marginX + gap
        Printer.CurrentY = Printer.CurrentY + gap
        Printer.Print fgMain.TextMatrix(Row, 3)

        imageCellY = y + gap
        y = Printer.CurrentY + 100
        
        Printer.CurrentX = imgCellX + gap
        Printer.CurrentY = imageCellY
        Set pic = LoadPicture(fgMain.TextMatrix(Row, 4))
        Printer.PaintPicture pic, Printer.CurrentX, Printer.CurrentY, Printer.ScaleWidth / 7, y - imageCellY - gap
       
    Next Row

    ' End the print job
    Printer.EndDoc
    
End Sub


Private Sub Form_Activate()
    Dim pic As StdPicture

    With fgMain
        .Rows = 21
        .Cols = 5
        .ColWidth(0) = 400
        .ColWidth(1) = 2000
        .ColWidth(2) = 800
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
    
        ' Set column headers
        .TextMatrix(0, 0) = "No"
        .TextMatrix(0, 1) = "title"
        .TextMatrix(0, 2) = "stock"
        .TextMatrix(0, 3) = "group"
        .TextMatrix(0, 4) = "image"
        
        For i = 1 To 20
            .RowHeight(i) = 600
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = "Hello World!" + Str(i)
            .TextMatrix(i, 2) = Str(Int((100 * Rnd) + 5)) & " pcs"
            .TextMatrix(i, 3) = "Buicut " + Str(i)
            .TextMatrix(i, 4) = "C:\test\" + Trim(Int(Rnd * 5 + 1)) + ".jpg"
            Set .CellPicture = pic
        Next
    End With
End Sub
