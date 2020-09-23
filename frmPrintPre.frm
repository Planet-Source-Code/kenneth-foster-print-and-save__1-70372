VERSION 5.00
Begin VB.Form frmPrintPre 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                           Preview And Print Form"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   5805
   Begin VB.CommandButton Command1 
      Caption         =   "Save as Bitmap"
      Height          =   540
      Left            =   45
      TabIndex        =   4
      Top             =   7755
      Width           =   1590
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   540
      Left            =   2040
      TabIndex        =   3
      Top             =   7755
      Width           =   1875
   End
   Begin VB.CommandButton cmdPrintAll 
      Caption         =   "Print Pictures"
      Height          =   540
      Left            =   4290
      TabIndex        =   2
      Top             =   7755
      Width           =   1425
   End
   Begin VB.PictureBox PicSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Left            =   7470
      ScaleHeight     =   705
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   9075
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7425
      Left            =   15
      ScaleHeight     =   495
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   0
      Top             =   180
      Width           =   5760
   End
End
Attribute VB_Name = "frmPrintPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'Printer code is from Microsoft site, modified a little for this program
'

Private Sub Form_Load()
   Dim dRatio As Double
   Dim x As Integer
   Me.Left = frmMain.Left + 650
   Me.Top = frmMain.Top
   dRatio = ScalePicPreviewToPrinterInches(Picture1)
   PrintRoutineAll Picture1, dRatio
End Sub

Private Sub Command1_Click()
   With frmEdit.cde
        .DialogTitle = "Save Picture"
        .Filter = "Bitmap (*.Bmp )|*.bmp"
        Picture1.Picture = Picture1.Image
        .ShowSave
        SavePicture Picture1.Picture, .FileName
        MsgBox "Picture saved...." & .FileName
    End With
End Sub

Private Sub cmdCancel_Click()
Dim x As Integer
    For x = 0 To 7
        frmMain.Shape1(x).Visible = True
     Next x
    Unload Me
End Sub

Private Sub cmdPrintAll_Click()   'print
         Printer.ScaleMode = vbInches
         PrintRoutineAll Printer
         Printer.EndDoc
End Sub

Private Function ScalePicPreviewToPrinterInches(picPreview As PictureBox) As Double

         Dim Ratio As Double ' Ratio between Printer and Picture
         Dim LRGap As Double, TBGap As Double
         Dim HeightRatio As Double, WidthRatio As Double
         Dim PgWidth As Double, PgHeight As Double
         Dim smtemp As Long

         ' Get the physical page size in Inches:
         PgWidth = Printer.Width / 1440
         PgHeight = Printer.Height / 1440

         ' Find the size of the non-printable area on the printer to
         ' use to offset coordinates. These formulas assume the
         ' printable area is centered on the page:
         smtemp = Printer.ScaleMode
         Printer.ScaleMode = vbInches
         LRGap = (PgWidth - Printer.ScaleWidth) / 2
         TBGap = (PgHeight - Printer.ScaleHeight) / 2
         Printer.ScaleMode = smtemp

         ' Scale PictureBox to Printer's printable area in Inches:
         picPreview.ScaleMode = vbInches

         ' Compare the height and with ratios to determine the
         ' Ratio to use and how to size the picture box:
         HeightRatio = picPreview.ScaleHeight / PgHeight
         WidthRatio = picPreview.ScaleWidth / PgWidth

         If HeightRatio < WidthRatio Then
            Ratio = HeightRatio
            smtemp = picPreview.Container.ScaleMode
            picPreview.Container.ScaleMode = vbInches
            picPreview.Width = PgWidth * Ratio
            picPreview.Container.ScaleMode = smtemp
         Else
            Ratio = WidthRatio
            smtemp = picPreview.Container.ScaleMode
            picPreview.Container.ScaleMode = vbInches
            picPreview.Height = PgHeight * Ratio
            picPreview.Container.ScaleMode = smtemp
         End If

         ' Set default properties of picture box to match printer
         ' There are many that you could add here:
         picPreview.Scale (0, 0)-(PgWidth, PgHeight)
         picPreview.Font.Name = Printer.Font.Name
         picPreview.FontSize = Printer.FontSize * Ratio
         picPreview.ForeColor = Printer.ForeColor
         picPreview.Cls

         ScalePicPreviewToPrinterInches = Ratio
End Function

      
Private Sub PrintRoutineAll(objPrint As Object, Optional Ratio As Double = 1)
         ' All dimensions in inches:
         Dim xPosition As Double      'horizontal (or left) position of picture
         Dim yPosition As Double      'vertical (or top) position of picture
         xPosition = 0.08
         yPosition = 0.15
         Dim i As Integer
         
         Dim picWidth As Double       'picture width
         Dim picHeight As Double      'picture height
         picWidth = 3.9
         picHeight = 2.48
         
         Dim xSpacing As Double       'horizontal spacing bewtween pictures
         Dim ySpacing As Double       'vertical spacing between pictures
         xSpacing = 0.32
         ySpacing = 0.25
         
     For i = 0 To 7
     
         ' Print some graphics to the control object
         frmMain.picDisplay(i).Picture = frmMain.picDisplay(i).Image
         PicSrc.Picture = frmMain.picDisplay(i).Picture
         'object.PaintPicture picture, x1, y1, width1, height1, x2, y2, width2, height2, opcode   '<-- general format
         objPrint.PaintPicture PicSrc.Picture, xPosition, yPosition, picWidth, picHeight
         
         xPosition = xPosition + (picWidth + xSpacing)       'next picture moves in the x direction

         If xPosition >= 8 Then                              'if xPosition is greater than 8 in., then start a new row
            xPosition = 0.08                                 'new row so x starts a the beginning
            yPosition = yPosition + (picHeight + ySpacing)   'y moves down one row
         End If
         
     Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub
