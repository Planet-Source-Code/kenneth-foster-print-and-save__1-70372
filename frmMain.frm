VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                     Select and Print Pictures"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5100
      TabIndex        =   15
      Top             =   7755
      Width           =   1200
   End
   Begin Project1.StrokeText StrokeText1 
      Height          =   840
      Left            =   780
      TabIndex        =   13
      Top             =   8460
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   1482
      Caption         =   "Print and Save"
      Shadow          =   -1  'True
      TStyle          =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Albert"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picPreview 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   9015
      ScaleHeight     =   1155
      ScaleWidth      =   1575
      TabIndex        =   11
      Top             =   3495
      Width           =   1575
   End
   Begin VB.CommandButton cmdSPP 
      Caption         =   "Show Print Preview"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2280
      TabIndex        =   10
      Top             =   7755
      Width           =   2490
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1680
      Index           =   7
      Left            =   3570
      ScaleHeight     =   1680
      ScaleWidth      =   2640
      TabIndex        =   9
      Top             =   5805
      Width           =   2640
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1680
      Index           =   6
      Left            =   720
      ScaleHeight     =   1680
      ScaleWidth      =   2640
      TabIndex        =   8
      Top             =   5805
      Width           =   2640
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1680
      Index           =   5
      Left            =   3570
      ScaleHeight     =   1680
      ScaleWidth      =   2640
      TabIndex        =   7
      Top             =   3960
      Width           =   2640
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1680
      Index           =   4
      Left            =   720
      ScaleHeight     =   1680
      ScaleWidth      =   2640
      TabIndex        =   6
      Top             =   3960
      Width           =   2640
   End
   Begin VB.CommandButton cmdAddPix 
      Caption         =   "Get Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   615
      TabIndex        =   1
      Top             =   7755
      Width           =   1395
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00FFFFFF&
      Height          =   7440
      Left            =   600
      ScaleHeight     =   7380
      ScaleWidth      =   5640
      TabIndex        =   0
      Top             =   195
      Width           =   5700
      Begin VB.PictureBox picDisplay 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1680
         Index           =   3
         Left            =   2940
         ScaleHeight     =   1680
         ScaleWidth      =   2640
         TabIndex        =   5
         Top             =   1905
         Width           =   2640
      End
      Begin VB.PictureBox picDisplay 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1680
         Index           =   2
         Left            =   90
         ScaleHeight     =   1680
         ScaleWidth      =   2640
         TabIndex        =   4
         Top             =   1905
         Width           =   2640
      End
      Begin VB.PictureBox picDisplay 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1680
         Index           =   1
         Left            =   2940
         ScaleHeight     =   1680
         ScaleWidth      =   2640
         TabIndex        =   3
         Top             =   60
         Width           =   2640
      End
      Begin VB.PictureBox picDisplay 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1680
         Index           =   0
         Left            =   90
         ScaleHeight     =   1680
         ScaleWidth      =   2640
         TabIndex        =   2
         Top             =   60
         Width           =   2640
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Height          =   1770
         Index           =   7
         Left            =   2910
         Top             =   5550
         Width           =   2715
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Height          =   1770
         Index           =   6
         Left            =   60
         Top             =   5550
         Width           =   2715
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Height          =   1755
         Index           =   5
         Left            =   2910
         Top             =   3705
         Width           =   2715
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Height          =   1755
         Index           =   4
         Left            =   60
         Top             =   3705
         Width           =   2715
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Height          =   1755
         Index           =   3
         Left            =   2910
         Top             =   1875
         Width           =   2715
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Height          =   1755
         Index           =   2
         Left            =   60
         Top             =   1875
         Width           =   2715
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Height          =   1755
         Index           =   1
         Left            =   2910
         Top             =   30
         Width           =   2715
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Height          =   1755
         Index           =   0
         Left            =   60
         Top             =   30
         Width           =   2715
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Dbl click on picture position or press 'Get Picture' button."
      Height          =   225
      Left            =   1560
      TabIndex        =   14
      Top             =   8355
      Width           =   4065
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Preview of Current           Selection"
      Height          =   390
      Left            =   9045
      TabIndex        =   12
      Top             =   6165
      Width           =   1395
   End
   Begin VB.Image Image1 
      Height          =   1155
      Left            =   9030
      Stretch         =   -1  'True
      Top             =   4740
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
'**                              Print and Save
'**                               Version 1.6.0
'**                               By Ken Foster
'**                                 April 6, 2008
'**                     Freeware--- no copyrights claimed
'*******************************************************************

'=============================================

Option Explicit

Public SelPos As Integer     'stores which picturebox has been selected

Private Sub Form_Load()
   Shape1(0).BorderColor = vbRed      'make first red rectangle red at start up
End Sub

Private Sub cmdAddPix_Click()
   frmEdit.Show
   frmMain.Hide
   frmEdit.cmdResetLeft_Click
   frmEdit.cmdResetMid_Click
End Sub

Private Sub cmdExit_Click()
   Unload frmEdit
   Unload frmPrintPre
   Unload Me
End Sub

Private Sub cmdSPP_Click()
Dim x As Integer
   For x = 0 To 7
      Shape1(x).Visible = False
   Next x
   frmPrintPre.Show
End Sub

Private Sub picDisplay_DblClick(Index As Integer)
   SelPos = Index
   cmdAddPix_Click
End Sub

Private Sub picDisplay_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim xx As Integer

   SelPos = Index
    For xx = 0 To 7                        'highlight selected rectangle in red, all others green
      Shape1(xx).BorderColor = &HC000&
   Next xx
   Shape1(Index).BorderColor = vbRed
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload frmPrintPre
   Unload frmEdit
   Unload Me
End Sub
