VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Image Effects"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   426
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Effects"
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   2460
      Width           =   3015
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   2775
         TabIndex        =   13
         Top             =   240
         Width           =   2775
         Begin VB.OptionButton Option1 
            Caption         =   "Alpha Blend"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   180
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Transparent Blt"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   420
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "GrayScale"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   660
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Saturation"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   900
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Color Fill"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   1140
            Width           =   1500
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1815
      Left            =   3240
      TabIndex        =   6
      Top             =   2460
      Width           =   3015
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   2835
         TabIndex        =   7
         Top             =   240
         Width           =   2835
         Begin VB.OptionButton Option2 
            Caption         =   "Icons.bmp "
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   1755
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Auto apply changes"
            Height          =   195
            Left            =   60
            TabIndex        =   9
            Top             =   180
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Flowers.jpg"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Value           =   -1  'True
            Width           =   1755
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Left picture box image file:"
            Height          =   195
            Left            =   60
            TabIndex        =   11
            Top             =   540
            Width           =   1920
         End
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Screen Shot"
      Height          =   375
      Left            =   3660
      TabIndex        =   5
      Top             =   4620
      Width           =   1275
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   5130
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5609
            Text            =   "Ready"
            TextSave        =   "Ready"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Value: 128"
            TextSave        =   "Value: 128"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Elapsed: 0"
            TextSave        =   "Elapsed: 0"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   4620
      Width           =   1215
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Value (right-click to reset)"
      Top             =   4620
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   327682
      LargeChange     =   25
      Max             =   255
      SelStart        =   128
      TickStyle       =   3
      TickFrequency   =   26
      Value           =   128
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   2250
      Left            =   3240
      ScaleHeight     =   146
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   1
      Top             =   120
      Width           =   3000
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2250
      Left            =   120
      ScaleHeight     =   146
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   0
      Top             =   120
      Width           =   3000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   8
      X2              =   416
      Y1              =   297
      Y2              =   297
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   8
      X2              =   416
      Y1              =   296
      Y2              =   296
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************
' Module:     Form1
' DateTime:   06/07/2006 13:34
' Author:     BioHazardMX
' Purpose:    cImageFX Demonstration Form
'***************************************************************************************

Private hMod As Long
Private WithEvents mImageFX As cImageFX
Attribute mImageFX.VB_VarHelpID = -1
Private cValue As Byte, cFX As Long

Private Sub Check1_Click()
  Command1.Enabled = Not CBool(Check1.Value)
End Sub

Private Sub Command1_Click()
  Picture2.Cls
  StatusBar1.Panels(1).Text = "Processing..."
  Select Case cFX
    Case 0
      mImageFX.FxAlphaBlend Picture2.hDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, cValue, vbMagenta
      Slider1.Enabled = True
    Case 1
      mImageFX.FxTransBlt Picture2.hDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbMagenta
      Slider1.Enabled = False
    Case 2
      mImageFX.FxGrayScale Picture2.hDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbMagenta
      Slider1.Enabled = False
    Case 3
      mImageFX.FxSaturation Picture2.hDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, cValue, vbMagenta
      Slider1.Enabled = True
    Case 4
      mImageFX.FxColorFill Picture2.hDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, cValue, vbHighlight, vbMagenta
      Slider1.Enabled = True
  End Select
  Picture2.Refresh
End Sub

Private Sub Command2_Click()
  Picture2.Cls
  StatusBar1.Panels(1).Text = "Working..."
  mImageFX.FxScreenShot Picture2.hDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
  Picture2.Refresh
End Sub

Private Sub Form_Initialize()
  hMod = LoadLibrary("Shell32.dll")
  InitCommonControls
End Sub
Private Sub Form_Load()
  Set mImageFX = New cImageFX
  Set Picture1.Picture = LoadPicture(App.Path & "\Images\Flowers.jpg")
  Set Picture2.Picture = LoadPicture(App.Path & "\Images\Sunset.jpg")
  Slider1_Scroll
  DisableResize hWnd
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set mImageFX = Nothing
  Call FreeLibrary(hMod)
End Sub

Private Sub mImageFX_Complete(ByVal lTimeMs As Long)
  StatusBar1.Panels(1).Text = "Ready"
  StatusBar1.Panels(3).Text = "Elapsed: " & lTimeMs & " ms"
End Sub

Private Sub Option1_Click(Index As Integer)
  cFX = Index
  If CBool(Check1.Value) Then Command1_Click
End Sub

Private Sub Option2_Click(Index As Integer)
  If Index = 0 Then Set Picture1.Picture = LoadPicture(App.Path & "\Images\Flowers.jpg")
  If Index = 1 Then Set Picture1.Picture = LoadPicture(App.Path & "\Images\Icons.bmp")
  If CBool(Check1.Value) Then Command1_Click
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then Slider1.Value = 128
  Slider1_Scroll
  If CBool(Check1.Value) Then Command1_Click
End Sub

Private Sub Slider1_Scroll()
  StatusBar1.Panels(2).Text = "Value: " & Slider1.Value
  cValue = Slider1.Value
End Sub
