VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmGDIPlusLOGO 
   Caption         =   "GDI+ Unicode Logo Generator"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   516
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   646
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Clipboard"
      Height          =   375
      Left            =   7560
      TabIndex        =   49
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7560
      TabIndex        =   48
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CheckBox chkEllipse 
      Caption         =   "Ellipse"
      Height          =   255
      Left            =   6480
      TabIndex        =   47
      Top             =   1320
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.TextBox TxtSize 
      Height          =   285
      Index           =   0
      Left            =   4980
      TabIndex        =   43
      Text            =   "640"
      ToolTipText     =   "Horizontal Logosize"
      Top             =   1260
      Width           =   615
   End
   Begin VB.TextBox TxtSize 
      Height          =   285
      Index           =   1
      Left            =   5700
      TabIndex        =   42
      Text            =   "256"
      ToolTipText     =   "Vertical Logosize"
      Top             =   1260
      Width           =   615
   End
   Begin VB.VScrollBar VscrY 
      Height          =   285
      Left            =   6180
      Max             =   -20
      Min             =   -2000
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1260
      Value           =   -20
      Width           =   135
   End
   Begin VB.VScrollBar VscrX 
      Height          =   285
      Left            =   5460
      Max             =   -20
      Min             =   -2000
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1260
      Value           =   -20
      Width           =   135
   End
   Begin VB.Frame Frame2 
      Caption         =   "Step 2 - Select Object"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      TabIndex        =   28
      Top             =   960
      Width           =   4635
      Begin VB.OptionButton optObject 
         Caption         =   "Text Fill"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   32
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optObject 
         Caption         =   "Text Border"
         Height          =   255
         Index           =   2
         Left            =   2340
         TabIndex        =   31
         Top             =   300
         Width           =   1155
      End
      Begin VB.OptionButton optObject 
         Caption         =   "Background"
         Height          =   255
         Index           =   1
         Left            =   1020
         TabIndex        =   30
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton optObject 
         Caption         =   "Border"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Step 3 - Set Object Properties (Loop to Step 2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   60
      TabIndex        =   15
      Top             =   1620
      Width           =   9555
      Begin VB.TextBox txtAlpha 
         Height          =   225
         Index           =   1
         Left            =   4140
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "255"
         ToolTipText     =   "Size of the Border"
         Top             =   1860
         Width           =   375
      End
      Begin VB.TextBox txtAlpha 
         Height          =   225
         Index           =   0
         Left            =   2640
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "255"
         ToolTipText     =   "Size of the Border"
         Top             =   1860
         Width           =   375
      End
      Begin VB.CheckBox chkGamma 
         Caption         =   "Gamma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   35
         ToolTipText     =   "Gamma Correction"
         Top             =   1860
         Width           =   975
      End
      Begin VB.TextBox txtBorder 
         Height          =   225
         Left            =   960
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "4"
         ToolTipText     =   "Border Size"
         Top             =   1860
         Width           =   375
      End
      Begin VB.FileListBox File 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   7620
         Pattern         =   "*.gif;*.jpg;*.bmp"
         TabIndex        =   21
         ToolTipText     =   "Texture Files"
         Top             =   480
         Width           =   1755
      End
      Begin VB.ComboBox cboStyle 
         Appearance      =   0  'Flat
         Height          =   1350
         ItemData        =   "frmGDIPlusLOGO.frx":0000
         Left            =   120
         List            =   "frmGDIPlusLOGO.frx":0013
         Style           =   1  'Simple Combo
         TabIndex        =   20
         Text            =   "cboStyle"
         Top             =   480
         Width           =   1395
      End
      Begin VB.ComboBox cboColor 
         Appearance      =   0  'Flat
         Height          =   1350
         Index           =   0
         ItemData        =   "frmGDIPlusLOGO.frx":004C
         Left            =   1620
         List            =   "frmGDIPlusLOGO.frx":058E
         Style           =   1  'Simple Combo
         TabIndex        =   19
         Text            =   "cboColor"
         Top             =   480
         Width           =   1395
      End
      Begin VB.ComboBox cboColor 
         Appearance      =   0  'Flat
         Height          =   1350
         Index           =   1
         ItemData        =   "frmGDIPlusLOGO.frx":0B94
         Left            =   3120
         List            =   "frmGDIPlusLOGO.frx":10D6
         Style           =   1  'Simple Combo
         TabIndex        =   18
         Text            =   "cboColor"
         Top             =   480
         Width           =   1395
      End
      Begin VB.ComboBox cboHatch 
         Appearance      =   0  'Flat
         Height          =   1350
         ItemData        =   "frmGDIPlusLOGO.frx":16DC
         Left            =   6120
         List            =   "frmGDIPlusLOGO.frx":17AA
         Style           =   1  'Simple Combo
         TabIndex        =   17
         Text            =   "cboHatch"
         Top             =   480
         Width           =   1395
      End
      Begin VB.ComboBox cboGradient 
         Appearance      =   0  'Flat
         Height          =   1350
         ItemData        =   "frmGDIPlusLOGO.frx":1E19
         Left            =   4620
         List            =   "frmGDIPlusLOGO.frx":1E29
         Style           =   1  'Simple Combo
         TabIndex        =   16
         Text            =   "cboGradient"
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label7 
         Caption         =   "End Alpha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   39
         ToolTipText     =   "Transparency 0-255"
         Top             =   1860
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Start Alpha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1620
         TabIndex        =   37
         ToolTipText     =   "Transparency 0-255"
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Label lblBorder 
         Caption         =   "Border"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label lblTexture 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Texture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8100
         TabIndex        =   27
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblStyle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Pattern Style"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   26
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblColor1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Start Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   25
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lblColor2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "End Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3465
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Hatch Style"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6315
         TabIndex        =   23
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Gradient Style"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4680
         TabIndex        =   22
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Frame FrmTab 
      Caption         =   "Step 1 - Set Text/Font Properties"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   9555
      Begin VB.CheckBox chkFont 
         Caption         =   "Strikethru"
         Height          =   255
         Index           =   3
         Left            =   7440
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkFont 
         Caption         =   "Underline"
         Height          =   255
         Index           =   2
         Left            =   7440
         TabIndex        =   13
         Top             =   180
         Width           =   975
      End
      Begin VB.CheckBox chkFont 
         Caption         =   "Italic"
         Height          =   255
         Index           =   1
         Left            =   6780
         TabIndex        =   12
         Top             =   480
         Value           =   1  'Checked
         Width           =   675
      End
      Begin VB.CheckBox chkFont 
         Caption         =   "Bold"
         Height          =   255
         Index           =   0
         Left            =   6780
         TabIndex        =   11
         Top             =   180
         Value           =   1  'Checked
         Width           =   675
      End
      Begin VB.CheckBox chkUnicode 
         Caption         =   "Unicode"
         Height          =   255
         Left            =   8460
         TabIndex        =   6
         Top             =   480
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.ComboBox cboAlign 
         Height          =   315
         ItemData        =   "frmGDIPlusLOGO.frx":1E66
         Left            =   3300
         List            =   "frmGDIPlusLOGO.frx":1E85
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Font Alignment"
         Top             =   450
         Width           =   1275
      End
      Begin VB.ComboBox cboFontSize 
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Text            =   "cboSize"
         ToolTipText     =   "Fontsize"
         Top             =   450
         Width           =   975
      End
      Begin VB.ComboBox cboFontName 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Fontname"
         Top             =   450
         Width           =   2100
      End
      Begin VB.CheckBox chkAlias 
         Caption         =   "Anti Alias"
         Height          =   255
         Left            =   8460
         TabIndex        =   2
         Top             =   180
         Value           =   1  'Checked
         Width           =   975
      End
      Begin MSForms.TextBox txtLogo 
         Height          =   675
         Left            =   4920
         TabIndex        =   46
         Top             =   180
         Width           =   1815
         VariousPropertyBits=   -1400879077
         Size            =   "3201;1191"
         Value           =   "Simplified Chinese"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label lblText 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4500
         TabIndex        =   10
         Top             =   180
         Width           =   405
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   825
         TabIndex        =   9
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   8
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Alignment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3420
         TabIndex        =   7
         Top             =   240
         Width           =   840
      End
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   8760
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   60
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   3840
      Width           =   9600
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Size Y"
      Height          =   195
      Left            =   5760
      TabIndex        =   45
      Top             =   1020
      Width           =   450
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Size X"
      Height          =   195
      Left            =   5040
      TabIndex        =   44
      Top             =   1020
      Width           =   450
   End
End
Attribute VB_Name = "frmGDIPlusLOGO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Four objects used:
'0 = Border
'1 = Background
'2 = Text Border
'3 = Text Fill
Dim GpInput             As GdiplusStartupInput
Dim Locked              As Boolean
Dim Init                As Boolean
Dim m_Border(3)         As Integer 'only m_Border(0) and m_Border(2) are used
Dim m_BackColor         As Long
Dim m_BackColorIdx      As Long
Dim m_Logo              As String
Dim m_CurrentObject     As Long
Dim m_EndAlpha(3)       As Long
Dim m_EndColor(3)       As Long
Dim m_EndColorIdx(3)    As Long
Dim m_EndFinal(3)       As Long
Dim m_Gamma(3)          As Boolean
Dim m_GradientMode(3)   As LinearGradientMode
Dim m_HatchStyle(3)     As HatchStyle
Dim m_StartAlpha(3)     As Long
Dim m_StartColor(3)     As Long
Dim m_StartColorIdx(3)  As Long
Dim m_StartFinal(3)     As Long
Dim m_Style(3)          As BrushType
Dim m_Texture(3)        As String
Dim m_TextureIndex(3)   As Long
Dim rct                 As RECTF
Dim TexturePath         As String
Dim token               As Long

Private Sub CboAlign_Change()
   RefreshLogo
End Sub

Private Sub CboAlign_Click()
   CboAlign_Change
End Sub

Private Sub cboColor_Change(Index As Integer)
   Select Case Index
      Case 1
         m_EndColor(m_CurrentObject) = cboColor(1).ItemData(cboColor(1).ListIndex)
         m_EndColorIdx(m_CurrentObject) = cboColor(1).ListIndex
      Case 0
         m_StartColor(m_CurrentObject) = cboColor(0).ItemData(cboColor(0).ListIndex)
         m_StartColorIdx(m_CurrentObject) = cboColor(0).ListIndex
      Case 2
         m_BackColorIdx = cboColor(2).ListIndex
         m_BackColor = Gdi2VbColor(cboColor(2).ItemData(m_BackColorIdx))
   End Select

   RefreshLogo
End Sub

Private Sub cboColor_Click(Index As Integer)
   cboColor_Change Index
End Sub

Private Sub cboGradient_Change()
   m_GradientMode(m_CurrentObject) = cboGradient.ListIndex
   RefreshLogo
End Sub

Private Sub cboGradient_Click()
   cboGradient_Change
End Sub

Private Sub cboHatch_Change()
   m_HatchStyle(m_CurrentObject) = cboHatch.ListIndex
   RefreshLogo
End Sub

Private Sub cboHatch_Click()
   cboHatch_Change
End Sub

Private Sub cboFontName_Change()
   picLogo.FontName = cboFontName.Text
   RefreshLogo
End Sub

Private Sub cboFontName_Click()
   cboFontName_Change
End Sub

Private Sub cboFontSize_Change()
   picLogo.FontSize = Val(cboFontSize)
   RefreshLogo
End Sub

Private Sub cboFontSize_Click()
   cboFontSize_Change
End Sub

Private Sub cboStyle_Change()
   m_Style(m_CurrentObject) = cboStyle.ListIndex
   SetVisibility m_Style(m_CurrentObject)
   RefreshLogo
End Sub

Private Sub cboStyle_Click()
   cboStyle_Change
End Sub

Private Sub ChkAlias_Click()
   RefreshLogo
End Sub

Private Sub chkEllipse_Click()
   RefreshLogo
End Sub

Private Sub chkFont_Click(Index As Integer)
   picLogo.FontBold = chkFont(0)
   picLogo.FontItalic = chkFont(1)
   picLogo.FontUnderline = chkFont(2)
   picLogo.FontStrikethru = chkFont(3)
   RefreshLogo
End Sub

Private Sub chkGamma_Click()
   m_Gamma(m_CurrentObject) = chkGamma * -1
   RefreshLogo
End Sub

Private Sub chkUnicode_Click()
   RefreshLogo
End Sub

Private Sub cmdClipboard_Click()
   Clipboard.Clear
   'copy the Image to the Clipboard
   Clipboard.SetData picLogo.Image
End Sub

Private Sub cmdSave_Click()
   With cmd
      On Error GoTo EndSave
      .FileName = vbNullString
      .Filter = "Bitmap|*.bmp"
      .ShowSave
      If LenB(.FileName) Then
         SavePicture picLogo.Image, .FileName
      End If
   End With
EndSave:
End Sub

Private Sub File_Click()
   m_Texture(m_CurrentObject) = TexturePath & File.FileName
   m_TextureIndex(m_CurrentObject) = File.ListIndex
   RefreshLogo
End Sub

Private Sub Form_Load()
   Dim StartFont        As Integer
   Dim i                As Integer
   Dim UniCrLf          As String
   Dim GpInput          As GdiplusStartupInput
   Dim fontFam          As Long
   Dim strFontName      As String
   Dim FS               As Long
   Dim str              As String
   Dim IsAvailable      As Long
   
   UniCrLf = Chr$(13) & Chr$(0) & Chr$(10) & Chr$(0)

   Me.Show
   Init = True

   rct.Width = picLogo.ScaleWidth
   rct.Height = picLogo.ScaleHeight
   TexturePath = AppPath & "textures\"

   m_Logo = StrConv("Simplified", vbUnicode) & UniCrLf & _
      StrConv("Chinese", vbUnicode) & UniCrLf & "ˇ˛sS`OegŸè?QÑvÓvÑv"
   txtLogo.Text = StrConv(m_Logo, vbFromUnicode)

   ' Load the GDI+ Dll
   GpInput.GdiplusVersion = 1
   If GdiplusStartup(token, GpInput) <> Ok Then
      MsgBox "Error loading GDI+!", vbCritical
      Unload Me
   End If
   
   'Get Supported Font names and set them to the Font combo Box
   For i = 0 To Screen.FontCount - 1
      strFontName = Screen.Fonts(i)
      GdipCreateFontFamilyFromName strFontName, 0, fontFam
      GdipIsStyleAvailable fontFam, FS, IsAvailable
      If IsAvailable Then
         cboFontName.AddItem strFontName
      End If
   Next
   
   ' Shutdown the GDI+ Dll
   GdiplusShutdown (token)
      
   For i = 0 To cboFontName.ListCount
      'Find Font in the sorted list
      If cboFontName.List(i) = "BlackChancery" Then
         StartFont = i
         Exit For
      ElseIf cboFontName.List(i) = "Times New Roman" Then
         StartFont = i
      End If
   Next
   'Set the standard Fontsizes
   For i = 12 To 256 Step 4
      cboFontSize.AddItem i
   Next

   File.path = TexturePath

   'preset alpha values, border sizes, textures
   For i = 0 To 3
      m_StartAlpha(i) = 255
      m_EndAlpha(i) = 255
      optObject(i).Value = True
      m_Style(i) = BrushTypeLinearGradient
      m_GradientMode(i) = LinearGradientModeVertical
      m_HatchStyle(i) = HatchStyleSmallGrid
      m_Border(i) = Choose(i + 1, 8, 0, 3, 0)
      File.ListIndex = Choose(i + 1, 11, 4, 18, 45)
      m_StartColor(i) = Black: m_StartColorIdx(i) = 7
      m_EndColor(i) = White: m_EndColorIdx(i) = 137
      'm_StartColor(i) = Red: m_StartColorIdx(i) = 113
      'm_EndColor(i) = Yellow: m_EndColorIdx(i) = 139
   Next

   m_EndColor(0) = LightSteelBlue: m_EndColorIdx(0) = 74
   m_StartColor(3) = White: m_StartColorIdx(3) = 137
   m_EndColor(3) = Black: m_EndColorIdx(3) = 7

   'Now activate the Settings
   cboFontName.ListIndex = StartFont 'Start Font
   cboFontSize = 48                  'Fontsize

   'picLogo.BackColor = vbBlack: m_BackColorIdx = 7
   'm_BackColor = vbBlack

   m_Style(3) = BrushTypeTextureFill
   m_Style(2) = BrushTypeSolidColor
   m_Style(1) = BrushTypeTextureFill
   m_Style(0) = BrushTypeLinearGradient
   'm_GradientMode(0) = LinearGradientModeHorizontalTri

   cboColor(0).ListIndex = 0
   cboColor(1).ListIndex = 0

   cboAlign.ListIndex = 4        'Font Align is Center

   optObject_Click 3
   Init = False
   RefreshLogo
End Sub

Private Sub optObject_Click(Index As Integer)
   'Disable updates
   Locked = True
   m_CurrentObject = Index
   'Set indices
   File.ListIndex = m_TextureIndex(Index)
   cboStyle.ListIndex = m_Style(Index)
   cboHatch.ListIndex = m_HatchStyle(Index)
   cboGradient.ListIndex = m_GradientMode(Index)
   cboColor(0).ListIndex = m_StartColorIdx(Index)
   cboColor(1).ListIndex = m_EndColorIdx(Index)
   txtBorder = m_Border(Index)
   txtAlpha(0) = m_StartAlpha(Index)
   txtAlpha(1) = m_EndAlpha(Index)
   chkGamma = Abs(m_Gamma(Index))
   'Hide border property for fills
   txtBorder.Enabled = Index Mod 2 = 0
   lblBorder.Enabled = txtBorder.Enabled

   SetVisibility m_Style(Index)

   'Enable Drawing
   Locked = False
   RefreshLogo

End Sub

Private Sub SetVisibility(curStyle As BrushType)
   'Reduce screen clutter by setting the visibility of
   'combo boxes that apply only to current brush type
   
   File.Visible = curStyle = BrushTypeTextureFill
   cboColor(0).Visible = Not (File.Visible)
   cboHatch.Visible = curStyle = BrushTypeHatchFill
   cboGradient.Visible = curStyle = BrushTypeLinearGradient
   
   If curStyle = BrushTypeSolidColor Or curStyle = BrushTypeTextureFill Then
      cboColor(1).Visible = False
   Else
      cboColor(1).Visible = True
   End If
   
   'Alpha visibility matches Start/End color combos
   txtAlpha(0).Visible = cboColor(0).Visible
   txtAlpha(1).Visible = cboColor(1).Visible
   
End Sub

Private Sub RefreshLogo()

   Dim i                As Integer
   Dim TR               As RECTF
   If (Not Locked) And (Not Init) Then
      For i = 0 To 3
         m_StartFinal(i) = ColorSetAlpha(m_StartColor(i), m_StartAlpha(i))
         m_EndFinal(i) = ColorSetAlpha(m_EndColor(i), m_EndAlpha(i))
      Next

      'Inflate rectangle inward for BorderWidth
      'Since penalignment default is center this actually
      'leaves a border of half the pen width.
      LSet TR = rct
      InflateRectF TR, -m_Border(0), -m_Border(0)

      picLogo.Cls

      ' Load the GDI+ Dll
      GpInput.GdiplusVersion = 1
      If GdiplusStartup(token, GpInput) <> Ok Then
         MsgBox "Error loading GDI+!", vbCritical
         Unload Me
      End If

      DrawGdipLogo picLogo.hDC, _
         TR, chkEllipse.Value = 1, m_Border(), m_Style(), _
         m_GradientMode(), m_HatchStyle(), _
         m_Logo, chkAlias.Value = 1, _
         picLogo, _
         cboAlign.ListIndex, _
         m_Gamma(), _
         m_StartFinal(), m_EndFinal(), _
         m_Texture(), chkUnicode, 0

      ' Shutdown the GDI+ Dll
      GdiplusShutdown (token)
   End If
End Sub

Private Sub txtAlpha_Change(Index As Integer)
   Dim iVal             As Integer

   iVal = Val(txtAlpha(Index))
   If iVal > 255 Then
      txtAlpha(Index) = 255
   ElseIf iVal < 0 Then
      txtAlpha(Index) = 0
   End If

   If Index = 0 Then 'Start
      m_StartAlpha(m_CurrentObject) = Val(txtAlpha(0))
      m_StartFinal(m_CurrentObject) = m_StartColor(m_CurrentObject) And _
         m_StartAlpha(m_CurrentObject)
   ElseIf Index = 1 Then 'End
      m_EndAlpha(m_CurrentObject) = Val(txtAlpha(1))
      m_EndFinal(m_CurrentObject) = m_EndColor(m_CurrentObject) And _
         m_EndAlpha(m_CurrentObject)
   End If

   RefreshLogo

End Sub

Private Sub TxtBorder_Change()
   If m_CurrentObject = 0 Or m_CurrentObject = 2 Then
      m_Border(m_CurrentObject) = Val(txtBorder)
      RefreshLogo
   End If
End Sub

Private Sub txtLogo_Change()
   m_Logo = txtLogo.Text
   RefreshLogo
End Sub

Private Sub TxtSize_Change(Index As Integer)
   Dim iVal             As Integer
   Dim iMax             As Integer

   iMax = IIf(Index, 256, 640)
   iVal = Val(TxtSize(Index))

   If iVal > iMax Then
      TxtSize(Index) = iMax
   End If

   picLogo.Move 4 + (640 - TxtSize(0)) \ 2, _
      256 + (256 - TxtSize(1)) \ 2, _
      TxtSize(0), _
      TxtSize(1)
   rct.Width = picLogo.ScaleWidth
   rct.Height = picLogo.ScaleHeight
   RefreshLogo
End Sub

Private Sub TxtSize_KeyPress(Index As Integer, KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) Then
      KeyAscii = 0
   End If
End Sub

