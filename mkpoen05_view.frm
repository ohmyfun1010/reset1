VERSION 5.00
Object = "{3B930683-5AF1-4F07-9CE8-CA8063E1F3DD}#2.0#0"; "XButton.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form mkpoen05_view 
   Caption         =   "Ã·ºÎÆÄÀÏ ºä(mkpoen05_view)"
   ClientHeight    =   12285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   12285
   ScaleWidth      =   11865
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   11805
      Left            =   45
      TabIndex        =   0
      Top             =   450
      Width           =   11760
      ExtentX         =   20743
      ExtentY         =   20823
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin Threed.SSPanel SSPanel20 
      Height          =   510
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   14595
      _Version        =   65536
      _ExtentX        =   25744
      _ExtentY        =   900
      _StockProps     =   15
      BackColor       =   -2147483637
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txt_filenm 
         Appearance      =   0  'Æò¸é
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  '»ç¿ë ¸øÇÔ
         Left            =   1410
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   4
         Top             =   45
         Width           =   3345
      End
      Begin VB.CommandButton cmb_alldel 
         Caption         =   "Àü Ã¼  »è Á¦"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   12825
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   90
         Width           =   1545
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   330
         Left            =   60
         TabIndex        =   3
         Top             =   45
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   582
         _StockProps     =   15
         Caption         =   "ÆÄÀÏÀÌ¸§"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XLibrary_XButton.XButton cmb_exit 
         Height          =   375
         Left            =   10935
         TabIndex        =   5
         Top             =   45
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   661
         BackColor1      =   12632256
         BackColor2      =   16777215
         BackColorEx     =   12632319
         BackGradientStyle=   2
         BackStyle       =   4
         BevelHeight     =   5
         BackGradientExPercent=   80
         BackGlassColorStyle=   1
         BackGradientAutoValue=   40
         BackGlassAutoValue=   70
         BackLightShadowShadowValue=   -30
         BackLightShadowLightValue=   30
         BorderStyle     =   1
         BorderWidth     =   1
         BorderColor     =   8421504
         EnabledColor    =   6579300
         MaskColor       =   13828096
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Á¾  ·á"
         TextWidthPos    =   2
         TextHeightPos   =   2
         TextWidthMargin =   5
         TextHeightMargin=   5
         IconPosition    =   2
         IconAndTextMargin=   0
         IconMaskColor   =   13828096
         MouseOverMargin =   2
         MouseOverEffectAutoValue=   -20
         MouseDownBorderEffectValue=   -40
         MouseDownDefaultValue=   20
         FocusDefaultMargin=   3
         FocusColor1     =   16777152
         FocusColor2     =   16777088
         FocusColorStyle =   1
         FocusColorMargin=   2
         FocusEffectAutoValue=   -20
         ToolTipBodyText =   "Á¶È¸"
         ToolTipTitleText=   ""
         ToolTipCentered =   -1  'True
         ToolTipBackColor=   12648447
         ToolTipExBackColor1=   12648447
         ToolTipExHoverTime=   1000
         ToolTipExPopupTime=   10000
         ToolTipExPopupPos=   0
         ToolTipExArrowWidth=   10
         ToolTipExArrowHeight=   15
         ToolTipExBorderRoundNum=   0
         ToolTipExPopupPosWMargin=   5
         ToolTipExPopupPosHMargin=   5
         ToolTipExBackColor2=   16777215
         ToolTipExBorderColor=   4210752
         ToolTipExTitleText=   "Title"
         ToolTipExIconAndTitleMargin=   5
         ToolTipExTitleAlign=   2
         BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTipExTopMargin=   5
         ToolTipExBottomMargin=   5
         ToolTipExLeftMargin=   5
         ToolTipExRightMargin=   5
         ToolTipExBodyText=   "Body Text"
         ToolTipExBodyTextColor=   4210752
         BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTipExTitleLineColor=   4210752
         ToolTipExTitleAndLineMargin=   5
         ToolTipExPostScriptText=   "PostScript"
         ToolTipExIconAndPostScriptMargin=   5
         ToolTipExPostScriptLineColor=   4210752
         BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTipExTitleLineShadow=   -1  'True
         ToolTipExTitleLine=   -1  'True
         ToolTipExTitleLineLeftMargin=   5
         ToolTipExTitleLineRightMargin=   5
         ToolTipExPostScriptLineShadow=   -1  'True
         ToolTipExPostScriptLine=   -1  'True
         ToolTipExPostScriptLineLeftMargin=   5
         ToolTipExPostScriptLineRightMargin=   5
         ToolTipExTitleAndBodyMargin=   5
         ToolTipExBodyAndPostScriptMargin=   5
         ToolTipExTitleTextBackColor=   16777215
         ToolTipExTitleIconMaskColor=   13828096
         ToolTipExTitleIconAndTextAlign=   2
         ToolTipExTitleIconAndTextMargin=   5
         ToolTipExPopupAutoPos=   -1  'True
         ToolTipExPostScriptAndLineMargin=   5
         ToolTipExPostScriptIconPos=   1
         ToolTipExPostScriptIconAndTextMargin=   5
         ToolTipExPostScriptIconAndTextAlign=   2
         ToolTipExPostScriptIconMaskColor=   13828096
         ToolTipExBodyTextBackColor=   16761024
      End
   End
End
Attribute VB_Name = "mkpoen05_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_exit_Click()
    Unload Me
End Sub




   
   


