VERSION 5.00
Object = "{3B930683-5AF1-4F07-9CE8-CA8063E1F3DD}#2.0#0"; "XButton.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form mkpoen05_print 
   Caption         =   "Form1"
   ClientHeight    =   12165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   12165
   ScaleWidth      =   10545
   Begin Threed.SSPanel SSPanel1 
      Height          =   12135
      Index           =   1
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   10485
      _Version        =   65536
      _ExtentX        =   18494
      _ExtentY        =   21405
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSCommand SSCommand1 
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   600
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "SSCommand1"
      End
      Begin Threed.SSCommand btn_msg 
         Height          =   375
         Left            =   7320
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "쪽지발송"
      End
      Begin VB.CheckBox chk_app 
         Caption         =   "결재자지정"
         Height          =   180
         Left            =   9120
         TabIndex        =   13
         Top             =   660
         Width           =   1215
      End
      Begin FPSpreadADO.fpSpread sprd_print2 
         Height          =   11025
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1320
         Width           =   10170
         _Version        =   458752
         _ExtentX        =   17939
         _ExtentY        =   19447
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   41
         MaxRows         =   41
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "mkpoen05_print.frx":0000
      End
      Begin FPSpreadADO.fpSpread sprd_print3 
         Height          =   465
         Left            =   6225
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Width           =   465
         _Version        =   458752
         _ExtentX        =   820
         _ExtentY        =   820
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   42
         MaxRows         =   43
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "mkpoen05_print.frx":13DEE
      End
      Begin VB.TextBox txt_seq1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2295
         MaxLength       =   3
         TabIndex        =   2
         Top             =   120
         Width           =   510
      End
      Begin VB.TextBox txt_dat1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   1
         Top             =   120
         Width           =   1185
      End
      Begin XLibrary_XButton.XButton btn_print 
         Height          =   435
         Left            =   9090
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   60
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   767
         BackColor1      =   12632256
         BackColor2      =   16777215
         BackColorEx     =   14737632
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
         Text            =   "출 력"
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
         ToolTipBodyText =   "조회"
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
      Begin MSComDlg.CommonDialog Comm1 
         Left            =   6840
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   930
         _Version        =   65536
         _ExtentX        =   1640
         _ExtentY        =   582
         _StockProps     =   15
         Caption         =   "등록번호"
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XLibrary_XButton.XButton btn_view 
         Height          =   435
         Left            =   2880
         TabIndex        =   3
         Top             =   60
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   767
         BackColor1      =   12632256
         BackColor2      =   16777215
         BackColorEx     =   14737632
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
         Text            =   "조 회"
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
         ToolTipBodyText =   "조회"
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
      Begin XLibrary_XButton.XButton btn_clear 
         Height          =   315
         Left            =   8340
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         BackColor1      =   12632256
         BackColor2      =   16777215
         BackColorEx     =   14737632
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
         Text            =   "CLEAR"
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
         ToolTipBodyText =   "조회"
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
      Begin FPSpreadADO.fpSpread sprd_print111 
         Height          =   465
         Left            =   6240
         TabIndex        =   7
         Top             =   135
         Visible         =   0   'False
         Width           =   525
         _Version        =   458752
         _ExtentX        =   926
         _ExtentY        =   820
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   41
         MaxRows         =   41
         ScrollBars      =   2
         SpreadDesigner  =   "mkpoen05_print.frx":266AA
      End
      Begin FPSpreadADO.fpSpread sprd_print222 
         Height          =   465
         Left            =   6240
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   150
         Visible         =   0   'False
         Width           =   570
         _Version        =   458752
         _ExtentX        =   1005
         _ExtentY        =   820
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   41
         MaxRows         =   41
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "mkpoen05_print.frx":389F9
      End
      Begin FPSpreadADO.fpSpread sprd_print1 
         Height          =   11025
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   10245
         _Version        =   458752
         _ExtentX        =   18071
         _ExtentY        =   19447
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   41
         MaxRows         =   41
         ScrollBars      =   2
         SpreadDesigner  =   "mkpoen05_print.frx":4A8A6
      End
      Begin VB.TextBox txt_dae 
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   300
      End
      Begin VB.Label msg 
         BackColor       =   &H00F0F0F0&
         Caption         =   "메시지를 출력합니다!"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   4635
         TabIndex        =   9
         Top             =   195
         Width           =   5580
      End
      Begin VB.Image Image1 
         Height          =   195
         Left            =   4320
         Picture         =   "mkpoen05_print.frx":5E613
         Top             =   180
         Width           =   210
      End
   End
End
Attribute VB_Name = "mkpoen05_print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SERVER_PATH As String = "/공구테스트_DATA/"

Dim aname(4) As String
Dim asab(4) As Integer
Dim mname(4)  As String
Dim msab(4)   As Integer
Dim pcnt As Integer     '어떤 프린트 출력할건지
Dim isab As Integer
Dim iname As String

Private Sub btn_msg_Click()
            
    On Error GoTo err_rtn

    Dim nextsab2 As Integer
    Dim nextmsab(4) As Integer
    Dim title    As String
    Dim resp     As Integer
    Dim ii As Integer

    If Len(txt_dat1) < 8 Then
        msg = "등록번호를 확인하세요."
        Exit Sub
    End If
    
    If Len(txt_seq1) = 0 Then
        msg = "등록번호를 확인하세요."
        Exit Sub
    End If
    
    resp = MsgBox("쪽지를 발송하시겠습니까 ??", vbQuestion + vbYesNo)
    If resp = vbNo Then Exit Sub
    
    'memo 문자열 작업============
    If pcnt = 2 Then
        sprd_print1.row = 4: sprd_print1.Col = 1: title = sprd_print1.Text
    ElseIf pcnt = 3 Then
        sprd_print2.row = 4: sprd_print2.Col = 1: title = sprd_print2.Text
    End If
    
    memo = "TEST DATA 가 등록되었습니다." & vbCrLf
    memo = memo & "확인후 본인 결재란에 결재 바랍니다." & vbCrLf & vbCrLf
    memo = memo & "등록번호: " & txt_dat1 & "-" & txt_seq1 & vbCrLf
    memo = memo & "담당자:" & iname & vbCrLf
    memo = memo & "TEST 제목: " & title & vbCrLf & vbCrLf
    memo = memo & "연결PG:Tol\mkpoen05"
    '============================
    
    For ii = 1 To 4
        nextmsab(ii) = 0
    Next ii

    sss = "       select *"
    sss = sss & "   from oth_applist"
    sss = sss & "  where apl_table = 'man_tooltesthd'"
    sss = sss & "    and apl_tdat = " & txt_dat1
    sss = sss & "    and apl_tseq = " & txt_seq1
    
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    
    If Rs.RecordCount < 1 Then
        Rs.Close
        Exit Sub
    Else
        
        '결재순서: 생산 -> 기술연구소 -> 전무님 -> 관련부서4
        
        '생산,기연,전무님 미결시 미결자 단독 쪽지발송
        If Rs!apl_1sab <> 0 And Rs!apl_1yn = "N" Then
            nextsab2 = Rs!apl_1sab
            GoTo 10
        End If
        
        If Rs!apl_2sab <> 0 And Rs!apl_2yn = "N" Then
            nextsab2 = Rs!apl_2sab
            GoTo 10
        End If
        
        If Rs!apl_3sab <> 0 And Rs!apl_3yn = "N" Then
            nextsab2 = Rs!apl_3sab
            GoTo 10
        End If
        
        '관련부서 미결시 미결자 단체 쪽지발송
        If Rs!apl_m1sab <> 0 And Rs!apl_m1yn = "N" Then
            nextmsab(1) = Rs!apl_m1sab
        End If
        
        If Rs!apl_m2sab <> 0 And Rs!apl_m2yn = "N" Then
            nextmsab(2) = Rs!apl_m2sab
        End If
        
        If Rs!apl_m3sab <> 0 And Rs!apl_m3yn = "N" Then
            nextmsab(3) = Rs!apl_m3sab
        End If
        
        If Rs!apl_m4sab <> 0 And Rs!apl_m4yn = "N" Then
            nextmsab(4) = Rs!apl_m4sab
        End If
        
        GoTo 11
        
    End If
    
10:

    Ws.BeginTrans

    sss = "insert into oth_automsg(ams_time,ams_sab,ams_rsab,ams_sendyn,ams_stime,ams_comment,ams_sprog)"
    sss = sss & " values(sysdate" & ","
    sss = sss & " 9999" & ","
    sss = sss & nextsab2 & ","
    sss = sss & " 'N'" & ","
    sss = sss & " 1" & ","
    sss = sss & "'" & memo & "'" & ","
    sss = sss & " 'TEST DATA 등록')"
    '
    db.Execute sss, 64
    
    Ws.CommitTrans
    
11:
    
    Ws.BeginTrans
    
    For ii = 1 To 4
        
        If nextmsab(ii) > 0 Then
            
            sss = "insert into oth_automsg(ams_time,ams_sab,ams_rsab,ams_sendyn,ams_stime,ams_comment,ams_sprog)"
            sss = sss & " values(sysdate" & ","
            sss = sss & " 9999" & ","
            sss = sss & nextmsab(ii) & ","
            sss = sss & " 'N'" & ","
            sss = sss & " 1" & ","
            sss = sss & "'" & memo & "'" & ","
            sss = sss & " 'TEST DATA 등록')"
            '
            db.Execute sss, 64
            
        End If
        
    Next ii
    
    Ws.CommitTrans
    
    msg = "쪽지가 발송되었습니다."

    Exit Sub

err_rtn:
    Ws.Rollback
    MsgBox (Err.Description)
    
End Sub

Private Sub btn_print_Click()

    Dim docno As String
    Dim kumcd As String
    Dim footer As String
    
    On Error GoTo err_chk
    '
    '
    '-------------------------------------------------------------------------------------
    '문서번호 - 세로
    '-------------------------------------------------------------------------------------
    Dim HLFNO As String
  
    'Printer.FontName = "Times New Roman"
    HLFNO = ""
    HLFNO = HLF_NO("MKPOEN05", 1)
                                                                    
    'Printer.FontSize = 7: Printer.CurrentX = 6: Printer.CurrentY = 280: Printer.Print HLFNO
    'Printer.FontSize = 9: Printer.CurrentX = 85: Printer.CurrentY = 280: Printer.Print "HY-LOK CORPORATION"
    'Printer.FontSize = 7: Printer.CurrentX = 170: Printer.CurrentY = 280: Printer.Print "A4(210mmx297mm)"
  
    '-------------------------------------------------------------------------------------
    'Printer.FontSize = 10
    'Printer.FontName = "바탕체"
    
    '출력전 결재버튼 비활성화, 도장칸으로 변경
    sss = "select nvl(apl_1yn,'N') as apl_1yn,nvl(apl_2yn,'N') as apl_2yn,nvl(apl_3yn,'N') as apl_3yn,nvl(apl_4yn,'N') as apl_4yn,nvl(apl_m1yn,'N') as apl_m1yn,"
    sss = sss & " nvl(apl_m1yn,'N') as apl_m1yn,nvl(apl_m2yn,'N') as apl_m2yn,nvl(apl_m3yn,'N') as apl_m3yn,nvl(apl_m4yn,'N') as apl_m4yn"
    sss = sss & " from oth_applist"
    sss = sss & " where apl_table = 'man_tooltesthd'"
    sss = sss & "   and apl_tdat = " & Val(txt_dat1)
    sss = sss & "   and apl_tseq = " & Val(txt_seq1)
    
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Rs.RecordCount > 0 Then
        
        If Rs!apl_1yn = "N" Then
            If pcnt = 2 Then
                sprd_print1.Col = 26
                sprd_print1.row = 2
                sprd_print1.CellType = CellTypeEdit
                sprd_print1.BackColor = RGB(255, 255, 255)
                sprd_print1.AddCellSpan 26, 2, 1, 1
                sprd_print1.AddCellSpan 26, 2, 4, 4
            ElseIf pcnt = 3 Then
                sprd_print2.Col = 26
                sprd_print2.row = 2
                sprd_print2.CellType = CellTypeEdit
                sprd_print2.BackColor = RGB(255, 255, 255)
                sprd_print2.AddCellSpan 26, 2, 1, 1
                sprd_print2.AddCellSpan 26, 2, 4, 4
            End If
        End If
        
        If Rs!apl_2yn = "N" Then
            If pcnt = 2 Then
                sprd_print1.Col = 30
                sprd_print1.row = 2
                sprd_print1.CellType = CellTypeEdit
                sprd_print1.BackColor = RGB(255, 255, 255)
                sprd_print1.AddCellSpan 30, 2, 1, 1
                sprd_print1.AddCellSpan 30, 2, 4, 4
            ElseIf pcnt = 3 Then
                sprd_print2.Col = 30
                sprd_print2.row = 2
                sprd_print2.CellType = CellTypeEdit
                sprd_print2.BackColor = RGB(255, 255, 255)
                sprd_print2.AddCellSpan 30, 2, 1, 1
                sprd_print2.AddCellSpan 30, 2, 4, 4
            End If
        End If
        
        If Rs!apl_3yn = "N" Then
            If pcnt = 2 Then
                sprd_print1.Col = 34
                sprd_print1.row = 2
                sprd_print1.CellType = CellTypeEdit
                sprd_print1.BackColor = RGB(255, 255, 255)
                sprd_print1.AddCellSpan 34, 2, 1, 1
                sprd_print1.AddCellSpan 34, 2, 4, 4
            ElseIf pcnt = 3 Then
                sprd_print2.Col = 34
                sprd_print2.row = 2
                sprd_print2.CellType = CellTypeEdit
                sprd_print2.BackColor = RGB(255, 255, 255)
                sprd_print2.AddCellSpan 34, 2, 1, 1
                sprd_print2.AddCellSpan 34, 2, 4, 4
            End If
        End If
        
        If Rs!apl_4yn = "N" Then
            If pcnt = 2 Then
                sprd_print1.Col = 38
                sprd_print1.row = 2
                sprd_print1.CellType = CellTypeEdit
                sprd_print1.BackColor = RGB(255, 255, 255)
                sprd_print1.AddCellSpan 38, 2, 1, 1
                sprd_print1.AddCellSpan 38, 2, 4, 4
            ElseIf pcnt = 3 Then
                sprd_print2.Col = 38
                sprd_print2.row = 2
                sprd_print2.CellType = CellTypeEdit
                sprd_print2.BackColor = RGB(255, 255, 255)
                sprd_print2.AddCellSpan 38, 2, 1, 1
                sprd_print2.AddCellSpan 38, 2, 4, 4
            End If
        End If
        
        If Rs!apl_m1yn = "N" Then
            If pcnt = 2 Then
                sprd_print1.Col = 34
                sprd_print1.row = 33
                sprd_print1.CellType = CellTypeEdit
                sprd_print1.BackColor = RGB(255, 255, 255)
                sprd_print1.AddCellSpan 34, 33, 4, 3
            Else
                sprd_print2.Col = 34
                sprd_print2.row = 33
                sprd_print2.CellType = CellTypeEdit
                sprd_print2.BackColor = RGB(255, 255, 255)
                sprd_print2.AddCellSpan 34, 33, 4, 3
            End If
        End If
        
        If Rs!apl_m2yn = "N" Then
            If pcnt = 2 Then
                sprd_print1.Col = 38
                sprd_print1.row = 33
                sprd_print1.CellType = CellTypeEdit
                sprd_print1.BackColor = RGB(255, 255, 255)
                sprd_print1.AddCellSpan 38, 33, 4, 3
            Else
                sprd_print2.Col = 38
                sprd_print2.row = 33
                sprd_print2.CellType = CellTypeEdit
                sprd_print2.BackColor = RGB(255, 255, 255)
                sprd_print2.AddCellSpan 38, 33, 4, 3
            End If
        End If
        
        If Rs!apl_m3yn = "N" Then
            If pcnt = 2 Then
                sprd_print1.Col = 34
                sprd_print1.row = 37
                sprd_print1.CellType = CellTypeEdit
                sprd_print1.BackColor = RGB(255, 255, 255)
                sprd_print1.AddCellSpan 34, 37, 4, 4
            Else
                sprd_print2.Col = 34
                sprd_print2.row = 37
                sprd_print2.CellType = CellTypeEdit
                sprd_print2.BackColor = RGB(255, 255, 255)
                sprd_print2.AddCellSpan 34, 37, 4, 4
            End If
        End If
        
        If Rs!apl_m4yn = "N" Then
            If pcnt = 2 Then
                sprd_print1.Col = 38
                sprd_print1.row = 37
                sprd_print1.CellType = CellTypeEdit
                sprd_print1.BackColor = RGB(255, 255, 255)
                sprd_print1.AddCellSpan 38, 37, 4, 4
            Else
                sprd_print2.Col = 38
                sprd_print2.row = 37
                sprd_print2.CellType = CellTypeEdit
                sprd_print2.BackColor = RGB(255, 255, 255)
                sprd_print2.AddCellSpan 38, 37, 4, 4
            End If
        End If
        
    End If
    
    If sprd_print1.Visible = True Then
        
        Comm1.CancelError = True
        Comm1.Action = 5
        
        sprd_print1.PrintOrientation = PrintOrientationPortrait
        
        sprd_print1.PrintBorder = True
        sprd_print1.PrintShadows = False
        sprd_print1.PrintMarginLeft = 800
        sprd_print1.PrintMarginRight = 600
        sprd_print1.PrintMarginTop = 300
        
        '2장씩 나올때================
'        sprd_print1.PrintShadows = False
'        sprd_print1.PrintMarginTop = 300
'        sprd_print1.PrintMarginBottom = 700
'        sprd_print1.PrintMarginLeft = 850
'        sprd_print1.PrintMarginRight = 650
        '============================
        
        sprd_print1.PrintHeader = "/n/n/fz""22""/l" & "/C/fb1" & "TOOL TEST DATA" & "/n/n/n/n/n"
        sprd_print1.PrintFooter = "/fn""Times New Roman""/l" & HLFNO & "/c" & "HY-LOK CORPORATION" & "/r" & "A4(210mmx297mm)         "
        
        sprd_print1.Action = ActionSmartPrint
        DoEvents
    
    ElseIf sprd_print2.Visible = True Then
        
        Comm1.CancelError = True
        Comm1.Action = 5
        
        sprd_print2.PrintOrientation = PrintOrientationPortrait
        
        sprd_print2.PrintBorder = True
        sprd_print2.PrintShadows = False
        sprd_print2.PrintMarginLeft = 800
        sprd_print2.PrintMarginRight = 600
        sprd_print2.PrintMarginTop = 500
        
        '2장씩 나올때================
'        sprd_print2.PrintShadows = False
'        sprd_print2.PrintMarginTop = 300
'        sprd_print2.PrintMarginBottom = 700
'        sprd_print2.PrintMarginLeft = 850
'        sprd_print2.PrintMarginRight = 650
        '============================
        
        sprd_print2.PrintHeader = "/n/n/fz""22""/l" & "/C/fb1" & "TOOL TEST DATA" & "/n/n/n/n/n"
        sprd_print2.PrintFooter = "/fn""Times New Roman""/l" & HLFNO & "/c" & "HY-LOK CORPORATION" & "/r" & "A4(210mmx297mm)         "
        
        sprd_print2.Action = ActionSmartPrint
        DoEvents
        
    ElseIf sprd_print3.Visible = True Then
      
        Comm1.CancelError = True
        Comm1.Action = 5
        
        sprd_print3.PrintOrientation = PrintOrientationPortrait
        
        sprd_print3.PrintShadows = False
        sprd_print3.PrintMarginTop = 300
        sprd_print3.PrintMarginBottom = 700
        sprd_print3.PrintMarginLeft = 850
        sprd_print3.PrintMarginRight = 650
        
        sprd_print3.PrintHeader = "/n/n/fz""22""/l" & "/C/fb1" & "TOOL TEST DATA" & "/n/n/n/n/n"
        sprd_print3.PrintFooter = "/fn""Times New Roman""/l" & HLFNO & "/c" & "HY-LOK CORPORATION" & "/r" & "A4(210mmx297mm)         "
        
        sprd_print3.Action = ActionSmartPrint
        DoEvents
    
    End If
    
    '변경된 결재버튼 복구
    Call app_init
    
    Rs.Close
    
Exit Sub

err_chk:

    Call app_init

    Printer.KillDoc
    msg = Err.Description
    
End Sub

Sub app_init()
    
    If pcnt = 2 Then
        
        With sprd_print1
            
            If Rs!apl_1yn = "N" Then
                .AddCellSpan 26, 2, 4, 2
                .AddCellSpan 26, 4, 4, 2
                .row = 2: .Col = 26: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""
                
                If asab(1) <> 0 Then
                    .TypeButtonText = aname(1)
                Else
                    .TypeButtonText = "Click"
                End If
                
            End If
             
            If Rs!apl_2yn = "N" Then
                .AddCellSpan 30, 2, 4, 2
                .AddCellSpan 30, 4, 4, 2
                .row = 2: .Col = 30: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""
                
                If asab(2) <> 0 Then
                    .TypeButtonText = aname(2)
                Else
                    .TypeButtonText = "Click"
                End If
                
            End If
            
            If Rs!apl_3yn = "N" Then
                .AddCellSpan 34, 2, 4, 2
                .AddCellSpan 34, 4, 4, 2
                .row = 2: .Col = 34: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""
                
                If asab(3) <> 0 Then
                    .TypeButtonText = aname(3)
                Else
                    .TypeButtonText = "Click"
                End If
                
            End If
            
            If Rs!apl_4yn = "N" Then
                .AddCellSpan 38, 2, 4, 2
                .AddCellSpan 38, 4, 4, 2
                .row = 2: .Col = 38: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""
                
                If asab(4) <> 0 Then
                    .TypeButtonText = aname(4)
                Else
                    .TypeButtonText = "Click"
                End If
                
            End If
            
            If Rs!apl_m1yn = "N" Then
                .AddCellSpan 34, 33, 4, 2
                .AddCellSpan 34, 35, 4, 1
                .row = 33: .Col = 34: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""
                
                If msab(1) <> 0 Then
                    .TypeButtonText = mname(1)
                Else
                    .TypeButtonText = "Click"
                End If
            End If
            
            If Rs!apl_m2yn = "N" Then
                .AddCellSpan 38, 33, 4, 2
                .AddCellSpan 38, 35, 4, 1
                .row = 33: .Col = 38: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""
                
                If msab(2) <> 0 Then
                    .TypeButtonText = mname(2)
                Else
                    .TypeButtonText = "Click"
                End If
            End If
            
            If Rs!apl_m3yn = "N" Then
                .AddCellSpan 34, 37, 4, 2
                .AddCellSpan 34, 39, 4, 2
                .row = 37: .Col = 34: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""
                
                If msab(3) <> 0 Then
                    .TypeButtonText = mname(3)
                Else
                    .TypeButtonText = "Click"
                End If
                
            End If
            
            If Rs!apl_m4yn = "N" Then
                .AddCellSpan 38, 37, 4, 2
                .AddCellSpan 38, 39, 4, 2
                .row = 37: .Col = 38: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""
                
                If msab(4) <> 0 Then
                    .TypeButtonText = mname(4)
                Else
                    .TypeButtonText = "Click"
                End If
                
            End If
            
        End With
        
    ElseIf pcnt = 3 Then
        
        With sprd_print2

            If Rs!apl_1yn = "N" Then
                .AddCellSpan 26, 2, 4, 2
                .AddCellSpan 26, 4, 4, 2
                .row = 2: .Col = 26: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""

                If asab(1) <> 0 Then
                    .TypeButtonText = aname(1)
                Else
                    .TypeButtonText = "Click"
                End If

            End If

            If Rs!apl_2yn = "N" Then
                .AddCellSpan 30, 2, 4, 2
                .AddCellSpan 30, 4, 4, 2
                .row = 2: .Col = 30: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""
                
                If asab(2) <> 0 Then
                    .TypeButtonText = aname(2)
                Else
                    .TypeButtonText = "Click"
                End If
                
            End If
            
            If Rs!apl_3yn = "N" Then
                .AddCellSpan 34, 2, 4, 2
                .AddCellSpan 34, 4, 4, 2
                .row = 2: .Col = 34: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""
                
                If asab(3) <> 0 Then
                    .TypeButtonText = aname(3)
                Else
                    .TypeButtonText = "Click"
                End If
                
            End If

            If Rs!apl_4yn = "N" Then
                .AddCellSpan 38, 2, 4, 2
                .AddCellSpan 38, 4, 4, 2
                .row = 2: .Col = 38: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""

                If asab(4) <> 0 Then
                    .TypeButtonText = aname(4)
                Else
                    .TypeButtonText = "Click"
                End If

            End If
            
            If Rs!apl_m1yn = "N" Then
                .AddCellSpan 34, 33, 4, 2
                .AddCellSpan 34, 35, 4, 1
                .row = 33: .Col = 34: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""
                
                If msab(1) <> 0 Then
                    .TypeButtonText = mname(1)
                Else
                    .TypeButtonText = "Click"
                End If
            End If
            
            If Rs!apl_m2yn = "N" Then
                .AddCellSpan 38, 33, 4, 2
                .AddCellSpan 38, 35, 4, 1
                .row = 33: .Col = 38: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""
                
                If msab(2) <> 0 Then
                    .TypeButtonText = mname(2)
                Else
                    .TypeButtonText = "Click"
                End If
            End If
            
            If Rs!apl_m3yn = "N" Then
                .AddCellSpan 34, 37, 4, 2
                .AddCellSpan 34, 39, 4, 2
                .row = 37: .Col = 34: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""
                
                If msab(3) <> 0 Then
                    .TypeButtonText = mname(3)
                Else
                    .TypeButtonText = "Click"
                End If
                
            End If
            
            If Rs!apl_m4yn = "N" Then
                .AddCellSpan 38, 37, 4, 2
                .AddCellSpan 38, 39, 4, 2
                .row = 37: .Col = 38: .CellType = CellTypeButton: .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .Text = ""
                
                If msab(4) <> 0 Then
                    .TypeButtonText = mname(4)
                Else
                    .TypeButtonText = "Click"
                End If
                
            End If

        End With
        
    End If
        
End Sub

Public Sub btn_view_Click()

    If IsNumeric(txt_dat1) = False Or Len(txt_dat1) <> 8 Then
        Call msg_display("등록 일자를 확인하세요!")
        txt_dat1.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txt_seq1) = False Or Len(txt_seq1) < 1 Then
        Call msg_display("등록 순번을 확인하세요!")
        txt_seq1.SetFocus
        Exit Sub
    End If
    
    sss = "       select count(*) cnt"
    sss = sss & "   from man_tooltesthd, man_tooltestds"
    sss = sss & "  where tth_dat = ttd_dat"
    sss = sss & "    and tth_seq = ttd_seq"
    sss = sss & "    and tth_dat = " & txt_dat1
    sss = sss & "    and tth_seq = " & txt_seq1
    sss = sss & "  order by ttd_lno"
    
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Rs.RecordCount < 1 Then
        Rs.Close
        Call msg_display("등록된 내역이 없습니다!")
        Exit Sub
    End If
    
    pcnt = Rs!cnt
    
    If Rs!cnt = 2 Then
        sprd_print1.Visible = True
        sprd_print2.Visible = False
        sprd_print3.Visible = False
        Call print1
    ElseIf Rs!cnt = 3 Then
        sprd_print1.Visible = False
        sprd_print2.Visible = True
        sprd_print3.Visible = False
        Call print2
    End If

End Sub

Private Sub print1()
    
    Dim purpose As String
    Dim resultOK As String
    Dim resultNG As String
    
    
    sss = "       select *"
    sss = sss & "   from man_tooltesthd, man_tooltestds"
    sss = sss & "  where tth_dat = ttd_dat"
    sss = sss & "    and tth_seq = ttd_seq"
    sss = sss & "    and tth_dat = " & txt_dat1
    sss = sss & "    and tth_seq = " & txt_seq1
    sss = sss & "  order by ttd_lno"
    
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Rs.RecordCount < 1 Then
        Rs.Close
        Call msg_display("등록된 내역이 없습니다!")
        Exit Sub
    End If
    
    Call clear1
    Call app
    
    If Not IsNull(Rs!tth_testno) Then sprd_print1.row = 1: sprd_print1.Col = 6: sprd_print1.Text = Rs!tth_testno
    If Not IsNull(Rs!tth_tdat) Then sprd_print1.row = 1: sprd_print1.Col = 16: sprd_print1.Text = Format(Rs!tth_tdat, "0000-00-00")
    If Not IsNull(Rs!tth_jubno) Then sprd_print1.row = 2: sprd_print1.Col = 6: sprd_print1.Text = Format(Rs!tth_jubno, "00000000-000")

    If Not IsNull(Rs!tth_title) Then sprd_print1.row = 4: sprd_print1.Col = 1: sprd_print1.Text = Rs!tth_title
    
    If Not IsNull(Rs!tth_tmcd) Then
        sprd_print1.row = 8: sprd_print1.Col = 7: sprd_print1.Text = Rs!tth_tmcd
    
        sss = "       select mhc_name, ems_mark, sok_name(mhc_sok) soknm "
        sss = sss & "   from man_machcd, eam_mast"
        sss = sss & "  where mhc_code = '" & Rs!tth_tmcd & "'"
        sss = sss & "    and mhc_code = ems_mcd(+)"
                
        Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        If Ks.RecordCount > 0 Then
            If Not IsNull(Ks!mhc_name) Then sprd_print1.row = 7: sprd_print1.Col = 7: sprd_print1.Text = Ks!mhc_name
            If Not IsNull(Ks!ems_mark) Then sprd_print1.row = 6: sprd_print1.Col = 7: sprd_print1.Text = Ks!ems_mark
            If Not IsNull(Ks!soknm) Then sprd_print1.row = 9: sprd_print1.Col = 7: sprd_print1.Text = Ks!soknm
        End If
        
        Ks.Close
    
    End If

    If Not IsNull(Rs!tth_tlot) Then
        
        sss = "       select dit_bpcd, dit_bpjil, dit_jacd, dit_jajil"
        sss = sss & "   from man_direct"
        sss = sss & "  where dit_lot = '" & Rs!tth_tlot & "'"
            
        Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        If Ks.RecordCount > 0 Then
            
            If Not IsNull(Ks!dit_bpcd) Then sprd_print1.row = 10: sprd_print1.Col = 7: sprd_print1.Text = Ks!dit_bpcd
            If Not IsNull(Ks!dit_bpjil) Then sprd_print1.row = 11: sprd_print1.Col = 7: sprd_print1.Text = Ks!dit_bpjil
            If Not IsNull(Ks!dit_jacd) Then sprd_print1.row = 12: sprd_print1.Col = 7: sprd_print1.Text = Ks!dit_jacd
            If Not IsNull(Ks!dit_jajil) Then sprd_print1.row = 13: sprd_print1.Col = 7: sprd_print1.Text = Ks!dit_jajil
        
        End If
    
        Ks.Close
    
    End If
    
    purpose = "        TEST 목적" & Chr(13) & Chr(13)
    
    If Rs!tth_pyn1 = "Y" Then
        purpose = purpose & "        1.공구 수명 ( ○ )"
    Else
        purpose = purpose & "        1.공구 수명 (     )"
    End If

    If Rs!tth_pyn2 = "Y" Then
        purpose = purpose & "             2.칩 처리 ( ○ )" & Chr(13)
    Else
        purpose = purpose & "             2.칩 처리 (     )" & Chr(13)
    End If


    If Rs!tth_pyn3 = "Y" Then
        purpose = purpose & "        3.시간 단축 ( ○ )"
    Else
        purpose = purpose & "        3.시간 단축 (     )"
    End If


    If Rs!tth_pyn4 = "Y" Then
        purpose = purpose & "             4.공구비 절감 ( ○ )" & Chr(13)
    Else
        purpose = purpose & "             4.공구비 절감 (     )" & Chr(13)
    End If
     
    
    If Rs!tth_pyn5 = "Y" Then
        purpose = purpose & "        5.기타 ( ○ )"
    Else
        purpose = purpose & "        5.기타 (     )"
    End If
    
    sprd_print1.row = 14: sprd_print1.Col = 1: sprd_print1.Text = purpose

    If Not IsNull(Rs!tth_tsab) Then
                                            
        sss = "       select sin_name, sin_sok, sok_name(sin_sok) soknm"
        sss = sss & "   from peo_sinbun"
        sss = sss & "  where sin_sab = " & Rs!tth_tsab
        'sss = sss & "    and sin_taedt = 0 "
            
        Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        If Ks.RecordCount > 0 Then
            sprd_print1.row = 17: sprd_print1.Col = 26: sprd_print1.Text = Ks!sin_name & "(" & Ks!soknm & ")"
        End If
        Ks.Close
        
    End If
    
    '도면 이미지
    If Not IsNull(Rs!tth_file1) Then
        If Right(Rs!tth_file1, 3) = "JPG" Then
           '
           '==============================
           ' FTP를 이용한 파일 다운로드
           '==============================
           If FTP_Connection Then
              '
              If Not FTP경로체크(SERVER_PATH) Then
                 Call FTP_DisConnect
                 MsgBox "서버경로를 찾을수 없습니다.(정보관리센터 문의)"
                 Exit Sub
              End If
              '
              GPath = "c:\jpeg\" & Rs!tth_file1
              '
              If FTP_Download(SERVER_PATH & Rs!tth_file1, GPath) Then
                 sprd_print1.row = 6: sprd_print1.Col = 21: sprd_print1.TypePictPicture = LoadPicture(GPath)
                 sprd_print1.TypePictStretch = True
                 sprd_print1.TypePictMaintainScale = True
              End If
              
              '
              Call FTP_DisConnect
              '
           End If
           '
           '======================================
          ' exists = ExistFile(N_Driver & ":\" & Rs!tth_file1)
          '  If exists = True Then
          '      FileCopy N_Driver & ":\" & Rs!tth_file1, "c:\temp\" & "Temp.jpg"
          '      GPath = "c:\temp\Temp.jpg"
          '      sprd_print1.Row = 6: sprd_print1.Col = 21: sprd_print1.TypePictPicture = LoadPicture(GPath)
          '
          '  End If
        End If
    End If
    
    If Not IsNull(Rs!tth_rmk) Then sprd_print1.row = 19: sprd_print1.Col = 34: sprd_print1.Text = Rs!tth_rmk
    If Not IsNull(Rs!tth_cmt) Then sprd_print1.row = 28: sprd_print1.Col = 34: sprd_print1.Text = Rs!tth_cmt
    
    'DESC
    Do While Not Rs.EOF
                 
        '기존공구
        If Rs!ttd_lno = 1 Then
            sprd_print1.Col = 7
            If Not IsNull(Rs!ttd_maker) Then sprd_print1.row = 18:  sprd_print1.Text = Rs!ttd_maker & "(기존)"
            If Not IsNull(Rs!ttd_tipstd) Then sprd_print1.row = 19: sprd_print1.Text = Rs!ttd_tipstd
            If Not IsNull(Rs!ttd_tipjil) Then sprd_print1.row = 20: sprd_print1.Text = Rs!ttd_tipjil
            If Not IsNull(Rs!ttd_holder) Then sprd_print1.row = 21: sprd_print1.Text = Rs!ttd_holder
            If Not IsNull(Rs!ttd_rcntmn) Then sprd_print1.row = 22: sprd_print1.Text = Rs!ttd_rcntmn
            If Not IsNull(Rs!ttd_rcntmx) Then sprd_print1.row = 22: sprd_print1.Text = sprd_print1.Text & "-" & Rs!ttd_rcntmx
            If Not IsNull(Rs!ttd_movmn) Then sprd_print1.row = 24:  sprd_print1.Text = Rs!ttd_movmn
            If Not IsNull(Rs!ttd_movmx) Then sprd_print1.row = 24:  sprd_print1.Text = sprd_print1.Text & "-" & Rs!ttd_movmx
            If Not IsNull(Rs!ttd_depth) Then sprd_print1.row = 23:  sprd_print1.Text = Rs!ttd_depth & "mm"
            If Not IsNull(Rs!ttd_tct) Then sprd_print1.row = 25:    sprd_print1.Text = Format(Rs!ttd_tct, "###,##0")
            If Not IsNull(Rs!ttd_pct) Then sprd_print1.row = 25:    sprd_print1.Col = 14: sprd_print1.Text = Format(Rs!ttd_pct, "###,##0")
                                           
            If Not IsNull(Rs!ttd_fluid) Then
                sprd_print1.row = 26: sprd_print1.Col = 7:
                If Rs!ttd_fluid = 1 Then
                    sprd_print1.Text = "수용성"
                Else
                    sprd_print1.Text = "비수용성"
                End If
            End If
                                        
            sprd_print1.Col = 7
            If Not IsNull(Rs!ttd_qty) Then sprd_print1.row = 27: sprd_print1.Text = Format(Rs!ttd_qty, "###,##0") & " Point"
            
            If Not IsNull(Rs!ttd_dan) Then
                If InStr(1, Rs!ttd_dan, ".", 0) <> 0 Then
                    sprd_print1.row = 28: sprd_print1.Text = Format(Rs!ttd_dan, "###,###,##0.0#") & "원"
                Else
                    sprd_print1.row = 28: sprd_print1.Text = Format(Rs!ttd_dan, "###,###,##0") & "원"
                    
                End If
            End If
            If Not IsNull(Rs!ttd_tldn) Then
                If InStr(1, Rs!ttd_tldn, ".", 0) <> 0 Then
                    sprd_print1.row = 29: sprd_print1.Text = Format(Rs!ttd_tldn, "###,###,##0.0#") & "원/Corner"
                Else
                    sprd_print1.row = 29: sprd_print1.Text = Format(Rs!ttd_tldn, "###,###,##0") & "원/Corner"
                End If
            End If
            If Not IsNull(Rs!ttd_chdn) Then
                If InStr(1, Rs!ttd_chdn, ".", 0) <> 0 Then
                    sprd_print1.row = 30: sprd_print1.Text = Format(Rs!ttd_chdn, "###,###,##0.0#") & "원"
                Else
                    sprd_print1.row = 30: sprd_print1.Text = Format(Rs!ttd_chdn, "###,###,##0") & "원/EA"
                End If
            End If
        
        End If
        
        '테스트1
        If Rs!ttd_lno = 2 Then
        
            sprd_print1.Col = 21
            If Not IsNull(Rs!ttd_maker) Then sprd_print1.row = 18:  sprd_print1.Text = Rs!ttd_maker & "(테스트)"
            If Not IsNull(Rs!ttd_tipstd) Then sprd_print1.row = 19: sprd_print1.Text = Rs!ttd_tipstd
            If Not IsNull(Rs!ttd_tipjil) Then sprd_print1.row = 20: sprd_print1.Text = Rs!ttd_tipjil
            If Not IsNull(Rs!ttd_holder) Then sprd_print1.row = 21: sprd_print1.Text = Rs!ttd_holder
            If Not IsNull(Rs!ttd_rcntmn) Then sprd_print1.row = 22: sprd_print1.Text = Rs!ttd_rcntmn
            If Not IsNull(Rs!ttd_rcntmx) Then sprd_print1.row = 22: sprd_print1.Text = sprd_print1.Text & "-" & Rs!ttd_rcntmx
            If Not IsNull(Rs!ttd_movmn) Then sprd_print1.row = 24:  sprd_print1.Text = Rs!ttd_movmn
            If Not IsNull(Rs!ttd_movmx) Then sprd_print1.row = 24:  sprd_print1.Text = sprd_print1.Text & "-" & Rs!ttd_movmx
            If Not IsNull(Rs!ttd_depth) Then sprd_print1.row = 23:  sprd_print1.Text = Rs!ttd_depth & "mm"
            If Not IsNull(Rs!ttd_tct) Then sprd_print1.row = 25:    sprd_print1.Text = Format(Rs!ttd_tct, "###,##0")
            If Not IsNull(Rs!ttd_pct) Then sprd_print1.row = 25:    sprd_print1.Col = 27: sprd_print1.Text = Format(Rs!ttd_pct, "###,##0")
            
            If Not IsNull(Rs!ttd_fluid) Then
                sprd_print1.row = 26: sprd_print1.Col = 21:
                If Rs!ttd_fluid = 1 Then
                    sprd_print1.Text = "수용성"
                Else
                    sprd_print1.Text = "비수용성"
                End If
            End If
            
            sprd_print1.Col = 21
            If Not IsNull(Rs!ttd_qty) Then sprd_print1.row = 27: sprd_print1.Text = Format(Rs!ttd_qty, "###,##0") & " Point"
            
            If Not IsNull(Rs!ttd_dan) Then
                If InStr(1, Rs!ttd_dan, ".", 0) <> 0 Then
                    sprd_print1.row = 28: sprd_print1.Text = Format(Rs!ttd_dan, "###,###,##0.0#") & "원"
                Else
                    sprd_print1.row = 28: sprd_print1.Text = Format(Rs!ttd_dan, "###,###,##0") & "원"
                    
                End If
            End If
            If Not IsNull(Rs!ttd_tldn) Then
                If InStr(1, Rs!ttd_tldn, ".", 0) <> 0 Then
                    sprd_print1.row = 29: sprd_print1.Text = Format(Rs!ttd_tldn, "###,###,##0.0#") & "원/Corner"
                Else
                    sprd_print1.row = 29: sprd_print1.Text = Format(Rs!ttd_tldn, "###,###,##0") & "원/Corner"
                End If
            End If
            If Not IsNull(Rs!ttd_chdn) Then
                If InStr(1, Rs!ttd_chdn, ".", 0) <> 0 Then
                    sprd_print1.row = 30: sprd_print1.Text = Format(Rs!ttd_chdn, "###,###,##0.0#") & "원"
                Else
                    sprd_print1.row = 30: sprd_print1.Text = Format(Rs!ttd_chdn, "###,###,##0") & "원/EA"
                End If
            End If
            
        End If
        
        '테스트2
        If Rs!ttd_lno = 4 Then
            
            sprd_print1.Col = 21
            If Not IsNull(Rs!ttd_maker) Then sprd_print1.row = 18: sprd_print1.Text = Rs!ttd_maker
            If Not IsNull(Rs!ttd_tipstd) Then sprd_print1.row = 19: sprd_print1.Text = Rs!ttd_tipstd
            If Not IsNull(Rs!ttd_tipjil) Then sprd_print1.row = 20: sprd_print1.Text = Rs!ttd_tipjil
            If Not IsNull(Rs!ttd_holder) Then sprd_print1.row = 21: sprd_print1.Text = Rs!ttd_holder
            If Not IsNull(Rs!ttd_rcntmn) Then sprd_print1.row = 22: sprd_print1.Text = Rs!ttd_rcntmn
            If Not IsNull(Rs!ttd_rcntmx) Then sprd_print1.row = 22: sprd_print1.Text = sprd_print1.Text & "-" & Rs!ttd_rcntmx
            If Not IsNull(Rs!ttd_movmn) Then sprd_print1.row = 24: sprd_print1.Text = Rs!ttd_movmn
            If Not IsNull(Rs!ttd_movmx) Then sprd_print1.row = 24: sprd_print1.Text = sprd_print1.Text & "-" & Rs!ttd_movmx
            If Not IsNull(Rs!ttd_depth) Then sprd_print1.row = 23: sprd_print1.Text = Rs!ttd_depth & "mm"
            If Not IsNull(Rs!ttd_tct) Then sprd_print1.row = 25: sprd_print1.Text = Format(Rs!ttd_tct, "###,##0")
            If Not IsNull(Rs!ttd_pct) Then sprd_print1.row = 25: sprd_print1.Col = 28: sprd_print1.Text = Format(Rs!ttd_pct, "###,##0")
            
            If Not IsNull(Rs!ttd_fluid) Then
                sprd_print1.row = 26: sprd_print1.Col = 21:
                If Rs!ttd_fluid = 1 Then
                    sprd_print1.Text = "수용성"
                Else
                    sprd_print1.Text = "비수용성"
                End If
            End If
            
            sprd_print1.Col = 7
            If Not IsNull(Rs!ttd_qty) Then sprd_print1.row = 27: sprd_print1.Text = Format(Rs!ttd_qty, "###,##0") & " Point"
            
            If Not IsNull(Rs!ttd_dan) Then
                If InStr(1, Rs!ttd_dan, ".", 0) <> 0 Then
                    sprd_print1.row = 28: sprd_print1.Text = Format(Rs!ttd_dan, "###,###,##0.0#") & "원"
                Else
                    sprd_print1.row = 28: sprd_print1.Text = Format(Rs!ttd_dan, "###,###,##0") & "원"
                    
                End If
            End If
            If Not IsNull(Rs!ttd_tldn) Then
                If InStr(1, Rs!ttd_tldn, ".", 0) <> 0 Then
                    sprd_print1.row = 29: sprd_print1.Text = Format(Rs!ttd_tldn, "###,###,##0.0#") & "원/Corner"
                Else
                    sprd_print1.row = 29: sprd_print1.Text = Format(Rs!ttd_tldn, "###,###,##0") & "원/Corner"
                End If
            End If
            If Not IsNull(Rs!ttd_chdn) Then
                If InStr(1, Rs!ttd_chdn, ".", 0) <> 0 Then
                    sprd_print1.row = 30: sprd_print1.Text = Format(Rs!ttd_chdn, "###,###,##0.0#") & "원"
                Else
                    sprd_print1.row = 30: sprd_print1.Text = Format(Rs!ttd_chdn, "###,###,##0") & "원/EA"
                End If
            End If
            
        End If
        
        If Rs!ttd_ryn = "Y" Then
               
            resultOK = ""
            resultNG = ""
            
            If Not IsNull(Rs!ttd_result) Then
                
                If Rs!ttd_result = "OK" Then
                    sprd_print1.row = 31: sprd_print1.Col = 7
                    sprd_print1.Text = "  O.K ( ○ )": 'sprd_print1.FontBold = True
                    sprd_print1.row = 36: sprd_print1.Col = 7
                    sprd_print1.Text = "  N.G (     )"
                End If
                If Rs!ttd_result = "NG" Then
                    sprd_print1.row = 31: sprd_print1.Col = 7
                    sprd_print1.Text = "  O.K (     )"
                    sprd_print1.row = 36: sprd_print1.Col = 7
                    sprd_print1.Text = "  N.G ( ○ )": 'sprd_print1.FontBold = True
                End If
            End If
                
            
            If Rs!ttd_result = "OK" Then
                
                If Rs!ttd_ryn1 = "Y" Then
                    resultOK = resultOK & "           1.공구수명 연장 ( ○ )"
                Else
                    resultOK = resultOK & "           1.공구수명 연장 (     )"
                End If
    
                If Rs!ttd_ryn4 = "Y" Then
                    resultOK = resultOK & "          4.시간 단축 ( ○ )" & Chr(13)
                Else
                    resultOK = resultOK & "          4.시간 단축 (     )" & Chr(13)
                End If
                    
                If Rs!ttd_ryn2 = "Y" Then
                    resultOK = resultOK & "           2.칩처리 양호 ( ○ )"
                Else
                    resultOK = resultOK & "           2.칩처리 양호 (     )"
                End If
                    
                If Rs!ttd_ryn5 = "Y" Then
                    resultOK = resultOK & "             5.기타 ( ○ )" & Chr(13)
                Else
                    resultOK = resultOK & "             5.기타 (     )" & Chr(13)
                End If
                    
                If Rs!ttd_ryn3 = "Y" Then
                    resultOK = resultOK & "           3.공구비 절감 ( ○ )"
                Else
                    resultOK = resultOK & "           3.공구비 절감 (     )"
                End If
                    
                resultNG = resultNG & "           1.결손 (     )"
                resultNG = resultNG & "                      4.공구비 상승 (     )" & Chr(13)
                resultNG = resultNG & "           2.마모 (     )"
                resultNG = resultNG & "                      5.기타 (     )" & Chr(13)
                resultNG = resultNG & "           3.칩처리 불량 (     )"
                
                
            Else
                
                resultNG = ""
                
                If Rs!ttd_ryn1 = "Y" Then
                    resultNG = resultNG & "           1.결손 ( ○ )"
                Else
                    resultNG = resultNG & "           1.결손 (     )"
                End If
    
                If Rs!ttd_ryn4 = "Y" Then
                    resultNG = resultNG & "                      4.공구비 상승 ( ○ )" & Chr(13)
                Else
                    resultNG = resultNG & "                      4.공구비 상승 (     )" & Chr(13)
                End If
                    
                If Rs!ttd_ryn2 = "Y" Then
                    resultNG = resultNG & "           2.마모 ( ○ )"
                Else
                    resultNG = resultNG & "           2.마모 (     )"
                End If
                    
                If Rs!ttd_ryn5 = "Y" Then
                    resultNG = resultNG & "                      5.기타 ( ○ )" & Chr(13)
                Else
                    resultNG = resultNG & "                      5.기타 (     )" & Chr(13)
                End If
                    
                If Rs!ttd_ryn3 = "Y" Then
                    resultNG = resultNG & "           3.칩처리 불량 ( ○ )"
                Else
                    resultNG = resultNG & "           3.칩처리 불량 (     )"
                End If
                
                resultOK = resultOK & "           1.공구수명 연장 (     )"
                resultOK = resultOK & "          4.시간 단축 (     )" & Chr(13)
                resultOK = resultOK & "           2.칩처리 양호 (     )"
                resultOK = resultOK & "             5.기타 (     )" & Chr(13)
                resultOK = resultOK & "           3.공구비 절감 (     )"
            
            End If
                
            sprd_print1.row = 31: sprd_print1.Col = 14: sprd_print1.Text = resultOK
            sprd_print1.row = 36: sprd_print1.Col = 14: sprd_print1.Text = resultNG
            
            sprd_print1.row = 41: sprd_print1.Col = 7
            If Rs!ttd_rtyn = "Y" Then
                sprd_print1.Text = "재 TEST 가능 여부 - OK"
            Else
                sprd_print1.Text = "재 TEST 가능 여부 - NOT"
            End If
            
            If Rs!ttd_lno = 1 Then
                sprd_print1.Col = 7
            ElseIf Rs!ttd_lno = 2 Then
                sprd_print1.Col = 21
            End If
                
            sprd_print1.row = 18: sprd_print1.FontBold = True: sprd_print1.RowHeight(18) = 14
            sprd_print1.row = 19: sprd_print1.FontBold = True: sprd_print1.RowHeight(19) = 14
                                                                    
            sprd_print1.row = 28: sprd_print1.FontBold = True: sprd_print1.RowHeight(28) = 14
            sprd_print1.row = 29: sprd_print1.FontBold = True: sprd_print1.RowHeight(29) = 14
            sprd_print1.row = 30: sprd_print1.FontBold = True: sprd_print1.RowHeight(30) = 14
                    
        End If
            
        Rs.MoveNext
    Loop
    
    Rs.Close
    
    'btn_view1.Enabled = False
    
    'txt_dat1.Enabled = False
    'txt_seq1.Enabled = False
    
    'btn_add1.Enabled = False
    'btn_mod1.Enabled = True
    'btn_del1.Enabled = True
    
    Call msg_display("조회완료!")
    
    
End Sub

Private Sub app()
        
    sss = "    select apl_table,apl_tdat,apl_tseq,"
    sss = sss & "     apl_isab,sinbun_name(apl_isab) as iname,"
    sss = sss & "     apl_1sab,apl_1yn,apl_1dat,sinbun_name(apl_1sab) as appname1,"
    sss = sss & "     apl_2sab,apl_2yn,apl_2dat,sinbun_name(apl_2sab) as appname2,"
    sss = sss & "     apl_3sab,apl_3yn,apl_3dat,sinbun_name(apl_3sab) as appname3,"
    sss = sss & "     apl_4sab,apl_4yn,apl_4dat,sinbun_name(apl_4sab) as appname4,"
    sss = sss & "     apl_alldat,"
    sss = sss & "     apl_m1sab,apl_m1yn,apl_m1dat,sinbun_name(apl_m1sab) as appmname1,"
    sss = sss & "     apl_m2sab,apl_m2yn,apl_m2dat,sinbun_name(apl_m2sab) as appmname2,"
    sss = sss & "     apl_m3sab,apl_m3yn,apl_m3dat,sinbun_name(apl_m3sab) as appmname3,"
    sss = sss & "     apl_m4sab,apl_m4yn,apl_m4dat,sinbun_name(apl_m4sab) as appmname4"
    sss = sss & " from oth_applist"
    sss = sss & " where apl_table = 'man_tooltesthd'"
    sss = sss & "  and apl_tdat = " & txt_dat1
    sss = sss & "  and apl_tseq = " & txt_seq1
    
    Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Ks.RecordCount < 1 Then
        Ks.Close
        Call msg_display("결재 내역이 없습니다!")
        Exit Sub
    End If
    
    '전역으로 사용할 결재자 사번 및 결재자 이름 초기화====
    aname(1) = ""
    aname(2) = ""
    aname(3) = ""
    aname(4) = ""
    mname(1) = ""
    mname(2) = ""
    mname(3) = ""
    mname(4) = ""
    iname = ""
        
    asab(1) = 0
    asab(2) = 0
    asab(3) = 0
    asab(4) = 0
    msab(1) = 0
    msab(2) = 0
    msab(3) = 0
    msab(4) = 0
    isab = 0
    '======================================================
    
    If pcnt = 2 Then
    
        With sprd_print1

            
            '담당자
            If Ks!apl_isab <> 0 Then
                
                .row = 2
                .Col = 22
                .CellType = CellTypeEdit: .Text = Ks!iname: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                iname = Ks!iname: isab = Ks!apl_isab
                    
                .row = 4: .Text = Format(Val(Ks!apl_tdat), "####/##/##")
            
            End If
            
            '생산
            If Ks!apl_1sab <> 0 Then
            
                .row = 2: .Col = 26
                .TypeButtonText = Ks!appname1: .CellTag = Ks!apl_1sab
                aname(1) = Ks!appname1: asab(1) = Ks!apl_1sab
            
                If Ks!apl_1dat > 0 Then
                    
                    .CellType = CellTypeEdit: .Text = Ks!appname1: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                        
                    .row = 4: .Text = Format(Val(Ks!apl_1dat), "####/##/##")
                    
                End If
            
            End If
            
            '기술연구소
            If Ks!apl_2sab <> 0 Then
                
                .row = 2: .Col = 30
                .TypeButtonText = Ks!appname2: .CellTag = Ks!apl_2sab
                aname(2) = Ks!appname2: asab(2) = Ks!apl_2sab
                
                If Ks!apl_2dat > 0 Then
                    
                    .CellType = CellTypeEdit: .Text = Ks!appname2: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                        
                    .row = 4: .Text = Format(Val(Ks!apl_2dat), "####/##/##")
                    
                End If
                
            End If
            
            '전무님
            If Ks!apl_3sab <> 0 Then
                
                .row = 2: .Col = 34
                .TypeButtonText = Ks!appname3: .CellTag = Ks!apl_3sab
                aname(3) = Ks!appname3: asab(3) = Ks!apl_3sab
                
                If Ks!apl_3dat > 0 Then
                    
                    .CellType = CellTypeEdit: .Text = Ks!appname3: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                        
                    .row = 4: .Text = Format(Val(Ks!apl_3dat), "####/##/##")
                    
                End If
                
            End If
            
            '사장님
            If Ks!apl_4sab <> 0 Then
                
                .row = 2: .Col = 38
                .CellTag = Ks!apl_4sab
                
                If Ks!appname4 = "문창환" Then
                    .TypeButtonText = "사장님"
                    aname(4) = "사장님"
                Else
                    .TypeButtonText = Ks!appname4
                    aname(4) = Ks!appname4
                End If
                
                asab(4) = Ks!apl_4sab
                
                '승인결재가 대결이 아닌경우
                If Ks!apl_4dat > 0 And Ks!apl_4dat <> 99999999 Then
                    
                    .CellType = CellTypeEdit: .Text = "사장님": .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(255, 214, 255)
                    .row = 4: .Text = Format(Val(Ks!apl_4dat), "####/##/##")
                
                '승인결재가 대결인 경우
                ElseIf Ks!apl_4dat = 99999999 Then
                    
                     .CellType = CellTypeEdit: .Text = "대결": .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(255, 214, 255)
                     .row = 4: .Text = Format(Val(Ks!apl_alldat), "####/##/##")
                    
                End If
                
            End If
            
            '검토1
            If Ks!apl_m1sab <> 0 Then
                
                .row = 33: .Col = 34
                .TypeButtonText = Ks!appmname1: .CellTag = Ks!apl_m1sab
                mname(1) = Ks!appmname1: msab(1) = Ks!apl_m1sab
                
                If Ks!apl_m1dat > 0 Then
                    
                    .CellType = CellTypeEdit: .Text = Ks!appmname1: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                        
                    .row = 35: .Text = Format(Val(Ks!apl_m1dat), "####/##/##")
                    
                End If
                
            End If
            
            '검토2
            If Ks!apl_m2sab <> 0 Then
                
                .row = 33: .Col = 38
                .TypeButtonText = Ks!appmname2: .CellTag = Ks!apl_m2sab
                mname(2) = Ks!appmname2: msab(2) = Ks!apl_m2sab
                
                If Ks!apl_m2dat > 0 Then
                    
                    .CellType = CellTypeEdit: .Text = Ks!appmname2: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                        
                    .row = 35: .Text = Format(Val(Ks!apl_m2dat), "####/##/##")
                    
                End If
                
            End If
            
            '검토3
            If Ks!apl_m3sab <> 0 Then
            
                .row = 37: .Col = 34
                .TypeButtonText = Ks!appmname3: .CellTag = Ks!apl_m3sab
                mname(3) = Ks!appmname3: msab(3) = Ks!apl_m3sab
                
                If Ks!apl_m3dat > 0 Then
                    
                    .CellType = CellTypeEdit: .Text = Ks!appmname3: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                        
                    .row = 39: .Text = Format(Val(Ks!apl_m3dat), "####/##/##")
                    
                End If
                
            End If
            
            '검토4
            If Ks!apl_m4sab <> 0 Then
                
                .row = 37: .Col = 38
                .TypeButtonText = Ks!appmname4: .CellTag = Ks!apl_m4sab
                mname(4) = Ks!appmname4: msab(4) = Ks!apl_m4sab
                
                If Ks!apl_m4dat > 0 Then
                    
                    .CellType = CellTypeEdit: .Text = Ks!appmname4: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                        
                    .row = 39: .Text = Format(Val(Ks!apl_m4dat), "####/##/##")
                    
                End If
                
            End If
        
        End With
        
    ElseIf pcnt = 3 Then
        
        With sprd_print2
            
            '담당자
            If Ks!apl_isab <> 0 Then
                
                .row = 2
                .Col = 22
                .CellType = CellTypeEdit: .Text = Ks!iname: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                iname = Ks!iname: isab = Ks!apl_isab
                    
                .row = 4: .Text = Format(Val(Ks!apl_tdat), "####/##/##")
            
            End If
            
            '생산
            If Ks!apl_1sab <> 0 Then
            
                .row = 2: .Col = 26
                .TypeButtonText = Ks!appname1: .CellTag = Ks!apl_1sab
                aname(1) = Ks!appname1: asab(1) = Ks!apl_1sab
            
                If Ks!apl_1dat > 0 Then
                    
                    .CellType = CellTypeEdit: .Text = Ks!appname1: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                        
                    .row = 4: .Text = Format(Val(Ks!apl_1dat), "####/##/##")
                    
                End If
            
            End If
            
            '기술연구소
            If Ks!apl_2sab <> 0 Then
                
                .row = 2: .Col = 30
                .TypeButtonText = Ks!appname2: .CellTag = Ks!apl_2sab
                aname(2) = Ks!appname2: asab(2) = Ks!apl_2sab
                
                If Ks!apl_2dat > 0 Then
                    
                    .CellType = CellTypeEdit: .Text = Ks!appname2: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                        
                    .row = 4: .Text = Format(Val(Ks!apl_2dat), "####/##/##")
                    
                End If
                
            End If
            
            '전무님
            If Ks!apl_3sab <> 0 Then
                
                .row = 2: .Col = 34
                .TypeButtonText = Ks!appname3: .CellTag = Ks!apl_3sab
                aname(3) = Ks!appname3: asab(3) = Ks!apl_3sab
                
                If Ks!apl_3dat > 0 Then
                    
                    .CellType = CellTypeEdit: .Text = Ks!appname3: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                        
                    .row = 4: .Text = Format(Val(Ks!apl_3dat), "####/##/##")
                    
                End If
                
            End If
            
            '사장님
            If Ks!apl_4sab <> 0 Then
                
                .row = 2: .Col = 38
                .CellTag = Ks!apl_4sab
                
                If Ks!appname4 = "문창환" Then
                    .TypeButtonText = "사장님"
                    aname(4) = "사장님"
                Else
                    .TypeButtonText = Ks!appname4
                    aname(4) = Ks!appname4
                End If
                
                asab(4) = Ks!apl_4sab
                
                '승인결재가 대결이 아닌경우
                If Ks!apl_4dat > 0 And Ks!apl_4dat <> 99999999 Then
                    
                    .CellType = CellTypeEdit: .Text = "사장님": .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(255, 214, 255)
                    .row = 4: .Text = Format(Val(Ks!apl_4dat), "####/##/##")
                
                '승인결재가 대결인 경우
                ElseIf Ks!apl_4dat = 99999999 Then
                    
                     .CellType = CellTypeEdit: .Text = "대결": .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(255, 214, 255)
                     .row = 4: .Text = Format(Val(Ks!apl_alldat), "####/##/##")
                    
                End If
                
            End If
            
            '검토1
            If Ks!apl_m1sab <> 0 Then
                
                .row = 33: .Col = 34
                .TypeButtonText = Ks!appmname1: .CellTag = Ks!apl_m1sab
                mname(1) = Ks!appmname1: msab(1) = Ks!apl_m1sab
                
                If Ks!apl_m1dat > 0 Then
                    
                    .CellType = CellTypeEdit: .Text = Ks!appmname1: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                        
                    .row = 35: .Text = Format(Val(Ks!apl_m1dat), "####/##/##")
                    
                End If
                
            End If
            
            '검토2
            If Ks!apl_m2sab <> 0 Then
                
                .row = 33: .Col = 38
                .TypeButtonText = Ks!appmname2: .CellTag = Ks!apl_m2sab
                mname(2) = Ks!appmname2: msab(2) = Ks!apl_m2sab
                
                If Ks!apl_m2dat > 0 Then
                    
                    .CellType = CellTypeEdit: .Text = Ks!appmname2: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                        
                    .row = 35: .Text = Format(Val(Ks!apl_m2dat), "####/##/##")
                    
                End If
                
            End If
            
            '검토3
            If Ks!apl_m3sab <> 0 Then
            
                .row = 37: .Col = 34
                .TypeButtonText = Ks!appmname3: .CellTag = Ks!apl_m3sab
                mname(3) = Ks!appmname3: msab(3) = Ks!apl_m3sab
                
                If Ks!apl_m3dat > 0 Then
                    
                    .CellType = CellTypeEdit: .Text = Ks!appmname3: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                        
                    .row = 39: .Text = Format(Val(Ks!apl_m3dat), "####/##/##")
                    
                End If
                
            End If
            
            '검토4
            If Ks!apl_m4sab <> 0 Then
                
                .row = 37: .Col = 38
                .TypeButtonText = Ks!appmname4: .CellTag = Ks!apl_m4sab
                mname(4) = Ks!appmname4: msab(4) = Ks!apl_m4sab
                
                If Ks!apl_m4dat > 0 Then
                    
                    .CellType = CellTypeEdit: .Text = Ks!appmname4: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                        
                    .row = 39: .Text = Format(Val(Ks!apl_m4dat), "####/##/##")
                    
                End If
                
            End If
            
        End With
        
    End If
    
    Ks.Close
    
End Sub

Private Sub btn_clear_Click()
    Call clear1
    Call clear2
End Sub

Private Sub clear1()
    
    Dim ii As Double
    
    sprd_print1.row = 1
    sprd_print1.Col = 6: sprd_print1.Text = ""      'TEST NO
    sprd_print1.Col = 16: sprd_print1.Text = ""     'DATE

    sprd_print1.row = 2
    sprd_print1.Col = 6: sprd_print1.Text = ""      '접수번호
    
    sprd_print1.row = 4
    sprd_print1.Col = 1: sprd_print1.Text = ""      'TEST 제목
    
    sprd_print1.row = 4
    
    sprd_print1.Col = 7                             '사용기계/피삭재
    For ii = 6 To 13
        sprd_print1.row = ii: sprd_print1.Text = ""
    Next ii

    sprd_print1.row = 14
    sprd_print1.Col = 1: sprd_print1.Text = ""      'TEST 목적
    
    sprd_print1.row = 17
    sprd_print1.Col = 26: sprd_print1.Text = ""     '작업자
    
    sprd_print1.row = 6:
    sprd_print1.Col = 21: sprd_print1.TypePictPicture = Nothing  '이미지
    
    sprd_print1.Col = 7                             '사용기계/피삭재
    For ii = 18 To 30
        sprd_print1.row = ii: sprd_print1.Text = ""
    Next ii
    
    sprd_print1.row = 25
    sprd_print1.Col = 14: sprd_print1.Text = ""

    sprd_print1.Col = 21                             '기존공구내역
    For ii = 18 To 30
        sprd_print1.row = ii: sprd_print1.Text = ""
    Next ii
    
    sprd_print1.row = 25                             '테스트1 내역
    sprd_print1.Col = 28: sprd_print1.Text = ""
    
    
    sprd_print1.row = 19
    sprd_print1.Col = 34: sprd_print1.Text = ""      '비고
    
    sprd_print1.row = 28
    sprd_print1.Col = 34: sprd_print1.Text = ""      '평가
    
    sprd_print1.row = 31
    sprd_print1.Col = 7: sprd_print1.Text = ""       '결과 OK
    
    sprd_print1.row = 31
    sprd_print1.Col = 14: sprd_print1.Text = ""      '결과이유
    
    sprd_print1.row = 36
    sprd_print1.Col = 7: sprd_print1.Text = ""       '결과 OK
    
    sprd_print1.row = 36
    sprd_print1.Col = 14: sprd_print1.Text = ""      '결과이유
    
    sprd_print1.row = 41
    sprd_print1.Col = 7: sprd_print1.Text = ""       '재 TEST여부
    
    sprd_print1.Col = 7
    sprd_print1.row = 18: sprd_print1.FontBold = False: sprd_print1.RowHeight(18) = 14
    sprd_print1.row = 19: sprd_print1.FontBold = False: sprd_print1.RowHeight(19) = 14
                    
    sprd_print1.row = 28: sprd_print1.FontBold = False: sprd_print1.RowHeight(28) = 14
    sprd_print1.row = 29: sprd_print1.FontBold = False: sprd_print1.RowHeight(29) = 14
    sprd_print1.row = 30: sprd_print1.FontBold = False: sprd_print1.RowHeight(30) = 14

    sprd_print1.Col = 21
    sprd_print1.row = 18: sprd_print1.FontBold = False: sprd_print1.RowHeight(18) = 14
    sprd_print1.row = 19: sprd_print1.FontBold = False: sprd_print1.RowHeight(19) = 14
                    
    sprd_print1.row = 28: sprd_print1.FontBold = False: sprd_print1.RowHeight(28) = 14
    sprd_print1.row = 29: sprd_print1.FontBold = False: sprd_print1.RowHeight(29) = 14
    sprd_print1.row = 30: sprd_print1.FontBold = False: sprd_print1.RowHeight(30) = 14
    
    With sprd_print1
    
        .row = 2
        .Col = 22: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 26: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 30: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 34: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 38: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        
        .row = 4
        .Col = 22: .Text = ""
        .Col = 26: .Text = ""
        .Col = 30: .Text = ""
        .Col = 34: .Text = ""
        .Col = 38: .Text = ""

        .row = 33
        .Col = 34: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 38: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        
        .row = 35
        .Col = 34: .Text = ""
        .Col = 38: .Text = ""
        
        .row = 37
        .Col = 34: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 38: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        
        .row = 39
        .Col = 34: .Text = ""
        .Col = 38: .Text = ""
        
    End With
    
End Sub

Private Sub print2()

    Dim purpose As String
    Dim resultOK As String
    Dim resultNG As String
    
'    sss = "       select *"
'    sss = sss & "   from man_tooltesthd, man_tooltestds"
'    sss = sss & "  where tth_dat = ttd_dat"
'    sss = sss & "    and tth_seq = ttd_seq"
'    sss = sss & "    and tth_dat = " & txt_dat1
'    sss = sss & "    and tth_seq = " & txt_seq1
'    sss = sss & "  order by ttd_lno"
    
    sss = "       select *"
    sss = sss & "   from man_tooltesthd, man_tooltestds"
    sss = sss & "  where tth_dat = ttd_dat"
    sss = sss & "    and tth_seq = ttd_seq"
    sss = sss & "    and tth_dat = " & txt_dat1
    sss = sss & "    and tth_seq = " & txt_seq1
    sss = sss & "  order by ttd_lno"
    
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Rs.RecordCount < 1 Then
        Rs.Close
        Call msg_display("등록된 내역이 없습니다!")
        Exit Sub
    End If
    
    Call clear2
    Call app
    
    If Not IsNull(Rs!tth_testno) Then sprd_print2.row = 1: sprd_print2.Col = 6: sprd_print2.Text = Rs!tth_testno
    If Not IsNull(Rs!tth_tdat) Then sprd_print2.row = 1: sprd_print2.Col = 16: sprd_print2.Text = Format(Rs!tth_tdat, "0000-00-00")
    If Not IsNull(Rs!tth_jubno) Then sprd_print2.row = 2: sprd_print2.Col = 6: sprd_print2.Text = Format(Rs!tth_jubno, "00000000-000")

    If Not IsNull(Rs!tth_title) Then sprd_print2.row = 4: sprd_print2.Col = 1: sprd_print2.Text = Rs!tth_title
    
    If Not IsNull(Rs!tth_tmcd) Then
        sprd_print2.row = 8: sprd_print2.Col = 7: sprd_print2.Text = Rs!tth_tmcd
    
        sss = "       select mhc_name, ems_mark, sok_name(mhc_sok) soknm "
        sss = sss & "   from man_machcd, eam_mast"
        sss = sss & "  where mhc_code = '" & Rs!tth_tmcd & "'"
        sss = sss & "    and mhc_code = ems_mcd(+)"
                
        Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        If Ks.RecordCount > 0 Then
            If Not IsNull(Ks!mhc_name) Then sprd_print2.row = 7: sprd_print2.Col = 7: sprd_print2.Text = Ks!mhc_name
            If Not IsNull(Ks!ems_mark) Then sprd_print2.row = 6: sprd_print2.Col = 7: sprd_print2.Text = Ks!ems_mark
            If Not IsNull(Ks!soknm) Then sprd_print2.row = 9: sprd_print2.Col = 7: sprd_print2.Text = Ks!soknm
        End If
        
        Ks.Close
    
    End If

    If Not IsNull(Rs!tth_tlot) Then
        
        sss = "       select dit_bpcd, dit_bpjil, dit_jacd, dit_jajil"
        sss = sss & "   from man_direct"
        sss = sss & "  where dit_lot = '" & Rs!tth_tlot & "'"
            
        Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        If Ks.RecordCount > 0 Then
            
            If Not IsNull(Ks!dit_bpcd) Then sprd_print2.row = 10: sprd_print2.Col = 7: sprd_print2.Text = Ks!dit_bpcd
            If Not IsNull(Ks!dit_bpjil) Then sprd_print2.row = 11: sprd_print2.Col = 7: sprd_print2.Text = Ks!dit_bpjil
            If Not IsNull(Ks!dit_jacd) Then sprd_print2.row = 12: sprd_print2.Col = 7: sprd_print2.Text = Ks!dit_jacd
            If Not IsNull(Ks!dit_jajil) Then sprd_print2.row = 13: sprd_print2.Col = 7: sprd_print2.Text = Ks!dit_jajil
        
        End If
    
        Ks.Close
    
    End If
    
    purpose = "        TEST 목적" & Chr(13) & Chr(13)
    
    If Rs!tth_pyn1 = "Y" Then
        purpose = purpose & "        1.공구 수명 ( ○ )"
    Else
        purpose = purpose & "        1.공구 수명 (     )"
    End If

    If Rs!tth_pyn2 = "Y" Then
        purpose = purpose & "             2.칩 처리 ( ○ )" & Chr(13)
    Else
        purpose = purpose & "             2.칩 처리 (     )" & Chr(13)
    End If


    If Rs!tth_pyn3 = "Y" Then
        purpose = purpose & "        3.시간 단축 ( ○ )"
    Else
        purpose = purpose & "        3.시간 단축 (     )"
    End If


    If Rs!tth_pyn4 = "Y" Then
        purpose = purpose & "             4.공구비 절감 ( ○ )" & Chr(13)
    Else
        purpose = purpose & "             4.공구비 절감 (     )" & Chr(13)
    End If


    If Rs!tth_pyn5 = "Y" Then
        purpose = purpose & "        5.기타 ( ○ )"
    Else
        purpose = purpose & "        5.기타 (     )"
    End If
    
    sprd_print2.row = 14: sprd_print2.Col = 1: sprd_print2.Text = purpose

    If Not IsNull(Rs!tth_tsab) Then
        
        sss = "       select sin_name, sin_sok, sok_name(sin_sok) soknm"
        sss = sss & "   from peo_sinbun"
        sss = sss & "  where sin_sab = " & Rs!tth_tsab
        sss = sss & "    and sin_taedt = 0 "
            
        Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        If Ks.RecordCount > 0 Then
            sprd_print2.row = 17: sprd_print2.Col = 26: sprd_print2.Text = Ks!sin_name & "(" & Ks!soknm & ")"
        End If
        Ks.Close
    
    End If
    
    '도면 이미지
    If Not IsNull(Rs!tth_file1) Then
        If Right(Rs!tth_file1, 3) = "JPG" Then
           '
           '==============================
           ' FTP를 이용한 파일 다운로드
           '==============================
           If FTP_Connection Then
              '
              If Not FTP경로체크(SERVER_PATH) Then
                 Call FTP_DisConnect
                 MsgBox "서버경로를 찾을수 없습니다.(정보관리센터 문의)"
                 Exit Sub
              End If
              '
              GPath = "c:\jpeg\" & Rs!tth_file1
              '
              If FTP_Download(SERVER_PATH & Rs!tth_file1, GPath) Then
                 sprd_print2.row = 6: sprd_print2.Col = 21: sprd_print2.TypePictPicture = LoadPicture(GPath)
                 sprd_print2.TypePictStretch = True
                 sprd_print2.TypePictMaintainScale = True
              End If
              '
              Call FTP_DisConnect
              '
           End If
           '
           '======================================
           ' exists = ExistFile(N_Driver & ":\" & Rs!tth_file1)
           ' If exists = True Then
           '     FileCopy N_Driver & ":\" & Rs!tth_file1, "c:\temp\" & "Temp.jpg"
           '     GPath = "c:\temp\Temp.jpg"
           '     sprd_print2.Row = 6: sprd_print2.Col = 21: sprd_print2.TypePictPicture = LoadPicture(GPath)
           ' End If
        End If
    End If
    
    If Not IsNull(Rs!tth_rmk) Then sprd_print2.row = 19: sprd_print2.Col = 34: sprd_print2.Text = Rs!tth_rmk
    If Not IsNull(Rs!tth_cmt) Then sprd_print2.row = 28: sprd_print2.Col = 34: sprd_print2.Text = Rs!tth_cmt
    
    'DESC
    Do While Not Rs.EOF
        
        '기존공구
        If Rs!ttd_lno = 1 Then
            sprd_print2.Col = 7
            If Not IsNull(Rs!ttd_maker) Then sprd_print2.row = 18: sprd_print2.Text = Rs!ttd_maker & "(기존)"
            If Not IsNull(Rs!ttd_tipstd) Then sprd_print2.row = 19: sprd_print2.Text = Rs!ttd_tipstd
            If Not IsNull(Rs!ttd_tipjil) Then sprd_print2.row = 20: sprd_print2.Text = Rs!ttd_tipjil
            If Not IsNull(Rs!ttd_holder) Then sprd_print2.row = 21: sprd_print2.Text = Rs!ttd_holder
            If Not IsNull(Rs!ttd_rcntmn) Then sprd_print2.row = 22: sprd_print2.Text = Rs!ttd_rcntmn
            If Not IsNull(Rs!ttd_rcntmx) Then sprd_print2.row = 22: sprd_print2.Text = sprd_print2.Text & "-" & Rs!ttd_rcntmx
            If Not IsNull(Rs!ttd_movmn) Then sprd_print2.row = 24: sprd_print2.Text = Rs!ttd_movmn
            If Not IsNull(Rs!ttd_movmx) Then sprd_print2.row = 24: sprd_print2.Text = sprd_print2.Text & "-" & Rs!ttd_movmx
            If Not IsNull(Rs!ttd_depth) Then sprd_print2.row = 23: sprd_print2.Text = Rs!ttd_depth & "mm"
            If Not IsNull(Rs!ttd_tct) Then sprd_print2.row = 25: sprd_print2.Text = Format(Rs!ttd_tct, "###,##0")
            If Not IsNull(Rs!ttd_pct) Then sprd_print2.row = 25: sprd_print2.Col = 12: sprd_print2.Text = Format(Rs!ttd_pct, "###,##0")
            
            If Not IsNull(Rs!ttd_fluid) Then
                sprd_print2.row = 26: sprd_print2.Col = 7:
                If Rs!ttd_fluid = 1 Then
                    sprd_print2.Text = "수용성"
                Else
                    sprd_print2.Text = "비수용성"
                End If
            End If
            
            sprd_print2.Col = 7
            If Not IsNull(Rs!ttd_qty) Then sprd_print2.row = 27: sprd_print2.Text = Format(Rs!ttd_qty, "###,##0") & " Point"
            'If Not IsNull(Rs!ttd_dan) Then sprd_print2.Row = 28: sprd_print2.Text = Format(Rs!ttd_dan, "###,###,###") & "원"
            'If Not IsNull(Rs!ttd_tldn) Then sprd_print2.Row = 29: sprd_print2.Text = Format(Rs!ttd_tldn, "###,###,###") & "원/EA"
            'If Not IsNull(Rs!ttd_chdn) Then sprd_print2.Row = 30: sprd_print2.Text = Format(Rs!ttd_chdn, "###,###,###") & "원"
            
            If Not IsNull(Rs!ttd_dan) Then
                If InStr(1, Rs!ttd_dan, ".", 0) <> 0 Then
                    sprd_print2.row = 28: sprd_print2.Text = Format(Rs!ttd_dan, "###,###,##0.0#") & "원"
                Else
                    sprd_print2.row = 28: sprd_print2.Text = Format(Rs!ttd_dan, "###,###,##0") & "원"
                    
                End If
            End If
            If Not IsNull(Rs!ttd_tldn) Then
                If InStr(1, Rs!ttd_tldn, ".", 0) <> 0 Then
                    sprd_print2.row = 29: sprd_print2.Text = Format(Rs!ttd_tldn, "###,###,##0.0#") & "원/Corner"
                Else
                    sprd_print2.row = 29: sprd_print2.Text = Format(Rs!ttd_tldn, "###,###,##0") & "원/Corner"
                End If
            End If
            If Not IsNull(Rs!ttd_chdn) Then
                If InStr(1, Rs!ttd_chdn, ".", 0) <> 0 Then
                    sprd_print2.row = 30: sprd_print2.Text = Format(Rs!ttd_chdn, "###,###,##0.0#") & "원"
                Else
                    sprd_print2.row = 30: sprd_print2.Text = Format(Rs!ttd_chdn, "###,###,##0") & "원/EA"
                End If
            End If
            
        End If
        
        '테스트1
        If Rs!ttd_lno = 2 Then
        
            sprd_print2.Col = 16
            If Not IsNull(Rs!ttd_maker) Then sprd_print2.row = 18: sprd_print2.Text = Rs!ttd_maker & "(테스트)"
            If Not IsNull(Rs!ttd_tipstd) Then sprd_print2.row = 19: sprd_print2.Text = Rs!ttd_tipstd
            If Not IsNull(Rs!ttd_tipjil) Then sprd_print2.row = 20: sprd_print2.Text = Rs!ttd_tipjil
            If Not IsNull(Rs!ttd_holder) Then sprd_print2.row = 21: sprd_print2.Text = Rs!ttd_holder
            If Not IsNull(Rs!ttd_rcntmn) Then sprd_print2.row = 22: sprd_print2.Text = Rs!ttd_rcntmn
            If Not IsNull(Rs!ttd_rcntmx) Then sprd_print2.row = 22: sprd_print2.Text = sprd_print2.Text & "-" & Rs!ttd_rcntmx
            If Not IsNull(Rs!ttd_movmn) Then sprd_print2.row = 24: sprd_print2.Text = Rs!ttd_movmn
            If Not IsNull(Rs!ttd_movmx) Then sprd_print2.row = 24: sprd_print2.Text = sprd_print2.Text & "-" & Rs!ttd_movmx
            If Not IsNull(Rs!ttd_depth) Then sprd_print2.row = 23: sprd_print2.Text = Rs!ttd_depth & "mm"
            If Not IsNull(Rs!ttd_tct) Then sprd_print2.row = 25: sprd_print2.Text = Format(Rs!ttd_tct, "###,##0")
            If Not IsNull(Rs!ttd_pct) Then sprd_print2.row = 25: sprd_print2.Col = 21: sprd_print2.Text = Format(Rs!ttd_pct, "###,##0")
            
            If Not IsNull(Rs!ttd_fluid) Then
                sprd_print2.row = 26: sprd_print2.Col = 16:
                If Rs!ttd_fluid = 1 Then
                    sprd_print2.Text = "수용성"
                Else
                    sprd_print2.Text = "비수용성"
                End If
            End If
            
            sprd_print2.Col = 16
            
            If Not IsNull(Rs!ttd_qty) Then sprd_print2.row = 27: sprd_print2.Text = Format(Rs!ttd_qty, "###,##0") & " Point"
            
            If Not IsNull(Rs!ttd_dan) Then
                If InStr(1, Rs!ttd_dan, ".", 0) <> 0 Then
                    sprd_print2.row = 28: sprd_print2.Text = Format(Rs!ttd_dan, "###,###,##0.0#") & "원"
                Else
                    sprd_print2.row = 28: sprd_print2.Text = Format(Rs!ttd_dan, "###,###,##0") & "원"
                    
                End If
            End If
            If Not IsNull(Rs!ttd_tldn) Then
                If InStr(1, Rs!ttd_tldn, ".", 0) <> 0 Then
                    sprd_print2.row = 29: sprd_print2.Text = Format(Rs!ttd_tldn, "###,###,##0.0#") & "원/Corner"
                Else
                    sprd_print2.row = 29: sprd_print2.Text = Format(Rs!ttd_tldn, "###,###,##0") & "원/Corner"
                End If
            End If
            If Not IsNull(Rs!ttd_chdn) Then
                If InStr(1, Rs!ttd_chdn, ".", 0) <> 0 Then
                    sprd_print2.row = 30: sprd_print2.Text = Format(Rs!ttd_chdn, "###,###,##0.0#") & "원"
                Else
                    sprd_print2.row = 30: sprd_print2.Text = Format(Rs!ttd_chdn, "###,###,##0") & "원/EA"
                End If
            End If
            
        End If
        
        '테스트2
        If Rs!ttd_lno = 3 Then
            
            sprd_print2.Col = 25
            If Not IsNull(Rs!ttd_maker) Then sprd_print2.row = 18:  sprd_print2.Text = Rs!ttd_maker & "(테스트)"
            If Not IsNull(Rs!ttd_tipstd) Then sprd_print2.row = 19: sprd_print2.Text = Rs!ttd_tipstd
            If Not IsNull(Rs!ttd_tipjil) Then sprd_print2.row = 20: sprd_print2.Text = Rs!ttd_tipjil
            If Not IsNull(Rs!ttd_holder) Then sprd_print2.row = 21: sprd_print2.Text = Rs!ttd_holder
            If Not IsNull(Rs!ttd_rcntmn) Then sprd_print2.row = 22: sprd_print2.Text = Rs!ttd_rcntmn
            If Not IsNull(Rs!ttd_rcntmx) Then sprd_print2.row = 22: sprd_print2.Text = sprd_print2.Text & "-" & Rs!ttd_rcntmx
            If Not IsNull(Rs!ttd_movmn) Then sprd_print2.row = 24:  sprd_print2.Text = Rs!ttd_movmn
            If Not IsNull(Rs!ttd_movmx) Then sprd_print2.row = 24:  sprd_print2.Text = sprd_print2.Text & "-" & Rs!ttd_movmx
            If Not IsNull(Rs!ttd_depth) Then sprd_print2.row = 23:  sprd_print2.Text = Rs!ttd_depth & "mm"
            If Not IsNull(Rs!ttd_tct) Then sprd_print2.row = 25:    sprd_print2.Text = Format(Rs!ttd_tct, "###,##0")
            If Not IsNull(Rs!ttd_pct) Then sprd_print2.row = 25:    sprd_print2.Col = 30: sprd_print2.Text = Format(Rs!ttd_pct, "###,##0")
            
            If Not IsNull(Rs!ttd_fluid) Then
                sprd_print2.row = 26: sprd_print2.Col = 25:
                If Rs!ttd_fluid = 1 Then
                    sprd_print2.Text = "수용성"
                Else
                    sprd_print2.Text = "비수용성"
                End If
            End If
            
            sprd_print2.Col = 25
            
            If Not IsNull(Rs!ttd_qty) Then sprd_print2.row = 27: sprd_print2.Text = Format(Rs!ttd_qty, "###,##0") & " Point"
            
            If Not IsNull(Rs!ttd_dan) Then
                If InStr(1, Rs!ttd_dan, ".", 0) <> 0 Then
                    sprd_print2.row = 28: sprd_print2.Text = Format(Rs!ttd_dan, "###,###,##0.0#") & "원"
                Else
                    sprd_print2.row = 28: sprd_print2.Text = Format(Rs!ttd_dan, "###,###,##0") & "원"
                    
                End If
            End If
            If Not IsNull(Rs!ttd_tldn) Then
                If InStr(1, Rs!ttd_tldn, ".", 0) <> 0 Then
                    sprd_print2.row = 29: sprd_print2.Text = Format(Rs!ttd_tldn, "###,###,##0.0#") & "원/Corner"
                Else
                    sprd_print2.row = 29: sprd_print2.Text = Format(Rs!ttd_tldn, "###,###,##0") & "원/Corner"
                End If
            End If
            If Not IsNull(Rs!ttd_chdn) Then
                If InStr(1, Rs!ttd_chdn, ".", 0) <> 0 Then
                    sprd_print2.row = 30: sprd_print2.Text = Format(Rs!ttd_chdn, "###,###,##0.0#") & "원"
                Else
                    sprd_print2.row = 30: sprd_print2.Text = Format(Rs!ttd_chdn, "###,###,##0") & "원/EA"
                End If
            End If
            
        End If
        
        If Rs!ttd_ryn = "Y" Then
               
            resultOK = ""
            resultNG = ""
            
            If Not IsNull(Rs!ttd_result) Then
                
                If Rs!ttd_result = "OK" Then
                    sprd_print2.row = 31: sprd_print2.Col = 7
                    sprd_print2.Text = "  O.K ( ○ )": 'sprd_print2.FontBold = True
                    sprd_print2.row = 36: sprd_print2.Col = 7
                    sprd_print2.Text = "  N.G (     )"
                End If
                If Rs!ttd_result = "NG" Then
                    sprd_print2.row = 31: sprd_print2.Col = 7
                    sprd_print2.Text = "  O.K (     )"
                    sprd_print2.row = 36: sprd_print2.Col = 7
                    sprd_print2.Text = "  N.G ( ○ )": 'sprd_print2.FontBold = True
                End If
            End If

            If Rs!ttd_result = "OK" Then
                
                If Rs!ttd_ryn1 = "Y" Then
                    resultOK = resultOK & "           1.공구수명 연장 ( ○ )"
                Else
                    resultOK = resultOK & "           1.공구수명 연장 (     )"
                End If
    
                If Rs!ttd_ryn4 = "Y" Then
                    resultOK = resultOK & "          4.시간 단축 ( ○ )" & Chr(13)
                Else
                    resultOK = resultOK & "          4.시간 단축 (     )" & Chr(13)
                End If
                    
                If Rs!ttd_ryn2 = "Y" Then
                    resultOK = resultOK & "           2.칩처리 양호 ( ○ )"
                Else
                    resultOK = resultOK & "           2.칩처리 양호 (     )"
                End If
                    
                If Rs!ttd_ryn5 = "Y" Then
                    resultOK = resultOK & "             5.기타 ( ○ )" & Chr(13)
                Else
                    resultOK = resultOK & "             5.기타 (     )" & Chr(13)
                End If
                    
                If Rs!ttd_ryn3 = "Y" Then
                    resultOK = resultOK & "           3.공구비 절감 ( ○ )"
                Else
                    resultOK = resultOK & "           3.공구비 절감 (     )"
                End If
                    
                resultNG = resultNG & "           1.결손 (     )"
                resultNG = resultNG & "                      4.공구비 상승 (     )" & Chr(13)
                resultNG = resultNG & "           2.마모 (     )"
                resultNG = resultNG & "                      5.기타 (     )" & Chr(13)
                resultNG = resultNG & "           3.칩처리 불량 (     )"
                
                
            Else
                
                resultNG = ""
                
                If Rs!ttd_ryn1 = "Y" Then
                    resultNG = resultNG & "           1.결손 ( ○ )"
                Else
                    resultNG = resultNG & "           1.결손 (     )"
                End If
    
                If Rs!ttd_ryn4 = "Y" Then
                    resultNG = resultNG & "                      4.공구비 상승 ( ○ )" & Chr(13)
                Else
                    resultNG = resultNG & "                      4.공구비 상승 (     )" & Chr(13)
                End If
                    
                If Rs!ttd_ryn2 = "Y" Then
                    resultNG = resultNG & "           2.마모 ( ○ )"
                Else
                    resultNG = resultNG & "           2.마모 (     )"
                End If
                                                 
                If Rs!ttd_ryn5 = "Y" Then
                    resultNG = resultNG & "                      5.기타 ( ○ )" & Chr(13)
                Else
                    resultNG = resultNG & "                      5.기타 (     )" & Chr(13)
                End If
                                                                   
                If Rs!ttd_ryn3 = "Y" Then
                    resultNG = resultNG & "           3.칩처리 불량 ( ○ )"
                Else
                    resultNG = resultNG & "           3.칩처리 불량 (     )"
                End If
                        
                resultOK = resultOK & "           1.공구수명 연장 (     )"
                resultOK = resultOK & "          4.시간 단축 (     )" & Chr(13)
                resultOK = resultOK & "           2.칩처리 양호 (     )"
                resultOK = resultOK & "             5.기타 (     )" & Chr(13)
                resultOK = resultOK & "           3.공구비 절감 (     )"

            End If
                
            sprd_print2.row = 31: sprd_print2.Col = 14: sprd_print2.Text = resultOK
            sprd_print2.row = 36: sprd_print2.Col = 14: sprd_print2.Text = resultNG
            
            sprd_print2.row = 41: sprd_print2.Col = 7
            If Rs!ttd_rtyn = "Y" Then
                sprd_print2.Text = "재 TEST 가능 여부 - OK"
            Else
                sprd_print2.Text = "재 TEST 가능 여부 - NOT"
            End If
                            
            If Rs!ttd_lno = 1 Then
                sprd_print2.Col = 7
            ElseIf Rs!ttd_lno = 2 Then
                sprd_print2.Col = 16
            ElseIf Rs!ttd_lno = 3 Then
                sprd_print2.Col = 25
            End If
                
            sprd_print2.row = 18: sprd_print2.FontBold = True: sprd_print2.RowHeight(18) = 14
            sprd_print2.row = 19: sprd_print2.FontBold = True: sprd_print2.RowHeight(19) = 14
                    
            sprd_print2.row = 28: sprd_print2.FontBold = True: sprd_print2.RowHeight(28) = 14
            sprd_print2.row = 29: sprd_print2.FontBold = True: sprd_print2.RowHeight(29) = 14
            sprd_print2.row = 30: sprd_print2.FontBold = True: sprd_print2.RowHeight(30) = 14
                
        End If
            
        Rs.MoveNext
    Loop
    
    Rs.Close
    
    'btn_view1.Enabled = False
    
    'txt_dat1.Enabled = False
    'txt_seq1.Enabled = False
    
    'btn_add1.Enabled = False
    'btn_mod1.Enabled = True
    'btn_del1.Enabled = True
    
    Call msg_display("조회완료!")
    
    
End Sub

Private Sub clear2()
    
    Dim ii As Double
    
    sprd_print2.row = 1
    sprd_print2.Col = 6: sprd_print2.Text = ""      'TEST NO
    sprd_print2.Col = 16: sprd_print2.Text = ""     'DATE

    sprd_print2.row = 2
    sprd_print2.Col = 6: sprd_print2.Text = ""      '접수번호
    
    sprd_print2.row = 4
    sprd_print2.Col = 1: sprd_print2.Text = ""      'TEST 제목
    
    sprd_print2.row = 4
    
    sprd_print2.Col = 7                             '사용기계/피삭재
    For ii = 6 To 13
        sprd_print2.row = ii: sprd_print2.Text = ""
    Next ii

    sprd_print2.row = 14
    sprd_print2.Col = 1: sprd_print2.Text = ""      'TEST 목적
    
    sprd_print2.row = 17
    sprd_print2.Col = 26: sprd_print2.Text = ""     '작업자
    
    sprd_print2.row = 6:
    sprd_print2.Col = 21: sprd_print2.TypePictPicture = Nothing  '이미지
    
    sprd_print2.Col = 7                              '기존공구내역
    For ii = 18 To 30
        sprd_print2.row = ii: sprd_print2.Text = ""
    Next ii
    sprd_print2.row = 25
    sprd_print2.Col = 12: sprd_print2.Text = ""

    sprd_print2.Col = 16                             '테스트1 내역
    For ii = 18 To 30
        sprd_print2.row = ii: sprd_print2.Text = ""
    Next ii
    sprd_print2.row = 25
    sprd_print2.Col = 21: sprd_print2.Text = ""

    sprd_print2.Col = 25                             '테스트2 내역
    For ii = 18 To 30
        sprd_print2.row = ii: sprd_print2.Text = ""
    Next ii
    sprd_print2.row = 25
    sprd_print2.Col = 30: sprd_print2.Text = ""
    
    
    sprd_print2.row = 19
    sprd_print2.Col = 34: sprd_print2.Text = ""      '비고
    
    sprd_print2.row = 28
    sprd_print2.Col = 34: sprd_print2.Text = ""      '평가
    
    sprd_print2.row = 31
    sprd_print2.Col = 7: sprd_print2.Text = ""       '결과 OK
    
    sprd_print2.row = 31
    sprd_print2.Col = 14: sprd_print2.Text = ""      '결과이유
    
    sprd_print2.row = 36
    sprd_print2.Col = 7: sprd_print2.Text = ""       '결과 OK
    
    sprd_print2.row = 36
    sprd_print2.Col = 14: sprd_print2.Text = ""      '결과이유
    
    sprd_print2.row = 41
    sprd_print2.Col = 7: sprd_print2.Text = ""       '재 TEST여부
    
    sprd_print2.Col = 7
    sprd_print2.row = 18: sprd_print2.FontBold = False: sprd_print2.RowHeight(18) = 14
    sprd_print2.row = 19: sprd_print2.FontBold = False: sprd_print2.RowHeight(19) = 14
                    
    sprd_print2.row = 28: sprd_print2.FontBold = False: sprd_print2.RowHeight(28) = 14
    sprd_print2.row = 29: sprd_print2.FontBold = False: sprd_print2.RowHeight(29) = 14
    sprd_print2.row = 30: sprd_print2.FontBold = False: sprd_print2.RowHeight(30) = 14

    sprd_print2.Col = 16
    sprd_print2.row = 18: sprd_print2.FontBold = False: sprd_print2.RowHeight(18) = 14
    sprd_print2.row = 19: sprd_print2.FontBold = False: sprd_print2.RowHeight(19) = 14
                    
    sprd_print2.row = 28: sprd_print2.FontBold = False: sprd_print2.RowHeight(28) = 14
    sprd_print2.row = 29: sprd_print2.FontBold = False: sprd_print2.RowHeight(29) = 14
    sprd_print2.row = 30: sprd_print2.FontBold = False: sprd_print2.RowHeight(30) = 14
    
    sprd_print2.Col = 25
    sprd_print2.row = 18: sprd_print2.FontBold = False: sprd_print2.RowHeight(18) = 14
    sprd_print2.row = 19: sprd_print2.FontBold = False: sprd_print2.RowHeight(19) = 14
                    
    sprd_print2.row = 28: sprd_print2.FontBold = False: sprd_print2.RowHeight(28) = 14
    sprd_print2.row = 29: sprd_print2.FontBold = False: sprd_print2.RowHeight(29) = 14
    sprd_print2.row = 30: sprd_print2.FontBold = False: sprd_print2.RowHeight(30) = 14
    
    With sprd_print2
    
        .row = 2
        .Col = 22: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 26: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 30: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 34: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 38: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        
        .row = 4
        .Col = 22: .Text = ""
        .Col = 26: .Text = ""
        .Col = 30: .Text = ""
        .Col = 34: .Text = ""
        .Col = 38: .Text = ""

        .row = 33
        .Col = 34: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 38: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        
        .row = 35
        .Col = 34: .Text = ""
        .Col = 38: .Text = ""
        
        .row = 37
        .Col = 34: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 38: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        
        .row = 39
        .Col = 34: .Text = ""
        .Col = 38: .Text = ""
        
    End With
    
End Sub

'파일존재유무 체크
Private Function ExistFile(FilePath As String) As Long
     If LenB(Dir$(FilePath)) Then
          ExistFile = 1&
     Else
          ExistFile = 0&
     End If
End Function

Public Sub msg_display(mass)
   
   Dim jj As Integer
   Dim msg_len As Integer
   Dim pausetime As Single
   Dim start As Single
   
   msg_len = Len(Trim(mass))
   msg.Caption = ""
   Beep
   
   For jj = 1 To msg_len
       msg.Caption = Space(msg_len - jj + 2) & LTrim(msg.Caption)
       msg.Caption = msg.Caption & Mid(mass, jj, 1)
  '
       pausetime = 0.01    ' 기간을 지정합니다.
       start = Timer       ' 시작 시간을 지정합니다.
       Do While Timer < start + pausetime
          DoEvents         ' 다른 프로시저로 넘깁니다.
       Loop
   Next

End Sub


Private Sub Form_Load()
   '
   sprd_print1.Left = 120
   sprd_print1.Top = 630
   
   sprd_print2.Left = 120
   sprd_print2.Top = 630
   
   sprd_print3.Left = 120
   sprd_print3.Top = 630
   
   If Job_Level > 8 Then
    chk_app.Visible = True
    btn_msg.Visible = True
   Else
    chk_app.Visible = False
    btn_msg.Visible = False
   End If
   

End Sub



Private Sub sprd_print1_ButtonClicked(ByVal Col As Long, ByVal row As Long, ByVal ButtonDown As Integer)
      
    On Error GoTo err_rtn
    
    Dim title       As String      '쪽지발송 시 사용할 title 변수
    Dim ii          As Integer
    
    If Len(txt_dat1) < 8 Then
        msg = "등록번호를 확인하세요."
        Exit Sub
    End If
    
    If Len(txt_seq1) = 0 Then
        msg = "등록번호를 확인하세요."
        Exit Sub
    End If
    
    '결재자 지정
    If chk_app.Value = 1 Then
        
        If (Col = 26 Or Col = 30 Or Col = 34) And row = 2 Then
            
            Call mkpoen05_app.index_Send(Col, row, 1)
            
            Load mkpoen05_app
            mkpoen05_app.Show 1
            
        End If
        
        If (Col = 34 Or Col = 38) And (row = 33 Or row = 37) Then
        
            Call mkpoen05_app.index_Send(Col, row, 1)
            
            Load mkpoen05_app
            mkpoen05_app.Show 1
            
        End If
    
    '결재
    Else
        
        With sprd_print1
            
            '결재쪽지 문자열 생성====
            .row = 4: .Col = 1: title = .Text
            
            '사장님 결재완료 시
            If row = 2 And Col = 38 Then
                memo = "TEST DATA 결재가 완료되었습니다." & vbCrLf & vbCrLf
                memo = memo & "등록번호: " & Val(txt_dat1) & "-" & Val(txt_seq1) & vbCrLf
                memo = memo & "담당자:" & iname & vbCrLf
                memo = memo & "TEST 제목: " & title & vbCrLf & vbCrLf
                memo = memo & "연결PG:Tol\mkpoen05"
            '관련부서 결재 완료시
            ElseIf (row = 33 And Col = 34) Or (row = 33 And Col = 38) Or (row = 37 And Col = 34) Or (row = 37 And Col = 38) Then
                memo = "TEST DATA 결재가 모두완료되었습니다. " & vbCrLf
                memo = memo & "사장님 결재를 진행해주세요!!" & vbCrLf & vbCrLf
                memo = memo & "등록번호: " & Val(txt_dat1) & "-" & Val(txt_seq1) & vbCrLf
                memo = memo & "담당자:" & iname & vbCrLf
                memo = memo & "TEST 제목: " & title & vbCrLf & vbCrLf
                memo = memo & "연결PG:Tol\mkpoen05"
            '그 외 결재시
            Else
                memo = "TEST DATA 가 등록되었습니다." & vbCrLf
                memo = memo & "확인후 본인 결재란에 결재 바랍니다." & vbCrLf & vbCrLf
                memo = memo & "등록번호: " & Val(txt_dat1) & "-" & Val(txt_seq1) & vbCrLf
                memo = memo & "담당자:" & iname & vbCrLf
                memo = memo & "TEST 제목: " & title & vbCrLf & vbCrLf
                memo = memo & "연결PG:Tol\mkpoen05"
            End If
            '========================
            
            Ws.BeginTrans
            
            '생산,기술연구소,전무님,사장님 결재
            If (row = 2 And Col = 26) Or (row = 2 And Col = 30) Or (row = 2 And Col = 34) Or (row = 2 And Col = 38) Then
            
                .Col = Col
                .row = 2
                
                '결재자 미지정시 결재 불가
                If .CellTag = 0 Then
                    msg = "결재자가 지정되지 않았습니다."
                    Exit Sub
                End If
                
                '전무님 승인결재 가능(대결 or 사장님 전자결재를 진행해야하기 때문)
                If Col = 38 And row = 2 And Gsab = asab(3) Then
                    
                    txt_dae.Text = ""
                    
                    Load mkpoen05_dae
                    mkpoen05_dae.Show 1
                    
                    If Val(txt_dae.Text) = 0 Then
                        Exit Sub
                    End If
                
                Else
                    If Gsab <> .CellTag Then
                        msg = "결재자만 결재할 수 있습니다."
                        Exit Sub
                    End If
                End If
                '
        
                sss = "update oth_applist set"
                If Col = 26 Then sss = sss & " apl_1yn = 'Y',"
                If Col = 26 Then sss = sss & " apl_1dat = " & Format(Now, "yyyymmdd")
                If Col = 30 Then sss = sss & " apl_2yn = 'Y',"
                If Col = 30 Then sss = sss & " apl_2dat = " & Format(Now, "yyyymmdd")
                If Col = 34 Then sss = sss & " apl_3yn = 'Y',"
                If Col = 34 Then sss = sss & " apl_3dat = " & Format(Now, "yyyymmdd")
                If Col = 38 Then sss = sss & " apl_4yn = 'Y',"
                If Col = 38 Then sss = sss & " apl_4dat = " & Val(txt_dae.Text)
                '대결
                If Col = 38 Then
                    If txt_dae.Text = 99999999 Then
                        sss = sss & "," & " apl_allsab = " & asab(3)
                        sss = sss & "," & " apl_allyn = 'Y'"
                        sss = sss & "," & " apl_alldat = " & Format(Now, "yyyymmdd")
                    End If
                End If
                sss = sss & " where apl_table = 'man_tooltesthd'"
                sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
                sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
                
                db.Execute sss, 64
                '
                '다음 결재자 쪽지 발송
                '
                If Col = 26 Or Col = 30 Or Col = 38 Then
                
                    sss = "insert into oth_automsg(ams_time,ams_sab,ams_rsab,ams_sendyn,ams_stime,ams_comment,ams_sprog)"
                    sss = sss & " values(sysdate" & ","
                    sss = sss & " 9999" & ","
                    If Col = 26 Then sss = sss & nextsab(2) & ","
                    If Col = 30 Then sss = sss & nextsab(3) & ","
                    If Col = 38 Then sss = sss & isab & ","
                    sss = sss & " 'N'" & ","
                    sss = sss & " 1" & ","
                    sss = sss & "'" & memo & "'" & ","
                    sss = sss & " 'TEST DATA 등록')"
                    '
                    db.Execute sss, 64
                
                ElseIf Col = 34 Then
                
                    For ii = 1 To 4
                        
                        If msab(ii) > 0 Then
                        
                            sss = "insert into oth_automsg(ams_time,ams_sab,ams_rsab,ams_sendyn,ams_stime,ams_comment,ams_sprog)"
                            sss = sss & " values(sysdate" & ","
                            sss = sss & " 9999" & ","
                            sss = sss & msab(ii) & ","
                            sss = sss & " 'N'" & ","
                            sss = sss & " 1" & ","
                            sss = sss & "'" & memo & "'" & ","
                            sss = sss & " 'TEST DATA 등록')"
                            '
                            db.Execute sss, 64
                        
                        End If
                        
                    Next ii
                
                End If
                '
                
                '결재도장 찍기====
                '결재자
                .CellType = CellTypeEdit: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter:
                If Col = 26 Then .Text = aname(1): .BackColor = RGB(215, 227, 188)
                If Col = 30 Then .Text = aname(2): .BackColor = RGB(215, 227, 188)
                If Col = 34 Then .Text = aname(3): .BackColor = RGB(215, 227, 188)
                If Col = 38 Then
                    .BackColor = RGB(255, 214, 255)
                    If txt_dae.Text = 99999999 Then
                        .Text = "대결"
                    Else
                        .Text = aname(4)
                    End If
                End If
                
                '결재일자
                .row = 4:
                If Col = 38 Then
                 
                    If txt_dae.Text = 99999999 Then
                        .Text = Format(Format(Now, "yyyymmdd"), "####/##/##")
                    Else
                        .Text = Format(txt_dae.Text, "####/##/##")
                    End If
                    
                Else
                
                    .Text = Format(Format(Now, "yyyymmdd"), "####/##/##")
                    
                End If
                '================
            
            '관련부서 결재
            ElseIf (row = 33 And Col = 34) Or (row = 33 And Col = 38) Or (row = 37 And Col = 34) Or (row = 37 And Col = 38) Then
                
                Dim mcnt As Integer '관련부서 결재자명수
                Dim myncnt As Integer '관련부서 결재완료수
                
                .Col = Col
                .row = row
                
                If Gsab <> .CellTag Then
                    
                    msg = "결재자만 결재할 수 있습니다."
                    Exit Sub
                
                Else
        
                    sss = "update oth_applist set"
                    If row = 33 And Col = 34 Then sss = sss & " apl_m1yn = 'Y',"
                    If row = 33 And Col = 34 Then sss = sss & " apl_m1dat = " & Format(Now, "yyyymmdd")
                    If row = 33 And Col = 38 Then sss = sss & " apl_m2yn = 'Y',"
                    If row = 33 And Col = 38 Then sss = sss & " apl_m2dat = " & Format(Now, "yyyymmdd")
                    If row = 37 And Col = 34 Then sss = sss & " apl_m3yn = 'Y',"
                    If row = 37 And Col = 34 Then sss = sss & " apl_m3dat = " & Format(Now, "yyyymmdd")
                    If row = 37 And Col = 38 Then sss = sss & " apl_m4yn = 'Y',"
                    If row = 37 And Col = 38 Then sss = sss & " apl_m4dat = " & Format(Now, "yyyymmdd")
                    sss = sss & " where apl_table = 'man_tooltesthd'"
                    sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
                    sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
                
                    db.Execute sss, 64
                    
                End If
                
                '결재도장 찍기 ====
                .CellType = CellTypeEdit: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188):
                If row = 33 And Col = 34 Then .Text = mname(1)
                If row = 33 And Col = 38 Then .Text = mname(2)
                If row = 37 And Col = 34 Then .Text = mname(3)
                If row = 37 And Col = 38 Then .Text = mname(4)
                
                If row = 33 And Col = 34 Then .row = 35: .Text = Format(Format(Now, "yyyymmdd"), "####/##/##")
                If row = 33 And Col = 38 Then .row = 35: .Text = Format(Format(Now, "yyyymmdd"), "####/##/##")
                If row = 37 And Col = 34 Then .row = 39: .Text = Format(Format(Now, "yyyymmdd"), "####/##/##")
                If row = 37 And Col = 38 Then .row = 39: .Text = Format(Format(Now, "yyyymmdd"), "####/##/##")
                '==================
                
                '관련부서 결재완료 확인 =====
                sss = "select apl_m1sab,apl_m2sab,apl_m3sab,apl_m4sab,"
                sss = sss & " apl_m1yn,apl_m2yn,apl_m3yn,apl_m4yn"
                sss = sss & " from oth_applist"
                sss = sss & " where apl_table = 'man_tooltesthd'"
                sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
                sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
                
                Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
                If Rs.RecordCount > 0 Then
                    
                    If Rs!apl_m1sab > 0 Then mcnt = mcnt + 1
                    If Rs!apl_m2sab > 0 Then mcnt = mcnt + 1
                    If Rs!apl_m3sab > 0 Then mcnt = mcnt + 1
                    If Rs!apl_m4sab > 0 Then mcnt = mcnt + 1
                    
                    If Rs!apl_m1yn = "Y" Then myncnt = myncnt + 1
                    If Rs!apl_m2yn = "Y" Then myncnt = myncnt + 1
                    If Rs!apl_m3yn = "Y" Then myncnt = myncnt + 1
                    If Rs!apl_m4yn = "Y" Then myncnt = myncnt + 1
                    
                End If
                '============================
                
                '결재완료 쪽지 발송
                If myncnt = mcnt Then
                    
                    sss = "insert into oth_automsg(ams_time,ams_sab,ams_rsab,ams_sendyn,ams_stime,ams_comment,ams_sprog)"
                    sss = sss & " values(sysdate" & ","
                    sss = sss & " 9999" & ","
                    sss = sss & isab & ","
                    sss = sss & " 'N'" & ","
                    sss = sss & " 1" & ","
                    sss = sss & "'" & memo & "'" & ","
                    sss = sss & " 'TEST DATA 등록')"
                    '
                    db.Execute sss, 64
                    
                    sss = "insert into oth_automsg(ams_time,ams_sab,ams_rsab,ams_sendyn,ams_stime,ams_comment,ams_sprog)"
                    sss = sss & " values(sysdate" & ","
                    sss = sss & " 9999" & ","
                    sss = sss & asab(3) & ","
                    sss = sss & " 'N'" & ","
                    sss = sss & " 1" & ","
                    sss = sss & "'" & memo & "'" & ","
                    sss = sss & " 'TEST DATA 등록')"
                    '
                    db.Execute sss, 64
                    
                End If
                
            End If
            
            Ws.CommitTrans
            
        End With
        
        msg = "결재가 완료되었습니다."
        
    End If
    
    txt_dat1.SetFocus
    
    Exit Sub

err_rtn:

    Ws.Rollback
    MsgBox (Err.Description)
    
End Sub



Private Sub sprd_print1_DblClick(ByVal Col As Long, ByVal row As Long)
    
    On Error GoTo err_rtn
    
    If Len(txt_dat1) < 8 Then
        msg = "등록번호를 확인하세요."
        Exit Sub
    End If
    
    If Len(txt_seq1) = 0 Then
        msg = "등록번호를 확인하세요."
        Exit Sub
    End If
    
    '생산부, 생산관리, 전무님 결재취소
    If row = 2 Or row = 3 Then
        
        '결재라인만 취소가능
        If Col < 22 Or Col > 41 Then
            Exit Sub
        End If
        
        '담당결재는 취소불가
        If Col = 22 Or Col = 23 Or Col = 24 Or Col = 25 Then
            Exit Sub
        End If
        
        'col 통합 작업===
        If Col = 26 Or Col = 27 Or Col = 28 Or Col = 29 Then
            Col = 26
        End If
    
        If Col = 30 Or Col = 31 Or Col = 32 Or Col = 33 Then
            Col = 30
        End If
    
        If Col = 34 Or Col = 35 Or Col = 36 Or Col = 37 Then
            Col = 34
        End If
    
        If Col = 38 Or Col = 39 Or Col = 40 Or Col = 41 Then
            Col = 38
        End If
        '================
        
        With sprd_print1
        
            .row = 2
            .Col = Col
            
            If Col = 38 And (row = 2 Or row = 3) And Gsab = asab(3) Then
            Else
                If Gsab <> .CellTag Then
                    msg = "결재자만 취소할 수 있습니다."
                    Exit Sub
                End If
            End If
            
            Ws.BeginTrans
        
            sss = "update oth_applist set"
            If Col = 26 Then sss = sss & " apl_1yn = 'N',"
            If Col = 26 Then sss = sss & " apl_1dat = 0"
            If Col = 30 Then sss = sss & " apl_2yn = 'N',"
            If Col = 30 Then sss = sss & " apl_2dat = 0"
            If Col = 34 Then sss = sss & " apl_3yn = 'N',"
            If Col = 34 Then sss = sss & " apl_3dat = 0"
            If Col = 38 Then sss = sss & " apl_4yn = 'N',"
            If Col = 38 Then sss = sss & " apl_4dat = 0,"
            If Col = 38 Then sss = sss & " apl_allyn = 'N',"
            If Col = 38 Then sss = sss & " apl_alldat = 0,"
            If Col = 38 Then sss = sss & " apl_allsab = 0"
            sss = sss & " where apl_table = 'man_tooltesthd'"
            sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
            sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
            
            db.Execute sss, 64
            
            Ws.CommitTrans
                
            .row = 2
            .Col = Col
            .Text = ""
            .CellType = CellTypeButton
            If Col = 38 Then
                .TypeButtonText = "사장님"
            Else
                If Col = 26 Then .TypeButtonText = aname(1)
                If Col = 30 Then .TypeButtonText = aname(2)
                If Col = 34 Then .TypeButtonText = aname(3)
            End If
    
            .Lock = False
    
            .row = 4
            .Text = ""
            
            msg = "결재가 취소되었습니다."
            
        End With
    
    '관련부서 결재 취소
    ElseIf row = 33 Or row = 34 Or row = 37 Or row = 38 Then
        
        '결재라인만 취소가능
        If Col < 34 Or Col > 41 Then
            Exit Sub
        End If
        
        'col 통합 작업===
        If Col = 34 Or Col = 35 Or Col = 36 Or Col = 37 Then
            Col = 34
        End If
        
        If Col = 38 Or Col = 39 Or Col = 40 Or Col = 41 Then
            Col = 38
        End If
        '================
        
        With sprd_print1
            
            If row = 33 Or row = 34 Then .row = 33
            If row = 37 Or row = 38 Then .row = 37
            
            .Col = Col
        
            If Gsab <> .CellTag Then
                msg = "결재자만 취소할 수 있습니다."
                Exit Sub
            End If
            
            Ws.BeginTrans
            
            sss = "update oth_applist set"
            
            '관련부서1(왼쪽 위)
            If Col = 34 And (row = 33 Or row = 34) Then sss = sss & " apl_m1yn = 'N',"
            If Col = 34 And (row = 33 Or row = 34) Then sss = sss & " apl_m1dat = 0"
            
            '관련부서2(오른쪽 위)
            If Col = 38 And (row = 33 Or row = 34) Then sss = sss & " apl_m2yn = 'N',"
            If Col = 38 And (row = 33 Or row = 34) Then sss = sss & " apl_m2dat = 0"
            
            '관련부서3(왼쪽 아래)
            If Col = 34 And (row = 37 Or row = 38) Then sss = sss & " apl_m3yn = 'N',"
            If Col = 34 And (row = 37 Or row = 38) Then sss = sss & " apl_m3dat = 0"
            
            '관련부서4(오른쪽 아래)
            If Col = 38 And (row = 37 Or row = 38) Then sss = sss & " apl_m4yn = 'N',"
            If Col = 38 And (row = 37 Or row = 38) Then sss = sss & " apl_m4dat = 0"
            
            sss = sss & " where apl_table = 'man_tooltesthd'"
            sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
            sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
            
            db.Execute sss, 64
        
            Ws.CommitTrans
        
            If row = 33 Or row = 34 Then
            
                .row = 33
                .Col = Col
                .CellType = CellTypeButton: .Lock = False
    
                If Col = 34 Then .TypeButtonText = mname(1)
                If Col = 38 Then .TypeButtonText = mname(2)
                
                .row = 35
                .Text = ""
            
            ElseIf row = 37 Or row = 38 Then
                
                .row = 37
                .Col = Col
                .CellType = CellTypeButton: .Lock = False
    
                If Col = 34 Then .TypeButtonText = mname(3)
                If Col = 38 Then .TypeButtonText = mname(4)
                
                .row = 39
                .Text = ""
                
            End If
                
            msg = "결재가 취소되었습니다."
        
        End With
        
    Else
        Exit Sub
    End If
    
    '이거 안하면 취소가 제대로안됨... 이유 모르겠음
    txt_dat1.SetFocus
            
    Exit Sub

err_rtn:
    Ws.Rollback
    MsgBox (Err.Description)
    
End Sub

Private Sub sprd_print2_ButtonClicked(ByVal Col As Long, ByVal row As Long, ByVal ButtonDown As Integer)
    
    On Error GoTo err_rtn
    
    Dim title       As String      '쪽지발송 시 사용할 title 변수
    
    If Len(txt_dat1) < 8 Then
        msg = "등록번호를 확인하세요."
        Exit Sub
    End If
    
    If Len(txt_seq1) = 0 Then
        msg = "등록번호를 확인하세요."
        Exit Sub
    End If
    
    '결재자 지정
    If chk_app.Value = 1 Then
        
        If (Col = 26 Or Col = 30 Or Col = 34) And row = 2 Then
            
            Call mkpoen05_app.index_Send(Col, row, 2)
            
            Load mkpoen05_app
            mkpoen05_app.Show 1
            
        End If
        
        If (Col = 34 Or Col = 38) And (row = 33 Or row = 37) Then
        
            Call mkpoen05_app.index_Send(Col, row, 2)
            
            Load mkpoen05_app
            mkpoen05_app.Show 1
            
        End If
    
    '결재
    Else
        
        With sprd_print2
            
            '결재쪽지 문자열 생성====
            .row = 4: .Col = 1: title = .Text
            
            '사장님 결재완료 시
            If row = 2 And Col = 38 Then
                memo = "TEST DATA 결재가 완료되었습니다." & vbCrLf & vbCrLf
                memo = memo & "등록번호: " & Val(txt_dat1) & "-" & Val(txt_seq1) & vbCrLf
                memo = memo & "담당자:" & iname & vbCrLf
                memo = memo & "TEST 제목: " & title & vbCrLf & vbCrLf
                memo = memo & "연결PG:Tol\mkpoen05"
            '관련부서 결재 완료시
            ElseIf (row = 33 And Col = 34) Or (row = 33 And Col = 38) Or (row = 37 And Col = 34) Or (row = 37 And Col = 38) Then
                memo = "TEST DATA 결재가 모두완료되었습니다. " & vbCrLf
                memo = memo & "사장님 결재를 진행해주세요!!" & vbCrLf & vbCrLf
                memo = memo & "등록번호: " & Val(txt_dat1) & "-" & Val(txt_seq1) & vbCrLf
                memo = memo & "담당자:" & iname & vbCrLf
                memo = memo & "TEST 제목: " & title & vbCrLf & vbCrLf
                memo = memo & "연결PG:Tol\mkpoen05"
            '그 외 결재시
            Else
                memo = "TEST DATA 가 등록되었습니다." & vbCrLf
                memo = memo & "확인후 본인 결재란에 결재 바랍니다." & vbCrLf & vbCrLf
                memo = memo & "등록번호: " & Val(txt_dat1) & "-" & Val(txt_seq1) & vbCrLf
                memo = memo & "담당자:" & iname & vbCrLf
                memo = memo & "TEST 제목: " & title & vbCrLf & vbCrLf
                memo = memo & "연결PG:Tol\mkpoen05"
            End If
            '========================
            
            Ws.BeginTrans
            
            '생산,기술연구소,전무님,사장님 결재
            If (row = 2 And Col = 26) Or (row = 2 And Col = 30) Or (row = 2 And Col = 34) Or (row = 2 And Col = 38) Then
            
                .Col = Col
                .row = 2
                
                '결재자 미지정시 결재 불가
                If .CellTag = 0 Then
                    msg = "결재자가 지정되지 않았습니다."
                    Exit Sub
                End If
                
                '전무님 승인결재 가능(대결 or 사장님 전자결재를 진행해야하기 때문)
                If Col = 38 And row = 2 And Gsab = asab(3) Then
                    
                    txt_dae.Text = ""
                    
                    Load mkpoen05_dae
                    mkpoen05_dae.Show 1
                    
                    If Val(txt_dae.Text) = 0 Then
                        Exit Sub
                    End If
                
                Else
                    If Gsab <> .CellTag Then
                        msg = "결재자만 결재할 수 있습니다."
                        Exit Sub
                    End If
                End If
                '
        
                sss = "update oth_applist set"
                If Col = 26 Then sss = sss & " apl_1yn = 'Y',"
                If Col = 26 Then sss = sss & " apl_1dat = " & Format(Now, "yyyymmdd")
                If Col = 30 Then sss = sss & " apl_2yn = 'Y',"
                If Col = 30 Then sss = sss & " apl_2dat = " & Format(Now, "yyyymmdd")
                If Col = 34 Then sss = sss & " apl_3yn = 'Y',"
                If Col = 34 Then sss = sss & " apl_3dat = " & Format(Now, "yyyymmdd")
                If Col = 38 Then sss = sss & " apl_4yn = 'Y',"
                If Col = 38 Then sss = sss & " apl_4dat = " & Val(txt_dae.Text)
                '대결
                If Col = 38 Then
                    If txt_dae.Text = 99999999 Then
                        sss = sss & "," & " apl_allsab = " & asab(3)
                        sss = sss & "," & " apl_allyn = 'Y'"
                        sss = sss & "," & " apl_alldat = " & Format(Now, "yyyymmdd")
                    End If
                End If
                sss = sss & " where apl_table = 'man_tooltesthd'"
                sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
                sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
                
                db.Execute sss, 64
                '
                '다음 결재자 쪽지 발송
                '
                If Col = 26 Or Col = 30 Or Col = 38 Then
                
                    sss = "insert into oth_automsg(ams_time,ams_sab,ams_rsab,ams_sendyn,ams_stime,ams_comment,ams_sprog)"
                    sss = sss & " values(sysdate" & ","
                    sss = sss & " 9999" & ","
                    If Col = 26 Then sss = sss & nextsab(2) & ","
                    If Col = 30 Then sss = sss & nextsab(3) & ","
                    If Col = 38 Then sss = sss & isab & ","
                    sss = sss & " 'N'" & ","
                    sss = sss & " 1" & ","
                    sss = sss & "'" & memo & "'" & ","
                    sss = sss & " 'TEST DATA 등록')"
                    '
                    db.Execute sss, 64
                
                ElseIf Col = 34 Then
                
                    For ii = 1 To 4
                        
                        If msab(ii) > 0 Then
                        
                            sss = "insert into oth_automsg(ams_time,ams_sab,ams_rsab,ams_sendyn,ams_stime,ams_comment,ams_sprog)"
                            sss = sss & " values(sysdate" & ","
                            sss = sss & " 9999" & ","
                            sss = sss & msab(ii) & ","
                            sss = sss & " 'N'" & ","
                            sss = sss & " 1" & ","
                            sss = sss & "'" & memo & "'" & ","
                            sss = sss & " 'TEST DATA 등록')"
                            '
                            db.Execute sss, 64
                        
                        End If
                        
                    Next ii
                
                End If
                '
                
                '결재도장 찍기====
                '결재자
                .CellType = CellTypeEdit: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter:
                If Col = 26 Then .Text = aname(1): .BackColor = RGB(215, 227, 188)
                If Col = 30 Then .Text = aname(2): .BackColor = RGB(215, 227, 188)
                If Col = 34 Then .Text = aname(3): .BackColor = RGB(215, 227, 188)
                If Col = 38 Then
                    .BackColor = RGB(255, 214, 255)
                    If txt_dae.Text = 99999999 Then
                        .Text = "대결"
                    Else
                        .Text = aname(4)
                    End If
                End If
                
                '결재일자
                .row = 4:
                If Col = 38 Then
                 
                    If txt_dae.Text = 99999999 Then
                        .Text = Format(Format(Now, "yyyymmdd"), "####/##/##")
                    Else
                        .Text = Format(txt_dae.Text, "####/##/##")
                    End If
                    
                Else
                
                    .Text = Format(Format(Now, "yyyymmdd"), "####/##/##")
                    
                End If
                '================
            
            '관련부서 결재
            ElseIf (row = 33 And Col = 34) Or (row = 33 And Col = 38) Or (row = 37 And Col = 34) Or (row = 37 And Col = 38) Then
                
                Dim mcnt As Integer '관련부서 결재자명수
                Dim myncnt As Integer '관련부서 결재완료수
                
                .Col = Col
                .row = row
                
                If Gsab <> .CellTag Then
                    
                    msg = "결재자만 결재할 수 있습니다."
                    Exit Sub
                
                Else
        
                    sss = "update oth_applist set"
                    If row = 33 And Col = 34 Then sss = sss & " apl_m1yn = 'Y',"
                    If row = 33 And Col = 34 Then sss = sss & " apl_m1dat = " & Format(Now, "yyyymmdd")
                    If row = 33 And Col = 38 Then sss = sss & " apl_m2yn = 'Y',"
                    If row = 33 And Col = 38 Then sss = sss & " apl_m2dat = " & Format(Now, "yyyymmdd")
                    If row = 37 And Col = 34 Then sss = sss & " apl_m3yn = 'Y',"
                    If row = 37 And Col = 34 Then sss = sss & " apl_m3dat = " & Format(Now, "yyyymmdd")
                    If row = 37 And Col = 38 Then sss = sss & " apl_m4yn = 'Y',"
                    If row = 37 And Col = 38 Then sss = sss & " apl_m4dat = " & Format(Now, "yyyymmdd")
                    sss = sss & " where apl_table = 'man_tooltesthd'"
                    sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
                    sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
                
                    db.Execute sss, 64
                    
                End If
                
                '결재도장 찍기 ====
                .CellType = CellTypeEdit: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188):
                If row = 33 And Col = 34 Then .Text = mname(1)
                If row = 33 And Col = 38 Then .Text = mname(2)
                If row = 37 And Col = 34 Then .Text = mname(3)
                If row = 37 And Col = 38 Then .Text = mname(4)
                
                If row = 33 And Col = 34 Then .row = 35: .Text = Format(Format(Now, "yyyymmdd"), "####/##/##")
                If row = 33 And Col = 38 Then .row = 35: .Text = Format(Format(Now, "yyyymmdd"), "####/##/##")
                If row = 37 And Col = 34 Then .row = 39: .Text = Format(Format(Now, "yyyymmdd"), "####/##/##")
                If row = 37 And Col = 38 Then .row = 39: .Text = Format(Format(Now, "yyyymmdd"), "####/##/##")
                '==================
                
                '관련부서 결재완료 확인 =====
                sss = "select apl_m1sab,apl_m2sab,apl_m3sab,apl_m4sab,"
                sss = sss & " apl_m1yn,apl_m2yn,apl_m3yn,apl_m4yn"
                sss = sss & " from oth_applist"
                sss = sss & " where apl_table = 'man_tooltesthd'"
                sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
                sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
                
                Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
                If Rs.RecordCount > 0 Then
                    
                    If Rs!apl_m1sab > 0 Then mcnt = mcnt + 1
                    If Rs!apl_m2sab > 0 Then mcnt = mcnt + 1
                    If Rs!apl_m3sab > 0 Then mcnt = mcnt + 1
                    If Rs!apl_m4sab > 0 Then mcnt = mcnt + 1
                    
                    If Rs!apl_m1yn = "Y" Then myncnt = myncnt + 1
                    If Rs!apl_m2yn = "Y" Then myncnt = myncnt + 1
                    If Rs!apl_m3yn = "Y" Then myncnt = myncnt + 1
                    If Rs!apl_m4yn = "Y" Then myncnt = myncnt + 1
                    
                End If
                '============================
                
                '결재완료 쪽지 발송
                If myncnt = mcnt Then
                    
                    sss = "insert into oth_automsg(ams_time,ams_sab,ams_rsab,ams_sendyn,ams_stime,ams_comment,ams_sprog)"
                    sss = sss & " values(sysdate" & ","
                    sss = sss & " 9999" & ","
                    sss = sss & isab & ","
                    sss = sss & " 'N'" & ","
                    sss = sss & " 1" & ","
                    sss = sss & "'" & memo & "'" & ","
                    sss = sss & " 'TEST DATA 등록')"
                    '
                    db.Execute sss, 64
                    
                    sss = "insert into oth_automsg(ams_time,ams_sab,ams_rsab,ams_sendyn,ams_stime,ams_comment,ams_sprog)"
                    sss = sss & " values(sysdate" & ","
                    sss = sss & " 9999" & ","
                    sss = sss & asab(3) & ","
                    sss = sss & " 'N'" & ","
                    sss = sss & " 1" & ","
                    sss = sss & "'" & memo & "'" & ","
                    sss = sss & " 'TEST DATA 등록')"
                    '
                    db.Execute sss, 64
                    
                End If
                
            End If
            
            Ws.CommitTrans
            
        End With
        
        msg = "결재가 완료되었습니다."
        
    End If
    
    txt_dat1.SetFocus
    
    Exit Sub

err_rtn:
    Ws.Rollback
    MsgBox (Err.Description)

End Sub

Private Sub sprd_print2_DblClick(ByVal Col As Long, ByVal row As Long)
        
    On Error GoTo err_rtn
    
    If Len(txt_dat1) < 8 Then
        msg = "등록번호를 확인하세요."
        Exit Sub
    End If
    
    If Len(txt_seq1) = 0 Then
        msg = "등록번호를 확인하세요."
        Exit Sub
    End If
    
        '생산부, 생산관리, 전무님 결재취소
    If row = 2 Or row = 3 Then
        
        '결재라인만 취소가능
        If Col < 22 Or Col > 41 Then
            Exit Sub
        End If
        
        '담당결재는 취소불가
        If Col = 22 Or Col = 23 Or Col = 24 Or Col = 25 Then
            Exit Sub
        End If
        
        'col 통합 작업===
        If Col = 26 Or Col = 27 Or Col = 28 Or Col = 29 Then
            Col = 26
        End If
    
        If Col = 30 Or Col = 31 Or Col = 32 Or Col = 33 Then
            Col = 30
        End If
    
        If Col = 34 Or Col = 35 Or Col = 36 Or Col = 37 Then
            Col = 34
        End If
    
        If Col = 38 Or Col = 39 Or Col = 40 Or Col = 41 Then
            Col = 38
        End If
        '================
        
        With sprd_print2
        
            .row = 2
            .Col = Col
            
            If Col = 38 And (row = 2 Or row = 3) And Gsab = asab(3) Then
            Else
                If Gsab <> .CellTag Then
                    msg = "결재자만 취소할 수 있습니다."
                    Exit Sub
                End If
            End If
            
            Ws.BeginTrans
        
            sss = "update oth_applist set"
            If Col = 26 Then sss = sss & " apl_1yn = 'N',"
            If Col = 26 Then sss = sss & " apl_1dat = 0"
            If Col = 30 Then sss = sss & " apl_2yn = 'N',"
            If Col = 30 Then sss = sss & " apl_2dat = 0"
            If Col = 34 Then sss = sss & " apl_3yn = 'N',"
            If Col = 34 Then sss = sss & " apl_3dat = 0"
            If Col = 38 Then sss = sss & " apl_4yn = 'N',"
            If Col = 38 Then sss = sss & " apl_4dat = 0,"
            If Col = 38 Then sss = sss & " apl_allyn = 'N',"
            If Col = 38 Then sss = sss & " apl_alldat = 0,"
            If Col = 38 Then sss = sss & " apl_allsab = 0"
            sss = sss & " where apl_table = 'man_tooltesthd'"
            sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
            sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
            
            db.Execute sss, 64
            
            Ws.CommitTrans
                
            .row = 2
            .Col = Col
            .Text = ""
            .CellType = CellTypeButton
            If Col = 38 Then
                .TypeButtonText = "사장님"
            Else
                If Col = 26 Then .TypeButtonText = aname(1)
                If Col = 30 Then .TypeButtonText = aname(2)
                If Col = 34 Then .TypeButtonText = aname(3)
            End If
    
            .Lock = False
    
            .row = 4
            .Text = ""
            
            msg = "결재가 취소되었습니다."
            
        End With
    
    '관련부서 결재 취소
    ElseIf row = 33 Or row = 34 Or row = 37 Or row = 38 Then
        
        '결재라인만 취소가능
        If Col < 34 Or Col > 41 Then
            Exit Sub
        End If
        
        'col 통합 작업===
        If Col = 34 Or Col = 35 Or Col = 36 Or Col = 37 Then
            Col = 34
        End If
        
        If Col = 38 Or Col = 39 Or Col = 40 Or Col = 41 Then
            Col = 38
        End If
        '================
        
        With sprd_print2
            
            If row = 33 Or row = 34 Then .row = 33
            If row = 37 Or row = 38 Then .row = 37
            
            .Col = Col
        
            If Gsab <> .CellTag Then
                msg = "결재자만 취소할 수 있습니다."
                Exit Sub
            End If
            
            Ws.BeginTrans
            
            sss = "update oth_applist set"
            
            '관련부서1(왼쪽 위)
            If Col = 34 And (row = 33 Or row = 34) Then sss = sss & " apl_m1yn = 'N',"
            If Col = 34 And (row = 33 Or row = 34) Then sss = sss & " apl_m1dat = 0"
            
            '관련부서2(오른쪽 위)
            If Col = 38 And (row = 33 Or row = 34) Then sss = sss & " apl_m2yn = 'N',"
            If Col = 38 And (row = 33 Or row = 34) Then sss = sss & " apl_m2dat = 0"
            
            '관련부서3(왼쪽 아래)
            If Col = 34 And (row = 37 Or row = 38) Then sss = sss & " apl_m3yn = 'N',"
            If Col = 34 And (row = 37 Or row = 38) Then sss = sss & " apl_m3dat = 0"
            
            '관련부서4(오른쪽 아래)
            If Col = 38 And (row = 37 Or row = 38) Then sss = sss & " apl_m4yn = 'N',"
            If Col = 38 And (row = 37 Or row = 38) Then sss = sss & " apl_m4dat = 0"
            
            sss = sss & " where apl_table = 'man_tooltesthd'"
            sss = sss & "   and apl_tdat = " & mkpoen05_print.txt_dat1
            sss = sss & "   and apl_tseq = " & mkpoen05_print.txt_seq1
            
            db.Execute sss, 64
        
            Ws.CommitTrans
        
            If row = 33 Or row = 34 Then
            
                .row = 33
                .Col = Col
                .CellType = CellTypeButton: .Lock = False
    
                If Col = 34 Then .TypeButtonText = mname(1)
                If Col = 38 Then .TypeButtonText = mname(2)
                
                .row = 35
                .Text = ""
            
            ElseIf row = 37 Or row = 38 Then
                
                .row = 37
                .Col = Col
                .CellType = CellTypeButton: .Lock = False
    
                If Col = 34 Then .TypeButtonText = mname(3)
                If Col = 38 Then .TypeButtonText = mname(4)
                
                .row = 39
                .Text = ""
                
            End If
                
            msg = "결재가 취소되었습니다."
        
        End With
        
    Else
        Exit Sub
    End If
    
    txt_dat1.SetFocus
            
    Exit Sub

err_rtn:
    Ws.Rollback
    MsgBox (Err.Description)
    
End Sub

Private Function nextsab(Index As Integer)
        
    '다음 결재자 산출
        
    Dim i As Integer
    
    For i = Index To 3
                    
        If asab(i) > 0 Then
            nextsab = asab(i)
            Exit For
        End If
                    
    Next i

End Function

Private Sub SSCommand1_Click()
    MsgBox asab(3)
End Sub
