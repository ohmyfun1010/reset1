VERSION 5.00
Object = "{3B930683-5AF1-4F07-9CE8-CA8063E1F3DD}#2.0#0"; "XButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.MDIForm mkpoen05MDI 
   BackColor       =   &H8000000C&
   Caption         =   "TOOL TEST DATA (mkpoen05)"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16740
   LinkTopic       =   "MDImain"
   LockControls    =   -1  'True
   ScrollBars      =   0   'False
   Begin Threed.SSPanel spl_left 
      Align           =   3  '왼쪽 맞춤
      Height          =   9765
      Left            =   0
      TabIndex        =   4
      Top             =   465
      Width           =   2745
      _Version        =   65536
      _ExtentX        =   4842
      _ExtentY        =   17224
      _StockProps     =   15
      Caption         =   "SSPanel1"
      ForeColor       =   -2147483630
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin XLibrary_XGroupBox.XGroupBox Xgb_menu 
         Height          =   9750
         Left            =   0
         Top             =   0
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   17198
         BackColor       =   16777215
         BorderColor     =   10526880
         BorderRoundNum  =   0
         BorderStyle     =   1
         TextColor       =   16777215
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "                     M E N U"
         TextPosition    =   0
         TextCustomMargin=   4
         GroupBoxStyle   =   1
         TextBarColor1   =   12757903
         TextBarStyle    =   3
         TextBarColor2   =   11767328
         TextBarSymbol   =   0   'False
         TextBarSymbolColor=   16777215
         TextBarHeightMargin=   10
         MouseCursor     =   0
         TextBarMouseCursor=   0
         IconandTextMargin=   4
         BodyColor       =   16777215
         Enabled         =   -1  'True
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   2115
            Top             =   9000
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mkpoen05MDI.frx":0000
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mkpoen05MDI.frx":2360
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mkpoen05MDI.frx":455A
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView menu 
            Height          =   6615
            Left            =   90
            TabIndex        =   5
            Top             =   450
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   11668
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   706
            LabelEdit       =   1
            Style           =   7
            HotTracking     =   -1  'True
            ImageList       =   "ImageList1"
            Appearance      =   0
         End
         Begin VB.Label lbl_logintime 
            BackStyle       =   0  '투명
            Caption         =   "login time"
            ForeColor       =   &H008080FF&
            Height          =   195
            Left            =   90
            TabIndex        =   12
            ToolTipText     =   "권한Level에 따라 메뉴가 출력됩니다!"
            Top             =   9450
            Width           =   2625
         End
         Begin VB.Label lbl_loginlv 
            BackStyle       =   0  '투명
            Caption         =   "LV"
            ForeColor       =   &H008080FF&
            Height          =   195
            Left            =   90
            TabIndex        =   11
            ToolTipText     =   "권한Level에 따라 메뉴가 출력됩니다!"
            Top             =   9225
            Width           =   1140
         End
         Begin VB.Image img_expand 
            Height          =   195
            Left            =   90
            Picture         =   "mkpoen05MDI.frx":68A0
            Top             =   90
            Width           =   210
         End
         Begin VB.Image img_reduce 
            Height          =   195
            Left            =   2475
            Picture         =   "mkpoen05MDI.frx":8AC2
            Top             =   90
            Width           =   210
         End
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  '위 맞춤
      Height          =   465
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   16740
      _Version        =   65536
      _ExtentX        =   29527
      _ExtentY        =   820
      _StockProps     =   15
      ForeColor       =   14737632
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Begin VB.TextBox txt_pwd 
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   3  '사용 못함
         Left            =   11340
         TabIndex        =   0
         Text            =   "PASSWORD"
         Top             =   90
         Width           =   1590
      End
      Begin VB.TextBox txt_sok 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   270
         IMEMode         =   3  '사용 못함
         Left            =   10755
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   90
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txt_sab 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   270
         IMEMode         =   3  '사용 못함
         Left            =   10125
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   90
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   270
         IMEMode         =   3  '사용 못함
         Left            =   9405
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   90
         Visible         =   0   'False
         Width           =   690
      End
      Begin XLibrary_XButton.XButton XBtn_login 
         Height          =   345
         Left            =   13005
         TabIndex        =   1
         Top             =   60
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         BackColor1      =   12757903
         BackColor2      =   16777215
         BackColorEx     =   11767328
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
         Text            =   "로그인"
         TextWidthPos    =   2
         TextHeightPos   =   2
         TextWidthMargin =   5
         TextHeightMargin=   5
         TextColor       =   16777215
         IconPosition    =   2
         IconAndTextMargin=   0
         IconMaskColor   =   13828096
         MouseOverColor2 =   65535
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
         ToolTipBodyText =   "XBUTTON 2"
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
      Begin XLibrary_XButton.XButton XBtn_logout 
         Height          =   345
         Left            =   14445
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   60
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         BackColor1      =   12757903
         BackColor2      =   16777215
         BackColorEx     =   11767328
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
         Text            =   "로그아웃"
         TextWidthPos    =   2
         TextHeightPos   =   2
         TextWidthMargin =   5
         TextHeightMargin=   5
         TextColor       =   16777215
         IconPosition    =   2
         IconAndTextMargin=   0
         IconMaskColor   =   13828096
         MouseOverColor2 =   65535
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
         ToolTipBodyText =   "XBUTTON 2"
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
      Begin XLibrary_XButton.XButton XBtn_exit 
         Height          =   345
         Left            =   15885
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   60
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   609
         BackColor1      =   33023
         BackColor2      =   16777215
         BackColorEx     =   255
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
         Text            =   "종료"
         TextWidthPos    =   2
         TextHeightPos   =   2
         TextWidthMargin =   5
         TextHeightMargin=   5
         TextColor       =   16777215
         IconPosition    =   2
         IconAndTextMargin=   0
         IconMaskColor   =   13828096
         MouseOverColor2 =   65535
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
         ToolTipBodyText =   "XBUTTON 2"
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
      Begin VB.Image Image1 
         Height          =   195
         Left            =   2745
         Picture         =   "mkpoen05MDI.frx":ACF7
         Top             =   135
         Width           =   210
      End
      Begin VB.Image Image2 
         Height          =   180
         Left            =   90
         Picture         =   "mkpoen05MDI.frx":CF26
         Top             =   135
         Width           =   795
      End
      Begin VB.Label msg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "메시지를 출력합니다!"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3060
         TabIndex        =   7
         Top             =   150
         Width           =   8205
      End
   End
End
Attribute VB_Name = "mkpoen05MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim loginlvl         As Integer
Dim sdat As String
Dim edat As String

Private Sub MDIForm_Load()
    
    Dim RetVal

    Set Ws = DBEngine.Workspaces(0)
    Set db = OpenDatabase("hd", False, False, "odbc;dsn=hd;uid=bnkjdb;pwd=dshhjy")
    Db_Open = 9
    
    If Len(Command) > 0 Then
        txt_pwd = Command
        txt_pwd.PasswordChar = "*"
    End If
    
    
    mkpoen05M.Left = 0
    mkpoen05M.Top = 0
    mkpoen05M.ZOrder 0
    mkpoen05M.Show
    
    XBtn_logout.Enabled = False
    
    menu.Enabled = False
    
    mkpoen05MDI.Top = 0
    mkpoen05MDI.Left = 0
    
End Sub

Private Sub txt_pwd_Click()
    txt_pwd = ""
    txt_pwd.PasswordChar = "*"
End Sub

Private Sub txt_pwd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub img_expand_Click()
    Xgb_menu.Width = 2760
    spl_left.Width = 2760
    menu.Visible = True
End Sub

Private Sub img_reduce_Click()
    Xgb_menu.Width = 400
    spl_left.Width = 400
    menu.Visible = False
End Sub

Private Sub XBtn_exit_Click()
    Call NetWork_delete(N_Driver)
    Unload Me
    End
End Sub

Private Sub cmb_exit_Click()
    Call NetWork_delete(N_Driver)
    Unload Me
    End
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call NetWork_delete(N_Driver)
    Unload Me
    End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Call NetWork_delete(N_Driver)
    Unload Me
    End
End Sub

Private Sub menu_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim frm As Form
    
    '1) 페이지1
    If Node = "1.TEST DATA 관리" Then
        mkpoen05A.ZOrder 0
        mkpoen05A.Left = 0
        mkpoen05A.Top = 0
        mkpoen05A.Show
        
        If Len(mkpoen05A.txt_sdat2) < 1 Then
            mkpoen05A.txt_sdat2 = sdat
            mkpoen05A.txt_edat2 = edat
        End If
        
        msg_display ("[TEST DATA 관리 이동]")
        DoEvents
    
    ElseIf Node = "2.TEST DATA 조회" Then
        mkpoen05B.ZOrder 0
        mkpoen05B.Left = 0
        mkpoen05B.Top = 0
        mkpoen05B.Show
        
        If Len(mkpoen05B.txt_sdat1) < 1 Then
            mkpoen05B.txt_sdat1 = sdat
            mkpoen05B.txt_edat1 = edat
        
            mkpoen05B.txt_sdat2 = sdat
            mkpoen05B.txt_edat2 = edat
        End If
        
        msg_display ("[TEST DATA 조회 이동]")
        DoEvents
        
    ElseIf Node = "3.공구 테스트관리" Then
        mkpoen05C.ZOrder 0
        mkpoen05C.Left = 0
        mkpoen05C.Top = 0
        mkpoen05C.Show
        
        msg_display ("[공구 테스트관리 이동]")
        DoEvents
    
    End If
    
End Sub

Private Sub XBtn_login_Click()
        
    On Error GoTo err_rtn
        
    txt_pwd = UCase(Trim(txt_pwd))
        
    If Len(Trim(txt_pwd)) < 1 Then
        msg = "비밀번호를 입력하세요!"
        txt_pwd.SetFocus
        Exit Sub
    End If
        
    sss = "      select sin_sok,sin_sab, sin_name, sin_sok, sysdate, "
    sss = sss & "       to_char(add_months(sysdate, -60), 'YYYYMM') || '01' Day3ma,"
    sss = sss & "       to_char(sysdate,'yyyymmdd') day"
    sss = sss & "  from peo_sinbun "
    sss = sss & " where sin_pwd = '" & UCase(txt_pwd) & "'"
    sss = sss & "   and sin_taedt < 1 "
    '
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Rs.RecordCount < 1 Then
        msg = "비밀번호를 다시 확인하세요!"
        txt_sab = ""
        txt_sok = ""
        txt_name = ""
        txt_pwd = ""
        txt_pwd.SetFocus
        Exit Sub
    End If
        
    txt_sab = Rs!sin_sab: Gsab = Rs!sin_sab
    txt_sok = Rs!sin_sok: Gsok = Rs!sin_sok
    txt_name = Rs!sin_name: GName = Rs!sin_name
    
    sdat = Rs!day3ma
    edat = Rs!Day
    
    '-------------------------------------------------------------------------------------------------------------------------------------------
    ' 예:                 ==> Lock 만 체크 (사번=0, 부서='', Version='')
    '     PG_Version_Check('mrsoen16', 0, '', '1.0.1')          ==> Lock 및 Version 만체크 (사변=0, 부서='')
    '     PG_Version_Check('mrsoen16', 487, 'B100', '')         ==> Lock 및 사번/부서로 권한체크 ( Version='')
    '     PG_Version_Check('mrsoen16', 487, 'B100', '1.0.1')    ==> Lock, Version 및 사번/부서로 권한체크  (사번 or 부서코드 하나만 넣어도 가능)
    '
    ' << Return 값 >>
    ' 1) 테이블에 pg_name 이 없으면 => "권한관리에 미등록된 프로그램 입니다. 사용불가."
    ' 2) Lock 걸려있을때            => "이 프로그램은 Lock 걸려있습니다. 사용불가"
    ' 3) Version 불일치             => "이 프로그램은 Version 이 낮습니다. HDRUN 하세요."
    ' 4) 권한이 없으면              => "0"
    ' 5) 사번 및 부서 권한이 있으면 => "1~10" 가지 해당 권한 을 문자열로 돌려줌. (사번 및 부서코드 없이 Lock 및 version 체크시는 "1" 을 return 함)
    ' 6) 오류 및 조회 불가          => "이프로그램은 사용불가 합니다."
    
    '프로그램 Version 체크
    
    Dim version As String
    Dim result  As String
    version = app.Major & "." & app.Minor & "." & app.Revision
       
    result = PG_Version_Check(app.EXEName, Val(Gsab), Gsok, version)
    '
    If IsNumeric(result) = False Then
       MsgBox result
       Exit Sub
    Else
       '
       Job_Level = result
       '
           'If Val(Job_Level) < 1 Then
           '   MsgBox ("프로그램 사용 권한이 없습니다.")
           '   Exit Sub
           'End If
       '
    End If
    '-------------------------------------------------------------------------------------------------------------------------------------
    
    lbl_logintime.Caption = Rs!sysdate
    
    Rs.Close
    
    '컨트롤 초기화
    msg = txt_name.Text & "님이 로그인 되었습니다!"
    XBtn_login.Enabled = False
    XBtn_logout.Enabled = True
    txt_pwd.Enabled = False
    menu.Enabled = True
    
    lbl_loginlv.Caption = "LV" & " " & Job_Level
    '메뉴초기화
    Call init_Menu
    
    '네트웍드라이버 연결
   ' If Len(N_Driver) < 1 Then
   '     Call NetWork_connect("공구테스트_DATA")
   ' End If
    '
    Exit Sub
    
err_rtn:
    msg = Err.Description
End Sub

Private Sub init_Menu()
   
    menu.Nodes.Add , , "mnu001", "메뉴", 3
    menu.Nodes.Add "mnu001", tvwChild, "submmnu101", "1.TEST DATA 관리", 2, 1
    menu.Nodes.Add "mnu001", tvwChild, "submmnu202", "2.TEST DATA 조회", 2, 1
    menu.Nodes.Add "mnu001", tvwChild, "submmnu303", "3.공구 테스트관리", 2, 1
    
    Call ExpandAllNodes(menu)
    
End Sub

Private Sub XBtn_logout_Click()
    
    msg = txt_name.Text & "님이 로그아웃 되었습니다."
    '컨트롤 초기화
    txt_sab = "": Gsab = 0
    txt_name = "": GName = ""
    txt_sok = "": Gsok = ""
    lbl_loginlv = "LV"
    lbl_logintime = "login time"
    XBtn_login.Enabled = True
    txt_pwd.Enabled = True
    txt_pwd.PasswordChar = "": txt_pwd = "PASSWORD"
    XBtn_logout.Enabled = False
    
    '모든 노드 제거
    menu.Nodes.Clear

    '메인화면 출력
    mkpoen05M.Left = 0
    mkpoen05M.Top = 0
    mkpoen05M.ZOrder 0
    mkpoen05M.Show
    
    Call NetWork_delete(N_Driver)
    N_Driver = ""
    
End Sub

'모든 노드 열기
Public Sub ExpandAllNodes(xtree As MSComctlLib.TreeView)
On Error Resume Next
    Dim xnode As MSComctlLib.Node
    
    For Each xnode In xtree.Nodes
        If (xnode.Children > 0) And Not (xnode.Expanded) Then
            xnode.Expanded = True
        End If
    Next xnode
    xtree.SelectedItem.EnsureVisible
End Sub

'모든 노드 닫기
Public Sub CollapseAllNodes(xtree As MSComctlLib.TreeView)
    Dim xnode As MSComctlLib.Node
    
    For Each xnode In xtree.Nodes
        If (xnode.Children > 0) And (xnode.Expanded) And (xnode.Index <> 1) Then
            xnode.Expanded = False
        End If
    Next xnode
End Sub

Public Sub msg_display(mass)
   
   Dim jj As Integer
   Dim msg_len As Integer
   Dim pausetime As Single
   Dim start As Single
   
   msg_len = Len(Trim(mass))
   mkpoen05MDI.msg.Caption = ""
   Beep
   
   For jj = 1 To msg_len
       mkpoen05MDI.msg.Caption = Space(msg_len - jj + 2) & LTrim(mkpoen05MDI.msg.Caption)
       mkpoen05MDI.msg.Caption = mkpoen05MDI.msg.Caption & Mid(mass, jj, 1)
  '
       pausetime = 0.01    ' 기간을 지정합니다.
       start = Timer       ' 시작 시간을 지정합니다.
       Do While Timer < start + pausetime
          DoEvents         ' 다른 프로시저로 넘깁니다.
       Loop
   Next

End Sub


