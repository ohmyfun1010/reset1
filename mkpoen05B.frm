VERSION 5.00
Object = "{3B930683-5AF1-4F07-9CE8-CA8063E1F3DD}#2.0#0"; "XButton.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form mkpoen05B 
   BackColor       =   &H80000018&
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   9810
   ClientLeft      =   2760
   ClientTop       =   2445
   ClientWidth     =   13965
   FillColor       =   &H00C0E0FF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel3 
      Height          =   9825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      _Version        =   65536
      _ExtentX        =   24791
      _ExtentY        =   17330
      _StockProps     =   15
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin TabDlg.SSTab SSTab1 
         Height          =   9780
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   17251
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "1. TEST DATA ��ȸ"
         TabPicture(0)   =   "mkpoen05B.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "sprd1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "2. TEST DATA �󼼳��� ��ȸ"
         TabPicture(1)   =   "mkpoen05B.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txt_edat2"
         Tab(1).Control(1)=   "txt_sdat2"
         Tab(1).Control(2)=   "sprd2"
         Tab(1).Control(3)=   "Frame1"
         Tab(1).Control(4)=   "btn_view2"
         Tab(1).Control(5)=   "SSPanel10"
         Tab(1).Control(6)=   "Label16"
         Tab(1).ControlCount=   7
         Begin VB.Frame Frame3 
            Height          =   855
            Left            =   120
            TabIndex        =   20
            Top             =   495
            Width           =   13710
            Begin VB.TextBox txt_edat1 
               Appearance      =   0  '���
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2415
               MaxLength       =   8
               TabIndex        =   2
               Top             =   300
               Width           =   1050
            End
            Begin VB.TextBox txt_sdat1 
               Appearance      =   0  '���
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1200
               MaxLength       =   8
               TabIndex        =   1
               Top             =   300
               Width           =   1050
            End
            Begin VB.Frame Frame5 
               Height          =   450
               Left            =   4620
               TabIndex        =   21
               Top             =   210
               Width           =   2250
               Begin Threed.SSOption opt_X 
                  Height          =   195
                  Left            =   90
                  TabIndex        =   3
                  TabStop         =   0   'False
                  Top             =   195
                  Width           =   600
                  _Version        =   65536
                  _ExtentX        =   1058
                  _ExtentY        =   344
                  _StockProps     =   78
                  Caption         =   "��ü"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   -1  'True
               End
               Begin Threed.SSOption opt_OK 
                  Height          =   195
                  Left            =   810
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   195
                  Width           =   645
                  _Version        =   65536
                  _ExtentX        =   1138
                  _ExtentY        =   344
                  _StockProps     =   78
                  Caption         =   "O.K"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSOption opt_NG 
                  Height          =   195
                  Left            =   1530
                  TabIndex        =   5
                  TabStop         =   0   'False
                  Top             =   195
                  Width           =   615
                  _Version        =   65536
                  _ExtentX        =   1085
                  _ExtentY        =   344
                  _StockProps     =   78
                  Caption         =   "N.G"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
            Begin VB.TextBox txt_test1 
               Appearance      =   0  '���
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   8010
               MaxLength       =   8
               TabIndex        =   6
               Top             =   300
               Width           =   1110
            End
            Begin XLibrary_XButton.XButton btn_view1 
               Height          =   375
               Left            =   9210
               TabIndex        =   7
               Top             =   270
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   661
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
               Text            =   "Ȯ ��"
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
               ToolTipBodyText =   "��ȸ"
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
            Begin Threed.SSPanel SSPanel1 
               Height          =   330
               Left            =   120
               TabIndex        =   22
               Top             =   300
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   582
               _StockProps     =   15
               Caption         =   "�������"
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSPanel SSPanel11 
               Height          =   330
               Left            =   3540
               TabIndex        =   23
               Top             =   300
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   582
               _StockProps     =   15
               Caption         =   "�������"
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSPanel SSPanel14 
               Height          =   330
               Left            =   6930
               TabIndex        =   24
               Top             =   300
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   582
               _StockProps     =   15
               Caption         =   "TEST NO."
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label1 
               Caption         =   "-"
               Height          =   285
               Left            =   2280
               TabIndex        =   25
               Top             =   360
               Width           =   150
            End
         End
         Begin VB.TextBox txt_edat2 
            Appearance      =   0  '���
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -73680
            MaxLength       =   8
            TabIndex        =   10
            Top             =   840
            Width           =   1050
         End
         Begin VB.TextBox txt_sdat2 
            Appearance      =   0  '���
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -74930
            MaxLength       =   8
            TabIndex        =   9
            Top             =   840
            Width           =   1050
         End
         Begin FPSpreadADO.fpSpread sprd2 
            Height          =   8265
            Left            =   -74910
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1290
            Width           =   3615
            _Version        =   458752
            _ExtentX        =   6376
            _ExtentY        =   14579
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   5
            SpreadDesigner  =   "mkpoen05B.frx":0038
         End
         Begin VB.Frame Frame1 
            Height          =   9345
            Left            =   -71220
            TabIndex        =   17
            Top             =   330
            Width           =   10125
            Begin VB.TextBox txt_dat2 
               Appearance      =   0  '���
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8610
               MaxLength       =   8
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   210
               Visible         =   0   'False
               Width           =   930
            End
            Begin VB.TextBox txt_seq2 
               Appearance      =   0  '���
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9600
               MaxLength       =   8
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   210
               Visible         =   0   'False
               Width           =   420
            End
            Begin FPSpreadADO.fpSpread sprd_print2 
               Height          =   9045
               Left            =   90
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   180
               Width           =   9930
               _Version        =   458752
               _ExtentX        =   17515
               _ExtentY        =   15954
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
               ScrollBars      =   0
               SpreadDesigner  =   "mkpoen05B.frx":06A1
            End
         End
         Begin XLibrary_XButton.XButton btn_view2 
            Height          =   700
            Left            =   -72540
            TabIndex        =   11
            Top             =   480
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   1244
            BackColor1      =   12632256
            BackColor2      =   16777215
            BackColorEx     =   16761024
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
            Text            =   "��ȸ"
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
            ToolTipBodyText =   "��ȸ"
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   330
            Left            =   -74940
            TabIndex        =   18
            Top             =   480
            Width           =   2310
            _Version        =   65536
            _ExtentX        =   4075
            _ExtentY        =   582
            _StockProps     =   15
            Caption         =   "�������"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FPSpreadADO.fpSpread sprd1 
            Height          =   7995
            Left            =   120
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1650
            Width           =   13710
            _Version        =   458752
            _ExtentX        =   24183
            _ExtentY        =   14102
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   15
            MaxRows         =   22
            SpreadDesigner  =   "mkpoen05B.frx":2DF4
         End
         Begin VB.Label Label6 
            Caption         =   "�� ��ưŬ��: ���γ��� Ȯ��"
            Height          =   255
            Left            =   420
            TabIndex        =   26
            Top             =   1410
            Width           =   2475
         End
         Begin VB.Label Label16 
            Caption         =   "-"
            Height          =   285
            Left            =   -73830
            TabIndex        =   19
            Top             =   900
            Width           =   150
         End
      End
   End
End
Attribute VB_Name = "mkpoen05B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '
    Const SERVER_PATH As String = "/�����׽�Ʈ_DATA/"
    '
    Dim ii          As Double
    Dim cnt         As Double

'===================================
'TAB1. TEST DATA ��ȸ
'===================================
Private Sub btn_view1_Click()
    '
    On Error GoTo err_rtn
    '
    Dim chk_ttdno As String
    
    If Len(txt_sdat1) <> 8 Or IsNumeric(txt_sdat1) = False Then
        Call msg_display("��� �������ڸ� Ȯ���ϼ���")
        txt_sdat1.SetFocus
        Exit Sub
    End If
    
    If Len(txt_edat1) <> 8 Or IsNumeric(txt_edat1) = False Then
        Call msg_display("��� �������ڸ� Ȯ���ϼ���")
        txt_edat1.SetFocus
        Exit Sub
    End If

    sss = "       select *"
    sss = sss & "   from man_tooltesthd, man_tooltestds"
    sss = sss & "  where tth_dat = ttd_dat"
    sss = sss & "    and tth_seq = ttd_seq"
    If Len(txt_sdat1) = 8 And Len(txt_edat1) = 8 Then
        sss = sss & "   and tth_dat between " & txt_sdat1 & " and " & txt_edat1
    End If
                
    If Len(txt_test1) > 0 Then sss = sss & " and tth_testno = '" & txt_test1 & "'"
    
    If opt_OK.Value = True Or opt_NG.Value = True Then
        sss = sss & "    and tth_dat || tth_seq in(select ttd_dat || ttd_seq "
        sss = sss & "                                from man_tooltestds"
        If opt_OK.Value = True Then sss = sss & "   where ttd_result = 'OK')"
        If opt_NG.Value = True Then sss = sss & "   where ttd_result = 'NG')"
    End If
                                    
    'sss = sss & "  order by tth_dat, tth_seq, ttd_lno"
    sss = sss & "  order by tth_testno, tth_dat, tth_seq, ttd_lno"
    
    sprd1.MaxRows = 0: cnt = 1

    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Rs.RecordCount < 1 Then
        Rs.Close
        Call msg_display("��ϵ� ������ �����ϴ�!")
        Exit Sub
    End If

    sprd1.MaxRows = 0: cnt = 0
                            
    Do While Not Rs.EOF
                  
        cnt = cnt + 1: sprd1.MaxRows = cnt: sprd1.row = cnt
                
        If cnt > 1 Then
            If chk_ttdno <> Rs!tth_dat & Rs!tth_seq Then
                cnt = cnt + 1: sprd1.MaxRows = cnt: sprd1.row = cnt
                
                sprd1.AddCellSpan 1, cnt - 1, 15, 1      'ǰ�� ����
                sprd1.RowHeight(cnt - 1) = 12
                sprd1.row = cnt - 1: sprd1.Col = 1: sprd1.BackColor = &H80000004
                sprd1.row = cnt
                chk_ttdno = Rs!tth_dat & Rs!tth_seq
                
            End If
        Else
            chk_ttdno = Rs!tth_dat & Rs!tth_seq
        End If
        
        If Rs!ttd_lno = 1 Then
                                 
            'If Not IsNull(Rs!tth_dat) Then sprd1.Col = 1: sprd1.Text = Rs!tth_dat & "-" & Format(Rs!tth_seq, "000")
            
            sprd1.Col = 1
            sprd1.CellType = CellTypeButton: sprd1.TypeButtonColor = &H80000018
            sprd1.TypeButtonText = Rs!tth_dat & "-" & Format(Rs!tth_seq, "000")
            
            If Not IsNull(Rs!ttd_lno) Then sprd1.Col = 2: sprd1.Text = Rs!ttd_lno
            If Not IsNull(Rs!tth_testno) Then sprd1.Col = 3: sprd1.Text = Rs!tth_testno
            If Not IsNull(Rs!tth_tdat) Then sprd1.Col = 4: sprd1.Text = Rs!tth_tdat
            If Not IsNull(Rs!tth_title) Then sprd1.Col = 5: sprd1.Text = Rs!tth_title
            If Not IsNull(Rs!tth_tmcd) Then sprd1.Col = 6: sprd1.Text = Rs!tth_tmcd
            If Not IsNull(Rs!tth_tlot) Then sprd1.Col = 7: sprd1.Text = Rs!tth_tlot
                             
        End If
        
        If Rs!ttd_gbn = 1 Then
            sprd1.Col = 8: sprd1.Text = "����"
        ElseIf Rs!ttd_gbn = 2 Then
            sprd1.Col = 8: sprd1.Text = "�׽�Ʈ"
        End If
            
        If Not IsNull(Rs!ttd_maker) Then sprd1.Col = 9: sprd1.Text = Rs!ttd_maker
        If Not IsNull(Rs!ttd_tipstd) Then sprd1.Col = 10: sprd1.Text = Rs!ttd_tipstd
        If Not IsNull(Rs!ttd_tipjil) Then sprd1.Col = 11: sprd1.Text = Rs!ttd_tipjil
        If Not IsNull(Rs!ttd_dan) Then sprd1.Col = 12: sprd1.Text = Rs!ttd_dan
        If Not IsNull(Rs!ttd_tldn) Then sprd1.Col = 13: sprd1.Text = Rs!ttd_tldn
        If Not IsNull(Rs!ttd_chdn) Then sprd1.Col = 14: sprd1.Text = Rs!ttd_chdn
        If Not IsNull(Rs!ttd_result) Then
            sprd1.Col = 15: sprd1.Text = Rs!ttd_result
            If Rs!ttd_result = "OK" Then
                sprd1.ForeColor = &HFF0000
            Else
                sprd1.ForeColor = &HFF&
            End If
        End If
        Rs.MoveNext
        
    Loop
    
    Rs.Close
    
    Exit Sub

err_rtn:

    MsgBox (Err.Description)
    
End Sub

Private Sub sprd1_ButtonClicked(ByVal Col As Long, ByVal row As Long, ByVal ButtonDown As Integer)

On Error GoTo err_rtn

    sprd1.row = row: sprd1.Col = Col
    
    If sprd1.CellType = CellTypeButton Then
    
        txt_dat2 = Left(sprd1.TypeButtonText, 8)
        txt_seq2 = Right(sprd1.TypeButtonText, 3)
        
        SSTab1.Tab = 1
        
        Call print2
        
    End If

Exit Sub

err_rtn:

    MsgBox (Err.Description)
    
End Sub

Private Sub txt_test1_LostFocus()
    txt_test1.Text = UCase(Trim(txt_test1))
End Sub

'===================================
'TAB2. TEST DATA ��ȸ
'===================================
Private Sub btn_view2_Click()

On Error GoTo err_rtn

    Dim chk_ttdno As String
    
    If Len(txt_sdat2) <> 8 Or IsNumeric(txt_sdat2) = False Then
        Call msg_display("��� �������ڸ� Ȯ���ϼ���")
        txt_sdat2.SetFocus
        Exit Sub
    End If
    
    If Len(txt_edat2) <> 8 Or IsNumeric(txt_edat2) = False Then
        Call msg_display("��� �������ڸ� Ȯ���ϼ���")
        txt_edat2.SetFocus
        Exit Sub
    End If
    
    sss = "       select tth_testno, tth_dat || lpad(tth_seq,3,'0') tth_no, tth_title,  max(ttd_result) ttd_result  "
    sss = sss & "   from man_tooltesthd, man_tooltestds"
    sss = sss & "  where tth_dat = ttd_dat"
    sss = sss & "    and tth_seq = ttd_seq"
    If Len(txt_sdat2) = 8 And Len(txt_edat2) = 8 Then
        sss = sss & "   and tth_dat between " & txt_sdat2 & " and " & txt_edat2
    End If
    
    sss = sss & "  group by tth_testno, tth_dat, tth_seq, tth_title"
    sss = sss & "  order by tth_testno, tth_dat, tth_seq"
    
    sprd2.MaxRows = 0: cnt = 1
    
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Rs.RecordCount < 1 Then
        Rs.Close
        Call msg_display("��ϵ� ������ �����ϴ�!")
        Exit Sub
    End If

    sprd2.MaxRows = 0: cnt = 0
    
    Do While Not Rs.EOF
                  
        cnt = cnt + 1: sprd2.MaxRows = cnt: sprd2.row = cnt

        If Not IsNull(Rs!tth_testno) Then sprd2.Col = 1: sprd2.Text = Rs!tth_testno
        If Not IsNull(Rs!tth_title) Then sprd2.Col = 2: sprd2.Text = Rs!tth_title
        If Not IsNull(Rs!ttd_result) Then
                                                        
            sprd2.Col = 3: sprd2.Text = Rs!ttd_result
            
            If Rs!ttd_result = "OK" Then
                sprd2.ForeColor = &HFF0000
            ElseIf Rs!ttd_result = "NG" Then
                sprd2.ForeColor = &HFF&
            End If
            sprd2.Col = 3: sprd2.Text = Rs!ttd_result
        End If
        
        If Not IsNull(Rs!tth_no) Then sprd2.Col = 5: sprd2.Text = Format(Rs!tth_no, "########-###")
        
        Rs.MoveNext
        
    Loop
    
    Rs.Close
    
    Call msg_display("��ȸ�Ϸ�!")

Exit Sub

err_rtn:

    MsgBox (Err.Description)
    
End Sub


Private Sub sprd2_ButtonClicked(ByVal Col As Long, ByVal row As Long, ByVal ButtonDown As Integer)
    
    sprd2.row = row:
    
    sprd2.Col = 5
    
    If Len(sprd2.Text) <> 12 Then
        Call msg_display("������ȸ�� �۾��ϼ���!")
        Exit Sub
    End If
        
    txt_dat2 = Left(sprd2.Text, 8)
    txt_seq2 = Right(sprd2.Text, 3)

    Call print2
    
End Sub



Private Sub print2()
    
    Dim purpose As String
    Dim resultOK As String
    Dim resultNG As String
    
    sss = "       select *"
    sss = sss & "   from man_tooltesthd, man_tooltestds"
    sss = sss & "  where tth_dat = ttd_dat"
    sss = sss & "    and tth_seq = ttd_seq"
    sss = sss & "    and tth_dat = " & txt_dat2
    sss = sss & "    and tth_seq = " & txt_seq2
    sss = sss & "  order by ttd_lno"
    
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Rs.RecordCount < 1 Then
        Rs.Close
        Call msg_display("��ϵ� ������ �����ϴ�!")
        Exit Sub
    End If
    
    Call clear2
    
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
    
    purpose = "        TEST ����" & Chr(13)
    
    If Rs!tth_pyn1 = "Y" Then
        purpose = purpose & "        1.���� ���� ( �� )"
    Else
        purpose = purpose & "        1.���� ���� (     )"
    End If

    If Rs!tth_pyn2 = "Y" Then
        purpose = purpose & "             2.Ĩ ó�� ( �� )" & Chr(13)
    Else
        purpose = purpose & "             2.Ĩ ó�� (     )" & Chr(13)
    End If


    If Rs!tth_pyn3 = "Y" Then
        purpose = purpose & "        3.�ð� ���� ( �� )"
    Else
        purpose = purpose & "        3.�ð� ���� (     )"
    End If


    If Rs!tth_pyn4 = "Y" Then
        purpose = purpose & "             4.������ ���� ( �� )" & Chr(13)
    Else
        purpose = purpose & "             4.������ ���� (     )" & Chr(13)
    End If


    If Rs!tth_pyn5 = "Y" Then
        purpose = purpose & "        5.��Ÿ ( �� )"
    Else
        purpose = purpose & "        5.��Ÿ (     )"
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
    
    '���� �̹���
    If Not IsNull(Rs!tth_file1) Then
        If Right(Rs!tth_file1, 3) = "JPG" Then
           '==============================
           ' FTP�� �̿��� ���� �ٿ�ε�
           '==============================
           If FTP_Connection Then
              '
              If Not FTP���üũ(SERVER_PATH) Then
                 Call FTP_DisConnect
                 MsgBox "������θ� ã���� �����ϴ�.(������������ ����)"
                 Exit Sub
              End If
              '
              If Not FTP_Download(SERVER_PATH & Rs!tth_file1, "c:\temp\Temp.jpg") Then
              Else
                 sprd_print2.row = 1: sprd_print2.Col = 21: sprd_print2.TypePictPicture = LoadPicture("c:\temp\Temp.jpg")
              End If
              '
              Call FTP_DisConnect
              '
           End If
           '
           '===========================================
           ' ��Ʈ��ũ ����̹��� �̿��� ���� �ٿ�ε�
           '===========================================
           ' exists = ExistFile(N_Driver & ":\" & Rs!tth_file1)
           ' If exists = True Then
           '     FileCopy N_Driver & ":\" & Rs!tth_file1, "c:\temp\" & "Temp.jpg"
           '     GPath = "c:\temp\Temp.jpg"
           '     sprd_print2.Row = 1: sprd_print2.Col = 21: sprd_print2.TypePictPicture = LoadPicture(GPath)
           ' End If
        End If
    End If
    
    If Not IsNull(Rs!tth_rmk) Then sprd_print2.row = 19: sprd_print2.Col = 34: sprd_print2.Text = Rs!tth_rmk
    If Not IsNull(Rs!tth_cmt) Then sprd_print2.row = 28: sprd_print2.Col = 34: sprd_print2.Text = Rs!tth_cmt
    
    If Not IsNull(Rs!tth_file1) Then
        sprd_print2.row = 41: sprd_print2.Col = 34: sprd_print2.CellTag = Rs!tth_file1
        sprd_print2.CellType = CellTypeButton: sprd_print2.TypeButtonColor = &H8000000F
        sprd_print2.TypeButtonText = "����"
        sprd_print2.Col = 35:: sprd_print2.CellTag = Rs!tth_file1
        
    End If
    If Not IsNull(Rs!tth_file2) Then
        sprd_print2.row = 41: sprd_print2.Col = 36: sprd_print2.CellTag = Rs!tth_file2
        sprd_print2.CellType = CellTypeButton: sprd_print2.TypeButtonColor = &H8000000F
        sprd_print2.TypeButtonText = "����"
        sprd_print2.Col = 37:: sprd_print2.CellTag = Rs!tth_file2
        
    End If
    If Not IsNull(Rs!tth_file3) Then
        sprd_print2.row = 41: sprd_print2.Col = 38: sprd_print2.CellTag = Rs!tth_file3
        sprd_print2.TypeButtonText = "����"
        sprd_print2.CellType = CellTypeButton: sprd_print2.TypeButtonColor = &H8000000F
        sprd_print2.Col = 39:: sprd_print2.CellTag = Rs!tth_file3
        
    End If
    If Not IsNull(Rs!tth_file4) Then
        sprd_print2.row = 41: sprd_print2.Col = 40: sprd_print2.CellTag = Rs!tth_file4
        sprd_print2.TypeButtonText = "����"
        sprd_print2.CellType = CellTypeButton: sprd_print2.TypeButtonColor = &H8000000F
        sprd_print2.Col = 41:: sprd_print2.CellTag = Rs!tth_file4
    End If
        
        
    'DESC
    Do While Not Rs.EOF
                                
        '��������
        If Rs!ttd_lno = 1 Then
            sprd_print2.Col = 7
            If Not IsNull(Rs!ttd_maker) Then sprd_print2.row = 18:  sprd_print2.Text = Rs!ttd_maker & "(����)"
            If Not IsNull(Rs!ttd_tipstd) Then sprd_print2.row = 19: sprd_print2.Text = Rs!ttd_tipstd
            If Not IsNull(Rs!ttd_tipjil) Then sprd_print2.row = 20: sprd_print2.Text = Rs!ttd_tipjil
            If Not IsNull(Rs!ttd_holder) Then sprd_print2.row = 21: sprd_print2.Text = Rs!ttd_holder
            If Not IsNull(Rs!ttd_rcntmn) Then sprd_print2.row = 22: sprd_print2.Text = Rs!ttd_rcntmn
            If Not IsNull(Rs!ttd_rcntmx) Then sprd_print2.row = 22: sprd_print2.Text = sprd_print2.Text & "-" & Rs!ttd_rcntmx
            If Not IsNull(Rs!ttd_movmn) Then sprd_print2.row = 24:  sprd_print2.Text = Rs!ttd_movmn
            If Not IsNull(Rs!ttd_movmx) Then sprd_print2.row = 24:  sprd_print2.Text = sprd_print2.Text & "-" & Rs!ttd_movmx
            If Not IsNull(Rs!ttd_depth) Then sprd_print2.row = 23:  sprd_print2.Text = Rs!ttd_depth & "m/m"
            If Not IsNull(Rs!ttd_tct) Then sprd_print2.row = 25:    sprd_print2.Text = Format(Rs!ttd_tct, "###,##0")
            If Not IsNull(Rs!ttd_pct) Then sprd_print2.row = 25:    sprd_print2.Col = 12: sprd_print2.Text = Format(Rs!ttd_pct, "###,##0")
            If Not IsNull(Rs!ttd_fluid) Then
                sprd_print2.row = 26: sprd_print2.Col = 7:
                If Rs!ttd_fluid = 1 Then
                    sprd_print2.Text = "���뼺"
                Else
                    sprd_print2.Text = "����뼺"
                End If
            End If
            
            sprd_print2.Col = 7
            If Not IsNull(Rs!ttd_qty) Then sprd_print2.row = 27: sprd_print2.Text = Format(Rs!ttd_qty, "###,##0") & " Point"
            
            'If Not IsNull(Rs!ttd_dan) Then sprd_print2.Row = 28: sprd_print2.Text = Format(Rs!ttd_dan, "###,###,###") & "��"
            'If Not IsNull(Rs!ttd_tldn) Then sprd_print2.Row = 29: sprd_print2.Text = Format(Rs!ttd_tldn, "###,###,###") & "��/EA"
            'If Not IsNull(Rs!ttd_chdn) Then sprd_print2.Row = 30: sprd_print2.Text = Format(Rs!ttd_chdn, "###,###,###") & "��"
            
            If Not IsNull(Rs!ttd_dan) Then
                If InStr(1, Rs!ttd_dan, ".", 0) <> 0 Then
                    sprd_print2.row = 28: sprd_print2.Text = Format(Rs!ttd_dan, "###,###,##0.0#") & "��"
                Else
                    sprd_print2.row = 28: sprd_print2.Text = Format(Rs!ttd_dan, "###,###,##0") & "��"
                    
                End If
            End If
            If Not IsNull(Rs!ttd_tldn) Then
                If InStr(1, Rs!ttd_tldn, ".", 0) <> 0 Then
                    sprd_print2.row = 29: sprd_print2.Text = Format(Rs!ttd_tldn, "###,###,##0.0#") & "��/Corner"
                Else
                    sprd_print2.row = 29: sprd_print2.Text = Format(Rs!ttd_tldn, "###,###,##0") & "��/Corner"
                End If
            End If
            If Not IsNull(Rs!ttd_chdn) Then
                If InStr(1, Rs!ttd_chdn, ".", 0) <> 0 Then
                    sprd_print2.row = 30: sprd_print2.Text = Format(Rs!ttd_chdn, "###,###,##0.0#") & "��"
                Else
                    sprd_print2.row = 30: sprd_print2.Text = Format(Rs!ttd_chdn, "###,###,##0") & "��/EA"
                End If
            End If
            
        End If
        
        '�׽�Ʈ1
        If Rs!ttd_lno = 2 Then
        
            sprd_print2.Col = 16
            If Not IsNull(Rs!ttd_maker) Then sprd_print2.row = 18:  sprd_print2.Text = Rs!ttd_maker & "(�׽�Ʈ)"
            If Not IsNull(Rs!ttd_tipstd) Then sprd_print2.row = 19: sprd_print2.Text = Rs!ttd_tipstd
            If Not IsNull(Rs!ttd_tipjil) Then sprd_print2.row = 20: sprd_print2.Text = Rs!ttd_tipjil
            If Not IsNull(Rs!ttd_holder) Then sprd_print2.row = 21: sprd_print2.Text = Rs!ttd_holder
            If Not IsNull(Rs!ttd_rcntmn) Then sprd_print2.row = 22: sprd_print2.Text = Rs!ttd_rcntmn
            If Not IsNull(Rs!ttd_rcntmx) Then sprd_print2.row = 22: sprd_print2.Text = sprd_print2.Text & "-" & Rs!ttd_rcntmx
            If Not IsNull(Rs!ttd_movmn) Then sprd_print2.row = 24:  sprd_print2.Text = Rs!ttd_movmn
            If Not IsNull(Rs!ttd_movmx) Then sprd_print2.row = 24:  sprd_print2.Text = sprd_print2.Text & "-" & Rs!ttd_movmx
            If Not IsNull(Rs!ttd_depth) Then sprd_print2.row = 23:  sprd_print2.Text = Rs!ttd_depth & "m/m"
            If Not IsNull(Rs!ttd_tct) Then sprd_print2.row = 25:    sprd_print2.Text = Format(Rs!ttd_tct, "###,##0")
            If Not IsNull(Rs!ttd_pct) Then sprd_print2.row = 25:    sprd_print2.Col = 21: sprd_print2.Text = Format(Rs!ttd_pct, "###,##0")
            
            If Not IsNull(Rs!ttd_fluid) Then
                sprd_print2.row = 26: sprd_print2.Col = 16:
                If Rs!ttd_fluid = 1 Then
                    sprd_print2.Text = "���뼺"
                Else
                    sprd_print2.Text = "����뼺"
                End If
            End If
            
            sprd_print2.Col = 16
            If Not IsNull(Rs!ttd_qty) Then sprd_print2.row = 27: sprd_print2.Text = Format(Rs!ttd_qty, "###,##0") & " Point"
 
            If Not IsNull(Rs!ttd_dan) Then
                If InStr(1, Rs!ttd_dan, ".", 0) <> 0 Then
                    sprd_print2.row = 28: sprd_print2.Text = Format(Rs!ttd_dan, "###,###,##0.0#") & "��"
                Else
                    sprd_print2.row = 28: sprd_print2.Text = Format(Rs!ttd_dan, "###,###,##0") & "��"
                    
                End If
            End If
            If Not IsNull(Rs!ttd_tldn) Then
                If InStr(1, Rs!ttd_tldn, ".", 0) <> 0 Then
                    sprd_print2.row = 29: sprd_print2.Text = Format(Rs!ttd_tldn, "###,###,##0.0#") & "��/Corner"
                Else
                    sprd_print2.row = 29: sprd_print2.Text = Format(Rs!ttd_tldn, "###,###,##0") & "��/Corner"
                End If
            End If
            If Not IsNull(Rs!ttd_chdn) Then
                If InStr(1, Rs!ttd_chdn, ".", 0) <> 0 Then
                    sprd_print2.row = 30: sprd_print2.Text = Format(Rs!ttd_chdn, "###,###,##0.0#") & "��"
                Else
                    sprd_print2.row = 30: sprd_print2.Text = Format(Rs!ttd_chdn, "###,###,##0") & "��/EA"
                End If
            End If
            
            
        End If
        
        '�׽�Ʈ2
        If Rs!ttd_lno = 3 Then
            
            sprd_print2.Col = 25
            If Not IsNull(Rs!ttd_maker) Then sprd_print2.row = 18:  sprd_print2.Text = Rs!ttd_maker & "(�׽�Ʈ)"
            If Not IsNull(Rs!ttd_tipstd) Then sprd_print2.row = 19: sprd_print2.Text = Rs!ttd_tipstd
            If Not IsNull(Rs!ttd_tipjil) Then sprd_print2.row = 20: sprd_print2.Text = Rs!ttd_tipjil
            If Not IsNull(Rs!ttd_holder) Then sprd_print2.row = 21: sprd_print2.Text = Rs!ttd_holder
            If Not IsNull(Rs!ttd_rcntmn) Then sprd_print2.row = 22: sprd_print2.Text = Rs!ttd_rcntmn
            If Not IsNull(Rs!ttd_rcntmx) Then sprd_print2.row = 22: sprd_print2.Text = sprd_print2.Text & "-" & Rs!ttd_rcntmx
            If Not IsNull(Rs!ttd_movmn) Then sprd_print2.row = 24:  sprd_print2.Text = Rs!ttd_movmn
            If Not IsNull(Rs!ttd_movmx) Then sprd_print2.row = 24:  sprd_print2.Text = sprd_print2.Text & "-" & Rs!ttd_movmx
            If Not IsNull(Rs!ttd_depth) Then sprd_print2.row = 23:  sprd_print2.Text = Rs!ttd_depth & "m/m"
            If Not IsNull(Rs!ttd_tct) Then sprd_print2.row = 25:    sprd_print2.Text = Format(Rs!ttd_tct, "###,##0")
            If Not IsNull(Rs!ttd_pct) Then sprd_print2.row = 25:    sprd_print2.Col = 30: sprd_print2.Text = Format(Rs!ttd_pct, "###,##0")
            
            If Not IsNull(Rs!ttd_fluid) Then
                sprd_print2.row = 26: sprd_print2.Col = 25:
                If Rs!ttd_fluid = 1 Then
                    sprd_print2.Text = "���뼺"
                Else
                    sprd_print2.Text = "����뼺"
                End If
            End If
            
            sprd_print2.Col = 25
            If Not IsNull(Rs!ttd_qty) Then sprd_print2.row = 27: sprd_print2.Text = Format(Rs!ttd_qty, "###,##0") & " Point"

            If Not IsNull(Rs!ttd_dan) Then
                If InStr(1, Rs!ttd_dan, ".", 0) <> 0 Then
                    sprd_print2.row = 28: sprd_print2.Text = Format(Rs!ttd_dan, "###,###,##0.0#") & "��"
                Else
                    sprd_print2.row = 28: sprd_print2.Text = Format(Rs!ttd_dan, "###,###,##0") & "��"
                End If
            End If
            If Not IsNull(Rs!ttd_tldn) Then
                If InStr(1, Rs!ttd_tldn, ".", 0) <> 0 Then
                    sprd_print2.row = 29: sprd_print2.Text = Format(Rs!ttd_tldn, "###,###,##0.0#") & "��/Corner"
                Else
                    sprd_print2.row = 29: sprd_print2.Text = Format(Rs!ttd_tldn, "###,###,##0") & "��/Corner"
                End If
            End If
            If Not IsNull(Rs!ttd_chdn) Then
                If InStr(1, Rs!ttd_chdn, ".", 0) <> 0 Then
                    sprd_print2.row = 30: sprd_print2.Text = Format(Rs!ttd_chdn, "###,###,##0.0#") & "��"
                Else
                    sprd_print2.row = 30: sprd_print2.Text = Format(Rs!ttd_chdn, "###,###,##0") & "��/EA"
                End If
            End If
            
        End If
        
        If Rs!ttd_ryn = "Y" Then
               
            resultOK = ""
            resultNG = ""
            
            If Not IsNull(Rs!ttd_result) Then
                
                If Rs!ttd_result = "OK" Then
                    sprd_print2.row = 31: sprd_print2.Col = 7
                    sprd_print2.Text = "  O.K ( �� )": 'sprd_print2.FontBold = True
                    sprd_print2.row = 36: sprd_print2.Col = 7
                    sprd_print2.Text = "  N.G (     )"
                End If
                If Rs!ttd_result = "NG" Then
                    sprd_print2.row = 31: sprd_print2.Col = 7
                    sprd_print2.Text = "  O.K (     )"
                    sprd_print2.row = 36: sprd_print2.Col = 7
                    sprd_print2.Text = "  N.G ( �� )": 'sprd_print2.FontBold = True
                End If
            End If
                
            
            If Rs!ttd_result = "OK" Then
                
                If Rs!ttd_ryn1 = "Y" Then
                    resultOK = resultOK & "           1.�������� ���� ( �� )"
                Else
                    resultOK = resultOK & "           1.�������� ���� (     )"
                End If
    
                If Rs!ttd_ryn4 = "Y" Then
                    resultOK = resultOK & "            4.�ð� ���� ( �� )" & Chr(13)
                Else
                    resultOK = resultOK & "            4.�ð� ���� (     )" & Chr(13)
                End If
                    
                If Rs!ttd_ryn2 = "Y" Then
                    resultOK = resultOK & "           2.Ĩó�� ��ȣ ( �� )"
                Else
                    resultOK = resultOK & "           2.Ĩó�� ��ȣ (     )"
                End If
                    
                If Rs!ttd_ryn5 = "Y" Then
                    resultOK = resultOK & "               5.��Ÿ ( �� )" & Chr(13)
                Else
                    resultOK = resultOK & "               5.��Ÿ (     )" & Chr(13)
                End If
                    
                If Rs!ttd_ryn3 = "Y" Then
                    resultOK = resultOK & "           3.������ ���� ( �� )"
                Else
                    resultOK = resultOK & "           3.������ ���� (     )"
                End If
                    
                resultNG = resultNG & "           1.��� (     )"
                resultNG = resultNG & "                          4.������ ��� (     )" & Chr(13)
                resultNG = resultNG & "           2.���� (     )"
                resultNG = resultNG & "                          5.��Ÿ (     )" & Chr(13)
                resultNG = resultNG & "           3.Ĩó�� �ҷ� (     )"
                
                
            Else
                
                resultNG = ""
                
                If Rs!ttd_ryn1 = "Y" Then
                    resultNG = resultNG & "           1.��� ( �� )"
                Else
                    resultNG = resultNG & "           1.��� (     )"
                End If
    
                If Rs!ttd_ryn4 = "Y" Then
                    resultNG = resultNG & "                          4.������ ��� ( �� )" & Chr(13)
                Else
                    resultNG = resultNG & "                          4.������ ��� (     )" & Chr(13)
                End If
                    
                If Rs!ttd_ryn2 = "Y" Then
                    resultNG = resultNG & "           2.���� ( �� )"
                Else
                    resultNG = resultNG & "           2.���� (     )"
                End If
                    
                If Rs!ttd_ryn5 = "Y" Then
                    resultNG = resultNG & "                          5.��Ÿ ( �� )" & Chr(13)
                Else
                    resultNG = resultNG & "                          5.��Ÿ (     )" & Chr(13)
                End If
                    
                If Rs!ttd_ryn3 = "Y" Then
                    resultNG = resultNG & "           3.Ĩó�� �ҷ� ( �� )"
                Else
                    resultNG = resultNG & "           3.Ĩó�� �ҷ� (     )"
                End If
                
                resultOK = resultOK & "           1.�������� ���� (     )"
                resultOK = resultOK & "            4.�ð� ���� (     )" & Chr(13)
                resultOK = resultOK & "           2.Ĩó�� ��ȣ (     )"
                resultOK = resultOK & "               5.��Ÿ (     )" & Chr(13)
                resultOK = resultOK & "           3.������ ���� (     )"
            
            End If
                
            sprd_print2.row = 31: sprd_print2.Col = 14: sprd_print2.Text = resultOK
            sprd_print2.row = 36: sprd_print2.Col = 14: sprd_print2.Text = resultNG
            
            sprd_print2.row = 41: sprd_print2.Col = 7
            If Rs!ttd_rtyn = "Y" Then
                sprd_print2.Text = "�� TEST ���� ���� - OK"
            Else
                sprd_print2.Text = "�� TEST ���� ���� - NOT"
            End If
            
            If Rs!ttd_lno = 1 Then
                sprd_print2.Col = 7
            ElseIf Rs!ttd_lno = 2 Then
                sprd_print2.Col = 16
            ElseIf Rs!ttd_lno = 3 Then
                sprd_print2.Col = 25
            End If
                
            sprd_print2.row = 18: sprd_print2.FontBold = True: sprd_print2.RowHeight(18) = 10.5
            sprd_print2.row = 19: sprd_print2.FontBold = True: sprd_print2.RowHeight(19) = 10.5
                    
            sprd_print2.row = 28: sprd_print2.FontBold = True: sprd_print2.RowHeight(28) = 10.5
            sprd_print2.row = 29: sprd_print2.FontBold = True: sprd_print2.RowHeight(29) = 10.5
            sprd_print2.row = 30: sprd_print2.FontBold = True: sprd_print2.RowHeight(30) = 10.5
                
        End If
            
        Rs.MoveNext
    Loop
    
    Rs.Close
    
    'btn_view1.Enabled = False
    
    'txt_dat2.Enabled = False
    'txt_seq2.Enabled = False
    
    'btn_add1.Enabled = False
    'btn_mod1.Enabled = True
    'btn_del1.Enabled = True
    
    Call msg_display("��ȸ�Ϸ�!")
    
    
End Sub

Private Sub clear2()
    
    Dim ii As Double
    
    sprd_print2.row = 1
    sprd_print2.Col = 6: sprd_print2.Text = ""      'TEST NO
    sprd_print2.Col = 16: sprd_print2.Text = ""     'DATE

    sprd_print2.row = 2
    sprd_print2.Col = 6: sprd_print2.Text = ""      '������ȣ
    
    sprd_print2.row = 4
    sprd_print2.Col = 1: sprd_print2.Text = ""      'TEST ����
    
    sprd_print2.row = 4
    
    sprd_print2.Col = 7                             '�����/�ǻ���
    For ii = 6 To 13
        sprd_print2.row = ii: sprd_print2.Text = ""
    Next ii

    sprd_print2.row = 14
    sprd_print2.Col = 1: sprd_print2.Text = ""      'TEST ����
    
    sprd_print2.row = 17
    sprd_print2.Col = 26: sprd_print2.Text = ""     '�۾���
    
    sprd_print2.row = 1:
    sprd_print2.Col = 21: sprd_print2.TypePictPicture = Nothing  '�̹���
    
    sprd_print2.Col = 7                              '������������
    For ii = 18 To 30
        sprd_print2.row = ii: sprd_print2.Text = ""
    Next ii
    sprd_print2.row = 25
    sprd_print2.Col = 12: sprd_print2.Text = ""

    sprd_print2.Col = 16                             '�׽�Ʈ1 ����
    For ii = 18 To 30
        sprd_print2.row = ii: sprd_print2.Text = ""
    Next ii
    sprd_print2.row = 25
    sprd_print2.Col = 21: sprd_print2.Text = ""
    
    sprd_print2.Col = 25                             '�׽�Ʈ2 ����
    For ii = 18 To 30
        sprd_print2.row = ii: sprd_print2.Text = ""
    Next ii
    sprd_print2.row = 25
    sprd_print2.Col = 30: sprd_print2.Text = ""
    
    
    sprd_print2.row = 19
    sprd_print2.Col = 34: sprd_print2.Text = ""      '���
    
    sprd_print2.row = 28
    sprd_print2.Col = 34: sprd_print2.Text = ""      '��
    
    sprd_print2.row = 31
    sprd_print2.Col = 7: sprd_print2.Text = ""       '��� OK
    
    sprd_print2.row = 31
    sprd_print2.Col = 14: sprd_print2.Text = ""      '�������
    
    sprd_print2.row = 36
    sprd_print2.Col = 7: sprd_print2.Text = ""       '��� OK
    
    sprd_print2.row = 36
    sprd_print2.Col = 14: sprd_print2.Text = ""      '�������
    
    sprd_print2.row = 41
    sprd_print2.Col = 7: sprd_print2.Text = ""       '�� TEST����
    
    sprd_print2.Col = 7
    sprd_print2.row = 18: sprd_print2.FontBold = False: sprd_print2.RowHeight(18) = 10.5
    sprd_print2.row = 19: sprd_print2.FontBold = False: sprd_print2.RowHeight(19) = 10.5
                    
    sprd_print2.row = 28: sprd_print2.FontBold = False: sprd_print2.RowHeight(28) = 10.5
    sprd_print2.row = 29: sprd_print2.FontBold = False: sprd_print2.RowHeight(29) = 10.5
    sprd_print2.row = 30: sprd_print2.FontBold = False: sprd_print2.RowHeight(30) = 10.5

    sprd_print2.Col = 16
    sprd_print2.row = 18: sprd_print2.FontBold = False: sprd_print2.RowHeight(18) = 10.5
    sprd_print2.row = 19: sprd_print2.FontBold = False: sprd_print2.RowHeight(19) = 10.5
                    
    sprd_print2.row = 28: sprd_print2.FontBold = False: sprd_print2.RowHeight(28) = 10.5
    sprd_print2.row = 29: sprd_print2.FontBold = False: sprd_print2.RowHeight(29) = 10.5
    sprd_print2.row = 30: sprd_print2.FontBold = False: sprd_print2.RowHeight(30) = 10.5
    
    sprd_print2.Col = 25
    sprd_print2.row = 18: sprd_print2.FontBold = False: sprd_print2.RowHeight(18) = 10.5
    sprd_print2.row = 19: sprd_print2.FontBold = False: sprd_print2.RowHeight(19) = 10.5
                    
    sprd_print2.row = 28: sprd_print2.FontBold = False: sprd_print2.RowHeight(28) = 10.5
    sprd_print2.row = 29: sprd_print2.FontBold = False: sprd_print2.RowHeight(29) = 10.5
    sprd_print2.row = 30: sprd_print2.FontBold = False: sprd_print2.RowHeight(30) = 10.5
    
    '÷������
    sprd_print2.row = 41
    sprd_print2.Col = 34: sprd_print2.CellType = CellTypeEdit
    sprd_print2.Col = 36: sprd_print2.CellType = CellTypeEdit
    sprd_print2.Col = 38: sprd_print2.CellType = CellTypeEdit
    sprd_print2.Col = 40: sprd_print2.CellType = CellTypeEdit
    
    sprd_print2.row = 41:
    sprd_print2.Col = 34:: sprd_print2.CellTag = ""
    sprd_print2.Col = 35:: sprd_print2.CellTag = ""
    sprd_print2.Col = 36:: sprd_print2.CellTag = ""
    sprd_print2.Col = 37:: sprd_print2.CellTag = ""
    sprd_print2.Col = 38:: sprd_print2.CellTag = ""
    sprd_print2.Col = 39:: sprd_print2.CellTag = ""
    sprd_print2.Col = 40:: sprd_print2.CellTag = ""
    sprd_print2.Col = 41:: sprd_print2.CellTag = ""
    
End Sub

'÷������ ����

Private Sub sprd_print2_ButtonClicked(ByVal Col As Long, ByVal row As Long, ByVal ButtonDown As Integer)
    
On Error GoTo err_rtn
        
    Dim isrno As String     '�������� ��ȣ
    Dim filenm1 As String   '����1
    Dim filenm2 As String   '����2
    
    Dim fs            As Object 'Scripting.FileSystemObject ��ü
    Dim lsSouurce     As String '������ ���
    Dim lsDestination As String '������ġ ���
    Dim F_NAME     As String
    
    sprd_print2.row = row: sprd_print2.Col = Col
    '
    If Len(sprd_print2.CellTag) > 2 Then
       '
       If Len(sprd_print2.CellTag) < 5 Then
           msg_display ("��ϵ� ÷�������� �����ϴ�!")
           Exit Sub
       End If
       '
       If Install_ACROBET = False Then
           MsgBox "��ũ�κ������� ��ġ�Ǿ� ���� �ʾ� ���Ⱑ �Ұ����մϴ�."
           Exit Sub
       End If
       '
       mkpoen05_view.Show
       'mkpoen05_view.txt_dat = txt_dat2
       'mkpoen05_view.txt_seq = Format(txt_seq2, "000")
       '
       '==============================
       ' FTP�� �̿��� ���� �ٿ�ε�
       '==============================
       If FTP_Connection Then
          '
          If Not FTP���üũ(SERVER_PATH) Then
             Call FTP_DisConnect
             MsgBox "������θ� ã���� �����ϴ�.(������������ ����)"
             Exit Sub
          End If
          '
          If Not FTP_Download(SERVER_PATH & sprd_print2.CellTag, "c:\temp\" & sprd_print2.CellTag) Then
          Else
             mkpoen05_view.WebBrowser1.Navigate "c:\temp\" & sprd_print2.CellTag
             mkpoen05_view.txt_filenm = sprd_print2.CellTag
          End If
          '
          Call FTP_DisConnect
          '
       End If
       '
       '===========================================
       ' ��Ʈ��ũ ����̹��� �̿��� ���� �ٿ�ε�
       '===========================================
       ' mkpoen05_view.WebBrowser1.Navigate N_Driver & ":\" & sprd_print2.CellTag
       ' mkpoen05_view.txt_filenm = sprd_print2.CellTag
        '
       Exit Sub
       '
    End If
    '
    Exit Sub
    '
err_rtn:
  MsgBox (Err.Description)

End Sub

'-------------------

'������������ üũ
Private Function ExistFile(FilePath As String) As Long
     If LenB(Dir$(FilePath)) Then
          ExistFile = 1&
     Else
          ExistFile = 0&
     End If
End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    If SSTab1.Tab = 0 Then
        txt_sdat1.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        txt_sdat2.SetFocus
    End If
    
End Sub

'ETC
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
  
       pausetime = 0.02    ' �Ⱓ�� �����մϴ�.
       start = Timer       ' ���� �ð��� �����մϴ�.
       Do While Timer < start + pausetime
          DoEvents         ' �ٸ� ���ν����� �ѱ�ϴ�.
       Loop
   Next

End Sub

'---------------------------
' ��ũ�κ�
'---------------------------
'ACROBET ��ġ ���� Ȯ��
Private Function Install_ACROBET() As Boolean
  On Error GoTo Error_Handler
  '
  Dim MyPath As String
  Dim MyName As String
  
  
  MyPath = "C:\Program Files\"
  '
  ii = 0
  '
  MyName = Dir(MyPath, vbDirectory)
  Do While MyName <> ""
    ' ���� ���丮�� �����ϴ� ���丮�� �����մϴ�.
    If MyName <> "." And MyName <> ".." Then
      ' MyName�� ���丮���� Ȯ���ϱ� ���ؼ� ��Ʈ��(bitwise) �񱳸� ����մϴ�.
      If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
        If Left(UCase(Trim(MyName)), 5) = "ADOBE" Then
          '
          MyPath = "C:\Program Files\Adobe\"
          MyName = Dir(MyPath, vbDirectory)
          
          '������ ����
          Do While MyName <> ""
            ' ���� ���丮�� �����ϴ� ���丮�� �����մϴ�.
            If MyName <> "." And MyName <> ".." Then
              ' MyName�� ���丮���� Ȯ���ϱ� ���ؼ� ��Ʈ��(bitwise) �񱳸� ����մϴ�.
              If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
                If Left(UCase(Trim(MyName)), 7) = "ACROBAT" Then
                  ii = 1
                  Exit Do
                'Reader 8.0 �϶� : C:\Program Files\Adobe\Reader 8.0\Reader\AdobeUpdateCheck.exe
                ElseIf Left(UCase(Trim(MyName)), 6) = "READER" Then
                  ii = 1
                  Exit Do
                End If
              End If
            End If
            '
            MyName = Dir  ' ���� �׸��� �о���Դϴ�.
            '
          Loop
          '
          If ii = 1 Then Exit Do
          '
        End If
      End If
    End If
    '
    MyName = Dir  ' ���� �׸��� �о���Դϴ�.
    '
  Loop
  
  '64��Ʈ ��
  If ii = 0 Then
  
    MyPath = "C:\Program Files (x86)\"
    ii = 0
  '
    MyName = Dir(MyPath, vbDirectory)
    Do While MyName <> ""
      ' ���� ���丮�� �����ϴ� ���丮�� �����մϴ�.
      If MyName <> "." And MyName <> ".." Then
        ' MyName�� ���丮���� Ȯ���ϱ� ���ؼ� ��Ʈ��(bitwise) �񱳸� ����մϴ�.
        If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
          If Left(UCase(Trim(MyName)), 5) = "ADOBE" Then
            '
            MyPath = "C:\Program Files (x86)\Adobe\"
            MyName = Dir(MyPath, vbDirectory)
            
            '������ ����
            Do While MyName <> ""
              ' ���� ���丮�� �����ϴ� ���丮�� �����մϴ�.
              If MyName <> "." And MyName <> ".." Then
                ' MyName�� ���丮���� Ȯ���ϱ� ���ؼ� ��Ʈ��(bitwise) �񱳸� ����մϴ�.
                If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
                  If Left(UCase(Trim(MyName)), 7) = "ACROBAT" Then
                    ii = 1
                    Exit Do
                  'Reader 8.0 �϶� : C:\Program Files\Adobe\Reader 8.0\Reader\AdobeUpdateCheck.exe
                  ElseIf Left(UCase(Trim(MyName)), 6) = "READER" Then
                    ii = 1
                    Exit Do
                  End If
                End If
              End If
              '
              MyName = Dir  ' ���� �׸��� �о���Դϴ�.
              '
            Loop
            '
            If ii = 1 Then Exit Do
            '
          End If
        End If
      End If
      '
      MyName = Dir  ' ���� �׸��� �о���Դϴ�.
      '
    Loop
  End If
  '
  '
  If ii = 0 Then
    Install_ACROBET = False
  Else
    Install_ACROBET = True
  End If
  
  Exit Function
  '
Error_Handler:
  Install_ACROBET = False
End Function
