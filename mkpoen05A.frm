VERSION 5.00
Object = "{3B930683-5AF1-4F07-9CE8-CA8063E1F3DD}#2.0#0"; "XButton.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form mkpoen05A 
   BackColor       =   &H80000004&
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   9795
   ClientLeft      =   585
   ClientTop       =   2850
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   13980
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel3 
      Height          =   9780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13965
      _Version        =   65536
      _ExtentX        =   24633
      _ExtentY        =   17251
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
         TabIndex        =   95
         Top             =   0
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   17251
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "1.TEST DATA ���"
         TabPicture(0)   =   "mkpoen05A.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frm13"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "1.TEST DATA ��ȸ"
         TabPicture(1)   =   "mkpoen05A.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label6"
         Tab(1).Control(1)=   "sprd2"
         Tab(1).Control(2)=   "Frame3"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "TEST DATA �������"
         TabPicture(2)   =   "mkpoen05A.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprd3"
         Tab(2).Control(1)=   "Frame6"
         Tab(2).ControlCount=   2
         Begin FPSpreadADO.fpSpread sprd3 
            Height          =   8295
            Left            =   -74760
            TabIndex        =   177
            Top             =   1320
            Width           =   13575
            _Version        =   458752
            _ExtentX        =   23945
            _ExtentY        =   14631
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
            MaxCols         =   11
            MaxRows         =   0
            SpreadDesigner  =   "mkpoen05A.frx":0054
         End
         Begin VB.Frame Frame6 
            Height          =   855
            Left            =   -74880
            TabIndex        =   171
            Top             =   360
            Width           =   13710
            Begin VB.TextBox txt_edat3 
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
               TabIndex        =   173
               Top             =   300
               Width           =   1050
            End
            Begin VB.TextBox txt_sdat3 
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
               TabIndex        =   172
               Top             =   300
               Width           =   1050
            End
            Begin XLibrary_XButton.XButton btn_view3 
               Height          =   375
               Left            =   3600
               TabIndex        =   174
               Top             =   280
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
            Begin Threed.SSPanel SSPanel17 
               Height          =   330
               Left            =   120
               TabIndex        =   175
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
            Begin VB.Label Label17 
               Caption         =   "-"
               Height          =   285
               Left            =   2280
               TabIndex        =   176
               Top             =   360
               Width           =   150
            End
         End
         Begin VB.Frame Frame3 
            Height          =   855
            Left            =   -74880
            TabIndex        =   111
            Top             =   495
            Width           =   13710
            Begin VB.TextBox txt_test2 
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
               TabIndex        =   107
               Top             =   300
               Width           =   1110
            End
            Begin VB.Frame Frame5 
               Height          =   450
               Left            =   4620
               TabIndex        =   164
               Top             =   210
               Width           =   2250
               Begin Threed.SSOption opt_X 
                  Height          =   195
                  Left            =   90
                  TabIndex        =   108
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
                  TabIndex        =   105
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
                  TabIndex        =   106
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
               Left            =   1200
               MaxLength       =   8
               TabIndex        =   103
               Top             =   300
               Width           =   1050
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
               Left            =   2415
               MaxLength       =   8
               TabIndex        =   104
               Top             =   300
               Width           =   1050
            End
            Begin XLibrary_XButton.XButton btn_view2 
               Height          =   375
               Left            =   9210
               TabIndex        =   109
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
            Begin Threed.SSPanel SSPanel10 
               Height          =   330
               Left            =   120
               TabIndex        =   159
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
               TabIndex        =   165
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
               TabIndex        =   166
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
            Begin VB.Label Label16 
               Caption         =   "-"
               Height          =   285
               Left            =   2280
               TabIndex        =   160
               Top             =   360
               Width           =   150
            End
         End
         Begin VB.Frame frm13 
            Height          =   9315
            Left            =   90
            TabIndex        =   96
            Top             =   330
            Width           =   13770
            Begin MSComDlg.CommonDialog Comm1 
               Left            =   13200
               Top             =   120
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.TextBox txt_dat1 
               Appearance      =   0  '���
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1170
               MaxLength       =   8
               TabIndex        =   1
               Top             =   240
               Width           =   1185
            End
            Begin VB.TextBox txt_seq1 
               Appearance      =   0  '���
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2385
               MaxLength       =   3
               TabIndex        =   2
               Top             =   240
               Width           =   510
            End
            Begin VB.Frame Frame4 
               Height          =   2715
               Left            =   210
               TabIndex        =   139
               Top             =   6480
               Width           =   4815
               Begin FPSpreadADO.fpSpread spd_file11 
                  Height          =   1710
                  Left            =   60
                  TabIndex        =   161
                  TabStop         =   0   'False
                  Top             =   690
                  Width           =   4680
                  _Version        =   458752
                  _ExtentX        =   8255
                  _ExtentY        =   3016
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
                  MaxCols         =   4
                  MaxRows         =   4
                  ScrollBars      =   0
                  SpreadDesigner  =   "mkpoen05A.frx":053F
               End
               Begin FPSpreadADO.fpSpread spd_file12 
                  Height          =   1710
                  Left            =   60
                  TabIndex        =   162
                  TabStop         =   0   'False
                  Top             =   690
                  Width           =   4670
                  _Version        =   458752
                  _ExtentX        =   8237
                  _ExtentY        =   3016
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
                  MaxCols         =   4
                  MaxRows         =   4
                  ScrollBars      =   0
                  SpreadDesigner  =   "mkpoen05A.frx":0A63
               End
               Begin Threed.SSPanel SSPanel12 
                  Height          =   330
                  Left            =   60
                  TabIndex        =   163
                  Top             =   300
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "÷������"
                  BackColor       =   8438015
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
               Begin VB.Label lbl_cmt1 
                  Caption         =   "�� ����Ŭ�� ÷������ ����"
                  Height          =   225
                  Left            =   510
                  TabIndex        =   168
                  Top             =   2460
                  Width           =   2955
               End
            End
            Begin VB.Frame Frame2 
               Height          =   5895
               Left            =   210
               TabIndex        =   129
               Top             =   600
               Width           =   4815
               Begin Threed.SSPanel SSPanel15 
                  Height          =   315
                  Left            =   2940
                  TabIndex        =   170
                  Top             =   210
                  Width           =   1635
                  _Version        =   65536
                  _ExtentX        =   2884
                  _ExtentY        =   556
                  _StockProps     =   15
                  BackColor       =   12640511
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "����"
                     Size            =   9.01
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.OptionButton opt_n_su 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "����"
                     Height          =   180
                     Left            =   870
                     TabIndex        =   8
                     TabStop         =   0   'False
                     Top             =   60
                     Width           =   705
                  End
                  Begin VB.OptionButton opt_n_ja 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "�ڵ�"
                     Height          =   180
                     Left            =   90
                     TabIndex        =   7
                     TabStop         =   0   'False
                     Top             =   60
                     Value           =   -1  'True
                     Width           =   855
                  End
               End
               Begin VB.TextBox txt_seq 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000F&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2355
                  MaxLength       =   3
                  TabIndex        =   6
                  Top             =   210
                  Width           =   510
               End
               Begin VB.TextBox txt_dat 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000F&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   11.25
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   1140
                  MaxLength       =   8
                  TabIndex        =   5
                  Top             =   210
                  Width           =   1185
               End
               Begin VB.TextBox txt_lotno1 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
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
                  Left            =   960
                  MaxLength       =   8
                  TabIndex        =   18
                  Top             =   3540
                  Width           =   1200
               End
               Begin VB.TextBox txt_tsoknm1 
                  Appearance      =   0  '���
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   3000
                  MaxLength       =   20
                  TabIndex        =   30
                  TabStop         =   0   'False
                  Top             =   5490
                  Width           =   1575
               End
               Begin VB.TextBox txt_jajil1 
                  Appearance      =   0  '���
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   3900
                  MaxLength       =   6
                  TabIndex        =   22
                  TabStop         =   0   'False
                  Top             =   4245
                  Width           =   795
               End
               Begin VB.TextBox txt_bpjil1 
                  Appearance      =   0  '���
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   3900
                  MaxLength       =   6
                  TabIndex        =   20
                  TabStop         =   0   'False
                  Top             =   3900
                  Width           =   795
               End
               Begin VB.TextBox txt_mnm1 
                  Appearance      =   0  '���
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2010
                  MaxLength       =   20
                  TabIndex        =   15
                  TabStop         =   0   'False
                  Top             =   2100
                  Width           =   2580
               End
               Begin VB.TextBox txt_jseq1 
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
                  Left            =   2355
                  MaxLength       =   3
                  TabIndex        =   12
                  Top             =   960
                  Width           =   510
               End
               Begin VB.TextBox txt_jdat1 
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
                  Left            =   1140
                  MaxLength       =   8
                  TabIndex        =   11
                  Top             =   960
                  Width           =   1185
               End
               Begin VB.TextBox txt_tnm1 
                  Appearance      =   0  '���
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   1050
                  MaxLength       =   20
                  TabIndex        =   28
                  Top             =   5490
                  Width           =   1035
               End
               Begin VB.TextBox txt_tsab1 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
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
                  Left            =   2130
                  MaxLength       =   7
                  TabIndex        =   29
                  Top             =   5490
                  Width           =   825
               End
               Begin VB.TextBox txt_tdat1 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
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
                  Left            =   3420
                  MaxLength       =   20
                  TabIndex        =   10
                  Top             =   600
                  Width           =   1155
               End
               Begin VB.TextBox txt_testno1 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
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
                  Left            =   1140
                  MaxLength       =   20
                  TabIndex        =   9
                  Top             =   600
                  Width           =   1185
               End
               Begin VB.CheckBox chk_pyn5 
                  Caption         =   "5.��Ÿ"
                  Height          =   375
                  Left            =   3150
                  TabIndex        =   26
                  Top             =   4860
                  Width           =   1365
               End
               Begin VB.CheckBox chk_pyn4 
                  Caption         =   "4.������ ����"
                  Height          =   375
                  Left            =   3150
                  TabIndex        =   24
                  Top             =   4605
                  Width           =   1425
               End
               Begin VB.CheckBox chk_pyn3 
                  Caption         =   "3.�ð� ����"
                  Height          =   375
                  Left            =   1620
                  TabIndex        =   27
                  Top             =   5130
                  Width           =   1365
               End
               Begin VB.CheckBox chk_pyn2 
                  Caption         =   "2.Ĩ ó��"
                  Height          =   375
                  Left            =   1620
                  TabIndex        =   25
                  Top             =   4860
                  Width           =   1215
               End
               Begin VB.CheckBox chk_pyn1 
                  Caption         =   "1.���� ����"
                  Height          =   405
                  Left            =   1620
                  TabIndex        =   23
                  Top             =   4590
                  Width           =   1245
               End
               Begin VB.TextBox txt_bpcd1 
                  Appearance      =   0  '���
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   960
                  MaxLength       =   24
                  TabIndex        =   19
                  TabStop         =   0   'False
                  Top             =   3900
                  Width           =   2925
               End
               Begin VB.TextBox txt_jacd1 
                  Appearance      =   0  '���
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   960
                  MaxLength       =   24
                  TabIndex        =   21
                  TabStop         =   0   'False
                  Top             =   4245
                  Width           =   2925
               End
               Begin VB.TextBox txt_mcd1 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
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
                  Left            =   1245
                  MaxLength       =   5
                  TabIndex        =   14
                  Top             =   2100
                  Width           =   735
               End
               Begin VB.TextBox txt_mmake1 
                  Appearance      =   0  '���
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   1245
                  MaxLength       =   20
                  TabIndex        =   16
                  TabStop         =   0   'False
                  Top             =   2445
                  Width           =   3345
               End
               Begin VB.TextBox txt_msok1 
                  Appearance      =   0  '���
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   1245
                  MaxLength       =   20
                  TabIndex        =   17
                  TabStop         =   0   'False
                  Top             =   2790
                  Width           =   3345
               End
               Begin VB.TextBox txt_title1 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
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
                  Left            =   1530
                  MaxLength       =   20
                  TabIndex        =   13
                  Top             =   1350
                  Width           =   3045
               End
               Begin Threed.SSPanel SSPanel4 
                  Height          =   330
                  Index           =   1
                  Left            =   90
                  TabIndex        =   130
                  Top             =   1350
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "TEST ����"
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
               Begin Threed.SSPanel SSPanel1 
                  Height          =   330
                  Index           =   3
                  Left            =   90
                  TabIndex        =   131
                  Top             =   1740
                  Width           =   4500
                  _Version        =   65536
                  _ExtentX        =   7937
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   " ��    ��           ��    �� "
                  BackColor       =   12640511
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
               Begin Threed.SSPanel SSPanel4 
                  Height          =   330
                  Index           =   2
                  Left            =   90
                  TabIndex        =   132
                  Top             =   2100
                  Width           =   1110
                  _Version        =   65536
                  _ExtentX        =   1958
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "����ڵ�"
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
               Begin Threed.SSPanel SSPanel7 
                  Height          =   330
                  Left            =   90
                  TabIndex        =   133
                  Top             =   2445
                  Width           =   1110
                  _Version        =   65536
                  _ExtentX        =   1958
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "MAKER"
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
               Begin Threed.SSPanel SSPanel8 
                  Height          =   330
                  Left            =   90
                  TabIndex        =   134
                  Top             =   2790
                  Width           =   1110
                  _Version        =   65536
                  _ExtentX        =   1958
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "�μ���"
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
               Begin Threed.SSPanel SSPanel1 
                  Height          =   330
                  Index           =   4
                  Left            =   90
                  TabIndex        =   135
                  Top             =   3180
                  Width           =   4500
                  _Version        =   65536
                  _ExtentX        =   7937
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "        ��           ��           ��      "
                  BackColor       =   12640511
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
               Begin Threed.SSPanel SSPanel4 
                  Height          =   330
                  Index           =   3
                  Left            =   90
                  TabIndex        =   136
                  Top             =   3900
                  Width           =   825
                  _Version        =   65536
                  _ExtentX        =   1455
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "��ǰ�ڵ�"
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
               Begin Threed.SSPanel SSPanel9 
                  Height          =   330
                  Left            =   90
                  TabIndex        =   137
                  Top             =   4245
                  Width           =   825
                  _Version        =   65536
                  _ExtentX        =   1455
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "�����ڵ�"
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
               Begin Threed.SSPanel SSPanel13 
                  Height          =   330
                  Left            =   90
                  TabIndex        =   138
                  Top             =   4650
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "TSET ����"
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
               Begin Threed.SSPanel SSPanel4 
                  Height          =   330
                  Index           =   4
                  Left            =   90
                  TabIndex        =   140
                  Top             =   600
                  Width           =   1020
                  _Version        =   65536
                  _ExtentX        =   1799
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
               Begin Threed.SSPanel SSPanel4 
                  Height          =   330
                  Index           =   5
                  Left            =   2370
                  TabIndex        =   141
                  Top             =   600
                  Width           =   1020
                  _Version        =   65536
                  _ExtentX        =   1799
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "�׽�Ʈ ����"
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
               Begin Threed.SSPanel SSPanel4 
                  Height          =   330
                  Index           =   6
                  Left            =   90
                  TabIndex        =   142
                  Top             =   960
                  Width           =   1020
                  _Version        =   65536
                  _ExtentX        =   1799
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "������ȣ"
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
               Begin Threed.SSPanel SSPanel16 
                  Height          =   330
                  Left            =   90
                  TabIndex        =   143
                  Top             =   5490
                  Width           =   930
                  _Version        =   65536
                  _ExtentX        =   1640
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "�۾���"
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
               Begin Threed.SSPanel SSPanel4 
                  Height          =   330
                  Index           =   9
                  Left            =   90
                  TabIndex        =   146
                  Top             =   3540
                  Width           =   825
                  _Version        =   65536
                  _ExtentX        =   1455
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "LOT NO."
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
               Begin Threed.SSPanel SSPanel1 
                  Height          =   330
                  Index           =   5
                  Left            =   90
                  TabIndex        =   169
                  Top             =   210
                  Width           =   1020
                  _Version        =   65536
                  _ExtentX        =   1799
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "��Ϲ�ȣ"
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
            End
            Begin VB.Frame Frame1 
               Height          =   7875
               Left            =   5100
               TabIndex        =   99
               Top             =   600
               Width           =   8355
               Begin VB.TextBox txt_movmn1 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2020
                  MaxLength       =   20
                  TabIndex        =   39
                  Top             =   2910
                  Width           =   855
               End
               Begin VB.TextBox txt_movmx1 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   3030
                  MaxLength       =   20
                  TabIndex        =   40
                  Top             =   2910
                  Width           =   855
               End
               Begin VB.TextBox txt_movmn2 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4020
                  MaxLength       =   20
                  TabIndex        =   56
                  Top             =   2910
                  Width           =   855
               End
               Begin VB.TextBox txt_movmx2 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   5040
                  MaxLength       =   20
                  TabIndex        =   57
                  Top             =   2910
                  Width           =   855
               End
               Begin VB.TextBox txt_movmn3 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   6000
                  MaxLength       =   20
                  TabIndex        =   74
                  Top             =   2910
                  Width           =   855
               End
               Begin VB.TextBox txt_movmx3 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   7020
                  MaxLength       =   20
                  TabIndex        =   75
                  Top             =   2910
                  Width           =   855
               End
               Begin VB.TextBox txt_rcntmn1 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2020
                  MaxLength       =   20
                  TabIndex        =   36
                  Top             =   2220
                  Width           =   855
               End
               Begin VB.TextBox txt_rcntmx1 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   3030
                  MaxLength       =   20
                  TabIndex        =   37
                  Top             =   2220
                  Width           =   855
               End
               Begin VB.TextBox txt_rcntmn2 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4020
                  MaxLength       =   20
                  TabIndex        =   53
                  Top             =   2220
                  Width           =   855
               End
               Begin VB.TextBox txt_rcntmx2 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   5040
                  MaxLength       =   20
                  TabIndex        =   54
                  Top             =   2220
                  Width           =   855
               End
               Begin VB.TextBox txt_rcntmn3 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   6000
                  MaxLength       =   20
                  TabIndex        =   71
                  Top             =   2220
                  Width           =   855
               End
               Begin VB.TextBox txt_rcntmx3 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   7020
                  MaxLength       =   20
                  TabIndex        =   72
                  Top             =   2220
                  Width           =   855
               End
               Begin VB.ComboBox cmb_fluid3 
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   6000
                  TabIndex        =   78
                  Text            =   "1.���뼺"
                  Top             =   3610
                  Width           =   1890
               End
               Begin VB.ComboBox cmb_fluid2 
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4020
                  TabIndex        =   60
                  Text            =   "1.���뼺"
                  Top             =   3610
                  Width           =   1890
               End
               Begin VB.ComboBox cmb_fluid1 
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  ItemData        =   "mkpoen05A.frx":0F3C
                  Left            =   2040
                  List            =   "mkpoen05A.frx":0F3E
                  TabIndex        =   43
                  Text            =   "1.���뼺"
                  Top             =   3610
                  Width           =   1890
               End
               Begin VB.TextBox txt_pct3 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   7020
                  MaxLength       =   20
                  TabIndex        =   77
                  Top             =   3255
                  Width           =   855
               End
               Begin VB.TextBox txt_tct3 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   6000
                  MaxLength       =   20
                  TabIndex        =   76
                  Top             =   3255
                  Width           =   855
               End
               Begin VB.TextBox txt_pct2 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   5040
                  MaxLength       =   20
                  TabIndex        =   59
                  Top             =   3255
                  Width           =   855
               End
               Begin VB.TextBox txt_tct2 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4020
                  MaxLength       =   20
                  TabIndex        =   58
                  Top             =   3255
                  Width           =   855
               End
               Begin VB.TextBox txt_pct1 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   3030
                  MaxLength       =   20
                  TabIndex        =   42
                  Top             =   3255
                  Width           =   855
               End
               Begin FPSpreadADO.fpSpread sprd_rmk1 
                  Height          =   2190
                  Left            =   4020
                  TabIndex        =   90
                  Top             =   5520
                  Width           =   3885
                  _Version        =   458752
                  _ExtentX        =   6853
                  _ExtentY        =   3863
                  _StockProps     =   64
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
                  MaxCols         =   2
                  MaxRows         =   1
                  RetainSelBlock  =   0   'False
                  ScrollBars      =   0
                  SelectBlockOptions=   0
                  SpreadDesigner  =   "mkpoen05A.frx":0F40
               End
               Begin VB.CheckBox chk_ryn5 
                  Caption         =   "5.��Ÿ"
                  Height          =   405
                  Left            =   1770
                  TabIndex        =   88
                  Top             =   6930
                  Width           =   1695
               End
               Begin VB.CheckBox chk_ryn4 
                  Caption         =   "4.�ð� ����"
                  Height          =   405
                  Left            =   1770
                  TabIndex        =   87
                  Top             =   6660
                  Width           =   1725
               End
               Begin VB.CheckBox chk_ryn3 
                  Caption         =   "3.������ ����"
                  Height          =   405
                  Left            =   1770
                  TabIndex        =   86
                  Top             =   6390
                  Width           =   1785
               End
               Begin VB.CheckBox chk_ryn2 
                  Caption         =   "2.Ĩó�� ��ȣ"
                  Height          =   405
                  Left            =   1770
                  TabIndex        =   85
                  Top             =   6120
                  Width           =   1665
               End
               Begin VB.CheckBox chk_ryn1 
                  Caption         =   "1.�������� ����"
                  Height          =   405
                  Left            =   1770
                  TabIndex        =   84
                  Top             =   5850
                  Width           =   1635
               End
               Begin VB.ComboBox cmb_rtyn1 
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1770
                  TabIndex        =   89
                  Text            =   "N"
                  Top             =   7395
                  Width           =   630
               End
               Begin VB.ComboBox cmb_result1 
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1770
                  TabIndex        =   83
                  Top             =   5510
                  Width           =   1200
               End
               Begin VB.TextBox txt_chdn3 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   6000
                  MaxLength       =   20
                  TabIndex        =   82
                  Top             =   5055
                  Width           =   1875
               End
               Begin VB.TextBox txt_tldn3 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   6000
                  MaxLength       =   20
                  TabIndex        =   81
                  Top             =   4710
                  Width           =   1875
               End
               Begin VB.TextBox txt_dan3 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   6000
                  MaxLength       =   20
                  TabIndex        =   80
                  Top             =   4365
                  Width           =   1875
               End
               Begin VB.TextBox txt_qty3 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   6000
                  MaxLength       =   7
                  TabIndex        =   79
                  Top             =   4020
                  Width           =   855
               End
               Begin VB.TextBox txt_depth3 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   6000
                  MaxLength       =   20
                  TabIndex        =   73
                  Top             =   2565
                  Width           =   855
               End
               Begin VB.TextBox txt_holder3 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   6000
                  MaxLength       =   20
                  TabIndex        =   70
                  Top             =   1785
                  Width           =   1875
               End
               Begin VB.TextBox txt_tipjil3 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   6000
                  MaxLength       =   20
                  TabIndex        =   69
                  Top             =   1440
                  Width           =   1875
               End
               Begin VB.TextBox txt_tipstd3 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   6000
                  MaxLength       =   20
                  TabIndex        =   68
                  Top             =   1095
                  Width           =   1875
               End
               Begin VB.TextBox txt_maker3 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   6000
                  MaxLength       =   20
                  TabIndex        =   67
                  Top             =   750
                  Width           =   1875
               End
               Begin VB.TextBox txt_chdn2 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4020
                  MaxLength       =   20
                  TabIndex        =   64
                  Top             =   5055
                  Width           =   1875
               End
               Begin VB.TextBox txt_tldn2 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4020
                  MaxLength       =   20
                  TabIndex        =   63
                  Top             =   4710
                  Width           =   1875
               End
               Begin VB.TextBox txt_dan2 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4020
                  MaxLength       =   20
                  TabIndex        =   62
                  Top             =   4365
                  Width           =   1875
               End
               Begin VB.TextBox txt_qty2 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4020
                  MaxLength       =   7
                  TabIndex        =   61
                  Top             =   4020
                  Width           =   885
               End
               Begin VB.TextBox txt_depth2 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4020
                  MaxLength       =   20
                  TabIndex        =   55
                  Top             =   2565
                  Width           =   855
               End
               Begin VB.TextBox txt_holder2 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4020
                  MaxLength       =   20
                  TabIndex        =   52
                  Top             =   1785
                  Width           =   1875
               End
               Begin VB.TextBox txt_tipjil2 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4020
                  MaxLength       =   20
                  TabIndex        =   51
                  Top             =   1440
                  Width           =   1875
               End
               Begin VB.TextBox txt_tipstd2 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4020
                  MaxLength       =   20
                  TabIndex        =   50
                  Top             =   1095
                  Width           =   1875
               End
               Begin VB.TextBox txt_maker2 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4020
                  MaxLength       =   20
                  TabIndex        =   49
                  Top             =   750
                  Width           =   1875
               End
               Begin VB.TextBox txt_chdn1 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2025
                  MaxLength       =   20
                  TabIndex        =   47
                  Top             =   5055
                  Width           =   1875
               End
               Begin VB.TextBox txt_tldn1 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2025
                  MaxLength       =   20
                  TabIndex        =   46
                  Top             =   4710
                  Width           =   1875
               End
               Begin VB.TextBox txt_dan1 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2025
                  MaxLength       =   20
                  TabIndex        =   45
                  Top             =   4365
                  Width           =   1875
               End
               Begin VB.TextBox txt_qty1 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2025
                  MaxLength       =   7
                  TabIndex        =   44
                  Top             =   4020
                  Width           =   795
               End
               Begin VB.TextBox txt_tct1 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2025
                  MaxLength       =   20
                  TabIndex        =   41
                  Top             =   3255
                  Width           =   855
               End
               Begin VB.TextBox txt_depth1 
                  Alignment       =   1  '������ ����
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2025
                  MaxLength       =   20
                  TabIndex        =   38
                  Top             =   2565
                  Width           =   855
               End
               Begin VB.TextBox txt_holder1 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2025
                  MaxLength       =   20
                  TabIndex        =   35
                  Top             =   1785
                  Width           =   1875
               End
               Begin VB.TextBox txt_tipjil1 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2025
                  MaxLength       =   20
                  TabIndex        =   34
                  Top             =   1440
                  Width           =   1875
               End
               Begin VB.TextBox txt_tipstd1 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2025
                  MaxLength       =   20
                  TabIndex        =   33
                  Top             =   1095
                  Width           =   1875
               End
               Begin VB.TextBox txt_maker1 
                  Appearance      =   0  '���
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9.75
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2025
                  MaxLength       =   20
                  TabIndex        =   32
                  Top             =   750
                  Width           =   1875
               End
               Begin Threed.SSPanel SSPanel1 
                  Height          =   1365
                  Index           =   1
                  Left            =   120
                  TabIndex        =   100
                  Top             =   750
                  Width           =   420
                  _Version        =   65536
                  _ExtentX        =   741
                  _ExtentY        =   2399
                  _StockProps     =   15
                  Caption         =   " ��    ��           ��    �� "
                  BackColor       =   12640511
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "����"
                     Size            =   8.97
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSPanel SSPanel1 
                  Height          =   1710
                  Index           =   2
                  Left            =   120
                  TabIndex        =   101
                  Top             =   2220
                  Width           =   420
                  _Version        =   65536
                  _ExtentX        =   741
                  _ExtentY        =   3016
                  _StockProps     =   15
                  Caption         =   " ��    ��           ��    �� "
                  BackColor       =   12640511
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
               Begin Threed.SSPanel SSPanel4 
                  Height          =   330
                  Index           =   0
                  Left            =   570
                  TabIndex        =   102
                  Top             =   750
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "MAKER"
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
               Begin Threed.SSPanel SSPanel18 
                  Height          =   330
                  Left            =   570
                  TabIndex        =   110
                  Top             =   1095
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "TIP �԰�"
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
               Begin Threed.SSPanel SSPanel22 
                  Height          =   330
                  Left            =   570
                  TabIndex        =   112
                  Top             =   1440
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "TIP ����"
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
               Begin Threed.SSPanel SSPanel23 
                  Height          =   330
                  Left            =   570
                  TabIndex        =   113
                  Top             =   1785
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "HOLDER"
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
               Begin Threed.SSPanel SSPanel24 
                  Height          =   330
                  Left            =   570
                  TabIndex        =   114
                  Top             =   2220
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "�д� ȸ����"
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
               Begin Threed.SSPanel SSPanel25 
                  Height          =   330
                  Left            =   570
                  TabIndex        =   115
                  Top             =   2565
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "���Է�"
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
               Begin Threed.SSPanel SSPanel26 
                  Height          =   330
                  Left            =   570
                  TabIndex        =   116
                  Top             =   2910
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "�̼�(mm/rev)"
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
               Begin Threed.SSPanel SSPanel27 
                  Height          =   330
                  Left            =   570
                  TabIndex        =   117
                  Top             =   3255
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "T.C/T - P.C/T"
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
               Begin Threed.SSPanel SSPanel28 
                  Height          =   330
                  Left            =   550
                  TabIndex        =   118
                  Top             =   3600
                  Width           =   1440
                  _Version        =   65536
                  _ExtentX        =   2540
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "��������"
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
               Begin Threed.SSPanel SSPanel29 
                  Height          =   330
                  Left            =   570
                  TabIndex        =   119
                  Top             =   4020
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "������"
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
               Begin Threed.SSPanel SSPanel30 
                  Height          =   330
                  Left            =   570
                  TabIndex        =   120
                  Top             =   4365
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "�����ܰ�/EA"
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
               Begin Threed.SSPanel SSPanel31 
                  Height          =   330
                  Left            =   570
                  TabIndex        =   121
                  Top             =   4710
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "������/Corner"
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
               Begin Threed.SSPanel SSPanel32 
                  Height          =   330
                  Left            =   570
                  TabIndex        =   122
                  Top             =   5055
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "������ȯ��"
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
               Begin Threed.SSPanel SSPanel33 
                  Height          =   330
                  Left            =   2010
                  TabIndex        =   123
                  Top             =   390
                  Width           =   1905
                  _Version        =   65536
                  _ExtentX        =   3360
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "�� ��"
                  BackColor       =   12640511
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.OptionButton opt_ryn 
                     BackColor       =   &H00C0E0FF&
                     Height          =   225
                     Index           =   1
                     Left            =   60
                     TabIndex        =   31
                     TabStop         =   0   'False
                     Top             =   60
                     Width           =   255
                  End
               End
               Begin Threed.SSPanel SSPanel34 
                  Height          =   330
                  Left            =   4020
                  TabIndex        =   124
                  Top             =   390
                  Width           =   1875
                  _Version        =   65536
                  _ExtentX        =   3307
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "�׽�Ʈ1"
                  BackColor       =   12640511
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.OptionButton opt_ryn 
                     BackColor       =   &H00C0E0FF&
                     Height          =   225
                     Index           =   2
                     Left            =   30
                     TabIndex        =   48
                     Top             =   60
                     Width           =   255
                  End
               End
               Begin Threed.SSPanel SSPanel35 
                  Height          =   330
                  Left            =   6000
                  TabIndex        =   125
                  Top             =   390
                  Width           =   1875
                  _Version        =   65536
                  _ExtentX        =   3307
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "�׽�Ʈ2"
                  BackColor       =   12640511
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.OptionButton opt_ryn 
                     BackColor       =   &H00C0E0FF&
                     Height          =   225
                     Index           =   3
                     Left            =   60
                     TabIndex        =   65
                     Top             =   60
                     Width           =   255
                  End
                  Begin VB.CheckBox chk_test2 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   255
                     Left            =   1590
                     TabIndex        =   66
                     TabStop         =   0   'False
                     Top             =   40
                     Width           =   225
                  End
               End
               Begin Threed.SSPanel SSPanel2 
                  Height          =   330
                  Left            =   570
                  TabIndex        =   126
                  Top             =   5490
                  Width           =   1140
                  _Version        =   65536
                  _ExtentX        =   2011
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "�� ��"
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
               Begin Threed.SSPanel SSPanel5 
                  Height          =   330
                  Left            =   570
                  TabIndex        =   127
                  Top             =   5880
                  Width           =   1140
                  _Version        =   65536
                  _ExtentX        =   2011
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "�� ��"
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
               Begin Threed.SSPanel SSPanel6 
                  Height          =   330
                  Left            =   570
                  TabIndex        =   128
                  Top             =   7380
                  Width           =   1140
                  _Version        =   65536
                  _ExtentX        =   2011
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "�� TEST"
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
               Begin VB.Label Label15 
                  Caption         =   "(Point)"
                  Height          =   285
                  Left            =   6900
                  TabIndex        =   158
                  Top             =   4140
                  Width           =   675
               End
               Begin VB.Label Label14 
                  Caption         =   "(Point)"
                  Height          =   285
                  Left            =   4950
                  TabIndex        =   157
                  Top             =   4140
                  Width           =   675
               End
               Begin VB.Label Label13 
                  Caption         =   "(Point)"
                  Height          =   285
                  Left            =   2850
                  TabIndex        =   156
                  Top             =   4140
                  Width           =   675
               End
               Begin VB.Label Label12 
                  Caption         =   "(mm)"
                  Height          =   285
                  Left            =   6870
                  TabIndex        =   155
                  Top             =   2670
                  Width           =   495
               End
               Begin VB.Label Label11 
                  Caption         =   "(mm)"
                  Height          =   285
                  Left            =   4890
                  TabIndex        =   154
                  Top             =   2670
                  Width           =   495
               End
               Begin VB.Label Label10 
                  Caption         =   "(mm)"
                  Height          =   285
                  Left            =   2910
                  TabIndex        =   153
                  Top             =   2670
                  Width           =   495
               End
               Begin VB.Label Label9 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   6885
                  TabIndex        =   152
                  Top             =   2985
                  Width           =   105
               End
               Begin VB.Label Label8 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   4905
                  TabIndex        =   151
                  Top             =   2985
                  Width           =   105
               End
               Begin VB.Label Label7 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   2895
                  TabIndex        =   150
                  Top             =   2985
                  Width           =   105
               End
               Begin VB.Label Label5 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   6880
                  TabIndex        =   149
                  Top             =   2280
                  Width           =   105
               End
               Begin VB.Label Label4 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   4900
                  TabIndex        =   148
                  Top             =   2280
                  Width           =   105
               End
               Begin VB.Label Label3 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   2890
                  TabIndex        =   147
                  Top             =   2280
                  Width           =   105
               End
               Begin VB.Label Label2 
                  Caption         =   "�׽�Ʈ2 �ۼ��� Ŭ���ϼ��� ��"
                  Height          =   285
                  Left            =   5350
                  TabIndex        =   145
                  Top             =   150
                  Width           =   2805
               End
               Begin VB.Label Label1 
                  Caption         =   "�� ����� �����ϼ���(����� Ŭ��)"
                  Height          =   315
                  Left            =   2100
                  TabIndex        =   144
                  Top             =   150
                  Width           =   2955
               End
            End
            Begin Threed.SSPanel SSPanel1 
               Height          =   330
               Index           =   0
               Left            =   210
               TabIndex        =   97
               Top             =   240
               Width           =   930
               _Version        =   65536
               _ExtentX        =   1640
               _ExtentY        =   582
               _StockProps     =   15
               Caption         =   "��Ϲ�ȣ"
               BackColor       =   12640511
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
            Begin XLibrary_XButton.XButton btn_add1 
               Height          =   555
               Left            =   9240
               TabIndex        =   92
               Top             =   8595
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   979
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
               Text            =   "�� ��"
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
            Begin XLibrary_XButton.XButton btn_mod1 
               Height          =   555
               Left            =   10680
               TabIndex        =   93
               Top             =   8595
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   979
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
               Enabled         =   0   'False
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
               Text            =   "�� ��"
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
            Begin XLibrary_XButton.XButton btn_del1 
               Height          =   555
               Left            =   12120
               TabIndex        =   94
               TabStop         =   0   'False
               Top             =   8595
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   979
               BackColor1      =   12632256
               BackColor2      =   16777215
               BackColorEx     =   8421631
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
               Enabled         =   0   'False
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
               Text            =   "�� ��"
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
            Begin XLibrary_XButton.XButton btn_clear1 
               Height          =   330
               Left            =   4005
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   240
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   582
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
               Text            =   "Clear"
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
            Begin XLibrary_XButton.XButton btn_view1 
               Height          =   330
               Left            =   3135
               TabIndex        =   3
               Top             =   240
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   582
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
            Begin XLibrary_XButton.XButton btn_prt1 
               Height          =   555
               Left            =   5100
               TabIndex        =   91
               TabStop         =   0   'False
               Top             =   8580
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   979
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
               Text            =   "����� �� ��"
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
               ToolTipBodyText =   "�� ��"
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
         Begin FPSpreadADO.fpSpread sprd2 
            Height          =   7995
            Left            =   -74880
            TabIndex        =   98
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
            SpreadDesigner  =   "mkpoen05A.frx":1272
         End
         Begin VB.Label Label6 
            Caption         =   "�� ����Ŭ��: ���γ��� Ȯ��"
            Height          =   255
            Left            =   -74580
            TabIndex        =   167
            Top             =   1410
            Width           =   2475
         End
      End
   End
End
Attribute VB_Name = "mkpoen05A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '
    Const SERVER_PATH As String = "/�����׽�Ʈ_DATA/"
    '
    Dim ii              As Double
    Dim cnt             As Double
    '
    Dim File_Path       As String
    
    

Private Sub btn_view3_Click()
    
    sss = "select a.*,apl_isab,sinbun_name(apl_isab) as iname,apl_1yn,apl_2yn,apl_3yn,apl_4yn,apl_m1yn,apl_tdat,apl_1dat,apl_2dat,apl_3dat,apl_4dat,apl_m1dat,apl_alldat,"
    sss = sss & " apl_1sab,apl_2sab,apl_3sab,apl_4sab,apl_m1sab,"
    sss = sss & " sinbun_name(apl_1sab) as aname1,sinbun_name(apl_2sab) as aname2,sinbun_name(apl_3sab) as aname3,sinbun_name(apl_4sab) as aname4,sinbun_name(apl_m1sab) as mname"
    sss = sss & " from man_tooltesthd a, oth_applist b"
    sss = sss & " where tth_dat between " & txt_sdat3 & " and " & txt_edat3
    sss = sss & "   and tth_dat = apl_tdat(+)"
    sss = sss & "   and tth_seq = apl_tseq(+)"
    sss = sss & "   and apl_table(+) = 'man_tooltesthd'"
    sss = sss & " order by tth_dat,tth_seq"
    
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Rs.RecordCount < 1 Then
        Rs.Close
        Call msg_display("��ϵ� ������ �����ϴ�!")
        Exit Sub
    End If
    
    sprd3.MaxRows = 0: cnt = 0
    '
    Do While Not Rs.EOF
       '
       cnt = cnt + 1: sprd3.MaxRows = cnt: sprd3.row = cnt
       '
       If Not IsNull(Rs!tth_dat) Then sprd3.Col = 1: sprd3.Text = Rs!tth_dat
       If Not IsNull(Rs!tth_seq) Then sprd3.Col = 2: sprd3.Text = Rs!tth_seq
       If Not IsNull(Rs!tth_tdat) Then sprd3.Col = 3: sprd3.Text = Rs!tth_tdat
       If Not IsNull(Rs!tth_testno) Then sprd3.Col = 4: sprd3.Text = Rs!tth_testno
       If Not IsNull(Rs!tth_title) Then sprd3.Col = 5: sprd3.Text = Rs!tth_title
       '
       If Not IsNull(Rs!apl_isab) Then sprd3.Col = 6: sprd3.Text = Rs!iname
       If Rs!apl_tdat > 0 Then sprd3.Text = Rs!iname & Chr(13) & Rs!apl_tdat: sprd3.BackColor = RGB(215, 227, 188)
       
       If Not IsNull(Rs!apl_m1sab) Then If Rs!apl_m1sab > 0 Then sprd3.Col = 7: sprd3.Text = Rs!mname
       If Rs!apl_m1yn = "Y" Then sprd3.Col = 7: sprd3.Text = Rs!mname & Chr(13) & Rs!apl_m1dat: sprd3.BackColor = RGB(215, 227, 188)
       
       If Not IsNull(Rs!apl_1sab) Then If Rs!apl_1sab > 0 Then sprd3.Col = 8: sprd3.Text = Rs!aname1
       If Rs!apl_1yn = "Y" Then sprd3.Col = 8: sprd3.Text = Rs!aname1 & Chr(13) & Rs!apl_1dat: sprd3.BackColor = RGB(215, 227, 188)
       
       If Not IsNull(Rs!apl_2sab) Then If Rs!apl_2sab > 0 Then sprd3.Col = 9: sprd3.Text = Rs!aname2
       If Rs!apl_2yn = "Y" Then sprd3.Col = 9: sprd3.Text = Rs!aname2 & Chr(13) & Rs!apl_2dat: sprd3.BackColor = RGB(215, 227, 188)
       
       If Not IsNull(Rs!apl_3sab) Then If Rs!apl_3sab > 0 Then sprd3.Col = 10: sprd3.Text = Rs!aname3
       If Rs!apl_3yn = "Y" Then sprd3.Col = 10: sprd3.Text = Rs!aname3 & Chr(13) & Rs!apl_3dat: sprd3.BackColor = RGB(215, 227, 188)
       
       If Not IsNull(Rs!apl_4sab) Then sprd3.Col = 11: sprd3.Text = "�����"
       If Rs!apl_4yn = "Y" Then sprd3.Col = 11: sprd3.Text = "�����" & Chr(13) & Rs!apl_4dat: sprd3.BackColor = RGB(255, 214, 255)
       
       If Rs!apl_4dat = 99999999 Then sprd3.Col = 11: sprd3.Text = "���" & Chr(13) & Rs!apl_alldat: sprd3.BackColor = RGB(255, 214, 255)
       
       sprd3.RowHeight(sprd3.row) = sprd3.MaxTextRowHeight(sprd3.row) + 2
       
       Rs.MoveNext
       '
    Loop
    Rs.Close
    
End Sub

Private Sub Form_Load()
    
    '���
    cmb_result1.Clear
    cmb_result1.AddItem ""
    cmb_result1.AddItem "O.K"
    cmb_result1.AddItem "N.G"
    cmb_result1.ListIndex = 0
    
    cmb_fluid1.Clear
    cmb_fluid1.AddItem ""
    cmb_fluid1.AddItem "1.���뼺"
    cmb_fluid1.AddItem "2.����뼺"
    cmb_fluid1.ListIndex = 0
    
    cmb_fluid2.Clear
    cmb_fluid2.AddItem ""
    cmb_fluid2.AddItem "1.���뼺"
    cmb_fluid2.AddItem "2.����뼺"
    cmb_fluid2.ListIndex = 0
    
    cmb_fluid3.Clear
    cmb_fluid3.AddItem ""
    cmb_fluid3.AddItem "1.���뼺"
    cmb_fluid3.AddItem "2.����뼺"
    cmb_fluid3.ListIndex = 0
              
    cmb_rtyn1.Clear
    cmb_rtyn1.AddItem ""
    cmb_rtyn1.AddItem "N"
    cmb_rtyn1.AddItem "Y"
    cmb_rtyn1.ListIndex = 0
    
    txt_sdat3.Text = 20221018
    txt_edat3.Text = 20221020
    
    Call clear_rtn
    
End Sub

'===================================
'TAB1. TEST DATA ���
'===================================

'------------------------------------
'������ȸ
'------------------------------------
Private Sub btn_view1_Click()
    
    If IsNumeric(txt_dat1) = False Or Len(txt_dat1) <> 8 Then
        Call msg_display("��� ���ڸ� Ȯ���ϼ���!")
        txt_dat1.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txt_seq1) = False Or Len(txt_seq1) < 1 Then
        Call msg_display("��� ������ Ȯ���ϼ���!")
        txt_seq1.SetFocus
        Exit Sub
    End If
    
    Call clear_rtn

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
        Call msg_display("��ϵ� ������ �����ϴ�!")
        Exit Sub
    End If
        
    If Not IsNull(Rs!tth_testno) Then txt_testno1 = Rs!tth_testno
    If Not IsNull(Rs!tth_tdat) Then txt_tdat1 = Rs!tth_tdat
    If Not IsNull(Rs!tth_jubno) Then txt_jdat1 = Left(Rs!tth_jubno, 8)
    If Not IsNull(Rs!tth_jubno) Then txt_jseq1 = Right(Rs!tth_jubno, 3)
    If Not IsNull(Rs!tth_title) Then txt_title1 = Rs!tth_title
    If Not IsNull(Rs!tth_tmcd) Then txt_mcd1 = Rs!tth_tmcd
    Call txt_mcd1_LostFocus
                  
    If Not IsNull(Rs!tth_tlot) Then txt_lotno1 = Rs!tth_tlot
    Call txt_lotno1_LostFocus
    
    If Rs!tth_pyn1 = "Y" Then chk_pyn1.Value = 1
    If Rs!tth_pyn2 = "Y" Then chk_pyn2.Value = 1
    If Rs!tth_pyn3 = "Y" Then chk_pyn3.Value = 1
    If Rs!tth_pyn4 = "Y" Then chk_pyn4.Value = 1
    If Rs!tth_pyn5 = "Y" Then chk_pyn5.Value = 1

    If Not IsNull(Rs!tth_tsab) Then txt_tsab1 = Rs!tth_tsab
    txt_tsab1_LostFocus

    sprd_rmk1.row = 1
    If Not IsNull(Rs!tth_rmk) Then sprd_rmk1.Col = 1: sprd_rmk1.Text = Rs!tth_rmk
    If Not IsNull(Rs!tth_cmt) Then sprd_rmk1.Col = 2: sprd_rmk1.Text = Rs!tth_cmt
    'DESC
    
    Do While Not Rs.EOF
        
        '��������
        If Rs!ttd_lno = 1 Then
        
            If Not IsNull(Rs!ttd_maker) Then txt_maker1 = Rs!ttd_maker
            If Not IsNull(Rs!ttd_tipstd) Then txt_tipstd1 = Rs!ttd_tipstd
            If Not IsNull(Rs!ttd_tipjil) Then txt_tipjil1 = Rs!ttd_tipjil
            If Not IsNull(Rs!ttd_holder) Then txt_holder1 = Rs!ttd_holder
            If Not IsNull(Rs!ttd_rcntmn) Then txt_rcntmn1 = Rs!ttd_rcntmn
            If Not IsNull(Rs!ttd_rcntmx) Then txt_rcntmx1 = Rs!ttd_rcntmx
            If Not IsNull(Rs!ttd_movmn) Then txt_movmn1 = Rs!ttd_movmn
            If Not IsNull(Rs!ttd_movmx) Then txt_movmx1 = Rs!ttd_movmx
            If Not IsNull(Rs!ttd_depth) Then txt_depth1 = Rs!ttd_depth
            If Not IsNull(Rs!ttd_tct) Then txt_tct1 = Rs!ttd_tct
            If Not IsNull(Rs!ttd_pct) Then txt_pct1 = Rs!ttd_pct
            If Not IsNull(Rs!ttd_fluid) Then cmb_fluid1.ListIndex = Rs!ttd_fluid
            If Not IsNull(Rs!ttd_qty) Then txt_qty1 = Rs!ttd_qty
            If Not IsNull(Rs!ttd_dan) Then txt_dan1 = Rs!ttd_dan
            If Not IsNull(Rs!ttd_tldn) Then txt_tldn1 = Rs!ttd_tldn
            If Not IsNull(Rs!ttd_chdn) Then txt_chdn1 = Rs!ttd_chdn
            
            If Rs!ttd_ryn = "Y" Then
                opt_ryn(1).Value = True
                
                If Not IsNull(Rs!ttd_result) Then
                    If Rs!ttd_result = "OK" Then cmb_result1.ListIndex = 1
                    If Rs!ttd_result = "NG" Then cmb_result1.ListIndex = 2
                End If
                
                If Rs!ttd_ryn1 = "Y" Then chk_ryn1.Value = Checked
                If Rs!ttd_ryn2 = "Y" Then chk_ryn2.Value = Checked
                If Rs!ttd_ryn3 = "Y" Then chk_ryn3.Value = Checked
                If Rs!ttd_ryn4 = "Y" Then chk_ryn4.Value = Checked
                If Rs!ttd_ryn5 = "Y" Then chk_ryn5.Value = Checked

                If Rs!ttd_rtyn = "Y" Then
                    cmb_rtyn1.ListIndex = 2
                Else
                    cmb_rtyn1.ListIndex = 1
                End If
                
            End If


            
        End If
        
        '�׽�Ʈ1
        If Rs!ttd_lno = 2 Then
        
            If Not IsNull(Rs!ttd_maker) Then txt_maker2 = Rs!ttd_maker
            If Not IsNull(Rs!ttd_tipstd) Then txt_tipstd2 = Rs!ttd_tipstd
            If Not IsNull(Rs!ttd_tipjil) Then txt_tipjil2 = Rs!ttd_tipjil
            If Not IsNull(Rs!ttd_holder) Then txt_holder2 = Rs!ttd_holder
            If Not IsNull(Rs!ttd_rcntmn) Then txt_rcntmn2 = Rs!ttd_rcntmn
            If Not IsNull(Rs!ttd_rcntmx) Then txt_rcntmx2 = Rs!ttd_rcntmx
            If Not IsNull(Rs!ttd_movmn) Then txt_movmn2 = Rs!ttd_movmn
            If Not IsNull(Rs!ttd_movmx) Then txt_movmx2 = Rs!ttd_movmx
            If Not IsNull(Rs!ttd_depth) Then txt_depth2 = Rs!ttd_depth
            If Not IsNull(Rs!ttd_tct) Then txt_tct2 = Rs!ttd_tct
            If Not IsNull(Rs!ttd_pct) Then txt_pct2 = Rs!ttd_pct
            If Not IsNull(Rs!ttd_fluid) Then cmb_fluid2.ListIndex = Rs!ttd_fluid
            If Not IsNull(Rs!ttd_qty) Then txt_qty2 = Rs!ttd_qty
            If Not IsNull(Rs!ttd_dan) Then txt_dan2 = Rs!ttd_dan
            If Not IsNull(Rs!ttd_tldn) Then txt_tldn2 = Rs!ttd_tldn
            If Not IsNull(Rs!ttd_chdn) Then txt_chdn2 = Rs!ttd_chdn
            
            If Rs!ttd_ryn = "Y" Then
                opt_ryn(2).Value = True
                
                If Not IsNull(Rs!ttd_result) Then
                    If Rs!ttd_result = "OK" Then cmb_result1.ListIndex = 1
                    If Rs!ttd_result = "NG" Then cmb_result1.ListIndex = 2
                End If
                
                If Rs!ttd_ryn1 = "Y" Then chk_ryn1.Value = Checked
                If Rs!ttd_ryn2 = "Y" Then chk_ryn2.Value = Checked
                If Rs!ttd_ryn3 = "Y" Then chk_ryn3.Value = Checked
                If Rs!ttd_ryn4 = "Y" Then chk_ryn4.Value = Checked
                If Rs!ttd_ryn5 = "Y" Then chk_ryn5.Value = Checked

                If Rs!ttd_rtyn = "Y" Then
                    cmb_rtyn1.ListIndex = 2
                Else
                    cmb_rtyn1.ListIndex = 1
                End If
                
            End If
            
        End If
        
        '�׽�Ʈ2
        If Rs!ttd_lno = 3 Then
            
            chk_test2.Value = 1
            
            If Not IsNull(Rs!ttd_maker) Then txt_maker3 = Rs!ttd_maker
            If Not IsNull(Rs!ttd_tipstd) Then txt_tipstd3 = Rs!ttd_tipstd
            If Not IsNull(Rs!ttd_tipjil) Then txt_tipjil3 = Rs!ttd_tipjil
            If Not IsNull(Rs!ttd_holder) Then txt_holder3 = Rs!ttd_holder
            If Not IsNull(Rs!ttd_rcntmn) Then txt_rcntmn3 = Rs!ttd_rcntmn
            If Not IsNull(Rs!ttd_rcntmx) Then txt_rcntmx3 = Rs!ttd_rcntmx
            If Not IsNull(Rs!ttd_movmn) Then txt_movmn3 = Rs!ttd_movmn
            If Not IsNull(Rs!ttd_movmx) Then txt_movmx3 = Rs!ttd_movmx
            If Not IsNull(Rs!ttd_depth) Then txt_depth3 = Rs!ttd_depth
            If Not IsNull(Rs!ttd_tct) Then txt_tct3 = Rs!ttd_tct
            If Not IsNull(Rs!ttd_pct) Then txt_pct3 = Rs!ttd_pct
            If Not IsNull(Rs!ttd_fluid) Then cmb_fluid3.ListIndex = Rs!ttd_fluid
            If Not IsNull(Rs!ttd_qty) Then txt_qty3 = Rs!ttd_qty
            If Not IsNull(Rs!ttd_dan) Then txt_dan3 = Rs!ttd_dan
            If Not IsNull(Rs!ttd_tldn) Then txt_tldn3 = Rs!ttd_tldn
            If Not IsNull(Rs!ttd_chdn) Then txt_chdn3 = Rs!ttd_chdn
            
            If Rs!ttd_ryn = "Y" Then
                opt_ryn(3).Value = True
                
                If Not IsNull(Rs!ttd_result) Then
                    If Rs!ttd_result = "OK" Then cmb_result1.ListIndex = 1
                    If Rs!ttd_result = "NG" Then cmb_result1.ListIndex = 2
                End If
                
                If Rs!ttd_ryn1 = "Y" Then chk_ryn1.Value = Checked
                If Rs!ttd_ryn2 = "Y" Then chk_ryn2.Value = Checked
                If Rs!ttd_ryn3 = "Y" Then chk_ryn3.Value = Checked
                If Rs!ttd_ryn4 = "Y" Then chk_ryn4.Value = Checked
                If Rs!ttd_ryn5 = "Y" Then chk_ryn5.Value = Checked
                
                If Rs!ttd_rtyn = "Y" Then
                    cmb_rtyn1.ListIndex = 2
                Else
                    cmb_rtyn1.ListIndex = 1
                End If
                
            End If
            
        End If
        
        Rs.MoveNext
    Loop
    
    Rs.Close
    
    Call file_view
    
    spd_file11.Visible = False
    spd_file12.Visible = True
    lbl_cmt1.Visible = True
    
    btn_view1.Enabled = False
    
    txt_dat1.Enabled = False
    txt_seq1.Enabled = False
    
    btn_add1.Enabled = False
    btn_mod1.Enabled = True
    btn_del1.Enabled = True
    
    btn_prt1.Enabled = True
    
    txt_dat.Enabled = False
    txt_seq.Enabled = False
    opt_n_ja.Value = False
    opt_n_su.Value = False
    
    opt_n_ja.Enabled = False
    opt_n_su.Enabled = False
    
    Call msg_display("��ȸ�Ϸ�!")
    
End Sub

    
'------------------------------------
'�������
'------------------------------------
Private Sub btn_add1_Click()
    '
    On Error GoTo err_rtn
    '
    Dim Dat As Double
    Dim seq As Double
    
    Dim FileChk1 As String   '÷������1
    Dim FileChk2 As String   '÷������2
    Dim FileChk3 As String   '÷������2
    Dim FileChk4 As String   '÷������2
    Dim memo     As String   '�����߼� �޸�
   
    If Job_Level < 9 Then
        MsgBox ("�۾������� �����ϴ�!")
        Exit Sub
    End If
    
    '÷������Ȯ��1
    FileChk1 = Check_file(1, 2)
    If FileChk1 = "X" Then Exit Sub
    FileChk2 = Check_file(2, 2)
    If FileChk2 = "X" Then Exit Sub
    FileChk3 = Check_file(3, 2)
    If FileChk3 = "X" Then Exit Sub
    FileChk4 = Check_file(4, 2)
    If FileChk4 = "X" Then Exit Sub
    
    If Check_Insert_Data = 9 Then Exit Sub
    
    
    '������� �ڵ��ο�
    If opt_n_ja.Value = True Then
    '��¥��ȸ
        Dat = today("YYYYMMDD")
        
        '===================================================================
        '������ȸ
        sss = "      select nvl(max(tth_seq),0) + 1 as seq from man_tooltesthd"
        sss = sss & " where tth_dat = " & Dat
        Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        seq = Rs!seq
        Rs.Close
        
    '������� �����ο�(�������� �ұ޵��)
    Else
        '��Ϲ�ȣ Ȯ��(���� ��¥�Է�)
        If IsDate(Left(txt_dat, 4) & "/" & Mid(txt_dat, 5, 2) & "/" & Right(txt_dat, 2)) = False Then
            Call msg_display("��Ϲ�ȣ ��¥�� Ȯ���ϼ���! (��Ͻ���)")
            txt_dat.SetFocus
            Exit Sub
        End If
        
        '��Ϲ�ȣ Ȯ��(���� �����Է�)
        If Val(Trim(txt_seq)) < 1 Or IsNumeric(txt_seq) = False Then
            Call msg_display("��Ϲ�ȣ ������ Ȯ���ϼ���! (��Ͻ���)")
            txt_seq.SetFocus
            Exit Sub
        End If
        
        '������Ϲ�ȣ Ȯ��
        sss = "       select * "
        sss = sss & "   from man_tooltesthd"
        sss = sss & "  where tth_dat = " & txt_dat
        sss = sss & "    and tth_seq = " & txt_seq
                    
        Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        If Rs.RecordCount > 0 Then
            Rs.Close
            Call msg_display("������Ϲ�ȣ�� �����մϴ�! (��Ͻ���)")
            txt_dat.SetFocus
            Exit Sub
        End If
    
        Rs.Close
        
        Dat = txt_dat
        seq = Val(Trim(txt_seq))
        
    End If
        
    Ws.BeginTrans
    
    '=======================
    'HEAD�Է�
    '=======================
    sss = "insert into man_tooltesthd("
    sss = sss & "         tth_dat,"     '�������
    sss = sss & "         tth_seq,"     '��ϼ���
    sss = sss & "         tth_testno,"  'TEST NO.
    sss = sss & "         tth_title,"   '�׽�Ʈ ����
    sss = sss & "         tth_tlot,"    '�׽�Ʈ ���� LOT No.
    sss = sss & "         tth_tmcd,"    '�׽�Ʈ ����ڵ�
    sss = sss & "         tth_tsab,"    '�׽�Ʈ �۾���
    sss = sss & "         tth_tdat,"    '�׽�Ʈ ����
    sss = sss & "         tth_jubno,"    '��������
    sss = sss & "         tth_pyn1,"    'Y/N - 1.��������
    sss = sss & "         tth_pyn2,"    'Y/N - 2.Ĩó��
    sss = sss & "         tth_pyn3,"    'Y/N - 3.�ð�����
    sss = sss & "         tth_pyn4,"    'Y/N - 4.������
    sss = sss & "         tth_pyn5,"    'Y/N - 5.��Ÿ

    sss = sss & "         tth_sab,"     '�Է���
    sss = sss & "         tth_rmk,"     '���
    sss = sss & "         tth_cmt,"     '��
                            
    sss = sss & "         tth_file1,"   '����1
    sss = sss & "         tth_file2,"   '����2
    sss = sss & "         tth_file3,"   '����3
    sss = sss & "         tth_file4,"   '����4
    
    sss = sss & "         tth_indte,"   '�Է�����
    sss = sss & "         tth_updte"    '��������
    sss = sss & ")values( "
    sss = sss & "        " & Dat & ","
    sss = sss & "        " & seq & ","
    sss = sss & "        '" & txt_testno1 & "',"
    sss = sss & "        '" & txt_title1 & "',"
    sss = sss & "        '" & txt_lotno1 & "',"
    sss = sss & "        '" & txt_mcd1 & "',"
    sss = sss & "        " & Val(txt_tsab1) & ","
    sss = sss & "        " & Val(txt_tdat1) & ","
    sss = sss & "        " & txt_jdat1 & Format(txt_jseq1, "000") & ","
    If chk_pyn1.Value = 1 Then
        sss = sss & "        'Y',"
    Else
        sss = sss & "        'N',"
    End If
    
    If chk_pyn2.Value = 1 Then
        sss = sss & "        'Y',"
    Else
        sss = sss & "        'N',"
    End If
    
    If chk_pyn3.Value = 1 Then
        sss = sss & "        'Y',"
    Else
        sss = sss & "        'N',"
    End If
    
    If chk_pyn4.Value = 1 Then
        sss = sss & "        'Y',"
    Else
        sss = sss & "        'N',"
    End If
    
    If chk_pyn5.Value = 1 Then
        sss = sss & "        'Y',"
    Else
        sss = sss & "        'N',"
    End If
    
    sss = sss & "        " & Val(Gsab) & ","
    sprd_rmk1.row = 1: sprd_rmk1.Col = 1
    sss = sss & "        '" & sprd_rmk1.Text & "',"
    sprd_rmk1.row = 1: sprd_rmk1.Col = 2
    sss = sss & "        '" & sprd_rmk1.Text & "',"
    
    '============÷������
    '===========================
    ' FTP�� �̿��� ���� ���ε�
    '===========================
    If FTP_Connection Then
       '
       If Not FTP���üũ(SERVER_PATH) Then
          Call FTP_DisConnect
          Ws.Rollback
          MsgBox "������θ� ã���� �����ϴ�.(������������ ����), ÷������ ����"
          Exit Sub
       End If
       '
       '÷������1(�̹�������)
       If FileChk1 = "Y" Then
          '
          spd_file11.row = 1: spd_file11.Col = 2
          If UCase(Right(spd_file11.Text, 3)) = "PDF" Then
             sss = sss & "       '" & Dat & Format(seq, "000") & "-001" & ".PDF" & "',"
             '
             File_Path = Dat & Format(seq, "000") & "-001" & ".PDF"
             '
             If Not FTP_Upload(spd_file11.Text, SERVER_PATH, File_Path) Then
                Ws.Rollback
                Call FTP_DisConnect
                MsgBox "���� ���ε忡 ���� �Ͽ����ϴ�. (������������ ����)"
                Exit Sub
             End If
             'FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-001" & ".PDF"  '��Ʈ�������� ��������
          ElseIf UCase(Right(spd_file11.Text, 3)) = "JPG" Then
             sss = sss & "       '" & Dat & Format(seq, "000") & "-001" & ".JPG" & "',"
             '
             File_Path = Dat & Format(seq, "000") & "-001" & ".JPG"
             '
             If Not FTP_Upload(spd_file11.Text, SERVER_PATH, File_Path) Then
                Ws.Rollback
                Call FTP_DisConnect
                MsgBox "���� ���ε忡 ���� �Ͽ����ϴ�. (������������ ����)"
                Exit Sub
             End If
            ' FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-001" & ".JPG"  '��Ʈ�������� ��������
          End If
       Else
           sss = sss & "       null,"
       End If
    
       '÷������2
       If FileChk2 = "Y" Then
          '
          spd_file11.row = 2: spd_file11.Col = 2
          If UCase(Right(spd_file11.Text, 3)) = "PDF" Then
             sss = sss & "       '" & Dat & Format(seq, "000") & "-002" & ".PDF" & "',"
             '
             File_Path = Dat & Format(seq, "000") & "-002" & ".PDF"
             '
             If Not FTP_Upload(spd_file11.Text, SERVER_PATH, File_Path) Then
                Ws.Rollback
                Call FTP_DisConnect
                MsgBox "���� ���ε忡 ���� �Ͽ����ϴ�. (������������ ����)"
                Exit Sub
             End If
            ' FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-002" & ".PDF"  '��Ʈ�������� ��������
          ElseIf UCase(Right(spd_file11.Text, 3)) = "JPG" Then
             sss = sss & "       '" & Dat & Format(seq, "000") & "-002" & ".JPG" & "'," '����1
             '
             File_Path = Dat & Format(seq, "000") & "-002" & ".JPG"
             '
             If Not FTP_Upload(spd_file11.Text, SERVER_PATH, File_Path) Then
                Ws.Rollback
                Call FTP_DisConnect
                MsgBox "���� ���ε忡 ���� �Ͽ����ϴ�. (������������ ����)"
                Exit Sub
             End If
            ' FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-002" & ".JPG"  '��Ʈ�������� ��������
          End If
       Else
          sss = sss & "       null,"
       End If
       '
       '÷������3
       If FileChk3 = "Y" Then
          '
          spd_file11.row = 3: spd_file11.Col = 2
          If UCase(Right(spd_file11.Text, 3)) = "PDF" Then
             sss = sss & "       '" & Dat & Format(seq, "000") & "-003" & ".PDF" & "',"
             '
             File_Path = Dat & Format(seq, "000") & "-003" & ".PDF"
             '
             If Not FTP_Upload(spd_file11.Text, SERVER_PATH, File_Path) Then
                Ws.Rollback
                Call FTP_DisConnect
                MsgBox "���� ���ε忡 ���� �Ͽ����ϴ�. (������������ ����)"
                Exit Sub
             End If
            ' FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-003" & ".PDF"  '��Ʈ�������� ��������
          ElseIf UCase(Right(spd_file11.Text, 3)) = "JPG" Then
             sss = sss & "       '" & Dat & Format(seq, "000") & "-003" & ".JPG" & "'," '����1
             '
             File_Path = Dat & Format(seq, "000") & "-003" & ".JPG"
             '
             If Not FTP_Upload(spd_file11.Text, SERVER_PATH, File_Path) Then
                Ws.Rollback
                Call FTP_DisConnect
                MsgBox "���� ���ε忡 ���� �Ͽ����ϴ�. (������������ ����)"
                Exit Sub
             End If
            ' FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-003" & ".JPG"  '��Ʈ�������� ��������
          End If
       Else
          sss = sss & "       null,"
       End If
       '
       '÷������4
       If FileChk4 = "Y" Then
          '
          spd_file11.row = 4: spd_file11.Col = 2
          If UCase(Right(spd_file11.Text, 3)) = "PDF" Then
             sss = sss & "       '" & Dat & Format(seq, "000") & "-004" & ".PDF" & "'," '����1
             '
             File_Path = Dat & Format(seq, "000") & "-004" & ".PDF"
             '
             If Not FTP_Upload(spd_file11.Text, SERVER_PATH, File_Path) Then
                Ws.Rollback
                Call FTP_DisConnect
                MsgBox "���� ���ε忡 ���� �Ͽ����ϴ�. (������������ ����)"
                Exit Sub
             End If
            ' FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-004" & ".PDF"  '��Ʈ�������� ��������
          ElseIf UCase(Right(spd_file11.Text, 3)) = "JPG" Then
             sss = sss & "       '" & Dat & Format(seq, "000") & "-004" & ".JPG" & "'," '����1
             '
             File_Path = Dat & Format(seq, "000") & "-004" & ".JPG"
             '
             If Not FTP_Upload(spd_file11.Text, SERVER_PATH, File_Path) Then
                Ws.Rollback
                Call FTP_DisConnect
                MsgBox "���� ���ε忡 ���� �Ͽ����ϴ�. (������������ ����)"
                Exit Sub
             End If
            ' FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-004" & ".JPG"  '��Ʈ�������� ��������
          End If
       Else
          sss = sss & "       null,"
       End If
       '
       Call FTP_DisConnect
       '
    End If
    '======================
    '
    '
    sss = sss & "        to_char(sysdate,'yyyymmdd'),"
    sss = sss & "        to_char(sysdate,'yyyymmdd')"
    sss = sss & "       )"
    
    db.Execute sss, 64
    
    '=======================
    'DESC�Է�
    '=======================
    '---------------------->
    '�����������
    '---------------------->
    sss = "insert into man_tooltestds("
    sss = sss & "         ttd_dat,"     '�������
    sss = sss & "         ttd_seq,"     '��ϼ���
    sss = sss & "         ttd_lno,"     '������
    sss = sss & "         ttd_gbn,"     '1:��������/2:�׽�Ʈ����
    sss = sss & "         ttd_ryn,"     'Y/N - �������
    
    sss = sss & "         ttd_maker,"   '������
    sss = sss & "         ttd_tipstd,"  '�� �԰�/�ڵ�
    sss = sss & "         ttd_tipjil,"  '�� ����
    sss = sss & "         ttd_holder,"  '����Ȧ��
    sss = sss & "         ttd_rcntmn,"  '�д�ȸ�� �ּ�
    sss = sss & "         ttd_rcntmx,"  '�д�ȸ�� �ִ�
    sss = sss & "         ttd_movmn,"   '�̼�(MM/REV) �ּ�
    sss = sss & "         ttd_movmx,"   '�̼�(MM/REV) �ִ�
    sss = sss & "         ttd_tct,"     'TCT
    sss = sss & "         ttd_pct,"     'PCT
                        
    sss = sss & "         ttd_depth,"   '�������
    
    sss = sss & "         ttd_fluid,"   '��������
    sss = sss & "         ttd_qty,"     '��������
    sss = sss & "         ttd_dan,"     '�ܰ�/EA
    sss = sss & "         ttd_tldn,"    '������/EA
    sss = sss & "         ttd_chdn,"    '��ȯ��/EA
    sss = sss & "         ttd_result,"  '��� OK/NG
    sss = sss & "         ttd_ryn1,"    'Y/N
    sss = sss & "         ttd_ryn2,"    'Y/N
    sss = sss & "         ttd_ryn3,"    'Y/N
    sss = sss & "         ttd_ryn4,"    'Y/N
    sss = sss & "         ttd_ryn5,"    'Y/N
    sss = sss & "         ttd_rtyn, "   '�� �׽�Ʈ ����
    sss = sss & "         ttd_rmk,"     '
    sss = sss & "         ttd_updte"    '
    
    sss = sss & ")values( "
    sss = sss & "        " & Dat & ","
    sss = sss & "        " & seq & ","
    sss = sss & "        1,"
    sss = sss & "        1,"
    If opt_ryn(1).Value = True Then
        sss = sss & "        'Y',"
    Else
        sss = sss & "        'N',"
    End If
                   
    sss = sss & "        '" & txt_maker1 & "',"
    sss = sss & "        '" & txt_tipstd1 & "',"
    sss = sss & "        '" & txt_tipjil1 & "',"
    sss = sss & "        '" & txt_holder1 & "',"
    
    If Val(txt_rcntmn1) > 0 Then    '�д�ȸ���� �ּҰ�
        sss = sss & "        " & Val(txt_rcntmn1) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_rcntmx1) > 0 Then    '�д�ȸ���� �ִ밪
        sss = sss & "        " & Val(txt_rcntmx1) & ","
    Else
        sss = sss & "        null,"
    End If
                                                        
    If Val(txt_movmn1) > 0 Then    '�̼� �ּҰ�
        sss = sss & "        " & Val(txt_movmn1) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_movmx1) > 0 Then    '�̼� �ִ밪
        sss = sss & "        " & Val(txt_movmx1) & ","
    Else
        sss = sss & "        null,"
    End If
                                               
    If Val(txt_tct1) > 0 Then      'TCP
        sss = sss & "        " & Val(txt_tct1) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_pct1) > 0 Then      'PCT
        sss = sss & "        " & Val(txt_pct1) & ","
    Else
        sss = sss & "        null,"
    End If
    
    If Val(txt_depth1) > 0 Then    '�������
        sss = sss & "        " & Val(txt_depth1) & ","
    Else
        sss = sss & "        null,"
    End If

    sss = sss & "        " & Left(cmb_fluid1, 1) & ","
    sss = sss & "        " & Val(txt_qty1) & ","
    sss = sss & "        " & Val(txt_dan1) & ","
    sss = sss & "        " & Val(txt_tldn1) & ","
    sss = sss & "        " & Val(txt_chdn1) & ","
    
    If opt_ryn(1).Value = True Then
        
        sss = sss & "        '" & Replace(cmb_result1, ".", "") & "',"
        
        If chk_ryn1.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn2.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn3.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn4.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn5.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        sss = sss & "       '" & cmb_rtyn1 & "',"
    Else
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
    End If
    
    sss = sss & "        '',"
    sss = sss & "        to_char(sysdate,'yyyymmdd')"
    sss = sss & " )"
    
    db.Execute sss, 64

    '---------------------->
    '��TEST1��
    '---------------------->
    sss = "insert into man_tooltestds("
    sss = sss & "         ttd_dat,"     '�������
    sss = sss & "         ttd_seq,"     '��ϼ���
    sss = sss & "         ttd_lno,"     '������
    sss = sss & "         ttd_gbn,"     '1:��������/2:�׽�Ʈ����
    sss = sss & "         ttd_ryn,"     'Y/N - �������
    
    sss = sss & "         ttd_maker,"   '������
    sss = sss & "         ttd_tipstd,"  '�� �԰�/�ڵ�
    sss = sss & "         ttd_tipjil,"  '�� ����
    sss = sss & "         ttd_holder,"  '����Ȧ��
    sss = sss & "         ttd_rcntmn,"  '�д�ȸ�� �ּ�
    sss = sss & "         ttd_rcntmx,"  '�д�ȸ�� �ִ�
    sss = sss & "         ttd_movmn,"   '�̼�(MM/REV) �ּ�
    sss = sss & "         ttd_movmx,"   '�̼�(MM/REV) �ִ�
    sss = sss & "         ttd_tct,"     'TCT
    sss = sss & "         ttd_pct,"     'PCT
                          
    sss = sss & "         ttd_depth,"   '�������
                          
    sss = sss & "         ttd_fluid,"   '��������
    sss = sss & "         ttd_qty,"     '��������
    sss = sss & "         ttd_dan,"     '�ܰ�/EA
    sss = sss & "         ttd_tldn,"    '������/EA
    sss = sss & "         ttd_chdn,"    '��ȯ��/EA
    sss = sss & "         ttd_result,"  '��� OK/NG
    sss = sss & "         ttd_ryn1,"    'Y/N
    sss = sss & "         ttd_ryn2,"    'Y/N
    sss = sss & "         ttd_ryn3,"    'Y/N
    sss = sss & "         ttd_ryn4,"    'Y/N
    sss = sss & "         ttd_ryn5,"    'Y/N
    sss = sss & "         ttd_rtyn, "   '�� �׽�Ʈ ����
    sss = sss & "         ttd_rmk,"     '
    sss = sss & "         ttd_updte"    '
    
    sss = sss & ")values( "
    sss = sss & "        " & Dat & ","
    sss = sss & "        " & seq & ","
    sss = sss & "        2,"
    sss = sss & "        2,"
    If opt_ryn(2).Value = True Then
        sss = sss & "        'Y',"
    Else
        sss = sss & "        'N',"
    End If
                             
    sss = sss & "        '" & txt_maker2 & "',"
    sss = sss & "        '" & txt_tipstd2 & "',"
    sss = sss & "        '" & txt_tipjil2 & "',"
    sss = sss & "        '" & txt_holder2 & "',"
    
    If Val(txt_rcntmn2) > 0 Then    '�д�ȸ���� �ּҰ�
        sss = sss & "        " & Val(txt_rcntmn2) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_rcntmx2) > 0 Then    '�д�ȸ���� �ִ밪
        sss = sss & "        " & Val(txt_rcntmx2) & ","
    Else
        sss = sss & "        null,"
    End If
                                                        
    If Val(txt_movmn2) > 0 Then    '�̼� �ּҰ�
        sss = sss & "        " & Val(txt_movmn2) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_movmx2) > 0 Then    '�̼� �ִ밪
        sss = sss & "        " & Val(txt_movmx2) & ","
    Else
        sss = sss & "        null,"
    End If
                                               
    If Val(txt_tct2) > 0 Then      'TCP
        sss = sss & "        " & Val(txt_tct2) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_pct2) > 0 Then      'PCT
        sss = sss & "        " & Val(txt_pct2) & ","
    Else
        sss = sss & "        null,"
    End If
    
    If Val(txt_depth2) > 0 Then    '�������
        sss = sss & "        " & Val(txt_depth2) & ","
    Else
        sss = sss & "        null,"
    End If
    
    sss = sss & "        " & Left(cmb_fluid2, 1) & ","
    sss = sss & "        " & Val(txt_qty2) & ","
    sss = sss & "        " & Val(txt_dan2) & ","
    sss = sss & "        " & Val(txt_tldn2) & ","
    sss = sss & "        " & Val(txt_chdn2) & ","
                           
    If opt_ryn(2).Value = True Then
        
        sss = sss & "        '" & Replace(cmb_result1, ".", "") & "',"
        
        If chk_ryn1.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn2.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn3.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn4.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn5.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        sss = sss & "       '" & cmb_rtyn1 & "',"
            
    Else
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
    End If
    
    sss = sss & "        '',"
    sss = sss & "        to_char(sysdate,'yyyymmdd')"
    sss = sss & " )"
    
    db.Execute sss, 64
    
    
    '---------------------->
    '��TEST2��
    '---------------------->
    
    If chk_test2.Value = 1 Then
    
    
        sss = "insert into man_tooltestds("
        sss = sss & "         ttd_dat,"     '�������
        sss = sss & "         ttd_seq,"     '��ϼ���
        sss = sss & "         ttd_lno,"     '������
        sss = sss & "         ttd_gbn,"     '1:��������/2:�׽�Ʈ����
        sss = sss & "         ttd_ryn,"     'Y/N - �������
        
        sss = sss & "         ttd_maker,"   '������
        sss = sss & "         ttd_tipstd,"  '�� �԰�/�ڵ�
        sss = sss & "         ttd_tipjil,"  '�� ����
        sss = sss & "         ttd_holder,"  '����Ȧ��
        sss = sss & "         ttd_rcntmn,"  '�д�ȸ�� �ּ�
        sss = sss & "         ttd_rcntmx,"  '�д�ȸ�� �ִ�
        sss = sss & "         ttd_movmn,"   '�̼�(MM/REV) �ּ�
        sss = sss & "         ttd_movmx,"   '�̼�(MM/REV) �ִ�
        sss = sss & "         ttd_tct,"     'TCT
        sss = sss & "         ttd_pct,"     'PCT
                                            
        sss = sss & "         ttd_depth,"   '�������
                                                            
        sss = sss & "         ttd_fluid,"   '��������
        sss = sss & "         ttd_qty,"     '��������
        sss = sss & "         ttd_dan,"     '�ܰ�/EA
        sss = sss & "         ttd_tldn,"    '������/EA
        sss = sss & "         ttd_chdn,"    '��ȯ��/EA
        sss = sss & "         ttd_result,"  '��� OK/NG
        sss = sss & "         ttd_ryn1,"    'Y/N
        sss = sss & "         ttd_ryn2,"    'Y/N
        sss = sss & "         ttd_ryn3,"    'Y/N
        sss = sss & "         ttd_ryn4,"    'Y/N
        sss = sss & "         ttd_ryn5,"    'Y/N
        sss = sss & "         ttd_rtyn, "   '�� �׽�Ʈ ����
        sss = sss & "         ttd_rmk,"     '
        sss = sss & "         ttd_updte"    '
        
        sss = sss & ")values( "
        sss = sss & "        " & Dat & ","
        sss = sss & "        " & seq & ","
        sss = sss & "        3,"
        sss = sss & "        2,"
        If opt_ryn(3).Value = True Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
                       
        sss = sss & "        '" & txt_maker3 & "',"
        sss = sss & "        '" & txt_tipstd3 & "',"
        sss = sss & "        '" & txt_tipjil3 & "',"
        sss = sss & "        '" & txt_holder3 & "',"
        
        If Val(txt_rcntmn3) > 0 Then    '�д�ȸ���� �ּҰ�
            sss = sss & "        " & Val(txt_rcntmn3) & ","
        Else
            sss = sss & "        null,"
        End If
        If Val(txt_rcntmx3) > 0 Then    '�д�ȸ���� �ִ밪
            sss = sss & "        " & Val(txt_rcntmx3) & ","
        Else
            sss = sss & "        null,"
        End If
                                                            
        If Val(txt_movmn3) > 0 Then    '�̼� �ּҰ�
            sss = sss & "        " & Val(txt_movmn3) & ","
        Else
            sss = sss & "        null,"
        End If
        If Val(txt_movmx3) > 0 Then    '�̼� �ִ밪
            sss = sss & "        " & Val(txt_movmx3) & ","
        Else
            sss = sss & "        null,"
        End If
                                                   
        If Val(txt_tct3) > 0 Then      'TCP
            sss = sss & "        " & Val(txt_tct3) & ","
        Else
            sss = sss & "        null,"
        End If
        If Val(txt_pct3) > 0 Then      'PCT
            sss = sss & "        " & Val(txt_pct3) & ","
        Else
            sss = sss & "        null,"
        End If
        
        If Val(txt_depth3) > 0 Then    '�������
            sss = sss & "        " & Val(txt_depth3) & ","
        Else
            sss = sss & "        null,"
        End If
    
        sss = sss & "        " & Left(cmb_fluid3, 1) & ","
        sss = sss & "        " & Val(txt_qty3) & ","
        sss = sss & "        " & Val(txt_dan3) & ","
        sss = sss & "        " & Val(txt_tldn3) & ","
        sss = sss & "        " & Val(txt_chdn3) & ","
        
        If opt_ryn(3).Value = True Then
            
            sss = sss & "        '" & Replace(cmb_result1, ".", "") & "',"
            
            If chk_ryn1.Value = 1 Then
                sss = sss & "        'Y',"
            Else
                sss = sss & "        'N',"
            End If
            If chk_ryn2.Value = 1 Then
                sss = sss & "        'Y',"
            Else
                sss = sss & "        'N',"
            End If
            If chk_ryn3.Value = 1 Then
                sss = sss & "        'Y',"
            Else
                sss = sss & "        'N',"
            End If
            If chk_ryn4.Value = 1 Then
                sss = sss & "        'Y',"
            Else
                sss = sss & "        'N',"
            End If
            If chk_ryn5.Value = 1 Then
                sss = sss & "        'Y',"
            Else
                sss = sss & "        'N',"
            End If
            sss = sss & "       '" & cmb_rtyn1 & "',"
                
        Else
            sss = sss & "       null,"
            sss = sss & "       null,"
            sss = sss & "       null,"
            sss = sss & "       null,"
            sss = sss & "       null,"
            sss = sss & "       null,"
            sss = sss & "       null,"
        End If
    
        sss = sss & "        '',"
        sss = sss & "        to_char(sysdate,'yyyymmdd')"
        sss = sss & " )"
        
        db.Execute sss, 64
    
    End If
    '
    '������ �̽�ȯ�븮��
    sss = "insert into oth_applist("
    sss = sss & "apl_table,"
    sss = sss & "apl_tdat,"
    sss = sss & "apl_tseq,"
    sss = sss & "apl_isab,"
    sss = sss & "apl_1yn,"
    sss = sss & "apl_2yn,"
    sss = sss & "apl_3yn,"
    sss = sss & "apl_4yn,"
    sss = sss & "apl_1sab,"
    sss = sss & "apl_2sab,"
    sss = sss & "apl_3sab,"
    sss = sss & "apl_4sab,"
    sss = sss & "apl_1dat,"
    sss = sss & "apl_2dat,"
    sss = sss & "apl_3dat,"
    sss = sss & "apl_4dat,"
    sss = sss & "apl_allsab,"
    sss = sss & "apl_allyn,"
    sss = sss & "apl_alldat,"
    
    sss = sss & "apl_m1yn,"
    sss = sss & "apl_m1sab,"
    sss = sss & "apl_m1dat,"
    
    sss = sss & "apl_m2yn,"
    sss = sss & "apl_m2sab,"
    sss = sss & "apl_m2dat,"
    
    sss = sss & "apl_m3yn,"
    sss = sss & "apl_m3sab,"
    sss = sss & "apl_m3dat,"
    
    sss = sss & "apl_m4yn,"
    sss = sss & "apl_m4sab,"
    sss = sss & "apl_m4dat"
    
    sss = sss & ") values ("
    sss = sss & "'man_tooltesthd'" & ","
    sss = sss & Dat & ","
    sss = sss & seq & ","
    sss = sss & Gsab & ","
    sss = sss & "'N','N','N','N',"
    sss = sss & "0,0,0,3,0,0,0,0,"
    sss = sss & "0,'N',0,"
    sss = sss & "'N',0,0,"
    sss = sss & "'N',0,0,"
    sss = sss & "'N',0,0,"
    sss = sss & "'N',0,0"
    sss = sss & ")"
    '
    db.Execute sss, 64
    '
    '
    Ws.CommitTrans
    '
    '
    txt_dat1 = Dat: txt_dat1.Enabled = False
    txt_seq1 = seq: txt_seq1.Enabled = False
    '
    Call file_view
    '
    spd_file11.Visible = False
    spd_file12.Visible = True
    lbl_cmt1.Visible = True
    '
    btn_add1.Enabled = False
    btn_mod1.Enabled = True
    btn_del1.Enabled = True
    btn_prt1.Enabled = True
                     
    txt_dat.Enabled = False
    txt_seq.Enabled = False
    opt_n_ja.Value = False
    opt_n_su.Value = False
    opt_n_ja.Enabled = False
    opt_n_su.Enabled = False
    '
    Call msg_display("��ϵǾ����ϴ�")
    '
    Exit Sub
    '
err_rtn:
    Ws.Rollback
    MsgBox (Err.Description)
End Sub

'��ϳ��� üũ
Public Function Check_Insert_Data()
    
    Dim str As String
    
    'TEST NO.Ȯ��
    If Len(Trim(txt_testno1)) < 1 Then
        Call msg_display("TEST NO.�� Ȯ���ϼ���! (��Ͻ���)")
        txt_testno1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    If LenB(StrConv(txt_testno1, vbFromUnicode)) > 20 Then
        Call msg_display("TEST NO. ���ڱ��̸� Ȯ���ϼ���! (20byte����)")
        txt_testno1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '�׽�Ʈ ����Ȯ��
    If IsDate(Left(txt_tdat1, 4) & "/" & Mid(txt_tdat1, 5, 2) & "/" & Right(txt_tdat1, 2)) = False Then
        Call msg_display("�׽�Ʈ ���ڸ� Ȯ���ϼ���! (��Ͻ���)")
        txt_tdat1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '������ȣ Ȯ��
    If IsDate(Left(txt_jdat1, 4) & "/" & Mid(txt_jdat1, 5, 2) & "/" & Right(txt_jdat1, 2)) = False Then
        Call msg_display("�������ڸ� Ȯ���ϼ���! (��Ͻ���)")
        txt_jdat1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '�������� Ȯ��
    If Len(Trim(txt_jseq1)) < 1 Or IsNumeric(txt_jseq1) = False Then
        Call msg_display("���������� Ȯ���ϼ���! (��Ͻ���)")
        txt_jseq1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    'TEST���� Ȯ��
    If Len(Trim(txt_title1)) < 1 Then
        Call msg_display("TEST������ �Է��ϼ���! (��Ͻ���)")
        txt_title1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    If LenB(StrConv(txt_title1, vbFromUnicode)) > 30 Then
        Call msg_display("TEST���� ���ڱ��̸� Ȯ���ϼ���! (30byte����)")
        txt_title1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '����ڵ�Ȯ��
    If Len(Trim(txt_mcd1)) < 1 Then
        Call msg_display("TEST ����ڵ带 �Է��ϼ���! (��Ͻ���)")
        txt_mcd1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '����üũ
    sss = "       select mhc_name, ems_mark, sok_name(mhc_sok) soknm "
    sss = sss & "   from man_machcd, eam_mast"
    sss = sss & "  where mhc_code = '" & Trim(txt_mcd1) & "'"
    sss = sss & "    and mhc_code = ems_mcd(+)"
                
    Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Ks.RecordCount < 1 Then
        Call msg_display("TEST ����ڵ带 Ȯ���ϼ���! (��Ͻ���)")
        txt_mcd1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
        
    Ks.Close
    
    'LOT NOȮ��
    If Len(Trim(txt_lotno1)) < 1 Then
        Call msg_display("LOT NO.�� Ȯ���ϼ���! (��Ͻ���)")
        txt_lotno1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '����üũ
    sss = "       select dit_bpcd, dit_bpjil, dit_jacd, dit_jajil"
    sss = sss & "   from man_direct"
    sss = sss & "  where dit_lot = '" & Trim(txt_lotno1) & "'"
            
    Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Ks.RecordCount < 1 Then
        Call msg_display("LOT NO.�� Ȯ���ϼ���! (��Ͻ���)")
        txt_lotno1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    Ks.Close
        
    '���Ȯ��
    '�������� Ȯ��
    If Len(Trim(txt_tsab1)) < 1 Or IsNumeric(txt_tsab1) = False Then
        Call msg_display("�۾��ڻ���� Ȯ���ϼ���! (��Ͻ���)")
        txt_tsab1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    sss = "       select sin_name, sin_sok, sok_name(sin_sok) soknm"
    sss = sss & "   from peo_sinbun"
    sss = sss & "  where sin_sab = " & Val(txt_tsab1)
    'sss = sss & "    and sin_taedt = 0 "
            
    Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Ks.RecordCount < 1 Then
        Call msg_display("����� Ȯ���ϼ���! (��Ͻ���)")
        txt_lotno1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    Ks.Close
    
    '=========================
    '����� ���������� Ȯ�Ρ�
    '=========================
    
    '------------------------>
    '��������
    '------------------------>
    'MAKER
    If Len(Trim(txt_maker1)) < 1 Then
        Call msg_display("[��������] MAKER�� �Է��ϼ���! (��Ͻ���)")
        txt_maker1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_maker1, vbFromUnicode)) > 20 Then
        Call msg_display("[��������] MAKER ���ڱ��̸� Ȯ���ϼ���! (20byte����)")
        txt_maker1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    'TIP�԰�/�ڵ�
    If Len(Trim(txt_tipstd1)) < 1 Then
        Call msg_display("[��������] TIP�԰��� �Է��ϼ���! (��Ͻ���)")
        txt_tipstd1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_tipstd1, vbFromUnicode)) > 30 Then
        Call msg_display("[��������] TIP�԰� ���ڱ��̸� Ȯ���ϼ���! (30byte����)")
        txt_tipstd1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    'TIP����
    If Len(Trim(txt_tipjil1)) < 1 Then
        Call msg_display("[��������] TIP������ �Է��ϼ���! (��Ͻ���)")
        txt_tipjil1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_tipjil1, vbFromUnicode)) > 10 Then
        Call msg_display("[��������] TIP���� ���ڱ��̸� Ȯ���ϼ���! (10byte����)")
        txt_tipjil1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    'HOLDER
    If Len(Trim(txt_holder1)) < 1 Then
        Call msg_display("[��������] HOLDER �ڵ带 �Է��ϼ���! (��Ͻ���)")
        txt_holder1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_holder1, vbFromUnicode)) > 30 Then
        Call msg_display("[��������] HOLDER �ڵ� ���ڱ��̸� Ȯ���ϼ���! (20byte����)")
        txt_holder1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '�д�ȸ���� �ּҰ�
    If Len(Trim(txt_rcntmn1)) > 0 Then
        If IsNumeric(txt_rcntmn1) = False Then
            Call msg_display("�д�ȸ���� �ּҰ��� ���ڸ� �Է��ϼ���! ")
            txt_rcntmn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_rcntmn1) > 9999 Or Val(txt_rcntmn1) < 0 Then
            Call msg_display("�д�ȸ���� �ּҰ��� 1~9999 ������ �Է°����մϴ�! ")
            txt_rcntmn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '�д�ȸ���� �ִ밪
    If Len(Trim(txt_rcntmx1)) > 0 Then
        If IsNumeric(txt_rcntmx1) = False Then
            Call msg_display("�д�ȸ���� �ִ밪�� ���ڸ� �Է��ϼ���! ")
            txt_rcntmx1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_rcntmx1) > 9999 Or Val(txt_rcntmx1) < 0 Then
            Call msg_display("�д�ȸ���� �ִ밪�� 1~9999 ������ �Է°����մϴ�! ")
            txt_rcntmx1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If

    '�������
    If Len(Trim(txt_depth1)) > 0 Then
        If IsNumeric(txt_depth1) = False Then
            Call msg_display("������̸� �Է��ϼ���! ")
            txt_depth1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_depth1) > 99 Or Val(txt_depth1) < 0 Then
            Call msg_display("������� ���� 0.01~99.99 ������ �Է°����մϴ�! ")
            txt_depth1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
        
    '�̼� �ּҰ�
    If Len(Trim(txt_movmn1)) > 0 Then
        If IsNumeric(txt_movmn1) = False Then
            Call msg_display("�̼ۼ� �ּҰ��� ���ڸ� �Է��ϼ���! ")
            txt_movmn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_movmn1) > 9999 Or Val(txt_movmn1) < 0 Then
            Call msg_display("�̼� �ּҰ��� 0.01~9999 ������ �Է°����մϴ�! ")
            txt_movmn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '�̼� �ִ밪
    If Len(Trim(txt_movmx1)) > 0 Then
        If IsNumeric(txt_movmx1) = False Then
            Call msg_display("�̼� �ִ밪�� ���ڸ� �Է��ϼ���! ")
            txt_movmx1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_movmx1) > 9999 Or Val(txt_movmx1) < 0 Then
            Call msg_display("�̼� �ִ밪�� 0.01~9999 ������ �Է°����մϴ�! ")
            txt_movmx1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    'TCP
    If Len(Trim(txt_tct1)) > 0 Then
        If IsNumeric(txt_tct1) = False Then
            Call msg_display("T.C/P���� ���ڸ� �Է��ϼ���! ")
            txt_tct1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_tct1) > 99999 Or Val(txt_tct1) < 0 Then
            Call msg_display("T.C/P���� 1~99999 ������ �Է°����մϴ�! ")
            txt_tct1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    'PCT
    If Len(Trim(txt_pct1)) > 0 Then
        If IsNumeric(txt_pct1) = False Then
            Call msg_display("P.C/P���� ���ڸ� �Է��ϼ���! ")
            txt_pct1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_pct1) > 99999 Or Val(txt_pct1) < 0 Then
            Call msg_display("P.C/P����  1~99999 ������ �Է°����մϴ�! ")
            txt_pct1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '�������� Ȯ��
    If Left(cmb_fluid1, 1) <> 1 And Left(cmb_fluid1, 2) <> 2 Then
        Call msg_display("���������� �����ϼ���! ")
        cmb_fluid1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '������Ȯ��
    If Len(Trim(txt_qty1)) > 0 Then
        If IsNumeric(txt_qty1) = False Then
            Call msg_display("�������� ���ڸ� �Է��ϼ���! ")
            txt_qty1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_qty1) > 99999 Or Val(txt_qty1) < 0 Then
            Call msg_display("��������  1~99999 ������ �Է°����մϴ�! ")
            txt_qty1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '�����ܰ�/EA
    If Len(Trim(txt_dan1)) > 0 Then
        If IsNumeric(txt_dan1) = False Then
            Call msg_display("�����ܰ��� ���ڸ� �Է��ϼ���! ")
            txt_dan1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_dan1) > 99999999 Or Val(txt_dan1) < 0 Then
            Call msg_display("�����ܰ���  1~99999999 ������ �Է°����մϴ�! ")
            txt_dan1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '������/EA
    If Len(Trim(txt_tldn1)) > 0 Then
        If IsNumeric(txt_tldn1) = False Then
            Call msg_display("������� ���ڸ� �Է��ϼ���! ")
            txt_tldn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_tldn1) > 99999999 Or Val(txt_tldn1) < 0 Then
            Call msg_display("�������  1~99999999 ������ �Է°����մϴ�! ")
            txt_tldn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '������ȯ��
    If Len(Trim(txt_chdn1)) > 0 Then
        If IsNumeric(txt_chdn1) = False Then
            Call msg_display("��ȯ��� ���ڸ� �Է��ϼ���! ")
            txt_chdn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_chdn1) > 99999 Or Val(txt_chdn1) < 0 Then
            Call msg_display("��ȯ���  1~99999 ������ �Է°����մϴ�! ")
            txt_chdn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '------------------------>
    '���׽�Ʈ1
    '------------------------>
    'MAKER
    If Len(Trim(txt_maker2)) < 1 Then
        Call msg_display("[��������] MAKER�� �Է��ϼ���! (��Ͻ���)")
        txt_maker2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_maker2, vbFromUnicode)) > 20 Then
        Call msg_display("[��������] MAKER ���ڱ��̸� Ȯ���ϼ���! (20byte����)")
        txt_maker2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    'TIP�԰�/�ڵ�
    If Len(Trim(txt_tipstd2)) < 1 Then
        Call msg_display("[��������] TIP�԰��� �Է��ϼ���! (��Ͻ���)")
        txt_tipstd2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_tipstd2, vbFromUnicode)) > 30 Then
        Call msg_display("[��������] TIP�԰� ���ڱ��̸� Ȯ���ϼ���! (30byte����)")
        txt_tipstd2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    'TIP����
    If Len(Trim(txt_tipjil2)) < 1 Then
        Call msg_display("[��������] TIP������ �Է��ϼ���! (��Ͻ���)")
        txt_tipjil2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_tipjil2, vbFromUnicode)) > 10 Then
        Call msg_display("[��������] TIP���� ���ڱ��̸� Ȯ���ϼ���! (10byte����)")
        txt_tipjil2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    'HOLDER
    If Len(Trim(txt_holder2)) < 1 Then
        Call msg_display("[��������] HOLDER �ڵ带 �Է��ϼ���! (��Ͻ���)")
        txt_holder2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_holder2, vbFromUnicode)) > 30 Then
        Call msg_display("[��������] HOLDER �ڵ� ���ڱ��̸� Ȯ���ϼ���! (20byte����)")
        txt_holder2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '�д�ȸ���� �ּҰ�
    If Len(Trim(txt_rcntmn2)) > 0 Then
        If IsNumeric(txt_rcntmn2) = False Then
            Call msg_display("�д�ȸ���� �ּҰ��� ���ڸ� �Է��ϼ���! ")
            txt_rcntmn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_rcntmn2) > 9999 Or Val(txt_rcntmn2) < 0 Then
            Call msg_display("�д�ȸ���� �ּҰ��� 1~9999 ������ �Է°����մϴ�! ")
            txt_rcntmn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '�д�ȸ���� �ִ밪
    If Len(Trim(txt_rcntmx2)) > 0 Then
        If IsNumeric(txt_rcntmx2) = False Then
            Call msg_display("�д�ȸ���� �ִ밪�� ���ڸ� �Է��ϼ���! ")
            txt_rcntmx2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_rcntmx2) > 9999 Or Val(txt_rcntmx2) < 0 Then
            Call msg_display("�д�ȸ���� �ִ밪�� 1~9999 ������ �Է°����մϴ�! ")
            txt_rcntmx2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '�������
    If Len(Trim(txt_depth2)) > 0 Then
        If IsNumeric(txt_depth2) = False Then
            Call msg_display("������̸� �Է��ϼ���! ")
            txt_depth2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_depth2) > 99 Or Val(txt_depth2) < 0 Then
            Call msg_display("������� ���� 0.01~99.99 ������ �Է°����մϴ�! ")
            txt_depth2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '�̼� �ּҰ�
    If Len(Trim(txt_movmn2)) > 0 Then
        If IsNumeric(txt_movmn2) = False Then
            Call msg_display("�̼ۼ� �ּҰ��� ���ڸ� �Է��ϼ���! ")
            txt_movmn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_movmn2) > 9999 Or Val(txt_movmn2) < 0 Then
            Call msg_display("�̼� �ּҰ��� 0.01~9999 ������ �Է°����մϴ�! ")
            txt_movmn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '�̼� �ִ밪
    If Len(Trim(txt_movmx2)) > 0 Then
        If IsNumeric(txt_movmx2) = False Then
            Call msg_display("�̼� �ִ밪�� ���ڸ� �Է��ϼ���! ")
            txt_movmx2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_movmx2) > 9999 Or Val(txt_movmx2) < 0 Then
            Call msg_display("�̼� �ִ밪�� 0.01~9999 ������ �Է°����մϴ�! ")
            txt_movmx2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    'TCP
    If Len(Trim(txt_tct2)) > 0 Then
        If IsNumeric(txt_tct2) = False Then
            Call msg_display("T.C/P���� ���ڸ� �Է��ϼ���! ")
            txt_tct2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_tct2) > 99999 Or Val(txt_tct2) < 0 Then
            Call msg_display("T.C/P���� 1~99999 ������ �Է°����մϴ�! ")
            txt_tct2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    'PCT
    If Len(Trim(txt_pct2)) > 0 Then
        If IsNumeric(txt_pct2) = False Then
            Call msg_display("P.C/P���� ���ڸ� �Է��ϼ���! ")
            txt_pct2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_pct2) > 99999 Or Val(txt_pct2) < 0 Then
            Call msg_display("P.C/P����  1~99999 ������ �Է°����մϴ�! ")
            txt_pct2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '�������� Ȯ��
    If Left(cmb_fluid2, 1) <> 1 And Left(cmb_fluid2, 2) <> 2 Then
        Call msg_display("���������� �����ϼ���! ")
        cmb_fluid2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '������Ȯ��
    If Len(Trim(txt_qty2)) > 0 Then
        If IsNumeric(txt_qty2) = False Then
            Call msg_display("������ ���� ���ڸ� �Է��ϼ���! ")
            txt_qty2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_qty2) > 99999 Or Val(txt_qty2) < 0 Then
            Call msg_display("��������  1~99999 ������ �Է°����մϴ�! ")
            txt_qty2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
                            
    '�����ܰ�/EA
    If Len(Trim(txt_dan2)) > 0 Then
        If IsNumeric(txt_dan2) = False Then
            Call msg_display("�����ܰ��� ���ڸ� �Է��ϼ���! ")
            txt_dan2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_dan2) > 99999999 Or Val(txt_dan2) < 0 Then
            Call msg_display("�����ܰ���  1~99999999 ������ �Է°����մϴ�! ")
            txt_dan2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
        
    '������/EA
    If Len(Trim(txt_tldn2)) > 0 Then
        If IsNumeric(txt_tldn1) = False Then
            Call msg_display("������� ���ڸ� �Է��ϼ���! ")
            txt_tldn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_tldn2) > 99999999 Or Val(txt_tldn2) < 0 Then
            Call msg_display("�������  1~99999999 ������ �Է°����մϴ�! ")
            txt_tldn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
                                
    '������ȯ��
    If Len(Trim(txt_chdn2)) > 0 Then
        If IsNumeric(txt_chdn2) = False Then
            Call msg_display("��ȯ��� ���ڸ� �Է��ϼ���! ")
            txt_chdn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_chdn2) > 99999 Or Val(txt_chdn2) < 0 Then
            Call msg_display("��ȯ���  1~99999 ������ �Է°����մϴ�! ")
            txt_chdn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '------------------------>
    '�׽�Ʈ2
    '------------------------>
    
    If chk_test2.Value = 1 Then
    
        'MAKER
        If Len(Trim(txt_maker3)) < 1 Then
            Call msg_display("[��������] MAKER�� �Է��ϼ���! (��Ͻ���)")
            txt_maker3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If LenB(StrConv(txt_maker3, vbFromUnicode)) > 20 Then
            Call msg_display("[��������] MAKER ���ڱ��̸� Ȯ���ϼ���! (20byte����)")
            txt_maker3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        
        'TIP�԰�/�ڵ�
        If Len(Trim(txt_tipstd3)) < 1 Then
            Call msg_display("[��������] TIP�԰��� �Է��ϼ���! (��Ͻ���)")
            txt_tipstd3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If LenB(StrConv(txt_tipstd3, vbFromUnicode)) > 30 Then
            Call msg_display("[��������] TIP�԰� ���ڱ��̸� Ȯ���ϼ���! (30byte����)")
            txt_tipstd3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        
        'TIP����
        If Len(Trim(txt_tipjil3)) < 1 Then
            Call msg_display("[��������] TIP������ �Է��ϼ���! (��Ͻ���)")
            txt_tipjil3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If LenB(StrConv(txt_tipjil3, vbFromUnicode)) > 10 Then
            Call msg_display("[��������] TIP���� ���ڱ��̸� Ȯ���ϼ���! (10byte����)")
            txt_tipjil3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        
        'HOLDER
        If Len(Trim(txt_holder3)) < 1 Then
            Call msg_display("[��������] HOLDER �ڵ带 �Է��ϼ���! (��Ͻ���)")
            txt_holder3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If LenB(StrConv(txt_holder3, vbFromUnicode)) > 30 Then
            Call msg_display("[��������] HOLDER �ڵ� ���ڱ��̸� Ȯ���ϼ���! (20byte����)")
            txt_holder3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        
        '�д�ȸ���� �ּҰ�
        If Len(Trim(txt_rcntmn3)) > 0 Then
            If IsNumeric(txt_rcntmn3) = False Then
                Call msg_display("�д�ȸ���� �ּҰ��� ���ڸ� �Է��ϼ���! ")
                txt_rcntmn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_rcntmn3) > 9999 Or Val(txt_rcntmn3) < 0 Then
                Call msg_display("�д�ȸ���� �ּҰ��� 1~9999 ������ �Է°����մϴ�! ")
                txt_rcntmn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
        
        '�д�ȸ���� �ִ밪
        If Len(Trim(txt_rcntmx3)) > 0 Then
            If IsNumeric(txt_rcntmx3) = False Then
                Call msg_display("�д�ȸ���� �ִ밪�� ���ڸ� �Է��ϼ���! ")
                txt_rcntmx3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_rcntmx3) > 9999 Or Val(txt_rcntmx3) < 0 Then
                Call msg_display("�д�ȸ���� �ִ밪�� 1~9999 ������ �Է°����մϴ�! ")
                txt_rcntmx3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
        
        '�������
        If Len(Trim(txt_depth3)) > 0 Then
            If IsNumeric(txt_depth3) = False Then
                Call msg_display("������̸� �Է��ϼ���! ")
                txt_depth3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_depth3) > 99 Or Val(txt_depth3) < 0 Then
                Call msg_display("������� ���� 0.01~99.99 ������ �Է°����մϴ�! ")
                txt_depth3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
            
        '�̼� �ּҰ�
        If Len(Trim(txt_movmn3)) > 0 Then
            If IsNumeric(txt_movmn3) = False Then
                Call msg_display("�̼ۼ� �ּҰ��� ���ڸ� �Է��ϼ���! ")
                txt_movmn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_movmn3) > 9999 Or Val(txt_movmn3) < 0 Then
                Call msg_display("�̼� �ּҰ��� 0.01~9999 ������ �Է°����մϴ�! ")
                txt_movmn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
                                                        
        '�̼� �ִ밪
        If Len(Trim(txt_movmx3)) > 0 Then
            If IsNumeric(txt_movmx3) = False Then
                Call msg_display("�̼� �ִ밪�� ���ڸ� �Է��ϼ���! ")
                txt_movmx3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_movmx3) > 9999 Or Val(txt_movmx3) < 0 Then
                Call msg_display("�̼� �ִ밪�� 0.01~99.99 ������ �Է°����մϴ�! ")
                txt_movmx3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
        
        'TCP
        If Len(Trim(txt_tct3)) > 0 Then
            If IsNumeric(txt_tct3) = False Then
                Call msg_display("T.C/P���� ���ڸ� �Է��ϼ���! ")
                txt_tct3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_tct3) > 99999 Or Val(txt_tct3) < 0 Then
                Call msg_display("T.C/P���� 1~99999 ������ �Է°����մϴ�! ")
                txt_tct3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
        
        'PCT
        If Len(Trim(txt_pct3)) > 0 Then
            If IsNumeric(txt_pct3) = False Then
                Call msg_display("P.C/P���� ���ڸ� �Է��ϼ���! ")
                txt_pct3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_pct3) > 99999 Or Val(txt_pct3) < 0 Then
                Call msg_display("P.C/P����  1~99999 ������ �Է°����մϴ�! ")
                txt_pct3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
        
        '�������� Ȯ��
        If Left(cmb_fluid3, 1) <> 1 And Left(cmb_fluid3, 1) <> 2 Then
            Call msg_display("���������� �����ϼ���! ")
            cmb_fluid3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        
        '������Ȯ��
        If Len(Trim(txt_qty3)) > 0 Then
            If IsNumeric(txt_qty2) = False Then
                Call msg_display("�������� ���ڸ� �Է��ϼ���! ")
                txt_qty3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_qty3) > 99999 Or Val(txt_qty3) < 0 Then
                Call msg_display("��������  1~99999 ������ �Է°����մϴ�! ")
                txt_qty3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
        
        '�����ܰ�/EA
        If Len(Trim(txt_dan3)) > 0 Then
            If IsNumeric(txt_dan3) = False Then
                Call msg_display("�����ܰ��� ���ڸ� �Է��ϼ���! ")
                txt_dan3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_dan3) > 99999999 Or Val(txt_dan3) < 0 Then
                Call msg_display("�����ܰ���  1~99999999 ������ �Է°����մϴ�! ")
                txt_dan3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
        
        '������/EA
        If Len(Trim(txt_tldn3)) > 0 Then
            If IsNumeric(txt_tldn3) = False Then
                Call msg_display("������� ���ڸ� �Է��ϼ���! ")
                txt_tldn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_tldn3) > 99999999 Or Val(txt_tldn3) < 0 Then
                Call msg_display("�������  1~99999999 ������ �Է°����մϴ�! ")
                txt_tldn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
                
        '������ȯ��
        If Len(Trim(txt_chdn3)) > 0 Then
            If IsNumeric(txt_chdn3) = False Then
                Call msg_display("��ȯ��� ���ڸ� �Է��ϼ���! ")
                txt_chdn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_chdn3) > 99999 Or Val(txt_chdn3) < 0 Then
                Call msg_display("��ȯ���  1~99999 ������ �Է°����մϴ�! ")
                txt_chdn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
    
    End If
    
    
    '============================
    '���
    '============================
    
    If cmb_result1 <> "O.K" And cmb_result1 <> "N.G" Then
        Call msg_display("����� O.K �Ǵ� N.G�� �����ϼ���!")
        cmb_result1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    If cmb_result1 = "N.G" Then
        
        If opt_ryn(1).Value <> True Then
            Call msg_display("����� N.G�϶� ���� ������ ���� �� ����ϼ���!")
            'opt_ryn(1).SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    
    ElseIf cmb_result1 = "O.K" Then
        
        If opt_ryn(2).Value = False And opt_ryn(3).Value = False Then
            Call msg_display("����� O.K�϶� �׽�Ʈ ������ ���� �� ����ϼ���!")
            'opt_ryn(2).SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        
        
    End If
    
    If chk_ryn1.Value = 0 And chk_ryn2.Value = 0 And chk_ryn3.Value = 0 And _
       chk_ryn4.Value = 0 And chk_ryn5.Value = 0 Then
        Call msg_display("��� ������ �ּ��� 1���̻� �����ϼ���!")
        cmb_result1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    If cmb_rtyn1 <> "Y" And cmb_rtyn1 <> "N" Then
        Call msg_display("�� TEST���ɿ��δ� ""Y"" �Ǵ� ""N""�� �����ϼ���!")
        cmb_rtyn1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '�������
    If opt_ryn(1).Value = False And opt_ryn(2).Value = False And opt_ryn(3).Value = False Then
        Call msg_display("��� ��������� �����ϼ���! (����� ����)")
        Check_Insert_Data = 9
        Exit Function
    End If
    
    sprd_rmk1.row = 1
    sprd_rmk1.Col = 1
    If LenB(StrConv(sprd_rmk1.Text, vbFromUnicode)) > 200 Then
        Call msg_display("������� ���ڱ��̰� ��ϴ�.! (200byte����)")
        sprd_rmk1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    sprd_rmk1.Col = 2
    If LenB(StrConv(sprd_rmk1.Text, vbFromUnicode)) > 200 Then
        Call msg_display("�򰡳��� ������ ���ڱ��̰� ��ϴ�.! (200byte����)")
        sprd_rmk1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
End Function

'-------------------------------
'��������
'-------------------------------
Private Sub btn_mod1_Click()
    '
    On Error GoTo err_rtn
    '
    If Job_Level < 9 Then
        MsgBox ("�۾������� �����ϴ�!")
        Exit Sub
    End If
    
    If IsNumeric(txt_dat1) = False Or Len(txt_dat1) <> 8 Then
        Call msg_display("��� ���ڸ� Ȯ���ϼ���!")
        txt_dat1.SetFocus
        Exit Sub
    End If
    '
    If IsNumeric(txt_seq1) = False Or Len(txt_seq1) < 1 Then
        Call msg_display("��� ������ Ȯ���ϼ���!")
        txt_seq1.SetFocus
        Exit Sub
    End If
    '
    sss = "select * from man_tooltesthd, man_tooltestds"
    sss = sss & " where tth_dat = ttd_dat"
    sss = sss & "   and tth_seq = ttd_seq"
    sss = sss & "   and tth_dat = " & txt_dat1
    sss = sss & "   and tth_seq = " & txt_seq1
    sss = sss & "   order by ttd_lno"
    '
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Rs.RecordCount < 1 Then
       Rs.Close
       Call msg_display("��ϵ� ������ �����ϴ�. �������!")
       Exit Sub
    End If
    '
    If Check_Insert_Data = 9 Then Exit Sub
    '
    '����
    If MsgBox("(" & txt_dat1 & "-" & Format(txt_seq1, "000") & ") ��ϵ� ������ �����Ͻðڽ��ϱ�?", vbYesNo) <> vbYes Then
        Rs.Close
        msg_display ("���� ��ҵǾ����ϴ�.")
        Exit Sub
    End If
    '
    '
    Ws.BeginTrans
    '
    '
    sss = "update man_tooltesthd set"
    '
    sss = sss & "         tth_testno = '" & txt_testno1 & "',"                        'TEST NO.
    sss = sss & "         tth_title = '" & txt_title1 & "',"                          '�׽�Ʈ ����
    sss = sss & "         tth_tlot = '" & txt_lotno1 & "',"                           '�׽�Ʈ ���� LOT No.
    sss = sss & "         tth_tmcd = '" & txt_mcd1 & "',"                             '�׽�Ʈ ����ڵ�
    sss = sss & "         tth_tsab = " & Val(txt_tsab1) & ","                         '�׽�Ʈ �۾���
    sss = sss & "         tth_tdat = " & Val(txt_tdat1) & ","                           '�׽�Ʈ ����
    sss = sss & "         tth_jubno = " & txt_jdat1 & Format(txt_jseq1, "000") & ","   '��������
                    
    If chk_pyn1.Value = 1 Then
        sss = sss & "         tth_pyn1 = 'Y',"         'Y/N - 1.��������
    Else
        sss = sss & "         tth_pyn1 = 'N',"         'Y/N - 1.��������
    End If
    If chk_pyn2.Value = 1 Then
        sss = sss & "         tth_pyn2 = 'Y',"         'Y/N - 2.Ĩ ó��
    Else
        sss = sss & "         tth_pyn2 = 'N',"         'Y/N - 2.Ĩ ó��
    End If
    If chk_pyn3.Value = 1 Then
        sss = sss & "         tth_pyn3 = 'Y',"         'Y/N - 3.�ð� ����
    Else
        sss = sss & "         tth_pyn3 = 'N',"         'Y/N - 3.�ð� ����
    End If
    If chk_pyn4.Value = 1 Then
        sss = sss & "         tth_pyn4 = 'Y',"         'Y/N - 4.������ ����
    Else
        sss = sss & "         tth_pyn4 = 'N',"         'Y/N - 4.������ ����
    End If
    If chk_pyn5.Value = 1 Then
        sss = sss & "         tth_pyn5 = 'Y',"         'Y/N - 5.��Ÿ
    Else
        sss = sss & "         tth_pyn5 = 'N',"         'Y/N - 5.��Ÿ
    End If

    sss = sss & "         tth_sab = " & Val(txt_tsab1) & ","        '�Է���
    sprd_rmk1.row = 1
    sprd_rmk1.Col = 1
    sss = sss & "         tth_rmk = '" & sprd_rmk1.Text & "',"      '���
    sprd_rmk1.Col = 2
    sss = sss & "         tth_cmt = '" & sprd_rmk1.Text & "',"      '��
    
    sss = sss & "         tth_updte = to_char(sysdate,'yyyymmdd')"  '��������
                          
    sss = sss & "   where tth_dat = " & txt_dat1
    sss = sss & "     and tth_seq = " & txt_seq1
    
    db.Execute sss, 64
    
    'DESC ������ �ٽ� �����.
    sss = "       delete from man_tooltestds"
    sss = sss & "  where ttd_dat = " & txt_dat1
    sss = sss & "    and ttd_seq = " & txt_seq1
    
    db.Execute sss, 64
    
    '=======================
    'DESC�Է�
    '=======================
    '---------------------->
    '�����������
    '---------------------->
    sss = "insert into man_tooltestds("
    sss = sss & "         ttd_dat,"     '�������
    sss = sss & "         ttd_seq,"     '��ϼ���
    sss = sss & "         ttd_lno,"     '������
    sss = sss & "         ttd_gbn,"     '1:��������/2:�׽�Ʈ����
    sss = sss & "         ttd_ryn,"     'Y/N - �������
    
    sss = sss & "         ttd_maker,"   '������
    sss = sss & "         ttd_tipstd,"  '�� �԰�/�ڵ�
    sss = sss & "         ttd_tipjil,"  '�� ����
    sss = sss & "         ttd_holder,"  '����Ȧ��
    sss = sss & "         ttd_rcntmn,"  '�д�ȸ�� �ּ�
    sss = sss & "         ttd_rcntmx,"  '�д�ȸ�� �ִ�
    sss = sss & "         ttd_movmn,"   '�̼�(MM/REV) �ּ�
    sss = sss & "         ttd_movmx,"   '�̼�(MM/REV) �ִ�
    sss = sss & "         ttd_tct,"     'TCT
    sss = sss & "         ttd_pct,"     'PCT
                        
    sss = sss & "         ttd_depth,"   '�������
    
    sss = sss & "         ttd_fluid,"   '��������
    sss = sss & "         ttd_qty,"     '��������
    sss = sss & "         ttd_dan,"     '�ܰ�/EA
    sss = sss & "         ttd_tldn,"    '������/EA
    sss = sss & "         ttd_chdn,"    '��ȯ��/EA
    sss = sss & "         ttd_result,"  '��� OK/NG
    sss = sss & "         ttd_ryn1,"    'Y/N
    sss = sss & "         ttd_ryn2,"    'Y/N
    sss = sss & "         ttd_ryn3,"    'Y/N
    sss = sss & "         ttd_ryn4,"    'Y/N
    sss = sss & "         ttd_ryn5,"    'Y/N
    sss = sss & "         ttd_rtyn,"
    sss = sss & "         ttd_rmk,"
    sss = sss & "         ttd_updte"    '
    
    sss = sss & ")values( "
    sss = sss & "        " & txt_dat1 & ","
    sss = sss & "        " & txt_seq1 & ","
    sss = sss & "        1,"
    sss = sss & "        1,"
    If opt_ryn(1).Value = True Then
        sss = sss & "        'Y',"
    Else
        sss = sss & "        'N',"
    End If
                   
    sss = sss & "        '" & txt_maker1 & "',"
    sss = sss & "        '" & txt_tipstd1 & "',"
    sss = sss & "        '" & txt_tipjil1 & "',"
    sss = sss & "        '" & txt_holder1 & "',"
    
    If Val(txt_rcntmn1) > 0 Then    '�д�ȸ���� �ּҰ�
        sss = sss & "        " & Val(txt_rcntmn1) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_rcntmx1) > 0 Then    '�д�ȸ���� �ִ밪
        sss = sss & "        " & Val(txt_rcntmx1) & ","
    Else
        sss = sss & "        null,"
    End If
                                                        
    If Val(txt_movmn1) > 0 Then    '�̼� �ּҰ�
        sss = sss & "        " & Val(txt_movmn1) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_movmx1) > 0 Then    '�̼� �ִ밪
        sss = sss & "        " & Val(txt_movmx1) & ","
    Else
        sss = sss & "        null,"
    End If
                                               
    If Val(txt_tct1) > 0 Then      'TCP
        sss = sss & "        " & Val(txt_tct1) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_pct1) > 0 Then      'PCT
        sss = sss & "        " & Val(txt_pct1) & ","
    Else
        sss = sss & "        null,"
    End If
    
    If Val(txt_depth1) > 0 Then    '�������
        sss = sss & "        " & Val(txt_depth1) & ","
    Else
        sss = sss & "        null,"
    End If

    sss = sss & "        " & Left(cmb_fluid1, 1) & ","
    sss = sss & "        " & Val(txt_qty1) & ","
    sss = sss & "        " & Val(txt_dan1) & ","
    sss = sss & "        " & Val(txt_tldn1) & ","
    sss = sss & "        " & Val(txt_chdn1) & ","
    
    If opt_ryn(1).Value = True Then
        
        sss = sss & "        '" & Replace(cmb_result1, ".", "") & "',"
                                                        
        If chk_ryn1.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn2.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn3.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn4.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn5.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        sss = sss & "       '" & cmb_rtyn1 & "',"
                                                                                                        
    Else
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
    End If
    
    sss = sss & "        '',"
    sss = sss & "        to_char(sysdate,'yyyymmdd')"
    sss = sss & " )"

    db.Execute sss, 64
    
    '---------------------->
    '��TEST1��
    '---------------------->
    sss = "insert into man_tooltestds("
    sss = sss & "         ttd_dat,"     '�������
    sss = sss & "         ttd_seq,"     '��ϼ���
    sss = sss & "         ttd_lno,"     '������
    sss = sss & "         ttd_gbn,"     '1:��������/2:�׽�Ʈ����
    sss = sss & "         ttd_ryn,"     'Y/N - �������
    
    sss = sss & "         ttd_maker,"   '������
    sss = sss & "         ttd_tipstd,"  '�� �԰�/�ڵ�
    sss = sss & "         ttd_tipjil,"  '�� ����
    sss = sss & "         ttd_holder,"  '����Ȧ��
    sss = sss & "         ttd_rcntmn,"  '�д�ȸ�� �ּ�
    sss = sss & "         ttd_rcntmx,"  '�д�ȸ�� �ִ�
    sss = sss & "         ttd_movmn,"   '�̼�(MM/REV) �ּ�
    sss = sss & "         ttd_movmx,"   '�̼�(MM/REV) �ִ�
    sss = sss & "         ttd_tct,"     'TCT
    sss = sss & "         ttd_pct,"     'PCT
                        
    sss = sss & "         ttd_depth,"   '�������
    
    sss = sss & "         ttd_fluid,"   '��������
    sss = sss & "         ttd_qty,"     '��������
    sss = sss & "         ttd_dan,"     '�ܰ�/EA
    sss = sss & "         ttd_tldn,"    '������/EA
    sss = sss & "         ttd_chdn,"    '��ȯ��/EA
    sss = sss & "         ttd_result,"  '��� OK/NG
    sss = sss & "         ttd_ryn1,"    'Y/N
    sss = sss & "         ttd_ryn2,"    'Y/N
    sss = sss & "         ttd_ryn3,"    'Y/N
    sss = sss & "         ttd_ryn4,"    'Y/N
    sss = sss & "         ttd_ryn5,"    'Y/N
    sss = sss & "         ttd_rtyn,     "
    sss = sss & "         ttd_rmk,"     '
    sss = sss & "         ttd_updte"    '
    
    sss = sss & ")values( "
    sss = sss & "        " & txt_dat1 & ","
    sss = sss & "        " & txt_seq1 & ","
    sss = sss & "        2,"
    sss = sss & "        2,"
    If opt_ryn(2).Value = True Then
        sss = sss & "        'Y',"
    Else
        sss = sss & "        'N',"
    End If
                   
    sss = sss & "        '" & txt_maker2 & "',"
    sss = sss & "        '" & txt_tipstd2 & "',"
    sss = sss & "        '" & txt_tipjil2 & "',"
    sss = sss & "        '" & txt_holder2 & "',"
    
    If Val(txt_rcntmn2) > 0 Then    '�д�ȸ���� �ּҰ�
        sss = sss & "        " & Val(txt_rcntmn2) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_rcntmx2) > 0 Then    '�д�ȸ���� �ִ밪
        sss = sss & "        " & Val(txt_rcntmx2) & ","
    Else
        sss = sss & "        null,"
    End If
                                                        
    If Val(txt_movmn2) > 0 Then    '�̼� �ּҰ�
        sss = sss & "        " & Val(txt_movmn2) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_movmx2) > 0 Then    '�̼� �ִ밪
        sss = sss & "        " & Val(txt_movmx2) & ","
    Else
        sss = sss & "        null,"
    End If
                                               
    If Val(txt_tct2) > 0 Then      'TCP
        sss = sss & "        " & Val(txt_tct2) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_pct2) > 0 Then      'PCT
        sss = sss & "        " & Val(txt_pct2) & ","
    Else
        sss = sss & "        null,"
    End If
    
    If Val(txt_depth2) > 0 Then    '�������
        sss = sss & "        " & Val(txt_depth2) & ","
    Else
        sss = sss & "        null,"
    End If

    sss = sss & "        " & Left(cmb_fluid2, 1) & ","
    sss = sss & "        " & Val(txt_qty2) & ","
    sss = sss & "        " & Val(txt_dan2) & ","
    sss = sss & "        " & Val(txt_tldn2) & ","
    sss = sss & "        " & Val(txt_chdn2) & ","
    
    If opt_ryn(2).Value = True Then
        
        sss = sss & "        '" & Replace(cmb_result1, ".", "") & "',"
        
        If chk_ryn1.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn2.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn3.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn4.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        If chk_ryn5.Value = 1 Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
        
        sss = sss & "       '" & cmb_rtyn1 & "',"
            
    Else
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
        sss = sss & "       null,"
    End If
    
    sss = sss & "        '',"
    sss = sss & "        to_char(sysdate,'yyyymmdd')"
    sss = sss & " )"
    
    db.Execute sss, 64
    
    
    '---------------------->
    '��TEST2��
    '---------------------->
    
    If chk_test2.Value = 1 Then
    
    
        sss = "insert into man_tooltestds("
        sss = sss & "         ttd_dat,"     '�������
        sss = sss & "         ttd_seq,"     '��ϼ���
        sss = sss & "         ttd_lno,"     '������
        sss = sss & "         ttd_gbn,"     '1:��������/2:�׽�Ʈ����
        sss = sss & "         ttd_ryn,"     'Y/N - �������
        
        sss = sss & "         ttd_maker,"   '������
        sss = sss & "         ttd_tipstd,"  '�� �԰�/�ڵ�
        sss = sss & "         ttd_tipjil,"  '�� ����
        sss = sss & "         ttd_holder,"  '����Ȧ��
        sss = sss & "         ttd_rcntmn,"  '�д�ȸ�� �ּ�
        sss = sss & "         ttd_rcntmx,"  '�д�ȸ�� �ִ�
        sss = sss & "         ttd_movmn,"   '�̼�(MM/REV) �ּ�
        sss = sss & "         ttd_movmx,"   '�̼�(MM/REV) �ִ�
        sss = sss & "         ttd_tct,"     'TCT
        sss = sss & "         ttd_pct,"     'PCT
                            
        sss = sss & "         ttd_depth,"   '�������
        
        sss = sss & "         ttd_fluid,"   '��������
        sss = sss & "         ttd_qty,"     '��������
        sss = sss & "         ttd_dan,"     '�ܰ�/EA
        sss = sss & "         ttd_tldn,"    '������/EA
        sss = sss & "         ttd_chdn,"    '��ȯ��/EA
        sss = sss & "         ttd_result,"  '��� OK/NG
        sss = sss & "         ttd_ryn1,"    'Y/N
        sss = sss & "         ttd_ryn2,"    'Y/N
        sss = sss & "         ttd_ryn3,"    'Y/N
        sss = sss & "         ttd_ryn4,"    'Y/N
        sss = sss & "         ttd_ryn5,"    'Y/N
        sss = sss & "         ttd_rtyn,"    '
        sss = sss & "         ttd_rmk,"     '
        sss = sss & "         ttd_updte"    '
        
        sss = sss & ")values( "
        sss = sss & "        " & txt_dat1 & ","
        sss = sss & "        " & txt_seq1 & ","
        sss = sss & "        3,"
        sss = sss & "        2,"
        If opt_ryn(3).Value = True Then
            sss = sss & "        'Y',"
        Else
            sss = sss & "        'N',"
        End If
                       
        sss = sss & "        '" & txt_maker3 & "',"
        sss = sss & "        '" & txt_tipstd3 & "',"
        sss = sss & "        '" & txt_tipjil3 & "',"
        sss = sss & "        '" & txt_holder3 & "',"
        
        If Val(txt_rcntmn3) > 0 Then    '�д�ȸ���� �ּҰ�
            sss = sss & "        " & Val(txt_rcntmn3) & ","
        Else
            sss = sss & "        null,"
        End If
        If Val(txt_rcntmx3) > 0 Then    '�д�ȸ���� �ִ밪
            sss = sss & "        " & Val(txt_rcntmx3) & ","
        Else
            sss = sss & "        null,"
        End If
                                                            
        If Val(txt_movmn3) > 0 Then    '�̼� �ּҰ�
            sss = sss & "        " & Val(txt_movmn3) & ","
        Else
            sss = sss & "        null,"
        End If
        If Val(txt_movmx3) > 0 Then    '�̼� �ִ밪
            sss = sss & "        " & Val(txt_movmx3) & ","
        Else
            sss = sss & "        null,"
        End If
                                                   
        If Val(txt_tct3) > 0 Then      'TCP
            sss = sss & "        " & Val(txt_tct3) & ","
        Else
            sss = sss & "        null,"
        End If
        If Val(txt_pct3) > 0 Then      'PCT
            sss = sss & "        " & Val(txt_pct3) & ","
        Else
            sss = sss & "        null,"
        End If
        
        If Val(txt_depth3) > 0 Then    '�������
            sss = sss & "        " & Val(txt_depth3) & ","
        Else
            sss = sss & "        null,"
        End If
    
        sss = sss & "        " & Left(cmb_fluid3, 1) & ","
        sss = sss & "        " & Val(txt_qty3) & ","
        sss = sss & "        " & Val(txt_dan3) & ","
        sss = sss & "        " & Val(txt_tldn3) & ","
        sss = sss & "        " & Val(txt_chdn3) & ","
        
        If opt_ryn(3).Value = True Then
            
            sss = sss & "        '" & Replace(cmb_result1, ".", "") & "',"
            
            If chk_ryn1.Value = 1 Then
                sss = sss & "        'Y',"
            Else
                sss = sss & "        'N',"
            End If
            If chk_ryn2.Value = 1 Then
                sss = sss & "        'Y',"
            Else
                sss = sss & "        'N',"
            End If
            If chk_ryn3.Value = 1 Then
                sss = sss & "        'Y',"
            Else
                sss = sss & "        'N',"
            End If
            If chk_ryn4.Value = 1 Then
                sss = sss & "        'Y',"
            Else
                sss = sss & "        'N',"
            End If
            If chk_ryn5.Value = 1 Then
                sss = sss & "        'Y',"
            Else
                sss = sss & "        'N',"
            End If
            sss = sss & "       '" & cmb_rtyn1 & "',"
        Else
            sss = sss & "       null,"
            sss = sss & "       null,"
            sss = sss & "       null,"
            sss = sss & "       null,"
            sss = sss & "       null,"
            sss = sss & "       null,"
            sss = sss & "       null,"
        End If
            
        sss = sss & "        '',"
        sss = sss & "        to_char(sysdate,'yyyymmdd')"
        sss = sss & " )"

        db.Execute sss, 64
            
    End If
    
    Ws.CommitTrans
    
    
    Call msg_display("��ϳ����� �����Ǿ����ϴ�.")
    
    Exit Sub
    
err_rtn:

    Ws.Rollback
    MsgBox (Err.Description)
    
End Sub

'-------------------------------
'��������
'-------------------------------
Private Sub btn_del1_Click()
    
On Error GoTo err_rtn
    
     If Job_Level < 9 Then
        MsgBox ("�۾������� �����ϴ�!")
        Exit Sub
    End If
    
    If IsNumeric(txt_dat1) = False Or Len(txt_dat1) <> 8 Then
        Call msg_display("��� ���ڸ� Ȯ���ϼ���!")
        txt_dat1.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txt_seq1) = False Or Len(txt_seq1) < 1 Then
        Call msg_display("��� ������ Ȯ���ϼ���!")
        txt_seq1.SetFocus
        Exit Sub
    End If
    
    sss = "select * from man_tooltesthd, man_tooltestds"
    sss = sss & "  where tth_dat = " & txt_dat1
    sss = sss & "    and tth_seq = " & txt_seq1
    sss = sss & "    and tth_dat = ttd_dat"
    sss = sss & "    and tth_seq = ttd_seq"
    '
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Rs.RecordCount < 1 Then
        Rs.Close
        Call msg_display("������ �����ϴ�. �����Ұ�!")
        Exit Sub
    End If
    '
    '����
    If MsgBox("(" & txt_dat1 & "-" & Format(txt_seq1, "000") & ") ��ϵ� ������ �����Ͻðڽ��ϱ�?", vbYesNo) <> vbYes Then
       Rs.Close
       msg_display ("��ҵǾ����ϴ�.")
       Exit Sub
    End If
    '
    '
    Ws.BeginTrans
    '
    '
    sss = "       delete from man_tooltesthd"
    sss = sss & "  where tth_dat = " & txt_dat1
    sss = sss & "    and tth_seq = " & txt_seq1
    '
    db.Execute sss, 64
    '
    '
    sss = "       delete from man_tooltestds"
    sss = sss & "  where ttd_dat = " & txt_dat1
    sss = sss & "    and ttd_seq = " & txt_seq1
    '
    db.Execute sss, 64
    '
    '
    sss = "       delete from oth_applist"
    sss = sss & "  where apl_table = 'man_tooltesthd'"
    sss = sss & "    and apl_tdat = " & txt_dat1
    sss = sss & "    and apl_tseq = " & txt_seq1
    '
    db.Execute sss, 64
    '
    '
    '÷������ ��ü ����
    If Not IsNull(Rs!tth_file1) Or Not IsNull(Rs!tth_file2) Or Not IsNull(Rs!tth_file3) Or Not IsNull(Rs!tth_file4) Then
       '
       '===========================
       ' FTP�� �̿��� ���� ����
       '===========================
       If FTP_Connection Then
          '
          If Not FTP���üũ(SERVER_PATH) Then
             Call FTP_DisConnect
             Ws.Rollback
             MsgBox "������θ� ã���� �����ϴ�.(������������ ����)"
             Exit Sub
          End If
          '
          If Not IsNull(Rs!tth_file1) Then
             If Not FTP_Delete(SERVER_PATH & Rs!tth_file1) Then
                Ws.Rollback
                Call FTP_DisConnect
                mkpoen05MDI.msg = "������ ������ �� �����ϴ�. (������������ ����)"
                MsgBox "������ ������ �� �����ϴ�. (������������ ����)"
                Exit Sub
             End If
          End If
          '
          If Not IsNull(Rs!tth_file2) Then
             If Not FTP_Delete(SERVER_PATH & Rs!tth_file2) Then
                Ws.Rollback
                Call FTP_DisConnect
                mkpoen05MDI.msg = "������ ������ �� �����ϴ�. (������������ ����)"
                MsgBox "������ ������ �� �����ϴ�. (������������ ����)"
                Exit Sub
             End If
          End If
          '
          If Not IsNull(Rs!tth_file3) Then
             If Not FTP_Delete(SERVER_PATH & Rs!tth_file3) Then
                Ws.Rollback
                Call FTP_DisConnect
                mkpoen05MDI.msg = "������ ������ �� �����ϴ�. (������������ ����)"
                MsgBox "������ ������ �� �����ϴ�. (������������ ����)"
                Exit Sub
             End If
          End If
          '
          If Not IsNull(Rs!tth_file4) Then
             If Not FTP_Delete(SERVER_PATH & Rs!tth_file4) Then
                Ws.Rollback
                Call FTP_DisConnect
                mkpoen05MDI.msg = "������ ������ �� �����ϴ�. (������������ ����)"
                MsgBox "������ ������ �� �����ϴ�. (������������ ����)"
                Exit Sub
             End If
          End If
          '
          Call FTP_DisConnect
          '
       End If
       '
       '======================================
       ' ��Ʈ��ũ ����̹��� �̿��� ���� ����
       '======================================
       ' Kill (N_Driver & ":\" & txt_dat1 & Format(txt_seq1, "000") & "-00" & "*")
        '
    End If
    '
    '
    Ws.CommitTrans
    '
    '
    Rs.Close

    '�ʱ�ȭ
    Call clear_rtn
    
    btn_view1.Enabled = True
    
    'txt_dat1 = dat: txt_dat1.Enabled = True
    'txt_seq1 = seq: txt_seq1.Enabled = True
    
    txt_dat1.Enabled = True
    txt_seq1.Enabled = True
    '
    btn_add1.Enabled = True
    btn_mod1.Enabled = False
    btn_del1.Enabled = False
    btn_prt1.Enabled = False
    '
    Call msg_display("������ �����Ǿ����ϴ�!")
    '
    txt_testno1.SetFocus
    '
    Exit Sub
    '
err_rtn:
    Ws.Rollback
    msg_display (Err.Description)
End Sub

'----------------------
'�ʱ�ȭ
'----------------------
Private Sub btn_clear1_Click()
    
    Call clear_rtn
    
    txt_dat1 = "": txt_dat1.Enabled = True
    txt_seq1 = "": txt_seq1.Enabled = True
    
    btn_view1.Enabled = True
    
    btn_add1.Enabled = True
    btn_mod1.Enabled = False
    btn_del1.Enabled = False
    
    txt_dat1.SetFocus
    
    txt_dat.Enabled = False
    txt_seq.Enabled = False
    
    opt_n_ja.Value = True
    opt_n_su.Value = False
    
    opt_n_ja.Enabled = True
    opt_n_su.Enabled = True
    
End Sub

Private Sub clear_rtn()

    txt_testno1 = ""
    txt_tdat1 = ""
    txt_jdat1 = ""
    txt_jseq1 = ""
    txt_title1 = ""
    txt_mcd1 = ""
    txt_mnm1 = ""
    txt_mmake1 = ""
    
    txt_msok1 = ""
    
    txt_lotno1 = ""
    txt_bpcd1 = ""
    txt_bpjil1 = ""
    txt_jacd1 = ""
    txt_jajil1 = ""
    
    chk_pyn1.Value = False
    chk_pyn2.Value = False
    chk_pyn3.Value = False
    chk_pyn4.Value = False
    chk_pyn5.Value = False
    
    txt_tsab1 = ""
    txt_tnm1 = ""
    txt_tsoknm1 = ""
    
    '������������
    opt_ryn(1).Value = False
    txt_maker1 = ""
    txt_tipstd1 = ""
    txt_tipjil1 = ""
    txt_holder1 = ""
    txt_rcntmn1 = ""
    txt_rcntmx1 = ""
    txt_depth1 = ""
    txt_movmn1 = ""
    txt_movmx1 = ""
    txt_tct1 = ""
    txt_pct1 = ""
    cmb_fluid1.ListIndex = 0
    txt_qty1 = ""
    txt_dan1 = ""
    txt_tldn1 = ""
    txt_chdn1 = ""
    
    '�׽�Ʈ1 ����
    opt_ryn(2).Value = False
    txt_maker2 = ""
    txt_tipstd2 = ""
    txt_tipjil2 = ""
    txt_holder2 = ""
    txt_rcntmn2 = ""
    txt_rcntmx2 = ""
    txt_depth2 = ""
    txt_movmn2 = ""
    txt_movmx2 = ""
    txt_tct2 = ""
    txt_pct2 = ""
    cmb_fluid2.ListIndex = 0
    txt_qty2 = ""
    txt_dan2 = ""
    txt_tldn2 = ""
    txt_chdn2 = ""
                
    '�׽�Ʈ2 ����
    opt_ryn(3).Value = False
    chk_test2.Value = False
    txt_maker3 = ""
    txt_tipstd3 = ""
    txt_tipjil3 = ""
    txt_holder3 = ""
    txt_rcntmn3 = ""
    txt_rcntmx3 = ""
    txt_depth3 = ""
    txt_movmn3 = ""
    txt_movmx3 = ""
    txt_tct3 = ""
    txt_pct3 = ""
    cmb_fluid3.ListIndex = 0
    txt_qty3 = ""
    txt_dan3 = ""
    txt_tldn3 = ""
    txt_chdn3 = ""
    
    opt_ryn(3).Enabled = False
    txt_maker3.Enabled = False
    txt_tipstd3.Enabled = False
    txt_tipjil3.Enabled = False
    txt_holder3.Enabled = False
    txt_rcntmn3.Enabled = False
    txt_rcntmx3.Enabled = False
    txt_depth3.Enabled = False
    txt_movmn3.Enabled = False
    txt_movmx3.Enabled = False
    txt_tct3.Enabled = False
    txt_pct3.Enabled = False
    cmb_fluid3.Enabled = False
    txt_qty3.Enabled = False
    txt_dan3.Enabled = False
    txt_tldn3.Enabled = False
    txt_chdn3.Enabled = False
    
    txt_maker3.BackColor = &H8000000F
    txt_tipstd3.BackColor = &H8000000F
    txt_tipjil3.BackColor = &H8000000F
    txt_holder3.BackColor = &H8000000F
    txt_rcntmn3.BackColor = &H8000000F
    txt_rcntmx3.BackColor = &H8000000F
    txt_depth3.BackColor = &H8000000F
    txt_movmn3.BackColor = &H8000000F
    txt_movmx3.BackColor = &H8000000F
    txt_tct3.BackColor = &H8000000F
    txt_pct3.BackColor = &H8000000F
    cmb_fluid3.BackColor = &H8000000F
    txt_qty3.BackColor = &H8000000F
    txt_dan3.BackColor = &H8000000F
    txt_tldn3.BackColor = &H8000000F
    txt_chdn3.BackColor = &H8000000F
    
    '���
    chk_ryn1.Value = False
    chk_ryn2.Value = False
    chk_ryn3.Value = False
    chk_ryn4.Value = False
    chk_ryn5.Value = False
    
    cmb_result1.ListIndex = 0
    cmb_rtyn1.ListIndex = 0
    
    '���/��
    sprd_rmk1.row = 1
    sprd_rmk1.Col = 1: sprd_rmk1.Text = "LOT.NO : " & vbCrLf & "�������� : " & vbCrLf & "�׽�Ʈ���� : "
    sprd_rmk1.Col = 2: sprd_rmk1.Text = ""
    
    For ii = 1 To 5
        spd_file11.row = ii: spd_file11.Col = 2: spd_file11.Text = ""
    Next ii
    
    For ii = 1 To 5
        spd_file12.row = ii: spd_file12.Col = 1: spd_file12.Text = ""
    Next ii
    
    spd_file11.Visible = True               '����(�űԿ�)
    spd_file12.Visible = False              '����(������)
    lbl_cmt1.Visible = False
    
    btn_prt1.Enabled = False
    
End Sub

'�׽�Ʈ2Ȱ��ȭ
Private Sub chk_test2_Click()
    
    If chk_test2.Value = Checked Then
        opt_ryn(3).Enabled = True
        txt_maker3.Enabled = True
        txt_tipstd3.Enabled = True
        txt_tipjil3.Enabled = True
        txt_holder3.Enabled = True
        txt_rcntmn3.Enabled = True
        txt_rcntmx3.Enabled = True
        txt_depth3.Enabled = True
        txt_movmn3.Enabled = True
        txt_movmx3.Enabled = True
        txt_tct3.Enabled = True
        txt_pct3.Enabled = True
        cmb_fluid3.Enabled = True
        txt_qty3.Enabled = True
        txt_dan3.Enabled = True
        txt_tldn3.Enabled = True
        txt_chdn3.Enabled = True
        
        txt_maker3.BackColor = &HFFFFFF
        txt_tipstd3.BackColor = &HFFFFFF
        txt_tipjil3.BackColor = &HFFFFFF
        txt_holder3.BackColor = &HFFFFFF
        txt_rcntmn3.BackColor = &HFFFFFF
        txt_rcntmx3.BackColor = &HFFFFFF
        txt_depth3.BackColor = &HFFFFFF
        txt_movmn3.BackColor = &HFFFFFF
        txt_movmx3.BackColor = &HFFFFFF
        txt_tct3.BackColor = &HFFFFFF
        txt_pct3.BackColor = &HFFFFFF
        cmb_fluid3.BackColor = &HFFFFFF
        txt_qty3.BackColor = &HFFFFFF
        txt_dan3.BackColor = &HFFFFFF
        txt_tldn3.BackColor = &HFFFFFF
        txt_chdn3.BackColor = &HFFFFFF
    Else
        opt_ryn(3).Enabled = False
        opt_ryn(3).Value = False
        chk_test2.Value = False
        txt_maker3 = ""
        txt_tipstd3 = ""
        txt_tipjil3 = ""
        txt_holder3 = ""
        txt_rcntmn3 = ""
        txt_rcntmx3 = ""
        txt_depth3 = ""
        txt_movmn3 = ""
        txt_movmx3 = ""
        txt_tct3 = ""
        txt_pct3 = ""
        cmb_fluid3.ListIndex = 0
        txt_qty3 = ""
        txt_dan3 = ""
        txt_tldn3 = ""
        txt_chdn3 = ""
        
        txt_maker3.Enabled = False
        txt_tipstd3.Enabled = False
        txt_tipjil3.Enabled = False
        txt_holder3.Enabled = False
        txt_rcntmn3.Enabled = False
        txt_rcntmx3.Enabled = False
        txt_depth3.Enabled = False
        txt_movmn3.Enabled = False
        txt_movmx3.Enabled = False
        txt_tct3.Enabled = False
        txt_pct3.Enabled = False
        cmb_fluid3.Enabled = False
        txt_qty3.Enabled = False
        txt_dan3.Enabled = False
        txt_tldn3.Enabled = False
        txt_chdn3.Enabled = False
        
        txt_maker3.BackColor = &H8000000F
        txt_tipstd3.BackColor = &H8000000F
        txt_tipjil3.BackColor = &H8000000F
        txt_holder3.BackColor = &H8000000F
        txt_rcntmn3.BackColor = &H8000000F
        txt_rcntmx3.BackColor = &H8000000F
        txt_depth3.BackColor = &H8000000F
        txt_movmn3.BackColor = &H8000000F
        txt_movmx3.BackColor = &H8000000F
        txt_tct3.BackColor = &H8000000F
        txt_pct3.BackColor = &H8000000F
        cmb_fluid3.BackColor = &H8000000F
        txt_qty3.BackColor = &H8000000F
        txt_dan3.BackColor = &H8000000F
        txt_tldn3.BackColor = &H8000000F
        txt_chdn3.BackColor = &H8000000F
    
    End If

End Sub

Private Sub opt_n_ja_Click()
    
    txt_dat.Enabled = False
    txt_dat.Text = ""
    
    txt_seq.Enabled = False
    txt_seq.Text = ""
    
    txt_dat.BackColor = &H8000000F
    txt_seq.BackColor = &H8000000F
    
End Sub

Private Sub opt_n_su_Click()
    
    txt_dat.Enabled = True
    txt_dat.Text = ""
    
    txt_seq.Enabled = True
    txt_seq.Text = ""
    
    txt_dat.BackColor = &HFFFFFF
    txt_seq.BackColor = &HFFFFFF
    
    txt_dat.SetFocus
    
End Sub

Private Sub opt_ryn_Click(Index As Integer)
    
    If Index = 1 Then
        opt_ryn(2).Value = False
        opt_ryn(3).Value = False
    End If
    
    If Index = 2 Then
        opt_ryn(1).Value = False
        opt_ryn(3).Value = False
    End If
        
    If Index = 3 Then
        opt_ryn(1).Value = False
        opt_ryn(2).Value = False
    End If
                                                                                            
End Sub


Private Sub sprd2_DblClick(ByVal Col As Long, ByVal row As Long)
    
    sprd2.row = row
    
    If Col = 1 And row <> 0 Then
        sprd2.Col = 1
        
        If Len(sprd2.Text) <> 12 Then Exit Sub
        
        'Call btn_clear1_Click
        txt_dat1 = Left(sprd2.Text, 8)
        txt_seq1 = Right(sprd2.Text, 3)
        SSTab1.Tab = 0
        Call btn_view1_Click
        
        SSTab1.Tab = 0
        
    End If
    
End Sub




'------------------------->
'����ڵ� ��ȸ
'------------------------->
Private Sub txt_mcd1_LostFocus()
'��� ��ȸ

    txt_mcd1 = UCase(Trim(txt_mcd1))
    
    If Len(Trim(txt_mcd1)) > 1 Then
        sss = "       select mhc_name, ems_mark, sok_name(mhc_sok) soknm "
        sss = sss & "   from man_machcd, eam_mast"
        sss = sss & "  where mhc_code = '" & txt_mcd1 & "'"
        sss = sss & "    and mhc_code = ems_mcd(+)"
            
        Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        If Ks.RecordCount < 1 Then
            txt_mnm1 = "": txt_mmake1 = "": txt_msok1 = ""
            Ks.Close: Call msg_display("����ڵ带 Ȯ���ϼ���!")
            txt_mcd1.SetFocus
            Exit Sub
        End If
        
        If Not IsNull(Ks!mhc_name) Then txt_mnm1 = Ks!mhc_name
        If Not IsNull(Ks!ems_mark) Then txt_mmake1 = Ks!ems_mark
        If Not IsNull(Ks!soknm) Then txt_msok1 = Ks!soknm
        Ks.Close
        
    Else
        txt_mnm1 = ""
        txt_mmake1 = ""
        txt_msok1 = ""
    End If


End Sub

'------------------------->
'LOT NO. ��ȸ
'------------------------->
Private Sub txt_lotno1_LostFocus()
'��� ��ȸ

    txt_lotno1 = UCase(Trim(txt_lotno1))
    
    If Len(Trim(txt_lotno1)) = 8 Then
        sss = "       select dit_bpcd, dit_bpjil, dit_jacd, dit_jajil"
        sss = sss & "   from man_direct"
        sss = sss & "  where dit_lot = '" & txt_lotno1 & "'"
            
        Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        If Ks.RecordCount < 1 Then
            txt_bpcd1 = "": txt_bpjil1 = ""
            txt_jacd1 = "": txt_jajil1 = ""
            Ks.Close: Call msg_display("LOT NO.�� Ȯ���ϼ���!")
            txt_lotno1.SetFocus
            Exit Sub
        End If
        
        If Not IsNull(Ks!dit_bpcd) Then txt_bpcd1 = Ks!dit_bpcd
        If Not IsNull(Ks!dit_bpjil) Then txt_bpjil1 = Ks!dit_bpjil
        If Not IsNull(Ks!dit_jacd) Then txt_jacd1 = Ks!dit_jacd
        If Not IsNull(Ks!dit_jajil) Then txt_jajil1 = Ks!dit_jajil
        Ks.Close
        
    Else
        txt_bpcd1 = ""
        txt_bpjil1 = ""
        txt_jacd1 = ""
        txt_jajil1 = ""
    End If


End Sub

'------------------------->
'��� ��ȸ
'------------------------->
Private Sub txt_tsab1_LostFocus()
    
    If Len(Trim(txt_tsab1)) > 0 Then
    
        txt_tsab1 = Val(txt_tsab1)
        
        sss = "       select sin_name, sin_sok, sok_name(sin_sok) soknm"
        sss = sss & "   from peo_sinbun"
        sss = sss & "  where sin_sab = " & txt_tsab1
        'sss = sss & "    and sin_taedt = 0 "
        sss = sss & "    and sin_sab < 5000"
        sss = sss & "  order by sin_sab desc "
            
        Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        If Ks.RecordCount < 1 Then
            txt_tsab1 = "": txt_tnm1 = "": txt_tsoknm1 = ""
            Ks.Close: Call msg_display("�۾��� ����� Ȯ���ϼ���!")
            txt_tsab1.SetFocus
            Exit Sub
        End If
        
        If Not IsNull(Ks!sin_name) Then txt_tnm1 = Ks!sin_name
        'If Not IsNull(ks!sin_sok) Then txt_rsok1 = ks!sin_sok
        If Not IsNull(Ks!soknm) Then txt_tsoknm1 = Ks!soknm
        
        Ks.Close
            
    Else
        txt_tsab1 = ""
        txt_tnm1 = ""
        txt_tsoknm1 = ""
    End If
    
End Sub

Private Sub txt_tnm1_LostFocus()
    
    If Len(Trim(txt_tnm1)) > 0 Then
    
        txt_tnm1 = Trim(txt_tnm1)
        
        sss = "       select sin_sab, sin_name, sin_sok, sok_name(sin_sok) soknm"
        sss = sss & "   from peo_sinbun"
        sss = sss & "  where sin_name = '" & txt_tnm1 & "'"
        'sss = sss & "    and sin_taedt = 0 "
        sss = sss & "    and sin_sab < 5000"
        sss = sss & "  order by sin_sab desc "
            
        Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        If Ks.RecordCount < 1 Then
            txt_tnm1 = "": txt_tnm1 = "": txt_tsoknm1 = ""
            Ks.Close: Call msg_display("�۾��� �̸��� Ȯ���ϼ���!")
            txt_tnm1.SetFocus
            Exit Sub
        End If
        
        If Not IsNull(Ks!sin_name) Then txt_tnm1 = Ks!sin_name
        If Not IsNull(Ks!sin_sab) Then txt_tsab1 = Ks!sin_sab
        If Not IsNull(Ks!soknm) Then txt_tsoknm1 = Ks!soknm
        
        Ks.Close
            
    Else
        txt_tsab1 = ""
        txt_tnm1 = ""
        txt_tsoknm1 = ""
    End If
    
End Sub

Private Sub btn_prt1_Click()
    
'    If Job_Level < 8 Then
'        MsgBox ("�۾������� �����ϴ�!")
'        Exit Sub
'    End If
    
    mkpoen05_print.txt_dat1 = txt_dat1
    mkpoen05_print.txt_seq1 = txt_seq1
    mkpoen05_print.btn_view_Click
    mkpoen05_print.Visible = True
    
End Sub

Private Sub cmb_result1_Click()
    '
    If cmb_result1 = "O.K" Then
       '
       chk_ryn1.Value = False
       chk_ryn2.Value = False
       chk_ryn3.Value = False
       chk_ryn4.Value = False
       chk_ryn5.Value = False
       '
       chk_ryn1.Caption = "1.�������� ����"
       chk_ryn2.Caption = "2.Ĩó�� ��ȣ"
       chk_ryn3.Caption = "3.������ ����"
       chk_ryn4.Caption = "4.�ð� ����"
       chk_ryn5.Caption = "5.��Ÿ"
       '
    End If
    '
    If cmb_result1 = "N.G" Then
       '
       chk_ryn1.Value = False
       chk_ryn2.Value = False
       chk_ryn3.Value = False
       chk_ryn4.Value = False
       chk_ryn5.Value = False
       '
       chk_ryn1.Caption = "1.���"
       chk_ryn2.Caption = "2.����"
       chk_ryn3.Caption = "3.Ĩó�� �ҷ�"
       chk_ryn4.Caption = "4.������ ���"
       chk_ryn5.Caption = "5.��Ÿ"
    End If
    '
End Sub

Private Sub txt_testno1_LostFocus()
    txt_testno1.Text = UCase(Trim(txt_testno1))
End Sub
Private Sub txt_maker1_LostFocus()
    txt_maker1 = UCase(Trim(txt_maker1))
End Sub
Private Sub txt_tipstd1_LostFocus()
    txt_tipstd1 = UCase(Trim(txt_tipstd1))
End Sub
Private Sub txt_tipjil1_LostFocus()
    txt_tipjil1 = UCase(Trim(txt_tipjil1))
End Sub
Private Sub txt_holder1_LostFocus()
    txt_holder1 = UCase(Trim(txt_holder1))
End Sub
Private Sub txt_maker2_LostFocus()
    txt_maker2 = UCase(Trim(txt_maker2))
End Sub
Private Sub txt_tipstd2_LostFocus()
    txt_tipstd2 = UCase(Trim(txt_tipstd2))
End Sub
Private Sub txt_tipjil2_LostFocus()
    txt_tipjil2 = UCase(Trim(txt_tipjil2))
End Sub
Private Sub txt_holder2_LostFocus()
    txt_holder2 = UCase(Trim(txt_holder2))
End Sub
Private Sub txt_maker3_LostFocus()
    txt_maker3 = UCase(Trim(txt_maker3))
End Sub
Private Sub txt_tipstd3_LostFocus()
    txt_tipstd3 = UCase(Trim(txt_tipstd3))
End Sub
Private Sub txt_tipjil3_LostFocus()
    txt_tipjil3 = UCase(Trim(txt_tipjil3))
End Sub
Private Sub txt_holder3_LostFocus()
    txt_holder3 = UCase(Trim(txt_holder3))
End Sub


'÷������ ��ȸ
Public Sub file_view()
    
    For ii = 1 To 5
        spd_file12.row = ii: spd_file12.Col = 1: spd_file12.Text = ""
    Next ii
    
    sss = "       select tth_file1,tth_file2,tth_file3, tth_file4"
    sss = sss & "   from man_tooltesthd "
    sss = sss & "  where tth_dat = " & txt_dat1
    sss = sss & "    and tth_seq = " & txt_seq1
    
    Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    
    If Ks.RecordCount < 1 Then
       Ks.Close
       Exit Sub
    End If
    
    '÷������1
    spd_file12.row = 1
    If Not IsNull(Ks!tth_file1) Then
        spd_file12.Col = 1: spd_file12.Text = Ks!tth_file1
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "����"
    Else
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "���"
    End If
    
    '÷������2
    spd_file12.row = 2
    If Not IsNull(Ks!tth_file2) Then
        spd_file12.Col = 1: spd_file12.Text = Ks!tth_file2
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "����"
    Else
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "���"
    End If
    
    '÷������2
    spd_file12.row = 3
    If Not IsNull(Ks!tth_file3) Then
        spd_file12.Col = 1: spd_file12.Text = Ks!tth_file3
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "����"
    Else
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "���"
    End If
    
    '÷������2
    spd_file12.row = 4
    If Not IsNull(Ks!tth_file4) Then
        spd_file12.Col = 1: spd_file12.Text = Ks!tth_file4
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "����"
    Else
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "���"
    End If
    
    Ks.Close
    
End Sub

'----------------------
'÷������ �߰� �� ���� - �űԵ��
'----------------------
Private Sub spd_file11_ButtonClicked(ByVal Col As Long, ByVal row As Long, ByVal ButtonDown As Integer)
    '
    On Error GoTo err_rtn
    '
    If Job_Level < 9 Then
        MsgBox ("�۾������� �����ϴ�!")
        Exit Sub
    End If
    '
    If Col = 1 Then
        
        spd_file11.row = row: spd_file11.Col = 2
                    
        spd_file11.Text = ""
        
        Comm1.CancelError = True
        Comm1.Flags = cdlOFNOverwritePrompt
        Comm1.DialogTitle = "�������� ��û ÷������"
        Comm1.InitDir = "C:\"
        Comm1.Filter = "PDF���� or JPG���� (*.pdf;*.jpg)|*.pdf;*.jpg"
        Comm1.ShowOpen
        
        If UCase(Right(Comm1.fileName, 3)) <> "PDF" And UCase(Right(Comm1.fileName, 3)) <> "JPG" Then
            MsgBox ("������ ���� Ȯ���� 'PDF' or 'JPG' �� �ƴմϴ�.")
            Exit Sub
        End If
        '
        spd_file11.Text = UCase(Comm1.fileName)
        '
        Call msg_display("÷�������� �߰��Ǿ����ϴ�!")
        Exit Sub
    
    End If
    
    If Col = 4 Then
        
        spd_file11.row = row: spd_file11.Col = 2
        If Len(spd_file11.Text) < 1 Then Exit Sub
        spd_file11.Text = ""
        Call msg_display("÷�������� �����Ǿ����ϴ�!")
        Exit Sub

    End If
    '
    '
    Exit Sub

err_rtn:
  mkpoen05MDI.msg = Err.Description

End Sub

'----------------------
'÷������ �߰� �� ���� - �������
'----------------------
Private Sub spd_file12_ButtonClicked(ByVal Col As Long, ByVal row As Long, ByVal ButtonDown As Integer)
    '
    On Error GoTo err_rtn
    '
    If Job_Level < 9 Then
        MsgBox ("�۾������� �����ϴ�!")
        Exit Sub
    End If

    Dim isrno As String     '�������� ��ȣ
    Dim filenm1 As String   '����1
    Dim filenm2 As String   '����2
    Dim filenm3 As String   '����3
    Dim filenm4 As String   '����4
    
    Dim fs            As Object 'Scripting.FileSystemObject ��ü
    Dim lsSouurce     As String '������ ���
    Dim lsDestination As String '������ġ ���
    Dim F_NAME     As String
    
    spd_file12.row = row: spd_file12.Col = Col
    
    'If Len(Gsab) < 1 Then
    '    txt_pwd.SetFocus
    '    Call msg_display("�α��� �� �۾��ϼ���!!")
    '    Exit Sub
    'End If
    
    'If Gsok <> "D100" And Gsok <> "D000" And Gsab <> "1476" Then
    '    Call msg_display("���źμ��� ���/����/���� �۾� �����մϴ�. �Ҽ��� Ȯ���ϼ���!")
    '    Exit Sub
    'End If
    
    '÷������ ���
    If spd_file12.TypeButtonText = "���" And Col = 3 Then
            
        sss = "       select tth_file1, tth_file2, tth_file3, tth_file4"
        sss = sss & "   from man_tooltesthd "
        sss = sss & "  where tth_dat = " & txt_dat1
        sss = sss & "    and tth_seq = " & txt_seq1
    
        Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    
        If Rs.RecordCount < 1 Then
            Rs.Close
            msg_display ("÷�ε���� ��ϳ����� �����ϴ�. ��Ϲ�ȣ�� Ȯ���ϼ���!")
            Exit Sub
        End If
            
        'chk
        Rs.Close
        '
        Comm1.CancelError = True
        Comm1.Flags = cdlOFNOverwritePrompt
        Comm1.DialogTitle = "÷������"
        Comm1.InitDir = "C:\"
        Comm1.Filter = "PDF���� or JPG���� (*.pdf;*.jpg)|*.pdf;*.jpg"
            Comm1.ShowOpen
        If UCase(Right(Comm1.fileName, 3)) <> "PDF" And UCase(Right(Comm1.fileName, 3)) <> "JPG" Then
            MsgBox ("������ ���� Ȯ���� 'PDF' �Ǵ� 'JPG' ������ �ƴմϴ�!")
            Exit Sub
        End If
        '
        If UCase(Right(Comm1.fileName, 3)) = "PDF" Then
            If row = 1 Then filenm1 = txt_dat1 & Format(txt_seq1, "000") & "-001" & ".PDF"
            If row = 2 Then filenm2 = txt_dat1 & Format(txt_seq1, "000") & "-002" & ".PDF"
            If row = 3 Then filenm3 = txt_dat1 & Format(txt_seq1, "000") & "-003" & ".PDF"
            If row = 4 Then filenm4 = txt_dat1 & Format(txt_seq1, "000") & "-004" & ".PDF"
        ElseIf UCase(Right(Comm1.fileName, 3)) = "JPG" Then
            If row = 1 Then filenm1 = txt_dat1 & Format(txt_seq1, "000") & "-001" & ".JPG"
            If row = 2 Then filenm2 = txt_dat1 & Format(txt_seq1, "000") & "-002" & ".JPG"
            If row = 3 Then filenm3 = txt_dat1 & Format(txt_seq1, "000") & "-003" & ".JPG"
            If row = 4 Then filenm4 = txt_dat1 & Format(txt_seq1, "000") & "-004" & ".JPG"
        End If
        '
       '===========================
       ' FTP�� �̿��� ���� ���ε�
       '===========================
       If FTP_Connection Then
          '
          If Not FTP���üũ(SERVER_PATH) Then
             Call FTP_DisConnect
             MsgBox "������θ� ã���� �����ϴ�.(������������ ����)"
             Exit Sub
          End If
          '
          File_Path = ""
          If row = 1 Then File_Path = filenm1
          If row = 2 Then File_Path = filenm2
          If row = 3 Then File_Path = filenm3
          If row = 4 Then File_Path = filenm4
          '
          '���ε�
          If Not FTP_Upload(Comm1.fileName, SERVER_PATH, File_Path) Then
             Call FTP_DisConnect
             mkpoen05MDI.msg = "���� ���ε忡 ���� �Ͽ����ϴ�. (������������ ����)"
             Exit Sub
          End If
          '
          Call FTP_DisConnect
          '
       End If
       '
       '========================================
       ' ��Ʈ��ũ ����̹��� �̿��� ���� ���ε�
       '========================================
       ' If Row = 1 Then FileCopy Comm1.filename, N_Driver & ":\" & filenm1
       ' If Row = 2 Then FileCopy Comm1.filename, N_Driver & ":\" & filenm2
       ' If Row = 3 Then FileCopy Comm1.filename, N_Driver & ":\" & filenm3
       ' If Row = 4 Then FileCopy Comm1.filename, N_Driver & ":\" & filenm4
        '
        '
        sss = "      update man_tooltesthd"
        If row = 1 Then sss = sss & "   set tth_file1 = '" & filenm1 & "'"
        If row = 2 Then sss = sss & "   set tth_file2 = '" & filenm2 & "'"
        If row = 3 Then sss = sss & "   set tth_file3 = '" & filenm3 & "'"
        If row = 4 Then sss = sss & "   set tth_file4 = '" & filenm4 & "'"
        sss = sss & "  where tth_dat = " & txt_dat1
        sss = sss & "    and tth_seq = " & txt_seq1
        '
        db.Execute sss, 64
        '
        '
        spd_file12.Col = 1
        If row = 1 Then spd_file12.Text = filenm1
        If row = 2 Then spd_file12.Text = filenm2
        If row = 3 Then spd_file12.Text = filenm3
        If row = 4 Then spd_file12.Text = filenm4
        '
        spd_file12.Col = 3:  spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "����"
        '
        msg_display ("÷�������� ��ϵǾ����ϴ�!")
        '
        Exit Sub
    End If
    
    '÷������ ����
    If spd_file12.TypeButtonText = "����" And Col = 3 Then
       '
       sss = "select tth_file1, tth_file2, tth_file3, tth_file4"
       sss = sss & " from man_tooltesthd "
       sss = sss & " where tth_dat = " & txt_dat1
       sss = sss & "   and tth_seq = " & txt_seq1
       '
       Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
       If Rs.RecordCount < 1 Then
          Rs.Close
          msg_display ("÷�λ����� ��Ϲ�ȣ�� �����ϴ�. ��Ϲ�ȣ�� Ȯ���ϼ���!")
           Exit Sub
       End If
       '
       If MsgBox("÷�������� �����Ͻðڽ��ϱ�?", vbYesNo) <> vbYes Then
           Rs.Close
           msg_display ("÷������ �����۾��� ��ҵǾ����ϴ�!")
           Exit Sub
       End If
       '
       '���� �������
       sss = "      update man_tooltesthd "
       If row = 1 Then sss = sss & "   set tth_file1 = null"
       If row = 2 Then sss = sss & "   set tth_file2 = null"
       If row = 3 Then sss = sss & "   set tth_file3 = null"
       If row = 4 Then sss = sss & "   set tth_file4 = null"
       sss = sss & "  where tth_dat = " & txt_dat1
       sss = sss & "    and tth_seq = " & txt_seq1
       '
       db.Execute sss, 64
       '
       '===========================
       ' FTP�� �̿��� ���� ����
       '===========================
       If FTP_Connection Then
          '
          If Not FTP���üũ(SERVER_PATH) Then
             Call FTP_DisConnect
             MsgBox "������θ� ã���� �����ϴ�.(������������ ����)"
             Exit Sub
          End If
          '
          If row = 1 Then File_Path = Rs!tth_file1
          If row = 2 Then File_Path = Rs!tth_file2
          If row = 3 Then File_Path = Rs!tth_file3
          If row = 4 Then File_Path = Rs!tth_file4
          '
          If Not FTP_Delete(SERVER_PATH & File_Path) Then
             Call FTP_DisConnect
             mkpoen05MDI.msg = "������ ������ �� �����ϴ�. (������������ ����)"
             MsgBox "������ ������ �� �����ϴ�. (������������ ����)"
             Exit Sub
          End If
          '
          Call FTP_DisConnect
          '
       End If
       '
       '======================================
       ' ��Ʈ��ũ ����̹��� �̿��� ���� ����
       '======================================
      ' If Row = 1 Then Kill (N_Driver & ":\" & Rs!tth_file1)
      ' If Row = 2 Then Kill (N_Driver & ":\" & Rs!tth_file2)
      ' If Row = 3 Then Kill (N_Driver & ":\" & Rs!tth_file3)
      ' If Row = 4 Then Kill (N_Driver & ":\" & Rs!tth_file4)
       '
       spd_file12.Col = 1: spd_file12.Text = ""
       spd_file12.Col = 3:  spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "���"
       Rs.Close
       '
       msg_display ("÷�������� �����Ǿ����ϴ�!")
       '
       Exit Sub
       '
    End If
    '
    '�������� ����
    If spd_file12.TypeButtonText = "����" And Col = 4 Then
       '
       spd_file12.row = row: spd_file12.Col = 1
       '
       '�����̸� ����
       If Len(spd_file12.Text) < 5 Then
           msg_display ("��ϵ� ÷�������� �����ϴ�!")
           Exit Sub
       End If
       '
       '�������� ����
      ' exists = ExistFile(N_Driver & ":\" & spd_file12.Text)
      ' If exists = False Then
      '    msg_display ("������ ÷�������� �����ϴ�!")
      '    Exit Sub
      ' End If
       '
       Comm1.CancelError = True
       Comm1.fileName = spd_file12.Text
       Comm1.Filter = "*.pdf|*.pdf"
       Comm1.ShowSave
       '
       If Comm1.fileName = "" Then Exit Sub
       '
       F_NAME = Comm1.fileName
       '
       'lsSouurce = N_Driver & ":\" & spd_file12.Text
       lsDestination = F_NAME
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
          If Not FTP_Download(SERVER_PATH & spd_file12.Text, lsDestination) Then
             '
          End If
          '
          Call FTP_DisConnect
          '
       End If
       '
       '===========================================
       ' ��Ʈ��ũ ����̹��� �̿��� ���� �ٿ�ε�
       '===========================================
      ' Set fs = CreateObject("Scripting.FileSystemObject")
      ' fs.CopyFile lsSouurce, lsDestination '����
       '
       msg_display ("÷�������� ����Ǿ����ϴ�!")
       '
       Exit Sub
       '
    End If
    '
    Exit Sub
    '
err_rtn:
  mkpoen05MDI.msg = Err.Description
End Sub

'÷������ ����
Private Sub spd_file12_DblClick(ByVal Col As Long, ByVal row As Long)
    
On Error GoTo err_rtn
        
    Dim isrno As String     '�������� ��ȣ
    Dim filenm1 As String   '����1
    Dim filenm2 As String   '����2
    
    Dim fs            As Object 'Scripting.FileSystemObject ��ü
    Dim lsSouurce     As String '������ ���
    Dim lsDestination As String '������ġ ���
    Dim F_NAME     As String
    
    spd_file12.row = row: spd_file12.Col = 2
    
    '
    If Len(spd_file12.Text) > 2 Then
        
        spd_file12.row = row: spd_file12.Col = 1
        
        If Len(spd_file12.Text) < 5 Then
            msg_display ("��ϵ� ÷�������� �����ϴ�!")
            Exit Sub
        End If
        
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
          If FTP_Download(SERVER_PATH & spd_file12.Text, "c:\jpeg\" & spd_file12.Text) Then
         ' Else
             mkpoen05_view.WebBrowser1.Navigate "c:\jpeg\" & spd_file12.Text
             mkpoen05_view.txt_filenm = spd_file12.Text
          End If
          '
          Call FTP_DisConnect
          '
       End If
       '
       '======================================
       ' ��Ʈ��ũ ����̹��� �̿��� ���� View
       '======================================
       ' mkpoen05_view.WebBrowser1.Navigate N_Driver & ":\" & spd_file12.Text
       ' mkpoen05_view.txt_filenm = spd_file12.Text
        '
        Exit Sub
        '
    End If
    '
    Exit Sub
    
err_rtn:
  mkpoen05MDI.msg = Err.Description

End Sub

'----------------------
'÷������ üũ
'----------------------
Public Function Check_file(i As Integer, j As Integer)
    
    spd_file11.row = i: spd_file11.Col = j
    
    If Len(spd_file11.Text) > 0 Then
            
        If Right(spd_file11.Text, 3) <> "PDF" And Right(spd_file11.Text, 3) <> "pdf" And Right(spd_file11.Text, 3) <> "JPG" And Right(spd_file11.Text, 3) <> "jpg" Then
                msg_display (i & "��°��: ÷�������� PDF �Ǵ� JPG�� ��� �����մϴ�!")
                Check_file = "X"
                Exit Function
        End If
        '����� ������ �ִ��� chk
        exists = ExistFile(spd_file11.Text)
        If exists = False Then
            msg_display (i & "��°��: ��ο� ������ �������� �ʽ��ϴ�. ��θ� �ٽ� Ȯ���ϼ���!")
            Check_file = "X"
            Exit Function
        End If
        
        Check_file = "Y"
        Exit Function
        
    End If
    
    Check_file = "N"

End Function

'������������ üũ
Private Function ExistFile(FilePath As String) As Long
     If LenB(Dir$(FilePath)) Then
          ExistFile = 1&
     Else
          ExistFile = 0&
     End If
End Function

Private Sub txt_dat1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_seq1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txt_dat_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_seq_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub opt_n_su_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub



Private Sub txt_testno1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tdat1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_jdat1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_gdat_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_jseq1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_title1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_mcd1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_lotno1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chk_pyn1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chk_pyn2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chk_pyn3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chk_pyn4_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chk_pyn5_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tnm1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tsab1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


Private Sub txt_maker1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tipstd1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tipjil1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_holder1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_rcntmn1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_rcntmx1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_depth1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_movmn1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_movmx1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tct1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_pct1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub cmb_fluid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_qty1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_dan1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tldn1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_chdn1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txt_maker2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tipstd2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tipjil2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_holder2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_rcntmn2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_rcntmx2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_depth2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_movmn2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_movmx2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tct2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_pct2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub cmb_fluid2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_qty2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_dan2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tldn2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_chdn2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txt_maker3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tipstd3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tipjil3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_holder3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_rcntmn3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_rcntmx3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_depth3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_movmn3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_movmx3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tct3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_pct3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub cmb_fluid3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_qty3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_dan3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_tldn3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_chdn3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmb_result1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chk_ryn1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chk_ryn2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chk_ryn3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chk_ryn4_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chk_ryn5_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub cmb_rtyn1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
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
    
    sss = "select * from man_tooltesthd, man_tooltestds"
    sss = sss & " where tth_dat = ttd_dat"
    sss = sss & "   and tth_seq = ttd_seq"
    If Len(txt_sdat2) = 8 And Len(txt_edat2) = 8 Then
        sss = sss & " and tth_dat between " & txt_sdat2 & " and " & txt_edat2
    End If
    
    If Len(txt_test2) > 0 Then sss = sss & " and tth_testno = '" & txt_test2 & "'"
    
    If opt_OK.Value = True Or opt_NG.Value = True Then
        sss = sss & "    and tth_dat || tth_seq in(select ttd_dat || ttd_seq "
        sss = sss & "                                from man_tooltestds"
        If opt_OK.Value = True Then sss = sss & "   where ttd_result = 'OK')"
        If opt_NG.Value = True Then sss = sss & "   where ttd_result = 'NG')"
    End If

    'sss = sss & "  order by tth_dat, tth_seq, ttd_lno"
    sss = sss & "  order by tth_testno, tth_dat, tth_seq, ttd_lno"
                            
    sprd2.MaxRows = 0: cnt = 1
    
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Rs.RecordCount < 1 Then
        Rs.Close
        Call msg_display("��ϵ� ������ �����ϴ�!")
        Exit Sub
    End If
    '
    sprd2.MaxRows = 0: cnt = 0
    '
    Do While Not Rs.EOF
       '
       cnt = cnt + 1: sprd2.MaxRows = cnt: sprd2.row = cnt
       '
       If cnt > 1 Then
          If chk_ttdno <> Rs!tth_dat & Rs!tth_seq Then
             cnt = cnt + 1: sprd2.MaxRows = cnt: sprd2.row = cnt
             '
             sprd2.AddCellSpan 1, cnt - 1, 15, 1      'ǰ�� ����
             sprd2.RowHeight(cnt - 1) = 12
             sprd2.row = cnt - 1: sprd2.Col = 1: sprd2.BackColor = &H80000004
             sprd2.row = cnt
             chk_ttdno = Rs!tth_dat & Rs!tth_seq
          End If
       Else
          chk_ttdno = Rs!tth_dat & Rs!tth_seq
       End If
       '
       If Rs!ttd_lno = 1 Then
           '
           If Not IsNull(Rs!tth_dat) Then sprd2.Col = 1: sprd2.Text = Rs!tth_dat & "-" & Format(Rs!tth_seq, "000")
           If Not IsNull(Rs!ttd_lno) Then sprd2.Col = 2: sprd2.Text = Rs!ttd_lno
           '
           If Not IsNull(Rs!tth_testno) Then sprd2.Col = 3: sprd2.Text = Rs!tth_testno
           If Not IsNull(Rs!tth_tdat) Then sprd2.Col = 4:   sprd2.Text = Rs!tth_tdat
           If Not IsNull(Rs!tth_title) Then sprd2.Col = 5:  sprd2.Text = Rs!tth_title
           If Not IsNull(Rs!tth_tmcd) Then sprd2.Col = 6:   sprd2.Text = Rs!tth_tmcd
           If Not IsNull(Rs!tth_tlot) Then sprd2.Col = 7:   sprd2.Text = Rs!tth_tlot
           '
           For ii = 1 To 7
               sprd2.Col = ii
               sprd2.BackColor = 15794174
           Next ii
           '
       End If
       '
       If Rs!ttd_gbn = 1 Then sprd2.Col = 8: sprd2.Text = "����"
       If Rs!ttd_gbn = 2 Then sprd2.Col = 8: sprd2.Text = "�׽�Ʈ"
       '
       If Not IsNull(Rs!ttd_maker) Then sprd2.Col = 9:   sprd2.Text = Rs!ttd_maker
       If Not IsNull(Rs!ttd_tipstd) Then sprd2.Col = 10: sprd2.Text = Rs!ttd_tipstd
       If Not IsNull(Rs!ttd_tipjil) Then sprd2.Col = 11: sprd2.Text = Rs!ttd_tipjil
       If Not IsNull(Rs!ttd_dan) Then sprd2.Col = 12:    sprd2.Text = Rs!ttd_dan
       If Not IsNull(Rs!ttd_tldn) Then sprd2.Col = 13:   sprd2.Text = Rs!ttd_tldn
       If Not IsNull(Rs!ttd_chdn) Then sprd2.Col = 14:   sprd2.Text = Rs!ttd_chdn
       '
       If Not IsNull(Rs!ttd_result) Then
          sprd2.Col = 15: sprd2.Text = Rs!ttd_result
          If Rs!ttd_result = "OK" Then
             sprd2.ForeColor = &HFF0000
          Else
             sprd2.ForeColor = &HFF&
          End If
       End If
       '
       Rs.MoveNext
       '
    Loop
    Rs.Close
    '
    Call msg_display("���� ������ Ȯ���ϼ���")
    txt_sdat2.SetFocus
    '
    Exit Sub
    '
err_rtn:
    MsgBox (Err.Description)
End Sub

Private Sub txt_test2_LostFocus()
    txt_test2.Text = UCase(Trim(txt_test2))
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    '
    If SSTab1.Tab = 0 Then
    
    ElseIf SSTab1.Tab = 1 Then
        txt_sdat2.SetFocus
    End If
    '
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

