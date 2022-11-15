VERSION 5.00
Object = "{3B930683-5AF1-4F07-9CE8-CA8063E1F3DD}#2.0#0"; "XButton.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form mkpoen05A 
   BackColor       =   &H80000004&
   BorderStyle     =   0  '없음
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
         Name            =   "굴림"
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
         TabCaption(0)   =   "1.TEST DATA 등록"
         TabPicture(0)   =   "mkpoen05A.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frm13"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "1.TEST DATA 조회"
         TabPicture(1)   =   "mkpoen05A.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label6"
         Tab(1).Control(1)=   "sprd2"
         Tab(1).Control(2)=   "Frame3"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "TEST DATA 결재관리"
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
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "굴림"
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
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "굴림"
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
               Text            =   "확 인"
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
               Caption         =   "등록일자"
               BackColor       =   14737632
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
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "굴림"
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
                  Caption         =   "전체"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림"
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
                     Name            =   "굴림"
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
                     Name            =   "굴림"
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
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "굴림"
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
               Appearance      =   0  '평면
               BeginProperty Font 
                  Name            =   "굴림"
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
               Text            =   "확 인"
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
               Caption         =   "등록일자"
               BackColor       =   14737632
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
               Caption         =   "결과구분"
               BackColor       =   14737632
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
                  Name            =   "굴림"
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
               Left            =   1170
               MaxLength       =   8
               TabIndex        =   1
               Top             =   240
               Width           =   1185
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
                  Caption         =   "첨부파일"
                  BackColor       =   8438015
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
               Begin VB.Label lbl_cmt1 
                  Caption         =   "▲ 더블클릭 첨부파일 보기"
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
                     Name            =   "굴림"
                     Size            =   9.01
                     Charset         =   129
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Begin VB.OptionButton opt_n_su 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "수동"
                     Height          =   180
                     Left            =   870
                     TabIndex        =   8
                     TabStop         =   0   'False
                     Top             =   60
                     Width           =   705
                  End
                  Begin VB.OptionButton opt_n_ja 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "자동"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000F&
                  Enabled         =   0   'False
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
                  Left            =   2355
                  MaxLength       =   3
                  TabIndex        =   6
                  Top             =   210
                  Width           =   510
               End
               Begin VB.TextBox txt_dat 
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000F&
                  Enabled         =   0   'False
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
                  Left            =   1140
                  MaxLength       =   8
                  TabIndex        =   5
                  Top             =   210
                  Width           =   1185
               End
               Begin VB.TextBox txt_lotno1 
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Caption         =   "5.기타"
                  Height          =   375
                  Left            =   3150
                  TabIndex        =   26
                  Top             =   4860
                  Width           =   1365
               End
               Begin VB.CheckBox chk_pyn4 
                  Caption         =   "4.공구비 절감"
                  Height          =   375
                  Left            =   3150
                  TabIndex        =   24
                  Top             =   4605
                  Width           =   1425
               End
               Begin VB.CheckBox chk_pyn3 
                  Caption         =   "3.시간 단축"
                  Height          =   375
                  Left            =   1620
                  TabIndex        =   27
                  Top             =   5130
                  Width           =   1365
               End
               Begin VB.CheckBox chk_pyn2 
                  Caption         =   "2.칩 처리"
                  Height          =   375
                  Left            =   1620
                  TabIndex        =   25
                  Top             =   4860
                  Width           =   1215
               End
               Begin VB.CheckBox chk_pyn1 
                  Caption         =   "1.공구 수명"
                  Height          =   405
                  Left            =   1620
                  TabIndex        =   23
                  Top             =   4590
                  Width           =   1245
               End
               Begin VB.TextBox txt_bpcd1 
                  Appearance      =   0  '평면
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Caption         =   "TEST 제목"
                  BackColor       =   14737632
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
                  Caption         =   " 사    용           기    계 "
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
                  Caption         =   "장비코드"
                  BackColor       =   14737632
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
                     Name            =   "굴림"
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
                  Caption         =   "부서명"
                  BackColor       =   14737632
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
                  Caption         =   "        피           삭           재      "
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
                  Caption         =   "부품코드"
                  BackColor       =   14737632
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
                  Caption         =   "자재코드"
                  BackColor       =   14737632
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
                  Caption         =   "TSET 목적"
                  BackColor       =   14737632
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
                     Name            =   "굴림"
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
                  Caption         =   "테스트 일자"
                  BackColor       =   14737632
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
                  Caption         =   "접수번호"
                  BackColor       =   14737632
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
                  Caption         =   "작업자"
                  BackColor       =   14737632
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
                     Name            =   "굴림"
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
                  Caption         =   "등록번호"
                  BackColor       =   14737632
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
            End
            Begin VB.Frame Frame1 
               Height          =   7875
               Left            =   5100
               TabIndex        =   99
               Top             =   600
               Width           =   8355
               Begin VB.TextBox txt_movmn1 
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                     Name            =   "굴림"
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
                  Text            =   "1.수용성"
                  Top             =   3610
                  Width           =   1890
               End
               Begin VB.ComboBox cmb_fluid2 
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Text            =   "1.수용성"
                  Top             =   3610
                  Width           =   1890
               End
               Begin VB.ComboBox cmb_fluid1 
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Text            =   "1.수용성"
                  Top             =   3610
                  Width           =   1890
               End
               Begin VB.TextBox txt_pct3 
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Caption         =   "5.기타"
                  Height          =   405
                  Left            =   1770
                  TabIndex        =   88
                  Top             =   6930
                  Width           =   1695
               End
               Begin VB.CheckBox chk_ryn4 
                  Caption         =   "4.시간 단축"
                  Height          =   405
                  Left            =   1770
                  TabIndex        =   87
                  Top             =   6660
                  Width           =   1725
               End
               Begin VB.CheckBox chk_ryn3 
                  Caption         =   "3.공구비 절감"
                  Height          =   405
                  Left            =   1770
                  TabIndex        =   86
                  Top             =   6390
                  Width           =   1785
               End
               Begin VB.CheckBox chk_ryn2 
                  Caption         =   "2.칩처리 양호"
                  Height          =   405
                  Left            =   1770
                  TabIndex        =   85
                  Top             =   6120
                  Width           =   1665
               End
               Begin VB.CheckBox chk_ryn1 
                  Caption         =   "1.공구수명 연장"
                  Height          =   405
                  Left            =   1770
                  TabIndex        =   84
                  Top             =   5850
                  Width           =   1635
               End
               Begin VB.ComboBox cmb_rtyn1 
                  BeginProperty Font 
                     Name            =   "굴림"
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
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Alignment       =   1  '오른쪽 맞춤
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Appearance      =   0  '평면
                  BackColor       =   &H8000000E&
                  BeginProperty Font 
                     Name            =   "굴림"
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
                  Caption         =   " 사    용           공    구 "
                  BackColor       =   12640511
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "굴림"
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
                  Caption         =   " 절    삭           조    건 "
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
                     Name            =   "굴림"
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
                  Caption         =   "TIP 규격"
                  BackColor       =   14737632
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
                  Caption         =   "TIP 재질"
                  BackColor       =   14737632
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
                     Name            =   "굴림"
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
                  Caption         =   "분당 회전수"
                  BackColor       =   14737632
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
                  Caption         =   "절입량"
                  BackColor       =   14737632
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
                  Caption         =   "이송(mm/rev)"
                  BackColor       =   14737632
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
                     Name            =   "굴림"
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
                  Caption         =   "절삭유제"
                  BackColor       =   14737632
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
                  Caption         =   "가공수"
                  BackColor       =   14737632
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
                  Caption         =   "공구단가/EA"
                  BackColor       =   14737632
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
                  Caption         =   "공구비/Corner"
                  BackColor       =   14737632
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
                  Caption         =   "공구교환비"
                  BackColor       =   14737632
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
                  Caption         =   "기 존"
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
                  Caption         =   "테스트1"
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
                  Caption         =   "테스트2"
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
                  Caption         =   "결 과"
                  BackColor       =   14737632
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
                  Caption         =   "이 유"
                  BackColor       =   14737632
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
                  Caption         =   "재 TEST"
                  BackColor       =   14737632
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
                  Caption         =   "테스트2 작성시 클릭하세요 ▼"
                  Height          =   285
                  Left            =   5350
                  TabIndex        =   145
                  Top             =   150
                  Width           =   2805
               End
               Begin VB.Label Label1 
                  Caption         =   "▼ 결과를 선택하세요(원모양 클릭)"
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
               Text            =   "등 록"
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
               Text            =   "수 정"
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
               Text            =   "삭 제"
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
               Text            =   "조회"
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
               Text            =   "결재용 출 력"
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
               ToolTipBodyText =   "출 력"
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
            Caption         =   "▼ 더블클릭: 세부내역 확인"
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
    Const SERVER_PATH As String = "/공구테스트_DATA/"
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
        Call msg_display("등록된 내역이 없습니다!")
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
       
       If Not IsNull(Rs!apl_4sab) Then sprd3.Col = 11: sprd3.Text = "사장님"
       If Rs!apl_4yn = "Y" Then sprd3.Col = 11: sprd3.Text = "사장님" & Chr(13) & Rs!apl_4dat: sprd3.BackColor = RGB(255, 214, 255)
       
       If Rs!apl_4dat = 99999999 Then sprd3.Col = 11: sprd3.Text = "대결" & Chr(13) & Rs!apl_alldat: sprd3.BackColor = RGB(255, 214, 255)
       
       sprd3.RowHeight(sprd3.row) = sprd3.MaxTextRowHeight(sprd3.row) + 2
       
       Rs.MoveNext
       '
    Loop
    Rs.Close
    
End Sub

Private Sub Form_Load()
    
    '결과
    cmb_result1.Clear
    cmb_result1.AddItem ""
    cmb_result1.AddItem "O.K"
    cmb_result1.AddItem "N.G"
    cmb_result1.ListIndex = 0
    
    cmb_fluid1.Clear
    cmb_fluid1.AddItem ""
    cmb_fluid1.AddItem "1.수용성"
    cmb_fluid1.AddItem "2.비수용성"
    cmb_fluid1.ListIndex = 0
    
    cmb_fluid2.Clear
    cmb_fluid2.AddItem ""
    cmb_fluid2.AddItem "1.수용성"
    cmb_fluid2.AddItem "2.비수용성"
    cmb_fluid2.ListIndex = 0
    
    cmb_fluid3.Clear
    cmb_fluid3.AddItem ""
    cmb_fluid3.AddItem "1.수용성"
    cmb_fluid3.AddItem "2.비수용성"
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
'TAB1. TEST DATA 등록
'===================================

'------------------------------------
'내역조회
'------------------------------------
Private Sub btn_view1_Click()
    
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
        Call msg_display("등록된 내역이 없습니다!")
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
        
        '기존공구
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
        
        '테스트1
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
        
        '테스트2
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
    
    Call msg_display("조회완료!")
    
End Sub

    
'------------------------------------
'내역등록
'------------------------------------
Private Sub btn_add1_Click()
    '
    On Error GoTo err_rtn
    '
    Dim Dat As Double
    Dim seq As Double
    
    Dim FileChk1 As String   '첨부파일1
    Dim FileChk2 As String   '첨부파일2
    Dim FileChk3 As String   '첨부파일2
    Dim FileChk4 As String   '첨부파일2
    Dim memo     As String   '쪽지발송 메모
   
    If Job_Level < 9 Then
        MsgBox ("작업권한이 없습니다!")
        Exit Sub
    End If
    
    '첨부파일확인1
    FileChk1 = Check_file(1, 2)
    If FileChk1 = "X" Then Exit Sub
    FileChk2 = Check_file(2, 2)
    If FileChk2 = "X" Then Exit Sub
    FileChk3 = Check_file(3, 2)
    If FileChk3 = "X" Then Exit Sub
    FileChk4 = Check_file(4, 2)
    If FileChk4 = "X" Then Exit Sub
    
    If Check_Insert_Data = 9 Then Exit Sub
    
    
    '등록일자 자동부여
    If opt_n_ja.Value = True Then
    '날짜조회
        Dat = today("YYYYMMDD")
        
        '===================================================================
        '순번조회
        sss = "      select nvl(max(tth_seq),0) + 1 as seq from man_tooltesthd"
        sss = sss & " where tth_dat = " & Dat
        Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        seq = Rs!seq
        Rs.Close
        
    '등록일자 수동부여(지난내역 소급등록)
    Else
        '등록번호 확인(수동 날짜입력)
        If IsDate(Left(txt_dat, 4) & "/" & Mid(txt_dat, 5, 2) & "/" & Right(txt_dat, 2)) = False Then
            Call msg_display("등록번호 날짜를 확인하세요! (등록실패)")
            txt_dat.SetFocus
            Exit Sub
        End If
        
        '등록번호 확인(수동 순번입력)
        If Val(Trim(txt_seq)) < 1 Or IsNumeric(txt_seq) = False Then
            Call msg_display("등록번호 순번을 확인하세요! (등록실패)")
            txt_seq.SetFocus
            Exit Sub
        End If
        
        '기존등록번호 확인
        sss = "       select * "
        sss = sss & "   from man_tooltesthd"
        sss = sss & "  where tth_dat = " & txt_dat
        sss = sss & "    and tth_seq = " & txt_seq
                    
        Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        If Rs.RecordCount > 0 Then
            Rs.Close
            Call msg_display("기존등록번호가 존재합니다! (등록실패)")
            txt_dat.SetFocus
            Exit Sub
        End If
    
        Rs.Close
        
        Dat = txt_dat
        seq = Val(Trim(txt_seq))
        
    End If
        
    Ws.BeginTrans
    
    '=======================
    'HEAD입력
    '=======================
    sss = "insert into man_tooltesthd("
    sss = sss & "         tth_dat,"     '등록일자
    sss = sss & "         tth_seq,"     '등록순번
    sss = sss & "         tth_testno,"  'TEST NO.
    sss = sss & "         tth_title,"   '테스트 제목
    sss = sss & "         tth_tlot,"    '테스트 적용 LOT No.
    sss = sss & "         tth_tmcd,"    '테스트 장비코드
    sss = sss & "         tth_tsab,"    '테스트 작업자
    sss = sss & "         tth_tdat,"    '테스트 일자
    sss = sss & "         tth_jubno,"    '접수일자
    sss = sss & "         tth_pyn1,"    'Y/N - 1.공구수명
    sss = sss & "         tth_pyn2,"    'Y/N - 2.칩처리
    sss = sss & "         tth_pyn3,"    'Y/N - 3.시간단축
    sss = sss & "         tth_pyn4,"    'Y/N - 4.공구비
    sss = sss & "         tth_pyn5,"    'Y/N - 5.기타

    sss = sss & "         tth_sab,"     '입력자
    sss = sss & "         tth_rmk,"     '비고
    sss = sss & "         tth_cmt,"     '평가
                            
    sss = sss & "         tth_file1,"   '파일1
    sss = sss & "         tth_file2,"   '파일2
    sss = sss & "         tth_file3,"   '파일3
    sss = sss & "         tth_file4,"   '파일4
    
    sss = sss & "         tth_indte,"   '입력일자
    sss = sss & "         tth_updte"    '수정일자
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
    
    '============첨부파일
    '===========================
    ' FTP를 이용한 파일 업로드
    '===========================
    If FTP_Connection Then
       '
       If Not FTP경로체크(SERVER_PATH) Then
          Call FTP_DisConnect
          Ws.Rollback
          MsgBox "서버경로를 찾을수 없습니다.(정보관리센터 문의), 첨부파일 오류"
          Exit Sub
       End If
       '
       '첨부파일1(이미지파일)
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
                MsgBox "파일 업로드에 실패 하였습니다. (정보관리센터 문의)"
                Exit Sub
             End If
             'FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-001" & ".PDF"  '네트웍폴더에 파일저장
          ElseIf UCase(Right(spd_file11.Text, 3)) = "JPG" Then
             sss = sss & "       '" & Dat & Format(seq, "000") & "-001" & ".JPG" & "',"
             '
             File_Path = Dat & Format(seq, "000") & "-001" & ".JPG"
             '
             If Not FTP_Upload(spd_file11.Text, SERVER_PATH, File_Path) Then
                Ws.Rollback
                Call FTP_DisConnect
                MsgBox "파일 업로드에 실패 하였습니다. (정보관리센터 문의)"
                Exit Sub
             End If
            ' FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-001" & ".JPG"  '네트웍폴더에 파일저장
          End If
       Else
           sss = sss & "       null,"
       End If
    
       '첨부파일2
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
                MsgBox "파일 업로드에 실패 하였습니다. (정보관리센터 문의)"
                Exit Sub
             End If
            ' FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-002" & ".PDF"  '네트웍폴더에 파일저장
          ElseIf UCase(Right(spd_file11.Text, 3)) = "JPG" Then
             sss = sss & "       '" & Dat & Format(seq, "000") & "-002" & ".JPG" & "'," '파일1
             '
             File_Path = Dat & Format(seq, "000") & "-002" & ".JPG"
             '
             If Not FTP_Upload(spd_file11.Text, SERVER_PATH, File_Path) Then
                Ws.Rollback
                Call FTP_DisConnect
                MsgBox "파일 업로드에 실패 하였습니다. (정보관리센터 문의)"
                Exit Sub
             End If
            ' FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-002" & ".JPG"  '네트웍폴더에 파일저장
          End If
       Else
          sss = sss & "       null,"
       End If
       '
       '첨부파일3
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
                MsgBox "파일 업로드에 실패 하였습니다. (정보관리센터 문의)"
                Exit Sub
             End If
            ' FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-003" & ".PDF"  '네트웍폴더에 파일저장
          ElseIf UCase(Right(spd_file11.Text, 3)) = "JPG" Then
             sss = sss & "       '" & Dat & Format(seq, "000") & "-003" & ".JPG" & "'," '파일1
             '
             File_Path = Dat & Format(seq, "000") & "-003" & ".JPG"
             '
             If Not FTP_Upload(spd_file11.Text, SERVER_PATH, File_Path) Then
                Ws.Rollback
                Call FTP_DisConnect
                MsgBox "파일 업로드에 실패 하였습니다. (정보관리센터 문의)"
                Exit Sub
             End If
            ' FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-003" & ".JPG"  '네트웍폴더에 파일저장
          End If
       Else
          sss = sss & "       null,"
       End If
       '
       '첨부파일4
       If FileChk4 = "Y" Then
          '
          spd_file11.row = 4: spd_file11.Col = 2
          If UCase(Right(spd_file11.Text, 3)) = "PDF" Then
             sss = sss & "       '" & Dat & Format(seq, "000") & "-004" & ".PDF" & "'," '파일1
             '
             File_Path = Dat & Format(seq, "000") & "-004" & ".PDF"
             '
             If Not FTP_Upload(spd_file11.Text, SERVER_PATH, File_Path) Then
                Ws.Rollback
                Call FTP_DisConnect
                MsgBox "파일 업로드에 실패 하였습니다. (정보관리센터 문의)"
                Exit Sub
             End If
            ' FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-004" & ".PDF"  '네트웍폴더에 파일저장
          ElseIf UCase(Right(spd_file11.Text, 3)) = "JPG" Then
             sss = sss & "       '" & Dat & Format(seq, "000") & "-004" & ".JPG" & "'," '파일1
             '
             File_Path = Dat & Format(seq, "000") & "-004" & ".JPG"
             '
             If Not FTP_Upload(spd_file11.Text, SERVER_PATH, File_Path) Then
                Ws.Rollback
                Call FTP_DisConnect
                MsgBox "파일 업로드에 실패 하였습니다. (정보관리센터 문의)"
                Exit Sub
             End If
            ' FileCopy spd_file11.Text, N_Driver & ":\" & dat & Format(seq, "000") & "-004" & ".JPG"  '네트웍폴더에 파일저장
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
    'DESC입력
    '=======================
    '---------------------->
    '▼기존공구▼
    '---------------------->
    sss = "insert into man_tooltestds("
    sss = sss & "         ttd_dat,"     '등록일자
    sss = sss & "         ttd_seq,"     '등록순번
    sss = sss & "         ttd_lno,"     '등록행번
    sss = sss & "         ttd_gbn,"     '1:기존공구/2:테스트공구
    sss = sss & "         ttd_ryn,"     'Y/N - 결과선택
    
    sss = sss & "         ttd_maker,"   '제조사
    sss = sss & "         ttd_tipstd,"  '팁 규격/코드
    sss = sss & "         ttd_tipjil,"  '팁 재질
    sss = sss & "         ttd_holder,"  '적용홀더
    sss = sss & "         ttd_rcntmn,"  '분당회전 최소
    sss = sss & "         ttd_rcntmx,"  '분당회전 최대
    sss = sss & "         ttd_movmn,"   '이송(MM/REV) 최소
    sss = sss & "         ttd_movmx,"   '이송(MM/REV) 최대
    sss = sss & "         ttd_tct,"     'TCT
    sss = sss & "         ttd_pct,"     'PCT
                        
    sss = sss & "         ttd_depth,"   '절삭깊이
    
    sss = sss & "         ttd_fluid,"   '절삭유제
    sss = sss & "         ttd_qty,"     '가공수량
    sss = sss & "         ttd_dan,"     '단가/EA
    sss = sss & "         ttd_tldn,"    '공구비/EA
    sss = sss & "         ttd_chdn,"    '교환비/EA
    sss = sss & "         ttd_result,"  '결과 OK/NG
    sss = sss & "         ttd_ryn1,"    'Y/N
    sss = sss & "         ttd_ryn2,"    'Y/N
    sss = sss & "         ttd_ryn3,"    'Y/N
    sss = sss & "         ttd_ryn4,"    'Y/N
    sss = sss & "         ttd_ryn5,"    'Y/N
    sss = sss & "         ttd_rtyn, "   '재 테스트 여부
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
    
    If Val(txt_rcntmn1) > 0 Then    '분당회전수 최소값
        sss = sss & "        " & Val(txt_rcntmn1) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_rcntmx1) > 0 Then    '분당회전수 최대값
        sss = sss & "        " & Val(txt_rcntmx1) & ","
    Else
        sss = sss & "        null,"
    End If
                                                        
    If Val(txt_movmn1) > 0 Then    '이송 최소값
        sss = sss & "        " & Val(txt_movmn1) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_movmx1) > 0 Then    '이송 최대값
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
    
    If Val(txt_depth1) > 0 Then    '절삭깊이
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
    '▼TEST1▼
    '---------------------->
    sss = "insert into man_tooltestds("
    sss = sss & "         ttd_dat,"     '등록일자
    sss = sss & "         ttd_seq,"     '등록순번
    sss = sss & "         ttd_lno,"     '등록행번
    sss = sss & "         ttd_gbn,"     '1:기존공구/2:테스트공구
    sss = sss & "         ttd_ryn,"     'Y/N - 결과선택
    
    sss = sss & "         ttd_maker,"   '제조사
    sss = sss & "         ttd_tipstd,"  '팁 규격/코드
    sss = sss & "         ttd_tipjil,"  '팁 재질
    sss = sss & "         ttd_holder,"  '적용홀더
    sss = sss & "         ttd_rcntmn,"  '분당회전 최소
    sss = sss & "         ttd_rcntmx,"  '분당회전 최대
    sss = sss & "         ttd_movmn,"   '이송(MM/REV) 최소
    sss = sss & "         ttd_movmx,"   '이송(MM/REV) 최대
    sss = sss & "         ttd_tct,"     'TCT
    sss = sss & "         ttd_pct,"     'PCT
                          
    sss = sss & "         ttd_depth,"   '절삭깊이
                          
    sss = sss & "         ttd_fluid,"   '절삭유제
    sss = sss & "         ttd_qty,"     '가공수량
    sss = sss & "         ttd_dan,"     '단가/EA
    sss = sss & "         ttd_tldn,"    '공구비/EA
    sss = sss & "         ttd_chdn,"    '교환비/EA
    sss = sss & "         ttd_result,"  '결과 OK/NG
    sss = sss & "         ttd_ryn1,"    'Y/N
    sss = sss & "         ttd_ryn2,"    'Y/N
    sss = sss & "         ttd_ryn3,"    'Y/N
    sss = sss & "         ttd_ryn4,"    'Y/N
    sss = sss & "         ttd_ryn5,"    'Y/N
    sss = sss & "         ttd_rtyn, "   '재 테스트 여부
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
    
    If Val(txt_rcntmn2) > 0 Then    '분당회전수 최소값
        sss = sss & "        " & Val(txt_rcntmn2) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_rcntmx2) > 0 Then    '분당회전수 최대값
        sss = sss & "        " & Val(txt_rcntmx2) & ","
    Else
        sss = sss & "        null,"
    End If
                                                        
    If Val(txt_movmn2) > 0 Then    '이송 최소값
        sss = sss & "        " & Val(txt_movmn2) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_movmx2) > 0 Then    '이송 최대값
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
    
    If Val(txt_depth2) > 0 Then    '절삭깊이
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
    '▼TEST2▼
    '---------------------->
    
    If chk_test2.Value = 1 Then
    
    
        sss = "insert into man_tooltestds("
        sss = sss & "         ttd_dat,"     '등록일자
        sss = sss & "         ttd_seq,"     '등록순번
        sss = sss & "         ttd_lno,"     '등록행번
        sss = sss & "         ttd_gbn,"     '1:기존공구/2:테스트공구
        sss = sss & "         ttd_ryn,"     'Y/N - 결과선택
        
        sss = sss & "         ttd_maker,"   '제조사
        sss = sss & "         ttd_tipstd,"  '팁 규격/코드
        sss = sss & "         ttd_tipjil,"  '팁 재질
        sss = sss & "         ttd_holder,"  '적용홀더
        sss = sss & "         ttd_rcntmn,"  '분당회전 최소
        sss = sss & "         ttd_rcntmx,"  '분당회전 최대
        sss = sss & "         ttd_movmn,"   '이송(MM/REV) 최소
        sss = sss & "         ttd_movmx,"   '이송(MM/REV) 최대
        sss = sss & "         ttd_tct,"     'TCT
        sss = sss & "         ttd_pct,"     'PCT
                                            
        sss = sss & "         ttd_depth,"   '절삭깊이
                                                            
        sss = sss & "         ttd_fluid,"   '절삭유제
        sss = sss & "         ttd_qty,"     '가공수량
        sss = sss & "         ttd_dan,"     '단가/EA
        sss = sss & "         ttd_tldn,"    '공구비/EA
        sss = sss & "         ttd_chdn,"    '교환비/EA
        sss = sss & "         ttd_result,"  '결과 OK/NG
        sss = sss & "         ttd_ryn1,"    'Y/N
        sss = sss & "         ttd_ryn2,"    'Y/N
        sss = sss & "         ttd_ryn3,"    'Y/N
        sss = sss & "         ttd_ryn4,"    'Y/N
        sss = sss & "         ttd_ryn5,"    'Y/N
        sss = sss & "         ttd_rtyn, "   '재 테스트 여부
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
        
        If Val(txt_rcntmn3) > 0 Then    '분당회전수 최소값
            sss = sss & "        " & Val(txt_rcntmn3) & ","
        Else
            sss = sss & "        null,"
        End If
        If Val(txt_rcntmx3) > 0 Then    '분당회전수 최대값
            sss = sss & "        " & Val(txt_rcntmx3) & ","
        Else
            sss = sss & "        null,"
        End If
                                                            
        If Val(txt_movmn3) > 0 Then    '이송 최소값
            sss = sss & "        " & Val(txt_movmn3) & ","
        Else
            sss = sss & "        null,"
        End If
        If Val(txt_movmx3) > 0 Then    '이송 최대값
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
        
        If Val(txt_depth3) > 0 Then    '절삭깊이
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
    '담당결재 이승환대리로
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
    Call msg_display("등록되었습니다")
    '
    Exit Sub
    '
err_rtn:
    Ws.Rollback
    MsgBox (Err.Description)
End Sub

'등록내역 체크
Public Function Check_Insert_Data()
    
    Dim str As String
    
    'TEST NO.확인
    If Len(Trim(txt_testno1)) < 1 Then
        Call msg_display("TEST NO.를 확인하세요! (등록실패)")
        txt_testno1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    If LenB(StrConv(txt_testno1, vbFromUnicode)) > 20 Then
        Call msg_display("TEST NO. 문자길이를 확인하세요! (20byte제한)")
        txt_testno1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '테스트 일자확인
    If IsDate(Left(txt_tdat1, 4) & "/" & Mid(txt_tdat1, 5, 2) & "/" & Right(txt_tdat1, 2)) = False Then
        Call msg_display("테스트 일자를 확인하세요! (등록실패)")
        txt_tdat1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '접수번호 확인
    If IsDate(Left(txt_jdat1, 4) & "/" & Mid(txt_jdat1, 5, 2) & "/" & Right(txt_jdat1, 2)) = False Then
        Call msg_display("접수일자를 확인하세요! (등록실패)")
        txt_jdat1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '접수순번 확인
    If Len(Trim(txt_jseq1)) < 1 Or IsNumeric(txt_jseq1) = False Then
        Call msg_display("접수순번을 확인하세요! (등록실패)")
        txt_jseq1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    'TEST제목 확인
    If Len(Trim(txt_title1)) < 1 Then
        Call msg_display("TEST제목을 입력하세요! (등록실패)")
        txt_title1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    If LenB(StrConv(txt_title1, vbFromUnicode)) > 30 Then
        Call msg_display("TEST제목 문자길이를 확인하세요! (30byte제한)")
        txt_title1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '장비코드확인
    If Len(Trim(txt_mcd1)) < 1 Then
        Call msg_display("TEST 장비코드를 입력하세요! (등록실패)")
        txt_mcd1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '쿼리체크
    sss = "       select mhc_name, ems_mark, sok_name(mhc_sok) soknm "
    sss = sss & "   from man_machcd, eam_mast"
    sss = sss & "  where mhc_code = '" & Trim(txt_mcd1) & "'"
    sss = sss & "    and mhc_code = ems_mcd(+)"
                
    Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Ks.RecordCount < 1 Then
        Call msg_display("TEST 장비코드를 확인하세요! (등록실패)")
        txt_mcd1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
        
    Ks.Close
    
    'LOT NO확인
    If Len(Trim(txt_lotno1)) < 1 Then
        Call msg_display("LOT NO.를 확인하세요! (등록실패)")
        txt_lotno1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '쿼리체크
    sss = "       select dit_bpcd, dit_bpjil, dit_jacd, dit_jajil"
    sss = sss & "   from man_direct"
    sss = sss & "  where dit_lot = '" & Trim(txt_lotno1) & "'"
            
    Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Ks.RecordCount < 1 Then
        Call msg_display("LOT NO.를 확인하세요! (등록실패)")
        txt_lotno1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    Ks.Close
        
    '사번확인
    '접수순번 확인
    If Len(Trim(txt_tsab1)) < 1 Or IsNumeric(txt_tsab1) = False Then
        Call msg_display("작업자사번을 확인하세요! (등록실패)")
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
        Call msg_display("사번을 확인하세요! (등록실패)")
        txt_lotno1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    Ks.Close
    
    '=========================
    '▼기존 공구데이터 확인▼
    '=========================
    
    '------------------------>
    '기존공구
    '------------------------>
    'MAKER
    If Len(Trim(txt_maker1)) < 1 Then
        Call msg_display("[기존공구] MAKER를 입력하세요! (등록실패)")
        txt_maker1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_maker1, vbFromUnicode)) > 20 Then
        Call msg_display("[기존공구] MAKER 문자길이를 확인하세요! (20byte제한)")
        txt_maker1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    'TIP규격/코드
    If Len(Trim(txt_tipstd1)) < 1 Then
        Call msg_display("[기존공구] TIP규격을 입력하세요! (등록실패)")
        txt_tipstd1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_tipstd1, vbFromUnicode)) > 30 Then
        Call msg_display("[기존공구] TIP규격 문자길이를 확인하세요! (30byte제한)")
        txt_tipstd1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    'TIP재질
    If Len(Trim(txt_tipjil1)) < 1 Then
        Call msg_display("[기존공구] TIP재질을 입력하세요! (등록실패)")
        txt_tipjil1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_tipjil1, vbFromUnicode)) > 10 Then
        Call msg_display("[기존공구] TIP재질 문자길이를 확인하세요! (10byte제한)")
        txt_tipjil1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    'HOLDER
    If Len(Trim(txt_holder1)) < 1 Then
        Call msg_display("[기존공구] HOLDER 코드를 입력하세요! (등록실패)")
        txt_holder1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_holder1, vbFromUnicode)) > 30 Then
        Call msg_display("[기존공구] HOLDER 코드 문자길이를 확인하세요! (20byte제한)")
        txt_holder1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '분당회전수 최소값
    If Len(Trim(txt_rcntmn1)) > 0 Then
        If IsNumeric(txt_rcntmn1) = False Then
            Call msg_display("분당회전수 최소값은 숫자만 입력하세요! ")
            txt_rcntmn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_rcntmn1) > 9999 Or Val(txt_rcntmn1) < 0 Then
            Call msg_display("분당회전수 최소값은 1~9999 범위로 입력가능합니다! ")
            txt_rcntmn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '분당회전수 최대값
    If Len(Trim(txt_rcntmx1)) > 0 Then
        If IsNumeric(txt_rcntmx1) = False Then
            Call msg_display("분당회전수 최대값은 숫자만 입력하세요! ")
            txt_rcntmx1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_rcntmx1) > 9999 Or Val(txt_rcntmx1) < 0 Then
            Call msg_display("분당회전수 최대값은 1~9999 범위로 입력가능합니다! ")
            txt_rcntmx1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If

    '절삭깊이
    If Len(Trim(txt_depth1)) > 0 Then
        If IsNumeric(txt_depth1) = False Then
            Call msg_display("절삭깊이를 입력하세요! ")
            txt_depth1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_depth1) > 99 Or Val(txt_depth1) < 0 Then
            Call msg_display("절삭깊이 값은 0.01~99.99 범위로 입력가능합니다! ")
            txt_depth1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
        
    '이송 최소값
    If Len(Trim(txt_movmn1)) > 0 Then
        If IsNumeric(txt_movmn1) = False Then
            Call msg_display("이송수 최소값은 숫자만 입력하세요! ")
            txt_movmn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_movmn1) > 9999 Or Val(txt_movmn1) < 0 Then
            Call msg_display("이송 최소값은 0.01~9999 범위로 입력가능합니다! ")
            txt_movmn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '이송 최대값
    If Len(Trim(txt_movmx1)) > 0 Then
        If IsNumeric(txt_movmx1) = False Then
            Call msg_display("이송 최대값은 숫자만 입력하세요! ")
            txt_movmx1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_movmx1) > 9999 Or Val(txt_movmx1) < 0 Then
            Call msg_display("이송 최대값은 0.01~9999 범위로 입력가능합니다! ")
            txt_movmx1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    'TCP
    If Len(Trim(txt_tct1)) > 0 Then
        If IsNumeric(txt_tct1) = False Then
            Call msg_display("T.C/P값은 숫자만 입력하세요! ")
            txt_tct1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_tct1) > 99999 Or Val(txt_tct1) < 0 Then
            Call msg_display("T.C/P값은 1~99999 범위로 입력가능합니다! ")
            txt_tct1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    'PCT
    If Len(Trim(txt_pct1)) > 0 Then
        If IsNumeric(txt_pct1) = False Then
            Call msg_display("P.C/P값은 숫자만 입력하세요! ")
            txt_pct1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_pct1) > 99999 Or Val(txt_pct1) < 0 Then
            Call msg_display("P.C/P값은  1~99999 범위로 입력가능합니다! ")
            txt_pct1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '절삭유제 확인
    If Left(cmb_fluid1, 1) <> 1 And Left(cmb_fluid1, 2) <> 2 Then
        Call msg_display("절삭유제를 선택하세요! ")
        cmb_fluid1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '가공수확인
    If Len(Trim(txt_qty1)) > 0 Then
        If IsNumeric(txt_qty1) = False Then
            Call msg_display("가공수는 숫자만 입력하세요! ")
            txt_qty1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_qty1) > 99999 Or Val(txt_qty1) < 0 Then
            Call msg_display("가공수는  1~99999 범위로 입력가능합니다! ")
            txt_qty1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '공구단가/EA
    If Len(Trim(txt_dan1)) > 0 Then
        If IsNumeric(txt_dan1) = False Then
            Call msg_display("공구단가는 숫자만 입력하세요! ")
            txt_dan1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_dan1) > 99999999 Or Val(txt_dan1) < 0 Then
            Call msg_display("공구단가는  1~99999999 범위로 입력가능합니다! ")
            txt_dan1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '공구비/EA
    If Len(Trim(txt_tldn1)) > 0 Then
        If IsNumeric(txt_tldn1) = False Then
            Call msg_display("공구비는 숫자만 입력하세요! ")
            txt_tldn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_tldn1) > 99999999 Or Val(txt_tldn1) < 0 Then
            Call msg_display("공구비는  1~99999999 범위로 입력가능합니다! ")
            txt_tldn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '공구교환비
    If Len(Trim(txt_chdn1)) > 0 Then
        If IsNumeric(txt_chdn1) = False Then
            Call msg_display("교환비는 숫자만 입력하세요! ")
            txt_chdn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_chdn1) > 99999 Or Val(txt_chdn1) < 0 Then
            Call msg_display("교환비는  1~99999 범위로 입력가능합니다! ")
            txt_chdn1.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '------------------------>
    '▼테스트1
    '------------------------>
    'MAKER
    If Len(Trim(txt_maker2)) < 1 Then
        Call msg_display("[기존공구] MAKER를 입력하세요! (등록실패)")
        txt_maker2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_maker2, vbFromUnicode)) > 20 Then
        Call msg_display("[기존공구] MAKER 문자길이를 확인하세요! (20byte제한)")
        txt_maker2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    'TIP규격/코드
    If Len(Trim(txt_tipstd2)) < 1 Then
        Call msg_display("[기존공구] TIP규격을 입력하세요! (등록실패)")
        txt_tipstd2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_tipstd2, vbFromUnicode)) > 30 Then
        Call msg_display("[기존공구] TIP규격 문자길이를 확인하세요! (30byte제한)")
        txt_tipstd2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    'TIP재질
    If Len(Trim(txt_tipjil2)) < 1 Then
        Call msg_display("[기존공구] TIP재질을 입력하세요! (등록실패)")
        txt_tipjil2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_tipjil2, vbFromUnicode)) > 10 Then
        Call msg_display("[기존공구] TIP재질 문자길이를 확인하세요! (10byte제한)")
        txt_tipjil2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    'HOLDER
    If Len(Trim(txt_holder2)) < 1 Then
        Call msg_display("[기존공구] HOLDER 코드를 입력하세요! (등록실패)")
        txt_holder2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    If LenB(StrConv(txt_holder2, vbFromUnicode)) > 30 Then
        Call msg_display("[기존공구] HOLDER 코드 문자길이를 확인하세요! (20byte제한)")
        txt_holder2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '분당회전수 최소값
    If Len(Trim(txt_rcntmn2)) > 0 Then
        If IsNumeric(txt_rcntmn2) = False Then
            Call msg_display("분당회전수 최소값은 숫자만 입력하세요! ")
            txt_rcntmn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_rcntmn2) > 9999 Or Val(txt_rcntmn2) < 0 Then
            Call msg_display("분당회전수 최소값은 1~9999 범위로 입력가능합니다! ")
            txt_rcntmn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '분당회전수 최대값
    If Len(Trim(txt_rcntmx2)) > 0 Then
        If IsNumeric(txt_rcntmx2) = False Then
            Call msg_display("분당회전수 최대값은 숫자만 입력하세요! ")
            txt_rcntmx2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_rcntmx2) > 9999 Or Val(txt_rcntmx2) < 0 Then
            Call msg_display("분당회전수 최대값은 1~9999 범위로 입력가능합니다! ")
            txt_rcntmx2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '절삭깊이
    If Len(Trim(txt_depth2)) > 0 Then
        If IsNumeric(txt_depth2) = False Then
            Call msg_display("절삭깊이를 입력하세요! ")
            txt_depth2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_depth2) > 99 Or Val(txt_depth2) < 0 Then
            Call msg_display("절삭깊이 값은 0.01~99.99 범위로 입력가능합니다! ")
            txt_depth2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '이송 최소값
    If Len(Trim(txt_movmn2)) > 0 Then
        If IsNumeric(txt_movmn2) = False Then
            Call msg_display("이송수 최소값은 숫자만 입력하세요! ")
            txt_movmn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_movmn2) > 9999 Or Val(txt_movmn2) < 0 Then
            Call msg_display("이송 최소값은 0.01~9999 범위로 입력가능합니다! ")
            txt_movmn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '이송 최대값
    If Len(Trim(txt_movmx2)) > 0 Then
        If IsNumeric(txt_movmx2) = False Then
            Call msg_display("이송 최대값은 숫자만 입력하세요! ")
            txt_movmx2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_movmx2) > 9999 Or Val(txt_movmx2) < 0 Then
            Call msg_display("이송 최대값은 0.01~9999 범위로 입력가능합니다! ")
            txt_movmx2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    'TCP
    If Len(Trim(txt_tct2)) > 0 Then
        If IsNumeric(txt_tct2) = False Then
            Call msg_display("T.C/P값은 숫자만 입력하세요! ")
            txt_tct2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_tct2) > 99999 Or Val(txt_tct2) < 0 Then
            Call msg_display("T.C/P값은 1~99999 범위로 입력가능합니다! ")
            txt_tct2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    'PCT
    If Len(Trim(txt_pct2)) > 0 Then
        If IsNumeric(txt_pct2) = False Then
            Call msg_display("P.C/P값은 숫자만 입력하세요! ")
            txt_pct2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_pct2) > 99999 Or Val(txt_pct2) < 0 Then
            Call msg_display("P.C/P값은  1~99999 범위로 입력가능합니다! ")
            txt_pct2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '절삭유제 확인
    If Left(cmb_fluid2, 1) <> 1 And Left(cmb_fluid2, 2) <> 2 Then
        Call msg_display("절삭유제를 선택하세요! ")
        cmb_fluid2.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '가공수확인
    If Len(Trim(txt_qty2)) > 0 Then
        If IsNumeric(txt_qty2) = False Then
            Call msg_display("가공수 값은 숫자만 입력하세요! ")
            txt_qty2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_qty2) > 99999 Or Val(txt_qty2) < 0 Then
            Call msg_display("가공수는  1~99999 범위로 입력가능합니다! ")
            txt_qty2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
                            
    '공구단가/EA
    If Len(Trim(txt_dan2)) > 0 Then
        If IsNumeric(txt_dan2) = False Then
            Call msg_display("공구단가는 숫자만 입력하세요! ")
            txt_dan2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_dan2) > 99999999 Or Val(txt_dan2) < 0 Then
            Call msg_display("공구단가는  1~99999999 범위로 입력가능합니다! ")
            txt_dan2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
        
    '공구비/EA
    If Len(Trim(txt_tldn2)) > 0 Then
        If IsNumeric(txt_tldn1) = False Then
            Call msg_display("공구비는 숫자만 입력하세요! ")
            txt_tldn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_tldn2) > 99999999 Or Val(txt_tldn2) < 0 Then
            Call msg_display("공구비는  1~99999999 범위로 입력가능합니다! ")
            txt_tldn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
                                
    '공구교환비
    If Len(Trim(txt_chdn2)) > 0 Then
        If IsNumeric(txt_chdn2) = False Then
            Call msg_display("교환비는 숫자만 입력하세요! ")
            txt_chdn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If Val(txt_chdn2) > 99999 Or Val(txt_chdn2) < 0 Then
            Call msg_display("교환비는  1~99999 범위로 입력가능합니다! ")
            txt_chdn2.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    End If
    
    '------------------------>
    '테스트2
    '------------------------>
    
    If chk_test2.Value = 1 Then
    
        'MAKER
        If Len(Trim(txt_maker3)) < 1 Then
            Call msg_display("[기존공구] MAKER를 입력하세요! (등록실패)")
            txt_maker3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If LenB(StrConv(txt_maker3, vbFromUnicode)) > 20 Then
            Call msg_display("[기존공구] MAKER 문자길이를 확인하세요! (20byte제한)")
            txt_maker3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        
        'TIP규격/코드
        If Len(Trim(txt_tipstd3)) < 1 Then
            Call msg_display("[기존공구] TIP규격을 입력하세요! (등록실패)")
            txt_tipstd3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If LenB(StrConv(txt_tipstd3, vbFromUnicode)) > 30 Then
            Call msg_display("[기존공구] TIP규격 문자길이를 확인하세요! (30byte제한)")
            txt_tipstd3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        
        'TIP재질
        If Len(Trim(txt_tipjil3)) < 1 Then
            Call msg_display("[기존공구] TIP재질을 입력하세요! (등록실패)")
            txt_tipjil3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If LenB(StrConv(txt_tipjil3, vbFromUnicode)) > 10 Then
            Call msg_display("[기존공구] TIP재질 문자길이를 확인하세요! (10byte제한)")
            txt_tipjil3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        
        'HOLDER
        If Len(Trim(txt_holder3)) < 1 Then
            Call msg_display("[기존공구] HOLDER 코드를 입력하세요! (등록실패)")
            txt_holder3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        If LenB(StrConv(txt_holder3, vbFromUnicode)) > 30 Then
            Call msg_display("[기존공구] HOLDER 코드 문자길이를 확인하세요! (20byte제한)")
            txt_holder3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        
        '분당회전수 최소값
        If Len(Trim(txt_rcntmn3)) > 0 Then
            If IsNumeric(txt_rcntmn3) = False Then
                Call msg_display("분당회전수 최소값은 숫자만 입력하세요! ")
                txt_rcntmn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_rcntmn3) > 9999 Or Val(txt_rcntmn3) < 0 Then
                Call msg_display("분당회전수 최소값은 1~9999 범위로 입력가능합니다! ")
                txt_rcntmn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
        
        '분당회전수 최대값
        If Len(Trim(txt_rcntmx3)) > 0 Then
            If IsNumeric(txt_rcntmx3) = False Then
                Call msg_display("분당회전수 최대값은 숫자만 입력하세요! ")
                txt_rcntmx3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_rcntmx3) > 9999 Or Val(txt_rcntmx3) < 0 Then
                Call msg_display("분당회전수 최대값은 1~9999 범위로 입력가능합니다! ")
                txt_rcntmx3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
        
        '절삭깊이
        If Len(Trim(txt_depth3)) > 0 Then
            If IsNumeric(txt_depth3) = False Then
                Call msg_display("절삭깊이를 입력하세요! ")
                txt_depth3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_depth3) > 99 Or Val(txt_depth3) < 0 Then
                Call msg_display("절삭깊이 값은 0.01~99.99 범위로 입력가능합니다! ")
                txt_depth3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
            
        '이송 최소값
        If Len(Trim(txt_movmn3)) > 0 Then
            If IsNumeric(txt_movmn3) = False Then
                Call msg_display("이송수 최소값은 숫자만 입력하세요! ")
                txt_movmn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_movmn3) > 9999 Or Val(txt_movmn3) < 0 Then
                Call msg_display("이송 최소값은 0.01~9999 범위로 입력가능합니다! ")
                txt_movmn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
                                                        
        '이송 최대값
        If Len(Trim(txt_movmx3)) > 0 Then
            If IsNumeric(txt_movmx3) = False Then
                Call msg_display("이송 최대값은 숫자만 입력하세요! ")
                txt_movmx3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_movmx3) > 9999 Or Val(txt_movmx3) < 0 Then
                Call msg_display("이송 최대값은 0.01~99.99 범위로 입력가능합니다! ")
                txt_movmx3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
        
        'TCP
        If Len(Trim(txt_tct3)) > 0 Then
            If IsNumeric(txt_tct3) = False Then
                Call msg_display("T.C/P값은 숫자만 입력하세요! ")
                txt_tct3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_tct3) > 99999 Or Val(txt_tct3) < 0 Then
                Call msg_display("T.C/P값은 1~99999 범위로 입력가능합니다! ")
                txt_tct3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
        
        'PCT
        If Len(Trim(txt_pct3)) > 0 Then
            If IsNumeric(txt_pct3) = False Then
                Call msg_display("P.C/P값은 숫자만 입력하세요! ")
                txt_pct3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_pct3) > 99999 Or Val(txt_pct3) < 0 Then
                Call msg_display("P.C/P값은  1~99999 범위로 입력가능합니다! ")
                txt_pct3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
        
        '절삭유제 확인
        If Left(cmb_fluid3, 1) <> 1 And Left(cmb_fluid3, 1) <> 2 Then
            Call msg_display("절삭유제를 선택하세요! ")
            cmb_fluid3.SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        
        '가공수확인
        If Len(Trim(txt_qty3)) > 0 Then
            If IsNumeric(txt_qty2) = False Then
                Call msg_display("가공수는 숫자만 입력하세요! ")
                txt_qty3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_qty3) > 99999 Or Val(txt_qty3) < 0 Then
                Call msg_display("가공수는  1~99999 범위로 입력가능합니다! ")
                txt_qty3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
        
        '공구단가/EA
        If Len(Trim(txt_dan3)) > 0 Then
            If IsNumeric(txt_dan3) = False Then
                Call msg_display("공구단가는 숫자만 입력하세요! ")
                txt_dan3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_dan3) > 99999999 Or Val(txt_dan3) < 0 Then
                Call msg_display("공구단가는  1~99999999 범위로 입력가능합니다! ")
                txt_dan3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
        
        '공구비/EA
        If Len(Trim(txt_tldn3)) > 0 Then
            If IsNumeric(txt_tldn3) = False Then
                Call msg_display("공구비는 숫자만 입력하세요! ")
                txt_tldn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_tldn3) > 99999999 Or Val(txt_tldn3) < 0 Then
                Call msg_display("공구비는  1~99999999 범위로 입력가능합니다! ")
                txt_tldn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
                
        '공구교환비
        If Len(Trim(txt_chdn3)) > 0 Then
            If IsNumeric(txt_chdn3) = False Then
                Call msg_display("교환비는 숫자만 입력하세요! ")
                txt_chdn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
            If Val(txt_chdn3) > 99999 Or Val(txt_chdn3) < 0 Then
                Call msg_display("교환비는  1~99999 범위로 입력가능합니다! ")
                txt_chdn3.SetFocus
                Check_Insert_Data = 9
                Exit Function
            End If
        End If
    
    End If
    
    
    '============================
    '결과
    '============================
    
    If cmb_result1 <> "O.K" And cmb_result1 <> "N.G" Then
        Call msg_display("결과는 O.K 또는 N.G를 선택하세요!")
        cmb_result1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    If cmb_result1 = "N.G" Then
        
        If opt_ryn(1).Value <> True Then
            Call msg_display("결과가 N.G일때 기존 공구를 선택 후 등록하세요!")
            'opt_ryn(1).SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
    
    ElseIf cmb_result1 = "O.K" Then
        
        If opt_ryn(2).Value = False And opt_ryn(3).Value = False Then
            Call msg_display("결과가 O.K일때 테스트 공구를 선택 후 등록하세요!")
            'opt_ryn(2).SetFocus
            Check_Insert_Data = 9
            Exit Function
        End If
        
        
    End If
    
    If chk_ryn1.Value = 0 And chk_ryn2.Value = 0 And chk_ryn3.Value = 0 And _
       chk_ryn4.Value = 0 And chk_ryn5.Value = 0 Then
        Call msg_display("결과 이유는 최소한 1개이상 선택하세요!")
        cmb_result1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    If cmb_rtyn1 <> "Y" And cmb_rtyn1 <> "N" Then
        Call msg_display("재 TEST가능여부는 ""Y"" 또는 ""N""을 선택하세요!")
        cmb_rtyn1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    '결과선택
    If opt_ryn(1).Value = False And opt_ryn(2).Value = False And opt_ryn(3).Value = False Then
        Call msg_display("결과 적용공구를 선택하세요! (원모양 선택)")
        Check_Insert_Data = 9
        Exit Function
    End If
    
    sprd_rmk1.row = 1
    sprd_rmk1.Col = 1
    If LenB(StrConv(sprd_rmk1.Text, vbFromUnicode)) > 200 Then
        Call msg_display("비고내용의 문자길이가 깁니다.! (200byte제한)")
        sprd_rmk1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
    sprd_rmk1.Col = 2
    If LenB(StrConv(sprd_rmk1.Text, vbFromUnicode)) > 200 Then
        Call msg_display("평가내용 문자의 문자길이가 깁니다.! (200byte제한)")
        sprd_rmk1.SetFocus
        Check_Insert_Data = 9
        Exit Function
    End If
    
End Function

'-------------------------------
'내역수정
'-------------------------------
Private Sub btn_mod1_Click()
    '
    On Error GoTo err_rtn
    '
    If Job_Level < 9 Then
        MsgBox ("작업권한이 없습니다!")
        Exit Sub
    End If
    
    If IsNumeric(txt_dat1) = False Or Len(txt_dat1) <> 8 Then
        Call msg_display("등록 일자를 확인하세요!")
        txt_dat1.SetFocus
        Exit Sub
    End If
    '
    If IsNumeric(txt_seq1) = False Or Len(txt_seq1) < 1 Then
        Call msg_display("등록 순번을 확인하세요!")
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
       Call msg_display("등록된 내역이 없습니다. 수정취소!")
       Exit Sub
    End If
    '
    If Check_Insert_Data = 9 Then Exit Sub
    '
    '수정
    If MsgBox("(" & txt_dat1 & "-" & Format(txt_seq1, "000") & ") 등록된 내역을 수정하시겠습니까?", vbYesNo) <> vbYes Then
        Rs.Close
        msg_display ("수정 취소되었습니다.")
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
    sss = sss & "         tth_title = '" & txt_title1 & "',"                          '테스트 제목
    sss = sss & "         tth_tlot = '" & txt_lotno1 & "',"                           '테스트 적용 LOT No.
    sss = sss & "         tth_tmcd = '" & txt_mcd1 & "',"                             '테스트 장비코드
    sss = sss & "         tth_tsab = " & Val(txt_tsab1) & ","                         '테스트 작업자
    sss = sss & "         tth_tdat = " & Val(txt_tdat1) & ","                           '테스트 일자
    sss = sss & "         tth_jubno = " & txt_jdat1 & Format(txt_jseq1, "000") & ","   '접수일자
                    
    If chk_pyn1.Value = 1 Then
        sss = sss & "         tth_pyn1 = 'Y',"         'Y/N - 1.공구수명
    Else
        sss = sss & "         tth_pyn1 = 'N',"         'Y/N - 1.공구수명
    End If
    If chk_pyn2.Value = 1 Then
        sss = sss & "         tth_pyn2 = 'Y',"         'Y/N - 2.칩 처리
    Else
        sss = sss & "         tth_pyn2 = 'N',"         'Y/N - 2.칩 처리
    End If
    If chk_pyn3.Value = 1 Then
        sss = sss & "         tth_pyn3 = 'Y',"         'Y/N - 3.시간 단축
    Else
        sss = sss & "         tth_pyn3 = 'N',"         'Y/N - 3.시간 단축
    End If
    If chk_pyn4.Value = 1 Then
        sss = sss & "         tth_pyn4 = 'Y',"         'Y/N - 4.공구비 절감
    Else
        sss = sss & "         tth_pyn4 = 'N',"         'Y/N - 4.공구비 절감
    End If
    If chk_pyn5.Value = 1 Then
        sss = sss & "         tth_pyn5 = 'Y',"         'Y/N - 5.기타
    Else
        sss = sss & "         tth_pyn5 = 'N',"         'Y/N - 5.기타
    End If

    sss = sss & "         tth_sab = " & Val(txt_tsab1) & ","        '입력자
    sprd_rmk1.row = 1
    sprd_rmk1.Col = 1
    sss = sss & "         tth_rmk = '" & sprd_rmk1.Text & "',"      '비고
    sprd_rmk1.Col = 2
    sss = sss & "         tth_cmt = '" & sprd_rmk1.Text & "',"      '평가
    
    sss = sss & "         tth_updte = to_char(sysdate,'yyyymmdd')"  '수정일자
                          
    sss = sss & "   where tth_dat = " & txt_dat1
    sss = sss & "     and tth_seq = " & txt_seq1
    
    db.Execute sss, 64
    
    'DESC 삭제후 다시 등록함.
    sss = "       delete from man_tooltestds"
    sss = sss & "  where ttd_dat = " & txt_dat1
    sss = sss & "    and ttd_seq = " & txt_seq1
    
    db.Execute sss, 64
    
    '=======================
    'DESC입력
    '=======================
    '---------------------->
    '▼기존공구▼
    '---------------------->
    sss = "insert into man_tooltestds("
    sss = sss & "         ttd_dat,"     '등록일자
    sss = sss & "         ttd_seq,"     '등록순번
    sss = sss & "         ttd_lno,"     '등록행번
    sss = sss & "         ttd_gbn,"     '1:기존공구/2:테스트공구
    sss = sss & "         ttd_ryn,"     'Y/N - 결과선택
    
    sss = sss & "         ttd_maker,"   '제조사
    sss = sss & "         ttd_tipstd,"  '팁 규격/코드
    sss = sss & "         ttd_tipjil,"  '팁 재질
    sss = sss & "         ttd_holder,"  '적용홀더
    sss = sss & "         ttd_rcntmn,"  '분당회전 최소
    sss = sss & "         ttd_rcntmx,"  '분당회전 최대
    sss = sss & "         ttd_movmn,"   '이송(MM/REV) 최소
    sss = sss & "         ttd_movmx,"   '이송(MM/REV) 최대
    sss = sss & "         ttd_tct,"     'TCT
    sss = sss & "         ttd_pct,"     'PCT
                        
    sss = sss & "         ttd_depth,"   '절삭깊이
    
    sss = sss & "         ttd_fluid,"   '절삭유제
    sss = sss & "         ttd_qty,"     '가공수량
    sss = sss & "         ttd_dan,"     '단가/EA
    sss = sss & "         ttd_tldn,"    '공구비/EA
    sss = sss & "         ttd_chdn,"    '교환비/EA
    sss = sss & "         ttd_result,"  '결과 OK/NG
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
    
    If Val(txt_rcntmn1) > 0 Then    '분당회전수 최소값
        sss = sss & "        " & Val(txt_rcntmn1) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_rcntmx1) > 0 Then    '분당회전수 최대값
        sss = sss & "        " & Val(txt_rcntmx1) & ","
    Else
        sss = sss & "        null,"
    End If
                                                        
    If Val(txt_movmn1) > 0 Then    '이송 최소값
        sss = sss & "        " & Val(txt_movmn1) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_movmx1) > 0 Then    '이송 최대값
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
    
    If Val(txt_depth1) > 0 Then    '절삭깊이
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
    '▼TEST1▼
    '---------------------->
    sss = "insert into man_tooltestds("
    sss = sss & "         ttd_dat,"     '등록일자
    sss = sss & "         ttd_seq,"     '등록순번
    sss = sss & "         ttd_lno,"     '등록행번
    sss = sss & "         ttd_gbn,"     '1:기존공구/2:테스트공구
    sss = sss & "         ttd_ryn,"     'Y/N - 결과선택
    
    sss = sss & "         ttd_maker,"   '제조사
    sss = sss & "         ttd_tipstd,"  '팁 규격/코드
    sss = sss & "         ttd_tipjil,"  '팁 재질
    sss = sss & "         ttd_holder,"  '적용홀더
    sss = sss & "         ttd_rcntmn,"  '분당회전 최소
    sss = sss & "         ttd_rcntmx,"  '분당회전 최대
    sss = sss & "         ttd_movmn,"   '이송(MM/REV) 최소
    sss = sss & "         ttd_movmx,"   '이송(MM/REV) 최대
    sss = sss & "         ttd_tct,"     'TCT
    sss = sss & "         ttd_pct,"     'PCT
                        
    sss = sss & "         ttd_depth,"   '절삭깊이
    
    sss = sss & "         ttd_fluid,"   '절삭유제
    sss = sss & "         ttd_qty,"     '가공수량
    sss = sss & "         ttd_dan,"     '단가/EA
    sss = sss & "         ttd_tldn,"    '공구비/EA
    sss = sss & "         ttd_chdn,"    '교환비/EA
    sss = sss & "         ttd_result,"  '결과 OK/NG
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
    
    If Val(txt_rcntmn2) > 0 Then    '분당회전수 최소값
        sss = sss & "        " & Val(txt_rcntmn2) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_rcntmx2) > 0 Then    '분당회전수 최대값
        sss = sss & "        " & Val(txt_rcntmx2) & ","
    Else
        sss = sss & "        null,"
    End If
                                                        
    If Val(txt_movmn2) > 0 Then    '이송 최소값
        sss = sss & "        " & Val(txt_movmn2) & ","
    Else
        sss = sss & "        null,"
    End If
    If Val(txt_movmx2) > 0 Then    '이송 최대값
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
    
    If Val(txt_depth2) > 0 Then    '절삭깊이
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
    '▼TEST2▼
    '---------------------->
    
    If chk_test2.Value = 1 Then
    
    
        sss = "insert into man_tooltestds("
        sss = sss & "         ttd_dat,"     '등록일자
        sss = sss & "         ttd_seq,"     '등록순번
        sss = sss & "         ttd_lno,"     '등록행번
        sss = sss & "         ttd_gbn,"     '1:기존공구/2:테스트공구
        sss = sss & "         ttd_ryn,"     'Y/N - 결과선택
        
        sss = sss & "         ttd_maker,"   '제조사
        sss = sss & "         ttd_tipstd,"  '팁 규격/코드
        sss = sss & "         ttd_tipjil,"  '팁 재질
        sss = sss & "         ttd_holder,"  '적용홀더
        sss = sss & "         ttd_rcntmn,"  '분당회전 최소
        sss = sss & "         ttd_rcntmx,"  '분당회전 최대
        sss = sss & "         ttd_movmn,"   '이송(MM/REV) 최소
        sss = sss & "         ttd_movmx,"   '이송(MM/REV) 최대
        sss = sss & "         ttd_tct,"     'TCT
        sss = sss & "         ttd_pct,"     'PCT
                            
        sss = sss & "         ttd_depth,"   '절삭깊이
        
        sss = sss & "         ttd_fluid,"   '절삭유제
        sss = sss & "         ttd_qty,"     '가공수량
        sss = sss & "         ttd_dan,"     '단가/EA
        sss = sss & "         ttd_tldn,"    '공구비/EA
        sss = sss & "         ttd_chdn,"    '교환비/EA
        sss = sss & "         ttd_result,"  '결과 OK/NG
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
        
        If Val(txt_rcntmn3) > 0 Then    '분당회전수 최소값
            sss = sss & "        " & Val(txt_rcntmn3) & ","
        Else
            sss = sss & "        null,"
        End If
        If Val(txt_rcntmx3) > 0 Then    '분당회전수 최대값
            sss = sss & "        " & Val(txt_rcntmx3) & ","
        Else
            sss = sss & "        null,"
        End If
                                                            
        If Val(txt_movmn3) > 0 Then    '이송 최소값
            sss = sss & "        " & Val(txt_movmn3) & ","
        Else
            sss = sss & "        null,"
        End If
        If Val(txt_movmx3) > 0 Then    '이송 최대값
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
        
        If Val(txt_depth3) > 0 Then    '절삭깊이
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
    
    
    Call msg_display("등록내역이 수정되었습니다.")
    
    Exit Sub
    
err_rtn:

    Ws.Rollback
    MsgBox (Err.Description)
    
End Sub

'-------------------------------
'내역삭제
'-------------------------------
Private Sub btn_del1_Click()
    
On Error GoTo err_rtn
    
     If Job_Level < 9 Then
        MsgBox ("작업권한이 없습니다!")
        Exit Sub
    End If
    
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
    
    sss = "select * from man_tooltesthd, man_tooltestds"
    sss = sss & "  where tth_dat = " & txt_dat1
    sss = sss & "    and tth_seq = " & txt_seq1
    sss = sss & "    and tth_dat = ttd_dat"
    sss = sss & "    and tth_seq = ttd_seq"
    '
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Rs.RecordCount < 1 Then
        Rs.Close
        Call msg_display("내역이 없습니다. 삭제불가!")
        Exit Sub
    End If
    '
    '삭제
    If MsgBox("(" & txt_dat1 & "-" & Format(txt_seq1, "000") & ") 등록된 내역을 삭제하시겠습니까?", vbYesNo) <> vbYes Then
       Rs.Close
       msg_display ("취소되었습니다.")
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
    '첨부파일 전체 삭제
    If Not IsNull(Rs!tth_file1) Or Not IsNull(Rs!tth_file2) Or Not IsNull(Rs!tth_file3) Or Not IsNull(Rs!tth_file4) Then
       '
       '===========================
       ' FTP를 이용한 파일 삭제
       '===========================
       If FTP_Connection Then
          '
          If Not FTP경로체크(SERVER_PATH) Then
             Call FTP_DisConnect
             Ws.Rollback
             MsgBox "서버경로를 찾을수 없습니다.(정보관리센터 문의)"
             Exit Sub
          End If
          '
          If Not IsNull(Rs!tth_file1) Then
             If Not FTP_Delete(SERVER_PATH & Rs!tth_file1) Then
                Ws.Rollback
                Call FTP_DisConnect
                mkpoen05MDI.msg = "파일을 삭제할 수 없습니다. (정보관리센터 문의)"
                MsgBox "파일을 삭제할 수 없습니다. (정보관리센터 문의)"
                Exit Sub
             End If
          End If
          '
          If Not IsNull(Rs!tth_file2) Then
             If Not FTP_Delete(SERVER_PATH & Rs!tth_file2) Then
                Ws.Rollback
                Call FTP_DisConnect
                mkpoen05MDI.msg = "파일을 삭제할 수 없습니다. (정보관리센터 문의)"
                MsgBox "파일을 삭제할 수 없습니다. (정보관리센터 문의)"
                Exit Sub
             End If
          End If
          '
          If Not IsNull(Rs!tth_file3) Then
             If Not FTP_Delete(SERVER_PATH & Rs!tth_file3) Then
                Ws.Rollback
                Call FTP_DisConnect
                mkpoen05MDI.msg = "파일을 삭제할 수 없습니다. (정보관리센터 문의)"
                MsgBox "파일을 삭제할 수 없습니다. (정보관리센터 문의)"
                Exit Sub
             End If
          End If
          '
          If Not IsNull(Rs!tth_file4) Then
             If Not FTP_Delete(SERVER_PATH & Rs!tth_file4) Then
                Ws.Rollback
                Call FTP_DisConnect
                mkpoen05MDI.msg = "파일을 삭제할 수 없습니다. (정보관리센터 문의)"
                MsgBox "파일을 삭제할 수 없습니다. (정보관리센터 문의)"
                Exit Sub
             End If
          End If
          '
          Call FTP_DisConnect
          '
       End If
       '
       '======================================
       ' 네트워크 드라이버를 이용한 파일 삭제
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

    '초기화
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
    Call msg_display("내역이 삭제되었습니다!")
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
'초기화
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
    
    '기존공구내역
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
    
    '테스트1 내역
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
                
    '테스트2 내역
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
    
    '결과
    chk_ryn1.Value = False
    chk_ryn2.Value = False
    chk_ryn3.Value = False
    chk_ryn4.Value = False
    chk_ryn5.Value = False
    
    cmb_result1.ListIndex = 0
    cmb_rtyn1.ListIndex = 0
    
    '비고/평가
    sprd_rmk1.row = 1
    sprd_rmk1.Col = 1: sprd_rmk1.Text = "LOT.NO : " & vbCrLf & "가공수량 : " & vbCrLf & "테스트수량 : "
    sprd_rmk1.Col = 2: sprd_rmk1.Text = ""
    
    For ii = 1 To 5
        spd_file11.row = ii: spd_file11.Col = 2: spd_file11.Text = ""
    Next ii
    
    For ii = 1 To 5
        spd_file12.row = ii: spd_file12.Col = 1: spd_file12.Text = ""
    Next ii
    
    spd_file11.Visible = True               '파일(신규용)
    spd_file12.Visible = False              '파일(수정용)
    lbl_cmt1.Visible = False
    
    btn_prt1.Enabled = False
    
End Sub

'테스트2활성화
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
'장비코드 조회
'------------------------->
Private Sub txt_mcd1_LostFocus()
'장비 조회

    txt_mcd1 = UCase(Trim(txt_mcd1))
    
    If Len(Trim(txt_mcd1)) > 1 Then
        sss = "       select mhc_name, ems_mark, sok_name(mhc_sok) soknm "
        sss = sss & "   from man_machcd, eam_mast"
        sss = sss & "  where mhc_code = '" & txt_mcd1 & "'"
        sss = sss & "    and mhc_code = ems_mcd(+)"
            
        Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        If Ks.RecordCount < 1 Then
            txt_mnm1 = "": txt_mmake1 = "": txt_msok1 = ""
            Ks.Close: Call msg_display("장비코드를 확인하세요!")
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
'LOT NO. 조회
'------------------------->
Private Sub txt_lotno1_LostFocus()
'장비 조회

    txt_lotno1 = UCase(Trim(txt_lotno1))
    
    If Len(Trim(txt_lotno1)) = 8 Then
        sss = "       select dit_bpcd, dit_bpjil, dit_jacd, dit_jajil"
        sss = sss & "   from man_direct"
        sss = sss & "  where dit_lot = '" & txt_lotno1 & "'"
            
        Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
        If Ks.RecordCount < 1 Then
            txt_bpcd1 = "": txt_bpjil1 = ""
            txt_jacd1 = "": txt_jajil1 = ""
            Ks.Close: Call msg_display("LOT NO.를 확인하세요!")
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
'사번 조회
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
            Ks.Close: Call msg_display("작업자 사번을 확인하세요!")
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
            Ks.Close: Call msg_display("작업자 이름을 확인하세요!")
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
'        MsgBox ("작업권한이 없습니다!")
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
       chk_ryn1.Caption = "1.공구수명 연장"
       chk_ryn2.Caption = "2.칩처리 양호"
       chk_ryn3.Caption = "3.공구비 절감"
       chk_ryn4.Caption = "4.시간 단축"
       chk_ryn5.Caption = "5.기타"
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
       chk_ryn1.Caption = "1.결손"
       chk_ryn2.Caption = "2.마모"
       chk_ryn3.Caption = "3.칩처리 불량"
       chk_ryn4.Caption = "4.공구비 상승"
       chk_ryn5.Caption = "5.기타"
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


'첨부파일 조회
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
    
    '첨부파일1
    spd_file12.row = 1
    If Not IsNull(Ks!tth_file1) Then
        spd_file12.Col = 1: spd_file12.Text = Ks!tth_file1
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "삭제"
    Else
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "등록"
    End If
    
    '첨부파일2
    spd_file12.row = 2
    If Not IsNull(Ks!tth_file2) Then
        spd_file12.Col = 1: spd_file12.Text = Ks!tth_file2
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "삭제"
    Else
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "등록"
    End If
    
    '첨부파일2
    spd_file12.row = 3
    If Not IsNull(Ks!tth_file3) Then
        spd_file12.Col = 1: spd_file12.Text = Ks!tth_file3
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "삭제"
    Else
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "등록"
    End If
    
    '첨부파일2
    spd_file12.row = 4
    If Not IsNull(Ks!tth_file4) Then
        spd_file12.Col = 1: spd_file12.Text = Ks!tth_file4
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "삭제"
    Else
        spd_file12.Col = 3: spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "등록"
    End If
    
    Ks.Close
    
End Sub

'----------------------
'첨부파일 추가 및 삭제 - 신규등록
'----------------------
Private Sub spd_file11_ButtonClicked(ByVal Col As Long, ByVal row As Long, ByVal ButtonDown As Integer)
    '
    On Error GoTo err_rtn
    '
    If Job_Level < 9 Then
        MsgBox ("작업권한이 없습니다!")
        Exit Sub
    End If
    '
    If Col = 1 Then
        
        spd_file11.row = row: spd_file11.Col = 2
                    
        spd_file11.Text = ""
        
        Comm1.CancelError = True
        Comm1.Flags = cdlOFNOverwritePrompt
        Comm1.DialogTitle = "보증보험 신청 첨부파일"
        Comm1.InitDir = "C:\"
        Comm1.Filter = "PDF파일 or JPG파일 (*.pdf;*.jpg)|*.pdf;*.jpg"
        Comm1.ShowOpen
        
        If UCase(Right(Comm1.fileName, 3)) <> "PDF" And UCase(Right(Comm1.fileName, 3)) <> "JPG" Then
            MsgBox ("선택한 파일 확장자 'PDF' or 'JPG' 가 아닙니다.")
            Exit Sub
        End If
        '
        spd_file11.Text = UCase(Comm1.fileName)
        '
        Call msg_display("첨부파일이 추가되었습니다!")
        Exit Sub
    
    End If
    
    If Col = 4 Then
        
        spd_file11.row = row: spd_file11.Col = 2
        If Len(spd_file11.Text) < 1 Then Exit Sub
        spd_file11.Text = ""
        Call msg_display("첨부파일이 삭제되었습니다!")
        Exit Sub

    End If
    '
    '
    Exit Sub

err_rtn:
  mkpoen05MDI.msg = Err.Description

End Sub

'----------------------
'첨부파일 추가 및 삭제 - 수정등록
'----------------------
Private Sub spd_file12_ButtonClicked(ByVal Col As Long, ByVal row As Long, ByVal ButtonDown As Integer)
    '
    On Error GoTo err_rtn
    '
    If Job_Level < 9 Then
        MsgBox ("작업권한이 없습니다!")
        Exit Sub
    End If

    Dim isrno As String     '보증보험 번호
    Dim filenm1 As String   '파일1
    Dim filenm2 As String   '파일2
    Dim filenm3 As String   '파일3
    Dim filenm4 As String   '파일4
    
    Dim fs            As Object 'Scripting.FileSystemObject 객체
    Dim lsSouurce     As String '복사대상 경로
    Dim lsDestination As String '복사위치 경로
    Dim F_NAME     As String
    
    spd_file12.row = row: spd_file12.Col = Col
    
    'If Len(Gsab) < 1 Then
    '    txt_pwd.SetFocus
    '    Call msg_display("로그인 후 작업하세요!!")
    '    Exit Sub
    'End If
    
    'If Gsok <> "D100" And Gsok <> "D000" And Gsab <> "1476" Then
    '    Call msg_display("구매부서만 등록/수정/삭제 작업 가능합니다. 소속을 확인하세요!")
    '    Exit Sub
    'End If
    
    '첨부파일 등록
    If spd_file12.TypeButtonText = "등록" And Col = 3 Then
            
        sss = "       select tth_file1, tth_file2, tth_file3, tth_file4"
        sss = sss & "   from man_tooltesthd "
        sss = sss & "  where tth_dat = " & txt_dat1
        sss = sss & "    and tth_seq = " & txt_seq1
    
        Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    
        If Rs.RecordCount < 1 Then
            Rs.Close
            msg_display ("첨부등록할 등록내역이 없습니다. 등록번호를 확인하세요!")
            Exit Sub
        End If
            
        'chk
        Rs.Close
        '
        Comm1.CancelError = True
        Comm1.Flags = cdlOFNOverwritePrompt
        Comm1.DialogTitle = "첨부파일"
        Comm1.InitDir = "C:\"
        Comm1.Filter = "PDF파일 or JPG파일 (*.pdf;*.jpg)|*.pdf;*.jpg"
            Comm1.ShowOpen
        If UCase(Right(Comm1.fileName, 3)) <> "PDF" And UCase(Right(Comm1.fileName, 3)) <> "JPG" Then
            MsgBox ("선택한 파일 확장자 'PDF' 또는 'JPG' 파일이 아닙니다!")
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
       ' FTP를 이용한 파일 업로드
       '===========================
       If FTP_Connection Then
          '
          If Not FTP경로체크(SERVER_PATH) Then
             Call FTP_DisConnect
             MsgBox "서버경로를 찾을수 없습니다.(정보관리센터 문의)"
             Exit Sub
          End If
          '
          File_Path = ""
          If row = 1 Then File_Path = filenm1
          If row = 2 Then File_Path = filenm2
          If row = 3 Then File_Path = filenm3
          If row = 4 Then File_Path = filenm4
          '
          '업로드
          If Not FTP_Upload(Comm1.fileName, SERVER_PATH, File_Path) Then
             Call FTP_DisConnect
             mkpoen05MDI.msg = "파일 업로드에 실패 하였습니다. (정보관리센터 문의)"
             Exit Sub
          End If
          '
          Call FTP_DisConnect
          '
       End If
       '
       '========================================
       ' 네트워크 드라이버를 이용한 파일 업로드
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
        spd_file12.Col = 3:  spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "삭제"
        '
        msg_display ("첨부파일이 등록되었습니다!")
        '
        Exit Sub
    End If
    
    '첨부파일 삭제
    If spd_file12.TypeButtonText = "삭제" And Col = 3 Then
       '
       sss = "select tth_file1, tth_file2, tth_file3, tth_file4"
       sss = sss & " from man_tooltesthd "
       sss = sss & " where tth_dat = " & txt_dat1
       sss = sss & "   and tth_seq = " & txt_seq1
       '
       Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
       If Rs.RecordCount < 1 Then
          Rs.Close
          msg_display ("첨부삭제할 등록번호가 없습니다. 등록번호를 확인하세요!")
           Exit Sub
       End If
       '
       If MsgBox("첨부파일을 삭제하시겠습니까?", vbYesNo) <> vbYes Then
           Rs.Close
           msg_display ("첨부파일 삭제작업이 취소되었습니다!")
           Exit Sub
       End If
       '
       '파일 삭제등록
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
       ' FTP를 이용한 파일 삭제
       '===========================
       If FTP_Connection Then
          '
          If Not FTP경로체크(SERVER_PATH) Then
             Call FTP_DisConnect
             MsgBox "서버경로를 찾을수 없습니다.(정보관리센터 문의)"
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
             mkpoen05MDI.msg = "파일을 삭제할 수 없습니다. (정보관리센터 문의)"
             MsgBox "파일을 삭제할 수 없습니다. (정보관리센터 문의)"
             Exit Sub
          End If
          '
          Call FTP_DisConnect
          '
       End If
       '
       '======================================
       ' 네트워크 드라이버를 이용한 파일 삭제
       '======================================
      ' If Row = 1 Then Kill (N_Driver & ":\" & Rs!tth_file1)
      ' If Row = 2 Then Kill (N_Driver & ":\" & Rs!tth_file2)
      ' If Row = 3 Then Kill (N_Driver & ":\" & Rs!tth_file3)
      ' If Row = 4 Then Kill (N_Driver & ":\" & Rs!tth_file4)
       '
       spd_file12.Col = 1: spd_file12.Text = ""
       spd_file12.Col = 3:  spd_file12.CellType = CellTypeButton: spd_file12.TypeButtonText = "등록"
       Rs.Close
       '
       msg_display ("첨부파일이 삭제되었습니다!")
       '
       Exit Sub
       '
    End If
    '
    '참부파일 저장
    If spd_file12.TypeButtonText = "저장" And Col = 4 Then
       '
       spd_file12.row = row: spd_file12.Col = 1
       '
       '파일이름 유무
       If Len(spd_file12.Text) < 5 Then
           msg_display ("등록된 첨부파일이 없습니다!")
           Exit Sub
       End If
       '
       '파일존재 유무
      ' exists = ExistFile(N_Driver & ":\" & spd_file12.Text)
      ' If exists = False Then
      '    msg_display ("저장할 첨부파일이 없습니다!")
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
          If Not FTP_Download(SERVER_PATH & spd_file12.Text, lsDestination) Then
             '
          End If
          '
          Call FTP_DisConnect
          '
       End If
       '
       '===========================================
       ' 네트워크 드라이버를 이용한 파일 다운로드
       '===========================================
      ' Set fs = CreateObject("Scripting.FileSystemObject")
      ' fs.CopyFile lsSouurce, lsDestination '복사
       '
       msg_display ("첨부파일이 저장되었습니다!")
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

'첨부파일 보기
Private Sub spd_file12_DblClick(ByVal Col As Long, ByVal row As Long)
    
On Error GoTo err_rtn
        
    Dim isrno As String     '보증보험 번호
    Dim filenm1 As String   '파일1
    Dim filenm2 As String   '파일2
    
    Dim fs            As Object 'Scripting.FileSystemObject 객체
    Dim lsSouurce     As String '복사대상 경로
    Dim lsDestination As String '복사위치 경로
    Dim F_NAME     As String
    
    spd_file12.row = row: spd_file12.Col = 2
    
    '
    If Len(spd_file12.Text) > 2 Then
        
        spd_file12.row = row: spd_file12.Col = 1
        
        If Len(spd_file12.Text) < 5 Then
            msg_display ("등록된 첨부파일이 없습니다!")
            Exit Sub
        End If
        
        If Install_ACROBET = False Then
            MsgBox "아크로벳리더가 설치되어 있지 않아 보기가 불가능합니다."
            Exit Sub
        End If
        '
        mkpoen05_view.Show
        'mkpoen05_view.txt_dat = txt_dat2
        'mkpoen05_view.txt_seq = Format(txt_seq2, "000")
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
       ' 네트워크 드라이버를 이용한 파일 View
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
'첨부파일 체크
'----------------------
Public Function Check_file(i As Integer, j As Integer)
    
    spd_file11.row = i: spd_file11.Col = j
    
    If Len(spd_file11.Text) > 0 Then
            
        If Right(spd_file11.Text, 3) <> "PDF" And Right(spd_file11.Text, 3) <> "pdf" And Right(spd_file11.Text, 3) <> "JPG" And Right(spd_file11.Text, 3) <> "jpg" Then
                msg_display (i & "번째행: 첨부파일은 PDF 또는 JPG만 등록 가능합니다!")
                Check_file = "X"
                Exit Function
        End If
        '등록할 파일이 있는지 chk
        exists = ExistFile(spd_file11.Text)
        If exists = False Then
            msg_display (i & "번째행: 경로에 파일이 존재하지 않습니다. 경로를 다시 확인하세요!")
            Check_file = "X"
            Exit Function
        End If
        
        Check_file = "Y"
        Exit Function
        
    End If
    
    Check_file = "N"

End Function

'파일존재유무 체크
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
'TAB2. TEST DATA 조회
'===================================
Private Sub btn_view2_Click()
   
On Error GoTo err_rtn

    Dim chk_ttdno As String
    
    If Len(txt_sdat2) <> 8 Or IsNumeric(txt_sdat2) = False Then
        Call msg_display("등록 시작일자를 확인하세요")
        txt_sdat2.SetFocus
        Exit Sub
    End If
    
    If Len(txt_edat2) <> 8 Or IsNumeric(txt_edat2) = False Then
        Call msg_display("등록 종료일자를 확인하세요")
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
        Call msg_display("등록된 내역이 없습니다!")
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
             sprd2.AddCellSpan 1, cnt - 1, 15, 1      '품명 병합
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
       If Rs!ttd_gbn = 1 Then sprd2.Col = 8: sprd2.Text = "기존"
       If Rs!ttd_gbn = 2 Then sprd2.Col = 8: sprd2.Text = "테스트"
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
    Call msg_display("산출 내역을 확인하세요")
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
  
       pausetime = 0.02    ' 기간을 지정합니다.
       start = Timer       ' 시작 시간을 지정합니다.
       Do While Timer < start + pausetime
          DoEvents         ' 다른 프로시저로 넘깁니다.
       Loop
   Next

End Sub


'---------------------------
' 아크로벳
'---------------------------
'ACROBET 설치 여부 확인
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
    ' 현재 디렉토리와 포함하는 디렉토리를 무시합니다.
    If MyName <> "." And MyName <> ".." Then
      ' MyName이 디렉토리인지 확인하기 위해서 비트별(bitwise) 비교를 사용합니다.
      If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
        If Left(UCase(Trim(MyName)), 5) = "ADOBE" Then
          '
          MyPath = "C:\Program Files\Adobe\"
          MyName = Dir(MyPath, vbDirectory)
          
          '폴더만 보기
          Do While MyName <> ""
            ' 현재 디렉토리와 포함하는 디렉토리를 무시합니다.
            If MyName <> "." And MyName <> ".." Then
              ' MyName이 디렉토리인지 확인하기 위해서 비트별(bitwise) 비교를 사용합니다.
              If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
                If Left(UCase(Trim(MyName)), 7) = "ACROBAT" Then
                  ii = 1
                  Exit Do
                'Reader 8.0 일때 : C:\Program Files\Adobe\Reader 8.0\Reader\AdobeUpdateCheck.exe
                ElseIf Left(UCase(Trim(MyName)), 6) = "READER" Then
                  ii = 1
                  Exit Do
                End If
              End If
            End If
            '
            MyName = Dir  ' 다음 항목을 읽어들입니다.
            '
          Loop
          '
          If ii = 1 Then Exit Do
          '
        End If
      End If
    End If
    '
    MyName = Dir  ' 다음 항목을 읽어들입니다.
    '
  Loop
  
  '64비트 용
  If ii = 0 Then
  
    MyPath = "C:\Program Files (x86)\"
    ii = 0
  '
    MyName = Dir(MyPath, vbDirectory)
    Do While MyName <> ""
      ' 현재 디렉토리와 포함하는 디렉토리를 무시합니다.
      If MyName <> "." And MyName <> ".." Then
        ' MyName이 디렉토리인지 확인하기 위해서 비트별(bitwise) 비교를 사용합니다.
        If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
          If Left(UCase(Trim(MyName)), 5) = "ADOBE" Then
            '
            MyPath = "C:\Program Files (x86)\Adobe\"
            MyName = Dir(MyPath, vbDirectory)
            
            '폴더만 보기
            Do While MyName <> ""
              ' 현재 디렉토리와 포함하는 디렉토리를 무시합니다.
              If MyName <> "." And MyName <> ".." Then
                ' MyName이 디렉토리인지 확인하기 위해서 비트별(bitwise) 비교를 사용합니다.
                If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
                  If Left(UCase(Trim(MyName)), 7) = "ACROBAT" Then
                    ii = 1
                    Exit Do
                  'Reader 8.0 일때 : C:\Program Files\Adobe\Reader 8.0\Reader\AdobeUpdateCheck.exe
                  ElseIf Left(UCase(Trim(MyName)), 6) = "READER" Then
                    ii = 1
                    Exit Do
                  End If
                End If
              End If
              '
              MyName = Dir  ' 다음 항목을 읽어들입니다.
              '
            Loop
            '
            If ii = 1 Then Exit Do
            '
          End If
        End If
      End If
      '
      MyName = Dir  ' 다음 항목을 읽어들입니다.
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

