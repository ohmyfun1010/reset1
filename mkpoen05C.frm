VERSION 5.00
Object = "{3B930683-5AF1-4F07-9CE8-CA8063E1F3DD}#2.0#0"; "XButton.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form mkpoen05C 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   9840
   ClientLeft      =   7215
   ClientTop       =   3210
   ClientWidth     =   14040
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   14040
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9855
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      _Version        =   65536
      _ExtentX        =   24791
      _ExtentY        =   17383
      _StockProps     =   15
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   9855
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   17383
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "공구 수명 테스트 관리"
         TabPicture(0)   =   "mkpoen05C.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SSFrame1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "공구 수명 테스트 조회"
         TabPicture(1)   =   "mkpoen05C.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         Begin Threed.SSFrame SSFrame1 
            Height          =   9375
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   13815
            _Version        =   65536
            _ExtentX        =   24368
            _ExtentY        =   16536
            _StockProps     =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin FPSpreadADO.fpSpread spd_app 
               Height          =   930
               Left            =   9720
               TabIndex        =   110
               Top             =   165
               Width           =   3930
               _Version        =   458752
               _ExtentX        =   6932
               _ExtentY        =   1640
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
               MaxCols         =   4
               MaxRows         =   2
               RetainSelBlock  =   0   'False
               ScrollBars      =   0
               SpreadDesigner  =   "mkpoen05C.frx":0038
            End
            Begin VB.TextBox txt_seq 
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
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
               Left            =   2715
               MaxLength       =   3
               TabIndex        =   56
               Top             =   500
               Width           =   510
            End
            Begin VB.TextBox txt_dat 
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
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
               Left            =   1470
               MaxLength       =   8
               TabIndex        =   55
               Top             =   500
               Width           =   1185
            End
            Begin Threed.SSFrame SSFrame5 
               Height          =   1455
               Left            =   120
               TabIndex        =   49
               Top             =   1080
               Width           =   13575
               _Version        =   65536
               _ExtentX        =   23945
               _ExtentY        =   2566
               _StockProps     =   14
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.TextBox txt_rmk 
                  Appearance      =   0  '평면
                  Height          =   690
                  Left            =   9240
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  '수직
                  TabIndex        =   6
                  Top             =   600
                  Width           =   4215
               End
               Begin VB.TextBox txt_lot 
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
                  Left            =   6120
                  MaxLength       =   20
                  TabIndex        =   4
                  Top             =   240
                  Width           =   1755
               End
               Begin VB.TextBox txt_gocd 
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
                  Left            =   9240
                  MaxLength       =   20
                  TabIndex        =   5
                  Top             =   240
                  Width           =   1755
               End
               Begin VB.TextBox txt_bpcd 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Left            =   6120
                  MaxLength       =   20
                  TabIndex        =   101
                  Top             =   960
                  Width           =   1755
               End
               Begin VB.TextBox txt_jacd 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Left            =   6120
                  MaxLength       =   20
                  TabIndex        =   100
                  Top             =   600
                  Width           =   1755
               End
               Begin VB.TextBox txt_sab 
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
                  Left            =   1350
                  MaxLength       =   20
                  TabIndex        =   1
                  Top             =   240
                  Width           =   825
               End
               Begin VB.TextBox txt_name 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Left            =   2235
                  MaxLength       =   7
                  TabIndex        =   51
                  Top             =   240
                  Width           =   1065
               End
               Begin VB.TextBox txt_mcdname 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Left            =   2235
                  MaxLength       =   20
                  TabIndex        =   50
                  TabStop         =   0   'False
                  Top             =   600
                  Width           =   2460
               End
               Begin VB.TextBox txt_mcd 
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
                  Left            =   1350
                  MaxLength       =   5
                  TabIndex        =   2
                  Top             =   600
                  Width           =   825
               End
               Begin VB.TextBox txt_title 
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
                  Left            =   1350
                  MaxLength       =   20
                  TabIndex        =   3
                  Top             =   960
                  Width           =   3390
               End
               Begin Threed.SSPanel SSPanel37 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   52
                  Top             =   240
                  Width           =   1200
                  _Version        =   65536
                  _ExtentX        =   2117
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
                  Index           =   1
                  Left            =   120
                  TabIndex        =   53
                  Top             =   600
                  Width           =   1200
                  _Version        =   65536
                  _ExtentX        =   2117
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
               Begin Threed.SSPanel SSPanel4 
                  Height          =   330
                  Index           =   3
                  Left            =   120
                  TabIndex        =   54
                  Top             =   960
                  Width           =   1200
                  _Version        =   65536
                  _ExtentX        =   2117
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "테스트명"
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
                  Left            =   4920
                  TabIndex        =   102
                  Top             =   240
                  Width           =   1170
                  _Version        =   65536
                  _ExtentX        =   2064
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
               Begin Threed.SSPanel SSPanel38 
                  Height          =   330
                  Left            =   8040
                  TabIndex        =   103
                  Top             =   240
                  Width           =   1170
                  _Version        =   65536
                  _ExtentX        =   2064
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "공정코드"
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
               Begin Threed.SSPanel SSPanel39 
                  Height          =   330
                  Left            =   4920
                  TabIndex        =   104
                  Top             =   960
                  Width           =   1170
                  _Version        =   65536
                  _ExtentX        =   2064
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "제품명"
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
               Begin Threed.SSPanel SSPanel40 
                  Height          =   330
                  Left            =   4920
                  TabIndex        =   105
                  Top             =   600
                  Width           =   1170
                  _Version        =   65536
                  _ExtentX        =   2064
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "소재규격"
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
               Begin Threed.SSPanel SSPanel41 
                  Height          =   690
                  Left            =   8040
                  TabIndex        =   106
                  Top             =   600
                  Width           =   1170
                  _Version        =   65536
                  _ExtentX        =   2064
                  _ExtentY        =   1217
                  _StockProps     =   15
                  Caption         =   "담당자 의견"
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
            Begin Threed.SSFrame SSFrame2 
               Height          =   5895
               Left            =   120
               TabIndex        =   57
               Top             =   2760
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   10398
               _StockProps     =   14
               Caption         =   "기존"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.ComboBox cmb_fluid 
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
                  Index           =   1
                  ItemData        =   "mkpoen05C.frx":03B5
                  Left            =   2280
                  List            =   "mkpoen05C.frx":03C2
                  TabIndex        =   18
                  Top             =   3120
                  Width           =   2010
               End
               Begin VB.TextBox txt_movmx 
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
                  Index           =   1
                  Left            =   3360
                  MaxLength       =   20
                  TabIndex        =   15
                  Top             =   2400
                  Width           =   915
               End
               Begin VB.TextBox txt_movmn 
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
                  Index           =   1
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   14
                  Top             =   2400
                  Width           =   915
               End
               Begin VB.TextBox txt_depth 
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
                  Index           =   1
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   13
                  Top             =   2040
                  Width           =   1995
               End
               Begin VB.TextBox txt_rcntmx 
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
                  Index           =   1
                  Left            =   3360
                  MaxLength       =   20
                  TabIndex        =   12
                  Top             =   1680
                  Width           =   915
               End
               Begin VB.TextBox txt_rcntmn 
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
                  Index           =   1
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   11
                  Top             =   1680
                  Width           =   915
               End
               Begin VB.TextBox txt_holder 
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
                  Index           =   1
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   10
                  Top             =   1320
                  Width           =   1995
               End
               Begin VB.TextBox txt_tipjil 
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
                  Index           =   1
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   9
                  Top             =   960
                  Width           =   1995
               End
               Begin VB.TextBox txt_tct 
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
                  Index           =   1
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   16
                  Top             =   2760
                  Width           =   915
               End
               Begin VB.TextBox txt_tipstd 
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
                  Index           =   1
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   8
                  Top             =   600
                  Width           =   1995
               End
               Begin VB.TextBox txt_pct 
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
                  Index           =   1
                  Left            =   3360
                  MaxLength       =   20
                  TabIndex        =   17
                  Top             =   2760
                  Width           =   915
               End
               Begin VB.TextBox txt_maker 
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
                  Index           =   1
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   7
                  Top             =   240
                  Width           =   1995
               End
               Begin FPSpreadADO.fpSpread spd1 
                  Height          =   2220
                  Left            =   120
                  TabIndex        =   19
                  Top             =   3480
                  Width           =   4170
                  _Version        =   458752
                  _ExtentX        =   7355
                  _ExtentY        =   3916
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
                  MaxCols         =   4
                  MaxRows         =   6
                  RetainSelBlock  =   0   'False
                  ScrollBars      =   0
                  SpreadDesigner  =   "mkpoen05C.frx":03DE
               End
               Begin Threed.SSPanel SSPanel13 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   58
                  Top             =   240
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
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
               Begin Threed.SSPanel SSPanel14 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   59
                  Top             =   600
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
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
               Begin Threed.SSPanel SSPanel15 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   60
                  Top             =   960
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
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
               Begin Threed.SSPanel SSPanel17 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   61
                  Top             =   1320
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "기존 H/D 규격(HOLDER)"
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
                  Left            =   120
                  TabIndex        =   62
                  Top             =   1680
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "RPM/V(분당회전수)"
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
               Begin Threed.SSPanel SSPanel19 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   63
                  Top             =   2040
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "DEPTH(절입량)"
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
               Begin Threed.SSPanel SSPanel20 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   64
                  Top             =   2400
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "FEEDRATE(이송)"
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
               Begin Threed.SSPanel SSPanel2 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   65
                  Top             =   2760
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "전체C/T,공정C/T"
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
                  Left            =   120
                  TabIndex        =   66
                  Top             =   3120
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "절삭유"
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
               Begin VB.Label Label3 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   3225
                  TabIndex        =   69
                  Top             =   1770
                  Width           =   105
               End
               Begin VB.Label Label1 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   3225
                  TabIndex        =   68
                  Top             =   2835
                  Width           =   105
               End
               Begin VB.Label Label2 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   3225
                  TabIndex        =   67
                  Top             =   2475
                  Width           =   105
               End
            End
            Begin Threed.SSFrame SSFrame3 
               Height          =   5895
               Left            =   4680
               TabIndex        =   70
               Top             =   2760
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   10398
               _StockProps     =   14
               Caption         =   "테스트1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.ComboBox cmb_fluid 
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
                  Index           =   2
                  ItemData        =   "mkpoen05C.frx":091A
                  Left            =   2280
                  List            =   "mkpoen05C.frx":0927
                  TabIndex        =   31
                  Top             =   3120
                  Width           =   2010
               End
               Begin VB.TextBox txt_maker 
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
                  Index           =   2
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   20
                  Top             =   240
                  Width           =   1995
               End
               Begin VB.TextBox txt_pct 
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
                  Index           =   2
                  Left            =   3360
                  MaxLength       =   20
                  TabIndex        =   30
                  Top             =   2760
                  Width           =   915
               End
               Begin VB.TextBox txt_tipstd 
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
                  Index           =   2
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   21
                  Top             =   600
                  Width           =   1995
               End
               Begin VB.TextBox txt_tct 
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
                  Index           =   2
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   29
                  Top             =   2760
                  Width           =   915
               End
               Begin VB.TextBox txt_tipjil 
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
                  Index           =   2
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   22
                  Top             =   960
                  Width           =   1995
               End
               Begin VB.TextBox txt_holder 
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
                  Index           =   2
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   23
                  Top             =   1320
                  Width           =   1995
               End
               Begin VB.TextBox txt_rcntmn 
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
                  Index           =   2
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   24
                  Top             =   1680
                  Width           =   915
               End
               Begin VB.TextBox txt_rcntmx 
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
                  Index           =   2
                  Left            =   3360
                  MaxLength       =   20
                  TabIndex        =   25
                  Top             =   1680
                  Width           =   915
               End
               Begin VB.TextBox txt_depth 
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
                  Index           =   2
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   26
                  Top             =   2040
                  Width           =   1995
               End
               Begin VB.TextBox txt_movmn 
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
                  Index           =   2
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   27
                  Top             =   2400
                  Width           =   915
               End
               Begin VB.TextBox txt_movmx 
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
                  Index           =   2
                  Left            =   3360
                  MaxLength       =   20
                  TabIndex        =   28
                  Top             =   2400
                  Width           =   915
               End
               Begin Threed.SSPanel SSPanel11 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   71
                  Top             =   240
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
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
               Begin Threed.SSPanel SSPanel12 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   72
                  Top             =   600
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
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
               Begin Threed.SSPanel SSPanel21 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   73
                  Top             =   960
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
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
               Begin Threed.SSPanel SSPanel22 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   74
                  Top             =   1320
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "기존 H/D 규격(HOLDER)"
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
                  Left            =   120
                  TabIndex        =   75
                  Top             =   1680
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "RPM/V(분당회전수)"
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
                  Left            =   120
                  TabIndex        =   76
                  Top             =   2040
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "DEPTH(절입량)"
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
                  Left            =   120
                  TabIndex        =   77
                  Top             =   2400
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "FEEDRATE(이송)"
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
                  Left            =   120
                  TabIndex        =   78
                  Top             =   2760
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "전체C/T,공정C/T"
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
                  Left            =   120
                  TabIndex        =   79
                  Top             =   3120
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "절삭유"
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
               Begin FPSpreadADO.fpSpread spd2 
                  Height          =   2220
                  Left            =   120
                  TabIndex        =   107
                  Top             =   3480
                  Width           =   4170
                  _Version        =   458752
                  _ExtentX        =   7355
                  _ExtentY        =   3916
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
                  MaxCols         =   4
                  MaxRows         =   6
                  RetainSelBlock  =   0   'False
                  ScrollBars      =   0
                  SpreadDesigner  =   "mkpoen05C.frx":0943
               End
               Begin VB.Label Label4 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   3225
                  TabIndex        =   82
                  Top             =   2475
                  Width           =   105
               End
               Begin VB.Label Label5 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   3225
                  TabIndex        =   81
                  Top             =   2835
                  Width           =   105
               End
               Begin VB.Label Label6 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   3225
                  TabIndex        =   80
                  Top             =   1770
                  Width           =   105
               End
            End
            Begin Threed.SSFrame SSFrame4 
               Height          =   6135
               Left            =   9240
               TabIndex        =   83
               Top             =   2520
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   10821
               _StockProps     =   14
               Caption         =   "테스트2"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.CheckBox chk_use 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "사용"
                  Height          =   180
                  Left            =   3600
                  TabIndex        =   32
                  Top             =   200
                  Width           =   735
               End
               Begin VB.ComboBox cmb_fluid 
                  BackColor       =   &H00E0E0E0&
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
                  Height          =   315
                  Index           =   3
                  ItemData        =   "mkpoen05C.frx":0E7F
                  Left            =   2280
                  List            =   "mkpoen05C.frx":0E8C
                  TabIndex        =   44
                  Top             =   3360
                  Width           =   2010
               End
               Begin VB.TextBox txt_movmx 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Index           =   3
                  Left            =   3360
                  MaxLength       =   20
                  TabIndex        =   41
                  Top             =   2640
                  Width           =   915
               End
               Begin VB.TextBox txt_movmn 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Index           =   3
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   40
                  Top             =   2640
                  Width           =   915
               End
               Begin VB.TextBox txt_depth 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Index           =   3
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   84
                  Top             =   2280
                  Width           =   1995
               End
               Begin VB.TextBox txt_rcntmx 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Index           =   3
                  Left            =   3360
                  MaxLength       =   20
                  TabIndex        =   38
                  Top             =   1920
                  Width           =   915
               End
               Begin VB.TextBox txt_rcntmn 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Index           =   3
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   37
                  Top             =   1920
                  Width           =   915
               End
               Begin VB.TextBox txt_holder 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Index           =   3
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   36
                  Top             =   1560
                  Width           =   1995
               End
               Begin VB.TextBox txt_tipjil 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Index           =   3
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   35
                  Top             =   1200
                  Width           =   1995
               End
               Begin VB.TextBox txt_tct 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Index           =   3
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   42
                  Top             =   3000
                  Width           =   915
               End
               Begin VB.TextBox txt_tipstd 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Index           =   3
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   34
                  Top             =   840
                  Width           =   1995
               End
               Begin VB.TextBox txt_pct 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Index           =   3
                  Left            =   3360
                  MaxLength       =   20
                  TabIndex        =   43
                  Top             =   3000
                  Width           =   915
               End
               Begin VB.TextBox txt_maker 
                  Appearance      =   0  '평면
                  BackColor       =   &H00E0E0E0&
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
                  Index           =   3
                  Left            =   2280
                  MaxLength       =   20
                  TabIndex        =   33
                  Top             =   480
                  Width           =   1995
               End
               Begin Threed.SSPanel SSPanel28 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   85
                  Top             =   480
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
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
               Begin Threed.SSPanel SSPanel29 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   86
                  Top             =   840
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
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
               Begin Threed.SSPanel SSPanel30 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   87
                  Top             =   1200
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
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
               Begin Threed.SSPanel SSPanel31 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   88
                  Top             =   1560
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "기존 H/D 규격(HOLDER)"
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
                  Left            =   120
                  TabIndex        =   89
                  Top             =   1920
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "RPM/V(분당회전수)"
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
                  Left            =   120
                  TabIndex        =   90
                  Top             =   2280
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "DEPTH(절입량)"
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
               Begin Threed.SSPanel SSPanel34 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   91
                  Top             =   2640
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "FEEDRATE(이송)"
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
               Begin Threed.SSPanel SSPanel35 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   92
                  Top             =   3000
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "전체C/T,공정C/T"
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
               Begin Threed.SSPanel SSPanel36 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   93
                  Top             =   3360
                  Width           =   2130
                  _Version        =   65536
                  _ExtentX        =   3757
                  _ExtentY        =   582
                  _StockProps     =   15
                  Caption         =   "절삭유"
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
               Begin FPSpreadADO.fpSpread spd3 
                  Height          =   2220
                  Left            =   120
                  TabIndex        =   108
                  Top             =   3720
                  Width           =   4170
                  _Version        =   458752
                  _ExtentX        =   7355
                  _ExtentY        =   3916
                  _StockProps     =   64
                  Enabled         =   0   'False
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
                  MaxCols         =   4
                  MaxRows         =   6
                  RetainSelBlock  =   0   'False
                  ScrollBars      =   0
                  SpreadDesigner  =   "mkpoen05C.frx":0EA8
               End
               Begin VB.Label Label7 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   3225
                  TabIndex        =   96
                  Top             =   2010
                  Width           =   105
               End
               Begin VB.Label Label8 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   3225
                  TabIndex        =   95
                  Top             =   3075
                  Width           =   105
               End
               Begin VB.Label Label9 
                  Caption         =   "-"
                  Height          =   195
                  Left            =   3225
                  TabIndex        =   94
                  Top             =   2715
                  Width           =   105
               End
            End
            Begin Threed.SSPanel SSPanel1 
               Height          =   330
               Index           =   1
               Left            =   240
               TabIndex        =   97
               Top             =   500
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   582
               _StockProps     =   15
               Caption         =   "테스트일자"
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
            Begin XLibrary_XButton.XButton btn_clear1 
               Height          =   330
               Left            =   4230
               TabIndex        =   98
               TabStop         =   0   'False
               Top             =   500
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
               Left            =   3360
               TabIndex        =   99
               Top             =   500
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
            Begin XLibrary_XButton.XButton btn_add 
               Height          =   435
               Left            =   9480
               TabIndex        =   45
               Top             =   8805
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   767
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
            Begin XLibrary_XButton.XButton btn_del 
               Height          =   435
               Left            =   12360
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   8805
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   767
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
            Begin XLibrary_XButton.XButton btn_mod1 
               Height          =   435
               Left            =   10920
               TabIndex        =   109
               Top             =   8805
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
         End
      End
   End
   Begin Threed.SSPanel SSPanel9 
      Height          =   330
      Left            =   120
      TabIndex        =   47
      Top             =   5040
      Width           =   1770
      _Version        =   65536
      _ExtentX        =   3122
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "RPM/V(분당회전수)"
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
Attribute VB_Name = "mkpoen05C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_add_Click()
    
    On Error GoTo err_rtn
    
    Dim seq As Integer
    Dim lno As Integer
    
    Dim ii As Integer
    Dim gcnt(12) As Integer
    Dim gcnt2(12) As Integer
    Dim gcnt3(12) As Integer
    Dim resp        As Integer
    
    resp = MsgBox("테스트시트를 등록하시겠습니까??", vbQuestion + vbYesNo)
    If resp = vbNo Then Exit Sub
    
    If Check_Data = 9 Then Exit Sub
    
    '가공수 담는 작업
    For ii = 1 To 12
        
        If ii < 7 Then
        
            spd1.Col = 2
            spd1.row = ii
        
            gcnt(ii) = Val(spd1.Text)
        
        End If
        
        If ii > 6 Then
            
            spd1.Col = 4
            spd1.row = ii - 6
            
            gcnt(ii) = Val(spd1.Text)
            
        End If
        
    Next ii
    
    For ii = 1 To 12
        
        If ii < 7 Then
        
            spd2.Col = 2
            spd2.row = ii
        
            gcnt2(ii) = Val(spd2.Text)
        
        End If
        
        If ii > 6 Then
            
            spd2.Col = 4
            spd2.row = ii - 6
            
            gcnt2(ii) = Val(spd2.Text)
            
        End If
        
    Next ii
    
    For ii = 1 To 12
        
        If ii < 7 Then
        
            spd3.Col = 2
            spd3.row = ii
        
            gcnt3(ii) = Val(spd3.Text)
        
        End If
        
        If ii > 6 Then
            
            spd3.Col = 4
            spd3.row = ii - 6
            
            gcnt3(ii) = Val(spd3.Text)
            
        End If
        
    Next ii
    
    '순번조회
    sss = "      select nvl(max(sth_seq),0) + 1 as seq from man_stesthd"
    sss = sss & " where sth_dat = " & Format(Now, "yyyymmdd")
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    seq = Rs!seq
    Rs.Close
    
    Ws.BeginTrans
    
    sss = "insert into man_stesthd("
    sss = sss & "                   sth_dat"
    sss = sss & "                  ,sth_seq"
    sss = sss & "                  ,sth_tsab"
    sss = sss & "                  ,sth_tmcd"
    sss = sss & "                  ,sth_title"
    sss = sss & "                  ,sth_lot"
    sss = sss & "                  ,sth_gocd"
    sss = sss & "                  ,sth_remark"
    sss = sss & "                 ) values ("
    sss = sss & Format(Now, "yyyymmdd")
    sss = sss & "," & seq
    sss = sss & "," & Val(txt_sab)
    sss = sss & ",'" & Trim(txt_mcd) & "'"
    sss = sss & ",'" & Trim(txt_title) & "'"
    sss = sss & ",'" & Trim(txt_lot) & "'"
    sss = sss & ",'" & Trim(txt_gocd) & "'"
    sss = sss & ",'" & Trim(txt_remark) & "'"
    sss = sss & ")"
    
    db.Execute sss, 64
    '
    For ii = 1 To 3
        
        lno = ii
        
        If ii = 3 And chk_use.Value = 0 Then GoTo nextstep
    
        sss = "insert into man_stestds("
        sss = sss & "                   std_dat"
        sss = sss & "                  ,std_seq"
        sss = sss & "                  ,std_lno"
        sss = sss & "                  ,std_maker"
        sss = sss & "                  ,std_tipstd"
        sss = sss & "                  ,std_tipjil"
        sss = sss & "                  ,std_holder"
        sss = sss & "                  ,std_rcntmn"
        sss = sss & "                  ,std_rcntmx"
        sss = sss & "                  ,std_depth"
        sss = sss & "                  ,std_movmn"
        sss = sss & "                  ,std_movmx"
        sss = sss & "                  ,std_tct"
        sss = sss & "                  ,std_pct"
        sss = sss & "                  ,std_fluid"
        sss = sss & "                  ,std_result1"
        sss = sss & "                  ,std_result2"
        sss = sss & "                  ,std_result3"
        sss = sss & "                  ,std_result4"
        sss = sss & "                  ,std_result5"
        sss = sss & "                  ,std_result6"
        sss = sss & "                  ,std_result7"
        sss = sss & "                  ,std_result8"
        sss = sss & "                  ,std_result9"
        sss = sss & "                  ,std_result10"
        sss = sss & "                  ,std_result11"
        sss = sss & "                  ,std_result12"
        sss = sss & "                  ) values ("
        sss = sss & Format(Now, "yyyymmdd")
        sss = sss & "," & seq
        sss = sss & "," & lno
        sss = sss & ",'" & Trim(txt_maker(ii)) & "'"
        sss = sss & ",'" & Trim(txt_tipstd(ii)) & "'"
        sss = sss & ",'" & Trim(txt_tipjil(ii)) & "'"
        sss = sss & ",'" & Trim(txt_holder(ii)) & "'"
        sss = sss & "," & Val(txt_rcntmn(ii))
        sss = sss & "," & Val(txt_rcntmx(ii))
        sss = sss & "," & Val(txt_depth(ii))
        sss = sss & "," & Val(txt_movmn(ii))
        sss = sss & "," & Val(txt_movmx(ii))
        sss = sss & "," & Val(txt_tct(ii))
        sss = sss & "," & Val(txt_pct(ii))
        sss = sss & "," & Val(Left(cmb_fluid(ii), 1))
        If lno = 1 Then
        sss = sss & "," & Val(gcnt(1))
        sss = sss & "," & Val(gcnt(2))
        sss = sss & "," & Val(gcnt(3))
        sss = sss & "," & Val(gcnt(4))
        sss = sss & "," & Val(gcnt(5))
        sss = sss & "," & Val(gcnt(6))
        sss = sss & "," & Val(gcnt(7))
        sss = sss & "," & Val(gcnt(8))
        sss = sss & "," & Val(gcnt(9))
        sss = sss & "," & Val(gcnt(10))
        sss = sss & "," & Val(gcnt(11))
        sss = sss & "," & Val(gcnt(12))
        ElseIf lno = 2 Then
        sss = sss & "," & Val(gcnt2(1))
        sss = sss & "," & Val(gcnt2(2))
        sss = sss & "," & Val(gcnt2(3))
        sss = sss & "," & Val(gcnt2(4))
        sss = sss & "," & Val(gcnt2(5))
        sss = sss & "," & Val(gcnt2(6))
        sss = sss & "," & Val(gcnt2(7))
        sss = sss & "," & Val(gcnt2(8))
        sss = sss & "," & Val(gcnt2(9))
        sss = sss & "," & Val(gcnt2(10))
        sss = sss & "," & Val(gcnt2(11))
        sss = sss & "," & Val(gcnt2(12))
        ElseIf lno = 3 Then
        sss = sss & "," & Val(gcnt3(1))
        sss = sss & "," & Val(gcnt3(2))
        sss = sss & "," & Val(gcnt3(3))
        sss = sss & "," & Val(gcnt3(4))
        sss = sss & "," & Val(gcnt3(5))
        sss = sss & "," & Val(gcnt3(6))
        sss = sss & "," & Val(gcnt3(7))
        sss = sss & "," & Val(gcnt3(8))
        sss = sss & "," & Val(gcnt3(9))
        sss = sss & "," & Val(gcnt3(10))
        sss = sss & "," & Val(gcnt3(11))
        sss = sss & "," & Val(gcnt3(12))
        End If
        sss = sss & ")"
    
        db.Execute sss, 64
        
nextstep:

    Next ii
    
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
    sss = sss & "apl_4dat"
    sss = sss & ") values ("
    sss = sss & "'man_stesthd'" & ","
    sss = sss & Format(Now, "yyyymmdd") & ","
    sss = sss & seq & ","
    sss = sss & Gsab & ","
    sss = sss & "'N','N','N','N',"
    sss = sss & "0,0,0,0,0,0,0,0"
    sss = sss & ")"
    
    db.Execute sss, 64
    
    Ws.CommitTrans
    
    mkpoen05MDI.msg = "테스트시트 등록이 완료되었습니다."
    
    Exit Sub
    
err_rtn:

    Ws.Rollback
    mkpoen05MDI.msg = Err.Description
    
End Sub

Function Check_Data()

    Dim ii As Integer
    
    If Len(txt_sab) = 0 Then
        Check_Data = 9
        mkpoen05MDI.msg = "작업자를 확인하세요."
        txt_sab.SetFocus
        Exit Function
    End If
    
    If Len(txt_name) = 0 Then
        Check_Data = 9
        mkpoen05MDI.msg = "작업자를 확인하세요."
        txt_sab.SetFocus
        Exit Function
    End If
    
    If Len(txt_mcd) = 0 Then
        Check_Data = 9
        mkpoen05MDI.msg = "장비코드를 확인하세요."
        txt_mcd.SetFocus
        Exit Function
    End If
    
    If Len(txt_mcdname) = 0 Then
        Check_Data = 9
        mkpoen05MDI.msg = "장비코드를 확인하세요."
        txt_mcd.SetFocus
        Exit Function
    End If
    
    If Len(txt_title) = 0 Then
        Check_Data = 9
        mkpoen05MDI.msg = "테스트명을 확인하세요."
        txt_title.SetFocus
        Exit Function
    End If
    
    If Len(txt_lot) = 0 Then
        Check_Data = 9
        mkpoen05MDI.msg = "LOT NO 를 확인하세요."
        txt_lot.SetFocus
        Exit Function
    End If
    
    If Len(txt_bpcd) = 0 Then
        Check_Data = 9
        mkpoen05MDI.msg = "LOT NO 를 확인하세요."
        txt_lot.SetFocus
        Exit Function
    End If
    
    If Len(txt_jacd) = 0 Then
        Check_Data = 9
        mkpoen05MDI.msg = "LOT NO 를 확인하세요."
        txt_jacd.SetFocus
        Exit Function
    End If
    
    If Len(txt_gocd) = 0 Then
        Check_Data = 9
        mkpoen05MDI.msg = "공정코드를 확인하세요."
        txt_gocd.SetFocus
        Exit Function
    End If
    
    For ii = 1 To 3
        
        If ii = 3 And chk_use.Value = 0 Then
            Exit Function
        End If
        
        If Len(txt_maker(ii)) = 0 Then
            Check_Data = 9
            If ii = 1 Then mkpoen05MDI.msg = "기존 MARKER를 확인하세요."
            If ii = 2 Then mkpoen05MDI.msg = "테스트1 MARKER를 확인하세요."
            If ii = 3 Then mkpoen05MDI.msg = "테스트2 MARKER를 확인하세요."
            txt_maker(ii).SetFocus
            Exit Function
        End If
        
        If Len(txt_tipstd(ii)) = 0 Then
            Check_Data = 9
            If ii = 1 Then mkpoen05MDI.msg = "기존 TIP 규격을 확인하세요."
            If ii = 2 Then mkpoen05MDI.msg = "테스트1 TIP 규격을 확인하세요."
            If ii = 3 Then mkpoen05MDI.msg = "테스트2 TIP 규격을 확인하세요."
            txt_tipstd(ii).SetFocus
            Exit Function
        End If
        
        If Len(txt_tipjil(ii)) = 0 Then
            Check_Data = 9
            If ii = 1 Then mkpoen05MDI.msg = "기존 TIP 재질을 확인하세요."
            If ii = 2 Then mkpoen05MDI.msg = "테스트1 TIP 재질을 확인하세요."
            If ii = 3 Then mkpoen05MDI.msg = "테스트2 TIP 재질을 확인하세요."
            txt_tipjil(ii).SetFocus
            Exit Function
        End If
        
        If Len(txt_holder(ii)) = 0 Then
            Check_Data = 9
            If ii = 1 Then mkpoen05MDI.msg = "기존 H/D 규격을 확인하세요."
            If ii = 2 Then mkpoen05MDI.msg = "테스트1 H/D 규격을 확인하세요."
            If ii = 3 Then mkpoen05MDI.msg = "테스트2 H/D 규격을 확인하세요."
            txt_holder(ii).SetFocus
            Exit Function
        End If
        
        If Len(txt_rcntmn(ii)) = 0 Then
            Check_Data = 9
            If ii = 1 Then mkpoen05MDI.msg = "기존 RPM/V(min) 를 확인하세요."
            If ii = 2 Then mkpoen05MDI.msg = "테스트1 RPM/V(min) 를 확인하세요."
            If ii = 3 Then mkpoen05MDI.msg = "테스트2 RPM/V(min) 를 확인하세요."
            txt_rcntmn(ii).SetFocus
            Exit Function
        End If
        
        If Len(txt_rcntmx(ii)) = 0 Then
            Check_Data = 9
            If ii = 1 Then mkpoen05MDI.msg = "기존 RPM/V(max) 를 확인하세요."
            If ii = 2 Then mkpoen05MDI.msg = "테스트1 RPM/V(max) 를 확인하세요."
            If ii = 3 Then mkpoen05MDI.msg = "테스트2 RPM/V(max) 를 확인하세요."
            txt_rcntmx(ii).SetFocus
            Exit Function
        End If
        
        If Len(txt_depth(ii)) = 0 Then
            Check_Data = 9
            If ii = 1 Then mkpoen05MDI.msg = "기존 DEPTH(접입량) 를 확인하세요."
            If ii = 2 Then mkpoen05MDI.msg = "테스트1 DEPTH(접입량) 를 확인하세요."
            If ii = 3 Then mkpoen05MDI.msg = "테스트2 DEPTH(접입량) 를 확인하세요."
            txt_depth(ii).SetFocus
            Exit Function
        End If
        
        If Len(txt_movmn(ii)) = 0 Then
            Check_Data = 9
            If ii = 1 Then mkpoen05MDI.msg = "기존 FEEDRATE(min) 를 확인하세요."
            If ii = 2 Then mkpoen05MDI.msg = "테스트1 FEEDRATE(min) 를 확인하세요."
            If ii = 3 Then mkpoen05MDI.msg = "테스트2 FEEDRATE(min) 를 확인하세요."
            txt_movmn(ii).SetFocus
            Exit Function
        End If
        
        If Len(txt_movmx(ii)) = 0 Then
            Check_Data = 9
            If ii = 1 Then mkpoen05MDI.msg = "기존 FEEDRATE(max) 를 확인하세요."
            If ii = 2 Then mkpoen05MDI.msg = "테스트1 FEEDRATE(max) 를 확인하세요."
            If ii = 3 Then mkpoen05MDI.msg = "테스트2 FEEDRATE(max) 를 확인하세요."
            txt_movmx(ii).SetFocus
            Exit Function
        End If
        
        If Len(txt_tct(ii)) = 0 Then
            Check_Data = 9
            If ii = 1 Then mkpoen05MDI.msg = "기존 전체C/T 를 확인하세요."
            If ii = 2 Then mkpoen05MDI.msg = "테스트1 전체C/T 를 확인하세요."
            If ii = 3 Then mkpoen05MDI.msg = "테스트2 전체C/T 를 확인하세요."
            txt_tct(ii).SetFocus
            Exit Function
        End If
        
        If Len(txt_pct(ii)) = 0 Then
            Check_Data = 9
            If ii = 1 Then mkpoen05MDI.msg = "기존 공정C/T 를 확인하세요."
            If ii = 2 Then mkpoen05MDI.msg = "테스트1 공정C/T 를 확인하세요."
            If ii = 3 Then mkpoen05MDI.msg = "테스트2 공정C/T 를 확인하세요."
            txt_pct(ii).SetFocus
            Exit Function
        End If
        
        If Len(cmb_fluid(ii)) = 0 Then
            Check_Data = 9
            If ii = 1 Then mkpoen05MDI.msg = "기존 절삭유를 확인하세요."
            If ii = 2 Then mkpoen05MDI.msg = "테스트1 절삭유를 확인하세요."
            If ii = 3 Then mkpoen05MDI.msg = "테스트2 절삭유를 확인하세요."
            cmb_fluid(ii).SetFocus
            Exit Function
        End If
        
    Next ii
    
End Function

Private Sub btn_clear1_Click()

    Dim ii As Integer
    
    With spd_app
    
        .row = 1
        .Col = 1: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 2: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 3: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        .Col = 4: .CellType = CellTypeButton: .TypeButtonText = "Click": .TypeButtonColor = RGB(234, 234, 234): .Lock = False: .CellTag = 0: .Text = ""
        
        .row = 2
        .Col = 1: .Text = ""
        .Col = 2: .Text = ""
        .Col = 3: .Text = ""
        .Col = 4: .Text = ""
        
    End With
    
    txt_sab.Text = ""
    txt_name.Text = ""
    txt_mcd.Text = ""
    txt_mcdname.Text = ""
    txt_title.Text = ""
    txt_lot.Text = ""
    txt_gocd.Text = ""
    txt_bpcd.Text = ""
    txt_jacd.Text = ""
    txt_rmk.Text = ""
    
    For ii = 1 To 3
        
        txt_maker(ii).Text = ""
        txt_tipstd(ii).Text = ""
        txt_tipjil(ii).Text = ""
        txt_holder(ii).Text = ""
        txt_rcntmn(ii).Text = ""
        txt_rcntmx(ii).Text = ""
        txt_depth(ii).Text = ""
        txt_movmn(ii).Text = ""
        txt_movmx(ii).Text = ""
        txt_tct(ii).Text = ""
        txt_pct(ii).Text = ""
        cmb_fluid(ii).ListIndex = 0
        
    Next ii
    
    chk_use.Value = 0
    
    With spd1
        
        .Col = 2
        For ii = 1 To 6
            .row = ii
            .Text = ""
        Next ii
        
        .Col = 4
        For ii = 1 To 6
            .row = ii
            .Text = ""
        Next ii
        
    End With
    
    With spd2
        
        .Col = 2
        For ii = 1 To 6
            .row = ii
            .Text = ""
        Next ii
        
        .Col = 4
        For ii = 1 To 6
            .row = ii
            .Text = ""
        Next ii
        
    End With
    
    With spd3
        
        .Col = 2
        For ii = 1 To 6
            .row = ii
            .Text = ""
        Next ii
        
        .Col = 4
        For ii = 1 To 6
            .row = ii
            .Text = ""
        Next ii
        
    End With
    
End Sub

Private Sub btn_del_Click()

    On Error GoTo err_rtn

    Dim resp        As Integer
    
    If Len(txt_dat) <> 8 Then
        mkpoen05MDI.msg = "테스트 일자를 확인하세요."
        txt_dat.SetFocus
        Exit Sub
    End If
        
    If Len(txt_seq) = 0 Then
        mkpoen05MDI.msg = "테스트 순번을 확인하세요."
        txt_seq.SetFocus
        Exit Sub
    End If
    
    resp = MsgBox("테스트시트를 삭제하시겠습니까??", vbQuestion + vbYesNo)
    If resp = vbNo Then Exit Sub
    
    sss = "select * from man_stesthd"
    sss = sss & " where sth_dat = " & Val(txt_dat)
    sss = sss & "   and sth_seq = " & Val(txt_seq)
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    
    If Rs.RecordCount = 0 Then
        Rs.Close
        mkpoen05MDI.msg = "존재하지 않는 테스트시트 입니다."
        Exit Sub
    End If
    
    Ws.BeginTrans
    
    sss = "delete from man_stesthd"
    sss = sss & " where sth_dat = " & Val(txt_dat)
    sss = sss & "   and sth_seq = " & Val(txt_seq)
    
    db.Execute sss, 64
    
    sss = "delete from man_stestds"
    sss = sss & " where std_dat = " & Val(txt_dat)
    sss = sss & "   and std_seq = " & Val(txt_seq)
    
    db.Execute sss, 64
    
    sss = "delete from oth_applist"
    sss = sss & " where apl_table = 'man_stesthd'"
    sss = sss & "   and apl_tdat = " & Val(txt_dat)
    sss = sss & "   and apl_tseq = " & Val(txt_seq)
    
    db.Execute sss, 64
    
    Ws.CommitTrans
    
    Call btn_clear1_Click
    
    mkpoen05MDI.msg = "테스트 시트가 삭제되었습니다."
    
    Exit Sub
    
err_rtn:

    Ws.Rollback
    mkpoen05MDI.msg = Err.Description
    
End Sub

Private Sub btn_mod1_Click()

    On Error GoTo err_rtn

    Dim resp        As Integer
    Dim gcnt(12) As Integer
    Dim gcnt2(12) As Integer
    Dim gcnt3(12) As Integer
    Dim ii As Integer
        
    If Len(txt_dat) <> 8 Then
        mkpoen05MDI.msg = "테스트 일자를 확인하세요."
        txt_dat.SetFocus
        Exit Sub
    End If
        
    If Len(txt_seq) = 0 Then
        mkpoen05MDI.msg = "테스트 순번을 확인하세요."
        txt_seq.SetFocus
        Exit Sub
    End If
    
    resp = MsgBox("테스트시트를 수정하시겠습니까??", vbQuestion + vbYesNo)
    If resp = vbNo Then Exit Sub

    '가공수 담는 작업
    For ii = 1 To 12
        
        If ii < 7 Then
        
            spd1.Col = 2
            spd1.row = ii
        
            gcnt(ii) = Val(spd1.Text)
        
        End If
        
        If ii > 6 Then
            
            spd1.Col = 4
            spd1.row = ii - 6
            
            gcnt(ii) = Val(spd1.Text)
            
        End If
        
    Next ii
    
    For ii = 1 To 12
        
        If ii < 7 Then
        
            spd2.Col = 2
            spd2.row = ii
        
            gcnt2(ii) = Val(spd2.Text)
        
        End If
        
        If ii > 6 Then
            
            spd2.Col = 4
            spd2.row = ii - 6
            
            gcnt2(ii) = Val(spd2.Text)
            
        End If
        
    Next ii
    
    For ii = 1 To 12
        
        If ii < 7 Then
        
            spd3.Col = 2
            spd3.row = ii
        
            gcnt3(ii) = Val(spd3.Text)
        
        End If
        
        If ii > 6 Then
            
            spd3.Col = 4
            spd3.row = ii - 6
            
            gcnt3(ii) = Val(spd3.Text)
            
        End If
        
    Next ii
    
    sss = "select * from man_stesthd"
    sss = sss & " where sth_dat = " & Val(txt_dat)
    sss = sss & "   and sth_seq = " & Val(txt_seq)
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    
    If Rs.RecordCount = 0 Then
        Rs.Close
        mkpoen05MDI.msg = "존재하지 않는 테스트시트 입니다."
        Exit Sub
    End If
    
    Ws.BeginTrans
    
    sss = "update man_stesthd"
    sss = sss & " set sth_tsab = " & Val(txt_sab)
    sss = sss & "," & "sth_tmcd = '" & Trim(txt_mcd) & "'"
    sss = sss & "," & "sth_title = '" & Trim(txt_title) & "'"
    sss = sss & "," & "sth_lot = '" & Trim(txt_lot) & "'"
    sss = sss & "," & "sth_gocd = '" & Trim(txt_gocd) & "'"
    sss = sss & "," & "sth_remark = '" & Trim(txt_remark) & "'"
    sss = sss & " where sth_dat = " & Val(txt_dat)
    sss = sss & "   and sth_seq = " & Val(txt_seq)
    
    db.Execute sss, 64
    
    sss = "delete from man_stestds"
    sss = sss & " where std_dat = " & Val(txt_dat)
    sss = sss & "   and std_seq = " & Val(txt_seq)
    
    db.Execute sss, 64
    
    For ii = 1 To 3
        
        lno = ii
        
        If ii = 3 And chk_use.Value = 0 Then GoTo nextstep
    
        sss = "insert into man_stestds("
        sss = sss & "                   std_dat"
        sss = sss & "                  ,std_seq"
        sss = sss & "                  ,std_lno"
        sss = sss & "                  ,std_maker"
        sss = sss & "                  ,std_tipstd"
        sss = sss & "                  ,std_tipjil"
        sss = sss & "                  ,std_holder"
        sss = sss & "                  ,std_rcntmn"
        sss = sss & "                  ,std_rcntmx"
        sss = sss & "                  ,std_depth"
        sss = sss & "                  ,std_movmn"
        sss = sss & "                  ,std_movmx"
        sss = sss & "                  ,std_tct"
        sss = sss & "                  ,std_pct"
        sss = sss & "                  ,std_fluid"
        sss = sss & "                  ,std_result1"
        sss = sss & "                  ,std_result2"
        sss = sss & "                  ,std_result3"
        sss = sss & "                  ,std_result4"
        sss = sss & "                  ,std_result5"
        sss = sss & "                  ,std_result6"
        sss = sss & "                  ,std_result7"
        sss = sss & "                  ,std_result8"
        sss = sss & "                  ,std_result9"
        sss = sss & "                  ,std_result10"
        sss = sss & "                  ,std_result11"
        sss = sss & "                  ,std_result12"
        sss = sss & "                  ) values ("
        sss = sss & Val(txt_dat)
        sss = sss & "," & Val(txt_seq)
        sss = sss & "," & lno
        sss = sss & ",'" & Trim(txt_maker(ii)) & "'"
        sss = sss & ",'" & Trim(txt_tipstd(ii)) & "'"
        sss = sss & ",'" & Trim(txt_tipjil(ii)) & "'"
        sss = sss & ",'" & Trim(txt_holder(ii)) & "'"
        sss = sss & "," & Val(txt_rcntmn(ii))
        sss = sss & "," & Val(txt_rcntmx(ii))
        sss = sss & "," & Val(txt_depth(ii))
        sss = sss & "," & Val(txt_movmn(ii))
        sss = sss & "," & Val(txt_movmx(ii))
        sss = sss & "," & Val(txt_tct(ii))
        sss = sss & "," & Val(txt_pct(ii))
        sss = sss & "," & Val(Left(cmb_fluid(ii), 1))
        If lno = 1 Then
        sss = sss & "," & Val(gcnt(1))
        sss = sss & "," & Val(gcnt(2))
        sss = sss & "," & Val(gcnt(3))
        sss = sss & "," & Val(gcnt(4))
        sss = sss & "," & Val(gcnt(5))
        sss = sss & "," & Val(gcnt(6))
        sss = sss & "," & Val(gcnt(7))
        sss = sss & "," & Val(gcnt(8))
        sss = sss & "," & Val(gcnt(9))
        sss = sss & "," & Val(gcnt(10))
        sss = sss & "," & Val(gcnt(11))
        sss = sss & "," & Val(gcnt(12))
        ElseIf lno = 2 Then
        sss = sss & "," & Val(gcnt2(1))
        sss = sss & "," & Val(gcnt2(2))
        sss = sss & "," & Val(gcnt2(3))
        sss = sss & "," & Val(gcnt2(4))
        sss = sss & "," & Val(gcnt2(5))
        sss = sss & "," & Val(gcnt2(6))
        sss = sss & "," & Val(gcnt2(7))
        sss = sss & "," & Val(gcnt2(8))
        sss = sss & "," & Val(gcnt2(9))
        sss = sss & "," & Val(gcnt2(10))
        sss = sss & "," & Val(gcnt2(11))
        sss = sss & "," & Val(gcnt2(12))
        ElseIf lno = 3 Then
        sss = sss & "," & Val(gcnt3(1))
        sss = sss & "," & Val(gcnt3(2))
        sss = sss & "," & Val(gcnt3(3))
        sss = sss & "," & Val(gcnt3(4))
        sss = sss & "," & Val(gcnt3(5))
        sss = sss & "," & Val(gcnt3(6))
        sss = sss & "," & Val(gcnt3(7))
        sss = sss & "," & Val(gcnt3(8))
        sss = sss & "," & Val(gcnt3(9))
        sss = sss & "," & Val(gcnt3(10))
        sss = sss & "," & Val(gcnt3(11))
        sss = sss & "," & Val(gcnt3(12))
        End If
        sss = sss & ")"
    
        db.Execute sss, 64
        
nextstep:

    Next ii
    
    Ws.CommitTrans
    
    mkpoen05MDI.msg = "테스트 시트가 수정되었습니다."
    
    Exit Sub
    
err_rtn:

    Ws.Rollback
    mkpoen05MDI.msg = Err.Description
    
End Sub

Private Sub btn_view1_Click()
    
    'read
    
    'write
    
    On Error GoTo err_rtn

    Dim lno As Integer
    Dim ii As Integer
    
    If Len(txt_dat) <> 8 Then
        mkpoen05MDI.msg = "테스트 일자를 확인하세요."
        txt_dat.SetFocus
        Exit Sub
    End If
        
    If Len(txt_seq) = 0 Then
        mkpoen05MDI.msg = "테스트 순번을 확인하세요."
        txt_seq.SetFocus
        Exit Sub
    End If
    
    sss = "select sth_dat,sth_seq,sth_tsab,sth_tmcd,sth_title,sth_lot,sth_gocd,sth_remark,"
    sss = sss & " apl_tdat,apl_isab,sinbun_name(apl_isab) as apl_iname,"
    sss = sss & " apl_1yn,apl_1dat,apl_1sab,sinbun_name(apl_1sab) as apl_1name,"
    sss = sss & " apl_2yn,apl_2dat,apl_2sab,sinbun_name(apl_2sab) as apl_2name,"
    sss = sss & " apl_3yn,apl_3dat,apl_3sab,sinbun_name(apl_3sab) as apl_3name"
    sss = sss & " from man_stesthd,oth_applist"
    sss = sss & " where sth_dat = " & Val(txt_dat)
    sss = sss & "   and sth_seq = " & Val(txt_seq)
    sss = sss & "   and apl_tdat = sth_dat"
    sss = sss & "   and apl_tseq = sth_seq"
    sss = sss & "   and apl_table = 'man_stesthd'"
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    
    If Rs.RecordCount < 1 Then
        Rs.Close
        mkpoen05MDI.msg = "테스트일자 순번을 확인하세요."
        txt_dat.SetFocus
        Exit Sub
    End If
    
    Call btn_clear1_Click
    
'    If Not IsNull(Rs!sth_dat) Then txt_tdat = Rs!sth_dat
'    If Not IsNull(Rs!sth_seq) Then txt_tseq = Rs!sth_seq
    If Not IsNull(Rs!sth_tsab) Then txt_sab = Rs!sth_tsab
    If Not IsNull(Rs!sth_tmcd) Then txt_mcd = Rs!sth_tmcd
    If Not IsNull(Rs!sth_title) Then txt_title = Rs!sth_title
    If Not IsNull(Rs!sth_lot) Then txt_lot = Rs!sth_lot
    If Not IsNull(Rs!sth_gocd) Then txt_gocd = Rs!sth_gocd
    If Not IsNull(Rs!sth_remark) Then txt_remark = Rs!sth_remark
    
    With spd_app
        
        '담당
        .Col = 1
        
        .row = 1
        .CellType = CellTypeEdit: .Text = Rs!apl_iname: .CellTag = Rs!apl_isab: .Lock = True
        .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
        
        .row = 2
        .Text = Format(Val(Rs!apl_tdat), "####/##/##")
        
        '검토1
        .Col = 2
        If Rs!apl_1sab > 0 Then
            
            If Rs!apl_1dat > 0 Then
                
                .row = 1: .CellType = CellTypeEdit: .Text = Rs!apl_1name: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                .row = 2: .Text = Format(Val(Rs!apl_1dat), "####/##/##")
            Else
                .CellType = CellTypeButton: .TypeButtonText = Rs!apl_1name: .CellTag = Rs!apl_1sab
                .row = 2: .Text = ""
            End If
            
        End If
        
        '검토2
        .Col = 3
        If Rs!apl_2sab > 0 Then
            
            If Rs!apl_2dat > 0 Then
                .row = 1: .CellType = CellTypeEdit: .Text = Rs!apl_2name: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
                .row = 2: .Text = Format(Val(Rs!apl_2dat), "####/##/##")
            Else
                .CellType = CellTypeButton: .TypeButtonText = Rs!apl_2name: .CellTag = Rs!apl_2sab
                .row = 2: .Text = ""
            End If
            
        End If
        
        '승인
        .Col = 4
        If Rs!apl_3sab > 0 Then
            
            If Rs!apl_3dat > 0 Then
                .row = 1: .CellType = CellTypeEdit: .Text = Rs!apl_3name: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(242, 220, 219)
                .row = 2: .Text = Format(Val(Rs!apl_3dat), "####/##/##")
            Else
                .CellType = CellTypeButton: .TypeButtonText = Rs!apl_3name: .CellTag = Rs!apl_3sab
                .row = 2: .Text = ""
            End If
            
        End If
        
'        .Col = 2
'        If Not IsNull(Rs!apl_tdat) Then .Text = Rs!apl_tdat
'
'        .Col = 1
'        If Not IsNull(Rs!apl_1sab) Then
'
'            If Rs!apl_1dat > 0 Then
'                .CellType = CellTypeEdit: .Text = Rs!apl_1name: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
'                .row = 2: .Text = Format(Val(Rs!apl_1dat), "####/##/##")
'            Else
'                .TypeButtonText = Rs!apl_1name: .CellTag = Rs!apl_1sab
'                .row = 2: .Text = ""
'            End If
'
'        End If
'
'        .Col = 1
'        If Not IsNull(Rs!apl_2sab) Then
'
'            If Rs!apl_2dat > 0 Then
'                .CellType = CellTypeEdit: .Text = Rs!apl_2name: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
'                .row = 2: .Text = Format(Val(Rs!apl_2dat), "####/##/##")
'            Else
'                .TypeButtonText = Rs!apl_2name: .CellTag = Rs!apl_2sab
'                .row = 2: .Text = ""
'            End If
'
'        End If
'
'        .Col = 1
'        If Not IsNull(Rs!apl_3sab) Then
'
'            If Rs!apl_3dat > 0 Then
'                .CellType = CellTypeEdit: .Text = Rs!apl_3name: .Lock = True: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter: .BackColor = RGB(215, 227, 188)
'                .row = 2: .Text = Format(Val(Rs!apl_3dat), "####/##/##")
'            Else
'                .TypeButtonText = Rs!apl_3name: .CellTag = Rs!apl_3sab
'                .row = 2: .Text = ""
'            End If
'
'        End If
        
    End With
    
    Rs.Close
    
    If Len(txt_mcd) > 0 Then Call txt_mcd_LostFocus
    If Len(txt_sab) > 0 Then Call txt_sab_LostFocus
    If Len(txt_lot) > 0 Then Call txt_lot_LostFocus
    
    sss = "select *"
    sss = sss & " from man_stestds"
    sss = sss & " where std_dat = " & Val(txt_dat)
    sss = sss & "   and std_seq = " & Val(txt_seq)
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    
    If Rs.RecordCount < 1 Then
        Rs.Close
        mkpoen05MDI.msg = "테스트일자 순번을 확인하세요."
        txt_dat.SetFocus
        Exit Sub
    End If
    
    Do While Not Rs.EOF
        
        If Not IsNull(Rs!std_lno) Then
        
            lno = Rs!std_lno
            
            If Not IsNull(Rs!std_maker) Then txt_maker(lno) = Rs!std_maker
            If Not IsNull(Rs!std_tipstd) Then txt_tipstd(lno) = Rs!std_tipstd
            If Not IsNull(Rs!std_tipjil) Then txt_tipjil(lno) = Rs!std_tipjil
            If Not IsNull(Rs!std_holder) Then txt_holder(lno) = Rs!std_holder
            If Not IsNull(Rs!std_rcntmn) Then txt_rcntmn(lno) = Rs!std_rcntmn
            If Not IsNull(Rs!std_rcntmx) Then txt_rcntmx(lno) = Rs!std_rcntmx
            If Not IsNull(Rs!std_depth) Then txt_depth(lno) = Rs!std_depth
            If Not IsNull(Rs!std_movmn) Then txt_movmn(lno) = Rs!std_movmn
            If Not IsNull(Rs!std_movmx) Then txt_movmx(lno) = Rs!std_movmx
            If Not IsNull(Rs!std_tct) Then txt_tct(lno) = Rs!std_tct
            If Not IsNull(Rs!std_pct) Then txt_pct(lno) = Rs!std_pct
            If Not IsNull(Rs!std_fluid) Then cmb_fluid(lno).ListIndex = Rs!std_fluid
            
            If lno = 1 Then
                With spd1
                
                    .Col = 2
                    
                    If Not IsNull(Rs!std_result1) Then .row = 1: .Text = Rs!std_result1
                    If Not IsNull(Rs!std_result2) Then .row = 2: .Text = Rs!std_result2
                    If Not IsNull(Rs!std_result3) Then .row = 3: .Text = Rs!std_result3
                    If Not IsNull(Rs!std_result4) Then .row = 4: .Text = Rs!std_result4
                    If Not IsNull(Rs!std_result5) Then .row = 5: .Text = Rs!std_result5
                    If Not IsNull(Rs!std_result6) Then .row = 6: .Text = Rs!std_result6
                    
                    .Col = 4
                    If Not IsNull(Rs!std_result7) Then .row = 1: .Text = Rs!std_result7
                    If Not IsNull(Rs!std_result8) Then .row = 2: .Text = Rs!std_result8
                    If Not IsNull(Rs!std_result9) Then .row = 3: .Text = Rs!std_result9
                    If Not IsNull(Rs!std_result10) Then .row = 4: .Text = Rs!std_result10
                    If Not IsNull(Rs!std_result11) Then .row = 5: .Text = Rs!std_result11
                    If Not IsNull(Rs!std_result12) Then .row = 6: .Text = Rs!std_result12
                    
                End With
            End If
            
            If lno = 2 Then
                With spd2
                
                    .Col = 2
                    
                    If Not IsNull(Rs!std_result1) Then .row = 1: .Text = Rs!std_result1
                    If Not IsNull(Rs!std_result2) Then .row = 2: .Text = Rs!std_result2
                    If Not IsNull(Rs!std_result3) Then .row = 3: .Text = Rs!std_result3
                    If Not IsNull(Rs!std_result4) Then .row = 4: .Text = Rs!std_result4
                    If Not IsNull(Rs!std_result5) Then .row = 5: .Text = Rs!std_result5
                    If Not IsNull(Rs!std_result6) Then .row = 6: .Text = Rs!std_result6
                    
                    .Col = 4
                    If Not IsNull(Rs!std_result7) Then .row = 1: .Text = Rs!std_result7
                    If Not IsNull(Rs!std_result8) Then .row = 2: .Text = Rs!std_result8
                    If Not IsNull(Rs!std_result9) Then .row = 3: .Text = Rs!std_result9
                    If Not IsNull(Rs!std_result10) Then .row = 4: .Text = Rs!std_result10
                    If Not IsNull(Rs!std_result11) Then .row = 5: .Text = Rs!std_result11
                    If Not IsNull(Rs!std_result12) Then .row = 6: .Text = Rs!std_result12
                    
                End With
            End If
            
            If lno = 3 Then
            
                chk_use.Value = 1
                
                With spd3
                
                    .Col = 2
                    
                    If Not IsNull(Rs!std_result1) Then .row = 1: .Text = Rs!std_result1
                    If Not IsNull(Rs!std_result2) Then .row = 2: .Text = Rs!std_result2
                    If Not IsNull(Rs!std_result3) Then .row = 3: .Text = Rs!std_result3
                    If Not IsNull(Rs!std_result4) Then .row = 4: .Text = Rs!std_result4
                    If Not IsNull(Rs!std_result5) Then .row = 5: .Text = Rs!std_result5
                    If Not IsNull(Rs!std_result6) Then .row = 6: .Text = Rs!std_result6
                    
                    .Col = 4
                    If Not IsNull(Rs!std_result7) Then .row = 1: .Text = Rs!std_result7
                    If Not IsNull(Rs!std_result8) Then .row = 2: .Text = Rs!std_result8
                    If Not IsNull(Rs!std_result9) Then .row = 3: .Text = Rs!std_result9
                    If Not IsNull(Rs!std_result10) Then .row = 4: .Text = Rs!std_result10
                    If Not IsNull(Rs!std_result11) Then .row = 5: .Text = Rs!std_result11
                    If Not IsNull(Rs!std_result12) Then .row = 6: .Text = Rs!std_result12
                    
                End With
            End If
            
        End If
        
        Rs.MoveNext
    Loop
    
    mkpoen05MDI.msg = "조회완료!"
    
    Exit Sub
    
err_rtn:

    mkpoen05MDI.msg = Err.Description
    
End Sub

Private Sub chk_use_Click()
    
    If chk_use.Value = 1 Then
        
        txt_maker(3).Enabled = True
        txt_tipstd(3).Enabled = True
        txt_tipjil(3).Enabled = True
        txt_holder(3).Enabled = True
        txt_rcntmn(3).Enabled = True
        txt_rcntmx(3).Enabled = True
        txt_depth(3).Enabled = True
        txt_movmn(3).Enabled = True
        txt_movmx(3).Enabled = True
        txt_tct(3).Enabled = True
        txt_pct(3).Enabled = True
        cmb_fluid(3).Enabled = True
        spd3.Enabled = True
        
        txt_maker(3).BackColor = 16777215
        txt_tipstd(3).BackColor = 16777215
        txt_tipjil(3).BackColor = 16777215
        txt_holder(3).BackColor = 16777215
        txt_rcntmn(3).BackColor = 16777215
        txt_rcntmx(3).BackColor = 16777215
        txt_depth(3).BackColor = 16777215
        txt_movmn(3).BackColor = 16777215
        txt_movmx(3).BackColor = 16777215
        txt_tct(3).BackColor = 16777215
        txt_pct(3).BackColor = 16777215
        cmb_fluid(3).BackColor = 16777215
        
    Else
    
        txt_maker(3).Enabled = False
        txt_tipstd(3).Enabled = False
        txt_tipjil(3).Enabled = False
        txt_holder(3).Enabled = False
        txt_rcntmn(3).Enabled = False
        txt_rcntmx(3).Enabled = False
        txt_depth(3).Enabled = False
        txt_movmn(3).Enabled = False
        txt_movmx(3).Enabled = False
        txt_tct(3).Enabled = False
        txt_pct(3).Enabled = False
        cmb_fluid(3).Enabled = False
        spd3.Enabled = False
        
        txt_maker(3).BackColor = 14737632
        txt_tipstd(3).BackColor = 14737632
        txt_tipjil(3).BackColor = 14737632
        txt_holder(3).BackColor = 14737632
        txt_rcntmn(3).BackColor = 14737632
        txt_rcntmx(3).BackColor = 14737632
        txt_depth(3).BackColor = 14737632
        txt_movmn(3).BackColor = 14737632
        txt_movmx(3).BackColor = 14737632
        txt_tct(3).BackColor = 14737632
        txt_pct(3).BackColor = 14737632
        cmb_fluid(3).BackColor = 14737632
        
    End If
    
End Sub

Private Sub spd_app_ButtonClicked(ByVal Col As Long, ByVal row As Long, ByVal ButtonDown As Integer)
    On Error GoTo err_rtn
                
    Dim resp        As Integer

    resp = MsgBox("결재하시겠습니까??", vbQuestion + vbYesNo)
    If resp = vbNo Then Exit Sub

    sss = "select * from man_stesthd where sth_dat = " & Val(txt_dat) & " and sth_seq = " & Val(txt_seq)
    
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    '
    If Rs.RecordCount < 1 Then
        Rs.Close
        mkpoen05MDI.msg = "테스트일자 순번을 확인하세요."
        txt_dat.SetFocus
        Exit Sub
    End If
    
    Rs.Close
    
    With spd_app
        
        .Col = Col
        .row = 1
        
        .CellTag = Gsab
        .CellType = CellTypeEdit: .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter
        .Text = GName
        .Lock = True
        
        If Col = 4 Then
            .BackColor = RGB(242, 220, 219)
        Else
            .BackColor = RGB(215, 227, 188)
        End If
        
        .row = 2
        .Text = Format(Format(Now, "yyyymmdd"), "####/##/##")
        
    End With
    
    Ws.BeginTrans
    
    sss = "update oth_applist set"
    sss = sss & " apl_" & Col - 1 & "yn = 'Y'" & ","
    sss = sss & " apl_" & Col - 1 & "dat = " & Format(Now, "yyyymmdd") & ","
    sss = sss & " apl_" & Col - 1 & "sab = " & Gsab
    sss = sss & " where apl_table = 'man_stesthd'"
    sss = sss & "   and apl_tdat = " & Val(txt_dat)
    sss = sss & "   and apl_tseq = " & Val(txt_seq)
    
    db.Execute sss, 64
    
    Ws.CommitTrans
    
    Exit Sub
    
err_rtn:

    Ws.Rollback
    mkpoen05MDI.msg = Err.Description
    
End Sub

Private Sub spd_app_DblClick(ByVal Col As Long, ByVal row As Long)

    On Error GoTo err_rtn
                
    Dim resp        As Integer

    resp = MsgBox("결재 취소하시겠습니까??", vbQuestion + vbYesNo)
    If resp = vbNo Then Exit Sub

    sss = "select * from man_stesthd where sth_dat = " & Val(txt_dat) & " and sth_seq = " & Val(txt_seq)
    
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    '
    If Rs.RecordCount < 1 Then
        Rs.Close
        mkpoen05MDI.msg = "테스트일자 순번을 확인하세요."
        txt_dat.SetFocus
        Exit Sub
    End If
    
    Rs.Close
    
    With spd_app
        
        .Col = Col
        .row = 1
        
        .CellTag = 0
        .CellType = CellTypeButton
        .Text = "Click"
        .TypeButtonColor = RGB(234, 234, 234)
        .Lock = False
      
        .row = 2
        .Text = ""
        
    End With
    
    Ws.BeginTrans
    
    sss = "update oth_applist set"
    sss = sss & " apl_" & Col - 1 & "yn = 'N'" & ","
    sss = sss & " apl_" & Col - 1 & "dat = 0,"
    sss = sss & " apl_" & Col - 1 & "sab = 0"
    sss = sss & " where apl_table = 'man_stesthd'"
    sss = sss & "   and apl_tdat = " & Val(txt_dat)
    sss = sss & "   and apl_tseq = " & Val(txt_seq)
    
    db.Execute sss, 64
    
    Ws.CommitTrans
    
    Exit Sub
    
err_rtn:

    Ws.Rollback
    mkpoen05MDI.msg = Err.Description
End Sub

'Private Sub Form_Load()
'
'    txt_tdat = Format(Now, "yyyymmdd")
'
'End Sub

Private Sub SSFrame1_Click()

End Sub

Private Sub txt_lot_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txt_lot_LostFocus()

    If Len(txt_lot) = 0 Then
        Exit Sub
    End If
    
    txt_lot = UCase(Trim(txt_lot))
    
    '순번조회
    sss = "      select * from man_direct"
    sss = sss & " where dit_lot = '" & txt_lot & "'"
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    
    If Rs.RecordCount < 1 Then
        txt_bpcd = "": txt_jacd = ""
        Rs.Close
        mkpoen05MDI.msg = "LOT NO를 확인하세요!"
        txt_lot.SetFocus
        Exit Sub
    End If
    
    If Not IsNull(Rs!dit_bpcd) Then txt_bpcd = Rs!dit_bpcd
    If Not IsNull(Rs!dit_jacd) Then txt_jacd = Rs!dit_jacd
    
    Rs.Close
    
End Sub

Private Sub txt_mcd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txt_mcd_LostFocus()
    
    If Len(txt_mcd) = 0 Then
        Exit Sub
    End If
    
    txt_mcd = UCase(Trim(txt_mcd))
    
    sss = "       select mhc_name"
    sss = sss & "   from man_machcd"
    sss = sss & "  where mhc_code = '" & txt_mcd & "'"
            
    Set Ks = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    If Ks.RecordCount < 1 Then
        txt_mcdname = ""
        Ks.Close: mkpoen05MDI.msg = "장비코드를 확인하세요!"
        txt_mcd.SetFocus
        Exit Sub
    End If
        
    If Not IsNull(Ks!mhc_name) Then txt_mcdname = Ks!mhc_name

    Ks.Close
    
End Sub

Private Sub txt_sab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txt_sab_LostFocus()
    
    If Len(txt_sab) = 0 Then
        Exit Sub
    End If
    
    sss = "      select sin_name "
    sss = sss & "  from peo_sinbun"
    sss = sss & " where sin_sab = " & Val(txt_sab)
    Set Rs = db.OpenRecordset(sss, dbOpenSnapshot, 64)
    
    If Rs.RecordCount < 1 Then
        txt_name = ""
        Rs.Close
        mkpoen05MDI.msg = "작업자 사번을 확인하세요!"
        txt_sab.SetFocus
        Exit Sub
    End If
    
    If Not IsNull(Rs!sin_name) Then txt_name = Rs!sin_name
    
    Rs.Close
    
End Sub
