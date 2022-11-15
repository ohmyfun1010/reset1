VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12285
   ClientLeft      =   6615
   ClientTop       =   1755
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   ScaleHeight     =   12285
   ScaleWidth      =   12075
   Begin Threed.SSPanel SSPanel1 
      Height          =   12225
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   11985
      _Version        =   65536
      _ExtentX        =   21140
      _ExtentY        =   21564
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   11415
         Left            =   420
         TabIndex        =   1
         Top             =   360
         Width           =   11055
         _Version        =   458752
         _ExtentX        =   19500
         _ExtentY        =   20135
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
         SpreadDesigner  =   "mkpoen05P.frx":0000
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
